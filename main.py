import hashlib
import json
import os
import unicodedata
from datetime import datetime
import smtplib
from email.message import EmailMessage
from pathlib import Path
import sys
from concurrent.futures import ThreadPoolExecutor, as_completed
import asyncio

import flet as ft
import pandas as pd
from sqlalchemy import create_engine, text
from urllib.parse import quote_plus
from dotenv import load_dotenv

# -------------------------------------------------------------------
# CONFIGURAÇÕES BÁSICAS
# -------------------------------------------------------------------
if getattr(sys, "frozen", False):
    BASE_DIR = Path(sys.executable).resolve().parent
else:
    BASE_DIR = Path(__file__).resolve().parent

load_dotenv(BASE_DIR / ".env")


def build_db_url(base_dir: Path) -> str:
    db_url_env = os.getenv("EQS_DB_URL")
    if db_url_env:
        return db_url_env

    backend = os.getenv("EQS_DB_BACKEND", "sqlite").lower()

    if backend == "postgres":
        host = os.getenv("EQS_DB_HOST", "localhost")
        port = os.getenv("EQS_DB_PORT", "5432")
        name = os.getenv("EQS_DB_NAME", "eqs_automate")
        user = os.getenv("EQS_DB_USER", "postgres")
        password = os.getenv("EQS_DB_PASSWORD", "")

        user_enc = quote_plus(user)
        password_enc = quote_plus(password)

        return f"postgresql+psycopg2://{user_enc}:{password_enc}@{host}:{port}/{name}"

    db_path = base_dir / "EQS_automate_database.db"
    return f"sqlite:///{db_path}"


DB_URL = build_db_url(BASE_DIR)
engine = create_engine(DB_URL, echo=False, future=True)

IS_POSTGRES = DB_URL.startswith("postgresql")

# Pasta de entrada
INPUT_FOLDER = os.getenv("EQS_INPUT_FOLDER", "").strip()
if not INPUT_FOLDER:
    INPUT_FOLDER = str(BASE_DIR / "storage")

WORKERS = int(os.getenv("EQS_WORKERS", "4"))

# SMTP / Email
SMTP_SERVER = os.getenv("EQS_SMTP_SERVER", "smtp.office365.com")
SMTP_PORT = int(os.getenv("EQS_SMTP_PORT", "587"))
SMTP_USER = os.getenv("EQS_SMTP_USER", "")
SMTP_PASSWORD = os.getenv("EQS_SMTP_PASSWORD", "")
REPORT_EMAIL = os.getenv(
    "EQS_REPORT_EMAIL",
    "controladoria@eqsengenharia.com.br",
)

# -------------------------------------------------------------------
# EMAIL – RELATÓRIO EM HTML
# -------------------------------------------------------------------


def send_email_report(
    results: list[dict],
    log_lines: list[str],
    logger=None,
):
    def log(msg: str):
        if logger:
            logger(msg)

    if not SMTP_USER or not SMTP_PASSWORD:
        log("Configurações de e-mail não definidas (SMTP_USER/SMTP_PASSWORD). E-mail não enviado.")
        return

    if not results:
        log("Nenhum arquivo processado, e-mail não será enviado.")
        return

    log(f"Preparando relatório em HTML para {REPORT_EMAIL}...")

    now_str = datetime.now().strftime("%d/%m/%Y %H:%M:%S")

    rows_html = ""
    for r in results:
        status_icon = "✅" if r["ok"] else "❌"
        status_color = "#4CAF50" if r["ok"] else "#FF5252"

        rows_html += f"""
        <tr>
            <td style="padding:4px 8px;border:1px solid #555;">{r['file_name']}</td>
            <td style="padding:4px 8px;border:1px solid #555;">{r.get('table_name') or '-'}</td>
            <td style="padding:4px 8px;border:1px solid #555;text-align:center;color:{status_color};">{status_icon}</td>
            <td style="padding:4px 8px;border:1px solid #555;">{r['message']}</td>
        </tr>
        """

    html = f"""
    <html>
    <body style="font-family: Arial, sans-serif; background-color:#202020; color:#FFFFFF;">
        <h2 style="margin-bottom:4px;">Relatório de Upload EQS Automate</h2>
        <p style="margin-top:0;">Data: {now_str}</p>

        <table style="border-collapse: collapse; width: 100%; margin-top:10px; font-size:13px;">
            <tr style="background-color:#333333;">
                <th style="padding:6px 8px;border:1px solid #555;">Arquivo</th>
                <th style="padding:6px 8px;border:1px solid #555;">Tabela</th>
                <th style="padding:6px 8px;border:1px solid #555;">Status</th>
                <th style="padding:6px 8px;border:1px solid #555;">Mensagem</th>
            </tr>
            {rows_html}
        </table>

        <p style="margin-top:16px;font-size:11px;color:#CCCCCC;">
            Este e-mail foi gerado automaticamente pelo EQS Automate Conversor - Excel → Banco.
        </p>
    </body>
    </html>
    """

    msg = EmailMessage()
    msg["From"] = SMTP_USER
    msg["To"] = REPORT_EMAIL
    msg["Subject"] = "[EQS Automate] Relatório de upload de arquivos Excel"
    msg.set_content("Seu cliente de e-mail não suporta HTML.")
    msg.add_alternative(html, subtype="html")

    if any(not r["ok"] for r in results):
        log("Foram detectados erros, anexando log_automacao.txt ao e-mail...")
        log_text = "\n".join(log_lines)
        msg.add_attachment(
            log_text.encode("utf-8"),
            maintype="text",
            subtype="plain",
            filename="log_automacao.txt",
        )

    try:
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()
            server.login(SMTP_USER, SMTP_PASSWORD)
            server.send_message(msg)
        log("E-mail de relatório enviado com sucesso.")
    except Exception as e:
        log(f"Falha ao enviar e-mail: {e}")


# -------------------------------------------------------------------
# FUNÇÕES DE SUPORTE (NORMALIZAÇÃO / BD)
# -------------------------------------------------------------------


def normalize_string(value: str) -> str:
    if not value:
        return ""

    value = str(value).strip().lower()

    value = (
        unicodedata.normalize("NFKD", value)
        .encode("ascii", "ignore")
        .decode("ascii")
    )

    separators = [" ", "-", ".", ",", ";", "/", "\\", ":", "?", "!", "%", "(", ")", "[", "]"]
    for sep in separators:
        value = value.replace(sep, "_")

    value = "".join(ch for ch in value if ch.isalnum() or ch == "_")

    while "__" in value:
        value = value.replace("__", "_")

    value = value.strip("_")

    if not value:
        value = "coluna"

    return value


def table_name_from_path(path: str) -> str:
    base = os.path.basename(path)
    name, _ = os.path.splitext(base)
    return normalize_string(name)


def infer_sql_type(dtype) -> str:

    if pd.api.types.is_integer_dtype(dtype):
        return "BIGINT" if IS_POSTGRES else "INTEGER"
    if pd.api.types.is_float_dtype(dtype):
        return "DOUBLE PRECISION" if IS_POSTGRES else "REAL"
    if pd.api.types.is_datetime64_any_dtype(dtype):
        return "TIMESTAMP"
    if pd.api.types.is_bool_dtype(dtype):
        return "INTEGER"
    return "TEXT"


def prepare_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    new_columns = {}
    used_names = set()

    for col in df.columns:
        base = normalize_string(col)
        name = base or "coluna"

        if name in used_names:
            suffix = 2
            new_name = f"{name}_{suffix}"
            while new_name in used_names:
                suffix += 1
                new_name = f"{name}_{suffix}"
            name = new_name

        used_names.add(name)
        new_columns[col] = name

    df = df.rename(columns=new_columns)

    for col in df.columns:
        if pd.api.types.is_datetime64_any_dtype(df[col]):
            df[col] = df[col].dt.strftime("%Y-%m-%d %H:%M:%S")

    df = df.dropna(axis=1, how="all")

    return df


def ensure_table(table_name: str, df: pd.DataFrame, logger=None):
    if logger:
        logger(f"Criando/verificando tabela '{table_name}' no banco de dados...")

    columns_def = []
    for col in df.columns:
        sql_type = infer_sql_type(df[col].dtype)
        columns_def.append(f'"{col}" {sql_type}')

    columns_def.append('_row_hash TEXT UNIQUE')

    columns_sql = ", ".join(columns_def)

    if IS_POSTGRES:
        id_column = "id BIGSERIAL PRIMARY KEY"
    else:
        id_column = "id INTEGER PRIMARY KEY AUTOINCREMENT"

    create_sql = f"""
    CREATE TABLE IF NOT EXISTS "{table_name}" (
        {id_column},
        {columns_sql}
    );
    """

    with engine.begin() as conn:
        conn.execute(text(create_sql))

    if logger:
        logger(f"Tabela '{table_name}' pronta para receber dados.")


def insert_rows(table_name: str, df: pd.DataFrame, logger=None):
    total = len(df)
    if logger:
        logger(f"Iniciando inserção de {total} linha(s) em '{table_name}'...")

    inserted = 0

    with engine.begin() as conn:
        for _, row in df.iterrows():
            data = row.to_dict()

            row_hash = hashlib.sha256(
                json.dumps(data, sort_keys=True, default=str).encode("utf-8")
            ).hexdigest()

            data["_row_hash"] = row_hash

            cols = list(data.keys())
            cols_sql = ", ".join(f'"{c}"' for c in cols)
            params_sql = ", ".join(f":{c}" for c in cols)

            sql = text(
                f"""
                INSERT INTO "{table_name}" ({cols_sql})
                VALUES ({params_sql})
                ON CONFLICT(_row_hash) DO NOTHING;
                """
            )

            result = conn.execute(sql, data)

            if result.rowcount and result.rowcount > 0:
                inserted += 1

    if logger:
        logger(
            f"Inserção concluída: {inserted}/{total} linha(s) inseridas "
            f"(as demais já existiam)."
        )


def process_excel_file(file_path: str, logger=None) -> dict:
    def log(msg: str):
        if logger:
            logger(msg)

    if not os.path.exists(file_path):
        msg = f"Arquivo não encontrado: {file_path}"
        log(msg)
        return {"ok": False, "table_name": None, "message": msg}

    try:
        log(f"Lendo arquivo Excel: {file_path}")
        df = pd.read_excel(file_path)
    except Exception as e:
        msg = f"Erro ao ler Excel: {e}"
        log(msg)
        return {"ok": False, "table_name": None, "message": msg}

    if df.empty:
        msg = "Planilha vazia, nada a importar."
        log(msg)
        return {"ok": False, "table_name": None, "message": msg}

    log("Tratando cabeçalhos e tipos de dados...")
    df = prepare_dataframe(df)

    table_name = table_name_from_path(file_path)
    log(f"Tabela alvo derivada do arquivo: '{table_name}'")

    try:
        ensure_table(table_name, df, logger=log)
        insert_rows(table_name, df, logger=log)
    except Exception as e:
        msg = f"Erro ao inserir dados no banco: {e}"
        log(msg)
        return {"ok": False, "table_name": table_name, "message": msg}

    msg = (
        f"Importação concluída para a tabela '{table_name}' "
        f"com {len(df)} linha(s) processadas (duplicadas ignoradas)."
    )
    log("Processo finalizado com sucesso.")
    return {"ok": True, "table_name": table_name, "message": msg}


def process_all_files_in_folder(
    folder_path: str,
    logger=None,
    workers: int = 4,
) -> list[dict]:
    def ui_log(msg: str):
        if logger:
            logger(msg)

    results: list[dict] = []

    if not folder_path or not os.path.isdir(folder_path):
        ui_log(f"Pasta de entrada inválida ou inexistente: {folder_path!r}")
        return results

    ui_log(f"Varredura iniciada na pasta: {folder_path}")

    excel_files = [
        os.path.join(folder_path, f)
        for f in os.listdir(folder_path)
        if f.lower().endswith((".xlsx", ".xls"))
    ]

    if not excel_files:
        ui_log("Nenhum arquivo Excel (*.xlsx / *.xls) encontrado na pasta.")
        return results

    ui_log(f"{len(excel_files)} arquivo(s) Excel encontrado(s).")
    ui_log(f"Processamento paralelo com até {workers} worker(s)...")

    def worker_task(file_path: str):
        local_logs: list[str] = []

        def thread_log(message: str):
            local_logs.append(message)

        result = process_excel_file(file_path, logger=thread_log)
        return file_path, result, local_logs

    with ThreadPoolExecutor(max_workers=workers) as executor:
        future_map = {
            executor.submit(worker_task, file_path): file_path
            for file_path in excel_files
        }

    for future in as_completed(future_map):
        file_path = future_map[future]
        file_name = os.path.basename(file_path)

        try:
            file_path_result, result, local_logs = future.result()

            for msg in local_logs:
                ui_log(f"[{file_name}] {msg}")

            status = "SUCESSO" if result["ok"] else "ERRO"
            ui_log(f"--- Finalizado {status} para o arquivo: {file_name} ---")

            results.append(
                {
                    "file_path": file_path_result,
                    "file_name": file_name,
                    "ok": result["ok"],
                    "table_name": result["table_name"],
                    "message": result["message"],
                }
            )

        except Exception as e:
            msg = f"Erro inesperado no processamento de {file_name}: {e}"
            ui_log(msg)
            results.append(
                {
                    "file_path": file_path,
                    "file_name": file_name,
                    "ok": False,
                    "table_name": None,
                    "message": msg,
                }
            )

    ui_log("Varredura da pasta concluída (execução paralela finalizada).")
    return results


# -------------------------------------------------------------------
# INTERFACE FLET
# -------------------------------------------------------------------


def main(page: ft.Page):
    page.title = "EQS Automate Conversor - Excel → Banco"
    page.window.height = 800
    page.window.width = 700
    page.window.resizable = False

    current_input_folder = INPUT_FOLDER

    status_text = ft.Text("Clique no botão para processar a pasta configurada no .env.")
    folder_text = ft.Text(
        f"Pasta configurada (EQS_INPUT_FOLDER): {current_input_folder or '[NÃO DEFINIDA]'}",
        size=12,
        color=ft.Colors.GREY,
    )
    workers_text = ft.Text(
        f"Workers paralelos (EQS_WORKERS): {WORKERS}",
        size=12,
        color=ft.Colors.GREY,
    )

    log_view = ft.ListView(
        expand=1,
        spacing=5,
        auto_scroll=True,
    )

    log_buffer: list[str] = []

    def ui_log(message: str):
        timestamp = datetime.now().strftime("%d-%m-%Y %H:%M:%S")
        line = f"{timestamp} - {message}"
        log_buffer.append(line)
        log_view.controls.append(ft.Text(line, size=12))
        page.update()

    def update_env_input_folder(new_folder: str):
        env_path = BASE_DIR / ".env"
        lines = []
        if env_path.exists():
            lines = env_path.read_text(encoding="utf-8").splitlines()

        found = False
        new_lines = []
        for line in lines:
            if line.strip().startswith("EQS_INPUT_FOLDER="):
                new_lines.append(f"EQS_INPUT_FOLDER={new_folder}")
                found = True
            else:
                new_lines.append(line)

        if not found:
            new_lines.append(f"EQS_INPUT_FOLDER={new_folder}")

        env_path.write_text("\n".join(new_lines) + "\n", encoding="utf-8")
        ui_log(f"EQS_INPUT_FOLDER atualizado no .env para: {new_folder}")

    def on_folder_result(e: ft.FilePickerResultEvent):
        nonlocal current_input_folder
        if e.path:
            current_input_folder = e.path
            folder_text.value = (
                f"Pasta configurada (EQS_INPUT_FOLDER): {current_input_folder}"
            )
            update_env_input_folder(current_input_folder)
            page.update()
        else:
            ui_log("Nenhuma pasta selecionada.")

    folder_picker = ft.FilePicker(on_result=on_folder_result)
    page.overlay.append(folder_picker)

    def choose_folder(e):
        folder_picker.get_directory_path()

    def start_auto_close():
        countdown_text = ft.Text("Fechando em 3...")

        dialog = ft.AlertDialog(
            modal=True,
            title=ft.Text("Processamento concluído"),
            content=countdown_text,
        )

        page.dialog = dialog
        dialog.open = True
        page.update()

        async def do_countdown():
            for i in range(3, 0, -1):
                countdown_text.value = f"Fechando em {i}..."
                page.update()
                await asyncio.sleep(1)
            await asyncio.sleep(0.5)
            page.window.close()

        page.run_task(do_countdown)

    def on_run_click(e):
        log_view.controls.clear()
        log_buffer.clear()
        page.update()

        if not current_input_folder:
            status_text.value = "EQS_INPUT_FOLDER não está definido no .env."
            ui_log(status_text.value)
            page.update()
            return

        status_text.value = "Iniciando processamento da pasta..."
        page.update()

        results = process_all_files_in_folder(
            current_input_folder,
            logger=ui_log,
            workers=WORKERS,
        )

        if not results:
            status_text.value = "Nenhum arquivo foi processado."
            page.update()
            return

        ok_count = sum(1 for r in results if r["ok"])
        err_count = len(results) - ok_count
        status_text.value = (
            f"Processamento concluído. Sucesso: {ok_count}, Erros: {err_count}."
        )
        page.update()

        ui_log("Gerando e-mail de relatório...")
        send_email_report(results, log_buffer, logger=ui_log)
        page.update()

        start_auto_close()

    run_button = ft.ElevatedButton(
        "Processar pasta de arquivos Excel",
        on_click=on_run_click,
    )

    choose_folder_button = ft.ElevatedButton(
        "Selecionar pasta de arquivos Excel",
        on_click=choose_folder,
    )

    page.add(
        ft.Column(
            [
                ft.Text(
                    "Importador Excel → Banco de Dados",
                    size=20,
                    weight=ft.FontWeight.BOLD,
                ),
                ft.Row(
                    [run_button, choose_folder_button],
                    spacing=10,
                ),
                status_text,
                folder_text,
                workers_text,
                ft.Text("Log do processo:", size=14, weight=ft.FontWeight.BOLD),
                ft.Container(
                    content=log_view,
                    height=600,
                    border=ft.border.all(1, ft.Colors.GREY),
                    padding=10,
                ),
            ],
            spacing=12,
            alignment=ft.MainAxisAlignment.START,
        )
    )


if __name__ == "__main__":
    ft.app(target=main)
