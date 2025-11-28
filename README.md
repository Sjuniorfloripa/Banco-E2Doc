
# 📁 EQS Automate Conversor – Excel → Banco + Relatório Automático

Automação desenvolvida para processar arquivos Excel em lote, converter dados para banco de dados SQL e enviar relatório por e-mail com status visual de sucesso ou falha.  
Criada com **Python + Flet**, com interface simples e pronta para uso por qualquer usuário, inclusive sem conhecimento técnico.

---

### 🔥 Funcionalidades principais

| Recurso | Status |
|--------|--------|
| Leitura automática de todos os Excel dentro de uma pasta | ✅ |
| Criação da tabela no banco caso ainda não exista | 🏗️ |
| Inserção de dados com prevenção contra duplicidade (via hash) | 🔐 |
| Log de todas as operações em tempo real no painel da interface | 📋 |
| Geração de relatório profissional em HTML (similar a dashboards E2Doc) | 📊 |
| Envio automático por e-mail (Outlook/Office 365) | 📩 |
| Anexo de log `.txt` apenas se houver erros | ⚠️ |
| Possibilidade de empacotamento em `.exe` para distribuição corporativa | 🚀 |

---

### ⚙️ Como funciona

1. Defina no `.env` o diretório onde estão os arquivos Excel.
2. Abra o programa (ou o executável gerado via *flet pack*).
3. Clique no botão **"Processar pasta"**.
4. A aplicação irá:
   - varrer todos os arquivos `.xlsx/.xls`,
   - processar e inserir no banco SQLite,
   - gerar tabela HTML com status por arquivo,
   - enviar relatório automático via e-mail corporativo.

*Caso algum arquivo falhe → o log completo é anexado automaticamente ao e-mail.*

---

### 🧱 Tecnologias utilizadas

| Tecnologia | Uso |
|----------|-----|
| Python | Núcleo da automação |
| Pandas | Tratamento das planilhas |
| SQLAlchemy | Comunicação com o banco SQLite |
| Flet | Interface gráfica |
| Python-dotenv | Leitura de configurações via `.env` |
| SMTP (Office365) | Envio do relatório por e-mail |

---

### 📦 Deploy como executável

O projeto pode ser convertido em `.exe` com:

```bash
flet pack main.py --name "EQS Automate Conversor"
```

Isso permite distribuição interna na empresa sem necessidade de instalar Python.

---

### 💼 Ideal para:

✔ Departamentos financeiros  
✔ Controladoria  
✔ RH e auditoria de documentação  
✔ Integrações e uploads recorrentes  
✔ Processamento repetitivo de planilhas

---
