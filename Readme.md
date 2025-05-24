## 📧 Automação de Envio de E-mails com Faturas em Python

Este projeto é um script em Python para automatizar o envio de e-mails com faturas em anexo. Ele lê dados de clientes e faturas de uma planilha Excel, monta e-mails personalizados em HTML com anexos e envia-os em massa, registrando falhas para análise posterior.

---

### ✨ Funcionalidades

- **Leitura de Dados via Excel:** Utiliza `pandas` para ler informações dos clientes de um arquivo `.xlsx`.
- **E-mails Personalizados:** Gera e-mails em HTML com assuntos dinâmicos para cada destinatário.
- **Envio com Anexos:** Anexa arquivos PDF de fatura a cada e-mail.
- **Assinatura em HTML com Imagem:** Incorpora uma assinatura profissional em HTML com imagem.
- **Suporte a CC:** Permite enviar cópias para outros destinatários especificados.
- **Log de Falhas:** Gera um arquivo `log_falhas_envios.xlsx` em caso de erro no envio.
- **Configuração Segura:** Credenciais sensíveis ficam em um arquivo `.py` ignorado pelo versionamento.

---

### 🛠️ Pré-requisitos

- Python 3.7+
- pip

---

### ⚙️ Instalação e Configuração

1. **Clone o Repositório**
    ```bash
    git clone https://github.com/bericktk/automacao-email.git
    cd automacao-email
    ```

2. **Instale as Dependências**
    ```bash
    pip install pandas openpyxl
    ```
    (Opcional: crie um `requirements.txt` com `pandas` e execute `pip install -r requirements.txt`)

3. **Configure suas Credenciais**

    Crie um arquivo chamado `dados.py` na raiz do projeto com o seguinte conteúdo:

    ```python
    # dados.py

    nomeDoRemetente = "Seu Nome ou Nome da Empresa"
    emailDoRemetente = "seu.email@gmail.com"
    # Use uma "Senha de App" gerada pelo Google, não sua senha normal.
    senhaAppGoogle = "sua_senha_de_app_aqui"
    servidorDoSMTP = "smtp.gmail.com"
    portaDoSMTP = 587
    ```

    > **Importante:** Para Gmail, gere uma "Senha de App" ([veja como](https://support.google.com/accounts/answer/185833)).

4. **Prepare a Planilha de Clientes**

    Renomeie sua planilha para `clientes.xlsx` ou ajuste a variável `arquivoPlanilhaFaturas` no script. A planilha deve conter uma aba chamada `Plan1` (ou ajuste `nomeAbaPlanilha`) com as colunas:

    | Nome         | Email                | ArquivoFatura                | Valor   | Vencimento   | EmailCopia                |
    |--------------|----------------------|------------------------------|---------|--------------|---------------------------|
    | Berick Dev   | contato@berick.dev   | faturas_pdf/fatura_berick.pdf| 150.00  | 2025-06-30   | financeiro@empresa.com    |

    - `EmailCopia` é opcional e pode conter múltiplos e-mails separados por `,` ou `;`.

5. **Adicione os Arquivos**

    - Coloque os PDFs das faturas na pasta indicada na planilha.
    - Coloque a imagem da assinatura (ex: `assinatura_email.png`) na raiz do projeto ou ajuste a variável `imagemAssinaturaEmail` no script.

---

### ▶️ Como Usar

Execute o script principal:

```bash
python main.py
```

O status de cada envio será exibido no terminal. Se houver falhas, será criado o arquivo `log_falhas_envios.xlsx`.

---

### 📁 Estrutura do Projeto

```
.
├── main.py                    # Script principal
├── dados.py                   # Credenciais (NÃO versionado)
├── clientes.xlsx              # Planilha de clientes
├── assinatura_email.png       # Imagem da assinatura
├── faturas_pdf/               # Pasta das faturas
│   └── fatura_exemplo.pdf
└── README.md                  # Este arquivo
```

---

### ⚖️ Licença

Este projeto está sob a Licença MIT. Veja o arquivo LICENSE para mais detalhes.