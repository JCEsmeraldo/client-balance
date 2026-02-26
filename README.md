# Client Balance

Sistema desktop de controle financeiro por cliente, desenvolvido em Python com CustomTkinter e SQLite.

## Funcionalidades

- **Clientes** — cadastro completo com nome, e-mail, telefone e observacoes; listagem em tabela com selecao por clique; acoes de salvar, editar e excluir.
- **Movimentacoes** — registro de entradas e saidas vinculadas a um cliente; data preenchida automaticamente; listagem geral.
- **Extrato** — visualizacao das movimentacoes de um cliente especifico ordenadas por data; saldo total calculado com cor dinamica (verde/vermelho).

## Tecnologias

| Tecnologia | Versao minima |
|---|---|
| Python | 3.10 |
| customtkinter | 5.2.0 |
| SQLite | embutido no Python |

## Instalacao

```bash
# 1. Clone ou copie os arquivos para uma pasta
cd client-balance

# 2. Crie e ative um ambiente virtual (recomendado)
python -m venv .venv
source .venv/bin/activate   # macOS / Linux
# .venv\Scripts\activate    # Windows

# 3. Instale as dependencias
pip install -r requirements.txt

# 4. Execute a aplicacao
python app.py
```

O arquivo `banco.db` e criado automaticamente na primeira execucao, na mesma pasta do `app.py`.

## Estrutura do projeto

```
client-balance/
├── app.py           # codigo-fonte completo
├── banco.db         # banco SQLite (gerado automaticamente)
├── requirements.txt # dependencias Python
└── README.md        # este arquivo
```

## Gerar executavel (.exe / .app)

```bash
pip install pyinstaller

# macOS — gera ClientBalance.app em dist/
pyinstaller --onefile --windowed --name ClientBalance app.py

# Windows — gera ClientBalance.exe em dist/
pyinstaller --onefile --windowed --name ClientBalance app.py
```

O executavel sera gerado na pasta `dist/`.
Apos o build, as pastas `build/` e o arquivo `ClientBalance.spec` podem ser descartados.

## Schema do banco

**clientes**

| Coluna | Tipo |
|---|---|
| id | INTEGER PK AUTOINCREMENT |
| nome | TEXT NOT NULL |
| email | TEXT |
| telefone | TEXT |
| observacoes | TEXT |

**movimentacoes**

| Coluna | Tipo |
|---|---|
| id | INTEGER PK AUTOINCREMENT |
| cliente_id | INTEGER FK |
| tipo | TEXT (Entrada / Saida) |
| valor | REAL |
| descricao | TEXT |
| data | TEXT (ISO 8601) |

Ao excluir um cliente, todas as suas movimentacoes sao removidas automaticamente via `ON DELETE CASCADE`.

