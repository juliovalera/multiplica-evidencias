# Multiplica Evidências

Aplicação desktop em Python para registrar encontros do Programa Multiplica, armazenar evidências localmente e gerar relatórios mensais em `DOCX`, com exportação opcional para `PDF`.

## Características do MVP

- Cadastro local do professor multiplicador, sem dependência de internet.
- Cadastro reutilizável de turmas em SQLite.
- Registro de encontros com data, pauta, turma, situação, participantes e observação.
- Colagem de prints da área de transferência ou anexação de imagens do computador.
- Preenchimento automático de horário do encontro a partir da turma.
- Checklist mensal antes da geração do relatório.
- Geração de relatório em `DOCX` inspirada no modelo visual informado.
- Exportação opcional para `PDF` com `LibreOffice` ou `Microsoft Word` + `pywin32`.
- Backup simples do banco local.

## Estrutura

- `main.py`: ponto de entrada da aplicação.
- `config.py`: caminhos do projeto e utilitários.
- `models.py`: estruturas de dados e listas padrão.
- `database.py`: persistência SQLite, CRUD e checklist mensal.
- `relatorio.py`: geração do relatório `DOCX` e conversão opcional para `PDF`.
- `interface/app.py`: interface desktop com `Tkinter`.
- `data/`: banco local criado automaticamente.
- `evidencias/`: imagens salvas por mês, turma e pauta.
- `saidas/`: relatórios e backups gerados.
- `modelos/`: local para o arquivo de referência visual do relatório.

## Requisitos

- Python 3.11 ou superior
- Windows recomendado para colagem de imagens e abertura rápida de arquivos
- Um arquivo de modelo em `modelos/evidencias.docx`, se desejar usar a mesma referência visual

## Instalação

1. Criar ambiente virtual:

```powershell
python -m venv .venv
```

2. Ativar o ambiente virtual:

```powershell
.venv\Scripts\Activate.ps1
```

3. Instalar dependências:

```powershell
pip install -r requirements.txt
```

## Como executar

```powershell
python main.py
```

## Como usar

1. Abra a aplicação e aceite os termos de uso no primeiro acesso.
2. Preencha os dados na aba `Professor Multiplicador`.
3. Cadastre ou atualize as turmas na aba `Turmas`.
4. Registre cada encontro na aba `Encontros`.
5. Use `Colar print` para evidências copiadas do Teams ou `Adicionar imagem` para arquivos salvos.
6. Selecione uma evidência para pré-visualização quando necessário.
7. Revise a aba `Checklist e relatório`.
8. Gere o arquivo `DOCX` e, se disponível, também o `PDF`.

## Modelo DOCX

- Coloque o arquivo de referência em `modelos/evidencias.docx`.
- O aplicativo recria a estrutura visual principal com `python-docx`.
- Para manter o repositório público sem dados pessoais, o arquivo de modelo não é versionado por padrão.

## Privacidade

- Todos os dados ficam apenas no computador local.
- Não há envio para internet, APIs externas ou upload de arquivos.
- Banco, evidências e saídas geradas não são publicados no repositório.
