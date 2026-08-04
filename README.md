# Gerador de Comprovantes

Gera comprovantes de pagamento em PDF a partir de uma planilha Excel e um
modelo Word, com interface gráfica (Tkinter). Funciona em **Windows, macOS
e Linux**.

## ⚠️ Ação necessária antes de usar

O arquivo `modelos/MODELO_COMPROVANTE.docx` está **vazio** (sem texto).
Adicione nele o texto do comprovante com os placeholders exatos:

```
{{Nome}}
{{CPF}}
{{Valor}}
```

Eles podem estar em parágrafos comuns, dentro de tabelas, ou em
cabeçalho/rodapé — o preenchedor cobre todos esses casos.

## Arquitetura

O projeto segue **arquitetura em camadas**, com Injeção de Dependência
manual e o Princípio da Inversão de Dependência (a camada de orquestração
depende de abstrações, não de implementações concretas):

```
src/
└── main.py                              # composition root: monta e injeta as dependências
└── gerador_comprovantes/
    ├── config.py                        # configurações centralizadas (Settings, imutável)
    ├── exceptions.py                    # exceções específicas do domínio
    ├── models.py                        # entidade de domínio: Beneficiario
    ├── logging_config.py                # configuração de logging
    ├── services/
    │   ├── formatadores.py              # funções puras (CPF, valor, normalização)
    │   ├── identificador_colunas.py     # heurística de identificação de colunas
    │   ├── leitor_excel.py              # leitura da planilha
    │   ├── preenchedor_word.py          # preenchimento do modelo .docx
    │   ├── conversor_pdf.py             # abstração + implementação LibreOffice (multiplataforma)
    │   ├── resolvedor_caminhos.py       # caminhos multiplataforma (dev vs. empacotado)
    │   └── gerador_comprovantes.py      # orquestrador do caso de uso (sem dependência de GUI)
    └── gui/
        └── app.py                       # camada de apresentação (Tkinter), roda em thread separada

tests/                                   # testes unitários com dublês (mocks) das dependências
```

### Por que essa organização

- **Separação de camadas**: apresentação (`gui/`) → orquestração
  (`gerador_comprovantes.py`) → regras de negócio (`identificador_colunas`,
  `models`) → integração externa (`conversor_pdf`, `leitor_excel`). Cada
  camada só conhece a camada logo abaixo dela.
- **Injeção de Dependência**: `GeradorDeComprovantes` recebe todos os
  serviços prontos pelo construtor, em vez de criá-los internamente. Isso
  permite testar a lógica de orquestração com dublês (mocks), sem precisar
  de Excel, Word ou LibreOffice de verdade — veja
  `tests/test_gerador_comprovantes.py`.
- **Observer pattern para progresso**: o orquestrador não conhece Tkinter.
  Ele reporta progresso por um callback (`on_progresso`), e quem chama
  decide o que fazer com isso (atualizar uma barra, imprimir no console,
  etc). No projeto original, a lógica de geração manipulava widgets
  diretamente — isso foi eliminado.
- **Interface abstrata para conversão de PDF** (`ConversorPdf`): a
  implementação atual usa LibreOffice, mas o resto do sistema depende só
  da abstração. Trocar de motor de conversão no futuro não exige mudar o
  orquestrador.
- **POO onde há estado/comportamento a encapsular** (`Beneficiario`,
  `IdentificadorDeColunas`, `ConversorLibreOffice`); **funções puras onde
  não há** (`formatadores.py`) — evita criar classes artificiais só por
  convenção.
- **Multiplataforma de verdade**: `ResolvedorDeCaminhos` trata os 3
  formatos de empacotamento do PyInstaller (script, `.app` do macOS,
  pasta com `.exe`/binário no Windows/Linux). `ConversorLibreOffice` busca
  o executável primeiro via `PATH` (`shutil.which`) e depois em caminhos
  padrão de instalação de cada sistema operacional.

## Rodando localmente

```bash
python -m venv .venv
source .venv/bin/activate          # Windows: .venv\Scripts\activate
pip install -r requirements-dev.txt

# rodar os testes
pytest

# rodar a aplicação
python src/main.py
```

Requer o [LibreOffice](https://www.libreoffice.org/download/download/)
instalado (usado para converter .docx em .pdf).

## Gerando o executável

O PyInstaller precisa ser rodado **no próprio sistema operacional de
destino** (não dá para gerar o `.exe` do Windows a partir do macOS, por
exemplo):

```bash
pip install pyinstaller
pyinstaller --name main --windowed --add-data "dados:dados" --add-data "modelos:modelos" src/main.py
```

No Windows, troque `:` por `;` no `--add-data`.
