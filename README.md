# robo-conferencia-sistemas-juridicos
Robô de conferência e batimento entre sistemas jurídicos (Legal One x ADVWIN)

## Batimento de publicações com a base do Legal One

Script que:

1. Extrai todos os números de processo no padrão CNJ (ex:
   `0118284-78.2026.8.04.1000`) encontrados no texto livre de uma coluna de
   uma planilha de publicações (padrão: coluna C).
2. Carrega os números de processo cadastrados em uma planilha de base do
   Legal One (a coluna de CNJ é detectada automaticamente, ou pode ser
   informada explicitamente).
3. Gera uma nova planilha `.xlsx` apenas com as publicações cujo CNJ bate com
   algum processo da base do Legal One, preservando todas as colunas
   originais da publicação e adicionando uma coluna `Número CNJ` com o
   processo encontrado. Publicações sem nenhum CNJ na base são descartadas.

### Instalação

```powershell
py -m pip install -r requirements.txt
```

### Uso

```powershell
py -m robo_conferencia.bater_legal_one caminho\publicacoes.xlsx caminho\base_legal_one.xlsx caminho\saida.xlsx
```

É preciso rodar com `src` no `PYTHONPATH` (ou instalar o pacote), por exemplo:

```powershell
$env:PYTHONPATH = "src"
py -m robo_conferencia.bater_legal_one publicacoes.xlsx legal_one.xlsx saida.xlsx
```

Parâmetros opcionais:

- `--coluna-texto`: coluna da planilha de publicações com o texto a analisar
  (letra do Excel, índice 0-based ou nome do cabeçalho). Padrão: `C`.
- `--aba-publicacoes`: nome da aba da planilha de publicações, se não for a
  primeira aba com dados.
- `--coluna-cnj-legal-one`: coluna da base do Legal One com o número do
  processo, caso a detecção automática não funcione.
- `--aba-legal-one`: nome da aba da planilha de base do Legal One, se não for
  a primeira aba com dados.

### Testes

```powershell
py -m pytest tests -q
```
