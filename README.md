# robo-conferencia-sistemas-juridicos
Robô de conferência e batimento entre sistemas jurídicos (Legal One x ADVWIN)

## Batimento de publicações com a base do Legal One

Script que:

1. Extrai TODOS os números de processo no padrão CNJ (ex:
   `0118284-78.2026.8.04.1000`, com ou sem pontuação) citados no texto livre
   de uma coluna de uma planilha de publicações (padrão: coluna C).
2. Carrega os números de processo cadastrados em uma planilha de base do
   Legal One (a coluna de CNJ é detectada automaticamente, ou pode ser
   informada explicitamente).
3. Gera uma nova planilha `.xlsx` com **todas** as publicações da planilha
   original (nenhuma é descartada), preservando todas as colunas originais e
   adicionando uma coluna por CNJ citado no texto (`Número CNJ` quando há
   apenas um por publicação em toda a planilha, ou `Número CNJ 1`,
   `Número CNJ 2`, ... quando alguma publicação cita mais de um). As células
   cujo CNJ foi encontrado na base do Legal One são pintadas de verde, para
   facilitar a conferência manual — publicações muito longas costumam citar
   vários processos apenas como referência, então o robô não tenta adivinhar
   qual é "o" processo da publicação: lista todos e destaca os que já estão
   cadastrados no Legal One.

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
