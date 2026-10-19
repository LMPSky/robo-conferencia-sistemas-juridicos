# robo-conferencia-sistemas-juridicos
Robô de conferência e batimento entre sistemas jurídicos (Legal One x ADVWIN)

## Batimento de cadastro de processos (Legal One x planilhas do GED)

O que o robô faz: diariamente, times diferentes (Advogado, Robô TI, Sistema
Cliente, Triagem Publicações) sobem no GED do Legal One planilhas "Planilha
de Cadastro - <tipo>" com processos que precisam ser cadastrados como novos
processos no Legal One. Este robô confere, para cada processo listado nessas
planilhas, se ele já foi efetivamente cadastrado na base do Legal One, e gera
um relatório de pendências para enviar ao time de cadastro.

### Como organizar as pastas de entrada

```
entrada/
  base_legal_one/        <- cole aqui o export da base de processos do Legal One (.xlsx)
  planilhas_cadastro/    <- cole aqui as planilhas "PLANILHA DE CADASTRO - ..." baixadas do GED (.xlsx), uma ou várias
  export_ged/             <- OPCIONAL: export de metadados do GED (.xls/.xlsx), usado só para
                              enriquecer o relatório com tipo e data de envio de cada planilha
saida/                    <- o robô grava aqui o relatório, nome 'Batimento_AAAAMMDD_HHMMSS.xlsx'
```

As pastas são criadas automaticamente (com um `.gitkeep`) na primeira
execução, caso ainda não existam.

### Como rodar

```powershell
py -m pip install -r requirements.txt
py batimento.py
```

Não é necessário passar argumentos: o robô sempre lê de `entrada/` e grava em
`saida/` relativos ao diretório onde é executado. Se faltar algum arquivo
obrigatório (base do Legal One ou planilhas de cadastro), o robô imprime uma
mensagem amigável pedindo para colocar o arquivo na pasta correta, em vez de
um stack trace.

### Como interpretar o relatório de saída

O arquivo `saida/Batimento_AAAAMMDD_HHMMSS.xlsx` tem 3 abas:

- **Resumo**: total de processos, cadastrados, pendentes e sem número
  identificado, agrupado por arquivo de origem e por tipo de planilha, com a
  porcentagem de cadastrado.
- **Pendencias**: apenas os processos com status `PENDENTE` (não encontrado
  na base do Legal One) ou `SEM NÚMERO IDENTIFICADO` (linha sem CNJ nem outro
  número preenchido, requer checagem manual) — esta é a aba a enviar ao time
  de cadastro (Diego, com cópia à Adriana Torres).
- **Detalhe completo**: todas as linhas processadas, com todas as colunas e o
  status de cada uma. As linhas são coloridas: vermelho claro para
  `PENDENTE`, amarelo claro para `SEM NÚMERO IDENTIFICADO` e verde claro para
  `CADASTRADO`.

### Testes

```powershell
py -m pytest tests -q
```

Os testes usam fixtures sintéticas (nenhum dado real do usuário) para validar
que processos que batem com a base do Legal One saem como `CADASTRADO`, que
processos que não batem saem como `PENDENTE`, e que uma linha sem CNJ e sem
outro número sai como `SEM NÚMERO IDENTIFICADO`.

## Triagem de grupos do GED baixados em lote (sem abrir cada grupo)

Quando os processos estão organizados em pastas/grupos dentro do sistema e o
download em massa só traz uma planilha com o *nome* de cada grupo (sem o
conteúdo), abrir cada grupo manualmente é inviável se houver muitos. Este
script extrai automaticamente o número CNJ do nome do grupo (quando ele
estiver presente, ainda que junto de outro texto) e já classifica o grupo
comparando com a base do Legal One — assim só é preciso abrir manualmente os
grupos cujo nome não tem nenhum CNJ identificável.

### Uso

```powershell
$env:PYTHONPATH = "src"
py -m robo_conferencia.listar_pendencias_grupos caminho\grupos_ged.xlsx caminho\base_legal_one.xlsx caminho\saida.xlsx
```

A planilha de saída tem todas as colunas originais do arquivo de grupos, mais
`Número CNJ` e `Status`:

- `CADASTRADO`: o CNJ extraído do nome do grupo já está na base do Legal One.
- `PENDENTE`: o CNJ extraído do nome do grupo não está na base do Legal One.
- `SEM NÚMERO IDENTIFICADO`: não foi possível extrair um CNJ do nome do
  grupo — esse grupo precisa ser aberto manualmente para verificação.

Parâmetros opcionais:

- `--coluna-texto`: coluna da planilha de grupos com o nome/texto de cada
  grupo (letra do Excel, índice 0-based ou nome do cabeçalho). Padrão: `A`.
- `--aba-grupos`: nome da aba da planilha de grupos, se não for a primeira.
- `--coluna-cnj-legal-one` / `--aba-legal-one`: iguais ao script de
  batimento de publicações, abaixo.

### Testes

```powershell
py -m pytest tests -q
```

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
