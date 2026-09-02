# Robô SIAPE

Automação em Python para download e tratamento dos dados de **remuneração de servidores públicos federais — SIAPE**, a partir dos arquivos disponibilizados pelo Portal da Transparência.

A aplicação recebe **ano e mês** como entrada, realiza o download do pacote correspondente, extrai e trata o CSV de remuneração e gera uma planilha Excel pronta para consulta, além de registrar toda a execução em log.

## Objetivo

O projeto foi desenvolvido para automatizar o processo de obtenção e preparação da base de remuneração do SIAPE, evitando tratamento manual de arquivos grandes.

O fluxo principal é:

```text
Ano + Mês
   ↓
Validação da competência
   ↓
Download do pacote SIAPE
   ↓
Validação do ZIP
   ↓
Extração
   ↓
Remoção de arquivos desnecessários
   ↓
Localização do CSV de remuneração
   ↓
Tratamento da base
   ↓
Geração do Excel
   ↓
Registro da execução
   ↓
Limpeza dos arquivos temporários
```

## Funcionalidades

- Seleção de ano.
- Seleção de mês.
- Validação do ano e do mês.
- Bloqueio de competências futuras.
- Verificação dos meses disponíveis.
- Download direto por HTTP com `requests`.
- Download sem Selenium e sem necessidade de abrir navegador.
- Uso de dois endereços de download:
  - Portal da Transparência;
  - endereço estático de dados da CGU.
- Download em arquivo temporário `.parcial`.
- Validação do conteúdo recebido como ZIP.
- Até **3 tentativas** de download em caso de falha.
- Pausa de 2 segundos entre tentativas.
- Extração automática do ZIP.
- Remoção de arquivos considerados desnecessários.
- Localização automática do CSV de remuneração.
- Suporte a diferentes codificações do CSV.
- Leitura do CSV com separador `;`.
- Identificação de colunas que possuem dados.
- Remoção de linhas completamente vazias.
- Preservação da base original em CSV.
- Geração de planilha Excel com `xlsxwriter`.
- Conversão de valores monetários para números.
- Formatação dos valores monetários em reais.
- Ajuste das larguras das colunas.
- Cabeçalho formatado.
- Congelamento da primeira linha.
- Filtro automático na planilha.
- Processamento utilizando modo `constant_memory` do `xlsxwriter`, adequado para grandes volumes de dados e significativamente mais rápido que a geração anterior com `openpyxl`.
- Registro detalhado da execução.
- Log com data, hora e nível (`INFO`, `WARNING` e `ERROR`).
- Limpeza dos arquivos temporários após a execução.
- Criação de uma pasta exclusiva para cada execução.
- Interface gráfica em PyQt6, com login local, painel de controle, histórico de execuções e configurações do usuário.
- Janela abre em tamanho maior por padrão (1250x750), sem botão de maximizar/tela cheia e sem poder redimensionar pelas bordas (removido por causar bug no layout).
- Botões do card "Ações rápidas" e da barra "Resumo dos logs / Perfil / Ajuda" com o mesmo estilo e tamanho, mantendo consistência visual entre os dois blocos.
- Diálogo de ajuda passo a passo, com prints reais da interface anotados (setas/realces) mostrando onde clicar.
- Empacotamento em executável (`.exe`) via PyInstaller, funcionando sem precisar do Python instalado na máquina de destino.

## Pacote utilizado

O robô está direcionado ao pacote:

```text
Servidores_SIAPE
```

A interface solicita apenas a **competência (ano e mês)**. O tipo do pacote não é uma opção escolhida pelo usuário.

O nome do arquivo baixado é montado a partir da competência e do tipo, por exemplo:

```text
202503_Servidores_SIAPE.zip
```

## Processamento dos dados

Após o download, o ZIP é extraído para uma área temporária.

O robô remove arquivos que não são necessários para o processamento da remuneração, incluindo arquivos relacionados a:

```text
afastamentos
cadastro
observacoes
observações
```

Em seguida, procura um arquivo CSV cujo nome contenha `remuneracao` ou `remuneração`.

Se o CSV de remuneração não for encontrado, a execução é interrompida e o erro é registrado.

## Tratamento do CSV

O CSV é lido utilizando `;` como delimitador.

O sistema tenta as seguintes codificações:

```text
UTF-8 com BOM
UTF-8
CP1252
Latin-1
```

Durante o tratamento:

- linhas completamente vazias são ignoradas;
- colunas sem nenhum valor preenchido são removidas;
- campos vazios continuam vazios;
- as colunas válidas são preservadas na ordem original;
- os valores monetários identificados são convertidos para números.

## Valores monetários

O tratamento dos valores monetários é direcionado à **estrutura específica da planilha SIAPE utilizada pelo projeto**.

Para identificar quais colunas devem ser tratadas como monetárias, o robô utiliza a lista `PALAVRAS_MONETARIAS`. A lista contém termos que aparecem nos nomes das colunas financeiras dessa base.

Atualmente, são considerados:

```python
PALAVRAS_MONETARIAS = [
    "remuneracao",
    "remuneração",
    "remunerações",
    "remuneracoes",
    "abate-teto",
    "gratificação",
    "gratificacao",
    "férias",
    "ferias",
    "irrf",
    "pss",
    "rpgs",
    "dedução",
    "deducao",
    "deduções",
    "deducoes",
    "pensão",
    "pensao",
    "fundo",
    "taxa",
    "verbas",
    "total",
]
```

Esses termos **não representam uma lista genérica para qualquer planilha**. Eles foram definidos de acordo com os campos presentes na planilha SIAPE utilizada pelo robô.

Por isso, palavras como `fundo`, `taxa`, `verbas` e `total` permanecem na lista de forma intencional: dentro da estrutura dessa base específica, elas são utilizadas para identificar campos que devem receber tratamento monetário.

O robô procura os termos nos nomes das colunas e, quando uma coluna é identificada como monetária, tenta converter os valores para números.

Exemplo:

```text
R$ 12.345,67
       ↓
12345.67
```

Depois da conversão, o Excel recebe a formatação monetária:

```text
R$ #,##0.00
```

Isso permite que os valores sejam reconhecidos como números na planilha e utilizados em operações como filtros, somas e médias.

## Estrutura do projeto

O projeto tem três "camadas" diferentes de arquivos: o que é código-fonte (versionado no GitHub), o que é gerado ao empacotar o `.exe`, e o que é gerado ao executar o robô. Cada uma é detalhada abaixo.

### 1. O que vai para o GitHub

```text
.
├── build.py
├── interface.py
├── robo_siape.py
├── README.md
├── requirements.txt
├── .gitignore
├── .env.example
└── assets/
    └── ajuda/
        ├── passo1_painel.png
        ├── passo2_periodo.png
        └── passo3_iniciar.png
```

Apenas código-fonte, documentação e arquivos de configuração de exemplo. Nada gerado por build ou execução é versionado — tudo isso está listado no `.gitignore`.

> Essa é a estrutura da branch `main`. A branch `dev` acrescenta apenas o
> `teste.py`, usado para estressar o robô — ver [Testes (branch dev)](#testes-branch-dev).

#### `interface.py`

Responsável pela interface gráfica da aplicação (PyQt6) e pela interação com o usuário: login local, cadastro, seleção de ano/mês, painel de controle com histórico e indicadores, acompanhamento da execução em tempo real, exibição dos logs e o diálogo de ajuda passo a passo (ver `assets/ajuda/` abaixo).

A janela principal abre em 1400x860, pode ser redimensionada livremente pelas bordas, mas não pode ser maximizada — o botão de maximizar foi removido (junto com o duplo clique na barra de título) porque o modo maximizado bugava o layout dos cartões. No card "Ações rápidas" e na barra "Resumo dos logs / Perfil / Ajuda" do painel, os botões usam o mesmo estilo (`compacto`), sem variação de tamanho entre os dois blocos, e o botão de ajuda que ficava duplicado ao lado de "Configurar robô" foi removido (a ação já existe na barra de atalhos).

#### `robo_siape.py`

Contém o pipeline completo da automação: download direto por HTTP, validação, extração do ZIP, tratamento do CSV, geração do Excel, logging e limpeza dos arquivos temporários.

#### `build.py`

Gera o executável (`RoboSIAPE.exe`) a partir de `interface.py`, usando PyInstaller em modo `--onefile --windowed`. Também limpa builds anteriores e adiciona metadados de versão ao executável (nome, descrição, versão), o que ajuda a reduzir falsos positivos de antivírus/SmartScreen.

#### `requirements.txt`

Lista as dependências externas necessárias para executar o projeto (e para gerar o `.exe`).

#### `.gitignore`

Impede que ambientes virtuais, arquivos temporários, logs, configurações locais, credenciais e resultados gerados sejam enviados ao Git.

#### `.env.example`

Modelo de variáveis de ambiente do projeto, sem armazenar segredos ou credenciais reais. Pode ser versionado normalmente.

#### `assets/ajuda/`

Prints reais da própria interface, já anotados com setas e círculos numerados indicando onde clicar. São usados pelo diálogo de ajuda (botão **"？ Ajuda"** no painel), que mostra esses três passos — iniciar uma nova execução, escolher o período e clicar em "INICIAR ROBÔ" — de forma navegável (Anterior/Próximo).

Rodando via `python interface.py`, precisa estar na mesma pasta que `interface.py` (ao lado dele), com esses nomes de arquivo exatos. Rodando como `.exe`, essa pasta já vai embutida dentro do executável — ver [Gerando o executável](#gerando-o-executável-exe). Em ambos os casos, se as imagens não forem encontradas, o diálogo apenas mostra um aviso de "imagem não encontrada" no lugar do print, em vez de quebrar.

### 2. Pasta após rodar o `build.py` (gera o `.exe`)

```text
Robo-SIAPE/
├── build/                  # cache intermediário gerado pelo PyInstaller
├── dist/
│   └── RoboSIAPE.exe       # executável final
├── saida/                  # criada automaticamente ao abrir o app, mesmo sem rodar o robô
├── RoboSIAPE.spec          # gerado pelo PyInstaller
├── build.py
├── interface.py
├── robo_siape.py
├── README.md
├── requirements.txt
├── .gitignore
├── .env.example
└── assets/
    └── ajuda/
        ├── passo1_painel.png
        ├── passo2_periodo.png
        └── passo3_iniciar.png
```

`build/`, `dist/`, `RoboSIAPE.spec` e `saida/` não são versionados.

A pasta `saida/` já aparece aqui porque é criada assim que a interface abre (independente de já ter rodado o robô). O caminho é resolvido de forma que a `saida/` sempre nasça na **raiz do projeto** (`Robo-SIAPE/`), nunca dentro de `dist/`, mesmo o `.exe` estando lá dentro — isso vale tanto rodando o `.exe` quanto rodando via `python interface.py`.

A pasta `assets/` continua na raiz porque é a partir dela que o `build.py` embute os prints **dentro do próprio** `RoboSIAPE.exe` (veja a nota sobre `--add-data` em [Gerando o executável](#gerando-o-executável-exe)) — não é preciso copiá-la manualmente para `dist/`.

### 3. Dentro de `saida/` após uma execução

```text
saida/
├── _processamento/
└── 2026-06_20260829_145613/
    ├── base_bruta_2026_06.csv
    ├── base_tratada_2026_06.xlsx
    └── execucao_2026_06.log
```

#### `_processamento/`

Área temporária de trabalho, usada para extrair o ZIP baixado e localizar o CSV de remuneração. É limpa ao final de cada execução, com sucesso ou erro.

#### `AAAA-MM_timestamp/` (uma pasta nova por execução)

Pasta exclusiva daquela execução, nomeada com a competência e um carimbo de data/hora — nunca é reaproveitada nem sobrescrita. Contém:

- `base_bruta_AAAA_MM.csv` — cópia do CSV original extraído do pacote, sem tratamento;
- `base_tratada_AAAA_MM.xlsx` — planilha final, tratada e formatada;
- `execucao_AAAA_MM.log` — log completo daquela execução.

Rodar o robô várias vezes acumula uma pasta por execução dentro de `saida/`.

## Instalação

Recomenda-se utilizar um ambiente virtual.

### Windows

```bash
python -m venv .venv
```

Ative o ambiente:

```bash
.venv\Scripts\activate
```

Instale as dependências:

```bash
pip install -r requirements.txt
```

## Execução (via código-fonte)

Execute a aplicação pela interface:

```bash
python interface.py
```

Depois:

1. informe ou selecione o ano;
2. selecione o mês;
3. inicie a execução;
4. aguarde o download e processamento;
5. consulte a pasta de saída (`saida/`).

## Gerando o executável (.exe)

Com as dependências instaladas (incluindo `pyinstaller`), rode:

```bash
python build.py
```

O executável final fica em `dist/RoboSIAPE.exe`. Basta copiar/mover **só o .exe** para a máquina de destino e executar, não é necessário levar o projeto inteiro, nem ter Python instalado nela.

> **Nota sobre a pasta de dados (`%APPDATA%\RoboSIAPE`):** os arquivos `usuarios.json`, `config.json` e `historico_execucoes.json` não ficam mais na pasta do projeto nem ao lado do `.exe`. Eles são salvos em `%APPDATA%\RoboSIAPE`, a pasta de dados de aplicativos do usuário do Windows — o mesmo padrão usado por programas como Chrome ou Word. Isso é criado automaticamente na primeira execução, tanto rodando via `python interface.py` quanto via `RoboSIAPE.exe` (não é exclusivo do `.exe`). A vantagem é que os dados ficam sempre no mesmo lugar, não importa de onde o executável seja aberto ou para onde seja movido/copiado, e o Windows nunca limpa essa pasta automaticamente (ao contrário de uma pasta temporária).

> **Prints do diálogo de ajuda embutidos no `.exe`:** se a pasta `assets/ajuda/` existir na raiz do projeto no momento do build, o `build.py` a embute dentro do próprio executável (via `--add-data` do PyInstaller). Em tempo de execução, o `.exe` extrai esses arquivos para uma pasta temporária (`sys._MEIPASS`) e o diálogo de ajuda os lê de lá — por isso não é preciso distribuir a pasta `assets/` separadamente junto com o `.exe`. Se `assets/ajuda/` não existir na hora do build, o `build.py` avisa no terminal e gera o `.exe` normalmente, só que sem os prints (o diálogo de ajuda mostra um aviso no lugar deles).

> **Nota sobre antivírus/SmartScreen:** executáveis gerados por PyInstaller sem assinatura digital podem ser sinalizados como suspeitos pelo Windows Defender/SmartScreen na primeira execução. Isso é um falso positivo comum desse tipo de empacotamento, não um problema no código. O `build.py` já toma algumas medidas para reduzir esse risco (sem compressão UPX, com metadados de versão no executável).

## Testes (branch `dev`)

> ⚠️ `teste.py` existe **apenas na branch `dev`** — não faz parte do código
> enviado para `main`/produção, e não deve ser mesclado para lá.

`teste.py` não é um teste unitário: ele roda o **pipeline completo**
(`executar_pipeline_completo`) para **todas** as combinações de ano/mês
conhecidas em `INDICES_ANOS`, uma atrás da outra, com o objetivo de
**estressar o robô** — verificar se ele se comporta bem numa bateria grande
de execuções reais, e não travar/quebrar em algum período específico.

Para rodar (a partir da branch `dev`, na raiz do projeto, junto de
`interface.py` e `robo_siape.py`):

```bash
python teste.py
```

Pontos de atenção:

- Ele baixa **de verdade** cada pacote disponível — pode demorar bastante e
  consumir banda, dependendo de quantos anos/meses existirem.
- Cada execução bem-sucedida gera uma pasta em `saida/`, exatamente como uma
  execução normal feita pela interface. Se quiser, apague `saida/` depois.
- Para restringir o teste a um intervalo específico (em vez de todos os
  anos), edite `ANOS_PARA_TESTAR` no topo do arquivo — ex.:
  `ANOS_PARA_TESTAR = ["2024", "2025", "2026"]`.
- Há uma pausa de `PAUSA_ENTRE_TESTES_SEGUNDOS` (padrão: 3s) entre uma
  execução e outra, para não sobrecarregar o servidor.
- Ao final, imprime um resumo agrupado por status (sucesso, sem dado
  publicado, período inválido, erro inesperado) e salva um log completo em
  `teste_todos_meses_<data>_<hora>.log`, na raiz do projeto.

## Requisitos

- Python 3.10 ou superior;
- conexão com a internet;
- acesso aos endereços de download do Portal da Transparência/CGU;
- espaço em disco suficiente para o pacote baixado e a planilha gerada.

## Segurança e versionamento

Dados produzidos localmente não devem ser versionados.

O `.gitignore` bloqueia arquivos como:

```text
.env
saida/
build/
dist/
*.spec
*.log
*.zip
```

> `usuarios.json`, `config.json` e `historico_execucoes.json` não ficam mais na pasta do projeto — eles são salvos em `%APPDATA%\RoboSIAPE` (ver [Nota sobre a pasta de dados](#gerando-o-executável-exe)), então nunca chegam a ser criados na raiz do repositório. As entradas correspondentes podem ser removidas do `.gitignore` se ainda estiverem lá de uma versão anterior do projeto.

O repositório deve conter apenas código, documentação, arquivos de configuração de exemplo e dependências.

## Decisões técnicas

### Requests em vez de Selenium

O download é realizado diretamente por HTTP utilizando `requests`. Não é necessário abrir navegador para acessar o arquivo, o que elimina a dependência de navegador, driver e automação de interface, além de tornar a execução mais rápida e headless por natureza.

### Arquivos temporários

O download utiliza arquivos temporários e somente promove o arquivo para o nome definitivo depois de verificar se ele é um ZIP válido.

### Processamento em streaming

A geração do Excel utiliza o modo `constant_memory` do `xlsxwriter`, evitando manter toda a planilha em memória — cada linha é escrita direto em disco e descartada da memória assim que não é mais necessária.

O projeto usava originalmente o modo `write_only` do `openpyxl` para o mesmo propósito, mas foi migrado para `xlsxwriter` por desempenho: em testes com a mesma formatação (cabeçalho, cores, larguras de coluna, formato monetário, congelamento de linha e autofiltro), o `xlsxwriter` em modo `constant_memory` processa a planilha entre 7x e 20x mais rápido que o `openpyxl` `write_only`, o que é especialmente relevante para os arquivos de remuneração do SIAPE, que costumam ter centenas de milhares de linhas.

### Caminho da pasta de saída independente de onde o robô é executado

Tanto rodando como script (`python interface.py`) quanto como executável (`RoboSIAPE.exe`), a pasta `saida/` é sempre resolvida em relação à raiz do projeto — nunca em relação à pasta temporária de extração do PyInstaller (`_MEIxxxxxx`) nem à pasta `dist/`. Isso é feito verificando se o processo está "congelado" (`sys.frozen`) e, nesse caso, usando `sys.executable` como referência (subindo um nível a partir de `dist/`) em vez de `__file__`.

### Separação de responsabilidades

O código separa as principais etapas em funções específicas:

```text
download
validação
extração
limpeza
localização do CSV
tratamento
conversão monetária
geração do Excel
logging
limpeza final
```

Isso facilita manutenção, testes e identificação de erros.


## Uso de Inteligência Artificial

Utilizei ferramentas de IA como apoio ao longo do desenvolvimento. A lógica, a estrutura do projeto e as decisões de implementação foram feitas por mim; a IA foi usada como ferramenta de auxílio nos seguintes pontos:

* **Download sem Selenium:** com apoio de IA, avaliei alternativas ao Selenium e migrei o download para requisições HTTP diretas (requests), eliminando a dependência de navegador e tornando a execução mais rápida e headless por natureza.
* **Empacotamento em executável:** apoio de IA para configurar o build.py com PyInstaller, incluindo ajustes para reduzir falsos positivos de antivírus/SmartScreen (remoção de compressão UPX, inclusão de metadados de versão) e correção do caminho da pasta de saída para funcionar corretamente tanto rodando como script quanto como .exe.
* **Git e GitHub**: usei IA (e vídeos no YouTube) como apoio para relembrar comandos e o fluxo de criação/organização do repositório, pois fazia tempo que eu não usava.
* **Debug e revisão do código e do README:** apoio de IA para revisar o código, encontrar inconsistências entre a documentação e o comportamento real do robô, e organizar a escrita deste README.
* **PyQt6:** com apoio de IA e vídeos explicativos no YouTube, aprendi a implementar o PyQt6 no projeto para deixar a interface gráfica mais limpa e com mais detalhes visuais.

Todo o código foi revisado, testado e é de meu conhecimento, posso explicar qualquer trecho da implementação.



## Histórico de commits

Os commits devem representar etapas reais do desenvolvimento do projeto.

Exemplos:

```text
feat: adiciona interface de login e seleção de período

feat: implementa automação de download do Portal da Transparência

feat: adiciona processamento da base de remuneração

feat: adiciona geração das planilhas Excel

feat: adiciona tratamento e formatação dos dados

feat: adiciona logging e retentativas de download

docs: adiciona README e documentação do projeto

chore: adiciona requirements e gitignore
```

Evite concentrar todo o desenvolvimento em um único commit como:

```text
projeto finalizado
```

Manter commits separados facilita acompanhar a evolução do projeto, identificar alterações e entender as diferentes etapas do desenvolvimento.

## Limitações atuais

- O pacote utilizado é `Servidores_SIAPE`.
- O sistema depende da disponibilidade dos arquivos para a competência escolhida.
- Para 2026, o código considera atualmente disponíveis apenas janeiro a junho.
- Alterações na estrutura dos arquivos disponibilizados pelo Portal da Transparência podem exigir ajustes no tratamento.
- O tratamento monetário é direcionado à estrutura da planilha SIAPE utilizada no projeto.
- O executável não é assinado digitalmente, podendo ser sinalizado por antivírus/SmartScreen na primeira execução em uma máquina nova.

## Licença

Este projeto foi desenvolvido para fins de estudo, avaliação técnica e automação do tratamento dos dados públicos do SIAPE.