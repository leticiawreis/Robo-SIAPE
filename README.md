# Robô SIAPE

Automação desenvolvida em Python para consultar, baixar, processar e formatar dados de remuneração de servidores públicos disponibilizados no Portal da Transparência.

O sistema possui uma interface gráfica para autenticação do usuário, seleção do período desejado e acompanhamento da execução em tempo real. A automação utiliza Selenium para acessar o Portal da Transparência, realizar o download da base correspondente ao período selecionado e gerar arquivos Excel organizados.

## Funcionalidades

* Login e cadastro de usuários localmente.
* Armazenamento das credenciais em arquivo local com senha protegida por hash.
* Seleção de ano e mês pela interface gráfica.
* Validação do período selecionado.
* Bloqueio de períodos futuros.
* Validação dos meses disponíveis para cada ano.
* Automação do acesso ao Portal da Transparência utilizando Selenium.
* Download automático da base de dados.
* Monitoramento do download até sua conclusão.
* Retentativa automática em caso de falha.
* Detecção de CAPTCHA.
* Pausa da automação para resolução manual do CAPTCHA quando necessário.
* Extração dos arquivos baixados.
* Identificação automática da base de remuneração.
* Remoção de linhas e colunas completamente vazias.
* Preservação de campos sem informação.
* Geração da planilha bruta.
* Geração da planilha tratada e formatada.
* Formatação de valores monetários.
* Ajuste automático da largura das colunas.
* Aplicação de filtros no cabeçalho.
* Congelamento da primeira linha.
* Registro detalhado da execução em arquivo de log.
* Limpeza dos arquivos temporários ao final do processamento.

## Tecnologias utilizadas

* Python
* PyQt6
* Selenium
* openpyxl
* Google Chrome
* JSON
* Logging

## Estrutura do projeto

```text
.
├── interface.py
├── robo_siape.py
├── requirements.txt
├── README.md
├── .gitignore
├── .env.example
└── saida/
```

### Responsabilidade dos arquivos

#### `interface.py`

Responsável pela interface gráfica do sistema.

Inclui:

* Tela de login.
* Cadastro de usuários.
* Seleção de ano e mês.
* Botão para iniciar o robô.
* Exibição do andamento da execução.
* Exibição dos logs.
* Comunicação com o processo de automação.

A interface foi desenvolvida utilizando PyQt6 e possui tema escuro com detalhes em amarelo.

#### `robo_siape.py`

Contém a lógica principal da automação.

É responsável por:

* Validar o período selecionado.
* Configurar o navegador.
* Acessar o Portal da Transparência.
* Navegar até os dados de servidores.
* Detectar e aguardar CAPTCHA quando necessário.
* Realizar o download.
* Controlar as tentativas de download.
* Extrair os arquivos.
* Localizar a base de remuneração.
* Tratar os dados.
* Criar as planilhas Excel.
* Registrar os eventos no log.
* Remover arquivos temporários.

#### `requirements.txt`

Contém as dependências externas utilizadas pelo projeto e suas respectivas versões.

#### `.gitignore`

Define arquivos e pastas que não devem ser enviados para o Git, como:

* ambientes virtuais;
* arquivos temporários;
* cache;
* `.env`;
* `usuarios.json`;
* arquivos gerados em `saida/`.

#### `.env.example`

Arquivo de referência para variáveis de ambiente, sem armazenar informações sensíveis.

#### `saida/`

Pasta destinada aos resultados gerados pelo robô.

Os arquivos intermediários utilizados durante o processamento ficam separados dos resultados finais e são removidos ao término da execução.

## Instalação

É recomendado utilizar um ambiente virtual para instalar as dependências do projeto.

```bash
python -m venv .venv
```

No Windows:

```bash
.venv\Scripts\activate
```

Depois, instale as dependências:

```bash
pip install -r requirements.txt
```

Também é necessário possuir o Google Chrome instalado na máquina.

## Execução

Para iniciar o sistema:

```bash
python interface.py
```

Na primeira utilização, é necessário cadastrar um usuário.

Depois:

1. Faça login.
2. Selecione o ano desejado.
3. Selecione o mês.
4. Clique em `Iniciar Robô`.
5. Aguarde o navegador ser aberto.
6. Caso o Portal da Transparência apresente um CAPTCHA, faça a resolução manualmente.
7. Aguarde o término do download e processamento.
8. Consulte os arquivos gerados na pasta `saida/`.

## Fluxo da automação

O processamento ocorre aproximadamente da seguinte forma:

```text
Login
  ↓
Seleção de ano e mês
  ↓
Validação do período
  ↓
Abertura do Portal da Transparência
  ↓
Navegação até os dados de servidores
  ↓
Verificação de CAPTCHA
  ↓
Download da base
  ↓
Validação do download
  ↓
Extração dos arquivos
  ↓
Localização da base de remuneração
  ↓
Criação da base bruta
  ↓
Tratamento dos dados
  ↓
Criação da base formatada
  ↓
Geração do log
  ↓
Limpeza dos arquivos temporários
```

## Validação do período

Antes de iniciar o processamento, o sistema verifica se o período selecionado é válido.

O robô não permite:

* períodos posteriores ao mês atual;
* meses que não estejam disponíveis para o ano selecionado;
* combinações de ano e mês incompatíveis com os dados disponibilizados pelo Portal.

Dessa forma, evita-se iniciar uma automação para uma base que ainda não está disponível.

## Download

O download é acompanhado automaticamente pelo sistema.

O arquivo somente é considerado concluído quando:

* um arquivo `.zip` é encontrado;
* o download não está mais em andamento;
* não existe mais um arquivo `.crdownload` correspondente ao download em execução.

Caso o download apresente uma falha, o sistema realiza novas tentativas automaticamente, respeitando o limite configurado de tentativas.

## CAPTCHA

O CAPTCHA não é automatizado.

Quando o Portal da Transparência solicita a verificação, o robô identifica a situação e pausa o processamento para que o usuário faça a resolução manualmente no navegador.

Após a resolução, a execução pode continuar normalmente.

### Sobre o modo headless (bônus não implementado)

O modo headless (navegador invisível) foi avaliado, mas não foi implementado propositalmente: o Portal da Transparência apresenta CAPTCHA em pontos do fluxo, e sua resolução manual depende de o navegador estar visível na tela. Rodar em headless exigiria contratar um serviço de resolução automática de CAPTCHA (ex. 2Captcha, Anti-Captcha), o que foge do escopo do case e introduziria uma dependência paga. Optou-se por manter o navegador visível para permitir a intervenção manual quando necessário.

## Processamento dos arquivos

Após o download, o arquivo compactado é extraído em uma área de processamento.

O sistema localiza automaticamente o arquivo correspondente à base de remuneração e utiliza seus dados para gerar os resultados.

Os arquivos intermediários utilizados durante essa etapa não permanecem como resultados finais.

## Base bruta

A base bruta representa os dados extraídos do Portal da Transparência antes do tratamento de organização da planilha.

Ela é mantida como referência para permitir uma comparação entre os dados originais e a versão final processada.

Exemplo:

```text
saida/

├── base_bruta_2025_03.csv
├── base_tratada_2025_03.xlsx
└── execucao_2025_03.log
```

## Base tratada

A base tratada é gerada a partir da base bruta e recebe o tratamento necessário para facilitar sua utilização.

Durante o processamento:

* linhas completamente vazias são removidas;
* colunas completamente vazias são removidas;
* valores ausentes são mantidos vazios;
* os dados são organizados em uma planilha Excel;
* o cabeçalho recebe formatação;
* filtros são adicionados;
* a primeira linha é congelada;
* as larguras das colunas são ajustadas;
* valores monetários são formatados quando identificados.

## Dados ausentes

Campos sem informação permanecem vazios.

O sistema não transforma automaticamente valores ausentes em `0`, evitando alterar o significado original dos dados.

## Linhas e colunas vazias

Uma linha é removida quando todos os seus campos estão vazios.

Uma coluna também é removida quando não possui nenhum valor preenchido em seus registros.

Dessa forma, a planilha final não fica com grandes áreas completamente vazias provenientes da base original.

## Formatação monetária

O sistema identifica colunas relacionadas a valores financeiros por meio do nome da coluna.

Entre os termos considerados estão, por exemplo:

* remuneração;
* vencimento;
* provento;
* desconto;
* líquido;
* bruto;
* valor;
* gratificação;
* auxílio;
* indenização.

Quando o conteúdo da coluna é numérico, os valores podem receber formatação monetária no Excel.

## Arquivos gerados

Para uma execução referente a março de 2025, por exemplo:

```text
saida/

├── _processamento/
├── base_bruta_2025_03.csv
├── base_tratada_2025_03.xlsx
└── execucao_2025_03.log
```

### `_processamento/`

Pasta utilizada durante a execução para armazenar os arquivos baixados e extraídos.

É uma área temporária e não representa o resultado final do robô.

Após o processamento, os arquivos temporários são removidos.

### `base_bruta_AAAA_MM.xlsx`

Planilha contendo a base extraída originalmente para o período selecionado, servindo como referência antes do tratamento.

### `base_tratada_AAAA_MM.xlsx`

Planilha final processada pelo robô.

Possui:

* cabeçalho formatado;
* filtros;
* primeira linha congelada;
* larguras de colunas ajustadas;
* tratamento das linhas e colunas vazias;
* formatação monetária quando aplicável.

### `execucao_AAAA_MM.log`

Arquivo que registra os acontecimentos da execução.

Cada registro possui data e hora e utiliza níveis de log como:

* `INFO`: informações normais do processamento;
* `WARNING`: situações recuperáveis ou que exigiram atenção;
* `ERROR`: erros que afetaram o processamento.

O sistema utiliza o módulo `logging` para registrar a execução, em vez de utilizar `print()` como mecanismo principal de log.

## Limpeza de arquivos temporários

Os arquivos utilizados apenas durante o processamento são mantidos na pasta de processamento e removidos ao final da execução.

A limpeza também é realizada quando ocorre uma falha, evitando o acúmulo de arquivos temporários de execuções anteriores.

Os resultados finais permanecem disponíveis na pasta `saida/`.

## Usuários e segurança

Os usuários cadastrados são armazenados localmente no arquivo:

```text
usuarios.json
```

As senhas não são armazenadas em texto puro. O sistema utiliza hash para armazenar as credenciais.

O arquivo `usuarios.json` não deve ser versionado no Git.

Por esse motivo, ele deve estar incluído no `.gitignore`.

## Dados sensíveis e Git

O projeto possui arquivos que não devem ser enviados para o repositório.

Entre eles:

```text
.env
usuarios.json
.venv/
__pycache__/
saida/
arquivos temporários
arquivos de cache
```

O `.env.example` pode ser versionado, pois serve apenas como modelo e não deve conter segredos ou credenciais reais.

## Tratamento de erros

O robô possui mecanismos para evitar que falhas pontuais encerrem imediatamente todo o processo.

Entre os mecanismos utilizados estão:

* retentativas de download;
* monitoramento do estado do download;
* verificação da existência dos arquivos;
* identificação de CAPTCHA;
* limpeza de arquivos temporários;
* registro de erros no log.

Quando uma falha não pode ser recuperada, ela é registrada como `ERROR` e o usuário é informado pela interface.

## Uso de Inteligência Artificial

Utilizei ferramentas de IA como apoio ao longo do desenvolvimento. A lógica, a estrutura do projeto e as decisões de implementação foram feitas por mim; a IA foi usada como ferramenta de auxílio nos seguintes pontos:

* **CAPTCHA:** confirmei com apoio de IA que não seria viável automatizar a resolução do CAPTCHA do Portal da Transparência sem contratar uma API paga de resolução; por isso essa etapa permanece manual, como documentado acima.
* **Git e GitHub:** usei IA (e vídeos no YouTube) como apoio para relembrar comandos e o fluxo de criação/organização do repositório, pois fazia tempo que eu não usava.
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

## Observações

O robô depende da disponibilidade e do funcionamento do Portal da Transparência.

Alterações na estrutura do site podem exigir ajustes na automação Selenium.

O CAPTCHA permanece como uma etapa manual sempre que for apresentado pelo Portal.

Os arquivos gerados pelo processamento não devem ser enviados ao repositório Git.