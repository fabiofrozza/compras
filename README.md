# Compras [![pt-br](https://img.shields.io/badge/lang-pt--br-yellow?style=plastic)](https://github.com/fabiofrozza/compras/tree/main/README.md)

[![en-us](https://img.shields.io/badge/lang-en--us-blue?style=plastic)](https://github.com/fabiofrozza/compras/tree/main/README.en-us.md)

Alguns scripts para auxiliar meus colegas nas atividades relacionadas a compras públicas.

## Conteúdo

* [`importacao`](./importacao): geração de arquivos para importação e criação dos pedidos de compras, resumo das importações para controle dos processos e produção de relatórios gerenciais.
* [`mapas`](./mapas): geração de listas de itens para licitação com base nos mapas de licitação de processos anteriores.
* [`catmat`](./catmat): verificação dos CATMATs e das respectivas margens de preferência (se houver) com base na lista de itens do Termo de Referência.
* [`atas`](./atas): geração de Atas de Registro de Preços com base nos relatórios de cadastramento dos fornecedores (obtidos no SICAF).
* [`fornecedores`](./fornecedores): geração de arquivo para atualização dos dados bancários dos fornecedores.
* [`powerbi`](./powerbi): geração de arquivo de dados para atualização dos painéis do Power BI.
* [`primeiro_uso`](./primeiro_uso): instalação do programa R e (opcionalmente) atualização  do PowerShell, utilizados pelos scripts.

## Iniciando

### Instalação

* [Baixe o conteúdo deste repositório](https://github.com/fabiofrozza/compras/archive/refs/heads/main.zip) na pasta desejada.
* Mantenha a mesma estrutura de pastas.
* Baixe a última versão do [R](https://cran.r-project.org/bin/windows/base/) e salve o arquivo na pasta `primeiro_uso`.
* Execute o script `primeiro_uso/primeiro_uso.exe`.
* _(Opcional, mas altamente recomendável)_ Baixe a última versão do [PowerShell](https://learn.microsoft.com/en-us/powershell/scripting/install/installing-powershell-on-windows?view=powershell-7.5#installing-the-zip-package) (arquivo .zip) e salve o arquivo na pasta `primeiro_uso`.
* Execute o script `primeiro_uso/opcional_powershell.exe`.

### Personalização

* Na pasta `_common`, renomeie o arquivo `.Renviron-MODELO` para `.Renviron`.
* Edite-o (no Bloco de Notas ou similar) seguindo as instruções do arquivo.
* Na pasta `_common/images`, substitua as imagens `company.png`, `department.png` e `lists_page.png` pelas da sua empresa.

Pronto! Os scripts estão prontos para serem usados!

## Algumas observações

Estes scripts são parte de um projeto pessoal que me rendeu muitas alegrias. 

Partindo de um conhecimento básico de R, iniciei, por necessidade, a criação do [`importacao`](./importacao) para agilizar a geração dos pedidos. Aos poucos, conforme o script funcionava, foram surgindo novas ideias e novos scripts. 

Assim, busquei alternativas já disponíveis aos meus colegas (como o PowerShell, pré-instalado no Windows, e, obviamente, os batches do CMD), facilitando ao máximo a utilização dos scripts por quem não tem nenhum conhecimento prévio de programação. Por este motivo, a estrutura é um pouco diferente do que seria um projeto padrão ou um pacote R. Contudo, os benefícios justificam esta decisão.

## Reconhecimento

Agradeço aos meus colegas pela paciência em serem meus _"testadores"_ e às suas ideias e sugestões. É para vocês tudo isto, o que fiz com muito carinho (como podem ver pelas mensagens engraçadinhas pra fechar as janelas 😄).

Obrigado a https://ascii.co.uk/ e https://www.asciiart.eu/ pelas artes em ASCII.

Obrigado a [Jonatas Emidio](https://github.com/jonatasemidio/multilanguage-readme-pattern) pelo template de readme multilínguas.
