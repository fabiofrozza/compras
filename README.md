# Compras [![pt-br](https://img.shields.io/badge/lang-pt--br-yellow?style=plastic)](https://github.com/fabiofrozza/compras/tree/main/README.md) [![en-us](https://img.shields.io/badge/lang-en--us-blue?style=plastic)](https://github.com/fabiofrozza/compras/tree/main/README.en-us.md)

[![GitHub Release](https://img.shields.io/github/v/release/fabiofrozza/compras)](https://github.com/fabiofrozza/compras/releases)

<p align="center">
<a href="#Funcionalidades">Funcionalidades</a> &nbsp;&bull;&nbsp;
<a href="#Instalação-e-personalização">Instalação e personalização</a> &nbsp;&bull;&nbsp;
<a href="#Uso">Uso</a> &nbsp;&bull;&nbsp;
<a href="#Algumas-observações">Algumas observações</a> &nbsp;&bull;&nbsp;
<a href="#Reconhecimento">Reconhecimento</a> &nbsp;&bull;&nbsp;
<a href="#Licença">Licença</a>
</p>

Aplicação web local que permite execução de scripts para auxiliar meus colegas nas atividades relacionadas a compras públicas.

## Funcionalidades

* **Atas** — geração de Atas de Registro de Preços com base nos relatórios de cadastramento dos fornecedores (obtidos no SICAF) e na lista de itens do Termo de Referência (TR).
* **CATMAT** — verificação dos CATMATs e das respectivas margens de preferência (se houver) com base na lista de itens do TR.
* **Fornecedores** — geração de arquivo para atualização dos dados bancários dos fornecedores.
* **Importação** — geração de arquivos para importação e criação dos pedidos de compras, resumo das importações para controle dos processos e produção de relatórios gerenciais.
* **Mapas** — geração de listas de itens para licitação com base nos mapas de licitação de processos anteriores.
* **Power BI** — geração de arquivo de dados para atualização dos painéis do Power BI.
* **Instalação** — gerenciamento da instalação do R.

## Instalação e personalização

1. [Baixe o conteúdo deste repositório](https://github.com/fabiofrozza/compras/archive/refs/heads/main.zip) e extraia na pasta desejada.

1. Renomeie os seguintes arquivos e edite-os no Bloco de Notas, seguindo as instruções no seu interior:
   * na pasta raiz: `.env-MODELO` para `.env`
   * na pasta `scripts/_common`: `.Renviron-MODELO` para `.Renviron`

1. Substitua as seguintes imagens pelas da sua empresa:
   * `public/img/company.png`
   * `public/img/department.png`
   * `scripts/_common/images/company.png`
   * `scripts/_common/images/department.png`

## Uso

Execute o arquivo `start.cmd`. O servidor será iniciado e o navegador abrirá automaticamente em `http://localhost:3000`.

A interface apresenta abas para cada funcionalidade. Basta preencher os campos necessários e acompanhar a execução dos scripts em tempo real pelo console integrado.

## Algumas observações

Este projeto é fruto de uma iniciativa pessoal que me rendeu muitas alegrias.

Partindo de um conhecimento básico de R, iniciei a criação dos scripts de importação para agilizar a geração dos pedidos. Aos poucos, conforme os scripts funcionavam, foram surgindo novas ideias e novas funcionalidades, até que o projeto ganhou uma interface web própria, substituindo os antigos scripts de linha de comando por uma experiência mais acessível e amigável.

A interface gráfica foi construída com **Bootstrap 5** e **JavaScript vanilla**. Considerando que se trata de uma aplicação de execução local, frameworks como React ou Vue agregariam complexidade e peso desnecessários.


## Reconhecimento

Agradeço aos meus colegas pela paciência em serem meus _"testadores"_ e às suas ideias e sugestões. É para vocês tudo isto, o que fiz com muito carinho (como podem ver pelas mensagens engraçadinhas pra fechar as janelas 😄).

Obrigado a https://ascii.co.uk/ e https://www.asciiart.eu/ pelas artes em ASCII.

Obrigado a [Jonatas Emidio](https://github.com/jonatasemidio/multilanguage-readme-pattern) pelo template de readme multilínguas.

Demais créditos estão junto ao código utilizado.

## Licença

Este projeto é de código aberto e está licenciado sob a licença MIT.