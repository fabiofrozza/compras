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

<p align="center">
  <img src="https://raw.githubusercontent.com/fabiofrozza/compras/main/docs/media/home.png" alt="Tela inicial" width="800">
</p>

## Funcionalidades

<p align="center">
  <img src="https://raw.githubusercontent.com/fabiofrozza/compras/main/docs/media/app.png" alt="Interface da aplicação" width="800">
</p>

* **Atas** — geração de Atas de Registro de Preços com base nos relatórios de cadastramento dos fornecedores (obtidos no SICAF) e na lista de itens do Termo de Referência (TR).
* **CATMAT** — verificação dos CATMATs e das respectivas margens de preferência (se houver) com base na lista de itens do TR.
* **Fornecedores** — geração de arquivo para atualização dos dados bancários dos fornecedores.
* **Importação** — geração de arquivos para importação e criação dos pedidos de compras, resumo das importações para controle dos processos e produção de relatórios gerenciais.
* **Mapas** — geração de listas de itens para licitação com base nos mapas de licitação de processos anteriores.
* **Power BI** — geração de arquivo de dados para atualização dos painéis do Power BI.
* **SNEs** — organização das certidões negativas dos fornecedores e da documentação para emissão de notas de empenho.
* **Instalação** — gerenciamento da instalação do R.

## Instalação e personalização

1. [Baixe o conteúdo deste repositório](https://github.com/fabiofrozza/compras/archive/refs/heads/main.zip) e extraia na pasta desejada.

1. Renomeie o arquivo `.env-MODELO` para `.env` e edite-o no Bloco de Notas (Windows) ou equivalente (Linux), seguindo as instruções no seu interior.

1. Substitua as imagens em `public/img/company.png` e `public/img/department.png` pelas da sua empresa.

1. **Somente no Linux:** baixe o executável do Node.js (v20 LTS ou superior) para Linux em [nodejs.org](https://nodejs.org/pt/download) e salve-o como `bin/node`. Em seguida, conceda permissão de execução ao iniciador e ao binário:
   ```bash
   chmod +x start-linux.sh bin/node
   ```

## Uso

- **Windows:** execute o arquivo `start.cmd`.
- **Linux:** execute `./start-linux.sh` no terminal.

O servidor será iniciado e o navegador abrirá automaticamente em `http://localhost:3000`.

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