# Compras [![en-us](https://img.shields.io/badge/lang-en--us-blue?style=plastic)](https://github.com/fabiofrozza/compras/tree/main/README.en-us.md) [![pt-br](https://img.shields.io/badge/lang-pt--br-yellow?style=plastic)](https://github.com/fabiofrozza/compras/tree/main/README.md) 

[![GitHub Release](https://img.shields.io/github/v/release/fabiofrozza/compras)](https://github.com/fabiofrozza/compras/releases)

<p align="center">
<a href="#Features">Features</a> &nbsp;&bull;&nbsp;
<a href="#Installation-and-customization">Installation and customization</a> &nbsp;&bull;&nbsp;
<a href="#Usage">Usage</a> &nbsp;&bull;&nbsp;
<a href="#Some-observations">Some observations</a> &nbsp;&bull;&nbsp;
<a href="#Acknowledgments">Acknowledgments</a> &nbsp;&bull;&nbsp;
<a href="#License">License</a>
</p>

A local web application that runs scripts to assist my colleagues in activities related to public procurement.

## Features

* **Atas** — generation of Price Registration Minutes based on supplier registration reports (obtained from SICAF) and the list of items from the Terms of Reference (TR).
* **CATMAT** — verification of CATMATs and their respective preference margins (if any) based on the list of items from the TR.
* **Fornecedores** — generation of a file to update suppliers' bank details.
* **Importação** — generation of files for importing and creating purchase orders, summary of imports for process control, and production of management reports.
* **Mapas** — generation of item lists for bidding based on bidding maps from previous processes.
* **Power BI** — generation of a data file to update Power BI dashboards.
* **Instalação** — R installation management.

## Installation and customization

1. [Download the content of this repository](https://github.com/fabiofrozza/compras/archive/refs/heads/main.zip) and extract it to the desired folder.

1. Rename the following files and edit them in Notepad, following the instructions inside:
   * in the root folder: `.env-MODELO` to `.env`
   * in the `scripts/_common` folder: `.Renviron-MODELO` to `.Renviron`

1. Replace the following images with your company's:
   * `public/img/company.png`
   * `public/img/department.png`
   * `scripts/_common/images/company.png`
   * `scripts/_common/images/department.png`

## Usage

Run the `start.cmd` file. The server will start and the browser will automatically open at `http://localhost:3000`.

The interface has tabs for each feature. Simply fill in the required fields and follow the script execution in real time through the integrated console.

## Some observations

This project is the result of a personal initiative that has brought me much joy.

Starting from a basic knowledge of R, I began creating the import scripts to speed up order generation. Gradually, as the scripts worked, new ideas and new features emerged, until the project gained its own web interface, replacing the old command-line scripts with a more accessible and user-friendly experience.

The graphical interface was built with **Bootstrap 5** and **vanilla JavaScript**. Considering that this is a locally-run application, frameworks like React or Vue would add unnecessary complexity and overhead.


## Acknowledgments

I thank my colleagues for their patience in being my _"testers"_ and for their ideas and suggestions. This is all for you, which I made with great affection (as you can see from the funny messages to close the windows 😄).

Thanks to https://ascii.co.uk/ and https://www.asciiart.eu/ for the ASCII art.

Thanks to [Jonatas Emidio](https://github.com/jonatasemidio/multilanguage-readme-pattern) for the multilingual readme template.

Other credits are alongside the code used.

## License

This project is open source and licensed under the MIT license.
