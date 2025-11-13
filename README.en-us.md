# Purchases [![pt-br](https://img.shields.io/badge/lang-pt--br-yellow?style=plastic)](https://github.com/fabiofrozza/compras/tree/main/README.md)

[![en-us](https://img.shields.io/badge/lang-en--us-blue?style=plastic)](https://github.com/fabiofrozza/compras/tree/main/README.en-us.md)

Some scripts to assist my colleagues in activities related to public procurement.

## Contents

* [`importacao`](./importacao): generation of files for importing and creating purchase orders, summary of imports for process control, and production of management reports.
* [`mapas`](./mapas): generation of item lists for bidding based on bidding maps from previous processes.
* [`catmat`](./catmat): verification of CATMATs and their respective preference margins (if any) based on the list of items in the Terms of Reference.
* [`atas`](./atas): generation of Price Registration Minutes based on supplier registration reports (obtained from SICAF).
* [`fornecedores`](./fornecedores): generation of a file to update suppliers' bank details.
* [`powerbi`](./powerbi): generation of a data file to update Power BI dashboards.
* [`primeiro_uso`](./primeiro_uso): installation of the R program and (optionally) updating PowerShell, used by the scripts.

## Getting Started

### Installation

* [Download the content of this repository](https://github.com/fabiofrozza/compras/archive/refs/heads/main.zip) to the desired folder.
* Keep the same folder structure.
* Download the latest version of [R](https://cran.r-project.org/bin/windows/base/) and save the file in the `primeiro_uso` folder.
* Run the `primeiro_uso/primeiro_uso.exe` script.
* _(Optional, but highly recommended)_ Download the latest version of [PowerShell](https://learn.microsoft.com/en-us/powershell/scripting/install/installing-powershell-on-windows?view=powershell-7.5#installing-the-zip-package) (.zip file) and save the file in the `primeiro_uso` folder.
* Run the `primeiro_uso/opcional_powershell.exe` script.

### Customization

* In the `_common` folder, rename the `.Renviron-MODELO` file to `.Renviron`.
* Edit it (in Notepad or similar) following the instructions in the file.
* In the `_common/images` folder, replace the `company.png`, `department.png`, and `lists_page.png` images with your company's.

Done! The scripts are ready to be used!

## Some observations

These scripts are part of a personal project that has brought me much joy.

Starting from a basic knowledge of R, I began, out of necessity, to create [`importacao`](./importacao) to speed up the generation of orders. Gradually, as the script worked, new ideas and new scripts emerged.

Thus, I looked for alternatives already available to my colleagues (like PowerShell, pre-installed on Windows, and, obviously, CMD batches), making it as easy as possible for those with no prior programming knowledge to use the scripts. For this reason, the structure is a bit different from what a standard project or an R package would be. However, the benefits justify this decision.

## Development

[Here](./_common/README.md) is some information about the code development.

## Acknowledgments

I thank my colleagues for their patience in being my _"testers"_ and for their ideas and suggestions. This is all for you, which I did with great affection (as you can see from the funny messages to close the windows 😄).

Thanks to https://www.flaticon.com for the icons and images.

Thanks to https://ascii.co.uk/ and https://www.asciiart.eu/ for the ASCII art.

Thanks to [Jonatas Emidio](https://github.com/jonatasemidio/multilanguage-readme-pattern) for the multilingual readme template.