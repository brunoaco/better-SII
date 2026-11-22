# better-SII

**Tampermonkey userscripts and a Python tool that make Chile's tax authority website ([www.sii.cl](https://www.sii.cl)) faster and easier to use.**

Accountants who manage many companies on the SII portal repeat the same clicks every day: switching between companies, downloading documents one by one, and fixing printouts. These scripts add the missing features directly to the SII pages, so a task that used to take dozens of clicks runs in one.

> Scripts and on-screen messages are in Spanish, because the SII portal and its users are Chilean.

---

## Scripts

| Script | SII page | What it does |
|---|---|---|
| **SII - Ordenar Empresas** (v2.1) | Company selector (`mipeSelEmpresa.cgi`) | Sorts the company drop-down list and adds a toggle between alphabetical order (name + RUT) and numeric order (RUT + name). The choice is remembered between visits. |
| **SII - Compras Pendientes** (v2.3) | Purchase and sales register (`consdcvinternetui`) | Checks pending purchase documents by company RUT and acknowledges them in bulk. Rules can be set per RUT and per document type (e.g. only credit notes), and a master switch turns it into a count-only mode. |
| **SII - Descarga CSV Mensual (Compra y Venta)** (v1.0) | Purchase and sales register (`consdcvinternetui`) | Loops through the 12 months of the selected year and downloads the purchase and sales CSV files for the selected company. |
| **SII - Descarga de PDFs** (v1.0) | Issued documents list (`mipeAdminDocsEmi.cgi`) | Downloads the PDFs of all issued documents in the list, after a confirmation window that shows the active company and the number of documents. |
| **SII BHE - Descarga Masiva PDFs** (v1.4) | Electronic fee receipts reports (BHE) | Bulk-downloads fee receipt (*boletas de honorarios*) PDFs from the annual, monthly or daily report, with an on-page control panel. |
| **SII F29 - Botón IMPRESIÓN** (v1.0) | Form 29 draft (`rfiInternet/consulta`) | Adds a print button that prints the Form 29 draft cleanly, without blank pages or cut-off columns. Paper size, orientation and margins are configurable. |

## Python tool

**`extraer_facturas_sii.py`**: once the received-document PDFs have been downloaded, this script reads them in bulk and builds an Excel report with each document's header (issuer, receiver, folio, type, date) and line items.

```bash
pip install pdfplumber openpyxl pandas
python extraer_facturas_sii.py <folder_with_pdfs> [output.xlsx]
```

## Installation

1. Install the [Tampermonkey](https://www.tampermonkey.net/) extension in your browser.
2. Open the script you want in this repository and copy its contents.
3. In Tampermonkey, choose **Create a new script**, paste the code and save.
4. Log in to SII and open the matching page. The script runs automatically.

## Tech

- JavaScript userscripts (DOM manipulation, AngularJS scope access, jQuery DataTables, `GM_download`)
- Python 3 (`pdfplumber`, `pandas`, `openpyxl`)
- Built with the help of AI tools (Claude.ai)

## Disclaimer

This is an independent project, not affiliated with or endorsed by the Servicio de Impuestos Internos (SII). The scripts only act on pages the user has already logged in to and run in the user's own browser. Changes to the SII website may break them.

## License

[Apache-2.0](LICENSE)
