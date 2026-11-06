# Power BI Dashboards: Public-Data Analysis Portfolio

A collection of Business Intelligence projects built with Power BI and Python on publicly available Brazilian data. Each folder is a self-contained project with its data, a Power BI report (`.pbix`) and a PDF export of the dashboard.

> The reports and data are in Brazilian Portuguese.

## Projects

| Project | What it analyzes | Files |
| --- | --- | --- |
| [`mais_medicos_2024`](mais_medicos_2024/) | The *Mais Médicos* (More Doctors) federal program: its reach and results across Brazilian regions, states and municipalities, using public health data. | [Report (.pbix)](mais_medicos_2024/bimaismedicos.pbix) · [PDF export](mais_medicos_2024/bimaismedicos.pdf) · [`data/`](mais_medicos_2024/data/) |
| [`concurso_buritis_2024`](concurso_buritis_2024/) | The 2024 municipal public-service exam of Buritis: applicants and the city's public servants, comparing applicants who are already municipal employees with those who are not. | [Report (.pbix)](concurso_buritis_2024/PAINEL_CONCURSO_BURITIS_2024.pbix) · [PDF export](concurso_buritis_2024/PAINEL_CONCURSO_BURITIS_2024.pdf) · [`data/`](concurso_buritis_2024/data/) · [`scripts/`](concurso_buritis_2024/scripts/) |

A third project is planned.

<!-- TODO: confirmar — a descrição do repositório no GitHub cita ETL com Pentaho e SQL, mas não há arquivos Pentaho ou SQL no repositório. Confirmar se esse ETL existe em outro lugar ou ajustar a descrição. -->
<!-- TODO: confirmar — indicar a fonte (portal/órgão) e a data de extração dos dados de cada projeto. -->

## Tech stack

- **Power BI Desktop** for the data model, visuals and dashboards (`.pbix` files)
- **Python** with `pandas` for data preparation (`concurso_buritis_2024/scripts/data_excel.py`)
- **Source data** as CSV, XLS/XLSX and PDF files
- **DAX** for measures in the reports

<!-- TODO: confirmar — o uso de DAX foi mantido do README original; os .pbix são binários e não foi possível verificá-lo aqui. -->

## Skills demonstrated

- Data analysis on large public datasets
- Dashboard design and data visualization
- Business Intelligence: turning raw data into decision-oriented insights
- Data preparation with Python

## Getting started

### View the dashboards

- Open the PDF export of any project for a quick look, or
- Open the `.pbix` file in [Power BI Desktop](https://powerbi.microsoft.com/desktop/) (Windows) for the interactive report.

### Run the data-preparation script

`concurso_buritis_2024/scripts/data_excel.py` merges every sheet of `buritis_concurso2024.xlsx` into a single sheet and saves it as `merged_buritis_concurso2024.xlsx`. It reads the input file and writes the output in the **current working directory**, so run it from the folder that contains the spreadsheet:

```bash
pip install pandas openpyxl
cd concurso_buritis_2024/data
python ../scripts/data_excel.py
```

## Repository structure

```
mais_medicos_2024/
├── bimaismedicos.pbix        Power BI report
├── bimaismedicos.pdf         Dashboard export
├── data/                     Source CSV/XLS files
└── img/                      Logos used in the report
concurso_buritis_2024/
├── PAINEL_CONCURSO_BURITIS_2024.pbix
├── PAINEL_CONCURSO_BURITIS_2024.pdf
├── data/                     Source spreadsheets, PDFs and merged output
├── img/
└── scripts/data_excel.py     Merges workbook sheets with pandas
```

## Contact

- <goncalvesdani54@gmail.com>
- [LinkedIn](https://www.linkedin.com/in/danielle-de-paula18/)
