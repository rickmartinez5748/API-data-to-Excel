# COVID-19 Data Analysis and Visualization (South America) with OpenPyXL

In this automation project, I built a Python script that fetches weather data from an API and exports it into an Excel file. I used OpenPyXL to generate line, bar, scatter, and pie charts directly within the spreadsheet. This workflow allowed me to automate data reporting, ensuring consistency and saving time. The project shows my capacity to combine Python programming with Excel automation tools.

The Excel file contains:
- Raw COVID-19 data for each South American country
- Four different chart types (Bar, Line, Pie, and Scatter) saved into separate sheets
- Automatically generated images embedded in the Excel workbook

---

## Features

- **Real-time API Data Retrieval**  
  Pulls current COVID-19 statistics from `https://disease.sh/v3/covid-19/countries`.
  
- **Data Filtering**  
  Automatically limits results to countries in **South America**.

- **Excel Workbook Creation**  
  Uses **OpenPyXL** to create a workbook with:
  - **Raw Data Sheet** – Includes country, continent, population, cases, deaths, recoveries, active cases, testing data, and per-million metrics.
  - **Bar Chart Sheet** – Total COVID-19 cases by country.
  - **Line Chart Sheet** – COVID-19 cases per million by country.
  - **Pie Chart Sheet** – Share of total COVID-19 deaths by country.
  - **Scatter Chart Sheet** – Relationship between total tests and total cases, with country labels.

- **Matplotlib Visualization**  
  Generates all charts dynamically, saves them as PNG images, and embeds them into the Excel workbook.

- **Completely Automated**  
  Running the script downloads the latest data, builds all charts, and saves the Excel file ready for analysis.

---

## Tech Stack

- **Python** – Core programming language
- **Requests** – API calls to fetch COVID-19 data
- **OpenPyXL** – Excel workbook creation and chart embedding
- **Matplotlib** – Chart generation and saving as images
- **disease.sh API** – Free COVID-19 data source
