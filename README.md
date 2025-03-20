# soce_scraper: Extract data from th SOCE (Sistema Oficial de Contratación Pública del Ecuador)
![Python](https://img.shields.io/badge/python-3.9-%233776AB.svg?style=for-the-badge&logo=python&logoColor=white)
![Flask](https://img.shields.io/badge/Flask-3.1.0-%23000.svg?style=for-the-badge&logo=flask&logoColor=white)
![Pandas](https://img.shields.io/badge/Pandas-2.2.3-%23150458.svg?style=for-the-badge&logo=pandas&logoColor=white)
![Requests](https://img.shields.io/badge/Requests-2.32.3-%23FF6F61.svg?style=for-the-badge&logo=requests&logoColor=white)

## Table of Contents
1. [Abstract](#abstract)
2. [Theoretical Framework](#theoretical-framework)
3. [Methodology](#methodology)
4. [Deployment](#deployment)
5. [Technical Implementation](#technical-implementation)
6. [Installation](#installation)
7. [Usage](#usage)
8. [License](#license)

## Abstract
soce_scraper is a web scraping tool designed to extract data from the Sistema Oficial de Contratación Pública del Ecuador (SOCE). It allows users to gather information on various procurement processes, including "Ínfima Cuantía", "Procesos", "Regimen Especial", and "Procedimiento especial". The extracted data is then exported to Excel files for further analysis.

## Theoretical Framework
- **Web Scraping**: The process of extracting data from websites using automated scripts.
- **Data Processing**: Utilizing libraries like Pandas to manipulate and analyze the scraped data.
- **Excel Exporting**: Using openpyxl to create Excel files with hyperlinks for easy access to detailed information.

## Methodology
The soce_scraper follows a structured approach:
1. **User  Input**: Users provide a date range and select the type of procurement process they wish to extract data for.
2. **Data Extraction**: The scraper sends requests to the SOCE website and retrieves the relevant data.
3. **Data Processing**: The scraped data is processed and formatted into a structured format.
4. **Excel Export**: The processed data is exported to Excel files, with appropriate formatting and hyperlinks.

## Deployment
soce_scraper is built using Flask, allowing it to run as a web application. The application processes user requests and returns the extracted data in Excel format.

## Technical Implementation
soce_scraper is developed in Python 3.9, utilizing:
- **Flask** for the web interface.
- **Requests** for making HTTP requests to the SOCE website.
- **Pandas** for data manipulation and analysis.
- **Openpyxl** for exporting data to Excel files.

## Installation
### Prerequisites
- **Python 3.9**: Ensure that Python is installed and configured in your system.

### Installation
- **Clone the repository**
  ```bash
  git clone https://github.com/nava2105/soce_scraper.git
  cd soce_scraper
  ```
- **Required Libraries**: Install the necessary libraries using pip:
  ```bash
  pip install -r requirements.txt
  ```
- **Run the application**
  ```bash
  python app.py
  ```
- **Access the application:** Open your web browser and navigate to http://localhost:5000

## Usage
Once the application is up and running, you can interact with the system by:
- Selecting the start and end dates for the data extraction.
- Choosing the type of procurement process.
- Submitting the form to generate and download the corresponding Excel file.

### License
This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.