# Advertising Data Tracker & Analysis

A Python-based application for organizing, processing, and analyzing advertising data through a structured data preparation and reporting workflow.

**Live Demo:** [Open the Advertising Data Tracker](https://ad-analysisv2.streamlit.app)

## Overview

The Advertising Data Tracker transforms raw advertising information into structured, analysis-ready data. It combines data processing, cleaning, and category mapping to improve data consistency and prepare records for analysis and reporting.

The project demonstrates how Python and spreadsheet-based workflows can help organize advertising information and make it easier to analyze.

## Problem Statement

Advertising data can contain inconsistent formats, categories, and record structures, making it difficult to compare information, identify patterns, and generate reliable reports. This project addresses that challenge by providing a structured workflow for preparing raw advertising information for analysis.

## Key Features

* **Data Processing:** Organizes raw advertising information into structured datasets.
* **Data Cleaning:** Improves data consistency and prepares records for analysis.
* **Category Mapping:** Supports standardized classification using category mappings.
* **Data Organization:** Separates raw inputs, processed data, and outputs.
* **Analytics and Reporting:** Prepares structured information for summaries and further analysis.
* **Interactive Interface:** Provides a web-based interface built with Streamlit.

## Workflow

The application follows a data preparation pipeline:

1. **Input:** Load advertising data from source files.
2. **Transformation:** Restructure raw information into a consistent tabular format.
3. **Cleaning:** Address inconsistencies and prepare records for analysis.
4. **Classification:** Apply category mappings to organize advertising records.
5. **Output:** Prepare processed data for reporting and analysis.

## Technology Stack

| Technology | Purpose                               |
| ---------- | ------------------------------------- |
| Python     | Application logic and data processing |
| Pandas     | Data manipulation and transformation  |
| Streamlit  | Interactive application interface     |
| Excel      | Data inputs and category mapping      |

## Project Structure

* `app.py` — Main application.
* `raw_inputs/` — Source data files.
* `warehouse/` — Organized data used in the processing workflow.
* `output/` — Processed outputs.
* `category_map.xlsx` — Category mapping reference.
* `requirements.txt` — Python dependencies.
* `.streamlit/config.toml` — Streamlit configuration.

## Getting Started

### Prerequisites

* Python
* pip
* Git

### Installation

1. Clone the repository:

   ```bash
   git clone https://github.com/andrinawatsemba/AD-analysisV1.3.git
   cd AD-analysisV1.3
   ```

2. Create a virtual environment:

   ```bash
   python -m venv .venv
   ```

3. Activate the environment.

   **Windows:**

   ```bash
   .venv\Scripts\activate
   ```

   **macOS/Linux:**

   ```bash
   source .venv/bin/activate
   ```

4. Install the dependencies:

   ```bash
   pip install -r requirements.txt
   ```

5. Launch the application:

   ```bash
   streamlit run app.py
   ```

   Open the local URL provided by Streamlit to access the application.

**Note:** Ensure that the required input files and supporting resources are available in the locations expected by the application.

## Skills Demonstrated

* Data cleaning and preprocessing
* Data wrangling and transformation
* Structured data management
* Category mapping and standardization
* Python and Pandas
* Interactive application development with Streamlit
* Data workflow organization

## Author

**Andrina Watsemba**

Data Science and Analytics | Data Analytics | Analytics Automation

---

*An applied data analytics project focused on transforming raw advertising information into structured, analysis-ready datasets.*
