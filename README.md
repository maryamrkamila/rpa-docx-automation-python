# rpa-docx-automation-python
Python-based automation for generating personalized Word and PDF documents from CSV data using dynamic table expansion with python-docx. Built with pandas and python-docx for high-volume document processing.

# 📄 Automated Document Generation Pipeline for Census Field Operations (RPA)

Python-based Robotic Process Automation (RPA) script to dynamically generate customized Microsoft Word (.docx) and PDF assignment letters for Field Supervisors using `pandas` and `python-docx`.

**Author:** Maryam Rahma Kusuma Kamila | [www.linkedin.com/in/maryamrkamila](https://www.linkedin.com/in/maryamrkamila)  
**Context:** Internship Project  

## 📌 Project Overview

This project automates the generation of Task Completion Documents for Field Supervisors (PML) during census operations. Instead of manually preparing hundreds of documents, this script generates personalized MS Word and PDF files directly from a structured CSV master dataset. The solution supports dynamic attachment tables, allowing each supervisor's document to adaptively include a varying number of supervised Field Enumerators (PPL) in a single execution.

## 🎯 The Problem & Business Objective

The manual administrative process faced a significant bottleneck:
* **Massive Scale:** Handling data for over 300 supervisors (PML) and 2,500+ field enumerators (PPL).
* **One-to-Many Relationship:** Each PML required an individual document featuring an attachment table with a varying number of rows (representing their specific PPLs).
* **Tool Limitations:** Traditional MS Word Mail Merge could not dynamically expand table rows, rendering it unusable for this workflow.
* **Operational Risk:** Manual repetitive copy-paste work posed a high risk of data mismatch and formatting inconsistencies.

**Objective:** To build a fully automated, zero-error pipeline that eliminates manual data entry and standardizes the output generation process.

## 🚀 Key Features

* **Dynamic Table Expansion:** Automatically generates and populates attachment table rows based on the exact number of PPLs assigned to a supervisor.
* **Native XML Preprocessing:** Built-in XML parsing logic to fix split placeholders caused by MS Word's background formatting, ensuring data is injected without destroying native formatting (e.g., preserving Tab spacings and center alignments).
* **Automated File Naming & Numbering:** Generates sequential document numbers (e.g., `B-0001`) and standardized file names based on regional codes.
* **Batch Conversion:** Seamlessly outputs both `.docx` and `.pdf` files into a centralized output directory in a single script execution.

## 🛠️ Tech Stack & Tools

* **Language:** Python
* **Data Processing:** `pandas`
* **Document Manipulation:** `python-docx`, `lxml` (XML Engine)
* **Format Conversion:** `docx2pdf`

## 📝 Technical Approach & Methodology

1. **ETL & Data Grouping:** Read the master CSV dataset using `pandas`. Utilized `groupby()` to aggregate the 2,500 PPL records accurately under their respective 300 PMLs. Data types (dtype) were strictly mapped to preserve leading zeros in regional codes.
2. **Template Engineering:** Designed a standardized MS Word template using static boilerplate text and dynamic placeholders (e.g., `{{ nama_pml }}`).
3. **Advanced Word XML Manipulation:** Rather than relying on standard templating libraries that often crash due to MS Word's hidden spell-check tags splitting the text runs, this script utilizes `lxml` to safely read, repair, and replace text at the root XML level.
4. **Row Cloning & Alignment:** Cloned the base template row, populated it with PPL data, and programmatically applied `WD_ALIGN_VERTICAL.CENTER` to specific cells before inserting them back into the document tree.

## 🏆 Achievements & Impact

* **Extreme Time Efficiency:** Transformed an administrative task that would have taken weeks of tedious manual labor into an automated script that runs in just a few minutes.
* **Zero Data-Entry Error:** Ensured 100% accuracy in matching PPL workloads to the correct PML supervisors by removing human intervention.
* **Scalable Framework:** Created a robust and reusable Python framework that can be seamlessly adapted for future census operations or similar reporting workflows.

## 📂 Repository Structure

```text
├── data/
│   └── master_data.csv            # Master dataset (Example/Dummy data)
├── template/
│   └── template_pml.docx          # MS Word base template with placeholders
├── output/                        # Auto-generated directory for output files
├── auto_pml_final.py              # Main Python RPA script
└── README.md                      # Project documentation
