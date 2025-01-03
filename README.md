# Bulk Document Generator (VBA)

## Description
A VBA-powered bulk document generator for automating the creation of personalized documents such as certificates, contracts, or letters. This tool merges data from an Excel sheet into a predefined Word template, significantly reducing the time and effort involved in generating repetitive documents. Ideal for corporate environments where document creation is a frequent task.

## Features
- Integrates with Microsoft Word and Excel
- Automatically fills placeholders in Word template with data from Excel
- Saves each generated document with a unique name
- Customizable for various document templates

## Prerequisites
Before using this project, ensure that you have the following:
- **Microsoft Excel and Word** installed on your machine.
- **Basic knowledge of VBA** for modification and customizations.
- **Access to Excel Data** with necessary information (like names, project names, etc.) in structured columns.

## Setup & Instructions

### 1. Download and Open the Project Files
- Clone the repository or download the ZIP file containing the project files.
- Open the `BulkDocumentGenerator.xlsm` file in Microsoft Excel. This file contains the VBA code that automates the process.

### 2. Modify the Excel Sheet
The Excel sheet should have the following column headers:
- **Name**: Name of the person the document will be generated for.
- **Project Name**: The name of the project associated with the document.
- **Date**: Date of the document creation or award.

Example Data:
| Name          | Date       | Contract ID | Project Name   |
|---------------|------------|-------------|----------------|
| John Doe      | 05/01/2025 | 1234        | Marketing Campaign |
| Jane Smith    | 10/02/2025 | 5678        | Website Development |

Ensure your data is filled in correctly under these columns.

### 3. Open and Configure the Word Template
The VBA code will reference a **Word document template** (e.g., `Template.docx`) where placeholders like `<Name>`, `<Project Name>`, and `<Date>` will be replaced by data from the Excel sheet.

Ensure that your template is set up correctly with placeholders in the format:
