# Minitab Examples from LinkedIn Learning

This repository contains examples from a LinkedIn Learning Minitab course taught by Professor Richard Chua. The examples featured here are utilized in his Six Sigma Green and Black Belt courses.
Seletcted examples were commented and exported in pdf format.

## Course Link
For more information and to access the course, visit the following link:  
[Six Sigma Green Belt Course on LinkedIn Learning](https://www.linkedin.com/learning/six-sigma-green-belt/)
[Quality Analytics Using MiniTab]
(https://www.linkedin.com/learning/quality-analytics-using-minitab)

_______________________________________________________________________

# Gage R&R Analysis Tool (Python)

This tool automates the creation of a Gage Repeatability and Reproducibility (Gage R&R) input form in an Excel file. It's designed to assist quality control professionals in performing measurement system analysis, particularly where Minitab is not available. The script prompts for user inputs to customize the study parameters, including the number of parts, operators, and replicates, and it randomizes the order of parts for each operator to minimize bias.

## Features

- **Excel File Creation**: Automatically creates a new Excel file or overwrites an existing one to store Gage R&R data.
- **Customizable Parameters**: Allows users to specify the number of parts, operators, and replicates.
- **Randomization of Runs**: Randomizes the order of parts for each operator and replicate to prevent order bias.
- **Easy Data Entry**: Generates a ready-to-use Excel file with placeholders for manual data entry.

## Prerequisites

Before you can run this script, you'll need the following installed on your system:
- Python (3.6 or later recommended)
- `openpyxl` library

You can install `openpyxl` using pip if it's not already installed:

```bash
pip install openpyxl