# Power Point Automation with Python
This project automates the process of updating a PowerPoint presentation that is delivered to clients on a monthly basis.
The data pipeline for this report normally queries a database using SQL; however, this repository uses a randomly generated sample data to demonstrate the automation process in a standalone example. 
This script automatically updates tables, charts, and text boxes in a PowerPoint template, reducing manual effort and ensuring consistent formatting. 

# Features
- Automates monthly reporting - eliminates manual editing efforts
- Populates tables, charts, and text boxes dynamically
- Uses random sample data to demontrate the process
- Supports custom formatting
- Easily adaptable for real business data sources 

# Tech Stack
- Language: Python
- Libraries:
    - pandas, numpy, datetime: data generation & manipulation
    - python-pptx: PowerPoint automation (tables, charts, text formatting)
    - shutil, os: file handling and path management
    - random: sample data generation for demonstration

# Project Structure
- Update_PowerPoint.ipynb #Main script to generate and update PowerPoint
- Report_Template.pptx #PowerPoint template to update
- Report_Template_Updated.pptx #PowerPoint updated version after running the script
- README.md # Project documentation

# How to Run the Project
1. Download the following files: 
    - Update_PowerPoint.ipynb (Jupyter Notebook)
    - Report_Template.pptx (PowerPoint template)
2. Open only the Jupyter Notebook file  
    - Make sure the PowerPoint file is closed while running the script to avoid file access errors 
3. Install the required libraries: pandas, numpy, python-pptx
4. Update the PowerPoint file path:
    - Update the first cell to set the path for reading the PowerPoint template
    - Update the last cell to set the path for saving the updated PowerPoint with all changes
5. Run all cells in order 
