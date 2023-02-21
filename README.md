# Lab Report PDF Generator

This script generates a PDF report based on data from an Excel file. It reads an Excel file and extracts data from the cells to display in a report. It then calculates a total score and risk factor based on the data and appends these values to the Excel file.

## Prerequisites

Before running this script, make sure that the following modules are installed:

- openpyxl
- reportlab

You can install these modules using pip by running the following command:

- pip install openpyxl
- pip install reportlab

## Usage

1. Save your Excel file in the same directory as this script. The Excel file should contain a sheet with the data you want to include in the report.
2. Open a command prompt or terminal and navigate to the directory where this script and the Excel file are saved.
3. Run the following command to generate the PDF report:
4. The PDF report will be saved in the same directory as the script.

## Notes

- This script assumes that the Excel file has a sheet with data starting from the second row. The first row should contain column headers.
- The script will add two columns to the end of the Excel file: "Total Score" and "Risk Factor".
- The "Total Score" is calculated as the sum of values in columns C, D, E, F, and H multiplied by their respective factors.
- The "Risk Factor" is calculated based on the "Total Score" as follows:

- If the total score is less than 30, the risk factor is "Low".
- If the total score is between 30 and 39, the risk factor is "Medium".
- If the total score is 40 or greater, the risk factor is "High".

## Contributing
### Github   
- [Sezgin](https://github.com/Sezgin3880)               
- [Azlaan](https://github.com/AzlaanIrshad)


### Linkedin
- [Sezgin](https://www.linkedin.com/in/sezgin-karaduman-619483221/)               
- [Azlaan](https://www.linkedin.com/in/azlaan-irshad/)
