# Salary Slip IKS Logics
 &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[![Image](https://ikslogics.com/wp-content/uploads/2020/08/white-and-green-logo.png)](https://ikslogics.com)
## Features!
* This project generates salary slip of IKS Logics employees in PDF using Python, HTML and CSS.
* Month and year will be picked automatically by system
`Important` if you don't generate the slips before month change you will not be able to generate it for that month.

## Prerequisite
* You must import excel file with employee details of *current month* to generate PDF salary slips for all employees for that *specific month*
* name of file should be **salary_sheet.xlsx**
* there must be a **logo.png** of IKS Logics

### Install the dependencies
- Install Python version 3.8.3 or higher can be download from [HERE](https://www.python.org/downloads/)
- set it's path automatically while installation.
- install requirements by running this command `pip install -r requirements.txt`
- download and install [wkhtmltopdf tool](https://wkhtmltopdf.org/downloads.html) and set the env path for *wkhtmltopdf*
- restart your ide tool.

### Development
- open project location in terminal
- run command `python main.py`
