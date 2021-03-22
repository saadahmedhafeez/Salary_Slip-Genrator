from pandas import ExcelFile
from datetime import datetime
import pdfkit
from bs4 import BeautifulSoup
from num2words import num2words
import os
import locale

current_month = datetime.now().strftime('%h')
current_year = datetime.now().year

xls = ExcelFile('salary_sheet.xlsx')
df = xls.parse(xls.sheet_names[0])
header = df.iloc[1]
df = df[2:]
df.columns = header
list_of_employees = df.to_dict('records')

options = {'enable-local-file-access': None}

locale.setlocale(locale.LC_ALL, '')


def generate_slip(list_of_dicts):
    file_count = 1
    temp_html = 'temp{}.html'
    filename = 'index.html'
    opened_file = open(filename, 'r')
    html_file = opened_file.read()
    opened_file.close()
    soup = BeautifulSoup(html_file, 'lxml')
    for emp_dict in list_of_dicts:
        try:
            date_time = soup.find(id='DateTime')
            date_time.string = ('Pay slip for the month of {}, {}'.format(current_month, current_year))
            employee_no = soup.find(id='employee_no')
            employee_no.string = (' {}'.format(emp_dict['Employee No']))
            employee_name = soup.find(id='employee_name')
            employee_name.string = (' {}'.format(emp_dict['Full Name']))
            designation = soup.find(id='designation')
            designation.string = (' {}'.format(emp_dict['Designation']))
            cnic = soup.find(id='cnic')
            cnic.string = (' {}'.format(emp_dict['CNIC #']))
            ac_no = soup.find(id='ac_no')
            ac_no.string = (' {}'.format(emp_dict['Account #']))
            basic_salary = soup.find(id='basic_salary')
            basic_salary_amount = round(emp_dict['Basic Salary'])
            basic_salary.string = ('{}'.format(f'{basic_salary_amount:n}'))
            income_tax = soup.find(id='income_tax')
            income_tax_amount = emp_dict['Income Tax/ Month'] * -1
            income_tax.string = ('{}'.format(f'{income_tax_amount:n}'))
            medical_allowance = soup.find(id='medical_allowance')
            medical_allowance_amount = round(emp_dict['Medical'])
            medical_allowance.string = ('{}'.format(f'{medical_allowance_amount:n}'))
            loan_returned = soup.find(id='loan_returned')
            loan_returned_amount = emp_dict['Loan/ Advance Returned']
            loan_returned.string = ('{}'.format(f'{loan_returned_amount:n}'))
            advance_loan = soup.find(id='advance_loan')
            advance_loan_amount = emp_dict['Advance/ Loan']
            advance_loan.string = ('{}'.format(f'{advance_loan_amount:n}'))
            lunch_expense = soup.find(id='lunch_expense')
            lunch_expense_amount = emp_dict['Lunch Expense'] * -1
            lunch_expense.string = ('{}'.format(f'{lunch_expense_amount:n}'))
            lunch_expense_co = soup.find(id='lunch_expense_co')
            lunch_expense_co_amount = emp_dict['Lunch .co']
            lunch_expense_co.string = ('{}'.format(f'{lunch_expense_co_amount:n}'))
            arrears = soup.find(id='arrears')
            arrears_amount = emp_dict['Arrears']
            arrears.string = ('{}'.format(f'{arrears_amount:n}'))
            wfh_deduction = soup.find(id='wfh_deduction')
            wfh_deduction_amount = emp_dict['WFH Deduction']
            wfh_deduction.string = ('{}'.format(f'{wfh_deduction_amount:n}'))
            internet = soup.find(id='internet')
            internet_amount = emp_dict['Internet']
            internet.string = ('{}'.format(f'{internet_amount:n}'))
            unpaid_leaves = soup.find(id='unpaid_leaves')
            unpaid_leaves.string = ('({})'.format(emp_dict['Unpaid Leaves']))
            unpaid_leaves_amount = soup.find(id='unpaid_leaves_amount')
            unpaid_leaves_amount2 = emp_dict['Unpaid Leaves Deduction']
            unpaid_leaves_amount.string = ('{}'.format(f'{unpaid_leaves_amount2:n}'))
            food_reimbursement = soup.find(id='food_reimbursement')
            food_reimbursement.string = ('{}'.format(''))
            travel_reimbursement = soup.find(id='travel_reimbursement')
            travel_reimbursement.string = ('{}'.format(''))
            eobi_co_share = soup.find(id='eobi_co_share')
            eobi_co_share.string = ('{}'.format(''))

            total_pay_sum = round(
                round(emp_dict['Basic Salary']) + round(emp_dict['Medical']) + emp_dict['Advance/ Loan']
                + emp_dict['Lunch .co'] + emp_dict['Arrears'] + emp_dict['Internet'])
            total_pay_deduct = ((emp_dict['Income Tax/ Month'] * -1) + emp_dict['Loan/ Advance Returned'] +
                                (emp_dict['Lunch Expense'] * -1) + emp_dict['WFH Deduction'] +
                                emp_dict['Unpaid Leaves Deduction'])
            total_receivable = soup.find(id='total_receivable')
            total_receivable.string = ('{}'.format('{}'.format(f'{total_pay_sum:n}')))
            total_deduction = soup.find(id='total_deduction')
            total_deduction.string = ('{}'.format('{}'.format(f'{total_pay_deduct:n}')))
            net_salary = soup.find(id='net_salary')
            net_salary_amount = round(emp_dict['Payable Salary'])
            net_salary.string = ('{}'.format(f'{net_salary_amount:n}'))
            amount_in_words = soup.find(id='amount_in_words')
            num_to_words = num2words(round(emp_dict['Payable Salary'])).capitalize()
            amount_in_words.string = ('{} rupees only/-'.format(num_to_words))

            opened_file = open(temp_html.format(file_count), 'w')
            opened_file.write(str(soup))
            opened_file.close()
            pdfkit.from_file(temp_html.format(file_count),
                             'generated_slips/{}-{} {}.pdf'.format(emp_dict['Full Name'], current_month, current_year),
                             options=options)
            os.remove(temp_html.format(file_count))
            file_count += 1
        except:
            print("Some value is missing!!!")


generate_slip(list_of_employees)
