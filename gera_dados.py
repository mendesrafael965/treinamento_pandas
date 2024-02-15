import csv
import xlsxwriter
from faker import Faker
from datetime import datetime, timedelta

fake = Faker('pt_BR')

# Função para gerar data de nascimento entre 18 e 65 anos atrás


def generate_birthdate():
    return fake.date_of_birth(minimum_age=18, maximum_age=65).strftime("%d/%m/%Y")

# Função para gerar data de admissão entre 1 e 10 anos atrás


def generate_hire_date():
    end_date = datetime.now()
    start_date = end_date - timedelta(days=365*10)
    return fake.date_time_between_dates(start_date, end_date).strftime("%d/%m/%Y")


# Gerar 100 registros de funcionários
employees = []
for i in range(1000):
    employee = {
        'NumCad': i + 1,
        'NomFun': fake.name(),
        'DatNas': generate_birthdate(),
        'DatAdm': generate_hire_date()
    }
    employees.append(employee)

# Inserir dados em csv

keys = employees[0].keys()

with open('employee.csv', 'w', newline='') as output_file:
    dict_writer = csv.DictWriter(output_file, keys)
    dict_writer.writeheader()
    dict_writer.writerows(employees)


# Inserir dados em excel
columns = list(employees[0].keys())
rows = [list(result.values()) for result in employees]

workbook = xlsxwriter.Workbook('employees.xlsx')
worksheet = workbook.add_worksheet('employees.xlsx')

date_format = workbook.add_format({'num_format': 'dd/mm/yy'})

for row in range(len(rows)):
    for column in range(len(columns)):
        if row == 0:
            worksheet.write(row, column, columns[column])
        else:
            if column > 1:
                worksheet.write(row, column, rows[row-1][column], date_format)
            else:
                worksheet.write(row, column, rows[row-1][column])
workbook.close()
