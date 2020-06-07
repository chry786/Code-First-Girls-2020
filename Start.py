import csv
import pandas as pd
import xlsxwriter

df = pd.read_csv("sales.csv")


# Reading in the data
def read_data():
    data = []

    with open('sales.csv', 'r') as sales_csv:
        spreadsheet = csv.DictReader(sales_csv)
        for row in spreadsheet:
            data.append(row)
    return data


# Total amount of sales
def total():
    data = read_data()

    sales = []
    for row in data:
        sales.append(int(row["sales"]))

    total_sales = sum(sales)
    print('Total sales: {}'.format(total_sales))


total()
print("\n")


# Calculates the least amount of sales
def least():
    data = read_data()
    month = []
    sold = []
    for row in data:
        sold.append(int(row["sales"]))
    for months in data:
        month.append(months["month"])
    minimum = min(sold)
    print("The month with the least amount of sales are: {}".format(month[sold.index(min(sold))]))
    print('The minimum amount of sales are: {}'.format(minimum))


least()
print("\n")


# Calculates the most amount of sales
def most():
    data = read_data()

    list = []
    month = []
    for row in data:
        list.append(int(row["sales"]))
    for months in data:
        month.append(months["month"])
    maximum = max(list)
    print("The month with the highest amount of sales are: {}".format(month[list.index(max(list))]))
    print('The maximum amount sold are: {}'.format(maximum))


most()
print("\n")


# Function to calculate the average amount of sales/expenditure
def average(lst):
    return sum(lst) / len(lst)


# Calculates the average amount of sales using the average function
def avg():
    data = read_data()
    sales = []

    for row in data:
        sales.append(int(row["sales"]))

    print("The average sales are: " + str(average(sales)))


avg()
print("\n")


# Calculates the average amount of expenditure using the average function
def exp():
    data = read_data()
    expenditure = []

    for row in data:
        expenditure.append(int(row["expenditure"]))

    print("The average expenditure is: " + str(average(expenditure)))


exp()
print("\n")

# Uses the Pandas packages describe() method to view basic statistical details of the sales & expenditure columns
print("Statistical details using Pandas Package")
sales_expenditure = df[['sales', 'expenditure']].describe()
print(sales_expenditure)

# Writing to an excel spreadsheet using the Xlwt module
df = pd.DataFrame(sales_expenditure)
writer = pd.ExcelWriter('Sales_Expenditure_Analysis.xlsx', engine='xlsxwriter')
df.to_excel(writer, sheet_name='Sheet1')
writer.save()


