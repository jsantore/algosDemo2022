import openpyxl
import numbers


def get_sheet(filename):
    excel_file = openpyxl.load_workbook(filename)
    return excel_file.active


def count_rows(datasheet):
    count = 0
    for row in datasheet.rows:
        #only count the rows with data
        if not isinstance(row[2].value, numbers.Number):
            continue
        count +=1
    return count

def find_largest(datasheet):
   # datasheet = get_sheet("MedianIncomeByStateCensusGov.xlsx")
    largest_row = None
    for row in datasheet.rows:
        if not isinstance(row[2].value, numbers.Number):
            continue
        if largest_row is None:
            largest_row = row
        if largest_row[1].value   <  row[1].value:
            largest_row = row
    return largest_row

def find_smallest(datasheet):
    smallest_row = None
    for row in datasheet.rows:
        if not isinstance(row[2].value, numbers.Number):
            continue
        if smallest_row is None:
            smallest_row = row
        if smallest_row[1].value > row[1].value:
            smallest_row = row
    return smallest_row

def find_average_income(datasheet):
    count =0
    sum = 0
    for row in datasheet.rows:
        if not isinstance(row[2].value, numbers.Number):
            continue
        sum += row[1].value
        count +=1
    average = sum/count
    return average

def main():
    datasheet = get_sheet("MedianIncomeByStateCensusGov.xlsx")
    num_rows = count_rows(datasheet)
    print(num_rows, " rows")
    biggest = find_largest(datasheet)
    print(f"the largest was {biggest[0].value} with income of {biggest[1].value}")
    smallest = find_smallest(datasheet)
    print(f"the smallest was {smallest[0].value} with a median income: {smallest[1].value}")
    ave_income = find_average_income(datasheet)
    print(f"average state income was {ave_income}")

# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    main()

# See PyCharm help at https://www.jetbrains.com/help/pycharm/
