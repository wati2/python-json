import sys
import pprint
import xlsxwriter
import json

class FileMangement:
    def __init__(self):
        self.json_file_name = "input.json"
        self.excel_file_name = "output.xlsx"
        self.array_x = []
        self.array_y = []
        self.array_sum = []

    def read_text_file(self):
        try:
            with open(self.json_file_name) as data_file:
                data = json.load(data_file)
                pprint.pprint(data)
                for each_axis in data["data"]:
                    x = int(each_axis["x"])
                    y = int(each_axis["y"])
                    self.array_x.append(x)
                    self.array_y.append(y)
                    self.array_sum.append(x + y)
                    pprint.pprint("x = {0}".format(x))
                    pprint.pprint("y = {0}".format(y))

        except:
            print("Inexpected error: ", sys.exc_info()[0])
            raise

    def save_to_xlsx(self):
        workbook = xlsxwriter.Workbook(self.excel_file_name)
        worksheet = workbook.add_worksheet()
        for index, value in enumerate(self.array_x):
            print(index, value)
            worksheet.write(index, 0, self.array_x[index]) # column 0
            worksheet.write(index, 1, self.array_y[index]) # column 1
            worksheet.write(index, 2, self.array_sum[index]) # column 1
        workbook.close()

if __name__ == "__main__":
    file_management = FileMangement()
    file_management.read_text_file()
    file_management.save_to_xlsx()