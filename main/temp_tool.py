import os,sys
import pprint
import xlrd

class temp_toolBox():

    def __init__(self, excelName, sheetNum):
        os.chdir(os.path.abspath(os.path.dirname(__file__)))
        self.excel_dir = os.path.split(os.getcwd())[0] + os.sep + "excels"
        self.excelName = excelName
        self.sheetNum = sheetNum
        self.LoD = []
        self.json_dict = {}

    def get_data_from_excel(self):
        execlPath = self.excel_dir + os.sep + self.excelName + '.xls'
        data = xlrd.open_workbook(execlPath)
        table = data.sheets()[self.sheetNum]
        nor = table.nrows
        nol = table.ncols
        dict = {}
        for i in range(2, nor):
            for j in range(nol):
                title = table.cell_value(0, j)
                value = table.cell_value(i, j)
                dict[title] = value
            yield dict



    def convert_LoD_to_json(self, root_name):
        if not self.LoD:
            print(f"Empty List")
        else:
            self.json_dict[root_name] = self.LoD

    def __str__(self):
        return  f"{__class__}'s self.excel_dir => {self.excel_dir}"


if __name__ == '__main__':


    print(f"os.path.abspath(os.path.dirname(__file__)) => {os.path.abspath(os.path.dirname(__file__))}")
    task1 = temp_toolBox("MIB3_DIAG_Key_Value_Pairs_CNS3.0",3)
    for i in task1.get_data_from_excel():
        task1.LoD.append(dict(i))
    task1.convert_LoD_to_json("ns_key_table")
    pprint.pprint(task1.json_dict)
