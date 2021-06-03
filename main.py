import pandas as pd
import openpyxl
def Manipulate(path):
    try:
        identifierurl = 'Identifiers.xlsx'
        df = pd.ExcelFile(identifierurl).parse('Sheet1')
        cl = {b for b in df['Identifiers']}
        excelsheet = path
        save = path
        name = ""
        if ('POST' in excelsheet):
            name = excelsheet[excelsheet.index('POST\\') + 5:]
        if ('PRE' in excelsheet):
            name = excelsheet[excelsheet.index('PRE\\') + 4:]
        print(name)
        if name != '':
            print('if')
            datasheet = openpyxl.load_workbook(excelsheet)
            sn = datasheet.sheetnames
            for nm in sn:
                if (nm != 'Cover' and nm != 'HOME'):
                    one = datasheet[nm]
                    count = 0
                    print(cl)
                    print(len(one[1]))
                    for a in range(len(one[1])):
                        print(count)
                        print(one[1][count].value)
                        if (one[1][count].value in cl):
                            print('in if')
                            one[1][count].value = "*" + one[1][count].value
                        count = count + 1
                    print(one[1][1].value)
            datasheet.save(save)
        else:
            print('Invalid Path, Files Only Present In PRE and POST Folders Will Be Read')
    except Exception as Error:
        if str( Error) == "substring not found":
            print('Invalid Path, Files Only Present In PRE and POST Folders Will Be Read')
        if '[Errno 2] No such file or directory' in str(Error):
            print('File with name Identifiers should be in pre or post folders')
        else:
            print(str(Error))
        #print(NameError)

Manipulate('POST\\2G Cell Frequency Data Template.xlsx')
Manipulate('PRE\\2G Cell Frequency Data Template.xlsx')

