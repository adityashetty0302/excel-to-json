import json as json
import os
import time
import xlrd

if __name__ == '__main__':

    os.chdir(r"C:\Users\Aditya.Shetty\Desktop\Aditya\samples")

    # By for logic
    t1 = time.time()
    wb = xlrd.open_workbook('t1.xlsx')
    workbook = []
    for x in range(0, wb.nsheets):

        sheet = wb.sheet_by_index(x)
        row_list = []
        for rownum in range(1, sheet.nrows):
            data = {}
            row_values = sheet.row_values(rownum)
            data["no"] = row_values[0]
            data["name"] = row_values[1]
            data["id"] = row_values[2]
            data["age"] = row_values[3]

            row_list.append(data)

        sheetdata = {}
        sheetdata[sheet.name] = row_list
        workbook.append(sheetdata)

    z = json.dumps({"t1": workbook})
    with open("data.json", "w") as f:
        f.write(z)
        f.close()

    print(z)
    t2 = time.time()
    print(t1, "\t", t2)
    print("\ntime by for logic: ", t2 - t1)

    print("\n")

    '''
    #By built in function
    t3=time.time()
    
    wb= xlrd.open_workbook('t1.xlsx')
    #print(wb.nsheets)
    
    #print(wb.sheets())
    
    for x in wb.sheets():
       
        data_list = []
    
        for rownum in range(1, x.nrows):
            data = OrderedDict()
            row_values = x.row_values(rownum)
            data["no"]= row_values[0]
            data["name"]= row_values[1]
            data["id"]= row_values[2]
            data["age"]= row_values[3]
            
            data_list.append(data)
            
            j=json.dumps(data_list)
            
        with open("data.json","a") as f:
            f.write(j)
            f.close()
            print(j)
            
    t4=time.time()
    print(t3,"\t",t4)
    print("\ntime by Built in function:",t4-t3)
    '''
