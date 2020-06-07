import gspread
import xlsxwriter
from oauth2client.service_account import ServiceAccountCredentials
from pprint import pprint
from xlrd import open_workbook
import time

print ('Execution Started...')
scope = ["https://spreadsheets.google.com/feeds",'https://www.googleapis.com/auth/spreadsheets',"https://www.googleapis.com/auth/drive.file","https://www.googleapis.com/auth/drive"]

creds = ServiceAccountCredentials.from_json_keyfile_name("creds.json", scope)

client = gspread.authorize(creds)

# Open the google spreadhseet
sheet = client.open("12LPPLUS_V1.0_7P5T").sheet1  

# Get a list of all records 
data = sheet.get_all_values()
#print ('Processed Row ', len(data))

# Excel sheet name
generated_excel_sheet_name = '12LPPLUS_V1.0_7P5T.xlsx' 
  
# Workbook() takes one, non-optional, argument which is the excel sheet name that we want to create. 
workbook = xlsxwriter.Workbook(generated_excel_sheet_name) 
  
# The workbook object is then used to add new worksheet via the add_worksheet() method. 
worksheet = workbook.add_worksheet()

start_index = 0
end_index = 5
row = 0
column = 0

# declared list for heading 
header = ['Process', 'Track', 'Threshold', 'Lg', 'PVT Corner', 'Temp', 'PDK Version', 'Vnom']
# heading values
empty_header_value = ['12LP+','7.5T','','','','','1.0','0.8'] 

# assume highest length of row is 80, each 5 contains a block until the next empty list
for start_index in range(0,80,5): 
    #print ('start_index ', start_index)
    #print ('end_index ', end_index)
    for j in data:
        #print (j[start_index : end_index])
        item = j[start_index : end_index] # each list contains 5 elements
        column = 0
        for content in item: # each content in item list
            if (content == 'hvt' or content == 'lvt' or content == 'rvt' or content == 'slvt'):
                for header_content in header: # heading list for block
                    worksheet.write(row, column, header_content) # write on excel sheet
                    column = column + 1
                row = row + 1
                column = 0
                for header_value in empty_header_value: # value for heading list
                    if(column == 2):
                        worksheet.write(row, column, content.upper()) # threshold value
                        column = column + 1
                    else:
                        worksheet.write(row, column, header_value)
                        column = column + 1
            elif(content == 'lg30' or content == 'lg34' or content == 'lg38' or content == 'lg14' or content == 'lg16' or content == 'lg18'):
                worksheet.write(row, 3, content[2:4])
            elif (content == 'tt_25' or content == 'tt_85' or content == 'TT_25' or content == 'TT_85' or content == 'ss_n40' or content == 'ff_125' or content == 'FFPG_125' or content == 'SSPG_n40'):
                pvt_temp_list = content.split('_')
                worksheet.write(row, 4, pvt_temp_list[0].upper())
                if (content == 'ss_n40' or content == 'SSPG_n40'):
                    worksheet.write(row, 5, '-' + pvt_temp_list[1][1:3] + 'C') # for corner in ss_n40, SSPG_n40, ff_125, FFPG_125
                else:
                    worksheet.write(row, 5, pvt_temp_list[1] + 'C') # for temp in ss_n40, SSPG_n40, ff_125, FFPG_125
            else:
                worksheet.write(row, column, content) # for heading and other values for vdd, delay, iddq and ceff
                column = column + 1
        row = row + 1
    end_index = end_index + 5

workbook.close() # close excel sheet
print ('Execution Completed!!!')
print ('Generated Excel Sheet Name: ', generated_excel_sheet_name, end = '')

'''time.sleep(1)
file = open_workbook(generated_excel_sheet_name).sheet_by_index(0)
time.sleep(1)

workbook_1 = xlsxwriter.Workbook('v21.xlsx')
worksheet_1 = workbook_1.add_worksheet()

for i in range(file.nrows):
    #print (i)
    #print(file.row_values(i))
    v = file.row_values(i)
    if 'failed' not in v:
        #sheet.insert_row(file.row_values(i),i + 1)
        for col in range(0, len(v)):
            print (v[col])
            worksheet_1.write(i, col, v[col])
        #continue
    else:
        continue
        #sheet.insert_row(file.row_values(i),i + 1)
        #for col in range(0, len(v)):
            #worksheet.write(i, col, v[col])
            #ws1.row(i).write(col, v[col])
        #time.sleep(1)
workbook_1.close() '''
