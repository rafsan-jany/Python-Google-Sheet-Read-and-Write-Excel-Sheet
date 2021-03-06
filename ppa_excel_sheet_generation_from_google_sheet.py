''' This program reads Google Sheet values using Google Sheet API and then writes data on a Excel Sheet in a specific format'''
''' command: python ppa_excel_sheet_generation_from_google_sheet.py 40LP 9T V1.1_4 1.1 All_data_sheet_for_ppa 40LP_V1.1_4_9T''' 


import gspread
import xlsxwriter
import sys
from oauth2client.service_account import ServiceAccountCredentials
from pprint import pprint
from xlrd import open_workbook
import time

# for command line input
process = sys.argv[1]
track = sys.argv[2]
pdk_version = sys.argv[3]
vnom = sys.argv[4]
google_spread_name = sys.argv[5]
google_work_sheet_name = sys.argv[6]
# print (google_work_sheet_name)

# declared list for heading 
header = ['Process', 'Track', 'Threshold', 'Lg', 'PVT Corner', 'Temp', 'PDK Version', 'Vnom']
# heading values
#empty_header_value = ['12LP+','7.5T','','','','','1.0','0.8'] 
empty_header_value = [process,track,'','','','',pdk_version,vnom]
# print (empty_header_value)

print ('Execution Started...')
scope = ["https://spreadsheets.google.com/feeds",'https://www.googleapis.com/auth/spreadsheets',"https://www.googleapis.com/auth/drive.file","https://www.googleapis.com/auth/drive"]

creds = ServiceAccountCredentials.from_json_keyfile_name("creds.json", scope)
client = gspread.authorize(creds)

#google_spread_name = "12LPPLUS_V1.0_7P5T"

# Open the google spreadhseet
#sheet = client.open(google_spread_name).sheet1
#print (client.open(google_spread_name).worksheets())

sheet = client.open(google_spread_name).worksheet(google_work_sheet_name) # name for sheet
#sheet = client.open_by_url('https://docs.google.com/spreadsheets/d/1zsbqFlKmmQHJJmhRM9N_64x4Fa-5RlYjKKHKbs59VSE/edit#gid=559268704').worksheet(google_work_sheet_name) # using google sheet link

# Get a list of all records 
data = sheet.get_all_values()
#print ('Processed Row ', len(data))

# Excel sheet name
#generated_excel_sheet_name = '12LPPLUS_V1.0_7P5T.xlsx'
generated_excel_sheet_name = google_work_sheet_name + '.xlsx' # same name as sheet name
  
# Workbook() takes one, non-optional, argument which is the excel sheet name that we want to create. 
workbook = xlsxwriter.Workbook(generated_excel_sheet_name) 
  
# The workbook object is then used to add new worksheet via the add_worksheet() method. 
worksheet = workbook.add_worksheet()

start_index = 0 
end_index = 5 
row = 0
column = 0

# assume highest length of row is 40, each 5 contains a block until the next empty list
for start_index in range(0,40,5): 
    #print ('start_index ', start_index)
    #print ('end_index ', end_index)
    for each_list in data:
        #print (each_list[start_index : end_index])
        item = each_list[start_index : end_index] # each item list contains 5 elements, each_list contains a row from google sheet
        column = 0
        if 'failed' not in item:
            # print (item)
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
                #elif(content == 'lg30' or content == 'lg34' or content == 'lg38' or content == 'lg14' or content == 'lg16' or content == 'lg18' or content == 'lg20'):
                elif (content in {'lg30', 'lg34', 'lg38', 'lg14', 'lg16', 'lg18', 'lg20', 'lg24', 'lg28', 'lg32', 'lg36', 'Lg14', 'Lg16', 'Lg40'}):
                    worksheet.write(row, 3, content[2:4] + 'nm')
                #elif (content == 'tt_25' or content == 'tt_85' or content == 'TT_25' or content == 'TT_85' or content == 'ss_n40' or content == 'ff_125' or content == 'FFPG_125' or content == 'SSPG_n40'):
                elif(content in {'tt_25', 'tt_85', 'TT_25', 'TT_85', 'ss_n40', 'ff_125', 'FFPG_125', 'SSPG_n40', 'ffg_125', 'ssg_n40'}):
                    pvt_temp_list = content.split('_')
                    corner_value = pvt_temp_list[0].upper()
                    worksheet.write(row, 4, pvt_temp_list[0].upper())
                    if (corner_value in {'SS', 'SSG', 'SSPG'}):
                        #worksheet.write(row, 7, '0.72') # vnom value set
                        worksheet.write(row, 5, '-' + pvt_temp_list[1][1:3] + 'C') # for temperature in ss_n40, SSPG_n40,  ssg_n40
                    elif(corner_value in {'FF', 'FFG', 'FFPG'}):
                        #worksheet.write(row, 7, '0.88') # vnom value set
                        worksheet.write(row, 5, pvt_temp_list[1] + 'C') # for temp in ff_125, ffg_125, FFPG_125
                    else:
                        #worksheet.write(row, 7, '0.8') # vnom value set
                        worksheet.write(row, 5, pvt_temp_list[1] + 'C') # for temp in tt_25, tt_85
                elif(content in {'tt25c', 'tt85c'}):
                    corner = content[0:2]
                    worksheet.write(row, 4, corner.upper())
                    temperature = content[2:4]
                    worksheet.write(row, 5, temperature + 'C')
                    #worksheet.write(row, 7, '0.8') # vnom value set
                elif(content in {'ffgp125c'}):
                    corner = content[0:4]
                    worksheet.write(row, 4, corner.upper())
                    temperature = content[4:7]
                    worksheet.write(row, 5, temperature + 'C')
                    #worksheet.write(row, 7, '0.88') # vnom value set
                elif content == 'ssgn40':
                    corner = content[0:3]
                    worksheet.write(row, 4, corner.upper())
                    temperature = content[4:6]
                    worksheet.write(row, 5, '-' + temperature + 'C')
                    #worksheet.write(row, 7, '0.72') # vnom value set
                else:
                    worksheet.write(row, column, content) # for heading and other values for vdd, delay, iddq and ceff
                    column = column + 1
            row = row + 1
        else:
            continue
    end_index = end_index + 5

workbook.close() # close excel sheet
print ('Execution Completed!!!')
print ('Generated Excel Sheet Name: ', generated_excel_sheet_name, end = '')
