import openpyxl, pprint

print('Opening workbook...')

wb = openpyxl.load_workbook('Report.xlsx')
sheet = wb['report1606497624523']

salesFunnel = {}

print('Reading rows...')

#for row in range(2, sheet.min_row + 1):
 #   leadId = sheet['A' + str(row)].value
  #  email = sheet['D' + str(row)].value
   # leadStatus = sheet['E' + str(row)].value
    #createdDate = sheet['F' + str(row)].value

def clean_list(sheet):
#Loop over e rows, and check each cell for a match.
    for row in range(2, sheet.max_row+1):
        for column in "D":
            cell_name = "{}{}".format(column, row)
        if sheet[cell_name].value == sheet[cell_name].value:
    #Check code should be here
            print(clean_list(sheet))


#salesFunnel.setdefault(leadId, {})
#salesFunnel[leadId].setdefault(email,{})
#salesFunnel[email].setdefault(leadStatus, {})
#salesFunnel[leadStatus].setdefault(createdDate, {})

print('Writing results...')

resultFile = open('reportResults.py', 'w')
resultFile.write('salesFunnelData = ' + pprint.pformat(sheet))
resultFile.close()
print('Done.')