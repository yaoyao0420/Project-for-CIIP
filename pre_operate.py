__author__ = 'ZHANG Junyao'
import xlrd
import xlwt
import chardet
import traceback
import openpyxl
def getColumnIndex(table, columnName):
    columnIndex = None
    #print table
    for i in range(table.ncols):
        #print columnName
        #print table.cell_value(0, i)
        if(table.cell_value(0, i) == columnName):
            columnIndex = i
            break
    return columnIndex

def readExcelDataByName(fileName, sheetName):
    #print fileName
    table = None
    errorMsg = ""
    try:
        data = xlrd.open_workbook(fileName)
        table = data.sheet_by_name(sheetName)
    except Exception as msg:
        errorMsg = msg
    return table, errorMsg

    
if __name__ == '__main__':

    xlsfile = 'E:/1.xls'
    table = readExcelDataByName(xlsfile, 'Sheet1')[0]
    #get certain columns' value and write them into a new worksheet named "Sheet1"   
    wbook = xlwt.Workbook(encoding = 'utf-8')
    # create a sheet which we want to write in
    worksheet = wbook.add_sheet(u'Sheet1',cell_overwrite_ok=True)
    post_id = []
    jurisdicti = []
    on_off_str = []
    osp_id = []
    streetname = []
    location = []
    
    #print (table.nrows)
    #print (table.cell_value(0, getColumnIndex(table,'POST_ID')))
    for i in range(table.nrows):
        post_id.append(table.cell_value(i, getColumnIndex(table,'POST_ID')))
        jurisdicti.append(table.cell_value(i, getColumnIndex(table,'JURISDICTI')))
        on_off_str.append(table.cell_value(i, getColumnIndex(table,'ON_OFF_STR')))
        osp_id.append(table.cell_value(i, getColumnIndex(table,'OSP_ID')))
        streetname.append(table.cell_value(i, getColumnIndex(table,'STREETNAME')))
        location.append(table.cell_value(i, getColumnIndex(table, 'LOCATION')))
        
    #print (post_id[1])
    for s in range(len(post_id)):
        worksheet.write(s,0,post_id[s])
        worksheet.write(s,1,jurisdicti[s])
        worksheet.write(s,2,on_off_str[s])
        worksheet.write(s,3,osp_id[s])
        worksheet.write(s,4,streetname[s])
        worksheet.write(s,5,location[s])

    wbook.save('E:/2.xls')

    #this time we will deal with the data with different character
    #for we want to get toatl meters per area, so we use Port, SFMTA, on street and off street this four indicators.
    #and we get four sheets:
    #(1) Port & on street (POST_ID, STREETNAME and LOCATION)
    #(2) Port & off street (POST_ID, OSP_ID and LOCATION)
    #(3) SFMTA & on street (POST_ID, STREETNAME and LOCATION)
    #(4) SFMTA & off street (POST_ID, OSP_ID and LOCATION)

    xlsfile2 = 'E:/2.xls'
    table2 = readExcelDataByName(xlsfile2, 'Sheet1')[0]

    wbook_ = xlwt.Workbook(encoding = 'utf-8')
    wsheet_1 = wbook_.add_sheet(u'PORT&ON',cell_overwrite_ok=True)
    wsheet_2 = wbook_.add_sheet(u'PORT&OFF',cell_overwrite_ok=True)
    wsheet_3 = wbook_.add_sheet(u'SFMTA&ON',cell_overwrite_ok=True)
    wsheet_4 = wbook_.add_sheet(u'SFMTA&OFF',cell_overwrite_ok=True)

    post_id_1 = []
    post_id_2 = []
    post_id_3 = []
    post_id_4 = []
    osp_id_1 = []
    streetname_1 = []
    osp_id_2 = []
    streetname_2 = []
    location_1 = []
    location_2 = []
    location_3 = []
    location_4 = []

    #streetname_1.append(table2.cell_value(0, getColumnIndex(table2,'STREETNAME')))
    #streetname_2.append(table2.cell_value(0, getColumnIndex(table2,'STREETNAME')))
    #osp_id_1.append(table2.cell_value(0, getColumnIndex(table2,'OSP_ID')))
    #osp_id_1.append(table2.cell_value(0, getColumnIndex(table2,'OSP_ID')))

    for i in range(table2.nrows):
        if i == 0:
            continue
        if table2.cell_value(i, getColumnIndex(table2,'ON_OFF_STR')) == 'ON':
            if table2.cell_value(i, getColumnIndex(table2,'JURISDICTI')) == 'PORT':
                post_id_1.append(table.cell_value(i, getColumnIndex(table2,'POST_ID')))
                streetname_1.append(table2.cell_value(i, getColumnIndex(table2,'STREETNAME')))
                location_1.append(table2.cell_value(i, getColumnIndex(table2,'LOCATION')))
            else:
                post_id_3.append(table.cell_value(i, getColumnIndex(table2,'POST_ID')))
                streetname_2.append(table2.cell_value(i, getColumnIndex(table2,'STREETNAME')))
                location_3.append(table2.cell_value(i, getColumnIndex(table2,'LOCATION')))

        else:
            if table2.cell_value(i, getColumnIndex(table2, 'JURISDICTI')) == 'PORT':
                post_id_2.append(table.cell_value(i, getColumnIndex(table2,'POST_ID')))
                osp_id_1.append(table2.cell_value(i, getColumnIndex(table2,'OSP_ID')))
                location_2.append(table2.cell_value(i, getColumnIndex(table2,'LOCATION')))
            else:
                post_id_4.append(table.cell_value(i, getColumnIndex(table2,'POST_ID')))
                osp_id_2.append(table2.cell_value(i, getColumnIndex(table2,'OSP_ID')))
                location_4.append(table2.cell_value(i, getColumnIndex(table2,'LOCATION')))
    
    for s in range(len(streetname_1)):
        wsheet_1.write(s,0,post_id_1[s])
        wsheet_1.write(s,1,streetname_1[s])
        wsheet_1.write(s,2,location_1[s])
    for s in range(len(streetname_2)):
        wsheet_3.write(s,0,post_id_3[s])
        wsheet_3.write(s,1,streetname_2[s])
        wsheet_3.write(s,2,location_3[s])
    for s in range(len(osp_id_1)):
        wsheet_2.write(s,0,post_id_2[s])
        wsheet_2.write(s,1,osp_id_1[s])
        wsheet_2.write(s,2,location_2[s])
    for s in range(len(osp_id_2)):
        wsheet_4.write(s,0,post_id_4[s])
        wsheet_4.write(s,1,osp_id_2[s])
        wsheet_4.write(s,2,location_4[s])


    wbook_.save('E:/3.xls')
            
                
                
    


    
