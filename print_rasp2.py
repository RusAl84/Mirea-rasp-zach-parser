import win32com.client

def str_count(text, substr):    #количество вхождений подстроки в строку
    return len(text.split(substr))-1


pos=2
fn=[u'c:\\rasp_zach\\1.xlsx',
    u'c:\\rasp_zach\\2.xlsx',
    u'c:\\rasp_zach\\3.xlsx',
    u'c:\\rasp_zach\\4.xlsx',
    u'c:\\rasp_zach\\5.xlsx']

ofn=u'c:\\rasp_zach\\zach.xlsx'
mas= [['23 декабря понедельник', 6, 16], 
      ['24 декабря вторник', 16, 28], 
      ['25 декабря среда', 28, 40],
      ['26 декабря четверг', 40, 52],
      ['27 декабря пятница', 52, 64],
      ['28 декабря суббота', 64, 76]]
Excel = win32com.client.Dispatch("Excel.Application")
owb = Excel.Workbooks.Open(ofn)
osheet = owb.ActiveSheet
for fi in range(0,5):
    print(fn[fi])
    wb = Excel.Workbooks.Open(fn[fi])
    sheet = wb.ActiveSheet
    i=1
    while i<300:
        rw=2 # строка с шифрами групп
        #gr=str(currentSheet.cell(row=rw, column=i).value)
        gr=str(sheet.Cells(rw,i).value)
        #print(gr)
        if len(gr)>4:
            if str_count(gr, '-')>1:
                for mm in range(0,6):
                    den=mas[mm][0]
                    for z in range(mas[mm][1],mas[mm][2]):
                        predm=str(sheet.Cells(z,i).value)
                        if len(predm)>7:
                            kto=str(sheet.Cells(z,i+2).value)
                            gde=str(sheet.Cells(z,i+3).value)
                            gde=gde.replace('.0','')
                            kurs=str(fi+1)
                            vremya=str(sheet.Cells(z,3).value)
                            osheet.Cells(pos,1).value=kto
                            osheet.Cells(pos,2).value=den
                            osheet.Cells(pos,3).value=vremya
                            osheet.Cells(pos,4).value=predm
                            osheet.Cells(pos,5).value=kurs
                            osheet.Cells(pos,6).value=gr
                            osheet.Cells(pos,7).value=gde
                            pos=pos+1
                            #print(den+'  '+gr+'  '+ predm+'  '+ kto+'  '+ gde)
        i=i+1
    wb.Close()    
owb.Save()
owb.Close()
#закрываем COM объект
Excel.Quit()