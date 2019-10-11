import openpyxl
import os
# rows = {
#     'Name of indicator':'d',
#     'categories':{
#         'Category':'',
#         'rows':{
#             'responsible unit':'',
#             'types of work':'',
#             'result ':'',
#             'necessary':'',
#             'done':'',
#             'percent':'',

#             }
#         }
#     }
fil = os.path.abspath(__file__)

tab={
        'indicator1':
            {
                'category1':[["1",'1','1','1','1','1',],['2','2','2','2','2','2',],['3','3','3','3','3','3',]],
                'category2':[["dsfs",'s','s','s','s','s',],['d','d','d','d','d','d',],['f','f','f','f','f','f',]],
                
            },
        'indicator2':
            {
                'category1':[["1",'1','1','1','1','1',],['2','2','2','2','2','2',],['3','3','3','3','3','3',]],
                'category2':[['s','s','s','s','s','s',],['d','d','d','d','d','d',],['f','f','f','f','f','f',]],
                
            },
        
    }


wb = openpyxl.Workbook()
ws = wb.active

ws.cell(row=1,column=2).value='Отчет о выполнении государственного задания в 2019 году'
ws.merge_cells(start_row=1,start_column=2,end_row=1,end_column=5)
ws.cell(row=1,column=6).value='III квартал'
ws.merge_cells(start_row=1,start_column=6,end_row=1,end_column=8)

rows=['Наименование показателя','Категория','Ответственное подразделение','Виды работ','Результат','Необходимо','Выполнено','Процент']
for i, value in enumerate(rows,1):
    ws.cell(row=2,column= i).value=value



sh=0
sc=0
for i, value_tab in enumerate(list(tab.keys()),1):
    ws.cell(row=3+sh,column=1).value=value_tab
    sc = 3+sh
    print(3+sh)
    for y, value_category in enumerate(list(tab[value_tab].keys()),1):
        ws.cell(row=3+sh,column=2).value=value_category


        for stroc, value_rows in enumerate(tab[value_tab][value_category],3):

            for stolb, row in enumerate(value_rows,3):

                for z, value in enumerate(row,1):
                    cell = ws.cell(row = stroc+sh,column=stolb )
                    cell.value = value
        ws.merge_cells(start_row=3+sh,start_column=2,end_row=sh+stroc,end_column=2)

        
        
        
        sh+=stroc-2
        

    ws.cell(row=3+sh,column=1).value='Итого:'
    ws.merge_cells(start_row=3+sh,start_column=1,end_row=3+sh,end_column=5)
    ws.merge_cells(start_row=sc,start_column=1,end_row=2+sh,end_column=1)
    
    print(sc)
    print(sh)
    sh+=1        
    
   

fil = os.path.join(fil,'..\..\Документы')
fil = os.path.abspath(fil)
os.chdir(path=fil)
wb.save('02_Ezhenedelny_otchet_GZ_2019_1.xlsx')



