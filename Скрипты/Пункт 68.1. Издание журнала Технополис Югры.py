from docxtpl import DocxTemplate



doc = DocxTemplate("C:/Users/furag/OneDrive/Рабочий стол/Python/Шаблоны/Пункт 68.1. Издание журнала Технополис Югры.docx")
context = {
        
        # 'tbl_contents':[
        #         {
        #                 'T1' : "World company",
        #                 'T2' : "World company",
        #                 'T3' : "World company",
        #         },
                
                
                
        # ],
        # 'tbl_contents2':[
        #         {
        #                 'T1' : "World company",
        #                 'T2' : "World company",
        #                 'T3' : "World company",
        #         },
        # ],
        # 'tbl_contents3':[
        #         {
        #                 'T1' : "World company",
        #                 'T2' : "World company",
        #                 'T3' : "World company",
        #                 'T4' : "World company",
        #         },
        # ],
        # 'tbl_contents4':[
        #         {
        #                 'T1' : "World company",
        #                 'T2' : "World company",
        #                 'T3' : "World company",
        #                 'T4' : "World company",
        #         },
        # ],
        'date':'sss',
        
        }
        
doc.render(context)
doc.save("C:/Users/furag/OneDrive/Рабочий стол/Python/Документы/generated_doc.docx")