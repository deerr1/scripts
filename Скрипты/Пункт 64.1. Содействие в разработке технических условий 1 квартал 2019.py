from docxtpl import DocxTemplate



doc = DocxTemplate("C:/Users/furag/OneDrive/Рабочий стол/Python/Шаблоны/Пункт 64.1. Содействие в разработке технических условий 1 квартал 2019.docx")
context = {
     'text' : 'World company',
     'date' : "World company",
     
        }
        
doc.render(context)
doc.save("C:/Users/furag/OneDrive/Рабочий стол/Python/Документы/generated_doc.docx")