from docxtpl import DocxTemplate



doc = DocxTemplate("C:/Users/furag/OneDrive/Рабочий стол/Python/Шаблоны/РЦК Деппром для прокуратуры11276-1-3.docx")
context = {
        
    'T1' : "World company",
    'T2' : "World company",
    'T3' : "World company",
    'T4' : "World company",
    'T5' : "World company",
    'T6' : "World company",
    'T7' : "World company",
    'T8' : "World company",
    'T9' : "World company",
    'T10' : "World company",
    'T11' : "World company",
    'T12' : "World company",
    'T13' : "World company",
    'T14' : "World company",
    'T15' : "World company",
    'T16' : "World company",
    'T17' : "World company",
        
        }
        
doc.render(context)
doc.save("C:/Users/furag/OneDrive/Рабочий стол/Python/Документы/generated_doc.docx")