from docxtpl import DocxTemplate
import os.path

fil = os.path.abspath(__file__)
fil = os.path.join(fil,'..\..\Шаблоны\Пункт 66.2. Мониторинг результатов реализации инновационных проектов.docx')
fil = os.path.abspath(fil)


doc = DocxTemplate(fil)
context = {
        
        'tbl_contents':[
                {
                        'number' : "World company",
                        'Applicant' : "World company",
                        'Application_Registration_Number' : "World company",
                },
                
                
                
        ],
        'tbl_contents2':[
                {
                        'number' : "World company",
                        'Applicant' : "World company",
                        'Project' : "World company",
                },
        ],
        'tbl_contents3':[
                {
                        'number' : "World company",
                        'Applicant' : "World company",
                        'Project' : "World company",
                        'Reward' : "World company",
                },
        ],
        'tbl_contents4':[
                {
                        'number' : "World company",
                        'Applicant_project_name' : "World company",
                        'Kit' : "World company",
                        'Equipment' : "World company",
                },
        ],
        'date':'sss',
        
        }
        
doc.render(context)
fil = os.path.join(fil,'..\..\Документы\generated_doc.docx')
fil = os.path.abspath(fil)
doc.save(fil)