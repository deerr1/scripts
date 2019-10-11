from docxtpl import DocxTemplate
import os.path

fil = os.path.abspath(__file__)
fil = os.path.join(fil,'..\..\Шаблоны\Пункт 65.1. Консультации по участию в программах Фонда 1 квартал 2019.docx')
fil = os.path.abspath(fil)

doc = DocxTemplate(fil)
context = {
        
        'Consultations_for_potential_participants_of_the_UMNIK_program':[
                {
                        'number' : "World company",
                        'Date' : "World company",
                        'Place' : "World company",
                        'Number_of_consulted' : "World company",
                },
                
                
                
        ],
        'summa':'22',
        'Applications_filed_for_the_program_Start_1_19_turn_1':[
                {
                        'number' : "World company",
                        'name_of_the_project' : "World company",
                        'Direction' : "World company",
                        'City' : "World company",
                        'date_of_application' : "World company",
                        'Applicant' : "World company",
                },
        ],
        'date':'ss'
        }
        
doc.render(context)
fil = os.path.join(fil,'..\..\Документы\generated_doc.docx')
fil = os.path.abspath(fil)
doc.save(fil)