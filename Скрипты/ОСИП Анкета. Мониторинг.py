from docxtpl import DocxTemplate
import os.path

fil = os.path.abspath(__file__)
fil = os.path.join(fil,'..\..\Шаблоны\ОСИП Анкета. Мониторинг.docx')
fil = os.path.abspath(fil)

doc = DocxTemplate(fil)
context = {
     'Project_implementation':[
                {
                        'Number' : "FF",
                        'event_title' : "FF",
                        'Actual_execution' : "FF",
                },
                
                
                
        ],

     'working_units' : "FF", 
     'workplaces' : "FF", 
     'volume1' : "FF", 
     'volume2' : "FF", 
     'volume3' : "FF", 
     'number_of_contracts' : "FF", 
     'amount_of_payments' : "FF", 
     'investment_size' : "FF", 
     'own_funds' : "FF", 
     'involved_funds' : "FF", 
     'gos_support' : "FF", 
     'extra_funds' : "FF", 
     'number_of_app' : "FF", 
     'number_of_doc' : "FF", 
     'research_costs' : "FF", 
     'number_of_dev' : "FF", 
     'number_of_samples' : "FF", 

     'production_research' : "FF", 
     'business_plan' : "FF", 
     'Assistance_in_patenting' : "FF", 
     'Participation_in_events' : "FF", 
     'Prototype_development' : "FF", 
     'Product_certification' : "FF", 
     'Rental_of_premises' : "FF", 
     'other' : "FF", 
        }
        
doc.render(context)
fil = os.path.join(fil,'..\..\Документы\generated_doc.docx')
fil = os.path.abspath(fil)
doc.save(fil)