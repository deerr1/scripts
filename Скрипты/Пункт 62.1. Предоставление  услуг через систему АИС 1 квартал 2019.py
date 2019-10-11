from docxtpl import DocxTemplate



doc = DocxTemplate("C:/Users/furag/OneDrive/Рабочий стол/Python/Шаблоны/Пункт 62.1. Предоставление  услуг через систему АИС 1 квартал 2019.docx")
context = {
     'event_title1' : 'World company',
     'implementation1' : "World company",
     'event_title2' : "World company" ,
     'implementation2' : "World company", 
     'event_title3' : "World company", 
     'implementation3' : "World company", 

     'working_units' : "World company", 
     'workplaces' : "World company", 
     'volume1' : "World company", 
     'volume2' : "World company", 
     'volume3' : "World company", 
     'number_of_contracts' : "World company", 
     'amount_of_payments' : "World company", 
     'investment_size' : "World company", 
     'own_funds' : "World company", 
     'involved_funds' : "World company", 
     'gos_support' : "World company", 
     'extra_funds' : "World company", 
     'number_of_app' : "World company", 
     'number_of_doc' : "World company", 
     'research_costs' : "World company", 
     'number_of_dev' : "World company", 
     'number_of_samples' : "World company", 

     'production_research' : "World company", 
     'business_plan' : "World company", 
     'Assistance_in_patenting' : "World company", 
     'Participation_in_events' : "World company", 
     'Prototype_development' : "World company", 
     'Product_certification' : "World company", 
     'Rental_of_premises' : "World company", 
     'other' : "World company", 
        }
        
doc.render(context)
doc.save("C:/Users/furag/OneDrive/Рабочий стол/Python/Документы/generated_doc.docx")