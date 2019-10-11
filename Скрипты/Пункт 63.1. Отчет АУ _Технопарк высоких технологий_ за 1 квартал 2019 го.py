from docxtpl import DocxTemplate
import os.path

fil = os.path.abspath(__file__)
fil = os.path.join(fil,'..\..\Шаблоны\Пункт 63.1. Отчет АУ _Технопарк высоких технологий_ за 1 квартал 2019 го....docx')
fil = os.path.abspath(fil)
doc = DocxTemplate(fil)
context = {
        'OKUD_form' : "World company",
        'date' : "World company",
        'Consolidated_Registry_Code' : "World company",
        'According_to_OKVED' : "World company",
        'Base_code' : "World company",
        
        'Information_on_the_actual_achievement_of_indicators':[
                {
                        'Unique_number' : "World company",
                        'Content_of_basic_service' : "World company",
                        'Consumer_category' : "World company",
                        'T4' : "World company",
                        'Service_form' : "World company",
                        'Service_charge' : "World company",
                        'name_indicator' : "World company",
                        'name' : "World company",
                        'OKEY_Code' : "World company",
                        'approved_in_the_state_assignment_for_the_year' : "World company",
                        'approved_in_the_state_assignment_at_the_reporting_date' : "World company",
                        'executed_at_the_reporting_date' : "World company",
                        'permissible_variation' : "World company",
                        'excess_deviation' : "World company",
                        'reason_for_rejection' : "World company",
                },
        ],
        'Information_on_the_actual_achievement_of_indicators_characterizing_the_volume_of_public_services1':[
                {
                        'Unique_registry_entry_number' : "World company",
                        'Content_of_basic_service' : "World company",
                        'Consumer_category' : "World company",
                        'T4' : "World company",
                        'Service_form' : "World company",
                        'Service_charge' : "World company",
                        'Name_of_indicator1' : "World company",
                        'Name_of_indicator2' : "World company",
                        'name' : "World company",
                        'OKEY_Code' : "World company",
                        'approved_in_the_state_assignment_for_the_year1' : "World company",
                        'approved_in_the_state_assignment_for_the_year2' : "World company",
                        'approved_in_the_state_assignment_at_the_reporting_date1' : "World company",
                        'approved_in_the_state_assignment_at_the_reporting_date2' : "World company",
                        'executed_at_the_reporting_date1' : "World company",
                        'executed_at_the_reporting_date2' : "World company",
                        'permissible_variation1' : "World company",
                        'permissible_variation2' : "World company",
                        'excess_deviation1' : "World company",
                        'excess_deviation2' : "World company",
                        'reason_for_rejection1' : "World company",
                        'reason_for_rejection2' : "World company",
                        'Average_annual_fee1' : "World company",
                        'Average_annual_fee2' : "World company",
                },
        ],
        'Information_on_the_actual_achievement_of_indicators_characterizing_the_volume_of_public_services2':[
                {
                        'Unique_number' : "World company",
                        'Content_of_basic_service' : "World company",
                        'Consumer_category' : "World company",
                        'T4' : "World company",
                        'Service_form' : "World company",
                        'Service_charge' : "World company",
                        'name_indicator' : "World company",
                        'name' : "World company",
                        'OKEY_Code' : "World company",
                        'approved_in_the_state_assignment_for_the_year' : "World company",
                        'approved_in_the_state_assignment_at_the_reporting_date' : "World company",
                        'executed_at_the_reporting_date' : "World company",
                        'permissible_variation' : "World company",
                        'excess_deviation' : "World company",
                        'reason_for_rejection' : "World company",
                        'Average_annual_fee' : "World company",
                },
        ],
        }
doc.render(context)
fil = os.path.join(fil,'..\..\Документы\generated_doc.docx')
fil = os.path.abspath(fil)
doc.save(fil)