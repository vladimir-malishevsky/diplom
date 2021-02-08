import requests
import json
import openpyxl

def get():
    with requests.get('http://127.0.0.1:8000/api/PatientAll/?format=json') as response:
        return json.loads(response.text)
# def get():
#     with open('data.json') as file:
#         return json.load(file)


book = openpyxl.Workbook()
sheet = book.active

sheet['A1'] = 'NAME'
sheet['B1'] = 'SURNAME'
sheet['C1'] = 'TREATMENT_SESSION_startSession'
sheet['D1'] = 'TREATMENT_SESSION_mainIll'
sheet['E1'] = 'TREATMENT_SESSION_doctor'
sheet['F1'] = 'TREATMENT_SESSION_comorbidity'
sheet['G1'] = 'STAGE_OF_TREATMENT_stageName'
sheet['H1'] = 'STAGE_OF_TREATMENT_startStage'
sheet['I1'] = 'STAGE_OF_TREATMENT_surgery'
sheet['J1'] = 'STAGE_OF_TREATMENT_pharmacotherapy'
sheet['K1'] = 'STAGE_OF_TREATMENT_physiotherapy'
sheet['L1'] = 'STAGE_OF_TREATMENT_electro_ultrasound_therapy'
sheet['M1'] = 'STAGE_OF_TREATMENT_stateOn'
sheet['N1'] = 'STAGE_OF_TREATMENT_laboratory_tests'
sheet['O1'] = 'STAGE_OF_TREATMENT_state_measurement'
sheet['P1'] = 'STAGE_OF_TREATMENT_endStage'
sheet['Q1'] = 'STAGE_OF_TREATMENT_endSession'

sheet.column_dimensions['A'].width = 15  # NAME
sheet.column_dimensions['B'].width = 15  # SURNAME
sheet.column_dimensions['C'].width = 35  # TREATMENT_SESSION_startSession
sheet.column_dimensions['D'].width = 70  # TREATMENT_SESSION_mainIll
sheet.column_dimensions['E'].width = 30  # TREATMENT_SESSION_doctor
sheet.column_dimensions['F'].width = 70  # TREATMENT_SESSION_comorbidity
sheet.column_dimensions['G'].width = 70  # STAGE_OF_TREATMENT_stageName
sheet.column_dimensions['H'].width = 70  # STAGE_OF_TREATMENT_startStage
sheet.column_dimensions['I'].width = 90  # STAGE_OF_TREATMENT_surgery
sheet.column_dimensions['J'].width = 90  # STAGE_OF_TREATMENT_pharmacotherapy
sheet.column_dimensions['K'].width = 70  # STAGE_OF_TREATMENT_physiotherapy
sheet.column_dimensions['L'].width = 70  # STAGE_OF_TREATMENT_electro_ultrasound_therapy
sheet.column_dimensions['M'].width = 30  # STAGE_OF_TREATMENT_stateOn
sheet.column_dimensions['N'].width = 200  # STAGE_OF_TREATMENT_state_laboratory_tests
sheet.column_dimensions['O'].width = 200  # STAGE_OF_TREATMENT_state_measurement
sheet.column_dimensions['P'].width = 60  # STAGE_OF_TREATMENT_endStage
sheet.column_dimensions['Q'].width = 60  # STAGE_OF_TREATMENT_endSession

row = 2
for patient in get():
    sheet[row][0].value = patient['name']
    sheet[row][1].value = patient['surname']
    for treatment_session in patient['treatment_session']:

        sheet[row][2].value = treatment_session['startSession']
        sheet[row][3].value = treatment_session['mainIll']
        sheet[row][4].value = treatment_session['doctor']

        comorbidity_str = ''
        for comorbidity in treatment_session['comorbidity']:
            comorbidity_str += comorbidity['nameIll'] + ' '
        sheet[row][5].value = comorbidity_str

        for stage_of_treatment in treatment_session['stage_of_treatment']:
            sheet[row][6].value = stage_of_treatment['stageName']
            sheet[row][7].value = stage_of_treatment['startStage']

            surgery_str = ''
            for surgery in stage_of_treatment['surgery']:
                surgery_str += surgery['DateIntervention'] + ' ' + surgery['nameInterven'] + ' '
                for sur_med_staff in surgery['sur_med_staff']:
                    surgery_str += sur_med_staff['medStaff']
            sheet[row][8].value = surgery_str

            pharmacotherapy_str = ''
            for pharmacotherapy in stage_of_treatment['pharmacotherapy']:
                pharmacotherapy_str += pharmacotherapy['namePill'] + ' ' + str(pharmacotherapy['dosePill']) + ' ' + \
                                       str(pharmacotherapy['unitPill']) + ' ' + str(pharmacotherapy['datePill']) + ' '
            sheet[row][9].value = pharmacotherapy_str

            physiotherapy_str = ''
            for physiotherapy in stage_of_treatment['physiotherapy']:  # test
                physiotherapy_str += physiotherapy['id_physiot'] + ' ' + \
                                     physiotherapy['id_stage'] + ' ' + \
                                     physiotherapy['name_physiotherapy'] + ' ' + \
                                     physiotherapy['value_physiotherapy'] + ' ' + \
                                     physiotherapy['unit_physiotherapy'] + ' ' + \
                                     physiotherapy['date_physiotherapy'] + ' ' + \
                                     physiotherapy['id_med_staff'] + ' '
            sheet[row][10].value = physiotherapy_str

            for electro_ultrasound_therapy in stage_of_treatment['electro_ultrasound_therapy']:  # test
                pass

            for state in stage_of_treatment['state']:
                sheet[row][12].value = state['state_on']

                laboratory_test_str = ''
                for laboratory_test in state['laboratory_test']:
                    laboratory_test_str += str(laboratory_test['nameTest']) + ' ' + str(laboratory_test['valueTest']) + ' ' + \
                                          str(laboratory_test['unitTest']) + ' ' + str(laboratory_test['dateTest']) + ' ' + \
                                          str(laboratory_test['laboratoryName']) + ' '
                    for lab_staff in laboratory_test['lab_staff']:  # test
                        laboratory_test_str += lab_staff['id_lab_test'] + ' ' + lab_staff['id_med_staff'] + ' '

                sheet[row][13].value = laboratory_test_str

                measurement_str = ''
                for measurement in state['measurement']:
                    measurement_str += str(measurement['nameMeasurement']) + ' ' + str(measurement['valueMeasurement']) + ' ' + \
                                      str(measurement['unitMeasurement']) + ' ' + str(measurement['dateMeasurement']) + ' ' + \
                                      str(measurement['medStaff']) + ' '
                sheet[row][14].value = measurement_str

                for image in state['image']:
                    pass


            sheet[row][15].value = stage_of_treatment['endStage']
        sheet[row][16].value = treatment_session['endSession']
    row += 1

book.save('test.xlsx')
book.close()

print("Finish")
