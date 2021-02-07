import requests
import json
import openpyxl
from openpyxl.styles import PatternFill

def get():
    with requests.get('http://127.0.0.1:8000/api/PatientAll/?format=json') as response:
        return json.loads(response.text)
# def get():
#     with open('data.json') as file:
#         return json.load(file)


book = openpyxl.Workbook()
sheet = book.active

sheet.merge_cells('A1:A4')
sheet.merge_cells('B1:B4')
sheet['A1'] = 'NAME'
sheet['B1'] = 'SURNAME'

sheet.merge_cells('C1:G1')
sheet['C1'] = 'TREATMENT_SESSION'

sheet.merge_cells('C2:C4')
sheet['C2'] = 'STARTSESSION'

sheet.merge_cells('D2:D4')
sheet['D2'] = 'MAINILL'

sheet.merge_cells('E2:E4')
sheet['E2'] = 'DOCTOR'

sheet.merge_cells('F2:F4')
sheet['F2'] = 'COMORBIDITY'

sheet.merge_cells('G2:P2')
sheet['G2'] = 'STAGE_OF_TREATMENT'

sheet.merge_cells('G3:G4')
sheet['G3'] = 'STAGENAME'

sheet.merge_cells('H3:H4')
sheet['H3'] = 'STARTSTAGE'

sheet.merge_cells('I3:I4')
sheet['I3'] = 'SURGERY'

sheet.merge_cells('J3:J4')
sheet['J3'] = 'PHARMACOTHERAPY'

sheet.merge_cells('K3:K4')
sheet['K3'] = 'PHYSIOTHERAPY'

sheet.merge_cells('L3:L4')
sheet['L3'] = 'ELECTRO_ULTRASOUND_THERAPY'

sheet.merge_cells('M3:O3')
sheet['M3'] = 'STATE'

sheet['M4'] = 'STATE ON'
sheet['N4'] = 'LABORATORY TEST'
sheet['O4'] = 'MEASUREMENT'

sheet.merge_cells('P3:P4')
sheet['P3'] = 'ENDSTAGE'

sheet.merge_cells('Q3:Q4')
sheet['Q3'] = 'ENDSESSION'

sheet.column_dimensions['A'].width = 15  # name
sheet.column_dimensions['B'].width = 15  # surname
sheet.column_dimensions['C'].width = 15  # start
sheet.column_dimensions['D'].width = 35  # mainil
sheet.column_dimensions['E'].width = 15  # doctor
sheet.column_dimensions['F'].width = 35  # comorbidity
sheet.column_dimensions['G'].width = 15  # stagename
sheet.column_dimensions['H'].width = 30  # startstage
sheet.column_dimensions['I'].width = 60  # surgery
sheet.column_dimensions['J'].width = 60  # pharmacotherapy
sheet.column_dimensions['K'].width = 15  # physiotherapy
sheet.column_dimensions['L'].width = 40  # electro_ultrasound_therapy
sheet.column_dimensions['M'].width = 20  # state
sheet.column_dimensions['N'].width = 85  #
sheet.column_dimensions['O'].width = 80  #
sheet.column_dimensions['P'].width = 60  #endstage
sheet.column_dimensions['Q'].width = 60  #endsession

row = 5
for patient in get():
    comorbidity_index, surgery_index, pharmacotherapy_index, physiotherapy_index, state_index = row, row, row, row, row
    sheet[row][0].value = patient['name']
    sheet[row][1].value = patient['surname']
    for treatment_session in patient['treatment_session']:

        sheet[row][2].value = treatment_session['startSession']
        sheet[row][3].value = treatment_session['mainIll']
        sheet[row][4].value = treatment_session['doctor']

        comorbidity_str = ''
        comorbidity_index = row
        for comorbidity in treatment_session['comorbidity']:
            comorbidity_str += comorbidity['nameIll']
            sheet[comorbidity_index][5].value = comorbidity_str
            comorbidity_index += 1

        for stage_of_treatment in treatment_session['stage_of_treatment']:
            sheet[row][6].value = stage_of_treatment['stageName']
            sheet[row][7].value = stage_of_treatment['startStage']

            surgery_str = ''
            surgery_index = row
            for surgery in stage_of_treatment['surgery']:
                surgery_str = surgery['DateIntervention'] + ' ' + surgery['nameInterven']
                for sur_med_staff in surgery['sur_med_staff']:
                    surgery_str += sur_med_staff['medStaff']
                sheet[surgery_index][8].value = surgery_str
                surgery_index += 1

            pharmacotherapy_str = ''
            pharmacotherapy_index = row
            for pharmacotherapy in stage_of_treatment['pharmacotherapy']:
                pharmacotherapy_str = pharmacotherapy['namePill'] + ' ' + str(pharmacotherapy['dosePill']) + ' ' + \
                                       str(pharmacotherapy['unitPill']) + ' ' + str(pharmacotherapy['datePill'])
                sheet[pharmacotherapy_index][9].value = pharmacotherapy_str
                pharmacotherapy_index += 1

            physiotherapy_str = ''
            physiotherapy_index = row
            for physiotherapy in stage_of_treatment['physiotherapy']:  # test
                physiotherapy_str = physiotherapy['id_physiot'] + ' ' + \
                                     physiotherapy['id_stage'] + ' ' + \
                                     physiotherapy['name_physiotherapy'] + ' ' + \
                                     physiotherapy['value_physiotherapy'] + ' ' + \
                                     physiotherapy['unit_physiotherapy'] + ' ' + \
                                     physiotherapy['date_physiotherapy'] + ' ' + \
                                     physiotherapy['id_med_staff']
                sheet[physiotherapy_index][10].value = physiotherapy_str
                physiotherapy_index += 1

            for electro_ultrasound_therapy in stage_of_treatment['electro_ultrasound_therapy']:  # test
                pass

            state_index = row
            for state in stage_of_treatment['state']:
                print(state['state_on'])
                sheet[state_index][12].value = state['state_on']

                laboratory_test_str = ''
                laboratory_test_index = state_index
                for laboratory_test in state['laboratory_test']:
                    laboratory_test_str = str(laboratory_test['nameTest']) + ' ' + str(laboratory_test['valueTest']) + ' ' + \
                                          str(laboratory_test['unitTest']) + ' ' + str(laboratory_test['dateTest']) + ' ' + \
                                          str(laboratory_test['laboratoryName'])
                    for lab_staff in laboratory_test['lab_staff']:  # test
                        laboratory_test_str += lab_staff['id_lab_test'] + ' ' + lab_staff['id_med_staff']

                    sheet[laboratory_test_index][13].value = laboratory_test_str
                    laboratory_test_index += 1

                measurement_str = ''
                measurement_index = state_index
                for measurement in state['measurement']:
                    measurement_str = str(measurement['nameMeasurement']) + ' ' + str(measurement['valueMeasurement']) + ' ' + \
                                      str(measurement['unitMeasurement']) + ' ' + str(measurement['dateMeasurement']) + ' ' + \
                                      str(measurement['medStaff'])
                    sheet[measurement_index][14].value = measurement_str
                    measurement_index += 1

                for image in state['image']:
                    pass

                if laboratory_test_index > measurement_index:
                    state_index = laboratory_test_index
                else:
                    state_index = measurement_index
                state_index += 1

            sheet[row][15].value = stage_of_treatment['endStage']
        sheet[row][16].value = treatment_session['endSession']
    arr = [comorbidity_index, surgery_index, pharmacotherapy_index, physiotherapy_index, state_index]
    max = 0
    for el in arr:
        if el > max:
            max = el
    row = max

book.save('test.xlsx')
book.close()
