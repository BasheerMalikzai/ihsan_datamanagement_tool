import pandas as pd
import openpyxl as xl
import time
import win32com.client
import os
import shutil
import numpy as np
import gspread
from oauth2client.service_account import ServiceAccountCredentials


def view_full_dataframe():

    pd.set_option('display.max_rows', 5000000)
    pd.set_option('display.max_columns', 5000000)
    pd.set_option('display.width', 50000000)


view_full_dataframe()
dictionary_df = {}


def update_communities_list(tbl_communities, tbl_target):
    """ Takes triggered communities from triggered communities list combines it with the communities in the
        choices list removes the duplicated communities to get only newly triggered communities and appends the
         newly added communities to the choices list, tbl_communities is the triggered communities table
         tbl_target is the xlsform where the choices list is to be updated with new communities"""

    communities_df = pd.read_excel(tbl_communities)
    communities = pd.read_excel(r'C:\FHIDatabase\xlsforms\CLTS\communities.xlsx')
    target_df = pd.read_excel(tbl_target, sheet_name='choices')
    wb = xl.load_workbook(tbl_target)
    ws = wb['choices']
    view_full_dataframe()
    communities_df = communities_df.rename(columns={'community_name': 'name', 'district': 'dst'})
    communities_df = communities_df[['name', 'dst']]
    communities_list = communities_df.values.tolist()
    rows = []
    column_list_name = []
    column_name = []
    column_label_english = []
    column_label_dari = []
    column_label_pashto = []
    column_ipn = []
    column_prv = []
    column_dst = []
    for record in communities_list:
        rows = ['community_name', record[0], record[0],record[0],record[0],'', '', record[1]]
        column_list_name.append(rows[0])
        column_name.append(rows[1])
        column_label_english.append(rows[2])
        column_label_dari.append(rows[3])
        column_label_pashto.append(rows[4])
        column_ipn.append(rows[5])
        column_prv.append(rows[6])
        column_dst.append(rows[7])
        dictionary_df = {'list name': column_list_name,
                         'name': column_name,
                         'label::English': column_label_english,
                         'label::Dari': column_label_dari,
                         'label::Pashto': column_label_pashto,
                         'ipn': column_ipn,
                         'prv': column_prv,
                         'dst': column_dst
                         }
    df = pd.DataFrame.from_dict(dictionary_df)
    target_df = target_df.append(df)
    target_df = target_df.append(communities)
    target_df = target_df.sort_values(by=['list name', 'name'])
    target_df = target_df[target_df['list name'] == 'community_name']
    target_df['concat'] = target_df['name'] + target_df['dst']
    target_df = target_df.drop_duplicates(subset='concat', keep=False)
    target_df = target_df.sort_values(by=['list name', 'name'])
    target_df = target_df.drop(columns=['concat'])
    target_df = target_df[['list name','name', 'label::English','label::Dari','label::Pashto','ipn','prv','dst']]
    data_list = target_df.values.tolist()
    for row in data_list:
        ws.append(row)
    wb.save(tbl_target)


def append_from_xls():

    communities = r'C:\FHIDatabase\xlsforms\communities_tobe_appended.xlsx'
    target_latrines_before_triggering = r'C:\FHIDatabase\xlsforms\CLTS\xlsform_primarylatrines_report.xlsx'
    target_certification = r'C:\FHIDatabase\xlsforms\CLTS\xlsForms_cltscertificationform.xlsx'
    target_claim = r'C:\FHIDatabase\xlsforms\CLTS\xlsForms_cltsClaiming.xlsx'
    target_clts_committee_estb = r'C:\FHIDatabase\xlsforms\CLTS\xlsForms_cltscommitteeestablishment.xlsx'
    target_sustainability = r'C:\FHIDatabase\xlsforms\CLTS\xlsForms_cltscommunitysustainablityplan.xlsx'
    target_demographics = r'C:\FHIDatabase\xlsforms\CLTS\xlsForms_cltsdemographics.xlsx'
    target_fhag_committee_estb = r'C:\FHIDatabase\xlsforms\CLTS\xlsForms_cltsfgagommitteeestablishment.xlsx'
    target_fhag_training = r'C:\FHIDatabase\xlsforms\CLTS\xlsForms_cltsfhagtraining.xlsx'
    target_households_visited_fhag = r'C:\FHIDatabase\xlsforms\CLTS\xlsForms_cltshhvisitedbyfhag.xlsx'
    target_final_latrines_report = r'C:\FHIDatabase\xlsforms\CLTS\xlsForms_cltslatrinereportsafterodfverfication.xlsx'
    target_meeting = r'C:\FHIDatabase\xlsforms\CLTS\xlsForms_cltsmeetingrecordform.xlsx'
    target_dysfunctional_water_source = r'C:\FHIDatabase\xlsforms\CLTS\xlsForms_cltsnumberofdysfunctionalwatersources.xlsx'
    target_dysfunctional_water_source_functional = r'C:\FHIDatabase\xlsforms\CLTS\xlsForms_cltsNumberofDysfunctionalwatersourcesmadefunctional.xlsx'
    target_post_odf_visits = r'C:\FHIDatabase\xlsforms\CLTS\xlsForms_cltspostodffollowup.xlsx'
    target_clts_training = r'C:\FHIDatabase\xlsforms\CLTS\xlsForms_cltstraining.xlsx'
    target_triggering_school = r'C:\FHIDatabase\xlsforms\CLTS\xlsForms_cltstriggeringinschool.xlsx'
    target_verification = r'C:\FHIDatabase\xlsforms\CLTS\xlsForms_cltsverification.xlsx'
    target_visits = r'C:\FHIDatabase\xlsforms\CLTS\xlsForms_cltsvisits.xlsx'
    target_baseline = r'C:\FHIDatabase\xlsforms\CLTS\Baseline_Form.xlsx'
    target_endline = r'C:\FHIDatabase\xlsforms\CLTS\Endline_Form.xlsx'
    start_time = time.time()
    update_communities_list(communities, target_latrines_before_triggering)
    update_communities_list(communities, target_certification)
    update_communities_list(communities, target_claim)
    update_communities_list(communities, target_clts_committee_estb)
    update_communities_list(communities, target_sustainability)
    update_communities_list(communities, target_demographics)
    update_communities_list(communities, target_fhag_committee_estb)
    update_communities_list(communities, target_fhag_training)
    update_communities_list(communities, target_households_visited_fhag)
    update_communities_list(communities, target_meeting)
    update_communities_list(communities, target_dysfunctional_water_source)
    update_communities_list(communities, target_dysfunctional_water_source_functional)
    update_communities_list(communities, target_post_odf_visits)
    update_communities_list(communities, target_clts_training)
    update_communities_list(communities, target_triggering_school)
    update_communities_list(communities, target_final_latrines_report)
    update_communities_list(communities, target_verification)
    update_communities_list(communities, target_visits)
    update_communities_list(communities, target_baseline)
    update_communities_list(communities, target_endline)
    end_time = time.time()
    print('All Communities appended to the choices list in ' + str(end_time-start_time) + ' seconds')


def update_all_communities():
    communities = r'C:\FHIDatabase\data\CLTS triggeringForm.xlsx'
    target_latrines_before_triggering = r'C:\FHIDatabase\xlsforms\CLTS\xlsform_primarylatrines_report.xlsx'
    target_certification = r'C:\FHIDatabase\xlsforms\CLTS\xlsForms_cltscertificationform.xlsx'
    target_claim = r'C:\FHIDatabase\xlsforms\CLTS\xlsForms_cltsClaiming.xlsx'
    target_clts_committee_estb = r'C:\FHIDatabase\xlsforms\CLTS\xlsForms_cltscommitteeestablishment.xlsx'
    target_sustainability = r'C:\FHIDatabase\xlsforms\CLTS\xlsForms_cltscommunitysustainablityplan.xlsx'
    target_demographics = r'C:\FHIDatabase\xlsforms\CLTS\xlsForms_cltsdemographics.xlsx'
    target_fhag_committee_estb = r'C:\FHIDatabase\xlsforms\CLTS\xlsForms_cltsfgagommitteeestablishment.xlsx'
    target_fhag_training = r'C:\FHIDatabase\xlsforms\CLTS\xlsForms_cltsfhagtraining.xlsx'
    target_households_visited_fhag = r'C:\FHIDatabase\xlsforms\CLTS\xlsForms_cltshhvisitedbyfhag.xlsx'
    target_final_latrines_report = r'C:\FHIDatabase\xlsforms\CLTS\xlsForms_cltslatrinereportsafterodfverfication.xlsx'
    target_meeting = r'C:\FHIDatabase\xlsforms\CLTS\xlsForms_cltsmeetingrecordform.xlsx'
    target_dysfunctional_water_source = r'C:\FHIDatabase\xlsforms\CLTS\xlsForms_cltsnumberofdysfunctionalwatersources.xlsx'
    target_dysfunctional_water_source_functional = r'C:\FHIDatabase\xlsforms\CLTS\xlsForms_cltsNumberofDysfunctionalwatersourcesmadefunctional.xlsx'
    target_post_odf_visits = r'C:\FHIDatabase\xlsforms\CLTS\xlsForms_cltspostodffollowup.xlsx'
    target_clts_training = r'C:\FHIDatabase\xlsforms\CLTS\xlsForms_cltstraining.xlsx'
    target_triggering_school = r'C:\FHIDatabase\xlsforms\CLTS\xlsForms_cltstriggeringinschool.xlsx'
    target_verification = r'C:\FHIDatabase\xlsforms\CLTS\xlsForms_cltsverification.xlsx'
    target_visits = r'C:\FHIDatabase\xlsforms\CLTS\xlsForms_cltsvisits.xlsx'
    target_baseline = r'C:\FHIDatabase\xlsforms\CLTS\Baseline_Form.xlsx'
    target_endline = r'C:\FHIDatabase\xlsforms\CLTS\Endline_Form.xlsx'
    start_time = time.time()
    update_communities_list(communities, target_latrines_before_triggering)
    update_communities_list(communities, target_certification)
    update_communities_list(communities, target_claim)
    update_communities_list(communities, target_clts_committee_estb)
    update_communities_list(communities, target_sustainability)
    update_communities_list(communities, target_demographics)
    update_communities_list(communities, target_fhag_committee_estb)
    update_communities_list(communities, target_fhag_training)
    update_communities_list(communities, target_households_visited_fhag)
    update_communities_list(communities, target_meeting)
    update_communities_list(communities, target_dysfunctional_water_source)
    update_communities_list(communities, target_dysfunctional_water_source_functional)
    update_communities_list(communities, target_post_odf_visits)
    update_communities_list(communities, target_clts_training)
    update_communities_list(communities, target_triggering_school)
    update_communities_list(communities, target_final_latrines_report)
    update_communities_list(communities, target_verification)
    update_communities_list(communities, target_visits)
    update_communities_list(communities, target_baseline)
    update_communities_list(communities, target_endline)
    end_time = time.time()
    print('All Communities appended to the choices list in ' + str(end_time-start_time) + ' seconds')


# Address for the clts data tables

def get_data(wb, ws,order):

    scope = ["https://spreadsheets.google.com/feeds",'https://www.googleapis.com/auth/spreadsheets',"https://www.googleapis.com/auth/drive.file","https://www.googleapis.com/auth/drive"]
    credentials = ServiceAccountCredentials.from_json_keyfile_name('C:\FHIDatabase\IHSAN_RTDC.json', scope)
    client = gspread.authorize(credentials)
    sheet = client.open(wb).worksheet(ws)
    data = sheet.get_all_records()
    df = pd.DataFrame.from_dict(data)
    df = df[order]
    return df




#Raw CLTS tables
triggering_table = r'C:\FHIDatabase\data\CLTS triggeringForm.xlsx'
demographics_table = r'C:\FHIDatabase\data\CLTS Demographics form.xlsx'
clts_committee_establishment_table = r'C:\FHIDatabase\data\CLTS community establish.xlsx'
clts_training_table = r'C:\FHIDatabase\data\CLTS Training.xlsx'
clts_committee_held_meetings_table = r'C:\FHIDatabase\data\CLTS Meetings .xlsx'
fhag_committee_establishment_table = r'C:\FHIDatabase\data\CLTS FHAG Committee Establishment form.xlsx'
fhag_training_table = r'C:\FHIDatabase\data\CLTS FHAG Training Form.xlsx'
fhag_hh_coverage_table = r'C:\FHIDatabase\data\CLTS FHAG Households coverage form.xlsx'
baseline_collected_table = r'C:\FHIDatabase\data\Baseline.xlsx'
latrines_before_triggering_table = r'C:\FHIDatabase\data\CLTS primary latrines report.xlsx'
triggering_in_school_table = r'C:\FHIDatabase\data\CLTS Triggering in School.xlsx'
claiming_odf_table = r'C:\FHIDatabase\data\CLTS Claiming Form.xlsx'
verified_odf_table = r'C:\FHIDatabase\data\CLTS_Verification.xlsx'
certified_odf_table = r'C:\FHIDatabase\data\CLTS Certification Form.xlsx'
endline_collected_table = r'C:\FHIDatabase\data\Endline.xlsx'
visits_table = r'C:\FHIDatabase\data\CLTS visits and latrines.xlsx'
final_latrines = r'C:\FHIDatabase\xlsforms\CLTS\xlsForms_cltslatrinereportsafterodfverfication.xlsx'
sustainability_plan_table = r'C:\FHIDatabase\data\CLTS Community sustainability Plan Development form.xlsx'
post_odf_visits_table = r'C:\FHIDatabase\data\CLTS Post ODF Follow up visit Forms.xlsx'
dysfunc_water_sources_table = r'C:\FHIDatabase\data\CLTS Number of dysfunctional water sources form.xlsx'
dysfunc_sources_functional_table = r'C:\FHIDatabase\data\CLTS Number of dysfunctional water sources made functional Record forms.xlsx'

# Clean CLTS Tabels
triggering_table_clean = r'C:\FHIDatabase\CLTS-Data\CLTS triggeringForm.xlsx'
demographics_table_clean = r'C:\FHIDatabase\CLTS-Data\CLTS Demographics form.xlsx'
clts_committee_establishment_table_clean = r'C:\FHIDatabase\CLTS-Data\CLTS community establish.xlsx'
clts_training_table_clean = r'C:\FHIDatabase\CLTS-Data\CLTS Training.xlsx'
clts_committee_held_meetings_table_clean = r'C:\FHIDatabase\CLTS-Data\CLTS Meetings .xlsx'
fhag_committee_establishment_table_clean= r'C:\FHIDatabase\CLTS-Data\CLTS FHAG Committee Establishment form.xlsx'
fhag_training_table_clean = r'C:\FHIDatabase\CLTS-Data\CLTS FHAG Training Form.xlsx'
fhag_hh_coverage_table_clean = r'C:\FHIDatabase\CLTS-Data\CLTS FHAG Households coverage form.xlsx'
baseline_collected_table_clean = r'C:\FHIDatabase\CLTS-Data\Baseline.xlsx'
latrines_before_triggering_table_clean = r'C:\FHIDatabase\CLTS-Data\CLTS primary latrines report.xlsx'
triggering_in_school_table_clean = r'C:\FHIDatabase\CLTS-Data\CLTS Triggering in School.xlsx'
claiming_odf_table_clean = r'C:\FHIDatabase\CLTS-Data\CLTS Claiming Form.xlsx'
verified_odf_table_clean = r'C:\FHIDatabase\CLTS-Data\CLTS_Verification.xlsx'
certified_odf_table_clean = r'C:\FHIDatabase\CLTS-Data\CLTS Certification Form.xlsx'
endline_collected_table_clean = r'C:\FHIDatabase\CLTS-Data\Endline.xlsx'
visits_table_clean = r'C:\FHIDatabase\CLTS-Data\CLTS visits and latrines.xlsx'
final_latrines_clean = r'C:\FHIDatabase\xlsforms\CLTS\xlsForms_cltslatrinereportsafterodfverfication.xlsx'
sustainability_plan_table_clean= r'C:\FHIDatabase\CLTS-Data\CLTS Community sustainability Plan Development form.xlsx'
post_odf_visits_table_clean = r'C:\FHIDatabase\CLTS-Data\CLTS Post ODF Follow up visit Forms.xlsx'
dysfunc_water_sources_table_clean = r'C:\FHIDatabase\CLTS-Data\CLTS Number of dysfunctional water sources form.xlsx'
dysfunc_sources_functional_table_clean = r'C:\FHIDatabase\CLTS-Data\CLTS Number of dysfunctional water sources made functional Record forms.xlsx'


def refresh(workbook):
    """ Refreshes one excel table that is linked to clts data in KoBoToolbox"""
    application = win32com.client.DispatchEx("Excel.application")
    workbook = application.Workbooks.open(workbook)
    workbook.RefreshAll()
    application.CalculateUntilAsyncQueriesDone()
    application.DisplayAlerts = False
    workbook.Save()
    application.Quit()


def refresh_all():

    """ Refreshes all excel workbooks that are linked to kobotoolbox"""
    start = time.time()
    refresh(triggering_table)
    refresh(demographics_table)
    refresh(clts_committee_establishment_table)
    refresh(clts_training_table)
    refresh(clts_committee_held_meetings_table)
    refresh(fhag_committee_establishment_table)
    refresh(fhag_training_table)
    refresh(fhag_hh_coverage_table)
    refresh(baseline_collected_table)
    refresh(latrines_before_triggering_table)
    refresh(triggering_table)
    refresh(claiming_odf_table)
    refresh(verified_odf_table)
    refresh(certified_odf_table)
    refresh(endline_collected_table)
    refresh(visits_table)
    refresh(sustainability_plan_table)
    refresh(post_odf_visits_table)
    refresh(dysfunc_water_sources_table)
    refresh(dysfunc_sources_functional_table)
    end = time.time()
    print(end-start)


# refresh_all()


def load_data():

    # r_1=['ipname','province','district','dstcode','cdc','village','community_name','lat','long','coordinates','_coordinates_latitude','_coordinates_longitude','_coordinates_altitude','_coordinates_precision','triggeringdate','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
    # triggering_table_AKF = get_data('CLTS-AKF','Triggering',r_1)
    # triggering_table_IRC = get_data('CLTS-IRC','Triggering',r_1)
    # triggering_table_CHA = get_data('CLTS-CHA','Triggering',r_1)
    # triggering_table_HADAAF = get_data('CLTS-HADAAF','Triggering',r_1)
    # triggering_table = triggering_table_AKF.append(triggering_table_CHA).append(triggering_table_HADAAF).append(triggering_table_IRC)
    # triggering_table = triggering_table[r_1]
    # # print(triggering_table)
    #
    # r_2=['ipname','province','district','dstcode','communituy_name','houseincommunity','nohouseholdincommunity','nomalepopincommuniyt','nofemalepopincommounity','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
    # demographics_table_AKF = get_data('CLTS-AKF','Demographics',r_2)
    # demographics_table_IRC = get_data('CLTS-IRC','Demographics',r_2)
    # demographics_table_CHA = get_data('CLTS-CHA','Demographics',r_2)
    # demographics_table_HADAAF = get_data('CLTS-HADAAF','Demographics',r_2)
    # demographics_table = demographics_table_AKF.append(demographics_table_CHA).append(demographics_table_HADAAF).append(demographics_table_IRC)
    # demographics_table = demographics_table[r_2]
    # # print(demographics_table)
    #
    # r_3=['ipname','province','district','dstcode','communituy_name','cltscomestab','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
    # clts_committee_establishment_table_AKF = get_data('CLTS-AKF','CLTS Comittee Establishment',r_3)
    # clts_committee_establishment_table_IRC = get_data('CLTS-IRC','CLTS Comittee Establishment',r_3)
    # clts_committee_establishment_table_CHA = get_data('CLTS-CHA','CLTS Comittee Establishment',r_3)
    # clts_committee_establishment_table_HADAAF = get_data('CLTS-HADAAF','CLTS Comittee Establishment',r_3)
    # clts_committee_establishment_table = clts_committee_establishment_table_AKF.append(clts_committee_establishment_table_CHA).append(clts_committee_establishment_table_HADAAF).append(clts_committee_establishment_table_IRC)
    # clts_committee_establishment_table = clts_committee_establishment_table[r_3]
    # # print(clts_committee_establishment_table)
    #
    # r_4=['ipname','province','district','dstcode','communituy_name','number_clts_trained','iecmaterailsused','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
    # clts_training_table_AKF = get_data('CLTS-AKF','CLTS Training',r_4)
    # clts_training_table_IRC = get_data('CLTS-IRC','CLTS Training',r_4)
    # clts_training_table_CHA = get_data('CLTS-CHA','CLTS Training',r_4)
    # clts_training_table_HADAAF = get_data('CLTS-HADAAF','CLTS Training',r_4)
    # clts_training_table = clts_training_table_AKF.append(clts_training_table_CHA).append(clts_training_table_HADAAF).append(clts_training_table_IRC)
    # clts_training_table = clts_training_table[r_4]
    # # print(clts_training_table)
    #
    # r_5=['ipname','province','district','dstcode','communituy_name','number_clts_meetings','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
    # clts_Meetings_Record_AKF = get_data('CLTS-AKF','CLTS Meetings Record',r_5)
    # clts_Meetings_Record_IRC = get_data('CLTS-IRC','CLTS Meetings Record',r_5)
    # clts_Meetings_Record_CHA = get_data('CLTS-CHA','CLTS Meetings Record',r_5)
    # clts_Meetings_Record_HADAAF = get_data('CLTS-HADAAF','CLTS Meetings Record',r_5)
    # clts_Meetings_Record = clts_Meetings_Record_AKF.append(clts_Meetings_Record_CHA).append(clts_Meetings_Record_HADAAF).append(clts_Meetings_Record_IRC)
    # clts_committee_held_meetings_table = clts_Meetings_Record[r_5]
    # # print(clts_committee_held_meetings_table)
    #
    # r_6=['ipname','province','district','dstcode','communituy_name','num_fhag','iecmaterailsused','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
    # fhag_training_AKF = get_data('CLTS-AKF','FHAG Training',r_6)
    # fhag_training_IRC = get_data('CLTS-IRC','FHAG Training',r_6)
    # fhag_training_CHA = get_data('CLTS-CHA','FHAG Training',r_6)
    # fhag_training_HADAAF = get_data('CLTS-HADAAF','FHAG Training',r_6)
    # fhag_training = fhag_training_AKF.append(fhag_training_CHA).append(fhag_training_CHA).append(fhag_training_HADAAF).append(fhag_training_IRC)
    # fhag_training_table = fhag_training[r_6]
    # # print(fhag_training_table)
    #
    # r_7=['ipname','province','district','dstcode','communituy_name','number_hhvisitedbyfhag','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
    # fhag_households_coverage_form_AKF = get_data('CLTS-AKF','FHAG Households Coverage Form',r_7)
    # fhag_households_coverage_form_IRC = get_data('CLTS-IRC','FHAG Households Coverage Form',r_7)
    # fhag_households_coverage_form_CHA = get_data('CLTS-CHA','FHAG Households Coverage Form',r_7)
    # fhag_households_coverage_form_HADAAF = get_data('CLTS-HADAAF','FHAG Households Coverage Form',r_7)
    # fhag_households_coverage_form = fhag_households_coverage_form_AKF.append(fhag_households_coverage_form_CHA).append(fhag_households_coverage_form_HADAAF).append(fhag_households_coverage_form_IRC)
    # fhag_hh_coverage_table = fhag_households_coverage_form[r_7]
    # # print(fhag_hh_coverage_table)
    #
    # r_8 = ['ipname','province','district','dstcode','community_name','household_head_name','household_head_fname','household_head_sex','household_head_contactno','no_male_memebers','no_female_memebers','latrine_exists','primaryimgae','latrine_type','latrine_type/flush','latrine_type/pit','latrine_type/ventilation_pipe_improved','latrine_type/composite','latrine_type/traditional','latrine_flush_type','pit_latrine_type','no_facility','latrine_status','latrine_facilities','latrine_facilities/door_curtain','latrine_facilities/vent_pipe','latrine_facilities/window_net','latrine_facilities/evacuation_door','latrine_facilities/cover_for_hole','latrine_facilities/hand_washing_facility','latrine_facilities/soap_ash','latrine_facilities/water','latrine_facilities/none','distance_watersource_latrine','latrine_improved','excreta_undersoil_6months','opendefecation_inyard','type_waste_managment','type_waste_managment/solide_waste_managment','type_waste_managment/wastewater_managment','type_waste_managment/no_waste_managment','yard_clean','water_source','water_source/spring_or_kariz_protected','water_source/spring_or_kariz_unprotected','water_source/stream_river_clean','water_source/stream_river_uncleaned','water_source/protected_well','water_source/unprotected_well','water_source/tap','water_source/other','other_treatment_method','distance_source_drinking_water','treatment_method','treatment_method/boile','treatment_method/bleach_chlorine','treatment_method/strain_through_cloth','treatment_method/water_filter','treatment_method/solar_dysinfiction','treatment_method/letit_sand_settle','treatment_method/no_tratment_method','treatment_method/other','other_treatment_method_001','dysfunctional_pump_availability','nohh_use_thispump','doyou_know_reason_dysfucntionality','reason-dysfunctionality','any_attempt_made_functionality','repair_level_required','is_repair_beneficial','date','remarks','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
    # baseline_AKF = get_data('CLTS-AKF','Baseline',r_8)
    # baseline_IRC = get_data('CLTS-IRC','Baseline',r_8)
    # baseline_CHA = get_data('CLTS-CHA','Baseline',r_8)
    # baseline_HADAAF = get_data('CLTS-HADAAF','Baseline',r_8)
    # baseline = baseline_AKF.append(baseline_CHA).append(baseline_HADAAF).append(baseline_IRC)
    # baseline_collected_table = baseline[r_8]
    # # print(baseline_collected_table)
    #
    # r_9=['ipname','province','district','dstcode','communituy_name','latrines_before_traiggering','existing_unimproved_Latrines','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
    # existing_latrines_report_AKF = get_data('CLTS-AKF','Existing Latrines Report',r_9)
    # existing_latrines_report_IRC = get_data('CLTS-IRC','Existing Latrines Report',r_9)
    # existing_latrines_report_CHA = get_data('CLTS-CHA','Existing Latrines Report',r_9)
    # existing_latrines_report_HADAAF = get_data('CLTS-HADAAF','Existing Latrines Report',r_9)
    # existing_latrines_report = existing_latrines_report_AKF.append(existing_latrines_report_CHA).append(existing_latrines_report_HADAAF).append(existing_latrines_report_IRC)
    # latrines_before_triggering_table = existing_latrines_report[r_9]
    # # print(latrines_before_triggering_table)
    #
    # r_10=['ipname','province','district','dstcode','communituy_name','schoolinthisvillage','hysanitseessinheld','nochilbenfrothissession','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
    # triggering_in_school_AKF = get_data('CLTS-AKF','Triggering in School',r_10)
    # triggering_in_school_IRC = get_data('CLTS-IRC','Triggering in School',r_10)
    # triggering_in_school_CHA = get_data('CLTS-CHA','Triggering in School',r_10)
    # triggering_in_school_HADAAF = get_data('CLTS-HADAAF','Triggering in School',r_10)
    # triggering_in_school = triggering_in_school_AKF.append(triggering_in_school_CHA).append(triggering_in_school_HADAAF).append(triggering_in_school_IRC)
    # triggering_in_school_table = triggering_in_school[r_10]
    # # print(triggering_in_school_table)
    #
    # r_11=['ipname','province','district','dstcode','communituy_name','claim_date','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
    # claiming_ODF_AKF = get_data('CLTS-AKF','Claiming ODF',r_11)
    # claiming_ODF_IRC = get_data('CLTS-IRC','Claiming ODF',r_11)
    # claiming_ODF_CHA = get_data('CLTS-CHA','Claiming ODF',r_11)
    # claiming_ODF_HADAAF = get_data('CLTS-HADAAF','Claiming ODF',r_11)
    # claiming_ODF = claiming_ODF_AKF.append(claiming_ODF_CHA).append(claiming_ODF_HADAAF).append(claiming_ODF_IRC)
    # claiming_odf_table = claiming_ODF[r_11]
    # # print(claiming_odf_table)
    #
    # r_12=['ipname','province','district','dstcode','communituy_name','verified_date','verifiedimage','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
    # verification_ODF_AKF = get_data('CLTS-AKF','Verification ODF',r_12)
    # verification_ODF_IRC = get_data('CLTS-IRC','Verification ODF',r_12)
    # verification_ODF_CHA = get_data('CLTS-CHA','Verification ODF',r_12)
    # verification_ODF_HADAAF = get_data('CLTS-HADAAF','Verification ODF',r_12)
    # verification_ODF = verification_ODF_AKF.append(verification_ODF_CHA).append(verification_ODF_HADAAF).append(verification_ODF_IRC)
    # verified_odf_table = verification_ODF[r_12]
    # # print(verified_odf_table)
    #
    # r_13=['ipname','province','district','dstcode','communituy_name','certification_date','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
    # certification_ODF_AKF = get_data('CLTS-AKF','Certification ODF',r_13)
    # certification_ODF_IRC = get_data('CLTS-IRC','Certification ODF',r_13)
    # certification_ODF_CHA = get_data('CLTS-CHA','Certification ODF',r_13)
    # certification_ODF_HADAAF = get_data('CLTS-HADAAF','Certification ODF',r_13)
    # certification_ODF = certification_ODF_AKF.append(certification_ODF_CHA).append(certification_ODF_HADAAF).append(certification_ODF_IRC)
    # certified_odf_table = certification_ODF[r_13]
    # # print(certification_ODF)
    #
    # r_14 = ['ipname','province','district','dstcode','community_name','household_head_name','household_head_fname','household_head_sex','household_head_contactno','no_male_memebers','no_female_memebers','latrine_exists','secondaryimage','latrine_type','latrine_type/flush','latrine_type/pit','latrine_type/ventilation_pipe_improved','latrine_type/composite','latrine_type/traditional','latrine_flush_type','pit_latrine_type','no_facility','latrine_status','latrine_facilities','latrine_facilities/door_curtain','latrine_facilities/vent_pipe','latrine_facilities/window_net','latrine_facilities/evacuation_door','latrine_facilities/cover_for_hole','latrine_facilities/hand_washing_facility','latrine_facilities/soap_ash','latrine_facilities/water','latrine_facilities/none','distance_watersource_latrine','latrine_improved','excreta_undersoil_6months','opendefecation_inyard','type_waste_managment','type_waste_managment/solide_waste_managment','type_waste_managment/wastewater_managment','type_waste_managment/no_waste_managment','yard_clean','water_source','water_source/spring_or_kariz_protected','water_source/spring_or_kariz_unprotected','water_source/stream_river_clean','water_source/stream_river_uncleaned','water_source/protected_well','water_source/unprotected_well','water_source/tap','water_source/other','other_treatment_method','distance_source_drinking_water','treatment_method','treatment_method/boile','treatment_method/bleach_chlorine','treatment_method/strain_through_cloth','treatment_method/water_filter','treatment_method/solar_dysinfiction','treatment_method/letit_sand_settle','treatment_method/no_treatment_method','treatment_method/other','other_treatment_method_001','dysfunctional_pump_availability','nohh_use_thispump','doyou_know_reason_dysfucntionality','reason-dysfunctionality','any_attempt_made_functionality','repair_level_required','is_repair_beneficial','date','remarks','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
    # endline_AKF = get_data('CLTS-AKF','Endline',r_14)
    # endline_IRC = get_data('CLTS-IRC','Endline',r_14)
    # endline_CHA = get_data('CLTS-CHA','Endline',r_14)
    # endline_HADAAF = get_data('CLTS-HADAAF','Endline',r_14)
    # endline = endline_AKF.append(endline_CHA).append(endline_HADAAF).append(endline_IRC)
    # endline_collected_table = endline[r_14]
    # # print(endline_collected_table)
    #
    # r_15=['ipname','province','district','dstcode','community_name','visitnumber','visit_date','no_latrines_newly_built','no_latrines_upgraded','no_male_benef','no_female_benef','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
    # Post_Triggering_Visits_AKF = get_data('CLTS-AKF','Post Triggering Visits',r_15)
    # Post_Triggering_Visits_IRC = get_data('CLTS-IRC','Post Triggering Visits',r_15)
    # Post_Triggering_Visits_CHA = get_data('CLTS-CHA','Post Triggering Visits',r_15)
    # Post_Triggering_Visits_HADAAF = get_data('CLTS-HADAAF','Post Triggering Visits',r_15)
    # Post_Triggering_Visits = Post_Triggering_Visits_AKF.append(Post_Triggering_Visits_CHA).append(Post_Triggering_Visits_HADAAF).append(Post_Triggering_Visits_IRC)
    # visits_table = Post_Triggering_Visits[r_15]
    # # print(visits_table)
    #
    #
    # # final_Latrines_Report_AKF = get_data('CLTS-AKF','Final Latrines Report',None)
    # # final_Latrines_Report_IRC = get_data('CLTS-IRC','Final Latrines Report',None)
    # # final_Latrines_Report_CHA = get_data('CLTS-CHA','Final Latrines Report',None)
    # # final_Latrines_Report_HADAAF = get_data('CLTS-HADAAF','Final Latrines Report',None)
    # # final_latrines = final_Latrines_Report_AKF.append(final_Latrines_Report_CHA).append(final_Latrines_Report_HADAAF).append(final_Latrines_Report_IRC)
    # # print(final_latrines)
    #
    # r_17=['ipname','province','district','dstcode','communituy_name','sustain_plan_developed','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
    # sustain_plan_Dev_AKF = get_data('CLTS-AKF','Sustain_plan_Dev',order=r_17)
    # sustain_plan_Dev_IRC = get_data('CLTS-IRC','Sustain_plan_Dev',order=r_17)
    # sustain_plan_Dev_CHA = get_data('CLTS-CHA','Sustain_plan_Dev',order=r_17)
    # sustain_plan_Dev_HADAAF = get_data('CLTS-HADAAF','Sustain_plan_Dev',order=r_17)
    # sustain_plan_Dev = sustain_plan_Dev_AKF.append(sustain_plan_Dev_CHA).append(sustain_plan_Dev_HADAAF).append(sustain_plan_Dev_IRC)
    # sustainability_plan_table = sustain_plan_Dev[r_17]
    # # print(sustainability_plan_table)
    #
    # r_18=['ipname','province','district','dstcode','community_name','podfvisitnumber','visit_date','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
    # post_ODF_Followup_visit_AKF = get_data('CLTS-AKF','Post ODF Followup visit',r_18)
    # post_ODF_Followup_visit_IRC = get_data('CLTS-IRC','Post ODF Followup visit',r_18)
    # post_ODF_Followup_visit_CHA = get_data('CLTS-CHA','Post ODF Followup visit',r_18)
    # post_ODF_Followup_visit_HADAAF = get_data('CLTS-HADAAF','Post ODF Followup visit',r_18)
    # post_ODF_Followup_visit = post_ODF_Followup_visit_AKF.append(post_ODF_Followup_visit_CHA).append(post_ODF_Followup_visit_HADAAF).append(post_ODF_Followup_visit_IRC)
    # post_odf_visits_table = post_ODF_Followup_visit[r_18]
    # # print(post_odf_visits_table)
    #
    # r_19=['ipname','province','district','dstcode','communituy_name','number_dysfwatersources','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
    # dysFunc_Water_Sources_AKF = get_data('CLTS-AKF','DysFunc Water Sources',r_19)
    # dysFunc_Water_Sources_IRC = get_data('CLTS-IRC','DysFunc Water Sources',r_19)
    # dysFunc_Water_Sources_CHA = get_data('CLTS-CHA','DysFunc Water Sources',r_19)
    # dysFunc_Water_Sources_HADAAF = get_data('CLTS-HADAAF','DysFunc Water Sources',r_19)
    # dysFunc_Water_Sources = dysFunc_Water_Sources_AKF.append(dysFunc_Water_Sources_CHA).append(dysFunc_Water_Sources_HADAAF).append(dysFunc_Water_Sources_IRC)
    # dysfunc_water_sources_table =dysFunc_Water_Sources[r_19]
    # # print(dysfunc_water_sources_table)
    #
    # r_20=['ipname','province','district','dstcode','communituy_name','number_dysfwatersourmf','benifwatersourmf','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
    # dysFunc_Made_Func_AKF = get_data('CLTS-AKF','DysFunc Made Func',r_20)
    # dysFunc_Made_Func_IRC = get_data('CLTS-IRC','DysFunc Made Func',r_20)
    # dysFunc_Made_Func_CHA = get_data('CLTS-CHA','DysFunc Made Func',r_20)
    # dysFunc_Made_Func_HADAAF = get_data('CLTS-HADAAF','DysFunc Made Func',r_20)
    # dysFunc_Made_Func = dysFunc_Made_Func_AKF.append(dysFunc_Made_Func_CHA).append(dysFunc_Made_Func_HADAAF).append(dysFunc_Made_Func_IRC)
    # dysfunc_sources_functional_table =dysFunc_Made_Func[r_20]
    # # print(dysfunc_sources_functional_table)
    #
    # r_21 = ['ipname','province','district','dstcode','communituy_name','fhagcomestab','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
    # fhag_committee_establishment_table_AKF = get_data('CLTS-AKF', 'FHAG Comitte Establishment', r_21)
    # fhag_committee_establishment_table_CHA = get_data('CLTS-CHA', 'FHAG Comitte Establishment', r_21)
    # fhag_committee_establishment_table_HADAAF = get_data('CLTS-HADAAF', 'FHAG Comitte Establishment', r_21)
    # fhag_committee_establishment_table_IRC = get_data('CLTS-IRC', 'FHAG Comitte Establishment', r_21)
    # fhag_committee_establishment_table = fhag_committee_establishment_table_AKF.append(fhag_committee_establishment_table_CHA).append(fhag_committee_establishment_table_HADAAF).append(fhag_committee_establishment_table_IRC)
    # fhag_committee_establishment_table = fhag_committee_establishment_table[r_21]
    # # print(fhag_committee_establishment_table)


    # Formats and combines all clts tables and exports them to a csv file.
    # triggering_table_clean = r'C:\FHIDatabase\CLTS-Data\Master Sheet.xlsx'
    # demographics_table_clean = r'C:\FHIDatabase\CLTS-Data\CLTS Demographics form.xlsx'
    # clts_committee_establishment_table_clean = r'C:\FHIDatabase\CLTS-Data\CLTS community establish.xlsx'
    # clts_training_table_clean = r'C:\FHIDatabase\CLTS-Data\CLTS Training.xlsx'
    # clts_committee_held_meetings_table_clean = r'C:\FHIDatabase\CLTS-Data\CLTS Meetings .xlsx'
    # fhag_committee_establishment_table_clean= r'C:\FHIDatabase\CLTS-Data\CLTS FHAG Committee Establishment form.xlsx'
    # fhag_training_table_clean = r'C:\FHIDatabase\CLTS-Data\CLTS FHAG Training Form.xlsx'
    # fhag_hh_coverage_table_clean = r'C:\FHIDatabase\CLTS-Data\CLTS FHAG Households coverage form.xlsx'
    # baseline_collected_table_clean = r'C:\FHIDatabase\CLTS-Data\Baseline.xlsx'
    # latrines_before_triggering_table_clean = r'C:\FHIDatabase\CLTS-Data\CLTS primary latrines report.xlsx'
    # triggering_in_school_table_clean = r'C:\FHIDatabase\CLTS-Data\CLTS Triggering in School.xlsx'


    df_triggering = pd.read_excel(triggering_table_clean)
    print(df_triggering.keys())
    df_triggering = df_triggering[['ipname', 'province', 'district', 'dstcode', 'cdc', 'village', 'community_name','triggeringdate', 'lat','long','_id']]
    df_triggering['concat'] = df_triggering['province'] + '-' + df_triggering['district'] + '-' + df_triggering['community_name']
    df_triggering = df_triggering.drop_duplicates(subset='concat', keep='first')
    # print(df_triggering)

    df_demographics = pd.read_excel(demographics_table_clean)
    df_demographics['concat'] = df_demographics['province'] + '-' + df_demographics['district'] + '-' + df_demographics['communituy_name']
    df_demographics = df_demographics[['concat','houseincommunity', 'nohouseholdincommunity', 'nomalepopincommuniyt', 'nofemalepopincommounity']]
    df_demographics = df_demographics.drop_duplicates(subset='concat', keep='first')
    df_1_2 = pd.merge(df_triggering, df_demographics, left_on='concat', right_on='concat', how='left')
    # print(df_1_2)

    df_clts_committee_establishment = pd.read_excel(clts_committee_establishment_table_clean)
    df_clts_committee_establishment['concat'] = df_clts_committee_establishment['province'] + '-' + df_clts_committee_establishment['district'] + '-' + df_clts_committee_establishment['communituy_name']
    df_clts_committee_establishment = df_clts_committee_establishment[['concat', 'cltscomestab']]
    df_clts_committee_establishment =df_clts_committee_establishment.drop_duplicates(subset='concat', keep='first')
    df_2_3 = pd.merge(df_1_2, df_clts_committee_establishment, left_on='concat', right_on='concat', how='left')
    # print(df_2_3)

    df_clts_training = pd.read_excel(clts_training_table_clean)
    df_clts_training['concat'] = df_clts_training['province'] + '-' + df_clts_training['district'] + '-' + df_clts_training['communituy_name']
    df_clts_training = df_clts_training[['concat', 'number_clts_trained', 'iecmaterailsused']]
    df_clts_training = df_clts_training.groupby('concat').sum()
    df_3_4 = pd.merge(df_2_3, df_clts_training, left_on='concat', right_on='concat', how='left')
    # print(df_3_4)

    df_meetings_held = pd.read_excel(clts_committee_held_meetings_table_clean)
    df_meetings_held['concat'] = df_meetings_held['province'] + '-' + df_meetings_held['district'] + '-' + df_meetings_held['communituy_name']
    df_meetings_held = df_meetings_held[['concat', 'number_clts_meetings']]
    df_meetings_held = df_meetings_held.groupby(['concat']).sum()
    df_4_5 = pd.merge(df_3_4, df_meetings_held, left_on='concat', right_on='concat', how='left')
    # print(df_4_5)

    df_fhag_committe_establishment = pd.read_excel(fhag_committee_establishment_table_clean)
    df_fhag_committe_establishment['concat'] = df_fhag_committe_establishment['province'] + '-' + df_fhag_committe_establishment['district'] + '-' + df_fhag_committe_establishment['communituy_name']
    df_fhag_committe_establishment = df_fhag_committe_establishment[['concat', 'fhagcomestab']]
    df_fhag_committe_establishment = df_fhag_committe_establishment.drop_duplicates(subset='concat', keep='first')
    df_5_6 = pd.merge(df_4_5,df_fhag_committe_establishment, right_on='concat', left_on='concat', how='left')
    # print(df_5_6)

    df_fhag_training = pd.read_excel(fhag_training_table_clean)
    df_fhag_training['concat'] = df_fhag_training['province'] + '-' + df_fhag_training['district'] + '-' + df_fhag_training['communituy_name']
    df_fhag_training = df_fhag_training[['concat', 'num_fhag', 'iecmaterailsused']]
    df_fhag_training = df_fhag_training.groupby('concat').sum()
    df_6_7 = pd.merge(df_5_6, df_fhag_training, left_on='concat', right_on='concat', how='left')
    # print(df_6_7)

    df_fhag_hh_coverage = pd.read_excel(fhag_hh_coverage_table_clean)
    df_fhag_hh_coverage['concat'] = df_fhag_hh_coverage['province'] + '-' + df_fhag_hh_coverage['district'] + '-' + df_fhag_hh_coverage['communituy_name']
    df_fhag_hh_coverage = df_fhag_hh_coverage[['concat', 'number_hhvisitedbyfhag']]
    df_fhag_hh_coverage = df_fhag_hh_coverage.drop_duplicates(subset='concat', keep='first')
    df_fhag_hh_coverage = df_fhag_hh_coverage.groupby('concat').sum()
    df_7_8 = pd.merge(df_6_7, df_fhag_hh_coverage, left_on='concat', right_on='concat', how='left')
    # print(df_7_8)

    df_baseline_collected = pd.read_excel(baseline_collected_table_clean)
    df_baseline_collected['concat'] = df_baseline_collected['province'] + '-' + df_baseline_collected['district'] + '-' + df_baseline_collected['community_name']
    df_baseline_collected = df_baseline_collected[['concat']]
    df_baseline_collected = df_baseline_collected.drop_duplicates(subset='concat', keep='first')
    df_baseline_collected['baselineCollected'] = 1
    df_8_9 = pd.merge(df_7_8, df_baseline_collected, left_on='concat', right_on='concat', how='left')
    # print(df_8_9)

    df_latrines_before_triggering = pd.read_excel(latrines_before_triggering_table_clean)
    df_latrines_before_triggering['concat'] = df_latrines_before_triggering['province'] + '-' + df_latrines_before_triggering['district'] + '-' + df_latrines_before_triggering['communituy_name']
    df_latrines_before_triggering = df_latrines_before_triggering[['concat', 'latrines_before_traiggering', 'existing_unimproved_Latrines']]
    df_latrines_before_triggering = df_latrines_before_triggering.groupby('concat').sum()
    df_9_10 = pd.merge(df_8_9, df_latrines_before_triggering, left_on='concat', right_on='concat', how='left')
    # print(df_9_10)

    df_visits = pd.read_excel(visits_table_clean)
    df_visits = df_visits.dropna(subset=['province'])
    print(df_visits)
    df_visits['concat'] = df_visits['province'] + '-' + df_visits['district'] + '-' + df_visits['community_name']
    df_visits['concat_visit'] = df_visits['province'] + '-' + df_visits['district'] + '-' + df_visits['community_name']+ df_visits['visitnumber']
    df_visits = df_visits.drop_duplicates(subset='concat_visit', keep='first')
    df_visits = df_visits[['concat', 'visitnumber', 'visit_date','no_latrines_newly_built', 'no_latrines_upgraded', 'no_male_benef', 'no_female_benef']]
    df_visits['visit_date']= df_visits['visit_date'].astype('datetime64[ns]')
    df_visits = df_visits.replace(np.nan, 'visit', regex=True)
    df_visits = df_visits.pivot(index='concat', columns='visitnumber', values=['visit_date','no_latrines_newly_built', 'no_latrines_upgraded', 'no_male_benef', 'no_female_benef'])
    df_10_11 = pd.merge(df_9_10,df_visits, left_on='concat', right_on='concat', how='left')
    # print(df_10_11)

    df_triggering_in_school = pd.read_excel(triggering_in_school_table_clean)
    df_triggering_in_school['concat'] = df_triggering_in_school['province'] + '-' + df_triggering_in_school['district'] + '-' + df_triggering_in_school['communituy_name']
    df_triggering_in_school = df_triggering_in_school[['concat','schoolinthisvillage', 'hysanitseessinheld', 'nochilbenfrothissession']]
    df_triggering_in_school = df_triggering_in_school.drop_duplicates(subset='concat', keep='first')
    df_11_12 = pd.merge(df_10_11, df_triggering_in_school, left_on='concat', right_on='concat', how='left')
    # print(df_11_12)
    # claiming_odf_table_clean = r'C:\FHIDatabase\CLTS-Data\CLTS Claiming Form.xlsx'
    # verified_odf_table_clean = r'C:\FHIDatabase\CLTS-Data\CLTS_Verification.xlsx'
    # certified_odf_table_clean = r'C:\FHIDatabase\CLTS-Data\CLTS Certification Form.xlsx'
    # endline_collected_table_clean = r'C:\FHIDatabase\CLTS-Data\Endline.xlsx'
    # visits_table_clean = r'C:\FHIDatabase\CLTS-Data\CLTS visits and latrines.xlsx'
    # final_latrines_clean = r'C:\FHIDatabase\xlsforms\CLTS\xlsForms_cltslatrinereportsafterodfverfication.xlsx'
    # sustainability_plan_table_clean= r'C:\FHIDatabase\CLTS-Data\CLTS Community sustainability Plan Development form.xlsx'
    # post_odf_visits_table_clean = r'C:\FHIDatabase\CLTS-Data\CLTS Post ODF Follow up visit Forms.xlsx'
    # dysfunc_water_sources_table_clean = r'C:\FHIDatabase\CLTS-Data\CLTS Number of dysfunctional water sources form.xlsx'
    # dysfunc_sources_functional_table_clean = r'C:\FHIDatabase\CLTS-Data\CLTS Number of dysfunctional water sources made functional Record forms.xlsx'
    df_claiming_odf = pd.read_excel(claiming_odf_table_clean)
    df_claiming_odf['concat'] = df_claiming_odf['province'] + '-' + df_claiming_odf['district'] + '-' + df_claiming_odf['communituy_name']
    df_claiming_odf = df_claiming_odf[['concat', 'claim_date']]
    df_claiming_odf = df_claiming_odf.drop_duplicates(subset='concat', keep='first')
    df_12_13 = pd.merge(df_11_12, df_claiming_odf, left_on='concat', right_on='concat', how= 'left')
    # print(df_12_13)

    df_verified_odf = pd.read_excel(verified_odf_table_clean)
    df_verified_odf['concat'] = df_verified_odf['province'] + '-' + df_verified_odf['district'] + '-' + df_verified_odf['communituy_name']
    df_verified_odf = df_verified_odf[['concat', 'verified_date']]
    df_verified_odf = df_verified_odf.drop_duplicates(subset='concat', keep='first')
    df_13_14 = pd.merge(df_12_13, df_verified_odf, left_on='concat', right_on='concat', how='left')
    # print(df_13_14)

    df_certification_odf = pd.read_excel(certified_odf_table_clean)
    df_certification_odf['concat'] = df_certification_odf['province'] + '-' + df_certification_odf['district'] + '-' + df_certification_odf['communituy_name']
    df_certification_odf = df_certification_odf[['concat', 'certification_date']]
    df_certification_odf = df_certification_odf.drop_duplicates(subset='concat', keep='first')
    df_14_15 = pd.merge(df_13_14, df_certification_odf, left_on='concat', right_on='concat', how= 'left')
    # print(df_14_15)

    df_endline_collected = pd.read_excel(endline_collected_table_clean)
    df_endline_collected['concat'] = df_endline_collected['province'] + '-' + df_endline_collected['district'] + '-' + df_endline_collected['community_name']
    df_endline_collected = df_endline_collected[['concat']]
    df_endline_collected = df_endline_collected.drop_duplicates(subset='concat', keep='first')
    df_endline_collected['EndlinelineCollected'] = 1
    df_15_16 = pd.merge(df_14_15, df_endline_collected, left_on='concat', right_on='concat', how='left')
    #print(df_15_16)

###
    # df_latrines_after_verification = pd.read_excel(visits_table)
    # df_latrines_after_verification['concat'] = df_latrines_after_verification['province'] + '-' + df_latrines_after_verification['district'] + '-' + df_latrines_after_verification['community_name']
    # df_latrines_after_verification = df_latrines_after_verification[['concat','no_latrines_newly_built', 'number_latrinesupgraded']]
    # df_latrines_after_verification = df_latrines_after_verification.groupby('concat').sum()
    # df_16_17 = pd.merge(df_15_16, df_latrines_after_verification, left_on='concat', right_on='concat', how='left')
    # # print(df_16_17)

    df_sustainability_plan = pd.read_excel(sustainability_plan_table_clean)
    df_sustainability_plan['concat'] = df_sustainability_plan['province'] + '-' + df_sustainability_plan['district'] + '-' + df_sustainability_plan['communituy_name']
    df_sustainability_plan = df_sustainability_plan[['concat', 'sustain_plan_developed']]
    df_sustainability_plan = df_sustainability_plan.drop_duplicates(subset='concat', keep='first')
    df_17_18 = pd.merge(df_15_16, df_sustainability_plan, right_on='concat', left_on='concat', how='left')
    # print(df_17_18)

    df_post_odf_visit = pd.read_excel(post_odf_visits_table_clean)
    df_post_odf_visit['concat'] = df_post_odf_visit['province'] + '-' + df_post_odf_visit['district'] + '-' + df_post_odf_visit['community_name']
    df_post_odf_visit['concat_visit'] = df_post_odf_visit['province'] + '-' + df_post_odf_visit['district'] + '-' + df_post_odf_visit['community_name'] + df_post_odf_visit['podfvisitnumber']
    df_post_odf_visit = df_post_odf_visit.drop_duplicates(subset='concat_visit', keep='first')
    df_post_odf_visit = df_post_odf_visit[['concat', 'podfvisitnumber', 'visit_date']]
    df_visits = df_visits.replace(np.nan, 'podfvisit', regex=True)
    df_post_odf_visit = df_post_odf_visit.pivot(index='concat', columns='podfvisitnumber', values='visit_date')
    df_18_19 = pd.merge(df_17_18,df_post_odf_visit, left_on='concat', right_on='concat', how='left')
    # print(df_18_19)

    df_dysfunc_water_sources = pd.read_excel(dysfunc_water_sources_table_clean)
    df_dysfunc_water_sources['concat'] = df_dysfunc_water_sources['province'] + '-' + df_dysfunc_water_sources['district'] + '-'+ df_dysfunc_water_sources['communituy_name']
    df_dysfunc_water_sources = df_dysfunc_water_sources[['concat', 'number_dysfwatersources']]
    df_dysfunc_water_sources = df_dysfunc_water_sources.groupby('concat').sum()
    df_19_20 = pd.merge(df_18_19,df_dysfunc_water_sources, left_on='concat', right_on='concat', how='left')
    # print(df_19_20)

    df_dysfunc_source_made_functional = pd.read_excel(dysfunc_sources_functional_table_clean)
    df_dysfunc_source_made_functional['concat'] = df_dysfunc_source_made_functional['province'] + '-' + df_dysfunc_source_made_functional['district'] + '-'+ df_dysfunc_source_made_functional['communituy_name']
    df_dysfunc_source_made_functional= df_dysfunc_source_made_functional[['concat', 'number_dysfwatersourmf', 'benifwatersourmf']]
    df_dysfunc_source_made_functional= df_dysfunc_source_made_functional.groupby('concat').sum()
    df_20_21 = pd.merge(df_19_20, df_dysfunc_source_made_functional, left_on='concat', right_on='concat', how='left')

    df_final = df_20_21
    df_final.set_index('ipname', inplace=True)
    df_final.to_excel(r'C:\FHIDatabase\CLTS-Data\final CLTS Data.xlsx')
    application = win32com.client.Dispatch("Excel.application")
    workbook = application.Workbooks.open(r'C:\FHIDatabase\CLTS-Data\final CLTS Data.xlsx')
    application.Visible = True


#
# view_full_dataframe()
# load_data()

def clts_mrrd_format():

    df_mrrd_report = pd.read_excel(r'C:\Users\mmalikzai\OneDrive - Family Health International\Project_Data\FY20-3rd-Quarter Data.xlsx')
    df_mrrd_report = df_mrrd_report[df_mrrd_report.columns.values].replace(np.nan,0)
    df_mrrd_report['sum_newly_built_latrines'] = df_mrrd_report["('no_latrines_newly_built', 'visit_1')"].astype('float64') + df_mrrd_report["('no_latrines_newly_built', 'visit_2')"].astype('float64') + df_mrrd_report["('no_latrines_newly_built', 'visit_3')"].astype('float64')  +df_mrrd_report["('no_latrines_newly_built', 'visit_4')"].astype('float64') +df_mrrd_report["('no_latrines_newly_built', 'visit_5')"].astype('float64') +df_mrrd_report["('no_latrines_newly_built', 'visit_6')"].astype('float64') +df_mrrd_report["('no_latrines_newly_built', 'visit_7')"].astype('float64') +df_mrrd_report["('no_latrines_newly_built', 'visit_8')"].astype('float64')
    df_mrrd_report['sum_upgraded_latrines'] = df_mrrd_report["('no_latrines_upgraded', 'visit_1')"].astype('float64')  + df_mrrd_report["('no_latrines_upgraded', 'visit_2')"].astype('float64')  + df_mrrd_report["('no_latrines_upgraded', 'visit_3')"].astype('float64')  +df_mrrd_report["('no_latrines_upgraded', 'visit_4')"].astype('float64') +df_mrrd_report["('no_latrines_upgraded', 'visit_5')"].astype('float64') +df_mrrd_report["('no_latrines_upgraded', 'visit_6')"].astype('float64') +df_mrrd_report["('no_latrines_upgraded', 'visit_7')"].astype('float64')+df_mrrd_report["('no_latrines_upgraded', 'visit_8')"].astype('float64')
    renamed_columns = {'ipname':'IPName','province':'Province','district':'District','dstcode':'DistCode','cdc':'cdc','village':'Village Name','community_name':'CommunityName','houseincommunity':'HouesInCommunity','nohouseholdincommunity':'HHinCommunity','nomalepopincommuniyt':'PopulationInCommunityMale','nofemalepopincommounity':'PopulationInCommunityFemale','latrines_before_traiggering':'LatrinesBeforeTreggering','existing_unimproved_Latrines':'ExistingUnimporvedLatrinebeforeTraiggering','triggeringdate':'DateCommunityTraiggered','sum_newly_built_latrines':'NewlyBuildwihHWF','sum_upgraded_latrines':'LatrineUpgradedtoImporvHWV','num_fhag':'FHAG_Trg_Hyg_memb','number_clts_trained':'CLTS_Memb_trg_Hyg',"('visit_date', 'visit_1')":'Visit1',"('visit_date', 'visit_2')":'VIsit2',"('visit_date', 'visit_3')":'Visit3',"('visit_date', 'visit_4')":'Visit4',"('visit_date', 'visit_5')":'Visit5',"('visit_date', 'visit_6')":'Visit6',"('visit_date', 'visit_7')":'Visit7',"('visit_date', 'visit_8')":'Visit8','verified_date':'DateCommunityVerifiedODF','certification_date':'DateCommunityCertifyODF'
                       }
    df_mrrd_report = df_mrrd_report[['ipname','province','district','dstcode','cdc','village','community_name','houseincommunity','nohouseholdincommunity','nomalepopincommuniyt','nofemalepopincommounity','latrines_before_traiggering','existing_unimproved_Latrines','triggeringdate','sum_newly_built_latrines','sum_upgraded_latrines','num_fhag','number_clts_trained',"('visit_date', 'visit_1')","('visit_date', 'visit_2')","('visit_date', 'visit_3')","('visit_date', 'visit_4')","('visit_date', 'visit_5')","('visit_date', 'visit_6')","('visit_date', 'visit_7')","('visit_date', 'visit_8')",'verified_date','certification_date'
                                     ]]
    df_mrrd_report = df_mrrd_report.rename(columns = renamed_columns)
    df_mrrd_report.set_index('IPName', inplace=True)
    df_mrrd_report.to_excel(r'C:\FHIDatabase\CLTS-Data\CLTS MRRD.xlsx')
    application = win32com.client.Dispatch("Excel.application")
    workbook = application.Workbooks.open(r'C:\FHIDatabase\CLTS-Data\CLTS MRRD.xlsx')
    application.Visible = True


def load_data_to_upload(from_excel, from_google, google_sheet_name, order, output, ipname):

    refresh(from_excel)
    scope = ["https://spreadsheets.google.com/feeds",'https://www.googleapis.com/auth/spreadsheets',"https://www.googleapis.com/auth/drive.file","https://www.googleapis.com/auth/drive"]
    credentials = ServiceAccountCredentials.from_json_keyfile_name('C:\FHIDatabase\IHSAN_RTDC.json', scope)
    client = gspread.authorize(credentials)
    sheet = client.open(from_google).worksheet(google_sheet_name)
    data = sheet.get_all_records()
    df_from_google = pd.DataFrame.from_dict(data)
    df_from_google = df_from_google[order]
    df_from_google = df_from_google[df_from_google['ipname'] == ipname]
    df_from_excel = pd.read_excel(from_excel)
    df_from_excel = df_from_excel[df_from_excel['ipname'] == ipname]
    df_from_excel = df_from_excel[order]
    df = df_from_excel.append(df_from_google)
    # df['ID+MetaID'] = df['_id']+df['']
    df = df.drop_duplicates(subset='meta/instanceID', keep=False)
    df = df[order]
    df.to_excel(output)
    application = win32com.client.Dispatch("Excel.application")
    workbook = application.Workbooks.open(output)
    application.Visible = True


def parse_clts_data(source, clean_data,order):

    refresh(source)
    df_source = pd.read_excel(source)
    df_source = df_source[order]
    df_clean_data = pd.read_excel(clean_data)
    df_clean_data = df_clean_data[order]
    df_data_to_parse = df_source.append(df_clean_data)
    df_data_to_parse = df_data_to_parse.drop_duplicates(subset='meta/instanceID', keep=False)
    df_final = df_clean_data.append(df_data_to_parse)
    df_final = df_final.drop_duplicates(subset='meta/instanceID', keep='first')
    df_final = df_final[order]
    df_final.set_index('ipname', inplace=True)
    df_final.to_excel(clean_data)


