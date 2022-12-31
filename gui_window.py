
import main
import tkinter as tk
import time
from tkinter import *
from tkinter.ttk import Combobox


# start = time.time()
# # main.refresh_all()
# # main.update_all_communities()
# # main.load_data()
# # end = time.time()
# print(end-start)

def refresh_communities_list():

    main.refresh(main.triggering_table)


def upload_IRC_cbnp():

    o = ['ipname','province','district','dstcode','village','hfname','hfcode','hftype','healthpost','healthpostcode','nohh_hvisit','reportdate','lat','long','scrn0_59scrnm','scrn0_59scrnf','pwrecievnuteduconsoulu19','pwrecievnuteduconsoulo19','nonpwrecivwashnuteduccoun','mam0_23','sam0_23','samchildrefered','samchildreferedreached','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
    main.load_data_to_upload(from_excel=r'C:\FHIDatabase\data\1.CBNP.xlsx', from_google='IRC', google_sheet_name='CBNP', order=o, output=r'C:\FHIDatabase\data\CBNP data to be uploade.xlsx', ipname='irc')


def upload_AKF_cbnp():

    o = ['ipname','province','district','dstcode','village','hfname','hfcode','hftype','healthpost','healthpostcode','nohh_hvisit','reportdate','lat','long','scrn0_59scrnm','scrn0_59scrnf','pwrecievnuteduconsoulu19','pwrecievnuteduconsoulo19','nonpwrecivwashnuteduccoun','mam0_23','sam0_23','samchildrefered','samchildreferedreached','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
    main.load_data_to_upload(from_excel=r'C:\FHIDatabase\data\1.CBNP.xlsx', from_google='AKF', google_sheet_name='CBNP', order=o, output=r'C:\FHIDatabase\data\CBNP data to be uploade.xlsx', ipname='akf')


def upload_CHA_cbnp():

    o = ['ipname','province','district','dstcode','village','hfname','hfcode','hftype','healthpost','healthpostcode','nohh_hvisit','reportdate','lat','long','scrn0_59scrnm','scrn0_59scrnf','pwrecievnuteduconsoulu19','pwrecievnuteduconsoulo19','nonpwrecivwashnuteduccoun','mam0_23','sam0_23','samchildrefered','samchildreferedreached','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
    main.load_data_to_upload(from_excel=r'C:\FHIDatabase\data\1.CBNP.xlsx', from_google='CHA', google_sheet_name='CBNP', order=o, output=r'C:\FHIDatabase\data\CBNP data to be uploade.xlsx', ipname='cha')


def upload_HADAAF_cbnp():

    o = ['ipname','province','district','dstcode','village','hfname','hfcode','hftype','healthpost','healthpostcode','nohh_hvisit','reportdate','lat','long','scrn0_59scrnm','scrn0_59scrnf','pwrecievnuteduconsoulu19','pwrecievnuteduconsoulo19','nonpwrecivwashnuteduccoun','mam0_23','sam0_23','samchildrefered','samchildreferedreached','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
    main.load_data_to_upload(from_excel=r'C:\FHIDatabase\data\1.CBNP.xlsx', from_google='HADAAF', google_sheet_name='CBNP', order=o, output=r'C:\FHIDatabase\data\CBNP data to be uploade.xlsx', ipname='hadaaf')


def upload_IRC_cbnp_chs():

    o = ['ipname','province','district','dstcode','hfname','hfcode','hftype','nohealthpostsreported','nochildreg','nochweigh','nochildnewadd','nochildrengainedweight','nohmevisitswithsam','foddemsessions','nosamreferred','nochildscreened','nochildwithsam6to59','nochildwithsam0to59sam_mam_refered','nochildwithsam0to59sam_mam_refered_reached','noantinatalvisit','nopostnatalvisit','reportdate','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
    main.load_data_to_upload(from_excel=r'C:\FHIDatabase\data\3.CBNPCHS.xlsx', from_google='IRC', google_sheet_name='CBNP-CHS', order=o, output=r'C:\FHIDatabase\data\CBNP data to be uploade.xlsx', ipname='irc')


def upload_AKF_cbnp_chs():

    o = ['ipname','province','district','dstcode','hfname','hfcode','hftype','nohealthpostsreported','nochildreg','nochweigh','nochildnewadd','nochildrengainedweight','nohmevisitswithsam','foddemsessions','nosamreferred','nochildscreened','nochildwithsam6to59','nochildwithsam0to59sam_mam_refered','nochildwithsam0to59sam_mam_refered_reached','noantinatalvisit','nopostnatalvisit','reportdate','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
    main.load_data_to_upload(from_excel=r'C:\FHIDatabase\data\3.CBNPCHS.xlsx', from_google='AKF', google_sheet_name='CBNP-CHS', order=o, output=r'C:\FHIDatabase\data\CBNP data to be uploade.xlsx', ipname='akf')


def upload_CHA_cbnp_chs():

    o = ['ipname','province','district','dstcode','hfname','hfcode','hftype','nohealthpostsreported','nochildreg','nochweigh','nochildnewadd','nochildrengainedweight','nohmevisitswithsam','foddemsessions','nosamreferred','nochildscreened','nochildwithsam6to59','nochildwithsam0to59sam_mam_refered','nochildwithsam0to59sam_mam_refered_reached','noantinatalvisit','nopostnatalvisit','reportdate','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
    main.load_data_to_upload(from_excel=r'C:\FHIDatabase\data\3.CBNPCHS.xlsx', from_google='CHA', google_sheet_name='CBNP-CHS', order=o, output=r'C:\FHIDatabase\data\CBNP data to be uploade.xlsx', ipname='cha')


def upload_HADAAF_cbnp_chs():

    o = ['ipname','province','district','dstcode','hfname','hfcode','hftype','nohealthpostsreported','nochildreg','nochweigh','nochildnewadd','nochildrengainedweight','nohmevisitswithsam','foddemsessions','nosamreferred','nochildscreened','nochildwithsam6to59','nochildwithsam0to59sam_mam_refered','nochildwithsam0to59sam_mam_refered_reached','noantinatalvisit','nopostnatalvisit','reportdate','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
    main.load_data_to_upload(from_excel=r'C:\FHIDatabase\data\3.CBNPCHS.xlsx', from_google='HADAAF', google_sheet_name='CBNP-CHS', order=o, output=r'C:\FHIDatabase\data\CBNP data to be uploade.xlsx', ipname='hadaaf')

def upload_IRC_Site_visit():

    o = ['ipname','province','district','dstcode','hf_site','actisited/Nutrition','actisited/clts','actisited/livelihood','actisited/operation','actisited/innovationfund','actisited/hf','actisited/compliance','hfname','hftype','vistype','vistype/program','vistype/monitoring','vistype/dqa','vistype/livelihood','vistype/operation','vistype/innovationfund','vistype/compliance','vistype/hf','vistype/tac','jvisit','sectore','sectore/moph','sectore/mrrd','sectore/mowa','sectore/mohra','sectore/fhi360','sectore/afsen','sectore/mail','sectore/moe','sectore/tpm','sectore/moec','sectore/mgtwell','sectore/nepa','visitsdate','enddate','findings','followup','lat_long','_lat_long_latitude','_lat_long_longitude','_lat_long_altitude','_lat_long_precision','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
    main.load_data_to_upload(from_excel=r'C:\FHIDatabase\data\3.Sitevisite.xlsx', from_google='IRC', google_sheet_name='SiteVisit', order=o, output=r'C:\FHIDatabase\data\CBNP data to be uploade.xlsx', ipname='irc')


def upload_HADAAF_Site_visit():

    o = ['ipname','province','district','dstcode','hf_site','actisited/Nutrition','actisited/clts','actisited/livelihood','actisited/operation','actisited/innovationfund','actisited/hf','actisited/compliance','hfname','hftype','vistype','vistype/program','vistype/monitoring','vistype/dqa','vistype/livelihood','vistype/operation','vistype/innovationfund','vistype/compliance','vistype/hf','vistype/tac','jvisit','sectore','sectore/moph','sectore/mrrd','sectore/mowa','sectore/mohra','sectore/fhi360','sectore/afsen','sectore/mail','sectore/moe','sectore/tpm','sectore/moec','sectore/mgtwell','sectore/nepa','visitsdate','enddate','findings','followup','lat_long','_lat_long_latitude','_lat_long_longitude','_lat_long_altitude','_lat_long_precision','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
    main.load_data_to_upload(from_excel=r'C:\FHIDatabase\data\3.Sitevisite.xlsx', from_google='HADAAF', google_sheet_name='SiteVisit', order=o, output=r'C:\FHIDatabase\data\CBNP data to be uploade.xlsx', ipname='hadaaf')

def upload_CHA_Site_visit():

    o = ['ipname','province','district','dstcode','hf_site','actisited/Nutrition','actisited/clts','actisited/livelihood','actisited/operation','actisited/innovationfund','actisited/hf','actisited/compliance','hfname','hftype','vistype','vistype/program','vistype/monitoring','vistype/dqa','vistype/livelihood','vistype/operation','vistype/innovationfund','vistype/compliance','vistype/hf','vistype/tac','jvisit','sectore','sectore/moph','sectore/mrrd','sectore/mowa','sectore/mohra','sectore/fhi360','sectore/afsen','sectore/mail','sectore/moe','sectore/tpm','sectore/moec','sectore/mgtwell','sectore/nepa','visitsdate','enddate','findings','followup','lat_long','_lat_long_latitude','_lat_long_longitude','_lat_long_altitude','_lat_long_precision','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
    main.load_data_to_upload(from_excel=r'C:\FHIDatabase\data\3.Sitevisite.xlsx', from_google='CHA', google_sheet_name='SiteVisit', order=o, output=r'C:\FHIDatabase\data\CBNP data to be uploade.xlsx', ipname='cha')


def upload_AKF_Site_visit():

    o = ['ipname','province','district','dstcode','hf_site','actisited/Nutrition','actisited/clts','actisited/livelihood','actisited/operation','actisited/innovationfund','actisited/hf','actisited/compliance','hfname','hftype','vistype','vistype/program','vistype/monitoring','vistype/dqa','vistype/livelihood','vistype/operation','vistype/innovationfund','vistype/compliance','vistype/hf','vistype/tac','jvisit','sectore','sectore/moph','sectore/mrrd','sectore/mowa','sectore/mohra','sectore/fhi360','sectore/afsen','sectore/mail','sectore/moe','sectore/tpm','sectore/moec','sectore/mgtwell','sectore/nepa','visitsdate','enddate','findings','followup','lat_long','_lat_long_latitude','_lat_long_longitude','_lat_long_altitude','_lat_long_precision','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
    main.load_data_to_upload(from_excel=r'C:\FHIDatabase\data\3.Sitevisite.xlsx', from_google='AKF', google_sheet_name='SiteVisit', order=o, output=r'C:\FHIDatabase\data\CBNP data to be uploade.xlsx', ipname='akf')


def upload_AKF_Supportive_Supervision():

    o = ['ipname','province','district','dstcode','hfname','hftype','lat','long','hfhpvidate','bphsimp','hfvi','ptidenti','orientation','orientation/iycf','orientation/imam','orientation/surveillance','orientation/referral','orientation/micronutrints','orientation/datamanagement','nostaffotiented','remark','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
    main.load_data_to_upload(from_excel=r'C:\FHIDatabase\data\4.SupprtiveSuperVision.xlsx', from_google='AKF', google_sheet_name='Supportive_Supervision', order=o, output=r'C:\FHIDatabase\data\CBNP data to be uploade.xlsx', ipname='akf')


def upload_HADAAF_Supportive_Supervision():

    o = ['ipname','province','district','dstcode','hfname','hftype','lat','long','hfhpvidate','bphsimp','hfvi','ptidenti','orientation','orientation/iycf','orientation/imam','orientation/surveillance','orientation/referral','orientation/micronutrints','orientation/datamanagement','nostaffotiented','remark','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
    main.load_data_to_upload(from_excel=r'C:\FHIDatabase\data\4.SupprtiveSuperVision.xlsx', from_google='HADAAF', google_sheet_name='Supportive_Supervision', order=o, output=r'C:\FHIDatabase\data\CBNP data to be uploade.xlsx', ipname='hadaaf')


def upload_CHA_Supportive_Supervision():

    o = ['ipname','province','district','dstcode','hfname','hftype','lat','long','hfhpvidate','bphsimp','hfvi','ptidenti','orientation','orientation/iycf','orientation/imam','orientation/surveillance','orientation/referral','orientation/micronutrints','orientation/datamanagement','nostaffotiented','remark','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
    main.load_data_to_upload(from_excel=r'C:\FHIDatabase\data\4.SupprtiveSuperVision.xlsx', from_google='CHA', google_sheet_name='Supportive_Supervision', order=o, output=r'C:\FHIDatabase\data\CBNP data to be uploade.xlsx', ipname='cha')


def upload_IRC_Supportive_Supervision():

    o = ['ipname','province','district','dstcode','hfname','hftype','lat','long','hfhpvidate','bphsimp','hfvi','ptidenti','orientation','orientation/iycf','orientation/imam','orientation/surveillance','orientation/referral','orientation/micronutrints','orientation/datamanagement','nostaffotiented','remark','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
    main.load_data_to_upload(from_excel=r'C:\FHIDatabase\data\4.SupprtiveSuperVision.xlsx', from_google='IRC', google_sheet_name='Supportive_Supervision', order=o, output=r'C:\FHIDatabase\data\CBNP data to be uploade.xlsx', ipname='irc')

def upload_CHA_SBCC_Report():

    o = ['ipname','province','district','dstcode','village','rdate','Intervention','Intervention/nutrition','Intervention/wash','Intervention/livelihood','iec_materials_type/billboards','iec_materials_type/banners','iec_materials_type/posters','iec_materials_type/leaflets','iec_materials_type/brochure','iec_materials_type/flip_chart','iec_materials_type/stickers','lat_long','_lat_long_latitude','_lat_long_longitude','_lat_long_altitude','_lat_long_precision','__version__','_version_','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
    main.load_data_to_upload(from_excel=r'C:\FHIDatabase\data\7.SBCC.xlsx', from_google='CHA', google_sheet_name='SBCC Report', order=o, output=r'C:\FHIDatabase\data\CBNP data to be uploade.xlsx', ipname='cha')

def upload_AKF_SBCC_Report():

    o = ['ipname','province','district','dstcode','village','rdate','Intervention','Intervention/nutrition','Intervention/wash','Intervention/livelihood','iec_materials_type/billboards','iec_materials_type/banners','iec_materials_type/posters','iec_materials_type/leaflets','iec_materials_type/brochure','iec_materials_type/flip_chart','iec_materials_type/stickers','lat_long','_lat_long_latitude','_lat_long_longitude','_lat_long_altitude','_lat_long_precision','__version__','_version_','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
    main.load_data_to_upload(from_excel=r'C:\FHIDatabase\data\7.SBCC.xlsx', from_google='AKF', google_sheet_name='SBCC Report', order=o, output=r'C:\FHIDatabase\data\CBNP data to be uploade.xlsx', ipname='akf')

def upload_HADAAF_SBCC_Report():

    o = ['ipname','province','district','dstcode','village','rdate','Intervention','Intervention/nutrition','Intervention/wash','Intervention/livelihood','iec_materials_type/billboards','iec_materials_type/banners','iec_materials_type/posters','iec_materials_type/leaflets','iec_materials_type/brochure','iec_materials_type/flip_chart','iec_materials_type/stickers','lat_long','_lat_long_latitude','_lat_long_longitude','_lat_long_altitude','_lat_long_precision','__version__','_version_','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
    main.load_data_to_upload(from_excel=r'C:\FHIDatabase\data\7.SBCC.xlsx', from_google='HADAAF', google_sheet_name='SBCC Report', order=o, output=r'C:\FHIDatabase\data\CBNP data to be uploade.xlsx', ipname='hadaaf')

def upload_IRC_SBCC_Report():

    o = ['ipname','province','district','dstcode','village','rdate','Intervention','Intervention/nutrition','Intervention/wash','Intervention/livelihood','iec_materials_type/billboards','iec_materials_type/banners','iec_materials_type/posters','iec_materials_type/leaflets','iec_materials_type/brochure','iec_materials_type/flip_chart','iec_materials_type/stickers','lat_long','_lat_long_latitude','_lat_long_longitude','_lat_long_altitude','_lat_long_precision','__version__','_version_','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
    main.load_data_to_upload(from_excel=r'C:\FHIDatabase\data\7.SBCC.xlsx', from_google='IRC', google_sheet_name='SBCC Report', order=o, output=r'C:\FHIDatabase\data\CBNP data to be uploade.xlsx', ipname='irc')

def upload_IRC_MobileCinema():

    o = ['ipname','province','district','dstcode','village','reportdate','school_community','schoolname','schooltype','session_covered','session_covered/session_covered_livelihood','session_covered/session_covered_nutrition','session_covered/session_covered_wash','lat','long','studentmale','studentfemale','studentfemalenonpregnant','teachermale','teacherfemalepregnant','teachernonpregnant','pwrnwleducu19','pwrnwleduco19','nonpwrnwleduc','__version__','_version_','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
    main.load_data_to_upload(from_excel=r'C:\FHIDatabase\data\6.MobileCinema.xlsx', from_google='IRC', google_sheet_name='MobileCinema', order=o, output=r'C:\FHIDatabase\data\CBNP data to be uploade.xlsx', ipname='irc')

def upload_AKF_MobileCinema():

    o = ['ipname','province','district','dstcode','village','reportdate','school_community','schoolname','schooltype','session_covered','session_covered/session_covered_livelihood','session_covered/session_covered_nutrition','session_covered/session_covered_wash','lat','long','studentmale','studentfemale','studentfemalenonpregnant','teachermale','teacherfemalepregnant','teachernonpregnant','pwrnwleducu19','pwrnwleduco19','nonpwrnwleduc','__version__','_version_','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
    main.load_data_to_upload(from_excel=r'C:\FHIDatabase\data\6.MobileCinema.xlsx', from_google='AKF', google_sheet_name='MobileCinema', order=o, output=r'C:\FHIDatabase\data\CBNP data to be uploade.xlsx', ipname='akf')

def upload_HADAAF_MobileCinema():

    o = ['ipname','province','district','dstcode','village','reportdate','school_community','schoolname','schooltype','session_covered','session_covered/session_covered_livelihood','session_covered/session_covered_nutrition','session_covered/session_covered_wash','lat','long','studentmale','studentfemale','studentfemalenonpregnant','teachermale','teacherfemalepregnant','teachernonpregnant','pwrnwleducu19','pwrnwleduco19','nonpwrnwleduc','__version__','_version_','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
    main.load_data_to_upload(from_excel=r'C:\FHIDatabase\data\6.MobileCinema.xlsx', from_google='HADAAF', google_sheet_name='MobileCinema', order=o, output=r'C:\FHIDatabase\data\CBNP data to be uploade.xlsx', ipname='hadaaf')


def upload_CHA_MobileCinema():

    o = ['ipname','province','district','dstcode','village','reportdate','school_community','schoolname','schooltype','session_covered','session_covered/session_covered_livelihood','session_covered/session_covered_nutrition','session_covered/session_covered_wash','lat','long','studentmale','studentfemale','studentfemalenonpregnant','teachermale','teacherfemalepregnant','teachernonpregnant','pwrnwleducu19','pwrnwleduco19','nonpwrnwleduc','__version__','_version_','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
    main.load_data_to_upload(from_excel=r'C:\FHIDatabase\data\6.MobileCinema.xlsx', from_google='CHA', google_sheet_name='MobileCinema', order=o, output=r'C:\FHIDatabase\data\CBNP data to be uploade.xlsx', ipname='cha')

def upload_IRC_Trainings():

    o = ['ipname','province','district','dstcode','village','lat','long','trgcategory','trgtype','trg_title','renewtrg','startdate','enddate','maletrained','fmaletrained','cdc','lelder','cltscommittee','fagh','farmer','chw','hshura','others','ips','moph','mrrd_prrd','mowa_dowa','mail_dail','moe_doe','mora','govtothers','dr','mw','nurse','consular','chs','journalist','religiousleader','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
    main.load_data_to_upload(from_excel=r'C:\FHIDatabase\data\5.Training.xlsx', from_google='IRC', google_sheet_name='Training', order=o, output=r'C:\FHIDatabase\data\CBNP data to be uploade.xlsx', ipname='irc')

def upload_AKF_Trainings():

    o = ['ipname','province','district','dstcode','village','lat','long','trgcategory','trgtype','trg_title','renewtrg','startdate','enddate','maletrained','fmaletrained','cdc','lelder','cltscommittee','fagh','farmer','chw','hshura','others','ips','moph','mrrd_prrd','mowa_dowa','mail_dail','moe_doe','mora','govtothers','dr','mw','nurse','consular','chs','journalist','religiousleader','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
    main.load_data_to_upload(from_excel=r'C:\FHIDatabase\data\5.Training.xlsx', from_google='AKF', google_sheet_name='Training', order=o, output=r'C:\FHIDatabase\data\CBNP data to be uploade.xlsx', ipname='akf')

def upload_CHA_Trainings():

    o = ['ipname','province','district','dstcode','village','lat','long','trgcategory','trgtype','trg_title','renewtrg','startdate','enddate','maletrained','fmaletrained','cdc','lelder','cltscommittee','fagh','farmer','chw','hshura','others','ips','moph','mrrd_prrd','mowa_dowa','mail_dail','moe_doe','mora','govtothers','dr','mw','nurse','consular','chs','journalist','religiousleader','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
    main.load_data_to_upload(from_excel=r'C:\FHIDatabase\data\5.Training.xlsx', from_google='CHA', google_sheet_name='Training', order=o, output=r'C:\FHIDatabase\data\CBNP data to be uploade.xlsx', ipname='cha')

def upload_HADAAF_Trainings():

    o = ['ipname','province','district','dstcode','village','lat','long','trgcategory','trgtype','trg_title','renewtrg','startdate','enddate','maletrained','fmaletrained','cdc','lelder','cltscommittee','fagh','farmer','chw','hshura','others','ips','moph','mrrd_prrd','mowa_dowa','mail_dail','moe_doe','mora','govtothers','dr','mw','nurse','consular','chs','journalist','religiousleader','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
    main.load_data_to_upload(from_excel=r'C:\FHIDatabase\data\5.Training.xlsx', from_google='HADAAF', google_sheet_name='Training', order=o, output=r'C:\FHIDatabase\data\CBNP data to be uploade.xlsx', ipname='hadaaf')


def upload_IRC_Baseline():

    o = ['ipname','province','district','dstcode','community_name','household_head_name','household_head_fname','household_head_sex','household_head_contactno','no_male_memebers','no_female_memebers','latrine_exists','primaryimgae','latrine_type','latrine_type/flush','latrine_type/pit','latrine_type/ventilation_pipe_improved','latrine_type/composite','latrine_type/traditional','latrine_flush_type','pit_latrine_type','no_facility','latrine_status','latrine_facilities','latrine_facilities/door_curtain','latrine_facilities/vent_pipe','latrine_facilities/window_net','latrine_facilities/evacuation_door','latrine_facilities/cover_for_hole','latrine_facilities/hand_washing_facility','latrine_facilities/soap_ash','latrine_facilities/water','latrine_facilities/none','distance_watersource_latrine','latrine_improved','excreta_undersoil_6months','opendefecation_inyard','type_waste_managment','type_waste_managment/solide_waste_managment','type_waste_managment/wastewater_managment','type_waste_managment/no_waste_managment','yard_clean','water_source','water_source/spring_or_kariz_protected','water_source/spring_or_kariz_unprotected','water_source/stream_river_clean','water_source/stream_river_uncleaned','water_source/protected_well','water_source/unprotected_well','water_source/tap','water_source/other','other_treatment_method','distance_source_drinking_water','treatment_method','treatment_method/boile','treatment_method/bleach_chlorine','treatment_method/strain_through_cloth','treatment_method/water_filter','treatment_method/solar_dysinfiction','treatment_method/letit_sand_settle','treatment_method/no_tratment_method','treatment_method/other','other_treatment_method_001','dysfunctional_pump_availability','nohh_use_thispump','doyou_know_reason_dysfucntionality','reason-dysfunctionality','any_attempt_made_functionality','repair_level_required','is_repair_beneficial','date','remarks','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
    main.load_data_to_upload(from_excel=r'C:\FHIDatabase\data\Baseline.xlsx', from_google='IRC', google_sheet_name='Baseline', order=o, output=r'C:\FHIDatabase\data\CBNP data to be uploade.xlsx', ipname='irc')


def upload_AKF_Baseline():

    o = ['ipname','province','district','dstcode','community_name','household_head_name','household_head_fname','household_head_sex','household_head_contactno','no_male_memebers','no_female_memebers','latrine_exists','primaryimgae','latrine_type','latrine_type/flush','latrine_type/pit','latrine_type/ventilation_pipe_improved','latrine_type/composite','latrine_type/traditional','latrine_flush_type','pit_latrine_type','no_facility','latrine_status','latrine_facilities','latrine_facilities/door_curtain','latrine_facilities/vent_pipe','latrine_facilities/window_net','latrine_facilities/evacuation_door','latrine_facilities/cover_for_hole','latrine_facilities/hand_washing_facility','latrine_facilities/soap_ash','latrine_facilities/water','latrine_facilities/none','distance_watersource_latrine','latrine_improved','excreta_undersoil_6months','opendefecation_inyard','type_waste_managment','type_waste_managment/solide_waste_managment','type_waste_managment/wastewater_managment','type_waste_managment/no_waste_managment','yard_clean','water_source','water_source/spring_or_kariz_protected','water_source/spring_or_kariz_unprotected','water_source/stream_river_clean','water_source/stream_river_uncleaned','water_source/protected_well','water_source/unprotected_well','water_source/tap','water_source/other','other_treatment_method','distance_source_drinking_water','treatment_method','treatment_method/boile','treatment_method/bleach_chlorine','treatment_method/strain_through_cloth','treatment_method/water_filter','treatment_method/solar_dysinfiction','treatment_method/letit_sand_settle','treatment_method/no_tratment_method','treatment_method/other','other_treatment_method_001','dysfunctional_pump_availability','nohh_use_thispump','doyou_know_reason_dysfucntionality','reason-dysfunctionality','any_attempt_made_functionality','repair_level_required','is_repair_beneficial','date','remarks','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
    main.load_data_to_upload(from_excel=r'C:\FHIDatabase\data\Baseline.xlsx', from_google='AKF', google_sheet_name='Baseline', order=o, output=r'C:\FHIDatabase\data\CBNP data to be uploade.xlsx', ipname='akf')

def upload_CHA_Baseline():

    o = ['ipname','province','district','dstcode','community_name','household_head_name','household_head_fname','household_head_sex','household_head_contactno','no_male_memebers','no_female_memebers','latrine_exists','primaryimgae','latrine_type','latrine_type/flush','latrine_type/pit','latrine_type/ventilation_pipe_improved','latrine_type/composite','latrine_type/traditional','latrine_flush_type','pit_latrine_type','no_facility','latrine_status','latrine_facilities','latrine_facilities/door_curtain','latrine_facilities/vent_pipe','latrine_facilities/window_net','latrine_facilities/evacuation_door','latrine_facilities/cover_for_hole','latrine_facilities/hand_washing_facility','latrine_facilities/soap_ash','latrine_facilities/water','latrine_facilities/none','distance_watersource_latrine','latrine_improved','excreta_undersoil_6months','opendefecation_inyard','type_waste_managment','type_waste_managment/solide_waste_managment','type_waste_managment/wastewater_managment','type_waste_managment/no_waste_managment','yard_clean','water_source','water_source/spring_or_kariz_protected','water_source/spring_or_kariz_unprotected','water_source/stream_river_clean','water_source/stream_river_uncleaned','water_source/protected_well','water_source/unprotected_well','water_source/tap','water_source/other','other_treatment_method','distance_source_drinking_water','treatment_method','treatment_method/boile','treatment_method/bleach_chlorine','treatment_method/strain_through_cloth','treatment_method/water_filter','treatment_method/solar_dysinfiction','treatment_method/letit_sand_settle','treatment_method/no_tratment_method','treatment_method/other','other_treatment_method_001','dysfunctional_pump_availability','nohh_use_thispump','doyou_know_reason_dysfucntionality','reason-dysfunctionality','any_attempt_made_functionality','repair_level_required','is_repair_beneficial','date','remarks','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
    main.load_data_to_upload(from_excel=r'C:\FHIDatabase\data\Baseline.xlsx', from_google='CHA', google_sheet_name='Baseline', order=o, output=r'C:\FHIDatabase\data\CBNP data to be uploade.xlsx', ipname='cha')

def upload_HADAAF_Baseline():

    o = ['ipname','province','district','dstcode','community_name','household_head_name','household_head_fname','household_head_sex','household_head_contactno','no_male_memebers','no_female_memebers','latrine_exists','primaryimgae','latrine_type','latrine_type/flush','latrine_type/pit','latrine_type/ventilation_pipe_improved','latrine_type/composite','latrine_type/traditional','latrine_flush_type','pit_latrine_type','no_facility','latrine_status','latrine_facilities','latrine_facilities/door_curtain','latrine_facilities/vent_pipe','latrine_facilities/window_net','latrine_facilities/evacuation_door','latrine_facilities/cover_for_hole','latrine_facilities/hand_washing_facility','latrine_facilities/soap_ash','latrine_facilities/water','latrine_facilities/none','distance_watersource_latrine','latrine_improved','excreta_undersoil_6months','opendefecation_inyard','type_waste_managment','type_waste_managment/solide_waste_managment','type_waste_managment/wastewater_managment','type_waste_managment/no_waste_managment','yard_clean','water_source','water_source/spring_or_kariz_protected','water_source/spring_or_kariz_unprotected','water_source/stream_river_clean','water_source/stream_river_uncleaned','water_source/protected_well','water_source/unprotected_well','water_source/tap','water_source/other','other_treatment_method','distance_source_drinking_water','treatment_method','treatment_method/boile','treatment_method/bleach_chlorine','treatment_method/strain_through_cloth','treatment_method/water_filter','treatment_method/solar_dysinfiction','treatment_method/letit_sand_settle','treatment_method/no_tratment_method','treatment_method/other','other_treatment_method_001','dysfunctional_pump_availability','nohh_use_thispump','doyou_know_reason_dysfucntionality','reason-dysfunctionality','any_attempt_made_functionality','repair_level_required','is_repair_beneficial','date','remarks','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
    main.load_data_to_upload(from_excel=r'C:\FHIDatabase\data\Baseline.xlsx', from_google='HADAAF', google_sheet_name='Baseline', order=o, output=r'C:\FHIDatabase\data\CBNP data to be uploade.xlsx', ipname='hadaaf')


def upload_IRC_Endline():
    
    o = ['ipname','province','district','dstcode','community_name','household_head_name','household_head_fname','household_head_sex','household_head_contactno','no_male_memebers','no_female_memebers','latrine_exists','secondaryimage','latrine_type','latrine_type/flush','latrine_type/pit','latrine_type/ventilation_pipe_improved','latrine_type/composite','latrine_type/traditional','latrine_flush_type','pit_latrine_type','no_facility','latrine_status','latrine_facilities','latrine_facilities/door_curtain','latrine_facilities/vent_pipe','latrine_facilities/window_net','latrine_facilities/evacuation_door','latrine_facilities/cover_for_hole','latrine_facilities/hand_washing_facility','latrine_facilities/soap_ash','latrine_facilities/water','latrine_facilities/none','distance_watersource_latrine','latrine_improved','excreta_undersoil_6months','opendefecation_inyard','type_waste_managment','type_waste_managment/solide_waste_managment','type_waste_managment/wastewater_managment','type_waste_managment/no_waste_managment','yard_clean','water_source','water_source/spring_or_kariz_protected','water_source/spring_or_kariz_unprotected','water_source/stream_river_clean','water_source/stream_river_uncleaned','water_source/protected_well','water_source/unprotected_well','water_source/tap','water_source/other','other_treatment_method','distance_source_drinking_water','treatment_method','treatment_method/boile','treatment_method/bleach_chlorine','treatment_method/strain_through_cloth','treatment_method/water_filter','treatment_method/solar_dysinfiction','treatment_method/letit_sand_settle','treatment_method/no_tratment_method','treatment_method/other','other_treatment_method_001','dysfunctional_pump_availability','nohh_use_thispump','doyou_know_reason_dysfucntionality','reason-dysfunctionality','any_attempt_made_functionality','repair_level_required','is_repair_beneficial','date','remarks','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
    main.load_data_to_upload(from_excel=r'C:\FHIDatabase\data\Endline.xlsx', from_google='IRC', google_sheet_name='Endline', order=o, output=r'C:\FHIDatabase\data\CBNP data to be uploade.xlsx', ipname='irc')

def upload_AKF_Endline():

    o = ['ipname','province','district','dstcode','community_name','household_head_name','household_head_fname','household_head_sex','household_head_contactno','no_male_memebers','no_female_memebers','latrine_exists','secondaryimage','latrine_type','latrine_type/flush','latrine_type/pit','latrine_type/ventilation_pipe_improved','latrine_type/composite','latrine_type/traditional','latrine_flush_type','pit_latrine_type','no_facility','latrine_status','latrine_facilities','latrine_facilities/door_curtain','latrine_facilities/vent_pipe','latrine_facilities/window_net','latrine_facilities/evacuation_door','latrine_facilities/cover_for_hole','latrine_facilities/hand_washing_facility','latrine_facilities/soap_ash','latrine_facilities/water','latrine_facilities/none','distance_watersource_latrine','latrine_improved','excreta_undersoil_6months','opendefecation_inyard','type_waste_managment','type_waste_managment/solide_waste_managment','type_waste_managment/wastewater_managment','type_waste_managment/no_waste_managment','yard_clean','water_source','water_source/spring_or_kariz_protected','water_source/spring_or_kariz_unprotected','water_source/stream_river_clean','water_source/stream_river_uncleaned','water_source/protected_well','water_source/unprotected_well','water_source/tap','water_source/other','other_treatment_method','distance_source_drinking_water','treatment_method','treatment_method/boile','treatment_method/bleach_chlorine','treatment_method/strain_through_cloth','treatment_method/water_filter','treatment_method/solar_dysinfiction','treatment_method/letit_sand_settle','treatment_method/no_tratment_method','treatment_method/other','other_treatment_method_001','dysfunctional_pump_availability','nohh_use_thispump','doyou_know_reason_dysfucntionality','reason-dysfunctionality','any_attempt_made_functionality','repair_level_required','is_repair_beneficial','date','remarks','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
    main.load_data_to_upload(from_excel=r'C:\FHIDatabase\data\Endline.xlsx', from_google='AKF', google_sheet_name='Endline', order=o, output=r'C:\FHIDatabase\data\CBNP data to be uploade.xlsx', ipname='akf')

def upload_CHA_Endline():

    o = ['ipname','province','district','dstcode','community_name','household_head_name','household_head_fname','household_head_sex','household_head_contactno','no_male_memebers','no_female_memebers','latrine_exists','secondaryimage','latrine_type','latrine_type/flush','latrine_type/pit','latrine_type/ventilation_pipe_improved','latrine_type/composite','latrine_type/traditional','latrine_flush_type','pit_latrine_type','no_facility','latrine_status','latrine_facilities','latrine_facilities/door_curtain','latrine_facilities/vent_pipe','latrine_facilities/window_net','latrine_facilities/evacuation_door','latrine_facilities/cover_for_hole','latrine_facilities/hand_washing_facility','latrine_facilities/soap_ash','latrine_facilities/water','latrine_facilities/none','distance_watersource_latrine','latrine_improved','excreta_undersoil_6months','opendefecation_inyard','type_waste_managment','type_waste_managment/solide_waste_managment','type_waste_managment/wastewater_managment','type_waste_managment/no_waste_managment','yard_clean','water_source','water_source/spring_or_kariz_protected','water_source/spring_or_kariz_unprotected','water_source/stream_river_clean','water_source/stream_river_uncleaned','water_source/protected_well','water_source/unprotected_well','water_source/tap','water_source/other','other_treatment_method','distance_source_drinking_water','treatment_method','treatment_method/boile','treatment_method/bleach_chlorine','treatment_method/strain_through_cloth','treatment_method/water_filter','treatment_method/solar_dysinfiction','treatment_method/letit_sand_settle','treatment_method/no_tratment_method','treatment_method/other','other_treatment_method_001','dysfunctional_pump_availability','nohh_use_thispump','doyou_know_reason_dysfucntionality','reason-dysfunctionality','any_attempt_made_functionality','repair_level_required','is_repair_beneficial','date','remarks','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
    main.load_data_to_upload(from_excel=r'C:\FHIDatabase\data\Endline.xlsx', from_google='IRC', google_sheet_name='Endline', order=o, output=r'C:\FHIDatabase\data\CBNP data to be uploade.xlsx', ipname='cha')

def upload_HADAAF_Endline():

    o = ['ipname','province','district','dstcode','community_name','household_head_name','household_head_fname','household_head_sex','household_head_contactno','no_male_memebers','no_female_memebers','latrine_exists','secondaryimage','latrine_type','latrine_type/flush','latrine_type/pit','latrine_type/ventilation_pipe_improved','latrine_type/composite','latrine_type/traditional','latrine_flush_type','pit_latrine_type','no_facility','latrine_status','latrine_facilities','latrine_facilities/door_curtain','latrine_facilities/vent_pipe','latrine_facilities/window_net','latrine_facilities/evacuation_door','latrine_facilities/cover_for_hole','latrine_facilities/hand_washing_facility','latrine_facilities/soap_ash','latrine_facilities/water','latrine_facilities/none','distance_watersource_latrine','latrine_improved','excreta_undersoil_6months','opendefecation_inyard','type_waste_managment','type_waste_managment/solide_waste_managment','type_waste_managment/wastewater_managment','type_waste_managment/no_waste_managment','yard_clean','water_source','water_source/spring_or_kariz_protected','water_source/spring_or_kariz_unprotected','water_source/stream_river_clean','water_source/stream_river_uncleaned','water_source/protected_well','water_source/unprotected_well','water_source/tap','water_source/other','other_treatment_method','distance_source_drinking_water','treatment_method','treatment_method/boile','treatment_method/bleach_chlorine','treatment_method/strain_through_cloth','treatment_method/water_filter','treatment_method/solar_dysinfiction','treatment_method/letit_sand_settle','treatment_method/no_tratment_method','treatment_method/other','other_treatment_method_001','dysfunctional_pump_availability','nohh_use_thispump','doyou_know_reason_dysfucntionality','reason-dysfunctionality','any_attempt_made_functionality','repair_level_required','is_repair_beneficial','date','remarks','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
    main.load_data_to_upload(from_excel=r'C:\FHIDatabase\data\Endline.xlsx', from_google='HADAAF', google_sheet_name='Endline', order=o, output=r'C:\FHIDatabase\data\CBNP data to be uploade.xlsx', ipname='hadaaf')

def upload_IRC_meeting():

    o = ['province','district','dstcode','village','ipname','lat_long','_lat_long_latitude','_lat_long_longitude','_lat_long_altitude','_lat_long_precision','title','mtype','sdate','edate','cdc','dda','dcc','pmc','localelders','mrrd_prrd','moph','mail_dail','moe_doe','mowa_dowa','mora','others','male','female','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes][province','district','dstcode','village','ipname','lat_long','_lat_long_latitude','_lat_long_longitude','_lat_long_altitude','_lat_long_precision','title','mtype','sdate','edate','cdc','dda','dcc','pmc','localelders','mrrd_prrd','moph','mail_dail','moe_doe','mowa_dowa','mora','others','male','female','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
    main.load_data_to_upload(from_excel=r'')
    
#AKF CLTS DATA UPLOADING FUNCTIONS

# def upload_AKF_Triggering():
#
#     o = ['ipname','province','district','dstcode','cdc','village','community_name','lat','long','coordinates','_coordinates_latitude','_coordinates_longitude','_coordinates_altitude','_coordinates_precision','triggeringdate','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
#     main.load_data_to_upload(from_excel=main.triggering_table, from_google = 'CLTS-AKF', google_sheet_name='Triggering', order=o, output=r'C:\FHIDatabase\data\CBNP data to be uploade.xlsx', ipname='akf')
#
# def upload_AKF_demographics():
#
#     o = ['ipname','province','district','dstcode','communituy_name','houseincommunity','nohouseholdincommunity','nomalepopincommuniyt','nofemalepopincommounity','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
#     main.load_data_to_upload(from_excel=main.demographics_table, from_google = 'CLTS-AKF', google_sheet_name='Demographics', order=o, output=r'C:\FHIDatabase\data\CBNP data to be uploade.xlsx', ipname='akf')
#
#
# def upload_AKF_Existing_latrines_report():
#
#     o = ['ipname','province','district','dstcode','communituy_name','latrines_before_traiggering','existing_unimproved_Latrines','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
#     main.load_data_to_upload(from_excel=main.latrines_before_triggering_table, from_google = 'CLTS-AKF', google_sheet_name='Existing Latrines Report', order=o, output=r'C:\FHIDatabase\data\CBNP data to be uploade.xlsx', ipname='akf')
#
# def upload_AKF_CLTS_Training():
#
#     o = ['ipname','province','district','dstcode','communituy_name','number_clts_trained','iecmaterailsused','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
#     main.load_data_to_upload(from_excel=main.clts_training_table, from_google = 'CLTS-AKF', google_sheet_name='CLTS Training', order=o, output=r'C:\FHIDatabase\data\CBNP data to be uploade.xlsx', ipname='akf')
#
# def upload_AKF_FHAG_Training():
#
#     o = ['ipname','province','district','dstcode','communituy_name','num_fhag','iecmaterailsused','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name', '_parent_index','_tags','_notes']
#     main.load_data_to_upload(from_excel=main.fhag_training_table, from_google = 'CLTS-AKF', google_sheet_name='FHAG Training', order=o, output=r'C:\FHIDatabase\data\CBNP data to be uploade.xlsx', ipname='akf')
#
#
# def upload_AKF_triggering_in_school():
#
#     o = ['ipname','province','district','dstcode','communituy_name','schoolinthisvillage','hysanitseessinheld','nochilbenfrothissession','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
#     main.load_data_to_upload(from_excel=main.triggering_in_school_table, from_google = 'CLTS-AKF', google_sheet_name='Triggering in School', order=o, output=r'C:\FHIDatabase\data\CBNP data to be uploade.xlsx', ipname='akf')
#
#
# def upload_AKF_FHAG_Households_coverage_form():
#
#     o = ['ipname','province','district','dstcode','communituy_name','number_hhvisitedbyfhag','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
#     main.load_data_to_upload(from_excel=main.fhag_hh_coverage_table, from_google = 'CLTS-AKF', google_sheet_name='FHAG Households Coverage Form', order=o, output=r'C:\FHIDatabase\data\CBNP data to be uploade.xlsx', ipname='akf')
#
# def upload_AKF_post_triggering_visits():
#
#     o = ['ipname','province','district','dstcode','community_name','visitnumber','visit_date','no_latrines_newly_built','no_latrines_upgraded','no_male_benef','no_female_benef','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags	_notes']
#     main.load_data_to_upload(from_excel=main.visits_table, from_google = 'CLTS-AKF', google_sheet_name='Post Triggering Visits', order=o, output=r'C:\FHIDatabase\data\CBNP data to be uploade.xlsx', ipname='akf')
#
# def upload_AKF_clts_meetings_record():
#
#     o = ['ipname','province','district','dstcode','communituy_name','number_clts_meetings','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
#     main.load_data_to_upload(from_excel=main.clts_committee_held_meetings_table, from_google = 'CLTS-AKF', google_sheet_name='CLTS Meetings Record', order=o, output=r'C:\FHIDatabase\data\CBNP data to be uploade.xlsx', ipname='akf')
#
# def upload_AKF_claiming():
#
#     o = ['ipname','province','district','dstcode','communituy_name','claim_date','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
#     main.load_data_to_upload(from_excel=main.claiming_odf_table, from_google = 'CLTS-AKF', google_sheet_name='Claiming ODF', order=o, output=r'C:\FHIDatabase\data\CBNP data to be uploade.xlsx', ipname='akf')
#
# def upload_AKF_verification():
#
#     o = ['ipname','province','district','dstcode','communituy_name','verified_date','verifiedimage','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
#     main.load_data_to_upload(from_excel=main.verified_odf_table, from_google = 'CLTS-AKF', google_sheet_name='Verification ODF', order=o, output=r'C:\FHIDatabase\data\CBNP data to be uploade.xlsx', ipname='akf')
#
#
# def upload_AKF_certification():
#
#     o = ['ipname','province','district','dstcode','communituy_name','certification_date','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
#     main.load_data_to_upload(from_excel=main.certified_odf_table, from_google = 'CLTS-AKF', google_sheet_name='Certification ODF', order=o, output=r'C:\FHIDatabase\data\CBNP data to be uploade.xlsx', ipname='akf')
#
# def upload_AKF_dysfunc_watersource():
#
#     o = ['ipname','province','district','dstcode','communituy_name','number_dysfwatersources','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
#     main.load_data_to_upload(from_excel=main.dysfunc_water_sources_table, from_google = 'CLTS-AKF', google_sheet_name='DysFunc Water Sources', order=o, output=r'C:\FHIDatabase\data\CBNP data to be uploade.xlsx', ipname='akf')
#
# def upload_AKF__dysfunc_watersource_madefunc():
#
#     o = ['ipname','province','district','dstcode','communituy_name','number_dysfwatersourmf','benifwatersourmf','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
#     main.load_data_to_upload(from_excel=main.dysfunc_sources_functional_table, from_google = 'CLTS-AKF', google_sheet_name='DysFunc Made Func', order=o, output=r'C:\FHIDatabase\data\CBNP data to be uploade.xlsx', ipname='akf')
#
# def upload_AKF_clts_comittee_establishment():
#
#     o = ['ipname','province','district','dstcode','communituy_name','cltscomestab','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
#     main.load_data_to_upload(from_excel=main.clts_committee_establishment_table, from_google = 'CLTS-AKF', google_sheet_name='CLTS Comittee Establishment', order=o, output=r'C:\FHIDatabase\data\CBNP data to be uploade.xlsx', ipname='akf')
#
#
# def upload_AKF_postodf_follow_ups():
#
#     o = ['ipname','province','district','dstcode','community_name','podfvisitnumber','visit_date','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
#     main.load_data_to_upload(from_excel=main.post_odf_visits_table, from_google = 'CLTS-AKF', google_sheet_name='Post ODF Followup visit', order=o, output=r'C:\FHIDatabase\data\CBNP data to be uploade.xlsx', ipname='akf')
#
#
# def upload_AKF_fhag_comittee_establishment():
#
#     o = ['ipname','province','district','dstcode','ommunituy_name','fhagcomestab','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
#     main.load_data_to_upload(from_excel=main.fhag_committee_establishment_table, from_google='CLTS-AKF', google_sheet_name='FHAG Comitte Establishment', order=o, output=r'C:\FHIDatabase\data\CBNP data to be uploade.xlsx', ipname='akf')
#
# def upload_AKF_Sustain_plan_Dev():
#
#     o = ['ipname','province','district','dstcode','communituy_name','sustain_plan_developed','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
#     main.load_data_to_upload(from_excel=main.fhag_committee_establishment_table, from_google='CLTS-AKF', google_sheet_name='FHAG Comitte Establishment', order=o, output=r'C:\FHIDatabase\data\CBNP data to be uploade.xlsx', ipname='akf')
#
# # IRC CLTS DATA UPLOADING FUNCTIONS
#
# def upload_IRC_Triggering():
#
#     o = ['ipname','province','district','dstcode','cdc','village','community_name','lat','long','coordinates','_coordinates_latitude','_coordinates_longitude','_coordinates_altitude','_coordinates_precision','triggeringdate','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
#     main.load_data_to_upload(from_excel=main.triggering_table, from_google='CLTS-IRC', google_sheet_name='Triggering', order=o, output=r'C:\FHIDatabase\data\CBNP data to be uploade.xlsx', ipname='irc')
#
# def upload_IRC_demographics():
#
#     o = ['ipname','province','district','dstcode','communituy_name','houseincommunity','nohouseholdincommunity','nomalepopincommuniyt','nofemalepopincommounity','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
#     main.load_data_to_upload(from_excel=main.demographics_table, from_google = 'CLTS-IRC', google_sheet_name='Demographics', order=o, output=r'C:\FHIDatabase\data\CBNP data to be uploade.xlsx', ipname='irc')
#
#
# def upload_IRC_Existing_latrines_report():
#
#     o = ['ipname','province','district','dstcode','communituy_name','latrines_before_traiggering','existing_unimproved_Latrines','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
#     main.load_data_to_upload(from_excel=main.latrines_before_triggering_table, from_google = 'CLTS-IRC', google_sheet_name='Existing Latrines Report', order=o, output=r'C:\FHIDatabase\data\CBNP data to be uploade.xlsx', ipname='irc')
#
# def upload_IRC_CLTS_Training():
#
#     o = ['ipname','province','district','dstcode','communituy_name','number_clts_trained','iecmaterailsused','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
#     main.load_data_to_upload(from_excel=main.clts_training_table, from_google = 'CLTS-IRC', google_sheet_name='CLTS Training', order=o, output=r'C:\FHIDatabase\data\CBNP data to be uploade.xlsx', ipname='irc')
#
# def upload_IRC_FHAG_Training():
#
#     o = ['ipname','province','district','dstcode','communituy_name','num_fhag','iecmaterailsused','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name', '_parent_index','_tags','_notes']
#     main.load_data_to_upload(from_excel=main.fhag_training_table, from_google = 'CLTS-IRC', google_sheet_name='FHAG Training', order=o, output=r'C:\FHIDatabase\data\CBNP data to be uploade.xlsx', ipname='irc')
#
#
# def upload_IRC_triggering_in_school():
#
#     o = ['ipname','province','district','dstcode','communituy_name','schoolinthisvillage','hysanitseessinheld','nochilbenfrothissession','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
#     main.load_data_to_upload(from_excel=main.triggering_in_school_table, from_google = 'CLTS-IRC', google_sheet_name='Triggering in School', order=o, output=r'C:\FHIDatabase\data\CBNP data to be uploade.xlsx', ipname='irc')
#
#
# def upload_IRC_FHAG_Households_coverage_form():
#
#     o = ['ipname','province','district','dstcode','communituy_name','number_hhvisitedbyfhag','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
#     main.load_data_to_upload(from_excel=main.fhag_hh_coverage_table, from_google = 'CLTS-IRC', google_sheet_name='FHAG Households Coverage Form', order=o, output=r'C:\FHIDatabase\data\CBNP data to be uploade.xlsx', ipname='irc')
#
# def upload_IRC_post_triggering_visits():
#
#     o = ['ipname','province','district','dstcode','community_name','visitnumber','visit_date','no_latrines_newly_built','no_latrines_upgraded','no_male_benef','no_female_benef','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags	_notes']
#     main.load_data_to_upload(from_excel=main.visits_table, from_google = 'CLTS-IRC', google_sheet_name='Post Triggering Visits', order=o, output=r'C:\FHIDatabase\data\CBNP data to be uploade.xlsx', ipname='irc')
#
# def upload_IRC_clts_meetings_record():
#
#     o = ['ipname','province','district','dstcode','communituy_name','number_clts_meetings','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
#     main.load_data_to_upload(from_excel=main.clts_committee_held_meetings_table, from_google = 'CLTS-IRC', google_sheet_name='CLTS Meetings Record', order=o, output=r'C:\FHIDatabase\data\CBNP data to be uploade.xlsx', ipname='irc')
#
# def upload_IRC_claiming():
#
#     o = ['ipname','province','district','dstcode','communituy_name','claim_date','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
#     main.load_data_to_upload(from_excel=main.claiming_odf_table, from_google = 'CLTS-IRC', google_sheet_name='Claiming ODF', order=o, output=r'C:\FHIDatabase\data\CBNP data to be uploade.xlsx', ipname='irc')
#
# def upload_IRC_verification():
#
#     o = ['ipname','province','district','dstcode','communituy_name','verified_date','verifiedimage','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
#     main.load_data_to_upload(from_excel=main.verified_odf_table, from_google = 'CLTS-IRC', google_sheet_name='Verification ODF', order=o, output=r'C:\FHIDatabase\data\CBNP data to be uploade.xlsx', ipname='irc')
#
#
# def upload_IRC_certification():
#
#     o = ['ipname','province','district','dstcode','communituy_name','certification_date','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
#     main.load_data_to_upload(from_excel=main.certified_odf_table, from_google = 'CLTS-IRC', google_sheet_name='Certification ODF', order=o, output=r'C:\FHIDatabase\data\CBNP data to be uploade.xlsx', ipname='irc')
#
# def upload_IRC_dysfunc_watersource():
#
#     o = ['ipname','province','district','dstcode','communituy_name','number_dysfwatersources','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
#     main.load_data_to_upload(from_excel=main.dysfunc_water_sources_table, from_google = 'CLTS-IRC', google_sheet_name='DysFunc Water Sources', order=o, output=r'C:\FHIDatabase\data\CBNP data to be uploade.xlsx', ipname='irc')
#
# def upload_IRC_dysfunc_watersource_madefunc():
#
#     o = ['ipname','province','district','dstcode','communituy_name','number_dysfwatersourmf','benifwatersourmf','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
#     main.load_data_to_upload(from_excel=main.dysfunc_sources_functional_table, from_google = 'CLTS-IRC', google_sheet_name='DysFunc Made Func', order=o, output=r'C:\FHIDatabase\data\CBNP data to be uploade.xlsx', ipname='irc')
#
# def upload_IRC_clts_comittee_establishment():
#
#     o = ['ipname','province','district','dstcode','communituy_name','cltscomestab','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
#     main.load_data_to_upload(from_excel=main.clts_committee_establishment_table, from_google = 'CLTS-IRC', google_sheet_name='CLTS Comittee Establishment', order=o, output=r'C:\FHIDatabase\data\CBNP data to be uploade.xlsx', ipname='irc')
#
#
# def upload_IRC_postodf_follow_ups():
#
#     o = ['ipname','province','district','dstcode','community_name','podfvisitnumber','visit_date','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
#     main.load_data_to_upload(from_excel=main.post_odf_visits_table, from_google = 'CLTS-IRC', google_sheet_name='Post ODF Followup visit', order=o, output=r'C:\FHIDatabase\data\CBNP data to be uploade.xlsx', ipname='irc')
#
#
# def upload_IRC_fhag_comittee_establishment():
#
#     o = ['ipname','province','district','dstcode','ommunituy_name','fhagcomestab','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
#     main.load_data_to_upload(from_excel=main.fhag_committee_establishment_table, from_google='CLTS-IRC', google_sheet_name='FHAG Comitte Establishment', order=o, output=r'C:\FHIDatabase\data\CBNP data to be uploade.xlsx', ipname='irc')
#
# def upload_IRC_Sustain_plan_Dev():
#
#     o = ['ipname','province','district','dstcode','communituy_name','sustain_plan_developed','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
#     main.load_data_to_upload(from_excel=main.fhag_committee_establishment_table, from_google='CLTS-IRC', google_sheet_name='FHAG Comitte Establishment', order=o, output=r'C:\FHIDatabase\data\CBNP data to be uploade.xlsx', ipname='irc')
#
#
#
# # CHA CLTS DATA UPLOADING FUNCTIONS
#
# def upload_CHA_Triggering():
#
#     o = ['ipname','province','district','dstcode','cdc','village','community_name','lat','long','coordinates','_coordinates_latitude','_coordinates_longitude','_coordinates_altitude','_coordinates_precision','triggeringdate','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
#     main.load_data_to_upload(from_excel=main.triggering_table, from_google='CLTS-CHA', google_sheet_name='Triggering', order=o, output=r'C:\FHIDatabase\data\CBNP data to be uploade.xlsx', ipname='cha')
#
# def upload_CHA_demographics():
#
#     o = ['ipname','province','district','dstcode','communituy_name','houseincommunity','nohouseholdincommunity','nomalepopincommuniyt','nofemalepopincommounity','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
#     main.load_data_to_upload(from_excel=main.demographics_table, from_google = 'CLTS-CHA', google_sheet_name='Demographics', order=o, output=r'C:\FHIDatabase\data\CBNP data to be uploade.xlsx', ipname='cha')
#
#
# def upload_CHA_Existing_latrines_report():
#
#     o = ['ipname','province','district','dstcode','communituy_name','latrines_before_traiggering','existing_unimproved_Latrines','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
#     main.load_data_to_upload(from_excel=main.latrines_before_triggering_table, from_google = 'CLTS-CHA', google_sheet_name='Existing Latrines Report', order=o, output=r'C:\FHIDatabase\data\CBNP data to be uploade.xlsx', ipname='cha')
#
# def upload_CHA_CLTS_Training():
#
#     o = ['ipname','province','district','dstcode','communituy_name','number_clts_trained','iecmaterailsused','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
#     main.load_data_to_upload(from_excel=main.clts_training_table, from_google = 'CLTS-CHA', google_sheet_name='CLTS Training', order=o, output=r'C:\FHIDatabase\data\CBNP data to be uploade.xlsx', ipname='cha')
#
# def upload_CHA_FHAG_Training():
#
#     o = ['ipname','province','district','dstcode','communituy_name','num_fhag','iecmaterailsused','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name', '_parent_index','_tags','_notes']
#     main.load_data_to_upload(from_excel=main.fhag_training_table, from_google = 'CLTS-CHA', google_sheet_name='FHAG Training', order=o, output=r'C:\FHIDatabase\data\CBNP data to be uploade.xlsx', ipname='cha')
#
#
# def upload_CHA_triggering_in_school():
#
#     o = ['ipname','province','district','dstcode','communituy_name','schoolinthisvillage','hysanitseessinheld','nochilbenfrothissession','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
#     main.load_data_to_upload(from_excel=main.triggering_in_school_table, from_google = 'CLTS-CHA', google_sheet_name='Triggering in School', order=o, output=r'C:\FHIDatabase\data\CBNP data to be uploade.xlsx', ipname='cha')
#
#
# def upload_CHA_FHAG_Households_coverage_form():
#
#     o = ['ipname','province','district','dstcode','communituy_name','number_hhvisitedbyfhag','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
#     main.load_data_to_upload(from_excel=main.fhag_hh_coverage_table, from_google = 'CLTS-CHA', google_sheet_name='FHAG Households Coverage Form', order=o, output=r'C:\FHIDatabase\data\CBNP data to be uploade.xlsx', ipname='cha')
#
# def upload_CHA_post_triggering_visits():
#
#     o = ['ipname','province','district','dstcode','community_name','visitnumber','visit_date','no_latrines_newly_built','no_latrines_upgraded','no_male_benef','no_female_benef','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags	_notes']
#     main.load_data_to_upload(from_excel=main.visits_table, from_google = 'CLTS-CHA', google_sheet_name='Post Triggering Visits', order=o, output=r'C:\FHIDatabase\data\CBNP data to be uploade.xlsx', ipname='cha')
#
# def upload_CHA_clts_meetings_record():
#
#     o = ['ipname','province','district','dstcode','communituy_name','number_clts_meetings','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
#     main.load_data_to_upload(from_excel=main.clts_committee_held_meetings_table, from_google = 'CLTS-CHA', google_sheet_name='CLTS Meetings Record', order=o, output=r'C:\FHIDatabase\data\CBNP data to be uploade.xlsx', ipname='cha')
#
# def upload_CHA_claiming():
#
#     o = ['ipname','province','district','dstcode','communituy_name','claim_date','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
#     main.load_data_to_upload(from_excel=main.claiming_odf_table, from_google = 'CLTS-CHA', google_sheet_name='Claiming ODF', order=o, output=r'C:\FHIDatabase\data\CBNP data to be uploade.xlsx', ipname='cha')
#
# def upload_CHA_verification():
#
#     o = ['ipname','province','district','dstcode','communituy_name','verified_date','verifiedimage','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
#     main.load_data_to_upload(from_excel=main.verified_odf_table, from_google = 'CLTS-CHA', google_sheet_name='Verification ODF', order=o, output=r'C:\FHIDatabase\data\CBNP data to be uploade.xlsx', ipname='cha')
#
#
# def upload_CHA_certification():
#
#     o = ['ipname','province','district','dstcode','communituy_name','certification_date','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
#     main.load_data_to_upload(from_excel=main.certified_odf_table, from_google = 'CLTS-CHA', google_sheet_name='Certification ODF', order=o, output=r'C:\FHIDatabase\data\CBNP data to be uploade.xlsx', ipname='cha')
#
# def upload_CHA_dysfunc_watersource():
#
#     o = ['ipname','province','district','dstcode','communituy_name','number_dysfwatersources','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
#     main.load_data_to_upload(from_excel=main.dysfunc_water_sources_table, from_google = 'CLTS-CHA', google_sheet_name='DysFunc Water Sources', order=o, output=r'C:\FHIDatabase\data\CBNP data to be uploade.xlsx', ipname='cha')
#
# def upload_CHA_dysfunc_watersource_madefunc():
#
#     o = ['ipname','province','district','dstcode','communituy_name','number_dysfwatersourmf','benifwatersourmf','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
#     main.load_data_to_upload(from_excel=main.dysfunc_sources_functional_table, from_google = 'CLTS-CHA', google_sheet_name='DysFunc Made Func', order=o, output=r'C:\FHIDatabase\data\CBNP data to be uploade.xlsx', ipname='cha')
#
# def upload_CHA_clts_comittee_establishment():
#
#     o = ['ipname','province','district','dstcode','communituy_name','cltscomestab','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
#     main.load_data_to_upload(from_excel=main.clts_committee_establishment_table, from_google = 'CLTS-CHA', google_sheet_name='CLTS Comittee Establishment', order=o, output=r'C:\FHIDatabase\data\CBNP data to be uploade.xlsx', ipname='cha')
#
#
# def upload_CHA_postodf_follow_ups():
#
#     o = ['ipname','province','district','dstcode','community_name','podfvisitnumber','visit_date','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
#     main.load_data_to_upload(from_excel=main.post_odf_visits_table, from_google = 'CLTS-CHA', google_sheet_name='Post ODF Followup visit', order=o, output=r'C:\FHIDatabase\data\CBNP data to be uploade.xlsx', ipname='cha')
#
#
# def upload_CHA_fhag_comittee_establishment():
#
#     o = ['ipname','province','district','dstcode','ommunituy_name','fhagcomestab','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
#     main.load_data_to_upload(from_excel=main.fhag_committee_establishment_table, from_google='CLTS-CHA', google_sheet_name='FHAG Comitte Establishment', order=o, output=r'C:\FHIDatabase\data\CBNP data to be uploade.xlsx', ipname='cha')
#
# def upload_CHA_Sustain_plan_Dev():
#
#     o = ['ipname','province','district','dstcode','communituy_name','sustain_plan_developed','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
#     main.load_data_to_upload(from_excel=main.fhag_committee_establishment_table, from_google='CLTS-IRC', google_sheet_name='FHAG Comitte Establishment', order=o, output=r'C:\FHIDatabase\data\CBNP data to be uploade.xlsx', ipname='irc')

# HADAAF CLTS data to be uploaded
# def upload_HADAAF_Triggering():
#
#     o = ['ipname','province','district','dstcode','cdc','village','community_name','lat','long','coordinates','_coordinates_latitude','_coordinates_longitude','_coordinates_altitude','_coordinates_precision','triggeringdate','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
#     main.load_data_to_upload(from_excel=main.triggering_table, from_google='CLTS-HADAAF', google_sheet_name='Triggering', order=o, output=r'C:\FHIDatabase\data\CBNP data to be uploade.xlsx', ipname='hadaaf')
#
# def upload_HADAAF_demographics():
#
#     o = ['ipname','province','district','dstcode','communituy_name','houseincommunity','nohouseholdincommunity','nomalepopincommuniyt','nofemalepopincommounity','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
#     main.load_data_to_upload(from_excel=main.demographics_table, from_google = 'CLTS-HADAAF', google_sheet_name='Demographics', order=o, output=r'C:\FHIDatabase\data\CBNP data to be uploade.xlsx', ipname='hadaaf')
#
#
# def upload_HADAAF_Existing_latrines_report():
#
#     o = ['ipname','province','district','dstcode','communituy_name','latrines_before_traiggering','existing_unimproved_Latrines','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
#     main.load_data_to_upload(from_excel=main.latrines_before_triggering_table, from_google = 'CLTS-HADAAF', google_sheet_name='Existing Latrines Report', order=o, output=r'C:\FHIDatabase\data\CBNP data to be uploade.xlsx', ipname='hadaaf')
#
# def upload_HADAAF_CLTS_Training():
#
#     o = ['ipname','province','district','dstcode','communituy_name','number_clts_trained','iecmaterailsused','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
#     main.load_data_to_upload(from_excel=main.clts_training_table, from_google = 'CLTS-HADAAF', google_sheet_name='CLTS Training', order=o, output=r'C:\FHIDatabase\data\CBNP data to be uploade.xlsx', ipname='hadaaf')
#
# def upload_HADAAF_FHAG_Training():
#
#     o = ['ipname','province','district','dstcode','communituy_name','num_fhag','iecmaterailsused','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name', '_parent_index','_tags','_notes']
#     main.load_data_to_upload(from_excel=main.fhag_training_table, from_google = 'CLTS-HADAAF', google_sheet_name='FHAG Training', order=o, output=r'C:\FHIDatabase\data\CBNP data to be uploade.xlsx', ipname='hadaaf')
#
#
# def upload_HADAAF_triggering_in_school():
#
#     o = ['ipname','province','district','dstcode','communituy_name','schoolinthisvillage','hysanitseessinheld','nochilbenfrothissession','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
#     main.load_data_to_upload(from_excel=main.triggering_in_school_table, from_google = 'CLTS-HADAAF', google_sheet_name='Triggering in School', order=o, output=r'C:\FHIDatabase\data\CBNP data to be uploade.xlsx', ipname='hadaaf')
#
#
# def upload_HADAAF_FHAG_Households_coverage_form():
#
#     o = ['ipname','province','district','dstcode','communituy_name','number_hhvisitedbyfhag','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
#     main.load_data_to_upload(from_excel=main.fhag_hh_coverage_table, from_google = 'CLTS-HADAAF', google_sheet_name='FHAG Households Coverage Form', order=o, output=r'C:\FHIDatabase\data\CBNP data to be uploade.xlsx', ipname='hadaaf')
#
# def upload_HADAAF_post_triggering_visits():
#
#     o = ['ipname','province','district','dstcode','community_name','visitnumber','visit_date','no_latrines_newly_built','no_latrines_upgraded','no_male_benef','no_female_benef','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags _notes']
#     main.load_data_to_upload(from_excel=main.visits_table, from_google = 'CLTS-HADAAF', google_sheet_name='Post Triggering Visits', order=o, output=r'C:\FHIDatabase\data\CBNP data to be uploade.xlsx', ipname='hadaaf')
#
# def upload_HADAAF_clts_meetings_record():
#
#     o = ['ipname','province','district','dstcode','communituy_name','number_clts_meetings','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
#     main.load_data_to_upload(from_excel=main.clts_committee_held_meetings_table, from_google = 'CLTS-HADAAF', google_sheet_name='CLTS Meetings Record', order=o, output=r'C:\FHIDatabase\data\CBNP data to be uploade.xlsx', ipname='hadaaf')
#
# def upload_HADAAF_claiming():
#
#     o = ['ipname','province','district','dstcode','communituy_name','claim_date','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
#     main.load_data_to_upload(from_excel=main.claiming_odf_table, from_google = 'CLTS-HADAAF', google_sheet_name='Claiming ODF', order=o, output=r'C:\FHIDatabase\data\CBNP data to be uploade.xlsx', ipname='hadaaf')
#
# def upload_HADAAF_verification():
#
#     o = ['ipname','province','district','dstcode','communituy_name','verified_date','verifiedimage','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
#     main.load_data_to_upload(from_excel=main.verified_odf_table, from_google = 'CLTS-HADAAF', google_sheet_name='Verification ODF', order=o, output=r'C:\FHIDatabase\data\CBNP data to be uploade.xlsx', ipname='hadaaf')
#
#
# def upload_HADAAF_certification():
#
#     o = ['ipname','province','district','dstcode','communituy_name','certification_date','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
#     main.load_data_to_upload(from_excel=main.certified_odf_table, from_google = 'CLTS-HADAAF', google_sheet_name='Certification ODF', order=o, output=r'C:\FHIDatabase\data\CBNP data to be uploade.xlsx', ipname='hadaaf')
#
# def upload_HADAAF_dysfunc_watersource():
#
#     o = ['ipname','province','district','dstcode','communituy_name','number_dysfwatersources','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
#     main.load_data_to_upload(from_excel=main.dysfunc_water_sources_table, from_google = 'CLTS-HADAAF', google_sheet_name='DysFunc Water Sources', order=o, output=r'C:\FHIDatabase\data\CBNP data to be uploade.xlsx', ipname='hadaaf')
#
# def upload_HADAAF_dysfunc_watersource_madefunc():
#
#     o = ['ipname','province','district','dstcode','communituy_name','number_dysfwatersourmf','benifwatersourmf','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
#     main.load_data_to_upload(from_excel=main.dysfunc_sources_functional_table, from_google = 'CLTS-HADAAF', google_sheet_name='DysFunc Made Func', order=o, output=r'C:\FHIDatabase\data\CBNP data to be uploade.xlsx', ipname='hadaaf')
#
# def upload_HADAAF_clts_comittee_establishment():
#
#     o = ['ipname','province','district','dstcode','communituy_name','cltscomestab','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
#     main.load_data_to_upload(from_excel=main.clts_committee_establishment_table, from_google = 'CLTS-HADAAF', google_sheet_name='CLTS Comittee Establishment', order=o, output=r'C:\FHIDatabase\data\CBNP data to be uploade.xlsx', ipname='hadaaf')
#
#
# def upload_HADAAF_postodf_follow_ups():
#
#     o = ['ipname','province','district','dstcode','community_name','podfvisitnumber','visit_date','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
#     main.load_data_to_upload(from_excel=main.post_odf_visits_table, from_google = 'CLTS-HADAAF', google_sheet_name='Post ODF Followup visit', order=o, output=r'C:\FHIDatabase\data\CBNP data to be uploade.xlsx', ipname='hadaaf')
#
#
# def upload_HADAAF_fhag_comittee_establishment():
#
#     o = ['ipname','province','district','dstcode','ommunituy_name','fhagcomestab','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
#     main.load_data_to_upload(from_excel=main.fhag_committee_establishment_table, from_google='CLTS-HADAAF', google_sheet_name='FHAG Comitte Establishment', order=o, output=r'C:\FHIDatabase\data\CBNP data to be uploade.xlsx', ipname='hadaaf')
#
# def upload_HADAAF_Sustain_plan_Dev():
#
#     o = ['ipname','province','district','dstcode','communituy_name','sustain_plan_developed','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
#     main.load_data_to_upload(from_excel=main.fhag_committee_establishment_table, from_google='CLTS-HADAAF', google_sheet_name='FHAG Comitte Establishment', order=o, output=r'C:\FHIDatabase\data\CBNP data to be uploade.xlsx', ipname='hadaaf')


# pasrse cbnp data to clean data

def upload_Triggering():

    o = ['ipname','province','district','dstcode','cdc','village','community_name','lat','long','coordinates','_coordinates_latitude','_coordinates_longitude','_coordinates_altitude','_coordinates_precision','triggeringdate','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
    main.parse_clts_data(source=main.triggering_table, clean_data=main.triggering_table_clean, order=o)

def upload_demographics():

    o = ['ipname','province','district','dstcode','communituy_name','houseincommunity','nohouseholdincommunity','nomalepopincommuniyt','nofemalepopincommounity','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
    main.parse_clts_data(source=main.demographics_table, clean_data=main.demographics_table_clean, order=o)


def upload_Baseline():

    o = ['ipname','province','district','dstcode','community_name','household_head_name','household_head_fname','household_head_sex','household_head_contactno','no_male_memebers','no_female_memebers','latrine_exists','primaryimgae','latrine_type','latrine_type/flush','latrine_type/pit','latrine_type/ventilation_pipe_improved','latrine_type/composite','latrine_type/traditional','latrine_flush_type','pit_latrine_type','no_facility','latrine_status','latrine_facilities','latrine_facilities/door_curtain','latrine_facilities/vent_pipe','latrine_facilities/window_net','latrine_facilities/evacuation_door','latrine_facilities/cover_for_hole','latrine_facilities/hand_washing_facility','latrine_facilities/soap_ash','latrine_facilities/water','latrine_facilities/none','distance_watersource_latrine','latrine_improved','excreta_undersoil_6months','opendefecation_inyard','type_waste_managment','type_waste_managment/solide_waste_managment','type_waste_managment/wastewater_managment','type_waste_managment/no_waste_managment','yard_clean','water_source','water_source/spring_or_kariz_protected','water_source/spring_or_kariz_unprotected','water_source/stream_river_clean','water_source/stream_river_uncleaned','water_source/protected_well','water_source/unprotected_well','water_source/tap','water_source/other','other_treatment_method','distance_source_drinking_water','treatment_method','treatment_method/boile','treatment_method/bleach_chlorine','treatment_method/strain_through_cloth','treatment_method/water_filter','treatment_method/solar_dysinfiction','treatment_method/letit_sand_settle','treatment_method/no_tratment_method','treatment_method/other','other_treatment_method_001','dysfunctional_pump_availability','nohh_use_thispump','doyou_know_reason_dysfucntionality','reason-dysfunctionality','any_attempt_made_functionality','repair_level_required','is_repair_beneficial','date','remarks','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
    main.parse_clts_data(source=main.baseline_collected_table, clean_data=main.baseline_collected_table_clean, order=o)


def upload_Existing_latrines_report():

    o = ['ipname','province','district','dstcode','communituy_name','latrines_before_traiggering','existing_unimproved_Latrines','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
    main.parse_clts_data(source=main.latrines_before_triggering_table, order=o, clean_data=main.latrines_before_triggering_table_clean)

def upload_CLTS_Training():

    o = ['ipname','province','district','dstcode','communituy_name','number_clts_trained','iecmaterailsused','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
    main.parse_clts_data(source=main.clts_training_table, clean_data=main.clts_training_table_clean, order=o)

def upload_FHAG_Training():

    o = ['ipname','province','district','dstcode','communituy_name','num_fhag','iecmaterailsused','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name', '_parent_index','_tags','_notes']
    main.parse_clts_data(source=main.fhag_training_table, clean_data=main.fhag_training_table_clean, order=o)


def upload_Triggering_in_school():

    o = ['ipname','province','district','dstcode','communituy_name','schoolinthisvillage','hysanitseessinheld','nochilbenfrothissession','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
    main.parse_clts_data(source=main.triggering_in_school_table, clean_data=main.triggering_in_school_table_clean, order=o)


def upload_FHAG_Households_coverage_form():

    o = ['ipname','province','district','dstcode','communituy_name','number_hhvisitedbyfhag','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
    main.parse_clts_data(source=main.fhag_hh_coverage_table, clean_data = main.fhag_hh_coverage_table_clean, order=o)

def upload_post_triggering_visits():

    o = ['ipname','province','district','dstcode','community_name','visitnumber','visit_date','no_latrines_newly_built','no_latrines_upgraded','no_male_benef','no_female_benef','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
    main.parse_clts_data(source=main.visits_table, clean_data = main.visits_table_clean, order=o)

def upload_clts_meetings_record():

    o = ['ipname','province','district','dstcode','communituy_name','number_clts_meetings','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
    main.parse_clts_data(source=main.clts_committee_held_meetings_table, clean_data=main.clts_committee_held_meetings_table_clean, order=o)

def upload_claiming():

    o = ['ipname','province','district','dstcode','communituy_name','claim_date','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
    main.parse_clts_data(source=main.claiming_odf_table, clean_data = main.claiming_odf_table_clean, order=o)

def upload_verification():

    o = ['ipname','province','district','dstcode','communituy_name','verified_date','verifiedimage','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
    main.parse_clts_data(source=main.verified_odf_table, clean_data = main.verified_odf_table_clean, order=o)


def upload_certification():

    o = ['ipname','province','district','dstcode','communituy_name','certification_date','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
    main.parse_clts_data(source=main.certified_odf_table, clean_data = main.certified_odf_table_clean, order=o)

def upload_dysfunc_watersource():

    o = ['ipname','province','district','dstcode','communituy_name','number_dysfwatersources','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
    main.parse_clts_data(source=main.dysfunc_water_sources_table, clean_data=main.dysfunc_water_sources_table_clean, order=o)

def upload_dysfunc_watersource_madefunc():

    o = ['ipname','province','district','dstcode','communituy_name','number_dysfwatersourmf','benifwatersourmf','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
    main.parse_clts_data(source=main.dysfunc_sources_functional_table, clean_data=main.dysfunc_sources_functional_table_clean, order=o)

def upload_clts_comittee_establishment():

    o = ['ipname','province','district','dstcode','communituy_name','cltscomestab','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
    main.parse_clts_data(source=main.clts_committee_establishment_table, clean_data = main.clts_committee_establishment_table_clean, order=o)


def upload_postodf_follow_ups():

    o = ['ipname','province','district','dstcode','community_name','podfvisitnumber','visit_date','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
    main.parse_clts_data(source=main.post_odf_visits_table, clean_data=main.post_odf_visits_table_clean, order=o)


def upload_fhag_comittee_establishment():

    o = ['ipname','province','district','dstcode','communituy_name','fhagcomestab','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
    main.parse_clts_data(source=main.fhag_committee_establishment_table, clean_data=main.fhag_committee_establishment_table_clean, order=o)

def upload_Sustain_plan_Dev():

    o = ['ipname','province','district','dstcode','communituy_name','sustain_plan_developed','__version__','meta/instanceID','_id','_uuid','_submission_time','_index','_parent_table_name','_parent_index','_tags','_notes']
    main.parse_clts_data(source=main.sustainability_plan_table, clean_data=main.sustainability_plan_table_clean, order=o)




def parse_data():

    sub = combo_subcontractors.get()
    data = combo_data.get()
    
    # AKF
    if sub == 'AKF' and data == 'CBNP':
        upload_AKF_cbnp()
    elif sub == 'AKF' and data == 'CBNP-CHS':
        upload_AKF_cbnp_chs()
    elif sub == 'AKF' and data == 'Trainings':
        upload_AKF_Trainings()
    elif sub == 'AKF' and data == 'Supportive Supervision':
        upload_AKF_Supportive_Supervision()
    elif sub == 'AKF' and data == 'SBCC':
        upload_AKF_SBCC_Report()
    elif sub == 'AKF' and data == 'Site Visit':
        upload_AKF_Site_visit()
    elif sub == 'AKF' and data == 'Mobile Cinema':
        upload_AKF_MobileCinema()
    # elif sub == 'AKF' and data == 'CLTS-Triggering':
    #     upload_AKF_Triggering()
    # elif sub == 'AKF' and data == 'CLTS-Demographics':
    #     upload_AKF_demographics()
    elif sub == 'AKF' and data == 'Baseline':
        upload_AKF_Baseline()
    # elif sub == 'AKF' and data == 'CLTS-Existing latrines Report':
    #     upload_AKF_Existing_latrines_report()
    # elif sub == 'AKF' and data == 'CLTS Training':
    #     upload_AKF_CLTS_Training()
    # elif sub == 'AKF' and data == 'CLTS FHAG Training':
    #     upload_AKF_FHAG_Training()
    # elif sub == 'AKF' and data == 'CLTS Triggering in School':
    #     upload_AKF_triggering_in_school()
    # elif sub == 'AKF' and data == 'CLTS FHAG Households Coverage':
    #     upload_AKF_FHAG_Households_coverage_form()
    # elif sub == 'AKF' and data == 'CLTS Post Triggering Vsist':
    #     upload_AKF_post_triggering_visits()
    # elif sub == 'AKF' and data == 'CLTS Claiming':
    #     upload_AKF_claiming()
    # elif sub == 'AKF' and data == 'CLTS Meetings Record':
    #     upload_AKF_clts_meetings_record()
    # elif sub == 'AKF' and data == 'CLTS Verification':
    #     upload_AKF_verification()
    # elif sub == 'AKF' and data == 'CLTS Certification':
    #     upload_AKF_certification()
    # elif sub == 'AKF' and data == 'CLTS Number of disfunctional water sources':
    #     upload_AKF_dysfunc_watersource()
    # elif sub == 'AKF' and data == 'CLTS Number of disfunctional water sources made functional':
    #     upload_AKF__dysfunc_watersource_madefunc()
    # elif sub == 'AKF' and data == 'CLTS Committee Establishment':
    #     upload_AKF_clts_comittee_establishment()
    # elif sub == 'AKF' and data == 'CLTS PostODF Followup Visists':
    #     upload_AKF_postodf_follow_ups()
    # elif sub == 'AKF' and data == 'CLTS FHAG Comittee Establishment':
    #     upload_AKF_fhag_comittee_establishment()
    # elif sub == 'AKF' and data == 'CLTS Community Sustainability plan Development':
    #     upload_AKF_Sustain_plan_Dev()
    elif sub == 'AKF' and data == 'Endline':
        upload_AKF_Endline()

    # IRC
    elif sub == 'IRC' and data == 'CBNP':
        upload_IRC_cbnp()
    elif sub == 'IRC' and data == 'CBNP-CHS':
        upload_IRC_cbnp_chs()
    elif sub == 'IRC' and data == 'Trainings':
        upload_IRC_Trainings()
    elif sub == 'IRC' and data == 'Supportive Supervision':
        upload_IRC_Supportive_Supervision()
    elif sub == 'IRC' and data == 'SBCC':
        upload_IRC_SBCC_Report()
    elif sub == 'IRC' and data == 'Site Visit':
        upload_IRC_Site_visit()
    elif sub == 'IRC' and data == 'Mobile Cinema':
        upload_IRC_MobileCinema()
    # elif sub == 'IRC' and data == 'CLTS-Triggering':
    #     upload_IRC_Triggering()
    # elif sub == 'IRC' and data == 'CLTS-Demographics':
    #     upload_IRC_demographics()
    elif sub == 'IRC' and data == 'Baseline':
        upload_IRC_Baseline()
    # elif sub == 'IRC' and data == 'CLTS-Existing latrines Report':
    #     upload_IRC_Existing_latrines_report()
    # elif sub == 'IRC' and data == 'CLTS Training':
    #     upload_IRC_CLTS_Training()
    # elif sub == 'IRC' and data == 'CLTS FHAG Training':
    #     upload_IRC_FHAG_Training()
    # elif sub == 'IRC' and data == 'CLTS Triggering in School':
    #     upload_IRC_triggering_in_school()
    # elif sub == 'IRC' and data == 'CLTS FHAG Households Coverage':
    #     upload_IRC_FHAG_Households_coverage_form()
    # elif sub == 'IRC' and data == 'CLTS Post Triggering Vsist':
    #     upload_IRC_post_triggering_visits()
    # elif sub == 'IRC' and data == 'CLTS Claiming':
    #     upload_IRC_claiming()
    # elif sub == 'IRC' and data == 'CLTS Meetings Record':
    #     upload_IRC_clts_meetings_record()
    # elif sub == 'IRC' and data == 'CLTS Verification':
    #     upload_IRC_verification()
    # elif sub == 'IRC' and data == 'CLTS Certification':
    #     upload_IRC_certification()
    # elif sub == 'IRC' and data == 'CLTS Number of disfunctional water sources':
    #     upload_IRC_dysfunc_watersource()
    # elif sub == 'IRC' and data == 'CLTS Number of disfunctional water sources made functional':
    #     upload_IRC_dysfunc_watersource_madefunc()
    # elif sub == 'IRC' and data == 'CLTS Committee Establishment':
    #     upload_IRC_clts_comittee_establishment()
    # elif sub == 'IRC' and data == 'CLTS PostODF Followup Visists':
    #     upload_IRC_postodf_follow_ups()
    # elif sub == 'IRC' and data == 'CLTS FHAG Comittee Establishment':
    #     upload_IRC_fhag_comittee_establishment()
    # elif sub == 'IRC' and data == 'CLTS Community Sustainability plan Development':
    #     upload_IRC_Sustain_plan_Dev()
    elif sub == 'IRC' and data == 'Endline':
        upload_IRC_Endline()

    # CHA
    elif sub == 'CHA' and data == 'CBNP':
        upload_CHA_cbnp()
    elif sub == 'CHA' and data == 'CBNP-CHS':
        upload_CHA_cbnp_chs()
    elif sub == 'CHA' and data == 'Trainings':
        upload_CHA_Trainings()
    elif sub == 'CHA' and data == 'Supportive Supervision':
        upload_CHA_Supportive_Supervision()
    elif sub == 'CHA' and data == 'SBCC':
        upload_CHA_SBCC_Report()
    elif sub == 'CHA' and data == 'Site Visit':
        upload_CHA_Site_visit()
    elif sub == 'CHA' and data == 'Mobile Cinema':
        upload_CHA_MobileCinema()
    # elif sub == 'CHA' and data == 'CLTS-Triggering':
    #     upload_CHA_Triggering()
    # elif sub == 'CHA' and data == 'CLTS-Demographics':
    #     upload_CHA_demographics()
    elif sub == 'CHA' and data == 'Baseline':
        upload_CHA_Baseline()
    # elif sub == 'CHA' and data == 'CLTS-Existing latrines Report':
    #     upload_CHA_Existing_latrines_report()
    # elif sub == 'CHA' and data == 'CLTS Training':
    #     upload_CHA_CLTS_Training()
    # elif sub == 'CHA' and data == 'CLTS FHAG Training':
    #     upload_CHA_FHAG_Training()
    # elif sub == 'CHA' and data == 'CLTS Triggering in School':
    #     upload_CHA_triggering_in_school()
    # elif sub == 'CHA' and data == 'CLTS FHAG Households Coverage':
    #     upload_CHA_FHAG_Households_coverage_form()
    # elif sub == 'CHA' and data == 'CLTS Post Triggering Vsist':
    #     upload_CHA_post_triggering_visits()
    # elif sub == 'CHA' and data == 'CLTS Claiming':
    #     upload_CHA_claiming()
    # elif sub == 'CHA' and data == 'CLTS Meetings Record':
    #     upload_CHA_clts_meetings_record()
    # elif sub == 'CHA' and data == 'CLTS Verification':
    #     upload_CHA_verification()
    # elif sub == 'CHA' and data == 'CLTS Certification':
    #     upload_CHA_certification()
    # elif sub == 'CHA' and data == 'CLTS Number of disfunctional water sources':
    #     upload_CHA_dysfunc_watersource()
    # elif sub == 'CHA' and data == 'CLTS Number of disfunctional water sources made functional':
    #     upload_CHA_dysfunc_watersource_madefunc()
    # elif sub == 'CHA' and data == 'CLTS Committee Establishment':
    #     upload_CHA_clts_comittee_establishment()
    # elif sub == 'CHA' and data == 'CLTS PostODF Followup Visists':
    #     upload_CHA_postodf_follow_ups()
    # elif sub == 'CHA' and data == 'CLTS FHAG Comittee Establishment':
    #     upload_CHA_fhag_comittee_establishment()
    # elif sub == 'CHA' and data == 'CLTS Community Sustainability plan Development':
    #     upload_CHA_Sustain_plan_Dev()
    elif sub == 'CHA' and data == 'Endline':
        upload_CHA_Endline()

    # HADAAF
    elif sub == 'HADAAF' and data == 'CBNP':
        upload_HADAAF_cbnp()
    elif sub == 'HADAAF' and data == 'CBNP-CHS':
        upload_HADAAF_cbnp_chs()
    elif sub == 'HADAAF' and data == 'Trainings':
        upload_HADAAF_Trainings()
    elif sub == 'HADAAF' and data == 'Supportive Supervision':
        upload_HADAAF_Supportive_Supervision()
    elif sub == 'HADAAF' and data == 'SBCC':
        upload_HADAAF_SBCC_Report()
    elif sub == 'HADAAF' and data == 'Site Visit':
        upload_HADAAF_Site_visit()
    elif sub == 'HADAAF' and data == 'Mobile Cinema':
        upload_HADAAF_MobileCinema()
    # elif sub == 'HADAAF' and data == 'CLTS-Triggering':
    #     upload_HADAAF_Triggering()
    # elif sub == 'HADAAF' and data == 'CLTS-Demographics':
    #     upload_HADAAF_demographics()
    elif sub == 'HADAAF' and data == 'Baseline':
        upload_HADAAF_Baseline()
    # elif sub == 'HADAAF' and data == 'CLTS-Existing latrines Report':
    #     upload_HADAAF_Existing_latrines_report()
    # elif sub == 'HADAAF' and data == 'CLTS Training':
    #     upload_HADAAF_CLTS_Training()
    # elif sub == 'HADAAF' and data == 'CLTS FHAG Training':
    #     upload_HADAAF_FHAG_Training()
    # elif sub == 'HADAAF' and data == 'CLTS Triggering in School':
    #     upload_HADAAF_triggering_in_school()
    # elif sub == 'HADAAF' and data == 'CLTS FHAG Households Coverage':
    #     upload_HADAAF_FHAG_Households_coverage_form()
    # elif sub == 'HADAAF' and data == 'CLTS Post Triggering Vsist':
    #     upload_HADAAF_post_triggering_visits()
    # elif sub == 'HADAAF' and data == 'CLTS Claiming':
    #     upload_HADAAF_claiming()
    # elif sub == 'HADAAF' and data == 'CLTS Meetings Record':
    #     upload_HADAAF_clts_meetings_record()
    # elif sub == 'HADAAF' and data == 'CLTS Verification':
    #     upload_HADAAF_verification()
    # elif sub == 'HADAAF' and data == 'CLTS Certification':
    #     upload_HADAAF_certification()
    # elif sub == 'HADAAF' and data == 'CLTS Number of disfunctional water sources':
    #     upload_HADAAF_dysfunc_watersource()
    # elif sub == 'HADAAF' and data == 'CLTS Number of disfunctional water sources made functional':
    #     upload_HADAAF_dysfunc_watersource_madefunc()
    # elif sub == 'HADAAF' and data == 'CLTS Committee Establishment':
    #     upload_HADAAF_clts_comittee_establishment()
    # elif sub == 'HADAAF' and data == 'CLTS PostODF Followup Visists':
    #     upload_HADAAF_postodf_follow_ups()
    # elif sub == 'HADAAF' and data == 'CLTS FHAG Comittee Establishment':
    #     upload_HADAAF_fhag_comittee_establishment()
    # elif sub == 'HADAAF' and data == 'CLTS Community Sustainability plan Development':
    #     upload_HADAAF_Sustain_plan_Dev()
    elif sub == 'HADAAF' and data == 'Endline':
        upload_HADAAF_Endline()
    else:
        print('No Data Here!')


def parse_clts_data():
        start = time.time()
        upload_Triggering()
        upload_demographics()
        upload_Baseline()
        upload_Existing_latrines_report()
        upload_CLTS_Training()
        upload_FHAG_Training()
        upload_Triggering_in_school()
        upload_FHAG_Households_coverage_form()
        upload_post_triggering_visits()
        upload_claiming()
        upload_clts_meetings_record()
        upload_verification()
        upload_certification()
        upload_dysfunc_watersource()
        upload_dysfunc_watersource_madefunc()
        upload_clts_comittee_establishment()
        upload_postodf_follow_ups()
        upload_fhag_comittee_establishment()
        upload_Sustain_plan_Dev()
        end = time.time()
        print('Data Successfully Parsed in {} seconde'.format(end-start))



mainWindow = tk.Tk()
mainWindow.title("Data Controlling Panel")
mainWindow.geometry('400x200+200+200')

mainWindow.columnconfigure(0, weight=1)
mainWindow.columnconfigure(1, weight=1)
mainWindow.columnconfigure(2, weight=1)
mainWindow.rowconfigure(0, weight=1)
mainWindow.rowconfigure(1, weight=1)
mainWindow.rowconfigure(2, weight=1)
mainWindow.rowconfigure(3, weight=1)
mainWindow.rowconfigure(4, weight=1)
mainWindow.rowconfigure(5, weight=1)
mainWindow.rowconfigure(6, weight=1)
mainWindow.rowconfigure(7, weight=1)
mainWindow.rowconfigure(8, weight=1)
mainWindow.rowconfigure(9, weight=1)
mainWindow.rowconfigure(10, weight=1)
mainWindow.rowconfigure(11, weight=1)
mainWindow.rowconfigure(12, weight=1)
mainWindow.rowconfigure(13, weight=1)

subcontractors = ['IRC', 'CHA', 'AKF', 'HADAAF']
# data = ['CBNP', 'CBNP-CHS', 'Trainings', 'Meetings', 'Supportive Supervision', 'SBCC', 'Site Visit', 'Mobile Cinema', 'CLTS-Triggering','CLTS-Demographics','CLTS-Baseline','CLTS-Existing latrines Report', 'CLTS Training',
#         'CLTS FHAG Training', 'CLTS Triggering in School', 'CLTS FHAG Households Coverage', 'CLTS Post Triggering Vsist', 'CLTS Claiming', 'CLTS Meetings Record', 'CLTS Final Latrines Report', 'CLTS Verification',
#         'CLTS Certification', 'CLTS Number of dysfunctional water sources', 'CLTS Number of disfunctional water sources made functional', 'CLTS Committee Establishment', 'CLTS PostODF Followup Visists', 'CLTS FHAG Comittee Establishment',
#         'CLTS Community Sustainability plan Development', 'Endline'
#         ]
data = ['CBNP', 'CBNP-CHS', 'Trainings', 'Meetings', 'Supportive Supervision', 'SBCC', 'Site Visit', 'Mobile Cinema','Baseline','Endline']


combo_subcontractors = Combobox(mainWindow, values=subcontractors)
combo_data = Combobox(mainWindow, values=data)
lbl_data = tk.Label(mainWindow, text='Data:  ')
lbl_sub = tk.Label(mainWindow,text='Subcontractor:  ')
btn_refreshCommunitiesList = tk.Button(mainWindow, height=1, width=20, text='Refresh Communities List', command=refresh_communities_list)
btn_refreshAll = tk.Button(mainWindow, height=1, width=20, text='Refresh All', command=main.refresh_all)
btn_updateAllCommunities = tk.Button(mainWindow, height=1, width=20, text='Update All Communities', command=main.update_all_communities)
btn_updatefromxls = tk.Button(mainWindow, height=1, width=20, text='Update from XLS', command=main.append_from_xls)
btn_loadData = tk.Button(mainWindow, height=1, width=20, text='Load Data', command=parse_data)
btn_parse_clts_data = tk.Button(mainWindow, height=1, width=20, text='Parse CLTS Data', command=parse_clts_data)
btn_loadCLTSData = tk.Button(mainWindow, height=1, width=20, text='Load CLTS Data', command=main.load_data)
btn_loadCLTSDataMRRDFormat = tk.Button(mainWindow, height=1, width=20, text='Load CLTS Data in MRRD format', command=main.clts_mrrd_format)
# btn_upload_IRC_cbnp = tk.Button(mainWindow, text='IRC CBNP to be uploaded', command=upload_IRC_cbnp)
# btn_upload_AKF_cbnp = tk.Button(mainWindow, text='AKF CBNP to be uploaded', command=upload_AKF_cbnp)
# btn_upload_CHA_cbnp = tk.Button(mainWindow, text='CHA CBNP to be uploaded', command=upload_CHA_cbnp)
# btn_upload_HADAAF_cbnp = tk.Button(mainWindow, text='HADAAF CBNP to be uploaded', command=upload_HADAAF_cbnp)
# btn_upload_IRC_cbnp_chs = tk.Button(mainWindow, text='IRC CBNP-CHS to be uploaded', command=upload_IRC_cbnp_chs)
# btn_upload_AKF_cbnp_chs = tk.Button(mainWindow, text='AKF CBNP-CHS to be uploaded', command=upload_AKF_cbnp_chs)
# btn_upload_CHA_cbnp_chs = tk.Button(mainWindow, text='CHA CBNP-CHS to be uploaded', command=upload_CHA_cbnp_chs)
# btn_upload_HADAAF_cbnp_chs = tk.Button(mainWindow, text='HADAAF CBNP-CHS to be uploaded', command=upload_HADAAF_cbnp_chs)
# btn_upload_IRC_SiteVisit = tk.Button(mainWindow, text='IRC Site Visit to be uploaded', command=upload_IRC_Site_visit)
# btn_upload_CHA_SiteVisit = tk.Button(mainWindow, text='CHA Site Visits to be uploaded', command=upload_CHA_Site_visit)
# btn_upload_AKF_SiteVisit = tk.Button(mainWindow, text='AKF Site Visits to be uploaded', command=upload_AKF_Site_visit)
# btn_upload_HADAAF_SiteVisit = tk.Button(mainWindow, text='HADAAF Site Visits to be uploaded', command=upload_HADAAF_Site_visit)
# btn_upload_IRC_Supportive_Supervision = tk.Button(mainWindow, text='IRC Supportive supervision to be uploaded', command=upload_IRC_Supportive_Supervision)
# btn_upload_AKF_Supportive_Supervision = tk.Button(mainWindow, text='AKF Supportive supervision to be uploaded', command=upload_AKF_Supportive_Supervision)
# btn_upload_CHA_Supportive_Supervision = tk.Button(mainWindow, text='CHA Supportive supervision to be uploaded', command=upload_CHA_Supportive_Supervision)
# btn_upload_HADAAF_Supportive_Supervision = tk.Button(mainWindow, text='HADAAF Supportive supervision to be uploaded', command=upload_HADAAF_Supportive_Supervision)
# btn_upload_IRC_SBCC = tk.Button(mainWindow, text='IRC SBCC to be uploaded', command=upload_IRC_SBCC_Report)
# btn_upload_AKF_SBCC = tk.Button(mainWindow, text='AKF SBCC to be uploaded', command=upload_AKF_SBCC_Report)
# btn_upload_CHA_SBCC= tk.Button(mainWindow, text='CHA SBCC to be uploaded', command=upload_CHA_SBCC_Report)
# btn_upload_HADAAF_SBCC = tk.Button(mainWindow, text='HADAAF SBCC to be uploaded', command=upload_HADAAF_SBCC_Report)
# btn_upload_IRC_MCinema = tk.Button(mainWindow, text='IRC Mobile Cinema to be uploaded', command=upload_IRC_MobileCinema)
# btn_upload_AKF_MCinema = tk.Button(mainWindow, text='AKF Mobile Cinema to be uploaded', command=upload_AKF_MobileCinema)
# btn_upload_CHA_MCinema = tk.Button(mainWindow, text='CHA Mobile Cinema to be uploaded', command=upload_CHA_MobileCinema)
# btn_upload_HADAAF_MCinema = tk.Button(mainWindow, text='HADAAF Mobile Cinema to be uploaded', command=upload_HADAAF_MobileCinema)
# btn_upload_IRC_Trainings = tk.Button(mainWindow, text='IRC Trainings to be uploaded', command=upload_IRC_Trainings)
# btn_upload_AKF_Trainings = tk.Button(mainWindow, text='AKF Trainings to be uploaded', command=upload_AKF_Trainings)
# btn_upload_CHA_Trainings = tk.Button(mainWindow, text='CHA Trainings to be uploaded', command=upload_CHA_Trainings)
# btn_upload_HADAAF_Trainings = tk.Button(mainWindow, text='HADAAF Trainings to be uploaded', command=upload_HADAAF_Trainings)
# btn_upload_IRC_Baseline = tk.Button(mainWindow, text='IRC Baseline to be uploaded', command=upload_IRC_Baseline)
# btn_upload_AKF_Baseline = tk.Button(mainWindow, text='AKF Baseline to be uploaded', command=upload_AKF_Baseline)
# btn_upload_CHA_Baseline = tk.Button(mainWindow, text='CHA Baseline to be uploaded', command=upload_CHA_Baseline)
# btn_upload_HADAAF_Baseline = tk.Button(mainWindow, text='HADAAF Baseline to be uploaded', command=upload_HADAAF_Baseline)
# btn_upload_IRC_Endline = tk.Button(mainWindow, text='IRC Endline to be uploaded', command=upload_IRC_Endline)
# btn_upload_AKF_Endline = tk.Button(mainWindow, text='AKF Endline to be uploaded', command=upload_AKF_Endline)
# btn_upload_CHA_Endline = tk.Button(mainWindow, text='CHA Endline to be uploaded', command=upload_CHA_Endline)
# btn_upload_HADAAF_Endline = tk.Button(mainWindow, text='HADAAF Endline to be uploaded', command=upload_HADAAF_Endline)

btn_refreshCommunitiesList.grid(row='0' ,column='0', columnspan=1)
btn_refreshAll.grid(row='1', column='0', columnspan=1)
btn_updateAllCommunities.grid(row='2', column='0', columnspan=1)
btn_updatefromxls .grid(row='3', column='0', columnspan=1)
btn_loadData.grid(row='4', column='0', columnspan=1)
combo_subcontractors.grid(row='1', column='2', sticky='w', columnspan=1)
lbl_sub.grid(row='1', column='1',sticky='e', columnspan=1)
combo_data.grid(row='2', column='2', sticky='w', columnspan=1)
lbl_data.grid(row='2', column='1',sticky='e', columnspan=1)
btn_parse_clts_data.grid(row='5',column='0', columnspan=1)
btn_loadCLTSData.grid(row='6',column='0', columnspan=1)
btn_loadCLTSDataMRRDFormat .grid(row='7',column='0', columnspan=1)
# btn_upload_IRC_cbnp.grid(row='1', column='2', columnspan=1)
# btn_upload_AKF_cbnp.grid(row='2', column='0', columnspan=1)
# btn_upload_CHA_cbnp.grid(row='2', column='1', columnspan=1)
# btn_upload_HADAAF_cbnp.grid(row='2', column='2', columnspan=1)
# btn_upload_IRC_cbnp_chs.grid(row='3', column='0', columnspan=1)
# btn_upload_AKF_cbnp_chs.grid(row='3', column='1', columnspan=1)
# btn_upload_CHA_cbnp_chs.grid(row='3', column='2', columnspan=1)
# btn_upload_HADAAF_cbnp_chs.grid(row='4', column='0', columnspan=1)
# btn_upload_IRC_SiteVisit.grid(row='4', column='1', columnspan=1)
# btn_upload_CHA_SiteVisit.grid(row='4', column='2', columnspan=1)
# btn_upload_AKF_SiteVisit.grid(row='5', column='0', columnspan=1)
# btn_upload_HADAAF_SiteVisit.grid(row='5', column='1', columnspan=1)
# btn_upload_IRC_Supportive_Supervision.grid(row='5', column='2', columnspan=1)
# btn_upload_AKF_Supportive_Supervision.grid(row='6', column='0', columnspan=1)
# btn_upload_CHA_Supportive_Supervision.grid(row='6', column='1', columnspan=1)
# btn_upload_HADAAF_Supportive_Supervision.grid(row='6', column='2', columnspan=1)
# btn_upload_IRC_SBCC.grid(row='7', column='0', columnspan=1)
# btn_upload_AKF_SBCC.grid(row='7', column='1', columnspan=1)
# btn_upload_CHA_SBCC.grid(row='7', column='2', columnspan=1)
# btn_upload_HADAAF_SBCC.grid(row='8', column='0', columnspan=1)
# btn_upload_IRC_MCinema.grid(row='8', column='1', columnspan=1)
# btn_upload_AKF_MCinema.grid(row='8', column='2', columnspan=1)
# btn_upload_CHA_MCinema.grid(row='9', column='0', columnspan=1)
# btn_upload_HADAAF_MCinema.grid(row='9', column='1', columnspan=1)
# btn_upload_IRC_Trainings.grid(row='9', column='2', columnspan=1)
# btn_upload_AKF_Trainings.grid(row='10', column='0', columnspan=1)
# btn_upload_CHA_Trainings.grid(row='10', column='1', columnspan=1)
# btn_upload_HADAAF_Trainings.grid(row='10', column='2', columnspan=1)
# btn_upload_IRC_Baseline.grid(row='11', column='0', columnspan=1)
# btn_upload_AKF_Baseline.grid(row='11', column='1', columnspan=1)
# btn_upload_CHA_Baseline.grid(row='11', column='2', columnspan=1)
# btn_upload_HADAAF_Baseline.grid(row='12', column='0', columnspan=1)
# btn_upload_IRC_Endline.grid(row='12', column='1', columnspan=1)
# btn_upload_AKF_Endline.grid(row='12', column='2', columnspan=1)
# btn_upload_CHA_Endline.grid(row='13', column='0', columnspan=1)
# btn_upload_HADAAF_Endline.grid(row='13', column='1', columnspan=1)


mainWindow.mainloop()
