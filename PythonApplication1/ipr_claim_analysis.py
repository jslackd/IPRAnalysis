"""Pull and assign information for targeted ipr claims"""
# Jonathan Slack
# jslackd@gmail.com

import os
import sys
from PIL import Image, ImageFilter
import pytesseract
import collections
import re
from wand.image import Image as IMG
import os
import xlsxwriter
import pandas as pd
import numpy as np
import PyPDF2
from multiprocessing.dummy import Pool as ThreadPool
from difflib import SequenceMatcher
import math
from datetime import date
from datetime import datetime
import urllib, json
import csv
import zipfile
import time
from shutil import copyfile

in_dir = "in_data"
fold = "all_iprs"
#fold = "test_docs2"
out_file = "ipr_read_data.xlsx"
out_file2 = "ipr_read_data+.xlsx"
out_file3 = "ipr_read_data++.xlsx"
temp_dir = "C:\\Users\\Johnny\\AppData\\Local\\Temp"
claim_file = "claim.tsv"
claimreduc_file = "claim_reduc.csv"

res = 400

# Declare data structures
ipr_data = collections.OrderedDict()
# Format: {filename.pdf (str): {attrib (str): val (---)}}
### These have already been found:
#       "trial_num(s)"  : ["IPR2015-00010"] or ["CBM2015-00004"] or ["PGR2015-00003"] or []
#       "trial_type"    : "IPR" or "CBM" or "PGR" or "Mult." or None
#       "dec_date"      : "12/16/2015" or None
#       "FWD?"          : True or False
#       "fd_type(s)"    : "FINAL WRITTEN DECISION" or "DECISION Termination of Trial" or None
#       "pat_num(s)"    : ["6658464"] or ["677234"] or ["43919"] or ["unknown"] or []
#       "pat_type(s)"   : ["PAT-B2"] or ["RE--E"] or ["D---"] or ["unknown"] or []
#       "mult_pat"      : True or False
#       "pet_name(s)"   : ["BIO-RAD LABORATORIES, INC.,"] or ["unknown"] or []
#       "ph_name(s)"    : ["CALIFORNIA INSTITUTE OF TECHNOLOGY"] or ["unknown"] or []
#       "no_issues"     : False or True (default is True)
#       "order_txt"     : "ORDERED that the joint motion to terminate the proceeding is GRANTED and . . ."
#       "order_disp(s)" : {"6658464": {"c-range": ["1-9,14"] , "disposition": ["unpatentable"]}}
#       "new_page?"     : False or True (default is False)
#       "expect_ccd"    : False or True (default is False)
#       "no_issues2"    : True or False (default is True)

iprclaim_data = collections.OrderedDict()
# Format: {patent number (str): {attrib (str): val (---)}}
### Finding/assigning the following:
#       "trial_num(s)"  : [["IPR2015-00010", "IPR2014-00213"],[...]]
#       "filename(s)"   : ["0001-IPR2013-00562 FD", "0001-IPR2014-01565 FWD Final"]
#       "dec_date(s)"   : ["12/16/2015", "12/21/2015"]
#       "pat_type"      : "PAT-B2"
#       "ph_name(s)"    : [["CALIFORNIA INSTITUTE OF TECHNOLOGY", "ACCENTURE, INC"]] 
#       "app_num"       : "10386691"
#       "continuation?" : False
#       "NBER_cats"     : {"cat_id" : "3", "subcat_id" : "32", "cat_name": "Drgs&Med", "subcat_name": "Surgery & Med Inst."}
#       "num_claims"    : 24
#       "no_issues3"    : True
#       "claim_msm?"    : False
#       "first_claim"   : {"text": "A means for ...", "orig_text": "A means for ..." "cat.": "independent", "root": None, "word_cnt": 45, "word_change": 13}
#### The following are disposed claims:
#       "1"             : {"text": "A means for ...", "dispo": "unpatentable", "date": "12/21/2015", "cat.": "independent", "root": None, "word_cnt": 45}
#       "2"             : {"text": "The means for ...","dispo": "unpatentable", "date": "12/21/2015", "cat.": "dependent", "root": "1", "word_cnt": 20}
#       "3"             : {"text": "The means for ..."., "dispo": "unpatentable", "date": "12/21/2015", "cat.": "dependent", "root": "1", "word_cnt": 22}
#       ...             : {...}
#       "24"            : {"text": "The means for ...", "dispo": "unpatentable", "date": "12/21/2015", "cat.": "dependent", "root": "16", "word_cnt": 19}

def create_dictionary_entry(fname):
    ipr_data[fname] = {
        "trial_num(s)": [], "trial_type": None, "fd_type(s)": None, "mult_pat": False,
        "dec_date": None, "pet_name(s)": [], "ph_name(s)": [], "pat_num(s)": None,
        "pat_type(s)": [], "order_txt": None, "order_disp(s)": collections.OrderedDict(), 
        "no_issues": True, "FWD?": False, "no_issues2": True, "expect_ccd": True
    }

def create_dictionary_entry2(fname):
    iprclaim_data[fname] = {
        "trial_num(s)": [], "filename(s)" : [], "dec_date(s)": [], "pat_type": None,
        "ph_name(s)": [], "NBER_cats": {}, "num_claims" : None, "first_claim": {"text":None,"orig_text":None,"cat.":"independent","root":None,"word_cnt":None,"word_change":None}, 
        "continuation?": False, "claim_msm?" : False, "no_issues3": True, "app_num": None
    }

def enter_claim_dict_entries(pat_num, claim_nums, claim_disp, date, existing = None):
    claims = []; dispos = []
    disptxts = claim_disp
    # Compile symmetric lists of claims and their dispositions
    n = 0
    for cntxts in claim_nums:
        cntxts = cntxts.replace(" ","")
        cls = cntxts.split(",")
        for cl in cls:
            hypos = cl.find("-")
            if hypos == -1:
                claims.append(cl)
                dispos.append(disptxts[n])
            else:
                claims.append(cl[:hypos])
                dispos.append(disptxts[n])
                claims.append(cl[hypos+1:])
                dispos.append(disptxts[n])
        n += 1

    if existing is None:
        m = 0
        for claim in claims:
            iprclaim_data[pat_num][claim] = {"text":None,"dispo":dispos[m],"date":date,"cat.":None,"root":None,"word_cnt":None}
            m += 1
    else:
        m = 0
        for claim in claims:
            if claim not in existing:
                iprclaim_data[pat_num][claim] = {"text":None,"dispo":dispos[m],"date":date,"cat.":None,"root":None,"word_cnt":None}
            else:
                old_date = iprclaim_data[pat_num][claim]["date"]
                dtold = datetime.strptime(old_date, "%m/%d/%Y")
                dtnew = datetime.strptime(date, "%m/%d/%Y")
                if dtnew > dtold:
                    iprclaim_data[pat_num][claim]["dispo"] = dispos[m]
                    iprclaim_data[pat_num][claim]["date"] = date
            m += 1
                
def pull_iprdata_ff(file_out):  
    # Read ipr data from 'IPR Trial Data'
    if os.path.isfile(file_out) == True:
        # Read existing file first sheet
        df = pd.read_excel(file_out, sheetname='IPR Trial Data')
        file_list_raw = df["Filename"].tolist()
        to_add_all = df.iloc[:,0:11]
        to_add_all = to_add_all.values.tolist()
    else:
        sys.exit("Error: ipr_read_data.xlsx DOES NOT EXIST")

    # Create a filename entry in ipr_data for each filename and pull data from file
    bin_trans = {"Yes": True, "No": False}
    bin_trans_rev = {"No": True, "Yes": False}
    keys = file_list_raw
    for key, data in zip(keys,to_add_all):

        create_dictionary_entry(key)

        trial_nums = str(data[0]).split("\n")
        if trial_nums != [""]:
            ipr_data[key]["trial_num(s)"] = trial_nums

        trial_type = str(data[1])
        if trial_type != "": 
            ipr_data[key]["trial_type"] = trial_type

        dec_date = str(data[2])
        if dec_date != "": 
            ipr_data[key]["dec_date"] = dec_date

        FWD = str(data[3])
        if FWD != "": 
            ipr_data[key]["FWD?"] = bin_trans[FWD]

        fd_types = str(data[4])
        if fd_types != "": 
            ipr_data[key]["fd_type(s)"] = fd_types

        pat_nums = str(data[5]).split("\n")
        if pat_nums != [""]:
            ipr_data[key]["pat_num(s)"] = pat_nums

        pat_types = str(data[6]).split("\n")
        if pat_types != [""]:
            ipr_data[key]["pat_type(s)"] = pat_types

        mult_pat = str(data[7])
        if mult_pat != "": 
            ipr_data[key]["mult_pat"] = bin_trans[mult_pat]

        pet_names = str(data[8]).split("\n")
        if pet_names != [""]:
            ipr_data[key]["pet_name(s)"] = pet_names

        ph_names = str(data[9]).split("\n")
        if ph_names != [""]:
            ipr_data[key]["ph_name(s)"] = ph_names

        no_issues = str(data[10])
        if no_issues != "": 
            ipr_data[key]["no_issues"] = bin_trans_rev[no_issues]

    # Read ipr data from 'IPR Patent Data'
    df = pd.read_excel(file_out, sheetname='IPR Patent Data')
    df = df.replace(np.nan, '', regex=True)
    file_list_raw = df["Filename"].tolist()
    to_add_all = df.iloc[:,0:13]
    to_add_all = to_add_all.values.tolist()

    # Pull data from 'IPR Patent Data' sheet
    bin_trans = {"Yes": True, "No": False}
    bin_trans_rev = {"No": True, "Yes": False}
    keys = file_list_raw; key_carry = ""
    for key, data in zip(keys,to_add_all):

        # We are in an order text row
        if key != "":

            order_txt = str(data[5])
            if order_txt != "":
                ipr_data[key]["order_txt"] = order_txt

            key_carry = key
        
        # We are in a patent row
        else:
            pat_n_ent = str(data[6])
            if pat_n_ent[-2:] == ".0": pat_n_ent = pat_n_ent [:-2]
            if pat_n_ent != "":
                ipr_data[key_carry]["order_disp(s)"][pat_n_ent] = {"c-range": [], "disposition": []}
            
            cranges = str(data[8]).split("\n")
            if cranges != [""]:
                ipr_data[key_carry]["order_disp(s)"][pat_n_ent]["c-range"] = cranges

            dispo = str(data[9]).split("\n")
            if dispo != [""]:
                ipr_data[key_carry]["order_disp(s)"][pat_n_ent]["disposition"] = dispo

            issue2 = str(data[10])
            if issue2 != "": 
                ipr_data[key_carry]["no_issues2"] = bin_trans_rev[issue2]

            ccderr = str(data[11])
            if ccderr != "": 
                ipr_data[key_carry]["expect_ccd"] = bin_trans[ccderr]

            mp_order = str(data[12])
            if mp_order != "": 
                ipr_data[key_carry]["new_page?"] = bin_trans[mp_order]

def transfer_ipr_data(targs):
    # Helper function for making sure our disposition vector is not empty
    def check_content(d):
        if len(d) == 0: return False
        for key in d.keys():
            if d[key]["c-range"] == [] or d[key]["c-range"] == [""]: return False
            if d[key]["c-range"] is None: return False
            if d[key]["disposition"] == [] or d[key]["disposition"] == [""]: return False
            if d[key]["disposition"] is None: return False
        return True

    # Loop through each targeted ipr, which should contain disposed patents (unless check_content is False)
    for targ in targs:
        dispos = ipr_data[targ]["order_disp(s)"]
        curr_keys = iprclaim_data.keys()

        # We have some empty values; don't write these
        if check_content(dispos) == False:
            continue
        
        pats = dispos.keys()
        j = 0
        for pat in pats:
            # We have a fresh patent number
            if pat not in curr_keys:
                # Record easy values into the dictionary
                create_dictionary_entry2(pat)
                iprclaim_data[pat]["trial_num(s)"].append(ipr_data[targ]["trial_num(s)"])
                iprclaim_data[pat]["filename(s)"].append(targ)
                iprclaim_data[pat]["dec_date(s)"].append(ipr_data[targ]["dec_date"])
                iprclaim_data[pat]["pat_type"] = ipr_data[targ]["pat_type(s)"][j]
                iprclaim_data[pat]["ph_name(s)"].append(ipr_data[targ]["ph_name(s)"])
                # Create dictionary entries for each claim, and change values
                date = ipr_data[targ]["dec_date"]
                crange_in = dispos[pat]["c-range"]
                dispo_in = dispos[pat]["disposition"]
                enter_claim_dict_entries(pat, crange_in, dispo_in,date)
            # We have a repeated patent number; must check value
            else:
                iprclaim_data[pat]["trial_num(s)"].append(ipr_data[targ]["trial_num(s)"])
                iprclaim_data[pat]["filename(s)"].append(targ)
                iprclaim_data[pat]["dec_date(s)"].append(ipr_data[targ]["dec_date"])
                iprclaim_data[pat]["ph_name(s)"].append(ipr_data[targ]["ph_name(s)"])
                # If needed created dictionary entries; if not, we must compare disposition and date
                existing =  iprclaim_data[pat]
                date = ipr_data[targ]["dec_date"]
                crange_in = dispos[pat]["c-range"]
                dispo_in = dispos[pat]["disposition"]
                enter_claim_dict_entries(pat, crange_in, dispo_in, date, existing)
            j+=1
                    
def patentsview_API_info(pat):
    # Pull patent information from PatentsView API
    if len(pat) == 0: 
        return {}, None, None, False
    error_ret = True
    ret1 = "%22patent_num_claims%22"
    ret2 = "%22nber_category_id%22"
    ret3 = "%22nber_subcategory_id%22"
    ret4 = "%22nber_category_title%22"
    ret5 = "%22nber_subcategory_title%22"
    ret6 = "%22app_number%22"
    fquery = ",".join([ret1,ret2,ret3,ret4,ret5,ret6])
    url_path = "http://www.patentsview.org/api/patents/query?q={%22patent_number%22:%22"+str(pat)+"%22}&f=["+fquery+"]"
    with urllib.request.urlopen(url_path) as url:
        datadown = json.loads(url.read().decode())
        if datadown["total_patent_count"] > 1: error_ret = False
        cat_id = datadown["patents"][0]["nbers"][0]["nber_category_id"]
        scat_id = datadown["patents"][0]["nbers"][0]["nber_subcategory_id"]
        cat_title = datadown["patents"][0]["nbers"][0]["nber_category_title"]
        scat_title = datadown["patents"][0]["nbers"][0]["nber_subcategory_title"]
        num_of_claims = datadown["patents"][0]["patent_num_claims"]
        app_num = datadown["patents"][0]["applications"][0]["app_number"]
        if "" in [cat_id, scat_id, cat_title, scat_title, num_of_claims, app_num]: error_ret = False
        nber_data = {"cat_id":cat_id, "subcat_id":scat_id, "cat_name":cat_title, "subcat_name":scat_title}
    return nber_data, num_of_claims, app_num, error_ret

def continuity_check(pat):
    # This function will download patent archive data, if needed, and check continuity data
    # Helper function for unziping files
    def unzip_folder(filenm, pat, appn):
        zip_ref = zipfile.ZipFile(filenm, 'r')
        out_folder = os.path.join(in_dir,"patent_data", pat, appn)
        os.makedirs(out_folder)
        zip_ref.extractall(out_folder)
        zip_ref.close()

    ##### Step A: Download application files ######
    app_num = iprclaim_data[pat]["app_num"]
    url = "http://storage.googleapis.com/uspto-pair/applications/" + app_num.replace(" ","") + ".zip"
    file_name = os.path.join(in_dir,"patent_data",pat, app_num + ".zip")
    missing_list = []

    # If folder for app exists, then skip this application number
    if os.path.isdir(os.path.join(in_dir,"patent_data",pat, app_num)) == True:
        pass
    # Only the zip file exists, so extract it and skip application number
    elif os.path.isfile(file_name) == True:
        # Unzip the file and delete the original zip file
        unzip_folder(file_name, pat, app_num)
        time.sleep(0.02)
        os.remove(file_name)

    # Download the file from `url` and save it locally under `file_name`:
    else:
        os.makedirs(os.path.join(in_dir,"patent_data",pat))
        skipper = False
        try:
            with urllib.request.urlopen(url) as response, open(file_name, 'wb') as out_file:
                data = response.read() # a `bytes` object
                out_file.write(data)
                out_file.close()               
        except TimeoutError:
            if os.path.isfile(file_name) == True:
                os.remove(file_name)
            os.makedirs(os.path.join(in_dir,"patent_data",pat,app_num))
            os.makedirs(os.path.join(in_dir,"patent_data",pat,app_num,app_num))
            copyfile(os.path.join(in_dir,"patent_data","-continuity_data.tsv"),os.path.join(in_dir,"patent_data",pat,app_num,app_num, "-continuity_data.tsv"))
            missing_list.append(pat)
            skipper = True
            #print("Patent: ",pat)
            #print("Application: ",app_num)
            #sys.exit("TimeoutError downloading application data. Try manually downloading the application")
        except urllib.error.HTTPError:
            if os.path.isfile(file_name) == True:
                os.remove(file_name)
            os.makedirs(os.path.join(in_dir,"patent_data",pat,app_num))
            os.makedirs(os.path.join(in_dir,"patent_data",pat,app_num,app_num))
            copyfile(os.path.join(in_dir,"patent_data","-continuity_data.tsv"),os.path.join(in_dir,"patent_data",pat,app_num,app_num, "-continuity_data.tsv"))
            missing_list.append(pat)
            skipper = True
            #print("Patent: ",pat)
            #print("Application: ",app_num)
            #sys.exit("HTTPError downloading application data. Try manually downloading the application")
        # Unzip the file and delete the original zip file
        if skipper == False:
            unzip_folder(file_name, pat, app_num)
            time.sleep(0.02)
            os.remove(file_name)

    ##### Step B: Analyze application files for continuity data ######
    cont_pull = False
    # Check if continuation info csv file
    path = os.path.join(in_dir,"patent_data", pat, app_num, app_num, app_num+"-continuity_data.tsv")
    if os.path.exists(path) == True:

        # Analyze the application info csv file
        with open(path) as tsvfile:
            reader = csv.DictReader(tsvfile, dialect='excel-tab')

            # Store data in a temporary DICTIONARY
            for row in reader:
                keys = list(row.items())
                cont_DESC = keys[0][1]
                # "cont_data" : <str>
                if "this application is a continuation of" in cont_DESC.lower():
                    cont_pull = True
                    return cont_pull, missing_list
                elif "this application is a continuational of" in cont_DESC.lower():
                    cont_pull = True
                    return cont_pull, missing_list
                elif "this application is a continuation in part of" in cont_DESC.lower():
                    cont_pull = True
                    return cont_pull, missing_list
                elif "this application is a divisional of" in cont_DESC.lower():
                    cont_pull = True
                    return cont_pull, missing_list
                elif "this application is a division of" in cont_DESC.lower():
                    cont_pull = True
                    return cont_pull, missing_list

    return cont_pull, missing_list

def patentsview_BULK_info(targets):
    # Will ananlyze Bulk claim PatensView file for info on our patent claims
    fpath = os.path.join(in_dir,"claim_data",claim_file)
    sbs1 = re.compile("[a-z]:[a-z]"); sbs2 = re.compile("[a-z];[a-z]"); sbs3 = re.compile("[a-z],[a-z]")
    clsb1 = re.compile("(?:of|in|from|to) claim [0-9]"); clsb2 = re.compile("(?:of|in|from) claims [0-9]")

    #tot_claims = 0
    #for targ in targets:
    #    keyss = iprclaim_data[targ]
    #    for ky in keyss:
    #        if ky.isdigit() == True:
    #           tot_claims += 1 
    #print(tot_claims)

    with open(fpath, encoding = "utf8", errors = "ignore") as tsvfile:
        tsvreader = csv.reader(tsvfile, delimiter="\t", )
        count = 0
        for line in tsvreader:
            # This line contains info on one of our patent claims
            if line[1] in targets and line[4] in iprclaim_data[line[1]]:
                # Clenup the text line for a more accurate word count
                textin = line[2]
                textin = textin.replace("  ", " ")
                textin = textin.replace("\n", " ")
                found1 = re.findall(sbs1,textin); found2 = re.findall(sbs2,textin); found3 = re.findall(sbs3,textin)
                for f1 in found1:
                    textin = textin.replace(f1,f1[0:2] + " " + f1[2])
                for f2 in found2:
                    textin = textin.replace(f2,f2[0:2] + " " + f2[2])
                for f3 in found3:
                    textin = textin.replace(f3,f3[0:2] + " " + f3[2])

                iprclaim_data[line[1]][line[4]]["text"] = textin
                clref1 = re.findall(clsb1,textin); clref2 = re.findall(clsb2,textin)
                if line[3] == "" and len(clref1) == 0 and len(clref2) == 0: 
                    deper = "independent"; root = None
                else: deper = "dependent"; root = line[3]
                iprclaim_data[line[1]][line[4]]["cat."] = deper
                iprclaim_data[line[1]][line[4]]["root"] = root
                iprclaim_data[line[1]][line[4]]["word_cnt"] = len(textin.split(" "))

                # If we are on the first claim, set first_claim values
                if line[4] == "1":
                    iprclaim_data[line[1]]["first_claim"]["text"] = textin
                    iprclaim_data[line[1]]["first_claim"]["cat."] = deper
                    iprclaim_data[line[1]]["first_claim"]["root"] = root
                    iprclaim_data[line[1]]["first_claim"]["word_cnt"] = len(textin.split(" "))

                #tot_claims -= 1
                print(count)
                #print(tot_claims)
                #print(textin)
                #print(len(textin.split(" ")))
                #print(deper)
                #print(root)
                #print("")

            count += 1

def claim_length_reduction_KUHN(targets):
    # Will ananlyze John Kuhn's claim reduction file for info on our patent claims
    fpath = os.path.join(in_dir,"claim_data",claimreduc_file)

    #tot_pats = len(targets)

    with open(fpath, encoding = "utf8") as tsvfile:
        tsvreader = csv.reader(tsvfile, delimiter=",")
        count = 0
        for line in tsvreader:
            # This line contains info on one of our patents
            if line[0] in targets:
                iprclaim_data[line[0]]["first_claim"]["word_change"] = int(line[2]) - int(line[1])

                #tot_pats -= 1
                #print(count)
                #print(tot_pats)
                #print(int(line[2]) - int(line[1]))
                #print("")

            count += 1

def pat_type_check(list_of_nums):
    # Helper function for screening out list of patents with PAT and not RE, D, etc.
    first3 = np.asarray([el[0:3] for el in list_of_nums[:]])
    for typef in first3:
        if typef != "PAT": return False
    return True

def write_ipr_data(data_in, data_in2, keys, keysnew, keys2, file_out):
    # Delete existing file if it exists
    if os.path.isfile(file_out) == True:
        os.remove(file_out)     

    # Open a workbook and our first worksheet for prior data
    workbook = xlsxwriter.Workbook(file_out)
    worksheet = workbook.add_worksheet(name = "IPR Trial Data")
    worksheet.set_column(0,0,20)
    worksheet.set_column(2,2,15)
    worksheet.set_column(5,6,12)
    worksheet.set_column(8,9,40)
    worksheet.set_column(4,4,20)
    worksheet.set_column(3,3,8)
    worksheet.set_column(7,7,8)
    worksheet.set_column(11,11,20)

    # Set formats
    header_format = workbook.add_format({'bold': True, 'font_color': 'black', 'align' : 'center'})
    header_format.set_text_wrap()
    text_format = workbook.add_format({'font_color': 'black', 'align' : 'vcenter'})
    text_format.set_text_wrap()
    text_format.set_align('center')
    text_format_sm = workbook.add_format({'font_color': 'black', 'align' : 'vcenter'})
    text_format_sm.set_text_wrap()
    text_format_sm.set_align('center')
    text_format_sm.set_font_size(10)
    date1_format = workbook.add_format({'font_color': 'black', 'num_format':'yyyy/mm/dd', 'align' : 'vcenter'})
    date1_format.set_text_wrap()
    date1_format.set_align('center')
    date2_format = workbook.add_format({'font_color': 'black', 'num_format':'mm/dd/yyyy', 'align' : 'vcenter'})
    date2_format.set_text_wrap()
    date2_format.set_align('center')
    int_format = workbook.add_format({'font_color': 'black', 'num_format':'0', 'align' : 'vcenter'})
    int_format.set_text_wrap() 
    int_format.set_align('center')
    float_format = workbook.add_format({'font_color': 'black', 'num_format':'0.00', 'align' : 'vcenter'})
    float_format.set_text_wrap()
    float_format.set_align('center')
    special_format = workbook.add_format({'font_color': 'red', 'align' : 'vcenter'})
    special_format.set_text_wrap()
    special_format.set_align('center')
    
    # Write headers for our first worksheet:
    worksheet.write('A1', 'Trial Number(s)', header_format)
    worksheet.write('B1', 'Trial Type', header_format)
    worksheet.write('C1', 'Trial Date', header_format)
    worksheet.write('D1', 'Final Written Decision?', header_format)
    worksheet.write('E1', 'Decision Type(s)', header_format)
    worksheet.write('F1', 'Associated Patent(s)', header_format)
    worksheet.write('G1', 'Associated Patent Type(s)', header_format)
    worksheet.write('H1', 'Multiple Patents?', header_format)
    worksheet.write('I1', 'Petitioner Name(s)', header_format)
    worksheet.write('J1', 'Patent Holder Name(s)', header_format)
    worksheet.write('K1', 'Issues?', header_format)
    worksheet.write('L1', 'Filename', header_format)

    # Write in data for each ipr on our first worksheet:
    bin_trans = {True: "Yes", False: "No", None: "No"}
    bin_trans_rev = {True: "No", False: "Yes", None: "Yes"}
    row = 1
    for key in keys:
        if data_in[key]["trial_num(s)"] is not None:
            worksheet.write_string(row,0,"\n".join(data_in[key]["trial_num(s)"]),text_format)
        if data_in[key]["trial_type"] is not None:
            worksheet.write_string(row,1,data_in[key]["trial_type"],text_format)
        if data_in[key]["dec_date"] is not None:
            worksheet.write_string(row,2,data_in[key]["dec_date"],text_format)
        worksheet.write_string(row,3,bin_trans[data_in[key]["FWD?"]],text_format)
        if data_in[key]["fd_type(s)"] is not None:
            worksheet.write_string(row,4,data_in[key]["fd_type(s)"],text_format)
        if data_in[key]["pat_num(s)"] is not None:
            worksheet.write_string(row,5,"\n".join(data_in[key]["pat_num(s)"]),text_format)
        if data_in[key]["pat_type(s)"] is not None:
            worksheet.write_string(row,6,"\n".join(data_in[key]["pat_type(s)"]),text_format)
        worksheet.write_string(row,7,bin_trans[data_in[key]["mult_pat"]],text_format)
        if data_in[key]["pet_name(s)"] is not None:
            worksheet.write_string(row,8,"\n".join(data_in[key]["pet_name(s)"]),text_format_sm)
        if data_in[key]["ph_name(s)"] is not None:
            worksheet.write_string(row,9,"\n".join(data_in[key]["ph_name(s)"]),text_format_sm)
        worksheet.write_string(row,10,bin_trans_rev[data_in[key]["no_issues"]],text_format)
        worksheet.write_string(row,11,key,text_format_sm)
        row += 1

    # Open a workbook and our first worksheet for prior data
    worksheet2 = workbook.add_worksheet(name = "IPR Patent Data")
    worksheet2.set_column(0,0,20)
    worksheet2.set_column(1,1,8)
    worksheet2.set_column(2,2,15)
    worksheet2.set_column(3,3,20)
    worksheet2.set_column(4,4,40)
    worksheet2.set_column(5,5,60)
    worksheet2.set_column(6,7,12)
    worksheet2.set_column(8,8,30)
    worksheet2.set_column(9,9,12)
    worksheet2.set_column(13,13,15)

    # Write headers for our second worksheet:
    worksheet2.write('A1', 'Trial Number(s)', header_format)
    worksheet2.write('B1', 'Trial Type', header_format)
    worksheet2.write('C1', 'Trial Date', header_format)
    worksheet2.write('D1', 'Decision Type(s)', header_format)
    worksheet2.write('E1', 'Trial Patent Holder Name(s)', header_format)  
    worksheet2.write('F1', 'Order Text', header_format)  
    worksheet2.write('G1', 'Associated Patent', header_format)
    worksheet2.write('H1', 'Associated Patent Type', header_format) 
    worksheet2.write('I1', 'Affected Claim(s)', header_format)
    worksheet2.write('J1', 'Claim(s) Disposition', header_format)
    worksheet2.write('K1', 'Order Issues?', header_format)
    worksheet2.write('L1', 'Claim Disp. Issues?', header_format)
    worksheet2.write('M1', 'Multi-Page Order?', header_format)
    worksheet2.write('N1', 'Filename', header_format)

    # Write in data for each ipr on our second worksheet:
    row = 1
    for key in keys2:
        if data_in[key]["trial_num(s)"] is not None:
            worksheet2.write_string(row,0,"\n".join(data_in[key]["trial_num(s)"]),text_format)
        if data_in[key]["trial_type"] is not None:
            worksheet2.write_string(row,1,data_in[key]["trial_type"],text_format)
        if data_in[key]["dec_date"] is not None:
            worksheet2.write_string(row,2,data_in[key]["dec_date"],text_format)
        if data_in[key]["fd_type(s)"] is not None:
            worksheet2.write_string(row,3,data_in[key]["fd_type(s)"],text_format)
        if data_in[key]["ph_name(s)"] is not None:
            worksheet2.write_string(row,4,"\n".join(data_in[key]["ph_name(s)"]),text_format_sm)
        if data_in[key]["order_txt"] is not None:
            worksheet2.write_string(row,5,data_in[key]["order_txt"],text_format)
        row_old = row
        i = 0
        #print(data_in[key]["pat_num(s)"])
        #print(data_in[key]["order_disp(s)"])
        #print(key)
        for pat_number in data_in[key]["pat_num(s)"]:
            row += 1
            worksheet2.write_string(row,6,pat_number,text_format)
            worksheet2.write_string(row,7,data_in[key]["pat_type(s)"][i],text_format)
            if pat_number in data_in[key]["order_disp(s)"]:
                if data_in[key]["order_disp(s)"][pat_number]["c-range"] is not None:
                    worksheet2.write_string(row,8,"\n".join(data_in[key]["order_disp(s)"][pat_number]["c-range"]),text_format)
                if data_in[key]["order_disp(s)"][pat_number]["disposition"] is not None:
                    worksheet2.write_string(row,9,"\n".join(data_in[key]["order_disp(s)"][pat_number]["disposition"]),text_format)
            worksheet2.write_string(row,10,bin_trans_rev[data_in[key]["no_issues2"]],text_format)
            worksheet2.write_string(row,11,bin_trans[data_in[key]["expect_ccd"]],text_format)
            worksheet2.write_string(row,12,bin_trans[data_in[key]["new_page?"]],text_format)
            i+=1

        worksheet2.write_string(row_old,13,key,text_format_sm)
        row +=1

    # Open a workbook and our first worksheet for prior data
    worksheet3 = workbook.add_worksheet(name = "IPR Claim Data")
    worksheet3.set_column(0,0,10)
    worksheet3.set_column(1,1,40)
    worksheet3.set_column(2,2,12)
    worksheet3.set_column(3,3,8)
    worksheet3.set_column(4,4,10)
    worksheet3.set_column(5,7,8)
    worksheet3.set_column(8,8,15)
    worksheet3.set_column(9,9,30)
    worksheet3.set_column(10,12,8)
    worksheet3.set_column(13,15,15)
    worksheet3.set_column(16,16,20)

    # Write headers for our second worksheet:
    worksheet3.write('A1', 'Patent No.', header_format)
    worksheet3.write('B1', 'Trial Patent Holder Name(s)', header_format)
    worksheet3.write('C1', 'Application No.', header_format)
    worksheet3.write('D1', 'Continuation?', header_format)
    worksheet3.write('E1', 'Patent Type', header_format)  
    worksheet3.write('F1', 'No. of Claims', header_format) 
    worksheet3.write('G1', 'Patent Issues?', header_format)
    worksheet3.write('H1', 'Claim Issues?', header_format)
    worksheet3.write('I1', 'NBER Category', header_format)  
    worksheet3.write('J1', 'NBER Subcategory', header_format)
    worksheet3.write('K1', 'Claim No.', header_format)
    worksheet3.write('L1', 'Word Count', header_format)
    worksheet3.write('M1', 'Word Reduc.', header_format)
    worksheet3.write('N1', 'Dependence', header_format)
    worksheet3.write('O1', 'Disposition', header_format)
    worksheet3.write('P1', 'Disp. Date', header_format)
    worksheet3.write('Q1', 'Filename(s)', header_format)

    # Write in data for each ipr patent on our third worksheet:
    row = 1
    for key in keysnew:
        worksheet3.write_string(row,0,key,text_format)
        if data_in2[key]["ph_name(s)"] is not None:
            worksheet3.write_string(row,1,"\n".join(";".join(sl) for sl in data_in2[key]["ph_name(s)"]),text_format_sm)
        if data_in2[key]["app_num"] is not None:
            worksheet3.write_string(row,2,data_in2[key]["app_num"],text_format)
        if data_in2[key]["continuation?"] is not None:
            worksheet3.write_string(row,3,bin_trans[data_in2[key]["continuation?"]],text_format)
        if data_in2[key]["pat_type"] is not None:
            worksheet3.write_string(row,4,data_in2[key]["pat_type"],text_format)
        if data_in2[key]["num_claims"] is not None:
            worksheet3.write_string(row,5,data_in2[key]["num_claims"],text_format)
        if data_in2[key]["no_issues3"] is not None:
            worksheet3.write_string(row,6,bin_trans_rev[data_in2[key]["no_issues3"]],text_format)
        if data_in2[key]["claim_msm?"] is not None:
            worksheet3.write_string(row,7,bin_trans[data_in2[key]["claim_msm?"]],text_format)
        if data_in2[key]["NBER_cats"]["cat_name"] is not None:
            worksheet3.write_string(row,8,data_in2[key]["NBER_cats"]["cat_name"],text_format)
        if data_in2[key]["NBER_cats"]["subcat_name"] is not None:
            worksheet3.write_string(row,9,data_in2[key]["NBER_cats"]["subcat_name"],text_format)

        row_old = row

        # First, compile list of claim numbers
        clist = []
        for keysub in data_in2[key]:
            if keysub.isdigit() == True: clist.append(keysub)

        # Print out first_claim information
        row += 1
        if data_in2[key]["ph_name(s)"] is not None:
            worksheet3.write_string(row,1,"\n".join(";".join(sl) for sl in data_in2[key]["ph_name(s)"]),text_format_sm)
        if data_in2[key]["app_num"] is not None:
            worksheet3.write_string(row,2,data_in2[key]["app_num"],text_format)
        if data_in2[key]["continuation?"] is not None:
            worksheet3.write_string(row,3,bin_trans[data_in2[key]["continuation?"]],text_format)
        if data_in2[key]["pat_type"] is not None:
            worksheet3.write_string(row,4,data_in2[key]["pat_type"],text_format)
        if data_in2[key]["num_claims"] is not None:
            worksheet3.write_string(row,5,data_in2[key]["num_claims"],text_format)
        if data_in2[key]["no_issues3"] is not None:
            worksheet3.write_string(row,6,bin_trans_rev[data_in2[key]["no_issues3"]],text_format)
        if data_in2[key]["claim_msm?"] is not None:
            worksheet3.write_string(row,7,bin_trans[data_in2[key]["claim_msm?"]],text_format)
        if data_in2[key]["NBER_cats"]["cat_name"] is not None:
            worksheet3.write_string(row,8,data_in2[key]["NBER_cats"]["cat_name"],text_format)
        if data_in2[key]["NBER_cats"]["subcat_name"] is not None:
            worksheet3.write_string(row,9,data_in2[key]["NBER_cats"]["subcat_name"],text_format)
        worksheet3.write_string(row,10,"first claim",text_format)
        if data_in2[key]["first_claim"]["word_cnt"] is not None:
            worksheet3.write_string(row,11,str(data_in2[key]["first_claim"]["word_cnt"]),text_format)   
        if data_in2[key]["first_claim"]["word_change"] is not None:
            worksheet3.write_string(row,13,str(data_in2[key]["first_claim"]["word_change"]),text_format)  
        if data_in2[key]["first_claim"]["cat."] is not None:
            worksheet3.write_string(row,13,data_in2[key]["first_claim"]["cat."],text_format)  
        worksheet3.write_string(row,14,"-",text_format) 
        worksheet3.write_string(row,15,"-",text_format)

        # Print out information for the rest of the claims
        for claim_number in clist:
            row += 1
            if data_in2[key]["ph_name(s)"] is not None:
                worksheet3.write_string(row,1,"\n".join(";".join(sl) for sl in data_in2[key]["ph_name(s)"]),text_format_sm)
            if data_in2[key]["app_num"] is not None:
                worksheet3.write_string(row,2,data_in2[key]["app_num"],text_format)
            if data_in2[key]["continuation?"] is not None:
                worksheet3.write_string(row,3,bin_trans[data_in2[key]["continuation?"]],text_format)
            if data_in2[key]["pat_type"] is not None:
                worksheet3.write_string(row,4,data_in2[key]["pat_type"],text_format)
            if data_in2[key]["num_claims"] is not None:
                worksheet3.write_string(row,5,data_in2[key]["num_claims"],text_format)
            if data_in2[key]["no_issues3"] is not None:
                worksheet3.write_string(row,6,bin_trans_rev[data_in2[key]["no_issues3"]],text_format)
            if data_in2[key]["claim_msm?"] is not None:
                worksheet3.write_string(row,7,bin_trans[data_in2[key]["claim_msm?"]],text_format)
            if data_in2[key]["NBER_cats"]["cat_name"] is not None:
                worksheet3.write_string(row,8,data_in2[key]["NBER_cats"]["cat_name"],text_format)
            if data_in2[key]["NBER_cats"]["subcat_name"] is not None:
                worksheet3.write_string(row,9,data_in2[key]["NBER_cats"]["subcat_name"],text_format)
            worksheet3.write_string(row,10,claim_number,text_format)
            if data_in2[key][claim_number]["word_cnt"] is not None:
                worksheet3.write_string(row,11,str(data_in2[key][claim_number]["word_cnt"]),text_format)       
            worksheet3.write_string(row,12,"-",text_format)        
            if data_in2[key][claim_number]["cat."] is not None:
                worksheet3.write_string(row,13,data_in2[key][claim_number]["cat."],text_format)   
            if data_in2[key][claim_number]["dispo"] is not None:
                print(data_in2[key][claim_number]["dispo"])
                worksheet3.write_string(row,14,data_in2[key][claim_number]["dispo"],text_format) 
            if data_in2[key][claim_number]["date"] is not None:
                worksheet3.write_string(row,15,data_in2[key][claim_number]["date"],text_format)

        worksheet3.write_string(row_old,16,"\n".join(data_in2[key]["filename(s)"]),text_format_sm)
        row +=1

    workbook.close()

def main():
    # Step 1: Read contents of existing ipr_read_data+ excel file and save to dictionary format
    pull_iprdata_ff(out_file2)

    # Step 2: Transfer parts of ipr_data into iprclaim_data
    # Old filtering conditions for printing
    tad = ipr_data.copy()
    # Must meet 4 conditions: 1. trial type is IPR or Mult. 2. At least one PAT- 3. FWD? is True
    tad = {k1: v1 for k1, v1 in tad.items() if (v1["FWD?"] == True and pat_type_check(v1["pat_type(s)"]) == True and 
            (v1["trial_type"] == "IPR" or v1["trial_type"] == "Mult." ))}
    tg_old = tad.keys()
    transfer_ipr_data(tg_old)

    # Analyze Patent and Claim Information
    target_pats = iprclaim_data.keys(); err_downs = []
    for target_pat in target_pats:
        # Step 3: Pull NBER category and subcategories, and number of claims from PatentsView
        NBER_cat, num_of_claims, app_num, error_free = patentsview_API_info(target_pat)
        iprclaim_data[target_pat]["NBER_cats"] = NBER_cat
        iprclaim_data[target_pat]["num_claims"] = num_of_claims
        iprclaim_data[target_pat]["app_num"] = app_num
        iprclaim_data[target_pat]["no_issues3"] = bool(iprclaim_data[target_pat]["no_issues3"] * error_free)

        print(target_pat)
        #print(NBER_cat)
        #print(num_of_claims)
        #if error_free == False: print("------ERROR------")
 
        # Step 4: Look up continuity data for each targeted patent (possible download from google archives)
        cont_flag, err_list = continuity_check(target_pat)
        iprclaim_data[target_pat]["continuation?"] = cont_flag
        err_downs.extend(err_list)
        
        #print(cont_flag)

    # Stop execution if we don't have complete continuity data files
    if len(err_downs) != 0:
        sys.exit("There are missing application continuity data files")
        for err in err_downs:
            print(err)

    # Step 5: Pull claim text and information from PatensView bulk data file
    patentsview_BULK_info(target_pats)

    # Step 6: Pull first claim length reduction from John Kuhn data set
    error_free = claim_length_reduction_KUHN(target_pats)

    # Step 7: Write all of our data to the output file
    write_ipr_data(ipr_data, iprclaim_data, ipr_data.keys(), iprclaim_data.keys(), tg_old, out_file3)

if __name__ == "__main__":
    main()