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

in_dir = "in_data"
#fold = "all_iprs"
fold = "test_docs2"
out_file = "ipr_read_data.xlsx"
out_file2 = "ipr_read_data+.xlsx"
out_file3 = "ipr_read_data++.xlsx"
temp_dir = "C:\\Users\\Johnny\\AppData\\Local\\Temp"

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
#       "continuation?" : False
#       "NBER_cats"     : {"cat_id" : "3", "subcat_id" : "32", "cat_name": "Drgs&Med", "subcat_name": "Surgery & Med Inst."}
#       "num_claims"    : 24
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
        "ph_name(s)": [], "NBER_cats": {}, "num_claims" : None, "first_claim": {}, 
        "continuation?": False, "claim_msm?" : False
    }

def enter_claim_dict_entries(pat_num, claim_nums, claim_disp, date, existing = None):
    claims = []; dispos = []
    disptxts = claim_disp
    # Compile symmetric lists of claims and their dispositions
    n = 0
    for cntxts in claim_nums:
        for cntxt in cntxts:
            cntxt = cntxt.replace(" ","")
            cls = cntxt.split(",")
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
                enter_claim_dict_entries(pat, dispos[pat]["c-range"], dispos[pat]["disposition"],date)
            # We have a repeated patent number; must check value
            else:
                iprclaim_data[pat]["trial_num(s)"].append(ipr_data[targ]["trial_num(s)"])
                iprclaim_data[pat]["filename(s)"].append(targ)
                iprclaim_data[pat]["dec_date(s)"].append(ipr_data[targ]["dec_date"])
                iprclaim_data[pat]["ph_name(s)"].append(ipr_data[targ]["ph_name(s)"])
                # If needed created dictionary entries; if not, we must compare disposition and date
                existing =  iprclaim_data[pat]
                date = ipr_data[targ]["dec_date"]
                enter_claim_dict_entries(pat, dispos[pat]["c-range"], dispos[pat]["disposition"], date, existing)
            j+=1
                    
def patentsview_API_info(pat):
    # Pull patent information from PatentsView API
    if len(pat) == 0: 
        return {}, None, False
    error_ret = True
    ret1 = "%22patent_num_claims%22"
    ret2 = "%22nber_category_id%22"
    ret3 = "%22nber_subcategory_id%22"
    ret4 = "%22nber_category_title%22"
    ret5 = "%22nber_subcategory_title%22"
    fquery = ",".join([ret1,ret2,ret3,ret4,ret5])
    url_path = "http://www.patentsview.org/api/patents/query?q={%22patent_number%22:%22"+str(pat)+"%22}&f=["+fquery+"]"
    with urllib.request.urlopen(url_path) as url:
        datadown = json.loads(url.read().decode())
        if datadown["total_patent_count"] > 1: error_ret = False
        cat_id = datadown["patents"][0]["nbers"][0]["nber_category_id"]
        scat_id = datadown["patents"][0]["nbers"][0]["nber_subcategory_id"]
        cat_title = datadown["patents"][0]["nbers"][0]["nber_category_title"]
        scat_title = datadown["patents"][0]["nbers"][0]["nber_subcategory_title"]
        num_of_claims = datadown["patents"][0]["patent_num_claims"]
        if "" in [cat_id, scat_id, cat_title, scat_title, num_of_claims]: error_ret = False
        nber_data = {"cat_id":cat_id, "subcat_id":scat_id, "cat_name":cat_title, "subcat_name":scat_title}
    return nber_data, num_of_claims, error_ret

def patentsview_BULK_info(pat):
    pass

def claim_length_reduction(pat):
    pass

def continuity_check(pat):
    pass

def pat_type_check(list_of_nums):
    # Helper function for screening out list of patents with PAT and not RE, D, etc.
    first3 = np.asarray([el[0:3] for el in list_of_nums[:]])
    for typef in first3:
        if typef != "PAT": return False
    return True

## NOT DONE #####
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
    worksheet2 = workbook.add_worksheet(name = "IPR Claim Data")

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

    target_pats = iprclaim_data.keys()
    for target_pat in target_pats:
        # Step 3: Pull NBER category and subcategories, and number of claims from PatentsView
        NBER_cat, num_of_claims, error_free = patentsview_API_info(target_pat)

        #print(target_pat)
        #print(NBER_cat)
        #print(num_of_claims)
        #if error_free == False: print("------ERROR------")

        # Step 4: Pull claim text and information from PatensView bulk data file
        claim_msm, error_free = patentsview_BULK_info(target)

        ## Step 5: Pull first claim length reduction from John Kuhn data set
        #error_free = claim_length_reduction(target)
 
        ## Step 6: Look up continuity data for each targeted patent (download from google archives)
        #cont_flag = continuity_check(target)

    breaking = 1
    # Step 7: Write all of our data to the output file
    write_ipr_data(ipr_data, iprclaim_data, ipr_data.keys(), iprclaim_data.keys(), tg_old, out_file3)

if __name__ == "__main__":
    main()