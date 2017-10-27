"""Reading IPR final decision claim disposition"""
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

in_dir = "in_data"
#fold = "all_iprs"
fold = "test_docs"
out_file = "ipr_read_data.xlsx"
out_file2 = "ipr_read_data+.xlsx"
temp_dir = "C:\\Users\\Johnny\\AppData\\Local\\Temp"

res = 400

# Declare data structure
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
### Finding the following:
#       "order_txt"     : "ORDERED that the joint motion to terminate the proceeding is GRANTED and . . ."
#       "order_disp(s)" : {"6658464": {"c-range": ["1-9,14"] , "disposition": ["unpatentable"]}}
#       "new_page?"     : False or True (default is False)
#       "expect_ccd"    : False or True (default is False)
#       "no_issues2"    : True or False (default is True)

def create_dictionary_entry(fname):
    ipr_data[fname] = {
        "trial_num(s)": [], "trial_type": None, "fd_type(s)": None, "mult_pat": False,
        "dec_date": None, "pet_name(s)": [], "ph_name(s)": [], "pat_num(s)": None,
        "pat_type(s)": [], "order_txt": None, "order_disp(s)": {}, "no_issues": True, "FWD?": False,
        "no_issues2": True, "expect_ccd": True
    }

def pull_iprdata_ff(file_out, dir, fold):  
    # Read ipr data from excel file
    if os.path.isfile(file_out) == True:
        # Read existing file
        df = pd.read_excel(file_out, sheetname='Sheet1')
        file_list_raw = df["Filename"].tolist()
        to_add_all = df.iloc[0:11]
        to_add_all = to_add_all.values.tolist()
    else:
        sys.exit("Error: ipr_read_data.xlsx DOES NOT EXIST")

    # Make sure each ipr in our data file is also in our data folder
    file_list = os.listdir(os.path.join(dir,fold))
    for file in file_list:
        if file not in file_list_raw:
            sys.exit("Error: file source and filenames in ipr_read_data.xlsx MISMATCH")

    # Create a filename entry in ipr_data for each filename and pull data from file
    bin_trans = {"Yes": True, "No": False}
    bin_trans_rev = {"No": True, "Yes": False}
    keys = file_list_raw
    for key, data in zip(keys,to_add_all):
        create_dictionary_entry(key)

        trial_nums = data[0].split("\n")
        if trial_nums != [""]:
            ipr_data[key]["trial_num(s)"] = trial_nums

        trial_type = data[1]
        if trial_type != "": 
            ipr_data[key]["trial_type"] = trial_type

        dec_date = data[2]
        if dec_date != "": 
            ipr_data[key]["dec_date"] = dec_date

        FWD = data[3]
        if FWD != "": 
            ipr_data[key]["FWD?"] = bin_trans[FWD]

        fd_types = data[4]
        if fd_types != "": 
            ipr_data[key]["fd_type(s)"] = fd_types

        pat_nums = data[5].split("\n")
        if pat_nums != [""]:
            ipr_data[key]["pat_num(s)"] = pat_nums

        pat_types = data[6].split("\n")
        if pat_types != [""]:
            ipr_data[key]["pat_type(s)"] = pat_types

        mult_pat = data[7]
        if mult_pat != "": 
            ipr_data[key]["mult_pat"] = bin_trans[mult_pat]

        pet_names = data[8].split("\n")
        if pet_names != [""]:
            ipr_data[key]["pet_name(s)"] = pet_names

        ph_names = data[9].split("\n")
        if ph_names != [""]:
            ipr_data[key]["ph_name(s)"] = ph_names

        no_issues = data[10]
        if no_issues != "": 
            ipr_data[key]["no_issues"] = bin_trans_rev[no_issues]

def pat_type_check(list_of_nums):
    # Helper function for screening out list of patents with PAT and not RE, D, etc.
    first3 = np.asarray([el[0:3] for el in list_of_nums[:]])
    for typef in first3:
        if typef != "PAT": return False
    return True

def pull_x_pages(subpath, fname, pages = -4):
    # Function for converting the last x pages of a document into pdf
    # Helper function for pooling ocr processing
    def pooled_tesseract_ocr(imagein):
        imagein = imagein.convert('L')
        imagein = imagein.filter(ImageFilter.SHARPEN)
        tessdata_dir_config = '--tessdata-dir "C:\\Program Files (x86)\\Tesseract-OCR\\tessdata" -oem 2 -psm 11'
        text_read = pytesseract.image_to_string(imagein, boxes = False, config=tessdata_dir_config)
        return text_read

    # First, find the number of pages in the pdf
    filestart = os.path.join(subpath, fname)
    filestartn = filestart.replace("\\","/")
    file = open(filestartn,'rb')
    reader = PyPDF2.PdfFileReader(file)
    lastpage = reader.getNumPages() - 1
    file.close()

    # Revise page range based on the number of pages in the doc
    if pages < 0 and pages*-1 > lastpage+1:
        pages = (lastpage+1)*-1
    if pages > 0 and pages > lastpage+1:
        pages = lastpage+1

    # Put pages to convert in a list
    if pages < 0:
        num = pages*-1
        subtr = np.array(list(range(0,num)))
        finalp = np.multiply(np.ones(np.size(subtr)),lastpage).astype(int)
        pages = finalp - subtr[::-1]
    elif pages > 0:
        pages = np.array(list(range(0,pages)))  

    # Convert x pages into png image
    text_out = []
    for page in pages:
        with IMG(filename = filestart + "["+str(page)+"]", resolution=res) as imgs:
            imgs.compression_quality = 99
            with imgs.sequence[0] as img:
                img.type = 'truecolor'
                IMG(img).save(filename = "fwd" + str(page) + ".png")

    # Cleanup (must be in admin mode)
    file_dump = os.listdir(temp_dir)
    for filed in file_dump:
        if "magick" in filed:
            try: os.remove(os.path.join(temp_dir,filed))
            except PermissionError:
                continue  

    # Read text using tesseract OCR
    image_vect = []
    for page in pages:
        image_vect.append(Image.open("fwd" + str(page) + ".png"))

    pool = ThreadPool(5)
    results = pool.map(pooled_tesseract_ocr, image_vect)
    i = 0
    for result in results:
        text_out.append(result)
        image_vect[i].close()
        os.remove("fwd" + str(pages[i]) + ".png")
        i += 1
    pool.close()
    pool.join()

    ## Convert x pages into png image
    #text_out = []
    #for page in pages:
    #    with IMG(filename = filestart + "["+str(page)+"]", resolution=res) as imgs:
    #        imgs.compression_quality = 99
    #        with imgs.sequence[0] as img:
    #            img.type = 'truecolor'
    #            IMG(img).save(filename = "fwd" + str(page) + ".png")

    #    # Read text using tesseract OCR
    #    imagein = Image.open("fwd" + str(page) + ".png")
    #    imagein = imagein.convert('L')
    #    imagein = imagein.filter(ImageFilter.SHARPEN)
    #    tessdata_dir_config = '--tessdata-dir "C:\\Program Files (x86)\\Tesseract-OCR\\tessdata" -oem 2 -psm 11'
    #    text_read = pytesseract.image_to_string(imagein, boxes = False, config=tessdata_dir_config)

    #    text_out.append(text_read)

    #    # Cleanup (must be in admin mode)
    #    os.remove("fwd" + str(page) + ".png")
    #    file_dump = os.listdir(temp_dir)
    #    for filed in file_dump:
    #        if "magick" in filed:
    #            try: os.remove(os.path.join(temp_dir,filed))
    #            except PermissionError:
    #                continue

    return text_out

def cleanup_text(text_list):
    revise = collections.OrderedDict()
    revise["0RDER"] = "ORDER"
    revise["oRDER"] = "ORDER"
    revise["C0NCLUSION"] = "CONCLUSION"
    revise["C0NCLUSI0N"] = "CONCLUSION"
    revise["CONCLUSI0N"] = "CONCLUSION"
    for i in range(0,len(text_list)):
        for entry in revise:
            text_list[i] = text_list[i].replace(entry, revise[entry])
    return text_list

def order_extract(text_list):
    def find_semicolons(texter,st):
        scs1 = text.find(";\n",st)
        scs2 = text.find("; and\n",st)
        scs3 = text.find(";and\n",st)

        if scs1 > 0 and scs2 > 0 and scs3 > 0:
            return min(scs1,scs2,scs3)
        elif scs1 > 0 and scs2 > 0:
            return min(scs1,scs2)
        elif scs1 > 0 and scs3 > 0:
            return min(scs1,scs3)
        elif scs2 > 0 and scs3 > 0:
            return min(scs2, scs3)
        elif scs1 == -1 and scs2 == -1 and scs3 == -1:
            return -1
        else:
            return max(scs1,scs2,scs3)

    # Start by cleaning out headers from each text string
    match = SequenceMatcher(None, text_list[0], text_list[1]).find_longest_match(0, len(text_list[0]), 0, len(text_list[1]))
    if match.a <= 1 and match.b <= 1 and match.size > 9:
        match_text = text_list[0][match.a: match.a + match.size]
        if match_text[-2:] == "/n/n": match_text = match_text[:-1]
        for g in range(0,len(text_list)):
            text_list[g] = text_list[g].replace(match_text,"")

    error_flag = True
    rev_list = []

    # First, search for CONCLUSION
    i = 0
    for text in text_list:
        if text.find("CONCLUSION") != -1:
            startl = i
            startpos = text.find("CONCLUSION")
            i = -1
            break
        i += 1
    # If CONCLUSION does not occur, then find ORDER
    if i != -1:
        i = 0
        for text in text_list:
            if text.find("ORDER") != -1:
                startl = i
                startpos = text.find("ORDER")
                i = -1
                break
    # Otherwise, search for ORDER starting at CONCLUSION position
    else:
        i = startl
        for text in text_list[startl:]:
            if i == startl:
                if text.find("ORDER",startpos) != -1:
                    startl = i
                    startpos = text.find("ORDER",startpos)
                    i = -1
                    break
            else:
                if text.find("ORDER") != -1:
                    startl = i
                    startpos = text.find("ORDER")
                    i = -1
                    break
            i += 1
    # If CONCLUSION and ORDER do not exist, then return empty and error
    if i != -1:
        return "", [], False, False
    else:
        text_list = text_list[startl:]
        text_list[0] = text_list[0][startpos:]

    # Cleanup start of "ORDER"
    if text_list[0].find("ORDER\n") == 0:
        text_list[0] = text_list[0][6:]
    elif text_list[0].find("ORDERED") == 0:
        pass
    else:
        pass
    # Cleanup line endings after periods.
    for j in range (0,len(text_list)):
        new_text = text_list[j]
        new_text = new_text.replace("U.S.\n", "U.S. ")
        new_text = new_text.replace("US.\n", "U.S. ")
        new_text = new_text.replace("No.\n", "No. ")
        new_text = new_text.replace("U.S.C.\n", "U.S.C. ")
        new_text = new_text.replace("C.F.R.\n", "C.F.R. ")
        new_text = new_text.replace("CFR.\n", "C.F.R. ")
        new_text = new_text.replace("C.P.R.\n", "C.F.R. ")
        new_text = new_text.replace("CPR.\n", "C.F.R. ")
        if new_text[-1] == ".":
            new_text = new_text + "\n"
        text_list[j] = new_text

    # Find end sets, and determine the end of the full order
    orders = []; handler = text_list.copy()
    i = 0; newp = False; period = False; carry_over = None
    for text in handler:
        i += 1
        # Exit upper loop based on period finding
        if period == True:
            break
        # Start unique search if carry_over != None
        if carry_over is not None:
            s2per = text.find(".\n",0) # possible position of end of Order
            s2col = find_semicolons(text,0) # possible position of end of order set
            if s2per == -1 and s2col == -1:
                # If we are on the last page, then append everything
                if i == len(handler):
                    orders.append(carry_over + text)
                    error_flag = False
                    break
                else:
                    carry_over = carry_over + text
                    error_flag = False
                    continue
            # "ORDER" followed by semicolon, but not period
            elif s2per == -1 and s2col != -1:
                orders.append(carry_over + text[0:s2col])
                carry_over = None
                breaker = 0; s2 = s2col
            # "ORDER" followed by period, but no semicolon
            elif s2per != -1 and s2col == -1:
                orders.append(carry_over + text[0:s2per])
                period = True
                continue
            # "ORDER" followed by period and semicolon
            else:
                if s2per < s2col:
                    orders.append(carry_over + text[0:s2per])
                    period = True
                    continue        
                else:
                    orders.append(carry_over + text[0:s2col])
                    carry_over = None
                    breaker = 0; s2 = s2col        
        else:
            breaker = 0
            s2 = 0

        while breaker == 0:
            s1 = text.find("ORDER",s2)
            if s1 != -1:
                s2per = text.find(".\n",s1) # possible position of end of Order
                s2col = find_semicolons(text,s1) # possible position of end of order set
                # "ORDER" exists, but no period or semicolon; continue to next page
                if s2per == -1 and s2col == -1:
                    # If we are on the last page, then append everything
                    if i == len(handler):
                        orders.append(text[s1:])
                        error_flag = False
                    else:
                        carry_over = text[s1:]
                    breaker = 1
                    newp = True
                    continue
                # "ORDER" followed by semicolon, but not period
                elif s2per == -1 and s2col != -1:
                    newp = True
                    orders.append(text[s1:s2col])
                    breaker = 0; s2 = s2col
                    continue
                # "ORDER" followed by period, but no semicolon
                elif s2per != -1 and s2col == -1:
                    orders.append(text[s1:s2per])
                    breaker = 1
                    period = True
                    continue
                # "ORDER" followed by period and semicolon
                else:
                    if s2per < s2col:
                        orders.append(text[s1:s2per])
                        breaker = 1
                        period = True
                        continue        
                    else:
                        orders.append(text[s1:s2col])
                        breaker = 0; s2 = s2col
                        continue           
            else:
                breaker = 1

    fulltxt = "; ".join(orders)
    fulltxt = fulltxt.replace("\n"," ")

    return fulltxt, orders, error_flag, newp

def order_claim_disposition(order_set, target):
    # Helper function for cleaning up the order text
    def cleanup_order(ordin):
        # Replace line and tab characters with spaces
        ordout = ordin.replace("\n"," ")
        ordout = ordout.replace("\t"," ")
        ordout = ordout.replace("  ", " ")
        # Correct phantom spaces in between numbers (i.e. "claim 13, 1 5, 17")
        sb = re.compile("[0-9] [0-9]")
        numseplist = re.findall(sb,ordout)
        for numsep in numseplist:
            f1 = numsep[0]; f2 = numsep[2]
            ordout = ordout.replace(numsep,f1+f2)
        # Replace "—" with "-"
        ordout = ordout.replace("—","-")
        return ordout
    # Helper function for finding verb phrases in an order text
    def find_verb(phrase_list):
        for phraser in phrase_list:
            for rekey in verbs.keys():
                flist = re.findall(rekey,phraser)
                if len(flist) != 0:
                    return verbs[rekey]
        return None
    # Helper function for finding adjective phrases in an order text 
    def find_adj(phrase_list):
        for phraser in phrase_list:
            for rekey in adject.keys():
                flist = re.findall(rekey,phraser)
                if len(flist) != 0:
                    return adject[rekey]
        return None
    # Helper function for finding patent numbers in an order text
    def find_patnum(phrase_list, patent_list):
        regexlist = []; err = True
        for p in patent_list:
            if len(p) != 7:
                err = False
                regexlist.append(re.compile("oaiwjeofijiaoiwer"))
            else:
                #regexlist.append(re.compile(p[0]+"(?:,|\s)\s*"+p[1]+p[2]+p[3]+"(?:,|\s)\s*"+p[4]+p[5]+p[6]))
                regexlist.append(re.compile(p[4]+p[5]+p[6]))
        for phraser in phrase_list:
            j = 0
            for regcomp in regexlist:
                finder = re.findall(regcomp,phraser)
                if len(finder) != 0:
                    return patent_list[j], err
                j+=1
        return None, False
    # Helper function for finding claim numbers in an order text
    def strip_claim_nums(claimtxt):
        claimtxt = claimtxt.replace(",","")
        claimtxt = claimtxt.replace("and","")
        claimtxt = claimtxt.replace("  ", " ")
        txtlist = claimtxt.split(" ")
        if txtlist[0].lower() == "claim" or txtlist[0].lower() == "claims":
            txtlist.pop(0)
        numstringlst = []
        for txt in txtlist:
            if txt[0].isdigit() == True and txt[-1].isdigit() == True:
                numstringlst.append(txt)
            else:
                break
        return ",".join(numstringlst)

    # Initialize reference vars
    pat_list = ipr_data[target]["pat_num(s)"]
    mult_pat = ipr_data[target]["mult_pat"]
    error_flag = True
    ccd_error = False

    # Create dictionary of verbiage and adjectives
    verbs = collections.OrderedDict()
    verbs[re.compile("has not been shown", re.IGNORECASE)] = False
    verbs[re.compile("has been shown (?!not)", re.IGNORECASE)] = True
    verbs[re.compile("have not been shown", re.IGNORECASE)] = False
    verbs[re.compile("have been shown (?!not)", re.IGNORECASE)] = True
    verbs[re.compile("is held not", re.IGNORECASE)] = False
    verbs[re.compile("is held (?!not)", re.IGNORECASE)] = True
    verbs[re.compile("are held not", re.IGNORECASE)] = False
    verbs[re.compile("are held (?!not)", re.IGNORECASE)] = True
    verbs[re.compile("are not", re.IGNORECASE)] = False
    verbs[re.compile("is not", re.IGNORECASE)] = False
    verbs[re.compile("are (?!not)", re.IGNORECASE)] = True
    verbs[re.compile("is (?!not)", re.IGNORECASE)] = True
    adject = collections.OrderedDict()
    adject[re.compile("unpatentable", re.IGNORECASE)] = False
    adject[re.compile("(?!un)patentable", re.IGNORECASE)] = True

    # Initialize the dictionary entries for our results
    dispos = collections.OrderedDict()
    for pat in pat_list:
        dispos[pat] = {"c-range": [], "disposition": []}
    
    # Loop through order set and extract claim disposition phrases
    cl = re.compile("claim(?=s| |$)",re.IGNORECASE)
    claim_order_ls = []
    for order in order_set:
        order = cleanup_order(order)
    
        # Find "claim" or "claims"
        cpos = [m.start() for m in re.finditer(cl, order)]
        cstrlist = []
        if len(cpos) == 1:
            cstrlist.append(order[cpos[0]:])
        elif len(cpos) > 1:
            for j in range(0,len(cpos)):
                if j != len(cpos) - 1:
                    cstrlist.append(order[cpos[j]:cpos[j+1]])
                else:
                    cstrlist.append(order[cpos[j]:])
        else:
            continue # no "claims" or "claim"
        claim_order_ls.append(cstrlist)

    # If we haven't found any claim language, then return nothing and an error
    if len(claim_order_ls) == 0:
        error_flag = False
        return ipr_data[target]["order_disp(s)"], False, error_flag

    # Analyze claim disposition phrases
    print(ipr_data[target]["order_txt"])
    print("")

    for phrase_set in claim_order_ls:
        for phrase in phrase_set:

            print("Mult. Pat:", mult_pat)

            # ONE PATENT CASE
            if mult_pat == False:
                # Start by finding verb triggers
                verb_logic = find_verb([phrase])
                if verb_logic is None:
                    verb_logic = find_verb(phrase_set)
                    if verb_logic is None:
                        ccd_error = True # no verb phrase at all
                        print("verb error")
                        print(phrase_set)
                        break
                # Now find the adjective trigger
                adj_logic = find_adj([phrase])
                if adj_logic is None:
                    adj_logic = find_adj(phrase_set)
                    if adj_logic is None:
                        ccd_error = True 
                        print("adj error")
                        print(phrase_set)
                        break
                # We have a a verb and adjective, now find claim numbers
                patnum = ipr_data[target]["pat_num(s)"][0]
                string_of_nums = strip_claim_nums(phrase)
                logic = (verb_logic == adj_logic)
                if logic == True: logicout = "patentable"
                else: logicout = "unpatentable"
                # Save claim dispositions to our temporary dictionary
                dispos[patnum]["c-range"].append(string_of_nums)
                dispos[patnum]["disposition"].append(logicout)

                print(logicout)
                print(string_of_nums)
                print(patnum)

            # MULTI PATENT CASE
            else:
                # Start by finding verb triggers
                verb_logic = find_verb([phrase])
                if verb_logic is None:
                    verb_logic = find_verb([phrase_set[-1]])
                    if verb_logic is None:
                        ccd_error = True # no verb phrase at all
                        print("verb error")
                        print(phrase_set)
                        break
                # Now find the adjective trigger
                adj_logic = find_adj([phrase])
                if adj_logic is None:
                    adj_logic = find_adj([phrase_set[-1]])
                    if adj_logic is None:
                        ccd_error = True 
                        print("adj error")
                        print(phrase_set)
                        break
                # Now find the patent number
                patnum, errpat = find_patnum([phrase], ipr_data[target]["pat_num(s)"])
                if patnum is None:
                    patnum, errpat = find_patnum([phrase_set[-1]], ipr_data[target]["pat_num(s)"])
                    if errpat == False: error_flag = False
                    if patnum is None:
                        ccd_error = True
                        print("patnum error")
                        print(phrase_set)
                        break
                # We have a verb, adjective, and patent, now find claim numbers
                string_of_nums = strip_claim_nums(phrase)
                logic = (verb_logic == adj_logic)
                if logic == True: logicout = "patentable"
                else: logicout = "unpatentable"
                # Save claim dispositions to our temporary dictionary
                dispos[patnum]["c-range"].append(string_of_nums)
                dispos[patnum]["disposition"].append(logicout)

                print(logicout)
                print(string_of_nums)
                print(patnum)

    return dispos, False, error_flag    

def write_ipr_data2(data_in, keys, keys2, file_out):
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
    worksheet2.set_column(8,8,15)
    worksheet2.set_column(9,9,10)
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
        for pat_number in data_in[key]["pat_num(s)"]:
            row += 1
            worksheet2.write_string(row,6,pat_number,text_format)
            worksheet2.write_string(row,7,data_in[key]["pat_type(s)"][i],text_format)
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
        
def main():
    # Step 1: Read contents of existing ipr_read_data excel file and save to dictionary format
    pull_iprdata_ff(out_file, in_dir,fold)

    # Step 2: Identify targeted IPR trials (filter 1)
    tad = ipr_data.copy()
    # Must meet 4 conditions: 1. trial type is IPR or Mult. 2. At least one PAT- 3. FWD? is True
    tad = {k1: v1 for k1, v1 in tad.items() if (v1["FWD?"] == True and pat_type_check(v1["pat_type(s)"]) == True and 
            (v1["trial_type"] == "IPR" or v1["trial_type"] == "Mult." ))}
    targets = tad.keys()

    # Step 3: Analyze claim disposition of final written decisions for targeted applications
    subpath = os.path.join(in_dir,fold)
    print_counter = 1
    for target in targets:
        print(print_counter)
        print_counter += 1

        print(target)

        # Pull text from last 4 pages of the pdf document
        text_read_list = pull_x_pages(subpath, target, -5)
        #print(text_read_list)

        # Clean up text for crucial keywords
        text_read_list = cleanup_text(text_read_list)

        # Extract full_order text; also, revise the text_read to only contain the order (but still in a list)
        full_order, revised_text_read, error_free, new_page = order_extract(text_read_list)
        ipr_data[target]["new_page?"] = new_page
        ipr_data[target]["no_issues2"] = bool(ipr_data[target]["no_issues2"] * error_free)
        ipr_data[target]["order_txt"] = full_order

        #print(full_order)
        #print(error_free)
        #print(new_page)
        #print("")

        # Analyze set of orders for claim disposition
        claim_dispo, ccd_issue, error_free = order_claim_disposition(revised_text_read, target)
        ipr_data[target]["order_disp(s)"] = claim_dispo
        ipr_data[target]["no_issues2"] = bool(ipr_data[target]["no_issues2"] * error_free)
        ipr_data[target]["expect_ccd"] = ccd_issue

    # Step 4: Write all of our data to the output file
    write_ipr_data2(ipr_data, ipr_data.keys(), targets, out_file2)

if __name__ == "__main__":
    main()