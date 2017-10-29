"""Reading IPR pdf files first page"""
# Jonathan Slack
# jslackd@gmail.com

import os
import sys
from PIL import Image, ImageEnhance, ImageFilter
import pytesseract
import collections
import re
from datetime import date
from datetime import datetime
from wand.image import Image as IMG
import os
import itertools
import xlsxwriter

in_dir = "in_data"
fold = "all_iprs"
#fold = "test_docs"
out_file = "ipr_read_data.xlsx"
temp_dir = "C:\\Users\\Johnny\\AppData\\Local\\Temp"

res = 400

# Declare data structure
ipr_data = collections.OrderedDict()
# Format: {filename.pdf (str): {attrib (str): val (---)}}
# Attributes:
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
#       "no_issues"     : False or True

def create_dictionary_entry(fname):
    ipr_data[fname] = {
        "trial_num(s)": [], "trial_type": None, "fd_type(s)": None, "mult_pat": False,
        "dec_date": None, "pet_name(s)": [], "ph_name(s)": [], "pat_num(s)": None,
        "pat_type(s)": [], "order_txt": None, "order_disp(s)": {}, "no_issues": True, "FWD?": False
    }

def read_desc_date(procstr):
    # This function will find the decision date for the final decision
    # If it fails, it will return None, False (False being for issue flag)
    
    # Start by finding outer bound based on "UNITED STATES" or "PATENT" or "before" or "patent"
    ender = None
    if procstr.find("UNITED STATES") != -1:
        pos1 = procstr.find("UNITED STATES") - 1
        if procstr[pos1] != "\n" and procstr[pos1] != " ": pos1 += 1
        ender = pos1
    elif procstr.find("PATENT") != -1:
        pos1 = procstr.find("PATENT") - 1
        if procstr[pos1] != "\n" and procstr[pos1] != " ": pos1 += 1
        ender = pos1
    elif procstr.lower().find("before") != -1:
        pos1 = procstr.lower().find("before") - 1
        if procstr[pos1] != "\n" and procstr[pos1] != " ": pos1 += 1
        ender = pos1
    elif procstr.lower().find("patent") != -1:
        pos1 = procstr.lower().find("patent") - 1
        if procstr[pos1] != "\n" and procstr[pos1] != " ": pos1 += 1
        ender = pos1
    # If we haven't set ender, choose the first half of the page
    if ender == None: ender = int(len(procstr)/2)

    # Starter is either <month> or the first "\n" character
    starter = None; mcheck = False
    months = ["January","February","March","April","May","June","July","August","September","October","November","December"]
    for month in months:
        if procstr.find(month,0,ender) != -1:
            mcheck = month
            break

    if mcheck != False:
        pos0 = procstr.find(mcheck,0,ender)
        starter = pos0
    elif procstr.find("\n",0,ender) != -1:
        pos0 = procstr.find("\n",0,ender)
        starter = pos0
    else:
        starter = 0    

    # Snip the string to only contain the date
    snip = procstr[starter:ender]
    # If we found a month, this will be easy
    if mcheck != False:
        if snip[0] == " ": snip = snip[1:] # removes extra space
        cut = snip.find("\n")
        snip = snip[0:cut].replace(" ","")
    # Otherwise, we need to manually search for date format
    else:
        splits = snip.split()
        year = None; month = None; day = None
        for i in range(2,len(splits)):
            pos_year = splits[i]; pos_day = splits[i-1].replace(",",""); pos_month = splits[i-2]
            if pos_year.isdigit() == True and pos_day.isdigit() == True and pos_month.isdigit() == False:
                year = pos_year; 
                day = pos_day; 
                month = pos_month

        if year == None or day == None or month == None:
            return None, False
        else:
            snip = month + day + "," + year

    # Cleanup snip for date processing
    if snip[0] == " " or snip[0] == "\n":
        snip = snip[1:]
    if snip[-1:] == " " or snip[-1:] == "\n":
        snip = snip[0:-1]

    # Pull date from targeted string
    try:
        ddate = datetime.strptime(snip, "%B%d,%Y")
        ddate = ddate.strftime("%m/%d/%Y")
    except ValueError:
        return snip, False

    return ddate, True

def read_petitioner_names(procstr):
    # This function will find the petitioner names on the first page of an IPR
    # If it fails, it will return [], False (False being for issue flag)

    # Find the first instance of "APPEAL BOARD" or "PATENT TRIAL AND" or "BEFORE THE"
    error_flag = True
    starter = None; ender = None
    if procstr.find("APPEAL BOARD") != -1:
        pos0 = procstr.find("APPEAL BOARD")
        pos1 = procstr[pos0::].find("\n") + pos0
        starter = pos1 + 1
    elif procstr.find("PATENT TRIAL AND") != -1:
        pos0 = procstr.find("PATENT TRIAL AND")
        pos1 = procstr[pos0::].find("\n") + pos0
        starter = pos1 + 1
    elif procstr.find("BEFORE THE") != -1:
        pos0 = procstr.find("BEFORE THE")
        pos1 = procstr[pos0::].find("\n") + pos0
        starter = pos1 + 1

    # Find the first instance of "petitioner"
    if starter is not None:
        pos2 = procstr[starter::].lower().find("petitioner") + starter
        if pos2 != -1:
            ender = pos2 - 1

    # If we haven't set ender, we have failed
    if ender is None:
        return [], False, 0

    # Cut snippet of text from "APPEAL BOARD" to "petitioner" 
    snip = procstr[starter:ender].replace("\n", " ")
    if snip[0] == " ": snip = snip[1::]

    # Split names preliminarily based on occurence of " and "
    splits = []
    and_occur = snip.find(" and ")
    if and_occur != -1:
        splits.append(snip[0:and_occur])
        splits.append(snip[and_occur+5::])
    else:
        splits.append(snip)

    # Crawl through the snip and check for separate names
    excs = ["LLC", "INC.", "CO.", "LTD.", "S.E.", "CORP.", "M.D.", "L.L.C."]
    petitioners = []
    for split in splits:
        pos = 0
        while pos < (len(split) - 1) and pos != -1:
            pos_old = pos
            pos = split.find(",", pos_old)
            if pos == -1:
                # A comma does not exist; add remaining string to list if it is not blank
                to_add = split[pos_old::]
                # Cleanup
                if to_add[0] == " ": to_add = to_add[1::]
                if to_add == "" or to_add == " ":
                    continue # exit if we have bad string
                lenner = len(to_add)
                if to_add[lenner-1] == ",": to_add = to_add[:-1]
                # Adding to list
                petitioners.append(to_add)             
            else:
                # A comma exists; add the bounded string to our list
                to_add = split[pos_old:pos]
                # Check for exceptional comma occurences; loop through possible exceptions
                exiter = False
                while exiter == False and pos < (len(split)-1):
                    manip = split[pos+1::].replace(" ","")
                    manip = manip.replace(",", "")
                    for exc in excs:
                        if manip.find(exc) == 0:
                            pos = len(exc) + split[pos::].find(exc) + pos
                            to_add = split[pos_old:pos]
                            break
                        if exc == "L.L.C.":
                            exiter = True
                if pos < (len(split)-1) and split[pos] != ",": 
                    pos_hold = split.find(",", pos)
                    if pos_hold == -1:
                        pos = len(split)
                        to_add = split[pos_old:pos]
                    else:
                        to_add = split[pos_old:pos]

                # Cleanup
                if len(to_add) == 0: error_flag = False
                if error_flag == True and to_add[0] == " ": to_add = to_add[1::]
                if error_flag == True and to_add[len(to_add)-1] == ",": to_add = to_add[:-1]
                # Adding to list
                petitioners.append(to_add)
                pos += 1
        
    # Cleanup names based on semicolons
    petitioners_new = []
    for pet in petitioners:
        clean = re.split(";",pet)
        indices = [i for i, x in enumerate(clean) if (x == "" or x == " ")]
        clean = [i for j, i in enumerate(clean) if j not in indices]
        if len(clean) > 1:
            petitioners_new.extend(clean)
        else:
            petitioners_new.append(pet)

    cnt = 0
    for pet in petitioners_new:
        if len(pet) == 0: continue
        if pet[0] == " ": 
            pet = pet[1::]
            petitioners_new[cnt] = pet
        if pet[-4:] == "CORR":
            pet = pet[:-1] + "P"
            petitioners_new[cnt] = pet
        if pet[0:5] == "J OHN":
            pet = "J" + pet[2:]
            petitioners_new[cnt] = pet
        cnt += 1

    return petitioners_new, error_flag, ender
    
def read_pholder_names(procstr_full, start_pos):
    # This function will find the patent holder names on the first page of an IPR
    # If it fails, it will return [], False (False being for issue flag)

    starter = None; ender = None
    procstr = procstr_full[start_pos::]
    # Find the first instance of "patent owner" or "patent 0wner" or "patentowner" and set as ender
    if procstr.lower().find("patent owner") != -1:
        pos1 = procstr.lower().find("patent owner") - 1
        if procstr[pos1] != "\n" and procstr[pos1] != " ": pos1 += 1
        ender = pos1
    elif procstr.lower().find("patent 0wner") != -1:
        pos1 = procstr.lower().find("patent 0wner") - 1
        if procstr[pos1] != "\n" and procstr[pos1] != " ": pos1 += 1
        ender = pos1
    elif procstr.lower().find("patentowner") != -1:
        pos1 = procstr.lower().find("patentowner") - 1
        if procstr[pos1] != "\n" and procstr[pos1] != " ": pos1 += 1
        ender = pos1

    # If we have not reset ender, we have failed
    if ender is None:
        return [], False, 0

    # Find the first instance of "v." or "V." or "v" + occuring after start_pos!=0
    if procstr.find("v.",0,ender) != -1:
        pos0 = procstr.find("v.",0,ender) + 2
        if procstr[pos0] == "\n": pos0 += 1
        starter = pos0
    elif procstr.find("V.",0,ender) != -1:
        pos0 = procstr.find("V.",0,ender) + 2
        if procstr[pos0] == "\n": pos0 += 1
        starter = pos0
    elif start_pos != 0 and procstr.lower().find("v",0,ender) != -1:
        pos0 = procstr.lower().find("v",0,ender) + 1
        if procstr[pos0] == "\n": pos0 += 1
        starter = pos0

    # If we have not reset starter, we have failed
    if starter is None:
        return [], False, 0

    # Cut snippet of text from "v." to "Patent Owner" 
    snip = procstr[starter:ender].replace("\n", " ")
    if snip[0] == " ": snip = snip[1::]

    # Split names preliminarily based on occurence of " and "
    splits = []
    and_occur = snip.find(" and ")
    if and_occur != -1:
        splits.append(snip[0:and_occur])
        splits.append(snip[and_occur+5::])
    else:
        splits.append(snip)

    # Crawl through the snip and check for separate names
    excs = ["LLC", "INC.", "CO.", "LTD.", "S.E.", "CORP.", "M.D.", "L.L.C."]
    pholders = []
    for split in splits:
        pos = 0
        while pos < (len(split) - 1) and pos != -1:
            pos_old = pos
            pos = split.find(",", pos_old)
            if pos == -1:
                # A comma does not exist; add remaining string to list if it is not blank
                to_add = split[pos_old::]
                # Cleanup
                if to_add[0] == " ": to_add = to_add[1::]
                if to_add == "" or to_add == " ":
                    continue # exit if we have bad string
                lenner = len(to_add)
                if to_add[lenner-1] == ",": to_add = to_add[:-1]
                # Adding to list
                pholders.append(to_add)             
            else:
                # A comma exists; add the bounded string to our list
                to_add = split[pos_old:pos]
                # Check for exceptional comma occurences; loop through possible exceptions
                exiter = False
                while exiter == False and pos < (len(split)-1):
                    manip = split[pos+1::].replace(" ","")
                    manip = manip.replace(",", "")
                    for exc in excs:
                        if manip.find(exc) == 0:
                            pos = len(exc) + split[pos::].find(exc) + pos
                            to_add = split[pos_old:pos]
                            break
                        if exc == "L.L.C.":
                            exiter = True
                if pos < (len(split)-1) and split[pos] != ",": 
                    pos_hold = split.find(",", pos)
                    if pos_hold == -1:
                        pos = len(split)
                        to_add = split[pos_old:pos]
                    else:
                        to_add = split[pos_old:pos]

                # Cleanup
                if to_add[0] == " ": to_add = to_add[1::]
                if to_add[len(to_add)-1] == ",": to_add = to_add[:-1]
                # Adding to list
                pholders.append(to_add)
                pos += 1
        
    # Cleanup names based on semicolons
    pholders_new = []
    for ph in pholders:
        clean = re.split(";",ph)
        indices = [i for i, x in enumerate(clean) if (x == "" or x == " ")]
        clean = [i for j, i in enumerate(clean) if j not in indices]
        if len(clean) > 1:
            pholders_new.extend(clean)
        else:
            pholders_new.append(ph)
    # Cleanup names based on spaces starting the string; "CORR" at the end; "J OHN" at the beginning
    cnt = 0
    for ph in pholders_new:
        if ph[0] == " " or ph[0] == "\n": 
            ph = ph[1::]
            pholders_new[cnt] = ph
        if ph[-4:] == "CORR":
            ph = ph[:-1] + "P"
            pholders_new[cnt] = ph
        if ph[0:5] == "J OHN":
            ph = "J" + ph[2:]
            pholders_new[cnt] = ph
        cnt += 1
    return pholders_new, True, (ender + start_pos)

def read_trial_nums(procstr_full, start_pos, target):
    # This function will find the trial numbers on the first page of an IPR
    # If it fails, it will return [], False (False being for issue flag)

    procstr = procstr_full[start_pos+1:].replace(" ","")
    error_flag = True
    # Find farthest outer bound based on "before" or "judge"
    if procstr.lower().find("before") != -1:
        f_end = procstr.lower().find("before")
    elif procstr.lower().find("judge") != -1:
        f_end = procstr.lower().find("judge")
    else:
        f_end = len(procstr) - 1

    # Find bounds for case number(s) and add save bounded snips
    snips = []
    formater = re.compile("[A-Z][A-Z][A-Z][0-9][0-9][0-9][0-9].[0-9][0-9][0-9][0-9][0-9]")
    pos0 = 0; pos1 = f_end
    while pos0 != -1 and pos0 < (len(procstr)-1):
        pos0 = procstr.lower().find("case",pos0,pos1)

        if pos0 != -1:
            pos1 = procstr.lower().find("patent",pos0,pos1)

            # WE have found proper bounds; add to snip list
            if pos1 != -1:
                snips.append(procstr[pos0+4:pos1])
                pos0 = pos1 + 1
                pos1 = f_end
            else:
                break
       
    ## No snips were found; we failed
    #if len(snips) == 0:
    #    return [], False

    # Pull case numbers from the compiled list
    # Also, compare to the existing entry (if it exists)
    trial_nums = []
    for snip in snips:
        found = re.findall(formater,snip)
        if len(found) != 0:
            trial_nums.extend(found)

    # Find any trial number occuring near the end of the document (no bounds)
    if f_end != (len(procstr) - 1):
        extra_str = procstr[f_end:]
        extra_found = re.findall(formater, extra_str)
        trial_nums.extend(extra_found)

    # Replace all "—" with "-" in the trial_nums we have found
    cnt = 0
    for trial_num in trial_nums:
        trial_num = trial_num.replace("—","-")
        trial_nums[cnt] = trial_num
        cnt += 1

    # Compare the preliminary trial number to what we have found thus far
    to_compare = ipr_data[target]["trial_num(s)"]
    if len(to_compare) != 0 and len(trial_nums) != 0:
        for comp in to_compare:
            if comp.lower() not in [tn.lower() for tn in trial_nums]:
                error_flag = False
                trial_nums.append(comp)
    elif len(to_compare) != 0 and len(trial_nums) == 0:
        error_flag = False
        for comp in to_compare:
            trial_nums.append(comp)
    elif len(to_compare) == 0 and len(trial_nums) == 0:
        error_flag = False

    return trial_nums, error_flag

def decide_trial_type(trials):
    # No data to go off of; return None for type type
    if len(trials) == 0:
        return None, False

    error_flag = True
    ttype = None
    for trial in trials:
        if trial.find("IPR") != -1:
            if ttype == None: ttype = "IPR"
            elif ttype == "IPR": pass
            else: ttype = "Mult."
        elif trial.find("CBM") != -1:
            if ttype == None: ttype = "CBM"
            elif ttype == "CBM": pass
            else: ttype = "Mult."
        elif trial.find("PGR") != -1:
            if ttype == None: ttype = "PGR"
            elif ttype == "PGR": pass
            else: ttype = "Mult."
        else:
            error_flag = False

    return ttype, error_flag

def read_pat_nums(procstr_full, start_pos):
    # This function will find the patent numbers on the first page of an IPR
    # If it fails, it will return [], [], False, False, f_end (False being for issue flag)

    procstr = procstr_full[start_pos+1:].replace(" ","")
    error_flag = True
    # Find farthest outer bound based on "before" or "judge" or "administrative"
    if procstr.lower().find("before") != -1:
        f_end = procstr.lower().find("before")
    elif procstr.lower().find("judge") != -1:
        f_end = procstr.lower().find("judge")
    elif procstr.lower().find("administrative") != -1:
        f_end = procstr.lower().find("administrative")
    else:
        f_end = len(procstr) - 1
    procstr = procstr[:f_end]

    # Find first location of "patent" w/o "patentowner"
    if procstr.lower().find("patent") != -1:
        pos0 = procstr.lower().find("patent")
        if procstr.lower().find("owner") == pos0 + 6:
            pos0 = procstr.lower().find("patent",pos0 + 6)
        snip = procstr[pos0:]
    # Two fail conditions: snip not long enough and no occurences of "patent"
    else:
        return [], [], False, False, (f_end + start_pos + 1)
    if len(snip) < 7:
        return [], [], False, False, (f_end + start_pos + 1)

    # Search for different types of patents
    patent_list = []
    pattype_list = []
    formatPAT = re.compile("[0-9].[0-9][0-9][0-9].[0-9][0-9][0-9]")
    PATposl = [m.start() for m in re.finditer(formatPAT, snip)]
    for PATpos in PATposl:
        ender = snip.find("\n",PATpos)
        patent_list.append(snip[PATpos:ender])
        pattype_list.append("PAT-")
        
    formatD = re.compile("D[0-9][0-9][0-9].[0-9][0-9][0-9]")
    Dposl = [m.start() for m in re.finditer(formatD, snip)]
    for Dpos in Dposl:
        ender = snip.find("\n",Dpos)
        patent_list.append(snip[Dpos:ender])
        pattype_list.append("D---")

    formatRE = re.compile("RE[0-9][0-9].[0-9][0-9][0-9]")
    REposl = [m.start() for m in re.finditer(formatRE, snip)]
    for REpos in REposl:
        ender = snip.find("\n",REpos)
        patent_list.append(snip[REpos:ender])
        pattype_list.append("RE--")

    if len(patent_list) == 0:
        error_flag = False

    # Cleanup patent numbers by removing all prefixes, commas, periods, or semicolons
    for i in range(0,len(patent_list)):
        cur = patent_list[i]
        # Remove prefixes
        pop = 0
        for x in cur:
            if x.isdigit() == False: pop += 1
            else: break
        cur = cur[pop:] 
        # Remove commas, periods, semicolons, parens, brackets
        cur = cur.replace(",",""); cur = cur.replace('.',""); cur = cur.replace(';',"")
        cur = cur.replace(')',""); cur = cur.replace(']',"")
        patent_list[i] = cur
        
    # Pull patent sub-type from the end of each string
    subtypes = ["A1","A2","A9","A","B1","B2","B3","B","C1","C2","C3","C","P1","P2","P3","P4","P9","S","E","D"]
    for i in range(0,len(patent_list)):
        pat = patent_list[i]
        for st in subtypes:
            loc = pat.find(st)
            if loc != -1:
                pat = pat[:loc]
                pattype_list[i] += st
                patent_list[i] = pat
                break

    if len(patent_list) > 1: mult_pat = True
    else: mult_pat = False

    return patent_list, pattype_list, mult_pat, error_flag, (f_end + start_pos + 1)

def read_dec_type(procstr_full, startpos, target):
    # This function will find the decision type(s) on the first page of an IPR
    # If it fails, it will return None, False, False (False being for issue flag)
    
    if startpos == len(procstr_full) - 1: startpos = int(len(procstr_full)/2)
    procstr = procstr_full[startpos:]
    procstr = procstr.replace("J udge", "Judge")

    # Look for a "35" or "37" preceded by a "\n"
    if procstr.find("\n35") != -1:
        pos1 = procstr.find("\n35")
    elif procstr.find("\n37") != -1:
        pos1 = procstr.find("\n37")
    elif procstr.find("§") != -1:
        ps = procstr.find("§")
        possibilites = [m.start() for m in re.finditer("\n", procstr[:ps])]
        pos1 = possibilites[-1:][0]
    elif procstr.find("U.S.C.") != -1:
        ps = procstr.find("U.S.C.")
        possibilites = [m.start() for m in re.finditer("\n", procstr[:ps])]
        pos1 = possibilites[-1:][0]
    elif procstr.find("C.F.R.") != -1:
        ps = procstr.find("C.F.R.")
        possibilites = [m.start() for m in re.finditer("\n", procstr[:ps])]
        pos1 = possibilites[-1:][0]
    else:
        return ipr_data[target]["fd_type(s)"], ipr_data[target]["FWD?"], False

    # Look for a "judge" or "administrative" or "patent" followed by first "\n"
    stopper = False
    if procstr.lower().find("judge",0,pos1) != -1:
        ps = [m.start() for m in re.finditer("judge", procstr[:pos1].lower())]
        phold = ps[-1:][0]
        pos0 = procstr[phold:pos1].find("\n")
        if pos0 != -1:
            pos0 += phold + 1
            stopper = True

    if stopper == False and procstr.lower().find("administrative",0,pos1) != -1:
        ps = [m.start() for m in re.finditer("administrative", procstr[:pos1].lower())]
        phold = ps[-1:][0]
        pos0 = procstr[phold:pos1].find("\n")
        if pos0 != -1:
            pos0 += phold + 1
            stopper = True

    if stopper == False and procstr.lower().find("patent",0,pos1) != -1:
        ps = [m.start() for m in re.finditer("patent", procstr[:pos1].lower())]
        phold = ps[-1:][0]
        pos0 = procstr[phold:pos1].find("\n")
        if pos0 != -1:
            pos0 += phold + 1
            stopper = True

    if stopper == False:
        return ipr_data[target]["fd_type(s)"], ipr_data[target]["FWD?"], False
    else:
        snip = procstr[pos0:pos1]

    # Clean up the results snip
    snip = snip.replace("\n", " ")
    if snip[0] == " ": snip = snip[1:]
    snip = snip.replace("J UDGMENT", "JUDGMENT")
    snip = snip.replace("0f", "of")
    snip = snip.replace("0F", "OF")

    # Check for occurences of final written decision
    score = 0
    words = ["final", "written", "decision"]
    for word in words:
        if snip.lower().find(word) != -1:
            score += 1

    if score >= 2: fwd = True
    else: fwd = False

    return snip, fwd, True

def write_ipr_data(data_in, keys, file_out):
    # Delete existing file if it exists
    if os.path.isfile(file_out) == True:
        os.remove(file_out)     

    # Open a workbook and a worksheet
    workbook = xlsxwriter.Workbook(file_out)
    worksheet = workbook.add_worksheet()
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
    
    # Write headers:
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

    # Write in data for each ipr
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

    workbook.close()
    
def main():
    # Step 1: compile list of ipr documents to analyze
    subpath = os.path.join(in_dir,fold)
    for file in os.listdir(subpath):
        create_dictionary_entry(file)

        # Step 2: set attributes based on filename
        # Preliminarily set "trial_num(s)" and "trial_type" based on the filename
        file_ident = file[5::]; switch = False
        if file_ident[0:3] == "IPR":
            ipr_data[file]["trial_type"] = "IPR"
            switch = True
        elif file_ident[0:3] == "CBM":
            ipr_data[file]["trial_type"] = "CBM"
            switch = True
        elif file_ident[0:3] == "PGR":
            ipr_data[file]["trial_type"] = "PGR"
            switch = True

        if switch == True:
            trial_hold = file_ident[0:13]
            if trial_hold[3:7].isdigit() == True and trial_hold[7] == '-' and trial_hold[8::].isdigit() == True:
                ipr_data[file]["trial_num(s)"].append(trial_hold)

        # Preliminiarily set "fd_type(s)" based on the filename
        if 'final written decision' in file_ident.lower():
            ipr_data[file]["fd_type(s)"] = "final written decision"
            ipr_data[file]["FWD?"] = True
        if 'terminating' in file_ident.lower() or 'termination' in file_ident.lower():
            if ipr_data[file]["fd_type(s)"] is None:
                ipr_data[file]["fd_type(s)"] = "termination of proceeding"
            else: ipr_data[file]["fd_type(s)"] += " and termination of proceeding"

    ######## NOT USING THIS ######### Step 3: Identify and target possible IPRs (round 1) ###########
    tad = ipr_data.copy()
    # Must meet 2 conditions: 1. Not CBM or PGR 2. Not "termination of proceeding"
    #tad = {k1: v1 for k1, v1 in tad.items() if (v1["trial_type"] != "CBM" and 
    #        v1["trial_type"] != "PGR" and "termination of proceeding" not in v1["fd_type(s)"])}
    targets = tad.keys()

    # Step 4: Analyze the contents of the first page of targets
    printcounter = 1
    for target in targets:
        print(printcounter)
        printcounter += 1
        # Convert first page into png image
        file = os.path.join(subpath, target + "[0]")
        with IMG(filename = file, resolution=res) as imgs:
            imgs.compression_quality = 99
            with imgs.sequence[0] as img:
                img.type = 'truecolor'
                IMG(img).save(filename = "in1.png")

        # Read text using tesseract OCR
        imagein = Image.open("in1.png")
        imagein = imagein.convert('L')
        imagein = imagein.filter(ImageFilter.SHARPEN)
        tessdata_dir_config = '--tessdata-dir "C:\\Program Files (x86)\\Tesseract-OCR\\tessdata" -oem 2 -psm 11'
        text_read = pytesseract.image_to_string(imagein, boxes = False, config=tessdata_dir_config)

        procstr = text_read
        ## Pull decision date from the first page
        ddate, error_free = read_desc_date(procstr)
        ipr_data[target]["dec_date"] = ddate
        ipr_data[target]["no_issues"] = bool(ipr_data[target]["no_issues"] * error_free)        

        #print(target)
        #print(ddate)
        #print(error_free)
        #print("")
        
        # Pull petitioner names from the first page
        petitioners, error_free, ph_start = read_petitioner_names(procstr)
        ipr_data[target]["pet_name(s)"].extend(petitioners)
        ipr_data[target]["no_issues"] = bool(ipr_data[target]["no_issues"] * error_free)

        # Pull patent holder names from the first page
        pholders, error_free, trialnum_start = read_pholder_names(procstr, ph_start)
        ipr_data[target]["ph_name(s)"].extend(pholders)
        ipr_data[target]["no_issues"] = bool(ipr_data[target]["no_issues"] * error_free)

        #print(target)
        #print("----Petitioners----")
        #for pet in petitioners:
        #    print(pet)
        #print("----Patent Holders----")
        #for ph in pholders:
        #    print(ph)
        
        # Pull trial number from the first page
        trial_nums, error_free = read_trial_nums(procstr, trialnum_start, target)
        ipr_data[target]["trial_num(s)"] = trial_nums
        ipr_data[target]["no_issues"] = bool(ipr_data[target]["no_issues"] * error_free)

        #if error_free == False:
        #    print(target)
        #    print(procstr)
        #    print(trial_nums)
        #    print(error_free)
        #    print("")

        # Figure out trial type based on list of trial numbers
        trial_type, error_free = decide_trial_type(trial_nums)
        ipr_data[target]["trial_type"] = trial_type
        ipr_data[target]["no_issues"] = bool(ipr_data[target]["no_issues"] * error_free)    
            
        #print(target)
        #print(trial_type)
        #print(error_free)
        #print("")

        # Pull patent numbers from the first page
        pat_nums, pat_types, mult_pat, error_free, dectypestarter = read_pat_nums(procstr, trialnum_start)
        ipr_data[target]["pat_num(s)"] = pat_nums
        ipr_data[target]["pat_type(s)"] = pat_types
        ipr_data[target]["mult_pat"] = mult_pat
        ipr_data[target]["no_issues"] = bool(ipr_data[target]["no_issues"] * error_free)    

        #print(target)
        #print("patents:")
        #for pat in pat_nums:
        #    print(pat)
        #    if len(pat) > 7: print("LONGER THAN 7 CHARs")
        #print(pat_types)
        #print(error_free)
        #print("")

        # Pull decision type(s) from the first page
        dec_types, fwd, error_free = read_dec_type(procstr, dectypestarter, target)
        ipr_data[target]["fd_type(s)"] = dec_types
        ipr_data[target]["FWD?"] = fwd
        ipr_data[target]["no_issues"] = bool(ipr_data[target]["no_issues"] * error_free)            

        #print(target)
        #print(dec_types)
        #print(fwd)
        #print(error_free)
        #print("")

        print(target)
        print("----Petitioners----")
        for pet in petitioners:
            print(pet)
        print("----Patent Holders----")
        for ph in pholders:
            print(ph)
        print("----------------------")
        print(trial_type)
        print("patents:")
        for pat in pat_nums:
            print(pat)
            if len(pat) > 7: print("LONGER THAN 7 CHARs")
        print(dec_types)
        print(fwd)
        print(error_free)
        print("")

        # Cleanup Wand's shit (temporary files)
        os.remove("in1.png")
        file_dump = os.listdir(temp_dir)
        for filed in file_dump:
            if "magick" in filed:
                try: os.remove(os.path.join(temp_dir,filed))
                except PermissionError:
                    continue
    
    ## Step 5: Output to excel file for manipulation
    write_ipr_data(ipr_data, targets, out_file)


if __name__ == "__main__":
    main()