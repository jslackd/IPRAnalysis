"""Reading IPR pdf files"""
# Jonathan Slack
# jslackd@gmail.com

import numpy as np
import cv2
import os
from math import pi
import random
import sys
from PIL import Image, ImageEnhance, ImageFilter
import pytesseract
import collections
import re
from datetime import date
from datetime import datetime
from multiprocessing.dummy import Pool as ThreadPool
from wand.image import Image as IMG
import os
import numpy as np
import PyPDF2
from multiprocessing.dummy import Pool as ThreadPool
import itertools
import re

in_dir = "in_data"
#fold = "all_iprs"
fold = "test_docs2"
out_file = "ipr_read_data.xlsx"

res = 400

# Declare data structure
ipr_data = collections.OrderedDict()
# Format: {filename.pdf (str): {attrib (str): val (---)}}
# Attributes:
#       "trial_num(s)"  : ["IPR2015-00010"] or ["CBM2015-00004"] or ["PGR2015-00003"] or []
#       "trial_type"    : "IPR" or "CBM" or "PGR" or None
#       "fd_type(s)"    : ["final written decision"] or ["judgment"] or ["termination of proceeding"] or ["unknown"] or []
#       "mult_pat"      : True or False
#       "dec_date"      : "12/16/2015" or None
#       "pet_name(s)"   : ["BIO-RAD LABORATORIES, INC.,"] or ["unknown"] or []
#       "ph_name(s)"    : ["CALIFORNIA INSTITUTE OF TECHNOLOGY"] or ["unknown"] or []
#       "pat_num(s)"    : ["6658464"] or ["43919"] or ["unknown"] or []
#       "pat_type(s)"   : ["B2"] or ["RE"] or ["unknown"] or []
#       "order_txt"     : "ORDERED that the joint motion to terminate the proceeding is GRANTED and . . ."
#       "order_disp(s)" : [["6658464", "unpatentable", [1,2,3,4,5,6,7,8,9,14]], [ ] ]  or []
#       "no_issues"     : False or True

def create_dictionary_entry(fname):
    ipr_data[fname] = {
        "trial_num(s)": [], "trial_type": None, "fd_type(s)": [], "mult_pat": False,
        "dec_date": None, "pet_name(s)": [], "ph_name(s)": [], "pat_num(s)": None,
        "pat_type(s)": [], "order_txt": None, "order_disp(s)": [], "no_issues": True
    }

def read_petitioner_names(procstr):
    # This function will find the petitioner names on the first page of an IPR
    # If it fails, it will return [], False (False being for issue flag)

    # Find the first instance of "APPEAL BOARD" or "PATENT TRIAL AND" or "BEFORE THE"
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
                if to_add[0] == " ": to_add = to_add[1::]
                if to_add[len(to_add)-1] == ",": to_add = to_add[:-1]
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

    return petitioners_new, True, ender
    
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
        return [], False

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
        return [], False

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
    return pholders_new, True

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
            trial_hold = file_ident[3:13]
            if trial_hold[0:4].isdigit() == True and trial_hold[4] == '-' and trial_hold[5::].isdigit() == True:
                ipr_data[file]["trial_num(s)"].append(trial_hold)

        # Preliminiarily set "fd_type(s)" based on the filename
        if 'final written decision' in file_ident.lower():
            ipr_data[file]["fd_type(s)"].append("final written decision")
        if 'terminating' in file_ident.lower() or 'termination' in file_ident.lower():
            ipr_data[file]["fd_type(s)"].append("termination of proceeding")

    # Step 3: Identify and target possible IPRs (round 1)
    tad = ipr_data.copy()
    # Must meet 2 conditions: 1. Not CBM or PGR 2. Not "termination of proceeding"
    tad = {k1: v1 for k1, v1 in tad.items() if (v1["trial_type"] != "CBM" and 
            v1["trial_type"] != "PGR" and "termination of proceeding" not in v1["fd_type(s)"])}
    targets = tad.keys()

    # Step 4: Analyze the contents of the first page of targets
    for target in targets:
        # Convert first page into png image
        file = os.path.join(subpath, target + "[0]")
        with IMG(filename = file, resolution=res) as imgs:
            imgs.compression_quality = 99
            img = imgs.sequence[0]
            img.type = 'truecolor'
            IMG(img).save(filename = "temporary.png")

        # Read text using tesseract OCR
        imagein = Image.open("temporary.png")
        imagein = imagein.convert('L')
        imagein = imagein.filter(ImageFilter.SHARPEN)
        tessdata_dir_config = '--tessdata-dir "C:\\Program Files (x86)\\Tesseract-OCR\\tessdata" -oem 2 -psm 11'
        text_read = pytesseract.image_to_string(imagein, boxes = False, config=tessdata_dir_config)

        # Pull petitioner names from the first page
        procstr = text_read
        petitioners, error_free, ph_start = read_petitioner_names(procstr)
        ipr_data[target]["pet_name(s)"].extend(petitioners)
        ipr_data[target]["no_issues"] = bool(ipr_data[target]["no_issues"] * error_free)

        # Pull patent holder names from the first page
        pholders, error_free = read_pholder_names(procstr, ph_start)
        ipr_data[target]["ph_name(s)"].extend(pholders)
        ipr_data[target]["no_issues"] = bool(ipr_data[target]["no_issues"] * error_free)
        
    
        print(target)
        print("----Petitioners----")
        for pet in petitioners:
            print(pet)
        print("----Patent Holders----")
        for ph in pholders:
            print(ph)
        print("\n\n")
            
        

        #print(target)
        #print(text_read)
        #print("")
        breakinghere = 1


if __name__ == "__main__":
    main()