"""Reading IPR final decision claim disposition"""
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
import pandas as pd
from pandas import ExcelWriter
from pandas import ExcelFile
import numpy as np
import PyPDF2
from multiprocessing.dummy import Pool as ThreadPool
import itertools

in_dir = "in_data"
#fold = "all_iprs"
fold = "test_docs"
out_file = "ipr_read_data.xlsx"
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
#       "order_disp(s)" : {"6658464": {"c-range": "1-9,14" , "disposition": "unpatentable"}}
#       "no_issues2"    : True or False (default is True)
#       "expect_ccd"    : False or True (default is False)

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
    worksheet2.set_column(8,8,5)
    worksheet2.set_column(9,9,15)
    worksheet2.set_column(10,10,15)

    # Write headers for our second worksheet:
    worksheet2.write('A1', 'Trial Number(s)', header_format)
    worksheet2.write('B1', 'Trial Type', header_format)
    worksheet2.write('C1', 'Trial Date', header_format)
    worksheet2.write('D1', 'Decision Type(s)', header_format)
    worksheet2.write('E1', 'Trial Patent Holder Name(s)', header_format)  
    worksheet2.write('F1', 'Relevant Order Text', header_format)  
    worksheet2.write('G1', 'Associated Patent', header_format)
    worksheet2.write('H1', 'Associated Patent Type', header_format) 
    worksheet2.write('I1', 'Affected Claim(s)', header_format)
    worksheet2.write('J1', 'Claim(s) Disposition', header_format) 
    worksheet2.write('K1', 'Filename', header_format)

    # Write in data for each ipr on our second worksheet:
    row = 1
    for key2 in keys2:
        pass

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
        # Pull text from last 4 pages of the pdf document
        text_read_list = pull_x_pages(subpath, target, -5)
        print(text_read_list)

        ## Clean up text for crucial keywords
        text_read_list = cleanup_text(text_read_list)

        # 

    ## Step X: Write all of our data to the output file
    #write_ipr_data2(ipr_data, ipr_data.keys(), targets, file_out)


if __name__ == "__main__":
    main()