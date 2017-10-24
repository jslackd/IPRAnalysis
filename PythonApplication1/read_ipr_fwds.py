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

in_dir = "in_data"
#fold = "all_iprs"
fold = "test_docs2"
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
#       "no_issues"     : False or True
### Finding the following:
#       "order_txt"     : "ORDERED that the joint motion to terminate the proceeding is GRANTED and . . ."
#       "order_disp(s)" : [["6658464", "unpatentable", [1,2,3,4,5,6,7,8,9,14]], [ ] ]  or []

def create_dictionary_entry(fname):
    ipr_data[fname] = {
        "trial_num(s)": [], "trial_type": None, "fd_type(s)": None, "mult_pat": False,
        "dec_date": None, "pet_name(s)": [], "ph_name(s)": [], "pat_num(s)": None,
        "pat_type(s)": [], "order_txt": None, "order_disp(s)": [], "no_issues": True, "FWD?": False
    }


def pull_iprdata_ff(file_out, dir, fold):
    # Make sure each ipr in our data file is also in our data folder
    
    # Read ipr data from excel file
    if os.path.isfile(file_out) == True:
        # Read existing file
        df = pd.read_excel(file_out, sheetname='Sheet1')
        ipr_list_raw = df["Trial Number(s)"].tolist()
        to_add_all = df.iloc[0:]
        to_add_all = to_add_all.values.tolist()
        print(ipr_list_raw)
        breaker = 1

    else:
        sys.exit("Error: ipr_read_data.xlsx does not exist")


def main():
    # Step 1: Read contents of existing ipr_read_data excel file and save to dictionary format
    ipr_data = pull_iprdata_ff(out_file, in_dir,fold)




    ## Step 1: compile list of ipr documents to analyze
    #subpath = os.path.join(in_dir,fold)
    #for file in os.listdir(subpath):
    #    create_dictionary_entry(file)


if __name__ == "__main__":
    main()