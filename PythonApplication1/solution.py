"""IPR Efficacy Analysis Solution"""
# Jonathan Slack
# jslackd@gmail.com

import numpy as np
import cv2
import os
from math import pi
import random
import sys
import csv
import collections
from datetime import date
from datetime import datetime
import statistics
import xlsxwriter
import pandas as pd
from pandas import ExcelWriter
from pandas import ExcelFile
import math

import pull_ipr_ptab

in_dir = "in_data"
fold = "claim_data"
file = "claim.tsv"


def main():
    print("Step 1")
    ## Step X: Pull down IPR final decision(s), if needed
    #pull_ipr_ptab.main()
    ## If there are application files missing, then stop execution

    path = os.path.join(in_dir, fold, file)
    count = 0

    with open(path, encoding = "utf8") as tsvfile:
        tsvreader = csv.reader(tsvfile, delimiter="\t")
        for line in tsvreader:
            #print(line)
            count += 1
            if count > 20000:
                break


        
    


if __name__ == "__main__":
    main()