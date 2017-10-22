"""Scripts for IPR pulling files from ptab"""
# Jonathan Slack
# jslackd@gmail.com

import os
import urllib.request
import zipfile
import pandas as pd
from pandas import ExcelWriter
from pandas import ExcelFile
import zipfile
import time
import sys

output_dir = "in_data"
final_dir = os.path.join(output_dir,"all_iprs")

# Helper function for unziping files
def unzip_folder(file_name):
    zip_ref = zipfile.ZipFile(os.path.join(output_dir,file_name), 'r')
    zip_ref.extractall(output_dir)
    zip_ref.close()

# Main function for downloading and sorting IPR files
def main():
    # Step 1: Get all IPR folders from PTAB archives
    max = 1947
    offset = 0
    
    while (offset + 100) < max:
        doc_name_start = "documents" + str(offset + 1) + "-" + str(offset + 100)

        # Folder exists so skip to the next possible file
        if os.path.isdir(os.path.join(output_dir,doc_name_start)) == True:
            offset += 100
            continue

        # The .zip file for ipr data already exists; just unzip and delete the file
        elif os.path.isfile(os.path.join(output_dir,doc_name_start + ".zip")) == True:
            unzip_folder(doc_name_start + ".zip")
            time.sleep(0.02)
            os.remove(os.path.join(output_dir,doc_name_start + ".zip"))
            offset += 100
            continue

        # No folder or zip folder, so download the folder from PTAB
        else:
            url = "https://ptabdata.uspto.gov/ptab-api/documents.zip?type=Final%20Decision&limit=100&offset=" + str(offset) +"&sort=-filingDatetime"
            try:
                with urllib.request.urlopen(url, capath = "ptabdatausptogov.p7c") as response, open(os.path.join(output_dir, doc_name_start + ".zip"), 'wb') as out_file:
                    data = response.read() # a `bytes` object
                    out_file.write(data)
                    out_file.close()
                  
            except urllib.error.HTTPError as e:
                if os.path.isfile(os.path.join(output_dir,doc_name_start + ".zip")) == True:
                    os.remove(os.path.join(output_dir,doc_name_start + ".zip"))   
                print(doc_name_start)
                print(e.code)
                print("")
                offset += 100
                continue
                #sys.exit("Execution has been terminated.")

            except TimeoutError:
                if os.path.isfile(os.path.join(output_dir,doc_name_start + ".zip")) == True:
                    os.remove(os.path.join(output_dir,doc_name_start + ".zip"))      
                print(doc_name_start)  
                print("Timeout error on download") 
                print("")
                offset += 100
                continue

            offset += 100

    # Step 2: Get all files from folders into main folder
    x = 1
    folder_list = os.listdir



    ###### NOT USEABLE (FOR COPY AND PASTE) ######
    ## Loop through app list and download apps
    #print_
    #for app_num_int in app_list:
    #    print(print_cnt/len(app_list)*100,"percent complete", end = "\r")
    #    print_cnt+=1

    #    app_num = str(app_num_int)
    #    app_num.replace(" ","")
    #    url = "http://storage.googleapis.com/uspto-pair/applications/" + app_num + ".zip"
    #    file_name = os.path.join(output_dir,app_num + ".zip")

    #    # If folder for app exists, then skip this application number
    #    if os.path.isdir(os.path.join(output_dir,app_num)) == True:
    #        continue

    #    # Only the zip file exists, so extract it and skip application number
    #    elif os.path.isfile(file_name) == True:
    #        # Unzip the file and delete the original zip file
    #        unzip_folder(file_name, app_num)
    #        time.sleep(0.02)
    #        os.remove(file_name)
    #        continue

    #    # Download the file from `url` and save it locally under `file_name`:
    #    else:
    #        try:
    #            with urllib.request.urlopen(url) as response, open(file_name, 'wb') as out_file:
    #                data = response.read() # a `bytes` object
    #                out_file.write(data)
    #                out_file.close()
    #            #response = requests.get(url, stream = True)
    #            #with open(file_name,'wb') as out_file:
    #            #    for chunk in response.iter_content(chunk_size=1024):
    #            #        if chunk:
    #            #            out_file.write(chunk)                    
    #        except TimeoutError:
    #            error_output_down.append(app_num)
    #            if os.path.isfile(file_name) == True:
    #                os.remove(file_name)                    
    #            continue
    #        except urllib.error.HTTPError:
    #            error_output_down.append(app_num) 
    #            if os.path.isfile(file_name) == True:
    #                os.remove(file_name) 
    #            continue

    #        # Unzip the file and delete the original zip file
    #        unzip_folder(file_name, app_num)
    #        time.sleep(0.02)
    #        os.remove(file_name)

if __name__ == "__main__":
    main()
