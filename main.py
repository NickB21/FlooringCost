import glob
import os
import pandas as pd
import numpy as np
from pandas import ExcelFile
from pandas import ExcelWriter
from datetime import datetime

# Location of Price Sheets Folders by Vendor
price_sheet_loc = 'C:/Users/nickh/Desktop/Flooring/2_PriceLists'
# Location of Updated Single Excel File
updated_sheet_loc = 'C:/Users/nickh/Desktop/Flooring/1_Combined_Price_List'


########################################################################################################
# Magic Formula
def carpet_equation(price):
    return (price * 2) + 2


########################################################################################################
# Disable Chained Assignments to not copy df strings
pd.options.mode.chained_assignment = None
pd.options.display.float_format = '{:, .2f}'.format

# Date
now = datetime.now()
year = now.strftime("%Y")
month = now.strftime("%m")
day = now.strftime("%d")
date = now.strftime("%Y-%m-%d")


# Functions

def newest_price_files(filepath):
    # Get new file from folder by time of edit
    list_of_files = glob.glob(filepath + '/*')
    latest_file = max(list_of_files, key=os.path.getctime)
    return latest_file


def folder_grab(price_sheet_location):
    # List of folders of vendors
    folders = os.listdir(price_sheet_location)
    for i, a in enumerate(folders):
        folders[i] = price_sheet_location + "/" + a
    return folders


def recent_file_list(recent_file):
    # Getting List of recent files
    for j, b in enumerate(recent_file):
        recent_file[j] = os.path.normpath(newest_price_files(b))
    latest_files = recent_file
    return latest_files


def csv_grab(folders_in, latest_files_in):
    i = 0
    csvs = []
    csvs_loc = []
    while i < len(latest_files_in):
        if latest_files_in[i].endswith(".csv") or latest_files_in[i].endswith(".xlsx") or latest_files_in[i].endswith(
                ".xls"):
            csvs.append(latest_files_in[i])
            csvs_loc.append(i)
            i = i + 1
        else:
            i = i + 1

    # CSVS folder retrieval
    csv_file = []
    i = 0
    while i < len(csvs_loc):
        csv_file.append(folders_in[csvs_loc[i]])
        i = i + 1
    return csv_file, csvs


def pdf_grab(folders_in, latest_files_in):
    i = 0
    pdfs = []
    pdfs_loc = []
    while i < len(latest_files_in):
        if latest_files_in[i].endswith(".pdf"):
            pdfs.append(latest_files_in[i])
            pdfs_loc.append(i)
            i = i + 1
        else:
            i = i + 1

    # PDF folder retrieval
    pdf_file = []
    i = 0
    while i < len(pdfs_loc):
        pdf_file.append(folders_in[pdfs_loc[i]])
        i = i + 1
    return pdf_file, pdfs


def print_separate():
    print(".")
    print(".")
    print(".")


def test(dataframe):
    where2save_test = updated_sheet_loc + '/' + date + '.xlsx'
    writer_test = ExcelWriter(where2save_test)
    dataframe.to_excel(writer_test, 'Sheet1', index=False)
    writer_test.save()


########################################################################################################
# Templates

def shaw_template(shaw_raw, vendors, j):
    # Obtain only columns needed
    shaw_d2 = shaw_raw[['style', 'stylename', 'color', 'colorname', 'selling company name', 'cutprice', 'rollprice', 'size']]
    # Brand
    shaw_d2.rename(columns={"selling company name": "brand"}, inplace=True)
    # Add Vendor
    shaw_d2.insert(loc=0, column='vendor', value=vendors[j])
    # Drop Color bc Mohawk doesn't include it
    shaw_d2.drop(["color", "colorname"], axis=1, inplace=True)
    shaw_d2 = shaw_d2.reindex(columns=['vendor', 'brand', 'style', 'stylename', 'cutprice', 'rollprice', 'size'])
    # Dropping Duplicate values in rows bc color was previously listed
    shaw_d2.drop_duplicates(subset=None, keep="first", inplace=True)
    # Keep only 12 AND 15ft roll sizes (removes most rugs & Misc)
    shaw_d2 = shaw_d2[shaw_d2['size'].str.startswith(('12', '15'), na=False)]

    # Pricing
    # sy to sf
    shaw_d2['cutprice'] = shaw_d2['rollprice'].div(9)
    # Markup
    shaw_d2['cutprice'] = shaw_d2['cutprice'].apply(carpet_equation)
    # Format
    shaw_d2['cutprice'] = shaw_d2['cutprice'].apply(lambda x: "${:.2f}".format(x))

    shaw_d2['rollprice'] = shaw_d2['rollprice'].div(9)
    shaw_d2['rollprice'] = shaw_d2['rollprice'].apply(carpet_equation)
    shaw_d2['rollprice'] = shaw_d2['rollprice'].apply(lambda x: "${:.2f}".format(x))

    # Alphabetize by stylename
    shaw_d2 = shaw_d2.sort_values(['brand', 'stylename'])

    # Formatting Size
    shaw_d2['size'] = shaw_d2['size'].astype(str).str[:2] + "00"
    shaw_d2['size'] = shaw_d2['size'].astype(str).str[:2] + " " + shaw_d2['size'].astype(str).str[2:]

    return shaw_d2


def mohawk_template(mohawk_df, vendors, k):
    mohawk_d2 = mohawk_df
    # Place blanks with NAs
    mohawk_d2 = mohawk_d2.replace(r'^\s*$', np.nan, regex=True)
    # Filter out blank rows
    mohawk_d2.dropna(subset=['Style#'], inplace=True)
    # Lowercase Header
    mohawk_d2 = mohawk_d2.rename(columns=str.lower)
    # Rid of spaces in header
    mohawk_d2.columns = mohawk_d2.columns.str.replace(' ', '', regex=True)
    # Rid of special characters in header
    mohawk_d2.columns = mohawk_d2.columns.str.replace('[#,@,&,.,(,)]', '', regex=True)
    # Rid of blank columns
    mohawk_d2.dropna(how='all', axis=1, inplace=True)
    # Filter out rows that repeat "Style#"
    mohawk_d2 = mohawk_d2[mohawk_d2['style'] != 'Style#']
    # Updating to most recent cost increase
    mohawk_d2['effectivedate'] = pd.to_datetime(mohawk_d2['effectivedate']).astype(str)
    mohawk_d2 = mohawk_d2.sort_values('effectivedate').drop_duplicates('style', keep="last", inplace=False)
    mohawk_d2 = mohawk_d2.sort_values('style')
    # Drop columns not used
    mohawk_d2.drop(["backing", "minqty", "rollpricesqyd", "cutpricesqyd"], axis=1, inplace=True)
    # Reindex to match Shaw Format
    mohawk_d2 = mohawk_d2.reindex(columns=['brand', 'style', 'stylename', 'cutpricesqft', 'rollpricesqft', 'size'])

    # Reformat column names
    mohawk_d2.rename(columns={"cutpricesqft": "cutprice", "rollpricesqft": "rollprice"}, inplace=True)
    mohawk_d2.insert(loc=0, column='vendor', value=vendors[k])

    mohawk_d2['cutprice'] = mohawk_d2['cutprice'].astype(float)
    mohawk_d2['cutprice'] = mohawk_d2['cutprice'].apply(carpet_equation)
    mohawk_d2['cutprice'] = mohawk_d2['cutprice'].apply(lambda x: "${:.2f}".format(x))

    mohawk_d2['rollprice'] = mohawk_d2['rollprice'].astype(float)
    mohawk_d2['rollprice'] = mohawk_d2['rollprice'].apply(carpet_equation)
    mohawk_d2['rollprice'] = mohawk_d2['rollprice'].apply(lambda x: "${:.2f}".format(x))

    mohawk_d2['size'] = mohawk_d2['size'].astype(str).str[:4]
    mohawk_d2['size'] = mohawk_d2['size'].astype(str).str[:2] + " " + mohawk_d2['size'].astype(str).str[2:]
    mohawk_d2 = mohawk_d2.sort_values(['brand', 'stylename'])

    return mohawk_d2


########################################################################################################
# Main
if __name__ == '__main__':
    folders = os.listdir(price_sheet_loc)
    temp = folder_grab(price_sheet_loc)
    folder_loc = temp
    latest_files = recent_file_list(temp)
    print("Today's Date:_________________", date)
    print("All Vendors: _________________", folders)
    print("Newest Files from all Vendors:", latest_files)

    print_separate()
    csvs_out = []
    csvs_file = []
    csvs_out, csvs_file = csv_grab(folders, latest_files)
    print("CSV Vendors:_________________", csvs_out)
    print("CSVs:________________________", csvs_file)

    print_separate()
    pdfs_out, pdfs_file = pdf_grab(folders, latest_files)
    print("PDF Vendors:_________________", pdfs_out)
    print("PDFs:________________________", pdfs_file)
    print_separate()

    # Loop through CSV Vendors
    i = 0
    while i < len(csvs_out):
        if csvs_out[i] == "Shaw":
            shaw_df = pd.read_excel(csvs_file[i])
            shaw_formatted = shaw_template(shaw_df, csvs_out, i)
            # test(shaw_formatted)
            i = i + 1

        elif csvs_out[i] == "Mohawk":
            print(csvs_file[i])
            mohawk_df = pd.read_excel(csvs_file[i], header=8)
            mohawk_formatted = mohawk_template(mohawk_df, csvs_out, i)
            # test(mohawk_formatted)
            i = i + 1

        else:
            i = i + 1

    vertical_concat = pd.concat([mohawk_formatted, shaw_formatted], axis=0)
    test(vertical_concat)

    # Working.....__________________________________________


    # # Saving Combined DF into CSV
    # where2save = updated_sheet_loc + '/' + date + '.xlsx'
    # writer = ExcelWriter(where2save)
    # d2.to_excel(writer, 'Sheet1', index=False)
    # writer.save()

    # print(pdfs, pdfs_loc)
    # data_csv = [folders, latest_files]
    # df_csv = pd.DataFrame(data_csv)
    # print("df_csv")
    # print(df_csv)

    # data_csvs = [folders, latest_files]
    # df_csvs = pd.DataFrame(data_csvs)
    # print("df_csvs")
    # print(df_csvs)

    # data_pdf = [folders, latest_files]
    # df_pdf = pd.DataFrame(data_pdf)
    # print("df_pdf")
    # print(df_pdf)

    # Separate DF into CSV and PDF
    # for column_name in df:
    #     print(column_name)
    #     print('------\n')
