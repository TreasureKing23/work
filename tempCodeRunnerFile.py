import pandas as pd
from openpyxl import load_workbook

# File paths
dwb_path = "Data.xlsx"
iwb_path = "Input.xlsx"
output_path="Updated.xlsx"

# Load Data 
dwb = pd.read_excel(dwb_path)

# Mapping from Data -> Input
field_map = {
    "ExamCode": "EXAMID",
    "ExamYear": "PERIODID",
    "EXCID":"CANID",
    "ERN": "ERN",
    "FirstName": "FNAME",
    "Lastname": "SURNAME",
    "Middlename": "MNAME",
    "Schoolcode": "SCHOOLID",
    "SchoolName": "SCHOOLN",
    "Gender": "GENDER",
    "DateofBirth": "DOB",
    "Gaurdian": "GNAME",
    "Address": "ADDR1",
    "Address_2": "ADDR2",
    "LMS_Account": "LMS",
    "PathNo": "PATH",
    "WardofState": "WARD",
    "Parish": "PARISH",
    "Region": "REGION",
    "Prefcode1": "PCODE1",
    "First_Preference": "PNAME1",
    "Prefcode2": "PCODE2",
    "Second_Preference": "PNAME2",
    "Prefcode3": "PCODE3",
    "Third_Preference": "PNAME3",
    "Prefcode4": "PCODE4",
    "Fourth_Preference": "PNAME4",
    "Prefcode5": "PCODE5",
    "Fifth_Preference": "PNAME5",
    "SRN": "SRN"
}

#load

df_dwb = pd.read_excel(dwb_path)
iwb= load_workbook(iwb_path)
ws1 = iwb.active

print(df_dwb.columns.tolist())