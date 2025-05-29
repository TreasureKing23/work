import pandas as pd
from openpyxl import load_workbook

# File paths
dwb_path = "Data.xlsx"
iwb_output_path = "Input.xlsx"

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

# Target fields for Input
iwb_fields = [
    "EXAMID", "PERIODID", "CANID", "ERN", "FNAME", "SURNAME", "MNAME", "SCHOOLID", "SCHOOLN",
    "GENDER", "DOB", "GNAME", "ADDR1", "ADDR2", "TEL_NIM", "EMAIL", "LMS", "PATH", "WARD",
    "SECTOR", "PARISH", "REGION", "PCODE1", "PNAME1", "PCODE2", "PNAME2", "PCODE3", "PNAME3",
    "PCODE4", "PNAME4", "PCODE5", "PNAME5", "CCODE1", "CNAME1", "CCODE2", "CNAME2",
    "PAPER01", "PAPER02", "PAPER03", "PAPER04", "PAPER05", "PAPER06", "PAPER07", "SRN",
    "USERFLD1", "USERFLD2", "USERFLD3", "USERFLD4", "USERFLD5"
]

# Create the new DataFrame for Input format
data_for_iwb = pd.DataFrame(columns=iwb_fields)

# Populate mapped fields
for dwb_col, iwb_col in field_map.items():
    if dwb_col in dwb.columns:
        data_for_iwb[iwb_col] = dwb[dwb_col]

# Fill blanks fields for now
data_for_iwb.fillna("", inplace=True)

# Save to new workbook (Input)
data_for_iwb.to_excel(iwb_output_path, index=False)
print(f"Input Workbook created: {iwb_output_path}")