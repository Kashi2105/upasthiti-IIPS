import gspread
from oauth2client.service_account import ServiceAccountCredentials
from collections import defaultdict
import pandas as pd
import re
def authenticate_and_get_sheet(sheet_url):
     # Authenticate using the Service Account JSON file
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    credentials = ServiceAccountCredentials.from_json_keyfile_name("Cred.json", scope)
    client = gspread.authorize(credentials)
    # Open the Google Sheet by URL
    workbook = client.open_by_url(sheet_url)
    return workbook

def get_sheet_names(workbook):
    """Gets the list of all sheet names from a Google Sheets workbook."""
    return [worksheet.title for worksheet in workbook.worksheets()]

def fetch_sheet_data(workbook,course,sheet_name): 
    worksheet = workbook.worksheet(sheet_name)  #getting data by sheet name
    data = worksheet.get_all_records()
    data = pd.DataFrame(data)
    print("data sent to get rolll")
    data=get_roll_name_and_clean(data,course)
    print("data recieved from get rolll")
    return data

def get_for_particular_student(data,student_name):
    grouped_data=data[data[student_name]=="Present"].groupby("Select Subject",observed=True)[student_name].count()
    grouped_data=grouped_data.reset_index()
    grouped_data["Total"]=grouped_data["Select Subject"].map(data.groupby("Select Subject")[student_name].count())
    grouped_data["Percentage"]=((grouped_data[student_name]/grouped_data["Total"])*100).round(2)
    print(grouped_data)
    return grouped_data

def get_subjectwise(data,subject):
    grouped_data=data.groupby("Select Subject",observed=True)[data.iloc[:,3:-2].columns].agg(lambda x: (x == "Present").sum())
    grouped_data=grouped_data.loc[subject,:].to_frame()
    grouped_data=grouped_data.rename(columns={subject:"Classes_Attended"})
    Total_classes=data[data["Select Subject"]==subject].shape[0]
    grouped_data["Percentage_Classes_Attended"]=(grouped_data["Classes_Attended"]/Total_classes*100).round(2)
    grouped_data.reset_index(inplace=True)
    grouped_data.rename(columns={"index":"Student"},inplace=True)
    print(Total_classes)
    return grouped_data

def get_overall_attendance(data):
    data["Select Subject"]=data["Select Subject"].astype("str")
    grouped_data=data.groupby("Select Subject",observed=True)[data.iloc[:,3:-2].columns].agg(lambda x: (x == "Present").sum())
    grouped_data=grouped_data.T
    grouped_data.index = grouped_data.index.astype(str)
    grouped_data=grouped_data.reset_index()
    grouped_data=grouped_data.rename(columns={"index":"Student"})
    grouped_data["Total"]=grouped_data.iloc[:,1:].sum(axis=1)
    return grouped_data

def filter_students_by_criteria(data,subject,criteria):
    data=get_subjectwise(data,subject)
    data=data[data["Percentage_Classes_Attended"]>=criteria]
    return data

def get_present(data,student_name,subject,date):
    date=pd.to_datetime(date)
    return data[(data["Select Subject"]==subject) & (data["Date"]=="2024-10-08")][student_name]

def get_roll_name_and_clean(data,course):
    Course={"MBA(MS)(5Y)":"IM", "MCA(5Y)":"IC", "MTech(IT)(5Y)":"IT", "B.Com.(Hons.)(3Y)":"IB"}
    col_dict={}
    columns=data.columns
    for col in columns:
        pattern = r'\[({}-2K[A-Za-z0-9-]+)\s*(.*)\]'.format(Course[course])
        match = re.search(pattern, col)
        if match:
            roll_number = match.group(1)  # Extract roll number
            name = match.group(2)  # Extract name
            new_col=roll_number +" "+ name
            col_dict[col]=new_col
    data=data.rename(columns=col_dict)
    data["Timestamp"]=pd.to_datetime(data["Timestamp"])
    data["Date"]=pd.to_datetime(data["Timestamp"].dt.date)
    data["Time"]=data["Timestamp"].dt.time
    data["Select Subject"]=data["Select Subject"].astype("category")
    data=data.drop(columns=["Timestamp","Remark"], errors="ignore")
    return data

if __name__ == "__main__":
    workbook=authenticate_and_get_sheet("https://docs.google.com/spreadsheets/d/1KFjU2WfAwAVTH3C9ZNWd3aZYGkm_nuWc-hePuB1lduY/edit?gid=1499235755#gid=1499235755")
    data=fetch_sheet_data(workbook,"MBA(MS)(5Y)","I Sem Sec-A")
    print(get_overall_attendance(data))