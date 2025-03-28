import gspread
from oauth2client.service_account import ServiceAccountCredentials
from collections import defaultdict
import pandas as pd
import re
subject_list=[]
def authenticate_and_get_sheet(sheet_url):
     # Authenticate using the Service Account JSON file
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    credentials = ServiceAccountCredentials.from_json_keyfile_name("/etc/secrets/Cred.json", scope)
    client = gspread.authorize(credentials)
    # Open the Google Sheet by URL
    workbook = client.open_by_url(sheet_url)
    return workbook

def get_sheet_names(workbook):
    """Gets the list of all sheet names from a Google Sheets workbook."""
    return [worksheet.title for worksheet in workbook.worksheets()]

def fetch_sheet_data(workbook,course,sheet_name,date1,date2): 
    worksheet = workbook.worksheet(sheet_name)  #getting data by sheet name
    data = worksheet.get_all_records()
    data = pd.DataFrame(data)
    print("data sent to get roll")
    data=get_roll_name_and_clean(data,course,date1,date2)
    print("data recieved from get roll")
    return data

def get_for_particular_student(data,student_name):
    grouped_data=data[data[student_name]=="Present"].groupby("Select Subject",observed=True)[student_name].count()
    grouped_data=grouped_data.reset_index()
    grouped_data["Total"]=grouped_data["Select Subject"].map(data.groupby("Select Subject")[student_name].count())
    grouped_data["Percentage"]=((grouped_data[student_name]/grouped_data["Total"])*100).round(2)
    print(grouped_data)
    return grouped_data

def get_subjectwise(data,subject):
    print("yehi nhi chal rha hai")
    grouped_data=data.groupby("Select Subject",observed=True)[data.iloc[:,3:-2].columns].agg(lambda x: (x == "Present").sum())
    grouped_data=grouped_data.loc[subject,:].to_frame()
    Total_classes=data[data["Select Subject"]==subject].shape[0]
    grouped_data=grouped_data.rename(columns={subject:f"Classes_Attended (out of ({Total_classes}))"})
    grouped_data["Percentage_Classes_Attended"]=(grouped_data[f"Classes_Attended (out of ({Total_classes}))"]/Total_classes*100).round(2)
    grouped_data.reset_index(inplace=True)
    grouped_data.rename(columns={"index":"Student"},inplace=True)
     #generating name and roll number column
    grouped_data[['Roll_number', 'Student_Name']] = grouped_data['Student'].str.split(' ', n=1, expand=True)
    #reordering the dataframe
    grouped_data = grouped_data[[grouped_data.columns[-2],grouped_data.columns[-1]] + list(grouped_data.columns[:-2])]
    #removing column named Student
    grouped_data=grouped_data.drop(columns=["Student"])
    print(Total_classes)
    return grouped_data

def get_overall_attendance(data):
    data["Select Subject"]=data["Select Subject"].astype("str")
    grouped_data=data.groupby("Select Subject",observed=True)[data.iloc[:,3:-2].columns].agg(lambda x: (x == "Present").sum())
    actual_grouped_data = data.groupby("Select Subject", observed=True)[data.iloc[:, 3:-2].columns].count()
    actual_grouped_dict = actual_grouped_data.iloc[:, 0].to_dict()
    actual_grouped_dict["Total"] = sum(actual_grouped_dict.values())
    if "" in actual_grouped_dict:
        del actual_grouped_dict[""]
    print(actual_grouped_dict)
    grouped_data=grouped_data.T
    grouped_data.index = grouped_data.index.astype(str)
    grouped_data=grouped_data.reset_index()
    if "" in grouped_data.columns:
        grouped_data = grouped_data.drop(columns=[""])
    grouped_data["Total"]=grouped_data.iloc[:,1:].sum(axis=1)
    # Compute percentage
    print(grouped_data.iloc[:,1:].columns)
    df_percentage = grouped_data.iloc[:,1:].div(actual_grouped_dict) * 100  # Compute percentage
    df_percentage =df_percentage.round(2)
    df_percentage["index"] = grouped_data["index"]
    print(df_percentage.round(2))
    #actual_data
    print(grouped_data)
    #renaming columns
    rename_dict_grouped_data={key: f"{key} (out of {value})" for key, value in actual_grouped_dict.items()}
    rename_dict_df_percentage={key: f"{key} (in percentage)" for key in actual_grouped_dict.keys()}
    rename_dict_grouped_data["index"]="Student"
    rename_dict_df_percentage["index"]="Student"
    grouped_data=grouped_data.rename(columns=rename_dict_grouped_data)
    df_percentage=df_percentage.rename(columns=rename_dict_df_percentage)
    #generating name and roll number column
    grouped_data[['Roll_number', 'Student_Name']] = grouped_data['Student'].str.split(' ', n=1, expand=True)
    df_percentage[['Roll_number', 'Student_Name']] = df_percentage['Student'].str.split(' ', n=1, expand=True)
    #reordering the dataframe
    grouped_data = grouped_data[[grouped_data.columns[-2],grouped_data.columns[-1]] + list(grouped_data.columns[:-2])]
    df_percentage = df_percentage[[df_percentage.columns[-2],df_percentage.columns[-1]] + list(df_percentage.columns[:-2])]
    #removing column named Student
    df_percentage=df_percentage.drop(columns=["Student"])
    grouped_data=grouped_data.drop(columns=["Student"])
    if "" in grouped_data.columns:
        grouped_data = grouped_data.drop(columns=[""])
    if "" in df_percentage.columns:
        df_percentage = df_percentage.drop(columns=[""])
    return (grouped_data,df_percentage)

def filter_students_by_criteria(data,subject,criteria):
    data=get_subjectwise(data,subject)
    data=data[data["Percentage_Classes_Attended"]>=criteria]
    return data

def get_present(data,student_name,subject,date):
    date=pd.to_datetime(date)
    return data[(data["Select Subject"]==subject) & (data["Date"]=="2024-10-08")][student_name]

def get_roll_name_and_clean(data,course,date1,date2):
    Course={"MBA(MS)(5Y)":"IM", "MCA(5Y)":"IC", "MTech(IT)(5Y)":"IT", "B.Com.(Hons.)(3Y)":"IB"}
    col_dict={}
    columns=data.columns
    for col in columns:
        pattern = r'.*\[(\d*\.*\s*)?({}-2K[A-Za-z0-9-]+)\s*(.*)\]'.format(Course[course])
        match = re.search(pattern, col)
        if match:
            roll_number = match.group(2)  # Extract roll number
            name = match.group(3)  # Extract name
            new_col=roll_number +" "+ name
            col_dict[col]=new_col
    data=data.rename(columns=col_dict)
    data["Timestamp"] = pd.to_datetime(data["Timestamp"], format="%m/%d/%Y %H:%M:%S", dayfirst=True)
    data["Date"]=pd.to_datetime(data["Timestamp"].dt.normalize())
    print(data["Date"])
    data["Time"]=data["Timestamp"].dt.time
    data["Select Subject"]=data["Select Subject"].astype("category")
    global subject_list
    subject_list=list(data["Select Subject"].cat.categories)
    data=data.drop(columns=["Timestamp","Remark"], errors="ignore")
    print("shape of OG data:",data.shape)
    filtered_data=data
    if date1 and date2:
        # Convert user input to datetime64[ns]
        date1 = pd.to_datetime(date1, format="%Y-%m-%d", dayfirst=True) if date1 else None
        date2 = pd.to_datetime(date2, format="%Y-%m-%d", dayfirst=True) if date2 else None
        # Find the closest available date if not found
        if date1 and not (data["Date"] == date1).any():
            print(date1)
            date1 = data[data["Date"] > date1]["Date"].min()  # Get next closest date
        if date2 and not (data["Date"] == date2).any():
            print(date2)
            date2 = data[data["Date"] < date2]["Date"].max()  # Get previous closest date
        print(date1,date2)
        # Filter DataFrame
        filtered_data = data[(data["Date"] >= date1) & (data["Date"] <= date2)]
        print(filtered_data["Date"])
    return filtered_data

def get_bet_dates(data,date1,date2):
   # Convert user input to datetime64[ns]
    date1 = pd.to_datetime(date1, format="%Y-%m-%d", dayfirst=True) if date1 else None
    date2 = pd.to_datetime(date2, format="%Y-%m-%d", dayfirst=True) if date2 else None
    # Find the closest available date if not found
    if date1 and not (data["Date"] == date1).any():
        date1 = data[data["Date"] > date1]["Date"].min()
        print(date1)# Get next closest date
    if date2 and not (data["Date"] == date2).any():
        date2 = data[data["Date"] < date2]["Date"].max()  # Get previous closest date
        print(date2)
    # Filter DataFrame
    filtered_data = data[(data["Date"] >= date1) & (data["Date"] <= date2)]
    print(filtered_data["Date"])
    print("shape of filtered data:",filtered_data.shape)
    date_filtered_data=get_overall_attendance(filtered_data)
    return date_filtered_data
    
if __name__ == "__main__":
    workbook=authenticate_and_get_sheet("https://docs.google.com/spreadsheets/d/1VyqArt8jWYCxgqDdKv7RBW7_KF9pBz5Hm3vsUA6iUq4/edit?gid=967537692#gid=967537692")
    data=fetch_sheet_data(workbook,"MCA(5Y)","Sem VIII-A","2025-02-01","2025-02-15")   
    print(get_subjectwise(data,"IC-812 Theory of Computation"))