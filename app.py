from flask import Flask, request, jsonify, render_template
from utils.spreadsheet_utils import (
    authenticate_and_get_sheet, get_sheet_names, fetch_sheet_data,
    get_for_particular_student, get_subjectwise, get_overall_attendance
)

app = Flask(__name__)

workbook = None
course = None

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/authenticate_and_get_workbook", methods=["POST"])
def authenticate_and_get_workbook():
    global workbook, course
    course = request.json.get("course")
    sheet_url = request.json.get("sheet_url")
    
    if not sheet_url:
        return jsonify({"error": "Sheet URL required"}), 400  
    
    workbook = authenticate_and_get_sheet(sheet_url)
    sheet_names = get_sheet_names(workbook)
    
    return jsonify({"sheets": sheet_names})

@app.route("/fetch_data", methods=["GET"])
def fetch_data():
    sheet_name = request.args.get("sheet")
    course = request.args.get("course")
    
    if not sheet_name:
        return jsonify({"error": "Sheet name required"}), 400
    
    data = fetch_sheet_data(workbook, course, sheet_name)
    return data.to_json(orient="records")

@app.route("/get_attendance_options", methods=["GET"])
def get_attendance_options():
    return jsonify({"options": ["Overall Attendance", "Attendance by Student", "Attendance by Subject"]})

@app.route("/overall_attendance", methods=["GET"])
def overall_attendance():
    sheet_name = request.args.get("sheet")
    course = request.args.get("course")
    
    data = fetch_sheet_data(workbook, course, sheet_name)
    result = get_overall_attendance(data)
    
    return result.to_json(orient="records")

@app.route("/student_attendance", methods=["GET"])
def student_attendance():
    sheet_name = request.args.get("sheet")
    course = request.args.get("course")
    
    data = fetch_sheet_data(workbook, course, sheet_name)
    student_list = data["Student Name"].unique().tolist()
    
    return jsonify({"students": student_list})

@app.route("/subjectwise_attendance", methods=["GET"])
def subjectwise_attendance():
    sheet_name = request.args.get("sheet")
    course = request.args.get("course")
    
    data = fetch_sheet_data(workbook, course, sheet_name)
    subject_list = data.columns[1:].tolist()
    
    return jsonify({"subjects": subject_list})

if __name__ == "__main__":
    app.run(debug=True)
