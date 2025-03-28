from flask import Flask, request, jsonify, render_template
import utils.spreadsheet_utils as ut 
from collections import OrderedDict
import json

app = Flask(__name__)

workbook = None
course = None
data= None
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
    
    workbook = ut.authenticate_and_get_sheet(sheet_url)
    sheet_names = ut.get_sheet_names(workbook)
    
    return jsonify({"sheets": sheet_names})

@app.route("/fetch_data", methods=["GET"])
def fetch_data():
    sheet_name = request.args.get("sheet")
    course = request.args.get("course")
    date1=request.args.get("date1")
    date2=request.args.get("date2")
    
    if not sheet_name:
        return jsonify({"error": "Sheet name required"}), 400

    global data 
    data = ut.fetch_sheet_data(workbook, course, sheet_name,date1,date2)
    
    return jsonify({"message": "Data initialized successfully"}), 200
    

@app.route("/get_attendance_options", methods=["GET"])
def get_attendance_options():
    return jsonify({"options": ["Overall Attendance", "Attendance by Student", "Attendance by Subject"]})

@app.route("/overall_attendance", methods=["GET"])
def overall_attendance():
    global data
    result = ut.get_overall_attendance(data.copy())
    result = {
        "actual": [OrderedDict(row) for row in result[0].to_dict(orient="records")],
        "percentage": [OrderedDict(row) for row in result[1].to_dict(orient="records")]
    }
    return json.dumps(result)


# API to get subjects list
@app.route('/get_subjects', methods=['GET'])
def get_subjects():
    print(ut.subject_list)
    return jsonify({"subjects":ut.subject_list})



# API endpoint to process attendance by subject
@app.route('/subjectwise_attendance', methods=['POST'])
def process_attendance():
    try: 
        subject = request.args.get("subject")

        if not subject:
            return jsonify({"error": "Subject is required"}), 400

        # Call function to process data
        global data
        data_by_subject = ut.get_subjectwise(data.copy(), subject)
        return data_by_subject.to_json(orient="records")
      
    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == "__main__":
    app.run(debug=True)
