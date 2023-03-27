import os

from flask import Flask, flash, redirect, render_template, request, send_file
from openpyxl import Workbook, load_workbook
from werkzeug.utils import secure_filename

UPLOAD_FOLDER = 'uploads'
ALLOWED_EXTENSIONS = {'xlsx'}

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['SECRET_KEY'] = 'super secret key'

def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/', methods=['GET', 'POST'])
def index():
    return render_template('index.html')

@app.route('/process', methods=['GET','POST'])
def process():
    
    if 'file' not in request.files:
            print("NO FILES")
    # Get the uploaded files
    file1 = request.files['file1']
    file2 = request.files['file2']
    
    filename1 = secure_filename(file1.filename)
    filename2 = secure_filename(file2.filename)
    # Save the files to disk
    file1.save(os.path.join(app.config['UPLOAD_FOLDER'], filename1))
    file2.save(os.path.join(app.config['UPLOAD_FOLDER'], filename2))

    # Load the Excel files and process the data (replace with your own processing code)
    wb_missing = load_workbook(os.path.join(app.config['UPLOAD_FOLDER'], filename1))
    ws_missing = wb_missing.worksheets[1]
    
    print(wb_missing)
    
    wb_caregiver = load_workbook(os.path.join(app.config['UPLOAD_FOLDER'], filename2))
    ws_caregiver = wb_caregiver.worksheets[1]
    
    new_wb = Workbook()
    new_ws = new_wb.active
    
    new_ws['A1'] = 'Mobile Number'
    new_ws['B1'] = 'Date'
    new_ws['C1'] = 'Aide'
# Create a dictionary to store the dates for each person
    dates_dict = {}

    # Loop through each row of the sheet
    for row in ws_missing.iter_rows(min_row=2, values_only=True):
        aide_code = row[8][-4::] #column I
        date = row[2] # column C
        # Add the date to the dictionary for this person
        if aide_code in dates_dict:
            if date not in dates_dict[aide_code]:
                dates_dict[aide_code].append(date)  
        else:
            dates_dict[aide_code] = [date]

    # Loop through the dictionary and update the sheet
    row_num = 2
    for aide_code, dates in dates_dict.items():
        # Write the name to column A
        new_ws.cell(row=row_num, column=2, value=", ".join(sorted(dates)))   
        new_ws.cell(row=row_num, column=3, value=aide_code)
        row_num += 1
        
        
        
    number_dict = {}
    
    for row in ws_caregiver.iter_rows(min_row=2, values_only=True):
        aide_code = str(row[2])[-4::] #column C
        if row[1] != "":
            phone_number = "+1" + row[1].replace("-", "") # column B
            number_dict[aide_code] = [phone_number]
       
        
        # if phone_number:
        #     number_dict[aide_code] = [phone_number]
    
    for col in new_ws.iter_cols(min_col=3, max_col=3):
        for cell in col:
            if cell.value in number_dict:
             new_ws.cell(row=cell.row, column=1, value=number_dict[cell.value][0]) 
    

    # Save the result workbook to disk
    output_filename = 'output.xlsx'
    new_wb.save(output_filename)

    # Send the output file to the user for download
    response = send_file(output_filename, as_attachment=True)

    # Delete the uploaded files and output file
    os.remove(os.path.join('uploads', file1.filename))
    os.remove(os.path.join('uploads', file2.filename))
    os.remove(output_filename)

    return response

if __name__ == '__main__':
    app.run(debug=False, port=80, host="0.0.0.0")
