import pandas as pd
import openpyxl
import io
from flask import Flask, render_template, request, make_response, url_for
import csv

app = Flask(__name__)

@app.route('/')
def home():
    return render_template('index.html')

@app.route('/convert', methods=['POST'])
def convert_excel_to_csv():
    # Get the uploaded file
    file = request.files['file']
    delimiter = request.form['delimiter']

    # Read the Excel file
    df = pd.read_excel(file)

   # Get the sheets to convert
    sheets = request.form.getlist('sheets')

    # Read the Excel file
    if len(sheets) == 0:
        # Convert all sheets
        df = pd.read_excel(file, sheet_name=None)
        sheet_names = list(df.keys())
        csv = pd.concat(df, ignore_index=True).to_csv(index=False, sep=delimiter)
    else:
        # Convert specific sheets
        df = pd.read_excel(file, sheet_name=sheets)
        sheet_names = sheets
        csv = df.to_csv(index=False, sep=delimiter)


    # Convert the CSV data to a file-like object
    output = io.StringIO(csv)

    # Create a response object with the CSV data as the body
    response = app.response_class(
        response=output.getvalue(),
        mimetype='text/csv',
        headers={'Content-Disposition': 'attachment; filename=output.csv'}
    )
    return response


if __name__ == '__main__':
    app.run(debug=True)