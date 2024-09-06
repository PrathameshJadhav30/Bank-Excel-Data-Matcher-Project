from flask import Flask, request, render_template, send_file
import pandas as pd
from io import BytesIO

app = Flask(__name__)

# Route for the homepage, rendering the upload form
@app.route('/')
def upload_form():
    return render_template('index.html')

# Route for handling the file upload and matching process
@app.route('/match', methods=['POST'])
def match_data():
    # Get the uploaded files and user inputs
    file1 = request.files['file1']
    file2 = request.files['file2']
    column = request.form['column']
    data_type = request.form['data_type']

    # Read the Excel files into pandas DataFrames
    df1 = pd.read_excel(file1)
    df2 = pd.read_excel(file2)

    # Try to match the data based on the common column
    try:
        # Matched data (inner join)
        matched_data = pd.merge(df1, df2, on=column, how='inner')

        # Unmatched data (rows in df1 but not in df2 or vice versa)
        unmatched_data1 = df1[~df1[column].isin(df2[column])]
        unmatched_data2 = df2[~df2[column].isin(df1[column])]
        unmatched_data = pd.concat([unmatched_data1, unmatched_data2])

    except KeyError:
        return f"Column '{column}' not found in one or both files."

    # Prepare the output in memory
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        if data_type == 'matched':
            matched_data.to_excel(writer, index=False, sheet_name="Matched Data")
        elif data_type == 'unmatched':
            unmatched_data.to_excel(writer, index=False, sheet_name="Unmatched Data")
        elif data_type == 'both':
            matched_data.to_excel(writer, index=False, sheet_name="Matched Data")
            unmatched_data.to_excel(writer, index=False, sheet_name="Unmatched Data")

    # Seek to the beginning of the BytesIO object so it can be sent as a file
    output.seek(0)

    # Return the Excel file as a downloadable response
    return send_file(output, download_name="data.xlsx", as_attachment=True)

if __name__ == "__main__":
    app.run(debug=True)
