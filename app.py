from flask import Flask, request, render_template, send_file
from openpyxl import load_workbook
from io import BytesIO
import base64

app = Flask(__name__)

# Encode the default Excel file in Base64
with open("Comps Data.xlsx", "rb") as f:
    encoded_excel = base64.b64encode(f.read()).decode('utf-8')

@app.route('/')
def home():
    return render_template("index.html", encoded_excel=encoded_excel)

@app.route('/modify_excel', methods=['POST'])
def modify_excel():
    try:
        # Decode the Base64 file to a binary stream
        uploaded_excel = request.form.get("uploaded_excel")
        binary_excel = base64.b64decode(uploaded_excel)
        workbook = load_workbook(BytesIO(binary_excel))

        # Get Base Year from the form
        base_year = request.form.get("baseYear")
        if not base_year.isdigit():
            return "Invalid Base Year", 400

        # Modify the Excel file (e.g., cell D3) and preserve formatting
        sheet = workbook.active
        cell_address = "D3"
        original_style = sheet[cell_address].style
        sheet[cell_address].value = int(base_year)
        sheet[cell_address].style = original_style

        # Save the modified file to memory
        output = BytesIO()
        workbook.save(output)
        output.seek(0)

        # Return the modified file for download
        return send_file(output, as_attachment=True, download_name="modified_excel_file.xlsx")
    except Exception as e:
        return f"Error: {e}", 500

if __name__ == "__main__":
    app.run(debug=True)
