from flask import Flask, render_template, request
from printCodeGen import generate_macro_code
from commonUtils import read_excel_file, log
from flask_cors import CORS

app = Flask(__name__)

CORS(app)

@app.route('/', strict_slashes=False)
def home():
    return render_template('home.html')

@app.route('/print', methods=['GET', 'POST'])
def print_route():
    if request.method == 'POST':

        sheetnumber = request.form['sheetnumber']
        column = request.form['column']
        xlfile = request.files['xl']
        
        if xlfile and column and sheetnumber:
            
            if not sheetnumber.isnumeric():
                log("Please provide a numeric sheet number")
                return render_template("print.html", error="Please provide a numeric sheet number")
            if not column.isalpha():
                log("Please provide a valid column name")
                return render_template("print.html", error="Please provide a valid column name")

            sheetnumber = int(sheetnumber)
            column = column.upper()
            
            vals = read_excel_file(xlfile, column=column, sheet_name=sheetnumber)
            
            if isinstance(vals, ValueError):
                log(str(vals))
                return str(vals)

            elif isinstance(vals, Exception):
                log(str(vals))
                return render_template("print.html", error="Unknow error occured\n"+str(vals))

            
            code = generate_macro_code(vals)
             
            if isinstance(code, TypeError):
                return render_template("print.html", error="The quantity column contains non-numeric values\n"+str(code))


            elif isinstance(code, Exception):
                return render_template("print.html", error="Unknow error occured\n"+str(code))

            print(code)
            return render_template("print.html", code=code)
            
            
        else:
            return render_template("print.html", error="No excel file attached")

        
    return render_template('print.html')


@app.route('/testing')
def testing():
    return 'asjihdiohdua huei9wa eajsnjaiehw8 aihdajhw0d a90oaidji2ojeiojd aouw9-e a90uw09 assdhh'

if __name__ == '__main__':
    app.run(debug=True, port=5322)
    
