from flask import Flask, request, render_template, jsonify, send_file
import subprocess
import os
import threading

app = Flask(__name__)

EXCEL_FILE = "ExcelFiles/results.xlsx"

@app.route('/')
def index():
    return render_template('index.html')

def run_scraper_thread(college, year, branch, low, high, semc):
    cmd = ["python", "scraper.py", college, year, branch, low, high, semc]
    subprocess.run(cmd)

@app.route('/run-scraper', methods=['POST'])
def run_scraper():
    data = request.form
    college = data.get('college')
    year = data.get('year')
    branch = data.get('branch')
    low = data.get('low')
    high = data.get('high')
    semc = data.get('semc')
    # Delete old file before starting new scrape
    if os.path.exists(EXCEL_FILE):
        os.remove(EXCEL_FILE)
    thread = threading.Thread(target=run_scraper_thread, args=(college, year, branch, low, high, semc))
    thread.start()
    return jsonify({"output": "Scraping started. Please wait..."})

@app.route('/download', methods=['GET', 'HEAD'])
def download():
    if os.path.exists(EXCEL_FILE):
        if request.method == 'HEAD':
            return '', 200
        return send_file(EXCEL_FILE, as_attachment=True)
    else:
        return "Excel file not found.", 404

if __name__ == '__main__':
    app.run(debug=True)