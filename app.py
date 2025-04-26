from flask import Flask, request, send_file
import io
import openpyxl

app = Flask(__name__)

@app.route('/webhook', methods=['POST'])
def webhook():
    data = request.json
    output = create_excel(data)
    return send_file(output, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                     as_attachment=True, download_name="施工計画書.xlsx")

def create_excel(data):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "計画書"

    ws.append(["分類", data.get("分類", "")])
    ws.append(["名称", data.get("名称", "")])
    ws.append(["場所", data.get("場所", "")])
    ws.append(["工期", data.get("工期", "")])

    file = io.BytesIO()
    wb.save(file)
    file.seek(0)
    return file

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=10000)
