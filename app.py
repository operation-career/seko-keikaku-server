from flask import Flask, request, jsonify, send_file
import io
import openpyxl
import requests

app = Flask(__name__)

# Webhookエンドポイント
@app.route('/webhook', methods=['POST'])
def webhook():
    try:
        # miiboから送られてきたJSONデータを受け取る
        data = request.get_json()

        # 必要なデータを取得
        file_url = data.get('file_url')
        vehicle_equipment = data.get('使用車両備品')
        work_description = data.get('主な作業内容')

        # ファイル（PDF）ダウンロードして保存
        file_response = requests.get(file_url)
        if file_response.status_code != 200:
            raise Exception("PDFファイルのダウンロードに失敗しました。")

        # ダウンロードしたPDFファイルの中身（バイナリデータ）
        pdf_content = file_response.content

        # --- 仮設定（本来はここでPDF解析して工事名・場所・工期を自動取得する想定） ---
        工事名 = "仮 工事名"
        工事場所 = "仮 工事場所"
        工期 = "仮 工期"
        # --------------------------------------------------------------------------

        # Excelファイルを作成
        output = create_excel(工事名, 工事場所, 工期, vehicle_equipment, work_description)

        return send_file(output, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                         as_attachment=True, download_name="施工計画書.xlsx")

    except Exception as e:
        print(f"エラー: {str(e)}")
        return jsonify({"status": "error", "message": str(e)}), 500

# Excelファイルを作成する関数
def create_excel(工事名, 工事場所, 工期, vehicle_equipment, work_description):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "計画書"

    # 情報を順番に書き込む
    ws.append(["工事名", 工事名])
    ws.append(["工事場所", 工事場所])
    ws.append(["工期", 工期])
    ws.append(["使用車両・備品", vehicle_equipment])
    ws.append(["主な作業内容", work_description])

    # メモリ上にExcelファイルを作成
    file = io.BytesIO()
    wb.save(file)
    file.seek(0)
    return file

# サーバー起動
if __name__ == '__main__':
    app.run(host='0.0.0.0', port=10000)
