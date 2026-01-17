#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
เว็บแอพพลิเคชั่นแปลงไฟล์ .xls เป็น CSV UTF-8
Web application to convert .xls files to CSV UTF-8 format

Developed by: นายปองดี ไชยจันดา (Pongdee Chaichanda)
Position: หัวหน้างานวัดผลและประเมินผล
Affiliation: กลุ่มบริหารวิชาการ โรงเรียนพยุห์วิทยา
Contact: Teacherpchai@gmail.com
Version: 1.0.0 (Beta)

Copyright (c) 2025 นายปองดี ไชยจันดา (Pongdee Chaichanda)
"""

import os
import tempfile
import zipfile
from pathlib import Path
from flask import Flask, render_template, request, send_file, jsonify, after_this_request
from werkzeug.utils import secure_filename
import pandas as pd
from io import StringIO, BytesIO

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024  # 100MB max file size
app.config['UPLOAD_FOLDER'] = tempfile.mkdtemp()

ALLOWED_EXTENSIONS = {'xls', 'xlsx'}


def allowed_file(filename):
    """ตรวจสอบว่าไฟล์ที่อัปโหลดเป็น .xls หรือ .xlsx"""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


def xls_to_csv_utf8(file_path):
    """
    แปลงไฟล์ .xls เป็น CSV UTF-8
    Returns: BytesIO object containing CSV data
    """
    try:
        # อ่านไฟล์ Excel โดยใช้ pandas
        try:
            # ลองอ่านเป็น Excel ก่อน
            df = pd.read_excel(file_path, sheet_name=0, engine=None, header=0)
        except Exception:
            # ถ้าไม่ได้ ลองอ่านเป็น HTML (กรณีที่ไฟล์เป็น HTML ที่บันทึกเป็น .xls)
            try:
                with open(file_path, 'rb') as f:
                    html_content = f.read()
                
                # ลอง decode ด้วย UTF-8 ก่อน
                try:
                    html_text = html_content.decode('utf-8')
                except UnicodeDecodeError:
                    # ลอง encoding อื่นๆ
                    encodings = ['cp874', 'tis-620', 'iso-8859-11', 'windows-874']
                    html_text = None
                    for enc in encodings:
                        try:
                            html_text = html_content.decode(enc)
                            break
                        except UnicodeDecodeError:
                            continue
                    
                    if html_text is None:
                        html_text = html_content.decode('utf-8', errors='replace')
                
                html_io = StringIO(html_text)
                dfs = pd.read_html(html_io, encoding='utf-8')
                if dfs:
                    df = dfs[0]
                else:
                    raise ValueError("ไม่พบตารางในไฟล์ HTML")
            except Exception as e:
                raise ValueError(f"ไม่สามารถอ่านไฟล์ได้: {str(e)}")
        
        # แปลงเป็น CSV UTF-8 with BOM
        output = BytesIO()
        csv_content = df.to_csv(index=False, header=True, encoding='utf-8-sig', sep=',')
        output.write(csv_content.encode('utf-8-sig'))
        output.seek(0)
        return output
    
    except Exception as e:
        raise Exception(f"เกิดข้อผิดพลาดในการแปลงไฟล์: {str(e)}")


@app.route('/')
def index():
    """หน้าแรก"""
    return render_template('index.html')


@app.route('/convert', methods=['POST'])
def convert_file():
    """แปลงไฟล์ที่อัปโหลด"""
    if 'files' not in request.files:
        return jsonify({'error': 'ไม่พบไฟล์ที่อัปโหลด'}), 400
    
    files = request.files.getlist('files')
    
    if not files or files[0].filename == '':
        return jsonify({'error': 'กรุณาเลือกไฟล์'}), 400
    
    # กรองเฉพาะไฟล์ที่อนุญาต
    valid_files = [f for f in files if allowed_file(f.filename)]
    
    if not valid_files:
        return jsonify({'error': 'กรุณาอัปโหลดไฟล์ .xls หรือ .xlsx เท่านั้น'}), 400
    
    # ถ้ามีไฟล์เดียว ส่งกลับไฟล์เดียว
    if len(valid_files) == 1:
        file = valid_files[0]
        filename = secure_filename(file.filename)
        
        # บันทึกไฟล์ชั่วคราว
        temp_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(temp_path)
        
        try:
            csv_data = xls_to_csv_utf8(temp_path)
            csv_filename = Path(filename).with_suffix('.csv').name
            
            @after_this_request
            def remove_file(response):
                try:
                    os.remove(temp_path)
                except Exception:
                    pass
                return response
            
            return send_file(
                csv_data,
                mimetype='text/csv',
                as_attachment=True,
                download_name=csv_filename
            )
        except Exception as e:
            if os.path.exists(temp_path):
                os.remove(temp_path)
            return jsonify({'error': str(e)}), 500
    
    # ถ้ามีหลายไฟล์ สร้าง ZIP
    else:
        zip_buffer = BytesIO()
        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
            temp_files = []
            for file in valid_files:
                filename = secure_filename(file.filename)
                temp_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
                file.save(temp_path)
                temp_files.append(temp_path)
                
                try:
                    csv_data = xls_to_csv_utf8(temp_path)
                    csv_filename = Path(filename).with_suffix('.csv').name
                    zip_file.writestr(csv_filename, csv_data.read())
                except Exception as e:
                    # ข้ามไฟล์ที่มีปัญหา
                    continue
            
            # ลบไฟล์ชั่วคราว
            for temp_path in temp_files:
                try:
                    os.remove(temp_path)
                except Exception:
                    pass
        
        zip_buffer.seek(0)
        return send_file(
            zip_buffer,
            mimetype='application/zip',
            as_attachment=True,
            download_name='converted_files.zip'
        )


if __name__ == '__main__':
    # สำหรับ local development
    app.run(debug=True, host='0.0.0.0', port=8000)
