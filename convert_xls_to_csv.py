#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
โปรแกรมแปลงไฟล์ .xls เป็น CSV UTF-8 (Comma delimited)
Convert .xls files to CSV UTF-8 (Comma delimited) format

Developed by: นายปองดี ไชยจันดา (Pongdee Chaichanda)
Position: หัวหน้างานวัดผลและประเมินผล
Affiliation: กลุ่มบริหารวิชาการ โรงเรียนพยุห์วิทยา
Contact: Teacherpchai@gmail.com
Version: 1.0.0 (Beta)

Copyright (c) 2025 นายปองดี ไชยจันดา (Pongdee Chaichanda)
"""

import os
import sys
import pandas as pd
from pathlib import Path
from io import StringIO


def xls_to_csv_utf8(input_file, output_file=None):
    """
    แปลงไฟล์ .xls เป็น CSV UTF-8
    
    Args:
        input_file: path ของไฟล์ .xls
        output_file: path ของไฟล์ .csv ที่ต้องการ (ถ้าไม่ระบุจะใช้ชื่อเดียวกับไฟล์ต้นฉบับ)
    """
    try:
        # ตรวจสอบว่าไฟล์ input มีอยู่จริง
        if not os.path.exists(input_file):
            raise FileNotFoundError(f"ไม่พบไฟล์: {input_file}")
        
        # ถ้าไม่ระบุ output_file ให้ใช้ชื่อเดียวกับ input แต่เปลี่ยนนามสกุลเป็น .csv
        if output_file is None:
            input_path = Path(input_file)
            output_file = input_path.with_suffix('.csv')
        
        # ตรวจสอบประเภทไฟล์จริงๆ
        input_path = Path(input_file)
        
        # อ่านไฟล์ Excel โดยใช้ pandas (รองรับทั้ง .xls, .xlsx และ HTML ที่บันทึกเป็น .xls)
        # pandas จะพยายามอ่านเป็น Excel ก่อน ถ้าไม่ได้จะลองอ่านเป็น HTML
        try:
            # ลองอ่านเป็น Excel ก่อน (ให้ pandas ตรวจจับ header อัตโนมัติ)
            df = pd.read_excel(input_file, sheet_name=0, engine=None, header=0)
        except Exception:
            # ถ้าไม่ได้ ลองอ่านเป็น HTML
            try:
                # อ่านไฟล์เป็น bytes ก่อน แล้ว decode เป็น UTF-8 เพื่อให้แน่ใจว่า encoding ถูกต้อง
                with open(input_file, 'rb') as f:
                    html_content = f.read()
                
                # ลอง decode ด้วย UTF-8 ก่อน
                try:
                    html_text = html_content.decode('utf-8')
                except UnicodeDecodeError:
                    # ถ้า UTF-8 ไม่ได้ ลอง encoding อื่นๆ
                    encodings = ['cp874', 'tis-620', 'iso-8859-11', 'windows-874']
                    html_text = None
                    for enc in encodings:
                        try:
                            html_text = html_content.decode(enc)
                            break
                        except UnicodeDecodeError:
                            continue
                    
                    if html_text is None:
                        # ถ้ายังไม่ได้ ให้ใช้ errors='replace' หรือ 'ignore'
                        html_text = html_content.decode('utf-8', errors='replace')
                
                # อ่าน HTML จาก string โดยใช้ StringIO
                html_io = StringIO(html_text)
                dfs = pd.read_html(html_io, encoding='utf-8')
                if dfs:
                    df = dfs[0]  # ใช้ table แรก
                else:
                    raise ValueError("ไม่พบตารางในไฟล์ HTML")
            except Exception as e:
                raise ValueError(f"ไม่สามารถอ่านไฟล์ได้: {str(e)}")
        
        # บันทึกเป็น CSV UTF-8 (Comma delimited) พร้อม BOM เพื่อให้ Excel รู้ว่าเป็น UTF-8
        # header=True เพื่อเก็บหัวคอลัมน์
        df.to_csv(output_file, index=False, header=True, encoding='utf-8-sig', sep=',')
        
        print(f"✓ แปลงไฟล์สำเร็จ: {input_file} → {output_file}")
        return output_file
    
    except Exception as e:
        print(f"✗ เกิดข้อผิดพลาด: {str(e)}", file=sys.stderr)
        return None


def batch_convert(input_directory, output_directory=None):
    """
    แปลงไฟล์ .xls ทั้งหมดในโฟลเดอร์
    
    Args:
        input_directory: path ของโฟลเดอร์ที่มีไฟล์ .xls
        output_directory: path ของโฟลเดอร์ที่ต้องการเก็บไฟล์ .csv (ถ้าไม่ระบุจะใช้โฟลเดอร์เดียวกับ input)
    """
    input_path = Path(input_directory)
    
    if not input_path.exists():
        print(f"✗ ไม่พบโฟลเดอร์: {input_directory}", file=sys.stderr)
        return
    
    if output_directory is None:
        output_path = input_path
    else:
        output_path = Path(output_directory)
        output_path.mkdir(parents=True, exist_ok=True)
    
    # ค้นหาไฟล์ .xls ทั้งหมด
    xls_files = list(input_path.glob('*.xls'))
    
    if not xls_files:
        print(f"ไม่พบไฟล์ .xls ในโฟลเดอร์: {input_directory}")
        return
    
    print(f"พบไฟล์ .xls ทั้งหมด {len(xls_files)} ไฟล์")
    
    for xls_file in xls_files:
        output_file = output_path / xls_file.with_suffix('.csv').name
        xls_to_csv_utf8(str(xls_file), str(output_file))


def main():
    """ฟังก์ชันหลัก"""
    if len(sys.argv) < 2:
        print("วิธีใช้:")
        print(f"  python {sys.argv[0]} <ไฟล์.xls> [ไฟล์.csv]")
        print(f"  python {sys.argv[0]} --batch <โฟลเดอร์> [โฟลเดอร์ผลลัพธ์]")
        print("\nตัวอย่าง:")
        print(f"  python {sys.argv[0]} data.xls")
        print(f"  python {sys.argv[0]} data.xls output.csv")
        print(f"  python {sys.argv[0]} --batch ./input_folder")
        print(f"  python {sys.argv[0]} --batch ./input_folder ./output_folder")
        sys.exit(1)
    
    if sys.argv[1] == '--batch':
        # แปลงทั้งโฟลเดอร์
        input_dir = sys.argv[2] if len(sys.argv) > 2 else '.'
        output_dir = sys.argv[3] if len(sys.argv) > 3 else None
        batch_convert(input_dir, output_dir)
    else:
        # แปลงไฟล์เดียว
        input_file = sys.argv[1]
        output_file = sys.argv[2] if len(sys.argv) > 2 else None
        xls_to_csv_utf8(input_file, output_file)


if __name__ == '__main__':
    main()
