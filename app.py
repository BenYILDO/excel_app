from flask import Flask, render_template, request, jsonify, send_file
import pandas as pd
import requests
from concurrent.futures import ThreadPoolExecutor
import threading
import time
from openpyxl.styles import PatternFill
from io import BytesIO
import openpyxl
import os

app = Flask(__name__)

# SSL doğrulamasını kapatma (production'da açılmalı)
import urllib3
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

def check_url(url):
    try:
        start_time = time.time()
        response = requests.get(url, timeout=10, verify=False)
        response_time = time.time() - start_time
        
        # İçeriği kontrol et
        content = response.text.strip()
        
        # Header bilgilerine göre çalışıp çalışmadığını kontrol et
        content_type = response.headers.get('content-type', 'Bilinmiyor')
        server = response.headers.get('server', 'Bilinmiyor')
        
        # Eğer content-type ve server bilgileri varsa çalışıyor demektir
        is_working = content_type != 'Bilinmiyor' and server != 'Bilinmiyor'
        
        return {
            'url': url,
            'status': response.status_code,
            'working': is_working,
            'response_time': f"{response_time:.2f} saniye",
            'content_type': content_type,
            'server': server,
            'content': content[:200] if content else ''  # İçeriği sınırla
        }
    except requests.exceptions.SSLError:
        return {
            'url': url,
            'status': 'SSL Hatası',
            'working': False,
            'error': 'SSL sertifika hatası'
        }
    except requests.exceptions.ConnectionError:
        return {
            'url': url,
            'status': 'Bağlantı Hatası',
            'working': False,
            'error': 'Sunucuya bağlanılamadı'
        }
    except requests.exceptions.Timeout:
        return {
            'url': url,
            'status': 'Zaman Aşımı',
            'working': False,
            'error': 'Sunucu yanıt vermedi'
        }
    except requests.RequestException as e:
        return {
            'url': url,
            'status': 'Hata',
            'working': False,
            'error': str(e)
        }

def process_urls(urls):
    with ThreadPoolExecutor(max_workers=5) as executor:  # Worker sayısını düşürdük
        results = list(executor.map(check_url, urls))
    return results

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/check_urls', methods=['POST'])
def check_urls():
    if 'file' in request.files:
        file = request.files['file']
        if file.filename.endswith('.xlsx'):
            df = pd.read_excel(file)
            urls = []
            if 'A' in df.columns:
                urls.extend(df['A'].dropna().tolist())
            if 'C' in df.columns:
                urls.extend(df['C'].dropna().tolist())
    else:
        urls = request.form.get('urls', '').split('\n')
        urls = [url.strip() for url in urls if url.strip()]

    results = process_urls(urls)
    return jsonify(results)

@app.route('/read_excel', methods=['POST'])
def read_excel():
    if 'file' not in request.files:
        return jsonify({'error': 'Dosya yüklenmedi'}), 400
    
    file = request.files['file']
    if not file.filename.endswith('.xlsx'):
        return jsonify({'error': 'Sadece Excel (.xlsx) dosyaları desteklenir'}), 400
    
    try:
        df = pd.read_excel(file)
        result = {
            'columns': [str(col) for col in df.columns],
            'data': df.fillna('').values.tolist()
        }
        return jsonify(result)
    except Exception as e:
        return jsonify({'error': f'Excel okuma hatası: {str(e)}'}), 500

@app.route('/export_excel', methods=['POST'])
def export_excel():
    try:
        data = request.json
        wb = openpyxl.Workbook()
        ws = wb.active
        
        # Başlıkları ekle
        for col, header in enumerate(data['headers'], start=1):
            ws.cell(row=1, column=col, value=header)
        
        # Verileri ekle ve renkleri ayarla
        green_fill = PatternFill(start_color='90EE90', end_color='90EE90', fill_type='solid')
        red_fill = PatternFill(start_color='FFB6C1', end_color='FFB6C1', fill_type='solid')
        
        for row_idx, row_data in enumerate(data['rows'], start=2):
            for col_idx, cell_data in enumerate(row_data, start=1):
                cell = ws.cell(row=row_idx, column=col_idx, value=cell_data.get('text', ''))
                if cell_data.get('working') is True:
                    cell.fill = green_fill
                elif cell_data.get('working') is False:
                    cell.fill = red_fill
        
        excel_file = BytesIO()
        wb.save(excel_file)
        excel_file.seek(0)
        
        return send_file(
            excel_file,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name='link_kontrol_sonuclari.xlsx'
        )
    except Exception as e:
        return jsonify({'error': str(e)}), 500

# Development sunucusu için
if __name__ == '__main__':
    # Render.com için port ayarı
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
