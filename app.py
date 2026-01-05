from flask import Flask, request, render_template, redirect, url_for, flash, session
import pandas as pd
import os
import tempfile
from openpyxl import load_workbook
from pyzbar import pyzbar
from PIL import Image
import pdfplumber
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import re
from datetime import datetime

app = Flask(__name__)
app.secret_key = 'your_secret_key_here'

# 用户配置
users = {
    '123abc': {'name': 'John', 'email': 'john@example.com'},
    '456def': {'name': 'Jane', 'email': 'jane@example.com'},
    '789xyz': {'name': 'Alex', 'email': 'alex@example.com'}
}

# 邮件配置
EMAIL_USER = os.environ.get('EMAIL_USER')
EMAIL_PASSWORD = os.environ.get('EMAIL_PASSWORD')
SMTP_SERVER = 'smtp.gmail.com'
SMTP_PORT = 587

# Excel 模板路径
EXCEL_TEMPLATE_PATH = 'transfer_template.xlsx'

def extract_text_from_image(image_path):
    """从图片中提取文本"""
    try:
        # 读取图片
        image = Image.open(image_path)
        
        # 尝试解码二维码
        decoded_objects = pyzbar.decode(image)
        if decoded_objects:
            return decoded_objects[0].data.decode('utf-8')
        
        # 如果没有二维码，返回空字符串
        return ""
    except Exception as e:
        print(f"图片处理错误: {e}")
        return ""

def extract_text_from_pdf(pdf_path):
    """从PDF中提取文本"""
    try:
        text = ""
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                text += page.extract_text() or ""
        return text
    except Exception as e:
        print(f"PDF处理错误: {e}")
        return ""

def parse_transfer_info(text):
    """解析提取的文本信息"""
    info = {
        'order_number': '',
        'date': '',
        'customer': '',
        'address': '',
        'items': []
    }
    
    # 简单的文本解析（根据实际格式调整）
    lines = text.split('\n')
    for line in lines:
        line = line.strip()
        if 'Order' in line or '订单' in line:
            # 提取订单号
            match = re.search(r'(\d+)', line)
            if match:
                info['order_number'] = match.group(1)
        elif 'Date' in line or '日期' in line:
            # 提取日期
            date_match = re.search(r'(\d{4}-\d{2}-\d{2}|\d{2}/\d{2}/\d{4})', line)
            if date_match:
                info['date'] = date_match.group(1)
        elif 'Customer' in line or '客户' in line:
            # 提取客户信息
            info['customer'] = line.split(':')[-1].strip()
        elif 'Address' in line or '地址' in line:
            # 提取地址信息
            info['address'] = line.split(':')[-1].strip()
    
    return info

def update_excel(info, username, tray_type, quantity):
    """更新Excel表格"""
    # 创建一个新的Excel文件
    df = pd.DataFrame({
        'Order Number': [info.get('order_number', '')],
        'Date': [info.get('date', '')],
        'Customer': [info.get('customer', '')],
        'Address': [info.get('address', '')],
        'Username': [username],
        'Tray Type': [tray_type],
        'Quantity': [quantity],
        'Status': ['Pending'],
        'Timestamp': [datetime.now().strftime('%Y-%m-%d %H:%M:%S')]
    })
    
    # 保存到临时文件
    temp_file = tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx')
    df.to_excel(temp_file.name, index=False)
    temp_file.close()
    
    return temp_file.name

def send_email_with_attachment(to_email, subject, body, attachment_path):
    """发送带附件的邮件"""
    try:
        msg = MIMEMultipart()
        msg['From'] = EMAIL_USER
        msg['To'] = to_email
        msg['Subject'] = subject
        
        msg.attach(MIMEText(body, 'plain'))
        
        with open(attachment_path, "rb") as attachment:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(attachment.read())
        
        encoders.encode_base64(part)
        part.add_header(
            'Content-Disposition',
            f'attachment; filename= {os.path.basename(attachment_path)}'
        )
        msg.attach(part)
        
        server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
        server.starttls()
        server.login(EMAIL_USER, EMAIL_PASSWORD)
        text = msg.as_string()
        server.sendmail(EMAIL_USER, to_email, text)
        server.quit()
        
        return True
    except Exception as e:
        print(f"邮件发送错误: {e}")
        return False

@app.route('/')
def index():
    return render_template('login.html')

@app.route('/login', methods=['POST'])
def login():
    username = request.form['username']
    if username in users:
        session['username'] = username
        session['user_info'] = users[username]
        return redirect(url_for('upload'))
    else:
        flash('Invalid username')
        return redirect(url_for('index'))

@app.route('/upload')
def upload():
    if 'username' not in session:
        return redirect(url_for('index'))
    return render_template('upload.html')

@app.route('/process', methods=['POST'])
def process():
    if 'username' not in session:
        return redirect(url_for('index'))
    
    username = session['username']
    user_info = session['user_info']
    
    if 'file' not in request.files:
        flash('No file selected')
        return redirect(url_for('upload'))
    
    file = request.files['file']
    if file.filename == '':
        flash('No file selected')
        return redirect(url_for('upload'))
    
    tray_type = request.form.get('tray_type')
    quantity = request.form.get('quantity')
    
    if not tray_type or not quantity:
        flash('Please select tray type and quantity')
        return redirect(url_for('upload'))
    
    # 保存上传的文件到临时位置
    temp_file = tempfile.NamedTemporaryFile(delete=False)
    file.save(temp_file.name)
    
    # 根据文件类型处理
    extracted_text = ""
    if file.filename.lower().endswith('.pdf'):
        extracted_text = extract_text_from_pdf(temp_file.name)
    else:
        extracted_text = extract_text_from_image(temp_file.name)
    
    # 解析信息
    info = parse_transfer_info(extracted_text)
    
    # 更新Excel
    excel_path = update_excel(info, username, tray_type, quantity)
    
    # 发送邮件
    subject = f"Transfer Order - {info.get('order_number', 'N/A')}"
    body = f"""
    New transfer order submitted by {username} ({user_info['email']})
    
    Order Number: {info.get('order_number', 'N/A')}
    Date: {info.get('date', 'N/A')}
    Customer: {info.get('customer', 'N/A')}
    Address: {info.get('address', 'N/A')}
    Tray Type: {tray_type}
    Quantity: {quantity}
    """
    
    success = send_email_with_attachment(user_info['email'], subject, body, excel_path)
    
    # 清理临时文件
    os.unlink(temp_file.name)
    os.unlink(excel_path)
    
    if success:
        flash('File processed successfully and email sent!')
    else:
        flash('File processed but email sending failed!')
    
    return redirect(url_for('upload'))

@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('index'))

if __name__ == '__main__':
    app.run(debug=True)
