"""Excel图片提取工具 - Flask主程序"""

import os
import json
import uuid
import shutil
import tempfile

from flask import Flask, render_template, request, jsonify, send_file, session, redirect
from werkzeug.utils import secure_filename

from services.excel_parser import (
    convert_xls_to_xlsx,
    get_sheet_names,
    get_headers,
    locate_images,
)
from services.image_extractor import ExtractConfig, extract_images

app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY', 'excel-image-extractor-secret')

# 访问密码（从环境变量读取，默认 DIY123456）
APP_PASSWORD = os.environ.get('APP_PASSWORD', 'DIY123456')

# 配置
UPLOAD_FOLDER = os.path.join(os.path.dirname(__file__), 'temp', 'uploads')
OUTPUT_FOLDER = os.path.join(os.path.dirname(__file__), 'temp', 'outputs')
STORE_FILE = os.path.join(os.path.dirname(__file__), 'temp', 'file_store.json')
ALLOWED_EXTENSIONS = {'xlsx', 'xls'}
MAX_FILE_SIZE = 1024 * 1024 * 1024  # 1GB

# 允许大文件上传
app.config['MAX_CONTENT_LENGTH'] = MAX_FILE_SIZE

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)


def _load_store() -> dict:
    """从JSON文件加载文件存储"""
    if os.path.exists(STORE_FILE):
        with open(STORE_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    return {}


def _save_store(store: dict):
    """保存文件存储到JSON"""
    os.makedirs(os.path.dirname(STORE_FILE), exist_ok=True)
    with open(STORE_FILE, 'w', encoding='utf-8') as f:
        json.dump(store, f, ensure_ascii=False)


def allowed_file(filename: str) -> bool:
    """检查文件扩展名是否合法"""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


@app.before_request
def require_login():
    """全局登录校验：静态资源和登录页免验证"""
    if request.path == '/login' or request.path.startswith('/static/'):
        return
    if 'authenticated' not in session:
        # API 请求返回 401 JSON，页面请求重定向到登录页
        if request.path.startswith('/api/') or request.is_json:
            return jsonify({"error": "未登录，请先访问 /login 输入密码"}), 401
        return redirect('/login')


@app.route('/login', methods=['GET', 'POST'])
def login():
    """简易登录页"""
    if request.method == 'POST':
        pwd = request.form.get('password', '').strip()
        if pwd == APP_PASSWORD:
            session['authenticated'] = True
            return redirect('/')
        return render_template_string(
            '<p style="color:#c45d3e;text-align:center;margin-top:20px;">密码错误</p>'
            '<p style="text-align:center;"><a href="/login">返回重试</a></p>'
        )
    return render_template_string('''
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>登录 - Excel 图片提取</title>
    <style>
        * { box-sizing: border-box; }
        body {
            margin: 0; padding: 0;
            font-family: "PingFang SC","Microsoft YaHei",sans-serif;
            background: #f3f2ef;
            display: flex; align-items: center; justify-content: center;
            min-height: 100vh;
        }
        .card {
            background: #fff; border-radius: 14px;
            padding: 40px 36px; width: 100%; max-width: 360px;
            box-shadow: 0 1px 2px rgba(0,0,0,.04), 0 4px 16px rgba(0,0,0,.04);
            text-align: center;
        }
        h2 { margin: 0 0 8px; font-size: 20px; color: #1a1a1a; }
        p.desc { margin: 0 0 24px; font-size: 13px; color: #8a8a8a; }
        input[type="password"] {
            width: 100%; padding: 12px 14px;
            border: 1.5px solid #e5e3df; border-radius: 10px;
            font-size: 14px; margin-bottom: 16px;
        }
        input[type="password"]:focus {
            outline: none; border-color: #c45d3e;
        }
        button {
            width: 100%; padding: 12px;
            background: #c45d3e; color: #fff;
            border: none; border-radius: 10px;
            font-size: 15px; font-weight: 600; cursor: pointer;
        }
        button:hover { opacity: .92; }
    </style>
</head>
<body>
    <div class="card">
        <h2>Excel 图片提取</h2>
        <p class="desc">请输入访问密码</p>
        <form method="post">
            <input type="password" name="password" placeholder="密码" autofocus>
            <button type="submit">进入</button>
        </form>
    </div>
</body>
</html>
''')


@app.route('/')
def index():
    """主页面"""
    return render_template('index.html')


@app.route('/api/upload', methods=['POST'])
def upload_file():
    """上传Excel文件并解析列名"""
    if 'file' not in request.files:
        return jsonify({"error": "未选择文件"}), 400

    file = request.files['file']
    if file.filename == '':
        return jsonify({"error": "未选择文件"}), 400

    if not allowed_file(file.filename):
        return jsonify({"error": "仅支持 .xlsx 和 .xls 格式"}), 400

    try:
        # 保存上传文件
        file_id = str(uuid.uuid4())[:8]
        ext = file.filename.rsplit('.', 1)[1].lower()
        filename = secure_filename(f"{file_id}.{ext}")
        filepath = os.path.join(UPLOAD_FOLDER, filename)
        file.save(filepath)

        # 如果是.xls，尝试转换为.xlsx
        if ext == 'xls':
            xlsx_path = convert_xls_to_xlsx(filepath)
            if xlsx_path:
                filepath = xlsx_path
                ext = 'xlsx'
            else:
                return jsonify({
                    "error": ".xls格式需要安装LibreOffice进行转换，请上传.xlsx格式或安装LibreOffice"
                }), 400

        # 解析Sheet名
        sheets = get_sheet_names(filepath)

        # 解析第一个Sheet的列名
        headers = get_headers(filepath, sheets[0]) if sheets else []

        # 获取第一个Sheet的图片数量
        image_count = len(locate_images(filepath, sheets[0])) if sheets else 0

        # 持久化文件信息
        store = _load_store()
        store[file_id] = {
            "filepath": filepath,
            "original_name": file.filename,
            "sheets": sheets,
        }
        _save_store(store)

        return jsonify({
            "file_id": file_id,
            "original_name": file.filename,
            "sheets": sheets,
            "default_sheet": sheets[0] if sheets else None,
            "headers": headers,
            "image_count": image_count,
        })

    except Exception as e:
        return jsonify({"error": f"文件解析失败: {str(e)}"}), 500


@app.route('/api/sheet-info', methods=['POST'])
def sheet_info():
    """获取指定Sheet的列名和图片信息"""
    data = request.get_json()
    file_id = data.get('file_id')
    sheet_name = data.get('sheet_name')

    store = _load_store()
    if file_id not in store:
        return jsonify({"error": "文件已过期，请重新上传"}), 400

    try:
        filepath = store[file_id]['filepath']
        headers = get_headers(filepath, sheet_name)
        image_count = len(locate_images(filepath, sheet_name))

        return jsonify({
            "headers": headers,
            "image_count": image_count,
        })
    except Exception as e:
        return jsonify({"error": f"Sheet解析失败: {str(e)}"}), 500


@app.route('/api/extract', methods=['POST'])
def extract():
    """执行图片提取"""
    data = request.get_json()
    file_id = data.get('file_id')

    store = _load_store()
    if file_id not in store:
        return jsonify({"error": "文件已过期，请重新上传"}), 400

    file_info = store[file_id]
    sheet_name = data.get('sheet_name', file_info['sheets'][0])
    image_col = data.get('image_col')
    name_col = data.get('name_col')
    prefix_col = data.get('prefix_col')
    manual_prefix = data.get('manual_prefix', '')

    if not image_col or not name_col:
        return jsonify({"error": "请选择图片列和商品名称列"}), 400

    try:
        task_id = str(uuid.uuid4())[:8]
        output_dir = os.path.join(OUTPUT_FOLDER, task_id)

        config = ExtractConfig(
            file_path=file_info['filepath'],
            sheet_name=sheet_name,
            image_col=int(image_col),
            name_col=int(name_col),
            prefix_col=int(prefix_col) if prefix_col else None,
            manual_prefix=manual_prefix,
        )

        result = extract_images(config, output_dir)

        # 持久化结果
        store = _load_store()
        store[task_id] = {
            "filepath": file_info['filepath'],
            "zip_path": result.zip_path,
            "result": {
                "total": result.total_images,
                "extracted": result.extracted,
                "skipped": result.skipped,
                "errors": result.errors,
            }
        }
        _save_store(store)

        return jsonify({
            "task_id": task_id,
            "total": result.total_images,
            "extracted": result.extracted,
            "skipped": result.skipped,
            "errors": result.errors,
            "download_url": f"/api/download/{task_id}" if result.zip_path else None,
        })

    except Exception as e:
        return jsonify({"error": f"图片提取失败: {str(e)}"}), 500


@app.route('/api/download/<task_id>')
def download(task_id):
    """下载ZIP文件"""
    store = _load_store()
    if task_id not in store:
        return jsonify({"error": "文件已过期"}), 404

    zip_path = store[task_id].get('zip_path')
    if not zip_path or not os.path.exists(zip_path):
        return jsonify({"error": "文件不存在"}), 404

    return send_file(
        zip_path,
        as_attachment=True,
        download_name=f"提取图片_{task_id}.zip",
        mimetype='application/zip'
    )


if __name__ == '__main__':
    app.run(debug=False, host='0.0.0.0', port=5001)
