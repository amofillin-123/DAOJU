from flask import Flask, request, send_file, render_template_string
import os
from script_generator import ScriptGenerator
import tempfile

# 创建Flask应用实例
application = Flask(__name__)
app = application  # 为了兼容性，同时提供app变量

# 简单的HTML模板
HTML_TEMPLATE = '''
<!DOCTYPE html>
<html>
<head>
    <title>PPT转换工具</title>
    <style>
        body {
            font-family: Arial, sans-serif;
            max-width: 800px;
            margin: 0 auto;
            padding: 20px;
        }
        .upload-form {
            border: 2px dashed #ccc;
            padding: 20px;
            text-align: center;
            margin: 20px 0;
        }
        .button {
            background-color: #4CAF50;
            color: white;
            padding: 10px 20px;
            border: none;
            border-radius: 4px;
            cursor: pointer;
        }
        .button:hover {
            background-color: #45a049;
        }
    </style>
</head>
<body>
    <h1>PPT转换工具</h1>
    <div class="upload-form">
        <form action="/convert" method="post" enctype="multipart/form-data">
            <p>请选择要转换的PPT文件 (.pptx)</p>
            <input type="file" name="file" accept=".pptx" required>
            <br><br>
            <input type="submit" value="转换" class="button">
        </form>
    </div>
</body>
</html>
'''

@app.route('/')
def index():
    return render_template_string(HTML_TEMPLATE)

@app.route('/convert', methods=['POST'])
def convert():
    if 'file' not in request.files:
        return '没有上传文件', 400
    
    file = request.files['file']
    if file.filename == '':
        return '没有选择文件', 400
    
    if not file.filename.endswith('.pptx'):
        return '请上传.pptx文件', 400

    # 创建临时目录存储文件
    with tempfile.TemporaryDirectory() as temp_dir:
        # 保存上传的文件
        pptx_path = os.path.join(temp_dir, file.filename)
        file.save(pptx_path)
        
        # 初始化转换器
        generator = ScriptGenerator()
        
        try:
            # 生成PDF文件名
            pdf_filename = os.path.splitext(file.filename)[0] + '_拍摄需求.pdf'
            pdf_path = os.path.join(temp_dir, pdf_filename)
            
            # 处理文件
            generator.process_file(pptx_path)
            
            # 返回生成的PDF文件
            return send_file(
                pdf_path,
                as_attachment=True,
                download_name=pdf_filename,
                mimetype='application/pdf'
            )
            
        except Exception as e:
            return f'转换过程中发生错误: {str(e)}', 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 8080))) 