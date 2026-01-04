# app.py - IAA休闲游戏成本收益测算Web应用
from flask import Flask, render_template, request, jsonify, Response
from calculations import calculate_metrics, export_daily_excel
import os
import json
from datetime import datetime

app = Flask(__name__)

# 项目保存目录
PROJECTS_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'projects')
# 数据保存目录
DATA_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'data')

# 确保项目目录存在
if not os.path.exists(PROJECTS_DIR):
    os.makedirs(PROJECTS_DIR)

"""将项目名称转换为安全的文件名，移除不安全字符"""
def get_safe_filename(project_name):
    unsafe_chars = ['/', '\\', ':', '*', '?', '"', '<', '>', '|']
    safe_name = project_name
    for char in unsafe_chars:
        safe_name = safe_name.replace(char, '_')
    return safe_name

"""根据项目名称获取项目文件的完整路径"""
def get_project_file_path(project_name):
    safe_name = get_safe_filename(project_name)
    return os.path.join(PROJECTS_DIR, f"{safe_name}.json")

"""验证项目请求数据"""
def validate_project_request(data):
    if not data:
        return None, (jsonify({'error': '未收到有效的请求数据'}), 400)
    
    project_name = data.get('project_name', '').strip()
    if not project_name:
        return None, (jsonify({'error': '项目名称不能为空'}), 400)
    
    return project_name, None

"""保存项目数据的辅助函数"""
def save_project_data(data):
    project_name = data.get('project_name', '').strip()
    if not project_name:
        return
    
    # 添加保存时间
    data['savedAt'] = datetime.now().isoformat()
    
    # 保存到文件
    file_path = get_project_file_path(project_name)
    with open(file_path, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

"""渲染主页面"""
@app.route('/')
def index():
    return render_template('index.html')

""" 计算 """
@app.route('/calculate', methods=['POST'])
def calculate():
    try:
        # 获取前端发送的JSON数据
        params = request.get_json()
        
        if not params:
            return jsonify({'error': '未收到有效的参数数据'}), 400
        
        # 调用核心计算函数
        results = calculate_metrics(params)
        
        # 计算完成后自动保存项目
        save_project_data(params)
        
        # 返回计算结果
        return jsonify(results)
    
    except Exception as e:
        return jsonify({'error': f'计算过程中发生错误: {str(e)}'}), 500

""" 导出每日数据 """
@app.route('/export_csv', methods=['GET'])
def export_csv():
    from urllib.parse import quote
    
    try:
        # 获取项目名称参数
        project_name = request.args.get('project_name', '')
        
        if not project_name:
            return jsonify({'error': '未提供项目名称'}), 400
        
        # 调用导出函数获取Excel二进制数据
        excel_data = export_daily_excel(project_name)
        
        # 对文件名进行URL编码，支持中文文件名
        filename = f'{project_name}_每日数据.xlsx'
        encoded_filename = quote(filename)
        
        # 返回Excel文件响应
        return Response(
            excel_data,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            headers={
                'Content-Disposition': f"attachment; filename*=UTF-8''{encoded_filename}",
                'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            }
        )
    
    except (FileNotFoundError, Exception) as e:
        return jsonify({'error': f'导出过程中发生错误: {str(e)}'}), 500


""" 保存项目到本地文件，保存为JSON文件 """
@app.route('/save_project', methods=['POST'])
def save_project():
    try:
        data = request.get_json()
        if not data:
            return jsonify({'error': '未收到有效的项目数据'}), 400
        
        project_name = data.get('project_name', '').strip()
        if not project_name:
            return jsonify({'error': '项目名称不能为空'}), 400
        
        save_project_data(data)
        
        return jsonify({'success': True, 'message': f'项目 "{project_name}" 已保存成功'})
    
    except Exception as e:
        return jsonify({'error': f'保存项目时发生错误: {str(e)}'}), 500

""" 列出所有已保存的项目 """
@app.route('/list_projects', methods=['GET'])
def list_projects():
    try:
        projects = []
        
        if os.path.exists(PROJECTS_DIR):
            for filename in os.listdir(PROJECTS_DIR):
                if filename.endswith('.json'):
                    file_path = os.path.join(PROJECTS_DIR, filename)
                    try:
                        with open(file_path, 'r', encoding='utf-8') as f:
                            data = json.load(f)
                            project_name = data.get('project_name', filename[:-5])
                            saved_at = data.get('savedAt', '')
                            projects.append({
                                'name': project_name,
                                'filename': filename,
                                'savedAt': saved_at
                            })
                    except:
                        # 如果文件读取失败，跳过
                        continue
        
        # 按保存时间倒序排列
        projects.sort(key=lambda x: x.get('savedAt', ''), reverse=True)
        
        return jsonify({'success': True, 'projects': projects})
    
    except Exception as e:
        return jsonify({'error': f'获取项目列表时发生错误: {str(e)}'}), 500

""" 加载指定项目 """
@app.route('/load_project', methods=['POST'])
def load_project():
    try:
        project_name, error = validate_project_request(request.get_json())
        if error:
            return error
        
        file_path = get_project_file_path(project_name)
        
        if not os.path.exists(file_path):
            return jsonify({'error': f'项目 "{project_name}" 不存在'}), 404
        
        with open(file_path, 'r', encoding='utf-8') as f:
            project_data = json.load(f)
        
        return jsonify({'success': True, 'data': project_data})
    
    except Exception as e:
        return jsonify({'error': f'加载项目时发生错误: {str(e)}'}), 500

""" 删除指定项目 """
@app.route('/delete_project', methods=['POST'])
def delete_project():
    try:
        project_name, error = validate_project_request(request.get_json())
        if error:
            return error
        
        file_path = get_project_file_path(project_name)
        
        if not os.path.exists(file_path):
            return jsonify({'error': f'项目 "{project_name}" 不存在'}), 404
        
        # 删除项目JSON文件
        os.remove(file_path)
        
        # 同时删除data文件夹中对应的pkl文件
        safe_name = get_safe_filename(project_name)
        pkl_file_path = os.path.join(DATA_DIR, f"{safe_name}_data.pkl")
        if os.path.exists(pkl_file_path):
            os.remove(pkl_file_path)
        
        return jsonify({'success': True, 'message': f'项目 "{project_name}" 已删除'})
    
    except Exception as e:
        return jsonify({'error': f'删除项目时发生错误: {str(e)}'}), 500


if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)
