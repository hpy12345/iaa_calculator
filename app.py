# app.py - IAA休闲游戏成本收益测算Web应用
import os
import re
import json
from datetime import datetime
from urllib.parse import quote

from flask import Flask, render_template, request, jsonify, Response
from calculations import calculate_metrics, export_daily_excel

app = Flask(__name__)

# ── 目录配置 ──────────────────────────────────────────────────────────────────
_BASE_DIR = os.path.dirname(os.path.abspath(__file__))

PROJECTS_DIR      = os.path.join(_BASE_DIR, 'projects')       # 项目保存目录
DATA_DIR          = os.path.join(_BASE_DIR, 'data')            # 计算数据缓存目录
ROI_DATA_DIR      = os.path.join(_BASE_DIR, 'roi_data')        # ROI曲线保存目录
RETENTION_DATA_DIR = os.path.join(_BASE_DIR, 'retention_data') # 留存率曲线保存目录

# 曲线类型 → 存储目录映射
_CURVE_DIR_MAP = {
    'roi':       ROI_DATA_DIR,
    'retention': RETENTION_DATA_DIR,
}

# 确保各目录存在
for _dir in [PROJECTS_DIR, DATA_DIR, ROI_DATA_DIR, RETENTION_DATA_DIR]:
    os.makedirs(_dir, exist_ok=True)

# ── 文件名 / 路径工具函数 ──────────────────────────────────────────────────────

# 不安全字符的正则，用于文件名清理
_UNSAFE_CHARS_RE = re.compile(r'[/\\:*?"<>|]')


def get_safe_filename(project_name: str) -> str:
    """将项目名称转换为安全的文件名，移除不安全字符。"""
    return _UNSAFE_CHARS_RE.sub('_', project_name)


def get_project_file_path(project_name: str) -> str:
    """根据项目名称获取项目 JSON 文件的完整路径。"""
    return os.path.join(PROJECTS_DIR, f"{get_safe_filename(project_name)}.json")


def _get_curve_dir(curve_type: str) -> str | None:
    """根据曲线类型返回对应的存储目录，类型无效时返回 None。"""
    return _CURVE_DIR_MAP.get(curve_type)


def _get_curve_file_path(curve_type: str, curve_id: str) -> str | None:
    """根据曲线类型和 ID 获取曲线文件路径，类型无效时返回 None。"""
    curve_dir = _get_curve_dir(curve_type)
    if not curve_dir:
        return None
    return os.path.join(curve_dir, f"{curve_id}.json")

# ── 请求验证 / 数据持久化工具函数 ─────────────────────────────────────────────

def validate_project_request(data: dict | None):
    """验证项目请求数据，返回 (project_name, None) 或 (None, error_response)。"""
    if not data:
        return None, (jsonify({'error': '未收到有效的请求数据'}), 400)

    project_name = data.get('project_name', '').strip()
    if not project_name:
        return None, (jsonify({'error': '项目名称不能为空'}), 400)

    return project_name, None


def save_project_data(data: dict) -> None:
    """将项目数据持久化到 JSON 文件，若项目名称为空则直接返回。"""
    project_name = data.get('project_name', '').strip()
    if not project_name:
        return

    data['savedAt'] = datetime.now().isoformat()
    file_path = get_project_file_path(project_name)
    with open(file_path, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def _load_curves_from_dir(curve_dir: str) -> list:
    """读取指定目录下所有曲线 JSON 文件，返回曲线数据列表（跳过损坏文件）。"""
    curves = []
    if not os.path.exists(curve_dir):
        return curves
    for filename in os.listdir(curve_dir):
        if not filename.endswith('.json'):
            continue
        file_path = os.path.join(curve_dir, filename)
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                curves.append(json.load(f))
        except Exception:
            continue
    return curves

# ── 路由：页面 ────────────────────────────────────────────────────────────────

@app.route('/')
def index():
    """渲染主页面。"""
    return render_template('index.html')

# ── 路由：计算 ────────────────────────────────────────────────────────────────

@app.route('/calculate', methods=['POST'])
def calculate():
    """接收前端参数，执行核心计算并自动保存项目，返回计算结果。"""
    try:
        params = request.get_json()
        if not params:
            return jsonify({'error': '未收到有效的参数数据'}), 400

        results = calculate_metrics(params)
        save_project_data(params)
        return jsonify(results)

    except Exception as e:
        return jsonify({'error': f'计算过程中发生错误: {str(e)}'}), 500

# ── 路由：导出 ────────────────────────────────────────────────────────────────

@app.route('/export_csv', methods=['GET'])
def export_csv():
    """导出指定项目的每日数据为 Excel 文件并返回下载响应。"""
    try:
        project_name = request.args.get('project_name', '')
        if not project_name:
            return jsonify({'error': '未提供项目名称'}), 400

        excel_data = export_daily_excel(project_name)

        filename = f'{project_name}_每日数据.xlsx'
        encoded_filename = quote(filename)

        return Response(
            excel_data,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            headers={
                'Content-Disposition': f"attachment; filename*=UTF-8''{encoded_filename}",
                'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            }
        )

    except Exception as e:
        return jsonify({'error': f'导出过程中发生错误: {str(e)}'}), 500

# ── 路由：项目管理 ────────────────────────────────────────────────────────────

@app.route('/save_project', methods=['POST'])
def save_project():
    """保存项目数据到本地 JSON 文件。"""
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


@app.route('/list_projects', methods=['GET'])
def list_projects():
    """列出所有已保存的项目，按保存时间倒序返回。"""
    try:
        projects = []
        for filename in os.listdir(PROJECTS_DIR):
            if not filename.endswith('.json'):
                continue
            file_path = os.path.join(PROJECTS_DIR, filename)
            try:
                with open(file_path, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                projects.append({
                    'name':    data.get('project_name', filename[:-5]),
                    'filename': filename,
                    'savedAt': data.get('savedAt', ''),
                })
            except Exception:
                continue

        projects.sort(key=lambda x: x.get('savedAt', ''), reverse=True)
        return jsonify({'success': True, 'projects': projects})

    except Exception as e:
        return jsonify({'error': f'获取项目列表时发生错误: {str(e)}'}), 500


@app.route('/load_project', methods=['POST'])
def load_project():
    """加载指定项目的数据并返回。"""
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


@app.route('/delete_project', methods=['POST'])
def delete_project():
    """删除指定项目的 JSON 文件及对应的数据缓存文件。"""
    try:
        project_name, error = validate_project_request(request.get_json())
        if error:
            return error

        file_path = get_project_file_path(project_name)
        if not os.path.exists(file_path):
            return jsonify({'error': f'项目 "{project_name}" 不存在'}), 404

        os.remove(file_path)

        # 同时删除 data 目录中对应的 pkl 缓存文件
        pkl_path = os.path.join(DATA_DIR, f"{get_safe_filename(project_name)}_data.pkl")
        if os.path.exists(pkl_path):
            os.remove(pkl_path)

        return jsonify({'success': True, 'message': f'项目 "{project_name}" 已删除'})

    except Exception as e:
        return jsonify({'error': f'删除项目时发生错误: {str(e)}'}), 500

# ── 路由：曲线管理 ────────────────────────────────────────────────────────────

@app.route('/save_curve', methods=['POST'])
def save_curve():
    """保存曲线数据（ROI 或留存率）到对应目录的 JSON 文件。"""
    try:
        data = request.get_json()
        if not data:
            return jsonify({'error': '未收到有效的曲线数据'}), 400

        curve_type = data.get('type', '').strip()
        if curve_type not in _CURVE_DIR_MAP:
            return jsonify({'error': '无效的曲线类型，必须为 roi 或 retention'}), 400

        curve_id = data.get('id', '').strip()
        if not curve_id:
            return jsonify({'error': '曲线ID不能为空'}), 400

        curve_name = data.get('name', '').strip()
        if not curve_name:
            return jsonify({'error': '曲线名称不能为空'}), 400

        points = data.get('points')
        if not isinstance(points, list):
            return jsonify({'error': '曲线数据点格式无效'}), 400

        now = datetime.now().isoformat()
        curve_data = {
            'id':        curve_id,
            'name':      curve_name,
            'type':      curve_type,
            'points':    points,
            'createdAt': data.get('createdAt', now),
            'updatedAt': now,
        }

        file_path = _get_curve_file_path(curve_type, curve_id)
        with open(file_path, 'w', encoding='utf-8') as f:
            json.dump(curve_data, f, ensure_ascii=False, indent=2)

        return jsonify({'success': True, 'message': f'曲线 "{curve_name}" 已保存成功', 'curve': curve_data})

    except Exception as e:
        return jsonify({'error': f'保存曲线时发生错误: {str(e)}'}), 500


@app.route('/list_curves', methods=['GET'])
def list_curves():
    """列出指定类型的所有曲线，按更新时间倒序返回。"""
    try:
        curve_type = request.args.get('type', '').strip()
        if curve_type not in _CURVE_DIR_MAP:
            return jsonify({'error': '无效的曲线类型，必须为 roi 或 retention'}), 400

        curves = _load_curves_from_dir(_get_curve_dir(curve_type))
        curves.sort(key=lambda x: x.get('updatedAt', ''))
        return jsonify({'success': True, 'curves': curves})

    except Exception as e:
        return jsonify({'error': f'获取曲线列表时发生错误: {str(e)}'}), 500


@app.route('/list_all_curves', methods=['GET'])
def list_all_curves():
    """列出所有类型（roi + retention）的曲线，按更新时间倒序返回。"""
    try:
        all_curves = []
        for curve_dir in _CURVE_DIR_MAP.values():
            all_curves.extend(_load_curves_from_dir(curve_dir))

        all_curves.sort(key=lambda x: x.get('updatedAt', ''))
        return jsonify({'success': True, 'curves': all_curves})

    except Exception as e:
        return jsonify({'error': f'获取曲线列表时发生错误: {str(e)}'}), 500


@app.route('/delete_curve', methods=['POST'])
def delete_curve():
    """删除指定曲线的 JSON 文件。"""
    try:
        data = request.get_json()
        if not data:
            return jsonify({'error': '未收到有效的请求数据'}), 400

        curve_type = data.get('type', '').strip()
        if curve_type not in _CURVE_DIR_MAP:
            return jsonify({'error': '无效的曲线类型，必须为 roi 或 retention'}), 400

        curve_id = data.get('id', '').strip()
        if not curve_id:
            return jsonify({'error': '曲线ID不能为空'}), 400

        file_path = _get_curve_file_path(curve_type, curve_id)
        if not file_path or not os.path.exists(file_path):
            return jsonify({'error': '曲线不存在'}), 404

        os.remove(file_path)
        return jsonify({'success': True, 'message': '曲线已删除'})

    except Exception as e:
        return jsonify({'error': f'删除曲线时发生错误: {str(e)}'}), 500


@app.route('/rename_curve', methods=['POST'])
def rename_curve():
    """重命名指定曲线。"""
    try:
        data = request.get_json()
        if not data:
            return jsonify({'error': '未收到有效的请求数据'}), 400

        curve_type = data.get('type', '').strip()
        if curve_type not in _CURVE_DIR_MAP:
            return jsonify({'error': '无效的曲线类型，必须为 roi 或 retention'}), 400

        curve_id = data.get('id', '').strip()
        if not curve_id:
            return jsonify({'error': '曲线ID不能为空'}), 400

        new_name = data.get('name', '').strip()
        if not new_name:
            return jsonify({'error': '新名称不能为空'}), 400

        file_path = _get_curve_file_path(curve_type, curve_id)
        if not file_path or not os.path.exists(file_path):
            return jsonify({'error': '曲线不存在'}), 404

        with open(file_path, 'r', encoding='utf-8') as f:
            curve_data = json.load(f)

        curve_data['name'] = new_name
        curve_data['updatedAt'] = datetime.now().isoformat()

        with open(file_path, 'w', encoding='utf-8') as f:
            json.dump(curve_data, f, ensure_ascii=False, indent=2)

        return jsonify({'success': True, 'message': f'曲线已重命名为 "{new_name}"', 'curve': curve_data})

    except Exception as e:
        return jsonify({'error': f'重命名曲线时发生错误: {str(e)}'}), 500


if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)