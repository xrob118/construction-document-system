"""文档模板管理 API 路由"""

import os
from datetime import datetime
from fastapi import APIRouter, UploadFile, File, HTTPException
from fastapi.responses import FileResponse

router = APIRouter()

# 模板目录（指向 backend/templates）
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TEMPLATES_DIR = os.path.join(BASE_DIR, 'templates')
os.makedirs(TEMPLATES_DIR, exist_ok=True)

# 文档类型与模板文件名映射
DOC_TYPES = {
    'construction_design': {
        'label': '施工组织设计',
        'ext': ['.doc', '.docx'],
        'template_names': ['施工组织设计模板'],
    },
    'survey': {
        'label': '项目勘察单',
        'ext': ['.doc', '.docx'],
        'template_names': ['项目勘察单模板'],
    },
    'tech_briefing': {
        'label': '技术交底',
        'ext': ['.doc', '.docx'],
        'template_names': ['技术交底模板'],
    },
    'safety_briefing': {
        'label': '安全交底',
        'ext': ['.doc', '.docx'],
        'template_names': ['安全交底模板'],
    },
    'gantt_chart': {
        'label': '施工进度横道图',
        'ext': ['.xlsx', '.xls'],
        'template_names': ['施工进度横道图模板'],
    }
}


def _find_template_file(template_names):
    """根据模板基础名称查找实际文件

    Args:
        template_names: 模板基础名称列表，如 ["施工组织设计模板"]

    Returns:
        找到的文件名（含扩展名），未找到返回 None
    """
    for base_name in template_names:
        for ext in ['.docx', '.doc', '.xlsx', '.xls']:
            filename = f"{base_name}{ext}"
            if os.path.exists(os.path.join(TEMPLATES_DIR, filename)):
                return filename
    return None


@router.get('/api/templates', summary='获取模板列表')
def get_templates():
    """获取所有模板的状态"""
    items = []
    for doc_type, info in DOC_TYPES.items():
        file_name = _find_template_file(info['template_names'])
        uploaded_at = None

        if file_name:
            file_path = os.path.join(TEMPLATES_DIR, file_name)
            mtime = os.path.getmtime(file_path)
            uploaded_at = datetime.fromtimestamp(mtime).strftime('%Y-%m-%d %H:%M:%S')

        items.append({
            'doc_type': doc_type,
            'label': info['label'],
            'file_name': file_name,
            'uploaded_at': uploaded_at
        })
    return {'items': items}


@router.post('/api/templates/upload', summary='上传文档模板')
async def upload_template(doc_type: str, file: UploadFile = File(...)):
    """上传文档模板，支持 .doc 和 .docx 格式"""
    if doc_type not in DOC_TYPES:
        raise HTTPException(status_code=400, detail='不支持的文档类型')

    # 检查文件扩展名
    info = DOC_TYPES[doc_type]
    _, ext = os.path.splitext(file.filename.lower())
    if ext not in info['ext']:
        raise HTTPException(status_code=400, detail=f'仅支持 {", ".join(info["ext"])} 格式')

    # 删除旧模板（匹配所有可能的扩展名）
    for base_name in info['template_names']:
        for old_ext in ['.docx', '.doc', '.xlsx', '.xls']:
            old_file = os.path.join(TEMPLATES_DIR, f"{base_name}{old_ext}")
            if os.path.exists(old_file):
                try:
                    os.remove(old_file)
                except Exception:
                    pass

    # 保存新模板，使用中文模板名 + 实际扩展名
    base_name = info['template_names'][0]
    new_filename = f'{base_name}{ext}'
    file_path = os.path.join(TEMPLATES_DIR, new_filename)

    content = await file.read()
    with open(file_path, 'wb') as f:
        f.write(content)

    return {'message': '上传成功', 'file_name': new_filename}


@router.get('/api/templates/download/{doc_type}', summary='下载文档模板')
def download_template(doc_type: str):
    """下载指定类型的模板文件"""
    if doc_type not in DOC_TYPES:
        raise HTTPException(status_code=404, detail='模板不存在')

    info = DOC_TYPES[doc_type]
    file_name = _find_template_file(info['template_names'])

    if not file_name:
        raise HTTPException(status_code=404, detail='该类型暂无模板文件')

    file_path = os.path.join(TEMPLATES_DIR, file_name)
    return FileResponse(
        path=file_path,
        filename=file_name,
        media_type='application/octet-stream'
    )


@router.delete('/api/templates/{doc_type}', summary='删除模板')
def delete_template(doc_type: str):
    """删除模板"""
    if doc_type not in DOC_TYPES:
        raise HTTPException(status_code=404, detail='模板不存在')

    info = DOC_TYPES[doc_type]
    deleted = False
    for base_name in info['template_names']:
        for ext in ['.docx', '.doc', '.xlsx', '.xls']:
            file_path = os.path.join(TEMPLATES_DIR, f"{base_name}{ext}")
            if os.path.exists(file_path):
                try:
                    os.remove(file_path)
                    deleted = True
                except Exception as e:
                    print(f'删除失败: {e}')

    return {'message': '删除成功', 'deleted': deleted}
