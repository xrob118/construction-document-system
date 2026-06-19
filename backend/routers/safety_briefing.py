"""安全交底 API 路由"""

import os
from fastapi import APIRouter, Depends, HTTPException
from sqlalchemy.orm import Session
from pydantic import BaseModel
from typing import List, Optional, Dict

from models import GeneratedDoc
from database import get_db, SessionLocal
from services.doc_generator import DocGenerator
from services.data_service import DataService

router = APIRouter()


class SafetyBriefingRequest(BaseModel):
    """安全交底生成请求体"""
    project_id: int
    worker_ids: List[int] = []
    content_items: List[str] = []
    worker_roles: Dict[int, str] = {}


@router.post("/api/safety-briefing/generate", summary="生成安全交底文档")
def generate_safety_briefing(request: SafetyBriefingRequest):
    """生成安全交底文档

    接收工程ID、人员ID列表和交底内容项列表，组合数据后生成文档。
    """
    try:
        # 组合数据
        data_service = DataService()
        data_dict = data_service.get_safety_briefing_data(
            project_id=request.project_id,
            worker_ids=request.worker_ids,
            content_items=request.content_items,
            worker_roles=request.worker_roles,
        )

        # 生成文档
        generator = DocGenerator()
        file_path, pdf_path = generator.generate_safety_briefing(data_dict)

        # 获取文件名用于下载
        filename = os.path.basename(file_path)

        # 保存生成记录到数据库
        db = SessionLocal()
        try:
            doc_record = GeneratedDoc(
                project_id=request.project_id,
                doc_type="safety_briefing",
                file_path=file_path,
                pdf_path=pdf_path,
            )
            db.add(doc_record)
            db.commit()
            db.refresh(doc_record)
            doc_id = doc_record.id
        finally:
            db.close()

        return {
            "message": "安全交底生成成功",
            "id": doc_id,
            "filename": filename,
            "file_path": file_path,
            "download_url": f"/output/{filename}",
        }
    except ValueError as e:
        raise HTTPException(status_code=404, detail=str(e))
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"文档生成失败: {str(e)}")


class BatchDeleteRequest(BaseModel):
    """批量删除请求体"""
    ids: List[int]


@router.delete("/api/safety-briefing/batch", summary="批量删除安全交底记录")
def batch_delete_safety_briefing(request: BatchDeleteRequest, db: Session = Depends(get_db)):
    """批量删除安全交底历史记录"""
    if not request.ids:
        raise HTTPException(status_code=400, detail="ids 不能为空")
    deleted_ids = []
    failed = []
    for doc_id in request.ids:
        try:
            record = db.query(GeneratedDoc).filter(GeneratedDoc.id == doc_id).first()
            if not record:
                failed.append({"id": doc_id, "reason": "记录不存在"})
                continue
            for path_attr in ("file_path", "pdf_path"):
                path_str = getattr(record, path_attr, None)
                if path_str and os.path.exists(path_str):
                    try:
                        os.remove(path_str)
                    except Exception:
                        pass
            db.delete(record)
            deleted_ids.append(doc_id)
        except Exception as e:
            failed.append({"id": doc_id, "reason": str(e)})
    try:
        db.commit()
    except Exception as e:
        db.rollback()
        raise HTTPException(status_code=500, detail=f"提交失败: {str(e)}")

    return {
        "message": "批量删除完成",
        "deleted_ids": deleted_ids,
        "failed": failed,
    }


@router.delete("/api/safety-briefing/{doc_id}", summary="删除安全交底记录")
def delete_safety_briefing(doc_id: int, db: Session = Depends(get_db)):
    """删除安全交底历史记录"""
    record = db.query(GeneratedDoc).filter(GeneratedDoc.id == doc_id).first()
    if not record:
        raise HTTPException(status_code=404, detail="记录不存在")

    for path_attr in ("file_path", "pdf_path"):
        path_str = getattr(record, path_attr, None)
        if path_str and os.path.exists(path_str):
            try:
                os.remove(path_str)
            except Exception:
                pass

    try:
        db.delete(record)
        db.commit()
    except Exception as e:
        db.rollback()
        raise HTTPException(status_code=500, detail=f"删除记录失败: {str(e)}")

    return {"message": "删除成功", "id": doc_id}


@router.get("/api/safety-briefing/history", summary="获取安全交底生成历史")
def get_safety_briefing_history(
    project_id: Optional[int] = None,
    db: Session = Depends(get_db),
):
    """获取安全交底的生成历史记录"""
    query = db.query(GeneratedDoc).filter(GeneratedDoc.doc_type == "safety_briefing")
    if project_id:
        query = query.filter(GeneratedDoc.project_id == project_id)
    records = query.order_by(GeneratedDoc.created_at.desc()).all()

    return {
        "items": [
            {
                "id": r.id,
                "project_id": r.project_id,
                "project_name": r.project.project_name if r.project else None,
                "project_code": r.project.project_code if r.project else None,
                "doc_type": r.doc_type,
                "file_path": r.file_path,
                "filename": os.path.basename(r.file_path) if r.file_path else None,
                "download_url": f"/output/{os.path.basename(r.file_path)}" if r.file_path else None,
                "created_at": str(r.created_at) if r.created_at else None,
            }
            for r in records
        ]
    }
