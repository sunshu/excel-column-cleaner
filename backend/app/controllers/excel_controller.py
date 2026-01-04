"""
Excel 控制器
处理 Excel 相关的 HTTP 请求
"""

from fastapi import APIRouter, UploadFile, File, Form, HTTPException
from fastapi.responses import StreamingResponse
import logging
import io
import os
from typing import List

from services.excel_service import ExcelService
from utils.file_utils import validate_excel_file, generate_filename

logger = logging.getLogger(__name__)
router = APIRouter()

# 创建 Excel 服务实例
excel_service = ExcelService()

@router.post("/excel/delete-columns")
async def delete_excel_columns(
    file: UploadFile = File(..., description="要处理的 Excel 文件"),
    columns: str = Form(..., description="要删除的列索引，用逗号分隔，如：3,5")
):
    """
    删除 Excel 文件中的指定列
    
    Args:
        file: 上传的 Excel 文件
        columns: 要删除的列索引字符串，如 "3,5,7"
    
    Returns:
        StreamingResponse: 处理后的 Excel 文件
    """
    try:
        # 验证文件
        if not file.filename:
            raise HTTPException(status_code=400, detail="未选择文件")
        
        # 验证文件类型
        if not validate_excel_file(file.filename):
            raise HTTPException(
                status_code=400, 
                detail="不支持的文件格式，请上传 .xlsx 或 .xls 文件"
            )
        
        # 解析列索引
        try:
            column_indices = parse_column_indices(columns)
        except ValueError as e:
            raise HTTPException(status_code=400, detail=str(e))
        
        logger.info(f"处理文件: {file.filename}, 删除列: {column_indices}")
        
        # 读取文件内容
        file_content = await file.read()
        
        # 处理 Excel 文件
        processed_content = excel_service.delete_columns(file_content, column_indices)
        
        # 生成新文件名
        new_filename = generate_filename(file.filename, "_processed")
        
        # 返回处理后的文件
        # 处理中文文件名编码问题
        from urllib.parse import quote
        encoded_filename = quote(new_filename.encode('utf-8'))
        
        return StreamingResponse(
            io.BytesIO(processed_content),
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={
                "Content-Disposition": f"attachment; filename*=UTF-8''{encoded_filename}"
            }
        )
        
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"处理 Excel 文件时出错: {str(e)}", exc_info=True)
        raise HTTPException(status_code=500, detail=f"处理文件时出错: {str(e)}")

def parse_column_indices(columns_str: str) -> List[int]:
    """
    解析列索引字符串
    
    Args:
        columns_str: 列索引字符串，如 "3,5,7"
    
    Returns:
        List[int]: 列索引列表
    
    Raises:
        ValueError: 当输入格式不正确时
    """
    if not columns_str.strip():
        raise ValueError("列索引不能为空")
    
    try:
        # 分割并清理空白字符
        column_parts = [part.strip() for part in columns_str.split(',') if part.strip()]
        
        if not column_parts:
            raise ValueError("列索引不能为空")
        
        # 转换为整数并验证
        column_indices = []
        for part in column_parts:
            if not part.isdigit():
                raise ValueError(f"列索引必须是正整数: {part}")
            
            col_index = int(part)
            if col_index < 1:
                raise ValueError(f"列索引必须大于 0: {col_index}")
            
            column_indices.append(col_index)
        
        # 去重并排序（降序，从右到左删除避免索引变化）
        column_indices = sorted(list(set(column_indices)), reverse=True)
        
        return column_indices
        
    except ValueError:
        raise
    except Exception as e:
        raise ValueError(f"解析列索引时出错: {str(e)}")

@router.post("/excel/preview")
async def preview_excel_columns(
    file: UploadFile = File(..., description="要预览的 Excel 文件")
):
    """
    预览 Excel 文件的列信息
    
    Args:
        file: 上传的 Excel 文件
    
    Returns:
        dict: 包含列信息的字典
    """
    try:
        # 验证文件
        if not file.filename:
            raise HTTPException(status_code=400, detail="未选择文件")
        
        # 验证文件类型
        if not validate_excel_file(file.filename):
            raise HTTPException(
                status_code=400, 
                detail="不支持的文件格式，请上传 .xlsx 或 .xls 文件"
            )
        
        logger.info(f"预览文件: {file.filename}")
        
        # 读取文件内容
        file_content = await file.read()
        
        # 获取列信息
        columns_info = excel_service.get_columns_info(file_content)
        
        return {
            "filename": file.filename,
            "columns": columns_info
        }
        
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"预览 Excel 文件时出错: {str(e)}", exc_info=True)
        raise HTTPException(status_code=500, detail=f"预览文件时出错: {str(e)}")

@router.get("/excel/info")
async def get_api_info():
    """获取 API 信息"""
    return {
        "name": "Excel 列删除 API",
        "version": "1.0.0",
        "endpoints": {
            "preview": {
                "method": "POST",
                "path": "/api/excel/preview",
                "description": "预览 Excel 文件的列信息"
            },
            "delete_columns": {
                "method": "POST",
                "path": "/api/excel/delete-columns",
                "description": "删除 Excel 文件中的指定列"
            }
        }
    }