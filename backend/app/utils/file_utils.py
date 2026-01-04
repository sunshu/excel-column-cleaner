"""
文件工具模块
提供文件处理相关的工具函数
"""

import os
import logging
from typing import Optional

logger = logging.getLogger(__name__)

def validate_excel_file(filename: str) -> bool:
    """
    验证文件是否为有效的 Excel 文件
    
    Args:
        filename: 文件名
    
    Returns:
        bool: 是否为有效的 Excel 文件
    """
    if not filename:
        return False
    
    # 检查文件扩展名
    valid_extensions = ['.xlsx', '.xls']
    file_ext = os.path.splitext(filename.lower())[1]
    
    return file_ext in valid_extensions

def generate_filename(original_filename: str, suffix: str = "_processed") -> str:
    """
    生成新的文件名
    
    Args:
        original_filename: 原始文件名
        suffix: 要添加的后缀
    
    Returns:
        str: 新的文件名
    """
    if not original_filename:
        return f"processed_file{suffix}.xlsx"
    
    # 清理文件名，确保安全
    clean_filename = sanitize_filename(original_filename)
    
    # 分离文件名和扩展名
    name, ext = os.path.splitext(clean_filename)
    
    # 如果没有扩展名，默认使用 .xlsx
    if not ext:
        ext = '.xlsx'
    
    # 生成新文件名
    new_filename = f"{name}{suffix}{ext}"
    
    return new_filename

def get_file_size_mb(file_size_bytes: int) -> float:
    """
    将字节转换为 MB
    
    Args:
        file_size_bytes: 文件大小（字节）
    
    Returns:
        float: 文件大小（MB）
    """
    return file_size_bytes / (1024 * 1024)

def validate_file_size(file_size_bytes: int, max_size_mb: float = 50) -> bool:
    """
    验证文件大小是否在允许范围内
    
    Args:
        file_size_bytes: 文件大小（字节）
        max_size_mb: 最大允许大小（MB）
    
    Returns:
        bool: 文件大小是否合法
    """
    file_size_mb = get_file_size_mb(file_size_bytes)
    return file_size_mb <= max_size_mb

def sanitize_filename(filename: str) -> str:
    """
    清理文件名，移除不安全的字符，保持中文字符
    
    Args:
        filename: 原始文件名
    
    Returns:
        str: 清理后的文件名
    """
    if not filename:
        return "unnamed_file"
    
    # 移除路径分隔符和其他不安全字符，但保留中文字符
    unsafe_chars = ['/', '\\', ':', '*', '?', '"', '<', '>', '|']
    clean_name = filename
    
    for char in unsafe_chars:
        clean_name = clean_name.replace(char, '_')
    
    # 移除前后空白字符
    clean_name = clean_name.strip()
    
    # 如果文件名为空，使用默认名称
    if not clean_name:
        clean_name = "unnamed_file"
    
    return clean_name

def ensure_directory_exists(directory_path: str) -> bool:
    """
    确保目录存在，如果不存在则创建
    
    Args:
        directory_path: 目录路径
    
    Returns:
        bool: 操作是否成功
    """
    try:
        os.makedirs(directory_path, exist_ok=True)
        return True
    except Exception as e:
        logger.error(f"创建目录失败 {directory_path}: {str(e)}")
        return False