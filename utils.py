"""工具函数模块"""
import subprocess
from typing import Union


def decode_subprocess_error(error: subprocess.CalledProcessError) -> str:
    """
    解码子进程错误信息，尝试多种编码

    Args:
        error: subprocess.CalledProcessError 异常对象

    Returns:
        解码后的错误信息字符串
    """
    error_msg = str(error)

    if error.stderr:
        # 如果 stderr 是字符串，直接返回
        if isinstance(error.stderr, str):
            return error.stderr

        # 如果是字节，尝试多种编码
        if isinstance(error.stderr, bytes):
            # 尝试 UTF-8
            try:
                return error.stderr.decode('utf-8')
            except UnicodeDecodeError:
                pass

            # 尝试 GBK (Windows 中文)
            try:
                return error.stderr.decode('gbk')
            except UnicodeDecodeError:
                pass

            # 最后使用 UTF-8 忽略错误
            return error.stderr.decode('utf-8', errors='ignore')

    return error_msg
