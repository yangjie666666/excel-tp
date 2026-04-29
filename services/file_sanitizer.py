"""文件名规范化模块 - 处理非法字符、空值、重复名、长度截断"""

import re

# Windows 文件名非法字符
ILLEGAL_CHARS = r'[\\/:*?"<>|]'
# 最大文件名长度（不含扩展名）
MAX_NAME_LENGTH = 200


def sanitize_filename(name: str, fallback_index: int = 0) -> str:
    """规范化文件名

    Args:
        name: 原始文件名（不含扩展名）
        fallback_index: 当name为空时使用的序号

    Returns:
        规范化后的文件名
    """
    if not name or not str(name).strip():
        return f"未命名_{fallback_index}"

    name = str(name)

    # 去除首尾空格和换行
    name = name.strip().replace('\n', '').replace('\r', '')

    # 替换非法字符为下划线
    name = re.sub(ILLEGAL_CHARS, '_', name)

    # 合并连续的下划线
    name = re.sub(r'_+', '_', name)

    # 去除首尾的下划线和连字符
    name = name.strip('_-')

    # 截断过长的名称
    if len(name) > MAX_NAME_LENGTH:
        name = name[:MAX_NAME_LENGTH]

    # 如果处理后为空，使用fallback
    if not name:
        return f"未命名_{fallback_index}"

    return name


def resolve_duplicates(filenames: list[str]) -> list[str]:
    """解决文件名重复问题

    Args:
        filenames: 文件名列表（含扩展名）

    Returns:
        去重后的文件名列表，重复项追加 _1, _2 等
    """
    seen: dict[str, int] = {}
    result = []

    for filename in filenames:
        # 分离名称和扩展名
        name, ext = _split_ext(filename)

        if filename in seen:
            seen[filename] += 1
            new_name = f"{name}_{seen[filename]}{ext}"
            # 确保新名称也不重复
            while new_name in seen:
                seen[filename] += 1
                new_name = f"{name}_{seen[filename]}{ext}"
            seen[new_name] = 0
            result.append(new_name)
        else:
            seen[filename] = 0
            result.append(filename)

    return result


def build_filename(
    name_col_value: str,
    ext: str,
    manual_prefix: str = "",
    field_prefix_value: str = "",
    fallback_index: int = 0,
    row_num: int | None = None,
    row_pad_width: int = 3,
) -> str:
    """构建最终文件名

    规则: {手动前缀}{字段前缀}{行号前缀}{商品名称}.{扩展名}

    Args:
        name_col_value: 商品名称列的值
        ext: 图片原始扩展名（含点号，如 .png）
        manual_prefix: 用户手动输入的前缀
        field_prefix_value: 选择的字段前缀值
        fallback_index: 空值时的序号
        row_num: Excel行号，作为固定前缀插入（如 001）
        row_pad_width: 行号零填充宽度，根据最大行号动态计算

    Returns:
        完整文件名
    """
    # 规范化各部分
    prefix = sanitize_filename(manual_prefix) if manual_prefix else ""
    field_pfx = sanitize_filename(field_prefix_value) if field_prefix_value else ""
    name = sanitize_filename(name_col_value, fallback_index)

    # 行号前缀（动态零填充，保证字符串排序正确）
    row_prefix = f"{row_num:0{row_pad_width}d}" if row_num is not None else ""

    # 拼接，各部分间用连字符连接
    parts = [p for p in [prefix, field_pfx, row_prefix, name] if p]
    filename = "-".join(parts)

    # 拼接扩展名
    if not ext.startswith('.'):
        ext = f'.{ext}'

    return f"{filename}{ext}"


def _split_ext(filename: str) -> tuple[str, str]:
    """分离文件名和扩展名"""
    if '.' in filename:
        idx = filename.rfind('.')
        return filename[:idx], filename[idx:]
    return filename, ''
