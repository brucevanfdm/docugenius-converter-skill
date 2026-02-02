#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
独立的文档转换脚本
- 将 Word、Excel、PowerPoint 和 PDF 文件转换为 Markdown 格式
- 将 Markdown 文件转换为 Word (.docx) 格式

依赖：
- python-docx: pip install python-docx
- openpyxl: pip install openpyxl
- python-pptx: pip install python-pptx
- pdfplumber: pip install pdfplumber
- Node.js: 用于 Markdown 转 DOCX（需要单独安装）
"""

import sys
import os
import json
import importlib
import re
import subprocess
import shutil
import io
import time
import hashlib
from pathlib import Path

SUPPORTED_EXTENSIONS = ['.docx', '.xlsx', '.pptx', '.pdf', '.md']
MAX_FILE_SIZE_BYTES = 100 * 1024 * 1024
NODE_CONVERT_TIMEOUT_SECONDS = 120
CACHE_EXPIRY_SECONDS = None  # 缓存永不过期，仅在 requirements.txt 变更时失效

def _configure_windows_stdio():
    """
    Windows 控制台经常使用 GBK/CP936，直接输出某些符号（如 ✓/✗）会触发 UnicodeEncodeError。

    策略：
    - 交互式控制台（isatty=True）：不强制切换编码，尽量保持用户终端显示正常；仅将 errors 调整为 replace，避免崩溃。
    - 非交互（被管道/测试框架捕获）：优先输出 UTF-8，保证机器可读；同样使用 errors=replace 兜底。
    """
    if sys.platform != "win32":
        return

    def _safe_reconfigure(stream, *, encoding=None, errors=None):
        if not hasattr(stream, "reconfigure"):
            return False
        try:
            kwargs = {}
            if encoding is not None:
                kwargs["encoding"] = encoding
            if errors is not None:
                kwargs["errors"] = errors
            stream.reconfigure(**kwargs)
            return True
        except Exception:
            return False

    def _safe_wrap(stream, *, encoding, errors):
        buffer = None
        if hasattr(stream, "detach"):
            try:
                buffer = stream.detach()
            except Exception:
                buffer = None
        if buffer is None:
            buffer = getattr(stream, "buffer", None)
        if buffer is None:
            return False
        try:
            wrapped = io.TextIOWrapper(buffer, encoding=encoding, errors=errors, line_buffering=True)
            if stream is sys.stdout:
                sys.stdout = wrapped
            elif stream is sys.stderr:
                sys.stderr = wrapped
            return True
        except Exception:
            return False

    is_tty = bool(getattr(sys.stdout, "isatty", lambda: False)())
    errors = "replace"

    if is_tty:
        if not _safe_reconfigure(sys.stdout, errors=errors):
            current_encoding = getattr(sys.stdout, "encoding", None) or "utf-8"
            _safe_wrap(sys.stdout, encoding=current_encoding, errors=errors)
        if not _safe_reconfigure(sys.stderr, errors=errors):
            current_encoding = getattr(sys.stderr, "encoding", None) or "utf-8"
            _safe_wrap(sys.stderr, encoding=current_encoding, errors=errors)
        return

    target_encoding = "utf-8"
    if not _safe_reconfigure(sys.stdout, encoding=target_encoding, errors=errors):
        _safe_wrap(sys.stdout, encoding=target_encoding, errors=errors)
    if not _safe_reconfigure(sys.stderr, encoding=target_encoding, errors=errors):
        _safe_wrap(sys.stderr, encoding=target_encoding, errors=errors)


_configure_windows_stdio()

_DEPENDENCIES_BY_EXT = {
    '.docx': [('docx', 'python-docx')],
    '.xlsx': [('openpyxl', 'openpyxl')],
    '.pptx': [('pptx', 'python-pptx')],
    '.pdf': [('pdfplumber', 'pdfplumber')],
    '.md': [],  # Markdown 转 DOCX 使用 Node.js，无 Python 依赖
}

# ==================== 缓存管理函数 ====================

def _find_project_root(start_dir: Path):
    """
    从给定目录开始向上查找项目根目录。

    判定依据（任一命中即可）：
    - SKILL.md（本仓库特征文件）
    - requirements.txt
    - .git（开发态仓库）
    """
    try:
        current = start_dir.resolve()
    except Exception:
        current = start_dir

    candidates = [current, *list(getattr(current, "parents", []))]
    for candidate in candidates:
        try:
            if (
                (candidate / 'SKILL.md').exists()
                or (candidate / 'requirements.txt').exists()
                or (candidate / '.git').exists()
            ):
                return candidate
        except Exception:
            continue
    return None

def get_cache_dir():
    """获取项目级缓存目录"""
    # 优先从当前工作目录向上找项目根目录；否则回退到脚本所在目录向上找；
    # 再不行（例如从系统目录运行），回退到用户级缓存目录，避免因权限问题导致缓存不可用。
    project_root = _find_project_root(Path.cwd())
    if project_root is None:
        project_root = _find_project_root(Path(__file__).resolve().parent)

    if project_root is not None:
        cache_dir = project_root / '.claude' / 'cache' / 'docugenius-converter'
    else:
        if sys.platform == "win32":
            base_dir = os.getenv("LOCALAPPDATA") or str(Path.home() / "AppData" / "Local")
        else:
            base_dir = os.getenv("XDG_CACHE_HOME") or str(Path.home() / ".cache")
        cache_dir = Path(base_dir) / 'docugenius-converter'

    # 创建缓存目录
    try:
        cache_dir.mkdir(parents=True, exist_ok=True)
    except Exception:
        # 如果无法创建，返回 None（禁用缓存）
        return None

    return cache_dir

def get_requirements_hash():
    """计算 requirements.txt 的 hash，用于检测依赖变更"""
    try:
        # 尝试找到 requirements.txt（按项目根目录判定，避免工作目录不同导致误判）
        project_root = _find_project_root(Path(__file__).resolve().parent)
        if project_root is None:
            project_root = Path(__file__).resolve().parent.parent

        req_file = project_root / 'requirements.txt'

        if not req_file.exists():
            return None

        with open(req_file, 'rb') as f:
            return hashlib.md5(f.read()).hexdigest()
    except Exception:
        return None

def load_dependency_cache():
    """加载依赖缓存"""
    cache_dir = get_cache_dir()
    if not cache_dir:
        return None

    cache_file = cache_dir / 'dependencies.json'

    if not cache_file.exists():
        return None

    try:
        with open(cache_file, 'r', encoding='utf-8') as f:
            cache = json.load(f)

        # 检查缓存是否过期（如果设置了过期时间）
        if CACHE_EXPIRY_SECONDS is not None:
            cache_time = cache.get('timestamp', 0)
            if time.time() - cache_time > CACHE_EXPIRY_SECONDS:
                return None

        # 检查 requirements.txt 是否变更
        current_hash = get_requirements_hash()
        cached_hash = cache.get('requirements_hash')
        if current_hash and cached_hash != current_hash:
            return None

        return cache
    except Exception:
        return None

def save_dependency_cache(cache_data):
    """保存依赖缓存"""
    cache_dir = get_cache_dir()
    if not cache_dir:
        return

    cache_file = cache_dir / 'dependencies.json'

    try:
        # 添加时间戳和 requirements hash
        cache_data['timestamp'] = time.time()
        cache_data['requirements_hash'] = get_requirements_hash()

        with open(cache_file, 'w', encoding='utf-8') as f:
            json.dump(cache_data, f, ensure_ascii=False, indent=2)
    except Exception:
        pass  # 缓存保存失败不影响主流程

# ==================== 依赖检查函数 ====================

def check_dependencies(file_ext=None):
    """
    检查必需的依赖是否已安装（默认检查全部；传入 file_ext 时仅检查该格式所需）

    使用项目级缓存提高性能：
    - 缓存位置：优先 <project>/.claude/cache/docugenius-converter/dependencies.json
      - 若无法识别项目根目录或无权限写入，则回退到用户级缓存目录（Windows: %LOCALAPPDATA%/docugenius-converter）
    - 缓存有效期：永不过期（仅在 requirements.txt 变更时失效）
    - 缓存失效：requirements.txt 变更时自动失效，或手动清除缓存
    """
    # 确定需要检查的依赖
    deps = _DEPENDENCIES_BY_EXT.get(file_ext) if file_ext else [
        ('docx', 'python-docx'),
        ('openpyxl', 'openpyxl'),
        ('pptx', 'python-pptx'),
        ('pdfplumber', 'pdfplumber'),
    ]
    if deps is None:
        deps = [
            ('docx', 'python-docx'),
            ('openpyxl', 'openpyxl'),
            ('pptx', 'python-pptx'),
            ('pdfplumber', 'pdfplumber'),
        ]

    # 生成缓存键
    cache_key = ','.join([pip_name for _, pip_name in deps])

    # 尝试从缓存加载
    cache = load_dependency_cache()
    if cache and 'dependencies' in cache:
        cached_deps = cache['dependencies']
        # 检查缓存中是否包含当前需要的依赖信息
        if cache_key in cached_deps:
            cached_result = cached_deps[cache_key]
            if cached_result.get('all_installed'):
                return True, None
            else:
                missing = cached_result.get('missing', [])
                return False, f"缺少依赖库: {', '.join(missing)}。请运行: pip install {' '.join(missing)}"

    # 缓存未命中，执行实际检查
    missing = []
    for module_name, pip_name in deps:
        try:
            importlib.import_module(module_name)
        except ImportError:
            missing.append(pip_name)

    # 构建结果
    if missing:
        result = {
            'all_installed': False,
            'missing': missing
        }
    else:
        result = {
            'all_installed': True,
            'missing': []
        }

    # 保存到缓存
    cache = cache or {'dependencies': {}}
    cache['dependencies'][cache_key] = result
    save_dependency_cache(cache)

    if missing:
        return False, f"缺少依赖库: {', '.join(missing)}。请运行: pip install {' '.join(missing)}"

    return True, None

def _normalize_table_cell(value):
    """将单元格内容规范化为安全的 Markdown 表格单元格文本"""
    if value is None:
        return ""
    text = str(value)
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    text = text.replace("\n", " ")
    text = text.strip()
    # Markdown 表格分隔符转义
    text = text.replace("|", "\\|")
    return text

def convert_docx(file_path):
    """转换 Word 文档，支持标题、格式和列表（含编号/层级）"""
    import docx

    doc = docx.Document(file_path)
    content = ""

    def get_numbering_info(para):
        """
        尝试从段落的 numPr / numbering.xml 解析列表信息

        Returns:
            None 或 {'level': int, 'ordered': bool}
        """
        try:
            p = getattr(para, "_p", None)
            ppr = getattr(p, "pPr", None) if p is not None else None
            num_pr = getattr(ppr, "numPr", None) if ppr is not None else None
            if num_pr is None:
                return None

            num_id_el = getattr(num_pr, "numId", None)
            ilvl_el = getattr(num_pr, "ilvl", None)
            level = int(ilvl_el.val) if ilvl_el is not None and getattr(ilvl_el, "val", None) is not None else 0

            num_fmt = None
            num_id = None
            if num_id_el is not None and getattr(num_id_el, "val", None) is not None:
                try:
                    num_id = int(num_id_el.val)
                except Exception:
                    num_id = None

            if num_id is not None:
                try:
                    numbering = doc.part.numbering_part.element
                    ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
                    num_nodes = numbering.xpath(f'.//w:num[@w:numId="{num_id}"]', namespaces=ns)
                    if num_nodes:
                        abstract_id_nodes = num_nodes[0].xpath('./w:abstractNumId', namespaces=ns)
                        if abstract_id_nodes:
                            abstract_id = abstract_id_nodes[0].get(f'{{{ns["w"]}}}val')
                            abstract_nodes = numbering.xpath(
                                f'.//w:abstractNum[@w:abstractNumId="{abstract_id}"]',
                                namespaces=ns,
                            )
                            if abstract_nodes:
                                lvl_nodes = abstract_nodes[0].xpath(f'./w:lvl[@w:ilvl="{level}"]', namespaces=ns)
                                if not lvl_nodes:
                                    lvl_nodes = abstract_nodes[0].xpath('./w:lvl', namespaces=ns)
                                if lvl_nodes:
                                    fmt_nodes = lvl_nodes[0].xpath('./w:numFmt', namespaces=ns)
                                    if fmt_nodes:
                                        num_fmt = fmt_nodes[0].get(f'{{{ns["w"]}}}val')
                except Exception:
                    num_fmt = None

            style = getattr(para, "style", None)
            style_name = getattr(style, "name", "") if style is not None else ""
            style_id = getattr(style, "style_id", "") if style is not None else ""
            style_hint = f"{style_name} {style_id}".lower()

            if (num_fmt or "").lower() == "bullet":
                ordered = False
            elif (num_fmt or ""):
                ordered = True
            elif "bullet" in style_hint or "项目符号" in style_hint or "符号" in style_hint:
                ordered = False
            elif "number" in style_hint or "编号" in style_hint:
                ordered = True
            else:
                ordered = True

            return {'level': max(level, 0), 'ordered': ordered}
        except Exception:
            return None

    def format_run_text(run):
        """格式化单个文本片段，添加Markdown格式标记"""
        text = run.text
        if not text:
            return ""

        # 应用粗体和斜体格式
        if run.bold and run.italic:
            text = f"***{text}***"
        elif run.bold:
            text = f"**{text}**"
        elif run.italic:
            text = f"*{text}*"

        return text

    def process_paragraph(para):
        """处理单个段落，识别标题、列表和格式"""
        if not para.text.strip():
            return ""

        style = para.style if hasattr(para, "style") else None
        style_name = style.name if style else ""
        style_id = getattr(style, "style_id", "") if style else ""

        # 先拼接富文本（列表项也需要保留粗体/斜体）
        formatted_text = ""
        for run in para.runs:
            formatted_text += format_run_text(run)
        text_value = (formatted_text.strip() or para.text.strip()).replace("\n", " ").strip()
        if not text_value:
            return ""

        # 识别标题层级
        heading_level = None
        if isinstance(style_id, str) and style_id:
            m = re.match(r"(?i)^heading(\d+)$", style_id.strip())
            if m:
                heading_level = int(m.group(1))
        if heading_level is None and isinstance(style_name, str) and style_name:
            # 兼容中文 Word 默认标题样式（例如“标题 1”）
            m = re.match(r"^(?:Heading|标题)\s*(\d+)$", style_name.strip())
            if m:
                heading_level = int(m.group(1))

        if heading_level is not None:
            heading_prefix = "#" * min(heading_level, 6)  # Markdown最多支持6级标题
            return f"{heading_prefix} {para.text.strip()}\n\n"

        # 检查是否是列表项
        # 优先使用 numPr + numbering.xml 解析列表编号格式与层级
        numbering_info = get_numbering_info(para)
        if numbering_info:
            indent = "    " * numbering_info["level"]
            marker = "1." if numbering_info["ordered"] else "-"
            return f"{indent}{marker} {text_value}\n"

        # 注意：python-docx对列表的支持有限，这里做基本处理（按样式兜底）
        style_name_str = style_name or ""
        style_id_str = style_id or ""
        if (isinstance(style_id_str, str) and style_id_str.startswith("List")) or (
            isinstance(style_name_str, str) and style_name_str.startswith("List")
        ):
            level = 0
            m = re.search(r"(\d+)$", style_name_str.strip())
            if m:
                level = max(int(m.group(1)) - 1, 0)
            indent = "    " * level
            if "Bullet" in style_id_str or "Bullet" in style_name_str or style_name_str == "List Bullet":
                return f"{indent}- {text_value}\n"
            if "Number" in style_id_str or "Number" in style_name_str or style_name_str == "List Number":
                return f"{indent}1. {text_value}\n"

        return text_value + "\n\n"

    # 处理文档中的所有元素（段落和表格）
    # 需要按照它们在文档中的顺序处理
    paragraphs_iter = iter(doc.paragraphs)
    tables_iter = iter(doc.tables)
    for element in doc.element.body:
        # 处理段落
        if element.tag.endswith('p'):
            para = next(paragraphs_iter, None)
            if para is not None:
                content += process_paragraph(para)

        # 处理表格
        elif element.tag.endswith('tbl'):
            table = next(tables_iter, None)
            if table is not None:
                for i, row in enumerate(table.rows):
                    row_data = [_normalize_table_cell(cell.text) for cell in row.cells]

                    if i == 0:  # 表头
                        content += "| " + " | ".join(row_data) + " |\n"
                        content += "| " + " | ".join(["---"] * len(row_data)) + " |\n"
                    else:
                        content += "| " + " | ".join(row_data) + " |\n"
                content += "\n"

    return content.strip()

def convert_xlsx(file_path):
    """转换 Excel 文件"""
    import openpyxl

    workbook = openpyxl.load_workbook(file_path, data_only=True)
    content = ""

    for sheet_name in workbook.sheetnames:
        if len(workbook.sheetnames) > 1:
            content += f"## {sheet_name}\n\n"

        worksheet = workbook[sheet_name]
        rows = list(worksheet.iter_rows(values_only=True))

        if rows:
            # 过滤空行
            non_empty_rows = []
            for row in rows:
                if any(cell is not None and str(cell).strip() for cell in row):
                    non_empty_rows.append(row)

            if non_empty_rows:
                def last_non_empty_col_count(row):
                    for idx in range(len(row) - 1, -1, -1):
                        cell = row[idx]
                        if cell is None:
                            continue
                        if str(cell).strip():
                            return idx + 1
                    return 0

                max_cols = max(last_non_empty_col_count(row) for row in non_empty_rows)

                if max_cols == 0:
                    # 如果没有非空单元格，跳过这个工作表
                    continue

                for i, row in enumerate(non_empty_rows):
                    row_data = []
                    for j in range(max_cols):
                        if j < len(row) and row[j] is not None:
                            row_data.append(_normalize_table_cell(row[j]))
                        else:
                            row_data.append("")

                    if i == 0:  # 表头
                        content += "| " + " | ".join(row_data) + " |\n"
                        content += "| " + " | ".join(["---"] * len(row_data)) + " |\n"
                    else:
                        content += "| " + " | ".join(row_data) + " |\n"
                content += "\n"

    return content.strip()

def convert_pptx(file_path):
    """转换 PowerPoint 文件"""
    import pptx

    presentation = pptx.Presentation(file_path)
    content = ""

    for i, slide in enumerate(presentation.slides, 1):
        if len(presentation.slides) > 1:
            content += f"## Slide {i}\n\n"

        for shape in slide.shapes:
            if hasattr(shape, "text") and shape.text.strip():
                content += shape.text.strip() + "\n\n"

        if i < len(presentation.slides):
            content += "---\n\n"

    return content.strip()

def convert_pdf(file_path):
    """转换 PDF 文件，支持文本和表格提取"""
    import pdfplumber

    content = ""
    with pdfplumber.open(file_path) as pdf:
        for i, page in enumerate(pdf.pages):
            if len(pdf.pages) > 1:
                content += f"## Page {i+1}\n\n"

            # 提取表格
            tables = page.extract_tables()
            if tables:
                for table in tables:
                    if table and len(table) > 0:
                        # 过滤空行和空单元格
                        filtered_table = []
                        for row in table:
                            if row and any(cell is not None and str(cell).strip() for cell in row):
                                # 清理单元格内容
                                cleaned_row = [_normalize_table_cell(cell) for cell in row]
                                filtered_table.append(cleaned_row)

                        if filtered_table:
                            max_cols = max(len(r) for r in filtered_table)
                            normalized = [r + [""] * (max_cols - len(r)) for r in filtered_table]
                            # 输出为Markdown表格
                            for idx, row in enumerate(normalized):
                                content += "| " + " | ".join(row) + " |\n"
                                if idx == 0:  # 在第一行后添加分隔符
                                    content += "| " + " | ".join(["---"] * len(row)) + " |\n"
                            content += "\n"

            # 提取文本（排除表格区域）
            # 注意：pdfplumber的extract_text会包含表格文本，这里我们简单处理
            text = page.extract_text()
            if text and text.strip():
                # 如果页面有表格，文本可能包含表格内容，这里做基本清理
                lines = text.split('\n')
                cleaned_lines = []
                for line in lines:
                    line = line.strip()
                    if line:
                        cleaned_lines.append(line)

                if cleaned_lines:
                    # 如果这个页面已经有表格了，添加一个分隔
                    if tables:
                        content += "### 文本内容\n\n"

                    content += '\n'.join(cleaned_lines) + "\n\n"

    return content.strip()

def convert_md(file_path, output_dir=None):
    """
    将 Markdown 文件转换为 DOCX 格式（通过 Node.js 脚本）

    Args:
        file_path: Markdown 文件路径
        output_dir: 可选的输出目录

    Returns:
        包含 'success'、'output_path' 和可选 'error' 的字典
    """
    # 检查 Node.js 是否可用
    node_cmd = shutil.which('node')
    if not node_cmd:
        return {
            'success': False,
            'error': '未找到 Node.js。Markdown 转 DOCX 需要 Node.js 环境。请安装 Node.js: https://nodejs.org/'
        }

    # 获取 Node.js 脚本路径
    script_dir = os.path.dirname(os.path.abspath(__file__))
    node_script = os.path.join(script_dir, 'md_to_docx', 'index.js')

    if not os.path.exists(node_script):
        return {
            'success': False,
            'error': f'Node.js 转换脚本不存在: {node_script}。请运行 npm install 安装依赖。'
        }

    # 检查 node_modules 是否存在
    node_modules = os.path.join(script_dir, 'md_to_docx', 'node_modules')
    if not os.path.exists(node_modules):
        return {
            'success': False,
            'error': f'Node.js 依赖未安装。请在 {os.path.join(script_dir, "md_to_docx")} 目录下运行: npm install'
        }

    try:
        # 调用 Node.js 脚本
        cmd = [node_cmd, node_script, file_path]
        if output_dir:
            cmd.append(output_dir)

        result = subprocess.run(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            universal_newlines=True,
            encoding='utf-8',
            timeout=NODE_CONVERT_TIMEOUT_SECONDS,
        )

        # 解析输出
        try:
            output = json.loads(result.stdout)
            return output
        except json.JSONDecodeError:
            if result.returncode == 0:
                return {
                    'success': True,
                    'output_path': result.stdout.strip(),
                    'message': '转换成功'
                }
            else:
                return {
                    'success': False,
                    'error': f'Node.js 脚本输出解析失败: {result.stdout}\n{result.stderr}'
                }

    except subprocess.TimeoutExpired:
        return {
            'success': False,
            'error': '转换超时（超过2分钟）'
        }
    except Exception as e:
        return {
            'success': False,
            'error': f'调用 Node.js 脚本失败: {str(e)}'
        }

def convert_document(file_path, extract_images=True, output_dir=None):
    """
    将文档转换为 Markdown 格式

    Args:
        file_path: 文档文件路径
        extract_images: 是否提取图片（当前版本暂不支持）
        output_dir: 可选的输出目录（默认为同目录下的 Markdown/ 子目录）

    Returns:
        包含 'success'、'markdown_content'、'output_path' 和可选 'error' 的字典
    """
    # 验证输入文件
    file_path = os.path.normpath(file_path)
    if not os.path.exists(file_path):
        return {
            'success': False,
            'error': f'文件不存在: {file_path}'
        }

    # 检查文件大小（限制为100MB）
    try:
        file_size = os.path.getsize(file_path)
        if file_size > MAX_FILE_SIZE_BYTES:
            return {
                'success': False,
                'error': f'文件过大: {file_size / (1024*1024):.2f}MB，超过限制 {MAX_FILE_SIZE_BYTES / (1024*1024):.0f}MB'
            }
    except OSError as e:
        return {
            'success': False,
            'error': f'无法读取文件大小: {str(e)}'
        }

    # 检查文件扩展名
    file_ext = os.path.splitext(file_path)[1].lower()
    if file_ext not in SUPPORTED_EXTENSIONS:
        return {
            'success': False,
            'error': f'不支持的文件格式: {file_ext}。支持的格式: {", ".join(SUPPORTED_EXTENSIONS)}'
        }

    # Markdown 转 DOCX 使用单独的处理流程
    if file_ext == '.md':
        return convert_md(file_path, output_dir)

    # 检查依赖（按格式按需检查，避免无关依赖阻塞）
    deps_ok, error_msg = check_dependencies(file_ext)
    if not deps_ok:
        return {
            'success': False,
            'error': error_msg
        }

    try:
        # 根据文件类型转换
        if file_ext == '.docx':
            markdown_content = convert_docx(file_path)
        elif file_ext == '.xlsx':
            markdown_content = convert_xlsx(file_path)
        elif file_ext == '.pptx':
            markdown_content = convert_pptx(file_path)
        elif file_ext == '.pdf':
            markdown_content = convert_pdf(file_path)
        else:
            return {
                'success': False,
                'error': f'不支持的文件类型: {file_ext}'
            }

        # 确定输出路径
        if output_dir:
            os.makedirs(output_dir, exist_ok=True)
            output_path = os.path.join(output_dir, os.path.splitext(os.path.basename(file_path))[0] + '.md')
        else:
            # 默认：在同目录下创建 Markdown/ 子目录
            file_dir = os.path.dirname(file_path) or '.'
            output_dir = os.path.join(file_dir, 'Markdown')
            os.makedirs(output_dir, exist_ok=True)
            output_path = os.path.join(output_dir, os.path.splitext(os.path.basename(file_path))[0] + '.md')

        # 保存 Markdown 文件
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(markdown_content)

        return {
            'success': True,
            'markdown_content': markdown_content,
            'output_path': output_path
        }

    except PermissionError as e:
        return {
            'success': False,
            'error': f'权限不足: 无法读取文件或写入输出目录 - {str(e)}'
        }
    except MemoryError:
        return {
            'success': False,
            'error': '内存不足: 文件可能过大，请尝试处理较小的文件'
        }
    except FileNotFoundError as e:
        return {
            'success': False,
            'error': f'文件未找到: {str(e)}'
        }
    except OSError as e:
        return {
            'success': False,
            'error': f'系统错误: {str(e)}'
        }
    except Exception as e:
        return {
            'success': False,
            'error': f'转换错误 ({type(e).__name__}): {str(e)}'
        }

def batch_convert(directory, recursive=True, extract_images=True, output_dir=None):
    """
    批量转换目录中的所有支持的文档

    Args:
        directory: 要扫描的目录
        recursive: 是否递归扫描子目录
        extract_images: 是否提取图片
        output_dir: 可选的输出目录

    Returns:
        转换结果列表
    """
    results = []

    if recursive:
        # 递归扫描
            for root, dirs, files in os.walk(directory):
                for file in files:
                    if os.path.splitext(file)[1].lower() in SUPPORTED_EXTENSIONS:
                        file_path = os.path.join(root, file)
                        result = convert_document(file_path, extract_images, output_dir)
                        results.append({
                            'file': file_path,
                        'result': result
                    })
    else:
        # 只扫描当前目录
        for file in os.listdir(directory):
            file_path = os.path.join(directory, file)
            if os.path.isfile(file_path) and os.path.splitext(file)[1].lower() in SUPPORTED_EXTENSIONS:
                result = convert_document(file_path, extract_images, output_dir)
                results.append({
                    'file': file_path,
                    'result': result
                })

    return results

def clear_cache():
    """清除依赖缓存"""
    cache_dir = get_cache_dir()
    if not cache_dir:
        return {
            'success': False,
            'message': '无法获取缓存目录'
        }

    cache_file = cache_dir / 'dependencies.json'

    if not cache_file.exists():
        return {
            'success': True,
            'message': '缓存不存在，无需清除'
        }

    try:
        cache_file.unlink()
        return {
            'success': True,
            'message': f'已清除缓存: {cache_file}'
        }
    except Exception as e:
        return {
            'success': False,
            'message': f'清除缓存失败: {str(e)}'
        }

def main():
    if len(sys.argv) < 2:
        print('用法: python convert_document.py <file_path> [extract_images] [output_dir]')
        print('  file_path: 文档文件路径')
        print('  extract_images: true/false (默认: true，当前版本暂不支持)')
        print('  output_dir: 可选的输出目录')
        print('')
        print('支持的格式:')
        print('  - Office/PDF 转 Markdown: .docx, .xlsx, .pptx, .pdf')
        print('  - Markdown 转 Word: .md')
        print('')
        print('批量转换: python convert_document.py --batch <directory> [recursive]')
        print('  directory: 要扫描的目录')
        print('  recursive: true/false (默认: true)')
        print('')
        print('清除缓存: python convert_document.py --clear-cache')
        print('  清除依赖检测缓存，强制重新检查依赖')
        sys.exit(1)

    # 清除缓存模式
    if sys.argv[1] == '--clear-cache':
        result = clear_cache()
        print(json.dumps(result, ensure_ascii=False, indent=2))
        sys.exit(0 if result['success'] else 1)

    # 批量转换模式
    if sys.argv[1] == '--batch':
        if len(sys.argv) < 3:
            print('错误: 批量转换需要指定目录')
            sys.exit(1)

        directory = sys.argv[2]
        recursive = sys.argv[3].lower() == 'true' if len(sys.argv) > 3 else True

        results = batch_convert(directory, recursive)

        # 输出结果统计
        success_count = sum(1 for r in results if r['result']['success'])
        total_count = len(results)

        print(json.dumps({
            'total': total_count,
            'success': success_count,
            'failed': total_count - success_count,
            'results': results
        }, ensure_ascii=False, indent=2))

        sys.exit(0 if success_count == total_count else 1)

    # 单文件转换模式
    file_path = sys.argv[1]
    extract_images = sys.argv[2].lower() == 'true' if len(sys.argv) > 2 else True
    output_dir = sys.argv[3] if len(sys.argv) > 3 else None

    result = convert_document(file_path, extract_images, output_dir)

    # 输出结果为 JSON
    print(json.dumps(result, ensure_ascii=False, indent=2))

    sys.exit(0 if result['success'] else 1)

if __name__ == '__main__':
    main()
