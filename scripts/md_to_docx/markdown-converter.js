/**
 * Markdown 到 HTML 转换器
 * 简化版 Markdown 解析器，支持常见语法
 */

/**
 * 将 Markdown 文本转换为 HTML
 * @param {string} markdown - Markdown 文本
 * @returns {string} HTML 字符串
 */
function markdownToHTML(markdown) {
  if (!markdown || typeof markdown !== 'string') {
    return '';
  }

  let html = markdown;

  // 1. 首先保护代码块
  const codeBlocks = [];
  html = html.replace(/```(\w*)\s*\n([\s\S]*?)```/g, (match, lang, code) => {
    const placeholder = `\x00CODEBLOCK${codeBlocks.length}\x00`;
    codeBlocks.push({ lang: lang || '', code: code.trim() });
    return placeholder;
  });

  // 2. 保护行内代码
  const inlineCodes = [];
  html = html.replace(/`([^`]+)`/g, (match, code) => {
    const placeholder = `\x00INLINECODE${inlineCodes.length}\x00`;
    inlineCodes.push(code);
    return placeholder;
  });

  // 处理标题
  html = html.replace(/^######\s+(.+)$/gm, '<h6>$1</h6>');
  html = html.replace(/^#####\s+(.+)$/gm, '<h5>$1</h5>');
  html = html.replace(/^####\s+(.+)$/gm, '<h4>$1</h4>');
  html = html.replace(/^###\s+(.+)$/gm, '<h3>$1</h3>');
  html = html.replace(/^##\s+(.+)$/gm, '<h2>$1</h2>');
  html = html.replace(/^#\s+(.+)$/gm, '<h1>$1</h1>');

  // 处理粗体和斜体
  html = html.replace(/\*\*\*(.+?)\*\*\*/g, '<strong><em>$1</em></strong>');
  html = html.replace(/\*\*(.+?)\*\*/g, '<strong>$1</strong>');
  html = html.replace(/\*(.+?)\*/g, '<em>$1</em>');
  html = html.replace(/___(.+?)___/g, '<strong><em>$1</em></strong>');
  html = html.replace(/__(.+?)__/g, '<strong>$1</strong>');
  html = html.replace(/_(.+?)_/g, '<em>$1</em>');

  // 处理删除线
  html = html.replace(/~~(.+?)~~/g, '<del>$1</del>');

  // 处理引用块
  html = html.replace(/^>\s+(.+)$/gm, '<blockquote>$1</blockquote>');
  html = html.replace(/<\/blockquote>\n<blockquote>/g, '\n');

  // 处理列表
  html = processLists(html);

  // 处理表格
  html = processTables(html);

  // 处理水平线
  html = html.replace(/^(-{3,}|_{3,}|\*{3,})$/gm, '<hr>');

  // 处理链接 [text](url)
  html = html.replace(/\[([^\]]+)\]\(([^)]+)\)/g, '<a href="$2">$1</a>');

  // 处理图片 ![alt](url)
  html = html.replace(/!\[([^\]]*)\]\(([^)]+)\)/g, '<img src="$2" alt="$1">');

  // 处理段落
  html = processParagraphs(html);

  // 恢复行内代码
  inlineCodes.forEach((code, i) => {
    html = html.replace(`\x00INLINECODE${i}\x00`, () => `<code>${escapeHTML(code)}</code>`);
  });

  // 恢复代码块
  codeBlocks.forEach((block, i) => {
    html = html.replace(`\x00CODEBLOCK${i}\x00`, () => `<pre><code class="language-${block.lang}">${escapeHTML(block.code)}</code></pre>`);
  });

  return html;
}

/**
 * 转义 HTML 特殊字符
 */
function escapeHTML(text) {
  const map = {
    '&': '&amp;',
    '<': '&lt;',
    '>': '&gt;',
    '"': '&quot;',
    "'": '&#039;'
  };
  return text.replace(/[&<>"']/g, char => map[char]);
}

/**
 * 处理列表
 */
function processLists(html) {
  const lines = html.split('\n');
  const result = [];
  const stack = [];

  const countLeadingSpaces = (s) => {
    let count = 0;
    for (let i = 0; i < s.length; i++) {
      if (s[i] === ' ') count++;
      else if (s[i] === '\t') count += 4;
      else break;
    }
    return count;
  };

  const closeToIndent = (targetIndent) => {
    while (stack.length > 0 && targetIndent < stack[stack.length - 1].indent) {
      const top = stack[stack.length - 1];
      if (top.hasOpenLi) {
        result.push('</li>');
        top.hasOpenLi = false;
      }
      result.push(`</${top.type}>`);
      stack.pop();
    }
  };

  const closeAll = () => closeToIndent(-1);

  const ensureList = (type, indent) => {
    if (stack.length === 0) {
      result.push(`<${type}>`);
      stack.push({ type, indent, hasOpenLi: false });
      return;
    }

    const top = stack[stack.length - 1];
    if (indent > top.indent) {
      result.push(`<${type}>`);
      stack.push({ type, indent, hasOpenLi: false });
      return;
    }

    if (indent === top.indent && type !== top.type) {
      if (top.hasOpenLi) {
        result.push('</li>');
        top.hasOpenLi = false;
      }
      result.push(`</${top.type}>`);
      stack.pop();
      result.push(`<${type}>`);
      stack.push({ type, indent, hasOpenLi: false });
    }
  };

  for (const line of lines) {
    const ulMatch = line.match(/^(\s*)[-*+]\s+(.+)$/);
    const olMatch = line.match(/^(\s*)\d+\.\s+(.+)$/);

    if (!ulMatch && !olMatch) {
      closeAll();
      result.push(line);
      continue;
    }

    const indentStr = (ulMatch || olMatch)[1] || '';
    const indent = countLeadingSpaces(indentStr);
    const type = ulMatch ? 'ul' : 'ol';
    const itemText = (ulMatch || olMatch)[2];

    closeToIndent(indent);
    ensureList(type, indent);

    const current = stack[stack.length - 1];
    if (current.hasOpenLi) {
      result.push('</li>');
      current.hasOpenLi = false;
    }

    result.push(`<li>${itemText}`);
    current.hasOpenLi = true;
  }

  closeAll();
  return result.join('\n');
}

/**
 * 处理表格
 */
function processTables(html) {
  const lines = html.split('\n');
  const result = [];
  let inTable = false;
  let tableRows = [];

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i].trim();

    if (line.startsWith('|') && line.endsWith('|')) {
      if (/^\|[\s:|-]+\|$/.test(line) && line.includes('-')) {
        continue;
      }

      if (!inTable) {
        inTable = true;
      }

      const cellContent = line.slice(1, -1);
      const cells = cellContent.split('|');
      tableRows.push(cells.map(cell => cell.trim()));
    } else {
      if (inTable) {
        result.push(buildTable(tableRows));
        tableRows = [];
        inTable = false;
      }
      result.push(line);
    }
  }

  if (inTable) {
    result.push(buildTable(tableRows));
  }

  return result.join('\n');
}

function buildTable(rows) {
  if (rows.length === 0) return '';

  const maxCols = Math.max(...rows.map(row => row.length));
  let html = '<table>\n';

  if (rows.length > 0) {
    html += '<thead>\n<tr>\n';
    const headerRow = rows[0];
    for (let i = 0; i < maxCols; i++) {
      const cell = i < headerRow.length ? headerRow[i] : '';
      html += `<th>${cell}</th>\n`;
    }
    html += '</tr>\n</thead>\n';
  }

  if (rows.length > 1) {
    html += '<tbody>\n';
    for (let i = 1; i < rows.length; i++) {
      html += '<tr>\n';
      const bodyRow = rows[i];
      for (let j = 0; j < maxCols; j++) {
        const cell = j < bodyRow.length ? bodyRow[j] : '';
        html += `<td>${cell}</td>\n`;
      }
      html += '</tr>\n';
    }
    html += '</tbody>\n';
  }

  html += '</table>';
  return html;
}

/**
 * 处理段落
 */
function processParagraphs(html) {
  const lines = html.split('\n');
  const result = [];
  let paragraph = [];

  const blockElements = ['<h1', '<h2', '<h3', '<h4', '<h5', '<h6', '<ul', '<ol', '<li', '<table', '<thead', '<tbody', '<tr', '<th', '<td', '<blockquote', '<pre', '<hr', '<div', '</li', '</ul', '</ol', '</table', '</thead', '</tbody', '</tr', '</blockquote', '</pre', '</div'];

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    const trimmedLine = line.trim();

    const isBlock = blockElements.some(tag => trimmedLine.startsWith(tag));

    if (isBlock || trimmedLine === '') {
      if (paragraph.length > 0) {
        result.push(`<p>${paragraph.join(' ')}</p>`);
        paragraph = [];
      }
      if (trimmedLine !== '') {
        result.push(line);
      }
    } else {
      paragraph.push(trimmedLine);
    }
  }

  if (paragraph.length > 0) {
    result.push(`<p>${paragraph.join(' ')}</p>`);
  }

  return result.join('\n');
}

module.exports = {
  markdownToHTML
};
