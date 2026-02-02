/**
 * HTML 到 DOCX 转换器（简化版）
 * 将 HTML 转换为 docx.js 组件结构
 * 适用于 Node.js 环境，使用 jsdom 提供 DOM 支持
 */

const { JSDOM } = require('jsdom');
const { Paragraph, TextRun, HeadingLevel, Table, TableRow, TableCell, WidthType, BorderStyle, AlignmentType, VerticalAlign } = require('docx');
const { charsToTwips } = require('./styles');

// Node 类型常量
const NODE_TYPE = {
  TEXT_NODE: 3,
  ELEMENT_NODE: 1
};

/**
 * 将 HTML 字符串转换为 docx 组件数组
 * @param {string} htmlString - HTML 字符串
 * @returns {Array} docx 组件数组
 */
function convertHTMLToDocx(htmlString) {
  const dom = new JSDOM(`<body>${htmlString}</body>`);
  const temp = dom.window.document.body;

  const children = [];

  for (const node of temp.childNodes) {
    const converted = convertNode(node);
    if (converted) {
      if (Array.isArray(converted)) {
        children.push(...converted);
      } else {
        children.push(converted);
      }
    }
  }

  return children.length > 0 ? children : [new Paragraph({ text: '' })];
}

/**
 * 转换单个 DOM 节点
 */
function convertNode(node) {
  // 文本节点
  if (node.nodeType === NODE_TYPE.TEXT_NODE) {
    const text = node.textContent.trim();
    if (text) {
      return new Paragraph({
        children: [new TextRun(text)],
        indent: { firstLine: charsToTwips(2) }
      });
    }
    return null;
  }

  // 元素节点
  if (node.nodeType === NODE_TYPE.ELEMENT_NODE) {
    const tagName = node.nodeName.toUpperCase();

    switch (tagName) {
      case 'H1':
        return new Paragraph({
          text: node.textContent,
          heading: HeadingLevel.HEADING_1
        });

      case 'H2':
        return new Paragraph({
          text: node.textContent,
          heading: HeadingLevel.HEADING_2
        });

      case 'H3':
        return new Paragraph({
          text: node.textContent,
          heading: HeadingLevel.HEADING_3
        });

      case 'H4':
      case 'H5':
      case 'H6':
        return new Paragraph({
          text: node.textContent,
          heading: HeadingLevel.HEADING_3
        });

      case 'P':
        return convertParagraph(node);

      case 'PRE':
        return convertCodeBlock(node);

      case 'BLOCKQUOTE':
        return convertBlockquote(node);

      case 'TABLE':
        return convertTable(node);

      case 'UL':
      case 'OL':
        return convertList(node);

      case 'IMG':
        return convertImage(node);

      case 'HR':
        return new Paragraph({
          children: [new TextRun({ text: '─'.repeat(50), color: 'CCCCCC' })],
          alignment: AlignmentType.CENTER
        });

      case 'BR':
        return new Paragraph({ text: '' });

      case 'DIV':
      case 'SECTION':
      case 'ARTICLE':
      case 'SPAN':
        return convertChildren(node);

      default:
        return convertChildren(node);
    }
  }

  return null;
}

/**
 * 转换段落元素
 */
function convertParagraph(pElement) {
  const runs = convertInlineNodes(pElement.childNodes);
  return new Paragraph({
    children: runs.length > 0 ? runs : [new TextRun('')],
    indent: { firstLine: charsToTwips(2) }
  });
}

/**
 * 转换代码块
 */
function convertCodeBlock(preElement) {
  const codeElement = preElement.querySelector('code') || preElement;
  const codeText = codeElement.textContent;
  const lines = codeText.split('\n');

  const textRuns = [];
  lines.forEach((line, index) => {
    textRuns.push(new TextRun({
      text: line || ' ',
      font: "Consolas",
      size: 22,
      color: "1F2937"
    }));

    if (index < lines.length - 1) {
      textRuns.push(new TextRun({ text: '', break: 1 }));
    }
  });

  return new Paragraph({
    children: textRuns,
    style: "CodeBlock"
  });
}

/**
 * 转换图片元素
 */
function convertImage(imgElement) {
  const alt = imgElement.getAttribute('alt') || '';
  const src = imgElement.getAttribute('src') || '';

  let description = '[图片]';
  if (alt) {
    description = `[图片: ${alt}]`;
  }
  if (src && !src.startsWith('data:')) {
    description += ` (${src})`;
  }

  return new Paragraph({
    children: [new TextRun({ text: description, italics: true, color: "6B7280" })],
    spacing: { before: 200, after: 200 }
  });
}

/**
 * 转换列表元素
 */
function convertList(listElement, level = 0) {
  const paragraphs = [];
  const isOrdered = listElement.nodeName === 'OL';
  const reference = isOrdered ? "numbered-list" : "bullet-list";

  const listItems = listElement.querySelectorAll(':scope > li');
  for (const li of listItems) {
    // 检查是否有嵌套列表
    const nestedList = li.querySelector(':scope > ul, :scope > ol');

    // 获取列表项的直接文本内容
    const runs = convertInlineNodes(li.childNodes, true);

    paragraphs.push(new Paragraph({
      children: runs.length > 0 ? runs : [new TextRun('')],
      numbering: { reference: reference, level: level }
    }));

    // 处理嵌套列表
    if (nestedList) {
      const nestedParagraphs = convertList(nestedList, level + 1);
      paragraphs.push(...nestedParagraphs);
    }
  }

  return paragraphs;
}

/**
 * 转换表格
 */
function convertTable(tableElement) {
  const rows = [];

  // 处理表头
  const thead = tableElement.querySelector('thead');
  if (thead) {
    const headerRows = thead.querySelectorAll('tr');
    headerRows.forEach(tr => {
      const cells = [];
      tr.querySelectorAll('th, td').forEach(cell => {
        const cellContent = convertCellContent(cell);

        // 表头加粗
        const headerCellContent = cellContent.map(item => {
          if (item.constructor.name === 'Paragraph') {
            const children = item.options?.children || [];
            const boldChildren = children.map(child => {
              if (child.constructor.name === 'TextRun') {
                return new TextRun({ ...child.options, bold: true, size: 24 });
              }
              return child;
            });
            return new Paragraph({ ...item.options, children: boldChildren, alignment: AlignmentType.CENTER });
          }
          return item;
        });

        cells.push(new TableCell({
          children: headerCellContent,
          shading: { fill: "E5E7EB" },
          verticalAlign: VerticalAlign.CENTER,
          margins: { top: 120, bottom: 120, left: 150, right: 150 }
        }));
      });

      if (cells.length > 0) {
        rows.push(new TableRow({ children: cells, tableHeader: true }));
      }
    });
  }

  // 处理表体
  const tbody = tableElement.querySelector('tbody') || tableElement;
  const bodyRows = tbody.querySelectorAll('tr');
  bodyRows.forEach(tr => {
    if (thead && thead.contains(tr)) return;

    const cells = [];
    tr.querySelectorAll('th, td').forEach(cell => {
      const cellContent = convertCellContent(cell);
      cells.push(new TableCell({
        children: cellContent,
        verticalAlign: VerticalAlign.CENTER,
        margins: { top: 100, bottom: 100, left: 150, right: 150 }
      }));
    });

    if (cells.length > 0) {
      rows.push(new TableRow({ children: cells }));
    }
  });

  return new Table({
    rows: rows,
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: {
      top: { style: BorderStyle.SINGLE, size: 6, color: "9CA3AF" },
      bottom: { style: BorderStyle.SINGLE, size: 6, color: "9CA3AF" },
      left: { style: BorderStyle.SINGLE, size: 6, color: "9CA3AF" },
      right: { style: BorderStyle.SINGLE, size: 6, color: "9CA3AF" },
      insideHorizontal: { style: BorderStyle.SINGLE, size: 4, color: "D1D5DB" },
      insideVertical: { style: BorderStyle.SINGLE, size: 4, color: "D1D5DB" }
    }
  });
}

/**
 * 转换表格单元格内容
 */
function convertCellContent(cellElement) {
  const runs = convertInlineNodes(cellElement.childNodes);
  return [new Paragraph({
    children: runs.length > 0 ? runs : [new TextRun('')],
    alignment: cellElement.nodeName === 'TH' ? AlignmentType.CENTER : AlignmentType.LEFT
  })];
}

/**
 * 转换引用块
 */
function convertBlockquote(blockquoteElement) {
  const paragraphs = [];

  for (const child of blockquoteElement.childNodes) {
    if (child.nodeName === 'P') {
      const runs = convertInlineNodes(child.childNodes);
      paragraphs.push(new Paragraph({ children: runs, style: "Quote" }));
    } else if (child.nodeType === NODE_TYPE.TEXT_NODE) {
      const text = child.textContent.trim();
      if (text) {
        paragraphs.push(new Paragraph({ children: [new TextRun(text)], style: "Quote" }));
      }
    }
  }

  return paragraphs.length > 0 ? paragraphs : [new Paragraph({
    text: blockquoteElement.textContent,
    style: "Quote"
  })];
}

/**
 * 递归转换容器的子元素
 */
function convertChildren(containerElement) {
  const children = [];

  for (const child of containerElement.childNodes) {
    const converted = convertNode(child);
    if (converted) {
      if (Array.isArray(converted)) {
        children.push(...converted);
      } else {
        children.push(converted);
      }
    }
  }

  return children;
}

/**
 * 转换内联元素为 TextRun 数组
 * @param {NodeList} nodes - 节点列表
 * @param {boolean} skipNestedLists - 是否跳过嵌套列表
 */
function convertInlineNodes(nodes, skipNestedLists = false) {
  const runs = [];

  for (const node of nodes) {
    if (node.nodeType === NODE_TYPE.TEXT_NODE) {
      const text = node.textContent;
      if (text && text.trim()) {
        runs.push(new TextRun(text));
      }
    } else if (node.nodeType === NODE_TYPE.ELEMENT_NODE) {
      const tagName = node.nodeName.toUpperCase();

      // 跳过嵌套列表
      if (skipNestedLists && (tagName === 'UL' || tagName === 'OL')) {
        continue;
      }

      switch (tagName) {
        case 'STRONG':
        case 'B':
          runs.push(new TextRun({ text: node.textContent, bold: true }));
          break;

        case 'EM':
        case 'I':
          runs.push(new TextRun({ text: node.textContent, italics: true }));
          break;

        case 'DEL':
        case 'S':
          runs.push(new TextRun({ text: node.textContent, strike: true }));
          break;

        case 'CODE':
          runs.push(new TextRun({
            text: node.textContent,
            font: "Consolas",
            size: 22,
            color: "DC2626"
          }));
          break;

        case 'A':
          runs.push(new TextRun({
            text: node.textContent,
            color: "2563EB",
            underline: {}
          }));
          break;

        case 'BR':
          runs.push(new TextRun({ text: '', break: 1 }));
          break;

        default:
          runs.push(...convertInlineNodes(node.childNodes, skipNestedLists));
      }
    }
  }

  return runs;
}

module.exports = {
  convertHTMLToDocx
};
