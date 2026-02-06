/**
 * 法律文件脱敏处理 - Node.js版本
 *
 * 功能：像Word的"查找替换"一样，遍历所有文本进行替换
 * 用法：node redact.js <input.docx> <rules.json> <output.docx>
 */

const fs = require('fs');
const path = require('path');

// 使用adm-zip来处理docx文件（docx本质是zip压缩包）
const AdmZip = require('adm-zip');

// 获取命令行参数
const args = process.argv.slice(2);
const inputFile = args[0];
const rulesFile = args[1];
const outputFile = args[2];

/**
 * 在XML文本中执行替换（支持跨标签替换）
 * @param {string} xmlContent - XML内容
 * @param {Array} redactions - 替换规则数组
 * @returns {string} - 替换后的XML
 */
function replaceInXML(xmlContent, redactions) {
    // 循环处理跨标签替换，直到没有更多匹配
    let result = xmlContent;
    let maxIterations = 100; // 防止无限循环
    let iteration = 0;
    let hasCrossTagMatch;

    do {
        hasCrossTagMatch = false;
        const beforeLength = result.length;
        result = handleCrossTagReplacement(result, redactions);
        const afterLength = result.length;

        // 如果XML长度变化，说明进行了替换
        if (beforeLength !== afterLength) {
            hasCrossTagMatch = true;
        }

        iteration++;
    } while (hasCrossTagMatch && iteration < maxIterations);

    // 然后处理单标签内的替换（作为补充）
    for (const item of redactions) {
        const { original, replacement } = item;
        result = result.replace(/(<w:t[^>]*>)([^<]*?)(<\/w:t>)/g, (match, openTag, content, closeTag) => {
            const newContent = content.split(original).join(replacement);
            return openTag + newContent + closeTag;
        });
    }

    return result;
}

/**
 * 处理跨标签的替换
 * @param {string} xmlContent - XML内容
 * @param {Array} redactions - 替换规则数组
 * @returns {string} - 替换后的XML
 */
function handleCrossTagReplacement(xmlContent, redactions) {
    // 找到所有<w:t>标签及其位置
    const tTagPattern = /<w:t[^>]*>([^<]*)<\/w:t>/g;
    const textNodes = [];
    let match;

    // 重置正则表达式的lastIndex
    tTagPattern.lastIndex = 0;

    while ((match = tTagPattern.exec(xmlContent)) !== null) {
        textNodes.push({
            fullMatch: match[0],           // 完整的<w:t>内容</w:t>
            content: match[1],             // 标签内的文本
            startIndex: match.index,       // 在XML中的开始位置
            endIndex: match.index + match[0].length
        });
    }

    // 标记哪些节点已被处理
    const processed = new Set();
    const matches = [];

    // 尝试拼接相邻标签，查找匹配
    for (let i = 0; i < textNodes.length; i++) {
        // 跳过已处理的节点
        if (processed.has(i)) continue;

        // 尝试拼接最多5个相邻标签
        for (let span = 1; span <= Math.min(5, textNodes.length - i); span++) {
            let combinedText = '';
            for (let j = 0; j < span; j++) {
                combinedText += textNodes[i + j].content;
            }

            // 检查是否有任何规则匹配
            for (const item of redactions) {
                const { original, replacement } = item;

                if (combinedText.includes(original)) {
                    // 找到original在combinedText中的起始位置
                    const matchStart = combinedText.indexOf(original);

                    // 计算从哪个节点开始，以及跨越多少个节点
                    let charCount = 0;
                    let startNodeIndex = i;
                    let endNodeIndex = i;

                    for (let j = 0; j < span; j++) {
                        const nodeLen = textNodes[i + j].content.length;
                        const nodeEnd = charCount + nodeLen;

                        if (charCount <= matchStart && matchStart < nodeEnd) {
                            startNodeIndex = i + j;
                        }

                        const matchEnd = matchStart + original.length;
                        if (charCount < matchEnd && matchEnd <= nodeEnd) {
                            endNodeIndex = i + j;
                            break;
                        }

                        if (charCount < matchEnd && nodeEnd < matchEnd) {
                            endNodeIndex = i + j;
                        }

                        charCount = nodeEnd;
                    }

                    const actualSpan = endNodeIndex - startNodeIndex + 1;

                    // 计算original在第一个节点中的起始偏移
                    let offsetInFirstNode = matchStart;
                    for (let j = 0; j < startNodeIndex - i; j++) {
                        offsetInFirstNode -= textNodes[i + j].content.length;
                    }

                    // 只处理起始节点未被处理的匹配
                    if (!processed.has(startNodeIndex)) {
                        // 记录匹配
                        matches.push({
                            startNodeIndex,
                            actualSpan,
                            original,
                            replacement,
                            offsetInFirstNode
                        });

                        // 标记所有涉及的节点为已处理
                        for (let j = startNodeIndex; j < startNodeIndex + actualSpan; j++) {
                            processed.add(j);
                        }

                        // 找到匹配后跳出内层循环
                        break;
                    }
                }
            }

            // 如果这个节点的起始位置已被处理，跳出span循环
            if (processed.has(i)) break;
        }
    }

    // 如果没有匹配，返回原始内容
    if (matches.length === 0) {
        return xmlContent;
    }

    // 从后向前处理匹配（避免位置偏移）
    let result = xmlContent;
    for (let i = matches.length - 1; i >= 0; i--) {
        const m = matches[i];
        result = performCrossTagReplace(result, textNodes, m.startNodeIndex, m.actualSpan, m.offsetInFirstNode, m.original, m.replacement);
    }

    return result;
}

/**
 * 执行跨标签替换
 * @param {string} xmlContent - 原始XML内容
 * @param {Array} textNodes - 文本节点数组
 * @param {number} startIndex - 开始<t:t>节点索引
 * @param {number} tSpan - 跨越的<t:t>节点数
 * @param {number} offsetInFirstNode - original在第一个节点中的起始偏移
 * @param {string} original - 要替换的原始文本
 * @param {string} replacement - 替换文本
 * @returns {string} - 替换后的XML
 */
function performCrossTagReplace(xmlContent, textNodes, startIndex, tSpan, offsetInFirstNode, original, replacement) {
    const firstTNode = textNodes[startIndex];
    const lastTNode = textNodes[startIndex + tSpan - 1];

    // 特殊情况：tSpan=1，表示单个标签内匹配，使用简单替换
    if (tSpan === 1) {
        const node = firstTNode;
        const nodeContent = node.content;

        if (offsetInFirstNode === 0) {
            // original从节点开头开始，直接替换
            const newContent = nodeContent.replace(original, replacement);
            const newNodeTag = node.fullMatch.replace(nodeContent, newContent);
            return xmlContent.substring(0, node.startIndex) + newNodeTag + xmlContent.substring(node.endIndex);
        } else {
            // original不从节点开头开始，需要保留前缀
            const prefix = nodeContent.substring(0, offsetInFirstNode);
            const suffix = nodeContent.substring(offsetInFirstNode + original.length);
            const newContent = prefix + replacement + suffix;
            const newNodeTag = node.fullMatch.replace(nodeContent, newContent);
            return xmlContent.substring(0, node.startIndex) + newNodeTag + xmlContent.substring(node.endIndex);
        }
    }

    // 跨节点情况：需要合并节点内容，替换，然后重新分配
    // 1. 获取所有相关节点的完整内容
    let fullContent = '';
    for (let i = 0; i < tSpan; i++) {
        fullContent += textNodes[startIndex + i].content;
    }

    // 2. 替换内容
    const newFullContent = fullContent.replace(original, replacement);

    // 3. 找到第一个和最后一个节点的<w:r>边界
    let beforeFirstT = xmlContent.substring(0, firstTNode.startIndex);
    let firstRStart = beforeFirstT.lastIndexOf('<w:r>');
    if (firstRStart === -1) firstRStart = firstTNode.startIndex;

    let afterLastT = xmlContent.substring(lastTNode.endIndex);
    let firstREnd = afterLastT.indexOf('</w:r>');
    if (firstREnd === -1) firstREnd = 0;
    const firstREndPos = lastTNode.endIndex + firstREnd + 6;

    // 4. 检查是否跨越多个run
    let lastRStart = firstRStart;
    let lastREndPos = firstREndPos;

    const betweenRuns = xmlContent.substring(firstREndPos, lastTNode.endIndex);
    if (betweenRuns.includes('<w:r>') && betweenRuns.includes('</w:r>')) {
        // 找到包含最后一个<t:t>的<w:r>起始位置
        let beforeLastT = xmlContent.substring(0, lastTNode.startIndex);
        lastRStart = beforeLastT.lastIndexOf('<w:r>');
        // 最后一个</w:r>的结束位置
        let afterLastT2 = xmlContent.substring(lastTNode.endIndex);
        let lastREnd = afterLastT2.indexOf('</w:r>');
        if (lastREnd === -1) lastREnd = 0;
        lastREndPos = lastTNode.endIndex + lastREnd + 6;
    }

    // 5. 提取整个XML块
    const blockStart = firstRStart;
    const blockEnd = lastREndPos;
    let xmlBlock = xmlContent.substring(blockStart, blockEnd);

    // 6. 在块内替换所有相关的<t:t>标签
    const tPattern = /<w:t[^>]*>([^<]*)<\/w:t>/g;
    let charIndex = 0; // 在newFullContent中的位置
    let nodeIndex = 0; // 当前处理第几个相关节点

    let newXmlBlock = xmlBlock.replace(tPattern, (match, content) => {
        // 检查这个match是否是我们的目标节点之一
        const matchPosInBlock = xmlBlock.indexOf(match);
        if (matchPosInBlock === -1) return match;

        const absolutePos = blockStart + matchPosInBlock;

        // 检查是否在目标节点范围内
        const isFirstNode = absolutePos >= firstTNode.startIndex && absolutePos < firstTNode.endIndex;
        const isLastNode = absolutePos >= lastTNode.startIndex && absolutePos < lastTNode.endIndex;
        const isMiddleNode = !isFirstNode && !isLastNode && nodeIndex > 0 && nodeIndex < tSpan - 1;

        if (isFirstNode || isLastNode || isMiddleNode) {
            // 这是目标节点之一
            if (nodeIndex === 0) {
                // 第一个节点：保留前缀 + replacement
                const firstNode = textNodes[startIndex];
                const prefix = firstNode.content.substring(0, offsetInFirstNode);
                nodeIndex++;
                return match.replace(content, prefix + replacement);
            } else {
                // 后续节点：清空内容
                nodeIndex++;
                return match.replace(content, '');
            }
        }

        return match;
    });

    return xmlContent.substring(0, blockStart) + newXmlBlock + xmlContent.substring(blockEnd);
}

/**
 * 处理docx文件
 * @param {string} inputFile - 输入文件路径
 * @param {Array} redactions - 替换规则
 * @param {string} outputFile - 输出文件路径
 */
function processDocx(inputFile, redactions, outputFile) {
    console.log(`正在处理: ${inputFile}`);
    console.log(`替换规则: ${redactions.length}条`);

    // 按original长度降序排序（先处理更长的文本）
    const sortedRedactions = [...redactions].sort((a, b) => b.original.length - a.original.length);

    // 使用AdmZip读取docx文件
    const zip = new AdmZip(inputFile);

    // 获取document.xml（主要内容）
    let documentXml = zip.readAsText('word/document.xml');

    // 执行替换
    documentXml = replaceInXML(documentXml, sortedRedactions);

    // 更新zip中的文件
    zip.addFile('word/document.xml', Buffer.from(documentXml, 'utf-8'));

    // 处理header和footer（如果有）
    const zipEntries = zip.getEntries();
    for (const entry of zipEntries) {
        if (entry.entryName.startsWith('word/header/') || entry.entryName.startsWith('word/footer/')) {
            let content = zip.readAsText(entry.entryName);
            content = replaceInXML(content, redactions);
            zip.addFile(entry.entryName, Buffer.from(content, 'utf-8'));
        }
    }

    // 写入输出文件
    zip.writeZip(outputFile);

    console.log(`成功: 脱敏文件已保存到 ${outputFile}`);
    console.log(`共处理 ${redactions.length} 条替换规则`);
}

/**
 * 主函数
 */
function main() {
    if (args.length < 3) {
        console.log('用法: node redact.js <input.docx> <rules.json> <output.docx>');
        console.log('');
        console.log('示例:');
        console.log('  node redact.js input.docx rules.json output.docx');
        console.log('');
        console.log('依赖安装:');
        console.log('  npm install adm-zip');
        process.exit(1);
    }

    // 检查输入文件是否存在
    if (!fs.existsSync(inputFile)) {
        console.error(`错误: 输入文件不存在: ${inputFile}`);
        process.exit(1);
    }

    // 读取规则文件
    let rules;
    try {
        const rulesContent = fs.readFileSync(rulesFile, 'utf-8');
        rules = JSON.parse(rulesContent);
    } catch (e) {
        console.error(`错误: 无法读取规则文件 ${rulesFile}`);
        console.error(e.message);
        process.exit(1);
    }

    // 获取替换规则
    const redactions = rules.redactions || [];
    if (redactions.length === 0) {
        console.error('错误: 规则文件中没有替换规则');
        process.exit(1);
    }

    // 处理文件
    try {
        processDocx(inputFile, redactions, outputFile);
    } catch (e) {
        console.error('错误: 处理失败');
        console.error(e);
        process.exit(1);
    }
}

// 运行主函数
if (require.main === module) {
    main();
}

module.exports = { processDocx, replaceInXML };
