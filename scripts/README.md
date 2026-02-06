# Python脚本使用说明

## 环境要求

- Python 3.7+
- python-docx库

## 安装依赖

```bash
pip install python-docx
```

或使用requirements.txt：

```bash
pip install -r requirements.txt
```

## 脚本说明

### redact.py - 脱敏脚本

将docx文件中的敏感信息进行替换脱敏。

#### 用法

```bash
python redact.py <input.docx> <rules.json> [-o output.docx] [--log log.json]
```

#### 参数

| 参数 | 必填 | 说明 |
|------|------|------|
| `input.docx` | 是 | 输入的原始docx文件 |
| `rules.json` | 是 | 脱敏规则配置文件（前端导出） |
| `-o output.docx` | 否 | 输出文件名（默认：input_脱敏.docx） |
| `--log log.json` | 否 | 输出处理日志（可选） |

#### 示例

```bash
# 基础用法
python redact.py 合同.docx rules-export.json

# 指定输出文件
python redact.py 合同.docx rules-export.json -o 合同_脱敏版.docx

# 输出处理日志
python redact.py 合同.docx rules-export.json -o 合同_脱敏版.docx --log process.log
```

#### 功能特点

- 保留原文格式（段落、表格、字体样式）
- 支持段落和表格中的内容脱敏
- 高亮显示脱敏内容（黄色背景）
- 生成处理日志

### restore.py - 还原脚本

根据比对文档将脱敏文件还原为原始内容。

#### 用法

```bash
python restore.py <redacted.docx> <mapping.md> [-o restored.docx]
```

#### 参数

| 参数 | 必填 | 说明 |
|------|------|------|
| `redacted.docx` | 是 | 脱敏后的docx文件 |
| `mapping.md` | 是 | 替换比对文档（前端导出） |
| `-o restored.docx` | 否 | 输出文件名（默认：redacted_还原.docx） |

#### 示例

```bash
# 基础用法
python restore.py 合同_脱敏版.docx mapping.md

# 指定输出文件
python restore.py 合同_脱敏版.docx mapping.md -o 合同_还原.docx
```

## 输出文件说明

### 脱敏脚本输出

1. **脱敏后docx文件**: 敏感信息已替换，格式保留
2. **处理日志JSON**: 记录脱敏处理详情

### 还原脚本输出

1. **还原后docx文件**: 内容恢复为原始状态

## 注意事项

1. **文件编码**: 确保rules.json和mapping.md使用UTF-8编码
2. **路径问题**: 文件路径含空格时请用引号包裹
3. **格式保留**: 复杂格式（如页眉页脚）可能需要手动调整
4. **备份文件**: 脱敏处理前建议备份原始文件

## 常见问题

### Q: 为什么脱敏后格式有变化？

A: python-docx在处理复杂格式时可能有差异，建议检查输出文件。

### Q: 可以批量处理多个文件吗？

A: 当前版本不支持批量处理，需要写脚本循环调用。

### Q: 还原后内容不完整怎么办？

A: 检查mapping.md是否完整匹配，确保替换文本一致。
