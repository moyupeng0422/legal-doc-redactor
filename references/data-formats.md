# 数据格式说明

## 1. 脱敏规则JSON格式

### 文件结构

```json
{
  "version": "1.0.0",
  "created": "2026-02-03T14:30:00",
  "description": "法律文件脱敏规则配置",
  "rules": [...],
  "contextRules": [...],
  "metadata": {...}
}
```

### 完整示例

```json
{
  "version": "1.0.0",
  "created": "2026-02-03T14:30:00",
  "description": "法律文件脱敏规则配置",
  "rules": [
    {
      "id": "company_name",
      "name": "公司名称",
      "category": "organization",
      "enabled": true,
      "patterns": [
        "([\\u4e00-\\u9fa5]{2,15})(公司|有限公司|股份有限公司|集团|企业|厂|店)"
      ],
      "replacement": "【${role}公司${index}】",
      "contextSensitive": true,
      "priority": 10,
      "description": "识别各类公司名称"
    },
    {
      "id": "date",
      "name": "日期",
      "category": "date",
      "enabled": true,
      "patterns": [
        "\\d{4}年\\d{1,2}月\\d{1,2}日",
        "\\d{4}-\\d{1,2}-\\d{1,2}",
        "\\d{4}/\\d{1,2}/\\d{1,2}"
      ],
      "replacement": "【${context}日期${index}】",
      "contextSensitive": true,
      "priority": 8,
      "description": "识别各种格式的日期"
    },
    {
      "id": "price",
      "name": "价格/金额",
      "category": "price",
      "enabled": true,
      "patterns": [
        "人民币\\s*\\d+(,\\d{3})*(\\.\\d+)?元",
        "\\d+(,\\d{3})*(\\.\\d+)?元",
        "RMB\\s*\\d+(,\\d{3})*(\\.\\d+)?",
        "\\$\\s*\\d+(,\\d{3})*(\\.\\d+)?"
      ],
      "replacement": "【${type}】",
      "contextSensitive": true,
      "priority": 9,
      "description": "识别价格和金额"
    }
  ],
  "contextRules": [
    {
      "id": "party_a",
      "name": "甲方/买方",
      "trigger": "甲方|买方|委托方|发包方|采购方",
      "category": "buyer",
      "prefix": "买方"
    },
    {
      "id": "party_b",
      "name": "乙方/卖方",
      "trigger": "乙方|卖方|受托方|承包方|供应商",
      "category": "seller",
      "prefix": "卖方"
    }
  ],
  "priceContextRules": [
    {
      "trigger": "合同总价|总金额|交易价格|成交价格",
      "type": "合同总价"
    },
    {
      "trigger": "单价|单位价格",
      "type": "单价"
    },
    {
      "trigger": "预付款|首付款|定金",
      "type": "预付款"
    }
  ],
  "metadata": {
    "totalRules": 3,
    "enabledRules": 3,
    "lastModified": "2026-02-03T14:30:00"
  }
}
```

### 字段说明

| 字段 | 类型 | 必填 | 说明 |
|------|------|------|------|
| `id` | string | 是 | 规则唯一标识符 |
| `name` | string | 是 | 规则显示名称 |
| `category` | string | 是 | 脱敏内容分类 |
| `enabled` | boolean | 是 | 是否启用该规则 |
| `patterns` | array | 是 | 正则表达式列表 |
| `replacement` | string | 是 | 替换模板（支持变量） |
| `contextSensitive` | boolean | 是 | 是否启用上下文敏感替换 |
| `priority` | number | 是 | 规则优先级（1-10） |
| `description` | string | 否 | 规则说明 |

### 替换模板变量

| 变量 | 说明 | 示例 |
|------|------|------|
| `${category}` | 内容分类 | organization |
| `${role}` | 上下文角色 | 买方/卖方/第三方 |
| `${index}` | 同类型序号 | 1, 2, 3... |
| `${type}` | 具体类型 | 单价/总价/预付款 |
| `${context}` | 上下文类型 | 签约/交付/生效 |
| `${docType}` | 文档类型 | 合同编号/申请号 |

## 2. 替换比对.md格式

### 文件结构

```markdown
# 脱敏内容替换比对

**文件**: {原始文件名}
**生成时间**: {时间戳}
**脱敏项总数**: {数量}

## 替换明细

| 序号 | 原文 | 替换 | 类型 | 位置 |
|------|------|------|------|------|
... ...

## 统计

- **分类1**: 数量
- **分类2**: 数量
...
```

### 完整示例

```markdown
# 脱敏内容替换比对

**文件**: 技术服务合同.docx
**生成时间**: 2026-02-03 14:30:00
**脱敏项总数**: 18

## 替换明细

| 序号 | 原文 | 替换 | 类型 | 位置 |
|------|------|------|------|------|
| 1 | 示例公司A有限公司 | 【买方公司A】 | organization | 段落2 |
| 2 | 示例公司B股份有限公司 | 【卖方公司B】 | organization | 段落3 |
| 3 | 2024年3月15日 | 【签约日期】 | date | 段落5 |
| 4 | 2024年6月30日 | 【交付日期1】 | date | 段落8 |
| 5 | 2024年12月31日 | 【项目结束日期】 | date | 段落12 |
| 6 | 人民币500,000元 | 【合同总价】 | price | 段落10 |
| 7 | 单价500元/件 | 【单价】 | price | 表格2行3 |
| 8 | 预付款100,000元 | 【预付款】 | price | 段落15 |
| 9 | 尾款50,000元 | 【尾款】 | price | 段落18 |
| 10 | HT-2024-001 | 【合同编号】 | document_id | 段落1 |
| 11 | 202310123456.7 | 【专利申请号】 | document_id | 段落7 |
| 12 | 智能制造系统集成项目 | 【项目名称1】 | project_name | 段落3 |
| 13 | 联系人：张三 | 【联系人姓名】 | person_name | 段落25 |
| 14 | 电话：138xxxx5678 | 【联系电话】 | contact | 段落25 |
| 15 | zhangsan@example.com | 【邮箱】 | contact | 段落26 |
| 16 | 示例省示例市示例区示例街道1号示例大厦 | 【地址1】 | address | 段落4 |
| 17 | 示例直辖市示例区示例路2号示例广场 | 【地址2】 | address | 段落5 |
| 18 | 微信：abc123 | 【微信号】 | contact | 段落27 |

## 统计

- **组织机构**: 2处
- **日期**: 3处
- **价格/金额**: 4处
- **文件编号**: 2处
- **项目名称**: 1处
- **人名**: 1处
- **联系方式**: 3处
- **地址**: 2处

## 还原说明

使用以下命令还原脱敏文件：

```bash
python scripts/restore.py redacted.docx mapping.md -o restored.docx
```
```

## 3. 导出规则JSON格式

### 用于Python脚本的规则格式

```json
{
  "sourceFile": "技术服务合同.docx",
  "exportTime": "2026-02-03T14:30:00",
  "redactions": [
    {
      "id": 1,
      "original": "示例公司A有限公司",
      "replacement": "【买方公司A】",
      "type": "organization",
      "location": "paragraph:2",
      "context": "party_a"
    },
    {
      "id": 2,
      "original": "2024年3月15日",
      "replacement": "【签约日期】",
      "type": "date",
      "location": "paragraph:5",
      "context": "signing_date"
    }
  ]
}
```

### 字段说明

| 字段 | 类型 | 说明 |
|------|------|------|
| `id` | number | 脱敏项唯一ID |
| `original` | string | 原始文本 |
| `replacement` | string | 替换文本 |
| `type` | string | 脱敏类型 |
| `location` | string | 位置标识 |
| `context` | string | 上下文标识 |

## 4. Python脚本输入/输出格式

### redact.py 输入

**命令行参数**:
```bash
python redact.py <input.docx> <rules.json> [-o output.docx]
```

**rules.json 格式**: 参见第1节

### redact.py 输出

1. **脱敏后docx文件**: 保留原格式，敏感信息已替换
2. **替换日志JSON**:
```json
{
  "inputFile": "input.docx",
  "outputFile": "output.docx",
  "processTime": "2026-02-03T14:30:00",
  "totalRedactions": 18,
  "redactions": [...]
}
```

### restore.py 输入

**命令行参数**:
```bash
python restore.py <redacted.docx> <mapping.md> [-o restored.docx]
```

**mapping.md 格式**: 参见第2节

### restore.py 输出

还原后的docx文件，内容恢复为原始状态

## 5. 内部数据结构

### 脱敏项对象

```javascript
{
  id: "uuid",
  original: "原始文本",
  replacement: "替换文本",
  type: "organization",
  startIndex: 100,
  endIndex: 125,
  location: "paragraph:2",
  context: {
    role: "buyer",
    index: 1
  }
}
```

### 文档解析结果

```javascript
{
  paragraphs: [
    {
      index: 0,
      text: "段落文本",
      redactions: [
        { /* 脱敏项对象 */ }
      ]
    }
  ],
  tables: [
    {
      index: 0,
      rows: 3,
      cols: 4,
      cells: [
        {
          row: 0,
          col: 0,
          text: "单元格文本",
          redactions: []
        }
      ]
    }
  ],
  metadata: {
    totalParagraphs: 30,
    totalTables: 5,
    totalRedactions: 18
  }
}
```

## 6. localStorage格式

### 优先级机制

**优先级顺序**：黑名单（最高）> 白名单 > 脱敏类别（内置+自定义）

| 优先级 | 类型 | 说明 |
|--------|------|------|
| 1（最高） | 黑名单 | 强制脱敏，不受白名单限制 |
| 2 | 白名单 | 覆盖检查（包含关系），优先于脱敏类别 |
| 3（最低） | 脱敏类别 | 内置规则 + 自定义类型 |

### 白名单数据（redaction_whitelist）

**存储位置**: `localStorage.redaction_whitelist`

**数据格式**: JSON数组字符串

**示例**:
```json
["示例公司A", "2024年3月15日", "公开邮箱@example.com", "计划"]
```

**覆盖检查**: 白名单采用包含关系检查，如白名单有"计划"，则"股权激励计划"不被识别。

**JavaScript操作**:
```javascript
// 读取
const whitelist = JSON.parse(localStorage.getItem('redaction_whitelist') || '[]');

// 写入
localStorage.setItem('redaction_whitelist', JSON.stringify([...whitelist, '新项']));

// 覆盖检查（检查内容是否包含白名单中的任何一项）
function isCoveredByWhitelist(content) {
    for (const item of whitelist) {
        if (content.includes(item) || item.includes(content)) {
            return true;
        }
    }
    return false;
}
```

### 黑名单数据（redaction_blacklist）

**存储位置**: `localStorage.redaction_blacklist`

**数据格式**: JSON数组字符串（对象格式，v1.1.0+）

**示例**:
```json
[
  {"original": "内部项目代号", "type": "project_name"},
  {"original": "特殊代码格式", "type": "file_code"},
  {"original": "内部术语", "type": "sensitive"}
]
```

**兼容性**: 自动兼容旧版字符串格式 `["内部项目代号", "特殊代码格式"]`

**JavaScript操作**:
```javascript
// 读取
const blacklist = JSON.parse(localStorage.getItem('redaction_blacklist') || '[]');

// 添加（对象格式）
blacklist.push({ original: '新项', type: 'organization' });
localStorage.setItem('redaction_blacklist', JSON.stringify(blacklist));

// 检查（遍历对象数组）
const inBlacklist = blacklist.some(item =>
    (typeof item === 'string' ? item : item.original) === '待检查内容'
);
```

### 自定义类型数据（redaction_customTypes）

**存储位置**: `localStorage.redaction_customTypes`

**数据格式**: JSON对象，键为类型名称，值为项目数组

**示例**:
```json
{
  "合同名称": [
    {"original": "技术服务合同", "id": 1736789123456},
    {"original": "软件开发协议", "id": 1736789123457}
  ],
  "产品型号": [
    {"original": "ABC-100", "id": 1736789123458},
    {"original": "XYZ-200", "id": 1736789123459}
  ]
}
```

**JavaScript操作**:
```javascript
// 读取
const customTypes = JSON.parse(localStorage.getItem('redaction_customTypes') || '{}');

// 添加类型
if (!customTypes['新类型']) {
    customTypes['新类型'] = [];
}
customTypes['新类型'].push({ original: '内容', id: Date.now() });
localStorage.setItem('redaction_customTypes', JSON.stringify(customTypes));

// 删除类型
delete customTypes['类型名'];
localStorage.setItem('redaction_customTypes', JSON.stringify(customTypes));
```

### 数据共享

白名单/黑名单/自定义类型数据在所有使用该工具的文件间共享，保持一致性。

## 7. 调试日志格式

### console.log覆盖机制

系统会覆盖 `console.log` 来收集所有日志输出：

```javascript
const originalLog = console.log;
const debugLogs = [];

console.log = function(...args) {
  debugLogs.push({
    timestamp: new Date().toISOString(),
    message: args.map(arg => typeof arg === 'object' ? JSON.stringify(arg) : arg).join(' ')
  });
  originalLog.apply(console, args);
};
```

### 日志内容

调试日志包含以下信息：

| 类型 | 说明 | 示例 |
|------|------|------|
| 规则匹配 | 显示哪些规则匹配到内容 | `匹配规则: organization, 文本: 示例公司` |
| 白名单检查 | 显示白名单过滤结果 | `白名单跳过: 示例公司A` |
| 黑名单检查 | 显示黑名单强制脱敏 | `黑名单强制脱敏: 内部代号` |
| 替换操作 | 显示替换详情 | `替换: 示例公司 → 【公司1】` |
| 错误警告 | 显示处理错误 | `警告: 未匹配到任何规则` |

### 导出格式

**文件名**: `debug_log_YYYYMMDD_HHMMSS.txt`

**内容格式**:
```
调试日志导出时间: 2026-02-06 14:30:00

[2026-02-06T14:30:00.123Z] 开始处理文档: 合同.docx
[2026-02-06T14:30:00.456Z] 段落数: 30, 表格数: 5
[2026-02-06T14:30:00.789Z] 匹配规则: organization, 文本: 示例公司A有限公司
[2026-02-06T14:30:01.012Z] 白名单检查: 不在白名单
[2026-02-06T14:30:01.234Z] 黑名单检查: 不在黑名单
[2026-02-06T14:30:01.456Z] 替换: 示例公司A有限公司 → 【公司1】
[2026-02-06T14:30:02.789Z] 处理完成: 总计18个脱敏项
```

### 导出操作

点击"📥 导出日志"按钮后：

1. 收集所有调试日志
2. 格式化为带时间戳的文本
3. 创建 Blob 对象
4. 触发浏览器下载
