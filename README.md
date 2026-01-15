# 智能教学助手系统

一个基于AI的机械原理课程教学辅助工具，用于自动批改作业、整理成绩和统计提交情况。

## 功能特性

### 📝 作业批改
- **二人小组作业批改**：使用视觉模型自动批改PDF格式的小组作业，支持多页面识别
- **课程论文批改**：使用文本模型批改课程论文，提供详细评语
- **智能评分**：自动给出分数和评语，支持未提交作业的标记

### 📊 成绩整理
- **二人小组成绩整理**：按学生名单顺序整理各章节成绩
- **十人小组成绩整理**：支持大型小组作业的成绩管理
- **自动匹配**：根据学号自动匹配学生分组信息

### 📈 统计功能
- **PDF文件统计**：统计各章节提交的PDF文件
- **提交情况分析**：识别未提交作业的学生和小组

## 项目结构

```
.
├── 批改作业.py                    # 批改二人小组作业
├── 批改论文.py                    # 批改课程论文
├── 整理二人小组作业成绩.py        # 整理二人小组成绩
├── 整理十人小组作业成绩.py        # 整理十人小组成绩
├── 统计小组作业pdf.py             # 统计小组作业PDF文件
├── 统计课程论文pdf.py             # 统计课程论文PDF文件
├── configs.yaml                   # 配置文件（需自行创建）
├── configs.yaml.example           # 配置文件示例
├── 整理二人小组作业成绩-使用说明.md  # 详细使用说明
└── README.md                      # 本文件
```

## 快速开始

### 1. 环境要求

- Python 3.7+
- 依赖库：
  ```bash
  pip install openpyxl xlrd PyMuPDF Pillow openai pyyaml
  ```

### 2. 配置API

复制配置文件示例并填入你的API信息：

```bash
cp configs.yaml.example configs.yaml
```

编辑 `configs.yaml`：

```yaml
openai:
  api_key: "your-api-key-here"
  base_url: "https://dashscope.aliyuncs.com/compatible-mode/v1"
  
models:
  vision: "qwen-vl-max"  # 视觉模型（批改作业用）
  text: "qwen-max"       # 文本模型（批改论文用）
```

### 3. 使用示例

#### 批改二人小组作业

```bash
python 批改作业.py
```

功能：
- 自动读取指定目录下的PDF作业
- 使用AI视觉模型批改
- 生成Excel格式的批改结果
- 支持断点续传（每5份保存一次）

#### 批改课程论文

```bash
python 批改论文.py
```

功能：
- 提取PDF文本内容
- 使用AI文本模型批改
- 从文件名提取学生信息
- 生成详细评语

#### 整理成绩

```bash
python 整理二人小组作业成绩.py
```

功能：
- 读取学生选课名单
- 匹配分组信息
- 整理各章节成绩
- 按学生顺序输出

## 核心功能说明

### 批改作业

**支持的功能：**
- ✅ 多页面PDF识别
- ✅ 自动评分（80-100分，5分为单位）
- ✅ 生成简短评语
- ✅ 未提交作业标记为0分
- ✅ 按小组序号排序
- ✅ 自动保存进度

**评分标准：**
- 内容完整性
- 准确性
- 规范性

### 批改论文

**支持的功能：**
- ✅ PDF文本提取
- ✅ 自动评分（70-100分，5分为单位）
- ✅ 生成详细评语（100-200字）
- ✅ 学生信息自动识别
- ✅ 断点续传

**评价维度：**
- 内容完整性和逻辑性
- 理论知识掌握程度
- 分析问题的深度
- 文字表达和格式规范性

### 成绩整理

**支持的功能：**
- ✅ 学号自动标准化
- ✅ 分组信息智能匹配
- ✅ 支持空组号行（继承上一组号）
- ✅ 支持单人小组
- ✅ 未分组学生标记
- ✅ 多章节成绩汇总

**输出格式：**
| 序号 | 学号 | 姓名 | 组号 | 第2章 | 第8章 | ... |
|------|------|------|------|-------|-------|-----|

## 文件路径配置

在各脚本开头可以修改文件路径：

```python
# 批改作业.py
root_dir = r"E:/Projects/Intelligent_Teaching_Assistant/小组作业/二人小组作业"
output_filename = r"E:/Projects/Intelligent_Teaching_Assistant/二人小组作业-批改.xlsx"

# 批改论文.py
target_dir = r"E:/Projects/Intelligent_Teaching_Assistant/课程论文"
output_filename = r"E:/Projects/Intelligent_Teaching_Assistant/课程论文-批改.xlsx"

# 整理成绩.py
student_list_file = r"E:\Projects\Intelligent_Teaching_Assistant\2024-2025-秋-302074030-02-机械原理-选课学生名单-2025-09-28.xls"
group_file = r"E:\Projects\Intelligent_Teaching_Assistant\分组-2025-2026-秋-302074030-02-机械原理.xlsx"
grade_file = r"E:\Projects\Intelligent_Teaching_Assistant\二人小组作业-批改.xlsx"
```

## 注意事项

1. **API配置**：确保 `configs.yaml` 文件已正确配置，该文件不会被上传到GitHub
2. **文件格式**：作业和论文需为PDF格式，文件名需符合规范
3. **学号格式**：系统会自动处理学号格式，去除空格和 `.0` 后缀
4. **分组表格式**：支持空组号行和单人小组，详见使用说明
5. **进度保存**：批改过程中会自动保存进度，避免数据丢失

## 常见问题

### Q: 如何排除已批改的章节？

在 `批改作业.py` 中修改 `excluded_paths` 列表：

```python
excluded_paths = [
    "第2章", 
    "第8章",
    # 添加更多要排除的章节
]
```

### Q: 如何调整评分范围？

修改对应脚本中的评分逻辑：

```python
# 批改作业.py - 80-100分
score = max(80, min(100, score))
score = round(score / 5) * 5

# 批改论文.py - 70-100分
score = max(70, min(100, score))
score = round(score / 5) * 5
```

### Q: 批改失败怎么办？

- 检查API配置是否正确
- 确认PDF文件是否损坏
- 查看错误日志信息
- 脚本会自动保存已批改的结果

### Q: 如何处理未提交作业？

运行批改脚本时，输入总组数：

```
请输入总共有多少个小组（直接回车跳过，只批改已提交的）: 40
```

系统会自动标记未提交的小组为0分。

## 技术栈

- **Python 3.7+**
- **OpenAI API**：调用大模型进行批改
- **PyMuPDF (fitz)**：PDF文件处理
- **openpyxl**：Excel文件读写
- **Pillow**：图像处理
- **PyYAML**：配置文件解析

## 更新日志

### v1.0 (2026-01-15)
- ✅ 实现作业和论文自动批改
- ✅ 实现成绩整理功能
- ✅ 支持PDF文件统计
- ✅ 支持未提交作业标记
- ✅ 支持断点续传

## 许可证

本项目仅供教学使用。

## 联系方式

如有问题或建议，请联系项目维护者。
