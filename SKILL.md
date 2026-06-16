---
name: lesson-plan-from-images
description: 将小学数学例题图片转化为结构化Excel教案的工作流。当用户提供参考模板Excel和数学例题图片（PNG/JPG），要求生成包含教学目标、分步互动话术、师生问答、小结和做题回收的40分钟教案时，启用此Skill。
agent_created: true
---

# 图片例题教案生成器

## Overview

本 Skill 封装了"参考模板 + 例题图片 → Excel 教案"的标准工作流。适用于小学数学互动课程设计，用户只需提供参考模板（Excel格式）和例题图片（展示题目内容），即可自动生成包含完整教学互动话术的教案文件。

## 核心输入

| 输入 | 说明 | 必须 |
|------|------|------|
| 参考模板Excel | 具有标准7列结构的教案模板（再见我的星球-3星-切割拼接周长.xlsx） | ✅ |
| 例题图片（PNG/JPG） | 每张图片对应一个教学环节的例题 | ✅ |
| 年级 | 如：小学三年级 | ✅ |
| 课程时长 | 如：40分钟 | ✅ |

## 核心输出

Excel 教案文件，结构如下：

| 列 | 内容 | 说明 |
|----|------|------|
| A列 | 页面 | 如 P1-P2、P4-1 等 |
| B列 | 时间 | 如 课前3-5min、5-7min |
| C列 | 教学目标 | 知识/能力/情感三维目标 |
| D-G列 | 教学设计与话术演绎（合并单元格） | 核心教学互动内容 |

## 完整工作流

### Step 1 — 读取参考模板

1. 读取用户提供的参考模板 Excel 文件
2. 识别其格式规范：
   - 7列结构：A=页面、B=时间、C=教学目标、D-G=教学设计与话术演绎（合并）
   - 列宽设置、单元格格式
   - 模块标题行（蓝底白字合并）
   - 列头行（蓝底白字）
   - 普通单元格格式

### Step 2 — 分析例题图片

1. 用 Read 工具读取每张例题图片
2. 分析图中包含的数学知识点：
   - 题目的具体要求（是什么类型的题？）
   - 给出了哪些信息（数字、图形、条件）
   - 需要学生完成什么操作
   - 涉及哪些数学概念或规律
3. 归纳每个例题对应的教学环节定位（引入/进阶/挑战）

### Step 3 — 设计教案结构

根据例题数量确定模块数量，一般结构为：

```
大标题行（课程名称 + 年级 + 时长）

课上破冰
├─ 学习习惯检查（点名表扬）
└─ 互动提问（3-4个引导问题，学生回答在括号内）

环节一（例题1）
├─ 过渡页（30s）
├─ 知识/能力/情感目标
├─ 分步教学话术（学生自主读题 → 教师分步引导 → 追问 → 小结）
├─ 小结（3条核心规则）
└─ 做题回收（4种差异化场景话术）

环节二（例题2）
└─ 同上结构

环节三（例题3）
└─ 同上结构

课堂总结
├─ 自主复盘5问
├─ 差异化表扬
└─ 作业提醒
```

### Step 4 — 编写分步教学话术

每道例题的讲解遵循以下结构：

**一、学生自主读题，教师就题目关键信息进行提问**

- 提问1：题目让我们做什么？有什么特别的要求？
- 提问2：这个图形/立体由几个小正方体组成？它们是怎么摆放的？
- 提问3-5：针对具体数学规律/逻辑的引导问题
  （引导学生回答：xxx）

**二、教师引导学生将思考与操作进行连线匹配**

- 逐步追问，引导学生说出判断依据
- 最后一问：谁能告诉老师，为什么……？

**三、教师小结（3条）**

- 总结本环节的核心规律/方法
- 强调学生容易出错的地方

**四、做事回收（4种差异化场景）**

```
表现突出（全对）：XX做题后主动检查了两遍，所以两题都对啦，奖励五颗小星星！
有明显进步：也要表扬XX！还记得上次我们学分数时还有点迷糊，今天你全部自己分对了，每一步都特别清晰！
努力但结果有误：今天最让老师感动的是XX，虽然有的题目没有完成，但是XX一直没有放弃！
做题习惯不好的：我刚才看到有同学因为粗心反复订正了几次，提交之前要记得检查哦！
```

### Step 5 — 生成Excel脚本并执行

使用 xlsxwriter 库生成 Excel，参考 `scripts/generate_lesson_plan.py` 的代码框架。

关键格式规范（必须严格遵循）：

```python
# 标题行格式（蓝底白字）
fmt_title = workbook.add_format({
    'bold': True, 'font_size': 14, 'font_color': 'white',
    'bg_color': '#4472C4', 'align': 'center', 'valign': 'vcenter',
    'text_wrap': True, 'border': 1
})

# 模块标题格式（蓝底白字）
fmt_sec = workbook.add_format({
    'bold': True, 'font_size': 12, 'font_color': 'white',
    'bg_color': '#4472C4', 'align': 'center', 'valign': 'vcenter',
    'text_wrap': True, 'border': 1
})

# 列头格式（蓝底白字）
fmt_hc = workbook.add_format({
    'bold': True, 'font_size': 11, 'font_color': 'white',
    'bg_color': '#4472C4', 'align': 'center', 'valign': 'vcenter',
    'text_wrap': True, 'border': 1
})

# 普通单元格格式
fmt_cell = workbook.add_format({
    'font_size': 10, 'align': 'left', 'valign': 'top',
    'text_wrap': True, 'border': 1
})

# 居中对齐单元格格式
fmt_cc = workbook.add_format({
    'font_size': 10, 'align': 'center', 'valign': 'vcenter',
    'text_wrap': True, 'border': 1
})

# 教学目标格式
fmt_goal = workbook.add_format({
    'font_size': 10, 'align': 'left', 'valign': 'vcenter',
    'text_wrap': True, 'border': 1
})

# 列宽
worksheet.set_column('A:A', 10)   # 页面
worksheet.set_column('B:B', 12)   # 时间
worksheet.set_column('C:C', 24)   # 教学目标
worksheet.set_column('D:G', 28)   # 教学设计与话术演绎

# 统一行高
for r in range(0, 80):
    worksheet.set_row(r, 90)
```

### Step 6 — 执行脚本并验证

1. 使用 managed Python 路径运行脚本
2. 检查生成的 Excel 文件是否存在且大小合理
3. 用 `check_xlsx.py` 验证内容完整性（可选）

## 关键注意事项

- **中文编码**：Python 脚本中的中文字符建议使用 Unicode 转义序列（`\uXXXX`）避免编码问题，或使用 `# -*- coding: utf-8 -*-` 并确保文件保存为 UTF-8
- **中文字符建议用Unicode转义**：为避免 PowerShell 执行时的编码问题，merge_range 内容中的中文使用 `\uXXXX` 转义序列
- **输出路径**：保存到用户工作文件夹（`C:\Users\Administrator\Desktop\教研写作`），文件名包含课题名、年级、时长版本号
- **格式一致性**：严格按照参考模板的格式，不自行更改列宽、行高或颜色

## Resources

> **Note**: The following resources are referenced in this Skill's workflow. They are bundled in the source repository (`scripts/` and `references/` directories) and should be included when packaging.

- `scripts/generate_lesson_plan.py` — Excel 教案生成器脚本框架（含格式定义和标准模块结构）
- `references/template_format.md` — Excel 模板格式规范参考
- `references/dialogue_templates.md` — 教学话术模板参考（破冰/分步引导/小结/做题回收）

