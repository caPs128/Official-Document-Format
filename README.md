# -
世界上最大的岛是领导
# 公文排版技能 (gongwen-format)

将 Markdown / 纯文本 转换为符合 **GB/T 9704-2012 党政机关公文格式** 的 Word 文档。

## 快速开始

```bash
# 1. 安装依赖
pip install python-docx

# 2. 复制生成脚本到工作目录
cp scripts/generate.py /path/to/your/project/

# 3. 修改 generate.py 中的 build_content(doc) 函数
#    ——填入你的实际文档内容

# 4. 运行
python generate.py

# 5. 输出文件：output.docx（可在 Word 或 WPS 中打开）
```

## 格式规范一览

| 要素 | 规范 | 备注 |
|------|------|------|
| 纸张 | A4 (210mm×297mm) | — |
| 上边距 | 37mm | — |
| 下边距 | 35mm | — |
| 左边距 | 28mm | — |
| 右边距 | 26mm | — |
| 版心 | 156mm×225mm | — |
| 标题字体 | 2号(22pt) 方正小标宋简体 | 居中，无加粗 |
| 正文字体 | 3号(16pt) 仿宋_GB2312 | 首行缩进2字符 |
| 一级标题 | 3号(16pt) 黑体 | 序号"一、" |
| 二级标题 | 3号(16pt) 楷体_GB2312 | 序号"（一）" |
| 数字/英文 | Times New Roman | 全文统一 |
| 行距 | 固定值 28.9 磅 | — |
| 每行字数 | 28 字 | 3号字体自然结果 |
| 每页行数 | 22 行 | 28.9pt行距自然结果 |

## 字体回退链

当系统未安装指定字体时，Word 自动使用：

```
方正小标宋简体 → 宋体 (SimSun)
仿宋_GB2312   → 仿宋 (FangSong)
楷体_GB2312   → 楷体 (KaiTi)
黑体          → 系统自带黑体 (SimHei)
```

## 文件结构

```
gongwen-format/
├── SKILL.md           # 技能定义（给 AI Agent 看的）
├── README.md          # 本文件（给人类看的）
└── scripts/
    └── generate.py    # 可复用的 Python 生成脚本
```

## 触发条件

在 Claude Code 中提到以下关键词会自动加载此技能：
- 公文格式 / 政府公文 / 红头文件 / 党政机关文档
- 正式公文排版 / GB/T 9704 / A4排版
- 28字22行 / 方正小标宋 / 仿宋三号

## 许可

参考实现，可自由使用和修改。
