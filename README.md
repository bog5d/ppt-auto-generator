# PPT自动生成器 v3.8

🚀 **一款基于AI的智能PPT生成工具**，支持从JSON配置自动生成专业PPT，包含AI图片生成、智能排版、多主题支持等功能。

## ✨ 核心功能

- 📄 **JSON驱动** - 通过JSON配置文件定义PPT结构和内容
- 🎨 **4种主题** - 军事庄重、科技蓝、自然绿、商务灰
- 🤖 **AI图片生成** - 集成硅基流动API，自动生成配图
- 📐 **6种布局** - 左文右图、右文左图、上文下图等
- 📝 **智能换行** - 长文本自动换行，避免溢出
- 💡 **金句避让** - 金句位置智能调整，不与图片重叠

## 📦 安装依赖

```bash
pip install python-pptx requests
```

## 🚀 快速开始

### 1. 运行程序

```bash
python ppt_generator_v3.8_完美版.py
```

### 2. 选择JSON配置文件

程序会提示你选择：
- `[1]` 使用内置示例
- `[2]` 指定自定义JSON文件

### 3. 生成图片（可选）

选择是否使用AI生成配图：
- `[1]` 是（推荐，使用硅基流动AI）
- `[4]` 否（使用占位图）

### 4. 选择主题并生成

选择喜欢的主题配色，输入输出文件名，即可生成PPT！

## 📋 JSON配置格式

```json
{
  "metadata": {
    "title": "演示标题",
    "theme": "tech_blue"
  },
  "slides": [
    {
      "type": "cover",
      "title": "主标题",
      "subtitle": "副标题",
      "slogan": "口号"
    },
    {
      "type": "section",
      "title": "章节标题"
    },
    {
      "type": "content_image",
      "title": "内容页标题",
      "bullets": [
        "要点1：说明内容",
        "要点2：说明内容"
      ],
      "image_prompt": "AI image generation prompt",
      "image_desc": "图片描述",
      "quote": "金句内容"
    },
    {
      "type": "ending",
      "title": "结束语",
      "bullets": ["总结1", "总结2"],
      "quote": "结束金句"
    }
  ]
}
```

## 🎨 支持的幻灯片类型

| 类型 | 说明 |
|------|------|
| `cover` | 封面页 |
| `section` | 章节分隔页 |
| `content_image` | 图文混排页 |
| `chart` | 图表页 |
| `ending` | 结束页 |

## 🖼️ AI图片生成

程序集成了硅基流动(SiliconFlow)的FLUX模型：
- 自动从幻灯片标题和内容生成英文提示词
- 支持API限流自动重试
- 生成1024x1024高质量图片

## 🧪 运行测试

```bash
python test_ppt_auto.py
```

## 📁 项目结构

```
├── ppt_generator_v3.8_完美版.py  # 主程序
├── test_ppt_auto.py              # 自动化测试
├── example_config.json           # 示例配置
└── README.md                     # 说明文档
```

## 📄 License

MIT License

## 🤝 贡献

欢迎提交Issue和Pull Request！

---

Made with ❤️ by AI资源指挥官
