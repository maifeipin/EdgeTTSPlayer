# EdgeTTSPlayer

基于 [edge-tts](https://github.com/rany2/edge-tts) 的文本转语音工具，支持**流式播放**和 **MP3 导出**。

## 功能

- 🎤 **14+ 中文语音** — 使用 Microsoft Edge 神经网络 TTS，语音质量接近真人
- ▶ **流式播放** — 文本自动按标点断句，边生成边播放，双缓冲预加载无缝衔接
- � **多格式支持** — 支持 TXT、Markdown、HTML、EPUB、MOBI、PDF、DOCX
- �📝 **实时编辑** — 加载文件后可直接编辑文本，修改自动保存
- 💾 **MP3 导出** — 支持单文件和批量转换
- ⚙️ **可调参数** — 语速、音量滑块，断句最大字数可配置
- 🧹 **自动清理** — 播放结束或停止后临时音频文件自动删除

## 支持格式

| 格式 | 扩展名 | 解析方式 |
|------|--------|---------|
| 纯文本 | `.txt` `.md` | 直接读取 |
| 网页 | `.html` `.htm` | BeautifulSoup 提取文本 |
| EPUB 电子书 | `.epub` | ebooklib 解析章节 |
| MOBI 电子书 | `.mobi` | mobi 库解包 + HTML 提取 |
| PDF 文档 | `.pdf` | PyPDF2 逐页提取 |
| Word 文档 | `.docx` | python-docx 段落提取 |

## 安装

```bash
# 克隆项目
git clone https://github.com/maifeipin/EdgeTTSPlayer.git
cd EdgeTTSPlayer

# 安装依赖
pip install -r requirements.txt
```

## 使用

```bash
python main.py
```

1. 点击 **选择文件** 加载文本 / 电子书 / 文档
2. 选择语音、调整语速和音量
3. 点击 **▶ 播放** 流式播放，或 **转换为MP3** 导出文件

## 依赖

- Python 3.10+
- [edge-tts](https://github.com/rany2/edge-tts) — Microsoft Edge 在线 TTS
- [pygame](https://www.pygame.org/) — 音频播放
- [ebooklib](https://github.com/aerkalov/ebooklib) — EPUB 解析
- [mobi](https://pypi.org/project/mobi/) — MOBI 解析
- [PyPDF2](https://pypi.org/project/PyPDF2/) — PDF 文本提取
- [python-docx](https://python-docx.readthedocs.io/) — DOCX 解析
- [beautifulsoup4](https://www.crummy.com/software/BeautifulSoup/) — HTML 文本提取

## 注意事项

- 需要**网络连接**（调用 Microsoft Edge 在线 TTS 服务，免费无限制）
- PDF 提取质量取决于 PDF 内容类型（扫描版 PDF 无法提取文本）
- Windows 推荐使用微软雅黑字体以获得最佳中文显示效果

## License

MIT
