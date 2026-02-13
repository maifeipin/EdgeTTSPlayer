# EdgeTTSPlayer

基于 [edge-tts](https://github.com/rany2/edge-tts) 的文本转语音工具，支持**流式播放**和 **MP3 导出**。

## 功能

- 🎤 **14+ 中文语音** — 使用 Microsoft Edge 神经网络 TTS，语音质量接近真人
- ▶ **流式播放** — 文本自动按标点断句，边生成边播放，双缓冲预加载无缝衔接
- 📝 **实时编辑** — 加载文本文件后可直接编辑，修改自动保存
- 💾 **MP3 导出** — 支持单文件和批量转换
- ⚙️ **可调参数** — 语速、音量滑块，断句最大字数可配置
- 🧹 **自动清理** — 播放结束或停止后临时音频文件自动删除

## 截图

![EdgeTTSPlayer](https://img.shields.io/badge/Python-3.10+-blue?logo=python&logoColor=white)
![License](https://img.shields.io/badge/license-MIT-green)

## 安装

```bash
# 克隆项目
git clone https://github.com/YOUR_USERNAME/EdgeTTSPlayer.git
cd EdgeTTSPlayer

# 安装依赖
pip install -r requirements.txt
```

## 使用

```bash
python main.py
```

1. 点击 **选择文件** 加载 `.txt` 文本
2. 选择语音、调整语速和音量
3. 点击 **▶ 播放** 流式播放，或 **转换为MP3** 导出文件

## 依赖

- Python 3.10+
- [edge-tts](https://github.com/rany2/edge-tts) — Microsoft Edge 在线 TTS 服务
- [pygame](https://www.pygame.org/) — 音频播放

## 注意事项

- 需要**网络连接**（调用 Microsoft Edge 在线 TTS 服务，免费无限制）
- Windows 推荐使用微软雅黑字体以获得最佳中文显示效果

## License

MIT
