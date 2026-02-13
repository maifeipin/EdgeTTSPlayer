import pathlib
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog, messagebox
import edge_tts
import asyncio
import threading
import os
import platform
import webbrowser
import re
import tempfile
import shutil
from datetime import datetime

import pygame


# edge-tts 默认中文语音
DEFAULT_VOICE = "zh-CN-XiaoxiaoNeural"

# 断句标点
SENTENCE_DELIMITERS = re.compile(r'(?<=[。！？；…!?;])|(?<=\n)')
CLAUSE_DELIMITERS = re.compile(r'(?<=[，、,])')


def split_text_to_chunks(text, max_length=200):
    """按标点断句，将文本拆为不超过 max_length 的片段列表。

    优先在句号等强分隔符断句，不够时在逗号等弱分隔符断，
    最坏情况按 max_length 硬切。
    """
    text = text.strip()
    if not text:
        return []

    # 第一轮：按强分隔符拆分
    raw_sentences = SENTENCE_DELIMITERS.split(text)
    raw_sentences = [s for s in raw_sentences if s.strip()]

    chunks = []
    buffer = ""

    for sentence in raw_sentences:
        sentence = sentence.strip()
        if not sentence:
            continue

        # 如果当前句子本身就超长，二次拆分
        if len(sentence) > max_length:
            # 先把 buffer 里攒的内容提交
            if buffer:
                chunks.append(buffer)
                buffer = ""
            # 按弱分隔符拆
            sub_parts = CLAUSE_DELIMITERS.split(sentence)
            sub_buf = ""
            for part in sub_parts:
                part = part.strip()
                if not part:
                    continue
                if len(sub_buf) + len(part) <= max_length:
                    sub_buf += part
                else:
                    if sub_buf:
                        chunks.append(sub_buf)
                    # 如果单个 part 仍然超长，硬切
                    while len(part) > max_length:
                        chunks.append(part[:max_length])
                        part = part[max_length:]
                    sub_buf = part
            if sub_buf:
                buffer = sub_buf
            continue

        # 正常情况：尝试把句子追加到 buffer
        if len(buffer) + len(sentence) <= max_length:
            buffer += sentence
        else:
            if buffer:
                chunks.append(buffer)
            buffer = sentence

    if buffer:
        chunks.append(buffer)

    return chunks


class Application(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("文本转语音转换器")
        self.geometry("1200x700")
        self.minsize(700, 400)
        self.resizable(True, True)

        self.file_path = tk.StringVar()
        self.output_dir = tk.StringVar()
        self.rate_var = tk.DoubleVar(value=50.00)
        self.volume_var = tk.DoubleVar(value=50.00)
        self.display_rate_var = tk.StringVar()
        self.display_volume_var = tk.StringVar()
        self.update_display_vars()
        self.rate_var.trace_add('write', self.update_display_vars)
        self.volume_var.trace_add('write', self.update_display_vars)

        # 语音列表（后台异步加载）
        self.voices = []
        self.voice_var = tk.StringVar(value=DEFAULT_VOICE)

        # 断句设置
        self.chunk_size_var = tk.IntVar(value=200)

        # 流式播放状态
        self._playback_stop = threading.Event()
        self._playback_thread = None
        self._temp_dir = None
        self._is_playing = False

        # 初始化 pygame mixer
        pygame.mixer.init()

        self.style = ttk.Style()
        self.style.theme_use('clam')
        self.style.configure('TButton', font=('微软雅黑', 10), padding=5)
        self.style.configure('TEntry', font=('微软雅黑', 10))
        self.style.configure('TLabel', font=('微软雅黑', 10))
        self.style.configure('TFrame', background='#f0f0f0')
        self.style.configure('TLabelframe', font=('微软雅黑', 10, 'bold'))
        self.style.configure('Accent.TButton', foreground='white', background='#4CAF50',
                             font=('微软雅黑', 11, 'bold'))
        self.style.map('Accent.TButton', background=[('active', '#45a049'), ('!active', '#4CAF50')])
        self.style.configure('Stop.TButton', foreground='white', background='#f44336',
                             font=('微软雅黑', 11, 'bold'))
        self.style.map('Stop.TButton', background=[('active', '#d32f2f'), ('!active', '#f44336')])

        self.init_ui()
        self.load_voices_async()

        # 关闭窗口时清理
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    def _on_close(self):
        """窗口关闭时停止播放并清理"""
        self.stop_playback()
        pygame.mixer.quit()
        self.destroy()

    def init_ui(self):
        main_frame = ttk.Frame(self, padding=20)
        main_frame.pack(fill=tk.BOTH, expand=True)

        title_frame = tk.Frame(main_frame, bd=0, highlightthickness=0, bg=self.cget('bg'))
        title_frame.pack(fill=tk.X, pady=(0, 20))
        title_label = tk.Label(
            title_frame,
            text="文本转语音转换器",
            font=('微软雅黑', 16, 'bold'),
            bg=self.cget('bg')
        )
        title_label.pack(side=tk.LEFT)

        help_btn = ttk.Button(title_frame, text="帮助", command=self.show_help, width=8)
        help_btn.pack(side=tk.RIGHT, padx=(10, 0))

        pw = ttk.PanedWindow(main_frame, orient=tk.HORIZONTAL)
        pw.pack(fill=tk.BOTH, expand=True)

        left_panel = ttk.Labelframe(pw, text="输入和设置", padding=15)
        pw.add(left_panel, weight=2)

        right_panel = ttk.Labelframe(pw, text="文本预览和转换", padding=15)
        pw.add(right_panel, weight=3)

        self.create_left_panel(left_panel)
        self.create_right_panel(right_panel)

        self.status_var = tk.StringVar()
        self.status_var.set("正在加载语音列表...")
        status_frame = ttk.Frame(main_frame)
        status_frame.pack(fill=tk.X, pady=(10, 0))
        status_label = ttk.Label(status_frame, textvariable=self.status_var, foreground='#666')
        status_label.pack(side=tk.LEFT)
        version_label = ttk.Label(status_frame, text="edge-tts · EdgeTTSPlayer", foreground='#999')
        version_label.pack(side=tk.RIGHT)

    def update_display_vars(self, *args):
        self.display_rate_var.set(f"{self.rate_var.get():.2f}")
        self.display_volume_var.set(f"{self.volume_var.get():.2f}")

    def create_left_panel(self, parent):
        file_frame = ttk.LabelFrame(parent, text="文本文件", padding=10)
        file_frame.pack(fill=tk.X, pady=(0, 10))

        entry_frame = ttk.Frame(file_frame)
        entry_frame.pack(fill=tk.X, expand=True)
        self.txt_path = ttk.Entry(entry_frame, textvariable=self.file_path, width=30)
        self.txt_path.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
        btn_sel = ttk.Button(entry_frame, text="选择文件", command=self.select_file, width=10)
        btn_sel.pack(side=tk.RIGHT)

        output_frame = ttk.LabelFrame(parent, text="输出设置", padding=10)
        output_frame.pack(fill=tk.X, pady=(0, 10))

        dir_frame = ttk.Frame(output_frame)
        dir_frame.pack(fill=tk.X, expand=True)
        self.txt_output_dir = ttk.Entry(dir_frame, textvariable=self.output_dir, width=30)
        self.txt_output_dir.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
        btn_dir = ttk.Button(dir_frame, text="选择目录", command=self.select_output_dir, width=10)
        btn_dir.pack(side=tk.RIGHT)

        voice_frame = ttk.LabelFrame(parent, text="语音设置", padding=10)
        voice_frame.pack(fill=tk.X, pady=(0, 10))

        # 语音选择下拉框
        voice_sel_frame = ttk.Frame(voice_frame)
        voice_sel_frame.pack(fill=tk.X, pady=(0, 5))
        ttk.Label(voice_sel_frame, text="语音:").pack(side=tk.LEFT)
        self.voice_combo = ttk.Combobox(voice_sel_frame, textvariable=self.voice_var,
                                        state='readonly', width=35)
        self.voice_combo.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(5, 0))
        self.voice_combo.set("加载中...")

        rate_frame = ttk.Frame(voice_frame)
        rate_frame.pack(fill=tk.X, pady=(5, 0))
        ttk.Label(rate_frame, text="语速:").pack(side=tk.LEFT)
        self.rate_scale = ttk.Scale(rate_frame, from_=1, to=100, variable=self.rate_var, orient=tk.HORIZONTAL)
        self.rate_scale.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(5, 0))
        self.rate_label = ttk.Label(rate_frame, textvariable=self.display_rate_var, width=6)
        self.rate_label.pack(side=tk.LEFT)

        volume_frame = ttk.Frame(voice_frame)
        volume_frame.pack(fill=tk.X, pady=(5, 0))
        ttk.Label(volume_frame, text="音量:").pack(side=tk.LEFT)
        self.volume_scale = ttk.Scale(volume_frame, from_=1, to=100, variable=self.volume_var, orient=tk.HORIZONTAL)
        self.volume_scale.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(5, 0))
        self.volume_label = ttk.Label(volume_frame, textvariable=self.display_volume_var, width=6)
        self.volume_label.pack(side=tk.LEFT)

        # 断句设置
        chunk_frame = ttk.LabelFrame(parent, text="断句设置", padding=10)
        chunk_frame.pack(fill=tk.X, pady=(0, 10))

        chunk_inner = ttk.Frame(chunk_frame)
        chunk_inner.pack(fill=tk.X)
        ttk.Label(chunk_inner, text="每片最大字数:").pack(side=tk.LEFT)
        self.chunk_spinbox = ttk.Spinbox(chunk_inner, from_=50, to=1000, increment=50,
                                         textvariable=self.chunk_size_var, width=8)
        self.chunk_spinbox.pack(side=tk.LEFT, padx=(5, 0))
        ttk.Label(chunk_inner, text="字", foreground='#999').pack(side=tk.LEFT, padx=(3, 0))

    def create_right_panel(self, parent):
        preview_frame = ttk.LabelFrame(parent, text="文本预览", padding=10)
        preview_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        scrollbar = ttk.Scrollbar(preview_frame, orient=tk.VERTICAL)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.text_preview = tk.Text(preview_frame, height=8, font=('微软雅黑', 10), wrap=tk.WORD,
                                    yscrollcommand=scrollbar.set)
        self.text_preview.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.text_preview.bind('<<Modified>>', self.on_text_modified)

        scrollbar.configure(command=self.text_preview.yview)

        convert_frame = ttk.Frame(parent)
        convert_frame.pack(fill=tk.X, pady=(10, 0))

        # 播放 / 停止按钮
        play_frame = ttk.Frame(convert_frame)
        play_frame.pack(fill=tk.X, pady=(0, 5))

        self.btn_play = ttk.Button(
            play_frame,
            text="▶ 播放",
            command=self.start_playback,
            style='Accent.TButton'
        )
        self.btn_play.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))

        self.btn_stop = ttk.Button(
            play_frame,
            text="■ 停止",
            command=self.stop_playback,
            style='Stop.TButton',
            state='disabled'
        )
        self.btn_stop.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(5, 0))

        # 播放状态
        self.play_status_var = tk.StringVar()
        self.play_status_label = ttk.Label(convert_frame, textvariable=self.play_status_var,
                                           foreground='#2196F3', font=('微软雅黑', 9))
        self.play_status_label.pack(fill=tk.X, pady=(0, 5))

        # 转换按钮
        self.btn_convert = ttk.Button(
            convert_frame,
            text="转换为MP3",
            command=self.convert_to_mp3,
            style='Accent.TButton'
        )
        self.btn_convert.pack(fill=tk.X, pady=(5, 0))

        btn_batch = ttk.Button(
            convert_frame,
            text="批量转换",
            command=self.batch_convert,
            style='Accent.TButton'
        )
        btn_batch.pack(fill=tk.X, pady=(5, 0))

        btn_open_dir = ttk.Button(
            convert_frame,
            text="打开输出目录",
            command=self.open_output_dir
        )
        btn_open_dir.pack(fill=tk.X, pady=(5, 0))

        self.progress = ttk.Progressbar(convert_frame, mode='indeterminate', length=200)

    # ====================== 语音加载 ======================

    def load_voices_async(self):
        """后台异步加载 edge-tts 语音列表"""
        def _load():
            try:
                loop = asyncio.new_event_loop()
                voices = loop.run_until_complete(edge_tts.list_voices())
                loop.close()

                # 只保留中文语音，按 ShortName 排序
                zh_voices = [v for v in voices if v["Locale"].startswith("zh-")]
                zh_voices.sort(key=lambda v: v["ShortName"])

                self.voices = zh_voices

                display_names = []
                for v in zh_voices:
                    gender = "女" if v["Gender"] == "Female" else "男"
                    locale = v["Locale"]
                    display_names.append(f"{v['ShortName']}  ({gender}, {locale})")

                self.after(0, lambda: self._update_voice_ui(display_names))
            except Exception as e:
                self.after(0, lambda: self.status_var.set(f"加载语音列表失败: {str(e)}"))

        threading.Thread(target=_load, daemon=True).start()

    def _update_voice_ui(self, display_names):
        """在主线程中更新语音下拉框"""
        self.voice_combo['values'] = display_names
        for i, v in enumerate(self.voices):
            if v["ShortName"] == DEFAULT_VOICE:
                self.voice_combo.current(i)
                break
        else:
            if display_names:
                self.voice_combo.current(0)
        self.status_var.set(f"准备就绪 — 已加载 {len(display_names)} 个中文语音")

    def get_selected_voice(self):
        """获取当前选中的语音 ShortName"""
        idx = self.voice_combo.current()
        if 0 <= idx < len(self.voices):
            return self.voices[idx]["ShortName"]
        return DEFAULT_VOICE

    # ====================== 参数映射 ======================

    def get_rate_string(self):
        """将滑块值 (1-100) 映射为 edge-tts 的 rate 参数字符串"""
        rate = self.rate_var.get()
        if rate <= 50:
            percent = int((rate - 50) / 50 * 50)
        else:
            percent = int((rate - 50) / 50 * 100)
        return f"{percent:+d}%"

    def get_volume_string(self):
        """将滑块值 (1-100) 映射为 edge-tts 的 volume 参数字符串"""
        volume = self.volume_var.get()
        percent = int((volume - 50) / 50 * 50)
        return f"{percent:+d}%"

    # ====================== 文件操作 ======================

    def select_file(self):
        current_file = self.file_path.get()
        if current_file and os.path.exists(current_file):
            initial_dir = os.path.dirname(current_file)
        else:
            initial_dir = str(pathlib.Path.home())

        txt_file = filedialog.askopenfilename(
            initialdir=initial_dir,
            title="选择文本文件",
            filetypes=(('文本文件', '*.txt'), ('所有文件', '*.*'))
        )
        if txt_file:
            self.file_path.set(txt_file)
            file_dir = os.path.dirname(txt_file)
            self.output_dir.set(file_dir)
            self.status_var.set(f"已选择文件: {pathlib.Path(txt_file).name}")
            try:
                with open(txt_file, 'r', encoding='utf8') as f:
                    content = f.read()
                self.text_preview.delete(1.0, tk.END)
                self.text_preview.insert(tk.END, content)
                self.text_preview.edit_modified(False)
            except Exception as e:
                messagebox.showwarning("警告", f"无法读取文件内容: {str(e)}")

    def select_output_dir(self):
        current_file = self.file_path.get()
        if current_file and os.path.exists(current_file):
            initial_dir = os.path.dirname(current_file)
        elif self.output_dir.get():
            initial_dir = self.output_dir.get()
        else:
            initial_dir = str(pathlib.Path.home())

        directory = filedialog.askdirectory(
            initialdir=initial_dir,
            title="选择输出目录"
        )
        if directory:
            self.output_dir.set(directory)
            self.status_var.set(f"输出目录设置为: {directory}")

    def on_text_modified(self, event):
        if self.text_preview.edit_modified():
            current_file = self.file_path.get()
            if current_file and os.path.exists(current_file):
                try:
                    content = self.text_preview.get(1.0, tk.END)
                    with open(current_file, 'w', encoding='utf8') as f:
                        f.write(content)
                    self.status_var.set(f"已自动保存修改到文件: {os.path.basename(current_file)}")
                except Exception as e:
                    self.status_var.set(f"自动保存失败: {str(e)}")
            self.text_preview.edit_modified(False)

    # ====================== 流式播放 ======================

    def start_playback(self):
        """开始流式播放：断句 → 双缓冲生成+播放"""
        text = self.text_preview.get(1.0, tk.END).strip()
        if not text:
            messagebox.showwarning("警告", "没有可播放的文本内容!")
            return

        # 如果正在播放，先停止
        if self._is_playing:
            self.stop_playback()

        max_len = self.chunk_size_var.get()
        chunks = split_text_to_chunks(text, max_len)
        if not chunks:
            messagebox.showwarning("警告", "文本断句后为空!")
            return

        # 创建临时目录
        self._temp_dir = tempfile.mkdtemp(prefix="tts_stream_")
        self._playback_stop.clear()
        self._is_playing = True

        # 更新 UI 状态
        self.btn_play.state(['disabled'])
        self.btn_stop.state(['!disabled'])
        self.btn_convert.state(['disabled'])
        self.play_status_var.set(f"准备播放... 共 {len(chunks)} 个片段")
        self.status_var.set(f"流式播放中 — 共 {len(chunks)} 个片段")

        voice = self.get_selected_voice()
        rate = self.get_rate_string()
        volume = self.get_volume_string()

        # 启动后台播放线程
        self._playback_thread = threading.Thread(
            target=self._playback_worker,
            args=(chunks, voice, rate, volume),
            daemon=True
        )
        self._playback_thread.start()

    def stop_playback(self):
        """停止播放并清理"""
        self._playback_stop.set()

        # 停止当前正在播放的音频
        try:
            pygame.mixer.music.stop()
        except Exception:
            pass

        # 等待线程结束
        if self._playback_thread and self._playback_thread.is_alive():
            self._playback_thread.join(timeout=3)

        self._cleanup_temp_dir()
        self._is_playing = False

        # 恢复 UI
        self.after(0, self._reset_play_ui)

    def _reset_play_ui(self):
        """在主线程恢复播放按钮状态"""
        self.btn_play.state(['!disabled'])
        self.btn_stop.state(['disabled'])
        self.btn_convert.state(['!disabled'])
        self.play_status_var.set("")

    def _playback_worker(self, chunks, voice, rate, volume):
        """后台线程：双缓冲生成+播放碎片"""
        loop = asyncio.new_event_loop()
        total = len(chunks)
        next_path = None

        try:
            # 预生成第一个片段
            if self._playback_stop.is_set():
                return
            first_path = os.path.join(self._temp_dir, "chunk_0.mp3")
            self.after(0, lambda: self.play_status_var.set(f"正在生成片段 1/{total}..."))
            loop.run_until_complete(self._generate_chunk_audio(chunks[0], first_path, voice, rate, volume))

            for i in range(total):
                if self._playback_stop.is_set():
                    return

                current_path = first_path if i == 0 else next_path

                # 异步预生成下一个片段（如果还有的话）
                next_path = None
                gen_thread = None
                if i + 1 < total:
                    next_path = os.path.join(self._temp_dir, f"chunk_{i + 1}.mp3")
                    gen_done = threading.Event()
                    gen_error = [None]

                    def _gen_next(idx=i + 1, path=next_path):
                        try:
                            gen_loop = asyncio.new_event_loop()
                            gen_loop.run_until_complete(
                                self._generate_chunk_audio(chunks[idx], path, voice, rate, volume)
                            )
                            gen_loop.close()
                        except Exception as e:
                            gen_error[0] = e
                        finally:
                            gen_done.set()

                    gen_thread = threading.Thread(target=_gen_next, daemon=True)
                    gen_thread.start()

                # 播放当前片段
                self.after(0, lambda idx=i: self.play_status_var.set(
                    f"▶ 正在播放 {idx + 1}/{total} 片段..."
                ))

                try:
                    pygame.mixer.music.load(current_path)
                    pygame.mixer.music.play()

                    # 等待播放完成（轮询检查停止标志）
                    while pygame.mixer.music.get_busy():
                        if self._playback_stop.is_set():
                            pygame.mixer.music.stop()
                            return
                        pygame.time.wait(100)
                except Exception as e:
                    self.after(0, lambda err=str(e): self.status_var.set(f"播放出错: {err}"))
                    return

                # 删除已播放完的临时文件
                try:
                    pygame.mixer.music.unload()
                    os.remove(current_path)
                except Exception:
                    pass

                # 等待下一个片段生成完成
                if gen_thread:
                    gen_done.wait()
                    if gen_error[0]:
                        self.after(0, lambda err=str(gen_error[0]): self.status_var.set(
                            f"生成片段出错: {err}"
                        ))
                        return

            # 播放完毕
            self.after(0, lambda: self.status_var.set("播放完毕"))

        except Exception as e:
            self.after(0, lambda err=str(e): self.status_var.set(f"流式播放出错: {err}"))
        finally:
            loop.close()
            self._cleanup_temp_dir()
            self._is_playing = False
            self.after(0, self._reset_play_ui)

    async def _generate_chunk_audio(self, text, output_path, voice, rate, volume):
        """调用 edge-tts 生成单个碎片音频"""
        communicate = edge_tts.Communicate(text, voice, rate=rate, volume=volume)
        await communicate.save(output_path)

    def _cleanup_temp_dir(self):
        """清理临时目录"""
        if self._temp_dir and os.path.isdir(self._temp_dir):
            try:
                shutil.rmtree(self._temp_dir, ignore_errors=True)
            except Exception:
                pass
            self._temp_dir = None

    # ====================== 转换逻辑 (edge-tts) ======================

    def convert_to_mp3(self):
        text = self.text_preview.get(1.0, tk.END).strip()
        if not text:
            messagebox.showwarning("警告", "没有可转换的文本内容!")
            return

        def convert_thread():
            try:
                self.btn_convert.state(['disabled'])
                self.progress.pack(fill=tk.X, pady=(10, 0))
                self.progress.start()
                self.status_var.set("正在转换...")
                self.update()

                voice = self.get_selected_voice()
                rate = self.get_rate_string()
                volume = self.get_volume_string()

                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                output_name = f"TTS_{timestamp}.mp3"
                output_dir = self.output_dir.get() or os.path.dirname(self.file_path.get()) or str(pathlib.Path.home())
                output_path = pathlib.Path(output_dir) / output_name

                loop = asyncio.new_event_loop()
                communicate = edge_tts.Communicate(text, voice, rate=rate, volume=volume)
                loop.run_until_complete(communicate.save(str(output_path)))
                loop.close()

                self.progress.stop()
                self.progress.pack_forget()
                self.status_var.set(f"转换成功! 文件已保存为: {output_path.name}")
                messagebox.showinfo("成功", f"文件转换成功!\n已保存为: {output_path.name}")
            except Exception as e:
                self.progress.stop()
                self.progress.pack_forget()
                self.status_var.set("转换失败")
                messagebox.showerror("错误", f"转换过程中出错:\n{str(e)}")
            finally:
                self.btn_convert.state(['!disabled'])

        threading.Thread(target=convert_thread, daemon=True).start()

    def batch_convert(self):
        files = filedialog.askopenfilenames(
            initialdir=os.path.dirname(self.file_path.get()) if self.file_path.get() else str(pathlib.Path.home()),
            title="选择多个文本文件",
            filetypes=(('文本文件', '*.txt'), ('所有文件', '*.*'))
        )
        if not files:
            return

        def batch_thread():
            try:
                self.btn_convert.state(['disabled'])
                self.progress.pack(fill=tk.X, pady=(10, 0))
                self.progress.start()
                self.status_var.set("正在批量转换...")
                self.update()

                success_count = 0
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                voice = self.get_selected_voice()
                rate = self.get_rate_string()
                volume = self.get_volume_string()

                loop = asyncio.new_event_loop()

                for i, file_path in enumerate(files, 1):
                    if not file_path:
                        continue
                    try:
                        with open(file_path, 'r', encoding='utf8') as f:
                            text = f.read().strip()
                        if not text:
                            continue

                        output_name = f"TTS_batch_{timestamp}_{i}.mp3"
                        output_dir = self.output_dir.get() or os.path.dirname(file_path) or str(pathlib.Path.home())
                        output_path = pathlib.Path(output_dir) / output_name

                        communicate = edge_tts.Communicate(text, voice, rate=rate, volume=volume)
                        loop.run_until_complete(communicate.save(str(output_path)))

                        success_count += 1
                        self.status_var.set(f"正在批量转换... 已完成 {i}/{len(files)}")
                        self.update()
                    except Exception as e:
                        self.status_var.set(f"转换 {os.path.basename(file_path)} 失败: {str(e)}")
                        continue

                loop.close()

                self.progress.stop()
                self.progress.pack_forget()
                self.status_var.set(f"批量转换完成! 成功转换 {success_count}/{len(files)} 个文件")
                messagebox.showinfo("完成", f"批量转换完成!\n成功转换 {success_count}/{len(files)} 个文件")
            except Exception as e:
                self.progress.stop()
                self.progress.pack_forget()
                self.status_var.set("批量转换失败")
                messagebox.showerror("错误", f"批量转换过程中出错:\n{str(e)}")
            finally:
                self.btn_convert.state(['!disabled'])

        threading.Thread(target=batch_thread, daemon=True).start()

    # ====================== 其他功能 ======================

    def open_output_dir(self):
        dir_path = self.output_dir.get()
        if not dir_path and self.file_path.get():
            dir_path = os.path.dirname(self.file_path.get())
        if not dir_path or not os.path.isdir(dir_path):
            messagebox.showwarning("警告", "无效的输出目录!")
            return
        try:
            if platform.system() == "Windows":
                os.startfile(dir_path)
            elif platform.system() == "Darwin":
                webbrowser.open(f"file://{dir_path}")
            else:
                webbrowser.open(dir_path)
        except Exception as e:
            messagebox.showerror("错误", f"无法打开目录:\n{str(e)}")

    def show_help(self):
        help_text = """文本转语音转换器使用说明（edge-tts 版）

1. 基本使用:
   - 点击"选择文件"按钮选择要转换的文本文件
   - 在右侧预览区域可以查看和编辑内容（自动保存）
   - 点击"转换为MP3"保存为文件

2. 流式播放（推荐）:
   - 点击 ▶ 播放，文本自动断句并连续播放
   - 播放时显示当前片段进度
   - 点击 ■ 停止即可中断，临时文件自动清理
   - "每片最大字数"控制每个片段的长度

3. 语音设置:
   - 语音: 从下拉框选择中文语音
   - 语速: 滑块中间为正常语速
   - 音量: 滑块中间为正常音量

4. 批量转换:
   - 点击"批量转换"选择多个文本文件

5. 注意事项:
   - 需要网络连接（Microsoft Edge 在线 TTS）
"""
        messagebox.showinfo("帮助", help_text)


if __name__ == "__main__":
    app = Application()
    app.mainloop()
