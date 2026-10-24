# -*- coding: utf-8 -*-
"""
论文格式一键排版 · 图形界面

用法：
    python gui.py

打包成 exe 后可直接双击运行。
"""

import os
import sys
import threading
import traceback
from pathlib import Path

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

try:
    import ctypes
    ctypes.windll.shcore.SetProcessDpiAwareness(1)
except Exception:
    pass

import format_docx as fd

APP_TITLE = "论文格式一键排版"
BG = "#F6F9FA"
INK = "#0F2A36"
MUTED = "#6B8592"
ACCENT = "#0E7490"


def base_dir():
    """exe 所在目录（打包后）或脚本所在目录"""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent
    return Path(__file__).parent


def find_rules():
    """优先用同目录的 rules.yaml，找不到就用打包进 exe 的默认规则"""
    local = base_dir() / "rules.yaml"
    if local.exists():
        return local, "rules.yaml（可编辑）"
    bundled = Path(getattr(sys, "_MEIPASS", base_dir())) / "rules.yaml"
    if bundled.exists():
        return bundled, "内置默认"
    return None, "未找到"


def describe(cfg):
    size = {12: "小四", 10.5: "五号", 14: "四号", 16: "三号"}.get(
        cfg.get("size"), f"{cfg.get('size')}pt")
    ls = cfg.get("line_spacing", 1.5)
    ls_txt = "单倍行距" if ls == 1 else f"{ls:g} 倍行距"
    ind = cfg.get("first_line_chars", 0)
    ind_txt = "无缩进" if not ind else f"缩进 {ind:g} 字符"
    return f"{cfg.get('font_cn', '宋体')} {size} · {ls_txt} · {ind_txt}"


class App:
    def __init__(self, root):
        self.root = root
        self.src_path = None
        self.rules_path, self.rules_label = find_rules()
        self.rules = None
        self.running = False

        root.title(APP_TITLE)
        root.configure(bg=BG)
        root.resizable(False, False)
        root.option_add("*Font", ("Microsoft YaHei UI", 10))

        self._build()
        self._load_rules()
        self._center(560, 400)

    # ---------------- 界面 ----------------
    def _build(self):
        pad = 28

        head = tk.Frame(self.root, bg=BG)
        head.pack(fill="x", padx=pad, pady=(24, 0))
        tk.Label(head, text=APP_TITLE, bg=BG, fg=INK,
                 font=("Microsoft YaHei UI", 16, "bold")).pack(anchor="w")
        tk.Label(head, text="把 Word 论文的字体、行距、缩进一次改对",
                 bg=BG, fg=MUTED, font=("Microsoft YaHei UI", 9)).pack(anchor="w", pady=(4, 0))

        body = tk.Frame(self.root, bg=BG)
        body.pack(fill="both", expand=True, padx=pad, pady=(20, 0))

        self.btn_pick = tk.Button(
            body, text="选择论文文件（.docx）", command=self.pick,
            bg="#FFFFFF", fg=INK, relief="solid", bd=1,
            font=("Microsoft YaHei UI", 10), cursor="hand2",
            activebackground="#EAF2F5", height=2)
        self.btn_pick.pack(fill="x")

        self.file_label = tk.Label(
            body, text="还没有选择文件", bg=BG, fg=MUTED,
            font=("Microsoft YaHei UI", 9), anchor="w", wraplength=480, justify="left")
        self.file_label.pack(fill="x", pady=(10, 0))

        self.btn_run = tk.Button(
            body, text="开始排版", command=self.run,
            bg=ACCENT, fg="#FFFFFF", relief="flat", bd=0,
            font=("Microsoft YaHei UI", 11, "bold"), cursor="hand2",
            activebackground="#0B5C73", activeforeground="#FFFFFF",
            state="disabled", height=2)
        self.btn_run.pack(fill="x", pady=(18, 0))

        self.status = tk.Label(
            body, text="", bg=BG, fg=MUTED,
            font=("Microsoft YaHei UI", 9), anchor="w", wraplength=480, justify="left")
        self.status.pack(fill="x", pady=(12, 0))

        foot = tk.Frame(self.root, bg="#EDF3F5")
        foot.pack(fill="x", side="bottom")
        self.rule_label = tk.Label(
            foot, text="", bg="#EDF3F5", fg=MUTED,
            font=("Microsoft YaHei UI", 9), anchor="w", padx=pad, pady=10)
        self.rule_label.pack(fill="x")

    def _center(self, w, h):
        self.root.update_idletasks()
        x = (self.root.winfo_screenwidth() - w) // 2
        y = (self.root.winfo_screenheight() - h) // 3
        self.root.geometry(f"{w}x{h}+{x}+{y}")

    def _load_rules(self):
        if not self.rules_path:
            self.rule_label.config(text="⚠ 找不到 rules.yaml，无法排版", fg="#A32D2D")
            return
        try:
            self.rules = fd.load_rules(self.rules_path)
            body = describe(self.rules.get("body", {}))
            self.rule_label.config(
                text=f"当前规则：{self.rules_label}　｜　正文 {body}")
        except Exception as e:
            self.rule_label.config(text=f"⚠ 规则文件读取失败：{e}", fg="#A32D2D")

    # ---------------- 交互 ----------------
    def pick(self):
        path = filedialog.askopenfilename(
            title="选择要排版的论文",
            filetypes=[("Word 文档", "*.docx"), ("所有文件", "*.*")])
        if not path:
            return
        if not path.lower().endswith(".docx"):
            messagebox.showwarning(
                APP_TITLE,
                "只支持 .docx 文件。\n\n如果是 .doc，请先在 Word 里「另存为」→ 选 .docx。")
            return
        self.src_path = Path(path)
        self.file_label.config(text=str(self.src_path), fg=INK)
        self.status.config(text="")
        self.btn_run.config(state="normal", bg=ACCENT)

    def run(self):
        if self.running or not self.src_path:
            return
        if not self.rules:
            messagebox.showerror(APP_TITLE, "规则文件没有加载成功，无法排版。")
            return

        out = self.src_path.with_name(self.src_path.stem + "-已排版.docx")
        if out.exists():
            if not messagebox.askyesno(
                    APP_TITLE,
                    f"「{out.name}」已存在，要覆盖吗？\n\n"
                    f"选「否」会另存为带序号的文件。"):
                i = 2
                while True:
                    cand = self.src_path.with_name(f"{self.src_path.stem}-已排版({i}).docx")
                    if not cand.exists():
                        out = cand
                        break
                    i += 1

        self.running = True
        self.btn_run.config(state="disabled", text="正在排版…")
        self.btn_pick.config(state="disabled")
        self.status.config(text="正在处理，请稍候…", fg=MUTED)

        threading.Thread(target=self._work, args=(out,), daemon=True).start()

    def _work(self, out):
        try:
            stats = fd.process(self.src_path, out, self.rules)
            self.root.after(0, lambda: self._done(stats, out))
        except Exception as e:
            detail = traceback.format_exc()
            self.root.after(0, lambda: self._fail(e, detail))

    def _done(self, stats, out):
        self.running = False
        self.btn_run.config(state="normal", text="开始排版")
        self.btn_pick.config(state="normal")

        lines = [f"已应用样式：{'、'.join(stats['applied'])}"]
        if stats["missing"]:
            lines.append(f"跳过（文档里没有）：{'、'.join(stats['missing'])}")
        lines.append("")
        lines.append(f"一级标题 {stats['h1']} 段 ｜ 二级标题 {stats['h2']} 段 ｜ 正文 {stats['body']} 段")
        for n in stats["notes"]:
            lines.append(f"· {n}")

        if stats["h1"] == 0:
            lines.append("")
            lines.append("⚠ 没检测到「标题 1」段落。")
            lines.append("说明你的标题是手动加粗做的，不是 Word 样式。")
            lines.append("在 Word 里选中标题，按 Ctrl+Alt+1 就能改过来。")

        self.status.config(text=f"完成：{out.name}", fg=ACCENT)

        if messagebox.askyesno(
                APP_TITLE,
                "\n".join(lines) + f"\n\n文件已保存到：\n{out}\n\n要打开所在文件夹吗？"):
            try:
                os.startfile(str(out.parent))
            except Exception:
                pass

    def _fail(self, exc, detail):
        self.running = False
        self.btn_run.config(state="normal", text="开始排版")
        self.btn_pick.config(state="normal")
        self.status.config(text="出错了", fg="#A32D2D")

        msg = str(exc)
        if "Permission" in msg or "being used" in msg:
            hint = "文件可能正在 Word 或 WPS 里打开着，先关掉再试。"
        elif isinstance(exc, FileNotFoundError):
            hint = "文件不存在，可能被移动或删除了。"
        elif isinstance(exc, ValueError):
            hint = msg
        else:
            hint = "详细信息：\n" + detail[-600:]
        messagebox.showerror(APP_TITLE, hint)


def main():
    # 带参数时走命令行模式（方便自动化测试与脚本调用）
    if len(sys.argv) > 1:
        fd.main()
        return

    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
