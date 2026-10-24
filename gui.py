# -*- coding: utf-8 -*-
"""
论文格式一键排版 · 图形界面

用法：
    python gui.py

打包成 exe 后可直接双击运行。

规则来源（优先级从高到低）：
    1. exe 同目录的「我的格式.txt」—— 用户在界面上保存过的自定义格式
    2. exe 同目录的 rules.yaml      —— 手工编辑的完整规则
    3. 打包进 exe 的默认 rules.yaml
"""

import os
import sys
import threading
import traceback
from pathlib import Path

import tkinter as tk
from tkinter import filedialog, messagebox

try:
    import ctypes
    ctypes.windll.shcore.SetProcessDpiAwareness(1)
except Exception:
    pass

import custom_format as cf
import format_docx as fd

APP_TITLE = "论文格式一键排版"
BG = "#F6F9FA"
INK = "#0F2A36"
MUTED = "#6B8592"
ACCENT = "#0E7490"
ERR = "#A32D2D"
OKC = "#0F6E56"
MY_FILE = "我的格式.txt"


def base_dir():
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent
    return Path(__file__).parent


def _default_rules(base):
    """自定义格式要盖在默认规则上；默认规则优先取同目录 rules.yaml，其次取打包内置"""
    for p in (base / "rules.yaml",
              Path(getattr(sys, "_MEIPASS", base)) / "rules.yaml"):
        if p.exists():
            try:
                return fd.load_rules(p)
            except Exception:
                continue
    return None


def load_rules():
    """按优先级加载：我的格式.txt → rules.yaml → 内置默认"""
    base = base_dir()

    mine = base / MY_FILE
    if mine.exists():
        try:
            text = mine.read_text(encoding="utf-8")
            override, errs, _ = cf.parse(text)
            default = _default_rules(base)
            if not errs and default:
                return cf.merge(default, override), text, MY_FILE
        except Exception:
            pass

    local = base / "rules.yaml"
    if local.exists():
        try:
            return fd.load_rules(local), "", "rules.yaml"
        except Exception:
            pass

    bundled = Path(getattr(sys, "_MEIPASS", base)) / "rules.yaml"
    if bundled.exists():
        try:
            return fd.load_rules(bundled), "", "内置默认"
        except Exception:
            pass

    return None, "", "未找到规则文件"


class App:
    def __init__(self, root):
        self.root = root
        self.src_path = None
        self.running = False

        self.base_rules, saved_text, self.source_label = load_rules()
        self.rules = self.base_rules
        self.override = {}

        root.title(APP_TITLE)
        root.configure(bg=BG)
        root.resizable(False, True)
        root.option_add("*Font", ("Microsoft YaHei UI", 10))

        self._build()
        self._refresh_rule_label()
        if saved_text:
            self.txt.insert("1.0", saved_text)
            self._apply_custom(quiet=True)
        self._center(600, 660)

    # ---------------- 界面 ----------------
    def _build(self):
        pad = 26

        head = tk.Frame(self.root, bg=BG)
        head.pack(fill="x", padx=pad, pady=(22, 0))
        tk.Label(head, text=APP_TITLE, bg=BG, fg=INK,
                 font=("Microsoft YaHei UI", 16, "bold")).pack(anchor="w")
        tk.Label(head, text="把 Word 论文的字体、行距、缩进一次改对",
                 bg=BG, fg=MUTED, font=("Microsoft YaHei UI", 9)).pack(anchor="w", pady=(3, 0))

        body = tk.Frame(self.root, bg=BG)
        body.pack(fill="both", expand=True, padx=pad, pady=(16, 0))

        self.btn_pick = tk.Button(
            body, text="选择论文文件（.docx）", command=self.pick,
            bg="#FFFFFF", fg=INK, relief="solid", bd=1,
            font=("Microsoft YaHei UI", 10), cursor="hand2",
            activebackground="#EAF2F5", height=2)
        self.btn_pick.pack(fill="x")

        self.file_label = tk.Label(
            body, text="还没有选择文件", bg=BG, fg=MUTED,
            font=("Microsoft YaHei UI", 9), anchor="w", wraplength=520, justify="left")
        self.file_label.pack(fill="x", pady=(8, 0))

        self.btn_run = tk.Button(
            body, text="开始排版", command=self.run,
            bg=ACCENT, fg="#FFFFFF", relief="flat", bd=0,
            font=("Microsoft YaHei UI", 11, "bold"), cursor="hand2",
            activebackground="#0B5C73", activeforeground="#FFFFFF",
            state="disabled", height=2)
        self.btn_run.pack(fill="x", pady=(14, 0))

        self.status = tk.Label(
            body, text="", bg=BG, fg=MUTED,
            font=("Microsoft YaHei UI", 9), anchor="w", wraplength=520, justify="left")
        self.status.pack(fill="x", pady=(8, 0))

        tk.Frame(body, bg="#DCE6EA", height=1).pack(fill="x", pady=(14, 0))

        # ---------- 自定义格式 ----------
        tk.Label(body, text="自定义格式（可选，一行一条）", bg=BG, fg=INK,
                 font=("Microsoft YaHei UI", 10, "bold")).pack(anchor="w", pady=(12, 2))
        tk.Label(body, text="按学校文件的说法写就行：冒号前是要改什么，后面是改成什么样。",
                 bg=BG, fg=MUTED, font=("Microsoft YaHei UI", 9)).pack(anchor="w")

        box = tk.Frame(body, bg="#FFFFFF", highlightbackground="#D2DFE5",
                       highlightthickness=1)
        box.pack(fill="x", pady=(6, 0))
        self.txt = tk.Text(box, height=4, wrap="word", bd=0, bg="#FFFFFF", fg=INK,
                           font=("Microsoft YaHei UI", 9), padx=8, pady=6,
                           insertbackground=INK)
        self.txt.pack(fill="x")

        bar = tk.Frame(body, bg=BG)
        bar.pack(fill="x", pady=(6, 0))

        self.btn_apply = tk.Button(
            bar, text="应用", command=self._apply_custom,
            bg="#FFFFFF", fg=INK, relief="solid", bd=1,
            font=("Microsoft YaHei UI", 9), cursor="hand2", width=6)
        self.btn_apply.pack(side="left")

        self.btn_sample = tk.Button(
            bar, text="插入示例", command=self._insert_sample,
            bg=BG, fg=ACCENT, relief="flat", bd=0,
            font=("Microsoft YaHei UI", 9), cursor="hand2")
        self.btn_sample.pack(side="left", padx=(8, 0))

        self.btn_save = tk.Button(
            bar, text="保存为默认", command=self._save,
            bg=BG, fg=ACCENT, relief="flat", bd=0,
            font=("Microsoft YaHei UI", 9), cursor="hand2")
        self.btn_save.pack(side="left", padx=(8, 0))

        self.btn_clear = tk.Button(
            bar, text="清空", command=self._clear,
            bg=BG, fg=MUTED, relief="flat", bd=0,
            font=("Microsoft YaHei UI", 9), cursor="hand2")
        self.btn_clear.pack(side="right")

        self.hint = tk.Label(
            body, text="", bg=BG, fg=MUTED, anchor="w",
            font=("Microsoft YaHei UI", 9), wraplength=540, justify="left")
        self.hint.pack(fill="x", pady=(6, 0))

        foot = tk.Frame(self.root, bg="#EDF3F5")
        foot.pack(fill="x", side="bottom")
        self.rule_label = tk.Label(
            foot, text="", bg="#EDF3F5", fg=MUTED, anchor="w",
            font=("Microsoft YaHei UI", 9), padx=pad, pady=9,
            wraplength=540, justify="left")
        self.rule_label.pack(fill="x")

    def _center(self, w, h):
        self.root.update_idletasks()
        x = (self.root.winfo_screenwidth() - w) // 2
        y = max(20, (self.root.winfo_screenheight() - h) // 12)
        self.root.geometry(f"{w}x{h}+{x}+{y}")

    # ---------------- 规则 ----------------
    def _refresh_rule_label(self):
        if self.override:
            self.rule_label.config(
                text=f"当前规则：自定义（{cf.summarize(self.override)}）", fg=OKC)
        else:
            self.rule_label.config(text=f"当前规则：{self.source_label}", fg=MUTED)

    def _set_hint(self, text, color=MUTED):
        self.hint.config(text=text, fg=color)

    def _apply_custom(self, quiet=False):
        text = self.txt.get("1.0", "end").strip()
        if not text:
            self.override = {}
            self.rules = self.base_rules
            self._refresh_rule_label()
            self._set_hint("没有填写自定义格式，使用默认规则。")
            return True

        override, errs, warns = cf.parse(text)
        if errs:
            more = f"（共 {len(errs)} 处问题）" if len(errs) > 1 else ""
            self._set_hint(errs[0] + more, ERR)
            return False

        if not override:
            self._set_hint("没认出任何格式项，请检查写法。", ERR)
            return False

        merged = cf.merge(self.base_rules, override) if self.base_rules else None
        if not merged:
            self._set_hint("默认规则没加载成功，无法应用自定义格式。", ERR)
            return False

        self.override = override
        self.rules = merged
        self._refresh_rule_label()
        if not quiet:
            msg = f"已应用：{cf.summarize(override)}"
            if warns:
                msg += f"\n（{warns[0]}，不影响结果）" if len(warns) == 1 else \
                       f"\n（还有 {len(warns)} 处内容被跳过，不影响结果）"
            self._set_hint(msg, OKC)
        return True

    def _insert_sample(self):
        self.txt.delete("1.0", "end")
        self.txt.insert("1.0", cf.SAMPLE)
        self._set_hint("已填入示例格式。改哪一行就改哪一项，改完点「应用」。")

    def _clear(self):
        self.txt.delete("1.0", "end")
        self.override = {}
        self.rules = self.base_rules
        self._refresh_rule_label()
        self._set_hint("已清空，回到默认规则。")

    def _save(self):
        if not self.override:
            messagebox.showinfo(APP_TITLE, "先填好自定义格式并点「应用」，再保存。")
            return
        path = base_dir() / MY_FILE
        try:
            path.write_text(self.txt.get("1.0", "end").strip() + "\n", encoding="utf-8")
        except Exception as e:
            messagebox.showerror(
                APP_TITLE,
                f"保存失败：{e}\n\n可能是这个文件夹没有写入权限。"
                f"把程序换个位置（比如桌面）再试。")
            return
        self.source_label = MY_FILE
        self._refresh_rule_label()
        self._set_hint(f"已保存到：{path}\n下次打开会自动读取。", OKC)

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

        # 排版前把输入框内容再应用一次，免得用户改了忘点「应用」
        if self.txt.get("1.0", "end").strip():
            if not self._apply_custom():
                messagebox.showerror(APP_TITLE, "自定义格式有问题，先改好再排版。")
                return

        if not self.rules:
            messagebox.showerror(APP_TITLE, "规则没有加载成功，无法排版。")
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
        # 注意：except 里的 e 在块结束会被 Python 删掉，
        # 延迟执行的回调必须用默认参数把它绑住，否则回调报 NameError，
        # 界面会永远停在「正在处理」。
        try:
            stats = fd.process(self.src_path, out, self.rules)
        except Exception as exc:
            detail = traceback.format_exc()
            self.root.after(0, lambda e=exc, d=detail: self._fail(e, d))
            return
        self.root.after(0, lambda s=stats, o=out: self._done(s, o))

    def _done(self, stats, out):
        self.running = False
        self.btn_run.config(state="normal", text="开始排版")
        self.btn_pick.config(state="normal")

        lines = []
        nzr = stats.get("normalize")
        if nzr:
            lines.append("这份文档原来没用 Word 样式，已先重建结构：")
            if nzr.get("md_heading"):
                lines.append(f"· {nzr['md_heading']} 个 Markdown 标题转成 Word 标题")
            if nzr.get("doc_title"):
                lines.append("· 开头的文档大标题改用 Title 样式")
            if nzr.get("aligned_headings"):
                lines.append(f"· 补齐 {nzr['aligned_headings']} 个漏掉样式的章节标题")
            if nzr.get("inferred_headings"):
                lines.append(f"· 按字号推断出 {nzr['inferred_headings']} 个标题")
            if nzr.get("md_bold"):
                lines.append(f"· 清理 {nzr['md_bold']} 处加粗符号")
            if nzr.get("md_removed"):
                lines.append(f"· 删除 {nzr['md_removed']} 段 Markdown 残留")
            if nzr.get("ai_note"):
                lines.append(f"· 删除 {nzr['ai_note']} 处 AI 生成工具的署名")
            if nzr.get("empty_removed"):
                lines.append(f"· 删除 {nzr['empty_removed']} 个空段落")
            if nzr.get("stripped_runs"):
                lines.append(f"· 清除 {nzr['stripped_runs']} 处手动格式")
            lines.append("")

        lines.append(f"已应用样式：{'、'.join(stats['applied'])}")
        if stats["missing"]:
            lines.append(f"跳过（文档里没有）：{'、'.join(stats['missing'])}")
        lines.append("")
        lines.append(f"一级标题 {stats['h1']} ｜ 二级 {stats['h2']} ｜ 三级 {stats['h3']} ｜ 正文 {stats['body']} 段")
        for n in stats["notes"]:
            lines.append(f"· {n}")

        if stats["h1"] == 0:
            lines.append("")
            lines.append("注意：没能识别出标题。")
            lines.append("如果文档里确实有标题，在 Word 里选中它，")
            lines.append("按 Ctrl+Alt+1 套用「标题 1」，再跑一次。")

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
        self.status.config(text="出错了", fg=ERR)

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
    if len(sys.argv) > 1:
        fd.main()
        return

    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
