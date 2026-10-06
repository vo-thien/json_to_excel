# -*- coding: utf-8 -*-
"""Giao diện chuyển đổi JSON sang Excel.

Chạy:  python json_to_excel_gui.py     (hoặc double-click "Chay tool.bat")
"""

from __future__ import annotations

import os
import queue
import subprocess
import sys
import threading
import traceback

import tkinter as tk
import tkinter.font as tkfont
from tkinter import filedialog, messagebox, ttk

import json_excel_core as core

PREVIEW_ROWS = 200
PREVIEW_CELL_CHARS = 120
TICKED, UNTICKED = "☑", "☐"

#: Hệ số tỉ lệ hiển thị (1.0 = 100%, 1.75 = 175%). Đặt trong apply_scaling().
SCALE = 1.0


def px(n: float) -> int:
    """Đổi một kích thước thiết kế (tính theo pixel ở 100%) sang pixel thật."""
    return int(round(n * SCALE))


def dpi_scale() -> float:
    """Tỉ lệ hiển thị thật của Windows (1.0 = 100%, 1.5 = 150%...).

    Tk luôn đoán màn hình là 96 DPI, nên phải tự hỏi Windows.
    """
    if sys.platform != "win32":
        return 1.0
    try:
        import ctypes

        return ctypes.windll.user32.GetDpiForSystem() / 96.0
    except Exception:  # noqa: BLE001 - Windows cũ hơn 1607 hoặc không phải Windows
        try:
            import ctypes

            hdc = ctypes.windll.user32.GetDC(0)
            dpi = ctypes.windll.gdi32.GetDeviceCaps(hdc, 88)  # LOGPIXELSX
            ctypes.windll.user32.ReleaseDC(0, hdc)
            return dpi / 96.0
        except Exception:  # noqa: BLE001
            return 1.0


def apply_scaling(root: tk.Misc) -> float:
    """Chỉnh cỡ chữ và chiều cao hàng theo tỉ lệ hiển thị. Trả về hệ số tỉ lệ.

    Điểm quan trọng: chiều cao hàng của Treeview là số pixel CỐ ĐỊNH (mặc định
    20) và không tự co theo cỡ chữ. Trên màn hình 125%/150%/175%, chữ cao hơn
    20px nên các hàng bị đè lên nhau. Phải đo chiều cao thật của font rồi đặt lại.

    Hệ số trả về được dùng cho mọi kích thước pixel khác (bề rộng cột, lề...) —
    quên nhân theo nó thì tiêu đề cột sẽ bị cắt cụt.
    """
    global SCALE
    SCALE = dpi_scale()
    base_px = max(11, int(round(12 * SCALE)))

    for name in ("TkDefaultFont", "TkTextFont", "TkMenuFont"):
        try:
            tkfont.nametofont(name).configure(size=-base_px)
        except tk.TclError:
            pass
    try:
        tkfont.nametofont("TkHeadingFont").configure(size=-base_px, weight="bold")
    except tk.TclError:
        pass

    style = ttk.Style(root)
    line = tkfont.nametofont("TkDefaultFont").metrics("linespace")
    style.configure("Treeview", rowheight=line + px(6))
    style.configure("Treeview.Heading", padding=(px(4), px(4)))
    return SCALE


def preview_text(value) -> str:
    """Chuỗi hiển thị gọn cho ô xem trước."""
    if value is None:
        return ""
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if isinstance(value, (dict, list)):
        text = core.json_text(value)
    else:
        text = str(value)
    text = core.CONTROL_CHARS.sub("", text).replace("\n", " ⏎ ").replace("\r", "")
    if len(text) > PREVIEW_CELL_CHARS:
        text = text[: PREVIEW_CELL_CHARS - 1] + "…"
    return text


class App(tk.Tk):
    def __init__(self, initial_file: str | None = None) -> None:
        super().__init__()
        apply_scaling(self)
        self.title("Chuyển đổi JSON sang Excel")

        # Cửa sổ phải nằm gọn trong vùng làm việc: ở 175% thì 1000x760 thiết kế
        # nở thành 1750x1330 pixel thật, tràn ra ngoài màn hình và bị cắt đáy.
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        w = min(px(1000), sw - px(40))
        h = min(px(760), sh - px(100))
        self.geometry(f"{w}x{h}+{max(0, (sw - w) // 2)}+{max(0, (sh - h) // 3)}")
        self.minsize(min(px(820), w), min(px(560), h))

        self.queue: queue.Queue = queue.Queue()
        self.candidates: list[core.ArrayCandidate] = []
        self.checked: set[int] = set()
        self.file_path: str | None = None
        self.busy = False
        self.last_out_dir: str | None = None

        self.flatten_var = tk.BooleanVar(value=True)
        self.depth_var = tk.IntVar(value=20)
        self.path_var = tk.StringVar(value="")
        self.status_var = tk.StringVar(value="Chưa chọn file nào.")

        self._build_ui()
        self.after(80, self._poll_queue)

        # Cho phép mở sẵn một file: kéo file JSON thả lên "Chay tool.bat".
        if initial_file and os.path.isfile(initial_file):
            self.after(50, lambda: self.load_file(initial_file))

    # ------------------------------------------------------------- giao diện

    def _build_ui(self) -> None:
        pad = {"padx": px(8), "pady": px(4)}
        self.columnconfigure(0, weight=1)
        # Mục 3 giữ chiều cao cố định (đủ cho ~6 bảng), mục 4 giãn hết phần còn lại.
        self.rowconfigure(3, weight=1)

        # 1. Chọn file
        lf_file = ttk.LabelFrame(self, text="1. Chọn file JSON")
        lf_file.grid(row=0, column=0, sticky="ew", **pad)
        lf_file.columnconfigure(0, weight=1)
        ttk.Entry(lf_file, textvariable=self.path_var).grid(
            row=0, column=0, sticky="ew", padx=px(6), pady=px(6)
        )
        self.btn_open = ttk.Button(lf_file, text="Chọn file…", command=self.on_open)
        self.btn_open.grid(row=0, column=1, padx=(0, px(6)), pady=px(6))

        # 2. Tùy chọn
        lf_opt = ttk.LabelFrame(self, text="2. Tùy chọn")
        lf_opt.grid(row=1, column=0, sticky="ew", **pad)
        ttk.Checkbutton(
            lf_opt,
            text="Làm phẳng JSON lồng nhau (user.address.city)",
            variable=self.flatten_var,
            command=self.on_options_changed,
        ).grid(row=0, column=0, sticky="w", padx=px(6), pady=px(6))
        ttk.Label(lf_opt, text="Độ sâu tối đa:").grid(row=0, column=1, padx=(px(16), px(4)))
        ttk.Spinbox(
            lf_opt, from_=1, to=100, width=5, textvariable=self.depth_var, command=self.on_options_changed
        ).grid(row=0, column=2, padx=(0, px(6)))
        ttk.Label(
            lf_opt,
            text="(vượt độ sâu thì phần lồng giữ nguyên dạng chuỗi JSON)",
            foreground="#666",
        ).grid(row=0, column=3, sticky="w", padx=px(6))

        # 3. Bảng dữ liệu tìm thấy
        lf_arr = ttk.LabelFrame(self, text="3. Chọn bảng dữ liệu muốn xuất")
        lf_arr.grid(row=2, column=0, sticky="ew", **pad)
        lf_arr.columnconfigure(0, weight=1)

        bar = ttk.Frame(lf_arr)
        bar.grid(row=0, column=0, sticky="ew", padx=px(6), pady=(px(6), 0))
        ttk.Button(bar, text="Chọn tất cả", command=lambda: self._set_all(True)).pack(side="left")
        ttk.Button(bar, text="Bỏ chọn", command=lambda: self._set_all(False)).pack(side="left", padx=px(6))
        ttk.Label(bar, text="Nhấn vào dòng để bật/tắt.", foreground="#666").pack(side="left", padx=px(12))

        self.tree_arr = ttk.Treeview(
            lf_arr, columns=("tick", "path", "rows", "cols", "sheet"), show="headings", height=6
        )
        for col, text, width, anchor in (
            ("tick", "", 40, "center"),
            ("path", "Đường dẫn trong JSON", 300, "w"),
            ("rows", "Số dòng", 100, "e"),
            ("cols", "Số cột", 90, "e"),
            ("sheet", "Tên sheet", 160, "w"),
        ):
            self.tree_arr.heading(col, text=text)
            # Bề rộng cột cũng phải nhân theo tỉ lệ, nếu không tiêu đề sẽ bị cắt cụt.
            self.tree_arr.column(col, width=px(width), minwidth=px(width), anchor=anchor)
        self.tree_arr.grid(row=1, column=0, sticky="ew", padx=px(6), pady=px(6))
        self.tree_arr.bind("<Button-1>", self.on_array_click)
        self.tree_arr.bind("<<TreeviewSelect>>", lambda _e: self.refresh_preview())
        yscroll = ttk.Scrollbar(lf_arr, orient="vertical", command=self.tree_arr.yview)
        yscroll.grid(row=1, column=1, sticky="ns", pady=px(6), padx=(0, px(6)))
        self.tree_arr.configure(yscrollcommand=yscroll.set)

        # 4. Xem trước
        lf_prev = ttk.LabelFrame(self, text="4. Xem trước")
        lf_prev.grid(row=3, column=0, sticky="nsew", **pad)
        lf_prev.columnconfigure(0, weight=1)
        lf_prev.rowconfigure(1, weight=1)
        self.lbl_prev = ttk.Label(lf_prev, text="—", foreground="#666")
        self.lbl_prev.grid(row=0, column=0, sticky="w", padx=px(6), pady=(px(6), 0))
        self.tree_prev = ttk.Treeview(lf_prev, show="headings", height=8)
        self.tree_prev.grid(row=1, column=0, sticky="nsew", padx=px(6), pady=px(6))
        xscroll = ttk.Scrollbar(lf_prev, orient="horizontal", command=self.tree_prev.xview)
        xscroll.grid(row=2, column=0, sticky="ew", padx=px(6), pady=(0, px(6)))
        self.tree_prev.configure(xscrollcommand=xscroll.set)

        # 5. Xuất
        bar2 = ttk.Frame(self)
        bar2.grid(row=4, column=0, sticky="ew", **pad)
        bar2.columnconfigure(1, weight=1)
        self.btn_export = ttk.Button(bar2, text="Xuất Excel…", command=self.on_export, state="disabled")
        self.btn_export.grid(row=0, column=0, padx=(0, px(10)))
        self.progress = ttk.Progressbar(bar2, mode="determinate", maximum=100)
        self.progress.grid(row=0, column=1, sticky="ew")
        self.btn_folder = ttk.Button(bar2, text="Mở thư mục", command=self.on_open_folder, state="disabled")
        self.btn_folder.grid(row=0, column=2, padx=(px(10), 0))
        ttk.Label(self, textvariable=self.status_var, foreground="#333").grid(
            row=5, column=0, sticky="w", padx=px(16), pady=(0, px(10))
        )

    # ------------------------------------------------------------- tiện ích

    def _set_busy(self, busy: bool, status: str | None = None) -> None:
        self.busy = busy
        state = "disabled" if busy else "normal"
        self.btn_open.configure(state=state)
        self.btn_export.configure(state="disabled" if busy or not self.checked else "normal")
        if status is not None:
            self.status_var.set(status)
        if busy:
            self.progress.configure(mode="indeterminate")
            self.progress.start(12)
        else:
            self.progress.stop()
            self.progress.configure(mode="determinate")
            self.progress["value"] = 0

    def _set_all(self, value: bool) -> None:
        self.checked = set(range(len(self.candidates))) if value else set()
        self._render_arrays()
        self.btn_export.configure(state="normal" if self.checked and not self.busy else "disabled")
        self.refresh_preview()

    def _selected(self) -> list[core.ArrayCandidate]:
        return [self.candidates[i] for i in sorted(self.checked) if 0 <= i < len(self.candidates)]

    # --------------------------------------------------------------- nạp file

    def on_open(self) -> None:
        path = filedialog.askopenfilename(
            title="Chọn file JSON",
            filetypes=[("File JSON", "*.json"), ("Tất cả các file", "*.*")],
        )
        if path:
            self.load_file(path)

    def load_file(self, path: str) -> None:
        self.file_path = path
        self.path_var.set(path)
        self.candidates = []
        self.checked = set()
        self._render_arrays()
        self._clear_preview()
        self.btn_export.configure(state="disabled")

        flatten, depth = self.flatten_var.get(), self.depth_var.get()
        self._set_busy(True, f"Đang đọc {os.path.basename(path)}…")

        def work() -> None:
            try:
                data = core.load_json(path)
                candidates = core.find_arrays(data)
                for c in candidates:
                    c.columns = core.scan_columns(c.rows, flatten, max_depth=depth)
                self.queue.put(("loaded", candidates, None))
            except Exception as exc:  # noqa: BLE001 - hiển thị mọi lỗi cho người dùng
                self.queue.put(("loaded", None, exc))

        threading.Thread(target=work, daemon=True).start()

    def _render_arrays(self) -> None:
        """Vẽ lại bảng theo đúng trạng thái `self.checked` (không tự tick gì)."""
        self.tree_arr.delete(*self.tree_arr.get_children())
        for i, c in enumerate(self.candidates):
            tick = TICKED if i in self.checked else UNTICKED
            sheet = c.sheet_name or core._default_sheet_name(c.path)
            self.tree_arr.insert(
                "", "end", iid=str(i), values=(tick, c.path, f"{c.count:,}", len(c.columns), sheet)
            )

    def _auto_select_biggest(self) -> None:
        """Tick sẵn mảng nhiều dòng nhất (ưu tiên mảng không rỗng)."""
        if not self.candidates:
            return
        non_empty = [i for i, c in enumerate(self.candidates) if c.count]
        pool = non_empty or range(len(self.candidates))
        self.checked = {max(pool, key=lambda i: self.candidates[i].count)}

    def on_array_click(self, event) -> None:
        if self.busy:
            return
        row = self.tree_arr.identify_row(event.y)
        if not row:
            return
        i = int(row)
        if i in self.checked:
            self.checked.discard(i)
            self.tree_arr.set(row, "tick", UNTICKED)
        else:
            self.checked.add(i)
            self.tree_arr.set(row, "tick", TICKED)
        self.btn_export.configure(state="normal" if self.checked else "disabled")
        self.refresh_preview()

    def on_options_changed(self) -> None:
        if not self.candidates or self.busy:
            return
        try:
            depth = max(1, int(self.depth_var.get()))
        except (tk.TclError, ValueError):
            return
        self.depth_var.set(depth)
        flatten = self.flatten_var.get()
        for c in self.candidates:
            c.columns = core.scan_columns(c.rows, flatten, max_depth=depth)
        for i, c in enumerate(self.candidates):
            self.tree_arr.set(str(i), "cols", len(c.columns))
        self.refresh_preview()

    # -------------------------------------------------------------- xem trước

    def _clear_preview(self) -> None:
        self.tree_prev.delete(*self.tree_prev.get_children())
        self.tree_prev["columns"] = ()
        self.lbl_prev.configure(text="—")

    def refresh_preview(self) -> None:
        sel = self.tree_arr.selection()
        idx = int(sel[0]) if sel else (min(self.checked) if self.checked else None)
        if idx is None or not (0 <= idx < len(self.candidates)):
            self._clear_preview()
            return
        cand = self.candidates[idx]
        columns = cand.columns
        if not columns:
            self._clear_preview()
            self.lbl_prev.configure(text=f"“{cand.path}” không có cột nào.")
            return

        self.tree_prev.delete(*self.tree_prev.get_children())
        self.tree_prev["columns"] = columns

        flatten, depth = self.flatten_var.get(), self.depth_var.get()
        shown = cand.rows[:PREVIEW_ROWS]
        rendered = []
        for rec in shown:
            norm = core.normalize_record(rec, flatten, max_depth=depth)
            rendered.append([preview_text(norm.get(c)) for c in columns])

        # Bề rộng cột tính theo NỘI DUNG (không phải tiêu đề), nếu không cột nào
        # cũng hẹp bằng tên cột và nội dung bị cắt hết.
        char_px = tkfont.nametofont("TkDefaultFont").measure("0") or px(8)
        for j, col in enumerate(columns):
            self.tree_prev.heading(col, text=col)
            widest = len(col)
            for row_vals in rendered[:60]:
                widest = max(widest, len(row_vals[j]))
            width = min(max(widest + 2, 10), 34) * char_px
            self.tree_prev.column(col, width=width, minwidth=char_px * 6, anchor="w", stretch=False)

        for row_vals in rendered:
            self.tree_prev.insert("", "end", values=row_vals)

        extra = "" if cand.count <= PREVIEW_ROWS else f" (còn {cand.count - PREVIEW_ROWS:,} dòng nữa)"
        self.lbl_prev.configure(
            text=f"“{cand.path}” — {cand.count:,} dòng × {len(columns)} cột, hiện {len(shown)} dòng đầu{extra}"
        )

    # ---------------------------------------------------------------- xuất

    def on_export(self) -> None:
        chosen = self._selected()
        if not chosen:
            messagebox.showwarning("Chưa chọn dữ liệu", "Bạn cần tick ít nhất một bảng dữ liệu.")
            return
        default = os.path.splitext(os.path.basename(self.file_path or "ket_qua.json"))[0] + ".xlsx"
        out = filedialog.asksaveasfilename(
            title="Lưu file Excel",
            defaultextension=".xlsx",
            initialfile=default,
            filetypes=[("Excel workbook", "*.xlsx")],
        )
        if not out:
            return
        self.last_out_dir = os.path.dirname(out)
        self.btn_folder.configure(state="normal")

        flatten, depth = self.flatten_var.get(), self.depth_var.get()
        total = sum(c.count for c in chosen)
        self._set_busy(True, f"Đang ghi {total:,} dòng…")

        def work() -> None:
            try:
                report = core.export_xlsx(
                    chosen,
                    out,
                    flatten=flatten,
                    max_depth=depth,
                    progress_cb=lambda done, tot: self.queue.put(("progress", done, tot)),
                )
                self.queue.put(("exported", report, None))
            except Exception as exc:  # noqa: BLE001
                self.queue.put(("exported", None, exc))

        threading.Thread(target=work, daemon=True).start()

    def on_open_folder(self) -> None:
        if not self.last_out_dir or not os.path.isdir(self.last_out_dir):
            return
        try:
            if sys.platform == "win32":
                os.startfile(self.last_out_dir)  # noqa: S606
            else:
                subprocess.Popen(["xdg-open", self.last_out_dir])
        except OSError as exc:
            messagebox.showerror("Không mở được thư mục", str(exc))

    # ------------------------------------------------------- vòng lặp sự kiện

    def _poll_queue(self) -> None:
        try:
            while True:
                kind, a, b = self.queue.get_nowait()
                if kind == "loaded":
                    self._on_loaded(a, b)
                elif kind == "progress":
                    if str(self.progress["mode"]) != "determinate":
                        self.progress.stop()  # dừng timer của chế độ chạy vô định
                        self.progress.configure(mode="determinate")
                    self.progress["value"] = (a / b * 100) if b else 0
                    self.status_var.set(f"Đang ghi {a:,}/{b:,} dòng…")
                elif kind == "exported":
                    self._on_exported(a, b)
        except queue.Empty:
            pass
        self.after(80, self._poll_queue)

    def _on_loaded(self, candidates, error) -> None:
        self._set_busy(False)
        if error is not None:
            self.status_var.set("Đọc file thất bại.")
            messagebox.showerror("Không đọc được file JSON", f"{type(error).__name__}: {error}")
            return
        self.candidates = candidates
        if not candidates:
            self.status_var.set("Không tìm thấy mảng dữ liệu nào trong file.")
            messagebox.showwarning(
                "Không có bảng dữ liệu",
                "File JSON này không chứa mảng nào để xuất.\n"
                "Nếu file chỉ có một object đơn lẻ, hãy bọc nó trong một mảng.",
            )
            return
        self._auto_select_biggest()
        self._render_arrays()
        self.btn_export.configure(state="normal" if self.checked else "disabled")
        self.refresh_preview()
        self.status_var.set(f"Đã đọc {len(candidates)} bảng dữ liệu từ {os.path.basename(self.file_path or '')}.")

    def _on_exported(self, report, error) -> None:
        self._set_busy(False)
        if error is not None:
            self.status_var.set("Xuất file thất bại.")
            messagebox.showerror("Không ghi được file Excel", f"{type(error).__name__}: {error}")
            return
        self.status_var.set(f"Xong: {report.total_rows:,} dòng trong {len(report.sheets)} sheet.")
        messagebox.showinfo("Hoàn tất", report.summary())
        self.btn_export.configure(state="normal" if self.checked else "disabled")


def _ensure_streams() -> None:
    """Bản .exe dựng ở chế độ --windowed không có console: sys.stdout là None.

    Gán tạm vào devnull để mọi lệnh print không làm chương trình chết.
    """
    for name in ("stdout", "stderr"):
        if getattr(sys, name, None) is None:
            try:
                setattr(sys, name, open(os.devnull, "w", encoding="utf-8"))
            except OSError:
                pass


def main(argv: list[str] | None = None) -> int:
    argv = sys.argv[1:] if argv is None else argv
    _ensure_streams()

    # Chế độ không giao diện, dùng để chuyển hàng loạt hoặc kiểm thử bản .exe:
    #   "JSON sang Excel.exe" --convert input.json -o output.xlsx
    if "--convert" in argv:
        try:
            return core._main([a for a in argv if a != "--convert"])
        except SystemExit as exc:
            return int(exc.code or 0)
        except Exception as exc:  # noqa: BLE001 - không có console để xem traceback
            try:
                root = tk.Tk()
                root.withdraw()
                messagebox.showerror("Lỗi chuyển đổi", f"{type(exc).__name__}: {exc}")
                root.destroy()
            except tk.TclError:
                pass
            return 1

    try:  # chữ nét hơn trên màn hình Windows hiển thị tỉ lệ
        import ctypes

        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:  # noqa: BLE001 - không phải Windows hoặc Windows cũ
        pass

    try:
        App(initial_file=argv[0] if argv else None).mainloop()
    except Exception:  # noqa: BLE001 - chạy bằng pythonw thì không có console để xem lỗi
        err = traceback.format_exc()
        try:
            root = tk.Tk()
            root.withdraw()
            messagebox.showerror("Lỗi không mong đợi", err)
            root.destroy()
        except tk.TclError:
            print(err, file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
