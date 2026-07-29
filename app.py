from __future__ import annotations

import threading
from pathlib import Path
from tkinter import Canvas, filedialog, messagebox

import customtkinter as ctk

from processor import STAGE_FUNCTIONS, load_address_master


APP_TITLE = "MPM to FM Data cleaner"
BG = "#FFF5FA"
PANEL = "#F4FAFF"
CARD = "#FFFFFF"
PINK = "#F8BBD0"
PINK_HOVER = "#F48FB1"
BLUE = "#A7D8FF"
BLUE_HOVER = "#7EC8F8"
TEXT = "#344054"
MUTED = "#667085"


class CleanerApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("980x660")
        self.minsize(860, 580)
        self.configure(fg_color=BG)

        ctk.set_appearance_mode("Light")

        self.selected_files: dict[int, Path] = {}
        self.file_labels: dict[int, ctk.CTkLabel] = {}
        self.buttons: list[ctk.CTkButton] = []

        self._build_ui()

    def _build_ui(self):
        root = ctk.CTkFrame(self, corner_radius=0, fg_color=BG)
        root.pack(fill="both", expand=True)

        header = ctk.CTkFrame(root, fg_color=BG)
        header.pack(fill="x", padx=22, pady=(16, 2))
        header.grid_columnconfigure(0, weight=1)

        title = ctk.CTkLabel(header, text=APP_TITLE, font=ctk.CTkFont(size=28, weight="bold"), text_color=TEXT)
        title.grid(row=0, column=0, sticky="w")

        mascot = Canvas(header, width=170, height=92, bg=BG, highlightthickness=0)
        mascot.grid(row=0, column=1, sticky="e")
        self._draw_mascots(mascot)

        panel = ctk.CTkScrollableFrame(root, fg_color=PANEL, corner_radius=14)
        panel.pack(fill="both", expand=True, padx=22, pady=(14, 14))

        stages = [
            (1, "Import Original File", "Export Check File 1", "Stage 1: Steps 1-6. If no check is needed, it continues to Step 11 automatically."),
            (2, "Import Check File 1", "Export Check File 2", "Stage 2: Steps 7-11"),
            (3, "Import Check File 2", "Export Check File 3", "Stage 3: Steps 12-14"),
            (4, "Import Check File 3", "Export Check File 4", "Stage 4: Steps 15-25"),
            (5, "Import Check File 4", "Export Check File 5", "Stage 5: Steps 26-31. Add Sheet1 manually after this export."),
            (6, "Import Check File 5", "Export Final Excel", "Stage 6: Step 32. Updates CustomersForFilemakerJP using Sheet1 lookup."),
        ]

        for stage, import_text, export_text, desc in stages:
            self._stage_row(panel, stage, import_text, export_text, desc)

        master = load_address_master()
        master_status = f"Address master: {Path(master['path']).name}" if master.get("path") else "Address master: not found"
        self.status = ctk.CTkLabel(root, text=f"Ready · {master_status}", anchor="w", text_color=TEXT)
        self.status.pack(fill="x", padx=22, pady=(0, 14))

    def _stage_row(self, parent, stage: int, import_text: str, export_text: str, desc: str):
        frame = ctk.CTkFrame(parent, corner_radius=14, fg_color=CARD, border_width=1, border_color="#E8EEF7")
        frame.pack(fill="x", padx=8, pady=8)
        frame.grid_columnconfigure(1, weight=1)

        badge = ctk.CTkLabel(
            frame,
            text=f"Stage {stage}",
            width=82,
            font=ctk.CTkFont(weight="bold"),
            text_color=TEXT,
            fg_color="#FFE4EF" if stage % 2 else "#E0F2FF",
            corner_radius=12,
        )
        badge.grid(row=0, column=0, rowspan=2, padx=14, pady=12, sticky="ns")

        desc_label = ctk.CTkLabel(frame, text=desc, anchor="w", font=ctk.CTkFont(size=14, weight="bold"), text_color=TEXT)
        desc_label.grid(row=0, column=1, sticky="ew", padx=(0, 12), pady=(12, 2))

        file_label = ctk.CTkLabel(frame, text="No file selected", anchor="w", text_color=MUTED)
        file_label.grid(row=1, column=1, sticky="ew", padx=(0, 12), pady=(2, 12))
        self.file_labels[stage] = file_label

        import_btn = ctk.CTkButton(
            frame,
            text=import_text,
            width=190,
            fg_color=BLUE,
            hover_color=BLUE_HOVER,
            text_color="#16324F",
            command=lambda s=stage: self.select_file(s),
        )
        import_btn.grid(row=0, column=2, padx=(0, 10), pady=12, sticky="e")

        export_btn = ctk.CTkButton(
            frame,
            text=export_text,
            width=190,
            fg_color=PINK,
            hover_color=PINK_HOVER,
            text_color="#4A1930",
            command=lambda s=stage: self.export_stage(s),
        )
        export_btn.grid(row=1, column=2, padx=(0, 10), pady=12, sticky="e")
        self.buttons.extend([import_btn, export_btn])

    def _draw_mascots(self, canvas: Canvas):
        # Simple pastel sky decorations.
        canvas.create_oval(16, 18, 58, 60, fill="#FFD976", outline="#F7B955", width=2)
        for x1, y1, x2, y2 in [(37, 4, 37, 14), (37, 64, 37, 76), (2, 39, 12, 39), (62, 39, 74, 39),
                               (12, 14, 20, 22), (56, 14, 64, 22), (12, 64, 20, 56), (56, 64, 64, 56)]:
            canvas.create_line(x1, y1, x2, y2, fill="#F7B955", width=2)
        canvas.create_oval(30, 32, 34, 36, fill=TEXT, outline="")
        canvas.create_oval(43, 32, 47, 36, fill=TEXT, outline="")
        canvas.create_arc(31, 37, 47, 49, start=200, extent=140, style="arc", outline=TEXT, width=2)

        canvas.create_oval(88, 18, 138, 68, fill="#A7D8FF", outline="#7EC8F8", width=2)
        canvas.create_oval(106, 12, 150, 62, fill=BG, outline=BG)
        canvas.create_oval(101, 38, 104, 41, fill="#7EC8F8", outline="")
        canvas.create_oval(113, 48, 116, 51, fill="#7EC8F8", outline="")

        def star(cx, cy, r, color):
            pts = [
                cx, cy - r, cx + r * 0.28, cy - r * 0.28, cx + r, cy,
                cx + r * 0.28, cy + r * 0.28, cx, cy + r,
                cx - r * 0.28, cy + r * 0.28, cx - r, cy,
                cx - r * 0.28, cy - r * 0.28,
            ]
            canvas.create_polygon(pts, fill=color, outline="")

        star(164, 20, 8, "#F8BBD0")

    def select_file(self, stage: int):
        file_path = filedialog.askopenfilename(
            title=f"Select Excel file for Stage {stage}",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
        )
        if not file_path:
            return
        path = Path(file_path)
        self.selected_files[stage] = path
        self.file_labels[stage].configure(text=str(path))
        self.set_status(f"Stage {stage} selected: {path.name}")

    def export_stage(self, stage: int):
        input_path = self.selected_files.get(stage)
        if not input_path:
            messagebox.showwarning("No file selected", f"Please import an Excel file for Stage {stage} first.")
            return

        default_suffix = "Final" if stage == 6 else f"Check_{stage}"
        if stage == 1:
            default_suffix = "Check_1_or_Auto_Check_2"
        default_name = f"{input_path.stem}_{default_suffix}.xlsx"
        output_path = filedialog.asksaveasfilename(
            title="Choose export location",
            initialfile=default_name,
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
        )
        if not output_path:
            return

        self.set_busy(True)
        self.set_status(f"Stage {stage} is running...")
        thread = threading.Thread(target=self._run_stage, args=(stage, input_path, Path(output_path)), daemon=True)
        thread.start()

    def _run_stage(self, stage: int, input_path: Path, output_path: Path):
        try:
            result = STAGE_FUNCTIONS[stage](input_path, output_path)
        except Exception as exc:
            self.after(0, lambda: self._stage_failed(stage, exc))
            return
        self.after(0, lambda: self._stage_done(stage, output_path, result))

    def _stage_done(self, stage: int, output_path: Path, result=None):
        self.set_busy(False)
        if stage == 1 and result == "stage2_check":
            msg = (
                "No rows had both BX and BS filled, so Check File 1 is not needed.\n\n"
                "The app continued through Steps 7-11 and exported the next review file.\n"
                "Next, use Stage 3 with this exported file."
            )
            self.set_status(f"Stage 1 skipped Check File 1 and exported Stage 2 review: {output_path}")
            messagebox.showinfo("Check skipped", f"{msg}\n\nSaved to:\n{output_path}")
            return
        self.set_status(f"Stage {stage} complete: {output_path}")
        messagebox.showinfo("Done", f"Stage {stage} exported:\n{output_path}")

    def _stage_failed(self, stage: int, exc: Exception):
        self.set_busy(False)
        self.set_status(f"Stage {stage} failed")
        messagebox.showerror("Processing failed", f"Stage {stage} failed:\n{exc}")

    def set_busy(self, busy: bool):
        state = "disabled" if busy else "normal"
        for button in self.buttons:
            button.configure(state=state)

    def set_status(self, text: str):
        self.status.configure(text=text)


if __name__ == "__main__":
    app = CleanerApp()
    app.mainloop()
