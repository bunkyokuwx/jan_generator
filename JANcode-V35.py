import os
import sys
import platform
import threading
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from PIL import Image, ImageDraw, ImageFont, ImageTk

# ==========================================
# クロスプラットフォームフォント設定
# ==========================================
SYS_FONT = "Meiryo" if platform.system() == "Windows" else "Hiragino Sans"

# ==========================================
# コア：自動フォントアウトライン化エンジン (TTF -> EPS)
# ==========================================
try:
    from fontTools.ttLib import TTFont
    from fontTools.pens.basePen import BasePen
    HAS_FONTTOOLS = True

    class EPSOutlinePen(BasePen):
        def __init__(self, glyphSet):
            super().__init__(glyphSet)
            self.commands = []

        def _moveTo(self, pt):
            self.commands.append(f"{pt[0]:.3f} {pt[1]:.3f} m")

        def _lineTo(self, pt):
            self.commands.append(f"{pt[0]:.3f} {pt[1]:.3f} l")

        def _curveToOne(self, pt1, pt2, pt3):
            self.commands.append(f"{pt1[0]:.3f} {pt1[1]:.3f} {pt2[0]:.3f} {pt2[1]:.3f} {pt3[0]:.3f} {pt3[1]:.3f} c")

        def _qCurveToOne(self, pt1, pt2):
            p0 = self._getCurrentPoint()
            cp1x = p0[0] + (2.0/3.0) * (pt1[0] - p0[0])
            cp1y = p0[1] + (2.0/3.0) * (pt1[1] - p0[1])
            cp2x = pt2[0] + (2.0/3.0) * (pt1[0] - pt2[0])
            cp2y = pt2[1] + (2.0/3.0) * (pt1[1] - pt2[1])
            self.commands.append(f"{cp1x:.3f} {cp1y:.3f} {cp2x:.3f} {cp2y:.3f} {pt2[0]:.3f} {pt2[1]:.3f} c")

        def _closePath(self):
            self.commands.append("d")

except ImportError:
    HAS_FONTTOOLS = False

# ==========================================
# JAN コアロジック
# ==========================================
L_CODES = ["0001101", "0011001", "0010011", "0111101", "0100011", "0110001", "0101111", "0111011", "0110111", "0001011"]
G_CODES = ["0100111", "0110011", "0011011", "0100001", "0011101", "0111001", "0000101", "0010001", "0001001", "0010111"]
R_CODES = ["1110010", "1100110", "1101100", "1000010", "1011100", "1001110", "1010000", "1000100", "1001000", "1110100"]
PARITY  = ["LLLLLL", "LLGLGG", "LLGGLG", "LLGGGL", "LGLLGG", "LGGLLG", "LGGGLL", "LGLGLG", "LGLGGL", "LGGLGL"]

def generate_ean13_binary(code13):
    first_digit = int(code13[0])
    parity_pattern = PARITY[first_digit]
    binary = "101" 
    for i in range(6):
        digit = int(code13[i+1])
        if parity_pattern[i] == 'L': binary += L_CODES[digit]
        else: binary += G_CODES[digit]
    binary += "01010" 
    for i in range(6, 12):
        digit = int(code13[i+1])
        binary += R_CODES[digit]
    binary += "101" 
    return binary

def get_ocrb_font_path():
    if hasattr(sys, '_MEIPASS'):
        bundled_font = os.path.join(sys._MEIPASS, "OCRB.ttf")
        if os.path.exists(bundled_font): return bundled_font
        
    local_font = os.path.join(os.path.abspath("."), "OCRB.ttf")
    if os.path.exists(local_font): return local_font

    font_names = ["OCRB.ttf", "OCRB.ttc", "OCRB.otf", "ocrb10.ttf", "ocrbreg.ttf"]
    if platform.system() == "Windows":
        import winreg
        try:
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts") as key:
                for i in range(winreg.QueryInfoKey(key)[1]):
                    name, value, _ = winreg.EnumValue(key, i)
                    if "ocrb" in name.lower() or "ocr-b" in name.lower():
                        path = value if os.path.isabs(value) else os.path.join(os.environ.get('WINDIR', 'C:\\Windows'), 'Fonts', value)
                        if os.path.exists(path): return path
        except Exception: pass
        
        fallback_dirs = [os.path.join(os.environ.get('WINDIR', 'C:\\Windows'), 'Fonts')]
        for d in fallback_dirs:
            for fname in font_names:
                p = os.path.join(d, fname)
                if os.path.exists(p): return p
    
    elif platform.system() == "Darwin":
        mac_dirs = [
            os.path.expanduser("~/Library/Fonts"),
            "/Library/Fonts",
            "/System/Library/Fonts"
        ]
        for d in mac_dirs:
            for fname in font_names:
                p = os.path.join(d, fname)
                if os.path.exists(p): return p
    return None

class JANCodeGenerator:
    def __init__(self):
        self.font_path = get_ocrb_font_path()

    def calculate_check_digit(self, code_12):
        s_odd = sum(int(code_12[i]) for i in range(0, 12, 2))
        s_even = sum(int(code_12[i]) for i in range(1, 12, 2))
        return (10 - ((s_odd + (s_even * 3)) % 10)) % 10

    def complete_jan_code(self, code_12):
        if len(code_12) != 12 or not code_12.isdigit(): return ""
        check_digit = self.calculate_check_digit(code_12)
        return f"{code_12}{check_digit}"

    def get_char_positions(self, jan_code, padding_x, module_w):
        positions = []
        positions.append((jan_code[0], padding_x - 8.0 * module_w))
        for i in range(6):
            char_center = padding_x + (3 + i * 7 + 3.5) * module_w
            positions.append((jan_code[i+1], char_center))
        for i in range(6):
            char_center = padding_x + (50 + i * 7 + 3.5) * module_w
            positions.append((jan_code[i+7], char_center))
        return positions

    def draw_ean13_png(self, jan_code, filepath, scale=1.0, add_frame=False):
        h_pt_new = 40.4787
        w_pt_base = 105.703 
        img_w = int(113 * 12 * scale) 
        img_h = int(img_w * (h_pt_new / w_pt_base))
        module_w = img_w / 113.0
        padding_x = 11 * module_w
        
        top_pad_pt = 0.935
        guard_bot_pt = 4.904
        norm_bot_pt = 9.581
        text_base_pt = 0.95
        font_size_pt_base = 11.0

        bar_y_top = img_h * (top_pad_pt / h_pt_new)
        guard_y_bottom = img_h * ((h_pt_new - guard_bot_pt) / h_pt_new)
        normal_y_bottom = img_h * ((h_pt_new - norm_bot_pt) / h_pt_new)
        text_y_baseline = img_h * ((h_pt_new - text_base_pt) / h_pt_new)
        font_size = int(img_h * (font_size_pt_base / h_pt_new))
        
        width_mm = 37.29 * scale
        dpi_val = img_w / (width_mm / 25.4) 

        img = Image.new("RGB", (img_w, img_h), "white")
        draw = ImageDraw.Draw(img)

        binary = generate_ean13_binary(jan_code)
        for i, bit in enumerate(binary):
            if bit == '1':
                is_guard = (i < 3) or (45 <= i <= 49) or (i > 91)
                y_bottom = guard_y_bottom if is_guard else normal_y_bottom
                x = padding_x + i * module_w
                draw.rectangle([x, bar_y_top, x + module_w - 1, y_bottom], fill="black")

        try:
            if self.font_path: font = ImageFont.truetype(self.font_path, font_size)
            else: font = ImageFont.load_default() 
        except IOError:
            font = ImageFont.load_default()

        positions = self.get_char_positions(jan_code, padding_x, module_w)
        for char, center_x in positions:
            try:
                bbox = draw.textbbox((0, 0), char, font=font)
                w = bbox[2] - bbox[0]
                h = bbox[3] - bbox[1]
            except AttributeError:
                w, h = draw.textsize(char, font=font)
            draw.text((center_x - w / 2, text_y_baseline - h), char, font=font, fill="black")

        if scale == 1.5 and add_frame:
            px_per_mm = img_w / width_mm
            thick_px = int(3.0 * px_per_mm)
            pad_x_px = int(3.5 * px_per_mm)
            pad_y_px = int(4.0 * px_per_mm)
            
            new_w = img_w + 2 * (thick_px + pad_x_px)
            new_h = img_h + 2 * (thick_px + pad_y_px)
            
            framed_img = Image.new("RGB", (new_w, new_h), "black")
            draw_frame = ImageDraw.Draw(framed_img)
            draw_frame.rectangle([thick_px, thick_px, new_w - thick_px - 1, new_h - thick_px - 1], fill="white")
            framed_img.paste(img, (thick_px + pad_x_px, thick_px + pad_y_px))
            img = framed_img

        img.save(filepath, dpi=(dpi_val, dpi_val))

    def draw_ean13_eps_vector(self, jan_code, filepath, scale=1.0, add_frame=False):
        h_pt_new = 40.4787
        w_pt_base = 105.703
        w_pt = w_pt_base * scale
        h_pt = h_pt_new * scale
        
        pt_per_mm = 72.0 / 25.4
        module_pt = w_pt / 113.0
        padding_x_pt = 11 * module_pt
        
        top_pad_pt = 0.935
        bar_y_top = (h_pt_new - top_pad_pt) * scale
        guard_y_bottom = 4.904 * scale
        normal_y_bottom = 9.581 * scale
        text_y_baseline = 0.95 * scale 
        font_size_pt = 11.0 * scale

        offset_x_pt = 0
        offset_y_pt = 0
        if scale == 1.5 and add_frame:
            thick_pt = 3.0 * pt_per_mm
            pad_x_pt = 3.5 * pt_per_mm
            pad_y_pt = 4.0 * pt_per_mm
            
            bbox_w = w_pt + 2 * (thick_pt + pad_x_pt)
            bbox_h = h_pt + 2 * (thick_pt + pad_y_pt)
            offset_x_pt = thick_pt + pad_x_pt
            offset_y_pt = thick_pt + pad_y_pt
        else:
            bbox_w = w_pt
            bbox_h = h_pt

        lines = [
            "%!PS-Adobe-3.0 EPSF-3.0",
            f"%%BoundingBox: 0 0 {int(bbox_w+1)} {int(bbox_h+1)}",
            "%%EndComments",
            "/m {moveto} def",
            "/l {lineto} def",
            "/c {curveto} def",
            "/n {newpath} def",
            "/d {closepath} def",
            "/f {eofill} def",
            "/cshow { dup stringwidth pop 2 div neg 0 rmoveto show } def"
        ]

        if scale == 1.5 and add_frame:
            lines.extend([
                "0 0 0 setrgbcolor",
                f"0 0 {bbox_w:.3f} {bbox_h:.3f} rectfill",
                "1 1 1 setrgbcolor",
                f"{thick_pt:.3f} {thick_pt:.3f} {(bbox_w - 2*thick_pt):.3f} {(bbox_h - 2*thick_pt):.3f} rectfill"
            ])

        lines.append("0 0 0 setrgbcolor")

        binary = generate_ean13_binary(jan_code)
        for i, bit in enumerate(binary):
            if bit == '1':
                is_guard = (i < 3) or (45 <= i <= 49) or (i > 91)
                y_bottom = guard_y_bottom if is_guard else normal_y_bottom
                bar_h = bar_y_top - y_bottom
                x = padding_x_pt + i * module_pt + offset_x_pt
                y = y_bottom + offset_y_pt
                lines.append(f"{x:.3f} {y:.3f} {module_pt:.3f} {bar_h:.3f} rectfill")

        positions = self.get_char_positions(jan_code, padding_x_pt, module_pt)
        eps_outlined = False
        
        if HAS_FONTTOOLS and self.font_path:
            font = None
            try:
                font = TTFont(self.font_path, fontNumber=0) if self.font_path.lower().endswith('.ttc') else TTFont(self.font_path)
                cmap = font.getBestCmap()
                glyphSet = font.getGlyphSet()
                upm = font['head'].unitsPerEm
                scale_factor = font_size_pt / upm

                for char, center_x in positions:
                    if ord(char) in cmap:
                        glyph_name = cmap[ord(char)]
                        glyph = glyphSet[glyph_name]
                        pen = EPSOutlinePen(glyphSet)
                        glyph.draw(pen)
                        
                        glyph_width = getattr(glyph, 'width', 0)
                        start_x = (center_x + offset_x_pt) - (glyph_width * scale_factor) / 2.0
                        cy = text_y_baseline + offset_y_pt
                        
                        lines.append("gsave")
                        lines.append(f"{start_x:.3f} {cy:.3f} translate")
                        lines.append(f"{scale_factor:.6f} {scale_factor:.6f} scale")
                        lines.append("n")
                        lines.extend(pen.commands)
                        lines.append("f")
                        lines.append("grestore")
                eps_outlined = True
            except Exception:
                pass
            finally:
                if font is not None:
                    font.close()

        if not eps_outlined:
            lines.append(f"/OCRB findfont {font_size_pt:.3f} scalefont setfont")
            for char, center_x in positions:
                cx = center_x + offset_x_pt
                cy = text_y_baseline + offset_y_pt
                lines.append(f"{cx:.3f} {cy:.3f} moveto ({char}) cshow")
            
        lines.append("%%EOF")
        
        with open(filepath, 'w', encoding='utf-8') as f:
            f.write("\n".join(lines))

# ==========================================
# UI コアコンポーネント
# ==========================================
class AutoScrollbar(ttk.Scrollbar):
    def set(self, lo, hi):
        if float(lo) <= 0.0 and float(hi) >= 1.0:
            self.grid_remove() 
        else:
            self.grid()
        super().set(lo, hi)

class RoundedButton(tk.Label):
    def __init__(self, master, text, command, width=100, height=32, radius=8, 
                 bg_color="#FFFFFF", fg_color="#374151", hover_bg="#E5E7EB", 
                 disabled_bg="#D1D5DB", parent_bg="#F8F9FA", font=(SYS_FONT, 10, "bold"), **kwargs):
        self.normal_img = self._create_img(width, height, radius, bg_color, parent_bg)
        self.hover_img = self._create_img(width, height, radius, hover_bg, parent_bg)
        self.disabled_img = self._create_img(width, height, radius, disabled_bg, parent_bg)
        
        super().__init__(master, text=text, image=self.normal_img, compound="center", 
                         bg=parent_bg, fg=fg_color, font=font, cursor="hand2", **kwargs)
        
        self.command = command
        self.fg_color = fg_color
        self.is_disabled = False

        self.bind("<Enter>", self.on_hover)
        self.bind("<Leave>", self.on_leave)
        self.bind("<ButtonRelease-1>", self.on_release)

    def _create_img(self, w, h, r, color, parent_bg):
        scale = 4
        img = Image.new("RGBA", (w * scale, h * scale), parent_bg)
        draw = ImageDraw.Draw(img)
        draw.rounded_rectangle((0, 0, w * scale - 1, h * scale - 1), radius=r * scale, fill=color)
        img = img.resize((w, h), Image.Resampling.LANCZOS)
        return ImageTk.PhotoImage(img)

    def on_hover(self, e):
        if not self.is_disabled: self.config(image=self.hover_img)
    def on_leave(self, e):
        if not self.is_disabled: self.config(image=self.normal_img)
    def on_release(self, e):
        if not self.is_disabled and self.command: self.command()
    def set_state(self, state):
        if state == tk.DISABLED:
            self.is_disabled = True
            self.config(image=self.disabled_img, cursor="arrow", fg="#9CA3AF")
        else:
            self.is_disabled = False
            self.config(image=self.normal_img, cursor="hand2", fg=self.fg_color)

class CustomCheckbutton(tk.Frame):
    def __init__(self, master, text, variable, command=None, bg="#F3F4F6", fg="#374151", size=22, **kwargs):
        super().__init__(master, bg=bg, **kwargs)
        self.variable = variable
        self.command = command
        self.is_disabled = False
        self.fg_color = fg
        self.bg_color = bg  

        self.img_off = self._create_img(checked=False, disabled=False, size=size)
        self.img_on = self._create_img(checked=True, disabled=False, size=size)
        self.img_off_dis = self._create_img(checked=False, disabled=True, size=size)
        self.img_on_dis = self._create_img(checked=True, disabled=True, size=size)

        self.icon = tk.Label(self, image=self.img_off if not variable.get() else self.img_on, bg=bg, cursor="hand2")
        self.icon.pack(side="left")
        
        self.lbl = tk.Label(self, text=text, bg=bg, fg=fg, font=(SYS_FONT, 10), cursor="hand2")
        self.lbl.pack(side="left", padx=(5, 0))

        self.icon.bind("<Button-1>", self.toggle)
        self.lbl.bind("<Button-1>", self.toggle)
        self.variable.trace_add("write", self._on_var_change)

    def _create_img(self, checked, disabled, size):
        scale = 4
        w, h = size * scale, size * scale
        img = Image.new("RGBA", (w, h), self.bg_color)
        draw = ImageDraw.Draw(img)
        
        box_color = "#E5E7EB" if disabled else "#FFFFFF"
        border_color = "#D1D5DB" if disabled else "#9CA3AF"
        check_bg = "#9CA3AF" if disabled else "#3B82F6"
        radius = 4 * scale
        
        if checked:
            draw.rounded_rectangle((0, 0, w-1, h-1), radius=radius, fill=check_bg)
            tick_pts = [(int(0.25*w), int(0.55*h)), (int(0.45*w), int(0.75*h)), (int(0.78*w), int(0.3*h))]
            draw.line(tick_pts, fill="#FFFFFF", width=int(0.12*w), joint="curve")
        else:
            draw.rounded_rectangle((0, 0, w-1, h-1), radius=radius, fill=box_color, outline=border_color, width=1*scale)
            
        img = img.resize((size, size), Image.Resampling.LANCZOS)
        return ImageTk.PhotoImage(img)

    def toggle(self, event=None):
        if not self.is_disabled:
            self.variable.set(not self.variable.get())
            if self.command: self.command()

    def _on_var_change(self, *args):
        if self.is_disabled:
            self.icon.config(image=self.img_on_dis if self.variable.get() else self.img_off_dis)
        else:
            self.icon.config(image=self.img_on if self.variable.get() else self.img_off)

    def config_state(self, state):
        if state == tk.DISABLED:
            self.is_disabled = True
            self.icon.config(cursor="arrow", image=self.img_on_dis if self.variable.get() else self.img_off_dis)
            self.lbl.config(cursor="arrow", fg="#9CA3AF")
        else:
            self.is_disabled = False
            self.icon.config(cursor="hand2", image=self.img_on if self.variable.get() else self.img_off)
            self.lbl.config(cursor="hand2", fg=self.fg_color)

# ==========================================
# 動的入力行
# ==========================================
class JanInputRow:
    def __init__(self, parent, generator, remove_callback):
        self.generator = generator
        self.remove_callback = remove_callback
        self.current_jan = ""
        
        self.container = tk.Frame(parent, bg="#F8F9FA")
        self.container.pack(fill=tk.X, pady=(0, 6))

        font_header = (SYS_FONT, 8)
        font_val = (SYS_FONT, 12, "bold")

        table_border = tk.Frame(self.container, bg="#D1D5DB", padx=1, pady=1)
        table_border.pack(side=tk.LEFT, anchor=tk.S)

        tk.Label(table_border, text="国コード", bg="#F1F5F9", fg="#475569", font=font_header, pady=2).grid(row=0, column=0, padx=(0,1), pady=(0,1), sticky="nsew")
        tk.Label(table_border, text="メーカー", bg="#F1F5F9", fg="#475569", font=font_header, pady=2).grid(row=0, column=1, padx=(0,1), pady=(0,1), sticky="nsew")
        tk.Label(table_border, text="商品コード", bg="#F1F5F9", fg="#475569", font=font_header, pady=2).grid(row=0, column=2, pady=(0,1), sticky="nsew")

        tk.Label(table_border, text=" 49 ", bg="#F3F4F6", fg="#1F2937", font=font_val, pady=2).grid(row=1, column=0, padx=(0,1), sticky="nsew")
        tk.Label(table_border, text=" 05343 ", bg="#F3F4F6", fg="#1F2937", font=font_val, pady=2).grid(row=1, column=1, padx=(0,1), sticky="nsew")
        
        entry_container = tk.Frame(table_border, bg="#FFFFFF")
        entry_container.grid(row=1, column=2, sticky="nsew")
        
        self.var_suffix = tk.StringVar()
        vcmd = (self.container.register(self._validate_input), '%P')
        self.entry_suffix = tk.Entry(entry_container, textvariable=self.var_suffix, width=6, font=font_val, relief="flat", justify="center", fg="#1F2937", highlightthickness=0, validate="key", validatecommand=vcmd)
        self.entry_suffix.pack(fill=tk.BOTH, expand=True, pady=2)
        self.var_suffix.trace_add("write", self._on_type)

        cd_border = tk.Frame(self.container, bg="#D1D5DB", padx=1, pady=1)
        cd_border.pack(side=tk.LEFT, anchor=tk.S, padx=(0, 0))
        
        self.var_check = tk.StringVar(value="")
        self.lbl_check = tk.Label(cd_border, textvariable=self.var_check, font=font_val, fg="#1F2937", bg="#F3F4F6", width=2, pady=2)
        self.lbl_check.pack()

        self.lbl_status = tk.Label(self.container, text="×", font=("Arial", 14, "bold"), fg="#DC2626", bg="#F8F9FA")
        self.lbl_status.pack(side=tk.LEFT, anchor=tk.S, padx=(4, 0), pady=(0, 4))

        tk.Label(self.container, text="→", font=("Arial", 14), fg="#9CA3AF", bg="#F8F9FA").pack(side=tk.LEFT, anchor=tk.S, padx=4, pady=(0, 4))

        res_border = tk.Frame(self.container, bg="#D1D5DB", padx=1, pady=1)
        res_border.pack(side=tk.LEFT, anchor=tk.S)

        res_container = tk.Frame(res_border, bg="#F3F4F6")
        res_container.pack()
        self.var_result = tk.StringVar()
        tk.Entry(res_container, textvariable=self.var_result, width=14, font=font_val, 
                 state="readonly", readonlybackground="#F3F4F6", relief="flat", highlightthickness=0).pack(padx=4, pady=2)

        RoundedButton(self.container, text="コピー", command=self._copy, width=60, height=26, radius=6).pack(side=tk.LEFT, anchor=tk.S, padx=(8, 4), pady=(0, 2))
        RoundedButton(self.container, text="削除", command=lambda: self.remove_callback(self), width=50, height=26, radius=6, 
                      bg_color="#FEE2E2", fg_color="#DC2626", hover_bg="#FCA5A5").pack(side=tk.LEFT, anchor=tk.S, pady=(0, 2))

        self._on_type()

    def _validate_input(self, new_value):
        if new_value == "":
            return True
        if new_value.isdigit() and len(new_value) <= 5:
            return True
        return False

    def _on_type(self, *args):
        suffix = self.var_suffix.get()
        if len(suffix) == 5:
            code_12 = "4905343" + suffix
            cd = self.generator.calculate_check_digit(code_12)
            self.current_jan = f"{code_12}{cd}"
            self.var_result.set(self.current_jan)
            self.var_check.set(str(cd))
            self.lbl_status.config(text="○", fg="#DC2626")
        else:
            self.current_jan = ""
            self.var_result.set("")
            self.var_check.set("") 
            self.lbl_status.config(text="×", fg="#DC2626")

    def _copy(self):
        if self.current_jan:
            self.container.clipboard_clear()
            self.container.clipboard_append(self.current_jan)

    def destroy(self):
        self.container.destroy()

# ==========================================
# メインプログラム
# ==========================================
class FlatJANApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("JAN コード作成ツール")
        self.minsize(680, 560)
        self.configure(bg="#F8F9FA")
        
        self.generator = JANCodeGenerator()
        self.rows = []

        self._build_ui()
        self._add_row()

        self._set_window_icon()
        self._force_foreground() 
        self.after(200, self._close_splash) 
        self.after(500, self._force_foreground)

    def _force_foreground(self):
        if platform.system() == "Darwin":
            try:
                os.system(f'''/usr/bin/osascript -e 'tell app "System Events" to set frontmost of every process whose unix id is {os.getpid()} to true' ''')
            except Exception:
                pass
        
        self.lift()
        self.attributes('-topmost', True)
        self.after(50, lambda: self.attributes('-topmost', False))
        self.focus_force()

    def _set_window_icon(self):
        try:
            is_mac = platform.system() == "Darwin"
            icon_name = "JAN_ICON.png" if is_mac else "JAN_ICON.ico"
            
            if hasattr(sys, '_MEIPASS'):
                icon_path = os.path.join(sys._MEIPASS, icon_name)
            else:
                icon_path = os.path.join(os.path.abspath("."), icon_name)

            if os.path.exists(icon_path):
                if is_mac:
                    img = tk.PhotoImage(file=icon_path)
                    self.iconphoto(True, img)
                else:
                    self.iconbitmap(default=icon_path)
        except Exception:
            pass

    def _close_splash(self):
        try:
            import pyi_splash
            pyi_splash.close()
        except ImportError:
            pass

    def _build_ui(self):
        main_frame = tk.Frame(self, bg="#F8F9FA")
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=(20, 10))

        tk.Label(main_frame, text="①商品コードを入力", font=(SYS_FONT, 14, "bold"), bg="#F8F9FA", fg="#1F2937").pack(anchor="w")
        tk.Label(main_frame, text="商品コードを枠内に入力し、コードが自動生成します。", font=(SYS_FONT, 9), bg="#F8F9FA", fg="#4B5563").pack(anchor="w", pady=(2, 2))
        
        tk.Label(main_frame, text="※ 入力枠には数字のみ入力可能です（最大5桁）。", font=(SYS_FONT, 9, "bold"), bg="#F8F9FA", fg="#DC2626").pack(anchor="w", pady=(0, 10))

        list_container = tk.Frame(main_frame, bg="#F8F9FA")
        list_container.pack(fill=tk.X)
        list_container.grid_rowconfigure(0, weight=1)
        list_container.grid_columnconfigure(0, weight=1) 

        self.canvas = tk.Canvas(list_container, borderwidth=0, highlightthickness=0, bg="#F8F9FA", height=62)
        self.scrollbar = AutoScrollbar(list_container, orient=tk.VERTICAL)
        self.scrollable_frame = tk.Frame(self.canvas, bg="#F8F9FA")

        self.canvas_window = self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        
        self.scrollable_frame.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.bind('<Configure>', lambda e: self.canvas.itemconfig(self.canvas_window, width=e.width))

        self.canvas.grid(row=0, column=0, sticky="nsew")
        self.scrollbar.grid(row=0, column=1, sticky="ns")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        self.scrollbar.config(command=self.canvas.yview)

        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)
        self.canvas.bind_all("<Button-4>", self._on_mousewheel)
        self.canvas.bind_all("<Button-5>", self._on_mousewheel)

        btn_action_frame = tk.Frame(main_frame, bg="#F8F9FA")
        btn_action_frame.pack(fill=tk.X, pady=15) 
        RoundedButton(btn_action_frame, text="＋", command=self._add_row, width=32, height=32, radius=16, font=("Arial", 16, "bold")).pack(side=tk.LEFT)
        # 再修正：バグを修正した必須のコードを復元
        RoundedButton(btn_action_frame, text="全てクリア", command=self._clear_all, width=100, height=32, radius=6, 
                      bg_color="#FEE2E2", fg_color="#DC2626", hover_bg="#FCA5A5").pack(side=tk.LEFT, padx=15)

        tk.Frame(main_frame, height=1, bg="#D1D5DB").pack(fill=tk.X, pady=(0, 10))

        seq_frame = tk.Frame(main_frame, bg="#F8F9FA")
        seq_frame.pack(fill=tk.X, pady=5)
        
        tk.Label(seq_frame, text="商品コード連番の一括作成", font=(SYS_FONT, 10, "bold"), bg="#F8F9FA", fg="#374151").pack(anchor="w", pady=(0, 5))
        
        input_area = tk.Frame(seq_frame, bg="#F8F9FA")
        input_area.pack(fill=tk.X)
        
        tk.Label(input_area, text="開始", font=(SYS_FONT, 10), bg="#F8F9FA").pack(side=tk.LEFT)
        self.entry_start = ttk.Entry(input_area, width=8, font=(SYS_FONT, 10))
        self.entry_start.pack(side=tk.LEFT, padx=5)
        
        tk.Label(input_area, text="〜終了", font=(SYS_FONT, 10), bg="#F8F9FA").pack(side=tk.LEFT)
        self.entry_end = ttk.Entry(input_area, width=8, font=(SYS_FONT, 10))
        self.entry_end.pack(side=tk.LEFT, padx=5)
        
        RoundedButton(input_area, text="一括作成", command=self._batch_add, width=80, height=28, radius=6).pack(side=tk.LEFT, padx=10)

        tk.Label(main_frame, text="②出力設定", font=(SYS_FONT, 14, "bold"), bg="#F8F9FA", fg="#1F2937").pack(anchor="w", pady=(20, 5))
        
        settings_card = tk.Frame(main_frame, bg="#F3F4F6", highlightthickness=1, highlightbackground="#D1D5DB")
        settings_card.pack(fill=tk.X, pady=(5, 0))

        inner_pad = tk.Frame(settings_card, bg="#F3F4F6")
        inner_pad.pack(fill=tk.BOTH, expand=True, padx=25, pady=20)
        inner_pad.columnconfigure(0, weight=1)
        inner_pad.columnconfigure(1, weight=1)

        self.var_png = tk.BooleanVar(value=True)
        self.var_eps = tk.BooleanVar(value=True)
        self.var_std = tk.BooleanVar(value=True)
        self.var_15x = tk.BooleanVar(value=False)
        self.var_frame = tk.BooleanVar(value=False)

        self.var_15x.trace_add("write", self._update_frame_state)

        CustomCheckbutton(inner_pad, text="PNG 形式で出力", variable=self.var_png).grid(row=0, column=0, sticky="w", pady=(0, 16))
        CustomCheckbutton(inner_pad, text="EPS 形式で出力", variable=self.var_eps).grid(row=0, column=1, sticky="w", pady=(0, 16))
        
        CustomCheckbutton(inner_pad, text="標準サイズで出力", variable=self.var_std).grid(row=1, column=0, sticky="w", pady=(0, 16))
        
        CustomCheckbutton(inner_pad, text="1.5 倍サイズで出力", variable=self.var_15x).grid(row=2, column=0, sticky="w")
        self.chk_frame = CustomCheckbutton(inner_pad, text="黒枠を追加", variable=self.var_frame)
        self.chk_frame.grid(row=2, column=1, sticky="w")
        self._update_frame_state()

        btn_frame = tk.Frame(self, bg="#F8F9FA")
        btn_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=(10, 10))
        
        self.btn_export = RoundedButton(btn_frame, text="出力", command=self._export, width=280, height=44, radius=8, 
                                        bg_color="#3B82F6", fg_color="#FFFFFF", hover_bg="#2563EB", font=(SYS_FONT, 14, "bold"))
        self.btn_export.pack()

    def _update_canvas_height(self):
        row_h = 70
        display_rows = min(len(self.rows), 3.5) 
        self.canvas.config(height=int(display_rows * row_h))

    def _on_mousewheel(self, event):
        try:
            lo, hi = self.scrollbar.get()
            if float(lo) <= 0.0 and float(hi) >= 1.0:
                return
        except ValueError:
            return

        if platform.system() == 'Windows':
            self.canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        elif platform.system() == 'Darwin':
            self.canvas.yview_scroll(int(-1*event.delta), "units")
        else:
            if event.num == 4:
                self.canvas.yview_scroll(-1, "units")
            elif event.num == 5:
                self.canvas.yview_scroll(1, "units")

    def _add_row(self):
        row = JanInputRow(self.scrollable_frame, self.generator, self._remove_row)
        self.rows.append(row)
        self._update_canvas_height()
        
        self.scrollable_frame.update_idletasks() 
        self.canvas.yview_moveto(1)

    def _remove_row(self, row_obj):
        row_obj.destroy()
        self.rows.remove(row_obj)
        self.scrollable_frame.update_idletasks() 
        
        if len(self.rows) == 0:
            self.canvas.yview_moveto(0)
            self._add_row()
        else:
            self._update_canvas_height()

    def _clear_all(self):
        if not messagebox.askyesno("確認", "すべての入力をクリアしますか？"): 
            return
        for row in self.rows:
            row.destroy()
        self.rows.clear()
        
        # 修正：削除後の高さを正確に反映するため再描画してスクロールをリセット
        self.scrollable_frame.update_idletasks() 
        self.canvas.yview_moveto(0) 
        self._add_row()

    def _batch_add(self):
        start = self.entry_start.get().strip()
        end = self.entry_end.get().strip()

        if not (start.isdigit() and len(start) == 5 and end.isdigit() and len(end) == 5):
            return messagebox.showwarning("エラー", "開始と終了は5桁の数字で入力してください。")

        s_int, e_int = int(start), int(end)
        if s_int > e_int:
            return messagebox.showwarning("エラー", "終了番号は開始番号より大きくしてください。")

        is_first_empty = len(self.rows) == 1 and not self.rows[0].var_suffix.get()

        for i, num in enumerate(range(s_int, e_int + 1)):
            suffix = f"{num:05d}"
            if i == 0 and is_first_empty:
                self.rows[0].var_suffix.set(suffix)
            else:
                row = JanInputRow(self.scrollable_frame, self.generator, self._remove_row)
                row.var_suffix.set(suffix)
                self.rows.append(row)
            
        self._update_canvas_height()
        
        self.scrollable_frame.update_idletasks()
        self.canvas.yview_moveto(1)
        
        self.entry_start.delete(0, tk.END)
        self.entry_end.delete(0, tk.END)

    def _update_frame_state(self, *args):
        if self.var_15x.get():
            self.chk_frame.config_state(tk.NORMAL)
        else:
            self.var_frame.set(False)
            self.chk_frame.config_state(tk.DISABLED)

    def _export(self):
        valid_jans = [r.current_jan for r in self.rows if r.current_jan]
        if not valid_jans:
            return messagebox.showwarning("警告", "出力する有効な JAN Code がありません。")
            
        if not self.var_png.get() and not self.var_eps.get():
            return messagebox.showwarning("警告", "PNG または EPS のどちらかを選択してください。")
            
        if not self.var_std.get() and not self.var_15x.get():
            return messagebox.showwarning("警告", "出力サイズ（標準 / 1.5倍）を選択してください。")

        dir_path = filedialog.askdirectory(title="保存先フォルダを選択")
        if not dir_path: return

        self.btn_export.set_state(tk.DISABLED)
        self.btn_export.config(text="出力中...")

        export_thread = threading.Thread(target=self._export_task, args=(valid_jans, dir_path))
        export_thread.daemon = True
        export_thread.start()

    def _export_task(self, valid_jans, dir_path):
        count = 0
        error_message = None
        
        try:
            for jan in valid_jans:
                if self.var_std.get():
                    if self.var_png.get():
                        self.generator.draw_ean13_png(jan, os.path.join(dir_path, f"jan_{jan}_std.png"), scale=1.0)
                    if self.var_eps.get():
                        self.generator.draw_ean13_eps_vector(jan, os.path.join(dir_path, f"jan_{jan}_std.eps"), scale=1.0)
                
                if self.var_15x.get():
                    if self.var_png.get():
                        self.generator.draw_ean13_png(jan, os.path.join(dir_path, f"jan_{jan}_15x.png"), scale=1.5, add_frame=False)
                    if self.var_eps.get():
                        self.generator.draw_ean13_eps_vector(jan, os.path.join(dir_path, f"jan_{jan}_15x.eps"), scale=1.5, add_frame=False)
                    
                    if self.var_frame.get():
                        if self.var_png.get():
                            self.generator.draw_ean13_png(jan, os.path.join(dir_path, f"jan_{jan}_15x_frame.png"), scale=1.5, add_frame=True)
                        if self.var_eps.get():
                            self.generator.draw_ean13_eps_vector(jan, os.path.join(dir_path, f"jan_{jan}_15x_frame.eps"), scale=1.5, add_frame=True)
                count += 1
        except Exception as e:
            error_message = str(e)

        self.after(0, lambda: self._export_finished(count, error_message))

    def _export_finished(self, count, error_message):
        self.btn_export.set_state(tk.NORMAL)
        self.btn_export.config(text="出力")
        
        if error_message:
            messagebox.showerror("エラー", f"出力中にエラーが発生しました:\n{error_message}")
        else:
            messagebox.showinfo("完了", f"{count} 個の商品コードの画像出力が完了しました。")

if __name__ == "__main__":
    app = FlatJANApp()
    app.mainloop()