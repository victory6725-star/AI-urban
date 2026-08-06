import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
import ezdxf
from dxf_analyzer import analyze_dxf_hatch_only, get_all_layers, detect_dxf_encoding
from report_generator import export_to_excel, export_to_hwp

import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.patches import Polygon as MplPolygon

class ProgressStream:
    def __init__(self, file_obj, total_bytes, progress_callback):
        self.file_obj = file_obj
        self.total_bytes = total_bytes
        self.bytes_read = 0
        self.progress_callback = progress_callback
        self.last_percent = -1
        
    def _update_progress(self, length):
        self.bytes_read += length
        # Character count to byte size estimation for UTF-8 / CP949 Korean texts
        est_bytes = self.bytes_read * 1.15
        percent = min(99, int(est_bytes / self.total_bytes * 100))
        if percent != self.last_percent:
            self.last_percent = percent
            self.progress_callback(percent)

    def read(self, size=-1):
        data = self.file_obj.read(size)
        if data:
            self._update_progress(len(data))
        return data
        
    def readline(self, limit=-1):
        data = self.file_obj.readline(limit)
        if data:
            self._update_progress(len(data))
        return data
        
    def __iter__(self):
        return self
        
    def __next__(self):
        line = self.readline()
        if not line:
            raise StopIteration
        return line

def get_hatch_points(hatch):
    paths = []
    try:
        for path in hatch.paths:
            pts = []
            if hasattr(path, 'vertices') and path.vertices:
                # Polyline path
                for v in path.vertices:
                    pts.append((v[0], v[1]))
            elif hasattr(path, 'edges') and path.edges:
                # Edge path
                for edge in path.edges:
                    edge_type = getattr(edge, 'EDGE_TYPE', None)
                    if edge_type == 'LineEdge':
                        pts.append((edge.start[0], edge.start[1]))
                        pts.append((edge.end[0], edge.end[1]))
                    elif edge_type == 'ArcEdge':
                        pts.append((edge.start[0], edge.start[1]))
                        pts.append((edge.end[0], edge.end[1]))
                    elif edge_type == 'SplineEdge':
                        for p in getattr(edge, 'control_points', []):
                            pts.append((p[0], p[1]))
                    else:
                        if hasattr(edge, 'start'):
                            pts.append((edge.start[0], edge.start[1]))
                        if hasattr(edge, 'end'):
                            pts.append((edge.end[0], edge.end[1]))
            
            cleaned_pts = []
            for p in pts:
                if not cleaned_pts or cleaned_pts[-1] != p:
                    cleaned_pts.append(p)
            if len(cleaned_pts) >= 3:
                paths.append(cleaned_pts)
    except Exception:
        pass
    return paths

class ModernApp:
    def __init__(self, root):
        self.root = root
        self.root.title("토지이용계획 DXF 면적 분석기")
        self.root.geometry("1280x750")
        self.root.resizable(False, False)
        
        # Color Palette
        self.primary_color = "#1F4E79"
        self.secondary_color = "#F2F2F2"
        self.accent_color = "#E06666"
        self.text_color = "#333333"
        
        self.dxf_path = tk.StringVar()
        self.boundary_layer = tk.StringVar()
        self.all_layers = []
        self.hatch_layer_counts = {}  # layer_name -> hatch count
        self.layer_check_vars = {}    # layer_name -> BooleanVar
        self.layer_rename_vars = {}   # layer_name -> StringVar (editable layer name)
        self.dxf_doc = None
        
        self.setup_ui()
        
    def setup_ui(self):
        # Configure style
        style = ttk.Style()
        style.theme_use('clam')
        style.configure('TLabel', font=('맑은 고딕', 10), foreground=self.text_color)
        style.configure('TButton', font=('맑은 고딕', 10, 'bold'), background=self.primary_color, foreground='white')
        style.map('TButton', background=[('active', '#1A4063')])
        
        # Header Panel
        header_frame = tk.Frame(self.root, bg=self.primary_color, height=85)
        header_frame.pack(fill=tk.X)
        header_frame.pack_propagate(False)
        
        title_label = tk.Label(header_frame, text="토지이용계획 DXF 면적 분석기", 
                                font=("맑은 고딕", 16, "bold"), fg="white", bg=self.primary_color)
        title_label.pack(pady=10)
        subtitle_label = tk.Label(header_frame, text="DXF 해치(HATCH)를 분석하여 레이어별 면적 산출 및 엑셀/한글 보고서 자동 생성", 
                                  font=("맑은 고딕", 9), fg="#D9D9D9", bg=self.primary_color)
        subtitle_label.pack()

        # Content Panel
        content_frame = tk.Frame(self.root, padx=20, pady=15)
        content_frame.pack(fill=tk.BOTH, expand=True)
        
        # Left Side Container
        left_container = tk.Frame(content_frame)
        left_container.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Right Side Container
        right_container = tk.Frame(content_frame, padx=15)
        right_container.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)
        
        # File Selection Row
        file_label = ttk.Label(left_container, text="DXF 파일 선택:")
        file_label.grid(row=0, column=0, sticky=tk.W, pady=3)
        
        self.file_entry = ttk.Entry(left_container, textvariable=self.dxf_path, width=54)
        self.file_entry.grid(row=1, column=0, columnspan=2, sticky=tk.W, pady=3)
        
        self.btn_browse = ttk.Button(left_container, text="찾아보기...", command=self.browse_dxf)
        self.btn_browse.grid(row=1, column=2, padx=10, pady=3)
        
        # Boundary Layer Selection Row
        layer_label = ttk.Label(left_container, text="구역계 (경계선) 레이어 지정:")
        layer_label.grid(row=2, column=0, sticky=tk.W, pady=10)
        
        self.layer_combo = ttk.Combobox(left_container, textvariable=self.boundary_layer, width=51, state="readonly")
        self.layer_combo.grid(row=3, column=0, columnspan=2, sticky=tk.W, pady=3)
        self.layer_combo.bind("<<ComboboxSelected>>", lambda e: self.draw_preview())
        
        # ============================================================
        # Layer Editor Section (해치 레이어 선택/편집)
        # ============================================================
        layer_editor_label = ttk.Label(left_container, text="분석 대상 레이어 선택 및 이름 수정 (해치 HATCH만 분석):", 
                                        font=('맑은 고딕', 10, 'bold'))
        layer_editor_label.grid(row=4, column=0, columnspan=3, sticky=tk.W, pady=(15, 3))
        
        # Layer list frame with scrollbar
        layer_list_frame = tk.Frame(left_container, relief=tk.SUNKEN, bd=1)
        layer_list_frame.grid(row=5, column=0, columnspan=3, sticky=tk.NSEW, pady=3)
        
        # Canvas + Scrollbar for checkbutton list
        self.layer_canvas = tk.Canvas(layer_list_frame, height=220, bg="white", highlightthickness=0)
        layer_scrollbar = ttk.Scrollbar(layer_list_frame, orient=tk.VERTICAL, command=self.layer_canvas.yview)
        self.layer_inner_frame = tk.Frame(self.layer_canvas, bg="white")
        
        self.layer_inner_frame.bind("<Configure>", lambda e: self.layer_canvas.configure(scrollregion=self.layer_canvas.bbox("all")))
        self.layer_canvas.create_window((0, 0), window=self.layer_inner_frame, anchor=tk.NW)
        self.layer_canvas.configure(yscrollcommand=layer_scrollbar.set)
        
        self.layer_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        layer_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Mouse wheel scrolling
        self.layer_canvas.bind("<Enter>", lambda e: self.layer_canvas.bind_all("<MouseWheel>", self._on_mousewheel))
        self.layer_canvas.bind("<Leave>", lambda e: self.layer_canvas.unbind_all("<MouseWheel>"))
        
        # Placeholder text
        self.layer_placeholder = tk.Label(self.layer_inner_frame, text="DXF 파일을 선택하면 레이어 목록이 표시됩니다.",
                                          font=('맑은 고딕', 9), fg="#AAAAAA", bg="white")
        self.layer_placeholder.pack(pady=30, padx=20)
        
        # Select All / Deselect All buttons
        btn_frame = tk.Frame(left_container)
        btn_frame.grid(row=6, column=0, columnspan=3, sticky=tk.W, pady=3)
        
        self.btn_select_all = ttk.Button(btn_frame, text="전체 선택", command=self.select_all_layers, width=11)
        self.btn_select_all.pack(side=tk.LEFT, padx=(0, 4))
        
        self.btn_deselect_all = ttk.Button(btn_frame, text="전체 해제", command=self.deselect_all_layers, width=11)
        self.btn_deselect_all.pack(side=tk.LEFT, padx=(0, 4))
        
        self.btn_select_hatch = ttk.Button(btn_frame, text="해치 레이어만 선택", command=self.select_hatch_only_layers, width=16)
        self.btn_select_hatch.pack(side=tk.LEFT, padx=(0, 4))
        
        self.btn_save_all = ttk.Button(btn_frame, text="전체 레이어 저장", command=lambda: self.save_modified_dxf('all'), width=16)
        self.btn_save_all.pack(side=tk.LEFT, padx=(0, 4))
        
        self.btn_save_selected = ttk.Button(btn_frame, text="선택 레이어만 저장", command=lambda: self.save_modified_dxf('selected'), width=17)
        self.btn_save_selected.pack(side=tk.LEFT, padx=(0, 4))
        
        self.layer_info_label = tk.Label(btn_frame, text="", font=('맑은 고딕', 9), fg="#666666")
        self.layer_info_label.pack(side=tk.LEFT, padx=10)
        
        # Execute Row
        self.btn_run = ttk.Button(left_container, text="▶  해치(HATCH) 면적 분석 및 보고서 생성 실행", command=self.run_analysis, width=50)
        self.btn_run.grid(row=7, column=0, columnspan=3, pady=15)
        
        # Status Label
        self.status_label = tk.Label(left_container, text="대기 중...", font=('맑은 고딕', 10, 'italic'), fg="#888888")
        self.status_label.grid(row=8, column=0, columnspan=3, pady=3)
        
        # ============================================================
        # Right Side CAD Preview Section
        # ============================================================
        preview_label = tk.Label(right_container, text="도면 미리보기 (선택된 레이어)", font=('맑은 고딕', 11, 'bold'), fg=self.primary_color)
        preview_label.pack(anchor=tk.W, pady=(0, 5))
        
        self.fig, self.ax = plt.subplots(figsize=(6, 6))
        self.fig.patch.set_facecolor('white')
        self.ax.axis('off')
        self.ax.set_aspect('equal')
        
        self.canvas = FigureCanvasTkAgg(self.fig, master=right_container)
        self.canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
        
        warning_label = tk.Label(right_container, text="※ 대용량 도면의 경우 첫 로딩 시 수 초가 소요되나,\n이후 레이어 선택 변경은 즉각 반영됩니다.", font=('맑은 고딕', 8), fg="#888888", justify=tk.LEFT)
        warning_label.pack(anchor=tk.W, pady=(5, 0))

    def _on_mousewheel(self, event):
        self.layer_canvas.yview_scroll(int(-1*(event.delta/120)), "units")

    def browse_dxf(self):
        filepath = filedialog.askopenfilename(
            title="토지이용계획 DXF 파일 선택",
            filetypes=[("AutoCAD DXF Files", "*.dxf")]
        )
        if filepath:
            normalized_path = os.path.normpath(filepath)
            self.dxf_path.set(normalized_path)
            self.status_label.config(text="도면 데이터 로딩 중... 0%", fg="#1F4E79")
            self.btn_run.config(state="disabled")
            self.btn_browse.config(state="disabled")
            self.root.update_idletasks()
            
            # Start background thread to load DXF file exactly once
            threading.Thread(target=self._bg_load_dxf, args=(normalized_path,), daemon=True).start()

    def _bg_load_dxf(self, filepath):
        try:
            encoding = detect_dxf_encoding(filepath)
            total_bytes = os.path.getsize(filepath)
            
            # Open file and wrap stream to capture progress
            with open(filepath, 'r', encoding=encoding, errors='ignore') as f:
                stream = ProgressStream(f, total_bytes, self._update_load_progress)
                doc = ezdxf.read(stream)
            
            # Force 100% when finished loading
            self._update_load_progress(100)
            
            # Extract hatch counts
            hatch_counts = {}
            msp = doc.modelspace()
            for hatch in msp.query('HATCH'):
                layer_name = hatch.dxf.layer
                hatch_counts[layer_name] = hatch_counts.get(layer_name, 0) + 1
                
            all_layers = sorted(set(l.dxf.name for l in doc.layers))
            
            # Notify main thread
            self.root.after(0, lambda: self._on_dxf_loaded(doc, all_layers, hatch_counts))
            
        except Exception as e:
            self.root.after(0, lambda: self._on_dxf_load_failed(str(e)))

    def _update_load_progress(self, percent):
        self.root.after(0, lambda: self.status_label.config(text=f"도면 데이터 로딩 중... {percent}%", fg="#1F4E79"))

    def _on_dxf_loaded(self, doc, all_layers, hatch_counts):
        self.dxf_doc = doc
        self.all_layers = all_layers
        self.hatch_layer_counts = hatch_counts
        
        self.layer_combo['values'] = self.all_layers
        
        # Guess boundary layer
        guess_layer = ""
        for l in self.all_layers:
            l_lower = l.lower()
            if "구역계" in l_lower or "경계" in l_lower or "boundary" in l_lower or "outline" in l_lower:
                guess_layer = l
                break
        
        if guess_layer:
            self.boundary_layer.set(guess_layer)
        elif self.all_layers:
            self.boundary_layer.set(self.all_layers[0])
            
        self._populate_layer_checkboxes()
        self.draw_preview()
        
        self.btn_run.config(state="normal")
        self.btn_browse.config(state="normal")
        self.status_label.config(text="도면 로딩 및 시각화 성공", fg="green")

    def _on_dxf_load_failed(self, error_msg):
        self.dxf_doc = None
        self.all_layers = []
        self.hatch_layer_counts = {}
        
        self.btn_run.config(state="normal")
        self.btn_browse.config(state="normal")
        self.status_label.config(text="도면 로드 실패", fg="red")
        messagebox.showerror("오류", f"도면을 불러오는 중 오류가 발생했습니다:\n{error_msg}")

    def _populate_layer_checkboxes(self):
        """Populate the layer editor panel with checkboxes and editable fields for each layer."""
        # Clear existing
        for widget in self.layer_inner_frame.winfo_children():
            widget.destroy()
        
        self.layer_check_vars = {}
        self.layer_rename_vars = {}
        
        if not self.all_layers:
            self.layer_placeholder = tk.Label(self.layer_inner_frame, text="레이어가 없습니다.",
                                              font=('맑은 고딕', 9), fg="#AAAAAA", bg="white")
            self.layer_placeholder.pack(pady=30, padx=20)
            self._update_layer_info()
            return
        
        # Header row
        header = tk.Frame(self.layer_inner_frame, bg="#E8EEF4")
        header.pack(fill=tk.X, padx=2, pady=(2, 0))
        
        tk.Label(header, text="선택", font=('맑은 고딕', 8, 'bold'), bg="#E8EEF4", width=6).pack(side=tk.LEFT, padx=2)
        tk.Label(header, text="원본 레이어명", font=('맑은 고딕', 8, 'bold'), bg="#E8EEF4", width=22, anchor=tk.W).pack(side=tk.LEFT, padx=2)
        tk.Label(header, text="→", font=('맑은 고딕', 8, 'bold'), bg="#E8EEF4", width=2).pack(side=tk.LEFT, padx=2)
        tk.Label(header, text="수정할 레이어명", font=('맑은 고딕', 8, 'bold'), bg="#E8EEF4", width=22, anchor=tk.W).pack(side=tk.LEFT, padx=2)
        tk.Label(header, text="해치 수", font=('맑은 고딕', 8, 'bold'), bg="#E8EEF4", width=8).pack(side=tk.LEFT, padx=2)
        tk.Label(header, text="상태", font=('맑은 고딕', 8, 'bold'), bg="#E8EEF4", width=10).pack(side=tk.LEFT, padx=2)
        
        for i, layer_name in enumerate(self.all_layers):
            has_hatch = layer_name in self.hatch_layer_counts
            hatch_count = self.hatch_layer_counts.get(layer_name, 0)
            
            # Default: check only layers that have hatches
            var = tk.BooleanVar(value=has_hatch)
            self.layer_check_vars[layer_name] = var
            
            # Default editable layer name is the original layer name
            rename_var = tk.StringVar(value=layer_name)
            self.layer_rename_vars[layer_name] = rename_var
            
            row_bg = "#FFFFFF" if i % 2 == 0 else "#F8FAFB"
            row_frame = tk.Frame(self.layer_inner_frame, bg=row_bg)
            row_frame.pack(fill=tk.X, padx=2)
            
            cb = tk.Checkbutton(row_frame, variable=var, bg=row_bg, activebackground=row_bg,
                                command=self._update_layer_info)
            cb.pack(side=tk.LEFT, padx=(10, 0))
            
            name_label = tk.Label(row_frame, text=layer_name, font=('맑은 고딕', 9), 
                                  bg=row_bg, fg=self.text_color if has_hatch else "#AAAAAA",
                                  width=22, anchor=tk.W)
            name_label.pack(side=tk.LEFT, padx=2)
            
            arrow_label = tk.Label(row_frame, text="→", font=('맑은 고딕', 9), bg=row_bg, fg="#888888", width=2)
            arrow_label.pack(side=tk.LEFT, padx=2)
            
            rename_entry = ttk.Entry(row_frame, textvariable=rename_var, width=20)
            rename_entry.pack(side=tk.LEFT, padx=2)
            rename_entry.bind("<FocusOut>", lambda e: self.draw_preview())
            rename_entry.bind("<Return>", lambda e: self.draw_preview())
            
            count_label = tk.Label(row_frame, text=str(hatch_count) if has_hatch else "-", 
                                   font=('맑은 고딕', 9), bg=row_bg,
                                   fg=self.primary_color if has_hatch else "#CCCCCC",
                                   width=8)
            count_label.pack(side=tk.LEFT, padx=2)
            
            status_text = "● 해치 있음" if has_hatch else "○ 해치 없음"
            status_color = "#2E7D32" if has_hatch else "#BBBBBB"
            status_label = tk.Label(row_frame, text=status_text, font=('맑은 고딕', 8),
                                    bg=row_bg, fg=status_color, width=10)
            status_label.pack(side=tk.LEFT, padx=2)
        
        self._update_layer_info()
    
    def _update_layer_info(self):
        """Update the info label showing selected layer count and redraw preview."""
        selected = sum(1 for v in self.layer_check_vars.values() if v.get())
        total = len(self.layer_check_vars)
        hatch_total = len(self.hatch_layer_counts)
        self.layer_info_label.config(
            text=f"선택: {selected}/{total}개  |  해치: {hatch_total}개"
        )
        self.draw_preview()
    
    def select_all_layers(self):
        for var in self.layer_check_vars.values():
            var.set(True)
        self._update_layer_info()
    
    def deselect_all_layers(self):
        for var in self.layer_check_vars.values():
            var.set(False)
        self._update_layer_info()
    
    def select_hatch_only_layers(self):
        """Select only layers that contain HATCH entities."""
        for layer_name, var in self.layer_check_vars.items():
            var.set(layer_name in self.hatch_layer_counts)
        self._update_layer_info()

    def draw_preview(self):
        if not hasattr(self, 'dxf_doc') or not self.dxf_doc:
            return
            
        self.ax.clear()
        
        # Get selected layers
        selected_layers = [name for name, var in self.layer_check_vars.items() if var.get()]
        boundary = self.boundary_layer.get()
        
        msp = self.dxf_doc.modelspace()
        
        # 1. Draw boundary if selected
        if boundary:
            boundary_polys = msp.query(f'*[layer=="{boundary}"]')
            for pl in boundary_polys:
                if pl.dxftype() in ('LWPOLYLINE', 'POLYLINE'):
                    try:
                        pts = [(p[0], p[1]) for p in pl.get_points(format='xy')]
                        if len(pts) >= 2:
                            x_coords = [p[0] for p in pts]
                            y_coords = [p[1] for p in pts]
                            if pl.is_closed:
                                x_coords.append(pts[0][0])
                                y_coords.append(pts[0][1])
                            self.ax.plot(x_coords, y_coords, color='#D32F2F', linewidth=2.5, label='구역계')
                    except Exception:
                        pass
        
        # 2. Draw hatches of selected layers
        hatches = msp.query('HATCH')
        
        # Assign colors to each unique modified layer name
        distinct_colors = ['#1F77B4', '#FF7F0E', '#2CA02C', '#D62728', '#9467BD', 
                            '#8C564B', '#E377C2', '#7F7F7F', '#BCBD22', '#17BECF',
                            '#AEC7E8', '#FFBB78', '#98DF8A', '#FF9896', '#C5B0D5',
                            '#C49C94', '#F7B6D2', '#C7C7C7', '#DBDB8D', '#9EDAE5']
        
        mod_colors = {}
        color_idx = 0
        
        for hatch in hatches:
            layer_name = hatch.dxf.layer
            if layer_name in selected_layers:
                mod_name = self.layer_rename_vars[layer_name].get().strip()
                if not mod_name:
                    mod_name = layer_name
                
                if mod_name not in mod_colors:
                    mod_colors[mod_name] = distinct_colors[color_idx % len(distinct_colors)]
                    color_idx += 1
                
                color = mod_colors[mod_name]
                
                paths = get_hatch_points(hatch)
                for path in paths:
                    poly = MplPolygon(path, closed=True, facecolor=color, edgecolor='#333333', alpha=0.6, linewidth=0.5)
                    self.ax.add_patch(poly)
                    
        # 3. Draw lines and polylines on selected layers (Render other vector objects if requested)
        for layer_name in selected_layers:
            # Query all LINE, LWPOLYLINE, and POLYLINE entities on this layer
            lines_and_polys = msp.query(f'LINE LWPOLYLINE POLYLINE[layer=="{layer_name}"]')
            
            mod_name = self.layer_rename_vars[layer_name].get().strip()
            if not mod_name:
                mod_name = layer_name
            color = mod_colors.get(mod_name, '#777777')
            
            for ent in lines_and_polys:
                if ent.dxftype() == 'LINE':
                    try:
                        x = [ent.dxf.start[0], ent.dxf.end[0]]
                        y = [ent.dxf.start[1], ent.dxf.end[1]]
                        self.ax.plot(x, y, color=color, linewidth=0.8, alpha=0.7)
                    except Exception:
                        pass
                elif ent.dxftype() in ('LWPOLYLINE', 'POLYLINE'):
                    try:
                        pts = [(p[0], p[1]) for p in ent.get_points(format='xy')]
                        if len(pts) >= 2:
                            x = [p[0] for p in pts]
                            y = [p[1] for p in pts]
                            if ent.is_closed:
                                x.append(pts[0][0])
                                y.append(pts[0][1])
                            self.ax.plot(x, y, color=color, linewidth=1.0, alpha=0.8)
                    except Exception:
                        pass
        
        # Adjust appearance
        self.ax.autoscale()
        self.ax.set_aspect('equal')
        self.ax.axis('off')
        
        # Legend (English and Korean labels mapping)
        if mod_colors:
            from matplotlib.patches import Patch
            legend_elements = [Patch(facecolor=c, edgecolor='#333333', label=name, alpha=0.6) for name, c in mod_colors.items()]
            # Set font compatible with Korean text to avoid block characters in Matplotlib legend
            self.ax.legend(handles=legend_elements, loc='upper right', prop={'family': 'Malgun Gothic', 'size': 8})
            
        self.canvas.draw()

    def save_modified_dxf(self, save_mode='all'):
        """
        save_mode='all'      : 레이어 이름 변경 사항을 전체 도면에 저장
        save_mode='selected' : 체크된 레이어의 엔티티만 포함하여 저장 (선택되지 않은 레이어는 제외)
        """
        dxf = self.dxf_path.get()
        if not dxf or not hasattr(self, 'dxf_doc') or not self.dxf_doc:
            messagebox.showwarning("경고", "먼저 DXF 도면 파일을 불러와 주세요.")
            return

        # Compile layer rename dictionary
        rename_dict = {}
        for orig_name, var in self.layer_rename_vars.items():
            mod_name = var.get().strip()
            if mod_name and mod_name != orig_name:
                rename_dict[orig_name] = mod_name

        # Get selected layers
        selected_layers = {name for name, var in self.layer_check_vars.items() if var.get()}

        if save_mode == 'selected' and not selected_layers:
            messagebox.showwarning("경고", "저장할 레이어를 최소 1개 이상 체크해 주세요.")
            return

        # Confirm message differs by mode
        if save_mode == 'all':
            confirm_msg = (
                f"레이어 이름 변경 사항을 전체 도면에 저장합니다.\n\n"
                f"• 변경된 레이어: {len(rename_dict)}개\n"
                f"• 도면 전체 레이어: {len(self.all_layers)}개 모두 포함"
            )
            suffix = "_레이어수정"
            title = "전체 레이어 DXF 저장"
        else:
            confirm_msg = (
                f"체크된 레이어({len(selected_layers)}개)의 엔티티만 포함하여 저장합니다.\n\n"
                f"• 선택된 레이어: {len(selected_layers)}개\n"
                f"• 제외 레이어: {len(self.all_layers) - len(selected_layers)}개 (엔티티 제외됨)\n\n"
                f"⚠ 체크되지 않은 레이어의 모든 엔티티는 저장 파일에 포함되지 않습니다."
            )
            suffix = "_선택레이어"
            title = "선택 레이어만 DXF 저장"

        if not messagebox.askyesno("확인", confirm_msg):
            return

        base_dir = os.path.dirname(dxf)
        base_name = os.path.splitext(os.path.basename(dxf))[0]
        default_out = os.path.join(base_dir, f"{base_name}{suffix}.dxf")

        filepath = filedialog.asksaveasfilename(
            title=title,
            initialfile=os.path.basename(default_out),
            filetypes=[("AutoCAD DXF Files", "*.dxf")],
            defaultextension=".dxf"
        )

        if filepath:
            self.status_label.config(text="CAD 레이어 저장 중... 0%", fg="#1F4E79")
            self.btn_run.config(state="disabled")
            self.btn_browse.config(state="disabled")
            self.btn_save_all.config(state="disabled")
            self.btn_save_selected.config(state="disabled")
            self.root.update_idletasks()

            layers_to_keep = selected_layers if save_mode == 'selected' else None
            threading.Thread(
                target=self._bg_save_modified_dxf,
                args=(dxf, filepath, rename_dict, layers_to_keep),
                daemon=True
            ).start()

    def _bg_save_modified_dxf(self, orig_path, new_path, rename_dict, layers_to_keep=None):
        """
        layers_to_keep=None  : 전체 레이어 저장 (레이어 이름 수정만 적용)
        layers_to_keep=set() : 해당 레이어 엔티티만 포함하여 저장
        """
        try:
            # Step 1: Load original DXF with recovery mode
            self.root.after(0, lambda: self.status_label.config(text="CAD 레이어 저장 중... 10% (원본 도면 읽기)", fg="#1F4E79"))
            try:
                doc, auditor = ezdxf.recover.readfile(orig_path)
            except Exception:
                doc = ezdxf.readfile(orig_path)
            msp = doc.modelspace()

            # Step 2: Create new layer entries for renamed layers
            self.root.after(0, lambda: self.status_label.config(text="CAD 레이어 저장 중... 30% (레이어 생성)", fg="#1F4E79"))
            for old_name, new_name in rename_dict.items():
                if new_name not in [l.dxf.name for l in doc.layers]:
                    try:
                        old_layer = doc.layers.get(old_name)
                        doc.layers.new(name=new_name, dxfattribs={
                            'color': old_layer.dxf.color,
                            'linetype': old_layer.dxf.linetype
                        })
                    except Exception:
                        doc.layers.new(name=new_name)

            # Step 3: Rename entities to new layer names
            self.root.after(0, lambda: self.status_label.config(text="CAD 레이어 저장 중... 50% (엔티티 이름 수정)", fg="#1F4E79"))
            for old_name, new_name in rename_dict.items():
                entities = msp.query(f'*[layer=="{old_name}"]')
                for ent in entities:
                    ent.dxf.layer = new_name

            # Step 4 (only for 'selected' mode): delete entities NOT in kept layers
            if layers_to_keep is not None:
                self.root.after(0, lambda: self.status_label.config(text="CAD 레이어 저장 중... 65% (미선택 엔티티 제외)", fg="#1F4E79"))
                # Build effective kept layer set (apply renames)
                effective_keep = set()
                for layer in layers_to_keep:
                    effective_keep.add(rename_dict.get(layer, layer))

                # Collect entities to delete
                to_delete = [ent for ent in msp if ent.dxf.layer not in effective_keep]
                for ent in to_delete:
                    try:
                        msp.delete_entity(ent)
                    except Exception:
                        pass

            # Step 5: Save
            self.root.after(0, lambda: self.status_label.config(text="CAD 레이어 저장 중... 75% (파일 저장)", fg="#1F4E79"))
            doc.saveas(new_path)

            # Step 6: Reload
            self.root.after(0, lambda: self.status_label.config(text="CAD 레이어 저장 중... 90% (수정본 열기)", fg="#1F4E79"))
            try:
                fresh_doc, _ = ezdxf.recover.readfile(new_path)
            except Exception:
                fresh_doc = ezdxf.readfile(new_path)

            hatch_counts = {}
            for hatch in fresh_doc.modelspace().query('HATCH'):
                layer_name = hatch.dxf.layer
                hatch_counts[layer_name] = hatch_counts.get(layer_name, 0) + 1
            all_layers = sorted(set(l.dxf.name for l in fresh_doc.layers))

            self.root.after(0, lambda: self.status_label.config(text="CAD 레이어 저장 중... 100%", fg="#1F4E79"))
            self.root.after(0, lambda: self._on_dxf_saved(fresh_doc, new_path, all_layers, hatch_counts))

        except Exception as e:
            self.root.after(0, lambda: self._on_dxf_save_failed(str(e)))

    def _on_dxf_saved(self, fresh_doc, new_path, all_layers, hatch_counts):
        self.dxf_doc = fresh_doc
        self.dxf_path.set(os.path.normpath(new_path))
        self.all_layers = all_layers
        self.hatch_layer_counts = hatch_counts
        self.layer_combo['values'] = self.all_layers

        # Refresh boundary layer guess
        guess_layer = ""
        for l in self.all_layers:
            l_lower = l.lower()
            if "구역계" in l_lower or "경계" in l_lower or "boundary" in l_lower or "outline" in l_lower:
                guess_layer = l
                break
        if guess_layer:
            self.boundary_layer.set(guess_layer)

        self._populate_layer_checkboxes()
        self.draw_preview()

        self.btn_run.config(state="normal")
        self.btn_browse.config(state="normal")
        self.btn_save_all.config(state="normal")
        self.btn_save_selected.config(state="normal")
        self.status_label.config(text="CAD 파일 저장 및 로드 완료", fg="green")
        messagebox.showinfo("성공", f"수정된 CAD 도면 파일이 성공적으로 저장되었습니다:\n{new_path}")

    def _on_dxf_save_failed(self, error_msg):
        self.btn_run.config(state="normal")
        self.btn_browse.config(state="normal")
        self.btn_save_all.config(state="normal")
        self.btn_save_selected.config(state="normal")
        self.status_label.config(text="CAD 저장 실패", fg="red")
        messagebox.showerror("오류", f"CAD 저장 중 에러가 발생했습니다:\n{error_msg}")


    def run_analysis(self):
        dxf = self.dxf_path.get()
        if not dxf:
            messagebox.showwarning("경고", "DXF 파일을 선택해 주세요.")
            return
        
        # Get selected layers (original names)
        selected_layers = [name for name, var in self.layer_check_vars.items() if var.get()]
        if not selected_layers:
            messagebox.showwarning("경고", "분석할 레이어를 최소 1개 이상 선택해 주세요.")
            return
            
        boundary = self.boundary_layer.get()
            
        self.btn_run.config(state="disabled")
        self.status_label.config(text="해치(HATCH) 면적 분석 중 (용량이 클수록 1분 가량 소요될 수 있습니다)...", fg="#1F4E79")
        self.root.update_idletasks()
        
        try:
            # 1. Analyze (HATCH ONLY with selected layers)
            raw_analysis_data = analyze_dxf_hatch_only(dxf, boundary if boundary else None, set(selected_layers))
            
            # 2. Map original layer names to modified names and aggregate
            mapped_layers = {}
            for orig_name, area in raw_analysis_data.get('layers', {}).items():
                mod_name = self.layer_rename_vars[orig_name].get().strip()
                if not mod_name:
                    mod_name = orig_name
                mapped_layers[mod_name] = mapped_layers.get(mod_name, 0.0) + area
            
            # Map boundary layer name if it is selected and renamed
            mapped_boundary = boundary
            if boundary in self.layer_rename_vars:
                mapped_boundary = self.layer_rename_vars[boundary].get().strip()
                if not mapped_boundary:
                    mapped_boundary = boundary
            
            analysis_data = {
                'layers': mapped_layers,
                'boundary_area': raw_analysis_data.get('boundary_area', 1.0)
            }
            
            # 3. Output path
            base_dir = os.path.dirname(dxf)
            base_name = os.path.splitext(os.path.basename(dxf))[0]
            
            excel_out = os.path.join(base_dir, f"{base_name}_토지이용계획표.xlsx")
            hwp_out = os.path.join(base_dir, f"{base_name}_토지이용계획표.hwp")
            
            # 4. Export Excel
            self.status_label.config(text="엑셀 파일 생성 중...", fg="#1F4E79")
            self.root.update_idletasks()
            export_to_excel(analysis_data, excel_out, mapped_boundary if mapped_boundary else None)
            
            # 5. Export HWP
            hwp_error_msg = ""
            self.status_label.config(text="한글(HWP) 파일 생성 중...", fg="#1F4E79")
            self.root.update_idletasks()
            try:
                export_to_hwp(analysis_data, hwp_out, mapped_boundary if mapped_boundary else None)
                hwp_created = True
            except Exception as e:
                hwp_created = False
                hwp_error_msg = str(e)
                
            self.btn_run.config(state="normal")
            
            # Success feedback
            layer_count = len(analysis_data.get('layers', {}))
            if hwp_created:
                self.status_label.config(text=f"모든 보고서 생성 완료! (해치 {layer_count}개 레이어 분석)", fg="green")
                messagebox.showinfo("성공", 
                    f"해치(HATCH) 면적 분석 및 보고서 생성이 완료되었습니다!\n\n"
                    f"• 분석 방법: 해치(HATCH)만 면적 산출 (폴리라인 제외)\n"
                    f"• 분석 레이어: {layer_count}개\n\n"
                    f"1. 엑셀: {excel_out}\n2. 한글: {hwp_out}")
            else:
                self.status_label.config(text="엑셀 성공 / 한글 실패", fg="orange")
                messagebox.showwarning("일부 실패", 
                    f"엑셀 파일은 정상 생성되었으나 한글 보고서 생성 중 오류가 발생했습니다.\n\n"
                    f"• 분석 방법: 해치(HATCH)만 면적 산출 (폴리라인 제외)\n"
                    f"• 분석 레이어: {layer_count}개\n\n"
                    f"1. 엑셀 생성 완료: {excel_out}\n"
                    f"2. 한글 오류 내용: {hwp_error_msg}\n\n"
                    f"*한글 프로그램이 설치되어 있고 보안 모듈 등록 또는 자동제어 팝업 승인을 완료했는지 확인해보세요.")
                
        except Exception as ex:
            self.btn_run.config(state="normal")
            self.status_label.config(text="오류 발생", fg="red")
            messagebox.showerror("오류", f"분석 실행 중 에러가 발생했습니다:\n{str(ex)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = ModernApp(root)
    root.mainloop()
