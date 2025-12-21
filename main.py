"""
================================================================================
FolderStatsTool - CÔNG CỤ THỐNG KÊ THƯ MỤC & ĐỔI TÊN HÀNG LOẠT
================================================================================
Version: 1.0.0
Author: ThanhRòm
License: MIT
================================================================================

MÔ TẢ CHUNG:
------------
Ứng dụng desktop (GUI) giúp thống kê nhanh dung lượng, số file của các thư mục
và hỗ trợ đổi tên hàng loạt file/folder thông qua Excel.

FINAL SOURCE - Đây là phiên bản hoàn chỉnh, đã được tối ưu và test kỹ lưỡng.

================================================================================
THƯ VIỆN SỬ DỤNG:
================================================================================
- tkinter: Giao diện đồ họa (có sẵn trong Python)
- pandas: Xử lý dữ liệu dạng bảng
- openpyxl: Đọc/ghi file Excel (.xlsx)

================================================================================
HƯỚNG DẪN CÀI ĐẶT:
================================================================================
pip install pandas openpyxl

================================================================================
TÍNH NĂNG CHÍNH:
================================================================================

BƯỚC 1 - CORE SCRIPT:
---------------------
- Quét đệ quy, tính toán kích thước và xuất Excel

BƯỚC 2 - HIGH PERFORMANCE & ASYNC:
----------------------------------
1. Thuật toán "Bottom-Up Dynamic Programming":
   - Sử dụng os.walk(topdown=False) thay vì quét đi quét lại
   - Tính dung lượng thư mục con sâu nhất trước, cộng dồn lên cha
   - Tốc độ tăng gấp hàng trăm lần với thư mục sâu

2. Giao diện phản hồi (Responsive GUI):
   - Tách biệt luồng xử lý khỏi luồng giao diện
   - Thanh tiến trình và Log thời gian thực
   - TreeView hiển thị trực quan

3. Tương tác dữ liệu:
   - Copy vùng chọn (Ctrl+C)
   - Dán đường dẫn trực tiếp

4. Báo cáo Excel thông minh:
   - Tự động định dạng độ rộng cột

BƯỚC 3 - BATCH RENAME:
----------------------
1. Đổi tên file/folder hàng loạt
2. Import/Export Excel để đổi tên
3. Inline editing trong TreeView

FINAL SOURCE - Đây là phiên bản hoàn chỉnh v1.0.0
================================================================================
"""

# ==============================================================================
# PHẦN 0: IMPORT THƯ VIỆN
# ==============================================================================
# Các thư viện chuẩn của Python
import os                   # Thao tác với hệ thống file
import sys                  # Tương tác với Python interpreter  
import datetime             # Xử lý ngày giờ
import subprocess           # Chạy các lệnh hệ thống
import threading            # Xử lý đa luồng
import queue                # Hàng đợi an toàn cho đa luồng
import time                 # Đo thời gian
import shutil               # Di chuyển file/folder

# Thư viện giao diện đồ họa (có sẵn trong Python)
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# ==============================================================================
# PHẦN 0.1: KIỂM TRA & CÀI ĐẶT THƯ VIỆN NGOÀI
# ==============================================================================
# Tự động cài đặt pandas và openpyxl nếu chưa có
try:
    import pandas as pd
    from openpyxl import load_workbook
except ImportError:
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pandas", "openpyxl"])
        import pandas as pd
        from openpyxl import load_workbook
    except:
        messagebox.showerror("Lỗi Thư viện", "Cần cài đặt thư viện: pandas, openpyxl")
        sys.exit()

# ==============================================================================
# PHẦN 0.2: CẤU HÌNH THÔNG BÁO KHỞI ĐỘNG
# ==============================================================================
# Đặt False để tắt thông báo khi khởi động ứng dụng
SHOW_STARTUP_INFO = True

# Thông tin ứng dụng
APP_INFO = {
    "name": "FolderStatsTool v1.0.1",
    "description": "Công cụ thống kê thư mục & đổi tên hàng loạt",
    "features": [
        "GUI Tkinter với Progress Bar",
        "Thuật toán Bottom-Up siêu nhanh",
        "TreeView với inline editing",
        "ĐỔI TÊN FILE/FOLDER HÀNG LOẠT",
        "DI CHUYỂN & XÓA FILE/FOLDER",
        "TẠO MỚI FILE/FOLDER",
        "LỌc song song nhiều cột",
        "Import/Export Excel"
    ],
    "author": "ThanhRòm"
}

def show_startup_info(parent=None):
    """
    Hiển thị hộp thoại giới thiệu ứng dụng khi khởi động.
    Có thể tắt bằng cách đặt SHOW_STARTUP_INFO = False
    
    Tham số:
        parent: Cửa sổ cha để căn giữa dialog (tùy chọn)
    """
    if not SHOW_STARTUP_INFO:
        return
    
    # Tạo nội dung thông báo
    msg = f"📁 {APP_INFO['name']}\n\n"
    msg += f"📝 {APP_INFO['description']}\n\n"
    msg += "✨ Tính năng:\n"
    for f in APP_INFO['features']:
        msg += f"  • {f}\n"
    msg += f"\n👤 Tác giả: {APP_INFO['author']}"
    
    if parent:
        # Tạo dialog tùy chỉnh để căn giữa
        dialog = tk.Toplevel(parent)
        dialog.title("Giới thiệu ứng dụng")
        dialog.transient(parent)
        dialog.grab_set()
        dialog.resizable(False, False)
        
        # Nội dung
        frame = ttk.Frame(dialog, padding=20)
        frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(frame, text=msg, justify=tk.LEFT, font=("Segoe UI", 10)).pack(pady=10)
        ttk.Button(frame, text="OK", command=dialog.destroy, width=15).pack(pady=10)
        
        # Căn giữa dialog trên parent
        dialog.update_idletasks()
        dialog_width = dialog.winfo_width()
        dialog_height = dialog.winfo_height()
        x = parent.winfo_x() + (parent.winfo_width() // 2) - (dialog_width // 2)
        y = parent.winfo_y() + (parent.winfo_height() // 2) - (dialog_height // 2)
        dialog.geometry(f"+{x}+{y}")
        
        dialog.wait_window()
    else:
        messagebox.showinfo("Giới thiệu ứng dụng", msg)

# ==============================================================================
# PHẦN 1: TIỆN ÍCH HỖ TRỢ (UTILS)
# ==============================================================================

def format_size(size_bytes):
    """
    Chuyển đổi byte sang định dạng đọc được (KB, MB, GB, TB).
    
    Tham số:
        size_bytes (int): Kích thước tính bằng byte
        
    Trả về:
        str: Chuỗi định dạng "X.XX đơn vị" (ví dụ: "1.50 GB")
    """
    if size_bytes == 0: 
        return "0 B"
    
    size_name = ("B", "KB", "MB", "GB", "TB")
    i = 0
    p = 1024  # Hệ số chia
    s = size_bytes
    
    # Chia liên tục cho 1024 đến khi nhỏ hơn 1024 hoặc hết đơn vị
    while s >= p and i < len(size_name) - 1:
        s /= p
        i += 1
    
    return f"{s:.2f} {size_name[i]}"

def format_duration(seconds):
    """
    Chuyển đổi số giây thành chuỗi thời gian đọc được.
    
    Tham số:
        seconds (float): Số giây
        
    Trả về:
        str: Chuỗi thời gian (ví dụ: "1 phút 30.00 giây")
    """
    if seconds < 60:
        return f"{seconds:.2f} giây"
    else:
        m, s = divmod(seconds, 60)
        h, m = divmod(m, 60)
        if h:
            return f"{int(h)} giờ {int(m)} phút {s:.2f} giây"
        return f"{int(m)} phút {s:.2f} giây"

# ==============================================================================
# PHẦN 2: LỚP GIAO DIỆN CHÍNH (GUI CLASS)
# ==============================================================================

class FolderStatsTool(tk.Tk):
    """
    Lớp chính của ứng dụng, kế thừa từ tk.Tk.
    Quản lý toàn bộ giao diện và logic nghiệp vụ.
    """
    
    def __init__(self):
        """
        Khởi tạo ứng dụng:
        - Cấu hình cửa sổ chính
        - Khai báo biến trạng thái
        - Tạo giao diện
        - Thiết lập sự kiện
        """
        super().__init__()
        
        # Thông báo khởi động sẽ hiển thị sau khi UI tải xong (xem cuối __init__)
        
        # Cấu hình cửa sổ chính
        self.title("FolderStatsTool v1.0.1 - Thống kê & Đổi tên hàng loạt - by ThanhRòm")
        
        # Lấy kích thước màn hình và đặt cửa sổ chiếm hết chiều cao
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        window_width = min(1200, screen_width - 200)  # Tối đa 1600px rộng
        window_height = screen_height - 80  # Trừ taskbar
        x_pos = (screen_width - window_width) // 2
        y_pos = 10
        self.geometry(f"{window_width}x{window_height}+{x_pos}+{y_pos}")
        self.minsize(1100, 700)          # Kích thước tối thiểu
        
        # --- Biến trạng thái giao diện ---
        self.path_var = tk.StringVar()              # Đường dẫn thư mục quét
        self.deep_scan_var = tk.BooleanVar(value=True)  # Quét sâu hay nhanh
        self.status_var = tk.StringVar(value="Sẵn sàng")  # Trạng thái
        self.stats_info_var = tk.StringVar(value="")      # Thống kê tổng
        
        # --- Biến lọc dữ liệu ---
        self.filter_text_var = tk.StringVar()       # Text lọc theo tên
        self.filter_type_var = tk.StringVar(value="Tất cả")  # Lọc theo loại
        # Lưu thông tin các filter đang active: {"name": {...}, "type": {...}, "size": {...}, "count": {...}, "delete": {...}}
        self.active_filters = {}
        # Lưu tên gốc của các cột
        self.original_headers = {
            "#0": "Tên file/folder",
            "#1": "Loại",
            "#2": "Dung lượng",
            "#3": "Files",
            "#4": "Xóa",
            "#5": "Tên mới",
            "#6": "Di chuyển đến..."
        }
        
        # --- Biến quản lý đa luồng ---
        self.queue = queue.Queue()      # Hàng đợi giao tiếp giữa các luồng
        self.is_scanning = False        # Cờ đang quét hay không
        self.stop_event = threading.Event()  # Sự kiện dừng quét
        self.start_time = 0             # Thời điểm bắt đầu quét
        
        # --- Dữ liệu lưu trữ ---
        # data_map: Dictionary lưu thông tin mỗi file/folder
        # Key: đường dẫn đầy đủ
        # Value: dict chứa path, name, type, size, count, children, new_name
        self.data_map = {}
        self.root_path_cache = ""       # Lưu đường dẫn gốc đang quét
        self.item_id_to_path = {}       # Map ID Treeview -> Path thực tế

        # Khởi tạo giao diện
        self.style_config()
        self.create_widgets()
        self.setup_bindings()
        
        # Hiển thị thông báo khởi động sau khi UI đã tải xong
        self.after(100, lambda: show_startup_info(self))

    def style_config(self):
        """
        Cấu hình style cho các widget ttk.
        Sử dụng theme 'clam' với các tùy chỉnh phông chữ.
        """
        style = ttk.Style()
        style.theme_use('clam')
        
        # Style cho TreeView
        style.configure("Treeview", rowheight=28, font=("Segoe UI", 10))
        style.configure("Treeview.Heading", font=("Segoe UI", 10, "bold"), background="#e1e1e1")
        
        # Style cho Progress Bar
        style.configure("Horizontal.TProgressbar", thickness=15)

    def create_widgets(self):
        """
        Tạo toàn bộ các widget giao diện.
        Chia thành 5 phần: Header, Toolbar, Status, Body (TreeView), Footer.
        """
        # ========== 1. HEADER - Khu vực nhập liệu ==========
        top_frame = ttk.Frame(self, padding="15 10 15 5")
        top_frame.pack(fill=tk.X)

        # Dòng tùy chọn quét
        opt_frame = ttk.Frame(top_frame)
        opt_frame.pack(fill=tk.X, pady=(0, 5))
        ttk.Checkbutton(opt_frame, text="Quét sâu (Deep Scan)", variable=self.deep_scan_var).pack(side=tk.LEFT)
        ttk.Label(opt_frame, text="(Tính tổng dung lượng toàn bộ thư mục con)", 
                  foreground="gray", font=("Segoe UI", 9, "italic")).pack(side=tk.LEFT, padx=5)

        # Dòng nhập đường dẫn
        input_frame = ttk.Frame(top_frame)
        input_frame.pack(fill=tk.X, pady=2)
        
        ttk.Button(input_frame, text="Chọn Thư mục...", command=self.browse_folder).pack(side=tk.LEFT, padx=(0, 5))
        self.entry_path = ttk.Entry(input_frame, textvariable=self.path_var, font=("Consolas", 10))
        self.entry_path.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5, ipady=4)
        
        self.btn_scan = ttk.Button(input_frame, text="BẮT ĐẦU QUÉT", command=self.start_scan_thread)
        self.btn_scan.pack(side=tk.LEFT, padx=(5, 0))

        # ========== 2. TOOLBAR - Công cụ ==========
        toolbar_frame = ttk.LabelFrame(self, text="Công cụ", padding="10 5")
        toolbar_frame.pack(fill=tk.X, padx=15, pady=5)
        
        # --- Thu gọn/Mở rộng ---
        ttk.Button(toolbar_frame, text="📂 Mở rộng", command=self.expand_all).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar_frame, text="📁 Thu gọn", command=self.collapse_all).pack(side=tk.LEFT, padx=(2, 15))
        
        # --- Xóa lọc (nút tiện ích) ---
        ttk.Button(toolbar_frame, text="🔍 Xóa lọc", command=self.clear_all_filters).pack(side=tk.LEFT, padx=2)
        
        # --- Thao tác file/folder (từ phải sang trái: Xóa, Tạo, Đổi tên, Di chuyển) ---
        ttk.Button(toolbar_frame, text="🗑️ XÓA", command=self.execute_delete_batch).pack(side=tk.RIGHT, padx=2)
        ttk.Button(toolbar_frame, text="➕ TẠO MỚI", command=self.show_create_dialog).pack(side=tk.RIGHT, padx=2)
        ttk.Button(toolbar_frame, text="✏️ ĐỔI TÊN", command=self.execute_rename_batch).pack(side=tk.RIGHT, padx=2)
        ttk.Button(toolbar_frame, text="📦 DI CHUYỂN", command=self.execute_move_batch).pack(side=tk.RIGHT, padx=2)

        # ========== 3. STATUS BAR - Thanh trạng thái ==========
        status_frame = ttk.Frame(self, padding="15 0 15 5")
        status_frame.pack(fill=tk.X)
        
        # Progress bar dạng indeerminate (không xác định tiến độ cụ thể)
        self.progress = ttk.Progressbar(status_frame, mode="indeterminate", style="Horizontal.TProgressbar")
        self.progress.pack(fill=tk.X, pady=(5, 2))
        
        # Dòng thông tin
        info_row = ttk.Frame(status_frame)
        info_row.pack(fill=tk.X)
        ttk.Label(info_row, textvariable=self.status_var, foreground="#0066cc", 
                  font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT)
        ttk.Label(info_row, textvariable=self.stats_info_var, 
                  font=("Segoe UI", 9, "bold")).pack(side=tk.RIGHT)

        # ========== 4. BODY - TreeView hiển thị dữ liệu ==========
        body_frame = ttk.Frame(self, padding="15 0 15 0")
        body_frame.pack(fill=tk.BOTH, expand=True)

        # Định nghĩa các cột của TreeView
        columns = ("type", "size_fmt", "count", "delete", "new_name", "new_path")
        self.tree = ttk.Treeview(body_frame, columns=columns, selectmode="extended")
        
        # Cấu hình cột "Tên file/folder" (cột cây, #0)
        self.tree.heading("#0", text="Tên file/folder", anchor="w")
        self.tree.column("#0", width=300, minwidth=150, stretch=True)
        
        # Cấu hình cột "Loại"
        self.tree.heading("type", text="Loại")
        self.tree.column("type", width=80, anchor="center", stretch=False)
        
        # Cấu hình cột "Dung lượng"
        self.tree.heading("size_fmt", text="Dung lượng")
        self.tree.column("size_fmt", width=120, anchor="e", stretch=False)
        
        # Cấu hình cột "Files" (số file)
        self.tree.heading("count", text="Files")
        self.tree.column("count", width=80, anchor="center", stretch=False)

        # Cấu hình cột "Xóa" (checkbox)
        self.tree.heading("delete", text="Xóa")
        self.tree.column("delete", width=50, anchor="center", stretch=False)

        # Cấu hình cột "Tên mới" (có thể chỉnh sửa)
        self.tree.heading("new_name", text="Tên mới", anchor="w")
        self.tree.column("new_name", width=150, minwidth=100, anchor="w", stretch=True)
        
        # Cấu hình cột "Đường dẫn mới" (di chuyển)
        self.tree.heading("new_path", text="Di chuyển đến...", anchor="w")
        self.tree.column("new_path", width=300, minwidth=100, anchor="w", stretch=True)

        # Scrollbar dọc và ngang
        ysb = ttk.Scrollbar(body_frame, orient=tk.VERTICAL, command=self.tree.yview)
        xsb = ttk.Scrollbar(body_frame, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscroll=ysb.set, xscroll=xsb.set)
        
        # Sắp xếp layout
        ysb.pack(side=tk.RIGHT, fill=tk.Y)
        xsb.pack(side=tk.BOTTOM, fill=tk.X)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # ========== 5. FOOTER - Các nút chức năng ==========
        footer_frame = ttk.Frame(self, padding="15 10")
        footer_frame.pack(fill=tk.X)
        ttk.Button(footer_frame, text="📋 Copy Clipboard", command=self.copy_selection).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(footer_frame, text="📤 Xuất Excel", command=self.export_excel).pack(side=tk.RIGHT, padx=(5, 0))
        ttk.Button(footer_frame, text="📥 Nhập Excel", command=self.import_rename_excel).pack(side=tk.RIGHT, padx=(5, 0))

    def setup_bindings(self):
        """
        Thiết lập các sự kiện (event bindings).
        """
        # Ctrl+C để copy vùng chọn
        self.tree.bind("<Control-c>", self.copy_selection)
        
        # Single click cho toggle checkbox xóa
        self.tree.bind("<Button-1>", self.on_tree_single_click)
        
        # Double click để chỉnh sửa tên mới và đường dẫn
        self.tree.bind("<Double-1>", self.on_tree_double_click)
        
        # Right-click trên header để filter
        self.tree.bind("<Button-3>", self.on_header_right_click)
        
        # Bắt đầu kiểm tra queue sau 100ms
        self.after(100, self.process_queue)

    # ==========================================================================
    # PHẦN 3: LOGIC CỐT LÕI (CORE)
    # ==========================================================================

    def browse_folder(self):
        """
        Mở dialog chọn thư mục và cập nhật vào ô nhập liệu.
        """
        folder = filedialog.askdirectory()
        if folder:
            self.path_var.set(folder)

    def process_queue(self):
        """
        Xử lý các tin nhắn từ luồng worker gửi về luồng GUI.
        Chạy liên tục mỗi 100ms để cập nhật giao diện.
        """
        try:
            while True:
                msg_type, data = self.queue.get_nowait()
                
                if msg_type == "status":
                    self.status_var.set(data)
                elif msg_type == "stats_update":
                    self.stats_info_var.set(data)
                elif msg_type == "done":
                    self.on_scan_finished(data)
                elif msg_type == "error":
                    self.is_scanning = False
                    self.progress.stop()
                    self.btn_scan.config(state="normal")
                    messagebox.showerror("Lỗi", data)
                    
        except queue.Empty:
            pass
        finally:
            # Lên lịch kiểm tra tiếp
            self.after(100, self.process_queue)

    def start_scan_thread(self):
        """
        Bắt đầu quét thư mục trong luồng riêng (worker thread).
        Tránh treo giao diện khi quét thư mục lớn.
        """
        path = self.path_var.get().strip()
        
        # Kiểm tra đường dẫn hợp lệ
        if not path or not os.path.exists(path):
            messagebox.showerror("Lỗi", "Đường dẫn không hợp lệ!")
            return
        
        # Tránh quét chồng chéo
        if self.is_scanning:
            return

        # Đặt trạng thái quét
        self.is_scanning = True
        self.start_time = time.time()
        self.stop_event.clear()
        self.btn_scan.config(state="disabled")
        self.progress.start(10)  # Bắt đầu animation progress bar
        self.status_var.set("Đang quét...")
        
        # Xóa dữ liệu cũ
        self.tree.delete(*self.tree.get_children())
        self.data_map = {}
        self.item_id_to_path = {}
        self.root_path_cache = path
        
        # Tạo và chạy luồng worker
        t = threading.Thread(target=self.worker_scan_logic, args=(path, self.deep_scan_var.get()))
        t.daemon = True  # Luồng sẽ tự động dừng khi ứng dụng đóng
        t.start()

    def on_scan_finished(self, result_data):
        """
        Xử lý khi quét hoàn tất.
        Cập nhật giao diện và hiển thị dữ liệu lên TreeView.
        
        Tham số:
            result_data (dict): Dữ liệu quét từ worker thread
        """
        self.is_scanning = False
        self.progress.stop()
        self.btn_scan.config(state="normal")
        self.status_var.set(f"Hoàn tất! Thời gian: {format_duration(time.time() - self.start_time)}")
        
        self.data_map = result_data
        root_data = self.data_map.get(self.root_path_cache)
        
        if root_data:
            # Hiển thị thống kê tổng
            self.stats_info_var.set(f"Tổng: {format_size(root_data['size'])} | {root_data['count']} files")
            
            # Độ sâu hiển thị: vô hạn nếu quét sâu, 1 nếu quét nhanh
            display_depth = float('inf') if self.deep_scan_var.get() else 1
            self.populate_tree("", self.root_path_cache, 0, display_depth)

    def populate_tree(self, parent_id, current_path, depth, max_depth):
        """
        Đệ quy thêm các node vào TreeView.
        
        Tham số:
            parent_id (str): ID của node cha trong TreeView
            current_path (str): Đường dẫn thư mục/file hiện tại
            depth (int): Độ sâu hiện tại
            max_depth (int): Độ sâu tối đa
        """
        if current_path not in self.data_map:
            return
            
        node = self.data_map[current_path]
        
        # Lấy các giá trị từ data
        delete_val = node.get('delete', '')
        new_name_val = node.get('new_name', '')
        new_path_val = node.get('new_path', '')
        
        # Tạo node trong TreeView (type, size, count, delete, new_name, new_path)
        values = (node['type'], format_size(node['size']), node['count'], delete_val, new_name_val, new_path_val)
        new_id = self.tree.insert(parent_id, "end", text=node['name'], values=values, open=True)
        
        # Lưu mapping ID -> Path
        self.item_id_to_path[new_id] = current_path
        
        # Đệ quy thêm các node con nếu chưa đạt độ sâu tối đa
        if depth < max_depth:
            children = [self.data_map[p] for p in node.get('children', []) if p in self.data_map]
            # Sắp xếp: Folder trước, sau đó theo dung lượng giảm dần
            children.sort(key=lambda x: (x['type'] == 'File', -x['size']))
            for child in children:
                self.populate_tree(new_id, child['path'], depth + 1, max_depth)

    def worker_scan_logic(self, start_path, is_deep_scan):
        """
        Logic quét thư mục chạy trong worker thread.
        Sử dụng thuật toán Bottom-Up để tính dung lượng chính xác và nhanh.
        
        Tham số:
            start_path (str): Đường dẫn thư mục bắt đầu quét
            is_deep_scan (bool): True nếu quét sâu, False nếu quét nhanh
        """
        self.queue.put(("status", "Đang khởi tạo..."))
        local_data = {}
        
        try:
            cnt = 0
            
            # Sử dụng os.walk với topdown=False (Bottom-Up)
            # Nghĩa là duyệt từ thư mục sâu nhất trước, sau đó lên thư mục cha
            for root, dirs, files in os.walk(start_path, topdown=False):
                if self.stop_event.is_set():
                    break
                    
                c_size = 0      # Tổng dung lượng thư mục hiện tại
                c_files = 0     # Tổng số file trong thư mục hiện tại
                c_paths = []    # Danh sách đường dẫn con
                
                # Xử lý các file trong thư mục
                for f in files:
                    fp = os.path.join(root, f)
                    try:
                        sz = os.path.getsize(fp)
                        c_size += sz
                        c_files += 1
                        local_data[fp] = {
                            'path': fp, 
                            'name': f, 
                            'type': 'File', 
                            'size': sz, 
                            'count': 0, 
                            'children': [], 
                            'new_name': ''
                        }
                        c_paths.append(fp)
                    except:
                        pass
                
                # Cộng dồn dung lượng từ các thư mục con (đã được tính trước)
                for d in dirs:
                    dp = os.path.join(root, d)
                    if dp in local_data:
                        c_size += local_data[dp]['size']
                        c_files += local_data[dp]['count']
                        c_paths.append(dp)
                
                # Lưu thông tin thư mục hiện tại
                local_data[root] = {
                    'path': root, 
                    'name': os.path.basename(root) if root != start_path else root,
                    'type': 'Folder', 
                    'size': c_size, 
                    'count': c_files, 
                    'children': c_paths,
                    'new_name': ''
                }
                
                cnt += 1
                # Cập nhật trạng thái mỗi 50 thư mục
                if cnt % 50 == 0:
                    self.queue.put(("status", f"Đang quét: {os.path.basename(root)}"))
            
            # Gửi kết quả về luồng GUI
            self.queue.put(("done", local_data))
            
        except Exception as e:
            self.queue.put(("error", str(e)))

    # ==========================================================================
    # PHẦN 4: TÍNH NĂNG CHỈNH SỬA TRỰC TIẾP (INLINE EDITING)
    # ==========================================================================
    
    def on_tree_single_click(self, event):
        """
        Xử lý single click - toggle checkbox cột Xóa.
        """
        region = self.tree.identify("region", event.x, event.y)
        if region != "cell":
            return
        
        column = self.tree.identify_column(event.x)
        # Chỉ xử lý cột delete (#4)
        if column != "#4":
            return
        
        item_id = self.tree.identify_row(event.y)
        if not item_id:
            return
        
        # Toggle giá trị ✓
        current_values = self.tree.item(item_id, "values")
        current_delete = current_values[3] if len(current_values) > 3 else ""
        
        new_val = "" if current_delete == "✓" else "✓"
        
        # Cập nhật UI
        self.tree.set(item_id, column="delete", value=new_val)
        
        # Cập nhật Data Map
        path = self.item_id_to_path.get(item_id)
        if path and path in self.data_map:
            self.data_map[path]['delete'] = new_val
    
    def on_tree_double_click(self, event):
        """
        Xử lý sự kiện double click vào cột new_name (#5) và new_path (#6).
        Cột delete (#4) dùng single-click toggle.
        Cột new_path (#6): Hiện dialog với cả Browse và nhập trực tiếp.
        """
        # Kiểm tra vùng click
        region = self.tree.identify("region", event.x, event.y)
        if region != "cell":
            return
        
        # Chỉ edit cột #5 (new_name) và #6 (new_path)
        column = self.tree.identify_column(event.x)
        editable_columns = {"#5": "new_name", "#6": "new_path"}
        
        if column not in editable_columns:
            return
        
        column_key = editable_columns[column]

        # Lấy item được click
        item_id = self.tree.identify_row(event.y)
        if not item_id:
            return
        
        # Lấy vị trí và kích thước ô
        bbox = self.tree.bbox(item_id, column)
        if not bbox:
            return
        x, y, width, height = bbox
        
        # Lấy giá trị hiện tại
        current_values = self.tree.item(item_id, "values")
        col_index = {"#5": 4, "#6": 5}[column]
        current_val = current_values[col_index] if len(current_values) > col_index else ""

        # Cột new_path (#6): Tạo Entry với nút Browse kế bên
        if column == "#6":
            # Tạo frame chứa Entry và Button
            edit_frame = tk.Frame(self.tree)
            edit_frame.place(x=x, y=y, width=width, height=height)
            
            # Entry chiếm phần lớn không gian
            entry = tk.Entry(edit_frame, font=("Segoe UI", 9))
            entry.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            entry.insert(0, current_val)
            entry.focus()
            entry.select_range(0, tk.END)
            
            # Biến để theo dõi trạng thái
            is_browsing = [False]
            
            # Nút Browse nhỏ gọn
            def browse():
                is_browsing[0] = True
                folder = filedialog.askdirectory(title="Chọn thư mục đích")
                is_browsing[0] = False
                if folder:
                    try:
                        if entry.winfo_exists():
                            entry.delete(0, tk.END)
                            entry.insert(0, folder)
                            entry.focus()
                    except:
                        pass
            
            browse_btn = tk.Button(edit_frame, text="📁", font=("Segoe UI", 8), 
                                   command=browse, width=2, relief="flat", bg="#e0e0e0")
            browse_btn.pack(side=tk.RIGHT, fill=tk.Y)
            
            def save_edit(event=None):
                try:
                    if edit_frame.winfo_exists():
                        new_val = entry.get()
                        self.tree.set(item_id, column="new_path", value=new_val)
                        path = self.item_id_to_path.get(item_id)
                        if path and path in self.data_map:
                            self.data_map[path]['new_path'] = new_val
                        edit_frame.destroy()
                except:
                    pass
            
            def cancel_edit(event=None):
                try:
                    if edit_frame.winfo_exists():
                        edit_frame.destroy()
                except:
                    pass
            
            def on_focus_out(event=None):
                # Không đóng nếu đang browse
                if is_browsing[0]:
                    return
                # Đợi một chút để check focus có chuyển sang browse button không
                edit_frame.after(150, lambda: save_edit() if not is_browsing[0] else None)
            
            entry.bind("<Return>", save_edit)
            entry.bind("<Escape>", cancel_edit)
            entry.bind("<FocusOut>", on_focus_out)
            return

        # Cột new_name (#5): Tạo Entry đơn giản
        entry = tk.Entry(self.tree, font=("Segoe UI", 10))
        entry.place(x=x, y=y, width=width, height=height)
        entry.insert(0, current_val)
        entry.focus()
        entry.select_range(0, tk.END)

        def save_edit(event=None):
            """Lưu giá trị mới khi nhấn Enter hoặc click ra ngoài."""
            new_val = entry.get()
            # Cập nhật UI
            self.tree.set(item_id, column=column_key, value=new_val)
            # Cập nhật Data Map
            path = self.item_id_to_path.get(item_id)
            if path and path in self.data_map:
                self.data_map[path][column_key] = new_val
            entry.destroy()

        def cancel_edit(event=None):
            """Hủy chỉnh sửa khi nhấn Escape."""
            entry.destroy()

        entry.bind("<Return>", save_edit)
        entry.bind("<FocusOut>", save_edit)
        entry.bind("<Escape>", cancel_edit)
    
    def on_header_right_click(self, event):
        """
        Xử lý right-click trên header để hiện filter popup.
        Hỗ trợ lọc song song nhiều cột.
        """
        region = self.tree.identify("region", event.x, event.y)
        if region != "heading":
            return
        
        column = self.tree.identify_column(event.x)
        
        # Tạo popup menu
        menu = tk.Menu(self, tearoff=0)
        
        # Hiển thị tất cả các filter đang active
        if self.active_filters:
            menu.add_command(label="🔵 Các bộ lọc đang áp dụng:", state="disabled")
            for col, info in self.active_filters.items():
                menu.add_command(label=f"   • {info['desc']}", state="disabled")
            menu.add_separator()
            menu.add_command(label="❌ Xóa TẤT CẢ bộ lọc", command=self.clear_all_filters)
            menu.add_separator()
        
        # Map column to filter key
        column_filter_map = {
            "#0": "name",
            "#1": "type", 
            "#2": "size",
            "#3": "count",
            "#4": "delete"
        }
        
        filter_key = column_filter_map.get(column)
        
        if column == "#0":  # Tên hiện tại (name)
            menu.add_command(label="🔍 Lọc theo tên...", command=self.show_name_filter_dialog)
            if "name" in self.active_filters:
                menu.add_command(label="❌ Xóa lọc cột này", command=lambda: self.clear_column_filter("name"))
        elif column == "#2":  # size_fmt
            menu.add_command(label="🔍 Lọc: = 0 B", command=lambda: self.filter_by_size(0))
            menu.add_command(label="🔍 Lọc: > 1 KB", command=lambda: self.filter_by_size(1024))
            menu.add_command(label="🔍 Lọc: > 1 MB", command=lambda: self.filter_by_size(1024*1024))
            menu.add_command(label="🔍 Lọc: > 10 MB", command=lambda: self.filter_by_size(10*1024*1024))
            menu.add_command(label="🔍 Lọc: > 100 MB", command=lambda: self.filter_by_size(100*1024*1024))
            menu.add_command(label="🔍 Lọc: > 1 GB", command=lambda: self.filter_by_size(1024*1024*1024))
            if "size" in self.active_filters:
                menu.add_separator()
                menu.add_command(label="❌ Xóa lọc cột này", command=lambda: self.clear_column_filter("size"))
        elif column == "#1":  # type
            menu.add_command(label="📁 Chỉ Folder", command=lambda: self.filter_by_type("Folder"))
            menu.add_command(label="📄 Chỉ File", command=lambda: self.filter_by_type("File"))
            if "type" in self.active_filters:
                menu.add_separator()
                menu.add_command(label="❌ Xóa lọc cột này", command=lambda: self.clear_column_filter("type"))
        elif column == "#3":  # count
            menu.add_command(label="🔍 Lọc: = 0 files", command=lambda: self.filter_by_count(0))
            menu.add_command(label="🔍 Lọc: > 10 files", command=lambda: self.filter_by_count(10))
            menu.add_command(label="🔍 Lọc: > 100 files", command=lambda: self.filter_by_count(100))
            menu.add_command(label="🔍 Lọc: > 1000 files", command=lambda: self.filter_by_count(1000))
            if "count" in self.active_filters:
                menu.add_separator()
                menu.add_command(label="❌ Xóa lọc cột này", command=lambda: self.clear_column_filter("count"))
        elif column == "#4":  # delete
            menu.add_command(label="✓ Chọn tất cả", command=self.select_all_delete)
            menu.add_command(label="☐ Bỏ chọn tất cả", command=self.deselect_all_delete)
            menu.add_separator()
            menu.add_command(label="🔍 Chỉ hiện đã chọn", command=self.filter_selected_delete)
            if "delete" in self.active_filters:
                menu.add_separator()
                menu.add_command(label="❌ Xóa lọc cột này", command=lambda: self.clear_column_filter("delete"))
        elif column == "#6":  # new_path
            menu.add_command(label="📁 Chọn thư mục đích cho tất cả...", command=self.browse_destination_for_all)
            menu.add_command(label="🗑️ Xóa đường dẫn đích tất cả", command=self.clear_all_destinations)
        else:
            return
        
        menu.post(event.x_root, event.y_root)
    
    def show_name_filter_dialog(self):
        """Hiển thị dialog nhập text để lọc theo tên."""
        dialog = tk.Toplevel(self)
        dialog.title("Lọc theo tên")
        dialog.resizable(False, False)
        dialog.transient(self)
        dialog.grab_set()
        
        # Kích thước dialog
        dialog_width = 350
        dialog_height = 140
        
        # Căn giữa dialog theo ứng dụng chính
        self.update_idletasks()
        x = self.winfo_x() + (self.winfo_width() // 2) - (dialog_width // 2)
        y = self.winfo_y() + (self.winfo_height() // 2) - (dialog_height // 2)
        dialog.geometry(f"{dialog_width}x{dialog_height}+{x}+{y}")
        
        ttk.Label(dialog, text="Nhập text để lọc:", font=("Segoe UI", 10)).pack(pady=(20, 10))
        entry = ttk.Entry(dialog, width=35, font=("Segoe UI", 10))
        entry.pack(pady=5, padx=20)
        entry.focus()
        
        # Điền giá trị hiện tại nếu có
        if "name" in self.active_filters:
            entry.insert(0, self.active_filters["name"].get("value", ""))
        
        def apply():
            text = entry.get().strip()
            if text:
                self.active_filters["name"] = {"value": text, "desc": f"Tên chứa '{text}'"}
                self.apply_combined_filters()
            dialog.destroy()
        
        def cancel():
            dialog.destroy()
        
        btn_frame = ttk.Frame(dialog)
        btn_frame.pack(pady=15)
        tk.Button(btn_frame, text="✓ Lọc", command=apply, width=10, height=2, relief="flat", bg="#e0e0e0").pack(side=tk.LEFT, padx=20)
        tk.Button(btn_frame, text="✗ Hủy", command=cancel, width=10, height=2, relief="flat", bg="#e0e0e0").pack(side=tk.LEFT, padx=20)
        
        entry.bind("<Return>", lambda e: apply())
    
    def browse_destination_for_all(self):
        """Chọn thư mục đích cho tất cả items hiện tại."""
        folder = filedialog.askdirectory(title="Chọn thư mục đích")
        if folder:
            for item_id in self.item_id_to_path:
                self.tree.set(item_id, column="new_path", value=folder)
                path = self.item_id_to_path.get(item_id)
                if path and path in self.data_map:
                    self.data_map[path]['new_path'] = folder
            self.status_var.set(f"Đã đặt đường dẫn đích: {folder}")
    
    def clear_all_destinations(self):
        """Xóa đường dẫn đích của tất cả."""
        for item_id in self.item_id_to_path:
            self.tree.set(item_id, column="new_path", value="")
            path = self.item_id_to_path.get(item_id)
            if path and path in self.data_map:
                self.data_map[path]['new_path'] = ""
        self.status_var.set("Đã xóa tất cả đường dẫn đích.")
    
    def update_filter_indicator(self):
        """Cập nhật indicator và header hiển thị filter đang active."""
        # Map từ filter key sang column ID và column name
        filter_to_col = {
            "name": ("#0", None),  # #0 không cần column name
            "type": ("#1", "type"),
            "size": ("#2", "size_fmt"),
            "count": ("#3", "count"),
            "delete": ("#4", "delete")
        }
        
        # Map column name để reset header
        col_name_to_original = {
            "#0": "Tên file/folder",
            "type": "Loại",
            "size_fmt": "Dung lượng",
            "count": "Files",
            "delete": "Xóa",
            "new_name": "Tên mới",
            "new_path": "Di chuyển đến..."
        }
        
        # Reset tất cả header về gốc
        self.tree.heading("#0", text="Tên file/folder")
        self.tree.heading("type", text="Loại")
        self.tree.heading("size_fmt", text="Dung lượng")
        self.tree.heading("count", text="Files")
        self.tree.heading("delete", text="Xóa")
        self.tree.heading("new_name", text="Tên mới")
        self.tree.heading("new_path", text="Di chuyển đến...")
        
        # Thêm biểu tượng cho các cột đang lọc
        for filter_key in self.active_filters:
            if filter_key in filter_to_col:
                col_id, col_name = filter_to_col[filter_key]
                original_text = col_name_to_original.get(col_name if col_name else col_id, "")
                new_text = f"🔍 {original_text}"
                if col_id == "#0":
                    self.tree.heading("#0", text=new_text)
                else:
                    self.tree.heading(col_name, text=new_text)
        
        # Cập nhật status bar
        if self.active_filters:
            filter_descs = [f['desc'] for f in self.active_filters.values()]
            self.status_var.set(f"🔵 Đang lọc: {' | '.join(filter_descs)}")
        else:
            self.status_var.set("Sẵn sàng")
    
    def clear_column_filter(self, column_key):
        """Xóa filter của một cột cụ thể."""
        if column_key in self.active_filters:
            del self.active_filters[column_key]
            self.apply_combined_filters()
    
    def clear_all_filters(self):
        """Xóa tất cả các filter."""
        self.active_filters = {}
        self.filter_text_var.set("")
        self.filter_type_var.set("Tất cả")
        self.update_filter_indicator()
        
        if self.data_map:
            self.tree.delete(*self.tree.get_children())
            self.item_id_to_path = {}
            display_depth = float('inf') if self.deep_scan_var.get() else 1
            self.populate_tree("", self.root_path_cache, 0, display_depth)
        
        self.status_var.set("Đã xóa tất cả bộ lọc - Hiển thị tất cả.")
    
    def apply_combined_filters(self):
        """
        Áp dụng tất cả các filter đang active cùng lúc.
        """
        if not self.data_map:
            return
        
        self.tree.delete(*self.tree.get_children())
        self.item_id_to_path = {}
        
        display_depth = float('inf') if self.deep_scan_var.get() else 1
        self.populate_tree_combined_filter("", self.root_path_cache, 0, display_depth)
        
        self.update_filter_indicator()
    
    def populate_tree_combined_filter(self, parent_id, current_path, depth, max_depth):
        """
        Populate TreeView với tất cả các filter đang active.
        """
        if current_path not in self.data_map:
            return
        
        node = self.data_map[current_path]
        
        # Kiểm tra tất cả các filter (trừ root)
        if depth > 0:
            # Filter theo name
            if "name" in self.active_filters:
                filter_text = self.active_filters["name"]["value"].lower()
                if filter_text not in node['name'].lower():
                    # Kiểm tra xem có con nào match không
                    if not self.has_matching_child(current_path, depth, max_depth):
                        return
            
            # Filter theo type
            if "type" in self.active_filters:
                filter_type = self.active_filters["type"]["value"]
                if node['type'] != filter_type:
                    # Folder vẫn hiển nếu có con match
                    if node['type'] == 'Folder':
                        if not self.has_matching_child(current_path, depth, max_depth):
                            return
                    else:
                        return
            
            # Filter theo size
            if "size" in self.active_filters:
                min_size = self.active_filters["size"]["value"]
                is_exact = self.active_filters["size"].get("exact", False)
                if is_exact:
                    # Lọc = 0
                    if node['size'] != 0:
                        return
                else:
                    # Lọc > min_size
                    if node['size'] < min_size:
                        return
            
            # Filter theo count
            if "count" in self.active_filters:
                min_count = self.active_filters["count"]["value"]
                is_exact = self.active_filters["count"].get("exact", False)
                if is_exact:
                    # Lọc = 0 files
                    if node['type'] == 'Folder' and node['count'] != 0:
                        return
                else:
                    # Lọc > min_count
                    if node['type'] == 'Folder' and node['count'] < min_count:
                        return
            
            # Filter theo delete
            if "delete" in self.active_filters:
                if node.get('delete') != "✓":
                    return
        
        delete_val = node.get('delete', '')
        new_name_val = node.get('new_name', '')
        new_path_val = node.get('new_path', '')
        
        values = (node['type'], format_size(node['size']), node['count'], delete_val, new_name_val, new_path_val)
        new_id = self.tree.insert(parent_id, "end", text=node['name'], values=values, open=True)
        self.item_id_to_path[new_id] = current_path
        
        if depth < max_depth:
            children = [self.data_map[p] for p in node.get('children', []) if p in self.data_map]
            children.sort(key=lambda x: (x['type'] == 'File', -x['size']))
            for child in children:
                self.populate_tree_combined_filter(new_id, child['path'], depth + 1, max_depth)
    
    def has_matching_child(self, path, depth, max_depth):
        """Kiểm tra xem path có con/cháu nào match tất cả filter không."""
        if path not in self.data_map:
            return False
        node = self.data_map[path]
        
        for child_path in node.get('children', []):
            if child_path in self.data_map:
                child = self.data_map[child_path]
                
                # Kiểm tra child có match tất cả filter không
                matches = True
                
                if "name" in self.active_filters:
                    filter_text = self.active_filters["name"]["value"].lower()
                    if filter_text not in child['name'].lower():
                        matches = False
                
                if "type" in self.active_filters and matches:
                    filter_type = self.active_filters["type"]["value"]
                    if child['type'] != filter_type:
                        matches = False
                
                if "size" in self.active_filters and matches:
                    min_size = self.active_filters["size"]["value"]
                    is_exact = self.active_filters["size"].get("exact", False)
                    if is_exact:
                        if child['size'] != 0:
                            matches = False
                    else:
                        if child['size'] < min_size:
                            matches = False
                
                if "count" in self.active_filters and matches:
                    min_count = self.active_filters["count"]["value"]
                    is_exact = self.active_filters["count"].get("exact", False)
                    if is_exact:
                        if child['type'] == 'Folder' and child['count'] != 0:
                            matches = False
                    else:
                        if child['type'] == 'Folder' and child['count'] < min_count:
                            matches = False
                
                if "delete" in self.active_filters and matches:
                    if child.get('delete') != "✓":
                        matches = False
                
                if matches:
                    return True
                
                # Kiểm tra đệ quy nếu là folder
                if child['type'] == 'Folder' and depth < max_depth:
                    if self.has_matching_child(child_path, depth + 1, max_depth):
                        return True
        
        return False

    def filter_by_size(self, min_size):
        """
        Lọc chỉ hiện các items có size > min_size (hoặc = 0 nếu min_size = 0).
        """
        size_label = format_size(min_size)
        if min_size == 0:
            self.active_filters["size"] = {"value": min_size, "desc": f"Dung lượng = 0 B", "exact": True}
        else:
            self.active_filters["size"] = {"value": min_size, "desc": f"Dung lượng > {size_label}", "exact": False}
        self.apply_combined_filters()
    
    def filter_by_type(self, type_filter):
        """
        Lọc chỉ hiện Folder hoặc File.
        """
        self.active_filters["type"] = {"value": type_filter, "desc": f"Loại: {type_filter}"}
        self.apply_combined_filters()
    
    def filter_by_count(self, min_count):
        """
        Lọc chỉ hiện các folder có số file > min_count (hoặc = 0 nếu min_count = 0).
        """
        if min_count == 0:
            self.active_filters["count"] = {"value": min_count, "desc": f"Số files = 0", "exact": True}
        else:
            self.active_filters["count"] = {"value": min_count, "desc": f"Số files > {min_count}", "exact": False}
        self.apply_combined_filters()
    
    def apply_toolbar_filter(self):
        """
        Áp dụng filter từ toolbar (filter_text_var và filter_type_var).
        """
        filter_text = self.filter_text_var.get().strip()
        filter_type = self.filter_type_var.get()
        
        # Cập nhật filter name
        if filter_text:
            self.active_filters["name"] = {"value": filter_text, "desc": f"Tên chứa '{filter_text}'"}
        elif "name" in self.active_filters:
            del self.active_filters["name"]
        
        # Cập nhật filter type
        if filter_type != "Tất cả":
            self.active_filters["type"] = {"value": filter_type, "desc": f"Loại: {filter_type}"}
        elif "type" in self.active_filters:
            del self.active_filters["type"]
        
        self.apply_combined_filters()
    
    def filter_selected_delete(self):
        """
        Chỉ hiện các items đã đánh dấu xóa.
        """
        self.active_filters["delete"] = {"value": True, "desc": "Đã chọn xóa"}
        self.apply_combined_filters()
    
    def select_all_delete(self):
        """
        Đánh dấu xóa tất cả items hiện thị.
        """
        for item_id in self.item_id_to_path:
            self.tree.set(item_id, column="delete", value="✓")
            path = self.item_id_to_path.get(item_id)
            if path and path in self.data_map:
                self.data_map[path]['delete'] = "✓"
        self.status_var.set("Đã chọn tất cả để xóa.")
    
    def deselect_all_delete(self):
        """
        Bỏ đánh dấu xóa tất cả.
        """
        for item_id in self.item_id_to_path:
            self.tree.set(item_id, column="delete", value="")
            path = self.item_id_to_path.get(item_id)
            if path and path in self.data_map:
                self.data_map[path]['delete'] = ""
        self.status_var.set("Đã bỏ chọn tất cả.")
    


    # ==========================================================================
    # PHẦN 5: XUẤT/NHẬP EXCEL (EXPORT/IMPORT)
    # ==========================================================================

    def export_excel(self):
        """
        Xuất dữ liệu ra file Excel để người dùng điền cột 'Tên mới'.
        """
        if not self.data_map:
            messagebox.showwarning("Cảnh báo", "Chưa có dữ liệu để xuất.")
            return

        # Tạo tên file với timestamp
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        output_file = f"FolderStats_RenameTemplate_{timestamp}.xlsx"
        
        try:
            export_data = []
            
            def flatten(path, depth=0):
                """Hàm đệ quy duyệt cây thư mục theo thứ tự."""
                if path not in self.data_map:
                    return
                node = self.data_map[path]
                
                # Thêm indent để hiện cấu trúc cây
                display_name = ("    " * depth) + node['name']
                
                export_data.append({
                    "Cấu trúc thư mục": display_name,
                    "Tên hiện tại": node['name'],
                    "Tên mới": node.get('new_name', ''),
                    "Xóa": node.get('delete', ''),
                    "Di chuyển đến": node.get('new_path', ''),
                    "Loại": node['type'],
                    "Dung lượng": format_size(node['size']),
                    "Đường dẫn": node['path']  # Key để map lại
                })
                
                # Đệ quy các con
                children = [self.data_map[p] for p in node.get('children', []) if p in self.data_map]
                children.sort(key=lambda x: (x['type'] == 'File', -x['size']))
                for child in children:
                    flatten(child['path'], depth + 1)
            
            flatten(self.root_path_cache)
            
            # Tạo DataFrame và xuất Excel
            df = pd.DataFrame(export_data)
            with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name="Data")
                ws = writer.sheets["Data"]
                
                # Import thêm các module cần thiết cho openpyxl
                from openpyxl.worksheet.table import Table, TableStyleInfo
                from openpyxl.utils import get_column_letter
                
                # Tự động điều chỉnh độ rộng cột
                for idx, col in enumerate(df.columns):
                    col_letter = get_column_letter(idx + 1)
                    max_len = max(df[col].astype(str).map(len).max(), len(col)) + 2
                    ws.column_dimensions[col_letter].width = min(max_len, 60)
                
                # Freeze pane - Cố định dòng tiêu đề
                ws.freeze_panes = "A2"
                
                # Tính toán vùng dữ liệu
                last_col = get_column_letter(len(df.columns))
                last_row = len(df) + 1
                
                # Table format - Đóng khung table với style
                table_ref = f"A1:{last_col}{last_row}"
                table = Table(displayName="FolderStatsData", ref=table_ref)
                style = TableStyleInfo(
                    name="TableStyleMedium9",  # Màu xanh dương nhạt
                    showFirstColumn=False,
                    showLastColumn=False,
                    showRowStripes=True,
                    showColumnStripes=False
                )
                table.tableStyleInfo = style
                ws.add_table(table)
                
                # Grouping theo cấu trúc thư mục (depth)
                # Tính depth cho mỗi dòng dựa vào số dấu cách trong "Cấu trúc thư mục"
                for row_idx, row_data in enumerate(export_data, start=2):
                    display_name = row_data["Cấu trúc thư mục"]
                    # Đếm số lượng indent (4 spaces = 1 level)
                    indent_count = len(display_name) - len(display_name.lstrip())
                    depth = indent_count // 4
                    
                    if depth > 0:
                        # Group dòng này với outline level tương ứng
                        ws.row_dimensions[row_idx].outlineLevel = min(depth, 7)  # Max 7 levels
                
                # Đóng tất cả groups mặc định
                ws.sheet_properties.outlinePr.summaryBelow = False
            
            messagebox.showinfo("Thành công", f"Đã xuất file: {output_file}\n\n✨ Tính năng mới:\n• AutoFilter trên header\n• Table format với màu xen kẽ\n• Group theo cấu trúc thư mục\n• Freeze dòng tiêu đề")
            
            # Mở file Excel (chỉ Windows)
            if os.name == 'nt':
                os.startfile(output_file)
            
        except Exception as e:
            messagebox.showerror("Lỗi Xuất Excel", str(e))

    def import_rename_excel(self):
        """
        Nhập dữ liệu đổi tên từ file Excel đã được điền cột 'Tên mới'.
        """
        if not self.data_map:
            messagebox.showwarning("Chú ý", "Vui lòng quét thư mục trước khi Import dữ liệu đổi tên.")
            return

        file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx;*.xls")])
        if not file_path:
            return

        try:
            df = pd.read_excel(file_path)
            
            # Kiểm tra cột bắt buộc
            required_cols = ["Đường dẫn", "Tên mới"]
            if not all(col in df.columns for col in required_cols):
                messagebox.showerror("Lỗi Format", 
                    "File Excel cần có cột: 'Đường dẫn' và 'Tên mới'.\nHãy dùng chức năng Xuất Excel để lấy mẫu.")
                return

            count_update = 0
            count_delete = 0
            count_move = 0
            
            # Duyệt file Excel và cập nhật data_map
            for index, row in df.iterrows():
                path = str(row["Đường dẫn"]).strip()
                new_name = str(row["Tên mới"]).strip()
                delete_val = str(row.get("Xóa", "")).strip() if "Xóa" in df.columns else ""
                new_path = str(row.get("Di chuyển đến", "")).strip() if "Di chuyển đến" in df.columns else ""
                
                if path not in self.data_map:
                    continue
                
                # Cập nhật tên mới
                if new_name and new_name.lower() != 'nan':
                    self.data_map[path]['new_name'] = new_name
                    count_update += 1
                
                # Cập nhật đánh dấu xóa
                if delete_val and delete_val.lower() not in ['nan', '']:
                    self.data_map[path]['delete'] = delete_val
                    count_delete += 1
                
                # Cập nhật đường dẫn di chuyển
                if new_path and new_path.lower() != 'nan':
                    self.data_map[path]['new_path'] = new_path
                    count_move += 1

            # Refresh lại TreeView
            self.tree.delete(*self.tree.get_children())
            self.item_id_to_path = {}
            display_depth = float('inf') if self.deep_scan_var.get() else 1
            self.populate_tree("", self.root_path_cache, 0, display_depth)
            
            msg = f"Đã cập nhật từ Excel:\\n• {count_update} mục có tên mới\\n• {count_delete} mục đánh dấu xóa\\n• {count_move} mục có đường dẫn di chuyển"
            messagebox.showinfo("Đã nhập", msg)

        except Exception as e:
            messagebox.showerror("Lỗi Import", str(e))

    def execute_rename_batch(self):
        """
        Thực hiện đổi tên hàng loạt các file/folder có 'Tên mới' khác tên cũ.
        """
        # Lọc danh sách cần đổi tên
        items_to_rename = []
        for path, data in self.data_map.items():
            new_name = data.get('new_name', '').strip()
            if new_name and new_name != data['name']:
                items_to_rename.append((path, new_name))
        
        if not items_to_rename:
            messagebox.showinfo("Thông báo", "Không tìm thấy mục nào có 'Tên mới' khác tên cũ.")
            return

        # Xác nhận trước khi thực hiện
        if not messagebox.askyesno("Xác nhận", 
            f"Bạn có chắc muốn đổi tên {len(items_to_rename)} mục không?\nLưu ý: Việc này sẽ thay đổi file hệ thống thật."):
            return

        success_count = 0
        errors = []

        # QUAN TRỌNG: Sắp xếp theo độ dài đường dẫn giảm dần
        # Path dài = nằm sâu hơn = cần rename trước
        # Nếu rename folder cha trước, path folder con sẽ bị sai
        items_to_rename.sort(key=lambda x: len(x[0]), reverse=True)

        for old_path, new_name in items_to_rename:
            try:
                folder = os.path.dirname(old_path)
                new_path = os.path.join(folder, new_name)
                
                os.rename(old_path, new_path)
                success_count += 1
                
            except Exception as e:
                errors.append(f"{os.path.basename(old_path)} -> {e}")

        # Hiển thị kết quả
        msg = f"Đã đổi tên thành công: {success_count}/{len(items_to_rename)}"
        if errors:
            msg += "\n\nLỗi:\n" + "\n".join(errors[:5])
            if len(errors) > 5:
                msg += "\n..."
        
        messagebox.showinfo("Kết quả", msg)
        
        # Quét lại để cập nhật đường dẫn mới
        self.status_var.set("Đã đổi tên xong. Đang quét lại dữ liệu mới...")
        self.start_scan_thread()

    # ==========================================================================
    # PHẦN 6: TIỆN ÍCH COPY
    # ==========================================================================
    
    def copy_selection(self, event=None):
        """
        Copy các dòng đã chọn vào clipboard.
        Format: Tên cũ [TAB] Tên mới [TAB] Dung lượng
        """
        selected = self.tree.selection()
        if not selected:
            return
            
        text = ""
        for i in selected:
            vals = self.tree.item(i, "values")
            name = self.tree.item(i, "text")
            # vals = (type, size, count, new_name)
            text += f"{name}\t{vals[3]}\t{vals[1]}\n"
            
        self.clipboard_clear()
        self.clipboard_append(text)
        self.status_var.set("Đã copy vùng chọn.")

    # ==========================================================================
    # PHẦN 7: TÍNH NĂNG MỞ RỘNG - LỌC, THU GỌN, XÓA, DI CHUYỂN
    # ==========================================================================
    
    def expand_all(self):
        """Mở rộng tất cả các node trong TreeView."""
        def expand_children(item):
            self.tree.item(item, open=True)
            for child in self.tree.get_children(item):
                expand_children(child)
        
        for item in self.tree.get_children():
            expand_children(item)
        self.status_var.set("Đã mở rộng tất cả.")
    
    def collapse_all(self):
        """Thu gọn tất cả các node trong TreeView."""
        def collapse_children(item):
            for child in self.tree.get_children(item):
                collapse_children(child)
            self.tree.item(item, open=False)
        
        for item in self.tree.get_children():
            collapse_children(item)
        self.status_var.set("Đã thu gọn tất cả.")
    
    def apply_filter(self):
        """Lọc TreeView theo tên và loại."""
        if not self.data_map:
            return
        
        filter_text = self.filter_text_var.get().lower().strip()
        filter_type = self.filter_type_var.get()
        
        # Xóa tree hiện tại và rebuild với filter
        self.tree.delete(*self.tree.get_children())
        self.item_id_to_path = {}
        
        display_depth = float('inf') if self.deep_scan_var.get() else 1
        self.populate_tree_filtered("", self.root_path_cache, 0, display_depth, filter_text, filter_type)
        
        count = len(self.tree.get_children())
        self.status_var.set(f"Đã lọc: {filter_text if filter_text else 'Tất cả'} | Loại: {filter_type}")
    
    def populate_tree_filtered(self, parent_id, current_path, depth, max_depth, filter_text, filter_type):
        """Populate TreeView với filter."""
        if current_path not in self.data_map:
            return
        
        node = self.data_map[current_path]
        
        # Kiểm tra filter
        name_match = not filter_text or filter_text in node['name'].lower()
        type_match = filter_type == "Tất cả" or filter_type == node['type']
        
        # Kiểm tra con cháu có match không
        has_matching_children = False
        if depth < max_depth:
            for child_path in node.get('children', []):
                if self.check_descendants_match(child_path, filter_text, filter_type, depth + 1, max_depth):
                    has_matching_children = True
                    break
        
        # Hiển thị node nếu match hoặc có con match
        if name_match and type_match or has_matching_children or depth == 0:
            delete_val = node.get('delete', '')
            new_name_val = node.get('new_name', '')
            new_path_val = node.get('new_path', '')
            
            values = (node['type'], format_size(node['size']), node['count'], delete_val, new_name_val, new_path_val)
            new_id = self.tree.insert(parent_id, "end", text=node['name'], values=values, open=True)
            self.item_id_to_path[new_id] = current_path
            
            # Đệ quy con
            if depth < max_depth:
                children = [self.data_map[p] for p in node.get('children', []) if p in self.data_map]
                children.sort(key=lambda x: (x['type'] == 'File', -x['size']))
                for child in children:
                    self.populate_tree_filtered(new_id, child['path'], depth + 1, max_depth, filter_text, filter_type)
    
    def check_descendants_match(self, path, filter_text, filter_type, depth, max_depth):
        """Kiểm tra xem path hoặc con cháu có match filter không."""
        if path not in self.data_map:
            return False
        node = self.data_map[path]
        
        name_match = not filter_text or filter_text in node['name'].lower()
        type_match = filter_type == "Tất cả" or filter_type == node['type']
        
        if name_match and type_match:
            return True
        
        if depth < max_depth:
            for child_path in node.get('children', []):
                if self.check_descendants_match(child_path, filter_text, filter_type, depth + 1, max_depth):
                    return True
        return False
    
    def clear_filter(self):
        """Xóa bộ lọc và hiển thị lại toàn bộ."""
        self.filter_text_var.set("")
        self.filter_type_var.set("Tất cả")
        self.active_filter = None  # Reset filter indicator
        
        if self.data_map:
            self.tree.delete(*self.tree.get_children())
            self.item_id_to_path = {}
            display_depth = float('inf') if self.deep_scan_var.get() else 1
            self.populate_tree("", self.root_path_cache, 0, display_depth)
        
        self.status_var.set("Đã xóa bộ lọc - Hiển thị tất cả.")
    
    def execute_delete_batch(self):
        """Xóa các file/folder có đánh dấu 'X' hoặc '✓' trong cột Xóa."""
        items_to_delete = []
        for path, data in self.data_map.items():
            delete_flag = str(data.get('delete', '')).strip().upper()
            if delete_flag in ['X', '✓', 'Y', 'YES', '1', 'TRUE']:
                items_to_delete.append(path)
        
        if not items_to_delete:
            messagebox.showinfo("Thông báo", "Không có mục nào được đánh dấu xóa.\nĐể xóa, click đúp vào cột 'Xóa' và nhập 'X'.")
            return
        
        # Xác nhận
        if not messagebox.askyesno("⚠️ XÁC NHẬN XÓA", 
            f"Bạn có chắc muốn XÓA VĨNH VIỄN {len(items_to_delete)} mục không?\n\n⚠️ CẢNH BÁO: Hành động này KHÔNG THỂ hoàn tác!"):
            return
        
        success_count = 0
        errors = []
        
        # Sắp xếp theo path dài trước (xóa từ sâu nhất)
        items_to_delete.sort(key=lambda x: len(x), reverse=True)
        
        for path in items_to_delete:
            try:
                if os.path.isdir(path):
                    shutil.rmtree(path)
                else:
                    os.remove(path)
                success_count += 1
            except Exception as e:
                errors.append(f"{os.path.basename(path)} -> {e}")
        
        # Hiển thị kết quả
        msg = f"Đã xóa thành công: {success_count}/{len(items_to_delete)}"
        if errors:
            msg += "\n\nLỗi:\n" + "\n".join(errors[:5])
            if len(errors) > 5:
                msg += "\n..."
        
        messagebox.showinfo("Kết quả Xóa", msg)
        
        # Quét lại
        self.status_var.set("Đã xóa xong. Đang quét lại...")
        self.start_scan_thread()
    
    def execute_move_batch(self):
        """
        Di chuyển các file/folder có đường dẫn mới.
        Tự động tạo thư mục đích nếu chưa tồn tại.
        """
        items_to_move = []
        folders_to_create = set()
        
        for path, data in self.data_map.items():
            new_path = str(data.get('new_path', '')).strip()
            if new_path:
                items_to_move.append((path, new_path))
                if not os.path.exists(new_path):
                    folders_to_create.add(new_path)
        
        if not items_to_move:
            messagebox.showinfo("Thông báo", "Không có mục nào có đường dẫn đích.\nĐể di chuyển, click đúp vào cột 'Di chuyển đến...' và nhập đường dẫn thư mục đích.")
            return
        
        # Xác nhận với thông tin về thư mục sẽ được tạo mới
        confirm_msg = f"Bạn có chắc muốn di chuyển {len(items_to_move)} mục không?"
        if folders_to_create:
            confirm_msg += f"\n\n📁 Sẽ tạo mới {len(folders_to_create)} thư mục:\n"
            for folder in list(folders_to_create)[:3]:
                confirm_msg += f"  • {folder}\n"
            if len(folders_to_create) > 3:
                confirm_msg += f"  ... và {len(folders_to_create) - 3} thư mục khác"
        
        if not messagebox.askyesno("Xác nhận Di chuyển", confirm_msg):
            return
        
        success_count = 0
        errors = []
        created_folders = []
        
        # Sắp xếp theo path dài trước (di chuyển từ sâu nhất)
        items_to_move.sort(key=lambda x: len(x[0]), reverse=True)
        
        for old_path, dest_folder in items_to_move:
            try:
                # Tự động tạo thư mục đích nếu chưa tồn tại
                if not os.path.exists(dest_folder):
                    os.makedirs(dest_folder, exist_ok=True)
                    if dest_folder not in created_folders:
                        created_folders.append(dest_folder)
                
                basename = os.path.basename(old_path)
                new_full_path = os.path.join(dest_folder, basename)
                shutil.move(old_path, new_full_path)
                success_count += 1
            except Exception as e:
                errors.append(f"{os.path.basename(old_path)} -> {e}")
        
        # Hiển thị kết quả
        msg = f"Đã di chuyển thành công: {success_count}/{len(items_to_move)}"
        if created_folders:
            msg += f"\n\n📁 Đã tạo {len(created_folders)} thư mục mới."
        if errors:
            msg += "\n\nLỗi:\n" + "\n".join(errors[:5])
            if len(errors) > 5:
                msg += "\n..."
        
        messagebox.showinfo("Kết quả Di chuyển", msg)
        
        # Quét lại
        self.status_var.set("Đã di chuyển xong. Đang quét lại...")
        self.start_scan_thread()
    
    def show_create_dialog(self):
        """
        Hiển thị dialog tạo file hoặc folder mới.
        Cho phép người dùng nhập tên và chọn loại (File/Folder).
        """
        if not self.root_path_cache:
            messagebox.showwarning("Thông báo", "Vui lòng quét thư mục trước khi tạo mới.")
            return
        
        # Xác định thư mục cha: dùng item đang chọn hoặc root
        selected = self.tree.selection()
        if selected:
            selected_id = selected[0]
            selected_path = self.item_id_to_path.get(selected_id, self.root_path_cache)
            # Nếu chọn file, lấy thư mục chứa file đó
            if os.path.isfile(selected_path):
                parent_folder = os.path.dirname(selected_path)
            else:
                parent_folder = selected_path
        else:
            parent_folder = self.root_path_cache
        
        dialog = tk.Toplevel(self)
        dialog.title("➕ Tạo File/Folder mới")
        dialog.geometry("450x220")
        dialog.resizable(False, False)
        dialog.transient(self)
        dialog.grab_set()
        
        # Header
        ttk.Label(dialog, text="➕ Tạo File/Folder mới", font=("Segoe UI", 11, "bold")).pack(pady=(15, 5))
        ttk.Label(dialog, text=f"Trong thư mục: {parent_folder}", font=("Segoe UI", 8, "italic"), 
                  foreground="gray", wraplength=400).pack(pady=(0, 10))
        
        # Nhập tên
        name_frame = ttk.Frame(dialog)
        name_frame.pack(fill=tk.X, padx=20, pady=5)
        ttk.Label(name_frame, text="Tên:", width=10).pack(side=tk.LEFT)
        name_var = tk.StringVar()
        name_entry = ttk.Entry(name_frame, textvariable=name_var, width=40)
        name_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        name_entry.focus()
        
        # Chọn loại
        type_frame = ttk.Frame(dialog)
        type_frame.pack(fill=tk.X, padx=20, pady=5)
        ttk.Label(type_frame, text="Loại:", width=10).pack(side=tk.LEFT)
        type_var = tk.StringVar(value="Folder")
        ttk.Radiobutton(type_frame, text="📁 Folder", variable=type_var, value="Folder").pack(side=tk.LEFT, padx=5)
        ttk.Radiobutton(type_frame, text="📄 File", variable=type_var, value="File").pack(side=tk.LEFT, padx=5)
        
        # Extension cho file (hiện khi chọn File)
        ext_frame = ttk.Frame(dialog)
        ext_frame.pack(fill=tk.X, padx=20, pady=5)
        ttk.Label(ext_frame, text="Phần mở rộng:", width=10).pack(side=tk.LEFT)
        ext_var = tk.StringVar(value=".txt")
        ext_combo = ttk.Combobox(ext_frame, textvariable=ext_var, width=15,
                                  values=[".txt", ".md", ".json", ".xml", ".csv", ".py", ".html", ".css", ".js", "Khác..."])
        ext_combo.pack(side=tk.LEFT, padx=5)
        ext_combo.config(state="disabled")
        
        def on_type_change(*args):
            if type_var.get() == "File":
                ext_combo.config(state="normal")
            else:
                ext_combo.config(state="disabled")
        
        type_var.trace('w', on_type_change)
        
        def create():
            name = name_var.get().strip()
            if not name:
                messagebox.showwarning("Lỗi", "Vui lòng nhập tên!")
                return
            
            item_type = type_var.get()
            
            # Nếu là file, thêm phần mở rộng nếu chưa có
            if item_type == "File":
                ext = ext_var.get().strip()
                if ext and ext != "Khác..." and not name.endswith(ext):
                    name += ext
            
            full_path = os.path.join(parent_folder, name)
            
            # Kiểm tra đã tồn tại chưa
            if os.path.exists(full_path):
                messagebox.showwarning("Lỗi", f"{item_type} '{name}' đã tồn tại!")
                return
            
            try:
                if item_type == "Folder":
                    os.makedirs(full_path)
                else:
                    # Tạo file rỗng
                    with open(full_path, 'w', encoding='utf-8') as f:
                        pass
                
                dialog.destroy()
                messagebox.showinfo("Thành công", f"Đã tạo {item_type}: {name}")
                
                # Quét lại để cập nhật TreeView
                self.start_scan_thread()
                
            except Exception as e:
                messagebox.showerror("Lỗi", f"Không thể tạo {item_type}: {e}")
        
        def cancel():
            dialog.destroy()
        
        # Buttons
        btn_frame = ttk.Frame(dialog)
        btn_frame.pack(pady=15)
        ttk.Button(btn_frame, text="✓ Tạo", command=create, width=12).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="✗ Hủy", command=cancel, width=12).pack(side=tk.LEFT, padx=5)
        
        name_entry.bind("<Return>", lambda e: create())

# ==============================================================================
# PHẦN 8: KHỞI CHẠY ỨNG DỤNG
# ==============================================================================

if __name__ == "__main__":
    app = FolderStatsTool()
    app.mainloop()

