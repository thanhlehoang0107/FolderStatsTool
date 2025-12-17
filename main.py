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
    "name": "FolderStatsTool v1.0.0",
    "description": "Công cụ thống kê thư mục & đổi tên hàng loạt",
    "features": [
        "GUI Tkinter với Progress Bar",
        "Thuật toán Bottom-Up siêu nhanh",
        "TreeView với inline editing",
        "ĐỔI TÊN FILE/FOLDER HÀNG LOẠT",
        "Import/Export Excel để đổi tên",
        "Hỗ trợ quét sâu/nhanh"
    ],
    "author": "ThanhRòm"
}

def show_startup_info():
    """
    Hiển thị hộp thoại giới thiệu ứng dụng khi khởi động.
    Có thể tắt bằng cách đặt SHOW_STARTUP_INFO = False
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
        
        # Hiển thị thông báo khởi động
        show_startup_info()
        
        # Cấu hình cửa sổ chính
        self.title("FolderStatsTool v1.0.0 - Thống kê & Đổi tên hàng loạt - by ThanhRòm")
        self.geometry("1200x750+50+50")  # Rộng x Cao + X + Y
        self.minsize(1100, 700)          # Kích thước tối thiểu
        
        # --- Biến trạng thái giao diện ---
        self.path_var = tk.StringVar()              # Đường dẫn thư mục quét
        self.deep_scan_var = tk.BooleanVar(value=True)  # Quét sâu hay nhanh
        self.status_var = tk.StringVar(value="Sẵn sàng")  # Trạng thái
        self.stats_info_var = tk.StringVar(value="")      # Thống kê tổng
        
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

        # ========== 2. TOOLBAR - Công cụ đổi tên ==========
        toolbar_frame = ttk.LabelFrame(self, text="Công cụ Đổi tên (Rename Tools)", padding="10 5")
        toolbar_frame.pack(fill=tk.X, padx=15, pady=5)
        
        ttk.Button(toolbar_frame, text="📥 Nhập từ Excel (Import)", command=self.import_rename_excel).pack(side=tk.LEFT, padx=5)
        ttk.Label(toolbar_frame, text="->", font=("Bold", 10)).pack(side=tk.LEFT)
        ttk.Label(toolbar_frame, text="Kiểm tra cột 'Tên mới' bên dưới", foreground="blue").pack(side=tk.LEFT, padx=5)
        
        self.btn_apply_rename = ttk.Button(toolbar_frame, text="⚡ THỰC HIỆN ĐỔI TÊN", command=self.execute_rename_batch)
        self.btn_apply_rename.pack(side=tk.RIGHT, padx=5)

        # ========== 3. STATUS BAR - Thanh trạng thái ==========
        status_frame = ttk.Frame(self, padding="15 0 15 5")
        status_frame.pack(fill=tk.X)
        
        # Progress bar dạng indeterminate (không xác định tiến độ cụ thể)
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
        columns = ("type", "size_fmt", "count", "new_name")
        self.tree = ttk.Treeview(body_frame, columns=columns, selectmode="extended")
        
        # Cấu hình cột "Tên hiện tại" (cột cây, #0)
        self.tree.heading("#0", text="Tên hiện tại", anchor="w")
        self.tree.column("#0", width=350, minwidth=200)
        
        # Cấu hình cột "Loại"
        self.tree.heading("type", text="Loại")
        self.tree.column("type", width=80, anchor="center")
        
        # Cấu hình cột "Dung lượng"
        self.tree.heading("size_fmt", text="Dung lượng")
        self.tree.column("size_fmt", width=120, anchor="e")
        
        # Cấu hình cột "Files" (số file)
        self.tree.heading("count", text="Files")
        self.tree.column("count", width=80, anchor="center")

        # Cấu hình cột "Tên mới" (có thể chỉnh sửa)
        self.tree.heading("new_name", text="Tên mới (Click đúp để sửa)", anchor="w")
        self.tree.column("new_name", width=300, anchor="w")

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
        ttk.Button(footer_frame, text="Copy Clipboard", command=self.copy_selection).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(footer_frame, text="Xuất Excel (Export)", command=self.export_excel).pack(side=tk.RIGHT, padx=(5, 0))

    def setup_bindings(self):
        """
        Thiết lập các sự kiện (event bindings).
        """
        # Ctrl+C để copy vùng chọn
        self.tree.bind("<Control-c>", self.copy_selection)
        
        # Double click để chỉnh sửa tên mới
        self.tree.bind("<Double-1>", self.on_tree_double_click)
        
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
        
        # Lấy tên mới nếu có (từ import hoặc edit trước)
        new_name_val = node.get('new_name', '')
        
        # Tạo node trong TreeView
        values = (node['type'], format_size(node['size']), node['count'], new_name_val)
        new_id = self.tree.insert(parent_id, "end", text=node['name'], values=values, open=(depth==0))
        
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
    
    def on_tree_double_click(self, event):
        """
        Xử lý sự kiện double click vào cột 'Tên mới'.
        Hiển thị ô nhập liệu để chỉnh sửa trực tiếp.
        """
        # Kiểm tra vùng click
        region = self.tree.identify("region", event.x, event.y)
        if region != "cell":
            return
        
        # Kiểm tra cột click (cột new_name là #4)
        column = self.tree.identify_column(event.x)
        if column != "#4":
            return

        # Lấy item được click
        item_id = self.tree.identify_row(event.y)
        if not item_id:
            return

        # Lấy vị trí và kích thước ô
        x, y, width, height = self.tree.bbox(item_id, column)
        
        # Lấy giá trị hiện tại
        current_values = self.tree.item(item_id, "values")
        current_new_name = current_values[3] if len(current_values) > 3 else ""

        # Tạo Entry widget đè lên ô
        entry = tk.Entry(self.tree, font=("Segoe UI", 10))
        entry.place(x=x, y=y, width=width, height=height)
        entry.insert(0, current_new_name)
        entry.focus()

        def save_edit(event=None):
            """Lưu giá trị mới khi nhấn Enter hoặc click ra ngoài."""
            new_val = entry.get()
            # Cập nhật UI
            self.tree.set(item_id, column="new_name", value=new_val)
            # Cập nhật Data Map
            path = self.item_id_to_path.get(item_id)
            if path and path in self.data_map:
                self.data_map[path]['new_name'] = new_val
            entry.destroy()

        def cancel_edit(event=None):
            """Hủy chỉnh sửa khi nhấn Escape."""
            entry.destroy()

        entry.bind("<Return>", save_edit)
        entry.bind("<FocusOut>", save_edit)
        entry.bind("<Escape>", cancel_edit)

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
                
                # Tự động điều chỉnh độ rộng cột
                for idx, col in enumerate(df.columns):
                    max_len = max(df[col].astype(str).map(len).max(), len(col)) + 2
                    ws.column_dimensions[chr(65 + idx)].width = min(max_len, 60)
            
            messagebox.showinfo("Thành công", f"Đã xuất file: {output_file}\nBạn hãy điền cột 'Tên mới' rồi Import lại.")
            
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
            
            # Duyệt file Excel và cập nhật data_map
            for index, row in df.iterrows():
                path = str(row["Đường dẫn"]).strip()
                new_name = str(row["Tên mới"]).strip()
                
                # Bỏ qua nếu tên mới rỗng hoặc là 'nan'
                if not new_name or new_name.lower() == 'nan':
                    continue
                
                if path in self.data_map:
                    self.data_map[path]['new_name'] = new_name
                    count_update += 1

            # Refresh lại TreeView
            self.tree.delete(*self.tree.get_children())
            self.item_id_to_path = {}
            display_depth = float('inf') if self.deep_scan_var.get() else 1
            self.populate_tree("", self.root_path_cache, 0, display_depth)
            
            messagebox.showinfo("Đã nhập", f"Đã cập nhật {count_update} mục có tên mới.\nNhấn 'THỰC HIỆN ĐỔI TÊN' để áp dụng.")

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

# ==============================================================================
# PHẦN 7: KHỞI CHẠY ỨNG DỤNG
# ==============================================================================

if __name__ == "__main__":
    app = FolderStatsTool()
    app.mainloop()
