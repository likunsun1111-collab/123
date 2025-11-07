import time
import mss
import cv2
import numpy as np
import pytesseract
import tkinter as tk
from tkinter import messagebox, filedialog, Checkbutton
import threading
import os
import winsound
import pyautogui
from PIL import Image, ImageTk
import win32api
import win32con

class AutoEnrollOCRMonitor:
    def __init__(self, root):
        self.root = root
        self.root.title("文字监控（带自动报名）")
        self.root.geometry("450x430")  # 适当加宽窗口避免控件拥挤
        self.root.resizable(True, True)
        
        # 核心参数（保持不变）
        self.monitor_area = None
        self.tesseract_path = "C:/Program Files/Tesseract-OCR/tesseract.exe"
        self.check_interval = 1.0
        self.monitoring = False
        self.is_calibrating = False
        self.color_channel = "灰"
        self.threshold_value = 229
        self.invert_binary = tk.BooleanVar(value=False)
        
        # 自动报名设置（保持不变）
        self.enroll_points = []
        self.auto_enroll = tk.BooleanVar(value=True)
        self.last_enroll_time = 0
        self.enroll_cooldown = 50
        self.click_interval = 0.06
        
        # 关键词（调整标签对应的值）
        self.enroll_keyword = "请尽快报名"
        self.success_keyword = "报名成功"  # 实际关键词值不变，仅标签文本修改
        self.victory_keyword = "全部阵亡"
        self.reward_success_keyword = "体力"  # 原"胜利结束"关键词值
        
        # 提示设置（保持不变）
        self.enroll_flash = tk.BooleanVar(value=True)
        self.enroll_sound = tk.BooleanVar(value=True)
        self.victory_flash = tk.BooleanVar(value=True)
        self.victory_sound = tk.BooleanVar(value=True)
        
        # 提示线程控制（保持不变）
        self.enroll_alert_running = False
        self.victory_alert_running = False
        self.countdown_running = False
        self.countdown_thread = None
        
        self.create_widgets()

    def create_widgets(self):
        # 1. Tesseract路径（保持不变）
        path_frame = tk.Frame(self.root)
        path_frame.pack(fill=tk.X, padx=3, pady=1)
        tk.Label(path_frame, text="Tesseract路径:", font=("微软雅黑", 8)).pack(side=tk.LEFT, padx=1)
        self.path_entry = tk.Entry(path_frame, width=20, font=("微软雅黑", 8))
        self.path_entry.pack(side=tk.LEFT, padx=1, fill=tk.X, expand=True)
        self.path_entry.insert(0, self.tesseract_path)
        tk.Button(path_frame, text="浏览", command=self.browse_tesseract, 
                 font=("微软雅黑", 8), width=4).pack(side=tk.LEFT)
        
        # 2. 通道选择+反转二值化（保持不变）
        channel_frame = tk.Frame(self.root)
        channel_frame.pack(fill=tk.X, padx=3, pady=1)
        tk.Label(channel_frame, text="颜色通道:", font=("微软雅黑", 8)).pack(side=tk.LEFT, padx=1)
        self.channel_var = tk.StringVar(value="灰")
        for ch in ["红", "绿", "蓝", "灰"]:
            tk.Radiobutton(channel_frame, text=ch, variable=self.channel_var, 
                          value=ch, command=self.update_channel, 
                          font=("微软雅黑", 8), indicatoron=0, width=3).pack(side=tk.LEFT, padx=0)
        Checkbutton(channel_frame, text="反转二值化", variable=self.invert_binary, 
                   font=("微软雅黑", 8)).pack(side=tk.RIGHT, padx=3)
        
        # 3. 阈值滑块（保持不变）
        threshold_frame = tk.Frame(self.root)
        threshold_frame.pack(fill=tk.X, padx=3, pady=1)
        tk.Label(threshold_frame, text="二值化阈值:", font=("微软雅黑", 8)).pack(side=tk.LEFT, padx=1)
        self.threshold_slider = tk.Scale(
            threshold_frame, from_=0, to=255, orient=tk.HORIZONTAL,
            length=140, command=self.update_threshold, resolution=1, showvalue=0
        )
        self.threshold_slider.set(self.threshold_value)
        self.threshold_slider.pack(side=tk.LEFT, padx=1, fill=tk.X, expand=True)
        self.threshold_val_label = tk.Label(threshold_frame, text=f"{self.threshold_value}", 
                                           font=("微软雅黑", 8), width=3)
        self.threshold_val_label.pack(side=tk.LEFT)
        
        # 4. 间隔滑块（保持不变）
        interval_frame = tk.Frame(self.root)
        interval_frame.pack(fill=tk.X, padx=3, pady=1)
        tk.Label(interval_frame, text="检测间隔(秒):", font=("微软雅黑", 8)).pack(side=tk.LEFT, padx=1)
        self.interval_slider = tk.Scale(
            interval_frame, from_=0.1, to=2.0, orient=tk.HORIZONTAL,
            length=140, command=self.update_interval, resolution=0.1, showvalue=1
        )
        self.interval_slider.set(self.check_interval)
        self.interval_slider.pack(side=tk.LEFT, padx=1, fill=tk.X, expand=True)
        self.interval_val_label = tk.Label(interval_frame, text=f"{self.check_interval:.1f}", 
                                          font=("微软雅黑", 8), width=3)
        self.interval_val_label.pack(side=tk.LEFT)
        
        # 5. 预览区域（保持不变）
        preview_frame = tk.Frame(self.root)
        preview_frame.pack(padx=3, pady=1, fill=tk.X)
        tk.Label(preview_frame, text="图像预览:", fg="blue", font=("微软雅黑", 8)).pack(anchor="w")
        self.preview_canvas = tk.Canvas(preview_frame, width=400, height=70, bg="gray")
        self.preview_canvas.pack(fill=tk.X)
        self.preview_label = tk.Label(self.preview_canvas, text="框选区域后显示预览", 
                                     bg="gray", fg="white", font=("微软雅黑", 8))
        self.preview_label.place(relx=0.5, rely=0.5, anchor="center")
        
        # 6. 关键词设置（核心调整区域）
        keyword_frame = tk.Frame(self.root)
        keyword_frame.pack(fill=tk.X, padx=3, pady=1)
        
        # 报名关键词行（移除自动报名和设置按钮）
        tk.Label(keyword_frame, text="报名关键词:", font=("微软雅黑", 8)).grid(row=0, column=0, padx=1, sticky="w")
        self.enroll_entry = tk.Entry(keyword_frame, width=12, font=("微软雅黑", 8))
        self.enroll_entry.grid(row=0, column=1, padx=1, sticky="w")
        self.enroll_entry.insert(0, self.enroll_keyword)
        Checkbutton(keyword_frame, text="闪烁", variable=self.enroll_flash, 
                   font=("微软雅黑", 8), width=3).grid(row=0, column=2, padx=0)
        Checkbutton(keyword_frame, text="提示音", variable=self.enroll_sound, 
                   font=("微软雅黑", 8), width=4).grid(row=0, column=3, padx=0)
        
        # 报名成功关键词行（新增自动报名和设置按钮）
        tk.Label(keyword_frame, text="报名成功关键词:", font=("微软雅黑", 8)).grid(row=1, column=0, padx=1, sticky="w")  # 标签修改
        self.success_entry = tk.Entry(keyword_frame, width=12, font=("微软雅黑", 8))
        self.success_entry.grid(row=1, column=1, padx=1, sticky="w")
        self.success_entry.insert(0, self.success_keyword)
        # 下移的自动报名复选框和设置按钮
        Checkbutton(keyword_frame, text="自动报名", variable=self.auto_enroll, 
                   font=("微软雅黑", 8), width=5).grid(row=1, column=2, padx=0)  # 移至本行
        self.set_points_btn = tk.Button(keyword_frame, text="设置报名点", width=8, 
                                       command=self.set_enroll_points, font=("微软雅黑", 8))
        self.set_points_btn.grid(row=1, column=3, padx=1)  # 移至本行
        
        # 胜利关键词行（保持不变）
        tk.Label(keyword_frame, text="胜利关键词:", font=("微软雅黑", 8)).grid(row=2, column=0, padx=1, sticky="w")
        self.victory_entry = tk.Entry(keyword_frame, width=12, font=("微软雅黑", 8))
        self.victory_entry.grid(row=2, column=1, padx=1, sticky="w")
        self.victory_entry.insert(0, self.victory_keyword)
        Checkbutton(keyword_frame, text="闪烁", variable=self.victory_flash, 
                   font=("微软雅黑", 8), width=3).grid(row=2, column=2, padx=0)
        Checkbutton(keyword_frame, text="提示音", variable=self.victory_sound, 
                   font=("微软雅黑", 8), width=4).grid(row=2, column=3, padx=0)
        
        # 领奖成功关键词行（标签修改）
        tk.Label(keyword_frame, text="领奖成功关键词:", font=("微软雅黑", 8)).grid(row=3, column=0, padx=1, sticky="w")  # 标签修改
        self.reward_success_entry = tk.Entry(keyword_frame, width=12, font=("微软雅黑", 8))  # 变量名调整
        self.reward_success_entry.grid(row=3, column=1, padx=1, sticky="w")
        self.reward_success_entry.insert(0, self.reward_success_keyword)
        
        # 7. 操作按钮（保持不变）
        btn_frame = tk.Frame(self.root)
        btn_frame.pack(pady=1)
        self.select_btn = tk.Button(btn_frame, text="框选监控区域", width=10, 
                                   command=self.select_monitor_area, font=("微软雅黑", 8))
        self.select_btn.pack(side=tk.LEFT, padx=1)
        self.start_btn = tk.Button(btn_frame, text="开始监控", width=10, 
                                  command=self.toggle_monitoring, font=("微软雅黑", 8))
        self.start_btn.pack(side=tk.LEFT, padx=1)
        
        # 8. 识别结果（保持不变）
        tk.Label(self.root, text="识别结果:", fg="blue", font=("微软雅黑", 8)).pack(anchor="w", padx=3, pady=0)
        self.recognized_text = tk.Text(self.root, height=1, width=50, 
                                      wrap=tk.WORD, font=("微软雅黑", 8))
        self.recognized_text.pack(padx=3, pady=0, fill=tk.X)
        self.recognized_text.insert(tk.END, "等待校准...")
        
        # 9. 状态提示（保持不变）
        self.status_label = tk.Label(self.root, text="监控中（灰，间隔1.0秒）| 报名点：0/3", 
                                    fg="gray", font=("微软雅黑", 7))
        self.status_label.pack(pady=1, fill=tk.X)

    # 以下方法仅调整与关键词相关的变量引用，核心逻辑不变
    def browse_tesseract(self):
        path = filedialog.askopenfilename(
            title="选择Tesseract可执行文件",
            filetypes=[("可执行文件", "tesseract.exe")]
        )
        if path:
            self.path_entry.delete(0, tk.END)
            self.path_entry.insert(0, path)
            self.tesseract_path = path

    def update_channel(self):
        self.color_channel = self.channel_var.get()
        self.status_label.config(text=f"监控中（{self.color_channel}，间隔{self.check_interval:.1f}秒）| 报名点：{len(self.enroll_points)}/3")

    def update_threshold(self, value):
        self.threshold_value = int(value)
        self.threshold_val_label.config(text=f"{self.threshold_value}")

    def update_interval(self, value):
        self.check_interval = float(value)
        self.interval_val_label.config(text=f"{self.check_interval:.1f}")
        self.status_label.config(text=f"监控中（{self.color_channel}，间隔{self.check_interval:.1f}秒）| 报名点：{len(self.enroll_points)}/3")

    def set_enroll_points(self):
        if self.monitoring:
            messagebox.showinfo("提示", "请先停止监控再设置报名点")
            return
        
        self.enroll_points = []
        selection_window = tk.Toplevel(self.root)
        selection_window.attributes("-fullscreen", True, "-topmost", True)
        selection_window.configure(bg="white")
        selection_window.attributes("-alpha", 0.5)
        
        canvas = tk.Canvas(selection_window, cursor="cross", bg="white", highlightthickness=0)
        canvas.pack(fill=tk.BOTH, expand=True)
        canvas.focus_set()
        
        guide_text = canvas.create_text(
            canvas.winfo_screenwidth()//2,
            canvas.winfo_screenheight()//2 - 50,
            text="请点击3个报名操作位置（依次点击）",
            font=("微软雅黑", 14, "bold"),
            fill="red"
        )
        count_text = canvas.create_text(
            canvas.winfo_screenwidth()//2,
            canvas.winfo_screenheight()//2 + 50,
            text=f"已选择: {len(self.enroll_points)}/3",
            font=("微软雅黑", 12),
            fill="blue"
        )
        
        def on_click(e):
            self.enroll_points.append((e.x_root, e.y_root))
            canvas.create_oval(e.x-15, e.y-15, e.x+15, e.y+15, 
                             fill="red", outline="yellow", width=3)
            canvas.create_text(e.x, e.y, text=str(len(self.enroll_points)), 
                             font=("微软雅黑", 12, "bold"), fill="white")
            canvas.itemconfig(count_text, text=f"已选择: {len(self.enroll_points)}/3")
            
            if len(self.enroll_points) >= 3:
                selection_window.destroy()
                messagebox.showinfo("成功", "3个报名点已设置完成！")
                self.status_label.config(text=f"监控中（{self.color_channel}，间隔{self.check_interval:.1f}秒）| 报名点：3/3")
        
        canvas.bind("<Button-1>", on_click)

    def select_monitor_area(self):
        selection_window = tk.Toplevel(self.root)
        selection_window.attributes("-fullscreen", True, "-alpha", 0.3, "-topmost", True)
        selection_window.configure(bg="white")
        
        canvas = tk.Canvas(selection_window, cursor="cross", bg="white", highlightthickness=0)
        canvas.pack(fill=tk.BOTH, expand=True)
        canvas.create_text(
            canvas.winfo_screenwidth()//2,
            canvas.winfo_screenheight()//2,
            text="拖动框选文字监控区域",
            font=("微软雅黑", 12),
            fill="red"
        )
        
        start_x = start_y = 0
        rect = None
        
        def on_press(e):
            nonlocal start_x, start_y, rect
            start_x, start_y = e.x, e.y
            rect = canvas.create_rectangle(0, 0, 0, 0, outline="red", width=3)
        
        def on_drag(e):
            nonlocal rect
            canvas.coords(rect, start_x, start_y, e.x, e.y)
        
        def on_release(e):
            nonlocal start_x, start_y
            self.monitor_area = {
                "top": min(start_y, e.y),
                "left": min(start_x, e.x),
                "width": abs(e.x - start_x),
                "height": abs(e.y - start_y)
            }
            selection_window.destroy()
            messagebox.showinfo("成功", "监控区域已设置")
            if hasattr(self, 'preview_label') and self.preview_label:
                self.preview_label.destroy()
            self.status_label.config(text=f"监控中（{self.color_channel}，间隔{self.check_interval:.1f}秒）| 报名点：{len(self.enroll_points)}/3")
            self.is_calibrating = True
            threading.Thread(target=self.calibration_loop, daemon=True).start()
        
        canvas.bind("<ButtonPress-1>", on_press)
        canvas.bind("<B1-Motion>", on_drag)
        canvas.bind("<ButtonRelease-1>", on_release)

    def update_preview(self):
        if self.processed_img is None:
            return
        try:
            h, w = self.processed_img.shape[:2]
            max_w = self.preview_canvas.winfo_width()
            max_h = self.preview_canvas.winfo_height()
            scale = min(max_w/w, max_h/h)
            resized_img = cv2.resize(self.processed_img, (int(w*scale), int(h*scale)))
            resized_img = cv2.cvtColor(resized_img, cv2.COLOR_GRAY2RGB)
            tk_img = ImageTk.PhotoImage(image=Image.fromarray(resized_img))
            
            self.preview_canvas.delete("all")
            self.preview_canvas.create_image(max_w//2, max_h//2, image=tk_img, anchor="center")
            self.preview_canvas.image = tk_img
        except:
            self.preview_canvas.delete("all")
            self.preview_canvas.create_text(self.preview_canvas.winfo_width()//2, 
                                           self.preview_canvas.winfo_height()//2, 
                                           text="预览错误", fill="red", font=("微软雅黑", 8))

    def calibration_loop(self):
        while self.is_calibrating and not self.monitoring and self.monitor_area:
            self.processed_img = self.preprocess_image()
            self.root.after(0, self.update_preview)
            text = self.capture_and_ocr()
            self.recognized_text.delete(1.0, tk.END)
            self.recognized_text.insert(tk.END, text)
            time.sleep(0.2)

    def toggle_monitoring(self):
        if self.monitoring:
            self.monitoring = False
            self.enroll_alert_running = False
            self.victory_alert_running = False
            self.countdown_running = False
            self.start_btn.config(text="开始监控")
            self.threshold_slider.config(state=tk.NORMAL)
            self.interval_slider.config(state=tk.NORMAL)
            self.set_points_btn.config(state=tk.NORMAL)
            self.is_calibrating = True
            threading.Thread(target=self.calibration_loop, daemon=True).start()
            self.status_label.config(text=f"监控已停止 | 报名点：{len(self.enroll_points)}/3", fg="red")
            self.restore_styles()
        else:
            if not os.path.exists(self.path_entry.get().strip()):
                messagebox.showerror("错误", "Tesseract路径无效")
                return
            if not self.monitor_area:
                messagebox.showerror("错误", "请先框选监控区域")
                return
            if self.auto_enroll.get() and len(self.enroll_points) < 3:
                messagebox.showerror("错误", "自动报名需要设置3个报名点")
                return
            
            tessdata = os.path.join(os.path.dirname(self.path_entry.get()), "tessdata")
            if not os.path.exists(os.path.join(tessdata, "chi_sim.traineddata")):
                messagebox.showerror("错误", "缺少中文语言包")
                return
            
            self.last_enroll_time = 0
            self.monitoring = True
            self.is_calibrating = False
            self.start_btn.config(text="停止监控")
            self.threshold_slider.config(state=tk.DISABLED)
            self.interval_slider.config(state=tk.DISABLED)
            self.set_points_btn.config(state=tk.DISABLED)
            self.status_label.config(
                text=f"监控中（{self.color_channel}，间隔{self.check_interval:.1f}秒）| 报名点：{len(self.enroll_points)}/3", 
                fg="green"
            )
            threading.Thread(target=self.monitoring_loop, daemon=True).start()

    def preprocess_image(self):
        try:
            with mss.mss() as sct:
                img = np.array(sct.grab(self.monitor_area))
            
            if self.color_channel == "灰":
                channel = cv2.cvtColor(img, cv2.COLOR_BGRA2GRAY)
            else:
                channels = cv2.split(img)
                channel = channels[2] if self.color_channel == "红" else \
                          channels[1] if self.color_channel == "绿" else channels[0]
            
            _, thresh = cv2.threshold(
                channel, self.threshold_value, 255,
                cv2.THRESH_BINARY_INV if self.invert_binary.get() else cv2.THRESH_BINARY
            )
            return thresh
        except:
            return None

    def capture_and_ocr(self):
        if self.processed_img is None:
            return "无图像"
        try:
            pytesseract.pytesseract.tesseract_cmd = self.tesseract_path
            return pytesseract.image_to_string(
                self.processed_img, lang="chi_sim", 
                config=r'--oem 3 --psm 11'
            ).strip() or "未识别到文字"
        except:
            return "识别失败"

    def auto_click_enroll_points(self):
        if not self.monitoring or len(self.enroll_points) != 3:
            return
        
        initial_x, initial_y = win32api.GetCursorPos()
        try:
            for i, (x, y) in enumerate(self.enroll_points, 1):
                pyautogui.moveTo(x, y, duration=0.01)
                self.lock_mouse(x, y)
                pyautogui.click()
                self.status_label.config(text=f"自动报名：点击第{i}个位置 ({x},{y})")
                time.sleep(self.click_interval)
        except Exception as e:
            self.status_label.config(text=f"自动点击失败：{str(e)}")
        finally:
            win32api.SetCursorPos((initial_x, initial_y))

    def lock_mouse(self, x, y):
        win32api.SetCursorPos((x, y))

    def play_sound(self, freq=1000, duration=500):
        try:
            winsound.Beep(freq, duration)
        except:
            pass

    def set_style(self, color_bg, color_fg, status_text):
        self.root.configure(bg=color_bg)
        for widget in [self.path_entry, self.enroll_entry, self.success_entry,
                      self.victory_entry, self.reward_success_entry, self.recognized_text]:  # 调整变量名
            widget.configure(bg=color_bg.replace("cc", "e6"))
        for label in self.root.winfo_children():
            if isinstance(label, tk.Label):
                label.configure(bg=color_bg, fg=color_fg)
        self.status_label.config(text=status_text)

    def set_green_style(self):
        self.set_style("#ccffcc", "#006600", f"检测到报名：{self.enroll_keyword}")

    def restore_styles(self):
        self.root.configure(bg=self.root.cget("bg"))
        for widget in [self.path_entry, self.enroll_entry, self.success_entry,
                      self.victory_entry, self.reward_success_entry, self.recognized_text]:  # 调整变量名
            widget.configure(bg="white")
        for label in self.root.winfo_children():
            if isinstance(label, tk.Label):
                label.configure(bg="SystemButtonFace", fg="black")
        self.status_label.config(
            text=f"监控中（{self.color_channel}，间隔{self.check_interval:.1f}秒）| 报名点：{len(self.enroll_points)}/3"
        )

    def enroll_alert_loop(self):
        start_time = time.time()
        self.enroll_alert_running = True
        while self.enroll_alert_running and self.monitoring:
            # 引用修改后的报名成功关键词输入框
            if (time.time() - start_time >= 30) or (self.success_entry.get() in self.recognized_text.get(1.0, tk.END)):
                self.enroll_alert_running = False
                self.root.after(0, self.restore_styles)
                break
            
            if self.enroll_flash.get():
                self.root.after(0, self.set_green_style)
            if self.enroll_sound.get():
                self.play_sound(1500, 300)
            
            time.sleep(2)
            
            if self.enroll_flash.get() and self.enroll_alert_running:
                self.root.after(0, self.restore_styles)
                time.sleep(0.1)

    def victory_alert_loop(self):
        start_time = time.time()
        self.victory_alert_running = True
        while self.victory_alert_running and self.monitoring:
            # 引用修改后的领奖成功关键词输入框
            if (time.time() - start_time >= 30) or (self.reward_success_entry.get() in self.recognized_text.get(1.0, tk.END)):
                self.victory_alert_running = False
                self.root.after(0, self.restore_styles)
                break
            
            if self.victory_flash.get():
                self.root.after(0, lambda: self.set_style("#ffcccc", "#cc0000", f"检测到胜利：{self.victory_keyword}"))
            if self.victory_sound.get():
                self.play_sound(800, 300)
            
            time.sleep(2)
            
            if self.victory_flash.get() and self.victory_alert_running:
                self.root.after(0, self.restore_styles)
                time.sleep(0.1)

    def victory_countdown(self):
        countdown = 30
        self.countdown_running = True
        while countdown > 0 and self.countdown_running and self.monitoring:
            self.root.after(0, lambda c=countdown: self.status_label.config(text=f"胜利倒计时：{c}秒"))
            time.sleep(1)
            countdown -= 1
        if self.monitoring:
            self.root.after(0, lambda: self.status_label.config(
                text=f"监控中（{self.color_channel}，间隔{self.check_interval:.1f}秒）| 报名点：{len(self.enroll_points)}/3"
            ))
        self.countdown_running = False

    def monitoring_loop(self):
        while self.monitoring:
            try:
                self.processed_img = self.preprocess_image()
                self.root.after(0, self.update_preview)
                
                text = self.capture_and_ocr()
                self.recognized_text.delete(1.0, tk.END)
                self.recognized_text.insert(tk.END, text)
                
                current_time = time.time()

                # 检测报名成功关键词（引用修改后的输入框）
                if self.success_entry.get() in text:
                    self.enroll_alert_running = False
                    self.root.after(0, self.restore_styles)

                # 检测领奖成功关键词（引用修改后的输入框）
                if self.reward_success_entry.get() in text:
                    self.victory_alert_running = False
                    self.countdown_running = False
                    self.root.after(0, self.restore_styles)

                # 检测报名关键词
                if (self.enroll_entry.get() in text and self.success_entry.get() not in text and
                    current_time - self.last_enroll_time >= self.enroll_cooldown and
                    not self.enroll_alert_running):
                    
                    self.last_enroll_time = current_time
                    if self.auto_enroll.get():
                        self.auto_click_enroll_points()
                    threading.Thread(target=self.enroll_alert_loop, daemon=True).start()

                # 检测胜利关键词
                if (self.victory_entry.get() in text and self.reward_success_entry.get() not in text and  # 调整变量
                    not self.victory_alert_running and not self.countdown_running):
                    
                    threading.Thread(target=self.victory_alert_loop, daemon=True).start()
                    self.countdown_thread = threading.Thread(target=self.victory_countdown, daemon=True)
                    self.countdown_thread.start()
            
            except Exception as e:
                self.recognized_text.delete(1.0, tk.END)
                self.recognized_text.insert(tk.END, f"监控错误：{str(e)}")
            
            time.sleep(self.check_interval)


if __name__ == "__main__":
    try:
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)
    except:
        pass
    
    try:
        root = tk.Tk()
        app = AutoEnrollOCRMonitor(root)
        root.mainloop()
    except Exception as e:
        messagebox.showerror("启动失败", f"错误：{str(e)}\n需安装依赖：pip install opencv-python pyautogui pytesseract pywin32 mss pillow")
        input("按回车退出...")
