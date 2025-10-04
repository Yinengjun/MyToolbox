import tkinter as tk
from tkinter import ttk, messagebox, Menu
import os
import json
import subprocess
import sys
from pathlib import Path
import platform

# Windows图标提取
if platform.system() == "Windows":
    try:
        import win32api
        import win32con
        import win32ui
        import win32gui
        from PIL import Image, ImageTk
        ICON_SUPPORT = True
    except:
        ICON_SUPPORT = False
else:
    ICON_SUPPORT = False

class MyToolbox:
    def __init__(self, root):
        self.root = root
        self.root.title("MyToolbox")
        self.root.geometry("1000x600")
        
        # 配置文件路径
        self.tools_dir = Path("tools")
        self.data_file = Path("toolbox_data.json")
        
        # 存储自定义名称和说明
        self.custom_data = self.load_data()
        
        # 图标缓存
        self.icon_cache = {}
        
        # 创建工具文件夹（如果不存在）
        self.tools_dir.mkdir(exist_ok=True)
        
        self.setup_ui()
        self.load_categories()
        
    def setup_ui(self):
        """设置用户界面"""
        # 主容器
        main_container = ttk.PanedWindow(self.root, orient=tk.HORIZONTAL)
        main_container.pack(fill=tk.BOTH, expand=True)
        
        # 左侧边栏
        sidebar_frame = ttk.Frame(main_container, width=200)
        main_container.add(sidebar_frame, weight=1)
        
        # 分类列表
        self.category_listbox = tk.Listbox(sidebar_frame, font=("微软雅黑", 11))
        self.category_listbox.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.category_listbox.bind("<<ListboxSelect>>", self.on_category_select)
        
        # 右侧内容区
        content_frame = ttk.Frame(main_container)
        main_container.add(content_frame, weight=4)
        
        # 工具网格区域（使用Canvas和Scrollbar）
        self.canvas = tk.Canvas(content_frame, bg="white")
        scrollbar = ttk.Scrollbar(content_frame, orient="vertical", command=self.canvas.yview)
        self.tools_frame = ttk.Frame(self.canvas)
        
        self.canvas.configure(yscrollcommand=scrollbar.set)
        
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        self.canvas_frame = self.canvas.create_window((0, 0), window=self.tools_frame, anchor="nw")
        
        # 更新滚动区域
        def configure_scroll_region(event):
            self.canvas.configure(scrollregion=self.canvas.bbox("all"))
            canvas_width = event.width - 10
            self.canvas.itemconfig(self.canvas_frame, width=canvas_width)
        
        self.canvas.bind("<Configure>", configure_scroll_region)
        
        # 鼠标滚轮支持
        def on_mousewheel(event):
            self.canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        self.canvas.bind_all("<MouseWheel>", on_mousewheel)
        
    def load_data(self):
        """从JSON文件加载数据"""
        if self.data_file.exists():
            try:
                with open(self.data_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except:
                return {}
        return {}
    
    def save_data(self):
        """保存数据到JSON文件"""
        with open(self.data_file, 'w', encoding='utf-8') as f:
            json.dump(self.custom_data, f, ensure_ascii=False, indent=2)
    
    def load_categories(self):
        """加载分类列表"""
        self.category_listbox.delete(0, tk.END)
        
        if not self.tools_dir.exists():
            return
        
        categories = [d.name for d in self.tools_dir.iterdir() if d.is_dir()]
        categories.sort()
        
        for category in categories:
            self.category_listbox.insert(tk.END, category)
        
        # 自动选择第一个分类
        if categories:
            self.category_listbox.selection_set(0)
            self.on_category_select(None)
    
    def on_category_select(self, event):
        """分类选择事件"""
        selection = self.category_listbox.curselection()
        if not selection:
            return
        
        category = self.category_listbox.get(selection[0])
        self.load_tools(category)
    
    def _hicon_to_photo(self, hicon, size=48):
        """HICON -> Tk PhotoImage, 并清理GDI资源"""
        if not hicon:
            return None

        ico_x = win32api.GetSystemMetrics(win32con.SM_CXICON)
        ico_y = win32api.GetSystemMetrics(win32con.SM_CYICON)

        hdc_screen = None
        dc = None
        memdc = None
        bmp = None
        try:
            hdc_screen = win32gui.GetDC(0)
            dc = win32ui.CreateDCFromHandle(hdc_screen)
            memdc = dc.CreateCompatibleDC()
            bmp = win32ui.CreateBitmap()
            bmp.CreateCompatibleBitmap(dc, ico_x, ico_y)
            memdc.SelectObject(bmp)
            memdc.DrawIcon((0, 0), hicon)

            # BGRX -> RGB
            bmpstr = bmp.GetBitmapBits(True)
            img = Image.frombuffer('RGB', (ico_x, ico_y), bmpstr, 'raw', 'BGRX', 0, 1)
            if size and size != ico_x:
                img = img.resize((size, size), Image.Resampling.LANCZOS)
            return ImageTk.PhotoImage(img)
        finally:
            # 释放GDI与图标句柄
            try:
                if hicon:
                    win32gui.DestroyIcon(hicon)
            except Exception:
                pass
            try:
                if memdc:
                    memdc.DeleteDC()
                if dc:
                    dc.DeleteDC()
            except Exception:
                pass
            try:
                if hdc_screen:
                    win32gui.ReleaseDC(0, hdc_screen)
            except Exception:
                pass
            try:
                if bmp:
                    win32gui.DeleteObject(bmp.GetHandle())
            except Exception:
                pass

    def get_file_icon(self, file_path):
        """获取文件图标(含.lnk解析, 普通文件关联图标), 带缓存"""
        if not ICON_SUPPORT:
            return None

        path = str(file_path)
        suffix = file_path.suffix.lower()
        icon_src = path     # 实际提取图标的来源路径(可能是lnk目标或IconLocation)
        index = 0           # 资源索引(lnk的IconLocation可能指定)
        # 解析 .lnk
        if suffix == ".lnk":
            try:
                import win32com.client
                wsh = win32com.client.Dispatch("WScript.Shell")
                sc = wsh.CreateShortcut(path)
                loc = sc.IconLocation or ""
                if loc:
                    # 形如: "C:\\Path\\app.exe,0"
                    parts = loc.split(',')
                    icon_src = parts[0].strip().strip('"')
                    if len(parts) > 1 and parts[1].strip():
                        try:
                            index = int(parts[1])
                        except ValueError:
                            index = 0
                elif sc.TargetPath:
                    icon_src = sc.TargetPath
            except Exception:
                # 解析失败则退回用链接文件本身
                icon_src = path
                index = 0

        # 缓存key: 对资源文件/lnk使用 "路径@索引"；其他类型用后缀共享
        if suffix in (".exe", ".dll", ".ico") or suffix == ".lnk":
            cache_key = f"{icon_src.lower()}@{index}"
        else:
            cache_key = suffix

        if cache_key in self.icon_cache:
            return self.icon_cache[cache_key]

        hicon = None

        # 1) 优先: 从资源文件提取(适用于 exe/dll/ico 或明确索引的lnk)
        try:
            large, small = win32gui.ExtractIconEx(icon_src, index)
            # 优先用small, 不够则用large
            if small:
                hicon = small[0]
                for h in small[1:]:
                    win32gui.DestroyIcon(h)
            elif large:
                hicon = large[0]
                for h in large[1:]:
                    win32gui.DestroyIcon(h)
        except Exception:
            hicon = None

        # 2) 兜底: 取系统关联图标(SHGetFileInfo)
        if not hicon:
            try:
                from win32com.shell import shell, shellcon
                flags = shellcon.SHGFI_ICON | shellcon.SHGFI_LARGEICON
                # 若文件不存在而只想按扩展名取图标, 可加: | shellcon.SHGFI_USEFILEATTRIBUTES
                ret = shell.SHGetFileInfo(icon_src, 0, flags)
                # 返回 (hIcon, iIcon, dwAttributes, displayName, typeName)
                if isinstance(ret, tuple) and ret and ret[0]:
                    hicon = ret[0]
            except Exception:
                hicon = None

        if not hicon:
            return None

        photo = self._hicon_to_photo(hicon, size=48)
        self.icon_cache[cache_key] = photo
        return photo
    
    def load_tools(self, category):
        """加载指定分类下的工具 - 网格布局"""
        # 清空当前工具列表
        for widget in self.tools_frame.winfo_children():
            widget.destroy()
        
        self.icon_cache.clear()  # 清空图标缓存
        
        category_path = self.tools_dir / category
        if not category_path.exists():
            return
        
        # 获取所有文件和.lnk（不包括子文件夹）
        files = []
        for item in category_path.iterdir():
            if item.is_file():
                files.append(item)
        
        files.sort(key=lambda x: x.name)
        
        if not files:
            ttk.Label(self.tools_frame, text="此分类下暂无工具", 
                     font=("微软雅黑", 10)).grid(row=0, column=0, pady=20, padx=20)
            return
        
        # 计算每行显示的工具数量
        items_per_row = 5
        
        # 显示每个工具（网格布局）
        for idx, file_path in enumerate(files):
            row = idx // items_per_row
            col = idx % items_per_row
            self.create_tool_grid_item(file_path, category, row, col)
    
    def create_tool_grid_item(self, file_path, category, row, col):
        """创建网格形式的工具项"""
        file_key = f"{category}/{file_path.name}"
        
        # 获取显示名称和说明
        custom_name = self.custom_data.get(file_key, {}).get("name", file_path.stem)
        description = self.custom_data.get(file_key, {}).get("description", "")
        
        # 工具项容器
        item_frame = ttk.Frame(self.tools_frame)
        item_frame.grid(row=row, column=col, padx=15, pady=15, sticky="n")
        
        # 图标
        icon_label = tk.Label(item_frame, bg="white", cursor="hand2")
        
        # 尝试获取文件图标
        icon = self.get_file_icon(file_path)
        if icon:
            icon_label.config(image=icon)
            icon_label.image = icon  # 保持引用
        else:
            # 默认图标（使用文字）
            icon_label.config(text="📦", font=("Segoe UI Emoji", 36), 
                            width=3, height=2, relief=tk.FLAT)
        
        icon_label.pack()
        
        # 名称标签
        name_label = tk.Label(item_frame, text=custom_name, 
                             font=("微软雅黑", 9), wraplength=100,
                             cursor="hand2", bg="white")
        name_label.pack()
        
        # 双击打开
        icon_label.bind("<Double-Button-1>", lambda e: self.run_tool(file_path))
        name_label.bind("<Double-Button-1>", lambda e: self.run_tool(file_path))
        
        # 右键菜单
        def show_context_menu(event):
            menu = Menu(self.root, tearoff=0)
            menu.add_command(label="打开", 
                           command=lambda: self.run_tool(file_path))
            menu.add_command(label="编辑信息", 
                           command=lambda: self.edit_tool_info(file_key, file_path.name, category))
            menu.post(event.x_root, event.y_root)
        
        icon_label.bind("<Button-3>", show_context_menu)
        name_label.bind("<Button-3>", show_context_menu)
        
        # 悬停显示说明
        if description:
            self.create_tooltip(icon_label, description)
            self.create_tooltip(name_label, description)
    
    def create_tooltip(self, widget, text):
        """创建工具提示"""
        def show_tooltip(event):
            tooltip = tk.Toplevel()
            tooltip.wm_overrideredirect(True)
            tooltip.wm_geometry(f"+{event.x_root+10}+{event.y_root+10}")
            
            label = tk.Label(tooltip, text=text, background="lightyellow", 
                           relief=tk.SOLID, borderwidth=1, padx=8, pady=5,
                           font=("微软雅黑", 9), wraplength=300)
            label.pack()
            
            widget.tooltip = tooltip
        
        def hide_tooltip(event):
            if hasattr(widget, 'tooltip'):
                widget.tooltip.destroy()
                delattr(widget, 'tooltip')
        
        widget.bind("<Enter>", show_tooltip)
        widget.bind("<Leave>", hide_tooltip)
    
    def run_tool(self, file_path):
        """运行工具"""
        try:
            if sys.platform == "win32":
                os.startfile(file_path)
            elif sys.platform == "darwin":
                subprocess.call(["open", file_path])
            else:
                subprocess.call(["xdg-open", file_path])
        except Exception as e:
            messagebox.showerror("错误", f"无法运行工具:\n{str(e)}")
    
    def edit_tool_info(self, file_key, original_name, category):
        """编辑工具信息"""
        # 创建编辑对话框
        dialog = tk.Toplevel(self.root)
        dialog.title("编辑工具信息")
        dialog.geometry("400x250")
        dialog.transient(self.root)
        dialog.grab_set()
        
        # 获取当前数据
        current_data = self.custom_data.get(file_key, {})
        original_stem = Path(original_name).stem
        current_name = current_data.get("name", original_stem)
        current_desc = current_data.get("description", "")
        
        # 自定义名称
        ttk.Label(dialog, text="自定义名称:", font=("微软雅黑", 10)).pack(pady=(20, 5))
        name_entry = ttk.Entry(dialog, font=("微软雅黑", 10), width=40)
        name_entry.pack(pady=5)
        name_entry.insert(0, current_name)
        
        # 解释说明
        ttk.Label(dialog, text="解释说明:", font=("微软雅黑", 10)).pack(pady=(10, 5))
        desc_text = tk.Text(dialog, font=("微软雅黑", 9), width=40, height=5)
        desc_text.pack(pady=5)
        desc_text.insert("1.0", current_desc)
        
        # 按钮
        button_frame = ttk.Frame(dialog)
        button_frame.pack(pady=20)
        
        def save_changes():
            new_name = name_entry.get().strip()
            new_desc = desc_text.get("1.0", tk.END).strip()
            
            if not new_name:
                new_name = original_stem
            
            self.custom_data[file_key] = {
                "name": new_name,
                "description": new_desc
            }
            
            self.save_data()
            dialog.destroy()
            # 重新加载当前分类
            self.load_tools(category)
        
        def reset_changes():
            if file_key in self.custom_data:
                del self.custom_data[file_key]
                self.save_data()
                dialog.destroy()
                self.load_tools(category)
        
        ttk.Button(button_frame, text="保存", command=save_changes).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="重置", command=reset_changes).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="取消", command=dialog.destroy).pack(side=tk.LEFT, padx=5)


def main():
    root = tk.Tk()
    app = MyToolbox(root)
    root.mainloop()


if __name__ == "__main__":
    main()