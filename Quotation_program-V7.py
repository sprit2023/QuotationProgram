import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import json
import os
from datetime import datetime
import sys

class QuotationApp:
    def __init__(self, root):
        self.root = root
        self.root.title("报价单系统-Sprit.Zeng V3.0-测试版")
        self.root.geometry("1200x1000")

        # 绑定窗口最小化和恢复事件
        self.root.bind("<Unmap>", self.on_minimize)  # 窗口最小化时触发
        self.root.bind("<Map>", self.on_restore)     # 窗口恢复时触发

        # 列对应关系
        self.COLUMN_MAPPING = {
            "物料编码": "物料编码",
            "物料名称": "物料名称",
            "规格型号": "规格型号",
            "数量": "数量",
            "含税单价": "含税单价"
        }

        # 存储完整的产品数据
        self.full_product_data = []

        # 历史记录文件路径
        self.history_file = self.get_history_file_path()

        # 初始化
        self.create_widgets()

        # 加载历史记录
        self.load_history_from_file()

    def on_minimize(self, event):
        """窗口最小化时的处理"""
        print("窗口已最小化")

    def on_restore(self, event):
        """窗口恢复时的处理"""
        print("窗口已恢复")
        self.root.deiconify()  # 强制显示窗口
        self.root.lift()       # 将窗口置于最上层
        self.root.focus_force()  # 强制聚焦窗口

    def get_history_file_path(self):
        """获取历史记录文件的路径"""
        if hasattr(sys, '_MEIPASS'):
            # 打包后的运行环境
            base_path = sys._MEIPASS
        else:
            # 开发环境
            base_path = os.path.abspath(".")

        # 将历史记录文件存储在程序目录下
        history_file = os.path.join(base_path, "quotation_history.json")

        # 如果文件不存在，则创建一个空文件
        if not os.path.exists(history_file):
            with open(history_file, "w", encoding="utf-8") as file:
                json.dump([], file)

        return history_file

    def create_widgets(self):
        # 主布局
        main_frame = tk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # 工具栏
        toolbar = tk.Frame(main_frame)
        toolbar.pack(fill=tk.X, pady=5)

        # 保存按钮
        btn_save = tk.Button(toolbar, text="保存报价单", command=self.save_quotation)
        btn_save.pack(side=tk.RIGHT, padx=5)

        # 导出按钮
        btn_export = tk.Button(toolbar, text="导出报价单", command=self.export_excel)
        btn_export.pack(side=tk.RIGHT, padx=5)

        # 导入按钮
        btn_import = tk.Button(toolbar, text="导入Excel", command=self.import_excel)
        btn_import.pack(side=tk.LEFT, padx=5)

        # 搜索框
        self.search_var = tk.StringVar()
        self.search_entry = ttk.Entry(toolbar, textvariable=self.search_var, width=40)
        self.search_entry.pack(side=tk.LEFT, padx=10)
        self.search_entry.bind("<KeyRelease>", self.filter_products)

        # 产品表格
        self.product_frame = tk.LabelFrame(main_frame, text="产品列表")
        self.product_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        self.create_product_table()

        # 报价单表格
        self.quotation_frame = tk.LabelFrame(main_frame, text="报价单")
        self.quotation_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        self.create_quotation_table()

        # 底部信息
        bottom_frame = tk.Frame(main_frame)
        bottom_frame.pack(fill=tk.X, pady=10)

        # 第一行：含税总价和毛利率
        row1_frame = tk.Frame(bottom_frame)
        row1_frame.pack(fill=tk.X, pady=5)

        # 含税总价
        tk.Label(row1_frame, text="含税总价：").pack(side=tk.LEFT, padx=(46, 0))
        self.total_label = ttk.Entry(row1_frame, font=("Arial", 14, "bold"), state="readonly", width=15)
        self.total_label.pack(side=tk.LEFT, padx=5)

        tk.Label(row1_frame, text="大写金额：").pack(side=tk.LEFT, padx=(20, 0))
        self.total_cn_label = ttk.Entry(row1_frame, font=("Arial", 14), state="readonly", width=40)
        self.total_cn_label.pack(side=tk.LEFT, padx=5)

        # 毛利率输入框
        tk.Label(row1_frame, text="毛利率（%）：").pack(side=tk.LEFT, padx=(40, 0))
        self.profit_margin_entry = ttk.Entry(row1_frame, width=10)
        self.profit_margin_entry.pack(side=tk.LEFT, padx=5)
        self.profit_margin_entry.bind("<KeyRelease>", self.calculate_total)

        # 第二行：最终含税总价
        row2_frame = tk.Frame(bottom_frame)
        row2_frame.pack(fill=tk.X, pady=5)

        tk.Label(row2_frame, text="最终含税总价：").pack(side=tk.LEFT, padx=(20, 0))
        self.final_total_label = ttk.Entry(row2_frame, font=("Arial", 14, "bold"), state="readonly", width=15)
        self.final_total_label.pack(side=tk.LEFT, padx=5)

        tk.Label(row2_frame, text="大写金额：").pack(side=tk.LEFT, padx=(20, 0))
        self.final_total_cn_label = ttk.Entry(row2_frame, font=("Arial", 14), state="readonly", width=40)
        self.final_total_cn_label.pack(side=tk.LEFT, padx=5)

        # 清空配置按钮
        btn_clear = tk.Button(row2_frame, text="清空配置", command=self.clear_quotation)
        btn_clear.pack(side=tk.LEFT, padx=(30, 0))

        # 主内容区域
        content_frame = tk.Frame(main_frame)
        content_frame.pack(fill=tk.BOTH, expand=True)

        # 左侧区域
        left_frame = tk.Frame(content_frame)
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # 选中商品信息
        self.selected_item_frame = tk.LabelFrame(left_frame, text="选中商品信息")
        self.selected_item_frame.pack(fill=tk.BOTH, expand=True, pady=10)

        self.selected_item_info = tk.Text(self.selected_item_frame, font=("Arial", 12))
        self.selected_item_info.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # 右侧历史记录区域
        history_frame = tk.LabelFrame(content_frame, text="历史报价记录")
        history_frame.pack(side=tk.RIGHT, fill=tk.BOTH, padx=10, pady=10, ipadx=10, ipady=10)

        # 历史记录表格
        self.history_tree = ttk.Treeview(history_frame, columns=("时间", "总金额", "毛利率", "操作"), show="headings")
        self.history_tree.heading("时间", text="时间")
        self.history_tree.heading("总金额", text="总金额")
        self.history_tree.heading("毛利率", text="毛利率")
        self.history_tree.heading("操作", text="操作")
        self.history_tree.column("时间", width=150)
        self.history_tree.column("总金额", width=100)
        self.history_tree.column("毛利率", width=80)
        self.history_tree.column("操作", width=50, anchor="center")
        self.history_tree.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # 绑定删除事件
        self.history_tree.bind("<Button-1>", self.delete_history_item)

        # 历史记录操作按钮区域
        btn_frame = tk.Frame(history_frame)
        btn_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=5)

        # 添加删除按钮
        btn_delete = tk.Button(btn_frame, text="删除记录", command=self.delete_history)
        btn_delete.pack(side=tk.LEFT, padx=5)

        # 绑定历史记录双击事件
        self.history_tree.bind("<Double-1>", self.load_history_quotation)

    def create_product_table(self):
        columns = list(self.COLUMN_MAPPING.values())
        self.product_tree = ttk.Treeview(self.product_frame, columns=columns, show="headings")

        # 设置列宽
        col_widths = [150, 200, 250, 80, 100]
        for col, width in zip(columns, col_widths):
            self.product_tree.heading(col, text=col)
            self.product_tree.column(col, width=width, anchor="center")

        # 添加滚动条
        scrollbar = ttk.Scrollbar(self.product_frame, orient=tk.VERTICAL, command=self.product_tree.yview)
        self.product_tree.configure(yscroll=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.product_tree.pack(fill=tk.BOTH, expand=True)

        # 绑定双击事件
        self.product_tree.bind("<Double-1>", self.add_to_quotation)

        # 绑定 Ctrl+C 复制事件
        self.product_tree.bind("<Control-c>", self.copy_selected_text)

    def create_quotation_table(self):
        columns = ("物料编码", "物料名称", "规格型号", "数量", "含税单价", "小计", "操作")
        self.quotation_tree = ttk.Treeview(self.quotation_frame, columns=columns, show="headings")

        # 设置列宽
        col_widths = [120, 150, 250, 80, 100, 100, 80]
        for col, width in zip(columns, col_widths):
            self.quotation_tree.heading(col, text=col)
            self.quotation_tree.column(col, width=width, anchor="center")

        # 添加滚动条
        scrollbar = ttk.Scrollbar(self.quotation_frame, orient=tk.VERTICAL, command=self.quotation_tree.yview)
        self.quotation_tree.configure(yscroll=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.quotation_tree.pack(fill=tk.BOTH, expand=True)

        # 绑定删除事件（仅通过删除按钮或 DEL 键）
        self.quotation_tree.bind("<Delete>", self.delete_item)

        # 绑定选中事件
        self.quotation_tree.bind("<<TreeviewSelect>>", self.show_selected_item_info)

        # 绑定双击事件，用于编辑数量和含税单价
        self.quotation_tree.bind("<Double-1>", self.edit_quotation_item)

        # 绑定操作列的点击事件
        self.quotation_tree.bind("<Button-1>", self.handle_operation_click)

        # 用于存储编辑框
        self.quotation_edit_entry = None
        self.quotation_edit_column = None
        self.quotation_edit_item = None

    def handle_operation_click(self, event):
        """处理操作列的点击事件"""
        region = self.quotation_tree.identify_region(event.x, event.y)
        if region != "cell":
            return

        column = self.quotation_tree.identify_column(event.x)
        item = self.quotation_tree.identify_row(event.y)

        # 只处理操作列（第7列）
        if column == "#7":
            self.quotation_tree.delete(item)
            self.calculate_total()  # 重新计算总价

    def edit_quotation_item(self, event):
        """双击报价单中的数量或含税单价列，进入编辑模式"""
        region = self.quotation_tree.identify_region(event.x, event.y)
        if region != "cell":
            return

        column = self.quotation_tree.identify_column(event.x)
        item = self.quotation_tree.identify_row(event.y)

        # 只允许编辑“数量”和“含税单价”列
        if column not in ("#4", "#5"):
            return

        # 获取当前单元格的值
        values = self.quotation_tree.item(item, "values")
        col_index = int(column[1:]) - 1  # 列索引从0开始
        cell_value = values[col_index]

        # 获取单元格的坐标和大小
        x, y, width, height = self.quotation_tree.bbox(item, column)

        # 创建编辑框
        self.quotation_edit_entry = ttk.Entry(self.quotation_frame, width=width)
        self.quotation_edit_entry.place(x=x, y=y, width=width, height=height)
        self.quotation_edit_entry.insert(0, cell_value)
        self.quotation_edit_entry.focus()

        # 绑定回车键和失去焦点事件
        self.quotation_edit_entry.bind("<Return>", lambda e: self.save_quotation_edit(item, col_index))
        self.quotation_edit_entry.bind("<FocusOut>", lambda e: self.save_quotation_edit(item, col_index))

        # 存储当前编辑的列和行
        self.quotation_edit_column = col_index
        self.quotation_edit_item = item

    def save_quotation_edit(self, item, col_index):
        """保存编辑后的值"""
        if self.quotation_edit_entry:
            new_value = self.quotation_edit_entry.get()

            # 更新报价单中的值
            values = list(self.quotation_tree.item(item, "values"))
            values[col_index] = new_value

            # 如果是数量或含税单价列，重新计算小计
            if col_index == 3:  # 数量列
                try:
                    quantity = float(new_value)
                    unit_price = float(values[4])
                    values[5] = f"{quantity * unit_price:.2f}"  # 更新小计
                except ValueError:
                    messagebox.showerror("错误", "请输入有效的数字！")
                    return
            elif col_index == 4:  # 含税单价列
                try:
                    unit_price = float(new_value)
                    quantity = float(values[3])
                    values[5] = f"{quantity * unit_price:.2f}"  # 更新小计
                    values[4] = f"{unit_price:.2f}"  # 确保含税单价保留两位小数
                except ValueError:
                    messagebox.showerror("错误", "请输入有效的数字！")
                    return

            # 更新行数据
            self.quotation_tree.item(item, values=values)

            # 销毁编辑框
            self.quotation_edit_entry.destroy()
            self.quotation_edit_entry = None

            # 重新计算总价
            self.calculate_total()

    def show_selected_item_info(self, event):
        """显示选中商品的详细信息"""
        selected_item = self.quotation_tree.selection()
        if not selected_item:
            return

        # 获取选中行的数据
        values = self.quotation_tree.item(selected_item, "values")

        # 清空原有内容
        self.selected_item_info.delete(1.0, tk.END)

        # 显示详细信息
        info = (
            f"物料编码: {values[0]}\n"
            f"物料名称: {values[1]}\n"
            f"规格型号: {values[2]}\n"
            f"数量: {values[3]}\n"
            f"含税单价: {values[4]}\n"
            f"小计: {values[5]}\n"
        )
        self.selected_item_info.insert(tk.END, info)

    def save_quotation(self):
        """保存当前报价单到历史记录"""
        # 获取当前报价单数据
        quotation_data = []
        for item in self.quotation_tree.get_children():
            values = self.quotation_tree.item(item, "values")
            quotation_data.append({
                "物料编码": values[0],
                "物料名称": values[1],
                "规格型号": values[2],
                "数量": values[3],
                "含税单价": values[4],
                "小计": values[5]
            })

        # 获取当前时间和毛利率
        current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        try:
            profit_margin = float(self.profit_margin_entry.get())
        except ValueError:
            profit_margin = 0.0

        # 获取总金额
        total_amount = self.total_label.get().replace(",", "")

        # 添加到历史记录
        self.history_tree.insert("", "end", values=(current_time, total_amount, f"{profit_margin:.2f}%", "删除"))

        # 保存到文件
        self.save_history_to_file(current_time, total_amount, profit_margin, quotation_data)

    def save_history_to_file(self, current_time, total_amount, profit_margin, quotation_data):
        """将历史记录保存到文件"""
        history_entry = {
            "时间": current_time,
            "总金额": total_amount,
            "毛利率": f"{profit_margin:.2f}%",
            "报价单详情": quotation_data  # 添加报价单详细数据
        }

        # 读取现有历史记录
        try:
            with open(self.history_file, "r", encoding="utf-8") as file:
                history = json.load(file)
        except FileNotFoundError:
            history = []

        # 添加新记录
        history.append(history_entry)

        # 保存更新后的历史记录
        with open(self.history_file, "w", encoding="utf-8") as file:
            json.dump(history, file, ensure_ascii=False, indent=4)

    def load_history_from_file(self):
        """从文件加载历史记录"""
        if not os.path.exists(self.history_file):
            return

        with open(self.history_file, "r", encoding="utf-8") as f:
            history_data = json.load(f)

        for entry in history_data:
            self.history_tree.insert("", "end", values=(entry["时间"], entry["总金额"], entry["毛利率"], "删除"))

    def export_excel(self):
        """导出当前报价单到 Excel 文件"""
        # 获取当前报价单数据
        quotation_data = []
        for item in self.quotation_tree.get_children():
            values = self.quotation_tree.item(item, "values")
            quotation_data.append({
                "物料编码": values[0],
                "物料名称": values[1],
                "规格型号": values[2],
                "数量": values[3],
                "含税单价": values[4],
                "小计": values[5]
            })

        # 获取含税总价、毛利率和最终含税总价
        total_amount = self.total_label.get().replace(",", "")
        profit_margin = self.profit_margin_entry.get()
        final_total = self.final_total_label.get().replace(",", "")

        # 创建 DataFrame
        df = pd.DataFrame(quotation_data)
        summary_df = pd.DataFrame({
            "物料编码": ["总计"],
            "物料名称": [""],
            "规格型号": [""],
            "数量": [""],
            "含税单价": [""],
            "小计": [total_amount]
        })
        profit_margin_row = pd.DataFrame({
            "物料编码": ["毛利率"],
            "物料名称": [""],
            "规格型号": [""],
            "数量": [""],
            "含税单价": [""],
            "小计": [f"{profit_margin}%"]
        })
        final_total_row = pd.DataFrame({
            "物料编码": ["最终含税总价"],
            "物料名称": [""],
            "规格型号": [""],
            "数量": [""],
            "含税单价": [""],
            "小计": [final_total]
        })

        # 合并所有数据
        result_df = pd.concat([df, summary_df, profit_margin_row, final_total_row], ignore_index=True)

        # 弹出保存文件对话框
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")]
        )

        if file_path:
            # 导出到 Excel
            result_df.to_excel(file_path, index=False)
            messagebox.showinfo("导出成功", f"报价单已成功导出到 {file_path}")

    def import_excel(self):
        file_path = filedialog.askopenfilename(
            title="选择 Excel 文件",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )

        if file_path:
            success, message = self.load_excel_data(file_path)
            if success:
                messagebox.showinfo("导入成功", message)
            else:
                messagebox.showerror("导入失败", message)

    def load_excel_data(self, file_path):
        try:
            # 强制将“规格型号”列读取为字符串类型
            df = pd.read_excel(file_path, dtype={"规格型号": str})

            # 检查列是否匹配
            required_columns = list(self.COLUMN_MAPPING.keys())
            if not all(col in df.columns for col in required_columns):
                return False, f"Excel 文件缺少必需的列：{required_columns}"

            # 清空现有产品表格
            self.product_tree.delete(*self.product_tree.get_children())

            # 将数据加载到产品表格中
            self.full_product_data = []  # 清空完整数据
            for index, row in df.iterrows():
                values = [row[col] for col in required_columns]
                # 确保含税单价保留两位小数
                values[4] = f"{float(values[4]):.2f}"
                self.product_tree.insert("", "end", values=values)
                self.full_product_data.append(values)  # 添加到完整数据

            return True, f"成功导入 {len(df)} 条产品数据！"
        except Exception as e:
            return False, f"导入 Excel 文件时出错：{e}"

    def filter_products(self, event=None):
        search_term = self.search_var.get().strip().lower()  # 获取搜索框内容并转换为小写

        # 清空当前显示的产品列表
        self.product_tree.delete(*self.product_tree.get_children())

        # 遍历完整的产品数据
        for values in self.full_product_data:
            spec = str(values[2]).lower()  # 获取“规格型号”列的值并转换为小写

            # 如果搜索内容为空或“规格型号”包含搜索内容，则显示该行
            if not search_term or search_term in spec:
                self.product_tree.insert("", "end", values=values)

    def add_to_quotation(self, event):
        """双击产品列表中的产品，将其添加到报价表中"""
        selected_item = self.product_tree.selection()
        if not selected_item:
            return

        # 获取选中行的数据
        item_values = self.product_tree.item(selected_item, "values")
        material_code = item_values[0]  # 物料编码
        material_name = item_values[1]  # 物料名称
        spec = item_values[2]  # 规格型号
        quantity = 1  # 默认数量为 1
        unit_price = float(item_values[4])  # 含税单价

        # 检查报价表中是否已存在该产品
        existing_item = None
        for item in self.quotation_tree.get_children():
            values = self.quotation_tree.item(item, "values")
            if values[0] == material_code:  # 通过物料编码判断是否已存在
                existing_item = item
                break

        if existing_item:
            # 如果已存在，更新数量
            existing_values = self.quotation_tree.item(existing_item, "values")
            new_quantity = int(existing_values[3]) + quantity  # 叠加数量
            subtotal = new_quantity * unit_price  # 重新计算小计
            self.quotation_tree.item(existing_item, values=(
                material_code, material_name, spec, new_quantity, f"{unit_price:.2f}", f"{subtotal:.2f}", "删除"
            ))
        else:
            # 如果不存在，新增一行
            subtotal = quantity * unit_price  # 计算小计
            self.quotation_tree.insert("", "end", values=(
                material_code, material_name, spec, quantity, f"{unit_price:.2f}", f"{subtotal:.2f}", "删除"
            ))

        # 更新总价
        self.calculate_total()

    def calculate_total(self, event=None):
        """计算报价表的总价"""
        total = 0.0
        for item in self.quotation_tree.get_children():
            values = self.quotation_tree.item(item, "values")
            subtotal = float(values[5])  # 小计
            total += subtotal

        # 更新含税总价（千分位表示）
        self.total_label.config(state="normal")
        self.total_label.delete(0, tk.END)
        self.total_label.insert(0, f"{total:,.2f}")  # 使用千分位表示
        self.total_label.config(state="readonly")

        # 更新大写金额
        self.total_cn_label.config(state="normal")
        self.total_cn_label.delete(0, tk.END)
        self.total_cn_label.insert(0, self.to_chinese_amount(total))
        self.total_cn_label.config(state="readonly")

        # 计算最终含税总价（考虑毛利率）
        try:
            profit_margin = float(self.profit_margin_entry.get())
            final_total = total * (1 + profit_margin / 100)
            self.final_total_label.config(state="normal")
            self.final_total_label.delete(0, tk.END)
            self.final_total_label.insert(0, f"{final_total:,.2f}")  # 使用千分位表示
            self.final_total_label.config(state="readonly")

            # 更新最终大写金额
            self.final_total_cn_label.config(state="normal")
            self.final_total_cn_label.delete(0, tk.END)
            self.final_total_cn_label.insert(0, self.to_chinese_amount(final_total))
            self.final_total_cn_label.config(state="readonly")
        except ValueError:
            pass

    def delete_item(self, event):
        """删除报价表中的行"""
        selected_item = self.quotation_tree.selection()
        if selected_item:
            self.quotation_tree.delete(selected_item)
            self.calculate_total()  # 重新计算总价

    def clear_quotation(self):
        """清空报价表"""
        for item in self.quotation_tree.get_children():
            self.quotation_tree.delete(item)
        self.calculate_total()  # 重置总价

    def delete_history_item(self, event):
        """删除历史记录中的某一行"""
        region = self.history_tree.identify_region(event.x, event.y)
        if region != "cell":
            return

        column = self.history_tree.identify_column(event.x)
        item = self.history_tree.identify_row(event.y)

        # 只处理操作列（第4列）
        if column == "#4":
            self.history_tree.delete(item)
            self.save_history_to_file()  # 保存更新后的历史记录

    def delete_history(self):
        """删除所有历史记录"""
        for item in self.history_tree.get_children():
            self.history_tree.delete(item)
        self.save_history_to_file()  # 保存更新后的历史记录

    def load_history_quotation(self, event):
        """加载历史记录中的报价单"""
        selected_item = self.history_tree.selection()
        if not selected_item:
            return

        # 获取选中行的数据
        values = self.history_tree.item(selected_item, "values")
        time_str = values[0]

        # 弹出确认对话框
        confirm = messagebox.askyesno("恢复报价单", f"是否恢复 {time_str} 的报价单？")
        if not confirm:
            return

        # 清空当前报价表
        self.clear_quotation()

        # 从历史记录文件中加载详细的报价单数据
        with open(self.history_file, "r", encoding="utf-8") as file:
            history = json.load(file)

        # 找到对应的历史记录
        for entry in history:
            if entry["时间"] == time_str:
                # 加载报价单详情
                quotation_data = entry["报价单详情"]
                for item_data in quotation_data:
                    material_code = item_data["物料编码"]
                    material_name = item_data["物料名称"]
                    spec = item_data["规格型号"]
                    quantity = item_data["数量"]
                    unit_price = item_data["含税单价"]
                    subtotal = item_data["小计"]

                    # 添加到报价表
                    self.quotation_tree.insert("", "end", values=(
                        material_code, material_name, spec, quantity, unit_price, subtotal, "删除"
                    ))

                # 更新毛利率
                profit_margin = float(entry["毛利率"].replace("%", ""))
                self.profit_margin_entry.delete(0, tk.END)
                self.profit_margin_entry.insert(0, f"{profit_margin:.2f}")

                # 更新总价
                self.calculate_total()

    def copy_selected_text(self, event):
        """复制选中文本到剪贴板"""
        widget = event.widget
        try:
            selected_text = widget.item(widget.selection()[0])['values']
            self.root.clipboard_clear()
            self.root.clipboard_append(str(selected_text))
        except IndexError:
            pass

    def to_chinese_amount(self, amount):
        # 金额中文大写转换
        units = ["", "拾", "佰", "仟", "万", "拾", "佰", "仟", "亿", "拾", "佰", "仟", "万", "拾", "佰", "仟", "亿"]
        nums = ["零", "壹", "贰", "叁", "肆", "伍", "陆", "柒", "捌", "玖"]
        decimal_units = ["角", "分"]

        # 处理整数部分
        integer_part = int(amount)
        integer_str = str(integer_part)
        result = ""
        zero_flag = False  # 标记是否需要添加“零”

        # 从高位到低位处理
        length = len(integer_str)
        for i, char in enumerate(integer_str):
            if char == '0':
                zero_flag = True
            else:
                if zero_flag:
                    result += nums[0]  # 添加“零”
                    zero_flag = False
                result += nums[int(char)] + units[length - i - 1]

        # 处理小数部分
        decimal_part = round(amount - integer_part, 2)
        decimal_str = ""
        if decimal_part > 0:
            decimal_str = "".join([
                nums[int(d)] + u 
                for d, u in zip(f"{decimal_part:.2f}".split('.')[1], decimal_units)
                if d != '0'
            ])

        # 拼接结果
        if not result:
            result = "零"
        if not decimal_str:
            decimal_str = "整"

        return result + "元" + decimal_str

if __name__ == "__main__":
    root = tk.Tk()
    app = QuotationApp(root)
    print("程序已启动，等待用户交互...")
    root.mainloop()