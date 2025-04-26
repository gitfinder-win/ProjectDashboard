import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
from matplotlib.gridspec import GridSpec
import os
import datetime
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
from matplotlib.widgets import Cursor
from data_processor import DataProcessor
import threading
import time
import platform

# Set Chinese font
# 设置更美观的中文字体和全局样式
plt.rcParams['font.sans-serif'] = ['Microsoft YaHei']  # 使用更现代的中文字体
plt.rcParams['axes.unicode_minus'] = False
plt.rcParams['font.size'] = 10  # 统一基础字号

class ProjectDashboard:
    def __init__(self, master):
        self.master = master
        self.master.title("项目看板")
        self.master.geometry("1280x720")
        self.master.configure(bg="#0a1a35")  # 使用更柔和的深蓝色背景
        
        # Data attributes
        self.data_processor = DataProcessor()
        self.current_year = datetime.datetime.now().year
        self.annotations = []  # 存储标注对象用于后续管理
        self.temp_annotations = []  # 存储临时悬停注释
        self.fixed_annotations = []  # 存储固定点击注释
        self.dept_bars = {}  # 初始化部门柱状图字典
        self.highlight_active = False  # 自动高亮状态
        self.highlight_index = 0  # 当前高亮部门索引
        self.highlight_timer = None  # 高亮定时器
        self.highlight_duration = 5  # 高亮显示时长(秒)
        
        # 防止休眠相关变量
        self.keep_awake_active = False  # 保持显示状态
        self.keep_awake_thread = None   # 保持显示线程
        self.scheduled_display = False  # 定时显示状态
        self.schedule_start_hour = 8    # 默认开始时间 8:00
        self.schedule_start_minute = 0
        self.schedule_end_hour = 18     # 默认结束时间 18:00
        self.schedule_end_minute = 0
        self.schedule_check_timer = None  # 检查时间段的定时器
        
        # Create UI
        self.create_widgets()
        
    def create_widgets(self):
        # Top frame for controls
        self.control_frame = tk.Frame(self.master, bg="#000720")
        self.control_frame.pack(fill=tk.X, padx=10, pady=10)
        
        # Load Excel button
        self.load_btn = tk.Button(self.control_frame, text="加载Excel数据", command=self.load_excel_file,
                                 bg="#4a89dc", fg="white", font=("Microsoft YaHei", 12, "bold"))
        self.load_btn.pack(side=tk.LEFT, padx=5)
        
        # Year selection
        self.year_label = tk.Label(self.control_frame, text="年份:", bg="#000720", fg="white", font=("SimHei", 12))
        self.year_label.pack(side=tk.LEFT, padx=5)
        
        self.year_var = tk.StringVar(value=str(self.current_year))
        self.year_entry = tk.Entry(self.control_frame, textvariable=self.year_var, width=6, font=("SimHei", 12))
        self.year_entry.pack(side=tk.LEFT, padx=5)
        
        # Update button
        self.update_btn = tk.Button(self.control_frame, text="更新看板", command=self.update_dashboard,
                                  bg="#FF9F45", fg="white", font=("SimHei", 12))
        self.update_btn.pack(side=tk.LEFT, padx=5)
        
        # Reset zoom button
        self.reset_zoom_btn = tk.Button(self.control_frame, text="重置缩放", command=self.reset_zoom,
                                     bg="#4a69bd", fg="white", font=("SimHei", 12))
        self.reset_zoom_btn.pack(side=tk.LEFT, padx=5)
        
        # 高亮持续时间设置
        self.duration_label = tk.Label(self.control_frame, text="高亮时长(秒):", bg="#000720", fg="white", font=("SimHei", 12))
        self.duration_label.pack(side=tk.LEFT, padx=5)
        
        self.duration_var = tk.StringVar(value=str(self.highlight_duration))
        self.duration_entry = tk.Entry(self.control_frame, textvariable=self.duration_var, width=3, font=("SimHei", 12))
        self.duration_entry.pack(side=tk.LEFT, padx=5)
        
        # Auto highlight toggle button
        self.highlight_btn = tk.Button(self.control_frame, text="开始高亮", command=self.toggle_highlight,
                                     bg="#37b24d", fg="white", font=("SimHei", 12))
        self.highlight_btn.pack(side=tk.LEFT, padx=5)
        
        # 添加分隔线
        separator = tk.Frame(self.control_frame, height=30, width=2, bg="gray")
        separator.pack(side=tk.LEFT, padx=10, fill=tk.Y)
        
        # 保持常显按钮
        self.keep_awake_btn = tk.Button(self.control_frame, text="保持常显", command=self.toggle_keep_awake,
                                     bg="#37b24d", fg="white", font=("SimHei", 12))
        self.keep_awake_btn.pack(side=tk.LEFT, padx=5)
        
        # 定时显示按钮
        self.schedule_btn = tk.Button(self.control_frame, text="定时显示", command=self.toggle_scheduled_display,
                                   bg="#37b24d", fg="white", font=("SimHei", 12))
        self.schedule_btn.pack(side=tk.LEFT, padx=5)
        
        # 设置时间段按钮
        self.time_settings_btn = tk.Button(self.control_frame, text="设置时间段", command=self.show_schedule_settings,
                                        bg="#4a89dc", fg="white", font=("SimHei", 12))
        self.time_settings_btn.pack(side=tk.LEFT, padx=5)
        
        # Canvas for showing the dashboard
        self.canvas_frame = tk.Frame(self.master, bg="#000720")
        self.canvas_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Create figure for plotting
        self.fig = plt.Figure(figsize=(12, 8), dpi=100, facecolor="#0a1a35")
        self.canvas = FigureCanvasTkAgg(self.fig, self.canvas_frame)
        self.canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
        
        # 添加matplotlib导航工具栏
        self.toolbar_frame = tk.Frame(self.master)
        self.toolbar_frame.pack(fill=tk.X)
        self.toolbar = NavigationToolbar2Tk(self.canvas, self.toolbar_frame)
        self.toolbar.config(background="#000720")
        self.toolbar._message_label.config(background="#000720", foreground="white")
        for button in self.toolbar.winfo_children():
            if isinstance(button, tk.Button):
                button.config(background="#000720", foreground="white")
                
        # 连接鼠标事件处理器
        self.fig.canvas.mpl_connect('motion_notify_event', self.on_hover)
        
        # Status bar
        self.status_var = tk.StringVar(value="就绪")
        self.status_bar = tk.Label(self.master, textvariable=self.status_var, 
                                 bd=1, relief=tk.SUNKEN, anchor=tk.W, bg="#000720", fg="white")
        self.status_bar.pack(fill=tk.X)
        
    def load_excel_file(self):
        file_path = filedialog.askopenfilename(title="选择Excel文件", 
                                             filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")])
        if not file_path:
            return
            
        try:
            self.status_var.set(f"正在加载: {os.path.basename(file_path)}")
            self.master.update()
            
            # Load and process data using DataProcessor
            if self.data_processor.load_excel(file_path):
                if self.data_processor.process_data():
                    self.status_var.set(f"数据已加载: {os.path.basename(file_path)}")
                    # 直接更新看板，移除成功提示
                    self.update_dashboard()
                else:
                    self.status_var.set("处理数据失败")
                    messagebox.showerror("错误", "处理Excel数据时出错")
            else:
                self.status_var.set("加载失败")
                messagebox.showerror("错误", "加载Excel文件失败")
            
        except Exception as e:
            self.status_var.set("加载失败")
            messagebox.showerror("错误", f"加载数据时出错: {str(e)}")
    
    def update_dashboard(self):
        if not hasattr(self.data_processor, 'processed_data') or not self.data_processor.processed_data:
            messagebox.showwarning("警告", "请先加载Excel数据")
            return
            
        try:
            year = int(self.year_var.get())
            self.data_processor.year = year
        except ValueError:
            messagebox.showerror("错误", "请输入有效年份")
            return
            
        self.status_var.set("正在更新看板...")
        self.master.update()
        
        # 清除之前的数据
        self.annotations = []
        self.dept_bars = {}  # 重置部门柱状图字典
        
        # Clear previous plots
        self.fig.clear()
        
        # Create title
        self.fig.suptitle(f"{year}年项目任务看板", fontsize=16, color="white", y=0.98)
        
        # Create grid for subplots
        gs = GridSpec(2, 1, figure=self.fig, height_ratios=[1, 1.5])
        
        # 1. Monthly completion rates trend chart (top)
        self.create_monthly_completion_chart(self.fig.add_subplot(gs[0, 0]))
        
        # 2. Department monthly metrics chart (bottom)
        self.create_department_monthly_metrics_chart(self.fig.add_subplot(gs[1, 0]))
        
        # Adjust layout
        self.fig.tight_layout(rect=[0, 0, 1, 0.95])
        self.canvas.draw()
        self.status_var.set("看板已更新")
    
    def create_monthly_completion_chart(self, ax):
        """Create a chart showing monthly completion rates for 4 departments"""
        ax.set_title("1-12月部门任务计划完成率趋势图", color="white", fontsize=12, pad=35)  # 增加 pad 值
        ax.set_facecolor("#101450")
        
        # Get monthly completion rates for 4 departments
        months, department_rates, department_names = self.data_processor.get_department_monthly_completion_rates(4)
        
        # Define colors for each department
        colors = ['#3a7ca5', '#d63031', '#00b894', '#fdcb6e']
        
        # Set up the plot area
        ax.set_ylim(0, 100)
        ax.set_xlim(0.5, len(months) + 0.5)
        
        # Convert month names to display format (1-12)
        month_numbers = [i+1 for i in range(len(months))]
        
        # Set up grid and ticks
        ax.grid(True, linestyle="--", alpha=0.3)
        ax.set_xticks(month_numbers)
        ax.set_xticklabels([str(num) for num in month_numbers])
        ax.set_xlabel("月份", color="white", fontsize=11)
        ax.set_ylabel("任务计划完成率 (%)", color="white", fontsize=11)
        
        # 添加80%水位线（不添加label和picker属性）
        ax.axhline(y=80, color='red', linestyle='--', linewidth=2, alpha=0.6, zorder=1)
        
        # 添加填充区域来区分80%以上和以下的区域
        ax.fill_between(month_numbers, 80, 100, color='green', alpha=0.1, zorder=0)
        ax.fill_between(month_numbers, 0, 80, color='red', alpha=0.1, zorder=0)
        
        # 存储部门平均完成率，用于图例
        dept_avg_rates = {}
        line_objects = {}
        
        # Plot lines for each department
        for i, (dept_name, rates) in enumerate(zip(department_names, department_rates)):
            color = colors[i % len(colors)]
            # Use only valid data points (not NaN) for plotting
            valid_months = []
            valid_rates = []
            
            for j, rate in enumerate(rates):
                if not np.isnan(rate):
                    valid_months.append(month_numbers[j])
                    valid_rates.append(rate)
            
            # Only plot if there are valid data points
            if valid_rates:
                # 计算平均完成率
                avg_rate = sum(valid_rates) / len(valid_rates)
                dept_avg_rates[dept_name] = avg_rate
                
                # 绘制折线图
                line, = ax.plot(valid_months, valid_rates, marker='o', color=color, linewidth=2, 
                               label=f"{dept_name} (总平均: {avg_rate:.2f}%)", picker=5)
                
                # 存储线对象属性
                line.dept_name = dept_name  # 确保设置正确的部门名称
                line.original_color = color
                line.original_linewidth = 2
                line.original_alpha = 1.0
                
                # 为每个数据点添加可悬停标记
                for x, y in zip(valid_months, valid_rates):
                    point, = ax.plot(x, y, 'o', color=color, markersize=8, alpha=0.7, picker=5)
                    # 将点和数据关联存储，用于悬停显示
                    point.dept_name = dept_name
                    point.x_value = x
                    point.y_value = y
                
                # Add department name labels directly at the end of each line with average rate
                if valid_months:
                    last_x = valid_months[-1]
                    last_y = valid_rates[-1]
                    ax.annotate(f"{dept_name}", 
                               xy=(last_x, last_y),
                               xytext=(last_x + 0.1, last_y),
                               color=color,
                               fontsize=8,
                               va='center',
                               zorder=5,
                               bbox=dict(boxstyle="round,pad=0.1", 
                                       fc="#101450", 
                                       ec="none",
                                       alpha=0.6))
        
        # 在图表上方添加图例
        legend = ax.legend(loc='upper center', 
                          bbox_to_anchor=(0.5, 1.25),
                          ncol=len(department_names), 
                          fontsize=9,
                          frameon=True,
                          facecolor='#101450',
                          edgecolor='white')
        
        # 设置图例文本颜色为白色
        for text in legend.get_texts():
            text.set_color('white')
        
        # Set text colors to white
        for text in ax.get_xticklabels() + ax.get_yticklabels():
            text.set_color("white")
        ax.spines['bottom'].set_color('white')
        ax.spines['top'].set_color('white') 
        ax.spines['right'].set_color('white')
        ax.spines['left'].set_color('white')
        ax.tick_params(axis='x', colors='white')
        ax.tick_params(axis='y', colors='white')
    
    def create_department_monthly_metrics_chart(self, ax):
        """Create a chart showing department monthly metrics (完成任务数, 输出物, 审签数)"""
        ax.set_title("部门月度任务指标统计", color="white", fontsize=10)
        ax.set_facecolor("#101450")
        
        # Get monthly metrics data for departments
        months, departments, metrics_data = self.data_processor.get_department_monthly_metrics()
        
        # 打印数据用于调试
        print("\n获取到的月度指标数据:")
        for dept in departments:
            print(f"{dept}:")
            for month in months:
                if dept in metrics_data and month in metrics_data[dept]:
                    metrics_values = metrics_data[dept][month]
                    print(f"  {month}: {metrics_values}")
                else:
                    print(f"  {month}: 无数据")
        
        # Convert month names to display format (1-12)
        month_numbers = [i+1 for i in range(len(months))]
        
        # Check if we have data
        if not metrics_data or not departments:
            ax.text(0.5, 0.5, "无可用数据", ha='center', va='center', color='white', fontsize=12)
            ax.axis('off')
            return
        
        # 计算每个部门的柱子宽度
        n_depts = len(departments)
        bar_width = 0.8 / n_depts  # 每个部门柱子的宽度
        
        # Define colors for each metric
        metric_colors = {
            "完成任务数": "#3a7ca5",  # Blue
            "输出物": "#d63031",      # Red
            "审签数": "#00b894"       # Green
        }
        
        # 存储每个部门各月份的柱状位置，用于添加部门标签
        dept_bar_positions = {}
        dept_bars = {}  # 存储每个部门的所有柱状图对象
        
        # 创建堆叠的柱状图
        for d_idx, dept in enumerate(departments):  # 正确的写法
            # 计算当前部门的x位置
            x_positions = [num - 0.4 + bar_width * d_idx + bar_width/2 for num in month_numbers]
            dept_bar_positions[dept] = x_positions
            dept_bars[dept] = []  # 初始化部门柱状图列表
            
            # 为每个月创建堆叠的柱状图
            bottom_values = [0] * len(months)  # 底部起始值
            
            # 绘制各指标的堆叠柱状图
            for metric in self.data_processor.metrics:
                # 收集各月份的数据
                values = []
                for month in months:
                    if dept in metrics_data and month in metrics_data[dept] and metric in metrics_data[dept][month]:
                        values.append(metrics_data[dept][month][metric])
                    else:
                        values.append(0)
                
                # 绘制条形图
                color = metric_colors.get(metric, "#ffffff")
                label = metric if d_idx == 0 else None  # 只为第一个部门添加指标图例
                bars = ax.bar(x_positions, values, bar_width * 0.9, 
                              bottom=bottom_values, color=color, picker=5, label=label)
                
                # 添加到部门柱状图列表
                dept_bars[dept].extend(bars)
                
                # 更新下一个指标的底部值
                bottom_values = [bottom + value for bottom, value in zip(bottom_values, values)]
                
                # 为每个柱状图段添加数据属性用于悬停和点击显示
                for i, bar in enumerate(bars):
                    bar.dept_name = dept
                    bar.metric_name = metric
                    bar.month_num = i + 1
                    bar.value = values[i]
                    bar.visible_annotation = False  # 标记是否固定显示注释
                    bar.is_highlighted = False  # 标记是否突出显示
        
        # 添加部门名称标签（优化位置到每月柱状图下方中心）
        for dept, positions in dept_bar_positions.items():
            for pos in positions:
                # 将标签左移，右对齐到柱状图右边缘
                label_x = pos + (bar_width * 0.45)  # 将标签右对齐到柱状图右边缘
                ax.text(label_x, -5, dept, ha='right', va='top', 
                       color="white", fontsize=8, rotation=45,
                       bbox=dict(boxstyle="round,pad=0.1", fc="#101450", ec="gray", alpha=0.6))
        
        # 添加图例（只显示指标）
        handles = []
        labels = []
        for metric, color in metric_colors.items():
            handle = plt.Rectangle((0,0), 1, 1, color=color, label=metric)
            handles.append(handle)
            labels.append(metric)
        
        ax.legend(handles=handles, labels=labels, loc='center left',  # 改为左侧位置
                  bbox_to_anchor=(-0.12, 0.5),  # 向左偏移
                  fontsize=8)
        
        # Set up grid and ticks
        ax.grid(True, linestyle="--", alpha=0.3)
        ax.set_xticks(month_numbers)
        ax.set_xticklabels([str(num) for num in month_numbers])
        ax.set_xlabel("月份", color="white", fontsize=9)
        ax.set_ylabel("数量", color="white", fontsize=9)
        
        # Set text colors to white
        for text in ax.get_xticklabels() + ax.get_yticklabels():
            text.set_color("white")
        ax.spines['bottom'].set_color('white')
        ax.spines['top'].set_color('white') 
        ax.spines['right'].set_color('white')
        ax.spines['left'].set_color('white')
        ax.tick_params(axis='x', colors='white')
        ax.tick_params(axis='y', colors='white')
        
        # 保存部门柱状图对象字典
        self.dept_bars = dept_bars
        
        # 连接鼠标事件
        self.fig.canvas.mpl_connect('button_press_event', self.on_click)
    
    def on_hover(self, event):
        """Handle hover event to show data on hover"""
        # 如果自动高亮功能正在运行，忽略鼠标悬停事件
        if self.highlight_active:
            return
            
        if event.inaxes is None:
            # 鼠标不在任何坐标轴内
            self._reset_all_line_styles()  # 重置所有线条样式
            self._reset_all_bar_highlights()  # 重置所有柱状图高亮
            self._remove_non_fixed_annotations()  # 移除非固定注释
            self.fig.canvas.draw_idle()
            return

        # 移除所有非固定注释
        self._remove_non_fixed_annotations()
        
        # 重置所有线条样式
        self._reset_all_line_styles()
        
        # 重置所有柱状图高亮
        self._reset_all_bar_highlights()
        
        if event.inaxes == self.fig.axes[0]:  # 趋势图
            # 检查是否悬停在线上
            for line in self.fig.axes[0].get_lines():
                if line.contains(event)[0]:
                    # 高亮线条
                    line.set_linewidth(3.0)
                    line.set_color('yellow')
                    
                    # 获取部门名称和数据
                    dept_name = getattr(line, 'dept_name', 'Unknown')
                    xdata, ydata = line.get_data()
                    xpos, ypos = event.xdata, event.ydata
                    distances = [(abs(x - xpos), i) for i, x in enumerate(xdata)]
                    closest_idx = min(distances, key=lambda x: x[0])[1]
                    month = int(xdata[closest_idx])
                    rate = ydata[closest_idx]
                    
                    # 创建临时注释，添加zorder确保显示在最上层
                    # 为12月份特殊处理文本位置，避免标签超出图表
                    if month == 12:
                        xytext = (-60, 10)  # 向左偏移文本位置
                    else:
                        xytext = (10, 10)
                        
                    annotation = self.fig.axes[0].annotate(
                        f"{dept_name}: {month}月 {rate:.1f}%",
                        xy=(xdata[closest_idx], ydata[closest_idx]),
                        xytext=xytext,
                        textcoords="offset points",
                        bbox=dict(boxstyle="round,pad=0.3", fc="yellow", ec="b", alpha=0.8),
                        color='black',
                        fontsize=9,
                        zorder=1000  # 确保显示在最上层
                    )
                    
                    self.temp_annotations.append(annotation)
                    break
        
        elif event.inaxes == self.fig.axes[1]:  # 柱状图
            # 检查是否悬停在柱子上
            for dept_name, bars in self.dept_bars.items():
                for bar in bars:
                    if bar.contains(event)[0]:
                        # 高亮部门的所有柱状图
                        self._highlight_department_bars(dept_name)
                        
                        # 获取数据用于注释
                        month_num = getattr(bar, 'month_num', 0)
                        metric_name = getattr(bar, 'metric_name', 'Unknown')
                        value = getattr(bar, 'value', 0)
                        
                        # 创建临时注释，添加zorder确保显示在最上层
                        annotation = self.fig.axes[1].annotate(
                            f"{dept_name}: {month_num}月 {metric_name} {value}",
                            xy=(bar.get_x() + bar.get_width()/2, bar.get_y() + bar.get_height()),
                            xytext=(0, 10),
                            textcoords="offset points",
                            ha='center',
                            bbox=dict(boxstyle="round,pad=0.3", fc="yellow", ec="b", alpha=0.8),
                            color='black',
                            fontsize=9,
                            zorder=1000  # 确保显示在最上层
                        )
                        self.temp_annotations.append(annotation)
                        
                        # 找到了包含事件的柱状图，退出循环
                        break
                else:
                    # 内层循环没有中断，继续下一个部门
                    continue
                # 内层循环中断，说明找到了匹配的部门，跳出外层循环
                break
        
        # 重绘图形
        self.fig.canvas.draw_idle()
        
    def on_click(self, event):
        """Handle click event to fix annotations"""
        # 如果自动高亮功能正在运行，忽略鼠标点击事件
        if self.highlight_active:
            return
            
        if event.inaxes is None:
            # 点击在图表外部，重置所有效果
            self._reset_all_bar_highlights()
            self._reset_all_line_styles()
            self._remove_non_fixed_annotations()
            self.fig.canvas.draw_idle()
            return
            
        if event.inaxes == self.fig.axes[1]:  # 柱状图区域
            clicked_bar = False
            
            # 检查是否点击在柱子上
            for dept_name, bars in self.dept_bars.items():
                for bar in bars:
                    if bar.contains(event)[0]:
                        clicked_bar = True
                        break
                if clicked_bar:
                    break
            
            if not clicked_bar:
                # 点击在柱状图区域但未点中柱子，重置效果
                for dept_bars in self.dept_bars.values():
                    for bar in dept_bars:
                        bar.set_alpha(1.0)  # 恢复所有柱子的透明度
                        bar.is_highlighted = False
                self._remove_non_fixed_annotations()
                self.fig.canvas.draw_idle()
                return
    
        if event.inaxes == self.fig.axes[0]:  # 趋势图
            # 检查是否点击了线
            for line in self.fig.axes[0].get_lines():
                if line.contains(event)[0]:
                    # 获取部门名称和数据
                    dept_name = getattr(line, 'dept_name', 'Unknown')
                    xdata, ydata = line.get_data()
                    
                    # 查找最接近的数据点
                    xpos, ypos = event.xdata, event.ydata
                    distances = [(abs(x - xpos), i) for i, x in enumerate(xdata)]
                    closest_idx = min(distances, key=lambda x: x[0])[1]
                    month = int(xdata[closest_idx])
                    
                    # 检查是否已经有固定注释
                    for annotation in self.fixed_annotations:
                        if hasattr(annotation, 'dept_name') and annotation.dept_name == dept_name and \
                           hasattr(annotation, 'month') and annotation.month == month:
                            # 如果已经有相同的固定注释，则移除它
                            annotation.remove()
                            self.fixed_annotations.remove(annotation)
                            self.fig.canvas.draw_idle()
                            return
                    
                    # 创建固定注释，添加zorder确保显示在最上层
                    rate = ydata[closest_idx]
                    
                    # 为12月份特殊处理文本位置，避免标签超出图表
                    if month == 12:
                        xytext = (-60, 10)  # 向左偏移文本位置
                        arrow_props = dict(arrowstyle="-", color="yellow", alpha=0.8, connectionstyle="arc3,rad=-0.2")
                    else:
                        xytext = (10, 10)
                        arrow_props = dict(arrowstyle="-", color="yellow", alpha=0.8)
                    
                    annotation = self.fig.axes[0].annotate(
                        f"{dept_name}: {month}月 {rate:.1f}%",
                        xy=(xdata[closest_idx], ydata[closest_idx]),
                        xytext=xytext,
                        textcoords="offset points",
                        bbox=dict(boxstyle="round,pad=0.3", fc="yellow", ec="b", alpha=0.8),
                        color='black',
                        fontsize=9,
                        zorder=1000,  # 确保显示在最上层
                        arrowprops=arrow_props
                    )
                    # 添加属性以便后续识别
                    annotation.dept_name = dept_name
                    annotation.month = month
                    annotation.visible_annotation = True
                    
                    # 添加到固定注释列表
                    self.fixed_annotations.append(annotation)
                    break
                    
        elif event.inaxes == self.fig.axes[1]:  # 柱状图
            # 检查是否点击了柱状图
            for dept_name, bars in self.dept_bars.items():
                found_bar = False
                for bar in bars:
                    if bar.contains(event)[0]:
                        # 获取数据
                        month_num = getattr(bar, 'month_num', 0)
                        metric_name = getattr(bar, 'metric_name', 'Unknown')
                        value = getattr(bar, 'value', 0)
                        
                        # 检查是否已有固定注释
                        for annotation in self.fixed_annotations:
                            if hasattr(annotation, 'dept_name') and annotation.dept_name == dept_name and \
                               hasattr(annotation, 'month_num') and annotation.month_num == month_num and \
                               hasattr(annotation, 'metric_name') and annotation.metric_name == metric_name:
                                # 如果已有相同注释，则移除
                                annotation.remove()
                                self.fixed_annotations.remove(annotation)
                                found_bar = True
                                break
                        
                        if not found_bar:
                            # 创建固定注释，添加zorder确保显示在最上层
                            annotation = self.fig.axes[1].annotate(
                                f"{dept_name}: {month_num}月 {metric_name} {value}",
                                xy=(bar.get_x() + bar.get_width()/2, bar.get_y() + bar.get_height()),
                                xytext=(0, 10),
                                textcoords="offset points",
                                ha='center',
                                bbox=dict(boxstyle="round,pad=0.3", fc="yellow", ec="b", alpha=0.8),
                                color='black',
                                fontsize=9,
                                zorder=1000  # 确保显示在最上层
                            )
                            # 添加属性
                            annotation.dept_name = dept_name
                            annotation.month_num = month_num
                            annotation.metric_name = metric_name
                            annotation.visible_annotation = True
                            
                            # 添加到固定注释列表
                            self.fixed_annotations.append(annotation)
                        
                        # 高亮部门的所有柱状图
                        self._highlight_department_bars(dept_name)
                        found_bar = True
                        break
                
                if found_bar:
                    break
        
        # 重绘图形
        self.fig.canvas.draw_idle()
        
    def _remove_non_fixed_annotations(self):
        """移除所有非固定注释"""
        for annotation in self.temp_annotations:
            if annotation in self.fig.axes[0].texts or annotation in self.fig.axes[1].texts:
                annotation.remove()
        self.temp_annotations = []
    
    def reset_zoom(self):
        """重置所有子图的缩放状态"""
        for ax in self.fig.axes:
            ax.set_autoscale_on(True)
            ax.relim()
            ax.autoscale_view()
        
        # 设置完成率图表的y轴范围为0-100
        if len(self.fig.axes) > 0:
            self.fig.axes[0].set_ylim(0, 100)
            
        self.canvas.draw()
        self.status_var.set("已重置缩放")
        
    def toggle_highlight(self):
        """Toggle automatic department highlighting"""
        self.highlight_active = not self.highlight_active
        
        if self.highlight_active:
            self.highlight_btn.config(text="停止高亮", bg="#e03131")
            self.start_highlight_cycle()
        else:
            self.highlight_btn.config(text="开始高亮", bg="#37b24d")
            self.stop_highlight_cycle()
            self.reset_highlight()
    
    def start_highlight_cycle(self):
        """Start cycling through departments with highlighting"""
        if not hasattr(self.data_processor, 'departments') or not self.data_processor.departments:
            return
            
        self.highlight_index = 0
        self.highlight_next_department()
    
    def highlight_next_department(self):
        """Highlight the next department in sequence"""
        if not self.highlight_active:
            return
            
        departments = self.data_processor.departments
        if not departments:
            return
            
        # Reset previous highlight
        self.reset_highlight()
        
        # Get current department
        dept = departments[self.highlight_index]
        
        # Highlight department in both charts
        self.highlight_department(dept)
        
        # Update index for next department
        self.highlight_index = (self.highlight_index + 1) % len(departments)
        
        # Get current duration setting
        try:
            duration = int(self.duration_var.get())
            if duration < 1:  # 确保至少为1秒
                duration = 1
            self.highlight_duration = duration
        except ValueError:
            # 如果输入无效，使用默认值
            duration = self.highlight_duration
        
        # Schedule next highlight with updated duration (convert to milliseconds)
        self.highlight_timer = self.master.after(duration * 1000, self.highlight_next_department)
    
    def stop_highlight_cycle(self):
        """Stop the automatic highlighting cycle"""
        if self.highlight_timer:
            self.master.after_cancel(self.highlight_timer)
            self.highlight_timer = None
    
    def reset_highlight(self):
        """Reset all highlights to default appearance"""
        for ax in self.fig.axes:
            for line in ax.lines:
                if hasattr(line, 'default_color'):
                    line.set_color(line.default_color)
                    line.set_linewidth(1.5)
                    
            for collection in ax.collections:
                if hasattr(collection, 'default_color'):
                    collection.set_color(collection.default_color)
                    collection.set_alpha(0.7)
        
        self.canvas.draw()
    
    def highlight_department(self, dept_name):
        """Highlight a specific department in all charts and show value labels"""
        # 移除所有非固定注释
        self._remove_non_fixed_annotations()
        
        # 重置所有线条样式
        self._reset_all_line_styles()
        
        # 重置所有柱状图高亮
        self._reset_all_bar_highlights()
        
        # 临时存储需要添加的注释
        temp_annotations = []
        
        # 1. 处理趋势图中的线条
        if len(self.fig.axes) > 0:
            ax0 = self.fig.axes[0]
            trend_annotations = []  # 保存趋势图的注释以便后续调整位置
            
            for line in ax0.get_lines():
                if hasattr(line, 'dept_name') and line.dept_name == dept_name:
                    # 高亮线条
                    line.set_linewidth(3.0)
                    line.set_color('yellow')
                    
                    # 获取数据用于注释
                    xdata, ydata = line.get_data()
                    
                    # 查找最新月份（最大的月份值）
                    valid_months = [int(x) for x, y in zip(xdata, ydata) if not np.isnan(y) and y > 0]
                    latest_month = max(valid_months) if valid_months else 0
                    
                    # 为每个有效数据点添加注释
                    for i, (x, y) in enumerate(zip(xdata, ydata)):
                        if not np.isnan(y) and y > 0:  # 只显示有效且大于0的值
                            month = int(x)
                            # 创建注释但暂不添加
                            # 只有最新月份显示完整部门名称，其它月份只显示数值
                            if month == latest_month:
                                text = f"{dept_name}: {month}月 {y:.1f}%"
                            else:
                                text = f"{month}月: {y:.1f}%"
                                
                            annotation = {
                                'x': x,
                                'y': y,
                                'text': text,
                                'month': month
                            }
                            trend_annotations.append(annotation)
            
            # 按月份排序注释，以便更好地调整布局
            trend_annotations.sort(key=lambda a: a['month'])
            
            # 调整注释位置以避免重叠 - 避免重复标签
            if trend_annotations:
                # 将注释分组为高值和低值
                high_values = [a for a in trend_annotations if a['y'] > 60]
                low_values = [a for a in trend_annotations if a['y'] <= 60]
                
                # 函数：创建交错的注释
                def create_staggered_annotations(annotations, base_y_offset, step, direction=1):
                    # 每个月份只保留一个注释，避免重复
                    unique_months = {}
                    for anno in annotations:
                        month = anno['month']
                        # 如果月份已存在，只保留值较大的那个
                        if month in unique_months:
                            if anno['y'] > unique_months[month]['y']:
                                unique_months[month] = anno
                        else:
                            unique_months[month] = anno
                    
                    # 对唯一的月份注释排序并创建
                    sorted_annos = sorted(unique_months.values(), key=lambda a: a['month'])
                    for i, anno in enumerate(sorted_annos):
                        month = anno['month']
                        y_offset = base_y_offset + (i % 3) * step * direction
                        
                        # 为12月特殊处理位置，避免超出右边界
                        if month == 12:
                            x_offset = -60  # 向左偏移，避免超出右边界
                            arrow_props = dict(arrowstyle="-", color="yellow", alpha=0.8, connectionstyle="arc3,rad=-0.2")
                        else:
                            x_offset = 10
                            arrow_props = dict(arrowstyle="-", color="yellow", alpha=0.8)
                        
                        annotation = ax0.annotate(
                            anno['text'],
                            xy=(anno['x'], anno['y']),
                            xytext=(x_offset, y_offset),
                            textcoords="offset points",
                            bbox=dict(boxstyle="round,pad=0.3", fc="yellow", ec="b", alpha=0.8),
                            color='black',
                            fontsize=9,
                            zorder=1000,
                            arrowprops=arrow_props
                        )
                        temp_annotations.append(annotation)
                
                # 处理高值和低值注释
                create_staggered_annotations(high_values, 10, 20, 1)  # 高值向上交错
                create_staggered_annotations(low_values, -25, 20, -1)  # 低值向下交错
        
        # 2. 处理柱状图
        if len(self.fig.axes) > 1 and hasattr(self, 'dept_bars'):
            ax1 = self.fig.axes[1]
            
            # 高亮部门的所有柱状图
            self._highlight_department_bars(dept_name)
            
            # 查找最新月份
            latest_month = 0
            if dept_name in self.dept_bars:
                valid_months = []
                for bar in self.dept_bars[dept_name]:
                    month_num = getattr(bar, 'month_num', 0)
                    value = getattr(bar, 'value', 0)
                    if value > 0:  # 只考虑有值的月份
                        valid_months.append(month_num)
                if valid_months:
                    latest_month = max(valid_months)
            
            # 为部门的每个柱状图添加注释 (仅最近一个月)
            if dept_name in self.dept_bars and latest_month > 0:
                # 按月份分组收集注释数据
                bar_annotations = []
                
                for bar in self.dept_bars[dept_name]:
                    # 获取数据用于注释
                    month_num = getattr(bar, 'month_num', 0)
                    metric_name = getattr(bar, 'metric_name', 'Unknown')
                    value = getattr(bar, 'value', 0)
                    
                    # 跳过非最近月份或值为0的柱状图
                    if month_num != latest_month or value == 0:
                        continue
                    
                    # 只添加最近一个月的注释
                    bar_annotations.append({
                        'x': bar.get_x() + bar.get_width()/2,
                        'y': bar.get_y() + bar.get_height(),
                        'text': f"{dept_name}: {month_num}月 {metric_name} {value}",
                        'metric': metric_name
                    })
                
                # 为最近一个月份创建注释，避免重叠
                for i, anno in enumerate(bar_annotations):
                    y_offset = 10 + i * 18  # 垂直交错
                    x_offset = 0
                    
                    # 判断是否需要使用连接线
                    use_arrow = (len(bar_annotations) > 1)
                    
                    arrow_props = dict(arrowstyle="-", color="yellow", alpha=0.8) if use_arrow else None
                    
                    annotation = ax1.annotate(
                        anno['text'],
                        xy=(anno['x'], anno['y']),
                        xytext=(x_offset, y_offset),
                        textcoords="offset points",
                        ha='center',
                        bbox=dict(boxstyle="round,pad=0.3", fc="yellow", ec="b", alpha=0.8),
                        color='black',
                        fontsize=9,
                        zorder=1000,
                        arrowprops=arrow_props
                    )
                    temp_annotations.append(annotation)
        
        # 存储临时注释
        self.temp_annotations = temp_annotations
        
        # 重绘图形
        self.fig.canvas.draw_idle()

    def _reset_all_line_styles(self):
        """重置所有趋势线的样式"""
        if not hasattr(self, 'fig') or not self.fig.axes:
            return
            
        for line in self.fig.axes[0].get_lines():
            if hasattr(line, 'original_color'):
                line.set_color(line.original_color)
                line.set_linewidth(line.original_linewidth)
            elif hasattr(line, 'dept_name'):
                dept_idx = list(self.data_processor.departments).index(line.dept_name) \
                    if line.dept_name in self.data_processor.departments else 0
                line.set_color(['#3a7ca5', '#d63031', '#00b894', '#fdcb6e'][dept_idx % 4])
  
    def _reset_all_bar_highlights(self):
        """重置所有柱状图的高亮状态"""
        if not hasattr(self, 'dept_bars'):
            return
            
        # 恢复所有柱状图的原始透明度
        for dept_bars in self.dept_bars.values():
            for bar in dept_bars:
                if hasattr(bar, 'is_highlighted') and bar.is_highlighted:
                    bar.set_alpha(1.0)  # 恢复原始透明度
                    bar.is_highlighted = False

    def _highlight_department_bars(self, dept_name):
        """高亮显示指定部门的所有柱状图"""
        # 降低所有柱状图的透明度
        for d_name, bars in self.dept_bars.items():
            for bar in bars:
                if d_name != dept_name:
                    bar.set_alpha(0.3)  # 降低非目标部门的透明度
                else:
                    bar.set_alpha(1.0)  # 确保目标部门完全不透明
                    bar.is_highlighted = True  # 标记为高亮状态

    # Add method to prevent computer sleep based on platform
    def prevent_sleep(self):
        """Prevent computer from sleeping based on platform"""
        while self.keep_awake_active:
            # Move mouse slightly each minute to prevent sleep
            if platform.system() == "Windows":
                try:
                    # On Windows, simulate keyboard input
                    import ctypes
                    ctypes.windll.kernel32.SetThreadExecutionState(0x80000002)  # ES_CONTINUOUS | ES_SYSTEM_REQUIRED
                except Exception as e:
                    print(f"无法阻止系统休眠: {e}")
            elif platform.system() == "Darwin":  # macOS
                try:
                    os.system("caffeinate -i -t 60 &")
                except:
                    pass
            elif platform.system() == "Linux":
                try:
                    os.system("xdg-screensaver reset")
                except:
                    pass
                
            # 检查是否在允许显示的时间段内
            if self.scheduled_display:
                current_time = datetime.datetime.now().time()
                start_time = datetime.time(hour=self.schedule_start_hour, minute=self.schedule_start_minute)
                end_time = datetime.time(hour=self.schedule_end_hour, minute=self.schedule_end_minute)
                
                # 如果当前时间不在设定范围内，且高亮处于激活状态，则停止高亮
                if not (start_time <= current_time <= end_time):
                    if self.highlight_active:
                        # 在tk主线程中停止高亮
                        self.master.after(0, self.stop_highlight_cycle)
                # 如果当前时间在设定范围内，且高亮未激活，则启动高亮
                elif start_time <= current_time <= end_time:
                    if not self.highlight_active:
                        # 在tk主线程中启动高亮
                        self.master.after(0, self.start_highlight_cycle)
            
            # 短暂休眠，避免过度占用CPU
            time.sleep(30)  # 每30秒检查一次
            
    def toggle_keep_awake(self):
        """Toggle keep awake functionality"""
        self.keep_awake_active = not self.keep_awake_active
        
        if self.keep_awake_active:
            self.keep_awake_btn.config(text="停止常显", bg="#e03131")
            # Start thread to prevent sleep
            self.keep_awake_thread = threading.Thread(target=self.prevent_sleep, daemon=True)
            self.keep_awake_thread.start()
        else:
            self.keep_awake_btn.config(text="保持常显", bg="#37b24d")
            # Thread will exit on its own when self.keep_awake_active becomes False
            
            # 如果定时显示也被关闭，则停止高亮循环
            if self.scheduled_display and self.highlight_active:
                self.stop_highlight_cycle()
                
    def toggle_scheduled_display(self):
        """Toggle scheduled display functionality"""
        self.scheduled_display = not self.scheduled_display
        
        if self.scheduled_display:
            self.schedule_btn.config(text="取消定时", bg="#e03131")
            # Start checking if we're in the scheduled time window
            self.check_schedule()
        else:
            self.schedule_btn.config(text="定时显示", bg="#37b24d")
            # If keep_awake is not active, stop the highlight cycle
            if not self.keep_awake_active and self.highlight_active:
                self.stop_highlight_cycle()
    
    def check_schedule(self):
        """Check if current time is within scheduled display window"""
        if not self.scheduled_display:
            return
            
        current_time = datetime.datetime.now().time()
        start_time = datetime.time(hour=self.schedule_start_hour, minute=self.schedule_start_minute)
        end_time = datetime.time(hour=self.schedule_end_hour, minute=self.schedule_end_minute)
        
        # Update display based on current time
        if start_time <= current_time <= end_time:
            # Start highlight if not already active
            if not self.highlight_active:
                self.start_highlight_cycle()
        else:
            # Stop highlight if active
            if self.highlight_active:
                self.stop_highlight_cycle()
        
        # Schedule next check in 1 minute
        self.schedule_check_timer = self.master.after(60000, self.check_schedule)
        
    def show_schedule_settings(self):
        """Show a dialog to set the scheduled display time window"""
        # Create toplevel window
        settings_window = tk.Toplevel(self.master)
        settings_window.title("设置显示时间段")
        settings_window.geometry("350x200")
        settings_window.configure(bg="#0a1a35")
        settings_window.resizable(False, False)
        settings_window.grab_set()  # Make it modal
        
        # Create frame for settings
        settings_frame = tk.Frame(settings_window, bg="#0a1a35")
        settings_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # Start time settings
        start_label = tk.Label(settings_frame, text="开始时间:", bg="#0a1a35", fg="white", font=("SimHei", 12))
        start_label.grid(row=0, column=0, padx=5, pady=10, sticky="w")
        
        start_hour_var = tk.StringVar(value=str(self.schedule_start_hour))
        start_hour = ttk.Spinbox(settings_frame, from_=0, to=23, width=3, textvariable=start_hour_var)
        start_hour.grid(row=0, column=1, padx=5, pady=10)
        
        start_hour_label = tk.Label(settings_frame, text="时", bg="#0a1a35", fg="white", font=("SimHei", 12))
        start_hour_label.grid(row=0, column=2, padx=2, pady=10)
        
        start_min_var = tk.StringVar(value=str(self.schedule_start_minute))
        start_min = ttk.Spinbox(settings_frame, from_=0, to=59, width=3, textvariable=start_min_var)
        start_min.grid(row=0, column=3, padx=5, pady=10)
        
        start_min_label = tk.Label(settings_frame, text="分", bg="#0a1a35", fg="white", font=("SimHei", 12))
        start_min_label.grid(row=0, column=4, padx=2, pady=10)
        
        # End time settings
        end_label = tk.Label(settings_frame, text="结束时间:", bg="#0a1a35", fg="white", font=("SimHei", 12))
        end_label.grid(row=1, column=0, padx=5, pady=10, sticky="w")
        
        end_hour_var = tk.StringVar(value=str(self.schedule_end_hour))
        end_hour = ttk.Spinbox(settings_frame, from_=0, to=23, width=3, textvariable=end_hour_var)
        end_hour.grid(row=1, column=1, padx=5, pady=10)
        
        end_hour_label = tk.Label(settings_frame, text="时", bg="#0a1a35", fg="white", font=("SimHei", 12))
        end_hour_label.grid(row=1, column=2, padx=2, pady=10)
        
        end_min_var = tk.StringVar(value=str(self.schedule_end_minute))
        end_min = ttk.Spinbox(settings_frame, from_=0, to=59, width=3, textvariable=end_min_var)
        end_min.grid(row=1, column=3, padx=5, pady=10)
        
        end_min_label = tk.Label(settings_frame, text="分", bg="#0a1a35", fg="white", font=("SimHei", 12))
        end_min_label.grid(row=1, column=4, padx=2, pady=10)
        
        # Function to save settings
        def save_settings():
            try:
                # Parse values
                start_h = int(start_hour_var.get())
                start_m = int(start_min_var.get())
                end_h = int(end_hour_var.get())
                end_m = int(end_min_var.get())
                
                # Validate ranges
                if 0 <= start_h <= 23 and 0 <= start_m <= 59 and 0 <= end_h <= 23 and 0 <= end_m <= 59:
                    # Save settings
                    self.schedule_start_hour = start_h
                    self.schedule_start_minute = start_m
                    self.schedule_end_hour = end_h
                    self.schedule_end_minute = end_m
                    
                    # If scheduled display is active, immediately check the schedule
                    if self.scheduled_display:
                        self.check_schedule()
                    
                    settings_window.destroy()
                else:
                    messagebox.showerror("错误", "请输入有效的时间范围（小时：0-23，分钟：0-59）")
            except ValueError:
                messagebox.showerror("错误", "请输入有效的整数时间值")
        
        # Save button
        save_btn = tk.Button(settings_frame, text="保存设置", 
                            command=save_settings,
                            bg="#4a89dc", fg="white", 
                            font=("SimHei", 12))
        save_btn.grid(row=2, column=0, columnspan=5, pady=20)

def main():
    root = tk.Tk()
    app = ProjectDashboard(root)
    
    # Add window close handler to clean up resources
    def on_closing():
        if app.keep_awake_active:
            app.keep_awake_active = False  # Stop the keep-awake thread
        
        if app.highlight_timer:
            root.after_cancel(app.highlight_timer)  # Cancel highlight timer
            
        if app.schedule_check_timer:
            root.after_cancel(app.schedule_check_timer)  # Cancel schedule check timer
            
        root.destroy()
        
    root.protocol("WM_DELETE_WINDOW", on_closing)
    root.mainloop()

if __name__ == "__main__":
    main()