import pandas as pd
import numpy as np
import re
from typing import Dict, List, Tuple, Any


class DataProcessor:
    """
    A class for processing Excel data for the project dashboard
    """
    
    def __init__(self):
        self.summary_data = None
        self.task_status_data = None
        self.departments = []
        self.months = []
        self.metrics = ["完成任务数", "输出物", "审签数"]
        self.year = None
        self.monthly_stats = {}
        self.processed_data = {}
        self.completion_data = {}
        
    def load_excel(self, file_path: str) -> bool:
        """
        Load data from Excel file
        
        Args:
            file_path: Path to the Excel file
            
        Returns:
            bool: True if successful, False otherwise
        """
        try:
            # Try to load both sheets if possible
            sheets_dict = pd.read_excel(file_path, sheet_name=None)
            
            # Check if required sheets exist
            if 'Summary' in sheets_dict and 'TaskStatus' in sheets_dict:
                self.summary_data = sheets_dict['Summary']
                self.task_status_data = sheets_dict['TaskStatus']
                print(f"Successfully loaded sheets: Summary ({self.summary_data.shape}) and TaskStatus ({self.task_status_data.shape})")
                return True
            else:
                # Find available sheets
                available_sheets = list(sheets_dict.keys())
                print(f"Required sheets 'Summary' and 'TaskStatus' not found. Available sheets: {available_sheets}")
                
                # Try to use the first sheet as Summary and second as TaskStatus if they exist
                if len(sheets_dict) >= 2:
                    self.summary_data = list(sheets_dict.values())[0]
                    self.task_status_data = list(sheets_dict.values())[1]
                    print(f"Using first sheet as Summary and second as TaskStatus")
                    return True
                elif len(sheets_dict) == 1:
                    # If only one sheet is available, use it as summary data
                    self.summary_data = list(sheets_dict.values())[0]
                    # Create an empty DataFrame for task status
                    self.task_status_data = pd.DataFrame()
                    print(f"Only one sheet found, using it as Summary data")
                    return True
                
                return False
        except Exception as e:
            print(f"Error loading Excel file: {str(e)}")
            self.summary_data = None
            self.task_status_data = None
            return False
    
    def process_data(self) -> bool:
        """
        Process the loaded Excel data
        
        Returns:
            bool: True if successful, False otherwise
        """
        if self.summary_data is None:
            return False
            
        try:
            # Process Summary data
            self._process_summary_data()
            
            # Process TaskStatus data if available
            if self.task_status_data is not None and not self.task_status_data.empty:
                self._process_task_status_data()
            
            # Set default year if not already set
            if self.year is None:
                current_year = pd.Timestamp.now().year
                self.year = current_year
            
            return True
        except Exception as e:
            print(f"Error processing data: {str(e)}")
            return False
    
    def _process_summary_data(self):
        """Process the Summary sheet data"""
        if self.summary_data is None or self.summary_data.empty:
            print("No Summary data to process")
            return
            
        # First, identify departments column
        dept_col = None
        for col in self.summary_data.columns:
            if isinstance(col, str) and '部门' in col:
                dept_col = col
                break
        
        if dept_col is None:
            # If no specific department column found, use the first column
            dept_col = self.summary_data.columns[0]
            
        # Extract departments
        self.departments = self.summary_data[dept_col].dropna().unique().tolist()
        print(f"Found departments: {self.departments}")
        
        # 月份及其对应的三个指标列的映射
        month_metrics_mapping = {}
        
        # 从截图可以看到Excel表格结构：每个月有固定的完成任务数、输出物、审签数三列
        # 首先找出所有月份列
        for col in self.summary_data.columns:
            col_str = str(col)
            month_match = re.search(r'(\d+)月', col_str)
            if month_match:
                month_num = int(month_match.group(1))
                month_name = f"{month_num}月"
                
                # 找到月份列后，它和它右边的两列应该是对应的三个指标
                if isinstance(col, int) or isinstance(col, str):
                    col_idx = list(self.summary_data.columns).index(col)
                    if col_idx + 2 < len(self.summary_data.columns):
                        # 从实际数据结构看，3列分别是完成任务数、输出物、审签数
                        metric_cols = [
                            self.summary_data.columns[col_idx],     # 完成任务数
                            self.summary_data.columns[col_idx + 1], # 输出物
                            self.summary_data.columns[col_idx + 2]  # 审签数
                        ]
                        month_metrics_mapping[month_num] = {
                            'month_name': month_name,
                            'metric_cols': metric_cols
                        }
        
        # 保存找到的月份
        sorted_months = sorted(month_metrics_mapping.keys())
        self.months = [f"{m}月" for m in sorted_months]
        
        print(f"Found months: {self.months}")
        print(f"Month metrics mapping: {month_metrics_mapping}")
        
        # 创建各部门各月份的指标数据结构
        self.processed_data = {}
        
        for dept in self.departments:
            self.processed_data[dept] = {}
            
            # 获取该部门的所有行
            dept_rows = self.summary_data[self.summary_data[dept_col] == dept]
            
            if dept_rows.empty:
                continue
            
            # 处理每个月的指标数据
            for month_num in sorted_months:
                month_info = month_metrics_mapping.get(month_num)
                if not month_info:
                    continue
                
                month_name = month_info['month_name']
                metric_cols = month_info['metric_cols']
                
                self.processed_data[dept][month_name] = {}
                
                # 将三个列映射到三个指标
                for i, metric in enumerate(self.metrics):
                    if i < len(metric_cols):
                        # 获取对应列的值
                        col = metric_cols[i]
                        if col in dept_rows.columns:
                            value = dept_rows[col].iloc[0]
                            if pd.isna(value):
                                value = 0
                            self.processed_data[dept][month_name][metric] = value
                        else:
                            self.processed_data[dept][month_name][metric] = 0
                    else:
                        self.processed_data[dept][month_name][metric] = 0
        
        # 输出提取到的指标数据用于调试
        print("\n提取到的部门月度指标数据:")
        for dept in self.departments:
            print(f"  {dept}:")
            for month in self.months:
                if dept in self.processed_data and month in self.processed_data[dept]:
                    print(f"    {month}: {self.processed_data[dept][month]}")
            
        # 计算月度汇总统计
        self._calculate_monthly_stats()
    
    def _process_task_status_data(self):
        """Process the TaskStatus sheet data to extract completion rates by department and month"""
        if self.task_status_data is None or self.task_status_data.empty:
            print("No TaskStatus data to process")
            return
            
        print("\n--- TaskStatus Data Analysis ---")
        print(f"TaskStatus data shape: {self.task_status_data.shape}")
        print("First few rows of TaskStatus data:")
        print(self.task_status_data.head())
            
        # Find department column
        dept_col = None
        for col in self.task_status_data.columns:
            if isinstance(col, str) and '部门' in col:
                dept_col = col
                break
        
        if dept_col is None:
            # If no specific department column found, use the first column
            dept_col = self.task_status_data.columns[0]
            
        print(f"Using '{dept_col}' as department column")
        if dept_col in self.task_status_data.columns:
            print(f"Department values in TaskStatus: {self.task_status_data[dept_col].unique()}")
        
        # Initialize completion data structure
        completion_data = {}
        for dept in self.departments:
            completion_data[dept] = {}
            # Initialize with NaN for all months (1-12)
            for month_num in range(1, 13):
                month_name = f"{month_num}月"
                completion_data[dept][month_name] = np.nan
        
        # Look for the '1~2月任务统计' column and related columns
        combined_columns = None
        for col in self.task_status_data.columns:
            if isinstance(col, str) and '1~2月' in col:
                combined_columns = col
                print(f"Found combined 1-2 month column: {combined_columns}")
                break
        
        # Find the completion rate column for 1-2月
        combined_completion_rate_col = None
        if combined_columns:
            # Check adjacent columns (up to 6 positions to the right) for the completion rate
            col_idx = list(self.task_status_data.columns).index(combined_columns)
            for i in range(col_idx, min(col_idx + 6, len(self.task_status_data.columns))):
                check_col = self.task_status_data.columns[i]
                if isinstance(check_col, str) and '计划任务完成率' in check_col:
                    combined_completion_rate_col = check_col
                    print(f"Found completion rate column for 1-2月: {combined_completion_rate_col}")
                    break
                # Also check for unnamed columns that might contain completion rate
                elif 'Unnamed:' in str(check_col):
                    # Check if the first row contains '计划任务完成率'
                    if i < len(self.task_status_data.columns) and len(self.task_status_data) > 0:
                        cell_value = self.task_status_data.iloc[0, i]
                        if isinstance(cell_value, str) and '计划任务完成率' in cell_value:
                            combined_completion_rate_col = check_col
                            print(f"Found completion rate column for 1-2月 in unnamed column: {combined_completion_rate_col}")
                            break
        
        # 部门完成率临时存储结构 {部门: {月份: [值1, 值2, ...], ...}, ...}
        dept_month_values = {}
        for dept in self.departments:
            dept_month_values[dept] = {}
            for month_num in range(1, 13):
                month_name = f"{month_num}月"
                dept_month_values[dept][month_name] = []
        
        # Process 1-2月 completion rates if found
        if combined_completion_rate_col:
            print(f"Processing 1-2月 completion rates from column: {combined_completion_rate_col}")
            for dept in self.departments:
                # 获取该部门的所有行
                dept_rows = self.task_status_data[self.task_status_data[dept_col] == dept]
                
                if dept_rows.empty or combined_completion_rate_col not in dept_rows.columns:
                    continue
                
                # 收集所有非NaN值
                valid_values = []
                for _, row in dept_rows.iterrows():
                    value = row[combined_completion_rate_col]
                    if not pd.isna(value):
                        # Convert to numeric if it's a string percentage
                        if isinstance(value, str) and '%' in value:
                            try:
                                value = float(value.strip('%'))
                            except:
                                print(f"  Could not convert '{value}' to float")
                                continue
                        elif isinstance(value, (int, float)):
                            # Ensure the value is a percentage
                            if value <= 1:
                                value = value * 100
                        valid_values.append(value)
                
                # 计算平均值
                if valid_values:
                    avg_value = sum(valid_values) / len(valid_values)
                    print(f"  Found {len(valid_values)} values for {dept}, average: {avg_value:.2f}%")
                    
                    # 将平均值添加到临时存储中
                    dept_month_values[dept]['1月'].append(avg_value)
                    dept_month_values[dept]['2月'].append(avg_value)
        
        # Process the remaining months (3-12月)
        for month_num in range(3, 13):
            month_name = f"{month_num}月"
            month_column = None
            
            # Find the column for this month
            for col in self.task_status_data.columns:
                if isinstance(col, str) and f"{month_num}月" in col:
                    month_column = col
                    print(f"Found column for {month_name}: {month_column}")
                    break
            
            if month_column:
                # Find completion rate column
                month_idx = list(self.task_status_data.columns).index(month_column)
                completion_col = None
                
                # Check adjacent columns for completion rate
                for i in range(month_idx, min(month_idx + 6, len(self.task_status_data.columns))):
                    check_col = self.task_status_data.columns[i]
                    if isinstance(check_col, str) and '计划任务完成率' in check_col:
                        completion_col = check_col
                        print(f"Found completion rate column for {month_name}: {completion_col}")
                        break
                    # Also check for unnamed columns
                    elif 'Unnamed:' in str(check_col):
                        # Check if the first row contains '计划任务完成率'
                        if i < len(self.task_status_data.columns) and len(self.task_status_data) > 0:
                            cell_value = self.task_status_data.iloc[0, i]
                            if isinstance(cell_value, str) and '计划任务完成率' in cell_value:
                                completion_col = check_col
                                print(f"Found completion rate column for {month_name} in unnamed column: {completion_col}")
                                break
                
                # Process completion rates for this month
                if completion_col:
                    print(f"Processing {month_name} completion rates from column: {completion_col}")
                    for dept in self.departments:
                        # 获取该部门的所有行
                        dept_rows = self.task_status_data[self.task_status_data[dept_col] == dept]
                        
                        if dept_rows.empty or completion_col not in dept_rows.columns:
                            continue
                        
                        # 收集所有非NaN值
                        valid_values = []
                        for _, row in dept_rows.iterrows():
                            value = row[completion_col]
                            if not pd.isna(value):
                                # Convert to numeric if it's a string percentage
                                if isinstance(value, str) and '%' in value:
                                    try:
                                        value = float(value.strip('%'))
                                    except:
                                        print(f"  Could not convert '{value}' to float")
                                        continue
                                elif isinstance(value, (int, float)):
                                    # Ensure the value is a percentage
                                    if value <= 1:
                                        value = value * 100
                                valid_values.append(value)
                        
                        # 计算平均值
                        if valid_values:
                            avg_value = sum(valid_values) / len(valid_values)
                            print(f"  Found {len(valid_values)} values for {dept}, average: {avg_value:.2f}%")
                            
                            # 将平均值添加到临时存储中
                            dept_month_values[dept][month_name].append(avg_value)
        
        # 计算最终平均值并存储到completion_data
        for dept in self.departments:
            for month_name, values in dept_month_values[dept].items():
                if values:  # 如果有值
                    completion_data[dept][month_name] = sum(values) / len(values)
        
        # Store the completion data
        self.completion_data = completion_data
        
        # 输出最终的部门月度完成率
        print("\n最终部门月度完成率:")
        for dept in self.departments:
            print(f"  {dept}:", end=" ")
            for month in range(1, 13):
                month_name = f"{month}月"
                if month_name in completion_data[dept] and not np.isnan(completion_data[dept][month_name]):
                    print(f"{month}月: {completion_data[dept][month_name]:.2f}%", end=", ")
            print()
        
        # Check if we have data for any month
        data_found = False
        for dept in self.departments:
            for month in range(1, 13):
                month_name = f"{month}月"
                if month_name in completion_data[dept] and not np.isnan(completion_data[dept][month_name]):
                    data_found = True
                    break
            if data_found:
                break
        
        # Print the summary of what we found
        if data_found:
            print("\nFound completion rate data:")
            for dept in self.departments:
                valid_months = 0
                print(f"  {dept}:", end=" ")
                for month in range(1, 13):
                    month_name = f"{month}月"
                    if month_name in completion_data[dept] and not np.isnan(completion_data[dept][month_name]):
                        print(f"{month}月: {completion_data[dept][month_name]:.2f}%", end=", ")
                        valid_months += 1
                print(f"(Total: {valid_months} months)")
        else:
            print("\nNo completion rate data found")
            print("\nSUGGESTIONS TO IMPROVE DATA EXTRACTION:")
            print("1. 确保Excel表格中有名为'部门'的列，并包含和Summary表相同的部门名称")
            print("2. 确保有形如'X月任务统计'的列标题，其中X是月份数字")
            print("3. 确保有包含'计划任务完成率'文本的列，或者在数据中有明显的完成率百分比")
            print("4. 确保完成率数据是数字格式，或带有%符号的文本")
            print("5. 如果使用合并单元格，确保在合并前填写了所有相关单元格")
    
    def _calculate_monthly_stats(self):
        """Calculate statistics for each month across departments"""
        if not self.processed_data:
            return
            
        for month in self.months:
            self.monthly_stats[month] = {metric: 0 for metric in self.metrics}
            
            for dept in self.departments:
                if dept in self.processed_data and month in self.processed_data[dept]:
                    for metric in self.metrics:
                        if metric in self.processed_data[dept][month]:
                            value = self.processed_data[dept][month][metric]
                            if pd.notna(value) and isinstance(value, (int, float)):
                                self.monthly_stats[month][metric] += value
    
    def get_department_monthly_completion_rates(self, num_departments=4) -> Tuple[List[str], List[List[float]], List[str]]:
        """
        Get the monthly completion rates for the top N departments
        
        Args:
            num_departments: Number of departments to include
            
        Returns:
            Tuple of:
            - Month names
            - List of completion rate lists for each department
            - Department names
        """
        if not hasattr(self, 'completion_data') or not self.completion_data:
            print("No completion rate data available")
            return self.months, [[50.0] * len(self.months)] * num_departments, self.departments[:num_departments]
        
        # Get top departments based on average completion rate (use all if we have fewer than requested)
        dept_avg_rates = {}
        for dept in self.departments:
            if dept in self.completion_data:
                rates = [self.completion_data[dept].get(month, np.nan) for month in self.months]
                # Filter out NaN values
                valid_rates = [r for r in rates if not np.isnan(r)]
                if valid_rates:
                    dept_avg_rates[dept] = sum(valid_rates) / len(valid_rates)
        
        # Sort departments by average rate and take top N
        sorted_depts = sorted(dept_avg_rates.items(), key=lambda x: x[1], reverse=True)
        top_depts = [dept for dept, _ in sorted_depts[:num_departments]]
        
        # If we don't have enough departments, use what we have
        if len(top_depts) < num_departments:
            remaining = [dept for dept in self.departments if dept not in top_depts]
            top_depts.extend(remaining[:num_departments - len(top_depts)])
        
        # Get completion rates for each selected department
        department_rates = []
        for dept in top_depts:
            rates = []
            for month in self.months:
                if dept in self.completion_data and month in self.completion_data[dept]:
                    rate = self.completion_data[dept][month]
                else:
                    rate = np.nan
                rates.append(rate)
            department_rates.append(rates)
        
        return self.months, department_rates, top_depts
    
    def get_department_monthly_metrics(self) -> Tuple[List[str], List[str], Dict[str, Dict[str, Dict[str, float]]]]:
        """
        Get monthly metrics (完成任务数, 输出物, 审签数) for each department
        
        Returns:
            Tuple of:
            - Month names
            - Department names
            - Dictionary of metrics by department and month
        """
        if not self.processed_data:
            print("No processed data available")
            return self.months, self.departments, {}
        
        return self.months, self.departments, self.processed_data 