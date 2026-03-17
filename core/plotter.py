import pandas as pd
import numpy as np
import datetime
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from tkinter import messagebox

# ==========================================
# SMPS 绘图逻辑
# ==========================================
def generate_smps_figure(filepath, x_label, y_label, text_info=None, density=1.30, font_size=12):
    try:
        if not filepath.endswith('.xlsx'):
            messagebox.showerror("Error", "Please select a .xlsx file.")
            return None

        df = pd.read_excel(filepath, header=None)
        header_row_idx, col_x_idx, col_y_idx = -1, -1, -1

        for idx, row in df.iterrows():
            row_strs = [str(val).strip() for val in row.values]
            found_start_time, found_total_conc = False, False
            for c_idx, val in enumerate(row_strs):
                if "start time" in val.lower():
                    found_start_time = True
                    temp_x_idx = c_idx
                if "Total Conc." in val: 
                    found_total_conc = True
                    temp_y_idx = c_idx
            if found_start_time and found_total_conc:
                header_row_idx, col_x_idx, col_y_idx = idx, temp_x_idx, temp_y_idx
                break
                
        if header_row_idx == -1:
            messagebox.showerror("Error", "Could not find 'Start Time' and 'Total Conc.' labels.")
            return None

        data_df = df.iloc[header_row_idx + 1:].copy()
        x_raw, y_raw = data_df.iloc[:, col_x_idx], data_df.iloc[:, col_y_idx]
        valid_mask = x_raw.notna() & y_raw.notna() & (x_raw != '') & (y_raw != '')
        x_raw, y_raw = x_raw[valid_mask], y_raw[valid_mask]

        y_data = pd.to_numeric(y_raw, errors='coerce')
        y_data = (y_data / 1000000000) * density 
        
        x_data = pd.to_datetime(x_raw.astype(str), format='%H:%M:%S', errors='coerce')

        final_mask = x_data.notna() & y_data.notna()
        x_data, y_data = x_data[final_mask], y_data[final_mask]

        if len(x_data) == 0:
            messagebox.showwarning("Warning", "No valid data found.")
            return None

        fig, ax = plt.subplots(figsize=(9, 6))
        ax.scatter(x_data, y_data, c='#1976D2', s=40, alpha=0.9, edgecolors='none')

        min_time, max_time = x_data.min().floor('h'), x_data.max().ceil('h')
        ax.set_xlim(min_time, max_time)
        ax.xaxis.set_major_locator(mdates.HourLocator())
        ax.xaxis.set_major_formatter(mdates.DateFormatter('%H:00'))

        ymin, ymax = ax.get_ylim()
        ax.set_ylim(ymin, ymax + (ymax - ymin) * 0.18) 
        
        # 动态应用字体大小
        ax.set_xlabel(x_label.replace('\\n', '\n'), fontsize=font_size, fontweight='bold')
        ax.set_ylabel(y_label.replace('\\n', '\n'), fontsize=font_size, fontweight='bold')
        
        # 刻度数字大小随之等比变化（比标题小1号）
        tick_size = max(8, font_size - 1)
        ax.tick_params(axis='both', which='major', labelsize=tick_size, direction='in', top=True, right=True)
        ax.grid(True, linestyle='--', alpha=0.6)

        if text_info and text_info[0] and text_info[1].strip():
            ax.text(0.03, 0.95, text_info[1].replace('\\n', '\n'), transform=ax.transAxes, 
                    color='#FF0000', fontsize=font_size, fontweight='bold', va='top', ha='left', zorder=10)

        plt.tight_layout()
        return fig
    except Exception as e:
        messagebox.showerror("Processing Error", str(e))
        return None

# ==========================================
# SMPS 质量积分计算逻辑
# ==========================================
def calculate_smps_mass(filepath, flow_rate_str, start_time_str, end_time_str, density=1.30):
    try:
        if not filepath.endswith('.xlsx'):
            messagebox.showerror("Error", "Please select a .xlsx file.")
            return None

        try:
            flow_rate = float(flow_rate_str)
            ref_date = datetime.datetime(2026, 1, 1) 
            t_start_input = datetime.datetime.strptime(start_time_str.strip(), "%H:%M")
            t_end_input = datetime.datetime.strptime(end_time_str.strip(), "%H:%M")
            
            target_start_sec = t_start_input.hour * 3600 + t_start_input.minute * 60
            target_end_sec = t_end_input.hour * 3600 + t_end_input.minute * 60
        except ValueError:
            messagebox.showerror("Input Error", "Check Flow Rate (number) and Time (HH:MM).")
            return None

        df = pd.read_excel(filepath, header=None)
        header_row_idx, col_x_idx, col_y_idx = -1, -1, -1
        for idx, row in df.iterrows():
            row_strs = [str(val).strip() for val in row.values]
            if any("start time" in s.lower() for s in row_strs) and any("Total Conc." in s for s in row_strs):
                header_row_idx = idx
                for c_idx, val in enumerate(row_strs):
                    if "start time" in val.lower(): col_x_idx = c_idx
                    if "Total Conc." in val: col_y_idx = c_idx
                break
        
        if header_row_idx == -1: return None

        data_df = df.iloc[header_row_idx + 1:].copy()
        x_raw = pd.to_datetime(data_df.iloc[:, col_x_idx].astype(str), format='%H:%M:%S', errors='coerce')
        y_raw = pd.to_numeric(data_df.iloc[:, col_y_idx], errors='coerce')
        
        valid_mask = x_raw.notna() & y_raw.notna()
        x_all = x_raw[valid_mask]
        
        y_all = (y_raw[valid_mask] / 1000000000) * density 
        
        all_secs = x_all.dt.hour * 3600 + x_all.dt.minute * 60 + x_all.dt.second

        y_start_interp = np.interp(target_start_sec, all_secs, y_all)
        y_end_interp = np.interp(target_end_sec, all_secs, y_all)

        middle_mask = (all_secs > target_start_sec) & (all_secs < target_end_sec)
        x_mid = all_secs[middle_mask].tolist()
        y_mid = y_all[middle_mask].tolist()
        
        final_x_secs = [target_start_sec] + x_mid + [target_end_sec]
        final_y_vals = [y_start_interp] + y_mid + [y_end_interp]
        
        x_minutes = np.array(final_x_secs) / 60.0
        y_values = np.array(final_y_vals)
        
        diff_x = np.diff(x_minutes)
        avg_y = (y_values[1:] + y_values[:-1]) / 2.0
        integral_val = np.sum(diff_x * avg_y)
        
        flow_m3_min = flow_rate * 0.001
        total_mass_ug = integral_val * flow_m3_min
        
        return total_mass_ug

    except Exception as e:
        messagebox.showerror("Calculation Error", str(e))
        return None

# ==========================================
# PTR 绘图逻辑
# ==========================================
def parse_ptr_file(filepath):
    try:
        if not filepath.endswith('.xlsx'):
            messagebox.showerror("Error", "Please select a .xlsx file.")
            return None
        df = pd.read_excel(filepath, header=None)
        header_row_idx = -1
        for idx, row in df.iterrows():
            if any("AbsTime" in str(val) for val in row.values):
                header_row_idx = idx
                break
        if header_row_idx == -1:
            messagebox.showerror("Error", "Could not find 'AbsTime'.")
            return None

        df.columns = df.iloc[header_row_idx]
        data_df = df.iloc[header_row_idx + 1:].copy()
        substance_cols = [c for c in data_df.columns if isinstance(c, str) and c.startswith('m')]
        
        x_raw = data_df['AbsTime']
        x_data = pd.to_datetime(x_raw, errors='coerce')
        valid_mask = x_data.notna()

        for col in substance_cols:
            data_df[col] = pd.to_numeric(data_df[col], errors='coerce')
            valid_mask &= data_df[col].notna()

        x_data, data_df = x_data[valid_mask], data_df[valid_mask]
        max_vals = {col: data_df[col].max() for col in substance_cols}
        return x_data, data_df, substance_cols, max_vals
    except Exception as e:
        messagebox.showerror("Processing Error", str(e))
        return None

def generate_ptr_figure_from_data(x_data, data_df, left_cols, right_cols, x_label, y_label_left, y_label_right, text_info=None, font_size=12):
    try:
        fig, ax1 = plt.subplots(figsize=(10, 7.5)) 
        ax2 = ax1.twinx()
        lines, labels = [], []
        distinct_colors = ['#4E79A7', '#F28E2B', '#E15759', '#76B7B2', '#59A14F', '#EDC949', '#AF7AA1', '#FF9DA7', '#9C755F', '#BAB0AB', '#396AB1', '#DA7C30', '#3E9651', '#CC2529', '#535154', '#6B4C9A', '#922428', '#948B3D', '#4C72B0', '#DD8452']

        color_idx = 0
        for col in left_cols:
            c = distinct_colors[color_idx % len(distinct_colors)]
            color_idx += 1
            line, = ax1.plot(x_data, data_df[col], linestyle='-', linewidth=1.5, label=col, color=c)
            lines.append(line); labels.append(col)

        for col in right_cols:
            c = distinct_colors[color_idx % len(distinct_colors)]
            color_idx += 1
            line, = ax2.plot(x_data, data_df[col], linestyle='-', linewidth=2.5, label=col + " (Right Axis)", color=c)
            lines.append(line); labels.append(col)

        min_time, max_time = x_data.min().floor('h'), x_data.max().ceil('h')
        ax1.set_xlim(min_time, max_time)
        ax1.xaxis.set_major_locator(mdates.HourLocator())
        ax1.xaxis.set_major_formatter(mdates.DateFormatter('%H:00'))

        ymin1, ymax1 = ax1.get_ylim()
        ax1.set_ylim(ymin1, ymax1 + (ymax1 - ymin1) * 0.18)
        ymin2, ymax2 = ax2.get_ylim()
        ax2.set_ylim(ymin2, ymax2 + (ymax2 - ymin2) * 0.18)

        # 动态应用字体大小
        ax1.set_xlabel(x_label.replace('\\n', '\n'), fontsize=font_size, fontweight='bold')
        ax1.set_ylabel(y_label_left.replace('\\n', '\n'), fontsize=font_size, fontweight='bold', color='black')
        ax2.set_ylabel(y_label_right.replace('\\n', '\n'), fontsize=font_size, fontweight='bold', color='black')

        tick_size = max(8, font_size - 1)
        ax1.grid(True, linestyle='--', alpha=0.6)
        ax1.tick_params(axis='both', which='major', labelsize=tick_size, direction='in')
        ax2.tick_params(axis='y', which='major', labelsize=tick_size, direction='in')

        if text_info and text_info[0] and text_info[1].strip():
            ax1.text(0.03, 0.95, text_info[1].replace('\\n', '\n'), transform=ax1.transAxes, color='#FF0000', fontsize=font_size, fontweight='bold', va='top', ha='left', zorder=10)

        fig.subplots_adjust(bottom=0.35)
        ax1.legend(lines, labels, loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=2, frameon=False, fontsize=9)
        return fig
    except Exception as e:
        messagebox.showerror("Plotting Error", str(e))
        return None

# ==========================================
# FTIR 绘图逻辑 
# ==========================================
def parse_ftir_file(filepath):
    try:
        if not filepath.endswith('.xlsx'):
            messagebox.showerror("Error", "Please select a .xlsx file.")
            return None

        try:
            df = pd.read_excel(filepath, header=None)
        except Exception as read_err:
            messagebox.showerror("Read Error", f"Cannot read Excel file. Is it open in another program?\n\n{str(read_err)}")
            return None
        
        header_row_idx = -1
        for idx, row in df.iterrows():
            if any(str(val).strip().lower() == 'local time' for val in row.values):
                header_row_idx = idx
                break

        if header_row_idx == -1:
            messagebox.showerror("Error", "Could not find 'local time' header in the file.")
            return None

        if header_row_idx + 1 >= len(df):
            messagebox.showerror("Data Structure Error", "Found 'local time' but no data below it!")
            return None

        headers = pd.Series([str(val).strip() for val in df.iloc[header_row_idx].values])
        multipliers = df.iloc[header_row_idx + 1]

        x_col_idx = (headers.str.lower() == 'local time').idxmax()
        
        substance_cols = []
        valid_col_indices = []
        mult_dict = {}
        
        for idx, col_name in enumerate(headers):
            col_name_lower = col_name.lower()
            if col_name_lower in ['local time', 'h2o', 'co2', 'nan', '']:
                continue
                
            try:
                val = multipliers.iloc[idx] 
                mult = float(val)
                if pd.notna(mult):
                    substance_cols.append(col_name)
                    valid_col_indices.append(idx)
                    mult_dict[col_name] = mult
            except (ValueError, TypeError):
                continue

        if not substance_cols:
            messagebox.showerror("Error", "No valid substances with numeric multipliers found.")
            return None

        data_start_idx = header_row_idx + 2
        if data_start_idx >= len(df):
            messagebox.showerror("Data Structure Error", "Found multipliers, but no time-series data below them. Please check the file.")
            return None

        data_df_raw = df.iloc[data_start_idx:].copy()
        
        x_raw = data_df_raw.iloc[:, x_col_idx]
        x_data = pd.to_datetime(x_raw, errors='coerce')
        valid_mask = x_data.notna()

        data_df = pd.DataFrame()
        for col_name, col_idx in zip(substance_cols, valid_col_indices):
            raw_vals = pd.to_numeric(data_df_raw.iloc[:, col_idx], errors='coerce')
            data_df[col_name] = raw_vals * mult_dict[col_name]
            valid_mask &= data_df[col_name].notna()

        x_data = x_data[valid_mask]
        data_df = data_df[valid_mask]

        if len(x_data) == 0:
            messagebox.showwarning("Warning", "No valid numeric data found below the multipliers.")
            return None

        max_vals = {col: data_df[col].max() for col in substance_cols}

        return x_data, data_df, substance_cols, max_vals

    except Exception as e:
        import traceback
        print(traceback.format_exc())
        messagebox.showerror("Processing Error", f"An error occurred:\n{str(e)}\n\nCheck terminal for full traceback.")
        return None

def generate_ftir_figure_from_data(x_data, data_df, left_cols, right_cols, x_label, y_label_left, y_label_right, text_info=None, font_size=12):
    try:
        fig, ax1 = plt.subplots(figsize=(10, 7.5)) 
        ax2 = ax1.twinx()

        lines = []
        labels = []

        distinct_colors = [
            '#4E79A7', '#F28E2B', '#E15759', '#76B7B2', '#59A14F', 
            '#EDC949', '#AF7AA1', '#FF9DA7', '#9C755F', '#BAB0AB',
            '#396AB1', '#DA7C30', '#3E9651', '#CC2529', '#535154'
        ]

        color_idx = 0

        for col in left_cols:
            c = distinct_colors[color_idx % len(distinct_colors)]
            color_idx += 1
            line, = ax1.plot(x_data, data_df[col], marker='o', markersize=4, linestyle='-', linewidth=1.5, label=col, color=c)
            lines.append(line)
            labels.append(col)

        for col in right_cols:
            c = distinct_colors[color_idx % len(distinct_colors)]
            color_idx += 1
            line, = ax2.plot(x_data, data_df[col], marker='o', markersize=5, linestyle='-', linewidth=2.5, label=col + " (Right Axis)", color=c)
            lines.append(line)
            labels.append(col)

        min_time = x_data.min().floor('h')
        max_time = x_data.max().ceil('h')
        ax1.set_xlim(min_time, max_time)
        ax1.xaxis.set_major_locator(mdates.HourLocator())
        ax1.xaxis.set_major_formatter(mdates.DateFormatter('%H:00'))

        ymin1, ymax1 = ax1.get_ylim()
        ax1.set_ylim(ymin1, ymax1 + (ymax1 - ymin1) * 0.18)
        
        ymin2, ymax2 = ax2.get_ylim()
        ax2.set_ylim(ymin2, ymax2 + (ymax2 - ymin2) * 0.18)

        # 动态应用字体大小
        ax1.set_xlabel(x_label.replace('\\n', '\n'), fontsize=font_size, fontweight='bold')
        ax1.set_ylabel(y_label_left.replace('\\n', '\n'), fontsize=font_size, fontweight='bold', color='black')
        ax2.set_ylabel(y_label_right.replace('\\n', '\n'), fontsize=font_size, fontweight='bold', color='black')

        tick_size = max(8, font_size - 1)
        ax1.grid(True, linestyle='--', alpha=0.6)
        ax1.tick_params(axis='both', which='major', labelsize=tick_size, direction='in')
        ax2.tick_params(axis='y', which='major', labelsize=tick_size, direction='in')

        if text_info and text_info[0] and text_info[1].strip():
            display_text = text_info[1].replace('\\n', '\n')
            ax1.text(0.03, 0.95, display_text, transform=ax1.transAxes, 
                    color='#FF0000', fontsize=font_size, fontweight='bold', va='top', ha='left', zorder=10)

        fig.subplots_adjust(bottom=0.35)
        ax1.legend(lines, labels, loc='upper center', bbox_to_anchor=(0.5, -0.15),
                   ncol=2, frameon=False, fontsize=9)

        return fig

    except Exception as e:
        messagebox.showerror("Plotting Error", f"An error occurred:\n{str(e)}")
        return None