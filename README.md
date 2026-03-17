**⚠️ Attention：此项目暂未经过检查，上传内容有待完善**


# 科学数据可视化终端 (Advanced Data Visualization Terminal)

一个专为大气科学数据分析设计的桌面应用程序，支持 SMPS、PTR 和 FTIR 数据的清洗、可视化和计算。使用现代化的 GUI 界面，帮助科研人员快速生成高质量的科学图表。

## ✨ 功能特性

### 📊 数据可视化
- **SMPS (扫描电迁移率颗粒物粒径谱仪)**: 绘制颗粒物质量浓度时间序列图，支持密度设置和质量积分计算。
- **PTR (质子转移反应质谱仪)**: 双 Y 轴绘图，自动剔除无效数据，支持前体物选择。
- **FTIR (傅里叶变换红外光谱仪)**: 智能解析复杂数据结构，自动过滤干扰气体，支持乘数应用。

### 🛠️ 高级功能
- 自定义坐标轴标签和字体大小
- 添加图表文本注释
- 高分辨率 PNG 保存 (300dpi)，适合期刊发表
- 实时预览和保存功能
- 质量积分计算器 (SMPS)

### 🎨 用户界面
- 现代暗色主题界面
- 标签页式布局，便于切换数据类型
- 响应式设计，支持窗口调整

## 🚀 安装与运行

### 系统要求
- Python 3.8 或更高版本
- macOS / Windows / Linux

### 安装步骤
1. **克隆项目**:
   ```bash
   git clone https://github.com/z-whan/Atmospheric-Chemistry-data-processor.git
   cd Atmospheric-Chemistry-data-processor
   ```

2. **创建虚拟环境** (推荐):
   ```bash
   python -m venv venv
   source venv/bin/activate  # macOS/Linux
   # 或 venv\Scripts\activate  # Windows
   ```

3. **安装依赖**:
   ```bash
   pip install -r requirements.txt
   ```

   主要依赖包:
   - `customtkinter` - 现代化 GUI 框架
   - `pandas` - 数据处理
   - `matplotlib` - 图表绘制
   - `numpy` - 数值计算
   - `openpyxl` - Excel 文件读取

4. **运行应用程序**:
   ```bash
   python main.py
   ```

## 📖 使用指南

### 基本操作
1. 启动应用后，选择相应的标签页 (SMPS / PTR / FTIR)
2. 点击 "Select Excel File" 选择数据文件 (.xlsx 格式)
3. 配置参数 (密度、坐标轴标签等)
4. 点击 "Preview Plot" 预览图表
5. 如满意，点击 "Save Plot (.png)" 保存

### 数据格式要求
- **SMPS**: Excel 文件包含 "Start Time" 和 "Total Conc." 列
- **PTR**: 包含 "AbsTime" 列和以 'm' 开头的物质列
- **FTIR**: 包含 "Local Time" 行，第一行标题，第二行乘数，第三行起数据

### 高级设置
- **自定义标签**: 勾选 "Custom X/Y-Axis Label" 并输入文本
- **文本注释**: 勾选 "Add text on top-left" 添加图表注释
- **换行**: 在标签中使用 `\n` 实现换行
- **质量计算**: 在 SMPS 标签页输入流速和时间范围计算总质量

## 📁 项目结构

```
DataProcessor/
├── main.py                 # 应用程序入口
├── main.spec               # PyInstaller 打包配置
├── README.md               # 项目文档
├── build/                  # 构建输出目录
├── venv_mac/               # macOS 虚拟环境
├── core/
│   └── plotter.py          # 数据处理和绘图核心逻辑
└── ui/
    └── main_window.py      # GUI 主窗口和界面逻辑
```

## 🐛 故障排除

### 常见问题
- **文件读取失败**: 确保 Excel 文件未被其他程序打开，格式正确
- **无数据显示**: 检查数据列名是否匹配预期格式
- **依赖安装失败**: 升级 pip (`pip install --upgrade pip`) 或使用 conda

### 错误报告
如果遇到问题，请检查终端输出或查看应用内的错误消息。

## 🤝 贡献

欢迎提交 Issue 和 Pull Request！

1. Fork 本项目
2. 创建特性分支 (`git checkout -b feature/AmazingFeature`)
3. 提交更改 (`git commit -m 'Add some AmazingFeature'`)
4. 推送到分支 (`git push origin feature/AmazingFeature`)
5. 打开 Pull Request

## 📄 许可证

本项目采用 MIT 许可证 - 查看 [LICENSE](LICENSE) 文件了解详情。

---

# Advanced Data Visualization Terminal (科学数据可视化终端)

A desktop application designed for atmospheric science data analysis, supporting visualization and calculation of SMPS, PTR, and FTIR data. Features a modern GUI interface to help researchers quickly generate high-quality scientific figures.

## ✨ Features

### 📊 Data Visualization
- **SMPS (Scanning Mobility Particle Sizer)**: Plot particle mass concentration time series with customizable density settings and mass integration calculations.
- **PTR (Proton Transfer Reaction Mass Spectrometry)**: Dual Y-axis plotting with automatic invalid data removal and precursor selection support.
- **FTIR (Fourier Transform Infrared Spectroscopy)**: Intelligent complex data structure parsing with automatic interfering gas filtering and multiplier application.

### 🛠️ Advanced Features
- Customizable axis labels and font sizes
- Add text annotations to plots
- High-resolution PNG export (300dpi) suitable for journal publication
- Real-time preview and save functionality
- Mass integration calculator (SMPS)

### 🎨 User Interface
- Modern dark theme interface
- Tab-based layout for easy data type switching
- Responsive design with window resizing support

## 🚀 Installation & Setup

### System Requirements
- Python 3.8 or higher
- macOS / Windows / Linux

### Installation Steps
1. **Clone the repository**:
   ```bash
   git clone https://github.com/z-whan/Atmospheric-Chemistry-data-processor.git
   cd Atmospheric-Chemistry-data-processor
   ```

2. **Create virtual environment** (recommended):
   ```bash
   python -m venv venv
   source venv/bin/activate  # macOS/Linux
   # or venv\Scripts\activate  # Windows
   ```

3. **Install dependencies**:
   ```bash
   pip install -r requirements.txt
   ```

   Key packages:
   - `customtkinter` - Modern GUI framework
   - `pandas` - Data processing
   - `matplotlib` - Chart plotting
   - `numpy` - Numerical computation
   - `openpyxl` - Excel file reading

4. **Run the application**:
   ```bash
   python main.py
   ```

## 📖 User Guide

### Basic Operation
1. After launching, select the appropriate tab (SMPS / PTR / FTIR)
2. Click "Select Excel File" to choose your data file (.xlsx format)
3. Configure parameters (density, axis labels, etc.)
4. Click "Preview Plot" to preview the figure
5. If satisfied, click "Save Plot (.png)" to save

### Data Format Requirements
- **SMPS**: Excel file containing "Start Time" and "Total Conc." columns
- **PTR**: Contains "AbsTime" column and substance columns starting with 'm'
- **FTIR**: Contains "Local Time" header row, multipliers in second row, data from third row onwards

### Advanced Settings
- **Custom Labels**: Check "Custom X/Y-Axis Label" and enter text
- **Text Annotation**: Check "Add text on top-left" to add plot notes
- **Line Breaks**: Use `\n` in labels to implement line breaks
- **Mass Calculation**: Enter flow rate and time range in SMPS tab to calculate total mass

## 📁 Project Structure

```
Atmospheric-Chemistry-data-processor/
├── main.py                 # Application entry point
├── README.md               # Project documentation
├── requirements.txt        # Python dependencies
├── core/
│   └── plotter.py          # Core data processing and plotting logic
└── ui/
    └── main_window.py      # GUI main window and interface logic
```

## 🐛 Troubleshooting

### Common Issues
- **File read failure**: Ensure Excel file is not open in other programs and format is correct
- **No data displayed**: Check if column names match expected format
- **Dependency installation failed**: Upgrade pip (`pip install --upgrade pip`) or use conda

### Error Reporting
If you encounter issues, check terminal output or error messages within the application.

## 🤝 Contributing

Issues and Pull Requests are welcome!

1. Fork this repository
2. Create a feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.


