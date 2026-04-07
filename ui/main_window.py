import os
import customtkinter as ctk
from tkinter import filedialog, messagebox
import matplotlib.pyplot as plt
from core.plotter import (generate_smps_figure, parse_ptr_file, generate_ptr_figure_from_data,
                          parse_ftir_file, generate_ftir_figure_from_data, calculate_smps_mass)

class PrecursorSelectionDialog(ctk.CTkToplevel):
    def __init__(self, parent, substance_cols, max_vals, title_text="Select Precursors"):
        super().__init__(parent)
        self.title(title_text)
        self.geometry("520x620")
        self.configure(fg_color="#111827")
        
        self.transient(parent)
        self.grab_set()

        self.selected_cols = []
        self.cancelled = True
        
        lbl = ctk.CTkLabel(
            self,
            text="Choose precursor species for the Right Y-Axis\n(sorted by max concentration)",
            font=ctk.CTkFont(size=14, weight="bold"),
            text_color="#E5E7EB"
        )
        lbl.pack(pady=(18, 12), padx=16)

        self.scroll_frame = ctk.CTkScrollableFrame(
            self,
            width=460,
            height=450,
            fg_color="#1F2937",
            corner_radius=12,
            border_width=1,
            border_color="#334155"
        )
        self.scroll_frame.pack(pady=10, padx=20, fill="both", expand=True)

        self.checkboxes = {}
        sorted_cols = sorted(substance_cols, key=lambda x: max_vals[x], reverse=True)

        for col in sorted_cols:
            var = ctk.BooleanVar(value=False)
            text = f"{col}  [ Max: {max_vals[col]:.2f} ]"
            cb = ctk.CTkCheckBox(
                self.scroll_frame,
                text=text,
                variable=var,
                text_color="#D1D5DB",
                fg_color="#0EA5E9",
                hover_color="#0284C7",
                border_color="#64748B"
            )
            cb.pack(pady=6, padx=10, anchor="w")
            self.checkboxes[col] = var

        btn_confirm = ctk.CTkButton(
            self,
            text="Confirm & Plot",
            height=38,
            corner_radius=10,
            fg_color="#0284C7",
            hover_color="#0369A1",
            font=ctk.CTkFont(size=13, weight="bold"),
            command=self.confirm
        )
        btn_confirm.pack(pady=15)

    def confirm(self):
        self.selected_cols = [col for col, var in self.checkboxes.items() if var.get()]
        self.cancelled = False
        self.destroy()

class DataVisualizerApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Advanced Data Visualization Terminal")
        self.geometry("760x940")
        self.minsize(700, 880)
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("dark-blue")

        self.ui = {
            "app_bg": "#0B1220",
            "panel_bg": "#111827",
            "card_bg": "#1F2937",
            "card_border": "#334155",
            "text_main": "#E5E7EB",
            "text_muted": "#94A3B8",
            "accent": "#0EA5E9",
            "accent_hover": "#0284C7",
            "success": "#16A34A",
            "success_hover": "#15803D",
            "warning": "#F97316"
        }

        self.configure(fg_color=self.ui["app_bg"])

        self.current_fig = None

        self.lbl_title = ctk.CTkLabel(
            self,
            text="Scientific Plotting Studio",
            font=ctk.CTkFont(size=28, weight="bold"),
            text_color=self.ui["text_main"]
        )
        self.lbl_title.pack(pady=(18, 4))

        self.lbl_subtitle = ctk.CTkLabel(
            self,
            text="SMPS / PTR / FTIR Data Workflow",
            font=ctk.CTkFont(size=13, weight="normal"),
            text_color=self.ui["text_muted"]
        )
        self.lbl_subtitle.pack(pady=(0, 10))

        self.btn_help = ctk.CTkButton(
            self, text="❓ Help", width=65, height=30, 
            corner_radius=15, fg_color="#1E293B", hover_color="#334155",
            font=ctk.CTkFont(weight="bold"), command=self.show_help
        )
        self.btn_help.place(relx=0.96, y=25, anchor="ne")

        self.tabview = ctk.CTkTabview(
            self,
            width=700,
            height=820,
            fg_color=self.ui["panel_bg"],
            segmented_button_selected_color=self.ui["accent"],
            segmented_button_selected_hover_color=self.ui["accent_hover"],
            segmented_button_unselected_color="#1E293B",
            segmented_button_unselected_hover_color="#334155",
            text_color=self.ui["text_main"]
        )
        self.tabview.pack(padx=20, pady=10, fill="both", expand=True)

        self.tab_smps = self.tabview.add("SMPS")
        self.tab_ptr = self.tabview.add("PTR")
        self.tab_ftir = self.tabview.add("FTIR")

        self.setup_smps_tab()
        self.setup_ptr_tab()
        self.setup_ftir_tab()

    def _style_entry(self, entry_widget):
        entry_widget.configure(
            height=34,
            corner_radius=8,
            fg_color="#0F172A",
            border_color=self.ui["card_border"],
            text_color=self.ui["text_main"]
        )

    def _style_primary_button(self, button_widget):
        button_widget.configure(
            corner_radius=10,
            fg_color=self.ui["accent"],
            hover_color=self.ui["accent_hover"],
            text_color="#FFFFFF"
        )

    def _style_success_button(self, button_widget):
        button_widget.configure(
            corner_radius=10,
            fg_color=self.ui["success"],
            hover_color=self.ui["success_hover"],
            text_color="#FFFFFF"
        )

    def _style_card(self, frame_widget):
        frame_widget.configure(
            fg_color=self.ui["card_bg"],
            corner_radius=10,
            border_width=1,
            border_color=self.ui["card_border"]
        )

    def show_help(self):
        help_window = ctk.CTkToplevel(self)
        help_window.title("Documentation / Help")
        help_window.geometry("600x550")
        help_window.transient(self)
        help_window.configure(fg_color=self.ui["panel_bg"])
        
        lbl = ctk.CTkLabel(help_window, text="User Manual (README.md)", font=ctk.CTkFont(size=18, weight="bold"))
        lbl.pack(pady=(15, 10))

        textbox = ctk.CTkTextbox(
            help_window,
            wrap="word",
            font=ctk.CTkFont(size=13),
            fg_color="#0F172A",
            border_color=self.ui["card_border"],
            border_width=1,
            text_color=self.ui["text_main"]
        )
        textbox.pack(fill="both", expand=True, padx=20, pady=(0, 20))
        
        base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        readme_path = os.path.join(base_dir, "README.md")
        
        if os.path.exists(readme_path):
            try:
                with open(readme_path, "r", encoding="utf-8") as f:
                    content = f.read()
                textbox.insert("0.0", content)
            except Exception as e:
                textbox.insert("0.0", f"Error reading README.md:\n{str(e)}")
        else:
            textbox.insert("0.0", "README.md not found in the project root directory.\n\nPlease create a 'README.md' file next to your 'main.py' to add instructions here.")
        
        textbox.configure(state="disabled")

    # ==========================================
    # SMPS TAB SETUP
    # ==========================================
    def setup_smps_tab(self):
        self.smps_selected_file = None

        frame_density = ctk.CTkFrame(self.tab_smps)
        self._style_card(frame_density)
        frame_density.pack(pady=(10, 0), padx=20, fill="x")
        lbl_density = ctk.CTkLabel(frame_density, text="Density (g/cm³):", font=ctk.CTkFont(weight="bold"))
        lbl_density.pack(side="left", padx=(0, 10))
        self.entry_density = ctk.CTkEntry(frame_density, width=100)
        self.entry_density.insert(0, "1.30")
        self.entry_density.pack(side="left")
        self._style_entry(self.entry_density)

        frame_file = ctk.CTkFrame(self.tab_smps)
        self._style_card(frame_file)
        frame_file.pack(pady=10, padx=20, fill="x")
        self.btn_browse_smps = ctk.CTkButton(frame_file, text="Select Excel File", command=self.browse_smps_file)
        self.btn_browse_smps.pack(side="left", padx=(0, 15))
        self._style_primary_button(self.btn_browse_smps)
        self.lbl_filename_smps = ctk.CTkLabel(frame_file, text="No file selected...", text_color=self.ui["text_muted"])
        self.lbl_filename_smps.pack(side="left", fill="x", expand=True)

        frame_x_label = ctk.CTkFrame(self.tab_smps)
        self._style_card(frame_x_label)
        frame_x_label.pack(pady=(5, 0), padx=20, fill="x")
        self.check_x_custom_var = ctk.BooleanVar(value=False)
        self.check_x_custom = ctk.CTkCheckBox(frame_x_label, text="Custom X-Axis Label", variable=self.check_x_custom_var)
        self.check_x_custom.pack(side="top", anchor="w", pady=(0, 5))
        self.entry_x_smps = ctk.CTkEntry(frame_x_label)
        self.entry_x_smps.insert(0, "Local Time")
        self.entry_x_smps.pack(side="top", fill="x")
        self._style_entry(self.entry_x_smps)

        frame_y_label = ctk.CTkFrame(self.tab_smps)
        self._style_card(frame_y_label)
        frame_y_label.pack(pady=(10, 0), padx=20, fill="x")
        self.check_y_custom_var = ctk.BooleanVar(value=False)
        self.check_y_custom = ctk.CTkCheckBox(frame_y_label, text="Custom Y-Axis Label", variable=self.check_y_custom_var)
        self.check_y_custom.pack(side="top", anchor="w", pady=(0, 5))
        self.entry_y_smps = ctk.CTkEntry(frame_y_label)
        self.entry_y_smps.insert(0, "SOA mass conc. ($\mu$g/m$^3$) calculated\nassuming d=1.30 g/cm$^3$")
        self.entry_y_smps.pack(side="top", fill="x")
        self._style_entry(self.entry_y_smps)

        frame_text = ctk.CTkFrame(self.tab_smps)
        self._style_card(frame_text)
        frame_text.pack(pady=(10, 5), padx=20, fill="x")
        self.check_text_smps_var = ctk.BooleanVar(value=False)
        self.check_text_smps = ctk.CTkCheckBox(frame_text, text="Add text on top-left of the plot", variable=self.check_text_smps_var)
        self.check_text_smps.pack(side="top", anchor="w", pady=(0, 5))
        self.entry_text_smps = ctk.CTkEntry(frame_text)
        self.entry_text_smps.insert(0, "Mass sampled on the filter was calculated to x $\\mu$g")
        self.entry_text_smps.pack(side="top", fill="x")
        self._style_entry(self.entry_text_smps)

        frame_font_smps = ctk.CTkFrame(self.tab_smps)
        self._style_card(frame_font_smps)
        frame_font_smps.pack(pady=(5, 5), padx=20, fill="x")
        lbl_font_smps = ctk.CTkLabel(frame_font_smps, text="Axis Font Size:")
        lbl_font_smps.pack(side="left", padx=(0, 10))
        self.entry_font_size_smps = ctk.CTkEntry(frame_font_smps, width=60)
        self.entry_font_size_smps.insert(0, "12")
        self.entry_font_size_smps.pack(side="left")
        self._style_entry(self.entry_font_size_smps)

        frame_buttons = ctk.CTkFrame(self.tab_smps)
        self._style_card(frame_buttons)
        frame_buttons.pack(pady=(5, 5), padx=20, fill="x")
        self.btn_preview_smps = ctk.CTkButton(frame_buttons, text="Preview Plot", height=40, font=ctk.CTkFont(size=14, weight="bold"), command=self.preview_smps_plot)
        self.btn_preview_smps.pack(side="left", fill="x", expand=True, padx=(0, 10))
        self.btn_save_smps = ctk.CTkButton(frame_buttons, text="Save Plot (.png)", height=40, font=ctk.CTkFont(size=14, weight="bold"), fg_color="#2E7D32", hover_color="#1B5E20", state="disabled", command=self.save_smps_plot)
        self.btn_save_smps.pack(side="right", fill="x", expand=True, padx=(10, 0))
        self._style_primary_button(self.btn_preview_smps)
        self._style_success_button(self.btn_save_smps)

        frame_calc = ctk.CTkFrame(self.tab_smps)
        self._style_card(frame_calc)
        frame_calc.pack(pady=(5, 10), padx=20, fill="x")
        
        lbl_calc_title = ctk.CTkLabel(frame_calc, text="Total Mass Calculator (Integration)", font=ctk.CTkFont(weight="bold"))
        lbl_calc_title.grid(row=0, column=0, columnspan=4, pady=(5, 5), padx=10, sticky="w")
        
        lbl_flow = ctk.CTkLabel(frame_calc, text="Flow Rate (L/min):")
        lbl_flow.grid(row=1, column=0, pady=5, padx=10, sticky="w")
        self.entry_flow = ctk.CTkEntry(frame_calc, width=80)
        self.entry_flow.insert(0, "16.7")
        self.entry_flow.grid(row=1, column=1, pady=5, padx=10, sticky="w")
        self._style_entry(self.entry_flow)
        
        lbl_time = ctk.CTkLabel(frame_calc, text="Time Range (HH:MM):")
        lbl_time.grid(row=2, column=0, pady=5, padx=10, sticky="w")
        
        frame_time_inputs = ctk.CTkFrame(frame_calc, fg_color="transparent")
        frame_time_inputs.grid(row=2, column=1, columnspan=2, pady=5, padx=10, sticky="w")
        
        self.entry_start_time = ctk.CTkEntry(frame_time_inputs, width=65, placeholder_text="e.g. 10:00")
        self.entry_start_time.pack(side="left")
        lbl_dash = ctk.CTkLabel(frame_time_inputs, text=" - ")
        lbl_dash.pack(side="left", padx=5)
        self.entry_end_time = ctk.CTkEntry(frame_time_inputs, width=65, placeholder_text="e.g. 11:30")
        self.entry_end_time.pack(side="left")
        self._style_entry(self.entry_start_time)
        self._style_entry(self.entry_end_time)
        
        self.btn_calc = ctk.CTkButton(frame_calc, text="Calculate", command=self.calculate_smps_mass_ui, width=100)
        self.btn_calc.grid(row=1, column=3, rowspan=2, pady=5, padx=10, sticky="e")
        self._style_primary_button(self.btn_calc)
        
        self.entry_calc_result = ctk.CTkEntry(
            frame_calc, 
            font=ctk.CTkFont(size=14, weight="bold"), 
            text_color="#4CAF50", 
            fg_color="transparent", 
            border_width=0         
        )
        self.entry_calc_result.grid(row=3, column=0, columnspan=4, pady=(5, 10), padx=10, sticky="ew")
        self.entry_calc_result.insert(0, "Result: -- μg")
        self.entry_calc_result.configure(state="readonly") 

    # ==========================================
    # PTR TAB SETUP
    # ==========================================
    def setup_ptr_tab(self):
        self.ptr_selected_file = None

        frame_file = ctk.CTkFrame(self.tab_ptr)
        self._style_card(frame_file)
        frame_file.pack(pady=10, padx=20, fill="x")
        self.btn_browse_ptr = ctk.CTkButton(frame_file, text="Select Excel File", command=self.browse_ptr_file)
        self.btn_browse_ptr.pack(side="left", padx=(0, 15))
        self._style_primary_button(self.btn_browse_ptr)
        self.lbl_filename_ptr = ctk.CTkLabel(frame_file, text="No file selected...", text_color=self.ui["text_muted"])
        self.lbl_filename_ptr.pack(side="left", fill="x", expand=True)

        frame_x_ptr = ctk.CTkFrame(self.tab_ptr)
        self._style_card(frame_x_ptr)
        frame_x_ptr.pack(pady=(5, 0), padx=20, fill="x")
        self.check_x_custom_ptr_var = ctk.BooleanVar(value=False)
        self.check_x_custom_ptr = ctk.CTkCheckBox(frame_x_ptr, text="Custom X-Axis Label", variable=self.check_x_custom_ptr_var)
        self.check_x_custom_ptr.pack(side="top", anchor="w", pady=(0, 5))
        self.entry_x_ptr = ctk.CTkEntry(frame_x_ptr)
        self.entry_x_ptr.insert(0, "Local Time")
        self.entry_x_ptr.pack(side="top", fill="x")
        self._style_entry(self.entry_x_ptr)

        frame_y_left_ptr = ctk.CTkFrame(self.tab_ptr)
        self._style_card(frame_y_left_ptr)
        frame_y_left_ptr.pack(pady=(10, 0), padx=20, fill="x")
        self.check_y_left_custom_ptr_var = ctk.BooleanVar(value=False)
        self.check_y_left_custom_ptr = ctk.CTkCheckBox(frame_y_left_ptr, text="Custom Left Y-Axis Label (Products)", variable=self.check_y_left_custom_ptr_var)
        self.check_y_left_custom_ptr.pack(side="top", anchor="w", pady=(0, 5))
        self.entry_y_left_ptr = ctk.CTkEntry(frame_y_left_ptr)
        self.entry_y_left_ptr.insert(0, "gases other than precursor (ppbv)")
        self.entry_y_left_ptr.pack(side="top", fill="x")
        self._style_entry(self.entry_y_left_ptr)

        frame_y_right_ptr = ctk.CTkFrame(self.tab_ptr)
        self._style_card(frame_y_right_ptr)
        frame_y_right_ptr.pack(pady=(10, 0), padx=20, fill="x")
        self.check_y_right_custom_ptr_var = ctk.BooleanVar(value=False)
        self.check_y_right_custom_ptr = ctk.CTkCheckBox(frame_y_right_ptr, text="Custom Right Y-Axis Label (Precursors)", variable=self.check_y_right_custom_ptr_var)
        self.check_y_right_custom_ptr.pack(side="top", anchor="w", pady=(0, 5))
        self.entry_y_right_ptr = ctk.CTkEntry(frame_y_right_ptr)
        self.entry_y_right_ptr.insert(0, "precursor (ppbv)")
        self.entry_y_right_ptr.pack(side="top", fill="x")
        self._style_entry(self.entry_y_right_ptr)

        frame_text = ctk.CTkFrame(self.tab_ptr)
        self._style_card(frame_text)
        frame_text.pack(pady=(10, 5), padx=20, fill="x")
        self.check_text_ptr_var = ctk.BooleanVar(value=False)
        self.check_text_ptr = ctk.CTkCheckBox(frame_text, text="Add text on top-left of the plot", variable=self.check_text_ptr_var)
        self.check_text_ptr.pack(side="top", anchor="w", pady=(0, 5))
        self.entry_text_ptr = ctk.CTkEntry(frame_text, placeholder_text="Enter text to display in red...")
        self.entry_text_ptr.pack(side="top", fill="x")
        self._style_entry(self.entry_text_ptr)

        frame_font_ptr = ctk.CTkFrame(self.tab_ptr)
        self._style_card(frame_font_ptr)
        frame_font_ptr.pack(pady=(5, 5), padx=20, fill="x")
        lbl_font_ptr = ctk.CTkLabel(frame_font_ptr, text="Axis Font Size:")
        lbl_font_ptr.pack(side="left", padx=(0, 10))
        self.entry_font_size_ptr = ctk.CTkEntry(frame_font_ptr, width=60)
        self.entry_font_size_ptr.insert(0, "12")
        self.entry_font_size_ptr.pack(side="left")
        self._style_entry(self.entry_font_size_ptr)

        frame_buttons = ctk.CTkFrame(self.tab_ptr)
        self._style_card(frame_buttons)
        frame_buttons.pack(pady=(15, 10), padx=20, fill="x")
        self.btn_preview_ptr = ctk.CTkButton(frame_buttons, text="Preview Plot", height=40, font=ctk.CTkFont(size=14, weight="bold"), command=self.preview_ptr_plot)
        self.btn_preview_ptr.pack(side="left", fill="x", expand=True, padx=(0, 10))
        self.btn_save_ptr = ctk.CTkButton(frame_buttons, text="Save Plot (.png)", height=40, font=ctk.CTkFont(size=14, weight="bold"), fg_color="#2E7D32", hover_color="#1B5E20", state="disabled", command=self.save_ptr_plot)
        self.btn_save_ptr.pack(side="right", fill="x", expand=True, padx=(10, 0))
        self._style_primary_button(self.btn_preview_ptr)
        self._style_success_button(self.btn_save_ptr)

    # ==========================================
    # FTIR TAB SETUP
    # ==========================================
    def setup_ftir_tab(self):
        self.ftir_selected_file = None

        frame_file = ctk.CTkFrame(self.tab_ftir)
        self._style_card(frame_file)
        frame_file.pack(pady=10, padx=20, fill="x")
        self.btn_browse_ftir = ctk.CTkButton(frame_file, text="Select Excel File", command=self.browse_ftir_file)
        self.btn_browse_ftir.pack(side="left", padx=(0, 15))
        self._style_primary_button(self.btn_browse_ftir)
        self.lbl_filename_ftir = ctk.CTkLabel(frame_file, text="No file selected...", text_color=self.ui["text_muted"])
        self.lbl_filename_ftir.pack(side="left", fill="x", expand=True)

        frame_x_ftir = ctk.CTkFrame(self.tab_ftir)
        self._style_card(frame_x_ftir)
        frame_x_ftir.pack(pady=(5, 0), padx=20, fill="x")
        self.check_x_custom_ftir_var = ctk.BooleanVar(value=False)
        self.check_x_custom_ftir = ctk.CTkCheckBox(frame_x_ftir, text="Custom X-Axis Label", variable=self.check_x_custom_ftir_var)
        self.check_x_custom_ftir.pack(side="top", anchor="w", pady=(0, 5))
        self.entry_x_ftir = ctk.CTkEntry(frame_x_ftir)
        self.entry_x_ftir.insert(0, "Local Time")
        self.entry_x_ftir.pack(side="top", fill="x")
        self._style_entry(self.entry_x_ftir)

        frame_y_left_ftir = ctk.CTkFrame(self.tab_ftir)
        self._style_card(frame_y_left_ftir)
        frame_y_left_ftir.pack(pady=(10, 0), padx=20, fill="x")
        self.check_y_left_custom_ftir_var = ctk.BooleanVar(value=False)
        self.check_y_left_custom_ftir = ctk.CTkCheckBox(frame_y_left_ftir, text="Custom Left Y-Axis Label (Products)", variable=self.check_y_left_custom_ftir_var)
        self.check_y_left_custom_ftir.pack(side="top", anchor="w", pady=(0, 5))
        self.entry_y_left_ftir = ctk.CTkEntry(frame_y_left_ftir)
        self.entry_y_left_ftir.insert(0, "gases other than precursors (ppbv)")
        self.entry_y_left_ftir.pack(side="top", fill="x")
        self._style_entry(self.entry_y_left_ftir)

        frame_y_right_ftir = ctk.CTkFrame(self.tab_ftir)
        self._style_card(frame_y_right_ftir)
        frame_y_right_ftir.pack(pady=(10, 0), padx=20, fill="x")
        self.check_y_right_custom_ftir_var = ctk.BooleanVar(value=False)
        self.check_y_right_custom_ftir = ctk.CTkCheckBox(frame_y_right_ftir, text="Custom Right Y-Axis Label (Precursors)", variable=self.check_y_right_custom_ftir_var)
        self.check_y_right_custom_ftir.pack(side="top", anchor="w", pady=(0, 5))
        self.entry_y_right_ftir = ctk.CTkEntry(frame_y_right_ftir)
        self.entry_y_right_ftir.insert(0, "precursors (ppbv)")
        self.entry_y_right_ftir.pack(side="top", fill="x")
        self._style_entry(self.entry_y_right_ftir)

        frame_text = ctk.CTkFrame(self.tab_ftir)
        self._style_card(frame_text)
        frame_text.pack(pady=(10, 5), padx=20, fill="x")
        self.check_text_ftir_var = ctk.BooleanVar(value=False)
        self.check_text_ftir = ctk.CTkCheckBox(frame_text, text="Add text on top-left of the plot", variable=self.check_text_ftir_var)
        self.check_text_ftir.pack(side="top", anchor="w", pady=(0, 5))
        self.entry_text_ftir = ctk.CTkEntry(frame_text, placeholder_text="Enter text to display in red...")
        self.entry_text_ftir.pack(side="top", fill="x")
        self._style_entry(self.entry_text_ftir)

        frame_font_ftir = ctk.CTkFrame(self.tab_ftir)
        self._style_card(frame_font_ftir)
        frame_font_ftir.pack(pady=(5, 5), padx=20, fill="x")
        lbl_font_ftir = ctk.CTkLabel(frame_font_ftir, text="Axis Font Size:")
        lbl_font_ftir.pack(side="left", padx=(0, 10))
        self.entry_font_size_ftir = ctk.CTkEntry(frame_font_ftir, width=60)
        self.entry_font_size_ftir.insert(0, "12")
        self.entry_font_size_ftir.pack(side="left")
        self._style_entry(self.entry_font_size_ftir)

        frame_buttons = ctk.CTkFrame(self.tab_ftir)
        self._style_card(frame_buttons)
        frame_buttons.pack(pady=(15, 10), padx=20, fill="x")
        self.btn_preview_ftir = ctk.CTkButton(frame_buttons, text="Preview Plot", height=40, font=ctk.CTkFont(size=14, weight="bold"), command=self.preview_ftir_plot)
        self.btn_preview_ftir.pack(side="left", fill="x", expand=True, padx=(0, 10))
        self.btn_save_ftir = ctk.CTkButton(frame_buttons, text="Save Plot (.png)", height=40, font=ctk.CTkFont(size=14, weight="bold"), fg_color="#2E7D32", hover_color="#1B5E20", state="disabled", command=self.save_ftir_plot)
        self.btn_save_ftir.pack(side="right", fill="x", expand=True, padx=(10, 0))
        self._style_primary_button(self.btn_preview_ftir)
        self._style_success_button(self.btn_save_ftir)

    # ==========================================
    # LOGIC: SMPS
    # ==========================================
    def browse_smps_file(self):
        filepath = filedialog.askopenfilename(title="Select SMPS Data File", filetypes=(("Excel Files", "*.xlsx"), ("All Files", "*.*")))
        if filepath:
            self.smps_selected_file = filepath
            self.lbl_filename_smps.configure(text=filepath.split("/")[-1], text_color="white")
            self.current_fig = None
            self.btn_save_smps.configure(state="disabled")

    def preview_smps_plot(self):
        if not getattr(self, 'smps_selected_file', None):
            self.lbl_filename_smps.configure(text="Please select an .xlsx file first!", text_color="#EF5350")
            return
            
        try:
            density = float(self.entry_density.get().strip())
        except ValueError:
            messagebox.showerror("Input Error", "Density must be a valid number.")
            return

        try:
            font_size = int(self.entry_font_size_smps.get().strip())
        except ValueError:
            font_size = 12

        if self.check_x_custom_var.get():
            x_label = self.entry_x_smps.get()
        else:
            x_label = "Local Time"

        if self.check_y_custom_var.get():
            y_label = self.entry_y_smps.get()
        else:
            y_label = f"SOA mass conc. ($\\mu$g/m$^3$) calculated\nassuming d={density} g/cm$^3$"

        if self.current_fig: plt.close(self.current_fig)
        text_info = (self.check_text_smps_var.get(), self.entry_text_smps.get())
        
        self.current_fig = generate_smps_figure(self.smps_selected_file, x_label, y_label, text_info, density, font_size)
        if self.current_fig:
            self.btn_save_smps.configure(state="normal")
            plt.show()

    def save_smps_plot(self):
        if self.current_fig and self.smps_selected_file:
            save_path = f"{os.path.splitext(self.smps_selected_file)[0]}.png"
            try:
                self.current_fig.savefig(save_path, dpi=300, bbox_inches='tight')
                messagebox.showinfo("Success", f"Plot saved to:\n\n{save_path}")
            except Exception as e:
                messagebox.showerror("Save Error", str(e))

    def _set_calc_result(self, text, color):
        self.entry_calc_result.configure(state="normal", text_color=color)
        self.entry_calc_result.delete(0, "end")
        self.entry_calc_result.insert(0, text)
        self.entry_calc_result.configure(state="readonly")

    def calculate_smps_mass_ui(self):
        if not getattr(self, 'smps_selected_file', None):
            self._set_calc_result("Error: Select file first!", "#EF5350")
            return
            
        try:
            density = float(self.entry_density.get().strip())
        except ValueError:
            self._set_calc_result("Error: Invalid density!", "#EF5350")
            return
            
        flow_rate = self.entry_flow.get()
        start_time = self.entry_start_time.get()
        end_time = self.entry_end_time.get()
        
        if not flow_rate or not start_time or not end_time:
            self._set_calc_result("Error: Fill all fields!", "#EF5350")
            return

        res = calculate_smps_mass(self.smps_selected_file, flow_rate, start_time, end_time, density)
        
        if res is not None:
            self._set_calc_result(f"Result: {res:.2f} μg", "#4CAF50")

    # ==========================================
    # LOGIC: PTR
    # ==========================================
    def browse_ptr_file(self):
        filepath = filedialog.askopenfilename(title="Select PTR Data File", filetypes=(("Excel Files", "*.xlsx"), ("All Files", "*.*")))
        if filepath:
            self.ptr_selected_file = filepath
            self.lbl_filename_ptr.configure(text=filepath.split("/")[-1], text_color="white")
            self.current_fig = None
            self.btn_save_ptr.configure(state="disabled")

    def preview_ptr_plot(self):
        if not getattr(self, 'ptr_selected_file', None):
            self.lbl_filename_ptr.configure(text="Please select an .xlsx file first!", text_color="#EF5350")
            return
            
        try:
            font_size = int(self.entry_font_size_ptr.get().strip())
        except ValueError:
            font_size = 12

        result = parse_ptr_file(self.ptr_selected_file)
        if not result: return
        x_data, data_df, substance_cols, max_vals = result

        dialog = PrecursorSelectionDialog(self, substance_cols, max_vals)
        self.wait_window(dialog)
        if dialog.cancelled: return

        right_cols = dialog.selected_cols
        left_cols = [c for c in substance_cols if c not in right_cols]

        if right_cols:
            default_right = f"{', '.join(right_cols)} (ppbv)"
            default_left = f"gases other than {', '.join(right_cols)} (ppbv)"
        else:
            default_right = "precursor (ppbv)"
            default_left = "gases other than precursor (ppbv)"

        x_label = self.entry_x_ptr.get() if self.check_x_custom_ptr_var.get() else "Local Time"
        y_label_left = self.entry_y_left_ptr.get() if self.check_y_left_custom_ptr_var.get() else default_left
        y_label_right = self.entry_y_right_ptr.get() if self.check_y_right_custom_ptr_var.get() else default_right

        if self.current_fig: plt.close(self.current_fig)
        text_info = (self.check_text_ptr_var.get(), self.entry_text_ptr.get())
        
        self.current_fig = generate_ptr_figure_from_data(
            x_data, data_df, left_cols, right_cols, 
            x_label, y_label_left, y_label_right, 
            text_info, font_size
        )
        if self.current_fig:
            self.btn_save_ptr.configure(state="normal")
            plt.show()

    def save_ptr_plot(self):
        if self.current_fig and self.ptr_selected_file:
            save_path = f"{os.path.splitext(self.ptr_selected_file)[0]}.png"
            try:
                self.current_fig.savefig(save_path, dpi=300, bbox_inches='tight')
                messagebox.showinfo("Success", f"Plot saved to:\n\n{save_path}")
            except Exception as e:
                messagebox.showerror("Save Error", str(e))

    # ==========================================
    # LOGIC: FTIR
    # ==========================================
    def browse_ftir_file(self):
        filepath = filedialog.askopenfilename(title="Select FTIR Data File", filetypes=(("Excel Files", "*.xlsx"), ("All Files", "*.*")))
        if filepath:
            self.ftir_selected_file = filepath
            self.lbl_filename_ftir.configure(text=filepath.split("/")[-1], text_color="white")
            self.current_fig = None
            self.btn_save_ftir.configure(state="disabled")

    def preview_ftir_plot(self):
        if not getattr(self, 'ftir_selected_file', None):
            self.lbl_filename_ftir.configure(text="Please select an .xlsx file first!", text_color="#EF5350")
            return
            
        try:
            font_size = int(self.entry_font_size_ftir.get().strip())
        except ValueError:
            font_size = 12
        
        result = parse_ftir_file(self.ftir_selected_file)
        if not result: return
        x_data, data_df, substance_cols, max_vals = result

        dialog = PrecursorSelectionDialog(self, substance_cols, max_vals, title_text="Select Precursors (FTIR)")
        self.wait_window(dialog)
        if dialog.cancelled: return

        right_cols = dialog.selected_cols
        left_cols = [c for c in substance_cols if c not in right_cols]

        if right_cols:
            default_right = f"{', '.join(right_cols)} (ppbv)"
            default_left = f"gases other than {', '.join(right_cols)} (ppbv)"
        else:
            default_right = "precursor (ppbv)"
            default_left = "gases other than precursors (ppbv)"

        x_label = self.entry_x_ftir.get() if self.check_x_custom_ftir_var.get() else "Local Time"
        y_label_left = self.entry_y_left_ftir.get() if self.check_y_left_custom_ftir_var.get() else default_left
        y_label_right = self.entry_y_right_ftir.get() if self.check_y_right_custom_ftir_var.get() else default_right

        if self.current_fig: plt.close(self.current_fig)
        
        text_info = (self.check_text_ftir_var.get(), self.entry_text_ftir.get())
        
        self.current_fig = generate_ftir_figure_from_data(
            x_data, data_df, left_cols, right_cols,
            x_label, y_label_left, y_label_right,
            text_info, font_size
        )
        
        if self.current_fig:
            self.btn_save_ftir.configure(state="normal")
            plt.show()

    def save_ftir_plot(self):
        if self.current_fig and self.ftir_selected_file:
            save_path = f"{os.path.splitext(self.ftir_selected_file)[0]}.png"
            try:
                self.current_fig.savefig(save_path, dpi=300, bbox_inches='tight')
                messagebox.showinfo("Success", f"Plot saved to:\n\n{save_path}")
            except Exception as e:
                messagebox.showerror("Save Error", str(e))