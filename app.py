import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext, simpledialog
import os
import sys
import platform
import subprocess
import time
import traceback
from pathlib import Path
import importlib.util
import warnings
import re 
from difflib import get_close_matches
import threading
import odf
from functools import lru_cache
try:
    from importlib import metadata as importlib_metadata
except ImportError:
    import importlib_metadata  # type: ignore
DEPENDENCY_VERSIONS = {
    "pandas": ">=2.2.3",
    "openpyxl": ">=3.1.4",
    "xlrd": ">=2.0.1",
    "pyxlsb": ">=1.0.9",
    "fuzzywuzzy": ">=0.18.0",
    "python-Levenshtein": ">=0.12.2",
    "XlsxWriter": ">=3.2.0",
    "pyarrow": ">=16.1.0",
    "pywin32": ">=310",
    "packaging": ">=23.2",  # <-- Added packaging
}

# Dependency details for UI.
DEPENDENCY_INFO = [
    {"name": "pandas", "import": "pandas", "required": True, "desc": "Data manipulation and analysis", "version": DEPENDENCY_VERSIONS["pandas"]},
    {"name": "openpyxl", "import": "openpyxl", "required": True, "desc": "Read/write .xlsx/.xlsm files", "version": DEPENDENCY_VERSIONS["openpyxl"]},
    {"name": "xlrd", "import": "xlrd", "required": True, "desc": "Read legacy .xls files", "version": DEPENDENCY_VERSIONS["xlrd"]},
    {"name": "packaging", "import": "packaging", "required": True, "desc": "Version parsing and comparison", "version": DEPENDENCY_VERSIONS["packaging"]},  # <-- Added packaging
    {"name": "pyxlsb", "import": "pyxlsb", "required": False, "desc": "Read .xlsb files (optional)", "version": DEPENDENCY_VERSIONS["pyxlsb"]},
    {"name": "fuzzywuzzy", "import": "fuzzywuzzy", "required": False, "desc": "Fuzzy string matching (optional)", "version": DEPENDENCY_VERSIONS["fuzzywuzzy"]},
    {"name": "python-Levenshtein", "import": "Levenshtein", "required": False, "desc": "Faster fuzzywuzzy (optional)", "version": DEPENDENCY_VERSIONS["python-Levenshtein"]},
    {"name": "XlsxWriter", "import": "xlsxwriter", "required": False, "desc": "Alternative Excel writer (optional)", "version": DEPENDENCY_VERSIONS["XlsxWriter"]},
    {"name": "pyarrow", "import": "pyarrow", "required": False, "desc": "Faster data operations (optional)", "version": DEPENDENCY_VERSIONS["pyarrow"]},
    {"name": "pywin32", "import": "win32api", "required": False, "desc": "Enterprise Excel/COM support (Windows only, optional)", "version": DEPENDENCY_VERSIONS["pywin32"]},
]

# Dependency flags
HAS_PANDAS = False
HAS_WIN32COM = False
HAS_PYTHONCOM = False
HAS_ODF = False
HAS_PYXLSB = False
HAS_OPENPYXL = False
HAS_XLRD = False
HAS_PACKAGING = False  # <-- Added packaging flag

# Dependency details for UI

# Try to import dependencies and set flags
try:
    import pandas as pd
    HAS_PANDAS = True
except ImportError:
    pass
try:
    import win32com.client
    HAS_WIN32COM = True
except ImportError:
    pass
try:
    import pythoncom
    HAS_PYTHONCOM = True
except ImportError:
    pass
try:
    import odf
    HAS_ODF = True
except ImportError:
    pass
try:
    import pyxlsb
    HAS_PYXLSB = True
except ImportError:
    pass
try:
    import openpyxl
    HAS_OPENPYXL = True
except ImportError:
    pass
try:
    import xlrd
    HAS_XLRD = True
except ImportError:
    pass
try:
    import packaging
    HAS_PACKAGING = True
except ImportError:
    pass

# When running as a compiled executable, we need to force these flags to True
# since the imports work differently in frozen environments
if getattr(sys, 'frozen', False):
    HAS_PANDAS = True
    HAS_OPENPYXL = True
    HAS_XLRD = True
    HAS_PACKAGING = True

# List of required dependencies (for core functionality)
REQUIRED_DEPENDENCIES = [
    ("pandas", HAS_PANDAS),
    ("openpyxl", HAS_OPENPYXL),
    ("xlrd", HAS_XLRD),
    ("packaging", HAS_PACKAGING),  # <-- Added packaging to required dependencies
]
# Optional dependencies for advanced Excel support
OPTIONAL_DEPENDENCIES = [
    ("win32com", HAS_WIN32COM),
    ("pythoncom", HAS_PYTHONCOM),
    ("odfpy", HAS_ODF),
    ("pyxlsb", HAS_PYXLSB),
]

def get_missing_dependencies():
    missing = [name for name, flag in REQUIRED_DEPENDENCIES if not flag]
    return missing

def get_missing_optional_dependencies():
    missing = [name for name, flag in OPTIONAL_DEPENDENCIES if not flag]
    return missing

def check_package_installed(pkg_name, import_name=None, version_spec=None):
    # If running as a frozen executable, consider core packages as installed
    if getattr(sys, 'frozen', False):
        core_packages = ["pandas", "openpyxl", "xlrd", "packaging"]
        if pkg_name in core_packages:
            return True
    
    # Use importlib to check for presence
    import importlib.util
    mod_name = import_name or pkg_name.replace('-', '_')
    
    # First try direct import which is more reliable
    try:
        __import__(mod_name)
        if not version_spec:
            return True
    except ImportError:
        # If direct import fails, check with find_spec
        spec = importlib.util.find_spec(mod_name)
        if spec is None:
            return False
        # If version is not specified, just presence is enough
        if not version_spec:
            return True
    
    # Try to get the installed version
    try:
        installed_version = importlib_metadata.version(pkg_name)
        # Parse version specifier (e.g., '>=1.2.3')
        import re
        from packaging import version as packaging_version
        match = re.match(r'(>=|<=|==|>|<|~=)?\s*([\d\.]+)', version_spec.strip())
        if match:
            op, required_version = match.groups()
            op = op or '=='
            if op == '==':
                return packaging_version.parse(installed_version) == packaging_version.parse(required_version)
            elif op == '>=':
                return packaging_version.parse(installed_version) >= packaging_version.parse(required_version)
            elif op == '<=':
                return packaging_version.parse(installed_version) <= packaging_version.parse(required_version)
            elif op == '>':
                return packaging_version.parse(installed_version) > packaging_version.parse(required_version)
            elif op == '<':
                return packaging_version.parse(installed_version) < packaging_version.parse(required_version)
            elif op == '~=':
                # Compatible release, e.g., ~=1.4 means >=1.4, ==1.*k
                return packaging_version.parse(installed_version) >= packaging_version.parse(required_version)
            else:
                return True  # Unknown operator, fallback to True
        else:
            return True  # If we can't parse, fallback to presence
    except Exception:
        # If we can't get version but we know the package is present, consider it installed
        return True

class ExcelComparisonTool:
    def __init__(self, root):
        self.root = root
        self.root.title("Excella")
        self.root.geometry("1000x700")
        
        # Set to open in full screen by default, cross-platform
        system = platform.system()
        if system == 'Windows':
            self.root.state('zoomed')
        elif system == 'Linux':
            try:
                self.root.attributes('-zoomed', True)
            except tk.TclError:
                self.root.state('normal')
        elif system == 'Darwin':  # macOS
            try:
                self.root.attributes('-fullscreen', True)
            except tk.TclError:
                self.root.state('normal')
        else:
            self.root.state('normal')
        
        # Variables
        self.master_file_path = tk.StringVar()
        self.secondary_file_path = tk.StringVar()
        self.master_df = None
        self.secondary_df = None
        self.master_columns = []
        self.secondary_columns = []
        self.result_df = None
        self.dependency_tab = None
        self.dependency_text = None
        self.dependency_install_btn = None
        
        self.setup_gui()
        # After GUI setup, check for missing dependencies
        self.check_and_handle_dependencies()
        
    def setup_gui(self):
        # Main frame
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(0, weight=1)

        # Create notebook directly in main frame
        self.notebook = ttk.Notebook(main_frame)
        self.notebook.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        main_frame.rowconfigure(0, weight=1)

        # Add dependencies tab first (if needed, will be shown later)
        self.dependency_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.dependency_tab, text="Dependencies")
        self.setup_dependency_tab(self.dependency_tab)

        # Tab 1: File Selection & Column Mapping
        self.mapping_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.mapping_tab, text="File & Column Mapping")
        self.setup_mapping_tab(self.mapping_tab)

        # Tab 2: Options & Processing
        self.options_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.options_tab, text="Options & Processing")
        self.setup_options_tab(self.options_tab)

    def setup_dependency_tab(self, parent):
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(3, weight=1)
        ttk.Label(parent, text="Dependencies", font=("Arial", 12, "bold")).grid(row=0, column=0, sticky=tk.W, pady=(10, 5), padx=10)
        
        # Different message for frozen vs non-frozen environment
        if getattr(sys, 'frozen', False):
            ttk.Label(parent, text="This is a packaged application with all dependencies included.\nNo installation is required.", font=("Arial", 9)).grid(row=1, column=0, sticky=tk.W, padx=10)
        else:
            ttk.Label(parent, text="Below is a list of required and optional dependencies for this tool.\nSelect any to install, or use Install All for all missing dependencies.", font=("Arial", 9)).grid(row=1, column=0, sticky=tk.W, padx=10)

        # Dependency grid with checkboxes
        grid_frame = ttk.Frame(parent)
        grid_frame.grid(row=2, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=10, pady=(5, 5))
        grid_frame.columnconfigure(0, weight=0)
        grid_frame.columnconfigure(1, weight=0)
        grid_frame.columnconfigure(2, weight=0)
        grid_frame.columnconfigure(3, weight=1)
        grid_frame.columnconfigure(4, weight=0)

        headers = ["Select", "Package", "Required", "Purpose", "Status"]
        for col, header in enumerate(headers):
            ttk.Label(grid_frame, text=header, font=("Arial", 10, "bold")).grid(row=0, column=col, sticky=tk.W, padx=2, pady=2)

        self.dep_vars = []
        self.dep_labels = []
        for i, dep in enumerate(DEPENDENCY_INFO):
            var = tk.BooleanVar()
            self.dep_vars.append(var)
            installed = check_package_installed(dep["name"], dep.get("import"), dep.get("version"))
            status = "Installed" if installed else "Missing"
            color = "green" if installed else "red"
            # Checkbox
            cb = ttk.Checkbutton(grid_frame, variable=var)
            cb.grid(row=i+1, column=0, sticky=tk.W, padx=2)
            if installed:
                cb.state(["disabled"])
            # Package name with version
            pkg_label = dep["name"]
            if dep.get("version"):
                pkg_label += f" ({dep['version']})"
            lbl_pkg = ttk.Label(grid_frame, text=pkg_label, foreground=color)
            lbl_pkg.grid(row=i+1, column=1, sticky=tk.W, padx=2)
            # Required
            lbl_req = ttk.Label(grid_frame, text="Yes" if dep["required"] else "No", foreground=color)
            lbl_req.grid(row=i+1, column=2, sticky=tk.W, padx=2)
            # Purpose
            lbl_purp = ttk.Label(grid_frame, text=dep["desc"], foreground=color)
            lbl_purp.grid(row=i+1, column=3, sticky=tk.W, padx=2)
            # Status
            lbl_stat = ttk.Label(grid_frame, text=status, foreground=color)
            lbl_stat.grid(row=i+1, column=4, sticky=tk.W, padx=2)
            self.dep_labels.append((lbl_pkg, lbl_req, lbl_purp, lbl_stat, cb))

        # Legend
        legend_frame = ttk.Frame(parent)
        legend_frame.grid(row=3, column=0, sticky=tk.W, padx=10, pady=(0, 5))
        ttk.Label(legend_frame, text="Legend:").pack(side=tk.LEFT)
        ttk.Label(legend_frame, text="  ", foreground="green").pack(side=tk.LEFT)
        ttk.Label(legend_frame, text="Installed  ").pack(side=tk.LEFT)
        ttk.Label(legend_frame, text="  ", foreground="red").pack(side=tk.LEFT)
        ttk.Label(legend_frame, text="Missing").pack(side=tk.LEFT)

        # Buttons
        btn_frame = ttk.Frame(parent)
        btn_frame.grid(row=4, column=0, sticky=tk.W, padx=10, pady=(0, 10))
        self.dependency_install_btn = ttk.Button(btn_frame, text="Install All Missing Dependencies", command=self.install_missing_dependencies)
        self.dependency_install_btn.pack(side=tk.LEFT, padx=(0, 10))
        self.dependency_selected_btn = ttk.Button(btn_frame, text="Install Selected Dependencies", command=self.install_selected_dependencies)
        self.dependency_selected_btn.pack(side=tk.LEFT, padx=(0, 10))
        self.dependency_refresh_btn = ttk.Button(btn_frame, text="Refresh", command=self.refresh_dependency_status)
        self.dependency_refresh_btn.pack(side=tk.LEFT, padx=(0, 10))
        self.dependency_bypass_btn = ttk.Button(btn_frame, text="Skip/Bypass (Enable App)", command=self.bypass_dependencies)
        self.dependency_bypass_btn.pack(side=tk.LEFT, padx=(0, 10))
        # Add uninstall button
        self.dependency_uninstall_btn = ttk.Button(
            btn_frame,
            text="Uninstall All Installed Dependencies",
            command=self.uninstall_installed_dependencies
        )
        self.dependency_uninstall_btn.pack(side=tk.LEFT, padx=(0, 10))

        # Output area
        self.dependency_text = scrolledtext.ScrolledText(parent, height=8, width=80, state=tk.DISABLED)
        self.dependency_text.grid(row=5, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=10, pady=(0, 10))
        parent.rowconfigure(5, weight=1)

        self.refresh_dependency_status()

    def refresh_dependency_status(self):
        # If running as a frozen executable, mark all dependencies as installed
        if getattr(sys, 'frozen', False):
            for i, dep in enumerate(DEPENDENCY_INFO):
                color = "green"
                self.dep_labels[i][0].config(foreground=color)  # pkg
                self.dep_labels[i][1].config(foreground=color)  # req
                self.dep_labels[i][2].config(foreground=color)  # purpose
                self.dep_labels[i][3].config(foreground=color, text="Included")  # status
                self.dep_labels[i][4].state(["disabled"])
                self.dep_vars[i].set(False)
            return
            
        # Normal dependency check for non-frozen environment
        for i, dep in enumerate(DEPENDENCY_INFO):
            installed = check_package_installed(dep["name"], dep.get("import"), dep.get("version"))
            color = "green" if installed else "red"
            self.dep_labels[i][0].config(foreground=color)  # pkg
            self.dep_labels[i][1].config(foreground=color)  # req
            self.dep_labels[i][2].config(foreground=color)  # purpose
            self.dep_labels[i][3].config(foreground=color, text="Installed" if installed else "Missing")  # status
            if installed:
                self.dep_labels[i][4].state(["disabled"])
                self.dep_vars[i].set(False)
            else:
                self.dep_labels[i][4].state(["!disabled"])

    def install_selected_dependencies(self):
        selected = [dep["name"] + dep["version"] if dep.get("version") else dep["name"]
                    for dep, var in zip(DEPENDENCY_INFO, self.dep_vars)
                    if var.get() and not check_package_installed(dep["name"], dep.get("import"), dep.get("version"))]
        if not selected:
            messagebox.showinfo("No Selection", "Please select at least one missing dependency to install.")
            return
        self._install_dependencies(selected)

    def install_missing_dependencies(self):
        # Install all missing dependencies (required and optional)
        missing = [dep["name"] + dep["version"] if dep.get("version") else dep["name"]
                  for dep in DEPENDENCY_INFO if not check_package_installed(dep["name"], dep.get("import"), dep.get("version"))]
        if not missing:
            return
        self._install_dependencies(missing)

    def _install_dependencies(self, dep_list):
        # Check if running as a frozen executable
        if getattr(sys, 'frozen', False):
            self.dependency_text.config(state=tk.NORMAL)
            self.dependency_text.delete(1.0, tk.END)
            self.dependency_text.insert(tk.END, "This is a packaged application. Dependencies are already included.\n")
            self.dependency_text.insert(tk.END, "No need to install additional packages.\n\n")
            self.dependency_text.insert(tk.END, "If you're experiencing issues, please try reinstalling the application.")
            self.dependency_text.see(tk.END)
            self.dependency_text.config(state=tk.DISABLED)
            messagebox.showinfo("Packaged Application", 
                               "This is a packaged application with all dependencies included.\n\n"
                               "No need to install additional packages.")
            return
            
        # Normal installation process for non-frozen environment
        self.dependency_install_btn.config(state=tk.DISABLED)
        self.dependency_selected_btn.config(state=tk.DISABLED)
        self.dependency_text.config(state=tk.NORMAL)
        self.dependency_text.delete(1.0, tk.END)
        self.dependency_text.insert(tk.END, f"Installing: {', '.join(dep_list)}\n\n")
        self.dependency_text.see(tk.END)
        self.dependency_text.update_idletasks()
        self.root.update_idletasks()
        def run_install():
            for dep in dep_list:
                self.dependency_text.insert(tk.END, f"Installing {dep}...\n")
                self.dependency_text.see(tk.END)
                self.dependency_text.update_idletasks()
                self.root.update_idletasks()
                try:
                    process = subprocess.Popen([
                        sys.executable, "-m", "pip", "install", dep
                    ], stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True)
                    for line in process.stdout:
                        self.dependency_text.insert(tk.END, line)
                        self.dependency_text.see(tk.END)
                        self.dependency_text.update_idletasks()
                        self.root.update_idletasks()
                    process.wait()
                    if process.returncode == 0:
                        self.dependency_text.insert(tk.END, f"{dep} installed successfully.\n\n")
                        # If pywin32, run post-install script
                        if dep.startswith("pywin32"):
                            self.dependency_text.insert(tk.END, "Running pywin32_postinstall...\n")
                            self.dependency_text.see(tk.END)
                            self.dependency_text.update_idletasks()
                            self.root.update_idletasks()
                            try:
                                subprocess.run([
                                    sys.executable, "-m", "pywin32_postinstall", "-install"
                                ], check=False)
                                self.dependency_text.insert(tk.END, "pywin32_postinstall completed.\n")
                            except Exception as e:
                                self.dependency_text.insert(
                                    tk.END,
                                    f"pywin32_postinstall not found or failed: {e}\n"
                                    "If you encounter issues, try running the postinstall script manually or consult pywin32 documentation.\n"
                                )
                    else:
                        self.dependency_text.insert(tk.END, f"Failed to install {dep}.\n\n")
                except Exception as e:
                    self.dependency_text.insert(tk.END, f"Error installing {dep}: {e}\n\n")
                self.dependency_text.see(tk.END)
                self.dependency_text.update_idletasks()
                self.root.update_idletasks()
            self.dependency_text.insert(tk.END, "\nInstallation complete. Please restart the application to use all features.")
            self.dependency_text.see(tk.END)
            self.dependency_text.update_idletasks()
            self.root.update_idletasks()
            self.refresh_dependency_status()
            self.dependency_install_btn.config(state=tk.NORMAL)
            self.dependency_selected_btn.config(state=tk.NORMAL)
            messagebox.showinfo("Restart Required", "Dependencies installed. Please restart the application.")
        import threading
        threading.Thread(target=run_install, daemon=True).start()

    def uninstall_installed_dependencies(self):
        # Uninstall all currently installed dependencies in DEPENDENCY_INFO
        installed = [dep["name"] for dep in DEPENDENCY_INFO if check_package_installed(dep["name"], dep.get("import"), dep.get("version"))]
        if not installed:
            messagebox.showinfo("No Installed Packages", "No dependencies are currently installed.")
            return
        confirm = messagebox.askyesno(
            "Confirm Uninstall",
            "Are you sure you want to uninstall ALL installed dependencies?\n\n"
            "This will remove all packages listed in the dependencies table.\n"
            "You may need to reinstall them to use the application."
        )
        if not confirm:
            return
        self._uninstall_dependencies(installed)

    def _uninstall_dependencies(self, dep_list):
        # Check if running as a frozen executable
        if getattr(sys, 'frozen', False):
            self.dependency_text.config(state=tk.NORMAL)
            self.dependency_text.delete(1.0, tk.END)
            self.dependency_text.insert(tk.END, "This is a packaged application. Dependencies cannot be uninstalled.\n")
            self.dependency_text.insert(tk.END, "The dependencies are embedded in the executable.\n\n")
            self.dependency_text.insert(tk.END, "If you're experiencing issues, please try reinstalling the application.")
            self.dependency_text.see(tk.END)
            self.dependency_text.config(state=tk.DISABLED)
            messagebox.showinfo("Packaged Application", 
                               "This is a packaged application with embedded dependencies.\n\n"
                               "Dependencies cannot be uninstalled from the executable.")
            return
            
        # Normal uninstallation process for non-frozen environment
        self.dependency_uninstall_btn.config(state=tk.DISABLED)
        self.dependency_install_btn.config(state=tk.DISABLED)
        self.dependency_selected_btn.config(state=tk.DISABLED)
        self.dependency_text.config(state=tk.NORMAL)
        self.dependency_text.delete(1.0, tk.END)
        self.dependency_text.insert(tk.END, f"Uninstalling: {', '.join(dep_list)}\n\n")
        self.dependency_text.see(tk.END)
        self.dependency_text.update_idletasks()
        self.root.update_idletasks()
        def run_uninstall():
            for dep in dep_list:
                self.dependency_text.insert(tk.END, f"Uninstalling {dep}...\n")
                self.dependency_text.see(tk.END)
                self.dependency_text.update_idletasks()
                self.root.update_idletasks()
                try:
                    process = subprocess.Popen([
                        sys.executable, "-m", "pip", "uninstall", "-y", dep
                    ], stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True)
                    for line in process.stdout:
                        self.dependency_text.insert(tk.END, line)
                        self.dependency_text.see(tk.END)
                        self.dependency_text.update_idletasks()
                        self.root.update_idletasks()
                    process.wait()
                    if process.returncode == 0:
                        self.dependency_text.insert(tk.END, f"{dep} uninstalled successfully.\n\n")
                    else:
                        self.dependency_text.insert(tk.END, f"Failed to uninstall {dep}.\n\n")
                except Exception as e:
                    self.dependency_text.insert(tk.END, f"Error uninstalling {dep}: {e}\n\n")
                self.dependency_text.see(tk.END)
                self.dependency_text.update_idletasks()
                self.root.update_idletasks()
            self.dependency_text.insert(tk.END, "\nUninstallation complete. You may need to reinstall dependencies to use the application.")
            self.dependency_text.see(tk.END)
            self.dependency_text.update_idletasks()
            self.root.update_idletasks()
            self.refresh_dependency_status()
            self.dependency_uninstall_btn.config(state=tk.NORMAL)
            self.dependency_install_btn.config(state=tk.NORMAL)
            self.dependency_selected_btn.config(state=tk.NORMAL)
            messagebox.showinfo("Uninstall Complete", "Dependencies uninstalled. Please reinstall them to use the application.")
        import threading
        threading.Thread(target=run_uninstall, daemon=True).start()

    def check_and_handle_dependencies(self):
        # If running as a frozen executable, bypass dependency check
        if getattr(sys, 'frozen', False):
            # Enable all tabs
            self.notebook.tab(1, state="normal")
            self.notebook.tab(2, state="normal")
            self.dependency_install_btn.config(state=tk.DISABLED)
            self.dependency_text.config(state=tk.NORMAL)
            self.dependency_text.delete(1.0, tk.END)
            self.dependency_text.insert(tk.END, "All dependencies are included, you can proceed to the application :)")
            self.dependency_text.config(state=tk.DISABLED)
            return
            
        # Normal dependency check for non-frozen environment
        missing = get_missing_dependencies()
        self.refresh_dependency_status()
        if missing:
            # Disable other tabs
            self.notebook.tab(1, state="disabled")
            self.notebook.tab(2, state="disabled")
            self.notebook.select(0)
            self.dependency_install_btn.config(state=tk.NORMAL)
            self.dependency_text.config(state=tk.NORMAL)
            self.dependency_text.delete(1.0, tk.END)
            self.dependency_text.insert(tk.END, "Please install the missing dependencies to use the application.\n")
            self.dependency_text.config(state=tk.DISABLED)
            # Show popup
            self.root.after(500, lambda: messagebox.showwarning(
                "Missing Dependencies",
                "Some required dependencies are missing. Please go to the 'Dependencies' tab to install them."
            ))
        else:
            # All dependencies present, enable all tabs
            self.notebook.tab(1, state="normal")
            self.notebook.tab(2, state="normal")
            self.dependency_install_btn.config(state=tk.DISABLED)
            self.dependency_text.config(state=tk.NORMAL)
            self.dependency_text.delete(1.0, tk.END)
            self.dependency_text.insert(tk.END, "It’s all working. Probably. Hit run and hope.")
            self.dependency_text.config(state=tk.DISABLED)

    def setup_mapping_tab(self, parent):
        """Setup the file selection and column mapping tab"""
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(1, weight=1)  # Give column mapping section more vertical space
        
        # Compact File Selection section - redesigned to take less space
        file_frame = ttk.Frame(parent)  # Removed LabelFrame to save space
        file_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 5))
        file_frame.columnconfigure(1, weight=1)
        
        # More compact file selection layout
        ttk.Label(file_frame, text="Reference File:").grid(row=0, column=0, sticky=tk.W, pady=2)
        file_entry1 = ttk.Entry(file_frame, textvariable=self.master_file_path)
        file_entry1.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=5)
        ttk.Button(file_frame, text="Browse", command=self.browse_master_file, width=8).grid(row=0, column=2)
        
        ttk.Label(file_frame, text="Target File:").grid(row=1, column=0, sticky=tk.W, pady=2)
        file_entry2 = ttk.Entry(file_frame, textvariable=self.secondary_file_path)
        file_entry2.grid(row=1, column=1, sticky=(tk.W, tk.E), padx=5)
        ttk.Button(file_frame, text="Browse", command=self.browse_secondary_file, width=8).grid(row=1, column=2)
        
        # Load files button on the same row to save space
        ttk.Button(file_frame, text="Load Files", command=self.load_files, width=12).grid(
            row=0, column=3, rowspan=2, padx=(10, 0), pady=5)
        
        # Improved Column Mapping section - with more vertical space available
        self.setup_improved_mapping(parent)

        # Add a smaller Results & Log section at the bottom
        # --- Header frame for Results & Log ---
        results_header = ttk.Frame(parent)
        results_header.grid(row=2, column=0, sticky=(tk.W, tk.E), pady=(8, 0), padx=2)
        results_header.columnconfigure(0, weight=1)
        ttk.Label(results_header, text="Results & Log", font=("Arial", 10, "bold")).grid(row=0, column=0, sticky=tk.W)
        ttk.Button(results_header, text="Clear Log", command=self.clear_logs, width=10).grid(row=0, column=1, sticky=tk.E, padx=(0, 2))
        # --- Log area below header ---
        results_frame = ttk.Frame(parent, padding="0")
        results_frame.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), padx=5)
        results_frame.columnconfigure(0, weight=1)
        results_frame.rowconfigure(0, weight=1)
        self.log_text_mapping = scrolledtext.ScrolledText(results_frame, height=8)
        self.log_text_mapping.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

    def setup_improved_mapping(self, parent):
        """Setup an improved column mapping interface with better layout"""
        # Initialize selection arrays
        self.selected_reference_primary = tk.StringVar()
        self.selected_target_columns = []
        self.selected_replace_columns = []
        self.selected_data_source = tk.StringVar()
        
        # Main mapping frame
        mapping_frame = ttk.LabelFrame(parent, text="Column Mapping", padding="10")
        mapping_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 5))
        mapping_frame.columnconfigure(0, weight=1)
        mapping_frame.columnconfigure(1, weight=1)
        mapping_frame.rowconfigure(1, weight=1)  # Give the selection section weight
    
        # Using a two-panel approach for better organization
        left_panel = ttk.Frame(mapping_frame)
        left_panel.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N), padx=(0, 5))
        left_panel.columnconfigure(1, weight=1)  # Combobox column gets extra space
        
        right_panel = ttk.Frame(mapping_frame)
        right_panel.grid(row=0, column=1, sticky=(tk.W, tk.E, tk.N))
        right_panel.columnconfigure(1, weight=1)  # Combobox column gets extra space
        
        # === LEFT PANEL: Reference file selections ===
        ttk.Label(left_panel, text="Reference File Columns", font=("Arial", 9, "bold")).grid(
            row=0, column=0, columnspan=2, sticky=tk.W, pady=(0, 5))
        
        ttk.Label(left_panel, text="Primary Match:").grid(row=1, column=0, sticky=tk.W, pady=2)
        self.ref_primary_combo = ttk.Combobox(left_panel, state="readonly")
        self.ref_primary_combo.grid(row=1, column=1, sticky=(tk.W, tk.E), pady=2)
        self.ref_primary_combo.bind('<<ComboboxSelected>>', self.on_reference_primary_selected)
        
        ttk.Label(left_panel, text="Data Source:").grid(row=2, column=0, sticky=tk.W, pady=2)
        self.ref_data_combo = ttk.Combobox(left_panel, state="readonly")
        self.ref_data_combo.grid(row=2, column=1, sticky=(tk.W, tk.E), pady=2)
        self.ref_data_combo.bind('<<ComboboxSelected>>', self.on_data_source_selected)
        
        # === RIGHT PANEL: Target file selections ===
        ttk.Label(right_panel, text="Target File Columns", font=("Arial", 9, "bold")).grid(
            row=0, column=0, columnspan=3, sticky=tk.W, pady=(0, 5))
        
        ttk.Label(right_panel, text="Primary Match:").grid(row=1, column=0, sticky=tk.W, pady=2)
        self.target_primary_combo = ttk.Combobox(right_panel, state="readonly")
        self.target_primary_combo.grid(row=1, column=1, sticky=(tk.W, tk.E), padx=(5, 5), pady=2)
        ttk.Button(right_panel, text="Add", command=self.on_target_primary_selected, width=8).grid(
            row=1, column=2, sticky=tk.E, pady=2)
        
        ttk.Label(right_panel, text="Additional Target:").grid(row=2, column=0, sticky=tk.W, pady=2)
        self.target_additional_combo = ttk.Combobox(right_panel, state="readonly")
        self.target_additional_combo.grid(row=2, column=1, sticky=(tk.W, tk.E), padx=(5, 5), pady=2)
        ttk.Button(right_panel, text="Add", command=self.add_target_column, width=8).grid(
            row=2, column=2, sticky=tk.E, pady=2)
        
        ttk.Label(right_panel, text="Replace Column:").grid(row=3, column=0, sticky=tk.W, pady=2)
        self.target_replace_combo = ttk.Combobox(right_panel, state="readonly")
        self.target_replace_combo.grid(row=3, column=1, sticky=(tk.W, tk.E), padx=(5, 5), pady=2)
        ttk.Button(right_panel, text="Add", command=self.add_replace_column, width=8).grid(
            row=3, column=2, sticky=tk.E, pady=2)
        
        # === SELECTION DISPLAY SECTION ===
        selection_section = ttk.Frame(mapping_frame)
        selection_section.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(10, 0))
        selection_section.columnconfigure(0, weight=1)
        selection_section.columnconfigure(1, weight=1)
        selection_section.columnconfigure(2, weight=1)
        selection_section.rowconfigure(0, weight=1)  # Allow expansion vertically
        
        # Target Columns Display with scrollbar
        target_frame = ttk.LabelFrame(selection_section, text="Selected Target Columns")
        target_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(0, 5))
        target_frame.columnconfigure(0, weight=1)
        target_frame.rowconfigure(0, weight=1)
        
        target_scroll = ttk.Scrollbar(target_frame, orient=tk.VERTICAL)
        self.target_listbox = tk.Listbox(target_frame, height=5, font=("Arial", 8), 
                                         yscrollcommand=target_scroll.set)
        target_scroll.config(command=self.target_listbox.yview)
        self.target_listbox.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        target_scroll.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.target_listbox.bind('<Double-Button-1>', self.remove_target_from_listbox)
        
        # Replace Columns Display with scrollbar
        replace_frame = ttk.LabelFrame(selection_section, text="Selected Replace Columns")
        replace_frame.grid(row=0, column=1, sticky=(tk.W, tk.E, tk.N, tk.S), padx=5)
        replace_frame.columnconfigure(0, weight=1)
        replace_frame.rowconfigure(0, weight=1)
        
        replace_scroll = ttk.Scrollbar(replace_frame, orient=tk.VERTICAL)
        self.replace_listbox = tk.Listbox(replace_frame, height=5, font=("Arial", 8),
                                          yscrollcommand=replace_scroll.set)
        replace_scroll.config(command=self.replace_listbox.yview)
        self.replace_listbox.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        replace_scroll.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.replace_listbox.bind('<Double-Button-1>', self.remove_replace_from_listbox)
        
        # Current Selection summary (Ref Primary and Data Source moved outside)
        summary_outer = ttk.Frame(selection_section)
        summary_outer.grid(row=0, column=2, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(5, 0))
        summary_outer.columnconfigure(0, weight=1)
        summary_outer.rowconfigure(1, weight=1)

        # Heading and labels outside the box
        heading_frame = ttk.Frame(summary_outer)
        heading_frame.grid(row=0, column=0, sticky=tk.W)
        ttk.Label(heading_frame, text="Current Selection", font=("Arial", 9, "bold")).pack(side=tk.LEFT)
        ttk.Label(heading_frame, text="  Ref Primary:", font=("Arial", 8, "bold")).pack(side=tk.LEFT)
        self.ref_primary_label = ttk.Label(heading_frame, text="None", foreground="blue", font=("Arial", 8))
        self.ref_primary_label.pack(side=tk.LEFT, padx=(2, 8))
        ttk.Label(heading_frame, text="Data Source:", font=("Arial", 8, "bold")).pack(side=tk.LEFT)
        self.data_source_label = ttk.Label(heading_frame, text="None", foreground="green", font=("Arial", 8))
        self.data_source_label.pack(side=tk.LEFT, padx=(2, 0))

        # Mapping info box with scrollbar
        summary_frame = ttk.LabelFrame(summary_outer, text="Data Mapping Chart")
        summary_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(2, 0))
        summary_frame.columnconfigure(0, weight=1)
        summary_frame.rowconfigure(0, weight=1)
        mapping_scroll = ttk.Scrollbar(summary_frame, orient=tk.VERTICAL)
        self.mapping_info = tk.Text(summary_frame, height=6, width=28, font=("Arial", 8), wrap=tk.WORD, yscrollcommand=mapping_scroll.set)
        self.mapping_info.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        mapping_scroll.config(command=self.mapping_info.yview)
        mapping_scroll.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.mapping_info.config(state=tk.DISABLED)
        
        # === CONTROL BUTTONS: Moved to bottom for better visibility ===
        control_frame = ttk.Frame(mapping_frame)
        control_frame.grid(row=2, column=0, columnspan=2, sticky=tk.E, pady=(10, 0))
        
        ttk.Button(control_frame, text="Auto-Detect Columns", command=self.auto_detect_columns, 
                   width=18).pack(side=tk.RIGHT, padx=5)
        ttk.Button(control_frame, text="Clear All", command=self.clear_all_selections, 
                   width=10).pack(side=tk.RIGHT, padx=5)

    def setup_options_tab(self, parent):
        """Setup the options and processing tab with a compact 2x2 grid and a tall results/logs section"""
        parent.columnconfigure(0, weight=1)
        parent.columnconfigure(1, weight=1)
        parent.rowconfigure(0, weight=1)
        parent.rowconfigure(1, weight=1)
        parent.rowconfigure(2, weight=3)  # For tall log area

        # --- 2x2 Grid for Options ---
        # Top-left: Matching Options
        options_frame = ttk.LabelFrame(parent, text="Matching Options", padding="10")
        options_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=5, pady=5)
        options_frame.columnconfigure(0, weight=1)

        self.fuzzy_matching = tk.BooleanVar(value=True)
        ttk.Checkbutton(options_frame, text="Enable fuzzy name matching", variable=self.fuzzy_matching).grid(row=0, column=0, sticky=tk.W, pady=2)
        threshold_frame = ttk.Frame(options_frame)
        threshold_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=2)
        ttk.Label(threshold_frame, text="Similarity threshold (0.1-1.0):").pack(side=tk.LEFT)
        self.similarity_threshold = tk.DoubleVar(value=0.8)
        similarity_spin = ttk.Spinbox(threshold_frame, from_=0.1, to=1.0, increment=0.1, width=6, textvariable=self.similarity_threshold)
        similarity_spin.pack(side=tk.LEFT, padx=(5, 0))

        # Top-right: Multi-Value Target Column Options
        multivalue_frame = ttk.LabelFrame(parent, text="Multi-Value Target Column Options", padding="10")
        multivalue_frame.grid(row=0, column=1, sticky=(tk.W, tk.E, tk.N, tk.S), padx=5, pady=5)
        multivalue_frame.columnconfigure(0, weight=1)

        self.enable_multivalue = tk.BooleanVar(value=False)
        ttk.Checkbutton(multivalue_frame, text="Enable multi-value processing", variable=self.enable_multivalue, command=self.toggle_multivalue_options).grid(row=0, column=0, sticky=tk.W, pady=2)
        delimiter_frame = ttk.Frame(multivalue_frame)
        delimiter_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=2)
        ttk.Label(delimiter_frame, text="Delimiter:").pack(side=tk.LEFT)
        self.target_delimiter = tk.StringVar(value=",")
        delimiter_entry = ttk.Entry(delimiter_frame, textvariable=self.target_delimiter, width=4)
        delimiter_entry.pack(side=tk.LEFT, padx=(5, 0))
        common_delims = [(",", "Comma"), (";", "Semicolon"), ("|", "Pipe"), (" ", "Space")]
        for delim, name in common_delims:
            ttk.Button(delimiter_frame, text=name, width=7, command=lambda d=delim: self.target_delimiter.set(d)).pack(side=tk.LEFT, padx=1)
        self.delimiter_frame = delimiter_frame

        # Bottom-left: Output Options
        structure_frame = ttk.LabelFrame(parent, text="Output Options", padding="10")
        structure_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=5, pady=5)
        structure_frame.columnconfigure(0, weight=1)
        self.preserve_structure = tk.BooleanVar(value=True)
        ttk.Checkbutton(structure_frame, text="Preserve original file structure (only update matched columns)", variable=self.preserve_structure).grid(row=0, column=0, sticky=tk.W, pady=2)

        # Bottom-right: Processing & Export
        process_export_frame = ttk.LabelFrame(parent, text="Processing & Export", padding="10")
        process_export_frame.grid(row=1, column=1, sticky=(tk.W, tk.E, tk.N, tk.S), padx=5, pady=5)
        process_export_frame.columnconfigure(0, weight=1)
        ttk.Button(process_export_frame, text="Process & Match Data", command=self.process_data, style="Accent.TButton", width=22).grid(row=0, column=0, pady=(0, 8), sticky=tk.EW)
        ttk.Button(process_export_frame, text="Export Results to Excel/CSV", command=self.export_results, width=22).grid(row=1, column=0, pady=(0, 2), sticky=tk.EW)

        # --- Move: Replace Column Values Section above log ---
        replace_col_frame = ttk.LabelFrame(parent, text="Replace Entire Column in Processed Target File", padding="10")
        replace_col_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), padx=5, pady=(5, 10))
        replace_col_frame.columnconfigure(1, weight=1)
        ttk.Label(replace_col_frame, text="Column:").grid(row=0, column=0, sticky=tk.W)
        self.replace_col_combo = ttk.Combobox(replace_col_frame, state="readonly", width=28)
        self.replace_col_combo.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=5)
        ttk.Label(replace_col_frame, text="Value:").grid(row=0, column=2, sticky=tk.W)
        self.replace_col_value = tk.StringVar()
        ttk.Entry(replace_col_frame, textvariable=self.replace_col_value, width=18).grid(row=0, column=3, sticky=(tk.W, tk.E), padx=5)
        self.replace_col_btn = ttk.Button(replace_col_frame, text="Replace Column", command=self.replace_entire_column, state="disabled")
        self.replace_col_btn.grid(row=0, column=4, padx=5)
        self.preview_col_btn = ttk.Button(replace_col_frame, text="Preview", command=self.preview_replace_column, state="disabled")
        self.preview_col_btn.grid(row=0, column=5, padx=5)
        self.undo_col_btn = ttk.Button(replace_col_frame, text="Undo", command=self.undo_replace_column, state="disabled")
        self.undo_col_btn.grid(row=0, column=6, padx=5)
        self._undo_col_data = None
        self._undo_col_name = None

        # --- Tall Results/Logs Section ---
        # --- Header frame for Results & Log ---
        results_header = ttk.Frame(parent)
        results_header.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E), padx=5, pady=(10, 0))
        results_header.columnconfigure(0, weight=1)
        ttk.Label(results_header, text="Results & Log", font=("Arial", 10, "bold")).grid(row=0, column=0, sticky=tk.W)
        ttk.Button(results_header, text="Clear Log", command=self.clear_logs, width=10).grid(row=0, column=1, sticky=tk.E, padx=(0, 2))
        # --- Log area below header ---
        results_frame = ttk.Frame(parent, padding="0")
        results_frame.grid(row=4, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), padx=5)
        results_frame.columnconfigure(0, weight=1)
        results_frame.rowconfigure(0, weight=1)
        self.log_text = scrolledtext.ScrolledText(results_frame, height=12)
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.toggle_multivalue_options()

    def update_replace_col_combo(self):
        # Update the replace_col_combo dropdown with columns from self.secondary_work
        if hasattr(self, 'secondary_work') and self.secondary_work is not None:
            cols = list(self.secondary_work.columns)
            self.replace_col_combo['values'] = cols
            self.replace_col_btn.config(state="normal")
            self.preview_col_btn.config(state="normal")
            self.log_message(f"Available columns for replacement: {cols}", "INFO")
            self.replace_col_combo.update_idletasks()
        else:
            self.replace_col_combo['values'] = []
            self.replace_col_btn.config(state="disabled")
            self.preview_col_btn.config(state="disabled")
            self.undo_col_btn.config(state="disabled")

    def preview_replace_column(self):
        # Show a preview of the first 10 rows as they would look after replacement
        if not hasattr(self, 'secondary_work') or self.secondary_work is None:
            messagebox.showwarning("No Data", "No processed target file available. Please process data first.")
            return
        col = self.replace_col_combo.get()
        val = self.replace_col_value.get()
        if not col:
            messagebox.showwarning("No Column Selected", "Please select a column to preview.")
            return
        preview_df = self.secondary_work.copy()
        preview_df[col] = val
        preview_rows = preview_df[[col]].head(10)
        preview_text = preview_rows.to_string(index=True, header=True)
        messagebox.showinfo("Preview Replacement", f"Preview of column '{col}' after replacement (first 10 rows):\n\n{preview_text}")
        # self.log_message(f"Previewed replacement for column '{col}' with value '{val}'.", "INFO")

    def replace_entire_column(self):
        # Replace all values in the selected column with the specified value
        if not hasattr(self, 'secondary_work') or self.secondary_work is None:
            messagebox.showwarning("No Data", "No processed target file available. Please process data first.")
            return
        col = self.replace_col_combo.get()
        val = self.replace_col_value.get()
        if not col:
            messagebox.showwarning("No Column Selected", "Please select a column to replace.")
            return
        # Save for undo
        self._undo_col_data = self.secondary_work[col].copy()
        self._undo_col_name = col
        self.secondary_work[col] = val
        self.log_message(f"All values in column '{col}' replaced with '{val}'.", "INFO")
        messagebox.showinfo("Column Replaced", f"All values in column '{col}' have been replaced with '{val}'.")
        self.undo_col_btn.config(state="normal")

    def undo_replace_column(self):
        # Undo the last column replacement
        if not hasattr(self, 'secondary_work') or self.secondary_work is None or self._undo_col_data is None or self._undo_col_name is None:
            messagebox.showwarning("Nothing to Undo", "No column replacement to undo.")
            return
        self.secondary_work[self._undo_col_name] = self._undo_col_data
        self.log_message(f"Undo: Restored previous values for column '{self._undo_col_name}'.", "INFO")
        messagebox.showinfo("Undo Complete", f"Previous values for column '{self._undo_col_name}' have been restored.")
        self._undo_col_data = None
        self._undo_col_name = None
        self.undo_col_btn.config(state="disabled")

    def toggle_multivalue_options(self):
        """Enable/disable multi-value delimiter options based on checkbox"""
        state = 'normal' if self.enable_multivalue.get() else 'disabled'
        for widget in self.delimiter_frame.winfo_children():
            if isinstance(widget, (ttk.Entry, ttk.Button)):
                widget.configure(state=state)

    def on_data_source_selected(self, event=None):
        """Handle data source column selection - only one allowed"""
        selected = self.ref_data_combo.get()
        if selected:
            self.selected_data_source.set(selected)
            self.data_source_label.config(text=selected)
            self.update_mapping_display()
            self.log_message(f"Data source column selected: {selected}")

    def on_reference_primary_selected(self, event=None):
        """Handle reference primary column selection"""
        selected = self.ref_primary_combo.get()
        if selected:
            self.selected_reference_primary.set(selected)
            self.ref_primary_label.config(text=selected)
            self.update_mapping_display()
            self.log_message(f"Reference primary column selected: {selected}")

    def on_target_primary_selected(self, event=None):
        """Handle target primary column selection"""
        selected = self.target_primary_combo.get()
        if selected and selected not in self.selected_target_columns:
            self.selected_target_columns.insert(0, selected)  # Insert at beginning as it's primary
            self.update_target_display()
            self.update_mapping_display()
            self.target_primary_combo.set("")  # Clear selection
            self.log_message(f"Target primary column selected: {selected}")

    def add_target_column(self):
        """Add additional target column for comparison"""
        selected = self.target_additional_combo.get()
        if selected and selected not in self.selected_target_columns:
            self.selected_target_columns.append(selected)
            self.update_target_display()
            self.update_mapping_display()
            self.target_additional_combo.set("")  # Clear selection
            self.log_message(f"Additional target column added: {selected}")

    def add_replace_column(self):
        """Add column to replace list"""
        selected = self.target_replace_combo.get()
        if selected and selected not in self.selected_replace_columns:
            self.selected_replace_columns.append(selected)
            self.update_replace_display()
            self.update_mapping_display()
            self.target_replace_combo.set("")  # Clear selection
            self.log_message(f"Replace column added: {selected}")

    def update_target_display(self):
        """Update target columns listbox"""
        self.target_listbox.delete(0, tk.END)
        for i, col in enumerate(self.selected_target_columns):
            display_text = f"{i+1}. {col}"
            if i == 0:
                display_text += " (Primary)"
            self.target_listbox.insert(tk.END, display_text)

    def update_replace_display(self):
        """Update replace columns listbox"""
        self.replace_listbox.delete(0, tk.END)
        for i, col in enumerate(self.selected_replace_columns):
            self.replace_listbox.insert(tk.END, f"{i+1}. {col}")

    def update_mapping_display(self):
        """Update the mapping information display"""
        self.mapping_info.config(state=tk.NORMAL)
        self.mapping_info.delete(1.0, tk.END)
        
        mapping_text = ""
        
        # Show individual target to replace mapping
        if self.selected_target_columns and self.selected_replace_columns:
            mapping_text += "Individual Mapping:\n"
            max_items = min(len(self.selected_target_columns), len(self.selected_replace_columns))
            for i in range(max_items):
                # Show full column names instead of truncating
                target = self.selected_target_columns[i]
                replace = self.selected_replace_columns[i]
                mapping_text += f"{i+1}. {target} →>> {replace}\n"
            
            # Show unmapped items
            if len(self.selected_target_columns) > len(self.selected_replace_columns):
                for i in range(len(self.selected_replace_columns), len(self.selected_target_columns)):
                    target = self.selected_target_columns[i]
                    mapping_text += f"{i+1}. {target} →>> [No Replace]\n"
            elif len(self.selected_replace_columns) > len(self.selected_target_columns):
                for i in range(len(self.selected_target_columns), len(self.selected_replace_columns)):
                    replace = self.selected_replace_columns[i]
                    mapping_text += f"{i+1}. [No Target] →>> {replace}\n"
        
        # Show data source info
        data_source = self.selected_data_source.get()
        if data_source:
            data_short = data_source
            mapping_text += f"\nData Source: {data_short}"
        
        self.mapping_info.insert(1.0, mapping_text)
        self.mapping_info.config(state=tk.DISABLED)

    def remove_target_from_listbox(self, event=None):
        """Remove selected target column via double-click"""
        selection = self.target_listbox.curselection()
        if selection:
            index = selection[0]
            if index < len(self.selected_target_columns):
                removed_col = self.selected_target_columns.pop(index)
                self.update_target_display()
                self.update_mapping_display()
                self.log_message(f"Target column removed: {removed_col}")

    def remove_replace_from_listbox(self, event=None):
        """Remove selected replace column via double-click"""
        selection = self.replace_listbox.curselection()
        if selection:
            index = selection[0]
            if index < len(self.selected_replace_columns):
                removed_col = self.selected_replace_columns.pop(index)
                self.update_replace_display()
                self.update_mapping_display()
                self.log_message(f"Replace column removed: {removed_col}")

    def clear_all_selections(self):
        """Clear all column selections"""
        self.selected_reference_primary.set("")
        self.selected_data_source.set("")
        self.ref_primary_label.config(text="None")
        self.data_source_label.config(text="None")
        self.selected_target_columns.clear()
        self.selected_replace_columns.clear()
        
        self.update_target_display()
        self.update_replace_display()
        self.update_mapping_display()
        
        self.log_message("All selections cleared")

    def auto_detect_columns(self):
        """Auto-detect and suggest column mappings"""
        if self.master_df is None or self.secondary_df is None:
            messagebox.showwarning("No Data", "Please load files first")
            return
        
        # Clear existing selections
        self.clear_all_selections()
        
        # Auto-detect columns
        master_cols = list(self.master_df.columns)
        secondary_cols = list(self.secondary_df.columns)
        
        # Look for name columns in reference
        name_keywords = ['name', 'naam', 'full name', 'fullname', 'participant']
        for col in master_cols:
            col_lower = col.lower()
            for keyword in name_keywords:
                if keyword in col_lower:
                    self.ref_primary_combo.set(col)
                    self.on_reference_primary_selected()
                    break
            if self.selected_reference_primary.get():
                break
        
        # Look for name columns in target
        for col in secondary_cols:
            col_lower = col.lower()
            for keyword in name_keywords:
                if keyword in col_lower:
                    self.target_primary_combo.set(col)
                    self.on_target_primary_selected()
                    break
            if self.selected_target_columns:
                break
        
        # Auto-add ID columns
        id_keywords = ['id', 'signum', 'employee id', 'emp id', 'userid', 'user id']
        
        # Add ID from reference as data source
        for col in master_cols:
            col_lower = col.lower()
            for keyword in id_keywords:
                if keyword in col_lower:
                    self.ref_data_combo.set(col)
                    self.on_data_source_selected()
                    break
            if self.selected_data_source.get():
                break
        
        # Add ID from target to replace
        for col in secondary_cols:
            col_lower = col.lower()
            for keyword in id_keywords:
                if keyword in col_lower:
                    self.target_replace_combo.set(col)
                    self.add_replace_column()
                    break
        
        self.log_message("Auto-detection completed")

    def update_column_dropdowns(self):
        """Update column dropdown menus with available columns"""
        if self.master_df is not None:
            master_cols = list(self.master_df.columns)
            self.ref_primary_combo['values'] = master_cols
            self.ref_data_combo['values'] = master_cols
            
        if self.secondary_df is not None:
            secondary_cols = list(self.secondary_df.columns)
            self.target_primary_combo['values'] = secondary_cols
            self.target_additional_combo['values'] = secondary_cols
            self.target_replace_combo['values'] = secondary_cols

    def log_message(self, message, level="INFO"):
        """Add message to log(s) with timestamp"""
        import datetime
        timestamp = datetime.datetime.now().strftime("%H:%M:%S")
        log_entry = f"[{timestamp}] {level}: {message}\n"
        # Write to both logs if both exist, else to the one present
        if hasattr(self, 'log_text') and self.log_text:
            self.log_text.insert(tk.END, log_entry)
            self.log_text.see(tk.END)
        if hasattr(self, 'log_text_mapping') and self.log_text_mapping:
            self.log_text_mapping.insert(tk.END, log_entry)
            self.log_text_mapping.see(tk.END)
        self.root.update_idletasks()
        
    def browse_master_file(self):
        filename = filedialog.askopenfilename(
            title="Select Master File",
            filetypes=[
                ("Excel files", "*.xlsx *.xls *.xlsm *.xlsb"),
                ("All files", "*.*")
            ]
        )
        if filename:
            self.master_file_path.set(filename)
            
    def browse_secondary_file(self):
        filename = filedialog.askopenfilename(
            title="Select Secondary File",
            filetypes=[
                ("Excel files", "*.xlsx *.xls *.xlsm *.xlsb"),
                ("All files", "*.*")
            ]
        )
        if filename:
            self.secondary_file_path.set(filename)
            
    def load_files(self):
        """Load and analyze both Excel files"""
        try:
            self.log_text.delete(1.0, tk.END)
            self.log_message("Starting file loading process...")
            
            # Validate file paths
            master_path = self.master_file_path.get().strip()
            secondary_path = self.secondary_file_path.get().strip()
            
            if not master_path or not secondary_path:
                raise ValueError("Please select both master and secondary files")
                
            if not os.path.exists(master_path):
                raise FileNotFoundError(f"Master file not found: {master_path}")
                
            if not os.path.exists(secondary_path):
                raise FileNotFoundError(f"Secondary file not found: {secondary_path}")
            
            # Load master file
            self.log_message(f"Loading master file: {os.path.basename(master_path)}")
            self.master_df = self.load_excel_file(master_path)
            self.log_message(f"Master file loaded successfully. Shape: {self.master_df.shape}")
            
            # Load secondary file
            self.log_message(f"Loading secondary file: {os.path.basename(secondary_path)}")
            self.secondary_df = self.load_excel_file(secondary_path)
            self.log_message(f"Secondary file loaded successfully. Shape: {self.secondary_df.shape}")
            
            # Update column dropdowns
            self.update_column_dropdowns()
            
            self.log_message("Files loaded successfully! Please select column mappings.")
            
        except Exception as e:
            error_msg = f"Error loading files: {str(e)}"
            self.log_message(error_msg, "ERROR")
            messagebox.showerror("Error", error_msg)
            
    def is_valid_excel_file(self, filepath):
        """Check if file appears to be a valid Excel file"""
        try:
            # Check file signature (magic numbers)
            with open(filepath, 'rb') as f:
                header = f.read(8)
                
            # Excel file signatures
            xlsx_sig = b'PK\x03\x04'  # XLSX files are ZIP files
            xls_sig = b'\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1'  # XLS files
            
            if header.startswith(xlsx_sig) or header.startswith(xls_sig):
                return True
                
            return False
        except Exception as e:
            self.log_message(f"Error checking file validity: {str(e)}", "WARNING")
            return False
    
    def load_excel_file(self, filepath):
        """Enhanced Excel file loader with enterprise support"""
        try:
            # When running as a frozen executable, ensure all required modules are imported
            if getattr(sys, 'frozen', False):
                # Import all required modules at once to avoid issues
                import pandas as pd
                import numpy as np
                import openpyxl
                import xlrd
                try:
                    import pyxlsb
                except ImportError:
                    pass
                try:
                    import win32com.client
                    import pythoncom
                except ImportError:
                    pass
                
                # Make pandas available globally
                global pd
            
            # Normalize file path to handle OneDrive paths better
            filepath = self.normalize_path(filepath)
            
            # Get file extension
            file_ext = Path(filepath).suffix.lower()
            password = None
            
            self.log_message(f"Detecting file format: {file_ext}")
            
            # Check if file is actually an Excel file
            if not self.is_valid_excel_file(filepath):
                self.log_message("File doesn't appear to be a valid Excel file. Checking if password protected...", "WARNING")
                # Ask for password if the file might be password protected
                password = self.prompt_for_password(filepath)
            
            # First try: Standard pandas methods
            try:
                df = self.load_with_pandas(filepath, password)
                if df is not None:
                    return df
            except Exception as e:
                self.log_message(f"Standard pandas loading failed: {str(e)}", "WARNING")
            
            # Second try: Use win32com if available (Windows only)
            if HAS_WIN32COM and platform.system() == 'Windows':
                try:
                    df = self.load_with_win32com(filepath, password)
                    if df is not None:
                        return df
                except Exception as e:
                    self.log_message(f"COM interface loading failed: {str(e)}", "WARNING")
            
            # Attempt to copy the file to a local temp location and try again
            try:
                self.log_message("Attempting to copy file to local temp location...", "INFO")
                temp_file = self.copy_to_temp(filepath)
                if temp_file:
                    try:
                        df = self.load_with_pandas(temp_file, password)
                        if df is not None:
                            # Clean up temp file
                            try:
                                os.remove(temp_file)
                            except:
                                pass
                            return df
                    except:
                        # Clean up temp file
                        try:
                            os.remove(temp_file)
                        except:
                            pass
            except Exception as e:
                self.log_message(f"Temp file approach failed: {str(e)}", "WARNING")
            
            # If we get here, all methods failed
            raise ValueError("Failed to load Excel file with any method. The file might be corrupted, password-protected, or in an unsupported format.")
            
        except Exception as e:
            raise Exception(f"Failed to load {filepath}: {str(e)}")
    
    def normalize_path(self, filepath):
        """Normalize file path to handle OneDrive paths better"""
        # Replace double slashes with single slash
        filepath = filepath.replace('//', '/')
        
        # Convert to proper Windows path if on Windows
        if platform.system() == 'Windows':
            filepath = filepath.replace('/', '\\')
            
        # Handle special OneDrive paths
        if platform.system() == 'Windows' and 'onedrive' in filepath.lower():
            # Try to ensure the path is absolute
            if not os.path.isabs(filepath):
                filepath = os.path.abspath(filepath)
                
            # Verify the file exists at this path
            if not os.path.exists(filepath):
                self.log_message(f"File doesn't exist at normalized path: {filepath}", "WARNING")
                # Try to get the OneDrive root folder
                onedrive_path = self.get_onedrive_path()
                if onedrive_path:
                    # Try to construct path using OneDrive root
                    onedrive_idx = filepath.lower().find('onedrive')
                    if onedrive_idx >= 0:
                        relative_path = filepath[onedrive_idx + len('onedrive'):]
                        # Remove any leading separators
                        relative_path = relative_path.lstrip('\\/')
                        new_path = os.path.join(onedrive_path, relative_path)
                        self.log_message(f"Trying OneDrive path: {new_path}", "INFO")
                        if os.path.exists(new_path):
                            return new_path
        
        return filepath
    
    def get_onedrive_path(self):
        """Try to get the OneDrive root path"""
        if platform.system() != 'Windows':
            return None
            
        # Common OneDrive locations
        possible_paths = [
            os.path.join(os.path.expanduser('~'), 'OneDrive'),
            os.path.join(os.path.expanduser('~'), 'OneDrive - Ericsson')
        ]
        
        for path in possible_paths:
            if os.path.exists(path) and os.path.isdir(path):
                return path
                
        return None
    
    def copy_to_temp(self, filepath):
        """Copy file to temporary location"""
        try:
            # Create temp directory if it doesn't exist
            temp_dir = os.path.join(os.path.expanduser('~'), 'temp_excel_files')
            os.makedirs(temp_dir, exist_ok=True)
            
            # Create a unique temp file name
            filename = os.path.basename(filepath)
            temp_file = os.path.join(temp_dir, f"temp_{int(time.time())}_{filename}")
            
            # Copy the file
            import shutil
            shutil.copy2(filepath, temp_file)
            
            self.log_message(f"File copied to temporary location: {temp_file}", "INFO")
            return temp_file
        except Exception as e:
            self.log_message(f"Failed to copy to temp location: {str(e)}", "WARNING")
            return None
    
    def check_and_install_dependencies(self):
        """Check for missing dependencies and offer to install them"""
        missing_deps = []
        
        # Check for odfpy
        try:
            import odf
        except ImportError:
            missing_deps.append("odfpy")
        
        # Check for pyxlsb
        try:
            import pyxlsb
        except ImportError:
            missing_deps.append("pyxlsb")
        
        # If there are missing dependencies, offer to install them
        if missing_deps:
            deps_str = ", ".join(missing_deps)
            self.log_message(f"Missing optional dependencies: {deps_str}", "WARNING")
            
            response = messagebox.askyesno(
                "Install Dependencies?",
                f"The following optional dependencies are missing: {deps_str}\n\n"
                "Would you like to install them now?"
            )
            
            if response:
                try:
                    self.log_message(f"Installing dependencies: {deps_str}", "INFO")
                    for dep in missing_deps:
                        subprocess.check_call([sys.executable, "-m", "pip", "install", dep])
                    self.log_message("Dependencies installed successfully", "INFO")
                    
                    # Inform user they may need to restart
                    messagebox.showinfo(
                        "Restart Recommended",
                        "Dependencies were installed successfully.\n"
                        "It's recommended to restart the application for changes to take effect."
                    )
                except Exception as e:
                    self.log_message(f"Failed to install dependencies: {str(e)}", "ERROR")
    
    def load_with_pandas(self, filepath, password=None):
        """Load Excel file using pandas with password support"""
        # Ensure pandas is imported
        if getattr(sys, 'frozen', False) and not 'pd' in globals():
            global pd
            import pandas as pd
            
        file_ext = Path(filepath).suffix.lower()
        
        if file_ext in ['.xlsx', '.xls', '.xlsm', '.xlsb']:
            try:
                # First try with default settings
                engine = None
                excel_kwargs = {}
                
                # Set engine based on file extension
                if file_ext == '.xlsx' or file_ext == '.xlsm':
                    engine = 'openpyxl'
                    # Ensure openpyxl is imported
                    if getattr(sys, 'frozen', False):
                        import openpyxl
                elif file_ext == '.xls':
                    engine = 'xlrd'
                    # Ensure xlrd is imported
                    if getattr(sys, 'frozen', False):
                        import xlrd
                elif file_ext == '.xlsb':
                    engine = 'pyxlsb'
                    # Ensure pyxlsb is imported
                    if getattr(sys, 'frozen', False):
                        try:
                            import pyxlsb
                        except ImportError:
                            pass  # Optional dependency
                
                if engine:
                    excel_kwargs['engine'] = engine
                
                # Handle password separately for openpyxl
                if password and engine == 'openpyxl':
                    # For openpyxl, we need to use a different approach
                    try:
                        import openpyxl
                        self.log_message("Loading password-protected file with openpyxl", "INFO")
                        wb = openpyxl.load_workbook(filepath, data_only=True, password=password)
                        sheet_names = wb.sheetnames
                        
                        if len(sheet_names) > 1:
                            sheet_name = self.select_sheet(sheet_names, os.path.basename(filepath))
                            if not sheet_name:
                                raise ValueError("No sheet selected")
                            ws = wb[sheet_name]
                        else:
                            ws = wb.active
                        
                        # Convert worksheet to dataframe preserving original data types
                        data = ws.values
                        cols = next(data)
                        data = list(data)
                        df = pd.DataFrame(data, columns=cols)
                        
                        # Clean column names but preserve all data as-is
                        df.columns = df.columns.astype(str).str.strip()
                        # Don't remove empty rows or convert data types - preserve structure
                        return df
                    except ImportError:
                        self.log_message("openpyxl not available for password handling", "WARNING")
                
                # Standard approach without password - preserve data types
                if len(pd.ExcelFile(filepath, **excel_kwargs).sheet_names) > 1:
                    sheet_name = self.select_sheet(pd.ExcelFile(filepath, **excel_kwargs).sheet_names, os.path.basename(filepath))
                    if not sheet_name:
                        raise ValueError("No sheet selected")
                    # Read without date parsing to preserve original formats
                    df = pd.read_excel(filepath, sheet_name=sheet_name, dtype=str, **excel_kwargs)
                else:
                    # Read without date parsing to preserve original formats
                    df = pd.read_excel(filepath, dtype=str, **excel_kwargs)
                
                # Clean column names but preserve data structure and formats
                df.columns = df.columns.astype(str).str.strip()
                # Don't remove empty rows - preserve original structure exactly
                return df
                
            except Exception as e:
                if password is None and "password" in str(e).lower():
                    # Try again with a password
                    password = self.prompt_for_password(filepath)
                    if password:
                        return self.load_with_pandas(filepath, password)
                
                # Try different engines if the default fails
                engines_to_try = []
                if file_ext == '.xlsx' or file_ext == '.xlsm':
                    engines_to_try = ['openpyxl', 'xlrd', 'odf']
                elif file_ext == '.xls':
                    engines_to_try = ['xlrd', 'openpyxl']
                elif file_ext == '.xlsb':
                    engines_to_try = ['pyxlsb']
                
                for engine in engines_to_try:
                    try:
                        self.log_message(f"Trying with engine: {engine}", "INFO")
                        excel_kwargs = {'engine': engine}
                        
                        # Ensure the engine module is imported
                        if getattr(sys, 'frozen', False):
                            if engine == 'openpyxl':
                                import openpyxl
                            elif engine == 'xlrd':
                                import xlrd
                            elif engine == 'pyxlsb':
                                try:
                                    import pyxlsb
                                except ImportError:
                                    continue  # Skip this engine if not available
                            elif engine == 'odf':
                                try:
                                    import odf
                                except ImportError:
                                    continue  # Skip this engine if not available
                        
                        excel_file = pd.ExcelFile(filepath, **excel_kwargs)
                        sheet_names = excel_file.sheet_names
                        
                        if len(sheet_names) > 1:
                            sheet_name = self.select_sheet(sheet_names, os.path.basename(filepath))
                            if not sheet_name:
                                raise ValueError("No sheet selected")
                            # Preserve data types by reading as string
                            df = pd.read_excel(filepath, sheet_name=sheet_name, dtype=str, **excel_kwargs)
                        else:
                            # Preserve data types by reading as string
                            df = pd.read_excel(filepath, dtype=str, **excel_kwargs)
                        
                        # Clean column names but preserve data
                        df.columns = df.columns.astype(str).str.strip()
                        # Don't remove empty rows - preserve structure
                        return df
                    except Exception as engine_error:
                        self.log_message(f"Engine {engine} failed: {str(engine_error)}", "WARNING")
                
                # All engines failed
                raise ValueError("Failed to load with any pandas engine")
                
        elif file_ext == '.csv':
            # Try different encodings for CSV
            encodings = ['utf-8', 'latin1', 'cp1252', 'iso-8859-1']
            for encoding in encodings:
                try:
                    # Read CSV as string to preserve formats
                    df = pd.read_csv(filepath, encoding=encoding, dtype=str)
                    return df
                except Exception:
                    pass
            
            raise ValueError("Failed to load CSV with any encoding")
        
        else:
            raise ValueError(f"Unsupported file format: {file_ext}")
    
    def load_with_win32com(self, filepath, password=None):
        """Load Excel file using COM interface (Windows only)"""
        # Ensure pandas is imported
        if getattr(sys, 'frozen', False) and not 'pd' in globals():
            global pd
            import pandas as pd
            
        if not HAS_WIN32COM:
            raise ImportError("win32com not available")
        
        # Ensure pythoncom is imported
        if getattr(sys, 'frozen', False):
            import pythoncom
            import win32com.client
            
        pythoncom.CoInitialize()
        try:
            self.log_message("Attempting to load file with Excel COM interface...", "INFO")
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            
            # Create a temporary file path for saving
            temp_dir = os.path.join(os.path.expanduser('~'), 'temp_excel_files')
            os.makedirs(temp_dir, exist_ok=True)
            temp_file = os.path.join(temp_dir, f"converted_excel_{int(time.time())}.xlsx")
            
            try:
                # Open the workbook
                workbook = None
                try:
                    if password:
                        workbook = excel.Workbooks.Open(filepath, Password=password)
                    else:
                        workbook = excel.Workbooks.Open(filepath)
                except Exception as e:
                    self.log_message(f"Failed to open workbook: {str(e)}", "WARNING")
                    # Try again with ReadOnly
                    try:
                        workbook = excel.Workbooks.Open(filepath, ReadOnly=True)
                    except Exception as e2:
                        self.log_message(f"Failed to open workbook in read-only mode: {str(e2)}", "WARNING")
                        raise
                
                if not workbook:
                    raise ValueError("Failed to open workbook with COM interface")
                
                # Let user select sheet if multiple sheets
                if workbook.Sheets.Count > 1:
                    sheet_names = [sheet.Name for sheet in workbook.Sheets]
                    selected_sheet = self.select_sheet(sheet_names, os.path.basename(filepath))
                    if not selected_sheet:
                        raise ValueError("No sheet selected")
                    sheet = workbook.Sheets(selected_sheet)
                else:
                    sheet = workbook.Sheets(1)
                
                # Save to a new format that pandas can read
                if os.path.exists(temp_file):
                    os.remove(temp_file)
                
                # Export data to CSV as an alternative approach
                try:
                    # Try to export using SaveAs method
                    sheet.Activate()
                    # Try XlFileFormat.xlOpenXMLWorkbook (51) for XLSX
                    workbook.SaveAs(temp_file, 51)
                    workbook.Close(False)
                    
                    # Now load with pandas
                    df = pd.read_excel(temp_file)
                except Exception as e:
                    self.log_message(f"SaveAs failed: {str(e)}", "WARNING")
                    
                    # Alternative approach - export to CSV
                    csv_temp_file = temp_file.replace('.xlsx', '.csv')
                    sheet.Activate()
                    workbook.SaveAs(csv_temp_file, 6)  # 6 = CSV format
                    workbook.Close(False)
                    
                    # Now load with pandas
                    df = pd.read_csv(csv_temp_file)
                    
                    # Clean up CSV file
                    try:
                        os.remove(csv_temp_file)
                    except:
                        pass
                
                # Clean column names
                df.columns = df.columns.astype(str).str.strip()
                # Remove empty rows
                df = df.dropna(how='all')
                
                self.log_message("Successfully loaded file using Excel COM interface", "INFO")
                return df
                
            except Exception as e:
                raise Exception(f"COM interface error: {str(e)}")
            finally:
                # Clean up
                excel.Quit()
                if os.path.exists(temp_file):
                    try:
                        os.remove(temp_file)
                    except:
                        pass
                    
        except Exception as e:
            raise Exception(f"COM interface error: {str(e)}")
        finally:
            pythoncom.CoUninitialize()

    def prompt_for_password(self, filepath):
        """Prompt user for password if file might be password protected"""
        filename = os.path.basename(filepath)
        response = messagebox.askyesno(
            "Password Protected?", 
            f"The file '{filename}' might be password protected. Does it require a password?"
        )
        
        if response:
            password = simpledialog.askstring(
                "Password Required", 
                f"Enter password for '{filename}':", 
                show='*'
            )
            return password
        return None
    
    def select_sheet(self, sheet_names, filename):
        """Dialog to select sheet from multiple sheets"""
        dialog = tk.Toplevel(self.root)
        dialog.title(f"Select Sheet - {filename}")
        dialog.geometry("300x200")
        dialog.transient(self.root)
        dialog.grab_set()
        
        selected_sheet = None
        
        def on_select():
            nonlocal selected_sheet
            selection = sheet_listbox.curselection()
            if selection:
                selected_sheet = sheet_names[selection[0]]
                dialog.destroy()
        
        ttk.Label(dialog, text="Select a sheet:").pack(pady=10)
        
        sheet_listbox = tk.Listbox(dialog)
        for sheet in sheet_names:
            sheet_listbox.insert(tk.END, sheet)
        sheet_listbox.pack(pady=10, padx=20, fill=tk.BOTH, expand=True)
        
        button_frame = ttk.Frame(dialog)
        button_frame.pack(pady=10)
        ttk.Button(button_frame, text="Select", command=on_select).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Cancel", command=dialog.destroy).pack(side=tk.LEFT, padx=5)
        
        # Select first sheet by default
        sheet_listbox.selection_set(0)
        
        dialog.wait_window()
        return selected_sheet
        
    def update_column_dropdowns(self):
        """Update column dropdown menus with available columns"""
        if self.master_df is not None:
            master_cols = list(self.master_df.columns)
            self.ref_primary_combo['values'] = master_cols
            self.ref_data_combo['values'] = master_cols
            
        if self.secondary_df is not None:
            secondary_cols = list(self.secondary_df.columns)
            self.target_primary_combo['values'] = secondary_cols
            self.target_additional_combo['values'] = secondary_cols
            self.target_replace_combo['values'] = secondary_cols

    def process_data(self):
        """Process and match data between files using advanced mapping"""
        try:
            self.log_message("Starting advanced data processing...")
            
            # Validate inputs
            if self.master_df is None or self.secondary_df is None:
                raise ValueError("Please load both files first")
            
            # Validate advanced mapping selections
            reference_primary = self.selected_reference_primary.get()
            if not reference_primary:
                raise ValueError("Please select a reference primary match column")
            
            if not self.selected_target_columns:
                raise ValueError("Please select at least one target column for comparison")
            
            if self.preserve_structure.get() and not self.selected_replace_columns:
                raise ValueError("Please select columns to replace when preserve structure is enabled")
            
            data_source = self.selected_data_source.get()
            if not data_source:
                raise ValueError("Please select a data source column")
            
            self.log_message(f"Reference primary: {reference_primary}")
            self.log_message(f"Target columns: {self.selected_target_columns}")
            self.log_message(f"Replace columns: {self.selected_replace_columns}")
            self.log_message(f"Data source column: {data_source}")
            
            if self.enable_multivalue.get():
                delimiter = self.target_delimiter.get()
                self.log_message(f"Multi-value processing enabled with delimiter: '{delimiter}'")
            
            # Create working copies
            master_work = self.master_df.copy()
            self.secondary_work = self.secondary_df.copy()
            
            # Clean names for comparison
            master_work['clean_reference'] = master_work[reference_primary].apply(self.clean_name)
            
            # Clean target columns for comparison
            for i, target_col in enumerate(self.selected_target_columns):
                self.secondary_work[f'clean_target_{i}'] = self.secondary_work[target_col].apply(self.clean_name)
            
            # Statistics
            exact_matches = 0
            fuzzy_matches = 0
            no_matches = 0
            
            if self.preserve_structure.get():
                # Add tracking columns
                self.secondary_work['Match_Type'] = ""
                self.secondary_work['Matched_Reference_Name'] = ""
                self.secondary_work['Confidence'] = ""
                self.secondary_work['Matched_Column'] = ""
                
                # Process each row in secondary file
                self.log_message(f"Processing {len(self.secondary_work)} records from target file...")
                
                for idx, row in self.secondary_work.iterrows():
                    # Track matches for each target column individually
                    matches_found = {}  # target_column_index -> (match_info, data_source_value)
                    
                    # Check each target column against reference primary
                    for i, target_col in enumerate(self.selected_target_columns):
                        target_value = row[target_col]
                        
                        if pd.isna(target_value) or str(target_value).strip() == "":
                            continue
                        
                        # Handle multi-value processing if enabled
                        if self.enable_multivalue.get():
                            # Check if the value contains any delimiters (comma, semicolon, pipe)
                            target_str = str(target_value)
                            delimiters_to_check = [',', ';', '|']
                            actual_delimiter = None
                            
                            # Find which delimiter is actually used in this value
                            for delim in delimiters_to_check:
                                if delim in target_str:
                                    actual_delimiter = delim
                                    break
                            
                            # Process multi-value if we found a delimiter
                            if actual_delimiter:
                                target_values = [val.strip() for val in target_str.split(actual_delimiter) if val.strip()]
                                
                                if len(target_values) > 1:
                                    self.log_message(f"Processing multi-value field with delimiter '{actual_delimiter}': {target_values}")
                                    
                                    # Process multiple values
                                    matched_results = []
                                    match_types = []
                                    for target_val in target_values:
                                        match_result = self.find_match_for_value(target_val, master_work, reference_primary, data_source)
                                        if match_result:
                                            matched_results.append(match_result['data_value'])
                                            match_types.append(match_result['type'])
                                            self.log_message(f"  '{target_val}' matched to '{match_result['data_value']}'")
                                        else:
                                            matched_results.append("match unfound")
                                            match_types.append("NONE")
                                            self.log_message(f"  '{target_val}' not matched")
                                    
                                    if matched_results:
                                        # Use the same delimiter that was found in the original data
                                        combined_data_value = actual_delimiter.join(matched_results)
                                        
                                        # Determine overall match type
                                        has_exact = "EXACT" in match_types
                                        has_fuzzy = "FUZZY" in match_types
                                        match_count = len([r for r in matched_results if r != 'match unfound'])
                                        
                                        match_type = "EXACT" if has_exact else ("FUZZY" if has_fuzzy else "REVIEW")
                                        
                                        matches_found[i] = {
                                            'type': match_type,
                                            'matched_name': f"Multi-value: {match_count}/{len(target_values)} matched",
                                            'confidence': 1.0 if has_exact else (0.8 if has_fuzzy else 0.0),
                                            'column': target_col,
                                            'data_value': combined_data_value
                                        }
                                        continue
                        
                        # Single value processing (original logic)
                        target_clean_name = self.clean_name(target_value)
                        if not target_clean_name:
                            continue
                        
                        # Try exact match
                        exact_match = master_work[master_work['clean_reference'] == target_clean_name]
                        
                        if not exact_match.empty:
                            matches_found[i] = {
                                'type': 'EXACT',
                                'matched_name': exact_match.iloc[0][reference_primary],
                                'confidence': 1.0,
                                'column': target_col,
                                'data_value': exact_match.iloc[0].get(data_source, "")
                            }
                        
                        # Try fuzzy matching if enabled and no exact match
                        elif self.fuzzy_matching.get():
                            master_names = master_work['clean_reference'].dropna().tolist()
                            if master_names:
                                close_matches = get_close_matches(
                                    target_clean_name,
                                    master_names,
                                    n=1,
                                    cutoff=self.similarity_threshold.get()
                                )
                                
                                if close_matches:
                                    matched_clean_name = close_matches[0]
                                    fuzzy_match = master_work[master_work['clean_reference'] == matched_clean_name]
                                    
                                    if not fuzzy_match.empty:
                                        from difflib import SequenceMatcher
                                        confidence = SequenceMatcher(None, target_clean_name, matched_clean_name).ratio()
                                        
                                        matches_found[i] = {
                                            'type': 'FUZZY',
                                            'matched_name': fuzzy_match.iloc[0][reference_primary],
                                            'confidence': confidence,
                                            'column': target_col,
                                            'data_value': fuzzy_match.iloc[0].get(data_source, "")
                                        }
                    
                    # Update replace columns based on individual matchess
                    if matches_found:
                        # Update each replace column based on its corresponding target column
                        for target_idx, match_info in matches_found.items():
                            if target_idx < len(self.selected_replace_columns):
                                replace_col = self.selected_replace_columns[target_idx]
                                if replace_col in self.secondary_work.columns:
                                    self.secondary_work.at[idx, replace_col] = match_info['data_value']
                        
                        # Use the best match for tracking (highest confidence)
                        best_match = max(matches_found.values(), key=lambda x: x['confidence'])
                        
                        # Update tracking with best match
                        self.secondary_work.at[idx, 'Match_Type'] = best_match['type']
                        self.secondary_work.at[idx, 'Matched_Reference_Name'] = best_match['matched_name']
                        self.secondary_work.at[idx, 'Confidence'] = round(best_match['confidence'], 3)
                        self.secondary_work.at[idx, 'Matched_Column'] = best_match['column']
                        
                        if best_match['type'] == "EXACT":
                            exact_matches += 1
                        else:
                            fuzzy_matches += 1
                    else:
                        no_matches += 1
                        self.secondary_work.at[idx, 'Match_Type'] = "REVIEW"
                
                # Clean up temporary columns
                cols_to_drop = ['clean_reference'] + [f'clean_target_{i}' for i in range(len(self.selected_target_columns))]
                master_work = master_work.drop(columns=[col for col in cols_to_drop if col in master_work.columns])
                self.secondary_work = self.secondary_work.drop(columns=[col for col in cols_to_drop if col in self.secondary_work.columns])
            
            # Log statistics
            total_records = len(self.secondary_work)
            self.log_message(f"\nAdvanced processing complete!")
            self.log_message(f"Total records processed: {total_records}")
            self.log_message(f"Exact matches: {exact_matches} ({exact_matches/total_records*100:.1f}%)" if total_records > 0 else "No records to process")
            self.log_message(f"Fuzzy matches: {fuzzy_matches} ({fuzzy_matches/total_records*100:.1f}%)" if total_records > 0 else "")
            self.log_message(f"No matches (Review): {no_matches} ({no_matches/total_records*100:.1f}%)" if total_records > 0 else "")
            
            # Show preview
            self.show_advanced_results_preview()
            
            # After processing, update the replace column UI
            self.update_replace_col_combo()
            
        except Exception as e:
            error_msg = f"Error processing data: {str(e)}\n{traceback.format_exc()}"
            self.log_message(error_msg, "ERROR")
            messagebox.showerror("Processing Error", str(e))

    def show_advanced_results_preview(self):
        """Show preview of advanced results in log"""
        if not self.selected_target_columns:
            return
        
        primary_target_col = self.selected_target_columns[0]
        self.log_message("\n--- ADVANCED RESULTS PREVIEW ---")
        
        if self.preserve_structure.get() and hasattr(self, 'secondary_work'):
            # Preview from updated secondary dataframe
            preview_df = self.secondary_work[self.secondary_work[primary_target_col].notna() & 
                                           (self.secondary_work[primary_target_col].astype(str).str.strip() != "")].head(10)
            
            for idx, row in preview_df.iterrows():
                # Show all target columns for this row
                target_info = []
                for i, target_col in enumerate(self.selected_target_columns):
                    if target_col in row and pd.notna(row[target_col]):
                        target_info.append(f"{target_col}: {row[target_col]}")
                
                match_type = row.get('Match_Type', '')
                matched_name = row.get('Matched_Reference_Name', '')
                matched_col = row.get('Matched_Column', '')
                
                # Show individual replaced values based on position mapping
                replaced_values = []
                for i, replace_col in enumerate(self.selected_replace_columns):
                    if replace_col in row and i < len(self.selected_target_columns):
                        target_name = self.selected_target_columns[i] if i < len(self.selected_target_columns) else "N/A"
                        replaced_values.append(f"{target_name[:10]}→{replace_col}: {row[replace_col]}")
                
                target_str = " | ".join(target_info[:2])  # Show first 2 target columns
                replaced_str = " | ".join(replaced_values[:2])  # Show first 2 mappings
                
                self.log_message(f"{idx+1}. Targets: {target_str}")
                self.log_message(f"    Best Match: {match_type} via {matched_col}: {matched_name}")
                if replaced_str:
                    self.log_message(f"    Individual Updates: {replaced_str}")
            
            non_empty_count = len(self.secondary_work[self.secondary_work[primary_target_col].notna() & 
                                                    (self.secondary_work[primary_target_col].astype(str).str.strip() != "")])
            if non_empty_count > 10:
                self.log_message(f"... and {non_empty_count - 10} more records")

    def clean_name(self, name):
        """Clean and normalize name for comparison"""
        if pd.isna(name) or name is None:
            return ""
        
        name = str(name).strip()
        # Remove extra spaces and convert to lowercase
        name = re.sub(r'\s+', ' ', name).lower()
        # Remove special characters except spaces and hyphens
        name = re.sub(r'[^\w\s\-]', '', name)
        return name

    def export_results(self):
        """Export results to Excel file"""
        try:
            if (self.preserve_structure.get() and not hasattr(self, 'secondary_work')) or \
               (not self.preserve_structure.get() and not hasattr(self, 'result_df')):
                messagebox.showwarning("No Results", "No results to export. Please process data first.")
                return
                
            filename = filedialog.asksaveasfilename(
                title="Save Results",
                defaultextension=".xlsx",
                filetypes=[
                    ("Excel files", "*.xlsx"),
                    ("CSV files", "*.csv"),
                    ("All files", "*.*")
                ]
            )
            
            if filename:
                file_ext = Path(filename).suffix.lower()
                
                if self.preserve_structure.get():
                    # Export updated secondary file with match tracking columns
                    export_df = self.secondary_work.copy()
                    
                    # Remove only the clean_name column but keep match tracking columns
                    if 'clean_name' in export_df.columns:
                        export_df = export_df.drop(columns=['clean_name'])
                    
                    if file_ext == '.xlsx':
                        # Use openpyxl to preserve formatting better
                        with pd.ExcelWriter(filename, engine='openpyxl') as writer:
                            export_df.to_excel(writer, sheet_name='Results', index=False)
                            
                            # Get the workbook and worksheet
                            workbook = writer.book
                            worksheet = writer.sheets['Results']
                            
                            # Auto-adjust column widths
                            for column in worksheet.columns:
                                max_length = 0
                                column_letter = column[0].column_letter
                                for cell in column:
                                    try:
                                        if len(str(cell.value)) > max_length:
                                            max_length = len(str(cell.value))
                                    except:
                                        pass
                                adjusted_width = min(max_length + 2, 50)
                                worksheet.column_dimensions[column_letter].width = adjusted_width
                        
                    elif file_ext == '.csv':
                        export_df.to_csv(filename, index=False)
                    else:
                        # Default to Excel
                        export_df.to_excel(filename, index=False)
                else:
                    # Original export behavior with multiple sheets
                    if file_ext == '.xlsx':
                        # Create Excel file with multiple sheets
                        with pd.ExcelWriter(filename, engine='openpyxl') as writer:
                            # Main results
                            self.result_df.to_excel(writer, sheet_name='Results', index=False)
                            
                            # Summary statistics
                            summary_data = {
                                'Metric': ['Total Records', 'Exact Matches', 'Fuzzy Matches', 'No Matches (Review)'],
                                'Count': [
                                    len(self.result_df),
                                    len(self.result_df[self.result_df['Match_Type'] == 'EXACT']),
                                    len(self.result_df[self.result_df['Match_Type'] == 'FUZZY']),
                                    len(self.result_df[self.result_df['Match_Type'] == 'REVIEW'])
                                ]
                            }
                            summary_df = pd.DataFrame(summary_data)
                            summary_df.to_excel(writer, sheet_name='Summary', index=False)
                            
                            # Records needing review
                            review_df = self.result_df[self.result_df['Match_Type'] == 'REVIEW']
                            if not review_df.empty:
                                review_df.to_excel(writer, sheet_name='Need_Review', index=False)
                                
                    elif file_ext == '.csv':
                        self.result_df.to_csv(filename, index=False)
                    else:
                        # Default to Excel
                        self.result_df.to_excel(filename, index=False)
                
                self.log_message(f"Results exported successfully to: {filename}")
                messagebox.showinfo("Export Complete", f"Results exported to:\n{filename}")
                
        except Exception as e:
            error_msg = f"Error exporting results: {str(e)}"
            self.log_message(error_msg, "ERROR")
            messagebox.showerror("Export Error", error_msg)

    def find_match_for_value(self, target_value, master_work, reference_primary, data_source, exact_only=False, fuzzy_only=False):
        """Find match for a single target value"""
        target_clean_name = self.clean_name(target_value)
        if not target_clean_name:
            return None
        
        # Try exact match
        if not fuzzy_only:
            exact_match = master_work[master_work['clean_reference'] == target_clean_name]
            if not exact_match.empty:
                return {
                    'type': 'EXACT',
                    'matched_name': exact_match.iloc[0][reference_primary],
                    'confidence': 1.0,
                    'data_value': exact_match.iloc[0].get(data_source, "")
                }
        
        # Try fuzzy matching if enabled and no exact match
        if not exact_only and self.fuzzy_matching.get():
            master_names = master_work['clean_reference'].dropna().tolist()
            if master_names:
                close_matches = get_close_matches(
                    target_clean_name,
                    master_names,
                    n=1,
                    cutoff=self.similarity_threshold.get()
                )
                
                if close_matches:
                    matched_clean_name = close_matches[0]
                    fuzzy_match = master_work[master_work['clean_reference'] == matched_clean_name]
                    
                    if not fuzzy_match.empty:
                        from difflib import SequenceMatcher
                        confidence = SequenceMatcher(None, target_clean_name, matched_clean_name).ratio()
                        
                        return {
                            'type': 'FUZZY',
                            'matched_name': fuzzy_match.iloc[0][reference_primary],
                            'confidence': confidence,
                            'data_value': fuzzy_match.iloc[0].get(data_source, "")
                        }
        
        return None

    def clear_logs(self):
        """Clear all log text widgets"""
        if hasattr(self, 'log_text') and self.log_text:
            self.log_text.delete(1.0, tk.END)
        if hasattr(self, 'log_text_mapping') and self.log_text_mapping:
            self.log_text_mapping.delete(1.0, tk.END)

    def bypass_dependencies(self):
        # Enable all main tabs regardless of dependency status
        self.notebook.tab(1, state="normal")
        self.notebook.tab(2, state="normal")
        # Optionally, show a warning message
        messagebox.showwarning("Bypass Activated", "Dependency check bypassed. Please ensure all required packages are installed for full functionality.")

def main():
    root = tk.Tk()
    # Set theme
    style = ttk.Style()
    style.theme_use('clam')
    # Create accent button style
    style.configure("Accent.TButton", font=("Arial", 10, "bold"))
    # Make sure buttons are always visible with a minimum width
    style.configure("TButton", padding=5)
    app = ExcelComparisonTool(root)
        # Copyright disclaimer
    app.log_message("© 2025 Excella - Excel Comparison Tool. Author: frenzywall. All rights reserved.", "INFO")
    
    # Display system info for debugging
    is_frozen = getattr(sys, 'frozen', False)
    
    # When running as executable, force pandas version display
    if is_frozen:
        pandas_version = pd.__version__
        app.log_message("Running as packaged executable", "INFO")
    elif HAS_PANDAS:
        try:
            pandas_version = pd.__version__
        except:
            pandas_version = "Unknown version"
    else:
        pandas_version = "Not installed"
        
    platform_info = f"Python {sys.version_info.major}.{sys.version_info.minor}.{sys.version_info.micro}, " \
                    f"Platform: {platform.system()} {platform.version()}, " \
                    f"Pandas: {pandas_version}"
    app.log_message(f"System info: {platform_info}", "INFO")
    
    if is_frozen or HAS_WIN32COM:
        app.log_message("COM interface support is available for Enterprise Excel files", "INFO")
    else:
        app.log_message("COM interface not available. Install pywin32 for better Enterprise Excel support", "INFO")
    
    root.mainloop()

if __name__ == "__main__":
    main()

    
