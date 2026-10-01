import win32com.client as win32
import os
import sys
import random
import datetime
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
import pandas as pd
from pathlib import Path
import customtkinter as ctk
from customtkinter import filedialog
import json
import warnings

try:
    import mplcursors
except ImportError:
    mplcursors = None

try:
    from fpdf import FPDF
except ImportError:
    FPDF = None

from deap import base, creator, tools, algorithms

# ==========================================
# 1. STREAM REDIRECTOR (IN-APP TERMINAL)
# ==========================================
class ConsoleRedirector:
    def __init__(self, textbox):
        self.textbox = textbox

    def write(self, text):
        self.textbox.configure(state="normal")
        self.textbox.insert("end", text)
        self.textbox.see("end")
        self.textbox.configure(state="disabled")

    def flush(self):
        pass


# ==========================================
# 2. THE BACKEND ENGINE (AI & ASPEN)
# ==========================================
simulation_cache = {}

def connect_to_aspen(aspen_filepath):
    try:
        aspen = win32.Dispatch('Apwn.Document')
        aspen.InitFromFile2(os.path.abspath(aspen_filepath))
        return aspen
    except Exception:
        return None

def scan_aspen_inventory(aspen):
    blocks, streams = [], []
    try:
        b_node = aspen.Tree.FindNode(r"\Data\Blocks")
        if b_node:
            for b in b_node.Elements: blocks.append(b.Name)
            
        s_node = aspen.Tree.FindNode(r"\Data\Streams")
        if s_node:
            for s in s_node.Elements: streams.append(s.Name)
    except Exception:
        pass
    return blocks, streams

def evaluate_flowsheet(individual, config, aspen):
    ind_key = tuple(round(x, 4) for x in individual)
    if ind_key in simulation_cache:
        return simulation_cache[ind_key]
        
    try:
        for i, input_var in enumerate(config["inputs"]):
            node = aspen.Tree.FindNode(input_var["path"])
            if node: node.Value = individual[i]
                
        aspen.Engine.Run2()
        
        error_node = aspen.Tree.FindNode(r"\Data\Results Summary\Run-Status\NERROR")
        if error_node and error_node.Value is not None and int(error_node.Value) > 0:
            failure_result = tuple([-999999.0] * len(config["objectives"]))
            simulation_cache[ind_key] = failure_result
            return failure_result
        
        for const in config["constraints"]:
            try:
                node = aspen.Tree.FindNode(const["path"])
                if node and node.Value is not None:
                    val = float(node.Value)
                    if abs(val) > 1e20: 
                        return tuple([-999999.0] * len(config["objectives"]))
                    if const["type"] == "MAX" and val > const["limit"]:
                        return tuple([-999999.0] * len(config["objectives"]))
                    if const["type"] == "MIN" and val < const["limit"]:
                        return tuple([-999999.0] * len(config["objectives"]))
                else:
                    return tuple([-999999.0] * len(config["objectives"]))
            except Exception:
                return tuple([-999999.0] * len(config["objectives"]))
        
        results = []
        for obj in config["objectives"]:
            try:
                node = aspen.Tree.FindNode(obj["path"])
                if node and node.Value is not None:
                    val = float(node.Value)
                    if abs(val) > 1e20: 
                        results.append(-999999.0)
                    else:
                        results.append(val * obj["weight"])
                else:
                    results.append(-999999.0)
            except Exception:
                results.append(-999999.0)
                
        final_result = tuple(results)
        simulation_cache[ind_key] = final_result
        return final_result
        
    except Exception:
        return tuple([-999999.0] * len(config["objectives"]))

def run_optimization(config, aspen, pop_size, generations, progress_callback=None, abort_check=None):
    num_objectives = len(config["objectives"])
    
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        creator.create("FitnessMulti", base.Fitness, weights=tuple([1.0] * num_objectives))
        creator.create("Individual", list, fitness=creator.FitnessMulti)

    toolbox = base.Toolbox()
    toolbox.register("attr_float", lambda bounds: random.uniform(bounds[0], bounds[1]))
    toolbox.register("individual", lambda: creator.Individual([toolbox.attr_float(var["bounds"]) for var in config["inputs"]]))
    toolbox.register("population", tools.initRepeat, list, toolbox.individual)
    toolbox.register("evaluate", evaluate_flowsheet, config=config, aspen=aspen)
    toolbox.register("mate", tools.cxSimulatedBinaryBounded, low=[v["bounds"][0] for v in config["inputs"]], up=[v["bounds"][1] for v in config["inputs"]], eta=20.0)
    toolbox.register("mutate", tools.mutPolynomialBounded, low=[v["bounds"][0] for v in config["inputs"]], up=[v["bounds"][1] for v in config["inputs"]], eta=20.0, indpb=1.0/len(config["inputs"]))
    toolbox.register("select", tools.selNSGA2)

    pop = toolbox.population(n=pop_size)
    
    print(f"[ENGINE] Generating initial population of {pop_size} plant models...")
    
    for i, ind in enumerate(pop):
        if abort_check and abort_check(): break
        ind.fitness.values = toolbox.evaluate(ind)
        if progress_callback: progress_callback(0, generations, i+1, pop_size)

    pop = [ind for ind in pop if ind.fitness.valid]

    for gen in range(generations):
        if abort_check and abort_check(): 
            print("[ENGINE] Emergency Stop Confirmed. Salvaging computed data...")
            break
            
        gen_num = gen + 1
        print(f"[EVOLUTION] Processing Generation {gen_num}/{generations}...")
        
        offspring = algorithms.varAnd(pop, toolbox, cxpb=0.7, mutpb=0.2)
        
        fits = []
        for i, ind in enumerate(offspring):
            if abort_check and abort_check(): break
            fit = toolbox.evaluate(ind)
            fits.append(fit)
            if progress_callback: progress_callback(gen_num, generations, i+1, len(offspring))
            
        for ind, fit in zip(offspring[:len(fits)], fits):
            ind.fitness.values = fit
            
        evaluated_offspring = [ind for ind in offspring if ind.fitness.valid]
        pop = toolbox.select(pop + evaluated_offspring, k=min(pop_size, len(pop + evaluated_offspring)))

    if not pop:
        return [], config

    pareto_front = tools.sortNondominated(pop, len(pop), first_front_only=True)[0]
    
    csv_data = []
    for ind in pareto_front:
        row = {}
        for j, var in enumerate(config["inputs"]):
            row[var['name']] = round(ind[j], 4)
        for k, obj in enumerate(config["objectives"]):
            row[obj['name']] = round(ind.fitness.values[k] / obj["weight"], 4)
        csv_data.append(row)
        
    df = pd.DataFrame(csv_data)
    
    desktop_path = Path.home() / "Desktop"
    if not desktop_path.exists():
        desktop_path = Path.home() / "OneDrive" / "Desktop"
        
    export_file = desktop_path / "Aspen_Pareto_Results.csv"
    df.to_csv(export_file, index=False)
    print(f"[FILE IO] Optimal dataset successfully written to: {export_file}")

    return pareto_front, config


# ==========================================
# 3. GRAPHICAL USER INTERFACE
# ==========================================
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

COMMON_INPUTS = ["TEMP", "PRES", "VFRAC", "HEAT-DUTY", "MASS-FLOW", "MOLE-FLOW", "MOLE-RR", "MASS-RR", "DIST-RATE", "BTM-RATE", "NTRAY", "FEED-STAGE"]
COMMON_OUTPUTS = ["RES_MOLEFLOW", "RES_MASSFLOW", "QCALC", "TEMP_OUT", "PRES_OUT", "DUTY", "MOLEFRAC", "MASSFRAC", "REB-DUTY", "COND-DUTY", "PURITY"]

class AspenOptimizerUI(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Methanol Research Optimizer — AI Edition")
        self.geometry("1100x820")
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self.aspen_link = None
        self.available_blocks = ["(Load File First)"]
        self.available_streams = ["(Load File First)"]
        
        self.var_rows = []
        self.obj_rows = []
        self.const_rows = []
        
        self.abort_flag = False
        self.final_pareto_data = None

        # --- SIDEBAR ---
        self.sidebar = ctk.CTkFrame(self, width=230, corner_radius=0)
        self.sidebar.grid(row=0, column=0, sticky="nsew")
        self.sidebar.grid_rowconfigure(5, weight=1) 

        self.logo_label = ctk.CTkLabel(self.sidebar, text="⚡ METHANOL AI\nRESEARCH", font=ctk.CTkFont(size=20, weight="bold"))
        self.logo_label.grid(row=0, column=0, padx=20, pady=(30, 5))
        self.version_label = ctk.CTkLabel(self.sidebar, text="Research Edition v1.0", text_color="gray", font=ctk.CTkFont(size=12))
        self.version_label.grid(row=1, column=0, padx=20, pady=(0, 15))

        # --- LIVE TELEMETRY IN SIDEBAR ---
        self.telem_frame = ctk.CTkFrame(self.sidebar, fg_color="gray15", corner_radius=8)
        self.telem_frame.grid(row=2, column=0, padx=15, pady=10, sticky="ew")
        
        ctk.CTkLabel(self.telem_frame, text="SYSTEM STATUS", font=ctk.CTkFont(size=11, weight="bold"), text_color="#3498db").pack(anchor="w", padx=10, pady=(8, 2))
        self.telem_status = ctk.CTkLabel(self.telem_frame, text="Idle / Standby", font=ctk.CTkFont(size=12))
        self.telem_status.pack(anchor="w", padx=10, pady=(0, 5))

        self.progress_bar = ctk.CTkProgressBar(self.telem_frame, orientation="horizontal", height=10)
        self.progress_bar.pack(fill="x", padx=10, pady=(5, 5))
        self.progress_bar.set(0.0)

        self.telem_pct = ctk.CTkLabel(self.telem_frame, text="Cycle Progress: 0%", font=ctk.CTkFont(size=11), text_color="gray")
        self.telem_pct.pack(anchor="w", padx=10, pady=(0, 8))

        # --- MAIN TABVIEW ---
        self.tabs = ctk.CTkTabview(self)
        self.tabs.grid(row=0, column=1, sticky="nsew", padx=20, pady=(10, 0))

        self.tab_engine = self.tabs.add("1. Engine & File")
        self.tab_vars = self.tabs.add("2. Variables")
        self.tab_objs = self.tabs.add("3. Objectives")
        self.tab_consts = self.tabs.add("4. Constraints")
        self.tab_results = self.tabs.add("5. Results Dashboard")

        self.setup_engine_tab()
        self.setup_vars_tab()
        self.setup_objs_tab()
        self.setup_consts_tab()
        self.setup_results_tab()

        # --- IN-APP COMMAND TERMINAL ---
        self.bottom_frame = ctk.CTkFrame(self, height=180, corner_radius=0, fg_color="gray10")
        self.bottom_frame.grid(row=1, column=0, columnspan=2, sticky="nsew")
        self.bottom_frame.grid_columnconfigure(0, weight=1)

        term_header = ctk.CTkFrame(self.bottom_frame, fg_color="transparent", height=25)
        term_header.pack(fill="x", padx=15, pady=(5, 0))
        ctk.CTkLabel(term_header, text="COMMAND CENTER CONSOLE OUTPUT", font=ctk.CTkFont(size=11, weight="bold"), text_color="#1d865c").pack(side="left")

        self.console_box = ctk.CTkTextbox(self.bottom_frame, height=140, font=ctk.CTkFont(family="Consolas", size=11), state="disabled", fg_color="black", text_color="#00ff66")
        self.console_box.pack(fill="both", expand=True, padx=15, pady=(2, 10))

        sys.stdout = ConsoleRedirector(self.console_box)
        print("[CONSOLE] Methanol Research AI initialized successfully.")

    def delete_row(self, row_data, row_list):
        row_data["frame"].destroy()
        row_list.remove(row_data)

    def setup_engine_tab(self):
        self.file_label = ctk.CTkLabel(self.tab_engine, text="Aspen Project File:", font=ctk.CTkFont(weight="bold"))
        self.file_label.pack(anchor="w", padx=20, pady=(15, 5))
        
        frame1 = ctk.CTkFrame(self.tab_engine, fg_color="transparent")
        frame1.pack(fill="x", padx=20)
        
        self.file_entry = ctk.CTkEntry(frame1, placeholder_text="Select your .apwz file...")
        self.file_entry.pack(side="left", fill="x", expand=True, padx=(0, 10))
        
        ctk.CTkButton(frame1, text="Browse", width=80, command=self.browse_file).pack(side="left", padx=(0, 10))
        ctk.CTkButton(frame1, text="Connect & Scan", width=130, fg_color="#b87333", hover_color="#8c5827", command=self.scan_file).pack(side="left")

        preset_frame = ctk.CTkFrame(self.tab_engine, fg_color="gray20", corner_radius=8)
        preset_frame.pack(fill="x", padx=20, pady=15)
        
        ctk.CTkLabel(preset_frame, text="Session Memory:", font=ctk.CTkFont(weight="bold")).pack(side="left", padx=15, pady=12)
        ctk.CTkButton(preset_frame, text="Save Config (.json)", fg_color="#27ae60", hover_color="#1e8449", command=self.save_config).pack(side="left", padx=(0, 10))
        ctk.CTkButton(preset_frame, text="Load Config (.json)", fg_color="#2980b9", hover_color="#1f618d", command=self.load_config).pack(side="left")
        
        self.ai_label = ctk.CTkLabel(self.tab_engine, text="Evolutionary Search Hyperparameters:", font=ctk.CTkFont(weight="bold"))
        self.ai_label.pack(anchor="w", padx=20, pady=(15, 5))

        frame2 = ctk.CTkFrame(self.tab_engine, fg_color="transparent")
        frame2.pack(fill="x", padx=20)
        
        ctk.CTkLabel(frame2, text="Population Size:").pack(side="left")
        self.pop_entry = ctk.CTkEntry(frame2, width=60)
        self.pop_entry.insert(0, "40")
        self.pop_entry.pack(side="left", padx=(10, 30))

        ctk.CTkLabel(frame2, text="Generations:").pack(side="left")
        self.gen_entry = ctk.CTkEntry(frame2, width=60)
        self.gen_entry.insert(0, "20")
        self.gen_entry.pack(side="left", padx=(10, 0))

        self.run_btn = ctk.CTkButton(
            self.tab_engine, text="LAUNCH INDUSTRIAL OPTIMIZATION", font=ctk.CTkFont(size=16, weight="bold"), 
            fg_color="#1d865c", hover_color="#145c3f", height=45, command=self.toggle_engine_state
        )
        self.run_btn.pack(fill="x", padx=20, pady=25)

    def setup_results_tab(self):
        self.results_header = ctk.CTkFrame(self.tab_results, fg_color="transparent", height=40)
        self.results_header.pack(fill="x", padx=10, pady=5)
        
        self.pdf_btn = ctk.CTkButton(self.results_header, text="📄 Export Executive PDF Report", fg_color="#8e44ad", hover_color="#732d91", state="disabled", command=self.generate_pdf)
        self.pdf_btn.pack(side="right")

        self.results_container = ctk.CTkFrame(self.tab_results, fg_color="transparent")
        self.results_container.pack(fill="both", expand=True, padx=10, pady=0)
        
        self.no_data_label = ctk.CTkLabel(
            self.results_container, 
            text="Execution Pending.\nRun optimization to generate your interactive Pareto frontier.", 
            font=ctk.CTkFont(size=14, slant="italic"), text_color="gray"
        )
        self.no_data_label.pack(expand=True)

    def save_config(self):
        data = {"engine": {"pop": self.pop_entry.get(), "gen": self.gen_entry.get()}, "vars": [], "objs": [], "consts": []}
        for row in self.var_rows: data["vars"].append({"block": row['block'].get(), "prop": row['prop'].get(), "min": row['min'].get(), "max": row['max'].get()})
        for row in self.const_rows: data["consts"].append({"target": row['target'].get(), "type": row['type'].get(), "prop": row['prop'].get(), "limit": row['limit'].get(), "val": row['val'].get()})
        filepath = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("JSON Files", "*.json")])
        if filepath:
            with open(filepath, 'w') as f: json.dump(data, f, indent=4)
            print(f"[SYSTEM] State configuration serialized to: {filepath}")

    def load_config(self):
        filepath = filedialog.askopenfilename(filetypes=[("JSON Files", "*.json")])
        if filepath:
            with open(filepath, 'r') as f: data = json.load(f)
            self.pop_entry.delete(0, 'end'); self.pop_entry.insert(0, data["engine"]["pop"])
            self.gen_entry.delete(0, 'end'); self.gen_entry.insert(0, data["engine"]["gen"])
            for row in list(self.var_rows): self.delete_row(row, self.var_rows)
            for row in list(self.const_rows): self.delete_row(row, self.const_rows)
            for v in data.get("vars", []): self.add_var_row(v["block"], v["prop"], v["min"], v["max"])
            for c in data.get("consts", []): self.add_const_row(c["target"], c["type"], c["prop"], c["limit"], c["val"])
            print(f"[SYSTEM] Configuration successfully loaded from: {filepath}")

    def browse_file(self):
        filename = filedialog.askopenfilename(filetypes=[("Aspen Plus Workspaces", "*.apwz")])
        if filename:
            self.file_entry.delete(0, 'end')
            self.file_entry.insert(0, filename)
            print(f"[ASPEN] Selected workspace: {filename}")

    def scan_file(self):
        filepath = self.file_entry.get()
        if not filepath or not os.path.exists(filepath): return
        print("[ASPEN] Initializing COM link and cataloging plant topology...")
        self.telem_status.configure(text="Scanning Topology...", text_color="#f39c12")
        self.update()
        if not self.aspen_link: self.aspen_link = connect_to_aspen(filepath)
        if self.aspen_link:
            blocks, streams = scan_aspen_inventory(self.aspen_link)
            self.available_blocks = blocks if blocks else ["None Found"]
            self.available_streams = streams if streams else ["None Found"]
            self.refresh_dropdowns()
            self.telem_status.configure(text="Connected / Ready", text_color="#1d865c")
            print(f"[ASPEN] Scan successful: Identified {len(blocks)} Blocks and {len(streams)} Streams.")

    def refresh_dropdowns(self):
        all_inventory = self.available_blocks + self.available_streams
        for row in self.var_rows: row['block'].configure(values=self.available_blocks)
        for row in self.const_rows: row['target'].configure(values=all_inventory)

    def setup_vars_tab(self):
        ctk.CTkButton(self.tab_vars, text="+ Add Variable", command=lambda: self.add_var_row("", "TEMP", "", "")).pack(anchor="w", padx=20, pady=10)
        self.vars_scroll = ctk.CTkScrollableFrame(self.tab_vars, fg_color="transparent")
        self.vars_scroll.pack(fill="both", expand=True, padx=10, pady=5)
        self.add_var_row("", "TEMP", "", "")

    def add_var_row(self, block_val, prop_val, min_val, max_val):
        row = ctk.CTkFrame(self.vars_scroll)
        row.pack(fill="x", pady=2)
        e_block = ctk.CTkComboBox(row, width=130, values=self.available_blocks)
        e_block.set(block_val if block_val else (self.available_blocks[0] if self.available_blocks else ""))
        e_block.pack(side="left", padx=5)
        e_prop = ctk.CTkComboBox(row, width=170, values=COMMON_INPUTS)
        e_prop.set(prop_val); e_prop.pack(side="left", padx=5)
        e_min = ctk.CTkEntry(row, width=100, placeholder_text="Min"); e_min.insert(0, min_val); e_min.pack(side="left", padx=5)
        e_max = ctk.CTkEntry(row, width=100, placeholder_text="Max"); e_max.insert(0, max_val); e_max.pack(side="left", padx=5)
        row_data = {"frame": row, "block": e_block, "prop": e_prop, "min": e_min, "max": e_max}
        self.var_rows.append(row_data)
        ctk.CTkButton(row, text="X", width=30, fg_color="#c0392b", hover_color="#922b21", command=lambda: self.delete_row(row_data, self.var_rows)).pack(side="left", padx=5)

    def setup_objs_tab(self):
        ctk.CTkLabel(self.tab_objs, text="RESEARCH PAPER OVERRIDE ACTIVE", font=ctk.CTkFont(size=14, weight="bold"), text_color="#f1c40f").pack(pady=(20, 5))
        ctk.CTkLabel(self.tab_objs, text="This edition is hardcoded for your Methanol Optimization paper.\nThe AI will automatically map:\n\n1. Minimize Stream 2 Molar Flow (Maximizes Methanol Yield)\n2. Minimize Stream 2 Temperature (Maps Energy Cost)", justify="left").pack(pady=10)

    def setup_consts_tab(self):
        ctk.CTkButton(self.tab_consts, text="+ Add Constraint", command=lambda: self.add_const_row("", "Stream", "RES_MASSFLOW", "MIN", "")).pack(anchor="w", padx=20, pady=10)
        self.consts_scroll = ctk.CTkScrollableFrame(self.tab_consts, fg_color="transparent")
        self.consts_scroll.pack(fill="both", expand=True, padx=10, pady=5)

    def add_const_row(self, target_val, type_val, prop_val, limit_val, val_val):
        row = ctk.CTkFrame(self.consts_scroll)
        row.pack(fill="x", pady=2)
        all_inventory = self.available_blocks + self.available_streams
        e_target = ctk.CTkComboBox(row, width=140, values=all_inventory)
        e_target.set(target_val if target_val else (all_inventory[0] if all_inventory else ""))
        e_target.pack(side="left", padx=5)
        dd_type = ctk.CTkOptionMenu(row, width=110, values=["Stream", "Block"]); dd_type.set(type_val); dd_type.pack(side="left", padx=5)
        e_prop = ctk.CTkComboBox(row, width=160, values=COMMON_OUTPUTS); e_prop.set(prop_val); e_prop.pack(side="left", padx=5)
        dd_limit = ctk.CTkOptionMenu(row, width=80, values=["MIN", "MAX"]); dd_limit.set(limit_val); dd_limit.pack(side="left", padx=5)
        e_val = ctk.CTkEntry(row, width=80); e_val.insert(0, val_val); e_val.pack(side="left", padx=5)
        row_data = {"frame": row, "target": e_target, "type": dd_type, "prop": e_prop, "limit": dd_limit, "val": e_val}
        self.const_rows.append(row_data)
        ctk.CTkButton(row, text="X", width=30, fg_color="#c0392b", hover_color="#922b21", command=lambda: self.delete_row(row_data, self.const_rows)).pack(side="left", padx=5)

    def check_abort(self):
        self.update() 
        return self.abort_flag

    def toggle_engine_state(self):
        if self.run_btn.cget("text") == "LAUNCH INDUSTRIAL OPTIMIZATION":
            self.abort_flag = False
            self.run_btn.configure(text="🛑 ABORT OPERATION (SAFE STOP)", fg_color="#c0392b", hover_color="#922b21")
            self.launch_engine()
        else:
            print("[SYSTEM] User initiated Emergency Stop. Instructing engine to break loops...")
            self.telem_status.configure(text="Halting Engine...", text_color="#e74c3c")
            self.abort_flag = True

    def reset_run_btn(self):
        self.run_btn.configure(text="LAUNCH INDUSTRIAL OPTIMIZATION", fg_color="#1d865c", hover_color="#145c3f")

    def launch_engine(self):
        filepath = self.file_entry.get()
        if not filepath: 
            print("[ERROR] Flowsheet not loaded.")
            self.reset_run_btn()
            return
            
        config = {"inputs": [], "objectives": [], "constraints": []}
        
        try:
            pop_size, generations = int(self.pop_entry.get()), int(self.gen_entry.get())
            
            for row in self.var_rows:
                b, p, m, mx = row['block'].get().strip(), row['prop'].get().strip(), row['min'].get().strip(), row['max'].get().strip()
                if not b or not p or not m or not mx or "Load File First" in b: continue
                vmin, vmax = float(m), float(mx)
                if vmin > vmax: vmin, vmax = vmax, vmin
                path_str = rf"\Data\Blocks\{b}\Input\{p}"
                config["inputs"].append({"name": f"{b}_{p}", "path": path_str, "bounds": (vmin, vmax)})
                

            # =========================================================================
            # THE RESEARCH PAPER OVERRIDE (YIELD VS SELECTIVITY)
            # =========================================================================
            config["objectives"] = [
                {
                    "name": "Stream2_TotalMoles (Yield)", 
                    "path": r"\Data\Streams\2\Output\RES_MOLEFLOW", 
                    "weight": -1.0  # Minimizing total moles maximizes conversion/yield
                },
                {
                    "name": "Reactor_Temp (Selectivity Control)", 
                    "path": r"\Data\Blocks\B1\Input\TEMP", 
                    "weight": -1.0  # Minimizing temp suppresses RWGS to maximize selectivity
                }
            ]
            # =========================================================================

            for row in self.const_rows:
                t, p, v = row['target'].get().strip(), row['prop'].get().strip(), row['val'].get().strip()
                if not t or not p or not v or "Load File First" in t: continue
                folder = "Streams" if row['type'].get() == "Stream" else "Blocks"
                if t in self.available_blocks: folder = "Blocks"
                if t in self.available_streams: folder = "Streams"
                path_str = rf"\Data\{folder}\{t}\Output\{p}"
                config["constraints"].append({"name": f"{t}_Const", "path": path_str, "type": row['limit'].get(), "limit": float(v)})
                
            if len(config["inputs"]) < 1:
                print("[ERROR] Requirement unmet: Need >=1 Decision Variable.")
                self.reset_run_btn()
                return
        except ValueError:
            print("[ERROR] Type parsing failure: Check bounds/limits format.")
            self.reset_run_btn()
            return

        print("[ENGINE] Preparing Aspen link for evolutionary run...")
        self.telem_status.configure(text="Simulating Models...", text_color="#2ecc71")
        self.progress_bar.set(0.0)
        self.update()

        try:
            if not self.aspen_link: self.aspen_link = connect_to_aspen(filepath)
            if not self.aspen_link:
                print("[ERROR] Connection to Aspen COM service dropped.")
                self.telem_status.configure(text="Connection Fault", text_color="red")
                return

            pareto_front, config = run_optimization(config, self.aspen_link, pop_size, generations, progress_callback=self.update_telemetry, abort_check=self.check_abort)
            
            if not pareto_front:
                self.telem_status.configure(text="Aborted (No Data)", text_color="#e74c3c")
                print("[SYSTEM] Operation halted before any valid plant designs were calculated.")
                return

            self.final_pareto_data = (pareto_front, config)
            self.pdf_btn.configure(state="normal")

            if self.abort_flag:
                self.telem_status.configure(text="Aborted (Data Salvaged)", text_color="#e74c3c")
                print("[SUCCESS] Operation aborted safely. Salvaged data rendered to dashboard.")
            else:
                self.telem_status.configure(text="Run Complete", text_color="#1d865c")
                self.progress_bar.set(1.0)
                print("[SUCCESS] Evolutionary cycle concluded. Rendering embedded Pareto front...")

            self.render_embedded_pareto(config, pareto_front)
            
        except Exception as e:
            print(f"[FATAL ERROR] Engine execution failed: {str(e)}")
            self.telem_status.configure(text="System Crash", text_color="red")
            
        finally:
            self.reset_run_btn()

    def update_telemetry(self, gen_curr, gen_total, ind_curr=0, ind_total=0):
        if gen_curr == 0:
            pct = (ind_curr / ind_total) * (1 / (gen_total + 1))
            self.telem_pct.configure(text=f"Init Population: {ind_curr}/{ind_total}")
        else:
            base_pct = gen_curr / (gen_total + 1)
            sub_pct = (ind_curr / ind_total) * (1 / (gen_total + 1))
            pct = base_pct + sub_pct
            self.telem_pct.configure(text=f"Cycle Progress: {int(pct * 100)}% (Gen {gen_curr}/{gen_total})")
        
        self.progress_bar.set(pct)
        self.update()

    def render_embedded_pareto(self, config, pareto_front):
        for widget in self.results_container.winfo_children(): widget.destroy()
        obj1_name, obj2_name = config["objectives"][0]["name"], config["objectives"][1]["name"]
        obj1_vals = [ind.fitness.values[0] / config["objectives"][0]["weight"] for ind in pareto_front]
        obj2_vals = [ind.fitness.values[1] / config["objectives"][1]["weight"] for ind in pareto_front]

        plt.style.use("dark_background")
        fig, ax = plt.subplots(figsize=(7, 4.5), dpi=100)
        fig.patch.set_facecolor("#1e1e1e"); ax.set_facecolor("#2b2b2b")
        scatter = ax.scatter(obj1_vals, obj2_vals, color="#00ffcc", edgecolor="white", s=65, label="Non-Dominated Optimal", zorder=3)
        ax.set_title("Optimal Temp vs. Methanol Yield Pareto Surface", fontsize=12, fontweight="bold", color="white", pad=12)
        ax.set_xlabel(f"{obj1_name}", fontsize=10, color="lightgray")
        ax.set_ylabel(f"{obj2_name}", fontsize=10, color="lightgray")
        ax.grid(True, linestyle="--", alpha=0.3, zorder=0)
        ax.legend(facecolor="#1e1e1e", edgecolor="gray")
        fig.tight_layout()

        if mplcursors:
            cursor = mplcursors.cursor(scatter, hover=True)
            @cursor.connect("add")
            def on_add(sel):
                ind = pareto_front[sel.index]
                text = "Optimal Operating Point:\n"
                for j, var in enumerate(config["inputs"]): text += f" • {var['name']}: {ind[j]:.2f}\n"
                for k, obj in enumerate(config["objectives"]): text += f" • {obj['name']}: {ind.fitness.values[k] / obj['weight']:.2f}\n"
                sel.annotation.set_text(text.strip())
                sel.annotation.get_bbox_patch().set(fc="#111111", alpha=0.95, edgecolor="#00ffcc", lw=1.5)

        canvas = FigureCanvasTkAgg(fig, master=self.results_container)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True)
        toolbar_frame = ctk.CTkFrame(self.results_container, height=35, fg_color="transparent")
        toolbar_frame.pack(fill="x", pady=(5, 0))
        NavigationToolbar2Tk(canvas, toolbar_frame).update()
        self.tabs.set("5. Results Dashboard")

    def generate_pdf(self):
        if not FPDF:
            print("[ERROR] PDF library 'fpdf' not found. Please run 'pip install fpdf' in your terminal.")
            return
        if not self.final_pareto_data: return
        
        print("[SYSTEM] Compiling Executive PDF Report...")
        pareto_front, config = self.final_pareto_data
        
        best_ind = max(pareto_front, key=lambda ind: ind.fitness.values[0])
        
        pdf = FPDF()
        pdf.add_page()
        pdf.set_font("Arial", 'B', 16)
        pdf.cell(200, 10, txt="AI Methanol Optimization - Research Summary", ln=True, align='C')
        
        pdf.set_font("Arial", 'I', 10)
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        pdf.cell(200, 10, txt=f"Generated via Methanol AI Research Edition | {timestamp}", ln=True, align='C')
        pdf.ln(10)
        
        pdf.set_font("Arial", 'B', 14)
        pdf.cell(200, 10, txt="1. Recommended Operating Parameters (Decision Variables)", ln=True)
        pdf.set_font("Arial", '', 12)
        for j, var in enumerate(config["inputs"]):
            pdf.cell(200, 8, txt=f"   - {var['name']}: {best_ind[j]:.3f}", ln=True)
        pdf.ln(5)
        
        pdf.set_font("Arial", 'B', 14)
        pdf.cell(200, 10, txt="2. Projected Methanol Yield Outcomes (Objectives)", ln=True)
        pdf.set_font("Arial", '', 12)
        for k, obj in enumerate(config["objectives"]):
            val = best_ind.fitness.values[k] / obj["weight"]
            pdf.cell(200, 8, txt=f"   - {obj['name']}: {val:.3f}", ln=True)
        
        desktop_path = Path.home() / "Desktop"
        if not desktop_path.exists(): desktop_path = Path.home() / "OneDrive" / "Desktop"
        pdf_path = desktop_path / "Methanol_Research_Report.pdf"
        
        pdf.output(str(pdf_path))
        print(f"[SUCCESS] PDF Report Generated: {pdf_path}")
        self.telem_status.configure(text="PDF Exported", text_color="#8e44ad")


if __name__ == "__main__":
    app = AspenOptimizerUI()
    app.mainloop()