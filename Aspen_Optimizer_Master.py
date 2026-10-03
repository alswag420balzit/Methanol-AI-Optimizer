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
    blocks, streams, components = [], [], []
    try:
        b_node = aspen.Tree.FindNode(r"\Data\Blocks")
        if b_node: blocks = [b.Name for b in b_node.Elements]
            
        s_node = aspen.Tree.FindNode(r"\Data\Streams")
        if s_node: streams = [s.Name for s in s_node.Elements]
        
        c_node = aspen.Tree.FindNode(r"\Data\Components\Specifications\Input\COMP-ID")
        if c_node: components = [c.Name for c in c_node.Elements]
    except Exception:
        pass
    
    if not components:
        components = ["H2", "CO", "CO2", "METHANOL", "WATER", "METHANE", "N2"]
        
    return blocks, streams, components

def build_aspen_path(is_input, t_type, target, prop, comp=""):
    if is_input:
        if t_type == "Block":
            return rf"\Data\Blocks\{target}\Input\{prop}"
        elif t_type == "Stream":
            if prop == "Component Flow": return rf"\Data\Streams\{target}\Input\FLOW\MIXED\{comp}"
            else: return rf"\Data\Streams\{target}\Input\{prop}"
    else:
        if t_type == "Block":
            return rf"\Data\Blocks\{target}\Output\{prop}"
        elif t_type == "Stream":
            if prop == "Component Mass Flow": return rf"\Data\Streams\{target}\Output\MASSFLOW\MIXED\{comp}"
            elif prop == "Component Mole Flow": return rf"\Data\Streams\{target}\Output\MOLEFLOW\MIXED\{comp}"
            elif prop == "Component Mass Frac": return rf"\Data\Streams\{target}\Output\MASSFRAC\MIXED\{comp}"
            elif prop == "Component Mole Frac": return rf"\Data\Streams\{target}\Output\MOLEFRAC\MIXED\{comp}"
            elif prop == "Total Mass Flow": return rf"\Data\Streams\{target}\Output\RES_MASSFLOW"
            elif prop == "Total Mole Flow": return rf"\Data\Streams\{target}\Output\RES_MOLEFLOW"
            else: return rf"\Data\Streams\{target}\Output\{prop}"

def get_actual_inputs(individual, config, normalize):
    vals = list(individual)
    if not normalize: return vals
    
    feed_indices = [i for i, var in enumerate(config["inputs"]) if "FLOW" in var["path"].upper()]
    if feed_indices:
        lurgi_reactive_mass = 43998.6 
        current_sum = sum(vals[i] for i in feed_indices)
        if current_sum > 0:
            for i in feed_indices:
                vals[i] = (vals[i] / current_sum) * lurgi_reactive_mass
    return vals

def evaluate_flowsheet(individual, config, aspen, normalize):
    actual_inputs = get_actual_inputs(individual, config, normalize)
    ind_key = tuple(round(x, 4) for x in actual_inputs)
    if ind_key in simulation_cache: return simulation_cache[ind_key]
        
    try:
        for i, input_var in enumerate(config["inputs"]):
            node = aspen.Tree.FindNode(input_var["path"])
            if node: node.Value = actual_inputs[i]
                
        aspen.Engine.Run2()
        
        error_node = aspen.Tree.FindNode(r"\Data\Results Summary\Run-Status\NERROR")
        if error_node and error_node.Value is not None and int(error_node.Value) > 0:
            failure = tuple([-999999.0] * len(config["objectives"]))
            simulation_cache[ind_key] = failure
            return failure
        
        for const in config["constraints"]:
            try:
                node = aspen.Tree.FindNode(const["path"])
                if node and node.Value is not None:
                    val = float(node.Value)
                    if abs(val) > 1e20 or (const["type"] == "MAX" and val > const["limit"]) or (const["type"] == "MIN" and val < const["limit"]):
                        return tuple([-999999.0] * len(config["objectives"]))
                else: return tuple([-999999.0] * len(config["objectives"]))
            except Exception: return tuple([-999999.0] * len(config["objectives"]))
        
        results = []
        for obj in config["objectives"]:
            try:
                node = aspen.Tree.FindNode(obj["path"])
                if node and node.Value is not None:
                    val = float(node.Value)
                    if abs(val) > 1e20: results.append(-999999.0)
                    else: results.append(val * obj["weight"])
                else: results.append(-999999.0)
            except Exception: results.append(-999999.0)
                
        final_result = tuple(results)
        simulation_cache[ind_key] = final_result
        return final_result
        
    except Exception: return tuple([-999999.0] * len(config["objectives"]))

def run_optimization(config, aspen, pop_size, generations, normalize, progress_callback=None, abort_check=None):
    num_objectives = len(config["objectives"])
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        creator.create("FitnessMulti", base.Fitness, weights=tuple([1.0] * num_objectives))
        creator.create("Individual", list, fitness=creator.FitnessMulti)

    toolbox = base.Toolbox()
    toolbox.register("attr_float", lambda bounds: random.uniform(bounds[0], bounds[1]))
    toolbox.register("individual", lambda: creator.Individual([toolbox.attr_float(var["bounds"]) for var in config["inputs"]]))
    toolbox.register("population", tools.initRepeat, list, toolbox.individual)
    toolbox.register("evaluate", evaluate_flowsheet, config=config, aspen=aspen, normalize=normalize)
    toolbox.register("mate", tools.cxSimulatedBinaryBounded, low=[v["bounds"][0] for v in config["inputs"]], up=[v["bounds"][1] for v in config["inputs"]], eta=20.0)
    toolbox.register("mutate", tools.mutPolynomialBounded, low=[v["bounds"][0] for v in config["inputs"]], up=[v["bounds"][1] for v in config["inputs"]], eta=20.0, indpb=1.0/len(config["inputs"]))
    toolbox.register("select", tools.selNSGA2)

    pop = toolbox.population(n=pop_size)
    generational_best = []
    
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
            
        for ind, fit in zip(offspring[:len(fits)], fits): ind.fitness.values = fit
            
        evaluated_offspring = [ind for ind in offspring if ind.fitness.valid]
        pop = toolbox.select(pop + evaluated_offspring, k=min(pop_size, len(pop + evaluated_offspring)))

        if pop:
            best_in_gen = max(pop, key=lambda ind: ind.fitness.values[0])
            gen_record = {"Generation": gen_num}
            actual_vals = get_actual_inputs(best_in_gen, config, normalize)
            for j, var in enumerate(config["inputs"]): gen_record[var['name']] = round(actual_vals[j], 4)
            for k, obj in enumerate(config["objectives"]): gen_record[obj['name']] = round(best_in_gen.fitness.values[k] / obj["weight"], 4)
            generational_best.append(gen_record)

    if not pop: return [], config

    pareto_front = tools.sortNondominated(pop, len(pop), first_front_only=True)[0]
    
    desktop_path = Path.home() / "Desktop"
    if not desktop_path.exists(): desktop_path = Path.home() / "OneDrive" / "Desktop"
    
    csv_data = []
    for ind in pareto_front:
        row = {}
        actual_vals = get_actual_inputs(ind, config, normalize)
        for j, var in enumerate(config["inputs"]): row[var['name']] = round(actual_vals[j], 4)
        for k, obj in enumerate(config["objectives"]): row[obj['name']] = round(ind.fitness.values[k] / obj["weight"], 4)
        csv_data.append(row)
        
    df_pareto = pd.DataFrame(csv_data)
    pareto_file = desktop_path / "Aspen_Pareto_Results.csv"
    df_pareto.to_csv(pareto_file, index=False)
    
    if generational_best:
        df_history = pd.DataFrame(generational_best)
        history_file = desktop_path / "Aspen_Generational_Learning.csv"
        df_history.to_csv(history_file, index=False)

    print(f"[FILE IO] Optimal dataset successfully written to: {pareto_file}")
    return pareto_front, config

# ==========================================
# 3. GRAPHICAL USER INTERFACE
# ==========================================
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

IN_PROPS = {"Block": ["TEMP", "PRES", "CTEMP", "HEAT-DUTY", "VFRAC", "NTRAY", "FEED-STAGE"], "Stream": ["Component Flow", "TEMP", "PRES"]}
OUT_PROPS = {"Block": ["TEMP_OUT", "PRES_OUT", "DUTY"], "Stream": ["Component Mass Flow", "Component Mole Flow", "Component Mass Frac", "Component Mole Frac", "Total Mass Flow", "Total Mole Flow", "TEMP", "PRES"]}

class AspenOptimizerUI(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Universal Aspen Optimizer — AI Edition")
        self.geometry("1180x820")
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self.aspen_link = None
        self.available_blocks = ["(Load File First)"]
        self.available_streams = ["(Load File First)"]
        self.available_comps = ["(Load File First)"]
        
        self.var_rows = []
        self.constant_rows = []
        self.obj_rows = []
        self.const_rows = []
        
        self.abort_flag = False
        self.final_pareto_data = None

        self.sidebar = ctk.CTkFrame(self, width=230, corner_radius=0)
        self.sidebar.grid(row=0, column=0, sticky="nsew")
        self.sidebar.grid_rowconfigure(5, weight=1) 

        self.logo_label = ctk.CTkLabel(self.sidebar, text="⚡ UNIVERSAL\nAI OPTIMIZER", font=ctk.CTkFont(size=20, weight="bold"))
        self.logo_label.grid(row=0, column=0, padx=20, pady=(30, 5))
        self.version_label = ctk.CTkLabel(self.sidebar, text="Unlimited Edition v2.0", text_color="gray", font=ctk.CTkFont(size=12))
        self.version_label.grid(row=1, column=0, padx=20, pady=(0, 15))

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

        self.tabs = ctk.CTkTabview(self)
        self.tabs.grid(row=0, column=1, sticky="nsew", padx=20, pady=(10, 0))

        self.tab_engine = self.tabs.add("1. Engine & File")
        self.tab_vars = self.tabs.add("2. Variables")
        self.tab_consts_setup = self.tabs.add("3. Constants")
        self.tab_objs = self.tabs.add("4. Objectives")
        self.tab_consts = self.tabs.add("5. Constraints")
        self.tab_results = self.tabs.add("6. Results")

        self.setup_engine_tab()
        self.setup_vars_tab()
        self.setup_constants_tab()
        self.setup_objs_tab()
        self.setup_consts_tab()
        self.setup_results_tab()

        self.bottom_frame = ctk.CTkFrame(self, height=180, corner_radius=0, fg_color="gray10")
        self.bottom_frame.grid(row=1, column=0, columnspan=2, sticky="nsew")
        self.bottom_frame.grid_columnconfigure(0, weight=1)

        term_header = ctk.CTkFrame(self.bottom_frame, fg_color="transparent", height=25)
        term_header.pack(fill="x", padx=15, pady=(5, 0))
        ctk.CTkLabel(term_header, text="COMMAND CENTER CONSOLE OUTPUT", font=ctk.CTkFont(size=11, weight="bold"), text_color="#1d865c").pack(side="left")

        self.console_box = ctk.CTkTextbox(self.bottom_frame, height=140, font=ctk.CTkFont(family="Consolas", size=11), state="disabled", fg_color="black", text_color="#00ff66")
        self.console_box.pack(fill="both", expand=True, padx=15, pady=(2, 10))

        sys.stdout = ConsoleRedirector(self.console_box)
        print("[CONSOLE] Universal AI Engine initialized (v2.0).")
        print("[INFO] Paths are now fully dynamic. Connect to Aspen to scan your component list.")

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
        
        ctk.CTkButton(preset_frame, text="Save Config", width=90, fg_color="#27ae60", hover_color="#1e8449", command=self.save_config).pack(side="left", padx=10, pady=12)
        ctk.CTkButton(preset_frame, text="Load Config", width=90, fg_color="#2980b9", hover_color="#1f618d", command=self.load_config).pack(side="left", padx=(0, 10))
        ctk.CTkButton(preset_frame, text="Test Paths", width=90, fg_color="#8e44ad", hover_color="#732d91", command=self.test_paths).pack(side="left", padx=(0, 10))
        ctk.CTkButton(preset_frame, text="Apply Constants to Aspen", width=140, fg_color="#c0392b", hover_color="#922b21", command=self.apply_constants).pack(side="left", padx=(0, 10))

        self.ai_label = ctk.CTkLabel(self.tab_engine, text="Evolutionary Search Hyperparameters:", font=ctk.CTkFont(weight="bold"))
        self.ai_label.pack(anchor="w", padx=20, pady=(15, 5))

        frame2 = ctk.CTkFrame(self.tab_engine, fg_color="transparent")
        frame2.pack(fill="x", padx=20)
        
        ctk.CTkLabel(frame2, text="Population Size:").pack(side="left")
        self.pop_entry = ctk.CTkEntry(frame2, width=60); self.pop_entry.insert(0, "40"); self.pop_entry.pack(side="left", padx=(10, 30))
        ctk.CTkLabel(frame2, text="Generations:").pack(side="left")
        self.gen_entry = ctk.CTkEntry(frame2, width=60); self.gen_entry.insert(0, "20"); self.gen_entry.pack(side="left", padx=(10, 30))
        
        self.norm_var = ctk.BooleanVar(value=False)
        self.norm_check = ctk.CTkCheckBox(frame2, text="Enable Fair Comparison Stoichiometric Normalization (43,998 kg/hr)", variable=self.norm_var)
        self.norm_check.pack(side="left", padx=(10, 0))

        self.run_btn = ctk.CTkButton(
            self.tab_engine, text="LAUNCH AI OPTIMIZATION", font=ctk.CTkFont(size=16, weight="bold"), 
            fg_color="#1d865c", hover_color="#145c3f", height=45, command=self.toggle_engine_state
        )
        self.run_btn.pack(fill="x", padx=20, pady=25)

    def browse_file(self):
        filename = filedialog.askopenfilename(filetypes=[("Aspen Plus Workspaces", "*.apwz")])
        if filename:
            self.file_entry.delete(0, 'end'); self.file_entry.insert(0, filename)
            print(f"[ASPEN] Selected workspace: {filename}")

    def scan_file(self):
        filepath = self.file_entry.get()
        if not filepath or not os.path.exists(filepath): return
        print("[ASPEN] Initializing COM link and cataloging plant topology...")
        self.telem_status.configure(text="Scanning Topology...", text_color="#f39c12"); self.update()
        if not self.aspen_link: self.aspen_link = connect_to_aspen(filepath)
        if self.aspen_link:
            blocks, streams, comps = scan_aspen_inventory(self.aspen_link)
            self.available_blocks = blocks if blocks else ["None Found"]
            self.available_streams = streams if streams else ["None Found"]
            self.available_comps = comps if comps else ["None Found"]
            self.refresh_dropdowns()
            self.telem_status.configure(text="Connected / Ready", text_color="#1d865c")
            print(f"[ASPEN] Scan successful: Identified {len(blocks)} Blocks, {len(streams)} Streams, and {len(comps)} Components.")

    def refresh_dropdowns(self):
        for row in self.var_rows + self.constant_rows: 
            row['target'].configure(values=self.available_blocks if row['type'].get() == "Block" else self.available_streams)
            row['comp'].configure(values=self.available_comps)
        for row in self.obj_rows + self.const_rows:
            row['target'].configure(values=self.available_blocks if row['type'].get() == "Block" else self.available_streams)
            row['comp'].configure(values=self.available_comps)

    def _create_dynamic_row(self, parent, is_input, mode_type, vals):
        row = ctk.CTkFrame(parent)
        row.pack(fill="x", pady=2)
        
        props_dict = IN_PROPS if is_input else OUT_PROPS
        
        e_type = ctk.CTkOptionMenu(row, width=90, values=["Block", "Stream"])
        e_target = ctk.CTkComboBox(row, width=130, values=self.available_blocks)
        e_prop = ctk.CTkComboBox(row, width=180, values=props_dict["Block"])
        e_comp = ctk.CTkComboBox(row, width=100, values=self.available_comps)
        
        def update_dropdowns(choice, prop_val=None):
            if choice == "Block":
                e_target.configure(values=self.available_blocks)
                if not e_target.get() in self.available_blocks: e_target.set(self.available_blocks[0] if self.available_blocks else "")
                e_prop.configure(values=props_dict["Block"])
                e_prop.set(prop_val if prop_val in props_dict["Block"] else props_dict["Block"][0])
                e_comp.set(""); e_comp.configure(state="disabled")
            else:
                e_target.configure(values=self.available_streams)
                if not e_target.get() in self.available_streams: e_target.set(self.available_streams[0] if self.available_streams else "")
                e_prop.configure(values=props_dict["Stream"])
                e_prop.set(prop_val if prop_val in props_dict["Stream"] else props_dict["Stream"][0])
                check_comp_state(e_prop.get())

        def check_comp_state(choice):
            if choice.startswith("Component"):
                e_comp.configure(state="normal")
                if not e_comp.get(): e_comp.set(self.available_comps[0] if self.available_comps else "")
            else:
                e_comp.set(""); e_comp.configure(state="disabled")

        e_type.configure(command=lambda c: update_dropdowns(c))
        e_prop.configure(command=check_comp_state)
        
        e_type.set(vals.get("type", "Block"))
        update_dropdowns(e_type.get(), vals.get("prop", ""))
        e_target.set(vals.get("target", ""))
        e_comp.set(vals.get("comp", ""))
        
        e_type.pack(side="left", padx=5); e_target.pack(side="left", padx=5)
        e_prop.pack(side="left", padx=5); e_comp.pack(side="left", padx=5)
        
        row_data = {"frame": row, "type": e_type, "target": e_target, "prop": e_prop, "comp": e_comp}
        return row, row_data

    def setup_vars_tab(self):
        ctk.CTkButton(self.tab_vars, text="+ Add Variable", command=lambda: self.add_var_row({})).pack(anchor="w", padx=20, pady=10)
        self.vars_scroll = ctk.CTkScrollableFrame(self.tab_vars, fg_color="transparent")
        self.vars_scroll.pack(fill="both", expand=True, padx=10, pady=5)

    def add_var_row(self, vals):
        row, row_data = self._create_dynamic_row(self.vars_scroll, True, "VAR", vals)
        e_min = ctk.CTkEntry(row, width=80, placeholder_text="Min"); e_min.insert(0, vals.get("min", "")); e_min.pack(side="left", padx=5)
        e_max = ctk.CTkEntry(row, width=80, placeholder_text="Max"); e_max.insert(0, vals.get("max", "")); e_max.pack(side="left", padx=5)
        row_data.update({"min": e_min, "max": e_max})
        self.var_rows.append(row_data)
        ctk.CTkButton(row, text="X", width=30, fg_color="#c0392b", hover_color="#922b21", command=lambda: self.delete_row(row_data, self.var_rows)).pack(side="left", padx=5)

    def setup_constants_tab(self):
        ctk.CTkButton(self.tab_consts_setup, text="+ Add Constant", command=lambda: self.add_constant_row({})).pack(anchor="w", padx=20, pady=10)
        self.constants_scroll = ctk.CTkScrollableFrame(self.tab_consts_setup, fg_color="transparent")
        self.constants_scroll.pack(fill="both", expand=True, padx=10, pady=5)

    def add_constant_row(self, vals):
        row, row_data = self._create_dynamic_row(self.constants_scroll, True, "CONST", vals)
        e_val = ctk.CTkEntry(row, width=120, placeholder_text="Static Value"); e_val.insert(0, vals.get("val", "")); e_val.pack(side="left", padx=5)
        row_data.update({"val": e_val})
        self.constant_rows.append(row_data)
        ctk.CTkButton(row, text="X", width=30, fg_color="#c0392b", hover_color="#922b21", command=lambda: self.delete_row(row_data, self.constant_rows)).pack(side="left", padx=5)

    def setup_objs_tab(self):
        ctk.CTkButton(self.tab_objs, text="+ Add Objective", command=lambda: self.add_obj_row({})).pack(anchor="w", padx=20, pady=10)
        self.objs_scroll = ctk.CTkScrollableFrame(self.tab_objs, fg_color="transparent")
        self.objs_scroll.pack(fill="both", expand=True, padx=10, pady=5)
        if not self.obj_rows:
            self.add_obj_row({"type": "Stream", "prop": "Component Mass Flow", "comp": "METHANOL", "goal": "MAX"})
            self.add_obj_row({"type": "Stream", "prop": "Component Mass Flow", "comp": "WATER", "goal": "MIN"})

    def add_obj_row(self, vals):
        row, row_data = self._create_dynamic_row(self.objs_scroll, False, "OBJ", vals)
        e_goal = ctk.CTkOptionMenu(row, width=80, values=["MAX", "MIN"]); e_goal.set(vals.get("goal", "MAX")); e_goal.pack(side="left", padx=5)
        row_data.update({"goal": e_goal})
        self.obj_rows.append(row_data)
        ctk.CTkButton(row, text="X", width=30, fg_color="#c0392b", hover_color="#922b21", command=lambda: self.delete_row(row_data, self.obj_rows)).pack(side="left", padx=5)

    def setup_consts_tab(self):
        ctk.CTkButton(self.tab_consts, text="+ Add Constraint", command=lambda: self.add_const_row({})).pack(anchor="w", padx=20, pady=10)
        self.consts_scroll = ctk.CTkScrollableFrame(self.tab_consts, fg_color="transparent")
        self.consts_scroll.pack(fill="both", expand=True, padx=10, pady=5)

    def add_const_row(self, vals):
        row, row_data = self._create_dynamic_row(self.consts_scroll, False, "LIMIT", vals)
        e_limit = ctk.CTkOptionMenu(row, width=80, values=["MIN", "MAX"]); e_limit.set(vals.get("limit", "MIN")); e_limit.pack(side="left", padx=5)
        e_val = ctk.CTkEntry(row, width=80, placeholder_text="Value"); e_val.insert(0, vals.get("val", "")); e_val.pack(side="left", padx=5)
        row_data.update({"limit": e_limit, "val": e_val})
        self.const_rows.append(row_data)
        ctk.CTkButton(row, text="X", width=30, fg_color="#c0392b", hover_color="#922b21", command=lambda: self.delete_row(row_data, self.const_rows)).pack(side="left", padx=5)

    def setup_results_tab(self):
        self.results_header = ctk.CTkFrame(self.tab_results, fg_color="transparent", height=40)
        self.results_header.pack(fill="x", padx=10, pady=5)
        self.pdf_btn = ctk.CTkButton(self.results_header, text="📄 Export Executive PDF Report", fg_color="#8e44ad", hover_color="#732d91", state="disabled", command=self.generate_pdf)
        self.pdf_btn.pack(side="right")
        self.results_container = ctk.CTkFrame(self.tab_results, fg_color="transparent")
        self.results_container.pack(fill="both", expand=True, padx=10, pady=0)
        ctk.CTkLabel(self.results_container, text="Execution Pending...", font=ctk.CTkFont(size=14, slant="italic"), text_color="gray").pack(expand=True)

    def apply_constants(self, silent=False):
        if not silent: print("\n" + "="*60 + "\n[SYSTEM] Injecting Static Constants into Aspen Baseline...")
        filepath = self.file_entry.get()
        if not filepath:
            if not silent: print("[ERROR] Please select your Aspen .apwz file first.")
            return False
        if not self.aspen_link: self.aspen_link = connect_to_aspen(filepath)
        if not self.aspen_link:
            if not silent: print("[ERROR] Could not connect to Aspen COM server.")
            return False
        if not self.constant_rows:
            if not silent: print("[INFO] No constants defined in Tab 3. Skipping injection.")
            return True

        try:
            for row in self.constant_rows:
                t_type, target, prop, comp, val_str = row['type'].get(), row['target'].get(), row['prop'].get(), row['comp'].get(), row['val'].get()
                if not target or not prop or not val_str or "Load File" in target: continue
                path_str = build_aspen_path(True, t_type, target, prop, comp)
                node = self.aspen_link.Tree.FindNode(path_str)
                if node:
                    node.Value = float(val_str)
                    if not silent: print(f" -> Set {path_str.split(chr(92))[-1]} to {val_str}")
                else:
                    if not silent: print(f" -> ERROR: Node not found {path_str}")
            self.aspen_link.Save()
            if not silent: print("[SUCCESS] Constants fully applied and saved.")
            return True
        except Exception as e:
            if not silent: print(f"[ERROR] Failed to apply constants: {str(e)}")
            return False
        finally:
            if not silent: print("="*60 + "\n")

    def test_paths(self):
        print("\n" + "="*60 + "\n[DIAGNOSTIC] Verifying Dynamic Aspen Paths...")
        filepath = self.file_entry.get()
        if not filepath:
            print("[ERROR] Please select your Aspen file first."); return
        if not self.aspen_link: self.aspen_link = connect_to_aspen(filepath)
        if not self.aspen_link:
            print("[ERROR] Could not connect."); return
            
        print(f"{'PARAMETER NAME':<40} | {'LIVE VALUE'}\n" + "-" * 65)
        all_checks = [(self.var_rows, True, "VAR"), (self.constant_rows, True, "CONST"), (self.obj_rows, False, "OBJ"), (self.const_rows, False, "LIMIT")]
        
        for rows, is_in, tag in all_checks:
            for row in rows:
                t_type, target, prop, comp = row['type'].get(), row['target'].get(), row['prop'].get(), row['comp'].get()
                if not target or not prop or "Load File" in target: continue
                path_str = build_aspen_path(is_in, t_type, target, prop, comp)
                short_name = f"{tag}: {target}_{comp if comp else prop.split(chr(92))[-1]}"
                try:
                    node = self.aspen_link.Tree.FindNode(path_str)
                    if node and node.Value is not None: print(f"{short_name:<40} | {node.Value:<15.4f}")
                    else: print(f"{short_name:<40} | ERROR: Node Not Found")
                except Exception: print(f"{short_name:<40} | ERROR: Invalid Path")
        print("="*60 + "\n")

    def save_config(self):
        data = {"engine": {"pop": self.pop_entry.get(), "gen": self.gen_entry.get(), "norm": self.norm_var.get()}, "vars": [], "constants": [], "objs": [], "consts": []}
        for r in self.var_rows: data["vars"].append({"type": r['type'].get(), "target": r['target'].get(), "prop": r['prop'].get(), "comp": r['comp'].get(), "min": r['min'].get(), "max": r['max'].get()})
        for r in self.constant_rows: data["constants"].append({"type": r['type'].get(), "target": r['target'].get(), "prop": r['prop'].get(), "comp": r['comp'].get(), "val": r['val'].get()})
        for r in self.obj_rows: data["objs"].append({"type": r['type'].get(), "target": r['target'].get(), "prop": r['prop'].get(), "comp": r['comp'].get(), "goal": r['goal'].get()})
        for r in self.const_rows: data["consts"].append({"type": r['type'].get(), "target": r['target'].get(), "prop": r['prop'].get(), "comp": r['comp'].get(), "limit": r['limit'].get(), "val": r['val'].get()})
        filepath = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("JSON Files", "*.json")])
        if filepath:
            with open(filepath, 'w') as f: json.dump(data, f, indent=4)
            print(f"[SYSTEM] Config saved: {filepath}")

    def load_config(self):
        filepath = filedialog.askopenfilename(filetypes=[("JSON Files", "*.json")])
        if filepath:
            with open(filepath, 'r') as f: data = json.load(f)
            self.pop_entry.delete(0, 'end'); self.pop_entry.insert(0, data["engine"]["pop"])
            self.gen_entry.delete(0, 'end'); self.gen_entry.insert(0, data["engine"]["gen"])
            self.norm_var.set(data["engine"].get("norm", False))
            for r in list(self.var_rows): self.delete_row(r, self.var_rows)
            for r in list(self.constant_rows): self.delete_row(r, self.constant_rows)
            for r in list(self.obj_rows): self.delete_row(r, self.obj_rows)
            for r in list(self.const_rows): self.delete_row(r, self.const_rows)
            for v in data.get("vars", []): self.add_var_row(v)
            for c in data.get("constants", []): self.add_constant_row(c)
            for o in data.get("objs", []): self.add_obj_row(o)
            for l in data.get("consts", []): self.add_const_row(l)
            print(f"[SYSTEM] Config loaded: {filepath}")

    def toggle_engine_state(self):
        if self.run_btn.cget("text") == "LAUNCH AI OPTIMIZATION":
            self.abort_flag = False
            self.run_btn.configure(text="🛑 ABORT OPERATION (SAFE STOP)", fg_color="#c0392b", hover_color="#922b21")
            self.launch_engine()
        else:
            print("[SYSTEM] User initiated Emergency Stop..."); self.abort_flag = True

    def reset_run_btn(self):
        self.run_btn.configure(text="LAUNCH AI OPTIMIZATION", fg_color="#1d865c", hover_color="#145c3f")

    def launch_engine(self):
        filepath = self.file_entry.get()
        if not filepath: 
            print("[ERROR] Flowsheet not loaded."); self.reset_run_btn(); return
            
        config = {"inputs": [], "objectives": [], "constraints": []}
        try:
            pop_size, generations = int(self.pop_entry.get()), int(self.gen_entry.get())
            normalize = self.norm_var.get()
            
            for row in self.var_rows:
                t_type, target, prop, comp, m, mx = row['type'].get(), row['target'].get(), row['prop'].get(), row['comp'].get(), row['min'].get(), row['max'].get()
                if not target or not m or not mx or "Load File" in target: continue
                path_str = build_aspen_path(True, t_type, target, prop, comp)
                name_str = f"{target}_{comp if comp else prop}"
                vmin, vmax = float(m), float(mx)
                if vmin > vmax: vmin, vmax = vmax, vmin
                config["inputs"].append({"name": name_str, "path": path_str, "bounds": (vmin, vmax)})
                
            for row in self.obj_rows:
                t_type, target, prop, comp, goal = row['type'].get(), row['target'].get(), row['prop'].get(), row['comp'].get(), row['goal'].get()
                if not target or "Load File" in target: continue
                path_str = build_aspen_path(False, t_type, target, prop, comp)
                name_str = f"{target}_{comp if comp else prop}"
                weight = 1.0 if goal == "MAX" else -1.0
                config["objectives"].append({"name": name_str, "path": path_str, "weight": weight})

            for row in self.const_rows:
                t_type, target, prop, comp, limit, val = row['type'].get(), row['target'].get(), row['prop'].get(), row['comp'].get(), row['limit'].get(), row['val'].get()
                if not target or not val or "Load File" in target: continue
                path_str = build_aspen_path(False, t_type, target, prop, comp)
                config["constraints"].append({"name": f"{target}_Limit", "path": path_str, "type": limit, "limit": float(val)})
                
            if len(config["inputs"]) < 1 or len(config["objectives"]) < 1:
                print("[ERROR] Need >=1 Variable and >=1 Objective."); self.reset_run_btn(); return
        except ValueError:
            print("[ERROR] Type parsing failure: Check number formats."); self.reset_run_btn(); return

        print("[ENGINE] Preparing Aspen link for evolutionary run...")
        self.telem_status.configure(text="Simulating...", text_color="#2ecc71"); self.progress_bar.set(0.0); self.update()

        try:
            if not self.aspen_link: self.aspen_link = connect_to_aspen(filepath)
            if not self.aspen_link: print("[ERROR] Aspen COM link failed."); return
            self.apply_constants(silent=True)

            pareto_front, config = run_optimization(config, self.aspen_link, pop_size, generations, normalize, progress_callback=self.update_telemetry, abort_check=lambda: self.abort_flag)
            
            if not pareto_front:
                self.telem_status.configure(text="Aborted", text_color="#e74c3c")
                print("[SYSTEM] Halted before valid designs were found."); return

            self.final_pareto_data = (pareto_front, config)
            self.pdf_btn.configure(state="normal")
            
            self.telem_status.configure(text="Run Complete", text_color="#1d865c")
            print("[SUCCESS] Evolutionary cycle concluded. Rendering Pareto front...")
            self.render_embedded_pareto(config, pareto_front)
            
        except Exception as e:
            print(f"[FATAL] Engine failed: {str(e)}")
        finally:
            self.reset_run_btn()

    def update_telemetry(self, gen_curr, gen_total, ind_curr, ind_total):
        if gen_curr == 0: pct = (ind_curr / ind_total) * (1 / (gen_total + 1))
        else: pct = (gen_curr / (gen_total + 1)) + ((ind_curr / ind_total) * (1 / (gen_total + 1)))
        self.telem_pct.configure(text=f"Progress: {int(pct * 100)}% (Gen {gen_curr}/{gen_total})")
        self.progress_bar.set(pct); self.update()

    def render_embedded_pareto(self, config, pareto_front):
        for w in self.results_container.winfo_children(): w.destroy()
        if len(config["objectives"]) < 2:
            ctk.CTkLabel(self.results_container, text="Data saved. Add 2+ objectives to plot Pareto curve.").pack(expand=True)
            return

        obj1, obj2 = config["objectives"][0], config["objectives"][1]
        obj1_vals = [ind.fitness.values[0] / obj1["weight"] for ind in pareto_front]
        obj2_vals = [ind.fitness.values[1] / obj2["weight"] for ind in pareto_front]

        plt.style.use("dark_background")
        fig, ax = plt.subplots(figsize=(7, 4.5), dpi=100)
        fig.patch.set_facecolor("#1e1e1e"); ax.set_facecolor("#2b2b2b")
        scatter = ax.scatter(obj1_vals, obj2_vals, color="#00ffcc", edgecolor="white", s=65, zorder=3)
        ax.set_title(f"{obj2['name']} vs. {obj1['name']}", fontsize=12, fontweight="bold", color="white", pad=12)
        ax.set_xlabel(obj1['name'], fontsize=10, color="lightgray")
        ax.set_ylabel(obj2['name'], fontsize=10, color="lightgray")
        ax.grid(True, linestyle="--", alpha=0.3, zorder=0)
        fig.tight_layout()

        if mplcursors:
            cursor = mplcursors.cursor(scatter, hover=True)
            @cursor.connect("add")
            def on_add(sel):
                ind = pareto_front[sel.index]
                text = "Optimal Point:\n"
                actual_vals = get_actual_inputs(ind, config, self.norm_var.get())
                for j, var in enumerate(config["inputs"]): text += f" • {var['name']}: {actual_vals[j]:.2f}\n"
                for k, obj in enumerate(config["objectives"]): text += f" • {obj['name']}: {ind.fitness.values[k] / obj['weight']:.2f}\n"
                sel.annotation.set_text(text.strip())
                sel.annotation.get_bbox_patch().set(fc="#111111", alpha=0.95, edgecolor="#00ffcc", lw=1.5)

        canvas = FigureCanvasTkAgg(fig, master=self.results_container)
        canvas.draw(); canvas.get_tk_widget().pack(fill="both", expand=True)
        toolbar_frame = ctk.CTkFrame(self.results_container, height=35, fg_color="transparent"); toolbar_frame.pack(fill="x", pady=(5, 0))
        NavigationToolbar2Tk(canvas, toolbar_frame).update()
        self.tabs.set("6. Results")

    def generate_pdf(self):
        if not FPDF or not self.final_pareto_data: return
        pareto_front, config = self.final_pareto_data
        best_ind = max(pareto_front, key=lambda ind: ind.fitness.values[0])
        
        pdf = FPDF()
        pdf.add_page(); pdf.set_font("Arial", 'B', 16)
        pdf.cell(200, 10, txt="AI Plant Optimization - Global Report", ln=True, align='C'); pdf.ln(10)
        
        pdf.set_font("Arial", 'B', 14); pdf.cell(200, 10, txt="1. Recommended Operating Parameters", ln=True)
        pdf.set_font("Arial", '', 12)
        actual_vals = get_actual_inputs(best_ind, config, self.norm_var.get())
        for j, var in enumerate(config["inputs"]): pdf.cell(200, 8, txt=f"   - {var['name']}: {actual_vals[j]:.3f}", ln=True)
        
        pdf.ln(5); pdf.set_font("Arial", 'B', 14); pdf.cell(200, 10, txt="2. Projected Yield Outcomes", ln=True)
        pdf.set_font("Arial", '', 12)
        for k, obj in enumerate(config["objectives"]):
            val = best_ind.fitness.values[k] / obj["weight"]
            pdf.cell(200, 8, txt=f"   - {obj['name']}: {val:.3f}", ln=True)
        
        desktop_path = Path.home() / "Desktop"
        if not desktop_path.exists(): desktop_path = Path.home() / "OneDrive" / "Desktop"
        pdf_path = desktop_path / "AI_Optimization_Report.pdf"
        pdf.output(str(pdf_path))
        print(f"[SUCCESS] Report Saved: {pdf_path}")

if __name__ == "__main__":
    app = AspenOptimizerUI()
    app.mainloop()