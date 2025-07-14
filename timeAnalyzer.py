import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import seaborn as sns
from difflib import SequenceMatcher
import json
import os
from datetime import datetime
import sqlite3
from pathlib import Path

class TimeStudyApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Plataforma de Análise de Estudo de Tempos")
        self.root.geometry("1200x800")
        self.root.configure(bg='#f0f0f0')
        
        # Initialize database
        self.init_database()
        
        # Data storage
        self.uploaded_files = []
        self.activities = []
        self.groups = []
        self.statistics_results = []
        
        # Create GUI
        self.create_widgets()
        
    def init_database(self):
        """Initialize SQLite database"""
        self.db_path = "time_study.db"
        self.conn = sqlite3.connect(self.db_path)
        cursor = self.conn.cursor()
        
        # Create tables
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS files (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                filename TEXT,
                original_name TEXT,
                data TEXT,
                activity_column TEXT,
                time_column TEXT,
                processed INTEGER DEFAULT 0,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        ''')
        
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS activities (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT,
                original_names TEXT,
                occurrences INTEGER,
                times TEXT,
                unified INTEGER DEFAULT 0,
                group_id INTEGER,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        ''')
        
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS groups (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT,
                color TEXT,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        ''')
        
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS statistics (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                activity_id INTEGER,
                group_id INTEGER,
                count INTEGER,
                min_val REAL,
                max_val REAL,
                median_val REAL,
                q1 REAL,
                q3 REAL,
                iqr REAL,
                upper_fence REAL,
                lower_fence REAL,
                outlier_count INTEGER,
                non_outlier_count INTEGER,
                overall_mean REAL,
                non_outlier_mean REAL,
                total_activity_time REAL,
                non_normalized_time REAL,
                normalized_time REAL,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        ''')
        
        self.conn.commit()
        
    def create_widgets(self):
        """Create the main GUI widgets"""
        # Create notebook for tabs
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Tab 1: File Upload
        self.upload_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.upload_frame, text="1. Upload de Arquivos")
        self.create_upload_tab()
        
        # Tab 2: Column Mapping
        self.mapping_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.mapping_frame, text="2. Mapeamento de Colunas")
        self.create_mapping_tab()
        
        # Tab 3: Activity Unification
        self.unification_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.unification_frame, text="3. Unificação de Atividades")
        self.create_unification_tab()
        
        # Tab 4: Activity Grouping
        self.grouping_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.grouping_frame, text="4. Agrupamento de Atividades")
        self.create_grouping_tab()
        
        # Tab 5: Statistical Analysis
        self.analysis_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.analysis_frame, text="5. Análise Estatística")
        self.create_analysis_tab()
        
        # Tab 6: Export
        self.export_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.export_frame, text="6. Exportação")
        self.create_export_tab()
        
    def create_upload_tab(self):
        """Create file upload tab"""
        title_label = ttk.Label(self.upload_frame, text="Upload de Arquivos", font=("Arial", 16, "bold"))
        title_label.pack(pady=10)
        
        instruction_label = ttk.Label(self.upload_frame, text="Selecione arquivos Excel (.xlsx) ou CSV (.csv) contendo dados de estudo de tempos")
        instruction_label.pack(pady=5)
        
        # Upload button
        upload_btn = ttk.Button(self.upload_frame, text="Selecionar Arquivos", command=self.upload_files)
        upload_btn.pack(pady=10)
        
        # Files list
        self.files_listbox = tk.Listbox(self.upload_frame, height=10)
        self.files_listbox.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
        
        # Preview button
        preview_btn = ttk.Button(self.upload_frame, text="Visualizar Arquivo Selecionado", command=self.preview_file)
        preview_btn.pack(pady=5)
        
        # Preview text area
        self.preview_text = scrolledtext.ScrolledText(self.upload_frame, height=8)
        self.preview_text.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
        
    def create_mapping_tab(self):
        """Create column mapping tab"""
        title_label = ttk.Label(self.mapping_frame, text="Mapeamento de Colunas", font=("Arial", 16, "bold"))
        title_label.pack(pady=10)
        
        instruction_label = ttk.Label(self.mapping_frame, text="Selecione as colunas de atividade e tempo para cada arquivo")
        instruction_label.pack(pady=5)
        
        # Mapping frame
        mapping_scroll_frame = ttk.Frame(self.mapping_frame)
        mapping_scroll_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
        
        self.mapping_widgets = {}
        
        # Process button
        process_btn = ttk.Button(self.mapping_frame, text="Processar Arquivos", command=self.process_files)
        process_btn.pack(pady=10)
        
    def create_unification_tab(self):
        """Create activity unification tab"""
        title_label = ttk.Label(self.unification_frame, text="Unificação de Atividades", font=("Arial", 16, "bold"))
        title_label.pack(pady=10)
        
        instruction_label = ttk.Label(self.unification_frame, text="Revise e aprove sugestões de unificação de atividades similares")
        instruction_label.pack(pady=5)
        
        # Suggestions frame
        self.suggestions_frame = ttk.Frame(self.unification_frame)
        self.suggestions_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
        
        # Apply unification button
        unify_btn = ttk.Button(self.unification_frame, text="Aplicar Unificações", command=self.apply_unifications)
        unify_btn.pack(pady=10)
        
    def create_grouping_tab(self):
        """Create activity grouping tab"""
        title_label = ttk.Label(self.grouping_frame, text="Agrupamento de Atividades", font=("Arial", 16, "bold"))
        title_label.pack(pady=10)
        
        # Group creation frame
        group_creation_frame = ttk.LabelFrame(self.grouping_frame, text="Criar Novo Grupo")
        group_creation_frame.pack(fill=tk.X, padx=20, pady=10)
        
        ttk.Label(group_creation_frame, text="Nome do Grupo:").pack(side=tk.LEFT, padx=5)
        self.group_name_entry = ttk.Entry(group_creation_frame, width=30)
        self.group_name_entry.pack(side=tk.LEFT, padx=5)
        
        create_group_btn = ttk.Button(group_creation_frame, text="Criar Grupo", command=self.create_group)
        create_group_btn.pack(side=tk.LEFT, padx=5)
        
        # Main grouping frame
        main_grouping_frame = ttk.Frame(self.grouping_frame)
        main_grouping_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
        
        # Activities list
        activities_frame = ttk.LabelFrame(main_grouping_frame, text="Atividades")
        activities_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 10))
        
        self.activities_listbox = tk.Listbox(activities_frame, selectmode=tk.MULTIPLE)
        self.activities_listbox.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Groups list
        groups_frame = ttk.LabelFrame(main_grouping_frame, text="Grupos")
        groups_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)
        
        self.groups_listbox = tk.Listbox(groups_frame)
        self.groups_listbox.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Assign button
        assign_btn = ttk.Button(self.grouping_frame, text="Atribuir Atividades ao Grupo Selecionado", command=self.assign_to_group)
        assign_btn.pack(pady=10)
        
    def create_analysis_tab(self):
        """Create statistical analysis tab"""
        title_label = ttk.Label(self.analysis_frame, text="Análise Estatística", font=("Arial", 16, "bold"))
        title_label.pack(pady=10)
        
        # Calculate statistics button
        calc_btn = ttk.Button(self.analysis_frame, text="Calcular Estatísticas", command=self.calculate_statistics)
        calc_btn.pack(pady=10)
        
        # Results frame
        results_frame = ttk.Frame(self.analysis_frame)
        results_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
        
        # Statistics table
        self.stats_tree = ttk.Treeview(results_frame, columns=('Tipo', 'Nome', 'Contagem', 'Média', 'Mediana', 'Q1', 'Q3', 'IQR'), show='headings')
        self.stats_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Configure columns
        for col in self.stats_tree['columns']:
            self.stats_tree.heading(col, text=col)
            self.stats_tree.column(col, width=100)
        
        # Scrollbar for treeview
        stats_scrollbar = ttk.Scrollbar(results_frame, orient=tk.VERTICAL, command=self.stats_tree.yview)
        stats_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.stats_tree.configure(yscrollcommand=stats_scrollbar.set)
        
        # Plot frame
        self.plot_frame = ttk.Frame(self.analysis_frame)
        self.plot_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
        
    def create_export_tab(self):
        """Create export tab"""
        title_label = ttk.Label(self.export_frame, text="Exportação de Resultados", font=("Arial", 16, "bold"))
        title_label.pack(pady=10)
        
        instruction_label = ttk.Label(self.export_frame, text="Exporte os resultados da análise para arquivo Excel")
        instruction_label.pack(pady=5)
        
        export_btn = ttk.Button(self.export_frame, text="Exportar para Excel", command=self.export_to_excel)
        export_btn.pack(pady=20)
        
        # Export status
        self.export_status_label = ttk.Label(self.export_frame, text="")
        self.export_status_label.pack(pady=10)
        
    def upload_files(self):
        """Handle file upload"""
        file_paths = filedialog.askopenfilenames(
            title="Selecionar Arquivos",
            filetypes=[("Excel files", "*.xlsx"), ("CSV files", "*.csv"), ("All files", "*.*")]
        )
        
        for file_path in file_paths:
            try:
                # Read file
                if file_path.endswith('.xlsx'):
                    df = pd.read_excel(file_path)
                elif file_path.endswith('.csv'):
                    df = pd.read_csv(file_path)
                else:
                    continue
                
                # Store in database
                filename = os.path.basename(file_path)
                data_json = df.to_json()
                
                cursor = self.conn.cursor()
                cursor.execute('''
                    INSERT INTO files (filename, original_name, data)
                    VALUES (?, ?, ?)
                ''', (filename, filename, data_json))
                self.conn.commit()
                
                # Add to listbox
                self.files_listbox.insert(tk.END, filename)
                
                # Update mapping tab
                self.update_mapping_tab()
                
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao carregar arquivo {file_path}: {str(e)}")
    
    def preview_file(self):
        """Preview selected file"""
        selection = self.files_listbox.curselection()
        if not selection:
            messagebox.showwarning("Aviso", "Selecione um arquivo para visualizar")
            return
        
        filename = self.files_listbox.get(selection[0])
        
        # Get file data from database
        cursor = self.conn.cursor()
        cursor.execute('SELECT data FROM files WHERE filename = ?', (filename,))
        result = cursor.fetchone()
        
        if result:
            df = pd.read_json(result[0])
            preview_text = f"Arquivo: {filename}\n"
            preview_text += f"Linhas: {len(df)}, Colunas: {len(df.columns)}\n\n"
            preview_text += f"Colunas: {', '.join(df.columns)}\n\n"
            preview_text += "Primeiras 5 linhas:\n"
            preview_text += df.head().to_string()
            
            self.preview_text.delete(1.0, tk.END)
            self.preview_text.insert(1.0, preview_text)
    
    def update_mapping_tab(self):
        """Update column mapping tab with uploaded files"""
        # Clear existing widgets
        for widget in self.mapping_widgets.values():
            for w in widget:
                w.destroy()
        self.mapping_widgets.clear()
        
        # Get files from database
        cursor = self.conn.cursor()
        cursor.execute('SELECT id, filename, data, processed FROM files')
        files = cursor.fetchall()
        
        row = 0
        for file_id, filename, data_json, processed in files:
            df = pd.read_json(data_json)
            columns = df.columns.tolist()
            
            # File label
            file_label = ttk.Label(self.mapping_frame, text=f"Arquivo: {filename}", font=("Arial", 12, "bold"))
            file_label.grid(row=row, column=0, columnspan=4, sticky=tk.W, padx=20, pady=(10, 5))
            
            # Activity column
            ttk.Label(self.mapping_frame, text="Coluna de Atividade:").grid(row=row+1, column=0, padx=20, pady=2, sticky=tk.W)
            activity_combo = ttk.Combobox(self.mapping_frame, values=columns, width=20)
            activity_combo.grid(row=row+1, column=1, padx=10, pady=2)
            
            # Time column
            ttk.Label(self.mapping_frame, text="Coluna de Tempo:").grid(row=row+1, column=2, padx=20, pady=2, sticky=tk.W)
            time_combo = ttk.Combobox(self.mapping_frame, values=columns, width=20)
            time_combo.grid(row=row+1, column=3, padx=10, pady=2)
            
            self.mapping_widgets[file_id] = [file_label, activity_combo, time_combo]
            row += 2
    
    def process_files(self):
        """Process files with column mappings"""
        cursor = self.conn.cursor()
        cursor.execute('SELECT id, filename, data FROM files WHERE processed = 0')
        files = cursor.fetchall()
        
        all_activities = {}
        
        for file_id, filename, data_json in files:
            if file_id not in self.mapping_widgets:
                continue
            
            widgets = self.mapping_widgets[file_id]
            activity_column = widgets[1].get()
            time_column = widgets[2].get()
            
            if not activity_column or not time_column:
                messagebox.showwarning("Aviso", f"Selecione as colunas para o arquivo {filename}")
                continue
            
            # Update file with column mapping
            cursor.execute('''
                UPDATE files SET activity_column = ?, time_column = ?, processed = 1
                WHERE id = ?
            ''', (activity_column, time_column, file_id))
            
            # Extract activities
            df = pd.read_json(data_json)
            for _, row in df.iterrows():
                activity_name = str(row[activity_column])
                time_value = float(row[time_column]) if pd.notna(row[time_column]) else 0
                
                if activity_name in all_activities:
                    all_activities[activity_name].append(time_value)
                else:
                    all_activities[activity_name] = [time_value]
        
        # Store activities in database
        for activity_name, times in all_activities.items():
            times_json = json.dumps(times)
            cursor.execute('''
                INSERT INTO activities (name, original_names, occurrences, times)
                VALUES (?, ?, ?, ?)
            ''', (activity_name, json.dumps([activity_name]), len(times), times_json))
        
        self.conn.commit()
        
        # Update unification tab
        self.update_unification_tab()
        
        messagebox.showinfo("Sucesso", "Arquivos processados com sucesso!")
    
    def update_unification_tab(self):
        """Update unification tab with similarity suggestions"""
        # Clear existing widgets
        for widget in self.suggestions_frame.winfo_children():
            widget.destroy()
        
        # Get activities from database
        cursor = self.conn.cursor()
        cursor.execute('SELECT name FROM activities WHERE unified = 0')
        activity_names = [row[0] for row in cursor.fetchall()]
        
        # Find similar activities
        suggestions = self.find_similar_activities(activity_names)
        
        if not suggestions:
            ttk.Label(self.suggestions_frame, text="Nenhuma sugestão de unificação encontrada").pack(pady=20)
            return
        
        # Create suggestion widgets
        self.unification_vars = {}
        row = 0
        
        for suggestion in suggestions:
            frame = ttk.LabelFrame(self.suggestions_frame, text=f"Sugestão {row + 1}")
            frame.pack(fill=tk.X, padx=10, pady=5)
            
            # Activities list
            activities_text = "Atividades similares:\n" + "\n".join(suggestion['activities'])
            ttk.Label(frame, text=activities_text).pack(anchor=tk.W, padx=10, pady=5)
            
            # Checkbox to accept
            var = tk.BooleanVar()
            checkbox = ttk.Checkbutton(frame, text="Aceitar unificação", variable=var)
            checkbox.pack(anchor=tk.W, padx=10, pady=2)
            
            # Unified name entry
            ttk.Label(frame, text="Nome unificado:").pack(anchor=tk.W, padx=10, pady=(5, 0))
            unified_entry = ttk.Entry(frame, width=40)
            unified_entry.pack(anchor=tk.W, padx=10, pady=(0, 5))
            unified_entry.insert(0, suggestion['activities'][0])  # Default to first activity name
            
            self.unification_vars[row] = {
                'var': var,
                'entry': unified_entry,
                'activities': suggestion['activities']
            }
            row += 1
    
    def find_similar_activities(self, activity_names, threshold=0.8):
        """Find similar activity names"""
        suggestions = []
        processed = set()
        
        for i, name1 in enumerate(activity_names):
            if name1 in processed:
                continue
            
            similar_group = [name1]
            for j, name2 in enumerate(activity_names):
                if i != j and name2 not in processed:
                    similarity = SequenceMatcher(None, name1.lower(), name2.lower()).ratio()
                    if similarity >= threshold:
                        similar_group.append(name2)
                        processed.add(name2)
            
            if len(similar_group) > 1:
                suggestions.append({'activities': similar_group})
                processed.add(name1)
        
        return suggestions
    
    def apply_unifications(self):
        """Apply selected unifications"""
        cursor = self.conn.cursor()
        
        for suggestion_data in self.unification_vars.values():
            if suggestion_data['var'].get():  # If checkbox is checked
                unified_name = suggestion_data['entry'].get()
                activities = suggestion_data['activities']
                
                if not unified_name:
                    continue
                
                # Merge activities
                all_times = []
                original_names = []
                
                for activity_name in activities:
                    cursor.execute('SELECT times, original_names FROM activities WHERE name = ?', (activity_name,))
                    result = cursor.fetchone()
                    if result:
                        times = json.loads(result[0])
                        orig_names = json.loads(result[1])
                        all_times.extend(times)
                        original_names.extend(orig_names)
                
                # Update first activity with merged data
                first_activity = activities[0]
                cursor.execute('''
                    UPDATE activities 
                    SET name = ?, original_names = ?, occurrences = ?, times = ?, unified = 1
                    WHERE name = ?
                ''', (unified_name, json.dumps(original_names), len(all_times), json.dumps(all_times), first_activity))
                
                # Delete other activities
                for activity_name in activities[1:]:
                    cursor.execute('DELETE FROM activities WHERE name = ?', (activity_name,))
        
        self.conn.commit()
        self.update_grouping_tab()
        messagebox.showinfo("Sucesso", "Unificações aplicadas com sucesso!")
    
    def update_grouping_tab(self):
        """Update grouping tab with activities and groups"""
        # Clear activities listbox
        self.activities_listbox.delete(0, tk.END)
        
        # Get activities from database
        cursor = self.conn.cursor()
        cursor.execute('SELECT name FROM activities')
        activities = cursor.fetchall()
        
        for activity in activities:
            self.activities_listbox.insert(tk.END, activity[0])
        
        # Update groups listbox
        self.update_groups_listbox()
    
    def update_groups_listbox(self):
        """Update groups listbox"""
        self.groups_listbox.delete(0, tk.END)
        
        cursor = self.conn.cursor()
        cursor.execute('SELECT name FROM groups')
        groups = cursor.fetchall()
        
        for group in groups:
            self.groups_listbox.insert(tk.END, group[0])
    
    def create_group(self):
        """Create a new group"""
        group_name = self.group_name_entry.get().strip()
        if not group_name:
            messagebox.showwarning("Aviso", "Digite um nome para o grupo")
            return
        
        cursor = self.conn.cursor()
        cursor.execute('INSERT INTO groups (name, color) VALUES (?, ?)', (group_name, '#3498db'))
        self.conn.commit()
        
        self.group_name_entry.delete(0, tk.END)
        self.update_groups_listbox()
        messagebox.showinfo("Sucesso", f"Grupo '{group_name}' criado com sucesso!")
    
    def assign_to_group(self):
        """Assign selected activities to selected group"""
        selected_activities = [self.activities_listbox.get(i) for i in self.activities_listbox.curselection()]
        selected_group = self.groups_listbox.curselection()
        
        if not selected_activities:
            messagebox.showwarning("Aviso", "Selecione pelo menos uma atividade")
            return
        
        if not selected_group:
            messagebox.showwarning("Aviso", "Selecione um grupo")
            return
        
        group_name = self.groups_listbox.get(selected_group[0])
        
        # Get group ID
        cursor = self.conn.cursor()
        cursor.execute('SELECT id FROM groups WHERE name = ?', (group_name,))
        group_id = cursor.fetchone()[0]
        
        # Update activities
        for activity_name in selected_activities:
            cursor.execute('UPDATE activities SET group_id = ? WHERE name = ?', (group_id, activity_name))
        
        self.conn.commit()
        messagebox.showinfo("Sucesso", f"Atividades atribuídas ao grupo '{group_name}' com sucesso!")
    
    def calculate_statistics(self):
        """Calculate statistics for activities and groups"""
        cursor = self.conn.cursor()
        
        # Clear previous statistics
        cursor.execute('DELETE FROM statistics')
        
        # Calculate for individual activities
        cursor.execute('SELECT id, name, times, group_id FROM activities')
        activities = cursor.fetchall()
        
        for activity_id, name, times_json, group_id in activities:
            times = json.loads(times_json)
            if times:
                stats = self.calculate_time_statistics(times)
                cursor.execute('''
                    INSERT INTO statistics (
                        activity_id, group_id, count, min_val, max_val, median_val,
                        q1, q3, iqr, upper_fence, lower_fence, outlier_count,
                        non_outlier_count, overall_mean, non_outlier_mean,
                        total_activity_time, non_normalized_time, normalized_time
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ''', (activity_id, group_id, *stats.values()))
        
        # Calculate for groups
        cursor.execute('SELECT id, name FROM groups')
        groups = cursor.fetchall()
        
        for group_id, group_name in groups:
            cursor.execute('SELECT times FROM activities WHERE group_id = ?', (group_id,))
            group_activities = cursor.fetchall()
            
            all_times = []
            for times_json, in group_activities:
                times = json.loads(times_json)
                all_times.extend(times)
            
            if all_times:
                stats = self.calculate_time_statistics(all_times)
                cursor.execute('''
                    INSERT INTO statistics (
                        activity_id, group_id, count, min_val, max_val, median_val,
                        q1, q3, iqr, upper_fence, lower_fence, outlier_count,
                        non_outlier_count, overall_mean, non_outlier_mean,
                        total_activity_time, non_normalized_time, normalized_time
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ''', (None, group_id, *stats.values()))
        
        self.conn.commit()
        self.update_statistics_display()
        self.create_box_plots()
        messagebox.showinfo("Sucesso", "Estatísticas calculadas com sucesso!")
    
    def calculate_time_statistics(self, times):
        """Calculate comprehensive statistics for time data"""
        times = np.array(times)
        times = times[~np.isnan(times)]  # Remove NaN values
        
        if len(times) == 0:
            return {
                'count': 0, 'min_val': 0, 'max_val': 0, 'median_val': 0,
                'q1': 0, 'q3': 0, 'iqr': 0, 'upper_fence': 0, 'lower_fence': 0,
                'outlier_count': 0, 'non_outlier_count': 0, 'overall_mean': 0,
                'non_outlier_mean': 0, 'total_activity_time': 0,
                'non_normalized_time': 0, 'normalized_time': 0
            }
        
        # Basic statistics
        count = len(times)
        min_val = np.min(times)
        max_val = np.max(times)
        median_val = np.median(times)
        q1 = np.percentile(times, 25)
        q3 = np.percentile(times, 75)
        iqr = q3 - q1
        
        # Outlier detection
        lower_fence = q1 - 1.5 * iqr
        upper_fence = q3 + 1.5 * iqr
        outliers = times[(times < lower_fence) | (times > upper_fence)]
        non_outliers = times[(times >= lower_fence) & (times <= upper_fence)]
        
        outlier_count = len(outliers)
        non_outlier_count = len(non_outliers)
        
        # Means
        overall_mean = np.mean(times)
        non_outlier_mean = np.mean(non_outliers) if len(non_outliers) > 0 else 0
        
        # Time calculations
        total_activity_time = np.sum(times)
        non_normalized_time = total_activity_time  # in seconds
        normalized_time = total_activity_time / 60  # in minutes
        
        return {
            'count': count,
            'min_val': min_val,
            'max_val': max_val,
            'median_val': median_val,
            'q1': q1,
            'q3': q3,
            'iqr': iqr,
            'upper_fence': upper_fence,
            'lower_fence': lower_fence,
            'outlier_count': outlier_count,
            'non_outlier_count': non_outlier_count,
            'overall_mean': overall_mean,
            'non_outlier_mean': non_outlier_mean,
            'total_activity_time': total_activity_time,
            'non_normalized_time': non_normalized_time,
            'normalized_time': normalized_time
        }
    
    def update_statistics_display(self):
        """Update statistics display in treeview"""
        # Clear existing items
        for item in self.stats_tree.get_children():
            self.stats_tree.delete(item)
        
        cursor = self.conn.cursor()
        
        # Get statistics with activity and group names
        cursor.execute('''
            SELECT s.*, a.name as activity_name, g.name as group_name
            FROM statistics s
            LEFT JOIN activities a ON s.activity_id = a.id
            LEFT JOIN groups g ON s.group_id = g.id
            ORDER BY s.group_id, s.activity_id
        ''')
        
        statistics = cursor.fetchall()
        
        for stat in statistics:
            if stat[1]:  # activity_id exists
                tipo = "ATIVIDADE"
                nome = stat[19]  # activity_name
            else:
                tipo = "GRUPO"
                nome = stat[20]  # group_name
            
            self.stats_tree.insert('', tk.END, values=(
                tipo, nome, stat[3], f"{stat[14]:.2f}", f"{stat[6]:.2f}",
                f"{stat[7]:.2f}", f"{stat[8]:.2f}", f"{stat[9]:.2f}"
            ))
    
    def create_box_plots(self):
        """Create box plots for visualization"""
        # Clear existing plots
        for widget in self.plot_frame.winfo_children():
            widget.destroy()
        
        cursor = self.conn.cursor()
        cursor.execute('SELECT name, times FROM activities')
        activities = cursor.fetchall()
        
        if not activities:
            return
        
        # Create matplotlib figure
        fig, ax = plt.subplots(figsize=(12, 6))
        
        data = []
        labels = []
        
        for name, times_json in activities:
            times = json.loads(times_json)
            if times:
                data.append(times)
                labels.append(name[:20])  # Truncate long names
        
        if data:
            box_plot = ax.boxplot(data, labels=labels, patch_artist=True)
            
            # Customize colors
            colors = plt.cm.Set3(np.linspace(0, 1, len(data)))
            for patch, color in zip(box_plot['boxes'], colors):
                patch.set_facecolor(color)
            
            ax.set_title('Distribuição de Tempos por Atividade')
            ax.set_ylabel('Tempo (segundos)')
            ax.set_xlabel('Atividades')
            plt.xticks(rotation=45, ha='right')
            plt.tight_layout()
            
            # Embed plot in tkinter
            canvas = FigureCanvasTkAgg(fig, self.plot_frame)
            canvas.draw()
            canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
    
    def export_to_excel(self):
        """Export results to Excel file"""
        file_path = filedialog.asksaveasfilename(
            title="Salvar Arquivo Excel",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")]
        )
        
        if not file_path:
            return
        
        try:
            cursor = self.conn.cursor()
            
            # Get comprehensive data
            cursor.execute('''
                SELECT 
                    CASE WHEN s.activity_id IS NULL THEN 'GRUPO' ELSE 'ATIVIDADE' END as tipo,
                    COALESCE(a.name, g.name) as nome,
                    COALESCE(g.name, 'SEM GRUPO') as grupo,
                    s.count, s.min_val, s.max_val, s.median_val,
                    s.q1, s.q3, s.iqr, s.upper_fence, s.lower_fence,
                    s.outlier_count, s.non_outlier_count, s.overall_mean,
                    s.non_outlier_mean, s.total_activity_time,
                    s.non_normalized_time, s.normalized_time
                FROM statistics s
                LEFT JOIN activities a ON s.activity_id = a.id
                LEFT JOIN groups g ON s.group_id = g.id
                ORDER BY g.name, a.name
            ''')
            
            data = cursor.fetchall()
            
            # Create DataFrame
            columns = [
                'Tipo', 'Nome', 'Grupo', 'Contagem', 'Mínimo', 'Máximo', 'Mediana',
                'Q1', 'Q3', 'IQR', 'Limite Superior', 'Limite Inferior',
                'Outliers', 'Não-Outliers', 'Média Geral', 'Média Sem Outliers',
                'Tempo Total', 'Tempo Não Normalizado', 'Tempo Normalizado'
            ]
            
            df = pd.DataFrame(data, columns=columns)
            
            # Export to Excel
            with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name='Análise Estatística', index=False)
            
            self.export_status_label.config(text=f"Arquivo exportado com sucesso: {file_path}")
            messagebox.showinfo("Sucesso", f"Arquivo exportado com sucesso!\n{file_path}")
            
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao exportar arquivo: {str(e)}")
    
    def __del__(self):
        """Close database connection"""
        if hasattr(self, 'conn'):
            self.conn.close()

def main():
    root = tk.Tk()
    app = TimeStudyApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
