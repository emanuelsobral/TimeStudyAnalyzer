import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import sqlite3
import os
from pathlib import Path
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import seaborn as sns
import numpy as np
from difflib import SequenceMatcher

class TimeStudyAnalyzer:
    def __init__(self, root):
        self.root = root
        self.root.title("Plataforma de Análise de Estudo de Tempos - Desktop")
        self.root.geometry("1200x800")
        
        # Inicializar banco de dados
        self.init_database()
        
        # Variáveis de estado
        self.uploaded_files = []
        self.processed_data = None
        self.activity_column = None
        self.time_column = None
        self.unified_activities = {}
        self.activity_groups = {}
        
        # Criar interface
        self.create_interface()
        
    def init_database(self):
        """Inicializa o banco de dados SQLite"""
        self.db_path = "time_study.db"
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        # Criar tabelas
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS files (
                id INTEGER PRIMARY KEY,
                filename TEXT NOT NULL,
                upload_date TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                data TEXT NOT NULL
            )
        ''')
        
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS activities (
                id INTEGER PRIMARY KEY,
                name TEXT NOT NULL,
                unified_name TEXT,
                group_name TEXT,
                time_seconds REAL NOT NULL
            )
        ''')
        
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS groups (
                id INTEGER PRIMARY KEY,
                name TEXT NOT NULL,
                color TEXT
            )
        ''')
        
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS statistics (
                id INTEGER PRIMARY KEY,
                activity_name TEXT NOT NULL,
                group_name TEXT,
                q1 REAL,
                median REAL,
                q3 REAL,
                iqr REAL,
                outliers TEXT,
                mean_all REAL,
                mean_no_outliers REAL,
                total_time REAL
            )
        ''')
        
        conn.commit()
        conn.close()
        
    def create_interface(self):
        """Cria a interface principal com abas"""
        # Criar notebook para abas
        self.notebook = ttk.Notebook(self.root)
        self.notebook.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        
        # Configurar grid weights
        self.root.grid_rowconfigure(0, weight=1)
        self.root.grid_columnconfigure(0, weight=1)
        
        # Criar abas
        self.create_upload_tab()
        self.create_mapping_tab()
        self.create_unification_tab()
        self.create_grouping_tab()
        self.create_analysis_tab()
        self.create_export_tab()
        
    def create_upload_tab(self):
        """Aba de upload de arquivos"""
        upload_frame = ttk.Frame(self.notebook)
        self.notebook.add(upload_frame, text="1. Upload de Arquivos")
        
        # Configurar grid
        upload_frame.grid_rowconfigure(1, weight=1)
        upload_frame.grid_columnconfigure(0, weight=1)
        
        # Título
        title_label = ttk.Label(upload_frame, text="Upload de Arquivos Excel/CSV", font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, pady=10)
        
        # Frame para botões
        button_frame = ttk.Frame(upload_frame)
        button_frame.grid(row=1, column=0, pady=10)
        
        # Botão de upload
        upload_btn = ttk.Button(button_frame, text="Selecionar Arquivos", command=self.upload_files)
        upload_btn.grid(row=0, column=0, padx=5)
        
        # Botão de limpar
        clear_btn = ttk.Button(button_frame, text="Limpar Lista", command=self.clear_files)
        clear_btn.grid(row=0, column=1, padx=5)
        
        # Lista de arquivos
        list_frame = ttk.Frame(upload_frame)
        list_frame.grid(row=2, column=0, sticky="nsew", pady=10)
        list_frame.grid_rowconfigure(0, weight=1)
        list_frame.grid_columnconfigure(0, weight=1)
        
        self.file_listbox = tk.Listbox(list_frame)
        self.file_listbox.grid(row=0, column=0, sticky="nsew")
        
        # Scrollbar para lista
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.file_listbox.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.file_listbox.configure(yscrollcommand=scrollbar.set)
        
        # Preview dos dados
        preview_label = ttk.Label(upload_frame, text="Preview dos Dados:")
        preview_label.grid(row=3, column=0, sticky="w", pady=(10, 5))
        
        preview_frame = ttk.Frame(upload_frame)
        preview_frame.grid(row=4, column=0, sticky="nsew", pady=5)
        preview_frame.grid_rowconfigure(0, weight=1)
        preview_frame.grid_columnconfigure(0, weight=1)
        
        # Treeview para preview
        self.preview_tree = ttk.Treeview(preview_frame)
        self.preview_tree.grid(row=0, column=0, sticky="nsew")
        
        # Scrollbars para preview
        h_scroll = ttk.Scrollbar(preview_frame, orient="horizontal", command=self.preview_tree.xview)
        h_scroll.grid(row=1, column=0, sticky="ew")
        self.preview_tree.configure(xscrollcommand=h_scroll.set)
        
        v_scroll = ttk.Scrollbar(preview_frame, orient="vertical", command=self.preview_tree.yview)
        v_scroll.grid(row=0, column=1, sticky="ns")
        self.preview_tree.configure(yscrollcommand=v_scroll.set)
        
        # Configurar weights para upload_frame
        upload_frame.grid_rowconfigure(2, weight=1)
        upload_frame.grid_rowconfigure(4, weight=2)
        
    def create_mapping_tab(self):
        """Aba de mapeamento de colunas"""
        mapping_frame = ttk.Frame(self.notebook)
        self.notebook.add(mapping_frame, text="2. Mapeamento de Colunas")
        
        # Configurar grid
        mapping_frame.grid_rowconfigure(2, weight=1)
        mapping_frame.grid_columnconfigure(0, weight=1)
        
        # Título
        title_label = ttk.Label(mapping_frame, text="Mapeamento de Colunas", font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, pady=10)
        
        # Frame para seleção
        selection_frame = ttk.Frame(mapping_frame)
        selection_frame.grid(row=1, column=0, pady=10)
        
        # Seleção de coluna de atividade
        ttk.Label(selection_frame, text="Coluna de Atividade:").grid(row=0, column=0, padx=5, sticky="w")
        self.activity_combo = ttk.Combobox(selection_frame, state="readonly")
        self.activity_combo.grid(row=0, column=1, padx=5)
        
        # Seleção de coluna de tempo
        ttk.Label(selection_frame, text="Coluna de Tempo:").grid(row=0, column=2, padx=5, sticky="w")
        self.time_combo = ttk.Combobox(selection_frame, state="readonly")
        self.time_combo.grid(row=0, column=3, padx=5)
        
        # Botão de processar
        process_btn = ttk.Button(selection_frame, text="Processar Dados", command=self.process_data)
        process_btn.grid(row=0, column=4, padx=10)
        
        # Botão de debug
        debug_btn = ttk.Button(selection_frame, text="Debug Dados", command=self.debug_data)
        debug_btn.grid(row=0, column=5, padx=5)
        
        # Preview dos dados processados
        preview_label = ttk.Label(mapping_frame, text="Dados Processados:")
        preview_label.grid(row=2, column=0, sticky="w", pady=(10, 5))
        
        processed_frame = ttk.Frame(mapping_frame)
        processed_frame.grid(row=3, column=0, sticky="nsew", pady=5)
        processed_frame.grid_rowconfigure(0, weight=1)
        processed_frame.grid_columnconfigure(0, weight=1)
        
        self.processed_tree = ttk.Treeview(processed_frame, columns=("Atividade", "Tempo"), show="headings")
        self.processed_tree.heading("Atividade", text="Atividade")
        self.processed_tree.heading("Tempo", text="Tempo (segundos)")
        self.processed_tree.grid(row=0, column=0, sticky="nsew")
        
        # Scrollbar
        processed_scroll = ttk.Scrollbar(processed_frame, orient="vertical", command=self.processed_tree.yview)
        processed_scroll.grid(row=0, column=1, sticky="ns")
        self.processed_tree.configure(yscrollcommand=processed_scroll.set)
        
        # Configurar weights
        mapping_frame.grid_rowconfigure(3, weight=1)
        
    def create_unification_tab(self):
        """Aba de unificação de atividades"""
        unification_frame = ttk.Frame(self.notebook)
        self.notebook.add(unification_frame, text="3. Unificação de Atividades")
        
        # Configurar grid
        unification_frame.grid_rowconfigure(2, weight=1)
        unification_frame.grid_columnconfigure(0, weight=1)
        
        # Título
        title_label = ttk.Label(unification_frame, text="Unificação de Atividades Similares", font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, pady=10)
        
        # Frame para botões
        buttons_frame = ttk.Frame(unification_frame)
        buttons_frame.grid(row=1, column=0, pady=5)
        
        # Botão para detectar similaridades
        detect_btn = ttk.Button(buttons_frame, text="Detectar Atividades Similares", command=self.detect_similarities)
        detect_btn.grid(row=0, column=0, padx=5)
        
        # Botão para atualizar após unificações
        refresh_btn = ttk.Button(buttons_frame, text="Atualizar Lista", command=self.detect_similarities)
        refresh_btn.grid(row=0, column=1, padx=5)
        
        # Frame para sugestões
        suggestions_frame = ttk.Frame(unification_frame)
        suggestions_frame.grid(row=2, column=0, sticky="nsew", pady=10)
        suggestions_frame.grid_rowconfigure(0, weight=1)
        suggestions_frame.grid_columnconfigure(0, weight=1)
        
        self.similarity_tree = ttk.Treeview(suggestions_frame, columns=("Atividades", "Similaridade", "Ação"), show="headings")
        self.similarity_tree.heading("Atividades", text="Atividades Similares")
        self.similarity_tree.heading("Similaridade", text="% Similaridade")
        self.similarity_tree.heading("Ação", text="Ação")
        self.similarity_tree.grid(row=0, column=0, sticky="nsew")
        
        # Scrollbar
        similarity_scroll = ttk.Scrollbar(suggestions_frame, orient="vertical", command=self.similarity_tree.yview)
        similarity_scroll.grid(row=0, column=1, sticky="ns")
        self.similarity_tree.configure(yscrollcommand=similarity_scroll.set)
        
        # Frame para botões de ação
        action_frame = ttk.Frame(unification_frame)
        action_frame.grid(row=3, column=0, pady=5)
        
        unify_btn = ttk.Button(action_frame, text="Unificar Selecionadas", command=self.unify_activities)
        unify_btn.grid(row=0, column=0, padx=5)
        
        skip_btn = ttk.Button(action_frame, text="Pular Selecionadas", command=self.skip_activities)
        skip_btn.grid(row=0, column=1, padx=5)
        
    def create_grouping_tab(self):
        """Aba de agrupamento de atividades"""
        grouping_frame = ttk.Frame(self.notebook)
        self.notebook.add(grouping_frame, text="4. Agrupamento")
        
        # Configurar grid
        grouping_frame.grid_rowconfigure(1, weight=1)
        grouping_frame.grid_columnconfigure(0, weight=1)
        grouping_frame.grid_columnconfigure(1, weight=1)
        
        # Título
        title_label = ttk.Label(grouping_frame, text="Agrupamento de Atividades", font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=2, pady=10)
        
        # Frame esquerdo - Atividades disponíveis
        left_frame = ttk.LabelFrame(grouping_frame, text="Atividades Disponíveis")
        left_frame.grid(row=1, column=0, sticky="nsew", padx=(0, 5), pady=5)
        left_frame.grid_rowconfigure(0, weight=1)
        left_frame.grid_columnconfigure(0, weight=1)
        
        self.available_listbox = tk.Listbox(left_frame, selectmode=tk.MULTIPLE)
        self.available_listbox.grid(row=0, column=0, sticky="nsew", padx=5, pady=5)
        
        # Frame direito - Grupos
        right_frame = ttk.LabelFrame(grouping_frame, text="Grupos")
        right_frame.grid(row=1, column=1, sticky="nsew", padx=(5, 0), pady=5)
        right_frame.grid_rowconfigure(1, weight=1)
        right_frame.grid_columnconfigure(0, weight=1)
        
        # Botões para grupos
        group_btn_frame = ttk.Frame(right_frame)
        group_btn_frame.grid(row=0, column=0, sticky="ew", padx=5, pady=5)
        
        new_group_btn = ttk.Button(group_btn_frame, text="Novo Grupo", command=self.create_new_group)
        new_group_btn.grid(row=0, column=0, padx=2)
        
        add_to_group_btn = ttk.Button(group_btn_frame, text="Adicionar ao Grupo", command=self.add_to_group)
        add_to_group_btn.grid(row=0, column=1, padx=2)
        
        # Tree para grupos
        self.group_tree = ttk.Treeview(right_frame)
        self.group_tree.grid(row=1, column=0, sticky="nsew", padx=5, pady=5)
        
        group_scroll = ttk.Scrollbar(right_frame, orient="vertical", command=self.group_tree.yview)
        group_scroll.grid(row=1, column=1, sticky="ns")
        self.group_tree.configure(yscrollcommand=group_scroll.set)
        
    def create_analysis_tab(self):
        """Aba de análise estatística"""
        analysis_frame = ttk.Frame(self.notebook)
        self.notebook.add(analysis_frame, text="5. Análise Estatística")
        
        # Configurar grid
        analysis_frame.grid_rowconfigure(1, weight=1)
        analysis_frame.grid_columnconfigure(0, weight=1)
        
        # Frame do cabeçalho para título e botão
        header_frame = ttk.Frame(analysis_frame)
        header_frame.grid(row=0, column=0, sticky="ew", padx=10, pady=(10, 0))
        header_frame.grid_columnconfigure(0, weight=1)
        
        # Título
        title_label = ttk.Label(header_frame, text="Análise Estatística", font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, sticky="w")
        
        # Botão de análise
        analyze_btn = ttk.Button(header_frame, text="Executar Análise", command=self.perform_analysis)
        analyze_btn.grid(row=0, column=1, sticky="e")
        
        # Frame da Tabela de resultados
        table_frame = ttk.LabelFrame(analysis_frame, text="Resultados Estatísticos")
        table_frame.grid(row=1, column=0, sticky="nsew", padx=10, pady=5)
        table_frame.grid_rowconfigure(0, weight=1)
        table_frame.grid_columnconfigure(0, weight=1)
        
        # Definir as colunas da tabela com a nova ordem
        cols = (
            "n", "std_dev", "min", "max", "median", "q1", "q3", "iqr", 
            "lower_fence", "upper_fence", "outlier_count", "non_outlier_count",
            "mean_all", "mean_no_outliers", "time_non_norm", "time_norm"
        )
        
        self.results_tree = ttk.Treeview(table_frame, columns=cols, show="tree headings")
        
        # Definir os cabeçalhos das colunas e larguras com a nova ordem
        col_headings = {
            "n": ("N", 40), "std_dev": ("Desvio Padrão", 90), "min": ("Mínimo", 80), 
            "max": ("Máximo", 80), "median": ("Mediana", 80), "q1": ("Q1", 80), 
            "q3": ("Q3", 80), "iqr": ("IQR", 80), "lower_fence": ("Limite Inf.", 90),
            "upper_fence": ("Limite Sup.", 90), "outlier_count": ("Outliers", 60),
            "non_outlier_count": ("Dentro Lim.", 70), "mean_all": ("Média Geral", 80),
            "mean_no_outliers": ("Média s/ Out.", 80),
            "time_non_norm": ("T. Não Norm. (min)", 120),
            "time_norm": ("T. Norm. (min)", 120)
        }
        
        self.results_tree.heading("#0", text="Atividade/Grupo")
        self.results_tree.column("#0", width=200, stretch=tk.NO)
        
        for col, (heading, width) in col_headings.items():
            self.results_tree.heading(col, text=heading)
            self.results_tree.column(col, width=width, anchor="center")

        self.results_tree.grid(row=0, column=0, sticky="nsew")
        
        # Scrollbar Vertical
        results_v_scroll = ttk.Scrollbar(table_frame, orient="vertical", command=self.results_tree.yview)
        results_v_scroll.grid(row=0, column=1, sticky="ns")
        self.results_tree.configure(yscrollcommand=results_v_scroll.set)
        
        # Scrollbar Horizontal
        results_h_scroll = ttk.Scrollbar(table_frame, orient="horizontal", command=self.results_tree.xview)
        results_h_scroll.grid(row=1, column=0, sticky="ew")
        self.results_tree.configure(xscrollcommand=results_h_scroll.set)

    def create_export_tab(self):
        """Aba de exportação"""
        export_frame = ttk.Frame(self.notebook)
        self.notebook.add(export_frame, text="6. Exportação")
        
        # Configurar grid
        export_frame.grid_rowconfigure(1, weight=1)
        export_frame.grid_columnconfigure(0, weight=1)
        
        # Título
        title_label = ttk.Label(export_frame, text="Exportação de Resultados", font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, pady=10)
        
        # Frame para opções
        options_frame = ttk.Frame(export_frame)
        options_frame.grid(row=1, column=0, pady=20)
        
        # Botão de exportar Excel
        export_excel_btn = ttk.Button(options_frame, text="Exportar para Excel", command=self.export_to_excel)
        export_excel_btn.grid(row=0, column=0, padx=10, pady=10)
        
        # Botão de exportar CSV
        export_csv_btn = ttk.Button(options_frame, text="Exportar para CSV (Simples)", command=self.export_to_csv)
        export_csv_btn.grid(row=0, column=1, padx=10, pady=10)
        
        # Label de status
        self.export_status = ttk.Label(export_frame, text="Pronto para exportar")
        self.export_status.grid(row=2, column=0, pady=10)
        
    def upload_files(self):
        """Função para upload de arquivos"""
        filetypes = [
            ("Arquivos Excel", "*.xlsx"),
            ("Arquivos CSV", "*.csv"),
            ("Todos os arquivos", "*.*")
        ]
        
        files = filedialog.askopenfilenames(
            title="Selecionar arquivos",
            filetypes=filetypes
        )
        
        for file_path in files:
            if file_path not in self.uploaded_files:
                self.uploaded_files.append(file_path)
                filename = os.path.basename(file_path)
                self.file_listbox.insert(tk.END, filename)
                
        self.update_preview()
        self.update_column_combos()
        
    def clear_files(self):
        """Limpar lista de arquivos"""
        self.uploaded_files.clear()
        self.file_listbox.delete(0, tk.END)
        self.clear_preview()
        
    def update_preview(self):
        """Atualizar preview dos dados"""
        if not self.uploaded_files:
            self.clear_preview()
            return
            
        try:
            # Ler primeiro arquivo para preview
            file_path = self.uploaded_files[0]
            if file_path.endswith('.xlsx'):
                df = pd.read_excel(file_path)
            else:
                df = pd.read_csv(file_path)
                
            # Limpar preview anterior
            self.clear_preview()
            
            # Configurar colunas
            self.preview_tree["columns"] = list(df.columns)
            self.preview_tree["show"] = "headings"
            
            for col in df.columns:
                self.preview_tree.heading(col, text=col)
                self.preview_tree.column(col, width=100)
                
            # Adicionar dados (primeiras 10 linhas)
            for index, row in df.head(10).iterrows():
                self.preview_tree.insert("", tk.END, values=list(row))
                
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao ler arquivo: {str(e)}")
            
    def clear_preview(self):
        """Limpar preview"""
        for item in self.preview_tree.get_children():
            self.preview_tree.delete(item)
            
    def update_column_combos(self):
        """Atualizar comboboxes de colunas"""
        if not self.uploaded_files:
            return
            
        try:
            file_path = self.uploaded_files[0]
            if file_path.endswith('.xlsx'):
                df = pd.read_excel(file_path)
            else:
                df = pd.read_csv(file_path)
                
            columns = list(df.columns)
            self.activity_combo['values'] = columns
            self.time_combo['values'] = columns
            
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao ler colunas: {str(e)}")

    def process_data(self):
        """Processar dados dos arquivos, garantindo que células vazias sejam ignoradas."""
        if not self.uploaded_files:
            messagebox.showwarning("Aviso", "Nenhum arquivo selecionado")
            return
            
        if not self.activity_combo.get() or not self.time_combo.get():
            messagebox.showwarning("Aviso", "Selecione as colunas de atividade e tempo")
            return
            
        try:
            all_data = []
            total_rows_read = 0
            files_processed = 0
            
            activity_col = self.activity_combo.get()
            time_col = self.time_combo.get()
            
            for file_path in self.uploaded_files:
                try:
                    # Ler arquivo
                    if file_path.endswith('.xlsx'):
                        df = pd.read_excel(file_path)
                    else:
                        try:
                            df = pd.read_csv(file_path, encoding='utf-8')
                        except UnicodeDecodeError:
                            try:
                                df = pd.read_csv(file_path, encoding='latin1')
                            except UnicodeDecodeError:
                                df = pd.read_csv(file_path, encoding='cp1252')
                    
                    total_rows_read += len(df)
                    
                    if activity_col in df.columns and time_col in df.columns:
                        # 1. Selecionar apenas as colunas de interesse
                        data = df[[activity_col, time_col]].copy()
                        data.columns = ['Atividade', 'Tempo']
                        
                        # --- INÍCIO DA LÓGICA DE LIMPEZA ---
                        # 2. Remover qualquer linha onde a Atividade ou o Tempo sejam nulos (célula vazia)
                        #    Isso garante que linhas com informações incompletas sejam descartadas.
                        data.dropna(subset=['Atividade', 'Tempo'], how='any', inplace=True)

                        # 3. Padronizar a coluna de atividade como texto e remover espaços
                        data['Atividade'] = data['Atividade'].astype(str).str.strip()

                        # 4. Remover linhas que, após a conversão, resultaram em strings vazias
                        data = data[data['Atividade'] != '']
                        # --- FIM DA LÓGICA DE LIMPEZA ---
                        
                        if not data.empty:
                            all_data.append(data)
                            files_processed += 1
                    else:
                        print(f"Colunas não encontradas no arquivo {os.path.basename(file_path)}")

                except Exception as e:
                    print(f"Erro ao processar arquivo {os.path.basename(file_path)}: {str(e)}")
                    continue
            
            if not all_data:
                messagebox.showerror("Erro", "Nenhum dado válido pôde ser extraído dos arquivos. Verifique se as colunas selecionadas estão corretas e contêm dados.")
                return

            # Concatenar todos os dados limpos
            self.processed_data = pd.concat(all_data, ignore_index=True)
            original_count = len(self.processed_data)
            
            # Função para converter a coluna de tempo para segundos
            def convert_time(time_val):
                if pd.isna(time_val): return None
                if isinstance(time_val, (int, float)): return float(time_val)
                
                time_str = str(time_val).strip()
                if ':' in time_str:
                    try:
                        parts = time_str.split(':')
                        total_seconds = 0
                        if len(parts) == 3: # HH:MM:SS
                            total_seconds = float(parts[0])*3600 + float(parts[1])*60 + float(parts[2])
                        elif len(parts) == 2: # MM:SS
                            total_seconds = float(parts[0])*60 + float(parts[1])
                        return total_seconds
                    except (ValueError, IndexError):
                        return None
                try:
                    return float(time_str.replace(',', '.'))
                except ValueError:
                    return None

            # Aplicar conversão de tempo e remover falhas
            self.processed_data['Tempo'] = self.processed_data['Tempo'].apply(convert_time)
            self.processed_data.dropna(subset=['Tempo'], inplace=True)
            
            # Remover tempos negativos ou zero
            self.processed_data = self.processed_data[self.processed_data['Tempo'] > 0]
            
            final_count = len(self.processed_data)
            removed_count = original_count - final_count

            if final_count > 0:
                self.update_processed_preview()
                self.update_available_activities()
                messagebox.showinfo("Sucesso", 
                    f"Dados processados com sucesso!\n\n"
                    f"• {files_processed} arquivo(s) processado(s).\n"
                    f"• {total_rows_read} linha(s) lida(s) no total.\n"
                    f"• {final_count} registro(s) válido(s) para análise.\n"
                    f"• {removed_count} registro(s) removido(s) por tempo inválido.")
            else:
                 messagebox.showerror("Erro", "Nenhum dado válido encontrado após a limpeza e conversão. Verifique o formato dos dados nas colunas de tempo.")

        except Exception as e:
            messagebox.showerror("Erro", f"Erro geral ao processar dados: {str(e)}")
            print(f"Erro detalhado: {str(e)}")

    def update_processed_preview(self):
        """Atualizar preview dos dados processados"""
        for item in self.processed_tree.get_children():
            self.processed_tree.delete(item)
            
        if self.processed_data is not None:
            for index, row in self.processed_data.head(20).iterrows():
                self.processed_tree.insert("", tk.END, values=(row['Atividade'], f"{row['Tempo']:.2f}"))
                
    def update_available_activities(self):
        """Atualizar lista de atividades disponíveis"""
        self.available_listbox.delete(0, tk.END)
        
        if self.processed_data is not None:
            unique_activities = self.processed_data['Atividade'].unique()
            
            # Atividades que já estão em grupos
            grouped_activities = set()
            for group_data in self.activity_groups.values():
                grouped_activities.update(group_data['activities'])
            
            # Adicionar apenas atividades que não estão em grupos
            for activity in sorted(unique_activities):
                if activity not in grouped_activities:
                    self.available_listbox.insert(tk.END, activity)
                    
    def detect_similarities(self):
        """Detectar atividades similares"""
        if self.processed_data is None:
            messagebox.showwarning("Aviso", "Processe os dados primeiro")
            return
            
        # Limpar árvore de similaridades
        for item in self.similarity_tree.get_children():
            self.similarity_tree.delete(item)
            
        unique_activities = list(self.processed_data['Atividade'].unique())
        similarities = []
        
        for i, activity1 in enumerate(unique_activities):
            for j, activity2 in enumerate(unique_activities[i+1:], i+1):
                similarity = self.calculate_similarity(activity1, activity2)
                if similarity > 0.7:  # Threshold de 70%
                    similarities.append((activity1, activity2, similarity))
                    
        # Ordenar por similaridade (maior primeiro)
        similarities.sort(key=lambda x: x[2], reverse=True)
        
        # Adicionar similaridades à árvore
        for activity1, activity2, similarity in similarities:
            self.similarity_tree.insert("", tk.END, values=(
                f"{activity1} ↔ {activity2}",
                f"{similarity:.1%}",
                "Pendente"
            ))
            
        if not similarities:
            messagebox.showinfo("Info", "Nenhuma atividade similar encontrada")
            
    def calculate_similarity(self, str1, str2):
        """Calcular similaridade entre duas strings"""
        return SequenceMatcher(None, str1.lower(), str2.lower()).ratio()
        
    def unify_activities(self):
        """Unificar atividades selecionadas"""
        selected_items = self.similarity_tree.selection()
        if not selected_items:
            messagebox.showwarning("Aviso", "Selecione atividades para unificar")
            return
            
        try:
            for item in selected_items:
                # Obter dados do item selecionado
                item_data = self.similarity_tree.item(item)
                activities_text = item_data['values'][0]
                
                # Extrair as duas atividades
                if '↔' in activities_text:
                    activity1, activity2 = activities_text.split(' ↔ ')
                    
                    # Perguntar qual nome usar para unificação
                    choice_window = tk.Toplevel(self.root)
                    choice_window.title("Escolher Nome da Atividade")
                    choice_window.geometry("400x200")
                    choice_window.transient(self.root)
                    choice_window.grab_set()
                    
                    ttk.Label(choice_window, text="Escolha o nome para a atividade unificada:").pack(pady=10)
                    
                    var = tk.StringVar(value=activity1)
                    
                    ttk.Radiobutton(choice_window, text=activity1, variable=var, value=activity1).pack(pady=5)
                    ttk.Radiobutton(choice_window, text=activity2, variable=var, value=activity2).pack(pady=5)
                    
                    # Frame para entrada customizada
                    custom_frame = ttk.Frame(choice_window)
                    custom_frame.pack(pady=5)
                    
                    ttk.Radiobutton(custom_frame, text="Nome personalizado:", variable=var, value="custom").pack(side=tk.LEFT)
                    custom_entry = ttk.Entry(custom_frame, width=20)
                    custom_entry.pack(side=tk.LEFT, padx=5)
                    
                    def confirm_unification():
                        chosen_name = var.get()
                        if chosen_name == "custom":
                            chosen_name = custom_entry.get().strip()
                            if not chosen_name:
                                messagebox.showerror("Erro", "Digite um nome personalizado")
                                return
                        
                        # Realizar unificação
                        self.processed_data.loc[self.processed_data['Atividade'] == activity2, 'Atividade'] = chosen_name
                        if activity1 != chosen_name:
                            self.processed_data.loc[self.processed_data['Atividade'] == activity1, 'Atividade'] = chosen_name
                        
                        # Armazenar unificação
                        self.unified_activities[activity1] = chosen_name
                        self.unified_activities[activity2] = chosen_name
                        
                        # Atualizar status do item
                        self.similarity_tree.set(item, "Ação", "Unificada")
                        
                        choice_window.destroy()
                        
                        # Atualizar interfaces
                        self.update_processed_preview()
                        self.update_available_activities()
                    
                    ttk.Button(choice_window, text="Confirmar", command=confirm_unification).pack(pady=10)
                    
                    # Centralizar janela
                    choice_window.update_idletasks()
                    x = (self.root.winfo_screenwidth() // 2) - (choice_window.winfo_width() // 2)
                    y = (self.root.winfo_screenheight() // 2) - (choice_window.winfo_height() // 2)
                    choice_window.geometry(f"+{x}+{y}")
                    
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao unificar atividades: {str(e)}")
    
    def skip_activities(self):
        """Pular atividades selecionadas"""
        selected_items = self.similarity_tree.selection()
        if not selected_items:
            messagebox.showwarning("Aviso", "Selecione atividades para pular")
            return
            
        try:
            for item in selected_items:
                # Atualizar status do item
                self.similarity_tree.set(item, "Ação", "Pulada")
                
            messagebox.showinfo("Sucesso", f"{len(selected_items)} sugestão(ões) de unificação pulada(s)")
            
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao pular atividades: {str(e)}")
        
    def create_new_group(self):
        """Criar novo grupo"""
        # Janela para criar novo grupo
        group_window = tk.Toplevel(self.root)
        group_window.title("Criar Novo Grupo")
        group_window.geometry("300x150")
        group_window.transient(self.root)
        group_window.grab_set()
        
        # Nome do grupo
        ttk.Label(group_window, text="Nome do Grupo:").pack(pady=10)
        name_entry = ttk.Entry(group_window, width=30)
        name_entry.pack(pady=5)
        name_entry.focus()
        
        # Cor do grupo
        ttk.Label(group_window, text="Cor do Grupo:").pack(pady=(10, 5))
        color_var = tk.StringVar(value="#3B82F6")
        color_frame = ttk.Frame(group_window)
        color_frame.pack(pady=5)
        
        colors = ["#3B82F6", "#EF4444", "#10B981", "#F59E0B", "#8B5CF6", "#F97316"]
        for i, color in enumerate(colors):
            btn = tk.Button(color_frame, bg=color, width=3, height=1,
                            command=lambda c=color: color_var.set(c))
            btn.grid(row=0, column=i, padx=2)
        
        def confirm_group():
            name = name_entry.get().strip()
            if not name:
                messagebox.showerror("Erro", "Digite um nome para o grupo")
                return
            
            if name in self.activity_groups:
                messagebox.showerror("Erro", "Já existe um grupo com este nome")
                return
            
            # Criar grupo
            self.activity_groups[name] = {
                'color': color_var.get(),
                'activities': []
            }
            
            # Atualizar interface
            self.update_group_tree()
            group_window.destroy()
            messagebox.showinfo("Sucesso", f"Grupo '{name}' criado com sucesso!")
        
        ttk.Button(group_window, text="Criar", command=confirm_group).pack(pady=15)
        
        # Centralizar janela
        group_window.update_idletasks()
        x = (self.root.winfo_screenwidth() // 2) - (group_window.winfo_width() // 2)
        y = (self.root.winfo_screenheight() // 2) - (group_window.winfo_height() // 2)
        group_window.geometry(f"+{x}+{y}")
        
    def add_to_group(self):
        """Adicionar atividades ao grupo"""
        # Verificar se há atividades selecionadas
        selected_indices = self.available_listbox.curselection()
        if not selected_indices:
            messagebox.showwarning("Aviso", "Selecione atividades para adicionar ao grupo")
            return
        
        # Verificar se há grupos criados
        if not self.activity_groups:
            messagebox.showwarning("Aviso", "Crie um grupo primeiro")
            return
        
        # Janela para selecionar grupo
        select_window = tk.Toplevel(self.root)
        select_window.title("Selecionar Grupo")
        select_window.geometry("250x200")
        select_window.transient(self.root)
        select_window.grab_set()
        
        ttk.Label(select_window, text="Selecione o grupo:").pack(pady=10)
        
        # Lista de grupos
        group_listbox = tk.Listbox(select_window)
        group_listbox.pack(pady=5, padx=10, fill=tk.BOTH, expand=True)
        
        for group_name in self.activity_groups.keys():
            group_listbox.insert(tk.END, group_name)
        
        def confirm_addition():
            group_selection = group_listbox.curselection()
            if not group_selection:
                messagebox.showerror("Erro", "Selecione um grupo")
                return
            
            group_name = group_listbox.get(group_selection[0])
            selected_activities = [self.available_listbox.get(i) for i in selected_indices]
            
            # Adicionar atividades ao grupo
            for activity in selected_activities:
                if activity not in self.activity_groups[group_name]['activities']:
                    self.activity_groups[group_name]['activities'].append(activity)
            
            # Atualizar interface
            self.update_group_tree()
            self.update_available_activities()
            select_window.destroy()
            
            messagebox.showinfo("Sucesso", 
                f"{len(selected_activities)} atividade(s) adicionada(s) ao grupo '{group_name}'")
        
        ttk.Button(select_window, text="Adicionar", command=confirm_addition).pack(pady=10)
        
        # Centralizar janela
        select_window.update_idletasks()
        x = (self.root.winfo_screenwidth() // 2) - (select_window.winfo_width() // 2)
        y = (self.root.winfo_screenheight() // 2) - (select_window.winfo_height() // 2)
        select_window.geometry(f"+{x}+{y}")
    
    def update_group_tree(self):
        """Atualizar árvore de grupos"""
        # Limpar árvore
        for item in self.group_tree.get_children():
            self.group_tree.delete(item)
        
        # Adicionar grupos e suas atividades
        for group_name, group_data in self.activity_groups.items():
            # Adicionar grupo principal
            group_item = self.group_tree.insert("", tk.END, text=f"📁 {group_name}", 
                                                values=(), open=True)
            
            # Adicionar atividades do grupo
            for activity in group_data['activities']:
                self.group_tree.insert(group_item, tk.END, text=f"  📊 {activity}", values=())
        
        # Atualizar lista de atividades disponíveis (remover as que já estão em grupos)
        self.update_available_activities()

    def format_seconds_to_hms(self, seconds):
        """Converte segundos para o formato HH:MM:SS, lidando com valores negativos."""
        if pd.isna(seconds):
            return "N/A"
        try:
            seconds = float(seconds)
            
            # Lida com o sinal
            sign = "-" if seconds < 0 else ""
            seconds = abs(seconds)
            
            hours, remainder = divmod(seconds, 3600)
            minutes, seconds_part = divmod(remainder, 60)
            
            return f"{sign}{int(hours):02d}:{int(minutes):02d}:{int(seconds_part):02d}"
            
        except (ValueError, TypeError):
            return "N/A"

    def _calculate_metrics(self, times_series):
        """Função auxiliar para calcular todas as métricas para uma série de tempos."""
        if times_series.empty:
            return None

        # Métricas básicas
        n = len(times_series)
        std_dev = times_series.std()
        min_time = times_series.min()
        max_time = times_series.max()
        q1 = times_series.quantile(0.25)
        median = times_series.median()
        q3 = times_series.quantile(0.75)
        iqr = q3 - q1
        mean_all = times_series.mean()

        # Outliers
        lower_bound = q1 - 1.5 * iqr
        upper_bound = q3 + 1.5 * iqr
        outliers = times_series[(times_series < lower_bound) | (times_series > upper_bound)]
        clean_times = times_series[(times_series >= lower_bound) & (times_series <= upper_bound)]
        
        outlier_count = len(outliers)
        non_outlier_count = len(clean_times)
        mean_no_outliers = clean_times.mean() if non_outlier_count > 0 else 0

        # Tempos em minutos
        time_non_norm = mean_all / 60
        time_norm = mean_no_outliers / 60
        
        # Formatação para exibição na nova ordem
        values = (
            n,
            self.format_seconds_to_hms(std_dev),
            self.format_seconds_to_hms(min_time),
            self.format_seconds_to_hms(max_time),
            self.format_seconds_to_hms(median),
            self.format_seconds_to_hms(q1),
            self.format_seconds_to_hms(q3),
            self.format_seconds_to_hms(iqr),
            self.format_seconds_to_hms(lower_bound),
            self.format_seconds_to_hms(upper_bound),
            outlier_count,
            non_outlier_count,
            self.format_seconds_to_hms(mean_all),
            self.format_seconds_to_hms(mean_no_outliers),
            f"{time_non_norm:.2f} min",
            f"{time_norm:.2f} min"
        )
        return values

    def perform_analysis(self):
        """Executar análise estatística."""
        if self.processed_data is None or self.processed_data.empty:
            messagebox.showwarning("Aviso", "Não há dados válidos para analisar. Processe os arquivos primeiro.")
            return

        try:
            # Limpar resultados anteriores
            for item in self.results_tree.get_children():
                self.results_tree.delete(item)

            # Analisar grupos (se existirem)
            if self.activity_groups:
                for group_name, group_data in self.activity_groups.items():
                    if group_data['activities']:
                        group_times = self.processed_data[self.processed_data['Atividade'].isin(group_data['activities'])]['Tempo']
                        metrics = self._calculate_metrics(group_times)
                        if metrics:
                            group_item = self.results_tree.insert("", tk.END, text=f"📁 {group_name}", values=metrics, open=True)
                            # Analisar atividades individuais do grupo
                            for activity in group_data['activities']:
                                activity_times = self.processed_data[self.processed_data['Atividade'] == activity]['Tempo']
                                activity_metrics = self._calculate_metrics(activity_times)
                                if activity_metrics:
                                    self.results_tree.insert(group_item, tk.END, text=f"  📊 {activity}", values=activity_metrics)

            # Analisar atividades não agrupadas
            grouped_activities = {act for group in self.activity_groups.values() for act in group['activities']}
            ungrouped_df = self.processed_data[~self.processed_data['Atividade'].isin(grouped_activities)]

            if not ungrouped_df.empty:
                # Criar um nó pai para atividades não agrupadas, se houver grupos.
                # Se não houver grupos, as atividades são listadas na raiz.
                parent_item = ""
                if self.activity_groups:
                        parent_item = self.results_tree.insert("", tk.END, text="📋 Atividades Não Agrupadas", open=True)

                for activity, times in ungrouped_df.groupby('Atividade')['Tempo']:
                    metrics = self._calculate_metrics(times)
                    if metrics:
                        self.results_tree.insert(parent_item, tk.END, text=f"📊 {activity}", values=metrics)

            messagebox.showinfo("Análise Concluída", "A análise estatística foi concluída com sucesso!")

        except Exception as e:
            import traceback
            error_details = traceback.format_exc()
            print(f"Erro detalhado na análise: {error_details}")
            messagebox.showerror("Erro", f"Erro na análise estatística:\n{str(e)}\n\nVerifique o console para mais detalhes.")
            
    def export_to_excel(self):
        """Exportar resultados para Excel no formato customizado."""
        if self.processed_data is None:
            messagebox.showwarning("Aviso", "Nenhum dado para exportar. Execute a análise primeiro.")
            return

        if not self.results_tree.get_children():
            messagebox.showwarning("Aviso", "Nenhuma análise encontrada para exportar. Execute a análise primeiro.")
            return

        try:
            filename = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Arquivos Excel", "*.xlsx"), ("Todos os Arquivos", "*.*")]
            )
            if not filename:
                return

            # --- Parte 1: Preparar dados brutos pivotados ---
            # Mapeamento de atividade para grupo
            activity_to_group = {activity: group_name
                                 for group_name, data in self.activity_groups.items()
                                 for activity in data['activities']}

            df_part1 = self.processed_data.copy()
            df_part1['Processos'] = df_part1['Atividade'].map(activity_to_group).fillna('')
            
            # Agrupar tempos em listas por atividade
            pivoted_times = df_part1.groupby(['Processos', 'Atividade'])['Tempo'].apply(list).reset_index(name='Tempos')
            
            # Expandir as listas de tempo em colunas "Amostra N"
            max_samples = pivoted_times['Tempos'].str.len().max()
            sample_cols = [f'Amostra {i+1}' for i in range(max_samples)]
            
            time_df = pd.DataFrame(pivoted_times['Tempos'].tolist(), index=pivoted_times.index, columns=sample_cols)
            
            # Juntar as informações com os tempos expandidos
            part1_df = pd.concat([pivoted_times[['Processos', 'Atividade']], time_df], axis=1)
            part1_df.insert(1, 'COD', '') # Adicionar coluna COD em branco
            part1_df.rename(columns={'Atividade': 'Atividades da Coleta'}, inplace=True)


            # --- Parte 2: Preparar tabela de análise ---
            analysis_data = []
            headers = ['Atividade/Grupo'] + [self.results_tree.heading(col)['text'] for col in self.results_tree['columns']]

            for item_id in self.results_tree.get_children():
                # Item pai (Grupo ou categoria principal)
                parent_text = self.results_tree.item(item_id, 'text').strip()
                parent_values = list(self.results_tree.item(item_id, 'values'))
                analysis_data.append([parent_text] + parent_values)

                # Itens filhos (atividades dentro do grupo)
                for child_id in self.results_tree.get_children(item_id):
                    child_text = "    " + self.results_tree.item(child_id, 'text').strip()
                    child_values = list(self.results_tree.item(child_id, 'values'))
                    analysis_data.append([child_text] + child_values)
            
            part2_df = pd.DataFrame(analysis_data, columns=headers)

            # --- Escrever no arquivo Excel ---
            with pd.ExcelWriter(filename, engine='openpyxl') as writer:
                part1_df.to_excel(writer, sheet_name='Exportação Completa', index=False)
                
                # Calcular a coluna de início para a segunda tabela
                start_col_part2 = part1_df.shape[1] + 1  # +1 para a coluna em branco
                
                part2_df.to_excel(writer, sheet_name='Exportação Completa', index=False, startrow=0, startcol=start_col_part2)

            self.export_status.config(text=f"Exportado para: {filename}")
            messagebox.showinfo("Sucesso", "Dados exportados com sucesso para o formato customizado!")

        except Exception as e:
            import traceback
            error_details = traceback.format_exc()
            print(f"Erro detalhado na exportação: {error_details}")
            messagebox.showerror("Erro", f"Erro ao exportar para Excel: {str(e)}")
            
    def export_to_csv(self):
        """Exportar resultados para CSV (formato simples)"""
        if self.processed_data is None:
            messagebox.showwarning("Aviso", "Nenhum dado para exportar")
            return
            
        try:
            filename = filedialog.asksaveasfilename(
                defaultextension=".csv",
                filetypes=[("CSV files", "*.csv")]
            )
            
            if filename:
                # Exportação CSV mantém o formato simples dos dados processados
                self.processed_data.to_csv(filename, index=False)
                self.export_status.config(text=f"Exportado para: {filename}")
                messagebox.showinfo("Sucesso", "Dados brutos exportados com sucesso para CSV!")
                
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao exportar: {str(e)}")
    
    def debug_data(self):
        """Função de debug para analisar os dados dos arquivos"""
        if not self.uploaded_files:
            messagebox.showwarning("Aviso", "Nenhum arquivo selecionado")
            return
            
        debug_info = "=== DEBUG DOS DADOS ===\n\n"
        
        for i, file_path in enumerate(self.uploaded_files):
            try:
                debug_info += f"ARQUIVO {i+1}: {os.path.basename(file_path)}\n"
                
                # Ler arquivo
                if file_path.endswith('.xlsx'):
                    df = pd.read_excel(file_path)
                else:
                    df = pd.read_csv(file_path)
                
                debug_info += f"• Dimensões: {df.shape[0]} linhas x {df.shape[1]} colunas\n"
                debug_info += f"• Colunas: {list(df.columns)}\n"
                
                # Se colunas foram selecionadas, analisar
                if self.activity_combo.get() and self.time_combo.get():
                    activity_col = self.activity_combo.get()
                    time_col = self.time_combo.get()
                    
                    if activity_col in df.columns:
                        debug_info += f"• Coluna atividade '{activity_col}': {df[activity_col].count()} valores não-nulos\n"
                        debug_info += f"   Exemplos: {list(df[activity_col].dropna().head(3))}\n"
                    else:
                        debug_info += f"• Coluna atividade '{activity_col}': NÃO ENCONTRADA\n"
                    
                    if time_col in df.columns:
                        debug_info += f"• Coluna tempo '{time_col}': {df[time_col].count()} valores não-nulos\n"
                        debug_info += f"   Exemplos: {list(df[time_col].dropna().head(3))}\n"
                        debug_info += f"   Tipos: {df[time_col].dtype}\n"
                        
                        # Detectar formato de tempo
                        sample_values = df[time_col].dropna().head(5)
                        time_formats = []
                        for val in sample_values:
                            val_str = str(val).strip()
                            if ':' in val_str:
                                parts = val_str.split(':')
                                if len(parts) == 3:
                                    time_formats.append("HH:MM:SS")
                                elif len(parts) == 2:
                                    time_formats.append("MM:SS")
                                else:
                                    time_formats.append("Formato desconhecido")
                            else:
                                try:
                                    float(val_str.replace(',', '.'))
                                    time_formats.append("Numérico")
                                except:
                                    time_formats.append("Texto")
                        
                        debug_info += f"   Formatos detectados: {time_formats}\n"
                    else:
                        debug_info += f"• Coluna tempo '{time_col}': NÃO ENCONTRADA\n"
                
                debug_info += f"• Primeiras 3 linhas:\n{df.head(3).to_string()}\n"
                debug_info += "\n" + "="*50 + "\n\n"
                
            except Exception as e:
                debug_info += f"ERRO ao ler arquivo: {str(e)}\n\n"
        
        # Mostrar em uma janela de texto
        debug_window = tk.Toplevel(self.root)
        debug_window.title("Debug dos Dados")
        debug_window.geometry("800x600")
        
        text_widget = tk.Text(debug_window, wrap=tk.WORD)
        scrollbar = ttk.Scrollbar(debug_window, orient="vertical", command=text_widget.yview)
        text_widget.configure(yscrollcommand=scrollbar.set)
        
        text_widget.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        text_widget.insert("1.0", debug_info)

def main():
    root = tk.Tk()
    app = TimeStudyAnalyzer(root)
    root.mainloop()

if __name__ == "__main__":
    main()
