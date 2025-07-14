
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
        analysis_frame.grid_columnconfigure(1, weight=1)
        
        # Título
        title_label = ttk.Label(analysis_frame, text="Análise Estatística", font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=2, pady=10)
        
        # Botão de análise
        analyze_btn = ttk.Button(analysis_frame, text="Executar Análise", command=self.perform_analysis)
        analyze_btn.grid(row=0, column=1, sticky="e", padx=10)
        
        # Frame esquerdo - Tabela de resultados
        left_frame = ttk.LabelFrame(analysis_frame, text="Resultados Estatísticos")
        left_frame.grid(row=1, column=0, sticky="nsew", padx=(0, 5), pady=5)
        left_frame.grid_rowconfigure(0, weight=1)
        left_frame.grid_columnconfigure(0, weight=1)
        
        self.results_tree = ttk.Treeview(left_frame, columns=("Q1", "Mediana", "Q3", "IQR", "Média", "Outliers"), show="tree headings")
        self.results_tree.heading("#0", text="Atividade/Grupo")
        self.results_tree.heading("Q1", text="Q1")
        self.results_tree.heading("Mediana", text="Mediana")
        self.results_tree.heading("Q3", text="Q3")
        self.results_tree.heading("IQR", text="IQR")
        self.results_tree.heading("Média", text="Média")
        self.results_tree.heading("Outliers", text="Outliers")
        self.results_tree.grid(row=0, column=0, sticky="nsew", padx=5, pady=5)
        
        results_scroll = ttk.Scrollbar(left_frame, orient="vertical", command=self.results_tree.yview)
        results_scroll.grid(row=0, column=1, sticky="ns")
        self.results_tree.configure(yscrollcommand=results_scroll.set)
        
        # Frame direito - Gráficos
        right_frame = ttk.LabelFrame(analysis_frame, text="Visualizações")
        right_frame.grid(row=1, column=1, sticky="nsew", padx=(5, 0), pady=5)
        right_frame.grid_rowconfigure(0, weight=1)
        right_frame.grid_columnconfigure(0, weight=1)
        
        # Frame para matplotlib
        self.plot_frame = ttk.Frame(right_frame)
        self.plot_frame.grid(row=0, column=0, sticky="nsew", padx=5, pady=5)
        self.plot_frame.grid_rowconfigure(0, weight=1)
        self.plot_frame.grid_columnconfigure(0, weight=1)
        
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
        export_csv_btn = ttk.Button(options_frame, text="Exportar para CSV", command=self.export_to_csv)
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
        """Processar dados dos arquivos"""
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
            
            for file_path in self.uploaded_files:
                try:
                    # Ler arquivo
                    if file_path.endswith('.xlsx'):
                        df = pd.read_excel(file_path)
                    else:
                        # Tentar diferentes encodings para CSV
                        try:
                            df = pd.read_csv(file_path, encoding='utf-8')
                        except UnicodeDecodeError:
                            try:
                                df = pd.read_csv(file_path, encoding='latin1')
                            except UnicodeDecodeError:
                                df = pd.read_csv(file_path, encoding='cp1252')
                    
                    print(f"Arquivo {os.path.basename(file_path)}: {len(df)} linhas lidas")
                    total_rows_read += len(df)
                    
                    # Extrair colunas selecionadas
                    activity_col = self.activity_combo.get()
                    time_col = self.time_combo.get()
                    
                    if activity_col in df.columns and time_col in df.columns:
                        # Criar cópia dos dados
                        data = df[[activity_col, time_col]].copy()
                        data.columns = ['Atividade', 'Tempo']
                        
                        # Limpar espaços em branco nas atividades
                        data['Atividade'] = data['Atividade'].astype(str).str.strip()
                        
                        # Filtrar linhas vazias ou nulas
                        data = data[data['Atividade'].notna()]
                        data = data[data['Atividade'] != '']
                        data = data[data['Atividade'] != 'nan']
                        
                        print(f"Arquivo {os.path.basename(file_path)}: {len(data)} linhas válidas após limpeza")
                        
                        if len(data) > 0:
                            all_data.append(data)
                            files_processed += 1
                    else:
                        print(f"Colunas não encontradas no arquivo {os.path.basename(file_path)}")
                        print(f"Colunas disponíveis: {list(df.columns)}")
                        
                except Exception as e:
                    print(f"Erro ao processar arquivo {os.path.basename(file_path)}: {str(e)}")
                    continue
                    
            print(f"Total de arquivos processados: {files_processed}")
            print(f"Total de linhas lidas: {total_rows_read}")
            
            if all_data:
                # Concatenar todos os dados
                self.processed_data = pd.concat(all_data, ignore_index=True)
                print(f"Dados concatenados: {len(self.processed_data)} linhas")
                
                # Converter tempo para numérico
                original_count = len(self.processed_data)
                
                # Tentar converter tempo - aceitar diferentes formatos
                def convert_time(time_val):
                    if pd.isna(time_val):
                        return None
                    
                    # Se já é numérico
                    if isinstance(time_val, (int, float)):
                        return float(time_val)
                    
                    # Se é string, tentar converter
                    time_str = str(time_val).strip()
                    
                    # Verificar se é formato de tempo (HH:MM:SS, MM:SS, etc.)
                    if ':' in time_str:
                        try:
                            # Dividir por ':'
                            parts = time_str.split(':')
                            total_seconds = 0
                            
                            if len(parts) == 3:  # HH:MM:SS
                                hours = float(parts[0])
                                minutes = float(parts[1])
                                seconds = float(parts[2])
                                total_seconds = hours * 3600 + minutes * 60 + seconds
                            elif len(parts) == 2:  # MM:SS
                                minutes = float(parts[0])
                                seconds = float(parts[1])
                                total_seconds = minutes * 60 + seconds
                            else:
                                return None
                                
                            return total_seconds
                            
                        except (ValueError, IndexError):
                            # Se falhar, tentar outros métodos
                            pass
                    
                    # Remover caracteres não numéricos comuns
                    time_str = time_str.replace(',', '.')  # Vírgula para ponto decimal
                    time_str = time_str.replace(' ', '')   # Remover espaços
                    
                    # Tentar converter para float
                    try:
                        return float(time_str)
                    except ValueError:
                        return None
                
                self.processed_data['Tempo'] = self.processed_data['Tempo'].apply(convert_time)
                
                # Remover linhas com tempo inválido
                self.processed_data = self.processed_data.dropna(subset=['Tempo'])
                
                # Remover tempos negativos ou zero
                self.processed_data = self.processed_data[self.processed_data['Tempo'] > 0]
                
                final_count = len(self.processed_data)
                print(f"Dados finais após validação: {final_count} linhas")
                print(f"Linhas removidas por dados inválidos: {original_count - final_count}")
                
                if final_count > 0:
                    self.update_processed_preview()
                    self.update_available_activities()
                    messagebox.showinfo("Sucesso", 
                        f"Dados processados com sucesso!\n"
                        f"• {files_processed} arquivo(s) processado(s)\n"
                        f"• {total_rows_read} linha(s) lida(s)\n"
                        f"• {final_count} registro(s) válido(s)\n"
                        f"• {original_count - final_count} registro(s) removido(s) por dados inválidos")
                else:
                    messagebox.showerror("Erro", 
                        f"Nenhum dado válido encontrado!\n"
                        f"• {files_processed} arquivo(s) processado(s)\n"
                        f"• {total_rows_read} linha(s) lida(s)\n"
                        f"• Todos os registros foram removidos por conterem dados inválidos\n\n"
                        f"Verifique se:\n"
                        f"• A coluna de tempo contém valores numéricos\n"
                        f"• A coluna de atividade não está vazia\n"
                        f"• Os dados não contêm apenas cabeçalhos")
            else:
                messagebox.showerror("Erro", 
                    f"Nenhum arquivo pôde ser processado!\n"
                    f"• {len(self.uploaded_files)} arquivo(s) selecionado(s)\n"
                    f"• 0 arquivo(s) processado(s)\n\n"
                    f"Verifique se:\n"
                    f"• Os arquivos não estão corrompidos\n"
                    f"• As colunas selecionadas existem nos arquivos\n"
                    f"• Os arquivos contêm dados válidos")
                
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
                    x = (choice_window.winfo_screenwidth() // 2) - (choice_window.winfo_width() // 2)
                    y = (choice_window.winfo_screenheight() // 2) - (choice_window.winfo_height() // 2)
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
        x = (group_window.winfo_screenwidth() // 2) - (group_window.winfo_width() // 2)
        y = (group_window.winfo_screenheight() // 2) - (group_window.winfo_height() // 2)
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
        x = (select_window.winfo_screenwidth() // 2) - (select_window.winfo_width() // 2)
        y = (select_window.winfo_screenheight() // 2) - (select_window.winfo_height() // 2)
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
        
    def perform_analysis(self):
        """Executar análise estatística"""
        if self.processed_data is None:
            messagebox.showwarning("Aviso", "Processe os dados primeiro")
            return
            
        try:
            # Limpar resultados anteriores
            for item in self.results_tree.get_children():
                self.results_tree.delete(item)
            
            # Verificar se há dados para analisar
            if len(self.processed_data) == 0:
                messagebox.showwarning("Aviso", "Não há dados válidos para analisar")
                return
            
            total_analyzed = 0
            
            # Analisar grupos primeiro (se existirem)
            if self.activity_groups:
                for group_name, group_data in self.activity_groups.items():
                    if group_data['activities']:
                        # Obter dados do grupo
                        group_times = self.processed_data[
                            self.processed_data['Atividade'].isin(group_data['activities'])
                        ]['Tempo']
                        
                        if len(group_times) > 0:
                            total_analyzed += len(group_times)
                            
                            # Calcular estatísticas do grupo
                            q1 = group_times.quantile(0.25)
                            median = group_times.median()
                            q3 = group_times.quantile(0.75)
                            iqr = q3 - q1
                            mean_all = group_times.mean()
                            
                            # Detectar outliers
                            lower_bound = q1 - 1.5 * iqr
                            upper_bound = q3 + 1.5 * iqr
                            outliers = group_times[(group_times < lower_bound) | (group_times > upper_bound)]
                            
                            # Média sem outliers
                            clean_times = group_times[(group_times >= lower_bound) & (group_times <= upper_bound)]
                            mean_no_outliers = clean_times.mean() if len(clean_times) > 0 else mean_all
                            
                            # Adicionar grupo à árvore
                            group_item = self.results_tree.insert("", tk.END, text=f"📁 {group_name} ({len(group_times)} registros)", values=(
                                f"{q1:.2f}s",
                                f"{median:.2f}s",
                                f"{q3:.2f}s",
                                f"{iqr:.2f}s",
                                f"{mean_no_outliers:.2f}s",
                                f"{len(outliers)} outliers"
                            ))
                            
                            # Analisar atividades individuais do grupo
                            for activity in group_data['activities']:
                                activity_times = self.processed_data[
                                    self.processed_data['Atividade'] == activity
                                ]['Tempo']
                                
                                if len(activity_times) > 0:
                                    a_q1 = activity_times.quantile(0.25)
                                    a_median = activity_times.median()
                                    a_q3 = activity_times.quantile(0.75)
                                    a_iqr = a_q3 - a_q1
                                    a_mean_all = activity_times.mean()
                                    
                                    a_lower_bound = a_q1 - 1.5 * a_iqr
                                    a_upper_bound = a_q3 + 1.5 * a_iqr
                                    a_outliers = activity_times[(activity_times < a_lower_bound) | (activity_times > a_upper_bound)]
                                    
                                    a_clean_times = activity_times[(activity_times >= a_lower_bound) & (activity_times <= a_upper_bound)]
                                    a_mean_no_outliers = a_clean_times.mean() if len(a_clean_times) > 0 else a_mean_all
                                    
                                    # Adicionar atividade como filho do grupo
                                    self.results_tree.insert(group_item, tk.END, text=f"  📊 {activity} ({len(activity_times)} reg.)", values=(
                                        f"{a_q1:.2f}s",
                                        f"{a_median:.2f}s",
                                        f"{a_q3:.2f}s",
                                        f"{a_iqr:.2f}s",
                                        f"{a_mean_no_outliers:.2f}s",
                                        f"{len(a_outliers)}"
                                    ))
            
            # Analisar atividades não agrupadas
            grouped_activities = set()
            for group_data in self.activity_groups.values():
                grouped_activities.update(group_data['activities'])
            
            ungrouped_activities = self.processed_data[
                ~self.processed_data['Atividade'].isin(grouped_activities)
            ]
            
            if len(ungrouped_activities) > 0:
                # Adicionar seção de atividades não agrupadas
                ungrouped_item = self.results_tree.insert("", tk.END, text=f"📋 Atividades Não Agrupadas ({len(ungrouped_activities)} registros)", values=())
                
                # Agrupar por atividade
                grouped = ungrouped_activities.groupby('Atividade')['Tempo']
                
                for activity, times in grouped:
                    if len(times) > 0:
                        total_analyzed += len(times)
                        
                        q1 = times.quantile(0.25)
                        median = times.median()
                        q3 = times.quantile(0.75)
                        iqr = q3 - q1
                        mean_all = times.mean()
                        
                        # Detectar outliers
                        lower_bound = q1 - 1.5 * iqr
                        upper_bound = q3 + 1.5 * iqr
                        outliers = times[(times < lower_bound) | (times > upper_bound)]
                        
                        # Média sem outliers
                        clean_times = times[(times >= lower_bound) & (times <= upper_bound)]
                        mean_no_outliers = clean_times.mean() if len(clean_times) > 0 else mean_all
                        
                        # Adicionar à árvore
                        self.results_tree.insert(ungrouped_item, tk.END, text=f"  📊 {activity} ({len(times)} reg.)", values=(
                            f"{q1:.2f}s",
                            f"{median:.2f}s",
                            f"{q3:.2f}s",
                            f"{iqr:.2f}s",
                            f"{mean_no_outliers:.2f}s",
                            f"{len(outliers)}"
                        ))
            
            # Se não há grupos, analisar todas as atividades
            if not self.activity_groups and len(ungrouped_activities) == 0:
                # Analisar todas as atividades
                all_item = self.results_tree.insert("", tk.END, text=f"📊 Todas as Atividades ({len(self.processed_data)} registros)", values=())
                
                grouped = self.processed_data.groupby('Atividade')['Tempo']
                
                for activity, times in grouped:
                    if len(times) > 0:
                        total_analyzed += len(times)
                        
                        q1 = times.quantile(0.25)
                        median = times.median()
                        q3 = times.quantile(0.75)
                        iqr = q3 - q1
                        mean_all = times.mean()
                        
                        # Detectar outliers
                        lower_bound = q1 - 1.5 * iqr
                        upper_bound = q3 + 1.5 * iqr
                        outliers = times[(times < lower_bound) | (times > upper_bound)]
                        
                        # Média sem outliers
                        clean_times = times[(times >= lower_bound) & (times <= upper_bound)]
                        mean_no_outliers = clean_times.mean() if len(clean_times) > 0 else mean_all
                        
                        # Adicionar à árvore
                        self.results_tree.insert(all_item, tk.END, text=f"  📊 {activity} ({len(times)} reg.)", values=(
                            f"{q1:.2f}s",
                            f"{median:.2f}s",
                            f"{q3:.2f}s",
                            f"{iqr:.2f}s",
                            f"{mean_no_outliers:.2f}s",
                            f"{len(outliers)}"
                        ))
                
            # Expandir todas as seções
            for item in self.results_tree.get_children():
                self.results_tree.item(item, open=True)
                
            # Criar gráfico
            self.create_boxplot()
            
            # Mostrar estatísticas gerais
            unique_activities = len(self.processed_data['Atividade'].unique())
            total_time = self.processed_data['Tempo'].sum()
            avg_time = self.processed_data['Tempo'].mean()
            
            messagebox.showinfo("Análise Concluída", 
                f"Análise estatística concluída com sucesso!\n\n"
                f"📊 Estatísticas Gerais:\n"
                f"• Total de registros analisados: {total_analyzed}\n"
                f"• Atividades únicas: {unique_activities}\n"
                f"• Tempo total: {total_time:.2f} segundos ({total_time/60:.1f} minutos)\n"
                f"• Tempo médio por registro: {avg_time:.2f} segundos\n"
                f"• Grupos criados: {len(self.activity_groups)}")
            
        except Exception as e:
            import traceback
            error_details = traceback.format_exc()
            print(f"Erro detalhado na análise: {error_details}")
            messagebox.showerror("Erro", f"Erro na análise estatística:\n{str(e)}\n\nVerifique o console para mais detalhes.")
            
    def create_boxplot(self):
        """Criar boxplot dos dados"""
        try:
            # Limpar frame anterior
            for widget in self.plot_frame.winfo_children():
                widget.destroy()
                
            # Verificar se há dados
            if self.processed_data is None or len(self.processed_data) == 0:
                # Mostrar mensagem de não há dados
                no_data_label = ttk.Label(self.plot_frame, text="Nenhum dado disponível para visualização", 
                                        font=("Arial", 12), foreground="gray")
                no_data_label.grid(row=0, column=0, sticky="nsew", padx=20, pady=20)
                return
            
            # Preparar dados para boxplot
            activities = []
            times_list = []
            
            grouped = self.processed_data.groupby('Atividade')['Tempo']
            for activity, times in grouped:
                if len(times) > 0:  # Só incluir atividades com dados
                    activities.append(activity[:20] + "..." if len(activity) > 20 else activity)  # Truncar nomes longos
                    times_list.append(times.values)
            
            if not times_list:
                # Mostrar mensagem de não há dados válidos
                no_data_label = ttk.Label(self.plot_frame, text="Nenhum dado válido para visualização", 
                                        font=("Arial", 12), foreground="gray")
                no_data_label.grid(row=0, column=0, sticky="nsew", padx=20, pady=20)
                return
            
            # Criar figura com tamanho ajustado
            fig, ax = plt.subplots(figsize=(10, 6))
            fig.patch.set_facecolor('white')
            
            # Criar boxplot
            bp = ax.boxplot(times_list, labels=activities, patch_artist=True)
            
            # Personalizar cores
            colors = ['lightblue', 'lightgreen', 'lightcoral', 'lightyellow', 'lightpink', 'lightgray']
            for patch, color in zip(bp['boxes'], colors * len(bp['boxes'])):
                patch.set_facecolor(color)
                patch.set_alpha(0.7)
            
            # Configurar gráfico
            ax.set_title('Distribuição de Tempos por Atividade', fontsize=14, fontweight='bold')
            ax.set_ylabel('Tempo (segundos)', fontsize=12)
            ax.set_xlabel('Atividades', fontsize=12)
            
            # Rotacionar labels se necessário
            if len(activities) > 5:
                ax.tick_params(axis='x', rotation=45)
            
            # Adicionar grid
            ax.grid(True, alpha=0.3)
            
            # Ajustar layout
            plt.tight_layout()
            
            # Incorporar no tkinter
            canvas = FigureCanvasTkAgg(fig, self.plot_frame)
            canvas.draw()
            canvas.get_tk_widget().grid(row=0, column=0, sticky="nsew")
            
            print(f"Boxplot criado com sucesso para {len(activities)} atividades")
            
        except Exception as e:
            import traceback
            error_details = traceback.format_exc()
            print(f"Erro detalhado ao criar gráfico: {error_details}")
            
            # Mostrar mensagem de erro no frame
            error_label = ttk.Label(self.plot_frame, text=f"Erro ao criar gráfico:\n{str(e)}", 
                                  font=("Arial", 10), foreground="red")
            error_label.grid(row=0, column=0, sticky="nsew", padx=20, pady=20)
            
    def export_to_excel(self):
        """Exportar resultados para Excel"""
        if self.processed_data is None:
            messagebox.showwarning("Aviso", "Nenhum dado para exportar")
            return
            
        try:
            filename = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx")]
            )
            
            if filename:
                with pd.ExcelWriter(filename, engine='openpyxl') as writer:
                    self.processed_data.to_excel(writer, sheet_name='Dados Processados', index=False)
                    
                self.export_status.config(text=f"Exportado para: {filename}")
                messagebox.showinfo("Sucesso", "Dados exportados com sucesso!")
                
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao exportar: {str(e)}")
            
    def export_to_csv(self):
        """Exportar resultados para CSV"""
        if self.processed_data is None:
            messagebox.showwarning("Aviso", "Nenhum dado para exportar")
            return
            
        try:
            filename = filedialog.asksaveasfilename(
                defaultextension=".csv",
                filetypes=[("CSV files", "*.csv")]
            )
            
            if filename:
                self.processed_data.to_csv(filename, index=False)
                self.export_status.config(text=f"Exportado para: {filename}")
                messagebox.showinfo("Sucesso", "Dados exportados com sucesso!")
                
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
                        debug_info += f"  Exemplos: {list(df[activity_col].dropna().head(3))}\n"
                    else:
                        debug_info += f"• Coluna atividade '{activity_col}': NÃO ENCONTRADA\n"
                    
                    if time_col in df.columns:
                        debug_info += f"• Coluna tempo '{time_col}': {df[time_col].count()} valores não-nulos\n"
                        debug_info += f"  Exemplos: {list(df[time_col].dropna().head(3))}\n"
                        debug_info += f"  Tipos: {df[time_col].dtype}\n"
                        
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
                        
                        debug_info += f"  Formatos detectados: {time_formats}\n"
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
