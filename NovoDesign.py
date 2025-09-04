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

# Configurações de design moderno
class ModernColors:
    # Cores principais
    PRIMARY = "#2563EB"          # Azul moderno
    PRIMARY_DARK = "#1D4ED8"     # Azul escuro
    PRIMARY_LIGHT = "#DBEAFE"    # Azul claro
    
    # Cores de fundo
    BACKGROUND = "#F8FAFC"       # Fundo principal
    SURFACE = "#FFFFFF"          # Superfície de cards
    SURFACE_VARIANT = "#F1F5F9"  # Variante de superfície
    
    # Cores de texto
    TEXT_PRIMARY = "#0F172A"     # Texto principal
    TEXT_SECONDARY = "#64748B"   # Texto secundário
    TEXT_DISABLED = "#94A3B8"    # Texto desabilitado
    
    # Cores de status
    SUCCESS = "#059669"          # Verde sucesso
    SUCCESS_LIGHT = "#D1FAE5"    # Verde claro
    WARNING = "#D97706"          # Laranja aviso
    WARNING_LIGHT = "#FED7AA"    # Laranja claro
    ERROR = "#DC2626"            # Vermelho erro
    ERROR_LIGHT = "#FEE2E2"      # Vermelho claro
    
    # Cores de acento
    ACCENT = "#7C3AED"           # Roxo
    ACCENT_LIGHT = "#EDE9FE"     # Roxo claro
    
    # Cores neutras
    BORDER = "#E2E8F0"           # Bordas
    DIVIDER = "#CBD5E1"          # Divisores
    SHADOW = "gray80"            # Sombras

class ModernStyles:
    def __init__(self):
        self.setup_styles()
    
    def setup_styles(self):
        """Configurar estilos modernos para ttk"""
        style = ttk.Style()
        
        # Configurar tema base
        style.theme_use('clam')
        
        # Estilo para Notebook (abas)
        style.configure('Modern.TNotebook',
                       background=ModernColors.BACKGROUND,
                       borderwidth=0)
        
        style.configure('Modern.TNotebook.Tab',
                       background=ModernColors.SURFACE_VARIANT,
                       foreground=ModernColors.TEXT_SECONDARY,
                       padding=[20, 12],
                       borderwidth=0,
                       focuscolor='none')
        
        style.map('Modern.TNotebook.Tab',
                 background=[('selected', ModernColors.SURFACE),
                           ('active', ModernColors.PRIMARY_LIGHT)],
                 foreground=[('selected', ModernColors.PRIMARY),
                           ('active', ModernColors.PRIMARY)])
        
        # Estilo para botões principais
        style.configure('Primary.TButton',
                       background=ModernColors.PRIMARY,
                       foreground='white',
                       padding=[16, 8],
                       borderwidth=0,
                       focuscolor='none',
                       font=('Segoe UI', 10, 'normal'))
        
        style.map('Primary.TButton',
                 background=[('active', ModernColors.PRIMARY_DARK),
                           ('pressed', ModernColors.PRIMARY_DARK)])
        
        # Estilo para botões secundários
        style.configure('Secondary.TButton',
                       background=ModernColors.SURFACE,
                       foreground=ModernColors.TEXT_PRIMARY,
                       padding=[16, 8],
                       borderwidth=1,
                       relief='solid',
                       focuscolor='none',
                       font=('Segoe UI', 10, 'normal'))
        
        style.map('Secondary.TButton',
                 background=[('active', ModernColors.SURFACE_VARIANT),
                           ('pressed', ModernColors.SURFACE_VARIANT)],
                 bordercolor=[('focus', ModernColors.PRIMARY)])
        
        # Estilo para labels de título
        style.configure('Title.TLabel',
                       background=ModernColors.BACKGROUND,
                       foreground=ModernColors.TEXT_PRIMARY,
                       font=('Segoe UI', 18, 'bold'))
        
        # Estilo para labels de subtítulo
        style.configure('Subtitle.TLabel',
                       background=ModernColors.BACKGROUND,
                       foreground=ModernColors.TEXT_SECONDARY,
                       font=('Segoe UI', 12, 'normal'))
        
        # Estilo para frames de card
        style.configure('Card.TFrame',
                       background=ModernColors.SURFACE,
                       relief='flat',
                       borderwidth=1)
        
        # Estilo para LabelFrames
        style.configure('Modern.TLabelframe',
                       background=ModernColors.SURFACE,
                       foreground=ModernColors.TEXT_PRIMARY,
                       borderwidth=1,
                       relief='solid',
                       font=('Segoe UI', 11, 'bold'))
        
        style.configure('Modern.TLabelframe.Label',
                       background=ModernColors.SURFACE,
                       foreground=ModernColors.TEXT_PRIMARY,
                       font=('Segoe UI', 11, 'bold'))
        
        # Estilo para Combobox
        style.configure('Modern.TCombobox',
                       selectbackground=ModernColors.PRIMARY_LIGHT,
                       fieldbackground=ModernColors.SURFACE,
                       background=ModernColors.SURFACE,
                       borderwidth=1,
                       relief='solid',
                       focuscolor='none')
        
        # Estilo para Treeview
        style.configure('Modern.Treeview',
                       background=ModernColors.SURFACE,
                       foreground=ModernColors.TEXT_PRIMARY,
                       fieldbackground=ModernColors.SURFACE,
                       borderwidth=1,
                       relief='solid')
        
        style.configure('Modern.Treeview.Heading',
                       background=ModernColors.SURFACE_VARIANT,
                       foreground=ModernColors.TEXT_PRIMARY,
                       borderwidth=1,
                       relief='solid',
                       font=('Segoe UI', 10, 'bold'))
        
        style.map('Modern.Treeview',
                 background=[('selected', ModernColors.PRIMARY_LIGHT)],
                 foreground=[('selected', ModernColors.PRIMARY)])
        
        # Estilo para Entry
        style.configure('Modern.TEntry',
                       fieldbackground=ModernColors.SURFACE,
                       borderwidth=1,
                       relief='solid',
                       focuscolor='none')
        
        style.map('Modern.TEntry',
                 bordercolor=[('focus', ModernColors.PRIMARY)])
        
        # Estilo para Progressbar
        style.configure('Modern.TProgressbar',
                       background=ModernColors.PRIMARY,
                       troughcolor=ModernColors.SURFACE_VARIANT,
                       borderwidth=0,
                       lightcolor=ModernColors.PRIMARY,
                       darkcolor=ModernColors.PRIMARY)

class TimeStudyAnalyzer:
    def __init__(self, root):
        self.root = root
        self.root.title("📊 Plataforma de Análise de Estudo de Tempos")
        self.root.geometry("1400x900")
        self.root.minsize(1200, 700)
        
        # Configurar cores e estilo moderno
        self.root.configure(bg=ModernColors.BACKGROUND)
        self.styles = ModernStyles()
        
        # Configurar ícone (se disponível)
        try:
            # Tentar definir um ícone da janela (opcional)
            pass
        except:
            pass
        
        # Inicializar banco de dados
        self.init_database()
        
        # Variáveis de estado
        self.uploaded_files = []
        self.processed_data = None
        self.activity_column = None
        self.time_column = None
        self.rework_column = None  # Nova coluna de retrabalho
        self.unified_activities = {}
        self.activity_groups = {}
        
        # Criar interface moderna
        self.create_modern_interface()
        
        # Centralizar janela na tela
        self.center_window()
        
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
        
    def center_window(self):
        """Centralizar janela na tela"""
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        pos_x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        pos_y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{pos_x}+{pos_y}')
    
    def create_modern_card(self, parent, title=None, subtitle=None):
        """Criar um card moderno com título e subtítulo opcionais"""
        # Frame principal do card
        card_frame = ttk.Frame(parent, style='Card.TFrame')
        
        if title or subtitle:
            # Header do card
            header_frame = ttk.Frame(card_frame, style='Card.TFrame')
            header_frame.pack(fill=tk.X, padx=20, pady=(20, 10))
            
            if title:
                title_label = ttk.Label(header_frame, text=title, style='Title.TLabel')
                title_label.pack(anchor='w')
            
            if subtitle:
                subtitle_label = ttk.Label(header_frame, text=subtitle, style='Subtitle.TLabel')
                subtitle_label.pack(anchor='w', pady=(5, 0))
            
            # Linha divisória
            separator = ttk.Separator(card_frame, orient='horizontal')
            separator.pack(fill=tk.X, padx=20)
        
        # Conteúdo do card
        content_frame = ttk.Frame(card_frame, style='Card.TFrame')
        content_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        return card_frame, content_frame
    
    def create_modern_interface(self):
        """Criar interface moderna com design melhorado"""
        # Container principal
        main_container = ttk.Frame(self.root)
        main_container.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # Header da aplicação
        self.create_header(main_container)
        
        # Notebook para abas com estilo moderno
        self.notebook = ttk.Notebook(main_container, style='Modern.TNotebook')
        self.notebook.pack(fill=tk.BOTH, expand=True, pady=(20, 0))
        
        # Configurar grid weights
        self.root.grid_rowconfigure(0, weight=1)
        self.root.grid_columnconfigure(0, weight=1)
        
        # Criar abas com design moderno
        self.create_modern_upload_tab()
        self.create_modern_mapping_tab()
        self.create_modern_unification_tab()
        self.create_modern_grouping_tab()
        self.create_modern_analysis_tab()
        self.create_modern_export_tab()
    
    def create_header(self, parent):
        """Criar header moderno da aplicação"""
        header_frame = ttk.Frame(parent)
        header_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Título principal
        main_title = ttk.Label(header_frame, 
                              text="📊 Plataforma de Análise de Estudo de Tempos",
                              font=('Segoe UI', 24, 'bold'),
                              foreground=ModernColors.TEXT_PRIMARY)
        main_title.pack(anchor='w')
        
        # Subtítulo
        subtitle = ttk.Label(header_frame,
                           text="Análise estatística avançada para otimização de processos industriais",
                           font=('Segoe UI', 12, 'normal'),
                           foreground=ModernColors.TEXT_SECONDARY)
        subtitle.pack(anchor='w', pady=(5, 0))
        
        # Linha divisória
        separator = ttk.Separator(header_frame, orient='horizontal')
        separator.pack(fill=tk.X, pady=(15, 0))
    def create_modern_upload_tab(self):
        """Aba de upload de arquivos com design moderno"""
        upload_frame = ttk.Frame(self.notebook)
        self.notebook.add(upload_frame, text="🗂️  Upload de Arquivos")
        
        # Configurar fundo
        upload_frame.configure(style='Card.TFrame')
        
        # Card principal
        card, content = self.create_modern_card(upload_frame, 
                                               "Upload de Arquivos",
                                               "Selecione os arquivos Excel ou CSV para análise")
        card.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # Configurar grid do conteúdo
        content.grid_rowconfigure(2, weight=1)
        content.grid_rowconfigure(4, weight=2)
        content.grid_columnconfigure(0, weight=1)
        
        # Frame para botões com design moderno
        button_frame = ttk.Frame(content)
        button_frame.grid(row=0, column=0, pady=(0, 20), sticky="ew")
        
        # Botão de upload principal
        upload_btn = ttk.Button(button_frame, 
                               text="📁 Selecionar Arquivos",
                               command=self.upload_files,
                               style='Primary.TButton')
        upload_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # Botão de limpar
        clear_btn = ttk.Button(button_frame,
                              text="🗑️ Limpar Lista",
                              command=self.clear_files,
                              style='Secondary.TButton')
        clear_btn.pack(side=tk.LEFT)
        
        # Status de arquivos
        self.upload_status = ttk.Label(button_frame,
                                     text="Nenhum arquivo selecionado",
                                     font=('Segoe UI', 10, 'normal'),
                                     foreground=ModernColors.TEXT_SECONDARY)
        self.upload_status.pack(side=tk.RIGHT)
        
        # Lista de arquivos com design moderno
        files_label = ttk.Label(content, text="📋 Arquivos Selecionados:",
                               font=('Segoe UI', 12, 'bold'),
                               foreground=ModernColors.TEXT_PRIMARY)
        files_label.grid(row=1, column=0, sticky="w", pady=(0, 10))
        
        list_frame = ttk.Frame(content)
        list_frame.grid(row=2, column=0, sticky="nsew", pady=(0, 20))
        list_frame.grid_rowconfigure(0, weight=1)
        list_frame.grid_columnconfigure(0, weight=1)
        
        # Listbox com estilo personalizado
        self.file_listbox = tk.Listbox(list_frame,
                                      bg=ModernColors.SURFACE,
                                      fg=ModernColors.TEXT_PRIMARY,
                                      selectbackground=ModernColors.PRIMARY_LIGHT,
                                      selectforeground=ModernColors.PRIMARY,
                                      borderwidth=1,
                                      relief='solid',
                                      font=('Segoe UI', 10, 'normal'),
                                      activestyle='none')
        self.file_listbox.grid(row=0, column=0, sticky="nsew")
        
        # Scrollbar moderna
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.file_listbox.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.file_listbox.configure(yscrollcommand=scrollbar.set)
        
        # Preview dos dados
        preview_label = ttk.Label(content, text="👁️ Preview dos Dados:",
                                 font=('Segoe UI', 12, 'bold'),
                                 foreground=ModernColors.TEXT_PRIMARY)
        preview_label.grid(row=3, column=0, sticky="w", pady=(0, 10))
        
        preview_frame = ttk.Frame(content)
        preview_frame.grid(row=4, column=0, sticky="nsew")
        preview_frame.grid_rowconfigure(0, weight=1)
        preview_frame.grid_columnconfigure(0, weight=1)
        
        # Treeview moderna para preview
        self.preview_tree = ttk.Treeview(preview_frame, style='Modern.Treeview')
        self.preview_tree.grid(row=0, column=0, sticky="nsew")
        
        # Scrollbars para preview
        h_scroll = ttk.Scrollbar(preview_frame, orient="horizontal", command=self.preview_tree.xview)
        h_scroll.grid(row=1, column=0, sticky="ew")
        self.preview_tree.configure(xscrollcommand=h_scroll.set)
        
        v_scroll = ttk.Scrollbar(preview_frame, orient="vertical", command=self.preview_tree.yview)
        v_scroll.grid(row=0, column=1, sticky="ns")
        self.preview_tree.configure(yscrollcommand=v_scroll.set)
    def create_modern_mapping_tab(self):
        """Aba de mapeamento de colunas com design moderno"""
        mapping_frame = ttk.Frame(self.notebook)
        self.notebook.add(mapping_frame, text="🔗  Mapeamento")
        
        # Card principal
        card, content = self.create_modern_card(mapping_frame,
                                               "Mapeamento de Colunas",
                                               "Configure quais colunas contêm as informações de atividade e tempo")
        card.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # Configurar grid
        content.grid_rowconfigure(4, weight=1)
        content.grid_columnconfigure(0, weight=1)
        
        # Frame para seleção com layout melhorado
        selection_frame = ttk.Frame(content)
        selection_frame.grid(row=0, column=0, pady=(0, 20), sticky="ew")
        selection_frame.grid_columnconfigure(1, weight=1)
        selection_frame.grid_columnconfigure(3, weight=1)
        selection_frame.grid_columnconfigure(5, weight=1)
        
        # Seleção de coluna de atividade
        ttk.Label(selection_frame, text="📊 Coluna de Atividade:",
                 font=('Segoe UI', 11, 'bold'),
                 foreground=ModernColors.TEXT_PRIMARY).grid(row=0, column=0, padx=(0, 10), sticky="w")
        
        self.activity_combo = ttk.Combobox(selection_frame, state="readonly", style='Modern.TCombobox', width=20)
        self.activity_combo.grid(row=0, column=1, padx=(0, 15), sticky="ew")
        
        # Seleção de coluna de tempo
        ttk.Label(selection_frame, text="⏱️ Coluna de Tempo:",
                 font=('Segoe UI', 11, 'bold'),
                 foreground=ModernColors.TEXT_PRIMARY).grid(row=0, column=2, padx=(0, 10), sticky="w")
        
        self.time_combo = ttk.Combobox(selection_frame, state="readonly", style='Modern.TCombobox', width=20)
        self.time_combo.grid(row=0, column=3, padx=(0, 15), sticky="ew")
        
        # Seleção de coluna de retrabalho (opcional)
        ttk.Label(selection_frame, text="🔄 Retrabalho (opcional):",
                 font=('Segoe UI', 11, 'bold'),
                 foreground=ModernColors.TEXT_PRIMARY).grid(row=0, column=4, padx=(0, 10), sticky="w")
        
        self.rework_combo = ttk.Combobox(selection_frame, state="readonly", style='Modern.TCombobox', width=20)
        self.rework_combo.grid(row=0, column=5, sticky="ew")
        
        # Adicionar tooltip para retrabalho
        rework_info = ttk.Label(selection_frame, text="ℹ️ Excluir linhas onde retrabalho = 1",
                               font=('Segoe UI', 9, 'italic'),
                               foreground=ModernColors.TEXT_SECONDARY)
        rework_info.grid(row=1, column=4, columnspan=2, sticky="w", pady=(2, 0))
        
        # Frame para botões
        button_frame = ttk.Frame(content)
        button_frame.grid(row=1, column=0, pady=(0, 20))
        
        # Botão de processar
        process_btn = ttk.Button(button_frame, text="⚙️ Processar Dados",
                                command=self.process_data, style='Primary.TButton')
        process_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # Botão de debug
        debug_btn = ttk.Button(button_frame, text="🔍 Debug Dados",
                              command=self.debug_data, style='Secondary.TButton')
        debug_btn.pack(side=tk.LEFT)
        
        # Frame para informações sobre retrabalho
        info_frame = ttk.LabelFrame(content, text="ℹ️ Informações sobre Retrabalho", style='Modern.TLabelframe')
        info_frame.grid(row=2, column=0, sticky="ew", pady=(0, 20))
        
        info_text = ttk.Label(info_frame,
                             text="• A coluna de retrabalho é opcional e funciona como filtro binário\n"
                                  "• Valores que representam retrabalho: 1, 1.0, True, true, TRUE\n"
                                  "• Linhas com retrabalho serão EXCLUÍDAS da análise\n"
                                  "• Valores vazios, 0, False são considerados como 'sem retrabalho'",
                             font=('Segoe UI', 10, 'normal'),
                             foreground=ModernColors.TEXT_SECONDARY,
                             justify=tk.LEFT)
        info_text.pack(anchor='w', padx=15, pady=10)
        
        # Preview dos dados processados
        preview_label = ttk.Label(content, text="✅ Dados Processados:",
                                 font=('Segoe UI', 12, 'bold'),
                                 foreground=ModernColors.TEXT_PRIMARY)
        preview_label.grid(row=3, column=0, sticky="w", pady=(0, 10))
        
        processed_frame = ttk.Frame(content)
        processed_frame.grid(row=4, column=0, sticky="nsew")
        processed_frame.grid_rowconfigure(0, weight=1)
        processed_frame.grid_columnconfigure(0, weight=1)
        
        self.processed_tree = ttk.Treeview(processed_frame, 
                                          columns=("Atividade", "Tempo"), 
                                          show="headings",
                                          style='Modern.Treeview')
        self.processed_tree.heading("Atividade", text="📋 Atividade")
        self.processed_tree.heading("Tempo", text="⏱️ Tempo (segundos)")
        self.processed_tree.grid(row=0, column=0, sticky="nsew")
        
        # Scrollbar
        processed_scroll = ttk.Scrollbar(processed_frame, orient="vertical", command=self.processed_tree.yview)
        processed_scroll.grid(row=0, column=1, sticky="ns")
        self.processed_tree.configure(yscrollcommand=processed_scroll.set)
        
        # Configurar grid weights
        content.grid_rowconfigure(4, weight=1)
    def create_modern_unification_tab(self):
        """Aba de unificação de atividades com design moderno"""
        unification_frame = ttk.Frame(self.notebook)
        self.notebook.add(unification_frame, text="🔄  Unificação")
        
        # Card principal
        card, content = self.create_modern_card(unification_frame,
                                               "Unificação de Atividades",
                                               "Detecte e unifique atividades similares para melhor análise")
        card.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # Configurar grid
        content.grid_rowconfigure(2, weight=1)
        content.grid_columnconfigure(0, weight=1)
        
        # Frame para botões com melhor layout
        buttons_frame = ttk.Frame(content)
        buttons_frame.grid(row=0, column=0, pady=(0, 20), sticky="ew")
        
        # Botão para detectar similaridades
        detect_btn = ttk.Button(buttons_frame, text="🔍 Detectar Similaridades",
                               command=self.detect_similarities, style='Primary.TButton')
        detect_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # Botão para atualizar
        refresh_btn = ttk.Button(buttons_frame, text="🔄 Atualizar Lista",
                                command=self.detect_similarities, style='Secondary.TButton')
        refresh_btn.pack(side=tk.LEFT)
        
        # Contador de similaridades
        self.similarity_count = ttk.Label(buttons_frame, text="",
                                         font=('Segoe UI', 10, 'normal'),
                                         foreground=ModernColors.TEXT_SECONDARY)
        self.similarity_count.pack(side=tk.RIGHT)
        
        # Label para sugestões
        suggestions_label = ttk.Label(content, text="🎯 Sugestões de Unificação:",
                                     font=('Segoe UI', 12, 'bold'),
                                     foreground=ModernColors.TEXT_PRIMARY)
        suggestions_label.grid(row=1, column=0, sticky="w", pady=(0, 10))
        
        # Frame para sugestões
        suggestions_frame = ttk.Frame(content)
        suggestions_frame.grid(row=2, column=0, sticky="nsew", pady=(0, 20))
        suggestions_frame.grid_rowconfigure(0, weight=1)
        suggestions_frame.grid_columnconfigure(0, weight=1)
        
        self.similarity_tree = ttk.Treeview(suggestions_frame, 
                                           columns=("Atividades", "Similaridade", "Ação"), 
                                           show="headings",
                                           style='Modern.Treeview')
        self.similarity_tree.heading("Atividades", text="🔗 Atividades Similares")
        self.similarity_tree.heading("Similaridade", text="📊 % Similaridade")
        self.similarity_tree.heading("Ação", text="⚡ Status")
        
        # Ajustar largura das colunas
        self.similarity_tree.column("Atividades", width=400)
        self.similarity_tree.column("Similaridade", width=120, anchor='center')
        self.similarity_tree.column("Ação", width=100, anchor='center')
        
        self.similarity_tree.grid(row=0, column=0, sticky="nsew")
        
        # Scrollbar
        similarity_scroll = ttk.Scrollbar(suggestions_frame, orient="vertical", command=self.similarity_tree.yview)
        similarity_scroll.grid(row=0, column=1, sticky="ns")
        self.similarity_tree.configure(yscrollcommand=similarity_scroll.set)
        
        # Frame para botões de ação
        action_frame = ttk.Frame(content)
        action_frame.grid(row=3, column=0, pady=(0, 10))
        
        unify_btn = ttk.Button(action_frame, text="✅ Unificar Selecionadas",
                              command=self.unify_activities, style='Primary.TButton')
        unify_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        skip_btn = ttk.Button(action_frame, text="⏭️ Pular Selecionadas",
                             command=self.skip_activities, style='Secondary.TButton')
        skip_btn.pack(side=tk.LEFT)
    def create_modern_grouping_tab(self):
        """Aba de agrupamento de atividades com design moderno"""
        grouping_frame = ttk.Frame(self.notebook)
        self.notebook.add(grouping_frame, text="📁  Agrupamento")
        
        # Card principal
        card, content = self.create_modern_card(grouping_frame,
                                               "Agrupamento de Atividades",
                                               "Organize atividades em grupos para análise categorizada")
        card.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # Configurar grid
        content.grid_rowconfigure(0, weight=1)
        content.grid_columnconfigure(0, weight=1)
        content.grid_columnconfigure(1, weight=1)
        
        # Frame esquerdo - Atividades disponíveis
        left_frame = ttk.LabelFrame(content, text="📋 Atividades Disponíveis", style='Modern.TLabelframe')
        left_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 10), pady=0)
        left_frame.grid_rowconfigure(1, weight=1)
        left_frame.grid_columnconfigure(0, weight=1)
        
        # Contador de atividades disponíveis
        self.available_count = ttk.Label(left_frame, text="",
                                        font=('Segoe UI', 10, 'normal'),
                                        foreground=ModernColors.TEXT_SECONDARY)
        self.available_count.grid(row=0, column=0, sticky="w", padx=10, pady=(10, 5))
        
        self.available_listbox = tk.Listbox(left_frame, 
                                           selectmode=tk.MULTIPLE,
                                           bg=ModernColors.SURFACE,
                                           fg=ModernColors.TEXT_PRIMARY,
                                           selectbackground=ModernColors.PRIMARY_LIGHT,
                                           selectforeground=ModernColors.PRIMARY,
                                           borderwidth=0,
                                           font=('Segoe UI', 10, 'normal'),
                                           activestyle='none')
        self.available_listbox.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0, 10))
        
        # Frame direito - Grupos
        right_frame = ttk.LabelFrame(content, text="📁 Grupos de Atividades", style='Modern.TLabelframe')
        right_frame.grid(row=0, column=1, sticky="nsew", padx=(10, 0), pady=0)
        right_frame.grid_rowconfigure(2, weight=1)
        right_frame.grid_columnconfigure(0, weight=1)
        
        # Contador de grupos
        self.groups_count = ttk.Label(right_frame, text="",
                                     font=('Segoe UI', 10, 'normal'),
                                     foreground=ModernColors.TEXT_SECONDARY)
        self.groups_count.grid(row=0, column=0, sticky="w", padx=10, pady=(10, 5))
        
        # Botões para grupos
        group_btn_frame = ttk.Frame(right_frame)
        group_btn_frame.grid(row=1, column=0, sticky="ew", padx=10, pady=(0, 10))
        
        new_group_btn = ttk.Button(group_btn_frame, text="➕ Novo Grupo",
                                  command=self.create_new_group, style='Primary.TButton')
        new_group_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        add_to_group_btn = ttk.Button(group_btn_frame, text="📝 Adicionar ao Grupo",
                                     command=self.add_to_group, style='Secondary.TButton')
        add_to_group_btn.pack(side=tk.LEFT)
        
        # Tree para grupos
        self.group_tree = ttk.Treeview(right_frame, style='Modern.Treeview')
        self.group_tree.grid(row=2, column=0, sticky="nsew", padx=10, pady=(0, 10))
        
        group_scroll = ttk.Scrollbar(right_frame, orient="vertical", command=self.group_tree.yview)
        group_scroll.grid(row=2, column=1, sticky="ns", pady=(0, 10))
        self.group_tree.configure(yscrollcommand=group_scroll.set)
    def create_modern_analysis_tab(self):
        """Aba de análise estatística com design moderno"""
        analysis_frame = ttk.Frame(self.notebook)
        self.notebook.add(analysis_frame, text="📊  Análise")
        
        # Card principal
        card, content = self.create_modern_card(analysis_frame,
                                               "Análise Estatística Avançada",
                                               "Visualize estatísticas detalhadas e métricas de performance")
        card.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # Configurar grid
        content.grid_rowconfigure(1, weight=1)
        content.grid_columnconfigure(0, weight=1)
        
        # Frame do cabeçalho para botões e status
        header_frame = ttk.Frame(content)
        header_frame.grid(row=0, column=0, sticky="ew", pady=(0, 20))
        header_frame.grid_columnconfigure(1, weight=1)
        
        # Botão de análise principal
        analyze_btn = ttk.Button(header_frame, text="🚀 Executar Análise Completa",
                                command=self.perform_analysis, style='Primary.TButton')
        analyze_btn.grid(row=0, column=0, sticky="w")
        
        # Status da análise
        self.analysis_status = ttk.Label(header_frame, text="Pronto para análise",
                                        font=('Segoe UI', 10, 'normal'),
                                        foreground=ModernColors.TEXT_SECONDARY)
        self.analysis_status.grid(row=0, column=1, sticky="e")
        
        # Frame da Tabela de resultados
        table_frame = ttk.LabelFrame(content, text="📈 Resultados Estatísticos Detalhados", style='Modern.TLabelframe')
        table_frame.grid(row=1, column=0, sticky="nsew")
        table_frame.grid_rowconfigure(0, weight=1)
        table_frame.grid_columnconfigure(0, weight=1)
        
        # Definir as colunas da tabela com a nova ordem
        cols = (
            "n", "std_dev", "min", "max", "median", "q1", "q3", "iqr", 
            "lower_fence", "upper_fence", "outlier_count", "non_outlier_count",
            "mean_all", "mean_no_outliers", "time_non_norm", "time_norm"
        )
        
        self.results_tree = ttk.Treeview(table_frame, columns=cols, show="tree headings", style='Modern.Treeview')
        
        # Definir os cabeçalhos das colunas e larguras com ícones
        col_headings = {
            "n": ("📊 N", 50), 
            "std_dev": ("📐 Desvio Padrão", 100), 
            "min": ("⬇️ Mínimo", 90), 
            "max": ("⬆️ Máximo", 90), 
            "median": ("🎯 Mediana", 90), 
            "q1": ("📊 Q1", 80), 
            "q3": ("📊 Q3", 80), 
            "iqr": ("📏 IQR", 80), 
            "lower_fence": ("🔻 Limite Inf.", 100),
            "upper_fence": ("🔺 Limite Sup.", 100), 
            "outlier_count": ("⚠️ Outliers", 80),
            "non_outlier_count": ("✅ Dentro Lim.", 90), 
            "mean_all": ("📊 Média Geral", 100),
            "mean_no_outliers": ("📈 Média s/ Out.", 100),
            "time_non_norm": ("⏱️ T. Não Norm. (min)", 130),
            "time_norm": ("✨ T. Norm. (min)", 120)
        }
        
        self.results_tree.heading("#0", text="🏷️ Atividade/Grupo")
        self.results_tree.column("#0", width=220, stretch=tk.NO)
        
        for col, (heading, width) in col_headings.items():
            self.results_tree.heading(col, text=heading)
            self.results_tree.column(col, width=width, anchor="center")

        self.results_tree.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        
        # Frame para scrollbars
        scroll_frame = ttk.Frame(table_frame)
        scroll_frame.grid(row=0, column=1, rowspan=2, sticky="ns")
        
        # Scrollbar Vertical
        results_v_scroll = ttk.Scrollbar(scroll_frame, orient="vertical", command=self.results_tree.yview)
        results_v_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.results_tree.configure(yscrollcommand=results_v_scroll.set)
        
        # Scrollbar Horizontal
        results_h_scroll = ttk.Scrollbar(table_frame, orient="horizontal", command=self.results_tree.xview)
        results_h_scroll.grid(row=1, column=0, sticky="ew", padx=10)
        self.results_tree.configure(xscrollcommand=results_h_scroll.set)

    def create_modern_export_tab(self):
        """Aba de exportação com design moderno"""
        export_frame = ttk.Frame(self.notebook)
        self.notebook.add(export_frame, text="📤  Exportação")
        
        # Card principal
        card, content = self.create_modern_card(export_frame,
                                               "Exportação de Resultados",
                                               "Exporte seus dados e análises em diferentes formatos")
        card.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # Configurar grid
        content.grid_rowconfigure(1, weight=1)
        content.grid_columnconfigure(0, weight=1)
        
        # Frame para opções de exportação
        options_frame = ttk.Frame(content)
        options_frame.grid(row=0, column=0, pady=(0, 30))
        
        # Cards para cada tipo de exportação
        excel_card = ttk.Frame(options_frame, style='Card.TFrame', relief='solid', borderwidth=1)
        excel_card.grid(row=0, column=0, padx=(0, 20), pady=10, sticky="nsew")
        
        # Conteúdo do card Excel
        excel_content = ttk.Frame(excel_card)
        excel_content.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        excel_icon = ttk.Label(excel_content, text="📊", font=('Segoe UI', 24, 'normal'),
                              foreground=ModernColors.SUCCESS)
        excel_icon.pack(pady=(0, 10))
        
        excel_title = ttk.Label(excel_content, text="Excel Completo",
                               font=('Segoe UI', 14, 'bold'),
                               foreground=ModernColors.TEXT_PRIMARY)
        excel_title.pack()
        
        excel_desc = ttk.Label(excel_content, text="Exporta dados formatados\ncom análise estatística",
                              font=('Segoe UI', 10, 'normal'),
                              foreground=ModernColors.TEXT_SECONDARY,
                              justify=tk.CENTER)
        excel_desc.pack(pady=(5, 15))
        
        export_excel_btn = ttk.Button(excel_content, text="📊 Exportar Excel",
                                     command=self.export_to_excel, style='Primary.TButton')
        export_excel_btn.pack()
        
        # Card CSV
        csv_card = ttk.Frame(options_frame, style='Card.TFrame', relief='solid', borderwidth=1)
        csv_card.grid(row=0, column=1, padx=(20, 0), pady=10, sticky="nsew")
        
        # Conteúdo do card CSV
        csv_content = ttk.Frame(csv_card)
        csv_content.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        csv_icon = ttk.Label(csv_content, text="📄", font=('Segoe UI', 24, 'normal'),
                            foreground=ModernColors.PRIMARY)
        csv_icon.pack(pady=(0, 10))
        
        csv_title = ttk.Label(csv_content, text="CSV Simples",
                             font=('Segoe UI', 14, 'bold'),
                             foreground=ModernColors.TEXT_PRIMARY)
        csv_title.pack()
        
        csv_desc = ttk.Label(csv_content, text="Exporta dados brutos\nem formato CSV",
                            font=('Segoe UI', 10, 'normal'),
                            foreground=ModernColors.TEXT_SECONDARY,
                            justify=tk.CENTER)
        csv_desc.pack(pady=(5, 15))
        
        export_csv_btn = ttk.Button(csv_content, text="📄 Exportar CSV",
                                   command=self.export_to_csv, style='Secondary.TButton')
        export_csv_btn.pack()
        
        # Status da exportação
        status_frame = ttk.Frame(content)
        status_frame.grid(row=1, column=0, sticky="ew")
        
        ttk.Label(status_frame, text="📋 Status:",
                 font=('Segoe UI', 12, 'bold'),
                 foreground=ModernColors.TEXT_PRIMARY).pack(anchor='w', pady=(0, 5))
        
        self.export_status = ttk.Label(status_frame, text="✅ Pronto para exportar",
                                      font=('Segoe UI', 11, 'normal'),
                                      foreground=ModernColors.SUCCESS)
        self.export_status.pack(anchor='w')
        
    def upload_files(self):
        """Função para upload de arquivos com feedback visual melhorado"""
        filetypes = [
            ("Arquivos Excel", "*.xlsx"),
            ("Arquivos CSV", "*.csv"),
            ("Todos os arquivos", "*.*")
        ]
        
        files = filedialog.askopenfilenames(
            title="🗂️ Selecionar arquivos para análise",
            filetypes=filetypes
        )
        
        new_files = 0
        for file_path in files:
            if file_path not in self.uploaded_files:
                self.uploaded_files.append(file_path)
                filename = os.path.basename(file_path)
                self.file_listbox.insert(tk.END, f"📄 {filename}")
                new_files += 1
        
        # Atualizar status
        total_files = len(self.uploaded_files)
        if total_files == 0:
            self.upload_status.config(text="Nenhum arquivo selecionado", foreground=ModernColors.TEXT_SECONDARY)
        elif new_files > 0:
            self.upload_status.config(text=f"✅ {total_files} arquivo(s) carregado(s) (+{new_files} novo(s))", 
                                    foreground=ModernColors.SUCCESS)
        else:
            self.upload_status.config(text=f"✅ {total_files} arquivo(s) carregado(s)", 
                                    foreground=ModernColors.SUCCESS)
                
        self.update_preview()
        self.update_column_combos()
        
    def clear_files(self):
        """Limpar lista de arquivos com confirmação"""
        if self.uploaded_files:
            result = messagebox.askyesno("🗑️ Confirmar", 
                                       "Deseja realmente limpar todos os arquivos selecionados?",
                                       icon='question')
            if result:
                self.uploaded_files.clear()
                self.file_listbox.delete(0, tk.END)
                self.upload_status.config(text="Nenhum arquivo selecionado", foreground=ModernColors.TEXT_SECONDARY)
                self.clear_preview()
        else:
            messagebox.showinfo("ℹ️ Informação", "Não há arquivos para limpar.")
            
    def update_preview(self):
        """Atualizar preview dos dados com melhor tratamento de erros"""
        if not self.uploaded_files:
            self.clear_preview()
            return
            
        try:
            # Ler primeiro arquivo para preview
            file_path = self.uploaded_files[0]
            filename = os.path.basename(file_path)
            
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
                self.preview_tree.heading(col, text=f"📊 {col}")
                self.preview_tree.column(col, width=120, anchor='center')
                
            # Adicionar dados (primeiras 10 linhas)
            for index, row in df.head(10).iterrows():
                values = []
                for val in row:
                    if pd.isna(val):
                        values.append("(vazio)")
                    else:
                        values.append(str(val)[:50])  # Limitar tamanho
                self.preview_tree.insert("", tk.END, values=values)
                
        except Exception as e:
            messagebox.showerror("❌ Erro", f"Erro ao ler arquivo {filename}:\n{str(e)}")
            
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
                
            columns = [f"📊 {col}" for col in df.columns]
            self.activity_combo['values'] = columns
            self.time_combo['values'] = columns
            
            # Para a coluna de retrabalho, adicionar uma opção vazia no início
            rework_columns = [""] + columns  # Primeira opção vazia para não selecionar
            self.rework_combo['values'] = rework_columns
            self.rework_combo.set("")  # Deixar vazio por padrão
            
        except Exception as e:
            messagebox.showerror("❌ Erro", f"Erro ao ler colunas: {str(e)}")

    def process_data(self):
        """Processar dados dos arquivos com melhor feedback visual e filtro de retrabalho"""
        if not self.uploaded_files:
            messagebox.showwarning("⚠️ Aviso", "Nenhum arquivo selecionado")
            return
            
        activity_col_display = self.activity_combo.get()
        time_col_display = self.time_combo.get()
        rework_col_display = self.rework_combo.get()
        
        if not activity_col_display or not time_col_display:
            messagebox.showwarning("⚠️ Aviso", "Selecione as colunas de atividade e tempo")
            return
        
        # Remover ícones dos nomes das colunas
        activity_col = activity_col_display.replace("📊 ", "")
        time_col = time_col_display.replace("📊 ", "")
        rework_col = rework_col_display.replace("📊 ", "") if rework_col_display else None
            
        try:
            all_data = []
            total_rows_read = 0
            files_processed = 0
            rework_filtered_count = 0  # Contador de linhas filtradas por retrabalho
            
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
                    
                    # Verificar se as colunas obrigatórias existem
                    if activity_col in df.columns and time_col in df.columns:
                        # Selecionar colunas necessárias
                        cols_to_select = [activity_col, time_col]
                        if rework_col and rework_col in df.columns:
                            cols_to_select.append(rework_col)
                            data = df[cols_to_select].copy()
                            data.columns = ['Atividade', 'Tempo', 'Retrabalho']
                        else:
                            data = df[cols_to_select].copy()
                            data.columns = ['Atividade', 'Tempo']
                        
                        # Aplicar filtro de retrabalho ANTES da limpeza geral
                        if rework_col and rework_col in df.columns:
                            before_filter = len(data)
                            # Filtrar linhas onde retrabalho é 1 (excluir essas linhas)
                            # Considerar também valores como "1", 1.0, True como retrabalho
                            data['Retrabalho'] = data['Retrabalho'].astype(str).str.strip()
                            rework_mask = data['Retrabalho'].isin(['1', '1.0', 'True', 'true', 'TRUE'])
                            data = data[~rework_mask]  # Manter apenas as que NÃO são retrabalho
                            after_filter = len(data)
                            rework_filtered_count += (before_filter - after_filter)
                            
                            # Remover a coluna de retrabalho após o filtro
                            data = data[['Atividade', 'Tempo']]
                        
                        # Limpeza de dados padrão
                        data.dropna(subset=['Atividade', 'Tempo'], how='any', inplace=True)
                        data['Atividade'] = data['Atividade'].astype(str).str.strip()
                        data = data[data['Atividade'] != '']
                        
                        if not data.empty:
                            all_data.append(data)
                            files_processed += 1
                    else:
                        missing_cols = []
                        if activity_col not in df.columns:
                            missing_cols.append(activity_col)
                        if time_col not in df.columns:
                            missing_cols.append(time_col)
                        print(f"Colunas não encontradas no arquivo {os.path.basename(file_path)}: {missing_cols}")

                except Exception as e:
                    print(f"Erro ao processar arquivo {os.path.basename(file_path)}: {str(e)}")
                    continue
            
            if not all_data:
                messagebox.showerror("❌ Erro", "Nenhum dado válido pôde ser extraído dos arquivos.\nVerifique se as colunas selecionadas estão corretas e contêm dados.")
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
                
                # Mensagem de sucesso com informações sobre retrabalho
                success_message = (f"Dados processados com sucesso!\n\n"
                                 f"📁 {files_processed} arquivo(s) processado(s)\n"
                                 f"📊 {total_rows_read} linha(s) lida(s) no total\n")
                
                if rework_col:
                    success_message += f"🔄 {rework_filtered_count} linha(s) excluída(s) por retrabalho\n"
                
                success_message += (f"✅ {final_count} registro(s) válido(s) para análise\n"
                                  f"🗑️ {removed_count} registro(s) removido(s) por serem inválidos")
                
                messagebox.showinfo("✅ Sucesso", success_message)
            else:
                 messagebox.showerror("❌ Erro", "Nenhum dado válido encontrado após a limpeza e conversão.\nVerifique o formato dos dados nas colunas de tempo.")

        except Exception as e:
            messagebox.showerror("❌ Erro", f"Erro geral ao processar dados:\n{str(e)}")
            print(f"Erro detalhado: {str(e)}")

    def update_processed_preview(self):
        """Atualizar preview dos dados processados com ícones"""
        for item in self.processed_tree.get_children():
            self.processed_tree.delete(item)
            
        if self.processed_data is not None:
            for index, row in self.processed_data.head(20).iterrows():
                self.processed_tree.insert("", tk.END, values=(f"📋 {row['Atividade']}", f"⏱️ {row['Tempo']:.2f}s"))
                
    def update_available_activities(self):
        """Atualizar lista de atividades disponíveis com contadores"""
        self.available_listbox.delete(0, tk.END)
        
        if self.processed_data is not None:
            unique_activities = self.processed_data['Atividade'].unique()
            
            # Atividades que já estão em grupos
            grouped_activities = set()
            for group_data in self.activity_groups.values():
                grouped_activities.update(group_data['activities'])
            
            # Adicionar apenas atividades que não estão em grupos
            available_count = 0
            for activity in sorted(unique_activities):
                if activity not in grouped_activities:
                    # Contar ocorrências da atividade
                    count = len(self.processed_data[self.processed_data['Atividade'] == activity])
                    self.available_listbox.insert(tk.END, f"📊 {activity} ({count})")
                    available_count += 1
            
            # Atualizar contador na interface
            if hasattr(self, 'available_count'):
                self.available_count.config(text=f"{available_count} atividade(s) disponível(eis)")
            
            # Atualizar contador de grupos
            if hasattr(self, 'groups_count'):
                group_count = len(self.activity_groups)
                total_grouped = sum(len(data['activities']) for data in self.activity_groups.values())
                self.groups_count.config(text=f"{group_count} grupo(s) • {total_grouped} atividade(s) agrupada(s)")
                    
    def detect_similarities(self):
        """Detectar atividades similares com melhor feedback"""
        if self.processed_data is None:
            messagebox.showwarning("⚠️ Aviso", "Processe os dados primeiro")
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
            # Adicionar contagem de ocorrências
            count1 = len(self.processed_data[self.processed_data['Atividade'] == activity1])
            count2 = len(self.processed_data[self.processed_data['Atividade'] == activity2])
            
            display_text = f"{activity1} ({count1}) ↔ {activity2} ({count2})"
            
            self.similarity_tree.insert("", tk.END, values=(
                display_text,
                f"{similarity:.1%}",
                "⏳ Pendente"
            ))
        
        # Atualizar contador
        if hasattr(self, 'similarity_count'):
            if similarities:
                self.similarity_count.config(text=f"🎯 {len(similarities)} similaridade(s) detectada(s)",
                                            foreground=ModernColors.SUCCESS)
            else:
                self.similarity_count.config(text="ℹ️ Nenhuma similaridade encontrada",
                                           foreground=ModernColors.TEXT_SECONDARY)
            
        if not similarities:
            messagebox.showinfo("ℹ️ Informação", "Nenhuma atividade similar encontrada com 70% ou mais de similaridade")
            
    def calculate_similarity(self, str1, str2):
        """Calcular similaridade entre duas strings"""
        return SequenceMatcher(None, str1.lower(), str2.lower()).ratio()
        
    def unify_activities(self):
        """Unificar atividades selecionadas com interface melhorada"""
        selected_items = self.similarity_tree.selection()
        if not selected_items:
            messagebox.showwarning("⚠️ Aviso", "Selecione atividades para unificar")
            return
            
        try:
            for item in selected_items:
                # Obter dados do item selecionado
                item_data = self.similarity_tree.item(item)
                activities_text = item_data['values'][0]
                
                # Extrair as duas atividades (removendo contadores)
                if '↔' in activities_text:
                    parts = activities_text.split(' ↔ ')
                    # Remover contadores no formato (N)
                    activity1 = parts[0].split(' (')[0]
                    activity2 = parts[1].split(' (')[0]
                    
                    # Criar janela de escolha modernizada
                    choice_window = tk.Toplevel(self.root)
                    choice_window.title("🔄 Unificar Atividades")
                    choice_window.geometry("500x350")
                    choice_window.configure(bg=ModernColors.SURFACE)
                    choice_window.transient(self.root)
                    choice_window.grab_set()
                    
                    # Header
                    header_frame = ttk.Frame(choice_window)
                    header_frame.pack(fill=tk.X, padx=20, pady=(20, 10))
                    
                    title_label = ttk.Label(header_frame, text="🔄 Unificar Atividades Similares",
                                           font=('Segoe UI', 16, 'bold'),
                                           foreground=ModernColors.TEXT_PRIMARY)
                    title_label.pack()
                    
                    subtitle_label = ttk.Label(header_frame, text="Escolha o nome para a atividade unificada:",
                                              font=('Segoe UI', 11, 'normal'),
                                              foreground=ModernColors.TEXT_SECONDARY)
                    subtitle_label.pack(pady=(5, 0))
                    
                    # Separator
                    separator = ttk.Separator(choice_window, orient='horizontal')
                    separator.pack(fill=tk.X, padx=20, pady=10)
                    
                    # Options frame
                    options_frame = ttk.Frame(choice_window)
                    options_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
                    
                    var = tk.StringVar(value=activity1)
                    
                    # Opção 1
                    option1_frame = ttk.Frame(options_frame, relief='solid', borderwidth=1)
                    option1_frame.pack(fill=tk.X, pady=(0, 10))
                    
                    radio1 = ttk.Radiobutton(option1_frame, text="", variable=var, value=activity1)
                    radio1.pack(side=tk.LEFT, padx=10, pady=10)
                    
                    label1 = ttk.Label(option1_frame, text=f"📋 {activity1}",
                                      font=('Segoe UI', 11, 'normal'),
                                      foreground=ModernColors.TEXT_PRIMARY)
                    label1.pack(side=tk.LEFT, pady=10)
                    
                    # Opção 2
                    option2_frame = ttk.Frame(options_frame, relief='solid', borderwidth=1)
                    option2_frame.pack(fill=tk.X, pady=(0, 10))
                    
                    radio2 = ttk.Radiobutton(option2_frame, text="", variable=var, value=activity2)
                    radio2.pack(side=tk.LEFT, padx=10, pady=10)
                    
                    label2 = ttk.Label(option2_frame, text=f"📋 {activity2}",
                                      font=('Segoe UI', 11, 'normal'),
                                      foreground=ModernColors.TEXT_PRIMARY)
                    label2.pack(side=tk.LEFT, pady=10)
                    
                    # Opção customizada
                    custom_frame = ttk.Frame(options_frame, relief='solid', borderwidth=1)
                    custom_frame.pack(fill=tk.X, pady=(0, 20))
                    
                    radio3 = ttk.Radiobutton(custom_frame, text="", variable=var, value="custom")
                    radio3.pack(side=tk.LEFT, padx=10, pady=10)
                    
                    custom_label = ttk.Label(custom_frame, text="✏️ Nome personalizado:",
                                           font=('Segoe UI', 11, 'normal'),
                                           foreground=ModernColors.TEXT_PRIMARY)
                    custom_label.pack(side=tk.LEFT, pady=10)
                    
                    custom_entry = ttk.Entry(custom_frame, width=30, style='Modern.TEntry')
                    custom_entry.pack(side=tk.RIGHT, padx=10, pady=10)
                    
                    def confirm_unification():
                        chosen_name = var.get()
                        if chosen_name == "custom":
                            chosen_name = custom_entry.get().strip()
                            if not chosen_name:
                                messagebox.showerror("❌ Erro", "Digite um nome personalizado")
                                return
                        
                        # Realizar unificação
                        self.processed_data.loc[self.processed_data['Atividade'] == activity2, 'Atividade'] = chosen_name
                        if activity1 != chosen_name:
                            self.processed_data.loc[self.processed_data['Atividade'] == activity1, 'Atividade'] = chosen_name
                        
                        # Armazenar unificação
                        self.unified_activities[activity1] = chosen_name
                        self.unified_activities[activity2] = chosen_name
                        
                        # Atualizar status do item
                        self.similarity_tree.set(item, "Ação", "✅ Unificada")
                        
                        choice_window.destroy()
                        
                        # Atualizar interfaces
                        self.update_processed_preview()
                        self.update_available_activities()
                        
                        messagebox.showinfo("✅ Sucesso", f"Atividades unificadas como:\n📋 {chosen_name}")
                    
                    # Botões
                    button_frame = ttk.Frame(choice_window)
                    button_frame.pack(fill=tk.X, padx=20, pady=(0, 20))
                    
                    cancel_btn = ttk.Button(button_frame, text="❌ Cancelar",
                                           command=choice_window.destroy, style='Secondary.TButton')
                    cancel_btn.pack(side=tk.RIGHT, padx=(10, 0))
                    
                    confirm_btn = ttk.Button(button_frame, text="✅ Confirmar Unificação",
                                           command=confirm_unification, style='Primary.TButton')
                    confirm_btn.pack(side=tk.RIGHT)
                    
                    # Centralizar janela
                    choice_window.update_idletasks()
                    x = (self.root.winfo_screenwidth() // 2) - (choice_window.winfo_width() // 2)
                    y = (self.root.winfo_screenheight() // 2) - (choice_window.winfo_height() // 2)
                    choice_window.geometry(f"+{x}+{y}")
                    
        except Exception as e:
            messagebox.showerror("❌ Erro", f"Erro ao unificar atividades:\n{str(e)}")
    
    def skip_activities(self):
        """Pular atividades selecionadas"""
        selected_items = self.similarity_tree.selection()
        if not selected_items:
            messagebox.showwarning("⚠️ Aviso", "Selecione atividades para pular")
            return
            
        try:
            for item in selected_items:
                # Atualizar status do item
                self.similarity_tree.set(item, "Ação", "⏭️ Pulada")
                
            messagebox.showinfo("ℹ️ Informação", f"✅ {len(selected_items)} sugestão(ões) de unificação pulada(s)")
            
        except Exception as e:
            messagebox.showerror("❌ Erro", f"Erro ao pular atividades:\n{str(e)}")
        
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
        """Atualizar árvore de grupos com ícones e contadores"""
        # Limpar árvore
        for item in self.group_tree.get_children():
            self.group_tree.delete(item)
        
        # Adicionar grupos e suas atividades
        for group_name, group_data in self.activity_groups.items():
            activity_count = len(group_data['activities'])
            # Adicionar grupo principal com contador
            group_item = self.group_tree.insert("", tk.END, 
                                               text=f"📁 {group_name} ({activity_count} atividade(s))", 
                                               values=(), open=True)
            
            # Adicionar atividades do grupo
            for activity in group_data['activities']:
                # Contar ocorrências da atividade se os dados estão processados
                count_text = ""
                if self.processed_data is not None:
                    count = len(self.processed_data[self.processed_data['Atividade'] == activity])
                    count_text = f" ({count})"
                
                self.group_tree.insert(group_item, tk.END, 
                                     text=f"  📊 {activity}{count_text}", values=())
        
        # Atualizar contadores
        self.update_available_activities()

    def perform_analysis(self):
        """Executar análise estatística com feedback visual melhorado"""
        if self.processed_data is None or self.processed_data.empty:
            messagebox.showwarning("⚠️ Aviso", "Não há dados válidos para analisar. Processe os arquivos primeiro.")
            return

        try:
            # Atualizar status
            if hasattr(self, 'analysis_status'):
                self.analysis_status.config(text="🔄 Executando análise...", foreground=ModernColors.WARNING)
            
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
                            group_item = self.results_tree.insert("", tk.END, 
                                                                 text=f"📁 {group_name}", 
                                                                 values=metrics, open=True)
                            # Analisar atividades individuais do grupo
                            for activity in group_data['activities']:
                                activity_times = self.processed_data[self.processed_data['Atividade'] == activity]['Tempo']
                                activity_metrics = self._calculate_metrics(activity_times)
                                if activity_metrics:
                                    self.results_tree.insert(group_item, tk.END, 
                                                            text=f"  📊 {activity}", 
                                                            values=activity_metrics)

            # Analisar atividades não agrupadas
            grouped_activities = {act for group in self.activity_groups.values() for act in group['activities']}
            ungrouped_df = self.processed_data[~self.processed_data['Atividade'].isin(grouped_activities)]

            if not ungrouped_df.empty:
                # Criar um nó pai para atividades não agrupadas, se houver grupos.
                parent_item = ""
                if self.activity_groups:
                    parent_item = self.results_tree.insert("", tk.END, 
                                                          text="📋 Atividades Não Agrupadas", 
                                                          open=True)

                for activity, times in ungrouped_df.groupby('Atividade')['Tempo']:
                    metrics = self._calculate_metrics(times)
                    if metrics:
                        self.results_tree.insert(parent_item, tk.END, 
                                                text=f"📊 {activity}", 
                                                values=metrics)

            # Atualizar status de sucesso
            if hasattr(self, 'analysis_status'):
                total_items = len(self.results_tree.get_children())
                self.analysis_status.config(text=f"✅ Análise concluída • {total_items} item(s) analisado(s)", 
                                          foreground=ModernColors.SUCCESS)
            
            messagebox.showinfo("✅ Análise Concluída", 
                               "A análise estatística foi concluída com sucesso!\n\n"
                               "📊 Todos os dados foram processados e as métricas calculadas.")

        except Exception as e:
            import traceback
            error_details = traceback.format_exc()
            print(f"Erro detalhado na análise: {error_details}")
            
            if hasattr(self, 'analysis_status'):
                self.analysis_status.config(text="❌ Erro na análise", foreground=ModernColors.ERROR)
            
            messagebox.showerror("❌ Erro", f"Erro na análise estatística:\n{str(e)}\n\nVerifique o console para mais detalhes.")
            
    def export_to_excel(self):
        """Exportar resultados para Excel com feedback melhorado"""
        if self.processed_data is None:
            messagebox.showwarning("⚠️ Aviso", "Nenhum dado para exportar. Execute a análise primeiro.")
            return

        if not self.results_tree.get_children():
            messagebox.showwarning("⚠️ Aviso", "Nenhuma análise encontrada para exportar. Execute a análise primeiro.")
            return

        try:
            filename = filedialog.asksaveasfilename(
                title="💾 Salvar arquivo Excel",
                defaultextension=".xlsx",
                filetypes=[("Arquivos Excel", "*.xlsx"), ("Todos os Arquivos", "*.*")]
            )
            if not filename:
                return

            # Atualizar status
            self.export_status.config(text="🔄 Exportando para Excel...", foreground=ModernColors.WARNING)

            # --- Parte 1: Preparar dados brutos pivotados ---
            # Mapeamento de atividade para grupo
            activity_to_group = {activity: group_name
                                 for group_name, data in self.activity_groups.items()
                                 for activity in data['activities']}

            df_part1_source = self.processed_data.copy()
            df_part1_source['Processos'] = df_part1_source['Atividade'].map(activity_to_group).fillna('')
            
            # Agrupar tempos em listas por atividade
            pivoted_times = df_part1_source.groupby(['Processos', 'Atividade'])['Tempo'].apply(list).reset_index(name='Tempos')
            
            # Expandir as listas de tempo em colunas "Amostra N"
            max_samples = 0
            if not pivoted_times.empty:
                max_samples_val = pivoted_times['Tempos'].str.len().max()
                if pd.notna(max_samples_val):
                    max_samples = int(max_samples_val)

            sample_cols = [f'Amostra {i+1}' for i in range(max_samples)]
            
            time_df = pd.DataFrame(pivoted_times['Tempos'].tolist(), index=pivoted_times.index, columns=sample_cols)
            
            # Juntar as informações com os tempos expandidos
            part1_df = pd.concat([pivoted_times[['Processos', 'Atividade']], time_df], axis=1)
            
            # Garantir que todas as atividades definidas nos grupos apareçam
            all_grouped_activities = {act for data in self.activity_groups.values() for act in data['activities']}
            
            exported_activities = set()
            if not part1_df.empty:
                exported_activities = set(part1_df['Atividade'])

            missing_activities = all_grouped_activities - exported_activities
            
            if missing_activities:
                missing_data = []
                for activity in sorted(list(missing_activities)):
                    group_name = activity_to_group.get(activity, '')
                    row = {'Processos': group_name, 'Atividade': activity}
                    missing_data.append(row)
                
                if missing_data:
                    missing_df = pd.DataFrame(missing_data)
                    part1_df = pd.concat([part1_df, missing_df], ignore_index=True)
            
            # Ordenar para manter os grupos juntos
            if not part1_df.empty:
                 part1_df = part1_df.sort_values(by=['Processos', 'Atividade']).reset_index(drop=True)

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

            self.export_status.config(text=f"✅ Exportado com sucesso: {os.path.basename(filename)}", 
                                    foreground=ModernColors.SUCCESS)
            messagebox.showinfo("✅ Sucesso", 
                               f"Dados exportados com sucesso!\n\n"
                               f"📊 Arquivo: {os.path.basename(filename)}\n"
                               f"📁 Local: {os.path.dirname(filename)}")

        except Exception as e:
            import traceback
            error_details = traceback.format_exc()
            print(f"Erro detalhado na exportação: {error_details}")
            self.export_status.config(text="❌ Erro na exportação", foreground=ModernColors.ERROR)
            messagebox.showerror("❌ Erro", f"Erro ao exportar para Excel:\n{str(e)}")
            
    def export_to_csv(self):
        """Exportar resultados para CSV com feedback melhorado"""
        if self.processed_data is None:
            messagebox.showwarning("⚠️ Aviso", "Nenhum dado para exportar")
            return
            
        try:
            filename = filedialog.asksaveasfilename(
                title="💾 Salvar arquivo CSV",
                defaultextension=".csv",
                filetypes=[("Arquivos CSV", "*.csv"), ("Todos os Arquivos", "*.*")]
            )
            
            if filename:
                self.export_status.config(text="🔄 Exportando para CSV...", foreground=ModernColors.WARNING)
                
                # Exportação CSV mantém o formato simples dos dados processados
                self.processed_data.to_csv(filename, index=False)
                
                self.export_status.config(text=f"✅ Exportado com sucesso: {os.path.basename(filename)}", 
                                        foreground=ModernColors.SUCCESS)
                messagebox.showinfo("✅ Sucesso", 
                                   f"Dados brutos exportados com sucesso!\n\n"
                                   f"📄 Arquivo: {os.path.basename(filename)}\n"
                                   f"📁 Local: {os.path.dirname(filename)}")
                
        except Exception as e:
            self.export_status.config(text="❌ Erro na exportação", foreground=ModernColors.ERROR)
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

    def create_new_group(self):
        """Criar novo grupo com interface modernizada"""
        # Janela para criar novo grupo
        group_window = tk.Toplevel(self.root)
        group_window.title("➕ Criar Novo Grupo")
        group_window.geometry("450x300")
        group_window.configure(bg=ModernColors.SURFACE)
        group_window.transient(self.root)
        group_window.grab_set()
        
        # Header
        header_frame = ttk.Frame(group_window)
        header_frame.pack(fill=tk.X, padx=20, pady=(20, 10))
        
        title_label = ttk.Label(header_frame, text="➕ Criar Novo Grupo",
                               font=('Segoe UI', 16, 'bold'),
                               foreground=ModernColors.TEXT_PRIMARY)
        title_label.pack()
        
        subtitle_label = ttk.Label(header_frame, text="Organize suas atividades em grupos personalizados",
                                  font=('Segoe UI', 11, 'normal'),
                                  foreground=ModernColors.TEXT_SECONDARY)
        subtitle_label.pack(pady=(5, 0))
        
        # Separator
        separator = ttk.Separator(group_window, orient='horizontal')
        separator.pack(fill=tk.X, padx=20, pady=10)
        
        # Content frame
        content_frame = ttk.Frame(group_window)
        content_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
        
        # Nome do grupo
        name_label = ttk.Label(content_frame, text="📝 Nome do Grupo:",
                              font=('Segoe UI', 12, 'bold'),
                              foreground=ModernColors.TEXT_PRIMARY)
        name_label.pack(anchor='w', pady=(0, 5))
        
        name_entry = ttk.Entry(content_frame, width=40, style='Modern.TEntry',
                              font=('Segoe UI', 11, 'normal'))
        name_entry.pack(fill=tk.X, pady=(0, 20))
        name_entry.focus()
        
        # Cor do grupo
        color_label = ttk.Label(content_frame, text="🎨 Cor do Grupo:",
                               font=('Segoe UI', 12, 'bold'),
                               foreground=ModernColors.TEXT_PRIMARY)
        color_label.pack(anchor='w', pady=(0, 10))
        
        color_var = tk.StringVar(value=ModernColors.PRIMARY)
        color_frame = ttk.Frame(content_frame)
        color_frame.pack(pady=(0, 20))
        
        colors = [
            ModernColors.PRIMARY, ModernColors.SUCCESS, ModernColors.WARNING, 
            ModernColors.ERROR, ModernColors.ACCENT, "#6366F1"
        ]
        color_names = ["Azul", "Verde", "Laranja", "Vermelho", "Roxo", "Índigo"]
        
        for i, (color, name) in enumerate(zip(colors, color_names)):
            btn = tk.Button(color_frame, bg=color, width=8, height=2,
                           relief='solid', borderwidth=2,
                           command=lambda c=color: color_var.set(c))
            btn.grid(row=0, column=i, padx=5)
            
            # Label do nome da cor
            name_label = ttk.Label(color_frame, text=name, font=('Segoe UI', 8, 'normal'),
                                  foreground=ModernColors.TEXT_SECONDARY)
            name_label.grid(row=1, column=i, pady=(2, 0))
        
        def confirm_group():
            name = name_entry.get().strip()
            if not name:
                messagebox.showerror("❌ Erro", "Digite um nome para o grupo")
                return
            
            if name in self.activity_groups:
                messagebox.showerror("❌ Erro", "Já existe um grupo com este nome")
                return
            
            # Criar grupo
            self.activity_groups[name] = {
                'color': color_var.get(),
                'activities': []
            }
            
            # Atualizar interface
            self.update_group_tree()
            group_window.destroy()
            messagebox.showinfo("✅ Sucesso", f"Grupo '{name}' criado com sucesso! 🎉")
        
        # Botões
        button_frame = ttk.Frame(group_window)
        button_frame.pack(fill=tk.X, padx=20, pady=(0, 20))
        
        cancel_btn = ttk.Button(button_frame, text="❌ Cancelar",
                               command=group_window.destroy, style='Secondary.TButton')
        cancel_btn.pack(side=tk.RIGHT, padx=(10, 0))
        
        create_btn = ttk.Button(button_frame, text="✅ Criar Grupo",
                               command=confirm_group, style='Primary.TButton')
        create_btn.pack(side=tk.RIGHT)
        
        # Bind Enter key
        group_window.bind('<Return>', lambda e: confirm_group())
        
        # Centralizar janela
        group_window.update_idletasks()
        x = (self.root.winfo_screenwidth() // 2) - (group_window.winfo_width() // 2)
        y = (self.root.winfo_screenheight() // 2) - (group_window.winfo_height() // 2)
        group_window.geometry(f"+{x}+{y}")
        
    def add_to_group(self):
        """Adicionar atividades ao grupo com interface melhorada"""
        # Verificar se há atividades selecionadas
        selected_indices = self.available_listbox.curselection()
        if not selected_indices:
            messagebox.showwarning("⚠️ Aviso", "Selecione atividades para adicionar ao grupo")
            return
        
        # Verificar se há grupos criados
        if not self.activity_groups:
            messagebox.showwarning("⚠️ Aviso", "Crie um grupo primeiro")
            return
        
        # Janela para selecionar grupo
        select_window = tk.Toplevel(self.root)
        select_window.title("📝 Adicionar ao Grupo")
        select_window.geometry("400x300")
        select_window.configure(bg=ModernColors.SURFACE)
        select_window.transient(self.root)
        select_window.grab_set()
        
        # Header
        header_frame = ttk.Frame(select_window)
        header_frame.pack(fill=tk.X, padx=20, pady=(20, 10))
        
        title_label = ttk.Label(header_frame, text="📝 Adicionar ao Grupo",
                               font=('Segoe UI', 14, 'bold'),
                               foreground=ModernColors.TEXT_PRIMARY)
        title_label.pack()
        
        selected_count = len(selected_indices)
        subtitle_label = ttk.Label(header_frame, 
                                  text=f"Selecione o grupo para {selected_count} atividade(s)",
                                  font=('Segoe UI', 10, 'normal'),
                                  foreground=ModernColors.TEXT_SECONDARY,
                                  background=ModernColors.SURFACE)
        subtitle_label.pack(pady=(5, 0))
        
        # Separator
        separator = ttk.Separator(select_window, orient='horizontal')
        separator.pack(fill=tk.X, padx=20, pady=10)
        
        # Lista de grupos
        content_frame = ttk.Frame(select_window)
        content_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
        
        groups_label = ttk.Label(content_frame, text="📁 Grupos Disponíveis:",
                                font=('Segoe UI', 11, 'bold'),
                                foreground=ModernColors.TEXT_PRIMARY,
                                background=ModernColors.SURFACE)
        groups_label.pack(anchor='w', pady=(0, 10))
        
        group_listbox = tk.Listbox(content_frame,
                                  bg=ModernColors.SURFACE,
                                  fg=ModernColors.TEXT_PRIMARY,
                                  selectbackground=ModernColors.PRIMARY_LIGHT,
                                  selectforeground=ModernColors.PRIMARY,
                                  borderwidth=1,
                                  relief='solid',
                                  font=('Segoe UI', 10, 'normal'),
                                  activestyle='none')
        group_listbox.pack(fill=tk.BOTH, expand=True, pady=(0, 20))
        
        for group_name, group_data in self.activity_groups.items():
            activity_count = len(group_data['activities'])
            group_listbox.insert(tk.END, f"📁 {group_name} ({activity_count} atividades)")
        
        def confirm_addition():
            group_selection = group_listbox.curselection()
            if not group_selection:
                messagebox.showerror("❌ Erro", "Selecione um grupo")
                return
            
            # Extrair nome do grupo (remover ícone e contador)
            selected_text = group_listbox.get(group_selection[0])
            group_name = selected_text.replace("📁 ", "").split(" (")[0]
            
            # Obter atividades selecionadas (remover ícones e contadores)
            selected_activities = []
            for i in selected_indices:
                activity_text = self.available_listbox.get(i)
                activity_name = activity_text.replace("📊 ", "").split(" (")[0]
                selected_activities.append(activity_name)
            
            # Adicionar atividades ao grupo
            for activity in selected_activities:
                if activity not in self.activity_groups[group_name]['activities']:
                    self.activity_groups[group_name]['activities'].append(activity)
            
            # Atualizar interface
            self.update_group_tree()
            self.update_available_activities()
            select_window.destroy()
            
            messagebox.showinfo("✅ Sucesso", 
                f"{len(selected_activities)} atividade(s) adicionada(s) ao grupo '{group_name}' 🎉")
        
        # Botões
        button_frame = ttk.Frame(select_window)
        button_frame.pack(fill=tk.X, padx=20, pady=(0, 20))
        
        cancel_btn = ttk.Button(button_frame, text="❌ Cancelar",
                               command=select_window.destroy, style='Secondary.TButton')
        cancel_btn.pack(side=tk.RIGHT, padx=(10, 0))
        
        add_btn = ttk.Button(button_frame, text="✅ Adicionar",
                            command=confirm_addition, style='Primary.TButton')
        add_btn.pack(side=tk.RIGHT)
        
        # Centralizar janela
        select_window.update_idletasks()
        x = (self.root.winfo_screenwidth() // 2) - (select_window.winfo_width() // 2)
        y = (self.root.winfo_screenheight() // 2) - (select_window.winfo_height() // 2)
        select_window.geometry(f"+{x}+{y}")
    def debug_data(self):
        """Função de debug para analisar os dados dos arquivos com interface melhorada"""
        if not self.uploaded_files:
            messagebox.showwarning("⚠️ Aviso", "Nenhum arquivo selecionado")
            return
            
        debug_info = "=== 🔍 DEBUG DOS DADOS ===\n\n"
        
        for i, file_path in enumerate(self.uploaded_files):
            try:
                debug_info += f"📁 ARQUIVO {i+1}: {os.path.basename(file_path)}\n"
                debug_info += f"📍 Caminho: {file_path}\n"
                
                # Ler arquivo
                if file_path.endswith('.xlsx'):
                    df = pd.read_excel(file_path)
                else:
                    df = pd.read_csv(file_path)
                
                debug_info += f"📊 Dimensões: {df.shape[0]} linhas x {df.shape[1]} colunas\n"
                debug_info += f"📋 Colunas: {list(df.columns)}\n"
                
                # Se colunas foram selecionadas, analisar
                activity_col_display = self.activity_combo.get()
                time_col_display = self.time_combo.get()
                rework_col_display = self.rework_combo.get()
                
                if activity_col_display and time_col_display:
                    activity_col = activity_col_display.replace("📊 ", "")
                    time_col = time_col_display.replace("📊 ", "")
                    rework_col = rework_col_display.replace("📊 ", "") if rework_col_display else None
                    
                    if activity_col in df.columns:
                        debug_info += f"✅ Coluna atividade '{activity_col}': {df[activity_col].count()} valores não-nulos\n"
                        examples = list(df[activity_col].dropna().head(3))
                        debug_info += f"   📝 Exemplos: {examples}\n"
                    else:
                        debug_info += f"❌ Coluna atividade '{activity_col}': NÃO ENCONTRADA\n"
                    
                    if time_col in df.columns:
                        debug_info += f"✅ Coluna tempo '{time_col}': {df[time_col].count()} valores não-nulos\n"
                        examples = list(df[time_col].dropna().head(3))
                        debug_info += f"   📝 Exemplos: {examples}\n"
                        debug_info += f"   🔍 Tipo de dados: {df[time_col].dtype}\n"
                        
                        # Detectar formato de tempo
                        sample_values = df[time_col].dropna().head(5)
                        time_formats = []
                        for val in sample_values:
                            val_str = str(val).strip()
                            if ':' in val_str:
                                parts = val_str.split(':')
                                if len(parts) == 3:
                                    time_formats.append("⏰ HH:MM:SS")
                                elif len(parts) == 2:
                                    time_formats.append("⏱️ MM:SS")
                                else:
                                    time_formats.append("❓ Formato desconhecido")
                            else:
                                try:
                                    float(val_str.replace(',', '.'))
                                    time_formats.append("🔢 Numérico")
                                except:
                                    time_formats.append("📝 Texto")
                        
                        debug_info += f"   🎯 Formatos detectados: {time_formats}\n"
                    else:
                        debug_info += f"❌ Coluna tempo '{time_col}': NÃO ENCONTRADA\n"
                    
                    # Análise da coluna de retrabalho (se selecionada)
                    if rework_col:
                        if rework_col in df.columns:
                            rework_count = df[rework_col].count()
                            debug_info += f"🔄 Coluna retrabalho '{rework_col}': {rework_count} valores não-nulos\n"
                            
                            # Analisar valores únicos na coluna de retrabalho
                            unique_values = df[rework_col].dropna().unique()
                            debug_info += f"   📊 Valores únicos: {list(unique_values)}\n"
                            
                            # Contar quantas linhas seriam filtradas
                            rework_series = df[rework_col].astype(str).str.strip()
                            rework_mask = rework_series.isin(['1', '1.0', 'True', 'true', 'TRUE'])
                            rework_lines = rework_mask.sum()
                            debug_info += f"   🗑️ Linhas que seriam excluídas (retrabalho=1): {rework_lines}\n"
                            debug_info += f"   ✅ Linhas que seriam mantidas: {len(df) - rework_lines}\n"
                        else:
                            debug_info += f"❌ Coluna retrabalho '{rework_col}': NÃO ENCONTRADA\n"
                    else:
                        debug_info += "ℹ️ Coluna retrabalho: NÃO SELECIONADA (opcional)\n"
                
                debug_info += f"📄 Primeiras 3 linhas:\n{df.head(3).to_string()}\n"
                debug_info += "\n" + "="*60 + "\n\n"
                
            except Exception as e:
                debug_info += f"❌ ERRO ao ler arquivo: {str(e)}\n\n"
        
        # Criar janela de debug modernizada
        debug_window = tk.Toplevel(self.root)
        debug_window.title("🔍 Debug dos Dados")
        debug_window.geometry("900x700")
        debug_window.configure(bg=ModernColors.SURFACE)
        
        # Header
        header_frame = ttk.Frame(debug_window)
        header_frame.pack(fill=tk.X, padx=20, pady=(20, 10))
        
        title_label = ttk.Label(header_frame, text="🔍 Debug dos Dados",
                               font=('Segoe UI', 16, 'bold'),
                               foreground=ModernColors.TEXT_PRIMARY,
                               background=ModernColors.SURFACE)
        title_label.pack()
        
        subtitle_label = ttk.Label(header_frame, text="Informações detalhadas sobre os arquivos carregados",
                                  font=('Segoe UI', 11, 'normal'),
                                  foreground=ModernColors.TEXT_SECONDARY,
                                  background=ModernColors.SURFACE)
        subtitle_label.pack(pady=(5, 0))
        
        # Separator
        separator = ttk.Separator(debug_window, orient='horizontal')
        separator.pack(fill=tk.X, padx=20, pady=10)
        
        # Content frame
        content_frame = ttk.Frame(debug_window)
        content_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
        
        # Text widget with scrollbar
        text_frame = ttk.Frame(content_frame)
        text_frame.pack(fill=tk.BOTH, expand=True)
        
        text_widget = tk.Text(text_frame, wrap=tk.WORD,
                             bg=ModernColors.SURFACE_VARIANT,
                             fg=ModernColors.TEXT_PRIMARY,
                             borderwidth=1,
                             relief='solid',
                             font=('Consolas', 10, 'normal'),
                             padx=10, pady=10)
        
        scrollbar_v = ttk.Scrollbar(text_frame, orient="vertical", command=text_widget.yview)
        scrollbar_h = ttk.Scrollbar(text_frame, orient="horizontal", command=text_widget.xview)
        
        text_widget.configure(yscrollcommand=scrollbar_v.set, xscrollcommand=scrollbar_h.set)
        
        text_widget.grid(row=0, column=0, sticky="nsew")
        scrollbar_v.grid(row=0, column=1, sticky="ns")
        scrollbar_h.grid(row=1, column=0, sticky="ew")
        
        text_frame.grid_rowconfigure(0, weight=1)
        text_frame.grid_columnconfigure(0, weight=1)
        
        text_widget.insert("1.0", debug_info)
        text_widget.config(state=tk.DISABLED)  # Make read-only
        
        # Button frame
        button_frame = ttk.Frame(debug_window)
        button_frame.pack(fill=tk.X, padx=20, pady=(10, 20))
        
        close_btn = ttk.Button(button_frame, text="✅ Fechar",
                              command=debug_window.destroy, style='Primary.TButton')
        close_btn.pack(side=tk.RIGHT)

def main():
    """Função principal com configurações melhoradas"""
    root = tk.Tk()
    
    # Configurar ícone da aplicação (se disponível)
    try:
        # Opcional: definir ícone personalizado
        pass
    except:
        pass
    
    # Criar aplicação
    app = TimeStudyAnalyzer(root)
    
    # Configurar comportamento de fechamento
    def on_closing():
        if messagebox.askokcancel("🚪 Sair", "Deseja realmente sair da aplicação?"):
            root.destroy()
    
    root.protocol("WM_DELETE_WINDOW", on_closing)
    
    # Iniciar loop principal
    root.mainloop()

if __name__ == "__main__":
    main()
