"""
DomBot Admissional
==================
Automação RPA para emissão de Contratos Admissionais no sistema Domínio Folha.

Autor: Hugo L. Almeida
Versão: 1.0
"""

import customtkinter as ctk
import pandas as pd
from pywinauto.application import Application
from pywinauto.keyboard import send_keys
from pywinauto import findwindows, timings
import win32gui
import win32con
import ctypes
import time
import logging
import json
import unicodedata
from datetime import datetime, timedelta
import os
import requests
from dotenv import load_dotenv
from pathlib import Path
from pdf_corrigir_spans import corrigir_pdf

load_dotenv()
import traceback
import threading
import sys
from importlib import util as _importlib_util
from typing import Optional, Tuple
import tkinter.messagebox as messagebox
from PIL import Image, ImageDraw

# Importa M.E.G_ONE (projeto separado) por caminho de arquivo, pois o nome
# "M.E.G_ONE.py" contém pontos e não pode ser importado como módulo normal.
_MEG_ONE_PATH = r"C:\Users\Canella e Santos\Desktop\Hugo\RPA\MEG_ONE\M.E.G_ONE.py"
_meg_one_spec = _importlib_util.spec_from_file_location("meg_one", _MEG_ONE_PATH)
meg_one = _importlib_util.module_from_spec(_meg_one_spec)
_meg_one_spec.loader.exec_module(meg_one)


class GUILogHandler(logging.Handler):
    """Handler de logging que encaminha mensagens para a interface gráfica."""

    def __init__(self, gui):
        super().__init__()
        self.gui = gui

    def emit(self, record):
        msg = self.format(record)
        self.gui.window.after(0, lambda: self.gui.adicionar_log(msg, record.levelno))


class AutomacaoGUI:
    """Interface gráfica principal da automação."""

    # Cores do tema
    CORES = {
        'sucesso': '#2ECC71',
        'erro': '#E74C3C',
        'aviso': '#F39C12',
        'info': '#3498DB',
        'texto': '#ECF0F1',
        'fundo_card': '#2C3E50',
        'fundo_escuro': '#1A252F',
        'destaque': '#1ABC9C',
        'processando': '#9B59B6',
    }

    def __init__(self):
        # Configuração do tema
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("green")

        self.window = ctk.CTk()
        self.window.title("DomBot - Admissional v1.0")
        self.window.geometry("800x550")
        self.window.minsize(750, 500)
        self.window.protocol("WM_DELETE_WINDOW", self.ao_fechar)

        # Flags para controle de execução
        self.executando = False
        self.pausa_solicitada = False
        self.thread_automacao = None

        # Flags do dashboard de revisão do Excel principal (Meg_One)
        self.aguardando_revisao_dashboard = False
        self.revisao_cancelada = False
        self.caminho_excel_revisado = None
        self.usar_relatorios_existentes = ctk.BooleanVar(value=False)

        # Modo manual: emite contratos a partir de uma planilha própria, pulando
        # toda a geração automática (Empregados/Contrato -> Meg_One -> revisão).
        # Usado para contratos específicos que fogem do padrão automático.
        self.modo_manual = False

        # Estatísticas
        self.stats = {
            'processados': 0,
            'sucesso': 0,
            'erros': 0,
            'puladas': 0,
            'tempo_inicio': None
        }

        # DataFrame carregado
        self.df_carregado = None

        # Configurar ícone
        self.set_window_icon()

        # Criar diretório de logs se não existir
        self.logs_dir = os.path.join(os.path.dirname(__file__), "logs")
        if not os.path.exists(self.logs_dir):
            os.makedirs(self.logs_dir)

        # Configurar logging para arquivos
        self.setup_file_logging()

        # Variáveis da interface
        self.arquivo_excel = ctk.StringVar()
        self.periodo_referencia = ctk.StringVar(value=(datetime.now() + timedelta(days=1)).strftime("%d/%m/%Y"))
        self.diretorio_salvamento = ctk.StringVar()
        self.status_var = ctk.StringVar(value="Aguardando início...")

        # Variáveis exclusivas do modo manual
        self.arquivo_excel_manual = ctk.StringVar()
        self.diretorio_manual = ctk.StringVar()
        self.linha_inicial_manual = ctk.StringVar(value="2")

        # Variáveis de controle
        self.total_linhas = 0
        self.linhas_processadas = 0
        self.linhas_com_erro = 0
        self.linhas_puladas = 0

        # Logger
        self.logger = logging.getLogger('AutomacaoDominio')
        self.logger.setLevel(logging.INFO)
        self.logger.handlers = []

        # Adicionar GUIHandler
        self.gui_handler = GUILogHandler(self)
        formatter = logging.Formatter('%(message)s')
        self.gui_handler.setFormatter(formatter)
        self.logger.addHandler(self.gui_handler)

        self.criar_interface()

    def setup_file_logging(self):
        """Configura o logging para arquivos"""
        data_atual = datetime.now().strftime("%Y-%m-%d")

        # Logger de sucesso
        self.success_logger = logging.getLogger('SuccessLog')
        self.success_logger.setLevel(logging.INFO)
        if not self.success_logger.handlers:
            success_handler = logging.FileHandler(
                os.path.join(self.logs_dir, f'success_{data_atual}.log'),
                encoding='utf-8'
            )
            success_handler.setFormatter(
                logging.Formatter('%(asctime)s - %(message)s', '%Y-%m-%d %H:%M:%S')
            )
            self.success_logger.addHandler(success_handler)

        # Logger de erro
        self.error_logger = logging.getLogger('ErrorLog')
        self.error_logger.setLevel(logging.ERROR)
        if not self.error_logger.handlers:
            error_handler = logging.FileHandler(
                os.path.join(self.logs_dir, f'error_{data_atual}.log'),
                encoding='utf-8'
            )
            error_handler.setFormatter(
                logging.Formatter('%(asctime)s - %(message)s', '%Y-%m-%d %H:%M:%S')
            )
            self.error_logger.addHandler(error_handler)

    def set_window_icon(self):
        """Configura o ícone da janela"""
        try:
            icon_path = os.path.join(os.path.dirname(__file__), "assets", "favicon.ico")
            if os.name == 'nt' and os.path.exists(icon_path):
                self.window.iconbitmap(icon_path)
        except Exception as e:
            print(f"Erro ao carregar ícone: {e}")

    def criar_interface(self):
        # Frame principal com grid
        self.window.grid_columnconfigure(0, weight=1)
        self.window.grid_rowconfigure(0, weight=1)

        main_frame = ctk.CTkFrame(self.window, fg_color="transparent")
        main_frame.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        main_frame.grid_columnconfigure(0, weight=1)
        main_frame.grid_rowconfigure(3, weight=1)

        # === HEADER ===
        self.criar_header(main_frame)

        # === PAINEL DE CONFIGURAÇÃO ===
        self.criar_painel_config(main_frame)

        # === PAINEL DE ESTATÍSTICAS ===
        self.criar_painel_estatisticas(main_frame)

        # === ÁREA DE CONTEÚDO (Abas) ===
        self.criar_area_conteudo(main_frame)

    def criar_header(self, parent):
        """Cria o cabeçalho com título e status"""
        header_frame = ctk.CTkFrame(parent, fg_color=self.CORES['fundo_card'], corner_radius=8)
        header_frame.grid(row=0, column=0, sticky="ew", pady=(0, 6))
        header_frame.grid_columnconfigure(1, weight=1)

        # Ícone/Logo com fundo branco circular
        logo_path = os.path.join(os.path.dirname(__file__), "assets", "DomBot_New.png")
        if os.path.exists(logo_path):
            size = 66
            circle_size = 44
            bg = Image.new("RGBA", (size, size), (0, 0, 0, 0))
            circle_mask = Image.new("L", (circle_size, circle_size), 0)
            ImageDraw.Draw(circle_mask).ellipse((0, 0, circle_size - 1, circle_size - 1), fill=255)
            circle = Image.new("RGBA", (circle_size, circle_size), (255, 255, 255, 255))
            circle_offset = (size - circle_size) // 2
            bg.paste(circle, (circle_offset, circle_offset), circle_mask)
            original = Image.open(logo_path).convert("RGBA")
            original = original.resize((size, size), Image.LANCZOS)
            bg.paste(original, (0, 0), original)
            logo_image = ctk.CTkImage(light_image=bg, dark_image=bg, size=(size, size))
            ctk.CTkLabel(header_frame, image=logo_image, text="").grid(row=0, column=0, padx=10, pady=8)
        else:
            logo_frame = ctk.CTkFrame(header_frame, fg_color=self.CORES['destaque'],
                                       width=44, height=44, corner_radius=22)
            logo_frame.grid(row=0, column=0, padx=10, pady=8)
            logo_frame.grid_propagate(False)
            ctk.CTkLabel(logo_frame, text="🤖", font=("Segoe UI Emoji", 18)).place(relx=0.5, rely=0.5, anchor="center")

        # Título
        ctk.CTkLabel(
            header_frame,
            text="DomBot - Admissional",
            font=ctk.CTkFont(size=16, weight="bold"),
            text_color=self.CORES['texto']
        ).grid(row=0, column=1, sticky="w", padx=5)

        # Status indicator
        self.status_frame = ctk.CTkFrame(header_frame, fg_color="transparent")
        self.status_frame.grid(row=0, column=2, padx=10)

        self.status_indicator = ctk.CTkFrame(
            self.status_frame,
            fg_color="#7F8C8D",
            width=10, height=10,
            corner_radius=5
        )
        self.status_indicator.pack(side="left", padx=(0, 6))

        self.status_label = ctk.CTkLabel(
            self.status_frame,
            textvariable=self.status_var,
            font=ctk.CTkFont(size=11),
            text_color="#95A5A6"
        )
        self.status_label.pack(side="left")

    def criar_painel_config(self, parent):
        """Cria o painel de configuração"""
        config_frame = ctk.CTkFrame(parent, fg_color=self.CORES['fundo_card'], corner_radius=8)
        config_frame.grid(row=1, column=0, sticky="ew", pady=(0, 6))
        config_frame.grid_columnconfigure(1, weight=1)

        # Linha única com tudo
        inner_frame = ctk.CTkFrame(config_frame, fg_color="transparent")
        inner_frame.grid(row=0, column=0, sticky="ew", padx=10, pady=8)
        inner_frame.grid_columnconfigure(1, weight=1)

        # Botões de controle
        self.btn_iniciar = ctk.CTkButton(
            inner_frame, text="▶ Iniciar", command=self.iniciar_automacao_thread,
            width=90, height=32, font=ctk.CTkFont(size=11, weight="bold"),
            fg_color=self.CORES['sucesso'], hover_color="#27AE60"
        )
        self.btn_iniciar.grid(row=0, column=0, padx=3)

        self.btn_pausar = ctk.CTkButton(
            inner_frame, text="⏸ Pausar", command=self.pausar_automacao,
            width=90, height=32, font=ctk.CTkFont(size=11, weight="bold"),
            fg_color=self.CORES['aviso'], hover_color="#E67E22", state="disabled"
        )
        self.btn_pausar.grid(row=0, column=1, padx=3)

        self.btn_parar = ctk.CTkButton(
            inner_frame, text="⏹ Parar", command=self.parar_automacao,
            width=90, height=32, font=ctk.CTkFont(size=11, weight="bold"),
            fg_color=self.CORES['erro'], hover_color="#C0392B", state="disabled"
        )
        self.btn_parar.grid(row=0, column=2, padx=(3, 0))

        # Segunda linha — Período de referência e Diretório de salvamento
        inner_frame2 = ctk.CTkFrame(config_frame, fg_color="transparent")
        inner_frame2.grid(row=1, column=0, sticky="ew", padx=10, pady=(0, 8))
        inner_frame2.grid_columnconfigure(2, weight=1)

        # Período de referência
        ctk.CTkLabel(
            inner_frame2, text="📅", font=ctk.CTkFont(size=14)
        ).grid(row=0, column=0, padx=(0, 5))

        ctk.CTkLabel(
            inner_frame2, text="Período:", font=ctk.CTkFont(size=11), text_color="#BDC3C7"
        ).grid(row=0, column=1, padx=(0, 3))

        self.entry_periodo = ctk.CTkEntry(
            inner_frame2, textvariable=self.periodo_referencia,
            width=100, height=32, font=ctk.CTkFont(size=11), justify="center",
            placeholder_text="dd/mm/aaaa"
        )
        self.entry_periodo.grid(row=0, column=2, sticky="w", padx=(0, 15))

        # Diretório de salvamento
        ctk.CTkLabel(
            inner_frame2, text="📂", font=ctk.CTkFont(size=14)
        ).grid(row=0, column=3, padx=(0, 5))

        self.entry_diretorio = ctk.CTkEntry(
            inner_frame2,
            textvariable=self.diretorio_salvamento,
            placeholder_text="Diretório onde os PDFs serão salvos...",
            height=32,
            font=ctk.CTkFont(size=11)
        )
        self.entry_diretorio.grid(row=0, column=4, sticky="ew", padx=(0, 8))
        inner_frame2.grid_columnconfigure(4, weight=1)

        ctk.CTkButton(
            inner_frame2, text="Selecionar", command=self.selecionar_diretorio,
            width=80, height=32, font=ctk.CTkFont(size=11),
            fg_color=self.CORES['info'], hover_color="#2980B9"
        ).grid(row=0, column=5)

        # Checkbox: pular geração dos relatórios Empregados/Contrato
        self.check_usar_existentes = ctk.CTkCheckBox(
            inner_frame2,
            text="Usar relatórios já existentes (pular Empregados/Contrato)",
            variable=self.usar_relatorios_existentes,
            font=ctk.CTkFont(size=11)
        )
        self.check_usar_existentes.grid(row=1, column=0, columnspan=6, sticky="w", pady=(8, 0))

    def criar_painel_estatisticas(self, parent):
        """Cria o painel de estatísticas"""
        stats_frame = ctk.CTkFrame(parent, fg_color=self.CORES['fundo_card'], corner_radius=8)
        stats_frame.grid(row=2, column=0, sticky="ew", pady=(0, 6))

        # Grid para os cards de estatísticas
        for i in range(5):
            stats_frame.grid_columnconfigure(i, weight=1)

        # Cards de estatísticas
        self.criar_stat_card(stats_frame, 0, "📊", "Total", "total_label", "0")
        self.criar_stat_card(stats_frame, 1, "✅", "Sucesso", "sucesso_label", "0", self.CORES['sucesso'])
        self.criar_stat_card(stats_frame, 2, "❌", "Erros", "erros_label", "0", self.CORES['erro'])
        self.criar_stat_card(stats_frame, 3, "🏢", "Empresa", "empresa_label", "-", self.CORES['info'])
        self.criar_stat_card(stats_frame, 4, "⏱", "Tempo", "tempo_label", "00:00:00", self.CORES['aviso'])

        # Barra de progresso
        progress_frame = ctk.CTkFrame(stats_frame, fg_color="transparent")
        progress_frame.grid(row=1, column=0, columnspan=5, sticky="ew", padx=10, pady=(2, 8))
        progress_frame.grid_columnconfigure(0, weight=1)

        self.progress_bar = ctk.CTkProgressBar(
            progress_frame, height=6, corner_radius=3, progress_color=self.CORES['destaque']
        )
        self.progress_bar.grid(row=0, column=0, sticky="ew")
        self.progress_bar.set(0)

        self.progress_label = ctk.CTkLabel(
            progress_frame, text="0%", font=ctk.CTkFont(size=10), text_color="#95A5A6"
        )
        self.progress_label.grid(row=0, column=1, padx=(8, 0))

    def criar_stat_card(self, parent, col, icon, titulo, attr_name, valor_inicial, cor=None):
        """Cria um card de estatística"""
        card = ctk.CTkFrame(parent, fg_color="transparent")
        card.grid(row=0, column=col, padx=5, pady=8)

        ctk.CTkLabel(
            card, text=f"{icon} {titulo}", font=ctk.CTkFont(size=10), text_color="#7F8C8D"
        ).pack()

        label = ctk.CTkLabel(
            card, text=valor_inicial, font=ctk.CTkFont(size=14, weight="bold"),
            text_color=cor if cor else self.CORES['texto']
        )
        label.pack()

        setattr(self, attr_name, label)

    def criar_area_conteudo(self, parent):
        """Cria a área de conteúdo com abas"""
        self.tabview = ctk.CTkTabview(
            parent, fg_color=self.CORES['fundo_card'],
            segmented_button_fg_color=self.CORES['fundo_escuro'],
            segmented_button_selected_color=self.CORES['destaque'],
            corner_radius=8, height=25
        )
        self.tabview.grid(row=3, column=0, sticky="nsew")

        tab_logs = self.tabview.add("📋 Logs")
        tab_preview = self.tabview.add("📊 Preview")
        tab_manual = self.tabview.add("✍ Manual")
        tab_nao_emitir = self.tabview.add("🚫 Não Emitir")

        self.criar_aba_logs(tab_logs)
        self.criar_aba_preview(tab_preview)
        self.criar_aba_manual(tab_manual)
        self.criar_aba_nao_emitir(tab_nao_emitir)

    def criar_aba_logs(self, parent):
        """Cria a aba de logs"""
        parent.grid_columnconfigure(0, weight=1)
        parent.grid_rowconfigure(0, weight=1)

        log_container = ctk.CTkFrame(parent, fg_color="transparent")
        log_container.grid(row=0, column=0, sticky="nsew", padx=3, pady=3)
        log_container.grid_columnconfigure(0, weight=1)
        log_container.grid_rowconfigure(0, weight=1)

        self.log_text = ctk.CTkTextbox(
            log_container, font=ctk.CTkFont(family="Consolas", size=11),
            fg_color=self.CORES['fundo_escuro'], corner_radius=6
        )
        self.log_text.grid(row=0, column=0, sticky="nsew")

        # Configurar tags de cores
        self.log_text._textbox.tag_config("sucesso", foreground=self.CORES['sucesso'])
        self.log_text._textbox.tag_config("erro", foreground=self.CORES['erro'])
        self.log_text._textbox.tag_config("aviso", foreground=self.CORES['aviso'])
        self.log_text._textbox.tag_config("info", foreground=self.CORES['info'])
        self.log_text._textbox.tag_config("processando", foreground=self.CORES['processando'])

        # Botões de controle do log
        btn_frame = ctk.CTkFrame(log_container, fg_color="transparent")
        btn_frame.grid(row=1, column=0, sticky="ew", pady=(5, 0))

        ctk.CTkButton(
            btn_frame, text="🗑 Limpar", command=self.limpar_logs,
            width=90, height=26, font=ctk.CTkFont(size=10),
            fg_color="#34495E", hover_color="#2C3E50"
        ).pack(side="left")

        ctk.CTkButton(
            btn_frame, text="💾 Exportar", command=self.exportar_logs,
            width=90, height=26, font=ctk.CTkFont(size=10),
            fg_color="#34495E", hover_color="#2C3E50"
        ).pack(side="left", padx=8)

    def criar_aba_preview(self, parent):
        """Cria a aba de preview do Excel"""
        parent.grid_columnconfigure(0, weight=1)
        parent.grid_rowconfigure(1, weight=1)

        info_frame = ctk.CTkFrame(parent, fg_color="transparent")
        info_frame.grid(row=0, column=0, sticky="ew", padx=3, pady=3)

        self.preview_info_label = ctk.CTkLabel(
            info_frame, text="Nenhum arquivo carregado",
            font=ctk.CTkFont(size=11), text_color="#95A5A6"
        )
        self.preview_info_label.pack(side="left")

        ctk.CTkButton(
            info_frame, text="🔄 Recarregar", command=self.carregar_preview,
            width=85, height=24, font=ctk.CTkFont(size=10),
            fg_color="#34495E", hover_color="#2C3E50"
        ).pack(side="right")

        ctk.CTkButton(
            info_frame, text="📁 Abrir arquivo", command=self.selecionar_arquivo,
            width=100, height=24, font=ctk.CTkFont(size=10),
            fg_color=self.CORES['info'], hover_color="#2980B9"
        ).pack(side="right", padx=(0, 8))

        self.preview_text = ctk.CTkTextbox(
            parent, font=ctk.CTkFont(family="Consolas", size=10),
            fg_color=self.CORES['fundo_escuro'], corner_radius=6
        )
        self.preview_text.grid(row=1, column=0, sticky="nsew", padx=3, pady=(0, 3))

    def criar_aba_manual(self, parent):
        """Cria a aba de emissão manual — emite contratos a partir de uma planilha
        própria, sem gerar os relatórios de Empregados/Contrato nem passar pelo
        Meg_One. Usada para contratos específicos que fogem do padrão automático."""
        parent.grid_columnconfigure(0, weight=1)
        parent.grid_rowconfigure(3, weight=1)

        ctk.CTkLabel(
            parent,
            text="Emite direto da sua planilha, pulando a geração automática dos relatórios "
                 "e o filtro de 'Não Emitir'.",
            font=ctk.CTkFont(size=10), text_color="#95A5A6",
            wraplength=700, justify="left"
        ).grid(row=0, column=0, sticky="w", padx=3, pady=(3, 6))

        # Planilha manual
        arq_frame = ctk.CTkFrame(parent, fg_color="transparent")
        arq_frame.grid(row=1, column=0, sticky="ew", padx=3)
        arq_frame.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(
            arq_frame, text="📁", font=ctk.CTkFont(size=14)
        ).grid(row=0, column=0, padx=(0, 5))

        ctk.CTkEntry(
            arq_frame, textvariable=self.arquivo_excel_manual,
            placeholder_text="Planilha com os contratos a emitir (Nº, Cod.Funcionário, Tipo de Contrato, Documento)...",
            height=30, font=ctk.CTkFont(size=11)
        ).grid(row=0, column=1, sticky="ew", padx=(0, 8))

        ctk.CTkButton(
            arq_frame, text="Procurar", command=self.selecionar_arquivo_manual,
            width=80, height=30, font=ctk.CTkFont(size=11),
            fg_color=self.CORES['info'], hover_color="#2980B9"
        ).grid(row=0, column=2, padx=(0, 12))

        ctk.CTkLabel(
            arq_frame, text="Linha:", font=ctk.CTkFont(size=11), text_color="#BDC3C7"
        ).grid(row=0, column=3, padx=(0, 3))

        ctk.CTkEntry(
            arq_frame, textvariable=self.linha_inicial_manual,
            width=50, height=30, font=ctk.CTkFont(size=11), justify="center"
        ).grid(row=0, column=4)

        # Diretório de salvamento próprio do modo manual
        dir_frame = ctk.CTkFrame(parent, fg_color="transparent")
        dir_frame.grid(row=2, column=0, sticky="ew", padx=3, pady=(8, 6))
        dir_frame.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(
            dir_frame, text="📂", font=ctk.CTkFont(size=14)
        ).grid(row=0, column=0, padx=(0, 5))

        ctk.CTkEntry(
            dir_frame, textvariable=self.diretorio_manual,
            placeholder_text="Diretório onde os PDFs manuais serão salvos...",
            height=30, font=ctk.CTkFont(size=11)
        ).grid(row=0, column=1, sticky="ew", padx=(0, 8))

        ctk.CTkButton(
            dir_frame, text="Selecionar", command=self.selecionar_diretorio_manual,
            width=90, height=30, font=ctk.CTkFont(size=11),
            fg_color=self.CORES['info'], hover_color="#2980B9"
        ).grid(row=0, column=2, padx=(0, 12))

        self.btn_emitir_manual = ctk.CTkButton(
            dir_frame, text="▶ Emitir manual", command=self.iniciar_manual_thread,
            width=130, height=30, font=ctk.CTkFont(size=11, weight="bold"),
            fg_color=self.CORES['sucesso'], hover_color="#27AE60"
        )
        self.btn_emitir_manual.grid(row=0, column=3)

        # Preview da planilha manual
        self.preview_manual_text = ctk.CTkTextbox(
            parent, font=ctk.CTkFont(family="Consolas", size=10),
            fg_color=self.CORES['fundo_escuro'], corner_radius=6
        )
        self.preview_manual_text.grid(row=3, column=0, sticky="nsew", padx=3, pady=(0, 3))

    def selecionar_arquivo_manual(self):
        """Seleciona a planilha do modo manual e carrega seu preview."""
        filename = ctk.filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx *.xls")],
            title="Selecione a planilha dos contratos manuais"
        )
        if filename:
            self.arquivo_excel_manual.set(filename)
            self.adicionar_log(f"[Manual] Planilha selecionada: {os.path.basename(filename)}", logging.INFO, "info")
            self.carregar_preview_manual()

    def selecionar_diretorio_manual(self):
        """Seleciona o diretório de salvamento dos PDFs manuais."""
        import subprocess
        try:
            script = (
                "Add-Type -AssemblyName System.Windows.Forms; "
                "$f = New-Object System.Windows.Forms.FolderBrowserDialog; "
                "$f.Description = 'Selecione o diretório para salvar os PDFs manuais'; "
                "$f.ShowNewFolderButton = $true; "
                "if ($f.ShowDialog() -eq 'OK') { $f.SelectedPath } else { '' }"
            )
            result = subprocess.run(
                ["powershell", "-Command", script],
                capture_output=True, text=True, timeout=120
            )
            diretorio = result.stdout.strip()
            if diretorio:
                self.diretorio_manual.set(diretorio)
                self.adicionar_log(f"[Manual] Diretório de salvamento: {diretorio}", logging.INFO, "info")
        except Exception as e:
            self.adicionar_log(f"[Manual] Erro ao selecionar diretório: {e}", logging.ERROR, "erro")

    def carregar_preview_manual(self):
        """Carrega o preview da planilha manual e valida as colunas obrigatórias."""
        caminho = self.arquivo_excel_manual.get()
        if not caminho:
            return

        try:
            df = pd.read_excel(caminho)
            self.preview_manual_text.delete("1.0", "end")

            header = " | ".join([f"{col:^15}" for col in df.columns[:6]])
            self.preview_manual_text.insert("end", f"{'─' * len(header)}\n")
            self.preview_manual_text.insert("end", f"{header}\n")
            self.preview_manual_text.insert("end", f"{'─' * len(header)}\n")

            for _, row in df.head(50).iterrows():
                row_text = " | ".join([f"{str(val)[:15]:^15}" for val in row.values[:6]])
                self.preview_manual_text.insert("end", f"{row_text}\n")

            if len(df) > 50:
                self.preview_manual_text.insert("end", f"\n... e mais {len(df) - 50} linhas\n")

            colunas_faltando = [
                col for col in ['Nº', 'Cod.Funcionário', 'Tipo de Contrato', 'Documento']
                if col not in df.columns
            ]
            if colunas_faltando:
                self.adicionar_log(
                    f"[Manual] Colunas obrigatórias não encontradas: {', '.join(colunas_faltando)}",
                    logging.WARNING, "aviso"
                )
            else:
                self.adicionar_log(
                    f"[Manual] Preview carregado: {len(df)} linhas. Todas as colunas obrigatórias encontradas",
                    logging.INFO, "sucesso"
                )
        except Exception as e:
            self.adicionar_log(f"[Manual] Erro ao carregar preview: {str(e)}", logging.ERROR, "erro")

    def criar_aba_nao_emitir(self, parent):
        """Cria a aba de gerenciamento da lista de empresas que nunca devem
        ser emitidas automaticamente (ex: Criat, contrato tratado manualmente)."""
        parent.grid_columnconfigure(0, weight=1)
        parent.grid_rowconfigure(1, weight=1)

        info_frame = ctk.CTkFrame(parent, fg_color="transparent")
        info_frame.grid(row=0, column=0, sticky="ew", padx=3, pady=3)

        ctk.CTkLabel(
            info_frame,
            text="Empresas cujo nome contenha um destes termos nunca serão emitidas (comparação parcial, sem diferenciar maiúsculas).",
            font=ctk.CTkFont(size=10), text_color="#95A5A6"
        ).pack(side="left")

        add_frame = ctk.CTkFrame(parent, fg_color="transparent")
        add_frame.grid(row=1, column=0, sticky="ew", padx=3, pady=(0, 5))
        add_frame.grid_columnconfigure(0, weight=1)

        self.entry_nao_emitir = ctk.CTkEntry(
            add_frame, placeholder_text="Nome ou trecho do nome da empresa (ex: CRIAT)",
            height=28, font=ctk.CTkFont(size=11)
        )
        self.entry_nao_emitir.grid(row=0, column=0, sticky="ew", padx=(0, 8))

        ctk.CTkButton(
            add_frame, text="Adicionar", command=self._adicionar_termo_nao_emitir,
            width=90, height=28, font=ctk.CTkFont(size=11),
            fg_color=self.CORES['sucesso'], hover_color="#27AE60"
        ).grid(row=0, column=1)

        self.lista_nao_emitir_frame = ctk.CTkScrollableFrame(parent, fg_color=self.CORES['fundo_escuro'])
        self.lista_nao_emitir_frame.grid(row=2, column=0, sticky="nsew", padx=3, pady=(0, 3))
        parent.grid_rowconfigure(2, weight=1)

        self._renderizar_lista_nao_emitir()

    def _renderizar_lista_nao_emitir(self):
        """Redesenha a lista de termos na aba 'Não Emitir' a partir do arquivo."""
        for widget in self.lista_nao_emitir_frame.winfo_children():
            widget.destroy()

        lista = self.carregar_lista_nao_emitir()
        for i, termo in enumerate(lista):
            linha = ctk.CTkFrame(self.lista_nao_emitir_frame, fg_color="transparent")
            linha.pack(fill="x", pady=2, padx=2)

            ctk.CTkLabel(linha, text=termo, font=ctk.CTkFont(size=11)).pack(side="left", padx=(4, 0))

            ctk.CTkButton(
                linha, text="Remover", width=70, height=22, font=ctk.CTkFont(size=9),
                fg_color=self.CORES['erro'], hover_color="#C0392B",
                command=lambda t=termo: self._remover_termo_nao_emitir(t)
            ).pack(side="right")

    def _adicionar_termo_nao_emitir(self):
        termo = self.entry_nao_emitir.get().strip()
        if not termo:
            return
        lista = self.carregar_lista_nao_emitir()
        if termo.upper() not in [t.upper() for t in lista]:
            lista.append(termo)
            self.salvar_lista_nao_emitir(lista)
            self.adicionar_log(f"Adicionado à lista de não emitir: '{termo}'", logging.INFO, "info")
        self.entry_nao_emitir.delete(0, "end")
        self._renderizar_lista_nao_emitir()

    def _remover_termo_nao_emitir(self, termo: str):
        lista = self.carregar_lista_nao_emitir()
        lista = [t for t in lista if t != termo]
        self.salvar_lista_nao_emitir(lista)
        self.adicionar_log(f"Removido da lista de não emitir: '{termo}'", logging.INFO, "info")
        self._renderizar_lista_nao_emitir()

    def selecionar_diretorio(self):
        """Abre diálogo para selecionar diretório de salvamento dos PDFs"""
        import subprocess
        try:
            # Usa PowerShell para abrir o diálogo de pasta nativo do Windows
            # Isso evita conflitos com o event loop do customtkinter
            script = (
                "Add-Type -AssemblyName System.Windows.Forms; "
                "$f = New-Object System.Windows.Forms.FolderBrowserDialog; "
                "$f.Description = 'Selecione o diretório para salvar os PDFs'; "
                "$f.ShowNewFolderButton = $true; "
                "if ($f.ShowDialog() -eq 'OK') { $f.SelectedPath } else { '' }"
            )
            result = subprocess.run(
                ["powershell", "-Command", script],
                capture_output=True, text=True, timeout=120
            )
            diretorio = result.stdout.strip()
            if diretorio:
                self.diretorio_salvamento.set(diretorio)
                self.adicionar_log(f"Diretório de salvamento: {diretorio}", logging.INFO, "info")
        except Exception as e:
            self.adicionar_log(f"Erro ao selecionar diretório: {e}", logging.ERROR, "erro")

    def selecionar_arquivo(self):
        """Abre diálogo para selecionar arquivo Excel"""
        filename = ctk.filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx *.xls")],
            title="Selecione o arquivo Excel"
        )
        if filename:
            self.arquivo_excel.set(filename)
            self.adicionar_log(f"Arquivo selecionado: {os.path.basename(filename)}", logging.INFO, "info")
            self.carregar_preview()

    def carregar_preview(self):
        """Carrega preview do arquivo Excel"""
        if not self.arquivo_excel.get():
            return

        try:
            self.df_carregado = pd.read_excel(self.arquivo_excel.get())
            total_linhas = len(self.df_carregado)

            # Atualizar info
            self.preview_info_label.configure(
                text=f"📄 {os.path.basename(self.arquivo_excel.get())} | {total_linhas} linhas | Colunas: {', '.join(self.df_carregado.columns[:6])}..."
            )

            # Atualizar estatística de total
            self.total_label.configure(text=str(total_linhas))

            # Mostrar preview
            self.preview_text.delete("1.0", "end")

            # Cabeçalho
            header = " | ".join([f"{col:^15}" for col in self.df_carregado.columns[:6]])
            self.preview_text.insert("end", f"{'─' * len(header)}\n")
            self.preview_text.insert("end", f"{header}\n")
            self.preview_text.insert("end", f"{'─' * len(header)}\n")

            # Dados (primeiras 50 linhas)
            for _, row in self.df_carregado.head(50).iterrows():
                row_text = " | ".join([f"{str(val)[:15]:^15}" for val in row.values[:6]])
                self.preview_text.insert("end", f"{row_text}\n")

            if total_linhas > 50:
                self.preview_text.insert("end", f"\n... e mais {total_linhas - 50} linhas\n")

            # Validar colunas necessárias
            colunas_necessarias = ['Nº', 'Cod.Funcionário', 'Tipo de Contrato', 'Documento']
            colunas_faltando = [col for col in colunas_necessarias if col not in self.df_carregado.columns]

            if colunas_faltando:
                self.adicionar_log(f"Colunas obrigatórias não encontradas: {', '.join(colunas_faltando)}", logging.WARNING, "aviso")
            else:
                self.adicionar_log(f"Preview carregado: {total_linhas} linhas. Todas as colunas obrigatórias encontradas", logging.INFO, "sucesso")

        except Exception as e:
            self.adicionar_log(f"Erro ao carregar preview: {str(e)}", logging.ERROR, "erro")

    def limpar_logs(self):
        """Limpa a área de logs"""
        self.log_text.delete("1.0", "end")
        self.adicionar_log("Log limpo", logging.INFO, "info")

    def exportar_logs(self):
        """Exporta logs para arquivo"""
        try:
            filename = ctk.filedialog.asksaveasfilename(
                defaultextension=".txt",
                filetypes=[("Text files", "*.txt"), ("All files", "*.*")],
                initialfilename=f"logs_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
            )
            if filename:
                with open(filename, 'w', encoding='utf-8') as f:
                    f.write(self.log_text.get("1.0", "end"))
                self.adicionar_log(f"Logs exportados para: {filename}", logging.INFO, "sucesso")
        except Exception as e:
            self.adicionar_log(f"Erro ao exportar logs: {str(e)}", logging.ERROR, "erro")

    def atualizar_progresso(self, atual, total):
        """Atualiza a barra de progresso"""
        porcentagem = atual / total if total > 0 else 0
        self.progress_bar.set(porcentagem)
        self.progress_label.configure(text=f"{porcentagem * 100:.1f}%")
        self.status_var.set(f"Processando: {atual}/{total}")
        self.window.update_idletasks()

    def atualizar_estatisticas(self):
        """Atualiza os cards de estatísticas"""
        self.sucesso_label.configure(text=str(self.linhas_processadas))
        self.erros_label.configure(text=str(self.linhas_com_erro))
        self.stats['processados'] = self.linhas_processadas + self.linhas_com_erro

    def atualizar_tempo(self):
        """Atualiza o tempo decorrido"""
        if self.stats['tempo_inicio']:
            elapsed = datetime.now() - self.stats['tempo_inicio']
            hours, remainder = divmod(int(elapsed.total_seconds()), 3600)
            minutes, seconds = divmod(remainder, 60)
            self.tempo_label.configure(text=f"{hours:02d}:{minutes:02d}:{seconds:02d}")
            if self.executando:
                self.window.after(1000, self.atualizar_tempo)

    def atualizar_status_indicator(self, status):
        """Atualiza o indicador de status visual"""
        cores = {
            'aguardando': '#7F8C8D',
            'executando': self.CORES['sucesso'],
            'pausado': self.CORES['aviso'],
            'erro': self.CORES['erro'],
            'concluido': self.CORES['info']
        }
        self.status_indicator.configure(fg_color=cores.get(status, '#7F8C8D'))

    def adicionar_log(self, mensagem, level=logging.INFO, tag=None):
        """Adiciona mensagem ao log visual com cores"""
        try:
            timestamp = datetime.now().strftime('%H:%M:%S')

            # Determinar tag baseado no nível se não especificado
            if tag is None:
                if level >= logging.ERROR:
                    tag = "erro"
                elif level >= logging.WARNING:
                    tag = "aviso"
                elif "sucesso" in mensagem.lower() or "processad" in mensagem.lower():
                    tag = "sucesso"
                else:
                    tag = "info"

            # Prefixo visual
            prefixos = {
                "sucesso": "✅",
                "erro": "❌",
                "aviso": "⚠️",
                "info": "ℹ️",
                "processando": "⏳"
            }
            prefixo = prefixos.get(tag, "•")

            # Inserir mensagem
            self.log_text.insert("end", f"[{timestamp}] {prefixo} ", tag)
            self.log_text.insert("end", f"{mensagem}\n", tag)
            self.log_text.see("end")
            self.window.update_idletasks()
        except Exception:
            pass

    def validar_entrada(self) -> Tuple[bool, str]:
        """Valida os dados de entrada. O Excel principal é gerado automaticamente
        pela automação (Empregados/Contrato -> Meg_One -> dashboard de revisão),
        então não é validado aqui — só os campos conhecidos antes de iniciar."""
        if not self.diretorio_salvamento.get():
            return False, "Selecione o diretório de salvamento"

        if self.usar_relatorios_existentes.get():
            diretorio = self.obter_pasta_dia()
            pasta_excels = os.path.join(diretorio, "Excels Base")
            data_ref = datetime.now().strftime("%Y-%m-%d")
            try:
                ref = datetime.strptime(self.periodo_referencia.get().strip(), "%d/%m/%Y")
                data_ref = ref.strftime("%Y-%m-%d")
            except Exception:
                pass
            caminho_empregados = os.path.join(pasta_excels, f"Empregados_Admissao_{data_ref}.xls")
            caminho_contrato = os.path.join(pasta_excels, f"ContratoPrazoDeterminado_{data_ref}.xls")
            if not os.path.exists(caminho_empregados):
                return False, f"Relatório de Empregados não encontrado: {caminho_empregados}"
            if not os.path.exists(caminho_contrato):
                return False, f"Relatório de Contrato não encontrado: {caminho_contrato}"

        return True, "Validação OK"

    def iniciar_automacao_thread(self):
        """Inicia a automação em uma thread separada"""
        if self.executando:
            self.adicionar_log("Automação já em execução", logging.WARNING, "aviso")
            return

        # Validar entrada
        valido, mensagem = self.validar_entrada()
        if not valido:
            self.adicionar_log(f"Erro de validação: {mensagem}", logging.ERROR, "erro")
            messagebox.showerror("Erro de Validação", mensagem)
            return

        # Resetar estatísticas
        self.linhas_processadas = 0
        self.linhas_com_erro = 0
        self.linhas_puladas = 0
        self.stats = {'processados': 0, 'sucesso': 0, 'erros': 0, 'puladas': 0, 'tempo_inicio': datetime.now()}
        self.sucesso_label.configure(text="0")
        self.erros_label.configure(text="0")

        self.thread_automacao = threading.Thread(target=self.iniciar_automacao)
        self.thread_automacao.daemon = True
        self.thread_automacao.start()

        # Atualizar interface
        self.btn_iniciar.configure(state="disabled")
        self.btn_emitir_manual.configure(state="disabled")
        self.btn_pausar.configure(state="normal")
        self.btn_parar.configure(state="normal")
        self.atualizar_status_indicator('executando')

        # Iniciar timer
        self.atualizar_tempo()

    def validar_entrada_manual(self) -> Tuple[bool, str]:
        """Valida os campos do modo manual. Diferente do automático, aqui a
        planilha é fornecida pelo usuário, então ela é validada de fato."""
        caminho = self.arquivo_excel_manual.get().strip()
        if not caminho:
            return False, "Selecione a planilha dos contratos manuais"

        if not os.path.exists(caminho):
            return False, f"Planilha não encontrada: {caminho}"

        if not self.diretorio_manual.get().strip():
            return False, "Selecione o diretório de salvamento dos PDFs manuais"

        try:
            linha = int(self.linha_inicial_manual.get())
        except ValueError:
            return False, "Linha inicial deve ser um número inteiro"
        if linha < 2:
            return False, "Linha inicial deve ser 2 ou maior (linha 1 é o cabeçalho)"

        try:
            df = pd.read_excel(caminho)
        except Exception as e:
            return False, f"Erro ao ler a planilha: {e}"

        colunas_faltando = [
            col for col in ['Nº', 'Cod.Funcionário', 'Tipo de Contrato', 'Documento']
            if col not in df.columns
        ]
        if colunas_faltando:
            return False, f"Colunas obrigatórias não encontradas: {', '.join(colunas_faltando)}"

        if linha - 2 >= len(df):
            return False, f"Linha inicial {linha} está além do fim da planilha ({len(df)} linhas de dados)"

        return True, "Validação OK"

    def iniciar_manual_thread(self):
        """Inicia a emissão manual em uma thread separada."""
        if self.executando:
            self.adicionar_log("Automação já em execução", logging.WARNING, "aviso")
            return

        valido, mensagem = self.validar_entrada_manual()
        if not valido:
            self.adicionar_log(f"[Manual] Erro de validação: {mensagem}", logging.ERROR, "erro")
            messagebox.showerror("Erro de Validação — Modo Manual", mensagem)
            return

        self.modo_manual = True

        # Resetar estatísticas
        self.linhas_processadas = 0
        self.linhas_com_erro = 0
        self.linhas_puladas = 0
        self.stats = {'processados': 0, 'sucesso': 0, 'erros': 0, 'puladas': 0, 'tempo_inicio': datetime.now()}
        self.sucesso_label.configure(text="0")
        self.erros_label.configure(text="0")

        self.thread_automacao = threading.Thread(target=self.iniciar_automacao_manual)
        self.thread_automacao.daemon = True
        self.thread_automacao.start()

        self.btn_iniciar.configure(state="disabled")
        self.btn_emitir_manual.configure(state="disabled")
        self.btn_pausar.configure(state="normal")
        self.btn_parar.configure(state="normal")
        self.atualizar_status_indicator('executando')
        self.atualizar_tempo()

    def pausar_automacao(self):
        """Pausa/retoma a automação"""
        if self.executando:
            self.pausa_solicitada = not self.pausa_solicitada
            if self.pausa_solicitada:
                self.btn_pausar.configure(text="▶  Retomar")
                self.status_var.set("Pausado")
                self.atualizar_status_indicator('pausado')
                self.adicionar_log("Automação pausada", logging.INFO, "aviso")
            else:
                self.btn_pausar.configure(text="⏸  Pausar")
                self.status_var.set("Em execução...")
                self.atualizar_status_indicator('executando')
                self.adicionar_log("Automação retomada", logging.INFO, "info")

    def parar_automacao(self):
        """Para a execução da automação"""
        if self.executando:
            self.executando = False
            self.pausa_solicitada = False
            self.adicionar_log("Solicitação de parada enviada. Aguardando conclusão...", logging.INFO, "aviso")
            self.status_var.set("Interrompendo...")
            self.atualizar_status_indicator('erro')

    def ao_fechar(self):
        """Tratamento do fechamento da janela"""
        if self.executando:
            if messagebox.askyesno("Confirmação",
                                   "Existe uma automação em execução. Deseja realmente sair?"):
                self.executando = False
                self.pausa_solicitada = False
                self.window.after(1000, self.window.destroy)
        else:
            self.window.destroy()

    def obter_pasta_dia(self) -> str:
        """Retorna o caminho da pasta do dia (dd.mm.aa) dentro do diretório raiz
        selecionado, criando-a (e a subpasta 'Excels Base') se não existir."""
        raiz = self.diretorio_salvamento.get()
        try:
            ref = datetime.strptime(self.periodo_referencia.get().strip(), "%d/%m/%Y")
        except Exception:
            ref = datetime.now()
        pasta_dia = os.path.join(raiz, ref.strftime("%d.%m.%y"))
        os.makedirs(os.path.join(pasta_dia, "Excels Base"), exist_ok=True)
        return pasta_dia

    def _caminho_lista_nao_emitir(self) -> str:
        return os.path.join(os.path.dirname(os.path.abspath(__file__)), "nao_emitir.json")

    def carregar_lista_nao_emitir(self) -> list:
        """Lê a lista de nomes de empresa que nunca devem ser emitidos
        (ex: 'CRIAT', tratada manualmente por outra equipe). Cria o arquivo
        com valor padrão se ainda não existir."""
        caminho = self._caminho_lista_nao_emitir()
        if not os.path.exists(caminho):
            padrao = ["CRIAT"]
            self.salvar_lista_nao_emitir(padrao)
            return padrao
        try:
            with open(caminho, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return []

    def salvar_lista_nao_emitir(self, lista: list) -> None:
        caminho = self._caminho_lista_nao_emitir()
        with open(caminho, "w", encoding="utf-8") as f:
            json.dump(lista, f, ensure_ascii=False, indent=2)

    def filtrar_nao_emitir(self, df: pd.DataFrame) -> pd.DataFrame:
        """Remove linhas cuja EMPRESAS contenha (case-insensitive) algum termo
        da lista de não emitir. Loga cada exclusão para o usuário saber que
        foi intencional, não erro."""
        lista = self.carregar_lista_nao_emitir()
        if not lista or 'EMPRESAS' not in df.columns:
            return df

        mascara_excluir = pd.Series(False, index=df.index)
        for termo in lista:
            termo = termo.strip()
            if not termo:
                continue
            mascara_termo = df['EMPRESAS'].astype(str).str.contains(termo, case=False, regex=False, na=False)
            for empresa in df.loc[mascara_termo, 'EMPRESAS'].unique():
                self.adicionar_log(f"🚫 Empresa '{empresa}' está na lista de não emitir — ignorada", logging.WARNING, "aviso")
            mascara_excluir |= mascara_termo

        return df[~mascara_excluir].reset_index(drop=True)

    def gerar_excel_principal_meg_one(self) -> Optional[str]:
        """Chama processar_dombot_admiss (M.E.G_ONE) usando os 2 relatórios já
        gerados (Empregados e Contrato por Prazo Determinado) para montar o Excel
        principal de entrada do loop de emissão. Retorna o caminho gerado, ou None em erro."""
        try:
            if not self.diretorio_salvamento.get():
                self.adicionar_log("Diretório de salvamento não definido — necessário para localizar os relatórios", logging.ERROR, "erro")
                return None

            diretorio = self.obter_pasta_dia()
            pasta_excels = os.path.join(diretorio, "Excels Base")

            ref = datetime.strptime(self.periodo_referencia.get().strip(), "%d/%m/%Y")
            data_ref = ref.strftime("%Y-%m-%d")
            periodo_meg_one = ref.strftime("%d.%m.%y")

            caminho_empregados = os.path.join(pasta_excels, f"Empregados_Admissao_{data_ref}.xls")
            caminho_contrato = os.path.join(pasta_excels, f"ContratoPrazoDeterminado_{data_ref}.xls")

            if not os.path.exists(caminho_empregados):
                self.adicionar_log(f"Relatório de Empregados não encontrado: {caminho_empregados}", logging.ERROR, "erro")
                return None
            if not os.path.exists(caminho_contrato):
                self.adicionar_log(f"Relatório de Contrato não encontrado: {caminho_contrato}", logging.ERROR, "erro")
                return None

            excel_saida = os.path.join(pasta_excels, f"ExcelPrincipal_{data_ref}.xlsx")

            total = meg_one.processar_dombot_admiss(
                caminho_empregados,
                caminho_contrato,
                excel_saida,
                log_callback=lambda msg: self.adicionar_log(f"[Meg_One] {msg}", logging.INFO, "info"),
                progress_callback=lambda p: None,
                pasta_destino=diretorio,
                periodo=periodo_meg_one,
            )
            self.adicionar_log(f"Meg_One gerou {total} registro(s) no Excel principal", logging.INFO, "info")
            return excel_saida

        except Exception as e:
            self.adicionar_log(f"Erro ao gerar Excel principal via Meg_One: {str(e)}", logging.ERROR, "erro")
            return None

    def abrir_dashboard_revisao(self, caminho_excel: str):
        """Abre a tela de revisão do Excel principal, permitindo corrigir o
        código de empresa (Nº) de cada linha antes de liberar a emissão de
        contratos. Roda na main thread (agendado via window.after)."""
        try:
            df = pd.read_excel(caminho_excel)
        except Exception as e:
            self.adicionar_log(f"Erro ao abrir Excel para revisão: {str(e)}", logging.ERROR, "erro")
            self.revisao_cancelada = True
            self.aguardando_revisao_dashboard = False
            return

        janela = ctk.CTkToplevel(self.window)
        janela.title("Revisão do Excel Principal")
        janela.geometry("1250x650")
        janela.grab_set()

        def on_close():
            if messagebox.askyesno("Cancelar", "Cancelar a automação? Nenhum contrato será emitido."):
                self.revisao_cancelada = True
                self.aguardando_revisao_dashboard = False
                janela.destroy()

        janela.protocol("WM_DELETE_WINDOW", on_close)

        def checar_parada():
            if not self.executando:
                janela.destroy()
                return
            janela.after(500, checar_parada)
        janela.after(500, checar_parada)

        aviso = ctk.CTkLabel(
            janela,
            text="Revise o código (Nº) de cada empresa antes de continuar. Linhas em vermelho ficaram sem código "
                 "(ambíguas — preencha ou remova). Use 'Remover' para excluir linhas que não devem ser emitidas "
                 "(ex: contrato tratado manualmente por outra equipe).",
            font=ctk.CTkFont(size=11), text_color=self.CORES['aviso'], wraplength=860
        )
        aviso.pack(fill="x", padx=10, pady=(10, 5))

        scroll = ctk.CTkScrollableFrame(janela)
        scroll.pack(fill="both", expand=True, padx=10, pady=5)

        colunas = ['Nº', 'EMPRESAS', 'Cod.Funcionário', 'Funcionário', 'Tipo de Contrato', 'Documento']
        for c, nome_col in enumerate(colunas):
            ctk.CTkLabel(scroll, text=nome_col, font=ctk.CTkFont(size=11, weight="bold")).grid(
                row=0, column=c, padx=4, pady=4, sticky="w"
            )
        ctk.CTkLabel(scroll, text="", font=ctk.CTkFont(size=11, weight="bold")).grid(
            row=0, column=len(colunas), padx=4, pady=4, sticky="w"
        )

        linhas_removidas = set()
        entries_numero = {}
        linha_widgets = {}

        def remover_linha(idx_df):
            linhas_removidas.add(idx_df)
            for widget in linha_widgets[idx_df]:
                widget.configure(state="disabled")
                try:
                    widget.configure(text_color="gray50")
                except Exception:
                    pass

        def formatar_numero(valor):
            """Formata o código de empresa sem .0 (colunas com célula vazia
            viram float64 ao salvar/reler Excel, ex: 676 -> 676.0)."""
            if pd.isna(valor) or str(valor).strip() == '':
                return ""
            try:
                return str(int(float(valor)))
            except (ValueError, TypeError):
                return str(valor).strip()

        for i, (idx_df, row) in enumerate(df.iterrows(), start=1):
            valor_numero = formatar_numero(row.get('Nº', ''))
            ambigua = not valor_numero

            entry = ctk.CTkEntry(scroll, width=70)
            entry.insert(0, valor_numero)
            if ambigua:
                entry.configure(border_color=self.CORES['erro'])
            entry.grid(row=i, column=0, padx=4, pady=2, sticky="w")
            entries_numero[idx_df] = entry

            widgets_linha = [entry]
            for c, nome_col in enumerate(colunas[1:], start=1):
                valor = "" if pd.isna(row.get(nome_col, '')) else str(row.get(nome_col, ''))
                lbl = ctk.CTkLabel(scroll, text=valor, font=ctk.CTkFont(size=10))
                lbl.grid(row=i, column=c, padx=4, pady=2, sticky="w")
                widgets_linha.append(lbl)

            btn_remover = ctk.CTkButton(
                scroll, text="Remover", width=70, height=22, font=ctk.CTkFont(size=9),
                fg_color=self.CORES['erro'], hover_color="#C0392B",
                command=lambda i=idx_df: remover_linha(i)
            )
            btn_remover.grid(row=i, column=len(colunas), padx=4, pady=2, sticky="w")
            widgets_linha.append(btn_remover)
            linha_widgets[idx_df] = widgets_linha

        botoes_frame = ctk.CTkFrame(janela, fg_color="transparent")
        botoes_frame.pack(fill="x", padx=10, pady=10)

        def continuar():
            pendentes = [
                idx_df for idx_df, entry in entries_numero.items()
                if idx_df not in linhas_removidas and not entry.get().strip()
            ]
            if pendentes:
                messagebox.showerror(
                    "Linhas ambíguas pendentes",
                    f"{len(pendentes)} linha(s) ainda estão sem código (Nº). "
                    "Preencha o código ou clique em 'Remover' antes de continuar."
                )
                return

            for idx_df, entry in entries_numero.items():
                if idx_df in linhas_removidas:
                    continue
                valor = entry.get().strip()
                if valor and not valor.isdigit():
                    messagebox.showerror("Valor inválido", f"O código '{valor}' não é numérico. Corrija antes de continuar.")
                    return

            for idx_df, entry in entries_numero.items():
                if idx_df not in linhas_removidas:
                    df.at[idx_df, 'Nº'] = entry.get().strip()

            df_final = df.drop(index=list(linhas_removidas))

            pasta = os.path.dirname(caminho_excel)
            base, _ = os.path.splitext(os.path.basename(caminho_excel))
            caminho_revisado = os.path.join(pasta, f"{base}_revisado.xlsx")
            df_final.to_excel(caminho_revisado, index=False)

            self.caminho_excel_revisado = caminho_revisado
            self.arquivo_excel.set(caminho_revisado)
            self.carregar_preview()
            self.aguardando_revisao_dashboard = False
            janela.destroy()

        def cancelar():
            if messagebox.askyesno("Cancelar", "Cancelar a automação? Nenhum contrato será emitido."):
                self.revisao_cancelada = True
                self.aguardando_revisao_dashboard = False
                janela.destroy()

        ctk.CTkButton(
            botoes_frame, text="Continuar", command=continuar,
            fg_color=self.CORES['sucesso'], hover_color="#27AE60"
        ).pack(side="right", padx=(5, 0))
        ctk.CTkButton(
            botoes_frame, text="Cancelar", command=cancelar,
            fg_color=self.CORES['erro'], hover_color="#C0392B"
        ).pack(side="right")

    def iniciar_automacao(self):
        """Método principal de automação"""
        try:
            self.adicionar_log("Iniciando automação...", logging.INFO, "processando")
            self.status_var.set("Em execução...")
            self.executando = True

            # Resetar barra de progresso
            self.progress_bar.set(0)

            # Iniciar automação
            automacao = DominioAutomation(self.logger, self)

            # Conectar ao Domínio
            if not automacao.connect_to_dominio():
                self.adicionar_log("Não foi possível conectar ao Domínio", logging.ERROR, "erro")
                return

            if self.usar_relatorios_existentes.get():
                self.adicionar_log("Usando relatórios já existentes — pulando Empregados/Contrato", logging.INFO, "info")
            else:
                # Gerar relatório de Empregados (Admissão) em Excel — roda uma vez, antes do loop
                self.adicionar_log("Gerando relatório de Empregados (Admissão)...", logging.INFO, "processando")
                if not automacao.processar_relatorio_empregados():
                    self.adicionar_log("Falha ao gerar relatório de Empregados — abortando execução", logging.ERROR, "erro")
                    return
                self.adicionar_log("Relatório de Empregados gerado com sucesso", logging.INFO, "sucesso")

                # Gerar relatório de Contrato por Prazo Determinado em Excel — roda uma vez, antes do loop
                self.adicionar_log("Gerando relatório de Contrato por Prazo Determinado...", logging.INFO, "processando")
                if not automacao.processar_contrato_prazo_determinado():
                    self.adicionar_log("Falha ao gerar relatório de Contrato por Prazo Determinado — abortando execução", logging.ERROR, "erro")
                    return
                self.adicionar_log("Relatório de Contrato por Prazo Determinado gerado com sucesso", logging.INFO, "sucesso")

            # Gerar Excel principal via Meg_One a partir dos dois relatórios
            self.adicionar_log("Gerando Excel principal via Meg_One...", logging.INFO, "processando")
            excel_principal = self.gerar_excel_principal_meg_one()
            if not excel_principal:
                self.adicionar_log("Falha ao gerar Excel principal via Meg_One — abortando execução", logging.ERROR, "erro")
                return
            self.adicionar_log(f"Excel principal gerado: {excel_principal}", logging.INFO, "sucesso")

            # Remove empresas da lista de "não emitir" (ex: Criat, contrato
            # tratado manualmente por outra equipe) antes de qualquer checagem.
            try:
                df_filtrado = pd.read_excel(excel_principal)
                df_filtrado = self.filtrar_nao_emitir(df_filtrado)
                df_filtrado.to_excel(excel_principal, index=False)
            except Exception as e:
                self.adicionar_log(f"Erro ao aplicar lista de não emitir: {str(e)}", logging.WARNING, "aviso")

            # Só pausa para revisão manual se houver linha ambígua (Nº vazio) —
            # caso contrário, segue direto para a emissão sem interromper o fluxo.
            try:
                df_check = pd.read_excel(excel_principal)
                tem_ambiguidade = df_check['Nº'].isna().any() or (df_check['Nº'].astype(str).str.strip() == '').any()
            except Exception:
                tem_ambiguidade = True  # na dúvida, força revisão

            if tem_ambiguidade:
                self.adicionar_log("Linha(s) ambígua(s) detectada(s) — abrindo revisão do Excel principal", logging.WARNING, "aviso")
                self.aguardando_revisao_dashboard = True
                self.revisao_cancelada = False
                self.caminho_excel_revisado = None
                self.window.after(0, self.abrir_dashboard_revisao, excel_principal)

                while self.aguardando_revisao_dashboard and self.executando:
                    time.sleep(0.5)

                if not self.executando or self.revisao_cancelada:
                    self.adicionar_log("Automação cancelada na revisão do Excel principal", logging.INFO, "aviso")
                    return

                self.arquivo_excel.set(self.caminho_excel_revisado or excel_principal)
            else:
                self.arquivo_excel.set(excel_principal)

            # Carregar Excel
            df = pd.read_excel(self.arquivo_excel.get())
            df_processar = df

            self.total_linhas = len(df_processar)
            self.adicionar_log(f"Arquivo carregado: {self.total_linhas} linhas para processar", logging.INFO, "info")
            self.total_label.configure(text=str(self.total_linhas))

            # Processar linhas
            for idx, (original_index, row) in enumerate(df_processar.iterrows()):
                # Verificar se deve parar
                if not self.executando:
                    self.adicionar_log("Automação interrompida pelo usuário", logging.INFO, "aviso")
                    break

                # Verificar pausa
                while self.pausa_solicitada and self.executando:
                    time.sleep(0.5)

                if not self.executando:
                    break

                # Atualizar progresso
                self.atualizar_progresso(idx + 1, self.total_linhas)

                linha_excel = original_index + 2

                # Linha sem código de empresa (não deveria chegar aqui — o dashboard
                # de revisão exige preencher ou remover — mas pula defensivamente
                # em vez de travar todo o loop com "cannot convert float NaN to integer")
                if pd.isna(row['Nº']) or str(row['Nº']).strip() == '':
                    self.linhas_puladas += 1
                    self.adicionar_log(
                        f"Linha {linha_excel} pulada: sem código de empresa (Nº vazio)",
                        logging.WARNING, "aviso"
                    )
                    self.atualizar_estatisticas()
                    continue

                # Atualizar empresa no card
                empresa_num = str(int(row['Nº']))
                self.empresa_label.configure(text=empresa_num[:20])

                try:
                    func_nome = row.get('Funcionário', 'N/A')
                    self.adicionar_log(
                        f"Processando linha {linha_excel} - Empresa {row['Nº']} - {func_nome}",
                        logging.INFO, "processando"
                    )

                    success = automacao.processar_linha(row, original_index, linha_excel)

                    if success:
                        self.linhas_processadas += 1
                        self.success_logger.info(f"Linha {linha_excel} - Empresa {row['Nº']} - {func_nome} - processada com sucesso")
                        self.adicionar_log(f"Linha {linha_excel} processada com sucesso", logging.INFO, "sucesso")
                    else:
                        self.linhas_com_erro += 1
                        self.error_logger.error(f"Linha {linha_excel} - Empresa {row['Nº']} - {func_nome} - erro no processamento")
                        self.adicionar_log(f"Erro na linha {linha_excel}", logging.ERROR, "erro")
                        time.sleep(2)  # aguarda sistema estabilizar antes de processar próxima linha

                    self.atualizar_estatisticas()

                except Exception as e:
                    self.linhas_com_erro += 1
                    erro_msg = f"Linha {linha_excel} - Erro: {str(e)}"
                    self.error_logger.error(erro_msg)
                    self.adicionar_log(erro_msg, logging.ERROR, "erro")
                    self.atualizar_estatisticas()

            # Finalização
            if self.executando:
                self.status_var.set("Processamento concluído")
                self.progress_bar.set(1.0)
                self.progress_label.configure(text="100%")
                self.atualizar_status_indicator('concluido')
                self.adicionar_log("Automação concluída!", logging.INFO, "sucesso")
                self.adicionar_log(f"Resumo: {self.linhas_processadas} processadas, {self.linhas_com_erro} com erro, {self.linhas_puladas} puladas", logging.INFO, "info")

                # Corrige spans fragmentados em todos os PDFs da pasta do dia
                try:
                    pasta_pdfs = self.obter_pasta_dia() if self.diretorio_salvamento.get() else ""
                    if pasta_pdfs and os.path.isdir(pasta_pdfs):
                        pdfs = list(Path(pasta_pdfs).glob("*.pdf"))
                        self.adicionar_log(f"Corrigindo spans em {len(pdfs)} PDF(s) da pasta...", logging.INFO, "info")
                        total_corrigidas = 0
                        total_com_fragmentacao = 0
                        total_paginas_removidas = 0
                        arquivos_com_pagina_branca = 0
                        falhas = 0
                        for pdf_path in pdfs:
                            try:
                                stats = corrigir_pdf(pdf_path, output_path=pdf_path)
                                if stats["fragmentadas"] > 0:
                                    total_com_fragmentacao += 1
                                    total_corrigidas += stats["corrigidas"]
                                removidas = stats.get("paginas_removidas", [])
                                if removidas:
                                    arquivos_com_pagina_branca += 1
                                    total_paginas_removidas += len(removidas)
                                    self.adicionar_log(
                                        f"📄 Página(s) em branco removida(s) de {pdf_path.name}: "
                                        f"{', '.join(str(p) for p in removidas)}",
                                        logging.INFO, "info"
                                    )
                            except Exception as e_corr:
                                falhas += 1
                                self.error_logger.error(f"Correção de spans falhou em {pdf_path.name}: {e_corr}")
                        self.adicionar_log(
                            f"Correção de spans concluída: {total_com_fragmentacao} arquivo(s) corrigido(s) "
                            f"({total_corrigidas} linha(s)), {falhas} falha(s)",
                            logging.INFO, "sucesso" if falhas == 0 else "aviso"
                        )
                        if total_paginas_removidas:
                            self.adicionar_log(
                                f"Páginas em branco removidas: {total_paginas_removidas} "
                                f"em {arquivos_com_pagina_branca} arquivo(s)",
                                logging.INFO, "sucesso"
                            )
                except Exception as e:
                    self.adicionar_log(f"Erro ao corrigir spans na pasta: {str(e)}", logging.WARNING, "aviso")

                # Enviar notificação ao Discord via webhook
                try:
                    data_referencia = self.periodo_referencia.get() or (datetime.now() + timedelta(days=1)).strftime("%d/%m/%Y")
                    diretorio_pdfs = self.obter_pasta_dia() if self.diretorio_salvamento.get() else os.path.dirname(os.path.abspath(self.arquivo_excel.get()))
                    mensagem = (
                        f"📋 **Contratos Admissionais Emitidos**\n\n"
                        f"📅 **Período de referência:** {data_referencia}\n"
                        f"📊 **Quantidade emitida:** {self.linhas_processadas}\n"
                        f"📂 **Diretório dos PDFs:** `{diretorio_pdfs}`\n\n"
                        f"✅ Emissão finalizada com sucesso!\n\n"
                        f"<@&1299044385899548752>"
                    )
                    webhook_url = os.getenv("DISCORD_WEBHOOK_URL")
                    requests.post(webhook_url, json={"content": mensagem}, timeout=10)
                    self.adicionar_log("Notificação enviada ao Discord", logging.INFO, "sucesso")
                except Exception as e:
                    self.adicionar_log(f"Erro ao enviar notificação ao Discord: {str(e)}", logging.WARNING, "aviso")

        except Exception as e:
            erro_msg = f"Erro crítico: {str(e)}"
            self.error_logger.error(erro_msg)
            self.adicionar_log(erro_msg, logging.ERROR, "erro")
            self.status_var.set("Erro no processamento")
            self.atualizar_status_indicator('erro')
        finally:
            self.executando = False
            self.pausa_solicitada = False
            self.atualizar_tempo()  # Atualização final do timer
            self.btn_iniciar.configure(state="normal")
            self.btn_emitir_manual.configure(state="normal")
            self.btn_pausar.configure(state="disabled", text="⏸ Pausar")
            self.btn_parar.configure(state="disabled")

    def iniciar_automacao_manual(self):
        """Emissão manual: lê a planilha fornecida pelo usuário e emite os
        contratos direto, sem gerar os relatórios de Empregados/Contrato, sem
        Meg_One, sem dashboard de revisão e sem o filtro de 'Não Emitir' — o
        modo manual é justamente a exceção a essas regras."""
        try:
            self.adicionar_log("Iniciando emissão MANUAL...", logging.INFO, "processando")
            self.status_var.set("Em execução (manual)...")
            self.executando = True
            self.progress_bar.set(0)

            linha_inicial = int(self.linha_inicial_manual.get())

            automacao = DominioAutomation(self.logger, self)

            if not automacao.connect_to_dominio():
                self.adicionar_log("Não foi possível conectar ao Domínio", logging.ERROR, "erro")
                return

            # Carregar planilha manual, aplicando a linha inicial (linha 2 = índice 0)
            df = pd.read_excel(self.arquivo_excel_manual.get())
            df_processar = df.iloc[linha_inicial - 2:]

            self.total_linhas = len(df_processar)
            self.adicionar_log(
                f"[Manual] Planilha carregada: {self.total_linhas} linhas para processar "
                f"(a partir da linha {linha_inicial})",
                logging.INFO, "info"
            )
            self.total_label.configure(text=str(self.total_linhas))

            for idx, (original_index, row) in enumerate(df_processar.iterrows()):
                if not self.executando:
                    self.adicionar_log("Emissão manual interrompida pelo usuário", logging.INFO, "aviso")
                    break

                while self.pausa_solicitada and self.executando:
                    time.sleep(0.5)

                if not self.executando:
                    break

                self.atualizar_progresso(idx + 1, self.total_linhas)

                linha_excel = original_index + 2

                if pd.isna(row['Nº']) or str(row['Nº']).strip() == '':
                    self.linhas_puladas += 1
                    self.adicionar_log(
                        f"[Manual] Linha {linha_excel} pulada: sem código de empresa (Nº vazio)",
                        logging.WARNING, "aviso"
                    )
                    self.atualizar_estatisticas()
                    continue

                empresa_num = str(int(row['Nº']))
                self.empresa_label.configure(text=empresa_num[:20])

                try:
                    func_nome = row.get('Funcionário', 'N/A')
                    self.adicionar_log(
                        f"[Manual] Processando linha {linha_excel} - Empresa {row['Nº']} - {func_nome}",
                        logging.INFO, "processando"
                    )

                    success = automacao.processar_linha(row, original_index, linha_excel)

                    if success:
                        self.linhas_processadas += 1
                        self.success_logger.info(
                            f"[MANUAL] Linha {linha_excel} - Empresa {row['Nº']} - {func_nome} - processada com sucesso"
                        )
                        self.adicionar_log(f"[Manual] Linha {linha_excel} processada com sucesso", logging.INFO, "sucesso")
                    else:
                        self.linhas_com_erro += 1
                        self.error_logger.error(
                            f"[MANUAL] Linha {linha_excel} - Empresa {row['Nº']} - {func_nome} - erro no processamento"
                        )
                        self.adicionar_log(f"[Manual] Erro na linha {linha_excel}", logging.ERROR, "erro")
                        time.sleep(2)  # aguarda sistema estabilizar antes da próxima linha

                    self.atualizar_estatisticas()

                except Exception as e:
                    self.linhas_com_erro += 1
                    erro_msg = f"[MANUAL] Linha {linha_excel} - Erro: {str(e)}"
                    self.error_logger.error(erro_msg)
                    self.adicionar_log(erro_msg, logging.ERROR, "erro")
                    self.atualizar_estatisticas()

            if self.executando:
                self.status_var.set("Emissão manual concluída")
                self.progress_bar.set(1.0)
                self.progress_label.configure(text="100%")
                self.atualizar_status_indicator('concluido')
                self.adicionar_log("Emissão manual concluída!", logging.INFO, "sucesso")
                self.adicionar_log(
                    f"Resumo: {self.linhas_processadas} processadas, "
                    f"{self.linhas_com_erro} com erro, {self.linhas_puladas} puladas",
                    logging.INFO, "info"
                )

                # Corrige spans fragmentados nos PDFs da pasta manual
                try:
                    pasta_pdfs = self.diretorio_manual.get()
                    if pasta_pdfs and os.path.isdir(pasta_pdfs):
                        pdfs = list(Path(pasta_pdfs).glob("*.pdf"))
                        self.adicionar_log(f"[Manual] Corrigindo spans em {len(pdfs)} PDF(s)...", logging.INFO, "info")
                        total_corrigidas = 0
                        total_com_fragmentacao = 0
                        total_paginas_removidas = 0
                        arquivos_com_pagina_branca = 0
                        falhas = 0
                        for pdf_path in pdfs:
                            try:
                                stats = corrigir_pdf(pdf_path, output_path=pdf_path)
                                if stats["fragmentadas"] > 0:
                                    total_com_fragmentacao += 1
                                    total_corrigidas += stats["corrigidas"]
                                removidas = stats.get("paginas_removidas", [])
                                if removidas:
                                    arquivos_com_pagina_branca += 1
                                    total_paginas_removidas += len(removidas)
                                    self.adicionar_log(
                                        f"[Manual] 📄 Página(s) em branco removida(s) de {pdf_path.name}: "
                                        f"{', '.join(str(p) for p in removidas)}",
                                        logging.INFO, "info"
                                    )
                            except Exception as e_corr:
                                falhas += 1
                                self.error_logger.error(f"[MANUAL] Correção de spans falhou em {pdf_path.name}: {e_corr}")
                        self.adicionar_log(
                            f"[Manual] Correção de spans concluída: {total_com_fragmentacao} arquivo(s) corrigido(s) "
                            f"({total_corrigidas} linha(s)), {falhas} falha(s)",
                            logging.INFO, "sucesso" if falhas == 0 else "aviso"
                        )
                        if total_paginas_removidas:
                            self.adicionar_log(
                                f"[Manual] Páginas em branco removidas: {total_paginas_removidas} "
                                f"em {arquivos_com_pagina_branca} arquivo(s)",
                                logging.INFO, "sucesso"
                            )
                except Exception as e:
                    self.adicionar_log(f"[Manual] Erro ao corrigir spans: {str(e)}", logging.WARNING, "aviso")

                # Notificação ao Discord
                try:
                    data_referencia = self.periodo_referencia.get() or (datetime.now() + timedelta(days=1)).strftime("%d/%m/%Y")
                    mensagem = (
                        f"📋 **Contratos Admissionais Emitidos (MANUAL)**\n\n"
                        f"📅 **Período de referência:** {data_referencia}\n"
                        f"📊 **Quantidade emitida:** {self.linhas_processadas}\n"
                        f"📂 **Diretório dos PDFs:** `{self.diretorio_manual.get()}`\n\n"
                        f"✅ Emissão manual finalizada com sucesso!\n\n"
                        f"<@&1299044385899548752>"
                    )
                    webhook_url = os.getenv("DISCORD_WEBHOOK_URL")
                    requests.post(webhook_url, json={"content": mensagem}, timeout=10)
                    self.adicionar_log("Notificação enviada ao Discord", logging.INFO, "sucesso")
                except Exception as e:
                    self.adicionar_log(f"Erro ao enviar notificação ao Discord: {str(e)}", logging.WARNING, "aviso")

        except Exception as e:
            erro_msg = f"Erro crítico na emissão manual: {str(e)}"
            self.error_logger.error(erro_msg)
            self.adicionar_log(erro_msg, logging.ERROR, "erro")
            self.status_var.set("Erro no processamento")
            self.atualizar_status_indicator('erro')
        finally:
            self.executando = False
            self.pausa_solicitada = False
            self.modo_manual = False
            self.atualizar_tempo()
            self.btn_iniciar.configure(state="normal")
            self.btn_emitir_manual.configure(state="normal")
            self.btn_pausar.configure(state="disabled", text="⏸ Pausar")
            self.btn_parar.configure(state="disabled")

    def executar(self):
        self.window.mainloop()

class DominioAutomation:
    def __init__(self, logger, gui):
        timings.Timings.window_find_timeout = 20
        self.app = None
        self.main_window = None
        self.logger = logger
        self.gui = gui
        self.empresa_atual = None

    def log(self, message):
        self.logger.info(message)

    def should_stop(self) -> bool:
        """Verifica se deve parar a execução"""
        return not self.gui.executando

    def check_pause(self):
        """Verifica e aguarda se pausado"""
        while self.gui.pausa_solicitada and self.gui.executando:
            time.sleep(0.5)

    def smart_sleep(self, seconds: float):
        """Sleep interruptível que verifica pausa/parada"""
        interval = 0.15
        elapsed = 0.0
        while elapsed < seconds:
            if self.should_stop():
                return False
            self.check_pause()
            if self.should_stop():
                return False
            sleep_time = min(interval, seconds - elapsed)
            time.sleep(sleep_time)
            elapsed += sleep_time
        return True

    def wait_for_condition(self, condition_fn, timeout: float = 30, poll_interval: float = 0.15, description: str = "") -> bool:
        """Polls condition_fn() até retornar True, ou timeout."""
        start = time.time()
        while time.time() - start < timeout:
            if self.should_stop():
                return False
            self.check_pause()
            try:
                if condition_fn():
                    if description:
                        self.log(f"{description} - concluido em {time.time() - start:.1f}s")
                    return True
            except Exception:
                pass
            time.sleep(poll_interval)
        if description:
            self.log(f"{description} - timeout apos {timeout}s")
        return False

    def calcular_periodo_mes_referencia(self) -> Tuple[str, str]:
        """Retorna (data_inicio, data_fim) — ambas a mesma data do campo
        self.gui.periodo_referencia (formato "%d/%m/%Y"), já que a automação
        roda diariamente (uma execução por dia, não por mês). Em caso de erro
        de parse, usa a data de hoje."""
        valor = ""
        try:
            valor = self.gui.periodo_referencia.get()
            ref = datetime.strptime(valor.strip(), "%d/%m/%Y")
        except Exception:
            self.log(f"⚠️ Não foi possível interpretar periodo_referencia ('{valor}'), usando data atual")
            ref = datetime.now()

        data = ref.strftime("%d/%m/%Y")
        self.log(f"📅 Período de admissão: {data}")
        return data, data

    def _force_focus(self, hwnd: int):
        """Força foco para uma janela contornando a restrição do Windows 10."""
        try:
            # Simula um pressionamento de Alt para desbloquear SetForegroundWindow
            ctypes.windll.user32.keybd_event(0x12, 0, 0, 0)
            ctypes.windll.user32.keybd_event(0x12, 0, 2, 0)
            # Só restaura se estiver minimizada — SW_RESTORE desmaximizaria uma
            # janela já maximizada, encolhendo-a sem necessidade.
            if win32gui.IsIconic(hwnd):
                win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            win32gui.SetForegroundWindow(hwnd)
            time.sleep(0.3)
        except Exception:
            pass

    def _window_exists(self, title: str, class_name: str) -> bool:
        """Verifica se janela com título/classe existe via win32gui."""
        try:
            result = [False]
            def callback(hwnd, _):
                if win32gui.IsWindowVisible(hwnd):
                    try:
                        if (win32gui.GetWindowText(hwnd) == title and
                                win32gui.GetClassName(hwnd) == class_name):
                            result[0] = True
                            return False
                    except Exception:
                        pass
                return True
            win32gui.EnumWindows(callback, None)
            return result[0]
        except Exception:
            return False

    def _window_exists_by_class(self, class_name: str) -> bool:
        """Verifica se janela com classe existe via win32gui (ignora título)."""
        try:
            result = [False]
            def callback(hwnd, _):
                if win32gui.IsWindowVisible(hwnd):
                    try:
                        if win32gui.GetClassName(hwnd) == class_name:
                            result[0] = True
                            return False
                    except Exception:
                        pass
                return True
            win32gui.EnumWindows(callback, None)
            return result[0]
        except Exception:
            return False

    def _save_dialog_exists(self) -> bool:
        """Verifica se a janela de salvamento existe procurando pelo elemento 'Salvar em:' (AutomationId 1091)."""
        try:
            result = [False]
            def callback(hwnd, _):
                if win32gui.IsWindowVisible(hwnd):
                    try:
                        if win32gui.GetClassName(hwnd) == "#32770":
                            # Procura pelo controle filho "Salvar em:" (Static, AutomationId=1091)
                            child = win32gui.FindWindowEx(hwnd, 0, "Static", "Salvar em:")
                            if child:
                                result[0] = True
                                return False
                    except Exception:
                        pass
                return True
            win32gui.EnumWindows(callback, None)
            return result[0]
        except Exception:
            return False

    def _any_error_dialog_visible(self) -> bool:
        """Verifica se há diálogo de erro visível via win32gui."""
        error_keywords = ("erro", "aviso", "atenção", "alerta", "warning", "error", "informação")
        try:
            result = [False]
            def callback(hwnd, _):
                if win32gui.IsWindowVisible(hwnd):
                    try:
                        cls = win32gui.GetClassName(hwnd)
                        if cls == "#32770":
                            title = win32gui.GetWindowText(hwnd).lower()
                            for kw in error_keywords:
                                if kw in title:
                                    result[0] = True
                                    return False
                    except Exception:
                        pass
                return True
            win32gui.EnumWindows(callback, None)
            return result[0]
        except Exception:
            return False

    def _is_connection_alive(self) -> bool:
        """Verifica se a conexão pywinauto ainda é válida."""
        if self.app is None or self.main_window is None:
            return False
        try:
            hwnd = self.main_window.handle
            if not win32gui.IsWindow(hwnd):
                return False
            win32gui.GetWindowText(hwnd)
            return True
        except Exception:
            return False

    def find_dominio_window(self) -> Optional[int]:
        """Encontra a janela do Domínio Folha"""
        try:
            self.log("🔍 Procurando janela do Domínio Folha...")

            try:
                all_windows = findwindows.find_windows()
                self.log(f"📋 Total de janelas abertas: {len(all_windows)}")

                for hwnd in all_windows:
                    try:
                        title = win32gui.GetWindowText(hwnd)
                        if "Domínio" in title and title:
                            self.log(f"🪟 Janela encontrada: '{title}'")
                            if "Folha" in title:
                                self.log(f"✅ Janela do Domínio Folha localizada!")
                                return hwnd
                    except Exception:
                        continue
            except Exception as e:
                self.log(f"⚠️ Erro ao listar janelas: {str(e)}")

            # Fallback: tentar o método original com regex
            windows = findwindows.find_windows(title_re=".*Domínio Folha.*")
            if windows:
                self.log(f"✅ Janela do Domínio encontrada via regex (total: {len(windows)})")
                return windows[0]

            self.log("❌ Nenhuma janela do Domínio Folha encontrada")
            return None
        except Exception as e:
            self.log(f"❌ Erro ao procurar janela do Domínio: {str(e)}")
            self.log(f"Traceback: {traceback.format_exc()}")
            return None

    def connect_to_dominio(self) -> bool:
        """Conecta à aplicação Domínio"""
        try:
            handle = self.find_dominio_window()
            if not handle:
                return False

            if win32gui.IsIconic(handle):
                win32gui.ShowWindow(handle, win32con.SW_RESTORE)
                time.sleep(1)

            win32gui.SetForegroundWindow(handle)
            time.sleep(0.5)

            self.app = Application(backend="uia").connect(handle=handle)
            self.main_window = self.app.window(handle=handle)

            self.log("✅ Conectado ao Domínio Folha com sucesso")
            return True

        except Exception as e:
            self.log(f"❌ Erro ao conectar ao Domínio: {str(e)}")
            return False

    def wait_for_window_close(self, window, window_title: str, timeout: int = 30) -> bool:
        """Espera até que uma janela seja fechada"""
        start_time = time.time()
        while time.time() - start_time < timeout:
            if self.should_stop():
                return False
            self.check_pause()

            try:
                if not window.exists() or not window.is_visible():
                    self.log(f"✅ Janela '{window_title}' fechada")
                    return True
            except Exception:
                return True

            self.handle_error_dialogs()
            time.sleep(0.15)

        self.log(f"⚠️ Timeout aguardando fechamento da janela '{window_title}'")
        return False

    def handle_empresa_change(self, empresa_num: str) -> bool:
        """Gerencia a troca de empresa"""
        try:
            if self.should_stop():
                return False

            self.log("📞 Solicitando troca de empresa (F8)")
            self.main_window.set_focus()
            if not self.smart_sleep(0.3):
                return False
            send_keys('{F8}')
            if not self.smart_sleep(2):
                return False

            max_attempts = 10
            troca_window = None

            for attempt in range(max_attempts):
                if self.should_stop():
                    return False
                self.check_pause()

                try:
                    troca_window = self.main_window.child_window(
                        title="Troca de empresas",
                        class_name="FNWND3190"
                    )

                    if troca_window.exists():
                        break

                    if not self.handle_error_dialogs():
                        self.cleanup_windows()
                        return False

                    if not self.smart_sleep(0.5):
                        return False
                except Exception:
                    if attempt == max_attempts - 1:
                        self.log("❌ Janela 'Troca de empresas' não encontrada (timeout)")
                        return False
                    if not self.smart_sleep(1):
                        return False

            if not troca_window:
                self.log("❌ Janela 'Troca de empresas' não encontrada")
                return False

            self.log(f"🏢 Alterando para empresa: {empresa_num}")

            # Resolve o wrapper antes de acessar o handle
            troca_wrapper = troca_window.wrapper_object()
            troca_hwnd = troca_wrapper.handle

            self._force_focus(troca_hwnd)
            if not self.smart_sleep(0.3):
                return False

            # Tenta localizar o campo Edit e preencher com set_focus + type_keys (mais confiável no Domínio)
            edit_hwnd = win32gui.FindWindowEx(troca_hwnd, 0, "Edit", None)
            if edit_hwnd:
                self.log("🖊️ Campo encontrado, preenchendo via pywinauto type_keys")
                try:
                    edit_wrapper = self.app.window(handle=edit_hwnd)
                    edit_wrapper.set_focus()
                    time.sleep(0.2)
                    # Limpa conteúdo atual e digita o número da empresa
                    edit_wrapper.type_keys('^a', with_spaces=False)
                    edit_wrapper.type_keys(empresa_num, with_spaces=False)
                    time.sleep(0.3)
                    edit_wrapper.type_keys('{ENTER}', with_spaces=False)
                except Exception:
                    # Fallback: send_keys direto na janela com foco
                    self.log("⚠️ Fallback para send_keys no campo Edit")
                    win32gui.SetForegroundWindow(edit_hwnd)
                    time.sleep(0.2)
                    send_keys('^a')
                    send_keys(empresa_num)
                    time.sleep(0.3)
                    send_keys('{ENTER}')
            else:
                self.log("⚠️ Campo Edit não encontrado, usando send_keys na janela")
                troca_wrapper.set_focus()
                if not self.smart_sleep(0.3):
                    return False
                send_keys('^a')
                send_keys(empresa_num)
                if not self.smart_sleep(0.3):
                    return False
                send_keys('{ENTER}')

            if not self.smart_sleep(1.5):
                return False

            if not self.handle_error_dialogs():
                self.cleanup_windows()
                return False

            self.wait_for_window_close(troca_window, "Troca de empresas")

            self.close_avisos_vencimento()

            return True

        except Exception as e:
            self.log(f"❌ Erro na troca de empresa: {str(e)}")
            return False

    def close_avisos_vencimento(self):
        """Fecha janela de avisos de vencimento se estiver aberta"""
        try:
            aviso_window = self.main_window.child_window(
                title="Avisos de Vencimento",
                class_name="FNWND3190"
            )

            if aviso_window.exists() and aviso_window.is_visible():
                self.log("📋 Fechando 'Avisos de Vencimento'")
                aviso_window.set_focus()
                send_keys('{ESC}')
                time.sleep(0.5)
                send_keys('{ESC}')
                time.sleep(0.5)
        except Exception:
            pass

    def processar_linha(self, row, index: int, linha_excel: int) -> bool:
        """Processa uma linha do Excel"""
        try:
            if self.should_stop():
                return False

            # Só reconectar se a conexão existente estiver quebrada
            if not self._is_connection_alive():
                handle = self.find_dominio_window()
                if not handle:
                    self.log("❌ Não foi possível localizar a janela do Domínio")
                    return False
                try:
                    self.app = Application(backend="uia").connect(handle=handle)
                    self.main_window = self.app.window(handle=handle)
                    self.log("✅ Reconectado ao Domínio com sucesso")
                except Exception as e:
                    self.log(f"❌ Erro ao reconectar: {str(e)}")
                    return False

            handle = self.main_window.handle

            if win32gui.IsIconic(handle):
                win32gui.ShowWindow(handle, win32con.SW_RESTORE)
                if not self.smart_sleep(0.5):
                    return False

            win32gui.SetForegroundWindow(handle)
            time.sleep(0.2)

            # Troca de empresa (só troca se for diferente da atual)
            empresa_num = str(int(row['Nº']))
            if empresa_num != self.empresa_atual:
                if not self.handle_empresa_change(empresa_num):
                    return False
                self.empresa_atual = empresa_num
            else:
                self.log(f"🏢 Empresa {empresa_num} já selecionada, pulando troca")

            if self.should_stop():
                return False
            self.check_pause()

            # Acessar Relatórios > Cadastrais
            self.log("📊 Acessando Relatórios > Cadastrais")
            self.main_window.set_focus()
            send_keys('%r')  # ALT+R
            if not self.smart_sleep(0.5):
                return False
            send_keys('c')  # Cadastrais
            if not self.smart_sleep(0.5):
                return False
            send_keys('{DOWN}{DOWN}{DOWN}')  # 3x seta para baixo
            if not self.smart_sleep(0.5):
                return False
            send_keys('{ENTER}')
            if not self.smart_sleep(1):
                return False

            # Processar no Relatórios Admissionais
            return self.processar_relatorio_admissional(row, linha_excel)

        except Exception as e:
            self.log(f"❌ Erro ao processar linha {linha_excel}: {str(e)}")
            return False

    def processar_relatorio_admissional(self, row, linha_excel: int) -> bool:
        """Processa o relatório admissional"""
        try:
            if self.should_stop():
                return False

            # Aguardar janela Relatórios Admissionais
            max_attempts = 10
            admiss_window = None

            for attempt in range(max_attempts):
                if self.should_stop():
                    return False
                self.check_pause()

                try:
                    admiss_window = self.main_window.child_window(
                        title="Relatórios Admissionais",
                        class_name="FNWND3190"
                    )

                    if admiss_window.exists():
                        break

                    if not self.handle_error_dialogs():
                        self.cleanup_windows()
                        return False

                    if not self.smart_sleep(1):
                        return False
                except Exception:
                    if attempt == max_attempts - 1:
                        self.log("❌ Relatórios Admissionais não encontrado (timeout)")
                        return False

            if not admiss_window:
                self.log("❌ Relatórios Admissionais não encontrado")
                return False

            self.log("📋 Relatórios Admissionais localizado")

            if self.should_stop():
                return False
            self.check_pause()

            # Preencher código do funcionário (3 TABs para chegar ao campo)
            cod_funcionario = str(int(row['Cod.Funcionário']))
            self.log(f"👤 Preenchendo código do funcionário: {cod_funcionario}")

            send_keys('{TAB}{TAB}{TAB}')
            time.sleep(0.3)
            send_keys(cod_funcionario)
            if not self.smart_sleep(0.5):
                return False

            # Acessar aba Relatórios via clique por coordenadas relativas
            self.log("📑 Acessando aba Relatórios")
            try:
                tab_control = admiss_window.child_window(auto_id="1004", class_name="PBTabControl32_100")
                # Clicar na aba "Relatórios" — segunda aba, posição aproximada (após aba "Geral")
                tab_control.click_input(coords=(70, 10))
                if not self.smart_sleep(0.5):
                    return False
            except Exception as e:
                self.log(f"⚠️ Erro ao clicar na aba Relatórios: {str(e)}")
                return False

            # Selecionar tipo de contrato
            tipo_contrato = str(row['Tipo de Contrato']).strip().upper()
            self.log(f"📝 Tipo de contrato: {'Experiência' if tipo_contrato == 'E' else 'Prazo Indeterminado'}")

            # Checkbox Contrato de Experiência (AutomationId: 1031)
            # Checkbox Contrato por Prazo Indeterminado (AutomationId: 1030)
            try:
                cb_experiencia = admiss_window.child_window(auto_id="1031", class_name="Button")
                cb_indeterminado = admiss_window.child_window(auto_id="1030", class_name="Button")

                if tipo_contrato == 'E':
                    # Marcar Experiência, desmarcar Indeterminado
                    if not cb_experiencia.get_toggle_state():
                        cb_experiencia.click_input()
                        time.sleep(0.3)
                    if cb_indeterminado.get_toggle_state():
                        cb_indeterminado.click_input()
                        time.sleep(0.3)
                    self.log("✅ Contrato de Experiência selecionado")
                elif tipo_contrato == 'I':
                    # Marcar Indeterminado, desmarcar Experiência
                    if not cb_indeterminado.get_toggle_state():
                        cb_indeterminado.click_input()
                        time.sleep(0.3)
                    if cb_experiencia.get_toggle_state():
                        cb_experiencia.click_input()
                        time.sleep(0.3)
                    self.log("✅ Contrato por Prazo Indeterminado selecionado")
                else:
                    self.log(f"❌ Tipo de contrato inválido: {tipo_contrato}")
                    return False

            except Exception as e:
                self.log(f"❌ Erro ao selecionar tipo de contrato: {str(e)}")
                return False

            if self.should_stop():
                return False
            self.check_pause()

            # Clicar no botão OK via atalho Alt+O
            self.log("⚡ Clicando em OK para gerar relatório")
            send_keys('%o')  # Alt+O

            # Enviar Ctrl+D em loop aguardando a janela "Salvar em PDF"
            self.log("📄 Aguardando janela de salvamento (enviando Ctrl+D periodicamente)...")
            timeout_total = 180  # 3 minutos no total
            intervalo_ctrl_d = 5  # enviar Ctrl+D a cada 5 segundos
            inicio = time.time()
            janela_encontrada = False

            while time.time() - inicio < timeout_total:
                if self.should_stop():
                    return False

                # Diálogo de "já existe / substituir" não é erro — confirma e segue
                self.handle_sobrescrever_dialog()

                # Verificar diálogos de erro/aviso
                if self._any_error_dialog_visible():
                    deve_continuar = self.handle_error_dialogs()
                    if deve_continuar:
                        # Aviso não crítico (ex: erro de gravação) — janela de salvamento deve abrir após fechar
                        self.log("🔄 Aviso tratado, continuando aguardo da janela de salvamento...")
                        time.sleep(0.5)
                        continue
                    else:
                        # Erro que aborta (ex: sem dados para emitir) — limpa e sai
                        self.log("⚠️ Diálogo de erro detectado, abortando linha")
                        self.cleanup_windows()
                        return False

                if self._window_exists("Salvar em PDF", "#32770") or self._save_dialog_exists():
                    janela_encontrada = True
                    break
                send_keys('^d')  # Ctrl+D
                for _ in range(int(intervalo_ctrl_d / 0.25)):
                    if self.should_stop():
                        return False
                    if self._any_error_dialog_visible():
                        break  # Volta pro while para tratar o diálogo
                    if self._window_exists("Salvar em PDF", "#32770") or self._save_dialog_exists():
                        janela_encontrada = True
                        break
                    time.sleep(0.25)
                if janela_encontrada:
                    break

            if not janela_encontrada:
                self.log("❌ Timeout aguardando janela de salvamento")
                return False

            # Verificar e tratar janela de erro
            if not self.handle_error_dialogs():
                self.cleanup_windows()
                return False

            # Localizar janela de salvamento e preencher nome
            return self.salvar_pdf(row, linha_excel)

        except Exception as e:
            self.log(f"❌ Erro no processamento do relatório admissional: {str(e)}")
            return False

    def _find_save_window_hwnd(self) -> int:
        """Retorna o hwnd da janela de salvamento de PDF, ou 0 se não encontrada."""
        result = [0]
        def callback(hwnd, _):
            if win32gui.IsWindowVisible(hwnd):
                try:
                    if win32gui.GetClassName(hwnd) == "#32770":
                        if win32gui.GetWindowText(hwnd) == "Salvar em PDF":
                            result[0] = hwnd
                            return False
                        child = win32gui.FindWindowEx(hwnd, 0, "Static", "Salvar em:")
                        if child:
                            result[0] = hwnd
                            return False
                except Exception:
                    pass
            return True
        try:
            win32gui.EnumWindows(callback, None)
        except Exception:
            pass
        return result[0]

    def _save_excel_dialog_exists(self) -> bool:
        """Verifica se o diálogo 'Selecione o arquivo' (exportação para Excel,
        disparado pelo botão de Excel na barra lateral, não por Ctrl+D) está
        visível. Confirmado: título exato 'Selecione o arquivo', combo Tipo
        já vem pré-selecionado em 'Arquivos do Excel (*.xls)'."""
        try:
            result = [False]
            def callback(hwnd, _):
                if win32gui.IsWindowVisible(hwnd):
                    try:
                        if win32gui.GetClassName(hwnd) == "#32770" and win32gui.GetWindowText(hwnd) == "Selecione o arquivo":
                            result[0] = True
                            return False
                    except Exception:
                        pass
                return True
            win32gui.EnumWindows(callback, None)
            return result[0]
        except Exception:
            return False

    def _find_save_excel_window_hwnd(self) -> int:
        """Retorna o hwnd do diálogo 'Selecione o arquivo', ou 0 se não encontrado."""
        result = [0]
        def callback(hwnd, _):
            if win32gui.IsWindowVisible(hwnd):
                try:
                    if win32gui.GetClassName(hwnd) == "#32770" and win32gui.GetWindowText(hwnd) == "Selecione o arquivo":
                        result[0] = hwnd
                        return False
                except Exception:
                    pass
            return True
        try:
            win32gui.EnumWindows(callback, None)
        except Exception:
            pass
        return result[0]

    def salvar_pdf(self, row, linha_excel: int) -> bool:
        """Salva o relatório como PDF"""
        try:
            if self.should_stop():
                return False

            self.log("💾 Configurando salvamento do PDF")

            # Busca a janela de salvamento pelo hwnd (pode ser top-level, não filha de main_window)
            save_hwnd = self._find_save_window_hwnd()
            if save_hwnd:
                save_app = Application(backend="uia").connect(handle=save_hwnd)
                save_window = save_app.window(handle=save_hwnd)
            else:
                # Fallback: tenta como filho de main_window
                self.log("🔍 Procurando janela de salvamento alternativa...")
                save_window = self.main_window.child_window(
                    title="Salvar em PDF",
                    class_name="#32770"
                )
                if not save_window.exists():
                    self.log("❌ Janela de salvamento não encontrada")
                    return False
                self.log("✅ Janela de salvamento encontrada via child_window")

            if self.should_stop():
                return False
            self.check_pause()

            # Navegar até campo de nome do arquivo
            nome_pdf = str(row['Documento'])
            # Remove apenas caracteres inválidos em nomes de arquivo (não inclui '\' e ':' pois são do diretório)
            for ch in r'*?"<>|':
                nome_pdf = nome_pdf.replace(ch, '-')

            # Se o usuário selecionou um diretório, montar caminho absoluto separando diretório do nome.
            # No modo manual o destino é o diretório próprio da aba Manual.
            if getattr(self.gui, 'modo_manual', False):
                diretorio = self.gui.diretorio_manual.get()
            else:
                diretorio = self.gui.diretorio_salvamento.get()
            if diretorio:
                caminho_completo = os.path.join(diretorio, nome_pdf)
                self.log(f"📝 Salvando em: {caminho_completo}")
            else:
                caminho_completo = nome_pdf
                self.log(f"📝 Nome do arquivo: {nome_pdf}")

            time.sleep(0.5)

            # Coloca o caminho diretamente no campo de nome via WM_SETTEXT (evita
            # que send_keys interprete '\', caracteres acentuados ou '-' como teclas especiais)
            self._force_focus(save_hwnd)
            filename_edit = win32gui.FindWindowEx(save_hwnd, 0, "Edit", None)
            if filename_edit:
                win32gui.SendMessage(filename_edit, win32con.WM_SETTEXT, 0, caminho_completo)
                time.sleep(0.2)
            else:
                # Fallback: seleciona tudo e digita (caminho sem caracteres especiais de controle)
                send_keys('^a')
                send_keys(caminho_completo, with_spaces=True)
            time.sleep(0.3)

            if self.should_stop():
                return False
            self.check_pause()

            # Confirmar com Enter
            self.log("💾 Salvando PDF")
            send_keys('{ENTER}')

            # Esperar o salvamento concluir, confirmando a substituição se o PDF
            # já existir (reemissão para a mesma data) — ver _criar_checador_salvamento.
            if not self.wait_for_condition(
                self._criar_checador_salvamento(save_hwnd),
                timeout=30,
                poll_interval=0.3,
                description="Aguardando salvamento do PDF"
            ):
                self.log("⚠️ Timeout aguardando janela fechar, continuando mesmo assim")

            self.log(f"✅ PDF salvo: {nome_pdf}")

            # Aguarda o sistema de arquivos confirmar a gravação antes de continuar
            time.sleep(5.0)

            # Correção de spans fragmentados é feita em lote ao final de toda a
            # emissão (ver corrigir_pasta_spans), não arquivo a arquivo aqui,
            # pois logo após salvar o PDF pode ainda estar com lock do app de origem.

            # Limpar janelas para próxima iteração
            self.cleanup_windows()

            return True

        except Exception as e:
            self.log(f"❌ Erro ao salvar PDF: {str(e)}")
            return False

    def salvar_excel_empregados(self) -> bool:
        """Salva o relatório de Empregados como Excel. Nome/local previsíveis
        (para facilitar integração futura com Meg_One, fora de escopo aqui).
        Extensão .xlsx assumida — confirmar na prática se o Domínio já adiciona
        a extensão sozinho ou se precisa ser explícita no nome digitado."""
        try:
            if self.should_stop():
                return False

            self.log("💾 Configurando salvamento do Excel de Empregados")

            save_hwnd = self._find_save_excel_window_hwnd()
            if not save_hwnd:
                self.log("❌ Janela de salvamento do Excel não encontrada")
                return False

            if self.should_stop():
                return False
            self.check_pause()

            data_inicio, _ = self.calcular_periodo_mes_referencia()
            mes_ano = datetime.strptime(data_inicio, "%d/%m/%Y").strftime("%Y-%m-%d")
            nome_arquivo = f"Empregados_Admissao_{mes_ano}.xls"

            diretorio = self.gui.obter_pasta_dia() if self.gui.diretorio_salvamento.get() else ""
            if diretorio:
                caminho_completo = os.path.join(diretorio, "Excels Base", nome_arquivo)
            else:
                caminho_completo = nome_arquivo
            self.log(f"📝 Salvando relatório de Empregados em: {caminho_completo}")

            time.sleep(0.5)
            self._force_focus(save_hwnd)
            filename_edit = win32gui.FindWindowEx(save_hwnd, 0, "Edit", None)
            if filename_edit:
                win32gui.SendMessage(filename_edit, win32con.WM_SETTEXT, 0, caminho_completo)
                time.sleep(0.2)
            else:
                send_keys('^a')
                send_keys(caminho_completo, with_spaces=True)
            time.sleep(0.3)

            if self.should_stop():
                return False
            self.check_pause()

            self.log("💾 Salvando Excel de Empregados")
            send_keys('{ENTER}')

            # O diálogo "já existe / substituir" é uma janela SEPARADA que só
            # aparece depois que 'Selecione o arquivo' fecha — ver _criar_checador_salvamento.
            if not self.wait_for_condition(
                self._criar_checador_salvamento(save_hwnd),
                timeout=60,
                poll_interval=0.3,
                description="Aguardando salvamento do Excel de Empregados"
            ):
                self.log("⚠️ Timeout aguardando janela do Excel fechar, continuando mesmo assim")

            self.log(f"✅ Excel de Empregados salvo: {nome_arquivo}")
            time.sleep(3.0)

            self.cleanup_windows()
            return True

        except Exception as e:
            self.log(f"❌ Erro ao salvar Excel de Empregados: {str(e)}")
            return False

    def salvar_excel_contrato_prazo_determinado(self) -> bool:
        """Salva o relatório de Contrato por Prazo Determinado como Excel.
        Mesma lógica de salvar_excel_empregados, mudando apenas o nome do arquivo."""
        try:
            if self.should_stop():
                return False

            self.log("💾 Configurando salvamento do Excel de Contrato por Prazo Determinado")

            save_hwnd = self._find_save_excel_window_hwnd()
            if not save_hwnd:
                self.log("❌ Janela de salvamento do Excel não encontrada")
                return False

            if self.should_stop():
                return False
            self.check_pause()

            data_inicio, _ = self.calcular_periodo_mes_referencia()
            mes_ano = datetime.strptime(data_inicio, "%d/%m/%Y").strftime("%Y-%m-%d")
            nome_arquivo = f"ContratoPrazoDeterminado_{mes_ano}.xls"

            diretorio = self.gui.obter_pasta_dia() if self.gui.diretorio_salvamento.get() else ""
            if diretorio:
                caminho_completo = os.path.join(diretorio, "Excels Base", nome_arquivo)
            else:
                caminho_completo = nome_arquivo
            self.log(f"📝 Salvando relatório de Contrato por Prazo Determinado em: {caminho_completo}")

            time.sleep(0.5)
            self._force_focus(save_hwnd)
            filename_edit = win32gui.FindWindowEx(save_hwnd, 0, "Edit", None)
            if filename_edit:
                win32gui.SendMessage(filename_edit, win32con.WM_SETTEXT, 0, caminho_completo)
                time.sleep(0.2)
            else:
                send_keys('^a')
                send_keys(caminho_completo, with_spaces=True)
            time.sleep(0.3)

            if self.should_stop():
                return False
            self.check_pause()

            self.log("💾 Salvando Excel de Contrato por Prazo Determinado")
            send_keys('{ENTER}')

            # O diálogo "já existe / substituir" é uma janela SEPARADA que só
            # aparece depois que 'Selecione o arquivo' fecha — por isso não basta
            # checar save_hwnd, senão a espera termina em 0s com o diálogo aberto.
            if not self.wait_for_condition(
                self._criar_checador_salvamento(save_hwnd),
                timeout=60,
                poll_interval=0.3,
                description="Aguardando salvamento do Excel de Contrato por Prazo Determinado"
            ):
                self.log("⚠️ Timeout aguardando janela do Excel fechar, continuando mesmo assim")

            self.log(f"✅ Excel de Contrato por Prazo Determinado salvo: {nome_arquivo}")
            time.sleep(3.0)

            self.cleanup_windows()
            return True

        except Exception as e:
            self.log(f"❌ Erro ao salvar Excel de Contrato por Prazo Determinado: {str(e)}")
            return False

    def _criar_checador_salvamento(self, save_hwnd: int, estabilidade: float = 4.0):
        """Cria a condição de fim de salvamento usada pelas esperas de PDF e Excel.

        O diálogo de 'já existe / substituir' é uma janela SEPARADA que só surge
        DEPOIS que a janela de salvamento fecha. Checar apenas save_hwnd faz a
        espera terminar instantaneamente ("concluído em 0.0s") com o diálogo
        ainda aberto na tela e o bot declarando sucesso indevidamente.

        Por isso, além de tratar o diálogo, exige que a tela fique limpa por
        `estabilidade` segundos antes de dar o salvamento por concluído — tempo
        para o diálogo atrasado aparecer e ser confirmado."""
        estado = {"limpo_desde": None}

        def checar() -> bool:
            # Confirma a substituição, se houver diálogo aguardando
            if self.handle_sobrescrever_dialog():
                estado["limpo_desde"] = None  # reinicia a contagem de estabilidade
                return False

            if win32gui.IsWindow(save_hwnd) and win32gui.IsWindowVisible(save_hwnd):
                estado["limpo_desde"] = None
                return False

            # Janela fechada e nenhum diálogo: aguarda o período de estabilidade
            agora = time.time()
            if estado["limpo_desde"] is None:
                estado["limpo_desde"] = agora
                return False

            return (agora - estado["limpo_desde"]) >= estabilidade

        return checar

    def handle_sobrescrever_dialog(self) -> bool:
        """Confirma o diálogo de 'arquivo já existe — deseja substituí-lo?' com Sim.

        Aparece sempre que uma emissão é reprocessada para a mesma data de
        referência (ex: rodou sábado, entraram novos funcionários e roda de novo
        para o mesmo sábado). Substituir é sempre o comportamento desejado, pois
        o relatório novo é o mais completo.

        Retorna True se tratou algum diálogo, False se não havia nenhum."""
        tratou = False
        try:
            # Pode haver mais de um em sequência (Excel + PDF), então repete
            # até não sobrar nenhum.
            for _ in range(5):
                found_hwnd = None
                found_text = ""

                def enum_callback(hwnd, _ctx):
                    nonlocal found_hwnd, found_text
                    if not win32gui.IsWindowVisible(hwnd):
                        return True
                    try:
                        if win32gui.GetClassName(hwnd) != "#32770":
                            return True
                        # Junta o texto de todos os Static filhos para achar a mensagem
                        textos = []

                        def child_cb(child_hwnd, _c):
                            try:
                                if win32gui.GetClassName(child_hwnd) == "Static":
                                    textos.append(win32gui.GetWindowText(child_hwnd))
                            except Exception:
                                pass
                            return True

                        win32gui.EnumChildWindows(hwnd, child_cb, None)
                        conteudo = " ".join(textos)
                        # Normaliza acentos: o texto pode vir com acentuação
                        # diferente conforme a codepage do sistema.
                        conteudo = unicodedata.normalize("NFKD", conteudo)
                        conteudo = "".join(c for c in conteudo if not unicodedata.combining(c)).lower()
                        if "ja existe" in conteudo and "substitu" in conteudo:
                            found_hwnd = hwnd
                            found_text = " ".join(textos)
                            return False
                    except Exception:
                        pass
                    return True

                win32gui.EnumWindows(enum_callback, None)

                if not found_hwnd:
                    break

                self.log(f"♻️ Arquivo já existe — substituindo: {found_text[:120]}")
                self._force_focus(found_hwnd)
                time.sleep(0.2)

                # Clica no botão "Sim" explicitamente (o diálogo tem Sim/Não,
                # não OK — depender do default do ENTER seria frágil).
                clicou = False
                try:
                    dlg_app = Application(backend="uia").connect(handle=found_hwnd)
                    dlg = dlg_app.window(handle=found_hwnd)
                    for titulo in ("&Sim", "Sim", "&Yes", "Yes"):
                        try:
                            btn = dlg.child_window(title=titulo, class_name="Button")
                            if btn.exists():
                                btn.click_input()
                                clicou = True
                                break
                        except Exception:
                            continue
                except Exception:
                    pass

                if not clicou:
                    # Fallback: 'S' é o acelerador de "Sim"; ENTER confirma o default
                    send_keys('%s')
                    time.sleep(0.2)
                    if win32gui.IsWindow(found_hwnd) and win32gui.IsWindowVisible(found_hwnd):
                        send_keys('{ENTER}')

                tratou = True
                time.sleep(0.6)

            return tratou

        except Exception as e:
            self.log(f"⚠️ Exceção ao tratar diálogo de substituição: {str(e)}")
            return tratou

    def handle_error_dialogs(self) -> bool:
        """Trata diálogos de erro que podem aparecer.
        Retorna True se deve continuar, False se deve abortar."""
        try:
            # O diálogo de "já existe / substituir" não é erro — trata e segue.
            if self.handle_sobrescrever_dialog():
                return True

            error_titles_lower = {"erro", "erro léxico", "aviso", "atenção",
                                  "informação", "alerta", "warning", "error"}

            found_hwnd = None
            found_title = None

            def enum_callback(hwnd, _):
                nonlocal found_hwnd, found_title
                if not win32gui.IsWindowVisible(hwnd):
                    return True
                try:
                    title = win32gui.GetWindowText(hwnd)
                    if not title:
                        return True
                    title_lower = title.strip().lower()
                    for err_title in error_titles_lower:
                        if title_lower == err_title or err_title in title_lower:
                            if win32gui.GetClassName(hwnd) == "#32770":
                                found_hwnd = hwnd
                                found_title = title
                                return False
                except Exception:
                    pass
                return True

            win32gui.EnumWindows(enum_callback, None)

            if found_hwnd is None:
                return True

            try:
                error_window = self.app.window(handle=found_hwnd)
            except Exception:
                win32gui.SetForegroundWindow(found_hwnd)
                send_keys('{ENTER}')
                time.sleep(0.3)
                return True

            message = ""
            try:
                message = error_window.window_text()
                try:
                    static_texts = error_window.children(class_name="Static")
                    for static in static_texts:
                        text = static.window_text()
                        if text:
                            message += " " + text
                except Exception:
                    pass
            except Exception:
                pass

            self.log(f"⚠️ Diálogo detectado: '{found_title}' - {message[:100] if message else 'sem mensagem'}")

            # Mensagens de aviso que, ao fechar, abrem a janela de salvamento (continua o fluxo)
            mensagens_continuar_salvamento = [
                "erro na gravação do relatório",
                "nome do caminho inválido",
                "caracteres não permitidos"
            ]

            message_lower = message.lower()
            for msg in mensagens_continuar_salvamento:
                if msg in message_lower:
                    self.log(f"⚠️ Aviso de gravação detectado (não crítico): {msg}")
                    error_window.set_focus()
                    send_keys('{ENTER}')
                    time.sleep(0.5)
                    return True  # Continua o fluxo — janela de salvamento vai abrir após fechar este aviso

            # Mensagens que indicam que não há dados — aborta a linha atual
            mensagens_abortar = [
                "sem dados para emitir",
                "nenhum registro encontrado",
                "não há dados",
                "registro não encontrado"
            ]

            for msg in mensagens_abortar:
                if msg in message_lower:
                    self.log(f"⚠️ Aviso não crítico: {msg}")
                    error_window.set_focus()
                    send_keys('{ENTER}')
                    time.sleep(0.5)
                    for _ in range(4):
                        send_keys('{ESC}')
                        time.sleep(0.5)
                    return False

            if "léxico" in found_title.lower():
                self.log("⚠️ Erro léxico detectado, fechando...")
                error_window.set_focus()
                for _ in range(3):
                    send_keys('{ESC}')
                    time.sleep(0.5)
                return True

            self.log(f"⚠️ Fechando diálogo '{found_title}'...")
            error_window.set_focus()
            time.sleep(0.2)

            try:
                ok_button = error_window.child_window(title="OK", class_name="Button")
                if ok_button.exists():
                    ok_button.click_input()
                    time.sleep(0.5)
                    if found_title.lower() in ("erro", "aviso"):
                        return False
                    return True
            except Exception:
                pass

            send_keys('{ENTER}')
            time.sleep(0.5)

            try:
                if error_window.exists():
                    send_keys('{ESC}')
                    time.sleep(0.3)
            except Exception:
                pass

            if found_title.lower() in ("erro", "aviso"):
                return False

            return True

        except Exception as e:
            self.log(f"⚠️ Exceção ao verificar diálogos: {str(e)}")
            return True


    def cleanup_windows(self):
        """Limpa e fecha janelas abertas"""
        try:
            self.log("🧹 Limpando janelas")

            try:
                main_hwnd = self.main_window.handle
                self._force_focus(main_hwnd)
            except Exception:
                pass

            for _ in range(4):
                send_keys('{ESC}')
                time.sleep(0.5)

            # Verificar se a janela Relatórios Admissionais ainda está aberta
            try:
                admiss_window = self.main_window.child_window(
                    title="Relatórios Admissionais",
                    class_name="FNWND3190"
                )

                if admiss_window.exists() and admiss_window.is_visible():
                    self.log("🔄 Fechando Relatórios Admissionais restante")
                    send_keys('{ESC}')
                    time.sleep(0.5)
            except Exception:
                pass

            time.sleep(1)

        except Exception as e:
            self.log(f"⚠️ Erro durante limpeza: {str(e)}")

    def processar_relatorio_empregados(self) -> bool:
        """Gera o relatório de Empregados (filtro Admissão do mês de referência)
        em Excel. Roda uma única vez, antes do loop de emissão de contratos.

        ATENÇÃO: várias etapas abaixo são hipóteses (marcadas # INCERTEZA) que
        precisam ser validadas rodando o robô com o Domínio Folha aberto e
        ajustadas conforme o comportamento real observado nos logs."""
        try:
            if self.should_stop():
                return False
            self.check_pause()

            # Fecha qualquer janela "Empregados" residual (ex: de uma tentativa
            # anterior que falhou) para evitar ambiguidade ao localizar a nova.
            for _ in range(5):
                try:
                    residual = self.main_window.child_window(title="Empregados", class_name="FNWND3190", found_index=0)
                    if not residual.exists():
                        break
                    residual.close()
                    time.sleep(0.5)
                except Exception:
                    break

            # --- 1) Abrir janela "Empregados" ---
            # Navegação confirmada pelo usuário: Alt+R (Relatórios) -> 'c' -> 'e' (Empregados).
            self.log("📇 Acessando menu para abrir Empregados")
            self.main_window.set_focus()
            send_keys('%r')  # Alt+R
            if not self.smart_sleep(0.5):
                return False
            send_keys('c')
            if not self.smart_sleep(0.5):
                return False
            send_keys('e')  # Empregados
            if not self.smart_sleep(0.5):
                return False
            send_keys('{ENTER}')
            if not self.smart_sleep(1):
                return False

            max_attempts = 10
            empregados_window = None
            for attempt in range(max_attempts):
                if self.should_stop():
                    return False
                self.check_pause()
                try:
                    # Pode haver mais de uma janela "Empregados" simultânea (ex: uma
                    # residual de execução anterior) — pega a mais recente (topo, index 0).
                    candidata = self.main_window.child_window(
                        title="Empregados", class_name="FNWND3190", found_index=0
                    )
                    if candidata.exists():
                        empregados_window = candidata
                        break
                    if not self.handle_error_dialogs():
                        self.cleanup_windows()
                        return False
                    if not self.smart_sleep(1):
                        return False
                except Exception:
                    if attempt == max_attempts - 1:
                        self.log("❌ Janela Empregados não encontrada (timeout)")
                        return False

            if not empregados_window:
                self.log("❌ Janela Empregados não encontrada")
                return False
            self.log("📇 Janela Empregados localizada")

            if self.should_stop():
                return False
            self.check_pause()

            # --- 2) Botão "Seleção..." ---
            send_keys('%s')  # Alt+S
            if not self.smart_sleep(0.8):
                return False

            # Janela top-level confirmada: título "Seleção de Empregados", classe FNWND3190.
            selecao_window = self.main_window.child_window(
                title="Seleção de Empregados", class_name="FNWND3190"
            )
            if not selecao_window.exists():
                self.log("❌ Janela 'Seleção de Empregados' não encontrada")
                return False

            # --- 3) Aba "Adicional" ---
            try:
                tab_control = selecao_window.child_window(auto_id="1002", class_name="PBTabControl32_100")
                tab_control.click_input(coords=(70, 10))  # coords aproximadas, ajustar se necessário
                if not self.smart_sleep(0.5):
                    return False
            except Exception as e:
                self.log(f"⚠️ Erro ao acessar aba Adicional: {str(e)}")
                return False

            # --- 4) Checkbox "Admissão:" ---
            try:
                cb_admissao = selecao_window.child_window(auto_id="1008", class_name="Button")
                if not cb_admissao.get_toggle_state():
                    cb_admissao.click_input()
                    if not self.smart_sleep(0.3):
                        return False
                self.log("✅ Checkbox Admissão marcado")
            except Exception as e:
                self.log(f"❌ Erro ao marcar checkbox Admissão: {str(e)}")
                return False

            # --- 5/6) Datas Início/Fim ---
            data_inicio, data_fim = self.calcular_periodo_mes_referencia()
            try:
                campo_inicio = selecao_window.child_window(auto_id="1009", class_name="PBEDIT190")
                campo_inicio.click_input()
                send_keys('^a')
                send_keys(data_inicio)
                if not self.smart_sleep(0.3):
                    return False

                campo_fim = selecao_window.child_window(auto_id="1010", class_name="PBEDIT190")
                campo_fim.click_input()
                send_keys('^a')
                send_keys(data_fim)
                if not self.smart_sleep(0.3):
                    return False
            except Exception as e:
                self.log(f"❌ Erro ao preencher datas de admissão: {str(e)}")
                return False

            # --- 7) OK da janela de Seleção ---
            send_keys('%o')  # Alt+O
            if not self.smart_sleep(1):
                return False
            if not self.handle_error_dialogs():
                self.cleanup_windows()
                return False

            # --- 8) Botão "Empresas..." ---
            send_keys('%e')  # Alt+E
            if not self.smart_sleep(0.8):
                return False

            # --- 9/10/11) Janela modal "Selecionar Empresas" ---
            try:
                empresas_window = self.main_window.child_window(class_name="FNWNS3190")
                if not empresas_window.exists():
                    self.log("❌ Janela Selecionar Empresas não encontrada")
                    return False

                btn_todas = empresas_window.child_window(auto_id="1006", class_name="Button")
                btn_todas.click_input()
                if not self.smart_sleep(0.3):
                    return False

                btn_ok_empresas = empresas_window.child_window(auto_id="1009", class_name="Button")
                btn_ok_empresas.click_input()
                if not self.smart_sleep(0.8):
                    return False
            except Exception as e:
                self.log(f"❌ Erro ao selecionar empresas: {str(e)}")
                return False

            if not self.handle_error_dialogs():
                self.cleanup_windows()
                return False

            # --- 12) "ok >> ok" final ---
            # INCERTEZA: pode ser redundante com o OK acima de "Selecionar Empresas" —
            # validar na prática se este Alt+O extra é necessário ou incorreto.
            send_keys('%o')  # HIPÓTESE — validar na prática
            if not self.smart_sleep(1):
                return False

            # --- 13) Clicar no ícone de exportar Excel (painel FNUDO3190, auto_id
            # 1020, na barra lateral da janela Empregados) e aguardar o diálogo
            # "Selecione o arquivo". Ctrl+D NÃO se aplica aqui — abre "Salvar em PDF".
            if not self.smart_sleep(2):
                return False
            self.log("📊 Clicando no ícone de exportar Excel")
            try:
                btn_excel = empregados_window.child_window(
                    auto_id="1020", class_name="FNUDO3190", found_index=0
                )
                btn_excel.click_input()
            except Exception as e:
                self.log(f"❌ Erro ao clicar no ícone de exportar Excel: {str(e)}")
                return False

            self.log("📊 Aguardando janela de salvamento...")
            timeout_total = 180
            inicio = time.time()
            janela_encontrada = False

            while time.time() - inicio < timeout_total:
                if self.should_stop():
                    return False

                # Diálogo de "já existe / substituir" não é erro — confirma e segue
                self.handle_sobrescrever_dialog()

                if self._any_error_dialog_visible():
                    deve_continuar = self.handle_error_dialogs()
                    if deve_continuar:
                        time.sleep(0.5)
                        continue
                    else:
                        self.log("⚠️ Diálogo de erro detectado ao gerar relatório de Empregados")
                        self.cleanup_windows()
                        return False

                if self._window_exists("Selecione o arquivo", "#32770") or self._save_excel_dialog_exists():
                    janela_encontrada = True
                    break

                time.sleep(0.25)

            if not janela_encontrada:
                self.log("❌ Timeout aguardando janela de salvamento do Excel")
                return False

            if not self.handle_error_dialogs():
                self.cleanup_windows()
                return False

            return self.salvar_excel_empregados()

        except Exception as e:
            self.log(f"❌ Erro ao gerar relatório de Empregados: {str(e)}")
            self.log(f"Traceback: {traceback.format_exc()}")
            return False

    def processar_contrato_prazo_determinado(self) -> bool:
        """Gera o relatório de Contrato por Prazo Determinado (filtro do mês de
        referência, contrato de experiência) em Excel. Roda uma única vez, logo
        após processar_relatorio_empregados e antes do loop de emissão de contratos."""
        try:
            if self.should_stop():
                return False
            self.check_pause()

            # Fecha qualquer janela "Contrato por Prazo Determinado" residual
            # (ex: de uma tentativa anterior) para evitar ambiguidade ao localizar.
            for _ in range(5):
                try:
                    residual = self.main_window.child_window(
                        title="Contrato por Prazo Determinado", class_name="FNWND3190", found_index=0
                    )
                    if not residual.exists():
                        break
                    residual.close()
                    time.sleep(0.5)
                except Exception:
                    break

            # --- 1) Abrir janela "Contrato por Prazo Determinado" via menu ---
            # Navegação confirmada pelo usuário: Alt+R (Relatórios) -> 'o' (Outros)
            # -> 'c' (Contrato) -> 'c' (Contrato de Prazo Determinado) -> Enter.
            self.log("📃 Acessando menu Relatórios > Outros > Contrato > Contrato de Prazo Determinado")
            self.main_window.set_focus()
            send_keys('%r')  # Alt+R
            if not self.smart_sleep(0.5):
                return False
            send_keys('o')  # Outros
            if not self.smart_sleep(0.5):
                return False
            send_keys('c')  # Contrato
            if not self.smart_sleep(0.5):
                return False
            send_keys('c')  # Contrato de Prazo Determinado
            if not self.smart_sleep(0.5):
                return False
            send_keys('{ENTER}')
            if not self.smart_sleep(1):
                return False

            max_attempts = 10
            contrato_window = None
            for attempt in range(max_attempts):
                if self.should_stop():
                    return False
                self.check_pause()
                try:
                    candidata = self.main_window.child_window(
                        title="Contrato por Prazo Determinado", class_name="FNWND3190", found_index=0
                    )
                    if candidata.exists():
                        contrato_window = candidata
                        break
                    if not self.handle_error_dialogs():
                        self.cleanup_windows()
                        return False
                    if not self.smart_sleep(1):
                        return False
                except Exception:
                    if attempt == max_attempts - 1:
                        self.log("❌ Janela Contrato por Prazo Determinado não encontrada (timeout)")
                        return False

            if not contrato_window:
                self.log("❌ Janela Contrato por Prazo Determinado não encontrada")
                return False
            self.log("📃 Janela Contrato por Prazo Determinado localizada")

            if self.should_stop():
                return False
            self.check_pause()

            # --- 2/3) Datas Início/Fim — campo Início já vem focado ao abrir a
            # janela, então digita direto e usa Tab para ir ao campo Fim
            # (AutomationId 1007), em vez de localizar por child_window. ---
            data_inicio, data_fim = self.calcular_periodo_mes_referencia()
            try:
                send_keys('^a')
                send_keys(data_inicio)
                if not self.smart_sleep(0.3):
                    return False
                send_keys('{TAB}')
                if not self.smart_sleep(0.3):
                    return False
                send_keys('^a')
                send_keys(data_fim)
                if not self.smart_sleep(0.3):
                    return False
            except Exception as e:
                self.log(f"❌ Erro ao preencher datas do contrato: {str(e)}")
                return False

            # --- 4) Radio "Contrato de experiência" (auto_id 6) ---
            # RadioButton usa SelectionItemPattern (is_selected), não TogglePattern
            # (get_toggle_state é de checkbox e falha aqui).
            try:
                radio_experiencia = contrato_window.child_window(auto_id="6", class_name="Button")
                if not radio_experiencia.is_selected():
                    radio_experiencia.click_input()
                    if not self.smart_sleep(0.3):
                        return False
                self.log("✅ Contrato de experiência selecionado")
            except Exception as e:
                self.log(f"❌ Erro ao selecionar Contrato de experiência: {str(e)}")
                return False

            if self.should_stop():
                return False
            self.check_pause()

            # --- 5) OK ---
            send_keys('%o')  # Alt+O
            if not self.smart_sleep(1):
                return False
            if not self.handle_error_dialogs():
                self.cleanup_windows()
                return False

            # --- 6) Clicar no ícone de exportar Excel (painel FNUDO3190, auto_id
            # 1020, mesma posição confirmada na janela Empregados) e aguardar o
            # diálogo "Selecione o arquivo" ---
            if not self.smart_sleep(2):
                return False
            self.log("📊 Clicando no ícone de exportar Excel")
            try:
                btn_excel = contrato_window.child_window(
                    auto_id="1020", class_name="FNUDO3190", found_index=0
                )
                btn_excel.click_input()
            except Exception as e:
                self.log(f"❌ Erro ao clicar no ícone de exportar Excel: {str(e)}")
                return False

            self.log("📊 Aguardando janela de salvamento...")
            timeout_total = 180
            inicio = time.time()
            janela_encontrada = False

            while time.time() - inicio < timeout_total:
                if self.should_stop():
                    return False

                # Diálogo de "já existe / substituir" não é erro — confirma e segue
                self.handle_sobrescrever_dialog()

                if self._any_error_dialog_visible():
                    deve_continuar = self.handle_error_dialogs()
                    if deve_continuar:
                        time.sleep(0.5)
                        continue
                    else:
                        self.log("⚠️ Diálogo de erro detectado ao gerar relatório de Contrato por Prazo Determinado")
                        self.cleanup_windows()
                        return False

                if self._window_exists("Selecione o arquivo", "#32770") or self._save_excel_dialog_exists():
                    janela_encontrada = True
                    break

                time.sleep(0.25)

            if not janela_encontrada:
                self.log("❌ Timeout aguardando janela de salvamento do Excel")
                return False

            if not self.handle_error_dialogs():
                self.cleanup_windows()
                return False

            return self.salvar_excel_contrato_prazo_determinado()

        except Exception as e:
            self.log(f"❌ Erro ao gerar relatório de Contrato por Prazo Determinado: {str(e)}")
            self.log(f"Traceback: {traceback.format_exc()}")
            return False

def main():
    """Função principal"""
    try:
        gui = AutomacaoGUI()
        gui.executar()
    except Exception as e:
        print(f"Erro crítico na aplicação: {str(e)}")
        traceback.print_exc()

if __name__ == "__main__":
    main()
