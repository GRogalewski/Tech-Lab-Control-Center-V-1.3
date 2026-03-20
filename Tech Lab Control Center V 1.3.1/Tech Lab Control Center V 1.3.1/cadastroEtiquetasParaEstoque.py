# biblioteca para datas
import customtkinter as ctk
import win32print
import win32ui
import sqlite3
import datetime
import csv
import os
import requests
import threading
import json
from tkinter import messagebox, scrolledtext, ttk
from tkinter.filedialog import asksaveasfilename

# Configurar aparência moderna
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("dark-blue")

class EtiquetaSimplificadaApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        self.title("Tech Lab Control Center V 1.4")
        self.geometry("1100x750") 
        self.minsize(900, 650)    
        
        # Configurações de impressão Zebra
        self.config_impressao = {
            'temperatura': 10,
            'velocidade': 5,
            'tonalidade': 5,
            'largura_etiqueta': 550,   
            'altura_etiqueta': 35,     
            'margem_esquerda': 2,      
            'margem_superior': 10
        }
        
        # Chaves Fixas do ServiceNow (TechCenter Lab)
        self.SYS_ID_FILA_LAB = "600096386fd2e6006bda17164b3ee459"
        self.SYS_ID_GUSTAVO = "dc0df06797c32e1080d1b537f053afc5"
        self.ARQUIVO_COOKIE = "cookie.txt"
        
        # Variável para controlar se o patrimônio foi encontrado
        self.patrimonio_encontrado = False
        
        # Conectar aos bancos de dados
        self.conectar_bancos()
        
        # Detectar impressoras Zebra automaticamente
        self.detectar_impressoras_zebra()
        
        self.criar_widgets()
        self.carregar_impressoras()
        self.carregar_modelos()

        # Iniciar atualização automática das estatísticas
        self.auto_refresh_stats()
        
    def conectar_bancos(self):
        try:
            self.conn_patrimonios = sqlite3.connect('patrimoniosCadastrados.db')
            self.cursor_patrimonios = self.conn_patrimonios.cursor()
            
            self.cursor_patrimonios.execute("PRAGMA table_info(patrimonios)")
            colunas = self.cursor_patrimonios.fetchall()
            colunas_existentes = [coluna[1] for coluna in colunas]
            
            if not colunas_existentes or 'modelo' not in colunas_existentes:
                if colunas_existentes:
                    self.cursor_patrimonios.execute("ALTER TABLE patrimonios RENAME TO patrimonios_old")
                    self.conn_patrimonios.commit()
                
                self.cursor_patrimonios.execute('''
                    CREATE TABLE IF NOT EXISTS patrimonios (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        patrimonio TEXT UNIQUE NOT NULL,
                        memoria TEXT NOT NULL,
                        modelo TEXT NOT NULL,
                        tipo TEXT NOT NULL,
                        estado TEXT NOT NULL,
                        pxe TEXT DEFAULT 'Não configurado',
                        autopilot TEXT DEFAULT 'Não configurado',
                        data_cadastro TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                        data_impressao TIMESTAMP
                    )
                ''')
                
                if colunas_existentes:
                    try:
                        self.cursor_patrimonios.execute('''
                            INSERT INTO patrimonios (patrimonio, memoria, modelo, tipo, estado, pxe, autopilot, data_cadastro, data_impressao)
                            SELECT 
                                patrimonio, 
                                memoria, 
                                COALESCE(modelo, 'T14') as modelo,
                                COALESCE(tipo, 'Não avaliado') as tipo,
                                COALESCE(estado, 'Excelente') as estado,
                                'Não configurado' as pxe,
                                'Não configurado' as autopilot,
                                data_cadastro, 
                                data_impressao
                            FROM patrimonios_old
                        ''')
                        self.cursor_patrimonios.execute("DROP TABLE patrimonios_old")
                    except sqlite3.Error as e:
                        print(f"Erro na migração: {e}")
                        self.cursor_patrimonios.execute("DROP TABLE IF EXISTS patrimonios_old")
                
                self.conn_patrimonios.commit()
            
            self.conn_modelos = sqlite3.connect('modelosDeMaquinas.db')
            self.cursor_modelos = self.conn_modelos.cursor()
            
            self.cursor_modelos.execute('''
                CREATE TABLE IF NOT EXISTS tipos_equipamento (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    modelo TEXT UNIQUE NOT NULL,
                    data_criacao TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
            ''')
            
            self.cursor_modelos.execute("SELECT COUNT(*) FROM tipos_equipamento")
            if self.cursor_modelos.fetchone()[0] == 0:
                modelos_padrao = ["T14", "Latitude 5440", "Latitude 5430", "Latitude 5420", "E14"]
                for modelo in modelos_padrao:
                    self.cursor_modelos.execute("INSERT OR IGNORE INTO tipos_equipamento (modelo) VALUES (?)", (modelo,))
                self.conn_modelos.commit()
            
        except sqlite3.Error as e:
            messagebox.showerror("Erro de Banco de Dados", f"Erro ao conectar com os bancos: {e}")
    
    def carregar_modelos(self):
        try:
            self.cursor_modelos.execute("SELECT modelo FROM tipos_equipamento ORDER BY modelo")
            modelos = self.cursor_modelos.fetchall()
            self.modelos_lista = [modelo[0] for modelo in modelos]
            if hasattr(self, 'modelo_dropdown'):
                self.modelo_dropdown.configure(values=self.modelos_lista)
                if self.modelos_lista:
                    self.modelo_dropdown.set(self.modelos_lista[0])
        except sqlite3.Error as e:
            modelos_padrao = ["T14", "Latitude 5440", "Latitude 5430", "Latitude 5420", "E14"]
            self.modelos_lista = modelos_padrao
            if hasattr(self, 'modelo_dropdown'):
                self.modelo_dropdown.configure(values=modelos_padrao)
                if modelos_padrao:
                    self.modelo_dropdown.set(modelos_padrao[0])
    
    def criar_widgets(self):
        self.tabview = ctk.CTkTabview(self, corner_radius=15)
        self.tabview.pack(fill="both", expand=True, padx=20, pady=20)
        
        self.tab_individual = self.tabview.add("Etiqueta Individual")
        self.tab_lote = self.tabview.add("Impressão em Lote")
        self.tab_estatisticas = self.tabview.add("Estatísticas")
        self.tab_modelos = self.tabview.add("Gerenciar Modelos")
        self.tab_config = self.tabview.add("Config. Impressão")
        
        # 🚀 ABAS DO SERVICENOW
        self.tab_gerenciar = self.tabview.add("Gerenciar Chamados")
        self.tab_snow = self.tabview.add("Fechar Chamados")
        
        self.criar_aba_individual()
        self.criar_aba_lote()
        self.criar_aba_estatisticas()
        self.criar_aba_gerenciar_modelos()
        self.criar_aba_configuracoes()
        self.criar_aba_snow() 
        self.criar_aba_gerenciar() 
        
    # =========================================================================
    # 🚀 MOTOR DO TECH LAB CONTROL CENTER (SERVICENOW)
    # =========================================================================
    
    def carregar_cookie_salvo(self):
        if os.path.exists(self.ARQUIVO_COOKIE):
            with open(self.ARQUIVO_COOKIE, 'r', encoding='utf-8') as f:
                return f.read().strip()
        return ""

    def salvar_cookie_local(self, cookie):
        with open(self.ARQUIVO_COOKIE, 'w', encoding='utf-8') as f:
            f.write(cookie)

    # -------------------------------------------------------------------------
    # ABA: GERENCIAR CHAMADOS (TRIAGEM EM LOTE)
    # -------------------------------------------------------------------------
    # -------------------------------------------------------------------------
    # ABA: GERENCIAR CHAMADOS (TRIAGEM EM LOTE COM CHECKLIST)
    # -------------------------------------------------------------------------
    # -------------------------------------------------------------------------
    # ABA: GERENCIAR CHAMADOS (TRIAGEM EM LOTE)
    # -------------------------------------------------------------------------
    def criar_aba_gerenciar(self):
        main_frame = ctk.CTkFrame(self.tab_gerenciar, corner_radius=20)
        main_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        header_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        header_frame.pack(fill="x", padx=20, pady=(20, 10))
        
        title_label = ctk.CTkLabel(header_frame, text="⚙️ Triagem em Lote de Chamados (Tech Lab)", font=ctk.CTkFont(size=24, weight="bold"))
        title_label.pack(side="left")
        
        content_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        content_frame.pack(fill="both", expand=True, padx=20, pady=10)
        
        left_panel = ctk.CTkFrame(content_frame, corner_radius=15)
        left_panel.pack(side="left", fill="both", expand=True, padx=(0, 10))
        
        right_panel = ctk.CTkFrame(content_frame, corner_radius=15)
        right_panel.pack(side="right", fill="both", expand=True, padx=(10, 0))
        
        ctk.CTkLabel(left_panel, text="1. Escolha a Etapa e Bipe os Equipamentos", font=ctk.CTkFont(weight="bold")).pack(anchor="w", padx=15, pady=(15, 5))
        ctk.CTkLabel(left_panel, text="As tags listadas serão assumidas, movidas para Work in Progress e a descrição será atualizada com a etapa escolhida.", text_color="#AAAAAA", wraplength=400, justify="left").pack(anchor="w", padx=15, pady=(0, 10))
        
        estagio_frame = ctk.CTkFrame(left_panel, fg_color="transparent")
        estagio_frame.pack(fill="x", padx=15, pady=(0, 10))
        
        ctk.CTkLabel(estagio_frame, text="Etapa atual:", font=ctk.CTkFont(weight="bold")).pack(side="left", padx=(0, 10))
        self.estagio_triagem_var = ctk.StringVar(value="Formatação")
        
        # Caixinha limpa, com Manutenção inclusa, mas sem puxar interface de checklist
        self.estagio_combo = ctk.CTkComboBox(estagio_frame, values=["Formatação", "Validação", "Limpeza", "Manutenção", "Acionado garantia"], variable=self.estagio_triagem_var, width=180)
        self.estagio_combo.pack(side="left")

        self.tags_triagem_text = scrolledtext.ScrolledText(left_panel, height=13, font=("Consolas", 12), bg="#2B2B2B", fg="white", insertbackground="white", relief="flat", bd=0)
        self.tags_triagem_text.pack(fill="both", expand=True, padx=15, pady=(15, 15))
        
        btn_frame_left = ctk.CTkFrame(left_panel, fg_color="transparent")
        btn_frame_left.pack(fill="x", padx=15, pady=(0, 15))
        
        self.btn_limpar_triagem = ctk.CTkButton(btn_frame_left, text="🗑️ Limpar Lista", command=lambda: self.tags_triagem_text.delete(1.0, "end"), fg_color="#E74C3C", hover_color="#C0392B", height=40)
        self.btn_limpar_triagem.pack(side="left", fill="x", expand=True, padx=(0, 5))
        
        self.btn_executar_triagem = ctk.CTkButton(btn_frame_left, text="🚀 Executar Triagem em Massa", command=self.iniciar_triagem, fg_color="#3498DB", hover_color="#2980B9", height=40)
        self.btn_executar_triagem.pack(side="right", fill="x", expand=True, padx=(5, 0))

        ctk.CTkLabel(right_panel, text="Log do Sistema", font=ctk.CTkFont(weight="bold")).pack(anchor="w", padx=15, pady=(15, 5))
        
        self.log_triagem = scrolledtext.ScrolledText(right_panel, font=("Consolas", 11), bg="#1E1E1E", fg="white", insertbackground="white", relief="flat", bd=0)
        self.log_triagem.pack(fill="both", expand=True, padx=15, pady=(0, 15))
        
        self.log_triagem.tag_config("red", foreground="#E74C3C")
        self.log_triagem.tag_config("green", foreground="#2AA876")
        self.log_triagem.tag_config("orange", foreground="#F39C12")
        self.log_triagem.tag_config("blue", foreground="#3498DB")
        self.log_triagem.tag_config("white", foreground="white")
        self.log_triagem.tag_config("purple", foreground="#9B59B6")
        
        self.log_triagem.insert("end", "Aguardando lote de Service Tags...\n", "white")
        
        self.triagem_progress = ctk.CTkProgressBar(right_panel)
        self.triagem_progress.pack(fill="x", padx=15, pady=(0, 15))
        self.triagem_progress.set(0)

    def _atualizar_log_triagem(self, mensagem, cor="white"):
        timestamp = f"[{datetime.datetime.now().strftime('%H:%M:%S')}] "
        self.log_triagem.insert("end", timestamp + mensagem, cor)
        self.log_triagem.see("end")

    def iniciar_triagem(self):
        tags_raw = self.tags_triagem_text.get(1.0, "end").strip()
        cookie = self.cookie_var.get().strip() 
        estagio_escolhido = self.estagio_triagem_var.get() 
        
        if not cookie:
            messagebox.showwarning("Atenção", "Por favor, cole o Cookie do ServiceNow na aba 'Fechar Chamados' primeiro.")
            return
            
        if not tags_raw:
            messagebox.showwarning("Atenção", "Bipe pelo menos uma Service Tag para processar.")
            return

        self.salvar_cookie_local(cookie)

        lista_tags = []
        for linha in tags_raw.split('\n'):
            linha = linha.strip()
            if linha:
                if ',' in linha:
                    lista_tags.extend([t.strip().upper() for t in linha.split(',') if t.strip()])
                else:
                    lista_tags.append(linha.upper())

        self.btn_executar_triagem.configure(state="disabled", text="⏳ Processando...")
        self.triagem_progress.set(0)
        self._atualizar_log_triagem(f"\n========================================\n", "blue")
        self._atualizar_log_triagem(f"🚀 INICIANDO {estagio_escolhido.upper()} ({len(lista_tags)} MÁQUINAS)\n", "purple")
        self._atualizar_log_triagem(f"========================================\n", "blue")
        
        threading.Thread(target=self._thread_executar_triagem, args=(lista_tags, cookie, estagio_escolhido), daemon=True).start()

    def _thread_executar_triagem(self, lista_tags, cookie, estagio_escolhido):
        headers = {
            "Accept": "application/json",
            "Content-Type": "application/json",
            "Cookie": cookie
        }
        url_base = "https://grupoboticario.service-now.com/api/now/table/sc_req_item"
        total = len(lista_tags)
        
        for idx, tag in enumerate(lista_tags):
            self.after(0, self._atualizar_log_triagem, f"\n🔍 Buscando {tag}...\n", "white")
            query_lupa = f"123TEXTQUERY321={tag}^active=true^assignment_group={self.SYS_ID_FILA_LAB}"
            
            try:
                resp_busca = requests.get(url_base, headers=headers, params={"sysparm_query": query_lupa, "sysparm_fields": "number,sys_id", "sysparm_limit": 1})
                
                if resp_busca.status_code == 401:
                    self.after(0, self._atualizar_log_triagem, f"❌ ERRO: Seu Cookie expirou ou é inválido.\n", "red")
                    break 
                    
                if resp_busca.status_code != 200 or not resp_busca.json().get('result'):
                    self.after(0, self._atualizar_log_triagem, f"⚠️ Ignorado: Nenhum chamado aberto para {tag}.\n", "orange")
                    self._atualizar_progresso_triagem(idx + 1, total)
                    continue
                    
                chamado = resp_busca.json()['result'][0]
                numero_ritm = chamado.get('number')
                sys_id_ritm = chamado.get('sys_id')
                
                self.after(0, self._atualizar_log_triagem, f"🎯 RITM {numero_ritm}: Aplicando '{estagio_escolhido}'...\n", "white")
                
                data_hoje = datetime.datetime.now().strftime("%d/%m/%Y")
                
                # Montagem limpa e direta do pacote
                if estagio_escolhido == "Acionado garantia":
                    novo_status = "300"
                    nova_descricao = f"{data_hoje} - Aguardando técnico"
                    dados_triagem = {
                        "state": novo_status, 
                        "assigned_to": self.SYS_ID_GUSTAVO,
                        "description": nova_descricao,
                        "comments": f"{data_hoje} - Acionado garantia", 
                        "correlation_id": "N/A"
                    }
                else:
                    novo_status = "2" 
                    nova_descricao = f"{data_hoje} - Em processo de {estagio_escolhido.lower()}"
                    dados_triagem = {
                        "state": novo_status, 
                        "assigned_to": self.SYS_ID_GUSTAVO,
                        "description": nova_descricao
                    }
                
                url_atualizar = f"{url_base}/{sys_id_ritm}"
                resp_atualizar = requests.patch(url_atualizar, headers=headers, data=json.dumps(dados_triagem))
                
                if resp_atualizar.status_code == 200:
                    self.after(0, self._atualizar_log_triagem, f"✅ SUCESSO: {numero_ritm} atualizado!\n", "green")
                else:
                    self.after(0, self._atualizar_log_triagem, f"❌ ERRO no {numero_ritm}: Code {resp_atualizar.status_code}\n", "red")
                    
            except Exception as e:
                self.after(0, self._atualizar_log_triagem, f"❌ Falha de Conexão na tag {tag}: {str(e)}\n", "red")
                
            self._atualizar_progresso_triagem(idx + 1, total)
            
        self.after(0, self._concluir_processo_triagem)

    def _atualizar_progresso_triagem(self, atual, total):
        progresso = atual / total
        self.after(0, lambda: self.triagem_progress.set(progresso))

    def _concluir_processo_triagem(self):
        self.btn_executar_triagem.configure(state="normal", text="🚀 Executar Triagem em Massa")
        self._atualizar_log_triagem(f"\n========================================\n", "blue")
        self._atualizar_log_triagem(f"🎉 LOTE DE TRIAGEM FINALIZADO!\n", "blue")
        self._atualizar_log_triagem(f"========================================\n", "blue")
        self.triagem_progress.set(1)

    # -------------------------------------------------------------------------
    # ABA: FECHAR CHAMADOS
    # -------------------------------------------------------------------------
    def criar_aba_snow(self):
        main_frame = ctk.CTkFrame(self.tab_snow, corner_radius=20)
        main_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        header_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        header_frame.pack(fill="x", padx=20, pady=(20, 10))
        
        title_label = ctk.CTkLabel(header_frame, text="Automação de Baixa (ServiceNow)", font=ctk.CTkFont(size=24, weight="bold"))
        title_label.pack(side="left")
        
        content_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        content_frame.pack(fill="both", expand=True, padx=20, pady=10)
        
        left_panel = ctk.CTkFrame(content_frame, corner_radius=15)
        left_panel.pack(side="left", fill="both", expand=True, padx=(0, 10))
        
        right_panel = ctk.CTkFrame(content_frame, corner_radius=15)
        right_panel.pack(side="right", fill="both", expand=True, padx=(10, 0))
        
        ctk.CTkLabel(left_panel, text="1. Autenticação (Cookie)", font=ctk.CTkFont(weight="bold")).pack(anchor="w", padx=15, pady=(15, 5))
        
        self.cookie_var = ctk.StringVar(value=self.carregar_cookie_salvo())
        self.cookie_entry = ctk.CTkEntry(left_panel, textvariable=self.cookie_var, show="*", placeholder_text="Cole o Cookie Atualizado do ServiceNow aqui...", height=35)
        self.cookie_entry.pack(fill="x", padx=15, pady=(0, 15))
        
        ctk.CTkLabel(left_panel, text="2. Bipar Equipamentos (Service Tags)", font=ctk.CTkFont(weight="bold")).pack(anchor="w", padx=15, pady=(5, 5))
        
        self.snow_tags_text = scrolledtext.ScrolledText(left_panel, height=15, font=("Consolas", 12), bg="#2B2B2B", fg="white", insertbackground="white", relief="flat", bd=0)
        self.snow_tags_text.pack(fill="both", expand=True, padx=15, pady=(0, 15))
        
        btn_frame_left = ctk.CTkFrame(left_panel, fg_color="transparent")
        btn_frame_left.pack(fill="x", padx=15, pady=(0, 15))
        
        self.btn_limpar_snow = ctk.CTkButton(btn_frame_left, text="🗑️ Limpar Lista", command=lambda: self.snow_tags_text.delete(1.0, "end"), fg_color="#E74C3C", hover_color="#C0392B", height=40)
        self.btn_limpar_snow.pack(side="left", fill="x", expand=True, padx=(0, 5))
        
        self.btn_executar_snow = ctk.CTkButton(btn_frame_left, text="🚀 Executar Baixas em Massa", command=self.iniciar_processo_baixas, fg_color="#2AA876", hover_color="#207A59", height=40)
        self.btn_executar_snow.pack(side="right", fill="x", expand=True, padx=(5, 0))

        header_right = ctk.CTkFrame(right_panel, fg_color="transparent")
        header_right.pack(fill="x", padx=15, pady=(15, 5))
        
        ctk.CTkLabel(header_right, text="Relatório de Processamento", font=ctk.CTkFont(weight="bold")).pack(side="left")
        self.lbl_status_processamento = ctk.CTkLabel(header_right, text="Status: Aguardando...", text_color="#AAAAAA")
        self.lbl_status_processamento.pack(side="right")
        
        style = ttk.Style()
        style.theme_use("default")
        style.configure("Treeview", background="#2b2b2b", foreground="white", fieldbackground="#2b2b2b", borderwidth=0, rowheight=25)
        style.map("Treeview", background=[('selected', '#1f538d')])
        
        tree_frame = ctk.CTkFrame(right_panel, fg_color="transparent")
        tree_frame.pack(fill="both", expand=True, padx=15, pady=(0, 15))
        
        self.tree_snow = ttk.Treeview(tree_frame, columns=("Tag", "RITM", "Status"), show='headings')
        self.tree_snow.heading("Tag", text="Service Tag")
        self.tree_snow.heading("RITM", text="Chamado (RITM)")
        self.tree_snow.heading("Status", text="Resultado")
        
        self.tree_snow.column("Tag", width=100, anchor="center")
        self.tree_snow.column("RITM", width=100, anchor="center")
        self.tree_snow.column("Status", width=200, anchor="w")
        self.tree_snow.pack(side="left", expand=True, fill="both")
        
        scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree_snow.yview)
        scrollbar.pack(side="right", fill="y")
        self.tree_snow.configure(yscrollcommand=scrollbar.set)
        
        self.snow_progress = ctk.CTkProgressBar(right_panel)
        self.snow_progress.pack(fill="x", padx=15, pady=(0, 15))
        self.snow_progress.set(0)

    def iniciar_processo_baixas(self):
        cookie = self.cookie_var.get().strip()
        tags_raw = self.snow_tags_text.get(1.0, "end").strip()
        
        if not cookie:
            messagebox.showwarning("Atenção", "Por favor, insira o Cookie do ServiceNow.")
            return
            
        if not tags_raw:
            messagebox.showwarning("Atenção", "Bipe pelo menos uma Service Tag para processar.")
            return

        self.salvar_cookie_local(cookie)
        
        lista_tags = []
        for linha in tags_raw.split('\n'):
            linha = linha.strip()
            if linha:
                if ',' in linha:
                    lista_tags.extend([t.strip().upper() for t in linha.split(',') if t.strip()])
                else:
                    lista_tags.append(linha.upper())
                    
        for item in self.tree_snow.get_children():
            self.tree_snow.delete(item)
            
        self.btn_executar_snow.configure(state="disabled", text="⏳ Processando...")
        self.snow_progress.set(0)
        self.lbl_status_processamento.configure(text=f"Processando 0 de {len(lista_tags)}...", text_color="#3498DB")
        
        threading.Thread(target=self._thread_executar_baixas, args=(lista_tags, cookie), daemon=True).start()

    def _thread_executar_baixas(self, lista_tags, cookie):
        headers = {
            "Accept": "application/json",
            "Content-Type": "application/json",
            "Cookie": cookie
        }
        url_base = "https://grupoboticario.service-now.com/api/now/table/sc_req_item"
        total = len(lista_tags)
        
        for idx, tag in enumerate(lista_tags):
            item_id = self.tree_snow.insert("", "end", values=(tag, "Buscando...", "Aguarde..."))
            self.tree_snow.see(item_id)
            
            query_lupa = f"123TEXTQUERY321={tag}^active=true^assignment_group={self.SYS_ID_FILA_LAB}"
            
            try:
                resp_busca = requests.get(url_base, headers=headers, params={"sysparm_query": query_lupa, "sysparm_fields": "number,sys_id", "sysparm_limit": 1})
                
                if resp_busca.status_code == 401:
                    self.tree_snow.item(item_id, values=(tag, "ERRO", "❌ Cookie Expirado/Inválido"))
                    break
                    
                if resp_busca.status_code != 200 or not resp_busca.json().get('result'):
                    self.tree_snow.item(item_id, values=(tag, "N/A", "⚠️ Não há chamados abertos"))
                    continue
                    
                chamado_alvo = resp_busca.json()['result'][0]
                numero_ritm = chamado_alvo.get('number')
                sys_id_ritm = chamado_alvo.get('sys_id')
                
                self.tree_snow.item(item_id, values=(tag, numero_ritm, "Fechando..."))
                
                url_fechar = f"{url_base}/{sys_id_ritm}"
                nota_fechamento = (
                    "Realizado a limpeza do equipamento, formatado e validado com imagem, entregue ao estoque.\n"
                    "********************************************************************************************************\n"
                    "Nossos serviços estão disponíveis 24 horas, todos os dias da semana, via Telefone (41) 0800 878 2774 ou Chat (https://autoatendimento.grupoboticario.com.br/), e você também pode contar com o apoio de nossa assistente virtual BETI (Área de Trabalho) e no Portal de Serviço TI."
                )
                dados_fechamento = {
                    "state": "3", 
                    "assigned_to": self.SYS_ID_GUSTAVO, 
                    "u_categoria_de_fechamento": "Equipamento Direcionado para Estoque",
                    "close_notes": nota_fechamento
                }
                
                resp_fechar = requests.patch(url_fechar, headers=headers, data=json.dumps(dados_fechamento))
                
                if resp_fechar.status_code == 200:
                    self.tree_snow.item(item_id, values=(tag, numero_ritm, "✅ SUCESSO (Fechado)"))
                else:
                    self.tree_snow.item(item_id, values=(tag, numero_ritm, f"❌ ERRO {resp_fechar.status_code} ao salvar"))
                    
            except Exception as e:
                self.tree_snow.item(item_id, values=(tag, "ERRO", "❌ Falha de Conexão"))
            
            progresso = (idx + 1) / total
            self.snow_progress.set(progresso)
            self.lbl_status_processamento.configure(text=f"Processando {idx + 1} de {total}...")
            
        self.after(0, self._concluir_processo_baixas)
        
    def _concluir_processo_baixas(self):
        self.btn_executar_snow.configure(state="normal", text="🚀 Executar Baixas em Massa")
        self.lbl_status_processamento.configure(text="Processamento Concluído!", text_color="#2AA876")
        self.snow_progress.set(1)
        
    # =========================================================================
    # RESTANTE DO SEU CÓDIGO ORIGINAL (INTACTO)
    # =========================================================================

    def criar_aba_individual(self):
        main_frame = ctk.CTkScrollableFrame(self.tab_individual, corner_radius=20)
        main_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        header_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        header_frame.pack(fill="x", padx=20, pady=(20, 10))
        
        title_label = ctk.CTkLabel(header_frame, text="Etiqueta Individual", font=ctk.CTkFont(size=24, weight="bold"))
        title_label.pack(side="left")
        
        self.status_indicator = ctk.CTkLabel(header_frame, text="●", text_color="#95A5A6", font=ctk.CTkFont(size=16))
        self.status_indicator.pack(side="right", padx=10)
        
        content_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        content_frame.pack(fill="both", expand=True, padx=20, pady=10)
        
        left_panel = ctk.CTkFrame(content_frame, corner_radius=15)
        left_panel.pack(side="left", fill="both", expand=True, padx=(0, 10))
        
        right_panel = ctk.CTkFrame(content_frame, corner_radius=15)
        right_panel.pack(side="right", fill="both", expand=True, padx=(10, 0))
        
        patrimonio_card = ctk.CTkFrame(left_panel, corner_radius=12)
        patrimonio_card.pack(fill="x", padx=15, pady=10)
        
        ctk.CTkLabel(patrimonio_card, text="Patrimônio", font=ctk.CTkFont(weight="bold")).pack(anchor="w", padx=15, pady=(15, 5))
        
        input_frame = ctk.CTkFrame(patrimonio_card, fg_color="transparent")
        input_frame.pack(fill="x", padx=15, pady=(0, 15))
        
        self.patrimonio_var = ctk.StringVar()
        self.patrimonio_entry = ctk.CTkEntry(input_frame, textvariable=self.patrimonio_var, placeholder_text="Digite ou escaneie o patrimônio...", height=40, font=ctk.CTkFont(size=14))
        self.patrimonio_entry.pack(fill="x")
        self.patrimonio_entry.bind("<Return>", self.buscar_patrimonio)
        self.patrimonio_entry.bind("<FocusOut>", self.buscar_patrimonio)
        
        self.alerta_frame = ctk.CTkFrame(patrimonio_card, fg_color="#E74C3C", corner_radius=8, height=40)
        self.alerta_label = ctk.CTkLabel(self.alerta_frame, text="⚠ PATRIMÔNIO NÃO ENCONTRADO ⚠", text_color="white", font=ctk.CTkFont(weight="bold"))
        
        config_frame = ctk.CTkFrame(left_panel, fg_color="transparent")
        config_frame.pack(fill="both", expand=True, padx=15, pady=10)
        
        memoria_card = ctk.CTkFrame(config_frame, corner_radius=12)
        memoria_card.pack(fill="x", pady=5)
        
        ctk.CTkLabel(memoria_card, text="Memória RAM", font=ctk.CTkFont(weight="bold")).pack(anchor="w", padx=15, pady=(15, 10))
        
        self.ram_var = ctk.StringVar(value="16GB")
        ram_buttons_frame = ctk.CTkFrame(memoria_card, fg_color="transparent")
        ram_buttons_frame.pack(fill="x", padx=15, pady=(0, 15))
        
        ram_default_frame = ctk.CTkFrame(ram_buttons_frame, fg_color="transparent")
        ram_default_frame.pack(side="top", fill="x")
        
        ctk.CTkRadioButton(ram_default_frame, text="16GB", variable=self.ram_var, value="16GB").pack(side="left", padx=(0, 10))
        ctk.CTkRadioButton(ram_default_frame, text="32GB", variable=self.ram_var, value="32GB").pack(side="left", padx=(0, 10))
        ctk.CTkRadioButton(ram_default_frame, text="64GB", variable=self.ram_var, value="64GB").pack(side="left")
        
        ram_custom_frame = ctk.CTkFrame(ram_buttons_frame, fg_color="transparent")
        ram_custom_frame.pack(side="top", fill="x", pady=(10, 0))
        
        ctk.CTkRadioButton(ram_custom_frame, text="Personalizado:", variable=self.ram_var, value="").pack(side="left", padx=(0, 10))
        
        self.ram_custom_var = ctk.StringVar()
        self.ram_custom_entry = ctk.CTkEntry(ram_custom_frame, textvariable=self.ram_custom_var, placeholder_text="Ex: 12", width=100)
        self.ram_custom_entry.pack(side="left")
        
        self.ram_custom_entry.bind("<KeyRelease>", self.validar_apenas_numeros)
        self.ram_custom_entry.bind("<FocusOut>", self.completar_memoria_gb)
        
        modelo_card = ctk.CTkFrame(config_frame, corner_radius=12)
        modelo_card.pack(fill="x", pady=5)
        
        ctk.CTkLabel(modelo_card, text="Modelo do Equipamento", font=ctk.CTkFont(weight="bold")).pack(anchor="w", padx=15, pady=(15, 10))
        
        self.modelo_var = ctk.StringVar(value="")
        self.modelo_dropdown = ctk.CTkComboBox(modelo_card, variable=self.modelo_var, height=35, dropdown_font=ctk.CTkFont(size=12))
        self.modelo_dropdown.pack(fill="x", padx=15, pady=(0, 15))
        
        row_frame = ctk.CTkFrame(config_frame, fg_color="transparent")
        row_frame.pack(fill="x", pady=5)
        
        tipo_card = ctk.CTkFrame(row_frame, corner_radius=12)
        tipo_card.pack(side="left", fill="x", expand=True, padx=(0, 5))
        
        ctk.CTkLabel(tipo_card, text="Tipo", font=ctk.CTkFont(weight="bold")).pack(anchor="w", padx=15, pady=(15, 10))
        
        self.tipo_var = ctk.StringVar(value="Não avaliado")
        tipo_buttons_frame = ctk.CTkFrame(tipo_card, fg_color="transparent")
        tipo_buttons_frame.pack(fill="x", padx=15, pady=(0, 15))
        
        ctk.CTkRadioButton(tipo_buttons_frame, text="Autopilot", variable=self.tipo_var, value="Autopilot").pack(side="left", padx=(0, 10))
        ctk.CTkRadioButton(tipo_buttons_frame, text="PXE", variable=self.tipo_var, value="PXE").pack(side="left", padx=(0, 10))
        ctk.CTkRadioButton(tipo_buttons_frame, text="Linux", variable=self.tipo_var, value="Linux").pack(side="left", padx=(0, 10))
        ctk.CTkRadioButton(tipo_buttons_frame, text="MacOs", variable=self.tipo_var, value="MacOs").pack(side="left", padx=(0, 10))
        
        estado_card = ctk.CTkFrame(row_frame, corner_radius=12)
        estado_card.pack(side="right", fill="x", expand=True, padx=(5, 0))
        
        ctk.CTkLabel(estado_card, text="Estado", font=ctk.CTkFont(weight="bold")).pack(anchor="w", padx=15, pady=(15, 10))
        
        self.estado_var = ctk.StringVar(value="Excelente")
        estado_buttons_frame = ctk.CTkFrame(estado_card, fg_color="transparent")
        estado_buttons_frame.pack(fill="x", padx=15, pady=(0, 15))
        
        ctk.CTkRadioButton(estado_buttons_frame, text="Excelente", variable=self.estado_var, value="Excelente").pack(anchor="w", pady=2)
        ctk.CTkRadioButton(estado_buttons_frame, text="Bom", variable=self.estado_var, value="Bom").pack(anchor="w", pady=2)
        ctk.CTkRadioButton(estado_buttons_frame, text="Emprestimo", variable=self.estado_var, value="Emprestimo").pack(anchor="w", pady=2)
        
        preview_card = ctk.CTkFrame(right_panel, corner_radius=12)
        preview_card.pack(fill="x", padx=15, pady=10)
        
        ctk.CTkLabel(preview_card, text="Pré-visualização", font=ctk.CTkFont(weight="bold")).pack(anchor="w", padx=15, pady=(15, 10))
        
        preview_content = ctk.CTkFrame(preview_card, fg_color="#2B2B2B", corner_radius=8, height=120)
        preview_content.pack(fill="x", padx=15, pady=(0, 15))
        
        self.preview_label = ctk.CTkLabel(preview_content, text="Pré-visualização da etiqueta aparecerá aqui", text_color="#AAAAAA", font=ctk.CTkFont(size=12))
        self.preview_label.pack(expand=True)
        
        print_card = ctk.CTkFrame(right_panel, corner_radius=12)
        print_card.pack(fill="x", padx=15, pady=10)
        
        ctk.CTkLabel(print_card, text="Configurações de Impressão", font=ctk.CTkFont(weight="bold")).pack(anchor="w", padx=15, pady=(15, 10))
        
        qtd_frame = ctk.CTkFrame(print_card, fg_color="transparent")
        qtd_frame.pack(fill="x", padx=15, pady=5)
        
        ctk.CTkLabel(qtd_frame, text="Cópias:").pack(side="left")
        self.qtd_var = ctk.StringVar(value="1")
        self.qtd_entry = ctk.CTkEntry(qtd_frame, textvariable=self.qtd_var, width=60)
        self.qtd_entry.pack(side="right")
        
        printer_frame = ctk.CTkFrame(print_card, fg_color="transparent")
        printer_frame.pack(fill="x", padx=15, pady=(5, 15))
        
        ctk.CTkLabel(printer_frame, text="Impressora:").pack(anchor="w")
        self.printer_var = ctk.StringVar(value="")
        self.printer_dropdown = ctk.CTkComboBox(printer_frame, variable=self.printer_var, height=35)
        self.printer_dropdown.pack(fill="x", pady=5)
        
        button_frame = ctk.CTkFrame(right_panel, fg_color="transparent")
        button_frame.pack(fill="x", padx=15, pady=20)
        
        self.imprimir_btn = ctk.CTkButton(button_frame, text="🖨️ Imprimir Etiqueta", command=self.imprimir_etiqueta_individual, fg_color="#95A5A6", hover_color="#7F8C8D", state="disabled", height=40, font=ctk.CTkFont(weight="bold"))
        self.imprimir_btn.pack(fill="x", pady=5)
        
        action_buttons_frame = ctk.CTkFrame(button_frame, fg_color="transparent")
        action_buttons_frame.pack(fill="x", pady=5)
        
        self.salvar_btn = ctk.CTkButton(action_buttons_frame, text="💾 Salvar no Banco", command=self.salvar_no_banco, fg_color="#3498DB", hover_color="#2980B9", height=35)
        self.salvar_btn.pack(side="left", fill="x", expand=True, padx=(0, 5))
        
        self.limpar_btn = ctk.CTkButton(action_buttons_frame, text="🗑️ Limpar", command=self.limpar_campos, fg_color="#E74C3C", hover_color="#C0392B", height=35)
        self.limpar_btn.pack(side="right", fill="x", expand=True, padx=(5, 0))

        self.excluir_btn = ctk.CTkButton(button_frame, text="🗑️ Excluir Registro do Banco", command=self.excluir_maquina, fg_color="#C0392B", hover_color="#922B21", height=35)
        self.excluir_btn.pack(fill="x", pady=(10, 5))
        
        status_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        status_frame.pack(fill="x", padx=20, pady=(10, 20))
        
        self.patrimonio_entry.focus_set()

    def criar_aba_estatisticas(self):
        for widget in self.tab_estatisticas.winfo_children():
            widget.destroy()

        main_frame = ctk.CTkScrollableFrame(self.tab_estatisticas, corner_radius=20)
        main_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        filter_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        filter_frame.pack(fill="x", padx=20, pady=10)
        
        ctk.CTkLabel(filter_frame, text="Período:", font=ctk.CTkFont(weight="bold")).pack(side="left", padx=(0, 5))
        self.filtro_estatistica_var = ctk.StringVar(value="Tudo")
        self.filtro_estatistica_combo = ctk.CTkComboBox(filter_frame, values=["Tudo", "Hoje", "Esta Semana", "Este Mês"], variable=self.filtro_estatistica_var, command=lambda x: self.atualizar_estatisticas(), width=130)
        self.filtro_estatistica_combo.pack(side="left", padx=(0, 15))

        ctk.CTkLabel(filter_frame, text="Modelo:", font=ctk.CTkFont(weight="bold")).pack(side="left", padx=(0, 5))
        self.filtro_modelo_var = ctk.StringVar(value="Todos")
        self.filtro_modelo_combo = ctk.CTkComboBox(filter_frame, values=["Todos"], variable=self.filtro_modelo_var, command=lambda x: self.atualizar_estatisticas(), width=150)
        self.filtro_modelo_combo.pack(side="left", padx=(0, 15))

        ctk.CTkLabel(filter_frame, text="Tipo:", font=ctk.CTkFont(weight="bold")).pack(side="left", padx=(0, 5))
        self.filtro_tipo_var = ctk.StringVar(value="Todos")
        self.filtro_tipo_combo = ctk.CTkComboBox(filter_frame, values=["Todos", "Autopilot", "PXE", "Linux", "MacOs", "Não avaliado"], variable=self.filtro_tipo_var, command=lambda x: self.atualizar_estatisticas(), width=130)
        self.filtro_tipo_combo.pack(side="left", padx=(0, 15))

        self.exportar_stats_btn = ctk.CTkButton(filter_frame, text="📊 Baixar Relatório", command=self.baixar_planilha_estatisticas, fg_color="#27AE60", hover_color="#1E8449", height=30)
        self.exportar_stats_btn.pack(side="right", padx=10)

        style = ttk.Style()
        style.theme_use("default")
        style.configure("Treeview", background="#2b2b2b", foreground="white", fieldbackground="#2b2b2b", borderwidth=0)
        style.map("Treeview", background=[('selected', '#1f538d')])
        
        self.tree_stats = ttk.Treeview(main_frame, columns=("Modelo", "Tipo", "Quantidade"), show='headings')
        self.tree_stats.heading("Modelo", text="Modelo")
        self.tree_stats.heading("Tipo", text="Tipo")
        self.tree_stats.heading("Quantidade", text="Quantidade")
        self.tree_stats.pack(expand=True, fill="both", padx=20, pady=10)

        self.total_stats_label = ctk.CTkLabel(main_frame, text="Total de Máquinas: 0", font=ctk.CTkFont(size=18, weight="bold"))
        self.total_stats_label.pack(anchor="e", padx=20, pady=(0, 10))
        
    def atualizar_estatisticas(self):
        for item in self.tree_stats.get_children():
            self.tree_stats.delete(item)
            
        periodo = self.filtro_estatistica_var.get()
        modelo_filtro = self.filtro_modelo_var.get()
        tipo_filtro = self.filtro_tipo_var.get()
        
        condicoes = []
        params = []
        
        hoje = datetime.date.today()
        if periodo == "Hoje":
            condicoes.append("date(data_cadastro) = ?")
            params.append(hoje.strftime("%Y-%m-%d"))
        elif periodo == "Esta Semana":
            dias_para_domingo = (hoje.weekday() + 1) % 7
            domingo = hoje - datetime.timedelta(days=dias_para_domingo)
            condicoes.append("date(data_cadastro) >= ?")
            params.append(domingo.strftime("%Y-%m-%d"))
        elif periodo == "Este Mês":
            condicoes.append("date(data_cadastro) >= ?")
            params.append(hoje.replace(day=1).strftime("%Y-%m-%d"))

        if modelo_filtro != "Todos":
            condicoes.append("modelo = ?")
            params.append(modelo_filtro)

        if tipo_filtro != "Todos":
            condicoes.append("tipo = ?")
            params.append(tipo_filtro)

        where_clause = ""
        if condicoes:
            where_clause = "WHERE " + " AND ".join(condicoes)

        try:
            query = f'''
                SELECT modelo, tipo, COUNT(*) 
                FROM patrimonios 
                {where_clause}
                GROUP BY modelo, tipo 
                ORDER BY COUNT(*) DESC
            '''
            if params:
                self.cursor_patrimonios.execute(query, tuple(params))
            else:
                self.cursor_patrimonios.execute(query)
                
            total_maquinas = 0
                
            for row in self.cursor_patrimonios.fetchall():
                self.tree_stats.insert("", "end", values=row)
                total_maquinas += row[2] 
                
            if hasattr(self, 'total_stats_label'):
                self.total_stats_label.configure(text=f"Total de Máquinas: {total_maquinas}")

            if hasattr(self, 'modelos_lista'):
                lista_atualizada = ["Todos"] + self.modelos_lista
                if self.filtro_modelo_combo.cget("values") != lista_atualizada:
                    self.filtro_modelo_combo.configure(values=lista_atualizada)
                
        except sqlite3.Error as e:
            print(f"Erro SQL ao atualizar estatísticas: {e}")
    
    def excluir_maquina(self):
        patrimonio_digitado = self.patrimonio_var.get().strip()
        
        if not patrimonio_digitado:
            messagebox.showwarning("Aviso", "Digite ou escaneie o número do patrimônio que deseja excluir.")
            return

        patrimonio_sem_zero = patrimonio_digitado.lstrip('0')

        try:
            self.cursor_patrimonios.execute("SELECT modelo, tipo, patrimonio FROM patrimonios WHERE patrimonio = ?", (patrimonio_digitado,))
            maquina = self.cursor_patrimonios.fetchone()
            
            if not maquina and patrimonio_digitado != patrimonio_sem_zero:
                self.cursor_patrimonios.execute("SELECT modelo, tipo, patrimonio FROM patrimonios WHERE patrimonio = ?", (patrimonio_sem_zero,))
                maquina = self.cursor_patrimonios.fetchone()
            
            if not maquina:
                messagebox.showerror("Erro", f"O patrimônio '{patrimonio_digitado}' não foi encontrado no sistema.")
                return
                
            patrimonio_banco = maquina[2]
                
            if not messagebox.askyesno("Confirmação", f"Tem certeza que deseja excluir a máquina {patrimonio_banco} ({maquina[0]} - {maquina[1]})?"):
                return
                
            dialog = ctk.CTkToplevel(self)
            dialog.title("Autenticação Necessária")
            dialog.geometry("400x220")
            dialog.transient(self) 
            dialog.grab_set()      
            
            dialog.update_idletasks()
            x = self.winfo_x() + (self.winfo_width() // 2) - (400 // 2)
            y = self.winfo_y() + (self.winfo_height() // 2) - (220 // 2)
            dialog.geometry(f"+{x}+{y}")

            ctk.CTkLabel(dialog, text="Digite a senha de administrador\npara confirmar a exclusão:", font=ctk.CTkFont(size=14, weight="bold")).pack(pady=(20, 10))
            
            entry_senha = ctk.CTkEntry(dialog, show="*", width=200, justify="center")
            entry_senha.pack(pady=10)
            entry_senha.focus_set()
            
            senha_digitada = [None] 
            
            def confirmar(event=None):
                senha_digitada[0] = entry_senha.get()
                dialog.destroy()
                
            def cancelar():
                dialog.destroy()
                
            entry_senha.bind("<Return>", confirmar)
            
            btn_frame = ctk.CTkFrame(dialog, fg_color="transparent")
            btn_frame.pack(pady=(10, 20))
            
            ctk.CTkButton(btn_frame, text="Confirmar", command=confirmar, width=100, fg_color="#2AA876", hover_color="#207A59").pack(side="left", padx=10)
            ctk.CTkButton(btn_frame, text="Cancelar", command=cancelar, width=100, fg_color="#E74C3C", hover_color="#C0392B").pack(side="right", padx=10)
            
            self.wait_window(dialog)
            
            senha = senha_digitada[0]
            
            if senha == "AdminTechLab":
                self.cursor_patrimonios.execute("DELETE FROM patrimonios WHERE patrimonio = ?", (patrimonio_banco,))
                self.conn_patrimonios.commit()
                
                self.limpar_campos()
                self.atualizar_estatisticas()
                
                messagebox.showinfo("Sucesso", f"Máquina {patrimonio_banco} excluída permanentemente!")
                
            elif senha is not None: 
                messagebox.showerror("Acesso Negado", "Senha incorreta! A exclusão foi cancelada.")
                
        except sqlite3.Error as e:
            messagebox.showerror("Erro de Banco de Dados", f"Erro ao tentar excluir: {e}")        

    def baixar_planilha_estatisticas(self):
        periodo = self.filtro_estatistica_var.get()
        modelo_filtro = self.filtro_modelo_var.get()
        tipo_filtro = self.filtro_tipo_var.get()
        
        condicoes = []
        params = []
        
        hoje = datetime.date.today()
        
        if periodo == "Hoje":
            condicoes.append("date(data_cadastro, 'localtime') = ?")
            params.append(hoje.strftime("%Y-%m-%d"))
        elif periodo == "Esta Semana":
            dias_para_domingo = (hoje.weekday() + 1) % 7
            domingo = hoje - datetime.timedelta(days=dias_para_domingo)
            condicoes.append("date(data_cadastro, 'localtime') >= ?")
            params.append(domingo.strftime("%Y-%m-%d"))
        elif periodo == "Este Mês":
            condicoes.append("date(data_cadastro, 'localtime') >= ?")
            params.append(hoje.replace(day=1).strftime("%Y-%m-%d"))

        if modelo_filtro != "Todos":
            condicoes.append("modelo = ?")
            params.append(modelo_filtro)

        if tipo_filtro != "Todos":
            condicoes.append("tipo = ?")
            params.append(tipo_filtro)

        where_clause = ""
        if condicoes:
            where_clause = "WHERE " + " AND ".join(condicoes)

        try:
            query = f'''
                SELECT patrimonio, modelo, tipo, memoria, estado, datetime(data_cadastro, 'localtime') 
                FROM patrimonios 
                {where_clause}
                ORDER BY data_cadastro DESC
            '''
            
            if params:
                self.cursor_patrimonios.execute(query, tuple(params))
            else:
                self.cursor_patrimonios.execute(query)
                
            maquinas = self.cursor_patrimonios.fetchall()
            
            if not maquinas:
                messagebox.showinfo("Vazio", "Nenhuma máquina encontrada para os filtros selecionados.")
                return
                
            nome_padrao = f"relatorio_estoque.csv"
            
            from tkinter import filedialog
            filename = filedialog.asksaveasfilename(
                defaultextension=".csv",
                filetypes=[("Planilha Excel (CSV)", "*.csv"), ("Todos os arquivos", "*.*")],
                initialfile=nome_padrao
            )
            
            if filename:
                with open(filename, 'w', newline='', encoding='utf-8-sig') as file:
                    writer = csv.writer(file, delimiter=';')
                    writer.writerow(['Patrimônio', 'Modelo', 'Tipo', 'Memória', 'Estado', 'Data/Hora de Cadastro'])
                    
                    for maquina in maquinas:
                        patrimonio, modelo, tipo, memoria, estado, data_cadastro = maquina
                        data_formatada = data_cadastro
                        try:
                            if data_cadastro:
                                dt = datetime.datetime.strptime(data_cadastro, "%Y-%m-%d %H:%M:%S")
                                data_formatada = dt.strftime("%d/%m/%Y %H:%M")
                        except ValueError: pass
                            
                        writer.writerow([patrimonio, modelo, tipo, memoria, estado, data_formatada])
                        
                messagebox.showinfo("Sucesso", f"Relatório exportado!\n{len(maquinas)} máquinas registradas na planilha.")
                
        except sqlite3.Error as e:
            messagebox.showerror("Erro no Banco", f"Falha ao consultar banco de dados: {e}")
    
    def auto_refresh_stats(self):
        self.atualizar_estatisticas()
        self.after(5000, self.auto_refresh_stats)

    def validar_apenas_numeros(self, event=None):
        self.ram_var.set("")
        
        valor_atual = self.ram_custom_var.get().upper()
        
        if valor_atual.endswith("GB"):
            apenas_numeros = ''.join(filter(str.isdigit, valor_atual[:-2]))
            if valor_atual != f"{apenas_numeros}GB":
                self.ram_custom_var.set(f"{apenas_numeros}GB")
        else:
            apenas_numeros = ''.join(filter(str.isdigit, valor_atual))
            if valor_atual != apenas_numeros:
                self.ram_custom_var.set(apenas_numeros)
                
        self.atualizar_preview()

    def completar_memoria_gb(self, event=None):
        valor = self.ram_custom_var.get()
        apenas_numeros = ''.join(filter(str.isdigit, valor))
        
        if apenas_numeros:
            self.ram_custom_var.set(f"{apenas_numeros}GB")
            self.ram_var.set("")
        else:
            self.ram_custom_var.set("")
            
        self.atualizar_preview()

    def criar_aba_lote(self):
        main_frame = ctk.CTkFrame(self.tab_lote, corner_radius=20)
        main_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        header_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        header_frame.pack(fill="x", padx=20, pady=(20, 10))
        
        title_label = ctk.CTkLabel(header_frame, text="Impressão em Lote", font=ctk.CTkFont(size=24, weight="bold"))
        title_label.pack(side="left")
        
        content_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        content_frame.pack(fill="both", expand=True, padx=20, pady=10)
        
        left_panel = ctk.CTkFrame(content_frame, corner_radius=15)
        left_panel.pack(side="left", fill="both", expand=True, padx=(0, 10))
        
        right_panel = ctk.CTkFrame(content_frame, corner_radius=15)
        right_panel.pack(side="right", fill="both", expand=False, padx=(10, 0), pady=10)
        right_panel.pack_propagate(False)
        right_panel.configure(width=300)
        
        input_card = ctk.CTkFrame(left_panel, corner_radius=12)
        input_card.pack(fill="both", expand=True, padx=15, pady=15)
        
        ctk.CTkLabel(input_card, text="Lista de Patrimônios", font=ctk.CTkFont(weight="bold")).pack(anchor="w", padx=15, pady=(15, 10))
        
        ctk.CTkLabel(input_card, text="Digite os patrimônios (um por linha ou separados por vírgula):", text_color="#AAAAAA").pack(anchor="w", padx=15, pady=(0, 10))
        
        self.patrimonios_text = scrolledtext.ScrolledText(input_card, height=15, font=("Consolas", 11), bg="#2B2B2B", fg="white", insertbackground="white", relief="flat", bd=0)
        self.patrimonios_text.pack(fill="both", expand=True, padx=15, pady=(0, 15))
        
        config_card = ctk.CTkFrame(right_panel, corner_radius=12)
        config_card.pack(fill="x", padx=15, pady=(0, 15))
        
        ctk.CTkLabel(config_card, text="Configurações", font=ctk.CTkFont(weight="bold")).pack(anchor="w", padx=15, pady=(15, 10))
        
        qtd_frame = ctk.CTkFrame(config_card, fg_color="transparent")
        qtd_frame.pack(fill="x", padx=15, pady=5)
        
        ctk.CTkLabel(qtd_frame, text="Cópias por patrimônio:").pack(anchor="w")
        self.qtd_lote_var = ctk.StringVar(value="1")
        self.qtd_lote_entry = ctk.CTkEntry(qtd_frame, textvariable=self.qtd_lote_var, width=80)
        self.qtd_lote_entry.pack(anchor="w", pady=5)
        
        printer_frame = ctk.CTkFrame(config_card, fg_color="transparent")
        printer_frame.pack(fill="x", padx=15, pady=(5, 15))
        
        ctk.CTkLabel(printer_frame, text="Impressora:").pack(anchor="w")
        self.printer_lote_var = ctk.StringVar(value="")
        self.printer_lote_dropdown = ctk.CTkComboBox(printer_frame, variable=self.printer_lote_var, height=35)
        self.printer_lote_dropdown.pack(fill="x", pady=5)
        
        action_card = ctk.CTkFrame(right_panel, corner_radius=12)
        action_card.pack(fill="x", padx=15, pady=(0, 15))
        
        ctk.CTkLabel(action_card, text="Ações", font=ctk.CTkFont(weight="bold")).pack(anchor="w", padx=15, pady=(15, 10))
        
        self.carregar_btn = ctk.CTkButton(action_card, text="📂 Carregar do Banco", command=self.carregar_patrimonios_banco, fg_color="#3498DB", hover_color="#2980B9", height=35)
        self.carregar_btn.pack(fill="x", padx=15, pady=5)
        
        self.baixar_planilha_btn = ctk.CTkButton(action_card, text="📊 Baixar Planilha", command=self.baixar_planilha_patrimonios, fg_color="#27AE60", hover_color="#1E8449", height=35)
        self.baixar_planilha_btn.pack(fill="x", padx=15, pady=5)
        
        self.imprimir_lote_btn = ctk.CTkButton(action_card, text="🖨️ Imprimir Lote", command=self.imprimir_em_lote, fg_color="#2AA876", hover_color="#207A59", height=35)
        self.imprimir_lote_btn.pack(fill="x", padx=15, pady=5)
        
        self.limpar_lote_btn = ctk.CTkButton(action_card, text="🗑️ Limpar Lista", command=self.limpar_lista_patrimonios, fg_color="#E74C3C", hover_color="#C0392B", height=35)
        self.limpar_lote_btn.pack(fill="x", padx=15, pady=5)
        
        status_card = ctk.CTkFrame(right_panel, corner_radius=12)
        status_card.pack(fill="x", padx=15, pady=(0, 15))
        
        ctk.CTkLabel(status_card, text="Status", font=ctk.CTkFont(weight="bold")).pack(anchor="w", padx=15, pady=(15, 10))
        
        self.status_lote_label = ctk.CTkLabel(status_card, text="Aguardando entrada de dados...", wraplength=250, justify="left")
        self.status_lote_label.pack(anchor="w", padx=15, pady=(0, 15))
    
    def criar_aba_gerenciar_modelos(self):
        main_frame = ctk.CTkFrame(self.tab_modelos, corner_radius=20)
        main_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        header_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        header_frame.pack(fill="x", padx=20, pady=(20, 10))
        
        title_label = ctk.CTkLabel(header_frame, text="Gerenciar Modelos de Equipamentos", font=ctk.CTkFont(size=24, weight="bold"))
        title_label.pack(side="left")
        
        content_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        content_frame.pack(fill="both", expand=True, padx=20, pady=10)
        
        left_panel = ctk.CTkFrame(content_frame, corner_radius=15)
        left_panel.pack(side="left", fill="both", expand=True, padx=(0, 10))
        
        right_panel = ctk.CTkFrame(content_frame, corner_radius=15)
        right_panel.pack(side="right", fill="both", expand=False, padx=(10, 0))
        right_panel.configure(width=300)
        
        ctk.CTkLabel(left_panel, text="Modelos Cadastrados", font=ctk.CTkFont(weight="bold")).pack(anchor="w", padx=15, pady=(15, 10))
        
        lista_frame = ctk.CTkFrame(left_panel, corner_radius=12)
        lista_frame.pack(fill="both", expand=True, padx=15, pady=(0, 15))
        
        self.modelos_listbox = ctk.CTkTextbox(lista_frame, wrap="word", font=ctk.CTkFont(size=12))
        self.modelos_listbox.pack(fill="both", expand=True, padx=15, pady=15)
        
        atualizar_btn = ctk.CTkButton(left_panel, text="🔄 Atualizar Lista", command=self.carregar_lista_modelos, height=35)
        atualizar_btn.pack(fill="x", padx=15, pady=(0, 15))
        
        ctk.CTkLabel(right_panel, text="Gerenciar Modelos", font=ctk.CTkFont(weight="bold")).pack(anchor="w", padx=15, pady=(15, 10))
        
        add_card = ctk.CTkFrame(right_panel, corner_radius=12)
        add_card.pack(fill="x", padx=15, pady=(0, 15))
        
        ctk.CTkLabel(add_card, text="Adicionar Novo Modelo", font=ctk.CTkFont(weight="bold")).pack(anchor="w", padx=15, pady=(15, 10))
        
        self.novo_modelo_var = ctk.StringVar()
        self.novo_modelo_entry = ctk.CTkEntry(add_card, textvariable=self.novo_modelo_var, placeholder_text="Digite o nome do novo modelo...", height=35)
        self.novo_modelo_entry.pack(fill="x", padx=15, pady=5)
        
        add_btn = ctk.CTkButton(add_card, text="➕ Adicionar Modelo", command=self.adicionar_modelo, fg_color="#2AA876", hover_color="#207A59", height=35)
        add_btn.pack(fill="x", padx=15, pady=(5, 15))
        
        remove_card = ctk.CTkFrame(right_panel, corner_radius=12)
        remove_card.pack(fill="x", padx=15, pady=(0, 15))
        
        ctk.CTkLabel(remove_card, text="Remover Modelo", font=ctk.CTkFont(weight="bold")).pack(anchor="w", padx=15, pady=(15, 10))
        
        self.modelo_remover_var = ctk.StringVar(value="")
        self.modelo_remover_dropdown = ctk.CTkComboBox(remove_card, variable=self.modelo_remover_var, height=35)
        self.modelo_remover_dropdown.pack(fill="x", padx=15, pady=5)
        
        remove_btn = ctk.CTkButton(remove_card, text="🗑️ Remover Modelo", command=self.remover_modelo, fg_color="#E74C3C", hover_color="#C0392B", height=35)
        remove_btn.pack(fill="x", padx=15, pady=(5, 15))
        
        status_card = ctk.CTkFrame(right_panel, corner_radius=12)
        status_card.pack(fill="x", padx=15, pady=(0, 15))
        
        ctk.CTkLabel(status_card, text="Status", font=ctk.CTkFont(weight="bold")).pack(anchor="w", padx=15, pady=(15, 10))
        
        self.status_modelos_label = ctk.CTkLabel(status_card, text="Gerencie os modelos de equipamentos aqui", wraplength=250, justify="left")
        self.status_modelos_label.pack(anchor="w", padx=15, pady=(0, 15))
        
        self.carregar_lista_modelos()
        self.atualizar_combobox_remover()
    
    def carregar_lista_modelos(self):
        try:
            self.cursor_modelos.execute("SELECT modelo FROM tipos_equipamento ORDER BY modelo")
            modelos = self.cursor_modelos.fetchall()
            
            self.modelos_listbox.delete(1.0, "end")
            
            if modelos:
                for modelo in modelos:
                    self.modelos_listbox.insert("end", f"• {modelo[0]}\n")
                self.status_modelos_label.configure(text=f"✓ {len(modelos)} modelos carregados")
            else:
                self.modelos_listbox.insert("end", "Nenhum modelo cadastrado")
                self.status_modelos_label.configure(text="Nenhum modelo cadastrado")
                
        except sqlite3.Error as e:
            self.modelos_listbox.delete(1.0, "end")
            self.modelos_listbox.insert("end", f"Erro ao carregar modelos: {e}")
            self.status_modelos_label.configure(text="Erro ao carregar modelos")
    
    def atualizar_combobox_remover(self):
        try:
            self.cursor_modelos.execute("SELECT modelo FROM tipos_equipamento ORDER BY modelo")
            modelos = self.cursor_modelos.fetchall()
            
            modelos_lista = [modelo[0] for modelo in modelos]
            self.modelo_remover_dropdown.configure(values=modelos_lista)
            if modelos_lista:
                self.modelo_remover_dropdown.set(modelos_lista[0])
            else:
                self.modelo_remover_dropdown.set("")
                
        except sqlite3.Error as e:
            self.modelo_remover_dropdown.configure(values=[])
            self.modelo_remover_dropdown.set("")
    
    def adicionar_modelo(self):
        novo_modelo = self.novo_modelo_var.get().strip()
        
        if not novo_modelo:
            messagebox.showwarning("Campo Vazio", "Digite o nome do modelo")
            return
        
        try:
            self.cursor_modelos.execute("INSERT INTO tipos_equipamento (modelo) VALUES (?)", (novo_modelo,))
            self.conn_modelos.commit()
            
            self.novo_modelo_var.set("")
            self.carregar_lista_modelos()
            self.atualizar_combobox_remover()
            self.carregar_modelos() 
            
            self.status_modelos_label.configure(text=f"✓ Modelo '{novo_modelo}' adicionado com sucesso!")
            
        except sqlite3.IntegrityError:
            messagebox.showwarning("Modelo Existente", f"O modelo '{novo_modelo}' já existe na base de dados!")
        except sqlite3.Error as e:
            messagebox.showerror("Erro", f"Erro ao adicionar modelo: {e}")
    
    def remover_modelo(self):
        modelo_remover = self.modelo_remover_var.get().strip()
        
        if not modelo_remover:
            messagebox.showwarning("Seleção Vazia", "Selecione um modelo para remover")
            return
        
        resposta = messagebox.askyesno(
            "Confirmar Remoção", 
            f"Tem certeza que deseja remover o modelo '{modelo_remover}'?\n\nEsta ação não pode ser desfeita."
        )
        
        if resposta:
            try:
                self.cursor_modelos.execute("DELETE FROM tipos_equipamento WHERE modelo = ?", (modelo_remover,))
                self.conn_modelos.commit()
                
                self.carregar_lista_modelos()
                self.atualizar_combobox_remover()
                self.carregar_modelos() 
                
                self.status_modelos_label.configure(text=f"✓ Modelo '{modelo_remover}' removido com sucesso!")
                
            except sqlite3.Error as e:
                messagebox.showerror("Erro", f"Erro ao remover modelo: {e}")
    
    def criar_aba_configuracoes(self):
        main_frame = ctk.CTkFrame(self.tab_config, corner_radius=20)
        main_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        header_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        header_frame.pack(fill="x", padx=20, pady=(20, 10))
        
        title_label = ctk.CTkLabel(header_frame, text="Configurações de Impressão Zebra", font=ctk.CTkFont(size=24, weight="bold"))
        title_label.pack(side="left")
        
        self.status_zebra_label = ctk.CTkLabel(header_frame, text="● Nenhuma impressora Zebra detectada", text_color="#E74C3C", font=ctk.CTkFont(size=12))
        self.status_zebra_label.pack(side="right", padx=10)
        
        content_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        content_frame.pack(fill="both", expand=True, padx=20, pady=10)
        
        left_panel = ctk.CTkFrame(content_frame, corner_radius=15)
        left_panel.pack(side="left", fill="both", expand=True, padx=(0, 10))
        
        right_panel = ctk.CTkFrame(content_frame, corner_radius=15)
        right_panel.pack(side="right", fill="both", expand=True, padx=(10, 0))
        
        params_card = ctk.CTkFrame(left_panel, corner_radius=12)
        params_card.pack(fill="x", padx=15, pady=15)
        
        ctk.CTkLabel(params_card, text="Parâmetros de Impressão", font=ctk.CTkFont(weight="bold")).pack(anchor="w", padx=15, pady=(15, 10))
        
        temp_frame = ctk.CTkFrame(params_card, fg_color="transparent")
        temp_frame.pack(fill="x", padx=15, pady=8)
        
        ctk.CTkLabel(temp_frame, text="Temperatura (Darkness):").pack(side="left")
        self.temp_var = ctk.StringVar(value=str(self.config_impressao['temperatura']))
        self.temp_entry = ctk.CTkEntry(temp_frame, textvariable=self.temp_var, width=60)
        self.temp_entry.pack(side="right")
        ctk.CTkLabel(temp_frame, text="(0-30)", text_color="#AAAAAA").pack(side="right", padx=5)
        
        speed_frame = ctk.CTkFrame(params_card, fg_color="transparent")
        speed_frame.pack(fill="x", padx=15, pady=8)
        
        ctk.CTkLabel(speed_frame, text="Velocidade:").pack(side="left")
        self.speed_var = ctk.StringVar(value=str(self.config_impressao['velocidade']))
        self.speed_entry = ctk.CTkEntry(speed_frame, textvariable=self.speed_var, width=60)
        self.speed_entry.pack(side="right")
        ctk.CTkLabel(speed_frame, text="(1-10)", text_color="#AAAAAA").pack(side="right", padx=5)
        
        tone_frame = ctk.CTkFrame(params_card, fg_color="transparent")
        tone_frame.pack(fill="x", padx=15, pady=8)
        
        ctk.CTkLabel(tone_frame, text="Tonalidade (Tear-off):").pack(side="left")
        self.tone_var = ctk.StringVar(value=str(self.config_impressao['tonalidade']))
        self.tone_entry = ctk.CTkEntry(tone_frame, textvariable=self.tone_var, width=60)
        self.tone_entry.pack(side="right")
        ctk.CTkLabel(tone_frame, text="(0-10)", text_color="#AAAAAA").pack(side="right", padx=5)
        
        etiqueta_card = ctk.CTkFrame(right_panel, corner_radius=12)
        etiqueta_card.pack(fill="x", padx=15, pady=15)
        
        ctk.CTkLabel(etiqueta_card, text="Dimensões da Etiqueta", font=ctk.CTkFont(weight="bold")).pack(anchor="w", padx=15, pady=(15, 10))
        
        width_frame = ctk.CTkFrame(etiqueta_card, fg_color="transparent")
        width_frame.pack(fill="x", padx=15, pady=8)
        
        ctk.CTkLabel(width_frame, text="Largura (mm):").pack(side="left")
        self.width_var = ctk.StringVar(value=str(self.config_impressao['largura_etiqueta']))
        self.width_entry = ctk.CTkEntry(width_frame, textvariable=self.width_var, width=60)
        self.width_entry.pack(side="right")
        
        height_frame = ctk.CTkFrame(etiqueta_card, fg_color="transparent")
        height_frame.pack(fill="x", padx=15, pady=8)
        
        ctk.CTkLabel(height_frame, text="Altura (mm):").pack(side="left")
        self.height_var = ctk.StringVar(value=str(self.config_impressao['altura_etiqueta']))
        self.height_entry = ctk.CTkEntry(height_frame, textvariable=self.height_var, width=60)
        self.height_entry.pack(side="right")
        
        margem_frame = ctk.CTkFrame(etiqueta_card, fg_color="transparent")
        margem_frame.pack(fill="x", padx=15, pady=8)
        
        ctk.CTkLabel(margem_frame, text="Margem Esquerda (mm):").pack(side="left")
        self.margin_var = ctk.StringVar(value=str(self.config_impressao['margem_esquerda']))
        self.margin_entry = ctk.CTkEntry(margem_frame, textvariable=self.margin_var, width=60)
        self.margin_entry.pack(side="right")
        
        action_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        action_frame.pack(fill="x", padx=20, pady=20)
        
        button_frame = ctk.CTkFrame(action_frame, fg_color="transparent")
        button_frame.pack(side="right")
        
        self.salvar_config_btn = ctk.CTkButton(button_frame, text="💾 Salvar Configurações", command=self.salvar_configuracoes, fg_color="#2AA876", hover_color="#229954", height=35)
        self.salvar_config_btn.pack(side="left", padx=5)
        
        self.aplicar_config_btn = ctk.CTkButton(button_frame, text="✓ Aplicar", command=self.aplicar_configuracoes, fg_color="#3498DB", hover_color="#2980B9", height=35)
        self.aplicar_config_btn.pack(side="left", padx=5)
        
        self.reset_config_btn = ctk.CTkButton(button_frame, text="🔄 Resetar Padrão", command=self.resetar_configuracoes, fg_color="#E74C3C", hover_color="#C0392B", height=35)
        self.reset_config_btn.pack(side="left", padx=5)
        
        self.test_config_btn = ctk.CTkButton(button_frame, text="🧪 Testar Config", command=self.testar_configuracoes, fg_color="#F39C12", hover_color="#D68910", height=35)
        self.test_config_btn.pack(side="left", padx=5)
        
        self.atualizar_status_zebra()
    
    def atualizar_status_zebra(self):
        if hasattr(self, 'impressoras_zebra') and self.impressoras_zebra:
            status_text = f"● {len(self.impressoras_zebra)} impressora(s) Zebra detectada(s)"
            self.status_zebra_label.configure(text=status_text, text_color="#2AA876")
        else:
            self.status_zebra_label.configure(text="● Nenhuma impressora Zebra detectada", text_color="#E74C3C")
    
    def salvar_configuracoes(self):
        try:
            self.aplicar_configuracoes()
            messagebox.showinfo("Sucesso", "Configurações salvas com sucesso!")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao salvar configurações: {e}")
    
    def aplicar_configuracoes(self):
        try:
            temperatura = int(self.temp_var.get())
            velocidade = int(self.speed_var.get())
            tonalidade = int(self.tone_var.get())
            largura = int(self.width_var.get())
            altura = int(self.height_var.get())
            margem = int(self.margin_var.get())
            
            if not (0 <= temperatura <= 30):
                raise ValueError("Temperatura deve estar entre 0 e 30")
            if not (1 <= velocidade <= 10):
                raise ValueError("Velocidade deve estar entre 1 e 10")
            if not (0 <= tonalidade <= 10):
                raise ValueError("Tonalidade deve estar entre 0 e 10")
            
            self.config_impressao.update({
                'temperatura': temperatura,
                'velocidade': velocidade,
                'tonalidade': tonalidade,
                'largura_etiqueta': largura,
                'altura_etiqueta': altura,
                'margem_esquerda': margem
            })
            
        except ValueError as e:
            messagebox.showwarning("Validação", f"Valor inválido: {e}")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao aplicar configurações: {e}")
    
    def resetar_configuracoes(self):
        default_config = {
            'temperatura': 10,
            'velocidade': 5,
            'tonalidade': 5,
            'largura_etiqueta': 550,
            'altura_etiqueta': 35,
            'margem_esquerda': 2,
            'margem_superior': 10
        }
        
        self.config_impressao.update(default_config)
        
        self.temp_var.set(str(default_config['temperatura']))
        self.speed_var.set(str(default_config['velocidade']))
        self.tone_var.set(str(default_config['tonalidade']))
        self.width_var.set(str(default_config['largura_etiqueta']))
        self.height_var.set(str(default_config['altura_etiqueta']))
        self.margin_var.set(str(default_config['margem_esquerda']))
        
        messagebox.showinfo("Reset", "Configurações resetadas para valores padrão!")
    
    def testar_configuracoes(self):
        try:
            self.aplicar_configuracoes()
            
            printer_name = self.printer_var.get()
            if not printer_name:
                messagebox.showwarning("Impressora", "Selecione uma impressora primeiro")
                return
            
            zpl_teste = self.gerar_codigo_zpl_lote(
                "TESTE123", 
                "16GB", 
                "T14", 
                "Autopilot", 
                "Excelente"
            )
            
            self.enviar_para_impressora(zpl_teste, printer_name)
            messagebox.showinfo("Teste", "Etiqueta de teste enviada para impressão!")
            
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao testar configurações: {e}")

    def buscar_patrimonio(self, event=None):
        patrimonio = self.patrimonio_var.get().strip()
        
        if patrimonio.startswith('0'):
            patrimonio = patrimonio.lstrip('0')
            self.patrimonio_var.set(patrimonio) 
            
        if not patrimonio:
            self.ocultar_alerta()
            self.habilitar_impressao(False)
            self.status_indicator.configure(text_color="#95A5A6")
            return
            
        try:
            self.cursor_patrimonios.execute('SELECT * FROM patrimonios WHERE LOWER(patrimonio) = LOWER(?)', (patrimonio,))
            resultado = self.cursor_patrimonios.fetchone()
            
            if resultado:
                self.ram_var.set(resultado[2])
                self.modelo_var.set(resultado[3])
                self.tipo_var.set(resultado[4])
                self.estado_var.set(resultado[5])
                
                if resultado[2] not in ["16GB", "32GB", "64GB"]:
                    self.ram_custom_var.set(resultado[2])
                    self.ram_var.set("")
                else:
                    self.ram_custom_var.set("")
                
                self.status_indicator.configure(text_color="#2AA876")
                self.ocultar_alerta()
                self.habilitar_impressao(True)
                self.patrimonio_encontrado = True
                
                self.atualizar_preview()
                
            else:
                self.status_indicator.configure(text_color="#E74C3C")
                self.mostrar_alerta()
                self.habilitar_impressao(False)
                self.patrimonio_encontrado = False
                self.preview_label.configure(text="Patrimônio não encontrado no banco")
                
        except sqlite3.Error as e:
            messagebox.showerror("Erro de Banco de Dados", f"Erro ao buscar patrimônio: {e}")
            self.habilitar_impressao(False)
            self.patrimonio_encontrado = False
    
    def atualizar_preview(self):
        patrimonio = self.patrimonio_var.get().strip()
        if patrimonio and self.patrimonio_encontrado:
            ram_value = self.ram_var.get()
            if ram_value == "":
                ram_value = self.ram_custom_var.get().strip()
            
            preview_text = f"""PATRIMONIO: {patrimonio}
MEMORIA: {ram_value}
MODELO: {self.modelo_var.get()}
TIPO: {self.tipo_var.get()}
ESTADO: {self.estado_var.get()}

[ CÓDIGO DE BARRAS ]"""
            self.preview_label.configure(text=preview_text)
    
    def mostrar_alerta(self):
        self.alerta_frame.pack(fill="x", padx=15, pady=5)
        self.alerta_label.pack(padx=15, pady=10)
    
    def ocultar_alerta(self):
        self.alerta_frame.pack_forget()
    
    def habilitar_impressao(self, habilitar):
        if habilitar:
            self.imprimir_btn.configure(state="normal", fg_color="#2AA876", hover_color="#207A59")
        else:
            self.imprimir_btn.configure(state="disabled", fg_color="#95A5A6", hover_color="#7F8C8D")
    
    def salvar_no_banco(self):
        patrimonio = self.patrimonio_var.get().strip()
        
        if patrimonio.startswith('0'):
            patrimonio = patrimonio.lstrip('0')
            self.patrimonio_var.set(patrimonio)
        
        ram_value = self.ram_var.get()
        if ram_value == "":
            ram_value = self.ram_custom_var.get().strip()
            if not ram_value:
                messagebox.showwarning("Campo Vazio", "Digite o valor da memória RAM")
                return
        
        modelo = self.modelo_var.get()
        tipo = self.tipo_var.get()
        estado = self.estado_var.get()
        
        if not patrimonio:
            messagebox.showwarning("Campo Vazio", "Digite o número do patrimônio")
            return
            
        if not modelo:
            messagebox.showwarning("Campo Vazio", "Selecione um modelo")
            return

        if tipo == "Não avaliado" or not tipo:
            messagebox.showwarning("Campo Vazio", "Por favor, selecione o Tipo de sistema (Autopilot, PXE, Linux ou MacOs) antes de salvar.")
            return

        modelo_upper = modelo.upper() 
        
        if "MAC" in modelo_upper:
            if tipo != "MacOs":
                messagebox.showerror("Erro de Validação", f"Equipamentos Apple ({modelo}) só podem ser cadastrados com o tipo 'MacOs'.")
                return
        else:
            if tipo == "MacOs":
                messagebox.showerror("Erro de Validação", f"O modelo '{modelo}' não é um dispositivo Apple e não pode ser cadastrado como 'MacOs'.")
                return
            
        try:
            self.cursor_patrimonios.execute("SELECT id, datetime(data_cadastro, 'localtime') FROM patrimonios WHERE patrimonio = ?", (patrimonio,))
            existe = self.cursor_patrimonios.fetchone()
            
            if existe:
                id_patrimonio, data_cadastro = existe
                data_str = "Data desconhecida"
                if data_cadastro:
                    try:
                        dt = datetime.datetime.strptime(data_cadastro, "%Y-%m-%d %H:%M:%S")
                        data_str = dt.strftime("%d/%m/%Y às %H:%M")
                    except ValueError:
                        data_str = data_cadastro 

                resposta = messagebox.askyesno(
                    "Patrimônio Já Cadastrado", 
                    f"A máquina com o patrimônio '{patrimonio}' já foi cadastrada no sistema em:\n\n{data_str}\n\nDeseja atualizar os dados desta máquina?"
                )
                
                if not resposta:
                    return
                
                self.cursor_patrimonios.execute('''
                    UPDATE patrimonios SET memoria = ?, modelo = ?, tipo = ?, estado = ?
                    WHERE patrimonio = ?
                ''', (ram_value, modelo, tipo, estado, patrimonio))
                mensagem_sucesso = "Dados da máquina atualizados com sucesso!"
                
            else:
                self.cursor_patrimonios.execute('''
                    INSERT INTO patrimonios (patrimonio, memoria, modelo, tipo, estado)
                    VALUES (?, ?, ?, ?, ?)
                ''', (patrimonio, ram_value, modelo, tipo, estado))
                mensagem_sucesso = "Nova máquina cadastrada com sucesso no banco de dados!"
                
            self.conn_patrimonios.commit()
            self.buscar_patrimonio()
            self.atualizar_estatisticas()
            messagebox.showinfo("Sucesso", mensagem_sucesso)
            
        except sqlite3.Error as e:
            messagebox.showerror("Erro de Banco de Dados", f"Erro ao salvar dados: {e}")
    
    def carregar_patrimonios_banco(self):
        try:
            self.cursor_patrimonios.execute('SELECT patrimonio, modelo, tipo, estado FROM patrimonios ORDER BY patrimonio')
            patrimonios = self.cursor_patrimonios.fetchall()
            
            if patrimonios:
                lista_patrimonios = [pat[0] for pat in patrimonios]
                self.patrimonios_text.delete(1.0, "end")
                self.patrimonios_text.insert(1.0, "\n".join(lista_patrimonios))
                
                total = len(patrimonios)
                status_text = f"✓ Carregados {total} patrimônios do banco"
                self.status_lote_label.configure(text=status_text)
            else:
                self.status_lote_label.configure(text="Nenhum patrimônio encontrado no banco")
                
        except sqlite3.Error as e:
            messagebox.showerror("Erro", f"Erro ao carregar patrimônios: {e}")
    
    def baixar_planilha_patrimonios(self):
        try:
            self.cursor_patrimonios.execute('SELECT patrimonio, modelo, tipo, estado FROM patrimonios ORDER BY patrimonio')
            patrimonios = self.cursor_patrimonios.fetchall()
            
            if not patrimonios:
                messagebox.showinfo("Informação", "Nenhum patrimônio encontrado no banco de dados")
                return
            
            import csv
            import io
            
            output = io.StringIO()
            writer = csv.writer(output)
            writer.writerow(['Patrimônio', 'Modelo', 'Tipo', 'Estado'])
            
            for patrimonio in patrimonios:
                writer.writerow(patrimonio)
            
            csv_content = output.getvalue()
            output.close()
            
            from tkinter import filedialog
            filename = filedialog.asksaveasfilename(
                defaultextension=".csv",
                filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
                initialfile="patrimonios_export.csv"
            )
            
            if filename:
                with open(filename, 'w', newline='', encoding='utf-8') as file:
                    file.write(csv_content)
                messagebox.showinfo("Sucesso", f"Planilha exportada com sucesso!\n{len(patrimonios)} patrimônios exportados.")
                
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao exportar planilha: {e}")
    
    def processar_lista_patrimonios(self):
        texto = self.patrimonios_text.get(1.0, "end").strip()
        if not texto:
            return []
        
        patrimonios = []
        for linha in texto.split('\n'):
            linha = linha.strip()
            if linha:
                if ',' in linha:
                    patrimonios.extend([pat.strip().lstrip('0') for pat in linha.split(',') if pat.strip()])
                else:
                    patrimonios.append(linha.lstrip('0'))
        
        return patrimonios
    
    def imprimir_em_lote(self):
        try:
            patrimonios = self.processar_lista_patrimonios()
            if not patrimonios:
                messagebox.showwarning("Lista Vazia", "Digite os patrimônios para impressão")
                return
            
            qtd_copias = int(self.qtd_lote_var.get() or "1")
            printer_name = self.printer_lote_var.get()
            
            if not printer_name:
                messagebox.showwarning("Impressora", "Selecione uma impressora")
                return
            
            self.status_lote_label.configure(text=f"🖨️ Imprimindo {len(patrimonios)} patrimônios...")
            self.imprimir_lote_btn.configure(state="disabled")
            self.update()
            
            sucessos = 0
            erros = []
            
            for patrimonio in patrimonios:
                patrimonio = patrimonio.strip()
                if patrimonio:
                    try:
                        self.cursor_patrimonios.execute('''
                            SELECT memoria, modelo, tipo, estado FROM patrimonios 
                            WHERE LOWER(patrimonio) = LOWER(?)
                        ''', (patrimonio,))
                        resultado = self.cursor_patrimonios.fetchone()
                        
                        if resultado:
                            ram, modelo, tipo, estado = resultado
                            for _ in range(qtd_copias):
                                zpl_code = self.gerar_codigo_zpl_lote(patrimonio, ram, modelo, tipo, estado)
                                self.enviar_para_impressora(zpl_code, printer_name)
                            
                            data_impressao = datetime.datetime.now().strftime("%Y-%m-%d")
                            self.cursor_patrimonios.execute('''
                                UPDATE patrimonios SET data_impressao = ? WHERE patrimonio = ?
                            ''', (data_impressao, patrimonio))
                            sucessos += 1
                        else:
                            erros.append(f"{patrimonio} - Não encontrado no banco")
                            
                    except Exception as e:
                        erros.append(f"{patrimonio} - {str(e)}")
            
            self.conn_patrimonios.commit()
            
            if erros:
                messagebox.showwarning("Impressão Concluída", 
                                     f"Impressão concluída!\n✓ Sucessos: {sucessos}\n✗ Erros: {len(erros)}\n\nErros:\n" + "\n".join(erros))
            else:
                messagebox.showinfo("Sucesso", f"✓ Todas as {sucessos} etiquetas impressas com sucesso!")
            
            self.status_lote_label.configure(text=f"✓ Concluído: {sucessos} sucessos, {len(erros)} erros")
            
        except Exception as e:
            messagebox.showerror("Erro", f"Erro na impressão em lote: {str(e)}")
        finally:
            self.imprimir_lote_btn.configure(state="normal")
    
    def imprimir_etiqueta_individual(self):
        if not self.patrimonio_encontrado:
            messagebox.showwarning("Patrimônio Não Encontrado", 
                                 "Não é possível imprimir. O patrimônio não foi encontrado no banco de dados.\n\n"
                                 "Salve o patrimônio no banco antes de imprimir.")
            return
            
        try:
            patrimonio = self.patrimonio_var.get().strip()
            
            if patrimonio.startswith('0'):
                patrimonio = patrimonio.lstrip('0')
                self.patrimonio_var.set(patrimonio)
            
            ram_value = self.ram_var.get()
            if ram_value == "":
                ram_value = self.ram_custom_var.get().strip()
            
            modelo = self.modelo_var.get()
            printer_name = self.printer_var.get()
            qtd_copias = int(self.qtd_var.get() or "1")
            self.status_indicator.configure(text_color="#F39C12")
            self.imprimir_btn.configure(state="disabled")
            self.update()
            
            for i in range(qtd_copias):
                zpl_code = self.gerar_codigo_zpl()
                self.enviar_para_impressora(zpl_code, printer_name)
            
            data_impressao = datetime.datetime.now().strftime("%Y-%m-%d ")
            self.cursor_patrimonios.execute('''
                UPDATE patrimonios SET data_impressao = ? WHERE patrimonio = ?
            ''', (data_impressao, patrimonio))
            self.conn_patrimonios.commit()
    
            self.status_indicator.configure(text_color="#2AA876")
            
        except Exception as e:
            self.status_indicator.configure(text_color="#E74C3C")
        finally:
            if self.patrimonio_encontrado:
                self.imprimir_btn.configure(state="normal")
    
    def enviar_para_impressora(self, zpl_code, printer_name):
        hprinter = win32print.OpenPrinter(printer_name)
        try:
            doc_info = ("Etiqueta", None, "RAW")
            job_id = win32print.StartDocPrinter(hprinter, 1, doc_info)
            win32print.StartPagePrinter(hprinter)
            win32print.WritePrinter(hprinter, zpl_code.encode('utf-8'))
            win32print.EndPagePrinter(hprinter)
            win32print.EndDocPrinter(hprinter)
        finally:
            win32print.ClosePrinter(hprinter)
    
    def gerar_codigo_zpl(self):
        patrimonio = self.patrimonio_var.get().strip()
        ram_value = self.ram_var.get()
        if ram_value == "":
            ram_value = self.ram_custom_var.get().strip()
        
        estado = self.estado_var.get()
        modelo = self.modelo_var.get()
        tipo = self.tipo_var.get()
        
        return self.gerar_codigo_zpl_lote(patrimonio, ram_value, modelo, tipo, estado)
    
    def gerar_codigo_zpl_lote(self, patrimonio, ram, modelo, tipo, estado):
        data_impressao = datetime.datetime.now().strftime("%d/%m/%Y ")
        
        temperatura = self.config_impressao.get('temperatura', 10)
        velocidade = self.config_impressao.get('velocidade', 5)
        tonalidade = self.config_impressao.get('tonalidade', 5)
        largura = self.config_impressao.get('largura_etiqueta', 550)
        altura = self.config_impressao.get('altura_etiqueta', 35)
        margem_esq = self.config_impressao.get('margem_esquerda', 2)
        margem_sup = self.config_impressao.get('margem_superior', 10)
        
        zpl_commands = []
        zpl_commands.append("^XA")
        zpl_commands.append("^MMT")  
        
        zpl_commands.append(f"^MD{temperatura}")  
        zpl_commands.append(f"^PR{velocidade}")   
        zpl_commands.append(f"^TO{tonalidade}")  
        
        zpl_commands.append(f"^PW{largura}")      
        zpl_commands.append(f"^LL{altura}")       
        
        pos_y1 = margem_sup + 14
        pos_y2 = margem_sup + 45
        pos_y3 = margem_sup + 75
        pos_x_col1 = margem_esq + 25
        pos_x_col2 = margem_esq + 325
        
        zpl_commands.append(f"^FO{pos_x_col1},{pos_y1}^A0N,25,25^FDPatrimonio: {patrimonio}^FS")
        zpl_commands.append(f"^FO{pos_x_col1},{pos_y2}^A0N,25,25^FDMem: {ram}^FS")
        zpl_commands.append(f"^FO{pos_x_col2},{pos_y1}^A0N,25,25^FDEstado: {estado}^FS")
        zpl_commands.append(f"^FO{pos_x_col1},{pos_y3}^A0N,25,25^FDModelo: {modelo}^FS")
        zpl_commands.append(f"^FO{pos_x_col2},{pos_y2}^A0N,25,25^FDTipo: {tipo}^FS")
        zpl_commands.append(f"^FO335,85^A0N,25,25^FDData:{data_impressao}^FS")
        zpl_commands.append(f"^FO35,120^BY2^BCN,60,Y,N,N^FD{patrimonio}^FS")
        zpl_commands.append("^PQ1")
        zpl_commands.append("^XZ")
        
        return "\n".join(zpl_commands)
    
    def limpar_campos(self):
        self.patrimonio_var.set("")
        self.ram_var.set("16GB")
        self.ram_custom_var.set("")
        self.modelo_var.set("")
        self.tipo_var.set("Não avaliado")
        self.estado_var.set("Excelente")
        self.qtd_var.set("1")
        self.status_indicator.configure(text_color="#95A5A6")
        self.ocultar_alerta()
        self.habilitar_impressao(False)
        self.patrimonio_encontrado = False
        self.preview_label.configure(text="Pré-visualização da etiqueta aparecerá aqui")
        self.patrimonio_entry.focus_set()
    
    def limpar_lista_patrimonios(self):
        self.patrimonios_text.delete(1.0, "end")
        self.status_lote_label.configure(text="Lista limpa. Digite os patrimônios para impressão.")

    def detectar_impressoras_zebra(self):
        try:
            impressoras = win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL)
            printer_names = [printer[2] for printer in impressoras]
            
            zebra_keywords = ["ZEBRA", "ZDESIGNER", "GC420", "GK420", "GX420", "ZD620", "ZD420", "ZT230", "ZT410"]
            
            self.impressoras_zebra = []
            self.impressora_detectada = None
            
            for printer_name in printer_names:
                printer_upper = printer_name.upper()
                if any(keyword in printer_upper for keyword in zebra_keywords):
                    self.impressoras_zebra.append(printer_name)
                    if not self.impressora_detectada:
                        self.impressora_detectada = printer_name
                        if "ZD620" in printer_upper or "ZD420" in printer_upper:
                            self.config_impressao.update({'temperatura': 12, 'velocidade': 6})
                        elif "ZT230" in printer_upper or "ZT410" in printer_upper:
                            self.config_impressao.update({'temperatura': 15, 'velocidade': 8})
                        else:  
                            self.config_impressao.update({'temperatura': 10, 'velocidade': 5})
            
            return self.impressoras_zebra
            
        except Exception as e:
            print(f"Erro ao detectar impressoras Zebra: {e}")
            return []
    
    def carregar_impressoras(self):
        try:
            impressoras = win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL)
            printer_names = [printer[2] for printer in impressoras]
            
            if hasattr(self, 'impressoras_zebra') and self.impressoras_zebra:
                self.printer_dropdown.configure(values=self.impressoras_zebra)
                self.printer_var.set(self.impressora_detectada or self.impressoras_zebra[0])
                if hasattr(self, 'printer_lote_dropdown'):
                    self.printer_lote_dropdown.configure(values=self.impressoras_zebra)
                    self.printer_lote_var.set(self.impressora_detectada or self.impressoras_zebra[0])
            elif printer_names:
                self.printer_dropdown.configure(values=printer_names)
                self.printer_var.set(printer_names[0])
                if hasattr(self, 'printer_lote_dropdown'):
                    self.printer_lote_dropdown.configure(values=printer_names)
                    self.printer_lote_var.set(printer_names[0])
                
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao carregar impressoras: {e}")

    def __del__(self):
        if hasattr(self, 'conn_patrimonios'):
            self.conn_patrimonios.close()
        if hasattr(self, 'conn_modelos'):
            self.conn_modelos.close()

if __name__ == "__main__":
    app = EtiquetaSimplificadaApp()
    app.mainloop()