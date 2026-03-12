import customtkinter as ctk
import win32print
import win32ui
import sqlite3
import datetime
from tkinter import messagebox

# Configurar aparência
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

class ManutencaoApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        self.title("Manutenção & Garantia")
        self.geometry("800x750")
        # Conectar ao banco de dados de manutenção
        self.conectar_banco()
        
        self.criar_widgets()
        self.carregar_impressoras()
        self.carregar_tipos_equipamento()
        
    def conectar_banco(self):
        """Conectar ao banco de dados SQLite para manutenção"""
        try:
            self.conn = sqlite3.connect('manutencoes.db')
            self.cursor = self.conn.cursor()
            
            # Criar tabela de manutenções se não existir
            self.cursor.execute('''
                CREATE TABLE IF NOT EXISTS manutencoes (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    patrimonio TEXT NOT NULL,
                    tipo_manutencao TEXT NOT NULL,
                    descricao TEXT NOT NULL,
                    tipo_equipamento TEXT NOT NULL,
                    data_cadastro TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    data_impressao TIMESTAMP
                )
            ''')
            
            # Criar tabela de tipos de equipamento se não existir
            self.cursor.execute('''
                CREATE TABLE IF NOT EXISTS tipos_equipamento (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    nome TEXT NOT NULL UNIQUE
                )
            ''')
            
            # Inserir tipos padrão se a tabela estiver vazia
            self.cursor.execute("SELECT COUNT(*) FROM tipos_equipamento")
            if self.cursor.fetchone()[0] == 0:
                tipos_padrao = ['Autopilot', 'Desktop', 'Notebook', 'Monitor', 'Impressora']
                for tipo in tipos_padrao:
                    self.cursor.execute("INSERT INTO tipos_equipamento (nome) VALUES (?)", (tipo,))
            
            self.conn.commit()
            
        except sqlite3.Error as e:
            messagebox.showerror("Erro de Banco de Dados", f"Erro ao conectar com o banco: {e}")
    
    def criar_widgets(self):
        # Frame principal
        main_frame = ctk.CTkFrame(self)
        main_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Título
        title_label = ctk.CTkLabel(main_frame, text="Impressão de Etiquetas para Manutenção e Garantia", 
                                  font=ctk.CTkFont(size=16, weight="bold"))
        title_label.pack(pady=10)
        
        # Campo do patrimônio
        patrimonio_frame = ctk.CTkFrame(main_frame)
        patrimonio_frame.pack(fill="x", padx=10, pady=5)
        
        ctk.CTkLabel(patrimonio_frame, text="Patrimônio:").pack(anchor="w", padx=10)
        self.patrimonio_var = ctk.StringVar()
        self.patrimonio_entry = ctk.CTkEntry(patrimonio_frame, textvariable=self.patrimonio_var,
                                           placeholder_text="Digite ou escaneie o patrimônio")
        self.patrimonio_entry.pack(pady=5, fill="x", padx=10)
        self.patrimonio_entry.bind("<Return>", self.buscar_manutencao)
        self.patrimonio_entry.bind("<FocusOut>", self.buscar_manutencao)
        
        # Frame para tipo de manutenção
        manutencao_frame = ctk.CTkFrame(main_frame)
        manutencao_frame.pack(fill="x", padx=10, pady=10)
        
        ctk.CTkLabel(manutencao_frame, text="Tipo de Manutenção:").pack(anchor="w", padx=10)
        self.tipo_manutencao_var = ctk.StringVar(value="Preventiva")
        ctk.CTkRadioButton(manutencao_frame, text="Preventiva", variable=self.tipo_manutencao_var, value="Preventiva").pack(anchor="w", padx=10, pady=5)
        ctk.CTkRadioButton(manutencao_frame, text="Corretiva", variable=self.tipo_manutencao_var, value="Corretiva").pack(anchor="w", padx=10)
        ctk.CTkRadioButton(manutencao_frame, text="Garantia", variable=self.tipo_manutencao_var, value="Garantia").pack(anchor="w", padx=10, pady=5)
        
        # Frame para tipo de equipamento
        equipamento_frame = ctk.CTkFrame(main_frame)
        equipamento_frame.pack(fill="x", padx=10, pady=5)
        
        ctk.CTkLabel(equipamento_frame, text="Tipo de Equipamento:").pack(anchor="w", padx=10)
        
        # Frame para combobox e botões de gerenciamento
        equipamento_controls_frame = ctk.CTkFrame(equipamento_frame)
        equipamento_controls_frame.pack(fill="x", pady=5, padx=10)
        
        self.tipo_equipamento_var = ctk.StringVar(value="")
        self.tipo_equipamento_cb = ctk.CTkComboBox(equipamento_controls_frame, 
                                                  variable=self.tipo_equipamento_var,
                                                  state="readonly")
        self.tipo_equipamento_cb.pack(side="left", fill="x", expand=True, padx=(0, 5))
        
        # Botões para gerenciar tipos de equipamento
        gerenciar_frame = ctk.CTkFrame(equipamento_controls_frame)
        gerenciar_frame.pack(side="right")
        
        self.adicionar_tipo_btn = ctk.CTkButton(gerenciar_frame, text="+", 
                                               width=30, height=30,
                                               command=self.adicionar_tipo_equipamento)
        self.adicionar_tipo_btn.pack(side="left", padx=2)
        
        self.remover_tipo_btn = ctk.CTkButton(gerenciar_frame, text="-", 
                                             width=30, height=30,
                                             command=self.remover_tipo_equipamento,
                                             fg_color="#E74C3C", hover_color="#C0392B")
        self.remover_tipo_btn.pack(side="left", padx=2)
        
        # Campo para descrição do problema
        descricao_frame = ctk.CTkFrame(main_frame)
        descricao_frame.pack(fill="x", padx=10, pady=10)
        
        ctk.CTkLabel(descricao_frame, text="Descrição do Problema:").pack(anchor="w", padx=10)
        self.descricao_text = ctk.CTkTextbox(descricao_frame, height=80)
        self.descricao_text.pack(pady=5, fill="x", padx=10)
        
        # Impressora
        printer_frame = ctk.CTkFrame(main_frame)
        printer_frame.pack(fill="x", padx=10, pady=10)
        
        ctk.CTkLabel(printer_frame, text="Impressora:").pack(anchor="w", padx=10)
        self.printer_var = ctk.StringVar(value="")
        self.printer_dropdown = ctk.CTkComboBox(printer_frame, variable=self.printer_var)
        self.printer_dropdown.pack(pady=5, fill="x", padx=10)
        
        # Botões
        button_frame = ctk.CTkFrame(main_frame)
        button_frame.pack(pady=15)
        
        self.imprimir_btn = ctk.CTkButton(button_frame, text="Imprimir Etiqueta", 
                                         command=self.imprimir_etiqueta,
                                         fg_color="#2AA876", hover_color="#207A59")
        self.imprimir_btn.pack(side="left", padx=5)
        
        self.limpar_btn = ctk.CTkButton(button_frame, text="Limpar", 
                                       command=self.limpar_campos,
                                       fg_color="#E74C3C", hover_color="#C0392B")
        self.limpar_btn.pack(side="left", padx=5)
        
        self.salvar_btn = ctk.CTkButton(button_frame, text="Salvar", 
                                       command=self.salvar_no_banco,
                                       fg_color="#3498DB", hover_color="#2980B9")
        self.salvar_btn.pack(side="left", padx=5)
        
        # Status
        self.status_label = ctk.CTkLabel(main_frame, text="Pronto para imprimir")
        self.status_label.pack(pady=5)
        
        # Informações do banco
        info_frame = ctk.CTkFrame(main_frame)
        info_frame.pack(fill="x", padx=10, pady=5)
        
        ctk.CTkLabel(info_frame, text="Informações do Banco:", font=ctk.CTkFont(weight="bold")).pack(anchor="w", pady=5, padx=10)
        self.info_label = ctk.CTkLabel(info_frame, text="Nenhum registro de manutenção carregado")
        self.info_label.pack(anchor="w", padx=10)
        
        # Foco no campo do patrimônio
        self.patrimonio_entry.focus_set()
        
    def carregar_tipos_equipamento(self):
        """Carrega os tipos de equipamento do banco de dados"""
        try:
            self.cursor.execute("SELECT nome FROM tipos_equipamento ORDER BY nome")
            resultados = self.cursor.fetchall()
            tipos = [resultado[0] for resultado in resultados]
            
            if tipos:
                self.tipo_equipamento_cb.configure(values=tipos)
                self.tipo_equipamento_var.set(tipos[0])
            else:
                self.tipo_equipamento_cb.configure(values=[])
                self.tipo_equipamento_var.set("")
                
        except sqlite3.Error as e:
            messagebox.showerror("Erro", f"Erro ao carregar tipos de equipamento: {e}")
    
    def adicionar_tipo_equipamento(self):
        """Abre diálogo para adicionar novo tipo de equipamento"""
        dialog = ctk.CTkInputDialog(text="Digite o novo tipo de equipamento:", title="Adicionar Tipo")
        novo_tipo = dialog.get_input()
        
        if novo_tipo and novo_tipo.strip():
            novo_tipo = novo_tipo.strip()
            try:
                self.cursor.execute("INSERT INTO tipos_equipamento (nome) VALUES (?)", (novo_tipo,))
                self.conn.commit()
                self.carregar_tipos_equipamento()
                self.status_label.configure(text=f"Tipo '{novo_tipo}' adicionado com sucesso!")
            except sqlite3.IntegrityError:
                messagebox.showerror("Erro", f"O tipo '{novo_tipo}' já existe!")
            except sqlite3.Error as e:
                messagebox.showerror("Erro", f"Erro ao adicionar tipo: {e}")
    
    def remover_tipo_equipamento(self):
        """Remove o tipo de equipamento selecionado"""
        tipo_selecionado = self.tipo_equipamento_var.get()
        
        if not tipo_selecionado:
            messagebox.showwarning("Aviso", "Nenhum tipo selecionado para remover")
            return
            
        confirmacao = messagebox.askyesno(
            "Confirmar Remoção", 
            f"Tem certeza que deseja remover o tipo '{tipo_selecionado}'?"
        )
        
        if confirmacao:
            try:
                self.cursor.execute("DELETE FROM tipos_equipamento WHERE nome = ?", (tipo_selecionado,))
                self.conn.commit()
                self.carregar_tipos_equipamento()
                self.status_label.configure(text=f"Tipo '{tipo_selecionado}' removido com sucesso!")
            except sqlite3.Error as e:
                messagebox.showerror("Erro", f"Erro ao remover tipo: {e}")
        
    def buscar_manutencao(self, event=None):
        """Busca um registro de manutenção no banco de dados"""
        patrimonio = self.patrimonio_var.get().strip()
        
        if not patrimonio:
            return
            
        try:
            self.cursor.execute("SELECT * FROM manutencoes WHERE patrimonio = ? ORDER BY data_cadastro DESC LIMIT 1", (patrimonio,))
            resultado = self.cursor.fetchone()
            
            if resultado:
                # Preencher os campos com os dados do banco
                self.tipo_manutencao_var.set(resultado[2])  # tipo de manutenção
                self.tipo_equipamento_var.set(resultado[4])  # tipo de equipamento
                self.descricao_text.delete("1.0", "end")
                self.descricao_text.insert("1.0", resultado[3])  # descrição
                
                data_cadastro = resultado[5]
                data_impressao = resultado[6] if resultado[6] else "Nunca"
                
                self.info_label.configure(text=f"Registro encontrado! Cadastrado em: {data_cadastro}, Última impressão: {data_impressao}")
                self.status_label.configure(text="Registro carregado do banco de dados")
            else:
                self.info_label.configure(text="Patrimônio não encontrado no banco. Preencha os dados.")
                self.status_label.configure(text="Novo registro - preencha os dados")
                
        except sqlite3.Error as e:
            messagebox.showerror("Erro de Banco de Dados", f"Erro ao buscar registro: {e}")
    
    def salvar_no_banco(self):
        """Salva ou atualiza as informações no banco de dados"""
        patrimonio = self.patrimonio_var.get().strip()
        tipo_manutencao = self.tipo_manutencao_var.get()
        descricao = self.descricao_text.get("1.0", "end-1c").strip()
        tipo_equipamento = self.tipo_equipamento_var.get()
        
        if not patrimonio:
            messagebox.showwarning("Campo Vazio", "Digite o número do patrimônio")
            return
            
        if not tipo_equipamento:
            messagebox.showwarning("Campo Vazio", "Selecione um tipo de equipamento")
            return
            
        if not descricao:
            messagebox.showwarning("Campo Vazio", "Digite a descrição do problema")
            return
            
        try:
            # Inserir novo registro (mantemos histórico)
            self.cursor.execute('''
                INSERT INTO manutencoes (patrimonio, tipo_manutencao, descricao, tipo_equipamento)
                VALUES (?, ?, ?, ?)
            ''', (patrimonio, tipo_manutencao, descricao, tipo_equipamento))
                
            self.conn.commit()
            self.status_label.configure(text="Registro de manutenção salvo no banco de dados")
            messagebox.showinfo("Sucesso", "Dados salvos com sucesso no banco de dados!")
            
        except sqlite3.Error as e:
            messagebox.showerror("Erro de Banco de Dados", f"Erro ao salvar dados: {e}")
    
    def carregar_impressoras(self):
        try:
            impressoras = win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL)
            printer_names = [printer[2] for printer in impressoras]
            
            zebra_printers = [name for name in printer_names if "ZEBRA" in name.upper() or "ZDESIGNER" in name.upper() or "GC420" in name.upper()]
            
            if zebra_printers:
                self.printer_dropdown.configure(values=zebra_printers)
                self.printer_var.set(zebra_printers[0])
            elif printer_names:
                self.printer_dropdown.configure(values=printer_names)
                self.printer_var.set(printer_names[0])
                
        except Exception as e:
            self.status_label.configure(text="Erro ao carregar impressoras")
    
    def gerar_codigo_zpl(self):
        """Gera código ZPL para etiqueta de manutenção (coordenadas mantidas)"""
        patrimonio = self.patrimonio_var.get().strip()
        tipo_manutencao = self.tipo_manutencao_var.get()
        tipo_equipamento = self.tipo_equipamento_var.get()
        descricao = self.descricao_text.get("1.0", "end-1c").strip()
        
        if not patrimonio:
            raise ValueError("Patrimônio é obrigatório")
        
        if not descricao:
            raise ValueError("Descrição do problema é obrigatória")
        
        zpl_commands = []
        zpl_commands.append("^XA")
        zpl_commands.append("^MMT")  # Modo Tear-off
        
        # Primeira linha: MANUTENÇÃO e Tipo de Equipamento
        zpl_commands.append(f"^FO80,30^A0N,30,30^FDManutencao^FS")
        zpl_commands.append(f"^FO300,30^A0N,30,30^FDEq: {tipo_equipamento}^FS")
        
        # Segunda linha: Tipo de Manutenção
        zpl_commands.append(f"^FO80,70^A0N,30,30^FDTipo: {tipo_manutencao}^FS")
        
        # Terceira linha: Descrição do problema (primeira parte)
        zpl_commands.append(f"^FO80,110^A0N,25,25^FDProblema:^FS")
        
        # Dividir a descrição em até 3 linhas
        descricao_linhas = []
        palavras = descricao.split()
        linha_atual = ""
        
        for palavra in palavras:
            if len(linha_atual) + len(palavra) <= 45:
                if linha_atual:
                    linha_atual += " " + palavra
                else:
                    linha_atual = palavra
            else:
                descricao_linhas.append(linha_atual)
                linha_atual = palavra
        
        if linha_atual:
            descricao_linhas.append(linha_atual)
        
        # Limitar a 3 linhas máximo
        descricao_linhas = descricao_linhas[:2]
        
        # Adicionar as linhas de descrição
        for i, linha in enumerate(descricao_linhas):
            y_pos = 140 + (i * 30)
            zpl_commands.append(f"^FO80,{y_pos}^A0N,25,25^FD{linha}^FS")
        
        # Finalizar
        zpl_commands.append("^PQ1")  # Apenas 1 cópia
        zpl_commands.append("^XZ")
        
        return "\n".join(zpl_commands)
    
    def imprimir_etiqueta(self):
        try:
            patrimonio = self.patrimonio_var.get().strip()
            if not patrimonio:
                self.status_label.configure(text="Digite o número do patrimônio")
                return
                
            tipo_equipamento = self.tipo_equipamento_var.get()
            if not tipo_equipamento:
                self.status_label.configure(text="Selecione um tipo de equipamento")
                return
                
            descricao = self.descricao_text.get("1.0", "end-1c").strip()
            if not descricao:
                self.status_label.configure(text="Digite a descrição do problema")
                return
                
            printer_name = self.printer_var.get()
            if not printer_name:
                self.status_label.configure(text="Selecione uma impressora")
                return
            
            self.status_label.configure(text="Imprimindo...")
            self.imprimir_btn.configure(state="disabled")
            self.update()
            
            # Gerar código ZPL
            zpl_code = self.gerar_codigo_zpl()
            
            # Obter handle da impressora
            hprinter = win32print.OpenPrinter(printer_name)
            try:
                # Iniciar documento de impressão
                doc_info = ("Etiqueta de Manutenção", None, "RAW")
                job_id = win32print.StartDocPrinter(hprinter, 1, doc_info)
                win32print.StartPagePrinter(hprinter)
                
                # Enviar comandos ZPL para a impressora
                win32print.WritePrinter(hprinter, zpl_code.encode('utf-8'))
                
                # Finalizar impressão
                win32print.EndPagePrinter(hprinter)
                win32print.EndDocPrinter(hprinter)
                
                # Atualizar data de impressão no banco de dados
                data_impressao = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                self.cursor.execute('''
                    UPDATE manutencoes 
                    SET data_impressao = ?
                    WHERE patrimonio = ? 
                    AND id = (SELECT MAX(id) FROM manutencoes WHERE patrimonio = ?)
                ''', (data_impressao, patrimonio, patrimonio))
                self.conn.commit()
                
                self.status_label.configure(text="Etiqueta de manutenção impressa com sucesso!")
                
            except Exception as e:
                self.status_label.configure(text=f"Erro: {str(e)}")
            finally:
                win32print.ClosePrinter(hprinter)
                self.imprimir_btn.configure(state="normal")
                
        except Exception as e:
            self.status_label.configure(text=f"Erro: {str(e)}")
            self.imprimir_btn.configure(state="normal")
    
    def limpar_campos(self):
        """Limpa todos os campos"""
        self.patrimonio_var.set("")
        self.tipo_manutencao_var.set("Preventiva")
        self.tipo_equipamento_var.set("")
        self.descricao_text.delete("1.0", "end")
        self.info_label.configure(text="Nenhum registro de manutenção carregado")
        self.patrimonio_entry.focus_set()
        self.status_label.configure(text="Campos limpos. Pronto para imprimir.")

    def __del__(self):
        """Fechar conexão com o banco ao destruir a aplicação"""
        if hasattr(self, 'conn'):
            self.conn.close()


if __name__ == "__main__":
    app = ManutencaoApp()
    app.mainloop()