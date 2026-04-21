import sys
import os
import pandas as pd
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog, filedialog
import sqlite3
from datetime import datetime
from reportlab.pdfgen import canvas
from reportlab.lib import colors

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

class SistemaExpedicaoMaster:
    def __init__(self, root):
        self.root = root
        self.root.title("BURDOG - Gestão de Fluxo de Expedição")
        self.root.geometry("1350x850")
        
        try:
            self.root.iconbitmap(resource_path("logo_prolav.ico"))
        except: pass
        
        # Cores e Estilos
        self.cor_fundo = "#1e272e"
        self.cor_card = "#2f3640"
        self.cor_destaque = "#3498db"
        self.cor_sucesso = "#27ae60"
        self.cor_alerta = "#e67e22"
        self.f_btn = ("Arial", 9, "bold")
        
        self.root.configure(bg=self.cor_fundo)
        self.configurar_estilos()
        self.criar_banco()
        
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(expand=True, fill='both', padx=5, pady=5)

        # --- RODAPÉ GLOBAL DE EXPEDIENTE ---
        self.frame_exp = tk.Frame(self.root, bg=self.cor_card, bd=1, relief="ridge")
        self.frame_exp.pack(side="bottom", fill="x", ipady=3)

        self.btn_ini_exp = tk.Button(self.frame_exp, text="▶ INICIAR EXPEDIENTE", bg=self.cor_sucesso, 
                                     fg="white", font=("Arial", 8, "bold"), command=self.registrar_inicio_exp)
        self.btn_ini_exp.pack(side="left", padx=10, pady=5)

        self.btn_fim_exp = tk.Button(self.frame_exp, text="⏹ ENCERRAR EXPEDIENTE", bg="#c0392b", 
                                     fg="white", font=("Arial", 8, "bold"), command=self.registrar_fim_exp)
        self.btn_fim_exp.pack(side="left", padx=5, pady=5)

        self.lbl_exp_status = tk.Label(self.frame_exp, text="STATUS: AGUARDANDO", 
                                       fg="yellow", bg=self.cor_card, font=("Arial", 8, "bold"))
        self.lbl_exp_status.pack(side="right", padx=15)

        self.btn_editar_global = tk.Button(self.frame_exp, text="📝 EDITAR PEDIDO", bg="#3498db", 
                                   fg="white", font=("Arial", 8, "bold"), command=self.abrir_edicao)
        self.btn_editar_global.pack(side="left", padx=5, pady=5)
        
        abas_config = [
            ("aba_cadastro", " CADASTRO "), ("aba_pendentes", " PENDENTES "),
            ("aba_em_sep", " EM SEPARAÇÃO "), ("aba_aguard_conf", " AGUARD. CONFERÊNCIA "),
            ("aba_em_conf", " EM CONFERÊNCIA "), ("aba_aguard_pecas", " AGUARD. PEÇAS "), 
            ("aba_prontos", " PRONTOS "), ("aba_historico", " HISTÓRICO "), 
            ("aba_pesquisa", " PESQUISA "), ("aba_config", " EQUIPE ")
        ]
        
        for var_nome, titulo in abas_config:
            aba = tk.Frame(self.notebook, bg=self.cor_fundo)
            self.notebook.add(aba, text=titulo)
            setattr(self, var_nome, aba)

        self.setup_aba_cadastro()
        self.setup_aba_pendentes()
        self.setup_aba_em_sep()
        self.setup_aba_aguard_conf()
        self.setup_aba_em_conf()
        self.setup_aba_aguard_pecas()
        self.setup_aba_prontos()
        self.setup_aba_historico()
        self.setup_aba_pesquisa()
        self.setup_aba_config()
        
        self.atualizar_todas_tabelas()
        self.controlar_acesso(liberar=False)

    def criar_banco(self):
        conn = sqlite3.connect('expedicao_v4.db'); c = conn.cursor()
        c.execute('''CREATE TABLE IF NOT EXISTS pedidos
                     (id INTEGER PRIMARY KEY AUTOINCREMENT, numero_pedido TEXT, cliente TEXT, 
                      cidade TEXT, transporte TEXT, volumes TEXT, nf TEXT,
                      separador TEXT, conferente TEXT, 
                      h_entrada TEXT, h_inicio_sep TEXT, h_fim_sep TEXT, 
                      h_inicio_conf TEXT, h_pronto TEXT, 
                      h_inicio_pecas TEXT, h_fim_pecas TEXT, h_arquivado TEXT,
                      status TEXT, observacoes TEXT)''')
        
        c.execute('''CREATE TABLE IF NOT EXISTS itens_pedido 
                     (id INTEGER PRIMARY KEY AUTOINCREMENT, id_pedido INTEGER, 
                      codigo_produto TEXT, descricao_produto TEXT, quantidade TEXT)''')
        
        c.execute('''CREATE TABLE IF NOT EXISTS funcionarios (id INTEGER PRIMARY KEY AUTOINCREMENT, nome TEXT, cargo TEXT)''')

        c.execute('''CREATE TABLE IF NOT EXISTS registro_expediente 
                     (id INTEGER PRIMARY KEY AUTOINCREMENT, tipo TEXT, data_hora TEXT)''')
        conn.commit(); conn.close()

    def controlar_acesso(self, liberar=False):
        estado = "normal" if liberar else "disabled"
        
        # Sua lista de abas corrigida
        abas = [self.aba_cadastro, self.aba_pendentes, self.aba_em_sep, 
                self.aba_aguard_conf, self.aba_em_conf, self.aba_aguard_pecas, 
                self.aba_prontos, self.aba_historico, self.aba_pesquisa]
        
        for aba in abas:
            # Pegamos todos os widgets dentro da aba, incluindo os que estão dentro de Frames
            for widget in aba.winfo_children():
                self._aplicar_estado(widget, estado)

    def _aplicar_estado(self, widget, estado):
        """Função auxiliar para aplicar o estado apenas onde é permitido"""
        # Lista de tipos de componentes que ACEITAM o bloqueio
        tipos_bloqueaveis = (tk.Button, tk.Entry, ttk.Combobox, ttk.Treeview, tk.Text, tk.Checkbutton)
        
        if isinstance(widget, tipos_bloqueaveis):
            try:
                widget.configure(state=estado)
            except:
                pass
        
        # Se for um Frame, temos que entrar nele e bloquear o que tem dentro (recursividade)
        elif isinstance(widget, (tk.Frame, tk.LabelFrame)):
            for child in widget.winfo_children():
                self._aplicar_estado(child, estado)

    def registrar_inicio_exp(self):
        agora = datetime.now().strftime("%d/%m/%Y %H:%M")
        conn = sqlite3.connect('expedicao_v4.db')
        c = conn.cursor()
        c.execute("INSERT INTO registro_expediente (tipo, data_hora) VALUES (?,?)", ("INICIO", agora))
        conn.commit()
        conn.close()
        
        self.lbl_exp_status.config(text=f"EXPEDIENTE INICIADO ÀS {agora[11:]}", fg="#2ecc71")
        
        # --- AQUI ESTÁ A MELHORIA ---
        self.controlar_acesso(liberar=True) # Destrava o sistema
        # ----------------------------
        
        messagebox.showinfo("Burdog", "Expediente retomado! Os cronômetros voltaram a contar.")

    def registrar_fim_exp(self):
        if not messagebox.askyesno("Encerrar", "Deseja realmente encerrar o expediente e pausar os tempos?"):
            return
            
        agora = datetime.now().strftime("%d/%m/%Y %H:%M")
        conn = sqlite3.connect('expedicao_v4.db')
        c = conn.cursor()
        c.execute("INSERT INTO registro_expediente (tipo, data_hora) VALUES (?,?)", ("FIM", agora))
        conn.commit()
        conn.close()
        
        self.lbl_exp_status.config(text=f"EXPEDIENTE ENCERRADO ÀS {agora[11:]}", fg="#e74c3c")
        
        # --- AQUI ESTÁ A MELHORIA ---
        self.controlar_acesso(liberar=False) # Trava o sistema novamente
        # ----------------------------
        
        messagebox.showwarning("Burdog", "Expediente pausado! Bom descanso.")  

    def calcular_diferenca(self, inicio, fim):
        if not inicio or not fim or str(inicio).strip() in ["", "--", "None"]:
            return "--"
        try:
            fmt = "%d/%m/%Y %H:%M"
            d_ini = datetime.strptime(str(inicio).strip(), fmt)
            d_fim = datetime.strptime(str(fim).strip(), fmt)
            
            segundos_totais = int((d_fim - d_ini).total_seconds())
            
            # Busca no banco se houve pausas de expediente entre o inicio e fim do pedido
            conn = sqlite3.connect('expedicao_v4.db')
            c = conn.cursor()
            c.execute("SELECT tipo, data_hora FROM registro_expediente WHERE data_hora >= ? AND data_hora <= ? ORDER BY data_hora", (inicio, fim))
            registros = c.fetchall()
            conn.close()
            
            tempo_pausado = 0
            pausa_inicio = None
            
            for tipo, dt_hr in registros:
                dt = datetime.strptime(dt_hr, fmt)
                if tipo == "FIM":
                    pausa_inicio = dt
                elif tipo == "INICIO" and pausa_inicio:
                    tempo_pausado += int((dt - pausa_inicio).total_seconds())
                    pausa_inicio = None
            
            segundos_finais = segundos_totais - tempo_pausado
            if segundos_finais < 0: segundos_finais = 0
            
            minutos = segundos_finais // 60
            if minutos < 60: return f"{minutos} min"
            return f"{minutos // 60}h {minutos % 60}m"
        except:
            return "--"

    # --- ABA CADASTRO ---
    def setup_aba_cadastro(self):
        frame = tk.Frame(self.aba_cadastro, bg=self.cor_card, bd=1, relief="solid")
        frame.place(relx=0.5, rely=0.5, anchor="center", width=500, height=450)
        tk.Label(frame, text="NOVO PEDIDO", fg=self.cor_destaque, bg=self.cor_card, font=("Arial", 16, "bold")).pack(pady=15)
        self.ent_ped = self.add_entry(frame, "Número do Pedido:")
        self.ent_cli = self.add_entry(frame, "Nome do Cliente:")
        self.ent_cid = self.add_entry(frame, "Cidade:")
        tk.Label(frame, text="Meio de Transporte:", fg="white", bg=self.cor_card).pack()
        self.cb_trans = ttk.Combobox(frame, values=["Transportadora", "Retira", "Nosso Carro"], state="readonly")
        self.cb_trans.pack(pady=5, ipady=3)
        tk.Button(frame, text="CADASTRAR", bg=self.cor_sucesso, fg="white", font=self.f_btn, command=self.acao_cadastrar).pack(pady=25, ipadx=40, ipady=5)

    def acao_cadastrar(self):
        agora = datetime.now().strftime("%d/%m/%Y %H:%M")
        ped, cli, cid, tra = self.ent_ped.get().strip(), self.ent_cli.get().strip(), self.ent_cid.get().strip(), self.cb_trans.get()
        if not ped or not cli:
            messagebox.showwarning("Atenção", "Número do Pedido e Cliente são obrigatórios."); return
        conn = sqlite3.connect('expedicao_v4.db'); c = conn.cursor()
        c.execute("INSERT INTO pedidos (numero_pedido, cliente, cidade, transporte, h_entrada, status) VALUES (?,?,?,?,?,?)", (ped, cli, cid, tra, agora, "PENDENTE"))
        conn.commit(); conn.close()
        messagebox.showinfo("Sucesso", f"Pedido #{ped} cadastrado!"); self.ent_ped.delete(0, 'end'); self.ent_cli.delete(0, 'end'); self.ent_cid.delete(0, 'end'); self.cb_trans.set(''); self.ent_ped.focus_set(); self.atualizar_todas_tabelas()

    # --- ABA PENDENTES ---
    def setup_aba_pendentes(self):
        self.tree_pend = self.criar_tabela(self.aba_pendentes, ("ID", "Pedido", "Cliente", "Entrada", "Transporte"))
        tk.Button(self.aba_pendentes, text="INICIAR SEPARAÇÃO", bg=self.cor_destaque, fg="white", font=self.f_btn, command=self.dlg_iniciar_sep).pack(pady=10)
        self.tree_pend.bind("<Double-1>", self.abrir_detalhes)
        self.lbl_total_pend = tk.Label(self.aba_pendentes, text="Total de pedidos: 0", 
                               fg=self.cor_destaque, bg=self.cor_fundo, font=("Arial", 10, "bold"))
        self.lbl_total_pend.pack(side="bottom", anchor="e", padx=10, pady=5)

    def dlg_iniciar_sep(self):
        sel = self.tree_pend.selection()
        if not sel: return
        id_ped = self.tree_pend.item(sel)['values'][0]
        win = tk.Toplevel(self.root); win.geometry("300x150"); win.title("Separador")
        cb = ttk.Combobox(win, values=self.obter_funcao_db("Separador"), state="readonly"); cb.pack(pady=20)
        tk.Button(win, text="INICIAR", command=lambda: [self.mover_status(id_ped, "EM SEPARAÇÃO", "h_inicio_sep", datetime.now().strftime("%d/%m/%Y %H:%M"), "separador", cb.get()), win.destroy()]).pack()

    # --- ABA EM SEPARAÇÃO ---
    def setup_aba_em_sep(self):
        self.tree_sep = self.criar_tabela(self.aba_em_sep, ("ID", "Pedido", "Cliente", "Início", "Separador"))
        f = tk.Frame(self.aba_em_sep, bg=self.cor_fundo); f.pack(pady=10)
        tk.Button(f, text="TERMINAR SEPARAÇÃO", bg=self.cor_sucesso, fg="white", font=self.f_btn, command=lambda: self.mover_status(self.tree_sep.item(self.tree_sep.selection())['values'][0], "AGUARDANDO CONFERENCIA", "h_fim_sep", datetime.now().strftime("%d/%m/%Y %H:%M"))).pack(side="left", padx=5)
        tk.Button(f, text="FALTA PEÇA", bg=self.cor_alerta, fg="white", font=self.f_btn, command=lambda: self.abrir_modal_falta(self.tree_sep.item(self.tree_sep.selection())['values'][0], "SEPARAÇÃO")).pack(side="left", padx=5)
        self.tree_sep.bind("<Double-1>", self.abrir_detalhes)
        self.lbl_total_sep = tk.Label(self.aba_em_sep, text="Total de pedidos: 0", 
                               fg=self.cor_destaque, bg=self.cor_fundo, font=("Arial", 10, "bold"))
        self.lbl_total_sep.pack(side="bottom", anchor="e", padx=10, pady=5)

    # --- ABA AGUARDANDO CONFERÊNCIA ---
    def setup_aba_aguard_conf(self):
        self.tree_conf = self.criar_tabela(self.aba_aguard_conf, ("ID", "Pedido", "Cliente", "Início Sep.", "T. Separação"))
        tk.Button(self.aba_aguard_conf, text="INICIAR CONFERÊNCIA", bg=self.cor_sucesso, fg="white", font=self.f_btn, command=self.dlg_iniciar_conf).pack(pady=10)
        self.tree_conf.bind("<Double-1>", self.abrir_detalhes)
        self.lbl_total_aguard_conf = tk.Label(self.aba_aguard_conf, text="Total de pedidos: 0", 
                               fg=self.cor_destaque, bg=self.cor_fundo, font=("Arial", 10, "bold"))
        self.lbl_total_aguard_conf.pack(side="bottom", anchor="e", padx=10, pady=5)

    def dlg_iniciar_conf(self):
        sel = self.tree_conf.selection()
        if not sel: return
        id_ped = self.tree_conf.item(sel)['values'][0]
        win = tk.Toplevel(self.root); win.geometry("300x180"); win.title("Conferente")
        cb = ttk.Combobox(win, values=self.obter_funcao_db("Conferente"), state="readonly"); cb.pack(pady=10)
        tk.Button(win, text="INICIAR", command=lambda: [self.mover_status(id_ped, "EM CONFERENCIA", "h_inicio_conf", datetime.now().strftime("%d/%m/%Y %H:%M"), "conferente", cb.get()), win.destroy()]).pack()

    # --- ABA EM CONFERÊNCIA ---
    def setup_aba_em_conf(self):
        self.tree_em_conf = self.criar_tabela(self.aba_em_conf, ("ID", "Pedido", "Cliente", "Conferente", "Início Conf."))
        f = tk.Frame(self.aba_em_conf, bg=self.cor_fundo); f.pack(pady=10)
        tk.Button(f, text="FINALIZAR CONFERÊNCIA", bg=self.cor_sucesso, fg="white", font=self.f_btn, command=lambda: self.mover_status(self.tree_em_conf.item(self.tree_em_conf.selection())['values'][0], "PRONTO", "h_pronto", datetime.now().strftime("%d/%m/%Y %H:%M"))).pack(side="left", padx=5)
        tk.Button(f, text="FALTA PEÇA", bg=self.cor_alerta, fg="white", font=self.f_btn, command=lambda: self.abrir_modal_falta(self.tree_em_conf.item(self.tree_em_conf.selection())['values'][0], "CONFERÊNCIA")).pack(side="left", padx=5)
        self.lbl_total_em_conf = tk.Label(self.aba_em_conf, text="Total de pedidos: 0", 
                               fg=self.cor_destaque, bg=self.cor_fundo, font=("Arial", 10, "bold"))
        self.lbl_total_em_conf.pack(side="bottom", anchor="e", padx=10, pady=5)

    # --- ABA AGUARDANDO PEÇAS ---
    def setup_aba_aguard_pecas(self):
        self.tree_pecas = self.criar_tabela(self.aba_aguard_pecas, ("ID", "Pedido", "Cliente", "Entrou em Falta", "Motivo"))
        tk.Button(self.aba_aguard_pecas, text="PEÇA CHEGOU", bg=self.cor_destaque, fg="white", font=self.f_btn, command=lambda: self.mover_status(self.tree_pecas.item(self.tree_pecas.selection())['values'][0], "AGUARDANDO CONFERENCIA", "h_fim_pecas", datetime.now().strftime("%d/%m/%Y %H:%M"))).pack(pady=10)
        self.tree_pecas.bind("<Double-1>", self.abrir_detalhes)
        self.lbl_total_aguard_pecas = tk.Label(self.aba_aguard_pecas, text="Total de pedidos: 0", 
                               fg=self.cor_destaque, bg=self.cor_fundo, font=("Arial", 10, "bold"))
        self.lbl_total_aguard_pecas.pack(side="bottom", anchor="e", padx=10, pady=5)

    # --- ABA PRONTOS ---
    def setup_aba_prontos(self):
        # --- NOVO: FRAME DE BUSCA NO TOPO ---
        f_busca = tk.Frame(self.aba_prontos, bg=self.cor_fundo, pady=5)
        f_busca.pack(fill="x")
        
        tk.Label(f_busca, text="🔍 Buscar Pedido ou Cliente:", fg="white", bg=self.cor_fundo, font=("Arial", 10, "bold")).pack(side="left", padx=10)
        self.ent_busca_prontos = tk.Entry(f_busca, width=35, font=("Arial", 11))
        self.ent_busca_prontos.pack(side="left", padx=5)
        
        # Faz a busca automática enquanto você digita
        self.ent_busca_prontos.bind("<KeyRelease>", lambda e: self.atualizar_aba_prontos())
        
        tk.Button(f_busca, text="LIMPAR", bg="#7f8c8d", fg="white", command=self.limpar_busca_prontos).pack(side="left", padx=5)

        # --- TABELA (TREEVIEW) ---
        self.tree_prontos = self.criar_tabela(self.aba_prontos, ("ID", "Pedido", "Cliente", "Volumes", "NF", "Conferente"))
        self.tree_prontos.bind("<Double-1>", self.abrir_detalhes)
        
        # --- FRAME DE BOTÕES DE AÇÃO ---
        f = tk.Frame(self.aba_prontos, bg=self.cor_fundo); f.pack(pady=10)
        tk.Button(f, text="VOLUMES", bg="#34495e", fg="white", font=self.f_btn, command=self.dlg_ins_vol).pack(side="left", padx=2)
        tk.Button(f, text="NF", bg="#9b59b6", fg="white", font=self.f_btn, command=self.dlg_ins_nf).pack(side="left", padx=2)
        tk.Button(f, text="ETIQUETA VOLUME", bg="#f39c12", fg="white", font=self.f_btn, command=self.gerar_etiqueta).pack(side="left", padx=2)
        tk.Button(f, text="ETIQ. PRODUTOS", bg="#16a085", fg="white", font=self.f_btn, command=self.abrir_modal_produtos).pack(side="left", padx=2)
        tk.Button(f, text="ARQUIVAR", bg=self.cor_sucesso, fg="white", font=self.f_btn, command=self.confirmar_arquivamento).pack(side="left", padx=2)
        
        self.lbl_total_prontos = tk.Label(self.aba_prontos, text="Total de pedidos: 0", 
                               fg=self.cor_destaque, bg=self.cor_fundo, font=("Arial", 10, "bold"))
        self.lbl_total_prontos.pack(side="bottom", anchor="e", padx=10, pady=5)

    def atualizar_aba_prontos(self):
        # Limpa a tabela atual
        for i in self.tree_prontos.get_children(): self.tree_prontos.delete(i)
        
        termo = self.ent_busca_prontos.get()
        
        conn = sqlite3.connect('expedicao_v4.db'); c = conn.cursor()
        
        # VOLTAMOS PARA O SEU PADRÃO: Filtra apenas por 'PRONTO'
        # Quando você arquiva, ele vira 'HISTORICO' e para de aparecer aqui sozinho.
        query = "SELECT id, numero_pedido, cliente, volumes, nf, conferente FROM pedidos WHERE status = 'PRONTO'"
        params = []

        if termo:
            query += " AND (numero_pedido LIKE ? OR cliente LIKE ?)"
            params.append(f"%{termo}%")
            params.append(f"%{termo}%")
        
        query += " ORDER BY h_pronto DESC"
        
        c.execute(query, tuple(params))
        rows = c.fetchall()
        for r in rows: self.tree_prontos.insert('', 'end', values=r)
        
        self.lbl_total_prontos.config(text=f"Total de pedidos: {len(rows)}")
        conn.close()

    def limpar_busca_prontos(self):
        self.ent_busca_prontos.delete(0, 'end')
        self.atualizar_aba_prontos()    

    # --- ABA HISTÓRICO ---
    def setup_aba_historico(self):
        self.tree_hist = self.criar_tabela(self.aba_historico, ("ID", "Pedido", "Cliente", "Volumes", "NF", "Conferente", "Arquivado Em"))
        self.tree_hist.bind("<Double-1>", self.abrir_detalhes)

    # --- ABA PESQUISA ATUALIZADA ---
    def setup_aba_pesquisa(self):
        f_filtros = tk.Frame(self.aba_pesquisa, bg=self.cor_card, pady=10)
        f_filtros.pack(fill="x")
        
        row1 = tk.Frame(f_filtros, bg=self.cor_card); row1.pack(fill="x", pady=2, padx=10)
        
        # NOVO: Filtro por Número do Pedido
        tk.Label(row1, text="Nº Pedido:", fg="white", bg=self.cor_card).pack(side="left")
        self.f_num = tk.Entry(row1, width=10); self.f_num.pack(side="left", padx=5)

        tk.Label(row1, text="Cliente:", fg="white", bg=self.cor_card).pack(side="left", padx=5)
        self.f_cli = tk.Entry(row1, width=20); self.f_cli.pack(side="left", padx=5)
        
        tk.Label(row1, text="Transporte:", fg="white", bg=self.cor_card).pack(side="left", padx=5)
        self.f_tra = ttk.Combobox(row1, values=["", "Transportadora", "Retira", "Nosso Carro"], width=15, state="readonly")
        self.f_tra.pack(side="left", padx=5)
        
        row2 = tk.Frame(f_filtros, bg=self.cor_card); row2.pack(fill="x", pady=2, padx=10)
        
        tk.Label(row2, text="Separador:", fg="white", bg=self.cor_card).pack(side="left")
        self.f_sep = ttk.Combobox(row2, values=[""] + self.obter_funcao_db("Separador"), width=15, state="readonly")
        self.f_sep.pack(side="left", padx=5)

        tk.Label(row2, text="Conferente:", fg="white", bg=self.cor_card).pack(side="left", padx=5)
        self.f_conf = ttk.Combobox(row2, values=[""] + self.obter_funcao_db("Conferente"), width=15, state="readonly")
        self.f_conf.pack(side="left", padx=5)

        # NOVO: Seletor de qual data pesquisar
        tk.Label(row2, text="Filtrar Data por:", fg="white", bg=self.cor_card).pack(side="left", padx=5)
        self.f_tipo_data = ttk.Combobox(row2, values=["Entrada", "Separação", "Pronto"], width=12, state="readonly")
        self.f_tipo_data.set("Entrada")
        self.f_tipo_data.pack(side="left", padx=5)

        row3 = tk.Frame(f_filtros, bg=self.cor_card); row3.pack(fill="x", pady=5, padx=10)

        tk.Label(row3, text="Início:", fg="white", bg=self.cor_card).pack(side="left")
        self.f_data_ini = tk.Entry(row3, width=12); self.f_data_ini.pack(side="left", padx=2)
        self.f_data_ini.bind("<KeyRelease>", self.formatar_data)
        
        tk.Label(row3, text="Fim:", fg="white", bg=self.cor_card).pack(side="left", padx=2)
        self.f_data_fim = tk.Entry(row3, width=12); self.f_data_fim.pack(side="left", padx=2)
        self.f_data_fim.bind("<KeyRelease>", self.formatar_data)
        
        tk.Button(row3, text="FILTRAR", bg=self.cor_destaque, fg="white", font=self.f_btn, command=self.acao_pesquisar).pack(side="left", padx=10)
        tk.Button(row3, text="LIMPAR", bg="#7f8c8d", fg="white", font=self.f_btn, command=self.limpar_filtros_pesquisa).pack(side="left")
        tk.Button(row3, text="📊 EXPORTAR RESUMO", bg="#27ae60", fg="white", font=self.f_btn, 
          command=self.exportar_produtividade).pack(side="left", padx=10)

        # Colunas com h_pronto para conferência visual
        colunas = ("ID", "Pedido", "Cliente", "Transporte", "Status", "Separador", "Conferente", "Entrada", "Finalizado", "Fim Sep.")
        self.tree_pesquisa = self.criar_tabela(self.aba_pesquisa, colunas)
        self.tree_pesquisa.bind("<Double-1>", self.abrir_detalhes)

    def acao_pesquisar(self):
        for i in self.tree_pesquisa.get_children(): self.tree_pesquisa.delete(i)
        conn = sqlite3.connect('expedicao_v4.db'); c = conn.cursor()
        
        # Selecionamos também o h_pronto para exibir na tabela
        query = "SELECT id, numero_pedido, cliente, transporte, status, separador, conferente, h_entrada, h_pronto, h_fim_sep FROM pedidos WHERE 1=1"
        params = []
        
        # Filtro por Número (Novo)
        if self.f_num.get(): 
            query += " AND numero_pedido LIKE ?"; params.append(f"%{self.f_num.get()}%")

        if self.f_cli.get(): query += " AND cliente LIKE ?"; params.append(f"%{self.f_cli.get()}%")
        if self.f_tra.get(): query += " AND transporte = ?"; params.append(self.f_tra.get())
        if self.f_sep.get(): query += " AND separador = ?"; params.append(self.f_sep.get())
        if self.f_conf.get(): query += " AND conferente = ?"; params.append(self.f_conf.get())
        
        # Define qual coluna de data usar baseada no Combobox
        # Se for "Pronto", usa h_pronto. Se for "Entrada", usa h_entrada.
        escolha = self.f_tipo_data.get()
    
        if escolha == "Pronto":
            col_data = "h_pronto"
        elif escolha == "Separação":
            col_data = "h_fim_sep"
        else:
            col_data = "h_entrada"
        
        # Filtro de Data Inicial
        if self.f_data_ini.get():
            d = self.f_data_ini.get()
            params.append(f"{d[6:10]}-{d[3:5]}-{d[0:2]}")
            query += f" AND substr({col_data}, 7, 4) || '-' || substr({col_data}, 4, 2) || '-' || substr({col_data}, 1, 2) >= ?"
            
        # Filtro de Data Final
        if self.f_data_fim.get():
            d = self.f_data_fim.get()
            params.append(f"{d[6:10]}-{d[3:5]}-{d[0:2]}")
            query += f" AND substr({col_data}, 7, 4) || '-' || substr({col_data}, 4, 2) || '-' || substr({col_data}, 1, 2) <= ?"
        
        c.execute(query, tuple(params))
        for r in c.fetchall(): self.tree_pesquisa.insert('', 'end', values=r)
        conn.close()

    def formatar_data(self, event):
        # Pega o widget que disparou o evento (ini ou fim)
        entry = event.widget
        conteudo = entry.get()
        
        # Remove qualquer coisa que não seja número
        apenas_numeros = "".join(filter(str.isdigit, conteudo))
        
        # Limita a 8 dígitos (DDMMYYYY)
        apenas_numeros = apenas_numeros[:8]
        
        data_formatada = ""
        for i, num in enumerate(apenas_numeros):
            if i == 2 or i == 4:
                data_formatada += "/"
            data_formatada += num
            
        # Atualiza o campo com a nova formatação
        entry.delete(0, "end")
        entry.insert(0, data_formatada)

    def limpar_filtros_pesquisa(self):
        self.f_num.delete(0, 'end') # Novo
        self.f_cli.delete(0, 'end')
        self.f_data_ini.delete(0, 'end')
        self.f_data_fim.delete(0, 'end')
        self.f_tra.set('')
        self.f_sep.set('')
        self.f_conf.set('')
        self.f_tipo_data.set('Entrada') # Novo
        self.acao_pesquisar()

    # --- MODAL DETALHES (RAIO-X) COM INTEGRAÇÃO DE PRODUTOS ---
    def abrir_detalhes(self, event):
        tree = event.widget; sel = tree.selection()
        if not sel: return
        id_ped = tree.item(sel)['values'][0]
        conn = sqlite3.connect('expedicao_v4.db'); c = conn.cursor()
        c.execute("SELECT * FROM pedidos WHERE id=?", (id_ped,)); p = c.fetchone()
        
        # Busca itens do pedido para exibir no modal
        c.execute("SELECT quantidade, descricao_produto FROM itens_pedido WHERE id_pedido=?", (id_ped,))
        itens = c.fetchall()
        conn.close()

        t_sep, t_conf = self.calcular_diferenca(p[10], p[11]), self.calcular_diferenca(p[12], p[13])
        win = tk.Toplevel(self.root); win.title(f"Raio-X Pedido #{p[1]}"); win.geometry("500x850"); win.configure(bg=self.cor_card)
        
        # Scrollbar para o modal caso tenha muitos produtos
        container = tk.Frame(win, bg=self.cor_card); container.pack(expand=True, fill="both")
        canvas_modal = tk.Canvas(container, bg=self.cor_card, highlightthickness=0)
        scroll_y = ttk.Scrollbar(container, orient="vertical", command=canvas_modal.yview)
        scroll_frame = tk.Frame(canvas_modal, bg=self.cor_card)
        
        scroll_frame.bind("<Configure>", lambda e: canvas_modal.configure(scrollregion=canvas_modal.bbox("all")))
        canvas_modal.create_window((0, 0), window=scroll_frame, anchor="nw", width=480)
        canvas_modal.configure(yscrollcommand=scroll_y.set)
        canvas_modal.pack(side="left", fill="both", expand=True); scroll_y.pack(side="right", fill="y")

        tk.Label(scroll_frame, text="RELATÓRIO DE FLUXO", font=("Arial", 14, "bold"), fg=self.cor_destaque, bg=self.cor_card).pack(pady=15)
        info = [
            ("CLIENTE", p[2], "white"), ("CIDADE", p[3], "white"), ("TRANSPORTE", p[4], "#f1c40f"),
            ("STATUS ATUAL", p[17], self.cor_sucesso), ("-"*50, "", "gray"),
            ("SEPARADOR", p[7], "white"), ("ENTRADA", p[9], "white"), ("INÍCIO SEP.", p[10], "white"),
            ("TEMPO SEPARAÇÃO", t_sep, "#3498db"), ("-"*50, "", "gray"),
            ("CONFERENTE", p[8], "white"), ("INÍCIO CONF.", p[12], "white"), ("FINALIZADO EM", p[13], "white"),
            ("TEMPO CONFERÊNCIA", t_conf, "#3498db"), ("-"*50, "", "gray"),
            ("VOLUMES", p[5] if p[5] else "0", "white"), ("NOTA FISCAL", p[6] if p[6] else "---", "white"),
            ("ARQUIVADO EM", p[16] if p[16] else "Ainda não arquivado", "white")
        ]
        for label, val, cor in info:
            if label.startswith("-"): tk.Label(scroll_frame, text=label, fg=cor, bg=self.cor_card).pack(); continue
            f = tk.Frame(scroll_frame, bg=self.cor_card); f.pack(fill="x", padx=40, pady=2)
            tk.Label(f, text=f"{label}:", fg="gray", bg=self.cor_card, font=("Arial", 9, "bold")).pack(side="left")
            tk.Label(f, text=f" {val}", fg=cor, bg=self.cor_card, font=("Arial", 10)).pack(side="left")

        if p[18] and str(p[18]).strip() != "":
            tk.Label(scroll_frame, text="-"*50, fg="gray", bg=self.cor_card).pack()
            f_obs = tk.Frame(scroll_frame, bg="#c0392b", padx=10, pady=5); f_obs.pack(fill="x", padx=30)
            tk.Label(f_obs, text="OBSERVAÇÕES / FALTAS:", fg="white", bg="#c0392b", font=("Arial", 9, "bold")).pack(anchor="w")
            tk.Label(f_obs, text=p[18], fg="white", bg="#c0392b", font=("Arial", 10), wraplength=400, justify="left").pack(anchor="w")

        # EXIBIÇÃO DOS PRODUTOS NO MODAL
        if itens:
            tk.Label(scroll_frame, text="-"*50, fg="gray", bg=self.cor_card).pack()
            tk.Label(scroll_frame, text="ITENS DO PEDIDO:", fg="cyan", bg=self.cor_card, font=("Arial", 10, "bold")).pack(pady=5)
            f_itens = tk.Frame(scroll_frame, bg="#34495e", padx=10, pady=10); f_itens.pack(fill="x", padx=30, pady=5)
            for qtd, desc in itens:
                tk.Label(f_itens, text=f"{qtd} CX - {desc}", fg="white", bg="#34495e", font=("Arial", 9)).pack(anchor="w")

    # --- ARQUIVAMENTO ---
    def confirmar_arquivamento(self):
        sel = self.tree_prontos.selection()
        if not sel: return
        id_ped = self.tree_prontos.item(sel)['values'][0]
        win = tk.Toplevel(self.root); win.title("Confirmação"); win.geometry("380x300"); win.configure(bg=self.cor_card); win.grab_set()
        tk.Label(win, text="PEDIDO COMPLETO?", fg="white", bg=self.cor_card, font=("Arial", 10, "bold")).pack(pady=15)
        var_falta = tk.StringVar(value="Não")
        def toggle_obs():
            if var_falta.get() == "Sim": txt_obs.pack(pady=5, padx=10); lbl_inst.pack()
            else: txt_obs.pack_forget(); lbl_inst.pack_forget()
        rb_frame = tk.Frame(win, bg=self.cor_card); rb_frame.pack()
        tk.Radiobutton(rb_frame, text="SIM", variable=var_falta, value="Não", bg=self.cor_card, fg="white", selectcolor="black", command=toggle_obs).pack(side="left", padx=10)
        tk.Radiobutton(rb_frame, text="NÃO", variable=var_falta, value="Sim", bg=self.cor_card, fg="white", selectcolor="black", command=toggle_obs).pack(side="left", padx=10)
        lbl_inst = tk.Label(win, text="Faltas:", fg="yellow", bg=self.cor_card, font=("Arial", 8))
        txt_obs = tk.Text(win, height=4, width=40); txt_obs.pack_forget()
        def finalizar():
            obs = "FALTA: " + txt_obs.get("1.0", "end-1c") if var_falta.get() == "Sim" else ""
            self.exec_sql("UPDATE pedidos SET status='HISTORICO', h_arquivado=?, observacoes=? WHERE id=?", (datetime.now().strftime("%d/%m/%Y %H:%M"), obs, id_ped)); win.destroy()
        tk.Button(win, text="CONFIRMAR", bg=self.cor_sucesso, fg="white", command=finalizar).pack(pady=20)

    # --- MODAL PRODUTOS COM PERSISTÊNCIA ---
    def abrir_modal_produtos(self):
        sel = self.tree_prontos.selection()
        if not sel:
            messagebox.showwarning("Atenção", "Selecione um pedido primeiro!"); return
        id_ped = self.tree_prontos.item(sel)['values'][0]
        conn = sqlite3.connect('expedicao_v4.db'); c = conn.cursor()
        c.execute("SELECT cliente FROM pedidos WHERE id=?", (id_ped,))
        res_cli = c.fetchone(); cliente_nome = res_cli[0] if res_cli else "N/A"
        c.execute("SELECT codigo_produto, quantidade FROM itens_pedido WHERE id_pedido=?", (id_ped,))
        itens_salvos = {row[0]: row[1] for row in c.fetchall()}; conn.close()

        produtos_base = [
            ("49", "AGITADOR 8 BLADES COMPLETO"), ("79", "AGITADOR 8 BLADES SIMPLES"),
            ("2433", "LM COM FILTRO"), ("65", "LM08"),
            ("2434", "LM COM BUCHA E PARAFUSO"), ("6539", "BRANCO COMPLETO"),
            ("6537", "INFERIOR BRANCO"), ("71", "COLORMAQ"),
            ("2960", "FACILITE"), ("6969", "AGITADOR BWK")
        ]

        win = tk.Toplevel(self.root); win.title(f"Produtos - {cliente_nome}"); win.geometry("480x650"); win.configure(bg=self.cor_card); win.grab_set()
        tk.Label(win, text=f"PEDIDO #{id_ped} - {cliente_nome}", fg="cyan", bg=self.cor_card, font=("Arial", 10, "bold")).pack(pady=10)

        main_frame = tk.Frame(win, bg=self.cor_card); main_frame.pack(expand=True, fill="both", padx=10)
        canvas_container = tk.Canvas(main_frame, bg=self.cor_card, highlightthickness=0)
        scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=canvas_container.yview)
        scrollable_frame = tk.Frame(canvas_container, bg=self.cor_card)
        scrollable_frame.bind("<Configure>", lambda e: canvas_container.configure(scrollregion=canvas_container.bbox("all")))
        canvas_container.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas_container.configure(yscrollcommand=scrollbar.set)

        entries = {}
        h_frame = tk.Frame(scrollable_frame, bg=self.cor_card); h_frame.pack(fill="x", padx=10, pady=5)
        tk.Label(h_frame, text="QTD CX", fg="gray", bg=self.cor_card, width=8).pack(side="left")
        tk.Label(h_frame, text="DESCRIÇÃO", fg="gray", bg=self.cor_card).pack(side="left", padx=25)

        for cod, desc in produtos_base:
            f = tk.Frame(scrollable_frame, bg=self.cor_card, pady=3); f.pack(fill="x", padx=10)
            ent = tk.Entry(f, width=8, justify="center", font=("Arial", 10, "bold"))
            ent.pack(side="left")
            if cod in itens_salvos: ent.insert(0, itens_salvos[cod])
            entries[cod] = (ent, desc)
            tk.Label(f, text=f"{cod} - {desc}", fg="white", bg=self.cor_card, font=("Arial", 10)).pack(side="left", padx=15)

        canvas_container.pack(side="left", expand=True, fill="both"); scrollbar.pack(side="right", fill="y")

        def acao_final(imprimir=True):
            dados_persist = []
            conn = sqlite3.connect('expedicao_v4.db'); c = conn.cursor()
            c.execute("DELETE FROM itens_pedido WHERE id_pedido=?", (id_ped,))
            for cod, (ent, desc) in entries.items():
                qtd = ent.get().strip()
                if qtd:
                    c.execute("INSERT INTO itens_pedido (id_pedido, codigo_produto, descricao_produto, quantidade) VALUES (?,?,?,?)", (id_ped, cod, desc, qtd))
                    dados_persist.append((qtd, cod, desc))
            conn.commit(); conn.close()
            if imprimir:
                if not dados_persist: messagebox.showwarning("Atenção", "Preencha ao menos uma quantidade!")
                else: self.gerar_pdf_produtos(dados_persist, id_ped, cliente_nome); win.destroy()
            else: messagebox.showinfo("Sucesso", "Quantidades salvas!"); win.destroy()

        btn_f = tk.Frame(win, bg=self.cor_card); btn_f.pack(pady=20, fill="x", padx=40)
        tk.Button(btn_f, text="💾 SÓ SALVAR", bg="#34495e", fg="white", font=self.f_btn, command=lambda: acao_final(False)).pack(side="left", expand=True, padx=5, ipady=5)
        tk.Button(btn_f, text="🖨️ SALVAR E GERAR PDF", bg=self.cor_sucesso, fg="white", font=self.f_btn, command=lambda: acao_final(True)).pack(side="left", expand=True, padx=5, ipady=5)

    def gerar_pdf_produtos(self, dados, id_ped, cliente):
        try:
            nome_pdf = f"produtos_pedido_{id_ped}.pdf"
            pdf = canvas.Canvas(nome_pdf, pagesize=(430, 480)) 
            pdf.setFont("Helvetica-Bold", 14)
            pdf.drawString(20, 455, f"CLIENTE: {cliente.upper()}")
            pdf.line(20, 450, 410, 450)
            x_start, y_start, row_height = 15, 430, 25
            col_widths = [90, 70, 240]
            pdf.setFont("Helvetica-Bold", 14)
            pdf.rect(x_start, y_start-row_height, col_widths[0], row_height)
            pdf.rect(x_start+col_widths[0], y_start-row_height, col_widths[1], row_height)
            pdf.rect(x_start+col_widths[0]+col_widths[1], y_start-row_height, col_widths[2], row_height)
            pdf.drawString(x_start+5, y_start-17, "QUANT. CX"); pdf.drawString(x_start+col_widths[0]+5, y_start-17, "CODIGO"); pdf.drawString(x_start+col_widths[0]+col_widths[1]+5, y_start-17, "DESCRICAO")
            curr_y = y_start - row_height
            pdf.setFont("Helvetica", 14)
            for qtd, cod, desc in dados:
                curr_y -= row_height
                pdf.rect(x_start, curr_y, col_widths[0], row_height)
                pdf.rect(x_start+col_widths[0], curr_y, col_widths[1], row_height)
                pdf.rect(x_start+col_widths[0]+col_widths[1], curr_y, col_widths[2], row_height)
                pdf.drawString(x_start+5, curr_y+7, str(qtd)); pdf.drawString(x_start+col_widths[0]+5, curr_y+7, str(cod)); pdf.drawString(x_start+col_widths[0]+col_widths[1]+5, curr_y+7, str(desc))
            pdf.save(); os.startfile(nome_pdf)
        except Exception as e: messagebox.showerror("Erro", f"Erro PDF: {e}")

    def gerar_etiqueta(self):
        sel = self.tree_prontos.selection()
        if not sel: return
        id_ped = self.tree_prontos.item(sel)['values'][0]
        conn = sqlite3.connect('expedicao_v4.db'); c = conn.cursor()
        c.execute("SELECT numero_pedido, cliente, separador, cidade, transporte, volumes, nf FROM pedidos WHERE id=?", (id_ped,))
        d = c.fetchone(); conn.close()
        if not d[5] or d[5] == "---": 
            messagebox.showwarning("Atenção", "Informe os volumes!"); return
        try:
            total_v = int(d[5]); nome_pdf = f"etiqueta_{d[0]}.pdf"; pdf = canvas.Canvas(nome_pdf, pagesize=(283, 283))
            for i in range(1, total_v + 1):
                pdf.setStrokeColor(colors.black); pdf.setLineWidth(1.5); pdf.roundRect(5, 5, 273, 273, 5, stroke=1)
                pdf.setFillColor(colors.black); pdf.rect(5, 233, 100, 45, fill=1)
                pdf.setFillColor(colors.white); pdf.setFont("Helvetica-Bold", 9); pdf.drawString(12, 262, "Nº DO PEDIDO")
                pdf.setFont("Helvetica-Bold", 24); pdf.drawString(12, 240, f"# {d[0]}")
                pdf.setFillColor(colors.black); pdf.setLineWidth(1); pdf.line(5, 233, 278, 233)
                pdf.setFont("Helvetica-Bold", 8); pdf.drawString(10, 220, "DESTINATÁRIO:")
                pdf.setFont("Helvetica-Bold", 14); pdf.drawString(10, 202, str(d[1]).upper()[:22])
                pdf.line(5, 185, 278, 185)
                pdf.setFont("Helvetica-Bold", 8); pdf.drawString(10, 172, "CIDADE:"); pdf.setFont("Helvetica-Bold", 12); pdf.drawString(10, 158, str(d[3]).upper()[:25])
                pdf.setFont("Helvetica-Bold", 8); pdf.drawString(10, 135, "TRANSPORTE:"); pdf.setFont("Helvetica-Bold", 12); pdf.drawString(10, 120, str(d[4]).upper()[:25])
                pdf.line(5, 105, 278, 105)
                pdf.setFont("Helvetica-Bold", 8); pdf.drawString(10, 90, "NF:"); pdf.setFont("Helvetica-Bold", 20); pdf.drawString(10, 68, str(d[6] if d[6] else "---"))
                pdf.setFillColor(colors.HexColor("#f1f2f6")); pdf.rect(160, 15, 110, 80, fill=1, stroke=1)
                pdf.setFillColor(colors.black); pdf.setFont("Helvetica-Bold", 10); pdf.drawCentredString(215, 80, "VOLUME(S)")
                pdf.setFont("Helvetica-Bold", 40); pdf.drawCentredString(215, 30, f"{i}/{total_v}"); pdf.showPage()
            pdf.save(); os.startfile(nome_pdf)
        except Exception as e: messagebox.showerror("Erro", f"Erro: {str(e)}")

    def configurar_estilos(self):
        s = ttk.Style(); s.theme_use('default')
        s.configure("Treeview", background=self.cor_card, foreground="white", fieldbackground=self.cor_card, rowheight=28)
        s.configure("Treeview.Heading", background=self.cor_destaque, foreground="white", font=("Arial", 9, "bold"))

    def add_entry(self, master, txt):
        tk.Label(master, text=txt, fg="white", bg=self.cor_card).pack()
        e = tk.Entry(master, justify='center', font=("Arial", 11)); e.pack(pady=5, ipady=3, padx=20, fill='x'); return e

    def criar_tabela(self, master, colunas):
        t = ttk.Treeview(master, columns=colunas, show='headings')
        for c in colunas: t.heading(c, text=c); t.column(c, anchor="center", width=130)
        t.pack(expand=True, fill='both', padx=10, pady=10); return t

    def exec_sql(self, query, params=()):
        conn = sqlite3.connect('expedicao_v4.db'); c = conn.cursor(); c.execute(query, params); conn.commit(); conn.close(); self.atualizar_todas_tabelas()

    def mover_status(self, pedido_id, novo_status, campo_hora=None, valor_hora=None, campo_nome=None, valor_nome=None):
        query = f"UPDATE pedidos SET status='{novo_status}'"
        params = []
        if campo_hora: query += f", {campo_hora}=?"; params.append(valor_hora)
        if campo_nome: query += f", {campo_nome}=?"; params.append(valor_nome)
        query += " WHERE id=?"; params.append(pedido_id); self.exec_sql(query, params)

    def obter_funcao_db(self, cargo):
        conn = sqlite3.connect('expedicao_v4.db'); c = conn.cursor(); c.execute("SELECT nome FROM funcionarios WHERE cargo=?", (cargo,))
        n = [r[0] for r in c.fetchall()]; conn.close(); return n

    def abrir_modal_falta(self, id_ped, origem):
        win = tk.Toplevel(self.root); win.title("Falta de Peça"); win.geometry("400x300"); win.grab_set()
        tk.Label(win, text=f"Motivo na {origem}:").pack(pady=10); txt = tk.Text(win, height=5); txt.pack(pady=10, padx=10)
        def salvar():
            self.exec_sql("UPDATE pedidos SET status='AGUARDANDO PECAS', h_inicio_pecas=?, observacoes=? WHERE id=?", (datetime.now().strftime("%d/%m/%Y %H:%M"), txt.get("1.0", "end"), id_ped)); win.destroy()
        tk.Button(win, text="ENVIAR PARA AGUARD. PEÇAS", command=salvar).pack()

    def dlg_ins_vol(self):
        sel = self.tree_prontos.selection()
        if not sel: return
        win = tk.Toplevel(self.root); win.title("Inserir Volumes"); win.geometry("250x120")
        e = tk.Entry(win, font=("Arial", 12), justify="center"); e.pack(pady=15); e.focus_set()
        tk.Button(win, text="SALVAR", bg=self.cor_sucesso, fg="white", font=self.f_btn, width=15,
                  command=lambda: [self.exec_sql("UPDATE pedidos SET volumes=? WHERE id=?", (e.get(), self.tree_prontos.item(sel)['values'][0])), win.destroy()]).pack()

    def dlg_ins_nf(self):
        sel = self.tree_prontos.selection()
        if not sel: return
        win = tk.Toplevel(self.root); win.title("Inserir NF"); win.geometry("250x120")
        e = tk.Entry(win, font=("Arial", 12), justify="center"); e.pack(pady=15); e.focus_set()
        tk.Button(win, text="SALVAR", bg=self.cor_sucesso, fg="white", font=self.f_btn, width=15,
                  command=lambda: [self.exec_sql("UPDATE pedidos SET nf=? WHERE id=?", (e.get(), self.tree_prontos.item(sel)['values'][0])), win.destroy()]).pack()

    def setup_aba_config(self):
        f = tk.Frame(self.aba_config, bg=self.cor_card, pady=20); f.pack(fill="x")
        self.en_nome = tk.Entry(f); self.en_nome.grid(row=0, column=0, padx=5)
        self.cb_cargo = ttk.Combobox(f, values=["Separador", "Conferente"], state="readonly"); self.cb_cargo.grid(row=0, column=1, padx=5)
        tk.Button(f, text="ADICIONAR", command=self.add_func).grid(row=0, column=2, padx=5)
        self.tree_func = self.criar_tabela(self.aba_config, ("ID", "Nome", "Cargo"))

    def abrir_edicao(self):
        id_ped = None
        tabelas = [self.tree_pend, self.tree_sep, self.tree_conf, self.tree_em_conf, self.tree_pecas, self.tree_prontos]
        
        for t in tabelas:
            sel = t.selection()
            if sel:
                item = t.item(sel[0])
                valores = item.get('values', [])
                if valores:
                    id_ped = valores[0]
                break
        
        if id_ped is None:
            messagebox.showwarning("Burdog", "Selecione um pedido para editar!")
            return

        try:
            conn = sqlite3.connect('expedicao_v4.db')
            c = conn.cursor()
            
            c.execute("SELECT numero_pedido, cliente, separador, conferente, transporte FROM pedidos WHERE id=?", (id_ped,))
            dados_atuais = c.fetchone()
            
            c.execute("SELECT nome FROM funcionarios ORDER BY nome")
            lista_equipe = [f[0] for f in c.fetchall()]
            
            conn.close()

            if not dados_atuais:
                return

            win_edit = tk.Toplevel(self.root)
            win_edit.title(f"Editando Pedido #{dados_atuais[0]}")
            win_edit.geometry("400x480")
            win_edit.configure(bg=self.cor_card)
            win_edit.grab_set()

            tk.Label(win_edit, text="EDITAR INFORMAÇÕES", font=("Arial", 12, "bold"), 
                    fg=self.cor_destaque, bg=self.cor_card).pack(pady=15)

            tk.Label(win_edit, text="Separador Responsável:", fg="gray", bg=self.cor_card).pack()
            cb_sep = ttk.Combobox(win_edit, values=lista_equipe, state="readonly", width=30)
            cb_sep.pack(pady=5)
            cb_sep.set(dados_atuais[2] if dados_atuais[2] else "")

            tk.Label(win_edit, text="Conferente Responsável:", fg="gray", bg=self.cor_card).pack()
            cb_conf = ttk.Combobox(win_edit, values=lista_equipe, state="readonly", width=30)
            cb_conf.pack(pady=5)
            cb_conf.set(dados_atuais[3] if dados_atuais[3] else "")

            tk.Label(win_edit, text="Nome do Cliente:", fg="gray", bg=self.cor_card).pack()
            ent_cliente = tk.Entry(win_edit, width=33)
            ent_cliente.insert(0, dados_atuais[1] if dados_atuais[1] else "")
            ent_cliente.pack(pady=5)

            tk.Label(win_edit, text="Número do Pedido:", fg="gray", bg=self.cor_card).pack()
            ent_numero = tk.Entry(win_edit, width=33)
            ent_numero.insert(0, dados_atuais[0] if dados_atuais[0] else "")
            ent_numero.pack(pady=5)

            def salvar():
                if messagebox.askyesno("Confirmar", "Deseja salvar as alterações?"):
                    conn_save = sqlite3.connect('expedicao_v4.db')
                    cs = conn_save.cursor()
                    cs.execute("""UPDATE pedidos SET
                            numero_pedido = ?, 
                            separador = ?, 
                            conferente = ?, 
                            cliente = ? 
                            WHERE id = ?""", 
                            (ent_numero.get(), cb_sep.get(), cb_conf.get(), ent_cliente.get(), id_ped))
                    conn_save.commit()
                    conn_save.close()
                    
                    messagebox.showinfo("Sucesso", "Pedido atualizado!")
                    win_edit.destroy()
                    self.atualizar_todas_tabelas()

            tk.Button(win_edit, text="💾 SALVAR ALTERAÇÕES", bg=self.cor_sucesso, fg="white", 
                    font=("Arial", 10, "bold"), command=salvar, width=25).pack(pady=30)

        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao acessar dados: {e}")

    def exportar_produtividade(self):
        conn = sqlite3.connect('expedicao_v4.db')
        
        # 1. Buscamos todas as colunas de tempo necessárias
        df = pd.read_sql_query("SELECT separador, conferente, h_entrada, h_pronto, h_fim_sep FROM pedidos", conn)
        conn.close()

        # 2. Define a data de referência baseada no que você selecionou no Combobox da tela
        escolha = self.f_tipo_data.get()
        
        if escolha == "Pronto":
            col_ref = "h_pronto"
        elif escolha == "Separação":
            col_ref = "h_fim_sep"
        else:
            col_ref = "h_entrada"

        # Converte para data real do Pandas
        df['data_ref'] = pd.to_datetime(df[col_ref], dayfirst=True, errors='coerce')
        
        # 3. Filtro de Período (Igual à sua tela de pesquisa)
        try:
            if self.f_data_ini.get():
                ini = pd.to_datetime(self.f_data_ini.get(), dayfirst=True)
            else:
                ini = df['data_ref'].min()

            if self.f_data_fim.get():
                fim = pd.to_datetime(self.f_data_fim.get(), dayfirst=True)
                # Ajusta para o final do dia (23:59:59) para não ignorar pedidos do último dia
                fim = fim.replace(hour=23, minute=59, second=59)
            else:
                fim = df['data_ref'].max()
                
            mask = (df['data_ref'] >= ini) & (df['data_ref'] <= fim)
            df_filtrado = df.loc[mask].copy()
            
        except Exception as e:
            messagebox.showerror("Erro de Data", "Verifique se as datas de início e fim estão corretas.")
            return

        # 4. Contagem de Produção (Agrupando por nome)
        resumo_separadores = df_filtrado['separador'].value_counts().reset_index()
        resumo_separadores.columns = ['Nome', 'Total Separado']

        resumo_conferentes = df_filtrado['conferente'].value_counts().reset_index()
        resumo_conferentes.columns = ['Nome', 'Total Conferido']

        # 5. Salvar em Excel
        caminho_arquivo = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            title="Salvar Relatório de Produtividade"
        )

        if caminho_arquivo:
            with pd.ExcelWriter(caminho_arquivo) as writer:
                resumo_separadores.to_excel(writer, sheet_name='Separadores', index=False)
                resumo_conferentes.to_excel(writer, sheet_name='Conferentes', index=False)
                # Aba de auditoria com os horários originais
                df_filtrado.drop(columns=['data_ref']).to_excel(writer, sheet_name='Relatorio_Geral', index=False)
            
            messagebox.showinfo("Sucesso", f"Relatório de {escolha} exportado!")

    def add_func(self):
        if self.en_nome.get(): self.exec_sql("INSERT INTO funcionarios (nome, cargo) VALUES (?,?)", (self.en_nome.get(), self.cb_cargo.get()))

    def atualizar_todas_tabelas(self):
        conn = sqlite3.connect('expedicao_v4.db'); c = conn.cursor()
        for t in [self.tree_pend, self.tree_sep, self.tree_conf, self.tree_em_conf, self.tree_pecas, self.tree_prontos, self.tree_hist, self.tree_func, self.tree_pesquisa]:
            for i in t.get_children(): t.delete(i)
        c.execute("SELECT id, numero_pedido, cliente, h_entrada, transporte FROM pedidos WHERE status='PENDENTE'")
        for r in c.fetchall(): self.tree_pend.insert('', 'end', values=r)
        c.execute("SELECT id, numero_pedido, cliente, h_inicio_sep, separador FROM pedidos WHERE status='EM SEPARAÇÃO'")
        for r in c.fetchall(): self.tree_sep.insert('', 'end', values=r)
        c.execute("SELECT id, numero_pedido, cliente, h_inicio_sep, h_fim_sep FROM pedidos WHERE status='AGUARDANDO CONFERENCIA'")
        for r in c.fetchall(): self.tree_conf.insert('', 'end', values=(r[0], r[1], r[2], r[3], self.calcular_diferenca(r[3], r[4])))
        c.execute("SELECT id, numero_pedido, cliente, conferente, h_inicio_conf FROM pedidos WHERE status='EM CONFERENCIA'")
        for r in c.fetchall(): self.tree_em_conf.insert('', 'end', values=r)
        c.execute("SELECT id, numero_pedido, cliente, h_inicio_pecas, observacoes FROM pedidos WHERE status='AGUARDANDO PECAS'")
        for r in c.fetchall(): self.tree_pecas.insert('', 'end', values=r)
        c.execute("SELECT id, numero_pedido, cliente, volumes, nf, conferente FROM pedidos WHERE status='PRONTO'")
        for r in c.fetchall(): self.tree_prontos.insert('', 'end', values=(r[0], r[1], r[2], r[3] if r[3] else "---", r[4] if r[4] else "---", r[5]))
        c.execute("SELECT id, numero_pedido, cliente, volumes, nf, conferente, h_arquivado FROM pedidos WHERE status='HISTORICO'")
        for r in c.fetchall(): self.tree_hist.insert('', 'end', values=r)
        c.execute("SELECT * FROM funcionarios")
        for r in c.fetchall(): self.tree_func.insert('', 'end', values=r)
        conn.close()
        self.lbl_total_pend.config(text=f"Total de pedidos: {len(self.tree_pend.get_children())}")
        self.lbl_total_sep.config(text=f"Total de pedidos: {len(self.tree_sep.get_children())}")
        self.lbl_total_aguard_conf.config(text=f"Total de pedidos: {len(self.tree_conf.get_children())}")
        self.lbl_total_em_conf.config(text=f"Total de pedidos: {len(self.tree_em_conf.get_children())}")
        self.lbl_total_aguard_pecas.config(text=f"Total de pedidos: {len(self.tree_pecas.get_children())}")
        self.lbl_total_prontos.config(text=f"Total de pedidos: {len(self.tree_prontos.get_children())}")

if __name__ == "__main__":
    root = tk.Tk(); app = SistemaExpedicaoMaster(root); root.mainloop()