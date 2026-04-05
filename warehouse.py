"""
Interface grafica para a aplicacao Warehouse.
Regista leituras (referencias), mantem consumos por referencia,
grava em CSV e (NOVO) grava em SQLite como fonte de verdade.
No final da sessao exporta relatorios CSV (detalhe e totais por referencia do dia).
Tambem permite exportar relatorios CSV diretamente da DB por dia ou intervalo.
Os caminhos de gravacao definem-se em config.ini (na mesma pasta da aplicacao ou do .exe).
"""

import datetime
import csv
import os
import sqlite3
import uuid
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
from typing import Dict, List, Tuple, Optional, Any

from app_paths import APP_DIR
from config_helpers import (
    carregar_caminhos,
    carregar_dropdowns,
    carregar_caminho_description,
    carregar_caminho_logo,
)
from csv_utils import (
    normalizar_referencia,
    carregar_descricoes,
    carregar_lotes_completos,
)
from db_utils import db_path, db_connect, db_init

# -------------------- Identidade visual --------------------
CORES = {
    "azul": "#0024D3",
    "azul_claro": "#00A9EB",
    "cinza_claro": "#8C8C8C",
    "cinza_escuro": "#575757",
    "branco": "#FFFFFF",
    "fundo": "#F5F5F5",
    "painel_titulo": "#E8F0FE",
    "verde": "#2E7D32",
    "vermelho": "#C62828",
}

# -------------------- App --------------------
class WarehouseApp:
    def __init__(self) -> None:
        self.root = tk.Tk()
        self.root.title("Warehouse Control - Consola de Leituras")
        self.root.state("zoomed")
        self.root.minsize(1020, 1000)
        self.root.configure(bg=CORES["fundo"])

        # Estado da sessão
        self.sessao_iniciada = False
        self.operador = ""
        self.projeto = ""
        self.turno = ""
        self.inicio_sessao = None
        self.consumos: Dict[str, int] = {}
        # Guardamos todas as leituras da sessao:
        # (id, referencia, quantidade, timestamp, comentario, lote, description)
        self.ultimas_leituras: List[tuple] = []
        self.logfile: Optional[str] = None
        self.log_dir, self.bom_path, self.db_dir = carregar_caminhos()
        self.description_path = carregar_caminho_description()
        self.logo_path = carregar_caminho_logo()
        self.descricoes_ref: Dict[str, str] = {}
        self._desc_mtime = None
        self._logo_img = None
        self._timer_duracao = None
        self._timer_foco_referencia = None

        # SQLite
        self.db_con: Optional[sqlite3.Connection] = None
        self.db_path: Optional[str] = None
        self.sessao_id: Optional[str] = None

        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        self._construir_interface()
        self._carregar_descricoes_se_necessario(force=True)

    def _on_close(self) -> None:
        """Handler de fecho da janela principal."""
        if self.sessao_iniciada:
            if not messagebox.askyesno(
                "Terminar aplicação",
                "Existe uma sessao em curso.\n\n"
                "Deseja terminar a sessao atual e fechar a aplicação?",
            ):
                return
            self._terminar_sessao()
        self.root.destroy()

    # -------------------- UI --------------------
    def _construir_interface(self) -> None:
        c = CORES

        # Header
        header = tk.Frame(self.root, bg=c["azul"], height=64)
        header.pack(fill=tk.X)
        header.pack_propagate(False)

        if os.path.isfile(self.logo_path):
            try:
                self._logo_img = tk.PhotoImage(file=self.logo_path)
                logo_label = tk.Label(header, image=self._logo_img, bg=c["azul"])
                logo_label.pack(side=tk.LEFT, padx=16, pady=12)
            except tk.TclError:
                self._logo_img = None

        if self._logo_img is None:
            tk.Label(header, text="FORVIA", font=("Segoe UI", 18, "bold"),
                     fg=c["branco"], bg=c["azul"]).pack(side=tk.LEFT, padx=16, pady=12)

        tk.Label(header, text="Warehouse Control - Consola de Leituras",
                 font=("Segoe UI", 14, "bold"), fg=c["branco"], bg=c["azul"]).pack(side=tk.LEFT, padx=20, pady=14)

        # Main
        main = tk.Frame(self.root, bg=c["fundo"], padx=12, pady=12)
        main.pack(fill=tk.BOTH, expand=True)

        left = tk.Frame(main, bg=c["fundo"])
        left.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Sessão
        self._painel_titulo(left, "Sessao")
        frame_sessao_linha = tk.Frame(left, bg=c["fundo"])
        frame_sessao_linha.pack(fill=tk.X, pady=(0, 10))

        frame_sessao = tk.Frame(frame_sessao_linha, bg=c["branco"], padx=12, pady=10, relief=tk.FLAT)
        frame_sessao.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        projetos_list, turnos_list = carregar_dropdowns()

        tk.Label(frame_sessao, text="Operador:", font=("Segoe UI", 9),
                 fg=c["cinza_escuro"], bg=c["branco"]).grid(row=0, column=0, sticky=tk.W, pady=2)
        self.entry_operador = tk.Entry(frame_sessao, width=18, font=("Segoe UI", 10), relief=tk.SOLID, bd=1)
        self.entry_operador.grid(row=0, column=1, padx=(8, 12), pady=2)

        tk.Label(frame_sessao, text="Projeto/Linha:", font=("Segoe UI", 9),
                 fg=c["cinza_escuro"], bg=c["branco"]).grid(row=1, column=0, sticky=tk.W, pady=2)
        self.combo_projeto = ttk.Combobox(frame_sessao, width=18, font=("Segoe UI", 10), values=projetos_list, state="readonly")
        if projetos_list:
            self.combo_projeto.set(projetos_list[0])
        self.combo_projeto.grid(row=1, column=1, padx=(8, 12), pady=2)

        tk.Label(frame_sessao, text="Turno:", font=("Segoe UI", 9),
                 fg=c["cinza_escuro"], bg=c["branco"]).grid(row=2, column=0, sticky=tk.W, pady=2)
        self.combo_turno = ttk.Combobox(frame_sessao, width=18, font=("Segoe UI", 10), values=turnos_list, state="readonly")
        if turnos_list:
            self.combo_turno.set(turnos_list[0])
        self.combo_turno.grid(row=2, column=1, padx=(8, 12), pady=2)

        self.label_sessao = tk.Label(frame_sessao, text="Introduza operador e clique em Iniciar sessao.",
                                     font=("Segoe UI", 9), fg=c["cinza_claro"], bg=c["branco"])
        self.label_sessao.grid(row=3, column=0, columnspan=2, sticky=tk.W, pady=(6, 4))

        self.btn_iniciar = tk.Button(frame_sessao, text="Iniciar sessao", font=("Segoe UI", 10, "bold"),
                                     fg=c["branco"], bg=c["azul"], activebackground=c["azul_claro"],
                                     activeforeground=c["branco"], relief=tk.FLAT, padx=12, pady=4,
                                     cursor="hand2", command=self._iniciar_sessao)
        self.btn_iniciar.grid(row=4, column=0, columnspan=2, pady=(4, 0))

        # Hora ao lado
        frame_hora = tk.Frame(frame_sessao_linha, bg=c["azul"], padx=16, pady=12, relief=tk.FLAT)
        frame_hora.pack(side=tk.LEFT, fill=tk.Y)

        self.label_hora = tk.Label(frame_hora, text="--:--:--", font=("Segoe UI", 22, "bold"),
                                   fg=c["branco"], bg=c["azul"])
        self.label_hora.pack()

        self.label_inicio_sessao = tk.Label(frame_hora, text="Sessao iniciada a\n--:--:--",
                                            font=("Segoe UI", 9), fg=c["branco"], bg=c["azul"],
                                            justify=tk.CENTER)
        self.label_inicio_sessao.pack(pady=(4, 0))
        self._atualizar_hora()

        # Leitura
        self._painel_titulo(left, "Leitura")
        frame_leitura = tk.Frame(left, bg=c["branco"], padx=12, pady=10, relief=tk.FLAT)
        frame_leitura.pack(fill=tk.X, pady=(0, 10))

        frame_leitura_cols = tk.Frame(frame_leitura, bg=c["branco"])
        frame_leitura_cols.pack(fill=tk.X)

        col_manual = tk.Frame(frame_leitura_cols, bg=c["branco"])
        col_manual.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 10))

        col_lote = tk.Frame(frame_leitura_cols, bg=c["branco"])
        col_lote.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(10, 0))

        tk.Label(col_manual, text="Referencia:", font=("Segoe UI", 9), fg=c["cinza_escuro"], bg=c["branco"]).pack(anchor=tk.W)
        self.entry_referencia = tk.Entry(col_manual, width=36, font=("Segoe UI", 12), relief=tk.SOLID, bd=1)
        self.entry_referencia.pack(fill=tk.X, pady=(2, 8))
        self.entry_referencia.bind("<Return>", lambda e: self._registar_leitura())
        self.entry_referencia.bind("<FocusIn>", lambda e: self.entry_referencia.select_range(0, tk.END))

        row_qty = tk.Frame(col_manual, bg=c["branco"])
        row_qty.pack(fill=tk.X, pady=(0, 8))
        tk.Label(row_qty, text="Quantidade:", font=("Segoe UI", 9), fg=c["cinza_escuro"], bg=c["branco"]).pack(side=tk.LEFT, padx=(0, 8))
        self.var_quantidade = tk.StringVar(value="1")
        self.spin_quantidade = tk.Spinbox(row_qty, from_=0, to=9999, width=8, textvariable=self.var_quantidade, font=("Segoe UI", 10))
        self.spin_quantidade.pack(side=tk.LEFT, padx=(0, 4))
        self.spin_quantidade.bind("<KeyRelease>", self._validar_quantidade_teclado)
        self.spin_quantidade.bind("<FocusOut>", self._normalizar_quantidade)
        tk.Button(row_qty, text="-", width=2, font=("Segoe UI", 10),
                  fg=c["cinza_escuro"], bg=c["fundo"], relief=tk.FLAT,
                  command=lambda: self._alterar_quantidade(-1)).pack(side=tk.LEFT, padx=2)
        tk.Button(row_qty, text="+", width=2, font=("Segoe UI", 10),
                  fg=c["cinza_escuro"], bg=c["fundo"], relief=tk.FLAT,
                  command=lambda: self._alterar_quantidade(1)).pack(side=tk.LEFT)

        tk.Label(col_manual, text="Comentario (opcional):", font=("Segoe UI", 9),
                 fg=c["cinza_escuro"], bg=c["branco"]).pack(anchor=tk.W, pady=(4, 2))
        self.text_comentario = tk.Text(col_manual, height=2, width=36, font=("Segoe UI", 10),
                                       relief=tk.SOLID, bd=1, wrap=tk.WORD)
        self.text_comentario.pack(fill=tk.X, pady=(0, 8))

        largura_btn_registo = 20
        self.btn_registar = tk.Button(col_manual, text="  REGISTAR (Enter)  ", width=largura_btn_registo,
                                      font=("Segoe UI", 11, "bold"),
                                      fg=c["branco"], bg=c["verde"], activebackground="#1B5E20",
                                      activeforeground=c["branco"], relief=tk.FLAT, padx=16, pady=6,
                                      cursor="hand2", command=self._registar_leitura)
        self.btn_registar.pack(anchor=tk.W, pady=(4, 0))

        # Lote completo
        tk.Label(col_lote, text="Lote completo:", font=("Segoe UI", 9), fg=c["cinza_escuro"], bg=c["branco"]).pack(anchor=tk.W)
        self.lotes_completos = carregar_lotes_completos(self.bom_path)
        nomes_lotes = list(self.lotes_completos.keys())
        self.nomes_lotes_todos = nomes_lotes
        self.combo_lote = ttk.Combobox(col_lote, width=36, font=("Segoe UI", 10), values=nomes_lotes, state="normal")
        if nomes_lotes:
            self.combo_lote.set(nomes_lotes[0])
        self.combo_lote.bind("<KeyRelease>", self._filtrar_lotes)
        self.combo_lote.pack(fill=tk.X, pady=(2, 8))

        tk.Frame(col_lote, bg=c["branco"], height=110).pack(fill=tk.X)

        self.btn_registar_lote = tk.Button(
            col_lote,
            text="  REGISTAR LOTE  ",
            width=largura_btn_registo,
            font=("Segoe UI", 11, "bold"),
            fg=c["branco"],
            bg=c["azul"],
            activebackground=c["azul_claro"],
            activeforeground=c["branco"],
            relief=tk.FLAT,
            padx=16,
            pady=6,
            cursor="hand2",
            command=self._registar_lote_completo,
            state=tk.NORMAL if nomes_lotes else tk.DISABLED,
        )
        self.btn_registar_lote.pack(anchor=tk.W, pady=(4, 0))

        # Leituras da sessao
        self._painel_titulo(left, "Leituras da sessao")
        frame_ultimas = tk.Frame(left, bg=c["branco"], padx=8, pady=8, relief=tk.FLAT)
        frame_ultimas.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        self.list_ultimas = tk.Listbox(frame_ultimas, height=10, font=("Consolas", 10),
                                       bg=c["branco"], fg=c["cinza_escuro"],
                                       selectbackground=c["azul"], selectforeground=c["branco"],
                                       relief=tk.FLAT, highlightthickness=0)
        scroll_ultimas = tk.Scrollbar(frame_ultimas, orient=tk.VERTICAL, command=self.list_ultimas.yview, bg=c["cinza_claro"])
        self.list_ultimas.configure(yscrollcommand=scroll_ultimas.set)
        self.list_ultimas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll_ultimas.pack(side=tk.RIGHT, fill=tk.Y)

        btn_eliminar = tk.Button(frame_ultimas, text="Eliminar leitura selecionada",
                                 font=("Segoe UI", 9), fg=c["azul"], bg=c["branco"],
                                 activeforeground=c["azul_claro"], relief=tk.FLAT,
                                 cursor="hand2", command=self._eliminar_leitura)
        btn_eliminar.pack(anchor=tk.W, pady=(6, 0))

        btn_editar_coment = tk.Button(frame_ultimas, text="Editar/adicionar comentario",
                                      font=("Segoe UI", 9), fg=c["azul"], bg=c["branco"],
                                      activeforeground=c["azul_claro"], relief=tk.FLAT,
                                      cursor="hand2", command=self._editar_comentario_leitura)
        btn_editar_coment.pack(anchor=tk.W, pady=(4, 0))

        # Coluna direita
        right = tk.Frame(main, bg=c["fundo"])
        right.pack(side=tk.LEFT, fill=tk.BOTH, expand=False, padx=(16, 0))

        self._painel_titulo(right, "Resumo da sessao")
        frame_resumo = tk.Frame(right, bg=c["branco"], padx=8, pady=8, relief=tk.FLAT)
        frame_resumo.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        self.text_resumo = tk.Text(frame_resumo, width=36, height=20, font=("Consolas", 10), state=tk.DISABLED,
                                   bg=c["branco"], fg=c["cinza_escuro"], relief=tk.FLAT, wrap=tk.WORD)
        scroll_resumo = tk.Scrollbar(frame_resumo, orient=tk.VERTICAL, command=self.text_resumo.yview, bg=c["cinza_claro"])
        self.text_resumo.configure(yscrollcommand=scroll_resumo.set)
        self.text_resumo.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll_resumo.pack(side=tk.RIGHT, fill=tk.Y)

        self._painel_titulo(right, "Relatorios")
        frame_relatorios = tk.Frame(right, bg=c["branco"], padx=10, pady=10, relief=tk.FLAT)
        frame_relatorios.pack(fill=tk.X, pady=(0, 10))
        tk.Label(
            frame_relatorios,
            text="Extracao CSV",
            font=("Segoe UI", 9),
            fg=c["cinza_escuro"],
            bg=c["branco"],
        ).pack(anchor=tk.W)
        self.btn_abrir_relatorios = tk.Button(
            frame_relatorios,
            text="  Abrir exportacao CSV  ",
            font=("Segoe UI", 10, "bold"),
            fg=c["branco"],
            bg=c["azul"],
            activebackground=c["azul_claro"],
            activeforeground=c["branco"],
            relief=tk.FLAT,
            padx=10,
            pady=4,
            cursor="hand2",
            command=self._abrir_janela_exportacao_csv_db,
        )
        self.btn_abrir_relatorios.pack(anchor=tk.W, pady=(8, 0))

        # Footer
        footer = tk.Frame(self.root, bg=c["branco"], padx=16, pady=12)
        footer.pack(fill=tk.X)

        self.label_total = tk.Label(footer, text=" Total do dia: 0", font=("Segoe UI", 11, "bold"),
                                    fg=c["azul"], bg=c["branco"])
        self.label_total.pack(side=tk.LEFT)

        self.label_leituras_sessao = tk.Label(footer, text=" Sessao: 0", font=("Segoe UI", 9),
                                              fg=c["cinza_escuro"], bg=c["branco"])
        self.label_leituras_sessao.pack(side=tk.LEFT, padx=(20, 0))

        self.label_refs = tk.Label(footer, text=" Ref. unicas: 0", font=("Segoe UI", 9),
                                   fg=c["cinza_escuro"], bg=c["branco"])
        self.label_refs.pack(side=tk.LEFT, padx=(12, 0))

        self.label_duracao = tk.Label(footer, text=" Duracao: 00:00:00", font=("Segoe UI", 9),
                                      fg=c["cinza_escuro"], bg=c["branco"])
        self.label_duracao.pack(side=tk.LEFT, padx=(12, 0))

        self.btn_terminar = tk.Button(footer, text="  Terminar sessao  ", font=("Segoe UI", 10, "bold"),
                                      fg=c["branco"], bg=c["vermelho"], activebackground="#B71C1C",
                                      activeforeground=c["branco"], relief=tk.FLAT, padx=12, pady=4,
                                      cursor="hand2", command=self._terminar_sessao)
        self.btn_terminar.pack(side=tk.RIGHT)

        footer2 = tk.Frame(self.root, bg=c["branco"], padx=16, pady=6)
        footer2.pack(fill=tk.X)
        tk.Label(footer2, text="Desenvolvido por Bruno Santos - 2026 - v5", font=("Segoe UI", 8),
                 fg=c["cinza_claro"], bg=c["branco"]).pack(anchor=tk.W)

        self._atualizar_resumo()
        self.root.bind_all("<FocusIn>", self._on_focus_change, add="+")
        self.entry_referencia.focus_set()
        self._agendar_retorno_referencia()

    def _painel_titulo(self, parent: tk.Widget, texto: str) -> None:
        f = tk.Frame(parent, bg=CORES["painel_titulo"], padx=10, pady=6)
        f.pack(fill=tk.X)
        tk.Label(f, text=texto, font=("Segoe UI", 10, "bold"),
                 fg=CORES["cinza_escuro"], bg=CORES["painel_titulo"]).pack(anchor=tk.W)

    # -------------------- Focus helpers --------------------
    def _on_focus_change(self, event: Optional[tk.Event] = None) -> None:
        self._agendar_retorno_referencia()

    def _agendar_retorno_referencia(self) -> None:
        if self._timer_foco_referencia:
            self.root.after_cancel(self._timer_foco_referencia)
        self._timer_foco_referencia = self.root.after(150000, self._retornar_foco_referencia)

    def _retornar_foco_referencia(self) -> None:
        self._timer_foco_referencia = None
        if self.root.state() == "iconic" or self.root.focus_displayof() is None:
            self._agendar_retorno_referencia()
            return
        if self.root.focus_get() != self.entry_referencia:
            try:
                self.entry_referencia.focus_set()
                self.entry_referencia.icursor(tk.END)
            except tk.TclError:
                pass
        self._agendar_retorno_referencia()

    # -------------------- Hora / duração --------------------
    def _atualizar_hora(self) -> None:
        now = datetime.datetime.now()
        self.label_hora.configure(text=now.strftime("%H:%M:%S"))
        if self.sessao_iniciada and self.inicio_sessao:
            self.label_inicio_sessao.configure(text=f"Sessao iniciada a\n{self.inicio_sessao.strftime('%H:%M:%S')}")
        else:
            self.label_inicio_sessao.configure(text="Sessao iniciada a\n--:--:--")
        self.root.after(1000, self._atualizar_hora)

    def _atualizar_duracao(self) -> None:
        if not self.sessao_iniciada or not self.inicio_sessao:
            self._timer_duracao = None
            return
        delta = datetime.datetime.now() - self.inicio_sessao
        h, r = divmod(int(delta.total_seconds()), 3600)
        m, s = divmod(r, 60)
        self.label_duracao.configure(text=f" Duracao: {h:02d}:{m:02d}:{s:02d}")
        self._timer_duracao = self.root.after(1000, self._atualizar_duracao)

    # -------------------- Sessão --------------------
    def _iniciar_sessao(self) -> None:
        op = self.entry_operador.get().strip()
        proj = self.combo_projeto.get().strip() if self.combo_projeto.get() else ""
        turno = self.combo_turno.get().strip() if self.combo_turno.get() else ""

        if not op:
            messagebox.showwarning("Campos em falta", "Preencha o Operador antes de iniciar.")
            return
        if not proj:
            messagebox.showwarning("Campos em falta", "Selecione Projeto/Linha antes de iniciar.")
            return

        # Fechar DB anterior (se existir)
        try:
            if self.db_con is not None:
                self.db_con.close()
        except Exception:
            pass
        self.db_con = None

        self.operador = op
        self.projeto = proj
        self.turno = turno
        self.inicio_sessao = datetime.datetime.now()
        self.log_dir, self.bom_path, self.db_dir = carregar_caminhos()
        self.description_path = carregar_caminho_description()
        self.logo_path = carregar_caminho_logo()

        os.makedirs(self.log_dir, exist_ok=True)
        db_parent = self.db_dir if not str(self.db_dir).lower().endswith(".db") else (os.path.dirname(self.db_dir) or APP_DIR)
        os.makedirs(db_parent, exist_ok=True)
        self.logfile = os.path.join(self.log_dir, f"log_{self.inicio_sessao.date()}.csv")

        # SQLite init
        try:
            self.sessao_id = uuid.uuid4().hex
            self.db_path = db_path(self.db_dir)
            self.db_con = db_connect(self.db_path)
            db_init(self.db_con)
        except sqlite3.Error as err:
            self.db_con = None
            messagebox.showerror("Erro DB", f"Não foi possível iniciar a base de dados.\n\n{err}")
            return

        self.consumos = {}
        self.ultimas_leituras.clear()
        self.sessao_iniciada = True
        self._carregar_descricoes_se_necessario(force=True)

        self.entry_operador.configure(state="disabled")
        self.combo_projeto.configure(state="disabled")
        self.combo_turno.configure(state="disabled")
        self.btn_iniciar.configure(state=tk.DISABLED)

        txt = f"Sessao iniciada as {self.inicio_sessao.strftime('%H:%M:%S')}  {self.operador} | {self.projeto}"
        if self.turno:
            txt += f" | {self.turno}"

        self.label_sessao.configure(text=txt, fg=CORES["verde"])
        self.label_inicio_sessao.configure(text=f"Sessao iniciada a\n{self.inicio_sessao.strftime('%H:%M:%S')}")
        self._atualizar_ultimas()
        self._atualizar_resumo()
        self._atualizar_duracao()
        self.entry_referencia.focus_set()

    # -------------------- Leitura --------------------
    def _is_inventario(self) -> bool:
        proj = self.projeto if self.sessao_iniciada else self.combo_projeto.get().strip()
        return proj.strip().lower() == "inventario"

    def _obter_quantidade(self) -> int:
        texto = self.spin_quantidade.get().strip()
        if not texto:
            return 0 if self._is_inventario() else 1
        try:
            n = int(texto)
        except ValueError:
            return 0 if self._is_inventario() else 1
        minimo = 0 if self._is_inventario() else 1
        return max(minimo, min(9999, n))

    def _validar_quantidade_teclado(self, event: Optional[tk.Event] = None) -> None:
        texto = self.spin_quantidade.get()
        if not texto:
            return
        filtrado = "".join(ch for ch in texto if ch.isdigit())
        if filtrado != texto:
            self.spin_quantidade.delete(0, tk.END)
            self.spin_quantidade.insert(0, filtrado)

    def _normalizar_quantidade(self, event: Optional[tk.Event] = None) -> None:
        q = self._obter_quantidade()
        self.var_quantidade.set(str(q))

    def _alterar_quantidade(self, delta: int) -> None:
        q = self._obter_quantidade()
        minimo = 0 if self._is_inventario() else 1
        q = max(minimo, min(9999, q + delta))
        self.var_quantidade.set(str(q))

    def _filtrar_lotes(self, event: Optional[tk.Event] = None) -> None:
        if not hasattr(self, "combo_lote") or not hasattr(self, "nomes_lotes_todos"):
            return
        texto = self.combo_lote.get().strip().lower()
        if texto:
            filtrados = [nome for nome in self.nomes_lotes_todos if texto in nome.lower()]
        else:
            filtrados = list(self.nomes_lotes_todos)
        self.combo_lote.configure(values=filtrados if filtrados else self.nomes_lotes_todos)

    def _registar_leitura(self) -> None:
        if not self.sessao_iniciada:
            messagebox.showwarning("Sessao nao iniciada", "Inicie a sessao antes de registar leituras.")
            return

        referencia = self.entry_referencia.get().strip().upper()
        if not referencia:
            return

        if referencia == "EXIT":
            self._terminar_sessao()
            return

        quantidade = self._obter_quantidade()
        comentario = self.text_comentario.get(1.0, tk.END).strip().replace("\n", " ").replace(";", ",")

        self._registar_item(referencia, quantidade, comentario)

        self.entry_referencia.delete(0, tk.END)
        self.var_quantidade.set("1")
        self.text_comentario.delete(1.0, tk.END)

        self._atualizar_ultimas()
        self._atualizar_resumo()
        self.entry_referencia.focus_set()

    def _registar_item(self, referencia: str, quantidade: int = 1, comentario: str = "", lote: str = "") -> None:
        """Regista no SQLite (fonte de verdade) + CSV (compatibilidade)."""
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self._carregar_descricoes_se_necessario()
        descricao_registo = self.descricoes_ref.get(normalizar_referencia(referencia), "").strip()

        # Memória (para o resumo) - igual
        self.consumos[referencia] = self.consumos.get(referencia, 0) + int(quantidade)

        # 1) SQLite
        row_id = None
        try:
            if self.db_con is None:
                raise sqlite3.Error("Ligação DB não inicializada.")
            cur = self.db_con.execute(
                """INSERT INTO leituras
                   (ts, operador, projeto, turno, referencia, description, quantidade, comentario, lote, sessao_id)
                   VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                (timestamp, self.operador, self.projeto, self.turno,
                 referencia, descricao_registo, int(quantidade), comentario, lote, self.sessao_id)
            )
            self.db_con.commit()
            row_id = cur.lastrowid
        except sqlite3.Error as err:
            messagebox.showerror("Erro DB", str(err))
            return

        # Mantem a leitura mais recente no topo da lista
        self.ultimas_leituras.insert(0, (row_id, referencia, int(quantidade), timestamp, comentario, lote, descricao_registo))

        # 2) CSV (mantém)
        try:
            escrever_cabecalho = (not os.path.isfile(self.logfile)) or (os.path.getsize(self.logfile) == 0)
            if not escrever_cabecalho:
                self._garantir_cabecalho_csv_com_lote()

            with open(self.logfile, mode="a", newline="", encoding="utf-8") as f:
                w = csv.writer(f, delimiter=";")
                if escrever_cabecalho:
                    w.writerow(["Data", "Operador", "Projeto", "Turno", "Referencia", "Description", "Quantidade", "Comentario", "Lote"])
                w.writerow([timestamp, self.operador, self.projeto, self.turno, referencia, descricao_registo, int(quantidade), comentario, lote])
                f.flush()
                os.fsync(f.fileno())
        except OSError as err:
            # DB já tem o registo; CSV é secundário
            messagebox.showwarning("Aviso CSV", f"Registo guardado na BD, mas falhou a escrita no CSV:\n\n{err}")

    def _garantir_cabecalho_csv_com_lote(self):
        """Atualiza CSV existente para o formato com colunas Description e Lote."""
        if not self.logfile or not os.path.isfile(self.logfile):
            return
        try:
            with open(self.logfile, mode="r", newline="", encoding="utf-8") as f:
                rows = list(csv.reader(f, delimiter=";"))
        except OSError:
            return
        if not rows:
            return

        header = rows[0]
        header_correto = (
            len(header) >= 9
            and header[0] == "Data"
            and header[1] == "Operador"
            and header[5] in ("Description", "Descricao")
            and header[8] == "Lote"
        )
        if header_correto:
            return
        if len(header) < 2 or header[0] != "Data" or header[1] != "Operador":
            return

        novos_rows = [["Data", "Operador", "Projeto", "Turno", "Referencia", "Description", "Quantidade", "Comentario", "Lote"]]
        for row in rows[1:]:
            if len(row) >= 9:
                row_n = row[:9]
            elif len(row) == 8:
                row_n = [row[0], row[1], row[2], row[3], row[4], "", row[5], row[6], row[7]]
            elif len(row) == 7:
                row_n = [row[0], row[1], row[2], row[3], row[4], "", row[5], row[6], ""]
            elif len(row) == 6:
                row_n = [row[0], row[1], row[2], "", row[3], "", row[4], row[5], ""]
            elif len(row) >= 5:
                row_n = [row[0], row[1], row[2], "", row[3], "", row[4], "", ""]
            else:
                continue
            novos_rows.append(row_n)

        try:
            with open(self.logfile, mode="w", newline="", encoding="utf-8") as f:
                w = csv.writer(f, delimiter=";")
                w.writerows(novos_rows)
        except OSError:
            pass

    def _registar_lote_completo(self):
        """Regista todas as referencias configuradas no lote selecionado."""
        if not self.sessao_iniciada:
            messagebox.showwarning("Sessao nao iniciada", "Inicie a sessao antes de registar leituras.")
            return

        nome_digitado = self.combo_lote.get().strip() if hasattr(self, "combo_lote") else ""
        if not nome_digitado:
            messagebox.showwarning("Lote em falta", "Selecione um lote completo.")
            return

        nome_lote = next((n for n in self.nomes_lotes_todos if n.lower() == nome_digitado.lower()), "")
        if not nome_lote:
            matches = [n for n in self.nomes_lotes_todos if nome_digitado.lower() in n.lower()]
            if len(matches) == 1:
                nome_lote = matches[0]
                self.combo_lote.set(nome_lote)
            else:
                messagebox.showwarning("Lote invalido", "Nao foi encontrado um lote unico com esse nome.")
                return

        itens = self.lotes_completos.get(nome_lote, [])
        if not itens:
            messagebox.showwarning("Lote invalido", "O lote selecionado nao tem referencias configuradas.")
            return

        for ref, qty in itens:
            self._registar_item(ref, qty, "", nome_lote)

        self.entry_referencia.delete(0, tk.END)
        self.var_quantidade.set("1")
        self.text_comentario.delete(1.0, tk.END)
        self._atualizar_ultimas()
        self._atualizar_resumo()
        self.entry_referencia.focus_set()

    # -------------------- Listas / Resumo --------------------
    def _atualizar_ultimas(self) -> None:
        self.list_ultimas.delete(0, tk.END)
        self._carregar_descricoes_se_necessario()
        total = len(self.ultimas_leituras)
        for i, item in enumerate(self.ultimas_leituras, 1):
            if isinstance(item, tuple) and len(item) >= 6 and isinstance(item[0], int):
                _, ref, qty, timestamp, coment, _, descricao_item = (item + ("",))[:7]
            else:
                # fallback antigo, se algum item ficar em formato legado
                ref, qty = item[0], (item[1] if len(item) > 1 else 1)
                timestamp = item[2] if len(item) > 2 else ""
                coment = item[3] if len(item) > 3 else ""
                descricao_item = ""

            hora = "--:--:--"
            if isinstance(timestamp, str) and timestamp:
                hora = timestamp.rsplit(" ", 1)[-1] if " " in timestamp else timestamp[-8:]

            ordem_sessao = total - i + 1
            descricao = descricao_item or self.descricoes_ref.get(normalizar_referencia(ref), "")
            comentario_txt = coment.strip() if isinstance(coment, str) else str(coment or "")
            texto = f"{hora} {ordem_sessao}. {ref} {descricao} {qty} {comentario_txt}".strip()
            self.list_ultimas.insert(tk.END, texto)

    def _carregar_descricoes_se_necessario(self, force: bool = False) -> None:
        path = carregar_caminho_description()
        self.description_path = path

        mtime = None
        if path and os.path.isfile(path):
            try:
                mtime = os.path.getmtime(path)
            except OSError:
                mtime = None

        if force or mtime != self._desc_mtime:
            self.descricoes_ref = carregar_descricoes(path)
            self._desc_mtime = mtime

    def _total_do_dia(self) -> int:
        """Total do dia via SQLite (se existir) com fallback CSV."""
        dia = datetime.date.today().strftime("%Y-%m-%d")

        # SQLite (usa sempre a BD se existir, mesmo fora de sessao)
        con: Optional[sqlite3.Connection] = None
        fechar_con = False
        try:
            try:
                con, fechar_con = self._obter_conexao_db_relatorio()
            except FileNotFoundError:
                con = None

            if con is not None:
                row = con.execute(
                    "SELECT COALESCE(SUM(quantidade), 0) FROM leituras WHERE ts LIKE ?",
                    (dia + "%",),
                ).fetchone()
                return int(row[0] or 0)
        except sqlite3.Error:
            pass
        finally:
            if fechar_con and con is not None:
                try:
                    con.close()
                except Exception:
                    pass

        # fallback CSV
        self.log_dir, _, _ = carregar_caminhos()
        path = os.path.join(self.log_dir, f"log_{datetime.date.today()}.csv")
        if not os.path.isfile(path):
            return 0
        total = 0
        try:
            with open(path, mode="r", newline="", encoding="utf-8") as f:
                reader = csv.reader(f, delimiter=";")
                for row in reader:
                    if len(row) < 5 or (row[0] == "Data" and row[1] == "Operador"):
                        continue
                    if len(row) >= 9:
                        qty_col = 6
                    elif len(row) >= 6:
                        qty_col = 5
                    else:
                        qty_col = 4
                    try:
                        total += int(row[qty_col])
                    except ValueError:
                        total += 1
        except OSError:
            pass
        return total

    def _atualizar_resumo(self) -> None:
        self.text_resumo.configure(state=tk.NORMAL)
        self.text_resumo.delete(1.0, tk.END)
        for ref, qty in sorted(self.consumos.items()):
            self.text_resumo.insert(tk.END, f"  {ref}    {qty}\n")
        self.text_resumo.configure(state=tk.DISABLED)

        total_sessao = sum(self.consumos.values())
        total_dia = self._total_do_dia()
        n_refs = len(self.consumos)

        self.label_total.configure(text=f" Total do dia: {total_dia}")
        self.label_leituras_sessao.configure(text=f" Sessao: {total_sessao}")
        self.label_refs.configure(text=f" Ref. unicas: {n_refs}")

        if self.sessao_iniciada and self.inicio_sessao:
            delta = datetime.datetime.now() - self.inicio_sessao
            h, r = divmod(int(delta.total_seconds()), 3600)
            m, s = divmod(r, 60)
            self.label_duracao.configure(text=f" Duracao: {h:02d}:{m:02d}:{s:02d}")
        else:
            self.label_duracao.configure(text=" Duracao: 00:00:00")

    # -------------------- Edicao / Eliminar --------------------
    def _obter_leitura_selecionada(self, acao: str):
        sel = self.list_ultimas.curselection()
        if not sel:
            messagebox.showinfo("Nada selecionado", f"Selecione uma leitura na lista para {acao}.")
            return None, None, None

        idx = int(sel[0])
        items = list(self.ultimas_leituras)
        if idx >= len(items):
            return None, None, None

        item = items[idx]
        if not (isinstance(item, tuple) and len(item) >= 6 and isinstance(item[0], int)):
            messagebox.showwarning("Leitura invalida", f"Esta leitura nao pode ser {acao} (formato antigo em memoria).")
            return None, None, None

        return idx, item, items

    def _atualizar_comentario_csv(self, timestamp, ref, qty, lote, novo_comentario):
        if not self.logfile or not os.path.isfile(self.logfile):
            return

        self._garantir_cabecalho_csv_com_lote()
        with open(self.logfile, mode="r", newline="", encoding="utf-8") as f:
            rows = list(csv.reader(f, delimiter=";"))

        def linha_coincide(row):
            if len(row) < 5 or row[0] == "Data":
                return False
            if row[0] != timestamp:
                return False
            if len(row) >= 9:
                ok = row[4].strip() == ref and str(row[6]).strip() == str(qty)
                if ok and lote:
                    return row[8].strip() == lote
                return ok
            if len(row) >= 8:
                ok = row[4].strip() == ref and str(row[5]).strip() == str(qty)
                if ok and lote:
                    return row[7].strip() == lote
                return ok
            return row[3].strip() == ref and str(row[4]).strip() == str(qty)

        matches = [i for i in range(len(rows)) if linha_coincide(rows[i])]
        if not matches:
            return

        idx = matches[-1]
        row = rows[idx]

        if len(row) >= 9:
            row[7] = novo_comentario
        elif len(row) == 8:
            row[6] = novo_comentario
        elif len(row) == 7:
            row[6] = novo_comentario
        else:
            row.append(novo_comentario)

        rows[idx] = row

        with open(self.logfile, mode="w", newline="", encoding="utf-8") as f:
            w = csv.writer(f, delimiter=";")
            w.writerows(rows)

    def _editar_comentario_leitura(self) -> None:
        if not self.sessao_iniciada:
            messagebox.showwarning("Sessao nao iniciada", "Inicie uma sessao primeiro.")
            return

        idx, item, items = self._obter_leitura_selecionada("editar")
        if item is None:
            return

        row_id, ref, qty, timestamp, comentario_atual, lote = item[:6]
        descricao_item = item[6] if len(item) >= 7 else ""

        novo_comentario = simpledialog.askstring(
            "Editar comentario",
            f"Referencia: {ref}  Quantidade: {qty}\n\nComentario:",
            initialvalue=comentario_atual or "",
            parent=self.root,
        )
        if novo_comentario is None:
            return

        novo_comentario = novo_comentario.strip().replace("\n", " ").replace(";", ",")

        try:
            if self.db_con is None:
                raise sqlite3.Error("Ligacao DB nao inicializada.")
            self.db_con.execute("UPDATE leituras SET comentario = ? WHERE id = ?", (novo_comentario, row_id))
            self.db_con.commit()
        except sqlite3.Error as err:
            messagebox.showerror("Erro ao editar (DB)", str(err))
            return

        items[idx] = (row_id, ref, qty, timestamp, novo_comentario, lote, descricao_item)
        self.ultimas_leituras = items

        try:
            self._atualizar_comentario_csv(timestamp, ref, qty, lote, novo_comentario)
        except OSError:
            pass

        self._atualizar_ultimas()
        self.list_ultimas.selection_clear(0, tk.END)
        self.list_ultimas.selection_set(idx)
        self.list_ultimas.see(idx)

    def _eliminar_leitura(self) -> None:
        """Remove a leitura selecionada da sessao e apaga no SQLite (e tenta refletir no CSV)."""
        if not self.sessao_iniciada:
            messagebox.showwarning("Sessao nao iniciada", "Inicie uma sessao primeiro.")
            return

        idx, item, items = self._obter_leitura_selecionada("eliminar")
        if item is None:
            return

        row_id, ref, qty, timestamp, _, _ = item[:6]

        if not messagebox.askyesno("Eliminar leitura", f"Eliminar registo:\n  {ref}  {qty}\n\nConfirma?"):
            return

        # 1) DB delete (preciso)
        try:
            if self.db_con is None:
                raise sqlite3.Error("Ligação DB não inicializada.")
            self.db_con.execute("DELETE FROM leituras WHERE id = ?", (row_id,))
            self.db_con.commit()
        except sqlite3.Error as err:
            messagebox.showerror("Erro ao eliminar (DB)", str(err))
            return

        # 2) Atualizar consumos em memória
        self.consumos[ref] = self.consumos.get(ref, 0) - int(qty)
        if self.consumos[ref] <= 0:
            del self.consumos[ref]

        # 3) Remover da lista da sessao
        items.pop(idx)
        self.ultimas_leituras = items

        # 4) Tentar refletir no CSV (best-effort)
        # Nota: CSV não tem id, removemos por timestamp+ref+qty (se houver duplicados no mesmo segundo, pode remover a última ocorrência).
        try:
            self._garantir_cabecalho_csv_com_lote()
            with open(self.logfile, mode="r", newline="", encoding="utf-8") as f:
                rows = list(csv.reader(f, delimiter=";"))

            def linha_coincide(row):
                if len(row) < 5 or row[0] == "Data":
                    return False
                if row[0] != timestamp:
                    return False
                if len(row) >= 9:
                    return row[4].strip() == ref and str(row[6]).strip() == str(qty)
                if len(row) >= 6:
                    return row[4].strip() == ref and str(row[5]) == str(qty)
                if len(row) >= 5:
                    return row[3].strip() == ref and str(row[4]) == str(qty)
                return False

            matches = [i for i in range(len(rows)) if linha_coincide(rows[i])]
            if matches:
                rows.pop(matches[-1])

            with open(self.logfile, mode="w", newline="", encoding="utf-8") as f:
                w = csv.writer(f, delimiter=";")
                w.writerows(rows)
        except OSError:
            pass

        self._atualizar_ultimas()
        self._atualizar_resumo()

    # -------------------- Relatorios CSV (DB) --------------------
    def _parse_data_relatorio(self, valor: Any) -> datetime.date:
        texto = str(valor or "").strip()
        if not texto:
            raise ValueError("Data em falta.")
        return datetime.datetime.strptime(texto, "%Y-%m-%d").date()

    def _obter_conexao_db_relatorio(self) -> Tuple[sqlite3.Connection, bool]:
        if self.db_con is not None:
            return self.db_con, False

        self.log_dir, self.bom_path, self.db_dir = carregar_caminhos()
        caminho_db = db_path(self.db_dir)
        if not os.path.isfile(caminho_db):
            raise FileNotFoundError(f"Base de dados nao encontrada:\n{os.path.abspath(caminho_db)}")

        con = db_connect(caminho_db)
        return con, True

    def _gerar_relatorio_csv_db(self, data_ini: datetime.date, data_fim: datetime.date) -> Tuple[str, str, int]:
        inicio_str = data_ini.strftime("%Y-%m-%d")
        fim_str = data_fim.strftime("%Y-%m-%d")
        limite_superior = (data_fim + datetime.timedelta(days=1)).strftime("%Y-%m-%d")

        if inicio_str == fim_str:
            base_nome = f"relatorio_db_{inicio_str}"
        else:
            base_nome = f"relatorio_db_{inicio_str}_a_{fim_str}"

        self.log_dir, _, _ = carregar_caminhos()
        os.makedirs(self.log_dir, exist_ok=True)
        detalhe_path = os.path.join(self.log_dir, f"{base_nome}_detalhe.csv")
        totais_path = os.path.join(self.log_dir, f"{base_nome}_totais.csv")

        con = None
        fechar_con = False
        try:
            con, fechar_con = self._obter_conexao_db_relatorio()

            linhas_detalhe = con.execute(
                """SELECT ts, operador, projeto, turno, referencia, description, quantidade, comentario, lote
                   FROM leituras
                   WHERE ts >= ? AND ts < ?
                   ORDER BY ts ASC""",
                (inicio_str, limite_superior),
            ).fetchall()

            totais_ref = con.execute(
                """SELECT referencia, COALESCE(MAX(description), ''), SUM(quantidade) AS total
                   FROM leituras
                   WHERE ts >= ? AND ts < ?
                   GROUP BY referencia
                   ORDER BY referencia ASC""",
                (inicio_str, limite_superior),
            ).fetchall()
        finally:
            if fechar_con and con is not None:
                try:
                    con.close()
                except Exception:
                    pass

        with open(detalhe_path, mode="w", newline="", encoding="utf-8") as f:
            w = csv.writer(f, delimiter=";")
            w.writerow(["Data", "Operador", "Projeto", "Turno", "Referencia", "Description", "Quantidade", "Comentario", "Lote"])
            w.writerows(linhas_detalhe)

        with open(totais_path, mode="w", newline="", encoding="utf-8") as f:
            w = csv.writer(f, delimiter=";")
            w.writerow(["Referencia", "Description", "Total"])

            total_geral = 0
            for ref, descricao, total in totais_ref:
                total_int = int(total or 0)
                w.writerow([ref, descricao, total_int])
                total_geral += total_int

            w.writerow([])
            w.writerow(["TOTAL GERAL", "", total_geral])

        return detalhe_path, totais_path, len(linhas_detalhe)

    def _abrir_janela_exportacao_csv_db(self) -> None:
        janela = tk.Toplevel(self.root)
        janela.title("Exportar relatorios CSV (DB)")
        janela.configure(bg=CORES["branco"])
        janela.resizable(False, False)
        janela.transient(self.root)
        janela.grab_set()

        hoje_str = datetime.date.today().strftime("%Y-%m-%d")
        modo_var = tk.StringVar(value="dia")
        data_ini_var = tk.StringVar(value=hoje_str)
        data_fim_var = tk.StringVar(value=hoje_str)

        frame = tk.Frame(janela, bg=CORES["branco"], padx=14, pady=12)
        frame.pack(fill=tk.BOTH, expand=True)

        tk.Label(
            frame,
            text="Exportacao para CSV",
            font=("Segoe UI", 11, "bold"),
            fg=CORES["cinza_escuro"],
            bg=CORES["branco"],
        ).grid(row=0, column=0, columnspan=3, sticky=tk.W, pady=(0, 8))

        tk.Radiobutton(
            frame,
            text="Um dia",
            variable=modo_var,
            value="dia",
            bg=CORES["branco"],
            fg=CORES["cinza_escuro"],
            activebackground=CORES["branco"],
            activeforeground=CORES["cinza_escuro"],
        ).grid(row=1, column=0, sticky=tk.W)
        tk.Radiobutton(
            frame,
            text="Intervalo de dias",
            variable=modo_var,
            value="intervalo",
            bg=CORES["branco"],
            fg=CORES["cinza_escuro"],
            activebackground=CORES["branco"],
            activeforeground=CORES["cinza_escuro"],
        ).grid(row=1, column=1, columnspan=2, sticky=tk.W, padx=(12, 0))

        tk.Label(
            frame,
            text="Data inicial (YYYY-MM-DD):",
            font=("Segoe UI", 9),
            fg=CORES["cinza_escuro"],
            bg=CORES["branco"],
        ).grid(row=2, column=0, sticky=tk.W, pady=(10, 2))
        entry_data_ini = tk.Entry(frame, textvariable=data_ini_var, width=16, font=("Segoe UI", 10), relief=tk.SOLID, bd=1)
        entry_data_ini.grid(row=2, column=1, sticky=tk.W, pady=(10, 2))

        tk.Label(
            frame,
            text="Data final (YYYY-MM-DD):",
            font=("Segoe UI", 9),
            fg=CORES["cinza_escuro"],
            bg=CORES["branco"],
        ).grid(row=3, column=0, sticky=tk.W, pady=(6, 2))
        entry_data_fim = tk.Entry(frame, textvariable=data_fim_var, width=16, font=("Segoe UI", 10), relief=tk.SOLID, bd=1)
        entry_data_fim.grid(row=3, column=1, sticky=tk.W, pady=(6, 2))

        def atualizar_estado_data_fim(*_):
            if modo_var.get() == "dia":
                data_fim_var.set(data_ini_var.get().strip())
                entry_data_fim.configure(state=tk.DISABLED)
            else:
                entry_data_fim.configure(state=tk.NORMAL)

        def exportar():
            try:
                data_ini = self._parse_data_relatorio(data_ini_var.get())
                if modo_var.get() == "dia":
                    data_fim = data_ini
                else:
                    data_fim = self._parse_data_relatorio(data_fim_var.get())
                    if data_fim < data_ini:
                        messagebox.showwarning(
                            "Intervalo invalido",
                            "A data final nao pode ser anterior a data inicial.",
                            parent=janela,
                        )
                        return

                detalhe_path, totais_path, registos = self._gerar_relatorio_csv_db(data_ini, data_fim)
            except ValueError:
                messagebox.showwarning(
                    "Data invalida",
                    "Use o formato YYYY-MM-DD (ex: 2026-03-01).",
                    parent=janela,
                )
                return
            except FileNotFoundError as err:
                messagebox.showerror("Relatorio CSV (DB)", str(err), parent=janela)
                return
            except sqlite3.Error as err:
                messagebox.showerror("Relatorio CSV (DB)", f"Erro ao ler a base de dados:\n\n{err}", parent=janela)
                return
            except OSError as err:
                messagebox.showerror("Relatorio CSV (DB)", f"Nao foi possivel gravar os ficheiros CSV:\n\n{err}", parent=janela)
                return

            msg = (
                "Relatorios CSV gerados com sucesso.\n\n"
                f"Detalhe: {os.path.abspath(detalhe_path)}\n"
                f"Totais: {os.path.abspath(totais_path)}\n\n"
                f"Registos exportados: {registos}"
            )
            if registos == 0:
                msg += "\n\nNao existem leituras no periodo selecionado."
            messagebox.showinfo("Relatorio CSV (DB)", msg, parent=janela)
            janela.destroy()

        botoes = tk.Frame(frame, bg=CORES["branco"])
        botoes.grid(row=4, column=0, columnspan=3, sticky=tk.W, pady=(12, 0))
        tk.Button(
            botoes,
            text="Exportar",
            font=("Segoe UI", 10, "bold"),
            fg=CORES["branco"],
            bg=CORES["azul"],
            activebackground=CORES["azul_claro"],
            activeforeground=CORES["branco"],
            relief=tk.FLAT,
            padx=14,
            pady=4,
            cursor="hand2",
            command=exportar,
        ).pack(side=tk.LEFT)
        tk.Button(
            botoes,
            text="Fechar",
            font=("Segoe UI", 10),
            fg=CORES["cinza_escuro"],
            bg=CORES["fundo"],
            activebackground=CORES["painel_titulo"],
            activeforeground=CORES["cinza_escuro"],
            relief=tk.FLAT,
            padx=12,
            pady=4,
            cursor="hand2",
            command=janela.destroy,
        ).pack(side=tk.LEFT, padx=(8, 0))

        modo_var.trace_add("write", atualizar_estado_data_fim)
        data_ini_var.trace_add("write", atualizar_estado_data_fim)
        atualizar_estado_data_fim()

        janela.bind("<Return>", lambda e: exportar())
        janela.bind("<Escape>", lambda e: janela.destroy())
        entry_data_ini.focus_set()
        entry_data_ini.select_range(0, tk.END)

    # -------------------- Terminar / Export CSV --------------------
    def _terminar_sessao(self) -> None:
        if self.sessao_iniciada:
            # Export CSV do dia a partir do SQLite
            if self.db_con is not None:
                self._exportar_csv_do_dia()

            self.sessao_iniciada = False

            if self._timer_duracao:
                self.root.after_cancel(self._timer_duracao)
                self._timer_duracao = None

            self.entry_operador.configure(state="normal")
            self.combo_projeto.configure(state="readonly")
            self.combo_turno.configure(state="readonly")
            self.btn_iniciar.configure(state=tk.NORMAL)

            self.label_sessao.configure(text="Sessao terminada. Pode iniciar uma nova sessao.", fg=CORES["cinza_claro"])
            self.label_inicio_sessao.configure(text="Sessao iniciada a\n--:--:--")

            # fechar DB com segurança
            try:
                if self.db_con is not None:
                    self.db_con.close()
            except Exception:
                pass
            self.db_con = None
            self.db_path = None
            self.sessao_id = None

            self._atualizar_resumo()
            messagebox.showinfo("Sessao terminada", "Sessao terminada.")

        self.entry_referencia.focus_set()

    def _exportar_csv_do_dia(self) -> None:
        """Gera os relatorios CSV do dia (detalhe + totais por referencia) a partir do SQLite."""
        if self.db_con is None:
            messagebox.showwarning("Relatorio CSV (DB)", "Base de dados nao disponivel. Nao foi possivel exportar.")
            return

        hoje = datetime.date.today()
        try:
            detalhe_path, totais_path, registos = self._gerar_relatorio_csv_db(hoje, hoje)
        except FileNotFoundError as err:
            messagebox.showerror("Relatorio CSV (DB)", str(err))
            return
        except sqlite3.Error as err:
            messagebox.showerror("Relatorio CSV (DB)", f"Erro ao ler a base de dados:\n\n{err}")
            return
        except OSError as err:
            messagebox.showerror("Relatorio CSV (DB)", f"Nao foi possivel gravar os ficheiros CSV:\n\n{err}")
            return

        msg = (
            "Relatorios CSV do dia gerados com sucesso.\n\n"
            f"Detalhe: {os.path.abspath(detalhe_path)}\n"
            f"Totais: {os.path.abspath(totais_path)}\n\n"
            f"Registos exportados: {registos}"
        )
        if registos == 0:
            msg += "\n\nNao existem leituras para o dia atual."
        messagebox.showinfo("Relatorio CSV (DB)", msg)

    # -------------------- Run --------------------
    def run(self) -> None:
        self.root.mainloop()


if __name__ == "__main__":
    app = WarehouseApp()
    app.run()
