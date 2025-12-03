import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
from sqlalchemy.orm import declarative_base, sessionmaker
from sqlalchemy import create_engine, text
from datetime import date, datetime
from urllib.parse import quote_plus

# ==========================================
# CONFIGURAÇÕES DO BANCO
# ==========================================
BD_SERVER = "192.168.38.28,1433"
BD_NAME = "facturas"
BD_USER = "paulo"
BD_PASSWORD = "loucoste9309323"
BD_DRIVER = "ODBC Driver 17 for SQL Server"

driver_encoded = quote_plus(BD_DRIVER)

DATABASE_URL = (
    f"mssql+pyodbc://{BD_USER}:{BD_PASSWORD}@{BD_SERVER}/{BD_NAME}"
    f"?driver={driver_encoded}"
)

engine = create_engine(
    DATABASE_URL,
    fast_executemany=True,
)
SessionLocal = sessionmaker(bind=engine)
Base = declarative_base()

# ==========================================
# TABELA DE MÊS DE NÚMERO PARA TEXTO
# ==========================================
dados6 = [
    ['1', 'Janeiro'],
    ['2', 'Fevereiro'],
    ['3', 'Março'],
    ['4', 'Abril'],
    ['5', 'Maio'],
    ['6', 'Junho'],
    ['7', 'Julho'],
    ['8', 'Agosto'],
    ['9', 'Setembro'],
    ['10', 'Outubro'],
    ['11', 'Novembro'],
    ['12', 'Dezembro'],
]
refmes = pd.DataFrame(dados6, columns=['Me', 'Mês'])
refmes.set_index('Me', inplace=True)

ordem_meses = [
    "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
    "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"
]


# ==========================================
# APLICAÇÃO TKINTER
# ==========================================

class ValorizadaApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Valorizada - Aplicativo Desktop")
        self.geometry("1200x700")

        # Estado
        self.df_importado = None
        self.df_filtrado = None

        # Layout principal: menu lateral + área de conteúdo
        self.columnconfigure(1, weight=1)
        self.rowconfigure(0, weight=1)

        self.create_sidebar()
        self.create_content_area()

        # Tela inicial
        self.show_inicio()

    # ---------------- SIDEBAR ----------------
    def create_sidebar(self):
        self.sidebar = tk.Frame(self, bg="#F0F0F0", width=200)
        self.sidebar.grid(row=0, column=0, sticky="nswe")
        self.sidebar.grid_propagate(False)

        tk.Label(self.sidebar, text="Menu", bg="#F0F0F0",
                 font=("Segoe UI", 14, "bold")).pack(pady=10)

        btn_inicio = tk.Button(self.sidebar, text="Início",
                               command=self.show_inicio, width=20)
        btn_inicio.pack(pady=5)

        btn_import = tk.Button(self.sidebar, text="Importação",
                               command=self.show_importacao, width=20)
        btn_import.pack(pady=5)

        btn_extracao = tk.Button(self.sidebar, text="Extração Valorizada",
                                 command=self.show_extracao, width=20)
        btn_extracao.pack(pady=5)

    # ------------- ÁREA DE CONTEÚDO ----------
    def create_content_area(self):
        self.content = tk.Frame(self)
        self.content.grid(row=0, column=1, sticky="nswe")
        self.content.rowconfigure(0, weight=1)
        self.content.columnconfigure(0, weight=1)

    def clear_content(self):
        for widget in self.content.winfo_children():
            widget.destroy()

    # ------------- TELA INÍCIO ----------
    def show_inicio(self):
        self.clear_content()
        frame = tk.Frame(self.content)
        frame.grid(sticky="nsew", padx=20, pady=20)

        tk.Label(frame, text="Importação de Dados Valorizada",
                 font=("Segoe UI", 18, "bold")).pack(pady=10)
        tk.Label(frame, text="Bem-vindo ao sistema de importação de dados.",
                 font=("Segoe UI", 12)).pack(pady=10)

    # ------------- TELA IMPORTAÇÃO ----------
    def show_importacao(self):
        self.clear_content()
        frame = tk.Frame(self.content)
        frame.grid(sticky="nsew", padx=20, pady=20)
        frame.columnconfigure(0, weight=1)
        frame.columnconfigure(1, weight=1)

        tk.Label(frame, text="Importação de Dados Valorizada",
                 font=("Segoe UI", 16, "bold")).grid(row=0, column=0, columnspan=2, pady=10, sticky="w")

        # Data
        tk.Label(frame, text="Definir Data (AAAA-MM-DD):").grid(row=1, column=0, sticky="w", pady=5)
        self.entry_data = tk.Entry(frame, width=15)
        self.entry_data.grid(row=1, column=1, sticky="w", pady=5)
        self.entry_data.insert(0, "2020-01-01")

        # Botão selecionar arquivo
        tk.Button(frame, text="Selecionar ficheiro Excel",
                  command=self.selecionar_arquivo).grid(row=2, column=0, columnspan=2, pady=10, sticky="w")

        # Pré-visualização
        tk.Label(frame, text="Pré-visualização:", font=("Segoe UI", 12, "bold")).grid(
            row=3, column=0, columnspan=2, sticky="w"
        )

        self.preview_tree = ttk.Treeview(frame, height=15)
        self.preview_tree.grid(row=4, column=0, columnspan=2, sticky="nsew", pady=5)

        frame.rowconfigure(4, weight=1)

        # Botão guardar
        tk.Button(frame, text="Guardar Facturação",
                  command=self.guardar_fatura).grid(row=5, column=0, columnspan=2, pady=10)

    def selecionar_arquivo(self):
        path = filedialog.askopenfilename(
            title="Selecionar Ficheiro Excel",
            filetypes=[("Excel", "*.xlsx")]
        )
        if not path:
            return

        try:
            df = pd.read_excel(path, engine="openpyxl")
            self.df_importado = df
            self.mostrar_preview(df)
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao ler o Excel:\n{e}")

    def mostrar_preview(self, df: pd.DataFrame):
        # Limpar Treeview
        for col in self.preview_tree.get_children():
            self.preview_tree.delete(col)
        self.preview_tree["columns"] = list(df.columns)
        self.preview_tree["show"] = "headings"

        for col in df.columns:
            self.preview_tree.heading(col, text=col)
            self.preview_tree.column(col, width=100)

        # Mostrar apenas primeiras linhas
        for _, row in df.head(50).iterrows():
            self.preview_tree.insert("", "end", values=list(row))

    def guardar_fatura(self):
        if self.df_importado is None:
            messagebox.showwarning("Aviso", "Nenhum ficheiro foi importado.")
            return

        data_str = self.entry_data.get().strip()
        try:
            data_obj = datetime.strptime(data_str, "%Y-%m-%d").date()
        except ValueError:
            messagebox.showerror("Erro", "Data inválida. Use o formato AAAA-MM-DD.")
            return

        data_df = pd.DataFrame({"Data": [data_obj]})
        df_final = pd.concat([self.df_importado, data_df], axis=1)
        df_final["Data"] = df_final["Data"].ffill()

        try:
            df_final.to_sql("Valorizada", con=engine, if_exists='append', index=False)
            messagebox.showinfo("Sucesso", "Informação carregada com sucesso!")
        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro na importação:\n{e}")

    # ------------- TELA EXTRAÇÃO ----------
    def show_extracao(self):
        self.clear_content()
        frame = tk.Frame(self.content)
        frame.grid(sticky="nsew", padx=20, pady=20)
        frame.rowconfigure(3, weight=1)
        frame.columnconfigure(1, weight=1)

        tk.Label(frame, text="Extração Valorizada",
                 font=("Segoe UI", 16, "bold")).grid(row=0, column=0, columnspan=3, pady=10, sticky="w")

        # Carregar dados da BD
        try:
            query = "SELECT * FROM Valorizada"
            valorizada = pd.read_sql(query, engine)
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao carregar dados da BD:\n{e}")
            return

        if valorizada.empty:
            messagebox.showinfo("Informação", "A tabela Valorizada está vazia.")
            return

        valorizada["Data"] = pd.to_datetime(valorizada["Data"])
        valorizada["Ano"] = valorizada["Data"].dt.year.astype(str)
        valorizada["Me"] = valorizada["Data"].dt.month.astype(str)
        valorizada["Data"] = valorizada["Data"].dt.date
        valorizada2 = pd.merge(valorizada, refmes, on="Me", how="left")

        # Guardar para uso nos handlers
        self.valorizada2 = valorizada2

        # Filtros
        tk.Label(frame, text="Ano(s):").grid(row=1, column=0, sticky="nw")
        self.list_anos = tk.Listbox(frame, selectmode="multiple", height=5, exportselection=False)
        self.list_anos.grid(row=2, column=0, sticky="nw", pady=5)

        anos_unicos = sorted(valorizada2["Ano"].unique())
        for a in anos_unicos:
            self.list_anos.insert(tk.END, a)

        tk.Label(frame, text="Mês(es):").grid(row=1, column=1, sticky="nw")
        self.list_meses = tk.Listbox(frame, selectmode="multiple", height=8, exportselection=False)
        self.list_meses.grid(row=2, column=1, sticky="nw", pady=5)

        for m in ordem_meses:
            self.list_meses.insert(tk.END, m)

        tk.Button(frame, text="Aplicar Filtro", command=self.aplicar_filtro).grid(row=2, column=2, padx=10, sticky="n")

        # Tabela de resultados
        self.tree_extracao = ttk.Treeview(frame)
        self.tree_extracao.grid(row=3, column=0, columnspan=3, sticky="nsew", pady=10)

        # Botão de exportação
        tk.Button(frame, text="Exportar CSV", command=self.exportar_csv).grid(row=4, column=0, columnspan=3, pady=5)

    def aplicar_filtro(self):
        if not hasattr(self, "valorizada2"):
            return

        df = self.valorizada2.copy()

        # Anos selecionados
        anos_sel_idx = self.list_anos.curselection()
        anos_sel = [self.list_anos.get(i) for i in anos_sel_idx]

        if anos_sel:
            df = df[df["Ano"].isin(anos_sel)]

        # Meses selecionados
        meses_sel_idx = self.list_meses.curselection()
        meses_sel = [self.list_meses.get(i) for i in meses_sel_idx]

        if meses_sel:
            df = df[df["Mês"].isin(meses_sel)]

        if df.empty:
            messagebox.showinfo("Informação", "Nenhum registo com os filtros selecionados.")
            return

        df["Itinerario"] = pd.to_numeric(df["Itinerario"], errors="coerce").astype("Int64")

        df = df.loc[:, [
            "Unidade", "Qtd_Docs", "Lote", "PIN", "Ordem", "CIL", "Nome",
            "Localidade", "Morada", "Aviso_Acesso", "Nr_Documento", "Valor",
            "Rota", "Roteiro", "Itinerario", "Regiao", "UC"
        ]]

        self.df_filtrado = df

        # Mostrar na Treeview
        for item in self.tree_extracao.get_children():
            self.tree_extracao.delete(item)

        self.tree_extracao["columns"] = list(df.columns)
        self.tree_extracao["show"] = "headings"

        for col in df.columns:
            self.tree_extracao.heading(col, text=col)
            self.tree_extracao.column(col, width=110)

        for _, row in df.iterrows():
            self.tree_extracao.insert("", "end", values=list(row))

    def exportar_csv(self):
        if self.df_filtrado is None or self.df_filtrado.empty:
            messagebox.showwarning("Aviso", "Não há dados filtrados para exportar.")
            return

        path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV", "*.csv")],
            title="Guardar ficheiro CSV"
        )
        if not path:
            return

        try:
            self.df_filtrado.to_csv(path, sep=";", decimal=",", index=False, encoding="utf-8-sig")
            messagebox.showinfo("Sucesso", f"Ficheiro guardado em:\n{path}")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao guardar CSV:\n{e}")


if __name__ == "__main__":
    app = ValorizadaApp()
    app.mainloop()
