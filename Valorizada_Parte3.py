import customtkinter as ctk
from tkinter import filedialog, messagebox
from tkinter import ttk
import pandas as pd
from sqlalchemy import create_engine, text
from sqlalchemy.orm import sessionmaker, declarative_base
from urllib.parse import quote_plus
from datetime import datetime

# ===========================================================
# CONFIGURAÇÃO DO BANCO
# ===========================================================

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

engine = create_engine(DATABASE_URL, fast_executemany=True)
SessionLocal = sessionmaker(bind=engine)
Base = declarative_base()

# ===========================================================
# REFERÊNCIA DOS MESES
# ===========================================================
dados6 = [
    ['1', 'Janeiro'], ['2', 'Fevereiro'], ['3', 'Março'], ['4', 'Abril'],
    ['5', 'Maio'], ['6', 'Junho'], ['7', 'Julho'], ['8', 'Agosto'],
    ['9', 'Setembro'], ['10', 'Outubro'], ['11', 'Novembro'], ['12', 'Dezembro'],
]

refmes = pd.DataFrame(dados6, columns=['Me', 'Mês'])
refmes.set_index('Me', inplace=True)

ordem_meses = [
    "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", 
    "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"
]

# ===========================================================
# INTERFACE CUSTOMTKINTER
# ===========================================================

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")


class ValorizadaApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Valorizada - Aplicativo Desktop")
        self.geometry("1300x750")

        # Frames principais
        self.sidebar = ctk.CTkFrame(self, width=220, corner_radius=10)
        self.sidebar.pack(side="left", fill="y", padx=10, pady=10)

        self.content = ctk.CTkFrame(self, corner_radius=10)
        self.content.pack(side="right", fill="both", expand=True, padx=10, pady=10)

        # Criar menu lateral
        self.menu_buttons()

        # Estado do aplicativo
        self.df_importado = None
        self.df_filtrado = None

        self.show_inicio()

    # =====================================================
    # MENU LATERAL
    # =====================================================
    def menu_buttons(self):
        title = ctk.CTkLabel(self.sidebar, text="Menu", font=("Segoe UI", 18, "bold"))
        title.pack(pady=15)

        ctk.CTkButton(self.sidebar, text="Início", width=200,
                       command=self.show_inicio).pack(pady=5)

        ctk.CTkButton(self.sidebar, text="Importação", width=200,
                       command=self.show_importacao).pack(pady=5)

        ctk.CTkButton(self.sidebar, text="Extração Valorizada", width=200,
                       command=self.show_extracao).pack(pady=5)

    # =====================================================
    # LIMPAR ÁREA PRINCIPAL
    # =====================================================
    def clear_content(self):
        for widget in self.content.winfo_children():
            widget.destroy()

    # =====================================================
    # TELA INÍCIO
    # =====================================================
    def show_inicio(self):
        self.clear_content()
        title = ctk.CTkLabel(self.content, text="Sistema de Importação Valorizada",
                             font=("Segoe UI", 24, "bold"))
        title.pack(pady=20)

        ctk.CTkLabel(self.content, text="Bem-vindo ao sistema.",
                     font=("Segoe UI", 14)).pack(pady=10)

    # =====================================================
    # IMPORTAÇÃO DE DADOS
    # =====================================================
    def show_importacao(self):
        self.clear_content()

        ctk.CTkLabel(self.content, text="Importação de Dados Valorizada",
                     font=("Segoe UI", 20, "bold")).pack(pady=10)

        frame = ctk.CTkFrame(self.content, corner_radius=10)
        frame.pack(fill="both", expand=True, padx=10, pady=10)

        # DATA
        ctk.CTkLabel(frame, text="Data (AAAA-MM-DD):").pack(pady=5)
        self.entry_data = ctk.CTkEntry(frame, width=150)
        self.entry_data.pack()
        self.entry_data.insert(0, "2020-01-01")

        # BOTÃO PARA ESCOLHER EXCEL
        ctk.CTkButton(frame, text="Selecionar Ficheiro Excel",
                      command=self.selecionar_arquivo).pack(pady=20)

        # PREVIEW
        self.tree_preview = ttk.Treeview(frame, height=15)
        self.tree_preview.pack(fill="both", expand=True, pady=10)

        # BOTÃO GUARDAR
        ctk.CTkButton(frame, text="Guardar Facturação",
                      command=self.guardar_fatura).pack(pady=20)

    def selecionar_arquivo(self):
        path = filedialog.askopenfilename(
            title="Selecionar ficheiro", filetypes=[("Excel", "*.xlsx")]
        )
        if not path:
            return

        try:
            df = pd.read_excel(path)
            self.df_importado = df
            self.preencher_tree(self.tree_preview, df)
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao carregar Excel:\n{e}")

    # Preencher tabela
    def preencher_tree(self, tree, df):
        tree.delete(*tree.get_children())
        tree["columns"] = df.columns.tolist()
        tree["show"] = "headings"

        for col in df.columns:
            tree.heading(col, text=col)
            tree.column(col, width=120)

        for _, row in df.iterrows():
            tree.insert("", "end", values=list(row))

    def guardar_fatura(self):
        if self.df_importado is None:
            messagebox.showwarning("Aviso", "Nenhum ficheiro importado.")
            return

        try:
            data = datetime.strptime(self.entry_data.get(), "%Y-%m-%d").date()
        except:
            messagebox.showerror("Erro", "Data inválida.")
            return

        df_final = self.df_importado.copy()
        df_final["Data"] = data

        try:
            df_final.to_sql("Valorizada", con=engine, if_exists='append', index=False)
            messagebox.showinfo("Sucesso", "Facturação carregada!")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao guardar na base de dados:\n{e}")

    # =====================================================
    # EXTRAÇÃO
    # =====================================================
    def show_extracao(self):
        self.clear_content()

        ctk.CTkLabel(self.content, text="Extração Valorizada",
                     font=("Segoe UI", 20, "bold")).pack(pady=10)

        df = pd.read_sql("SELECT * FROM Valorizada", engine)

        df["Data"] = pd.to_datetime(df["Data"])
        df["Ano"] = df["Data"].dt.year.astype(str)
        df["Me"] = df["Data"].dt.month.astype(str)

        df2 = pd.merge(df, refmes, on="Me", how="left")

        self.df_base = df2

        # FRAME DE FILTROS
        filtro = ctk.CTkFrame(self.content)
        filtro.pack(fill="x", padx=10, pady=10)

        # Ano
        ctk.CTkLabel(filtro, text="Ano:").grid(row=0, column=0, padx=10)
        anos = sorted(df2["Ano"].unique())
        self.combo_ano = ctk.CTkComboBox(filtro, values=anos)
        self.combo_ano.grid(row=0, column=1, padx=10)

        # Mês
        ctk.CTkLabel(filtro, text="Mês:").grid(row=0, column=2, padx=10)
        self.combo_mes = ctk.CTkComboBox(filtro, values=ordem_meses)
        self.combo_mes.grid(row=0, column=3, padx=10)

        ctk.CTkButton(filtro, text="Aplicar Filtro",
                      command=self.aplicar_filtro).grid(row=0, column=4, padx=10)
        
        ctk.CTkButton(filtro, text="Limpar Filtro",
              fg_color="gray20",
              hover_color="gray30",
              command=self.limpar_filtro).grid(row=0, column=5, padx=10)

        # Tabela de resultados
        self.tree_extracao = ttk.Treeview(self.content, height=20)
        self.tree_extracao.pack(fill="both", expand=True, padx=10, pady=10)

        # Botão CSV
        ctk.CTkButton(self.content, text="Exportar CSV",
                      command=self.exportar_csv).pack(pady=10)

    def aplicar_filtro(self):
        ano = self.combo_ano.get()
        mes = self.combo_mes.get()

        df = self.df_base.copy()
        df = df[df["Ano"] == ano]
        df = df[df["Mês"] == mes]

        df["Itinerario"] = pd.to_numeric(df["Itinerario"], errors="coerce").astype("Int64")

        # Guardar o dataframe completo (com colunas internas)
        self.df_filtrado = df

        # Remover colunas internas antes de exibir
        df_exibir = df.drop(columns=["Data", "Ano", "Me", "Mês"])

        # Guardar o dataframe a exportar
        self.df_exibir = df_exibir

        # Mostrar tabela
        self.preencher_tree(self.tree_extracao, df_exibir)
    
    def limpar_filtro(self):
        # Limpar seleções
        self.combo_ano.set("")
        self.combo_mes.set("")

        # Limpar árvores
        for item in self.tree_extracao.get_children():
            self.tree_extracao.delete(item)

        # Limpar memória das filtragens
        self.df_filtrado = None
        self.df_exibir = None

    def exportar_csv(self):
        if not hasattr(self, "df_exibir") or self.df_exibir.empty:
            messagebox.showinfo("Aviso", "Nenhum dado para exportar.")
            return

        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[
                ("Excel .xlsx", "*.xlsx"),
                ("CSV .csv", "*.csv")
            ],
            title="Guardar ficheiro"
        )

        if not path:
            return

        try:
            if path.endswith(".csv"):
                self.df_exibir.to_csv(path, sep=";", decimal=",", index=False, encoding="utf-8-sig")
            else:
                self.df_exibir.to_excel(path, index=False, engine="openpyxl")

            messagebox.showinfo("Sucesso", f"Ficheiro guardado em:\n{path}")

        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao guardar ficheiro:\n{e}")


# EXECUTAR A APLICAÇÃO
if __name__ == "__main__":
    app = ValorizadaApp()
    app.mainloop()
