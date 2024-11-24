import sqlite3
import tkinter as tk
from tkinter import messagebox
from tkinter import ttk
import csv
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
import datetime


# Inicializar banco de dados
def inicializar_banco():
    conexao = sqlite3.connect("estoque_dental.db")
    cursor = conexao.cursor()

    # Criar tabela de insumos
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS insumos (
            codigo INTEGER PRIMARY KEY AUTOINCREMENT,
            nome TEXT NOT NULL,
            quantidade INTEGER NOT NULL,
            categoria TEXT NOT NULL,
            unidade TEXT NOT NULL,
            fornecedor TEXT,
            quantidade_minima INTEGER NOT NULL
        )
    ''')
    
    # Criar tabela de histórico
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS historico (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            insumo_codigo INTEGER,
            tipo TEXT NOT NULL,
            quantidade INTEGER NOT NULL,
            data TEXT NOT NULL,
            FOREIGN KEY (insumo_codigo) REFERENCES insumos(codigo)
        )
    ''')
    
    conexao.commit()
    conexao.close()


# Função para validar data no formato DD/MM/AAAA
def validar_data(data):
    try:
        datetime.datetime.strptime(data, "%d/%m/%Y")
        return True
    except ValueError:
        return False


# Criar menu dinâmico
def criar_menu(root, botoes):
    for texto, comando in botoes:
        btn = tk.Button(root, text=texto, font=("Arial", 14), width=30, command=comando)
        btn.pack(pady=5)


# Tela para registrar insumos
def tela_registrar_insumos(root):
    root.withdraw()  # Ocultar janela principal
    janela = tk.Toplevel()
    janela.title("Registrar Novos Insumos")
    janela.geometry("400x400")

    campos = [
        ("Nome do Insumo", "nome"),
        ("Quantidade Inicial", "quantidade"),
        ("Categoria", "categoria"),
        ("Unidade de Medida", "unidade"),
        ("Fornecedor", "fornecedor"),
        ("Quantidade Mínima", "quantidade_minima")
    ]

    entradas = {}
    for texto, chave in campos:
        label = tk.Label(janela, text=texto)
        label.pack()
        entrada = tk.Entry(janela)
        entrada.pack()
        entradas[chave] = entrada

    # Função para salvar insumo no banco
    def salvar_insumo():
        dados = {chave: entradas[chave].get().strip() for chave in entradas}

        if not dados["nome"] or not dados["quantidade"].isdigit() or not dados["quantidade_minima"].isdigit():
            messagebox.showerror("Erro", "Preencha todos os campos obrigatórios corretamente!")
            return

        conexao = sqlite3.connect("estoque_dental.db")
        cursor = conexao.cursor()
        try:
            cursor.execute('''
                INSERT INTO insumos (nome, quantidade, categoria, unidade, fornecedor, quantidade_minima)
                VALUES (?, ?, ?, ?, ?, ?)
            ''', (dados["nome"], int(dados["quantidade"]), dados["categoria"],
                  dados["unidade"], dados["fornecedor"], int(dados["quantidade_minima"])))
            conexao.commit()
            messagebox.showinfo("Sucesso", "Insumo cadastrado com sucesso!")
        except sqlite3.Error as e:
            messagebox.showerror("Erro de Banco de Dados", str(e))
        finally:
            conexao.close()

        janela.destroy()
        root.deiconify()

    btn_salvar = tk.Button(janela, text="Salvar", command=salvar_insumo)
    btn_salvar.pack(pady=10)

    btn_cancelar = tk.Button(janela, text="Cancelar", command=lambda: [janela.destroy(), root.deiconify()])
    btn_cancelar.pack(pady=5)


# Tela para monitorar estoque
def tela_monitorar_estoque(root):
    root.withdraw()
    janela = tk.Toplevel()
    janela.title("Monitorar Estoque")
    janela.geometry("800x600")

    titulo = tk.Label(janela, text="Estoque Atual", font=("Arial", 18, "bold"))
    titulo.pack(pady=10)

    frame_tabela = tk.Frame(janela)
    frame_tabela.pack(fill=tk.BOTH, expand=True)

    colunas = ("Código", "Nome", "Quantidade", "Categoria", "Unidade", "Fornecedor", "Mínimo")
    tabela = ttk.Treeview(frame_tabela, columns=colunas, show="headings")
    for col in colunas:
        tabela.heading(col, text=col)
        tabela.column(col, width=100)
    tabela.pack(fill=tk.BOTH, expand=True, pady=10)

    def carregar_dados():
        tabela.delete(*tabela.get_children())
        conexao = sqlite3.connect("estoque_dental.db")
        cursor = conexao.cursor()
        try:
            cursor.execute("SELECT * FROM insumos")
            for linha in cursor.fetchall():
                codigo, nome, quantidade, categoria, unidade, fornecedor, minimo = linha
                cor = "red" if quantidade <= minimo else "black"
                tabela.insert("", tk.END, values=(codigo, nome, quantidade, categoria, unidade, fornecedor, minimo),
                              tags=(cor,))
            tabela.tag_configure("red", foreground="red")
        except sqlite3.Error as e:
            messagebox.showerror("Erro de Banco de Dados", str(e))
        finally:
            conexao.close()

    carregar_dados()

    # Filtro de busca
    filtro_label = tk.Label(janela, text="Buscar Insumo:")
    filtro_label.pack()
    filtro_entrada = tk.Entry(janela)
    filtro_entrada.pack()

    def filtrar_dados():
        busca = filtro_entrada.get().lower()
        tabela.delete(*tabela.get_children())
        conexao = sqlite3.connect("estoque_dental.db")
        cursor = conexao.cursor()
        try:
            cursor.execute("SELECT * FROM insumos WHERE LOWER(nome) LIKE ?", ('%' + busca + '%',))
            for linha in cursor.fetchall():
                codigo, nome, quantidade, categoria, unidade, fornecedor, minimo = linha
                cor = "red" if quantidade <= minimo else "black"
                tabela.insert("", tk.END, values=(codigo, nome, quantidade, categoria, unidade, fornecedor, minimo),
                              tags=(cor,))
            tabela.tag_configure("red", foreground="red")
        except sqlite3.Error as e:
            messagebox.showerror("Erro de Banco de Dados", str(e))
        finally:
            conexao.close()

    btn_filtrar = tk.Button(janela, text="Filtrar", command=filtrar_dados)
    btn_filtrar.pack(pady=5)

    btn_voltar = tk.Button(janela, text="Voltar", command=lambda: [janela.destroy(), root.deiconify()])
    btn_voltar.pack(pady=10)

# Tela para movimentação de estoque
def tela_movimentacao_estoque(root):
    root.withdraw()
    janela = tk.Toplevel()
    janela.title("Movimentação de Estoque")
    janela.geometry("600x600")

    titulo = tk.Label(janela, text="Registrar Movimentação de Estoque", font=("Arial", 18, "bold"))
    titulo.pack(pady=10)

    # Seleção de insumos
    frame_selecao = tk.Frame(janela)
    frame_selecao.pack(pady=10)
    label_insumo = tk.Label(frame_selecao, text="Selecionar Insumo:")
    label_insumo.pack(side=tk.LEFT, padx=5)
    combo_insumos = ttk.Combobox(frame_selecao, state="readonly", width=30)
    combo_insumos.pack(side=tk.LEFT, padx=5)

    def carregar_insumos():
        conexao = sqlite3.connect("estoque_dental.db")
        cursor = conexao.cursor()
        try:
            cursor.execute("SELECT codigo, nome FROM insumos")
            insumos = cursor.fetchall()
            combo_insumos['values'] = [f"{codigo} - {nome}" for codigo, nome in insumos]
        except sqlite3.Error as e:
            messagebox.showerror("Erro de Banco de Dados", str(e))
        finally:
            conexao.close()

    carregar_insumos()

    # Campos para entrada de dados
    frame_quantidade = tk.Frame(janela)
    frame_quantidade.pack(pady=10)

    label_quantidade = tk.Label(frame_quantidade, text="Quantidade:")
    label_quantidade.pack(side=tk.LEFT, padx=5)
    entrada_quantidade = tk.Entry(frame_quantidade, width=10)
    entrada_quantidade.pack(side=tk.LEFT, padx=5)

    label_data = tk.Label(frame_quantidade, text="Data (DD/MM/AAAA):")
    label_data.pack(side=tk.LEFT, padx=5)
    entrada_data = tk.Entry(frame_quantidade, width=15)
    entrada_data.pack(side=tk.LEFT, padx=5)

    # Registrar entrada de insumo
    def registrar_entrada():
        if not combo_insumos.get() or not entrada_quantidade.get().isdigit() or not validar_data(entrada_data.get()):
            messagebox.showerror("Erro", "Preencha todos os campos corretamente!")
            return

        codigo = int(combo_insumos.get().split(" - ")[0])
        quantidade = int(entrada_quantidade.get())
        data = entrada_data.get()

        conexao = sqlite3.connect("estoque_dental.db")
        cursor = conexao.cursor()
        try:
            cursor.execute("UPDATE insumos SET quantidade = quantidade + ? WHERE codigo = ?", (quantidade, codigo))
            cursor.execute('''
                INSERT INTO historico (insumo_codigo, tipo, quantidade, data)
                VALUES (?, "Entrada", ?, ?)
            ''', (codigo, quantidade, data))
            conexao.commit()
            messagebox.showinfo("Sucesso", "Entrada registrada com sucesso!")
            carregar_insumos()
        except sqlite3.Error as e:
            messagebox.showerror("Erro de Banco de Dados", str(e))
        finally:
            conexao.close()

    # Registrar saída de insumo
    def registrar_saida():
        if not combo_insumos.get() or not entrada_quantidade.get().isdigit() or not validar_data(entrada_data.get()):
            messagebox.showerror("Erro", "Preencha todos os campos corretamente!")
            return

        codigo = int(combo_insumos.get().split(" - ")[0])
        quantidade = int(entrada_quantidade.get())
        data = entrada_data.get()

        conexao = sqlite3.connect("estoque_dental.db")
        cursor = conexao.cursor()
        try:
            cursor.execute("SELECT quantidade FROM insumos WHERE codigo = ?", (codigo,))
            quantidade_atual = cursor.fetchone()[0]
            if quantidade > quantidade_atual:
                messagebox.showerror("Erro", "Quantidade em estoque insuficiente!")
                return

            cursor.execute("UPDATE insumos SET quantidade = quantidade - ? WHERE codigo = ?", (quantidade, codigo))
            cursor.execute('''
                INSERT INTO historico (insumo_codigo, tipo, quantidade, data)
                VALUES (?, "Saída", ?, ?)
            ''', (codigo, quantidade, data))
            conexao.commit()
            messagebox.showinfo("Sucesso", "Saída registrada com sucesso!")
            carregar_insumos()
        except sqlite3.Error as e:
            messagebox.showerror("Erro de Banco de Dados", str(e))
        finally:
            conexao.close()

    btn_entrada = tk.Button(janela, text="Registrar Entrada", command=registrar_entrada)
    btn_entrada.pack(pady=5)

    btn_saida = tk.Button(janela, text="Registrar Saída", command=registrar_saida)
    btn_saida.pack(pady=5)

    btn_voltar = tk.Button(janela, text="Voltar", command=lambda: [janela.destroy(), root.deiconify()])
    btn_voltar.pack(pady=10)


# Tela para relatórios
def tela_relatorios(root):
    root.withdraw()
    janela = tk.Toplevel()
    janela.title("Relatórios")
    janela.geometry("600x500")

    titulo = tk.Label(janela, text="Relatórios", font=("Arial", 18, "bold"))
    titulo.pack(pady=10)

    # Gerar relatório em CSV
    def gerar_csv():
        conexao = sqlite3.connect("estoque_dental.db")
        cursor = conexao.cursor()
        try:
            cursor.execute('''
                SELECT insumos.nome, historico.tipo, historico.quantidade, historico.data
                FROM historico
                INNER JOIN insumos ON historico.insumo_codigo = insumos.codigo
            ''')
            dados = cursor.fetchall()
            nome_arquivo = "relatorio_estoque.csv"
            with open(nome_arquivo, mode="w", newline="", encoding="utf-8") as arquivo_csv:
                escritor = csv.writer(arquivo_csv)
                escritor.writerow(["Nome do Insumo", "Tipo", "Quantidade", "Data"])
                escritor.writerows(dados)
            messagebox.showinfo("Sucesso", f"Relatório CSV gerado: {nome_arquivo}")
        except sqlite3.Error as e:
            messagebox.showerror("Erro de Banco de Dados", str(e))
        finally:
            conexao.close()

    # Gerar relatório em PDF
    def gerar_pdf():
        conexao = sqlite3.connect("estoque_dental.db")
        cursor = conexao.cursor()
        try:
            cursor.execute('''
                SELECT insumos.nome, historico.tipo, historico.quantidade, historico.data
                FROM historico
                INNER JOIN insumos ON historico.insumo_codigo = insumos.codigo
            ''')
            dados = cursor.fetchall()
            nome_arquivo = "relatorio_estoque.pdf"
            pdf = canvas.Canvas(nome_arquivo, pagesize=letter)
            pdf.setFont("Helvetica-Bold", 16)
            pdf.drawString(200, 750, "Relatório de Estoque")
            pdf.setFont("Helvetica", 12)

            y = 700
            for nome, tipo, quantidade, data in dados:
                if y < 50:
                    pdf.showPage()
                    pdf.setFont("Helvetica", 12)
                    y = 750
                pdf.drawString(50, y, nome)
                pdf.drawString(200, y, tipo)
                pdf.drawString(300, y, str(quantidade))
                pdf.drawString(400, y, data)
                y -= 20

            pdf.save()
            messagebox.showinfo("Sucesso", f"Relatório PDF gerado: {nome_arquivo}")
        except sqlite3.Error as e:
            messagebox.showerror("Erro de Banco de Dados", str(e))
        finally:
            conexao.close()

    btn_csv = tk.Button(janela, text="Gerar CSV", command=gerar_csv)
    btn_csv.pack(pady=10)

    btn_pdf = tk.Button(janela, text="Gerar PDF", command=gerar_pdf)
    btn_pdf.pack(pady=10)

    btn_voltar = tk.Button(janela, text="Voltar", command=lambda: [janela.destroy(), root.deiconify()])
    btn_voltar.pack(pady=20)


def iniciar_aplicativo():
    inicializar_banco()
    root = tk.Tk()
    root.title("Controle de Estoque - Clínica Odontológica")
    root.geometry("600x400")

    titulo = tk.Label(root, text="Menu Principal", font=("Arial", 18, "bold"))
    titulo.pack(pady=20)

    botoes = [
        ("Registrar Novos Insumos", lambda: tela_registrar_insumos(root)),
        ("Monitorar Estoque", lambda: tela_monitorar_estoque(root)),
        ("Movimentar Estoque", lambda: tela_movimentacao_estoque(root)),
        ("Gerar Relatórios", lambda: tela_relatorios(root))
    ]
    criar_menu(root, botoes)

    root.mainloop()


# Iniciar o programa
iniciar_aplicativo()
