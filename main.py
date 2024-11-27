import sqlite3
import tkinter as tk
from tkinter import messagebox
from tkinter import ttk
import pandas as pd
import csv
from datetime import datetime, timedelta
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter

# Adaptadores personalizados para datetime
def adapt_datetime(dt):
    return dt.isoformat()

def convert_datetime(s):
    return datetime.fromisoformat(s.decode())

sqlite3.register_adapter(datetime, adapt_datetime)
sqlite3.register_converter("datetime", convert_datetime)


# Função para carregar dados da planilha para o banco de dados
def carregar_planilha_para_banco(planilha_path):
    df = pd.read_excel(planilha_path, sheet_name="Página1")

    conexao = sqlite3.connect("estoque_dental.db", detect_types=sqlite3.PARSE_DECLTYPES)
    cursor = conexao.cursor()

    # Tabela de insumos
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS insumos (
            codigo TEXT PRIMARY KEY,
            nome TEXT NOT NULL,
            quantidade INTEGER NOT NULL,
            validade TEXT,
            localizacao TEXT,
            observacao TEXT
        )
    ''')

    # Tabela de histórico de movimentações
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS historico (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            insumo_codigo TEXT NOT NULL,
            tipo TEXT NOT NULL,
            quantidade INTEGER NOT NULL,
            data datetime NOT NULL,
            FOREIGN KEY (insumo_codigo) REFERENCES insumos (codigo)
        )
    ''')

    # Inserir os dados da planilha na tabela de insumos
    for _, linha in df.iterrows():
        try:
            cursor.execute('''
                INSERT OR REPLACE INTO insumos (codigo, nome, quantidade, validade, localizacao, observacao)
                VALUES (?, ?, ?, ?, ?, ?)
            ''', (
                linha["CÓDIGO"],
                linha["ÍTEM"],
                int(linha["QUANTIDADE"]) if str(linha["QUANTIDADE"]).isdigit() else 0,
                str(linha["VALIDADE"]) if pd.notna(linha["VALIDADE"]) else "INDETERMINADO",
                linha["ESTANTE/PRATELEIRA"],
                linha["OBSERVAÇÃO"] if pd.notna(linha["OBSERVAÇÃO"]) else None
            ))
        except Exception as e:
            print(f"Erro ao inserir linha {linha}: {e}")

    conexao.commit()
    conexao.close()



# Função para exportar os dados do banco para a planilha Excel
def exportar_para_excel():
    conexao = sqlite3.connect("estoque_dental.db")
    cursor = conexao.cursor()
    
    # Buscar dados da tabela de insumos
    cursor.execute("SELECT * FROM insumos")
    dados = cursor.fetchall()
    
    # Criar um DataFrame com os dados
    df = pd.DataFrame(dados, columns=["CÓDIGO", "ÍTEM", "QUANTIDADE", "VALIDADE", "ESTANTE/PRATELEIRA", "OBSERVAÇÃO"])
    
    # Exportar para a planilha
    nome_arquivo = "ESTOQUE_ATUALIZADO.xlsx"
    df.to_excel(nome_arquivo, index=False, sheet_name="Página1")
    messagebox.showinfo("Sucesso", f"Planilha exportada: {nome_arquivo}")
    conexao.close()


# Função para converter as datas de validade corretamente (ignorando o tempo)
def converter_data(data):
    try:
        return datetime.strptime(data, "%Y-%m-%d %H:%M:%S").date()  # Ignora o tempo
    except ValueError:
        try:
            return datetime.strptime(data, "%Y-%m-%d").date()  # Apenas a data
        except ValueError:
            return None


# Função para centralizar a janela
def centralizar_janela(root):
    root.update_idletasks()
    largura_janela = root.winfo_width()
    altura_janela = root.winfo_height()
    largura_tela = root.winfo_screenwidth()
    altura_tela = root.winfo_screenheight()

    x = (largura_tela // 2) - (largura_janela // 2)
    y = (altura_tela // 2) - (altura_janela // 2)
    root.geometry(f'{largura_janela}x{altura_janela}+{x}+{y}')



# Função para registrar novos insumos manualmente
def tela_registrar_insumos(root):
    root.withdraw()
    janela = tk.Toplevel()
    janela.title("Registrar Novo Insumo")
    janela.geometry("800x600")

    centralizar_janela(janela)

    campos = [
        ("Código do Insumo", "codigo"),
        ("Nome do Insumo", "nome"),
        ("Quantidade Inicial", "quantidade"),
        ("Validade (DD/MM/AAAA)", "validade"),
        ("Localização", "localizacao"),
        ("Observação", "observacao")
    ]

    entradas = {}
    for texto, chave in campos:
        label = tk.Label(janela, text=texto)
        label.pack()
        entrada = tk.Entry(janela)
        entrada.pack()
        entradas[chave] = entrada

    def salvar_insumo():
        dados = {chave: entradas[chave].get().strip() for chave in entradas}

        if not dados["codigo"] or not dados["nome"] or not dados["quantidade"].isdigit():
            messagebox.showerror("Erro", "Preencha todos os campos obrigatórios corretamente!")
            return

        # Validar a data de validade
        if dados["validade"]:
            try:
                datetime.strptime(dados["validade"], "%d/%m/%Y")
            except ValueError:
                messagebox.showerror("Erro", "Formato de data inválido! Use DD/MM/AAAA.")
                return

        conexao = sqlite3.connect("estoque_dental.db")
        cursor = conexao.cursor()
        try:
            cursor.execute('''
                INSERT INTO insumos (codigo, nome, quantidade, validade, localizacao, observacao)
                VALUES (?, ?, ?, ?, ?, ?)
            ''', (
                dados["codigo"],
                dados["nome"],
                int(dados["quantidade"]),
                dados["validade"] if dados["validade"] else "INDETERMINADO",
                dados["localizacao"],
                dados["observacao"]
            ))
            conexao.commit()
            messagebox.showinfo("Sucesso", "Insumo cadastrado com sucesso!")
            exportar_para_excel()  # Atualizar a planilha Excel
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


# Função para editar um insumo
def editar_insumo(rowid, carregar_dados, conexao):
    cursor = conexao.cursor()
    
    cursor.execute("SELECT rowid, * FROM insumos WHERE rowid = ?", (rowid,))
    insumo = cursor.fetchone()
    
    if insumo:
        janela_editar = tk.Toplevel()
        janela_editar.title("Editar Insumo")
        janela_editar.geometry("400x400")
        centralizar_janela(janela_editar)

        campos = [
            ("Código do Insumo", "codigo", insumo[1]),
            ("Nome do Insumo", "nome", insumo[2]),
            ("Quantidade", "quantidade", insumo[3]),
            ("Validade (DD/MM/AAAA)", "validade", insumo[4]),
            ("Localização", "localizacao", insumo[5]),
            ("Observação", "observacao", insumo[6])
        ]

        entradas = {}
        for texto, chave, valor in campos:
            label = tk.Label(janela_editar, text=texto)
            label.pack()
            entrada = tk.Entry(janela_editar)
            entrada.insert(0, str(valor))  # Convertendo o valor para string
            entrada.pack()
            entradas[chave] = entrada

        def salvar_edicao():
            dados = {chave: entradas[chave].get().strip() for chave in entradas}

            if not dados["nome"] or not dados["quantidade"].isdigit():
                messagebox.showerror("Erro", "Preencha todos os campos obrigatórios corretamente!")
                return

            # Validar a data de validade
            if dados["validade"]:
                try:
                    datetime.strptime(dados["validade"], "%d/%m/%Y")
                except ValueError:
                    messagebox.showerror("Erro", "Formato de data inválido! Use DD/MM/AAAA.")
                    return

            try:
                cursor.execute('''
                    UPDATE insumos
                    SET codigo = ?, nome = ?, quantidade = ?, validade = ?, localizacao = ?, observacao = ?
                    WHERE rowid = ?
                ''', (
                    dados["codigo"] if dados["codigo"] else None,
                    dados["nome"],
                    int(dados["quantidade"]),
                    dados["validade"] if dados["validade"] else "INDETERMINADO",
                    dados["localizacao"],
                    dados["observacao"],
                    rowid
                ))
                conexao.commit()
                messagebox.showinfo("Sucesso", "Insumo editado com sucesso!")
                exportar_para_excel()  # Atualizar a planilha Excel
                carregar_dados()  # Atualizar os dados na tabela de monitoramento
            except sqlite3.Error as e:
                messagebox.showerror("Erro de Banco de Dados", str(e))
            finally:
                janela_editar.destroy()
                carregar_dados()  # Recarregar dados na tabela de monitoramento

        btn_salvar = tk.Button(janela_editar, text="Salvar", command=salvar_edicao)
        btn_salvar.pack(pady=10)
        btn_cancelar = tk.Button(janela_editar, text="Cancelar", command=janela_editar.destroy)
        btn_cancelar.pack(pady=5)
    else:
        messagebox.showerror("Erro", "Insumo não encontrado!")

# Função para excluir um insumo
def excluir_insumo(rowid, carregar_dados, conexao):
    if messagebox.askokcancel("Confirmação", "Tem certeza de que deseja excluir este insumo?"):
        cursor = conexao.cursor()

        try:
            cursor.execute("DELETE FROM insumos WHERE rowid = ?", (rowid,))
            conexao.commit()
            messagebox.showinfo("Sucesso", "Insumo excluído com sucesso!")
            exportar_para_excel()  # Atualizar a planilha Excel
            carregar_dados()  # Atualizar os dados na tabela de monitoramento
        except sqlite3.Error as e:
            messagebox.showerror("Erro de Banco de Dados", str(e))
        finally:
            carregar_dados()  # Recarregar dados na tabela de monitoramento

# Função para monitorar estoque (atualizado)
def tela_monitorar_estoque(root):
    root.withdraw()
    janela = tk.Toplevel()
    janela.title("Monitorar Estoque")
    janela.geometry("800x600")

    centralizar_janela(janela)

    titulo = tk.Label(janela, text="Estoque Atual", font=("Arial", 18, "bold"))
    titulo.pack(pady=10)

    frame_tabela = tk.Frame(janela)
    frame_tabela.pack(fill=tk.BOTH, expand=True)

    colunas = ("Código", "Nome", "Quantidade", "Validade", "Localização", "Observação", "rowid")
    tabela = ttk.Treeview(frame_tabela, columns=colunas, show="headings")
    for col in colunas[:-1]:  # Excluindo rowid da exibição
        tabela.heading(col, text=col)
        tabela.column(col, width=120)
    tabela.pack(fill=tk.BOTH, expand=True, pady=10)

    conexao = sqlite3.connect("estoque_dental.db")

    def carregar_dados():
        tabela.delete(*tabela.get_children())
        cursor = conexao.cursor()
        cursor.execute("SELECT rowid, * FROM insumos")
        for linha in cursor.fetchall():
            quantidade = linha[3]
            if quantidade >= 20:
                cor = "#d4edda"  # Verde Claro
                tag = "verde"
            elif quantidade > 10:
                cor = "#fff3cd"  # Amarelo Claro
                tag = "amarelo"
            else:
                cor = "#f8d7da"  # Vermelho Claro
                tag = "vermelho"
            tabela.insert("", tk.END, values=linha[1:] + (linha[0],), tags=(tag,))  # rowid no final
            tabela.tag_configure(tag, background=cor)

    carregar_dados()

    def editar_selecionado():
        item_selecionado = tabela.selection()
        if item_selecionado:
            rowid = tabela.item(item_selecionado)["values"][-1]
            editar_insumo(rowid, carregar_dados, conexao)
        else:
            messagebox.showerror("Erro", "Selecione um insumo para editar!")

    def excluir_selecionado():
        item_selecionado = tabela.selection()
        if item_selecionado:
            rowid = tabela.item(item_selecionado)["values"][-1]
            excluir_insumo(rowid, carregar_dados, conexao)
        else:
            messagebox.showerror("Erro", "Selecione um insumo para excluir!")

    btn_editar = tk.Button(janela, text="Editar Insumo", command=editar_selecionado)
    btn_editar.pack(pady=5)
    btn_excluir = tk.Button(janela, text="Excluir Insumo", command=excluir_selecionado)
    btn_excluir.pack(pady=5)
    btn_voltar = tk.Button(janela, text="Voltar", command=lambda: [conexao.close(), root.deiconify()])
    btn_voltar.pack(pady=10)

    # Filtro de busca
    filtro_label = tk.Label(janela, text="Buscar Insumo:")
    filtro_label.pack()
    filtro_entrada = tk.Entry(janela)
    filtro_entrada.pack()

    def filtrar_dados():
        busca = filtro_entrada.get().lower()
        tabela.delete(*tabela.get_children())
        cursor = conexao.cursor()
        cursor.execute("SELECT rowid, * FROM insumos WHERE LOWER(nome) LIKE ?", ('%' + busca + '%',))
        for linha in cursor.fetchall():
            quantidade = linha[3]
            if quantidade >= 20:
                cor = "#d4edda"  # Verde Claro
                tag = "verde"
            elif quantidade > 10:
                cor = "#fff3cd"  # Amarelo Claro
                tag = "amarelo"
            else:
                cor = "#f8d7da"  # Vermelho Claro
                tag = "vermelho"
            tabela.insert("", tk.END, values=linha[1:] + (linha[0],), tags=(tag,))  # rowid no final
            tabela.tag_configure(tag, background=cor)

    btn_filtrar = tk.Button(janela, text="Filtrar", command=filtrar_dados)
    btn_filtrar.pack(pady=5)

    btn_voltar.pack(pady=10)

    janela.protocol("WM_DELETE_WINDOW", lambda: [conexao.close(), root.deiconify()])


# Função para movimentação de estoque
def tela_movimentacao_estoque(root):
    root.withdraw()
    janela = tk.Toplevel()
    janela.title("Movimentação de Estoque")
    janela.geometry("800x600")

    centralizar_janela(janela)

    titulo = tk.Label(janela, text="Registrar Movimentação de Estoque", font=("Arial", 18, "bold"))
    titulo.pack(pady=10)

    frame_selecao = tk.Frame(janela)
    frame_selecao.pack(pady=10)
    label_insumo = tk.Label(frame_selecao, text="Selecionar Insumo:")
    label_insumo.pack(side=tk.LEFT, padx=5)
    combo_insumos = ttk.Combobox(frame_selecao, state="readonly", width=30)
    combo_insumos.pack(side=tk.LEFT, padx=5)

    def carregar_insumos():
        conexao = sqlite3.connect("estoque_dental.db")
        cursor = conexao.cursor()
        cursor.execute("SELECT codigo, nome FROM insumos")
        insumos = cursor.fetchall()
        conexao.close()
        combo_insumos['values'] = [f"{codigo} - {nome}" for codigo, nome in insumos]

    carregar_insumos()

    frame_quantidade = tk.Frame(janela)
    frame_quantidade.pack(pady=10)
    label_quantidade = tk.Label(frame_quantidade, text="Quantidade:")
    label_quantidade.pack(side=tk.LEFT, padx=5)
    entrada_quantidade = tk.Entry(frame_quantidade, width=10)
    entrada_quantidade.pack(side=tk.LEFT, padx=5)

    # Registrar entrada
    def registrar_entrada():
        if not combo_insumos.get() or not entrada_quantidade.get().isdigit():
            messagebox.showerror("Erro", "Preencha todos os campos corretamente!")
            return

        quantidade = int(entrada_quantidade.get())
        codigo = combo_insumos.get().split(" - ")[0]

        conexao = sqlite3.connect("estoque_dental.db")
        cursor = conexao.cursor()
        cursor.execute("UPDATE insumos SET quantidade = quantidade + ? WHERE codigo = ?", (quantidade, codigo))
        cursor.execute('''
            INSERT INTO historico (insumo_codigo, tipo, quantidade, data)
            VALUES (?, ?, ?, ?)
        ''', (codigo, "Entrada", quantidade, datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
        conexao.commit()
        conexao.close()
        messagebox.showinfo("Sucesso", "Entrada registrada com sucesso!")
        carregar_insumos()

    # Registrar saída
    def registrar_saida():
        if not combo_insumos.get() or not entrada_quantidade.get().isdigit():
            messagebox.showerror("Erro", "Preencha todos os campos corretamente!")
            return

        quantidade = int(entrada_quantidade.get())
        codigo = combo_insumos.get().split(" - ")[0]

        conexao = sqlite3.connect("estoque_dental.db")
        cursor = conexao.cursor()
        cursor.execute("SELECT quantidade FROM insumos WHERE codigo = ?", (codigo,))
        quantidade_atual = cursor.fetchone()[0]

        if quantidade > quantidade_atual:
            messagebox.showerror("Erro", "Quantidade em estoque insuficiente!")
            conexao.close()
            return

        cursor.execute("UPDATE insumos SET quantidade = quantidade - ? WHERE codigo = ?", (quantidade, codigo))
        cursor.execute('''
            INSERT INTO historico (insumo_codigo, tipo, quantidade, data)
            VALUES (?, ?, ?, ?)
        ''', (codigo, "Saída", quantidade, datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
        conexao.commit()
        conexao.close()
        messagebox.showinfo("Sucesso", "Saída registrada com sucesso!")
        carregar_insumos()

    btn_entrada = tk.Button(janela, text="Registrar Entrada", command=registrar_entrada)
    btn_entrada.pack(pady=5)
    btn_saida = tk.Button(janela, text="Registrar Saída", command=registrar_saida)
    btn_saida.pack(pady=5)

    btn_voltar = tk.Button(janela, text="Voltar", command=lambda: [janela.destroy(), root.deiconify()])
    btn_voltar.pack(pady=10)


# Função para histórico de movimentações
def tela_historico(root):
    root.withdraw()
    janela = tk.Toplevel()
    janela.title("Histórico de Movimentações")
    janela.geometry("800x600")

    centralizar_janela(janela)

    titulo = tk.Label(janela, text="Histórico de Movimentações", font=("Arial", 18, "bold"))
    titulo.pack(pady=10)

    frame_tabela = tk.Frame(janela)
    frame_tabela.pack(fill=tk.BOTH, expand=True)

    colunas = ("ID", "Código do Insumo", "Nome do Insumo", "Tipo", "Quantidade", "Data")
    tabela = ttk.Treeview(frame_tabela, columns=colunas, show="headings")
    for col in colunas:
        tabela.heading(col, text=col)
        tabela.column(col, width=120)
    tabela.pack(fill=tk.BOTH, expand=True, pady=10)

    def carregar_historico():
        tabela.delete(*tabela.get_children())
        conexao = sqlite3.connect("estoque_dental.db")
        cursor = conexao.cursor()
        cursor.execute('''
            SELECT historico.id, historico.insumo_codigo, insumos.nome, historico.tipo, historico.quantidade, historico.data 
            FROM historico 
            JOIN insumos ON historico.insumo_codigo = insumos.codigo 
            ORDER BY historico.data DESC
        ''')
        for linha in cursor.fetchall():
            tabela.insert("", tk.END, values=linha)
        conexao.close()

    carregar_historico()

    btn_voltar = tk.Button(janela, text="Voltar", command=lambda: [janela.destroy(), root.deiconify()])
    btn_voltar.pack(pady=10)


# Função para verificar validade
def tela_alertas_validade(root):
    root.withdraw()
    janela = tk.Toplevel()
    janela.title("Alertas de Validade")
    janela.geometry("800x600")

    centralizar_janela(janela)

    titulo = tk.Label(janela, text="Itens Vencidos ou Próximos da Validade", font=("Arial", 18, "bold"))
    titulo.pack(pady=10)

    frame_tabela = tk.Frame(janela)
    frame_tabela.pack(fill=tk.BOTH, expand=True)

    colunas = ("Código", "Nome", "Quantidade", "Validade", "Localização")
    tabela = ttk.Treeview(frame_tabela, columns=colunas, show="headings")
    for col in colunas:
        tabela.heading(col, text=col)
        tabela.column(col, width=120)
    tabela.pack(fill=tk.BOTH, expand=True, pady=10)

    def carregar_alertas():
        tabela.delete(*tabela.get_children())
        hoje = datetime.now().date()  # Converte a data de hoje para o formato date (sem tempo)
        alerta_prazo = hoje + timedelta(days=30)  # Itens próximos a vencer em 30 dias
        conexao = sqlite3.connect("estoque_dental.db")
        cursor = conexao.cursor()
        cursor.execute("SELECT * FROM insumos WHERE validade != 'INDETERMINADO'")
        for linha in cursor.fetchall():
            validade = converter_data(linha[3])  # Convertendo a data de validade
            if validade and (validade < hoje or validade <= alerta_prazo):
                tabela.insert("", tk.END, values=linha[:5])
        conexao.close()

    carregar_alertas()

    btn_voltar = tk.Button(janela, text="Voltar", command=lambda: [janela.destroy(), root.deiconify()])
    btn_voltar.pack(pady=10)


def gerar_relatorio_pdf():
    conexao = sqlite3.connect("estoque_dental.db")
    cursor = conexao.cursor()
    
    cursor.execute('''
        SELECT historico.id, historico.insumo_codigo, insumos.nome, historico.tipo, historico.quantidade, historico.data 
        FROM historico 
        JOIN insumos ON historico.insumo_codigo = insumos.codigo 
        ORDER BY historico.data DESC
    ''')
    
    movimentacoes = cursor.fetchall()
    conexao.close()

    nome_arquivo = "relatorio_movimentacoes.pdf"
    c = canvas.Canvas(nome_arquivo, pagesize=letter)
    largura, altura = letter

    c.setFont("Helvetica-Bold", 14)
    c.drawString(30, altura - 30, "Relatório de Movimentações de Estoque")

    c.setFont("Helvetica", 12)
    linha = altura - 60
    for movimentacao in movimentacoes:
        texto = f"ID: {movimentacao[0]} | Código: {movimentacao[1]} | Nome: {movimentacao[2]} | Tipo: {movimentacao[3]} | Quantidade: {movimentacao[4]} | Data: {movimentacao[5]}"
        c.drawString(30, linha, texto)
        linha -= 20
        if linha < 40:
            c.showPage()
            c.setFont("Helvetica", 12)
            linha = altura - 40

    c.save()
    messagebox.showinfo("Sucesso", f"Relatório gerado: {nome_arquivo}")


def tela_gerar_relatorio(root):
    root.withdraw()
    janela = tk.Toplevel()
    janela.title("Gerar Relatório")
    janela.geometry("800x600")

    centralizar_janela(janela)

    label = ttk.Label(janela, text="Deseja gerar o relatório de todas as movimentações em PDF?")
    label.pack(pady=20)

    def confirmar_gerar_relatorio():
        gerar_relatorio_pdf()
        janela.destroy()
        root.deiconify()

    btn_sim = ttk.Button(janela, text="Sim", command=confirmar_gerar_relatorio)
    btn_sim.pack(pady=5)

    btn_nao = ttk.Button(janela, text="Não", command=lambda: [janela.destroy(), root.deiconify()])
    btn_nao.pack(pady=5)



# Aplicar estilo ao aplicativo
def aplicar_estilos():
    estilo = ttk.Style()
    estilo.theme_use('clam')

    # Configuração de fontes
    estilo.configure('TLabel', font=('Arial', 12))
    estilo.configure('TButton', font=('Arial', 12), padding=6)
    estilo.configure('Treeview', font=('Arial', 10), rowheight=25)
    estilo.configure('TEntry', font=('Arial', 12))

    # Estilo dos botões
    estilo.configure('TButton', background='#4CAF50', foreground='white', borderwidth=1)
    estilo.map('TButton', background=[('active', '#45a049')])

# Menu principal
def iniciar_aplicativo(planilha_path):
    carregar_planilha_para_banco(planilha_path)

    root = tk.Tk()
    aplicar_estilos()  # Aplicar os estilos
    root.title("Controle de Estoque - Clínica Odontológica")
    root.geometry("800x600")

    centralizar_janela(root)

    titulo = ttk.Label(root, text="Menu Principal", font=("Arial", 18, "bold"))
    titulo.pack(pady=20)

    botoes = [
        ("Registrar Novo Insumo", lambda: tela_registrar_insumos(root)),
        ("Monitorar Estoque", lambda: tela_monitorar_estoque(root)),
        ("Movimentar Estoque", lambda: tela_movimentacao_estoque(root)),
        ("Histórico de Movimentações", lambda: tela_historico(root)),
        ("Alertas de Validade", lambda: tela_alertas_validade(root)),
        ("Gerar Relatório", lambda: tela_gerar_relatorio(root))
    ]

    for texto, comando in botoes:
        btn = ttk.Button(root, text=texto, width=30, command=comando)
        btn.pack(pady=5)

    btn_sair = ttk.Button(root, text="Sair", width=30, command=root.quit)
    btn_sair.pack(pady=5)

    root.mainloop()


# Caminho relativo da planilha
planilha_path = "ESTOQUE.xlsx"
iniciar_aplicativo(planilha_path)