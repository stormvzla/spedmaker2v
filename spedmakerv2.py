import json
import os
from tkinter import *
from tkinter import messagebox

# Nome do arquivo de configuração
CONFIG_FILE = "empresas_salvas.json"

def salvar_dados_empresa(nome_empresa, cnpj, ie):
    """Salva os dados da empresa em um arquivo JSON."""
    # Carregar empresas existentes
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, "r", encoding="utf-8") as arquivo:
            empresas = json.load(arquivo)
    else:
        empresas = []

    # Verificar se a empresa já existe
    for empresa in empresas:
        if empresa["cnpj"] == cnpj:
            messagebox.showinfo("Aviso", "Empresa já cadastrada.")
            return

    # Adicionar nova empresa
    empresas.append({
        "nome_empresa": nome_empresa,
        "cnpj": cnpj,
        "ie": ie
    })

    # Salvar no arquivo
    with open(CONFIG_FILE, "w", encoding="utf-8") as arquivo:
        json.dump(empresas, arquivo, ensure_ascii=False, indent=4)

    messagebox.showinfo("Sucesso", "Dados da empresa salvos com sucesso!")
    atualizar_dropdown_empresas()  # Atualiza o dropdown após salvar

def carregar_empresas_salvas():
    """Carrega as empresas salvas do arquivo JSON."""
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, "r", encoding="utf-8") as arquivo:
            return json.load(arquivo)
    return []

def atualizar_dropdown_empresas():
    """Atualiza o dropdown de empresas com as empresas salvas."""
    empresas_salvas = carregar_empresas_salvas()
    if empresas_salvas:
        nomes_empresas = [empresa["nome_empresa"] for empresa in empresas_salvas]
        menu_empresas.delete(0, "end")  # Limpa o menu atual
        for nome in nomes_empresas:
            menu_empresas.add_command(label=nome, command=lambda value=nome: combo_empresas.set(value))
        combo_empresas.set("Selecione uma empresa")
    else:
        combo_empresas.set("Nenhuma empresa salva")

def preencher_campos_empresa(event=None):
    """Preenche os campos com os dados da empresa selecionada no dropdown."""
    empresa_selecionada = combo_empresas.get()
    if not empresa_selecionada or empresa_selecionada == "Selecione uma empresa" or empresa_selecionada == "Nenhuma empresa salva":
        return

    # Limpar campos antes de preencher
    entry_nome_empresa.delete(0, END)
    entry_cnpj.delete(0, END)
    entry_ie.delete(0, END)

    # Preencher campos com os dados da empresa selecionada
    empresas_salvas = carregar_empresas_salvas()
    for empresa in empresas_salvas:
        if empresa["nome_empresa"] == empresa_selecionada:
            entry_nome_empresa.insert(0, empresa["nome_empresa"])
            entry_cnpj.insert(0, empresa["cnpj"])
            entry_ie.insert(0, empresa["ie"])
            break

def gerar_arquivo_sped(nome_empresa, cnpj, ie, mes, ano):
    # Nome do arquivo
    nome_arquivo = f"SPED_ICMS_{mes:02d}{ano}.txt"
    
    # Lista para armazenar os registros
    registros = []

    # Função auxiliar para formatar as linhas corretamente
    def formatar_linha(linha):
        return f"|{linha}|"

    # Bloco 0 - Abertura, Identificação e Referências
    data_inicio = f"01{mes:02d}{ano}"  # Primeiro dia do mês no formato DDMMAAAA
    data_fim = f"31{mes:02d}{ano}"    # Último dia do mês (considerando meses com 31 dias)
    registros.append(formatar_linha("0000|019|0|{data_inicio}|{data_fim}|{nome_empresa}|{cnpj}||RJ|{ie}|3304557|||A|1".format(
        nome_empresa=nome_empresa,
        cnpj=cnpj,
        ie=ie,
        data_inicio=data_inicio,
        data_fim=data_fim
    )))
    registros.append(formatar_linha("0001|0"))  # Indicador de movimento: 0 = Sem operações
    registros.append(formatar_linha("0005|{nome_empresa}|20090907|AV RIO BRANCO|1|SAL 901|CENTRO|2131508041||GIGAPECAS2024@GMAIL.COM".format(
        nome_empresa=nome_empresa
    )))
    registros.append(formatar_linha("0100|{nome_empresa}|09285232760|16113416712|{cnpj}||||||||GIGAPECAS2024@GMAIL.COM|3304557".format(
        nome_empresa=nome_empresa,
        cnpj=cnpj
    )))
    registros.append(formatar_linha("0990|5"))  # Total de linhas do Bloco 0

    # Bloco B - Outras Declarações
    registros.append(formatar_linha("B001|1"))  # Indicador de movimento: 1 = Com operações
    registros.append(formatar_linha("B990|2"))  # Total de linhas do Bloco B

    # Bloco C - Documentos Fiscais
    registros.append(formatar_linha("C001|1"))  # Indicador de movimento: 1 = Com operações
    registros.append(formatar_linha("C990|2"))  # Total de linhas do Bloco C

    # Bloco D - Operações com ICMS-ST
    registros.append(formatar_linha("D001|1"))  # Indicador de movimento: 1 = Com operações
    registros.append(formatar_linha("D990|2"))  # Total de linhas do Bloco D

    # Bloco E - Apuração do ICMS
    registros.append(formatar_linha("E001|0"))  # Indicador de movimento: 0 = Sem operações
    registros.append(formatar_linha("E100|{data_inicio}|{data_fim}".format(data_inicio=data_inicio, data_fim=data_fim)))
    registros.append(formatar_linha("E110|0|0|0|0|0|0|0|0|0|0|0|0|0|0"))
    registros.append(formatar_linha("E990|4"))  # Total de linhas do Bloco E

    # Bloco G - Controle de Créditos de ICMS
    registros.append(formatar_linha("G001|1"))  # Indicador de movimento: 1 = Com operações
    registros.append(formatar_linha("G990|2"))  # Total de linhas do Bloco G

    # Bloco H - Inventário de Estoque
    registros.append(formatar_linha("H001|1"))  # Indicador de movimento: 1 = Com operações
    registros.append(formatar_linha("H990|2"))  # Total de linhas do Bloco H

    # Bloco K - Produção e Estoque
    registros.append(formatar_linha("K001|1"))  # Indicador de movimento: 1 = Com operações
    registros.append(formatar_linha("K990|2"))  # Total de linhas do Bloco K

    # Bloco 1 - Informações Complementares
    registros.append(formatar_linha("1001|0"))  # Indicador de movimento: 0 = Sem operações
    registros.append(formatar_linha("1010|N|N|N|N|N|N|N|N|N|N|N|N|N"))
    registros.append(formatar_linha("1990|3"))  # Total de linhas do Bloco 1

    # Bloco 9 - Encerramento
    registros.append(formatar_linha("9001|0"))  # Indicador de movimento: 0 = Sem operações
    # Contagem de registros
    registros_contagem = [
        ("0000", 1), ("0001", 1), ("0005", 1), ("0100", 1), ("0990", 1),
        ("B001", 1), ("B990", 1),
        ("C001", 1), ("C990", 1),
        ("D001", 1), ("D990", 1),
        ("E001", 1), ("E100", 1), ("E110", 1), ("E990", 1),
        ("G001", 1), ("G990", 1),
        ("H001", 1), ("H990", 1),
        ("K001", 1), ("K990", 1),
        ("1001", 1), ("1010", 1), ("1990", 1),
        ("9001", 1), ("9990", 1), ("9999", 1), ("9900", 28)
    ]
    for registro, quantidade in registros_contagem:
        registros.append(formatar_linha(f"9900|{registro}|{quantidade}"))
    registros.append(formatar_linha("9990|31"))  # Total de linhas do Bloco 9
    registros.append(formatar_linha("9999|55"))  # Total geral de linhas

    # Escrever no arquivo
    with open(nome_arquivo, "w", encoding="utf-8") as arquivo:
        for registro in registros:
            arquivo.write(registro + "\n")

def gerar_todos_os_arquivos_do_ano():
    # Obter valores dos campos
    nome_empresa = entry_nome_empresa.get().strip()
    cnpj = entry_cnpj.get().strip()
    ie = entry_ie.get().strip()
    ano = entry_ano.get().strip()

    # Validar campos
    if not nome_empresa or not cnpj or not ie or not ano:
        messagebox.showerror("Erro", "Preencha todos os campos.")
        return

    try:
        ano = int(ano)
        if ano <= 0:
            raise ValueError
    except ValueError:
        messagebox.showerror("Erro", "Ano inválido.")
        return

    # Gerar arquivos para todos os meses do ano
    for mes in range(1, 13):
        gerar_arquivo_sped(nome_empresa, cnpj, ie, mes, ano)

    messagebox.showinfo("Sucesso", f"Todos os arquivos do ano {ano} foram gerados com sucesso!")

def on_gerar_arquivo():
    # Obter valores dos campos
    nome_empresa = entry_nome_empresa.get().strip()
    cnpj = entry_cnpj.get().strip()
    ie = entry_ie.get().strip()
    mes = combo_mes.get()
    ano = entry_ano.get().strip()

    # Validar campos
    if not nome_empresa or not cnpj or not ie or not mes or not ano:
        messagebox.showerror("Erro", "Preencha todos os campos.")
        return

    try:
        mes = int(mes)
        ano = int(ano)
        if not (1 <= mes <= 12) or ano <= 0:
            raise ValueError
    except ValueError:
        messagebox.showerror("Erro", "Mês ou ano inválido.")
        return

    # Gerar arquivo SPED
    gerar_arquivo_sped(nome_empresa, cnpj, ie, mes, ano)

# Criar janela principal
root = Tk()
root.title("Gerador de Arquivos SPED ICMS")
root.geometry("450x400")

# Carregar empresas salvas
empresas_salvas = carregar_empresas_salvas()

# Dropdown para empresas salvas
Label(root, text="Selecione uma Empresa:").pack(pady=5)
combo_empresas = StringVar(root)
combo_empresas.set("Selecione uma empresa")  # Valor padrão
dropdown_empresas = OptionMenu(root, combo_empresas, "Selecione uma empresa")
dropdown_empresas.pack(pady=5)
menu_empresas = dropdown_empresas["menu"]
atualizar_dropdown_empresas()  # Inicializa o dropdown com empresas salvas
combo_empresas.trace("w", lambda *args: preencher_campos_empresa())  # Atualiza campos ao selecionar

# Labels e Entradas
Label(root, text="Nome da Empresa:").pack(pady=5)
entry_nome_empresa = Entry(root, width=40)
entry_nome_empresa.pack(pady=5)

Label(root, text="CNPJ (sem pontos ou traços):").pack(pady=5)
entry_cnpj = Entry(root, width=40)
entry_cnpj.pack(pady=5)

Label(root, text="Inscrição Estadual:").pack(pady=5)
entry_ie = Entry(root, width=40)
entry_ie.pack(pady=5)

# Botão para salvar empresa
Button(root, text="Salvar Empresa", command=lambda: salvar_dados_empresa(
    entry_nome_empresa.get().strip(),
    entry_cnpj.get().strip(),
    entry_ie.get().strip()
)).pack(pady=10)

Label(root, text="Mês:").pack(pady=5)
combo_mes = StringVar(root)
combo_mes.set("1")  # Valor padrão
OptionMenu(root, combo_mes, *range(1, 13)).pack(pady=5)

Label(root, text="Ano:").pack(pady=5)
entry_ano = Entry(root, width=10)
entry_ano.pack(pady=5)

# Botão para gerar arquivo único
Button(root, text="Gerar Arquivo SPED", command=on_gerar_arquivo).pack(pady=10)

# Botão para gerar todos os arquivos do ano
Button(root, text="Gerar Todos os Arquivos do Ano", command=gerar_todos_os_arquivos_do_ano).pack(pady=10)

# Executar loop principal
root.mainloop()