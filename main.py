import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
from datetime import datetime, timedelta
import openpyxl
import time

def selecionar_arquivo():
    filename = filedialog.askopenfilename(filetypes=[("Arquivos Excel", "*.xlsx")])
    if filename:
        entrada_arquivo.delete(0, tk.END)
        entrada_arquivo.insert(0, filename)
        btn_executar.config(state=tk.NORMAL)

def aplicar_bordas(planilha):
    for row in planilha.iter_rows():
        for cell in row:
            cell.border = openpyxl.styles.Border(
                left=openpyxl.styles.Side(border_style='thin', color='000000'),
                right=openpyxl.styles.Side(border_style='thin', color='000000'),
                top=openpyxl.styles.Side(border_style='thin', color='000000'),
                bottom=openpyxl.styles.Side(border_style='thin', color='000000')
            )

def substituir_valor_zero(valor):
    data_padrao = datetime(2000, 1, 1)  # Data padrão para substituir os valores zero
    if valor == 0:
        return data_padrao
    else:
        return valor

def limpar_posto(posto):
    # Esta função remove números e ':' da string, deixando apenas letras e caracteres especiais
    if isinstance(posto, str):
        return ''.join(filter(lambda x: not x.isdigit() and x != ':', posto))
    else:
        return posto

def renomear_colunas(df, renomear):
    if renomear == "Sim":
        # Exibir caixa de diálogo para o usuário inserir o novo nome e local do arquivo
        novo_nome_arquivo = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Arquivos Excel", "*.xlsx")])
        if novo_nome_arquivo:
            # Salvar o arquivo com o novo nome e local
            df.to_excel(novo_nome_arquivo, index=False)
            messagebox.showinfo("Sucesso", "Base extraida com sucesso!")
    else:
        # Salvar o arquivo no mesmo local e com o mesmo nome
        caminho_arquivo = entrada_arquivo.get()
        df.to_excel(caminho_arquivo, index=False)
        messagebox.showinfo("Sucesso", "Base extraida com sucesso!")

def executar_robo():
    caminho_arquivo = entrada_arquivo.get()
    
    if not caminho_arquivo:
        messagebox.showerror("Erro", "Por favor, selecione um arquivo Excel primeiro.")
        return
    
    btn_executar.config(state=tk.DISABLED)
    
    try:
        text_cmd.config(state=tk.NORMAL)
        text_cmd.delete(1.0, tk.END)
        text_cmd.insert(tk.END, "Iniciando processo...\n")
        text_cmd.config(state=tk.DISABLED)
        
        time.sleep(2)
        
        df = pd.read_excel(caminho_arquivo, header=41)
        
        # Substituir valores 0 na coluna 'Data Posto' pela data padrão
        df['Data Posto'] = df['Data Posto'].apply(substituir_valor_zero)
        
        text_cmd.config(state=tk.NORMAL)
        text_cmd.insert(tk.END, "Apagando linhas e colunas indesejadas...\n")
        text_cmd.config(state=tk.DISABLED)
        
        colunas_remover = ['Abertura', 'Posicionamento Interno', 'Reinc.', 'R. CO', 'Posicio.', 'Tec.']
        df = df.drop(columns=colunas_remover)
        
        text_cmd.config(state=tk.NORMAL)
        text_cmd.insert(tk.END, "Removendo Cliente: Correios\n")
        text_cmd.config(state=tk.DISABLED)
        
        df = df[~df['Cliente'].astype(str).str.contains('CORREIOS', case=False)]
        
        text_cmd.config(state=tk.NORMAL)
        text_cmd.insert(tk.END, "Ordenando por Data Posto...\n")
        text_cmd.config(state=tk.DISABLED)
        
        df.sort_values(by=['Data Posto'], inplace=True)
        
        text_cmd.config(state=tk.NORMAL)
        text_cmd.insert(tk.END, "Formatando Datas...\n")
        text_cmd.config(state=tk.DISABLED)
        
        df['Data Posto'] = df['Data Posto'].apply(lambda x: datetime.strftime(x, '%d/%m/%y %H:%M'))
        
        text_cmd.config(state=tk.NORMAL)
        text_cmd.insert(tk.END, "Limpando valores da coluna 'Posto'...\n")
        text_cmd.config(state=tk.DISABLED)
        
        df['Posto'] = df['Posto'].apply(limpar_posto)  # Aplica a função limpar_posto() na coluna 'Posto'
        
        text_cmd.config(state=tk.NORMAL)
        text_cmd.insert(tk.END, "Ajustando colunas...\n")
        text_cmd.config(state=tk.DISABLED)
        
        renomear_colunas(df, renomear_var.get())  # Chama a função de renomeação com base na seleção do usuário
        
        text_cmd.config(state=tk.NORMAL)
        text_cmd.insert(tk.END, "Base extraida com sucesso.\n")
        text_cmd.config(state=tk.DISABLED)
        
    except Exception as e:
        messagebox.showerror("Erro", f"Falha ao executar o robô: {str(e)}")
        text_cmd.config(state=tk.NORMAL)
        text_cmd.insert(tk.END, f"Erro: {str(e)}\n")
        text_cmd.config(state=tk.DISABLED)
    
    btn_executar.config(state=tk.NORMAL)

def fechar_aplicacao():
    root.destroy()

root = tk.Tk()
root.title("MACRO OI 360")
root.configure(bg='#000000')

style = ttk.Style()
style.configure('Flat.TCheckbutton', background='#000000', foreground='white', bordercolor='#000000')

frame = tk.Frame(root, bg='#000000')
frame.pack(padx=10, pady=10, fill='both', expand=True)

lbl_titulo1 = tk.Label(frame, text="MACRO OI 360", bg='#000000', fg='white', font=('Segoe UI', 18, 'bold'))
lbl_titulo1.grid(row=0, column=0, columnspan=3, padx=5, pady=10, sticky='nsew')

lbl_version = tk.Label(frame, text="Version 1.1", bg='#000000', fg='white', font=('Segoe UI', 8))
lbl_version.grid(row=999, column=0, padx=5, pady=5, sticky='sw')

lbl_titulo2 = tk.Label(frame, text="Selecione o Arquivo Excel", bg='#000000', fg='white', font=('Segoe UI', 14, 'bold'))
lbl_titulo2.grid(row=1, column=0, columnspan=3, padx=5, pady=10, sticky='w')

lbl_local_planilha = tk.Label(frame, text="Local da Planilha", bg='#000000', fg='white', font=('Segoe UI', 12))
lbl_local_planilha.grid(row=2, column=0, padx=5, pady=10, sticky='w')

entrada_arquivo = tk.Entry(frame, bg='#000000', fg='white', width=40, font=('Segoe UI', 12))
entrada_arquivo.grid(row=2, column=1, padx=5, pady=10, sticky='we')
entrada_arquivo.insert(0, "Diretório do arquivo")
entrada_arquivo.config(fg='grey')

btn_selecionar_arquivo = ttk.Button(frame, text="Selecionar", command=selecionar_arquivo)
btn_selecionar_arquivo.grid(row=2, column=2, padx=5, pady=10, sticky='e')

renomear_var = tk.StringVar(value="Não")
lbl_renomear = tk.Label(frame, text="Deseja renomear?", bg='#000000', fg='white', font=('Segoe UI', 12))
lbl_renomear.grid(row=3, column=0, columnspan=3, padx=5, pady=10, sticky='w')

radiobtn_sim = ttk.Radiobutton(frame, text="Sim", variable=renomear_var, value="Sim", style='Flat.TCheckbutton')
radiobtn_sim.grid(row=4, column=0, padx=5, pady=5, sticky='w')

radiobtn_nao = ttk.Radiobutton(frame, text="Não", variable=renomear_var, value="Não", style='Flat.TCheckbutton')
radiobtn_nao.grid(row=4, column=1, padx=5, pady=5, sticky='w')

text_cmd = tk.Text(frame, bg='#000000', fg='white', font=('Segoe UI', 10), wrap='word', height=8)
text_cmd.grid(row=5, column=0, columnspan=3, padx=5, pady=10, sticky='nsew')
text_cmd.tag_configure("success", foreground="green")
text_cmd.insert(tk.END, "CMD: Acompanhando o processo...\n")
text_cmd.config(state=tk.DISABLED)

btn_executar = ttk.Button(frame, text="Executar", command=executar_robo, state=tk.DISABLED)
btn_executar.grid(row=6, column=1, padx=5, pady=10, sticky='we')

btn_fechar = ttk.Button(frame, text="Fechar", command=fechar_aplicacao)
btn_fechar.grid(row=7, column=1, padx=5, pady=10, sticky='we')

root.mainloop()
