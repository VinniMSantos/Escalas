import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkcalendar import Calendar
import pandas as pd
from datetime import timedelta

# Variáveis globais para armazenar os dados da planilha e do período
selected_file = None
selected_sheet = None
start_date = None
end_date = None

# Função para abrir a tela principal após a seleção de período
def abrir_tela_principal():
    janela_principal = tk.Toplevel()
    janela_principal.title("Inserção de Escala")
    janela_principal.geometry("1000x600")

    # Exibe o período selecionado
    label_periodo = tk.Label(janela_principal, text=f"Período Selecionado: {start_date} até {end_date}", font=("Arial", 12))
    label_periodo.pack(pady=10)

    # Frame para a pré-visualização da planilha
    frame_tree = tk.Frame(janela_principal)
    frame_tree.pack(fill=tk.BOTH, padx=10, pady=10)

    # Criação da pré-visualização da planilha com inputs na primeira linha
    tree = ttk.Treeview(frame_tree, columns=("Localização", "Técnico", "Hora/Data Início", "Hora/Data Fim", "Turno", "Objetivo"), show='headings', height=10)
    tree.column("Localização", anchor=tk.CENTER, width=120)
    tree.column("Técnico", anchor=tk.CENTER, width=120)
    tree.column("Hora/Data Início", anchor=tk.CENTER, width=150)
    tree.column("Hora/Data Fim", anchor=tk.CENTER, width=150)
    tree.column("Turno", anchor=tk.CENTER, width=100)
    tree.column("Objetivo", anchor=tk.CENTER, width=120)

    # Definir os títulos das colunas
    tree.heading("Localização", text="Localização")
    tree.heading("Técnico", text="Técnico")
    tree.heading("Hora/Data Início", text="Hora/Data Início")
    tree.heading("Hora/Data Fim", text="Hora/Data Fim")
    tree.heading("Turno", text="Turno")
    tree.heading("Objetivo", text="Objetivo")
    tree.pack(fill=tk.BOTH, expand=True)

    # Inputs para edição na primeira linha da tabela
    inputs_frame = tk.Frame(janela_principal)
    inputs_frame.pack(fill=tk.X, padx=10, pady=5)

    entry_localizacao = ttk.Combobox(inputs_frame, values=["Escritório", "Sobreaviso", "Unidade", "Folga", "Férias"], width=15)
    entry_localizacao.grid(row=0, column=0, padx=5)
    
    entry_tecnico = ttk.Combobox(inputs_frame, values=["Técnico 1", "Técnico 2", "Técnico 3", "Técnico 4"], width=15)
    entry_tecnico.grid(row=0, column=1, padx=5)

    entry_hora_inicio = tk.Entry(inputs_frame, width=20)
    entry_hora_inicio.grid(row=0, column=2, padx=5)

    entry_hora_fim = tk.Entry(inputs_frame, width=20)
    entry_hora_fim.grid(row=0, column=3, padx=5)

    entry_turno = tk.Entry(inputs_frame, width=10)
    entry_turno.grid(row=0, column=4, padx=5)

    entry_objetivo = ttk.Combobox(inputs_frame, values=["Visita de rotina", "Treinamento", "Outros"], width=15)
    entry_objetivo.grid(row=0, column=5, padx=5)

    # Função para calcular turno automaticamente
    def calcular_turno():
        hora_inicio = int(entry_hora_inicio.get().split(":")[0])
        if 6 <= hora_inicio <= 16:
            entry_turno.delete(0, tk.END)
            entry_turno.insert(0, "Diurno")
        else:
            entry_turno.delete(0, tk.END)
            entry_turno.insert(0, "Noturno")

    entry_hora_inicio.bind("<FocusOut>", lambda event: calcular_turno())

    # Função para adicionar os dados à próxima linha da planilha
    def adicionar_linha():
        nova_entrada = [
            entry_localizacao.get(),
            entry_tecnico.get(),
            entry_hora_inicio.get(),
            entry_hora_fim.get(),
            entry_turno.get(),
            entry_objetivo.get()
        ]
        tree.insert('', 'end', values=nova_entrada)

        # Limpa os campos de input para a próxima inserção
        entry_localizacao.set('')
        entry_tecnico.set('')
        entry_hora_inicio.delete(0, tk.END)
        entry_hora_fim.delete(0, tk.END)
        entry_turno.delete(0, tk.END)
        entry_objetivo.set('')

    # Botão de confirmação (visto) no final da linha de inserção
    btn_visto = tk.Button(inputs_frame, text="✔️", command=adicionar_linha, font=("Arial", 12), bg="green", fg="white", width=3)
    btn_visto.grid(row=0, column=6, padx=10)

# Função para confirmar o período selecionado
def confirmar_periodo():
    global start_date, end_date
    if janela_periodo.start_date and janela_periodo.end_date:
        start_date = janela_periodo.start_date.strftime('%d/%m/%Y')
        end_date = janela_periodo.end_date.strftime('%d/%m/%Y')
        messagebox.showinfo("Período Selecionado", f"Período de {start_date} até {end_date} selecionado.")
        janela_periodo.destroy()
        abrir_tela_principal()  # Abre a tela principal após a seleção do período
    else:
        messagebox.showerror("Erro", "Selecione um período válido.")

# Função para selecionar arquivo e sheet
def selecionar_arquivo_e_sheet():
    global selected_file, selected_sheet
    arquivo = filedialog.askopenfilename(title="Selecione a planilha", filetypes=[("Arquivos Excel", "*.xlsx")])
    if arquivo:
        selected_file = arquivo
        label_arquivo.config(text=arquivo)
        try:
            planilha = pd.ExcelFile(arquivo)
            abas = planilha.sheet_names
            if abas:
                combobox_abas['values'] = abas
                combobox_abas.current(0)
                combobox_abas.config(state="readonly")
                selected_sheet = abas[0]
            else:
                messagebox.showerror("Erro", "Nenhuma aba encontrada na planilha.")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao carregar a planilha: {str(e)}")

def on_select_sheet(event):
    global selected_sheet
    selected_sheet = combobox_abas.get()

def confirmar_selecao():
    if not selected_file:
        messagebox.showerror("Erro", "Nenhum arquivo selecionado.")
        return
    if not selected_sheet:
        messagebox.showerror("Erro", "Nenhuma aba selecionada.")
        return
    messagebox.showinfo("Sucesso", f"Arquivo e aba selecionados com sucesso!")
    janela.withdraw()  # Esconde a janela principal
    janela_selecao_periodo()  # Abre a janela para selecionar o período

def janela_selecao_periodo():
    global janela_periodo
    janela_periodo = tk.Toplevel()
    janela_periodo.title("Seleção de Período")
    janela_periodo.geometry("500x500")
    
    label_instrucao = tk.Label(janela_periodo, text="Selecione o período clicando na data de início e na data de fim", font=("Arial", 12))
    label_instrucao.pack(pady=10)
    
    calendario = Calendar(janela_periodo, selectmode='day', date_pattern='dd/mm/yyyy', showweeknumbers=False)
    calendario.pack(pady=10)
    
    janela_periodo.start_date = None
    janela_periodo.end_date = None

    def on_date_click(event):
        selected = calendario.selection_get()
        if not janela_periodo.start_date:
            janela_periodo.start_date = selected
            janela_periodo.end_date = None
            calendario.calevent_remove('selected_range')
            calendario.calevent_create(selected, 'Início', 'selected_range')
            label_selecao.config(text=f"Data Inicial: {selected.strftime('%d/%m/%Y')}")
        elif not janela_periodo.end_date:
            if selected < janela_periodo.start_date:
                messagebox.showerror("Erro", "A data final não pode ser anterior à data inicial.")
                return
            if (selected - janela_periodo.start_date).days > 31:
                messagebox.showerror("Erro", "O período selecionado não pode exceder um mês.")
                return
            janela_periodo.end_date = selected
            calendario.calevent_remove('selected_range')
            current_date = janela_periodo.start_date
            while current_date <= janela_periodo.end_date:
                calendario.calevent_create(current_date, 'Período', 'selected_range')
                current_date += timedelta(days=1)
            label_selecao.config(text=f"Período: {janela_periodo.start_date.strftime('%d/%m/%Y')} até {janela_periodo.end_date.strftime('%d/%m/%Y')}")
        else:
            calendario.calevent_remove('selected_range')
            janela_periodo.start_date = selected
            janela_periodo.end_date = None
            calendario.calevent_create(selected, 'Início', 'selected_range')
            label_selecao.config(text=f"Data Inicial: {selected.strftime('%d/%m/%Y')}")

    calendario.bind("<<CalendarSelected>>", on_date_click)
    
    label_selecao = tk.Label(janela_periodo, text="Nenhum período selecionado", font=("Arial", 12))
    label_selecao.pack(pady=10)
    
    btn_confirmar = tk.Button(janela_periodo, text="Selecionar Período", command=confirmar_periodo, font=("Arial", 12), bg="green", fg="white")
    btn_confirmar.pack(pady=20)

# Interface gráfica inicial
janela = tk.Tk()
janela.title("Sistema de Escalas")
janela.geometry("500x300")

titulo = tk.Label(janela, text="Sistema de Escalas", font=("Arial", 16, "bold"))
titulo.pack(pady=10)

btn_selecionar_arquivo = tk.Button(janela, text="Selecionar Planilha Excel", command=selecionar_arquivo_e_sheet, font=("Arial", 12), bg="blue", fg="white")
btn_selecionar_arquivo.pack(pady=20)

label_arquivo = tk.Label(janela, text="Nenhum arquivo selecionado", font=("Arial", 10))
label_arquivo.pack(pady=10)

label_abas = tk.Label(janela, text="Selecione a aba:", font=("Arial", 12))
label_abas.pack(pady=5)

combobox_abas = ttk.Combobox(janela, values=[], state="disabled", font=("Arial", 12))
combobox_abas.pack(pady=5)
combobox_abas.bind("<<ComboboxSelected>>", on_select_sheet)

btn_confirmar = tk.Button(janela, text="Confirmar", command=confirmar_selecao, font=("Arial", 12), bg="green", fg="white")
btn_confirmar.pack(pady=20)

janela.mainloop()
