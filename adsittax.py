import tkinter as tk
from tkinter import messagebox
import openpyxl
from datetime import datetime

# Tabela de valores da Sittax
tabela_sittax = {
    5: 210.00,   # Valor ajustado para 5 CNPJs
    10: 245.00,  # Aproximadamente ajustado entre 5 e 15
    15: 280.00,  # Valor base para 15 CNPJs
    20: 314.00,  # Aproximadamente ajustado entre 15 e 30
    25: 348.00,  # Aproximadamente ajustado entre 20 e 30
    30: 389.00,  # Valor base para 30 CNPJs
    35: 421.00,  # Aproximadamente ajustado entre 30 e 50
    40: 455.00,  # Aproximadamente ajustado entre 35 e 50
    45: 470.00,  # Aproximadamente ajustado entre 40 e 50
    50: 497.00,  # Valor base para 50 CNPJs
    55: 523.00,  # Aproximadamente ajustado entre 50 e 70
    60: 549.00,  # Aproximadamente ajustado entre 50 e 70
    65: 582.00,  # Aproximadamente ajustado entre 60 e 70
    70: 635.00,  # Valor base para 70 CNPJs
    75: 665.00,  # Aproximadamente ajustado entre 70 e 100
    80: 695.00,  # Aproximadamente ajustado entre 70 e 100
    85: 725.00,  # Aproximadamente ajustado entre 80 e 100
    90: 755.00,  # Aproximadamente ajustado entre 80 e 100
    95: 785.00,  # Aproximadamente ajustado entre 90 e 100
    100: 795.00,  # Valor base para 100 CNPJs
    120: 845.00,  # Interpolado entre 100 e 150
    140: 899.00,  # Interpolado entre 100 e 150
    150: 999.00,  # Valor base para 150 CNPJs
    170: 1045.00, # Interpolado entre 150 e 200
    190: 1065.00, # Interpolado entre 150 e 200
    200: 1085.00, # Valor base para 200 CNPJs
    220: 1210.00, # Interpolado entre 200 e 300
    240: 1250.00, # Interpolado entre 200 e 300
    260: 1290.00, # Interpolado entre 200 e 300
    280: 1330.00, # Interpolado entre 200 e 300
    300: 1476.00, # Valor base para 300 CNPJs
    320: 1500.00, # Interpolado entre 300 e 500
    340: 1560.00, # Interpolado entre 300 e 500
    360: 1620.00, # Interpolado entre 300 e 500
    380: 1680.00, # Interpolado entre 300 e 500
    400: 1740.00, # Interpolado entre 300 e 500
    420: 1800.00, # Interpolado entre 300 e 500
    440: 1860.00, # Interpolado entre 300 e 500
    460: 1920.00, # Interpolado entre 300 e 500
    480: 1950.00, # Interpolado entre 300 e 500
    500: 1960.00, # Valor base para 500 CNPJs
    520: 2000.00, # Interpolado entre 500 e 1000
    540: 2040.00, # Interpolado entre 500 e 1000
    560: 2080.00, # Interpolado entre 500 e 1000
    580: 2120.00, # Interpolado entre 500 e 1000
    600: 2160.00, # Interpolado entre 500 e 1000
    620: 2200.00, # Interpolado entre 500 e 1000
    640: 2240.00, # Interpolado entre 500 e 1000
    660: 2280.00, # Interpolado entre 500 e 1000
    680: 2320.00, # Interpolado entre 500 e 1000
    700: 2360.00, # Interpolado entre 500 e 1000
    720: 2400.00, # Interpolado entre 500 e 1000
    740: 2440.00, # Interpolado entre 500 e 1000
    760: 2480.00, # Interpolado entre 500 e 1000
    780: 2520.00, # Interpolado entre 500 e 1000
    800: 2560.00, # Interpolado entre 500 e 1000
    820: 2600.00, # Interpolado entre 500 e 1000
    840: 2640.00, # Interpolado entre 500 e 1000
    860: 2680.00, # Interpolado entre 500 e 1000
    880: 2720.00, # Interpolado entre 500 e 1000
    900: 2760.00, # Interpolado entre 500 e 1000
    920: 2800.00, # Interpolado entre 500 e 1000
    940: 2840.00, # Interpolado entre 500 e 1000
    960: 2880.00, # Interpolado entre 500 e 1000
    980: 2920.00, # Interpolado entre 500 e 1000
    1000: 2500.00  # Valor base para 1000 CNPJs
}

# Função para salvar os dados no Excel
def salvar_excel(nome_cliente, mensalidade_atual, nova_mensalidade, desconto_aplicado, diferenca, cnpjs_atual, cnpjs_novo):
    try:
        # Abrir ou criar o arquivo Excel
        try:
            workbook = openpyxl.load_workbook("propostas_clientes_sittax.xlsx")
        except FileNotFoundError:
            workbook = openpyxl.Workbook()
            sheet = workbook.active
            # Criar cabeçalhos
            sheet.append(["Data", "Nome do Cliente", "Mensalidade Atual", "Nova Mensalidade", 
                          "Desconto Aplicado", "Diferença", "CNPJs Atuais", "CNPJs Novos"])
        else:
            sheet = workbook.active
        
        # Inserir uma nova linha com os dados
        sheet.append([datetime.now().strftime("%d/%m/%Y"),
                      nome_cliente,
                      f"R$ {mensalidade_atual:.2f}",
                      f"R$ {nova_mensalidade:.2f}",
                      "Sim" if desconto_aplicado else "Não",
                      f"R$ {diferenca:.2f}",
                      cnpjs_atual,
                      cnpjs_novo])
        
        # Salvar o arquivo Excel
        workbook.save("propostas_clientes_sittax.xlsx")
        messagebox.showinfo("Sucesso", "Dados salvos no Excel com sucesso!")
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao salvar no Excel: {e}")

# Função principal
def calcular_novo_valor():
    try:
        # Captura os dados da interface
        nome_cliente = entry_nome_cliente.get()
        mensalidade_atual = float(entry_mensalidade.get().replace(",", "."))
        cnpjs_atual = int(entry_cnpjs_atual.get())
        cnpjs_novo = int(entry_cnpjs_novo.get())
        desconto_aplicado = var_desconto.get() == 1

        # Determinar o novo valor pela tabela
        novo_valor = None
        for cnpj, valor in tabela_sittax.items():
            if cnpjs_novo <= cnpj:
                novo_valor = valor
                break
        
        if novo_valor is None:
            messagebox.showerror("Erro", "Quantidade de CNPJs não está na tabela da Sittax.")
            return

        # Calcular a diferença
        diferenca = novo_valor - mensalidade_atual

        # Exibir o resultado
        label_resultado.config(text=f"Novo valor da mensalidade: R$ {novo_valor:.2f}\nDiferença: R$ {diferenca:.2f}")

        # Salvar os dados no Excel
        salvar_excel(nome_cliente, mensalidade_atual, novo_valor, desconto_aplicado, diferenca, cnpjs_atual, cnpjs_novo)

    except ValueError:
        messagebox.showerror("Erro", "Por favor, preencha todos os campos corretamente.")

# Interface gráfica
root = tk.Tk()
root.title("Calculadora de Propostas Sittax")

# Campos de entrada
tk.Label(root, text="Nome do Cliente:").grid(row=0, column=0)
entry_nome_cliente = tk.Entry(root)
entry_nome_cliente.grid(row=0, column=1)

tk.Label(root, text="Mensalidade Atual (R$):").grid(row=1, column=0)
entry_mensalidade = tk.Entry(root)
entry_mensalidade.grid(row=1, column=1)

tk.Label(root, text="CNPJs Atuais:").grid(row=2, column=0)
entry_cnpjs_atual = tk.Entry(root)
entry_cnpjs_atual.grid(row=2, column=1)

tk.Label(root, text="CNPJs Novos:").grid(row=3, column=0)
entry_cnpjs_novo = tk.Entry(root)
entry_cnpjs_novo.grid(row=3, column=1)

# Opções de desconto
var_desconto = tk.IntVar()
tk.Label(root, text="Desconto aplicável:").grid(row=4, column=0)
tk.Radiobutton(root, text="Sim", variable=var_desconto, value=1).grid(row=4, column=1)
tk.Radiobutton(root, text="Não", variable=var_desconto, value=0).grid(row=4, column=2)

# Resultado
label_resultado = tk.Label(root, text="")
label_resultado.grid(row=5, column=0, columnspan=3)

# Botão para calcular
btn_calcular = tk.Button(root, text="Calcular", command=calcular_novo_valor)
btn_calcular.grid(row=6, column=0, columnspan=3)

root.mainloop()
