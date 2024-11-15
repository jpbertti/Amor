import tkinter as tk
from tkinter import messagebox
from datetime import datetime
import locale
import pandas as pd
import os
from tkcalendar import DateEntry  # Importa o DateEntry da tkcalendar
# pip install tkcalendar

# Define o local para usar o padrão brasileiro
locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')

# Função para calcular a série vermelha (hemoglobina, hemácias, hematócrito)
def calcular_serie_vermelha(genero, hemoglobina, hemacias, hematocrito):
    resultado = (hemoglobina * 3) / hematocrito
    
    if resultado == 1:
        status = "Normal e Hidratada"
    elif resultado < 1:
        status = "Desidratada"
    else:
        status = "Retenção hídrica"

    # Verifica se os valores estão dentro dos intervalos normais baseados no gênero
    if genero == "Homem":
        if not (4.5 <= hemacias <= 6.0):
            status += " (Eritrócitos fora do valor ideal)"
        if not (14.0 <= hemoglobina <= 16.5):
            status += " (Hemoglobina fora do valor ideal)"
        if not (42 <= hematocrito <= 52):
            status += " (Hematócrito fora do valor ideal)"
    elif genero == "Mulher":
        if not (4.0 <= hemacias <= 5.5):
            status += " (Eritrócitos fora do valor ideal)"
        if not (13.5 <= hemoglobina <= 15.5):
            status += " (Hemoglobina fora do valor ideal)"
        if not (39 <= hematocrito <= 47):
            status += " (Hematócrito fora do valor ideal)"
    
    return resultado, status

# Função para calcular a série branca (leucócitos, neutrófilos, linfócitos, monócitos, eosinófilos, basófilos)
def calcular_serie_branca(leucocitos, neutrofilos, linfocitos, monocitos, eosinofilos, basofilos):
    status = ""

    # Calcula a razão neutrófilos/linfócitos
    try:
        razao_neutrofilos_linfocitos = neutrofilos / linfocitos
    except ZeroDivisionError:
        razao_neutrofilos_linfocitos = float('inf')  # Evitar divisão por zero

    # Determina o status baseado na razão neutrófilos/linfócitos
    if razao_neutrofilos_linfocitos <= 1.5:
        status = "Ideal"
    elif 1.5 < razao_neutrofilos_linfocitos <= 2.5:
        status = "Alerta"
    else:
        status = "Inflamado"

    # Verifica se os valores estão dentro dos intervalos normais e ajusta o status conforme necessário
    if not (4000 <= leucocitos <= 6500):
        status += " (Leucócitos fora do valor ideal)"
    if not (45 <= neutrofilos <= 55):
        status += " (Neutrófilos fora do valor ideal)"
    if not (25 <= linfocitos <= 35):
        status += " (Linfócitos fora do valor ideal)"
    if not (3 <= monocitos <= 8):
        status += " (Monócitos fora do valor ideal)"
    if eosinofilos > 1:
        status += " (Eosinófilos fora do valor ideal)"
    if basofilos > 0.5:
        status += " (Basófilos fora do valor ideal)"
    
    return status, razao_neutrofilos_linfocitos

# Função para calcular os níveis de Vitamina D
def calcular_vitaminaD(nivel_vitaminaD):
    if nivel_vitaminaD < 20:
        status = "Deficiência"
    elif 20 <= nivel_vitaminaD <= 29:
        status = "Insuficiência"
    elif 30 <= nivel_vitaminaD <= 100:
        status = "Suficiência"
    else:
        status = "Alta"

    return status

# Função para salvar os resultados no Excel
def salvar_excel(data, nome, genero, hemoglobina, hemacias, hematocrito, leucocitos, neutrofilos, linfocitos, monocitos, eosinofilos, basofilos, resultado_serie_vermelha, status_vermelha, resultado_serie_branca, status_branca, nivel_vitaminaD, status_vitaminaD):
    # Cria a pasta com o nome do usuário, se não existir
    pasta = f"{nome.title()}"
    if not os.path.exists(pasta):
        os.makedirs(pasta)

    # Formata o nome do arquivo com o nome informado pelo usuário
    nome_arquivo = f"{data} - {nome} - Analises Laboratoriais.xlsx"
    caminho_arquivo = os.path.join(pasta, nome_arquivo)

    # Cria um DataFrame com os dados da série vermelha
    df_vermelha = pd.DataFrame({
        'Data': [data],
        'Nome': [nome],
        'Gênero': [genero],
        'Hemoglobina (g/dL)': [hemoglobina],
        'Hemácias (milhões/µL)': [hemacias],
        'Hematócrito (%)': [hematocrito],
        'Resultado Série Vermelha': [resultado_serie_vermelha],
        'Status Série Vermelha': [status_vermelha]
    })

    # Define os valores de referência com base no gênero
    if genero == "Mulher":
        valores_ideais_vermelha = pd.DataFrame({
            'Nome': 'Valor de Referência',
            'Gênero': ['Mulher'],
            'Hemoglobina (g/dL)': ['13,5-15,5'],
            'Hemácias (milhões/µL)': ['4,0-5,5'],
            'Hematócrito (%)': ['39-47']
        })
    else:  # Homem
        valores_ideais_vermelha = pd.DataFrame({
            'Nome': 'Valor de Referência',
            'Gênero': ['Homem'],
            'Hemoglobina (g/dL)': ['14,0-16,5'],
            'Hemácias (milhões/µL)': ['4,5-6,0'],
            'Hematócrito (%)': ['42-52']
        })

    df_vermelha = pd.concat([df_vermelha, valores_ideais_vermelha], ignore_index=True)

    # Cria um DataFrame com os dados da série branca
    df_branca = pd.DataFrame({
        'Data': [data],
        'Nome': [nome],
        'Gênero': [genero],
        'Leucócitos': [leucocitos],
        'Neutrófilos (%)': [neutrofilos],
        'Linfócitos (%)': [linfocitos],
        'Monócitos (%)': [monocitos],
        'Eosinófilos (%)': [eosinofilos],
        'Basófilos (%)': [basofilos],
        'Resultado Série Branca': [resultado_serie_branca],
        'Status Série Branca': [status_branca]
    })

    # Adiciona os valores de referência
    valores_ideais_branca = pd.DataFrame({
        'Nome': 'Valor de Referência',
        'Gênero': 'Mulher | Homem',
        'Leucócitos': '4000-6500',
        'Neutrófilos (%)': '45-55%',
        'Linfócitos (%)': '25-35%',
        'Monócitos (%)': '3-8%',
        'Eosinófilos (%)': '<1%',
        'Basófilos (%)': '<0,5%'
    }, index=[0])

    df_branca = pd.concat([df_branca, valores_ideais_branca], ignore_index=True)

    # Cria um DataFrame com os dados da Vitamina D
    df_vitaminaD = pd.DataFrame({
        'Data': [data],
        'Nome': [nome],
        'Nível de Vitamina D (ng/mL)': [nivel_vitaminaD],
        'Status Vitamina D': [status_vitaminaD]
    })

    # Adiciona os valores de referência para Vitamina D
    valores_ideais_vitaminaD = pd.DataFrame({
        'Nome': 'Valor de Referência',
        'Nível de Vitamina D (ng/mL)': ['<20', '20-29', '30-100', '>100'],
        'Status Vitamina D': ['Deficiência', 'Insuficiência', 'Suficiência', 'Alta']
    })

    df_vitaminaD = pd.concat([df_vitaminaD, valores_ideais_vitaminaD], ignore_index=True)

    # Salva os DataFrames em diferentes sheets
    with pd.ExcelWriter(caminho_arquivo) as writer:
        df_vermelha.to_excel(writer, sheet_name='Série Vermelha', index=False)
        df_branca.to_excel(writer, sheet_name='Série Branca', index=False)
        df_vitaminaD.to_excel(writer, sheet_name='Vitamina D', index=False)


# Função chamada ao confirmar o formulário
def ao_confirmar():
    nome = entry_nome.get().strip()
    genero = genero_var.get()
    hemoglobina = entry_hemoglobina.get().strip()
    hemacias = entry_hemacias.get().strip()
    hematocrito = entry_hematocrito.get().strip()
    leucocitos = entry_leucocitos.get().strip()
    neutrofilos = entry_neutrofilos.get().strip()
    linfocitos = entry_linfocitos.get().strip()
    monocitos = entry_monocitos.get().strip()
    eosinofilos = entry_eosinofilos.get().strip()
    basofilos = entry_basofilos.get().strip()
    vitaminaD = entry_vitaminaD.get().strip()
    data = entry_data.get_date().strftime('%d-%m-%Y')

    # Valida se todos os campos estão preenchidos
    if not (nome and genero and hemoglobina and hemacias and hematocrito and leucocitos and neutrofilos and linfocitos and monocitos and eosinofilos and basofilos and vitaminaD):
        messagebox.showwarning("Aviso", "Todos os campos devem ser preenchidos.")
        return

    try:
        hemoglobina = float(hemoglobina.replace(',', '.'))
        hemacias = float(hemacias.replace(',', '.'))
        hematocrito = float(hematocrito.replace(',', '.'))
        leucocitos = float(leucocitos.replace(',', '.'))
        neutrofilos = float(neutrofilos.replace(',', '.'))
        linfocitos = float(linfocitos.replace(',', '.'))
        monocitos = float(monocitos.replace(',', '.'))
        eosinofilos = float(eosinofilos.replace(',', '.'))
        basofilos = float(basofilos.replace(',', '.'))
        vitaminaD = float(vitaminaD.replace(',', '.'))
    except ValueError:
        messagebox.showwarning("Aviso", "Por favor, insira valores numéricos válidos.")
        return

    # Calcula os resultados
    resultado_serie_vermelha, status_vermelha = calcular_serie_vermelha(genero, hemoglobina, hemacias, hematocrito)
    status_branca, resultado_serie_branca = calcular_serie_branca(leucocitos, neutrofilos, linfocitos, monocitos, eosinofilos, basofilos)
    status_vitaminaD = calcular_vitaminaD(vitaminaD)

    # Exibe os resultados em uma mensagem
    messagebox.showinfo("Resultado", f"Resultado Série Vermelha: {resultado_serie_vermelha:.2f}\nStatus Série Vermelha: {status_vermelha}\n\nResultado Série Branca: {resultado_serie_branca:.2f}\nStatus Série Branca: {status_branca}\n\nNível de Vitamina D: {vitaminaD:.2f} ng/mL\nStatus Vitamina D: {status_vitaminaD}")

    # Salva os resultados no Excel
    salvar_excel(data, nome, genero, hemoglobina, hemacias, hematocrito, leucocitos, neutrofilos, linfocitos, monocitos, eosinofilos, basofilos, resultado_serie_vermelha, status_vermelha, resultado_serie_branca, status_branca, vitaminaD, status_vitaminaD)

# Criação da interface gráfica
janela = tk.Tk()
janela.title("Analises Laboratoriais")

# Define o tamanho da janela (largura x altura)
janela.geometry("350x450")

# Nome
tk.Label(janela, text="Nome:").grid(row=0, column=0)
entry_nome = tk.Entry(janela, width=30)
entry_nome.grid(row=0, column=1, padx=10, pady=5)

# Data
tk.Label(janela, text="Data:").grid(row=1, column=0)
entry_data = DateEntry(janela, date_pattern='dd/mm/yyyy')
entry_data.grid(row=1, column=1, padx=10, pady=5)

# Gênero
tk.Label(janela, text="Gênero:").grid(row=2, column=0)
genero_var = tk.StringVar()
genero_var.set("Mulher")
genero_opcoes = tk.OptionMenu(janela, genero_var, "Mulher", "Homem")
genero_opcoes.grid(row=2, column=1, padx=10, pady=5)

# Hemoglobina
tk.Label(janela, text="Hemoglobina (g/dL):").grid(row=3, column=0)
entry_hemoglobina = tk.Entry(janela)
entry_hemoglobina.grid(row=3, column=1, padx=10, pady=5)

# Hemácias
tk.Label(janela, text="Hemácias (milhões/µL):").grid(row=4, column=0)
entry_hemacias = tk.Entry(janela)
entry_hemacias.grid(row=4, column=1, padx=10, pady=5)

# Hematócrito
tk.Label(janela, text="Hematócrito (%):").grid(row=5, column=0)
entry_hematocrito = tk.Entry(janela)
entry_hematocrito.grid(row=5, column=1, padx=10, pady=5)

# Leucócitos
tk.Label(janela, text="Leucócitos:").grid(row=6, column=0)
entry_leucocitos = tk.Entry(janela)
entry_leucocitos.grid(row=6, column=1, padx=10, pady=5)

# Neutrófilos
tk.Label(janela, text="Neutrófilos (%):").grid(row=7, column=0)
entry_neutrofilos = tk.Entry(janela)
entry_neutrofilos.grid(row=7, column=1, padx=10, pady=5)

# Linfócitos
tk.Label(janela, text="Linfócitos (%):").grid(row=8, column=0)
entry_linfocitos = tk.Entry(janela)
entry_linfocitos.grid(row=8, column=1, padx=10, pady=5)

# Monócitos
tk.Label(janela, text="Monócitos (%):").grid(row=9, column=0)
entry_monocitos = tk.Entry(janela)
entry_monocitos.grid(row=9, column=1, padx=10, pady=5)

# Eosinófilos
tk.Label(janela, text="Eosinófilos (%):").grid(row=10, column=0)
entry_eosinofilos = tk.Entry(janela)
entry_eosinofilos.grid(row=10, column=1, padx=10, pady=5)

# Basófilos
tk.Label(janela, text="Basófilos (%):").grid(row=11, column=0)
entry_basofilos = tk.Entry(janela)
entry_basofilos.grid(row=11, column=1, padx=10, pady=5)

# Vitamina D
tk.Label(janela, text="Vitamina D (ng/mL):").grid(row=12, column=0)
entry_vitaminaD = tk.Entry(janela)
entry_vitaminaD.grid(row=12, column=1, padx=10, pady=5)

# Botão Confirmar
botao_confirmar = tk.Button(janela, text="Fazer Analise Laboratoriais", command=ao_confirmar)
botao_confirmar.grid(row=13, column=0, columnspan=2, padx=10, pady=20)

janela.mainloop()
