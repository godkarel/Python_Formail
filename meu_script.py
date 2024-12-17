import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
import re
import csv
import smtplib
import geocoder


from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

def montar_mensagem():
    # Coleta os dados dos campos de entrada fixos
    campos = {
        "Nome": entry_nome.get(),
        "Circuito / Cliente": entry_circuito_cliente.get(),
        "Endereço": entry_endereco.get(),
        "Projeto": entry_projeto.get(),
        "OS Nº": entry_os_numero.get(),
        "Distância Lançamento Aéreo Projetado (m)": entry_dist_aereo_proj.get(),
        "Distância Lançamento Aéreo Realizado (m)": entry_dist_aereo_real.get(),
        "Distância Lançamento Cordoalha (m)": entry_dist_cordoalha.get(),
        "Postes Equipados (unid.)": entry_postes_equipados.get(),
        "Qtde. CAIXA CEO Instalada Projetada (unid.)": entry_qtde_caixa_proj.get(),
        "Qtde. CAIXA CEO Instalada Realizada (unid.)": entry_qtde_caixa_real.get(),
        "Capacidade Cabo Lançado": entry_capacidade_cabo.get(),
        "Designação do Cabo / nº Cabo": entry_designacao_cabo.get(),
        "HUB / SITE": entry_hub_site.get(),
        "Sangria Realizada (unid.)": entry_sangria_real.get(),
        "Fusão Cabo Realizado (unid.)": entry_fusao_cabo_real.get(),
        "Pontas Reservadas Projetadas (unid.)": entry_pontas_proj.get(),
        "Pontas Reservadas Realizadas (unid.)": entry_pontas_real.get(),
        "Lançamento Tubete Projetado (m)": entry_lancamento_tubete_proj.get(),
        "Lançamento Tubete Realizado (m)": entry_lancamento_tubete_real.get(),
        "Lançamento Duto Realizado (m)": entry_lancamento_duto_real.get(),
        "Pendências": entry_pendencias.get(),
    }

    # Monta a mensagem apenas com campos fixos não vazios
    mensagem = "Relatório de Gastos OS\n\n"
    for titulo, valor in campos.items():
        if valor:  # Verifica se o campo não está vazio
            mensagem += f"{titulo}: {valor}\n"

    # Adiciona os valores dinâmicos das entradas
    mensagem += "\nEntradas Dinâmicas:\n"
    for linha, (entry_capacidade, entry_cabo, entry_fibra) in entradas.items():
        capacidade = entry_capacidade.get().strip()
        cabo = entry_cabo.get().strip()
        fibra = entry_fibra.get().strip()

        # Verifica se algum dos campos dessa linha foi preenchido
        if capacidade or cabo or fibra:
            mensagem += f"Linha {linha}:\n"
            if capacidade:
                mensagem += f"  - Capacidade: {capacidade}\n"
            if cabo:
                mensagem += f"  - Cabo: {cabo}\n"
            if fibra:
                mensagem += f"  - Fibra: {fibra}\n"

    return mensagem

def limpar_formulario():
    entry_nome.delete(0, 'end'),
    entry_circuito_cliente.delete(0, 'end')
    entry_endereco.delete(0, 'end')
    entry_projeto.delete(0, 'end')
    entry_os_numero.delete(0, 'end')
    entry_dist_aereo_proj.delete(0, 'end')
    entry_dist_aereo_real.delete(0, 'end')
    entry_dist_cordoalha.delete(0, 'end')
    entry_postes_equipados.delete(0, 'end')
    entry_qtde_caixa_proj.delete(0, 'end')
    entry_qtde_caixa_real.delete(0, 'end')
    entry_capacidade_cabo.delete(0, 'end')
    entry_designacao_cabo.delete(0, 'end')
    entry_hub_site.delete(0, 'end')
    entry_sangria_real.delete(0, 'end')
    entry_fusao_cabo_real.delete(0, 'end')
    entry_pontas_proj.delete(0, 'end')
    entry_pontas_real.delete(0, 'end')
    entry_lancamento_tubete_proj.delete(0, 'end')
    entry_lancamento_tubete_real.delete(0, 'end')
    entry_lancamento_duto_real.delete(0, 'end')
    entry_pendencias.delete(0, 'end')

def enviar_email():
    # Obtém os valores dos campos de entrada
    assunto = "Relatório de Gastos OS"
    mensagem = montar_mensagem()  # Chama a função para montar a mensagem
    destinatario = "karel.jogos@gmail.com"
    
    remetente = 'testparapython@gmail.com'
    senha = 'klei ijdi xzma shoh'

    # Configura o servidor SMTP
    servidor = smtplib.SMTP('smtp.gmail.com', 587)
    servidor.starttls()
    servidor.login(remetente, senha)

    # Cria o e-mail
    email = MIMEMultipart()
    email['From'] = remetente
    email['To'] = destinatario
    email['Subject'] = assunto

    # Adiciona o corpo do e-mail

    print(mensagem)
    email.attach(MIMEText(mensagem, 'plain'))

    # Envia o e-mail
    servidor.sendmail(remetente, destinatario, email.as_string())
    servidor.quit()

    print('E-mail enviado com sucesso!')

# Criar a janela principal
window = tk.Tk()
window.title("Formulário Completo")

# Criar o widget Notebook (abas)
notebook = ttk.Notebook(window)
notebook.pack(expand=True, fill='both')

# Criar o frame da primeira aba (Formulário)
tab1 = ttk.Frame(notebook)
notebook.add(tab1, text="Formulário")

# Criar o frame da segunda aba (Outro Conteúdo)
tab2 = ttk.Frame(notebook)
notebook.add(tab2, text="Outra Aba")

# Definir tamanho da fonte
large_font = ('Arial', 12, 'bold')

# Nome
label_nome = tk.Label(tab1, text="Nome", font=large_font)
label_nome.grid(row=0, column=0, columnspan=4, padx=10, pady=5, sticky='ew')
entry_nome = tk.Entry(tab1, font=large_font, justify='center')
entry_nome.grid(row=1, column=0, columnspan=4, padx=10, pady=5, sticky='ew')

# Bloco 1
label_circuito_cliente = tk.Label(tab1, text="Circuito / Cliente", font=large_font)
label_circuito_cliente.grid(row=6, column=0, padx=10, pady=5, sticky='ew')
entry_circuito_cliente = tk.Entry(tab1, font=large_font, justify='center')
entry_circuito_cliente.grid(row=7, column=0, padx=10, pady=5, sticky='ew')

label_endereco = tk.Label(tab1, text="Endereço", font=large_font)
label_endereco.grid(row=6, column=1, padx=10, pady=5, sticky='ew')
entry_endereco = tk.Entry(tab1, font=large_font, justify='center')
entry_endereco.grid(row=7, column=1, padx=10, pady=5, sticky='ew')

label_projeto = tk.Label(tab1, text="Projeto", font=large_font)
label_projeto.grid(row=6, column=2, padx=10, pady=5, sticky='ew')
entry_projeto = tk.Entry(tab1, font=large_font, justify='center')
entry_projeto.grid(row=7, column=2, padx=10, pady=5, sticky='ew')

label_os_numero = tk.Label(tab1, text="OS Nº", font=large_font)
label_os_numero.grid(row=6, column=3, padx=10, pady=5, sticky='ew')
entry_os_numero = tk.Entry(tab1, font=large_font, justify='center')
entry_os_numero.grid(row=7, column=3, padx=10, pady=5, sticky='ew')

# Bloco 2
label_dist_aereo_proj = tk.Label(tab1, text="Distância Lançamento Aéreo Projetado (m)", font=large_font)
label_dist_aereo_proj.grid(row=9, column=0, padx=10, pady=5, sticky='ew')
entry_dist_aereo_proj = tk.Entry(tab1, font=large_font, justify='center')
entry_dist_aereo_proj.grid(row=10, column=0, padx=10, pady=5, sticky='ew')

label_dist_aereo_real = tk.Label(tab1, text="Distância Lançamento Aéreo Realizado (m)", font=large_font)
label_dist_aereo_real.grid(row=9, column=1, padx=10, pady=5, sticky='ew')
entry_dist_aereo_real = tk.Entry(tab1, font=large_font, justify='center')
entry_dist_aereo_real.grid(row=10, column=1, padx=10, pady=5, sticky='ew')

label_dist_cordoalha = tk.Label(tab1, text="Distância Lançamento Cordoalha (m)", font=large_font)
label_dist_cordoalha.grid(row=9, column=2, padx=10, pady=5, sticky='ew')
entry_dist_cordoalha = tk.Entry(tab1, font=large_font, justify='center')
entry_dist_cordoalha.grid(row=10, column=2, padx=10, pady=5, sticky='ew')

label_postes_equipados = tk.Label(tab1, text="Postes Equipados (unid.)", font=large_font)
label_postes_equipados.grid(row=9, column=3, padx=10, pady=5, sticky='ew')
entry_postes_equipados = tk.Entry(tab1, font=large_font, justify='center')
entry_postes_equipados.grid(row=10, column=3, padx=10, pady=5, sticky='ew')

# Bloco 3
label_qtde_caixa_proj = tk.Label(tab1, text="Qtde. CAIXA \ CEO Instalada Projetada (unid.)")
label_qtde_caixa_proj.grid(row=11, column=0, padx=10, pady=5, sticky='ew')
entry_qtde_caixa_proj = tk.Entry(tab1, justify='center')
entry_qtde_caixa_proj.grid(row=12, column=0, padx=10, pady=5, sticky='ew')

label_qtde_caixa_real = tk.Label(tab1, text="Qtde. CAIXA \ CEO Instalada Realizada (unid.)")
label_qtde_caixa_real.grid(row=11, column=1, padx=10, pady=5, sticky='ew')
entry_qtde_caixa_real = tk.Entry(tab1, justify='center')
entry_qtde_caixa_real.grid(row=12, column=1, padx=10, pady=5, sticky='ew')

label_capacidade_cabo = tk.Label(tab1, text="Capacidade Cabo Lançado")
label_capacidade_cabo.grid(row=11, column=2, padx=10, pady=5, sticky='ew')
entry_capacidade_cabo = tk.Entry(tab1, justify='center')
entry_capacidade_cabo.grid(row=12, column=2, padx=10, pady=5, sticky='ew')

label_designacao_cabo = tk.Label(tab1, text="Designação do Cabo / nº Cabo")
label_designacao_cabo.grid(row=11, column=3, padx=10, pady=5, sticky='ew')
entry_designacao_cabo = tk.Entry(tab1, justify='center')
entry_designacao_cabo.grid(row=12, column=3, padx=10, pady=5, sticky='ew')

# Bloco 4

label_hub_site = tk.Label(tab1, text="HUB / SITE")
label_hub_site.grid(row=13, column=0, padx=10, pady=5, sticky='ew')
entry_hub_site = tk.Entry(tab1, justify='center')
entry_hub_site.grid(row=14, column=0, padx=10, pady=5, sticky='ew')

label_sangria_real = tk.Label(tab1, text="Sangria Realizada (unid.)")
label_sangria_real.grid(row=13, column=1, padx=10, pady=5, sticky='ew')
entry_sangria_real = tk.Entry(tab1, justify='center')
entry_sangria_real.grid(row=14, column=1, padx=10, pady=5, sticky='ew')

label_fusao_cabo_real = tk.Label(tab1, text="Fusão Cabo Realizado (unid.)")
label_fusao_cabo_real.grid(row=13, column=2, padx=10, pady=5, sticky='ew')
entry_fusao_cabo_real = tk.Entry(tab1, justify='center')
entry_fusao_cabo_real.grid(row=14, column=2, padx=10, pady=5, sticky='ew')

label_pontas_proj = tk.Label(tab1, text="Pontas Reservadas Projetadas (unid.)")
label_pontas_proj.grid(row=13, column=3, padx=10, pady=5, sticky='ew')
entry_pontas_proj = tk.Entry(tab1, justify='center')
entry_pontas_proj.grid(row=14, column=3, padx=10, pady=5, sticky='ew')


# Bloco 5
label_lancamento_duto_real = tk.Label(tab1, text="Lançamento Duto Realizado (m)")
label_lancamento_duto_real.grid(row=15, column=0, padx=10, pady=5, sticky='ew')
entry_lancamento_duto_real = tk.Entry(tab1, justify='center')
entry_lancamento_duto_real.grid(row=16, column=0, padx=10, pady=5, sticky='ew')

label_pontas_real = tk.Label(tab1, text="Pontas Reservadas Realizadas (unid.)")
label_pontas_real.grid(row=15, column=1, padx=10, pady=5, sticky='ew')
entry_pontas_real = tk.Entry(tab1, justify='center')
entry_pontas_real.grid(row=16, column=1, padx=10, pady=5, sticky='ew')

label_lancamento_tubete_proj = tk.Label(tab1, text="Lançamento Tubete Projetado (m)")
label_lancamento_tubete_proj.grid(row=15, column=2, padx=10, pady=5, sticky='ew')
entry_lancamento_tubete_proj = tk.Entry(tab1, justify='center')
entry_lancamento_tubete_proj.grid(row=16, column=2, padx=10, pady=5, sticky='ew')

label_lancamento_tubete_real = tk.Label(tab1, text="Lançamento Tubete Realizado (m)")
label_lancamento_tubete_real.grid(row=15, column=3, padx=10, pady=5, sticky='ew')
entry_lancamento_tubete_real = tk.Entry(tab1, justify='center')
entry_lancamento_tubete_real.grid(row=16, column=3, padx=10, pady=5, sticky='ew')

#Bloco 6
label_pendencias = tk.Label(tab1, text="Pendências")
label_pendencias.grid(row=17, column=0, padx=10, pady=5, sticky='ew')
entry_pendencias = tk.Entry(tab1, justify='center')
entry_pendencias.grid(row=18, column=0, padx=10, pady=5, sticky='ew')

# Mensagem
label_mensagem = tk.Label(tab1, text="Mensagem", font=large_font)
label_mensagem.grid(row=20, column=0, columnspan=6, padx=10, pady=(15, 5), sticky='ew')  # Espaçamento acima

entry_mensagem = tk.Text(tab1, height=5, font=large_font)
entry_mensagem.grid(row=21, column=0, columnspan=6, padx=10, pady=5, sticky='ew')  # Centralizando com columnspan

# Espaçamento entre o cabeçalho e as entradas
label_empty_space = tk.Label(tab1, text="", font=large_font)  # Espaço vazio
label_empty_space.grid(row=30, column=0, padx=10, pady=(5, 15), sticky='ew')  # Adiciona um espaço em branco

# Botões de ação (Opcional: Salvar, Enviar, etc.)
button_salvar = tk.Button(tab1, text="Salvar", width=10, font=large_font)
button_salvar.grid(row=30, column=1, padx=10, pady=20, sticky='ew')

# Finalizando o layout
button_submit = tk.Button(tab1, text="Enviar", command=enviar_email, font=large_font)
button_submit.grid(row=30, column=2, padx=10, pady=20, sticky='ew')

# Adicionar um botão para limpar o formulário
botao_limpar = tk.Button(tab1, text="Limpar", command=limpar_formulario, font=large_font)
botao_limpar.grid(row=30, column=3, padx=10, pady=20, sticky='ew')



# Cabeçalho: DETALHAMENTO
label_detalhamento = tk.Label(tab2, text="DETALHAMENTO", font=large_font, bg='darkblue', fg='white')
label_detalhamento.grid(row=0, column=0, columnspan=6, padx=10, pady=(15, 5), sticky='ew')  # Espaçamento acima

# Bloco: CAIXA \ CEO e ENDEREÇO
label_caixa_ceo = tk.Label(tab2, text="CAIXA \ CEO", font=large_font, bg='darkblue', fg='white')
label_caixa_ceo.grid(row=1, column=0, padx=10, pady=(15, 5), sticky='e')  # Espaçamento acima

entry_caixa_ceo = tk.Entry(tab2, font=large_font)
entry_caixa_ceo.grid(row=1, column=1, padx=10, pady=5, sticky='ew')

label_endereco2 = tk.Label(tab2, text="ENDEREÇO", font=large_font, bg='darkblue', fg='white')
label_endereco2.grid(row=1, column=2, padx=10, pady=(15, 5), sticky='e')  # Espaçamento acima

entry_endereco2 = tk.Entry(tab2, font=large_font)
entry_endereco2.grid(row=1, column=3, padx=10, pady=5, sticky='ew')

# Adicionando o botão com ícone do Maps ao lado do endereço
button_location = tk.PhotoImage(file="icon_maps.png")  # Substitua com o caminho para o ícone desejado
map_button = tk.Button(tab2, image=button_location, command=lambda: obter_localizacao(entry_endereco2))
map_button.grid(row=1, column=4, padx=10, pady=5, sticky='w')

# Função para obter a localização atual usando geocoder
def obter_localizacao(entry):
    # Obtendo a localização usando o geocoder
    location = geocoder.osm(entry.get())

    # Preenchendo latitude e longitude no campo de entrada
    entry.delete(0, tk.END)  # Limpar o campo
    entry.insert(0, f"Lat: {location.lat}, Lng: {location.lng}")

# Cabeçalho: ENTRADA
label_entrada_cabecalho = tk.Label(tab2, text="ENTRADA", font=large_font, bg='darkblue', fg='white')
label_entrada_cabecalho.grid(row=2, column=0, columnspan=6, padx=10, pady=(15, 5), sticky='ew')  # Espaçamento acima

def criar_entrada(linha):
    # Cria um novo conjunto de campos na linha especificada
    label_capacidade = tk.Label(tab2, text="CAPACIDADE", font=large_font)
    label_capacidade.grid(row=linha, column=0, padx=10, pady=5, sticky='e')

    entry_capacidade = tk.Entry(tab2, font=large_font)
    entry_capacidade.grid(row=linha, column=1, padx=10, pady=5, sticky='ew')
    entry_capacidade.bind("<KeyRelease>", lambda event: verificar_entrada(entry_capacidade, linha))

    label_cabo = tk.Label(tab2, text="CABO", font=large_font)
    label_cabo.grid(row=linha, column=2, padx=10, pady=5, sticky='e')

    entry_cabo = tk.Entry(tab2, font=large_font)
    entry_cabo.grid(row=linha, column=3, padx=10, pady=5, sticky='ew')

    label_fibra = tk.Label(tab2, text="FIBRA", font=large_font)
    label_fibra.grid(row=linha, column=4, padx=10, pady=5, sticky='e')

    entry_fibra = tk.Entry(tab2, font=large_font)
    entry_fibra.grid(row=linha, column=5, padx=10, pady=5, sticky='ew')

    # Adiciona as entradas ao dicionário
    entradas[linha] = (entry_capacidade, entry_cabo, entry_fibra)

def verificar_entrada(entry, linha):
    # Verifica se o campo atual já gerou a próxima linha
    if linha + 1 not in entradas and entry.get().strip():  # Só adiciona se não está vazio
        criar_entrada(linha + 1)

# Configuração inicial da interface
root = tk.Tk()
root.title("Gerador de Entradas Dinâmicas")

large_font = ("Arial", 12)

tab2 = tk.Frame(root)
tab2.pack(fill="both", expand=True)

# Dicionário para armazenar todas as entradas
entradas = {}

# Cria as primeiras entradas
criar_entrada(3)

root.mainloop()

# Cabeçalho: SAÍDA
label_saida_cabecalho = tk.Label(tab2, text="SAÍDA", font=large_font, bg='darkblue', fg='white')
label_saida_cabecalho.grid(row=6, column=0, columnspan=6, padx=10, pady=(15, 5), sticky='ew')  # Espaçamento acima

# Cabeçalho: EQUIPE DE FUSÃO
label_equipe_fusao_cabecalho = tk.Label(tab2, text="EQUIPE DE FUSÃO", font=large_font, bg='darkblue', fg='white')
label_equipe_fusao_cabecalho.grid(row=10, column=0, columnspan=6, padx=10, pady=(15, 5), sticky='ew')  # Espaçamento acima

# Centralizar a coluna
tab2.grid_columnconfigure(0, weight=1)
tab2.grid_columnconfigure(1, weight=1)
tab2.grid_columnconfigure(2, weight=1)
tab2.grid_columnconfigure(3, weight=1)


# Iniciar o loop principal da aplicação
tab1.mainloop()

# Função principal para iniciar o formulário
def iniciar_formulario():
    tab1.mainloop()
