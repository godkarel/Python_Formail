import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
import re
import csv
import smtplib

from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

def montar_mensagem():
    # Coleta os dados dos campos de entrada
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
        "Tecnologia": entry_tecnologia.get(),
        "HUB / SITE": entry_hub_site.get(),
        "Rota / Anel": entry_rota_anel.get(),
        "Cabo": entry_cabo.get(),
        "Rack": entry_rack.get(),
        "Frame": entry_frame.get(),
        "Splitter Primário": entry_splitter_primario.get(),
        "Splitter Secundário": entry_splitter_secundario.get(),
        "Posição": entry_posicao.get(),
        "Fibra": entry_fibra.get(),
        "Atenuação": entry_atenuacao.get(),
        "Tipo de Caixa a ser Instalada": entry_tipo_caixa.get(),
        "Terminação Projetada (unid.)": entry_terminacao_proj.get(),
        "Terminação Realizada (unid.)": entry_terminacao_real.get(),
        "Sangria Projetada (unid.)": entry_sangria_proj.get(),
        "Sangria Realizada (unid.)": entry_sangria_real.get(),
        "Lançamento Cabo Projetado (m)": entry_lancamento_cabo_proj.get(),
        "Lançamento Cabo Realizado (m)": entry_lancamento_cabo_real.get(),
        "Conectorização Projetada (unid.)": entry_conector_proj.get(),
        "Conectorização Realizada (unid.)": entry_conector_real.get(),
        "Fusão Cabo Projetado (unid.)": entry_fusao_cabo_proj.get(),
        "Fusão Cabo Realizado (unid.)": entry_fusao_cabo_real.get(),
        "Fibra Protegida Projetada (m)": entry_fibra_prot_proj.get(),
        "Fibra Protegida Realizada (m)": entry_fibra_prot_real.get(),
        "Pontas Reservadas Projetadas (unid.)": entry_pontas_proj.get(),
        "Pontas Reservadas Realizadas (unid.)": entry_pontas_real.get(),
        "Lançamento Tubete Projetado (m)": entry_lancamento_tubete_proj.get(),
        "Lançamento Tubete Realizado (m)": entry_lancamento_tubete_real.get(),
        "Lançamento Duto Projetado (m)": entry_lancamento_duto_proj.get(),
        "Lançamento Duto Realizado (m)": entry_lancamento_duto_real.get(),
        "Pendências": entry_pendencias.get(),
        "Comentários": entry_comentarios.get(),
    }

    # Monta a mensagem apenas com campos não vazios
    mensagem = "Relatório de Gastos OS\n\n"
    for titulo, valor in campos.items():
        if valor:  # Verifica se o campo não está vazio
            mensagem += f"{titulo}: {valor}\n"

    return mensagem

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
label_tecnologia = tk.Label(tab1, text="Tecnologia")
label_tecnologia.grid(row=13, column=0, padx=10, pady=5, sticky='ew')
entry_tecnologia = tk.Entry(tab1, justify='center')
entry_tecnologia.grid(row=14, column=0, padx=10, pady=5, sticky='ew')

label_hub_site = tk.Label(tab1, text="HUB / SITE")
label_hub_site.grid(row=13, column=1, padx=10, pady=5, sticky='ew')
entry_hub_site = tk.Entry(tab1, justify='center')
entry_hub_site.grid(row=14, column=1, padx=10, pady=5, sticky='ew')

label_rota_anel = tk.Label(tab1, text="Rota / Anel")
label_rota_anel.grid(row=13, column=2, padx=10, pady=5, sticky='ew')
entry_rota_anel = tk.Entry(tab1, justify='center')
entry_rota_anel.grid(row=14, column=2, padx=10, pady=5, sticky='ew')

label_cabo = tk.Label(tab1, text="Cabo")
label_cabo.grid(row=13, column=3, padx=10, pady=5, sticky='ew')
entry_cabo = tk.Entry(tab1, justify='center')
entry_cabo.grid(row=14, column=3, padx=10, pady=5, sticky='ew')

# Bloco 5
label_rack = tk.Label(tab1, text="Rack")
label_rack.grid(row=15, column=0, padx=10, pady=5, sticky='ew')
entry_rack = tk.Entry(tab1, justify='center')
entry_rack.grid(row=16, column=0, padx=10, pady=5, sticky='ew')

label_frame = tk.Label(tab1, text="Frame")
label_frame.grid(row=15, column=1, padx=10, pady=5, sticky='ew')
entry_frame = tk.Entry(tab1, justify='center')
entry_frame.grid(row=16, column=1, padx=10, pady=5, sticky='ew')

label_splitter_primario = tk.Label(tab1, text="Splitter Primário")
label_splitter_primario.grid(row=15, column=2, padx=10, pady=5, sticky='ew')
entry_splitter_primario = tk.Entry(tab1, justify='center')
entry_splitter_primario.grid(row=16, column=2, padx=10, pady=5, sticky='ew')

label_splitter_secundario = tk.Label(tab1, text="Splitter Secundário")
label_splitter_secundario.grid(row=15, column=3, padx=10, pady=5, sticky='ew')
entry_splitter_secundario = tk.Entry(tab1, justify='center')
entry_splitter_secundario.grid(row=16, column=3, padx=10, pady=5, sticky='ew')

label_posicao = tk.Label(tab1, text="Posição")
label_posicao.grid(row=17, column=0, padx=10, pady=5, sticky='ew')
entry_posicao = tk.Entry(tab1, justify='center')
entry_posicao.grid(row=18, column=0, padx=10, pady=5, sticky='ew')

label_fibra = tk.Label(tab1, text="Fibra")
label_fibra.grid(row=17, column=1, padx=10, pady=5, sticky='ew')
entry_fibra = tk.Entry(tab1, justify='center')
entry_fibra.grid(row=18, column=1, padx=10, pady=5, sticky='ew')

label_atenuacao = tk.Label(tab1, text="Atenuação")
label_atenuacao.grid(row=17, column=2, padx=10, pady=5, sticky='ew')
entry_atenuacao = tk.Entry(tab1, justify='center')
entry_atenuacao.grid(row=18, column=2, padx=10, pady=5, sticky='ew')

label_tipo_caixa = tk.Label(tab1, text="Tipo de Caixa a ser Instalada")
label_tipo_caixa.grid(row=17, column=3, padx=10, pady=5, sticky='ew')
entry_tipo_caixa = tk.Entry(tab1, justify='center')
entry_tipo_caixa.grid(row=18, column=3, padx=10, pady=5, sticky='ew')

# Bloco 6
label_terminacao_proj = tk.Label(tab1, text="Terminação Projetada (unid.)")
label_terminacao_proj.grid(row=19, column=0, padx=10, pady=5, sticky='ew')
entry_terminacao_proj = tk.Entry(tab1, justify='center')
entry_terminacao_proj.grid(row=20, column=0, padx=10, pady=5, sticky='ew')

label_terminacao_real = tk.Label(tab1, text="Terminação Realizada (unid.)")
label_terminacao_real.grid(row=19, column=1, padx=10, pady=5, sticky='ew')
entry_terminacao_real = tk.Entry(tab1, justify='center')
entry_terminacao_real.grid(row=20, column=1, padx=10, pady=5, sticky='ew')

label_sangria_proj = tk.Label(tab1, text="Sangria Projetada (unid.)")
label_sangria_proj.grid(row=19, column=2, padx=10, pady=5, sticky='ew')
entry_sangria_proj = tk.Entry(tab1, justify='center')
entry_sangria_proj.grid(row=20, column=2, padx=10, pady=5, sticky='ew')

label_sangria_real = tk.Label(tab1, text="Sangria Realizada (unid.)")
label_sangria_real.grid(row=19, column=3, padx=10, pady=5, sticky='ew')
entry_sangria_real = tk.Entry(tab1, justify='center')
entry_sangria_real.grid(row=20, column=3, padx=10, pady=5, sticky='ew')

# Bloco 7# Bloco 7# Bloco 7# Bloco 7# Bloco 7# Bloco 7# Bloco 7# Bloco 7# Bloco 7# Bloco 7# Bloco 7# Bloco 7
label_lancamento_cabo_proj = tk.Label(tab1, text="Lançamento Cabo Projetado (m)")
label_lancamento_cabo_proj.grid(row=21, column=0, padx=10, pady=5, sticky='ew')
entry_lancamento_cabo_proj = tk.Entry(tab1, justify='center')
entry_lancamento_cabo_proj.grid(row=22, column=0, padx=10, pady=5, sticky='ew')

label_lancamento_cabo_real = tk.Label(tab1, text="Lançamento Cabo Realizado (m)")
label_lancamento_cabo_real.grid(row=21, column=1, padx=10, pady=5, sticky='ew')
entry_lancamento_cabo_real = tk.Entry(tab1, justify='center')
entry_lancamento_cabo_real.grid(row=22, column=1, padx=10, pady=5, sticky='ew')

label_conector_proj = tk.Label(tab1, text="Conectorização Projetada (unid.)")
label_conector_proj.grid(row=21, column=2, padx=10, pady=5, sticky='ew')
entry_conector_proj = tk.Entry(tab1, justify='center')
entry_conector_proj.grid(row=22, column=2, padx=10, pady=5, sticky='ew')

label_conector_real = tk.Label(tab1, text="Conectorização Realizada (unid.)")
label_conector_real.grid(row=21, column=3, padx=10, pady=5, sticky='ew')
entry_conector_real = tk.Entry(tab1, justify='center')
entry_conector_real.grid(row=22, column=3, padx=10, pady=5, sticky='ew')

# Bloco 8
label_fusao_cabo_proj = tk.Label(tab1, text="Fusão Cabo Projetado (unid.)")
label_fusao_cabo_proj.grid(row=23, column=0, padx=10, pady=5, sticky='ew')
entry_fusao_cabo_proj = tk.Entry(tab1, justify='center')
entry_fusao_cabo_proj.grid(row=24, column=0, padx=10, pady=5, sticky='ew')

label_fusao_cabo_real = tk.Label(tab1, text="Fusão Cabo Realizado (unid.)")
label_fusao_cabo_real.grid(row=23, column=1, padx=10, pady=5, sticky='ew')
entry_fusao_cabo_real = tk.Entry(tab1, justify='center')
entry_fusao_cabo_real.grid(row=24, column=1, padx=10, pady=5, sticky='ew')

label_fibra_prot_proj = tk.Label(tab1, text="Fibra Protegida Projetada (m)")
label_fibra_prot_proj.grid(row=23, column=2, padx=10, pady=5, sticky='ew')
entry_fibra_prot_proj = tk.Entry(tab1, justify='center')
entry_fibra_prot_proj.grid(row=24, column=2, padx=10, pady=5, sticky='ew')

label_fibra_prot_real = tk.Label(tab1, text="Fibra Protegida Realizada (m)")
label_fibra_prot_real.grid(row=23, column=3, padx=10, pady=5, sticky='ew')
entry_fibra_prot_real = tk.Entry(tab1, justify='center')
entry_fibra_prot_real.grid(row=24, column=3, padx=10, pady=5, sticky='ew')

# Bloco 9
label_pontas_proj = tk.Label(tab1, text="Pontas Reservadas Projetadas (unid.)")
label_pontas_proj.grid(row=25, column=0, padx=10, pady=5, sticky='ew')
entry_pontas_proj = tk.Entry(tab1, justify='center')
entry_pontas_proj.grid(row=26, column=0, padx=10, pady=5, sticky='ew')

label_pontas_real = tk.Label(tab1, text="Pontas Reservadas Realizadas (unid.)")
label_pontas_real.grid(row=25, column=1, padx=10, pady=5, sticky='ew')
entry_pontas_real = tk.Entry(tab1, justify='center')
entry_pontas_real.grid(row=26, column=1, padx=10, pady=5, sticky='ew')

label_lancamento_tubete_proj = tk.Label(tab1, text="Lançamento Tubete Projetado (m)")
label_lancamento_tubete_proj.grid(row=25, column=2, padx=10, pady=5, sticky='ew')
entry_lancamento_tubete_proj = tk.Entry(tab1, justify='center')
entry_lancamento_tubete_proj.grid(row=26, column=2, padx=10, pady=5, sticky='ew')

label_lancamento_tubete_real = tk.Label(tab1, text="Lançamento Tubete Realizado (m)")
label_lancamento_tubete_real.grid(row=25, column=3, padx=10, pady=5, sticky='ew')
entry_lancamento_tubete_real = tk.Entry(tab1, justify='center')
entry_lancamento_tubete_real.grid(row=26, column=3, padx=10, pady=5, sticky='ew')

# Bloco 10
label_lancamento_duto_proj = tk.Label(tab1, text="Lançamento Duto Projetado (m)")
label_lancamento_duto_proj.grid(row=27, column=0, padx=10, pady=5, sticky='ew')
entry_lancamento_duto_proj = tk.Entry(tab1, justify='center')
entry_lancamento_duto_proj.grid(row=28, column=0, padx=10, pady=5, sticky='ew')

label_lancamento_duto_real = tk.Label(tab1, text="Lançamento Duto Realizado (m)")
label_lancamento_duto_real.grid(row=27, column=1, padx=10, pady=5, sticky='ew')
entry_lancamento_duto_real = tk.Entry(tab1, justify='center')
entry_lancamento_duto_real.grid(row=28, column=1, padx=10, pady=5, sticky='ew')

label_pendencias = tk.Label(tab1, text="Pendências")
label_pendencias.grid(row=27, column=2, padx=10, pady=5, sticky='ew')
entry_pendencias = tk.Entry(tab1, justify='center')
entry_pendencias.grid(row=28, column=2, padx=10, pady=5, sticky='ew')

label_comentarios = tk.Label(tab1, text="Comentários")
label_comentarios.grid(row=27, column=3, padx=10, pady=5, sticky='ew')
entry_comentarios = tk.Entry(tab1, justify='center')
entry_comentarios.grid(row=28, column=3, padx=10, pady=5, sticky='ew')


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

# Cabeçalho: ENTRADA
label_entrada_cabecalho = tk.Label(tab2, text="ENTRADA", font=large_font, bg='darkblue', fg='white')
label_entrada_cabecalho.grid(row=2, column=0, columnspan=6, padx=10, pady=(15, 5), sticky='ew')  # Espaçamento acima

# Informações de ENTRADA na mesma linha
for i in range(3):  # Adiciona três entradas
    row = 3 + i  # Ajusta a linha para cada entrada
    label_capacidade = tk.Label(tab2, text="CAPACIDADE", font=large_font)
    label_capacidade.grid(row=row, column=0, padx=10, pady=5, sticky='e')

    entry_capacidade = tk.Entry(tab2, font=large_font)
    entry_capacidade.grid(row=row, column=1, padx=10, pady=5, sticky='ew')

    label_cabo1 = tk.Label(tab2, text="CABO", font=large_font)
    label_cabo1.grid(row=row, column=2, padx=10, pady=5, sticky='e')

    entry_cabo1 = tk.Entry(tab2, font=large_font)
    entry_cabo1.grid(row=row, column=3, padx=10, pady=5, sticky='ew')

    label_fibra1 = tk.Label(tab2, text="FIBRA", font=large_font)
    label_fibra1.grid(row=row, column=4, padx=10, pady=5, sticky='e')

    entry_fibra1 = tk.Entry(tab2, font=large_font)
    entry_fibra1.grid(row=row, column=5, padx=10, pady=5, sticky='ew')

# Cabeçalho: SAÍDA
label_saida_cabecalho = tk.Label(tab2, text="SAÍDA", font=large_font, bg='darkblue', fg='white')
label_saida_cabecalho.grid(row=6, column=0, columnspan=6, padx=10, pady=(15, 5), sticky='ew')  # Espaçamento acima

# Informações de SAÍDA na mesma linha
for i in range(3):  # Adiciona três entradas
    row = 7 + i  # Ajusta a linha para cada entrada
    label_capacidade_saida = tk.Label(tab2, text="CAPACIDADE", font=large_font)
    label_capacidade_saida.grid(row=row, column=0, padx=10, pady=5, sticky='e')

    entry_capacidade_saida = tk.Entry(tab2, font=large_font)
    entry_capacidade_saida.grid(row=row, column=1, padx=10, pady=5, sticky='ew')

    label_cabo_saida = tk.Label(tab2, text="CABO", font=large_font)
    label_cabo_saida.grid(row=row, column=2, padx=10, pady=5, sticky='e')

    entry_cabo_saida = tk.Entry(tab2, font=large_font)
    entry_cabo_saida.grid(row=row, column=3, padx=10, pady=5, sticky='ew')

    label_fibra_saida = tk.Label(tab2, text="FIBRA", font=large_font)
    label_fibra_saida.grid(row=row, column=4, padx=10, pady=5, sticky='e')

    entry_fibra_saida = tk.Entry(tab2, font=large_font)
    entry_fibra_saida.grid(row=row, column=5, padx=10, pady=5, sticky='ew')

# Cabeçalho: EQUIPE DE FUSÃO
label_equipe_fusao_cabecalho = tk.Label(tab2, text="EQUIPE DE FUSÃO", font=large_font, bg='darkblue', fg='white')
label_equipe_fusao_cabecalho.grid(row=10, column=0, columnspan=6, padx=10, pady=(15, 5), sticky='ew')  # Espaçamento acima

# Espaçamento entre o cabeçalho e as entradas
label_empty_space = tk.Label(tab2, text="", font=large_font)  # Espaço vazio
label_empty_space.grid(row=11, column=0, padx=10, pady=(5, 15), sticky='ew')  # Adiciona um espaço em branco

# Informações da EQUIPE DE FUSÃO na mesma linha
for i in range(3):  # Adiciona três entradas
    row = 11 + i  # Ajusta a linha para cada entrada
    label_tipo_caixa1 = tk.Label(tab2, text="TIPO DE CAIXA", font=large_font)
    label_tipo_caixa1.grid(row=row, column=0, padx=10, pady=5, sticky='e')

    entry_tipo_caixa1 = tk.Entry(tab2, font=large_font)
    entry_tipo_caixa1.grid(row=row, column=1, padx=10, pady=5, sticky='ew')

    label_situacao_caixa = tk.Label(tab2, text="SITUAÇÃO CAIXA", font=large_font)
    label_situacao_caixa.grid(row=row, column=2, padx=10, pady=5, sticky='e')

    entry_situacao_caixa = tk.Entry(tab2, font=large_font)
    entry_situacao_caixa.grid(row=row, column=3, padx=10, pady=5, sticky='ew')

    label_observacao = tk.Label(tab2, text="OBSERVAÇÃO", font=large_font)
    label_observacao.grid(row=row, column=4, padx=10, pady=5, sticky='e')

    entry_observacao = tk.Entry(tab2, font=large_font)
    entry_observacao.grid(row=row, column=5, padx=10, pady=5, sticky='ew')

# Mensagem
label_mensagem = tk.Label(tab2, text="Mensagem", font=large_font)
label_mensagem.grid(row=14, column=0, columnspan=4, padx=10, pady=(15, 5), sticky='ew')  # Espaçamento acima

entry_mensagem = tk.Text(tab2, height=5, font=large_font)
entry_mensagem.grid(row=15, column=0, columnspan=4, padx=10, pady=5, sticky='ew')  # Centralizando com columnspan

# Botões de ação (Opcional: Salvar, Enviar, etc.)
button_salvar = tk.Button(tab2, text="Salvar", width=10)
button_salvar.grid(row=17, column=1, padx=10, pady=20, sticky='ew')

# Finalizando o layout
button_submit = tk.Button(tab2, text="Enviar", command=enviar_email, font=large_font)
button_submit.grid(row=17, column=2, padx=10, pady=20, sticky='ew')

def limpar_formulario():
    entry_nome.delete(0, tk.END)
    entry_email.delete(0, tk.END)
    entry_mensagem.delete("1.0", tk.END)
    # Limpe os outros campos também...

# Adicionar um botão para limpar o formulário
botao_limpar = tk.Button(tab2, text="Limpar", command=limpar_formulario, font=large_font)
botao_limpar.grid(row=17, column=3, padx=10, pady=20, sticky='ew')


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
