import pandas as pd
import datetime
import pywhatkit as kit
import os
import tkinter as tk
from tkinter import simpledialog, messagebox

def cadastrar_cliente():
    nome = simpledialog.askstring("Input", "Nome do Cliente:")
    processo = simpledialog.askstring("Input", "Número do Processo:")
    
    # Obter valor total do contrato
    valor_input = simpledialog.askstring("Input", "Valor Total do Contrato (use ponto ou vírgula): R$")
    valor_input = valor_input.replace(',', '.')  # Trocar vírgula por ponto
    valor_total_contrato = float(valor_input)  # Converte para float
    
    # Define a data de pagamento da 1ª parcela
    data_primeira_parcela = simpledialog.askstring("Input", "Data de pagamento da 1ª parcela (dd/mm/aaaa):")
    data_primeira_parcela = datetime.datetime.strptime(data_primeira_parcela, "%d/%m/%Y")
    
    parcelas = simpledialog.askinteger("Input", "Quantidade de Parcelas:")
    
    # Calcula o valor de cada parcela
    valor_parcela = valor_total_contrato / parcelas
    
    celular = simpledialog.askstring("Input", "Número do Celular (com DDD):")

    # Calcula datas de vencimento
    datas_vencimento = [data_primeira_parcela + datetime.timedelta(days=30*i) for i in range(parcelas)]
    
    # Cria um dicionário para os dados do cliente
    cliente_data = {
        "Nome": nome,
        "Processo": processo,
        "Valor Total do Contrato": valor_total_contrato,
        "Valor de Cada Parcela": valor_parcela,
        "Data Primeira Parcela": data_primeira_parcela,
        "Parcelas": parcelas,
        "Celular": celular
    }
    
    # Adiciona cada data de vencimento em uma coluna separada
    for idx, data in enumerate(datas_vencimento):
        cliente_data[f"Data Parcela {idx+1}"] = data

    return cliente_data

def enviar_mensagem(celular, mensagem):
    # Formata o celular para garantir que está no formato correto
    celular_formatado = f"{celular}"
    try:
        kit.sendwhatmsg(f"+{celular_formatado}", mensagem, datetime.datetime.now().hour, datetime.datetime.now().minute + 2)
        messagebox.showinfo("Info", "Mensagem enviada com sucesso!")
    except Exception as e:
        messagebox.showerror("Erro", f"Falha ao enviar mensagem: {e}")

def main():
    clientes = []
    celulares_vencidos = []  # Lista para armazenar celulares com parcelas vencidas
    caminho_diretorio = "clientes.xlsx"

    # Tenta ler dados existentes para verificar vencimentos
    if os.path.exists(caminho_diretorio):
        df_existente = pd.read_excel(caminho_diretorio, engine='openpyxl')
        clientes_existentes = df_existente.to_dict(orient='records')
        # Checar vencimentos
        hoje = datetime.datetime.now().date()
        for cliente in clientes_existentes:
            for key in cliente.keys():
                if "Data Parcela" in key:
                    data_venc = cliente[key]
                    if isinstance(data_venc, pd.Timestamp):  # Verifica se é tipo Timestamp
                        data_venc = data_venc.to_pydatetime().date()  # Converte para objeto date
                        # Exibe alerta se a data de vencimento é anterior à data de hoje
                        if data_venc < hoje:
                            alert_message = f"Parcela vencida para {cliente['Nome']} em {data_venc.strftime('%d/%m/%Y')}!"
                            messagebox.showwarning("Alerta de Vencimento", alert_message)
                            celulares_vencidos.append(cliente['Celular'])  # Adiciona celular à lista
                        elif data_venc == hoje:
                            messagebox.showwarning("Alerta de Vencimento", f"Parcela vencendo hoje para {cliente['Nome']} em {data_venc.strftime('%d/%m/%Y')}")
    else:
        clientes_existentes = []

    # Pergunta se o usuário deseja enviar mensagens para parcelas vencidas antes de cadastrar novos clientes
    if celulares_vencidos:
        enviar_mensagem_vencidos = messagebox.askyesno("Enviar Mensagens Vencidas", "Deseja enviar mensagens para as parcelas vencidas?")
        if enviar_mensagem_vencidos:
            for celular in celulares_vencidos:
                mensagem_vencida = (f"Olá, aqui é da Ramos Campos Sociedade de Advogados \n\n"
                                    "Lembramos que sua(s) parcela(s) estão vencida(s). "
                                    "Por favor, entre em contato para regularizar sua situação, a fim de evitar a interrupção dos serviços jurídicos nos termos do art. 190 do CPC.")
                enviar_mensagem(celular, mensagem_vencida)

    while True:
        cliente = cadastrar_cliente()
        clientes.append(cliente)

        enviar = messagebox.askyesno("Enviar Mensagem", "Deseja enviar mensagem de cobrança?")
        if enviar:
            data_parcela_1 = cliente['Data Primeira Parcela'].strftime('%d/%m/%Y')
            mensagem = (f"Olá {cliente['Nome']},\n\n"
                        f"Lembramos que o pagamento do contrato referente ao processo {cliente['Processo']} "
                        f"no valor total de R$ {cliente['Valor Total do Contrato']:.2f} está programado para "
                        f"vencer em {data_parcela_1}.\n"
                        f"Valor de cada parcela é de R$ {cliente['Valor de Cada Parcela']:.2f}.")
            enviar_mensagem(cliente['Celular'], mensagem)
            messagebox.showinfo("Info", "Mensagem enviada!")

        continuar = messagebox.askyesno("Cadastrar Cliente", "Deseja cadastrar outro cliente?")
        if not continuar:
            break

    # Combina os novos clientes com os existentes
    clientes_existentes.extend(clientes)
    df_final = pd.DataFrame(clientes_existentes)

    # Salvar os dados em um arquivo Excel
    df_final.to_excel(caminho_diretorio, index=False, engine='openpyxl')
    messagebox.showinfo("Info", f"Dados salvos em {caminho_diretorio}")

if __name__ == "__main__":
    # Inicializa a aplicação Tkinter
    root = tk.Tk()
    root.withdraw()  # Não mostra a janela principal
    main()

