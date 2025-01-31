import ttkbootstrap as ttk
from ttkbootstrap.constants import *
import json
import win32print  # Para manipular impressoras no Windows

# Função para carregar configurações do arquivo JSON
def carregar_configuracoes():
    try:
        with open("config.json", "r") as file:
            return json.load(file)
    except FileNotFoundError:
        return {}

# Função para salvar configurações no arquivo JSON
def salvar_configuracoes(config):
    with open("config.json", "w") as file:
        json.dump(config, file)

# Variável global para armazenar o número da NF processado
numero_nf_global = ""

# Função para processar a chave NF
def processar_chave(event=None):
    global numero_nf_global
    chave = entrada_chave.get()
    if chave.isdigit():  # Verifica se a entrada é numérica
        if len(chave) > 9:  # Caso seja uma chave de nota fiscal
            if len(chave) == 44:  # Somente processa se a chave tiver 44 dígitos
                bloco_relevante = chave[25:34]
                numero_nf = bloco_relevante.lstrip("0")
                numero_nf_global = numero_nf  # Atualiza o número processado
                resultado.set(f"Número da NF: {numero_nf}")
                entrada_volumes.focus()  # Mover o foco para o próximo campo
            elif len(chave) < 44:  # Ainda não tem 44 dígitos
                resultado.set("Aguardando a chave completa...")
            else:  # Caso tenha mais de 44 dígitos
                resultado.set("Erro: Tamanho inválido!")
        else:  # Caso seja um número de menos de 9 dígitos (inserção manual)
            numero_nf_global = chave  # Insere o número manualmente digitado
            resultado.set(f"Número da NF: {chave}")
            entrada_volumes.focus()  # Mover o foco para o próximo campo
    else:  # Caso não seja numérico
        resultado.set("Erro: Entrada inválida!")

# Função para enviar ZPL para a impressora
def enviar_para_impressora(event=None):
    global numero_nf_global
    volumes = entrada_volumes.get()
    if volumes.isdigit():
        volumes = int(volumes)
        impressora = configuracoes.get("ultima_impressora")
        if not impressora:
            resultado.set("Erro: Nenhuma impressora selecionada!")
            return

        handle = win32print.OpenPrinter(impressora)
        job = win32print.StartDocPrinter(handle, 1, ("Etiqueta NF", None, "RAW"))
        try:
            for i in range(1, volumes + 1):
                zpl = f"""
^XA
~TA000
~JSN
^LT0
^MNW
^MTT
^PON
^PMN
^LH0,0
^JMA
^PR4,4
~SD15
^JUS
^LRN
^CI27
^PA0,1,1,0
^XZ
^XA
^MMT
^PW719
^LL320
^LS0
^FT28,234^A0N,112,112^FH\^CI28^FDNF:^FS^CI27
^FT191,228^A0N,112,112^FH\^CI28^FD{numero_nf_global}^FS^CI27
^FT546,96^A0N,46,46^FH\^CI28^FDVol.{i}/{volumes}^FS^CI27
^LRY^FO3,275^GB715,0,44^FS
^LRY^FO3,0^GB715,0,42^FS
^LRN
^XZ
                """
                print(f"Etiqueta {i}/{volumes}:\n{zpl}")
                win32print.StartPagePrinter(handle)
                win32print.WritePrinter(handle, zpl.encode("utf-8"))
                win32print.EndPagePrinter(handle)
            resultado.set(f"Impressão de {volumes} etiquetas enviada para {impressora}")
        except Exception as e:
            resultado.set(f"Erro na impressão: {e}")
        finally:
            win32print.EndDocPrinter(handle)
            win32print.ClosePrinter(handle)

        limpar_campos()  # Limpar os campos e preparar para próxima NF
    else:
        resultado.set("Erro: Volume inválido!")

# Função para limpar campos
def limpar_campos():
    entrada_chave.delete(0, END)
    entrada_volumes.delete(0, END)
    resultado.set("")
    entrada_chave.focus()  # Foco no campo de chave/NF

# Função para selecionar impressora
def selecionar_impressora():
    def salvar_impressora():
        selecionada = lista_impressoras.get()
        configuracoes["ultima_impressora"] = selecionada
        salvar_configuracoes(configuracoes)
        janela_impressoras.destroy()

    impressoras = win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL, None, 1)
    nomes_impressoras = [imp[2] for imp in impressoras]

    janela_impressoras = ttk.Toplevel(app)
    janela_impressoras.title("Selecionar Impressora")
    janela_impressoras.geometry("400x300")
    janela_impressoras.transient(app)
    janela_impressoras.grab_set()

    lbl = ttk.Label(janela_impressoras, text="Selecione a impressora desejada:", font=("Helvetica", 12))
    lbl.pack(pady=10)

    lista_impressoras = ttk.Combobox(janela_impressoras, values=nomes_impressoras, font=("Helvetica", 12))
    lista_impressoras.pack(pady=10)
    if "ultima_impressora" in configuracoes:
        lista_impressoras.set(configuracoes["ultima_impressora"])

    btn_salvar = ttk.Button(janela_impressoras, text="Salvar", command=salvar_impressora, bootstyle=SUCCESS)
    btn_salvar.pack(pady=10)
# Função de validação para aceitar apenas números inteiros e até 3 dígitos
def validar_numero_int(P):
    # Verifica se P é vazio ou contém apenas números e se tem até 3 dígitos
    if P == "" or (P.isdigit() and len(P) <= 3):
        return True
    return False

# Carregar configurações ao iniciar o programa
configuracoes = carregar_configuracoes()

# Criando a janela principal
app = ttk.Window(themename="flatly")
app.title("Leitor de Notas Fiscais")
app.geometry("500x400")
app.resizable(False, False)

# Título
titulo = ttk.Label(app, text="", font=("Helvetica", 16, "bold"))
titulo.pack(pady=10)

# Exibição do resultado
resultado = ttk.StringVar()
lbl_resultado = ttk.Label(app, textvariable=resultado, font=("Helvetica", 28,"bold"), bootstyle=INFO)
lbl_resultado.pack(pady=10)

# Frame de entrada
frame_entrada = ttk.Frame(app)
frame_entrada.pack(pady=10, padx=10, fill=X)

# Label e campo para chave ou NF
lbl_chave = ttk.Label(frame_entrada, text="Chave ou NF:", font=("Helvetica", 12,"bold"))
lbl_chave.pack(anchor=S, pady=(0, 2))

entrada_chave = ttk.Entry(frame_entrada, font=("Helvetica", 12,"bold"))
entrada_chave.pack(side=TOP, fill=Y, expand=False, padx=5)
entrada_chave.bind("<Return>", processar_chave)
entrada_chave.focus()

# Frame para volumes
frame_volumes = ttk.Frame(app)
frame_volumes.pack(pady=10, padx=10, fill=Y)

# Label e campo para volumes
lbl_volumes = ttk.Label(frame_volumes, text="Volume:", font=("Helvetica", 12,"bold"))
lbl_volumes.pack(anchor=S, pady=(0, 2))

# Definindo o comando de validação
validacao = app.register(validar_numero_int)

# Criando o campo de entrada com validação
entrada_volumes = ttk.Entry(frame_volumes, font=("Helvetica", 12, "bold"),
                             validate="key", validatecommand=(validacao, "%P"))
entrada_volumes.pack(side=TOP, fill=Y, expand=False, padx=5)
entrada_volumes.bind("<Return>", enviar_para_impressora)
# Frame para seleção de impressora
frame_impressora = ttk.Frame(app)
frame_impressora.pack(pady=5, padx=50, fill=X)

# Label e lista suspensa de impressoras
lbl_impressora = ttk.Label(frame_impressora, text="Impressora:", font=("Helvetica", 8,"bold"))
lbl_impressora.pack(side=TOP, padx=5)

impressoras = win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL, None, 1)
nomes_impressoras = [imp[2] for imp in impressoras]

lista_impressoras = ttk.Combobox(frame_impressora, values=nomes_impressoras, font=("Helvetica", 8,"bold"))
lista_impressoras.pack(side=TOP, fill=X, expand=False, padx=5)

# Configuração da impressora selecionada
if "ultima_impressora" in configuracoes:
    lista_impressoras.set(configuracoes["ultima_impressora"])
else:
    lista_impressoras.set(nomes_impressoras[0] if nomes_impressoras else "")

# Botões de controle
frame_botoes = ttk.Frame(app)
frame_botoes.pack(pady=20)

# Botão de impressão
btn_imprimir = ttk.Button(frame_botoes, text="Imprimir", command=enviar_para_impressora, bootstyle=PRIMARY)
btn_imprimir.pack(side=LEFT, padx=10)

btn_limpar = ttk.Button(frame_botoes, text="Limpar", command=limpar_campos, bootstyle=SECONDARY)
btn_limpar.pack(side=LEFT, padx=10)

btn_fechar = ttk.Button(frame_botoes, text="Fechar", command=app.quit, bootstyle=DANGER)
btn_fechar.pack(side=LEFT, padx=10)

# Iniciar a aplicação
app.mainloop()