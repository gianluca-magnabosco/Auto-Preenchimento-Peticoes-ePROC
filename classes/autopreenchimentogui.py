import os
import re
import tkinter as tk
from tkinter import messagebox
from tkinter import ttk
from tkinter.constants import CENTER
from functools import partial
from functions.aux_functions import connectDataBase
from classes.processos import Processos
from classes.documentos import Documentos


class AutoPreenchimentoGUI(Processos):

    def centerWindow(self, width, height, window):
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x = (screen_width / 2) - (width / 2)
        y = (screen_height / 2) - (height / 2)
        window.geometry("%dx%d+%d+%d" % (width, height, x, y))


    def onCloseRoot(self):
        close = messagebox.askokcancel("Confirmação", "Tem certeza que deseja fechar o programa?")
        if close:
            self.root.destroy()


    def initGUI(self):
        self.root = tk.Tk()

        iconFile = "img/icon.ico"
        self.centerWindow(860, 640, self.root)
        self.root.resizable(False, False)
        self.root.title("Auto Peticionamento")

        self.root.protocol("WM_DELETE_WINDOW", self.onCloseRoot)
        self.root.iconbitmap(default = iconFile)
        rootBackgroundImage = tk.PhotoImage(file = "img/rootBackground.png")
        rootBackgroundLabel = tk.Label(self.root, image = rootBackgroundImage, bg = "white")
        rootBackgroundLabel.place(relx = 0.5, rely = 0.5, anchor = CENTER)

        self.initComponents()
        self.getData()

        self.gerarDocumento()

        self.root.mainloop()


    def initComponents(self):
        self.initButtons()
        self.initLabels()
        self.gerarWindowBackGround = tk.PhotoImage(file = "img/gerarWindowBackground.png")
        self.alterarWindowBackGround = tk.PhotoImage(file = "img/alterarWindowBackground.png")
        self.adicionarWindowBackGround = tk.PhotoImage(file = "img/adicionarWindowBackground.png")

    
    def initButtons(self):
        self.buttonStyle = ttk.Style()
        self.buttonStyle.configure("E.TButton", background = "white", foreground = "black", font = ("Helvetica", 12))
        self.buttonStyle.configure("W.TButton", background = "white", foreground = "black", font = ("Arial", 9))
        
        self.abrirButton = ttk.Button(self.root, style = "W.TButton", text = "Abrir", command = self.openDirectory, width = 27.75)
        self.abrirButton.pack()
        self.abrirButton.place(relx = 0.27, rely = 0.85, anchor = CENTER)

        self.alterarButton = ttk.Button(self.root, style = "W.TButton", text= "Alterar", command = self.alterarCliente, width = 27.75)
        self.alterarButton.pack()
        self.alterarButton.place(relx = 0.83, rely = 0.44, anchor = CENTER)

        self.adicionarButton = ttk.Button(self.root, style = "W.TButton", text = "Adicionar", command = self.adicionarProcesso, width = 27.75)
        self.adicionarButton.pack()
        self.adicionarButton.place(relx = 0.81, rely = 0.73, anchor = CENTER)

        self.gerarButton = ttk.Button(self.root, style = "E.TButton", text = "Gerar Petições", command = self.gerarDocumento, width = 27.75)
        self.gerarButton.pack()
        self.gerarButton.place(relx = 0.46, rely = 0.665, anchor = CENTER)


    def initLabels(self):
        self.creditsLabel = tk.Label(text = "Programa criado por: Gianluca Notari Magnabosco da Silva", font = ("", 7), bg = "white")
        self.creditsLabel.pack()
        self.creditsLabel.place(relx = 0.84, rely = 0.98, anchor = CENTER)


    def validarNumeroProcesso(self, num_processo, window):
        num_processo = num_processo.strip()

        if len(num_processo) > 25:
            num_processo = num_processo[:25]

        if not re.match("^\d{7}\-?\d{2}\.?\d{4}\.?\d\.?\d{2}\.?\d{4}", num_processo):
            messagebox.showerror(title = "Erro", message = "Processo inválido!") 
            window.attributes("-topmost", True)
            return None
        
        if len(num_processo) < 25 and len(num_processo) >= 20:
            edit_num_processo = list(str(num_processo))
            
            if edit_num_processo[7] != '-':
                edit_num_processo.insert(7, '-')
            
            if edit_num_processo[10] != '.':
                edit_num_processo.insert(10, '.')
            
            if edit_num_processo[15] != '.':
                edit_num_processo.insert(15, '.')

            if edit_num_processo[17] != '.':
                edit_num_processo.insert(17, '.')

            if edit_num_processo[20] != '.':
                edit_num_processo.insert(20, '.')

            num_processo = "".join(edit_num_processo)
        
        
        return num_processo


    def openDirectory(self):
        localPath = os.getcwd()
        os.startfile(os.path.join(localPath, "arquivos"))

    
    def validarAlteracaoCliente(self, num_processo):
        self.alterarWindow.attributes("-topmost", False)
        processo_atual = num_processo.get()

        processo_atual = self.validarNumeroProcesso(processo_atual, self.alterarWindow)

        if processo_atual is None:
            return

        con = connectDataBase()
        cur = con.cursor()
        
        cur.execute(f"SELECT * FROM processos WHERE num_processo = '{processo_atual}';")
        result = cur.fetchall()

        if not result:
            messagebox.showerror(title = "Erro", message = "Processo não encontrado!") 
            self.alterarWindow.attributes("-topmost", True)
            return


        cliente_atual = result[0][1]
        parte_adversa_atual = result[0][2]
        
        is_cliente = messagebox.askyesno(title = f"Processo nº {processo_atual}", message = f"{cliente_atual} é seu cliente nesse processo?")
        
        if not is_cliente:
            cur.execute(f"UPDATE processos SET cliente = '{parte_adversa_atual}' WHERE num_processo = '{processo_atual}';")
            cur.execute(f"UPDATE processos SET parte_adversa = '{cliente_atual}' WHERE num_processo = '{processo_atual}';")
            messagebox.showinfo(title = "Sucesso!", message = "Processo atualizado com sucesso!")
            
        if is_cliente:
            messagebox.showerror(title = "Nada foi alterado", message = "Nada foi atualizado!")

        con.commit()
        cur.close()
        con.close()

        self.onCloseAlterarWindow()


    def onCloseAlterarWindow(self):
        self.alterarWindow.destroy()
        self.root.attributes("-topmost", True)
        self.root.wm_state("normal")
    

    def alterarCliente(self):
        self.root.wm_state("iconic")

        self.alterarWindow = tk.Toplevel(self.root)
        self.alterarWindow.title("Alterar cliente")

        self.centerWindow(350, 135, self.alterarWindow)
        self.alterarWindow.resizable(False, False)
        self.alterarWindow.protocol("WM_DELETE_WINDOW", self.onCloseAlterarWindow)
        backgroundLabel = tk.Label(self.alterarWindow, image = self.alterarWindowBackGround, bg = "white")
        backgroundLabel.place(relx = 0.5, rely = 0.5, anchor = CENTER)

        num_processo = tk.StringVar()
        self.alterarWindowEntry = ttk.Entry(self.alterarWindow, textvariable = num_processo)
        self.alterarWindowEntry.place(relx = 0.635, rely = 0.475, anchor = CENTER, width = 200)

        validateInput = partial(self.validarAlteracaoCliente, num_processo) 

        buttonStyle = ttk.Style()
        buttonStyle.configure("W.TButton", background = "#a3cae7", foreground = "black", font = ("Open Sans", 9))
        self.alterarWindowButton = ttk.Button(self.alterarWindow, style = "W.TButton", text = "Confirma", command = validateInput)
        self.alterarWindowButton.place(relx = 0.5, rely = 0.79, anchor = CENTER, width = 60) 
            

        # add RETURN key handler
        def handler(e):
            validateInput()
        self.alterarWindow.bind("<Return>", handler)



    def validarAdicaoProcesso(self, num_processo, cliente_input, adversa_input, cidade_input):
        self.adicionarWindow.attributes("-topmost", False)
        
        processo_atual = num_processo.get()
        cliente = cliente_input.get()
        adversa = adversa_input.get()
        cidade = cidade_input.get()
        
        processo_atual = self.validarNumeroProcesso(processo_atual, self.adicionarWindow)

        if processo_atual is None:
            return
        
        if len(cliente) < 5 or len(adversa) < 5 or len(cidade) < 5:
            messagebox.showerror(title = "Erro", message = "Erro!\nValor(es) inválidos")
            self.adicionarWindow.attributes("-topmost", True)   
            return

        con = connectDataBase()
        cur = con.cursor()
        
        try: 
            cur.execute(f"INSERT INTO processos(num_processo, cliente, parte_adversa, cidade) VALUES ('{processo_atual}', '{cliente}', '{adversa}', '{cidade}');")
        except:
            messagebox.showwarning(title = "Erro", message = "Erro!\nTente novamente")
            self.adicionarWindow.attributes("-topmost", True)
            return          

        con.commit()
        cur.close()
        con.close()

        messagebox.showinfo(title = "Sucesso!", message = "Processo adicionado com sucesso!")                

        self.onCloseAdicionarWindow()


    def onCloseAdicionarWindow(self):
        self.adicionarWindow.destroy()
        self.root.attributes("-topmost", True)
        self.root.wm_state("normal")


    def adicionarProcesso(self):
        self.root.wm_state("iconic")

        self.adicionarWindow = tk.Toplevel(self.root)
        self.adicionarWindow.title("Adicionar processo")

        self.centerWindow(525, 273, self.adicionarWindow)
        self.adicionarWindow.resizable(False, False)
        self.adicionarWindow.protocol("WM_DELETE_WINDOW", self.onCloseAdicionarWindow)

        backgroundLabel = tk.Label(self.adicionarWindow, image = self.adicionarWindowBackGround, bg = "white")
        backgroundLabel.place(relx = 0.5, rely = 0.5, anchor = CENTER)

        num_processo = tk.StringVar()
        num_processoEntry = ttk.Entry(self.adicionarWindow, textvariable = num_processo)
        num_processoEntry.place(relx = 0.544, rely = 0.323, anchor = CENTER, width = 253)
        
        cliente_input = tk.StringVar()
        cliente_inputEntry = ttk.Entry(self.adicionarWindow, textvariable = cliente_input)
        cliente_inputEntry.place(relx = 0.507, rely = 0.469, anchor = CENTER, width = 293)

        adversa_input = tk.StringVar()
        adversa_inputEntry = ttk.Entry(self.adicionarWindow, textvariable = adversa_input)
        adversa_inputEntry.place(relx = 0.518, rely = 0.613, anchor = CENTER, width = 284)

        cidade_input = tk.StringVar()
        cidade_inputEntry = ttk.Entry(self.adicionarWindow, textvariable = cidade_input)
        cidade_inputEntry.place(relx = 0.509, rely = 0.76, anchor = CENTER, width = 292)

        validateInput = partial(self.validarAdicaoProcesso, num_processo, cliente_input, adversa_input, cidade_input) 
        
        buttonStyle = ttk.Style()
        buttonStyle.configure("W.TButton", background = "#a3cae7", foreground = "black", font = ("Open Sans", 9))
        confirmButton = ttk.Button(self.adicionarWindow, style = "W.TButton", text = "Confirma", command = validateInput)
        confirmButton.place(relx = 0.5, rely = 0.89, anchor = CENTER, width = 60) 
            
        # add RETURN key handler
        def handler(e):
            validateInput()
        self.adicionarWindow.bind("<Return>", handler)



    def validarGeracaoDocumento(self, num_processo):
        radioButtonVariable = self.radioButtonVar.get()
        processo_atual = num_processo.get()
        
        processo_atual = self.validarNumeroProcesso(processo_atual, self.gerarWindow)

        if processo_atual is None:
            return

        try:
            index = self.numero_processos.index(processo_atual)
        except:
            messagebox.showerror(title = "Erro", message = "Processo não encontrado!")
            return 


        cliente_temp = self.clientes[index]
        adversa_temp = self.adversas[index]

        con = connectDataBase()
        cur = con.cursor()

        cur.execute(f"SELECT num_processo FROM processos WHERE num_processo = '{processo_atual}';")
        result = cur.fetchall()

        if not result:
            is_cliente = messagebox.askyesnocancel(title = f"Processo nº {processo_atual}", message = f"{cliente_temp} é seu cliente nesse processo?")
            if is_cliente:
                cliente_final = cliente_temp
                adversa_final = adversa_temp

            if not is_cliente:
                cliente_final = adversa_temp
                adversa_final = cliente_temp

            cur.execute(f"INSERT INTO processos (num_processo, cliente, parte_adversa, cidade) values ('{self.numero_processos[index]}', '{cliente_final}', '{adversa_final}', '{self.cidades[index]}');")


        con.commit()
        cur.close()
        con.close()

        documento = Documentos(index)
        if radioButtonVariable == 1:
            documento.gerarCumprSentenca()
        if radioButtonVariable == 2:
            documento.gerarPeticao()

        messagebox.showinfo(title = "Sucesso!", message = "Arquivo gerado com sucesso!")

            

    def onCloseGerarWindow(self):
        self.gerarWindow.destroy()
        self.root.wm_state("normal")


    def gerarDocumento(self):
        self.root.wm_state("iconic")

        self.gerarWindow = tk.Toplevel(self.root)
        self.gerarWindow.title("Gerador de Petição")
        self.gerarWindow.protocol("WM_DELETE_WINDOW", self.onCloseGerarWindow)

        self.centerWindow(360, 150, self.gerarWindow)
        self.gerarWindow.resizable(False, False)

        backgroundLabel = tk.Label(self.gerarWindow, image = self.gerarWindowBackGround, bg = "white")
        backgroundLabel.place(relx = 0.5, rely = 0.5, anchor = CENTER)
    
        num_processo = tk.StringVar()
        num_processoEntry = ttk.Entry(self.gerarWindow, textvariable = num_processo)
        num_processoEntry.place(relx = 0.6, rely = 0.355, anchor = CENTER, width = 200)

        self.radioButtonVar = tk.IntVar()
        self.radioButtonVar.set(2)
        radioButtonStyle = ttk.Style()
        radioButtonStyle.configure("Wild.TRadiobutton", background = "#ffffff", foreground = "black")
        cumprimentoSentencaRadioButton = ttk.Radiobutton(self.gerarWindow, variable = self.radioButtonVar, style = "Wild.TRadiobutton", value = 1)
        cumprimentoSentencaRadioButton.place(relx = 0.16, rely = 0.55, anchor = CENTER)
        peticaoRadioButton = ttk.Radiobutton(self.gerarWindow, variable = self.radioButtonVar, style = "Wild.TRadiobutton", value = 2)
        peticaoRadioButton.place(relx = 0.71, rely = 0.55, anchor = CENTER)

        validateInput = partial(self.validarGeracaoDocumento, num_processo)        

        buttonStyle = ttk.Style()
        buttonStyle.configure("W.TButton", background = "#a3cae7", foreground = "black", font = ("Open Sans", 9))
        confirm_button = ttk.Button(self.gerarWindow, style = "W.TButton", text = "Confirma", command = validateInput)
        confirm_button.place(relx = 0.5, rely = 0.79, anchor = CENTER, width = 60) 
            
        # add RETURN key handler
        def handler(e):
            validateInput()
        self.gerarWindow.bind("<Return>", handler)
        