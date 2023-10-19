import json
import tkinter as tk
from tkinter import (simpledialog, messagebox)
from openpyxl import Workbook

#Função pra salvar o arquivo(recebe como parametro o nome do arquivo a ser salvo, e o arquivo alterado para salvar). Abre o arquivo com esse nome, e salva ele com o arquivo alterado.
def salvarArquivo(nomeArquivo=str, arquivoAlterado=dict):
    with open(nomeArquivo, 'w') as arquivo:
        json.dump(arquivoAlterado, arquivo, indent=4)

#Função para carregar o arquivo(recebe como parametro o nome do arquivo a abrir). abre o arquivo e retorna ele.
def carregarArquivo(nomeArquivo=str):
    with open(nomeArquivo, 'r') as arquivo:
        conteudoArquivo=json.load(arquivo)
    return conteudoArquivo

# função add aluno aos dados
def addAluno(janela):
    #Tenta abrir o arquivo, se nao existir cria um novo arquivo, salva ele e depois abre.
    try:
        conteudoArquivo=carregarArquivo('listaAlunos.txt')
    except:
        
        conteudoArquivo={'alunos': []}
        salvarArquivo('listaAlunos.txt', conteudoArquivo)

        messagebox.showerror('Arquivo','Arquivo de Dados Inexistente! Novo Arquivo Criado com Sucesso!')
        
        conteudoArquivo=carregarArquivo('listaAlunos.txt')

    #Preenche campos não obrigatorios com a mensagem "Campo não obrigatorio"
    def preencherCampo(evento):
        entrada = evento.widget
        if entrada.get() == entrada.default_text:
            entrada.delete(0, tk.END)

    def restaurarCampo(evento):
        entrada = evento.widget
        if entrada.get() == "":
            entrada.insert(0, entrada.default_text)

    def cancelar():
        janelaCadastro.destroy()

    #reliza o cadastro do novo aluno
    def cadastrar():
        if inputNome.get() != '' and inputSobrenome.get() !='' and  inputIdade.get()!='':
            aluno={'nome': inputNome.get().strip().lower().capitalize(), 'sobrenome': inputSobrenome.get().strip().lower().capitalize()}
            try:
                idade=int(inputIdade.get().strip())
                aluno['idade']=idade
            except:
                idade=False
                messagebox.showerror('Erro','Dígite Apenas Números Inteiros no Campo de Idade!')
                janelaCadastro.focus_force()
            if idade:
                if inputNota.get() != inputNota.default_text and inputNota.get() !='':
                    nota=inputNota.get().strip().replace(',','.')
                    try:
                        nota=float(nota)
                        aluno['notas']=[nota]
                        nota=True
                    except:
                        messagebox.showerror('Erro','Dígite Apenas Númenos no Campo de Nota!')
                        janelaCadastro.focus_force()
                else:
                    nota=False
                if nota == False or nota==True:
                    conteudoArquivo['alunos'].append(aluno)
                    salvarArquivo('listaAlunos.txt', conteudoArquivo)  
                    messagebox.showinfo('Arquivo', f'Aluno {aluno["nome"]} {aluno["sobrenome"]} Cadastrado Com Sucesso!')
                    janelaCadastro.destroy()
        else:
            messagebox.showerror('Erro', 'Campo de Nome, Sobrenome e Idade São Obrigatórios!')
            janelaCadastro.focus_force()

    # Cria a janela para o preenchimento do formulario do aluno a ser adicionado ao sistema       
    janelaCadastro=tk.Toplevel(janela)
    janelaCadastro.title('Cadastrar Novo Aluno')

    tk.Label(janelaCadastro,text='Nome').grid(row=0,column=0,padx=5,pady=(5,0))
    inputNome=tk.Entry(janelaCadastro, width=30)
    inputNome.grid(row=0,column=1,padx=5,pady=(5,0))
    inputNome.focus_set()

    tk.Label(janelaCadastro,text='Sobrenome').grid(row=1,column=0,padx=5)
    inputSobrenome=tk.Entry(janelaCadastro, width=30)
    inputSobrenome.grid(row=1,column=1,padx=5)

    tk.Label(janelaCadastro,text='Idade').grid(row=2,column=0,padx=5)
    inputIdade=tk.Entry(janelaCadastro, width=30)
    inputIdade.grid(row=2,column=1,padx=5)

    tk.Label(janelaCadastro,text='Média / Nota').grid(row=3,column=0,padx=5,pady=(0,5))
    inputNota=tk.Entry(janelaCadastro, width=30)
    inputNota.grid(row=3,column=1,padx=5,pady=(0,5))
    inputNota.default_text = 'Campo Não Obrigatório!' 
    inputNota.insert(0, inputNota.default_text)
    inputNota.bind("<FocusIn>", preencherCampo)
    inputNota.bind("<FocusOut>", restaurarCampo)

    coluna=tk.Label(janelaCadastro,text='',width=10)
    coluna.grid(row=0,column=2,rowspan=4)

    botaoCadastrar= tk.Button(janelaCadastro, text='Cadastrar Aluno', command=cadastrar, default="active")
    botaoCadastrar.grid(row=4,column=0,columnspan=2,pady=(0,5),padx=100)

    janelaCadastro.bind('<Return>', lambda event=None: botaoCadastrar.invoke())

    botaoCancelar= tk.Button(janelaCadastro, text='Cancelar', command=cancelar)
    botaoCancelar.grid(row=4,column=1, pady=(0,5),padx=(130,0)) 

def alteraInfo(janela):
    try:
        conteudoArquivo=carregarArquivo('listaAlunos.txt')
    except:
        messagebox.showerror('Arquivo','Arquivo de Dados Inexistente! Execulte um Novo Cadastro de Aluno!')
        return
    achou=False
    while not achou:
        alunoAlterar=simpledialog.askstring('Alteração', f'{"   "*5}Nome do Aluno a Ser Alterado{"   "*5}',parent=janela)
        if not alunoAlterar:
            break     
        else:   
            #pergunta o nome e ve se algum aluno no sistema tem esse nome
            nome=alunoAlterar.strip().lower().capitalize()
            sobrenome=False
            if ' ' in nome:
                nomeCompleto=nome.split(' ')
                nome=nomeCompleto[0]
                sobrenome=nomeCompleto[1].capitalize()
            if sobrenome:
                for n, a in enumerate(conteudoArquivo['alunos'],start=1):
                    if a['nome'] == nome and a['sobrenome'] == sobrenome:
                        if not achou:
                            achou={n: a}#quando achou algum aluno com esse nome add um dict com a key numero do aluno no sistema e seu nome de conteudo
                        else:
                            achou[n]=a#se achou mais de 1 com o mesmo nome add uma nova key ao dict
            if not sobrenome:
                #checa a quantidade de letras digitadas caso o nome esteja incompleto
                quantidadeLetrasEscritas=len(nome) 
                for n, a in enumerate(conteudoArquivo['alunos'],start=1):
                    if a['nome'][:quantidadeLetrasEscritas] == nome:
                        if not achou:
                            achou={n: a}#quando achou algum aluno com esse nome add um dict com a key numero do aluno no sistema e seu nome de conteudo
                        else:
                            achou[n]=a#se achou mais de 1 com o mesmo nome add uma nova key ao dict
            if nome and not achou:
                messagebox.showerror('Erro',f'Aluno {nome} não Cadastrado no Sistema!')
            # depois de achar. se tem mais de 1 aluno com esse nome vamos criar uma lista pra escolha visual do usuario
            if achou:
                if len(achou) == 1: # se tem so 1 aluno com esse nome continua
                    def salvarAluno():
                        alunoAlterar['nome']=inputNome.get().lower().capitalize()
                        alunoAlterar['sobrenome']=inputSobrenome.get().lower().capitalize()
                        try:
                            idade=int(inputIdade.get())
                            alunoAlterar['idade']=idade
                        except:
                            messagebox.showerror("Erro","Digíte Somente Números Inteiros no Campo de Idade!")
                            janelaAlteracao.focus_force()
                            idade=None
                        notas=inputNotas.get()
                        if notas != "":
                            try:
                                notas=notas.replace(",",".")
                                notasBack=notas.split(" ")
                                notas=[]
                                for nota in notasBack:
                                    nota=float(nota)
                                    notas.append(nota)
                                alunoAlterar['notas']=notas
                            except:
                                messagebox.showerror("Erro","Digíte Somente Números no Campo de Notas!")
                                janelaAlteracao.focus_force()
                        if notas=='':
                            notas=True
                        if idade and notas:
                            conteudoArquivo["alunos"][keyAluno]=alunoAlterar
                            salvarArquivo('listaAlunos.txt', conteudoArquivo)
                            messagebox.showinfo('Arquivo', 'Informações Alteradas com Sucesso!')
                            janelaAlteracao.destroy()
                        
                    keyAluno = list(achou.keys())[0]
                    keyAluno-=1
                    alunoAlterar=list(achou.values())[0]   #pega o dados do aluno
        
                    janelaAlteracao=tk.Toplevel(janela)
                    janelaAlteracao.focus_force()
                    janelaAlteracao.title(f'Alteração do Aluno {alunoAlterar["nome"]} {alunoAlterar["sobrenome"]}')
        
                    tk.Label(janelaAlteracao,text='Nome').grid(row=0,column=0, sticky='e')
                    inputNome=tk.Entry(janelaAlteracao)
                    inputNome.grid(row=0,column=1)
                    inputNome.insert(0, alunoAlterar['nome'])
                    inputNome.bind("<FocusOut>", lambda event: inputNome.insert(0, alunoAlterar['nome']) if not inputNome.get() else None)

                    tk.Label(janelaAlteracao,text='Sobrenome').grid(row=1,column=0, sticky='e')
                    inputSobrenome=tk.Entry(janelaAlteracao)
                    inputSobrenome.grid(row=1,column=1)
                    inputSobrenome.insert(0, alunoAlterar['sobrenome'])
                    inputSobrenome.bind("<FocusOut>", lambda event: inputSobrenome.insert(0, alunoAlterar['sobrenome']) if not inputSobrenome.get() else None)

                    tk.Label(janelaAlteracao,text='Idade').grid(row=2,column=0, sticky='e')
                    inputIdade=tk.Entry(janelaAlteracao)
                    inputIdade.grid(row=2,column=1)
                    inputIdade.insert(0, alunoAlterar['idade'])
                    inputIdade.bind("<FocusOut>", lambda event: inputIdade.insert(0, alunoAlterar['idade']) if not inputIdade.get() else None)

                    tk.Label(janelaAlteracao,text='Notas').grid(row=3,column=0, sticky='e')
                    inputNotas=tk.Entry(janelaAlteracao)
                    inputNotas.grid(row=3,column=1)
                    if "notas" in alunoAlterar:
                        inputNotas.insert(0, alunoAlterar['notas'])
                        inputNotas.bind("<FocusOut>", lambda event: inputNotas.insert(0, alunoAlterar['notas']) if not inputNotas.get() else None)

                    label=tk.Label(janelaAlteracao,text="Atenção Com o Campo Notas. As Notas Devem Ser Separadas Por Espaços E Devem Ser Apenas Números!")
                    label.grid(row=4, column=0, columnspan=4)

                    botaoSalvar=tk.Button(janelaAlteracao, text='Salvar', command=salvarAluno, padx=17, anchor='center', default="active")
                    botaoSalvar.grid(row=3, column=2, sticky='e')

                    janelaAlteracao.bind('<Return>', lambda event=None: botaoSalvar.invoke())

                elif len(achou) > 1:    # mais de 1 aluno com esse nome
                #função que vai remover o aluno baseado na seleção do aluno em uma lista com os alunos com o mesmo nome

                    def alterarAlunoSelecionado():
                        def salvarAluno():
                            alunoAlterar['nome']=inputNome.get()
                            alunoAlterar['sobrenome']=inputSobrenome.get()
                            try:
                                idade=int(inputIdade.get())
                                alunoAlterar['idade']=idade
                            except:
                                messagebox.showerror("Erro","Digíte Somente Números Inteiros no Campo de Idade!")
                                janelaAlteracao.focus_force()
                                idade=None
                            notas=inputNotas.get()
                            if notas != "":
                                try:
                                    notas=notas.replace(",",".")
                                    notasBack=notas.split(" ")
                                    notas=[]
                                    for nota in notasBack:
                                        nota=float(nota)
                                        notas.append(nota)
                                    alunoAlterar['notas']=notas
                                except:
                                    messagebox.showerror("Erro","Digíte Somente Números no Campo de Notas!")
                                    janelaAlteracao.focus_force()
                            if notas=='':
                                notas=True
                            if idade and notas:
                                conteudoArquivo["alunos"][keyAluno]=alunoAlterar
                                salvarArquivo('listaAlunos.txt', conteudoArquivo)
                                messagebox.showinfo('Arquivo', 'Informações Alteradas com Sucesso!')
                                janelaSelecionarAluno.destroy()# fecha a janela da lista
                                janelaAlteracao.destroy()
                
                        alunoSelecionado=lista.curselection() #pega o aluno clicado na lista
                        if alunoSelecionado:
                            indiceLista=alunoSelecionado[0] #posição do aluno na lista
                            if indiceLista!=lista.size()-1 and indiceLista!=0: #ve se a posição clicada na lista nao é a ultima nem a primeira pois o ultimo item dessa lista é uma frase informativa e o primeiro sao as informações
                                keyAluno=list(achou.keys())[indiceLista-1]    #pega o numero do aluno no sistema! que tambem é a key dele no dict 'achou'
                                alunoAlterar=list(achou.values())[indiceLista-1]   #pega o conteudo
                                keyAluno-=1

                                janelaAlteracao=tk.Toplevel(janela)
                                janelaAlteracao.focus_force()
                                janelaAlteracao.title(f'Alteração do Aluno {alunoAlterar["nome"]} {alunoAlterar["sobrenome"]}')
                    
                                tk.Label(janelaAlteracao,text='Nome').grid(row=0,column=0, sticky='e')
                                inputNome=tk.Entry(janelaAlteracao)
                                inputNome.grid(row=0,column=1)
                                inputNome.insert(0, alunoAlterar['nome'])
                                inputNome.bind("<FocusOut>", lambda event: inputNome.insert(0, alunoAlterar['nome']) if not inputNome.get() else None)

                                tk.Label(janelaAlteracao,text='Sobrenome').grid(row=1,column=0, sticky='e')
                                inputSobrenome=tk.Entry(janelaAlteracao)
                                inputSobrenome.grid(row=1,column=1)
                                inputSobrenome.insert(0, alunoAlterar['sobrenome'])
                                inputSobrenome.bind("<FocusOut>", lambda event: inputSobrenome.insert(0, alunoAlterar['sobrenome']) if not inputSobrenome.get() else None)

                                tk.Label(janelaAlteracao,text='Idade').grid(row=2,column=0, sticky='e')
                                inputIdade=tk.Entry(janelaAlteracao)
                                inputIdade.grid(row=2,column=1)
                                inputIdade.insert(0, alunoAlterar['idade'])
                                inputIdade.bind("<FocusOut>", lambda event: inputIdade.insert(0, alunoAlterar['idade']) if not inputIdade.get() else None)

                                tk.Label(janelaAlteracao,text='Notas').grid(row=3,column=0, sticky='e')
                                inputNotas=tk.Entry(janelaAlteracao)
                                inputNotas.grid(row=3,column=1)
                                if "notas" in alunoAlterar:
                                    inputNotas.insert(0, alunoAlterar['notas'])
                                    inputNotas.bind("<FocusOut>", lambda event: inputNotas.insert(0, alunoAlterar['notas']) if not inputNotas.get() else None)

                                label=tk.Label(janelaAlteracao,text="Atenção Com o Campo Notas. As Notas Devem Ser Separadas Por Espaços E Devem Ser Apenas Números!")
                                label.grid(row=4, column=0, columnspan=4)

                                botaoSalvar=tk.Button(janelaAlteracao, text='Salvar', command=salvarAluno, padx=17, anchor='center', default="active")
                                botaoSalvar.grid(row=3, column=2, sticky='e')

                                janelaAlteracao.bind('<Return>', lambda event=None: botaoSalvar.invoke())

                    messagebox.showwarning('Alunos',f'{len(achou)} Alunos Com o Nome {alunoAlterar} Localizados no Sistema!')
                    janelaSelecionarAluno=tk.Toplevel(janela)  #cria a janela para a lista de alunos com o mesmo nome
                    janelaSelecionarAluno.title(f'Lista de Alunos {alunoAlterar}')
                    fontMonoespaco=("Consolas", 10)
                    lista = tk.Listbox(janelaSelecionarAluno, height=15, width=61, font=fontMonoespaco)#cria a lista 
                    lista.grid(column=0, row=0, padx=10, pady=10)
                    #add os itens na lista
                    lista.insert(tk.END, f'{"Num":^5}{"Nome":<50}{"Idade":^5}') #add os alunos na lista
                    for n, a in achou.items():
                        lista.insert(tk.END, f'''{n:^5}{f'{a["nome"]} {a["sobrenome"]}':<50}{a["idade"]:^5}''')  
                    lista.insert(tk.END, '>>> Num Referente a Chegada do Aluno no Sistema!')
                    #botao
                    keyAluno=None
                    alunoAlterar=None
                    botaoSelecionar = tk.Button(janelaSelecionarAluno, text='Selecionar Aluno', command=alterarAlunoSelecionado)
                    botaoSelecionar.grid(column=0, row=1, padx=10, pady=10)  
                    lista.bind("<Double-Button-1>", lambda event=None: botaoSelecionar.invoke())

#função para adicionar nota para somente 1 aluno no sistema
def provaIndividual(janela):
    try:
        #abre o arquivo .txt com os dados na variavel conteudoArquivo
        conteudoArquivo=carregarArquivo('listaAlunos.txt')
        achou=False
        #pergunta o nome e ve se algum aluno no sistema tem esse nome
        while not achou:
            nome=simpledialog.askstring('Prova', f'{"   "*5}Nome do Aluno{"   "*5}',parent=janela)
            if not nome:
                break     
            else:  
                nome=nome.strip().lower().capitalize()
                sobrenome=False
                for l in nome:
                    if l ==' ':
                        nomeCompleto=nome.split(' ')
                        nome=nomeCompleto[0]
                        sobrenome=nomeCompleto[1].capitalize()
                if sobrenome:
                    for n, a in enumerate(conteudoArquivo['alunos'],start=1):
                        if a['nome'] == nome and a['sobrenome'] == sobrenome:
                            if not achou:
                                #quando achou algum aluno com esse nome add um dict com a key numero do aluno no sistema e seu nome de conteudo
                                achou={n: a}
                            else:
                                #se achou mais de 1 com o mesmo nome add uma nova key ao dict
                                achou[n]=a
                if not sobrenome:
                    #checa a quantidade de letras digitadas caso o nome esteja incompleto
                    quantidadeLetrasEscritas=len(nome)  
                    for n, a in enumerate(conteudoArquivo['alunos'],start=1):
                        if len(a['nome'])>=quantidadeLetrasEscritas:
                            if a['nome'][:quantidadeLetrasEscritas] == nome:
                                if not achou:
                                    #quando achou algum aluno com esse nome add um dict com a key numero do aluno no sistema e seu nome de conteudo
                                    achou={n: a}
                                else:
                                    #se achou mais de 1 com o mesmo nome add uma nova key ao dict
                                    achou[n]=a
                if nome and not achou:
                    messagebox.showerror('Erro',f'Aluno {nome} não Cadastrado no Sistema!')
        # depois de achar. se tem mais de 1 aluno com esse nome vamos criar uma lista pra escolha visual do usuario
        if achou:
            # se tem so 1 aluno com esse nome continua
            if len(achou) == 1: 
                #pega o numero referente a posição do aluno na list em dados
                achou=list(achou.values())[0]  
                while True:
                    nota=simpledialog.askstring(f'Aluno {achou["nome"]} {achou["sobrenome"]}', f'{"   "*5}Nota Para o Aluno {achou["nome"]}{"   "*5}',parent=janela)
                    if not nota:
                        break
                    else:
                        try:
                            nota=nota.replace(',','.')
                            nota=float(nota)
                            try:
                                #add a nota no sistema
                                achou['notas'].append(nota)
                            except:
                                achou['notas']=[nota]
                            messagebox.showinfo('Arquivo', f'Nota Adicionada ao Aluno {achou["nome"]} {achou["sobrenome"]}!')
                            salvarArquivo('listaAlunos.txt', conteudoArquivo)
                            break
                        except:
                            messagebox.showerror('Erro','Apenas Números no Campo de Notas!')
            # mais de 1 aluno com esse nome
            elif len(achou) > 1:    
                #função que vai remover o aluno baseado na seleção do aluno em uma lista com os alunos com o mesmo nome
                def notaAlunoSelecionado():
                    #pega o aluno clicado na lista
                    alunoSelecionado=lista.curselection() 
                    if alunoSelecionado:
                        #posição do aluno na lista
                        indiceLista=alunoSelecionado[0] 
                        #ve se a posição clicada na lista nao é a ultima pois o ultimo item dessa lista é uma frase informativa
                        if indiceLista!=lista.size()-1 and indiceLista!=0: 
                            #pega o numero do aluno no sistema! que tambem é a key dele no dict 'achou'
                            keyAluno=list(achou.keys())[indiceLista-1]   
                            keyAluno-=1
                            #pega o conteudo
                            aluno=list(achou.values())[indiceLista-1]
                            # sabemmos o aluno pega a nota e add igual antes   
                            while True:
                                nota=simpledialog.askstring(f'Aluno {aluno["nome"]} {aluno["sobrenome"]}', f'{"   "*5}Nota Para o Aluno {aluno["nome"]}{"   "*5}',parent=janelaSelecionarAluno)
                                if nota:
                                    nota=nota.replace(',','.')
                                    try:
                                        nota=float(nota)
                                        try:
                                            aluno['notas'].append(nota)
                                        except:
                                            aluno['notas']=[nota]
                                        break
                                    except:
                                        messagebox.showerror('Erro','Apenas Números no Campo de Notas!')
                                else:
                                    break
                            if nota:
                                messagebox.showinfo('Arquivo', f'Nota Adicionada ao Aluno {aluno["nome"]} {aluno["sobrenome"]}!')
                                salvarArquivo('listaAlunos.txt', conteudoArquivo)
                                # fecha a janela da lista
                                janelaSelecionarAluno.destroy()

                messagebox.showwarning('Arquivo',f'{len(achou)} Alunos Com o Nome {nome} Localizados no Sistema!')

                #cria a janela para a lista de alunos com o mesmo nome
                janelaSelecionarAluno=tk.Toplevel(janela)  
                janelaSelecionarAluno.title(f'Lista de Alunos {nome}')

                fontMonoespaco=("Consolas", 10)
                #cria a lista 
                lista = tk.Listbox(janelaSelecionarAluno, height=15, width=61, font=fontMonoespaco)
                lista.grid(column=0, row=0, padx=10, pady=10)

                #add os itens na lista
                lista.insert(tk.END, f'{"Num":^5}{"Nome":<50}{"Idade":^5}') 
                #add os alunos na lista
                for n, a in achou.items():
                    lista.insert(tk.END, f'''{n:^5}{f'{a["nome"]} {a["sobrenome"]}':<50}{a["idade"]:^5}''')
                lista.insert(tk.END, '>>> Num Referente a Chegada do Aluno no Sistema!')
                
                #botao
                botaoSelecionar = tk.Button(janelaSelecionarAluno, text='Selecionar Aluno', command=notaAlunoSelecionado)
                botaoSelecionar.grid(column=0, row=1, padx=10, pady=10)
                lista.bind("<Double-Button-1>", lambda event=None: botaoSelecionar.invoke())
    except:
        messagebox.showerror('Arquivo','Arquivo de Dados Inexistente! Execulte um Novo Cadastro de Aluno!')

#função para adicionar notas a todos os alunos do sistema
def provaColetiva(janela):
        #abre o arquivo .txt com os dados na variavel conteudoArquivo
    try:
        conteudoArquivo=carregarArquivo('listaAlunos.txt')
        def obter_nome(aluno):
            return (aluno['nome'],aluno['sobrenome'])
        alunosOrdenados=sorted(conteudoArquivo['alunos'],key=obter_nome)

    except:
        messagebox.showerror('Arquivo','Arquivo de Dados Inexistente! Execulte um Novo Cadastro de Aluno!')
        return

    def checkNotas():
        try:
            for entry in lista_entries:
                nota = entry.get()
                if nota !="":
                    nota=float(nota)
            salvarNotas()
        except:
            messagebox.showerror('Erro','Apenas Números nos Campos de Notas')
            janelaProva.focus_force()

    def salvarNotas():
        for i, entry in enumerate(lista_entries):
            nota = entry.get()
            if nota !="":
                nota=float(nota)
                for n, aluno in enumerate(alunosOrdenados):
                    if n == i:
                        aluno["notas"].append(nota)
        alunosOrdenadosSalvar={"alunos": alunosOrdenados}
        messagebox.showinfo('Arquivo', f'Notas Adicionadas com Sucesso!')
        salvarArquivo("listaAlunos.txt", alunosOrdenadosSalvar)
        janelaProva.destroy()

    janelaProva = tk.Toplevel(janela)
    janelaProva.title("Prova Coletiva")

    # Crie uma lista para armazenar as variáveis dos Entry
    lista_entries = []

    for i, aluno in enumerate(alunosOrdenados, start=1):
        tk.Label(janelaProva, text=f'Nota para {aluno["nome"]} {aluno["sobrenome"]}').grid(row=i, column=0, sticky='w', padx=(10,0))
        entry_nota = tk.Entry(janelaProva)
        if i == 1:
            entry_nota.focus_set()
        entry_nota.grid(row=i, column=1, padx=(0,10))
        lista_entries.append(entry_nota)

        # Crie um botão para salvar as notas
        botao_salvar = tk.Button(janelaProva, text="Salvar", command=checkNotas)
        botao_salvar.grid(row=len(alunosOrdenados) + 1, columnspan=2, pady=10)

    janelaProva.bind('<Return>', lambda event=None: botao_salvar.invoke())

# Listar alunos dos sistema
def listarAlunos(janela):
    def obter_nome(aluno):
        return (aluno['nome'],aluno['sobrenome'])
    try:
        conteudoArquivo=carregarArquivo('listaAlunos.txt')
        alunosOrdenados=sorted(conteudoArquivo['alunos'],key=obter_nome)

        #cria uma janela para a lista
        janelaAlunos=tk.Toplevel(janela)    
        janelaAlunos.title('Lista de Alunos')
        
        #font monoespaçada
        fontMonoespaco=("Consolas", 10) 
        #cria a lista
        lista = tk.Listbox(janelaAlunos, height=15, width=61, font=fontMonoespaco)  
        lista.grid(column=0, row=0, padx=10, pady=10)
        lista.insert(tk.END, f'{"Nome":<15}{"Sobrenome":<25}{"Idade":^5}') 
        for aluno in alunosOrdenados:
            #add os alunos na lista
            lista.insert(tk.END, f'{aluno["nome"]:<15}{aluno["sobrenome"]:<25}{aluno["idade"]:^5}')
        
        altura_lista = len(alunosOrdenados)
        lista.config(width=45, height=int(altura_lista))
    except:
        messagebox.showerror('Arquivo','Arquivo de Dados Inexistente! Execulte um Novo Cadastro de Aluno!')

#listar alunos no sistema com suas notas e situação. tambem pode se criar um arquivo excel com as mesmas informações
def encerrarAno(janela):
    def obter_nome(aluno):
        return (aluno['nome'],aluno['sobrenome'])

    def gerarExcel(janela):
        # Crie um novo arquivo Excel
        workbook = Workbook()
        sheet = workbook.active
    
        # Adicione cabeçalhos
        sheet['A1'] = 'Nome'
        sheet['B1'] = 'Sobrenome'
        sheet['C1'] = 'Notas'
        sheet['D1'] = 'Situação'
    
        # Preencha os dados
        row = 2  # Iniciar na segunda linha
        for aluno in alunosOrdenados:
            nome = aluno["nome"]
            sobrenome = aluno["sobrenome"]
            if "notas" in aluno:
                notas=[]
                for nota in aluno["notas"]:
                    nota=str(nota).replace(".",",")
                    notas.append(nota)
                notas = ' '.join(notas)
                quantidadeNotas = len(aluno["notas"])
                soma = sum(aluno["notas"])
                media = soma / quantidadeNotas
                if media >= 7:
                    situacao = "aprovado"
                else:
                    situacao = "reprovado"
            else:
                notas = "Sem notas"
                situacao = "Sem notas"
    
            sheet[f'A{row}'] = nome
            sheet[f'B{row}'] = sobrenome
            sheet[f'C{row}'] = notas
            sheet[f'D{row}'] = situacao
            row += 1
    
        # Salve o arquivo Excel
        workbook.save('dados_alunos.xlsx')
    
        messagebox.showinfo("Arquivo", "Arquivo Excel Criado Com Sucesso!")

        janela.focus_force()


    try:
        conteudoArquivo=carregarArquivo('listaAlunos.txt')
        alunosOrdenados=sorted(conteudoArquivo['alunos'],key=obter_nome)
        
        #cria uma janela para a lista
        janelaAlunos=tk.Toplevel(janela)    
        janelaAlunos.title('Lista de Alunos')
        
        for n, aluno in enumerate(alunosOrdenados,start=1):
            if "notas" in aluno:
                soma=0
                for nota in aluno["notas"]:
                    soma += nota
                soma=soma/len(aluno["notas"])
                if soma > 7:
                    situacao="Aprovado"
                else:
                    situacao="Reprovado"
            else:
                situacao="Sem Notas"
            if n == 1:
                #font monoespaçada
                fontMonoespaco=("Consolas", 10) 
                #cria a lista
                lista = tk.Listbox(janelaAlunos, height=15, width=61, font=fontMonoespaco)  
                lista.grid(column=0, row=0, padx=10, pady=10)
                lista.insert(tk.END, f'{"Nome":<15}{"Sobrenome":<15}{"Notas":^20}{"Situação":>10}')
                if "notas" in aluno:
                    soma=0
                    notas=[]
                    for nota in aluno["notas"]:
                        soma += nota
                        nota=str(nota).replace(".",",")
                        notas.append(nota)
                    soma=soma/len(aluno["notas"])
                    notas = ' '.join(notas)
                    if soma > 7:
                        situacao="Aprovado"
                    else:
                        situacao="Reprovado"
                    #add os alunos na lista
                    lista.insert(tk.END, f'{aluno["nome"]:<15}{aluno["sobrenome"]:<15}{notas:^20}{situacao:>10}')
                else:
                    notas = "Sem notas"
                    situacao = "Sem notas"
                    lista.insert(tk.END, f'{aluno["nome"]:<15}{aluno["sobrenome"]:<15}{notas:^20}{situacao:>10}')
            else:
                if "notas" in aluno:
                    soma=0
                    notas=[]
                    for nota in aluno["notas"]:
                        soma += nota
                        nota=str(nota).replace(".",",")
                        notas.append(nota)
                    soma=soma/len(aluno["notas"])
                    if len(aluno["notas"]) <= 4:
                        notas = ' '.join(notas)
                    else:    
                        notas = ' '.join(notas[:4])
                        notas = notas + " ..."
                    if soma > 7:
                        situacao="Aprovado"
                    else:
                        situacao="Reprovado"
                    
                    #add os alunos na lista
                    lista.insert(tk.END, f'{aluno["nome"]:<15}{aluno["sobrenome"]:<15}{notas:^20}{situacao:>10}')
                else:
                    notas = "Sem notas"
                    situacao = "Sem notas"
                    lista.insert(tk.END, f'{aluno["nome"]:<15}{aluno["sobrenome"]:<15}{notas:^20}{situacao:>10}')
        gerarArquivo = tk.Button(janelaAlunos, text="Arquivo Excel", command=lambda: gerarExcel(janelaAlunos), padx=62, anchor='center')
        gerarArquivo.grid(column=0, row=1, padx=10, pady=10)


    except:
        messagebox.showerror('Arquivo','Arquivo de Dados Inexistente! Execulte um Novo Cadastro de Aluno!')

# Função para remover Aluno
def removerAluno():
    #abre o arquivo .txt com os dados na variavel conteudoArquivo
    try:
        conteudoArquivo=carregarArquivo('listaAlunos.txt')
    except:
        messagebox.showerror('Arquivo','Arquivo de Dados Inexistente! Execulte um Novo Cadastro de Aluno!')
        return
    achou=False
    while not achou:
        aluno=simpledialog.askstring('Alteração', f'{"   "*5}Nome do Aluno a Ser Removido{"   "*5}')
        if not aluno:
            break     
        else:   #pergunta o nome e ve se algum aluno no sistema tem esse nome
            nome=aluno.strip().lower().capitalize()
            sobrenome=False
            if ' ' in nome:
                nomeCompleto=nome.split(' ')
                nome=nomeCompleto[0]
                sobrenome=nomeCompleto[1].capitalize()
            if sobrenome:
                for n, a in enumerate(conteudoArquivo['alunos'],start=1):
                    if a['nome'] == nome and a['sobrenome'] == sobrenome:
                        if not achou:
                            achou={n: a}#quando achou algum aluno com esse nome add um dict com a key numero do aluno no sistema e seu nome de conteudo
                        else:
                            achou[n]=a#se achou mais de 1 com o mesmo nome add uma nova key ao dict
            if not sobrenome:
                quantidadeLetrasEscritas=len(nome)  #otima correção! checa a quantidade de letras digitadas caso o nome esteja incompleto
                for n, a in enumerate(conteudoArquivo['alunos'],start=1):
                    if a['nome'][:quantidadeLetrasEscritas] == nome:
                        if not achou:
                            achou={n: a}#quando achou algum aluno com esse nome add um dict com a key numero do aluno no sistema e seu nome de conteudo
                        else:
                            achou[n]=a#se achou mais de 1 com o mesmo nome add uma nova key ao dict
            if nome and not achou:
                messagebox.showerror('Erro',f'Aluno {nome} não Cadastrado no Sistema!')
            # depois de achar. se tem mais de 1 aluno com esse nome vamos criar uma lista pra escolha visual do usuario
            if achou:
                if len(achou) == 1: # se tem so 1 aluno com esee nome continua
                    num=list(achou.keys())[0]
                    conteudoArquivo['alunos'].pop(num-1)
                    messagebox.showinfo('Arquivo', 'Aluno Removido do Sistema!')
                elif len(achou) > 1:    # mais de 1 aluno com esse nome
                    #função que vai remover o aluno baseado na seleção do aluno em uma lista com os alunos com o mesmo nome
                    def removerAlunoSelecionado():
                        alunoSelecionado=lista.curselection() #pega o aluno clicado na lista
                        if alunoSelecionado:
                            indiceLista=alunoSelecionado[0] #posição do aluno na lista
                            if indiceLista!=lista.size()-1 and indiceLista!=0: #ve se a posição clicada na lista nao é a ultima pois o ultimo item dessa lista é uma frase informativa
                                keyAluno=list(achou.keys())[indiceLista]    #pega o numero do aluno no sistema 
                                lista.delete(indiceLista)   #deleta da lista
                                messagebox.showinfo('Arquivo', 'Aluno Removido do Sistema!')
                                keyAluno-=2
                                # remove o aluno do sistema, salva, e fecha a janela da lista
                                conteudoArquivo['alunos'].pop(keyAluno)
                                salvarArquivo('listaAlunos.txt', conteudoArquivo)
                                janelaRemoverAluno.destroy()
                    
                    messagebox.showwarning('Arquivo',f'{len(achou)} Alunos Com o Nome {aluno.capitalize()} Localizados no Sistema!')
                    janelaRemoverAluno=tk.Tk()  #cria a janela para a lista de alunos com o mesmo nome
                    janelaRemoverAluno.title(f'Lista de Alunos {aluno.capitalize()}')
                    fontMonoespaco=("Consolas", 10)
                    lista = tk.Listbox(janelaRemoverAluno, height=15, width=50, font=fontMonoespaco)#cria a lista 
                    lista.grid(column=0, row=0, padx=10, pady=10)
                    #add os itens na lista
                    lista.insert(tk.END, f'{"Num":^5}{"Nome":<25}{"Idade":^5}')
                    for n, a in achou.items():
                        lista.insert(tk.END, f'{n:^5}{a["nome"].capitalize():<25}{a["idade"]:^5}')
                    lista.insert(tk.END, '>>> Num Referente a Chegada do Aluno no Sistema!')
                    #botao
                    botaoRemover = tk.Button(janelaRemoverAluno, text='Remover Selecionado', command=removerAlunoSelecionado)
                    botaoRemover.grid(column=0, row=1, padx=10, pady=10)  
    salvarArquivo('listaAlunos.txt', conteudoArquivo)
