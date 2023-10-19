#importações
import tkinter as tk
import functions

# Criar janela principal
janelaPrincipal = tk.Tk()
janelaPrincipal.title("Gerenciamento Escolar")
janelaPrincipal.geometry('300x300')

#criar Botões
#botao que adiciona novo aluno no sistema
botaoAddAluno=tk.Button(janelaPrincipal, text='Cadastrar\nNovo Aluno', command=lambda: functions.addAluno(janelaPrincipal), padx=29, pady=10, anchor='center')
botaoAddAluno.grid(column=0, row=0, padx=(10,5), pady=(10,5))

#botao que altera informações de um aluno cadastrado no sistema
botaoAlterarAluno=tk.Button(janelaPrincipal, text='Alterar informações\nde Aluno', command=lambda: functions.alteraInfo(janelaPrincipal), padx=10, pady=10, anchor='center')
botaoAlterarAluno.grid(column=1, row=0, padx=(5,10), pady=(10,5))

#botao para adicionar nota a um aluno especifico no sistema
botaoProvaIndividual=tk.Button(janelaPrincipal, text='Nota\nIndividual', command=lambda: functions.provaIndividual(janelaPrincipal), padx=36, pady=10, anchor='center')
botaoProvaIndividual.grid(column=0, row=1, padx=(10,5), pady=(5,5))

#botao para adicionar notas a todos os alunos no sistema (um de cada vez)
botaoProvaColetiva=tk.Button(janelaPrincipal, text='Nota\nColetiva', command=lambda: functions.provaColetiva(janelaPrincipal) , padx=41, pady=10,  anchor='center')
botaoProvaColetiva.grid(column=1, row=1, padx=(5,10), pady=(5,5))

#botao para ver uma lista dos alunos (apenas nome e idade)
botaoListarAlunos = tk.Button(janelaPrincipal, text="Lista\nde\nAlunos", command=lambda: functions.listarAlunos(janelaPrincipal), padx=43, pady=2,  anchor='center')
botaoListarAlunos.grid(column=0, row=2, padx=(10,5), pady=(5,5))

#botao para ver uma lista dos alunos (nome, notas e situação)
#aqui voce tambem pode criar um arquivo excel com as irformaçoes vizualizadas
botaoListarAlunosComSituacao = tk.Button(janelaPrincipal, text="Lista de Alunos\ncom Situação", command=lambda: functions.encerrarAno(janelaPrincipal), padx=23, pady=9, anchor='center')
botaoListarAlunosComSituacao.grid(column=1, row=2, padx=(5,10), pady=(5,5))

#botao para remover um aluno do sistema
botaoRemoveAluno = tk.Button(janelaPrincipal, text="Remover\nAluno", command=functions.removerAluno, padx=39, pady=10 ,anchor='center')
botaoRemoveAluno.grid(column=0, row=3, pady=(5,10), columnspan=2)

# Iniciar loop da interface gráfica
janelaPrincipal.mainloop()
