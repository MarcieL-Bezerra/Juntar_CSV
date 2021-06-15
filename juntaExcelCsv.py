import os
import tkinter as tk
import tkinter.filedialog as fdlg
import tkinter.messagebox as tkMessageBox
from tkinter import *
from datetime import date
import pandas as pd
from tkinter import ttk
from time import sleep

#root = tk.Tk()
tinicial = tk.Tk()
tinicial.geometry("800x500+200+100")
tinicial.title("JUNTA ARQUIVOS - SIS")
tinicial.resizable(width=False, height=False)
tinicial['bg'] = '#49A'
tinicial.iconphoto(True, PhotoImage(file='./arquivos/quebra.png'))
image=PhotoImage(file='./arquivos/junta.png')

#importante para progressbar
s = ttk.Style() 
s.theme_use('default') 
s.configure("SKyBlue1.Horizontal.TProgressbar", foreground='DarkSeaGreen3', background='white')

data_atual = date.today()
data_atual = data_atual.strftime('%d-%m-%Y')

#confirma a execulsão da função

chkthalesjunta = BooleanVar()
chkthalesjunta.set(0)
chkrelthalesjunta = Checkbutton(tinicial, bd=4,text='Junta CSV',bg="SKyBlue1",font=('arial',10,'bold'),var=chkthalesjunta)
chkrelthalesjunta.grid(row=11, column=1)

chkexceljunta = BooleanVar()
chkexceljunta.set(0)
chkrelexceljunta = Checkbutton(tinicial, bd=4,text='Junta Excel',bg="SKyBlue1",font=('arial',10,'bold'),var=chkexceljunta)
chkrelexceljunta.grid(row=11, column=2)


def juntacsv():
	progress1.start(10)
	if chkthalesjunta.get() == 1:
		try:
			#aqui seleciona os arquivos
			path = fdlg.askopenfilenames()
			df = pd.DataFrame()

			for f in path:
				data=pd.read_csv(f,sep = '|') #sep = '|'
				df = df.append(data)
			tkMessageBox.showinfo("Selecionar Pasta", message= "Selecione Pasta para salvar!")

			#aqui seleciona a pasta a ser colocada o novo arquivo
			opcoes = {}                # as opções são definidas em um dicionário
			opcoes['initialdir'] = ''    # será o diretório atual
			opcoes['parent'] = tinicial
			opcoes['title'] = 'Diálogo que retorna o nome do diretório selecionado'
			caminhoinicial = fdlg.askdirectory(**opcoes)

			df.to_excel(caminhoinicial +'/Relatorio-Thalles-' + data_atual + '.xlsx', index=False)

			tkMessageBox.showinfo("Junção Finalizada", message= "Realizado com sucesso!")
			progress1.stop()
		except :
			tkMessageBox.showinfo("Erro", message= "Tente Novamente!")
			progress1.stop()
	else:
		tkMessageBox.showinfo("Não Selecionado", message= "Nenhuma ação realizada!")
		progress1.stop()


def juntaexcel():
	progress1.start(10)
	if chkexceljunta.get() == 1:
		try:
			#aqui seleciona os arquivos
			path = fdlg.askopenfilenames()
			df = pd.DataFrame()

			for f in path:
				data=pd.read_excel(f) #sep = '|'
				df = df.append(data)
			tkMessageBox.showinfo("Selecionar Pasta", message= "Selecione Pasta para salvar!")

			#aqui seleciona a pasta a ser colocada o novo arquivo
			opcoes = {}                # as opções são definidas em um dicionário
			opcoes['initialdir'] = ''    # será o diretório atual
			opcoes['parent'] = tinicial
			opcoes['title'] = 'Diálogo que retorna o nome do diretório selecionado'
			caminhoinicial = fdlg.askdirectory(**opcoes)

			df.to_excel(caminhoinicial +'/Relatorio-Excell-' + data_atual + '.xlsx', index=False)

			tkMessageBox.showinfo("Junção Finalizada", message= "Realizado com sucesso!")
			progress1.stop()
		except :
			tkMessageBox.showinfo("Erro", message= "Tente Novamente!")
			progress1.stop()
	else:
		tkMessageBox.showinfo("Não Selecionado", message= "Nenhuma ação realizada1!")
		progress1.stop()

def confirma():
	if chkthalesjunta.get() == 1:
		chkexceljunta.set(0)
		juntacsv()

	elif chkexceljunta.get() == 1:
		chkthalesjunta.set(0)
		juntaexcel()

	else:
		tkMessageBox.showinfo("Não Selecionado", message= "Nenhuma ação realizada!")
		progress1.stop()







robozinho = Label(tinicial, image = image,width=800, height=450,bg ="white")
robozinho.grid(row=10,columnspan =10)

cmdCadastrar=Button(tinicial,bd=4,bg = 'DarkSeaGreen3',fg='black',text='Selecionar os Arquivos',font=('arial',18,'bold'),width=25,height=2,
	command = confirma).grid(row=10,columnspan=10)

progress1 =ttk.Progressbar(robozinho, orient=VERTICAL, length=450, style="SKyBlue1.Horizontal.TProgressbar",mode='determinate')
progress1.place(relx=0.005, rely = 0)


tinicial.mainloop()