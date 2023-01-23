from tkinter import *
login = Tk()
login.title = "Login"
Label(login,text="Usuário").pack()
usuario = Entry()
usuario.pack()
Label(login,text="Senha:").pack()
senha = Entry()
senha.pack()
entrar = Button(login,text="Entrar",command=login.destroy)
entrar.pack()

login.mainloop()