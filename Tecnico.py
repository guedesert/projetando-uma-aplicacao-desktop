from Login import Login
class Tecnico (Login):
    def __init__(self, cpf, email, senha, nome, telefone, equipe, turno):
        super().__init__(cpf, email, senha)
        self.nome = nome
        self.telefone = telefone
        self.equipe = equipe
        self.turno = turno