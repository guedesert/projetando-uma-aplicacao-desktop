from Ferramenta import Ferramenta
from Tecnico import Tecnico
class Reserva:
    def __init__(self, id, ferramenta, tecnico):
        self.id = id
        self.ferramenta = [ferramenta]
        self.tecnico = tecnico