# models/contato_model.py
class Contato:
    def __init__(self, nome="", numero="", status="Aguardando"):
        self.nome = nome
        self.numero = numero
        self.status = status

    def to_row(self):
        return [self.nome, self.numero, self.status]