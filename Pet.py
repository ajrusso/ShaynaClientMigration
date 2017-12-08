#  Author: BornSmasher
#  Date: 12/04/17
#
#  Class: Pet
#      def: RevelationPet Object
#
###############################################################################


class Pet:

    def __init__(self, breed=None, sex=None, type=None, vet=None, weight=None,
                       dead=None, name=None, vaccinated=None, age=None,
                       dietary=None, medication=None, microchip=None,
                       registration=None):
        self.breed = (breed, 17)
        self.sex = (sex, 18)
        self.type = (type, 15)
        self.vet = (vet, 28)
        self.weight = (weight, 23)
        self.dead = (dead, 16)
        self.name = (name, 14)
        self.vaccinated = (vaccinated, 24)
        self.age = (age, 19)
        self.dietary = (dietary, 25)
        self.medication = (medication, 26)
        self.microchip = (microchip, 21)
        self.registartion = (registration, 20)

    def print2ws(self, ws, row):
        members = [attr for attr in dir(self) if not callable(getattr(self, attr)) and not attr.startswith("__")]
        for member in members:
            col = self.__dict__[member][1]
            val = self.__dict__[member][0]
            ws.cell(row=row, column=col).value = val