#  Author: BornSmasher
#  Date: 12/04/17
#
#  Class: Pet
#      def: RevelationPet Object
#
###############################################################################


class Pet:

    def __init__(self, breed=None, sex=None, type=None, vet=None, weight=None, spayed_neutered=None,
                 dead=None, name=None, vaccinated=None, bordetella=None, dhpp=None,
                 fvrcp=None, flea_and_tick=None, parvo=None, rabies=None, leptospirosis=None,
                 age=None, dietary=None, medication=None, microchip=None, registration=None):
        self.breed = (breed, 17)
        self.sex = (sex, 18)
        self.type = (type, 15)
        self.vet = (vet, 35)
        self.weight = (weight, 23)
        self.dead = (dead, 16)
        self.name = (name, 14)
        self.vaccinated = (vaccinated, 24)
        self.bordetella = (bordetella, 25)
        self.dhpp = (dhpp, 26)
        self.fvrcp = (fvrcp, 27)
        self.flea_and_tick = (flea_and_tick, 28)
        self.parvo = (parvo, 29)
        self.rabies = (rabies, 30)
        self.leptospirosis = (leptospirosis, 31)
        self.age = (age, 19)
        self.dietary = (dietary, 32)
        self.medication = (medication, 33)
        self.microchip = (microchip, 21)
        self.registration = (registration, 20)
        self.spayed_neutered = (spayed_neutered, 22)

    def print2ws(self, ws, row):
        members = [attr for attr in dir(self) if not callable(getattr(self, attr)) and not attr.startswith("__")]
        for member in members:
            col = self.__dict__[member][1]
            val = self.__dict__[member][0]
            ws.cell(row=row, column=col).value = val