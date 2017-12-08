#  Author: BornSmasher
#  Date: 12/04/17
#
#  Customer Class
#      def: RevelationPet Customer Object
###############################################################################


from Pet import Pet


class Customer:
    def __init__(self, client, pets):
        self.Address_line_1 = (self.get_address_line_1(client), 7)
        self.Address_line_2 = (self.get_address_line_2(client), 8)
        self.City = (self.get_city(client), 9)
        self.County = (self.get_county(client), 10)
        self.Custid = (self.get_custid(client), 1)
        self.Email = (self.get_email(client), 3)
        self.Emergency_Contact = (self.get_emergency_contact(client), 12)
        self.Emergency_Phone = (self.get_emergency_phone(client), 13)
        self.Mobile_Phone = (self.get_mobile_phone(client), 5)
        self.Name = (self.get_name(client), 2)
        self.Notes = (self.get_notes(client), 6)
        self.Postcode_Zip = (self.get_postcode_zip(client), 11)
        self.Telephone = (self.get_telephone(client), 4)
        self.Total_Amount_Spent = (self.get_total_amount_spent(client), 29)
        self.Pets = self.get_pets(pets)


    def get_address_line_1(self, client):
        return client["Address"]

    def get_address_line_2(self, client):
        return client["Address2"]

    def get_city(self, client):
        return client["City"]

    def get_county(self, client):
        return None

    def get_custid(self, client):
        return client["ClientID"]

    def get_email(self, client):
        return client["E-mail"]

    def get_emergency_contact(self, client):
        return client["EmergencyContactName"]

    def get_emergency_phone(self, client):
        return client["EmergencyContactNumber"]

    def get_name(self, client):
        return str(client["FirstName"]) + str(client["LastName"])

    def get_notes(self, client):
        return client["Comment"]

    def get_postcode_zip(self, client):
        return ""

    def get_telephone(self, client):
        return client["PrimaryPhoneNumber"]

    def get_mobile_phone(self, client):
        return client["CellPhone"]

    def get_total_amount_spent(self, client):
        return 0

    # Finds Pets Associated With a Customers Account
    def get_pets(self, pets):
        pet_objects = []
        for pet in pets:
            pet_objects.append(Pet(breed=pet["Breed"],
                                   sex=pet["Gender"],
                                   type=pet["Type"],
                                   vet=pet["Vet"],
                                   weight=pet["Weight"],
                                   dead=pet["Dead"],
                                   name=pet["Name"],
                                   vaccinated="",
                                   age="",
                                   dietary="",
                                   medication="",
                                   microchip="",
                                   registration=""))
        return pet_objects

    # Adds a Customer to an Excel Worksheet in RevelationPet Format
    def print2ws(self, ws, row):
        members = [attr for attr in dir(self) if not callable(getattr(self, attr)) and not attr.startswith("__")]
        for member in members:
            if member == "Pets":
                pass
            else:
                col = self.__dict__[member][1]
                val = self.__dict__[member][0]
                ws.cell(row=row, column=col).value = val
