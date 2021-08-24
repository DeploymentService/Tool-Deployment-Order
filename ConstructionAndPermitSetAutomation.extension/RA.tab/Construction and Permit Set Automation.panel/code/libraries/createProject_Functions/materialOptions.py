class MaterialOption():

    def __init__(self):
        pass

    @classmethod
    def CreateNewOptionSet(self):
        materialOption_Template = {
                                   "Walls" : {
                                               "Exterior" : None,
                                               "Interior" : None
                                             },
                                   
                                   "Floors" : {
                                               "Exterior" : None,
                                               "Interior" : None,
                                               "Balcony" : None
                                              },
   
                                   "Roofs" : {
                                               "Main" : None
                                             }
                                   }

        return materialOption_Template