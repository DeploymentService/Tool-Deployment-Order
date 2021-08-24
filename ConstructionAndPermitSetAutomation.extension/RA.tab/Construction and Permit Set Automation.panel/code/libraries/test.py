test = {"Disciplines" : {"Architectural" : "ViewDiscipline.Architectural", 
                             "Structural" : "ViewDiscipline.Structural", 
                             "Mechanical" : "ViewDiscipline.Mechanical", 
                             "Electrical" : "ViewDiscipline.Electrical", 
                             "Plumbing" : "ViewDiscipline.Plumbing", 
                             "Coordination" : "ViewDiscipline.Coordination"},

            "View Uses" : ["Permit", "Construction", "Presentation", "Working View"],
            "View Categories" : ["Building", "Furniture/Casework"]}

for disciplineName, discipline in test["Disciplines"].items():

    print(f"{disciplineName} is the kind of {discipline}")