from agricolletto.models.tes import Agricolletto

def agricolletto():
    agri = Agricolletto()
    agri.convert_pdf()
    result = agri.shaping_pdf()
    print(result)