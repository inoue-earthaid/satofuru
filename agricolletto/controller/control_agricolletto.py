from agricolletto.models.agricollet_models import Agricolletto

def agricolletto():
    agri = Agricolletto()
    agri.convert_pdf()
    agri.shaping_pdf()
    agri.reflect_spreadsheet()