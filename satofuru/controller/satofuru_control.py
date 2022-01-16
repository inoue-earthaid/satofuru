import os
import sys
sys.path.append(os.path.dirname(os.path.dirname(os.path.dirname(__file__))))

from satofuru.models.satofuru_models import Satofuru

def control():
    satofuru = Satofuru()
    satofuru.login()
    satofuru.download()
    satofuru.quit()
    satofuru.get_values()
    satofuru.sp_import()