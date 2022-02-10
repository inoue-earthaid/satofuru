from time import sleep
from do.models.do_models import Do


def exe_process():
    do = Do()
    # do.login()
    # sleep(3)
    # do.filtering_export_csv()
    # sleep(3)
    # do.export_csv()
    # sleep(3)
    do.quit()
    do.import_csv()
    yamato_import_csv = do.export_yamato_csv()

    
if __name__ == '__main__':
    exe_process()