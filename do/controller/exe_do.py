from time import sleep
from do.models.do_models import Do


def exe_process():
    do = Do()
    do.login()
    sleep(1)
    do.search_status_request()
    sleep(1)
    do.change_request_to_receipt()
    sleep(1)
    do.filtering_export_csv()
    sleep(1)
    is_none = do.export_csv()
    sleep(1)
    do.quit()
    if is_none:
        is_yamato_export_none = do.export_yamato_csv()
        if is_yamato_export_none:
            return
        return 
    do.import_csv()
    is_yamato_export_none = do.export_yamato_csv()
    if is_yamato_export_none:
        return
    sleep(2)

    
if __name__ == '__main__':
    exe_process()