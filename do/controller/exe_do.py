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
    if is_none:
        do.quit()
        is_yamato_export_none = do.export_yamato_csv()
        if is_yamato_export_none:
            return
        return 
    sleep(30)
    do.quit()
    is_import_error = do.import_csv()
    if is_import_error:
        return 
        
    is_yamato_export_none = do.export_yamato_csv()
    if is_yamato_export_none:
        return
    sleep(2)

    
if __name__ == '__main__':
    exe_process()