from time import sleep
from do.do_models import Do


def execution():
    do = Do()
    do.login()
    sleep(2)
    do.search_status_request()
    not_order = do.change_request_to_receipt()
    if not_order:
        return
    do.export_csv()
    
if __name__ == '__main__':
    execution()