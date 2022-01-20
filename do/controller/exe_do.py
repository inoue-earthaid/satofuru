from time import sleep
from do.models.do_models import Do


def exe_process():
    do = Do()
    do.login()
    sleep(3)
    do.search_status_request()
    do.change_request_to_receipt()
    
if __name__ == '__main__':
    exe_process()