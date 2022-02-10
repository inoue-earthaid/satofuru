import os
from agricolletto.controller import control_agricolletto

def main():
    control_agricolletto.agricolletto()

if __name__ == '__main__':
    main()
    os.remove(r'C:\Users\ooaka\Downloads\委託業者日次実績表1075.PDF')