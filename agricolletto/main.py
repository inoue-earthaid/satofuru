import os
import sys
sys.path.append(os.path.dirname(os.path.dirname(__file__)))
from agricolletto.controller import control_agricolletto

def main():
    control_agricolletto.agricolletto()

if __name__ == '__main__':
    main()
    file_path = r'C:\Users\ooaka\Downloads\委託業者日次実績表1075.PDF'
    os.remove(file_path)
