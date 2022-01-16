import os
import sys
sys.path.append(os.path.dirname(os.path.dirname(__file__)))

from satofuru.controller import satofuru_control

def main():
    satofuru_control.control()

if __name__ == '__main__':
    main()