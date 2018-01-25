import sys
from extract_routes import Extraction

""" This is used as follows:

python example.py /home/test/directory_full_of_xls test.xls ACC_ON
"""

def main():
    src = str(sys.argv[1])
    dest = str(sys.argv[2])
    value = str(sys.argv[3])
    Extraction(src, dest, value)

if __name__ == "__main__":
    main()
