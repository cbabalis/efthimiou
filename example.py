import sys
from extract_routes import Extraction


def main():
    src = str(sys.argv[1])
    dest = str(sys.argv[2])
    value = str(sys.argv[3])
    Extraction(src, dest, value)

if __name__ == "__main__":
    main()
