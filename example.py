from extract_routes import Extraction


def main():
    src = str(sys.argv[1])
    dest = str(sys.argv[2])
    Extraction(src, dest)

if __name__ == "__main__":
    main()
