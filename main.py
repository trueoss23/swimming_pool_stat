from file_svc.file_svc_uc import FileSvc


def main():
    fs = FileSvc()
    workbook = fs.parse_excel_file("common/swim_igor.xlsx")


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    main()

# See PyCharm help at https://www.jetbrains.com/help/pycharm/
