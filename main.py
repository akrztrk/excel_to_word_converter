from converter.controller import ConverterController

def main():
    excel_path = "data/test_v3.xlsx"
    controller = ConverterController(excel_path)
    controller.convert_all_sheets()
    print("Tüm sayfalar Word'e dönüştürüldü.")

if __name__ == "__main__":
    main()