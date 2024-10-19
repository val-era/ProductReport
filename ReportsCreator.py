import BSReport
import SalesAnalysReport
import UploadPhoto

class MainStart:
    def __init__(self):
        pass

    def start(self):
        action = input(
            "Добро пожаловать в программу для создания отчетов, подробную инструкцию вы можете изучить на "
            "https://github.com/val-era/ProductReport\n"
            "Выберите действие, которое хотите выполнить. Для выбора введите цифру ниже после [+]\n\n"
            "Работа с отчетами:\n"
            "Для формирования отчетов Best Sellers и Sales Analyses введите 1\n"
            "Для загрузки фотографий в выбранный вами файл введите 2\n\n"
            "Работа с фотографиями:\n"
            "Для записи фотографий в базу данных введите 3\n"
            "Для замены существующих фотографий в базе данных введите 4\n"
            "Для сохранения фотографий в папку из вашего списка введите 5\n"
            "Для проверки соответствия фотографий в папке, списка в эксель и удаления лишних фото из папки введите 6\n"
            "Для удаления фотографий из базы данных по списку введите 7\n\n[+]"
        )

        if action == "1":
            file = input("Введите путь к файлу эксель по образцу из инструкции\n[+]")
            BSReport.Processing().process(file)
            SalesAnalysReport.Processing().process(file)

        if action == "2":
            UploadPhoto.GetImage().HoHo()
        if action == "3":
            UploadPhoto.GetImage().open_folder()
        if action == "4":
            UploadPhoto.GetImage().replace_photo()
        if action == "5":
            UploadPhoto.GetImage().save_photo_infolder()
        if action == "6":
            UploadPhoto.GetImage().check_photo_list()
        if action == "7":
            UploadPhoto.GetImage().del_photo_file()


if __name__ == "__main__":
    start = MainStart()
    start.start()