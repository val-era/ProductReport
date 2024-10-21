Добро пожаловать в программу для создания отчетов

При запуске в консоли изначально неоходимо выбрать отчет для создания
**Выберите действие, которое хотите выполнить. Для выбора введите цифру ниже после [+] и нажмите ENTER**

            Работа с отчетами:
            Для формирования отчетов Best Sellers и Sales Analyses введите 1
            Для загрузки фотографий в выбранный вами файл введите 2
            
            Работа с фотографиями:
            Для записи фотографий в базу данных введите 3
            Для замены существующих фотографий в базе данных введите 4
            Для сохранения фотографий в папку из вашего списка введите 5
            Для проверки соответствия фотографий в папке, списка в эксель и удаления лишних фото из папки введите 6
            Для удаления фотографий из базы данных по списку введите 7

На вход принимается **абсолютный путь к файлу в формате C:/User/file.xslx или если он находится в одной папке с программой то file.xslx**


Для  первого пункта загрузочный файл должен быть **СТРОГО в формате ниже с теми-же именами колонок**:
[Магазин, Артикул, Наименование, Модель ANTA, Тип продукта,	Количество,	Категория, Сезон, SKU,	RRP, RRP Amount, Grade,	Пол, Коллекция, Актуальный заказ, Сумма со скидкой, Кол-во продаж]
80% Данных колонок формируется в отчете по остаткам, https://github.com/val-era/StockGrade2.

![image](https://github.com/user-attachments/assets/cda6a747-a3f1-4a6a-8163-00a439f6e1ca)

**Обратите внимание [] означают строгое соответствие записываемых значений в формате внутри скобочек**

**Магазин** - Наименование магазина
**Артикул** - Артикул товара
**Наименование** - можно оставить пустым
**Модель ANTA** - имя модели
**Тип продукта** - группа продукта (Куртка и тп)
**Количество** - кол-во штучек на остатке
**Категория** - категория товара в формате **[APP, FTW, ACC]**
**Сезон** - сезон товара (Можно оставить пустым)
**SKU** - **[1]** используется для подсчета артикулов
**RRP, RRP Amount** - цена и суммма стока в деньгах, необходима для усовершенствования отчета и дальнейших метрик, в данный момент не задействована
**Grade** **[A, B, C]** - грейд остатков товара
**Пол** - **[MEN, WOMEN, UNISEX, KIDS]**
**Коллекция** - коллекция продукта
**Актуальный заказ** - нахождение артикула в матрице магазина **[Актуально, Не актуально]**
**Сумма со скидкой** - сумма продаж артикула со скидкой
**Кол-во продаж** - сумма проданных штук

По данному файлу формируются два отчета и сохраняются в папку с программой. Имена отчетов (**Best Sellers Report.xlsx**, **Sales Analyses Report.xlsx**)

Обратите внимание, что для выгрузки фотографий в отчет изначально необходимо их загрузить в базу данных, которая хранится в папке с отчетами. **Не удаляйте БД и не меняйте ее содержимое сторонними программами**. Она может весить много, в зависимости от кол-во фотографий. Бд будет называться ProductReportDB.db. Не удаляйте ее, иначе у вас не будет фотографий.

Основные метрики отчета расположены на 1 вкладке отчета.
