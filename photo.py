import pandas as pd
import os

file_name = "C:\\Users\\User\\Desktop\\Файл для фото.xlsx"
photos = input("Введите путь к папке с фотографиями \n {+}")
df = pd.read_excel(file_name)
df_photo = []
for index, row in df.iterrows():
    for i in row.values:
        df_photo.append(i)
directory = os.fsencode(photos)
for file in os.listdir(directory):
    filename = os.fsdecode(file)
    file_name = f"{photos}\\{filename}"
    art = (filename[: filename.find('.')])
    if art in df_photo:
        pass
    else:
        os.remove(file_name)



