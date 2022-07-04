#Библиотека для работы с ссылками
import webbrowser as web
#Библиотека для преобразования страницу в url
from urllib.parse import quote
#Библиотека для эмуляции нажатий клавиши
import pyautogui as pg
#
from io import BytesIO
import win32clipboard
from PIL import Image

#Библиотека для работы с Exel
import pandas as pd
#Библиотека для работы со ввременем
from datetime import datetime, date
import time

#Наименование столбцов таблицы
slovar = [
        "номер телефона",
        "сообщение",
        "часы",
        "минуты",
        "секунды",
        "изображение"
        ]

#Выгрузка таблицы
ttk_table = pd.read_excel("ttk.xlsx", names = slovar)

class What():
    def __init__(self, what_table):
        self.what_table = ttk_table
        self.sleep: int
        self.wait_time = 13

    def _time(self, time_hour, time_min, time_sec):
        current_time = time.localtime()
        left_time = datetime.strptime(
        f"{time_hour}:{time_min}:{time_sec}", "%H:%M:%S"
    ) - datetime.strptime(
        f"{current_time.tm_hour}:{current_time.tm_min}:{current_time.tm_sec}",
        "%H:%M:%S",
    )
        sleep_time = left_time.seconds - self.wait_time
        return(sleep_time)

    def _run_script(self):
        for nn in self.what_table.itertuples(index=False):
            self.sleep = self._time(time_hour = nn[2], time_min = nn[3], time_sec = nn[4])
            if nn[5] == "None":
                self._f_message(f"+{nn[0]}", nn[1], self.sleep)                                 
            else:
                self._f_image(f"+{nn[0]}", self.sleep, nn[5])

    def _f_message(self, phone_no, message, sleep):
        print(f"Отправка начнется через {sleep}с.")
        time.sleep(sleep)
        web.open(f"https://web.whatsapp.com/send?phone={phone_no}&text={quote(message)}")
        time.sleep(self.wait_time)
        pg.click()
        time.sleep(1)
        pg.keyDown("enter")
        time.sleep(3)
        pg.hotkey('ctrl', 'w')

    def _f_image(self, phone_no, sleep, image):
        print(f"Отправка начнется через {sleep}с.")
        time.sleep(sleep)
        self._open_image(image)
        web.open(f"https://web.whatsapp.com/send?phone={phone_no}")
        time.sleep(self.wait_time)
        pg.click()
        pg.hotkey('ctrl', 'v')
        time.sleep(1)
        pg.keyDown("enter")
        time.sleep(3)
        pg.hotkey('ctrl', 'w')

    def _copy_image(self, clip_type, data):
        win32clipboard.OpenClipboard()
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardData(clip_type, data)
        win32clipboard.CloseClipboard()

    def _open_image(self, image):
        filepath = image
        image = Image.open(filepath)
        output = BytesIO()
        image.convert("RGB").save(output, "BMP")
        data = output.getvalue()[14:]
        output.close()
        self._copy_image(win32clipboard.CF_DIB, data)




if __name__ == "__main__":
    run = What(ttk_table)
    run._run_script()
