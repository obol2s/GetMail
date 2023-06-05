import email
import imaplib
import os.path
import time
from email.header import decode_header
from win32com.client import Dispatch

mail_pass = 'password'
username = '@mail.ru'
imap_server = 'imap.mail.ru'
attachment = r'D:\path'


def get_mail():
    imap = imaplib.IMAP4_SSL(imap_server)
    imap.login(username, mail_pass)
    imap.select('Remains')

    _, data = imap.uid('search', 'UNSEEN')  # Поиск непрочитанных писем.

    if data[0]:  # Если найдены непрочитанные письма:
        # Закрываем открытые файлы.
        xl = Dispatch('Excel.Application')  # Создание экземпляра приложения Excel
        workbooks = xl.Workbooks  # Получение коллекции открытых книг

        # Закрытие каждой книги в коллекции, только если путь к файлу соответствует директории
        for workbook in workbooks:
            if os.path.dirname(workbook.FullName) == attachment:
                workbook.Close(False)

        # Удаляем предыдущие файлы из папки с вложениями.
        if os.path.exists(attachment):
            for f in os.listdir(attachment):
                os.remove(os.path.join(attachment, f))

            email_uid = data[0].split()[-1]  # Получаем UID последнего непрочитанного письма.
            _, msg = imap.uid('fetch', email_uid, '(RFC822)')  # Получаем письмо по UID (кортеж байт).
            msg = msg[0][1]  # Получаем объект bytes сообщения.
            msg_set = email.message_from_bytes(msg)  # Преобразуем объект bytes в объект сообщения.

            for part in msg_set.walk():  # Перебор всех частей сообщения.
                if part.get_content_disposition() == 'attachment':  # Если часть - вложение:
                    part.get_filename()  # Получаем имя файла вложения.
                    try:
                        filename = decode_header(part.get_filename())[0][0].decode('utf-8')  # Декодируем имя файла.
                    except AttributeError:
                        filename = part.get_filename()
                    print(filename)  # Выводим имя файла вложения.
                    with open(os.path.join(attachment, filename), 'wb') as new_file:  # Открываем файл для записи.
                        new_file.write(part.get_payload(decode=True))

        for num in data[0].split():
            imap.uid('store', num, '+FLAGS', '\\SEEN')


while True:
    get_mail()
    time.sleep(5)
