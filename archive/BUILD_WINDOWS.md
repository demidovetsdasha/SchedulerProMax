# Сборка ScheduleParser.exe на Windows

Инструкция по сборке программы «Парсер расписания Excel → Word» в `.exe` для Windows.

> Собирать нужно **на самой Windows**: PyInstaller упаковывает Python и библиотеки той
> системы, на которой запущен. На macOS получится программа только для macOS.

## Что нужно

- Windows 10 или 11, 64-бит (PySide6 не поддерживает более старые версии).
- Python 3.12 или 3.13.
- Интернет — для установки библиотек.

## 1. Установить Python

1. Скачайте установщик с <https://www.python.org/downloads/windows/>
   («Windows installer (64-bit)»).
2. В первом окне установщика **обязательно** отметьте
   **«Add python.exe to PATH»**, затем нажмите «Install Now».
3. Проверьте установку: откройте командную строку (`Win + R` → `cmd` → Enter) и выполните:
   ```bat
   py --version
   ```
   Должна вывестись версия, например `Python 3.13.x`.

## 2. Скопировать проект

Скопируйте папку проекта на компьютер с Windows, например в `C:\ScheduleParser`.

Нужны только эти файлы и папки:

```
main.py
requirements.txt
schedule_app\
config\
templates\
```

**Не копируйте папку `.venv`**: она собрана под macOS и на Windows не заработает.
Папку `__pycache__` тоже можно не копировать.

## 3. Собрать exe

Откройте командную строку в папке проекта. Проще всего открыть папку в Проводнике,
ввести в адресной строке `cmd` и нажать Enter. Затем по очереди выполните:

```bat
py -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt pyinstaller
pyinstaller --noconfirm --windowed --name ScheduleParser --distpath output_exe main.py
```

- Первая команда создаёт виртуальное окружение `.venv`.
- После второй в начале строки появится `(.venv)`, значит окружение активно.
- Третья устанавливает библиотеки (PySide6, openpyxl, python-docx, xlrd, PyInstaller).
  Это может занять несколько минут.
- Четвёртая собирает программу в папку `output_exe\ScheduleParser`.

## 4. Положить конфиг и шаблон рядом с exe

Программа читает конфиг и шаблон Word из папок рядом с `ScheduleParser.exe`:

```bat
xcopy /E /I /Y config output_exe\ScheduleParser\config
xcopy /E /I /Y templates output_exe\ScheduleParser\templates
```

## 5. Проверить

Запустите `output_exe\ScheduleParser\ScheduleParser.exe`. Откройте Excel-файл
расписания, отметьте подгруппы на странице и нажмите «Сформировать Word».

Готовая папка должна выглядеть так:

```
output_exe\ScheduleParser\
    ScheduleParser.exe
    _internal\            ← библиотеки программы
    config\faculties.json
    templates\template.docx
```

## Передача программы другим

Передавайте **всю папку** `output_exe\ScheduleParser`, например в zip-архиве.
Один `ScheduleParser.exe` без папки `_internal` не запустится.
Python на компьютерах пользователей не нужен.

## Настройка после сборки

`config\faculties.json` остаётся обычным текстовым файлом рядом с exe. Пересобирать
программу после его правки не нужно:

1. Откройте файл в Блокноте (или кнопкой «Открыть конфиг» в программе).
2. Измените, например, `deputy_dean` у нужного факультета и сохраните файл
   в кодировке UTF-8.
3. Нажмите в программе «Перечитать конфиг».

Шаблон Word можно заменить так же: положите новый файл в
`templates\template.docx`.

## Вариант: один exe-файл

Если удобнее один файл вместо папки, добавьте `--onefile`:

```bat
pyinstaller --noconfirm --windowed --onefile --name ScheduleParser --distpath output_exe main.py
xcopy /E /I /Y config output_exe\config
xcopy /E /I /Y templates output_exe\templates
```

Папки `config` и `templates` всё равно должны лежать рядом с `ScheduleParser.exe`.
Минусы: программа дольше запускается (распаковывается при каждом старте),
и антивирусы чаще ошибочно на неё реагируют.

## Пересборка после изменения кода

```bat
cd C:\ScheduleParser
.venv\Scripts\activate
pyinstaller --noconfirm --windowed --name ScheduleParser --distpath output_exe main.py
xcopy /E /I /Y config output_exe\ScheduleParser\config
xcopy /E /I /Y templates output_exe\ScheduleParser\templates
```

Шаги 1 и `pip install` повторять не нужно. Внимание: последние две команды
перезапишут `config` в папке программы версией из проекта.

## Возможные проблемы

| Проблема | Решение |
|---|---|
| `'py' не является внутренней или внешней командой` | Python не добавлен в PATH. Переустановите Python с отметкой «Add python.exe to PATH». |
| `.venv\Scripts\activate` выдаёт ошибку в PowerShell | Используйте `cmd`, а не PowerShell, или выполните в PowerShell `Set-ExecutionPolicy -Scope CurrentUser RemoteSigned`. |
| «Windows защитила ваш компьютер» при запуске | У программы нет цифровой подписи. Нажмите «Подробнее» → «Выполнить в любом случае». |
| Антивирус удаляет или блокирует exe | Ложное срабатывание на PyInstaller. Добавьте папку программы в исключения антивируса; со сборкой папкой (без `--onefile`) это случается реже. |
| «Не найден шаблон» в программе | Нет `templates\template.docx` рядом с exe — повторите шаг 4. |
| Программа не запускается без сообщений | Соберите без `--windowed` и запустите exe из `cmd`, чтобы увидеть текст ошибки. |
| Ошибка при сохранении Word-файла | Закройте этот файл в Word и сформируйте заново. |
