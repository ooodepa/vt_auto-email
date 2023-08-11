## Как запустить программу на Windows 10

1. Установить [Python 3.11.4](https://www.python.org/downloads/).
1. Устанавливаем зависимости Python через командную строку:
    1. Открываем окно выполнить (`Win` + `R`).
    1. Вводим `cmd` и жмём `Enter` - открылась командная строка.
    1. Вводим команды в командную строку.
        ```cmd
        pip install pandas
        pip install openpyxl
        ```
1. Запускаем программу нажав на файл `start.bat`, либо выполняй действия ниже:
    1. Открываем папку с проектом в проводнике (`explorer.exe`).
    1. Жмём по пути (например, `D:\**\vt_auto-email\src`).
    1. Пишем `cmd` и жмём `Enter`  - открылась командная строка.
    1. Запускаем программу скриптом
        ```cmd
        python AutoEmail.py
        ```