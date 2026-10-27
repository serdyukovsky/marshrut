"""Запускалка для Windows (из неё собирается «Маршрутный лист.exe»).

Сама программа — файл marshrut.py рядом с exe. Его можно обновлять без пересборки exe:
программа при запуске сама скачивает новую версию с GitHub.
"""
import os
import runpy
import sys

# Все библиотеки программы импортируются здесь, чтобы сборщик положил их внутрь exe.
import customtkinter  # noqa: F401
import openpyxl  # noqa: F401
import pypdf  # noqa: F401
import pypdfium2  # noqa: F401
import requests
import xlrd  # noqa: F401

URL = "https://raw.githubusercontent.com/serdyukovsky/marshrut/main/marshrut.py"
HERE = os.path.dirname(sys.executable if getattr(sys, "frozen", False) else os.path.abspath(__file__))
SCRIPT = os.path.join(HERE, "marshrut.py")


def say(text):
    import tkinter
    from tkinter import messagebox
    root = tkinter.Tk()
    root.withdraw()
    messagebox.showerror("Маршрутный лист", text)
    root.destroy()


def download():
    r = requests.get(URL, timeout=30)
    r.raise_for_status()
    compile(r.content.decode("utf-8"), "marshrut.py", "exec")
    with open(SCRIPT + ".new", "wb") as f:
        f.write(r.content)
    os.replace(SCRIPT + ".new", SCRIPT)


def main():
    if not os.path.exists(SCRIPT):
        try:
            download()
        except Exception:
            say("Не нашёл файл marshrut.py рядом с программой и не смог скачать его.\n"
                "Проверь интернет или скачай программу заново:\n"
                "github.com/serdyukovsky/marshrut")
            return
    sys.argv[0] = SCRIPT
    try:
        runpy.run_path(SCRIPT, run_name="__main__")
    except ImportError as e:
        say("Новой версии программы нужна библиотека, которой нет в этом exe (%s).\n"
            "Скачай свежую версию для Windows: github.com/serdyukovsky/marshrut" % e)


if __name__ == "__main__":
    main()
