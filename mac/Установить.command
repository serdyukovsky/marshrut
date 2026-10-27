#!/bin/bash
# Установка «Маршрутный лист» на Mac. Запускается один раз.
cd "$(dirname "$0")" || exit 1
echo "Устанавливаю «Маршрутный лист»…"
echo

PY=""
for p in python3.13 python3.12 python3.11 python3.10 python3; do
  if command -v "$p" >/dev/null 2>&1 && "$p" -c "import tkinter" >/dev/null 2>&1; then PY="$p"; break; fi
done
if [ -z "$PY" ]; then
  echo "Нужен Python с сайта python.org — сейчас открою страницу загрузки."
  echo "Установи его (кнопка «Download Python 3.x»), потом снова запусти этот файл."
  open "https://www.python.org/downloads/macos/"
  read -r -p "Нажми Enter, чтобы закрыть окно."
  exit 1
fi
echo "Python: $("$PY" --version)"

"$PY" -m venv .venv || { echo "Не получилось создать окружение Python."; read -r; exit 1; }
.venv/bin/python -m pip install -q --upgrade pip
.venv/bin/python -m pip install -q -r requirements.txt || { echo "Не получилось установить библиотеки. Проверь интернет."; read -r; exit 1; }

if [ ! -f marshrut.py ]; then
  curl -fsSL "https://raw.githubusercontent.com/serdyukovsky/marshrut/main/marshrut.py" -o marshrut.py
fi

cat > "Маршрутный лист.command" <<'EOF'
#!/bin/bash
# Запуск программы. Окно Терминала можно закрыть — программа продолжит работать.
cd "$(dirname "$0")" || exit 1
nohup .venv/bin/python marshrut.py >/dev/null 2>&1 &
disown
EOF
chmod +x "Маршрутный лист.command"

echo
echo "Готово! Дальше запускай программу файлом «Маршрутный лист.command» в этой папке."
open "Маршрутный лист.command"
