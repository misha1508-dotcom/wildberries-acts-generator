#!/bin/bash

# Получаем директорию где находится скрипт
DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"
cd "$DIR"

# Проверяем установлен ли Python
if ! command -v python3 &> /dev/null; then
    osascript -e 'display dialog "Python 3 не установлен!\n\nПожалуйста, установите Python с python.org" buttons {"OK"} default button "OK" with icon stop'
    exit 1
fi

# Проверяем установлены ли зависимости
if ! python3 -c "import flask" 2>/dev/null; then
    osascript -e 'display dialog "Установка зависимостей...\n\nЭто займет несколько минут." buttons {"OK"} default button "OK" with icon note'

    # Открываем терминал для установки зависимостей
    osascript <<END
        tell application "Terminal"
            activate
            do script "cd '$DIR' && pip3 install -r requirements.txt && echo '\n\nЗависимости установлены! Закройте это окно и запустите приложение снова.' && read -p 'Нажмите Enter для закрытия...'"
        end tell
END
    exit 0
fi

# Запускаем приложение
python3 desktop_app.py

# Если произошла ошибка
if [ $? -ne 0 ]; then
    osascript -e 'display dialog "Произошла ошибка при запуске приложения.\n\nПроверьте установку Python и зависимостей." buttons {"OK"} default button "OK" with icon stop'
fi
