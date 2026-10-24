#!/usr/bin/env bash
# Установка скиллов на macOS (работает и на Linux).
# Запуск из папки репозитория:  bash install_mac.sh
set -e
cd "$(dirname "$0")"

echo "1/4 Проверяю Python 3..."
if ! command -v python3 >/dev/null 2>&1; then
  echo "Python 3 не найден. Установите: xcode-select --install (или с python.org), потом запустите скрипт снова."; exit 1
fi

echo "2/4 Ставлю openpyxl (нужен для Excel)..."
python3 -c "import openpyxl" 2>/dev/null || \
  python3 -m pip install --user -r requirements.txt 2>/dev/null || \
  python3 -m pip install --user --break-system-packages -r requirements.txt
python3 -c "import openpyxl; print('openpyxl', openpyxl.__version__)"

echo "3/4 Копирую скиллы в ~/.claude/skills ..."
mkdir -p "$HOME/.claude/skills"
for s in avtoru-reseller-finder avtoru-outreach; do
  rm -rf "$HOME/.claude/skills/$s"
  cp -R "skills/$s" "$HOME/.claude/skills/$s"
  echo "  установлен: $s"
done

echo "4/4 Проверяю скрипты..."
python3 tests/smoke_test.py

cat <<'TXT'

Готово. Осталось сделать руками:
 1. Поставить Google Chrome и расширение Claude из Chrome Web Store, войти тем же аккаунтом Claude.
 2. В расширении выдать постоянное разрешение для auto.ru.
 3. Для рассылки: самому войти в аккаунт auto.ru в Chrome.
 4. Для автозапуска: чтобы Mac не засыпал, откройте Терминал и держите запущенным `caffeinate -s`
    или в Системных настройках поставьте «Не переходить в режим сна при питании от сети».
 5. Когда macOS спросит доступ к Рабочему столу, нажмите «Разрешить».
Затем скажите Claude: «найди перекупов на Авто.ру».
TXT
