#!/usr/bin/env python3
"""Находит рабочий стол, создаёт на нём папку AutoRu_Resellers и печатает JSON с состоянием.

Использование:
  python3 setup_folder.py                       # найти рабочий стол, создать папку, показать состояние
  python3 setup_folder.py --dir /путь/к/папке   # использовать готовую папку (например, смонтированную: $HOME/mnt/Desktop/AutoRu_Resellers)
  python3 setup_folder.py --save-criteria '{"region":"moskovskaya_oblast","models":[["toyota","camry"]],"per_model":8,"min_garage":2}'

Печатает: folder, desktop, writable, files (уже готовые файлы), next_number, criteria (последние сохранённые или null).
"""
import sys, os, json, re, datetime
from pathlib import Path

FOLDER_NAME = 'AutoRu_Resellers'
FILE_RE = re.compile(r'^(\d{3})_resellers_(\d{4}-\d{2}-\d{2})\.xlsx$')

def desktop_path():
    home = Path.home()
    if sys.platform == 'win32':
        try:  # настоящий путь рабочего стола (учитывает перенос в OneDrive)
            import ctypes
            from ctypes import wintypes
            buf = ctypes.create_unicode_buffer(wintypes.MAX_PATH)
            if ctypes.windll.shell32.SHGetFolderPathW(None, 0x0000, None, 0, buf) == 0 and buf.value:
                return Path(buf.value)
        except Exception:
            pass
    for cand in [home / 'Desktop', home / 'OneDrive' / 'Desktop', home / 'OneDrive' / 'Рабочий стол', home / 'Рабочий стол']:
        if cand.is_dir():
            return cand
    return None

def main():
    out = {}
    if '--dir' in sys.argv:
        folder = Path(sys.argv[sys.argv.index('--dir') + 1]).expanduser()
        out['desktop'] = None
    else:
        d = desktop_path()
        if d is None:
            print(json.dumps({'error': 'desktop_not_found', 'hint': 'рабочий стол не найден: спросите путь у пользователя и вызовите с --dir'}, ensure_ascii=False))
            sys.exit(2)
        folder = d / FOLDER_NAME
        out['desktop'] = str(d)
    folder.mkdir(parents=True, exist_ok=True)
    try:
        t = folder / '.write_test'; t.write_text('ok'); t.unlink(); out['writable'] = True
    except Exception as e:
        out['writable'] = False; out['write_error'] = str(e)
    cfg = folder / 'criteria.json'
    if '--save-criteria' in sys.argv:
        data = json.loads(sys.argv[sys.argv.index('--save-criteria') + 1])
        data['saved_at'] = datetime.date.today().isoformat()
        cfg.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding='utf-8')
    files = sorted(p.name for p in folder.iterdir() if FILE_RE.match(p.name))
    nums = [int(FILE_RE.match(n).group(1)) for n in files]
    out.update(folder=str(folder), files=files, next_number=(max(nums) + 1 if nums else 1),
               criteria=(json.loads(cfg.read_text(encoding='utf-8')) if cfg.exists() else None))
    sys.stdout.reconfigure(encoding='utf-8')
    print(json.dumps(out, ensure_ascii=False))

if __name__ == '__main__':
    main()
