#!/usr/bin/env python3
"""Быстрая проверка скриптов: создание папки, нумерация файлов по дате, отсев дублей между файлами.
Запуск: python3 tests/smoke_test.py"""
import json, subprocess, sys, tempfile, os, re
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
PY = sys.executable

def run(*args, env=None):
    out = subprocess.run([PY, *map(str, args)], capture_output=True, text=True, env=env, encoding='utf-8')
    assert out.returncode == 0, out.stderr
    return json.loads(out.stdout)

def main():
    with tempfile.TemporaryDirectory() as tmp:
        tmp = Path(tmp)
        folder = tmp / 'AutoRu_Resellers'
        st = run(ROOT / 'scripts/setup_folder.py', '--dir', folder)
        assert st['writable'] and st['next_number'] == 1 and st['files'] == []
        rows = tmp / 'rows.txt'
        rows.write_text('toyota camry|AAA111|Тест А|3|3|Toyota Camry; Kia Rio; Lada|/toyota/camry/1-a/\n'
                        'kia rio|BBB222|Тест Б|2|2|Kia Rio; Skoda|/kia/rio/2-b/\n', encoding='utf-8')
        r1 = run(ROOT / 'scripts/build_xlsx.py', rows, folder, '--run-note', 'test 1')
        assert r1['added'] == 2 and r1['duplicates_skipped'] == 0
        assert re.search(r'001_resellers_\d{4}-\d{2}-\d{2}\.xlsx$', r1['file'])
        r2 = run(ROOT / 'scripts/build_xlsx.py', rows, folder)          # те же продавцы: файла быть не должно
        assert r2['file'] is None and r2['duplicates_skipped'] == 2
        rows.write_text(rows.read_text(encoding='utf-8') + 'kia rio|CCC333|Тест В|2|2|A; B|/kia/rio/3-c/\n', encoding='utf-8')
        r3 = run(ROOT / 'scripts/build_xlsx.py', rows, folder)          # один новый -> файл 002
        assert r3['added'] == 1 and r3['duplicates_skipped'] == 2 and '002_resellers_' in r3['file']
        ids = run(ROOT / 'scripts/known_ids.py', folder)
        assert sorted(ids) == ['AAA111', 'BBB222', 'CCC333'], ids
        st = run(ROOT / 'scripts/setup_folder.py', '--dir', folder, '--save-criteria', '{"region":"moskva","models":[["kia","rio"]]}')
        assert st['next_number'] == 3 and st['criteria']['region'] == 'moskva'
    print('OK: все проверки пройдены')

if __name__ == '__main__':
    main()
