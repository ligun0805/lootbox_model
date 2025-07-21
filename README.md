# Lootbox Model Simulator

Этот проект позволяет моделировать выпадение предметов из лутбоксов, строить **Pity System** (механику жалости), рассчитывать экономическую модель и генерировать готовые Excel и Word отчёты с данными и графиками.

---

## Структура проекта

```
lootbox_model.py      # Основной скрипт симуляции
requirements.txt      # Список зависимостей для pip
README.md             # Инструкция по запуску (этот файл)
outputs/              # Папка, куда будут сохраняться результаты
```

---

## Требования

* Python 3.7 или новее
* pip (или pip3)
* Рекомендуется виртуальное окружение (venv)

Зависимости указаны в `requirements.txt`:

```text
pandas
python-docx
openpyxl
XlsxWriter
```

---

## Установка

1. **Клонируйте или скопируйте** папку проекта на локальный диск.

2. **Создайте виртуальное окружение** (рекомендуется):

   ```bash
   python3 -m venv .venv   # или `python -m venv .venv`
   # Windows (PowerShell):
   .\.venv\Scripts\Activate.ps1
   # Windows (CMD):
   .\.venv\Scripts\activate.bat
   # Linux/macOS:
   source .venv/bin/activate
   ```

3. **Установите зависимости**:

   ```bash
   pip install -r requirements.txt
   ```

---

## Запуск скрипта

```bash
python lootbox_model.py [-n N]
```

* `-n, --opens` — количество открытий лутбокса для симуляции (по умолчанию `100000`).

### Примеры

* **50000 открытий**:

  ```bash
  python lootbox_model.py -n 50000
  ```

* **По умолчанию (100000 открытий)**:

  ```bash
  python lootbox_model.py
  ```

---

## Результаты

После запуска создаётся новая папка:

```
outputs/YYYYMMDD_HHMMSS/
```

где:

* `YYYYMMDD_HHMMSS` — метка времени запуска.

Внутри неё генерируются:

* **lootbox\_model.xlsx** с листами:

  * `DropRates` — базовые вероятности
  * `PitySystem` — кривые роста шансов (Common, Rare, Epic, Legendary)
  * `SimulationResults` — реальные частоты выпадений
  * `EconomicModel` — сценарии доходности

  - встроенные графики для `PitySystem` и `SimulationResults`

* **lootbox\_model.docx** — подробное описание механик и таблицы по тем же данным.

---

## Сборка исполняемого файла

### Windows и macOS (PyInstaller)

1. Установите PyInstaller в том же (виртуальном) окружении:

   ```bash
   pip install pyinstaller
   ```

2. Сгенерируйте `.exe` / бинарь:

   ```bash
   pyinstaller --onefile \
     --hidden-import=openpyxl \
     --collect-all pandas \
     --collect-all openpyxl \
     lootbox_model.py
   ```

3. Готовый файл будет в `dist/lootbox_model.exe` (Windows) или `dist/lootbox_model` (macOS).

4. Запускайте:

   ```bash
   dist/lootbox_model.exe -n 50000
   ```

---

