# ТЗ: новая база данных поставщиков вина

## 1. Цель

Создать чистую и проверяемую базу поставщиков вина для анализа ресторанных винных карт.

База должна позволять:

- находить точного поставщика позиции из винной карты;
- сравнивать ресторанную цену с закупочной;
- показывать альтернативных поставщиков;
- отделять уверенные совпадения от спорных;
- исключать служебные строки, аксессуары, наборы и не-винные товары.

Старая база `wine_finder_bot/prices.db` используется только как источник/референс. Новая база собирается заново.

## 2. Источники данных

Основной источник:

- `wine_finder_bot/pricelists/*`

Поддерживаемые форматы:

- `.xlsx`
- `.xls`
- `.csv` желательно во v2

Каждый файл считается прайсом одного поставщика. Название поставщика по умолчанию берётся из имени файла, но должно поддерживаться ручное переименование через конфиг.

Нужен конфиг:

- `supplier_db/suppliers.yaml`

В нём:

- каноническое имя поставщика;
- список файлов/масок файлов;
- правила колонок, если автоопределение ошибается;
- флаг активности прайса;
- дата обновления, если известна.

## 3. Выходные артефакты

Скрипт сборки должен создавать:

- `supplier_db/suppliers.db` — основная SQLite-база;
- `supplier_db/suppliers.json` — компактный индекс для браузерного анализатора;
- `supplier_db/rejected_rows.csv` — строки, исключённые фильтрами;
- `supplier_db/build_report.csv` — отчёт качества по каждому источнику;
- `supplier_db/aliases.json` — словарь алиасов производителей и поставщиков.

Браузерный анализатор должен использовать `supplier_db/suppliers.json` вместо старого `hedonis/suppliers.json`.

## 4. Схема SQLite

### Таблица `supplier_items`

Обязательные поля:

- `id` INTEGER PRIMARY KEY
- `supplier` TEXT NOT NULL
- `source_file` TEXT NOT NULL
- `source_sheet` TEXT
- `source_row` INTEGER
- `raw_name` TEXT NOT NULL
- `raw_price` TEXT
- `price` REAL
- `currency` TEXT DEFAULT `RUB`
- `producer` TEXT
- `producer_normalized` TEXT
- `wine_name` TEXT
- `cuvee` TEXT
- `vintage` INTEGER
- `is_nv` INTEGER DEFAULT 0
- `volume_l` REAL
- `country` TEXT
- `region` TEXT
- `appellation` TEXT
- `color` TEXT
- `style` TEXT
- `grape_varieties` TEXT
- `normalized_name` TEXT NOT NULL
- `search_tokens` TEXT NOT NULL
- `match_key` TEXT
- `quality_score` INTEGER DEFAULT 0
- `created_at` TEXT NOT NULL
- `updated_at` TEXT NOT NULL

### Таблица `rejected_rows`

Поля:

- `id` INTEGER PRIMARY KEY
- `supplier` TEXT
- `source_file` TEXT
- `source_sheet` TEXT
- `source_row` INTEGER
- `raw_values` TEXT
- `reason` TEXT NOT NULL
- `created_at` TEXT NOT NULL

### Таблица `aliases`

Поля:

- `id` INTEGER PRIMARY KEY
- `kind` TEXT NOT NULL — `producer`, `supplier`, `region`, `country`
- `alias` TEXT NOT NULL
- `canonical` TEXT NOT NULL
- `source` TEXT — `manual`, `auto`, `import`

## 5. Импорт и парсинг

Скрипт:

- `supplier_db/build_supplier_db.py`

Команда:

```bash
python3 supplier_db/build_supplier_db.py
```

Алгоритм:

1. Прочитать `supplier_db/suppliers.yaml`.
2. Для каждого активного прайса открыть все листы.
3. Найти строку заголовков.
4. Определить колонки:
   - название;
   - цена;
   - год;
   - объём;
   - страна;
   - регион;
   - цвет/тип.
5. Прочитать строки после заголовка.
6. Отфильтровать мусор.
7. Нормализовать вино.
8. Записать принятые строки в `supplier_items`.
9. Записать отфильтрованные строки в `rejected_rows`.
10. Собрать `suppliers.json`.
11. Сформировать `build_report.csv`.

Если в конфиге для файла задана ручная карта колонок, она важнее автоопределения.

## 6. Правила фильтрации

Строка должна быть исключена, если:

- нет названия;
- название короче 4 символов;
- нет цены или цена не числовая;
- цена меньше `500 ₽`, кроме явно разрешённых категорий;
- строка похожа на заголовок раздела;
- строка является упаковкой, аксессуаром или служебной позицией.

Ключевые слова для исключения:

- `коробка`
- `подарочная упаковка`
- `упаковка`
- `пакет`
- `набор`
- `бокал`
- `пробка`
- `штопор`
- `декантер`
- `сертификат`
- `доставка`
- `депозит`
- `тара`
- `pos`
- `gift box`, если строка не содержит полноценного названия вина

Крепкий алкоголь, пиво, вода и безалкогольные напитки исключаются для v1.

## 7. Нормализация вина

Для каждой принятой строки нужно извлечь:

- производитель;
- название/кюве;
- год;
- NV-флаг;
- объём;
- страна;
- регион/апелласьон, если есть;
- цвет/стиль, если есть.

Минимальные эвристики:

- `2010–2035` считать винтажом;
- `NV`, `N/V`, `non-vintage` считать невинтажным вином;
- `0.75`, `0,75`, `750 ml`, `750мл` приводить к `0.75`;
- латинскую часть после `/` считать важным алиасом;
- кириллицу и латиницу хранить вместе, но токенизировать отдельно.

`normalized_name`:

- lower-case;
- без диакритики;
- `ё -> е`;
- только буквы/цифры/пробелы;
- схлопнутые пробелы.

`search_tokens`:

- JSON-массив уникальных токенов;
- исключить стоп-слова: `wine`, `вино`, `domaine`, `chateau`, `brut`, `rouge`, `blanc`, `doc`, `docg`, `igt`, `aoc`, `nv` и т.п.;
- исключить числа, кроме значимых частей названия.

`match_key`:

```text
producer_normalized + cuvee_normalized + vintage_or_nv + volume_l
```

## 8. Дедупликация

Дубликаты внутри одного поставщика объединяются, если совпадают:

- `supplier`;
- `producer_normalized`;
- `wine_name / cuvee`;
- `vintage` или `is_nv`;
- `volume_l`.

Если цены отличаются:

- хранить минимальную цену как `price`;
- в будущем добавить таблицу истории цен.

Дубликаты между поставщиками не объединяются: это разные предложения рынка.

## 9. Качество данных

Для каждой строки считать `quality_score` от 0 до 100:

- `+25` есть производитель;
- `+25` есть название/кюве;
- `+15` есть цена;
- `+10` есть год или NV;
- `+10` есть объём;
- `+5` есть страна;
- `+5` есть регион/апелласьон;
- `+5` есть цвет/стиль.

Строки с `quality_score < 45` не должны попадать в браузерный `suppliers.json`, но могут оставаться в SQLite для ручной проверки.

## 10. JSON-экспорт для анализатора

Файл:

- `supplier_db/suppliers.json`

Формат:

```json
{
  "version": 1,
  "generated_at": "2026-06-27T12:00:00Z",
  "count": 0,
  "suppliers": [],
  "items": []
}
```

Элемент `items`:

```json
{
  "id": 1,
  "supplier": "Лудинг",
  "producer": "Dom Perignon",
  "wine_name": "Vintage Brut",
  "cuvee": "Vintage",
  "vintage": 2013,
  "is_nv": false,
  "volume_l": 0.75,
  "country": "Франция",
  "region": "Champagne",
  "appellation": "Champagne AOC",
  "color": "Игристое",
  "price": 41681,
  "currency": "RUB",
  "raw_name": "Dom Perignon Vintage Moet Hennessey in gift box",
  "normalized_name": "dom perignon vintage moet hennessey gift box",
  "search_tokens": ["dom", "perignon", "vintage", "moet", "hennessey"]
}
```

JSON не должен содержать:

- rejected rows;
- служебные товары;
- строки с `quality_score < 45`;
- пустые `supplier`, `raw_name`, `price`.

## 11. Матчинг винной карты

Анализатор должен сравнивать позицию карты с базой по этапам:

1. Производитель.
2. Название/кюве.
3. Винтаж или NV.
4. Объём.
5. Цена как sanity check.

Статусы:

- `match` — уверенное совпадение;
- `review` — похожая позиция, нужна ручная проверка;
- `missing` — нет надёжного совпадения.

Ориентировочные пороги:

- `match >= 85`
- `review 60–84`
- `missing < 60`

Запрещено считать уверенным совпадением строку, если:

- совпал только регион или сорт;
- не совпал производитель;
- найден аксессуар/набор/коробка;
- цена поставщика подозрительно мала для категории.

## 12. Отчёт сборки

`build_report.csv` должен содержать:

- `supplier`
- `source_file`
- `source_sheet`
- `rows_total`
- `rows_accepted`
- `rows_rejected`
- `without_price`
- `without_producer`
- `without_vintage_or_nv`
- `without_volume`
- `avg_quality_score`
- `notes`

Цель качества для v1:

- не менее `80%` принятых строк имеют цену;
- не менее `70%` имеют производителя;
- не менее `60%` имеют год/NV;
- не менее `70%` имеют объём;
- rejected rows доступны для ручного аудита.

## 13. Приёмка

База считается готовой для v1, если:

- сборка запускается одной командой;
- создаются все выходные артефакты;
- `suppliers.db` содержит не менее `25` поставщиков;
- `suppliers.json` загружается в `hedonis/analyzer.html`;
- служебные товары не попадают в JSON;
- тестовые позиции находятся корректно:
  - `Dom Perignon Brut 2013`;
  - `Laherte Freres Ultradition Brut NV`;
  - `Chateau Gloria St.-Julien 2017`;
  - `Opus One 2018`;
  - `Chateauneuf-du-Pape La Bernardine M. Chapoutier 2021`;
- для каждой тестовой позиции видны до 3 альтернативных поставщиков;
- для спорных совпадений статус `review`, а не `match`.

## 14. Структура проекта

```text
supplier_db/
  TECH_SPEC.md
  suppliers.yaml
  build_supplier_db.py
  suppliers.db
  suppliers.json
  rejected_rows.csv
  build_report.csv
  aliases.json
```

## 15. Ограничения v1

В v1 не требуется:

- история изменения цен;
- онлайн-обновление прайсов;
- web-admin для ручной правки;
- OCR прайсов из PDF;
- ML-модель для парсинга каждой строки.

Эти функции можно добавить после появления стабильной чистой базы.
