# ukragro-prices
Автоматичне об’єднання щоденних цін UkrAgroConsult.

## Як запустити
- **GitHub:** Actions → «Оновити загальний Excel» → Run workflow. Готовий файл — в Artifacts запуску.
  Галочка «full_rebuild» перечитує всі файли з `input/` з нуля.

## Що де лежить
| Шлях | Призначення |
|---|---|
| `input/` | Щоденні Excel-файли UkrAgroConsult (аркуші `ДД.ММ.РРРР_min_max`). |
| `data/prices_db.csv` | Накопичена база всіх цін. Оновлюється й комітиться після кожного запуску. |
| `data/processed_files.csv` | Які файли вже в базі (назва + SHA256). Незмінені файли повторно не читаються. |
| `output/UKRAGRO_ALL_PRICES.xlsx` | Результат. |
