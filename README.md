# Price Mail Bot

Отдельный проект (не `debt_bot`):
- по таймеру меняет одну случайную цену в Excel на `+1`/`-1` RUB по очереди
- отправляет обновленный файл на почту Яндекса
- принимает новый `.xlsx` через Telegram и атомарно обновляет серверный файл без сброса таймера

## Быстрый старт

1. Скопируй `.env.example` в `.env` и заполни:
   - `PRICE_BOT_TOKEN`
   - `PRICE_BOT_ALLOWED_IDS`
   - `YANDEX_SMTP_LOGIN`
   - `YANDEX_SMTP_PASSWORD` (пароль приложения Яндекса)
   - `YANDEX_EMAIL_TO`

2. Положи исходный файл прайса:
   - по умолчанию: `./data/price.xlsx`
   - или настрой через `PRICE_FILE_PATH`

3. Запусти:
```bash
docker-compose up -d --build
```

Логи:
```bash
docker-compose logs -f
```

## Команды бота

- `/start` — описание работы
- `/status` — текущий статус (файл/интервал/почта)
- отправить `.xlsx` — обновить рабочий файл без сброса таймера
