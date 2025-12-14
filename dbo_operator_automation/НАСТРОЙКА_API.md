# Настройка Mailu API для автоматизации

## Включение API в Mailu

### 1. Включите API в mailu.env

На удаленной машине (10.18.2.6) отредактируйте `mailu.env`:

```bash
# Измените:
API=false
# На:
API=true

# Добавьте API токен (опционально, но рекомендуется):
API_TOKEN=ваш_случайный_токен_здесь
```

### 2. Сгенерируйте API токен (рекомендуется)

```bash
# Генерируем случайный токен
python3 -c "import secrets; print(secrets.token_urlsafe(32))"
```

Скопируйте результат в `API_TOKEN` в `mailu.env`.

### 3. Перезапустите Mailu

```bash
docker compose restart admin
# или полный перезапуск
docker compose down
docker compose up -d
```

### 4. Проверьте доступность API

```bash
# Проверка health endpoint
curl http://10.18.2.6/api/health

# Или через браузер
http://10.18.2.6/api/health
```

## Использование скрипта

### Запуск автоматизации через API:

```bash
python3 dbo_automation_api.py
```

## Преимущества API над IMAP:

1. ✅ Работает через HTTP/HTTPS (порты 80/443 уже открыты)
2. ✅ Не требует открытия портов 143/993
3. ✅ Проще для файрвола
4. ✅ Более современный подход

## Недостатки:

1. ⚠️ API должен быть включен в Mailu
2. ⚠️ Нужен API токен (опционально)
3. ⚠️ Mailu API может иметь ограничения по сравнению с IMAP

## Альтернатива: Webmail API

Если Mailu API не работает, можно использовать Webmail (Roundcube) API, который доступен через HTTP:

```python
webmail_url = "http://10.18.2.6/webmail"
```

## Документация Mailu API

См. официальную документацию:
- https://mailu.io/master/api.html

