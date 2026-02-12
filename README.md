# Teacher Blog (Word -> Jekyll)

Проект конвертирует контент из `content/content.docx` в многостраничный сайт Jekyll:
- главная страница,
- список новостей с пагинацией (по 10 на страницу),
- отдельные страницы новостей,
- встроенные изображения из Word.

## Требования

- Ruby + Bundler
- Jekyll (через `bundle`)
- Python 3

## Установка

В корне проекта выполните:

```powershell
bundle install
```

## Основной сценарий работы

1. Обновите файл `content/content.docx`.
2. Сгенерируйте сайт из Word:

```powershell
python scripts\docx_to_jekyll.py
```

3. Запустите локальный сервер Jekyll:

```powershell
bundle exec jekyll serve --livereload
```

Сайт будет доступен по адресу `http://127.0.0.1:4000/`.

## Что генерирует скрипт

После запуска `scripts/docx_to_jekyll.py` обновляются/создаются:
- `index.md`
- `news/index.md` и дополнительные страницы `news/page/<n>/index.md` (если новостей больше 10)
- файлы новостей в `_news/`
- изображения новостей в `assets/images/news/`
- служебные файлы и шаблоны Jekyll (если их нет)

## Полезные команды

Сборка без запуска сервера:

```powershell
bundle exec jekyll build
```

Очистка и пересборка:

```powershell
bundle exec jekyll clean
python scripts\docx_to_jekyll.py
bundle exec jekyll build
```
