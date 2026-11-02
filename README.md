# Работа для зумеров — «Норм или стрём?»

Статический Telegram Mini App с 20 карточками-вопросами и свайпами: вправо — «норм», влево — «стрём».

## Файлы проекта

- `index.html` — интерфейс карточек и отправка открытия/ответов.
- `transport-banner.svg` — пример рекламного баннера в карточке Московского транспорта.
- `Code.gs` — Apps Script: создаёт таблицу, принимает события и проверяет подпись Telegram `initData`.
- `PROJECT_GUIDE.md` — подробные шаги настройки, публикации и дальнейших обновлений.

## Google Таблица и Apps Script

Проект Apps Script **«команда транспорта»** обновлён. Функция `setup()` уже выполнена; она создала таблицу с вкладками `Sessions`, `Responses` и `Summary`:

[Открыть таблицу результатов](https://docs.google.com/spreadsheets/d/1t_M-pIP0cMEWS3zqJQZqX3vlxrYa6JopaVe_o41aCAU/edit)

- `Sessions` записывает дату открытия, Telegram ID и ID сессии.
- `Responses` записывает каждый ответ вместе с Telegram ID и конкретным ID/текстом вопроса.
- `Summary` автоматически считает «Норм» и «Стрём» по каждому вопросу формулами `COUNTIFS`.

Apps Script развернут в двух режимах: приватный endpoint **«Только я»** сохранён, а отдельный endpoint для Mini App доступен извне и исполняется от имени владельца. Он принимает запись только с корректной подписью Telegram `initData` не старше 24 часов. [Публичный `/exec`](https://script.google.com/macros/s/AKfycbwRe5CmVRT-AsHEB-d2IFlAOz8arWwBxQCrg8JQGlE79Af0_81D8OdqdAXYTSXX6tOD/exec) уже подключён в `APPS_SCRIPT_URL`. Токен хранится в свойствах Apps Script, не в репозитории.

## GitHub и Vercel

GitHub-репозиторий проекта: [mostransport-team](https://github.com/ilushkakush-coder/mostransport-team). После загрузки исходников сюда он станет источником версий кода и документации. Когда репозиторий будет связан с Vercel, каждый push в основную ветку будет автоматически обновлять сайт.

Production-сайт: [https://mostransport-team.vercel.app](https://mostransport-team.vercel.app). Vercel подключён к этому репозиторию и автоматически обновит сайт после push в `main`. Укажите эту ссылку как URL Telegram Mini App, затем протестируйте открытие и сохранение свайпов. Инструкции находятся в `PROJECT_GUIDE.md`.
