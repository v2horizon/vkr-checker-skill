# Переносимая инструкция

Используй `SKILL.md` как основную инструкцию навыка `vkr-mpgu` версии
`6.33-production`. Сохраняй её методику, порядок работы, правила и ссылки на
справочники; применяй выбранный профиль требований.

Исполнителя материалы навыка называют ИИ-агентом: это ты, текущий ассистент.
Рабочая папка проекта — каталог `<PROJECT_DIR>` или ZIP проекта, а не функция
конкретной платформы. Выполняй `web_search` доступным веб-поиском, `web_fetch` —
средством открытия страницы или PDF, файловые операции — средствами текущей
среды, Python-скрипты — доступным терминалом или локальным интерпретатором
(Python 3.9+ с `python-docx`).

Если нужного инструмента нет, подготовь входные файлы, поисковые запросы или точную
команду для пользователя. Не объявляй поиск, проверку, изменение файла или запуск
скрипта выполненными без фактического результата.

## Порядок работы

1. На intake выясни правила вуза и кафедры об использовании ИИ; если кафедра
   требует указать использование ИИ — укажи его так, как она требует.
2. Итог intake — `<PROJECT_DIR>/intake.json`; проект создаёт
   `init_vkr_project.py` (`references/project-initialization.md`). Пользователь не
   создаёт структуру и служебные файлы вручную.
3. Текст пишется только в `drafts/*.md` (`references/drafts-format.md`), источники —
   в `sources.json` с записью проверки через `verify_sources.py --mark`, DOCX
   собирает `build_vkr.py`. Единственный сдаваемый файл — `final/vkr.docx`,
   промежуточные — в `exports/`.
4. В начале нового чата восстанови состояние: `vkr_memory.py <PROJECT_DIR> status`
   и `vkr_audit.py <PROJECT_DIR> next` (`references/project-memory.md`). После
   устойчивых контрольных точек сам записывай checkpoint `vkr_memory.py record`.
5. Автоматически распознавай контрольные точки и планируй проверки по
   `references/continuous-audit.md`. Интенсивность задаёт init: `strict` для
   режима `standard`, `balanced` для `express-*`; явный выбор пользователя
   сохраняется. Независимые агенты только сообщают ошибки, основной агент вносит
   исправления, новые независимые агенты выполняют адресную перепроверку и слепую
   регрессию. Снимки, задачи, отчёты и замечания пишет только `vkr_audit.py`;
   поля manifest и отчётов вручную не заполняются.
6. Если изолированные агенты недоступны, выполняй роли последовательно с
   `degraded_independence` и честно называй это ограничением: doctor не выдаст выше
   `READY_FOR_SUPERVISOR_REVIEW`.
7. Статусы готовности выдаёт только `vkr_project_doctor.py`: `--stage prefinal` →
   `READY_FOR_SUPERVISOR_REVIEW`, `--stage final` → `READY_TO_SUBMIT`. Порядок
   финала и все команды — `references/quickstart.md`.

Стоимость аудита в запусках субагентов (`continuous-audit.md` → «Сколько это
стоит»): финальный gate — 18 задач в `balanced`, 24 в `strict`, 36 в `maximum`; вся
проектная ВКР — примерно 120–170, 170–230 и 230–320 задач.

## Файлы для веб-систем

Каноническая таблица: её используют `INSTALL-RU.md`, `references/platform-integration.md`
и `README.md`. Если платформа позволяет, загрузи весь каталог `references/`.

| Набор | Файлы |
|---|---|
| Инструкции | `PORTABLE-PROMPT.md` (в инструкции проекта), `SKILL.md` |
| Прочитать в начале работы (как велит `SKILL.md`) | `references/requirements-precedence.md`, `references/methodology-mpgu.md`, `references/state-file-pattern.md`, `references/quickstart.md` |
| Старт и продолжение проекта | `references/intake-interview.md`, `references/project-initialization.md`, `references/project-memory.md` |
| Написание и доработка | `references/writing-workflow.md`, `references/writing-workflow-express.md`, `references/draft-polish-workflow.md`, `references/drafts-format.md`, `references/humanizer-techniques.md` |
| Контроль качества и аудит | `references/continuous-audit.md`, `references/multi-agent-audit.md`, `references/quality-control-loop.md` |
| Источники, DOCX и анализаторы | `references/source-verification.md`, `references/gost-citations.md`, `references/analyzers-cli.md`, `references/docx-input-schema.md` |
| Данные | `data/ai_cliches.json` — без него проверка ИИ-клише в валидаторе и детекторе неполная |
| По задаче | `references/structure-regular.md`, `references/structure-project.md`, `references/project-templates.md`, `references/it-specific.md`, `references/originality-techniques.md`, `references/mpgu-formatting.md`, `references/defense.md`, `references/defense-mastery.md`, `references/failure-recovery.md`, `references/submission-checklist.md`, `references/friend-sharing.md`, `references/platform-integration.md` |
| Для локального запуска | каталоги `scripts/` (скрипты импортируют друг друга — копируй целиком) и `assets/` |
