# Автоматическая инициализация проекта ВКР

Основной агент создаёт рабочую структуру сам. Пользователь не должен вручную
создавать каталоги, пустые файлы, manifest или JSON-заготовки.

## Когда запускать

Запускай инициализацию, если пользователь начинает ВКР с нуля и в выбранной
папке нет `vkr-state.md`. Если пользователь приносит готовый DOCX, после
инициализации перенеси его в черновики через `import_docx.py`: дальнейший путь
общий. Для точечной проверки одного файла проект не нужен. Если state уже есть,
прочитай его и продолжай без повторной инициализации.

## Откуда запускать команды

Все скрипты вызываются по абсолютному пути установленного скилла, текущий
каталог значения не имеет:

```bash
python <SKILL_DIR>/scripts/init_vkr_project.py <PROJECT_DIR> --config <PROJECT_DIR>-intake.json --json
```

`<SKILL_DIR>` — папка, где лежит `SKILL.md`; `<PROJECT_DIR>` — отдельная папка
пользователя. Файл intake годится в любом из двух мест: рядом с проектом
(`<PROJECT_DIR>-intake.json`) или внутри ещё пустой папки проекта
(`<PROJECT_DIR>/intake.json`, так в `quickstart.md` и `intake-interview.md`);
во втором случае init учитывает его в базовом индексе памяти. Относительный `<PROJECT_DIR>` разрешается от текущего каталога, а
в выводе `project_root` всегда абсолютный.

В Windows PowerShell 5.1 с кодовой страницей 866/1251 перехваченный вывод
искажает кириллические пути. Используй `-o <файл.json>` (JSON в UTF-8 без BOM)
или перед запуском `$OutputEncoding = [Console]::OutputEncoding = [Text.UTF8Encoding]::new()`.
Ключ `-o` есть у `init_vkr_project.py`, `vkr_memory.py`, `vkr_project_doctor.py`
и `vkr_audit.py`.

## Выбор каталога

1. Если пользователь открыл отдельную папку именно для ВКР, используй её.
2. Если текущая папка содержит посторонний проект или материалы, создай рядом
   дочерний каталог `vkr-<краткий-slug-темы-или-фамилии>`.
3. Если данных для имени нет, используй `vkr-project`; при коллизии добавь
   суффикс `-2`, `-3` и далее.
4. Для разных студентов — разные корневые каталоги.

Проект нельзя создать внутри установленной папки скилла: скрипт отказывает с
кодом 2, в том числе для относительного пути, запущенного из папки скилла.

## Порядок

1. Проведи intake по `intake-interview.md`.
2. Выбери профиль требований, тип ВКР, режим и `audit_intensity`.
3. Запиши ответы в JSON-файл вне будущего проекта (или рядом с ним) и запусти
   `init_vkr_project.py` с `--config`. Для ФИО и темы предпочитай JSON, чтобы
   не потерять кавычки.
4. Проверь JSON-результат: `status=initialized`, `created`, `warnings`.
5. Дополни `vkr-state.md` ответами intake по схеме `state-file-pattern.md`.
6. Запусти `vkr_project_doctor.py <PROJECT_DIR> --stage draft --json` — свежий
   проект проходит `draft` (предупреждения о незаполненном паспорте и справка о
   маркерах в заготовках нормальны). В паспорте state поля титула, которые
   сборка ставит на титульный лист (автор, институт, кафедра, направление,
   научный руководитель), без значения записаны как «не указано» (их ищет
   doctor); вуз (сборка подставляет МПГУ) и правила кафедры об использовании
   ИИ — как «—».
7. Сформируй `plan.md`, запусти аудит точки intake/plan (`continuous-audit.md`)
   и покажи пользователю краткое резюме, путь проекта и следующий шаг.

После успешной инициализации сразу продолжай работу. Не заканчивай ответ
инструкцией «создайте файлы».

## Конфигурация intake

```json
{
  "topic": "Информационная система сопровождения практики",
  "vkr_type": "project",
  "profile": "mpgu-09-project",
  "mode": "standard",
  "audit_intensity": "strict",
  "deadline": "2027-06-01",
  "collective": false,
  "members": [],
  "ai_rules": "…",
  "title_page": {
    "university": "Московский педагогический государственный университет",
    "institute": "Институт математики и информатики",
    "department": "Кафедра …",
    "program_code": "09.03.02",
    "program_name": "Информационные системы и технологии",
    "program_profile": "…",
    "work_type": "Выпускная квалификационная работа",
    "author": "Иванов Иван Иванович",
    "group": "…",
    "supervisor": "Петрова А. Б.",
    "supervisor_title": "канд. пед. наук, доцент",
    "head_of_department": "…",
    "head_title": "…",
    "city": "Москва",
    "year": "2027"
  }
}
```

Правила:

- `profile`: `generic`, `mpgu-09-project` или `mpgu-09-regular` (алиас ключа —
  `validation_profile`). `mpgu-09-project` требует `vkr_type=project`,
  `mpgu-09-regular` — `regular`; противоречие — ошибка с кодом 2. Если
  `vkr_type` не указан, он берётся из профиля МПГУ, для `generic` — `project`.
- `mode`: `standard`, `express-14d`, `express-7d`, `express-4d`.
- `audit_intensity`: `balanced`, `strict`, `maximum`. Без явного значения
  `standard` получает `strict`, `express-*` — `balanced`.
- `collective` — только JSON `true` или `false`; `members` — массив строк.
- `deadline` — `YYYY-MM-DD` или пустая строка.
- `ai_rules` — строка: правила вуза и кафедры об использовании ИИ (запрет,
  допустимые виды помощи, нужно ли и где указывать использование ИИ). Init
  переносит её в `vkr-state.md` → «Требования и структура» строкой «Правила
  кафедры об использовании ИИ»; пустая строка или отсутствие ключа — «—». Поле
  необязательное: doctor его не требует. Если правила неизвестны, пиши «уточнить
  у научного руководителя», а не «не указано» (так doctor помечает пустой паспорт).
- `title_page` — только ключи из примера; по умолчанию `work_type`
  «Выпускная квалификационная работа», `city` «Москва», `year` — текущий год.
  Обязательны `author` и поля, которые сборка при пустом значении заменяет
  маркером `[ЗАПОЛНИТЬ: …]`: `institute`, `program_code`, `program_name`,
  `program_profile`, `supervisor_title`, `supervisor`, `department`,
  `head_title`, `head_of_department` (doctor `TITLE_PAGE_INCOMPLETE`).
- Прежние плоские ключи `student_name`, `institution`, `institute`,
  `program_code`, `program_name`, `supervisor`, `department`, `group` читаются
  как поля `title_page`.
- Неизвестные ключи — ошибка с перечнем допустимых, чтобы опечатка не
  подменилась значением по умолчанию.

CLI-параметры `--topic`, `--vkr-type`, `--profile`, `--mode`,
`--audit-intensity`, `--deadline`, `--collective`, `--student-name`,
`--institution`, `--program-code` дополняют конфигурацию.

## Что создаётся

```text
<PROJECT_DIR>/
├── vkr-project.json            # каноничная конфигурация (профиль, тип, режим, титульный лист)
├── vkr-state.md
├── plan.md
├── sources.json
├── drafts/                     # annotation, introduction, chapter-1…3, conclusion
├── sources/materials/          # для mpgu-*: копия methodology-09-03-02-2024.docx
├── evidence/
│   ├── index.json              # {methodology, sources, product, pilot, figures, defense, approvals}
│   ├── methodology/ product/ pilot/ figures/ defense/ approvals/
├── audit/
│   ├── manifest.json  snapshot-inputs.json  findings.json
│   ├── briefs/  reports/  snapshots/
├── memory/                     # handoff.json, handoff.md, roadmap.md, decisions.jsonl,
│                               # claims-register.json, artifact-index.json
├── logs/                       # activity.jsonl, tool-runs.jsonl
├── final/                      # здесь будет единственный сдаваемый vkr.docx
├── exports/                    # промежуточные сборки и отчёты анализаторов
└── backups/checkpoints/
```

Для профилей `mpgu-*` init копирует методичку в
`sources/materials/methodology-09-03-02-2024.docx` и записывает в
`evidence/index.json` → `methodology` запись
`{"id": "methodology-09-03-02-2024", "path": "sources/materials/methodology-09-03-02-2024.docx"}`.

`drafts/chapter-3.md` создаётся для `vkr_type=project`. Doctor требует третью
главу только для профиля `mpgu-09-project`; при другом утверждённом числе глав
основной агент добавляет или удаляет пустые заготовки до начала написания и
фиксирует решение в state. Индекс памяти сразу содержит хеши созданных файлов,
поэтому `vkr_memory.py status` после init возвращает `ok`.

## Защита существующих материалов

`init_vkr_project.py` идемпотентен:

- создаёт недостающие каталоги и файлы и пропускает существующие
  (`skipped_existing`);
- не имеет режима перезаписи;
- если `vkr-project.json` уже есть, берёт из него профиль, тип, режим,
  интенсивность и `collective`, которые не указаны явно; явное расхождение —
  `status=conflict` с перечнем полей и код 2, ничего не записывается;
- если есть только `vkr-state.md` прежней версии, те же значения читаются из
  state (`adopted` в выводе), чтобы новый `vkr-project.json` не противоречил
  проекту;
- если `evidence/index.json` уже был без записи методички, в `warnings`
  выводится запись, которую нужно добавить.

Коды возврата: 0 — `initialized` или `dry_run`; 2 — ошибка конфигурации,
конфликт или ошибка ввода-вывода (`{"status": "error", "error": …}`).

Если папка уже содержит черновики или источники, сначала проинвентаризируй их,
затем запусти инициализатор и перенеси сведения в state, `sources.json` и
`drafts/`; готовый DOCX — через `import_docx.py`. После обновления скилла на
существующем проекте выполни `vkr_memory.py <PROJECT_DIR> record --rebaseline`.

## Среда без доступа к файлам

Если платформа может создавать вложения, основной агент формирует тот же проект
как ZIP и отдаёт пользователю одним файлом. Если нельзя, сохрани state и план в
доступной памяти, подготовь bootstrap-блок и честно отметь
`filesystem_setup_pending: true`. Не заявляй, что структура создана на диске.

Независимые субагенты не выполняют инициализацию и не меняют рабочие файлы.
