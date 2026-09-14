# vkr-mpgu 6.33-production — AI-скилл для подготовки ВКР

AI-скилл для ИИ-агентов и ассистентов: подготовка, оформление и проверка выпускной квалификационной работы по требованиям МПГУ, профилю программы и ГОСТ. Скилл не привязан к одной платформе — он работает в любой ИИ-системе, которой можно передать постоянную инструкцию и файлы проекта: агентные CLI (Codex, Claude Code, Gemini CLI, OpenCode), редакторы с ИИ (Cursor, Windsurf, GitHub Copilot), десктоп- и веб-чаты, API и локальные агентные фреймворки. Нормативная инструкция для агента одна и та же везде — [vkr-mpgu/SKILL.md](./vkr-mpgu/SKILL.md).

> **Быстрый выбор:** скачайте один файл поставки — `vkr-mpgu-6.33-production.zip` или `vkr-mpgu-6.33-production.skill` — со страницы [Releases](https://github.com/v2horizon/vkr-checker-skill/releases) (те же файлы лежат рядом с этим README). Устанавливать оба файла не нужно. Инструкция для пользователя — [USER-GUIDE-RU.md](./USER-GUIDE-RU.md).

## Что находится в репозитории

| Путь | Назначение |
|---|---|
| [USER-GUIDE-RU.md](./USER-GUIDE-RU.md) | **Инструкция для пользователя**: что понадобится, установка, первый запуск, как идёт работа, правки в Word, типичные проблемы |
| [vkr-mpgu/SKILL.md](./vkr-mpgu/SKILL.md) | Основная инструкция, которую читает ИИ-агент |
| [vkr-mpgu/PORTABLE-PROMPT.md](./vkr-mpgu/PORTABLE-PROMPT.md) | Переносимая инструкция и набор файлов для веб-систем и API |
| [vkr-mpgu/INSTALL-RU.md](./vkr-mpgu/INSTALL-RU.md) | Подробная инструкция установки и обновления |
| [vkr-mpgu/references/](./vkr-mpgu/references/) | Методика, профили требований, рабочие процессы и протоколы аудита; сквозной путь — `references/quickstart.md` |
| [vkr-mpgu/scripts/](./vkr-mpgu/scripts/) | Инициализация проекта, сборка и импорт DOCX, валидатор, аудит, doctor, память, анализаторы |
| [vkr-mpgu/assets/](./vkr-mpgu/assets/) | Шаблоны форм и документ базовой методички |
| [vkr-mpgu/data/](./vkr-mpgu/data/) | Словарь ИИ-клише для валидатора и детектора (без него проверка клише неполная) |
| [tests/](./tests/) | Регрессионные и контрактные тесты скриптов и документации |
| [AGENTS.md](./AGENTS.md), [CLAUDE.md](./CLAUDE.md), [GEMINI.md](./GEMINI.md) | Точки входа для ИИ-систем, которые читают эти файлы в корне репозитория |
| `.cursor/rules/vkr-mpgu.mdc`, `.github/copilot-instructions.md` | Точки входа для Cursor и GitHub Copilot |
| `AUDIT-FIXES-*-RU.md` | Отчёты о проверке версий: исправления, результаты тестов, SHA-256 поставки |

Сам скилл намеренно остаётся папкой `vkr-mpgu/`: `SKILL.md`, справочники, скрипты и шаблоны должны сохраняться как отдельные файлы, иначе перестанут работать ссылки и запуск утилит.

## Что умеет скилл

- проводит intake-интервью (включая правила кафедры об использовании ИИ) и создаёт отдельный проект ВКР по `intake.json`;
- ведёт текст в Markdown-черновиках `drafts/*.md` и собирает единственный сдаваемый файл `final/vkr.docx` с базовым профилем МПГУ: A4, Times New Roman 14 pt, поля 35/10/20/20 мм, интервал 1,5;
- переносит готовый DOCX в черновики и возвращает в них ручные правки из Word (`import_docx.py`);
- ведёт реестр источников с записью фактической проверки, реестр утверждений, state, roadmap, handoff и журналы;
- проверяет оформление валидатором, текст — анализаторами шаблонности и вопросами комиссии;
- проводит независимый read-only аудит девяти ролей через `vkr_audit.py`: аудиторы только находят ошибки, исправляет основной агент, затем новые аудиторы делают адресную перепроверку и слепую регрессию;
- выдаёт машинный статус готовности (`vkr_project_doctor.py`): `READY_FOR_SUPERVISOR_REVIEW` или `READY_TO_SUBMIT`;
- продолжает работу в новом чате по сохранённой памяти проекта.

Скилл не выдумывает результаты проекта, источники, пилотирование или подписи. Если платформа не умеет создавать независимые контексты, аудит выполняется последовательно, помечается `degraded_independence`, и итог не выше `READY_FOR_SUPERVISOR_REVIEW`.

## Установка для пользователя

Требования: **Python 3.9 или новее** (в macOS/Linux — возможно, `python3`), **`python-docx`** и **`PyYAML`**:

```bash
python -m pip install python-docx pyyaml
```

Автоматическое обновление оглавления работает в Windows с Microsoft Word; без Word оглавление обновляется вручную (F9). В Windows PowerShell 5.1 читайте JSON-отчёты из файлов `-o` или включите UTF-8: `$OutputEncoding = [Console]::OutputEncoding = [Text.UTF8Encoding]::new()`.

### Порядок обновления

1. Закройте задачи и сессии, которые используют старый `vkr-mpgu`, и программы, открывшие файлы из папки скилла.
2. Перенесите старую папку за пределы каталога skills целиком одной операцией (блоки ниже: `[IO.Directory]::Move` в PowerShell, `mv` в bash). Если файл занят, блок останавливается с сообщением «Старая папка не перенесена, ничего не изменено» (в bash перед ним печатается ошибка `mv`): закройте программы и повторите блок; новую версию поверх старой не распаковывайте.
3. Распакуйте новую версию без перезаписи: блоки перед распаковкой проверяют, что старой папки в `skills` больше нет, и поверх неё не распаковывают.
4. Проверьте **версию**, а не только наличие `SKILL.md`: файл содержит `6.33-production`.
5. Для Claude Code и Claude Desktop отключите прежний `vkr-mpgu`, подключённый к аккаунту платформы (в сессии виден как `anthropic-skills:vkr-mpgu`), иначе загрузятся две версии.
6. Начните новую задачу или сессию.

### Вариант A — импорт файла .skill

Если клиент показывает **Import skill** или **Install skill**, отключите или удалите прежний `vkr-mpgu`, выберите `vkr-mpgu-6.33-production.skill` и откройте новую задачу.

```text
Codex: $vkr-mpgu помоги начать ВКР с нуля
Claude Code: /vkr-mpgu помоги начать ВКР с нуля
Claude Desktop и Claude.ai: помоги начать ВКР с нуля
```

### Вариант B — ZIP для Codex и Claude Code

Windows / Codex (одним блоком; при ошибке блок останавливается):

```powershell
& {
  $ErrorActionPreference = 'Stop'
  $zip = "$HOME\Downloads\vkr-mpgu-6.33-production.zip"
  $skills = "$HOME\.codex\skills"   # Claude Code: "$HOME\.claude\skills"
  $target = Join-Path $skills 'vkr-mpgu'
  if (-not (Test-Path -LiteralPath $zip)) { throw "Нет файла поставки: $zip" }
  New-Item -ItemType Directory -Force $skills | Out-Null
  if (Test-Path -LiteralPath $target) {
    $backup = "$HOME\vkr-mpgu-backup-$(Get-Date -Format yyyyMMdd-HHmmss)"
    try {
      [IO.Directory]::Move($target, $backup)
    } catch {
      throw "Старая папка не перенесена, ничего не изменено: $($_.Exception.Message) Закройте программы, которые держат файлы из $target, и повторите блок."
    }
  }
  Expand-Archive -LiteralPath $zip -DestinationPath $skills
  if (Select-String -LiteralPath (Join-Path $target 'SKILL.md') -Pattern '6.33-production' -SimpleMatch -Quiet) {
    'vkr-mpgu 6.33-production установлен'
  } else {
    throw 'В SKILL.md нет версии 6.33-production: установка не завершена'
  }
}
```

Проверка версии:

```powershell
Select-String -LiteralPath "$HOME\.codex\skills\vkr-mpgu\SKILL.md" -Pattern '6.33-production' -SimpleMatch -Quiet
```

Ожидается `True`. Затем начните новую задачу и напишите `$vkr-mpgu помоги начать ВКР с нуля`.

macOS / Linux:

```bash
(
  set -e
  zip="$HOME/Downloads/vkr-mpgu-6.33-production.zip"
  skills="$HOME/.codex/skills"   # Claude Code: "$HOME/.claude/skills"
  [ -f "$zip" ]
  mkdir -p "$skills"
  if [ -e "$skills/vkr-mpgu" ]; then
    mv "$skills/vkr-mpgu" "$HOME/vkr-mpgu-backup-$(date +%Y%m%d-%H%M%S)" || true
  fi
  [ ! -e "$skills/vkr-mpgu" ] || { echo "Старая папка не перенесена, ничего не изменено" >&2; exit 1; }
  unzip -q -n "$zip" -d "$skills"
  grep -q '6.33-production' "$skills/vkr-mpgu/SKILL.md"
  echo "vkr-mpgu 6.33-production установлен"
)
```

Проверка версии (bash):

```bash
grep -q '6.33-production' ~/.codex/skills/vkr-mpgu/SKILL.md && echo ok
```

Ожидается `ok`.

Claude Code: те же блоки, но в строке `$skills = …` (PowerShell) или `skills=…` (bash) замените каталог `.codex` на `.claude`, как подсказывает комментарий в блоке; итоговый путь `~/.claude/skills/vkr-mpgu/SKILL.md`, проверка — `grep -q '6.33-production' ~/.claude/skills/vkr-mpgu/SKILL.md && echo ok`. Затем новая сессия и `/vkr-mpgu`. Подробно — [vkr-mpgu/INSTALL-RU.md](./vkr-mpgu/INSTALL-RU.md).

### Вариант C — из клонированного репозитория

```bash
git clone <URL-ЭТОГО-РЕПОЗИТОРИЯ>
cd <КАТАЛОГ-РЕПОЗИТОРИЯ>
```

Скопируйте каталог `vkr-mpgu/` в каталог навыков вашей системы (перед этим перенесите старую версию за пределы каталога skills). В Codex итоговый путь `~/.codex/skills/vkr-mpgu/SKILL.md`, в Claude Code — `~/.claude/skills/vkr-mpgu/SKILL.md`. Не копируйте только README: он объясняет установку, но не заменяет файлы скилла.

## Установка для веб-систем и API

Если платформа не импортирует навыки и не поддерживает локальную папку:

1. Передайте `vkr-mpgu/PORTABLE-PROMPT.md` как постоянную инструкцию проекта.
2. Загрузите файлы знаний по таблице `PORTABLE-PROMPT.md` → «Файлы для веб-систем»: в обязательный набор входят `SKILL.md`, всё, что `SKILL.md` велит прочитать в начале работы, справочники рабочего пути и `data/ai_cliches.json`.
3. Каталоги `scripts/` и `assets/` предоставьте рабочей директории, где агент может запускать Python.
4. Убедитесь, что агент действительно умеет читать и писать файлы, запускать Python и пользоваться веб-поиском. Если инструмента нет, агент готовит команду или список действий и не сообщает о невыполненной проверке как о выполненной.

Полный ZIP предпочтительнее, потому что сохраняет все пути и шаблоны. Отличия платформ — `vkr-mpgu/references/platform-integration.md`.

## Использование в других ИИ-системах

Скилл — это обычные файлы: нормативная инструкция [vkr-mpgu/SKILL.md](./vkr-mpgu/SKILL.md), справочники, скрипты и шаблоны. Подключить их можно к любой системе, которая умеет держать постоянную инструкцию и читать файлы.

| Платформа | Как подключить |
|---|---|
| Claude Code, Claude Desktop, Claude.ai | импорт файла `.skill` или папка `~/.claude/skills/vkr-mpgu`; вызов `/vkr-mpgu …` или обычный запрос о ВКР |
| Codex | папка `~/.codex/skills/vkr-mpgu`; вызов `$vkr-mpgu …` |
| Cursor | правило `.cursor/rules/vkr-mpgu.mdc` (клонируйте репозиторий рядом с проектом ВКР или скопируйте правило в свой проект) |
| Windsurf, OpenCode и другие агенты с `AGENTS.md` | файл [AGENTS.md](./AGENTS.md) в корне клона |
| GitHub Copilot | файл `.github/copilot-instructions.md` в репозитории (плюс `AGENTS.md`) |
| Gemini CLI | файл [GEMINI.md](./GEMINI.md) в корне клона |
| ChatGPT, Gemini, любой другой веб-чат или API | [vkr-mpgu/PORTABLE-PROMPT.md](./vkr-mpgu/PORTABLE-PROMPT.md) как постоянная инструкция проекта, файлы — через хранилище знаний или вложения (таблица «Файлы для веб-систем») |
| LangChain, CrewAI, AutoGen и подобные фреймворки | `vkr-mpgu/SKILL.md` как системная инструкция; агенту нужны инструменты чтения и записи файлов, запуска Python и веб-поиска |

Во всех случаях нормативной остаётся `vkr-mpgu/SKILL.md`, а файлы точек входа только направляют к ней и повторяют ключевые правила. Рабочие файлы ВКР храните в отдельной папке проекта, а не в клоне скилла.

**Про независимый аудит.** Роли аудиторов должны выполняться в отдельных контекстах (субагентах). Если платформа их не создаёт, аудит выполняется последовательно в том же контексте, помечается `degraded_independence`, и готовность не поднимается выше `READY_FOR_SUPERVISOR_REVIEW`.

## Инструкция для ИИ-агента

Блок `AI_AGENT_BOOTSTRAP` ниже — единственный источник правил запуска: файлы точек входа (`AGENTS.md`, `CLAUDE.md`, `GEMINI.md`, `.cursor/rules/vkr-mpgu.mdc`, `.github/copilot-instructions.md`) ссылаются на него и не хранят собственных копий. Блок можно передать агенту вместе с репозиторием:

```text
AI_AGENT_BOOTSTRAP

1. Прочитай vkr-mpgu/SKILL.md до начала работы; сквозной путь с командами —
   vkr-mpgu/references/quickstart.md.
2. Считай README описанием установки, а SKILL.md — нормативной инструкцией.
3. Не редактируй установленную папку скилла во время работы пользователя.
   Рабочие файлы ВКР храни в отдельной папке проекта <PROJECT_DIR>.
4. Если проект уже содержит vkr-state.md, прочитай state, memory/handoff.json,
   memory/roadmap.md, memory/claims-register.json, audit/manifest.json и
   audit/findings.json, затем запусти vkr_memory.py status и vkr_audit.py next.
5. Если проект новый, проведи intake (включая правила кафедры об ИИ), запиши
   intake.json и запусти init_vkr_project.py. Пользователь не создаёт структуру вручную.
6. Текст пиши в drafts/*.md, DOCX собирай build_vkr.py; сдаётся только final/vkr.docx.
7. Проверяй источники фактически доступным веб-поиском и записывай итог
   verify_sources.py --mark. Не создавай фиктивные URL, результаты, пилоты или подписи.
8. Аудит — только командами vkr_audit.py: аудиторы read-only, исправляет основной
   агент, служебные поля manifest и отчётов вручную не пишутся.
9. После любого исправления: пересборка, новый снимок, адресная перепроверка и
   новые волны.
10. Перед заявлением о готовности запусти:
    python <SKILL_DIR>/scripts/vkr_project_doctor.py <PROJECT_DIR> --stage final --json
    READY_TO_SUBMIT выдаёт только doctor; FAIL означает, что работа не готова.
11. Не обещай абсолютную безошибочность: финальная проверка кафедры и научного
    руководителя остаётся обязательной.
AI_AGENT_BOOTSTRAP_END
```

## Первый запуск: ВКР с нуля

1. Установите скилл и начните новую задачу.
2. Напишите запрос (синтаксис по платформам — выше).
3. Ответьте на intake: программа, тема, срок, тип ВКР, продукт, источники, фактические результаты, правила кафедры об ИИ.
4. Основной агент создаст проект и заполнит служебные файлы.
5. Работайте по выбранному режиму: стандартный или экспресс.
6. После глав, перед передачей научному руководителю и перед сдачей дождитесь независимого аудита и исправлений основным агентом.

Пользователь не создаёт вручную каталоги `drafts`, `sources`, `evidence`, `audit`, `memory`, `logs`, `final` и `exports`.

## Сквозной путь и полезные команды

`<SKILL_DIR>` — абсолютный путь к установленной папке `vkr-mpgu`; `<PROJECT_DIR>` — отдельная папка проекта. Полный порядок с ожидаемыми результатами — [vkr-mpgu/references/quickstart.md](./vkr-mpgu/references/quickstart.md).

```bash
python <SKILL_DIR>/scripts/init_vkr_project.py <PROJECT_DIR> --config <PROJECT_DIR>/intake.json --json
python <SKILL_DIR>/scripts/vkr_audit.py <PROJECT_DIR> plan --kind primary --preset plan --json
python <SKILL_DIR>/scripts/verify_sources.py --report <PROJECT_DIR>/sources.json
python <SKILL_DIR>/scripts/build_vkr.py <PROJECT_DIR>
python <SKILL_DIR>/scripts/update_docx_fields.py <PROJECT_DIR>/final/vkr.docx
python <SKILL_DIR>/scripts/clean_docx_metadata.py <PROJECT_DIR>/final/vkr.docx --in-place
python <SKILL_DIR>/scripts/vkr_audit.py <PROJECT_DIR> validate --json
python <SKILL_DIR>/scripts/ai_detection_heuristic.py <PROJECT_DIR>/final/vkr.docx -o <PROJECT_DIR>/exports/ai-detection.json
python <SKILL_DIR>/scripts/content_ownership_check.py <PROJECT_DIR>/final/vkr.docx --all -o <PROJECT_DIR>/exports/ownership-questions.json
python <SKILL_DIR>/scripts/vkr_project_doctor.py <PROJECT_DIR> --stage prefinal --json
python <SKILL_DIR>/scripts/vkr_project_doctor.py <PROJECT_DIR> --stage final --json
python <SKILL_DIR>/scripts/import_docx.py <PROJECT_DIR> --docx <PROJECT_DIR>/final/vkr.docx --update
```

Единственный сдаваемый файл — `final/vkr.docx`; промежуточные DOCX и отчёты — в `exports/`. Оглавление обновляется и метаданные очищаются до `validate` и снимка аудита.

## Продолжение работы в новом чате

Откройте ту же папку проекта или передайте агенту ZIP проекта. В системе, которая принимает только вложения, передайте файлы из `continuation_files` вывода `vkr_memory.py status`:

```text
vkr-state.md
vkr-project.json
plan.md
sources.json
evidence/index.json
memory/handoff.json
memory/handoff.md
memory/roadmap.md
memory/claims-register.json
memory/artifact-index.json
audit/manifest.json
audit/findings.json
audit/snapshot-inputs.json
drafts/*.md
```

Агент сначала запускает `vkr_memory.py <PROJECT_DIR> status --json` и `vkr_audit.py <PROJECT_DIR> next --stage …` и только потом продолжает работу. Если файл изменился вне агента, прежние проверки к новому снимку не относятся.

### Обновление скилла при незавершённом проекте

После установки новой версии агент проверяет `status` (ожидаемо `needs_rebaseline`) и doctor на стадии `draft`, затем переиндексирует файлы:

```bash
python <SKILL_DIR>/scripts/vkr_memory.py <PROJECT_DIR> record --rebaseline --json
```

Записи аудита, сделанные вручную по протоколу до 6.33, doctor показывает как `AUDIT_LEGACY_RUNS` и не засчитывает: нужные проверки проводятся заново через `vkr_audit.py`.

## Структура проекта ВКР

```text
<PROJECT_DIR>/
├── vkr-project.json      конфигурация: профиль, тип, режим, интенсивность аудита, титул
├── vkr-state.md          читаемое состояние проекта
├── plan.md
├── sources.json          реестр источников
├── drafts/               annotation.md, introduction.md, chapter-N.md, conclusion.md, appendix-N.md
├── evidence/             index.json и доказательства: methodology, product, pilot, figures, defense, approvals
├── sources/materials/    оригиналы материалов; для mpgu-* — копия методички
├── audit/                снимки, задания, отчёты, реестр замечаний (пишет vkr_audit.py)
├── memory/               handoff, roadmap, решения, реестр утверждений, индекс файлов
├── logs/                 журналы checkpoint и запусков
├── exports/              промежуточные сборки и отчёты анализаторов
├── final/                vkr.docx — единственный сдаваемый файл
└── backups/checkpoints/  резервные копии перед сборкой, импортом и обновлением полей
```

## Проверенная поставка и тесты

- Версия: `6.33-production`.
- SHA-256 файлов поставки и результаты тестов записаны в отчёте о проверке версии (`AUDIT-FIXES-*-RU.md`).

Разработчик может повторить проверку из корня репозитория (Python 3.9+ с `python-docx`):

```bash
python -m unittest discover -s tests -t .
```

Отдельный файл тестов тоже запускается напрямую, например `python tests/test_docs_contract_633.py`. Тесты не ходят в сеть и не пишут в каталог скилла; тесты, которым нужен Word, используют закоммиченные фикстуры из `tests/fixtures/`.

## Ограничения и честное использование

Скилл помогает подготовить текст, оформление и проверки, но не заменяет студента, научного руководителя или требования кафедры. Пользователь обязан понимать работу, подтвердить фактические результаты, проверить источники и соблюдать правила кафедры об использовании ИИ. Эвристика AI-текста не является гарантией прохождения конкретного детектора. Реальная готовность ВКР определяется фактическими данными и прохождением требований вашей программы.

## Если что-то не работает

1. Проверьте, что `SKILL.md` находится непосредственно в `.../skills/vkr-mpgu/SKILL.md`, без вложенности `vkr-mpgu/vkr-mpgu`.
2. Проверьте версию: `SKILL.md` содержит `6.33-production` (команды выше).
3. Убедитесь, что не загружены две версии (папка в `skills` и навык, подключённый в аккаунте платформы).
4. Начните новую задачу после установки или обновления скилла.
5. Если платформа не поддерживает `.skill`, используйте ZIP.
6. Если независимые агенты недоступны, аудит выполняется последовательно с `degraded_independence`.
7. Для подробной диагностики прочитайте [vkr-mpgu/INSTALL-RU.md](./vkr-mpgu/INSTALL-RU.md) и `vkr-mpgu/references/failure-recovery.md`.

## Лицензия и исходная методика

Перед публикацией репозитория добавьте лицензию, разрешённую автором исходной методики, и сохраните сведения об авторстве. Документ методички в `assets/` не является частью лицензии кода скилла; не удаляйте официальные методические материалы из `assets/` и не выдавайте примеры скилла за фактические результаты пользователя.
