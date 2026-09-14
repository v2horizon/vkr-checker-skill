# Установка vkr-mpgu 6.33-production

В поставке находятся два файла с одинаковым содержимым:

- `vkr-mpgu-6.33-production.skill` — контейнер для клиента, который явно
  поддерживает импорт файлов `.skill`;
- `vkr-mpgu-6.33-production.zip` — универсальный ZIP для ручной установки,
  агентных CLI (Codex, Claude Code, Gemini CLI), редакторов с ИИ, веб-систем и API.

**Выбери один вариант. Устанавливать оба файла не нужно.** Если в интерфейсе нет
команды импорта `.skill`, используй ZIP.

## Требования

- **Python 3.9 или новее.** В macOS и Linux команда может называться `python3`.
- **`python-docx`** — сборка, импорт и проверка DOCX. `--help` скриптов работает и без
  него, но основной путь без `python-docx` невозможен.
- **`PyYAML`** — чтение реестра источников в YAML; ставится вместе с `python-docx`.

```bash
python -m pip install python-docx pyyaml
```

- Для автоматического обновления оглавления (`update_docx_fields.py`) — Windows и
  Microsoft Word; без них оглавление обновляется вручную (F9 в Word).
- В Windows PowerShell 5.1 кириллица в перехваченном выводе искажается: читай
  JSON-отчёты из файлов `-o` или перед запуском выполни
  `$OutputEncoding = [Console]::OutputEncoding = [Text.UTF8Encoding]::new()`.

## Обновление: общий порядок

1. **Закрой задачи и сессии**, которые используют старый `vkr-mpgu` (Codex, Claude
   Code, Claude Desktop), а также программы, открывшие файлы из папки скилла
   (Word, Проводник, редактор).
2. **Перенеси старую папку за пределы каталога skills** (например, в домашнюю
   папку с датой в имени). Внутри `skills` не должно остаться второй копии.
3. **Остановись при ошибке переноса.** Блоки ниже переносят папку целиком одной
   операцией (PowerShell — `[IO.Directory]::Move`, bash — `mv`): на одном диске она
   переносится вся или не переносится вовсе. Если файл занят (открыт в Word,
   Проводнике, редакторе или используется сессией), блок останавливается с
   сообщением «Старая папка не перенесена, ничего не изменено» (в bash перед ним
   печатается ошибка `mv`): старая папка остаётся на месте целиком, резервная не
   создаётся. Закрой программы и повтори блок.
4. **Распакуй новую версию** без перезаписи. Блоки перед распаковкой проверяют, что
   старой папки в `skills` больше нет: PowerShell при её наличии даёт ошибку «файл уже
   существует», bash-блок останавливается с тем же сообщением, что и при переносе
   (`unzip -n` сам по себе существующие файлы молча пропускает, поэтому проверка нужна).
   В обоих случаях вернись к шагу 3.
5. **Проверь версию, а не только наличие `SKILL.md`**: файл должен содержать
   `6.33-production`.
6. Начни **новую** задачу или сессию.

## Codex: установка ZIP

### Windows (PowerShell)

Команды выполняются одним блоком: при любой ошибке блок останавливается.

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

Отдельная проверка версии в любой момент:

```powershell
Select-String -LiteralPath "$HOME\.codex\skills\vkr-mpgu\SKILL.md" -Pattern '6.33-production' -SimpleMatch -Quiet
```

Ожидается `True`. Затем открой новую задачу Codex и напиши
`$vkr-mpgu помоги начать ВКР с нуля`.

### macOS и Linux

Команды выполняются в подоболочке с `set -e`: при ошибке переноса или распаковки
она останавливается, не закрывая терминал.

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

Если сообщения «установлен» нет — установка не завершена. Отдельная проверка версии
в любой момент:

```bash
grep -q '6.33-production' ~/.codex/skills/vkr-mpgu/SKILL.md && echo ok
```

Ожидается `ok`. Затем открой новую задачу и вызови `$vkr-mpgu`.

## Claude Code и Claude Desktop

1. **Отключи прежний `vkr-mpgu`, подключённый к аккаунту платформы.** Навык,
   добавленный в аккаунт раньше (например, импортом `.skill`), синхронизируется в
   Claude Code и Claude Desktop и виден в сессии как `anthropic-skills:vkr-mpgu`.
   Отключи или удали его в разделе навыков (Skills) настроек аккаунта платформы, иначе
   одновременно загрузятся две версии с одинаковыми триггерами.
2. **Локальная установка для Claude Code** — те же команды, что для Codex, с
   каталогом `~/.claude/skills` вместо `~/.codex/skills` (в PowerShell —
   `"$HOME\.claude\skills"`). Итоговый путь: `~/.claude/skills/vkr-mpgu/SKILL.md`;
   старая папка переносится за пределы `~/.claude/skills`.
3. Проверка версии: `grep -q '6.33-production' ~/.claude/skills/vkr-mpgu/SKILL.md && echo ok`
   (в PowerShell — `Select-String` как выше, с путём `.claude`).
4. Новая сессия; вызов — `/vkr-mpgu помоги начать ВКР с нуля`. Убедись, что в
   списке навыков сессии нет второго `vkr-mpgu`, и спроси «какая версия скилла
   vkr-mpgu загружена?» — ожидается `6.33-production`.

Для независимого аудита Claude Code запускает субагентов (Task) в отдельных
контекстах (`references/platform-integration.md`).

## Импорт `.skill`

Если клиент показывает команду **Import skill**, **Install skill** или принимает
файл `.skill`: сначала отключи или удали прежний `vkr-mpgu`, затем выбери
`vkr-mpgu-6.33-production.skill` и открой новую задачу. Вызов: в Codex —
`$vkr-mpgu …`; в Claude Code — `/vkr-mpgu …`; в Claude Desktop и веб-интерфейсе —
обычный запрос о ВКР («помоги начать ВКР с нуля»).

Если такой команды нет или импорт завершился ошибкой, не переименовывай внутренние
файлы: используй ручную установку ZIP выше. Сам `.skill` также является
ZIP-контейнером с корневой папкой `vkr-mpgu`; для распаковки средствами, которые не
узнают расширение `.skill`, скопируй файл и переименуй копию в `.zip`.

## ChatGPT, Claude.ai без импорта, Gemini и другие веб-системы

Веб-системы без импорта навыков не устанавливают папку целиком:

1. Распакуй ZIP.
2. Помести содержимое `PORTABLE-PROMPT.md` в системные или пользовательские
   инструкции проекта.
3. Загрузи файлы знаний по таблице `PORTABLE-PROMPT.md` → «Файлы для веб-систем»:
   обязательный набор включает всё, что `SKILL.md` велит прочитать в начале работы,
   справочники рабочего пути и `data/ai_cliches.json`.
4. DOCX из `assets/` и каталог `scripts/` сохраняй для локального запуска.

Если платформа не создаёт независимые контексты, скилл остаётся работоспособным,
но аудит получает отметку `degraded_independence` и итог не выше
`READY_FOR_SUPERVISOR_REVIEW`.

## API и агентные фреймворки

Передай `PORTABLE-PROMPT.md` как постоянную инструкцию, `SKILL.md` и файлы из той же
таблицы — через файловый поиск или базу знаний. Оркестратор должен уметь запускать
скрипты, создавать изолированные read-only задания аудиторов и возвращать их
JSON-отчёты основному агенту для `vkr_audit.py record`.

## Первый запуск

Пользователь пишет запрос (синтаксис вызова — выше) и отвечает на intake.
Основной агент сам создаёт отдельную папку проекта и всю служебную структуру;
вручную её создавать не нужно. Сквозной путь с командами —
`references/quickstart.md`.

В следующей локальной сессии открой ту же папку проекта. В веб-системе приложи ZIP
проекта или файлы из `continuation_files` вывода `vkr_memory.py status`
(`references/project-memory.md` → «Завершение и передача в новый чат»).

## Обновление при незавершённом проекте

Проекты прежних версий продолжаются без повторного интервью:

1. Установи новую версию по шагам выше.
2. В новой сессии агент запускает `vkr_memory.py <PROJECT_DIR> status --json` —
   ожидаемо `needs_rebaseline` — и doctor на стадии `draft`, разбирает изменения и
   переиндексирует файлы:

```bash
python <SKILL_DIR>/scripts/vkr_memory.py <PROJECT_DIR> record --rebaseline --json
```

3. Записи аудита, сделанные вручную по протоколу до 6.33, doctor показывает как
   `AUDIT_LEGACY_RUNS` и не засчитывает: нужные проверки проводятся заново через
   `vkr_audit.py`.
4. Готовый DOCX прежней версии, если текст правился только в нём, переносится в
   черновики через `import_docx.py` (`references/draft-polish-workflow.md`).

## Проверка установки

Правильная структура начинается так:

```text
vkr-mpgu/
├── SKILL.md
├── INSTALL-RU.md
├── PORTABLE-PROMPT.md
├── agents/
├── assets/
├── data/
├── references/
└── scripts/
```

Частая ошибка — лишняя вложенность `.../skills/vkr-mpgu/vkr-mpgu/SKILL.md`:
перемести внутреннюю папку на один уровень выше. Вторая частая ошибка — старая и
новая версии одновременно (папка в `skills` и навык, подключённый в аккаунте платформы).

Проверка по содержимому — `SKILL.md` содержит `6.33-production` (команды выше).
Для разработчика — полный набор тестов из корня репозитория:

```bash
python -m unittest discover -s tests -t .
```

SHA-256 официальных файлов поставки `.skill` и `.zip` должен совпадать со значением
в отчёте о проверке версии рядом с архивами.
