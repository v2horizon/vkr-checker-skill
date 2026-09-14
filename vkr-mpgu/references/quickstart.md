# Quickstart: сквозной путь от intake до сдачи

Это канонический runbook версии `6.33-production`: каждая команда ниже проверена
запуском на тестовом проекте, после каждой указан ожидаемый результат. Команды
запускает основной ИИ-агент; пользователь отвечает на вопросы, даёт факты и файлы.
Подробности каждого шага — в файлах, на которые ведут ссылки.

## Шаг 0. Ожидания и обозначения

Скилл **не заменит** работу студента целиком. Он помогает написать и оформить
текст, собрать и проверить библиографию, проверить работу и подготовиться к
защите. **Не делает** за студента: реальный продукт, пилотирование (выборка и
результаты — только реальные; минимум 5 участников методичка 09.03.02 задаёт лишь
для частной образовательной практики), согласование темы и саму защиту.

- `<SKILL_DIR>` — абсолютный путь к папке установленного скилла (где лежит `SKILL.md`).
- `<PROJECT_DIR>` — отдельная папка проекта ВКР, не внутри скилла.
- В macOS/Linux вместо `python` может понадобиться `python3`.
- Коды возврата всех скриптов: 0 — успех; 1 — найдены ошибки или проверка не
  пройдена; 2 — ошибка использования, ввода-вывода или зависимости.
- `--json` печатает только JSON; `-o FILE` пишет отчёт в UTF-8. В Windows
  PowerShell 5.1 перехваченный вывод искажает кириллицу — читай JSON из файла `-o`
  или сначала выполни `$OutputEncoding = [Console]::OutputEncoding = [Text.UTF8Encoding]::new()`.
- Нужны Python 3.9+ и `python-docx` (для YAML-реестра — ещё `PyYAML`); см. `INSTALL-RU.md`.

Единственный сдаваемый файл — `<PROJECT_DIR>/final/vkr.docx`. Промежуточные DOCX и
отчёты анализаторов — только в `<PROJECT_DIR>/exports/`.

## Шаг 1. Intake и создание проекта

1. Интервью по `intake-interview.md`: вуз, программа, тема, формат, срок, продукт,
   пилот, материалы, **правила кафедры об использовании ИИ**. Режим по сроку —
   таблица в `SKILL.md` → «С чего начать любой диалог».
2. Итог интервью агент записывает в `<PROJECT_DIR>/intake.json` по схеме
   `project-initialization.md` → «Конфигурация intake»: `topic`, `vkr_type`,
   `profile`, `mode`, `deadline`, `collective` (JSON `true`/`false`), `members` и
   все поля `title_page` (`university`, `institute`, `department`, `program_code`,
   `program_name`, `program_profile`, `work_type`, `author`, `group`, `supervisor`,
   `supervisor_title`, `head_of_department`, `head_title`, `city`, `year`).
   Файл можно положить и рядом с проектом, как в примере `project-initialization.md`.
3. Создание проекта:

```bash
python <SKILL_DIR>/scripts/init_vkr_project.py <PROJECT_DIR> --config <PROJECT_DIR>/intake.json --json
```

Ожидается: код 0, `"status": "initialized"`, в `created` — `vkr-project.json`,
`vkr-state.md`, `plan.md`, `sources.json`, `drafts/*.md`, `evidence/index.json`,
`audit/…`, `memory/…`, `logs/…`; для `mpgu-*` — копия методички в
`sources/materials/`. `audit_intensity` без явного значения: `strict` для
`standard`, `balanced` для `express-*`. Код 2 — ошибка конфигурации (например,
`profile` противоречит `vkr_type`, `collective` не bool, неизвестный ключ) или
конфликт с существующим проектом.

4. Агент дописывает `vkr-state.md` ответами intake (`state-file-pattern.md`) и
   проверяет проект:

```bash
python <SKILL_DIR>/scripts/vkr_project_doctor.py <PROJECT_DIR> --stage draft --json
python <SKILL_DIR>/scripts/vkr_memory.py <PROJECT_DIR> status --json
```

Ожидается: doctor — код 0, `"status": "PASS"`, `"readiness": "NOT_READY"`, справки
`PLACEHOLDER_FOUND` о маркерах в заготовках нормальны. Память — после дописывания state код 1 и
`"status": "external_changes_detected"` с `vkr-state.md` в `changed` (и `"handoff_stale": true`):
это норма, state только что дописан агентом после init (если запустить `status` до
правки state, будет `ok` и код 0 — тоже норма). Изменение фиксирует checkpoint (event —
`project-memory.md` → «Формат event»), после него `status` — `ok`:

```bash
python <SKILL_DIR>/scripts/vkr_memory.py <PROJECT_DIR> record --event <EVENT_JSON> --all-changed --json
```

5. Аудит точки «intake завершён» (матрица `continuous-audit.md`): материалы
   продукта и пилота от пользователя сначала кладутся в `evidence/` и записываются в
   `evidence/index.json`, затем:

```bash
python <SKILL_DIR>/scripts/vkr_audit.py <PROJECT_DIR> plan --kind primary --preset intake --json
```

Ожидается: `run_id` вида `R001-primary-intake`, задачи `MET`, `EVD`, `TEC`; дальше
`brief`, `record`, `close` — как в шаге 2.

## Шаг 2. План и аудит плана

Агент заполняет `plan.md` (тема → цель → задачи → главы → ожидаемые результаты),
затем планирует независимую проверку точки «план»:

```bash
python <SKILL_DIR>/scripts/vkr_audit.py <PROJECT_DIR> plan --kind primary --preset plan --json
python <SKILL_DIR>/scripts/vkr_audit.py <PROJECT_DIR> brief --run <RUN_ID> --json
```

Ожидается: `"status": "planned"`, `run_id` вида `R002-primary-plan` (номер — следующий после аудита intake), в `tasks` —
задачи ролей (`strict`: `MET`, `LOG-A`, `LOG-B`, `EVD`, `TEC`); `brief` пишет
`audit/briefs/<RUN_ID>/<TASK_ID>.md`. Каждое задание агент отдаёт отдельному
субагенту-аудитору в чистом контексте. Ответ аудитора (один JSON) агент сохраняет
в файл вне `audit/` и записывает:

```bash
python <SKILL_DIR>/scripts/vkr_audit.py <PROJECT_DIR> record --run <RUN_ID> --task <TASK_ID> --auditor-id <ID> --report <FILE> --json
python <SKILL_DIR>/scripts/vkr_audit.py <PROJECT_DIR> close --run <RUN_ID> --json
```

Ожидается: `record` — код 0, `"status": "recorded"`, `task_status`, найденные
замечания `F-<GATE>-<nnn>` в `new_findings`; код 1 — отчёт отклонён (схема,
`AUDITOR_NOT_INDEPENDENT`, `SNAPSHOT_STALE` — файлы менялись после `plan`).
`close` — код 0 и `"status": "complete"`, либо код 1 и `"status": "reported"` с
`open_high_findings`. Пока run не закрыт, файлы проекта не меняются. Без
субагентов роли выполняются последовательно с `--independence degraded_independence`
(см. `continuous-audit.md`).

## Шаг 3. Черновики, источники, утверждения, контрольные точки

- **Текст** — только в `drafts/*.md` по `drafts-format.md`: `annotation.md`,
  `introduction.md`, `chapter-1.md` … `chapter-N.md`, `conclusion.md`,
  `appendix-N.md`; ссылки `[@id]`, `[@id, с. 23]`; таблицы, листинги и рисунки —
  по разметке формата. Порядок и темп — `writing-workflow.md` или
  `writing-workflow-express.md`.
- **Источники** — в `sources.json` (схема — `gost-citations.md` → «Реестр
  источников»). Проверка по `source-verification.md`:

```bash
python <SKILL_DIR>/scripts/verify_sources.py --queries <PROJECT_DIR>/sources.json
python <SKILL_DIR>/scripts/verify_sources.py --mark <PROJECT_DIR>/sources.json --id <ID> --status confirmed --method catalog --url <URL>
python <SKILL_DIR>/scripts/verify_sources.py --report <PROJECT_DIR>/sources.json
```

Ожидается: `--queries` — запросы для непроверенных записей (код 0; поиск выполняет
агент через `web_search`); `--mark` — код 0, скрипт пишет `status` и
`verification` с текущим `checked_at` и хешем реквизитов `fields_sha256` (для
`suspicious`/`rejected` нужен `--note`, иначе код 2); `--report` — код 1, пока есть
`pending`/`suspicious`, `confirmed` без полного `verification` или `confirmed_stale`
(реквизиты правили после `--mark` — перепроверить и отметить заново), и код 0,
когда всё подтверждено.

- **Утверждения** — ключевые тезисы, числа и результаты в
  `memory/claims-register.json` (`project-memory.md` → «Реестр утверждений»);
  доказательства — файлы в `evidence/` и записи `evidence/index.json`.
- **Контрольные точки аудита** — по матрице `continuous-audit.md` на границах
  смысловых блоков (подраздел в `strict`, глава, пакет 10–15 источников, продукт,
  пилот, введение, заключение, сборка DOCX, пакет правок), например:

```bash
python <SKILL_DIR>/scripts/vkr_audit.py <PROJECT_DIR> plan --kind primary --preset chapter --checkpoint chapter-2 --json
python <SKILL_DIR>/scripts/vkr_audit.py <PROJECT_DIR> next --stage draft --json
```

Ожидается: `plan` — задачи пресета (глава в `strict`: `MET`, `LOG-A`, `LOG-B`,
`SRC`, `LNG`, `STY`, `OWN-QUESTIONS`, для проекта ещё `EVD`, `TEC`); `next` —
незакрытые run, ошибки doctor и `next_actions` с готовыми командами.

- **Checkpoint памяти** после законченного подраздела, главы, run аудита и в
  конце сессии: агент пишет event JSON (`project-memory.md` → «Формат event»;
  изменённые черновики, реестры и новые отчёты `audit/reports/…` перечислять в
  `files` не нужно — их берёт `--all-changed` из того же расчёта, что `status`) и
  запускает:

```bash
python <SKILL_DIR>/scripts/vkr_memory.py <PROJECT_DIR> record --event <EVENT_JSON> --all-changed --json
```

Ожидается: код 0, `"status": "recorded"`; после этого `status` снова `ok`.

## Шаг 4. Сборка и проверка DOCX

```bash
python <SKILL_DIR>/scripts/build_vkr.py <PROJECT_DIR>
python <SKILL_DIR>/scripts/update_docx_fields.py <PROJECT_DIR>/final/vkr.docx
python <SKILL_DIR>/scripts/clean_docx_metadata.py <PROJECT_DIR>/final/vkr.docx --in-place
python <SKILL_DIR>/scripts/docx_integrity.py <PROJECT_DIR>/final/vkr.docx
python <SKILL_DIR>/scripts/vkr_audit.py <PROJECT_DIR> validate --json
```

Ожидается:
1. `build_vkr.py` — код 0, `OK: собрано final/vkr.docx`; пишет
   `exports/vkr-content.json` и `exports/citation-map.json`, прежний DOCX копирует в
   `backups/checkpoints/<UTC>-build/`. Код 1 — ошибки содержания (неизвестный
   `[@id]`, цитируется `rejected` или запись только с `raw_text`, нет обязательного
   черновика), файлы не перезаписаны. Код 2 `OUTPUT_LOCKED` — `final/vkr.docx`
   открыт в Word: закрыть файл и повторить; сборка отказывает до любых записей,
   ничего не изменено.
2. `update_docx_fields.py` — в Windows с Microsoft Word код 0, `status: updated`,
   оглавление и номера страниц обновлены, копия — в
   `backups/checkpoints/<UTC>-fields/`, число страниц — в `pages` и
   `exports/docx-fields.json` (там же `word_opened: true` — Word документ
   действительно открыл). Без Word — код 1 и ручной шаг: открыть
   `final/vkr.docx` в Word, Ctrl+A, F9 («Обновить целиком» для оглавления),
   сохранить под тем же именем и закрыть. Код 2 — ошибка ввода-вывода, чаще всего
   `final/vkr.docx` открыт в Word («закрой файл в Word»): закрыть файл и повторить;
   Word не запускается, файл не меняется. Код 2 со `status: open_failed` — Word
   запустился, но отказался открыть документ: разметка DOCX нарушена, проверь её
   `docx_integrity.py` и пересобери файл (`build_vkr.py`, правки из прежнего DOCX —
   `import_docx.py --update`). При тайм-ауте останавливается только Word,
   запущенный этой командой.
3. `clean_docx_metadata.py --in-place` — код 0; файл очищен на месте, `created` и
   `modified` — время очистки (дата создания в прошлое не сдвигается). Если
   чистить нечего, файл не переписывается: `"status": "already_clean"`. Без
   `--in-place` рядом появится второй DOCX, и doctor final откажет.
4. `docx_integrity.py` — структурная проверка пакета без Word: код 0 — проблем нет;
   код 1 — найдены причины, по которым Word назовёт файл повреждённым (список в
   выводе; пересобрать шаги 1–3); код 2 — файл не найден, не читается или не ZIP.
5. `vkr_audit.py validate` — код 0, `"status": "validated"`, `summary.errors: 0`,
   отчёт `audit/automated-validation.json` с профилем из `vkr-project.json`. Код 1 —
   ошибки валидатора в `errors_preview`: исправить черновики и повторить шаг 4.
   Без обновления полей валидатор сообщает «Оглавление НЕ обновлено».
   Предупреждения (`summary.warnings`, первые — в `warnings_preview`, все — в
   `audit/automated-validation.json`, пункты с `"ok": false` и `"severity": "warning"`)
   gate не блокируют, но их обязательно прочитать и по каждому решить: исправить в
   черновиках (затем пересборка и повтор шага 4) или записать обоснование в state.
   Doctor показывает их как `VALIDATION_WARNINGS`.

Число страниц из `update_docx_fields.py` вносится в аннотацию, затем пересборка и
повтор шага 4 (`writing-workflow.md` → «Этап 6: Аннотация»); doctor prefinal/final
сверяет объём в аннотации с числом страниц самого DOCX (`docProps/app.xml`,
его пишет Word) и с `exports/docx-fields.json` (`ANNOTATION_PAGES_MISMATCH`;
расхождение этих двух источников — `DOCX_PAGES_MISMATCH`).

**После шага 4 файл больше не трогать.** Повторные `build_vkr.py`,
`update_docx_fields.py` или `clean_docx_metadata.py --in-place` переписывают байты
`final/vkr.docx`: снимок аудита устаревает (`AUDIT_SNAPSHOT_STALE`) и волны
проверок приходится проводить заново. Повторная очистка уже чистого файла
безопасна: она отвечает `"status": "already_clean"` и файл не переписывает.

Анализаторы (самопроверка; отчёты — в `exports/`, `analyzers-cli.md`):

```bash
python <SKILL_DIR>/scripts/ai_detection_heuristic.py <PROJECT_DIR>/final/vkr.docx -o <PROJECT_DIR>/exports/ai-detection.json
python <SKILL_DIR>/scripts/content_ownership_check.py <PROJECT_DIR>/final/vkr.docx --all -o <PROJECT_DIR>/exports/ownership-questions.json
```

Ожидается: код 0 и уровень риска `LOW`/`MODERATE`/`HIGH`/`CRITICAL`; код 1 —
`INSUFFICIENT_DATA` (меньше 300 слов основного текста или разделы не распознаны).
HIGH и CRITICAL — повод переписать места по `humanizer-techniques.md` → «Калибровка»,
а не доказательство авторства.

## Шаг 5. Версия для научного руководителя

```bash
python <SKILL_DIR>/scripts/vkr_audit.py <PROJECT_DIR> snapshot --json
python <SKILL_DIR>/scripts/vkr_audit.py <PROJECT_DIR> plan --kind primary --preset prefinal --json
python <SKILL_DIR>/scripts/vkr_project_doctor.py <PROJECT_DIR> --stage prefinal --json
```

После `plan` — `brief`, аудиторы, `record`, `close`, как в шаге 2 (девять ролей с
`full_gate`). Ожидается: doctor — код 0, `"status": "PASS"`,
`"readiness": "READY_FOR_SUPERVISOR_REVIEW"`. На prefinal `[ПРОВЕРИТЬ…]` и
неподтверждённые источники — предупреждения со списком; другие маркеры, короткие
обязательные черновики (меньше 300 значимых символов), открытые BLOCKER/MAJOR,
незакрытые run, `final/vkr.docx`, не совпадающий с черновиками
(`FINAL_DOCX_OUTDATED`), и объём в аннотации, расходящийся с DOCX
(`ANNOTATION_PAGES_MISMATCH`), — ошибки. Без prefinal-волны в `strict`/`maximum`
doctor предупреждает `AUDIT_PREFINAL_WAVE_MISSING` и в `next_actions` вместо «можно
передавать» напоминает о волне. Руководителю передаётся `final/vkr.docx`.

## Шаг 6. Финал

Предусловия: ответы Test-30 записаны (шаг 7); `verify_sources.py --report` даёт
код 0; все утверждения `confirmed` с доказательствами; данные титула в
`vkr-project.json` заполнены; последний `final/vkr.docx` прошёл шаг 4.

```bash
python <SKILL_DIR>/scripts/vkr_audit.py <PROJECT_DIR> snapshot --json
python <SKILL_DIR>/scripts/vkr_audit.py <PROJECT_DIR> plan --kind primary --preset final --json
```

Затем `brief`, независимые аудиторы, `record` по каждой задаче, `close`.
Ожидается: в `balanced` 9 задач, в `strict` 15 (B-реплики `MET`, `SRC`, `LOG`, `EVD`,
`TEC`, `DOC`), в `maximum` 18; все с `audit_lens: full_gate`.

**Если есть замечания** (`close` → `reported`):
1. `python <SKILL_DIR>/scripts/vkr_audit.py <PROJECT_DIR> findings --open --json` — список.
2. Исправляет только основной агент: черновики, реестры, доказательства; факты,
   которых нет, запрашиваются у пользователя (`USER_EVIDENCE_REQUIRED`).
3. Пересборка и проверка — весь шаг 4, затем новый снимок:

```bash
python <SKILL_DIR>/scripts/vkr_audit.py <PROJECT_DIR> snapshot --json
python <SKILL_DIR>/scripts/vkr_audit.py <PROJECT_DIR> plan --kind targeted_recheck --findings <F-ID> --json
```

4. Адресную перепроверку выполняет другой аудитор; в отчёте — `rechecks` с
   вердиктом. Ожидается: в выводе `record` — `rechecks[].finding_status: "resolved"`
   (у каждого перепроверенного замечания; поля `finding_status` верхнего уровня нет),
   затем `close`.
5. Новая primary-волна и более поздняя слепая волна на том же, уже исправленном
   снимке (адресной перепроверки без новой primary недостаточно):

```bash
python <SKILL_DIR>/scripts/vkr_audit.py <PROJECT_DIR> plan --kind primary --preset final --json
python <SKILL_DIR>/scripts/vkr_audit.py <PROJECT_DIR> plan --kind blind_regression --preset final --json
```

Каждую — `brief`, `record`, `close`. Слепую волну выполняют аудиторы, которые не
участвовали в primary этого снимка (9 задач; в `maximum` — 18).

**Итог:**

```bash
python <SKILL_DIR>/scripts/vkr_memory.py <PROJECT_DIR> record --event <EVENT_JSON> --all-changed --json
python <SKILL_DIR>/scripts/vkr_project_doctor.py <PROJECT_DIR> --stage final --json
```

Ожидается: код 0, `"status": "PASS"`, `"readiness": "READY_TO_SUBMIT"`. При
`degraded_independence` — не выше `READY_FOR_SUPERVISOR_REVIEW`. Код 1 — FAIL с
кодами находок и `next_actions` (например `AUDIT_SNAPSHOT_STALE` — файлы менялись
после снимка: новый снимок и новые волны; `VALIDATION_ERRORS`; `AUDIT_BLIND_WAVE_MISSING`).
Запись state и памяти после снимка его не меняет: раздел «Независимый аудит»,
журнал и история исключены из проекции state.
После финального аудита не открывай `final/vkr.docx` на запись и не пересохраняй
его: любое изменение байтов требует нового снимка и финальных волн; для просмотра
открывай копию или режим «только для чтения».

## Шаг 7. Защита

- Презентация, речь и банк вопросов — `defense.md`; владение материалом — `defense-mastery.md`.
- Test-30 и репетицию вопросов комиссии проводит основной агент с пользователем.
  Ответы пользователя своими словами сохраняются файлом, например
  `evidence/defense/test30-2027-05-10.md`, и вносятся в `evidence/index.json` →
  `defense`: `{"id": "defense-test30", "path": "evidence/defense/test30-2027-05-10.md"}`.
- Это делается **до** финального снимка: файлы `evidence/` входят в снимок, а роль
  `OWN` в обеих финальных волнах оценивает записанные ответы. Без них
  `plan --preset final` отказывает с `USER_EVIDENCE_REQUIRED`.

## Шаг 8. Правки в Word и чужой черновик

**Пользователь правил `final/vkr.docx` в Word.** Правки возвращаются в черновики,
иначе следующая сборка их затрёт:

```bash
python <SKILL_DIR>/scripts/import_docx.py <PROJECT_DIR> --docx <PROJECT_DIR>/final/vkr.docx --update
```

Ожидается: код 0, в `files.updated` — изменённые черновики; числа `[N]` переведены
обратно в `[@id]` по `exports/citation-map.json`; `sources.json` не меняется
(правки списка литературы вносятся в реестр). Затем шаг 4 и новые проверки.
Код 1 со `status: "fail"` — импорт остановлен, черновики не изменены:
- «список литературы в DOCX не совпадает с последней сборкой» — добавь/исправь
  источник в `sources.json` и пересобери; правки текста перенеси после этого;
- «черновики изменены после сборки: …» — если главная правка в черновике,
  пересобери; если главная правка в Word —
  `python <SKILL_DIR>/scripts/import_docx.py <PROJECT_DIR> --docx <PROJECT_DIR>/final/vkr.docx --update --overwrite-drafts`
  (прежние черновики — в `backups/checkpoints/<UTC>-import/`).

**Доработка чужого готового DOCX:** шаг 1 (init) →

```bash
python <SKILL_DIR>/scripts/import_docx.py <PROJECT_DIR> --docx <FILE>
```

Ожидается: код 0, черновики `drafts/*.md` созданы из разделов DOCX, рисунки — в
`evidence/figures/`, список литературы — записи `{"id": N, "raw_text": …,
"status": "pending"}` в `sources.json`, отчёт `audit/import-report.json`. Код 1 —
элементы с потерей содержимого (формулы, SmartArt, сноски; `status: "incomplete"`)
перечислены в отчёте и переносятся вручную; либо (`status: "fail"`) номера списка
литературы неоднозначны — повторяющийся номер, абзац без номера: импорт
остановлен, ничего не записано, причины в `sources.problems`; исправь список в
копии DOCX и повтори. Дальше — `draft-polish-workflow.md` и тот же путь с шага 3.

## Шаг 9. Продолжение в новом чате

```bash
python <SKILL_DIR>/scripts/vkr_memory.py <PROJECT_DIR> status --json
python <SKILL_DIR>/scripts/vkr_audit.py <PROJECT_DIR> next --stage draft --json
```

`--stage` — по текущему этапу (`draft`, `prefinal` или `final`). Ожидается: `status`
— `ok` (код 0) или `external_changes_detected`/`needs_rebaseline`/`recovery_required`
(код 1, разбор — `project-memory.md`); `next` — что делать дальше. `"handoff_stale": true`
в выводе обеих команд значит, что файлы или журнал аудита изменились после
checkpoint, записавшего handoff (чат оборвался до записи памяти): первый пункт
`next_actions` велит не продолжать по `next_actions` handoff, а прочитать
`vkr-state.md` и `audit/manifest.json`, записать `record --event <EVENT_JSON> --all-changed`
с актуальными `current_task` и `next_actions`; после этого флаг снимается. В локальной
агентной системе достаточно открыть корень проекта; в системе с вложениями —
файлы из `continuation_files` вывода `status` (список — `project-memory.md` →
«Завершение и передача в новый чат») или ZIP проекта.

## Типичные ошибки

- «Напиши мне диплом за час» — не получится: серьёзная работа требует 30–50 часов
  времени студента.
- Правка `final/vkr.docx` в Word без `import_docx.py --update` — следующая сборка
  затрёт правки.
- Второй DOCX в `final/` (копия, `vkr-clean.docx`, «версия для научрука») — doctor
  final откажет; промежуточные файлы — в `exports/`.
- Выдуманные источники, числа или результаты пилота — ни один gate их не пропустит,
  а комиссия спросит.
- Пропуск Test-30 — роль `OWN` не может пройти финал, и защита превращается в лотерею.

## Типичный тайминг

| Этап | Время |
|------|-------|
| Intake-интервью | 1 час |
| Введение | 2–3 часа |
| Теоретическая глава | 6–10 часов |
| Аналитическая глава | 8–12 часов (+ сбор данных) |
| Проектная глава | 10–15 часов (+ продукт) |
| Заключение и аннотация | 2 часа |
| Проверка источников | 3–8 часов |
| Сборка и проверка DOCX | 1 час |
| Правки научного руководителя | 5–15 часов |
| Презентация и речь | 4–6 часов |
| Test-30 и репетиции защиты | 3–5 часов |
| **Итого** | **45–75 часов чистого времени** |

Это без самого продукта и пилотирования. Время на субагентов-аудиторов — по
таблице «Сколько это стоит» в `continuous-audit.md`.
