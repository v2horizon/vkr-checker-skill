# Долговременная память, roadmap и журналы

История чата не является надёжной памятью проекта. Основной агент сохраняет
рабочее состояние в файлах проекта, чтобы новый чат мог восстановить контекст,
проверить изменения и продолжить работу без повторного интервью.

## Структура памяти

```text
memory/
├── handoff.json             # машинно-читаемое состояние для нового чата (пишет vkr_memory.py)
├── handoff.md               # короткое человеческое резюме (пишет vkr_memory.py)
├── roadmap.md               # этапы, статусы и ближайшие шаги (ведёт основной агент)
├── decisions.jsonl          # неизменяемый журнал решений (пишет vkr_memory.py)
├── claims-register.json     # ключевые утверждения и их доказательства (ведёт основной агент)
└── artifact-index.json      # последние известные SHA-256 файлов проекта (пишет vkr_memory.py)
logs/
├── activity.jsonl           # append-only журнал checkpoint (пишет vkr_memory.py)
└── tool-runs.jsonl          # события event_type=tool_run из vkr_memory.py record; без них файл пуст
```

`handoff.*`, `decisions.jsonl`, `artifact-index.json` и `logs/*` — служебные
файлы: их пишет только скрипт, и они не считаются внешними изменениями.

## Что записывать

На каждой контрольной точке основной агент фиксирует: что завершено и какие
файлы изменены; текущее задание и 1–5 следующих действий; принятые решения,
основание и кем они подтверждены; открытые вопросы и блокеры; факты, которые
требуют данных пользователя или согласования руководителя.

Не сохраняй скрытую цепочку рассуждений, пароли, токены, cookies, секретные
конфигурации и полные выводы инструментов. Сохраняй краткое проверяемое
основание решения и ссылку на артефакт.

## Автоматический старт каждой сессии

До написания нового текста основной агент:

1. читает `vkr-state.md`, `memory/handoff.json` и `memory/handoff.md`;
2. читает roadmap, решения и утверждения текущего раздела;
3. запускает

   ```bash
   python <SKILL_DIR>/scripts/vkr_memory.py <PROJECT_DIR> status --json
   python <SKILL_DIR>/scripts/vkr_audit.py <PROJECT_DIR> next --stage draft --json
   ```

4. разбирает внешние изменения: изменённые, исчезнувшие и новые файлы;
5. сообщает пользователю кратко: где остановились, что изменилось, что будет
   сделано сейчас.

`status` возвращает:

| `status` | Значение | Код |
|---|---|---|
| `ok` | файлы совпадают с последним checkpoint | 0 |
| `external_changes_detected` | есть изменённые, исчезнувшие или новые файлы в `drafts/`, `final/`, `evidence/`, `sources/`, `audit/reports/`, корне проекта, `memory/roadmap.md` или `memory/claims-register.json` | 1 |
| `needs_rebaseline` | индекс создан прежней версией скилла | 1 |
| `recovery_required` | нет файлов памяти или прервана запись checkpoint | 1 |
| `error` | повреждённые входные данные; поле `error` называет файл | 2 |

Если `handoff` повреждён, восстанови его из state, roadmap, журналов и
`audit/manifest.json`. Не придумывай прошлое состояние.

## Автоматическая запись checkpoint

После законченного подраздела, главы, пакета источников, продукта или пилота,
сборки DOCX, пакета исправлений, аудита и в конце сессии основной агент пишет
короткий event JSON и запускает:

```bash
python <SKILL_DIR>/scripts/vkr_memory.py <PROJECT_DIR> record --event <EVENT_JSON> --all-changed --json
```

`--all-changed` берёт изменённые, новые и удалённые файлы из того же расчёта, что
`status`: черновики, реестры, `final/vkr.docx`, доказательства и отчёты
`audit/reports/…`, которые записал `vkr_audit.py`. Перечислять их не нужно; поля
`files` и `removed_files` события, если заданы, добавляются к найденному. Без флага
в checkpoint попадают только файлы из этих полей.

Скрипт сам вычисляет хеши файлов, обновляет индекс, handoff и решения и
последним шагом дописывает событие в `logs/activity.jsonl`. Если запись
прервалась, `status` покажет `recovery_required`; повтори ту же команду — событие
и решения не задвоятся. Повтор уже записанного события возвращает
`already_recorded`. Пользователь эту команду вручную не запускает.

Event JSON можно держать внутри проекта (например, `<PROJECT_DIR>/event.json`):
`record` его не индексирует и запоминает как служебный, поэтому после записи
checkpoint `status` не считает этот файл внешним изменением, даже когда в него
записано следующее событие.

Для запуска штатного инструмента запиши событие `tool_run` с именем скрипта,
exit status и кратким результатом.

## Формат event

```json
{
  "event_type": "subsection_completed",
  "summary": "Завершён подраздел 2.3 и проверены связанные источники",
  "current_phase": "chapter-2",
  "current_task": "Исправление замечаний F-LOG-002 и F-SRC-004",
  "last_completed": "drafts/chapter-2.md: раздел 2.3",
  "next_actions": ["исправить F-LOG-002", "plan --kind targeted_recheck --findings F-LOG-002"],
  "open_questions": [],
  "open_blockers": ["F-SRC-004: нужен оригинал статьи"],
  "files": ["drafts/chapter-2.md", "sources.json", "audit/reports/R004-primary-chapter-2/R004-LOG-A.json"],
  "removed_files": [],
  "decisions": [{
    "decision": "Сохранить две метрики пилота",
    "basis": "Они присутствуют в исходной таблице",
    "confirmed_by": "user_evidence"
  }]
}
```

`files` — пути относительно корня проекта; служебные файлы памяти и журналов в
них не указываются. С `--all-changed` поля `files` и `removed_files` можно не
заполнять. `removed_files` снимает с учёта удалённые или
переименованные файлы, в том числе ошибочно записанные служебные пути. Хеши,
`event_id` и метки времени скрипт добавляет сам; `event_id` можно передать,
только если нужен собственный идентификатор вида `evt-…`.

## Переиндексация после обновления скилла

Проект, созданный прежней версией, получает `needs_rebaseline`. Проверь
изменения (`status`, doctor `draft`), затем:

```bash
python <SKILL_DIR>/scripts/vkr_memory.py <PROJECT_DIR> record --rebaseline --json
```

Команда заново индексирует все отслеживаемые файлы и удаляет из индекса
служебные пути. До переиндексации `record --all-changed` отказывает (код 2).

Аудит прежней версии не засчитывается. Первая команда `vkr_audit.py`,
которая пишет manifest, один раз переносит run протокола до 6.33 в
`legacy_runs` (с контрольной суммой и копией `audit/manifest.json.bak`); doctor
показывает их как `AUDIT_LEGACY_RUNS`. Отчёты 6.32 из `audit/reports` переносятся
в `audit/legacy/` командой:

```bash
python <SKILL_DIR>/scripts/vkr_audit.py <PROJECT_DIR> migrate-legacy --json
```

Та же команда переносит плоские поля титула 6.32 из `vkr-project.json`
(`student_name`, `institution`, `institute`, `program_code`, `program_name`,
`supervisor`, `department`, `group`) в `title_page`, если его ещё нет, и
перечисляет в `title_page_missing` поля, которые нужно дописать до сборки.

Замечания прежних отчётов проверяются заново волнами `plan` → `record` →
`close`. Run, дописанный в `legacy_runs` уже нового manifest, или разрыв
нумерации doctor показывает как `AUDIT_RUN_TAMPERED`/`AUDIT_SEQ_GAP`: такой
manifest восстанавливают из `audit/manifest.json.bak` или переносят весь аудит в
`audit/legacy/<дата>/` и проводят волны заново.

## Roadmap

Roadmap ведётся по результатам, а не по календарным обещаниям:

- `NOT_STARTED` — данных достаточно для планирования, работа не начата;
- `IN_PROGRESS` — основной агент работает над этапом;
- `WAITING_USER` — нужен факт, файл или решение пользователя;
- `WAITING_SUPERVISOR` — нужно согласование руководителя (замечание класса
  `SUPERVISOR_APPROVAL`);
- `IN_AUDIT` — снимок зафиксирован, run не закрыт;
- `FIXING` — основной агент исправляет замечания;
- `VERIFIED` — точка закрыта на текущем снимке;
- `INVALIDATED` — зависимый файл изменился, проверку нужно повторить.

Готовность работы в целом — только поле `readiness` doctor
(`READY_TO_SUBMIT`, `READY_FOR_SUPERVISOR_REVIEW`, `NOT_READY`).

## Реестр утверждений

`memory/claims-register.json` — массив ключевых утверждений:

```json
[{
  "claim_id": "CH2-CLAIM-014",
  "text": "Краткая формулировка тезиса",
  "location": "Глава II, 2.3",
  "status": "confirmed",
  "source_ids": ["ivanov2020"],
  "evidence_ids": ["pilot-table-02"],
  "evidence": []
}]
```

- `status`: `pending`, `confirmed`, `partial`, `unsupported`, `invalidated`.
  Старые `supported` и `verified` читаются как `confirmed`, но пиши `confirmed`.
- `source_ids` — id из `sources.json`; на финале источник должен быть
  `confirmed`.
- `evidence_ids` — id записей `evidence/index.json`; id источника сюда не пишется.
- `evidence` — необязательный массив путей внутри `evidence/` или
  `sources/materials/`, уже перечисленных в `evidence/index.json`; ссылки на
  `vkr-state.md`, черновики, `exports/` и другие служебные файлы не являются
  доказательством.
- На финале у каждого утверждения есть `claim_id`, `text`, `location`, статус
  `confirmed` и хотя бы одна ссылка на подтверждённый источник или доказательство;
  пустые файлы доказательств не принимаются, URL без локальной копии doctor
  отмечает предупреждением.

Итог аудита и хеши снимка в реестр утверждений не пишутся: они хранятся в
`audit/manifest.json` и `audit/findings.json`. Реестр утверждений входит в
снимок, поэтому его правка после финального аудита требует новых волн.

При изменении источника или данных найди зависимые утверждения, поставь им
`invalidated` или `pending` и запланируй проверки `SRC`, `LOG`, `EVD` или `TEC`.

## Реестр доказательств

`evidence/index.json` — объект с группами `methodology`, `sources`, `product`,
`pilot`, `figures`, `defense`, `approvals` (письменные одобрения руководителя для
`waive`); каждая группа — массив записей. Группа `pilot` означает «пилот или
эксперимент»: данные замеров, расчётов и вычислительных экспериментов обычной ВКР
кладутся в неё; при отсутствии продукта и пилота `product` и `pilot` остаются
пустыми:

```json
{
  "methodology": [{"id": "methodology-09-03-02-2024", "path": "sources/materials/methodology-09-03-02-2024.docx"}],
  "sources": [],
  "product": [{"id": "product-demo", "path": "evidence/product/demo-2026-05.mp4", "title": "Демонстрация функций"}],
  "pilot": [{"id": "pilot-table-02", "path": "evidence/pilot/results.csv"}],
  "figures": [],
  "defense": [{"id": "defense-test30", "path": "evidence/defense/test30-2026-06-01.md"}]
}
```

Запись: `id` и ровно одно из `path` (файл внутри `evidence/` или
`sources/materials/`) или `url`; необязательные `kind`, `title`, `note`,
`checked_at`. Другие ключи (например `file`) — ошибка doctor. Каждый локальный
файл автоматически входит в снимок аудита.

## Завершение и передача в новый чат

Перед окончанием сессии или ожидаемым сокращением контекста основной агент
сохраняет файлы, записывает checkpoint (`next_actions` — точное следующее
действие) и проверяет `status`. `handoff.md` скрипт формирует сам.

Файлы для продолжения в новом чате (их же перечисляет `status` в
`continuation_files`): `vkr-state.md`, `vkr-project.json`, `plan.md`,
`sources.json`, `evidence/index.json`, `memory/handoff.json`,
`memory/handoff.md`, `memory/roadmap.md`, `memory/claims-register.json`,
`memory/artifact-index.json`, `audit/manifest.json`, `audit/findings.json`,
`audit/snapshot-inputs.json`, текущие `drafts/*.md`. В локальной агентной
системе достаточно открыть корень проекта.

Новый чат не доверяет одному текстовому резюме: он сверяет handoff с файлами,
журналами и `audit/manifest.json`. Независимые аудиторы могут читать снимок, но
не читают и не меняют memory, roadmap и журналы.
