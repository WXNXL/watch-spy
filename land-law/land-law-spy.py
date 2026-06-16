#!/usr/bin/env python3
"""
╔══════════════════════════════════════════════════════════════════╗
║  ГЕНЕРАТОР ШПОРГАЛОК-ЗАДАЧ ДЛЯ GALAXY WATCH 4                  ║
║  Конвертирует .docx с задачами → набор HTML-файлов для Wear OS ║
║                                                                  ║
║  Использование:                                                  ║
║    python3 land-law-spy.py задачи.docx                           ║
║    python3 land-law-spy.py задачи.docx --output ./land-output    ║
║                                                                  ║
║  Требования: pandoc установлен в системе                        ║
║    Ubuntu/Debian: sudo apt install pandoc                        ║
║    macOS:         brew install pandoc                            ║
╚══════════════════════════════════════════════════════════════════╝
"""

import re, os, sys, shutil, subprocess, html as html_mod

# ═══════════════════════════════════════════════════════════════════
#  НАСТРОЙКИ — меняй под свой документ
# ═══════════════════════════════════════════════════════════════════

# Заголовок сайта (отображается на главной странице)
SITE_TITLE = "Задачи по земельному праву"

# Алфавит для панели навигации.
ALPHABET = "АБВГДЕЖЗИКЛМНОПРСТУФХЦЧШЩЭЮЯ"

# Цветовая схема AMOLED
COLOR_BG         = "#000000"
COLOR_TEXT        = "#e0e0e0"
COLOR_ACCENT      = "#ef9a9a"   # Акценты — тёплый красноватый (отличается от билетов)
COLOR_LINK_BG     = "#0a0a0a"
COLOR_LINK_ACTIVE = "#3a1a1a"
COLOR_BORDER      = "#333333"
COLOR_DEF         = "#ffcc80"
COLOR_DIM         = "#999999"
COLOR_HEADER_SUB  = "#b0bec5"
COLOR_TOPIC       = "#80cbc4"   # Цвет заголовков тем
COLOR_ANSWER      = "#a5d6a7"   # Цвет метки «Ответ»

# Размеры шрифтов — крупно для 396x396 круглого экрана!
FONT_BODY     = "20px"
FONT_H1       = "24px"
FONT_H2       = "22px"
FONT_ALPHA    = "18px"
FONT_TASK     = "19px"

# Минимальная высота кликабельной области (px)
TOUCH_TARGET = "56px"

# Отступы для круглого экрана
PADDING_SIDE = "max(22px, 14%)"

# ═══════════════════════════════════════════════════════════════════
#  КОНЕЦ НАСТРОЕК — дальше код, который обычно менять не нужно
# ═══════════════════════════════════════════════════════════════════


def pandoc_to_markdown(docx_path):
    """Конвертирует .docx в Markdown через Pandoc."""
    result = subprocess.run(
        ["pandoc", docx_path, "-t", "markdown", "--wrap=none"],
        capture_output=True, text=True, encoding="utf-8"
    )
    if result.returncode != 0:
        print(f"Ошибка Pandoc: {result.stderr}")
        sys.exit(1)
    return result.stdout


def clean_text(raw):
    """Очищает текст от Pandoc-артефактов."""
    t = re.sub(r'\[([^\]]*)\]\{\.mark\}', r'\1', raw)
    t = re.sub(r'\{\.underline\}', '', t)
    t = re.sub(r'\{#[^}]*\}', '', t)
    t = re.sub(r'\{\.TOC-Heading\}', '', t)
    t = re.sub(r'\{[^}]*\}', '', t)
    t = re.sub(r'\[([^\]]*)\]\(#[^)]*\)', r'\1', t)
    t = re.sub(r'\[([^\]]*)\]\(https?://[^)]*\)', r'\1', t)
    return t.strip()


def extract_number(title):
    """Извлекает номер из заголовка вида '**1. Текст**'."""
    m = re.match(r'^(\d+)\.\s*(.*)', title)
    if m:
        return int(m.group(1)), m.group(2).rstrip('.')
    return None


def get_first_letter(name):
    """Определяет букву для алфавитной навигации."""
    if name:
        ch = name[0].upper()
        if ch in ALPHABET:
            return ch
    return "#"


def parse_tasks(md_text):
    """
    Разбивает Markdown на задачи.
    Ищет строки вида **N. Текст задачи** (номер + точка + текст, всё в жирном).
    Также отслеживает заголовки тем вида **Задачи к теме N**.

    Важное отличие от билетов: задачи внутри ответов содержат
    подпункты вида **2)** или **2.**, которые НЕ должны считаться
    отдельными задачами. Поэтому мы требуем:
      - номер должен быть строго больше предыдущего найденного
      - после номера обязателен текст длиной >= 30 символов
        (условие задачи всегда — длинная ситуация)
    """
    tasks = []
    lines = md_text.split("\n")
    i = 0
    current_topic = "Без темы"
    last_num = 0

    while i < len(lines):
        line = lines[i]
        stripped = line.strip()

        # Проверяем заголовок темы: **Задачи к теме ...**
        topic_match = re.match(r'^\*{2}\s*Задачи к теме\s+(.+?)\s*\*{2}', stripped)
        if topic_match:
            raw_topic = topic_match.group(1).strip()
            # Очищаем от OCR-артефактов: "5нноа" → "5", "б" → "6"
            raw_topic = re.sub(r'нноа|мот|щег|экран', '', raw_topic, flags=re.IGNORECASE)
            # Заменяем русскую «б» на «6» (частый OCR-артефакт)
            if raw_topic.strip() == 'б':
                raw_topic = '6'
            current_topic = clean_text(raw_topic.strip())
            i += 1
            continue

        # Проверяем задачу: **N. Текст**
        # Требуем: N > last_num (монотонно возрастающие),
        #          после N. идёт текст (а не скобка) и он достаточно длинный
        task_match = re.match(r'^\*{2}(\d+)\.\s*(.*?)\*{2}', stripped)
        if task_match:
            num = int(task_match.group(1))
            name_raw = task_match.group(2).strip()

            # Пропускаем если: номер не возрастает или текст слишком короткий
            if num <= last_num or len(name_raw) < 30:
                i += 1
                continue

            last_num = num

            # Собираем тело задачи — всё до следующей задачи или темы
            body_lines = []
            i += 1
            while i < len(lines):
                next_line = lines[i].strip()

                # Новая тема
                if re.match(r'^\*{2}\s*Задачи к теме', next_line):
                    break

                # Возможная новая задача — проверяем по тем же критериям
                next_task = re.match(r'^\*{2}(\d+)\.\s*(.*?)\*{2}', next_line)
                if next_task:
                    next_num = int(next_task.group(1))
                    next_name = next_task.group(2).strip()
                    if next_num > num and len(next_name) >= 30:
                        break

                # Стоп: начало дополнительного раздела без номеров
                # (в конце документа есть задачи с маркером «Условие:»)
                if re.match(r'^\d+\\?\.\s+[А-ЯЁ]', next_line) and i + 2 < len(lines):
                    # Проверяем: через 1-2 строки есть «Условие:»?
                    lookahead = "\n".join(lines[i+1:i+3])
                    if 'Условие:' in lookahead:
                        break

                body_lines.append(lines[i])
                i += 1

            body = "\n".join(body_lines).strip()
            name = clean_text(name_raw) if name_raw else f"Задача {num}"

            tasks.append({
                "num": num,
                "name": name,
                "body": body,
                "topic": current_topic
            })
        else:
            i += 1

    return tasks


# ─── Markdown → HTML ──────────────────────────────────────────────

def format_inline(text):
    """Обработка inline-разметки."""
    # Pandoc-артефакты
    text = re.sub(r'\[([^\]]*)\]\{\.underline\}', r'<em>\1</em>', text)
    text = re.sub(r'\[([^\]]*)\]\{\.mark\}', r'\1', text)
    text = re.sub(r'\{\.mark\}', '', text)
    text = re.sub(r'\{\.underline\}', '', text)
    text = re.sub(r'\{#[^}]*\}', '', text)
    text = re.sub(r'\{\.TOC-Heading\}', '', text)
    # Ссылки → просто текст
    text = re.sub(r'\[([^\]]*)\]\(#[^)]*\)', r'\1', text)
    text = re.sub(r'\[([^\]]*)\]\(https?://[^)]*\)', r'\1', text)
    # Форматирование
    text = re.sub(r'\*\*(.+?)\*\*', r'<b>\1</b>', text)
    text = re.sub(r'(?<!\*)\*(?!\*)(.+?)(?<!\*)\*(?!\*)', r'<em>\1</em>', text)
    text = re.sub(r'~~(.+?)~~', '', text)
    text = re.sub(r'`(.+?)`', r'<code>\1</code>', text)
    # Определения ||
    text = text.replace('||', '<span class="def">')
    if '<span class="def">' in text and '</span>' not in text:
        text += '</span>'
    # Тире
    text = text.replace('---', '\u2014')
    text = text.replace('--', '\u2013')
    return text


def is_table_separator(line):
    s = line.strip()
    if s.startswith('+') and re.match(r'^[+\-=+:]+$', s):
        return True
    if re.match(r'^\|[\s\-:|]+\|$', s):
        return True
    return False


def parse_table_cells(line):
    s = line.strip()
    if s.startswith('|'): s = s[1:]
    if s.endswith('|'): s = s[:-1]
    cells = [c.strip() for c in s.split('|')]
    return [c for c in cells if c]


def md_to_html(text):
    """Конвертация Markdown тела задачи в HTML."""
    lines = text.split("\n")
    out = []
    in_ul = False
    para = []
    in_table = False
    table_rows = []

    def flush_para():
        nonlocal para
        if para:
            c = " ".join(para)
            if c.strip(): out.append(f"<p>{c}</p>")
            para = []

    def close_ul():
        nonlocal in_ul
        if in_ul: out.append("</ul>"); in_ul = False

    def flush_table():
        nonlocal in_table, table_rows
        if table_rows:
            out.append('<div class="table-simple">')
            for ri, row in enumerate(table_rows):
                for cell in row:
                    cls = "table-header" if ri == 0 else "table-cell"
                    out.append(f'<p class="{cls}">{format_inline(cell)}</p>')
            table_rows = []
        in_table = False

    for line in lines:
        s = line.strip()

        if not s:
            flush_para(); close_ul(); flush_table(); continue

        if is_table_separator(s):
            flush_para(); close_ul(); in_table = True; continue

        if '|' in s and not is_table_separator(s) and in_table:
            flush_para(); close_ul()
            cells = parse_table_cells(s)
            if cells: table_rows.append(cells)
            continue

        if in_table: flush_table()

        # Цитата
        bq = re.match(r'^>\s*(.*)', s)
        if bq:
            flush_para(); close_ul()
            c = format_inline(bq.group(1))
            out.append(f'<p class="quote">{c}</p>' if c.strip() else '<br>')
            continue

        # Нумерованный список
        ol = re.match(r'^(\d+)\\?\.\s+(.*)', s)
        if ol:
            flush_para(); close_ul()
            out.append(f'<p class="list-item"><b>{ol.group(1)}.</b> {format_inline(ol.group(2))}</p>')
            continue

        # Маркированный список
        ul = re.match(r'^[-*]\s+(.*)', s)
        if ul:
            flush_para()
            if not in_ul: out.append("<ul>"); in_ul = True
            out.append(f"<li>{format_inline(ul.group(1))}</li>")
            continue

        # H2-H6
        h = re.match(r'^(#{2,6})\s+(.*)', s)
        if h:
            flush_para(); close_ul()
            lvl = min(len(h.group(1)), 6)
            out.append(f"<h{lvl}>{format_inline(h.group(2))}</h{lvl}>")
            continue

        # Строка с **Ответ** — выделяем визуально
        answer_match = re.match(r'^\*{2}Ответ[:\s]*(.*?)\*{0,2}$', s)
        if answer_match:
            flush_para(); close_ul()
            answer_text = answer_match.group(1).strip()
            if answer_text:
                out.append(f'<div class="answer-label">Ответ</div>')
                out.append(f'<p class="answer-text">{format_inline(answer_text)}</p>')
            else:
                out.append(f'<div class="answer-label">Ответ</div>')
            continue

        # Абзац
        para.append(format_inline(s))

    flush_para(); close_ul(); flush_table()
    return "\n".join(out)


# ─── CSS ───────────────────────────────────────────────────────────

def build_css():
    return f"""
* {{ margin:0; padding:0; box-sizing:border-box; }}
html {{ -webkit-text-size-adjust:100%; }}

body {{
    background:{COLOR_BG};
    color:{COLOR_TEXT};
    font-family:system-ui,-apple-system,sans-serif;
    font-size:{FONT_BODY};
    line-height:1.7;
    padding:24px {PADDING_SIDE};
    -webkit-tap-highlight-color:transparent;
    overflow-y:auto;
}}

/* Алфавитная панель */
.alpha-bar {{
    display:flex; flex-wrap:wrap; justify-content:center; gap:3px;
    padding:8px 0; margin-bottom:12px; border-bottom:1px solid {COLOR_BORDER};
}}
.alpha-bar a {{
    display:flex; align-items:center; justify-content:center;
    width:34px; height:34px; color:{COLOR_ACCENT}; text-decoration:none;
    font-size:14px; font-weight:bold; border-radius:50%;
    background:#111; border:1px solid {COLOR_BORDER};
}}
.alpha-bar a:active {{ background:{COLOR_ACCENT}; color:#000; }}
.alpha-bar a.no-items {{ opacity:0.2; pointer-events:none; }}

/* Разделитель секций */
.section-label {{
    color:{COLOR_ACCENT}; font-size:22px; font-weight:bold;
    text-align:center; padding:16px 0 10px 0;
    border-top:2px solid {COLOR_ACCENT};
    margin-top:20px; margin-bottom:10px;
}}

/* Заголовок темы в списке */
.topic-label {{
    color:{COLOR_TOPIC}; font-size:20px; font-weight:bold;
    text-align:center; padding:14px 0 8px 0;
    border-top:1px solid {COLOR_TOPIC};
    margin-top:16px; margin-bottom:8px;
}}

/* Буквенный якорь */
.letter-anchor {{
    padding:10px 0 6px 0; color:{COLOR_ACCENT}; font-size:24px;
    font-weight:bold; text-align:center; border-bottom:1px solid {COLOR_BORDER};
    margin:14px 0 10px 0;
}}

/* Ссылка-задача */
.task-link {{
    display:block; padding:16px 14px; margin-bottom:8px;
    color:#fff; text-decoration:none; font-size:{FONT_TASK};
    line-height:1.6; border-bottom:1px solid #1a1a1a; border-radius:10px;
    background:{COLOR_LINK_BG}; min-height:{TOUCH_TARGET};
}}
.task-link:active {{ background:{COLOR_LINK_ACTIVE}; }}
.task-link .num {{ color:{COLOR_ACCENT}; font-weight:bold; margin-right:10px; }}

h1 {{ font-size:{FONT_H1}; color:{COLOR_ACCENT}; text-align:center; padding:14px 0; line-height:1.3; word-wrap:break-word; }}
h2,h3,h4,h5,h6 {{ color:{COLOR_HEADER_SUB}; font-size:{FONT_H2}; text-align:center; margin:14px 0 10px 0; }}

p {{ margin-bottom:14px; text-align:left; word-wrap:break-word; }}
p.list-item {{ padding-left:16px; border-left:3px solid {COLOR_ACCENT}; margin-bottom:12px; }}
p.quote {{ padding-left:16px; border-left:3px solid #555; color:{COLOR_DIM}; margin-bottom:12px; font-size:17px; }}

ul,ol {{ margin:8px 0 12px 16px; }}
li {{ margin-bottom:8px; word-wrap:break-word; }}

.table-simple {{ margin:12px 0; padding:10px; border:1px solid {COLOR_BORDER}; border-radius:8px; background:#0a0a0a; }}
.table-header {{ color:{COLOR_ACCENT}; font-weight:bold; font-size:17px; margin-bottom:8px; text-align:center; border-bottom:1px solid {COLOR_BORDER}; padding-bottom:6px; }}
.table-cell {{ color:{COLOR_TEXT}; font-size:16px; margin-bottom:8px; padding-bottom:6px; border-bottom:1px solid #1a1a1a; }}

.def {{ color:{COLOR_DEF}; font-weight:bold; }}
b,strong {{ color:#fff; }}
em {{ color:{COLOR_HEADER_SUB}; font-style:italic; }}
code {{ color:#ce93d8; font-size:16px; }}

/* Блок «Ответ» — выделен визуально */
.answer-label {{
    color:{COLOR_ANSWER}; font-size:22px; font-weight:bold;
    text-align:center; padding:12px 0 6px 0;
    border-top:2px solid {COLOR_ANSWER};
    margin-top:18px; margin-bottom:8px;
}}
.answer-text {{
    border-left:4px solid {COLOR_ANSWER};
    padding-left:14px; margin-bottom:14px;
}}

/* Тег темы на странице задачи */
.topic-tag {{
    display:inline-block; background:#1a1a1a; color:{COLOR_TOPIC};
    font-size:16px; padding:4px 12px; border-radius:20px;
    margin:8px auto; text-align:center; border:1px solid {COLOR_TOPIC};
}}

.back-link {{
    display:block; text-align:center; padding:18px; margin-top:24px;
    color:{COLOR_ACCENT}; text-decoration:none; font-size:20px; font-weight:bold;
    border:1px solid {COLOR_BORDER}; border-radius:12px; background:#0a0a0a;
    min-height:{TOUCH_TARGET};
}}
.back-link:active {{ background:{COLOR_LINK_ACTIVE}; }}

.main-title {{ text-align:center; font-size:24px; color:{COLOR_ACCENT}; padding:12px 0 8px 0; }}
.task-count {{ text-align:center; color:#666; font-size:17px; margin-bottom:12px; }}
"""


# ─── Сборка HTML ───────────────────────────────────────────────────

def make_html(title, body_html):
    return f"""<!DOCTYPE html>
<html lang="ru">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1.0,maximum-scale=1.0,user-scalable=no">
<title>{html_mod.escape(title)}</title>
<style>{build_css()}</style>
</head>
<body>
{body_html}
</body>
</html>"""


def build_index(tasks, topic_groups, sorted_topics, letter_groups, sorted_letters, all_alpha_letters):
    """Генерирует главную страницу с темами + алфавитом + нумерованным списком."""

    # ── Заголовок ──
    body = f'<div class="main-title">{html_mod.escape(SITE_TITLE)}</div>\n'
    body += f'<div class="task-count">Всего задач: {len(tasks)}</div>\n'

    # ── Алфавитная панель ──
    alpha = '<nav class="alpha-bar">'
    for l in all_alpha_letters:
        if l in letter_groups:
            alpha += f'<a href="#letter-{l}">{l}</a>'
        else:
            alpha += f'<a class="no-items">{l}</a>'
    alpha += '</nav>'
    body += alpha

    # ── По номерам ──
    body += '<div class="section-label">По номерам</div>\n'
    for task in tasks:
        slug = f"z{task['num']:02d}"
        name = task["name"]
        if len(name) > 55: name = name[:52] + "..."
        body += f'<a href="q/{slug}.html" class="task-link"><span class="num">{task["num"]}.</span>{html_mod.escape(name)}</a>\n'

    # ── По темам ──
    body += '<div class="section-label">По темам</div>\n'
    for topic in sorted_topics:
        body += f'<div id="topic-{html_mod.escape(topic)}" class="topic-label">Тема {html_mod.escape(topic)}</div>'
        for task in topic_groups[topic]:
            slug = f"z{task['num']:02d}"
            name = task["name"]
            if len(name) > 55: name = name[:52] + "..."
            body += f'<a href="q/{slug}.html" class="task-link"><span class="num">{task["num"]}.</span>{html_mod.escape(name)}</a>\n'

    # ── По алфавиту ──
    body += '<div class="section-label">По алфавиту</div>\n'
    for letter in sorted_letters:
        body += f'<div id="letter-{letter}" class="letter-anchor">{letter}</div>'
        for task in letter_groups[letter]:
            slug = f"z{task['num']:02d}"
            name = task["name"]
            if len(name) > 55: name = name[:52] + "..."
            body += f'<a href="q/{slug}.html" class="task-link"><span class="num">{task["num"]}.</span>{html_mod.escape(name)}</a>\n'

    return make_html(SITE_TITLE, body)


def build_task(task):
    """Генерирует страницу задачи."""
    body = f"<h1>Задача {task['num']}. {html_mod.escape(task['name'])}</h1>\n"
    body += f'<div style="text-align:center"><span class="topic-tag">Тема {html_mod.escape(task["topic"])}</span></div>\n'
    body += md_to_html(task["body"])
    body += '\n<a href="../index.html" class="back-link">\u2190 К задачам</a>'
    return make_html(f"Задача {task['num']}", body)


# ─── Главный запуск ───────────────────────────────────────────────

def main():
    if len(sys.argv) < 2:
        print("Использование: python3 land-law-spy.py файл.docx [--output ./папка]")
        print()
        print("Примеры:")
        print("  python3 land-law-spy.py задачи.docx")
        print("  python3 land-law-spy.py задачи.docx --output ./land-output")
        sys.exit(1)

    docx_path = sys.argv[1]
    output_dir = "./land-law-output"

    if "--output" in sys.argv:
        idx = sys.argv.index("--output")
        if idx + 1 < len(sys.argv):
            output_dir = sys.argv[idx + 1]

    if not os.path.exists(docx_path):
        print(f"Файл не найден: {docx_path}")
        sys.exit(1)

    pages_dir = os.path.join(output_dir, "q")

    print(f"📄 Файл: {docx_path}")
    print(f"📁 Папка: {output_dir}")

    # 1. Pandoc
    print("⏳ Конвертация .docx → Markdown...")
    md_text = pandoc_to_markdown(docx_path)

    # 2. Парсинг задач
    print("⏳ Парсинг задач...")
    tasks = parse_tasks(md_text)
    if not tasks:
        print("❌ Не найдено пронумерованных задач (**N. Текст**).")
        print("   Проверьте формат документа.")
        sys.exit(1)

    print(f"✅ Найдено задач: {len(tasks)}")

    # 3. Группировка по темам
    topic_groups = {}
    for task in tasks:
        topic_groups.setdefault(task["topic"], []).append(task)

    # Сортировка тем: извлекаем число из "темы" если можно
    def topic_sort_key(topic):
        m = re.search(r'(\d+)', topic)
        return int(m.group(1)) if m else 999
    sorted_topics = sorted(topic_groups.keys(), key=topic_sort_key)

    # 4. Группировка по буквам
    letter_groups = {}
    for task in tasks:
        l = get_first_letter(task["name"])
        letter_groups.setdefault(l, []).append(task)

    sorted_letters = sorted(letter_groups.keys(),
        key=lambda x: ALPHABET.find(x) if x in ALPHABET else 99)

    all_alpha_letters = list(ALPHABET)

    # 5. Очистка и создание папок (не трогаем .git!)
    def safe_clean(directory):
        """Удаляет только сгенерированные файлы, не трогая .git и другое."""
        for item in os.listdir(directory):
            path = os.path.join(directory, item)
            if item == ".git":
                continue
            if os.path.isdir(path):
                shutil.rmtree(path, ignore_errors=True)
            else:
                try:
                    os.remove(path)
                except OSError:
                    pass

    if os.path.exists(output_dir):
        safe_clean(output_dir)
    os.makedirs(pages_dir, exist_ok=True)

    # 6. Генерация файлов
    print("⏳ Генерация HTML-файлов...")

    index_html = build_index(tasks, topic_groups, sorted_topics, letter_groups, sorted_letters, all_alpha_letters)
    with open(os.path.join(output_dir, "index.html"), "w", encoding="utf-8") as f:
        f.write(index_html)

    for task in tasks:
        page_html = build_task(task)
        filepath = os.path.join(pages_dir, f"z{task['num']:02d}.html")
        with open(filepath, "w", encoding="utf-8") as f:
            f.write(page_html)

    print(f"✅ Готово! Файлов: {len(tasks)} + index.html")
    print(f"📦 Размер: {os.path.getsize(output_dir + '/index.html') // 1024}KB + q/")
    print()
    print("🚀 Загрузи папку на GitHub Pages / Vercel / Netlify")
    print("📱 Открой index.html на часах")


if __name__ == "__main__":
    main()
