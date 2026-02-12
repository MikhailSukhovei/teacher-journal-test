from __future__ import annotations

import math
import re
import zipfile
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field
from pathlib import Path

DOCX_PATH = Path("content/content.docx")
OUTPUT_ROOT = Path(".")
NEWS_PER_PAGE = 10

NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "pr": "http://schemas.openxmlformats.org/package/2006/relationships",
}

CYR_MAP = {
    "а": "a", "б": "b", "в": "v", "г": "g", "д": "d", "е": "e", "ё": "e", "ж": "zh",
    "з": "z", "и": "i", "й": "y", "к": "k", "л": "l", "м": "m", "н": "n", "о": "o",
    "п": "p", "р": "r", "с": "s", "т": "t", "у": "u", "ф": "f", "х": "h", "ц": "ts",
    "ч": "ch", "ш": "sh", "щ": "sch", "ъ": "", "ы": "y", "ь": "", "э": "e", "ю": "yu",
    "я": "ya",
}


@dataclass
class RawParagraph:
    style: str | None
    text: str
    image_rel_ids: list[str] = field(default_factory=list)


@dataclass
class NewsItem:
    index: int
    title: str
    body: str
    image_rel_ids: list[str]


@dataclass
class SiteData:
    home_title: str
    site_title: str
    site_header_title: str
    menu_items: list[str]
    home_content: str
    news_section_title: str
    news_items: list[NewsItem]
    section_pages: list["SectionPage"]


@dataclass
class SectionPage:
    title: str
    slug: str
    body: str


def transliterate(value: str) -> str:
    out: list[str] = []
    for ch in value.lower():
        out.append(CYR_MAP.get(ch, ch))
    return "".join(out)


def slugify(value: str, fallback: str) -> str:
    translit = transliterate(value)
    slug = re.sub(r"[^a-z0-9]+", "-", translit).strip("-")
    return slug or fallback


def paragraph_style(para: ET.Element) -> str | None:
    ppr = para.find("w:pPr", NS)
    if ppr is None:
        return None
    pstyle = ppr.find("w:pStyle", NS)
    if pstyle is None:
        return None
    return pstyle.attrib.get(f"{{{NS['w']}}}val")


def parse_docx(docx_path: Path) -> tuple[SiteData, dict[str, str], zipfile.ZipFile]:
    archive = zipfile.ZipFile(docx_path)
    document_root = ET.fromstring(archive.read("word/document.xml"))
    rels_root = ET.fromstring(archive.read("word/_rels/document.xml.rels"))

    rel_to_target: dict[str, str] = {}
    for rel in rels_root:
        rel_id = rel.attrib.get("Id")
        target = rel.attrib.get("Target")
        rel_type = rel.attrib.get("Type", "")
        if rel_id and target and rel_type.endswith("/image"):
            rel_to_target[rel_id] = target

    raw_paragraphs: list[RawParagraph] = []
    for para in document_root.findall(".//w:body/w:p", NS):
        style = paragraph_style(para)
        text = "".join(t.text or "" for t in para.findall(".//w:t", NS)).strip()
        rel_ids = [
            blip.attrib.get(f"{{{NS['r']}}}embed", "")
            for blip in para.findall(".//a:blip", NS)
            if blip.attrib.get(f"{{{NS['r']}}}embed")
        ]
        if text or rel_ids:
            raw_paragraphs.append(RawParagraph(style=style, text=text, image_rel_ids=rel_ids))

    if not raw_paragraphs:
        raise ValueError("В content.docx не найдено содержимое")

    home_title = raw_paragraphs[0].text

    menu_items: list[str] = []
    home_content_parts: list[str] = []
    news_section_title = "Новости"
    news_items: list[NewsItem] = []
    section_content_by_name: dict[str, list[str]] = {}

    current_section: str | None = None
    current_news_title: str | None = None
    current_news_parts: list[str] = []
    current_news_images: list[str] = []

    for paragraph in raw_paragraphs[1:]:
        if paragraph.style in {"2", "Heading2"} and paragraph.text:
            if current_news_title:
                news_items.append(
                    NewsItem(
                        index=len(news_items) + 1,
                        title=current_news_title,
                        body="\n\n".join(current_news_parts).strip(),
                        image_rel_ids=list(current_news_images),
                    )
                )
                current_news_title = None
                current_news_parts = []
                current_news_images = []

            current_section = paragraph.text.lower()
            if paragraph.text.lower() == "новости":
                news_section_title = paragraph.text
            elif paragraph.text.lower() not in {"меню", "контент"}:
                section_content_by_name.setdefault(current_section, [])
            continue

        if current_section == "меню":
            if paragraph.style in {"3", "Heading3"} and paragraph.text:
                menu_items.append(paragraph.text)
            continue

        if current_section == "контент":
            if paragraph.text:
                home_content_parts.append(paragraph.text)
            continue

        if current_section == "новости":
            if paragraph.style in {"3", "Heading3"} and paragraph.text:
                if current_news_title:
                    news_items.append(
                        NewsItem(
                            index=len(news_items) + 1,
                            title=current_news_title,
                            body="\n\n".join(current_news_parts).strip(),
                            image_rel_ids=list(current_news_images),
                        )
                    )
                    current_news_parts = []
                    current_news_images = []
                current_news_title = paragraph.text
                continue

            if current_news_title:
                if paragraph.text:
                    current_news_parts.append(paragraph.text)
                for rel_id in paragraph.image_rel_ids:
                    if rel_id not in current_news_images:
                        current_news_images.append(rel_id)
            continue

        if current_section and current_section not in {"меню", "контент", "новости"}:
            if paragraph.style in {"3", "Heading3"} and paragraph.text:
                section_content_by_name[current_section].append(f"### {paragraph.text}")
            elif paragraph.text:
                section_content_by_name[current_section].append(paragraph.text)

    if current_news_title:
        news_items.append(
            NewsItem(
                index=len(news_items) + 1,
                title=current_news_title,
                body="\n\n".join(current_news_parts).strip(),
                image_rel_ids=list(current_news_images),
            )
        )

    if not menu_items:
        menu_items = [news_section_title]

    title_lines = section_content_by_name.get("название", [])
    site_title = title_lines[0].strip() if title_lines else home_title
    site_header_title = "<br>".join(line.strip() for line in title_lines if line.strip()) or site_title

    section_pages: list[SectionPage] = []
    for item in menu_items:
        if slugify(item, "menu-item") == "novosti":
            continue
        section_key = item.lower()
        section_body = "\n\n".join(section_content_by_name.get(section_key, [])).strip()
        section_pages.append(
            SectionPage(
                title=item,
                slug=slugify(item, "section"),
                body=section_body,
            )
        )

    site_data = SiteData(
        home_title=home_title,
        site_title=site_title,
        site_header_title=site_header_title,
        menu_items=menu_items,
        home_content="\n\n".join(home_content_parts).strip(),
        news_section_title=news_section_title,
        news_items=news_items,
        section_pages=section_pages,
    )
    return site_data, rel_to_target, archive


def write_text(path: Path, content: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(content, encoding="utf-8")


def shell_quote(value: str) -> str:
    return value.replace('"', '&quot;')


def yaml_quote(value: str) -> str:
    return value.replace("\\", "\\\\").replace('"', '\\"')


def render_site(data: SiteData, rel_to_target: dict[str, str], archive: zipfile.ZipFile) -> None:
    config = f"""title: "{yaml_quote(data.site_title)}"
header_title: "{yaml_quote(data.site_header_title)}"
lang: ru
markdown: kramdown
collections:
  news:
    output: true
    permalink: /news/:name/
defaults:
  - scope:
      path: ""
      type: news
    values:
      layout: news_item
"""
    write_text(OUTPUT_ROOT / "_config.yml", config)

    menu_lines: list[str] = []
    for item in data.menu_items:
        slug = slugify(item, "menu-item")
        if slug == "novosti":
            menu_lines.append("  - label: \"Новости\"\n    url: /news/")
        else:
            menu_lines.append(f"  - label: \"{shell_quote(item)}\"\n    url: /{slug}/")
    write_text(OUTPUT_ROOT / "_data/menu.yml", "\n".join(menu_lines) + "\n")

    css = """body {
  margin: 0;
  font-family: Arial, sans-serif;
  color: #111;
  background: #f7f7f7;
}

.site-header {
  background: #fff;
  border-bottom: 1px solid #ddd;
}

.site-header__inner,
.page-content {
  max-width: 980px;
  margin: 0 auto;
  padding: 16px;
}

.site-title {
  margin: 0 0 12px;
}

.site-title a {
  color: inherit;
  text-decoration: none;
}

.menu {
  display: flex;
  gap: 12px;
  padding: 0;
  margin: 0;
  list-style: none;
}

.menu a {
  text-decoration: none;
  color: #005ea2;
  font-weight: 700;
}

.menu a.active {
  color: #111;
}

.content-block {
  background: #fff;
  border: 1px solid #ddd;
  padding: 20px;
}

.news-item {
  display: grid;
  grid-template-columns: 1fr 280px;
  gap: 16px;
  border-bottom: 1px solid #e5e5e5;
  padding: 14px 0;
}

.news-item:last-child {
  border-bottom: 0;
}

.news-item img {
  max-width: 100%;
  border-radius: 4px;
}

.pagination {
  display: flex;
  gap: 8px;
  margin-top: 20px;
  flex-wrap: wrap;
}

.pagination a,
.pagination span {
  border: 1px solid #ccc;
  padding: 6px 10px;
  text-decoration: none;
  color: #111;
}

.pagination .active {
  background: #111;
  color: #fff;
}

.news-hero img {
  width: 100%;
  max-height: 520px;
  object-fit: contain;
  background: #fafafa;
  border: 1px solid #ddd;
}

.news-gallery {
  margin-top: 18px;
}

.news-gallery__header {
  display: flex;
  align-items: center;
  justify-content: space-between;
  margin-bottom: 10px;
}

.news-gallery__controls {
  display: flex;
  gap: 8px;
}

.news-gallery__controls button {
  border: 1px solid #bbb;
  background: #fff;
  padding: 6px 10px;
  cursor: pointer;
}

.news-gallery__track {
  display: flex;
  gap: 12px;
  overflow-x: auto;
  scroll-behavior: smooth;
  padding-bottom: 6px;
}

.news-gallery__item {
  flex: 0 0 240px;
}

.news-gallery__item img {
  width: 100%;
  height: 150px;
  object-fit: cover;
  border: 1px solid #ddd;
}

@media (max-width: 720px) {
  .news-item {
    grid-template-columns: 1fr;
  }
}
"""
    write_text(OUTPUT_ROOT / "assets/css/site.css", css)

    base_layout = """<!doctype html>
<html lang="ru">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>{{ page.title }} | {{ site.title }}</title>
  <link rel="stylesheet" href="{{ '/assets/css/site.css' | relative_url }}">
</head>
<body>
  <header class="site-header">
    <div class="site-header__inner">
      <h1 class="site-title"><a href="{{ '/' | relative_url }}">{% if site.header_title %}{{ site.header_title }}{% else %}{{ site.title }}{% endif %}</a></h1>
      <nav>
        <ul class="menu">
          {% for item in site.data.menu %}
            {% assign is_active = false %}
            {% if page.url == item.url %}
              {% assign is_active = true %}
            {% elsif item.url != '/' and page.url contains item.url %}
              {% assign is_active = true %}
            {% endif %}
            <li><a href="{{ item.url | relative_url }}" class="{% if is_active %}active{% endif %}">{{ item.label }}</a></li>
          {% endfor %}
        </ul>
      </nav>
    </div>
  </header>
  <main class="page-content">
    {{ content }}
  </main>
</body>
</html>
"""
    write_text(OUTPUT_ROOT / "_layouts/base.html", base_layout)

    page_layout = """---
layout: base
---
<section class="content-block">
  <h2>{{ page.title }}</h2>
  {{ content }}
</section>
"""
    write_text(OUTPUT_ROOT / "_layouts/page.html", page_layout)

    news_list_layout = """---
layout: base
---
<section class="content-block">
  <h2>{{ page.title }}</h2>
  {% assign ordered_news = site.news | sort: 'order' %}
  {% assign page_news = ordered_news | slice: page.offset, 10 %}
  {% for news in page_news %}
    <article class="news-item">
      <div>
        <h3><a href="{{ news.url | relative_url }}">{{ news.title }}</a></h3>
        {% if news.excerpt_text %}<p>{{ news.excerpt_text }}</p>{% endif %}
      </div>
      {% if news.image %}
      <div>
        <img src="{{ news.image | relative_url }}" alt="{{ news.title }}">
      </div>
      {% endif %}
    </article>
  {% endfor %}

  {% if page.total_pages > 1 %}
  <nav class="pagination">
    {% if page.prev_url %}<a href="{{ page.prev_url | relative_url }}">Назад</a>{% endif %}
    {% for page_number in (1..page.total_pages) %}
      {% if page_number == page.current_page %}
        <span class="active">{{ page_number }}</span>
      {% elsif page_number == 1 %}
        <a href="{{ '/news/' | relative_url }}">{{ page_number }}</a>
      {% else %}
        <a href="{{ '/news/page/' | append: page_number | append: '/' | relative_url }}">{{ page_number }}</a>
      {% endif %}
    {% endfor %}
    {% if page.next_url %}<a href="{{ page.next_url | relative_url }}">Вперед</a>{% endif %}
  </nav>
  {% endif %}
</section>
"""
    write_text(OUTPUT_ROOT / "_layouts/news_list.html", news_list_layout)

    news_item_layout = """---
layout: base
---
<section class="content-block">
  <h2>{{ page.title }}</h2>
  {% assign gallery_images = page.images %}
  {% if gallery_images and gallery_images.size > 0 %}
    <p class="news-hero"><img src="{{ gallery_images[0] | relative_url }}" alt="{{ page.title }}"></p>
  {% elsif page.image %}
    <p class="news-hero"><img src="{{ page.image | relative_url }}" alt="{{ page.title }}"></p>
    {% assign gallery_images = page.image | split: '|' %}
  {% endif %}
  {{ content }}
  {% if gallery_images and gallery_images.size > 0 %}
    <section class="news-gallery">
      <div class="news-gallery__header">
        <strong>Галерея</strong>
        <div class="news-gallery__controls">
          <button type="button" class="news-gallery__prev" aria-label="Прокрутить влево">&lt;</button>
          <button type="button" class="news-gallery__next" aria-label="Прокрутить вправо">&gt;</button>
        </div>
      </div>
      <div class="news-gallery__track">
        {% for image in gallery_images %}
          <a class="news-gallery__item" href="{{ image | relative_url }}" target="_blank" rel="noopener">
            <img src="{{ image | relative_url }}" alt="{{ page.title }} {{ forloop.index }}">
          </a>
        {% endfor %}
      </div>
    </section>
    <script>
      (function () {
        var gallery = document.currentScript.closest('.news-gallery');
        if (!gallery) return;
        var track = gallery.querySelector('.news-gallery__track');
        var prev = gallery.querySelector('.news-gallery__prev');
        var next = gallery.querySelector('.news-gallery__next');
        var step = 260;
        prev.addEventListener('click', function () { track.scrollBy({ left: -step, behavior: 'smooth' }); });
        next.addEventListener('click', function () { track.scrollBy({ left: step, behavior: 'smooth' }); });
      })();
    </script>
  {% endif %}
</section>
"""
    write_text(OUTPUT_ROOT / "_layouts/news_item.html", news_item_layout)

    home_page = f"""---
layout: page
title: \"{shell_quote(data.home_title)}\"
permalink: /
---
{data.home_content}
"""
    write_text(OUTPUT_ROOT / "index.md", home_page)

    for section_page in data.section_pages:
        section_content = f"""---
layout: page
title: \"{shell_quote(section_page.title)}\"
permalink: /{section_page.slug}/
---
{section_page.body}
"""
        write_text(OUTPUT_ROOT / section_page.slug / "index.md", section_content)

    (OUTPUT_ROOT / "_news").mkdir(parents=True, exist_ok=True)
    (OUTPUT_ROOT / "assets/images/news").mkdir(parents=True, exist_ok=True)

    rendered_news = []
    for item in data.news_items:
        slug = slugify(item.title, f"news-{item.index}")
        image_paths: list[str] = []
        for idx, rel_id in enumerate(item.image_rel_ids, start=1):
            if rel_id not in rel_to_target:
                continue
            target = rel_to_target[rel_id].lstrip("/")
            source_name = f"word/{target}"
            ext = Path(target).suffix or ".jpg"
            file_name = f"{slug}{ext}" if idx == 1 else f"{slug}-{idx}{ext}"
            output_image = OUTPUT_ROOT / "assets/images/news" / file_name
            output_image.write_bytes(archive.read(source_name))
            image_paths.append(f"/assets/images/news/{file_name}")

        excerpt_text = item.body.split("\n\n", 1)[0].strip()
        image_path = image_paths[0] if image_paths else ""
        images_yaml = "".join(f'\n  - "{shell_quote(path)}"' for path in image_paths)
        content = f"""---
layout: news_item
title: \"{shell_quote(item.title)}\"
order: {item.index}
excerpt_text: \"{shell_quote(excerpt_text)}\"
image: \"{image_path}\"
images:{images_yaml if images_yaml else " []"}
---
{item.body}
"""
        write_text(OUTPUT_ROOT / "_news" / f"{slug}.md", content)
        rendered_news.append((item.index, item.title, slug))

    total_pages = max(1, math.ceil(len(rendered_news) / NEWS_PER_PAGE))
    for page_number in range(1, total_pages + 1):
        offset = (page_number - 1) * NEWS_PER_PAGE
        prev_url = ""
        next_url = ""
        if page_number > 1:
            prev_url = "/news/" if page_number == 2 else f"/news/page/{page_number - 1}/"
        if page_number < total_pages:
            next_url = f"/news/page/{page_number + 1}/"

        page_front_matter = f"""---
layout: news_list
title: \"{shell_quote(data.news_section_title)}\"
permalink: {'/news/' if page_number == 1 else f'/news/page/{page_number}/'}
current_page: {page_number}
total_pages: {total_pages}
offset: {offset}
prev_url: \"{prev_url}\"
next_url: \"{next_url}\"
---
"""
        target = OUTPUT_ROOT / "news" / "index.md" if page_number == 1 else OUTPUT_ROOT / "news" / "page" / str(page_number) / "index.md"
        write_text(target, page_front_matter)


def main() -> None:
    data, rel_to_target, archive = parse_docx(DOCX_PATH)
    try:
        render_site(data, rel_to_target, archive)
    finally:
        archive.close()
    print(
        "Готово: сгенерированы главная, "
        f"{len(data.news_items)} новостей и {len(data.section_pages)} дополнительных разделов."
    )


if __name__ == "__main__":
    main()
