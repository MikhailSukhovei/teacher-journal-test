from __future__ import annotations

import math
import re
import shutil
import zipfile
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field
from pathlib import Path

DOCX_PATH = Path("content/content.docx")
OUTPUT_ROOT = Path(".")
ITEMS_PER_PAGE = 10

NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
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
class SectionItem:
    index: int
    title: str
    body: str
    image_rel_ids: list[str] = field(default_factory=list)
    event_date: str = ""
    excerpt: str = ""
    slug: str = ""
    detail_url: str = ""
    image_paths: list[str] = field(default_factory=list)


@dataclass
class MenuSection:
    title: str
    slug: str
    items: list[SectionItem] = field(default_factory=list)


@dataclass
class SiteData:
    home_title: str
    site_title: str
    site_header_title_main: str
    site_header_title_sub: str
    site_footer_text: str
    site_logo_rel_id: str | None
    home_content_blocks: list[RawParagraph]
    menu_sections: list[MenuSection]
    home_preview_section_keys: list[str]


def transliterate(value: str) -> str:
    return "".join(CYR_MAP.get(ch, ch) for ch in value.lower())


def slugify(value: str, fallback: str) -> str:
    slug = re.sub(r"[^a-z0-9]+", "-", transliterate(value)).strip("-")
    return slug or fallback


def paragraph_style(para: ET.Element) -> str | None:
    ppr = para.find("w:pPr", NS)
    if ppr is None:
        return None
    pstyle = ppr.find("w:pStyle", NS)
    if pstyle is None:
        return None
    return pstyle.attrib.get(f"{{{NS['w']}}}val")


def normalize_section(name: str) -> str:
    return name.strip().lower()


def is_footer_section_key(section_key: str) -> bool:
    return "нижний" in section_key and ("колонтитул" in section_key or "колонтикул" in section_key)


def item_excerpt(body: str) -> str:
    first = body.split("\n\n", 1)[0].strip()
    if len(first) > 240:
        return first[:237].rstrip() + "..."
    return first


def parse_heading4_date(text: str) -> tuple[bool, str]:
    normalized = normalize_section(text)
    if not normalized.startswith("дата"):
        return False, ""
    value = text[len("Дата"):].strip(" :.-")
    return True, value


def build_items_for_section(section_title: str, section_slug: str, paragraphs: list[RawParagraph]) -> list[SectionItem]:
    items: list[SectionItem] = []
    current_title: str | None = None
    current_body: list[str] = []
    current_images: list[str] = []
    current_date = ""
    waiting_date_value = False

    def flush_item() -> None:
        nonlocal current_title, current_body, current_images, current_date, waiting_date_value
        if not current_title:
            return
        body = "\n\n".join(part for part in current_body if part).strip()
        idx = len(items) + 1
        title = current_title
        items.append(
            SectionItem(
                index=idx,
                title=title,
                body=body,
                image_rel_ids=list(current_images),
                event_date=current_date,
                excerpt=item_excerpt(body),
                slug=slugify(title, f"{section_slug}-item-{idx}"),
            )
        )
        current_title = None
        current_body = []
        current_images = []
        current_date = ""
        waiting_date_value = False

    for paragraph in paragraphs:
        if paragraph.style in {"3", "Heading3"} and paragraph.text:
            flush_item()
            current_title = paragraph.text
            continue

        if paragraph.style in {"4", "Heading4"} and paragraph.text:
            is_date_heading, inline_date = parse_heading4_date(paragraph.text)
            if is_date_heading:
                if inline_date:
                    current_date = inline_date
                    waiting_date_value = False
                else:
                    waiting_date_value = True
                continue

        if current_title is None:
            # Fallback: section without heading-3 still forms a single item.
            current_title = section_title

        if waiting_date_value and paragraph.text:
            current_date = paragraph.text
            waiting_date_value = False
            continue

        if paragraph.text:
            current_body.append(paragraph.text)
        for rel_id in paragraph.image_rel_ids:
            if rel_id not in current_images:
                current_images.append(rel_id)

    flush_item()
    return items


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
    title_lines: list[str] = []
    title_image_rel_ids: list[str] = []
    footer_lines: list[str] = []
    home_content_blocks: list[RawParagraph] = []
    section_paragraphs: dict[str, list[RawParagraph]] = {}
    section_titles: dict[str, str] = {}
    section_order: list[str] = []
    home_preview_section_keys: list[str] = []

    current_section: str | None = None
    current_h1: str | None = normalize_section(home_title)

    for paragraph in raw_paragraphs[1:]:
        if paragraph.text and paragraph.style in {"1", "Heading1"}:
            current_h1 = normalize_section(paragraph.text)

        if paragraph.text and paragraph.style in {"1", "Heading1", "2", "Heading2"}:
            section_key = normalize_section(paragraph.text)
            base_sections = {"меню", "контент", "название"}
            menu_key_set = {normalize_section(item) for item in menu_items}

            if (
                paragraph.style in {"2", "Heading2"}
                and current_h1 == normalize_section(home_title)
                and is_footer_section_key(section_key)
            ):
                current_section = "__footer__"
                continue

            if (
                paragraph.style in {"2", "Heading2"}
                and current_h1 == normalize_section(home_title)
                and section_key not in base_sections
                and not is_footer_section_key(section_key)
            ):
                if section_key not in home_preview_section_keys:
                    home_preview_section_keys.append(section_key)
                current_section = None
                continue

            is_base_section = paragraph.style in {"2", "Heading2"} and section_key in base_sections
            is_dynamic_section = False
            if paragraph.style in {"1", "Heading1"}:
                if menu_key_set:
                    is_dynamic_section = section_key in menu_key_set
                else:
                    is_dynamic_section = section_key != normalize_section(home_title) and section_key not in base_sections
            elif paragraph.style in {"2", "Heading2"} and section_key not in base_sections:
                # Backward compatibility: old content structure where section headers were H2.
                is_dynamic_section = section_key in menu_key_set

            if is_base_section or is_dynamic_section:
                current_section = section_key
                if section_key not in section_titles:
                    section_titles[section_key] = paragraph.text
                if section_key not in base_sections and section_key not in section_paragraphs:
                    section_paragraphs[section_key] = []
                    section_order.append(section_key)
                continue

        if current_section == "меню":
            if paragraph.style in {"3", "Heading3"} and paragraph.text:
                menu_items.append(paragraph.text)
            continue

        if current_section == "название":
            if paragraph.text:
                title_lines.append(paragraph.text)
            for rel_id in paragraph.image_rel_ids:
                if rel_id not in title_image_rel_ids:
                    title_image_rel_ids.append(rel_id)
            continue

        if current_section == "контент":
            home_content_blocks.append(paragraph)
            continue

        if current_section == "__footer__":
            if paragraph.text:
                footer_lines.append(paragraph.text)
            continue

        if current_section and current_section in section_paragraphs:
            section_paragraphs[current_section].append(paragraph)

    if not menu_items:
        menu_items = [section_titles[key] for key in section_order]

    normalized_title_lines = [line.strip() for line in title_lines if line.strip()]
    site_title = normalized_title_lines[0] if normalized_title_lines else home_title
    site_header_title_main = normalized_title_lines[0] if normalized_title_lines else site_title
    site_header_title_sub = "<br>".join(normalized_title_lines[1:]) if len(normalized_title_lines) > 1 else ""
    site_footer_text = "<br>".join(line.strip() for line in footer_lines if line.strip())
    site_logo_rel_id = title_image_rel_ids[0] if title_image_rel_ids else None

    menu_sections: list[MenuSection] = []
    for i, menu_name in enumerate(menu_items, start=1):
        section_key = normalize_section(menu_name)
        section_title = section_titles.get(section_key, menu_name)
        if section_key == "новости":
            section_slug = "news"
        else:
            section_slug = slugify(section_title, f"section-{i}")
        section_items = build_items_for_section(section_title, section_slug, section_paragraphs.get(section_key, []))
        menu_sections.append(MenuSection(title=section_title, slug=section_slug, items=section_items))

    site_data = SiteData(
        home_title=home_title,
        site_title=site_title,
        site_header_title_main=site_header_title_main,
        site_header_title_sub=site_header_title_sub,
        site_footer_text=site_footer_text,
        site_logo_rel_id=site_logo_rel_id,
        home_content_blocks=home_content_blocks,
        menu_sections=menu_sections,
        home_preview_section_keys=home_preview_section_keys,
    )
    return site_data, rel_to_target, archive


def write_text(path: Path, content: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(content, encoding="utf-8")


def remove_path(path: Path) -> None:
    if path.is_dir():
        shutil.rmtree(path)
    elif path.exists():
        path.unlink()


def shell_quote(value: str) -> str:
    return value.replace('"', '&quot;')


def yaml_quote(value: str) -> str:
    return value.replace("\\", "\\\\").replace('"', '\\"')


def build_items_yaml(items: list[SectionItem]) -> str:
    if not items:
        return "[]"
    lines: list[str] = []
    for item in items:
        lines.append(f'  - title: "{yaml_quote(item.title)}"')
        lines.append(f'    date: "{yaml_quote(item.event_date)}"')
        lines.append(f'    excerpt: "{yaml_quote(item.excerpt)}"')
        lines.append(f'    url: "{yaml_quote(item.detail_url)}"')
        first_image = item.image_paths[0] if item.image_paths else ""
        lines.append(f'    image: "{yaml_quote(first_image)}"')
    return "\n".join(lines)


def build_home_featured_yaml(
    featured_sections: list[tuple[str, list[tuple[str, str, str, str]]]]
) -> str:
    if not featured_sections:
        return "sections: []\n"

    lines = ["sections:"]
    for section_title, items in featured_sections:
        lines.append(f'  - title: "{yaml_quote(section_title)}"')
        lines.append("    items:")
        for title, event_date, image, url in items:
            lines.append(f'      - title: "{yaml_quote(title)}"')
            lines.append(f'        date: "{yaml_quote(event_date)}"')
            lines.append(f'        image: "{yaml_quote(image)}"')
            lines.append(f'        url: "{yaml_quote(url)}"')
    return "\n".join(lines) + "\n"


def build_home_content_markdown(
    blocks: list[RawParagraph],
    rel_to_target: dict[str, str],
    archive: zipfile.ZipFile,
) -> str:
    lines: list[str] = []
    home_image_dir = OUTPUT_ROOT / "assets" / "images" / "home"
    home_image_dir.mkdir(parents=True, exist_ok=True)
    image_counter = 1

    for block in blocks:
        if block.text:
            lines.append(block.text)
            lines.append("")
        for rel_id in block.image_rel_ids:
            target = rel_to_target.get(rel_id)
            if not target:
                continue
            normalized_target = target.lstrip("/")
            source_name = f"word/{normalized_target}"
            ext = Path(normalized_target).suffix or ".jpg"
            output_name = f"home-content-{image_counter}{ext}"
            output_path = home_image_dir / output_name
            output_path.write_bytes(archive.read(source_name))
            lines.append(f"![Иллюстрация](/assets/images/home/{output_name})")
            lines.append("")
            image_counter += 1

    while lines and not lines[-1].strip():
        lines.pop()
    return "\n".join(lines)


def render_site(data: SiteData, rel_to_target: dict[str, str], archive: zipfile.ZipFile) -> None:
    # Cleanup artifacts from the old \"news collection\" approach.
    remove_path(OUTPUT_ROOT / "_news")
    remove_path(OUTPUT_ROOT / "news")
    remove_path(OUTPUT_ROOT / "_layouts" / "news_list.html")
    remove_path(OUTPUT_ROOT / "_layouts" / "news_item.html")

    header_logo = ""
    if data.site_logo_rel_id and data.site_logo_rel_id in rel_to_target:
        target = rel_to_target[data.site_logo_rel_id].lstrip("/")
        source_name = f"word/{target}"
        ext = Path(target).suffix or ".png"
        logo_name = f"site-title-logo{ext}"
        logo_output = OUTPUT_ROOT / "assets" / "images" / logo_name
        logo_output.parent.mkdir(parents=True, exist_ok=True)
        logo_output.write_bytes(archive.read(source_name))
        header_logo = f"/assets/images/{logo_name}"

    config = f"""title: \"{yaml_quote(data.site_title)}\"
header_title_main: \"{yaml_quote(data.site_header_title_main)}\"
header_title_sub: \"{yaml_quote(data.site_header_title_sub)}\"
footer_text: \"{yaml_quote(data.site_footer_text)}\"
header_logo: \"{yaml_quote(header_logo)}\"
lang: ru
markdown: kramdown
"""
    write_text(OUTPUT_ROOT / "_config.yml", config)

    menu_lines: list[str] = []
    for section in data.menu_sections:
        menu_lines.append(f'  - label: "{shell_quote(section.title)}"\n    url: /{section.slug}/')
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

.site-title__link {
  display: flex;
  align-items: flex-start;
  gap: 12px;
}

.site-title__logo {
  width: 84px;
  height: 84px;
  object-fit: contain;
  flex: 0 0 auto;
  align-self: flex-start;
}

.site-title__text {
  line-height: 1.2;
}

.site-title__main {
  display: block;
}

.site-title__sub {
  display: block;
  margin-top: 4px;
  font-size: 0.82em;
  color: #666;
  line-height: 1.25;
}

.menu {
  display: flex;
  gap: 12px;
  padding: 0;
  margin: 0;
  list-style: none;
  flex-wrap: wrap;
}

.menu a {
  text-decoration: none;
  color: #005ea2;
  font-weight: 700;
}

.menu a.active {
  color: #111;
}

.site-footer {
  margin-top: 20px;
  border-top: 1px solid #ddd;
  background: #fff;
}

.site-footer__inner {
  max-width: 980px;
  margin: 0 auto;
  padding: 14px 16px 20px;
  font-size: 14px;
  color: #444;
  line-height: 1.35;
}

.content-block {
  background: #fff;
  border: 1px solid #ddd;
  padding: 20px;
}

.list-item {
  display: grid;
  grid-template-columns: 1fr 260px;
  gap: 14px;
  align-items: start;
  border-bottom: 1px solid #e5e5e5;
  padding: 14px 0;
}

.list-item:last-child {
  border-bottom: 0;
}

.list-item h3 {
  margin: 0 0 8px;
}

.list-item p {
  margin: 0;
}

.list-item__date {
  font-size: 12px;
  color: #333;
  padding: 4px 0;
  height: fit-content;
  min-height: 28px;
  text-align: left;
}

.list-item--with-date {
  display: grid;
  grid-template-columns: 140px 1fr;
  gap: 12px;
  align-items: start;
}

.list-item__content {
  text-align: left;
}

.list-item__image img {
  width: 100%;
  height: 160px;
  object-fit: cover;
  border: 1px solid #ddd;
}

.home-preview {
  margin-top: 16px;
}

.home-preview__grid {
  display: grid;
  grid-template-columns: repeat(3, minmax(0, 1fr));
  gap: 14px;
}

.home-preview__tile {
  border: 1px solid #ddd;
  background: #fff;
  padding: 10px;
}

.home-preview__tile-link {
  color: inherit;
  text-decoration: none;
  display: block;
}

.home-preview__image {
  width: 100%;
  height: 160px;
  object-fit: cover;
  border: 1px solid #ddd;
  background: #fafafa;
}

.home-preview__date {
  margin-top: 8px;
  font-size: 12px;
  color: #666;
}

.home-preview__title {
  margin-top: 6px;
  font-weight: 700;
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
  .list-item {
    grid-template-columns: 1fr;
  }

  .list-item--with-date {
    grid-template-columns: 1fr;
  }

  .home-preview__grid {
    grid-template-columns: 1fr;
  }

  .news-gallery__item {
    flex-basis: 200px;
  }
}
"""
    write_text(OUTPUT_ROOT / "assets/css/site.css", css)

    base_layout = """<!doctype html>
<html lang=\"ru\">
<head>
  <meta charset=\"utf-8\">
  <meta name=\"viewport\" content=\"width=device-width, initial-scale=1\">
  <title>{{ page.title }} | {{ site.title }}</title>
  <link rel=\"stylesheet\" href=\"{{ '/assets/css/site.css' | relative_url }}\">
</head>
<body>
  <header class=\"site-header\">
    <div class=\"site-header__inner\">
      <h1 class=\"site-title\">
        <a class=\"site-title__link\" href=\"{{ '/' | relative_url }}\">
          {% if site.header_logo and site.header_logo != '' %}
            <img class=\"site-title__logo\" src=\"{{ site.header_logo | relative_url }}\" alt=\"{{ site.title }}\">
          {% endif %}
          <span class=\"site-title__text\">
            <span class=\"site-title__main\">{% if site.header_title_main %}{{ site.header_title_main }}{% else %}{{ site.title }}{% endif %}</span>
            {% if site.header_title_sub and site.header_title_sub != '' %}
              <span class=\"site-title__sub\">{{ site.header_title_sub }}</span>
            {% endif %}
          </span>
        </a>
      </h1>
      <nav>
        <ul class=\"menu\">
          {% for item in site.data.menu %}
            {% assign is_active = false %}
            {% if page.url == item.url %}
              {% assign is_active = true %}
            {% elsif item.url != '/' and page.url contains item.url %}
              {% assign is_active = true %}
            {% endif %}
            <li><a href=\"{{ item.url | relative_url }}\" class=\"{% if is_active %}active{% endif %}\">{{ item.label }}</a></li>
          {% endfor %}
        </ul>
      </nav>
    </div>
  </header>
  <main class=\"page-content\">
    {{ content }}
  </main>
  {% if site.footer_text and site.footer_text != '' %}
  <footer class=\"site-footer\">
    <div class=\"site-footer__inner\">{{ site.footer_text }}</div>
  </footer>
  {% endif %}
</body>
</html>
"""
    write_text(OUTPUT_ROOT / "_layouts/base.html", base_layout)

    page_layout = """---
layout: base
---
<section class=\"content-block\">
  <h2>{{ page.title }}</h2>
  {{ content }}
</section>
"""
    write_text(OUTPUT_ROOT / "_layouts/page.html", page_layout)

    home_layout = """---
layout: base
---
<section class=\"content-block\">
  {{ content }}
</section>
{% assign featured_sections = site.data.home_featured.sections %}
{% if featured_sections and featured_sections.size > 0 %}
  {% for section in featured_sections %}
    <section class=\"content-block home-preview\">
      <h2>{{ section.title }}</h2>
      <div class=\"home-preview__grid\">
        {% for item in section.items %}
          <article class=\"home-preview__tile\">
            {% if item.url and item.url != '' %}<a class=\"home-preview__tile-link\" href=\"{{ item.url | relative_url }}\">{% endif %}
              {% if item.image and item.image != '' %}
                <img class=\"home-preview__image\" src=\"{{ item.image | relative_url }}\" alt=\"{{ item.title }}\">
              {% endif %}
              {% if item.date and item.date != '' %}<div class=\"home-preview__date\">{{ item.date }}</div>{% endif %}
              <div class=\"home-preview__title\">{{ item.title }}</div>
            {% if item.url and item.url != '' %}</a>{% endif %}
          </article>
        {% endfor %}
      </div>
    </section>
  {% endfor %}
{% endif %}
"""
    write_text(OUTPUT_ROOT / "_layouts/home.html", home_layout)

    menu_list_layout = """---
layout: base
---
<section class=\"content-block\">
  <h2>{{ page.title }}</h2>
  {% assign has_dates = false %}
  {% for i in page.items %}
    {% if i.date and i.date != '' %}
      {% assign has_dates = true %}
      {% break %}
    {% endif %}
  {% endfor %}

  {% for item in page.items %}
    <article class=\"list-item\">
      <div class=\"{% if has_dates %}list-item--with-date{% endif %}\">
        {% if has_dates %}
          <div class=\"list-item__date\">{{ item.date }}</div>
        {% endif %}
        <div class=\"list-item__content\">
          {% if item.url and item.url != '' %}
            <h3><a href=\"{{ item.url | relative_url }}\">{{ item.title }}</a></h3>
          {% else %}
            <h3>{{ item.title }}</h3>
          {% endif %}
          {% if item.image == '' or item.image == nil %}
            {% if item.excerpt != '' and item.excerpt != nil %}<p>{{ item.excerpt }}</p>{% endif %}
          {% endif %}
        </div>
      </div>
      {% if item.image and item.image != '' %}
        <div class=\"list-item__image\">
          <img src=\"{{ item.image | relative_url }}\" alt=\"{{ item.title }}\">
        </div>
      {% endif %}
    </article>
  {% endfor %}

  {% if page.total_pages > 1 %}
  <nav class=\"pagination\">
    {% if page.prev_url %}<a href=\"{{ page.prev_url | relative_url }}\">Назад</a>{% endif %}
    {% for page_number in (1..page.total_pages) %}
      {% if page_number == page.current_page %}
        <span class=\"active\">{{ page_number }}</span>
      {% elsif page_number == 1 %}
        <a href=\"{{ page.base_url | relative_url }}\">{{ page_number }}</a>
      {% else %}
        <a href=\"{{ page.base_url | append: 'page/' | append: page_number | append: '/' | relative_url }}\">{{ page_number }}</a>
      {% endif %}
    {% endfor %}
    {% if page.next_url %}<a href=\"{{ page.next_url | relative_url }}\">Вперед</a>{% endif %}
  </nav>
  {% endif %}
</section>
"""
    write_text(OUTPUT_ROOT / "_layouts/menu_list.html", menu_list_layout)

    menu_detail_layout = """---
layout: base
---
<section class=\"content-block\">
  <h2>{{ page.title }}</h2>
  {% assign gallery_images = page.images %}
  {% if gallery_images and gallery_images.size > 0 %}
    <p class=\"news-hero\"><img src=\"{{ gallery_images[0] | relative_url }}\" alt=\"{{ page.title }}\"></p>
  {% endif %}
  {{ content }}
  {% if gallery_images and gallery_images.size > 0 %}
    <section class=\"news-gallery\">
      <div class=\"news-gallery__header\">
        <strong>Галерея</strong>
        <div class=\"news-gallery__controls\">
          <button type=\"button\" class=\"news-gallery__prev\" aria-label=\"Прокрутить влево\">&lt;</button>
          <button type=\"button\" class=\"news-gallery__next\" aria-label=\"Прокрутить вправо\">&gt;</button>
        </div>
      </div>
      <div class=\"news-gallery__track\">
        {% for image in gallery_images %}
          <a class=\"news-gallery__item\" href=\"{{ image | relative_url }}\" target=\"_blank\" rel=\"noopener\">
            <img src=\"{{ image | relative_url }}\" alt=\"{{ page.title }} {{ forloop.index }}\">
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
    write_text(OUTPUT_ROOT / "_layouts/menu_detail.html", menu_detail_layout)

    home_content_markdown = build_home_content_markdown(data.home_content_blocks, rel_to_target, archive)

    home_page = f"""---
layout: home
title: \"{shell_quote(data.home_title)}\"
permalink: /
---
{home_content_markdown}
"""
    write_text(OUTPUT_ROOT / "index.md", home_page)

    (OUTPUT_ROOT / "assets/images/news").mkdir(parents=True, exist_ok=True)

    for section in data.menu_sections:
        for item in section.items:
            if not item.image_rel_ids:
                continue

            image_paths: list[str] = []
            for image_index, rel_id in enumerate(item.image_rel_ids, start=1):
                if rel_id not in rel_to_target:
                    continue
                target = rel_to_target[rel_id].lstrip("/")
                source_name = f"word/{target}"
                ext = Path(target).suffix or ".jpg"
                file_name = f"{section.slug}-{item.slug}-{image_index}{ext}"
                output_image = OUTPUT_ROOT / "assets/images/news" / file_name
                output_image.write_bytes(archive.read(source_name))
                image_paths.append(f"/assets/images/news/{file_name}")

            if not image_paths:
                continue

            item.image_paths = image_paths
            item.detail_url = f"/{section.slug}/{item.slug}/"
            images_yaml = "\n".join(f'  - "{shell_quote(path)}"' for path in image_paths)
            detail_content = f"""---
layout: menu_detail
title: \"{shell_quote(item.title)}\"
permalink: {item.detail_url}
images:
{images_yaml}
---
{item.body}
"""
            write_text(OUTPUT_ROOT / section.slug / item.slug / "index.md", detail_content)

    section_by_key = {normalize_section(section.title): section for section in data.menu_sections}
    featured_sections: list[tuple[str, list[tuple[str, str, str, str]]]] = []
    for section_key in data.home_preview_section_keys:
        section = section_by_key.get(section_key)
        if not section:
            continue
        preview_items: list[tuple[str, str, str, str]] = []
        for item in section.items[:6]:
            preview_image = item.image_paths[0] if item.image_paths else ""
            preview_items.append((item.title, item.event_date, preview_image, item.detail_url))
        if preview_items:
            featured_sections.append((section.title, preview_items))
    write_text(OUTPUT_ROOT / "_data" / "home_featured.yml", build_home_featured_yaml(featured_sections))

    for section in data.menu_sections:
        total_pages = max(1, math.ceil(len(section.items) / ITEMS_PER_PAGE))
        for page_number in range(1, total_pages + 1):
            offset = (page_number - 1) * ITEMS_PER_PAGE
            page_items = section.items[offset: offset + ITEMS_PER_PAGE]
            prev_url = ""
            next_url = ""
            base_url = f"/{section.slug}/"
            if page_number > 1:
                prev_url = base_url if page_number == 2 else f"/{section.slug}/page/{page_number - 1}/"
            if page_number < total_pages:
                next_url = f"/{section.slug}/page/{page_number + 1}/"

            page_front_matter = f"""---
layout: menu_list
title: \"{shell_quote(section.title)}\"
permalink: {base_url if page_number == 1 else f'/{section.slug}/page/{page_number}/'}
base_url: {base_url}
current_page: {page_number}
total_pages: {total_pages}
prev_url: \"{prev_url}\"
next_url: \"{next_url}\"
items:
{build_items_yaml(page_items)}
---
"""
            target = (
                OUTPUT_ROOT / section.slug / "index.md"
                if page_number == 1
                else OUTPUT_ROOT / section.slug / "page" / str(page_number) / "index.md"
            )
            write_text(target, page_front_matter)


def main() -> None:
    data, rel_to_target, archive = parse_docx(DOCX_PATH)
    try:
        render_site(data, rel_to_target, archive)
    finally:
        archive.close()

    total_items = sum(len(section.items) for section in data.menu_sections)
    detail_items = sum(1 for section in data.menu_sections for item in section.items if item.detail_url)
    print(
        "Готово: сгенерированы главная, "
        f"{len(data.menu_sections)} разделов меню, {total_items} элементов и {detail_items} детальных страниц."
    )


if __name__ == "__main__":
    main()
