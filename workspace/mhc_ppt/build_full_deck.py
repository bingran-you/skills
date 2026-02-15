import re
import json
import hashlib
from pathlib import Path
from email import policy
from email.parser import BytesParser
from bs4 import BeautifulSoup, NavigableString
from PIL import Image

mhtml_files = [
    '刚刚，梁文锋署名，DeepSeek元旦新论文要开启架构新篇章.mhtml',
    '梁文锋DeepSeek新论文！接棒何恺明和字节，又稳了稳AI的“地基”.mhtml',
    '租了8张H100，他成功复现了DeepSeek的mHC，结果比官方报告更炸裂.mhtml'
]

base = Path('workspace/mhc_ppt')
extracted = base / 'mhtml_ordered'
extracted.mkdir(parents=True, exist_ok=True)

slides_dir = base / 'slides_full_integrated'
slides_dir.mkdir(parents=True, exist_ok=True)

assets_dir = base / 'assets_full_integrated'
assets_dir.mkdir(parents=True, exist_ok=True)

# Utilities

def sanitize(name: str) -> str:
    return re.sub(r'[\\/:*?"<>|]', '_', name)


def decode_html_part(part, html_bytes: bytes) -> str:
    charset = part.get_content_charset() or 'utf-8'
    html = html_bytes.decode(charset, errors='replace')
    m = re.search(r'charset=([\w-]+)', html, re.I)
    if m:
        meta_charset = m.group(1).lower()
        if meta_charset and meta_charset != charset.lower():
            try:
                html = html_bytes.decode(meta_charset, errors='replace')
            except Exception:
                pass
    return html


def save_image(content_location: str, content_type: str, data: bytes, out_dir: Path) -> Path:
    ext = None
    if content_type == 'image/png':
        ext = '.png'
    elif content_type == 'image/jpeg':
        ext = '.jpg'
    elif content_type == 'image/webp':
        ext = '.webp'
    else:
        ext = '.img'
    h = hashlib.md5(content_location.encode('utf-8')).hexdigest()[:16]
    out_path = out_dir / f"img_{h}{ext}"
    out_path.write_bytes(data)
    return out_path


def normalize_url(url: str) -> str:
    if not url:
        return url
    url = url.split('#')[0]
    url = url.split('?')[0]
    return url


def flatten_blocks(node):
    for child in node.children:
        if isinstance(child, NavigableString):
            continue
        if child.name in ['h2', 'h3', 'h4', 'p', 'blockquote', 'img']:
            yield child
        elif child.name in ['section', 'div', 'article']:
            yield from flatten_blocks(child)
        else:
            # handle other containers
            if hasattr(child, 'children'):
                yield from flatten_blocks(child)


def extract_elements(html: str, img_map: dict):
    soup = BeautifulSoup(html, 'html.parser')
    content = soup.find('div', id='js_content') or soup.find('div', class_='rich_media_content') or soup.body
    elements = []

    if not content:
        return elements

    for block in flatten_blocks(content):
        if block.name == 'img':
            src = block.get('data-src') or block.get('src') or ''
            path = img_map.get(src) or img_map.get(normalize_url(src))
            if path:
                elements.append({'type': 'image', 'src': str(path), 'caption': (block.get('data-caption') or '').strip()})
            continue

        if block.name in ['h2', 'h3', 'h4']:
            text = block.get_text(' ', strip=True)
            text = re.sub(r'\s+', ' ', text)
            if text:
                elements.append({'type': 'heading', 'text': text})
            continue

        if block.name in ['p', 'blockquote']:
            # handle mixed content with images
            if block.find('img'):
                # iterate children in order
                buf = []
                for child in block.children:
                    if isinstance(child, NavigableString):
                        t = str(child).strip()
                        if t:
                            buf.append(t)
                    elif getattr(child, 'name', None) == 'img':
                        # flush text before image
                        if buf:
                            text = re.sub(r'\s+', ' ', ' '.join(buf)).strip()
                            if text:
                                elements.append({'type': 'text', 'text': text})
                            buf = []
                        src = child.get('data-src') or child.get('src') or ''
                        path = img_map.get(src) or img_map.get(normalize_url(src))
                        if path:
                            caption = (child.get('data-caption') or '').strip()
                            elements.append({'type': 'image', 'src': str(path), 'caption': caption})
                    else:
                        t = child.get_text(' ', strip=True) if hasattr(child, 'get_text') else ''
                        if t:
                            buf.append(t)
                if buf:
                    text = re.sub(r'\s+', ' ', ' '.join(buf)).strip()
                    if text:
                        elements.append({'type': 'text', 'text': text})
            else:
                text = block.get_text(' ', strip=True)
                text = re.sub(r'\s+', ' ', text)
                if text:
                    elements.append({'type': 'text', 'text': text})
    return elements


def split_text(text, max_len=180):
    # split by punctuation while preserving meaning
    parts = re.split(r'([。！？；;])', text)
    chunks = []
    buf = ''
    for i in range(0, len(parts), 2):
        seg = parts[i].strip()
        punct = parts[i+1] if i+1 < len(parts) else ''
        if not seg:
            continue
        piece = (seg + punct)
        if len(buf) + len(piece) <= max_len:
            buf += piece
        else:
            if buf:
                chunks.append(buf)
            # if piece too long, hard split
            if len(piece) > max_len:
                for j in range(0, len(piece), max_len):
                    chunks.append(piece[j:j+max_len])
                buf = ''
            else:
                buf = piece
    if buf:
        chunks.append(buf)
    return chunks


def build_slides_from_elements(elements, article_label):
    slides = []
    current = []
    char_count = 0
    max_chars = 300

    def flush():
        nonlocal current, char_count
        if current:
            slides.append({'type': 'text', 'items': current.copy(), 'article': article_label})
            current = []
            char_count = 0

    for el in elements:
        if el['type'] == 'image':
            # flush current text first
            flush()
            slides.append({'type': 'image', 'src': el['src'], 'caption': el.get('caption') or '', 'article': article_label})
            continue

        if el['type'] == 'heading':
            # start new slide for headings to keep structure clear
            flush()
            current.append({'type': 'heading', 'text': el['text']})
            char_count += len(el['text']) + 10
            continue

        if el['type'] == 'text':
            chunks = split_text(el['text']) if len(el['text']) > 180 else [el['text']]
            for chunk in chunks:
                if char_count + len(chunk) > max_chars:
                    flush()
                current.append({'type': 'text', 'text': chunk})
                char_count += len(chunk) + 2

    flush()
    return slides


slide_files = []

# Global cover slide will be created separately

slide_index = 1

all_articles = []

for mhtml in mhtml_files:
    mhtml_path = Path(mhtml)
    slug = mhtml_path.stem
    article_dir = extracted / sanitize(slug)
    assets = article_dir / 'assets'
    assets.mkdir(parents=True, exist_ok=True)

    msg = BytesParser(policy=policy.default).parsebytes(mhtml_path.read_bytes())

    html_bytes = None
    html_part = None
    img_map = {}

    for part in msg.walk():
        ctype = part.get_content_type()
        if ctype == 'text/html' and html_bytes is None:
            html_bytes = part.get_payload(decode=True)
            html_part = part
        if ctype.startswith('image/'):
            loc = part.get('Content-Location')
            if not loc:
                continue
            data = part.get_payload(decode=True)
            if not data:
                continue
            out_path = save_image(loc, ctype, data, assets)
            img_map[loc] = out_path
            img_map[normalize_url(loc)] = out_path

    if html_bytes is None:
        continue

    html = decode_html_part(html_part, html_bytes)
    (article_dir / 'article.html').write_text(html, encoding='utf-8')

    soup = BeautifulSoup(html, 'html.parser')
    title = None
    if soup.title and soup.title.get_text(strip=True):
        title = soup.title.get_text(strip=True)
    h1 = soup.find('h1', id='activity-name')
    if h1 and h1.get_text(strip=True):
        title = h1.get_text(strip=True)
    if not title:
        title = slug

    author = None
    author_tag = soup.find(id='js_author_name')
    if author_tag:
        author = author_tag.get_text(strip=True)
    publish_time = None
    time_tag = soup.find(id='publish_time')
    if time_tag:
        publish_time = time_tag.get_text(strip=True)

    elements = extract_elements(html, img_map)
    all_articles.append({
        'title': title,
        'author': author,
        'time': publish_time,
        'elements': elements,
    })

# Copy all images to a unified assets directory for PPTX generation
for article in all_articles:
    for el in article['elements']:
        if el['type'] == 'image':
            src = Path(el['src'])
            # keep original filename to avoid collisions
            dest = assets_dir / src.name
            if src.suffix.lower() == '.webp':
                # convert webp to png for safer embedding
                dest = assets_dir / f"{src.stem}.png"
                if not dest.exists():
                    with Image.open(src) as im:
                        im.save(dest)
            else:
                if not dest.exists():
                    dest.write_bytes(src.read_bytes())
            # use relative path from slides directory
            el['src'] = f"../{assets_dir.name}/{dest.name}"

# Ensure cover image exists
cover_src = base / 'assets' / 'deepseek_banner.jpg'
cover_dest = assets_dir / 'deepseek_banner.jpg'
if cover_src.exists() and not cover_dest.exists():
    cover_dest.write_bytes(cover_src.read_bytes())

# Build slides list
slides = []

# Cover slide
slides.append({'type': 'cover'})

for article in all_articles:
    slides.append({'type': 'article_title', 'title': article['title'], 'author': article['author'], 'time': article['time']})
    slides.extend(build_slides_from_elements(article['elements'], article['title']))

# Generate HTML slides

def write_slide(path: Path, html: str):
    path.write_text(html, encoding='utf-8')
    return str(path)


def header_bar(title: str):
    return f"<div class=\"header\"><h1>{title}</h1></div>"

slide_num = 1
for s in slides:
    fname = f"slide{slide_num:03d}.html"
    out_path = slides_dir / fname

    if s['type'] == 'cover':
        html = f"""<!DOCTYPE html>
<html>
<head>
<style>
html {{ background: #ffffff; }}
body {{
  width: 720pt; height: 405pt; margin: 0; padding: 0;
  background: #181B24; font-family: Arial, sans-serif;
  display: flex;
}}
.hero {{ width: 720pt; height: 200pt; }}
.hero img {{ width: 720pt; height: 200pt; object-fit: cover; }}
.title-block {{ padding: 18pt 36pt 0 36pt; }}
.title-block h1 {{ color: #ffffff; font-size: 36pt; margin: 0 0 6pt 0; }}
.subtitle {{ color: #B165FB; font-size: 18pt; margin: 0 0 6pt 0; }}
.subtitle2 {{ color: #ffffff; font-size: 16pt; margin: 0 0 10pt 0; }}
.meta {{ color: #cccccc; font-size: 12pt; margin: 0; }}
.accent {{ height: 6pt; background: #B165FB; width: 720pt; }}
</style>
</head>
<body>
<div>
  <div class=\"hero\">
    <img src=\"../{assets_dir.name}/deepseek_banner.jpg\" onerror=\"this.style.display='none'\" />
  </div>
  <div class=\"accent\"></div>
  <div class=\"title-block\">
    <h1>DeepSeek mHC</h1>
    <p class=\"subtitle\">Manifold-Constrained Hyper-Connections</p>
    <p class=\"subtitle2\">基于 3 篇公众号文章完整整理</p>
    <p class=\"meta\">2026 年 1 月</p>
  </div>
</div>
</body>
</html>"""
    elif s['type'] == 'article_title':
        author = s.get('author') or '作者信息未展示'
        time = s.get('time') or '发布时间未展示'
        html = f"""<!DOCTYPE html>
<html>
<head>
<style>
html {{ background: #ffffff; }}
body {{
  width: 720pt; height: 405pt; margin: 0; padding: 0;
  background: #F4F6F6; font-family: Arial, sans-serif;
  display: flex;
}}
.header {{ background: #181B24; height: 54pt; display: flex; align-items: center; padding: 0 36pt; border-bottom: 4pt solid #B165FB; }}
.header h1 {{ color: #ffffff; font-size: 22pt; margin: 0; }}
.content {{ padding: 40pt 48pt; }}
.title {{ font-size: 24pt; color: #181B24; margin: 0 0 12pt 0; font-weight: bold; }}
.meta {{ font-size: 14pt; color: #333333; margin: 0 0 6pt 0; }}
</style>
</head>
<body>
<div class=\"slide\">
  {header_bar(s['title'])}
  <div class=\"content\">
    <p class=\"title\">{s['title']}</p>
    <p class=\"meta\">作者：{author}</p>
    <p class=\"meta\">发布时间：{time}</p>
  </div>
</div>
</body>
</html>"""
    elif s['type'] == 'text':
        items_html = []
        for item in s['items']:
            if item['type'] == 'heading':
                items_html.append(f"<p class=\"heading\">{item['text']}</p>")
            else:
                items_html.append(f"<p class=\"text\">{item['text']}</p>")
        content_html = "\n".join(items_html)
        html = f"""<!DOCTYPE html>
<html>
<head>
<style>
html {{ background: #ffffff; }}
body {{
  width: 720pt; height: 405pt; margin: 0; padding: 0;
  background: #F4F6F6; font-family: Arial, sans-serif;
  display: flex;
}}
.slide {{ width: 720pt; height: 405pt; }}
.header {{ background: #181B24; height: 54pt; display: flex; align-items: center; padding: 0 36pt; border-bottom: 4pt solid #B165FB; }}
.header h1 {{ color: #ffffff; font-size: 18pt; margin: 0; }}
.content {{ padding: 18pt 36pt 26pt 36pt; }}
.heading {{ font-size: 18pt; font-weight: bold; color: #181B24; margin: 0 0 8pt 0; }}
.text {{ font-size: 13pt; color: #333333; line-height: 1.5; margin: 0 0 8pt 0; }}
</style>
</head>
<body>
<div class=\"slide\">
  {header_bar(s['article'])}
  <div class=\"content\">
    {content_html}
  </div>
</div>
</body>
</html>"""
    elif s['type'] == 'image':
        src = s['src']
        caption = s.get('caption') or ''
        # decide layout based on aspect ratio
        width_pt = 620
        height_pt = 280
        try:
            with Image.open(src) as im:
                w, h = im.size
            if h > w * 1.2:
                width_pt = 320
                height_pt = 280
        except Exception:
            pass
        img_style = f"width: {width_pt}pt; height: {height_pt}pt; object-fit: contain;"
        cap_html = f"<p class=\"caption\">{caption}</p>" if caption else ""
        html = f"""<!DOCTYPE html>
<html>
<head>
<style>
html {{ background: #ffffff; }}
body {{
  width: 720pt; height: 405pt; margin: 0; padding: 0;
  background: #F4F6F6; font-family: Arial, sans-serif;
  display: flex;
}}
.slide {{ width: 720pt; height: 405pt; }}
.header {{ background: #181B24; height: 54pt; display: flex; align-items: center; padding: 0 36pt; border-bottom: 4pt solid #B165FB; }}
.header h1 {{ color: #ffffff; font-size: 18pt; margin: 0; }}
.content {{ padding: 18pt 24pt; display: flex; flex-direction: column; align-items: center; }}
.image-box {{ background: #ffffff; padding: 6pt; border: 1pt solid #DDDDDD; }}
.caption {{ font-size: 11pt; color: #666666; margin: 6pt 0 0 0; text-align: center; }}
</style>
</head>
<body>
<div class=\"slide\">
  {header_bar(s['article'])}
  <div class=\"content\">
    <div class=\"image-box\">
      <img src=\"{src}\" style=\"{img_style}\" />
    </div>
    {cap_html}
  </div>
</div>
</body>
</html>"""
    else:
        continue

    write_slide(out_path, html)
    slide_files.append(str(out_path))
    slide_num += 1

manifest = {
    'slides_dir': str(slides_dir),
    'slides_count': len(slide_files),
    'slides': slide_files,
}
(base / 'slides_full_integrated_manifest.json').write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding='utf-8')
print(f"Generated {len(slide_files)} slides -> {slides_dir}")
