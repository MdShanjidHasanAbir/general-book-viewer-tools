"""
Book Parser Script
==================
Parses a book .txt file and .json (table of contents) file,
produces a single output JSON that the HTML viewer can load directly.

Usage:
    python parser.py <book.txt> <toc.json> [-o output.json]

The .txt file uses these tags:
    <page_number>...</page_number>  - page boundaries and labels
    <heading>...</heading>          - section headings
    <note>...</note>                - footnotes
    <br>                            - line breaks

The .json file is an array of chapters:
    [{"chapter": "...", "sections": ["...", ...]}, ...]

Output JSON structure:
    {
        "toc": [...],           // table of contents (from input JSON)
        "pages": [              // parsed and formatted pages
            {
                "label": "67",
                "originalLabels": ["67"],
                "html": "<div class='page-text'>...</div>"
            },
            ...
        ]
    }
"""

import json
import os
import re
import sys
import argparse

from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side


def classify_line(line):
    """Classify a line as arabic, empty, or other based on character ratios."""
    stripped = re.sub(r'<[^>]+>', '', line).strip()
    if not stripped:
        return 'empty', line

    arabic_count = 0
    total = len(stripped)
    for ch in stripped:
        if '\u0600' <= ch <= '\u06FF' or '\u0750' <= ch <= '\u077F' or \
           '\uFB50' <= ch <= '\uFDFF' or '\uFE70' <= ch <= '\uFEFF':
            arabic_count += 1

    if arabic_count > 0 and arabic_count / total > 0.6:
        return 'arabic', line.strip()

    return 'other', line


def format_content(raw):
    """Convert raw page content into styled HTML."""
    html = raw

    # Convert <heading> tags
    def replace_heading(m):
        return f'<div class="heading">{m.group(1).strip()}</div>'

    html = re.sub(r'<heading>([\s\S]*?)</heading>', replace_heading, html, flags=re.IGNORECASE)
    # Keep headings separate so they never get merged into an arabic-text block.
    html = re.sub(
        r'\s*(<div class="heading">[\s\S]*?</div>)\s*',
        r'<br>\1<br>',
        html,
        flags=re.IGNORECASE,
    )

    # Split by <br> and classify lines
    lines = html.split('<br>')
    classified = [classify_line(line) for line in lines]

    # Merge consecutive arabic lines into one block
    processed = []
    arabic_buffer = []

    for line_type, content in classified:
        if line_type == 'arabic':
            arabic_buffer.append(content)
        else:
            if arabic_buffer:
                processed.append(f'<div class="arabic-text">{" ".join(arabic_buffer)}</div>')
                arabic_buffer = []
            processed.append(content)

    if arabic_buffer:
        processed.append(f'<div class="arabic-text">{" ".join(arabic_buffer)}</div>')

    html = '<br>'.join(processed)

    # Clean up excessive line breaks
    html = re.sub(r'(<br>\s*){3,}', '<br><br>', html)
    html = re.sub(r'^(<br>\s*)+', '', html)

    return f'<div class="page-text">{html}</div>'


def parse_book(text):
    """Parse the raw text file into section-based pages."""
    # Split by <page_number> tags
    raw_parts = re.split(r'<page_number>', text, flags=re.IGNORECASE)

    raw_pages = []
    for i in range(1, len(raw_parts)):
        chunk = raw_parts[i]
        close_idx = chunk.find('</page_number>')
        if close_idx == -1:
            continue

        page_num_raw = chunk[:close_idx].strip()
        content = chunk[close_idx + len('</page_number>'):].strip()

        # Extract page label (Bengali or ASCII digits)
        num_match = re.search(r'[০-৯\d]+', page_num_raw)
        page_label = num_match.group(0) if num_match else page_num_raw

        # Extract notes
        notes = []
        def collect_note(m):
            notes.append(m.group(1).strip())
            return ''

        content = re.sub(r'<note>([\s\S]*?)</note>', collect_note, content, flags=re.IGNORECASE)

        has_heading = bool(re.search(r'<heading>', content, re.IGNORECASE))

        raw_pages.append({
            'label': page_label,
            'fullLabel': page_num_raw,
            'content': content,
            'notes': notes,
            'hasHeading': has_heading,
        })

    # Group pages into sections (a section starts at a <heading> or at index 0)
    sections = []
    for i, page in enumerate(raw_pages):
        if page['hasHeading'] or i == 0:
            sections.append([page])
        else:
            sections[-1].append(page)

    # Build one combined page per section
    result = []
    for section_pages in sections:
        combined_html = ''
        all_notes = []
        raw_texts = []

        for pg in section_pages:
            combined_html += format_content(pg['content'])
            all_notes.extend(pg['notes'])
            raw_texts.append(pg['content'])

        # Append footnotes
        if all_notes:
            combined_html += (
                '<div class="section-footnotes">'
                '<div class="section-footnotes-label">তথ্যসূত্র</div>'
                f'<div class="note">{"<br>".join(all_notes)}</div>'
                '</div>'
            )

        # Page label: range if section spans multiple original pages
        first_label = section_pages[0]['label']
        last_label = section_pages[-1]['label']
        label = first_label if first_label == last_label else f'{first_label}-{last_label}'

        original_labels = [p['label'] for p in section_pages]

        # Extract plain text content (strip heading and HTML tags)
        raw_content = '\n'.join(raw_texts)
        raw_content_no_heading = re.sub(r'<heading>[\s\S]*?</heading>', '', raw_content, flags=re.IGNORECASE)
        plain_content = re.sub(r'<[^>]+>', '', raw_content_no_heading).strip()
        plain_content = re.sub(r'\n{3,}', '\n\n', plain_content)

        # Extract heading text
        heading_match = re.search(r'<heading>([\s\S]*?)</heading>', raw_content, re.IGNORECASE)
        section_heading = heading_match.group(1).strip() if heading_match else ''

        plain_notes = '\n'.join(all_notes)
        plain_notes = re.sub(r'<[^>]+>', '', plain_notes).strip()

        result.append({
            'label': label,
            'originalLabels': original_labels,
            'html': combined_html,
            'plainContent': plain_content,
            'plainNotes': plain_notes,
            'sectionHeading': section_heading,
        })

    return result


def auto_detect_files(directory):
    """Auto-detect .txt and .json files in the given directory."""
    import glob

    txt_files = glob.glob(os.path.join(directory, '*.txt'))
    json_files = [f for f in glob.glob(os.path.join(directory, '*.json'))
                  if not os.path.basename(f).startswith('book-output')]

    if not txt_files:
        print('Error: No .txt file found in the current folder.')
        sys.exit(1)
    if not json_files:
        print('Error: No .json file found in the current folder.')
        sys.exit(1)

    if len(txt_files) > 1:
        print('Multiple .txt files found:')
        for i, f in enumerate(txt_files, 1):
            print(f'  {i}. {os.path.basename(f)}')
        choice = input('Select .txt file number: ').strip()
        txt_file = txt_files[int(choice) - 1]
    else:
        txt_file = txt_files[0]

    if len(json_files) > 1:
        print('Multiple .json files found:')
        for i, f in enumerate(json_files, 1):
            print(f'  {i}. {os.path.basename(f)}')
        choice = input('Select .json file number: ').strip()
        json_file = json_files[int(choice) - 1]
    else:
        json_file = json_files[0]

    return txt_file, json_file


def generate_xlsx(toc_data, pages, output_path):
    """Generate an xlsx file mapping chapters -> sections -> content -> notes."""
    wb = Workbook()
    ws = wb.active
    ws.title = 'Book Content'

    # Header row
    headers = ['chapter_id', 'chapter_name', 'section_id', 'section_name', 'content', 'notes']
    header_font = Font(bold=True, size=12)
    header_fill = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')
    header_font_white = Font(bold=True, size=12, color='FFFFFF')
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin'),
    )

    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=header)
        cell.font = header_font_white
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border = thin_border

    # Map pages to chapters sequentially using TOC
    page_idx = 0
    row = 2
    for ch_idx, chapter in enumerate(toc_data):
        chapter_id = ch_idx + 1
        chapter_name = chapter['chapter']

        for sec_idx, section_name in enumerate(chapter['sections']):
            section_id = sec_idx + 1
            content = ''
            notes = ''

            if page_idx < len(pages):
                pg = pages[page_idx]
                content = pg['plainContent']
                notes = pg['plainNotes']
                page_idx += 1

            ws.cell(row=row, column=1, value=chapter_id).border = thin_border
            ws.cell(row=row, column=2, value=chapter_name).border = thin_border
            ws.cell(row=row, column=3, value=section_id).border = thin_border
            ws.cell(row=row, column=4, value=section_name).border = thin_border

            content_cell = ws.cell(row=row, column=5, value=content)
            content_cell.alignment = Alignment(wrap_text=True, vertical='top')
            content_cell.border = thin_border

            notes_cell = ws.cell(row=row, column=6, value=notes)
            notes_cell.alignment = Alignment(wrap_text=True, vertical='top')
            notes_cell.border = thin_border

            row += 1

    # Set column widths
    ws.column_dimensions['A'].width = 12
    ws.column_dimensions['B'].width = 35
    ws.column_dimensions['C'].width = 12
    ws.column_dimensions['D'].width = 40
    ws.column_dimensions['E'].width = 60
    ws.column_dimensions['F'].width = 40

    # Freeze header row
    ws.freeze_panes = 'A2'

    wb.save(output_path)
    return row - 2  # number of data rows


def main():
    parser = argparse.ArgumentParser(
        description='Parse book .txt and .json files into a single output JSON for the HTML viewer.'
    )
    parser.add_argument('txt_file', nargs='?', default=None, help='Path to the book .txt file')
    parser.add_argument('json_file', nargs='?', default=None, help='Path to the table of contents .json file')
    parser.add_argument('-o', '--output', default='book-output.json',
                        help='Output JSON file path (default: book-output.json)')

    args = parser.parse_args()

    # Auto-detect files if not provided
    if args.txt_file and args.json_file:
        txt_file = args.txt_file
        json_file = args.json_file
    else:
        script_dir = os.path.dirname(os.path.abspath(__file__))
        txt_file, json_file = auto_detect_files(script_dir)
        print(f'Auto-detected: {os.path.basename(txt_file)} + {os.path.basename(json_file)}')

    # Read input files
    with open(txt_file, 'r', encoding='utf-8') as f:
        text_data = f.read()

    with open(json_file, 'r', encoding='utf-8') as f:
        toc_data = json.load(f)

    # Parse the book
    pages = parse_book(text_data)

    # Build JSON output (strip extra fields not needed by the viewer)
    json_pages = []
    for pg in pages:
        json_pages.append({
            'label': pg['label'],
            'originalLabels': pg['originalLabels'],
            'html': pg['html'],
        })

    output = {
        'toc': toc_data,
        'pages': json_pages,
    }

    # Write JSON output
    with open(args.output, 'w', encoding='utf-8') as f:
        json.dump(output, f, ensure_ascii=False, indent=2)

    print(f'Parsed {len(pages)} pages from "{os.path.basename(txt_file)}"')
    print(f'TOC: {len(toc_data)} chapters from "{os.path.basename(json_file)}"')
    print(f'Output written to "{args.output}"')

    # Generate xlsx file
    xlsx_path = os.path.splitext(args.output)[0] + '.xlsx'
    xlsx_rows = generate_xlsx(toc_data, pages, xlsx_path)
    print(f'XLSX written to "{xlsx_path}" ({xlsx_rows} rows)')


if __name__ == '__main__':
    main()
