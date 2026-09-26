"""Export source formula crops and flag suspicious edges for visual review."""

import argparse
import json
import math
from pathlib import Path

import numpy as np
from PIL import Image, ImageDraw, ImageFont

from measure_math_geometry import ink_box, native_formula


def inspect_crop(source, box):
    if not isinstance(box, list) or len(box) != 4 or not all(isinstance(v, (int, float)) and math.isfinite(v) for v in box):
        raise ValueError(f"Invalid source box: {box}")
    x, y, w, h = map(int, box)
    if min(x, y) < 0 or min(w, h) <= 0 or x + w > source.width or y + h > source.height:
        raise ValueError(f"Out-of-bounds source box: {box}")
    crop = source.crop((x, y, x + w, y + h)).convert('RGB')
    ink = ink_box(np.asarray(crop))
    edges = []
    if ink:
        for edge, close in [('left', ink[0] <= 2), ('top', ink[1] <= 2),
                            ('right', ink[2] >= w - 2), ('bottom', ink[3] >= h - 2)]:
            if close:
                edges.append(edge)
    return crop, ink, edges


def export_review(run, output, selected=None):
    if output.exists():
        raise FileExistsError('Use a fresh review folder to preserve evidence')
    pages = [run / 'pages' / f'page_{n:03}' for n in selected] if selected else sorted((run / 'pages').glob('page_*'))
    if not pages:
        raise ValueError('No pages selected')
    output.mkdir(parents=True)
    (output / 'crops').mkdir()
    rows = []
    for page in pages:
        manifest = json.loads((page / 'manifest.json').read_text(encoding='utf-8-sig'))
        source = Image.open(page / 'source.png').convert('RGB')
        seen = set()
        for formula in manifest.get('formula_inventory', []):
            if not native_formula(formula):
                continue
            ident = formula['id']
            if ident in seen:
                raise ValueError(f'Duplicate formula id in {page.name}: {ident}')
            seen.add(ident)
            crop, ink, edges = inspect_crop(source, formula['box_px'])
            filename = f'crops/{page.name}_{len(seen):04}.png'
            crop.save(output / filename)
            rows.append({'page': int(page.name.split('_')[-1]), 'id': ident,
                         'box_px': formula['box_px'], 'latex': formula.get('latex', ''),
                         'crop': filename, 'ink': ink, 'edge_candidates': edges,
                         'needs_close_review': bool(edges) or ink is None})
    font = ImageFont.load_default(size=18)
    boards = []
    for start in range(0, len(rows), 12):
        canvas = Image.new('RGB', (2100, 1040), '#dddddd')
        draw = ImageDraw.Draw(canvas)
        for index, row in enumerate(rows[start:start + 12]):
            x = (index % 3) * 700
            y = (index // 3) * 260
            crop = Image.open(output / row['crop'])
            crop.thumbnail((680, 210))
            canvas.paste(crop, (x + 10, y + 40))
            color = '#b00020' if row['needs_close_review'] else '#000000'
            draw.text((x + 10, y + 5), f"p{row['page']:03} {row['id']}"[:58], font=font, fill=color)
        filename = f'board_{start // 12 + 1:03}.png'
        canvas.save(output / filename)
        boards.append(filename)
    report = {'formulas': len(rows), 'edge_or_empty_candidates': sum(r['needs_close_review'] for r in rows),
              'visual_review_required': True, 'boards': boards, 'items': rows,
              'note': 'Review every crop for complete symbols and neighboring-content contamination. Edge flags are candidates, not automatic failures or acceptance.'}
    (output / 'report.json').write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding='utf-8')
    return report


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--run', type=Path, required=True)
    parser.add_argument('--output', type=Path, required=True)
    parser.add_argument('--page', type=int, action='append', help='One-based source page; repeat for multiple pages; default all pages')
    args = parser.parse_args()
    result = export_review(args.run, args.output, args.page)
    print(f"formulas={result['formulas']} edge_or_empty_candidates={result['edge_or_empty_candidates']} review={args.output / 'report.json'}")
