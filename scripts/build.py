"""รวม src/app.src.html + src/parser.js (+ data/sample.json ถ้ามี) เป็น index.html ไฟล์เดียว"""
import json, pathlib
root = pathlib.Path(__file__).resolve().parent.parent
sample = root / 'data' / 'sample.json'
data = sample.read_text(encoding='utf-8') if sample.exists() else json.dumps({'meta': {}, 'factors': []})
if '</' in data:
    raise SystemExit('sample.json must not contain "</"')
html = (root / 'src' / 'app.src.html').read_text(encoding='utf-8')
html = html.replace('/*SAMPLE*/', data).replace('/*PARSER*/', (root / 'src' / 'parser.js').read_text(encoding='utf-8'))
head = ('<!doctype html><html lang="th"><head><meta charset="utf-8">'
        '<meta name="viewport" content="width=device-width,initial-scale=1,viewport-fit=cover">'
        '<style>body{margin:0}img{max-width:100%}[hidden]{display:none!important}</style></head><body>')
(root / 'index.html').write_text(head + html + '</body></html>', encoding='utf-8')
print('built index.html', 'with sample data' if sample.exists() else 'without sample data')
