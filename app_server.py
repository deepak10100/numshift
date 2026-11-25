from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from email.parser import BytesParser
from email.policy import default
from io import BytesIO, StringIO
from html.parser import HTMLParser
from pathlib import Path
import csv, json, os, re, zipfile, xml.etree.ElementTree as ET, subprocess, tempfile, shutil, uuid
from urllib.parse import urlparse, parse_qs

ROOT = Path(__file__).resolve().parent
EXPORT_DIR = Path(os.environ.get('EXPORT_DIR', str(ROOT/'exports')))
EXPORTS_BY_TOKEN = {}
MAX_UPLOAD = 20 * 1024 * 1024

class TextHTMLParser(HTMLParser):
    def __init__(self):
        super().__init__(); self.parts=[]; self.title_parts=[]; self.h1_parts=[]; self.mode=[]; self.skip=0
    def handle_starttag(self, tag, attrs):
        if tag in ('script','style','noscript'): self.skip += 1
        if tag in ('title','h1'): self.mode.append(tag)
        if tag in ('p','div','br','li','tr','h1','h2','h3'): self.parts.append('\n')
    def handle_endtag(self, tag):
        if tag in ('script','style','noscript') and self.skip: self.skip -= 1
        if self.mode and self.mode[-1] == tag: self.mode.pop()
        if tag in ('p','div','li','tr','h1','h2','h3'): self.parts.append('\n')
    def handle_data(self, data):
        if self.skip: return
        self.parts.append(data)
        if self.mode and self.mode[-1] == 'title': self.title_parts.append(data)
        if self.mode and self.mode[-1] == 'h1': self.h1_parts.append(data)

def decode_text(data):
    try: return data.decode('utf-8-sig')
    except UnicodeDecodeError: return data.decode('latin-1', errors='replace')

def first_heading(text):
    for line in text.splitlines():
        line=line.strip()
        m=re.match(r'^#{1,6}\s+(.+?)\s*#*$', line)
        if m: return m.group(1).strip()
        if line.startswith('<h1'):
            m=re.search(r'<h1[^>]*>(.*?)</h1>', line, re.I)
            if m: return re.sub('<[^>]+>','',m.group(1)).strip()
    return ''

def first_title_candidate(text):
    """Fallback when metadata/explicit heading is absent: use a short first meaningful line."""
    for line in text.splitlines():
        line=re.sub(r'<[^>]+>', '', line).strip().lstrip('\ufeff')
        line=re.sub(r'^[-*+>#\s]+', '', line).strip()
        if not line or len(line)>120 or len(line)<3: continue
        if re.fullmatch(r'[\W_\d]+', line): continue
        return line
    return ''

def extract_upload(filename, data):
    ext=Path(filename).suffix.lower(); metadata_title=''; heading=''; content=''
    if ext in ('.txt','.md','.csv','.json','.html','.htm'):
        raw=decode_text(data)
        if ext in ('.html','.htm'):
            parser=TextHTMLParser(); parser.feed(raw)
            content=re.sub(r'\n[ \t]*\n+', '\n\n', ''.join(parser.parts)).strip()
            metadata_title=' '.join(' '.join(parser.title_parts).split())
            heading=' '.join(' '.join(parser.h1_parts).split())
        elif ext=='.csv':
            rows=list(csv.reader(StringIO(raw)))
            content='\n'.join('\t'.join(row) for row in rows)
        elif ext=='.json':
            obj=json.loads(raw)
            if isinstance(obj,dict): metadata_title=str(obj.get('title') or obj.get('name') or '')
            content=json.dumps(obj,ensure_ascii=False,indent=2)
        else:
            content=raw
            heading=first_heading(raw)
    elif ext=='.docx':
        with zipfile.ZipFile(BytesIO(data)) as z:
            doc=ET.fromstring(z.read('word/document.xml'))
            ns={'w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
            paras=[]; headings=[]
            for p in doc.findall('.//w:p',ns):
                value=''.join(t.text or '' for t in p.findall('.//w:t',ns)).strip()
                style=p.find('./w:pPr/w:pStyle',ns)
                style_name=style.attrib.get('{%s}val'%ns['w'],'').lower() if style is not None else ''
                if value:
                    paras.append(value)
                    if style_name.startswith('heading') or style_name in ('title','subtitle'): headings.append(value)
            content='\n'.join(paras)
            heading=headings[0] if headings else ''
            try:
                core=ET.fromstring(z.read('docProps/core.xml'))
                title_el=core.find('.//{http://purl.org/dc/elements/1.1/}title')
                metadata_title=(title_el.text or '').strip() if title_el is not None else ''
            except KeyError: pass
    elif ext=='.xlsx':
        from openpyxl import load_workbook
        wb=load_workbook(BytesIO(data),data_only=True,read_only=True)
        metadata_title=(wb.properties.title or '').strip()
        chunks=[]
        for ws in wb.worksheets:
            chunks.append(f'[{ws.title}]')
            for row in ws.iter_rows(values_only=True):
                vals=[str(v) if v is not None else '' for v in row]
                if any(vals): chunks.append('\t'.join(vals).rstrip())
        content='\n'.join(chunks)
        heading=wb.sheetnames[0] if wb.sheetnames else ''
        wb.close()
    elif ext=='.pdf':
        from pypdf import PdfReader
        reader=PdfReader(BytesIO(data))
        meta=reader.metadata
        metadata_title=(str(meta.title).strip() if meta and meta.title else '')
        if metadata_title.casefold() in ('untitled','untitled document'):
            metadata_title=''
        # Keep PDF extractor line breaks: they reflect the page's text layout.
        # pypdf's layout mode reconstructs line/column placement instead of
        # returning each positioned text fragment as an isolated line.
        content='\n\n'.join((page.extract_text(extraction_mode='layout') or '').strip() for page in reader.pages).strip()
        method='embedded text'
        # OCR only when the text layer is effectively empty. Short but valid selectable PDFs must not be overwritten by noisy OCR output.
        if len(re.sub(r'\s','',content)) < 4:
            if not shutil.which('pdftoppm') or not shutil.which('tesseract'):
                raise ValueError('This PDF appears scanned, but OCR tools are not available on the server.')
            pages_to_ocr=min(len(reader.pages),20)
            with tempfile.TemporaryDirectory(prefix='textreplace-ocr-') as tmp:
                pdf_path=os.path.join(tmp,'input.pdf'); prefix=os.path.join(tmp,'page')
                with open(pdf_path,'wb') as f: f.write(data)
                subprocess.run(['pdftoppm','-f','1','-l',str(pages_to_ocr),'-r','180','-png',pdf_path,prefix],check=True,stdout=subprocess.PIPE,stderr=subprocess.PIPE,timeout=90)
                images=sorted(Path(tmp).glob('page-*.png'))
                texts=[]
                for image in images:
                    result=subprocess.run(['tesseract',str(image),'stdout','-l','eng+spa','--psm','3'],check=True,stdout=subprocess.PIPE,stderr=subprocess.PIPE,timeout=45)
                    texts.append(result.stdout.decode('utf-8',errors='replace').strip())
                content='\n\n'.join(t.strip() for t in texts if t).strip()
                method='OCR (English + Spanish)'
        heading=first_heading(content)
    else:
        raise ValueError('Unsupported format. Supported: TXT, MD, CSV, JSON, HTML, PDF, DOCX, XLSX.')
    # Content-based first-line title inference is intentionally limited to PDFs.
    inferred_pdf_title = first_title_candidate(content) if ext=='.pdf' else ''
    title=metadata_title or heading or inferred_pdf_title or Path(filename).stem or 'Untitled document'
    result={'title':title,'content':content,'filename':Path(filename).name,'extension':ext.lstrip('.').upper()}
    if ext=='.pdf': result['method']=method
    return result

class Handler(SimpleHTTPRequestHandler):
    def __init__(self,*args,**kwargs): super().__init__(*args,directory=str(ROOT),**kwargs)
    def end_headers(self):
        self.send_header('Cache-Control','no-store, no-cache, must-revalidate, max-age=0')
        self.send_header('Pragma','no-cache')
        super().end_headers()
    def do_GET(self):
        if self.path.startswith('/api/download-export'):
            token=parse_qs(urlparse(self.path).query).get('token',[''])[0]
            path=EXPORTS_BY_TOKEN.get(token)
            if not token or not path or not path.is_file():
                self.send_error(404,'This export link is invalid or has expired.'); return
            payload=path.read_bytes()
            self.send_response(200); self.send_header('Content-Type','application/pdf'); self.send_header('Content-Disposition','inline; filename="numshift-result.pdf"'); self.send_header('Content-Length',str(len(payload))); self.end_headers(); self.wfile.write(payload); return
        # Never expose generated exports as public static files.
        if self.path.startswith('/exports/'):
            self.send_error(404); return
        super().do_GET()
    def do_POST(self):
        if self.path=='/api/export-pdf':
            self.do_export_pdf(); return
        if self.path!='/api/extract':
            self.send_error(404); return
        length=int(self.headers.get('Content-Length','0'))
        if length<=0 or length>MAX_UPLOAD:
            self.send_json(413,{'error':'File is empty or exceeds the 20 MB upload limit.'}); return
        body=self.rfile.read(length)
        wrapped=(f'Content-Type: {self.headers.get("Content-Type")}\r\nMIME-Version: 1.0\r\n\r\n').encode()+body
        try:
            msg=BytesParser(policy=default).parsebytes(wrapped)
            part=next((p for p in msg.iter_parts() if p.get_filename()),None)
            if part is None: raise ValueError('No file found in upload.')
            result=extract_upload(part.get_filename(),part.get_payload(decode=True) or b'')
            self.send_json(200,result)
        except ValueError as e:
            self.send_json(415,{'error':str(e)})
        except Exception as e:
            self.send_json(400,{'error':f'Could not read this file: {e}'})
    def do_export_pdf(self):
        length=int(self.headers.get('Content-Length','0'))
        if length<=0 or length>MAX_UPLOAD:
            self.send_json(413,{'error':'Export is empty or exceeds the size limit.'}); return
        try:
            request=json.loads(self.rfile.read(length).decode('utf-8'))
            text=str(request.get('content',''))
            if not text:
                self.send_json(400,{'error':'There is no result content to export.'}); return
            from reportlab.lib.pagesizes import A4
            from reportlab.lib.units import mm
            from reportlab.lib.styles import ParagraphStyle
            from reportlab.lib.enums import TA_LEFT
            from reportlab.pdfbase import pdfmetrics
            from reportlab.pdfbase.ttfonts import TTFont
            from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
            from xml.sax.saxutils import escape
            font_path='/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf'
            if 'UltraExportSans' not in pdfmetrics.getRegisteredFontNames():
                pdfmetrics.registerFont(TTFont('UltraExportSans',font_path))
            output=BytesIO()
            doc=SimpleDocTemplate(output,pagesize=A4,leftMargin=20*mm,rightMargin=20*mm,topMargin=25*mm,bottomMargin=25*mm,title='')
            style=ParagraphStyle('ExportText',fontName='UltraExportSans',fontSize=10.5,leading=15,alignment=TA_LEFT,spaceAfter=0,wordWrap='CJK')
            story=[]
            for line in text.replace('\\r\\n','\\n').replace('\\r','\\n').split('\\n'):
                if not line.strip(): story.append(Spacer(1,7))
                else: story.append(Paragraph(escape(line).replace('\\t','&nbsp;&nbsp;&nbsp;&nbsp;'),style))
            doc.build(story)
            payload=output.getvalue()
            EXPORT_DIR.mkdir(parents=True,exist_ok=True)
            token=uuid.uuid4().hex
            export_path=EXPORT_DIR/(token+'.pdf')
            export_path.write_bytes(payload)
            EXPORTS_BY_TOKEN[token]=export_path
            # Keep only a small number of recent, unguessable per-export links in memory.
            while len(EXPORTS_BY_TOKEN)>100:
                oldest=next(iter(EXPORTS_BY_TOKEN)); old_path=EXPORTS_BY_TOKEN.pop(oldest)
                try: old_path.unlink(missing_ok=True)
                except OSError: pass
            self.send_json(200,{'download':'/api/download-export?token='+token,'bytes':len(payload)})
        except Exception as e:
            self.send_json(500,{'error':f'Could not create PDF: {e}'})
    def send_json(self,status,obj):
        payload=json.dumps(obj,ensure_ascii=False).encode('utf-8')
        self.send_response(status); self.send_header('Content-Type','application/json; charset=utf-8'); self.send_header('Content-Length',str(len(payload))); self.end_headers(); self.wfile.write(payload)

if __name__=='__main__':
    os.chdir(ROOT)
    port=int(os.environ.get('PORT','4173'))
    ThreadingHTTPServer(('0.0.0.0',port),Handler).serve_forever()
