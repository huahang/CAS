import hashlib
import html
import json
import pathlib
import re
import os
import shutil
import subprocess
import zipfile
import xml.etree.ElementTree as ET
from html.parser import HTMLParser

ROOT=pathlib.Path(__file__).resolve().parents[1]
OUT=ROOT/'核验'

class TextParser(HTMLParser):
    def __init__(self, selector=None):
        super().__init__(convert_charrefs=True)
        self.selector=selector;self.active=selector is None;self.depth=0;self.skip=0;self.pieces=[];self.done=False;self.root_tag=None
        self.tables=[]
    def handle_starttag(self, tag, attrs):
        classes=dict(attrs).get('class','').split()
        if self.selector and not self.active and not self.done and self.selector.lstrip('.') in classes:
            self.active=True;self.depth=1;self.root_tag=tag;return
        if self.active:
            if tag in ['table','tr','td','th']:
                self.tables.append((tag,dict(attrs).get('rowspan','1'),dict(attrs).get('colspan','1')))
            if self.selector and tag==self.root_tag:self.depth+=1
            if tag in ['script','style']:self.skip+=1
    def handle_endtag(self,tag):
        if not self.active:return
        if tag in ['script','style']:self.skip=max(0,self.skip-1)
        if self.selector and tag==self.root_tag:
            self.depth-=1
            if self.depth==0:self.active=False;self.done=True
    def handle_data(self,data):
        if self.active and not self.skip:self.pieces.append(data)

def text_of_html(raw,selector=None):
    parser=TextParser(selector);parser.feed(raw);return ''.join(parser.pieces)
def tables_of_html(raw,selector=None):
    parser=TextParser(selector);parser.feed(raw);return parser.tables
def compact(t):return re.sub(r'\s+','',t)
def mdtext(raw):
    raw=re.sub(r'<!--.*?-->','',raw,flags=re.S)
    raw=re.sub(r'^#{1,6} ','',raw,flags=re.M)
    raw=re.sub(r'\\([\\*_\[\]`])',r'\1',raw)
    return compact(text_of_html(raw))

def original_text(seg):
    source=ROOT/seg['source_path']
    if seg['type']=='pdf':
        raw=subprocess.check_output(['pdftotext','-layout',str(source),'-']).decode()
        pages=[];removed=[]
        for i,page in enumerate(raw.split('\f')):
            lines=page.splitlines(); nonempty=[j for j,l in enumerate(lines) if l.strip()]
            if not nonempty:continue
            last=nonempty[-1]
            if re.fullmatch(r'[-—]?\s*\d+\s*[-—]?',lines[last].strip()):
                removed.append(dict(page=i+1,text=lines[last].strip()));lines.pop(last)
            pages.append('\n'.join(lines).rstrip())
        assert removed==seg['removed_page_numbers'],(source,'页码记录不符')
        text='\n'.join(pages)
        correction=seg.get('reading_order_correction')
        if correction:
            assert text.count(correction['before'])==1
            text=text.replace(correction['before'],correction['after'])
        return text
    if seg['type']=='doc':
        antiword=os.environ.get('CAS_ANTIWORD') or shutil.which('antiword')
        if not antiword:raise RuntimeError('需要 antiword, 或用 CAS_ANTIWORD 指定其路径')
        return subprocess.check_output([antiword,'-m','UTF-8.txt','-w','0',str(source)]).decode()
    if seg['type']=='docx':
        with zipfile.ZipFile(source) as archive:
            tree=ET.fromstring(archive.read('word/document.xml'))
        ns={'w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
        assert not tree.findall('.//w:numPr',ns),(source,'需要单独核对自动编号')
        return ''.join(t.text or '' for t in tree.findall('.//w:body//w:t',ns))
    raise ValueError(seg['type'])

results=[]
records=json.loads((OUT/'正文清单.json').read_text())
for doc in records:
    reference=[];expected_tables=[]
    for seg in doc['sources']:
        if seg['type']=='html':
            raw=(ROOT/seg['source_path']).read_bytes().decode(seg['encoding'])
            title=text_of_html(raw,'.title_con').strip()
            assert compact(title)==compact(seg['title']),(doc['path'],'网页标题不符')
            text=title+text_of_html(raw,seg['selector'])
            expected_tables+=tables_of_html(raw,seg['selector'])
        else:
            text=original_text(seg)
            assert compact(text)==compact((ROOT/seg['extracted_path']).read_text()),(doc['path'],'重新读取原件与提取快照不符')
        c=compact(text)
        digest=hashlib.sha256(c.encode()).hexdigest()
        assert digest==seg['text_sha256'],(doc['path'],'独立源提取与编纂源不符',digest,seg['text_sha256'])
        reference.append(c)
    expected=''.join(reference)
    actual=mdtext((ROOT/doc['path']).read_text())
    assert tables_of_html((ROOT/doc['path']).read_text())==expected_tables,(doc['path'],'表格行列或合并单元格不符')
    success=actual==expected
    if not success:
        import difflib
        sm=difflib.SequenceMatcher(None,expected,actual,autojunk=False)
        changes=[(op,expected[a:b],actual[c:d]) for op,a,b,c,d in sm.get_opcodes() if op!='equal']
        print('差异',doc['path'],changes[:8])
    results.append(dict(path=doc['path'],characters=len(expected),source_sha256=hashlib.sha256(expected.encode()).hexdigest(),mirror_sha256=hashlib.sha256(actual.encode()).hexdigest(),逐字比较='一致' if success else '不一致'))

attachments=json.loads((ROOT/'附件/清单.json').read_text())['附件']
for a in attachments:
    data=(ROOT/a['path']).read_bytes()
    assert len(data)==a['bytes'] and hashlib.sha256(data).hexdigest()==a['sha256'],a['path']
    if a.get('archive_member'):
        unrar=os.environ.get('CAS_UNRAR') or shutil.which('unrar-free')
        if not unrar:raise RuntimeError('需要 unrar-free, 或用 CAS_UNRAR 指定其路径')
        member=subprocess.check_output([unrar,'-P','-i','/dev/null',str(ROOT/a['archive_path']),a['archive_member']])
        assert member==data,(a['path'],'与官方压缩包成员字节不符')
for page in json.loads((OUT/'网页清单.json').read_text()):
    if page.get('status')!=200:continue
    data=(ROOT/page['path']).read_bytes()
    assert len(data)==page['bytes'] and hashlib.sha256(data).hexdigest()==page['sha256'],page['path']

# 索引中的所有正文和附件都必须可达, 没有遗漏入库版本.
readme=(ROOT/'README.md').read_text()
local_links=set(re.findall(r'\]\(<([^>]+)>\)',readme))
assert all((ROOT/p).is_file() for p in local_links),'README本地链接失效'
assert {r['path'] for r in records} <= local_links,'README漏列正文'
assert {a['path'] for a in attachments} <= local_links,'README漏列附件'
standard_rows=[];interpretation_rows=[]
for line in readme.splitlines():
    if re.match(r'\| \d+ \|',line):
        cols=[c.strip() for c in line.strip('|').split('|')]
        if len(cols)==7:standard_rows.append(int(cols[0]))
        elif len(cols)==5:interpretation_rows.append(int(cols[0]))
assert standard_rows==list(range(43)),standard_rows
assert interpretation_rows==list(range(1,21)),interpretation_rows
assert {str(p.relative_to(ROOT)) for folder in ['准则','解释'] for p in (ROOT/folder).glob('*.md')} == {r['path'] for r in records},'正文清单与文件集合不一致'
expected_versions={('准则',0,2006),('准则',0,2014)}
expected_versions.update(('准则',n,2006) for n in range(1,39))
revisions={2:[2014],7:[2019],9:[2014],12:[2019],14:[2017],16:[2017],21:[2018],22:[2017],23:[2017],24:[2017],25:[2020],30:[2014,2026],33:[2014],37:[2014,2017],39:[2014],40:[2014],41:[2014],42:[2017]}
expected_versions.update(('准则',n,y) for n,years in revisions.items() for y in years)
expected_versions.update(('解释',r['number'],r['year']) for r in records if r['kind']=='解释')
assert {(r['kind'],r['number'],r['year']) for r in records}==expected_versions,'准则版本集合不完整'
assert len([r for r in records if r['kind']=='解释'])==20
for r in records:
    assert str(r['year'])==r['issue_date'][:4],(r['path'],'版本年份与印发年份不同')
    assert pathlib.Path(r['path']).name==f"{r['name']}({r['year']}年版).md",(r['path'],'名称不符合规范')
page_records={p['url']:p for p in json.loads((OUT/'网页清单.json').read_text())}
urls=set(re.findall(r'https?://[^\s)<>]+',readme))
assert all(u in page_records and page_records[u].get('status')==200 for u in urls),'README官网链接未全部核验'
for r in records:
    basis=page_records[r['version_basis_url']]
    original=(ROOT/basis['path']).read_bytes().decode(basis['encoding'])
    if r['number']==0 and r['kind']=='准则':
        title=text_of_html(original,'.title_con') or text_of_html(original)
        assert f"第{33 if r['year']==2006 else 76}号" in compact(title)
    else:
        assert compact(r['document_number']) in compact(text_of_html(original)),(r['path'],'文号依据不符')
    assert compact(r['issue_date_literal']) in compact(text_of_html(original)),(r['path'],'印发日期依据不符')

(OUT/'逐字核对结果.json').write_text(json.dumps(dict(核验日期='2026-10-01',方法='去除来源注释、Markdown标记、HTML标签及Unicode排版空白后, 按Unicode码点逐字比较; 不归一化文字、数字或标点',正文=results),ensure_ascii=False,indent=2)+'\n')
print('逐字一致',sum(r['逐字比较']=='一致' for r in results),'/',len(results),'附件哈希一致',len(attachments),'README独立官网URL',len(urls))
assert all(r['逐字比较']=='一致' for r in results)
