from pathlib import Path
from datetime import date,timedelta
import json,math,copy
from docx import Document
from docx.shared import Cm,Pt,RGBColor
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from PIL import Image,ImageDraw,ImageFont
R=Path(__file__).resolve().parents[2];O=R/'docs/阶段二';T=R/'tmp/t28word';T.mkdir(parents=True,exist_ok=True)
p=O/'T28_跟踪甘特图.docx';d=Document(p)
# Re-running is safe: reuse the original B source when the file already contains inserted images.
if d.inline_shapes:
 import subprocess,io
 d=Document(io.BytesIO(subprocess.check_output(['git','-c','safe.directory='+str(R),'show','e691e71:T28_跟踪甘特图.docx'])))
base=json.loads((O/'D阶段1交接数据.json').read_text(encoding='utf-8'))['baseline_activities'];L={a['id']:a for a in base};actual={}
for t in Document(O/'T28附_状态日活动实际进度方案.docx').tables:
 for r in t.rows:
  c=[v.text for v in r.cells]
  if len(c)==11 and c[0] in L:actual[c[0]]=c
weekly=[t for t in d.tables if len(t.columns)>20];assert len(weekly)==5
groups=[]
for t in weekly:
 ids=[r.cells[0].text for r in t.rows[2:]];groups.append(ids)
assert sum(len(x) for x in groups)==82 and len(set(q for g in groups for q in g))==82
start=date(2026,12,1);end=date(2027,5,31);status=date(2027,2,19);hol={date(2027,1,1),date(2027,4,5)}|{date(2027,2,i) for i in range(5,14)}|{date(2027,5,i) for i in range(1,6)}
def work(x):return x.weekday()<5 and x not in hol
def parse(s):
 m,k=map(int,s.split('/'));return date(2026 if m==12 else 2027,m,k)
font=lambda n:ImageFont.truetype('C:/Windows/Fonts/msyh.ttc',n)
f=font(28);small=font(24);W=2200;left=540;scale=13.2;rh=62;top=125
workdays=[start+timedelta(days=i) for i in range((end-start).days+1) if work(start+timedelta(days=i))]
from bisect import bisect_left
roles={'PM':'项目经理','B1':'技术负责人','B2':'后端二','FE':'前端','AI':'AI算法','QA':'测试','BA':'需求分析','EXT':'甲方'}
def panel(ids,num):
 im=Image.new('RGB',(W,top+len(ids)*rh+14),'white');dr=ImageDraw.Draw(im)
 def text(x,y,s,ft=f):dr.text((x,y),s,font=ft,fill='black')
 def xx(v):return left+bisect_left(workdays,v)*scale
 for i in range(0,len(workdays),10):
  x=left+i*scale;dr.line((x,85,x,im.height),fill='#dddddd');text(x,12,str(i),small);text(x,44,workdays[i].strftime('%m-%d'),small)
 text(10,72,'活动编码 / 岗位');text(350,72,'状态 / 进度')
 def bar(a,b,y,pct=None):
  x1=xx(a);x2=xx(b+timedelta(days=1));height=14
  dr.rectangle((x1,y,max(x1+2,x2),y+height),outline='black',fill='white',width=1)
  if pct is not None and pct>0:dr.rectangle((x1,y,x1+(x2-x1)*pct,y+height),fill='black')
 for j,aid in enumerate(ids):
  y=top+j*rh;loe=aid=='1.3.1-LOE';a=L[aid] if not loe else dict(name='持续进度成本与风险跟踪',role='PM',start=str(start),finish=str(end),d=119)
  text(10,y+12,aid+' '+a['role'])
  c=actual.get(aid);st=c[6] if c else '持续' if loe else '未开始';pct=c[7] if c else 'LOE' if loe else '0%';text(350,y,st,small);text(350,y+28,pct,small)
  ps=date.fromisoformat(a['start']);pf=date.fromisoformat(a['finish'])
  if a['d']==0:
   cx=xx(ps)+4;cy=y+13;dr.polygon([(cx-5,cy),(cx,cy-5),(cx+5,cy),(cx,cy+5)],outline='black',fill='white')
  else:bar(ps,pf,y+6)
  if loe:bar(start,status,y+38,1)
  elif c and c[4]!='—':
   ast=parse(c[4]);af=parse(c[5]) if c[5]!='—' else status;bar(ast,af,y+38,int(pct[:-1])/100)
   late=sum(work(pf+timedelta(days=i)) for i in range(1,(af-pf).days+1)) if af>pf else 0
   if late:text(xx(af)+scale+6,y+34,('已逾' if st=='进行中' else '晚')+str(late)+'工作日',small)
  elif ps<=status:text(xx(status)+scale+8,y+34,'应开未开',small)
  pass # Keep a plain white plotting area, without row rules.
 sx=xx(status+timedelta(days=1))
 for y in range(110,im.height,22):dr.line((sx,y,sx,y+13),fill='black',width=3)
 text(sx+8,82,'状态日 02/19',small)
 path=T/f'part{num}.png';im.save(path);return path
# Replace the five weekly P/A/X grids with eight daily bar panels.
body=d.element.body;first=next(p for p in d.paragraphs if p.text=='项目管理、需求与设计（1.x–3.x）')._p;last=next(p for p in d.paragraphs if p.text=='状态日活动实际记录')._p
nodes=list(body);sidx=nodes.index(first)
# Remove preceding manual page-break paragraph as well.
if sidx and nodes[sidx-1].xpath('.//w:br'):first=nodes[sidx-1];sidx-=1
idx=nodes.index(last)
for el in nodes[sidx:idx]:body.remove(el)
# Append paragraphs normally, then move before the actual-record heading.
def para(text='',style=None):
 pp=d.add_paragraph(text,style);last.addprevious(pp._p);return pp
names=['项目管理 需求与设计','开发实现一','开发实现二','开发实现三与测试','部署 移交与验收'];number=0
for label,ids in zip(names,groups):
 chunks=math.ceil(len(ids)/14);size=math.ceil(len(ids)/chunks)
 for k in range(0,len(ids),size):
  ch=ids[k:k+size];number+=1
  pp=para(f'跟踪甘特图 {number}  {label}','Heading 1');pp.paragraph_format.space_after=Pt(3);pp.paragraph_format.page_break_before=True
  pp=para('上条空心为T19原基准，下条为实际；进行中按完成比例填充。横轴为累计工作日及日期，已扣除周末与假日；虚线为2027-02-19结束。');pp.paragraph_format.space_after=Pt(3)
  pic=panel(ch,number);pp=para();pp.paragraph_format.space_after=Pt(0);pp.paragraph_format.line_spacing=1;pp.add_run().add_picture(str(pic),width=Cm(27.2))
  pp=para('活动全称见T15，实际起止和原因见后附38项记录表。未开始无实际条，进行中跨度不代表连续投入。');pp.paragraph_format.space_after=Pt(0)
# Start the original actual record table on a fresh page.
last.get_or_add_pPr().append(OxmlElement('w:pageBreakBefore'))
for pp in d.paragraphs:
 if pp.text.startswith('T28  |'):pp.text='T28  |  V2.2  |  2026-11-23  |  数据编制：刘毅辉（成员B）  分段绘图排版：邓皓天（成员D）'
 if pp.text.startswith('每个活动按日历周展示：'):pp.text='各活动以日历日期横条展示：上条为T19 V0.2原基准，下条为状态日前实际记录。进行中活动条内按完成比例填充；未开始活动不绘制实际条。虚线准确标出2027-02-19结束。横轴按累计工作日显示，扣除周末及法定假日，横条连续绘制。'
 if pp.text.startswith('后续页面按 WBS 阶段'):pp.text='以下按WBS阶段分成8幅横条图，覆盖81项原基准网络活动和1项独立LOE支持工作。各页重复活动表头、完整时间轴及图例。CR-01与更新基准不混入本图；具体实际日期和偏差原因保留原记录表。'
for t in d.tables:
 if t.cell(0,0).text=='版本':
  cs=t.add_row().cells
  for cc,v in zip(cs,['V2.1','2026-11-23','将P/A/X周表替换为8幅横条图；保留原始数据。','邓皓天（成员D）']):cc.text=v
for t in d.tables:
 if t.cell(0,0).text=='版本':
  cs=t.add_row().cells
  for cc,v in zip(cs,['V2.2','2026-11-23','改用连续黑白横条、累计工作日轴及简洁活动标签，清除逐日小格和底纹；数据不变。','邓皓天（成员D）']):cc.text=v
# Keep the opening page compact; panel legends already explain the grouping.
for pp in list(d.paragraphs):
 if pp.text=='图表分组说明' or pp.text.startswith('以下按WBS阶段分成8幅'):pp._p.getparent().remove(pp._p)
# Keep consistent black typography and prevent header-only table breaks.
for name in ['Title','Heading 1','Heading 2','Normal']:
 st=d.styles[name];st.font.color.rgb=RGBColor(0,0,0)
for el in list(d.styles.element.xpath('.//w:pBdr')):el.getparent().remove(el)
for t in d.tables:
 for ri,row in enumerate(t.rows):
  row._tr.get_or_add_trPr().append(OxmlElement('w:cantSplit'))
  for cell in row.cells:
   for pp in cell.paragraphs:
    if ri==0:pp.paragraph_format.keep_with_next=True
 for c in t.rows[0].cells:
  for pp in c.paragraphs:
   for rr in pp.runs:rr.bold=True
assert len(d.inline_shapes)==8
# Original record contents remain intact.
assert any(len(t.rows)==39 and len(t.columns)==8 for t in d.tables)
try:d.save(p);print('saved original')
except PermissionError:
 p=O/'T28_跟踪甘特图_连续条排版.docx';d.save(p);print('saved unlocked layout copy',p)
print('inserted',number,'panels; retained actual table, summary and source notes')
