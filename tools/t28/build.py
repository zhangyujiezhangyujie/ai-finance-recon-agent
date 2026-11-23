from pathlib import Path
from datetime import date,timedelta
import json,math
from collections import Counter
from docx import Document
from PIL import Image,ImageDraw,ImageFont
R=Path(__file__).resolve().parents[2];O=R/'docs/阶段二'
base=json.loads((O/'D阶段1交接数据.json').read_text(encoding='utf-8'))['baseline_activities']
actual={};budget={}
for t in Document(O/'T28附_状态日活动实际进度方案.docx').tables:
 for r in t.rows:
  c=[v.text for v in r.cells]
  if len(c)==11 and c[0] in {a['id'] for a in base}:actual[c[0]]=c;budget[c[0]]=(c[8],c[9])
bactual={}
for t in Document(O/'T28_跟踪甘特图.docx').tables:
 for r in t.rows:
  c=[v.text for v in r.cells]
  if len(c)==8 and c[0] in actual:bactual[c[0]]=c
assert len(base)==81 and len(actual)==len(bactual)==38
for k,c in actual.items():
 b=bactual[k];assert b[2:7]==[c[3],c[4],c[5],c[6],c[7]],k
counts=Counter(c[6] for c in actual.values());assert counts=={'已完成':28,'进行中':2,'未开始':8}
start=date(2026,12,1);end=date(2027,5,31);status=date(2027,2,19)
def parse(s):
 m,d=map(int,s.split('/'));return date(2026 if m==12 else 2027,m,d)
holidays={date(2027,1,1),date(2027,4,5)}|{date(2027,2,d) for d in range(5,14)}|{date(2027,5,d) for d in range(1,6)}
def work(d):return d.weekday()<5 and d not in holidays
def wd_late(a,b):return sum(work(a+timedelta(days=i)) for i in range(1,(b-a).days+1)) if b>a else 0
workdays=[start+timedelta(days=i) for i in range((end-start).days+1) if work(start+timedelta(days=i))]
from bisect import bisect_left
N=len(workdays);L=1560;unit=23;W=L+N*unit+80;top=410;rh=66
rows=sorted(base,key=lambda a:(a['start'],a['ES'],a['id']))
rows.insert(4,dict(id='1.3.1-LOE',name='持续进度成本与风险跟踪',role='PM',start=str(start),finish=str(end),d=119))
H=top+len(rows)*rh+650
im=Image.new('RGB',(W,H),'white');dr=ImageDraw.Draw(im)
def font(n):return ImageFont.truetype('C:/Windows/Fonts/msyh.ttc',n)
f=font(23);small=font(21);title=font(55);bold=font(28)
def text(x,y,s,ft=f):dr.text((x,y),s,font=ft,fill='black')
def xx(d):return L+bisect_left(workdays,d)*unit
text(35,25,'T28 状态日跟踪甘特图',title)
text(35,110,'原基准 T19 V0.2  ·  状态日 2027-02-19  ·  数据：B 的 T28 V2.0 与 T28附 V1.0  ·  绘图：成员D 2026-11-23',bold)
text(35,164,'81项网络活动：28项已完成 / 2项进行中 / 8项应开工但未开始 / 43项尚未到基准开始日；1.3.1为独立支持工作。',bold)
# Legend
for x,fill,label in [(35,'white','上条：原基准'),(390,'black','下条：实际已完成部分')]:
 dr.rectangle((x,221,x+110,239),fill=fill,outline='black',width=2);text(x+126,211,label)
dr.rectangle((1040,221,1150,239),outline='black',width=2);dr.rectangle((1040,221,1090,239),fill='black');text(1166,211,'进行中：条内按完成比例填充；不表示连续投入')
text(35,262,'菱形：零工期基准活动  |  横轴：累计工作日，扣除周末及假日  |  虚竖线：状态日结束  |  未开始活动无实际条')
# Sparse cumulative workday grid, matching the T19/T35 visual treatment.
for i in range(0,N,10):
 x=L+i*unit;dr.line((x,top-52,x,top+len(rows)*rh),fill='#dddddd',width=1);text(x,top-95,str(i),small);text(x,top-55,workdays[i].strftime('%m-%d'),small)
text(20,top-48,'活动编码',bold);text(230,top-48,'活动名称与岗位',bold);text(850,top-48,'状态 / 完成率',bold);text(1110,top-48,'基准 / 实际日期',bold)
roles={'PM':'项目经理','BA':'需求分析','B1':'技术负责人','B2':'后端（二）','FE':'前端','AI':'AI算法','QA':'测试','EXT':'甲方'}
def bar(a,b,y,progress=0,baseline=False,loe=False):
 x1=xx(a)+2;x2=xx(b)+unit-2
 x1=xx(a);x2=xx(b+timedelta(days=1))
 dr.rectangle((x1,y,max(x1+2,x2),y+13),fill='white',outline='black',width=1)
 if not baseline and progress>0:dr.rectangle((x1,y,x1+(x2-x1)*progress,y+13),fill='black')
for j,a in enumerate(rows):
 y=top+j*rh;pass # No dense row rules.
 name=a['name'].replace('实施与成果提交','').replace('服务实现与自测','（服务）').replace('界面联调与自测','（界面）')
 # Wrapping avoids truncating long activity names.
 text(20,y+8,a['id']);text(230,y+2,name[:23]);text(230,y+30,(name[23:]+'  ' if len(name)>23 else '')+roles.get(a['role'],a['role']),small)
 ps=date.fromisoformat(a['start']);pf=date.fromisoformat(a['finish']);c=actual.get(a['id']);loe=a['id']=='1.3.1-LOE'
 if loe:st='持续支持 LOE';pct='不设完成率';ast=start;af=status;progress=1
 elif c:
  st=c[6];pct=c[7];ast=parse(c[4]) if c[4]!='—' else None;af=parse(c[5]) if c[5]!='—' else status if ast else None;progress=int(pct[:-1])/100
 else:st='未开始（未到期）';pct='0%';ast=af=None;progress=0
 text(850,y+3,st);text(850,y+31,pct,small)
 text(1110,y+3,'基 '+ps.strftime('%m/%d')+'—'+pf.strftime('%m/%d'),small)
 text(1110,y+31,'实 '+ast.strftime('%m/%d')+'—'+('状态日' if st in ['进行中','持续支持 LOE'] else af.strftime('%m/%d')) if ast else '实 —',small)
 if not a['d'] and not loe:
  cx=xx(ps)+unit/2;cy=y+16;dr.polygon([(cx-7,cy),(cx,cy-7),(cx+7,cy),(cx,cy+7)],outline='black',fill='white')
 else:bar(ps,pf,y+8,baseline=True)
 if ast:
  bar(ast,af,y+34,progress,loe=loe)
  late=wd_late(pf,af)
  if late and not loe:text(xx(af)+unit+8,y+29,('逾基准 '+str(late)+'工作日' if st=='进行中' else '完成晚 '+str(late)+'工作日'),small)
 elif ps<=status:
  text(xx(status)+unit+12,y+29,'应开工未开始',small)
 # R06 absence marks for B2 task in progress, no inferred productive effort.
 if a['id']=='4.1.3-S':
  for d in [date(2027,2,k) for k in [1,2,3,4,15,16]]:
   q=xx(d);dr.line((q,y+32,q+unit,y+50),fill='black',width=2);dr.line((q,y+50,q+unit,y+32),fill='black',width=2)
# Status line uses end of Feb19, actual data stops here.
sx=xx(status+timedelta(days=1))
for y in range(top-2,top+len(rows)*rh,22):dr.line((sx,y,sx,y+13),fill='black',width=3)
text(sx+8,top-25,'状态日 02/19 结束',bold)
y=top+len(rows)*rh+35
text(35,y,'读图与核对说明',bold)
notes=[
 '1. 本图比较原T19与状态日前实绩；不含CR-01新增活动，不使用T35更新基准的预测日期。',
 '2. 实际条起止表示已记录的实际日期跨度；进行中活动截至状态日，填充仅表示68%或45%的完成比例，不能据此反算实际人天。',
 '3. 后端（二）在02/01—02/04、02/15—02/16缺勤，共6个工作日；4.1.3-S条中以叉号标注。OCR返工位于4.1.5-S，完成45%。',
 '4. 进行中条后的“逾基准”只表示状态日已超过原计划完成日的工作日数，不是最终工期延误预测；等待与返工原因见T28实际记录表。',
 '5. 1.3.1持续监控独立展示，不计入81项网络活动；25人天、37500元；状态日PV与EV均为16369.97元，延期不追加投入。',
 '6. 28项完成、2项进行中、8项应开未开及43项未来活动已逐项核对；全部38项实绩日期、状态和完成率与B的T28记录一致。',
 '7. 三个已完成里程碑：MS0 2026-12-01、MS1 2026-12-18、MS2 2027-01-08，均按原基准完成；后续里程碑不混用新版日期。',
 '8. 挣值口径：活动级EV155218.07元与F4的155218.75元相差0.68元，T29及预测仍按F4；本图只展示进度，不重新估算挣值。',
 'V0.2  2026-11-23  邓皓天：复用T19/T35的Pillow横条绘图方式，为B的T28补充完整基准与实际对照图。'
]
for s in notes:text(35,y+55,s,small);y+=45
out=O/'T28_跟踪甘特图.png';im.save(out,dpi=(300,300));print(out,im.size)
# Visual QA crops are temporary, not deliverables.
T=R/'tmp/t28';T.mkdir(exist_ok=True,parents=True)
im.resize((1400,round(H*1400/W))).save(T/'overview.png')
im.crop((0,0,W,950)).save(T/'top.png')
ix=next(i for i,a in enumerate(rows) if a['id']=='4.1.3-S');im.crop((0,top+(ix-2)*rh,W,top+(ix+7)*rh)).save(T/'delays.png')
