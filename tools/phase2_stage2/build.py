from pathlib import Path
import json,math,copy
from datetime import date,timedelta
from docx import Document
from docx.shared import Cm,Pt,RGBColor
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from PIL import Image,ImageDraw,ImageFont
R=Path.cwd();O=R/'docs/阶段二';T=R/'tmp/phase2_stage2';T.mkdir(parents=True,exist_ok=True)
x=json.loads((O/'D阶段1交接数据.json').read_text(encoding='utf-8'));base=x['baseline_activities'];actual={}
for tb in Document(O/'T28附_状态日活动实际进度方案.docx').tables:
 for row in tb.rows:
  c=[v.text for v in row.cells]
  if len(c)==11 and c[0] in {a['id'] for a in base}:actual[c[0]]=c
hol={date(2027,1,1),date(2027,4,5),date(2027,6,9)}|{date(2027,2,d) for d in range(5,14)}|{date(2027,5,d) for d in range(1,6)}
days=[];d=date(2026,12,1)
while len(days)<230:
 if d.weekday()<5 and d not in hol:days.append(d)
 d+=timedelta(days=1)
def ds(v,end=False):return str(days[max(0,math.ceil(round(v,6))-1 if end else math.floor(round(v,6)))])
def actualidx(v):
 m,day=map(int,v.split('/'));return days.index(date(2026 if m==12 else 2027,m,day))
A=[]
for b in base:
 a=copy.deepcopy(b);c=actual.get(a['id']);a['done']=bool(c and c[6]=='已完成');a['pct']=c[7] if c else '0%'
 if a['done']:
  a['fixed']=[actualidx(c[4]),actualidx(c[5])+1];a['dur']=a['fixed'][1]-a['fixed'][0];a['basis']='固定实绩';a['tri']='实绩'
 else:
  tri=[a['O'],a['M'],a['P']]
  if a['id']=='4.1.3-S':tri=[.5,1,2]
  if a['id']=='4.1.5-S':tri=[1,2.2,4]
  a['tri']='/'.join(map(str,tri));a['dur']=round((tri[0]+4*tri[1]+tri[2])/6,2) if a['role']!='EXT' else a['d'];a['basis']='剩余工程估算' if c and c[6]=='进行中' else '沿用原三点'
 A.append(a)
cr=[('3.1.2','两级权限与流程设计','FE',[1,1,2],[]),('4.6.1','两级授权返修','B2',[1,2,3],['CR01-01']),('4.4.4','两级状态机与接口','B1',[2,3,5],['CR01-02','4.4.4-S']),('4.4.4','两级工作台联调','FE',[1,2,3],['CR01-03','4.4.4-F']),('4.4.4','两级统计适配','B1',[.5,1,2],['CR01-03']),('5.2.1','两级链路功能回归','QA',[1,2,3],['CR01-04','CR01-05','5.2.1-A']),('5.3.2','两级安全留痕验证','QA',[.5,1,2],['CR01-06','5.3.2-A']),('5.4.1','验收脚本资料适配','BA',[.5,1,2],['CR01-07','7.1.1-A'])]
for i,(w,n,r,tri,pr) in enumerate(cr,1):A.append(dict(id=f'CR01-{i:02}',wbs=w,name=n,role=r,O=tri[0],M=tri[1],P=tri[2],tri='/'.join(map(str,tri)),dur=round((tri[0]+4*tri[1]+tri[2])/6,2),pred=pr,done=False,pct='0%',basis='CR01增量'))
L={a['id']:a for a in A}
L['4.1.3-S']['pred']+=['CR01-02'] # avoid B2 overlap by placing short permission repair first
L['4.1.3-S']['pred'].remove('CR01-02')
L['CR01-02']['pred']+=['4.1.3-S'];L['4.1.2-S']['pred']+=['CR01-02']
L['5.1.1-A']['pred']+=['CR01-04','CR01-05'];L['5.2.2-A']['pred']+=['CR01-06'];L['6.2.2-A']['pred']+=['CR01-07'];L['5.4.1-A']['pred']+=['CR01-08']
# Preserve frontend resource chain despite a completed successor in the actual record.
L['4.1.2-F']['pred']+=['4.1.3-F']
L['4.1.3-F']['pred']+=['CR01-01']
# CR QA work follows original QA tasks and is inserted before subsequent QA work.
def schedule(items):
 q=copy.deepcopy(items);lookup={a['id']:a for a in q};seq=[];pending=q[:]
 while pending:
  ready=[a for a in pending if all(p in {v['id'] for v in seq} for p in a['pred'])]
  if not ready:raise ValueError('cycle')
  for a in ready:
   a['ES']=a['fixed'][0] if a['done'] else max([52]+[lookup[p]['EF'] for p in a['pred']]);a['EF']=a['fixed'][1] if a['done'] else round(a['ES']+a['dur'],2);seq.append(a);pending.remove(a)
 finish=max(a['EF'] for a in q)
 for a in reversed(seq):
  succ=[b for b in q if a['id'] in b['pred'] and not b['done']];a['LF']=min([b['LS'] for b in succ],default=finish);a['LS']=round(a['LF']-a['dur'],2);a['TF']=round(a['LS']-a['ES'],2);a['FF']=round(min([b['ES'] for b in succ],default=finish)-a['EF'],2)
 return seq,finish
S,finish=schedule(A)

def new(title,code,ver='V2.0'):
 d=Document();s=d.sections[0];s.top_margin=s.bottom_margin=Cm(1.6);s.left_margin=s.right_margin=Cm(1.6)
 for sn in ['Normal','Title','Heading 1','Heading 2']:
  st=d.styles[sn];st.font.name='宋体';st._element.get_or_add_rPr().rFonts.set(qn('w:eastAsia'),'宋体');st.font.color.rgb=RGBColor(0,0,0);st.font.size=Pt(10 if sn=='Normal' else 18 if sn=='Title' else 12);st.paragraph_format.space_after=Pt(4)
 [el.getparent().remove(el) for el in list(d.styles.element.xpath('.//w:pBdr'))];d.add_paragraph(title,'Title');d.add_paragraph(f'{code}  {ver}  2026-11-22  邓皓天 成员D')
 return d

def p(d,s):d.add_paragraph(s)
def h(d,s):d.add_paragraph(s,'Heading 1')
def table(d,heads,rows):
 t=d.add_table(rows=1,cols=len(heads));t.style='Table Grid'
 for c,v in zip(t.rows[0].cells,heads):c.text=str(v)
 repeat=OxmlElement('w:tblHeader');t.rows[0]._tr.get_or_add_trPr().append(repeat)
 for row in rows:
  for c,v in zip(t.add_row().cells,row):c.text=str(v)
 for i,r in enumerate(t.rows):
  r._tr.get_or_add_trPr().append(OxmlElement('w:cantSplit'))
  for c in r.cells:
   for pp in c.paragraphs:
    for rr in pp.runs:rr.font.size=Pt(8);rr.font.bold=i==0
 return t

def rev(d,msg):h(d,'修订记录');table(d,['版本','日期','修订人','主要变更'],[['V2.0','2026-11-22','邓皓天',msg]])

# Approved combination 1; no scope reduction.
from collections import defaultdict
from decimal import Decimal,ROUND_HALF_UP
import html
SOURCE_URL='https://www.forestry.gov.cn/lyj/1/szxx/20261105/695102.html'
def cash(v):return float(Decimal(str(v)).quantize(Decimal('.01'),rounding=ROUND_HALF_UP))
def money(v):return f'{v:,.2f}'
crash=[]
for a in A:
 if not a['done'] and a['role'] in ['B1','B2','FE'] and a['dur']>=2:
  cut=round(a['dur']*.25,2);crash.append(dict(id=a['id'],wbs=a['wbs'],role=a['role'],original=a['dur'],cut=cut,after=round(a['dur']-cut,2),added_effort=round(cut*1.5,3),cost=cash(cut*2250)));a['dur']=round(a['dur']-cut,2)
lk={a['id']:a for a in A};lk['5.3.1-A']['pred']=['5.1.1-A'];lk['5.3.1-A']['role']='QA2';lk['5.3.1-A']['dur']=6.20
lk['5.3.2-A']['pred']=list(dict.fromkeys(lk['5.3.2-A']['pred']+['5.2.2-A','CR01-06']))
S,finish=schedule(A);L={a['id']:a for a in S};rem=[a for a in S if not a['done']]
assert ds(L['7.3.2-A']['EF'],True)=='2027-06-08' and ds(finish,True)=='2027-06-11'
assert math.ceil(finish)-119==8
for i,a in enumerate(rem):
 for b in rem[i+1:]:assert a['role']!=b['role'] or min(a['EF'],b['EF'])-max(a['ES'],b['ES'])<.001
critical=[a['id'] for a in rem if a['TF']==0]
ms={'MS3':ds(max(L[k]['EF'] for k in ['4.4.4-F','CR01-04','CR01-05']),True),'MS4':ds(max(L[k]['EF'] for k in ['5.3.2-A','CR01-07']),True),'MS5':ds(L['5.4.2-A']['EF'],True),'MS6':ds(L['7.3.2-A']['EF'],True),'项目结束':ds(finish,True)}
assert ms['MS4']=='2027-05-10'
intro='状态日2027-02-19；以已批准组合1 赶工与测试快速跟进为唯一更新进度基准。范围不变，保留4.5，不开CR-02。项目经理于2026-11-22选定，项目发起人（模拟）批准进度基准变更，延期8个工作日。'
cal='日历沿用项目假设表B3/B4，每周五天，调休周末不工作；增补2027-06-09端午节，放假1天。依据国办发明电〔2026〕10号（2026-11-05），国家林草局转引中国政府网：'+SOURCE_URL+'。累计工作日映射取整，不对每个活动单独取整；原项目窗口119工作日，新结束映射第127工作日。'
resource='1.3.1延期监控投入不追加，原25人天及37500元保持不变，延长期在原投入内统筹。必须落实的资源：借调1名测试人员QA2，承担5.3.1-A回归，原QA执行功能、安全与部署验证，最多两条测试链并行。回归原5.17人天已在原人工预算内，20%返工追加1.034人天即1551.00元，工期6.20日；基础工作转岗不重复计费。后端（二）、技术负责人、前端须按逐活动追加人天落实赶工补人；赶工成本系数为每缩短1活动日追加1.5人天，不代表每岗只需补0.5人。'
font=ImageFont.truetype('C:/Windows/Fonts/simsun.ttc',18)
def revision(d,ver,msg,prior=None):
 h(d,'修订记录');rows=prior or []
 table(d,['版本','日期','修订人','主要变更'],rows+[[ver,'2026-11-22','邓皓天',msg]])
def pagedtable(d,heads,rows,n=20):
 for k in range(0,len(rows),n):
  if k:d.add_page_break()
  table(d,heads,rows[k:k+n])
def chart(d):
 for k in range(0,len(S),24):
  d.add_page_break();h(d,'组合1更新基准甘特图 '+str(k//24+1));ch=S[k:k+24];im=Image.new('RGB',(1600,95+32*len(ch)),'white');dr=ImageDraw.Draw(im)
  for tick in range(0,math.ceil(finish)+1,10):
   xx=300+tick*9;dr.line((xx,55,xx,im.height),fill='#ddd');dr.text((xx,5),str(tick),font=font,fill='black');dr.text((xx,26),ds(tick)[5:],font=font,fill='black')
  for j,a in enumerate(ch):
   y=70+j*32;dr.text((0,y-8),a['id']+' '+a['role'],font=font,fill='black');dr.rectangle((300+a['ES']*9,y,300+max(a['EF'],a['ES']+.2)*9,y+13),outline='black',fill='black' if a['done'] or a['TF']==0 else 'white')
  pp=T/f'gantt{k}.png';im.save(pp);d.add_picture(str(pp),width=Cm(17.5));p(d,'横轴为累计工作日，上行工作日序号、下行对应日期；非工作日不占刻度。黑条表示已完成或关键活动。精确日历起止见HTML明细。')
# T15
D=new('活动估算与批准进度基准','T15','V2.0');p(D,intro);p(D,cal);p(D,resource)
p(D,'原O M P仍表示工作量估算，期望工期为(O+4M+P)/6，保留两位。两项在制的剩余工程估算分别0.5/1/2和1/2.2/4日。压缩前期望日与批准工期分列，不能以压缩后工期反算原预算人工。已完成28项固定实绩；61项剩余活动重新排程。历史活动的LS/LF/TF/FF不再评估。')
pagedtable(D,['活动 岗位','O M P','原期望日','批准日','批准开始 完成','紧前'],[[a['id']+' '+a['role'],a['tri'],next((c['original'] for c in crash if c['id']==a['id']),5.17 if a['id']=='5.3.1-A' else a['dur']),a['dur'],ds(a['ES'])+'\n'+ds(a['EF'],True),'、'.join(a['pred']) or '无'] for a in S],14)
revision(D,'V2.0','按批准组合1重排，补端午，分列PERT与批准工期；替换同版本未批准计算稿。');D.save(O/'T15_活动清单工期估算与前后关系.docx')
# T16 AOA decomposed complete event network
D=new('网络图第2版与关键路径分析','T16','V2.1');p(D,intro);table(D,['里程碑','日期','完成依据'],[[k,v,'5.3.2-A与CR01-07均完成' if k=='MS4' else '沿用批准基准'] for k,v in ms.items()]);p(D,cal);p(D,'剩余网络起点为52，终点为'+str(finish)+'。ES EF LS LF以开工第0工作日起算；TF=LS−ES，FF=min紧后ES−EF；多条紧后取最小值。网络包含技术和同岗资源顺序。已完成活动只固定实际日期，不纳入剩余反推。QA2承担回归，与原QA功能测试并行，安全测试汇合两链后才可开始。')
p(D,'关键路径 '+ ' → '.join(critical));p(D,'AOA采用双代号：圆圈为事件，实箭线为有工期活动，虚箭线为零工期紧前连接；按活动行分解绘制，相同事件号是同一节点。为避免跨页线交叉，虚线源事件号列于行首，所有虚箭线在事件连接表明列。起点1连接无剩余紧前的活动，所有终止活动连接统一终点999。')
event={a['id']:(2*i+2,2*i+3) for i,a in enumerate(rem)}
for k in range(0,len(rem),11):
 D.add_page_break();ch=rem[k:k+11];im=Image.new('RGB',(1250,60+80*len(ch)),'white');dr=ImageDraw.Draw(im)
 for j,a in enumerate(ch):
  y=50+j*80;ii,jj=event[a['id']];pred=[event[q][1] for q in a['pred'] if q in event] or [1]
  dr.text((0,y-18),'虚线源 '+','.join(map(str,pred)),font=font,fill='black')
  for xx in range(250,380,12):dr.line((xx,y,xx+6,y),fill='black')
  dr.polygon([(390,y),(380,y-5),(380,y+5)],fill='black')
  for xx,ev in [(420,ii),(1160,jj)]:dr.ellipse((xx-24,y-24,xx+24,y+24),outline='black',width=2);dr.text((xx-16,y-10),str(ev),font=font,fill='black')
  dr.line((444,y,1136,y),fill='black',width=3 if a['TF']==0 else 1);dr.polygon([(1136,y),(1125,y-5),(1125,y+5)],fill='black');dr.text((470,y-24),a['id']+'  '+str(a['dur'])+'日',font=font,fill='black')
 path=T/f'aoa{k}.png';im.save(path);D.add_picture(str(path),width=Cm(17.5));table(D,['活动 i→j','ES EF','LS LF','TF FF'],[[a['id']+' '+str(event[a['id']]),f"{a['ES']:.2f} / {a['EF']:.2f}",f"{a['LS']:.2f} / {a['LF']:.2f}",f"{a['TF']:.2f} / {a['FF']:.2f}"] for a in ch])
D.add_page_break();h(D,'完整虚活动事件连接表');rows=[]
for a in rem:
 for q in ([event[v][1] for v in a['pred'] if v in event] or [1]):rows.append([f'{q}→{event[a["id"]][0]}','0',a['id']+'紧前汇合'])
for a in rem:
 if not any(a['id'] in b['pred'] for b in rem):rows.append([f'{event[a["id"]][1]}→999','0','统一终点'])
pagedtable(D,['虚箭线事件','工期','含义'],rows,30);revision(D,'V2.1','按A确认改MS4为5月10日，基于5.3.2-A及CR01-07；活动网络和浮动不变。',[['V2.0','2026-11-22','邓皓天','发布组合1网络图第2版']]);D.save(O/'T16_AOA网络图关键路径与浮动时间.docx')
# T35 only selected baseline
D=new('进度压缩方案与更新基准甘特图','T35','V1.1');p(D,intro);p(D,cal);table(D,['里程碑','批准日期','依据'],[[k,v,'5.3.2-A与CR01-07均完成，不按6.2.2-A' if k=='MS4' else '沿用批准基准'] for k,v in ms.items()]);p(D,resource)
h(D,'审批与落实');table(D,['事项','记录'],[['项目经理决策','项目经理 2026-11-22 选定组合1，范围不变，不开CR-02'],['进度基准批准','项目发起人（模拟）批准进度基准变更，延期8个工作日'],['测试资源','借调QA2，入场时间不晚于'+ds(L['5.3.1-A']['ES'])+'，人员姓名及独立测试环境由项目经理落实'],['赶工资源','后端（二） 技术负责人 前端按活动清单补人，逐项确认入场时间及投入'],['成本授权','本文件记录进度批准；已按MR-02批准，总预算543961.07元；B核对执行口径']])
D.add_page_break();h(D,'人工追加费用');p(D,'35项赶工累计缩短22.57活动·天，追加33.855人天，费用50782.50元；测试返工1.034人天1551.00元，合计34.889人天52333.50元。净项目缩短量与活动缩短量不同，不能按净节省天数收费。')
pagedtable(D,['活动 岗位','原日','批准日','缩短日','追加人天','追加人工元'],[[c['id']+' '+c['role'],c['original'],c['after'],c['cut'],c['added_effort'],money(c['cost'])] for c in crash],18)
chart(D);revision(D,'V1.1','修正MS4口径为5月10日，同步MR-02预算批准及1.3.1投入不追加。',[['V0.1','2026-11-22','邓皓天','首次比较四情景'],['V0.2','2026-11-22','邓皓天','增加组合方案和成本明细'],['V1.0','2026-11-22','邓皓天','定稿组合1，批准延期8工作日']]);D.save(O/'T35_进度压缩与更新基准甘特图.docx')
# T37 cost bridge, WBS, monthly allocation from updated activity windows.
base_mr=420782.81;cr_labor=20375.0;cr_reserve=1018.75;extra=52333.50;newbase=cash(base_mr+cr_labor+cr_reserve+extra);management=cash(newbase*.10);budget=cash(newbase+management);gap=cash(newbase-459112.50);request=cash(budget-459112.50)
# package original cost table is authoritative; costs remain PERT.
pkg={}
for t in Document(R/'docs/T21_成本估计.docx').tables:
 for row in t.rows:
  c=[v.text for v in row.cells]
  if len(c)==6 and c[0].count('.')==2:
   pkg[c[0]]=dict(name=c[1],labor=float(c[3].replace(',','')),other=float(c[4].replace(',','')),cr=0,cr_reserve=0,extra=0)
assert len(pkg)==59
for a in S:
 if a['id'].startswith('CR01'):
  cost=(a['O']+4*a['M']+a['P'])/6*1500;pkg[a['wbs']]['cr']+=cost;pkg[a['wbs']]['cr_reserve']+=cost*.05
for c in crash:pkg[c['wbs']]['extra']+=c['cost']
pkg['5.3.1']['extra']+=1551
bridge=[['MR-01后成本基准',money(base_mr),'417375 + 3407.81'],['CR01新增人工',money(cr_labor),'163/12人天 × 1500'],['CR01增量应急储备',money(cr_reserve),'20375 × 5%'],['赶工与测试追加',money(extra),'50782.50 + 1551.00'],['新成本基准及含风险额度EAC',money(newbase),'已按MR-02批准；B核对执行口径'],['已批准管理储备',money(management),'新成本基准 × 10%'],['已批准总预算',money(budget),'新基准 + 已批准管理储备'],['最低资金缺口',money(gap),'新基准 − 原总预算'],['已批准追加预算',money(request),'已批准总预算 − 原总预算']]
D=new('成本估计第2版','T21','V2.1');p(D,intro);p(D,'状态：已按MR-02批准，2026-11-22生效。以MR01批准后的420782.81元为起点，加入CR01增量21393.75元和选定方案追加人工52333.50元；新成本基准及含风险额度EAC为'+money(newbase)+'元。此处EAC是组合1资源按期落实、剩余工作按更新预算完成的管理预测，包含CR01应急额度；MR-02已批准追加预算84848.57元，总预算543961.07元。')
table(D,['项目','金额元','依据'],bridge)
p(D,'与T29区分：原BAC417375元、原CPI20/23对应历史EAC479981.25元，仍保留作偏差预警，不直接加到本次新基准中。MR01将3407.81元转入基准、减少管理储备，总预算不变；不得再重复追加该笔已批准损失。按新MR基准作纯绩效外推为483900.23元，尚未含CR01和组合1，不等同本稿的选定方案EAC。')
p(D,'工程成本边界：原人工367500元加CR01人工20375元、赶工与返工52333.50元为440208.50元；原直接费用30000元暂不变。原应急19875元已耗尽，MR01转入3407.81元属于损失弥补，CR01仅另留1018.75元。对已确定的赶工费用不机械再加5%应急，以免偏离项目经理指定桥接口径。管理储备按统一10%另列，已按MR-02批准为49451.01元。1.3.1延期监控投入不追加，原25人天及37500元保持不变。')
D.add_page_break();h(D,'按WBS自下而上核对');rows=[]
for w,v in pkg.items():
 orig=(v['labor']+v['other'])*1.05;mr=3407.81 if w=='1.3.1' else 0
 rows.append([w,v['name'],money(orig),money(v['cr']+v['cr_reserve']),money(v['extra']),money(orig+mr+v['cr']+v['cr_reserve']+v['extra'])])
p(D,'原基准按原包成本计5%应急；MR01 3407.81元暂在1.3.1监控控制账单列，不作为该包新增人工，B核账时再追溯到R01/R06实际费用，不改变总额。展示分币尾差由汇总行统一控制。')
pagedtable(D,['WBS','工作包','原基准','CR01含储备','赶工测试','新控制额'],rows,20);table(D,['汇总','原基准','MR01转入','CR01','赶工测试','新基准'],[['合计','417375.00','3407.81','21393.75','52333.50',money(newbase)]])
p(D,'待B核对：EAC采用管理纠偏预测的合理性、实际费用入账与MR01损失归属、历史CPI外推对照、月度分摊。1.3.1延期监控投入不追加，在原25人天及37500元内统筹。6月云环境及技术监控服务按原直接费用额度控制，若合同或用量确有超额，另行补估并报批，不属于1.3.1追加人工。');revision(D,'V2.1','登记MR-02批准，明确总预算及1.3.1延期投入不追加。',[['V2.0','2026-11-22','邓皓天','建立组合1成本桥接和WBS预算']]);D.save(O/'T21_成本估计.docx')
# month: unchanged historical PV through status, move only unearned planned balance.
hist=defaultdict(int)
for v in x['daily_pv']:
 if v['date']<='2027-02-19':hist[v['date'][:7]]+=v['pv_cents']
weights=defaultdict(float)
for a in rem:
 cost=a.get('cost',0)
 if a['id'].startswith('CR01'):continue
 for ix in range(math.floor(a['ES']),math.ceil(a['EF'])):
  overlap=max(0,min(ix+1,a['EF'])-max(ix,a['ES']));weights[str(days[ix])[:7]]+=cost*overlap/a['dur']
def alloc(amount,weights):
 keys=sorted(weights);vals={k:cash(amount*weights[k]/sum(weights.values())) for k in keys};vals[keys[-1]]=cash(vals[keys[-1]]+amount-sum(vals.values()));return vals
remaining=alloc(417375-sum(hist.values())/100,weights);monthly=defaultdict(lambda:defaultdict(float))
for m,v in hist.items():monthly[m]['orig']+=v/100
for m,v in remaining.items():monthly[m]['orig']+=v
monthly['2027-02']['mr']=3407.81
for a in rem:
 parts=[]
 if a['id'].startswith('CR01'):parts.append(('cr',(a['O']+4*a['M']+a['P'])/6*1500*1.05))
 if a['id']=='5.3.1-A':parts.append(('extra',1551))
 for c in crash:
  if c['id']==a['id']:parts.append(('extra',c['cost']))
 for key,amount in parts:
  w=defaultdict(float)
  for ix in range(math.floor(a['ES']),math.ceil(a['EF'])):w[str(days[ix])[:7]]+=max(0,min(ix+1,a['EF'])-max(ix,a['ES']))
  for m,v in alloc(amount,w).items():monthly[m][key]+=v
# aggregate CR exact per activity rounding may leave a cent; absorb it in last CR month.
last=max(m for m in monthly if monthly[m]['cr']);monthly[last]['cr']=cash(monthly[last]['cr']+21393.75-sum(v['cr'] for v in monthly.values()))
monthrows=[];cum=0
for m,v in sorted(monthly.items()):
 total=cash(sum(v.values()));cum=cash(cum+total);monthrows.append([m,*[cash(v[k]) for k in ['orig','mr','cr','extra']],total,cum])
assert abs(cum-newbase)<.011,(cum,newbase)
D=new('成本预算第2版','T22','V2.1');p(D,intro);p(D,'状态：已按MR-02批准，2026-11-22生效。成本基准494510.06元、管理储备49451.01元、总预算543961.07元；批准追加预算84848.57元。B核对执行口径，原基准绩效保留历史对照。');table(D,['项目','金额元','依据'],bridge)
D.add_page_break();h(D,'月度已批准成本基准');p(D,'保留状态日前逐日PV累计194023.44元的历史分摊；原BAC剩余223351.56元按新日历中原剩余活动成本权重重新分摊，包含原其他费用、储备及支持余额，是资金时段分摊约定，不代表这些金额全部是直接人工。MR01转入列2月；CR01和追加人工按对应新活动工作日分摊，末月吸收分币尾差。实际AC178501.56元与PV分别报告，不改写历史实际。')
table(D,['月份','原BAC重排','MR01','CR01含应急','赶工测试','月基准','累计'],[[r[0],*[money(v) for v in r[1:]]] for r in monthrows])
h(D,'追加预算与储备处理');p(D,'MR01后剩余管理储备38329.69元；CR01含应急与组合1合计73727.25元，超过剩余管理储备35397.56元。最低需新增该金额才能覆盖新基准，但管理储备将无余额。按统一10%管理储备规则，已批准总预算543961.07元，追加84848.57元；其中最低缺口35397.56元，另恢复已批准管理储备49451.01元。本次资金转用及追加已按MR-02批准；后续动用管理储备仍须逐笔报发起人批准。1.3.1延期监控投入不追加，维持原25人天及37500元。')
p(D,'新EAC494510.06元，状态日AC178501.56元，ETC316008.50元。该EAC含CR01应急1018.75元；不含新管理储备49451.01元，也不重复叠加T29历史EAC。A已发布MR-02及新版假设表，B核对执行分摊并由A同步T24，T29原基准绩效另留历史对照。');revision(D,'V2.1','登记MR-02预算批准及1.3.1不追加，月度金额保持不变。',[['V2.0','2026-11-22','邓皓天','编制组合1成本预算送审稿']]);D.save(O/'T22_成本预算.docx')
# Shared data, selected scenario only; old alternatives remain in Git history.
metrics={'MR01_BAC':base_mr,'CR01_labor':cr_labor,'CR01_contingency':cr_reserve,'crash_and_test':extra,'new_BAC':newbase,'EAC_including_CR01_contingency':newbase,'AC':178501.56,'ETC':cash(newbase-178501.56),'management_reserve_approved':management,'total_budget_approved':budget,'minimum_funding_gap':gap,'additional_budget_approved':request,'status':'已按MR-02批准'}
payload=dict(version='V1.1',date='2026-11-22',status='组合1进度基准已批准 成本已按MR-02批准',calendar={'holidays':sorted(map(str,hol)),'source':SOURCE_URL,'finish_delay_workdays':8},selected_scenario='组合1 赶工与测试快速跟进',milestones=ms,scenarios=[dict(name='组合1 赶工与测试快速跟进',finish=finish,date=ds(finish,True),cost_delta=extra,activities=S)],critical_path=critical,crash_cost_details=crash,cost_metrics=metrics,monitoring={'wbs':'1.3.1','effort_pd':25,'labor_cost':37500,'delay_added_effort_pd':0,'decision':'延期监控投入不追加'},milestone_basis={'MS4':['5.3.2-A','CR01-07']},monthly_cost_baseline_approved=monthrows)
(O/'D阶段2交接数据.json').write_bytes((json.dumps(payload,ensure_ascii=False,indent=2)+'\n').encode('utf-8'))
def ht(heads,rows):return '<table><tr>'+''.join('<th>'+html.escape(str(v))+'</th>' for v in heads)+'</tr>'+''.join('<tr>'+''.join('<td>'+html.escape(str(v))+'</td>' for v in row)+'</tr>' for row in rows)+'</table>'
page='<meta charset="utf-8"><title>组合1批准基准明细</title><style>body{font:14px Microsoft YaHei;color:black}table{border-collapse:collapse}td,th{border:1px solid #888;padding:5px}th{background:white}</style><h1>组合1批准进度基准 V1.1</h1><p>'+intro+'</p><p>'+cal+'</p><p>'+resource+'</p>'+ht(['里程碑','日期'],ms.items())+'<h2>活动及网络</h2>'+ht(['活动','名称','岗位','批准日','ES','EF','LS','LF','TF','FF','开始','完成','紧前'],[[a['id'],a['name'],a['role'],a['dur'],*[a[k] if not a['done'] or k in ['ES','EF'] else '历史固定 不评估' for k in ['ES','EF','LS','LF','TF','FF']],ds(a['ES']),ds(a['EF'],True),'、'.join(a['pred'])] for a in S])+'<h2>逐活动赶工费用</h2>'+ht(['活动','岗位','原日','批准日','缩短日','追加人天','人工元'],[[c[k] for k in ['id','role','original','after','cut','added_effort','cost']] for c in crash])+'<h2>T37成本桥接</h2>'+ht(['项目','金额','依据'],bridge)+'<h2>月度已批准基准</h2>'+ht(['月份','原BAC重排','MR01','CR01','追加人工','月基准','累计'],monthrows)
(O/'D阶段2_进度计算明细.html').write_bytes((page+'\n').encode('utf-8'))
print(ms);print(metrics);print('critical',' → '.join(critical));print('QA and resource overlap checks passed')
