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
hol={date(2027,1,1),date(2027,4,5)}|{date(2027,2,d) for d in range(5,14)}|{date(2027,5,d) for d in range(1,6)}
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
# Alternatives are engineering proposals, not authorizations.
scenarios=[('不压缩',S,finish,0,'完整范围，保留资源顺序')]
for typ in ['赶工','快速跟进','裁剪建议项']:
 z=copy.deepcopy(A);cost=0
 for a in z:
  if a['done']:continue
  if typ=='赶工' and a['role'] in ['B2','FE','B1'] and a['dur']>=2:
   cut=round(a['dur']*.25,2);a['dur']=round(a['dur']-cut,2);cost+=cut*1500*1.5
  if typ=='快速跟进' and a['id'].endswith('-F'):
   # Model early skeleton work as a separate zero-resource lead; retain interface acceptance dependency.
   cut=round(a['dur']*.25,2);a['dur']-=cut;cost+=cut*1500*.2
  if typ=='裁剪建议项' and a['wbs'] in ['4.5.1','4.5.2']:
   cost-=a['cost'];a['dur']=0
 seq,f=schedule(z);scenarios.append((typ,seq,f,round(cost,2),{'赶工':'假设可分割任务补充0.5名同岗短期资源，效率75%；需职能经理落实，实际可用性待确认','快速跟进':'界面25%骨架在等待期预制；预计20%返工，界面最终验收依赖不变；须确认等待期容量','裁剪建议项':'仅模拟取消4.5.1和4.5.2建议项，保留零时长传递关系；须C确认及甲方批准'}[typ]))
# Document helpers

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
intro='依据项目假设表v2.2、已冻结T13 V2.0、已批准CR-01及T28附v1.0编制。状态日2027-02-19，下一个工作日为2月22日。仅完成大作业二阶段2中成员D的T33、T34和T35。'
d=new('活动三点估算与剩余工期','T15');p(d,intro)
p(d,'已完成28项活动固定实绩，2项在制活动仅估剩余，51项原未完成活动及8项变更活动纳入重排。原三点值不再将最可能人天直接当作工期；按F9用期望值保留两位参与计算。单岗满负荷，项目经理活动沿用原日历占用比例；期望工期与成本期望人天分别管理。')
p(d,'两项在制剩余三点为D工程判断，待C和岗位负责人复核：4.1.3-S取0.5/1/2日，考虑模板澄清；4.1.5-S取1/2.2/4日，考虑OCR返工。68%和45%只表示挣值完成度，不直接换算剩余工期。CR01三点沿用批准附件。甲方确认活动按约定1日固定。')
p(d,'资源修正：CR01-02放在4.1.3-S后、4.1.2-S前；CR01-03在原4.4.4-S后；前端与测试增量插入各自原链，新增验收脚本在7.1.1-A后。已完成4.1.4-F不再等待尚未完成4.1.3-F，固定实绩；后续4.1.2-F另补该资源关系。')
for chunk in [S[i:i+18] for i in range(0,len(S),18)]:
 if chunk!=S[:18]:d.add_page_break()
 table(d,['活动与岗位','三点日','期望日或实绩','状态与依据','紧前'],[[a['id']+' '+a['role'],a['tri'],a['dur'],a['pct']+' '+a['basis'],'、'.join(a['pred']) or '无'] for a in chunk])
rev(d,'固定历史实绩，新增剩余估算及8项CR01活动，保留编码并重排资源。');d.save(O/'T15_活动清单工期估算与前后关系.docx')
# AOA: each activity has unique start/end event; dependencies are zero-time dummy arrows.
rem=[a for a in S if not a['done']];event={a['id']:(2*i+2,2*i+3) for i,a in enumerate(rem)}
d=new('剩余工作双代号网络与关键路径','T16');p(d,intro);p(d,'采用AOA双代号：圆圈为事件，实箭线为活动，虚箭线工期0表示紧前关系；事件1为状态日后起点。图按行分解，同编号事件为同一节点，跨行虚箭线用事件号引用，完整连接见紧前事件表。原已完成活动不参与剩余网络计算，历史日期见T15。时间以项目开工第0日计，状态日后起点为52。')
p(d,f'不压缩项目结束为{ds(finish,True)}，网络终点{finish:.2f}工作日。总浮动TF=LS−ES，自由浮动FF=紧后最早开始最小值−EF；末活动以项目终点计算。关键活动见下表TF为0的行。关键链含资源顺序，不能理解为纯技术依赖。')
font=ImageFont.truetype('C:/Windows/Fonts/simsun.ttc',18)
for k in range(0,len(rem),9):
 if k:d.add_page_break()
 chunk=rem[k:k+9];im=Image.new('RGB',(1250,90*len(chunk)+30),'white');dr=ImageDraw.Draw(im)
 for j,a in enumerate(chunk):
  y=j*90+55;i,e=event[a['id']];pr=[event[v][1] for v in a['pred'] if v in event] or [1]
  dr.text((5,y-28),'虚箭线来自 '+','.join(map(str,pr)),font=font,fill='black');[dr.line((xx,y,xx+6,y),fill='black',width=1) for xx in range(285,415,12)];dr.polygon([(420,y),(410,y-5),(410,y+5)],fill='black')
  for xx,nn in [(445,i),(1140,e)]:dr.ellipse((xx-23,y-23,xx+23,y+23),outline='black',width=2);dr.text((xx-14,y-11),str(nn),font=font,fill='black')
  dr.line((468,y,1117,y),fill='black',width=3 if a['TF']==0 else 1);dr.polygon([(1117,y),(1105,y-5),(1105,y+5)],fill='black');dr.text((490,y-27),f"{a['id']}  {a['dur']:.2f}日  TF {a['TF']:.2f}",font=font,fill='black')
 path=T/f'aoa{k}.png';im.save(path);d.add_picture(str(path),width=Cm(17.5))
 table(d,['活动 i→j','ES / EF','LS / LF','TF / FF'],[[f"{a['id']} {event[a['id']][0]}→{event[a['id']][1]}",f"{a['ES']:.2f} / {a['EF']:.2f}",f"{a['LS']:.2f} / {a['LF']:.2f}",f"{a['TF']:.2f} / {a['FF']:.2f}"] for a in chunk])
h(d,'关键活动序列');p(d,' → '.join(a['id'] for a in rem if a['TF']==0));p(d,'图中起点1到每个无剩余紧前的活动起事件设0工期虚活动。终止活动完成事件连接统一终点，终点工期为0。上述规则唯一确定全部虚箭线；关键活动同时存在技术与资源约束。');rev(d,'按状态日重算ES EF LS LF TF FF，建立剩余AOA事件。');d.save(O/'T16_AOA网络图关键路径与浮动时间.docx')
# D dictionary only, preserve original entries and replace schedule fields.
d=new('成员D工作包字典更新','T14 D附件');p(d,intro);p(d,'本附件替换全组字典中成员D原有4.1.4和4.3.2两条；其他成员条目及B负责的新增条目由各自汇总，不能将本附件当作全组完整字典。')
for tb in Document(R/'docs/T14_WBS字典.docx').tables:
 rows=[[c.text for c in r.cells] for r in tb.rows]
 if not any(len(r)>1 and (r[1].startswith('4.1.4 OCR') or r[1].startswith('4.3.2 差异')) for r in rows):continue
 w='4.1.4' if '4.1.4' in rows[1][1] else '4.3.2';aa=[a for a in S if a['wbs']==w]
 for r in rows:
  if r[0]=='工期':r[1]='；'.join(f"{a['id']}：{a['dur']:.2f}日，{ds(a['ES'])}至{ds(a['EF'],True)}，TF={a['TF']:.2f}" for a in aa)+'。已完成按实际，其余按PERT；人天不等同跨度。'
  if r[0]=='备注':r[1]+=(' 状态日全部完成；OCR返工属于4.1.5，CR01无直接增量，保留历史成本。' if w=='4.1.4' else ' 状态日未开始；CR01无直接增量，受后端关键链和前端排队影响。以下日期为不压缩预测，批准T35后同步更新。')
 d.add_page_break() if w=='4.3.2' else None
 h(d,w+' 工作包');table(d,rows[0],rows[1:])
rev(d,'更新D两条字典工期、状态日进度和变更影响。');d.save(O/'T14_WBS字典_成员D.docx')
# Compare and draw all scenario schedules; recommended conditional, no false approval.
d=new('进度压缩方案与更新基准甘特图','T35','V0.1');p(d,intro);p(d,'结论：以下方案均为计算与审批稿。原T19继续有效；项目经理选定方案并由发起人批准后，方可发布V2.0更新基准。赶工资源、快速跟进拆分和范围裁剪均需相应责任人落实。')
table(d,['方案','项目结束','较不压缩节省日','增减人工元','风险与前提'],[[n,ds(f,True),f'{finish-f:.2f}',f'{c:,.2f}',risk] for n,s,f,c,risk in scenarios])
p(d,'赶工费用按缩短日×1.5人天×1500元试算；单活动缩短25%，约需0.5名同岗人员，含交接损耗，未计延期云费用及储备。快速跟进把界面25%骨架从联调移到等待期，表列20%返工增量；须按细分任务落实空闲容量，不能直接削减必须测试。裁剪仅按4.5两包全部删除测算，节约成本沿用原PERT成本。三种费用均供T37整合，不能直接冲减已用应急储备；原EAC 479,981.25元仍为绩效外推。')
best=min(scenarios[1:],key=lambda v:v[2]);p(d,f'建议优先评审{best[0]}情景，其预测结束{ds(best[2],True)}。无方案自动获得批准；若仍晚于5月31日，应协商延期或重新组合可核实措施。不能以透支周末或未经批准删减必须项保日期。')
for n,seq,f,c,risk in scenarios:
 h(d,n+' 里程碑预测');lk={a['id']:a for a in seq};ids=[('MS3','4.4.4-F'),('MS4','6.2.2-A'),('MS5','5.4.2-A'),('MS6','7.3.2-A'),('收尾','1.4.1-A')]
 table(d,['里程碑','预测日期','说明'],[[ms,ds(max(lk[v]['EF'] for v in ['4.4.4-F','CR01-04','CR01-05']) if ms=='MS3' else lk[aid]['EF'],True),'MS3须同时完成CR01-04/05' if ms=='MS3' else '按对应活动完成'] for ms,aid in ids])
 # MS3 gate corrected to include change implementation.
 for start in range(0,len(seq),18):
  d.add_page_break();h(d,n+' 甘特图 '+str(start//18+1));ch=seq[start:start+18];im=Image.new('RGB',(1500,75+len(ch)*32),'white');dr=ImageDraw.Draw(im)
  for tick in range(0,math.ceil(f)+1,10):
   xx=260+tick*7;dr.line((xx,40,xx,im.height),fill='#dddddd');dr.text((xx,5),str(tick),font=font,fill='black')
  for j,a in enumerate(ch):
   y=55+j*32;dr.text((0,y-8),a['id'],font=font,fill='black');dr.rectangle((260+a['ES']*7,y,260+max(a['EF'],a['ES']+.2)*7,y+12),outline='black',fill='black' if a['done'] or a['TF']==0 else 'white')
  pp=T/f'gantt{n}{start}.png';im.save(pp);d.add_picture(str(pp),width=Cm(17.5));p(d,'横轴为从2026-12-01起累计工作日，52为状态日后起点。非工作日不占刻度；黑条为已完成或关键活动，空心条为其他活动。准确日历起止及所有活动详见配套HTML。')
h(d,'审批与交接');table(d,['事项','责任人','当前状态'],[['选择压缩情景及确认资源','项目经理','待选择'],['裁剪范围或两级变更业务核验','C与甲方','CR01已批；裁剪未批'],['更新进度基准','项目发起人','待批准'],['回写T20 T07 T24及T37预算','A与D及B','待方案批准后同步']]);rev(d,'V0.1比较三种压缩方式并给出可审阅情景甘特图；尚未替换原基准。');d.tables[-1].cell(1,0).text='V0.1';d.save(O/'T35_进度压缩与更新基准甘特图.docx')
import html
page='<meta charset="utf-8"><title>阶段2进度计算明细</title><style>body{font:14px Microsoft YaHei;color:black}table{border-collapse:collapse}td,th{border:1px solid #999;padding:5px}th{background:white}</style><h1>阶段2进度计算明细</h1><p>审批稿，原基准仍有效。日历按项目假设表；状态日52之后排程。历史实际不重写。</p>'
for n,seq,f,c,risk in scenarios:
 page+=f'<h2>{n}</h2><p>{risk}</p><table><tr>'+''.join('<th>'+v+'</th>' for v in ['活动','名称','岗位','日','ES','EF','LS','LF','TF','FF','开始','完成','紧前'])+'</tr>'
 for a in seq:page+='<tr>'+''.join('<td>'+html.escape(str(v))+'</td>' for v in [a['id'],a['name'],a['role'],a['dur'],*[round(a[k],2) for k in ['ES','EF','LS','LF','TF','FF']],ds(a['ES']),ds(a['EF'],True),'、'.join(a['pred'])])+'</tr>'
 page+='</table>'
(O/'D阶段2_进度计算明细.html').write_text(page,encoding='utf-8');(O/'D阶段2交接数据.json').write_text(json.dumps({'date':'2026-11-22','status':'方案审批稿','scenarios':[dict(name=n,finish=f,date=ds(f,True),cost_delta=c,activities=s) for n,s,f,c,risk in scenarios]},ensure_ascii=False,indent=2),encoding='utf-8')
print([(n,ds(f,True),c) for n,s,f,c,r in scenarios]);print('done',sum(a['done'] for a in S),'remaining',len(rem))
