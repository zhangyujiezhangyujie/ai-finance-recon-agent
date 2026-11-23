from pathlib import Path
from datetime import date,timedelta
from collections import defaultdict,deque
from fractions import Fraction as F
from decimal import Decimal,ROUND_HALF_UP
import json,math,copy,hashlib,subprocess,io
from docx import Document
from docx.shared import Cm,Pt
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
R=Path.cwd();O=R/'docs/阶段二';old=json.loads((O/'D阶段1交接数据.json').read_text(encoding='utf-8'));new=json.loads((O/'D阶段2交接数据.json').read_text(encoding='utf-8'));original_budget=R/'docs/T22_成本预算.docx'
if not original_budget.exists():original_budget=R/'T22_成本预算.docx'
original_doc=Document(original_budget)
src={'monthly':[[c.text for c in row.cells] for row in next(t for t in original_doc.tables if t.cell(0,0).text=='月份').rows[1:7]]}
months=['2026-12','2027-01','2027-02','2027-03','2027-04','2027-05','2027-06'];labels=['1 项目管理','2 需求分析','3 系统设计','4 开发实现','5 测试','6 部署与上线','7 培训与交付','应急储备']
def cent(v):return int((Decimal(str(v))*100).quantize(Decimal('1'),rounding=ROUND_HALF_UP))
# Exact original schedule cost weights.
days=[date.fromisoformat(v['date']) for v in old['daily_pv']];mi={m:j for j,m in enumerate(months)}
seed=[[0.0]*6 for _ in range(8)]
for a in old['baseline_activities']:
 ds=days[a['ES']:a['EF']] if a['d'] else [days[0]]
 for dt in ds:seed[int(a['wbs'][0])-1][mi[dt.strftime('%Y-%m')]]+=a['cost']/len(ds)
for ds,eff in old['monitor_daily_effort'].items():seed[0][mi[ds[:7]]]+=float(F(eff))*1500
other=[(6,[2000]*6),(4,[0,300,1200,1800,1800,900]),(4,[0,300,1200,1800,1800,900]),(6,[600]*6),(3,[2400,0,0,0,0,0])]
for w,dist in other:
 for j,v in enumerate(dist):seed[w-1][j]+=v
for j,r in enumerate(src['monthly'][:6]):seed[7][j]=float(r[4].replace(',',''))
rowtot=[59250,25000,35900,175000,44500,38100,19750,19875];coltot=[cent(r[5].replace(',','')) for r in src['monthly'][:6]]
# Proportional balancing preserves approved row and column controls; zero weights stay zero.
def balance(seed,rows,cols):
 assert sum(rows)==sum(cols);a=copy.deepcopy(seed)
 for i in range(len(rows)):
  if rows[i]==0:a[i]=[0.0]*len(cols)
 for iteration in range(10000):
  for i,r in enumerate(rows):
   ss=sum(a[i]);
   if r and not ss:raise ValueError('row has no timing support')
   a[i]=[v*r/ss if ss else 0 for v in a[i]]
  for j,c in enumerate(cols):
   ss=sum(a[i][j] for i in range(len(rows)))
   if c and not ss:raise ValueError('month has no timing support')
   for i in range(len(rows)):a[i][j]=a[i][j]*c/ss if ss else 0
  if max(abs(sum(a[i])-rows[i]) for i in range(len(rows)))<1e-6:break
 else:raise ValueError('balancing not converged')
 out=[[math.floor(max(0,v)+1e-9) for v in r] for r in a];rd=[rows[i]-sum(out[i]) for i in range(len(rows))];cd=[cols[j]-sum(out[i][j] for i in range(len(rows))) for j in range(len(cols))]
 # Unit-capacity bipartite flow rounds cells by at most one cent, preserving both margins.
 n=len(rows);m=len(cols);source=n+m;sink=source+1;cap=defaultdict(int);adj=defaultdict(list)
 def edge(u,v,c):adj[u].append(v);adj[v].append(u);cap[u,v]=c
 for i,v in enumerate(rd):edge(source,i,v)
 for j,v in enumerate(cd):edge(n+j,sink,v)
 for i in range(n):
  for j in sorted(range(m),key=lambda j:-(a[i][j]-out[i][j])):
   if a[i][j]>0:edge(i,n+j,1)
 flow=0
 while True:
  prev={source:None};q=deque([source])
  while q and sink not in prev:
   u=q.popleft()
   for v in adj[u]:
    if v not in prev and cap[u,v]>0:prev[v]=u;q.append(v)
  if sink not in prev:break
  v=sink
  while prev[v] is not None:u=prev[v];cap[u,v]-=1;cap[v,u]+=1;v=u
  flow+=1
 assert flow==sum(rd),(flow,rd,cd)
 for i in range(n):
  for j in range(m):out[i][j]+=cap[n+j,i]
 assert [sum(r) for r in out]==rows
 assert [sum(out[i][j] for i in range(n)) for j in range(m)]==cols
 return out
orig=balance(seed,[cent(v) for v in rowtot],coltot)
# Keep original WBS allocations in Dec/Jan; allocate the unspent original BAC over revised activity windows.
ndays=[];dt=date(2026,12,1);hol=set(new['calendar']['holidays'])
while len(ndays)<210:
 if dt.weekday()<5 and str(dt) not in hol:ndays.append(dt)
 dt+=timedelta(days=1)
weights=[[0.0]*5 for _ in range(8)]
for a in new['scenarios'][0]['activities']:
 if a['done'] or a['id'].startswith('CR01'):continue
 for ix in range(math.floor(a['ES']),math.ceil(a['EF'])):
  dt=ndays[ix];j=mi[dt.strftime('%Y-%m')]-2
  if 0<=j<5:weights[int(a['wbs'][0])-1][j]+=a.get('cost',0)*max(0,min(ix+1,a['EF'])-max(ix,a['ES']))/a['dur']
# Support and remaining nonlabor may be scheduled throughout the remaining approved window.
for j in range(5):weights[0][j]+=1;weights[7][j]+=new['monthly_cost_baseline_approved'][j+2][1]*.05
for a in old['baseline_activities']:
 if a['ES']<52:
  for ix in range(a['ES'],min(a['EF'],52)):
   if days[ix].strftime('%Y-%m')=='2027-02':weights[int(a['wbs'][0])-1][0]+=a['cost']/max(a['d'],1)
for j in range(5):
 weights[3][j]+=[2400,3600,3600,1800,0][j] # Original remaining OCR/API monthly fee weights; no added cost.
 weights[5][j]+=2600 # Remaining cloud/storage fees use equal monthly weights; total remains controlled.
rt=[cent(rowtot[i])-sum(orig[i][:2]) for i in range(8)];ct=[cent(r[1]) for r in new['monthly_cost_baseline_approved'][2:]]
rest=balance(weights,rt,ct);revised=[orig[i][:2]+rest[i] for i in range(8)]
# Allocate approved MR01, CR01 and crash/test costs to their original WBS, leaving only CR contingency in reserve row.
def allocate_cents(amount,w):
 keys=sorted(w);v={k:round(amount*w[k]/sum(w.values())) for k in keys};v[keys[-1]]+=amount-sum(v.values());return v
for a in new['scenarios'][0]['activities']:
 parts=[]
 if a['id'].startswith('CR01'):
  cost=cent(float((F(str(a['O']))+4*F(str(a['M']))+F(str(a['P'])))/6*1500));parts.extend([(int(a['wbs'][0])-1,cost),(7,round(cost*.05))])
 if a['id']=='5.3.1-A':parts.append((4,155100))
 for c in new['crash_cost_details']:
  if c['id']==a['id']:parts.append((int(a['wbs'][0])-1,cent(c['cost'])))
 for row,amt in parts:
  w=defaultdict(float)
  for ix in range(math.floor(a['ES']),math.ceil(a['EF'])):w[mi[ndays[ix].strftime('%Y-%m')]]+=max(0,min(ix+1,a['EF'])-max(ix,a['ES']))
  for j,v in allocate_cents(amt,w).items():revised[row][j]+=v
revised[0][2]+=340781
# Existing approved source table distributes CR and crash costs with activity-cent tails.
# Adjust within each cost category's eligible WBS to match those already approved monthly totals.
targetcols=[cent(r[5]) for r in new['monthly_cost_baseline_approved']];targetrows=[sum(r) for r in revised]
# Only cent-level column differences should arise from independent component rounding.
delta=[targetcols[j]-sum(revised[i][j] for i in range(8)) for j in range(7)]
assert max(map(abs,delta))<10,delta
for jp in range(7):
 while delta[jp]>0:
  jm=next(j for j,v in enumerate(delta) if v<0)
  row=next(i for i in [7,0,4,3,5,6] if revised[i][jm]>0 and revised[i][jp]>0)
  revised[row][jm]-=1;revised[row][jp]+=1;delta[jm]+=1;delta[jp]-=1
assert [sum(r) for r in revised]==targetrows
assert [sum(revised[i][j] for i in range(8)) for j in range(7)]==targetcols
assert sum(targetrows)==49451006 and sum(map(sum,orig))==41737500
# Scope assurance for snapshots.
backup=R/'backups/模板对照第一阶段_20261123';backup.mkdir(parents=True,exist_ok=True)
files=['T22_成本预算.docx']
for name in files:
 b=backup/name
 if not b.exists():
  baseline=subprocess.check_output(['git','-c','safe.directory='+str(R),'show','03723286f9a4f76b296c59586911a343f78c48b2:docs/阶段二/T22_成本预算.docx'])
  b.write_bytes(baseline)
# Cost template structure: WBS categories x months, row totals, reserve row, column totals.
d=Document(backup/'T22_成本预算.docx')
for p in d.paragraphs:
 if p.text.startswith('T22  V2.1'):p.text=p.text.replace('V2.1','V2.2').replace('2026-11-22','2026-11-23')
# Put appendices ahead of the existing revision record.
anchor=next(p._p for p in d.paragraphs if p.text=='修订记录')
def paragraph(s,style=None):
 p=d.add_paragraph(s,style);anchor.addprevious(p._p);return p

def matrix(title,mat,ms):
 p=paragraph(title,'Heading 1');p.paragraph_format.page_break_before=True
 paragraph('单位：元。行按WBS一级分支列示，另列应急储备。'+('CR-01及赶工、测试追加计入对应分支，MR-01转入计入1项目管理。' if len(ms)==7 else '本历史表不含MR-01、CR-01和组合1追加费用。')+'管理储备不进入成本基准。')
 t=d.add_table(rows=1,cols=len(ms)+2);t.style='Table Grid';t.autofit=False
 rows=[['WBS一级分支',*ms,'合计']]+[[labels[i],*[f'{v/100:,.2f}' for v in mat[i]],f'{sum(mat[i])/100:,.2f}'] for i in range(8)]+[['月合计',*[f'{sum(mat[i][j] for i in range(8))/100:,.2f}' for j in range(len(ms))],f'{sum(map(sum,mat))/100:,.2f}']]
 for ir,row in enumerate(rows):
  cells=t.rows[0].cells if ir==0 else t.add_row().cells
  for j,(c,v) in enumerate(zip(cells,row)):
   c.text=v;c.width=Cm(3 if j==0 else (17.8-3)/(len(ms)+1))
   for pp in c.paragraphs:
    pp.paragraph_format.space_after=Pt(3);pp.paragraph_format.line_spacing=1.05
    for rr in pp.runs:rr.font.size=Pt(8);rr.bold=ir in [0,len(rows)-1]
  t.rows[ir]._tr.get_or_add_trPr().append(OxmlElement('w:cantSplit'))
 t.rows[0]._tr.get_or_add_trPr().append(OxmlElement('w:tblHeader'));anchor.addprevious(t._tbl)
matrix('WBS一级与月份成本基准',revised,months)
paragraph('本表不改变已批准资金来源表的月基准、EAC494510.06元和总预算543961.07元。新增的是工作类别维度：保留2026年12月及2027年1月原分摊；原BAC剩余部分以更新活动成本时间权重作比例平衡，使原WBS一级总成本与既有各月原BAC重排控制数同时守恒；原其他费用及监控余额一并纳入分摊权重。CR-01及赶工、测试追加按原活动日期分摊，末位分币尾差受行列合计双重校正。此为资金分摊，不是逐月新增工作量或实际费用。')
paragraph('核对：新基准月合计分别为 '+ '、'.join(f'{v/100:,.2f}' for v in targetcols)+' 元；合计494510.06元。1.3.1仍25人天及37500元，MR-01转入的3407.81元属于风险损失控制账，不是监控追加人工。')
matrix('原基准WBS一级与月份历史对照',orig,months[:6])
paragraph('原基准依T19活动日期、原PERT人工、1.3.1逐日监控投入和T22第一版其他费用计划分摊；应急储备单列。按原分支成本与原T22每月金额作分币平衡，总计417375.00元。供B作T29原计划曲线，不能用新基准替代历史BAC。')
paragraph('交接B：新旧矩阵可直接读取配套T22_WBS月度成本基准.json，金额以整数分保存，并带行合计、列合计及累计值。新表覆盖2026-12至2027-06；原表覆盖2026-12至2027-05。状态日原PV194023.44元、EV155218.75元、AC178501.56元保持不变，不因新增矩阵重算或改写。')
t=d.tables[-2] if False else next(t for t in d.tables if t.cell(0,0).text=='版本')
cs=t.add_row().cells
for c,v in zip(cs,['V2.2','2026-11-23','邓皓天','对照cost baseline模板新增新旧两张WBS一级×月份矩阵和JSON交接，行列核对一致，原批准金额不变。']):c.text=v
for c in cs:
 for pp in c.paragraphs:
  for rr in pp.runs:rr.font.size=Pt(8)
d.save(O/'T22_成本预算.docx')
# Final data artifact, not a spreadsheet installation or recalculation dependency.
def pack(mat,ms):return dict(months=ms,row_labels=labels,values_cents=mat,row_totals_cents=[sum(r) for r in mat],monthly_totals_cents=[sum(mat[i][j] for i in range(8)) for j in range(len(ms))],total_cents=sum(map(sum,mat)))
payload=dict(version='V2.2',date='2026-11-23',unit='分',source='T22 V2.2; approved monthly totals unchanged',allocation='成本时段权重比例平衡及分币双边守恒；非新成本估算',approved=pack(revised,months),original=pack(orig,months[:6]),checks={'approved_total_cents':49451006,'original_total_cents':41737500,'row_column_controls_passed':True})
(O/'T22_WBS月度成本基准.json').write_bytes((json.dumps(payload,ensure_ascii=False,indent=2)+'\n').encode('utf-8'))
print('original rows',list(zip(labels,[sum(r)/100 for r in orig])));print('revised rows',list(zip(labels,[sum(r)/100 for r in revised])));print('monthly',targetcols);print('all matrix controls passed')
