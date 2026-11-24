from pathlib import Path
from fractions import Fraction as F
from datetime import date,timedelta
from collections import defaultdict
import math,json,sys
from docx import Document
from docx.shared import Cm,Pt,RGBColor
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
R=Path.cwd();O=R/'docs/阶段二';T=R/'tmp/t39';T.mkdir(parents=True,exist_ok=True)
x=json.loads((O/'D阶段2交接数据.json').read_text(encoding='utf-8'));old=json.loads((O/'D阶段1交接数据.json').read_text(encoding='utf-8'));tasks=x['scenarios'][0]['activities'];hol=set(x['calendar']['holidays']);days=[];dt=date(2026,12,1)
while len(days)<210:
 if dt.weekday()<5 and str(dt) not in hol:days.append(dt)
 dt+=timedelta(days=1)
months=[f'2026-12']+[f'2027-{i:02}' for i in range(1,7)];md={m:sum(d.weekday()<5 and str(d) not in hol for d in [date.fromisoformat(m+'-01')+timedelta(days=j) for j in range(31)] if d.strftime('%Y-%m')==m) for m in months}
def dt(v,end=False):return str(days[math.ceil(float(v))-1 if end else math.floor(float(v))])
roles={'PM':'项目经理','BA':'需求分析','B1':'后端技术负责人','B2':'后端开发二','FE':'前端开发','AI':'AI算法','QA':'原测试','QA2':'借调测试QA2','B1_POOL':'技术负责人同岗补人','B2_POOL':'后端二同岗补人','FE_POOL':'前端同岗补人'}
resource_order=list(roles);monthly=defaultdict(lambda:defaultdict(F));segments=defaultdict(list);records=[];rounding=[];crash={c['id']:c for c in x['crash_cost_details']}
def spread(role,pd,es,ef,aid):
 if pd==0:return
 assert ef>es
 density=pd/(ef-es);segments[role].append((es,ef,density,aid))
 for ix in range(math.floor(es),math.ceil(ef)):
  overlap=max(F(0),min(F(ix+1),ef)-max(F(ix),es));monthly[role][str(days[ix])[:7]]+=density*overlap
base_total=F(25);crtotal=F(0);added=F(0);qa_rework=F(31,30)
for a in tasks:
 base=(F(str(a['O']))+4*F(str(a['M']))+F(str(a['P'])))/6 if a['role']!='EXT' else F(0)
 if a['id'].startswith('CR01'):crtotal+=base
 else:base_total+=base
 extra=F(str(crash[a['id']]['added_effort'])) if a['id'] in crash else F(0);added+=extra
 rework=qa_rework if a['id']=='5.3.1-A' else F(0);es=F(str(a['ES']));ef=F(str(a['EF']));duration=ef-es;role=a['role']
 # Already completed entries retain budget weights; actual hours are not inferred.
 remaining_base=base;historical_partial=F(0)
 if a.get('pct') not in ['0%','100%',None] and not a['done']:
  tri=[F(v) for v in a['tri'].split('/')];remaining_base=(tri[0]+4*tri[1]+tri[2])/6;historical_partial=base-remaining_base
  hs=F(days.index(date.fromisoformat(a['start'])));he=F(days.index(date.fromisoformat(a['finish']))+1);assert he<=52
  spread(role,historical_partial,hs,he,a['id']+'-历史预算权重')
 core=remaining_base+rework;pool=extra;precision_only=False
 if not a['done'] and role in ['B1','B2','FE'] and core>duration and core-duration>F(5,1000):pool+=core-duration;core=duration
 if not a['done'] and core>duration and core-duration<=F(5,1000):precision_only=True;rounding.append(a['id'])
 if role!='EXT':spread(role,core,es,ef,a['id'])
 if pool:
  assert role in ['B1','B2','FE'];spread(role+'_POOL',pool,es,ef,a['id'])
 expected=base+extra+rework
 rec=dict(activity_id=a['id'],wbs=a['wbs'],name=a['name'],completed=a['done'],role=role,position=roles.get(role,'甲方业务代表'),start=dt(es),finish=dt(ef,True),ES=str(es),EF=str(ef),duration_workdays=float(duration),base_persondays=float(base),historical_partial_persondays=float(historical_partial),remaining_base_persondays=float(remaining_base),remaining_expected_persondays=float(remaining_base+extra+rework) if not a['done'] else 0,crash_added_persondays=float(extra),QA2_rework_persondays=float(rework),expected_persondays=float(expected),expected_persondays_exact=str(expected),core_assigned_persondays=float(core),supplement_assigned_persondays=float(pool),mean_total_FTE=float((expected if a['done'] else remaining_base+extra+rework)/duration) if duration else None,required_core_people=0 if role=='EXT' else 1,required_supplement_people=math.ceil(pool/duration) if pool else 0,required_total_people=(1+math.ceil(pool/duration) if pool else 1) if role!='EXT' else None,rounding_only_overload=precision_only,requirement_status='预算期望权重 非历史实际工时' if a['done'] else '按批准排程测算 待项目经理落实资源',source='CR-01增量' if a['id'].startswith('CR01') else '原活动')
 records.append(rec)
assert base_total==245 and crtotal==F(163,12) and added==F('33.855')
# Preserve historical support weights through the status date; spread remaining 25pd over PM spare capacity.
hist_support=F(0)
for date_str,v in old['monitor_daily_effort'].items():
 if date_str<='2027-02-19':monthly['PM'][date_str[:7]]+=F(v);hist_support+=F(v)
end=max(F(str(a['EF'])) for a in tasks);bounds={F(52),end}
for ix in range(52,math.ceil(end)+1):bounds.add(min(F(ix),end))
for es,ef,load,aid in segments['PM']:
 if ef>52:bounds.update([max(es,F(52)),min(ef,end)])
bs=sorted(v for v in bounds if 52<=v<=end);capacity=[]
for u,v in zip(bs,bs[1:]):
 busy=sum((load for es,ef,load,aid in segments['PM'] if es<(u+v)/2<ef),F(0));capacity.append((u,v,max(F(0),1-busy)))
avail=sum(((v-u)*cap for u,v,cap in capacity),F(0));remain_support=F(25)-hist_support;assert avail>=remain_support
for u,v,cap in capacity:spread('PM',(v-u)*cap*remain_support/avail,u,v,'1.3.1-LOE')
assert len([r for r in records if not r['completed']])==61
assert len([r for r in records if not r['completed'] and r['source']=='CR-01增量'])==8
expected_total=F(245)+crtotal+added+qa_rework
assert sum((sum(v.values()) for v in monthly.values()),F(0))==expected_total
# Detect peaks in remaining period, not the non-executable historical milestone cost weights.
peaks={}
for role,segs in segments.items():
 edges=sorted({max(F(52),es) for es,ef,load,aid in segs if ef>52}|{ef for es,ef,load,aid in segs if ef>52});peak=F(0)
 for u,v in zip(edges,edges[1:]):peak=max(peak,sum((load for es,ef,load,aid in segs if es<(u+v)/2<ef),F(0)))
 peaks[role]=float(peak)
overloads=[]
ai_edges=sorted({max(F(52),es) for es,ef,load,aid in segments['AI'] if ef>52}|{ef for es,ef,load,aid in segments['AI'] if ef>52})
for u,v in zip(ai_edges,ai_edges[1:]):
 active=[(load,aid) for es,ef,load,aid in segments['AI'] if es<(u+v)/2<ef]
 load=sum((z[0] for z in active),F(0))
 if load>F('1.005'):overloads.append(dict(role='AI',start=dt(u),finish=dt(v,True),ES=str(u),EF=str(v),FTE=float(load),activities=[z[1] for z in active],required_people=math.ceil(load),status='待项目经理落实并行协作或批准调整排程'))
assert all(peaks.get(role,0)<=1.005 for role in ['PM','BA','B1','B2','FE','QA','QA2','AI']),peaks
monthly_pd={r:{m:float(monthly[r][m]) for m in months} for r in resource_order};monthly_pm={r:{m:float(monthly[r][m]/md[m]) for m in months} for r in resource_order}
finance=dict(original_labor=367500,CR01_labor=20375,crash_labor=50782.5,QA2_rework_approved=1551,approved_labor_total=440208.5,personday_times_1500=float(expected_total*1500),approved_precision_difference=1,effort_total_persondays=float(expected_total),effort_total_exact=str(expected_total),scope='只核对人工；不含其他费用 应急储备 MR转入及管理储备')
assert expected_total*1500+1==F('440208.50')
payload=dict(version='V0.1',date='2026-11-24',source_commit='9897ad3',status='D资源估算交接稿 人员配置待A落实 非实际工时',unit_personmonth='人月=预算期望人天/该月完整项目工作日历可用工作日；为平均等效人数 非实际人头',months=months,calendar_workdays=md,activities=records,remaining_activities=[r for r in records if not r['completed']],support=dict(wbs='1.3.1',expected_persondays=25,already_allocated_persondays=float(hist_support),remaining_persondays=float(remain_support),extension_added_persondays=0),monthly_persondays=monthly_pd,monthly_personmonths=monthly_pm,remaining_period_peak_FTE=peaks,unresolved_capacity_overlaps=overloads,finance_reconciliation=finance,rounding_only_cases=rounding,QA2_execution_sensitivity='第三批分批预测96.24来自D容量附件，尚待测试岗确认；本表按批准89.57至95.77计量。分批调整在同一4月内，不改变月度人天。')
(O/'T39_D活动资源估算.json').write_bytes((json.dumps(payload,ensure_ascii=False,indent=2)+'\n').encode('utf-8'))
# Create a concise resource handoff, following template Date/Project/WBS/name/description slots.
d=Document();sec=d.sections[0];sec.page_width=Cm(29.7);sec.page_height=Cm(21);sec.top_margin=sec.bottom_margin=Cm(1.3);sec.left_margin=sec.right_margin=Cm(1.3)
for sn in ['Normal','Title','Heading 1']:
 st=d.styles[sn];st.font.name='宋体';st._element.get_or_add_rPr().rFonts.set(qn('w:eastAsia'),'宋体');st.font.color.rgb=RGBColor(0,0,0);st.font.size=Pt(11 if sn=='Normal' else 18 if sn=='Title' else 12);st.paragraph_format.space_after=Pt(4)
for e in list(d.styles.element.xpath('.//w:pBdr')):e.getparent().remove(e)
def p(s):d.add_paragraph(s)
def h(s):d.add_paragraph(s,'Heading 1')
def tab(heads,rows):
 t=d.add_table(rows=1,cols=len(heads));t.style='Table Grid'
 for i,vals in enumerate([heads]+rows):
  rr=t.rows[0] if i==0 else t.add_row();rr._tr.get_or_add_trPr().append(OxmlElement('w:cantSplit'))
  for c,v in zip(rr.cells,vals):
   c.text=str(v)
   for pp in c.paragraphs:
    pp.paragraph_format.space_after=Pt(2);pp.paragraph_format.line_spacing=1.02
    for run in pp.runs:run.font.size=Pt(10);run.bold=i==0
 t.rows[0]._tr.get_or_add_trPr().append(OxmlElement('w:tblHeader'));return t
def n(v):return f'{float(v):.4f}'
d.add_paragraph('活动资源估算与月度投入交接','Title');p('项目 AI财务票据整理与异常对账智能体  T39 D输入  V0.1  2026-11-24  编制 邓皓天 成员D  接收 项目经理')
p('供项目经理编制T39资源管理计划及资源直方图。依activity resource requirements.doc列出WBS编码 名称与资源描述，并按resource histogram.xls提供岗位×月份投入。剩余61项已含8项CR-01，不重复增加8项；全周期另提供28项已完成预算权重及1.3.1支持工作，便于月度汇总。日期沿用批准T15/T35，不改基准。')
tab(['版本','日期','编制人','主要变更'],[['V0.1','2026-11-24','邓皓天','按教师模板提供61项剩余活动资源、全周期人天/人月矩阵、补人容量及人工核对，交A汇总T39。']])
h('一 估算与换算口径')
p('在制活动4.1.3-S与4.1.5-S按C复核的剩余三点估算，分别为13/12和2.30人天；原全量期望人天与剩余量之差仅作为历史预算权重，分摊在原日期。它不代表实际工时；不得把全量工作量重复压到剩余工期。后续表仅列剩余需求，全周期矩阵另保留历史部分。')
p('人天表示有效投入，工期表示活动跨度。综合单价1500元/人天；全周期原期望245人天（含1.3.1的25人天），CR-01新增13.5833人天，赶工新增33.8550人天，QA2返工按精确PERT新增1.0333人天，总计'+n(expected_total)+'人天。计价核对见末页；不把其他费用与储备折成人天。')
p('月度人月=当月期望投入人天÷该月完整工作日历工作日数，是月平均等效全职人数，可用于直方图；不是名册人数，不使用未约定的20人天/人月常数。6月分母用全月可用日数，项目只排至6月11日。分摊按活动与每日区间交集，保留精确分数计算，展示4位，合计以精确值为准。')
tab(['月份']+months,[['全月工作日']+[md[m] for m in months],['项目窗口内工作日']+[sum(dt.strftime('%Y-%m')==m for dt in days[:math.ceil(end)]) for m in months]])
p('每项主岗1人；期望基础工作量超过压缩后1人容量时，超出部分与新增赶工人天一起计入同岗补人池，不把超负荷藏在主岗中。人数为并行定员建议（向上取整），平均FTE另列。少于0.005人天的超量仅源自期望工期保留2位的舍入，标记精度项，不据此增加一整人。QA2为独立借调测试，不挤占原QA的CR01-06等活动。')
p('1.3.1仍25人天、不追加。状态日前按原逐日支持权重保留，状态日后在项目经理离散活动的剩余容量内重新分摊。历史完成活动只表示预算期望权重，不是实际工时；MS0前准备工作在源数据中为启动日费用锚点，不能据此认定当天有足够执行容量。')
d.add_page_break();h('二 岗位与月份投入 人月')
tab(['岗位 / 资源池']+months+['人天合计'],[[roles[r]]+[n(monthly[r][m]/md[m]) for m in months]+[n(sum(monthly[r].values()))] for r in resource_order]+[['月合计']+[n(sum(monthly[r][m] for r in resource_order)/md[m]) for m in months]+[n(expected_total)]])
h('三 同口径人天与人工核对')
tab(['项目','人天或金额','说明'],[['原基准期望投入','245.0000人天 / 367500.00元','包含监控25人天；QA2承接原回归31/6人天不重复计费'],['CR-01','13.5833人天 / 20375.00元','8项增量归所属岗位'],['赶工','33.8550人天 / 50782.50元','补人池还承接主岗原工作量因压缩产生的容量外溢；该部分不是新增成本'],['QA2返工','1.0333人天 / 批准1551.00元','精确PERT×20%对应1550元；已批金额1元精度差单列'],['合计',n(expected_total)+'人天 / 批准人工440208.50元','人天×1500=440207.50元，加已批精度差1元；不把1元算成新增能力']])
p('JSON同时给出岗位×月份人天和人月、逐活动基础/赶工/返工分解、主岗与补人池分摊。岗位“所需人数”为执行定员，月度人月为平均投入，两者不得混作同一指标；同岗补人池每人可在不重叠活动间复用。')
remaining=[r for r in records if not r['completed']]
for k in range(0,len(remaining),9):
 d.add_page_break();h('四 剩余活动资源需求 '+str(k//9+1));p('全部为批准更新基准日期；主岗与补人须同期开工。工期舍入精度项见星号；甲方确认活动不计承建方人工。')
 tab(['活动 / WBS与名称','岗位与所需人数','剩余基础人天','赶工 / 返工新增','剩余合计人天','开始 / 完成'],[[r['activity_id']+'\n'+r['wbs']+' '+r['name'],r['position']+' 主岗'+str(r['required_core_people'])+'人'+(' + 同岗补人'+str(r['required_supplement_people'])+'人' if r['required_supplement_people'] else '')+(' *舍入' if r['rounding_only_overload'] else ''),n(r['remaining_base_persondays']),n(r['crash_added_persondays'])+' / '+n(r['QA2_rework_persondays']),n(r['remaining_expected_persondays']),r['start']+'\n'+r['finish']] for r in remaining[k:k+9]])
d.add_page_break();h('五 资源落实与约束交接')
tab(['资源','剩余区间峰值FTE','落实事项'],[[roles[r],n(peaks.get(r,0)),'沿用原岗，核对实际可用性' if not r.endswith('_POOL') and r!='QA2' else '项目经理落实人员姓名、到岗日、权限、交接及同岗协作；不能仅以月平均人数证明每日够用'] for r in resource_order])
assert not overloads
p('本表是预算期望资源建议，不是实际出勤或新增人员已到岗证明。所有主岗剩余区间峰值均不高于1.005FTE，微差为工期舍入。三类补人池另列容量。项目经理仍需落实QA2、赶工补人来源与到岗日，核对原岗实际可用性；补人池人数向上取整，按不重叠任务复用，不额外增加已批人工。若需新增工作量或改变日期，另走变更评估。')
p('QA2按批准汇总活动[89.57,95.77)分摊。D第三批容量附件另提出分批预测96.24，尚待测试岗确认；两者都在4月内，月度投入不变。若批准拆分回归子活动，再同步活动表及网络，不在资源估算中私改日期。用例及证据复用仍按C确认的分批门槛。')
p('T22 WBS×月矩阵是资金时段分配；本表是岗位工作量时段分配，二者不要求各月按1500元逐格相等，须核对全周期人工总额及对应WBS增量。不得用成本基准494510.06元或总预算543961.07元直接折算人员数。')
p('待A确认：月度人月采用实际工作日历的换算、逐活动所需人数及补人来源、历史预算权重与真实出勤的区别。补入T39后可直接用JSON的monthly_personmonths绘制资源直方图，monthly_persondays保留核对。新建议如改变已批工期或预算，应另走变更评估。')

# Embed compatibility normalization to keep the generator portable without unrelated scripts.
z=d.settings.element.find(qn('w:zoom'))
if z is not None:z.set(qn('w:percent'),'100')
d.save(O/'T39_D活动资源估算.docx')
print('remaining',len(remaining),'CR',8,'total PD',float(expected_total),'finance diff',1,'calendar',md,'peaks',peaks)
