from pathlib import Path
from datetime import date, timedelta
from fractions import Fraction
from collections import defaultdict
from decimal import Decimal, ROUND_HALF_UP
import json, shutil, ast, math
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from PIL import Image, ImageDraw, ImageFont

ROOT=Path(__file__).resolve().parents[2]; OUT=ROOT/'docs'; TMP=ROOT/'tmp/phase3'; TMP.mkdir(parents=True,exist_ok=True)
DATE='2026-11-15'
src=json.loads((ROOT/'tools/phase2/计算数据.json').read_text(encoding='utf-8'))
tasks=src['tasks']; lookup={a['id']:a for a in tasks}
roles={'PM':'项目经理','BA':'需求分析师','B1':'后端技术负责人','B2':'后端开发二','FE':'前端开发','AI':'AI算法','QA':'测试工程师','EXT':'甲方业务代表'}
holiday={date(2027,1,1),date(2027,4,5)}
holiday.update(date(2027,2,d) for d in range(5,14)); holiday.update(date(2027,5,d) for d in range(1,6))
days=[]; d=date(2026,12,1)
while len(days)<125:
    if d.weekday()<5 and d not in holiday:days.append(d)
    d+=timedelta(days=1)
window=[d for d in days if d<=date(2027,5,31)]
assert len(window)==119 and days[119]==date(2027,6,1)
for a in tasks:
    a['start']=days[a['ES']]; a['finish']=days[a['EF']-1] if a['d'] else days[0]
    a['exact_effort']=Fraction(a['O']+4*a['M']+a['P'],6) if a['role']!='EXT' else Fraction(0)
    a['exact_cost']=a['exact_effort']*1500
assert sum((a['exact_effort'] for a in tasks),Fraction(25))==245
assert sum(a['effort'] for a in tasks)+25==231
assert sum(a['exact_cost'] for a in tasks)+37500==367500
for a in tasks:
    assert a['EF']-a['ES']==a['d']
    assert all(lookup[p]['EF']<=a['ES'] for p in a['pred'])

# Spread support within remaining PM capacity, using the existing 25-person-day budget.
pmload=defaultdict(Fraction)
for a in tasks:
    if a['role']=='PM' and a['d']:
        for day in days[a['ES']:a['EF']]:pmload[day]+=Fraction(a['planned_effort'],a['d'])
available={d:1-pmload[d] for d in window}; total_available=sum(available.values())
assert total_available>=25
monitor={d:Fraction(25)*v/total_available for d,v in available.items()}
assert sum(monitor.values())==25
assert all(pmload[d]+monitor[d]<=1 for d in window)

names={}
for t in Document(OUT/'T13_WBS定稿.docx').tables:
    for row in t.rows:
        c=[x.text.strip() for x in row.cells]
        if c and c[0].count('.')==2 and c[0] in {a['wbs'] for a in tasks}|{'1.3.1'}: names[c[0]]=c[1]
assert len(names)==59

monthly=defaultdict(lambda:defaultdict(Fraction)); package=defaultdict(lambda:defaultdict(Fraction))
for a in tasks:
    occupied=days[a['ES']:a['EF']] if a['d'] else [days[0]]
    for day in occupied:monthly[day.strftime('%Y-%m')]['labor']+=a['exact_cost']/len(occupied)
    package[a['wbs']]['effort']+=a['exact_effort'];package[a['wbs']]['labor']+=a['exact_cost']
for day,eff in monitor.items():monthly[day.strftime('%Y-%m')]['labor']+=eff*1500
package['1.3.1']['effort']=Fraction(25);package['1.3.1']['labor']=Fraction(37500)
months=['2026-12','2027-01','2027-02','2027-03','2027-04','2027-05','2027-06']
other=[('云服务器与数据库','6.1.1',6,2000,[2000]*6,'2个测试实例×750元/月＋数据库500元/月，连续6个月'),('OCR调用','4.1.4',120000,0.05,[0,300,1200,1800,1800,900],'12万页×0.05元/页，覆盖开发、回归及演示'),('大模型API调用','4.4.3',60000000,0.0001,[0,300,1200,1800,1800,900],'6000万token×0.0001元/token，设置调用上限'),('存储备份与监控','6.1.2',6,600,[600]*6,'按600元/月×6个月，保留测试数据与日志'),('软件许可','3.1.1',1,2400,[2400,0,0,0,0,0],'优先开源；专用测试工具课程估算2400元/项目')]
for title,wbs,qty,unit,dist,basis in other:
    assert sum(dist)==round(qty*unit)
    package[wbs]['other']+=sum(dist)
    for m,v in zip(months,dist):monthly[m]['other']+=v
assert sum(p['other'] for p in package.values())==30000
estimate=397500; contingency=Fraction(19875); baseline=417375; management=Fraction(83475,2); budget=Fraction(918225,2)
assert Fraction(estimate)*Fraction(5,100)==contingency and Fraction(baseline)*Fraction(1,10)==management

def money(v):return f'{float(v):,.2f}'
def num(v):return f'{float(v):.4f}'.rstrip('0').rstrip('.')
def rounded(v):
    v=Fraction(v)
    return int((Decimal(v.numerator)*100/Decimal(v.denominator)).quantize(Decimal('1'),rounding=ROUND_HALF_UP))
def cents_distribute(vals,total):
    # Last month absorbs only cent-level rounding; exact fractions remain in source.
    out=[rounded(v) for v in vals]; out[-1]+=rounded(total)-sum(out); return out
labor_c=cents_distribute([monthly[m]['labor'] for m in months],367500)
other_c=[int(monthly[m]['other']*100) for m in months]
est_c=[a+b for a,b in zip(labor_c,other_c)]
reserve_c=cents_distribute([(monthly[m]['labor']+monthly[m]['other'])*Fraction(5,100) for m in months],contingency)
base_c=[a+b for a,b in zip(est_c,reserve_c)];assert sum(base_c)==baseline*100

# Reuse established black-font, unshaded table formatting without executing phase-two builder.
tree=ast.parse((ROOT/'tools/phase2/build.py').read_text(encoding='utf-8'))
for node in tree.body:
    if isinstance(node,ast.FunctionDef) and node.name in ['p','table']:exec(compile(ast.Module(body=[node],type_ignores=[]),'<format>','exec'))
def heading(doc,text):doc.add_paragraph(text,'Heading 1')
def newdoc(title,code):
    doc=Document();sec=doc.sections[0];sec.page_width=Inches(8.27);sec.page_height=Inches(11.69)
    sec.top_margin=sec.bottom_margin=Inches(.65);sec.left_margin=sec.right_margin=Inches(.65)
    for st in doc.styles:
        if st.type==1:
            st.font.name='宋体';st.font.size=Pt(10.5);st.font.color.rgb=RGBColor(0,0,0);st.element.get_or_add_rPr().rFonts.set(qn('w:eastAsia'),'宋体')
        for node in list(st.element.iter()):
            if node.tag in (qn('w:pBdr'),qn('w:shd')):node.getparent().remove(node)
    for sn,sz in [('Title',20),('Heading 1',14),('Heading 2',12)]:
        doc.styles[sn].font.size=Pt(sz);doc.styles[sn].font.bold=True
    doc.styles['Normal'].paragraph_format.space_after=Pt(5);doc.styles['Normal'].paragraph_format.line_spacing=1.1
    p(doc,title,'Title');p(doc,'AI财务票据整理与异常对账智能体');p(doc,f'{code}  V0.1  {DATE}  编制 邓皓天 成员D')
    table(doc,['版本','日期','修订人','主要变更'],[['V0.1',DATE,'邓皓天','完成阶段三首次编制，落实日历、期望计价、月度预算及依赖核对。']],[.6,.95,.7,4.7],9)
    footer=sec.footer.paragraphs[0];footer.alignment=2;footer.add_run(code+' 第 ');f=OxmlElement('w:fldSimple');f.set(qn('w:instr'),'PAGE');footer._p.append(f);footer.add_run(' 页')
    return doc
common='编制依据：项目假设表v1.3、T13 V1.0冻结基线FinAgent-WBS-BL-1.0、T14汇总字典v1.0、T15/T16既有V0.3计算情景及本轮V0.4说明更新、T17 v1.0、T07 v1.2。7个模拟岗位独立排程，完整覆盖38条必须实现与16条建议实现需求；不合并四位真人兼岗容量。'
calendar_note='工作日历采用每周一至周五、每日8小时。元旦1月1—3日、春节2月5—13日、清明4月3—5日、劳动节5月1—5日停工；按项目假设表B3保留周末休息，官方调休上班日1月31日、2月14日、5月8日仍不安排项目工作。窗口共130个周一至周五，扣除落在平日的11个假期日，余119工作日。'
source='日历核实来源：国务院办公厅2026年11月5日通知，人社部中国就业网2026年11月9日转载：https://chinajob.mohrss.gov.cn/h5/c/2026-11-09/635850.shtml，查询日期2026-11-15。采用官方假期但不采用周末补班，系遵守现有项目周末不上班约束，并非声称等同全国调休日历。'
ms=[('MS0','项目启动',date(2026,12,1),date(2026,12,1),'MS0前准备已具备'),('MS1','需求基线确认',date(2026,12,18),days[13],'2.3.2-C'),('MS2','系统设计评审通过',date(2027,1,22),days[27],'3.3.2-A及全部设计包'),('MS3','开发完成及功能冻结',date(2027,3,19),days[76],'全部4.*活动完成'),('MS4','系统测试完成',date(2027,4,23),days[94],'5.3.2-A'),('MS5','用户验收测试通过',date(2027,5,14),days[103],'5.4.2-A'),('MS6','项目验收移交',date(2027,5,31),days[116],'7.3.2-A')]

# Gantt images: every activity has an exact-date register alongside the monthly bars.
fontpath='C:/Windows/Fonts/msyh.ttc';font=ImageFont.truetype(fontpath,24);small=ImageFont.truetype(fontpath,22)
charttasks=sorted(tasks,key=lambda a:(a['ES'],a['EF'],a['id']))
charts=[]
for part in range(4):
    chunk=charttasks[part*21:(part+1)*21];W=1500;H=110+len(chunk)*39
    im=Image.new('RGB',(W,H),'white');dr=ImageDraw.Draw(im);left=630;right=1485;top=110
    start=date(2026,12,1);end=date(2027,6,2);scale=(right-left)/(end-start).days
    for m in [date(2026,12,1),date(2027,1,1),date(2027,2,1),date(2027,3,1),date(2027,4,1),date(2027,5,1),date(2027,6,1)]:
        x=left+(m-start).days*scale;dr.line((x,top-12,x,H-10),fill='#cccccc');dr.text((min(x+3,W-52),55),str(m.month)+'月',font=small,fill='black')
    dr.text((12,16),f'活动基准甘特图 {part+1}/4  2026年12月至2027年6月  实条关键 空心非关键 菱形零工期',font=font,fill='black')
    for i,a in enumerate(chunk):
        y=top+i*39; label=f"{a['id']} {a['name']}"
        while dr.textlength(label,font=small)>600:label=label[:-2]+'…'
        dr.text((12,y-6),label,font=small,fill='black');x=left+(a['start']-start).days*scale;xx=left+(a['finish']-start).days*scale+scale
        if not a['d']:dr.polygon([(x,y),(x+5,y+6),(x,y+12),(x-5,y+6)],fill='black')
        else:dr.rectangle((x,y,xx,y+15),outline='black',fill='black' if a['TF']==0 else 'white',width=2)
        dr.line((0,y+29,W,y+29),fill='#eeeeee')
    dest=TMP/f'gantt-{part+1}.png';im.save(dest);charts.append(dest)

doc=newdoc('基准甘特图与工作日历','T19')
p(doc,'本文件将81项网络活动和1项持续支持工作映射到日历，形成可评审的进度基准初稿。原120工作日情景保持不变，验收移交为2027-05-27，管理收尾为2027-06-01，较项目窗口晚1工作日；正式基准发布须由项目经理处理该差异。')
p(doc,common);heading(doc,'一 工作日历与排程规则');p(doc,calendar_note);p(doc,source)
table(doc,['月份','平日数','扣除假期平日','可用工作日'],[[m, sum(d.weekday()<5 for d in [date(int(m[:4]),int(m[5:]),1)+timedelta(days=i) for i in range(31)] if d.strftime('%Y-%m')==m),sum(d.strftime('%Y-%m')==m and d.weekday()<5 for d in holiday),sum(d.strftime('%Y-%m')==m for d in window)] for m in months[:6]],[1.2,1.5,2,2.2],9)
p(doc,'映射规则：相对区间[ES,EF)的首日为第ES+1个可用工作日，末日为第EF个可用工作日。FS无滞后关系表示紧后从下一可用工作日开始；零工期准备前提记在MS0起点，投入仍计费。工期采用最可能情景，人工计价采用未取整PERT期望；两者分列。')
doc.add_page_break();heading(doc,'二 里程碑及章程日期差异')
table(doc,['编号','里程碑','T07计划','本次排程','日历日偏差'],[[k,n,str(old),str(new),f'{(new-old).days:+d}'] for k,n,old,new,b in ms],[.5,2.05,1.25,1.25,1.3],9)
for k,n,old,new,b in ms:p(doc,f'{k}完成依据：{b}。')
p(doc,'MS6对应对外交付验收，1.4.1-A管理收尾仍在其后，2027-05-28至2027-06-01，占3工作日。不能仅凭MS6早于5月31日就宣称全部工作按期结束。项目经理需选择维持现日历并批准收尾顺延，或批准调整日历/资源/活动方案后由D重排；本稿不擅改项目假设。采用官方3个周末补班是另一情景，需先批准B3例外，不能直接混入本基准。')
heading(doc,'三 甘特图')
p(doc,'横轴为日历月份，条形跨越的周末和假期不计工作量。条形仅用于浏览，精确起止及依赖见活动登记表；已沿用T16资源顺序，关键性以原TF=0为准。')
for i,img in enumerate(charts):
    if i:doc.add_page_break()
    doc.add_picture(str(img),width=Inches(6.9))
heading(doc,'四 活动日期及依赖登记')
for part in range(4):
    chunk=charttasks[part*21:(part+1)*21]
    table(doc,['活动','岗位','工期','开始','完成','TF','紧前活动'],[[a['id'],roles[a['role']],a['d'],str(a['start']),str(a['finish']),a['TF'],'、'.join(a['pred']) or '无'] for a in chunk],[.8,.75,.35,.9,.9,.3,2.35],8)
heading(doc,'五 持续监控与岗位负荷')
p(doc,'1.3.1-A保留25人天和37,500元，覆盖2026-12-01至2027-05-31。按项目经理每日剩余容量比例分摊，离散活动投入按计划人天/工期均匀分配；每日离散投入加监控不超过1人天。25人天是工作量，不等于必须独占25个整天；周内按下表拆分，假期不排工，评审日优先预留容量。6月1日仅管理收尾，不额外增加监控估算。')
weeks=defaultdict(list)
for d in window:weeks[d-timedelta(days=d.weekday())].append(d)
table(doc,['周起始','实际工作日期','监控计划人天','监控人工预算元'],[[str(w),str(ds[0])+'至'+str(ds[-1]),num(sum(monitor[d] for d in ds)),money(sum(monitor[d] for d in ds)*1500)] for w,ds in sorted(weeks.items())],[1.1,2.5,1.5,1.4],9)
p(doc,'本表周合计显示值可能有四位小数舍入，源数据分数合计严格为25人天。其余6个承建岗位沿用T15串行资源关系，无同岗位重叠；7个模拟岗位与T17保持一致。')
heading(doc,'六 评审与交接')
p(doc,'A项目经理：确认收尾晚1工作日的处置，按本表MS0—MS6生成T20并回写T07；确认新预算后回填假设表D1/D3。C：复核范围与活动交接，确认T11 OI-6按单级审核关闭。B：用T21/T22总预算更新财务论证。D：确认后维护正式进度与成本基准，未经批准不删除建议功能。')
doc.save(OUT/'T19_基准甘特图.docx')

doc=newdoc('成本估计','T21')
p(doc,'采用WBS自下而上估算，建设成本估计397,500.00元，其中直接人工367,500.00元、其他直接费用30,000.00元。按统一储备公式测得总预算459,112.50元。费用单价与用量为课程模拟估算，待项目经理评审；不作为供应商报价。')
p(doc,common);heading(doc,'一 估算口径与边界')
p(doc,'工期120工作日，最可能计划投入231人天，PERT期望投入245人天。人工成本按每个活动的(O+4M+P)/6×1,500元计算，先保留分数精度，再汇总；不能把231人天或120天作为计价工作量。含MS0前准备6,?元的实际投入，具体金额见下表；其网络剩余工期为0不代表免费。甲方确认活动不计承建方人工。')
# Replace placeholder with exact source total before save.
prep=sum(a['exact_cost'] for a in tasks if a['d']==0)
for par in doc.paragraphs:
    if '6,?元' in par.text:par.text=par.text.replace('6,?元',money(prep)+'元')
p(doc,'综合人天单价已含工资、社保与公司管理费，不另计间接管理费；项目经理25人天是项目直接工时，不能因此删去。上线后每年15万元运维费用仅用于商业论证，不纳入建设预算。其他费用覆盖开发建设期，不与年度运维重复。')
heading(doc,'二 岗位人工估计')
table(doc,['岗位','计划人天','期望人天','人工成本元'],[[roles[r],sum(a['effort'] for a in tasks if a['role']==r)+(25 if r=='PM' else 0),num(sum((a['exact_effort'] for a in tasks if a['role']==r),Fraction(25 if r=='PM' else 0))),money(sum(a['exact_cost'] for a in tasks if a['role']==r)+(37500 if r=='PM' else 0))] for r in roles if r!='EXT'],[2.3,1.3,1.4,1.6],9)
heading(doc,'三 其他直接费用估计')
table(doc,['费用','归属WBS','用量与单价依据','金额元'],[[n,w,b,money(sum(dist))] for n,w,q,u,dist,b in other],[1.2,.8,3.4,1.2],9)
p(doc,'以上为D基于建设期6个月、虚构测试数据与开发验证量提出的估算假设。服务器、存储跨工作包共用但只在指定归属包入账；调用费用计入接入包用于归集开发及测试调用，不在测试包重复计算。用量与价格未签约，执行前应以报价或计量日志复核。6月1日仅管理文档收尾，不延长云资源使用，故未追加云费；若延长测试环境须走预算变更。')
heading(doc,'四 WBS自下而上汇总')
table(doc,['一级分支','人工元','其他费用元','成本估计元'],[[str(i),money(sum(v['labor'] for k,v in package.items() if k.startswith(str(i)+'.'))),money(sum(v['other'] for k,v in package.items() if k.startswith(str(i)+'.'))),money(sum(v['labor']+v['other'] for k,v in package.items() if k.startswith(str(i)+'.')))] for i in range(1,8)],[1.4,1.7,1.7,1.8],9)
table(doc,['WBS','工作包','期望人天','人工元','其他元','小计元'],[[k,names[k],num(v['effort']),money(v['labor']),money(v['other']),money(v['labor']+v['other'])] for k,v in sorted(package.items(),key=lambda kv:tuple(map(int,kv[0].split('.'))))],[.55,2.4,.7,1,.8,1.15],8)
heading(doc,'五 储备与总预算')
table(doc,['组成','计算公式','金额元'],[['成本估计','人工＋其他直接费用',money(estimate)],['应急储备','成本估计×5%',money(contingency)],['成本基准','成本估计＋应急储备',money(baseline)],['管理储备','成本基准×10%',money(management)],['总预算','成本基准＋管理储备',money(budget)]],[1.4,3.6,1.6],10)
p(doc,'应急储备19,875元对应OCR/模型降级与重复验证、API用量偏差、集成返工等已识别风险，由项目经理在成本基准内动用；风险具体登记号由A在整合风险登记册时关联。管理储备41,737.50元用于未识别风险，不列入成本基准，动用须经项目发起人批准。')
heading(doc,'六 不确定性及确认事项')
p(doc,'其他直接费用用量若整体增加20%，成本估计增加6,000元，总预算按同公式增加6,930元。此敏感性仅用于评审，不改变本稿基值。人力预算已采用PERT期望，5%应急储备为假设表统一要求，未另加三点估计差额。')
p(doc,'A项目经理需评审D提出的30,000元直接费用并回填假设表D1、D3及T07；本轮不擅改唯一取数源。B需以总预算459,112.50元作为第0年投资，不再使用100万元上限情景。C如调整范围必须告知D重算；本轮未裁剪建议功能。')
doc.save(OUT/'T21_成本估计.docx')

doc=newdoc('成本预算与月度成本基准','T22')
p(doc,'将T21建设成本按T19日历分摊，得到成本基准417,375.00元、管理储备41,737.50元及总预算459,112.50元。因原网络管理收尾至2027-06-01，预算另列6月尾项，不隐藏在5月。正式基准须待项目经理确认日期与费用估计后发布。')
p(doc,common);heading(doc,'一 分摊规则')
p(doc,'人工按活动PERT期望成本在其占用的有效工作日均匀分摊；MS0前3项准备投入作为期初建设投入归集在12月1日，不声称当天完成这些前置工作。持续监控按T19逐日容量分配。此表是计划值PV，不是工资付款或供应商结算预测，也不把期望人天用作资源占用量。')
p(doc,'服务器、数据库、存储备份按12月至5月逐月计费；软件许可12月一次计入；OCR和大模型按开发测试用量分配。应急储备按各月成本估计的5%分摊，仅作基准预留，发生风险时经批准转入实际活动；管理储备单独列示，不预摊进月度成本基准。采用分数计算，显示金额保留分，末月吸收分币尾差。')
heading(doc,'二 月度成本基准')
cum=0;rows=[]
for i,m in enumerate(months):
    cum+=base_c[i];rows.append([m,money(Fraction(labor_c[i],100)),money(Fraction(other_c[i],100)),money(Fraction(est_c[i],100)),money(Fraction(reserve_c[i],100)),money(Fraction(base_c[i],100)),money(Fraction(cum,100))])
rows.append(['合计',money(367500),money(30000),money(estimate),money(contingency),money(baseline),money(baseline)])
table(doc,['月份','人工元','其他元','成本估计元','应急储备元','成本基准元','累计基准元'],rows,[.75,.95,.8,1.05,.95,1.05,1.05],8)
heading(doc,'三 其他直接费用月度安排')
table(doc,['费用']+months[:6]+['合计'],[[n]+[money(v) for v in dist]+[money(sum(dist))] for n,w,q,u,dist,b in other],[1.3]+[.75]*6+[.8],8)
heading(doc,'四 累计成本曲线')
im=Image.new('RGB',(1550,720),'white');dr=ImageDraw.Draw(im);xs=[150+i*210 for i in range(7)];base_y=590;top=80
for val in range(0,450001,100000):
    y=base_y-val/450000*(base_y-top);dr.line((140,y,1470,y),fill='#cccccc');dr.text((10,y-12),f'{val/10000:.0f}万元',font=font,fill='black')
cum=0;points=[]
for i,m in enumerate(months):
    cum+=base_c[i];y=base_y-(cum/100)/450000*(base_y-top);points.append((xs[i],y));dr.text((xs[i]-45,620),m,font=font,fill='black');dr.text((xs[i]-45,y-30),f'{cum/100/10000:.2f}',font=small,fill='black')
dr.line(points,fill='black',width=4)
for x,y in points:dr.ellipse((x-5,y-5,x+5,y+5),fill='black')
dr.text((180,20),'月末累计成本基准 单位万元 不含管理储备',font=font,fill='black');im.save(TMP/'cost-curve.png');doc.add_picture(str(TMP/'cost-curve.png'),width=Inches(6.8))
heading(doc,'五 资金边界与监控')
table(doc,['控制口径','金额元','控制方式'],[['成本估计',money(estimate),'建设期人工与其他直接费用'],['应急储备',money(contingency),'包含在月度成本基准，项目经理按风险事项批准使用'],['成本基准',money(baseline),'月末累计PV对照实际成本，分析偏差和完工预测'],['管理储备',money(management),'基准外单列，项目发起人批准后按变更流程转入基准'],['总预算',money(budget),'项目建设资金授权总额，待批准']],[1.25,1.4,4],9)
p(doc,'每周由项目经理收集活动完成量、实际人天、API/云资源使用量；每月由D汇总计划值与实际成本并预测完工费用。超过基准或需要动用管理储备时，由A按T07权限提交审批；不得用管理储备自动冲销偏差。')
heading(doc,'六 预算交接与一致性检查')
p(doc,'本表7个月累计基准等于T21成本基准，管理储备未计入曲线；59个包人工与其他费用逐项汇总等于397,500元。25人天监控只计一次。6月尾项来自1.4.1-A最后1个工作日，若A批准调整日历并使收尾进入5月，D需重算月度分摊，不能仅移动合计数。')
p(doc,'A确认总预算和日期差异后，用本表回写章程与整合计划；B按总预算459,112.50元回算NPV/ROI/回收期，并独立计入年运维15万元；C确认范围不变，维护需求与验收口径。')
doc.save(OUT/'T22_成本预算.docx')

# Correct obsolete freeze statements only, retaining all estimates and network data.
backup=ROOT/'backups'/'阶段3修订前_20261115';backup.mkdir(parents=True,exist_ok=True)
for filename in ['T15_活动清单工期估算与前后关系.docx','T16_AOA网络图关键路径与浮动时间.docx']:
    path=OUT/filename
    if not (backup/filename).exists():shutil.copy2(path,backup/filename)
    doc=Document(backup/filename)
    replacements={
      '完成本轮修订且C确认25条NFR映射后，T13可冻结；目前仅缺C对新增NFR映射的确认，不将条件批准误写为已冻结。':'T13已于2026-11-14补齐NFR追溯并冻结V1.0，基线FinAgent-WBS-BL-1.0；本文件沿用既有活动估算，不重复申请冻结确认。',
      '已落实项目经理四项确认；T13待C确认25条NFR映射后冻结。确认意见按用户转述登记，未代签。':'已落实项目经理四项确认；T13已冻结V1.0，基线FinAgent-WBS-BL-1.0。确认意见按既有记录登记，未代签。',
      'T13仍待C确认新增NFR映射后冻结，不伪造签名或提前标为V1.0。':'T13已补齐追溯并冻结V1.0；本文件仅同步状态，不改变原估算。',
      'T13 未冻结前，本表为同范围测算情景的可复核草稿。':'T13现已冻结V1.0，本表作为同范围估算输入供阶段三使用。',
      '取得C确认后记录确认日期，将T13定稿为V1.0并冻结；正式冻结前保留V0.3。':'已完成：T13于2026-11-14冻结V1.0，见其冻结记录。本表原估算保持不变。',
      'T13冻结前':'T13冻结状态',
      'C确认后':'已冻结',
      '和 C 的 T14 字典 v1.4':'和 T14 汇总字典 v1.0',
      'C字典成本由C按同一确认回写':'T14汇总字典已统一期望成本',
      '按单级审核在下一版T11关闭OI-6；字典保留3/8计划人天，成本按19/6、49/6期望人天同步。':'T11的OI-6仍由C按单级审核关闭；T14汇总字典已保留3/8计划人天，成本按19/6、49/6期望人天统一。',
      '共25次':'合计25人天',
      '每周约1人天，共25次':'每周约1人天，合计25人天',
      'T19需把25次监控分配到具体周/工作日':'T19需把25人天监控分配到具体周/工作日',
    }
    pars=list(doc.paragraphs)+[p for t in doc.tables for row in t.rows for c in row.cells for p in c.paragraphs]
    for par in pars:
        text=par.text
        for old,new in replacements.items():text=text.replace(old,new)
        if 'T15  编制人' in text or 'T16  编制人' in text:text=text.replace('V0.3','V0.4').replace('2026-11-14',DATE)
        if text!=par.text:
            if par.runs:
                par.runs[0].text=text
                for run in par.runs[1:]:run.text=''
            else:par.add_run(text)
    revision=doc.tables[0].add_row().cells
    for c,v in zip(revision,['V0.4',DATE,'邓皓天','同步T13已冻结状态，明确监控按25人天分摊；阶段三日历结果见T19，原活动工期与网络数值不变。']):c.text=v
    heading(doc,'阶段三日历与估算衔接')
    p(doc,'T19按官方假期并保留周末休息映射，窗口119工作日，原120工作日情景管理收尾为2027-06-01，验收移交2027-05-27，差异交项目经理确认。231计划人天用于投入安排，245期望人天用于计价，人工367,500元。监控25人天分配见T19，非强制25次整日工作。分工PDF附录16周和100万元等旧草稿口径作废，统一以项目假设表v1.3为准。')
    doc.save(path)

data={'date':DATE,'version':'V0.1','calendar_source':source,'calendar_rule':calendar_note,'window_workdays':119,'network_duration':120,'project_finish':'2027-06-01','acceptance':'2027-05-27','estimate':estimate,'contingency':float(contingency),'baseline':baseline,'management_reserve':float(management),'budget':float(budget),'planned_effort':231,'expected_effort':245,'labor':367500,'other':30000,'monthly':rows,'milestones':[[k,n,str(o),str(nw),b] for k,n,o,nw,b in ms],'activities':[{k:(str(v) if isinstance(v,(date,Fraction)) else v) for k,v in a.items()} for a in tasks],'monitor_daily':{str(d):str(e) for d,e in monitor.items()},'checks':{'dependencies':True,'pm_daily_capacity':True,'wbs_count':59,'monthly_sum':sum(base_c)/100}}
(ROOT/'tools/phase3/计算数据.json').write_text(json.dumps(data,ensure_ascii=False,indent=2),encoding='utf-8')
print(json.dumps({k:v for k,v in data.items() if k not in ['activities','monitor_daily','calendar_source','calendar_rule']},ensure_ascii=False,indent=2))
