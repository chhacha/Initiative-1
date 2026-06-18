#!/usr/bin/env python3
"""Add native Excel charts (Stabilized Portfolio) to the Graphs tab via pure
zip/XML surgery so NOTHING else in the workbook is altered."""
import zipfile
from openpyxl.utils import get_column_letter as col

SRC = 'model_original.xlsx'
OUT = 'UAE_Industrial_Strategy_Financial_Model_0626_charts.xlsx'
USD  = "'UAE Platform Cash Flow (USD)'"
DEBT = "' Debt Schedule'"          # leading space is part of the sheet name
GR   = "Graphs"

YEARS = list(range(2026, 2037))    # 11 yrs
NY = len(YEARS)

NOI      = [29587613,30793546,32507856,33046715,33896556,35395796,36443511,37522239,38632897,39776431,24868384]
NETEXIT  = [0,0,0,0,0,0,0,0,0,195200459,355613421]
CONTRIB  = [-190744748,0,0,0,0,0,0,0,0,0,0]
DISTRIB  = [0,9394268,11242780,12021633,13083478,14736276,15978188,17248320,18547510,209674973,185348535]
NETCF    = [-198492477,7486820,9335333,10114185,11176030,12828829,14070740,15340873,16640062,205815521,179884954]
DRAW     = [214536785,0,0,0,0,0,0,0,0,0,0]
INTEREST = [-2424041,-14375960,-14087469,-13798979,-13510489,-13221999,-12933508,-12645018,-12356528,-12068037,-11779547]
AMORT    = [-708650,-4251899,-4251899,-4251899,-4251899,-4251899,-4251899,-4251899,-4251899,-4251899,-4251899]
BULLET   = [0,0,0,0,0,0,0,0,0,0,-171309142]
CUM=[];s=0
for v in NETCF: s+=v; CUM.append(s)
PRIN=[];b=0
for i in range(NY): b+=DRAW[i]+AMORT[i]+BULLET[i]; PRIN.append(b)
IRR  = [0.1276724398136139,0.11278300881385803,0.10125557780265806,0.09331141412258148]
MOIC = [2.6594491687737376,2.4317966870678105,2.1183242151856,2.1183242151856]
RET_LABELS = ['Gross Levered','Net Levered','Gross Unlevered','Net Unlevered']

# ---------- helper data block (cols Y.. ; data cols AA..AK) ----------
LBLCOL=25; C0=27; YEAR_ROW=4
def dcol(i): return col(C0+i)
rows_xml={}
def put(r,c,xml): rows_xml.setdefault(r,{})[c]=xml
def esc(t): return t.replace('&','&amp;').replace('<','&lt;').replace('>','&gt;')
def escA(t): return esc(t).replace('"','&quot;')
def cnum(ref,f,v): return f'<c r="{ref}"><f>{esc(f)}</f><v>{v}</v></c>'
def cstr(ref,t): return f'<c r="{ref}" t="inlineStr"><is><t xml:space="preserve">{esc(t)}</t></is></c>'

put(2,LBLCOL,cstr(f'{col(LBLCOL)}2','CHART SOURCE DATA - Stabilized Portfolio (USD) [auto-linked to model]'))
put(YEAR_ROW,LBLCOL,cstr(f'{col(LBLCOL)}{YEAR_ROW}','Year'))
for i,y in enumerate(YEARS):
    put(YEAR_ROW,C0+i,cnum(f'{dcol(i)}{YEAR_ROW}',f'{USD}!{col(5+i)}1',y))
def usd_f(row): return lambda i:f'{USD}!{col(5+i)}{row}'
def debt_f(row): return lambda i:f'{DEBT}!{col(4+i)}{row}'
series=[(5,'Net Operating Income',NOI,usd_f(29)),(6,'Net Exit Proceeds',NETEXIT,usd_f(32)),
        (7,'Contributions',CONTRIB,usd_f(48)),(8,'Distributions',DISTRIB,usd_f(49)),
        (9,'Net Levered Cash Flow',NETCF,usd_f(84)),(11,'Debt Drawdown',DRAW,debt_f(34)),
        (12,'Interest Expense',INTEREST,debt_f(37)),(13,'Scheduled Amortization',AMORT,debt_f(35)),
        (14,'Bullet Repayment',BULLET,debt_f(36))]
for rn,lab,vals,fn in series:
    put(rn,LBLCOL,cstr(f'{col(LBLCOL)}{rn}',lab))
    for i in range(NY): put(rn,C0+i,cnum(f'{dcol(i)}{rn}',fn(i),vals[i]))
put(10,LBLCOL,cstr(f'{col(LBLCOL)}10','Cumulative Net Levered CF'))
for i in range(NY):
    f=f'{dcol(0)}9' if i==0 else f'{dcol(i-1)}10+{dcol(i)}9'
    put(10,C0+i,cnum(f'{dcol(i)}10',f,CUM[i]))
put(15,LBLCOL,cstr(f'{col(LBLCOL)}15','Debt Principal Balance (EoP)'))
for i in range(NY):
    base=f'{dcol(i)}11+{dcol(i)}13+{dcol(i)}14'
    f=base if i==0 else f'{dcol(i-1)}15+{base}'
    put(15,C0+i,cnum(f'{dcol(i)}15',f,PRIN[i]))
put(17,LBLCOL,cstr(f'{col(LBLCOL)}17','Returns'))
for i,lab in enumerate(RET_LABELS): put(17,C0+i,cstr(f'{dcol(i)}17',lab))
for key,rn,fs,vals in [('IRR',18,[f'{USD}!D58',f'{USD}!D100',f'{USD}!J58',f'{USD}!J100'],IRR),
                       ('MOIC',19,[f'{USD}!D59',f'{USD}!D101',f'{USD}!J59',f'{USD}!J101'],MOIC)]:
    put(rn,LBLCOL,cstr(f'{col(LBLCOL)}{rn}',key))
    for i in range(4): put(rn,C0+i,cnum(f'{dcol(i)}{rn}',fs[i],vals[i]))

def helper_rows_xml():
    out=[]
    for r in sorted(rows_xml):
        cs=rows_xml[r]
        out.append(f'<row r="{r}" spans="1:186">{"".join(cs[c] for c in sorted(cs))}</row>')
    return ''.join(out)

# ---------------- chart XML generation ----------------
CNS='xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
def rng(c1,c2,r): return f"'{GR}'!${c1}${r}:${c2}${r}"
def cellref(c,r): return f"'{GR}'!${c}${r}"
def numcache(vals,fmt='General'):
    pts=''.join(f'<c:pt idx="{i}"><c:v>{v}</c:v></c:pt>' for i,v in enumerate(vals))
    return f'<c:numCache><c:formatCode>{fmt}</c:formatCode><c:ptCount val="{len(vals)}"/>{pts}</c:numCache>'
def strcache(vals):
    pts=''.join(f'<c:pt idx="{i}"><c:v>{esc(str(v))}</c:v></c:pt>' for i,v in enumerate(vals))
    return f'<c:strCache><c:ptCount val="{len(vals)}"/>{pts}</c:strCache>'

def cat_block(ref,vals,is_str):
    if is_str:
        return f'<c:cat><c:strRef><c:f>{ref}</c:f>{strcache(vals)}</c:strRef></c:cat>'
    return f'<c:cat><c:numRef><c:f>{ref}</c:f>{numcache(vals)}</c:numRef></c:cat>'

def ser_block(idx,name,name_ref,val_ref,vals,is_line,cat_ref,cat_vals,cat_is_str,valfmt):
    tx=f'<c:tx><c:strRef><c:f>{name_ref}</c:f>{strcache([name])}</c:strRef></c:tx>'
    marker='<c:marker><c:symbol val="circle"/><c:size val="5"/></c:marker>' if is_line else ''
    val=f'<c:val><c:numRef><c:f>{val_ref}</c:f>{numcache(vals,valfmt)}</c:numRef></c:val>'
    smooth='<c:smooth val="0"/>' if is_line else ''
    return f'<c:ser><c:idx val="{idx}"/><c:order val="{idx}"/>{tx}{marker}{cat_block(cat_ref,cat_vals,cat_is_str)}{val}{smooth}</c:ser>'

CATAX=111111111; VALAX=222222222
def chart_xml(title,ctype,sers,valfmt='General'):
    is_line=(ctype=='line')
    plot=''.join(sers)
    if is_line:
        body=f'<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>{plot}<c:marker val="1"/><c:axId val="{CATAX}"/><c:axId val="{VALAX}"/></c:lineChart>'
    else:
        body=f'<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>{plot}<c:gapWidth val="80"/><c:axId val="{CATAX}"/><c:axId val="{VALAX}"/></c:barChart>'
    catax=f'<c:catAx><c:axId val="{CATAX}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:crossAx val="{VALAX}"/></c:catAx>'
    valax=f'<c:valAx><c:axId val="{VALAX}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:numFmt formatCode="{escA(valfmt)}" sourceLinked="0"/><c:crossAx val="{CATAX}"/></c:valAx>'
    titlexml=f'<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1200" b="1"/></a:pPr><a:r><a:rPr lang="en-US" sz="1200" b="1"/><a:t>{esc(title)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>'
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            f'<c:chartSpace {CNS}><c:chart>{titlexml}<c:autoTitleDeleted val="0"/>'
            f'<c:plotArea><c:layout/>{body}{catax}{valax}<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr></c:plotArea>'
            f'<c:legend><c:legendPos val="b"/><c:overlay val="0"/></c:legend>'
            f'<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart></c:chartSpace>')

# year cat ref / vals
YREF=rng('AA','AK',YEAR_ROW); YV=YEARS
RREF=rng('AA','AD',17); RV=RET_LABELS
USD0='#,##0,,"M"'   # millions display
PCT='0.0%'; MULT='0.0"x"'

def line_one(title,row,label):
    return chart_xml(title,'line',[ser_block(0,label,cellref('Y',row),rng('AA','AK',row),globals_vals(row),True,YREF,YV,False,USD0)],USD0)

def globals_vals(row):
    m={5:NOI,6:NETEXIT,7:CONTRIB,8:DISTRIB,9:NETCF,10:CUM,11:DRAW,12:INTEREST,13:AMORT,14:BULLET,15:PRIN}
    return m[row]

charts=[]   # (title, xml)
# 1 NOI line
charts.append(('Stabilized Portfolio - Net Operating Income (USD)',
    chart_xml('Stabilized Portfolio - Net Operating Income (USD)','line',
      [ser_block(0,'NOI',cellref('Y',5),rng('AA','AK',5),NOI,True,YREF,YV,False,USD0)],USD0)))
# 2 Contributions vs Distributions bar
charts.append(('Levered Cash Flow - Contributions vs Distributions (USD)',
    chart_xml('Levered Cash Flow - Contributions vs Distributions (USD)','bar',
      [ser_block(0,'Contributions',cellref('Y',7),rng('AA','AK',7),CONTRIB,False,YREF,YV,False,USD0),
       ser_block(1,'Distributions',cellref('Y',8),rng('AA','AK',8),DISTRIB,False,YREF,YV,False,USD0)],USD0)))
# 3 J-curve line
charts.append(('Cumulative Net Levered Cash Flow (J-Curve, USD)',
    chart_xml('Cumulative Net Levered Cash Flow (J-Curve, USD)','line',
      [ser_block(0,'Cumulative Net Levered CF',cellref('Y',10),rng('AA','AK',10),CUM,True,YREF,YV,False,USD0)],USD0)))
# 4 Debt service bar (interest, amort, bullet)
charts.append(('Annual Debt Service (USD)',
    chart_xml('Annual Debt Service (USD)','bar',
      [ser_block(0,'Interest Expense',cellref('Y',12),rng('AA','AK',12),INTEREST,False,YREF,YV,False,USD0),
       ser_block(1,'Scheduled Amortization',cellref('Y',13),rng('AA','AK',13),AMORT,False,YREF,YV,False,USD0),
       ser_block(2,'Bullet Repayment',cellref('Y',14),rng('AA','AK',14),BULLET,False,YREF,YV,False,USD0)],USD0)))
# 5 Principal balance line
charts.append(('Debt Principal Balance Outstanding (USD)',
    chart_xml('Debt Principal Balance Outstanding (USD)','line',
      [ser_block(0,'Principal Balance',cellref('Y',15),rng('AA','AK',15),PRIN,True,YREF,YV,False,USD0)],USD0)))
# 6 Exit proceeds bar
charts.append(('Net Exit Proceeds by Year (USD)',
    chart_xml('Net Exit Proceeds by Year (USD)','bar',
      [ser_block(0,'Net Exit Proceeds',cellref('Y',6),rng('AA','AK',6),NETEXIT,False,YREF,YV,False,USD0)],USD0)))
# 7 IRR bar
charts.append(('Internal Rate of Return by Strategy',
    chart_xml('Internal Rate of Return by Strategy','bar',
      [ser_block(0,'IRR',cellref('Y',18),rng('AA','AD',18),IRR,False,RREF,RV,True,PCT)],PCT)))
# 8 MOIC bar
charts.append(('Equity Multiple (MOIC) by Strategy',
    chart_xml('Equity Multiple (MOIC) by Strategy','bar',
      [ser_block(0,'MOIC',cellref('Y',19),rng('AA','AD',19),MOIC,False,RREF,RV,True,MULT)],MULT)))

print('built',len(charts),'charts')

# ---------------- drawing anchors ----------------
def anchor(fc,fr,tc,tr,cid,name,rid):
    return (f'<xdr:twoCellAnchor>'
            f'<xdr:from><xdr:col>{fc}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>{fr}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>'
            f'<xdr:to><xdr:col>{tc}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>{tr}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>'
            f'<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="{cid}" name="{esc(name)}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>'
            f'<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>'
            f'<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">'
            f'<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="{rid}"/>'
            f'</a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>')

positions=[]
r=32
for k in range(4):
    positions.append((1,r,10,r+15))      # left
    positions.append((12,r,21,r+15))     # right
    r+=17
anchors=''; drels=''; cts=''
chart_start=3  # chart1,chart2 already exist
new_parts={}
for n,(title,xml) in enumerate(charts):
    cn=chart_start+n
    part=f'xl/charts/chart{cn}.xml'
    new_parts[part]=xml
    rid=f'rId{100+n}'
    fc,fr,tc,tr=positions[n]
    anchors+=anchor(fc,fr,tc,tr,200+n,f'Chart {cn}',rid)
    drels+=f'<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart{cn}.xml"/>'
    cts+=f'<Override PartName="/xl/charts/chart{cn}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>'

# ---------------- assemble new zip ----------------
zin=zipfile.ZipFile(SRC)
names=zin.namelist()
sheet3=zin.read('xl/worksheets/sheet3.xml').decode('utf8')
sheet3=sheet3.replace('</sheetData>', helper_rows_xml()+'</sheetData>')
sheet3=sheet3.replace('<dimension ref="A1:GD48"/>','<dimension ref="A1:GD98"/>')
draw1=zin.read('xl/drawings/drawing1.xml').decode('utf8')
draw1=draw1.replace('</xdr:wsDr>', anchors+'</xdr:wsDr>')
drawrels=zin.read('xl/drawings/_rels/drawing1.xml.rels').decode('utf8')
drawrels=drawrels.replace('</Relationships>', drels+'</Relationships>')
ct=zin.read('[Content_Types].xml').decode('utf8')
ct=ct.replace('</Types>', cts+'</Types>')

repl={'xl/worksheets/sheet3.xml':sheet3,'xl/drawings/drawing1.xml':draw1,
      'xl/drawings/_rels/drawing1.xml.rels':drawrels,'[Content_Types].xml':ct}

zout=zipfile.ZipFile(OUT,'w',zipfile.ZIP_DEFLATED)
for item in zin.infolist():
    data=repl.get(item.filename)
    if data is not None: zout.writestr(item,data)
    else: zout.writestr(item,zin.read(item.filename))
for part,xml in new_parts.items():
    zout.writestr(part,xml)
zout.close(); zin.close()
print('WROTE',OUT)
