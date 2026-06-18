#!/usr/bin/env python3
"""Add Stabilized-Portfolio sensitivity tables + tornado chart to the
Sensitivities tab. Operates on the charts file, writes the final file.
All values come from a return-engine validated to reproduce the model's
exact base case (Gross Levered IRR 12.767%, MOIC 2.6594x)."""
import zipfile, datetime
from openpyxl.utils import get_column_letter as gcol

IN  = 'UAE_Industrial_Strategy_Financial_Model_0626_charts.xlsx'
OUT = 'UAE_Industrial_Strategy_Financial_Model_0626_FINAL.xlsx'
SH  = 'Sensitivities'

# ---------------- validated return engine ----------------
dates=[datetime.date(y,12,31) for y in range(2026,2037)]; d0=dates[0]
days=[(d-d0).days/365.0 for d in dates]; N=11
Acq=[-429073569,0,0,0,0,0,0,0,0,0,0]
NOI=[29587613,30793546,32507856,33046715,33896556,35395796,36443511,37522239,38632897,39776431,24868384]
Tax=[-2662885,-2771419,-2925707,-2974204,-3050690,-3185622,-3279916,-3377002,-3476961,-8981980,-7792681]
Exit=[0,0,0,0,0,0,0,0,0,195200459,355613421]
Draw=[214536785,0,0,0,0,0,0,0,0,0,0]
Int=[-2424041,-14375960,-14087469,-13798979,-13510489,-13221999,-12933508,-12645018,-12356528,-12068037,-11779547]
Amort=[-708650,-4251899,-4251899,-4251899,-4251899,-4251899,-4251899,-4251899,-4251899,-4251899,-4251899]
Repay=[0,0,0,0,0,0,0,0,0,0,-171309142]
BASE_CAP,BASE_LTV,BASE_RATE=0.075,0.5,0.07
exit_tax=[0.0]*N; exit_tax[9]=Tax[9]-(-3580000.0); exit_tax[10]=Tax[10]-(-2200000.0)
def xnpv(r,cf): return sum(cf[i]/(1+r)**days[i] for i in range(N))
def xirr(cf):
    lo,hi=-0.99,2.0; flo=xnpv(lo,cf)
    for _ in range(200):
        mid=(lo+hi)/2; fm=xnpv(mid,cf)
        if abs(fm)<1e-4: return mid
        if (flo<0)==(fm<0): lo,flo=mid,fm
        else: hi=mid
    return mid
def moic(cf):
    return sum(x for x in cf if x>0)/-sum(x for x in cf if x<0)
def lev_cf(cap=BASE_CAP,ltv=BASE_LTV,rate=BASE_RATE):
    capf=BASE_CAP/cap; ltvf=ltv/BASE_LTV; ratef=rate/BASE_RATE; cf=[]
    for i in range(N):
        ex=Exit[i]*capf; tx=Tax[i]+exit_tax[i]*(capf-1)
        unlev=Acq[i]+NOI[i]+tx+ex
        debt=(Draw[i]+Int[i]*ratef+Amort[i]+Repay[i])*ltvf
        cf.append(unlev+debt)
    return cf
def irr(**k): return round(xirr(lev_cf(**k))*100,2)
def mo(**k): return round(moic(lev_cf(**k)),2)

CAPS=[0.060,0.065,0.070,0.075,0.080,0.085,0.090]
LTVS=[0.40,0.45,0.50,0.55,0.60,0.65]
RATES=[0.05,0.06,0.07,0.08,0.09]
LTVS2=[0.40,0.45,0.50,0.55,0.60]

# ---------------- cell helpers ----------------
cells={}
def put(r,c,xml): cells.setdefault(r,{})[c]=xml
def esc(t): return str(t).replace('&','&amp;').replace('<','&lt;').replace('>','&gt;')
def S(r,c,t): put(r,c,f'<c r="{gcol(c)}{r}" t="inlineStr"><is><t xml:space="preserve">{esc(t)}</t></is></c>')
def Nn(r,c,v): put(r,c,f'<c r="{gcol(c)}{r}"><v>{v}</v></c>')

# ---------------- layout ----------------
S(50,3,'SENSITIVITY ANALYSIS  -  Stabilized Portfolio (Gross Levered Returns)')
S(51,3,'Base case: Exit Cap 7.5%, LTV 50%, Interest 7.0%  ->  IRR 12.77% , MOIC 2.66x .  Net Levered runs ~1.5pts below Gross.')
S(52,3,'Values computed from the model\'s own cash flows (engine reproduces base case exactly). Cap rate is the dominant driver.')

# One-way: Exit Cap Rate (cols C,D,E)
r0=55
S(r0-1,3,'Exit Cap Rate Sensitivity')
S(r0,3,'Cap Rate (%)'); S(r0,4,'IRR (%)'); S(r0,5,'MOIC (x)')
for i,c in enumerate(CAPS):
    Nn(r0+1+i,3,round(c*100,1)); Nn(r0+1+i,4,irr(cap=c)); Nn(r0+1+i,5,mo(cap=c))
    if abs(c-BASE_CAP)<1e-9: S(r0+1+i,6,'<- base')

# One-way: LTV (cols H,I,J)
S(r0-1,8,'Loan-to-Value (LTV) Sensitivity')
S(r0,8,'LTV (%)'); S(r0,9,'IRR (%)'); S(r0,10,'MOIC (x)')
for i,l in enumerate(LTVS):
    Nn(r0+1+i,8,round(l*100,0)); Nn(r0+1+i,9,irr(ltv=l)); Nn(r0+1+i,10,mo(ltv=l))
    if abs(l-BASE_LTV)<1e-9: S(r0+1+i,11,'<- base')

# One-way: Interest Rate (cols M,N,O)
S(r0-1,13,'Interest Rate Sensitivity')
S(r0,13,'Interest (%)'); S(r0,14,'IRR (%)'); S(r0,15,'MOIC (x)')
for i,rt in enumerate(RATES):
    Nn(r0+1+i,13,round(rt*100,0)); Nn(r0+1+i,14,irr(rate=rt)); Nn(r0+1+i,15,mo(rate=rt))
    if abs(rt-BASE_RATE)<1e-9: S(r0+1+i,16,'<- base')

# Two-way: Cap (rows) x LTV (cols) -> IRR
t0=66
S(t0-1,3,'Two-Way: Gross Levered IRR (%)  -  Exit Cap Rate (down) x LTV (across)')
S(t0,3,'Cap \\ LTV')
for j,l in enumerate(LTVS2): Nn(t0,4+j,round(l*100,0))
for i,c in enumerate(CAPS):
    Nn(t0+1+i,3,round(c*100,1))
    for j,l in enumerate(LTVS2):
        Nn(t0+1+i,4+j,irr(cap=c,ltv=l))

# Tornado data block (cols R..U ; 18..21) ordered narrowest->widest so widest on top
def swing(name,down,up):
    return (name,round(down,2),round(up,2),round(abs(up-down),2))
tor=[swing('LTV (40% / 65%)', irr(ltv=0.40), irr(ltv=0.65)),
     swing('Interest (9% / 5%)', irr(rate=0.09), irr(rate=0.05)),
     swing('Exit Cap (9% / 6%)', irr(cap=0.090), irr(cap=0.060))]
tor.sort(key=lambda x:x[3])   # ascending swing
TR=55
S(TR-1,18,'Tornado data (Gross Lev IRR %)')
S(TR,18,'Driver'); S(TR,19,'Downside'); S(TR,20,'Upside'); S(TR,21,'Swing')
for i,(nm,dn,up,sw) in enumerate(tor):
    S(TR+1+i,18,nm); Nn(TR+1+i,19,dn); Nn(TR+1+i,20,up); Nn(TR+1+i,21,sw)
BASE_IRR=12.77
Nn(TR+5,19,BASE_IRR); S(TR+5,18,'Base IRR')

print('engine base check: IRR',irr(),'MOIC',mo())
print('tornado order:',[t[0] for t in tor])

# ---------------- tornado chart XML ----------------
CNS='xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
def ref(c1,c2,r): return f"'{SH}'!${gcol(c1)}${r}:${gcol(c2)}${r}"
def vref(c,r1,r2): return f"'{SH}'!${gcol(c)}${r1}:${gcol(c)}${r2}"
cat_names=[t[0] for t in tor]
def strcache(vals):
    p=''.join(f'<c:pt idx="{i}"><c:v>{esc(v)}</c:v></c:pt>' for i,v in enumerate(vals))
    return f'<c:strCache><c:ptCount val="{len(vals)}"/>{p}</c:strCache>'
def numcache(vals):
    p=''.join(f'<c:pt idx="{i}"><c:v>{v}</c:v></c:pt>' for i,v in enumerate(vals))
    return f'<c:numCache><c:formatCode>0.0</c:formatCode><c:ptCount val="{len(vals)}"/>{p}</c:numCache>'
catref=vref(18,TR+1,TR+3); catvals=cat_names
def ser(idx,name,namecol,valcol,vals):
    tx=f'<c:tx><c:strRef><c:f>{cellref(namecol)}</c:f>{strcache([name])}</c:strRef></c:tx>'
    cat=f'<c:cat><c:strRef><c:f>{catref}</c:f>{strcache(catvals)}</c:strRef></c:cat>'
    val=f'<c:val><c:numRef><c:f>{vref(valcol,TR+1,TR+3)}</c:f>{numcache(vals)}</c:numRef></c:val>'
    return f'<c:ser><c:idx val="{idx}"/><c:order val="{idx}"/>{tx}{cat}{val}</c:ser>'
def cellref(c): return f"'{SH}'!${gcol(c)}${TR}"
CATAX=333333331; VALAX=333333332
downs=[t[1] for t in tor]; ups=[t[2] for t in tor]
sers=ser(0,'Downside (adverse)',19,19,downs)+ser(1,'Upside (favourable)',20,20,ups)
bar=f'<c:barChart><c:barDir val="bar"/><c:grouping val="clustered"/><c:varyColors val="0"/>{sers}<c:gapWidth val="60"/><c:axId val="{CATAX}"/><c:axId val="{VALAX}"/></c:barChart>'
catax=f'<c:catAx><c:axId val="{CATAX}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:crossAx val="{VALAX}"/></c:catAx>'
valax=f'<c:valAx><c:axId val="{VALAX}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:numFmt formatCode="0.0" sourceLinked="0"/><c:crossAx val="{CATAX}"/></c:valAx>'
title='<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1200" b="1"/></a:pPr><a:r><a:rPr lang="en-US" sz="1200" b="1"/><a:t>Tornado - Gross Levered IRR Sensitivity (%)</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>'
tornado_xml=(f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<c:chartSpace {CNS}><c:chart>{title}<c:autoTitleDeleted val="0"/>'
    f'<c:plotArea><c:layout/>{bar}{catax}{valax}<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr></c:plotArea>'
    f'<c:legend><c:legendPos val="b"/><c:overlay val="0"/></c:legend><c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart></c:chartSpace>')

# ---------------- assemble ----------------
def rows_xml():
    out=[]
    for r in sorted(cells):
        cs=cells[r]; out.append(f'<row r="{r}" spans="1:186">{"".join(cs[c] for c in sorted(cs))}</row>')
    return ''.join(out)

zin=zipfile.ZipFile(IN)
s4=zin.read('xl/worksheets/sheet4.xml').decode('utf8')
s4=s4.replace('</sheetData>', rows_xml()+'</sheetData>').replace('<dimension ref="A1:GD48"/>','<dimension ref="A1:GD75"/>')
d2=zin.read('xl/drawings/drawing2.xml').decode('utf8')
anchor=('<xdr:twoCellAnchor><xdr:from><xdr:col>16</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>65</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>'
        '<xdr:to><xdr:col>26</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>83</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>'
        '<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="300" name="Tornado"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>'
        '<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>'
        '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">'
        '<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId200"/>'
        '</a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>')
d2=d2.replace('</xdr:wsDr>', anchor+'</xdr:wsDr>')
d2rels=zin.read('xl/drawings/_rels/drawing2.xml.rels').decode('utf8')
d2rels=d2rels.replace('</Relationships>','<Relationship Id="rId200" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart11.xml"/></Relationships>')
ct=zin.read('[Content_Types].xml').decode('utf8')
ct=ct.replace('</Types>','<Override PartName="/xl/charts/chart11.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/></Types>')
repl={'xl/worksheets/sheet4.xml':s4,'xl/drawings/drawing2.xml':d2,'xl/drawings/_rels/drawing2.xml.rels':d2rels,'[Content_Types].xml':ct}
zout=zipfile.ZipFile(OUT,'w',zipfile.ZIP_DEFLATED)
for it in zin.infolist():
    zout.writestr(it, repl.get(it.filename, zin.read(it.filename)))
zout.writestr('xl/charts/chart11.xml', tornado_xml)
zout.close(); zin.close()
print('WROTE',OUT)
