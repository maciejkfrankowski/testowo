import openpyxl, datetime, sys
plik = sys.argv[1] if len(sys.argv)>1 else 'NEU.xlsx'
rows=list(openpyxl.load_workbook(plik,read_only=True,data_only=True)['Spr Fin'].iter_rows(min_row=1,max_row=110,values_only=True))
K={str(r[0]).strip():r for r in rows if r[0] and isinstance(r[0],str)}
hdr=rows[0]; daty=K['balance_date']
kol=[i for i in range(len(hdr)) if isinstance(daty[i],datetime.datetime) and isinstance(K['revenues'][i],(int,float)) and K['revenues'][i]]
print(f'{plik}: kolumn z danymi {len(kol)}  ({hdr[kol[0]]} ... {hdr[kol[-1]]})')

g=lambda k,i:(K[k][i] or 0) if isinstance(K.get(k,[None]*99)[i],(int,float)) else 0
def wc(i,j):
    dlug=lambda x: g('current_liabilities',x)-g('current_other_liabilities',x)
    inne=lambda x: g('current_other_liabilities',x)+g('reckoning',x)
    return ((g('inventory',j)-g('inventory',i))+(g('current_receivables',j)-g('current_receivables',i))
            +(dlug(i)-dlug(j))+(inne(i)-inne(j)))

# 1. narastanie przychodow w obrebie roku
lata={}
for i in kol: lata.setdefault(daty[i].year,[]).append(i)
rosnie=lam=0
for rok,ii in lata.items():
    if len(ii)<2: continue
    v=[K['revenues'][x] for x in sorted(ii,key=lambda x:daty[x])]
    if all(v[k]<=v[k+1] for k in range(len(v)-1)): rosnie+=1
    else: lam+=1
print(f'1. RZiS narastajaco: lat zgodnych {rosnie}, lamiacych {lam}')

# 2. kapital obrotowy: od konca roku czy od poprzedniego kwartalu
zg_rok=zg_kw=ani=0
for poz,i in enumerate(kol):
    if poz==0 or K['change_in_working_capital'][i] is None: continue
    w=K['change_in_working_capital'][i]
    if not isinstance(w,(int,float)): continue
    prev=kol[poz-1]
    rok=[x for x in kol[:poz] if daty[x].month==12]
    ry=rok[-1] if rok else prev
    a,b=wc(i,prev),wc(i,ry)
    if b==w: zg_rok+=1
    elif a==w: zg_kw+=1
    else: ani+=1
print(f'2. Kapital obrotowy: od KONCA ROKU {zg_rok}, od poprz. kwartalu {zg_kw}, zaden {ani}')

# 3. pola meta
for k in ['currency_id','year_profit','own_share','discontinued_profit','extraordinary_profit','share_amount','raport_date']:
    if k not in K: continue
    v=[K[k][i] for i in kol]
    nie=sum(1 for x in v if x not in (None,0,''))
    prz=sorted({str(x)[:12] for x in v if x not in (None,'')})[:4]
    print(f'3. {k:22} wypelnione w {nie}/{len(v)}  przyklady: {prz}')

# 4. straty
straty=[hdr[i] for i in kol if isinstance(K['net_profit'][i],(int,float)) and K['net_profit'][i]<0]
print(f'4. okresy ze STRATA netto: {len(straty)}  {straty[:8]}')
