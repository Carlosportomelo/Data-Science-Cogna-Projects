import pandas as pd
P='outputs/visão resumida/hubspot_resumido_Dash.xlsx'
df=pd.read_excel(P,sheet_name='Blend_Agregado_Dash')
df.columns=[c.strip() for c in df.columns]
invest_col=next((c for c in df.columns if 'invest' in c.lower() or 'valor' in c.lower()), None)
vol_col=next((c for c in df.columns if 'volume' in c.lower() or 'negocio' in c.lower()), None)
date_col=next((c for c in df.columns if 'data' in c.lower()), None)
canal_col=next((c for c in df.columns if 'canal' in c.lower() or 'origem' in c.lower()), None)
paid=['social pago','pesquisa paga','facebook','instagram','paid social','google ads','paid search','cpc']
df['canal_clean']=df[canal_col].astype(str).str.lower() if canal_col else ''
mask_paid=df['canal_clean'].isin(paid)
df[invest_col]=pd.to_numeric(df[invest_col].astype(str).str.replace(',',''),errors='coerce').fillna(0)
df[vol_col]=pd.to_numeric(df[vol_col].astype(str).str.replace(',',''),errors='coerce').fillna(0)
# total
total_invest=df.loc[mask_paid,invest_col].sum()
total_leads=df.loc[mask_paid,vol_col].sum()
cpl=total_invest/total_leads if total_leads>0 else None
print('Investimento pago: R$ {:,.2f}'.format(total_invest))
print('Leads (Volume_Total_Negocios) pago: {}'.format(int(total_leads)))
print('CPL pago: R$ {:,.2f}'.format(cpl) if cpl else 'CPL indefinido')
# jan 2026
if date_col:
    df[date_col]=pd.to_datetime(df[date_col],errors='coerce')
    m=df[(df[date_col].dt.year==2026)&(df[date_col].dt.month==1)&mask_paid]
    if not m.empty:
        ti=m[invest_col].sum(); tl=m[vol_col].sum()
        print('\nJan/2026 -> Invest: R$ {:,.2f}'.format(ti))
        print('Jan/2026 -> Leads: {}'.format(int(tl)))
        print('Jan/2026 CPL: R$ {:,.2f}'.format(ti/tl) if tl>0 else 'Jan/2026 CPL indefinido')
