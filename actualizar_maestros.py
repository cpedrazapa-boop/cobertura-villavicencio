import openpyxl, json, sys
wb=openpyxl.load_workbook(sys.argv[1],data_only=True,read_only=True)
ws=wb['Data']; it=ws.iter_rows(values_only=True); hdr=list(next(it))
STATIC=['Tipo PT','Marca','Lines','Centro','Material','Descripcion','UM','tipo','presentacio','conte',
 'Cant entrega Pendien','Cantidad en Pedidos','FC TOTAL CD','Forescast KA','Volumen Solicitado','Volumen recibido2',
 'Entregado TAT','Entregado KA','Entregado HE','Entregado Socios','Fecha Prox Prod','Planta','Htls Prod Plan','Riesgo traspasos']
idx={c:hdr.index(c) for c in STATIC}
bd={}
for r in wb['BD PT'].iter_rows(min_row=2,values_only=True):
    if r[0] is not None and r[0] not in bd: bd[r[0]]=(r[2],r[3])
def clean(v):
    if v is None: return None
    if isinstance(v,str) and v.startswith('#'): return None
    if hasattr(v,'isoformat'): return str(v)[:10]
    return v
items=[]
for r in it:
    if r[idx['Material']] is None: continue
    o={c:clean(r[idx[c]]) for c in STATIC}
    for c in ['Marca','Lines','UM','tipo','Tipo PT']:
        v=o[c]; o[c]='' if (v in (None,0)) else str(v).strip()
    o['UM']=o['UM'].upper(); o['tipo']=o['tipo'].replace('Tw','TW').replace('pet','Pet')
    c3,d4=bd.get(o['Material'],(None,None)); o['_cjEst']=c3; o['_undEst']=d4
    items.append(o)
json.dump({'items':items},open(sys.argv[2],'w'),ensure_ascii=False,separators=(',',':'))
print(len(items))
