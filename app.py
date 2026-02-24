import os, re, shutil, datetime, uuid, json, copy, calendar
from pathlib import Path
from collections import defaultdict
from flask import Flask, render_template, request, jsonify, send_file, after_this_request

try:
    import pandas as pd
    from openpyxl import load_workbook, Workbook
    from openpyxl.styles import Font, Alignment, PatternFill
    import xlrd
except ImportError:
    pass

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024
UPLOAD_FOLDER = '/tmp/plannipro'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

MOIS_ABBR = {
    'août':8,'aout':8,'sept':9,'septembre':9,'oct':10,'octobre':10,
    'nov':11,'novembre':11,'déc':12,'dec':12,'décembre':12,
    'janv':1,'jan':1,'janvier':1,'févr':2,'fev':2,'février':2,
    'mars':3,'avr':4,'avril':4,'mai':5,'juin':6,'juil':7,'juillet':7
}
MOIS_FR = ['','Janvier','Février','Mars','Avril','Mai','Juin',
           'Juillet','Août','Septembre','Octobre','Novembre','Décembre']

FERIES_2026_2027 = {
    datetime.date(2026,11,1), datetime.date(2026,11,11), datetime.date(2026,12,25),
    datetime.date(2027,1,1),  datetime.date(2027,4,5),   datetime.date(2027,5,1),
    datetime.date(2027,5,8),  datetime.date(2027,5,13),  datetime.date(2027,5,24),
    datetime.date(2027,7,14), datetime.date(2027,8,15),
}

DEFAULT_COLORS = [
    'FFEE7E32','FFFFCC99','FFFFD243','FFDDFFDD',
    'FFB8D4F0','FFFFC0CB','FFD4B8F0','FFB8F0D4',
]

# ── Parsers planning classe ──────────────────
def colors_match(rgb):
    if not rgb: return False
    return abs(rgb[0]-166)<20 and abs(rgb[1]-202)<20 and abs(rgb[2]-240)<20

def parse_planning_classe_xls(filepath):
    try:
        wb = xlrd.open_workbook(filepath, formatting_info=True)
    except Exception as e:
        return {'nom': Path(filepath).stem, 'jours': []}
    ws = wb.sheet_by_index(0)
    nom = next((ws.cell_value(0,c) for c in range(ws.ncols) if isinstance(ws.cell_value(0,c),str) and len(ws.cell_value(0,c).strip())>3), Path(filepath).stem)
    jours = []
    blocks = []
    for c in range(ws.ncols):
        v = ws.cell_value(4,c)
        if isinstance(v,str):
            for m,mn in MOIS_ABBR.items():
                if m in v.lower():
                    yr = re.search(r'(20\d\d)',v)
                    if yr: blocks.append({'cj':c+1,'cd':c+2,'y':int(yr.group(1)),'m':mn})
                    break
    for b in blocks:
        for r in range(5,ws.nrows):
            if b['cj']>=ws.ncols or b['cd']>=ws.ncols: continue
            try:
                xf=wb.xf_list[ws.cell_xf_index(r,b['cj'])]
                rgb=wb.colour_map.get(xf.background.pattern_colour_index)
                if not colors_match(rgb): continue
            except: continue
            dv=ws.cell_value(r,b['cd'])
            dn=int(dv) if isinstance(dv,float) and dv>0 else (int(dv.strip()) if isinstance(dv,str) and dv.strip().isdigit() else None)
            if not dn or not (1<=dn<=31): continue
            try:
                d=datetime.date(b['y'],b['m'],dn)
                if d.weekday()<5: jours.append(d.strftime('%Y-%m-%d'))
            except: pass
    return {'nom':str(nom).strip(),'jours':sorted(set(jours))}

def parse_planning_classe_xlsx(filepath):
    try:
        wb=load_workbook(filepath,data_only=True)
    except: return {'nom':Path(filepath).stem,'jours':[]}
    ws=wb.active
    nom=next((ws.cell(row=1,column=c).value for c in range(1,ws.max_column+1) if ws.cell(row=1,column=c).value and isinstance(ws.cell(row=1,column=c).value,str) and len(ws.cell(row=1,column=c).value.strip())>3),Path(filepath).stem)
    jours=[]
    blocks=[]
    for c in range(1,ws.max_column+1):
        v=ws.cell(row=5,column=c).value
        if isinstance(v,str):
            for m,mn in MOIS_ABBR.items():
                if m in v.lower():
                    yr=re.search(r'(20\d\d)',v)
                    if yr: blocks.append({'cj':c+1,'cd':c+2,'y':int(yr.group(1)),'m':mn})
                    break
    for b in blocks:
        for r in range(6,ws.max_row+1):
            cell=ws.cell(row=r,column=b['cj'])
            bg=cell.fill.fgColor.rgb if cell.fill and cell.fill.fgColor and cell.fill.fgColor.type=='rgb' else None
            if bg not in ('FFA6CAF0','A6CAF0'): continue
            dv=ws.cell(row=r,column=b['cd']).value
            dn=dv.day if isinstance(dv,datetime.datetime) else (int(dv) if isinstance(dv,(int,float)) else (int(dv.strip()) if isinstance(dv,str) and dv.strip().isdigit() else None))
            if not dn or not (1<=dn<=31): continue
            try:
                d=datetime.date(b['y'],b['m'],dn)
                if d.weekday()<5: jours.append(d.strftime('%Y-%m-%d'))
            except: pass
    return {'nom':str(nom).strip(),'jours':sorted(set(jours))}

def parse_planning_classe(fp):
    return parse_planning_classe_xls(fp) if Path(fp).suffix.lower()=='.xls' else parse_planning_classe_xlsx(fp)

# ── Parser dispos ─────────────────────────────
def parse_disponibilite(filepath):
    try: df=pd.read_excel(filepath,sheet_name=0,header=None)
    except: return {'nom':Path(filepath).stem,'dispo':{}}
    nom=Path(filepath).stem; dispo={}
    def is_av(v): return False if v is None or (isinstance(v,float) and pd.isna(v)) else str(v).strip().upper() in ['X','4','✓','OUI','O']
    month_cols=[]
    for ci,val in enumerate(df.iloc[0]):
        if isinstance(val,str):
            for m,mn in MOIS_ABBR.items():
                if m in val.lower():
                    yr=re.search(r'(20\d\d)',val)
                    if yr: month_cols.append({'ci':ci,'m':mn,'y':int(yr.group(1))})
                    break
    ign=['sam','dim','férié','ferie','férie','fériés']
    for mc in month_cols:
        ci,mn,yr=mc['ci'],mc['m'],mc['y']
        if not yr: continue
        for ri in range(2,len(df)):
            row=df.iloc[ri]
            da=str(row.iloc[0]).strip().lower() if pd.notna(row.iloc[0]) else ''
            if da in ign: continue
            dr=row.iloc[1]
            dn=int(dr) if isinstance(dr,(int,float)) and not pd.isna(dr) else None
            if not dn or not (1<=dn<=31): continue
            try:
                d=datetime.date(yr,mn,dn)
                if d.weekday()>=5: continue
                ds=d.strftime('%Y-%m-%d')
                mv=row.iloc[ci] if ci<len(row) else None
                pv=row.iloc[ci+1] if ci+1<len(row) else None
                if str(mv).strip().lower() in ign: continue
                dispo[ds]={'matin':is_av(mv),'pm':is_av(pv)}
            except: pass
    return {'nom':nom,'dispo':dispo}

# ── Parser tableau formateurs ─────────────────
def parse_tableau_formateurs(filepath):
    assignments=defaultdict(list)
    def process(rows):
        cur=None
        for row in rows:
            v0=str(row[0]).strip() if row[0] else ''
            if re.match(r'(BTS|BAC|CGC|NDRC|GPME|Master|RDC|RH|EC)\s',v0,re.I): cur=v0
            elif cur and v0 and len(row)>1 and row[1]:
                mat=str(row[1]).strip()
                try: h=float(str(row[-1]).split('+')[0].strip())
                except: h=0
                if h>0 and mat and mat.lower() not in ['nan','total']:
                    assignments[cur].append({'formateur':v0,'matiere':mat,'heures':h,'heures_faites':0})
    ext=Path(filepath).suffix.lower()
    if ext=='.xls':
        try:
            wb=xlrd.open_workbook(filepath)
            si=next((i for i,n in enumerate(wb.sheet_names()) if 'mois' in n.lower()),len(wb.sheet_names())-1)
            ws=wb.sheet_by_index(si)
            process([[ws.cell_value(r,c) for c in range(ws.ncols)] for r in range(ws.nrows)])
        except: pass
    else:
        try:
            wb=load_workbook(filepath,data_only=True)
            sn=next((s for s in wb.sheetnames if 'mois' in s.lower()),wb.sheetnames[-1])
            process(list(wb[sn].iter_rows(values_only=True)))
        except: pass
    return dict(assignments)

# ── Moteur assignation ────────────────────────
def assigner(planning_classes,dispos_formateurs,affectations):
    dispo_idx={d['nom']:d['dispo'] for d in dispos_formateurs}
    result=defaultdict(dict); stats={'assigned':0,'warn':0}
    for ci in planning_classes:
        cn=ci['nom']; jours=ci['jours']
        aff=next((v for k,v in affectations.items() if k.strip().lower() in cn.lower() or cn.lower() in k.lower()),None)
        if not aff:
            for j in jours: result[j][cn]={'formateur':'?','matiere':'?','slot':'matin'}
            continue
        si=0
        for j in jours:
            slot=['matin','pm'][si%2]; si+=1; assigned=None
            for e in sorted(aff,key=lambda x:x['heures_faites']):
                pd_=dispo_idx.get(e['formateur'],{}); jd=pd_.get(j,{})
                if e['heures']-e['heures_faites']<=0: continue
                if not pd_ or jd.get(slot,False): assigned=e; break
            if assigned:
                assigned['heures_faites']+=4
                result[j][cn]={'formateur':assigned['formateur'],'matiere':assigned['matiere'],'slot':slot}
                stats['assigned']+=1
            else:
                result[j][cn]={'formateur':'⚠️','matiere':'','slot':slot}; stats['warn']+=1
    heures=defaultdict(lambda:defaultdict(int))
    for dv in result.values():
        for cl,inf in dv.items():
            if inf['formateur'] not in ['?','⚠️']: heures[inf['formateur']][inf['matiere']]+=4
    return dict(result),stats,dict(heures)

# ── Écriture planning assigné ─────────────────
def ecrire_planning(template_path,assignment,mois_cibles,output_path):
    shutil.copy(template_path,output_path)
    wb=load_workbook(output_path); filled=0
    kw=['BTS','BAC','EC ','CGC','NDRC','GPME','RDC','RH','Master']
    jf={'lundi':0,'mardi':1,'mercredi':2,'jeudi':3,'vendredi':4,'lun':0,'mar':1,'mer':2,'jeu':3,'ven':4}
    for sn in wb.sheetnames:
        ws=wb[sn]; titre=''
        for ri in range(1,4):
            for ci in range(1,ws.max_column+1):
                v=ws.cell(row=ri,column=ci).value
                if v and isinstance(v,str) and len(v.strip())>3: titre=v.upper(); break
        mn=None; yr=None
        for m,mnum in MOIS_ABBR.items():
            if m.upper() in titre:
                y=re.search(r'(20\d\d)',titre)
                if y: yr=int(y.group(1)); mn=mnum; break
        if not mn or f'{mn}/{yr}' not in mois_cibles: continue
        cc={}
        for ri in range(1,6):
            for ci in range(1,ws.max_column+1):
                v=ws.cell(row=ri,column=ci).value
                if isinstance(v,str) and any(k in v for k in kw): cc[v.strip()]=ci
        dr={}
        for ri in range(1,ws.max_row+1):
            for ci in range(1,4):
                v=ws.cell(row=ri,column=ci).value
                if isinstance(v,str) and v.strip().lower() in jf:
                    for lri in [ri,ri+1]:
                        for lci in range(1,4):
                            dv=ws.cell(row=lri,column=lci).value
                            if isinstance(dv,(int,float)) and 1<=int(dv)<=31:
                                try:
                                    d=datetime.date(yr,mn,int(dv)); ds=d.strftime('%Y-%m-%d')
                                    if ds not in dr: dr[ds]=ri
                                except: pass
        for ds,ri in dr.items():
            if ds not in assignment: continue
            for cn,ci in cc.items():
                for k,v in assignment[ds].items():
                    if k.strip().lower() in cn.lower() or cn.lower() in k.strip().lower():
                        cell=ws.cell(row=ri,column=ci)
                        cell.value='⚠️' if v['formateur']=='⚠️' else f"{v['formateur']}\n{v['matiere']}"
                        old=cell.font; cell.font=Font(name=old.name or 'Calibri',size=old.size or 9,bold=True,color=old.color)
                        cell.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True)
                        filled+=1; break
    wb.save(output_path); return filled

# ── Génération template colorié ───────────────
def detect_structure(ws):
    kw=['BTS','BAC','EC ','CGC','NDRC','GPME','RDC','RH','Master']
    cc={}; colors={}
    for ri in range(1,6):
        for ci in range(1,ws.max_column+1):
            cell=ws.cell(row=ri,column=ci); v=cell.value
            if isinstance(v,str) and any(k in v for k in kw):
                nom=v.strip(); cc[nom]=ci
                if cell.fill and cell.fill.fgColor and cell.fill.fgColor.type=='rgb':
                    c=cell.fill.fgColor.rgb
                    if c not in ('00000000','FFFFFFFF'): colors[nom]=c
    # Chercher couleurs dans données si non trouvées dans headers
    for nom,ci in cc.items():
        if nom in colors: continue
        for ri in range(5,min(25,ws.max_row+1)):
            cell=ws.cell(row=ri,column=ci)
            if cell.fill and cell.fill.fgColor and cell.fill.fgColor.type=='rgb':
                c=cell.fill.fgColor.rgb
                if c not in ('00000000','FFFFFFFF',None): colors[nom]=c; break
    jf_set={'lundi','mardi','mercredi','jeudi','vendredi','lun','mar','mer','jeu','ven'}
    fdr=6; dlc=2
    for ri in range(3,20):
        for ci in range(1,5):
            v=ws.cell(row=ri,column=ci).value
            if isinstance(v,str) and v.strip().lower() in jf_set:
                fdr=ri; dlc=ci; break
        else: continue
        break
    return {'class_cols':cc,'class_colors':colors,'first_data_row':fdr,'day_label_col':dlc}

def generer_template_colorie(template_path, planning_classes, annee_debut, output_path):
    wb_tpl=load_workbook(template_path); ws_tpl=wb_tpl.active
    struct=detect_structure(ws_tpl)
    cc=struct['class_cols']; colors=struct['class_colors']
    fdr=struct['first_data_row']; dlc=struct['day_label_col']

    # Assigner couleurs manquantes
    ci=0
    for nom in cc:
        if nom not in colors:
            colors[nom]=DEFAULT_COLORS[ci%len(DEFAULT_COLORS)]; ci+=1

    # Index des jours par classe
    cours_set={}
    for pc in planning_classes:
        cours_set[pc['nom']]=set(pc['jours'])

    mois_scolaire=[(9,annee_debut),(10,annee_debut),(11,annee_debut),(12,annee_debut),
                   (1,annee_debut+1),(2,annee_debut+1),(3,annee_debut+1),(4,annee_debut+1),
                   (5,annee_debut+1),(6,annee_debut+1),(7,annee_debut+1),(8,annee_debut+1)]

    jf_nom={0:'Lundi',1:'Mardi',2:'Mercredi',3:'Jeudi',4:'Vendredi'}
    wb_out=Workbook(); wb_out.remove(wb_out.active)
    total=0

    for (mois,annee) in mois_scolaire:
        nom_onglet=f"{MOIS_FR[mois]} {annee}"
        wb_tmp=load_workbook(template_path); ws_src=wb_tmp.active
        ws=wb_out.create_sheet(title=nom_onglet)

        # Copier le template
        for row in ws_src.iter_rows():
            for cell in row:
                nc=ws.cell(row=cell.row,column=cell.column)
                nc.value=cell.value
                if cell.has_style:
                    nc.font=copy.copy(cell.font); nc.fill=copy.copy(cell.fill)
                    nc.border=copy.copy(cell.border); nc.alignment=copy.copy(cell.alignment)
                    nc.number_format=cell.number_format
        for cl,cd in ws_src.column_dimensions.items():
            ws.column_dimensions[cl].width=cd.width
        for ri,rd in ws_src.row_dimensions.items():
            ws.row_dimensions[ri].height=rd.height
        for mg in ws_src.merged_cells.ranges:
            ws.merge_cells(str(mg))

        # Mettre à jour le titre
        for ri in range(1,4):
            for ci in range(1,ws.max_column+1):
                v=ws.cell(row=ri,column=ci).value
                if isinstance(v,str) and any(m in v.lower() for m in MOIS_ABBR):
                    if re.search(r'20\d\d',v):
                        ws.cell(row=ri,column=ci).value=f"{MOIS_FR[mois].upper()} {annee}"

        # Jours ouvrés du mois
        nb_j=calendar.monthrange(annee,mois)[1]
        jours_ouvres=[datetime.date(annee,mois,j) for j in range(1,nb_j+1)
                      if datetime.date(annee,mois,j).weekday()<5 and datetime.date(annee,mois,j) not in FERIES_2026_2027]

        # Écrire les jours et colorier
        row_idx=fdr
        for d in jours_ouvres:
            ds=d.strftime('%Y-%m-%d')
            ws.cell(row=row_idx,column=dlc).value=jf_nom[d.weekday()]
            ws.cell(row=row_idx,column=dlc+1).value=d.day

            for classe_nom,col_idx in cc.items():
                a_cours=any(
                    ds in pc['jours']
                    for pc in planning_classes
                    if pc['nom'] and (pc['nom'].strip().lower() in classe_nom.lower() or classe_nom.lower() in pc['nom'].strip().lower())
                )
                cell=ws.cell(row=row_idx,column=col_idx)
                if a_cours:
                    color=colors.get(classe_nom,'FFD9D9D9')
                    if len(color)==6: color='FF'+color
                    cell.fill=PatternFill(start_color=color,end_color=color,fill_type='solid')
                    total+=1
                else:
                    cell.fill=PatternFill(fill_type=None); cell.value=None
            row_idx+=1

    wb_out.save(output_path)
    return total, len(mois_scolaire)

# ── Routes ────────────────────────────────────
@app.route('/')
def index():
    return render_template('index.html')

@app.route('/generer', methods=['POST'])
def generer():
    sid=str(uuid.uuid4())[:8]; wd=os.path.join(UPLOAD_FOLDER,sid); os.makedirs(wd)
    def save(f,sub=''):
        d=os.path.join(wd,sub) if sub else wd; os.makedirs(d,exist_ok=True)
        p=os.path.join(d,f.filename); f.save(p); return p
    try:
        cf=request.files.getlist('classes'); df=request.files.getlist('dispos')
        ff=request.files.get('formateurs'); tf=request.files.get('template')
        mois=set(json.loads(request.form.get('mois','[]')))
        if not all([cf,df,ff,tf,mois]): return jsonify({'error':'Fichiers manquants'}),400
        cp=[save(f,'classes') for f in cf if f.filename]
        dp=[save(f,'dispos') for f in df if f.filename]
        fp=save(ff); tp=save(tf)
        pcs=[c for c in [parse_planning_classe(p) for p in cp] if c]
        dispos=[parse_disponibilite(p) for p in dp]
        aff=parse_tableau_formateurs(fp)
        assignment,stats,heures=assigner(pcs,dispos,aff)
        on=f"Planning_{'_'.join(m.replace('/','') for m in sorted(mois))}.xlsx"
        op=os.path.join(wd,on)
        ecrire_planning(tp,assignment,mois,op)
        return jsonify({'sessions_assignees':stats['assigned'],'creneaux_sans_prof':stats['warn'],
            'formateurs_actifs':len(heures),'classes':len(pcs),
            'mois':', '.join(f"{MOIS_FR[int(m.split('/')[0])]} {m.split('/')[1]}" for m in mois),
            'fichier':on,'session_id':sid,'heures_formateurs':{p:dict(m) for p,m in heures.items()}})
    except Exception as e:
        import traceback; return jsonify({'error':str(e),'trace':traceback.format_exc()}),500

@app.route('/generer-template', methods=['POST'])
def generer_template():
    sid=str(uuid.uuid4())[:8]; wd=os.path.join(UPLOAD_FOLDER,sid); os.makedirs(wd)
    def save(f,sub=''):
        d=os.path.join(wd,sub) if sub else wd; os.makedirs(d,exist_ok=True)
        p=os.path.join(d,f.filename); f.save(p); return p
    try:
        cf=request.files.getlist('classes'); tf=request.files.get('template')
        annee=int(request.form.get('annee',2026))
        if not cf or not tf: return jsonify({'error':'Fichiers manquants (classes + template)'}),400
        cp=[save(f,'classes') for f in cf if f.filename]; tp=save(tf)
        pcs=[c for c in [parse_planning_classe(p) for p in cp] if c]
        on=f"Template_Affichage_{annee}_{annee+1}.xlsx"; op=os.path.join(wd,on)
        total,nb=generer_template_colorie(tp,pcs,annee,op)
        return jsonify({'fichier':on,'session_id':sid,'nb_mois':nb,'cases_colories':total,
            'classes':len(pcs),'noms_classes':[c['nom'] for c in pcs if c.get('nom')]})
    except Exception as e:
        import traceback; return jsonify({'error':str(e),'trace':traceback.format_exc()}),500

@app.route('/telecharger/<session_id>/<filename>')
def telecharger(session_id,filename):
    path=os.path.join(UPLOAD_FOLDER,session_id,filename)
    if not os.path.exists(path): return "Fichier introuvable",404
    @after_this_request
    def cleanup(r):
        try: shutil.rmtree(os.path.join(UPLOAD_FOLDER,session_id))
        except: pass
        return r
    return send_file(path,as_attachment=True,download_name=filename)

if __name__=='__main__':
    app.run(debug=True,host='0.0.0.0',port=5000)
