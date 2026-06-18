#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
import_tb_questionnaires.py (v6.1) — Export CSV corrigé
Usage: python import_tb_questionnaires.py <dossier_pdf>
Dépendance: pip install PyPDF2
"""
import os, sys, re, csv
from datetime import datetime

try:
    import PyPDF2
except ImportError:
    sys.exit("ERREUR : pip install PyPDF2")

CAS_COLS = ["ID_Cas","Date_Creation","Date_MAJ","Statut","Date_Episode","Nom","Prenom","DDN","Sexe","Site_Infectieux","VIH","Contagieux","Tx_Date_Debut","Tx_TOD","Tx_Duree_Prevue","Tx_Date_Fin","Medecin","Prochain_Suivi_Medical","Contacts_Nombre","Contacts","Commentaires","Prochain_Suivi_DSPUB","Responsable_Enquete","Adresse_Usager","Telephone_Usager","Cellulaire_Usager"]
CONT_COLS = ["ID_Contact","ID_Cas_Index","Date_Creation","Date_MAJ","Statut","Nom","Prenom","DDN","Sexe","Cas_Index_Nom_Prenom","Cas_Index_Region","Cas_Index_Test_VIH","Date_Episode_Cas","Lien_Avec_Cas","Type_Contact","TCT1_Fait","TCT1_Date","TCT1_Resultat","TCT2_Fait","TCT2_Date","TCT2_Resultat","Medecin","RX_Poumons_Fait","RX_Poumons_Date","RX_Poumons_Resultat","Traitement","Tx_Date_Debut","Tx_Duree_Prevue","Tx_Date_Fin","Commentaires","Prochain_Suivi","Responsable"]

EMPTY_VALUES = {
    "", "N/A", "n/a",
    "Choisissez s'il vous pla\u00eet",
    "Choisissez s\u2019il vous pla\u00eet",
    "Choisissez s'il vous plait",
    "/Off", "Off", None
}

def is_empty(val):
    if not val: return True
    return val.strip() in EMPTY_VALUES or "Choisissez" in val

# ── Extraction PDF ──────────────────────────────────────────
def extraire_champs_pdf(path):
    champs = {}
    with open(path, "rb") as f:
        reader = PyPDF2.PdfReader(f)
        fields = reader.get_fields()
        if not fields: return champs
        for key, field in fields.items():
            val = field.get("/V", "")
            if isinstance(val, bytes):
                try: val = val.decode("utf-8")
                except: val = str(val)
            if val is None: val = ""
            champs[key] = str(val).strip()
    return champs

def v(ch, *frags):
    for key, val in ch.items():
        if all(f in key for f in frags):
            if is_empty(val): return ""
            return val
    return ""

def v_check(ch, *frags):
    raw = v(ch, *frags)
    return "checked" in raw.lower() if raw else False

# ── Nettoyage ───────────────────────────────────────────────
def nc(raw):
    """Enlève le préfixe numérique: '104 - HAÏTI' → 'HAÏTI'"""
    if not raw: return ""
    m = re.match(r'^\d+\s*-\s*(.+)$', raw.strip())
    return m.group(1).strip() if m else raw.strip()

def sexe_nam(nam):
    if not nam: return "Inconnu"
    c = nam.replace(" ", "")
    if len(c) < 8: return "Inconnu"
    d = c[4:]
    if len(d) < 4 or not d[:4].isdigit(): return "Inconnu"
    mois = int(d[2:4])
    if mois > 50: return "F"
    elif 1 <= mois <= 12: return "M"
    return "Inconnu"

# ── Mappings avec accents corrects ──────────────────────────
def m_statut(raw):
    if not raw: return ""
    return nc(raw)

def m_sexe(raw, nam=""):
    if raw:
        r = raw.lower().strip()
        if "masculin" in r or r == "m": return "M"
        if "féminin" in r or "feminin" in r or r == "f": return "F"
        if "autre" in r: return "Autre"
    if nam: return sexe_nam(nam)
    return "Inconnu"

def m_vih(raw):
    if not raw: return "Non testé"
    c = nc(raw).lower()
    if c == "oui" or "positif" in c: return "Positif"
    if c == "non" or "négatif" in c or "negatif" in c: return "Négatif"
    if "inconnu" in c: return "Inconnu"
    if "non fait" in c or "non testé" in c: return "Non testé"
    return "Non testé"

def m_site(raw):
    if not raw: return ""
    r = nc(raw).lower()
    p = any(w in r for w in ["pulmonaire","poumon","respiratoire",
        "siège de l'infection","siege de l'infection","symptômes compatibles",
        "surveillance médicale","surveillance medicale"])
    e = any(w in r for w in ["extrapulmonaire","ganglionnaire","méningé",
        "osseu","génito","pleural","péritonéal","miliaire"])
    if p and e: return "Pulmonaire et extrapulmonaire"
    if e: return "Extrapulmonaire"
    if p: return "Pulmonaire"
    return "Pulmonaire"

def m_contagieux(raw_rx):
    """Basé sur le résultat RX (champ 4208)"""
    if not raw_rx: return "Inconnu"
    c = nc(raw_rx).lower()
    if "cavitaire" in c and "non cavitaire" not in c: return "Oui"
    if "non cavitaire" in c: return "Inconnu"
    if "normale" in c: return "Non"
    if "non fait" in c: return "Donnée inconnue"
    return "Inconnu"

def m_tod(ch):
    # Priorité TOD (Oui) sur autoadministré (Non)
    for s in ["3.3858","3.6492","3.3859","3.3860","3.3861"]:
        if v_check(ch, f"COMPOSANT.QUESTION.{s}"): return "Oui"
    if v_check(ch, "COMPOSANT.QUESTION.3.3862"): return "Non"
    return ""

# ── Résistances (checkboxes 6931-6950) ──────────────────────
def extraire_resistances(ch):
    res_map = {
        "6931":"INH","6932":"RMP","6933":"SM","6934":"EMB",
        "6935":"PAS","6936":"PZA","6937":"ETHI","6938":"CAP",
        "6939":"KAN","6940":"CIP","6941":"AK","6942":"(modifié)",
        "6943":"RFB","6944":"OFL","6945":"MFX","6946":"LEV",
        "6947":"CS","6948":"CF","6949":"LZD","6950":"Autre"
    }
    resistances = []
    for code, nom in res_map.items():
        if v_check(ch, f"COMPOSANT.QUESTION.3.{code}"):
            resistances.append(nom)
    return resistances


# ── Extraction coordonnées usager ──────────────────────────
def extraire_adresse_usager(ch):
    """Compose l'adresse depuis adresse_27, ville_30, code_33."""
    adresse = v(ch, "identification.adresse_27")
    ville   = v(ch, "identification.ville_30")
    cp      = v(ch, "identification.adresse.m34.codepostal.code_33")
    parties = [p for p in (adresse, ville, cp) if p]
    return ", ".join(parties)

def extraire_telephone_usager(ch):
    """Extrait le téléphone depuis identification.coordonnee.telephones.0_39."""
    return v(ch, "identification.coordonnee.telephones.0_39")

def extraire_cellulaire_usager(ch):
    """Extrait le cellulaire depuis tout champ contenant 'cellulaire'."""
    for key, val in ch.items():
        if "cellulaire" in key.lower():
            if not is_empty(val):
                return val.strip()
    return ""

# ── Commentaires enrichis ───────────────────────────────────
def commentaires(ch):
    p = []
    
    # Identifiants
    sigmi = v(ch, "dsp.n.sigmi"); nam = v(ch, "identification.nam")
    if sigmi or nam:
        p.append(f"[SI-GMI: {sigmi} | NAM: {nam}]")
    
    # Pays/immigration
    pays = nc(v(ch, "QUESTION.3.2482"))
    arrivee = v(ch, "QUESTION.3.37261")
    arrivee_date = v(ch, "QUESTION.3.2596")
    immig = nc(v(ch, "QUESTION.3.2477"))
    if pays:
        l = f"Pays: {pays}"
        if arrivee_date: l += f" | Arrivée: {arrivee_date}"
        elif arrivee: l += f" | Arrivée: {arrivee}"
        if immig: l += f" | {immig}"
        p.append(l)
    
    # Emploi - filtrer les lignes vides
    occ = nc(v(ch, "QUESTION.3.24293"))
    emp = v(ch, "QUESTION.3.24295")
    fn = v(ch, "QUESTION.3.24297")
    if fn or emp:  # seulement si fonction ou employeur non vide
        l = "Emploi:"
        if fn: l += f" {fn}"
        if occ: l += f" ({occ})"
        if emp: l += f" — {emp}"
        p.append(l)
    elif occ and occ.lower() not in ["travailleur","étudiant","retraité",
         "à la maison ou sans emploi (précisez)","enfant en service de garde"]:
        p.append(f"Emploi: {occ}")
    
    # RX pulmonaire (champ 4208 = résultat RÉEL de la RX)
    rx_result = nc(v(ch, "QUESTION.3.4208"))
    rx_detail = v(ch, "QUESTION.3.4209")
    if rx_result:
        l = f"RX poumons: {rx_result}"
        if rx_detail: l += f" — {rx_detail[:250]}"
        p.append(l)
    
    # TDM
    tdm = nc(v(ch, "QUESTION.3.27262"))
    if tdm: p.append(f"TDM: {tdm}")
    
    # RX antérieure anormale (champ 5147 + détail 5148)
    rx_ant = nc(v(ch, "QUESTION.3.5147"))
    rx_ant_detail = v(ch, "QUESTION.3.5148")
    if rx_ant and "oui" in rx_ant.lower():
        l = "RX antérieure anormale: Oui"
        if rx_ant_detail: l += f" — {rx_ant_detail[:250]}"
        p.append(l)
    
    # VIH
    vih = nc(v(ch, "QUESTION.3.4979"))
    if vih and vih.lower() not in ["non"]:
        p.append(f"VIH: {vih}")
    
    # Diabète
    diab = nc(v(ch, "QUESTION.3.4820"))
    diabd = v(ch, "QUESTION.3.4821")
    if diab and "oui" in diab.lower():
        p.append(f"Diabète: Oui ({diabd})" if diabd else "Diabète: Oui")
    
    # Cancer
    cancer = nc(v(ch, "QUESTION.3.4981"))
    cancerd = v(ch, "QUESTION.3.4982")
    if cancer and "oui" in cancer.lower():
        p.append(f"Cancer: Oui ({cancerd})" if cancerd else "Cancer: Oui")
    
    # Insuffisance rénale
    renal = nc(v(ch, "QUESTION.3.4983"))
    renald = v(ch, "QUESTION.3.5146")
    if renal and "oui" in renal.lower():
        p.append(f"Insuffisance rénale: Oui ({renald})" if renald else "Insuffisance rénale: Oui")
    
    # BCG
    bcg = nc(v(ch, "QUESTION.3.8133"))
    bcgd = v(ch, "QUESTION.3.37263")
    if bcg and "oui" in bcg.lower():
        p.append(f"BCG: Oui ({bcgd})" if bcgd else "BCG: Oui")
    
    # Résistances aux antibiotiques
    resistances = extraire_resistances(ch)
    if resistances:
        p.append(f"⚠️ RÉSISTANCE: {', '.join(resistances)}")
    
    # Médication
    meds = []
    for code, nom, dc in [("8602","INH","8603"),("8844","RMP","8845"),
        ("8848","EMB","8849"),("9098","PZA","9099"),("9376","B6/Pyridoxine",None)]:
        mv = nc(v(ch, f"QUESTION.3.{code}"))
        if mv and "oui" in mv.lower():
            dose = v(ch, f"QUESTION.3.{dc}") if dc else ""
            meds.append(f"{nom} {dose}".strip() if dose else nom)
    if meds: p.append(f"Médication: {' + '.join(meds)}")
    
    # Observance
    obs = nc(v(ch, "QUESTION.3.28736"))
    if obs: p.append(f"Observance: {obs}")
    
    # Pharmacie
    ph = v(ch, "QUESTION.3.7250"); pht = v(ch, "QUESTION.3.7251")
    if ph: p.append(f"Pharmacie: {ph}" + (f" ({pht})" if pht else ""))
    
    # Médecin 2
    md2 = v(ch, "QUESTION.3.563") or v(ch, "QUESTION.3.502")
    md2s = v(ch, "QUESTION.3.564") or v(ch, "QUESTION.3.503")
    if md2: p.append(f"Médecin 2: {md2}" + (f" ({md2s})" if md2s else ""))
    
    # Vols pendant contagiosité
    vols = []
    d1=v(ch,"QUESTION.3.26238")
    if d1:
        dep1=v(ch,"QUESTION.3.26241"); arr1=v(ch,"QUESTION.3.26242")
        n1=v(ch,"QUESTION.3.26239"); c1=v(ch,"QUESTION.3.26240")
        vols.append(f"{dep1}→{arr1} {d1} ({c1} {n1})".strip())
    d2=v(ch,"QUESTION.3.26244")
    if d2:
        dep2=v(ch,"QUESTION.3.26247"); arr2=v(ch,"QUESTION.3.26248")
        n2=v(ch,"QUESTION.3.26246")
        vols.append(f"{dep2}→{arr2} {d2} ({n2})".strip())
    d3=v(ch,"QUESTION.3.26250")
    if d3:
        dep3=v(ch,"QUESTION.3.26253"); arr3=v(ch,"QUESTION.3.26254")
        n3=v(ch,"QUESTION.3.26251"); c3=v(ch,"QUESTION.3.26252")
        vols.append(f"{dep3}→{arr3} {d3} ({c3} {n3})".strip())
    if vols: p.append("Vols contagiosité: " + " | ".join(vols))
    
    # Contagiosité
    dc1=v(ch,"QUESTION.3.12151"); dc2=v(ch,"QUESTION.3.12152")
    if dc1 or dc2: p.append(f"Contagiosité: {dc1} → {dc2}")
    
    # Dates
    dd=v(ch,"QUESTION.3.21119")
    if dd: p.append(f"Diagnostic: {dd}")
    dv=v(ch,"QUESTION.3.1480")
    if dv: p.append(f"Validation CIM-10: {dv}")
    
    return "\n".join(p)

# ── Mapping PDF → CAS ───────────────────────────────────────
def mapper_cas(ch):
    now = datetime.now().strftime("%Y-%m-%d")
    nam = v(ch, "identification.nam")
    return {
        "ID_Cas": v(ch, "dsp.n.sigmi"),
        "Date_Creation": now, "Date_MAJ": now,
        "Statut": m_statut(v(ch, "QUESTION.3.15755")),
        "Date_Episode": v(ch, "QUESTION.3.101") or v(ch, "QUESTION.3.21119") or v(ch, "QUESTION.3.8366") or v(ch, "QUESTION.3.12151"),
        "Nom": v(ch, "identification.nom"),
        "Prenom": v(ch, "identification.prenom"),
        "DDN": v(ch, "identification.ddn"),
        "Sexe": m_sexe(v(ch, "identification.sexe"), nam),
        "Site_Infectieux": m_site(v(ch, "QUESTION.3.27573")),
        "VIH": m_vih(v(ch, "QUESTION.3.4979")),
        "Contagieux": m_contagieux(v(ch, "QUESTION.3.4208")),
        "Tx_Date_Debut": v(ch, "QUESTION.3.8366"),
        "Tx_TOD": m_tod(ch),
        "Tx_Duree_Prevue": "",
        "Tx_Date_Fin": v(ch, "QUESTION.3.8367"),
        "Medecin": v(ch, "QUESTION.3.439"),
        "Prochain_Suivi_Medical": "",
        "Contacts_Nombre": 0, "Contacts": "",
        "Commentaires": commentaires(ch),
        "Prochain_Suivi_DSPUB": "",
        "Responsable_Enquete": v(ch, "enqueteur.e.responsable"),
        "Adresse_Usager": extraire_adresse_usager(ch),
        "Telephone_Usager": extraire_telephone_usager(ch),
        "Cellulaire_Usager": extraire_cellulaire_usager(ch),
    }

# ── Mapping PDF → CONTACTS ──────────────────────────────────
# Note: les contacts réels sont dans le module "Gestion des contacts"
# de SI-GMI, PAS dans le questionnaire PDF. Le questionnaire ne contient
# que les "contacts avec un cas connu" (section 10) — qui est le cas INDEX
# qui a contaminé ce patient (pas ses propres contacts).
# On n'extrait donc PAS de contacts du questionnaire.
def mapper_contacts(ch, cas):
    return []

# ── CSV ─────────────────────────────────────────────────────
def ecrire_csv(path, cols, rows):
    with open(path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f, delimiter=";", quoting=csv.QUOTE_MINIMAL)
        w.writerow(cols)
        for row in rows:
            line = []
            for c in cols:
                val = str(row.get(c, ""))
                val = val.replace("\r\n", "\\n").replace("\n", "\\n").replace("\r", "")
                line.append(val)
            w.writerow(line)

# ── Main ────────────────────────────────────────────────────
def main():
    if len(sys.argv) < 2:
        print("Usage: python import_tb_questionnaires.py <dossier_pdf>")
        sys.exit(1)
    dossier = sys.argv[1]
    if not os.path.isdir(dossier):
        sys.exit(f"ERREUR : '{dossier}' introuvable")
    pdfs = sorted([os.path.join(dossier, f) for f in os.listdir(dossier)
                   if f.lower().endswith(".pdf")])
    if not pdfs:
        sys.exit("ERREUR : aucun PDF")
    
    tous_cas, tous_cont, res = [], [], []
    for pdf in pdfs:
        nom = os.path.basename(pdf)
        try:
            ch = extraire_champs_pdf(pdf)
            if not ch:
                res.append(f"  ❌ {nom} — Aucun champ AcroForm")
                continue
            cas = mapper_cas(ch)
            if not cas["ID_Cas"]:
                res.append(f"  ❌ {nom} — SI-GMI manquant")
                continue
            if not cas["Nom"]:
                res.append(f"  ❌ {nom} — Nom manquant")
                continue
            
            cont = mapper_contacts(ch, cas)
            cas["Contacts_Nombre"] = len(cont)
            cas["Contacts"] = ";".join(c["ID_Contact"] for c in cont)
            
            tous_cas.append(cas)
            tous_cont.extend(cont)
            res.append(f"  ✅ {nom} — {cas['ID_Cas']} ({cas['Nom']} {cas['Prenom']})")
        except Exception as e:
            res.append(f"  ❌ {nom} — {e}")
    
    csv_cas = os.path.join(dossier, "_import_cas.csv")
    csv_cont = os.path.join(dossier, "_import_contacts.csv")
    ecrire_csv(csv_cas, CAS_COLS, tous_cas)
    ecrire_csv(csv_cont, CONT_COLS, tous_cont)
    
    print("=" * 60)
    print(f"  RAPPORT v6.1 — {datetime.now().strftime('%Y-%m-%d %H:%M')}")
    print(f"  Cas: {len(tous_cas)} | Contacts: {len(tous_cont)}")
    print("=" * 60)
    for r in res: print(r)
    print(f"\n  📄 {csv_cas}")
    print(f"  📄 {csv_cont}")
    print("\n  → Alt+F8 → ImporterCSV dans Excel")

if __name__ == "__main__":
    main()
