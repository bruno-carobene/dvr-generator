# document_generator.py
from datetime import datetime
from docx import Document
from docx.shared import RGBColor, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import parse_xml
import os

# Database agenti chimici
db_chimico = {
    "Acidi per laboratori didattici": ["2-Medio", "H314"], 
    "Acido cloridrico": ["2-Medio", "H315 - H335"],
    "Acido solforico al 30%": ["2-Medio", "GHS05, H209, H314"], 
    "Acquaragia": ["2-Medio", "H304"],
    "Agenti pulitori sgrassanti": ["2-Medio", "H412, H304, H226, H336, H229"], 
    "Alcool etilico": ["1-Basso", "H225, H226, H319"],
    "Alghicida": ["3-Alto", "H314, H318"], 
    "Ammoniaca": ["3-Alto", "H315, H319"],
    "Antiadesivo siliconico": ["1-Basso", "H222, H229"], 
    "Anticorrosivo": ["2-Medio", "H22, H229,H315, H412"],
    "Antigelo": ["1-Basso", "H302, H304, H315, H318, H351"], 
    "Antigelo Permanente": ["1-Basso", "H302"],
    "Antiruggine": ["2-Medio", "H314, H315, H319"], 
    "Antiruggine liquido": ["1-Basso", "H226, H373, H315"],
    "Argon": ["1-Basso", "H280"], 
    "Azoto": ["1-Basso", "H281"],
    "Benzina": ["1-Basso", "H224, H304, H340, H350"], 
    "Blu di prussia": ["2-Medio", "H302, H315, H319"],
    "Candeggina": ["2-Medio", "H315, H319"], 
    "Catalizzatore vernici veicoli": ["2-Medio", "H226, H332, H304, H412"],
    "Collodio": ["1-Basso", "H226, H319, H335"], 
    "Correttore di pH": ["3-Alto", "H302, H314, H318, H400"],
    "Detergente disincrostante forni": ["2-Medio", "H315, H319"], 
    "Detergente igienizzante clima": ["1-Basso", "H225, H319"],
    "Detergente stoviglie a mano": ["3-Alto", "H302, H315, H318"], 
    "Detergente lavastoviglie": ["3-Alto", "H319, H315"],
    "Detergente lucidatura carrozzerie": ["1-Basso", "H225, H319, H336"], 
    "Detergente per pavimenti": ["1-Basso", "H315, H318"],
    "Detergente per superfici diluito": ["1-Basso", "-"], 
    "Detergente per WC": ["1-Basso", "H314, H335"],
    "Detergente speciale offset": ["1-Basso", "H226, H304, H336, H411"], 
    "Detersivo per lavatrice": ["1-Basso", "-"],
    "Diluente per inchiostri": ["2-Medio", "H226, H304, H335, H336"], 
    "Diluenti Nitro Antinebbia": ["3-Alto", "H225, H361d, H373"],
    "Flocculante": ["2-Medio", "H318"], 
    "Flussante": ["2-Medio", "H319, H336, H225"],
    "Fondo verniciatura veicoli": ["2-Medio", "H226, H314, H373, H412"], 
    "Fumi di saldatura": ["2-Medio", "n.a."],
    "Gasolio": ["1-Basso", "H226, H304, H332, H351, H411"], 
    "Glicole etilenico": ["1-Basso", "H302"],
    "Grasso lubrificante": ["1-Basso", "-"], 
    "Inchiostri per offset": ["2-Medio", "H315, H318, "],
    "Indurente vernici veicoli": ["2-Medio", "H226, H332, H317, H360, H412"], 
    "Legante basi verniciatura": ["2-Medio", "-"],
    "Loctite-401": ["1-Basso", "H315, H319, H335"], 
    "Lubrificanti spray (Svitol/Grasso)": ["1-Basso", "H223, H304, H411"],
    "Malta": ["2-Medio", "H318, H315, H317, H335"], 
    "Oli lubrificanti": ["1-Basso", "H315, H318, H336"],
    "Olio per impastare inchiostri": ["1-Basso", "H225, EUH066"], 
    "Pasta per riscontro": ["1-Basso", "-"],
    "Pasta per saldare i chip": ["2-Medio", "H302, H315, H317, H410"], 
    "Pittura ad acqua": ["2-Medio", "-"],
    "Polveri da molatura": ["2-Medio", "-"], 
    "Primer verniciatura veicoli": ["2-Medio", "H225, H315, H373, H412"],
    "Pulitore contatti elettrici": ["2-Medio", "H222, H315, H319, H411"], 
    "Reagenti": ["3-Alto", "Varia"],
    "Rivestimento trasparente veicoli": ["2-Medio", "H226, H317, H336, H412"], 
    "Sbloccante spray": ["1-Basso", "H223, H336, H229"],
    "Sepiolite": ["2-Medio", "H318, H302, H315"], 
    "Silicone spray": ["2-Medio", "H222, H315, H336, H411"],
    "Soluzione disinfettante": ["1-Basso", "H302, H318, H319, H336"], 
    "Solventi": ["3-Alto", "H304, H336"],
    "Tinta per capelli": ["2-Medio", "n.a."], 
    "Toner": ["1-Basso", "-"],
    "Total clean": ["1-Basso", "H319"], 
    "Vernice acqua spruzzo": ["2-Medio", "H319"],
    "Vernice spray": ["2-Medio", "H222, H319, H336, H411"], 
    "Vernici per offset": ["2-Medio", "H301, H314, H317, H411"]
}

# Testo del sommario statico (dal tuo file)
SOMMARIO_STATICO = """SOMMARIO

Introduzione
Obiettivi del documento
Chi ha partecipato alla redazione del documento
Procedura di identificazione e analisi dei rischi e definizione dei controlli
    Identificazione dei centri/fonti di pericolo per la sicurezza e la salute dei lavoratori
    Identificazione dei lavoratori (o di terzi) esposti a rischi potenziali
    Valutazione dei rischi, dal punto di vista qualitativo e quantitativo
    Studio sulla possibilità di eliminare i rischi
    Programma delle misure ritenute opportune per garantire il miglioramento nel tempo dei livelli di sicurezza e procedura per l'attuazione
Elenco dei pericoli considerati
Criteri di quantificazione del rischio
    Probabilità
    Danno
    Rischio
    Quantificazione dei rischi specifici
Prescrizioni legali
Gestione del documento

L'azienda
Anagrafica aziendale
Il Sistema di sicurezza aziendale
Descrizione strutturale della sede di lavoro
    Descrizione generale dei locali
    Attività affidate a terzi
    Attività svolte presso terzi
Attrezzature e agenti chimici impiegati
Elenco delle attrezzature impiegate
Agenti chimici

Allegati
Ambienti di lavoro
Attrezzature
Mansioni"""

def imposta_colore_cella(cella, colore_hex):
    """Imposta colore di sfondo cella"""
    xml_string = f'<w:shd xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:fill="{colore_hex}"/>'
    shading_elm = parse_xml(xml_string)
    cella._tc.get_or_add_tcPr().append(shading_elm)

def sostituisci_mantieni_formato(paragrafo, placeholder, valore):
    """Sostituisce il placeholder mantenendo la formattazione del primo run"""
    if placeholder not in paragrafo.text:
        return False
    
    for run in paragrafo.runs:
        if placeholder in run.text:
            font_name = run.font.name
            font_size = run.font.size
            bold = run.font.bold
            italic = run.font.italic
            color = run.font.color.rgb if run.font.color else None
            
            run.text = run.text.replace(placeholder, str(valore))
            
            if font_name:
                run.font.name = font_name
            if font_size:
                run.font.size = font_size
            if bold:
                run.font.bold = bold
            if italic:
                run.font.italic = italic
            if color:
                run.font.color.rgb = color
            
            return True
    
    full_text = ''.join([r.text for r in paragrafo.runs])
    if placeholder in full_text:
        first_run = paragrafo.runs[0]
        for run in paragrafo.runs[1:]:
            run.text = ""
        first_run.text = full_text.replace(placeholder, str(valore))
        return True
    
    return False

def compila_segnaposto(doc, dati):
    """Sostituisce segnaposto mantenendo la formattazione"""
    for p in doc.paragraphs:
        for key, value in dati.items():
            placeholder = f"{{{{{key}}}}}"
            if placeholder in p.text:
                sostituisci_mantieni_formato(p, placeholder, str(value))
    
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for key, value in dati.items():
                    placeholder = f"{{{{{key}}}}}"
                    if placeholder in cell.text:
                        for p in cell.paragraphs:
                            sostituisci_mantieni_formato(p, placeholder, str(value))

def inserisci_tabella_chimica(doc, segnaposto, lista_scelti, db):
    """Inserisce tabella agenti chimici"""
    for p in doc.paragraphs:
        if segnaposto in p.text:
            p.text = p.text.replace(segnaposto, "")
            table = doc.add_table(rows=1, cols=3)
            table.style = 'Table Grid'

            hdr_cells = table.rows[0].cells
            for i, t in enumerate(['Prodotto', 'Rischio', 'Classificazione']):
                hdr_cells[i].text = t
                imposta_colore_cella(hdr_cells[i], "D9D9D9")
                hdr_cells[i].paragraphs[0].runs[0].bold = True

            for nome_prod in lista_scelti:
                if nome_prod in db:
                    row_cells = table.add_row().cells
                    row_cells[0].text = nome_prod
                    rischio = db[nome_prod][0]
                    row_cells[1].text = rischio
                    row_cells[2].text = db[nome_prod][1]

                    if "Alto" in rischio: 
                        colore = "FF9999"
                    elif "Medio" in rischio: 
                        colore = "FFFFCC"
                    elif "Basso" in rischio: 
                        colore = "CCFFCC"
                    else: 
                        colore = "FFFFFF"

                    imposta_colore_cella(row_cells[1], colore)

            p._element.addnext(table._element)

def rimuovi_sommario_dinamico(doc):
    """Rimuove il campo sommario (TOC) dal documento"""
    # Cerca e rimuove i campi TOC
    for p in doc.paragraphs[:]:  # Copia la lista per poter rimuovere
        # Cerca il campo TOC (ha caratteri speciali)
        if 'TOC' in p.text or 'SOMMARIO' in p.text.upper():
            # Verifica se è un campo dinamico (ha fldChar)
            if p._element.xpath('.//w:fldChar'):
                p.text = ""  # Svuota il paragrafo
                continue
        
        # Rimuove anche "Nessuna voce di sommario trovata"
        if 'Nessuna voce di sommario trovata' in p.text:
            p.text = ""

def aggiungi_sommario_statico(doc):
    """Aggiunge il sommario statico a pagina 2"""
    # Trova l'elemento dopo la prima pagina (dopo un salto pagina o alla fine della prima sezione)
    # Inseriamo dopo il primo paragrafo che troviamo con "Documento di Valutazione" o simile
    
    target_para = None
    for i, p in enumerate(doc.paragraphs):
        if 'DOCUMENTO' in p.text.upper() or 'VALUTAZIONE' in p.text.upper():
            # Cerca il prossimo salto pagina o fine sezione
            target_para = p
            break
    
    if target_para is None:
        # Se non troviamo il target, inseriamo dopo il primo paragrafo
        target_para = doc.paragraphs[0] if doc.paragraphs else None
    
    if target_para:
        # Aggiungi salto pagina
        target_para._element.addnext(parse_xml(r'<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:r><w:br w:type="page"/></w:r></w:p>'))
        
        # Aggiungi titolo SOMMARIO
        p_sommario = doc.add_paragraph()
        p_sommario._element.getprevious().addnext(p_sommario._element)
        run = p_sommario.add_run("SOMMARIO")
        run.bold = True
        run.font.size = Pt(16)
        p_sommario.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        # Aggiungi righe del sommario
        righe = SOMMARIO_STATICO.split('\n')
        for riga in righe[1:]:  # Salta il primo "SOMMARIO"
            if riga.strip():
                p = doc.add_paragraph()
                p_sommario._element.addnext(p._element)
                p_sommario = p  # Aggiorna riferimento per il prossimo
                
                # Gestione indentazione
                if riga.startswith('    ') or riga.startswith('\t'):
                    # Sotto-sezione (indentata)
                    p.paragraph_format.left_indent = Pt(36)
                    run = p.add_run(riga.strip())
                    run.font.size = Pt(10)
                else:
                    # Sezione principale
                    run = p.add_run(riga.strip())
                    run.bold = True
                    run.font.size = Pt(11)

def formatta_elenco_bullettato(lista):
    """Formatta elenco con bullet e a capo, senza trattini"""
    if not lista:
        return ""
    
    # Formatta nomi
    voci_pulite = [v.replace("_", " ").capitalize() for v in lista]
    
    # Unisci con a capo (senza trattini iniziali, solo bullet verranno aggiunti dal template)
    return "\n".join(voci_pulite)

def copia_elementi_sicuro(src_doc, dest_doc):
    """Copia elementi da un documento all'altro in modo sicuro"""
    for element in src_doc.element.body:
        dest_doc.element.body.append(element)

def genera_dvr(azienda_data, ambienti, attrezzature, mansioni, agenti_chimici, templates_dir):
    """
    Funzione principale che genera il documento DVR
    """
    # Prepara i dati
    data_di_oggi = datetime.now().strftime("%d/%m/%Y")
    azienda_data["DATA"] = data_di_oggi
    
    # Formatta le liste CON a capo (per elenchi puntati)
    azienda_data["LISTA_AMBIENTI"] = formatta_elenco_bullettato(ambienti)
    azienda_data["LISTA_MANSIONI"] = formatta_elenco_bullettato(mansioni)
    azienda_data["LISTA_ATTREZZATURE"] = formatta_elenco_bullettato(attrezzature)
    azienda_data["LISTA_CHIMICI"] = formatta_elenco_bullettato(agenti_chimici)
    
    # Percorso template
    template_path = os.path.join(templates_dir, 'Template_Base.docx')
    
    if not os.path.exists(template_path):
        raise FileNotFoundError(f"Template non trovato: {template_path}")
    
    # 1. Carica template master
    master_doc = Document(template_path)
    
    # 2. Rimuovi sommario dinamico esistente
    rimuovi_sommario_dinamico(master_doc)
    
    # 3. Aggiungi sommario statico
    aggiungi_sommario_statico(master_doc)
    
    # 4. Compila segnaposti anagrafici (mantenendo formattazione)
    compila_segnaposto(master_doc, azienda_data)
    
    # 5. Inserisci tabella chimica
    inserisci_tabella_chimica(master_doc, "{{TABELLA_CHIMICA}}", agenti_chimici, db_chimico)
    
    # 6. Assembla moduli (metodo sicuro senza docxcompose)
    print("Assemblaggio moduli in corso...")
    
    # Raccogli tutti i moduli da aggiungere
    moduli_da_aggiungere = []
    
    for ambiente in ambienti:
        mod_path = os.path.join(templates_dir, f"{ambiente}.docx")
        if os.path.exists(mod_path):
            moduli_da_aggiungere.append(("ambiente", ambiente, mod_path))
    
    for att in attrezzature:
        mod_path = os.path.join(templates_dir, f"{att}.docx")
        if os.path.exists(mod_path):
            moduli_da_aggiungere.append(("attrezzatura", att, mod_path))
    
    for mans in mansioni:
        mod_path = os.path.join(templates_dir, f"{mans}.docx")
        if os.path.exists(mod_path):
            moduli_da_aggiungere.append(("mansione", mans, mod_path))
    
    # Aggiungi moduli uno per uno
    for tipo, nome, mod_path in moduli_da_aggiungere:
        try:
            mod_doc = Document(mod_path)
            # Aggiungi salto pagina
            master_doc.add_page_break()
            # Copia paragrafi uno per uno (più sicuro)
            for para in mod_doc.paragraphs:
                new_para = master_doc.add_paragraph()
                new_para.text = para.text
                # Copia stile
                if para.style:
                    try:
                        new_para.style = para.style.name
                    except:
                        pass
                # Copia formattazione
                if para.runs:
                    for run in para.runs:
                        new_run = new_para.add_run(run.text)
                        new_run.bold = run.bold
                        new_run.italic = run.italic
                        new_run.font.size = run.font.size
                        new_run.font.name = run.font.name
            print(f"  ✓ Aggiunto {tipo}: {nome}")
        except Exception as e:
            print(f"  ✗ Errore con {nome}: {e}")
    
    # 7. Salva in memoria (bytes)
    from io import BytesIO
    buffer = BytesIO()
    master_doc.save(buffer)
    buffer.seek(0)
    
    return buffer
