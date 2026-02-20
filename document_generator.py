# document_generator.py
from datetime import datetime
from docx import Document
from docx.shared import RGBColor
from docx.oxml.ns import qn
from docx.oxml import parse_xml
import os
import sys

# Fix per docxcompose su Python 3.13
try:
    from docxcompose.composer import Composer
except ImportError as e:
    print(f"Errore importazione Composer: {e}")
    # Fallback: installazione runtime se necessario
    import subprocess
    subprocess.check_call([sys.executable, "-m", "pip", "install", "docxcompose==1.1.2"])
    from docxcompose.composer import Composer

# Database agenti chimici (tuo codice originale)
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
    "Inchiostri per offset": ["2-Medio", "H315, H318, H412"],
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

def imposta_colore_cella(cella, colore_hex):
    """Imposta colore di sfondo cella"""
    xml_string = f'<w:shd xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:fill="{colore_hex}"/>'
    shading_elm = parse_xml(xml_string)
    cella._tc.get_or_add_tcPr().append(shading_elm)

def sostituisci_mantieni_formato(paragrafo, placeholder, valore):
    """
    Sostituisce il placeholder mantenendo la formattazione del primo run
    """
    if placeholder not in paragrafo.text:
        return False
    
    # Trova il run con il placeholder
    for run in paragrafo.runs:
        if placeholder in run.text:
            # Salva la formattazione
            font_name = run.font.name
            font_size = run.font.size
            bold = run.font.bold
            italic = run.font.italic
            color = run.font.color.rgb if run.font.color else None
            
            # Sostituisci testo
            run.text = run.text.replace(placeholder, str(valore))
            
            # Ripristina formattazione
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
    
    # Se il placeholder è spezzato su più run, ricostruisci
    full_text = ''.join([r.text for r in paragrafo.runs])
    if placeholder in full_text:
        # Cancella tutti i run tranne il primo
        first_run = paragrafo.runs[0]
        for run in paragrafo.runs[1:]:
            run.text = ""
        
        # Sostituisci nel primo run mantenendo la sua formattazione
        first_run.text = full_text.replace(placeholder, str(valore))
        return True
    
    return False

def compila_segnaposto(doc, dati):
    """Sostituisce segnaposto mantenendo la formattazione"""
    # Sostituzione nei paragrafi
    for p in doc.paragraphs:
        for key, value in dati.items():
            placeholder = f"{{{{{key}}}}}"
            if placeholder in p.text:
                sostituisci_mantieni_formato(p, placeholder, str(value))
    
    # Sostituzione nelle tabelle
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

            # Intestazione
            hdr_cells = table.rows[0].cells
            for i, t in enumerate(['Prodotto', 'Rischio', 'Classificazione']):
                hdr_cells[i].text = t
                imposta_colore_cella(hdr_cells[i], "D9D9D9")
                hdr_cells[i].paragraphs[0].runs[0].bold = True

            # Righe dal Database
            for nome_prod in lista_scelti:
                if nome_prod in db:
                    row_cells = table.add_row().cells
                    row_cells[0].text = nome_prod
                    rischio = db[nome_prod][0]
                    row_cells[1].text = rischio
                    row_cells[2].text = db[nome_prod][1]

                    # Logica Colori
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

def formatta_per_word(lista):
    """Trasforma nomi file in testo leggibile"""
    if not lista:
        return "Nessuno"
    voci_pulite = [v.replace("_", " ").capitalize() for v in lista]
    return "\n".join([f"- {v}" for v in voci_pulite])

def genera_dvr(azienda_data, ambienti, attrezzature, mansioni, agenti_chimici, templates_dir):
    """
    Funzione principale che genera il documento DVR usando docxcompose
    """
    # Prepara i dati
    data_di_oggi = datetime.now().strftime("%d/%m/%Y")
    azienda_data["DATA"] = data_di_oggi
    
    # Formatta le liste
    azienda_data["LISTA_AMBIENTI"] = formatta_per_word(ambienti)
    azienda_data["LISTA_MANSIONI"] = formatta_per_word(mansioni)
    azienda_data["LISTA_ATTREZZATURE"] = formatta_per_word(attrezzature)
    azienda_data["LISTA_CHIMICI"] = formatta_per_word(agenti_chimici)
    
    # Percorso template
    template_path = os.path.join(templates_dir, 'Template_Base.docx')
    
    if not os.path.exists(template_path):
        raise FileNotFoundError(f"Template non trovato: {template_path}")
    
    # 1. Carica template master
    master_doc = Document(template_path)
    
    # 2. Compila segnaposti anagrafici (mantenendo formattazione)
    compila_segnaposto(master_doc, azienda_data)
    
    # 3. Inserisci tabella chimica
    inserisci_tabella_chimica(master_doc, "{{TABELLA_CHIMICA}}", agenti_chimici, db_chimico)
    
    # 4. Assembla moduli con Composer (metodo sicuro)
    print("Assemblaggio moduli in corso...")
    composer = Composer(master_doc)
    
    # Aggiungi moduli ambienti
    for ambiente in ambienti:
        mod_path = os.path.join(templates_dir, f"{ambiente}.docx")
        if os.path.exists(mod_path):
            try:
                doc = Document(mod_path)
                composer.append(doc)
                print(f"  ✓ Aggiunto ambiente: {ambiente}")
            except Exception as e:
                print(f"  ✗ Errore con {ambiente}: {e}")
    
    # Aggiungi moduli attrezzature  
    for att in attrezzature:
        mod_path = os.path.join(templates_dir, f"{att}.docx")
        if os.path.exists(mod_path):
            try:
                doc = Document(mod_path)
                composer.append(doc)
                print(f"  ✓ Aggiunta attrezzatura: {att}")
            except Exception as e:
                print(f"  ✗ Errore con {att}: {e}")
    
    # Aggiungi moduli mansioni
    for mans in mansioni:
        mod_path = os.path.join(templates_dir, f"{mans}.docx")
        if os.path.exists(mod_path):
            try:
                doc = Document(mod_path)
                composer.append(doc)
                print(f"  ✓ Aggiunta mansione: {mans}")
            except Exception as e:
                print(f"  ✗ Errore con {mans}: {e}")
    
    # 5. Salva in memoria (bytes)
    from io import BytesIO
    buffer = BytesIO()
    composer.save(buffer)
    buffer.seek(0)
    
    return buffer
