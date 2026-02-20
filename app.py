import streamlit as st
from document_generator import genera_dvr
from datetime import datetime
import os

# === CONFIGURAZIONE PASSWORD ===
PASSWORD_CORRETTA = "easyworkdvr26"  # Cambia con la password che vuoi

# === FUNZIONE LOGIN ===
def check_password():
    """Verifica password e gestisce login"""
    
    # Inizializza stato sessione
    if "password_correct" not in st.session_state:
        st.session_state.password_correct = False
    
    # Se già loggato, mostra app
    if st.session_state.password_correct:
        return True
    
    # Mostra schermata login
    col1, col2, col3 = st.columns([1, 2, 1])
    
    with col2:
        st.markdown("## 🔐 Accesso Riservato")
        st.markdown("---")
        
        # Logo Easywork
        st.image("https://raw.githubusercontent.com/bruno-carobene/dvr-generator/main/assets/logo-easywork.png", 
                 use_column_width=True)
        
        st.markdown("---")
        st.markdown("### Generatore DVR")
        st.markdown("Inserisci la password per accedere all'applicazione")
        
        password = st.text_input("Password", type="password", key="password_input")
        
        if st.button("Accedi", use_container_width=True):
            if password == PASSWORD_CORRETTA:
                st.session_state.password_correct = True
                st.success("✅ Accesso consentito!")
                st.rerun()
            else:
                st.error("❌ Password errata. Riprova.")
        
        # Footer login
        st.markdown("---")
        st.caption("© Easywork - Tutti i diritti riservati")
    
    return False

# === VERIFICA ACCESSO ===
if not check_password():
    st.stop()  # Blocca esecuzione se non loggato

# === INIZIO APP (dopo login) ===
st.set_page_config(
    page_title="Generatore DVR",
    page_icon="📋",
    layout="wide"
)

# Logo in header (più piccolo)
col1, col2, col3 = st.columns([1, 3, 1])
with col2:
    st.image("https://raw.githubusercontent.com/bruno-carobene/dvr-generator/main/assets/logo-easywork.png", 
             width=200)

st.markdown("---")

st.title("Generatore Documento di Valutazione Rischi (DVR). Un progetto Easywork Italia S,r.l.")
st.markdown("Compila il modulo per generare il documento di valutazione rischi personalizzato.")

# Pulsante logout (opzionale, in sidebar)
with st.sidebar:
    if st.button("🚪 Logout"):
        st.session_state.password_correct = False
        st.rerun()
    st.markdown("---")

# Inizializza session state per i dati
if 'azienda_data' not in st.session_state:
    st.session_state.azienda_data = {}

# === SEZIONE 1: ANAGRAFICA AZIENDA ===
st.header("🏢 Anagrafica Aziendale")

col1, col2 = st.columns(2)

with col1:
    nome = st.text_input("Nome dell'azienda *", key="nome")
    tipologia = st.text_input("Tipologia di azienda", key="tipologia")
    ateco = st.text_input("Codice ATECO", key="ateco")
    datore = st.text_input("Datore di lavoro *", key="datore")
    rspp = st.text_input("RSPP", key="rspp")
    rls = st.text_input("RLS", key="rls")
    
with col2:
    piva = st.text_input("Partita IVA", key="piva")
    cf = st.text_input("Codice Fiscale", key="cf")
    medico = st.text_input("Medico competente", key="medico")
    antincendio = st.text_input("Incaricato antincendio", key="antincendio")
    primo_soccorso = st.text_input("Incaricato primo soccorso", key="primo_soccorso")
    orario = st.text_input("Orario di lavoro", key="orario")

# Indirizzi
st.subheader("📍 Sedi")
col3, col4 = st.columns(2)
with col3:
    st.markdown("**Sede Legale**")
    ind_legale = st.text_input("Via e numero civico", key="ind_legale")
    citta_legale = st.text_input("Comune", key="citta_legale")
    prov_legale = st.text_input("Provincia", key="prov_legale")

with col4:
    st.markdown("**Sede Operativa**")
    ind_op = st.text_input("Via e numero civico ", key="ind_op")
    citta_op = st.text_input("Comune ", key="citta_op")
    prov_op = st.text_input("Provincia ", key="prov_op")

# Descrizione attività
st.subheader("📝 Descrizione Attività")
attivita_desc = st.text_area("Descrivi brevemente il settore di attività", key="attivita")
locali = st.text_area("Struttura dei locali", key="locali")
terzi = st.text_input("Attività affidate a terzi", key="terzi")
terzi_svolte = st.text_input("Attività svolte presso terzi", key="terzi_svolte")

# === SEZIONE 2: AMBIENTI ===
st.header("🏭 Ambienti Aziendali")

st.markdown("Seleziona gli ambienti presenti:")

col_amb1, col_amb2, col_amb3 = st.columns(3)

with col_amb1:
    ufficio = st.checkbox("Ufficio")
    magazzino = st.checkbox("Magazzino")
    palestre = st.checkbox("Palestre e locali annessi")
    sala_attivita = st.checkbox("Sala attività associazione")

with col_amb2:
    dehors = st.checkbox("Dehors esterno - tensostruttura")
    area_bar = st.checkbox("Area bar e ristoro")
    locale_quadri = st.checkbox("Locale quadri elettrici")
    locali_spogliatoio = st.checkbox("Locali spogliatoio")

with col_amb3:
    tribuna = st.checkbox("Tribuna")
    locali_caldaia = st.checkbox("Locali caldaia")
    altro_amb = st.text_input("Altro (specificare)")

# === SEZIONE 3: ATTREZZATURE ===
st.header("🔧 Attrezzature Impiegate")

st.markdown("**Cucina e Ristorazione**")
col_att1, col_att2, col_att3 = st.columns(3)
with col_att1:
    att_cucina = st.checkbox("Attrezzature cucina")
    att_cucina_man = st.checkbox("Attrezzi manuali cucina")
    forno_micro = st.checkbox("Forno microonde")
with col_att2:
    forno_elettr = st.checkbox("Forno elettrico")
    frigorifero = st.checkbox("Frigorifero")
    macchina_caffe = st.checkbox("Macchina caffè bar")
with col_att3:
    macchina_ghiaccio = st.checkbox("Macchina ghiaccio")
    piastra_toast = st.checkbox("Piastra toast/panini")

st.markdown("**Ufficio e Tecnologia**")
col_att4, col_att5, col_att6 = st.columns(3)
with col_att4:
    registratore = st.checkbox("Registratore cassa")
    videoterm = st.checkbox("Videoterminali")
    att_ufficio = st.checkbox("Attrezzi manuali ufficio")
with col_att5:
    stampante_laser = st.checkbox("Stampante laser")
    stampante_ink = st.checkbox("Stampante getto inchiostro")
    telefono = st.checkbox("Telefono/Cellulare")
with col_att6:
    pc = st.checkbox("PC fisso/portatile")
    tablet = st.checkbox("Tablet")

st.markdown("**Logistica, Sport e Pulizia**")
col_att7, col_att8 = st.columns(2)
with col_att7:
    scaffali = st.checkbox("Scaffali metallici")
    att_sport = st.checkbox("Attrezzi sportivi")
    att_pulizia = st.checkbox("Attrezzi pulizia")
with col_att8:
    autoveicoli = st.checkbox("Autoveicoli")
    motoveicoli = st.checkbox("Motoveicoli")
    scala = st.checkbox("Scala portatile doppia")

altro_att = st.text_input("Altre attrezzature (specificare)")

# === SEZIONE 4: MANSIONI ===
st.header("👷 Mansioni Presenti")

st.markdown("**Ristorazione e Accoglienza**")
col_man1, col_man2 = st.columns(2)
with col_man1:
    addetto_bar = st.checkbox("Addetto bar")
    aiuto_cuoco = st.checkbox("Addetto cucina/aiutocuoco")
    resp_sala = st.checkbox("Responsabile sala")
with col_man2:
    reception = st.checkbox("Addetto reception")
    customer = st.checkbox("Addetto customer service")

st.markdown("**Amministrazione e Dirigenza**")
col_man3, col_man4 = st.columns(2)
with col_man3:
    dirigente = st.checkbox("Dirigente")
    imp_admin = st.checkbox("Impiegato amministrativo")
with col_man4:
    impiegato = st.checkbox("Impiegato")
    segretario = st.checkbox("Segretario/a")

st.markdown("**Sport e Benessere**")
col_man5, col_man6 = st.columns(2)
with col_man5:
    istruttore = st.checkbox("Istruttore/allenatore/preparatore")
    bagnino = st.checkbox("Bagnino")
with col_man6:
    estetista = st.checkbox("Estetista")
    parrucchiere = st.checkbox("Parrucchiere")

st.markdown("**Servizi e Manutenzione**")
col_man7, col_man8 = st.columns(2)
with col_man7:
    add_servizi = st.checkbox("Addetto ai servizi")
    add_pulizia = st.checkbox("Addetto pulizia base")
    minuta_man = st.checkbox("Addetto minuta manutenzione")
with col_man8:
    manutenzione = st.checkbox("Addetto manutenzione/tuttofare")
    magazzino = st.checkbox("Addetto magazzino")
    operaio = st.checkbox("Operaio generico")

altra_mansione = st.text_input("Altra mansione (specificare)")

# === SEZIONE 5: AGENTI CHIMICI ===
st.header("🧪 Agenti Chimici")

st.markdown("**Pulizia, Igiene e Ristorazione**")
col_ch1, col_ch2, col_ch3 = st.columns(3)
with col_ch1:
    acidi = st.checkbox("Acidi per laboratori didattici")
    alcool = st.checkbox("Alcool etilico")
    ammoniaca = st.checkbox("Ammoniaca")
with col_ch2:
    candeggina = st.checkbox("Candeggina")
    det_forni = st.checkbox("Detergente disincrostante forni")
    det_stov = st.checkbox("Detergente stoviglie")
with col_ch3:
    det_lavast = st.checkbox("Detergente lavastoviglie")
    det_pav = st.checkbox("Detergente pavimenti")
    det_wc = st.checkbox("Detergente WC")

st.markdown("**Officina, Manutenzione e Gas**")
col_ch4, col_ch5 = st.columns(2)
with col_ch4:
    anticorrosivo = st.checkbox("Anticorrosivo")
    antiruggine = st.checkbox("Antiruggine")
    argon = st.checkbox("Argon")
    azoto = st.checkbox("Azoto")
with col_ch5:
    fumi_sald = st.checkbox("Fumi di saldatura")
    grasso_lub = st.checkbox("Grasso lubrificante")
    lubr_spray = st.checkbox("Lubrificanti spray")
    polveri = st.checkbox("Polveri da molatura")

st.markdown("**Verniciatura e Carrozzeria**")
col_ch6, col_ch7 = st.columns(2)
with col_ch6:
    acquaragia = st.checkbox("Acquaragia")
    catalizzatore = st.checkbox("Catalizzatore vernici")
    det_carrozz = st.checkbox("Detergente carrozzerie")
with col_ch7:
    fondo_vern = st.checkbox("Fondo verniciatura")
    primer = st.checkbox("Primer verniciatura")
    vernice_spray = st.checkbox("Vernice spray")

st.markdown("**Carburanti e Altro**")
col_ch8, col_ch9 = st.columns(2)
with col_ch8:
    benzina = st.checkbox("Benzina")
    gasolio = st.checkbox("Gasolio")
with col_ch9:
    toner = st.checkbox("Toner")
    tinta = st.checkbox("Tinta per capelli")

# === PULSANTE GENERAZIONE ===
st.divider()
st.header("🚀 Generazione Documento")

if st.button("Genera DVR", type="primary", use_container_width=True):
    
    # Verifica campi obbligatori
    if not nome or not datore:
        st.error("❌ Compila i campi obbligatori (Nome azienda e Datore di lavoro)")
    else:
        with st.spinner("Generazione documento in corso..."):
            
            # Prepara dizionario azienda
            azienda_data = {
                "nome": nome,
                "tipologia": tipologia,
                "ateco": ateco,
                "Datore_di_lavoro": datore,
                "RSPP": rspp,
                "RLS": rls,
                "indirizzo_legale": ind_legale,
                "citta_legale": citta_legale,
                "provincia_legale": prov_legale,
                "indirizzo_operativo": ind_op,
                "citta_operativa": citta_op,
                "provincia_operativa": prov_op,
                "Incaricato_antincendio": antincendio,
                "Indirizzo_sede": ind_op or ind_legale,
                "Attivita": attivita_desc,
                "Partita_Iva": piva,
                "Codice_fiscale": cf,
                "Orario": orario,
                "Medico": medico,
                "Incaricato_primo_soccorso": primo_soccorso,
                "locali": locali,
                "terzi": terzi,
                "terzi_svolte": terzi_svolte
            }
            
            # Prepara liste
            ambienti = []
            if ufficio: ambienti.append("ufficio")
            if magazzino: ambienti.append("magazzino")
            if palestre: ambienti.append("palestre_locali_annessi")
            if sala_attivita: ambienti.append("sala_attivita_associazione")
            if dehors: ambienti.append("dehors_esterno")
            if area_bar: ambienti.append("area_bar_ristoro")
            if locale_quadri: ambienti.append("locale_quadri_elettrici")
            if locali_spogliatoio: ambienti.append("locali_spogliatoio")
            if tribuna: ambienti.append("tribuna")
            if locali_caldaia: ambienti.append("locali_caldaia")
            if altro_amb: ambienti.append(altro_amb.lower().replace(" ", "_"))
            
            attrezzature = []
            if att_cucina: attrezzature.append("attrezzature_cucina")
            if att_cucina_man: attrezzature.append("attrezzi_manuali_cucina")
            if forno_micro: attrezzature.append("forno_microonde")
            if forno_elettr: attrezzature.append("forno_elettrico")
            if frigorifero: attrezzature.append("frigorifero")
            if macchina_caffe: attrezzature.append("macchina_caffe_bar")
            if macchina_ghiaccio: attrezzature.append("macchina_ghiaccio")
            if piastra_toast: attrezzature.append("piastra_toast_panini")
            if registratore: attrezzature.append("registratore_cassa")
            if videoterm: attrezzature.append("videoterminali")
            if att_ufficio: attrezzature.append("attrezzi_manuali_ufficio")
            if stampante_laser: attrezzature.append("stampante_laser")
            if stampante_ink: attrezzature.append("stampante_getto_inchiostro")
            if telefono: attrezzature.append("telefono_cellulare")
            if pc: attrezzature.append("pc_fisso_portatile")
            if tablet: attrezzature.append("tablet")
            if scaffali: attrezzature.append("scaffali_metallici")
            if att_sport: attrezzature.append("attrezzi_sportivi")
            if att_pulizia: attrezzature.append("attrezzi_pulizia")
            if autoveicoli: attrezzature.append("autoveicoli")
            if motoveicoli: attrezzature.append("motoveicoli")
            if scala: attrezzature.append("scala_portatile_doppia")
            if altro_att: attrezzature.append(altro_att.upper())
            
            mansioni = []
            if addetto_bar: mansioni.append("addetto_bar")
            if aiuto_cuoco: mansioni.append("addetto_cucina_aiutocuoco")
            if resp_sala: mansioni.append("responsabile_sala")
            if reception: mansioni.append("addetto_reception")
            if customer: mansioni.append("addetto_customer_service")
            if dirigente: mansioni.append("dirigente")
            if imp_admin: mansioni.append("impiegato_amministrativo")
            if impiegato: mansioni.append("impiegato")
            if segretario: mansioni.append("segretario_a")
            if istruttore: mansioni.append("istruttore_allenatore_preparatore")
            if bagnino: mansioni.append("bagnino")
            if estetista: mansioni.append("estetista")
            if parrucchiere: mansioni.append("parrucchiere")
            if add_servizi: mansioni.append("addetto_ai_servizi")
            if add_pulizia: mansioni.append("addetto_pulizia_base")
            if minuta_man: mansioni.append("addetto_minuta_manutenzione")
            if manutenzione: mansioni.append("addetto_manutenzione_tuttofare_officina")
            if magazzino: mansioni.append("addetto_magazzino")
            if operaio: mansioni.append("operaio_generico")
            if altra_mansione: mansioni.append(altra_mansione.upper())
            
            # Mappatura agenti chimici
            m_chem = {
                "acidi": "Acidi per laboratori didattici",
                "alcool": "Alcool etilico",
                "ammoniaca": "Ammoniaca",
                "candeggina": "Candeggina",
                "det_forni": "Detergente disincrostante forni",
                "det_stov": "Detergente stoviglie a mano",
                "det_lavast": "Detergente lavastoviglie",
                "det_pav": "Detergente per pavimenti",
                "det_wc": "Detergente per WC",
                "anticorrosivo": "Anticorrosivo",
                "antiruggine": "Antiruggine",
                "argon": "Argon",
                "azoto": "Azoto",
                "fumi_sald": "Fumi di saldatura",
                "grasso_lub": "Grasso lubrificante",
                "lubr_spray": "Lubrificanti spray (Svitol/Grasso)",
                "polveri": "Polveri da molatura",
                "acquaragia": "Acquaragia",
                "catalizzatore": "Catalizzatore vernici veicoli",
                "det_carrozz": "Detergente lucidatura carrozzerie",
                "fondo_vern": "Fondo verniciatura veicoli",
                "primer": "Primer verniciatura veicoli",
                "vernice_spray": "Vernice spray",
                "benzina": "Benzina",
                "gasolio": "Gasolio",
                "toner": "Toner",
                "tinta": "Tinta per capelli"
            }
            
            agenti_chimici = []
            for var, name in m_chem.items():
                if locals().get(var):
                    agenti_chimici.append(name)
            
            try:
                # Genera documento
                templates_dir = "templates"
                doc_buffer = genera_dvr(
                    azienda_data, 
                    ambienti, 
                    attrezzature, 
                    mansioni, 
                    agenti_chimici,
                    templates_dir
                )
                
                # Nome file
                filename = f"DVR_{nome.replace(' ', '_')}_{datetime.now().strftime('%d-%m-%Y')}.docx"
                
                st.success("✅ Documento generato con successo!")
                
                # Download button
                st.download_button(
                    label="📥 Scarica Documento DVR",
                    data=doc_buffer,
                    file_name=filename,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )
                
                # Istruzioni sommario
                st.info("""
                📋 **Aggiornare il sommario:**  
                Dopo aver aperto il documento in Word:
                1. Seleziona tutto (**Ctrl+A**)
                2. Premi **F9** (o clicca sul sommario → tasto destro → "Aggiorna campo")
                3. Scegli "Aggiorna intero sommario"
                """)
                
                # Riepilogo
                with st.expander("📋 Riepilogo selezioni"):
                    st.write(f"**Ambienti:** {len(ambienti)} selezionati")
                    st.write(f"**Attrezzature:** {len(attrezzature)} selezionate")  
                    st.write(f"**Mansioni:** {len(mansioni)} selezionate")
                    st.write(f"**Agenti chimici:** {len(agenti_chimici)} selezionati")
                    
            except Exception as e:
                st.error(f"❌ Errore durante la generazione: {str(e)}")
                st.exception(e)

