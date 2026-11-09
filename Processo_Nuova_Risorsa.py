import streamlit as st
import csv
import pandas as pd
import io
import unicodedata
import zipfile

# ------------------------------------------------------------
# Caricamento configurazione da Excel caricato dall'utente
# ------------------------------------------------------------
def load_config_from_bytes(data: bytes):
    cfg = pd.read_excel(io.BytesIO(data), sheet_name=None, engine="openpyxl")

    ris = cfg.get("Risorsa Interna", pd.DataFrame())

    if not ris.empty:
        ris = ris.rename(columns={c: str(c).strip() for c in ris.columns})
        if "Tipo" in ris.columns and "TIPO" not in ris.columns:
            ris = ris.rename(columns={"Tipo": "TIPO"})
        for col in ["Section", "Key/App", "Label/Gruppi/Value", "TIPO"]:
            if col in ris.columns:
                ris[col] = ris[col].fillna("").astype(str).str.strip()

    required = {"Section", "Key/App", "Label/Gruppi/Value"}
    if not ris.empty and required.issubset(ris.columns):
        ou_df = ris[ris["Section"] == "OU"][["Key/App", "Label/Gruppi/Value"]].rename(
            columns={"Key/App": "key", "Label/Gruppi/Value": "label"}
        )
        ou_options = dict(zip(ou_df["key"], ou_df["label"]))

        grp_df = ris[ris["Section"] == "InserimentoGruppi"][["Key/App", "Label/Gruppi/Value"]].rename(
            columns={"Key/App": "app", "Label/Gruppi/Value": "gruppi"}
        )
        gruppi = dict(zip(grp_df["app"], grp_df["gruppi"]))

        def_df = ris[ris["Section"] == "Defaults"][["Key/App", "Label/Gruppi/Value"]].rename(
            columns={"Key/App": "key", "Label/Gruppi/Value": "value"}
        )
        defaults = dict(zip(def_df["key"], def_df["value"]))
    else:
        ou_options, gruppi, defaults = {}, {}, {}

    # Tutte le righe Defaults con TIPO = PEL, nello stesso ordine del foglio.
    pel_items = []
    if not ris.empty and required.issubset(ris.columns) and "TIPO" in ris.columns:
        pel_df = ris[
            (ris["Section"] == "Defaults") &
            (ris["TIPO"].str.upper() == "PEL")
        ][["Key/App", "Label/Gruppi/Value"]]
        for _, row in pel_df.iterrows():
            key = str(row["Key/App"]).strip()
            value = str(row["Label/Gruppi/Value"]).strip()
            if value:
                pel_items.append((key, value))

    organigramma = {}
    if "organigramma" in cfg:
        org = cfg["organigramma"].iloc[:, :2].dropna(how="all")
        org.columns = ["label", "value"]
        organigramma = dict(zip(org["label"], org["value"]))

    return ou_options, gruppi, defaults, organigramma, pel_items


# ------------------------------------------------------------
# Configurazione pagina
# ------------------------------------------------------------
st.set_page_config(page_title="1.1 Nuova Risorsa Interna")
st.title("1.1 Nuova Risorsa Interna")

config_file = st.file_uploader(
    "Carica il file di configurazione (config.xlsx)",
    type=["xlsx"],
    help="Deve contenere il foglio 'Risorsa Interna' con campo Section"
)
if not config_file:
    st.warning("Per favore carica il file di configurazione per continuare.")
    st.stop()

ou_options, gruppi, defaults, organigramma, pel_items = load_config_from_bytes(config_file.read())

# ------------------------------------------------------------
# Valori letti dal file di configurazione
# ------------------------------------------------------------
def split_values(value):
    return [item.strip() for item in str(value).split(";") if item.strip()]


dl_standard = split_values(defaults.get("dl_standard", ""))
dl_vip = split_values(defaults.get("dl_vip", ""))

# Gruppi AD da tutte le chiavi grp_o365_* presenti nei Defaults.
o365_groups = []
for key, value in defaults.items():
    if str(key).startswith("grp_o365_") and value:
        for group in split_values(value):
            if group.startswith("365 "):
                group = "O" + group
            if group not in o365_groups:
                o365_groups.append(group)

ARCHIVE_PATH = r"\\srv_dati.consip.tesoro.it\AreaCondivisa\DEPSI\IC\AD_Modifiche"


# ------------------------------------------------------------
# Utility
# ------------------------------------------------------------
def normalize_name(value: str) -> str:
    nfkd = unicodedata.normalize("NFKD", value)
    ascii_str = nfkd.encode("ASCII", "ignore").decode()
    return ascii_str.replace(" ", "").replace("'", "").lower()


def genera_samaccountname(nome, cognome, secondo_nome="", secondo_cognome="", esterno=False):
    n = normalize_name(nome)
    sn = normalize_name(secondo_nome)
    c = normalize_name(cognome)
    sc = normalize_name(secondo_cognome)
    suffix = ".ext" if esterno else ""
    limit = 16 if esterno else 20

    candidate = f"{n}{sn}.{c}{sc}"
    if len(candidate) <= limit:
        return candidate + suffix

    candidate = f"{n[:1]}{sn[:1]}.{c}{sc}"
    if len(candidate) <= limit:
        return candidate + suffix

    base = f"{n[:1]}{sn[:1]}.{c}"
    return base[:limit] + suffix


def build_full_name(cognome, secondo_cognome, nome, secondo_nome, esterno=False):
    parts = [p for p in [cognome, secondo_cognome, nome, secondo_nome] if p]
    return " ".join(parts) + (" (esterno)" if esterno else "")


def unique_in_order(values):
    result = []
    for value in values:
        value = str(value).strip()
        if value and value not in result:
            result.append(value)
    return result


def build_entraid_groups(pel_items, selected_key):
    """
    Gruppi Entra ID ricavati dalle righe Defaults con TIPO = PEL.

    Regole mantenute:
    - dl_standard e dl_vip sono liste di distribuzione e non gruppi Entra ID;
    - pillole resta un canale e non viene scritto nel file gruppi Entra ID;
    - grp_salesforce_welfare si applica soltanto agli utenti standard;
    - tutte le altre righe PEL sono riportate una per riga.
    """
    groups = []
    for key, value in pel_items:
        key = str(key).strip()
        value = str(value).strip()
        if not value:
            continue
        if key in ("dl_standard", "dl_vip", "pillole"):
            continue
        if key == "grp_salesforce_welfare" and selected_key != "utenti_standard":
            continue
        groups.extend(split_values(value))
    return unique_in_order(groups)


def build_pillole_value(pel_items):
    for key, value in pel_items:
        if str(key).strip() == "pillole" and str(value).strip():
            return str(value).strip()
    return ""


def prepara_csv_standard(row, header):
    """CSV strutturato con quoting standard per i valori che lo richiedono."""
    buffer = io.StringIO(newline="")
    writer = csv.writer(buffer, quoting=csv.QUOTE_MINIMAL, lineterminator="\n")
    writer.writerow(header)
    writer.writerow(row)
    return buffer.getvalue()


def prepara_csv_righe(rows):
    """CSV semplice a due colonne, senza intestazione."""
    buffer = io.StringIO(newline="")
    writer = csv.writer(buffer, quoting=csv.QUOTE_MINIMAL, lineterminator="\n")
    writer.writerows(rows)
    return buffer.getvalue()


HEADER_UTENTE = [
    "sAMAccountName", "Creation", "OU", "Name", "DisplayName", "cn", "GivenName",
    "Surname", "employeeNumber", "employeeID", "department", "Description",
    "passwordNeverExpired", "ExpireDate", "userprincipalname", "mail", "mobile",
    "RimozioneGruppo", "InserimentoGruppo", "disable", "moveToOU", "telephoneNumber", "company"
]

HEADER_COMPUTER = [
    "Computer", "OU", "add_mail", "remove_mail", "add_mobile", "remove_mobile",
    "add_userprincipalname", "remove_userprincipalname", "disable", "moveToOU"
]


# ------------------------------------------------------------
# Input, invariati rispetto al flusso esistente
# ------------------------------------------------------------
st.subheader("Modulo Inserimento Nuova Risorsa Interna")
employee_id = st.text_input("Matricola", defaults.get("employee_id_default", "")).strip()
cognome = st.text_input("Cognome").strip().capitalize()
secondo_cognome = st.text_input("Secondo Cognome").strip().capitalize()
nome = st.text_input("Nome").strip().capitalize()
secondo_nome = st.text_input("Secondo Nome").strip().capitalize()
codice_fiscale = st.text_input("Codice Fiscale", "").strip()

if organigramma:
    dept_label = st.selectbox("Sigla Divisione-Area", ["-- Seleziona --"] + list(organigramma.keys()))
    department = (
        organigramma.get(dept_label, "")
        if dept_label != "-- Seleziona --"
        else defaults.get("department_default", "")
    )
else:
    department = st.text_input("Sigla Divisione-Area", defaults.get("department_default", "")).strip()

numero_telefono = st.text_input("Mobile (+39 già inserito)", "").replace(" ", "")
description = st.text_input("PC (lascia vuoto per <PC>)", "<PC>").strip()
resident_flag = st.checkbox("È Resident?")
numero_fisso = st.text_input("Numero fisso Resident (+39 già inserito)", "").strip() if resident_flag else ""
telephone_number = (
    f"+39 {numero_fisso}"
    if resident_flag and numero_fisso
    else defaults.get("telephone_interna", "")
)

ou_vals = list(ou_options.values()) if ou_options else []
default_ou = defaults.get("ou_default", ou_vals[0] if ou_vals else "")
index_default = ou_vals.index(default_ou) if default_ou in ou_vals else 0

label_ou = (
    st.selectbox("Tipologia Utente", ou_vals, index=index_default)
    if ou_vals
    else st.text_input("Tipologia Utente", "")
)
selected_key = list(ou_options.keys())[ou_vals.index(label_ou)] if ou_vals else ""
ou_value = ou_options[selected_key] if ou_vals else ""
company = defaults.get("company_interna", "")

data_operativa = st.text_input("Data operatività (gg/mm/aaaa)", "").strip()
profilazione = st.checkbox("Profilazione SM?")
sm_lines = [line.strip() for line in st.text_area("SM (una per riga)", "").splitlines() if line.strip()] if profilazione else []

dl_list = dl_standard if selected_key == "utenti_standard" else dl_vip if selected_key == "utenti_vip" else []


# ------------------------------------------------------------
# Anteprima casella
# ------------------------------------------------------------
if st.button("Template per Posta Elettronica"):
    sam = genera_samaccountname(nome, cognome, secondo_nome, secondo_cognome, False)
    cn = build_full_name(cognome, secondo_cognome, nome, secondo_nome, False)

    st.markdown("Ciao.  \nRichiedo cortesemente la definizione di una casella di posta come sottoindicato.")
    st.dataframe(
        pd.DataFrame(
            [
                ["Tipo Utenza", "Remota"],
                ["Utenza", sam],
                ["Alias", sam],
                ["Display name", cn],
                ["Common name", cn],
                ["e-mail", f"{sam}@consip.it"],
                ["e-mail secondaria", f"{sam}@consipspa.mail.onmicrosoft.com"],
            ],
            columns=["Campo", "Valore"]
        ),
        hide_index=True,
        use_container_width=True
    )

    if dl_list:
        st.markdown(f"Il giorno **{data_operativa}** occorre inserire la casella nelle DL:")
        for dl in dl_list:
            st.markdown(f"- {dl}")

    if profilazione:
        st.markdown("Profilare su SM:")
        for sm in sm_lines:
            st.markdown(f"- {sm}")

    st.markdown("**Aggiungere utenza ai gruppi Entra ID:**")
    for group in build_entraid_groups(pel_items, selected_key):
        st.markdown(f"- {group}")

    pillole_value = build_pillole_value(pel_items)
    if pillole_value:
        st.markdown(f"- canale {pillole_value}")

    st.markdown("Grazie  \nSaluti")


# ------------------------------------------------------------
# Generazione nuovo bundle di provisioning
# ------------------------------------------------------------
if st.button("Genera CSV"):
    if not cognome or not nome:
        st.error("Inserire almeno Cognome e Nome prima di generare i file.")
        st.stop()

    sam = genera_samaccountname(nome, cognome, secondo_nome, secondo_cognome, False)
    cn = build_full_name(cognome, secondo_cognome, nome, secondo_nome, False)
    given = f"{nome} {secondo_nome}".strip()
    surname = f"{cognome} {secondo_cognome}".strip()
    mobile = f"+39 {numero_telefono}" if numero_telefono else ""

    # Nome leggibile richiesto nei file 1-4.
    persona_nome_file = " ".join(
        part for part in [cognome, secondo_cognome, nome, secondo_nome] if part
    )
    # Nome con underscore richiesto per il file computer e per il bundle.
    persona_nome_underscore = "_".join(
        part for part in [cognome, secondo_cognome, nome, secondo_nome] if part
    )

    # 1 - Utenza: stesso contenuto del precedente _utente.csv.
    row_utente = [
        sam, "SI", ou_value, cn, cn, cn, given, surname,
        codice_fiscale, employee_id, department, description or "<PC>",
        "No", "", f"{sam}@consip.it", f"{sam}@consip.it", mobile,
        "", "", "", "", telephone_number, company
    ]
    csv_utente = prepara_csv_standard(row_utente, HEADER_UTENTE)

    # 2 - Casella: formato Campo,Valore.
    casella_rows = [
        ["Tipo Utenza", "Remota"],
        ["Utenza", sam],
        ["Alias", sam],
        ["Display name", cn],
        ["Common name", cn],
        ["e-mail", f"{sam}@consip.it"],
        ["e-mail secondaria", f"{sam}@consipspa.mail.onmicrosoft.com"],
    ]
    csv_casella = prepara_csv_righe(casella_rows)

    # 3 - Profilazione gruppi AD.
    inserimento_gruppo_val = gruppi.get("interna", "") or ""
    gruppi_interna = []
    for group in split_values(inserimento_gruppo_val):
        if group.startswith("365 "):
            group = "O" + group
        gruppi_interna.append(group)

    gruppi_ad = unique_in_order(o365_groups + gruppi_interna)
    row_gruppi_ad = [""] * len(HEADER_UTENTE)
    row_gruppi_ad[0] = sam
    row_gruppi_ad[HEADER_UTENTE.index("InserimentoGruppo")] = ";".join(gruppi_ad)
    csv_gruppi_ad = prepara_csv_standard(row_gruppi_ad, HEADER_UTENTE)

    # 4 - Gruppi Entra ID, uno per riga.
    gruppi_entraid = build_entraid_groups(pel_items, selected_key)
    entra_lines = [f"Per l'utenza {sam} aggiungere i gruppi EntraID:"]
    entra_lines.extend(gruppi_entraid)
    csv_entraid = "\n".join(entra_lines) + "\n"

    # 5 - Computer: stesso contenuto del precedente _computer.csv.
    row_computer = [
        description or "", "", f"{sam}@consip.it", "", mobile, "", cn, "", "", ""
    ]
    csv_computer = prepara_csv_standard(row_computer, HEADER_COMPUTER)

    filename_utente = f"1_{persona_nome_file}_utenza.csv"
    filename_casella = f"2_{persona_nome_file}_casella.csv"
    filename_gruppi_ad = f"3_{persona_nome_file}_profilazione_Gruppi_AD.csv"
    filename_entraid = f"4_{persona_nome_file}_profilazione_Gruppi_EntraID.csv"
    filename_computer = f"5_{persona_nome_underscore}_computer.csv"

    # Primo testo richiesto: provisioning ordinato dei file 1-4.
    subject_provisioning = f"02 - Provisioning Utente (utenza/Casella/Profilazione) - {cognome} {nome}"
    message_provisioning = (
        "Ciao,\n"
        "si richiede di elaborare i file contenuti nella cartella compressa .zip nel seguente ordine:\n"
        f"1. {filename_utente} per la creazione di un'utenza On prem su AD\n"
        f"2. {filename_casella} per la creazione di un'utenza ibrida come indicato\n"
        f"3. {filename_gruppi_ad} per inserimento dei gruppi AD\n"
        f"4. {filename_entraid} per inserimento dei gruppi EntraID\n"
    )

    # Secondo testo richiesto: liste di distribuzione standard verso GIIT.
    subject_dl = "08 - Lista di distribuzione standard -> GIIT"
    dl_lines = [
        f"{sam}@consip.it in data {data_operativa} inserire nelle DL:"
    ]
    dl_lines.extend(dl_list)
    message_dl = "\n".join(dl_lines) + "\n"

    # Anteprime a video.
    st.subheader(subject_provisioning)
    st.text(message_provisioning)

    st.subheader(subject_dl)
    st.text(message_dl)

    st.subheader("Anteprima 1 - Utenza")
    st.dataframe(pd.DataFrame([row_utente], columns=HEADER_UTENTE), use_container_width=True)

    st.subheader("Anteprima 2 - Casella")
    st.dataframe(pd.DataFrame(casella_rows, columns=["Campo", "Valore"]), hide_index=True, use_container_width=True)

    st.subheader("Anteprima 3 - Gruppi AD")
    st.dataframe(pd.DataFrame([row_gruppi_ad], columns=HEADER_UTENTE), use_container_width=True)

    st.subheader("Anteprima 4 - Gruppi Entra ID")
    st.text(csv_entraid)

    st.subheader("Anteprima 5 - Computer")
    st.dataframe(pd.DataFrame([row_computer], columns=HEADER_COMPUTER), use_container_width=True)

    # Download singoli file.
    st.download_button("Scarica 1 - Utenza", csv_utente, filename_utente, "text/csv")
    st.download_button("Scarica 2 - Casella", csv_casella, filename_casella, "text/csv")
    st.download_button("Scarica 3 - Gruppi AD", csv_gruppi_ad, filename_gruppi_ad, "text/csv")
    st.download_button("Scarica 4 - Gruppi Entra ID", csv_entraid, filename_entraid, "text/csv")
    st.download_button("Scarica 5 - Computer", csv_computer, filename_computer, "text/csv")

    # Bundle ZIP con tutti e cinque i file.
    # Il file 5 è incluso nello ZIP, ma non è citato nel corpo del messaggio 02.
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
        zip_file.writestr(filename_utente, csv_utente)
        zip_file.writestr(filename_casella, csv_casella)
        zip_file.writestr(filename_gruppi_ad, csv_gruppi_ad)
        zip_file.writestr(filename_entraid, csv_entraid)
        zip_file.writestr(filename_computer, csv_computer)

    zip_buffer.seek(0)
    bundle_name = f"Provisioning_{persona_nome_underscore}.zip"
    st.download_button(
        "Scarica bundle completo ZIP",
        data=zip_buffer.getvalue(),
        file_name=bundle_name,
        mime="application/zip"
    )

    st.success(f"Bundle generato correttamente per '{sam}'.")
