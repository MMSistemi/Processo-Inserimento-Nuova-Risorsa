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
    cfg = pd.read_excel(
        io.BytesIO(data),
        sheet_name=None,
        engine="openpyxl"
    )

    if "Risorsa Interna" in cfg:
        ris = cfg["Risorsa Interna"]
    else:
        ris = pd.DataFrame()

    # Normalizzazione colonne
    if not ris.empty:
        ris = ris.rename(
            columns={
                column: str(column).strip()
                for column in ris.columns
            }
        )

        if "Tipo" in ris.columns and "TIPO" not in ris.columns:
            ris = ris.rename(
                columns={"Tipo": "TIPO"}
            )

        for column in [
            "Section",
            "Key/App",
            "Label/Gruppi/Value",
            "TIPO"
        ]:
            if column in ris.columns:
                ris[column] = (
                    ris[column]
                    .fillna("")
                    .astype(str)
                    .str.strip()
                )

    required_columns = {
        "Section",
        "Key/App",
        "Label/Gruppi/Value"
    }

    if not ris.empty and required_columns.issubset(ris.columns):

        # OU
        ou_df = ris[
            ris["Section"] == "OU"
        ][
            ["Key/App", "Label/Gruppi/Value"]
        ].rename(
            columns={
                "Key/App": "key",
                "Label/Gruppi/Value": "label"
            }
        )

        ou_options = dict(
            zip(
                ou_df["key"],
                ou_df["label"]
            )
        )

        # Inserimento gruppi
        gruppi_df = ris[
            ris["Section"] == "InserimentoGruppi"
        ][
            ["Key/App", "Label/Gruppi/Value"]
        ].rename(
            columns={
                "Key/App": "app",
                "Label/Gruppi/Value": "gruppi"
            }
        )

        gruppi = dict(
            zip(
                gruppi_df["app"],
                gruppi_df["gruppi"]
            )
        )

        # Defaults
        defaults_df = ris[
            ris["Section"] == "Defaults"
        ][
            ["Key/App", "Label/Gruppi/Value"]
        ].rename(
            columns={
                "Key/App": "key",
                "Label/Gruppi/Value": "value"
            }
        )

        defaults = dict(
            zip(
                defaults_df["key"],
                defaults_df["value"]
            )
        )

    else:
        ou_options = {}
        gruppi = {}
        defaults = {}

    # Estrazione di tutte le righe Defaults con TIPO = PEL
    # Viene mantenuto l'ordine originale del foglio Excel.
    pel_items = []

    if (
        not ris.empty
        and required_columns.issubset(ris.columns)
        and "TIPO" in ris.columns
    ):
        pel_df = ris[
            (ris["Section"] == "Defaults")
            & (ris["TIPO"].str.upper() == "PEL")
        ][
            ["Key/App", "Label/Gruppi/Value"]
        ]

        for _, row in pel_df.iterrows():
            key = str(row["Key/App"]).strip()
            value = str(
                row["Label/Gruppi/Value"]
            ).strip()

            if value:
                pel_items.append(
                    (key, value)
                )

    # Estrazione organigramma
    organigramma = {}

    if "organigramma" in cfg:
        organigramma_df = (
            cfg["organigramma"]
            .iloc[:, :2]
            .dropna(how="all")
        )

        organigramma_df.columns = [
            "label",
            "value"
        ]

        organigramma = dict(
            zip(
                organigramma_df["label"],
                organigramma_df["value"]
            )
        )

    return (
        ou_options,
        gruppi,
        defaults,
        organigramma,
        pel_items
    )


# ------------------------------------------------------------
# Configurazione pagina
# ------------------------------------------------------------
st.set_page_config(
    page_title="1.1 Nuova Risorsa Interna"
)

st.title(
    "1.1 Nuova Risorsa Interna"
)


# ------------------------------------------------------------
# Caricamento file di configurazione
# ------------------------------------------------------------
config_file = st.file_uploader(
    "Carica il file di configurazione (config.xlsx)",
    type=["xlsx"],
    help=(
        "Deve contenere il foglio 'Risorsa Interna' "
        "con il campo Section"
    )
)

if not config_file:
    st.warning(
        "Per favore carica il file di configurazione "
        "per continuare."
    )
    st.stop()


(
    ou_options,
    gruppi,
    defaults,
    organigramma,
    pel_items
) = load_config_from_bytes(
    config_file.read()
)


# ------------------------------------------------------------
# Funzioni di gestione configurazione
# ------------------------------------------------------------
def split_values(value):
    return [
        item.strip()
        for item in str(value).split(";")
        if item.strip()
    ]


# Liste di distribuzione
dl_standard = split_values(
    defaults.get(
        "dl_standard",
        ""
    )
)

dl_vip = split_values(
    defaults.get(
        "dl_vip",
        ""
    )
)


# ------------------------------------------------------------
# Gruppi AD letti da tutte le chiavi grp_o365_*
# ------------------------------------------------------------
o365_groups = []

for key, value in defaults.items():
    if str(key).startswith("grp_o365_") and value:
        for group in split_values(value):

            # Correzione automatica nel caso manchi la O iniziale
            if group.startswith("365 "):
                group = "O" + group

            if group not in o365_groups:
                o365_groups.append(group)


# Percorso archivio
ARCHIVE_PATH = (
    r"\\srv_dati.consip.tesoro.it"
    r"\AreaCondivisa\DEPSI\IC\AD_Modifiche"
)


# ------------------------------------------------------------
# Funzioni di utilità
# ------------------------------------------------------------
def normalize_name(value: str) -> str:
    normalized = unicodedata.normalize(
        "NFKD",
        value
    )

    ascii_value = (
        normalized
        .encode(
            "ASCII",
            "ignore"
        )
        .decode()
    )

    return (
        ascii_value
        .replace(" ", "")
        .replace("'", "")
        .lower()
    )


def genera_samaccountname(
    nome,
    cognome,
    secondo_nome="",
    secondo_cognome="",
    esterno=False
):
    nome_normalizzato = normalize_name(nome)
    secondo_nome_normalizzato = normalize_name(
        secondo_nome
    )
    cognome_normalizzato = normalize_name(
        cognome
    )
    secondo_cognome_normalizzato = normalize_name(
        secondo_cognome
    )

    suffix = ".ext" if esterno else ""
    limit = 16 if esterno else 20

    candidate_1 = (
        f"{nome_normalizzato}"
        f"{secondo_nome_normalizzato}."
        f"{cognome_normalizzato}"
        f"{secondo_cognome_normalizzato}"
    )

    if len(candidate_1) <= limit:
        return candidate_1 + suffix

    candidate_2 = (
        f"{nome_normalizzato[:1]}"
        f"{secondo_nome_normalizzato[:1]}."
        f"{cognome_normalizzato}"
        f"{secondo_cognome_normalizzato}"
    )

    if len(candidate_2) <= limit:
        return candidate_2 + suffix

    base = (
        f"{nome_normalizzato[:1]}"
        f"{secondo_nome_normalizzato[:1]}."
        f"{cognome_normalizzato}"
    )

    return base[:limit] + suffix


def build_full_name(
    cognome,
    secondo_cognome,
    nome,
    secondo_nome,
    esterno=False
):
    parts = [
        part
        for part in [
            cognome,
            secondo_cognome,
            nome,
            secondo_nome
        ]
        if part
    ]

    full_name = " ".join(parts)

    if esterno:
        full_name += " (esterno)"

    return full_name


def unique_in_order(values):
    result = []

    for value in values:
        value = str(value).strip()

        if value and value not in result:
            result.append(value)

    return result


# ------------------------------------------------------------
# Costruzione gruppi Entra ID
# ------------------------------------------------------------
def build_entraid_groups(
    pel_items,
    selected_key
):
    """
    Regole:
    - dl_standard e dl_vip non sono gruppi Entra ID;
    - pillole è un canale e non viene inserito;
    - grp_salesforce_welfare viene inserito solo
      per gli utenti standard;
    - tutte le altre righe PEL vengono riportate
      una per riga;
    - viene mantenuto l'ordine del file di configurazione.
    """

    entra_groups = []

    for key, value in pel_items:
        key = str(key).strip()
        value = str(value).strip()

        if not value:
            continue

        if key in (
            "dl_standard",
            "dl_vip",
            "pillole"
        ):
            continue

        if (
            key == "grp_salesforce_welfare"
            and selected_key != "utenti_standard"
        ):
            continue

        entra_groups.extend(
            split_values(value)
        )

    return unique_in_order(
        entra_groups
    )


# ------------------------------------------------------------
# Costruzione CSV con logica precedente
# ------------------------------------------------------------
def auto_quote(
    fields,
    quotechar='"',
    predicate=lambda value: " " in value
):
    """
    Racchiude tra doppi apici ogni valore che contiene
    almeno uno spazio.

    Esempio:
    Spagnuolo Luca -> "Spagnuolo Luca"
    """

    output = []

    for field in fields:
        value = str(field)

        if predicate(value):
            output.append(
                f'{quotechar}{value}{quotechar}'
            )
        else:
            output.append(value)

    return output


def prepara_csv_standard(
    row,
    header
):
    """
    Mantiene la costruzione precedente dei CSV:
    ogni campo contenente almeno uno spazio viene
    racchiuso tra doppi apici.
    """

    buffer = io.StringIO()

    writer = csv.writer(
        buffer,
        quoting=csv.QUOTE_NONE,
        escapechar="\\",
        lineterminator="\n"
    )

    quoted_row = auto_quote(
        row,
        quotechar='"',
        predicate=lambda value: " " in value
    )

    writer.writerow(header)
    writer.writerow(quoted_row)

    return buffer.getvalue()


# ------------------------------------------------------------
# Header CSV
# ------------------------------------------------------------
HEADER_UTENTE = [
    "sAMAccountName",
    "Creation",
    "OU",
    "Name",
    "DisplayName",
    "cn",
    "GivenName",
    "Surname",
    "employeeNumber",
    "employeeID",
    "department",
    "Description",
    "passwordNeverExpired",
    "ExpireDate",
    "userprincipalname",
    "mail",
    "mobile",
    "RimozioneGruppo",
    "InserimentoGruppo",
    "disable",
    "moveToOU",
    "telephoneNumber",
    "company"
]


HEADER_COMPUTER = [
    "Computer",
    "OU",
    "add_mail",
    "remove_mail",
    "add_mobile",
    "remove_mobile",
    "add_userprincipalname",
    "remove_userprincipalname",
    "disable",
    "moveToOU"
]


# ------------------------------------------------------------
# Modulo inserimento nuova risorsa interna
# ------------------------------------------------------------
st.subheader(
    "Modulo Inserimento Nuova Risorsa Interna"
)

employee_id = st.text_input(
    "Matricola",
    defaults.get(
        "employee_id_default",
        ""
    )
).strip()

cognome = (
    st.text_input("Cognome")
    .strip()
    .capitalize()
)

secondo_cognome = (
    st.text_input("Secondo Cognome")
    .strip()
    .capitalize()
)

nome = (
    st.text_input("Nome")
    .strip()
    .capitalize()
)

secondo_nome = (
    st.text_input("Secondo Nome")
    .strip()
    .capitalize()
)

codice_fiscale = st.text_input(
    "Codice Fiscale",
    ""
).strip()


# ------------------------------------------------------------
# Dropdown organigramma
# ------------------------------------------------------------
if organigramma:
    department_label = st.selectbox(
        "Sigla Divisione-Area",
        ["-- Seleziona --"]
        + list(
            organigramma.keys()
        )
    )

    if department_label != "-- Seleziona --":
        department = organigramma.get(
            department_label,
            ""
        )
    else:
        department = defaults.get(
            "department_default",
            ""
        )
else:
    department = st.text_input(
        "Sigla Divisione-Area",
        defaults.get(
            "department_default",
            ""
        )
    ).strip()


numero_telefono = (
    st.text_input(
        "Mobile (+39 già inserito)",
        ""
    )
    .replace(" ", "")
)

description = st.text_input(
    "PC (lascia vuoto per <PC>)",
    "<PC>"
).strip()

resident_flag = st.checkbox(
    "È Resident?"
)

if resident_flag:
    numero_fisso = st.text_input(
        "Numero fisso Resident (+39 già inserito)",
        ""
    ).strip()
else:
    numero_fisso = ""

if resident_flag and numero_fisso:
    telephone_number = (
        f"+39 {numero_fisso}"
    )
else:
    telephone_number = defaults.get(
        "telephone_interna",
        ""
    )


# ------------------------------------------------------------
# Tipologia utente e OU
# ------------------------------------------------------------
ou_values = (
    list(ou_options.values())
    if ou_options
    else []
)

default_ou = defaults.get(
    "ou_default",
    ou_values[0]
    if ou_values
    else ""
)

index_default = 0

if default_ou in ou_values:
    index_default = ou_values.index(
        default_ou
    )

if ou_values:
    label_ou = st.selectbox(
        "Tipologia Utente",
        ou_values,
        index=index_default
    )
else:
    label_ou = st.text_input(
        "Tipologia Utente",
        ""
    )

if ou_values:
    selected_key = list(
        ou_options.keys()
    )[ou_values.index(label_ou)]

    ou_value = ou_options[
        selected_key
    ]
else:
    selected_key = ""
    ou_value = ""


company = defaults.get(
    "company_interna",
    ""
)


data_operativa = st.text_input(
    "Data operatività (gg/mm/aaaa)",
    ""
).strip()


# ------------------------------------------------------------
# Dati aggiuntivi per il flusso operativo completo
# ------------------------------------------------------------
mouse_id = st.text_input(
    "Mouse ID",
    "",
    help=(
        "Campo facoltativo. Se lasciato vuoto, "
        "nel testo verrà riportato #id."
    )
).strip()

mouse_sn = st.text_input(
    "Mouse SN",
    "",
    help=(
        "Campo facoltativo. Se lasciato vuoto, "
        "nel testo verrà riportato #sn."
    )
).strip()

hub_id = st.text_input(
    "HUB Sandberg ID",
    "",
    help=(
        "Campo facoltativo. Se lasciato vuoto, "
        "nel testo verrà riportato #id."
    )
).strip()

firma_digitale = st.checkbox(
    "Firma digitale?"
)


profilazione = st.checkbox(
    "Profilazione SM?"
)

if profilazione:
    sm_lines = [
        line.strip()
        for line in st.text_area(
            "SM (una per riga)",
            ""
        ).splitlines()
        if line.strip()
    ]
else:
    sm_lines = []


# ------------------------------------------------------------
# Liste di distribuzione in base alla tipologia utente
# ------------------------------------------------------------
if selected_key == "utenti_standard":
    dl_list = dl_standard
elif selected_key == "utenti_vip":
    dl_list = dl_vip
else:
    dl_list = []


# ------------------------------------------------------------
# Generazione file e bundle ZIP
# Rimane unicamente il pulsante Genera CSV
# ------------------------------------------------------------
if st.button("Genera CSV"):

    if not cognome or not nome:
        st.error(
            "Inserire almeno Cognome e Nome "
            "prima di generare i file."
        )
        st.stop()

    sam = genera_samaccountname(
        nome,
        cognome,
        secondo_nome,
        secondo_cognome,
        False
    )

    cn = build_full_name(
        cognome,
        secondo_cognome,
        nome,
        secondo_nome,
        False
    )

    given_name = (
        f"{nome} {secondo_nome}"
        .strip()
    )

    surname = (
        f"{cognome} {secondo_cognome}"
        .strip()
    )

    if numero_telefono:
        mobile = (
            f"+39 {numero_telefono}"
        )
    else:
        mobile = ""

    # Nome leggibile per i file 1-4
    persona_nome_file = " ".join(
        part
        for part in [
            cognome,
            secondo_cognome,
            nome,
            secondo_nome
        ]
        if part
    )

    # Nome con underscore per file computer e ZIP
    persona_nome_underscore = "_".join(
        part
        for part in [
            cognome,
            secondo_cognome,
            nome,
            secondo_nome
        ]
        if part
    )

    # --------------------------------------------------------
    # 1 - File utenza
    # Stessa struttura del precedente _utente.csv
    # --------------------------------------------------------
    row_utente = [
        sam,
        "SI",
        ou_value,
        cn,
        cn,
        cn,
        given_name,
        surname,
        codice_fiscale,
        employee_id,
        department,
        description or "<PC>",
        "No",
        "",
        f"{sam}@consip.it",
        f"{sam}@consip.it",
        mobile,
        "",
        "",
        "",
        "",
        telephone_number,
        company
    ]

    csv_utente = prepara_csv_standard(
        row_utente,
        HEADER_UTENTE
    )

    # --------------------------------------------------------
    # 2 - File casella
    # Formato testuale esatto con separatore ---
    # --------------------------------------------------------
    casella_lines = [
        "Tipo Utenza --- Remota",
        f"Utenza --- {sam}",
        f"Alias --- {sam}",
        f"Display name --- {cn}",
        f"Common name --- {cn}",
        f"e-mail --- {sam}@consip.it",
        (
            "e-mail secondaria --- "
            f"{sam}@consipspa.mail.onmicrosoft.com"
        )
    ]

    csv_casella = (
        "\n".join(casella_lines)
        + "\n"
    )

    # --------------------------------------------------------
    # 3 - File profilazione gruppi AD
    # --------------------------------------------------------
    inserimento_gruppo_value = (
        gruppi.get(
            "interna",
            ""
        )
        or ""
    )

    gruppi_interna = []

    for group in split_values(
        inserimento_gruppo_value
    ):
        if group.startswith("365 "):
            group = "O" + group

        gruppi_interna.append(
            group
        )

    gruppi_ad = unique_in_order(
        o365_groups
        + gruppi_interna
    )

    row_gruppi_ad = [
        ""
    ] * len(HEADER_UTENTE)

    row_gruppi_ad[0] = sam

    inserimento_gruppo_index = (
        HEADER_UTENTE.index(
            "InserimentoGruppo"
        )
    )

    row_gruppi_ad[
        inserimento_gruppo_index
    ] = ";".join(gruppi_ad)

    csv_gruppi_ad = prepara_csv_standard(
        row_gruppi_ad,
        HEADER_UTENTE
    )

    # --------------------------------------------------------
    # 4 - File profilazione gruppi Entra ID
    # --------------------------------------------------------
    gruppi_entraid = build_entraid_groups(
        pel_items,
        selected_key
    )

    entra_lines = [
        (
            f"Per l'utenza {sam} "
            "aggiungere i gruppi EntraID:"
        )
    ]

    entra_lines.extend(
        gruppi_entraid
    )

    csv_entraid = (
        "\n".join(entra_lines)
        + "\n"
    )

    # --------------------------------------------------------
    # 5 - File computer
    # Stessa struttura del precedente _computer.csv
    # --------------------------------------------------------
    row_computer = [
        description or "",
        "",
        f"{sam}@consip.it",
        "",
        mobile,
        "",
        cn,
        "",
        "",
        ""
    ]

    csv_computer = prepara_csv_standard(
        row_computer,
        HEADER_COMPUTER
    )

    # --------------------------------------------------------
    # Nomi dei file
    # --------------------------------------------------------
    filename_utente = (
        f"1_{persona_nome_file}_utenza.csv"
    )

    filename_casella = (
        f"2_{persona_nome_file}_casella.csv"
    )

    filename_gruppi_ad = (
        f"3_{persona_nome_file}"
        "_profilazione_Gruppi_AD.csv"
    )

    filename_entraid = (
        f"4_{persona_nome_file}"
        "_profilazione_Gruppi_EntraID.csv"
    )

    filename_computer = (
        f"5_{persona_nome_underscore}"
        "_computer.csv"
    )

    # --------------------------------------------------------
    # 02 - Provisioning utente
    # --------------------------------------------------------
    subject_provisioning = (
        "02 - Provisioning Utente "
        "(utenza/Casella/Profilazione) - "
        f"{cn} -> GIIT"
    )

    message_provisioning = (
        "Ciao,\n"
        "si richiede di elaborare i file contenuti "
        "nella cartella compressa .zip "
        "nel seguente ordine:\n"
        f"1. {filename_utente} per la creazione "
        "di un'utenza On prem su AD\n"
        f"2. {filename_casella} per la creazione "
        "di un'utenza ibrida come indicato\n"
        f"3. {filename_gruppi_ad} per inserimento "
        "dei gruppi AD\n"
        f"4. {filename_entraid} per inserimento "
        "dei gruppi EntraID\n"
    )

    # --------------------------------------------------------
    # 08 - Liste di distribuzione
    # --------------------------------------------------------
    subject_dl = (
        "08 - Lista di distribuzione standard -> GIIT"
    )

    dl_lines = [
        (
            f"{sam}@consip.it "
            f"in data {data_operativa} "
            "inserire nelle DL:"
        )
    ]

    dl_lines.extend(
        dl_list
    )

    message_dl = (
        "\n".join(dl_lines)
        + "\n"
    )

    # --------------------------------------------------------
    # Dati utilizzati nella guida del flusso
    # --------------------------------------------------------
    pc_asset = (
        description
        or "<PC>"
    )

    mouse_id_output = (
        mouse_id
        or "#id"
    )

    mouse_sn_output = (
        mouse_sn
        or "#sn"
    )

    hub_id_output = (
        hub_id
        or "#id"
    )

    mail_utente = (
        f"{sam}@consip.it"
    )

    # Nome + cognome + un carattere per il punto
    lunghezza_nome_cognome_sic = (
        len(normalize_name(nome))
        + len(normalize_name(cognome))
        + 1
    )

    verifica_sic_necessaria = (
        lunghezza_nome_cognome_sic > 20
    )

    # --------------------------------------------------------
    # 01 - Assegnazione Asset
    # --------------------------------------------------------
    subject_assegnazione_asset = (
        f"01 - Assegnazione Asset - {cn} -> GSI"
    )

    message_assegnazione_asset = (
        f"Per la nuova risorsa {cn} assegnare asset:\n"
        f"- PC Nome asset {pc_asset}\n"
        f"- mouse ID {mouse_id_output} "
        f"e SN {mouse_sn_output}\n"
        f"- HUB Sandberg ID {hub_id_output}\n"
    )

    # --------------------------------------------------------
    # 03 - Provisioning PC
    # --------------------------------------------------------
    subject_provisioning_pc = (
        f"03 - Provisioning PC - {cn} -> GHDW"
    )

    message_provisioning_pc = (
        "Ciao,\n"
        f"per {cn} ritirare dal 599 "
        "per CONFIGURAZIONE i seguenti asset:\n"
        f"- PC Nome asset {pc_asset}\n"
        f"- mouse ID {mouse_id_output} "
        f"e SN {mouse_sn_output}\n"
        f"- HUB Sandberg ID {hub_id_output}\n"
        "\n"
        f"la risorsa prende servizio {data_operativa}\n"
    )

    # --------------------------------------------------------
    # 04 - Caricamento foto
    # --------------------------------------------------------
    subject_caricamento_foto = (
        f"04 - Caricamento Foto - {cn} -> GA"
    )

    message_caricamento_foto = (
        "Ciao,\n"
        "si richiede il caricamento della foto "
        f"in allegato per {cn}.\n"
    )

    # --------------------------------------------------------
    # 05 - Firma digitale
    # --------------------------------------------------------
    subject_firma_digitale = (
        f"05 - Firma digitale - {cn} -> MM"
    )

    message_firma_digitale = (
        f"Assegnare Firma Digitale a {cn}.\n"
    )

    # --------------------------------------------------------
    # 06 - Verifica ed eventuale modifica SIC
    # --------------------------------------------------------
    subject_verifica_sic = (
        "06 - Verifica ed Eventuale Modifica SIC - "
        f"{cn} -> GA"
    )

    message_verifica_sic = (
        "Ciao, si richiede verifica che per utente "
        f"{cn} sia indicato sul SIC con la seguente "
        f'email "{mail_utente}".\n'
    )

    # --------------------------------------------------------
    # 07 - Accesso intranet PSN
    # --------------------------------------------------------
    subject_accesso_psn = (
        "07 - Accesso intranet PSN per dipendente "
        "-> GITT -> GA"
    )

    message_accesso_psn = (
        "Attività momentaneamente esclusa dal flusso "
        "perché attualmente gestita via email.\n"
    )

    # --------------------------------------------------------
    # 09 - Firma Asset
    # --------------------------------------------------------
    subject_firma_asset = (
        f"09 - Firma Asset - {cn} -> GHDW"
    )

    message_firma_asset = (
        "Ciao,\n"
        f"si richiede il giorno {data_operativa} "
        "di far firmare gli asset ritirati dal 599 "
        f"per la consegna a {cn}.\n"
    )

    # --------------------------------------------------------
    # Costruzione della guida completa
    # --------------------------------------------------------
    guida_flusso_lines = [
        subject_assegnazione_asset,
        message_assegnazione_asset,

        subject_provisioning,
        message_provisioning,

        subject_provisioning_pc,
        message_provisioning_pc,

        subject_caricamento_foto,
        message_caricamento_foto
    ]

    # Il punto 05 viene aggiunto solo se selezionato
    if firma_digitale:
        guida_flusso_lines.extend(
            [
                subject_firma_digitale,
                message_firma_digitale
            ]
        )

    # Il punto 06 viene aggiunto solo se il limite
    # nome + cognome + 1 supera 20 caratteri
    if verifica_sic_necessaria:
        guida_flusso_lines.extend(
            [
                subject_verifica_sic,
                message_verifica_sic
            ]
        )

    # Punto 07 momentaneamente escluso,
    # ma mostrato come nota operativa
    guida_flusso_lines.extend(
        [
            subject_accesso_psn,
            message_accesso_psn,

            subject_dl,
            message_dl,

            subject_firma_asset,
            message_firma_asset
        ]
    )

    guida_flusso = "\n".join(
        guida_flusso_lines
    )

    # --------------------------------------------------------
    # Visualizzazione guida pratica completa
    # --------------------------------------------------------
    st.subheader(
        "Guida pratica flusso operativo 01-09"
    )

    st.text(
        guida_flusso
    )

    # --------------------------------------------------------
    # Informazioni sui punti condizionali
    # --------------------------------------------------------
    if not firma_digitale:
        st.info(
            "Punto 05 non visualizzato: "
            "la firma digitale non è stata selezionata."
        )

    if not verifica_sic_necessaria:
        st.info(
            "Punto 06 non visualizzato: "
            f"nome + cognome + 1 corrisponde a "
            f"{lunghezza_nome_cognome_sic} caratteri "
            "e non supera il limite di 20."
        )

    # --------------------------------------------------------
    # Visualizzazione messaggi già esistenti
    # --------------------------------------------------------
    st.subheader(
        subject_provisioning
    )

    st.text(
        message_provisioning
    )

    st.subheader(
        subject_dl
    )

    st.text(
        message_dl
    )

    # --------------------------------------------------------
    # Anteprima file 1
    # --------------------------------------------------------
    st.subheader(
        "Anteprima 1 - Utenza"
    )

    st.dataframe(
        pd.DataFrame(
            [row_utente],
            columns=HEADER_UTENTE
        ),
        use_container_width=True
    )

    # --------------------------------------------------------
    # Anteprima file 2
    # --------------------------------------------------------
    st.subheader(
        "Anteprima 2 - Casella"
    )

    st.text(
        csv_casella
    )

    # --------------------------------------------------------
    # Anteprima file 3
    # --------------------------------------------------------
    st.subheader(
        "Anteprima 3 - Gruppi AD"
    )

    st.dataframe(
        pd.DataFrame(
            [row_gruppi_ad],
            columns=HEADER_UTENTE
        ),
        use_container_width=True
    )

    # --------------------------------------------------------
    # Anteprima file 4
    # --------------------------------------------------------
    st.subheader(
        "Anteprima 4 - Gruppi Entra ID"
    )

    st.text(
        csv_entraid
    )

    # --------------------------------------------------------
    # Anteprima file 5
    # --------------------------------------------------------
    st.subheader(
        "Anteprima 5 - Computer"
    )

    st.dataframe(
        pd.DataFrame(
            [row_computer],
            columns=HEADER_COMPUTER
        ),
        use_container_width=True
    )

    # --------------------------------------------------------
    # Download singoli file
    # --------------------------------------------------------
    st.download_button(
        "Scarica 1 - Utenza",
        data=csv_utente,
        file_name=filename_utente,
        mime="text/csv"
    )

    st.download_button(
        "Scarica 2 - Casella",
        data=csv_casella,
        file_name=filename_casella,
        mime="text/csv"
    )

    st.download_button(
        "Scarica 3 - Gruppi AD",
        data=csv_gruppi_ad,
        file_name=filename_gruppi_ad,
        mime="text/csv"
    )

    st.download_button(
        "Scarica 4 - Gruppi Entra ID",
        data=csv_entraid,
        file_name=filename_entraid,
        mime="text/csv"
    )

    st.download_button(
        "Scarica 5 - Computer",
        data=csv_computer,
        file_name=filename_computer,
        mime="text/csv"
    )

    # --------------------------------------------------------
    # Creazione bundle ZIP
    # --------------------------------------------------------
    zip_buffer = io.BytesIO()

    with zipfile.ZipFile(
        zip_buffer,
        "w",
        zipfile.ZIP_DEFLATED
    ) as zip_file:

        zip_file.writestr(
            filename_utente,
            csv_utente
        )

        zip_file.writestr(
            filename_casella,
            csv_casella
        )

        zip_file.writestr(
            filename_gruppi_ad,
            csv_gruppi_ad
        )

        zip_file.writestr(
            filename_entraid,
            csv_entraid
        )

        zip_file.writestr(
            filename_computer,
            csv_computer
        )

        zip_file.writestr(
            (
                f"Guida_Flusso_"
                f"{persona_nome_underscore}.txt"
            ),
            guida_flusso
        )

    zip_buffer.seek(0)

    bundle_name = (
        f"Provisioning_"
        f"{persona_nome_underscore}.zip"
    )

    st.download_button(
        "Scarica bundle completo ZIP",
        data=zip_buffer.getvalue(),
        file_name=bundle_name,
        mime="application/zip"
    )

    st.success(
        f"Bundle generato correttamente per '{sam}'."
    )
