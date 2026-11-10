import streamlit as st
import csv
import pandas as pd
import io
import unicodedata
import zipfile


# ------------------------------------------------------------
# Caricamento configurazione da Excel
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

    # --------------------------------------------------------
    # Normalizzazione colonne
    # --------------------------------------------------------
    if not ris.empty:

        ris = ris.rename(
            columns={
                column: str(column).strip()
                for column in ris.columns
            }
        )

        if "Tipo" in ris.columns and "TIPO" not in ris.columns:
            ris = ris.rename(
                columns={
                    "Tipo": "TIPO"
                }
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

    if (
        not ris.empty
        and required_columns.issubset(ris.columns)
    ):

        # ----------------------------------------------------
        # OU
        # ----------------------------------------------------
        ou_df = ris[
            ris["Section"] == "OU"
        ][
            [
                "Key/App",
                "Label/Gruppi/Value"
            ]
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

        # ----------------------------------------------------
        # Inserimento gruppi
        # ----------------------------------------------------
        gruppi_df = ris[
            ris["Section"] == "InserimentoGruppi"
        ][
            [
                "Key/App",
                "Label/Gruppi/Value"
            ]
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

        # ----------------------------------------------------
        # Defaults
        # ----------------------------------------------------
        defaults_df = ris[
            ris["Section"] == "Defaults"
        ][
            [
                "Key/App",
                "Label/Gruppi/Value"
            ]
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

    # --------------------------------------------------------
    # Estrazione delle righe Defaults con TIPO = PEL
    # Mantiene l'ordine originale del foglio Excel
    # --------------------------------------------------------
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
            [
                "Key/App",
                "Label/Gruppi/Value"
            ]
        ]

        for _, row in pel_df.iterrows():

            key = str(
                row["Key/App"]
            ).strip()

            value = str(
                row["Label/Gruppi/Value"]
            ).strip()

            if value:
                pel_items.append(
                    (key, value)
                )

    # --------------------------------------------------------
    # Estrazione organigramma
    # --------------------------------------------------------
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
    page_title="1.1 Nuova Risorsa Interna",
    layout="wide"
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
        "con il campo Section."
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


def unique_in_order(values):

    result = []

    for value in values:

        value = str(value).strip()

        if value and value not in result:
            result.append(value)

    return result


# ------------------------------------------------------------
# Liste di distribuzione
# ------------------------------------------------------------
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
# Gruppi AD letti dalle chiavi grp_o365_*
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


# ------------------------------------------------------------
# Percorso archivio
# ------------------------------------------------------------
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

    nome_normalizzato = normalize_name(
        nome
    )

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

    full_name = " ".join(
        parts
    )

    if esterno:
        full_name += " (esterno)"

    return full_name


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
    - viene mantenuto l'ordine della configurazione.
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
# Costruzione CSV
# ------------------------------------------------------------
def auto_quote(
    fields,
    quotechar='"',
    predicate=lambda value: " " in value
):

    """
    Racchiude tra doppi apici ogni valore
    che contiene almeno uno spazio.
    """

    output = []

    for field in fields:

        value = str(field)

        if predicate(value):

            output.append(
                f'{quotechar}{value}{quotechar}'
            )

        else:

            output.append(
                value
            )

    return output


def prepara_csv_standard(
    row,
    header
):

    """
    Mantiene la costruzione precedente:
    ogni campo contenente almeno uno spazio
    viene racchiuso tra doppi apici.
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

    writer.writerow(
        header
    )

    writer.writerow(
        quoted_row
    )

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

colonna_1, colonna_2 = st.columns(2)


with colonna_1:

    employee_id = st.text_input(
        "Matricola",
        defaults.get(
            "employee_id_default",
            ""
        )
    ).strip()

    cognome = (
        st.text_input(
            "Cognome"
        )
        .strip()
        .title()
    )

    secondo_cognome = (
        st.text_input(
            "Secondo Cognome"
        )
        .strip()
        .title()
    )

    nome = (
        st.text_input(
            "Nome"
        )
        .strip()
        .title()
    )

    secondo_nome = (
        st.text_input(
            "Secondo Nome"
        )
        .strip()
        .title()
    )

    codice_fiscale = st.text_input(
        "Codice Fiscale",
        ""
    ).strip()


with colonna_2:

    numero_telefono = (
        st.text_input(
            "Mobile (+39 già inserito)",
            ""
        )
        .replace(" ", "")
        .strip()
    )

    description = st.text_input(
        "PC (lascia vuoto per <PC>)",
        "<PC>"
    ).strip()

    data_operativa = st.text_input(
        "Data operatività (gg/mm/aaaa)",
        ""
    ).strip()

    mouse_id = st.text_input(
        "Mouse ID",
        ""
    ).strip()

    mouse_sn = st.text_input(
        "Mouse SN",
        ""
    ).strip()

    hub_id = st.text_input(
        "HUB Sandberg ID",
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


# ------------------------------------------------------------
# Resident e numero fisso
# ------------------------------------------------------------
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
    list(
        ou_options.values()
    )
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
    )[
        ou_values.index(
            label_ou
        )
    ]

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


# ------------------------------------------------------------
# Firma digitale
# ------------------------------------------------------------
firma_digitale = st.checkbox(
    "Firma digitale?"
)


# ------------------------------------------------------------
# Profilazione SM
# ------------------------------------------------------------
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
# Liste di distribuzione per tipologia utente
# ------------------------------------------------------------
if selected_key == "utenti_standard":

    dl_list = dl_standard

elif selected_key == "utenti_vip":

    dl_list = dl_vip

else:

    dl_list = []


# ------------------------------------------------------------
# Generazione flusso completo
# ------------------------------------------------------------
if st.button(
    "Genera CSV",
    type="primary"
):

    # --------------------------------------------------------
    # Validazioni
    # --------------------------------------------------------
    if not cognome or not nome:

        st.error(
            "Inserire almeno Cognome e Nome "
            "prima di generare i file."
        )

        st.stop()

    if not data_operativa:

        st.error(
            "Inserire la Data operatività."
        )

        st.stop()

    # --------------------------------------------------------
    # Dati principali
    # --------------------------------------------------------
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

    pc_asset = (
        description
        or "<PC>"
    )

    data_operativa_output = (
        data_operativa
        or "<data inizio>"
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

    # ========================================================
    # FILE 1 - UTENZA
    # ========================================================
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
        pc_asset,
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

    # ========================================================
    # FILE 2 - CASELLA
    # ========================================================
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

    # Eventuale profilazione SM
    if profilazione and sm_lines:

        casella_lines.append("")
        casella_lines.append(
            "Profilare su SM:"
        )

        for sm in sm_lines:

            casella_lines.append(
                f"- {sm}"
            )

    csv_casella = (
        "\n".join(
            casella_lines
        )
        + "\n"
    )

    # ========================================================
    # FILE 3 - PROFILAZIONE GRUPPI AD
    # ========================================================
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
    ] * len(
        HEADER_UTENTE
    )

    row_gruppi_ad[0] = sam

    inserimento_gruppo_index = (
        HEADER_UTENTE.index(
            "InserimentoGruppo"
        )
    )

    row_gruppi_ad[
        inserimento_gruppo_index
    ] = ";".join(
        gruppi_ad
    )

    csv_gruppi_ad = prepara_csv_standard(
        row_gruppi_ad,
        HEADER_UTENTE
    )

    # ========================================================
    # FILE 4 - PROFILAZIONE GRUPPI ENTRA ID
    # ========================================================
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

    if gruppi_entraid:

        entra_lines.extend(
            gruppi_entraid
        )

    else:

        entra_lines.append(
            "Nessun gruppo EntraID previsto."
        )

    csv_entraid = (
        "\n".join(
            entra_lines
        )
        + "\n"
    )

    # ========================================================
    # FILE 5 - COMPUTER
    # Da elaborare alla data operatività
    # ========================================================
    row_computer = [
        pc_asset,
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

    # ========================================================
    # FILE 6 - LISTE DI DISTRIBUZIONE
    # Da elaborare alla data operatività
    # ========================================================
    dl_lines = [
        (
            f"{sam}@consip.it "
            f"in data {data_operativa_output} "
            "inserire nelle DL:"
        )
    ]

    if dl_list:

        dl_lines.extend(
            dl_list
        )

    else:

        dl_lines.append(
            "Nessuna lista di distribuzione prevista."
        )

    csv_dl = (
        "\n".join(
            dl_lines
        )
        + "\n"
    )

    # ========================================================
    # NOMI FILE
    # ========================================================
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

    filename_dl = (
        f"6_{persona_nome_file}_DL.csv"
    )

    # ========================================================
    # 01 - GSI
    # ASSEGNAZIONE ASSET
    # ========================================================
    subject_gsi = (
        "01 - Assegnazione Asset - "
        f"{cn} → GSI"
    )

    message_gsi = (
        f"Per la nuova risorsa {cn} assegnare asset:\n"
        f"- PC Nome asset {pc_asset}\n"
        f"- mouse ID {mouse_id_output} "
        f"e SN {mouse_sn_output}\n"
        f"- HUB Sandberg ID {hub_id_output}"
    )

    # ========================================================
    # 02 - GIIT
    # Provisioning Utente + Computer + DL
    # ========================================================
    subject_giit = (
        "02 - Provisioning Utente "
        "(utenza/Casella/Profilazione) + "
        "Modifica AD Computer + "
        "Liste di Distribuzione - "
        f"{cn} → GIIT"
    )

    message_giit = (
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
        f"5. {filename_computer} per Aggiornamento Computer su AD con riferimenti utente, effettuare questa attività il giorno {data_operativa}."\n"
        f"6. {filename_dl} per l'inserimento "
        "nelle Liste di Distribuzione , effettuare questa attività il giorno {data_operativa}.\n"
    )

    # ========================================================
    # 03 - GHDW
    # Provisioning PC + Firma Asset
    # ========================================================
    subject_ghdw = (
        "03 - Provisioning PC e Firma Asset - "
        f"{cn} → GHDW"
    )

    message_ghdw = (
        "Ciao,\n"
        f"per {cn} ritirare dal 599 "
        "per CONFIGURAZIONE i seguenti asset:\n"
        f"- PC Nome asset {pc_asset}\n"
        f"- mouse ID {mouse_id_output} "
        f"e SN {mouse_sn_output}\n"
        f"- HUB Sandberg ID {hub_id_output}\n"
        "\n"
        f"La risorsa prende servizio "
        f"{data_operativa_output}.\n"
        "\n"
        "Attività da effettuare:\n"
        f"1. Configurare gli asset indicati.\n"
        f"2. Il giorno {data_operativa_output} "
        "far firmare gli asset ritirati dal 599 "
        f"per la consegna a {cn}."
    )

    # ========================================================
    # 04 - GA
    # Foto + SIC + Accesso intranet PSN
    # ========================================================
    subject_ga = (
        "04 - Caricamento Foto, Verifica SIC e "
        "Accesso intranet PSN - "
        f"{cn} → GA"
    )

    mail_utente = (
        f"{sam}@consip.it"
    )

    message_ga = (
        "Ciao,\n"
        "si richiedono le seguenti attività:\n"
        f"1. Caricare la foto in allegato per {cn}.\n"
        f"2. Verificare che per l'utente {cn} "
        "sia indicata sul SIC la seguente email: "
        f'"{mail_utente}".\n'
        "3. Accesso intranet PSN: attività "
        "momentaneamente esclusa dal flusso, "
        "perché attualmente gestita via email."
    )

    # ========================================================
    # 05 - MM
    # Firma digitale
    # ========================================================
    subject_mm = (
        "05 - Firma digitale - "
        f"{cn} → MM"
    )

    if firma_digitale:

        message_mm = (
            f"Assegnare Firma Digitale a {cn}."
        )

    else:

        message_mm = (
            "Firma digitale non prevista "
            "per questa risorsa."
        )

    # ========================================================
    # GUIDA COMPLETA
    # Raggruppata per area di competenza
    # ========================================================
    guida_flusso_lines = [
        subject_gsi,
        message_gsi,

        subject_giit,
        message_giit,

        subject_ghdw,
        message_ghdw,

        subject_ga,
        message_ga,

        subject_mm,
        message_mm
    ]

    guida_flusso = "\n\n".join(
        guida_flusso_lines
    )

    filename_guida = (
        f"Guida_Flusso_"
        f"{persona_nome_underscore}.txt"
    )

    # ========================================================
    # CREAZIONE BUNDLE ZIP
    # ========================================================
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
            filename_dl,
            csv_dl
        )

        zip_file.writestr(
            filename_guida,
            guida_flusso
        )

    zip_buffer.seek(0)

    bundle_name = (
        f"Provisioning_"
        f"{persona_nome_underscore}.zip"
    )

    # ========================================================
    # VISUALIZZAZIONE DEL FLUSSO
    # ========================================================

    # ========================================================
    # 01 - GSI
    # ========================================================
    st.subheader(
        subject_gsi
    )

    st.text(
        message_gsi
    )

    st.divider()

    # ========================================================
    # 02 - GIIT
    # ========================================================
    st.subheader(
        subject_giit
    )

    st.text(
        message_giit
    )

    st.markdown(
        "**File di competenza GIIT:**"
    )

    # --------------------------------------------------------
    # File 1 - Utenza
    # --------------------------------------------------------
    st.markdown(
        f"**1. {filename_utente}**"
    )

    st.dataframe(
        pd.DataFrame(
            [row_utente],
            columns=HEADER_UTENTE
        ),
        use_container_width=True
    )

    st.download_button(
        "Scarica 1 - Utenza",
        data=csv_utente,
        file_name=filename_utente,
        mime="text/csv",
        key="download_utente"
    )

    # --------------------------------------------------------
    # File 2 - Casella
    # --------------------------------------------------------
    st.markdown(
        f"**2. {filename_casella}**"
    )

    st.text(
        csv_casella
    )

    st.download_button(
        "Scarica 2 - Casella",
        data=csv_casella,
        file_name=filename_casella,
        mime="text/csv",
        key="download_casella"
    )

    # --------------------------------------------------------
    # File 3 - Gruppi AD
    # --------------------------------------------------------
    st.markdown(
        f"**3. {filename_gruppi_ad}**"
    )

    st.dataframe(
        pd.DataFrame(
            [row_gruppi_ad],
            columns=HEADER_UTENTE
        ),
        use_container_width=True
    )

    st.download_button(
        "Scarica 3 - Gruppi AD",
        data=csv_gruppi_ad,
        file_name=filename_gruppi_ad,
        mime="text/csv",
        key="download_gruppi_ad"
    )

    # --------------------------------------------------------
    # File 4 - Gruppi Entra ID
    # --------------------------------------------------------
    st.markdown(
        f"**4. {filename_entraid}**"
    )

    st.text(
        csv_entraid
    )

    st.download_button(
        "Scarica 4 - Gruppi Entra ID",
        data=csv_entraid,
        file_name=filename_entraid,
        mime="text/csv",
        key="download_entraid"
    )

    # --------------------------------------------------------
    # File 5 - Computer
    # --------------------------------------------------------
    st.markdown(
        f"**5. {filename_computer}**"
    )

    st.info(
        f"Il file deve essere elaborato il giorno "
        f"{data_operativa_output}."
    )

    st.dataframe(
        pd.DataFrame(
            [row_computer],
            columns=HEADER_COMPUTER
        ),
        use_container_width=True
    )

    st.download_button(
        "Scarica 5 - Computer",
        data=csv_computer,
        file_name=filename_computer,
        mime="text/csv",
        key="download_computer"
    )

    # --------------------------------------------------------
    # File 6 - Liste di distribuzione
    # --------------------------------------------------------
    st.markdown(
        f"**6. {filename_dl}**"
    )

    st.info(
        f"Il file deve essere elaborato il giorno "
        f"{data_operativa_output}."
    )

    st.text(
        csv_dl
    )

    st.download_button(
        "Scarica 6 - Liste di Distribuzione",
        data=csv_dl,
        file_name=filename_dl,
        mime="text/csv",
        key="download_dl"
    )

    # --------------------------------------------------------
    # Bundle completo
    # --------------------------------------------------------
    st.markdown(
        "**Bundle completo del provisioning:**"
    )

    st.download_button(
        "📦 Scarica bundle completo ZIP",
        data=zip_buffer.getvalue(),
        file_name=bundle_name,
        mime="application/zip",
        key="download_bundle"
    )

    st.divider()

    # ========================================================
    # 03 - GHDW
    # ========================================================
    st.subheader(
        subject_ghdw
    )

    st.text(
        message_ghdw
    )

    st.divider()

    # ========================================================
    # 04 - GA
    # ========================================================
    st.subheader(
        subject_ga
    )

    st.text(
        message_ga
    )

    st.divider()

    # ========================================================
    # 05 - MM
    # ========================================================
    st.subheader(
        subject_mm
    )

    if firma_digitale:

        st.text(
            message_mm
        )

    else:

        st.info(
            message_mm
        )

    st.success(
        f"Flusso completo generato correttamente per '{sam}'."
    )
