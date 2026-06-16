# -----------------------------------------------------------------------------
# INSTRUKCJA URUCHOMIENIA:
# 1. Wymagane biblioteki: pip install openpyxl pandas streamlit xlsxwriter
# 2. Uruchomienie (dla dużych plików):
#    streamlit run ranking.py --server.maxUploadSize=2000
# -----------------------------------------------------------------------------

import streamlit as st
import pandas as pd
import io
import gc


# Ustawienie szerokiego układu strony
st.set_page_config(page_title="Śladem pieczątek", layout="wide")

def main():

    # -------------------------------------------------------------------------
    # CUSTOM CSS - Modern & Minimalist
    # -------------------------------------------------------------------------
    st.markdown("""
        <style>
            /* Global styles */
            * {
                text-transform: uppercase;
            }
            .main {
                background-color: #f9fafb;
            }
            .stButton>button {
                border-radius: 8px;
                border: none;
                padding: 0.5rem 1rem;
                background-color: #374151;
                color: white;
                transition: all 0.3s ease;
                font-weight: 700;
                letter-spacing: 0.05em;
            }
            .stButton>button:hover {
                background-color: #111827;
                box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);
                transform: translateY(-1px);
            }
            .stTextInput>div>div>input {
                border-radius: 8px;
                background-color: #ffffff;
                border: 1px solid #d1d5db;
            }
            .stSelectbox>div>div>div {
                border-radius: 8px;
                background-color: #ffffff;
            }
            /* Sidebar styling */
            [data-testid="stSidebar"] {
                background-color: #f3f4f6;
                border-right: 1px solid #d1d5db;
            }
            .sidebar .sidebar-content {
                padding: 2rem 1rem;
            }
            /* Tabs styling - REMOVED since tabs are gone */
            
            /* Dataframe styling */
            .stDataFrame {
                border-radius: 12px;
                overflow: hidden;
                border: 1px solid #d1d5db;
                background-color: #ffffff;
            }
            /* Scrollable container styling - ZMNIEJSZONY PADDING */
            [data-testid="stVerticalBlock"] > div > div > div > [data-testid="stVerticalBlock"] {
                padding: 0.1rem !important;
            }
            /* Moved button below the field */
            [data-testid="stFileUploader"] section {
                border: none !important;
                background-color: transparent !important;
                display: flex !important;
                flex-direction: column !important;
                align-items: center !important;
                padding: 0 !important;
            }
            /* The 'field' area (drag and drop zone) */
            [data-testid="stFileUploader"] section > div {
                border: 2px dashed #374151;
                border-radius: 12px;
                background-color: #ffffff;
                padding: 2.5rem 1rem !important;
                width: 100%;
                text-align: center;
                transition: all 0.3s ease;
            }
            /* Hide original text elements more aggressively */
            [data-testid="stFileUploader"] section > div,
            [data-testid="stFileUploader"] section small,
            [data-testid="stFileUploader"] section span {
                font-size: 0 !important;
                line-height: 0 !important;
                color: transparent !important;
                visibility: hidden !important;
            }
            /* Show ONLY our translated text */
            [data-testid="stFileUploader"] section > div::before {
                content: "PRZECIĄGNIJ I UPUŚĆ PLIKI TUTAJ";
                font-size: 1rem !important;
                font-weight: 700 !important;
                color: #374151 !important;
                visibility: visible !important;
                line-height: 1.5 !important;
                display: block !important;
            }
            [data-testid="stFileUploader"] section > div:hover {
                border-color: #111827;
                background-color: #f9fafb;
            }
            /* The button below the zone */
            [data-testid="stFileUploader"] button {
                margin-top: 1rem !important;
                background-color: #374151 !important;
                color: white !important;
                border-radius: 6px !important;
                width: auto !important;
                visibility: visible !important;
                font-size: 0 !important; /* Hide original text */
            }
            [data-testid="stFileUploader"] button * {
                font-size: 0 !important;
                visibility: hidden !important;
            }
            [data-testid="stFileUploader"] button::before {
                content: "PRZEGLĄDAJ PLIKI";
                font-size: 0.8rem !important;
                font-weight: 700 !important;
                padding: 0.5rem 1.5rem !important;
                visibility: visible !important;
                line-height: 1.5 !important;
                display: block !important;
                color: white !important;
            }
            [data-testid="stFileUploader"] button:hover {
                background-color: #111827 !important;
            }
            /* Hide any remaining default uploader parts */
            [data-testid="stFileUploader"] ul {
                display: none !important;
            }
            [data-testid="stFileUploaderPagination"] {
                display: none !important;
            }
            [data-testid="stFileUploader"] [data-testid="stMarkdownContainer"] {
                /* Optional: hide other status text if needed */
            }
            /* Consistent Sidebar Labels */
            [data-testid="stSidebar"] label, 
            [data-testid="stSidebar"] .st-at,
            [data-testid="stSidebar"] .st-ae,
            [data-testid="stSidebar"] [data-testid="stMarkdownContainer"] p {
                font-size: 14px !important;
                font-weight: 700 !important;
            }
            /* Move Logo to top corner */
            [data-testid="stSidebarNav"] {
                display: none;
            }
            [data-testid="stSidebar"] [data-testid="stVerticalBlock"] {
                padding-top: 0rem !important;
                gap: 0rem !important;
            }
            /* Maximize main content area */
            [data-testid="stMainBlockContainer"] {
                padding-top: 3.5rem !important;
                padding-right: 1rem !important;
                padding-left: 1rem !important;
                padding-bottom: 1rem !important;
                max-width: 98% !important;
            }
            /* ====== KOMPAKTOWANIE UKŁADU ====== */
            /* Globalne: zero gap między elementami pionowymi */
            [data-testid="stVerticalBlock"] {
                gap: 0rem !important;
            }
            /* Padding kontenerów z ramką (border=True) */
            [data-testid="stVerticalBlock"] > div > div > div > [data-testid="stVerticalBlock"] {
                padding: 0.1rem !important;
            }
            /* Zmniejszenie marginesów wrappera każdego widgetu */
            [data-testid="stElementContainer"] {
                margin-bottom: 0rem !important;
                margin-top: 0rem !important;
                padding-bottom: 0rem !important;
                padding-top: 0rem !important;
            }
            /* Zmniejszenie etykiet widgetów */
            [data-testid="stWidgetLabel"] {
                margin-bottom: 0rem !important;
                min-height: 0rem !important;
                line-height: 1.2 !important;
            }
            /* Kolumny poziome */
            [data-testid="stHorizontalBlock"] {
                gap: 0.25rem !important;
                align-items: flex-start !important;
            }
            [data-testid="column"] {
                padding-left: 0.15rem !important;
                padding-right: 0.15rem !important;
            }
            /* Usuń padding z bloków wewnątrz kolumn */
            [data-testid="column"] [data-testid="stVerticalBlock"] > div {
                padding-top: 0 !important;
                padding-bottom: 0 !important;
            }
            /* Zmniejsz radio */
            [data-testid="stRadio"] > div {
                gap: 0.1rem !important;
                padding: 0 !important;
            }
            [data-testid="stRadio"] label {
                padding: 0.1rem 0 !important;
            }
            /* Sidebar */
            [data-testid="stSidebar"] [data-testid="stVerticalBlock"] {
                gap: 0.2rem !important;
            }
            [data-testid="stSidebar"] [data-testid="stElementContainer"] {
                margin-bottom: 0rem !important;
                padding-bottom: 0rem !important;
            }
            /* Fix overlapping in sidebar */
            [data-testid="stSidebar"] .stSelectbox {
                margin-top: 0.5rem !important;
            }
            /* Italic placeholders & Lowercase for inputs */
            @keyframes spin {
                0% { transform: rotate(0deg); }
                100% { transform: rotate(360deg); }
            }
            /* Styl dla małego przycisku pobierania */
            .small-btn div[data-testid="stButton"] button, 
            .small-btn div.stDownloadButton button {
                padding: 0.25rem 0.75rem !important;
                font-size: 0.7rem !important;
                min-height: unset !important;
                height: auto !important;
                margin-top: 0.2rem !important;
            }
            ::placeholder {
                font-style: italic;
                text-transform: none !important;
                opacity: 0.7;
                font-size: 0.75rem !important;
            }
            input {
                text-transform: none !important;
                font-size: 0.75rem !important;
            }
            /* Zmniejszenie czcionki opcji radio (poprawione: powiększone o 10% względem poprzedniej zmiany) */
            div[data-testid="stRadio"] label p {
                font-size: 0.77rem !important;
            }
            div[data-testid="stDialog"] div[role="dialog"] {
                width: 95vw !important;
                max-width: none !important;
            }
            div[data-testid="stDialog"] div[role="dialog"] > div {
                max-width: none !important;
                width: 100% !important;
            }
        </style>
    """, unsafe_allow_html=True)

    # =========================================================================
    # FUNKCJE DIALOGOWE
    # =========================================================================
    @st.dialog("📊 PODSUMOWANIE WYSZUKIWANIA", width="large")
    def show_summary(df):
        st.write("Statystyki dla znalezionych rekordów:")
        
        # 1. Organy
        st.subheader("🏢 ORGANY WYDAJĄCE DECYZJE")
        cols_org = ['nazwa_organu', 'adres_organu']
        available_org = [c for c in cols_org if c in df.columns]
        if len(available_org) == 2:
            summary_org = df.groupby(available_org).size().reset_index(name='LICZBA WYSTĄPIEŃ')
            summary_org = summary_org.sort_values('LICZBA WYSTĄPIEŃ', ascending=False)
            st.dataframe(summary_org, hide_index=True, width="stretch")
        else:
            missing = set(cols_org) - set(available_org)
            st.info(f"Brak kolumn do podsumowania organów: {', '.join(missing)}")

        # 2. Inwestycje
        st.subheader("🏗️ RODZAJE INWESTYCJI")
        cols_inv = ['rodzaj_inwestycji', 'kategoria']
        available_inv = [c for c in cols_inv if c in df.columns]
        if len(available_inv) == 2:
            summary_inv = df.groupby(available_inv).size().reset_index(name='LICZBA WYSTĄPIEŃ')
            summary_inv = summary_inv.sort_values('LICZBA WYSTĄPIEŃ', ascending=False)
            st.dataframe(summary_inv, hide_index=True, width="stretch")
        else:
            missing = set(cols_inv) - set(available_inv)
            st.info(f"Brak kolumn do podsumowania inwestycji: {', '.join(missing)}")
        
        if st.button("ZAMKNIJ"):
            st.rerun()

    @st.dialog("ZASADY I PRYWATNOŚĆ", width="large")
    def show_terms():
        st.markdown("""
        ### Zasady użytkowania
        Aplikacja 'Śladem pieczątek' służy wyłącznie do wspomagania analizy statystycznej danych z Głównego Urzędu Nadzoru Budowlanego.
        
        ### Własność intelektualna
        Wszelkie rozwiązania techniczne, kod źródłowy oraz układ interfejsu stanowią własność intelektualną autora. **Kopiowanie, powielanie lub redystrybucja narzędzia bez zgody jest zabroniona.**
        
        ### Odpowiedzialność
        Narzędzie jest udostępnione w stanie takim, w jakim jest. Użytkownik korzysta z aplikacji na **własną odpowiedzialność**. Autor nie ponosi odpowiedzialności za rzetelność danych źródłowych (CSV) ani za skutki decyzji podjętych na podstawie wygenerowanych zestawień.
        
        ### Prywatność i dane
        Aplikacja **nie zbiera, nie przechowuje i nie przesyła dalej** żadnych danych wgranych przez użytkownika. Wszystkie operacje odbywają się lokalnie w ramach bieżącej sesji. Po zamknięciu przeglądarki lub odświeżeniu strony, wgrane dane są całkowicie usuwane z pamięci tymczasowej serwera.
        """)
        
        st.write("---")
        if st.button("AKCEPTUJĘ", use_container_width=True):
            st.session_state.terms_dismissed = True
            st.rerun()


    # Inicjalizacja zmiennych stanu sesji (zawsze na początku, aby uniknąć błędów dostępu)
    if "ranking_df" not in st.session_state:
        st.session_state.ranking_df = None
    if "search_df" not in st.session_state:
        st.session_state.search_df = None
    if "terms_accepted" not in st.session_state:
        st.session_state.terms_accepted = False
    if "terms_dismissed" not in st.session_state:
        st.session_state.terms_dismissed = False

    # Automatyczne wyświetlanie zasad przy pierwszym wejściu
    if not st.session_state.terms_accepted and not st.session_state.terms_dismissed:
        show_terms()


    # -------------------------------------------------------------------------
    # PANEL BOCZNY (SIDEBAR) - Konfiguracja
    # -------------------------------------------------------------------------
    with st.sidebar:
        st.markdown("<h2 style='text-align: center; color: #374151; margin-top: -1rem; font-size: 1.63rem; font-weight: 900;'>ŚLADEM <span style='color: #fbbf24; font-size: 1.2em; vertical-align: middle;'>●</span> PIECZĄTEK</h2>", unsafe_allow_html=True)
        
        
        st.markdown("---")
        uploaded_files = st.file_uploader(
            "WGRAJ PLIKI CSV LUB TXT", 
            type=['csv', 'txt'], 
            accept_multiple_files=True,
            label_visibility="collapsed"
        )
        
        if uploaded_files:
            for file in uploaded_files:
                st.markdown(f"<span style='color: #374151; font-weight: 700;'>✔ {file.name}</span>", unsafe_allow_html=True)
        
        st.markdown("---")
        csv_sep = ";"
        
        grouping_choice = st.selectbox(
            "GRUPUJ DUPLIKATY", 
            options=["TAK", "NIE"],
            index=0,
            help="GRUPUJ DECYZJE O TYM SAMYM NUMERZE URZĘDOWYM, CZYLI NP. DOTYCZĄCE ZABUDOWY SZEREGOWEJ LUB BLIŹNIACZEJ"
        )
        use_grouping = grouping_choice == "TAK"
        
        st.markdown("---")
        if st.button("ZASADY I PRYWATNOŚĆ", key="btn_terms", use_container_width=True):
            show_terms()

    # Hardcoded column names
    csv_name_col = "projektant_imie"
    csv_surname_col = "projektant_nazwisko"
    csv_license_col = "projektant_numer_uprawnien"
    csv_date_col = "data_wydania_decyzji"
    csv_unique_col = "numer_urzad"

    # -------------------------------------------------------------------------
    # GŁÓWNY PANEL - DYNAMIKA (WITAJ LUB WIDOK DANYCH)
    # -------------------------------------------------------------------------
    if not uploaded_files:
        st.session_state.ranking_df = None
        st.session_state.search_df = None
        with st.container(height=850, border=True):
            st.markdown("""
                <div style='display: flex; flex-direction: column; height: 750px; align-items: center; justify-content: center; text-align: center; color: #6b7280;'>
                    <div style='font-weight: 700; font-size: 18px; margin-bottom: 1.5rem; color: #374151;'>ZACZNIJ OD WGRANIA PLIKÓW CSV PO LEWEJ STRONIE.</div>
                    <div style='font-size: 13px; margin-top: 2rem; padding: 1.5rem; border-top: 1px solid #e5e7eb;'>
                        DANE DO ANALIZY MOŻNA POBRAĆ ZE STRONY: <br><br>
                        <a href='https://wyszukiwarka.gunb.gov.pl/pobranie.html' target='_blank' 
                           style='color: #111827; font-weight: 700; text-decoration: underline; font-size: 15px;'>
                           wyszukiwarka.gunb.gov.pl/pobranie.html
                        </a>
                    </div>
                </div>
            """, unsafe_allow_html=True)
    else:
        col_left, col_right = st.columns(2)

        # =========================================================================
        # LEWA KOLUMNA: RANKING
        # =========================================================================
        with col_left:
            with st.container(height=850, border=True):
                # Miejsce na komunikaty statusowe (pełna szerokość)
                status_ranking = st.empty()
                
                c1, c2 = st.columns(2)
                with c1:
                    if st.button("GENERUJ RANKING", key="btn_ranking"):
                        with st.spinner("⚙️ TRWA PRZETWARZANIE DANYCH..."):
                            st.session_state.ranking_df = analyze_statistics(
                                uploaded_files, 
                                csv_name_col, csv_surname_col, csv_license_col, csv_date_col, csv_unique_col, 
                                csv_sep, use_grouping,
                                status_container=status_ranking
                            )
                
                with c2:
                    if st.session_state.ranking_df is not None:
                        import io
                        buffer = io.BytesIO()
                        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                            df_export = st.session_state.ranking_df.reset_index()
                            df_export.to_excel(writer, index=False, sheet_name='Ranking')
                            worksheet = writer.sheets['Ranking']
                            for idx, col in enumerate(df_export.columns):
                                max_len = min(max(int(df_export[col].astype(str).str.len().fillna(0).max()), len(str(col))) + 2, 50)
                                worksheet.set_column(idx, idx, max_len)
                        st.download_button("EXPORT DANYCH DO EXCELA", buffer.getvalue(), "ranking.xlsx", "application/vnd.ms-excel", width="stretch")
                
                # Display ranking if it exists in session state
                if st.session_state.ranking_df is not None:
                    pivot = st.session_state.ranking_df
                    
                    # Dynamiczne filtrowanie po nazwisku
                    search_rank = st.text_input("🔍 FILTRUJ RANKING PO NAZWISKU", placeholder="wpisz nazwisko...", label_visibility="collapsed")
                    
                    if search_rank:
                        # Filtrowanie indeksu MultiIndex (poziom S_Nazwisko)
                        filtered_pivot = pivot[pivot.index.get_level_values('S_Nazwisko').str.contains(search_rank, case=False, na=False)]
                    else:
                        filtered_pivot = pivot
                    
                    st.dataframe(filtered_pivot.head(1000), width="stretch", height=700)

        # =========================================================================
        # PRAWA KOLUMNA: WYSZUKIWARKA
        # =========================================================================
        with col_right:
            with st.container(height=850, border=True):
                # Miejsce na komunikaty statusowe wyszukiwarki (pełna szerokość)
                status_search = st.empty()
                
                # FORMULARZ WYSZUKIWANIA
                with st.form("search_form", border=False):
                    # Pola tekstowe
                    col_nazwisko, col_imie, col_uprawnienia, col_rok = st.columns([1, 1, 1, 0.7])
                    with col_nazwisko:
                        search_surname = st.text_input("NAZWISKO", placeholder="nazwisko", label_visibility="collapsed")
                    with col_imie:
                        search_name = st.text_input("IMIĘ", placeholder="imię", label_visibility="collapsed")
                    with col_uprawnienia:
                        search_license = st.text_input("NR UPRAWNIEŃ", placeholder="nr uprawnień", label_visibility="collapsed")
                    with col_rok:
                        search_year = st.text_input("ROK", placeholder="rok", label_visibility="collapsed")
                    
                    # Tryb szukania
                    search_mode = st.radio("TRYB SZUKANIA:", ["DOKŁADNE DOPASOWANIE", "ZAWIERA FRAZĘ"], horizontal=True, label_visibility="collapsed")
                    
                    # Przycisk wyszukiwania (Submit)
                    submitted = st.form_submit_button("SZUKAJ W BAZIE", use_container_width=True)
                
                if submitted:
                    if not uploaded_files:
                        st.error("BRAK WGRANYCH PLIKÓW CSV!")
                    elif not (search_name or search_surname or search_license or search_year):
                        st.warning("PODAJ CO NAJMNIEJ JEDNO KRYTERIUM WYSZUKIWANIA.")
                    else:
                        with st.spinner("⚙️ TRWA PRZESZUKIWANIE BAZY..."):
                            st.session_state.search_df = perform_detailed_search(
                                uploaded_files, 
                                search_name, search_surname, search_license, search_year,
                                csv_name_col, csv_surname_col, csv_license_col, csv_date_col, csv_unique_col,
                                csv_sep,
                                search_mode, use_grouping,
                                status_container=status_search
                            )
                
                # Dodatkowe akcje (Export/Podsumowanie) widoczne po wyszukaniu
                if st.session_state.search_df is not None and not st.session_state.search_df.empty:
                    c1, c2 = st.columns(2)
                    with c1:
                        res_df = st.session_state.search_df
                        import io
                        buffer = io.BytesIO()
                        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                            res_df.to_excel(writer, index=False, sheet_name='Wyniki')
                            worksheet = writer.sheets['Wyniki']
                            for idx, col in enumerate(res_df.columns):
                                max_len = min(max(int(res_df[col].astype(str).str.len().fillna(0).max()), len(str(col))) + 2, 50)
                                worksheet.set_column(idx, idx, max_len)
                        st.download_button("EXPORT DANYCH DO EXCELA", buffer.getvalue(), "wyniki_szukania.xlsx", "application/vnd.ms-excel", use_container_width=True)
                    
                    with c2:
                        if st.button("PODSUMOWANIE", key="btn_summary", use_container_width=True):
                            show_summary(st.session_state.search_df)

                # Wyświetlanie wyników
                if st.session_state.search_df is not None:
                    res_df = st.session_state.search_df
                    if not res_df.empty:
                        st.dataframe(res_df, width="stretch", height=600)
                    else:
                        st.warning("BRAK WYNIKÓW DLA PODANYCH KRYTERIÓW.")

    # -------------------------------------------------------------------------
    # STAŁY KOMUNIKAT OSTRZEGAWCZY
    # -------------------------------------------------------------------------
    st.markdown("""
        <div style="
            background-color: #fff1f2;
            border: 2px solid #dc2626;
            border-radius: 10px;
            padding: 1rem 1.5rem;
            margin-top: 1rem;
        ">
            <p style="
                color: #dc2626;
                font-weight: 700;
                font-size: 0.82rem;
                margin: 0;
                line-height: 1.7;
                letter-spacing: 0.03em;
            ">
                ⚠️ WAŻNE — WERYFIKACJA WYNIKÓW:<br>
                Każdy rekord uzyskany za pomocą tej aplikacji należy bezwzględnie zweryfikować w dodatkowych, niezależnych źródłach 
                przed wyciągnięciem jakichkolwiek wniosków. Aplikacja nie została przetestowana dla wszystkich możliwych wariantów 
                danych i może zawierać błędy. Wyniki jej działania mogą dotyczyć osób fizycznych — ich nieprawidłowa interpretacja 
                lub upublicznienie bez weryfikacji może poważnie naruszyć czyjąś reputację. Użytkownik ponosi pełną odpowiedzialność 
                za sposób wykorzystania uzyskanych informacji.
            </p>
        </div>
    """, unsafe_allow_html=True)


# -----------------------------------------------------------------------------
# FUNKCJE LOGIKI
# -----------------------------------------------------------------------------

def analyze_statistics(files, col_imie, col_nazwisko, col_upr, col_data, col_unique, sep, use_grouping, status_container=None):
    """Generuje ranking z uwzględnieniem grupowania po numerze sprawy"""
    progress = st.progress(0)
    
    # Helper to show messages in container or directly
    def show_msg(type, text):
        if status_container:
            if type == 'success': status_container.success(text)
            elif type == 'warning': status_container.warning(text)
            elif type == 'error': status_container.error(text)
        else:
            if type == 'success': st.success(text)
            elif type == 'warning': st.warning(text)
            elif type == 'error': st.error(text)
    data_frames = []
    total_initial_rows = 0
    
    cols_to_use = [col_imie, col_nazwisko, col_upr, col_data]
    if use_grouping:
        cols_to_use.append(col_unique)
    
    for i, file in enumerate(files):
        try:
            file.seek(0)
            # Próba wczytania z różnymi kodowaniami
            df = None
            for enc in ['utf-8', 'utf-8-sig', 'cp1250', 'iso-8859-2']:
                try:
                    file.seek(0)
                    df = pd.read_csv(file, sep=sep, usecols=cols_to_use, dtype=str, on_bad_lines='skip', encoding=enc)
                    break
                except (UnicodeDecodeError, ValueError):
                    continue
            
            if df is None:
                # Jeśli nadal brak, spróbujmy bez usecols, żeby zobaczyć co jest w środku i wyrzucić błąd
                file.seek(0)
                df_check = pd.read_csv(file, sep=sep, nrows=0)
                missing = [c for c in cols_to_use if c not in df_check.columns]
                if missing:
                    st.warning(f"⚠️ Plik `{file.name}`: Brak kolumn: {missing}. Sprawdź separator i nazwy kolumn.")
                else:
                    st.warning(f"⚠️ Plik `{file.name}`: Nie udało się dopasować kodowania lub wystąpił inny błąd.")
                continue

            df['S_Imie'] = df[col_imie].str.strip().str.title()
            df['S_Nazwisko'] = df[col_nazwisko].str.strip().str.title()
            df['S_Upr'] = df[col_upr].str.strip()
            
            df['temp_date'] = pd.to_datetime(df[col_data], errors='coerce', dayfirst=False)
            df['Rok'] = df['temp_date'].dt.year.fillna(0).astype(int)
            
            total_initial_rows += len(df)
            
            if use_grouping:
                df['S_Unique'] = df[col_unique].str.strip()
                # Optymalizacja: deduplikacja już na poziomie pojedynczego pliku
                df = df.drop_duplicates(subset=['S_Nazwisko', 'S_Imie', 'S_Upr', 'S_Unique'])
            
            # Zostawiamy tylko to co niezbędne do rankingu
            keep_cols = ['S_Nazwisko', 'S_Imie', 'S_Upr', 'Rok']
            if use_grouping: keep_cols.append('S_Unique')
            df = df[keep_cols]
            
            data_frames.append(df)
            del df
        except Exception as e:
            st.warning(f"⚠️ Błąd w pliku `{file.name}`: {e}")
        progress.progress((i + 1) / len(files))

    if not data_frames:
        st.error("Błąd wczytywania danych.")
        return

    full_df = pd.concat(data_frames, ignore_index=True)
    
    # Grupowanie (Deduplikacja)
    if use_grouping:
        full_df = full_df.drop_duplicates(subset=['S_Nazwisko', 'S_Imie', 'S_Upr', 'S_Unique'])
        show_msg('success', f"Zredukowano z {total_initial_rows:,} do {len(full_df):,} unikalnych wpisów.")

    # Pivot
    pivot = full_df.pivot_table(index=['S_Nazwisko', 'S_Imie', 'S_Upr'], columns='Rok', aggfunc='size', fill_value=0)
    
    # Konwersja nazw kolumn na string (naprawia błąd "mixed type" w Streamlit)
    pivot.columns = pivot.columns.astype(str)
    pivot['SUMA'] = pivot.sum(axis=1)
    pivot = pivot.sort_values('SUMA', ascending=False)
    
    # Reorganizacja kolumn: SUMA na początku, lata malejąco, rok nieznany na końcu
    years = [c for c in pivot.columns if c != 'SUMA' and c != 0]
    years.sort(reverse=True)
    
    final_cols = ['SUMA'] + years
    if 0 in pivot.columns:
        final_cols.append(0)
    
    pivot = pivot[final_cols]
    
    # Przeniesienie zmiany nazwy tutaj, przed zwróceniem
    if 0 in pivot.columns:
        pivot = pivot.rename(columns={0: 'ROK NIEZNANY'})

    return pivot


def perform_detailed_search(files, s_name, s_surname, s_license, s_year, col_imie, col_nazwisko, col_upr, col_date, col_unique, sep, mode, deduplicate, status_container=None):
    """Wyszukiwanie szczegółowe z opcją deduplikacji"""
    results = []
    progress = st.progress(0)
    
    # Helper to show messages in container or directly
    def show_msg(type, text):
        if status_container:
            if type == 'success': status_container.success(text)
            elif type == 'warning': status_container.warning(text)
            elif type == 'error': status_container.error(text)
        else:
            if type == 'success': st.success(text)
            elif type == 'warning': st.warning(text)
            elif type == 'error': st.error(text)
    
    target_name = s_name.strip().lower() if s_name else ""
    target_surname = s_surname.strip().lower() if s_surname else ""
    target_license = s_license.strip().lower() if s_license else ""
    target_year = s_year.strip() if s_year else ""

    for i, file in enumerate(files):
        try:
            file.seek(0)
            # Próba wczytania z różnymi kodowaniami
            df = None
            for enc in ['utf-8', 'utf-8-sig', 'cp1250', 'iso-8859-2']:
                try:
                    file.seek(0)
                    df = pd.read_csv(file, sep=sep, dtype=str, on_bad_lines='skip', encoding=enc)
                    break
                except UnicodeDecodeError:
                    continue
            
            if df is None:
                st.warning(f"⚠️ Plik `{file.name}`: Nie udało się rozpoznać kodowania pliku.")
                continue

            # Sprawdzenie czy wymagane kolumny istnieją
            for c in [col_imie, col_nazwisko, col_upr]:
                if c not in df.columns:
                    st.warning(f"⚠️ Plik `{file.name}`: Brak kolumny `{c}`. Sprawdź separator.")
                    break
            else:
                temp_imie = df[col_imie].astype(str).str.strip().str.lower()
                temp_nazwisko = df[col_nazwisko].astype(str).str.strip().str.lower()
                temp_upr = df[col_upr].astype(str).str.strip().str.lower()
                
                condition = pd.Series([True] * len(df))
                
                if target_surname:
                    if mode == "DOKŁADNE DOPASOWANIE": condition &= (temp_nazwisko == target_surname)
                    else: condition &= (temp_nazwisko.str.contains(target_surname, na=False))
                
                if target_name:
                    if mode == "DOKŁADNE DOPASOWANIE": condition &= (temp_imie == target_name)
                    else: condition &= (temp_imie.str.contains(target_name, na=False))
                    
                if target_license:
                     if mode == "DOKŁADNE DOPASOWANIE": condition &= (temp_upr == target_license)
                     else: condition &= (temp_upr.str.contains(target_license, na=False))
                
                if target_year:
                    # Ekstrakcja roku z daty
                    extracted_year = df[col_date].astype(str).str.extract(r'(\d{4})')[0]
                    condition &= (extracted_year == target_year)
                
                matching = df[condition]
                
                if not matching.empty:
                    matching.insert(0, 'Plik_Źródłowy', file.name)
                    results.append(matching)
                
        except Exception as e:
            st.warning(f"⚠️ Błąd w pliku `{file.name}`: {e}")
        
        # Agresywne czyszczenie pamięci po każdym pliku
        if 'df' in locals(): del df
        progress.progress((i+1)/len(files))

    if results:
        final_df = pd.concat(results, ignore_index=True)
        final_df.index = final_df.index + 1
        
        # --- NOWA SEKCJA DEDUPLIKACJI W WYSZUKIWARCE ---
        if deduplicate:
            if col_unique in final_df.columns:
                before = len(final_df)
                # Usuwamy duplikaty biorąc pod uwagę: Numer Urzędu ORAZ Dane Osobowe.
                # Dzięki temu nie usuniemy dwóch RÓŻNYCH osób pracujących przy tej samej sprawie.
                final_df = final_df.drop_duplicates(subset=[col_unique, col_nazwisko, col_imie, col_upr])
                after = len(final_df)
                show_msg('success', f"✅ ZNALEZIONO {before} WIERSZY. PO USUNIĘCIU DUPLIKATÓW (WG '{col_unique}'): {after}.")
            else:
                show_msg('warning', f"⚠️ NIE MOŻNA POGRUPOWAĆ WYNIKÓW - BRAK KOLUMNY '{col_unique}'.")
        else:
            show_msg('success', f"ZNALEZIONO {len(final_df)} WIERSZY (BEZ USUWANIA DUPLIKATÓW).")
            
        return final_df
    else:
        return pd.DataFrame()

if __name__ == "__main__":
    main()