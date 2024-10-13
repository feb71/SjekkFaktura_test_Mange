import streamlit as st
import pdfplumber
import pandas as pd
import re
from io import BytesIO

st.set_page_config(page_title="Streamlit App", layout="wide", initial_sidebar_state="expanded")

# Funksjon for å lese fakturanummer fra PDF
def get_invoice_number(file):
    try:
        with pdfplumber.open(file) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                match = re.search(r"Fakturanummer\s*[:\-]?\s*(\d+)", text, re.IGNORECASE)
                if match:
                    return match.group(1)
        return None
    except Exception as e:
        st.error(f"Kunne ikke lese fakturanummer fra PDF: {e}")
        return None

# Funksjon for å lese PDF-filen og hente ut relevante data
def extract_data_from_pdf(file, doc_type, invoice_number=None):
    try:
        with pdfplumber.open(file) as pdf:
            data = []
            start_reading = False

            for page in pdf.pages:
                text = page.extract_text()
                if text is None:
                    st.error(f"Ingen tekst funnet på side {page.page_number} i PDF-filen.")
                    continue
                
                lines = text.split('\n')
                for line in lines:
                    if doc_type == "Faktura" and "Artikkel" in line:
                        start_reading = True
                        continue

                    if start_reading:
                        columns = line.split()
                        if len(columns) >= 5:
                            item_number = columns[1]
                            if not item_number.isdigit():
                                continue

                            # Trekke ut beskrivelse, og antall hvis det finnes på slutten av beskrivelsen
                            description = " ".join(columns[2:-4])
                            try:
                                # Hvis antall er inkludert i beskrivelsen (for varer kun i faktura)
                                antall_fra_beskrivelse = re.search(r'(\d+)\s*$', description)
                                if antall_fra_beskrivelse:
                                    quantity = float(antall_fra_beskrivelse.group(1).replace('.', '').replace(',', '.'))
                                    description = re.sub(r'\s*\d+\s*$', '', description)
                                else:
                                    quantity = float(columns[-4].replace('.', '').replace(',', '.')) if columns[-4].replace('.', '').replace(',', '').isdigit() else columns[-4]
                                
                                unit_price = float(columns[-3].replace('.', '').replace(',', '.')) if columns[-3].replace('.', '').replace(',', '').isdigit() else columns[-3]
                                discount = float(columns[-2].replace('.', '').replace(',', '.')) if columns[-2].replace('.', '').replace(',', '').isdigit() else 0  # Sett rabatt til 0 hvis tom
                                total_price = float(columns[-1].replace('.', '').replace(',', '.')) if columns[-1].replace('.', '').replace(',', '').isdigit() else columns[-1]
                            except ValueError as e:
                                st.error(f"Kunne ikke konvertere til flyttall: {e}")
                                continue

                            unique_id = f"{invoice_number}_{item_number}" if invoice_number else item_number
                            data.append({
                                "UnikID": unique_id,
                                "Varenummer": item_number,
                                "Beskrivelse_Faktura": description,
                                "Antall_Faktura": quantity,
                                "Enhetspris_Faktura": unit_price,
                                "Rabatt": discount,
                                "Beløp_Faktura": total_price,
                                "Type": doc_type
                            })
            if len(data) == 0:
                st.error("Ingen data ble funnet i PDF-filen.")
                
            return pd.DataFrame(data)
    except Exception as e:
        st.error(f"Kunne ikke lese data fra PDF: {e}")
        return pd.DataFrame()

# Funksjon for å konvertere DataFrame til en Excel-fil
def convert_df_to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
    return output.getvalue()

def main():
    st.title("Sammenlign Faktura mot Tilbud")
    st.markdown("""<style>.dataframe th {font-weight: bold !important;}</style>""", unsafe_allow_html=True)

    # Opprett tre kolonner
    col1, col2, col3 = st.columns([1, 5, 1])

    with col1:
        st.header("Last opp filer")
        invoice_files = st.file_uploader("Last opp fakturaer fra Brødrene Dahl", type="pdf", accept_multiple_files=True)
        offer_file = st.file_uploader("Last opp tilbud fra Brødrene Dahl (Excel)", type="xlsx")

    if invoice_files and offer_file:
        # Les tilbudet fra Excel-filen
        with col1:
            st.info("Laster inn tilbud fra Excel-filen...")
        offer_data = pd.read_excel(offer_file)

        # Riktige kolonnenavn fra Excel-filen for tilbud
        offer_data.rename(columns={
            'VARENR': 'Varenummer',
            'BESKRIVELSE': 'Beskrivelse_Tilbud',
            'ANTALL': 'Antall_Tilbud',
            'ENHET': 'Enhet_Tilbud',
            'ENHETSPRIS': 'Enhetspris_Tilbud',
            'TOTALPRIS': 'Totalt pris'
        }, inplace=True)

        all_invoice_data = pd.DataFrame()  # Samler all fakturadata

        # Løp gjennom alle fakturafiler
        for invoice_file in invoice_files:
            # Hent fakturanummer
            with col1:
                invoice_number = get_invoice_number(invoice_file)
                if invoice_number:
                    st.success(f"Fakturanummer funnet: {invoice_number}")
                else:
                    st.error(f"Kunne ikke finne fakturanummer for {invoice_file.name}")
                    continue  # Gå til neste faktura hvis ingen fakturanummer ble funnet

            # Ekstraher data fra PDF-filer
            with col1:
                st.info(f"Laster inn faktura: {invoice_file.name}")
            invoice_data = extract_data_from_pdf(invoice_file, "Faktura", invoice_number)

            # Legg til fakturadataene i den totale dataframen
            all_invoice_data = pd.concat([all_invoice_data, invoice_data], ignore_index=True)

        if not all_invoice_data.empty and not offer_data.empty:
            with col2:
                st.write("Sammenligner data...")

            # Merge faktura- og tilbudsdataene
            merged_data = pd.merge(offer_data, all_invoice_data, on="Varenummer", how='outer', suffixes=('_Tilbud', '_Faktura'))

            # Konverter kolonner til numerisk der det er relevant
            merged_data["Antall_Faktura"] = pd.to_numeric(merged_data["Antall_Faktura"], errors='coerce')
            merged_data["Antall_Tilbud"] = pd.to_numeric(merged_data["Antall_Tilbud"], errors='coerce')
            merged_data["Enhetspris_Faktura"] = pd.to_numeric(merged_data["Enhetspris_Faktura"], errors='coerce')
            merged_data["Enhetspris_Tilbud"] = pd.to_numeric(merged_data["Enhetspris_Tilbud"], errors='coerce')

            # Flytt verdier fra "Rabatt" til "Enhetspris_Faktura" der det er feil
            merged_data["Enhetspris_Faktura"] = merged_data.apply(
                lambda row: row["Rabatt"] if pd.isna(row["Enhetspris_Faktura"]) and not pd.isna(row["Rabatt"]) else row["Enhetspris_Faktura"],
                axis=1
            )

            # Fjern verdiene fra rabattkolonnen der de er flyttet
            merged_data["Rabatt"] = merged_data.apply(
                lambda row: None if row["Enhetspris_Faktura"] == row["Rabatt"] else row["Rabatt"],
                axis=1
            )

            # Finne avvik
            merged_data["Avvik_Antall"] = merged_data["Antall_Faktura"] - merged_data["Antall_Tilbud"]
            merged_data["Avvik_Enhetspris"] = merged_data["Enhetspris_Faktura"] - merged_data["Enhetspris_Tilbud"]
            merged_data["Prosentvis_økning"] = ((merged_data["Enhetspris_Faktura"] - merged_data["Enhetspris_Tilbud"]) / merged_data["Enhetspris_Tilbud"]) * 100

            # Filtrer avvik
            avvik = merged_data[(merged_data["Avvik_Antall"].notna() & (merged_data["Avvik_Antall"] != 0)) |
                                (merged_data["Avvik_Enhetspris"].notna() & (merged_data["Avvik_Enhetspris"] != 0))]

            with col2:
                st.subheader("Avvik mellom Faktura og Tilbud")
                st.dataframe(avvik)

            # Artikler som finnes i faktura, men ikke i tilbud
            only_in_invoice = merged_data[merged_data['Enhetspris_Tilbud'].isna()]
            with col2:
                st.subheader("Varenummer som finnes i faktura, men ikke i tilbud")
                st.dataframe(only_in_invoice)

            # Lagre kun artikkeldataene til XLSX
            all_items = all_invoice_data[["UnikID", "Varenummer", "Beskrivelse_Faktura", "Antall_Faktura", "Enhetspris_Faktura", "Beløp_Faktura", "Rabatt"]]
            
            excel_data = convert_df_to_excel(all_items)

            with col3:
                st.download_button(
                    label="Last ned avviksrapport som Excel",
                    data=convert_df_to_excel(avvik),
                    file_name="avvik_rapport.xlsx"
                )
                
                st.download_button(
                    label="Last ned alle varenummer som Excel",
                    data=excel_data,
                    file_name="faktura_varer.xlsx",
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                )

                # Lag en Excel-fil med varenummer som finnes i faktura, men ikke i tilbud
                only_in_invoice_data = convert_df_to_excel(only_in_invoice)
                st.download_button(
                    label="Last ned varenummer som ikke eksiterer i tilbudet",
                    data=only_in_invoice_data,
                    file_name="varer_kun_i_faktura.xlsx",
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                )
        else:
            st.error("Kunne ikke lese data fra tilbudet eller fakturaene.")
    else:
        st.error("Last opp både fakturaer og tilbud.")

if __name__ == "__main__":
    main()

