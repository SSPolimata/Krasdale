import streamlit as st
import pandas as pd
import string
import mailchimp_marketing as MailchimpMarketing
from mailchimp_marketing.api_client import ApiClientError

# Configuración de la página
st.set_page_config(
    page_title="Krasdale - Spreadsheet to MailChimp",
    page_icon="📄",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# CSS personalizado para el tema claro
st.markdown("""
    <style>
        [data-testid="stAppViewContainer"] { background-color: #ffffff; }
        [data-testid="stSidebar"] { background-color: #f8f9fa; }
        .main { padding: 0rem 1rem; background-color: #ffffff; }
        .stTitle { font-size: 3rem; color: #2c3e50; padding-bottom: 1rem; }
        .stAlert { padding: 1rem; margin: 1rem 0; border-radius: 0.5rem; }
        .css-1v0mbdj.ebxwdo61 { margin-top: 2rem; }
        .stMarkdown { color: #2c3e50; }
        [data-testid="stMetricValue"] { color: #2c3e50; background-color: #ffffff; }
        .stDataFrame { background-color: #ffffff; }
        .stButton button { background-color: #2c3e50; color: #ffffff; border-radius: 0.5rem; }
        [data-testid="stFileUploader"] { background-color: #f8f9fa; padding: 1rem; border-radius: 0.5rem; }
        .streamlit-expanderHeader { background-color: #f8f9fa; color: #2c3e50; }
    </style>
""", unsafe_allow_html=True)

# Protección por contraseña

def check_password():
    def password_entered():
        if st.session_state["password"] == st.secrets["password"]:
            st.session_state["password_correct"] = True
            del st.session_state["password"]
        else:
            st.session_state["password_correct"] = False

    if "password_correct" not in st.session_state:
        st.markdown("""
            <div style='text-align: center; padding: 1rem;'>
                <h1 style='color: #2c3e50;'>Krasdale</h1>
                <h2 style='color: #7f8c8d;'>Spreadsheet to MailChimp</h2>
            </div>
        """, unsafe_allow_html=True)
        st.text_input(
            "Please enter the password to access the application",
            type="password",
            on_change=password_entered,
            key="password"
        )
        return False
    return st.session_state["password_correct"]

def excel_column_names(n):
    """Generate Excel-style column names (A, B, ..., Z, AA, AB, ...) for n columns."""
    names = []
    for i in range(n):
        name = ''
        col = i
        while True:
            name = chr(65 + (col % 26)) + name
            col = col // 26 - 1
            if col < 0:
                break
        names.append(name)
    return names

# Mailchimp config
MAILCHIMP_API_KEY = st.secrets["mailchimp_api_key"]
MAILCHIMP_SERVER = st.secrets["mailchimp_server"]
LISTS = {
    "Bravo NY": "0a06e5f3d3",
    "Bravo FL": "eab6821d7c",
    "CTown": "7a827d6afc"
}

def add_contacts_to_mailchimp(df, lists, extra_fields_map=None):
    import hashlib
    client = MailchimpMarketing.Client()
    client.set_config({
        "api_key": MAILCHIMP_API_KEY,
        "server": MAILCHIMP_SERVER
    })
    results = {}
    for list_name, list_id in lists.items():
        success, failed = 0, 0
        st.info(f"Processing list: {list_name} (ID: {list_id}) with {len(df)} contacts...")
        for idx, row in df.iterrows():
            # Por defecto, email en columna B
            email = str(row['B']).strip()
            member_info = {}
            if extra_fields_map and list_name in extra_fields_map:
                email = str(row[extra_fields_map[list_name]['email']]).strip()
                merge_fields = {}
                for field, col in extra_fields_map[list_name].items():
                    if field == 'email':
                        continue
                    merge_fields[field] = str(row[col]).strip()
                member_info = {
                    "email_address": email,
                    "status": "subscribed",
                    "merge_fields": merge_fields
                }
            else:
                member_info = {
                    "email_address": email,
                    "status": "subscribed"
                }
            if pd.isna(email) or '@' not in email:
                st.warning(f"Skipping row {idx+1}: invalid email '{email}'")
                failed += 1
                continue
            st.write(f"Adding to {list_name}: {member_info}")
            try:
                client.lists.add_list_member(list_id, member_info)
                st.success(f"Contact {email} added to {list_name}.")
                success += 1
            except ApiClientError as error:
                # Si el error es Member Exists, hacer update
                if error.status_code == 400 and 'Member Exists' in error.text:
                    # Mailchimp requiere el hash MD5 del email en minúsculas
                    email_hash = hashlib.md5(email.lower().encode('utf-8')).hexdigest()
                    try:
                        client.lists.update_list_member(list_id, email_hash, member_info)
                        st.success(f"Contact {email} updated in {list_name}.")
                        success += 1
                    except ApiClientError as update_error:
                        st.error(f"Failed to update {email} in {list_name}: {update_error.text}")
                        failed += 1
                else:
                    st.error(f"Failed to add {email} to {list_name}: {error.text}")
                    failed += 1
        results[list_name] = {"success": success, "failed": failed}
    return results

def main():
    if not check_password():
        st.error("⚠️ Password incorrect. Please try again.")
        return

    st.markdown("""
        <div style='text-align: center; padding: 1rem;'>
            <h1 style='color: #2c3e50;'>Krasdale</h1>
            <h2 style='color: #7f8c8d;'>Spreadsheet to MailChimp</h2>
        </div>
    """, unsafe_allow_html=True)

    st.info(f"Mailchimp List IDs in use: Bravo NY: {LISTS['Bravo NY']}, Bravo FL: {LISTS['Bravo FL']}, CTown: {LISTS['CTown']}")

    uploaded_file = st.file_uploader("Choose a xlsx or csv file (without headers)", type=['xlsx', 'csv'])

    df = None
    filtered_df = None
    if uploaded_file is not None:
        try:
            if uploaded_file.name.endswith('.xlsx'):
                df = pd.read_excel(uploaded_file, header=None)
            elif uploaded_file.name.endswith('.csv'):
                df = pd.read_csv(uploaded_file, header=None, on_bad_lines='skip')
            # Limpiar celdas con formato ="..."
            df = df.applymap(lambda x: str(x)[2:-1] if isinstance(x, str) and x.startswith('="') and x.endswith('"') else x)
            col_names = excel_column_names(df.shape[1])
            df.columns = col_names
            # Formatear columna J a 5 dígitos con ceros a la izquierda
            if 'J' in df.columns:
                df['J'] = df['J'].apply(lambda x: str(x).zfill(5) if pd.notna(x) and str(x).isdigit() and len(str(x)) < 5 else str(x))
            if 'L' in df.columns:
                filtered_df = df[df['L'].astype(str).str.lower() == 'active']
                st.success("File uploaded and filtered successfully. Preview:")
                st.dataframe(filtered_df.head())
                if not filtered_df.empty:
                    # Asignar leads a listas según reglas
                    bravo_ny = filtered_df[(filtered_df['A'].astype(str).str.upper() == 'BRAVO') & (filtered_df['B'].astype(str).str.startswith(('43', '043')))]
                    bravo_fl = filtered_df[(filtered_df['A'].astype(str).str.upper() == 'BRAVO') & (filtered_df['B'].astype(str).str.startswith(('45', '045')))]
                    ctown = filtered_df[(filtered_df['A'].astype(str).str.upper() == 'CTOWN') & (filtered_df['B'].astype(str).str.startswith(('41', '041')))]

                    st.info(f"Bravo NY: {len(bravo_ny)} leads | Bravo FL: {len(bravo_fl)} leads | CTown: {len(ctown)} leads")

                    if st.button("Upload filtered contacts to Mailchimp lists"):
                        with st.spinner("Uploading contacts to Mailchimp..."):
                            results = {}
                            extra_fields_map = {
                                "Bravo NY": {
                                    "email": "C",
                                    "FNAME": "D",
                                    "LNAME": "E",
                                    "ADDRESS": "F",
                                    "ZIPCODE": "J"
                                },
                                "Bravo FL": {
                                    "email": "C",
                                    "FNAME": "D",
                                    "LNAME": "E",
                                    "ADDRESS": "F",
                                    "MMERGE10": "J",  # Full Address Zip
                                    "MMERGE11": "J"   # Zip
                                },
                                "CTown": {
                                    "email": "C",
                                    "FNAME": "D",
                                    "LNAME": "E",
                                    "ADDRESS": "F",
                                    "ZIPCODE": "J"
                                }
                            }
                            if not bravo_ny.empty:
                                results['Bravo NY'] = add_contacts_to_mailchimp(bravo_ny, {"Bravo NY": LISTS["Bravo NY"]}, extra_fields_map)['Bravo NY']
                            else:
                                results['Bravo NY'] = {"success": 0, "failed": 0}
                            if not bravo_fl.empty:
                                results['Bravo FL'] = add_contacts_to_mailchimp(bravo_fl, {"Bravo FL": LISTS["Bravo FL"]}, extra_fields_map)['Bravo FL']
                            else:
                                results['Bravo FL'] = {"success": 0, "failed": 0}
                            if not ctown.empty:
                                results['CTown'] = add_contacts_to_mailchimp(ctown, {"CTown": LISTS["CTown"]}, extra_fields_map)['CTown']
                            else:
                                results['CTown'] = {"success": 0, "failed": 0}
                        for list_name, res in results.items():
                            st.info(f"List '{list_name}': {res['success']} contacts added, {res['failed']} failed.")
            else:
                st.warning("Column 'L' does not exist in the uploaded file. The file must have at least 12 columns.")
                st.dataframe(df.head())
        except Exception as e:
            st.error(f"Error reading the file: {e}")

if __name__ == "__main__":
    main() 
