import streamlit as st
import pandas as pd
import string
import mailchimp_marketing as MailchimpMarketing
from mailchimp_marketing.api_client import ApiClientError
import json
import gspread
from google.oauth2.service_account import Credentials

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

# Google Sheets config
GOOGLE_SHEETS_CREDENTIALS = {
    "type": "service_account",
    "project_id": "main-guild-437619-m2",
    "private_key_id": st.secrets["private_key_id"],
    "private_key": st.secrets["google_credentials"],
    "client_email": "systems-specialist@main-guild-437619-m2.iam.gserviceaccount.com",
    "client_id": "101632772111851962527",
    "auth_uri": "https://accounts.google.com/o/oauth2/auth",
    "token_uri": "https://oauth2.googleapis.com/token",
    "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
    "client_x509_cert_url": "https://www.googleapis.com/robot/v1/metadata/x509/systems-specialist%40main-guild-437619-m2.iam.gserviceaccount.com",
    "universe_domain": "googleapis.com"
}

GOOGLE_SHEET_ID = "15adUqyAx4bO-Gz7RweIxxSHziKn3iVwa5e8Ac-iypf0"
SHEET_NAMES = {
    "Bravo NY": "Bravo NY",
    "Bravo FL": "Bravo FL", 
    "CTown": "CTOWN"
}

def test_google_sheets_access():
    """
    Prueba el acceso a Google Sheets para verificar que las credenciales funcionan
    """
    try:
        scopes = [
            'https://www.googleapis.com/auth/spreadsheets',
            'https://www.googleapis.com/auth/drive'
        ]
        creds = Credentials.from_service_account_info(GOOGLE_SHEETS_CREDENTIALS, scopes=scopes)
        client = gspread.authorize(creds)
        
        # Intentar abrir la hoja de cálculo
        sheet = client.open_by_key(GOOGLE_SHEET_ID)
        
        # Verificar que las hojas existen
        for list_name, sheet_name in SHEET_NAMES.items():
            try:
                worksheet = sheet.worksheet(sheet_name)
                st.success(f"✅ Access verified for sheet: {sheet_name}")
            except Exception as e:
                st.error(f"❌ Cannot access sheet '{sheet_name}': {str(e)}")
                return False
        
        return True
        
    except Exception as e:
        st.error(f"❌ Google Sheets access test failed: {str(e)}")
        return False

def save_to_google_sheets(list_name, contacts_data, results):
    """
    Guarda los resultados de la carga en Google Sheets
    contacts_data: lista de diccionarios con la información de los contactos
    results: diccionario con 'success' y 'failed' counts
    """
    try:
        # Configurar credenciales con scopes correctos
        scopes = [
            'https://www.googleapis.com/auth/spreadsheets',
            'https://www.googleapis.com/auth/drive'
        ]
        creds = Credentials.from_service_account_info(GOOGLE_SHEETS_CREDENTIALS, scopes=scopes)
        client = gspread.authorize(creds)
        
        # Abrir la hoja de cálculo
        sheet = client.open_by_key(GOOGLE_SHEET_ID)
        worksheet = sheet.worksheet(SHEET_NAMES[list_name])
        
        # Preparar datos para insertar
        rows_to_insert = []
        for contact in contacts_data:
            email = contact.get('email_address', '')
            first_name = contact.get('merge_fields', {}).get('FNAME', '')
            phone = contact.get('merge_fields', {}).get('PHONE', '')  # Si existe campo phone
            check_status = "✅" if contact.get('uploaded', False) else "❌"
            
            rows_to_insert.append([email, first_name, phone, check_status])
        
        # Insertar datos en la hoja
        if rows_to_insert:
            worksheet.append_rows(rows_to_insert)
            st.success(f"✅ {len(rows_to_insert)} records saved to Google Sheets ({list_name})")
        else:
            st.warning(f"No data to save for {list_name}")
            
    except Exception as e:
        st.error(f"❌ Error saving to Google Sheets ({list_name}): {str(e)}")
        # Mostrar más detalles del error para debugging
        import traceback
        st.error(f"Full error details: {traceback.format_exc()}")

def add_contacts_to_mailchimp(df, lists, extra_fields_map=None, chunk_size=1000):
    import hashlib
    import time
    client = MailchimpMarketing.Client()
    client.set_config({
        "api_key": MAILCHIMP_API_KEY,
        "server": MAILCHIMP_SERVER
    })
    results = {}
    
    for list_name, list_id in lists.items():
        success, failed = 0, 0
        st.info(f"Processing list: {list_name} (ID: {list_id}) with {len(df)} contacts...")
        
        # Preparar todos los contactos
        valid_contacts = []
        invalid_emails = []
        contacts_data = []  # Para guardar en Google Sheets
        
        for idx, row in df.iterrows():
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
                invalid_emails.append(idx)
                failed += 1
                # Agregar a contacts_data como fallido
                contacts_data.append({
                    'email_address': email,
                    'merge_fields': member_info.get('merge_fields', {}),
                    'uploaded': False
                })
                continue
                
            valid_contacts.append(member_info)
            # Agregar a contacts_data para tracking
            contacts_data.append({
                'email_address': email,
                'merge_fields': member_info.get('merge_fields', {}),
                'uploaded': False  # Se actualizará después
            })
        
        if valid_contacts:
            # Dividir en bloques
            total_contacts = len(valid_contacts)
            num_chunks = (total_contacts + chunk_size - 1) // chunk_size  # Redondear hacia arriba
            
            st.info(f"Dividing {total_contacts} contacts into {num_chunks} chunks of {chunk_size} contacts each...")
            
            chunk_start_idx = 0  # Para rastrear el índice en contacts_data
            
            for chunk_idx in range(num_chunks):
                start_idx = chunk_idx * chunk_size
                end_idx = min((chunk_idx + 1) * chunk_size, total_contacts)
                chunk_contacts = valid_contacts[start_idx:end_idx]
                
                st.info(f"Processing chunk {chunk_idx + 1}/{num_chunks} ({len(chunk_contacts)} contacts)...")
                
                # Preparar operaciones para este bloque
                operations = []
                for contact in chunk_contacts:
                    operations.append({
                        "method": "POST",
                        "path": f"/lists/{list_id}/members",
                        "body": json.dumps(contact)
                    })
                
                # Intentar cargar el bloque con reintentos
                max_retries = 3
                retry_count = 0
                chunk_success = False
                
                while retry_count < max_retries and not chunk_success:
                    try:
                        response = client.batches.start({"operations": operations})
                        batch_id = response["id"]
                        
                        st.info(f"Batch started with ID: {batch_id}")
                        
                        # Monitorear el estado del batch
                        status = "pending"
                        attempts = 0
                        max_attempts = 60  # Aumentar a 2 minutos
                        
                        while status in ["pending", "processing", "started"] and attempts < max_attempts:
                            time.sleep(3)  # Aumentar intervalo a 3 segundos
                            try:
                                batch_status = client.batches.status(batch_id)
                                status = batch_status["status"]
                                
                                # Mostrar información detallada del estado
                                if "total_operations" in batch_status:
                                    completed = batch_status.get("finished_operations", 0)
                                    total = batch_status.get("total_operations", 0)
                                    st.info(f"Chunk {chunk_idx + 1}: Status={status}, Progress={completed}/{total}")
                                else:
                                    st.info(f"Chunk {chunk_idx + 1}: Status={status} (attempt {attempts}/{max_attempts})")
                                    
                            except ApiClientError as status_error:
                                st.warning(f"Error checking batch status: {status_error.text}")
                                break
                                
                            attempts += 1
                            
                            # Mostrar progreso cada 10 segundos
                            if attempts % 3 == 0:
                                st.info(f"Chunk {chunk_idx + 1}: Still processing... (attempt {attempts}/{max_attempts})")
                        
                        if status == "finished":
                            chunk_success = True
                            success += len(chunk_contacts)
                            # Marcar como exitoso en contacts_data
                            for i in range(len(chunk_contacts)):
                                contacts_data[chunk_start_idx + i]['uploaded'] = True
                            st.success(f"Chunk {chunk_idx + 1}/{num_chunks} completed successfully. {len(chunk_contacts)} contacts uploaded.")
                        elif status == "started":
                            # Si está en "started" por mucho tiempo, considerarlo como fallido
                            st.error(f"Chunk {chunk_idx + 1} stuck in 'started' status. Moving to next chunk.")
                            failed += len(chunk_contacts)
                        else:
                            st.error(f"Chunk {chunk_idx + 1} failed. Final status: {status}")
                            failed += len(chunk_contacts)
                            
                    except ApiClientError as error:
                        retry_count += 1
                        st.warning(f"Chunk {chunk_idx + 1} failed (attempt {retry_count}/{max_retries}): {error.text}")
                        
                        if retry_count < max_retries:
                            st.info(f"Retrying chunk {chunk_idx + 1} in 10 seconds...")
                            time.sleep(10)  # Aumentar tiempo de espera
                        else:
                            st.error(f"Chunk {chunk_idx + 1} failed after {max_retries} attempts. Moving to next chunk.")
                            failed += len(chunk_contacts)
                
                chunk_start_idx += len(chunk_contacts)
                
                # Pausa entre bloques para evitar rate limiting
                if chunk_idx < num_chunks - 1:
                    st.info("Waiting 5 seconds before next chunk...")
                    time.sleep(5)  # Aumentar pausa entre bloques
            
            st.success(f"Bulk upload completed for {list_name}. Total: {success} successful, {failed} failed.")
            
            # Guardar resultados en Google Sheets
            st.info(f"Saving results to Google Sheets for {list_name}...")
            save_to_google_sheets(list_name, contacts_data, {"success": success, "failed": failed})
        else:
            st.warning(f"No valid contacts found for {list_name}")
            
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
                                    "ZIPCODE": "J",
                                    "PHONE": "K"
                                },
                                "Bravo FL": {
                                    "email": "C",
                                    "FNAME": "D",
                                    "LNAME": "E",
                                    "ADDRESS": "F",
                                    "MMERGE10": "J",  # Full Address Zip
                                    "MMERGE11": "J",   # Zip
                                    "PHONE": "K"
                                },
                                "CTown": {
                                    "email": "C",
                                    "FNAME": "D",
                                    "LNAME": "E",
                                    "ADDRESS": "F",
                                    "ZIPCODE": "J",
                                    "PHONE": "K"
                                }
                            }
                            if not bravo_ny.empty:
                                results['Bravo NY'] = add_contacts_to_mailchimp(bravo_ny, {"Bravo NY": LISTS["Bravo NY"]}, extra_fields_map, 700)['Bravo NY']
                            else:
                                results['Bravo NY'] = {"success": 0, "failed": 0}
                            if not bravo_fl.empty:
                                results['Bravo FL'] = add_contacts_to_mailchimp(bravo_fl, {"Bravo FL": LISTS["Bravo FL"]}, extra_fields_map, 700)['Bravo FL']
                            else:
                                results['Bravo FL'] = {"success": 0, "failed": 0}
                            if not ctown.empty:
                                results['CTown'] = add_contacts_to_mailchimp(ctown, {"CTown": LISTS["CTown"]}, extra_fields_map, 700)['CTown']
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
