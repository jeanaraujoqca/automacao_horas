import streamlit as st
import msal
import requests 
import pandas as pd
from datetime import datetime
from cryptography.hazmat.primitives import serialization
from cryptography.hazmat.backends import default_backend
import os

# Configurações do cliente (de preferência, carregue essas variáveis de ambiente)
client_id = os.getenv('CLIENT_ID')
client_secret = os.getenv('CLIENT_SECRET')
tenant_id = os.getenv('TENANT_ID')
cert_path = 'caminho/para/seu/certificado.pem'  # Atualize o caminho do certificado
cert_password = os.getenv('CERT_PASSWORD').encode()  # Converta a senha em bytes

# Inicialize Streamlit
st.title("Upload e Envio de Dados para SharePoint")

# Upload do arquivo Excel
uploaded_file = st.file_uploader("Escolha o arquivo Excel", type="xlsx")
if uploaded_file:
    df = pd.read_excel(uploaded_file)
    st.write("Pré-visualização dos dados:", df.head())  # Mostra uma prévia dos dados

    if st.button("Enviar para SharePoint"):
        # Carregar chave privada do certificado
        with open(cert_path, 'rb') as pem_file:
            private_key = serialization.load_pem_private_key(
                pem_file.read(),
                password=None,
                backend=default_backend()
            )

        # Obter token de autenticação
        authority = f'https://login.microsoftonline.com/{tenant_id}'
        scope = ['https://queirozcavalcanti.sharepoint.com/.default']
        app = msal.ConfidentialClientApplication(client_id, authority=authority,
                                                 client_credential={
                                                    "private_key": private_key,
                                                    "thumbprint": "seu_thumbprint"
                                                 })
        token_response = app.acquire_token_for_client(scopes=scope)
        
        if 'access_token' in token_response:
            access_token = token_response['access_token']
            headers = {
                'Authorization': f'Bearer {access_token}',
                'Accept': 'application/json;odata=verbose',
                'Content-Type': 'application/json;odata=verbose'
            }
            st.success("Token de acesso obtido com sucesso!")
        else:
            st.error("Erro ao obter token")

        # Percorre o DataFrame
        for index, row in df.iterrows():
            try:
                email = row['EMAIL']
                unidade = row['UNIDADE']
                treinamento = row['TREINAMENTO']
                carga_horaria = str(row['CARGA HORARIA'])
                tipo_treinamento = row['TIPO DO TREINAMENTO']
                inicio_convertido = datetime.strptime(row['INICIO DO TREINAMENTO'], "%d/%m/%Y").strftime("%Y-%m-%dT%H:%M:%S")
                termino_convertido = datetime.strptime(row['TERMINO DO TREINAMENTO'], "%d/%m/%Y").strftime("%Y-%m-%dT%H:%M:%S")
                categoria = row['CATEGORIA']
                instituicao_instrutor = row['INSTITUIÇÃO/INSTRUTOR']
                
                # Obter ID do usuário no SharePoint
                user_url = f"https://queirozcavalcanti.sharepoint.com/sites/qca360/_api/web/siteusers/getbyemail('{email}')"
                response = requests.get(user_url, headers=headers)
                
                if response.status_code == 200:
                    correct_user_id = response.json()['d']['Id']
                else:
                    st.error(f"Erro ao garantir o usuário para o email {email}: {response.status_code}")
                    continue

                # Dados do item a serem adicionados
                item_data = {
                    "__metadata": {"type": "SP.Data.Treinamentos_x005f_qcaListItem"},
                    "NOMEDOINTEGRANTEId": correct_user_id,
                    "Title": treinamento,
                    "CARGAHORARIA": carga_horaria,
                    "TIPO_x0020_DO_x0020_TREINAMENTO_": tipo_treinamento,
                    "INICIO_x0020_DO_x0020_TREINAMENT": inicio_convertido,
                    "TERMINO_x0020_DO_x0020_TREINAMEN": termino_convertido,
                    "TIPO_": categoria,
                    "INSTITUI_x00c7__x00c3_O_x002f_IN": instituicao_instrutor,
                    "UNIDADE": unidade,
                    "E_x002d_MAILId": correct_user_id
                }

                add_item_url = f"https://queirozcavalcanti.sharepoint.com/sites/qca360/_api/web/lists/getbytitle('Treinamento de atividades')/items"
                response = requests.post(add_item_url, headers=headers, json=item_data)

                if response.status_code == 201:
                    st.success(f"Item adicionado com sucesso para {email}")
                else:
                    st.error(f"Erro ao adicionar item para {email}: {response.status_code}")
            except Exception as e:
                st.error(f"Erro ao processar linha {index}: {str(e)}")
