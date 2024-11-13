import streamlit as st
import msal
import requests 
import pandas as pd
from datetime import datetime
from cryptography.hazmat.primitives import serialization
from cryptography.hazmat.backends import default_backend
import os
import tempfile
import base64

# Configuração da página deve ser a primeira linha de Streamlit
st.set_page_config(
    page_title="Automação de Horas",
    page_icon='qca_logo_2.png',
    layout="wide",
)

def customize_page():
    st.markdown(
        """
        <style>
        /* Fundo da página */
        html, body, .stApp {
            background-image: url("https://raw.githubusercontent.com/jeanaraujoqca/automacao_horas/refs/heads/main/bg_dark.png");
            background-size: cover;
            background-repeat: no-repeat;
            background-attachment: fixed;
            height: 100vh;
            margin: 0;
            padding: 0;
        }

        /* Transparência no cabeçalho */
        header, .css-18e3th9, .css-1d391kg, [data-testid="stHeader"] {
            background-color: rgba(0, 0, 0, 0) !important;
            color: #ffffff !important; /* Texto branco */
        }

        /* Estilo dos títulos e subtítulos em branco */
        h1 {
            color: #ffffff !important;  /* Texto em branco */
            font-size: 28px;
        }

        /* Caixa de upload de arquivo em branco */
        .stFileUploader {
            background-color: #ffffff !important;  /* Fundo branco */
            border: 2px solid #ffffff !important;   /* Borda branca */
            color: #000000 !important;              /* Texto preto para contraste dentro da caixa */
            border-radius: 10px;                    /* Bordas arredondadas */
            padding: 10px;
        }
        </style>
        """,
        unsafe_allow_html=True
    )

# Aplicar o fundo da página
customize_page()

# Título da página
st.title("Automação de Lançamento de Horas de Treinamento no SharePoint")

# Função para exibir mensagens de erro amigáveis
def mostrar_erro_personalizado(mensagem, sugestao=None):
    st.error(f"🚨 {mensagem}")
    if sugestao:
        st.info(f"💡 {sugestao}")

# Carregar as variáveis de ambiente
variaveis = {
    "CLIENT_ID": os.getenv('CLIENT_ID'),
    "TENANT_ID": os.getenv('TENANT_ID'),
    "CERT_PASSWORD": os.getenv('CERT_PASSWORD', '').encode(),  # Converta a senha em bytes
    "THUMBPRINT": os.getenv('THUMBPRINT'),
    "CERTIFICADO_BASE64": os.getenv("CERTIFICADO_BASE64")
}

# Verifique se há variáveis de ambiente ausentes e exiba um erro para cada uma
missing_vars = [var for var, value in variaveis.items() if not value]

if missing_vars:
    mostrar_erro_personalizado(
        "Algumas variáveis de ambiente estão ausentes e são necessárias para a execução.",
        "Verifique as configurações e certifique-se de que todas as variáveis de ambiente estão corretamente configuradas."
    )
    for var in missing_vars:
        st.write(f"- **{var}** está ausente")
    st.stop()  # Interrompe a execução se houver variáveis de ambiente ausentes

cert_pem = base64.b64decode(variaveis["CERTIFICADO_BASE64"])

# Salve o certificado temporariamente
with tempfile.NamedTemporaryFile(delete=False, suffix=".pem") as temp_cert_file:
    temp_cert_file.write(cert_pem)
    temp_cert_path = temp_cert_file.name

# Função para obter token de autenticação com tratamento de erro
def obter_token():
    try:
        # Carregue a chave privada do certificado temporário
        with open(temp_cert_path, 'rb') as pem_file:
            private_key = serialization.load_pem_private_key(
                pem_file.read(),
                password=None,  # Coloque a senha se o PEM estiver protegido
                backend=default_backend()
            )

        # Configuração MSAL
        authority = f'https://login.microsoftonline.com/{variaveis["TENANT_ID"]}'
        scope = ['https://queirozcavalcanti.sharepoint.com/.default']
        app = msal.ConfidentialClientApplication(
            variaveis["CLIENT_ID"], authority=authority,
            client_credential={
                "private_key": private_key,
                "thumbprint": variaveis["THUMBPRINT"]
            }
        )
        token_response = app.acquire_token_for_client(scopes=scope)

        if 'access_token' in token_response:
            return token_response['access_token']
        else:
            mostrar_erro_personalizado(
                "Falha ao obter token de acesso.",
                "Verifique as credenciais e tente novamente."
            )
            st.stop()
    except Exception as e:
        mostrar_erro_personalizado(
            f"Erro ao obter token de autenticação: {str(e)}",
            "Verifique as configurações do certificado e do tenant."
        )
        st.stop()
    finally:
        os.remove(temp_cert_path)  # Remover arquivo temporário após uso

# Obter e validar o token
access_token = obter_token()
headers = {
    'Authorization': f'Bearer {access_token}',
    'Accept': 'application/json;odata=verbose',
    'Content-Type': 'application/json;odata=verbose'
}
st.success("Token de acesso obtido com sucesso!")

# Upload do arquivo Excel
uploaded_file = st.file_uploader("Escolha o arquivo Excel", type="xlsx")
if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file)
        st.write("Pré-visualização dos dados:", df.head())  # Mostra uma prévia dos dados
    except Exception as e:
        mostrar_erro_personalizado(
            "Erro ao carregar o arquivo Excel.",
            "Verifique se o arquivo está no formato correto e tente novamente."
        )

    if st.button("Enviar para SharePoint"):
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
                    mostrar_erro_personalizado(
                        f"Erro ao buscar o usuário para o email {email}",
                        "Certifique-se de que o email está correto e registrado no sistema."
                    )
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
                    mostrar_erro_personalizado(
                        f"Erro ao adicionar item para {email}",
                        f"Código de status: {response.status_code}. Verifique as configurações do SharePoint."
                    )
            except Exception as e:
                mostrar_erro_personalizado(
                    f"Erro ao processar para {email}: {str(e)}",
                    "Verifique o conteúdo e formato do arquivo."
                )
