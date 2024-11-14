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
from io import BytesIO
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

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
            color: #ffffff !important;
        }
        /* Estilo dos títulos e subtítulos em branco */
        h1 {
            color: #ffffff !important;
            font-size: 32px;
        }

        p {
            color: #FF0000;
            font-weight: bold;
            font-size: 14px;
        }
        
        /* Caixa de upload de arquivo em branco */
        .stFileUploader {
            background-color: #ffffff !important;
            border: 2px solid #ffffff !important;
            color: #000000 !important;
            border-radius: 10px;
            padding: 10px;
        }
        </style>
        """,
        unsafe_allow_html=True
    )

customize_page()

st.title("Automação de Lançamento de Horas de Treinamento")

# Carregar as variáveis de ambiente
client_id = os.getenv('CLIENT_ID')
tenant_id = os.getenv('TENANT_ID')
cert_password = os.getenv('CERT_PASSWORD', '').encode()
thumbprint = os.getenv('THUMBPRINT')
cert_base64 = os.getenv("CERTIFICADO_BASE64")
email_jean = os.getenv("EMAIL_JEAN")
email_dani = os.getenv("EMAIL_DANI")
senha_email = os.getenv("SENHA_EMAIL")
email_tanus = os.getenv("EMAIL_TANUS")

# Função para obter token de autenticação
def obter_token():
    try:
        authority = f'https://login.microsoftonline.com/{tenant_id}'
        scope = ['https://queirozcavalcanti.sharepoint.com/.default']
        app = msal.ConfidentialClientApplication(
            client_id, authority=authority,
            client_credential={
                "private_key": serialization.load_pem_private_key(
                    base64.b64decode(cert_base64), 
                    password=None, 
                    backend=default_backend()
                ),
                "thumbprint": thumbprint
            }
        )
        token_response = app.acquire_token_for_client(scopes=scope)
        if 'access_token' in token_response:
            return token_response['access_token']
        else:
            raise ValueError("Erro ao obter token de acesso.")
    except Exception as e:
        st.error(f"Erro ao obter token de autenticação: {str(e)}")
        st.stop()

# Função para enviar e-mail
def enviar_email(relatorio, nome, equipe, total_sucesso, total_erro):
    remetente = email_jean  # Substitua pelo seu e-mail
    destinatario = email_dani, email_jean, email_tanus  # Substitua pelo e-mail destinatário
    senha = senha_email  # Substitua pela senha do e-mail

    # Configuração do e-mail
    mensagem = MIMEMultipart()
    mensagem["From"] = remetente
    mensagem["To"] = destinatario
    mensagem["Subject"] = "Relatório de Utilização - Automação de Horas"

    # Corpo do e-mail
    corpo = f"""
    Segue o relatório sobre a utilização do sistema de lançamento de treinamentos.

    Usuário: {nome}
    Equipe: {equipe}
    Total de Treinamentos Lançados com Sucesso: {total_sucesso}
    Total de Treinamentos com Erro: {total_erro}

    Em anexo, relatório detalhado.
    """
    mensagem.attach(MIMEText(corpo, "plain"))

    # Anexar o relatório
    part = MIMEBase("application", "octet-stream")
    part.set_payload(relatorio.getvalue())
    encoders.encode_base64(part)
    part.add_header("Content-Disposition", "attachment; filename=relatorio_resultados.xlsx")
    mensagem.attach(part)

    # Envio do e-mail
    with smtplib.SMTP("smtp-mail.outlook.com", 587) as servidor:
        servidor.starttls()
        servidor.login(remetente, senha)
        servidor.sendmail(remetente, destinatario, mensagem.as_string())

# Campos de entrada para nome e equipe
nome = st.text_input("Nome:")
equipe = st.text_input("Equipe:")

# Obter e validar o token
access_token = obter_token()
headers = {
    'Authorization': f'Bearer {access_token}',
    'Accept': 'application/json;odata=verbose',
    'Content-Type': 'application/json;odata=verbose'
}

# Upload do arquivo Excel
uploaded_file = st.file_uploader("Escolha o arquivo Excel:", type="xlsx")
if uploaded_file and nome and equipe:
    df = pd.read_excel(uploaded_file)
    st.write("Pré-visualização dos dados:", df.head())

    # Verificar colunas necessárias
    required_columns = ['EMAIL', 'UNIDADE', 'TREINAMENTO', 'CARGA HORARIA', 'TIPO DO TREINAMENTO', 'INICIO DO TREINAMENTO', 'TERMINO DO TREINAMENTO', 'CATEGORIA', 'INSTITUIÇÃO/INSTRUTOR']
    missing_columns = [col for col in required_columns if col not in df.columns]
    
    if missing_columns:
        st.error(f"Erro: As seguintes colunas estão ausentes no arquivo: {', '.join(missing_columns)}")
        st.stop()

    if st.button("Enviar para SharePoint"):
        if nome and equipe:
            st.write("Aguarde, estamos lançando os treinamentos...")
    
            # Lista para armazenar o status de cada treinamento
            resultados = []
            total_sucesso = 0
            total_erro = 0
    
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
                        raise ValueError(f"Erro ao buscar o usuário para o email {email}: {response.status_code}")
    
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
                        "E_x002d_MAILId": correct_user_id,
                    }
    
                    add_item_url = f"https://queirozcavalcanti.sharepoint.com/sites/qca360/_api/web/lists/getbytitle('Treinamento de atividades')/items"
                    response = requests.post(add_item_url, headers=headers, json=item_data)
    
                    if response.status_code == 201:
                        resultados.append({"Email": email, "Treinamento": treinamento, "Status": "Sucesso", "Mensagem": "Item adicionado com sucesso"})
                        total_sucesso += 1
                    else:
                        resultados.append({"Email": email, "Treinamento": treinamento, "Status": "Erro", "Mensagem": f"Erro ao adicionar item: {response.status_code}"})
                        total_erro += 1
                except Exception as e:
                    resultados.append({"Email": email, "Treinamento": treinamento, "Status": "Erro", "Mensagem": str(e)})
                    total_erro += 1
        else:
            st.error("Preencha os campos de nome e equipe.")

        # Gerar o relatório em um DataFrame
        df_resultados = pd.DataFrame(resultados)

        # Salvar o relatório como um arquivo Excel em memória
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_resultados.to_excel(writer, index=False, sheet_name='Resultados')
        output.seek(0)

        # Enviar e-mail com o relatório
        enviar_email(output, nome, equipe, total_sucesso, total_erro)

        # Exibir botão para download do relatório
        st.success("Processamento concluído! Baixe o relatório abaixo.")
        st.download_button(
            label="Baixar Relatório de Resultados",
            data=output,
            file_name="relatorio_resultados.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
