import webbrowser
import pyautogui
import pyperclip
from time import sleep
import datetime
import os
from urllib.parse import quote
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# ==================== CONFIGURAÇÕES ====================
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
SPREADSHEET_ID = '16O04A4ERu3Twi7OQD6W0Zg9X7j_Uwit9KQcs8wAf2Tw'
RANGE_NAME = 'Mensagens do dia!A1:G500'
INTERVALO_ENTRE_MENSAGENS = 20  # segundos entre cada envio
TEMPO_ESPERA_CARREGAR = 60  # tempo para carregar WhatsApp Web

# ==================== AUTENTICAÇÃO GOOGLE ====================
def autenticar_google():
    """Autentica com Google Sheets API"""
    creds = None
    
    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)
    
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                "client_secret.json", SCOPES
            )
            creds = flow.run_local_server(port=0)
        
        with open("token.json", "w") as token:
            token.write(creds.to_json())
    
    return creds

# ==================== BUSCAR DADOS ====================
def buscar_dados_planilha(creds):
    """Busca dados da planilha do Google Sheets"""
    try:
        service = build("sheets", "v4", credentials=creds)
        sheet = service.spreadsheets()
        
        result = sheet.values().get(
            spreadsheetId=SPREADSHEET_ID,
            range=RANGE_NAME
        ).execute()
        
        values = result.get("values", [])
        
        if not values:
            print("⚠️ Nenhum dado encontrado na planilha.")
            return []
        
        print(f"✅ {len(values)} linhas carregadas da planilha.")
        return values
        
    except HttpError as err:
        print(f"❌ Erro ao acessar Google Sheets: {err}")
        return []

# ==================== ENVIAR MENSAGEM ====================
def enviar_mensagem(telefone, mensagem, nome=""):
    """
    Envia mensagem usando webbrowser + pyautogui
    
    Funciona assim:
    1. Abre WhatsApp Web com o número
    2. Aguarda carregar
    3. Usa pyautogui para clicar e enviar
    """
    try:
        # Limpa o número (remove espaços, traços, parênteses)
        telefone_limpo = ''.join(filter(str.isdigit, telefone))
        
        # Adiciona código do país se não tiver (Brasil = +55)
        if not telefone_limpo.startswith('55'):
            telefone_limpo = '55' + telefone_limpo
        
        print(f"📤 Preparando envio para {nome} ({telefone_limpo})...")
        
        # Monta a URL do WhatsApp Web
        url = f'https://web.whatsapp.com/send?phone={telefone_limpo}&text={quote(mensagem)}'
        
        # Abre no navegador padrão
        webbrowser.open(url)
        
        # Aguarda o WhatsApp Web carregar completamente
        print(f"⏳ Aguardando {TEMPO_ESPERA_CARREGAR} segundos para carregar...")
        sleep(TEMPO_ESPERA_CARREGAR)
        
        # Usa pyperclip para garantir que a mensagem está correta
        # (às vezes a URL encoding pode dar problema com caracteres especiais)
        pyperclip.copy(mensagem)
        
        # Clica no campo de mensagem (geralmente já está focado)
        # e cola a mensagem novamente para garantir
        pyautogui.click()  # Clica na tela para focar
        sleep(0.5)
        
        # Pressiona Tab algumas vezes para garantir foco no campo de mensagem
        # (caso o campo não esteja focado)
        pyautogui.press('tab')
        sleep(0.3)
        
        # Aperta Enter para enviar
        pyautogui.press('enter')
        sleep(2)
        
        print(f"✅ Mensagem enviada para {nome} ({telefone})")
        
        # Fecha a aba (Ctrl+W)
        pyautogui.hotkey('ctrl', 'w')
        sleep(1)
        
        return True
        
    except Exception as e:
        print(f"❌ Erro ao enviar para {nome} ({telefone}): {str(e)}")
        return False

# ==================== REGISTRAR ERRO ====================
def registrar_erro(nome, telefone, erro):
    """Registra erros em arquivo CSV"""
    data_hora = datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    
    # Cria o arquivo com cabeçalho se não existir
    if not os.path.exists('erros.csv'):
        with open('erros.csv', 'w', encoding='utf-8') as arquivo:
            arquivo.write('Data/Hora,Nome,Telefone,Erro\n')
    
    with open('erros.csv', 'a', encoding='utf-8') as arquivo:
        arquivo.write(f'{data_hora},{nome},{telefone},{erro}\n')

# ==================== FUNÇÃO PRINCIPAL ====================
def main():
    print("="*60)
    print("🤖 ROBÔ DE ENVIO DE MENSAGENS - WhatsApp")
    print("="*60)
    print("\n⚠️ ATENÇÃO:")
    print("   1. Certifique-se de estar LOGADO no WhatsApp Web")
    print("   2. O navegador será aberto automaticamente")
    print("   3. NÃO mexa no mouse/teclado durante o envio")
    print("   4. Posicione esta janela de forma que não atrapalhe")
    print("="*60)
    
    input("\n✋ Pressione ENTER para começar em 5 segundos...")
    print("⏳ Iniciando em 5 segundos...")
    sleep(5)
    
    # 1. Autenticar Google
    print("\n📊 Conectando ao Google Sheets...")
    creds = autenticar_google()
    
    # 2. Buscar dados
    values = buscar_dados_planilha(creds)
    if not values:
        return
    
    # 3. Enviar mensagens
    print("\n📤 Iniciando envio de mensagens...\n")
    enviadas = 0
    erros_count = 0
    
    # Começa na linha 5 (índice 5, linha 6 da planilha)
    for i in range(5, len(values)):
        row = values[i]
        
        # Extrai dados com segurança
        telefone = row[0].strip() if len(row) > 0 else ""
        nome = row[1].strip() if len(row) > 1 else "Sem nome"
        mensagem = row[4] if len(row) > 4 else ""
        
        # Valida dados obrigatórios
        if not telefone or not mensagem:
            print(f"⚠️ Linha {i+1}: Dados incompletos - Telefone: {telefone}, Mensagem: {'Sim' if mensagem else 'Não'}")
            continue
        
        # Envia mensagem
        sucesso = enviar_mensagem(telefone, mensagem, nome)
        
        if sucesso:
            enviadas += 1
        else:
            erros_count += 1
            registrar_erro(nome, telefone, "Falha no envio")
        
        # Pausa entre envios para evitar bloqueio do WhatsApp
        if i < len(values) - 1:  # Não espera após a última mensagem
            print(f"⏳ Aguardando {INTERVALO_ENTRE_MENSAGENS} segundos antes do próximo envio...")
            sleep(INTERVALO_ENTRE_MENSAGENS)
    
    # 4. Resumo
    print("\n" + "="*60)
    print("📊 RESUMO DO ENVIO")
    print("="*60)
    print(f"✅ Mensagens enviadas com sucesso: {enviadas}")
    print(f"❌ Mensagens com erro: {erros_count}")
    print(f"📝 Total processado: {enviadas + erros_count}")
    
    if erros_count > 0:
        print(f"\n⚠️ Verifique o arquivo 'erros.csv' para detalhes dos erros")
    
    print("\n✅ Processo finalizado!")

# ==================== EXECUTAR ====================
if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\n⚠️ Processo interrompido pelo usuário.")
    except Exception as e:
        print(f"\n\n❌ Erro crítico: {e}")
