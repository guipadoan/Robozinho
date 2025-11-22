import pywhatkit as kit
from time import sleep
import datetime
import os
import json
import pytz
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# Configurar fuso horário de Brasília
FUSO_HORARIO = pytz.timezone('America/Sao_Paulo')

# ==================== CONFIGURAÇÕES ====================
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
SPREADSHEET_ID = '1PEr-cjNy99QtJWnVAPPwR43NkAesHXQLZSF0QukcLW4'
RANGE_NAME = 'Robozinho!A1:E10000'
INTERVALO_ENTRE_MENSAGENS = 20  # segundos entre cada envio

# Configurações de horário (Horário de Brasília)
HORA_INICIO = 8   # Começa às 8h
HORA_FIM = 22     # Para às 22h (não às 21h!)

# Arquivo para salvar progresso
CHECKPOINT_FILE = "progresso.json"

# ==================== GERENCIAMENTO DE CHECKPOINT ====================
def carregar_progresso():
    """
    Carrega o progresso salvo (última linha processada)
    Retorna a linha de onde deve continuar
    """
    if os.path.exists(CHECKPOINT_FILE):
        try:
            with open(CHECKPOINT_FILE, 'r') as f:
                dados = json.load(f)
                ultima_linha = dados.get('ultima_linha', 5)
                print(f"📍 Checkpoint encontrado! Continuando da linha {ultima_linha + 1}")
                return ultima_linha + 1  # Próxima linha a processar
        except:
            print("⚠️ Erro ao ler checkpoint. Começando do início.")
            return 5
    else:
        print("📝 Nenhum checkpoint encontrado. Começando do início (linha 6).")
        return 5

def salvar_progresso(linha_atual):
    """
    Salva o progresso atual (última linha processada)
    """
    agora_brasilia = datetime.datetime.now(FUSO_HORARIO)
    dados = {
        'ultima_linha': linha_atual,
        'data_hora': agora_brasilia.strftime("%d/%m/%Y %H:%M:%S")
    }
    with open(CHECKPOINT_FILE, 'w') as f:
        json.dump(dados, f, indent=2)
    print(f"💾 Progresso salvo: linha {linha_atual}")

def limpar_progresso():
    """
    Remove o arquivo de checkpoint (quando terminar todas as mensagens)
    """
    if os.path.exists(CHECKPOINT_FILE):
        os.remove(CHECKPOINT_FILE)
        print("✅ Checkpoint removido - todas as mensagens foram enviadas!")

def esta_no_horario_permitido():
    """
    Verifica se está dentro do horário permitido (8h às 22h) - Horário de Brasília
    """
    # Pega hora atual no fuso de Brasília
    agora_brasilia = datetime.datetime.now(FUSO_HORARIO)
    hora_atual = agora_brasilia.hour
    
    # Debug - mostra hora atual
    print(f"🕐 Hora atual (Brasília): {agora_brasilia.strftime('%H:%M:%S')}")
    
    esta_no_horario = HORA_INICIO <= hora_atual < HORA_FIM
    
    if not esta_no_horario:
        print(f"⚠️ Fora do horário permitido ({HORA_INICIO}h - {HORA_FIM}h)")
    
    return esta_no_horario

def aguardar_proximo_horario():
    """
    Aguarda até o próximo horário permitido (8h do próximo dia) - Horário de Brasília
    """
    agora = datetime.datetime.now(FUSO_HORARIO)
    
    # Se já passou da hora de fim, aguarda até hora de início do próximo dia
    if agora.hour >= HORA_FIM:
        proximo_inicio = agora.replace(hour=HORA_INICIO, minute=0, second=0, microsecond=0) + datetime.timedelta(days=1)
    else:
        # Se for antes da hora de início, aguarda até hora de início de hoje
        proximo_inicio = agora.replace(hour=HORA_INICIO, minute=0, second=0, microsecond=0)
    
    tempo_espera = (proximo_inicio - agora).total_seconds()
    
    print(f"\n⏰ Fora do horário permitido ({HORA_INICIO}h - {HORA_FIM}h)")
    print(f"🕐 Hora atual (Brasília): {agora.strftime('%d/%m/%Y %H:%M:%S')}")
    print(f"⏳ Aguardando até {proximo_inicio.strftime('%d/%m/%Y %H:%M')}")
    print(f"   (aproximadamente {int(tempo_espera / 3600)} horas e {int((tempo_espera % 3600) / 60)} minutos)")
    
    sleep(tempo_espera)

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
    """Envia mensagem usando pywhatkit"""
    try:
        telefone_limpo = ''.join(filter(str.isdigit, telefone))
        
        if not telefone_limpo.startswith('55'):
            telefone_limpo = '55' + telefone_limpo
        
        print(f"📤 Preparando envio para {nome} ({telefone_limpo})...")
        
        kit.sendwhatmsg_instantly(
            phone_no=f'+{telefone_limpo}',
            message=mensagem,
            wait_time=45,
            tab_close=True,
            close_time=5
        )
        
        print(f"✅ Mensagem enviada para {nome} ({telefone})")
        return True
        
    except Exception as e:
        print(f"❌ Erro ao enviar para {nome} ({telefone}): {str(e)}")
        return False

# ==================== REGISTRAR ERRO ====================
def registrar_erro(nome, telefone, erro):
    """Registra erros em arquivo CSV"""
    agora_brasilia = datetime.datetime.now(FUSO_HORARIO)
    data_hora = agora_brasilia.strftime("%d/%m/%Y %H:%M:%S")
    
    if not os.path.exists('erros.csv'):
        with open('erros.csv', 'w', encoding='utf-8') as arquivo:
            arquivo.write('Data/Hora,Nome,Telefone,Erro\n')
    
    with open('erros.csv', 'a', encoding='utf-8') as arquivo:
        arquivo.write(f'{data_hora},{nome},{telefone},{erro}\n')

# ==================== FUNÇÃO PRINCIPAL ====================
def main():
    print("="*70)
    print("🤖 ROBÔ DE ENVIO DE MENSAGENS - WhatsApp com Checkpoint")
    print("="*70)
    print(f"\n⏰ Horário de funcionamento: {HORA_INICIO}h às {HORA_FIM}h")
    print("💾 Sistema de checkpoint ativo (continua de onde parou)")
    print("\n⚠️ ATENÇÃO:")
    print("   1. Certifique-se de estar LOGADO no WhatsApp Web")
    print("   2. O robô vai parar às 21h e retomar às 8h automaticamente")
    print("   3. Mantenha o script rodando (use screen ou deixe o terminal aberto)")
    print("="*70)
    
    input("\n✋ Pressione ENTER para começar...")
    
    # Loop principal que roda continuamente
    while True:
        # Verifica se está no horário permitido
        if not esta_no_horario_permitido():
            aguardar_proximo_horario()
            continue
        
        # Autenticar Google
        print("\n📊 Conectando ao Google Sheets...")
        creds = autenticar_google()
        
        # Buscar dados
        values = buscar_dados_planilha(creds)
        if not values:
            print("⚠️ Nenhum dado na planilha. Aguardando próximo dia...")
            aguardar_proximo_horario()
            continue
        
        # Carregar progresso (de onde parou)
        linha_inicial = carregar_progresso()
        
        # Enviar mensagens
        print("\n📤 Iniciando envio de mensagens...\n")
        enviadas = 0
        erros_count = 0
        
        # Processa da linha salva até o final
        for i in range(linha_inicial, len(values)):
            # IMPORTANTE: Verifica horário antes de cada envio
            if not esta_no_horario_permitido():
                print(f"\n🕐 Horário limite atingido ({HORA_FIM}h)!")
                print(f"💾 Salvando progresso na linha {i}...")
                salvar_progresso(i)
                print("😴 Pausando até amanhã às 8h...")
                aguardar_proximo_horario()
                break  # Sai do loop e reinicia do checkpoint amanhã
            
            row = values[i]
            
            # Extrai dados (ajuste os índices conforme sua planilha)
            telefone = row[3].strip() if len(row) > 3 else ""
            nome = row[2].strip() if len(row) > 2 else "Sem nome"
            mensagem = row[4] if len(row) > 1 else ""  # Ajuste o índice da mensagem
            
            # Valida dados obrigatórios
            if not telefone or not mensagem:
                print(f"⚠️ Linha {i+1}: Dados incompletos")
                salvar_progresso(i)  # Salva mesmo se pular
                continue
            
            # Envia mensagem
            print(f"\n📍 Processando linha {i+1} de {len(values)}")
            sucesso = enviar_mensagem(telefone, mensagem, nome)
            
            if sucesso:
                enviadas += 1
            else:
                erros_count += 1
                registrar_erro(nome, telefone, "Falha no envio")
            
            # Salva progresso após cada envio
            salvar_progresso(i)
            
            # Pausa entre envios
            print(f"⏳ Aguardando {INTERVALO_ENTRE_MENSAGENS}s antes do próximo...")
            sleep(INTERVALO_ENTRE_MENSAGENS)
        
        # Se chegou aqui, terminou todas as linhas!
        if i >= len(values) - 1:
            print("\n" + "="*70)
            print("🎉 TODAS AS MENSAGENS FORAM ENVIADAS!")
            print("="*70)
            print(f"✅ Mensagens enviadas: {enviadas}")
            print(f"❌ Erros: {erros_count}")
            print(f"📝 Total processado: {enviadas + erros_count}")
            
            if erros_count > 0:
                print(f"\n⚠️ Verifique 'erros.csv' para detalhes dos erros")
            
            # Remove checkpoint pois terminou
            limpar_progresso()
            
            print("\n✅ Processo totalmente finalizado!")
            break  # Sai do loop principal e encerra o programa

# ==================== EXECUTAR ====================
if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\n⚠️ Processo interrompido pelo usuário.")
        print("💾 O progresso foi salvo. Execute novamente para continuar de onde parou.")
    except Exception as e:
        print(f"\n\n❌ Erro crítico: {e}")
        print("💾 O progresso foi salvo. Execute novamente para continuar de onde parou.")
