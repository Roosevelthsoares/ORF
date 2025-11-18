from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import time
import json
import requests
import pandas as pd


# ============================================================
# 1) CONFIGURAÇÕES DO LOGIN
# ============================================================
USER_EMAIL = "fiscal.pulsar@4cta.eb.mil.br"
USER_PASSWORD = "K809(F4a[?"

LOGIN_URL = "https://sport.pulsarconnect.io/login"
STARLINK_URL = "https://sport.pulsarconnect.io/starlink/starlinkMap?sideNav=true&interval=MTD&startDate=1761955200000&endDate=1763408670718&selectedAccount=All&serviceLineAccess=All&starlinkTab=true&page=1&size=25&sortBy=usageCB&sortOrder=desc&search="

# ============================================================
# 2) ABRIR SELENIUM (HEADLESS) COM LOGGING
# ============================================================
options = Options()
# Desabilitando headless temporariamente para debug - reative depois
# options.add_argument("--headless=new")  # modo invisível
options.add_argument("--disable-gpu")
options.add_argument("--window-size=1920,1080")
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")

# Habilitar logging de performance para capturar requisições
options.set_capability('goog:loggingPrefs', {'performance': 'ALL'})

driver = webdriver.Chrome(
    service=Service(ChromeDriverManager().install()),
    options=options
)

print("➡️ Abrindo página de login...")
driver.get(LOGIN_URL)

# ============================================================
# 3) DIGITAR LOGIN E SENHA
# ============================================================
print("➡️ Aguardando página carregar...")

# Espera explícita até o campo de email aparecer (máximo 20 segundos)
wait = WebDriverWait(driver, 20)

# Aguardar a página carregar completamente
time.sleep(5)

try:
    # Encontrar campos usando os nomes corretos
    email_input = wait.until(EC.presence_of_element_located((By.NAME, "userName")))
    pass_input = wait.until(EC.presence_of_element_located((By.NAME, "password")))
    
    print("✔ Campos de login encontrados!")
    print("➡️ Preenchendo login...")
    
    email_input.send_keys(USER_EMAIL)
    pass_input.send_keys(USER_PASSWORD)
    
    # Encontrar e clicar no botão de login
    login_button = driver.find_element(By.XPATH, "//button[contains(@class, 'loginButton')]")
    login_button.click()
    
    # aguardar redirecionamento (mais tempo em headless)
    print("➡️ Aguardando login...")
    time.sleep(12)
    
except Exception as e:
    print(f"❌ Erro ao fazer login: {e}")
    
    # Salvar screenshot para debug
    driver.save_screenshot("erro_login.png")
    print("📸 Screenshot salvo: erro_login.png")
    
    driver.quit()
    exit()

# ============================================================
# 4) PEGAR TOKEN DO LOCALSTORAGE
# ============================================================
print("➡️ Obtendo token...")

# Aguardar mais tempo para o login completar e token ser armazenado
max_tentativas = 10
logged_in_user = None

for tentativa in range(max_tentativas):
    logged_in_user = driver.execute_script(
        "return window.localStorage.getItem('loggedInUser');"
    )
    
    if logged_in_user:
        print("✔ Token encontrado!")
        break
    
    print(f"⏳ Aguardando token... (tentativa {tentativa + 1}/{max_tentativas})")
    time.sleep(2)

if not logged_in_user:
    print("❌ Token não encontrado no localStorage")
    print("Verificando todos os items do localStorage:")
    all_storage = driver.execute_script(
        "return JSON.stringify(window.localStorage);"
    )
    print(all_storage)
    
    # Verificar se há erro de login
    driver.save_screenshot("erro_token.png")
    print("📸 Screenshot salvo: erro_token.png")
    
    driver.quit()
    exit()

# converter em JSON
logged_data = json.loads(logged_in_user)
access_token = logged_data["data"]["access_token"]
print("✔ Token extraído com sucesso!")

# ============================================================
# 5) NAVEGAR PARA A PÁGINA DO STARLINK
# ============================================================
print("➡️ Navegando para a página do Starlink...")
driver.get(STARLINK_URL)

# Aguardar a página carregar BASTANTE tempo para todos os dados aparecerem
print("⏳ Aguardando 20 segundos para a página carregar completamente...")
time.sleep(20)

# Habilitar log de rede para capturar requisições
driver.execute_cdp_cmd('Network.enable', {})

print("➡️ Aguardando requisições XHR com dados...")

# Aguardar mais um pouco para garantir que as requisições sejam feitas
time.sleep(5)

# Procurar e clicar no botão de reload/refresh na página
print("➡️ Procurando botão de reload/refresh...")
try:
    # Tentar encontrar botão de reload (vários possíveis seletores)
    possible_selectors = [
        "//button[contains(@aria-label, 'reload')]",
        "//button[contains(@aria-label, 'refresh')]",
        "//button[contains(@title, 'Reload')]",
        "//button[contains(@title, 'Refresh')]",
        "//button[contains(., 'Reload')]",
        "//button[contains(., 'Refresh')]",
        "//button[contains(@class, 'reload')]",
        "//button[contains(@class, 'refresh')]",
        "//*[contains(@class, 'MuiIconButton')][contains(@aria-label, 'reload')]",
        "//*[@data-testid='refresh-button']",
    ]
    
    reload_button = None
    for selector in possible_selectors:
        try:
            reload_button = driver.find_element(By.XPATH, selector)
            if reload_button:
                print(f"✔ Botão de reload encontrado!")
                reload_button.click()
                print("✔ Botão clicado!")
                time.sleep(8)
                break
        except:
            continue
    
    if not reload_button:
        print("⚠ Botão de reload não encontrado, usando reload do navegador...")
        driver.execute_script("window.location.reload();")
        time.sleep(8)
        
except Exception as e:
    print(f"⚠ Erro ao procurar botão: {e}")
    driver.execute_script("window.location.reload();")
    time.sleep(8)

# Rolar a página para disparar carregamento de dados
print("➡️ Interagindo com a página...")
driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
time.sleep(3)
driver.execute_script("window.scrollTo(0, 0);")
time.sleep(3)

# Tentar pegar os logs de rede
logs = driver.get_log('performance')

# Filtrar requisições XHR relevantes (APIs que retornam JSON)
api_calls = []
for entry in logs:
    try:
        log = json.loads(entry['message'])['message']
        
        # Capturar tanto requisições quanto respostas
        if 'Network.requestWillBeSent' in log['method']:
            request = log['params'].get('request', {})
            url = request.get('url', '')
            
            # Capturar requisições para APIs de dados
            if 'api.k4mobility.com' in url and ('/query' in url or '/starlink' in url):
                request_data = {
                    'url': url,
                    'method': request.get('method', 'GET'),
                    'postData': request.get('postData', None),
                    'requestId': log['params'].get('requestId', '')
                }
                
                # Evitar duplicatas
                if not any(api['url'] == url for api in api_calls):
                    api_calls.append(request_data)
                    
                    if '/query' in url or 'collectv=' in url:
                        print(f"⭐ REQUISIÇÃO de dados: {request.get('method', 'GET')} {url[:100]}...")
                        
        elif 'Network.responseReceived' in log['method']:
            if 'response' in log['params']:
                response = log['params']['response']
                url = response.get('url', '')
                mime_type = response.get('mimeType', '')
                
                # Capturar respostas de APIs que retornam JSON
                if 'api.k4mobility.com' in url and 'application/json' in mime_type:
                    # Verificar se já foi capturada como requisição
                    existing = next((api for api in api_calls if api['url'] == url), None)
                    if not existing:
                        api_calls.append({
                            'url': url,
                            'method': 'GET',
                            'postData': None,
                            'requestId': log['params'].get('requestId', ''),
                            'mimeType': mime_type
                        })
                        
                        if '/query' in url or 'collectv=' in url:
                            print(f"⭐ RESPOSTA de dados: {url[:100]}...")
                        else:
                            print(f"🔍 API encontrada: {url[:100]}...")
    except Exception as e:
        continue

print(f"✔ {len(api_calls)} requisições de API encontradas")

# ============================================================
# 5.5) EXTRAIR DADOS DA TABELA VISÍVEL NA PÁGINA
# ============================================================
print("\n➡️ Extraindo dados da tabela na página...")

all_data = []

try:
    # Aguardar a tabela carregar
    wait = WebDriverWait(driver, 30)
    
    # Localizar todas as linhas da tabela
    # Procurar por linhas que contenham SERVICE LINE e STATUS
    rows = driver.find_elements(By.XPATH, "//tr[contains(@class, 'MuiTableRow') or contains(@role, 'row')]")
    
    print(f"🔍 {len(rows)} linhas encontradas na tabela")
    
    for row in rows:
        try:
            # Extrair células da linha
            cells = row.find_elements(By.TAG_NAME, "td")
            
            if len(cells) < 3:  # Precisa ter pelo menos 3 colunas
                continue
            
            # 1ª coluna: SERVICE LINE (nome do terminal)
            service_line_text = cells[0].text.strip()
            
            if not service_line_text or service_line_text == "SERVICE LINE":
                continue
            
            # 2ª coluna: STATUS (com os círculos)
            status_cell = cells[1]
            
            # Procurar pelo KIT ID dentro da célula de status
            kit_elements = status_cell.find_elements(By.XPATH, ".//*[contains(text(), 'KIT')]")
            kit_id = ""
            if kit_elements:
                kit_id = kit_elements[0].text.strip()
            
            # Procurar pelos círculos de status (svg ou span com cores)
            circles = status_cell.find_elements(By.XPATH, ".//span[contains(@class, 'MuiSvgIcon') or @role='img'] | .//svg | .//*[contains(@style, 'background') or contains(@style, 'color')]")
            
            # Pegar especificamente o 2º círculo (índice 1)
            second_circle_status = "DESCONHECIDO"
            if len(circles) >= 2:
                second_circle = circles[1]
                
                # Verificar a cor (pode estar no style, class, ou como SVG path)
                circle_html = second_circle.get_attribute('outerHTML')
                style = second_circle.get_attribute('style') or ''
                class_name = second_circle.get_attribute('class') or ''
                
                # Determinar se é verde ou vermelho
                if 'green' in style.lower() or 'green' in class_name.lower() or '#00ff00' in style.lower() or 'rgb(0' in style.lower():
                    second_circle_status = "VERDE"
                elif 'red' in style.lower() or 'red' in class_name.lower() or '#ff0000' in style.lower() or 'rgb(255, 0' in style.lower():
                    second_circle_status = "VERMELHO"
                else:
                    # Tentar verificar pelo SVG path fill
                    svg_paths = second_circle.find_elements(By.TAG_NAME, "path")
                    for path in svg_paths:
                        fill = path.get_attribute('fill') or ''
                        if 'green' in fill.lower() or '#0f0' in fill or '#00ff00' in fill:
                            second_circle_status = "VERDE"
                            break
                        elif 'red' in fill.lower() or '#f00' in fill or '#ff0000' in fill:
                            second_circle_status = "VERMELHO"
                            break
            
            # 3ª coluna: USAGE (GB)/PLAN
            usage_text = ""
            if len(cells) >= 3:
                usage_text = cells[2].text.strip()
            
            # Adicionar aos dados
            if service_line_text:
                all_data.append({
                    'SERVICE_LINE': service_line_text,
                    'KIT_ID': kit_id,
                    'STATUS_2ND_CIRCLE': second_circle_status,
                    'USAGE_GB_PLAN': usage_text,
                    'OBSERVACAO': ''  # Coluna em branco para preenchimento manual
                })
                
        except Exception as e:
            continue
    
    print(f"✔ {len(all_data)} registros extraídos da tabela!")
    
except Exception as e:
    print(f"❌ Erro ao extrair tabela: {e}")

# Se não conseguiu dados da tabela, tentar pelas APIs
if not all_data:
    print("\n⚠ Tentando obter dados das APIs...")
    
    # Salvar URLs para debug
    with open("debug_urls.txt", "w", encoding="utf-8") as f:
        for api_call in api_calls:
            f.write(f"{api_call['url']}\n\n")
    print(f"📝 URLs salvas em debug_urls.txt")
    
    # Código anterior de APIs...
    all_data_api = []
    
    for api_call in api_calls:
    try:
        # Fazer a requisição com o token
        headers = {
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json"
        }
        
        # Usar POST se houver postData, senão GET
        if api_call.get('method') == 'POST' and api_call.get('postData'):
            response = requests.post(api_call['url'], headers=headers, data=api_call['postData'])
        else:
            response = requests.get(api_call['url'], headers=headers)
        
        if response.status_code == 200:
            data = response.json()
            
            # Debug: verificar estrutura dos dados
            print(f"\n🔍 Analisando: {api_call['url'][:80]}...")
            
            # Verificar se há dados úteis
            if isinstance(data, dict):
                print(f"   Chaves encontradas: {list(data.keys())}")
                
                if 'data' in data:
                    if isinstance(data['data'], list) and len(data['data']) > 0:
                        all_data.extend(data['data'])
                        print(f"✔ {len(data['data'])} registros obtidos!")
                    elif isinstance(data['data'], dict):
                        # Se data é um dict, pode ter listas dentro
                        for key, value in data['data'].items():
                            if isinstance(value, list) and len(value) > 0:
                                all_data.extend(value)
                                print(f"✔ {len(value)} registros obtidos de data['{key}']!")
                                
                elif 'results' in data and isinstance(data['results'], list):
                    all_data.extend(data['results'])
                    print(f"✔ {len(data['results'])} registros obtidos de 'results'!")
                    
                # Verificar outras possíveis chaves com arrays
                for key in ['users', 'accounts', 'terminals', 'serviceLines', 'items']:
                    if key in data and isinstance(data[key], list) and len(data[key]) > 0:
                        all_data.extend(data[key])
                        print(f"✔ {len(data[key])} registros obtidos de '{key}'!")
                        
            elif isinstance(data, list) and len(data) > 0:
                all_data.extend(data)
                print(f"✔ {len(data)} registros obtidos (lista direta)!")
                
    except Exception as e:
        print(f"⚠ Erro ao processar {api_call['url'][:60]}: {e}")
        continue

driver.quit()

# ============================================================
# 6) EXPORTAR PARA CSV/EXCEL
# ============================================================
if all_data:
    df = pd.DataFrame(all_data)
    print(f"✔ Total de {len(df)} registros coletados!")
else:
    print("⚠ Nenhum dado foi coletado. Salvando estrutura vazia...")
    df = pd.DataFrame()

df.to_csv("starlink_dados.csv", index=False)
df.to_excel("starlink_dados.xlsx", index=False)

print("📁 Arquivos salvos:")
print("- starlink_dados.csv")
print("- starlink_dados.xlsx")
print(f"Colunas: {list(df.columns) if len(df.columns) > 0 else 'Nenhuma'}")
