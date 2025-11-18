from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import time
import pandas as pd

# ============================================================
# CONFIGURAÇÕES
# ============================================================
USER_EMAIL = "fiscal.pulsar@4cta.eb.mil.br"
USER_PASSWORD = "K809(F4a[?"
LOGIN_URL = "https://sport.pulsarconnect.io/login"
STARLINK_URL = "https://sport.pulsarconnect.io/starlink/starlinkMap?sideNav=true&interval=MTD&startDate=1761955200000&endDate=1763408670718&selectedAccount=All&serviceLineAccess=All&starlinkTab=true&page=1&size=25&sortBy=usageCB&sortOrder=desc&search="

# ============================================================
# ABRIR NAVEGADOR
# ============================================================
options = Options()
# Descomente para modo invisível após testar:
# options.add_argument("--headless=new")
options.add_argument("--disable-gpu")
options.add_argument("--window-size=1920,1080")
options.add_argument("--no-sandbox")

driver = webdriver.Chrome(
    service=Service(ChromeDriverManager().install()),
    options=options
)

print("➡️ Abrindo página de login...")
driver.get(LOGIN_URL)

# ============================================================
# FAZER LOGIN
# ============================================================
wait = WebDriverWait(driver, 20)
time.sleep(5)

try:
    print("➡️ Fazendo login...")
    email_input = wait.until(EC.presence_of_element_located((By.NAME, "userName")))
    pass_input = wait.until(EC.presence_of_element_located((By.NAME, "password")))
    
    email_input.send_keys(USER_EMAIL)
    pass_input.send_keys(USER_PASSWORD)
    
    login_button = driver.find_element(By.XPATH, "//button[contains(@class, 'loginButton')]")
    login_button.click()
    
    print("⏳ Aguardando login...")
    time.sleep(12)
    
except Exception as e:
    print(f"❌ Erro no login: {e}")
    driver.save_screenshot("erro_login.png")
    driver.quit()
    exit()

# ============================================================
# NAVEGAR PARA STARLINK
# ============================================================
print("➡️ Navegando para página do Starlink...")
driver.get(STARLINK_URL)

print("⏳ Aguardando página carregar...")
time.sleep(15)

# Tentar clicar em elementos que possam carregar os dados
print("➡️ Procurando por botões/filtros que carregam dados...")
try:
    # Procurar por botão de "Apply", "Search", "Load", etc
    possible_buttons = driver.find_elements(By.XPATH, 
        "//button[contains(text(), 'Apply')] | //button[contains(text(), 'Search')] | //button[contains(text(), 'Load')] | //button[contains(@aria-label, 'search')] | //button[contains(@type, 'submit')]")
    
    if possible_buttons:
        print(f"✔ {len(possible_buttons)} botões encontrados, clicando...")
        for btn in possible_buttons[:2]:  # Clicar nos 2 primeiros
            try:
                btn.click()
                time.sleep(3)
            except:
                pass
except Exception as e:
    print(f"⚠ Nenhum botão encontrado: {e}")

# Aguardar mais tempo para dados carregarem
print("⏳ Aguardando dados carregarem (mais 20 segundos)...")
time.sleep(20)

# Tentar scroll para disparar lazy loading
driver.execute_script("window.scrollTo(0, 500);")
time.sleep(2)
driver.execute_script("window.scrollTo(0, 0);")
time.sleep(2)

# ============================================================
# EXTRAIR DADOS DA TABELA
# ============================================================
print("\n📊 Extraindo dados da tabela...")

# Salvar HTML para debug
print("📄 Salvando HTML da página...")
with open("pagina_starlink.html", "w", encoding="utf-8") as f:
    f.write(driver.page_source)
print("✔ HTML salvo em: pagina_starlink.html")

driver.save_screenshot("pagina_starlink.png")
print("✔ Screenshot salvo em: pagina_starlink.png")

all_data = []

try:
    # Procurar por todas as linhas da tabela
    # Múltiplos seletores possíveis para maior compatibilidade
    row_selectors = [
        "//tbody/tr",
        "//tr[@role='row']",
        "//div[contains(@class, 'MuiTableRow')]",
        "//tr[contains(@class, 'MuiTableRow')]"
    ]
    
    rows = []
    for selector in row_selectors:
        try:
            found_rows = driver.find_elements(By.XPATH, selector)
            if found_rows and len(found_rows) > 0:
                rows = found_rows
                print(f"✔ {len(rows)} linhas encontradas usando: {selector}")
                break
        except:
            continue
    
    if not rows:
        print("⚠ Nenhuma linha encontrada. Salvando HTML para debug...")
        with open("pagina_starlink.html", "w", encoding="utf-8") as f:
            f.write(driver.page_source)
        print("📄 HTML salvo em: pagina_starlink.html")
    
    for idx, row in enumerate(rows):
        try:
            # Extrair todas as células
            cells = row.find_elements(By.TAG_NAME, "td")
            
            if len(cells) < 2:
                continue
            
            # 1ª coluna: SERVICE LINE
            service_line = cells[0].text.strip()
            
            # Pular cabeçalhos
            if not service_line or service_line.upper() == "SERVICE LINE" or len(service_line) < 3:
                continue
            
            # 2ª coluna: STATUS (com os círculos e KIT ID)
            status_cell = cells[1]
            
            # Procurar KIT ID
            kit_id = ""
            kit_pattern_elements = status_cell.find_elements(By.XPATH, ".//*[contains(text(), 'KIT')]")
            if kit_pattern_elements:
                kit_id = kit_pattern_elements[0].text.strip()
            
            # Procurar círculos de status
            # Círculos podem ser: SVG circles, spans coloridos, ou ícones
            status_indicators = status_cell.find_elements(By.XPATH, 
                ".//span[contains(@class, 'MuiSvgIcon')] | .//svg | .//*[local-name()='circle'] | .//span[contains(@class, 'icon')] | .//span[contains(@style, 'background')]")
            
            second_circle_status = "DESCONHECIDO"
            
            # Analisar o 2º indicador (índice 1)
            if len(status_indicators) >= 2:
                second_indicator = status_indicators[1]
                
                # Obter informações sobre a cor
                outer_html = second_indicator.get_attribute('outerHTML')
                style_attr = second_indicator.get_attribute('style') or ''
                class_attr = second_indicator.get_attribute('class') or ''
                
                # Análise de cor
                html_lower = (outer_html + style_attr + class_attr).lower()
                
                # Verde: vários tons possíveis
                green_indicators = ['green', '#0f0', '#00ff00', '#00e676', '#4caf50', 'rgb(0,', 'success', 'online']
                # Vermelho: vários tons possíveis  
                red_indicators = ['red', '#f00', '#ff0000', '#f44336', '#e53935', 'rgb(255, 0', 'error', 'offline', 'alert']
                
                if any(indicator in html_lower for indicator in green_indicators):
                    second_circle_status = "VERDE"
                elif any(indicator in html_lower for indicator in red_indicators):
                    second_circle_status = "VERMELHO"
            
            # 3ª coluna: USAGE
            usage = ""
            if len(cells) >= 3:
                usage = cells[2].text.strip()
            
            # Adicionar dados
            row_data = {
                'SERVICE_LINE': service_line,
                'KIT_ID': kit_id,
                'STATUS': second_circle_status,
                'USAGE_GB_PLAN': usage,
                'OBSERVACAO': ''  # Para preenchimento manual
            }
            
            all_data.append(row_data)
            print(f"✔ Linha {idx+1}: {service_line[:30]} | {kit_id} | {second_circle_status}")
            
        except Exception as e:
            print(f"⚠ Erro na linha {idx}: {e}")
            continue
    
    print(f"\n✅ Total extraído: {len(all_data)} registros")
    
except Exception as e:
    print(f"❌ Erro ao extrair tabela: {e}")
    driver.save_screenshot("erro_extracao.png")

driver.quit()

# ============================================================
# SALVAR EM PLANILHA
# ============================================================
if all_data:
    df = pd.DataFrame(all_data)
    
    # Salvar em Excel e CSV
    df.to_excel("starlink_status_report.xlsx", index=False)
    df.to_csv("starlink_status_report.csv", index=False)
    
    print("\n📁 Arquivos gerados:")
    print("  ✔ starlink_status_report.xlsx")
    print("  ✔ starlink_status_report.csv")
    print(f"\n📊 Resumo:")
    print(f"  Total de registros: {len(df)}")
    if 'STATUS' in df.columns:
        print(f"  STATUS VERDE: {len(df[df['STATUS'] == 'VERDE'])}")
        print(f"  STATUS VERMELHO: {len(df[df['STATUS'] == 'VERMELHO'])}")
        print(f"  STATUS DESCONHECIDO: {len(df[df['STATUS'] == 'DESCONHECIDO'])}")
else:
    print("\n⚠ Nenhum dado foi extraído!")
    print("Verifique os arquivos de debug gerados.")
