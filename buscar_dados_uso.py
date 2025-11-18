"""
Script alternativo: Busca dados de uso do Starlink via API
Baseado nos dados já coletados de serviceLines
"""
import pandas as pd
import requests
import json
from datetime import datetime, timedelta

# Ler o CSV com os serviceLines já coletados
df_service_lines = pd.read_csv("starlink_dados.csv")

print(f"✔ {len(df_service_lines)} service lines encontradas")
print(f"Colunas: {list(df_service_lines.columns)}")

# Fazer login para obter token
USER_EMAIL = "fiscal.pulsar@4cta.eb.mil.br"
USER_PASSWORD = "K809(F4a[?"

login_url = "https://api.k4mobility.com/iam/v3/user/login"
login_payload = {
    "username": USER_EMAIL,
    "password": USER_PASSWORD  
}

print("\n➡️ Fazendo login na API...")
response = requests.post(login_url, json=login_payload)

if response.status_code == 200:
    login_data = response.json()
    access_token = login_data["data"]["access_token"]
    print("✔ Token obtido!")
    
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }
    
    # Tentar buscar dados de uso para cada service line
    print("\n➡️ Buscando dados de uso/consumo...")
    
    all_usage_data = []
    
    for idx, row in df_service_lines.head(5).iterrows():  # Testar com 5 primeiros
        service_line_id = row['id']
        name = row['name']
        
        print(f"\n📡 {name} ({service_line_id})")
        
        # Tentar diferentes endpoints para dados de uso
        usage_endpoints = [
            f"https://api.k4mobility.com/starlink/usage/{service_line_id}",
            f"https://api.k4mobility.com/starlink/serviceLine/{service_line_id}/usage",
            f"https://api.k4mobility.com/starlink/serviceLine/{service_line_id}/metrics",
            f"https://api.k4mobility.com/starlink/serviceLine/{service_line_id}/data",
        ]
        
        for endpoint in usage_endpoints:
            try:
                resp = requests.get(endpoint, headers=headers, timeout=5)
                if resp.status_code == 200:
                    data = resp.json()
                    print(f"   ✔ Dados encontrados em: {endpoint.split('/')[-2:]}")
                    all_usage_data.append({
                        'service_line_id': service_line_id,
                        'name': name,
                        'data': data
                    })
                    break
            except Exception as e:
                continue
    
    if all_usage_data:
        print(f"\n✔ {len(all_usage_data)} service lines com dados de uso!")
        # Salvar dados
        with open("starlink_usage_data.json", "w", encoding="utf-8") as f:
            json.dump(all_usage_data, f, indent=2, ensure_ascii=False)
        print("📁 Dados salvos em: starlink_usage_data.json")
    else:
        print("\n⚠ Nenhum dado de uso encontrado nos endpoints testados")
        print("Pode ser necessário usar o método de captura via navegador")
        
else:
    print(f"❌ Erro no login: {response.status_code}")
    print(response.text)
