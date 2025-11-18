import requests
import json
import pandas as pd
from test import USER_EMAIL, USER_PASSWORD

# Fazer login para obter token
login_url = "https://api.k4mobility.com/iam/v3/user/login"
login_payload = {
    "username": USER_EMAIL,
    "password": USER_PASSWORD
}

print("➡️ Fazendo login na API...")
response = requests.post(login_url, json=login_payload)

if response.status_code == 200:
    login_data = response.json()
    access_token = login_data["data"]["access_token"]
    print("✔ Token obtido!")
    
    headers = {"Authorization": f"Bearer {access_token}"}
    
    # Tentar diferentes endpoints da API Starlink
    endpoints = [
        "https://api.k4mobility.com/starlink/serviceLine/dpId/DP-0833",
        "https://api.k4mobility.com/starlink/usage",
        "https://api.k4mobility.com/starlink/data",
        "https://api.k4mobility.com/starlink/metrics",
        "https://api.k4mobility.com/store/pg/query",
        "https://api.k4mobility.com/store/ch/query",
    ]
    
    for endpoint in endpoints:
        print(f"\n🔍 Testando: {endpoint}")
        try:
            resp = requests.get(endpoint, headers=headers)
            print(f"   Status: {resp.status_code}")
            if resp.status_code == 200:
                data = resp.json()
                print(f"   Chaves: {list(data.keys()) if isinstance(data, dict) else 'Lista'}")
                if isinstance(data, dict) and 'data' in data:
                    print(f"   data.keys(): {list(data['data'].keys()) if isinstance(data['data'], dict) else f'{len(data['data'])} items'}")
        except Exception as e:
            print(f"   Erro: {e}")
else:
    print(f"❌ Erro no login: {response.status_code}")
    print(response.text)
