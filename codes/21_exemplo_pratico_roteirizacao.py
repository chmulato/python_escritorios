# ================================================================
# Exemplo Prático 21: Roteirização com base em distância (API Google Maps ou OpenRoute)
#
# Este script calcula a rota entre dois pontos usando a API do Google Maps
# ou a API do OpenRoute, exibindo distância e tempo estimado.
#
# Autor: Christian Mulato
# Data: Junho de 2025
#
# ATENÇÃO:
# - pip install requests
# - Substitua as chaves de API pelos seus dados reais.
# ================================================================
import requests

# Função para calcular a rota utilizando a API do Google Maps
def calcular_rota_google_maps(origem, destino, chave_api):
    url = f"https://maps.googleapis.com/maps/api/directions/json?origin={origem}&destination={destino}&key={chave_api}"
    response = requests.get(url)
    dados = response.json()
    if dados['status'] == 'OK':
        rota = dados['routes'][0]
        distancia = rota['legs'][0]['distance']['text']
        tempo = rota['legs'][0]['duration']['text']
        return f"Distância: {distancia}, Tempo: {tempo}"
    else:
        return "Erro ao calcular a rota."

# Função para calcular a rota utilizando a API do OpenRoute
def calcular_rota_openroute(origem, destino, chave_api):
    url = f"https://api.openrouteservice.org/v2/directions/driving-car?api_key={chave_api}&start={origem}&end={destino}"
    response = requests.get(url)
    dados = response.json()
    if 'routes' in dados and dados['routes']:
        rota = dados['routes'][0]
        distancia = rota['summary']['distance'] / 1000  # metros para km
        tempo = rota['summary']['duration'] / 60  # segundos para minutos
        return f"Distância: {distancia:.2f} km, Tempo: {tempo:.1f} min"
    else:
        return "Erro ao calcular a rota."

# Exemplo de uso
origem = "-46.633308,-23.550520"  # São Paulo (longitude,latitude)
destino = "-43.172896,-22.906847"  # Rio de Janeiro (longitude,latitude)

# Google Maps
chave_api_google = 'sua_chave_api_google_maps'
resultado_google_maps = calcular_rota_google_maps(origem, destino, chave_api_google)
print("Google Maps:", resultado_google_maps)

# OpenRoute
chave_api_openroute = 'sua_chave_api_openroute'
resultado_openroute = calcular_rota_openroute(origem, destino, chave_api_openroute)
print("OpenRoute:", resultado_openroute)
