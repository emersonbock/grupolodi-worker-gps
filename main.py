import requests
import time
import os
import pandas as pd
import math
import csv  # <-- Adicionado para gerar a planilha
from zoneinfo import ZoneInfo
from datetime import datetime

# --- CONFIGURAÇÃO ---
# É altamente recomendável usar variáveis de ambiente para dados sensíveis.
# Se preferir, pode colar os valores diretamente aqui, mas não é o ideal por segurança.

# 1. Informações do Telegram
TELEGRAM_BOT_NAME = "frota_grupolodi_bot"
TELEGRAM_BOT_TOKEN = os.getenv("TELEGRAM_BOT_TOKEN", "8339036107:AAHQ27Z9wxcZgSEKNErlrftDCHJuailCnoE")
TELEGRAM_CHAT_ID = os.getenv("TELEGRAM_CHAT_ID", "-1002876367400")

# 2. Informações da Requisição
# ATENÇÃO: É necessário ter uma sessão de login ativa.
URL_API = "https://seguro.linkmonitoramento.com.br/app/index_xml.php?app_modulo=central_monitoramento&app_comando=monitorar_grupo_veiculo"

HEADERS = {
    'Accept': 'application/json, text/javascript, */*; q=0.01',
    'Accept-Encoding': 'gzip, deflate, br, zstd',
    'Accept-Language': 'pt-BR,pt;q=0.8,en-US;q=0.5,en;q=0.3',
    'Connection': 'keep-alive',
    'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
    'Host': 'seguro.linkmonitoramento.com.br',
    'Origin': 'https://seguro.linkmonitoramento.com.br',
    'Referer': 'https://seguro.linkmonitoramento.com.br/app/',
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:141.0) Gecko/20100101 Firefox/141.0',
    'X-Requested-With': 'XMLHttpRequest',
}
USER_AGENTS = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:141.0) Gecko/20100101 Firefox/141.0'}

PAYLOAD = "id_grupo=0&tipo=atualizar"

BASE_DE_CIDADES = [
    {"nome": "Miritituba - PA", "coordenadas": [-4.299317867324192, -55.95680008343704], "raio_km": 1.75},
    {"nome": "Santarém - PA", "coordenadas": [-2.4513435658922713, -54.73016688472612], "raio_km": 5.75},
    {"nome": "Itaituba - PA", "coordenadas": [-4.253218906491497, -56.00047168576537], "raio_km": 3},
    {"nome": "Trairão BR163 - PA", "coordenadas": [-4.7022085347195635, -55.996299517899146], "raio_km": 1.25},
    {"nome": "KM-30 BR163 - PA", "coordenadas": [-4.346104541095031, -55.78528514498654], "raio_km": 4},
    {"nome": "Caracol BR163 - PA", "coordenadas": [-5.018077317571283, -56.186974385535606], "raio_km": 2.5},
    {"nome": "Castelo dos Sonhos - PA", "coordenadas": [-8.316628389688903, -55.10059496430152], "raio_km": 2},
    {"nome": "Novo Progresso - PA", "coordenadas": [-7.039034825060456, -55.41511479964626], "raio_km": 2},
    {"nome": "Morais de Almeida - PA", "coordenadas": [-6.221954444889193, -55.62672098933037], "raio_km": 2.15},

    {"nome": "Sinop - MT", "coordenadas": [-11.854303329070625, -55.5063449823773], "raio_km": 5.25},
    {"nome": "Cláudia - MT", "coordenadas": [-11.505328438533565, -54.8754462328951], "raio_km": 2.65},
    {"nome": "Nobres - MT", "coordenadas": [-14.72111385657074, -56.32280244532029], "raio_km": 3.25},
    {"nome": "Itaúba - MT", "coordenadas": [-11.008770639483659, -55.242263937213885], "raio_km": 1.5},
    {"nome": "Nova Santa Helena - MT", "coordenadas": [-10.84812668059357, -55.18516665997486], "raio_km": 1},
    {"nome": "Peixoto de Azevedo - MT", "coordenadas": [-10.245274763144353, -54.98972907399532], "raio_km": 2.5},
    {"nome": "Matupá - MT", "coordenadas": [-10.17098292782133, -54.93357751064267], "raio_km": 2.5},
    {"nome": "Guarantã do Norte - MT", "coordenadas": [-9.952927542196868, -54.90489044242835], "raio_km": 3},
    {"nome": "União do Sul - MT", "coordenadas": [-11.53340847131758, -54.365663182830595], "raio_km": 3.15},
    {"nome": "Santa Carmem - MT", "coordenadas": [-11.975878313370082, -55.27873347895608], "raio_km": 1.85},
    {"nome": "Sorriso - MT", "coordenadas": [-12.562164390824588, -55.72185910088999], "raio_km": 2.25},
    {"nome": "Lucas do Rio Verde - MT", "coordenadas": [-13.074967904557624, -55.92273534917346], "raio_km": 2.50},
    {"nome": "Rosário Oeste - MT", "coordenadas": [-14.824960433313441, -56.426966650907985], "raio_km": 2},
    {"nome": "Feliz Natal - MT", "coordenadas": [-12.379763834401508, -54.93336992971628], "raio_km": 3},
    {"nome": "Marcelândia - MT", "coordenadas": [-11.086559923906941, -54.516745614125874], "raio_km": 3.25}
]

BASE_DE_LOCAIS = [
    {"nome": "Edifício Grupo Lodi", "lat": -23.550520, "lon": -46.633308, "raio_km": 0.75},
    {"nome": "Faz. Esprança", "lat": -23.59543, "lon": -46.68378, "raio_km": 1},
    {"nome": "Faz. São João", "lat": -23.54135, "lon": -46.54992, "raio_km": 1.5},
    {"nome": "Faz. Vô Amantino", "lat": -11.662623527640909, "lon": -55.244362751210225, "raio_km": 1.5},
    {"nome": "BritaMix", "lat": -14.674496902043648, "lon": -56.29667359859004, "raio_km": 1.5},
    {"nome": "Fertitex - Santarém", "lat": -2.5296782188664935, "lon": -54.72329619624749, "raio_km": 0.55},
    {"nome": "Posto Trevão - Santarém", "lat": -2.47488620255129, "lon": -54.729219777438495, "raio_km": 0.35},
    {"nome": "Posto Ipiranga - Santarém", "lat": -2.6038870242083894, "lon": -54.72009588421867, "raio_km": 0.25},
    {"nome": "Campo Rico - Comércio de Fertilizantes", "lat": -2.4630038663356704, "lon": -54.72558902158942,
     "raio_km": 0.50},
    {"nome": "Porto de Santarém", "lat": -2.416450585287248, "lon": -54.7380990640844, "raio_km": 1.50},
    {"nome": "Hidrovias - Miritituba", "lat": -4.271238592885619, "lon": -55.94419927892871, "raio_km": 0.25},
    {"nome": "Cianport - Miritituba", "lat": -4.280770594625185, "lon": -55.95228721352993, "raio_km": 0.25},
    {"nome": "Posto Mirian - Miritituba", "lat": -4.339272203201447, "lon": -55.95828706036294, "raio_km": 0.45},
    {"nome": "Posto Trevão - Miritituba", "lat": -4.368824724446036, "lon": -55.96214945517425, "raio_km": 0.40},
    {"nome": "Arm. Real Agro", "lat": -11.514953053030165, "lon": -54.40088782597076, "raio_km": 0.50},
    {"nome": "Posto Trevão - Caracol", "lat": -5.032212569295955, "lon": -56.18759882258732, "raio_km": 0.45}

]


# --- FUNÇÕES ---

def iniciar_sessao_link(email, senha):
    """Realiza o login e retorna um objeto de sessão autenticado."""
    login_url = "https://seguro.linkmonitoramento.com.br/app/includes/confirm_jquery.php"
    sessao = requests.Session()
    sessao.headers.update(USER_AGENTS)

    try:
        sessao.get("https://seguro.linkmonitoramento.com.br/app/", timeout=30)

        payload = {'email': email, 'senha': senha, 'memorizar': '1'}
        print(f"Tentando fazer login com o usuário: {email}...")
        response = sessao.post(login_url, data=payload, timeout=30)
        response.raise_for_status()

        dados_resposta = response.json()
        print("Resposta do login (JSON):", dados_resposta)

        if "Sucesso" in dados_resposta.get("mensagem"):
            print("Login bem-sucedido!")
            return sessao, dados_resposta
        else:
            print(f"Falha no login: {dados_resposta.get('mensagem', 'Erro desconhecido')}")
            return None, dados_resposta

    except requests.exceptions.RequestException as e:
        print(f"Ocorreu um erro na requisição de login: {e}")
        return None, None
    except requests.exceptions.JSONDecodeError:
        print("Erro: A resposta do login não é um JSON válido. Verifique a URL e o payload.")
        print("Conteúdo da resposta:", response.text)
        return None, None


def calcular_distancia_haversine(lat1, lon1, lat2, lon2):
    """Calcula a distância em km entre dois pontos geográficos."""
    raio_terra_km = 6371.0
    lat1_rad, lon1_rad, lat2_rad, lon2_rad = map(math.radians, [lat1, lon1, lat2, lon2])
    dlon = lon2_rad - lon1_rad
    dlat = lat2_rad - lat1_rad
    a = math.sin(dlat / 2) ** 2 + math.cos(lat1_rad) * math.cos(lat2_rad) * math.sin(dlon / 2) ** 2
    c = 2 * math.atan2(math.sqrt(a), math.sqrt(1 - a))
    return raio_terra_km * c


def obter_localizacao_conhecida(lat_veiculo, lon_veiculo):
    """Verifica se um veículo está dentro do raio de um local específico."""
    try:
        lat_veiculo_f, lon_veiculo_f = float(lat_veiculo), float(lon_veiculo)
    except (ValueError, TypeError):
        return "Coordenadas Inválidas"

    for local in BASE_DE_LOCAIS:
        distancia = calcular_distancia_haversine(local['lat'], local['lon'], lat_veiculo_f, lon_veiculo_f)
        if distancia <= local['raio_km']:
            return local['nome']
    return ""


def encontrar_cidade_mais_proxima(lat_veiculo, lon_veiculo):
    """Encontra a cidade mais próxima da base de dados."""
    try:
        lat_veiculo_f, lon_veiculo_f = float(lat_veiculo), float(lon_veiculo)
    except (ValueError, TypeError):
        return None, float('inf')

    cidade_proxima, menor_distancia = None, float('inf')
    for cidade in BASE_DE_CIDADES:
        dist = calcular_distancia_haversine(cidade['coordenadas'][0], cidade['coordenadas'][1], lat_veiculo_f,
                                            lon_veiculo_f)
        if dist < menor_distancia:
            menor_distancia = dist
            cidade_proxima = cidade['nome']
    return cidade_proxima, menor_distancia


def enviar_mensagem_telegram(mensagem):
    """Envia uma mensagem para o chat do Telegram."""
    url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendMessage"
    payload = {'chat_id': TELEGRAM_CHAT_ID, 'text': mensagem, 'parse_mode': 'HTML', 'disable_web_page_preview': True,
               'disable_notification': True}
    try:
        response = requests.post(url, json=payload, timeout=30)
        response.raise_for_status()
        print("Mensagem enviada com sucesso para o Telegram!")
    except requests.exceptions.RequestException as e:
        print(f"Erro ao enviar mensagem para o Telegram: {e}")


def converter_formato_data(data_string):
    """Converte data de UTC para o fuso 'America/Manaus'."""
    try:
        formato_entrada, formato_saida = '%Y-%m-%d %H:%M:%S', '%d/%m/%Y %H:%M'
        fuso_utc, fuso_manaus = ZoneInfo("UTC"), ZoneInfo("America/Manaus")
        objeto_data_naive = datetime.strptime(data_string, formato_entrada)
        objeto_data_convertido = objeto_data_naive.replace(tzinfo=fuso_utc).astimezone(fuso_manaus)
        return objeto_data_convertido.strftime(formato_saida)
    except (ValueError, TypeError):
        return "Data Inválida"


def converter_formato_data_resumido(data_string):
    """Converte data de UTC para o fuso 'America/Manaus'."""
    try:
        formato_entrada, formato_saida = '%Y-%m-%d %H:%M:%S', '%d/%m %H:%M'
        fuso_utc, fuso_manaus = ZoneInfo("UTC"), ZoneInfo("America/Manaus")
        objeto_data_naive = datetime.strptime(data_string, formato_entrada)
        objeto_data_convertido = objeto_data_naive.replace(tzinfo=fuso_utc).astimezone(fuso_manaus)
        return objeto_data_convertido.strftime(formato_saida)
    except (ValueError, TypeError):
        return "Data Inválida"


# --- NOVAS FUNÇÕES E FUNÇÕES MODIFICADAS ---

def processar_dados_veiculos(dados_brutos):
    """
    Processa a lista de veículos da API, enriquecendo com informações de localização.
    Retorna uma lista de dicionários com dados prontos para uso.
    """
    veiculos_processados = []
    if not dados_brutos:
        return veiculos_processados

    for veiculo_bruto in dados_brutos:
        lat = veiculo_bruto.get('latitude')
        lon = veiculo_bruto.get('longitude')

        # Encontra a cidade e o local
        local_conhecido = obter_localizacao_conhecida(lat, lon)
        cidade_proxima, dist_km = encontrar_cidade_mais_proxima(lat, lon)
        cidade_formatada = f"{cidade_proxima} ({dist_km:.1f} km)" if cidade_proxima else ""

        # Monta um dicionário limpo com todas as informações necessárias
        dados_veiculo = {
            'placa': veiculo_bruto.get('rotulo', '').strip().replace("-", ""),
            'horario': converter_formato_data(veiculo_bruto.get('data_hora', '')),
            'horario_resumido': converter_formato_data_resumido(veiculo_bruto.get('data_hora', '')),
            'cidade': cidade_formatada,
            'local': local_conhecido,
            'latitude': lat,
            'longitude': lon,
            'ignicao': "🟢 Ligado" if veiculo_bruto.get('ignicao') == '1' else "🔘 Desligado",
            'velocidade': veiculo_bruto.get('velocidade', ''),
            'motorista': veiculo_bruto.get('nome_motorista', ''),
            'link_maps': f"https://www.google.com/maps?q={lat},{lon}"
        }
        veiculos_processados.append(dados_veiculo)

    return veiculos_processados


def montar_mensagem(dados_processados):
    """Formata os dados processados em uma mensagem de texto para o Telegram."""
    if not dados_processados:
        return "Nenhum dado de veículo foi recebido na resposta."

    agora = datetime.now().strftime('%d/%m/%Y %H:%M:%S')
    mensagem = f"<b>🛰️ Relatório de Veículos - {agora}</b>\n\n"

    for veiculo in dados_processados:
        texto_localizacao = ""
        if veiculo['local'] != "":
            texto_localizacao += f"Local: <b>{veiculo['local']}</b>\n"
        texto_localizacao += f"Cidade: <b>{veiculo['cidade']}</b>\n"

        mensagem += (
            f"<b><code>{veiculo['placa']}</code></b>  👨🏻‍🌾  <code>{veiculo.get('motorista', 'N/A')}</code>\n"
            f"{veiculo['ignicao']} | {veiculo['velocidade']} km/h | {veiculo['horario_resumido']}\n"
            f"📍 {texto_localizacao}"
            f"➡️ <a href='{veiculo['link_maps']}'>Ver no Maps</a>\n"
            "\n"
        )
    return mensagem


def gerar_planilha_excel(dados_processados: list[dict], nome_arquivo="localizacao_dos_veiculos.xlsx"):
    """
    Gera uma planilha Excel (.xlsx) com os dados dos veículos.
    O arquivo é sobrescrito a cada nova execução.

    Args:
        dados_processados (list[dict]): Uma lista de dicionários, onde cada dicionário representa um veículo.
        nome_arquivo (str): O nome do arquivo Excel a ser gerado. Deve terminar com .xlsx.
    """
    if not nome_arquivo.endswith('.xlsx'):
        print(f"Aviso: O nome do arquivo '{nome_arquivo}' não termina com .xlsx. Corrigindo para você.")
        nome_arquivo += '.xlsx'

    # Cria um caminho seguro e portátil para salvar o arquivo em um subdiretório 'relatorios'
    try:
        diretorio_saida = f"C:\\Users\emers\OneDrive\Documentos\Grupo Lodi\Frota Pesada\Atualizacoes\\"
        caminho_completo = diretorio_saida + nome_arquivo
    except Exception as e:
        print(f"Erro ao criar o diretório de saída: {e}")
        return

    print(f"Gerando planilha Excel em: {caminho_completo}")

    if not dados_processados:
        print("Nenhum dado para gerar a planilha.")
        return

    try:
        # 1. Converte a lista de dicionários para um DataFrame do Pandas
        # O Pandas automaticamente usa as chaves dos dicionários como cabeçalho das colunas.
        df = pd.DataFrame(dados_processados)

        # 2. Salva o DataFrame em um arquivo Excel (.xlsx)
        #    - index=False: Impede que o Pandas salve o índice (0, 1, 2...) como uma coluna na planilha.
        #    - sheet_name: Define um nome para a aba da planilha.
        df.to_excel(
            caminho_completo,
            index=False,
            sheet_name='Relatorio de Veiculos'
        )

        print(f"Planilha '{nome_arquivo}' gerada com sucesso com {len(dados_processados)} registros.")

    except ImportError:
        print("Erro: A biblioteca 'openpyxl' é necessária para gerar arquivos .xlsx.")
        print("Por favor, instale-a com o comando: pip install openpyxl")
    except Exception as e:
        print(f"Ocorreu um erro inesperado ao gerar a planilha: {e}")


def gerar_planilha_csv(dados_processados, nome_arquivo="relatorio_veiculos.csv"):
    """
    Gera um arquivo CSV com os dados dos veículos.
    O arquivo é sobrescrito a cada execução.
    """
    arquivo = f"C:\\Users\emers\OneDrive\Documentos\Grupo Lodi\Frota Pesada\Atualizacoes\\{nome_arquivo}"
    print(dados_processados)
    print(f"Gerando planilha CSV: {nome_arquivo}...")
    if not dados_processados:
        print("Nenhum dado para gerar a planilha.")
        return

    # Cabeçalho da planilha, conforme solicitado
    cabecalho = ['placa', 'horario', 'cidade', 'local', 'latitude', 'longitude']

    try:
        # 'w' para modo de escrita (sobrescreve o arquivo)
        # newline='' evita linhas em branco extras no Windows
        # encoding='utf-8' para suportar caracteres especiais
        with open(arquivo, mode='w', newline='', encoding='utf-8') as arquivo_csv:
            # Usa DictWriter para facilitar a escrita a partir de dicionários
            writer = csv.DictWriter(arquivo_csv, fieldnames=cabecalho)

            # Escreve a primeira linha (cabeçalho)
            writer.writeheader()

            # Escreve uma linha para cada veículo
            for veiculo in dados_processados:
                # O DictWriter usa as chaves do dicionário que correspondem ao 'fieldnames'
                campos = {"placa": veiculo['placa'], "horario": veiculo['horario'], "cidade": veiculo['cidade'],
                          "local": veiculo['local'], "latitude": veiculo['latitude'], "longitude": veiculo['longitude']}
                writer.writerow(campos)

        print(f"Planilha '{nome_arquivo}' gerada com sucesso com {len(dados_processados)} registros.")
    except IOError as e:
        print(f"Erro de I/O ao gerar a planilha CSV: {e}")
    except Exception as e:
        print(f"Ocorreu um erro inesperado ao gerar a planilha: {e}")


def executar_tarefa(sessao_ativa):
    """Função principal que busca dados, envia para o Telegram e gera a planilha."""
    print(f"[{datetime.now().strftime('%d/%m/%Y %H:%M:%S')}] Executando a verificação...")

    try:
        sessao_ativa.headers.update(HEADERS)
        response = sessao_ativa.post(URL_API, data=PAYLOAD)
        response.raise_for_status()
        dados_brutos = response.json()

        # 1. Processar os dados brutos para enriquecê-los
        dados_processados = processar_dados_veiculos(dados_brutos)

        # 2. Gerar a mensagem para o Telegram usando os dados processados
        mensagem_formatada = montar_mensagem(dados_processados)
        enviar_mensagem_telegram(mensagem_formatada)

        # 3. Gerar a planilha CSV usando os mesmos dados processados
        gerar_planilha_excel(dados_processados)

    except requests.exceptions.HTTPError as e:
        print(f"Erro HTTP na requisição: {e.response.status_code} - {e.response.text}")
        enviar_mensagem_telegram(
            f"<b>⚠️ Erro na Automação</b>\n\nOcorreu um erro HTTP: {e.response.status_code}.\nVerifique se a sessão de login ainda é válida.")
    except requests.exceptions.RequestException as e:
        print(f"Erro de conexão ou timeout: {e}")
        enviar_mensagem_telegram(f"<b>⚠️ Erro na Automação</b>\n\nNão foi possível conectar à API: {e}")
    except ValueError:
        print("Erro: A resposta não é um JSON válido. Pode ser um erro de login/sessão expirada.")
        if 'response' in locals():
            print(f"Resposta recebida: {response.text}")
        enviar_mensagem_telegram(
            "<b>⚠️ Erro na Automação</b>\n\nA resposta da API não foi um JSON válido. A sessão de login pode ter expirado.")


# --- Bloco de Execução ---
if __name__ == "__main__":
    print(">>> Iniciando automação de monitoramento <<<")

    # Substitua com suas credenciais reais
    meu_email = "transportes@grupolodi.com.br"
    minha_senha = "Fr0ta#010725"

    sessao_ativa, resultado_login = iniciar_sessao_link(meu_email, minha_senha)

    if not sessao_ativa:
        print("\nFalha no processo de login. O script será encerrado.")
        enviar_mensagem_telegram(
            "<b>❌ Falha no Login</b>\n\nNão foi possível iniciar a sessão no sistema de monitoramento. Verifique as credenciais ou a estabilidade do serviço. A automação foi interrompida.")
    else:
        print("\n" + "=" * 40)
        print("Login realizado e sessão criada com sucesso!")
        resposta_painel = sessao_ativa.get(
            "https://seguro.linkmonitoramento.com.br/app/index_full.php?app_modulo=central_monitoramento&app_comando=inicial_monitorar&app_codigo=")
        resposta_painel.raise_for_status()
        print("Iniciando o monitoramento contínuo...")
        print("=" * 40 + "\n")

        # Loop principal de execução
        while True:
            executar_tarefa(sessao_ativa)

            # Intervalo de tempo entre as atualizações (em segundos)
            # 1860 segundos = 31 minutos
            intervalo_segundos = 1500
            print(f"Aguardando {intervalo_segundos / 60:.1f} minutos para a próxima verificação...")
            time.sleep(intervalo_segundos)