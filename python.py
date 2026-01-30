import pdfplumber
import re
import os
import shutil
import requests
import json
from datetime import datetime

# ================= CONFIGURAÇÃO =================
URL_WEBAPP = "https://script.google.com/macros/s/AKfycbzke-sTVigX4hkUkqLaTYRL0WDi_P-JAhc4PPsjwf0GuwSz_92lx43fQVM07XiiNBrjbA/exec"

PASTA_ENTRADA = './pedidos'
PASTA_LIDOS = './pedidos/lidos'
# =================================================

def converter_data_curta(data_str):
    """Converte DD/MM/YY para DD/MM/YYYY."""
    if not data_str:
        return datetime.now().strftime("%d/%m/%Y")
    parts = data_str.strip().split('/')
    if len(parts) == 3 and len(parts[2]) == 2:
        return f"{parts[0]}/{parts[1]}/20{parts[2]}"
    return data_str

def limpar_valor_monetario(texto):
    if not texto: return 0.0
    texto = texto.lower().replace('r$', '').replace('total', '').strip()
    if ',' in texto and '.' in texto:
        texto = texto.replace('.', '').replace(',', '.')
    elif ',' in texto:
        texto = texto.replace(',', '.')
    try: return float(texto)
    except: return 0.0

def identificar_unidade(texto):
    texto_upper = texto.upper()
    if re.search(r'\d+,\d+\s*(PR|PRS|PAR|PARES)\b', texto_upper): return "PAR"
    if re.search(r'\d+,\d+\s*(M|MTS|METRO|METROS)\b', texto_upper): return "METRO"
    if re.search(r'\b(PR|PRS|PAR|PARES)\b', texto_upper): return "PAR"
    if re.search(r'\b(M|MTS|METRO|METROS)\b', texto_upper): return "METRO"
    return "UNID"

# ================= EXTRAÇÃO DA ORDEM DE COMPRA =================

def extrair_ordem_compra(texto):
    """
    Extrai o número da Ordem de Compra do PDF.
    Usa \d{6,} para pular pedaços do CNPJ (94, 316, 999, 0009, 83)
    e capturar somente o número real da OC (6+ dígitos).
    """
    match = re.search(r'Ordem\s+(?:de\s+)?[Cc]ompra[\s\S]{0,50}?(\d{6,})', texto)
    if match:
        return match.group(1)
    return "N/D"

# ================= FUNÇÕES DE LOCAL DE ENTREGA =================

def extrair_local_dass(texto):
    texto_upper = texto.upper()

    match_cabecalho = re.search(r'DASS\s+(NE-\d{2})\s+([A-ZÀ-ÿ]+)', texto_upper)
    if match_cabecalho:
        codigo = match_cabecalho.group(1)
        cidade = match_cabecalho.group(2)
        return f"{cidade.title()} ({codigo})"

    cidades_encontradas = re.findall(r'CIDADE:\s*([^\n]+)', texto_upper)
    match_codigo = re.search(r'(NE-\d{2})', texto_upper)
    codigo_str = f" ({match_codigo.group(1)})" if match_codigo else ""

    for c in cidades_encontradas:
        c_limpa = c.replace("- BRAZIL", "").replace("BRAZIL", "")
        c_limpa = c_limpa.split("-")[0].strip()
        c_limpa = c_limpa.split("CEP")[0].strip()
        if "EUSEBIO" in c_limpa or "CRUZ DAS ALMAS" in c_limpa or "MARFIM" in c_limpa:
            continue
        if len(c_limpa) > 3:
            return f"{c_limpa.title()}{codigo_str}"

    return "N/D"

def extrair_local_dilly(texto):
    texto_upper = texto.upper()
    match_generico = re.search(r',\s*([A-Z\s]+)-[A-Z]{2}', texto_upper)
    if match_generico:
        cidade_encontrada = match_generico.group(1).strip()
        if len(cidade_encontrada) < 30 and "MARFIM" not in cidade_encontrada:
            return cidade_encontrada.title()
    if "BREJO" in texto_upper: return "Brejo Santo"
    if "MORADA" in texto_upper: return "Morada Nova"
    if "QUIXERAMOBIM" in texto_upper: return "Quixeramobim"
    return "N/D"

def extrair_local_aniger(texto):
    texto_upper = texto.upper()
    if re.search(r'QUIXERAMOBIM', texto_upper): return "Quixeramobim"
    if re.search(r'IVOTI', texto_upper): return "Ivoti"
    return "N/D"

# ================= PROCESSAMENTO POR CLIENTE =================

def processar_dilly(texto_completo, nome_arquivo):
    match_emissao = re.search(r'Data Emissão:\s*(\d{2}/\d{2}/\d{4})', texto_completo)
    data_rec = match_emissao.group(1) if match_emissao else datetime.now().strftime("%d/%m/%Y")

    match_entrega_tab = re.search(r'Previsão.*?(\d{2}/\d{2}/\d{4})', texto_completo, re.DOTALL)
    data_ped = match_entrega_tab.group(1) if match_entrega_tab else data_rec

    match_marca = re.search(r'Marca:\s*([^\s]+)', texto_completo)
    marca = match_marca.group(1).strip() if match_marca else "DILLY"

    qtd = 0
    valor = 0.0
    match_qtd = re.search(r'Quantidade Total:\s*([\d\.,]+)', texto_completo)
    if match_qtd: qtd = int(limpar_valor_monetario(match_qtd.group(1)))
    match_valor = re.search(r'Total\s*R\$([\d\.,]+)', texto_completo)
    if match_valor: valor = limpar_valor_monetario(match_valor.group(1))

    # --- 3. ORDEM DE COMPRA ---
    match_ordem = re.search(r'Ordem\s+(?:de\s+)?[Cc]ompra[\s\S]{0,50}?(\d{6,})', texto_completo)
    ordem_compra = match_ordem.group(1) if match_ordem else "N/D"

    valor_formatado = f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

    return {
        "dataPedido": data_ped,
        "dataRecebimento": data_rec,
        "arquivo": nome_arquivo,
        "cliente": "DILLY SPORTS",
        "marca": marca,
        "local": extrair_local_dilly(texto_completo),
        "qtd": qtd,
        "unidade": identificar_unidade(texto_completo),
        "valor": valor_formatado,
        "ordemCompra": ordem_compra
    }

def processar_aniger(texto_completo, nome_arquivo):
    match_emissao = re.search(r'Emissão:\s*(\d{2}/\d{2}/\d{4})', texto_completo)
    if not match_emissao:
        match_emissao = re.search(r'Emissão:.*?(\d{2}/\d{2}/\d{4})', texto_completo, re.DOTALL)
    data_rec_str = match_emissao.group(1) if match_emissao else datetime.now().strftime("%d/%m/%Y")

    todas_datas = re.findall(r'(\d{2}/\d{2}/\d{4})', texto_completo)
    data_ped_str = data_rec_str

    try:
        data_rec_obj = datetime.strptime(data_rec_str, "%d/%m/%Y")
        for d_str in todas_datas:
            try:
                d_obj = datetime.strptime(d_str, "%d/%m/%Y")
                if (d_obj - data_rec_obj).days > 5:
                    data_ped_str = d_str
                    break
            except: continue
    except: pass

    marca = "ANIGER"
    if "NIKE" in texto_completo.upper(): marca = "NIKE (Aniger)"

    qtd = 0
    valor = 0.0
    match_totais = re.search(r'Totais\s+([\d\.,]+).*?([\d\.,]+)', texto_completo, re.DOTALL)
    if match_totais:
        qtd = int(limpar_valor_monetario(match_totais.group(1)))
        valor = limpar_valor_monetario(match_totais.group(2))

    # --- 3. ORDEM DE COMPRA ---
    match_ordem = re.search(r'Ordem\s+(?:de\s+)?[Cc]ompra[\s\S]{0,50}?(\d{6,})', texto_completo)
    ordem_compra = match_ordem.group(1) if match_ordem else "N/D"

    valor_formatado = f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

    return {
        "dataPedido": data_ped_str,
        "dataRecebimento": data_rec_str,
        "arquivo": nome_arquivo,
        "cliente": "ANIGER",
        "marca": marca,
        "local": extrair_local_aniger(texto_completo),
        "qtd": qtd,
        "unidade": identificar_unidade(texto_completo),
        "valor": valor_formatado,
        "ordemCompra": ordem_compra
    }

def processar_dass(texto_completo, nome_arquivo):
    # --- 1. DATA DE RECEBIMENTO (Data da Emissão) ---
    match_emissao = re.search(r'Data da emissão:\s*(\d{2}/\d{2}/\d{4})', texto_completo, re.IGNORECASE)
    if match_emissao:
        data_rec = match_emissao.group(1)
    else:
        match_header = re.search(r'Hora.*?Data\s*(\d{2}/\d{2}/\d{4})', texto_completo, re.DOTALL)
        data_rec = match_header.group(1) if match_header else datetime.now().strftime("%d/%m/%Y")

    # --- 2. DATA DO PEDIDO (Entrega) ---
    idx_inicio = texto_completo.find("Prev. Ent.")
    texto_busca = texto_completo[idx_inicio:] if idx_inicio != -1 else texto_completo
    match_entrega = re.search(r'\d{8}.*?(\d{2}/\d{2}/\d{4})', texto_busca, re.DOTALL)
    data_ped = match_entrega.group(1) if match_entrega else data_rec

    # --- 3. ORDEM DE COMPRA ---
    match_ordem = re.search(r'Ordem\s+(?:de\s+)?[Cc]ompra[\s\S]{0,50}?(\d{6,})', texto_completo)
    ordem_compra = match_ordem.group(1) if match_ordem else "N/D"

    # --- DADOS GERAIS ---
    match_marca = re.search(r'Marca:\s*([^\n]+)', texto_completo)
    marca = match_marca.group(1).strip() if match_marca else "N/D"

    valor = 0.0
    qtd = 0
    match_val = re.search(r'Total valor:\s*([\d\.,]+)', texto_completo)
    if match_val: valor = limpar_valor_monetario(match_val.group(1))
    match_qtd = re.search(r'Total peças:\s*([\d\.,]+)', texto_completo)
    if match_qtd: qtd = int(limpar_valor_monetario(match_qtd.group(1)))

    valor_formatado = f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

    return {
        "dataPedido": data_ped,
        "dataRecebimento": data_rec,
        "arquivo": nome_arquivo,
        "cliente": "Grupo DASS",
        "marca": marca,
        "local": extrair_local_dass(texto_completo),
        "qtd": qtd,
        "unidade": identificar_unidade(texto_completo),
        "valor": valor_formatado,
        "ordemCompra": ordem_compra
    }

def processar_dakota(pages, nome_arquivo):
    """
    Processa PDF da DAKOTA. Cada linha da tabela gera um pedido separado.
    Usa extract_tables() do pdfplumber para ler a tabela estruturada.
    Colunas: #, Prioridade, Filial, OC, Emissão, Entrega, Limite, Comprador, Material, Unid, Qtd, Saldo
    """
    pedidos = []
    compradores_conhecidos = ('saimon', 'ccarlos')

    for page in pages:
        tables = page.extract_tables()
        if not tables:
            continue

        for table in tables:
            for row in table:
                if not row or len(row) < 8:
                    continue

                # Linhas de dados começam com número sequencial
                primeiro = str(row[0] or '').strip()
                if not primeiro.isdigit():
                    continue

                # Varrer cada célula e identificar pelo padrão
                filial = ""
                oc = ""
                datas = []
                unidade = "UNID"
                qtd = 0

                for cell in row:
                    val = str(cell or '').strip()
                    if not val or val == primeiro:
                        continue

                    # OC: dígitos + letra no final (41110T, 79703D, 79632D)
                    if not oc and re.match(r'^\d+[A-Za-z]$', val):
                        oc = val
                        continue

                    # Datas DD/MM/YY
                    if re.match(r'^\d{2}/\d{2}/\d{2}$', val):
                        datas.append(val)
                        continue

                    # Unidade (PR = PAR, MT = METRO)
                    if val.upper() in ('PR', 'MT'):
                        unidade = 'PAR' if val.upper() == 'PR' else 'METRO'
                        continue

                    # Pular compradores conhecidos
                    if val.lower() in compradores_conhecidos:
                        continue

                    # Pular descrição de material (começa com código numérico)
                    if re.match(r'^\d{4,}', val):
                        continue

                    # Filial: nome de cidade (só letras, 4+ caracteres)
                    if not filial and re.match(r'^[A-ZÀ-ÿa-zà-ÿ\s]+$', val) and len(val.strip()) >= 4:
                        filial = val
                        continue

                if not oc:
                    continue

                # Quantidade: primeiro número decimal encontrado (Qtd. OC)
                for cell in row:
                    val = str(cell or '').strip()
                    if val == primeiro or not val:
                        continue
                    if re.match(r'^[\d\.,]+$', val):
                        try:
                            num = int(float(val.replace('.', '').replace(',', '.')))
                            if num > 0:
                                qtd = num
                                break
                        except:
                            continue

                # Primeira data = emissão, segunda = entrega
                emissao = converter_data_curta(datas[0]) if datas else datetime.now().strftime("%d/%m/%Y")
                entrega = converter_data_curta(datas[1]) if len(datas) >= 2 else emissao

                pedidos.append({
                    "dataPedido": entrega,
                    "dataRecebimento": emissao,
                    "arquivo": nome_arquivo,
                    "cliente": "DAKOTA",
                    "marca": "DAKOTA",
                    "local": filial.strip().title(),
                    "qtd": qtd,
                    "unidade": unidade,
                    "valor": "R$ 0,00",
                    "ordemCompra": oc
                })

    return pedidos if pedidos else None

# ================= CONTROLADOR PRINCIPAL =================

def processar_pdf_inteligente(caminho_arquivo, nome_arquivo):
    try:
        with pdfplumber.open(caminho_arquivo) as pdf:
            texto_completo = ""
            for page in pdf.pages:
                texto_completo += page.extract_text() or ""

            texto_upper = texto_completo.upper()

            if "DILLY" in texto_upper:
                return [processar_dilly(texto_completo, nome_arquivo)]
            elif "ANIGER" in texto_upper:
                return [processar_aniger(texto_completo, nome_arquivo)]
            elif "DASS" in texto_upper or "01287588" in texto_completo:
                return [processar_dass(texto_completo, nome_arquivo)]
            elif "DAKOTA" in texto_upper:
                return processar_dakota(pdf.pages, nome_arquivo)
            else:
                return None

    except Exception as e:
        print(f"Erro ao abrir {nome_arquivo}: {e}")
        return []

def mover_arquivos_processados(lista_arquivos):
    if not os.path.exists(PASTA_LIDOS): os.makedirs(PASTA_LIDOS)
    print(f"\n📦 Movendo arquivos processados para: {PASTA_LIDOS}")
    for arquivo in set(lista_arquivos):
        try:
            caminho_origem = os.path.join(PASTA_ENTRADA, arquivo)
            caminho_destino = os.path.join(PASTA_LIDOS, arquivo)
            if os.path.exists(caminho_destino): os.remove(caminho_destino)
            shutil.move(caminho_origem, caminho_destino)
        except: pass

def main():
    if not os.path.exists(PASTA_ENTRADA):
        os.makedirs(PASTA_ENTRADA)
        print(f"Pasta criada. Coloque PDFs em '{PASTA_ENTRADA}'.")
        return

    todos_pedidos_para_envio = []
    arquivos_para_mover = []

    arquivos = [f for f in os.listdir(PASTA_ENTRADA) if f.lower().endswith('.pdf')]

    print(f"📂 Processando {len(arquivos)} arquivos...")
    print("-" * 95)
    print(f"{'EMISSÃO':<12} | {'ENTREGA':<12} | {'OC':<12} | {'CLIENTE':<15} | {'MARCA':<15} | {'VALOR'}")
    print("-" * 95)

    for arq in arquivos:
        lista_pedidos = processar_pdf_inteligente(os.path.join(PASTA_ENTRADA, arq), arq)

        if lista_pedidos:
            for p in lista_pedidos:
                todos_pedidos_para_envio.append(p)
                print(f"✅ {p['dataRecebimento']:<12} | {p['dataPedido']:<12} | {p['ordemCompra']:<12} | {p['cliente'][:15]:<15} | {p['marca'][:15]:<15} | {p['valor']}")
            arquivos_para_mover.append(arq)
        else:
            print(f"⚠️  Ignorado: {arq}")

    if todos_pedidos_para_envio:
        print("-" * 95)
        print(f"📤 Enviando {len(todos_pedidos_para_envio)} pedidos para Google Sheets...")

        try:
            response = requests.post(URL_WEBAPP, json={"pedidos": todos_pedidos_para_envio}, timeout=30)

            print(f"\n📡 Status: {response.status_code}")

            if response.status_code == 200:
                print(f"☁️  SUCESSO! Google recebeu os dados.")
                mover_arquivos_processados(arquivos_para_mover)
            else:
                print(f"❌ Erro HTTP {response.status_code}: {response.text}")

        except Exception as e:
            print(f"\n❌ Erro de conexão: {e}")
    else:
        print("\n⚠️  Nenhum pedido válido encontrado.")

    input("\nPressione ENTER para fechar...")

if __name__ == "__main__":
    main()
