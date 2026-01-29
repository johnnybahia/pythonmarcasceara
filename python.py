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
    """
    Identificação de unidade reforçada.
    1. Procura unidade colada ou próxima de números (comum na Aniger).
    2. Procura palavras soltas.
    """
    texto_upper = texto.upper()

    # REGEX DE ELITE: Procura um número (com vírgula) seguido da unidade
    # Ex: "1.000,00 PR" ou "500,00PR"
    if re.search(r'\d+,\d+\s*(PR|PRS|PAR|PARES)\b', texto_upper): return "PAR"
    if re.search(r'\d+,\d+\s*(M|MTS|METRO|METROS)\b', texto_upper): return "METRO"

    # Fallback: Procura a palavra solta (para DASS e outros)
    if re.search(r'\b(PR|PRS|PAR|PARES)\b', texto_upper): return "PAR"
    if re.search(r'\b(M|MTS|METRO|METROS)\b', texto_upper): return "METRO"
    
    return "UNID"

# ================= FUNÇÕES ESPECÍFICAS DE LOCALIZAÇÃO =================

def extrair_local_dass(texto):
    """
    Lógica de APRENDIZADO para DASS.
    1. Procura no cabeçalho o padrão 'DASS NE-XX NomeDaCidade'.
    2. Se falhar, procura o campo 'Cidade:' mas IGNORA o endereço da Marfim/Fornecedor.
    """
    texto_upper = texto.upper()
    
    # 1. TENTATIVA PELO CABEÇALHO (A Mais Segura)
    match_cabecalho = re.search(r'DASS\s+(NE-\d{2})\s+([A-ZÀ-ÿ]+)', texto_upper)
    if match_cabecalho:
        codigo = match_cabecalho.group(1) # NE-02
        cidade = match_cabecalho.group(2) # ITAPIPOCA
        return f"{cidade.title()} ({codigo})"
    
    # 2. TENTATIVA PELO CAMPO 'CIDADE:' (Com Filtro Anti-Erro)
    cidades_encontradas = re.findall(r'CIDADE:\s*([^\n]+)', texto_upper)
    
    match_codigo = re.search(r'(NE-\d{2})', texto_upper)
    codigo_str = f" ({match_codigo.group(1)})" if match_codigo else ""

    for c in cidades_encontradas:
        # Limpeza pesada
        c_limpa = c.replace("- BRAZIL", "").replace("BRAZIL", "")
        c_limpa = c_limpa.split("-")[0].strip()
        c_limpa = c_limpa.split("CEP")[0].strip()
        
        # O FILTRO DE SEGURANÇA: Pula endereço do Fornecedor
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

    return {
        "dataPedido": data_ped,
        "dataRecebimento": data_rec,
        "arquivo": nome_arquivo,
        "cliente": "DILLY SPORTS",
        "marca": marca,
        "local": extrair_local_dilly(texto_completo),
        "qtd": qtd,
        "unidade": identificar_unidade(texto_completo),
        "valor_raw": valor
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

    return {
        "dataPedido": data_ped_str,
        "dataRecebimento": data_rec_str,
        "arquivo": nome_arquivo,
        "cliente": "ANIGER",
        "marca": marca,
        "local": extrair_local_aniger(texto_completo),
        "qtd": qtd,
        "unidade": identificar_unidade(texto_completo), # ✅ Nova Lógica de Unidade
        "valor_raw": valor
    }

def processar_dass(texto_completo, nome_arquivo):
    match_emissao = re.search(r'Data da emissão:\s*(\d{2}/\d{2}/\d{4})', texto_completo, re.IGNORECASE)
    if match_emissao:
        data_rec = match_emissao.group(1)
    else:
        match_header = re.search(r'Hora.*?Data\s*(\d{2}/\d{2}/\d{4})', texto_completo, re.DOTALL)
        data_rec = match_header.group(1) if match_header else datetime.now().strftime("%d/%m/%Y")

    idx_inicio = texto_completo.find("Prev. Ent.")
    texto_busca = texto_completo[idx_inicio:] if idx_inicio != -1 else texto_completo
    match_entrega = re.search(r'\d{8}.*?(\d{2}/\d{2}/\d{4})', texto_busca, re.DOTALL)
    data_ped = match_entrega.group(1) if match_entrega else data_rec

    match_marca = re.search(r'Marca:\s*([^\n]+)', texto_completo)
    marca = match_marca.group(1).strip() if match_marca else "N/D"

    valor = 0.0
    qtd = 0
    match_val = re.search(r'Total valor:\s*([\d\.,]+)', texto_completo)
    if match_val: valor = limpar_valor_monetario(match_val.group(1))
    match_qtd = re.search(r'Total peças:\s*([\d\.,]+)', texto_completo)
    if match_qtd: qtd = int(limpar_valor_monetario(match_qtd.group(1)))

    return {
        "dataPedido": data_ped,
        "dataRecebimento": data_rec,
        "arquivo": nome_arquivo,
        "cliente": "Grupo DASS",
        "marca": marca,
        "local": extrair_local_dass(texto_completo),
        "qtd": qtd,
        "unidade": identificar_unidade(texto_completo),
        "valor_raw": valor
    }

# ================= CONTROLADOR PRINCIPAL =================

def processar_pdf_inteligente(caminho_arquivo, nome_arquivo):
    try:
        with pdfplumber.open(caminho_arquivo) as pdf:
            texto_completo = ""
            for page in pdf.pages:
                texto_completo += page.extract_text() or ""
            
            if "DILLY" in texto_completo.upper():
                dados = processar_dilly(texto_completo, nome_arquivo)
            elif "ANIGER" in texto_completo.upper():
                dados = processar_aniger(texto_completo, nome_arquivo)
            elif "DASS" in texto_completo.upper() or "01287588" in texto_completo:
                dados = processar_dass(texto_completo, nome_arquivo)
            else:
                return None

            dados["valor"] = f"R$ {dados['valor_raw']:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            del dados["valor_raw"] 

            return [dados]

    except Exception as e:
        print(f"Erro ao ler {nome_arquivo}: {e}")
        return []

def mover_arquivos(lista_arquivos):
    if not os.path.exists(PASTA_LIDOS): os.makedirs(PASTA_LIDOS)
    for arquivo in set(lista_arquivos):
        try:
            shutil.move(os.path.join(PASTA_ENTRADA, arquivo), os.path.join(PASTA_LIDOS, arquivo))
        except: pass

def main():
    if not os.path.exists(PASTA_ENTRADA):
        os.makedirs(PASTA_ENTRADA)
        print("Pasta criada.")
        return

    arquivos = [f for f in os.listdir(PASTA_ENTRADA) if f.lower().endswith('.pdf')]
    print(f"📂 Processando {len(arquivos)} arquivos...")
    
    todos_pedidos = []
    arquivos_ok = []

    print("-" * 80)
    print(f"{'CLIENTE':<12} | {'ENTREGA':<12} | {'LOCAL':<15} | {'VALOR':<15}")
    print("-" * 80)

    for arq in arquivos:
        pedidos = processar_pdf_inteligente(os.path.join(PASTA_ENTRADA, arq), arq)
        if pedidos:
            for p in pedidos:
                todos_pedidos.append(p)
                print(f"✅ {p['cliente'][:12]:<12} | {p['dataPedido']:<12} | {p['local'][:15]:<15} | {p['valor']:<15}")
            arquivos_ok.append(arq)
        else:
            print(f"⚠️  Desconhecido: {arq}")

    if todos_pedidos:
        print("\n📤 Enviando...")
        try:
            r = requests.post(URL_WEBAPP, json={"pedidos": todos_pedidos})
            if r.status_code == 200:
                print("☁️ SUCESSO! Dados enviados.")
                mover_arquivos(arquivos_ok)
            else:
                print(f"❌ Erro HTTP {r.status_code}")
        except Exception as e:
            print(f"❌ Erro Conexão: {e}")
    
    input("\nENTER para sair...")

if __name__ == "__main__":
    main()
