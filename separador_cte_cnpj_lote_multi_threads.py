import os
import shutil
import zipfile
import time
import win32api
import win32con
import win32gui
import sys
import re
import concurrent.futures
from tqdm import tqdm
from collections import defaultdict
from threading import Lock
import ctypes
import xml.etree.ElementTree as ET

# Obtém o diretório onde o script está localizado
SCRIPT_DIR = os.path.dirname(os.path.abspath(sys.argv[0]))
PASTA_ORIGEM = os.path.join(SCRIPT_DIR, "1.A Separar")
PASTA_DESTINO = os.path.join(SCRIPT_DIR, "0.Por CNPJ")
PASTA_DUPLICADOS = os.path.join(PASTA_ORIGEM, "1.Duplicados")

# Configurações de performance
REGEX_CNPJ = re.compile(
    rb'<emit[^>]*>.*?<CNPJ>(\d{14})</CNPJ>',
    re.DOTALL | re.IGNORECASE
)
CHUNK_SIZE = 4096 * 16  # 64KB para leitura de arquivos
MAX_WORKERS = os.cpu_count() * 4  # Número de threads paralelas

# Variáveis globais para sincronização
contadores_cnpj = defaultdict(int)
duplicado = 0
contadores_lock = Lock()
erros_lock = Lock()

def criar_pastas_necessarias():
    """Cria as pastas necessárias se não existirem"""
    os.makedirs(PASTA_ORIGEM, exist_ok=True)
    os.makedirs(PASTA_DESTINO, exist_ok=True)
    os.makedirs(os.path.join(PASTA_ORIGEM, "0.Erros"), exist_ok=True)
    os.makedirs(PASTA_DUPLICADOS, exist_ok=True)

class FLASHWINFO(ctypes.Structure):
    _fields_ = [
        ('cbSize', ctypes.c_uint),
        ('hwnd', ctypes.c_void_p),
        ('dwFlags', ctypes.c_uint),
        ('uCount', ctypes.c_uint),
        ('dwTimeout', ctypes.c_uint)
    ]

def flash_window(hwnd):
    """Faz a janela piscar na barra de tarefas até receber foco"""
    try:
        flash_info = FLASHWINFO()
        flash_info.cbSize = ctypes.sizeof(FLASHWINFO)
        flash_info.hwnd = hwnd
        flash_info.dwFlags = win32con.FLASHW_ALL | win32con.FLASHW_TIMERNOFG
        flash_info.uCount = 0  # 0 = até receber foco
        flash_info.dwTimeout = 0

        ctypes.windll.user32.FlashWindowEx(ctypes.byref(flash_info))
    except Exception as e:
        print(f"Erro ao piscar janela: {e}")

def mostrar_popup(mensagem, titulo="Alerta"):
    """Mostra popup informativo com efeito de piscar"""
    try:
        hwnd = win32api.MessageBox(0, mensagem, titulo, 
                                 win32con.MB_OK | win32con.MB_ICONINFORMATION | 
                                 win32con.MB_SETFOREGROUND)
        flash_window(hwnd)
    except Exception as e:
        print(f"\n=== {titulo} ===\n{mensagem}\n")

def mostrar_popup_confirmacao(mensagem, titulo="Confirmação"):
    """Mostra popup com botões Sim/Não que pisca na barra de tarefas"""
    try:
        resposta = win32api.MessageBox(0, mensagem, titulo, 
                                     win32con.MB_YESNO | win32con.MB_ICONQUESTION | 
                                     win32con.MB_SETFOREGROUND)
        flash_window(win32gui.GetForegroundWindow())
        return resposta == win32con.IDYES
    except Exception as e:
        print(f"\n=== {titulo} ===\n{mensagem}\n")
        return False

def configurar_encoding():
    if sys.stdout.encoding != 'utf-8':
        try:
            sys.stdout.reconfigure(encoding='utf-8')
        except:
            pass

def contar_lotes_para_compactar(pasta_destino):
    """Conta quantos lotes existem para compactar"""
    lotes = 0
    if os.path.exists(pasta_destino):
        for cnpj_folder in os.listdir(pasta_destino):
            cnpj_path = os.path.join(pasta_destino, cnpj_folder)
            if os.path.isdir(cnpj_path):
                for lote_folder in os.listdir(cnpj_path):
                    lote_path = os.path.join(cnpj_path, lote_folder)
                    if os.path.isdir(lote_path) and lote_folder.startswith('lote_'):
                        lotes += 1
    return lotes

def compactar_lotes(pasta_destino, manter_pastas=False):
    """Compacta todos os lotes em arquivos ZIP com processamento paralelo"""
    lotes_compactados = 0
    lotes_para_compactar = []
    
    if os.path.exists(pasta_destino):
        for cnpj_folder in os.listdir(pasta_destino):
            cnpj_path = os.path.join(pasta_destino, cnpj_folder)
            if os.path.isdir(cnpj_path):
                for lote_folder in os.listdir(cnpj_path):
                    lote_path = os.path.join(cnpj_path, lote_folder)
                    if os.path.isdir(lote_path) and lote_folder.startswith('lote_'):
                        lotes_para_compactar.append(lote_path)
    
    if not lotes_para_compactar:
        print("\nNenhum lote encontrado para compactar!")
        return 0
    
    # Configuração de paralelismo para compactação
    workers_compactacao = min(4, os.cpu_count() or 1)  # Máximo de 4 workers
    
    def compactar_lote(lote_path):
        """Função que compacta um único lote"""
        nome_zip = f"{os.path.basename(lote_path)}.zip"
        caminho_zip = os.path.join(os.path.dirname(lote_path), nome_zip)
        
        try:
            with zipfile.ZipFile(caminho_zip, 'w', zipfile.ZIP_DEFLATED) as zipf:
                for root, _, files in os.walk(lote_path):
                    for file in files:
                        file_path = os.path.join(root, file)
                        arcname = os.path.relpath(file_path, lote_path)
                        zipf.write(file_path, arcname)
            
            # Remove a pasta apenas se não for para manter as pastas
            if not manter_pastas:
                shutil.rmtree(lote_path)
            return True
        except Exception as e:
            print(f"\nErro ao compactar {lote_path}: {e}")
            return False
    
    # Processamento paralelo dos lotes
    with tqdm(total=len(lotes_para_compactar), unit='lote', desc="Compactando") as pbar:
        with concurrent.futures.ThreadPoolExecutor(max_workers=workers_compactacao) as executor:
            futures = {}
            
            # Enfileirar todos os lotes
            for lote in lotes_para_compactar:
                future = executor.submit(compactar_lote, lote)
                futures[future] = lote
            
            # Processar resultados conforme completam
            for future in concurrent.futures.as_completed(futures):
                if future.result():
                    lotes_compactados += 1
                pbar.update(1)
    
    return lotes_compactados

def extrair_cnpj_rapido(caminho_arquivo):
    """Extrai CNPJ usando regex com leitura otimizada em chunks"""
    try:
        with open(caminho_arquivo, 'rb') as f:
            buffer = b''
            while True:
                chunk = f.read(CHUNK_SIZE)
                if not chunk:
                    break
                    
                buffer += chunk
                match = REGEX_CNPJ.search(buffer)
                if match:
                    return match.group(1).decode('utf-8')
                
                # Manter apenas a última parte do buffer para a próxima iteração
                buffer = buffer[-5000:]
        return None
    except Exception as e:
        print(f"Erro ao ler arquivo {caminho_arquivo}: {e}")
        return None

def extrair_cnpj_confiavel(caminho_arquivo):
    """Extrai CNPJ usando parser XML com namespace correto"""
    try:
        tree = ET.parse(caminho_arquivo)
        root = tree.getroot()
        
        # Namespace oficial do CT-e
        ns = {'ns': 'http://www.portalfiscal.inf.br/cte'}
        
        # Busca segura com namespace
        emitente = root.find('.//ns:emit', ns)
        if emitente is not None:
            cnpj_element = emitente.find('ns:CNPJ', ns)
            if cnpj_element is not None and cnpj_element.text:
                return cnpj_element.text.strip()
        
        # Fallback para busca sem namespace
        emitente_fallback = root.find('.//emit')
        if emitente_fallback is not None:
            cnpj_fallback = emitente_fallback.find('CNPJ')
            if cnpj_fallback is not None and cnpj_fallback.text:
                return cnpj_fallback.text.strip()
        
        return None
    except Exception as e:
        print(f"Erro no parser XML: {e}")
        return None

def criar_arquivo_log_erros(pasta_erros, qtd_erros):
    """Cria um arquivo de log com a quantidade de erros encontrados"""
    caminho_log = os.path.join(pasta_erros, f"0.Erros_{qtd_erros}.txt")
    with open(caminho_log, 'w') as f:
        f.write(f"Total de arquivos com erro: {qtd_erros}\n")
        f.write(f"Data do processamento: {time.strftime('%d/%m/%Y %H:%M:%S')}\n")

def extrair_cnpj_otimizado(caminho_arquivo):
    """Combina regex rápida com parser XML para confiabilidade"""
    # Primeira tentativa: regex rápida
    cnpj_regex = extrair_cnpj_rapido(caminho_arquivo)
    if cnpj_regex:
        return cnpj_regex
    
    # Fallback: parser XML com namespace
    return extrair_cnpj_confiavel(caminho_arquivo)

def renomear_arquivo_existente(caminho_destino):
    """Renomeia arquivo existente com contador sequencial e move para pasta de duplicados"""
    global duplicado
    try:
        # Extrai nome e extensão do arquivo
        nome_arquivo = os.path.basename(caminho_destino)
        nome_base, extensao = os.path.splitext(nome_arquivo)
        
        # Contador para nomes duplicados
        duplicados = 1
        novo_caminho = os.path.join(PASTA_DUPLICADOS, f"{nome_base} ({duplicados}){extensao}")
        
        # Encontra o próximo número disponível
        while os.path.exists(novo_caminho):
            duplicados += 1
            novo_caminho = os.path.join(PASTA_DUPLICADOS, f"{nome_base} ({duplicados}){extensao}")
        
        # Move o arquivo existente para a pasta de duplicados
        shutil.move(caminho_destino, novo_caminho)
        duplicado += 1
        return True
    except Exception as e:
        print(f"Erro ao renomear arquivo existente: {e}")
        return False

def processar_arquivo(caminho_completo, pasta_erros):
    """Processa um único arquivo XML (função usada por threads)"""
    global contadores_cnpj, duplicado
    
    try:
        # Extração otimizada de CNPJ
        cnpj = extrair_cnpj_otimizado(caminho_completo)
        
        if not cnpj:
            raise ValueError("CNPJ não encontrado no XML")
        
        # Atualização segura dos contadores
        with contadores_lock:
            contadores_cnpj[cnpj] += 1
            numero_lote = (contadores_cnpj[cnpj] - 1) // 50000 + 1
            
            pasta_cnpj = os.path.join(PASTA_DESTINO, cnpj)
            pasta_lote = os.path.join(pasta_cnpj, f"lote_{numero_lote}")
            os.makedirs(pasta_lote, exist_ok=True)
            
            destino = os.path.join(pasta_lote, os.path.basename(caminho_completo))
            
            # Tratamento de arquivos duplicados
            if os.path.exists(destino):
                renomear_arquivo_existente(destino)
            
            shutil.move(caminho_completo, destino)
        
        return True
    
    except Exception as e:
        # Tratamento seguro de erros
        with erros_lock:
            destino_erro = os.path.join(pasta_erros, os.path.basename(caminho_completo))
            shutil.move(caminho_completo, destino_erro)
        return False

def organizar_cte_por_emitente():
    """
    Organiza arquivos XML de CT-e em pastas por CNPJ com processamento paralelo
    """
    global contadores_cnpj, duplicado
    
    pasta_erros = os.path.join(PASTA_ORIGEM, "0.Erros")

    mensagem_pos_separacao = ""
    
    # Verificar se há arquivos para processar
    print(f"\nVerificando arquivos na pasta: {PASTA_ORIGEM}")
    total_arquivos = []
    for raiz, _, arquivos in os.walk(PASTA_ORIGEM):
        # Ignorar pasta de erros
        if os.path.basename(raiz) == "0.Erros" and os.path.dirname(raiz) == PASTA_ORIGEM:
            continue
            
        for arquivo in arquivos:
            if arquivo.lower().endswith('.xml'):
                total_arquivos.append(os.path.join(raiz, arquivo))
    
    tem_arquivos = len(total_arquivos) > 0
    total_lotes = contar_lotes_para_compactar(PASTA_DESTINO)

    if not tem_arquivos and total_lotes == 0:
        mostrar_popup("Nenhum arquivo XML encontrado para separar e nenhum lote para compactar!", "Aviso")
        return
    
    # Contadores
    processados = 0
    erros = 0
    lotes_compactados = 0

    # Resetar contadores globais
    contadores_cnpj = defaultdict(int)
    duplicado = 0

    # Processar XMLs se existirem
    if tem_arquivos:
        print(f"\nProcessando {len(total_arquivos)} arquivos XML com {MAX_WORKERS} threads paralelas...")

        # Inicializar contadores para CNPJs conhecidos
        for cnpj_folder in os.listdir(PASTA_DESTINO):
            if os.path.isdir(os.path.join(PASTA_DESTINO, cnpj_folder)):
                # Contar arquivos existentes para cada CNPJ
                contador = 0
                for root, _, files in os.walk(os.path.join(PASTA_DESTINO, cnpj_folder)):
                    contador += len([f for f in files if f.lower().endswith('.xml')])
                contadores_cnpj[cnpj_folder] = contador
        
        # Processamento paralelo com ThreadPoolExecutor
        with tqdm(total=len(total_arquivos), unit='arquivo', desc="Separando CT-es") as pbar:
            with concurrent.futures.ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
                # Iniciar todas as tarefas
                futures = {executor.submit(processar_arquivo, arq, pasta_erros): arq for arq in total_arquivos}
                
                # Processar resultados conforme ficam prontos
                for future in concurrent.futures.as_completed(futures):
                    try:
                        if future.result():
                            processados += 1
                        else:
                            erros += 1
                    except Exception as e:
                        erros += 1
                        print(f"Erro no processamento paralelo: {e}")
                    
                    pbar.update(1)
                    pbar.set_postfix({'OK': processados, 'Erros': erros, 'duplicados': duplicado})
        
        # Remover pastas vazias
        for raiz, dirs, _ in os.walk(PASTA_ORIGEM, topdown=False):
            for dir in dirs:
                dir_path = os.path.join(raiz, dir)
                if dir != os.path.basename(PASTA_DESTINO) and dir != "0.Erros" and dir != "1.Duplicados":
                    try:
                        if not os.listdir(dir_path):
                            os.rmdir(dir_path)
                    except:
                        pass

        mensagem_pos_separacao = f"Todos os {len(total_arquivos)} arquivos XML foram separados.\n\n"
    else:
        mensagem_pos_separacao = "Nenhum arquivo XML encontrado para separar.\n\n"
    
    # Verifica novamente lotes após processamento
    total_lotes = contar_lotes_para_compactar(PASTA_DESTINO)
    tem_lotes = total_lotes > 0
    
    if tem_lotes:
        mensagem_pos_separacao += f"Foram encontrados {total_lotes} lotes.\nDeseja compactá-los agora?"
        
        if mostrar_popup_confirmacao(mensagem_pos_separacao, "Manter pastas?"):
            resposta = win32api.MessageBox(
                0,
                "Deseja manter as pastas após a compactação?\n\n"
                "• Sim: Manter pastas e arquivos ZIP (ambos)\n"
                "• Não: Manter apenas arquivos ZIP (pastas serão excluídas)",
                "Opções de Compactação",
                win32con.MB_YESNOCANCEL | win32con.MB_ICONQUESTION | win32con.MB_SETFOREGROUND
            )
            
            if resposta == win32con.IDYES:  # Manter ambos
                print("\nCompactando e mantendo pastas lotes...")
                lotes_compactados = compactar_lotes(PASTA_DESTINO, manter_pastas=True)
                opcao = "Mantidas pastas e ZIPs"
                
            elif resposta == win32con.IDNO:  # Manter apenas ZIP
                print("\nCompactando e removendo pastas lotes...")
                lotes_compactados = compactar_lotes(PASTA_DESTINO, manter_pastas=False)
                opcao = "Mantidos apenas ZIPs"
                
            else:  # IDCANCEL
                print("\nCompactação cancelada pelo usuário.")
                opcao = "Nenhuma opção de compactação"
                lotes_compactados = 0
        else:
            print("\nCompactação cancelada pelo usuário.")
            opcao = "Nenhuma opção de compactação"
            lotes_compactados = 0
    else:
        mensagem_pos_separacao += "Nenhum lote encontrado para compactar."
        mostrar_popup(mensagem_pos_separacao, "Status")
        opcao = "Nenhuma opção de compactação"
        lotes_compactados = 0
    
    # Cria log de erros se necessário
    if erros > 0:
        criar_arquivo_log_erros(pasta_erros, erros)

    # Popup final com resultados
    tempo_total = time.time() - inicio
    mensagem_final = (f"Processo finalizado!\n\n"
                     f"✓ Arquivos processados: {processados}\n"
                     f"✗ Arquivos com erro: {erros}\n"
                     f"👥 Arquivos duplicados: {duplicado}\n"
                     f"📦 Lotes compactados: {lotes_compactados}\n"
                     f"⚙ Opção: {opcao}\n"
                     f"⏱ Tempo total: {tempo_total:.2f}s\n\n"
                     f"Pasta do script: {SCRIPT_DIR}")
    
    print("\n" + "="*50)
    # Versão ASCII para o console
    mensagem_console = mensagem_final.replace('✓', '[PROCESSADOS]').replace('✗', '[ERRO]') \
                                     .replace('👥', '[DUPLICADOS]').replace('📦', '[COMPACTADOS]') \
                                     .replace('⚙', '[OPÇÃO]').replace('⏱', '[TEMPO]')
    print(mensagem_console)
    print("="*50)
    
    mostrar_popup(mensagem_final, "Processo Concluído")

if __name__ == "__main__":
    configurar_encoding()
    criar_pastas_necessarias()

    print("=== ORGANIZADOR DE CT-es PORTÁTIL ===")
    print(f"Local do script: {SCRIPT_DIR}")
    print(f"Pasta origem (XMLs): {PASTA_ORIGEM}")
    print(f"Pasta destino (CNPJ): {PASTA_DESTINO}")
    print(f"Pasta duplicados: {PASTA_DUPLICADOS}")
    print(f"Número de threads: {MAX_WORKERS}\n")

    inicio = time.time()
    organizar_cte_por_emitente()