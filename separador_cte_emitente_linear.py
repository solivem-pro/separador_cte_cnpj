import os
import shutil
import zipfile
from xml.etree import ElementTree as ET
import time
from tqdm import tqdm
import win32api
import win32con
import win32gui
import sys
import ctypes

# Obtém o diretório onde o script está localizado
SCRIPT_DIR = os.path.dirname(os.path.abspath(sys.argv[0]))
PASTA_ORIGEM = os.path.join(SCRIPT_DIR, "1.A Separar")
PASTA_DESTINO = os.path.join(SCRIPT_DIR, "0.Por CNPJ")
PASTA_ERROS = os.path.join(PASTA_DESTINO, "0.Erros")
PASTA_DUPLICADOS = os.path.join(PASTA_DESTINO, "1.Duplicados")
RELATORIO_DIR = os.path.join(PASTA_DESTINO, f"0.relatorio.txt")

def criar_pastas_necessarias():
    """Cria as pastas necessárias se não existirem"""
    os.makedirs(PASTA_ORIGEM, exist_ok=True)
    os.makedirs(PASTA_DESTINO, exist_ok=True)
    os.makedirs(os.path.join(PASTA_DESTINO, "0.Erros"), exist_ok=True)
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

def mostrar_popup_opcoes_compactacao(mensagem, titulo="Opções de Compactação"):
    """Mostra popup com três opções de ação para compactação"""
    try:
        resposta = win32api.MessageBox(0, mensagem, titulo,
                                     win32con.MB_ICONQUESTION | win32con.MB_SETFOREGROUND | 
                                     win32con.MB_YESNOCANCEL)
        flash_window(win32gui.GetForegroundWindow())
        return resposta
    except Exception as e:
        print(f"\n=== {titulo} ===\n{mensagem}\n")
        return win32con.IDCANCEL

def configurar_encoding():
    if sys.stdout.encoding != 'utf-8':
        try:
            sys.stdout.reconfigure(encoding='utf-8')
        except:
            pass

def contar_xmls(pasta):
    """Conta todos os arquivos XML em uma pasta e subpastas"""
    count = 0
    for root, _, files in os.walk(pasta):
        count += sum(1 for f in files if f.lower().endswith('.xml'))
    return count

def criar_arquivo_log_erros(PASTA_ERROS, qtd_erros):
    """Cria um arquivo de log com a quantidade de erros encontrados"""
    caminho_log = os.path.join(PASTA_ERROS, f"0.Erros_{qtd_erros}.txt")
    with open(caminho_log, 'w') as f:
        f.write(f"Total de arquivos com erro: {qtd_erros}\n")
        f.write(f"Data do processamento: {time.strftime('%d/%m/%Y %H:%M:%S')}\n")

def contar_lotes_para_compactar(pasta_destino):
    """Conta quantos lotes existem para compactar"""
    lotes = 0
    if os.path.exists(pasta_destino):
        for cnpj_folder in os.listdir(pasta_destino):
            cnpj_path = os.path.join(pasta_destino, cnpj_folder)
            if os.path.isdir(cnpj_path):
                for lote_folder in os.listdir(cnpj_path):
                    lote_path = os.path.join(cnpj_path, lote_folder)
                    if os.path.isdir(lote_path) and lote_folder.startswith('20'):
                        lotes += 1
    return lotes

def compactar_lotes(pasta_destino, manter_pastas=False):
    """
    Compacta todos os lotes em arquivos ZIP com barra de progresso
    :param manter_pastas: Se True, mantém as pastas originais após compactação
    """
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
    
    with tqdm(total=len(lotes_para_compactar), unit='lote', desc="Compactando") as pbar:
        for lote_path in lotes_para_compactar:
            nome_zip = f"{os.path.basename(lote_path)}.zip"
            caminho_zip = os.path.join(os.path.dirname(lote_path), nome_zip)
            
            # Verifica se o ZIP já existe e remove para recriar
            if os.path.exists(caminho_zip):
                os.remove(caminho_zip)
            
            with zipfile.ZipFile(caminho_zip, 'w', zipfile.ZIP_DEFLATED) as zipf:
                for root, _, files in os.walk(lote_path):
                    for file in files:
                        file_path = os.path.join(root, file)
                        arcname = os.path.relpath(file_path, lote_path)
                        zipf.write(file_path, arcname)
            
            # Remove a pasta original apenas se não for para manter
            if not manter_pastas:
                shutil.rmtree(lote_path)
            
            lotes_compactados += 1
            pbar.update(1)
    
    return lotes_compactados

def abrir_arquivos():
    """Abre a planilha de controle e o arquivo de erros usando o aplicativo padrão do sistema"""
    try:
        # Abrir a planilha de controle
        if os.path.exists(RELATORIO_DIR):
            os.startfile(RELATORIO_DIR)
        else:
            print(f"Arquivo da planilha não encontrado: {RELATORIO_DIR}")
            
    except Exception as e:
        print(f"Erro ao abrir arquivos: {str(e)}")

def renomear_arquivo_existente(caminho_destino, PASTA_DUPLICADOS):
    """Renomeia arquivo existente com contador sequencial e move para pasta de duplicados"""
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
        return True
    except Exception as e:
        print(f"Erro ao renomear arquivo existente: {e}")
        return False

def gerar_relatorio_por_cnpj(pasta_destino):
    """
    Gera um relatório com a quantidade de pastas de data e arquivos XML em cada pasta de CNPJ.
    Retorna uma string formatada para incluir no relatório.
    """
    relatorio = []
    if os.path.exists(pasta_destino):
        for cnpj_folder in sorted(os.listdir(pasta_destino)):
            cnpj_path = os.path.join(pasta_destino, cnpj_folder)
            if os.path.isdir(cnpj_path):
                datas = []
                for data_folder in sorted(os.listdir(cnpj_path)):
                    data_path = os.path.join(cnpj_path, data_folder)
                    if os.path.isdir(data_path):
                        qtd_xml = sum(1 for f in os.listdir(data_path) if f.lower().endswith('.xml'))
                        datas.append(f"    {data_folder}: {qtd_xml} arquivo(s)")
                relatorio.append(f"- CNPJ {cnpj_folder}:\n" + "\n".join(datas))
    return "\n\n".join(relatorio)

def organizar_cte_por_emitente():
    """
    Organiza arquivos XML de CT-e em pastas por CNPJ e data de emissão (AAAA-MM-DD).
    Usa pastas relativas ao local onde o script está salvo.
    """

    print(f"\nVerificando arquivos na pasta: {PASTA_ORIGEM}")
    total_arquivos = contar_xmls(PASTA_ORIGEM)
    tem_xmls = total_arquivos > 0
    total_lotes = contar_lotes_para_compactar(PASTA_DESTINO)
    tem_lotes = total_lotes > 0
    
    if not tem_xmls and not tem_lotes:
        mostrar_popup("Nenhum arquivo XML encontrado para separar e nenhum lote para compactar!", "Aviso")
        return
    
    processados = 0
    erros = 0
    duplicados = 0
    
    if tem_xmls:
        print(f"\nProcessando {total_arquivos} arquivos XML de {PASTA_ORIGEM}...")

        with tqdm(total=total_arquivos, unit='arquivo', desc="Separando CT-es") as progresso:
            for raiz, _, arquivos in os.walk(PASTA_ORIGEM):
                for arquivo in arquivos:
                    if arquivo.lower().endswith('.xml'):
                        caminho_completo = os.path.join(raiz, arquivo)
                        try:
                            tree = ET.parse(caminho_completo)
                            root = tree.getroot()
                            ns = {'ns': 'http://www.portalfiscal.inf.br/cte'}
                            
                            cnpj = root.find('.//ns:emit/ns:CNPJ', ns).text
                            
                            # Obtém a data de emissão no formato AAAA-MM-DD
                            dhEmi = root.find('.//ns:ide/ns:dhEmi', ns).text
                            data_emissao = dhEmi.split("T")[0]  # Pega só a parte da data

                            pasta_cnpj = os.path.join(PASTA_DESTINO, cnpj)
                            pasta_data = os.path.join(pasta_cnpj, data_emissao)
                            os.makedirs(pasta_data, exist_ok=True)
                            
                            caminho_destino = os.path.join(pasta_data, arquivo)
                            
                            if os.path.exists(caminho_destino):
                                if renomear_arquivo_existente(caminho_destino, PASTA_DUPLICADOS):
                                    duplicados += 1
                            
                            shutil.move(caminho_completo, caminho_destino)
                            processados += 1
                            
                        except Exception as e:
                            erros += 1
                            shutil.move(caminho_completo, os.path.join(PASTA_ERROS, arquivo))
                        
                        progresso.update(1)
                        progresso.set_postfix({'OK': processados, 'Erros': erros, 'Duplicados': duplicados})

        # Relatório final
        relatorio_cnpj = gerar_relatorio_por_cnpj(PASTA_DESTINO)
        validador = contar_xmls(PASTA_DESTINO)
        """Cria um arquivo de registro dos xmls separados"""
        RELATORIO_DIR = os.path.join(PASTA_DESTINO, f"0.relatorio.txt")
        with open(RELATORIO_DIR, 'w') as f:
            f.write(f"Data do processamento: {time.strftime('%d/%m/%Y %H:%M:%S')}\n")
            f.write(f"Total de arquivos XML encontrados: {total_arquivos}\n")
            f.write(f"Total de arquivos XML processados: {processados}\n")
            f.write(f"Total de arquivos XML com erros: {erros}\n")
            f.write(f"Total de arquivos XML duplicados: {duplicados}\n")
            f.write(f"Total de arquivos XML efetivamente separados: {validador}\n")
            f.write("|"+"--"*30 +"|"+"\n"*3)
            f.write("Relatório de arquivos separados por CNPJ e Data de Emissão:\n\n")
            f.write(relatorio_cnpj)

        # Remove pastas vazias
        for raiz, dirs, _ in os.walk(PASTA_ORIGEM, topdown=False):
            for dir in dirs:
                dir_path = os.path.join(raiz, dir)
                if dir not in ["0.Erros", "1.Duplicados"]:
                    try:
                        if not os.listdir(dir_path):
                            os.rmdir(dir_path)
                    except:
                        pass
        mensagem_pos_separacao = f"Todos os {total_arquivos} arquivos XML foram separados.\n\n"
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
        criar_arquivo_log_erros(PASTA_ERROS, erros)

    # Popup final com resultados
    tempo_total = time.time() - inicio
    mensagem_final = (f"Processo finalizado!\n\n"
                     f"✓ Arquivos processados: {processados}\n"
                     f"✗ Arquivos com erro: {erros}\n"
                     f"👥 Arquivos duplicados: {duplicados}\n"
                     f"📦 Lotes compactados: {lotes_compactados}\n"
                     f"⚙ Opção: {opcao}\n"
                     f"⏱ Tempo total: {tempo_total:.2f}s\n\n"
                     f"Deseja abrir o relatório final detalhado agora?")
    
    print("\n" + "="*50)
    # Versão ASCII para o console
    mensagem_console = mensagem_final.replace('✓', '[PROCESSADOS]').replace('✗', '[ERRO]') \
                                     .replace('👥', '[DUPLICADOS]').replace('📦', '[COMPACTADOS]') \
                                     .replace('⚙', '[OPÇÃO]').replace('⏱', '[TEMPO]')
    print(mensagem_console)
    print("="*50)
    
    if mostrar_popup_confirmacao(mensagem_final, "Processo Concluído"):
            abrir_arquivos()

if __name__ == "__main__":
    configurar_encoding()
    criar_pastas_necessarias()

    print("=== ORGANIZADOR DE CT-es PORTÁTIL ===")
    print(f"Local do script: {SCRIPT_DIR}")
    print(f"Pasta origem (XMLs): {PASTA_ORIGEM}")
    print(f"Pasta destino (CNPJ): {PASTA_DESTINO}")
    print(f"Pasta duplicados: {PASTA_DUPLICADOS}\n")

    inicio = time.time()
    organizar_cte_por_emitente()