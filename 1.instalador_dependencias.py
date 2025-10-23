import os
import sys
import subprocess
import platform
import tkinter as tk
from tkinter import messagebox

# Obtém o diretório onde o script está localizado
SCRIPT_DIR = os.path.dirname(os.path.abspath(sys.argv[0]))
PASTA_ORIGEM = os.path.join(SCRIPT_DIR, "1.A Separar")
PASTA_DESTINO = os.path.join(SCRIPT_DIR, "0.Por CNPJ")

def criar_pastas_necessarias():
    """Cria as pastas necessárias se não existirem"""
    os.makedirs(PASTA_ORIGEM, exist_ok=True)
    os.makedirs(PASTA_DESTINO, exist_ok=True)
    os.makedirs(os.path.join(PASTA_ORIGEM, "Erros"), exist_ok=True)

def atualizar_pip():
    """Atualiza o pip antes de outras instalações"""
    print("\nAtualizando pip...")
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "--upgrade", "pip"])
        print("✓ Pip atualizado com sucesso")
    except subprocess.CalledProcessError as e:
        print(f"✗ Falha ao atualizar pip: {str(e)}")

def verificar_instalar_dependencias():
    """Verifica e instala todas as dependências necessárias"""
    dependencias = [
        'tqdm',
        'pywin32'
    ]

    sistema_operacional = platform.system()
    python_exec = sys.executable
    falhas = []
    
    print("\n=== VERIFICANDO DEPENDÊNCIAS ===")
    print(f"Sistema Operacional: {sistema_operacional}")
    print(f"Python: {python_exec}")
    
    try:
        import pip
        atualizar_pip()
    except ImportError:
        print("\nERRO: Pip não está instalado. Instale o pip primeiro.")
        return False
    
    # Verifica e instala cada dependência
    for pacote in dependencias:
        try:
            __import__(pacote.split('==')[0])
            print(f"✔ {pacote} já está instalado")
        except ImportError:
            print(f"\nInstalando {pacote}...")
            try:
                subprocess.check_call([python_exec, "-m", "pip", "install", pacote], stdout=subprocess.DEVNULL)
                print(f"✔ {pacote} instalado com sucesso")
            except subprocess.CalledProcessError:
                print(f"✖ Falha ao instalar {pacote}")
                falhas.append(pacote)
    
    return len(falhas) == 0, falhas

def formatar_lista_falhas(falhas):
    """Formata a lista de falhas com bullets e quebras de linha"""
    if not falhas:
        return "Nenhuma falha encontrada"
    
    lista_formatada = ""
    for i, falha in enumerate(falhas, 1):
        lista_formatada += f"• {falha}"
        if i < len(falhas):
            lista_formatada += "\n"
    
    return lista_formatada
            
def mostrar_popup(mensagem, titulo="Instalação"):
    """Mostra popup informativo usando Tkinter"""
    try:
        root = tk.Tk()
        root.withdraw()  # Oculta a janela principal
        messagebox.showinfo(titulo, mensagem)
        root.destroy()
    except Exception as e:
        print(f"\nErro ao exibir popup: {str(e)}")
        print(f"\n=== {titulo} ===\n{mensagem}\n")

def main():
    print("=== CONFIGURADOR AUTOMÁTICO ===")
    print("Criando estrutura de pastas...")
    criar_pastas_necessarias()
    
    sucesso, falhas = verificar_instalar_dependencias()
    
    if sucesso:
        mensagem = "Todas as dependências foram instaladas com sucesso!\n\nAgora você pode executar o programa principal."
        print(f"\n{mensagem}")
        mostrar_popup(mensagem, "Instalação Completa")
    else:
        lista_falhas = formatar_lista_falhas(falhas)
        mensagem = (
            "Houve problemas na instalação das dependências.\n\n"
            "Dependências com falha:\n"
            f"{lista_falhas}\n\n"
            "Consulte as mensagens acima para corrigir."
        )
        
        # Versão para console
        print("\n" + "="*50)
        print("ERROS ENCONTRADOS:")
        print(lista_falhas.replace("• ", "- "))
        print("="*50)
        
        # Versão para popup
        mostrar_popup(mensagem, "Erro na Instalação")
        sys.exit(1)

if __name__ == "__main__":
    main()