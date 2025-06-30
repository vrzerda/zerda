#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import difflib
import sys
from pathlib import Path
import argparse
import tkinter as tk
from tkinter import filedialog, messagebox

def selecionar_arquivos():
    """
    Abre janelas de diálogo para o usuário selecionar os dois arquivos
    """
    # Cria uma janela root oculta
    root = tk.Tk()
    root.withdraw()  # Esconde a janela principal
    
    print("🔍 COMPARADOR DE ARQUIVOS DE TEXTO")
    print("=" * 40)
    print("📁 Selecione os arquivos para comparar...")
    
    # Seleciona o primeiro arquivo
    print("\n1️⃣ Selecione o PRIMEIRO arquivo...")
    arquivo1 = filedialog.askopenfilename(
        title="Selecione o PRIMEIRO arquivo para comparar",
        filetypes=[
            ("Arquivos de texto", "*.txt"),
            ("Todos os arquivos", "*.*")
        ]
    )
    
    if not arquivo1:
        print("❌ Nenhum arquivo selecionado. Operação cancelada.")
        root.destroy()
        return None, None, None
    
    print(f"✅ Primeiro arquivo: {Path(arquivo1).name}")
    
    # Seleciona o segundo arquivo
    print("\n2️⃣ Selecione o SEGUNDO arquivo...")
    arquivo2 = filedialog.askopenfilename(
        title="Selecione o SEGUNDO arquivo para comparar",
        filetypes=[
            ("Arquivos de texto", "*.txt"),
            ("Todos os arquivos", "*.*")
        ],
        initialdir=Path(arquivo1).parent  # Abre na mesma pasta do primeiro arquivo
    )
    
    if not arquivo2:
        print("❌ Segundo arquivo não selecionado. Operação cancelada.")
        root.destroy()
        return None, None, None
    
    print(f"✅ Segundo arquivo: {Path(arquivo2).name}")
    
    # Pergunta sobre o tipo de comparação
    escolha = messagebox.askyesnocancel(
        "Tipo de Comparação",
        "Como você quer ver a comparação?\n\n"
        "🎨 SIM = Gerar arquivo HTML (visual, lado a lado)\n"
        "📝 NÃO = Mostrar no terminal (texto simples)\n"
        "❌ CANCELAR = Sair"
    )
    
    root.destroy()
    
    if escolha is None:  # Cancelou
        print("❌ Operação cancelada pelo usuário.")
        return None, None, None
    
    return arquivo1, arquivo2, escolha

def ler_arquivo_seguro(caminho_arquivo):
    """
    Lê um arquivo de forma segura, tentando diferentes codificações
    """
    encodings = ['utf-8', 'latin-1', 'cp1252', 'iso-8859-1']
    
    for encoding in encodings:
        try:
            with open(caminho_arquivo, 'r', encoding=encoding) as f:
                conteudo = f.readlines()
            print(f"📖 Arquivo lido com codificação: {encoding}")
            return conteudo, encoding
        except UnicodeDecodeError:
            continue
        except Exception as e:
            print(f"❌ Erro ao ler arquivo: {e}")
            return None, None
    
    print(f"❌ Não foi possível ler o arquivo {caminho_arquivo} com nenhuma codificação")
    return None, None

def comparar_arquivos_terminal(arquivo1, arquivo2):
    """
    Compara dois arquivos e mostra as diferenças no terminal
    """
    print(f"\n🔄 Iniciando comparação...")
    print(f"📁 Arquivo 1: {Path(arquivo1).name}")
    print(f"📁 Arquivo 2: {Path(arquivo2).name}")
    
    # Lê os arquivos
    linhas1, enc1 = ler_arquivo_seguro(arquivo1)
    linhas2, enc2 = ler_arquivo_seguro(arquivo2)
    
    if linhas1 is None or linhas2 is None:
        print("❌ Erro ao ler um dos arquivos.")
        return
    
    print(f"📊 Arquivo 1: {len(linhas1)} linhas")
    print(f"📊 Arquivo 2: {len(linhas2)} linhas")
    
    # Verifica se são idênticos
    if linhas1 == linhas2:
        print("\n✅ OS ARQUIVOS SÃO IDÊNTICOS! Não há diferenças.")
        return
    
    print("\n" + "=" * 60)
    print("🔍 DIFERENÇAS ENCONTRADAS:")
    print("=" * 60)
    
    # Usa difflib para comparação detalhada
    differ = difflib.unified_diff(
        linhas1, 
        linhas2, 
        fromfile=f"📄 {Path(arquivo1).name}",
        tofile=f"📄 {Path(arquivo2).name}",
        lineterm='',
        n=3
    )
    
    contador_diffs = 0
    linhas_removidas = 0
    linhas_adicionadas = 0
    
    for linha in differ:
        contador_diffs += 1
        if linha.startswith('---') or linha.startswith('+++'):
            print(f"📁 {linha}")
        elif linha.startswith('@@'):
            print(f"\n📍 {linha}")
        elif linha.startswith('-') and not linha.startswith('---'):
            print(f"🔴 REMOVIDO: {linha[1:].rstrip()}")
            linhas_removidas += 1
        elif linha.startswith('+') and not linha.startswith('+++'):
            print(f"🟢 ADICIONADO: {linha[1:].rstrip()}")
            linhas_adicionadas += 1
        elif linha.startswith(' '):
            print(f"⚪ CONTEXTO: {linha[1:].rstrip()}")
    
    # Resumo
    print("\n" + "=" * 60)
    print("📈 RESUMO DAS DIFERENÇAS:")
    print(f"   🔴 Linhas removidas: {linhas_removidas}")
    print(f"   🟢 Linhas adicionadas: {linhas_adicionadas}")
    print(f"   📊 Total de mudanças: {linhas_removidas + linhas_adicionadas}")
    
    if contador_diffs == 0:
        print("⚠️ Aviso: Nenhuma diferença detectada pelo difflib, mas arquivos não são idênticos.")
        print("   Isso pode indicar diferenças sutis como quebras de linha ou codificação.")

def comparar_arquivos_html(arquivo1, arquivo2):
    """
    Gera uma comparação HTML lado a lado
    """
    print(f"\n🎨 Gerando comparação HTML...")
    
    # Lê os arquivos
    linhas1, enc1 = ler_arquivo_seguro(arquivo1)
    linhas2, enc2 = ler_arquivo_seguro(arquivo2)
    
    if linhas1 is None or linhas2 is None:
        print("❌ Erro ao ler um dos arquivos.")
        return
    
    # Gera HTML
    differ = difflib.HtmlDiff(wrapcolumn=80)
    html_diff = differ.make_file(
        linhas1, 
        linhas2, 
        fromdesc=f"{Path(arquivo1).name} ({enc1})",
        todesc=f"{Path(arquivo2).name} ({enc2})",
        context=True,
        numlines=3
    )
    
    # Salva arquivo HTML
    nome_saida = "comparacao_arquivos.html"
    try:
        with open(nome_saida, 'w', encoding='utf-8') as f:
            f.write(html_diff)
        
        print(f"✅ Comparação HTML salva em: {Path(nome_saida).absolute()}")
        print("🌐 Abra o arquivo HTML no seu navegador para ver a comparação visual!")
        
        # Tenta abrir automaticamente
        import webbrowser
        try:
            webbrowser.open(f"file://{Path(nome_saida).absolute()}")
            print("🚀 Abrindo no navegador...")
        except:
            print("📝 Abra manualmente o arquivo HTML no navegador.")
            
    except Exception as e:
        print(f"❌ Erro ao salvar arquivo HTML: {e}")

def modo_interativo():
    """
    Modo interativo com seleção gráfica de arquivos
    """
    resultado = selecionar_arquivos()
    
    if resultado[0] is None:  # Usuário cancelou
        return
    
    arquivo1, arquivo2, gerar_html = resultado
    
    print(f"\n🔄 Processando comparação...")
    print(f"📁 Arquivo 1: {arquivo1}")
    print(f"📁 Arquivo 2: {arquivo2}")
    
    if gerar_html:
        comparar_arquivos_html(arquivo1, arquivo2)
    else:
        comparar_arquivos_terminal(arquivo1, arquivo2)

def modo_linha_comando():
    """
    Modo linha de comando tradicional
    """
    parser = argparse.ArgumentParser(
        description='Compara dois arquivos de texto linha por linha',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Exemplos de uso:
  python comparador.py arquivo1.txt arquivo2.txt
  python comparador.py arquivo1.txt arquivo2.txt --html
  python comparador.py --gui (abre interface gráfica)
        """
    )
    
    parser.add_argument('arquivo1', nargs='?', help='Primeiro arquivo para comparar')
    parser.add_argument('arquivo2', nargs='?', help='Segundo arquivo para comparar')
    parser.add_argument('--html', action='store_true',
                       help='Gerar comparação HTML lado a lado')
    parser.add_argument('--gui', action='store_true',
                       help='Usar interface gráfica para seleção de arquivos')
    
    args = parser.parse_args()
    
    # Se --gui foi especificado ou nenhum arquivo foi fornecido
    if args.gui or (not args.arquivo1 or not args.arquivo2):
        modo_interativo()
        return
    
    # Limpa caminhos de aspas
    arquivo1 = args.arquivo1.strip('"\'')
    arquivo2 = args.arquivo2.strip('"\'')
    
    # Verifica se os arquivos existem
    if not Path(arquivo1).exists():
        print(f"❌ Erro: Arquivo '{arquivo1}' não encontrado")
        sys.exit(1)
    
    if not Path(arquivo2).exists():
        print(f"❌ Erro: Arquivo '{arquivo2}' não encontrado")
        sys.exit(1)
    
    if args.html:
        comparar_arquivos_html(arquivo1, arquivo2)
    else:
        comparar_arquivos_terminal(arquivo1, arquivo2)

def main():
    """
    Função principal que decide qual modo usar
    """
    # Se não há argumentos de linha de comando, usa modo interativo
    if len(sys.argv) == 1:
        modo_interativo()
    else:
        modo_linha_comando()

if __name__ == "__main__":
    main()