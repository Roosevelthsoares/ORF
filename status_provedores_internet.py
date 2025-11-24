#!/usr/bin/env python3
"""
Script de Monitoramento de Provedores de Internet e BBI
Gera relatório em formato ODT com status das conexões

Instalação:
  sudo apt install python3-odf mtr-tiny -y
  
Ou com venv:
  python3 -m venv venv
  source venv/bin/activate
  pip install odfpy
"""

import subprocess
import re
import sys
from datetime import datetime

try:
    from odf.opendocument import OpenDocumentText
    from odf.style import Style, TableColumnProperties, TableCellProperties, ParagraphProperties, TextProperties
    from odf.text import P
    from odf.table import Table, TableColumn, TableRow, TableCell
except ImportError:
    print("ERRO: Biblioteca odfpy não encontrada!")
    print("\nPara instalar, escolha uma opção:\n")
    print("Opção 1 (recomendado para WSL/Ubuntu):")
    print("  sudo apt update")
    print("  sudo apt install python3-odf mtr-tiny -y\n")
    print("Opção 2 (usando venv):")
    print("  python3 -m venv venv")
    print("  source venv/bin/activate")
    print("  pip install odfpy\n")
    sys.exit(1)

# Configurações dos links a serem testados
PROVEDORES = [
    {
        "link": "EMPRESA CLARO*",
        "teste": "ping 200.213.232.73",
        "link_wan": "INTERNET",
        "timeout": 5,
        "count": 4
    },
    {
        "link": "EMPRESA NWHERE*",
        "teste": "ping 186.233.94.33",
        "link_wan": "INTERNET",
        "timeout": 5,
        "count": 4
    },
    {
        "link": "ANÚNCIO BGP*",
        "teste": "mtr 177.8.82.1",
        "link_wan": "ENW",
        "timeout": 5,
        "count": 4
    },
    {
        "link": "4ªCTA ⇔ 7ªCTA**",
        "teste": "mtr 10.67.4.35",
        "link_wan": "METRO",
        "timeout": 5,
        "count": 4,
        "observacao": "(deve passar por 172.30.192.129)"
    },
    {
        "link": "4ªCTA ⇔ 6ªCTA**",
        "teste": "mtr 10.56.67.163",
        "link_wan": "METRO",
        "timeout": 5,
        "count": 4,
        "observacao": "(deve passar por 172.30.192.101)"
    },
    {
        "link": "4ªCTA ⇔ 41ªCT**",
        "teste": "mtr 10.89.36.95",
        "link_wan": "METRO",
        "timeout": 5,
        "count": 4,
        "observacao": "(deve passar por 172.30.192.194)"
    }
]

TUNEIS = [
    {
        "tunel": "4ª CTA ⇔ HMAM (Cisco)",
        "teste": "mtr 10.79.1.46",
        "observacao": "(observar se chega ao destino com apenas 3 saltos pela vpn)",
        "timeout": 5,
        "count": 4
    }
]


def executar_ping(host, count=4, timeout=5):
    """Executa comando ping e retorna sucesso/falha"""
    try:
        cmd = ['ping', '-c', str(count), '-W', str(timeout), host]
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=timeout+5)
        return result.returncode == 0, result.stdout
    except Exception as e:
        return False, str(e)


def executar_mtr(host, count=4, timeout=5):
    """Executa comando mtr e retorna sucesso/falha"""
    try:
        # Tentar mtr primeiro
        cmd = ['mtr', '-r', '-c', str(count), '-w', host]
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=timeout+10)
        return result.returncode == 0, result.stdout
    except FileNotFoundError:
        # Se mtr não estiver disponível, usar traceroute como fallback
        print(f"    [mtr não disponível, usando traceroute]")
        try:
            cmd = ['traceroute', '-m', '15', '-w', str(timeout), host]
            result = subprocess.run(cmd, capture_output=True, text=True, timeout=timeout+10)
            return result.returncode == 0, result.stdout
        except Exception as e:
            return False, f"traceroute error: {str(e)}"
    except Exception as e:
        return False, str(e)


def executar_teste(teste_cmd, timeout=5, count=4):
    """Executa o comando de teste apropriado"""
    partes = teste_cmd.split()
    comando = partes[0]
    host = partes[1] if len(partes) > 1 else ""
    
    if comando == "ping":
        return executar_ping(host, count, timeout)
    elif comando == "mtr":
        return executar_mtr(host, count, timeout)
    else:
        return False, f"Comando desconhecido: {comando}"


def criar_estilos(doc):
    # Estilo para célula com fundo verde (sucesso) + bordas
    style_green = Style(name="CellGreen", family="table-cell")
    style_green.addElement(TableCellProperties(
        backgroundcolor="#00ff00",
        border="0.74pt solid #000000"
    ))
    doc.automaticstyles.addElement(style_green)

    # Estilo para célula com fundo vermelho (erro) + bordas
    style_red = Style(name="CellRed", family="table-cell")
    style_red.addElement(TableCellProperties(
        backgroundcolor="#ff0000",
        border="0.74pt solid #000000"
    ))
    doc.automaticstyles.addElement(style_red)

    # Estilo para célula de cabeçalho (cinza) + bordas
    style_header = Style(name="CellHeader", family="table-cell")
    style_header.addElement(TableCellProperties(
        backgroundcolor="#cccccc",
        border="0.74pt solid #000000"
    ))
    doc.automaticstyles.addElement(style_header)

    # Estilo de texto em negrito
    style_bold = Style(name="BoldText", family="text")
    style_bold.addElement(TextProperties(fontweight="bold"))
    doc.automaticstyles.addElement(style_bold)

    # Estilo de texto centralizado
    style_center = Style(name="CenterText", family="paragraph")
    style_center.addElement(ParagraphProperties(textalign="center"))
    doc.styles.addElement(style_center)

    # Estilo de texto centralizado + negrito
    style_center_bold = Style(name="CenterBoldText", family="paragraph")
    style_center_bold.addElement(ParagraphProperties(textalign="center"))
    style_center_bold.addElement(TextProperties(fontweight="bold"))
    doc.styles.addElement(style_center_bold)

    # Estilo com bordas completas
    style_border = Style(name="CellBorder", family="table-cell")
    style_border.addElement(TableCellProperties(
        border="0.74pt solid #000000"
    ))
    doc.automaticstyles.addElement(style_border)

    return (
        style_green, style_red, style_header,
        style_bold, style_center, style_center_bold,
        style_border
    )



def criar_relatorio_odt(resultados_provedores, resultados_tuneis, arquivo_saida="relatorio_rede.odt"):
    """Cria o documento ODT com os resultados"""
    doc = OpenDocumentText()
    
    # Criar estilos
    style_green, style_red, style_header, style_bold, style_center, style_center_bold, style_border = criar_estilos(doc)
    
    # Título principal
    titulo = P(stylename=style_center_bold, text="1. STATUS DE PROVEDORES DE INTERNET E BBI")
    doc.text.addElement(titulo)

    doc.text.addElement(P(text=""))  # espaço
    
    # Tabela 1 - Status de Provedores
    tabela1 = Table(name="TabelaProvedores")
    
    # Colunas
    for _ in range(4):
        tabela1.addElement(TableColumn())
    
    # Cabeçalho
# Cabeçalho da tabela de provedores
    linha_header = TableRow()
    headers = ["LINK", "TESTE", "LINK WAN", "STATUS"]

    for header in headers:
        cell = TableCell(stylename=style_header)
        cell.addElement(P(stylename=style_center_bold, text=header))
        linha_header.addElement(cell)

    tabela1.addElement(linha_header)

    
    # Linhas de dados
    for resultado in resultados_provedores:
        linha = TableRow()

        # LINK
        cell = TableCell(stylename=style_border)
        cell.addElement(P(stylename=style_center_bold, text=resultado['link']))
        linha.addElement(cell)

        # TESTE + observação
        cell = TableCell(stylename=style_border)
        texto_teste = resultado['teste']
        if 'observacao' in resultado:
            texto_teste += f"\n{resultado['observacao']}"
        cell.addElement(P(stylename=style_center, text=texto_teste))
        linha.addElement(cell)

        # LINK WAN
        cell = TableCell(stylename=style_border)
        cell.addElement(P(stylename=style_center, text=resultado['link_wan']))
        linha.addElement(cell)

        # STATUS colorido
        if resultado['sucesso']:
            cell = TableCell(stylename=style_green)
        else:
            cell = TableCell(stylename=style_red)
        cell.addElement(P(stylename=style_center_bold, text=""))
        linha.addElement(cell)

        tabela1.addElement(linha)

    
    doc.text.addElement(tabela1)
    
    # Notas de rodapé
    doc.text.addElement(P(text=""))
    doc.text.addElement(P(text="*ping / traceroute da internet para (atestar a rota ao AS via operadora)."))
    doc.text.addElement(P(text="**traceroute da EBNET."))
    
    # Seção de túneis
    doc.text.addElement(P(text=""))
    doc.text.addElement(P(stylename=style_center_bold, text="Status dos túneis configurados para contingência da REME MAO ao 4º CTA"))
    
    # Tabela de túneis
    tabela_tuneis = Table(name="TabelaTuneis")
    for _ in range(3):
        tabela_tuneis.addElement(TableColumn())
    
    linha_header_tuneis = TableRow()
    for header in ["TÚNEL", "TESTE", "STATUS"]:
        cell = TableCell(stylename=style_header)
        cell.addElement(P(stylename=style_center_bold, text=header))
        linha_header_tuneis.addElement(cell)
    tabela_tuneis.addElement(linha_header_tuneis)
    
    # Linhas de túneis
    for resultado in resultados_tuneis:
        linha = TableRow()

        # TÚNEL
        cell = TableCell(stylename=style_border)
        cell.addElement(P(stylename=style_center_bold, text=resultado['tunel']))
        linha.addElement(cell)

        # TESTE
        cell = TableCell(stylename=style_border)
        texto_teste = resultado['teste']
        if 'observacao' in resultado:
            texto_teste += f"\n{resultado['observacao']}"
        cell.addElement(P(stylename=style_center, text=texto_teste))
        linha.addElement(cell)

        # STATUS
        if resultado['sucesso']:
            cell = TableCell(stylename=style_green)
        else:
            cell = TableCell(stylename=style_red)
        cell.addElement(P(stylename=style_center_bold, text=""))
        linha.addElement(cell)

        tabela_tuneis.addElement(linha)

    
    doc.text.addElement(tabela_tuneis)
    
    # Salvar documento
    doc.save(arquivo_saida)
    print(f"\n✓ Relatório gerado: {arquivo_saida}")


def main():
    """Função principal"""
    print("=" * 60)
    print("MONITORAMENTO DE PROVEDORES E TÚNEIS")
    print("=" * 60)
    print(f"Iniciado em: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}\n")
    
    # Testar provedores
    print("Testando provedores...")
    resultados_provedores = []
    
    for provedor in PROVEDORES:
        print(f"  • {provedor['link']}...", end=" ")
        sucesso, output = executar_teste(
            provedor['teste'],
            provedor['timeout'],
            provedor['count']
        )
        
        resultado = {
            'link': provedor['link'],
            'teste': provedor['teste'],
            'link_wan': provedor['link_wan'],
            'sucesso': sucesso,
            'output': output
        }
        
        if 'observacao' in provedor:
            resultado['observacao'] = provedor['observacao']
        
        if not sucesso:
            resultado['ocorrencia'] = 'Falha na conexão'
        
        resultados_provedores.append(resultado)
        print("✓ OK" if sucesso else "✗ FALHA")
    
    # Testar túneis
    print("\nTestando túneis...")
    resultados_tuneis = []
    
    for tunel in TUNEIS:
        print(f"  • {tunel['tunel']}...", end=" ")
        sucesso, output = executar_teste(
            tunel['teste'],
            tunel['timeout'],
            tunel['count']
        )
        
        resultado = {
            'tunel': tunel['tunel'],
            'teste': tunel['teste'],
            'observacao': tunel.get('observacao', ''),
            'sucesso': sucesso,
            'output': output
        }
        
        if not sucesso:
            resultado['ocorrencia'] = 'Falha na conexão'
        
        resultados_tuneis.append(resultado)
        print("✓ OK" if sucesso else "✗ FALHA")
    
    # Gerar relatório
    print("\nGerando relatório ODT...")
    criar_relatorio_odt(resultados_provedores, resultados_tuneis)
    
    print("\n" + "=" * 60)
    print("MONITORAMENTO CONCLUÍDO")
    print("=" * 60)


if __name__ == "__main__":
    main()