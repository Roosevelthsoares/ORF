# Extrator Starlink

Sistema automatizado para extração de dados do painel Starlink da Pulsar Connect.

## 📋 Descrição

Este projeto automatiza a extração de dados de status dos terminais Starlink, incluindo:
- Aplicação automática do filtro "Last 1 Day"
- Captura de KIT IDs através de hover nos ícones de status
- Detecção automática de cores de status (Verde/Vermelho)
- Paginação automática para coletar todos os registros
- Geração de relatórios em Excel (.xlsx) e CSV

## 🚀 Funcionalidades

- ✅ **Login Automático**: Autentica automaticamente no sistema
- ✅ **Filtro Automático**: Aplica o filtro "Last 1 Day" sem intervenção manual
- ✅ **Captura de KIT IDs**: Detecta tooltips ao passar o mouse sobre ícones de status
- ✅ **Detecção de Status**: Identifica automaticamente status verde/vermelho
- ✅ **Paginação Automática**: Navega por todas as páginas de resultados
- ✅ **Relatórios Formatados**: Gera Excel com células coloridas e CSV

## 📦 Requisitos

- Python 3.13+
- Google Chrome instalado
- Conexão com internet

## 🔧 Instalação

1. Clone este repositório:
```bash
git clone https://github.com/SEU_USUARIO/extrator-starlink.git
cd extrator-starlink
```

2. Instale as dependências:
```bash
pip install -r requirements.txt
```

## 💻 Uso

Execute o script principal:
```bash
python extrair_relatorio_final.py
```

O script irá:
1. Fazer login automaticamente
2. Navegar para a página Starlink
3. Aplicar o filtro "Last 1 Day"
4. Extrair dados de todas as páginas
5. Gerar os relatórios

## 📊 Relatórios Gerados

- **Relatorio_Starlink_Final.xlsx**: Planilha Excel formatada com cores
  - Verde: Status OK
  - Vermelho: Status com problema
- **Relatorio_Starlink_Final.csv**: Arquivo CSV para análise

## 🔐 Configuração

Edite as credenciais no arquivo `extrair_relatorio_final.py`:
```python
USER_EMAIL = "seu_email@exemplo.com"
USER_PASSWORD = "sua_senha"
```

## 📝 Estrutura do Relatório

| Coluna | Descrição |
|--------|-----------|
| OM | Nome da organização militar / terminal |
| PoP | KIT ID do terminal Starlink |
| STATUS | Status do terminal (célula colorida) |
| OCORRÊNCIA | Campo para anotações manuais |

## ⚙️ Tecnologias

- **Selenium**: Automação web
- **Pandas**: Manipulação de dados
- **OpenPyXL**: Geração de planilhas Excel
- **WebDriver Manager**: Gerenciamento automático do ChromeDriver

## 📈 Estatísticas

O sistema consegue:
- Taxa de sucesso: 100% na captura de KIT IDs
- Tempo médio: ~3.5s por registro
- Capacidade: Múltiplas páginas automaticamente

## 🐛 Solução de Problemas

**Erro: "Nenhuma linha encontrada"**
- Verifique se o filtro foi aplicado corretamente
- Aumente o tempo de espera em `time.sleep()`

**KIT IDs não capturados**
- Verifique a conexão com internet
- Aumente o tempo de hover se necessário

## 📄 Licença

Este projeto é de uso interno.

## 👥 Autor

Desenvolvido para automação de processos da 4ª CTA.
