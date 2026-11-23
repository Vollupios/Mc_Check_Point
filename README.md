# Sistema de Controle de Ponto

Este projeto é um sistema completo para análise e visualização de dados de controle de ponto de funcionários da Marcos Contabilidade.

## 📋 Visão Geral

O sistema é composto por dois componentes principais:
1. **Extrator de Dados** (`extrator.py`) - Script Python que lê a planilha do Cartão Ponto (.xls/.xlsx) exportada do sistema e grava os meses no `ponto.html` (também lê o formato HTML antigo)
2. **Interface Web** (`ponto.html`) - Dashboard interativo para visualização e análise dos dados

## 🔧 Como Funciona

### Arquivo HTML (`ponto.html`)

O arquivo HTML é um relatório completo e interativo que contém:

- **Dados Incorporados**: Os dados de todos os funcionários são armazenados diretamente no JavaScript do arquivo como objetos JSON
- **Interface Responsiva**: Desenvolvida com Tailwind CSS para uma experiência visual moderna
- **Visualizações Interativas**: Utiliza Chart.js para gráficos dinâmicos
- **Múltiplas Visualizações**:
  - **Resumo Geral**: Estatísticas consolidadas de todos os funcionários
  - **Análise Individual**: Dados detalhados por funcionário
  - **Validação**: Verificação de inconsistências nos registros

#### Estrutura dos Dados no HTML

```javascript
// dados de um mês (ex.: extractedMonthsData["2026-10"].data)
{
  "NOME_FUNCIONARIO": {
    "department": "DEPARTAMENTO",
    "sheetNumber": "123",
    "scheduledEntryTime": "07:30",
    "scheduledLeaveTime": "17:00",
    "records": [
      {
        "date": "2026-10-01",
        "ent1": "7:23",
        "sai2": "17:56"
      }
    ],
    // opcional: só aparece quando o horário muda em algum dia da semana
    // (chave = dia da semana: 1 = segunda ... 5 = sexta)
    "weeklySchedule": {
      "4": { "entry": "07:00", "leave": "16:00" }
    }
  }
}
```

Os meses gerados pelo extrator ficam no objeto `extractedMonthsData`, entre os marcadores
`// <<DADOS_EXTRAIDOS_INICIO>>` e `// <<DADOS_EXTRAIDOS_FIM>>`, no formato
`{ "2026-01": { "name": "Janeiro 2026", "data": { ...funcionários... } } }`.
O seletor de meses é montado automaticamente, em ordem de calendário, e o relatório abre no mês mais recente.

**Resumo anual:** a aba *Resumo Anual* consolida todos os meses lançados do ano mais recente (ex.: Jan a Out 2026):
indicadores do ano, % de entradas com atraso por mês, tabela mês a mês, % de atraso por departamento e
quantidade de atrasos por funcionário em cada mês (com filtro por departamento). O seletor também tem a
opção *Acumulado* no final, que faz as demais abas (Visão Geral, Ranking, Individual etc.) mostrarem o ano todo.
Tudo é recalculado sozinho quando um mês novo é adicionado pelo extrator.

### Sistema de Ponto

O controle de ponto funciona da seguinte forma:

#### Horários de Trabalho
- **Entrada Padrão**: Varia por funcionário (07:00, 07:30, 08:00)
- **Saída Padrão**: Varia por funcionário (17:00, 17:30, 17:50)
- **Departamentos**: Fiscal, Contábil, Pessoal, Simples Nacional, Apoio

#### Registros de Ponto
- **ent1**: Horário de entrada (primeiro registro do dia)
- **sai2**: Horário de saída (último registro do dia)
- **Formato**: HH:MM (24 horas)
- **Marcações Especiais**: 
  - `*` indica horário ajustado/corrigido
  - Registros podem ter observações especiais

#### Análises Disponíveis
1. **Pontualidade**: Comparação entre horário previsto e real
2. **Horas Extras**: Cálculo de tempo excedente
3. **Atrasos**: Identificação de chegadas tardias
4. **Frequência**: Análise de presença por período

## 🚀 Como Usar

### Atualizar com uma nova planilha do Cartão Ponto (forma recomendada)

1. Exporte o Cartão Ponto do sistema (ex.: `ponto_ate_m_s_10.xls`) e coloque na pasta do projeto.
2. Instale a dependência uma única vez: `pip install openpyxl`
3. Rode:
   - **Windows**: arraste a planilha para cima de `atualizar_ponto.bat` (ou `atualizar_ponto.bat ponto_ate_m_s_10.xls`)
   - **Linux/macOS**: `./run_extrator.sh ponto_ate_m_s_10.xls`
   - **Direto**: `python extrator.py ponto_ate_m_s_10.xls --html ponto.html`
4. Abra o `ponto.html` no navegador.

Regras usadas na leitura da planilha:
- Cada mês vira uma opção no seletor; meses já existentes com a mesma chave são substituídos, os demais são mantidos.
- **Entrada** = Ent. 1. **Saída** = última batida do dia (Saí. 3 quando houver, senão Saí. 2 ou Saí. 1).
- Dias com FÉRIAS, Feriado, ATESTAD, FOLGA, EXTERNO, GERENTE, LICENCA etc. no lugar da entrada não entram no cálculo.
- O horário previsto é o de segunda-feira; se o funcionário tiver horário diferente em algum dia útil, ele é usado nesse dia (`weeklySchedule`).
- O funcionário só aparece no mês se já tinha sido admitido e se há algo lançado no cartão naquele mês.

### Pré-requisitos

- Python 3.7+
- Bibliotecas Python:
  - `openpyxl` (planilha do Cartão Ponto)
  - `beautifulsoup4` (apenas para o formato HTML antigo)
  - `json`
  - `re`

### Instalação

1. Clone o repositório:
```bash
git clone <url-do-repositorio>
cd PONTO
```

2. Instale as dependências:
```bash
pip install beautifulsoup4
```

### Executando o Extrator

#### Método 1: Script Bash (Recomendado)
```bash
./run_extrator.sh ponto.html
```

#### Método 2: Python Direto
```bash
python extrator.py ponto.html
```

### Visualizando o Relatório

1. Abra o arquivo `ponto.html` em qualquer navegador web moderno
2. Navegue pelas diferentes abas:
   - **Resumo**: Visão geral dos dados
   - **Individual**: Análise por funcionário
   - **Validação**: Verificação de inconsistências

## 📊 Funcionalidades do Extrator

### Capacidades de Extração

O `extrator.py` é capaz de processar dois tipos de arquivos HTML:

1. **HTML com JavaScript Incorporado** (formato atual):
   - Extrai dados diretamente do objeto JavaScript `extractedMonthsData`
   - Mais rápido e confiável
   - Suporta estrutura JSON complexa

2. **HTML Tradicional com Tabelas**:
   - Parser HTML para tabelas estruturadas
   - Fallback para arquivos HTML antigos
   - Extração baseada em padrões de texto

### Saída de Dados

O extrator gera um arquivo `pontos_todos_funcionarios.json` contendo:

```json
{
  "NOME_FUNCIONARIO": {
    "department": "DEPARTAMENTO",
    "sheetNumber": "123",
    "scheduledEntryTime": "07:30",
    "scheduledLeaveTime": "17:00",
    "records": [
      {
        "date": "2026-10-01",
        "ent1": "7:23",
        "sai2": "17:56"
      }
    ]
  }
}
```

### Normalização de Dados

- **Datas**: Convertidas de DD/MM/YY para YYYY-MM-DD
- **Horários**: Mantidos no formato HH:MM
- **Nomes**: Preservados conforme original
- **Departamentos**: Categorizados automaticamente

## 📁 Estrutura do Projeto

```
PONTO/
├── extrator.py          # Script principal de extração
├── ponto.html          # Relatório interativo
├── run_extrator.sh     # Script de execução
├── atualizar_ponto.bat # Atualização no Windows (arraste a planilha)
└── README.md           # Documentação
```

## 🔍 Detalhes Técnicos

### Tecnologias Utilizadas

**Frontend (ponto.html):**
- HTML5 + CSS3
- Tailwind CSS (via CDN)
- Chart.js (via CDN)
- JavaScript ES6+

**Backend (extrator.py):**
- Python 3.7+
- BeautifulSoup4 para parsing HTML
- Regex para extração de dados
- JSON para serialização

### Compatibilidade

- **Navegadores**: Chrome, Firefox, Safari, Edge (versões modernas)
- **Sistemas**: Windows, macOS, Linux
- **Python**: 3.7 ou superior

## 🛠️ Customização

### Adicionando Novos Funcionários

Os dados são atualizados rodando o extrator com a planilha nova do Cartão Ponto (veja "Como Usar").

### Modificando Relatórios

As visualizações podem ser customizadas editando as funções JavaScript responsáveis pela geração de gráficos e tabelas.

### Configurando Horários

Os horários padrão podem ser ajustados para cada funcionário nos dados de configuração.

## 📈 Métricas e Análises

O sistema calcula automaticamente:

- **Taxa de Pontualidade**: Percentual de chegadas no horário
- **Horas Extras Médias**: Tempo médio de permanência extra
- **Padrões de Comportamento**: Análise temporal dos registros
- **Alertas de Inconsistência**: Identificação de registros anômalos

## 🤝 Contribuição

Para contribuir com o projeto:

1. Faça um fork do repositório
2. Crie uma branch para sua feature (`git checkout -b feature/nova-funcionalidade`)
3. Commit suas mudanças (`git commit -am 'Adiciona nova funcionalidade'`)
4. Push para a branch (`git push origin feature/nova-funcionalidade`)
5. Abra um Pull Request

## 📝 Licença

Este projeto é de uso interno da Marcos Contabilidade. Todos os direitos reservados.

## 📞 Suporte

Para dúvidas ou suporte técnico, entre em contato com a equipe de TI.

---

**Desenvolvido para Marcos Contabilidade** 🏢
