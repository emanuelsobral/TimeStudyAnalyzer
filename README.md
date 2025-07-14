
# Plataforma de Análise de Estudo de Tempos - Aplicação Desktop

## Descrição
Esta é uma aplicação desktop em Python para análise de dados de estudo de tempos. A aplicação permite upload de arquivos Excel/CSV, processamento de dados, unificação de atividades, agrupamento customizado e análise estatística completa.

## Funcionalidades

### 1. Upload de Arquivos
- Suporte para arquivos Excel (.xlsx) e CSV (.csv)
- Visualização prévia dos dados
- Armazenamento local em banco SQLite

### 2. Mapeamento de Colunas
- Seleção interativa de colunas de atividade e tempo
- Validação de dados
- Processamento automático

### 3. Unificação de Atividades
- Detecção automática de atividades similares
- Interface para aprovação de sugestões
- Renomeação customizada de atividades unificadas

### 4. Agrupamento de Atividades
- Criação de grupos customizados
- Atribuição de atividades aos grupos
- Organização hierárquica dos dados

### 5. Análise Estatística
- Cálculos estatísticos completos (quartis, mediana, IQR)
- Detecção de outliers
- Visualização com box plots
- Normalização de tempo (segundos para minutos)

### 6. Exportação
- Exportação para Excel com dados completos
- Estrutura hierárquica (grupos e atividades)
- Formatação profissional

## Instalação

1. Instale o Python 3.7 ou superior
2. Instale as dependências:
   ```bash
   pip install -r requirements.txt
   ```

## Execução

Execute o aplicativo com:
```bash
python desktop_app.py
```

## Estrutura dos Dados

### Banco de Dados SQLite
A aplicação utiliza SQLite para armazenamento local com as seguintes tabelas:
- `files`: Arquivos uploadados
- `activities`: Atividades processadas
- `groups`: Grupos de atividades
- `statistics`: Resultados estatísticos

### Cálculos Estatísticos
- **Quartis**: Q1, Q3, Mediana
- **IQR**: Intervalo interquartil
- **Outliers**: Método de cerca (fence method)
- **Médias**: Geral e sem outliers
- **Tempo Total**: Soma de todos os tempos
- **Normalização**: Conversão segundos → minutos

## Interface do Usuário

A interface é organizada em 6 abas:
1. **Upload de Arquivos**: Seleção e visualização de arquivos
2. **Mapeamento de Colunas**: Definição de colunas relevantes
3. **Unificação de Atividades**: Consolidação de atividades similares
4. **Agrupamento**: Organização em grupos customizados
5. **Análise Estatística**: Cálculos e visualizações
6. **Exportação**: Geração de relatórios em Excel

## Tecnologias Utilizadas
- **Python 3.7+**
- **tkinter**: Interface gráfica
- **pandas**: Manipulação de dados
- **matplotlib/seaborn**: Visualizações
- **SQLite**: Banco de dados local
- **openpyxl**: Manipulação de arquivos Excel

## Vantagens da Versão Desktop
- Funcionamento offline
- Processamento local seguro
- Interface responsiva
- Armazenamento local dos dados
- Sem dependência de servidor web