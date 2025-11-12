# Indicador Operacional - CDs

Sistema web para análise de capacidade dos Centros de Distribuição.

## 🚀 Funcionalidades

- Upload de arquivo Excel com dados operacionais
- Seleção do dia de análise (1-31)
- Visualização de capacidade geral, pallet e caixas por CD
- Interface responsiva e intuitiva

## 📋 Requisitos

- Python 3.8+
- Flask
- openpyxl
- pandas

## 🔧 Instalação Local

1. Clone o repositório ou extraia os arquivos
2. Instale as dependências:
```bash
pip install -r requirements.txt
```

3. Execute a aplicação:
```bash
python app.py
```

4. Acesse no navegador: `http://localhost:5000`

## 📦 Deploy no Render

1. Crie uma conta no [Render](https://render.com)
2. Conecte seu repositório GitHub
3. Crie um novo Web Service
4. Configure:
   - **Build Command**: `pip install -r requirements.txt`
   - **Start Command**: `gunicorn app:app`
   - **Environment**: Python 3

## 📊 Estrutura do Excel

O arquivo Excel deve conter:
- Abas numeradas de 1 a 31 (dias do mês)
- Coluna B (linhas 4-11): Nomes dos CDs
- Coluna X: Valor X para cálculo de capacidade
- Coluna C: Valor C para cálculo de capacidade
- Coluna AH: Capacidade de pallet
- Coluna AM: Capacidade de caixas

**Fórmula da Capacidade Geral**: (X / C) × 100

## 🎨 Indicadores Visuais

- 🟢 Verde: Capacidade < 70%
- 🟡 Amarelo: Capacidade entre 70% e 90%
- 🔴 Vermelho: Capacidade ≥ 90%

## 📁 Estrutura do Projeto

```
Indicador Operacional/
├── app.py                 # Aplicação Flask principal
├── requirements.txt       # Dependências Python
├── templates/
│   └── index.html        # Interface web
├── static/
│   └── style.css         # Estilos CSS
└── utils/
    └── excel_processor.py # Processamento do Excel
```

## 🔒 Segurança

- Upload limitado a 16MB
- Apenas arquivos .xlsx e .xls permitidos
- Arquivos são removidos após processamento
- Nenhum dado é armazenado permanentemente

## 📝 Licença

Uso interno - Unilever
