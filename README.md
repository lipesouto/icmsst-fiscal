# 🧾 OmniAI Fiscal - Exclusão ICMS-ST

Aplicação web para cálculo de créditos de PIS/COFINS decorrentes da exclusão do ICMS-ST da base de cálculo.

![Python](https://img.shields.io/badge/Python-3.9+-blue.svg)
![Streamlit](https://img.shields.io/badge/Streamlit-1.28+-red.svg)
![License](https://img.shields.io/badge/License-Proprietary-green.svg)

## 🎯 Funcionalidades

- ✅ Upload de base de produtos (Excel com NCM e MVA)
- ✅ Upload de múltiplos arquivos SPED Contribuições
- ✅ Seleção de CFOPs elegíveis configurável
- ✅ Processamento em lote com barra de progresso
- ✅ Geração automática de:
  - 📊 Excel consolidado (De/Para por mês)
  - 📄 Relatório PDF executivo
  - 📦 SPEDs retificados (um por mês)
  - 🔧 JSON para integração via API

## 🚀 Instalação Local

### Pré-requisitos

- Python 3.9+
- pip

### Passos

```bash
# Clone o repositório
git clone https://github.com/omniai/fiscal-icmsst.git
cd fiscal-icmsst

# Crie um ambiente virtual
python -m venv venv
source venv/bin/activate  # Linux/Mac
# ou
venv\Scripts\activate  # Windows

# Instale as dependências
pip install -r requirements.txt

# Execute a aplicação
streamlit run app.py
```

A aplicação estará disponível em `http://localhost:8501`

## 🐳 Deploy com Docker

```bash
# Build da imagem
docker build -t omniai-fiscal .

# Execute o container
docker run -p 8501:8501 omniai-fiscal
```

## ☁️ Deploy no Streamlit Cloud

1. Faça fork do repositório
2. Acesse [share.streamlit.io](https://share.streamlit.io)
3. Conecte seu GitHub
4. Selecione o repositório e o arquivo `app.py`
5. Clique em "Deploy"

## 📁 Formato dos Arquivos

### Base de Produtos (Excel)

O arquivo Excel deve conter as seguintes colunas (case-insensitive):

| Coluna | Obrigatório | Descrição |
|--------|-------------|-----------|
| `NCM` | Sim* | Código NCM de 8 dígitos |
| `Capitulo` | Sim* | Primeiros 4 dígitos do NCM |
| `Item` | Sim* | Últimos 4 dígitos do NCM |
| `MVA` ou `IVA/MVA` | Sim | Margem de Valor Agregado (%) |
| `Aliquota Entrada` | Não | Alíquota ICMS (default: 18%) |

*NCM pode ser informado diretamente OU reconstruído de Capitulo+Item

### Arquivos SPED

- Formato: SPED Contribuições (TXT)
- Encoding: Latin-1 (ISO-8859-1)
- Nomenclatura sugerida: `SPED_CONTRIBUICOES_MM_YYYY.txt`

## ⚙️ Configurações

### CFOPs Elegíveis

Por padrão, apenas o CFOP 5405 está selecionado:

| CFOP | Descrição | Uso |
|------|-----------|-----|
| **5405** | Venda ST - Substituído | ✅ Default |
| 5403 | Venda ST - Substituto | ⬜ Opcional |
| 5401 | Venda Produção ST | ⬜ Opcional |
| 5102 | Venda Revenda | ⬜ Opcional |

## 📊 Metodologia de Cálculo

1. **Identificação**: Registros C870 com CFOPs selecionados
2. **Enriquecimento**: Associação NCM → MVA via base de produtos
3. **Cálculo ICMS-ST**:
   - Base ICMS-ST = Valor Item × (1 + MVA%)
   - ICMS-ST = Base ICMS-ST × Alíquota ICMS%
4. **Nova Base PIS/COFINS**: BC Original - ICMS-ST
5. **Crédito**: Diferença entre tributos originais e recalculados

## 🔒 Segurança

- Todos os dados são processados localmente no navegador/servidor
- Nenhum dado é armazenado permanentemente
- Arquivos são descartados após o processamento
- Compatível com LGPD

## 📝 Licença

Software proprietário - OmniAI © 2025

## 🤝 Suporte

Para suporte técnico ou comercial:
- Email: suporte@omniai.com.br
- WhatsApp: (11) 99999-9999
