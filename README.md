# Bolão Copa do Mundo 2026

Scripts em Python para gerenciamento de um bolão da Copa do Mundo utilizando Google Forms e planilhas Excel.

## Funcionalidades

### bolao_pontos.py

- Baixa respostas do Google Forms
- Lê resultados dos jogos da planilha local
- Calcula pontuação dos apostadores
- Atualiza ranking automaticamente
- Destaca acertos por cores

### bolao_apostas.py

- Baixa respostas do Google Forms
- Copia apostas para a aba "apostas"
- Formata previsões em formato legível

## Requisitos

- Python 3.10+
- Google Forms com respostas exportáveis
- Planilha Excel compatível

## Instalação

Clone o projeto:

```bash
git clone https://github.com/luckstax7/bolao-copa-2026.git
cd bolao-copa-2026
```

Instale as dependências:

```bash
pip install -r requirements.txt
```

Crie o arquivo de configuração:

```bash
cp config.example.json config.json
```

Edite o arquivo:

```json
{
    "id_forms": "SEU_ID_DO_FORMS",
    "aba_forms": 0,
    "planilha_local": "Resultados_bolão.xlsx"
}
```

## Uso

Atualizar pontuação:

```bash
python bolao_pontos.py
```

Atualizar apostas:

```bash
python bolao_apostas.py
```

## Estrutura

```text
.
├── bolao_pontos.py
├── bolao_apostas.py
├── config.example.json
├── requirements.txt
├── README.md
└── .gitignore
```

## Status

### bolao_pontos.py

Estável e funcional.

### bolao_apostas.py

Funcional para a maior parte dos casos, porém ainda possui alguns pontos pendentes de validação e tratamento de exceções. alguns casos as apostas tem ficado em colunas trocadas o que pode prejudicar o andamento do programa

## Licença

MIT