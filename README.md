# AIA Desktop — Autaza Inspection Analytics

Aplicação desktop em Python para registro, análise e rastreabilidade de inspeções industriais, incluindo importação de planilhas, métricas de desempenho, comparação entre coletas e exportação de relatórios.

## Base publicada

- **Versão principal:** 1.1.6
- **Arquivo principal:** `src/Autaza_Inspection_Analytics.py`
- **Base temporal:** 1º de setembro de 2026

A versão 1.1.6 é a versão mais recente disponível neste conjunto de arquivos. Materiais de versões anteriores foram preservados em [`docs/historico/`](docs/historico/).

## Requisitos

- Windows 10/11 (para os scripts `.bat`)
- Python 3.10 ou superior
- Dependências em `requirements.txt`

## Instalação e execução

```powershell
py -3 -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
pip install -r requirements.txt
python src\Autaza_Inspection_Analytics.py
```

No Prompt de Comando do Windows, após instalar as dependências, também é possível executar `scripts\run_windows.bat`.

## Validação

Execute a verificação estática básica com:

```powershell
python tests\test_smoke.py
```

## Geração do executável Windows

Depois de instalar as dependências, inclusive `pyinstaller`, execute:

```powershell
python scripts\build_windows.py
```

O pacote será criado em `dist/`. Artefatos de compilação e executáveis não são versionados.

## Estrutura

- `src/` — código-fonte principal
- `assets/` — ícones e imagens usados pela aplicação
- `scripts/` — execução e empacotamento
- `tests/` — validações reproduzíveis
- `docs/historico/` — notas de versões anteriores

## Segurança e privacidade

Este repositório não inclui credenciais, tokens, arquivos temporários, dados de inspeção, builds nem executáveis. Antes de compartilhar dados gerados pela aplicação, revise-os para evitar a exposição de informações industriais sensíveis.
