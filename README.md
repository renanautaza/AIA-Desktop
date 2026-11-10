# AIA Desktop — Autaza Inspection Analytics

Aplicação desktop em Python para registro, análise e rastreabilidade de inspeções industriais, incluindo importação de planilhas, métricas de desempenho, comparação entre coletas e exportação de relatórios.

## Versão atualizada disponível

- **Versão atualizada em validação:** AIA Desktop **V1.3.0 RC**
- **Código principal da RC:** [releases/AIA_V1_3_0_RC/AIA_V1_3_0_RC.py](releases/AIA_V1_3_0_RC/AIA_V1_3_0_RC.py)
- **Status:** Release Candidate — ainda não substitui a versão oficial estável.
- **Validação local:** 31/31 verificações aprovadas.

A V1.3.0 RC destaca D-I como indicador principal, atualiza o relatório PDF, introduz rastreabilidade de CSV/Excel por linha e trata D-O com pareamento explícito por Correspondentes. Consulte o [pacote da V1.3.0 RC](releases/AIA_V1_3_0_RC/) para código, scripts, validações, checksums e documentação.

## Versão oficial estável

- **Última versão oficial:** AIA Desktop **V1.2.0**
- **Fechamento da release:** 10 de setembro de 2026
- **Indicador de referência:** [VERSAO_OFICIAL_ATUAL.md](VERSAO_OFICIAL_ATUAL.md)

A V1.2.0 segue preservada como baseline oficial. A V1.3.0 RC somente deve ser promovida a oficial após a validação de campo prevista no seu relatório.

## Histórico preservado

- **V1.3.0 RC (atualizada em validação):** [releases/AIA_V1_3_0_RC/](releases/AIA_V1_3_0_RC/)
- **V1.1.6 (base inicialmente publicada):** [src/Autaza_Inspection_Analytics.py](src/Autaza_Inspection_Analytics.py)
- **Notas e materiais históricos:** [docs/historico/](docs/historico/)

Nenhuma versão anterior foi removida ou sobrescrita.

## Requisitos

- Windows 10/11 (para os scripts .bat)
- Python 3.10 ou superior
- Dependências na pasta da versão escolhida: requirements.txt

## Execução da V1.3.0 RC

No Windows, abra a pasta releases/AIA_V1_3_0_RC e execute INSTALAR_DEPENDENCIAS_AIA.bat uma única vez; depois use INICIAR_AIA.bat.

Alternativamente, instale as dependências do pacote e execute AIA_V1_3_0_RC.py.

## Validação

Para a V1.3.0 RC, execute VALIDAR_AIA.py na pasta da release.

## Segurança e privacidade

Este repositório não inclui credenciais, tokens, dados de inspeção, caches Python, builds nem executáveis. Antes de compartilhar dados gerados pela aplicação, revise-os para evitar exposição de informações industriais sensíveis.
