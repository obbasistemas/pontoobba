# Ponto OBBA

Sistema de ponto eletrônico da OBBA - Solução em Gestão Condominial.

## Arquivos

- **`app_ponto.html`** — App do funcionário. Login por código + PIN. Marca Entrada, Saída Almoço, Volta Almoço e Saída Final, com geolocalização. Sequência de marcações se ajusta automaticamente por dia da semana (seg-sex com intervalo, sábado sem intervalo, domingo e feriados como folga).
- **`painel_admin.html`** — Painel de administração OBBA. Login por usuário + senha. Três áreas:
  - **Hoje**: status de marcação de todos os funcionários ativos no dia
  - **Registros**: busca por funcionário/período, edição e exclusão de marcações
  - **Feriados**: cadastro de feriados (viram folga automaticamente no app do funcionário)
- **`Codigo_AppsScript.gs`** — Backend (Google Apps Script), conectado a uma planilha Google Sheets que serve como banco de dados.

## Como implantar

1. Crie uma planilha no Google Sheets.
2. Extensões → Apps Script → cole o conteúdo de `Codigo_AppsScript.gs`.
3. Rode a função `inicializarPlanilha` uma vez (cria as abas Funcionarios, Jornadas, RegistroPonto, Feriados e Administradores, com dados de exemplo).
4. Implantar → Nova implantação → tipo "App da Web":
   - Executar como: **Eu**
   - Quem pode acessar: **Qualquer pessoa**
5. Copie a URL gerada (termina em `/exec`) e cole na constante `APPS_SCRIPT_URL`, no topo do `<script>`, em **ambos** `app_ponto.html` e `painel_admin.html`.

Sempre que o código do `.gs` for alterado, é necessário criar uma **nova versão** da implantação (Implantar → Gerenciar implantações → editar → Nova versão) para as mudanças valerem — a URL continua a mesma.

## Planilha (estrutura das abas)

- **Funcionarios**: Codigo, Nome, PIN, Funcao, JornadaID, Status
- **Jornadas**: JornadaID, Descricao, horários de seg-sex e sábado, tolerância
- **RegistroPonto**: ID, Codigo, Nome, Data, Tipo, Hora, Latitude, Longitude, Timestamp
- **Feriados**: Data, Descricao
- **Administradores**: Usuario, Senha, Nome

## Credenciais de exemplo (trocar antes de uso real)

- Funcionário: código `8`, PIN `1234` (Enoque Sanches Borges)
- Administração: usuário `admin`, senha `obba2026`

## Próximos passos planejados

- Espelho de ponto mensal (relatório com atrasos, horas extras e faltas)
- Cadastro de justificativas de falta/atraso
