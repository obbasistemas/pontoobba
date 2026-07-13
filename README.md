# Ponto OBBA

Sistema de ponto eletrônico da OBBA - Solução em Gestão Condominial.

## Arquivos

- **`app_ponto.html`** — App do funcionário. Login por código + PIN. Marca Entrada, Saída Almoço, Volta Almoço e Saída Final, com geolocalização. Sequência de marcações se ajusta automaticamente por dia da semana (seg-sex com intervalo, sábado sem intervalo, domingo e feriados como folga).
- **`painel_admin.html`** — Painel de administração OBBA. Login por usuário + senha. Quatro áreas:
  - **Hoje**: status de marcação de todos os funcionários ativos no dia
  - **Registros**: busca por funcionário/período, edição e exclusão de marcações
  - **Folha**: prévia e exportação da folha mensal no mesmo modelo do formulário em papel, com download em imagem e PDF
  - **Feriados**: cadastro de feriados (viram folga automaticamente no app do funcionário)
- **`Codigo_AppsScript.gs`** — Backend (Google Apps Script), conectado a uma planilha Google Sheets que serve como banco de dados.

## Como implantar

1. Crie uma planilha no Google Sheets.
2. Extensões → Apps Script → cole o conteúdo de `Codigo_AppsScript.gs`.
3. Rode a função `inicializarPlanilha` uma vez (cria as abas Funcionarios, Jornadas, RegistroPonto, Feriados e Administradores, com dados de exemplo e garante as colunas extras da folha mensal).
4. Implantar → Nova implantação → tipo "App da Web":
   - Executar como: **Eu**
   - Quem pode acessar: **Qualquer pessoa**
5. Copie a URL gerada (termina em `/exec`) e cole na constante `APPS_SCRIPT_URL`, no topo do `<script>`, em **ambos** `app_ponto.html` e `painel_admin.html`.

Sempre que o código do `.gs` for alterado, é necessário criar uma **nova versão** da implantação (Implantar → Gerenciar implantações → editar → Nova versão) para as mudanças valerem — a URL continua a mesma.

## Planilha (estrutura das abas)

- **Funcionarios**: Codigo, Nome, PIN, Funcao, JornadaID, Status, Depto / Setor / Secao, Endereco, Bairro, Cidade, Estado, CEP, Hr/Mes, CTPS, CBO
- **Jornadas**: JornadaID, Descricao, horários de seg-sex e sábado, tolerância
- **RegistroPonto**: ID, Codigo, Nome, Data, Tipo, Hora, Latitude, Longitude, Timestamp
- **Feriados**: Data, Descricao
- **Administradores**: Usuario, Senha, Nome

## Colunas da folha mensal

Para a folha mensal sair no mesmo modelo do papel, a aba `Funcionarios` deve conter estas colunas:

| Coluna | Uso na folha |
|---|---|
| `Codigo` | Código do funcionário |
| `Nome` | Nome do funcionário |
| `PIN` | Login do app |
| `Funcao` | Campo função |
| `JornadaID` | Jornada usada para expediente e intervalo |
| `Status` | Ativo/Inativo |
| `Depto / Setor / Secao` | Campo departamento/setor/seção |
| `Endereco` | Endereço do funcionário |
| `Bairro` | Bairro do funcionário |
| `Cidade` | Cidade do funcionário |
| `Estado` | UF/estado do funcionário |
| `CEP` | CEP do funcionário |
| `Hr/Mes` | Carga horária mensal |
| `CTPS` | Número da CTPS |
| `CBO` | Código CBO |

Se a sua planilha já existe, basta atualizar o `Codigo_AppsScript.gs` e rodar `inicializarPlanilha` novamente. O script agora acrescenta automaticamente as colunas que estiverem faltando na aba `Funcionarios`.

## Folha mensal

No painel administrativo, a aba `Folha` permite:

- escolher o funcionário
- escolher a competência
- visualizar a folha mensal no mesmo modelo do papel
- baixar a folha em imagem (`PNG`)
- baixar a folha em `PDF`

Os horários são preenchidos a partir dos registros da aba `RegistroPonto`:

- `Entrada`
- `Saída Almoço`
- `Volta Almoço`
- `Saída Final`

Sábados, domingos e feriados aparecem automaticamente na coluna de visto quando não houver expediente.

## Credenciais de exemplo (trocar antes de uso real)

- Funcionário: código `8`, PIN `1234` (Enoque Sanches Borges)
- Administração: usuário `admin`, senha `obba2026`

## Próximos passos planejados

- Cadastro de justificativas de falta/atraso
- Cálculo automático de atrasos, horas extras e faltas na folha mensal
