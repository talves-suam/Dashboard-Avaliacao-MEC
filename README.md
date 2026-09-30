# Dashboard Avaliação MEC — UNISUAM

Web app para **Google Apps Script**, com interface institucional (laranja/azul UNISUAM) e persistência em **JSON no Google Drive** (sem serviços externos pagos).

## Sobre armazenamento de dados

| Opção | Viável? | Observação |
|--------|---------|------------|
| Dados embutidos no código (`SeedData.gs`) | Parcial | Bom como *seed* inicial. Ruim para manutenção contínua (exige republicar o app a cada alteração). |
| **JSON em pasta do Drive** (implementado) | **Sim** | Gratuito, nativo do Apps Script, sem rede externa. Pasta: `Dashboard Avaliacao MEC - Dados` → `cursos_data.json`. |
| Google Sheets como banco | Sim | Alternativa gratuita nativa; útil se a equipe preferir editar em planilha. |
| PropertiesService | Não para este volume | Limite ~9 KB por propriedade. |

**Conclusão:** manter os dados no Drive (como pedido) é a melhor opção gratuita e sem dependência externa. O `SeedData.gs` só inicializa a base na primeira execução.

## Arquivos do projeto

| Arquivo | Função |
|---------|--------|
| `Code.gs` | Backend: `doGet`, CRUD, importação Excel via Drive API |
| `SeedData.gs` | Seed com 32 cursos da planilha base (metadados) |
| `Index.html` | Shell da página |
| `Stylesheet.html` | CSS |
| `JavaScript.html` | Lógica do dashboard |
| `LogoData.html` | Logo UNISUAM em base64 |
| `appsscript.json` | Manifesto (timezone, Drive advanced service) |

## Deploy no Google Apps Script

1. Acesse [script.google.com](https://script.google.com) → **Novo projeto**.
2. Renomeie o projeto (ex.: `Dashboard Avaliacao MEC`).
3. Crie os arquivos com os mesmos nomes e cole o conteúdo de cada um.
4. Em **Serviços** (+), ative **Drive API** (serviço avançado `Drive`).
5. **Implantar** → **Nova implantação** → tipo **App da Web**:
   - Executar como: **Eu**
   - Quem tem acesso: domínio UNISUAM / usuários do Google conforme a política interna
6. Abra a URL da implantação. Na primeira carga, o script cria a pasta e o JSON no Drive.

### Importação Excel

O botão **Incluir curso** aceita `.xlsx` com as abas:

- `Painel Conceito`
- `Painel Indicadores`
- `Painel de Detalhes` (agregados são recalculados no app; a aba não é obrigatória)

O arquivo anexado `Cópia de Visitas 2023 - 2024 - 2025.xlsx` contém **fórmulas** (`VLOOKUP` / `FILTER`) apontando para uma aba `Indicadores` que **não está na cópia** — por isso Conceito e notas aparecem como `#REF!`.  

Para popular as notas, importe uma planilha com **valores calculados** (cole especial → valores) ou a planilha original completa com a aba `Indicadores`.

## Funcionalidades

- Filtros: Código e-Mec, Curso, Unidade, Ano Visita, Ato Regulatório, Conceito Final Faixa
- Visão agregada: KPIs nota 5 / 4 / &lt; 4 + avaliações por ano + indicadores extremos
- Visão por código/curso: KPIs Conceito Contínuo e Faixa + apenas Painel de Indicadores
- Painéis: Conceito, Indicadores (1.1–3.18), Detalhes (&lt; 3 por ano / com CC &lt; 3)
- Cadastro por Excel ou formulário manual

## Preview local

Abra `preview.html` no navegador para validar o layout (usa dados demo; importação Excel só funciona no Apps Script).
