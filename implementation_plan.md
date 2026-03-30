# rpmaster_excel — CLI API para Excel

Criar uma API de linha de comando que manipula arquivos Excel (.xlsx) via ClosedXML, com retorno em JSON.

## Arquitetura Proposta

```
Program.cs          → Parser de argumentos + roteamento de comandos
ExcelEngine.cs      → Lógica principal (abrir/salvar workbook)
Commands/
  ReadCommand.cs    → Leitura de célula, range, coluna/linha inteira, end-to-end
  WriteCommand.cs   → Escrita de valor ou fórmula em célula/range
  FormulaCommand.cs → Aplicar fórmulas
  FilterCommand.cs  → Aplicar filtros em tabelas
  InfoCommand.cs    → Listar planilhas, tabelas, informações gerais
OutputFormatter.cs  → Formatar saída (JSON / XML)
```

## Sintaxe dos Comandos

### 1. READ — Leitura

```bash
# Ler uma célula
rpmaster_excel.exe read --file "C:\plan.xlsx" --sheet "Sheet1" --cell "A1" --format json

# Ler um range
rpmaster_excel.exe read --file "C:\plan.xlsx" --sheet "Sheet1" --range "A1:D10" --format json

# Ler coluna inteira (end-to-end, até última linha com dados)
rpmaster_excel.exe read --file "C:\plan.xlsx" --sheet "Sheet1" --column "A" --format json

# Ler linha inteira (end-to-end, até última coluna com dados)
rpmaster_excel.exe read --file "C:\plan.xlsx" --sheet "Sheet1" --row 1 --format json

# Ler múltiplas colunas end-to-end (ex: A até D)
rpmaster_excel.exe read --file "C:\plan.xlsx" --sheet "Sheet1" --columns "A:D" --format json

# Ler múltiplas linhas end-to-end
rpmaster_excel.exe read --file "C:\plan.xlsx" --sheet "Sheet1" --rows "1:10" --format json
```

### 2. WRITE — Escrita

```bash
# Escrever valor em célula
rpmaster_excel.exe write --file "C:\plan.xlsx" --sheet "Sheet1" --cell "A1" --value "Hello"

# Escrever fórmula em célula
rpmaster_excel.exe write --file "C:\plan.xlsx" --sheet "Sheet1" --cell "A1" --value "=SUM(B1:B10)" --formula

# Escrever valor em range (preenche todas as células do range)
rpmaster_excel.exe write --file "C:\plan.xlsx" --sheet "Sheet1" --range "A1:A10" --value "Test"
```

### 3. FORMULA — Aplicar fórmulas

```bash
# Aplicar fórmula em célula
rpmaster_excel.exe formula --file "C:\plan.xlsx" --sheet "Sheet1" --cell "C1" --expr "=A1+B1"

# Aplicar fórmula em range (fórmula relativa será ajustada)
rpmaster_excel.exe formula --file "C:\plan.xlsx" --sheet "Sheet1" --range "C1:C10" --expr "=A1+B1"
```

### 4. FILTER — Filtros

```bash
# Aplicar autofilter em um range
rpmaster_excel.exe filter --file "C:\plan.xlsx" --sheet "Sheet1" --range "A1:D100" --apply

# Remover filtro
rpmaster_excel.exe filter --file "C:\plan.xlsx" --sheet "Sheet1" --remove
```

### 5. INFO — Informações

```bash
# Listar todas as planilhas
rpmaster_excel.exe info --file "C:\plan.xlsx" --list-sheets --format json

# Listar todas as tabelas (ClosedXML Tables)
rpmaster_excel.exe info --file "C:\plan.xlsx" --list-tables --format json

# Informações gerais do arquivo
rpmaster_excel.exe info --file "C:\plan.xlsx" --summary --format json
```

### 6. MACRO — Chamar Macro (via ClosedXML, limitado a manipulação de dados)

```bash
rpmaster_excel.exe macro --file "C:\plan.xlsx" --name "MinhaRotina"
```

> [!IMPORTANT]
> **ClosedXML NÃO suporta execução de macros VBA diretamente.** ClosedXML trabalha apenas com arquivos `.xlsx` (sem macros). Para executar macros VBA, seria necessário usar COM Interop com Excel instalado. Posso implementar um comando `macro` que execute via COM Interop (precisa do Excel instalado na máquina), ou posso pular esse comando. O que prefere?

## Formato de Retorno

### JSON (padrão)
```json
{
  "success": true,
  "command": "read",
  "data": [
    {"row": 1, "col": "A", "value": "Nome"},
    {"row": 1, "col": "B", "value": "Idade"}
  ],
  "message": null
}
```

### Em caso de erro
```json
{
  "success": false,
  "command": "read",
  "data": null,
  "message": "Arquivo não encontrado: C:\\plan.xlsx"
}
```

## Proposed Changes

### Program.cs
- [MODIFY] [Program.cs](file:///c:/Users/sergio.ssilva/source/repos/rpmaster_excel/rpmaster_excel/Program.cs)
  - Parser de argumentos (`args[]`) com switch no primeiro argumento (read/write/formula/filter/info/macro)
  - Extrai flags `--file`, `--sheet`, `--cell`, `--range`, `--column`, `--row`, `--columns`, `--rows`, `--value`, `--formula`, `--format`, `--expr`, etc.
  - Try/catch global com retorno de erro em JSON
  - Exibe help se chamado sem args ou com `--help`

### [NEW] ExcelEngine.cs
- Classe para abrir/fechar workbook
- Métodos: `OpenWorkbook(path)`, `GetWorksheet(name)`, `SaveWorkbook()`

### [NEW] Commands/ReadCommand.cs
- Ler célula única
- Ler range retangular
- Ler coluna end-to-end (detecta `LastCellUsed()`)
- Ler linha end-to-end
- Ler múltiplas colunas (ex: A:D)
- Ler múltiplas linhas

### [NEW] Commands/WriteCommand.cs
- Escrever valor ou fórmula em célula
- Escrever valor em range

### [NEW] Commands/FormulaCommand.cs
- Aplicar fórmula em célula
- Aplicar fórmula em range

### [NEW] Commands/FilterCommand.cs
- Aplicar autofilter em range
- Remover filtro

### [NEW] Commands/InfoCommand.cs
- Listar planilhas
- Listar tabelas (ClosedXML tables)
- Resumo do arquivo (sheets, tamanho, etc.)

### [NEW] OutputFormatter.cs
- Serializar resultado em JSON ou XML
- Usa `Newtonsoft.Json` / `System.Text.Json`

### [MODIFY] rpmaster_excel.csproj
- Adicionar novos `.cs` ao `<ItemGroup>` de compilação

## Open Questions

> [!IMPORTANT]
> 1. **Macros VBA**: ClosedXML não executa macros. Quer que eu implemente via COM Interop (precisa do Excel instalado)? Ou posso pular o comando `macro`?

> [!NOTE]
> 2. **Formato XML**: Você mencionou `type_return:json,xml`. Quer realmente suporte a XML, ou JSON é suficiente? Se quiser XML, vou implementar com `System.Xml.Linq`.

> [!NOTE]
> 3. **Newtonsoft.Json vs System.Text.Json**: Vi que o projeto tem `System.Text.Json` instalado mas não Newtonsoft. Posso usar `System.Text.Json` mesmo? Ou quer que eu adicione Newtonsoft?

## Verification Plan

### Automated Tests
- Build do projeto com `msbuild`
- Teste manual com um arquivo `.xlsx` de exemplo

### Manual Verification
- Rodar cada comando e validar o JSON de saída
