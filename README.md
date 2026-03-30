# rpmaster_excel — CLI API para Excel

Uma interface de linha de comando poderosa e elegante para manipular arquivos Excel (.xlsx, .xlsm) utilizando **ClosedXML** e **COM Interop** (para macros).

## 🚀 Funcionalidades

- **Leitura Inteligente**: Leia células, ranges, ou colunas/linhas inteiras (*end-to-end*).
- **Escrita Flexível**: Insira valores ou fórmulas em células individuais ou ranges.
- **Gestão de Fórmulas**: Aplique fórmulas do Excel programaticamente.
- **Filtros**: Ative ou remova AutoFilters em suas planilhas.
- **Inspeção**: Liste planilhas, tabelas e obtenha resumos do arquivo.
- **Retorno Estruturado**: Saída em **JSON** (padrão) ou **XML**.
- **Macros VBA**: Execute macros existentes (requer Excel instalado).

## 🛠️ Comandos

### 1. Read (Leitura)
Lê dados da planilha de forma flexível.

```bash
# Ler uma célula específica
rpmaster_excel.exe read --file "planilha.xlsx" --sheet "Dados" --cell A1

# Ler um range
rpmaster_excel.exe read --file "planilha.xlsx" --sheet "Dados" --range A1:B10

# Ler colunas de ponta a ponta (até a última linha usada)
rpmaster_excel.exe read --file "planilha.xlsx" --sheet "Dados" --columns A:C

# Ler linhas de ponta a ponta (até a última coluna usada)
rpmaster_excel.exe read --file "planilha.xlsx" --sheet "Dados" --rows 1:5
```

### 2. Write (Escrita)
Escreve valores ou fórmulas.

```bash
# Escrever valor
rpmaster_excel.exe write --file "planilha.xlsx" --sheet "Plan1" --cell B2 --value "Olá Mundo"

# Escrever fórmula
rpmaster_excel.exe write --file "planilha.xlsx" --sheet "Plan1" --cell C2 --value "=A1*10" --formula

# Salvar como um novo arquivo
rpmaster_excel.exe write --file "origem.xlsx" --cell A1 --value "Texto" --output "destino.xlsx"
```

### 3. Formula (Fórmulas)
Aplica fórmulas em blocos.

```bash
rpmaster_excel.exe formula --file "dados.xlsx" --range D2:D100 --expr "=B2+C2"
```

### 4. Filter (Filtros)
Gerencia o AutoFilter do Excel.

```bash
# Aplicar filtro em um range
rpmaster_excel.exe filter --file "dados.xlsx" --range A1:G100 --apply

# Remover filtros da planilha
rpmaster_excel.exe filter --file "dados.xlsx" --remove
```

### 5. Info (Informações)
Obtém metadados do arquivo.

```bash
# Listar todas as abas (sheets)
rpmaster_excel.exe info --file "vendas.xlsx" --list-sheets

# Listar todas as Tabelas (ListObjects)
rpmaster_excel.exe info --file "vendas.xlsx" --list-tables

# Resumo geral do arquivo
rpmaster_excel.exe info --file "vendas.xlsx" --summary
```

### 6. Macro (VBA)
Executa rotinas VBA. *Requer Excel instalado.*

```bash
rpmaster_excel.exe macro --file "automation.xlsm" --name "ProcessarDados" --arg1 "param1"
```

---

## 📦 Formato de Saída

### JSON (Default)
```json
{
  "success": true,
  "command": "read",
  "data": [
    { "row": 1, "col": "A", "address": "A1", "value": "ID", "type": "Text" },
    { "row": 1, "col": "B", "address": "B1", "value": "Nome", "type": "Text" }
  ],
  "message": null
}
```

### XML
Use `--format xml` para obter a saída em formato XML.

---

## 🏗️ Compilação

O projeto é baseado em **.NET Framework 4.7.2**.
Certifique-se de restaurar os pacotes NuGet antes de compilar:

```bash
nuget restore rpmaster_excel.slnx
msbuild rpmaster_excel.slnx /p:Configuration=Release
```

## 📝 Requisitos

- .NET Framework 4.7.2+
- **ClosedXML** (para manipulação de .xlsx)
- **Microsoft Excel** (apenas para o comando `macro`)
