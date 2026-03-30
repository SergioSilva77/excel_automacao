using System;
using System.Collections.Generic;
using System.Linq;
using rpmaster_excel.Commands;

namespace rpmaster_excel
{
    internal class Program
    {
        static int Main(string[] args)
        {
            try
            {
                if (args.Length == 0 || args[0] == "--help" || args[0] == "-h" || args[0] == "help")
                {
                    ShowHelp();
                    return 0;
                }

                var command = args[0].ToLowerInvariant();
                var parsedArgs = ParseArguments(args.Skip(1).ToArray());
                var format = parsedArgs.GetValueOrDefault("--format") ?? "json";

                CommandResult result;

                switch (command)
                {
                    case "read":
                        result = ReadCommand.Execute(parsedArgs);
                        break;

                    case "write":
                        result = WriteCommand.Execute(parsedArgs);
                        break;

                    case "formula":
                        result = FormulaCommand.Execute(parsedArgs);
                        break;

                    case "filter":
                        result = FilterCommand.Execute(parsedArgs);
                        break;

                    case "info":
                        result = InfoCommand.Execute(parsedArgs);
                        break;

                    case "macro":
                        result = MacroCommand.Execute(parsedArgs);
                        break;

                    default:
                        result = CommandResult.Error(command, $"Comando desconhecido: '{command}'. Use --help para ver os comandos disponíveis.");
                        break;
                }

                Console.OutputEncoding = System.Text.Encoding.UTF8;
                Console.WriteLine(OutputFormatter.Format(result, format));

                return result.Success ? 0 : 1;
            }
            catch (Exception ex)
            {
                var errorResult = CommandResult.Error(
                    args.Length > 0 ? args[0] : "unknown",
                    $"{ex.GetType().Name}: {ex.Message}"
                );
                Console.OutputEncoding = System.Text.Encoding.UTF8;
                Console.WriteLine(OutputFormatter.Format(errorResult, "json"));
                return 1;
            }
        }

        /// <summary>
        /// Parseia argumentos no formato --key value ou --flag.
        /// </summary>
        static Dictionary<string, string> ParseArguments(string[] args)
        {
            var dict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            for (int i = 0; i < args.Length; i++)
            {
                var arg = args[i];

                if (arg.StartsWith("--"))
                {
                    // Flags sem valor (true/false)
                    if (i + 1 >= args.Length || args[i + 1].StartsWith("--"))
                    {
                        dict[arg.ToLowerInvariant()] = "true";
                    }
                    else
                    {
                        dict[arg.ToLowerInvariant()] = args[i + 1];
                        i++; // pula o valor
                    }
                }
            }

            return dict;
        }

        static void ShowHelp()
        {
            var help = @"
╔══════════════════════════════════════════════════════════════════╗
║                    rpmaster_excel - CLI API                      ║
║              Manipulação de arquivos Excel (.xlsx)                ║
╚══════════════════════════════════════════════════════════════════╝

COMANDOS DISPONÍVEIS:
─────────────────────

  read      Lê dados de células, ranges, colunas ou linhas
  write     Escreve valores ou fórmulas em células/ranges
  formula   Aplica fórmulas em células ou ranges
  filter    Aplica ou remove filtros (AutoFilter)
  info      Lista planilhas, tabelas e informações do arquivo
  macro     Executa macros VBA (requer Excel instalado)

═══════════════════════════════════════════════════════════════════

READ — Leitura de dados
───────────────────────
  rpmaster_excel.exe read --file <caminho> --sheet <nome> [opções]

  --cell <endereço>       Ler uma célula (ex: A1)
  --range <range>         Ler um range (ex: A1:D10)
  --column <letra>        Ler coluna inteira end-to-end (ex: A)
  --columns <range>       Ler múltiplas colunas (ex: A:D)
  --row <número>          Ler linha inteira end-to-end (ex: 1)
  --rows <range>          Ler múltiplas linhas (ex: 1:10)
  --format <json|xml>     Formato de saída (padrão: json)

WRITE — Escrita de dados
────────────────────────
  rpmaster_excel.exe write --file <caminho> --sheet <nome> [opções]

  --cell <endereço>       Célula destino (ex: A1)
  --range <range>         Range destino (ex: A1:A10)
  --value <valor>         Valor a escrever
  --formula               Flag: trata --value como fórmula
  --output <caminho>      Salvar em outro arquivo (opcional)
  --format <json|xml>     Formato de saída (padrão: json)

FORMULA — Aplicar fórmulas
──────────────────────────
  rpmaster_excel.exe formula --file <caminho> --sheet <nome> [opções]

  --cell <endereço>       Célula destino
  --range <range>         Range destino
  --expr <fórmula>        Expressão (ex: =SUM(A1:A10))
  --output <caminho>      Salvar em outro arquivo (opcional)
  --format <json|xml>     Formato de saída (padrão: json)

FILTER — Gerenciar filtros
──────────────────────────
  rpmaster_excel.exe filter --file <caminho> --sheet <nome> [opções]

  --range <range>         Range para aplicar filtro (ex: A1:D100)
  --apply                 Aplicar AutoFilter
  --remove                Remover filtro existente
  --output <caminho>      Salvar em outro arquivo (opcional)
  --format <json|xml>     Formato de saída (padrão: json)

INFO — Informações do arquivo
─────────────────────────────
  rpmaster_excel.exe info --file <caminho> [opções]

  --list-sheets           Lista todas as planilhas
  --list-tables           Lista todas as tabelas
  --sheet <nome>          Filtrar tabelas por planilha (opcional)
  --summary               Resumo geral do arquivo
  --format <json|xml>     Formato de saída (padrão: json)

MACRO — Executar macro VBA (COM Interop)
────────────────────────────────────────
  rpmaster_excel.exe macro --file <caminho> [opções]

  --name <nome>           Nome da macro (ex: MinhaRotina)
  --arg1 <valor>          Argumento 1 (opcional)
  --arg2 <valor>          Argumento 2 (opcional)
  --arg3 <valor>          Argumento 3 (opcional)
  --format <json|xml>     Formato de saída (padrão: json)

═══════════════════════════════════════════════════════════════════

EXEMPLOS:
─────────

  # Ler célula A1
  rpmaster_excel.exe read --file ""C:\plan.xlsx"" --sheet ""Sheet1"" --cell A1

  # Ler colunas A até D (end-to-end)
  rpmaster_excel.exe read --file ""C:\plan.xlsx"" --sheet ""Sheet1"" --columns A:D

  # Escrever fórmula
  rpmaster_excel.exe write --file ""C:\plan.xlsx"" --sheet ""Sheet1"" --cell A1 --value ""=SUM(B1:B10)"" --formula

  # Listar planilhas em XML
  rpmaster_excel.exe info --file ""C:\plan.xlsx"" --list-sheets --format xml

  # Executar macro VBA
  rpmaster_excel.exe macro --file ""C:\plan.xlsm"" --name ""MinhaRotina""
";
            Console.WriteLine(help);
        }
    }
}
