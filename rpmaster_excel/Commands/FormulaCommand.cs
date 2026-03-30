using System;
using System.Collections.Generic;
using ClosedXML.Excel;

namespace rpmaster_excel.Commands
{
    /// <summary>
    /// Comando para aplicar fórmulas em células ou ranges.
    /// </summary>
    public static class FormulaCommand
    {
        public static CommandResult Execute(Dictionary<string, string> args)
        {
            var file = args.GetValueOrDefault("--file");
            var sheet = args.GetValueOrDefault("--sheet");
            var cell = args.GetValueOrDefault("--cell");
            var range = args.GetValueOrDefault("--range");
            var expr = args.GetValueOrDefault("--expr");
            var output = args.GetValueOrDefault("--output");

            if (string.IsNullOrEmpty(expr))
                return CommandResult.Error("formula", "O parâmetro --expr é obrigatório.");

            // Remove o '=' inicial se presente (ClosedXML não precisa dele)
            var formula = expr.StartsWith("=") ? expr.Substring(1) : expr;

            using (var engine = new ExcelEngine())
            {
                engine.Open(file);
                var ws = engine.GetWorksheet(sheet);
                int count = 0;

                if (!string.IsNullOrEmpty(cell))
                {
                    ws.Cell(cell).FormulaA1 = formula;
                    count = 1;
                }
                else if (!string.IsNullOrEmpty(range))
                {
                    var rng = ws.Range(range);
                    foreach (var row in rng.Rows())
                    {
                        foreach (var c in row.Cells())
                        {
                            c.FormulaA1 = formula;
                            count++;
                        }
                    }
                }
                else
                {
                    return CommandResult.Error("formula", "Especifique --cell ou --range.");
                }

                engine.Save(output);

                var data = new Dictionary<string, object>
                {
                    { "cells_affected", count },
                    { "formula", "=" + formula },
                    { "target", cell ?? range }
                };

                return CommandResult.Ok("formula", data, $"Fórmula aplicada em {count} célula(s).");
            }
        }
    }
}
