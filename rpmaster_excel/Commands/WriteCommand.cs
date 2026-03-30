using System;
using System.Collections.Generic;
using ClosedXML.Excel;

namespace rpmaster_excel.Commands
{
    /// <summary>
    /// Comando de escrita: valor ou fórmula em célula/range.
    /// </summary>
    public static class WriteCommand
    {
        public static CommandResult Execute(Dictionary<string, string> args)
        {
            var file = args.GetValueOrDefault("--file");
            var sheet = args.GetValueOrDefault("--sheet");
            var cell = args.GetValueOrDefault("--cell");
            var range = args.GetValueOrDefault("--range");
            var value = args.GetValueOrDefault("--value");
            var isFormula = args.ContainsKey("--formula");
            var output = args.GetValueOrDefault("--output"); // salvar em outro arquivo

            if (string.IsNullOrEmpty(value))
                return CommandResult.Error("write", "O parâmetro --value é obrigatório.");

            using (var engine = new ExcelEngine())
            {
                engine.Open(file);
                var ws = engine.GetWorksheet(sheet);
                int cellsWritten = 0;

                if (!string.IsNullOrEmpty(cell))
                {
                    WriteToCell(ws, cell, value, isFormula);
                    cellsWritten = 1;
                }
                else if (!string.IsNullOrEmpty(range))
                {
                    cellsWritten = WriteToRange(ws, range, value, isFormula);
                }
                else
                {
                    return CommandResult.Error("write", "Especifique --cell ou --range.");
                }

                engine.Save(output);

                var data = new Dictionary<string, object>
                {
                    { "cells_written", cellsWritten },
                    { "target", cell ?? range },
                    { "is_formula", isFormula }
                };

                return CommandResult.Ok("write", data, $"{cellsWritten} célula(s) escrita(s) com sucesso.");
            }
        }

        private static void WriteToCell(IXLWorksheet ws, string cellAddress, string value, bool isFormula)
        {
            var c = ws.Cell(cellAddress);
            if (isFormula)
            {
                c.FormulaA1 = value.StartsWith("=") ? value.Substring(1) : value;
            }
            else
            {
                c.Value = value;
            }
        }

        private static int WriteToRange(IXLWorksheet ws, string rangeAddress, string value, bool isFormula)
        {
            var rng = ws.Range(rangeAddress);
            int count = 0;

            foreach (var row in rng.Rows())
            {
                foreach (var cell in row.Cells())
                {
                    if (isFormula)
                    {
                        cell.FormulaA1 = value.StartsWith("=") ? value.Substring(1) : value;
                    }
                    else
                    {
                        cell.Value = value;
                    }
                    count++;
                }
            }

            return count;
        }
    }
}
