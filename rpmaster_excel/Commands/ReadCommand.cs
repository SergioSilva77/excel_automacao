using System;
using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel;

namespace rpmaster_excel.Commands
{
    /// <summary>
    /// Comando de leitura: célula, range, coluna(s), linha(s), end-to-end.
    /// </summary>
    public static class ReadCommand
    {
        /// <summary>
        /// Executa o comando read baseado nos argumentos parseados.
        /// </summary>
        public static CommandResult Execute(Dictionary<string, string> args)
        {
            var file = args.GetValueOrDefault("--file");
            var sheet = args.GetValueOrDefault("--sheet");
            var cell = args.GetValueOrDefault("--cell");
            var range = args.GetValueOrDefault("--range");
            var column = args.GetValueOrDefault("--column");
            var row = args.GetValueOrDefault("--row");
            var columns = args.GetValueOrDefault("--columns");
            var rows = args.GetValueOrDefault("--rows");

            using (var engine = new ExcelEngine())
            {
                engine.Open(file);
                var ws = engine.GetWorksheet(sheet);

                List<Dictionary<string, object>> data;

                if (!string.IsNullOrEmpty(cell))
                {
                    data = ReadCell(ws, cell);
                }
                else if (!string.IsNullOrEmpty(range))
                {
                    data = ReadRange(ws, range);
                }
                else if (!string.IsNullOrEmpty(column))
                {
                    data = ReadColumnEndToEnd(ws, column);
                }
                else if (!string.IsNullOrEmpty(columns))
                {
                    data = ReadColumnsEndToEnd(ws, columns);
                }
                else if (!string.IsNullOrEmpty(row))
                {
                    data = ReadRowEndToEnd(ws, int.Parse(row));
                }
                else if (!string.IsNullOrEmpty(rows))
                {
                    data = ReadRowsEndToEnd(ws, rows);
                }
                else
                {
                    return CommandResult.Error("read", "Especifique --cell, --range, --column, --columns, --row ou --rows.");
                }

                return CommandResult.Ok("read", data);
            }
        }

        /// <summary>
        /// Lê uma única célula.
        /// </summary>
        private static List<Dictionary<string, object>> ReadCell(IXLWorksheet ws, string cellAddress)
        {
            var c = ws.Cell(cellAddress);
            return new List<Dictionary<string, object>>
            {
                CellToDict(c)
            };
        }

        /// <summary>
        /// Lê um range retangular (ex: A1:D10).
        /// </summary>
        private static List<Dictionary<string, object>> ReadRange(IXLWorksheet ws, string rangeAddress)
        {
            var rng = ws.Range(rangeAddress);
            var result = new List<Dictionary<string, object>>();

            foreach (var row in rng.Rows())
            {
                foreach (var cell in row.Cells())
                {
                    result.Add(CellToDict(cell));
                }
            }

            return result;
        }

        /// <summary>
        /// Lê uma coluna end-to-end (da primeira até a última célula com dados).
        /// </summary>
        private static List<Dictionary<string, object>> ReadColumnEndToEnd(IXLWorksheet ws, string colLetter)
        {
            var result = new List<Dictionary<string, object>>();
            var lastRow = ws.LastRowUsed()?.RowNumber() ?? 0;
            if (lastRow == 0) return result;

            var colNumber = XLHelper.GetColumnNumberFromLetter(colLetter);

            for (int r = 1; r <= lastRow; r++)
            {
                var cell = ws.Cell(r, colNumber);
                if (!cell.IsEmpty())
                    result.Add(CellToDict(cell));
            }

            return result;
        }

        /// <summary>
        /// Lê múltiplas colunas end-to-end (ex: "A:D").
        /// </summary>
        private static List<Dictionary<string, object>> ReadColumnsEndToEnd(IXLWorksheet ws, string colRange)
        {
            var parts = colRange.Split(':');
            if (parts.Length != 2)
                throw new ArgumentException("Formato de colunas inválido. Use ex: A:D");

            var startCol = XLHelper.GetColumnNumberFromLetter(parts[0].Trim());
            var endCol = XLHelper.GetColumnNumberFromLetter(parts[1].Trim());
            var lastRow = ws.LastRowUsed()?.RowNumber() ?? 0;

            var result = new List<Dictionary<string, object>>();
            if (lastRow == 0) return result;

            for (int r = 1; r <= lastRow; r++)
            {
                for (int c = startCol; c <= endCol; c++)
                {
                    var cell = ws.Cell(r, c);
                    if (!cell.IsEmpty())
                        result.Add(CellToDict(cell));
                }
            }

            return result;
        }

        /// <summary>
        /// Lê uma linha end-to-end (da primeira até a última coluna com dados).
        /// </summary>
        private static List<Dictionary<string, object>> ReadRowEndToEnd(IXLWorksheet ws, int rowNumber)
        {
            var result = new List<Dictionary<string, object>>();
            var lastCol = ws.LastColumnUsed()?.ColumnNumber() ?? 0;
            if (lastCol == 0) return result;

            for (int c = 1; c <= lastCol; c++)
            {
                var cell = ws.Cell(rowNumber, c);
                if (!cell.IsEmpty())
                    result.Add(CellToDict(cell));
            }

            return result;
        }

        /// <summary>
        /// Lê múltiplas linhas end-to-end (ex: "1:10").
        /// </summary>
        private static List<Dictionary<string, object>> ReadRowsEndToEnd(IXLWorksheet ws, string rowRange)
        {
            var parts = rowRange.Split(':');
            if (parts.Length != 2)
                throw new ArgumentException("Formato de linhas inválido. Use ex: 1:10");

            var startRow = int.Parse(parts[0].Trim());
            var endRow = int.Parse(parts[1].Trim());
            var lastCol = ws.LastColumnUsed()?.ColumnNumber() ?? 0;

            var result = new List<Dictionary<string, object>>();
            if (lastCol == 0) return result;

            for (int r = startRow; r <= endRow; r++)
            {
                for (int c = 1; c <= lastCol; c++)
                {
                    var cell = ws.Cell(r, c);
                    if (!cell.IsEmpty())
                        result.Add(CellToDict(cell));
                }
            }

            return result;
        }

        /// <summary>
        /// Converte uma célula em dicionário para serialização.
        /// </summary>
        private static Dictionary<string, object> CellToDict(IXLCell cell)
        {
            var dict = new Dictionary<string, object>
            {
                { "row", cell.Address.RowNumber },
                { "col", cell.Address.ColumnLetter },
                { "address", cell.Address.ToString() },
                { "value", cell.IsEmpty() ? null : cell.Value.ToString() },
                { "type", cell.DataType.ToString() }
            };

            if (cell.HasFormula)
            {
                dict["formula"] = cell.FormulaA1;
            }

            return dict;
        }
    }

    /// <summary>
    /// Extensão para obter valores de dicionário com fallback null.
    /// </summary>
    public static class DictionaryExtensions
    {
        public static string GetValueOrDefault(this Dictionary<string, string> dict, string key)
        {
            return dict.ContainsKey(key) ? dict[key] : null;
        }
    }
}
