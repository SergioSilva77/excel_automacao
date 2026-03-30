using System;
using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel;

namespace rpmaster_excel.Commands
{
    /// <summary>
    /// Comando para listar planilhas, tabelas e informações do workbook.
    /// </summary>
    public static class InfoCommand
    {
        public static CommandResult Execute(Dictionary<string, string> args)
        {
            var file = args.GetValueOrDefault("--file");
            var listSheets = args.ContainsKey("--list-sheets");
            var listTables = args.ContainsKey("--list-tables");
            var summary = args.ContainsKey("--summary");

            using (var engine = new ExcelEngine())
            {
                engine.Open(file);
                var wb = engine.Workbook;

                if (listSheets)
                {
                    return ListSheets(wb);
                }
                else if (listTables)
                {
                    var sheet = args.GetValueOrDefault("--sheet");
                    return ListTables(wb, sheet);
                }
                else if (summary)
                {
                    return GetSummary(wb, file);
                }
                else
                {
                    return CommandResult.Error("info", "Especifique --list-sheets, --list-tables ou --summary.");
                }
            }
        }

        private static CommandResult ListSheets(XLWorkbook wb)
        {
            var sheets = new List<Dictionary<string, object>>();

            foreach (var ws in wb.Worksheets)
            {
                var lastRow = ws.LastRowUsed()?.RowNumber() ?? 0;
                var lastCol = ws.LastColumnUsed()?.ColumnNumber() ?? 0;

                sheets.Add(new Dictionary<string, object>
                {
                    { "name", ws.Name },
                    { "position", ws.Position },
                    { "last_row", lastRow },
                    { "last_column", lastCol },
                    { "visibility", ws.Visibility.ToString() },
                    { "tables_count", ws.Tables.Count() }
                });
            }

            return CommandResult.Ok("info", sheets, $"{sheets.Count} planilha(s) encontrada(s).");
        }

        private static CommandResult ListTables(XLWorkbook wb, string sheetName)
        {
            var tables = new List<Dictionary<string, object>>();

            IEnumerable<IXLWorksheet> worksheets;

            if (!string.IsNullOrEmpty(sheetName))
            {
                var ws = wb.Worksheets.TryGetWorksheet(sheetName, out var found) ? found : null;
                if (ws == null)
                    return CommandResult.Error("info", $"Planilha '{sheetName}' não encontrada.");
                worksheets = new[] { ws };
            }
            else
            {
                worksheets = wb.Worksheets;
            }

            foreach (var ws in worksheets)
            {
                foreach (var table in ws.Tables)
                {
                    tables.Add(new Dictionary<string, object>
                    {
                        { "name", table.Name },
                        { "sheet", ws.Name },
                        { "range", table.RangeAddress.ToString() },
                        { "rows", table.RowCount() },
                        { "columns", table.ColumnCount() },
                        { "show_totals", table.ShowTotalsRow },
                        { "show_autofilter", table.ShowAutoFilter }
                    });
                }
            }

            return CommandResult.Ok("info", tables, $"{tables.Count} tabela(s) encontrada(s).");
        }

        private static CommandResult GetSummary(XLWorkbook wb, string filePath)
        {
            var fileInfo = new System.IO.FileInfo(filePath);
            var totalTables = 0;

            foreach (var ws in wb.Worksheets)
            {
                totalTables += ws.Tables.Count();
            }

            var data = new Dictionary<string, object>
            {
                { "file", filePath },
                { "file_size_kb", Math.Round(fileInfo.Length / 1024.0, 2) },
                { "sheets_count", wb.Worksheets.Count },
                { "tables_count", totalTables },
                { "sheet_names", wb.Worksheets.Select(w => w.Name).ToList() },
                { "created", fileInfo.CreationTime.ToString("yyyy-MM-dd HH:mm:ss") },
                { "modified", fileInfo.LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss") }
            };

            return CommandResult.Ok("info", data, "Resumo do arquivo.");
        }
    }
}
