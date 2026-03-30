using System;
using System.Collections.Generic;
using ClosedXML.Excel;

namespace rpmaster_excel.Commands
{
    /// <summary>
    /// Comando para aplicar e remover filtros (AutoFilter) em planilhas.
    /// </summary>
    public static class FilterCommand
    {
        public static CommandResult Execute(Dictionary<string, string> args)
        {
            var file = args.GetValueOrDefault("--file");
            var sheet = args.GetValueOrDefault("--sheet");
            var range = args.GetValueOrDefault("--range");
            var remove = args.ContainsKey("--remove");
            var apply = args.ContainsKey("--apply");
            var output = args.GetValueOrDefault("--output");

            using (var engine = new ExcelEngine())
            {
                engine.Open(file);
                var ws = engine.GetWorksheet(sheet);

                if (remove)
                {
                    // Remove autofilter da planilha
                    if (ws.AutoFilter != null && ws.AutoFilter.IsEnabled)
                    {
                        ws.AutoFilter.Clear();
                    }

                    engine.Save(output);

                    return CommandResult.Ok("filter", new Dictionary<string, object>
                    {
                        { "action", "removed" },
                        { "sheet", ws.Name }
                    }, "Filtro removido com sucesso.");
                }
                else if (apply)
                {
                    if (string.IsNullOrEmpty(range))
                        return CommandResult.Error("filter", "O parâmetro --range é obrigatório para aplicar filtro.");

                    var rng = ws.Range(range);
                    rng.SetAutoFilter();

                    engine.Save(output);

                    return CommandResult.Ok("filter", new Dictionary<string, object>
                    {
                        { "action", "applied" },
                        { "range", range },
                        { "sheet", ws.Name }
                    }, $"AutoFilter aplicado no range {range}.");
                }
                else
                {
                    return CommandResult.Error("filter", "Especifique --apply ou --remove.");
                }
            }
        }
    }
}
