using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace rpmaster_excel.Commands
{
    /// <summary>
    /// Comando para executar macros VBA via COM Interop.
    /// Requer Microsoft Excel instalado na máquina.
    /// </summary>
    public static class MacroCommand
    {
        public static CommandResult Execute(Dictionary<string, string> args)
        {
            var file = args.GetValueOrDefault("--file");
            var macroName = args.GetValueOrDefault("--name");
            var arg1 = args.GetValueOrDefault("--arg1");
            var arg2 = args.GetValueOrDefault("--arg2");
            var arg3 = args.GetValueOrDefault("--arg3");

            if (string.IsNullOrEmpty(file))
                return CommandResult.Error("macro", "O parâmetro --file é obrigatório.");

            if (string.IsNullOrEmpty(macroName))
                return CommandResult.Error("macro", "O parâmetro --name é obrigatório.");

            if (!System.IO.File.Exists(file))
                return CommandResult.Error("macro", $"Arquivo não encontrado: {file}");

            // Resolve para caminho absoluto (COM Interop exige)
            file = System.IO.Path.GetFullPath(file);

            dynamic excelApp = null;
            dynamic workbook = null;

            try
            {
                // Cria instância do Excel via COM
                var excelType = Type.GetTypeFromProgID("Excel.Application");
                if (excelType == null)
                    return CommandResult.Error("macro", "Microsoft Excel não está instalado nesta máquina. COM Interop indisponível.");

                excelApp = Activator.CreateInstance(excelType);
                excelApp.Visible = false;
                excelApp.DisplayAlerts = false;

                // Abre o workbook (xlsm, xlsb, xls suportam macros)
                workbook = excelApp.Workbooks.Open(file);

                // Monta argumentos para a macro
                var macroArgs = new List<object>();
                if (!string.IsNullOrEmpty(arg1)) macroArgs.Add(arg1);
                if (!string.IsNullOrEmpty(arg2)) macroArgs.Add(arg2);
                if (!string.IsNullOrEmpty(arg3)) macroArgs.Add(arg3);

                // Executa a macro
                object result;
                switch (macroArgs.Count)
                {
                    case 0:
                        result = excelApp.Run(macroName);
                        break;
                    case 1:
                        result = excelApp.Run(macroName, macroArgs[0]);
                        break;
                    case 2:
                        result = excelApp.Run(macroName, macroArgs[0], macroArgs[1]);
                        break;
                    case 3:
                        result = excelApp.Run(macroName, macroArgs[0], macroArgs[1], macroArgs[2]);
                        break;
                    default:
                        result = excelApp.Run(macroName);
                        break;
                }

                // Salva e fecha
                workbook.Save();
                workbook.Close(false);
                excelApp.Quit();

                var data = new Dictionary<string, object>
                {
                    { "macro", macroName },
                    { "file", file },
                    { "result", result?.ToString() },
                    { "arguments_count", macroArgs.Count }
                };

                return CommandResult.Ok("macro", data, $"Macro '{macroName}' executada com sucesso.");
            }
            catch (COMException ex)
            {
                return CommandResult.Error("macro", $"Erro COM ao executar macro: {ex.Message}");
            }
            catch (Exception ex)
            {
                return CommandResult.Error("macro", $"Erro ao executar macro: {ex.Message}");
            }
            finally
            {
                // Limpa recursos COM
                if (workbook != null)
                {
                    try { workbook.Close(false); } catch { }
                    Marshal.ReleaseComObject(workbook);
                }
                if (excelApp != null)
                {
                    try { excelApp.Quit(); } catch { }
                    Marshal.ReleaseComObject(excelApp);
                }
            }
        }
    }
}
