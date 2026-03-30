using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Xml.Linq;

namespace rpmaster_excel
{
    /// <summary>
    /// Resultado padrão retornado por todos os comandos.
    /// </summary>
    public class CommandResult
    {
        [JsonPropertyName("success")]
        public bool Success { get; set; }

        [JsonPropertyName("command")]
        public string Command { get; set; }

        [JsonPropertyName("data")]
        public object Data { get; set; }

        [JsonPropertyName("message")]
        public string Message { get; set; }

        public static CommandResult Ok(string command, object data, string message = null)
        {
            return new CommandResult { Success = true, Command = command, Data = data, Message = message };
        }

        public static CommandResult Error(string command, string message)
        {
            return new CommandResult { Success = false, Command = command, Data = null, Message = message };
        }
    }

    /// <summary>
    /// Formata a saída em JSON ou XML.
    /// </summary>
    public static class OutputFormatter
    {
        public static string Format(CommandResult result, string format)
        {
            format = (format ?? "json").ToLowerInvariant();

            switch (format)
            {
                case "xml":
                    return ToXml(result);
                case "json":
                default:
                    return ToJson(result);
            }
        }

        private static string ToJson(CommandResult result)
        {
            var options = new JsonSerializerOptions
            {
                WriteIndented = true,
                DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
                Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping
            };
            return JsonSerializer.Serialize(result, options);
        }

        private static string ToXml(CommandResult result)
        {
            var root = new XElement("result",
                new XElement("success", result.Success),
                new XElement("command", result.Command),
                new XElement("message", result.Message ?? "")
            );

            if (result.Data != null)
            {
                // Serializa data como JSON string dentro de um CDATA para preservar estrutura complexa
                var jsonData = JsonSerializer.Serialize(result.Data, new JsonSerializerOptions
                {
                    WriteIndented = false,
                    Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping
                });

                if (result.Data is IEnumerable<Dictionary<string, object>> list)
                {
                    var dataEl = new XElement("data");
                    foreach (var row in list)
                    {
                        var rowEl = new XElement("row");
                        foreach (var kv in row)
                        {
                            rowEl.Add(new XElement(SanitizeXmlName(kv.Key), kv.Value?.ToString() ?? ""));
                        }
                        dataEl.Add(rowEl);
                    }
                    root.Add(dataEl);
                }
                else
                {
                    root.Add(new XElement("data", new XCData(jsonData)));
                }
            }

            return root.ToString();
        }

        private static string SanitizeXmlName(string name)
        {
            if (string.IsNullOrEmpty(name)) return "field";
            // XML element names can't start with a number
            var sanitized = new string(name.Select(c => char.IsLetterOrDigit(c) || c == '_' ? c : '_').ToArray());
            if (char.IsDigit(sanitized[0])) sanitized = "_" + sanitized;
            return sanitized;
        }
    }
}
