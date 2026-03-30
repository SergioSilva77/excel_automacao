using System;
using ClosedXML.Excel;

namespace rpmaster_excel
{
    /// <summary>
    /// Motor principal para abrir, manipular e salvar workbooks Excel.
    /// </summary>
    public class ExcelEngine : IDisposable
    {
        private XLWorkbook _workbook;
        private string _filePath;

        public XLWorkbook Workbook => _workbook;
        public string FilePath => _filePath;

        /// <summary>
        /// Abre um workbook a partir do caminho informado.
        /// </summary>
        public void Open(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath))
                throw new ArgumentException("O caminho do arquivo é obrigatório.");

            if (!System.IO.File.Exists(filePath))
                throw new System.IO.FileNotFoundException($"Arquivo não encontrado: {filePath}");

            _filePath = filePath;
            _workbook = new XLWorkbook(filePath);
        }

        /// <summary>
        /// Cria um novo workbook vazio (para escrita em arquivo novo).
        /// </summary>
        public void Create(string filePath)
        {
            _filePath = filePath;
            _workbook = new XLWorkbook();
        }

        /// <summary>
        /// Obtém uma worksheet pelo nome. Se não informado, retorna a primeira.
        /// </summary>
        public IXLWorksheet GetWorksheet(string sheetName = null)
        {
            if (_workbook == null)
                throw new InvalidOperationException("Nenhum workbook aberto.");

            if (string.IsNullOrWhiteSpace(sheetName))
                return _workbook.Worksheets.Worksheet(1);

            if (!_workbook.Worksheets.TryGetWorksheet(sheetName, out var ws))
                throw new ArgumentException($"Planilha '{sheetName}' não encontrada.");

            return ws;
        }

        /// <summary>
        /// Salva o workbook no caminho original ou em novo caminho.
        /// </summary>
        public void Save(string outputPath = null)
        {
            if (_workbook == null)
                throw new InvalidOperationException("Nenhum workbook aberto.");

            _workbook.SaveAs(outputPath ?? _filePath);
        }

        public void Dispose()
        {
            _workbook?.Dispose();
        }
    }
}
