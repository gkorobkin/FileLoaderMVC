using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using FileHelpers;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace FileLoaderMVC
{
    public class BatchPreTradeProcessor : IBatchProcessor
    {
        public void Process(string file)
        {
            if (string.IsNullOrWhiteSpace(file))
                throw new ArgumentNullException(nameof(file));
            if (!File.Exists(file))
                throw new ArgumentException($"File not exists: {file}");

            switch (Path.GetExtension(file).ToLowerInvariant())
            {
                case ".xlsx":
                case ".xlsm":
                case ".xls":
                    ProcessExcelBatch(file);
                    break;
                case ".csv":
                case ".txt":
                    ProcessCSVBatch(file);
                    break;
                default:
                    var msg = $"Unknow file to batch process: {Path.GetFileName(file)}";
                    Trace.TraceError(msg);
                    throw new ApplicationException(msg);
            }
        }

        private void ProcessCSVBatch(string file)
        {
            var engine = new FileHelperEngine<BatchPreTradeRecord>();

            var records = engine.ReadFile(file);
            if (!records.Any())
            {
                Trace.TraceWarning($"No data rows were read from the file: {file}");
                return;
            }

            foreach (var rec in records)
            {
                ProcessPreTradeRecord(rec);
            }
        }

        private void ProcessExcelBatch(string path)
        {
            var records = ReadRecords(path).ToList();
            if (!records.Any())
            {
                Trace.TraceWarning($"No data rows were read from the file: {path}");
                return;
            }

            foreach (var rec in records)
            {
                ProcessPreTradeRecord(rec);
            }
        }

        private IEnumerable<BatchPreTradeRecord> ReadRecords(string path)
        {
            switch (Path.GetExtension(path)?.ToLowerInvariant())
            {
                case ".xlsx":
                case ".xlsm":
                    return GetRecordsXlsx(path);
                case ".xls":
                    return GetRecordsXls(path);
                default:
                    throw new ApplicationException($"Non Excel file: {path}");
            }
        }

        private static IEnumerable<BatchPreTradeRecord> GetRecordsXlsx(string path, int headerRows = 1)
        {
            var workBook = InitWorkBookXlsx(path);
            var sheet = workBook.GetSheetAt(0);
            var rows = sheet.GetRowEnumerator();

            //  skiping header row(s)
            for (var i = 0; i < headerRows; ++i)
            {
                if (!rows.MoveNext()) yield break;
            }

            //  fill data values
            while (rows.MoveNext())
                yield return GetRecord((IRow)rows.Current);
        }

        private static BatchPreTradeRecord GetRecord(IRow row)
        {
            var rec = new BatchPreTradeRecord
            {
                BlockId = (int) row.GetCell(0).NumericCellValue,
                Amount = (decimal) row.GetCell(5).NumericCellValue
            };

            var childIdCell = row.GetCell(1);
            if (childIdCell?.CellType == CellType.Numeric)
                rec.ChildId = (int) childIdCell.NumericCellValue;

            var ccy1Cell = row.GetCell(2);
            if (ccy1Cell?.CellType == CellType.String)
                rec.Ccy1 = ccy1Cell.StringCellValue;

            var ccy2Cell = row.GetCell(3);
            if (ccy2Cell?.CellType == CellType.String)
                rec.Ccy2 = ccy2Cell.StringCellValue;

            var buySellCell = row.GetCell(4);
            if (buySellCell?.CellType == CellType.String)
                rec.BuySell = buySellCell.StringCellValue;

            var tradeDateCell = row.GetCell(6);
            if (tradeDateCell?.CellType == CellType.Numeric &&
                DateUtil.IsCellDateFormatted(tradeDateCell))
            {
                rec.TradeDate = tradeDateCell.DateCellValue;
            }

            var valueDateCell = row.GetCell(7);
            if (valueDateCell?.CellType == CellType.Numeric &&
                DateUtil.IsCellDateFormatted(valueDateCell))
            {
                rec.ValueDate = valueDateCell.DateCellValue;
            }
            return rec;
        }

        private static IEnumerable<BatchPreTradeRecord> GetRecordsXls(string path, int headerRows = 0)
        {
            HSSFWorkbook workBook = InitWorkBookXls(path);
            var sheet = workBook.GetSheetAt(0);
            var rows = sheet.GetRowEnumerator();

            //  skiping header row(s)
            for (var i = 0; i < headerRows; ++i)
            {
                if (!rows.MoveNext()) yield break;
            }

            //  fill data values
            while (rows.MoveNext())
                yield return GetRecord((IRow)rows.Current);
        }

        private static XSSFWorkbook InitWorkBookXlsx(string path)
        {
            using (var file = new FileStream(path, FileMode.Open, FileAccess.Read))
            {
                return new XSSFWorkbook(file);
            }
        }

        private static HSSFWorkbook InitWorkBookXls(string path)
        {
            using (var file = new FileStream(path, FileMode.Open, FileAccess.Read))
            {
                return new HSSFWorkbook(file);
            }
        }

        private static void ProcessPreTradeRecord(BatchPreTradeRecord rec)
        {
            Trace.WriteLine(rec);
        }
    }
}