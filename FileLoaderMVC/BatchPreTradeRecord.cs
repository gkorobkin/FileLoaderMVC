using System;
using System.ComponentModel.DataAnnotations;
using FileHelpers;

namespace FileLoaderMVC
{
    [DelimitedRecord(",")]
    public class BatchPreTradeRecord
    {
        [Required]
        public int BlockId { get; set; }

//        [FieldNullValue(0)]
        public int? ChildId { get; set; }

//        [FieldNullValue("")]
        public string Ccy1 { get; set; }

//        [FieldNullValue("")]
        public string Ccy2 { get; set; }

//        [FieldNullValue("")]
        public string BuySell { get; set; }

        [Required]
        public decimal Amount { get; set; }

        [FieldOptional]
        [FieldConverter(ConverterKind.Date, "yyyy-MM-dd")]
        public DateTime? TradeDate { get; set; }

        [FieldOptional]
        [FieldConverter(ConverterKind.Date, "yyyy-MM-dd")]
        public DateTime? ValueDate { get; set; }

        public override string ToString()
        {
            return $"Block: {BlockId}, child: {ChildId}, ccy1: {Ccy1}, ccy2: {Ccy2}, Buy/Sell: {BuySell}, amount: {Amount:N}, trade date: {TradeDate:d}, value date: {ValueDate:d}";
        }
    }
}
