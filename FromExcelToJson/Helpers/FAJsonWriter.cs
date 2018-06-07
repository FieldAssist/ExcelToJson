using Newtonsoft.Json;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace FromExcelToJson.Helpers
{
    public abstract class FAJsonWriter
    {
        private readonly ExcelWorksheet ws;
        protected readonly Action<string> log;
        protected readonly ExcelCellAddress start;
        protected readonly ExcelCellAddress end;
        protected readonly Dictionary<int, string> fieldNames;
        protected readonly int firstRow;
        public int Count { get; protected set; }
        public abstract string CreateJson(int jsonStartRow, int jsonendRow);
        public int EndRow => end.Row;
        public FAJsonWriter(ExcelWorksheet ws, Action<string> logger)
        {
            this.ws = ws;
            this.log = logger;
            start = ws.Dimension.Start;
            end = ws.Dimension.End;


            fieldNames = new Dictionary<int, string>();
            firstRow = start.Row;
            for (int x = start.Column; x <= end.Column; x++)
            {
                fieldNames.Add(x, Regex.Replace(GetCellStringValue(start.Row, x), @"\s+", ""));
            }
            firstRow++;
        }
        protected string GetCellStringValue(int row, int col)
        {
            object cVal = ws.Cells[row, col].Value;
            if (cVal == null)
                return null;
            else
                return cVal.ToString();
        }
    }


    public class FlatJsonWriter : FAJsonWriter
    {
        public string ResultText { get; private set; }

        public FlatJsonWriter(ExcelWorksheet ws, Action<string> logger) : base(ws, logger)
        {
            Count = 0;
            ResultText = "";
        }

        public override string CreateJson(int jsonStartRow, int jsonendRow)
        {
            StringBuilder sb = new StringBuilder();
            StringWriter sw = new StringWriter(sb);
            JsonWriter jsonWriter = null;
            jsonWriter = new JsonTextWriter(sw)
            {
                //Use indentation for readability.
                Formatting = Newtonsoft.Json.Formatting.Indented
            };

            jsonWriter.WriteStartArray();
            for (int row = jsonStartRow; row <= jsonendRow && row <= end.Row; row++)
            {
                Count++;
                if (Count % 1000 == 0 || row >= end.Row)
                {
                    log($"Done upto Row {Count}\n{ResultText}");
                }
                jsonWriter.WriteStartObject();
                for (int col = start.Column; col <= end.Column; col++)
                {
                    jsonWriter.WritePropertyName(fieldNames[col]);
                    jsonWriter.WriteValue(GetCellStringValue(row, col));
                }
                jsonWriter.WriteEndObject();

            }
            jsonWriter.WriteEndArray();
            //jsonWriter.WriteEndObject();
            jsonWriter.Close();
            sw.Close();
            return sb.ToString();
        }
    }

    public class TwoColumnGroupedJsonWriter : FAJsonWriter
    {
        private readonly ExcelWorksheet ws;

        public TwoColumnGroupedJsonWriter(ExcelWorksheet ws, Action<string> logger) : base(ws, logger)
        {
            this.ws = ws;
        }

        public override string CreateJson(int jsonStartRow, int jsonendRow)
        {
            throw new NotImplementedException();
        }
    }
}
