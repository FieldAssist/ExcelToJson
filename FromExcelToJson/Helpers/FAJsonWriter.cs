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
        protected readonly ExcelCellAddress start;
        protected readonly ExcelCellAddress end;
        protected readonly Dictionary<int, string> fieldNames;
        public int Count { get; protected set; }
        public abstract string CreateJson(int jsonStartRow, int jsonendRow);
        public int EndRow => end.Row;
        public FAJsonWriter(ExcelWorksheet ws)
        {
            this.ws = ws;
            start = ws.Dimension.Start;
            end = ws.Dimension.End;
            fieldNames = new Dictionary<int, string>();
            for (int x = start.Column; x <= end.Column; x++)
            {
                fieldNames.Add(x, Regex.Replace(GetCellStringValue(start.Row, x), @"\s+", ""));
            }
        }
        public abstract bool IsValid();
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

        public FlatJsonWriter(ExcelWorksheet ws) : base(ws)
        {
            Count = 0;
            ResultText = "";
        }

        public override string CreateJson(int jsonStartRow, int jsonendRow)
        {
            if (!IsValid())
            {
                throw new Exception("Check Your Sheet... Some Issue!");
            }
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

        public override bool IsValid()
        {
            return start.Column == 1 && end.Column >= 1 && start.Row == 1 && end.Row > 1;
        }
    }

    public class TwoColumnGroupedJsonWriter : FAJsonWriter
    {
        private class DataModel
        {
            public string Col1 { get; set; }
            public string Col2 { get; set; }
        }

        private readonly ExcelWorksheet ws;

        public TwoColumnGroupedJsonWriter(ExcelWorksheet ws) : base(ws)
        {
            this.ws = ws;

        }
        public override bool IsValid()
        {
            return start.Column == 1 && end.Column == 2 && start.Row == 1 && end.Row > 1;
        }
        public override string CreateJson(int jsonStartRow, int jsonendRow)
        {
            if (!IsValid())
            {
                throw new Exception("For Two Column Group, You need to have exactly two columns");
            }

            var data = new List<DataModel>();

            for (int row = jsonStartRow; row <= jsonendRow && row <= end.Row; row++)
            {
                Count++;
                data.Add(new DataModel
                {
                    Col1 = GetCellStringValue(row, 1),
                    Col2 = GetCellStringValue(row, 2),
                });
            }

            var dictionary = data.GroupBy(d => d.Col1).ToDictionary(d => d.Key, d => d.Select(v => v.Col2));
            var arrayDictionary = dictionary.Select(d => new Dictionary<string, object> { [fieldNames[1]] = d.Key, [fieldNames[2]] = d.Value }).ToList();
            StringBuilder sb = new StringBuilder();
            StringWriter sw = new StringWriter(sb);
            var jsonWriter = new JsonTextWriter(sw)
            {
                //Use indentation for readability.
                Formatting = Formatting.Indented
            };
            jsonWriter.WriteRaw(JsonConvert.SerializeObject(arrayDictionary));
            jsonWriter.Close();
            sw.Close();
            return sb.ToString();
        }
    }
}
