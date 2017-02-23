using Microsoft.Win32;
using Newtonsoft.Json;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Security.Permissions;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Input;
using System.Windows.Threading;

namespace FromExcelToJson
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {

        #region Public Fields


        public const string FLD_ExcelFileName = "ExcelFileName";
        public const string FLD_FirstRowHasFieldNames = "FirstRowHasFieldNames";

        public const string FLD_ResultText = "ResultText";
        public const string FLD_BaseURL = "BaseURL";
        public const string FLD_ApiURL = "ApiURL";
        public const string FLD_Credentials = "Credentials";
        public const string FLD_IsGet = "IsGet";
        public const string FLD_IsPost = "IsPost";


        public static readonly DependencyProperty ExcelFileNameProperty = DependencyProperty.Register(
            FLD_ExcelFileName, typeof(string), typeof(MainWindow), new FrameworkPropertyMetadata("", FrameworkPropertyMetadataOptions.BindsTwoWayByDefault));
        public static readonly DependencyProperty BaseURLProperty = DependencyProperty.Register(
            FLD_BaseURL, typeof(string), typeof(MainWindow), new FrameworkPropertyMetadata("", FrameworkPropertyMetadataOptions.BindsTwoWayByDefault));
        public static readonly DependencyProperty ApiURLProperty = DependencyProperty.Register(
            FLD_ApiURL, typeof(string), typeof(MainWindow), new FrameworkPropertyMetadata("", FrameworkPropertyMetadataOptions.BindsTwoWayByDefault));
        public static readonly DependencyProperty FirstRowHasFieldNamesProperty = DependencyProperty.Register(
            FLD_FirstRowHasFieldNames, typeof(bool), typeof(MainWindow), new FrameworkPropertyMetadata(true, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault));
        public static readonly DependencyProperty ResultTextProperty = DependencyProperty.Register(
            FLD_ResultText, typeof(string), typeof(MainWindow), new FrameworkPropertyMetadata("", FrameworkPropertyMetadataOptions.BindsTwoWayByDefault));
        public static readonly DependencyProperty CredentialsProperty = DependencyProperty.Register(
                    FLD_Credentials, typeof(string), typeof(MainWindow), new FrameworkPropertyMetadata("", FrameworkPropertyMetadataOptions.BindsTwoWayByDefault));
        public static readonly DependencyProperty IsGetProperty = DependencyProperty.Register(
                    FLD_IsGet, typeof(bool), typeof(MainWindow), new FrameworkPropertyMetadata(false, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault));
        public static readonly DependencyProperty IsPostProperty = DependencyProperty.Register(
                            FLD_IsPost, typeof(bool), typeof(MainWindow), new FrameworkPropertyMetadata(false, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault));

        #endregion Public Fields

        #region Public Constructors

        public MainWindow()
        {
            InitializeComponent();
            this.DataContext = this;
        }

        #endregion Public Constructors

        #region Public Properties

        /// <summary>
        /// Excel file name
        /// </summary>
        public string ExcelFileName
        {
            get
            {
                return (string)this.GetValue(ExcelFileNameProperty);
            }
            set
            {
                this.SetValue(ExcelFileNameProperty, value);
            }
        }

        public string BaseURL
        {
            get
            {
                return (string)this.GetValue(BaseURLProperty);
            }
            set
            {
                this.SetValue(BaseURLProperty, value);
            }
        }

        public string ApiURL
        {
            get
            {
                return (string)this.GetValue(ApiURLProperty);
            }
            set
            {
                this.SetValue(ApiURLProperty, value);
            }
        }
        public bool FirstRowHasFieldNames
        {
            get
            {
                return (bool)this.GetValue(FirstRowHasFieldNamesProperty);
            }
            set
            {
                this.SetValue(FirstRowHasFieldNamesProperty, value);
            }
        }

        public string Credentials
        {
            get
            {
                return (string)this.GetValue(CredentialsProperty);
            }
            set
            {
                this.SetValue(CredentialsProperty, value);
            }
        }
        public string ResultText
        {
            get
            {
                return (string)this.GetValue(ResultTextProperty);
            }
            set
            {
                this.SetValue(ResultTextProperty, value);
            }
        }
        public bool IsGet
        {
            get
            {
                return (bool)this.GetValue(IsGetProperty);
            }
            set
            {
                this.SetValue(IsPostProperty, value);
            }
        }
        public bool IsPost
        {
            get
            {
                return (bool)this.GetValue(IsPostProperty);
            }
            set
            {
                this.SetValue(IsPostProperty, value);
            }
        }
        #endregion Public Properties

        #region Public Methods

        [SecurityPermission(SecurityAction.Demand, Flags = SecurityPermissionFlag.UnmanagedCode)]
        public void DoEvents()
        {
            DispatcherFrame frame = new DispatcherFrame();
            Dispatcher.CurrentDispatcher.BeginInvoke(DispatcherPriority.Background,
                new DispatcherOperationCallback(DoExitFrame), frame);
            Dispatcher.PushFrame(frame);
        }

        public object DoExitFrame(object f)
        {
            ((DispatcherFrame)f).Continue = false;

            return null;
        }

        #endregion Public Methods

        #region Private Methods
        private static string GetCellStringValue(ExcelWorksheet wks, int row, int col)
        {
            object cVal = wks.Cells[row, col].Value;
            if (cVal == null)
                return null;
            else
                return cVal.ToString();
        }

        private void GenerateJson_Click(object sender, RoutedEventArgs e)
        {
            Cursor c = this.Cursor;
            try
            {
                this.Cursor = Cursors.Wait;
                DoEvents();
                if (IsPost)
                {
                    FileInfo infile = new FileInfo(ExcelFileName);
                    using (ExcelPackage exp = new ExcelPackage(infile))
                    {
                        if (exp.Workbook.Worksheets.Count > 0)
                        {
                            ExcelWorksheet ws = exp.Workbook.Worksheets.First();
                            var start = ws.Dimension.Start;
                            var end = ws.Dimension.End;

                            Dictionary<int, string> fieldNames = new Dictionary<int, string>();
                            int firstRow = start.Row;
                            if (FirstRowHasFieldNames)
                            {
                                for (int x = start.Column; x <= end.Column; x++)
                                {
                                    //fieldNames.Add(x, GetCellStringValue(ws, x, start.Row));
                                    fieldNames.Add(x, Regex.Replace(GetCellStringValue(ws, start.Row, x), @"\s+", ""));
                                }
                                firstRow++;
                            }
                            else
                            {
                                for (int x = start.Column; x <= end.Column; x++)
                                {
                                    //fieldNames.Add(x, string.Format("Column_{0}", x));
                                    fieldNames.Add(x, Regex.Replace(GetCellStringValue(ws, start.Row, x), @"\s+", ""));
                                }
                                firstRow++;
                            }

                            var count = GenerateJsonFile(ExcelFileName, ws, start, end, fieldNames, firstRow);
                            MessageBox.Show($"Total Entries Sent - {count}", "Job Done");

                        }
                        else
                        {
                            MessageBox.Show("Looks like there are no worksheets!");
                        }
                    }
                }
                else if (IsGet)
                {
                    var response = CallAPI("", BaseURL, ApiURL, Credentials, false);
                    FileInfo infile = new FileInfo(ExcelFileName);
                    string outputFilejson = ExcelFileName.Replace(infile.Extension, $"_Response.json");
                    File.WriteAllText(outputFilejson, response.Content.ReadAsStringAsync().Result);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                this.Cursor = c;
            }
        }

        private long GenerateJsonFile(string outputFile, ExcelWorksheet ws,
            ExcelCellAddress start, ExcelCellAddress end, Dictionary<int, string> fieldNames, int firstRow)
        {
            long count = 0;
            FileInfo infile = new FileInfo(outputFile);
            for (int jsonStartRow = firstRow, jsonendRow = jsonStartRow + 1499; jsonStartRow <= end.Row; jsonStartRow += 1500, jsonendRow += 1500)
            {
                StringBuilder sb = new StringBuilder();
                StringWriter sw = new StringWriter(sb);
                JsonWriter jsonWriter = null;
                jsonWriter = new JsonTextWriter(sw);

                //Use indentation for readability.
                jsonWriter.Formatting = Newtonsoft.Json.Formatting.Indented;


                // jsonWriter.WriteStartObject();
                //jsonWriter.WritePropertyName(CollectionElement);
                jsonWriter.WriteStartArray();
                for (int row = jsonStartRow; row <= jsonendRow && row <= end.Row; row++)
                {
                    count++;
                    if (count % 1000 == 0 || row >= end.Row)
                    {
                        ResultText = $"Done upto Row {count}\n{ResultText}";
                        DoEvents();
                    }
                    jsonWriter.WriteStartObject();
                    for (int col = start.Column; col <= end.Column; col++)
                    {
                        jsonWriter.WritePropertyName(fieldNames[col]);
                        jsonWriter.WriteValue(GetCellStringValue(ws, row, col));
                    }
                    jsonWriter.WriteEndObject();

                }
                jsonWriter.WriteEndArray();
                //jsonWriter.WriteEndObject();
                jsonWriter.Close();
                sw.Close();
                string outputFilejson = outputFile.Replace(infile.Extension, $"_{jsonStartRow}-{(end.Row > jsonendRow ? jsonendRow : end.Row)}.json");
                string responseFilejson = outputFile.Replace(infile.Extension, $"_Response_{jsonStartRow}-{(end.Row > jsonendRow ? jsonendRow : end.Row)}.json");
                File.WriteAllText(outputFilejson, sb.ToString());

                var response = CallAPI(sb.ToString(), BaseURL, ApiURL, Credentials);//Enter the details to call the Api
                File.WriteAllText(responseFilejson, response.Content.ReadAsStringAsync().Result);
                //ResultText = File.ReadAllText(outputFile);
            }
            return (count);
        }

        private HttpResponseMessage CallAPI(string body, string baseURL, string api, string credentialString, bool IsXML = false)
        {
            HttpClient client = new HttpClient();
            client.BaseAddress = new Uri(baseURL);
            var credentials = Encoding.ASCII.GetBytes(credentialString);
            client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Basic", Convert.ToBase64String(credentials));
            if (IsXML)
            {
                client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/xml"));
            }
            else
            {
                client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
            }
            if (IsGet)
            {
                var response = client.GetAsync(api).Result;
                return response;
                //return $"{baseURL}-{Api}-{credentialString}--Get";
            }
            else if (IsPost)
            {
                if (IsXML)
                {
                    //var validXML = JsonConvert.DeserializeObject<List<Dictionary<string, string>>>(body);

                    var request = new HttpRequestMessage(HttpMethod.Post, api);
                    request.Content = new StringContent(body, Encoding.UTF8, "application/xml");
                    var response = client.SendAsync(request).Result;
                    return response;
                    //var response = client.PostAsXmlAsync(Api, validXML).Result;
                    //return $"{baseURL}-{api}-{credentialString}-Post-XML";
                }
                else
                {
                    var validJson = JsonConvert.DeserializeObject<List<Dictionary<string, string>>>(body);
                    var request = new HttpRequestMessage(HttpMethod.Post, api);
                    request.Content = new StringContent(body, Encoding.UTF8, "application/xml");
                    var response = client.SendAsync(request).Result;
                    //var response = client.PostAsJsonAsync(Api, validJson).Result;
                    return response;
                    //return $"{baseURL}-{api}-{credentialString}-Post-JSON";
                }
            }
            else
                return new HttpResponseMessage(System.Net.HttpStatusCode.BadRequest) { ReasonPhrase = "No Method Specified" };
        }



        private void GetExcelFile_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Title = "Indicate the excel file containing the data";
            ofd.Filter = "Excel file (*.xlsx)|*.xlsx|All files (*.*)|*.*";
            ofd.Multiselect = false;
            bool? opened = ofd.ShowDialog(this);
            if (opened.HasValue && opened.Value)
            {
                ExcelFileName = ofd.FileName;
            }
        }

        #endregion Private Methods
    }

    //public class SkuNorms
    //{
    //    public string ProductERPId { get; set; }
    //    public string RetailerCode { get; set; }
    //    public int ProductType { set; get; }
    //}
}