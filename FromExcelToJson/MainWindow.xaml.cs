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

        /// <summary>
        /// The row xml element container
        /// </summary>
        public const string FLD_CollectionElement = "CollectionElement";

        /// <summary>
        /// Excel file name
        /// </summary>
        public const string FLD_ExcelFileName = "ExcelFileName";

        /// <summary>
        /// Indicates if the first row has field names
        /// </summary>
        public const string FLD_FirstRowHasFieldNames = "FirstRowHasFieldNames";

        /// <summary>
        /// Result text
        /// </summary>
        public const string FLD_ResultText = "ResultText";

        /// <summary>
        /// The row xml element container
        /// </summary>
        public static readonly DependencyProperty CollectionElementProperty = DependencyProperty.Register(
            FLD_CollectionElement, typeof(string), typeof(MainWindow), new FrameworkPropertyMetadata("Cities", FrameworkPropertyMetadataOptions.BindsTwoWayByDefault));

        /// <summary>
        /// Excel file name
        /// </summary>
        public static readonly DependencyProperty ExcelFileNameProperty = DependencyProperty.Register(
            FLD_ExcelFileName, typeof(string), typeof(MainWindow), new FrameworkPropertyMetadata("", FrameworkPropertyMetadataOptions.BindsTwoWayByDefault));

        /// <summary>
        /// Indicates if the first row has field names
        /// </summary>
        public static readonly DependencyProperty FirstRowHasFieldNamesProperty = DependencyProperty.Register(
            FLD_FirstRowHasFieldNames, typeof(bool), typeof(MainWindow), new FrameworkPropertyMetadata(true, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault));

        /// <summary>
        /// Result text
        /// </summary>
        public static readonly DependencyProperty ResultTextProperty = DependencyProperty.Register(
            FLD_ResultText, typeof(string), typeof(MainWindow), new FrameworkPropertyMetadata("", FrameworkPropertyMetadataOptions.BindsTwoWayByDefault));

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
        /// The row xml element container
        /// </summary>
        public string CollectionElement
        {
            get
            {
                return (string)this.GetValue(CollectionElementProperty);
            }
            set
            {
                this.SetValue(CollectionElementProperty, value);
            }
        }

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

        /// <summary>
        /// Indicates if the first row has field names
        /// </summary>
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

        /// <summary>
        /// Result text
        /// </summary>
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

        #endregion Public Properties

        #region Public Methods

        [SecurityPermissionAttribute(SecurityAction.Demand, Flags = SecurityPermissionFlag.UnmanagedCode)]
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

        /// <summary>
        /// Gets the excel cell value as a string
        /// </summary>
        /// <param name="wks">The WKS.</param>
        /// <param name="row">The row.</param>
        /// <param name="col">The col.</param>
        /// <returns></returns>
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
                FileInfo infile = new FileInfo(ExcelFileName);
                string outputFile = ExcelFileName.Replace(infile.Extension, ".json");
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

                        GenerateJsonFile(outputFile, ws, start, end, fieldNames, firstRow);
                        if (MessageBox.Show("Do You want to open output file?", "Open the file", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
                        {
                            Process.Start(outputFile);
                        }
                    }
                    else
                    {
                        MessageBox.Show("Looks like there are no worksheets!");
                    }
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

        private void GenerateJsonFile(string outputFile, ExcelWorksheet ws,
            ExcelCellAddress start, ExcelCellAddress end, Dictionary<int, string> fieldNames, int firstRow)
        {
            var iiii = 0;
            for (int jsonStartRow = 4501, jsonendRow = jsonStartRow + 1499; jsonStartRow <= end.Row; jsonStartRow += 1500, jsonendRow += 1500)
            {
                StringBuilder sb = new StringBuilder();
                StringWriter sw = new StringWriter(sb);
                JsonWriter jsonWriter = null;
                jsonWriter = new JsonTextWriter(sw);

                //Use indentation for readability.
                jsonWriter.Formatting = Newtonsoft.Json.Formatting.Indented;


                int count = 0;
                int countDoEvents = 0;
                // jsonWriter.WriteStartObject();
                //jsonWriter.WritePropertyName(CollectionElement);
                jsonWriter.WriteStartArray();
                for (int row = jsonStartRow; row <= jsonendRow && row <= end.Row; row++)
                {
                    count++;
                    countDoEvents++;
                    if (countDoEvents >= 10)
                    {
                        ResultText = "Reading record " + count;
                        countDoEvents = 0;
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


                ResultText = "Ended reading writing file " + outputFile;
                countDoEvents = 0;

                sw.Close();
                File.WriteAllText(outputFile, sb.ToString());
                //ResultText = File.ReadAllText(outputFile);
                CallAPI(sb.ToString());
            }

            iiii++;


            //return (count);
        }

        private bool CallAPI(string json)
        {
            HttpClient client = new HttpClient();
            client.BaseAddress = new Uri("http://api-debug.fieldassist.in/");

            client.DefaultRequestHeaders.Accept.Add(
               new MediaTypeWithQualityHeaderValue("application/json"));

            // if using xml
            // client.DefaultRequestHeaders.Accept.Add(
            //   new MediaTypeWithQualityHeaderValue("application/xml"));

            var credentials = Encoding.ASCII.GetBytes("JockeyIndia-Demo:TNU3RuVatbCx71gwMHxJlz0kP1sk7zttFxkX5dlrQA2p2");
            client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Basic", Convert.ToBase64String(credentials));

            var validJson = JsonConvert.DeserializeObject<List<SkuNorms>>(json);

            var response = client.PostAsJsonAsync("api/V3/SkuSalesData/UploadSkuNorms", validJson).Result;

            // if using xml
            // var response = client.PostAsXmlAsync("api/products/Create", product).Result;

            if (response.IsSuccessStatusCode)
            {
                // product added
                return true;
            }
            else
            {
                // call function to log error if http status is not 200

                return false;
            }

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

    public class SkuNorms
    {
        public string ProductERPId { get; set; }
        public string RetailerCode { get; set; }
        public int ProductType { set; get; }
    }
}