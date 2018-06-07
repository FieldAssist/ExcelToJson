using FromExcelToJson.Helpers;
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

        public const string FLD_ResultText = "ResultText";
        public const string FLD_BaseURL = "BaseURL";
        public const string FLD_ApiURL = "ApiURL";
        public const string FLD_Credentials = "Credentials";
        public const string FLD_IsGet = "IsGet";
        public const string FLD_IsPost = "IsPost";
        public const string FLD_SplitInterval = "SplitInterval";
        public const string FLD_OnlyToJson = "OnlyToJson";
        public const string FLD_TwoColumnGroupJson = "TwoColumnGroupJson";



        public static readonly DependencyProperty ExcelFileNameProperty = DependencyProperty.Register(
            FLD_ExcelFileName, typeof(string), typeof(MainWindow), new FrameworkPropertyMetadata("", FrameworkPropertyMetadataOptions.BindsTwoWayByDefault));
        public static readonly DependencyProperty BaseURLProperty = DependencyProperty.Register(
            FLD_BaseURL, typeof(string), typeof(MainWindow), new FrameworkPropertyMetadata("", FrameworkPropertyMetadataOptions.BindsTwoWayByDefault));
        public static readonly DependencyProperty ApiURLProperty = DependencyProperty.Register(
            FLD_ApiURL, typeof(string), typeof(MainWindow), new FrameworkPropertyMetadata("", FrameworkPropertyMetadataOptions.BindsTwoWayByDefault));
        public static readonly DependencyProperty ResultTextProperty = DependencyProperty.Register(
            FLD_ResultText, typeof(string), typeof(MainWindow), new FrameworkPropertyMetadata("", FrameworkPropertyMetadataOptions.BindsTwoWayByDefault));
        public static readonly DependencyProperty CredentialsProperty = DependencyProperty.Register(
                    FLD_Credentials, typeof(string), typeof(MainWindow), new FrameworkPropertyMetadata("", FrameworkPropertyMetadataOptions.BindsTwoWayByDefault));
        public static readonly DependencyProperty IsGetProperty = DependencyProperty.Register(
                    FLD_IsGet, typeof(bool), typeof(MainWindow), new FrameworkPropertyMetadata(false, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault));
        public static readonly DependencyProperty IsPostProperty = DependencyProperty.Register(
                            FLD_IsPost, typeof(bool), typeof(MainWindow), new FrameworkPropertyMetadata(false, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault));
        public static readonly DependencyProperty SplitIntervalProperty = DependencyProperty.Register(
                            FLD_SplitInterval, typeof(int), typeof(MainWindow), new FrameworkPropertyMetadata(1500, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault));
        public static readonly DependencyProperty OnlyToJsonProperty = DependencyProperty.Register(
                            FLD_OnlyToJson, typeof(bool), typeof(MainWindow), new FrameworkPropertyMetadata(false, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault));
        public static readonly DependencyProperty TwoColumnGroupJsonProperty = DependencyProperty.Register(
                    FLD_TwoColumnGroupJson, typeof(bool), typeof(MainWindow), new FrameworkPropertyMetadata(false, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault));
        #endregion Public Fields

        #region Public Constructors

        public MainWindow()
        {
            InitializeComponent();
            this.DataContext = this;
            this.Credentials = "<USER>:<Pass>";
            this.BaseURL = "http://<subdomain>.fieldassist.in";
            this.SplitInterval = 1500;
            this.OnlyToJson = true;
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
        public int SplitInterval
        {
            get
            {
                return (int)this.GetValue(SplitIntervalProperty);
            }
            set
            {
                this.SetValue(SplitIntervalProperty, value);
            }
        }
        public bool TwoColumnGroupJson
        {
            get
            {
                return (bool)this.GetValue(TwoColumnGroupJsonProperty);
            }
            set
            {
                this.SetValue(TwoColumnGroupJsonProperty, value);
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

        public bool OnlyToJson
        {
            get
            {
                return (bool)this.GetValue(OnlyToJsonProperty);
            }
            set
            {
                this.SetValue(OnlyToJsonProperty, value);
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


        private void GenerateJson_Click(object sender, RoutedEventArgs e)
        {
            Cursor c = this.Cursor;
            try
            {
                this.Cursor = Cursors.Wait;
                DoEvents();
                if (IsPost || OnlyToJson)
                {
                    FileInfo infile = new FileInfo(ExcelFileName);
                    using (ExcelPackage exp = new ExcelPackage(infile))
                    {
                        if (exp.Workbook.Worksheets.Count > 0)
                        {
                            ExcelWorksheet ws = exp.Workbook.Worksheets.First();
                            var count = PostCallJson(ExcelFileName, ws);
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
                    WriteJsonToFile(outputFilejson, response.Content.ReadAsStringAsync().Result);
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
        private enum Action
        {
            GetCall,
            PostCall,
            OnlySave
        }

        private void LogToConsole(string str)
        {
            ResultText = $"{str}\n{ResultText}";
            DoEvents();
        }
        private long PostCallJson(string outputFile, ExcelWorksheet ws)
        {
            FileInfo infile = new FileInfo(outputFile);
            FAJsonWriter jsonWriter = TwoColumnGroupJson ? (FAJsonWriter)new TwoColumnGroupedJsonWriter(ws) : new FlatJsonWriter(ws);
            var action = OnlyToJson ? Action.OnlySave : Action.PostCall;

            var endRow = jsonWriter.EndRow;
            for (int jsonStartRow = 2, jsonendRow = jsonStartRow + SplitInterval - 1; jsonStartRow <= endRow; jsonStartRow += SplitInterval, jsonendRow += SplitInterval)
            {
                var jsonOutput = jsonWriter.CreateJson(jsonStartRow, jsonendRow);
                string outputFilejson = outputFile.Replace(infile.Extension, $"_{jsonStartRow}-{(endRow > jsonendRow ? jsonendRow : endRow)}.json");
                string responseFilejson = outputFile.Replace(infile.Extension, $"_Response_{jsonStartRow}-{(endRow > jsonendRow ? jsonendRow : endRow)}.json");
                switch (action)
                {
                    case Action.PostCall:
                        WriteJsonToFile(outputFilejson, jsonOutput);
                        var response = CallAPI(jsonOutput, BaseURL, ApiURL, Credentials);//Enter the details to call the Api
                        WriteJsonToFile(responseFilejson, response.Content.ReadAsStringAsync().Result);
                        break;
                    case Action.OnlySave:
                        WriteJsonToFile(outputFilejson, jsonOutput);
                        break;
                }
                LogToConsole($"Done upto Row {jsonendRow}");
            }
            return (endRow - 1);
        }

        private void WriteJsonToFile(string path, string Value)
        {
            File.WriteAllText(path, Value);
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
                    //var validJson = JsonConvert.DeserializeObject<List<Dictionary<string, string>>>(body);
                    var request = new HttpRequestMessage(HttpMethod.Post, api);
                    request.Content = new StringContent(body, Encoding.UTF8, "application/json");
                    try
                    {
                        var response = client.SendAsync(request).Result;
                        //var response = client.PostAsJsonAsync(Api, validJson).Result;
                        return response;
                    }
                    catch
                    {
                        return new HttpResponseMessage(System.Net.HttpStatusCode.InternalServerError);
                    }
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