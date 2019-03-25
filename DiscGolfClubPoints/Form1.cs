using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Threading;

namespace DiscGolfClubPoints
{
    public partial class Form1 : Form
    {
        static string[] Scopes = { SheetsService.Scope.SpreadsheetsReadonly };
        static string ApplicationName = "Disc Golf Club Points";
        static SheetsService service;
        static Dictionary<string, int> origNames;

        public Form1()
        {
            InitializeComponent();
            origNames = new Dictionary<string, int>();
            connectToSheets();
            fillNameCombo();
        }

        private void connectToSheets()
        {
            UserCredential credential;

            using (var stream = new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
            {
                string credPath = "token.json";
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                Console.WriteLine("Credential file saved to: " + credPath);
            }

            service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });
        }

        private void fillNameCombo()
        {
            string spreadSheetId = "1a68GxhJQDjvzBhBOLpJbch3Jj33bZlYE9FfHrboFB_E";
            string range = "Sheet1!A2:B";
            string curName;
            int curPoints;
            SpreadsheetsResource.ValuesResource.GetRequest request =
                service.Spreadsheets.Values.Get(spreadSheetId, range);

            ValueRange response = request.Execute();
            IList<IList<Object>> values = response.Values;
            if (values != null && values.Count > 0)
            {
                foreach (var row in values)
                {
                    try
                    {
                        curName = row[0].ToString();
                        curPoints = int.Parse(row[1].ToString());
                        origNames.Add(curName, curPoints);
                        nameCombo.Items.Add(curName);
                    } catch
                    {
                        //TODO error message something wasnt formated correctly.
                    }
                }
            }
            else
            {
                //TODO error message no data found
            }
        }

        private void enterButton_Click(object sender, EventArgs e)
        {
            if (nameCombo.Text.Length < 0)
            {
                //TODO error message and change to = instead of <
                return;
            }

            addName("Ian");
        }

        private void addName(string name)
        {
            
        }
    }
}
