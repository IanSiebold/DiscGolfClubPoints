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
using System.Diagnostics;

namespace DiscGolfClubPoints
{
    public partial class Form1 : Form
    {
        static string[] Scopes = { SheetsService.Scope.Spreadsheets };
        static string ApplicationName = "Disc Golf Club Points";
        static string spreadSheetId = "1a68GxhJQDjvzBhBOLpJbch3Jj33bZlYE9FfHrboFB_E";
        static SheetsService service;
        static Dictionary<string, int> origNames;
        static Dictionary<string, int> updatedNames;
        static Boolean needsUpdate;

        public Form1()
        {
            InitializeComponent();
            needsUpdate = false;
            origNames = new Dictionary<string, int>();
            updatedNames = new Dictionary<string, int>();
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
            //TODO eventually change this from hard coded to a find method on the first row
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
                        if (!origNames.Keys.Contains(curName))
                        {
                            origNames.Add(curName, curPoints);
                        }
                        else
                        {
                            origNames[curName] += curPoints;
                        }
                        nameCombo.Items.Add(curName);
                    }
                    catch
                    {
                        //TODO error message something wasnt formated correctly.
                    }
                }
                updatedNames = origNames;
            }
            else
            {
                //TODO error message no data found
            }
        }

        private void enterButton_Click(object sender, EventArgs e)
        {
            Stopwatch sw = new Stopwatch();

            needsUpdate = true;
            string enteredName = nameCombo.Text;
            if (enteredName.Length == 0)
            {
                return;
            }
            else if (updatedNames.Keys.Contains(enteredName))
            {
                updatedNames[enteredName] += 1;
            }
            else
            {
                updatedNames.Add(enteredName, 1);
                nameCombo.Items.Add(enteredName);
            }
            nameCombo.Text = nameCombo.Text + " has been logged.";
            toggleAllButtons();
            sw.Start();
            for (int i = 0; ; i++)
            {
                if (i % 10000 == 0)
                {
                    sw.Stop();
                    if (sw.ElapsedMilliseconds > 3000)
                    {
                        break;
                    } else
                    {
                        sw.Start();
                    }
                }
            }
            nameCombo.Text = "";
            toggleAllButtons();
            
        }

        private void updateButton_Click(object sender, EventArgs e)
        {
            updateSheets();
        }

        private void updateSheets()
        {
            toggleAllButtons();
            if (needsUpdate)
            {
                string range1 = "Sheet1!A2:A";
                string range2 = "Sheet1!B2:B";

                clearColumn(range1);
                clearColumn(range2);

                ValueRange valueRange = new ValueRange();
                valueRange.MajorDimension = "COLUMNS";

                var oblist = updatedNames.Keys.Cast<object>().ToList();
                valueRange.Values = new List<IList<object>> { oblist };

                SpreadsheetsResource.ValuesResource.UpdateRequest update = service.Spreadsheets.Values.Update(valueRange, spreadSheetId, range1);
                update.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
                update.Execute();

                oblist = updatedNames.Values.Cast<object>().ToList();
                valueRange.Values = new List<IList<object>> { oblist };

                update = service.Spreadsheets.Values.Update(valueRange, spreadSheetId, range2);
                update.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
                update.Execute();
            }
            toggleAllButtons();
        }

        private void clearColumn(string range)
        {
            ClearValuesRequest clearRequest = new ClearValuesRequest();
            SpreadsheetsResource.ValuesResource.ClearRequest req = service.Spreadsheets.Values.Clear(clearRequest, spreadSheetId, range);
            req.Execute();
        }

        private void toggleAllButtons()
        {
            if (enterButton.Enabled)
            {
                enterButton.Enabled = false;
            }
            else
            {
                enterButton.Enabled = true;
            }

            if (updateButton.Enabled)
            {
                updateButton.Enabled = false;
            }
            else
            {
                updateButton.Enabled = true;
            }

            if (nameCombo.Enabled)
            {
                nameCombo.Enabled = false;
            }
            else
            {
                nameCombo.Enabled = true;
            }
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (needsUpdate)
            {
                //TODO message saying you didn't save. Ask them to save
            }
        }
    }
}
