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
using System.Text.RegularExpressions;
using System.Threading;
using System.Diagnostics;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Calendar.v3;
using Google.Apis.Calendar.v3.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;

namespace WindowsFormsApplication4
{


    public partial class Form1 : Form
    {
        static string[] Scopes = { CalendarService.Scope.Calendar };
        static string ApplicationName = "Google Calendar API .NET Quickstart";
        Form2 frm = new Form2();

        public Form1()
        {

            InitializeComponent();
            Color test = Color.FromArgb(120, 154, 166);

            menuStrip1.BackColor = test;
            this.BackColor = test;
            button1.BackColor = test;
            button2.BackColor = test;
            label1.BackColor = test;
            label2.BackColor = test;
            statusStrip1.BackColor = test;
           
        }


        //Strings & Ints get initialized here so that they can be used outside the loops
        public static string Texter;
        public static string StartTijd;
        public static string EindTijd;
        public static int Dag;
        public static string Post;
        public static string CompleteString;//For logging
        public static string CompleteString2;//For logging
        public static string googleStart;
        public static string googleEind;
        public static DateTime gStartDt;//Converts extracted time to DT Object
        public static DateTime gEindDt;
        public static int Dag2;
        public static int Maand2;


        //Convert month string to string so it can be used in the DateTime
        public static int Maand = 11;
        public static int y = DateTime.Now.Year;
        //Initialize a list to contain the filtered text
        public static List<string> ListText = new List<string>();
        public static List<string> ListPost = new List<string>();
        public static List<string> ListStart = new List<string>();
        public static List<string> ListEind = new List<string>();



        public void PdfConvertor()
        {
            OpenFileDialog BestandenZoeker = new OpenFileDialog();

            BestandenZoeker.ShowDialog();

            Process Proces = new Process();
            ProcessStartInfo PSI = new ProcessStartInfo();
            PSI.FileName = @"pdftotext.exe";
            PSI.Arguments = @" -table " + BestandenZoeker.FileName + " Planning2Txt.txt";
            Proces.StartInfo = PSI;
            Proces.Start();
            Proces.WaitForExit();
            Proces.Close();
            StreamReader Lezer;

            Lezer = new StreamReader(@"Planning2Txt.txt");


            //StreamWriter writeLog = new StreamWriter("Log.txt");

            //Loop that checks if document is finished reading or not
            while ((Texter = Lezer.ReadLine()) != null)
            {

                //Get month
                if (Texter.Contains("JANUARI")) { Maand = 1; }
                if (Texter.Contains("FEBRUARI")) { Maand = 2; }
                if (Texter.Contains("MAART")) { Maand = 3; }
                if (Texter.Contains("APRIL")) { Maand = 4; }
                if (Texter.Contains("MEI")) { Maand = 5; }
                if (Texter.Contains("JUNI")) { Maand = 6; }
                if (Texter.Contains("JULI")) { Maand = 7; }
                if (Texter.Contains("AUGUSTUS")) { Maand = 8; }
                if (Texter.Contains("SEPTEMBER")) { Maand = 9; }
                if (Texter.Contains("OKTOBER")) { Maand = 10; }
                if (Texter.Contains("NOVEMBER")) { Maand = 11; }
                if (Texter.Contains("DECEMBER")) { Maand = 12; }

                //If statement to filter the lines that contain useful info
                Regex Useful = new Regex(@"\s\w\w\s\s\s\s\d");

                if (Useful.IsMatch(Texter))
                {
                    ListText.Add(Texter);
                    //writeLog.WriteLine(Texter);
                }
            }

            Lezer.Close();
            //writeLog.Close();
            File.Delete("Planning2Txt.txt");//So in the future a new file is generated
            //Pull info from each extracted line from .txt file
            foreach (var item in ListText)
            {
                //BEGIN UUR
                Regex RegStartTijd = new Regex(@"\d\d?():\d\d");
                Match MatchStartTijd = RegStartTijd.Match(item);
                StartTijd = MatchStartTijd.Value;
                DateTime StartTijdDT = Convert.ToDateTime(StartTijd);
                string googleStart = string.Format("{0:HH:mm}", StartTijdDT);

                //DAG
                Regex RegDag = new Regex(@"\d\d");
                Match MatchDag = RegDag.Match(item);
                Dag = Convert.ToInt32(MatchDag.Value);

                //POST
                Regex RegPost = new Regex(@"(\w{3,6})");
                Match MatchPost = RegPost.Match(item);
                Post = MatchPost.Value;
                ListPost.Add(Post);
                Debug.WriteLine(Post);

                //EINDUUR
                Regex RegEindTijd = new Regex(@"([-]\s\s?)(\d?\d[:]\d\d)");
                Match MatchEindTijd = RegEindTijd.Match(item);
                EindTijd = MatchEindTijd.Value;
                Regex EnkelUur = new Regex(@"(\d?\d[:]\d\d)");
                Match MatchEnkelUur = EnkelUur.Match(EindTijd);
                EindTijd = MatchEnkelUur.Value;
                DateTime EindTijdDT = Convert.ToDateTime(EindTijd);
                string googleEind = string.Format("{0:HH:mm}", EindTijdDT);

                DateTime Proof = new DateTime(y, Maand, Dag, 8, 0, 0);
                
                //OMZETTING INDIEN NACHTEN
                if (EindTijdDT.Hour < Proof.Hour)
                {
                    Maand2 = Maand;
                    if (Dag==DateTime.DaysInMonth(y, Maand))
                    {
                        
                        Maand2 = Maand + 1;
                    }
                    if (Maand2!=Maand)
                        //Als de maanden niet gelijk zijn wil dit zeggen dat de nacht valt op een overgang tussen twee maanden
                        //Dus dan moet de dag waarop de shift eindigd op 1 worden gezet.
                    {
                        Dag2 = 1;
                    }
                    else
                    //Indien het een gewone nacht is moet de dag + 1 worden
                    {
                        Dag2 = Dag+1;
                    }
                    
                }
                else //Als het niet om nachten gaat is de dag waarop de shift begint en eindigd dezelfde
                {
                    Dag2 = Dag;
                    Maand2 = Maand;
                }

                DateTime Datum = new DateTime(y, Maand, Dag);
                string DatumToString = Convert.ToString(Datum.ToShortDateString());

                DateTime Datum2 = new DateTime(y, Maand2, Dag2);
                string DatumToString2 = Convert.ToString(Datum2.ToShortDateString());

                //OUTPUT WINDOW FEEDBACK
                CompleteString = string.Format("Datum: {0,7} Post: {1,7}", DatumToString, Post);
                CompleteString2 = string.Format("Start: {0,7} Eind: {1,7}", StartTijd, EindTijd);
                label1.Text += CompleteString + "\n";//Visual feedback of end result
                label2.Text += CompleteString2 + "\n";

                //GOOGLE API PARSE
                googleStart = string.Format(DatumToString + " " + googleStart);
                gStartDt = Convert.ToDateTime(googleStart);
                googleStart = string.Format("{0:s}", gStartDt);
                ListStart.Add(googleStart);

                googleEind = string.Format(DatumToString2 + " " + googleEind);
                gEindDt = Convert.ToDateTime(googleEind);
                googleEind = string.Format("{0:s}", gEindDt);
                ListEind.Add(googleEind);


                Thread.Sleep(50);

            }

        }
        public void GoogleSender()
        {
            for (int i = 0; i < ListPost.Count; i++)
            {


                UserCredential credential;

                using (var stream = new FileStream("client_secret.json", FileMode.Open, FileAccess.Read))
                {
                    string credPath = System.Environment.GetFolderPath(System.Environment.SpecialFolder.Personal);
                    credPath = Path.Combine(credPath, ".credentials/calendar-dotnet-quickstart.json");
                    credential = GoogleWebAuthorizationBroker.AuthorizeAsync(GoogleClientSecrets.Load(stream).Secrets, Scopes, "user", CancellationToken.None, new FileDataStore(credPath, true)).Result;
                }

                var service = new CalendarService(new BaseClientService.Initializer() { HttpClientInitializer = credential, ApplicationName = ApplicationName, });

                {
                    E.Add(new EventReminder { Method = "popup", Minutes = Convert.ToInt32( frm.ReminderMinutes), });

                    Event newEvent = new Event()
                    {
                        Summary = ListPost[i],

                        Start = new EventDateTime()
                        {
                            DateTime = DateTime.Parse(ListStart[i]),
                        },

                        End = new EventDateTime()
                        {
                            DateTime = DateTime.Parse(ListEind[i]),
                        },

                        Location = "BEATRIJSLAAN 100, 2050 ANTWERPEN",

                        Reminders = new Event.RemindersData()
                        {
                           Overrides = E, 
                           UseDefault = false,
                        },


                        


                    };

                    String calendarId = "primary";
                    EventsResource.InsertRequest request = service.Events.Insert(newEvent, calendarId);
                    Event createdEvent = request.Execute();
                }

            }

        }

        public List<EventReminder> E = new List<EventReminder>();




        private void button1_Click(object sender, EventArgs e)
        {
            
            try
            { PdfConvertor();}
            catch (Exception)
            {

               toolStripStatusLabel1.Text = "Error : please supply program with readable pdf & correct filename!";

            }



        }

        private void button2_Click(object sender, EventArgs e)
        {
            GoogleSender();

            MessageBox.Show("Data uploaded. \nPlease reload program to start again!");
            button2.Visible = false;

        }

        private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Deckx Stijn & Pdf convertor");
        }

        private void configureToolStripMenuItem_Click(object sender, EventArgs e)
        {
            
            frm.Show();
        }
    }

}


