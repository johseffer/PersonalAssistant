using AudioSwitcher.AudioApi.CoreAudio;
using Microsoft.Owin.Hosting;
using Microsoft.Speech.Recognition;
using Microsoft.Speech.Synthesis;
using System;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PersonalAssistant
{
    internal class Program
    {
        private static readonly SpeechSynthesizer ss = new SpeechSynthesizer();
        private static SpeechRecognitionEngine sre;
        private static CoreAudioDevice defaultPlaybackDevice;
        private static readonly bool done = false;
        private static bool speechOn = false;
        private static bool loginRequested = false;
        private static bool locked = true;
        private static readonly System.Windows.Forms.Timer t1 = new System.Windows.Forms.Timer();
        private static Form loadForm;
        public const int KEYEVENTF_EXTENTEDKEY = 1;
        public const int KEYEVENTF_KEYUP = 0;
        public const int VK_MEDIA_NEXT_TRACK = 0xB0;
        public const int VK_MEDIA_PLAY_PAUSE = 0xB3;
        public const int VK_MEDIA_PREV_TRACK = 0xB1;

        [DllImport("user32.dll")]
        public static extern void keybd_event(byte virtualKey, byte scanCode, uint flags, IntPtr extraInfo);


        [DllImport("user32.dll")]
        public static extern bool LockWorkStation();

        private static void Main(string[] args)
        {
            try
            {
                SystemInitialize();
                Application.Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.ReadLine();
            }
        }

        private static void StartServer()
        {
            try
            {
                var ServerURI = $"http://localhost:8060";
                var SignalR = WebApp.Start<Startup>(ServerURI);
            }
            catch (Exception ex)
            {

            }
        }

        private static void SystemInitialize()
        {
            Thread thread = new Thread(() => OpenLoadForm());
            thread.Start();

            ss.SetOutputToDefaultAudioDevice();
            ss.SelectVoice("Microsoft Server Speech Text to Speech Voice (pt-BR, Heloisa)");
            ss.Speak("Inicializando Sistema. Aguarde.");
            Console.WriteLine("\nSystem Initializing ...");

            defaultPlaybackDevice = new CoreAudioController().DefaultPlaybackDevice;

            CultureInfo ci = new CultureInfo("pt-br");
            sre = new SpeechRecognitionEngine(ci);
            sre.SetInputToDefaultAudioDevice();
            sre.SpeechRecognized += sre_SpeechRecognized;

            LoadCommands();
            StartServer();

            sre.RecognizeAsync(RecognizeMode.Multiple); // multiple grammars

            Console.WriteLine("\nSystem started!");
            loadForm.Invoke(new Action(() => { loadForm.Close(); }));
            //thread.Abort();
        }

        private static void sre_SpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {
            string txt = e.Result.Text;
            float confidence = e.Result.Confidence; // consider implicit cast to double
            Console.WriteLine("\nRecognized: " + txt);

            if (!IsValidConfidence(confidence))
            {
                return;
            }
            else if (txt.IndexOf("desativar") < 0 && txt.IndexOf("ativar") >= 0)
            {
                Activate();
            }
            else if (txt.IndexOf("desativar") >= 0)
            {
                Desactivate();
            }
            else if (txt.IndexOf("Papaya") >= 0 && locked && loginRequested)
            {
                DoPassword();
            }
            else if (speechOn == false || locked)
            {
                return;
            }
            else if (txt.IndexOf("Tela Cheia") >= 0)
            {
                SendKeys.SendWait("^f");
            }
            else if (txt.IndexOf("Sair do Assistente") >= 0)
            {
                ss.Speak("Até Logo!");
                Environment.Exit(200);
            }
            else if (txt.IndexOf("Fechar Aba") >= 0)
            {
                SendKeys.SendWait("^w");
            }
            else if (txt.IndexOf("Nova Aba") >= 0)
            {
                SendKeys.SendWait("^t");
            }
            else if (txt.IndexOf("Bloquear") >= 0)
            {
                LockWorkStation();
                Desactivate();
            }
            else if (txt.IndexOf("Play") >= 0 || txt.IndexOf("Pause") >= 0)
            {
                keybd_event(VK_MEDIA_PLAY_PAUSE, 0, KEYEVENTF_EXTENTEDKEY, IntPtr.Zero);
            }
            else if (txt.IndexOf("Próxima") >= 0)
            {
                keybd_event(VK_MEDIA_NEXT_TRACK, 0, KEYEVENTF_EXTENTEDKEY, IntPtr.Zero);
            }
            else if (txt.IndexOf("Anterior") >= 0)
            {
                keybd_event(VK_MEDIA_PREV_TRACK, 0, KEYEVENTF_EXTENTEDKEY, IntPtr.Zero);
            }
            else if (txt.IndexOf("Abrir") >= 0)
            {
                DoOpen(txt);
            }
            else if (txt.IndexOf("Volume") >= 0)
            {
                SetVolume(txt);
            }
        }

        private static void OpenLoadForm()
        {
            loadForm = new Form();
            loadForm.Opacity = 0;
            loadForm.FormClosing += FormLoad_FormClosing;
            loadForm.Shown += new EventHandler((s, e) => { LoadForm_Shown(loadForm, e); });
            loadForm.BackColor = Color.Black;
            loadForm.WindowState = FormWindowState.Normal;
            //loadForm.Location = Screen.AllScreens[0].WorkingArea.Location;
            loadForm.FormBorderStyle = FormBorderStyle.None;
            loadForm.Bounds = Screen.PrimaryScreen.Bounds;
            loadForm.MinimizeBox = false;
            loadForm.MaximizeBox = false;
            PictureBox pb = new PictureBox
            {
                SizeMode = PictureBoxSizeMode.AutoSize
            };
            Bitmap image = PersonalAssistant.Properties.Resources.skull2;
            pb.Location = new Point((loadForm.ClientSize.Width / 2) - (image.Width / 2), (loadForm.ClientSize.Height / 2) - (image.Height / 2));
            pb.Image = image;
            loadForm.Controls.Add(pb);
            loadForm.Show();
            Application.Run(loadForm);
        }

        private static async void LoadForm_Shown(object sender, EventArgs e)
        {
            await FadeIn(loadForm, 80);
        }

        private static void FormLoad_FormClosing(object sender, FormClosingEventArgs e)
        {
            e.Cancel = true;    //cancel the event so the form won't be closed

            t1.Tick += new EventHandler((s, ev) => FormLoadFadeOut(sender, ev));  //this calls the fade out function
            t1.Start();

            if (((Form)sender).Opacity == 0)  //if the form is completly transparent
            {
                e.Cancel = false;   //resume the event - the program can be closed
                ss.Speak("Sistema inicializado.");
            }
        }

        private static void FormLoadFadeOut(object sender, EventArgs e)
        {
            if (((Form)sender).Opacity <= 0)
            {
                t1.Stop();
                ((Form)sender).Close();
            }
            else
            {
                ((Form)sender).Opacity -= 0.05;
            }
        }

        private static async Task FadeIn(Form o, int interval = 80)
        {
            //Object is not fully invisible. Fade it in
            while (o.Opacity < 1.0)
            {
                await Task.Delay(interval);
                o.Opacity += 0.05;
            }
            o.Opacity = 1; //make fully visible       
        }

        private static void LoadCommands()
        {
            sre.LoadGrammarAsync(GetStartStopCommand());
            sre.LoadGrammarAsync(GetPasswordCommand());
            sre.LoadGrammarAsync(GetLockCommand());
            sre.LoadGrammarAsync(GetExitCommand());
            sre.LoadGrammarAsync(GetSpotifyPlayCommand());
            sre.LoadGrammarAsync(GetSpotifyPauseCommand());
            sre.LoadGrammarAsync(GetSpotifyNextCommand());
            sre.LoadGrammarAsync(GetSpotifyPrevCommand());
            sre.LoadGrammarAsync(GetVolumeCommand());
            sre.LoadGrammarAsync(GetFullScreenCommand());
            sre.LoadGrammarAsync(GetNewTabCommand());
            sre.LoadGrammarAsync(GetCloseTabCommand());
            sre.LoadGrammarAsync(GetOpenProgrammCommand());
        }

        private static Grammar GetFullScreenCommand()
        {
            GrammarBuilder gb_fullscreen = new GrammarBuilder();
            gb_fullscreen.Append("Tela Cheia");
            return new Grammar(gb_fullscreen);
        }

        private static Grammar GetCloseTabCommand()
        {
            GrammarBuilder gb_fullscreen = new GrammarBuilder();
            gb_fullscreen.Append("Fechar Aba");
            return new Grammar(gb_fullscreen);
        }

        private static Grammar GetNewTabCommand()
        {
            GrammarBuilder gb_fullscreen = new GrammarBuilder();
            gb_fullscreen.Append("Nova Aba");
            return new Grammar(gb_fullscreen);
        }

        private static Grammar GetSpotifyPlayCommand()
        {
            GrammarBuilder gb_StartStop = new GrammarBuilder();
            gb_StartStop.Append("Play");
            return new Grammar(gb_StartStop);
        }

        private static Grammar GetSpotifyPauseCommand()
        {
            GrammarBuilder gb_StartStop = new GrammarBuilder();
            gb_StartStop.Append("Pause");
            return new Grammar(gb_StartStop);
        }

        private static Grammar GetSpotifyNextCommand()
        {
            GrammarBuilder gb_StartStop = new GrammarBuilder();
            gb_StartStop.Append("Próxima");
            return new Grammar(gb_StartStop);
        }

        private static Grammar GetSpotifyPrevCommand()
        {
            GrammarBuilder gb_StartStop = new GrammarBuilder();
            gb_StartStop.Append("Anterior");
            return new Grammar(gb_StartStop);
        }

        private static Grammar GetPasswordCommand()
        {
            GrammarBuilder gb_StartStop = new GrammarBuilder();
            gb_StartStop.Append("Papaya");
            return new Grammar(gb_StartStop);
        }

        private static Grammar GetStartStopCommand()
        {
            Choices ch_StartStopCommands = new Choices();
            ch_StartStopCommands.Add("ativar");
            ch_StartStopCommands.Add("desativar");

            GrammarBuilder gb_StartStop = new GrammarBuilder();
            gb_StartStop.Append(ch_StartStopCommands);
            return new Grammar(gb_StartStop);
        }

        private static Grammar GetLockCommand()
        {
            GrammarBuilder gb_Lock = new GrammarBuilder();
            gb_Lock.Append("Bloquear");
            return new Grammar(gb_Lock);
        }

        private static Grammar GetExitCommand()
        {
            GrammarBuilder gb_Exit = new GrammarBuilder();
            gb_Exit.Append("Sair do Assistente");
            return new Grammar(gb_Exit);
        }

        private static Grammar GetVolumeCommand()
        {
            Choices ch_VolumeChoices = new Choices();
            ch_VolumeChoices.Add("0");
            ch_VolumeChoices.Add("20");
            ch_VolumeChoices.Add("40");
            ch_VolumeChoices.Add("60");
            ch_VolumeChoices.Add("80");
            ch_VolumeChoices.Add("100");

            GrammarBuilder gb_VolumeSet = new GrammarBuilder();
            gb_VolumeSet.Append("Volume");
            gb_VolumeSet.Append(ch_VolumeChoices);
            return new Grammar(gb_VolumeSet);
        }

        private static Grammar GetOpenProgrammCommand()
        {
            Choices ch_OpenProgram = new Choices();
            ch_OpenProgram.Add("Visual Studio");
            ch_OpenProgram.Add("Netflix");
            ch_OpenProgram.Add("Twitch");
            ch_OpenProgram.Add("Spotify");

            GrammarBuilder gb_OpenProgram = new GrammarBuilder();
            gb_OpenProgram.Append("Abrir");
            gb_OpenProgram.Append(ch_OpenProgram);
            return new Grammar(gb_OpenProgram);
        }

        private static bool IsValidConfidence(float confidence)
        {
            return confidence > 0.40;
        }

        private static void Activate()
        {
            if (!speechOn)
            {
                RequestLogin();
            }
        }

        private static void Desactivate()
        {
            Console.WriteLine("Reconhecimento desativado.");
            ss.Speak("Reconhecimento desativado.");
            speechOn = false;
            locked = true;
            loginRequested = false;
        }

        private static void DoPassword()
        {
            locked = false;
            speechOn = true;
            ss.Speak("Acesso autorizado. Bem Vindo jô zefer.");
        }

        private static void DoOpen(string txt)
        {
            if (txt.Split(' ').Length <= 1)
            {
                ss.Speak("Formato do comando incorreto.");
                Console.WriteLine("Incorrect parameters! Range: [10, 20, 40, 60, 80, 100]");
            }
            else
            {
                string programmName = txt.Substring(6, txt.Length - 6);
                Console.WriteLine($"Openning {programmName}.");
                switch (programmName.ToUpper())
                {
                    case "VISUAL STUDIO":
                        {
                            Process.Start(@"C:\Program Files (x86)\Microsoft Visual Studio\2019\Preview\Common7\IDE\devenv.exe");
                            break;
                        }
                    case "TWITCH":
                        {
                            Process.Start(@"https://www.twitch.tv/");
                            break;
                        }
                    case "NETFLIX":
                        {
                            Process.Start(@"https://www.netflix.com/");
                            break;
                        }
                    case "SPOTIFY":
                        {
                            var fileName = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), @"Spotify\Spotify.exe");
                            if (File.Exists(fileName))
                            {
                                Process.Start(fileName);
                            }
                            else
                            {
                                ss.Speak("Aplicativo spotify não encontrado!");
                            }
                            break;
                        }
                    default:
                        {
                            break;
                        }
                }
            }
        }

        private static void SetVolume(string txt)
        {
            int volume = 0;
            if (txt.Split(' ').Length <= 1 || !int.TryParse(txt.Split(' ')[1], out volume))
            {
                ss.Speak("Formato do comando incorreto.");
                Console.WriteLine("Incorrect parameters! Range: [10, 20, 40, 60, 80, 100]");
            }
            else
            {
                Console.WriteLine($"Setting volume to {volume}");
                Debug.WriteLine("Current Volume:" + defaultPlaybackDevice.Volume);
                defaultPlaybackDevice.Volume = volume;
            }
        }

        private static void RequestLogin()
        {
            Console.WriteLine("Informe sua senha:");
            ss.Speak("Informe sua senha:");
            loginRequested = true;
        }
    }
}
