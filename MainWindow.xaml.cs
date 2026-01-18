using Microsoft.Win32;
using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Threading;
using Windows.Media.Capture;
using Windows.Media.SpeechRecognition;
using Windows.Networking.Connectivity;


namespace wpf3625_147pm
{
    public class CheckInternetConnection
    {
        public bool ConnectInternet()
        {
            ConnectionProfile internetProfile = NetworkInformation.GetInternetConnectionProfile();
            if (internetProfile == null)
            {
                return false;
            }
            var level = internetProfile.GetNetworkConnectivityLevel();
            return level == NetworkConnectivityLevel.InternetAccess;
        }
    }


    public partial class MainWindow : Window
    {
        private MediaCapture mediaCapture;
        private SpeechRecognizer speechRecognizer;
        private DispatcherTimer dispatcherTimer;
        private StringBuilder stringBuilder;
        private string appDirectory;


        public MainWindow()
        {
            InitializeComponent();
            CheckInternetConnection();
            stringBuilder = new StringBuilder();
            string documentsPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments); documentsPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            string appDirectoryName = "Gabscriber";
            appDirectory = Path.Combine(documentsPath, appDirectoryName);
            Directory.CreateDirectory(appDirectory);
        }


        private void CheckInternetConnection()
        {
            CheckInternetConnection internetChecker = new CheckInternetConnection();
            if (internetChecker.ConnectInternet() == false)
            {
                InternetNotice.Visibility = Visibility.Visible;
                ButtonStart.IsEnabled = false;
                ButtonStart.Opacity = 0.5;
            }

            else
            {
                InternetNotice.Visibility = Visibility.Collapsed;
                ButtonStart.IsEnabled = true;
                ButtonStart.Opacity = 1;
            }
        }

        
        private async Task InitMicSpeech()
        {            
            try
            {
                // Microphone
                MediaCapture mediaCapture = new MediaCapture();
                var settings = new MediaCaptureInitializationSettings();
                settings.StreamingCaptureMode = StreamingCaptureMode.Audio;
                await mediaCapture.InitializeAsync(settings);

                // Speech 
                speechRecognizer = new SpeechRecognizer();
                await speechRecognizer.CompileConstraintsAsync();
                speechRecognizer.ContinuousRecognitionSession.AutoStopSilenceTimeout = TimeSpan.FromMilliseconds(0);
                SpeechRecognitionCompilationResult result = await speechRecognizer.CompileConstraintsAsync();
                speechRecognizer.ContinuousRecognitionSession.Completed += ContinuousRecognitionSession_Completed;
                speechRecognizer.ContinuousRecognitionSession.ResultGenerated += ContinuousRecognitionSession_ResultGenerated;
                MicNotice.Visibility = Visibility.Collapsed;
                SpeechNotice.Visibility = Visibility.Collapsed;
                TbScrollViewer.Visibility = Visibility.Visible;
                TbTranscription.Visibility = Visibility.Visible;
                LabelReminder.Visibility = Visibility.Visible;
                TbStatusState.Visibility = Visibility.Visible;
                ButtonStart.IsEnabled = false;
                ButtonSave.IsEnabled = true;
                ButtonSave.Opacity = 1;
                ButtonStop.IsEnabled = true;
                ButtonStop.Opacity = 1;

                if (speechRecognizer.State == SpeechRecognizerState.Idle)
                {
                    await speechRecognizer.ContinuousRecognitionSession.StartAsync();
                }
            }

            catch (System.UnauthorizedAccessException)
            {
                MicNotice.Visibility = Visibility.Visible;
            }

            catch (Exception)
            {
                SpeechNotice.Visibility = Visibility.Visible;
            }
        }


        private void ButtonFaqClick(object sender, RoutedEventArgs e)
        {
            WindowFaq newWindow = new();
            newWindow.Show();
        }


        private async void ButtonStartClick(object sender, RoutedEventArgs e)
        {
            await InitMicSpeech();
            dispatcherTimer = new DispatcherTimer();
            dispatcherTimer.Tick += new EventHandler(Dispatcher_BreakNotice);
            dispatcherTimer.Interval = new TimeSpan(0, 10, 0);
            dispatcherTimer.Start();
        }


        private void Dispatcher_BreakNotice(object sender, EventArgs e)
        {
            string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            string filenameBreak = $"gabscriber_break{timestamp}.txt";
            string filePathBreak = Path.Combine(appDirectory, filenameBreak);
            File.WriteAllText(filePathBreak, TbTranscription.Text);
            BreakNotice.Visibility = Visibility.Visible;
            
        }


        private async void ContinuousRecognitionSession_ResultGenerated(SpeechContinuousRecognitionSession sender, SpeechContinuousRecognitionResultGeneratedEventArgs args)
        {
            if (args.Result.Confidence == SpeechRecognitionConfidence.Medium || args.Result.Confidence == SpeechRecognitionConfidence.High)
            {
                stringBuilder.Append(args.Result.Text + " ");
                await Dispatcher.InvokeAsync(new Action(() =>
                {
                    TbTranscription.Text = stringBuilder.ToString();
                    TbScrollViewer.ScrollToBottom();
                }));
            }
        }


        private async void ContinuousRecognitionSession_Completed(SpeechContinuousRecognitionSession sender, SpeechContinuousRecognitionCompletedEventArgs args)
        {
            speechRecognizer = new SpeechRecognizer();
            await speechRecognizer.CompileConstraintsAsync();
            speechRecognizer.ContinuousRecognitionSession.AutoStopSilenceTimeout = TimeSpan.FromMilliseconds(0);
            SpeechRecognitionCompilationResult result = await speechRecognizer.CompileConstraintsAsync();
            speechRecognizer.ContinuousRecognitionSession.Completed += ContinuousRecognitionSession_Completed;
            speechRecognizer.ContinuousRecognitionSession.ResultGenerated += ContinuousRecognitionSession_ResultGenerated;
            await speechRecognizer.ContinuousRecognitionSession.StartAsync();
        }


        private void ContinueInternetClick(object sender, RoutedEventArgs e)
        {
            CheckInternetConnection();
        }


        private async void ContinueMicClick(object sender, RoutedEventArgs e)
        {
            await InitMicSpeech();
        }


        private async void ContinueSpeechClick(object sender, RoutedEventArgs e)
        {
            await InitMicSpeech();
        }


        private void ContinueBreakTimeClick(object sender, RoutedEventArgs e)
        {
            System.Diagnostics.Process.Start(System.Diagnostics.Process.GetCurrentProcess().ProcessName);
            App.Current.Shutdown(); 
        }


        private void ButtonSaveClick(object sender, RoutedEventArgs e)
        {
            ButtonSave.IsEnabled = false;
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*";

            if (saveFileDialog.ShowDialog() == true)
            {
                File.WriteAllText(saveFileDialog.FileName, TbTranscription.Text);
            }
        }


        private void ButtonStopClick(object sender, RoutedEventArgs e)
        {
            ButtonStop.IsEnabled = false;
            string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            string filenameStop = $"gabscriber_stop{timestamp}.txt";
            string filePathStop = Path.Combine(appDirectory, filenameStop);
            File.WriteAllText(filePathStop, TbTranscription.Text); 
            System.Diagnostics.Process.Start(System.Diagnostics.Process.GetCurrentProcess().ProcessName);
            App.Current.Shutdown();
        }


        private void ButtonCloseClick(object sender, RoutedEventArgs e)
        {
            Application.Current.Shutdown();
        }
    }
}
