using Microsoft.ProjectOxford.Face;
using Microsoft.ProjectOxford.Face.Contract;
using Microsoft.ProjectOxford.SpeakerRecognition;
using Microsoft.ProjectOxford.SpeakerRecognition.Contract;
using Microsoft.ProjectOxford.SpeakerRecognition.Contract.Identification;
using NAudio.Wave;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Speech.Synthesis;
using System.Speech.Recognition;
using System.Data.SqlClient;
using System.Threading;
using System.Windows.Threading;
using Microsoft.ProjectOxford.SpeakerRecognition.Contract.Verification;

namespace FaceAndVoiceRegistration
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private webCam webcam;
        private string _selectedFile = "";
        private WaveIn _waveIn;
        private WaveFileWriter _fileWriter;
        private readonly String GroupName = "Write your group name";
        private Person p = new Person();
        private List<Guid> enrolllist = new List<Guid>();
        private SpeechSynthesizer speechSynthesizer;
        private string faceIdentifiedUserName;
        private string voiceIdentifiedUserName;
        private int accountNo;
        private readonly SqlConnection conn = new SqlConnection(@"Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=C:\Users\MK185299\Documents\Visual Studio 2017\Projects\FaceAndVoiceRegistrationSln\FaceAndVoiceRegistration\CustomerDatabase.mdf;Integrated Security=True");
        private SqlCommand cmd = new SqlCommand();
        private SqlDataReader dr;
        private DispatcherTimer timer;
        private int time = 0;
        private Guid verificationId;
        private string userPhrase;
        private readonly string faceAPISubscriptionKey = "Paste your face API subscription key here";
        private readonly string speakerAPISubscriptionKey = "paste your speaker recognition API subscription key here";
        
        public MainWindow()
        {
            InitializeComponent();
            InitializeRecorder();
            Task.Delay(2000);
            webcam = new webCam();
            webcam.InitializeWebCam(ref webImage);
            webcam.Start();
            Task.Delay(2000);
            InitializeRecorder();
            Task.Delay(2000);
            speechSynthesizer = new SpeechSynthesizer();
            speechSynthesizer.Rate = -1;
            timer = new DispatcherTimer();
            timer.Interval = new TimeSpan(0, 0, 1);
            timer.Tick += Timer_Tick;
        }

        private void Timer_Tick(object sender, EventArgs e)
        {
            try
            {
                time++;
                Console.WriteLine(time);
                if (time == 3)
                {
                    //time = 8;
                    timer.Stop();
                    voiceIdentification();
                    
                }
                //if (time == 13)
                //{
                //    timer.Stop();
                //    stopRecord();
                //}
            }
            catch(Exception ex)
            {
                Console.WriteLine("Error : " + ex.Message);
            }
        }


        private void stopRecord()
        {
            if (_waveIn != null)
            {
                _waveIn.StopRecording();
            }
        }

        private async void Window_Loaded(object sender, RoutedEventArgs e)
        {
            try
            {
                
                speechSynthesizer.SpeakAsync("Hi Visitor, Welcome to the new world of NCR.");
                await Task.Delay(1000);
                speechSynthesizer.SpeakAsync("To verify your face. Please put your face clearly infront of the ATM.");

                SpeakerIdentificationServiceClient _serviceClient = new SpeakerIdentificationServiceClient(speakerAPISubscriptionKey);
                
                bool groupExists = false;

                var faceServiceClient = new FaceServiceClient(faceAPISubscriptionKey);
                // Test whether the group already exists
                try
                {

                    Title = String.Format("Request: Group {0} will be used to build a person database. Checking whether the group exists.", GroupName);
                    Console.WriteLine("Request: Group {0} will be used to build a person database. Checking whether the group exists.", GroupName);

                    await faceServiceClient.GetPersonGroupAsync(GroupName);
                    groupExists = true;
                    Title = String.Format("Response: Group {0} exists.", GroupName);
                    Console.WriteLine("Response: Group {0} exists.", GroupName);
                }
                catch (FaceAPIException ex)
                {
                    if (ex.ErrorCode != "PersonGroupNotFound")
                    {
                        Title = String.Format("Response: {0}. {1}", ex.ErrorCode, ex.ErrorMessage);
                        return;
                    }
                    else
                    {
                        Title = String.Format("Response: Group {0} did not exist previously.", GroupName);
                    }
                }

                if (groupExists)
                {
                    Title = String.Format("Success..... Now your Group  {0} ready to use.", GroupName);
                    webcam.Start();
                    return;
                }

                else
                {
                    Console.WriteLine("Group did not exist. First you need to create a group");
                }
            }
            catch(Exception ex)
            {
                Console.WriteLine("Error : ", ex.Message);
            }
        }


        /// <summary>
        /// Initialize NAudio recorder instance
        /// </summary>
        private void InitializeRecorder()
        {
            _waveIn = new WaveIn();
            _waveIn.DeviceNumber = 0;
            int sampleRate = 16000; // 16 kHz
            int channels = 1; // mono
            _waveIn.WaveFormat = new WaveFormat(sampleRate, channels);
            _waveIn.DataAvailable += waveIn_DataAvailable;
            _waveIn.RecordingStopped += waveSource_RecordingStopped;
        }


        /// <summary>
        /// A method that's called whenever there's a chunk of audio is recorded
        /// </summary>
        /// <param name="sender">The sender object responsible for the event</param>
        /// <param name="e">The arguments of the event object</param>
        [MethodImpl(MethodImplOptions.Synchronized)]
        private void waveIn_DataAvailable(object sender, WaveInEventArgs e)
        {
            if (_fileWriter == null)
            {
                Console.WriteLine("Error");
            }
            _fileWriter.Write(e.Buffer, 0, e.BytesRecorded);
        }


        /// <summary>
        /// A listener called when the recording stops
        /// </summary>
        /// <param name="sender">Sender object responsible for event</param>
        /// <param name="e">A set of arguments sent to the listener</param>
        private void waveSource_RecordingStopped(object sender, StoppedEventArgs e)
        {
            _fileWriter.Dispose();
            _fileWriter = null;

            //Dispose recorder object
            _waveIn.Dispose();
            InitializeRecorder();

        }
        
        
        private async void faceIdentifyBtn_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                faceIdentifyBtn.IsEnabled = false;
                //capture photo than save.
                captureImage.Source = webImage.Source;
                Helper.SaveImageCapture((BitmapSource)captureImage.Source);

                string getDirectory = Directory.GetCurrentDirectory();
                string filePath = getDirectory + "\\test1.jpg";

                System.Drawing.Image image1 = System.Drawing.Image.FromFile(filePath);

                var faceServiceClient = new FaceServiceClient(faceAPISubscriptionKey);
                try
                {
                    Title = String.Format("Request: Training group \"{0}\"", GroupName);
                    await faceServiceClient.TrainPersonGroupAsync(GroupName);

                    TrainingStatus trainingStatus = null;
                    while (true)
                    {
                        await Task.Delay(1000);
                        trainingStatus = await faceServiceClient.GetPersonGroupTrainingStatusAsync(GroupName);
                        Title = String.Format("Response: {0}. Group \"{1}\" training process is {2}", "Success", GroupName, trainingStatus.Status);
                        if (trainingStatus.Status.ToString() != "running")
                        {
                            break;
                        }
                    }
                }
                catch (FaceAPIException ex)
                {

                    Title = String.Format("Response: {0}. {1}", ex.ErrorCode, ex.ErrorMessage);
                    faceIdentifyBtn.IsEnabled = true;
                }

                Title = "Detecting....";

                using (Stream s = File.OpenRead(filePath))
                {
                    var faces = await faceServiceClient.DetectAsync(s);
                    var faceIds = faces.Select(face => face.FaceId).ToArray();

                    var faceRects = faces.Select(face => face.FaceRectangle);
                    FaceRectangle[] faceRect = faceRects.ToArray();
                    if (faceRect.Length == 1)
                    {
                        Title = String.Format("Detection Finished. {0} face(s) detected", faceRect.Length);
                        speechSynthesizer.SpeakAsync("We have detected.");
                        speechSynthesizer.SpeakAsync(faceRect.Length.ToString());
                        speechSynthesizer.SpeakAsync("face.");
                        speechSynthesizer.SpeakAsync("Please Wait we are identifying your face.");

                        await Task.Delay(3000);
                        Title = "Identifying.....";
                        try
                        {
                            Console.WriteLine("Group Name is : {0}, faceIds is : {1}", GroupName, faceIds);
                            var results = await faceServiceClient.IdentifyAsync(GroupName, faceIds);

                            foreach (var identifyResult in results)
                            {
                                Title = String.Format("Result of face: {0}", identifyResult.FaceId);

                                if (identifyResult.Candidates.Length == 0)
                                {
                                    Title = String.Format("No one identified");
                                    MessageBox.Show("Hi, Make sure you have registered your face. Try to register now.");
                                    speechSynthesizer.SpeakAsync("Sorry. No one identified.");
                                    speechSynthesizer.SpeakAsync("Please make sure you have previously registered your face with us.");
                                    faceIdentifyBtn.IsEnabled = false;
                                    return;
                                }
                                else
                                {
                                    // Get top 1 among all candidates returned
                                    var candidateId = identifyResult.Candidates[0].PersonId;
                                    var person = await faceServiceClient.GetPersonAsync(GroupName, candidateId);
                                    faceIdentifiedUserName = person.Name.ToString();
                                    Title = String.Format("Identified as {0}", person.Name);
                                    
                                    speechSynthesizer.Speak("Hi.");
                                    speechSynthesizer.Speak(person.Name.ToString());
                                    speechSynthesizer.Speak("Now you need to verify your voice.");
                                    //speechSynthesizer.Speak("To verify your voice. Say like that.");
                                    //speechSynthesizer.Speak("My voice is stronger than my password. Verify my voice.");
                                    speechSynthesizer.Speak("Please speak your phrase.");
                                    speechSynthesizer.Speak("Now Start to Speak.");
                                    faceIdentifyBtn.IsEnabled = false;

                                    try
                                    {
                                        if (WaveIn.DeviceCount == 0)
                                        {
                                            throw new Exception("Cannot detect microphone.");
                                        }

                                        //save file.
                                        _selectedFile = getDirectory + "\\Sample2.wav";
                                        _fileWriter = new NAudio.Wave.WaveFileWriter(_selectedFile, _waveIn.WaveFormat);
                                        Console.WriteLine("Start Speak.");
                                        _waveIn.StartRecording();
                                        timer.Start();
                                        Title = String.Format("Recording...");
                                        GC.Collect();
                                    }
                                    catch (Exception ge)
                                    {
                                        Console.WriteLine("Error: " + ge.Message);
                                        GC.Collect();
                                    }
                                    
                                    
                                }
                            }
                            GC.Collect();
                        }
                        catch (FaceAPIException ex)
                        {
                            Title = String.Format("Failed...Try Again.");
                            speechSynthesizer.SpeakAsync("First register your face.");
                            Console.WriteLine("Error : {0} ", ex.Message);
                            image1.Dispose();
                            File.Delete(filePath);
                            GC.Collect();
                            return;
                        }
                    }
                    else if(faceRect.Length >1)
                    {
                        Title = String.Format("More than one faces detected. Make sure only one face is in the photo. Try again..");
                        speechSynthesizer.SpeakAsync("More than one faces detected. Make sure only one face is in the photo. Try again..");
                        faceIdentifyBtn.IsEnabled = true;
                        return;
                    }
                    else
                    {
                        Title = String.Format("No one detected in the photo. Please make sure your face is infront of the webcam. Try again with the correct photo.");
                        speechSynthesizer.SpeakAsync("No one detected. Please make sure your face is infront of the webcam. Try again with the correct photo.");
                        faceIdentifyBtn.IsEnabled = true;
                        return;
                    }

                    image1.Dispose();
                    File.Delete(filePath);
                    GC.Collect();
                    await Task.Delay(2000);
                }
            }
            catch(Exception ex)
            {
                Console.WriteLine("Error : ", ex.Message);
                faceIdentifyBtn.IsEnabled = true;
                GC.Collect();
            }
        }

        private async void voiceIdentification()
        {
            try
            {
                _identificationResultStckPnl.Visibility = Visibility.Hidden;
                if (_waveIn != null)
                {
                    _waveIn.StopRecording();
                }

                TimeSpan timeBetweenSaveAndIdentify = TimeSpan.FromSeconds(5.0);
                await Task.Delay(timeBetweenSaveAndIdentify);

                SpeakerIdentificationServiceClient _serviceClient = new SpeakerIdentificationServiceClient(speakerAPISubscriptionKey);

                List<Guid> list = new List<Guid>();
                Microsoft.ProjectOxford.SpeakerRecognition.Contract.Identification.Profile[] allProfiles = await _serviceClient.GetProfilesAsync();
                int itemsCount = 0;
                foreach (Microsoft.ProjectOxford.SpeakerRecognition.Contract.Identification.Profile profile in allProfiles)
                {
                    list.Add(profile.ProfileId);
                    itemsCount++;
                }
                Guid[] selectedIds = new Guid[itemsCount];
                for (int i = 0; i < itemsCount; i++)
                {
                    selectedIds[i] = list[i];
                }
                if (_selectedFile == "")
                    throw new Exception("No File Selected.");

                speechSynthesizer.SpeakAsync("Please wait we are verifying your voice.");
                Title = String.Format("Identifying File...");
                OperationLocation processPollingLocation;
                Console.WriteLine("Selected file is : {0}", _selectedFile);
                using (Stream audioStream = File.OpenRead(_selectedFile))
                {
                    //_selectedFile = "";
                    Console.WriteLine("Start");
                    Console.WriteLine("Audio File is : {0}", audioStream);
                    processPollingLocation = await _serviceClient.IdentifyAsync(audioStream, selectedIds, true);
                    Console.WriteLine("ProcesPolling Location : {0}", processPollingLocation);
                    Console.WriteLine("Done");
                }

                IdentificationOperation identificationResponse = null;
                int numOfRetries = 10;
                TimeSpan timeBetweenRetries = TimeSpan.FromSeconds(5.0);
                while (numOfRetries > 0)
                {
                    await Task.Delay(timeBetweenRetries);
                    identificationResponse = await _serviceClient.CheckIdentificationStatusAsync(processPollingLocation);
                    Console.WriteLine("Response is : {0}", identificationResponse);

                    if (identificationResponse.Status == Microsoft.ProjectOxford.SpeakerRecognition.Contract.Identification.Status.Succeeded)
                    {
                        break;
                    }
                    else if (identificationResponse.Status == Microsoft.ProjectOxford.SpeakerRecognition.Contract.Identification.Status.Failed)
                    {
                        Console.WriteLine("In");
                        speechSynthesizer.SpeakAsync("Failed. Please make sure your voice is registered.");
                        throw new IdentificationException(identificationResponse.Message);
                    }
                    numOfRetries--;
                }
                if (numOfRetries <= 0)
                {
                    throw new IdentificationException("Identification operation timeout.");
                }

                Title = String.Format("Identification Done.");

                conn.Open();
                SqlCommand cmd = conn.CreateCommand();
                cmd.CommandType = System.Data.CommandType.Text;
                cmd.CommandText = "Select AccountNo, CustomerName From AccountDetails where AccountNo = (Select AccountNo From AuthenticationDetails where VoiceId = '" + identificationResponse.ProcessingResult.IdentifiedProfileId.ToString() + "')";
                dr = cmd.ExecuteReader();
                if (dr.HasRows)
                {
                    while (dr.Read())
                    {
                        accountNo = dr.GetInt32(0);
                        voiceIdentifiedUserName = dr[1].ToString();
                        Console.WriteLine("Account No is : " + accountNo);
                        Console.WriteLine("Identified as :" + voiceIdentifiedUserName);
                        _identificationResultTxtBlk.Text = voiceIdentifiedUserName;
                    }
                }
                dr.Close();
                conn.Close();
                if (_identificationResultTxtBlk.Text == "")
                {
                    _identificationResultTxtBlk.Text = identificationResponse.ProcessingResult.IdentifiedProfileId.ToString();
                    speechSynthesizer.SpeakAsync("Sorry we have not found your data.");
                    return;
                }
                else
                {

                    if (faceIdentifiedUserName == voiceIdentifiedUserName)
                    {
                        Console.WriteLine("Selected file is : {0}", _selectedFile);

                        Stream stream = File.OpenRead(_selectedFile);
                        verifySpeaker(stream);
                        //speechSynthesizer.SpeakAsync("Hi.");
                        //speechSynthesizer.SpeakAsync(_identificationResultTxtBlk.Text.ToString());
                        //speechSynthesizer.SpeakAsync("Thanks to verify your face and voice.");
                        //speechSynthesizer.SpeakAsync("Now you can do your transactions");
                        
                    }
                    else
                    {
                        speechSynthesizer.SpeakAsync("Sorry we have found different voice identity from your face identity.");
                        return;
                    }
                    _identificationConfidenceTxtBlk.Text = identificationResponse.ProcessingResult.Confidence.ToString();
                    _identificationResultStckPnl.Visibility = Visibility.Visible;
                    GC.Collect();
                }
            }
            catch (IdentificationException ex)
            {
                Console.WriteLine("Speaker Identification Error : " + ex.Message);
                GC.Collect();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error : " + ex.Message);
                GC.Collect();
            }
        }


        private async void verifySpeaker(Stream audioStream)
        {
            try
            {
                conn.Open();
                SqlCommand cmd = conn.CreateCommand();
                cmd.CommandType = System.Data.CommandType.Text;
                cmd.CommandText = "Select VerifyId, Phrase From AuthenticationDetails where AccountNo = '" + accountNo + "'";
                dr = cmd.ExecuteReader();
                if (dr.HasRows)
                {
                    while (dr.Read())
                    {
                        verificationId = dr.GetGuid(0);
                        userPhrase = dr[1].ToString();
                        Console.WriteLine("Verification Id is : " + verificationId);
                        Console.WriteLine("User Phrase is :" + userPhrase);
                    }
                }
                dr.Close();
                conn.Close();
                //verificationId = Guid.Parse("6c3a49ee-aa36-4a45-b7de-068cd96516bc");
                SpeakerVerificationServiceClient verServiceClient = new SpeakerVerificationServiceClient(speakerAPISubscriptionKey);
                Title = String.Format("Verifying....");
                Console.WriteLine("Verifying....");
                Verification response = await verServiceClient.VerifyAsync(audioStream, verificationId);
                Title = String.Format("Verification Done.");
                Console.WriteLine("Verrification Done.");
                //statusResTxt.Text = response.Result.ToString();
                //confTxt.Text = response.Confidence.ToString();
                Console.WriteLine("Response Result is : " + response.Result.ToString());
                Console.WriteLine("Response Confidence is : " + response.Confidence.ToString());
                if (response.Result == Result.Accept)
                {
                    //statusResTxt.Background = Brushes.Green;
                    //statusResTxt.Foreground = Brushes.White;
                    speechSynthesizer.SpeakAsync("Hi.");
                    speechSynthesizer.SpeakAsync(_identificationResultTxtBlk.Text.ToString());
                    speechSynthesizer.SpeakAsync("Thanks to verify your face and voice.");
                    speechSynthesizer.SpeakAsync("Now you can do your transactions");
                    //return;
                }
                else
                {
                    //statusResTxt.Background = Brushes.Red;
                    //statusResTxt.Foreground = Brushes.White;
                    speechSynthesizer.SpeakAsync("Sorry verification failed.");
                    speechSynthesizer.SpeakAsync("Please speak correct phrase.");
                    return;
                }
            }
            catch (VerificationException exception)
            {
                Console.WriteLine("Cannot verify speaker: " + exception.Message);
            }
            catch (Exception ex)
            {

                Console.WriteLine("Error : " + ex.Message);
            }
        }


        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            if(speechSynthesizer.State == SynthesizerState.Speaking)
            {
                speechSynthesizer.Dispose();
            }
        }

        private void verifyBtn_Click(object sender, RoutedEventArgs e)
        {
            voiceIdentification();
        }

        private void recordBtn_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (WaveIn.DeviceCount == 0)
                {
                    throw new Exception("Cannot detect microphone.");
                }

                //save file.
                string getDirectory = Directory.GetCurrentDirectory();
                string _selectedFile = getDirectory + "\\Sample3.wav";
                _fileWriter = new NAudio.Wave.WaveFileWriter(this._selectedFile, _waveIn.WaveFormat);
                _waveIn.StartRecording();
                //timer.Start();
                Title = String.Format("Recording...");
                Stream streams = File.OpenRead(_selectedFile);
                //verifySpeaker(streams);
                GC.Collect();
            }
            catch (Exception ge)
            {
                Console.WriteLine("Error: " + ge.Message);
                GC.Collect();
            }
        }
    }
}
