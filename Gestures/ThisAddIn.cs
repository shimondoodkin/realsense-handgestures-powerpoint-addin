using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

/*
desired list of gesture commands:
 * prev slide swipe
 * next slide swipe
 * new slide high five 
desired list of voice commands:
 * prev slide
 * previouis slide
 * create new slide
 */
namespace Gestures
{
    public partial class ThisAddIn
    {
        public volatile bool stop = false;
        public volatile bool pause = false;
        //private volatile bool closing = false;
        //private bool indictatemode = false;
        private bool disconnected = false;
        private PXCMGesture.GeoNode[][] nodes = new PXCMGesture.GeoNode[2][] { new PXCMGesture.GeoNode[11], new PXCMGesture.GeoNode[11] };
        private PXCMGesture.Gesture[] gestures = new PXCMGesture.Gesture[2];
        private static System.Timers.Timer unbusytimer;
        public MyUtilMPipeline ppp = null;

        private bool busy = false;
        /*
         * example:
                            if (!busy)
                            {
                                busy = true;
                                unbusytimer.Enabled = true;
                                newslide();
                            }
        */
        private void OnTimedEvent(object source, System.Timers.ElapsedEventArgs e)
        {
            unbusytimer.Enabled = false;
            busy = false;
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            unbusytimer = new System.Timers.Timer(2000);
            unbusytimer.Elapsed += new System.Timers.ElapsedEventHandler(OnTimedEvent);
            unbusytimer.Enabled = false;
            Start();
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            Stop();
            unbusytimer.Dispose();
            System.Threading.Thread.Sleep(5);
        }

        private delegate void DisplayGesturesDelegate(PXCMGesture.Gesture[] gestures);
        public void DisplayGestures(PXCMGesture.Gesture[] gestures)
        {
            Random Ramdom = new Random();



            int randomNumber = Ramdom.Next(0, 100);
            //if (!Gesture.Checked) return; // maybe used

            bool gotcommand = false;

            //Gesture1.Invoke(new DisplayGesturesDelegate(delegate(PXCMGesture.Gesture[] data)
            //{
            if (gestures != null) if (gestures[0].label > 0)
                {
                    if (gestures[0].label == PXCMGesture.Gesture.Label.LABEL_NAV_SWIPE_RIGHT)
                    {
                        gotcommand = true;
                        if (!busy)
                        {
                            busy = true;
                            unbusytimer.Interval = 500;
                            unbusytimer.Enabled = true;
                            gotoSlide(1);
                        }
                    }

                    if (gestures[0].label == PXCMGesture.Gesture.Label.LABEL_NAV_SWIPE_LEFT)
                    {
                        gotcommand = true;
                        if (!busy)
                        {
                            busy = true;
                            unbusytimer.Interval = 500;
                            unbusytimer.Enabled = true;
                            gotoSlide(-1);
                        }
                    }

                    if (gestures[0].label == PXCMGesture.Gesture.Label.LABEL_NAV_SWIPE_DOWN)
                    {
                        gotcommand = true;
                        if (!busy)
                        {
                            busy = true;
                            unbusytimer.Interval = 500;
                            unbusytimer.Enabled = true;
                            gotoSlide(1);
                        }
                    }

                    if (gestures[0].label == PXCMGesture.Gesture.Label.LABEL_NAV_SWIPE_UP)
                    {
                        gotcommand = true;
                        if (!busy)
                        {
                            busy = true;
                            unbusytimer.Interval = 500;
                            unbusytimer.Enabled = true;
                            gotoSlide(-1);
                        }
                    }

                    if (gestures[0].label == PXCMGesture.Gesture.Label.LABEL_POSE_BIG5)
                    {
                        gotcommand = true;
                        if (!busy)
                        {
                            busy = true;
                            unbusytimer.Interval = 2000;
                            unbusytimer.Enabled = true;
                            newslide();
                        }
                    }


                    if (gestures[0].label == PXCMGesture.Gesture.Label.LABEL_POSE_THUMB_UP)
                    {
                        gotcommand = true;
                        if (!busy)
                        {
                            busy = true;
                            unbusytimer.Interval = 2000;
                            unbusytimer.Enabled = true;
                            startshow();
                        }
                    }



                    if (gestures[0].label == PXCMGesture.Gesture.Label.LABEL_HAND_WAVE)
                    {
                        gotcommand = true;
                        if (!busy)
                        {
                            busy = true;
                            unbusytimer.Interval = 2000;
                            unbusytimer.Enabled = true;
                            endshow();
                        }
                    }

                    //System.Console.WriteLine(gestures[0].label.ToString());
                    //  Globals.ThisAddIn.Application.ActivePresentation.Slides[1].Shapes.Title.TextFrame.TextRange.Text += "right " + gestures[0].label.ToString() + "\n";
                    //Globals.ThisAddIn.Application.ActivePresentation.Slides[1].Shapes.Title.TextFrame.TextRange.Text = "right " + gestures[0].label.ToString();
                    System.Diagnostics.Debug.Print("right " + gestures[0].label.ToString());
                    // // required
                    //Gesture1.Image = (Bitmap)pictures[data[0].label];
                    //Gesture1.Invalidate();
                    //  timer.Start();
                }
            //}), new object[] { gestures });

            //Gesture2.Invoke(new DisplayGesturesDelegate(delegate(PXCMGesture.Gesture[] data)
            //{
            if (!gotcommand)
                if (gestures != null) if (gestures[1].label > 0)
                    {
                        if (gestures[1].label == PXCMGesture.Gesture.Label.LABEL_NAV_SWIPE_RIGHT)
                        {
                            if (!busy)
                            {
                                busy = true;
                                unbusytimer.Interval = 500;
                                unbusytimer.Enabled = true;
                                gotoSlide(-1);
                            }
                        }

                        if (gestures[1].label == PXCMGesture.Gesture.Label.LABEL_NAV_SWIPE_LEFT)
                        {
                            if (!busy)
                            {
                                busy = true;
                                unbusytimer.Interval = 500;
                                unbusytimer.Enabled = true;
                                gotoSlide(1);
                            }
                        }


                        if (gestures[1].label == PXCMGesture.Gesture.Label.LABEL_NAV_SWIPE_DOWN)
                        {
                            gotcommand = true;
                            if (!busy)
                            {
                                busy = true;
                                unbusytimer.Interval = 500;
                                unbusytimer.Enabled = true;
                                gotoSlide(1);
                            }
                        }

                        if (gestures[1].label == PXCMGesture.Gesture.Label.LABEL_NAV_SWIPE_UP)
                        {
                            gotcommand = true;
                            if (!busy)
                            {
                                busy = true;
                                unbusytimer.Interval = 500;
                                unbusytimer.Enabled = true;
                                gotoSlide(-1);
                            }
                        }

                        if (gestures[1].label == PXCMGesture.Gesture.Label.LABEL_POSE_BIG5)
                        {
                            gotcommand = true;
                            if (!busy)
                            {
                                busy = true;
                                unbusytimer.Interval = 2000;
                                unbusytimer.Enabled = true;
                                newslide();
                            }

                        }

                        if (gestures[1].label == PXCMGesture.Gesture.Label.LABEL_POSE_THUMB_UP)
                        {
                            gotcommand = true;
                            if (!busy)
                            {
                                busy = true;
                                unbusytimer.Interval = 2000;
                                unbusytimer.Enabled = true;
                                startshow();
                            }
                        }



                        if (gestures[1].label == PXCMGesture.Gesture.Label.LABEL_HAND_WAVE)
                        {
                            gotcommand = true;
                            if (!busy)
                            {
                                busy = true;
                                unbusytimer.Interval = 2000;
                                unbusytimer.Enabled = true;
                                endshow();
                            }
                        }


                        //System.Console.WriteLine(gestures[1].label.ToString());

                        System.Diagnostics.Debug.Print("left " + gestures[1].label.ToString());
                        //data[1].label// required
                        //     Gesture2.Image = (Bitmap)pictures[data[1].label];
                        //     Gesture2.Invalidate();
                        //     timer.Start();
                    }
            //}), new object[] { gestures });

        }

        private void DisplayGesture(PXCMGesture gesture)
        {
            gesture.QueryGestureData(0, PXCMGesture.GeoNode.Label.LABEL_BODY_HAND_PRIMARY, 0, out gestures[0]);
            gesture.QueryGestureData(0, PXCMGesture.GeoNode.Label.LABEL_BODY_HAND_SECONDARY, 0, out gestures[1]);
            DisplayGestures(gestures);
        }


        delegate void DoRecognitionCompleted();
        private void DoRecognition()
        {
            //GestureRecognition gr = new GestureRecognition(this);
            //if (simpleToolStripMenuItem.Checked)
            //{
            this.SimplePipeline();
            //}
            //else
            //{
            //gr.AdvancedPipeline();
            //}
            //this.Invoke(new DoRecognitionCompleted(
            //    delegate
            //    {
            //Start.Enabled = true;
            //Stop.Enabled = false;
            //MainMenu.Enabled = true;
            //if (closing) Close();
            //    }
            // ));
        }



        //private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        // {
        //    stop = true;
        //   e.Cancel = Stop.Enabled;
        //   closing = true;
        // }

        //private void Stop_Click(object sender, EventArgs e)
        private void Stop()
        {
            stop = true;
        }

        private bool DisplayDeviceConnection(bool state)
        {
            if (state)
            {
                // if (!disconnected) form.UpdateStatus("Device Disconnected");
                disconnected = true;
            }
            else
            {
                //if (disconnected) form.UpdateStatus("Device Reconnected");
                disconnected = false;
            }
            return disconnected;
        }


        //private void Start_Click(object sender, EventArgs e)
        private void Start()
        {
            //MainMenu.Enabled = false;
            //Start.Enabled = false;
            //Stop.Enabled = true;

            stop = false;
            System.Threading.Thread thread = new System.Threading.Thread(DoRecognition);
            thread.Start();
            //System.Threading.Thread.Sleep(5);
        }

        public void SimplePipeline()
        {

            //UtilMPipeline pp = null;
            MyUtilMPipeline pp = null;
            disconnected = false;

            bool sts = true;
            /* gesture */
            /* Set Source */
            //if (form.GetRecordState())
            //{
            //pp = new UtilMPipeline(form.GetRecordFile(), true);
            //pp.QueryCapture().SetFilter(form.GetCheckedDevice());
            //}
            //else if (form.GetPlaybackState())
            //{
            //pp = new UtilMPipeline(form.GetPlaybackFile(), false);
            //}
            //else
            //{
            //pp = new UtilMPipeline();
            pp = new MyUtilMPipeline();
            pp.SetForm(this);
            //   pp.QueryCapture().SetFilter(form.GetCheckedDevice());
            //}

            /* Set Module */
            
            pp.EnableGesture(/*form.GetCheckedModule()*/);
            /* end egesture*/

            /* speech recognition*/
            /* Set Audio Source */
            // pp.QueryCapture().SetFilter("Microphone Array (Creative GestureCam)"/*form.GetCheckedSource()*/);

            /* Set Module */
            pp.EnableVoiceRecognition(/*"Voice Recognition (Nuance*)"*//*form.GetCheckedModule()*/);

            /* Set Language */
            //pp.SetProfileIndex(form.GetCheckedLanguage());
            // pp.SetProfileIndex(0);

            /* Set Command/Control or Dictation */
            /*
            if (form.IsCommandControl())
            {
                string[] cmds = form.GetCommands();
                if (cmds == null)
                {
                    form.PrintStatus("No Command List. Dictation instead.");
                    pp.SetVoiceDictation();
                }
                else
                {*/

            pp.SetVoiceCommands(cmds);
            //pp.SetVoiceDictation();


            /*
                }
            }
            else
            {
                pp.SetVoiceDictation();
            }*/
            /* end speech recognition*/

            /* Initialization */
            //form.UpdateStatus("Init Started");
            bool initok = false;
            try
            {
                pp.Init();
                initok = true;
            }
            catch (Exception e)
            {
                pp = null;
                System.Diagnostics.Debug.Print("Exception - Init Failed");
            }

            if (initok)
            {
                // form.UpdateStatus("Streaming");
                pp.QueryCapture().device.SetProperty(PXCMCapture.Device.Property.PROPERTY_AUDIO_MIX_LEVEL, 0.2f);
                while (!this.stop)
                {
                    if (this.pause)
                        System.Threading.Thread.Sleep(50);
                    else
                    {
                        if (!pp.AcquireFrame(true)) break;
                        if (!DisplayDeviceConnection(pp.IsDisconnected()))
                        {
                            /* Display Results */
                            
                                PXCMGesture gesture = pp.QueryGesture();
                                PXCMImage depth = pp.QueryImage(PXCMImage.ImageType.IMAGE_TYPE_DEPTH);
                                //DisplayPicture(depth, gesture);
                                //DisplayGeoNodes(gesture);
                                DisplayGesture(gesture);
                            
                            //form.UpdatePanel();
                        }
                        pp.ReleaseFrame();
                    }
                }
                pp.Close();
               // pp.Dispose();
            }
            else
            {
                //form.UpdateStatus("Init Failed");
                sts = false;
            }

            if (!sts) System.Diagnostics.Debug.Print("Init Failed");


            //if (sts) form.UpdateStatus("Stopped");
        }

        /* audio */
        private string[] cmds = {
                    "next slide", 
                    "powerpoint next slide",
                    "powerpoint next slide please", 
             //       "powerpoint prev slide please",
                    "powerpoint previous slide",
                    "powerpoint previous slide please",
                    "nextslide",
                    //"prevslide",
                    //"prev slide",
                    //"prew slide",
                    //"prewslide",
                    "previous slide",
                    "create new slide",
                    "add new slide",
                    "add slide",
                    "delete slide",
                    "undo",
                    "un do",
                    "redo",
                    "re do",
                    "powerpoint start slideshow",
                    "powerpoint begin slideshow",
                    "powerpoint exit slideshow",
                    "start slideshow",
                    "exit slideshow"
                    };
        public void OnRecognized(PXCMVoiceRecognition.Recognition data)
        {
            if (data.label < 0)
            {
                System.Diagnostics.Debug.Print("dictation: " + data.dictation);
            }
            else
            {
                for (int i = 0; i < 4; i++)
                {
                    int label = data.nBest[i].label;
                    uint confidence = data.nBest[i].confidence;
                    if (label < 0 || confidence == 0) continue;
                    if (i == 0 && confidence > 30)
                    {
                        if (cmds[label] == "next slide" || cmds[label] == "nextslide" || cmds[label] == "powerpoint next slide" || cmds[label] == "powerpoint next slide please")
                            gotoSlide(1);

                        if (cmds[label] == "previous slide" || cmds[label] == "powerpoint previous slide" || cmds[label] == "powerpoint previous slide please")
                            gotoSlide(-1);

                        if (cmds[label] == "create new slide" || cmds[label] == "add new slide" && confidence > 40 || cmds[label] == "add slide")
                            newslide();

                        if (cmds[label] == "delete slide" && confidence > 40)
                            deleteslide();

                        if (cmds[label] == "undo" || cmds[label] == "un do")
                            undo();

                        if (cmds[label] == "redo" || cmds[label] == "re do")
                            redo();

                        if (cmds[label] == "powerpoint start slideshow" || cmds[label] == "powerpoint begin slideshow" || cmds[label] == "start slideshow")
                            startshow();

                        if (cmds[label] == "powerpoint exit slideshow" || cmds[label] == "exit slideshow")
                            endshow();

                    }
                    System.Diagnostics.Debug.Print("dictation: " + data.dictation + " > " + Convert.ToString(i) + " label: " + cmds[label] + ", confidence:" + Convert.ToString(confidence));
                }
            }
        }




        public void OnAlert(PXCMVoiceRecognition.Alert data)
        {
            System.Diagnostics.Debug.Print(AlertToString(data.label));
        }

        public string AlertToString(PXCMVoiceRecognition.Alert.Label label)
        {
            switch (label)
            {
                case PXCMVoiceRecognition.Alert.Label.LABEL_SNR_LOW: return "SNR_LOW";
                case PXCMVoiceRecognition.Alert.Label.LABEL_SPEECH_UNRECOGNIZABLE: return "SPEECH_UNRECOGNIZABLE";
                case PXCMVoiceRecognition.Alert.Label.LABEL_VOLUME_HIGH: return "VOLUME_HIGH";
                case PXCMVoiceRecognition.Alert.Label.LABEL_VOLUME_LOW: return "VOLUME_LOW";
            }
            return "UNKNOWN";
        }
        /*end audio*/

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
