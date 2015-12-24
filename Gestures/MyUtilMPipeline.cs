using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Gestures
{
public class MyUtilMPipeline : UtilMPipeline
    {
    public void SetForm(ThisAddIn form)
        {
            this.form = form;
            form.ppp = this;
        }

        public void SetProfileIndex(uint pidx)
        {
            this.pidx = pidx;
        }

        public override void OnVoiceRecognitionSetup(ref PXCMVoiceRecognition.ProfileInfo pinfo)
        {
            //form.OnVoiceRecognitionSetup(ref pinfo);
            //QueryVoiceRecognition().QueryProfile(pidx, out pinfo);
        }

        public override void OnRecognized(ref PXCMVoiceRecognition.Recognition data)
        {
            form.OnRecognized(data);
            /*
            if (data.label < 0)
            {
                form.PrintConsole(data.dictation);
            }
            else
            {
                form.ClearScores();
                for (int i = 0; i < 4; i++)
                {
                    int label = data.nBest[i].label;
                    uint confidence = data.nBest[i].confidence;
                    if (label < 0 || confidence == 0) continue;
                    form.SetScore(label,confidence);
                }
            }*/
        }

        public override void OnAlert(ref PXCMVoiceRecognition.Alert data)
        {
            form.OnAlert( data);
            /*
            form.OnRecognized(data);
            form.PrintStatus(form.AlertToString(data.label));
            */
        }

        protected uint pidx;
        protected ThisAddIn form;
    }
}
