using System;
using System.Collections.Generic;
using System.Linq;
using System.Media;
using System.Text;
using System.Threading.Tasks;

namespace CharacterConverttoCharacterPics
{
    public class warnings
    {
        public static void playBeep() { SystemSounds.Beep.Play(); }//https://blog.kkbruce.net/2019/03/csharpformusicplay.html#.YHx9O-gzai4
        //https://analystcave.com/vba-status-bar-progress-bar-sounds-emails-alerts-vba/#:~:text=The%20VBA%20Status%20Bar%20is%20a%20panel%20that,Bar%20we%20need%20to%20Enable%20it%20using%20Application.DisplayStatusBar%3A
        public static void playSound()
        {//Public Declare Function sndPlaySound32 Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
            System.Media.SoundPlayer splay = new SoundPlayer(
                @"G:\我的雲端硬碟\修行\佛號\南無阿彌陀佛_2014_(印光大師版_六字四音_女聲唱頌_旋律優美) (mp3cut.net).wav");
            try
            {
                splay.Play();
                //播放聲音、音效、音樂
                //sndPlaySound32 "c:\Windows\Media\Alarm08.wav", &H0 '"C:\Windows\Media\Chimes.wav", &H0
                //        sndPlaySound32 "C:\Program Files (x86)\Microsoft Office\Office16\MEDIA\LYNC_ringtone2.wav", &H0
                //       sndPlaySound32 "C:\Program Files (x86)\Microsoft Office\Office16\MEDIA\LYNC_fsringing.wav", &H0
            }
            catch (Exception)
            {

                //throw;
            }
        }

    }
}
