using Schedule_WPF.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace Schedule_WPF
{
    /// <summary>
    /// Interaction logic for ChangeTimeDialog.xaml
    /// </summary>
    /// 
    public partial class ChangeTimeDialog : Window
    {
        
        public ChangeTimeDialog()
        {
            InitializeComponent();
        }

        private void Submit_Click(object sender, RoutedEventArgs e)
        {
            Application.Current.Resources["Set_ChangeTime_Success"] = false;

            if (allRequiredFields())
            {
                Application.Current.Resources["Set_ChangeTime_Success"] = true;

                string startTime = TimePicker.Text.ToString();
                int timeTable = TimeTable.SelectedIndex;
                string classBreak = Break.Text.ToString();
                int timeToChange = ChangedTimes.SelectedIndex;
                string endTime;
                string frontTimePaddedS = "";
                string backTimePaddedS = "";
                string frontTimeApply;
                string backTimeApply;
                string frontTimeApplyS;
                string backTimeApplyS;
                string meridian = "";
                int frontTimeZero = 0;
                string frontTimePadded = "";
                int backTimeZero = 0;
                string backTimePadded = "";

                int backTimePure = 0;
                string endTimePure;
                int backTimePureZero = 0;
                string backTimePurePadded = "";
                string backTimePureApply = "";


                int frontTimePure = 0;
                int frontTimePureZero = 0;
                string frontTimePurePadded = "";
                string frontTimePureApply = "";
                string meridianPure = meridian;


                string[] breakClassBreak = classBreak.Split(':');
                int breakHours = Int32.Parse(breakClassBreak[0]);
                int breakMinutes = Int32.Parse(breakClassBreak[1]);
                breakHours = breakHours * 60;

                
                int increment = Int32.Parse(Increment.Text);

                int incrementPure = increment;


                char[] removeChar = { ':', ' ' };

                string[] time = startTime.Split(removeChar);
                string[] timeS = startTime.Split(removeChar);
                int frontTimeS = int.Parse(timeS[0]);
                int backTimeS = int.Parse(timeS[1]);
                int frontTime = int.Parse(time[0]);
                int backTime = int.Parse(time[1]);
                frontTimePure = frontTime;
                backTimePure = backTime;
                meridian = time[2].ToString();
                bool meridianCheck = false;

                if (frontTime == 12)
                {
                    meridianCheck = true;
                }

                if (frontTimeS < 10 && frontTimeS.ToString().Length == 1)
                {
                    frontTimeZero = frontTimeS.ToString("D").Length + 1;
                    frontTimePaddedS = frontTimeS.ToString("D" + frontTimeZero.ToString());
                }
                if (backTimeS < 10 && backTimeS.ToString().Length == 1)
                {
                    backTimeZero = backTimeS.ToString("D").Length + 1;
                    backTimePaddedS = backTimeS.ToString("D" + backTimeZero.ToString());
                }
                if (frontTimePaddedS != "")
                {
                    frontTimeApplyS = frontTimePaddedS.ToString();

                    if (backTimePaddedS != "")
                    {

                        backTimeApplyS = backTimePaddedS.ToString();
                        startTime = frontTimeApplyS + ":" + backTimeApplyS;
                    }
                    else
                    {
                        backTimeApplyS = backTimeS.ToString();
                        startTime = frontTimeApplyS + ":" + backTimeApplyS;
                    }
                }
                else if (backTimePaddedS != "")
                {
                    frontTimeApplyS = frontTimeS.ToString();
                    backTimeApplyS = backTimePaddedS.ToString();
                    startTime = frontTimeApplyS + ":" + backTimeApplyS;
                }
                else
                {
                    frontTimeApplyS = frontTimeS.ToString();
                    backTimeApplyS = backTimeS.ToString();
                    startTime = frontTimeApplyS + ":" + backTimeApplyS;
                }



                backTimePure = backTime + incrementPure;

                while (backTimePure >= 60)
                {
                    frontTimePure = frontTimePure + 1;
                    backTimePure = backTimePure - 60;
                }

                if (backTimePure < 10 && backTimePure.ToString().Length == 1)
                {
                    backTimePureZero = backTimePure.ToString("D").Length + 1;
                    backTimePurePadded = backTimePure.ToString("D" + backTimePureZero.ToString());
                }


                while (backTimePure >= 60)
                {
                    backTimePure = backTimePure - 60;
                    frontTimePure = frontTimePure + 1;
                }
                
                while (frontTimePure > 12)
                {
                    frontTimePure = frontTimePure - 12;

                }
                if (frontTimePure < 10 && frontTimePure.ToString().Length == 1)
                {
                    frontTimePureZero = frontTimePure.ToString("D").Length + 1;
                    frontTimePurePadded = frontTimePure.ToString("D" + frontTimePureZero.ToString());
                }
                if (backTimePure < 10 && backTimePure.ToString().Length == 1)
                {
                    backTimePureZero = backTimePure.ToString("D").Length + 1;
                    backTimePurePadded = backTimePure.ToString("D" + backTimePureZero.ToString());
                }


                if (frontTimePurePadded != "")
                {
                    frontTimePureApply = frontTimePurePadded.ToString();

                    if (backTimePurePadded != "")
                    {

                        backTimePureApply = backTimePurePadded.ToString();
                        endTimePure = frontTimePureApply + ":" + backTimePureApply;
                    }
                    else
                    {
                        backTimePureApply = backTimePure.ToString();
                        endTimePure = frontTimePureApply + ":" + backTimePureApply;
                    }
                }
                else if (backTimePurePadded != "")
                {
                    frontTimePureApply = frontTimePure.ToString();
                    backTimePureApply = backTimePurePadded.ToString();
                    endTimePure = frontTimePureApply + ":" + backTimePureApply;
                }
                else
                {
                    frontTimePureApply = frontTimePure.ToString();
                    backTimePureApply = backTimePure.ToString();
                    endTimePure = frontTimePureApply + ":" + backTimePureApply;
                }



                Application.Current.Resources["Set_End_Time_Pure"] = endTimePure;

                backTime = backTime + increment + breakHours + breakMinutes; // start time of next class
                

                
                while (backTime >= 60)
                {
                    backTime = backTime - 60;
                    frontTime = frontTime + 1;
                }
                
                while (frontTime > 12)
                {
                    frontTime = frontTime - 12;
                }
                if (frontTime < 10 && frontTime.ToString().Length == 1)
                {
                    frontTimeZero = frontTime.ToString("D").Length + 1;
                    frontTimePadded = frontTime.ToString("D" + frontTimeZero.ToString());
                }
                if (backTime < 10 && backTime.ToString().Length == 1)
                {
                    backTimeZero = backTime.ToString("D").Length + 1;
                    backTimePadded = backTime.ToString("D" + backTimeZero.ToString());
                }
                if (frontTimePadded != "")
                {
                    frontTimeApply = frontTimePadded.ToString();

                    if (backTimePadded != "")
                    {

                        backTimeApply = backTimePadded.ToString();
                        endTime = frontTimeApply + ":" + backTimeApply;
                    }
                    else
                    {
                        backTimeApply = backTime.ToString();
                        endTime = frontTimeApply + ":" + backTimeApply;
                    }
                }
                else if (backTimePadded != "")
                {
                    frontTimeApply = frontTime.ToString();
                    backTimeApply = backTimePadded.ToString();
                    endTime = frontTimeApply + ":" + backTimeApply;
                }
                else
                {
                    frontTimeApply = frontTime.ToString();
                    backTimeApply = backTime.ToString();
                    endTime = frontTimeApply + ":" + backTimeApply;
                }

                Application.Current.Resources["Set_Start_Time"] = startTime;
                Application.Current.Resources["Set_End_Time"] = endTime;
                Application.Current.Resources["Set_Meridian"] = meridian;
                Application.Current.Resources["Set_TimeTable"] = timeTable;
                Application.Current.Resources["Set_TimeChange"] = timeToChange;
                this.Close();
            }

        }
        private bool allRequiredFields()
        {
            bool success = true;



            if (TimePicker.Text == null)
            {

                Start_Time_Invalid.Visibility = Visibility.Visible;
                success = false;
            }
            else
            {
                try
                {
                    char[] removeChar = { ':', ' ' };
                    string startTime = TimePicker.Text.ToString();
                    string[] time = startTime.Split(removeChar);
                    int frontTime = int.Parse(time[0]);
                    int backTime = int.Parse(time[1]);
                    

                }
                catch (Exception ex)
                {
                    Start_Time_Invalid.Visibility = Visibility.Visible;
                    success = false;
                }
            }

            
            if (Break.Text == "" || Break.Text == null)
            {
                Break_Time_Invalid.Visibility = Visibility.Visible;
                success = false;
            }

            try
            {
                int increment = Int32.Parse(Increment.Text);
            }
            catch(Exception ex)
            {
                ClassLengthWarning.Visibility = Visibility.Visible;
                success = false;
            }

            return success;
        }

        private void ChangedTimes_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

        }

        private void Increment_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

        }
    }
}