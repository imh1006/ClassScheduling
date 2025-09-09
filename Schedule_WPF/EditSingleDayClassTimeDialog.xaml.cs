using Schedule_WPF.Models;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
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
    /// Interaction logic for EditSingleDayClassTimeDialog.xaml
    /// </summary>
    public partial class EditSingleDayClassTimeDialog : Window
    {

        Timeslot[] times_MWF = { new Timeslot("08:00", "08:50", "AM"), new Timeslot("09:00", "09:50", "AM"), new Timeslot("10:00", "10:50", "AM"), new Timeslot("11:00", "11:50", "AM"), new Timeslot("12:00", "12:50", "PM"), new Timeslot("01:00", "01:50", "PM"), new Timeslot("02:00", "02:50", "PM"), new Timeslot("03:00", "03:50", "PM"), new Timeslot("04:00", "04:50", "PM"), new Timeslot("05:00", "05:50", "PM"), new Timeslot("06:00", "06:50", "PM") };
        Timeslot[] times_TR = { new Timeslot("08:00", "09:15", "AM"), new Timeslot("09:30", "10:45", "AM"), new Timeslot("11:00", "12:15", "AM"), new Timeslot("12:30", "01:45", "PM"), new Timeslot("02:00", "03:15", "PM"), new Timeslot("03:30", "04:45", "PM"), new Timeslot("05:00", "06:15", "PM") };
        ClassList classList = (ClassList)Application.Current.FindResource("Classes_List_View");
        ClassRoomList classrooms = (ClassRoomList)System.Windows.Application.Current.FindResource("ClassRoom_List_View");
        Classes targetClass;
        Timeslot selectedTime;
        string days;

        public ObservableCollection<ComboBoxItem> classroomOptions { get; set; }
        public ComboBoxItem classroom { get; set; }

        public EditSingleDayClassTimeDialog(Classes target)
        {
            InitializeComponent();
            targetClass = target;
            DataContext = this;

            classroomOptions = new ObservableCollection<ComboBoxItem>();

            for (int i = 0; i < classrooms.Count; i++)
            {
                string cont = classrooms[i].ClassID.ToString();

                if (targetClass.Classroom == classrooms[i])
                {

                    var itemOne = new ComboBoxItem { Content = cont };
                    classroom = itemOne;
                    classroomOptions.Add(itemOne);
                }
                else
                {
                    var item = new ComboBoxItem { Content = cont };
                    classroomOptions.Add(item);
                }
            }

            if (targetClass.ClassDay == "M")
            {
                TimeComboBox.SelectedIndex = 0;
            }
            if (targetClass.ClassDay == "T")
            {
                TimeComboBox.SelectedIndex = 1;
            }
            if (targetClass.ClassDay == "W")
            {
                TimeComboBox.SelectedIndex = 2;
            }
            if (targetClass.ClassDay == "R")
            {
                TimeComboBox.SelectedIndex = 3;
            }
            if (targetClass.ClassDay == "F")
            {
                TimeComboBox.SelectedIndex = 4;
            }
        }

        

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            
            if (validateTime())
            {
                days = TimeComboBox.Text;
                if (days != "Thursday")
                {
                    days = days.Substring(0, 1);
                }
                else if (days == "Thursday")
                {
                    days = "R";
                }
                targetClass.StartTime = selectedTime;
                targetClass.ClassDay = days;
                
                for (int i = 0; i < classrooms.Count(); i++)
                {
                    if (classrooms[i].ClassID == Classroom.Text)
                    {
                        targetClass.Classroom = classrooms[i];
                    }
                }


                this.Close();
            }
        }

        private bool validateTime()
        {
            bool valid = true;
            bool timeConflict = false;

            
            // if any combobox is empty = invalid
            if (StartingTime.Text == null || EndingTime.Text == null)
            {
                valid = false;
                noTime.Visibility = Visibility.Visible;
                
            }
            else
            {
                string startTime = StartingTime.Text.Substring(0, 5);
                string endTime = EndingTime.Text.Substring(0, 5);
                string startTimeFix = "";
                string endTimeFix = "";
               
                string meridian = StartingTime.Text.Substring(5);
                meridian = meridian.Trim();

                string[] Time = startTime.Split(':');
                string frontTime = Time[0];
                string backTime = Time[1];

                if (frontTime.Length == 1)
                {
                    frontTime = "0" + frontTime;
                    startTimeFix = frontTime + ":";
                }
                else
                {
                    startTimeFix = frontTime + ":";
                }

                char[] fixStartTime = backTime.ToCharArray().Where(c => !char.IsWhiteSpace(c)).ToArray();

                for (int i = 0; i < fixStartTime.Length; i++)
                {
                    startTimeFix = startTimeFix + fixStartTime[i];
                }

                string[] TimeEnd = endTime.Split(':');
                string frontTimeEnd = TimeEnd[0];
                string backTimeEnd = TimeEnd[1];

                if (frontTimeEnd.Length == 1)
                {
                    frontTimeEnd = "0" + frontTimeEnd;
                    endTimeFix = frontTimeEnd + ":";
                }
                else
                {
                    endTimeFix = frontTimeEnd + ":";
                }

                char[] fixEndTime = backTimeEnd.ToCharArray().Where(c => !char.IsWhiteSpace(c)).ToArray();

                for (int i = 0; i < fixEndTime.Length; i++)
                {
                    endTimeFix = endTimeFix + fixEndTime[i];
                }



                selectedTime = new Timeslot(startTimeFix, endTimeFix, meridian);

                days = TimeComboBox.Text;
                if (days != "Thursday")
                {
                    days = days.Substring(0, 1);
                }
                else if (days == "Thursday")
                {
                    days = "R";
                }


                // go through classlist and see if the professor is assigned at the same time and day
                for (int i = 0; i < classList.Count; i++)
                {
                    if (classList[i].Prof.FullName == targetClass.Prof.FullName)
                    {
                        if (classList[i].ClassDay == "MWF" && days == "M" || classList[i].ClassDay == "MWF" && days == "W" || classList[i].ClassDay == "MWF" && days == "F" || classList[i].ClassDay == "M" && days == "M" || classList[i].ClassDay == "W" && days == "W" || classList[i].ClassDay == "F" && days == "F" || classList[i].ClassDay == "TR" && days == "T" || classList[i].ClassDay == "TR" && days == "R" || classList[i].ClassDay == "T" && days == "T" || classList[i].ClassDay == "R" && days == "R")
                        {
                            //MessageBox.Show("Prof Hit: " + targetClass.Prof.LastName);
                            if (classList[i].StartTime.Time == startTimeFix && classList[i].TextBoxName != targetClass.TextBoxName) 
                            {
                                timeConflict = true;
                                break;
                            }
                            //MessageBox.Show("No Match\nTarget: " + selectedTime.FullTime + "\nClassList: " + classList[i].StartTime.FullTime);
                        }
                    }
                }
                // go through classlist and see if the classroom is assigned at the same time and day
                for (int i = 0; i < classList.Count(); i++)
                {
                    if (classList[i].Classroom.ClassID == Classroom.Text)
                    {
                        if (classList[i].ClassDay == "MWF" && days == "M" || classList[i].ClassDay == "MWF" && days == "W" || classList[i].ClassDay == "MWF" && days == "F" || classList[i].ClassDay == "M" && days == "M" || classList[i].ClassDay == "W" && days == "W" || classList[i].ClassDay == "F" && days == "F" || classList[i].ClassDay == "TR" && days == "T" || classList[i].ClassDay == "TR" && days == "R" || classList[i].ClassDay == "T" && days == "T" || classList[i].ClassDay == "R" && days == "R" || days == "TR" && classList[i].ClassDay == "T" || days == "TR" && classList[i].ClassDay == "R" || days == "TR" && classList[i].ClassDay == "TR" || days == "MWF" && classList[i].ClassDay == "M" || days == "MWF" && classList[i].ClassDay == "W" || days == "MWF" && classList[i].ClassDay == "F" || days == "MWF" && classList[i].ClassDay == "MWF")
                        {
                            //MessageBox.Show("Prof Hit: " + targetClass.Prof.LastName);
                            if (classList[i].StartTime.Time == startTimeFix && classList[i].TextBoxName != targetClass.TextBoxName)
                            {
                                timeConflict = true;
                                break;
                            }
                            //MessageBox.Show("No Match\nTarget: " + selectedTime.FullTime + "\nClassList: " + classList[i].StartTime.FullTime);
                        }
                    }
                }
                if (timeConflict)
                {
                    Time_Required.Visibility = Visibility.Hidden;
                    Time_Invalid.Visibility = Visibility.Visible;
                    valid = false;
                }
                else
                {
                    Time_Required.Visibility = Visibility.Hidden;
                    Time_Invalid.Visibility = Visibility.Hidden;
                }
            }
            return valid;
        }

    }
}
