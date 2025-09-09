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
    /// Interaction logic for EditClassDialog.xaml
    /// </summary>
    public partial class EditClassDialog : Window
    {
        Classes targetClass = null;
        string originalCRN = "";
        bool originalOnline;
        bool originalAssigned;
        Professors oldProfessor;

        public EditClassDialog(Classes _class)
        {
            InitializeComponent();
            Application.Current.Resources["Set_Class_Success"] = false;
            Application.Current.Resources["Edit_Class_Check"] = false;
            targetClass = _class;
            originalCRN = _class.CRN;
            originalOnline = _class.Online;
            originalAssigned = _class.isAssigned;
            ProfessorList profs = (ProfessorList)Application.Current.FindResource("Professor_List_View");
            Prof_Text.ItemsSource = profs;
            ClassList classList = (ClassList)System.Windows.Application.Current.FindResource("Classes_List_View");
            ClassList unassignedClasses = (ClassList)System.Windows.Application.Current.FindResource("Unassigned_Classes_List_View");
            ClassList onlineClasses = (ClassList)System.Windows.Application.Current.FindResource("Online_Classes_List_View");
            ClassList appointmentClasses = (ClassList)System.Windows.Application.Current.FindResource("Appointment_Classes_List_View");
            ClassList appointment2Classes = (ClassList)System.Windows.Application.Current.FindResource("Appointment2_Classes_List_View");
            ClassList deletedClasses = (ClassList)System.Windows.Application.Current.FindResource("Deleted_Classes_List_View");


            oldProfessor = _class.Prof;

            // Initialize fields with available data from class
            Classes c1 = _class;
            Term_Text.Text = _class.Term.ToString();
            Session_Text.Text = _class.Session.ToString();
            CRN_Text.Text = _class.CRN.ToString();
            Dept_Text.Text = _class.DeptName;
            ClassNum_Text.Text = _class.ClassNumber.ToString();
            Section_Text.Text = _class.SectionNumber.ToString();
            Name_Text.Text = _class.ClassName;
            Credits_Text.Text = _class.Credits.ToString();
            Waitlist_Text.Text = _class.Waitlist.ToString();
            Crosslist_Text.Text = _class.Crosslist.ToString();
            MaxSeats_Text.Text = _class.MaxSeats.ToString();
            ProjSeats_Text.Text= _class.ProjSeats.ToString();
            Enrolled_Text.Text = _class.Enrolled.ToString();
            StartDatePicker.Text = _class.StartDate.ToString();
            EndDatePicker.Text = _class.EndDate.ToString();
            RoomCap_Text.Text = _class.Classroom.AvailableSeats.ToString();
            Building_Text.Text = _class.Classroom.Location.ToString();
            Room_Text.Text = _class.Classroom.RoomNum.ToString();
            Days_Text.Text = _class.ClassDay.ToString();


            string start = _class.StartTime.Start;
            string end = _class.StartTime.End;
            string meridian = "";

            if (start != "-- ")
            {
                string checkMeridian = start.Substring(0, 2);
                string checkMeridianEnd = end.Substring(0, 2);
                int starting = Int32.Parse(checkMeridian);
                int ending = Int32.Parse(checkMeridianEnd);

                if (starting < 12 && ending >= 12)
                {
                    end = end.Substring(0, end.Length - 2);
                    meridian = "PM";
                }
                else
                {
                    meridian = _class.StartTime.Meridian.ToString();
                }
            }

            Times_Text.Text = _class.StartTime.Start.ToString() + " " + "- " + _class.StartTime.EndTime.ToString() + " " + meridian;


            int profIndex;
            for (profIndex = 0; profIndex < profs.Count; profIndex++)
            {
                if (profs[profIndex].FullName == _class.Prof.FullName)
                {
                    break;
                }
            }
            Prof_Text.SelectedIndex = profIndex;
            if (_class.Online)
            {
                Online_Box.IsChecked = true;
            }
            else if (_class.isAppointment)
            {
                if (_class.Classroom.Location == "APPT")
                {
                    Appointment_Box.IsChecked = true;
                }
                else
                {
                    Appointment2_Box.IsChecked = true;
                }
            }
            else
            {
                InClass_Box.IsChecked = true;
            }
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            if (allRequiredFields() && targetClass != null)
            {
                if (oldProfessor.FullName != ((Professors)Prof_Text.SelectedItem).FullName && !((bool)Online_Box.IsChecked) && targetClass.isAssigned)
                {
                    //MessageBox.Show("Here!"); // just flag it to main
                    Application.Current.Resources["Edit_Class_Check"] = true;
                }
                Application.Current.Resources["Set_Class_Term"] = Int32.Parse(Term_Text.Text.ToString());
                Application.Current.Resources["Set_Class_Session"] = Int32.Parse(Session_Text.Text.ToString());
                Application.Current.Resources["Set_Class_Success"] = true;
                Application.Current.Resources["Set_Class_CRN"] = CRN_Text.Text;
                Application.Current.Resources["Set_Class_Dept"] = Dept_Text.Text;
                Application.Current.Resources["Set_Class_Number"] = Int32.Parse(ClassNum_Text.Text.ToString());
                Application.Current.Resources["Set_Class_Section"] = Int32.Parse(Section_Text.Text.ToString());
                Application.Current.Resources["Set_Class_Name"] = Name_Text.Text;
                Application.Current.Resources["Set_Class_Credits"] = Int32.Parse(Credits_Text.Text.ToString());
                Application.Current.Resources["Set_Class_Professor"] = ((Professors)Prof_Text.SelectedItem).SRUID;
                Application.Current.Resources["Set_Class_Online"] = (bool)Online_Box.IsChecked;
                Application.Current.Resources["Set_Class_Appointment"] = (bool)Appointment_Box.IsChecked;
                Application.Current.Resources["Set_Class_Appointment2"] = (bool)Appointment2_Box.IsChecked;
                Application.Current.Resources["Set_Class_Enrolled"] = Int32.Parse(Enrolled_Text.Text.ToString());
                Application.Current.Resources["Set_Class_Waitlist"] = Int32.Parse(Waitlist_Text.Text.ToString());
                Application.Current.Resources["Set_Class_Crosslist"] = Crosslist_Text.Text.ToString();
                Application.Current.Resources["Set_Class_maxSeats"] = Int32.Parse(MaxSeats_Text.Text.ToString());
                Application.Current.Resources["Set_Class_projSeats"] = Int32.Parse(ProjSeats_Text.Text.ToString());
                Application.Current.Resources["Set_Class_StartDate"] = StartDatePicker.Text;
                Application.Current.Resources["Set_Class_EndDate"] = EndDatePicker.Text;

                // Close the window
                this.Close();
            }
        }

        private bool allRequiredFields()
        {
            bool success = true;
            int tmp;

            if (Prof_Text.Text == "")
            {
                Prof_image.Visibility = Visibility.Visible;
                success = false;
            }

            //Class Term
            if (Term_Text.Text == "")
            {
                Term_Required.Visibility = Visibility.Visible;
                Term_Invalid.Visibility = Visibility.Hidden;

                success = false;
            }
            else if (!Int32.TryParse(Term_Text.Text, out tmp))
            {
                Term_Required.Visibility = Visibility.Hidden;
                Term_Invalid.Visibility = Visibility.Visible;
                success = false;
            }
            else
            {
                if (tmp > 0)
                {
                    Term_Invalid.Visibility = Visibility.Hidden;
                    Session_Required.Visibility = Visibility.Hidden;
                }
                else
                {
                    Term_Invalid.Visibility = Visibility.Visible;
                    Term_Required.Visibility = Visibility.Hidden;
                    success = false;
                }
            }

            //Class Session
            if (Session_Text.Text == "")
            {
                Session_Required.Visibility = Visibility.Visible;
                Session_Invalid.Visibility = Visibility.Hidden;

                success = false;
            }
            else if (!Int32.TryParse(Session_Text.Text, out tmp))
            {
                Session_Required.Visibility = Visibility.Hidden;
                Session_Invalid.Visibility = Visibility.Visible;
                success = false;
            }
            else
            {
                if (tmp > 0)
                {
                    Session_Invalid.Visibility = Visibility.Hidden;
                    Session_Required.Visibility = Visibility.Hidden;
                }
                else
                {
                    Session_Invalid.Visibility = Visibility.Visible;
                    Session_Required.Visibility = Visibility.Hidden;
                    success = false;
                }
            }




            // Class CRN
            if (CRN_Text.Text == "")
            {
                CRN_Required.Visibility = Visibility.Visible;
                CRN_Invalid.Visibility = Visibility.Hidden;
                success = false;
            }
            else if (!Int32.TryParse(CRN_Text.Text, out tmp))
            {
                CRN_Text.Text = CRN_Text.Text.ToUpper();
                if (CRN_Text.Text != "NEW")
                {
                    CRN_Invalid.Visibility = Visibility.Visible;
                    CRN_Required.Visibility = Visibility.Hidden;
                    success = false;
                }
            }
            else
            {
                CRN_Invalid.Visibility = Visibility.Hidden;
                CRN_Required.Visibility = Visibility.Hidden;
            }
            // Department Name
            if (Dept_Text.Text == "")
            {
                Dept_Required.Visibility = Visibility.Visible;
                Dept_Invalid.Visibility = Visibility.Hidden;
                success = false;
            }
            else
            {
                if (Dept_Text.Text.Length != 4)
                {
                    Dept_Required.Visibility = Visibility.Hidden;
                    Dept_Invalid.Visibility = Visibility.Visible;
                    success = false;
                }
                else
                {
                    Dept_Required.Visibility = Visibility.Hidden;
                    Dept_Invalid.Visibility = Visibility.Hidden;
                    Dept_Text.Text = Dept_Text.Text.ToUpper();
                }
            }
            // Class Number
            if (ClassNum_Text.Text == "")
            {
                Number_Required.Visibility = Visibility.Visible;
                Number_Invalid.Visibility = Visibility.Hidden;

                success = false;
            }
            else if (!Int32.TryParse(ClassNum_Text.Text, out tmp))
            {
                Number_Invalid.Visibility = Visibility.Visible;
                Number_Required.Visibility = Visibility.Hidden;
                success = false;
            }
            else
            {
                if (tmp > 99 && tmp < 1000)
                {
                    Number_Invalid.Visibility = Visibility.Hidden;
                    Number_Required.Visibility = Visibility.Hidden;
                }
                else
                {
                    Number_Invalid.Visibility = Visibility.Visible;
                    Number_Required.Visibility = Visibility.Hidden;
                    success = false;
                }
            }

            ////// Class Section
            if (Section_Text.Text == "")
            {
                Section_Required.Visibility = Visibility.Visible;
                Section_Invalid.Visibility = Visibility.Hidden;
                success = false;
            }
            else if (!Int32.TryParse(Section_Text.Text, out tmp))
            {
                Section_Invalid.Visibility = Visibility.Visible;
                Section_Required.Visibility = Visibility.Hidden;
                success = false;
            }
            else
            {
                int sectionNum = Int32.Parse(Section_Text.Text);
                if (sectionNum > 0)
                {
                    Section_Invalid.Visibility = Visibility.Hidden;
                    Section_Required.Visibility = Visibility.Hidden;
                }
                else
                {
                    Section_Invalid.Visibility = Visibility.Visible;
                    Section_Required.Visibility = Visibility.Hidden;
                    success = false;
                }
            }
            // Class Name
            if (Name_Text.Text == "")
            {
                Name_Required.Visibility = Visibility.Visible;
                success = false;
            }
            else
            {
                Name_Required.Visibility = Visibility.Hidden;
            }


            // Class Credits
            if (Credits_Text.Text == "")
            {
                Credits_Required.Visibility = Visibility.Visible;
                Credits_Invalid.Visibility = Visibility.Hidden;
                success = false;
            }
            else if (!Int32.TryParse(Credits_Text.Text, out tmp))
            {
                Credits_Invalid.Visibility = Visibility.Visible;
                Credits_Required.Visibility = Visibility.Hidden;
                success = false;
            }
            else
            {
                if (tmp > 0 && tmp < 15)
                {
                    Credits_Invalid.Visibility = Visibility.Hidden;
                    Credits_Required.Visibility = Visibility.Hidden;
                }
                else
                {
                    Credits_Invalid.Visibility = Visibility.Visible;
                    Credits_Required.Visibility = Visibility.Hidden;
                    success = false;
                }
            }

            // Waitlist
            if (Waitlist_Text.Text == "")
            {
                Waitlist_Required.Visibility = Visibility.Visible;
                Waitlist_Invalid.Visibility = Visibility.Hidden;
                success = false;
            }
            else if (!Int32.TryParse(Waitlist_Text.Text, out tmp))
            {
                Waitlist_Required.Visibility = Visibility.Hidden;
                Waitlist_Invalid.Visibility = Visibility.Visible;
                success = false;
            }
            else
            {
                if (tmp >= 0)
                {
                    Waitlist_Required.Visibility = Visibility.Hidden;
                    Waitlist_Invalid.Visibility = Visibility.Hidden;
                }
                else
                {
                    Waitlist_Required.Visibility = Visibility.Hidden;
                    Waitlist_Invalid.Visibility = Visibility.Visible;
                    success = false;
                }
            }
            // Proj Seats
            if (!Int32.TryParse(ProjSeats_Text.Text, out tmp))
            {
                ProjSeats_Invalid.Visibility = Visibility.Visible;
                success = false;
            }
            else
            {
                if (tmp >= 0)
                {
                    ProjSeats_Invalid.Visibility = Visibility.Hidden;
                    if (Int32.Parse(RoomCap_Text.Text) == 0)
                    {
                        ProjSeats_Invalid.Visibility = Visibility.Hidden;
                    }
                    else if (tmp <= Int32.Parse(RoomCap_Text.Text))
                    {
                        ProjSeats_Invalid.Visibility = Visibility.Hidden;
                    }
                    else
                    {
                        ProjSeats_Invalid.Visibility = Visibility.Visible;
                        success = false;
                    }
                }
                else
                {
                    ProjSeats_Invalid.Visibility = Visibility.Visible;
                    success = false;
                }
            }

            // Max Seats
            if (!Int32.TryParse(MaxSeats_Text.Text, out tmp))
            {
                MaxSeats_Invalid.Visibility = Visibility.Visible;
                success = false;
            }
            else
            {
                if (tmp >= 0)
                {
                    MaxSeats_Invalid.Visibility = Visibility.Hidden;
                    if (Int32.Parse(RoomCap_Text.Text) == 0)
                    {
                        MaxSeats_Invalid.Visibility = Visibility.Hidden;
                    }
                    else if (tmp <= Int32.Parse(RoomCap_Text.Text))
                    {
                        MaxSeats_Invalid.Visibility = Visibility.Hidden;
                    }
                    else
                    {
                        MaxSeats_Invalid.Visibility = Visibility.Visible;
                        success = false;
                    }
                }
                else
                {
                    MaxSeats_Invalid.Visibility = Visibility.Visible;
                    success = false;
                }
            }

            // Enrolled
            if (Enrolled_Text.Text == "")
            {
                Enrolled_Required.Visibility = Visibility.Visible;
                Enrolled_Invalid.Visibility = Visibility.Hidden;
                success = false;
            }
            else if (!Int32.TryParse(Enrolled_Text.Text, out tmp))
            {
                Enrolled_Required.Visibility = Visibility.Visible;
                Enrolled_Invalid.Visibility = Visibility.Hidden;
                success = false;
            }
            else
            {
                if (tmp >= 0)
                {
                    Enrolled_Required.Visibility = Visibility.Hidden;
                    Enrolled_Invalid.Visibility = Visibility.Hidden;
                }
                else
                {
                    Enrolled_Required.Visibility = Visibility.Visible;
                    Enrolled_Invalid.Visibility = Visibility.Hidden;
                    success = false;
                }
            }

            return success;
        }

        private void Add_Prof_Click(object sender, RoutedEventArgs e)
        {
            AddProfessorDialog addProfDialog = new AddProfessorDialog();
            addProfDialog.Owner = this;
            addProfDialog.ShowDialog();

            ProfessorList professors = (ProfessorList)Application.Current.Resources["Professor_List_View"];

            if ((bool)Application.Current.Resources["Set_Prof_Success"])
            {
                // Get data about new professor
                string fName = (string)Application.Current.Resources["Set_Prof_FN"];
                string lName = (string)Application.Current.Resources["Set_Prof_LN"];
                string id = (string)Application.Current.Resources["Set_Prof_ID"];
                string colorString = (string)Application.Current.Resources["Set_Prof_Color"];
                // Create new professor object and add to professor list
                Professors prof = new Professors(fName, lName, id);
                prof.profRGB = new RGB_Color(colorString);
                professors.Add(prof);
                // Set the new professor as the comboBox selected value
                Prof_Text.SelectedIndex = (professors.Count - 1);
                // Reset Success flag for the addProfessorDialog
                Application.Current.Resources["Set_Prof_Success"] = false;
            }
        }
        private void CorrectDates(object sender, SelectionChangedEventArgs e)
        {

            if (EndDatePicker.Text == null)
            {
                EndDatePicker.SelectedDate = StartDatePicker.SelectedDate.Value.AddDays(30);
                Application.Current.Resources["Set_Class_EndDate"] = EndDatePicker.Text;
            }
            if (StartDatePicker.Text == null)
            {
                StartDatePicker.SelectedDate = EndDatePicker.SelectedDate.Value.AddDays(-30);
                Application.Current.Resources["Set_Class_StartDate"] = StartDatePicker.Text;
            }

            var test1 = EndDatePicker.Text.ToString().CompareTo(StartDatePicker.Text.ToString());

            if ((test1 == 0) || (test1 == -1))
            {
                EndDatePicker.SelectedDate = StartDatePicker.SelectedDate.Value.AddDays(30);
                EndDate_Invalid.Visibility = Visibility.Visible;
                Application.Current.Resources["Set_Class_EndDate"] = EndDatePicker.Text;
                return;
            }
            else
            {
                EndDate_Invalid.Visibility = Visibility.Hidden;
                Application.Current.Resources["Set_Class_EndDate"] = EndDatePicker.Text;
            }



            var test2 = StartDatePicker.Text.ToString().CompareTo(EndDatePicker.Text.ToString());

            if ((test2 == 0) || (test2 == 1))
            {
                StartDatePicker.SelectedDate = EndDatePicker.SelectedDate.Value.AddDays(-30);
                EndDate_Invalid.Visibility = Visibility.Visible;
                Application.Current.Resources["Set_Class_StartDate"] = StartDatePicker.Text;
                return;
            }
            else
            {
                EndDate_Invalid.Visibility = Visibility.Hidden;
                Application.Current.Resources["Set_Class_StartDate"] = StartDatePicker.Text;
            }
        }
    }
}
