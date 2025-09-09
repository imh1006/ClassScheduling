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
    /// Interaction logic for AddProfessorDialog.xaml
    /// </summary>
    public partial class AddProfessorDialog : Window
    {
        ProfessorList professors = (ProfessorList)Application.Current.FindResource("Professor_List_View");
        
        public AddProfessorDialog()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            
            if (allRequiredFields())
            {
                string first = FirstName.Text;
                string last = LastName.Text;
                string id = ID.Text;
                string color = colorPicker.SelectedColor.ToString();

                Application.Current.Resources["Set_Prof_FN"] = first;
                Application.Current.Resources["Set_Prof_LN"] = last;
                Application.Current.Resources["Set_Prof_ID"] = id;
                Application.Current.Resources["Set_Prof_Color"] = color;
                Application.Current.Resources["Set_Prof_Success"] = true;
                
                Close();
            }
        }



        private bool allRequiredFields()
        {
            bool success = true;
            // First Name
            if (FirstName.Text == "")
            {
                FName_Duplicate.Visibility = Visibility.Hidden;
                FirstName_Required.Visibility = Visibility.Visible;
                success = false;
            }
            else
            {
                FName_Duplicate.Visibility = Visibility.Hidden;
                FirstName_Required.Visibility = Visibility.Hidden;
            }
            // Last Name
            if (LastName.Text == "")
            {
                LName_Duplicate.Visibility = Visibility.Hidden;
                LastName_Required.Visibility = Visibility.Visible;
                success = false;
            }
            else
            {
                LName_Duplicate.Visibility = Visibility.Hidden;
                LastName_Required.Visibility = Visibility.Hidden;
            }
            string newname = LastName.Text.ToUpper() + ", " + FirstName.Text.ToUpper();
            for (int i = 0; i < professors.Count; i++)
            {
                
                if (newname == professors[i].LastName.ToUpper() + ", " + professors[i].FirstName.ToUpper())
                {
                    FName_Duplicate.Visibility = Visibility.Visible;
                    LName_Duplicate.Visibility = Visibility.Visible;
                    success = false;
                    break;

                }



            }
            // SRU ID
            if (ID.Text == "")
            {
                ID_Duplicate.Visibility = Visibility.Hidden;
                ID_Required.Visibility = Visibility.Visible;
                ID_Invalid.Visibility = Visibility.Hidden;
                success = false;
            }
            else
            {
                if (ID.Text.Length != 9 || ID.Text.Substring(0, 2) != "A0")
                {
                    ID_Duplicate.Visibility = Visibility.Hidden;
                    ID_Invalid.Visibility = Visibility.Visible;
                    ID_Required.Visibility = Visibility.Hidden;
                    success = false;
                }
                else
                {

                    for (int i = 0; i < professors.Count; i++)
                    {
                        
                        if (ID.Text == professors[i].SRUID)
                        {
                            ID_Duplicate.Visibility = Visibility.Visible;
                            ID_Invalid.Visibility = Visibility.Hidden;
                            ID_Required.Visibility = Visibility.Hidden;
                            success = false;
                            break;

                        }
                        else
                        {
                            ID_Duplicate.Visibility = Visibility.Hidden;
                            ID_Invalid.Visibility = Visibility.Hidden;
                            ID_Required.Visibility = Visibility.Hidden;
                        }


                    }
                    
                }
            }
            // Color
            if (colorPicker.SelectedColor.ToString() == "")
            {
                Color_Required.Visibility = Visibility.Visible;
                Color_Invalid.Visibility = Visibility.Hidden;
                success = false;
            }
            else
            {
                RGB_Color tempColor = new RGB_Color(colorPicker.SelectedColor.ToString());
                if (isColorTaken(tempColor))
                {
                    Color_Invalid.Visibility = Visibility.Visible;
                    Color_Required.Visibility = Visibility.Hidden;
                    success = false;
                }
                else
                {
                    Color_Invalid.Visibility = Visibility.Hidden;
                    Color_Required.Visibility = Visibility.Hidden;
                }
            }

            return success;
        }

        public bool isColorTaken(RGB_Color color)
        {
            for (int i = 0; i < professors.Count; i++)
            {
                if (withinColorRange(color, professors[i].profRGB))
                {
                    return true;
                }
            }
            return false;
        }
        public bool withinColorRange(RGB_Color c1, RGB_Color c2)
        {
            int threshold = 65;
            if (Math.Abs(c1.R - c2.R) <= threshold && Math.Abs(c1.G - c2.G) <= threshold && Math.Abs(c1.B - c2.B) <= threshold)
            {
                return true;
            }
            return false;
        }
    }
}
