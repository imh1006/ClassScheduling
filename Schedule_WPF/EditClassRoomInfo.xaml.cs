using Schedule_WPF.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.ComponentModel;
using System.Runtime.CompilerServices;
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
    /// Interaction logic for EditClassRoomInfo.xaml
    /// </summary>
    public partial class EditClassRoomInfo : INotifyPropertyChanged
    {
        private string buttonBuildingName;
        private int buttonRoomNum;
        private int buttonCapacity;
        private string buttonNotes;
        private string newLocation;
        private int newRoomNum;
        private int newSeating;
        private string newNotes;
        private bool changeClasses = false;
        public EditClassRoomInfo(string bldg, int room, int maxSeats, string notes)
        {
            InitializeComponent();
            DataContext = this;
            buttonBuildingName = bldg;
            buttonRoomNum = room;
            buttonCapacity = maxSeats;
            buttonNotes = notes;
        }
        public string CurrentBuilding
        {
            get
            {
                return buttonBuildingName;
            }
            set
            {
                buttonBuildingName = value;
                OnPropertyChanged();
            }
        }
        public int CurrentRoom
        {
            get
            {
                return buttonRoomNum;
            }
            set
            {
                buttonRoomNum = value;
                OnPropertyChanged();
            }
        }
        public int CurrentCapacity
        {
            get
            {
                return buttonCapacity;
            }
            set
            {
                buttonCapacity = value;
                OnPropertyChanged();
            }
        }
        public string CurrentNotes
        {
            get
            {
                return buttonNotes;
            }
            set
            {
                buttonNotes = value;
                OnPropertyChanged();
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;
        private void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
        private void SubmitData(object sender, RoutedEventArgs e)
        {
            if (allRequiredFields())
            {
                if (UpdateClassRoom(newLocation, newRoomNum, newSeating, newNotes) == true)
                {
                    changeClasses = true;
                    Application.Current.Resources["Set_ClassRoom_Bldg"] = newLocation;
                    Application.Current.Resources["Set_ClassRoom_Num"] = newRoomNum;
                    Application.Current.Resources["Set_ClassRoom_Seats"] = newSeating;
                    Application.Current.Resources["Set_ClassRoom_Notes"] = newNotes;
                    Application.Current.Resources["Set_ClassRoom_Success"] = true;

                }
                
                this.Close();
            }
        }

        private bool allRequiredFields()
        {
            bool success = true;
            int tmp;

            // Building name
            if (Building_Text.Text == "")
            {
                Building_Required.Visibility = Visibility.Visible;
                Building_Invalid.Visibility = Visibility.Hidden;
                success = false;
            }
            else
            {
                if (Building_Text.Text.Contains(" ") || Building_Text.Text.Length != 3)
                {
                    Building_Required.Visibility = Visibility.Hidden;
                    Building_Invalid.Visibility = Visibility.Visible;
                    success = false;
                }
                else
                {
                    Building_Invalid.Visibility = Visibility.Hidden;
                    Building_Required.Visibility = Visibility.Hidden;
                    newLocation = Building_Text.Text;
                }
            }
            // Room Number
            if (Number_Text.Text == "")
            {
                Number_Required.Visibility = Visibility.Visible;
                Number_Invalid.Visibility = Visibility.Hidden;
                success = false;
            }
            else if (!Int32.TryParse(Number_Text.Text, out tmp))
            {
                Number_Invalid.Visibility = Visibility.Visible;
                Number_Required.Visibility = Visibility.Hidden;
                success = false;
            }
            else
            {
                Number_Invalid.Visibility = Visibility.Hidden;
                Number_Required.Visibility = Visibility.Hidden;
                newRoomNum = Int32.Parse(Number_Text.Text);
            }

            // Room Capacity
            if (Seats_Text.Text != "")
            {
                if (!Int32.TryParse(Seats_Text.Text, out tmp))
                {
                    Seats_Invalid.Visibility = Visibility.Visible;
                    success = false;
                }
                else if(!(tmp >= 0))
                {
                    Seats_Invalid.Visibility = Visibility.Visible;
                    success = false;
                }
                else
                {
                    Seats_Invalid.Visibility = Visibility.Hidden;
                    newSeating = Int32.Parse(Seats_Text.Text);
                }
            }
            else
            {
                Seats_Text.Text = "0";
                newSeating = Int32.Parse(Seats_Text.Text);
            }

            // Notes
            if (Notes_Text.Text != null)
            {
                newNotes = Notes_Text.Text.ToString();
            }
            else
            {
                newNotes = "";
            }
            return success;
        }

        public bool UpdateClassRoom(string bldg, int roomNum, int maxCap, string notes)
        {
            ClassRoomList classrooms = (ClassRoomList)System.Windows.Application.Current.FindResource("ClassRoom_List_View");
            string bldgID, tempRoomLabel, inputRoomLabel, roomNotes;
            int roomID, capacity;
            inputRoomLabel = NewBuilding + "-" + NewRoom; //this is what is in the text boxes when submitted
            bool change = false;

            for (int n = 0; n < classrooms.Count; n++)
            {
                //cycles through each room in list to find the one that user originally clicked on
                bldgID = classrooms[n].Location;
                roomID = classrooms[n].RoomNum;
                capacity = classrooms[n].AvailableSeats;
                roomNotes = classrooms[n].Notes;
                change = false;

                //this is the indexed room
                tempRoomLabel = (bldgID + "-" + roomID);

                //This is if the new building and room num already exist
                if (tempRoomLabel == inputRoomLabel)
                {
                    if (maxCap != capacity) //if the two room capacities do not match
                    {
                        change = true;
                    }
                    if (!notes.Equals(roomNotes)) //if the two room notes do not match
                    {
                        change = true;
                    }
                    return change;
                }
                //if you are at the end of the list of current classrooms and have not gotten a match, add the classroom to the list
                if ((n == (classrooms.Count - 1) && (tempRoomLabel != inputRoomLabel)))
                {
                    change = true;
                    return change;
                }
            }
            return change;
        }

        public bool ChangeClasses()
        {
            return changeClasses;
        }

        public string NewBuilding
        {
            get
            {
                return newLocation;
            }
            set
            {
                newLocation = value;
            }
        }
        public int NewRoom
        {
            get
            {
                return newRoomNum;
            }
            set
            {
                newRoomNum = value;
            }
        }
        public int NewCapacity
        {
            get
            {
                return newSeating;
            }
            set
            {
                newSeating = value;
            }
        }
        public string NewNotes
        {
            get
            {
                return newNotes;
            }
            set
            {
                newNotes = value;
            }
        }
    }
    
}
