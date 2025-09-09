using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace Schedule_WPF.Models
{
    public class Classes : INotifyPropertyChanged
    {
        private string _DeptName;
        private int _ClassNumber;
        private int _SectionNumber;
        private string _ClassName;
        private string _ClassDay;
        private Timeslot _StartTime;
        private int _SeatsTaken;
        private int _Credits;
        private string _CRN;
        private bool _Online;
        private bool _isAssigned;
        private bool _isAppointment;
        private bool _hasChanged;
        private bool _excludeCredits;
        private string _Notes;
        private string _SectionNotes;
        private Professors _Prof;
        private ClassRoom _Classroom;
        private int _preferenceLevel;
        private string _preferenceMessage;
        private string _preferenceCode;
        private int _maxSeats;
        private int _projSeats;
        private List<bool> _changedData;
        private int _Session;
        private int _Term;
        private int _Waitlist;
        private string _Crosslist;
        private int _Enrolled;
        private string _StartDate;
        private string _EndDate;
        private string _Building;
        private string _RoomCap;
        private string _Room;

        public Classes()
        {
            ChangedData = new List<bool>();
            for (int i = 0; i < 33; i++)
            {
                ChangedData.Add(false);
            }
            Term = 0;
            Session = 0;
            CRN = "";
            DeptName = "";
            ClassNumber = 0;
            SectionNumber = 0;
            ClassName = "";
            Credits = 0;
            ClassDay = "";
            StartTime = new Timeslot();
            SeatsTaken = 0;
            Classroom = new ClassRoom();
            Prof = new Professors();
            isAssigned = false;
            Online = false;
            isAppointment = false;
            hasChanged = false;
            excludeCredits = false;
            ExtraData = new List<string>();
            Notes = "";
            SectionNotes = "";
            isCrossFirst = false;
            MaxSeats = 0;
            ProjSeats = 0;
            Enrolled = 0;
            Waitlist = 0;
            Crosslist = "";
            StartDate = "0";
            EndDate = "0";
            Building = "";
            Room = "";
            RoomCap = "";

            for (int i = 0; i < 33; i++)
            {
                ChangedData[i] = false;
            }
        }

        public Classes(int term, int session, string crn, string deptName, int classNum, int secNum, string className, int credits,
            string classDay, Timeslot startTime, int seatsTaken, ClassRoom classroom, Professors professor, bool online, bool appointment, bool changed, string sectionNotes, string notes, List<string> extras, int maxseats, int projseats, int enrolled, int waitlist, string crosslist, string startDate, string endDate, string building, string room, string roomCap)
        {
            ChangedData = new List<bool>();
            for (int i = 0; i < 33; i++)
            {
                ChangedData.Add(false);
            }
            Term = term;
            Session = session;
            CRN = crn;
            DeptName = deptName;
            ClassNumber = classNum;
            SectionNumber = secNum;
            ClassName = className;
            Credits = credits;
            ClassDay = classDay;
            StartTime = startTime;
            SeatsTaken = seatsTaken;
            Classroom = classroom;
            Prof = professor;
            isAssigned = false;
            Online = online;
            isAppointment = appointment;
            hasChanged = changed;
            ExtraData = extras;

            if (extras.Count == 0)
            {
                // Initialize
                for (int i = 0; i < 15; i++)
                {
                    extras.Add("");
                }
            }
            Notes = notes;
            SectionNotes = sectionNotes;
            PreferenceLevel = 0;
            PreferenceMessage = "";
            PreferenceCode = "";
            isCrossFirst = false;
            MaxSeats = maxseats;
            ProjSeats = projseats;
            Enrolled = enrolled;
            Waitlist = waitlist;
            Crosslist = crosslist;
            StartDate = startDate;
            EndDate = endDate;
            Building = building;
            Room = room;
            RoomCap = roomCap;


            for (int i = 0; i < 33; i++)
            {
                ChangedData[i] = false;
            }
        }

        public Classes DeepCopy()
        {
            List<string> extraCopy = new List<string>();
            for (int i = 0; i < ExtraData.Count; i++)
            {
                extraCopy.Add(ExtraData[i]);
            }
            Classes deepcopy = new Classes(Term, Session, CRN, DeptName, ClassNumber, SectionNumber, ClassName, Credits, ClassDay, StartTime, SeatsTaken, Classroom, Prof, Online, isAppointment, hasChanged, SectionNotes, Notes, extraCopy, MaxSeats, ProjSeats, Enrolled, Waitlist, Crosslist, StartDate, EndDate, Building, Room, RoomCap);
            List<bool> changedCopy = new List<bool>();
            for (int i = 0; i < ChangedData.Count; i++)
            {
                changedCopy.Add(ChangedData[i]);
            }
            deepcopy.ChangedData = changedCopy;
            return deepcopy;
        }

        public byte[] Serialize()
        {
            using (MemoryStream m = new MemoryStream())
            {
                using (BinaryWriter writer = new BinaryWriter(m))
                {
                    writer.Write(Term + Session + CRN + DeptName + ClassNumber + SectionNumber + ClassName + ClassDay + StartTime.FullTime + SeatsTaken + Credits + Online + isAssigned + isAppointment + excludeCredits + hasChanged + Prof.FullName + Classroom.ClassID + Notes + SectionNotes + isHidden + MaxSeats + ProjSeats + Enrolled + Waitlist + Crosslist + StartDate + EndDate);
                }
                return m.ToArray();
            }
        }
        public int Term { get { return _Term; } set { _Term = value; OnPropertyChanged("Term"); ChangedData[0] = true; } }
        public int Session { get { return _Session; } set { _Session = value; OnPropertyChanged("Session"); ChangedData[1] = true; } }
        public string DeptName { get { return _DeptName; } set { _DeptName = value; OnPropertyChanged("DeptName"); ChangedData[2] = true; } }
        public int ClassNumber { get { return _ClassNumber; } set { _ClassNumber = value; OnPropertyChanged("ClassNumber"); ChangedData[3] = true; } }
        public int SectionNumber { get { return _SectionNumber; } set { _SectionNumber = value; OnPropertyChanged("SectionNumber"); ChangedData[4] = true; } }
        public string ClassName { get { return _ClassName; } set { _ClassName = value; OnPropertyChanged("ClassName"); ChangedData[6] = true; } }
        public string ClassDay { get { return _ClassDay; } set { _ClassDay = value; OnPropertyChanged("ClassDay"); ChangedData[15] = true; } }
        public Timeslot StartTime { get { return _StartTime; } set { _StartTime = value; OnPropertyChanged("StartTime"); ChangedData[16] = true; ChangedData[17] = true; } }
        public int SeatsTaken { get { return _SeatsTaken; } set { _SeatsTaken = value; OnPropertyChanged("SeatsTaken"); ChangedData[12] = true; } }
        public int Credits { get { return _Credits; } set { _Credits = value; OnPropertyChanged("Credits"); ChangedData[8] = true; } }
        public string CRN { get { return _CRN; } set { _CRN = value; OnPropertyChanged("CRN"); ChangedData[5] = true; } }
        public bool Online { get { return _Online; } set { _Online = value; OnPropertyChanged("Online"); } }
        public bool isAssigned { get { return _isAssigned; } set { _isAssigned = value; OnPropertyChanged("isAssigned"); } }
        public bool isAppointment { get { return _isAppointment; } set { _isAppointment = value; excludeCredits = value; OnPropertyChanged("isAppointment"); } }
        public bool hasChanged { get { return _hasChanged; } set { _hasChanged = value; OnPropertyChanged("hasChanged"); } }
        public bool excludeCredits { get { return _excludeCredits; } set { _excludeCredits = value; OnPropertyChanged("excludeCredits"); } }
        public Professors Prof { get { return _Prof; } set { _Prof = value; OnPropertyChanged("Prof"); ChangedData[21] = true; ChangedData[22] = true; } }
        public ClassRoom Classroom { get { return _Classroom; } set { _Classroom = value; OnPropertyChanged("Classroom"); ChangedData[19] = true; ChangedData[20] = true; ChangedData[22] = true; } }
        public string TextBoxName { get { return DeptName + " " + ClassNumber + " [" + SectionNumber + "] " + PreferenceCodeFormatted; } }
        public int SeatsLeft { get { return Classroom.AvailableSeats - SeatsTaken; } }
        public string ClassID { get { return CRN + ClassName + SectionNumber + ClassNumber; } }
        public string ToolTipText { get { return "Name: " + ClassName + "\nProfessor: " + Prof.FullName + "\nTime: " + StartTime.FullTime + PreferenceMessageFormatted + HiddenMessage; } }
        public string Notes { get { return _Notes; } set { _Notes = value; OnPropertyChanged("Notes"); ChangedData[32] = true; } }
        public string SectionNotes { get { return _SectionNotes; } set { _SectionNotes = value; OnPropertyChanged("SectionNotes"); ChangedData[31] = true; } }
        public int PreferenceLevel { get { return _preferenceLevel; } set { _preferenceLevel = value; OnPropertyChanged("PreferenceLevel"); } }
        public string PreferenceMessage { get { return _preferenceMessage; } set { _preferenceMessage = value; OnPropertyChanged("PreferenceMessage"); } }
        public string PreferenceMessageFormatted { get { if (PreferenceLevel < 0) { return "\nPreference: " + PreferenceMessage; } else { return ""; } } }
        public string PreferenceCode { get { return _preferenceCode; } set { _preferenceCode = value; OnPropertyChanged("PreferenceCode"); } }
        public string PreferenceCodeFormatted { get { if (PreferenceLevel < 0) { return _preferenceCode; } else { return ""; } } }
        public bool isCrossListed { get { if (ExtraData[1] != "") { return true; } else { return false; } } } //CrossList
        //public string CrossListCode { get { return ExtraData[0]; } }
        public bool isHidden { get { if (MaxSeats == 0) { return true; } else { return false; } } }
        public string HiddenMessage { get { if (isHidden) { return "\n[ HIDDEN ]"; } else { return ""; } } }
        public int MaxSeats { get { return _maxSeats; } set { _maxSeats = value; OnPropertyChanged("MaxSeats"); ChangedData[9] = true; } }
        public int ProjSeats { get { return _projSeats; } set { _projSeats = value; OnPropertyChanged("ProjSeats"); ChangedData[11] = true; } }
        public List<string> ExtraData { get; set; }
        public bool isCrossFirst { get; set; }
        public List<bool> ChangedData { get { return _changedData; } set { _changedData = value; } }
        public int Enrolled { get { return _Enrolled; } set { _Enrolled = value; OnPropertyChanged("Enrolled"); ChangedData[12] = true;  } }
        public int Waitlist { get { return _Waitlist; } set { _Waitlist = value; OnPropertyChanged("Waitlist"); ChangedData[13] = true; } }
        public string Crosslist { get { return _Crosslist; } set { _Crosslist = value; OnPropertyChanged("Crosslist"); ChangedData[7] = true; } }
        public string StartDate { get { return _StartDate; } set { _StartDate = value; OnPropertyChanged("StartDate"); ChangedData[14] = true; } }
        public string EndDate { get { return _EndDate; } set { _EndDate = value; OnPropertyChanged("EndDate"); ChangedData[18] = true; } }
        public string Building { get { return _Building; } set { _Building = value;} }
        public string Room { get { return _Room; } set { _Room = value; } }
        public string RoomCap { get { return _RoomCap; } set { _RoomCap = value; } }



        public string getSectionString()
        {
            string output = "";
            if (SectionNumber > 0)
            {
                if (SectionNumber < 10)
                {
                    output = output + "0" + SectionNumber;
                }
                else
                {
                    output = output + SectionNumber;
                }
            }
            return output;
        }

        public event PropertyChangedEventHandler PropertyChanged;

        public void OnPropertyChanged(string propertyName)
        {
            if (PropertyChanged != null)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
            }
        }
    }
}
