using System;
using System.Data;
using System.Collections.Generic;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Xml.Serialization;
using Schedule_WPF.Properties;
using System.ComponentModel;
using Schedule_WPF.Models;
using Microsoft.Win32;
using ClosedXML.Excel;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Windows.Data;
using System.Globalization;
using System.Xml;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Word;
using System.Text.RegularExpressions;
using System.Windows.Forms;
//using Xceed.Wpf.Toolkit;



namespace Schedule_WPF
{
    /// <summary>
    /// Main Window of the program
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        ////////////// GLOBAL VARIABLES ////////////////

        Timeslot[] times_MWF = { new Timeslot("08:00", "08:50", "AM"), new Timeslot("09:00", "09:50", "AM"), new Timeslot("10:00", "10:50", "AM"), new Timeslot("11:00", "11:50", "AM"), new Timeslot("12:00", "12:50", "PM"), new Timeslot("01:00", "01:50", "PM"), new Timeslot("02:00", "02:50", "PM"), new Timeslot("03:00", "03:50", "PM"), new Timeslot("04:00", "04:50", "PM"), new Timeslot("05:00", "05:50", "PM"), new Timeslot("06:00", "06:50", "PM"), new Timeslot("07:00", "07:50", "PM"), new Timeslot("08:00", "08:50", "PM") };
        Timeslot[] times_TR = { new Timeslot("08:00", "09:15", "AM"), new Timeslot("09:30", "10:45", "AM"), new Timeslot("11:00", "12:15", "AM"), new Timeslot("12:30", "01:45", "PM"), new Timeslot("02:00", "03:15", "PM"), new Timeslot("03:30", "04:45", "PM"), new Timeslot("06:00", "07:15", "PM"), new Timeslot("07:30", "08:45", "PM") };
        
        // Night Classes - 2.5 hour timeslots  for weekdays 
        Timeslot[] times_Night_Monday = { new Timeslot("06:00", "08:30", "PM"), new Timeslot("08:45", "11:15", "PM") };
        Timeslot[] times_Night_Tuesday = { new Timeslot("06:00", "08:30", "PM"), new Timeslot("08:45", "11:15", "PM") };
        Timeslot[] times_Night_Wednesday = { new Timeslot("06:00", "08:30", "PM"), new Timeslot("08:45", "11:15", "PM") };
        Timeslot[] times_Night_Thursday = { new Timeslot("06:00", "08:30", "PM"), new Timeslot("08:45", "11:15", "PM") };
        Timeslot[] times_Night_Friday = { new Timeslot("06:00", "08:30", "PM"), new Timeslot("08:45", "11:15", "PM") };

        int[] classGridLocation = new int[99];

        List<Timeslot> times_Default = new List<Timeslot>(); //stores all timeslots information made from excel file
        List<string> times_Default_Timetable = new List<string>();
        List<int> times_Default_Room = new List<int>();
        List<int> times_Default_Row = new List<int>();

        List<Timeslot> masterTimeslotList = new List<Timeslot>(); //stores all changed timeslot information
        List<int> masterTimetableList = new List<int>();
        List<int> masterClassRoomList = new List<int>();

        


        int defaultRowCountMWF = 1;
        int defaultRowCountTR = 1;

        List<string> labelNames = new List<string>();// stores names of label to assign classes to correct slot
        List<int> labelColumn = new List<int>();
        List<int> labelRow = new List<int>();
        List<string> labelTimeTable = new List<string>();
        int labelCount = 0;

        Excel.Font headerFont;
        Excel.Font[] columnFont = new Excel.Font[99];

        string headerFontName;
        string headerFontStyle;
        int headerFontSize;

        string[] colFontName = new string[99]; //stores attributes of excel file
        string[] colFontStyle = new string[99];
        int[] colFontSize = new int[99];

        int columnCount = 0;
        public static int INCREMENT = 50;
        public static int INCREMENTTR = 75;
        
        List<int> changedRoomNum = new List<int>(); // saves room number
        List<string> changedStartTime = new List<string>(); //stores changed timeslot start time
        List<string> changedEndTime = new List<string>();
        List<int> changedRow = new List<int>(); //stores changed time slot row
        List<int> changedColumn = new List<int>(); // stores changed timeslot column
        List<string> changedMeridian = new List<string>(); // stores changed timeslot meridian
        List<int> changedTimeTable = new List<int>();
        int count = 0; //stores number of changed timeslots

        List<int> autoChangedRoomNum = new List<int>(); // saves room number
        List<string> autoChangedStartTime = new List<string>(); //stores auto changed timeslot start time
        List<string> autoChangedEndTime = new List<string>();
        List<int> autoChangedRow = new List<int>(); //stores auto changed time slot row
        List<int> autoChangedColumn = new List<int>(); // stores auto changed timeslot column
        List<string> autoChangedMeridian = new List<string>(); // stores auto changed timeslot meridian
        List<int> autoChangedTimeTable = new List<int>();
        int autoCount = 0; //stores number of changed auto timeslots
        int errorCount = 0; //store number of errors on startup

        ClassRoomList classrooms = (ClassRoomList)System.Windows.Application.Current.FindResource("ClassRoom_List_View");
        ProfessorList professors = (ProfessorList)System.Windows.Application.Current.FindResource("Professor_List_View");
        ProfessorPreferenceList professorPreferences = (ProfessorPreferenceList)System.Windows.Application.Current.FindResource("ProfPreference_List_View");
        ClassGroupList classGroupings = (ClassGroupList)System.Windows.Application.Current.FindResource("ClassGroup_List_View");
        ClassList classList = (ClassList)System.Windows.Application.Current.FindResource("Classes_List_View");
        ClassList unassignedClasses = (ClassList)System.Windows.Application.Current.FindResource("Unassigned_Classes_List_View");
        ClassList onlineClasses = (ClassList)System.Windows.Application.Current.FindResource("Online_Classes_List_View");
        ClassList appointmentClasses = (ClassList)System.Windows.Application.Current.FindResource("Appointment_Classes_List_View");
        ClassList appointment2Classes = (ClassList)System.Windows.Application.Current.FindResource("Appointment2_Classes_List_View");
        ClassList deletedClasses = (ClassList)System.Windows.Application.Current.FindResource("Deleted_Classes_List_View");
        ClassList singleDayClasses = (ClassList)System.Windows.Application.Current.FindResource("Single_Day_Classes_List_View");
        List<string> excelHeaders = new List<string>();
        List<Type> excelTypes = new List<Type>();
        List<ClassesHash> hashedClasses = new List<ClassesHash>();
        string filePath, latestHashDigest, colorFilePath;
        RGB_Color[] colorPalette = { new RGB_Color(244, 67, 54), new RGB_Color(156, 39, 176), new RGB_Color(63, 81, 181), new RGB_Color(3, 169, 244), new RGB_Color(0, 150, 136), new RGB_Color(139, 195, 74), new RGB_Color(255, 235, 59), new RGB_Color(255, 152, 0), new RGB_Color(233, 30, 99), new RGB_Color(103, 58, 183), new RGB_Color(33, 150, 243), new RGB_Color(0, 188, 212), new RGB_Color(76, 175, 80), new RGB_Color(205, 220, 57), new RGB_Color(255, 193, 7), new RGB_Color(255, 87, 34) };
        Pairs colorPairs = (Pairs)System.Windows.Application.Current.FindResource("ColorPairs_List_View");
        string term, termString, termYear, errorMSG;

        string savefileName;

        ////////////// START OF EXECUTION ////////////////

        public MainWindow()
        {
            InitializeComponent();
            InitializeErrorLog();
            filePath = System.Windows.Application.Current.Resources["FilePath"].ToString(); // make local copy of path to excel file (initialized by FileSelect window)
            ReadExcel(filePath);
            AssignProfColors();
            DrawTimeTables();
            FillDerivedLists();
            UpdateProfessorCapacity();
            BindData();
            GenerateClassListHashes();
            ErrorLog();
            startupErrors();
            latestHashDigest = ComputeSha256Hash(classList.Serialize()); // initialize hash digest of classlist (used to see if changes have been made before closing application)
            Helper.CloseUniqueWindow<FileSelect>(); // Close File Select window
            
        }


        ////////////// FUNCTIONS ////////////////





        /// <summary>
        /// Reads in Class Scheduling information, create classes objects and append them to classList
        /// </summary>
        /// <param name="file"></param>
        /// 



        public void ReadExcel(string file)
        {
            //read in headers file
            string fileName = "Headers.xlsx";
            FileInfo f = new FileInfo(fileName);
            string name = f.FullName;
            var headerFile = new XLWorkbook(name);
            var headerWorksheet = headerFile.Worksheet("Headers");
            int numberHeaders = headerWorksheet.RowsUsed().Count();
            //Excel.Worksheet classroomInfo = headerFile.Worksheet["ClassroomInfo"];
            var classroomInfo = headerFile.Worksheet("ClassroomInfo");
            int numberRoomHeaders = classroomInfo.RowsUsed().Count();
            var professorInfo = headerFile.Worksheet("ProfessorInfo");
            int numberProfHeaders = professorInfo.RowsUsed().Count();

            //read headers from headers file into list
            var headerList = new List<string>();
            for (int i = 0; i < numberHeaders; i++)
            {
                headerList.Add(headerWorksheet.Row(i + 1).Cell(1).GetValue<string>());
            }

            //read room headers from classroom worksheet into list
            var roomHeaderList = new List<string>();
            for (int i = 0; i < numberRoomHeaders; i++)
            {
                roomHeaderList.Add(classroomInfo.Column(i + 1).Cell(1).GetValue<string>());
            }

            //read professor headers from professor worksheet into list
            var profHeaderList = new List<string>();
            for (int i = 0; i < numberProfHeaders; i++)
            {
                profHeaderList.Add(professorInfo.Column(i + 1).Cell(1).GetValue<string>());
            }


            using (var ignore = new XLWorkbook())
            {
                // Select Worksheet
                Excel.Application oExcel = new Microsoft.Office.Interop.Excel.Application();
                Excel.Workbook excelWorkbook = oExcel.Workbooks.Open(file);
                Excel.Worksheet worksheet = (Excel.Worksheet)excelWorkbook.Worksheets[1];
                Excel.Range range = worksheet.UsedRange;
                var classInfo = worksheet;



                int columns = range.Columns.Count;
                var rows = range.Rows.Count;
                var roomWorksheet = headerFile.Worksheet("ClassroomInfo");
                int roomColumns = numberRoomHeaders;
                var roomRows = roomWorksheet.RangeUsed().RowsUsed().Skip(1);

                var profWorksheet = headerFile.Worksheet("ProfessorInfo");
                int profColumns = numberProfHeaders;
                var profRows = profWorksheet.RangeUsed().RowsUsed().Skip(1);



                //read headers from class information file into list
                var excelHeaderList = new List<string>();
                int invalidHeaders = 0;
                for (int i = 0; i < columns; i++)
                {
                    try
                    {
                        excelHeaderList.Add((string)(worksheet.Cells[1, i + 1] as Excel.Range).Value);
                    }
                    catch (Exception ex)
                    {
                        //unable to read in headers
                        invalidHeaders = invalidHeaders + 1;
                        break;
                    }

                }



                string missingheadersstring = "Missing Headers: ";
                int duplicateCount = 0;
                //Check if headers match header file
                int headerCount = 0;
                //if able to read in headers 
                if (invalidHeaders > 0)
                {
                    headerCount = headerList.Count - 1;

                }
                else
                {

                    for (int i = 0; i < excelHeaderList.Count; i++)
                    {
                        for (int j = 0; j < headerList.Count; j++)
                        {
                            if (excelHeaderList.ElementAt(i) == headerList.ElementAt(j))
                            {
                                headerCount++;
                                break;

                            }
                            else if (excelHeaderList.ElementAt(i) == headerList.ElementAt(j) + "(2)")
                            {
                                headerCount++;
                                duplicateCount++;
                                break;

                            }
                            else if (j == headerList.Count - 1 && excelHeaderList.ElementAt(i) != headerList.ElementAt(j))
                            {
                               // missingheadersstring = missingheadersstring + headerWorksheet.Row(i + 1).Cell(1).GetValue<string>() + "\n";

                            }

                        }
                    }
                }


                //check why headers are wrong
                if (headerCount < headerList.Count)
                {
                    //write out to error log
                    if (invalidHeaders > 0)
                    {
                        System.Windows.MessageBox.Show("Unable to read in headers. Please choose a new file or refer to the user manual for detailed infromation on reading in excel files.");
                        String errorMSG = "Unable to read in headers. Please choose a new file or refer to the user manual for detailed infromation on reading in excel files.";
                    }
                    else
                    {
                        System.Windows.MessageBox.Show("File is missing needed headers. Please refer to the User Manual for the complete list of needed headers");
                            
                        String errorMSG = "File is missing needed headers. Please refer to the User Manual for the complete list of needed headers";
                    }

                    System.Windows.MessageBox.Show("Please select a new file or exit the program.");
                    WriteErrorLog(errorMSG);
                    SaveErrorHistory();
                    oExcel.Workbooks.Close();
                    System.Windows.Forms.Application.Restart();
                    //Helper.CloseUniqueWindow<FileSelect>(); // Close File Select window

                    System.Environment.Exit(0);


                }





                // Index headers based on location in class information excel file. 
                //If a new header is added and needed, then you will need to create a new index to search for its location. 
                int indexTerm = excelHeaderList.IndexOf("Term") + 1; //1
                int indexSESSION = excelHeaderList.IndexOf("SESSION") + 1; //2 
                int indexSUBJ = excelHeaderList.IndexOf("SUBJ") + 1; //3
                int indexCOURSE = excelHeaderList.IndexOf("COURSE") + 1; //4
                int indexSECT = excelHeaderList.IndexOf("SECT") + 1; //5
                int indexCRN = excelHeaderList.IndexOf("CRN") + 1; //6
                int indexDESC = excelHeaderList.IndexOf("DESCRIPTION") + 1; //7
                int indexCross = excelHeaderList.IndexOf("CrossList") + 1; //8
                int indexCREDITS = excelHeaderList.IndexOf("CREDITS") + 1; //9
                int indexMAXS = excelHeaderList.IndexOf("MAX SEATS") + 1; //10
                int indexWait = excelHeaderList.IndexOf("Wait List") + 1; //11
                int indexProjSeats = excelHeaderList.IndexOf("ProjSeats") + 1; //12
                int indexEnroll = excelHeaderList.IndexOf("Enrolled") + 1; //13
                int indexStart = excelHeaderList.IndexOf("Start") + 1; //14
                int indexEnd = excelHeaderList.IndexOf("End") + 1; //15
                int indexDAYS = excelHeaderList.IndexOf("DAYS") + 1;//16
                int indexBEGIN = excelHeaderList.IndexOf("BEGIN") + 1;//17

                int indexEND2;
                if (duplicateCount == 0)
                {
                    indexEND2 = excelHeaderList.IndexOf("END") + 1;//18
                }
                else
                {
                    indexEND2 = excelHeaderList.IndexOf("END(2)") + 1;//18
                }

                int indexRCAP = excelHeaderList.IndexOf("ROOM_CAP") + 1;//19
                int indexBLDG = excelHeaderList.IndexOf("BLDG") + 1;//20
                int indexROOM = excelHeaderList.IndexOf("ROOM") + 1;//21
                int indexFaculty = excelHeaderList.IndexOf("Faculty Name") + 1;//22
                int indexFacultyID = excelHeaderList.IndexOf("Faculty ID") + 1;//23
                int indexFacultyNum = excelHeaderList.IndexOf("# Faculty") + 1;//24
                int indexMthd = excelHeaderList.IndexOf("Instruction Mthd") + 1; //25
                int indexHIP = excelHeaderList.IndexOf("HIP") + 1; //26
                int indexPedagogy = excelHeaderList.IndexOf("PPedagogy") + 1; //27
                int indexPSYC = excelHeaderList.IndexOf("PSYC") + 1; //28
                int indexHONR = excelHeaderList.IndexOf("HONR") + 1; //29
                int indexHNRS = excelHeaderList.IndexOf("HNRS") + 1; //30 
                int indexSPEC = excelHeaderList.IndexOf("SPEC") + 1; //31
                int indexSNOTES = excelHeaderList.IndexOf("SECTION_NOTES") + 1; //32
                int indexNOTES = excelHeaderList.IndexOf("NOTES") + 1; //33



                int indexClassroomCap = roomHeaderList.IndexOf("ROOM_CAP") + 1;//room Worksheet
                int indexClassroomBLDG = roomHeaderList.IndexOf("BLDG") + 1;//room Worksheet
                int indexClassroomNUM = roomHeaderList.IndexOf("ROOM") + 1; //room Worksheet
                int indexClassroomNOTES = roomHeaderList.IndexOf("Notes") + 1; //room Worksheet



                int indexProfName = profHeaderList.IndexOf("Name") + 1;//professor Worksheet
                int indexProfID = profHeaderList.IndexOf("ID") + 1;//professor Worksheet
                int indexMaxCredit = profHeaderList.IndexOf("Max Credit") + 1; //professor Worksheet
                int indexMaxPrep = profHeaderList.IndexOf("Max Prep") + 1; //professor Worksheet

                //Check information from default and master scheduling sheet

               


                // Determine term
                string termCode = "";
                try
                {
                    termCode = (worksheet.Cells[2, indexTerm] as Excel.Range).Value.ToString();
                }
                catch (Exception ex)
                {
                    System.Windows.MessageBox.Show("Invalid term code contained in excel sheet. Please refer to the user manual for an example term code that is able to be accepted by this program.");
                    errorMSG = "Invalid term code contained in excel sheet. Please refer to the user manual for an example term code that is able to be accepted by this program.";
                    System.Windows.MessageBox.Show("Please select a new file or exit the program.");
                    WriteErrorLog(errorMSG);
                    SaveErrorHistory();
                    oExcel.Workbooks.Close();
                    System.Windows.Forms.Application.Restart();
                    //Helper.CloseUniqueWindow<FileSelect>(); // Close File Select window

                    System.Environment.Exit(0);


                }
                if (termCode.Length == 6)
                {
                    termYear = termCode.Substring(0, termCode.Length - 2);
                    TermYearBox.Text = termYear;
                    term = termCode.Substring(termCode.Length - 2);
                    switch (term)
                    {
                        case "01":
                            termString = "Spring";
                            TermComboBox.SelectedIndex = 0;
                            break;
                        case "06":
                            termString = "Summer";
                            TermComboBox.SelectedIndex = 1;
                            break;
                        case "09":
                            termString = "Fall";
                            TermComboBox.SelectedIndex = 2;
                            break;
                        case "12":
                            termString = "Winter";
                            TermComboBox.SelectedIndex = 3;
                            break;
                        default:
                            termString = "None";
                            TermComboBox.SelectedIndex = 0;
                            break;
                    }
                    //MessageBox.Show("Term: " + term + "\nYear: " + termYear);

                }
                else
                {
                    term = "00";
                    termString = "None";
                    termYear = "0000";
                }

                // Populate excel headers array
                Excel.Range headerRow = range.Rows[1];
                string cellValue = "";
                for (int i = 0; i < numberHeaders; i++)
                {
                    cellValue = (string)(headerRow.Cells[i + 1] as Excel.Range).Value;
                    for (int n = 0; n < excelHeaders.Count; n++)
                    {
                        if (excelHeaders[n].ToUpper() == cellValue.ToUpper()) // if there is a duplicate column name
                        {
                            cellValue = cellValue + "(2)";
                            break;
                        }
                    }
                    excelHeaders.Add(cellValue);
                }

                // Create Professors
                int sruid_indexer = 0;
                bool professorFound;
                for (var i = 2; i < rows; i++)
                {
                    string fullName, lastName, firstName, SRUID;

                    try
                    {
                        if ((string)(worksheet.Cells[i, indexFaculty] as Excel.Range).Value != null)
                        {
                            fullName = (string)(worksheet.Cells[i, indexFaculty] as Excel.Range).Value;
                            if (fullName != "" && fullName.Contains(","))
                            {
                                professorFound = false;
                                for (int n = 0; n < professors.Count; n++)
                                {
                                    if (professors[n].FullName == fullName)
                                    {
                                        professorFound = true;
                                        break;
                                    }
                                }
                                if (!professorFound)
                                {
                                    lastName = fullName.Split(',')[0];
                                    firstName = fullName.Split(',')[1].Remove(0, 1);
                                    string length = (string)(worksheet.Cells[i, indexFacultyID] as Excel.Range).Value;

                                    if (length != null)
                                    {
                                        if ((string)(worksheet.Cells[i, indexFacultyID] as Excel.Range).Value != "" && length.Length == 9)
                                        {
                                            SRUID = (string)(worksheet.Cells[i, indexFacultyID] as Excel.Range).Value;
                                        }
                                        else
                                        {
                                            SRUID = "A0" + sruid_indexer;
                                            sruid_indexer++;
                                        }

                                        if (firstName != "None")
                                        {
                                            professors.Add(new Professors(firstName, lastName, SRUID));
                                            
                                        }
                                    }
                                }
                            }

                        }
                        
                    }

                    

                    catch(Exception ex)
                    {
                        System.Windows.MessageBox.Show("Invalid Professor name at row " + i + ".Please refer to the user manual for an example professor name that is able to be accepted by this program.");
                        errorMSG = "Invalid Professor name at row " + i + ". Please refer to the user manual for an example professor name that is able to be accepted by this program. ";
                        System.Windows.MessageBox.Show("Please select a new file or exit the program.");
                        WriteErrorLog(errorMSG);
                        SaveErrorHistory();
                        oExcel.Workbooks.Close();
                        System.Windows.Forms.Application.Restart();
                        //Helper.CloseUniqueWindow<FileSelect>(); // Close File Select window

                        System.Environment.Exit(0);


                    }


                    
                    
                }

                for (int i = 0; i < professors.Count(); i++)
                {
                    for (int j = 0; j < i; j++)
                    {
                        if (professors[j].LastName != professors[i].LastName)
                        {
                            
                            string professorA = professors[i].LastName;
                            string professorB = professors[j].LastName;

                            char[] sortingA = professorA.ToCharArray();
                            char[] sortingB = professorB.ToCharArray();

                            int sortedA = (int)sortingA[0] % 32;
                            int sortedB = (int)sortingB[0] % 32;

                            if (sortedB > sortedA)
                            {
                                Professors tmp = professors[i];
                                professors[i] = professors[j];
                                professors[j] = tmp;
                            }

                        }
                    }
                }

                // Create Classrooms
                int parseResult, room, capacity;
                bool classroomFound;
                string bldg, roomNotes;
                foreach (var row in roomRows)
                {
                    room = -1;
                    capacity = 0;
                    roomNotes = "";

                    try
                    {
                        if (!row.Cell(indexClassroomBLDG).IsEmpty())
                        {
                            bldg = row.Cell(indexClassroomBLDG).GetValue<string>().ToUpper();
                            if (bldg != "WEB" && !bldg.Contains("APPT"))
                            {
                                if (!row.Cell(indexClassroomNUM).IsEmpty() && int.TryParse(row.Cell(indexClassroomNUM).GetValue<string>(), out parseResult))
                                {
                                    room = parseResult;
                                    if (!row.Cell(indexClassroomCap).IsEmpty() && int.TryParse(row.Cell(indexClassroomCap).GetValue<string>(), out parseResult))
                                    {
                                        capacity = parseResult;
                                    }
                                    if (!row.Cell(indexClassroomNOTES).IsEmpty())
                                    {
                                        roomNotes = row.Cell(indexClassroomNOTES).GetValue<string>();
                                    }
                                }
                                classroomFound = false;
                                for (int n = 0; n < classrooms.Count; n++)
                                {
                                    if (classrooms[n].ClassID == (bldg + room))
                                    {
                                        classroomFound = true;
                                        break;
                                    }
                                }
                                if (!classroomFound)
                                {
                                    classrooms.Add(new ClassRoom(bldg, room, capacity, roomNotes));
                                }
                            }
                        }

                    }
                    catch(Exception ex)
                    {
                        System.Windows.MessageBox.Show("Unable to create classroom. Please refer to the user manual for example classroom information that is able to be accepted by this program.");
                        errorMSG = "Unable to create classroom. Please refer to the user manual for example classroom information that is able to be accepted by this program.";
                        System.Windows.MessageBox.Show("Please select a new file or exit the program.");
                        WriteErrorLog(errorMSG);
                        SaveErrorHistory();
                        oExcel.Workbooks.Close();
                        System.Windows.Forms.Application.Restart();
                        //Helper.CloseUniqueWindow<FileSelect>(); // Close File Select window

                        System.Environment.Exit(0);


                    }

                    
                }

                // Create Classes
                int Session = 0, ClassNum, Section, Credits, SeatsTaken, Term = 0, maxSeats = 0, projSeats = 0, Enrolled = 0, Waitlist = 0;
                int duplicate_CRN_indexer = -1;
                string CRN, Dept, ClassName, ClassDay, classID, profName, notes, sectionNotes, Crosslist = "", StartDate = "", EndDate = "", Building = "", Room = "", RoomCap = "";
                bool Online, Appoint, Changed;
                List<string> CRN_List = new List<string>();
                List<string> CrossListCodes = new List<string>();
                for (var i = 2; i <= rows; i++)
                {
                    ClassNum = -1;
                    Section = -1;
                    Credits = 0;
                    SeatsTaken = 0;
                    Dept = "";
                    ClassName = "";
                    ClassDay = "";
                    Professors prof = new Professors();
                    ClassRoom classroom = new ClassRoom();
                    Timeslot time = new Timeslot();
                    Online = false;
                    Appoint = false;
                    Changed = false;
                    notes = "";
                    sectionNotes = "";
                    Crosslist = "";


                    // CRN 
                    // Primary Key, if CRN is empty, do not enter record.
                    // If CRN is not a number, assign a unique negative value, tracked by duplicate_CRN_indexer
                    if ((worksheet.Cells[i, indexCRN] as Excel.Range).Value == null)
                    {
                        if (((worksheet.Cells[i, indexCOURSE] as Excel.Range).Value) != null && (worksheet.Cells[i, indexSECT] as Excel.Range).Value != null && (worksheet.Cells[i, indexSUBJ] as Excel.Range).Value != null)
                        {

                            errorMSG = "Missing CRN value for class: " + (worksheet.Cells[i, indexSUBJ] as Excel.Range).Value.ToString() + " " + (worksheet.Cells[i, indexCOURSE] as Excel.Range).Value.ToString() + " , section: " + (worksheet.Cells[i, indexSECT] as Excel.Range).Value.ToString();
                            WriteErrorLog(errorMSG);
                        }
                        else if (((worksheet.Cells[i, indexCOURSE] as Excel.Range).Value) != null && (worksheet.Cells[i, indexSECT] as Excel.Range).Value == null && (worksheet.Cells[i, indexSUBJ] as Excel.Range).Value != null)
                        {
                            errorMSG = "Missing CRN value for class: " + (worksheet.Cells[i, indexSUBJ] as Excel.Range).Value.ToString() + " " + (worksheet.Cells[i, indexCOURSE] as Excel.Range).Value.ToString();
                            WriteErrorLog(errorMSG);

                        }
                        else
                        {
                            errorMSG = "Missing value for a CRN. Not enought information was given to know which class.";
                            WriteErrorLog(errorMSG);
                        }
                    }


                    // CRN Cont.


                    if ((string)(worksheet.Cells[i, indexCRN] as Excel.Range).Value != null)
                    {
                        parseResult = -1;
                        CRN = "";
                        try
                        {
                            if (int.TryParse((string)(worksheet.Cells[i, indexCRN] as Excel.Range).Value, out parseResult))
                            {
                                CRN = (string)(worksheet.Cells[i, indexCRN] as Excel.Range).Value;
                                bool duplicate_CRN = false;
                                for (int n = 0; n < CRN_List.Count; n++)
                                {
                                    if (CRN_List[n] == CRN)
                                    {
                                        duplicate_CRN = true;
                                        break;
                                    }
                                }
                                if (!duplicate_CRN)
                                {
                                    CRN_List.Add(CRN);
                                }
                                else
                                {
                                    //System.Windows.MessageBox.Show("Duplicate CRN found in excel file./nCRN: " + CRN);
                                    errorMSG = "Duplicate CRN found in excel file. CRN: " + CRN;
                                    WriteErrorLog(errorMSG);
                                    CRN = duplicate_CRN_indexer.ToString(); // HANDLE THIS BETTER
                                    duplicate_CRN_indexer--;
                                }
                            }
                            else
                            {
                                string textCRN = (string)(worksheet.Cells[i, indexCRN] as Excel.Range).Value.ToUpper();
                                if (textCRN == "NEW")
                                {
                                    CRN = textCRN;
                                }
                                else
                                {
                                    //System.Windows.MessageBox.Show("CRN field is not a number or a string!./nCRN: " + (string)(worksheet.Cells[i, indexCRN] as Excel.Range).Value);
                                    errorMSG = "CRN field is not a number or a string!./ nCRN: " + (string)(worksheet.Cells[i, indexCRN] as Excel.Range).Value;
                                    WriteErrorLog(errorMSG);
                                    CRN = duplicate_CRN_indexer.ToString();
                                    duplicate_CRN_indexer--;
                                }
                            }

                        }
                        catch (Exception ex)
                        {
                            System.Windows.MessageBox.Show("Unable to understand the CRN at row " + i + ". Please refer to the user manual for example CRN information that is able to be accepted by this program.");
                            errorMSG = "Unable to understand the CRN at row " + i + ". Please refer to the user manual for example CRN information that is able to be accepted by this program.";
                            System.Windows.MessageBox.Show("Please select a new file or exit the program.");
                            WriteErrorLog(errorMSG);
                            SaveErrorHistory();
                            oExcel.Workbooks.Close();
                            System.Windows.Forms.Application.Restart();
                            //Helper.CloseUniqueWindow<FileSelect>(); // Close File Select window

                            System.Environment.Exit(0);

                        }


                        // DEPT
                        Dept = "";
                        try
                        {
                            if ((string)(worksheet.Cells[i, indexSUBJ] as Excel.Range).Value != "")
                            {
                                Dept = (string)(worksheet.Cells[i, indexSUBJ] as Excel.Range).Value.ToUpper();
                            }
                        }
                        catch(Exception ex)
                        {
                            System.Windows.MessageBox.Show("Unable to understand the Department at row " + i + ". Please refer to the user manual for example Department information that is able to be accepted by this program.");
                            errorMSG = "Unable to understand the Department at row " + i + ". Please refer to the user manual for example Department information that is able to be accepted by this program.";
                            System.Windows.MessageBox.Show("Please select a new file or exit the program.");
                            WriteErrorLog(errorMSG);
                            SaveErrorHistory();
                            oExcel.Workbooks.Close();
                            System.Windows.Forms.Application.Restart();
                            //Helper.CloseUniqueWindow<FileSelect>(); // Close File Select window

                            System.Environment.Exit(0);

                        }
                        // CLASS NUM
                        ClassNum = 0;
                        try
                        {
                            if ((string)(worksheet.Cells[i, indexSUBJ] as Excel.Range).Value != "")
                            {
                                int classNumPulled = 0;
                                classNumPulled = Convert.ToInt32((worksheet.Cells[i, indexCOURSE] as Excel.Range).Value);
                                if (classNumPulled > 0)
                                {
                                    ClassNum = classNumPulled;
                                }
                            }

                        }
                        catch(Exception ex)
                        {
                            System.Windows.MessageBox.Show("Unable to understand the Class Number at row " + i + ". Please refer to the user manual for example Class information that is able to be accepted by this program.");
                            errorMSG = "Unable to understand the Class Number at row " + i + ". Please refer to the user manual for example Class information that is able to be accepted by this program.";
                            System.Windows.MessageBox.Show("Please select a new file or exit the program.");
                            WriteErrorLog(errorMSG);
                            SaveErrorHistory();
                            oExcel.Workbooks.Close();
                            System.Windows.Forms.Application.Restart();
                            //Helper.CloseUniqueWindow<FileSelect>(); // Close File Select window

                            System.Environment.Exit(0);

                        }

                        // CLASS NAME
                        ClassName = "";
                        try
                        {
                            if ((string)(worksheet.Cells[i, indexDESC] as Excel.Range).Value != "")
                            {
                                ClassName = (string)(worksheet.Cells[i, indexDESC] as Excel.Range).Value;
                            }

                        }
                        catch(Exception ex)
                        {
                            System.Windows.MessageBox.Show("Unable to understand the Class Name at row " + i + ". Please refer to the user manual for example Class information that is able to be accepted by this program.");
                            errorMSG = "Unable to understand the Class Name at row " + i + ". Please refer to the user manual for example Class information that is able to be accepted by this program.";
                            System.Windows.MessageBox.Show("Please select a new file or exit the program.");
                            WriteErrorLog(errorMSG);
                            SaveErrorHistory();
                            oExcel.Workbooks.Close();
                            System.Windows.Forms.Application.Restart();
                            //Helper.CloseUniqueWindow<FileSelect>(); // Close File Select window

                            System.Environment.Exit(0);

                        }


                        // SECTION
                        Section = 0;
                        try 
                        {
                            if ((string)(worksheet.Cells[i, indexSECT] as Excel.Range).Value != "")
                            {
                                if (int.TryParse((string)(worksheet.Cells[i, indexSECT] as Excel.Range).Value, out parseResult))
                                {
                                    if (parseResult > 0)
                                    {
                                        Section = parseResult;
                                    }
                                }
                            }
                        }
                        catch(Exception ex)
                        {
                            System.Windows.MessageBox.Show("Unable to understand the Section at row " + i + ". Please refer to the user manual for example Section information that is able to be accepted by this program.");
                            errorMSG = "Unable to understand the Section at row " + i + ". Please refer to the user manual for example Section information that is able to be accepted by this program.";
                            System.Windows.MessageBox.Show("Please select a new file or exit the program.");
                            WriteErrorLog(errorMSG);
                            SaveErrorHistory();
                            oExcel.Workbooks.Close();
                            System.Windows.Forms.Application.Restart();
                            //Helper.CloseUniqueWindow<FileSelect>(); // Close File Select window

                            System.Environment.Exit(0);

                        }

                        // CREDITS
                        Credits = 0;
                        try
                        {
                            int? creditsPulled = null;
                            creditsPulled = Convert.ToInt32((worksheet.Cells[i, indexCREDITS] as Excel.Range).Value);
                            if (creditsPulled != null)
                            {
                                Credits = Convert.ToInt32((worksheet.Cells[i, indexCREDITS] as Excel.Range).Value);
                            }
                        }
                        catch(Exception ex)
                        {
                            System.Windows.MessageBox.Show("Unable to understand the Credits at row " + i + ". Please refer to the user manual for example Credits information that is able to be accepted by this program.");
                            errorMSG = "Unable to understand the Credits at row " + i + ". Please refer to the user manual for example Credits information that is able to be accepted by this program.";
                            System.Windows.MessageBox.Show("Please select a new file or exit the program.");
                            WriteErrorLog(errorMSG);
                            SaveErrorHistory();
                            oExcel.Workbooks.Close();
                            System.Windows.Forms.Application.Restart();
                            //Helper.CloseUniqueWindow<FileSelect>(); // Close File Select window

                            System.Environment.Exit(0);

                        }

                        // SEATS TAKEN
                        try
                        {
                            int? seatsPulled = null;
                            seatsPulled = Convert.ToInt32((worksheet.Cells[i, indexEnroll] as Excel.Range).Value);
                            if (seatsPulled != null)
                            {
                                SeatsTaken = Convert.ToInt32((worksheet.Cells[i, indexEnroll] as Excel.Range).Value);
                            }
                        }
                        catch(Exception ex)
                        {
                            System.Windows.MessageBox.Show("Unable to understand the Credits at row " + i + ". Please refer to the user manual for example Credits information that is able to be accepted by this program.");
                            errorMSG = "Unable to understand the Credits at row " + i + ". Please refer to the user manual for example Credits information that is able to be accepted by this program.";
                            System.Windows.MessageBox.Show("Please select a new file or exit the program.");
                            WriteErrorLog(errorMSG);
                            SaveErrorHistory();
                            oExcel.Workbooks.Close();
                            System.Windows.Forms.Application.Restart();
                            //Helper.CloseUniqueWindow<FileSelect>(); // Close File Select window

                            System.Environment.Exit(0);

                        }

                        // CLASSDAY
                        ClassDay = "";
                        try
                        {
                            if ((string)(worksheet.Cells[i, indexDAYS] as Excel.Range).Value != "" && (string)(worksheet.Cells[i, indexDAYS] as Excel.Range).Value != null)
                            {
                                ClassDay = (string)(worksheet.Cells[i, indexDAYS] as Excel.Range).Value;
                                times_Default_Timetable.Add(ClassDay);
                            }
                        }
                        catch(Exception ex)
                        {
                            System.Windows.MessageBox.Show("Unable to understand the Class Day information at row " + i + ". Please refer to the user manual for example Class information that is able to be accepted by this program.");
                            errorMSG = "Unable to understand the Class Day information at row " + i + ". Please refer to the user manual for example Class information that is able to be accepted by this program.";
                            System.Windows.MessageBox.Show("Please select a new file or exit the program.");
                            WriteErrorLog(errorMSG);
                            SaveErrorHistory();
                            oExcel.Workbooks.Close();
                            System.Windows.Forms.Application.Restart();
                            //Helper.CloseUniqueWindow<FileSelect>(); // Close File Select window

                            System.Environment.Exit(0);

                        }

                        // Determine Professor
                        profName = "";
                        try
                        {
                            if ((string)(worksheet.Cells[i, indexFaculty] as Excel.Range).Value != "")
                            {
                                profName = (string)(worksheet.Cells[i, indexFaculty] as Excel.Range).Value;
                                for (int n = 0; n < professors.Count; n++)
                                {
                                    if (professors[n].FullName == profName)
                                    {
                                        prof = professors[n];
                                        break;
                                    }
                                }
                            }
                        }
                        catch(Exception ex)
                        {
                            System.Windows.MessageBox.Show("Unable to understand the Professor name at row " + i + ". Please refer to the user manual for example Professor information that is able to be accepted by this program.");
                            errorMSG = "Unable to understand the Professor name at row " + i + ". Please refer to the user manual for example Professor information that is able to be accepted by this program.";
                            System.Windows.MessageBox.Show("Please select a new file or exit the program.");
                            WriteErrorLog(errorMSG);
                            SaveErrorHistory();
                            oExcel.Workbooks.Close();
                            System.Windows.Forms.Application.Restart();
                            //Helper.CloseUniqueWindow<FileSelect>(); // Close File Select window

                            System.Environment.Exit(0);
                        }
                        
                        // Determine ClassRoom
                        try
                        {
                            if ((string)(worksheet.Cells[i, indexBLDG] as Excel.Range).Value != "" && (string)(worksheet.Cells[i, indexBLDG] as Excel.Range).Value != null)
                            {
                                bldg = (string)(worksheet.Cells[i, indexBLDG] as Excel.Range).Value.ToUpper();
                                if (bldg != "WEB" && !bldg.Contains("APPT"))
                                {
                                    room = -1;
                                    int? roomPulled = Convert.ToInt32((worksheet.Cells[i, indexROOM] as Excel.Range).Value);
                                    if (roomPulled != null)
                                    {

                                        room = (int)roomPulled;

                                    }
                                    times_Default_Room.Add(room);
                                    classID = bldg + room;
                                    for (int n = 0; n < classrooms.Count; n++)
                                    {
                                        if (classrooms[n].ClassID == classID)
                                        {
                                            classroom = classrooms[n];
                                            break;
                                        }
                                    }
                                }
                                else
                                {
                                    if (bldg == "WEB")
                                    {
                                        classroom = new ClassRoom("WEB", 999);
                                        Online = true;
                                    }
                                    else if (bldg.Contains("APPT"))
                                    {
                                        if (bldg == "APPT")
                                        {
                                            classroom = new ClassRoom("APPT", 0);
                                        }
                                        else if (bldg == "APPT2")
                                        {
                                            classroom = new ClassRoom("APPT2", 0);
                                        }
                                        Appoint = true;
                                    }
                                }
                            }

                        }
                        catch(Exception ex)
                        {
                            System.Windows.MessageBox.Show("Unable to understand the Classroom Information at row " + i + ". Please refer to the user manual for example Classroom information that is able to be accepted by this program.");
                            errorMSG = "Unable to understand the Classroom Information at row " + i + ". Please refer to the user manual for example Classroom information that is able to be accepted by this program.";
                            System.Windows.MessageBox.Show("Please select a new file or exit the program.");
                            WriteErrorLog(errorMSG);
                            SaveErrorHistory();
                            oExcel.Workbooks.Close();
                            System.Windows.Forms.Application.Restart();
                            //Helper.CloseUniqueWindow<FileSelect>(); // Close File Select window

                            System.Environment.Exit(0);

                        }

                        // Determine TimeSlot

                        try
                        {
                            if ((string)(worksheet.Cells[i, indexBEGIN] as Excel.Range).Value != null)
                            {
                                string meridianPart = "";
                                string rawTime = (string)(worksheet.Cells[i, indexBEGIN] as Excel.Range).Value;
                                string timePart = formatTime(rawTime.Split(' ')[0]);

                                string endTimeRaw = (string)(worksheet.Cells[i, indexEND2] as Excel.Range).Value;
                                string endTime = formatTime(endTimeRaw.Split(' ')[0]);
                                try
                                {
                                    meridianPart = rawTime.Split(' ')[1];
                                }
                                catch
                                {
                                    meridianPart = rawTime.Substring(4);
                                    if (meridianPart.Any(char.IsDigit))
                                    {
                                        meridianPart = meridianPart.Substring(1);
                                    }
                                }
                                string startMeridian = "";

                                char[] fixMeridian = meridianPart.ToCharArray().Where(c => !char.IsWhiteSpace(c)).ToArray();

                                for (int j = 0; j < fixMeridian.Length; j++)
                                {
                                    startMeridian = startMeridian + fixMeridian[j];
                                }

                                string meridian = startMeridian;
                                if (meridian == "A.M.")
                                {
                                    meridian = "AM";
                                }
                                else if (meridian == "P.M.")
                                {
                                    meridian = "PM";
                                }

                                times_Default.Add(new Timeslot(timePart, endTime, meridian));
                                int classroomCheck = Int32.Parse((worksheet.Cells[i, indexROOM] as Excel.Range).Value.ToString());
                                time = DetermineTime(timePart, ClassDay, classroomCheck, meridian);

                            }


                        }
                        catch(Exception ex)
                        {
                            System.Windows.MessageBox.Show("Unable to understand the Timeslot Information at row " + i + ". Please refer to the user manual for example Time information that is able to be accepted by this program.");
                            errorMSG = "Unable to understand the Timeslot Information at row " + i + ". Please refer to the user manual for example time information that is able to be accepted by this program.";
                            System.Windows.MessageBox.Show("Please select a new file or exit the program.");
                            WriteErrorLog(errorMSG);
                            SaveErrorHistory();
                            oExcel.Workbooks.Close();
                            System.Windows.Forms.Application.Restart();
                            //Helper.CloseUniqueWindow<FileSelect>(); // Close File Select window

                            System.Environment.Exit(0);
                        }
                        
                        // Determine if it is higlighted red (changed) or not in the excel file
                        string cellColor = worksheet.Cells[i, indexTerm].Interior.Color.ToString();
                        if (cellColor == "13619199")
                        {
                            //MessageBox.Show("Excel Read: Class set to RED");
                            Changed = true;
                        }




                        // Crosslist handling

                        bool isCrossFirst = false;
                        try
                        {
                           
                            string crossCode = (string)(worksheet.Cells[i, indexCross] as Excel.Range).Value;
                            if ((crossCode != "") && (crossCode != null))
                            {
                                for (int j = 0; j < CrossListCodes.Count; j++)
                                {


                                    if (crossCode == CrossListCodes[j])
                                    {
                                        break;
                                    }
                                    if (j == (CrossListCodes.Count - 1))
                                    {
                                        isCrossFirst = true;
                                        CrossListCodes.Add(crossCode);
                                    }
                                }
                                if (CrossListCodes.Count == 0)
                                {
                                    isCrossFirst = true;
                                    CrossListCodes.Add(crossCode);
                                }
                            }
                        }
                        catch(Exception ex)
                        {
                            System.Windows.MessageBox.Show("Unable to understand the Crosslist Information at row " + i + ". Please refer to the user manual for example CrossList information that is able to be accepted by this program.");
                            errorMSG = "Unable to understand the CrossList Information at row " + i + ". Please refer to the user manual for example CrossList information that is able to be accepted by this program.";
                            System.Windows.MessageBox.Show("Please select a new file or exit the program.");
                            WriteErrorLog(errorMSG);
                            SaveErrorHistory();
                            oExcel.Workbooks.Close();
                            System.Windows.Forms.Application.Restart();
                            //Helper.CloseUniqueWindow<FileSelect>(); // Close File Select window

                            System.Environment.Exit(0);

                        }
                        
                        //SESSION

                        int outputS;
                        try
                        {
                            if ((string)(worksheet.Cells[i, indexSESSION] as Excel.Range).Value != "")
                            {
                                var sessionPulled = (worksheet.Cells[i, indexSESSION] as Excel.Range).Value;
                                bool isNumeric = Int32.TryParse(sessionPulled, out outputS);
                                if (isNumeric)
                                {
                                    Session = Int32.Parse(sessionPulled);
                                }
                                if (!isNumeric)
                                {
                                    errorMSG = "Incompatible SESSION value for class: " + (worksheet.Cells[i, indexSUBJ] as Excel.Range).Value.ToString() + " " + (worksheet.Cells[i, indexCOURSE] as Excel.Range).Value.ToString() + " , section: " + (worksheet.Cells[i, indexSECT] as Excel.Range).Value.ToString();
                                    WriteErrorLog(errorMSG);
                                }
                                isNumeric = false; //reset value to false
                            }
                            if ((string)(worksheet.Cells[i, indexSESSION] as Excel.Range).Value == "")
                            {
                                errorMSG = "Missing SESSION value for class: " + (worksheet.Cells[i, indexSUBJ] as Excel.Range).Value.ToString() + " " + (worksheet.Cells[i, indexCOURSE] as Excel.Range).Value.ToString() + " , section: " + (worksheet.Cells[i, indexSECT] as Excel.Range).Value.ToString();
                                WriteErrorLog(errorMSG);

                            }
                        }
                        catch(Exception ex)
                        {
                            System.Windows.MessageBox.Show("Unable to understand the Session Information at row " + i + ". Please refer to the user manual for example Session information that is able to be accepted by this program.");
                            errorMSG = "Unable to understand the Session Information at row " + i + ". Please refer to the user manual for example Session information that is able to be accepted by this program.";
                            System.Windows.MessageBox.Show("Please select a new file or exit the program.");
                            WriteErrorLog(errorMSG);
                            SaveErrorHistory();
                            oExcel.Workbooks.Close();
                            System.Windows.Forms.Application.Restart();
                            //Helper.CloseUniqueWindow<FileSelect>(); // Close File Select window

                            System.Environment.Exit(0);
                        }

                        

                        //Term

                        int outputT;

                        try
                        {
                            if ((string)(worksheet.Cells[i, indexTerm] as Excel.Range).Value != "")
                            {
                                var TermPulled = (worksheet.Cells[i, indexTerm] as Excel.Range).Value;
                                bool isNumeric = Int32.TryParse(TermPulled, out outputT);
                                if (isNumeric)
                                {
                                    Term = Int32.Parse(TermPulled);
                                }
                                if (!isNumeric)
                                {
                                    errorMSG = "Incompatible Term value for class: " + (worksheet.Cells[i, indexSUBJ] as Excel.Range).Value.ToString() + " " + (worksheet.Cells[i, indexCOURSE] as Excel.Range).Value.ToString() + " , section: " + (worksheet.Cells[i, indexSECT] as Excel.Range).Value.ToString();
                                    WriteErrorLog(errorMSG);
                                }
                            }
                            if ((string)(worksheet.Cells[i, indexTerm] as Excel.Range).Value == "")
                            {
                                errorMSG = "Missing Term value for class: " + (worksheet.Cells[i, indexSUBJ] as Excel.Range).Value.ToString() + " " + (worksheet.Cells[i, indexCOURSE] as Excel.Range).Value.ToString() + " , section: " + (worksheet.Cells[i, indexSECT] as Excel.Range).Value.ToString();
                                WriteErrorLog(errorMSG);

                            }
                        }
                        catch(Exception e)
                        {
                            System.Windows.MessageBox.Show("Invalid term code contained in excel file. Please refer to the user manual for an example term code that is able to be accepted by this program.");
                            errorMSG = "Invalid term code contained in excel file. Please refer to the user manual for an example term code that is able to be accepted by this program.";
                            System.Windows.MessageBox.Show("Please select a new file or exit the program.");
                            WriteErrorLog(errorMSG);
                            SaveErrorHistory();
                            oExcel.Workbooks.Close();
                            System.Windows.Forms.Application.Restart();
                            //Helper.CloseUniqueWindow<FileSelect>(); // Close File Select window

                            System.Environment.Exit(0);
                        }
                        

                        // Max Seats
                        try
                        {
                            int? mSeatsPulled = null;
                            mSeatsPulled = Convert.ToInt32((worksheet.Cells[i, indexMAXS] as Excel.Range).Value);
                            if (mSeatsPulled != null)
                            {
                                maxSeats = Convert.ToInt32((worksheet.Cells[i, indexMAXS] as Excel.Range).Value);
                            }

                        }
                        catch(Exception ex)
                        {
                            System.Windows.MessageBox.Show("Invalid Max Seats contained in row " + i + ". Please refer to the user manual for an example Max Seats that is able to be accepted by this program.");
                            errorMSG = "Invalid Max Seats contained in row " + i + ". Please refer to the user manual for an example Max Seats that is able to be accepted by this program.";
                            System.Windows.MessageBox.Show("Please select a new file or exit the program.");
                            WriteErrorLog(errorMSG);
                            SaveErrorHistory();
                            oExcel.Workbooks.Close();
                            System.Windows.Forms.Application.Restart();
                            //Helper.CloseUniqueWindow<FileSelect>(); // Close File Select window

                            System.Environment.Exit(0);

                        }


                        // Proj Seats
                        try
                        {
                            int? pSeatsPulled = null;
                            pSeatsPulled = Convert.ToInt32((worksheet.Cells[i, indexProjSeats] as Excel.Range).Value);
                            if (pSeatsPulled != null)
                            {
                                projSeats = Convert.ToInt32((worksheet.Cells[i, indexProjSeats] as Excel.Range).Value);
                            }
                        }
                        catch(Exception ex)
                        {
                            System.Windows.MessageBox.Show("Invalid Projected Seats contained in row " + i + ". Please refer to the user manual for an example Projected Seats that is able to be accepted by this program.");
                            errorMSG = "Invalid Projected Seats contained in row " + i + ". Please refer to the user manual for an example Projected Seats that is able to be accepted by this program.";
                            System.Windows.MessageBox.Show("Please select a new file or exit the program.");
                            WriteErrorLog(errorMSG);
                            SaveErrorHistory();
                            oExcel.Workbooks.Close();
                            System.Windows.Forms.Application.Restart();
                            //Helper.CloseUniqueWindow<FileSelect>(); // Close File Select window

                            System.Environment.Exit(0);
                        }



                        // Enrolled
                        try
                        {
                            int? enrollPulled = null;
                            enrollPulled = Convert.ToInt32((worksheet.Cells[i, indexEnroll] as Excel.Range).Value);
                            if (enrollPulled != null)
                            {
                                Enrolled = Convert.ToInt32((worksheet.Cells[i, indexEnroll] as Excel.Range).Value);
                            }
                        }
                        catch (Exception ex)
                        {
                            System.Windows.MessageBox.Show("Invalid Enrolled value contained in row " + i + ". Please refer to the user manual for an example Enrolled value that is able to be accepted by this program.");
                            errorMSG = "Invalid Enrolled value contained in row " + i + ".. Please refer to the user manual for an example Enrolled value that is able to be accepted by this program.";
                            System.Windows.MessageBox.Show("Please select a new file or exit the program.");
                            WriteErrorLog(errorMSG);
                            SaveErrorHistory();
                            oExcel.Workbooks.Close();
                            System.Windows.Forms.Application.Restart();
                            //Helper.CloseUniqueWindow<FileSelect>(); // Close File Select window

                            System.Environment.Exit(0);
                        }


                        // Waitlist
                        try
                        {
                            int? waitlistPulled = null;
                            waitlistPulled = Convert.ToInt32((worksheet.Cells[i, indexWait] as Excel.Range).Value);
                            if (waitlistPulled != null)
                            {
                                Waitlist = Convert.ToInt32((worksheet.Cells[i, indexWait] as Excel.Range).Value);
                            }
                        }
                        catch (Exception ex)
                        {
                            System.Windows.MessageBox.Show("Invalid Waitlist value contained in row " + i + ". Please refer to the user manual for an example Wiatlist value that is able to be accepted by this program.");
                            errorMSG = "Invalid Waitlist value contained in row " + i + ". Please refer to the user manual for an example Waitlist value that is able to be accepted by this program.";
                            System.Windows.MessageBox.Show("Please select a new file or exit the program.");
                            WriteErrorLog(errorMSG);
                            SaveErrorHistory();
                            oExcel.Workbooks.Close();
                            System.Windows.Forms.Application.Restart();
                            //Helper.CloseUniqueWindow<FileSelect>(); // Close File Select window

                            System.Environment.Exit(0);
                        }


                        //Crosslist
                        try
                        {
                            if ((string)(worksheet.Cells[i, indexCross] as Excel.Range).Value != null)
                            {
                                Crosslist = (string)(worksheet.Cells[i, indexCross] as Excel.Range).Value;

                            }
                        }
                        catch (Exception ex)
                        {
                            System.Windows.MessageBox.Show("Invalid Crosslist value contained in row " + i + ". Please refer to the user manual for an example Crosslist value that is able to be accepted by this program.");
                            errorMSG = "Invalid Crosslist value contained in row " + i + ". Please refer to the user manual for an example Crosslist value that is able to be accepted by this program.";
                            System.Windows.MessageBox.Show("Please select a new file or exit the program.");
                            WriteErrorLog(errorMSG);
                            SaveErrorHistory();
                            oExcel.Workbooks.Close();
                            System.Windows.Forms.Application.Restart();
                            //Helper.CloseUniqueWindow<FileSelect>(); // Close File Select window

                            System.Environment.Exit(0);
                        }


                        // Start Date
                        try
                        {
                            if ((string)(worksheet.Cells[i, indexStart] as Excel.Range).Value.ToString() != null)
                            {
                                string[] TempStartDate = (worksheet.Cells[i, indexStart] as Excel.Range).Value.ToString().Split(' ');
                                StartDate = TempStartDate[0];
                            }
                        }
                        catch (Exception ex)
                        {
                            System.Windows.MessageBox.Show("Invalid Start Date contained in row " + i + ". Please refer to the user manual for an example Start Date that is able to be accepted by this program.");
                            errorMSG = "Invalid Start Date contained in row " + i + ". Please refer to the user manual for an example Start Date that is able to be accepted by this program.";
                            System.Windows.MessageBox.Show("Please select a new file or exit the program.");
                            WriteErrorLog(errorMSG);
                            SaveErrorHistory();
                            oExcel.Workbooks.Close();
                            System.Windows.Forms.Application.Restart();
                            //Helper.CloseUniqueWindow<FileSelect>(); // Close File Select window

                            System.Environment.Exit(0);
                        }


                        // End Date
                        try
                        {
                            if ((string)(worksheet.Cells[i, indexEnd] as Excel.Range).Value.ToString() != null)
                            {
                                string[] TempEndDate = (worksheet.Cells[i, indexEnd] as Excel.Range).Value.ToString().Split(' ');
                                EndDate = TempEndDate[0];
                            }
                        }
                        catch (Exception ex)
                        {
                            System.Windows.MessageBox.Show("Invalid End Date contained in row " + i + ". Please refer to the user manual for an example End Date that is able to be accepted by this program.");
                            errorMSG = "Invalid End Date contained in row " + i + ". Please refer to the user manual for an example End Date that is able to be accepted by this program.";
                            System.Windows.MessageBox.Show("Please select a new file or exit the program.");
                            WriteErrorLog(errorMSG);
                            SaveErrorHistory();
                            oExcel.Workbooks.Close();
                            System.Windows.Forms.Application.Restart();
                            //Helper.CloseUniqueWindow<FileSelect>(); // Close File Select window

                            System.Environment.Exit(0);
                        }


                        //Building
                        try
                        {
                            if ((string)(worksheet.Cells[i, indexBLDG] as Excel.Range).Value.ToString() != null)
                            {
                                Building = (worksheet.Cells[i, indexBLDG] as Excel.Range).Value.ToString();
                            }
                        }
                        catch (Exception ex)
                        {
                            System.Windows.MessageBox.Show("Invalid Building contained in row " + i + ". Please refer to the user manual for an example Building that is able to be accepted by this program.");
                            errorMSG = "Invalid Building contained in row " + i + ". Please refer to the user manual for an example Building that is able to be accepted by this program.";
                            System.Windows.MessageBox.Show("Please select a new file or exit the program.");
                            WriteErrorLog(errorMSG);
                            SaveErrorHistory();
                            oExcel.Workbooks.Close();
                            System.Windows.Forms.Application.Restart();
                            //Helper.CloseUniqueWindow<FileSelect>(); // Close File Select window

                            System.Environment.Exit(0);
                        }

                        //Room Num

                        try
                        {
                            if ((worksheet.Cells[i, indexROOM] as Excel.Range).Value != null)
                            {
                                if ((worksheet.Cells[i, indexROOM] as Excel.Range).Value.ToString().Equals("WEB"))
                                {
                                    Room = "WEB";
                                }
                                else if ((worksheet.Cells[i, indexROOM] as Excel.Range).Value.ToString().Equals("APPT"))
                                {
                                    Room = "APPT";
                                }
                                else if ((worksheet.Cells[i, indexROOM] as Excel.Range).Value.ToString().Equals("APPT2"))
                                {
                                    Room = "APPT2";
                                }
                                else
                                {
                                    Room = (worksheet.Cells[i, indexROOM] as Excel.Range).Value.ToString();
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            System.Windows.MessageBox.Show("Invalid Room type contained in row " + i + ". Please refer to the user manual for an example Room type that is able to be accepted by this program.");
                            errorMSG = "Invalid Room type contained in row " + i + ". Please refer to the user manual for an example Room type that is able to be accepted by this program.";
                            System.Windows.MessageBox.Show("Please select a new file or exit the program.");
                            WriteErrorLog(errorMSG);
                            SaveErrorHistory();
                            oExcel.Workbooks.Close();
                            System.Windows.Forms.Application.Restart();
                            //Helper.CloseUniqueWindow<FileSelect>(); // Close File Select window

                            System.Environment.Exit(0);
                        }



                        //Room Capacity
                        try
                        {
                            if ((worksheet.Cells[i, indexRCAP] as Excel.Range).Value != null)
                            {
                                RoomCap = (worksheet.Cells[i, indexRCAP] as Excel.Range).Value.ToString();
                            }
                        }
                        catch (Exception ex)
                        {
                            System.Windows.MessageBox.Show("Invalid Room capacity contained in row " + i + ". Please refer to the user manual for an example Room capacity that is able to be accepted by this program.");
                            errorMSG = "Invalid Room capacity contained in row " + i + ". Please refer to the user manual for an example Room capacity that is able to be accepted by this program.";
                            System.Windows.MessageBox.Show("Please select a new file or exit the program.");
                            WriteErrorLog(errorMSG);
                            SaveErrorHistory();
                            oExcel.Workbooks.Close();
                            System.Windows.Forms.Application.Restart();
                            //Helper.CloseUniqueWindow<FileSelect>(); // Close File Select window

                            System.Environment.Exit(0);
                        }



                        // Get remaining extra data
                        List<string> extras = new List<string>();

                        for (int x = indexFacultyNum; x <= indexFacultyNum; x++) // Faculty Num //0
                        {
                            if ((worksheet.Cells[i, indexFacultyNum] as Excel.Range).Value != null)
                            {
                                extras.Add((string)(worksheet.Cells[i, indexFacultyNum] as Excel.Range).Value.ToString());
                            }
                        }


                        for (int x = indexHIP; x <= indexHIP; x++) // HIP // 1
                        {
                            if ((worksheet.Cells[i, x] as Excel.Range).Value != null)
                            {
                                extras.Add((string)(worksheet.Cells[i, x] as Excel.Range).Value.ToString());
                            }
                            else
                            {
                                extras.Add("");
                            }
                        }


                        for (int x = indexPedagogy; x <= indexPedagogy; x++) // Pedagogy //2
                        {
                            if ((worksheet.Cells[i, x] as Excel.Range).Value != null)
                            {
                                extras.Add((string)(worksheet.Cells[i, x] as Excel.Range).Value.ToString());
                            }
                            else
                            {
                                extras.Add("");
                            }
                        }

                        for (int x = indexPSYC; x <= indexPSYC; x++) // PSYC // 3
                        {
                            if ((worksheet.Cells[i, x] as Excel.Range).Value != null)
                            {
                                extras.Add((string)(worksheet.Cells[i, x] as Excel.Range).Value.ToString());
                            }
                            else
                            {
                                extras.Add("");
                            }
                        }
                        for (int x = indexHONR; x <= indexHONR; x++) // HONR // 4
                        {
                            if ((worksheet.Cells[i, x] as Excel.Range).Value != null)
                            {
                                extras.Add((string)(worksheet.Cells[i, x] as Excel.Range).Value.ToString());
                            }
                            else
                            {
                                extras.Add("");
                            }
                        }

                        for (int x = indexHNRS; x <= indexHNRS; x++) // HNRS // 5
                        {
                            if ((worksheet.Cells[i, x] as Excel.Range).Value != null)
                            {
                                extras.Add((string)(worksheet.Cells[i, x] as Excel.Range).Value.ToString());
                            }
                            else
                            {
                                extras.Add("");
                            }
                        }

                        for (int x = indexSPEC; x <= indexSPEC; x++) // HNRS // 5
                        {
                            if ((worksheet.Cells[i, x] as Excel.Range).Value != null)
                            {
                                extras.Add((string)(worksheet.Cells[i, x] as Excel.Range).Value.ToString());
                            }
                            else
                            {
                                extras.Add("");
                            }
                        }

                        try
                        {
                            sectionNotes = (string)(worksheet.Cells[i, indexSNOTES] as Excel.Range).Value;
                        }
                        catch (Exception ex)
                        {
                            System.Windows.MessageBox.Show("Invalid Section Notes contained in row " + i + ". Please refer to the user manual for an example Section Notes that is able to be accepted by this program.");
                            errorMSG = "Invalid Section Notes contained in row " + i + ". Please refer to the user manual for an example Section Notes that is able to be accepted by this program.";
                            System.Windows.MessageBox.Show("Please select a new file or exit the program.");
                            WriteErrorLog(errorMSG);
                            SaveErrorHistory();
                            oExcel.Workbooks.Close();
                            System.Windows.Forms.Application.Restart();
                            //Helper.CloseUniqueWindow<FileSelect>(); // Close File Select window

                            System.Environment.Exit(0);
                        }

                        try
                        {
                            notes = (string)(worksheet.Cells[i, indexNOTES] as Excel.Range).Value;
                        }
                        catch (Exception ex)
                        {
                            System.Windows.MessageBox.Show("Invalid Notes contained in row " + i + ". Please refer to the user manual for an example Notes that is able to be accepted by this program.");
                            errorMSG = "Invalid Notes contained in row " + i + ". Please refer to the user manual for an example Notes that is able to be accepted by this program.";
                            System.Windows.MessageBox.Show("Please select a new file or exit the program.");
                            WriteErrorLog(errorMSG);
                            SaveErrorHistory();
                            oExcel.Workbooks.Close();
                            System.Windows.Forms.Application.Restart();
                            //Helper.CloseUniqueWindow<FileSelect>(); // Close File Select window

                            System.Environment.Exit(0);
                        }


                        // Create class and add to classlist

                        try
                        {
                            Classes tmpClass = new Classes(Term, Session, CRN, Dept, ClassNum, Section, ClassName, Credits, ClassDay, time, SeatsTaken, classroom, prof, Online, Appoint, Changed, notes, sectionNotes, extras, maxSeats, projSeats, Enrolled, Waitlist, Crosslist, StartDate, EndDate, Building, Room, RoomCap);
                            tmpClass.isCrossFirst = isCrossFirst;
                            // check if it is a deleted class
                            if (worksheet.Cells[i, indexTerm].Font.Strikethrough == false)
                            {
                                bool dontAdd = false;
                              
                                if (tmpClass.CRN == "-1")
                                {
                                    for (int j = 0; j < classList.Count; j++)
                                    {
                                        if (classList[j].ClassName == tmpClass.ClassName && tmpClass.SectionNumber == classList[j].SectionNumber && tmpClass.Term == classList[j].Term)
                                        {
                                            errorMSG = "Duplicate Class " + tmpClass.ClassName + " Found, only one instance of this class was added to the class list";
                                            WriteErrorLog(errorMSG);
                                            dontAdd = true;
                                        }
                                    }
                                }
                                
                                if (dontAdd == false)
                                {
                                    classList.Add(tmpClass);
                                }
                            }
                            else
                            {
                                deletedClasses.Add(tmpClass);
                            }

                        }
                        catch (Exception ex)
                        {
                            System.Windows.MessageBox.Show("There has been an issue when attempting to add a class to the classlist. Please refer to the user manual for an example excel file that is able to be accepted by this program.");
                            errorMSG = "There has been an issue when attempting to add a class to the classlist. Please refer to the user manual for an example excel file that is able to be accepted by this program.";
                            System.Windows.MessageBox.Show("Please select a new file or exit the program.");
                            WriteErrorLog(errorMSG);
                            SaveErrorHistory();
                            oExcel.Workbooks.Close();
                            System.Windows.Forms.Application.Restart();
                            //Helper.CloseUniqueWindow<FileSelect>(); // Close File Select window

                            System.Environment.Exit(0);
                        }
                    }
                }

                for (int i = 0; i < classList.Count(); i++)
                {
                    for (int j = 0; j < i; j++)
                    {
                        if (classList[i].DeptName != classList[j].DeptName)
                        {
                            if (classList[j].DeptName == "CYBR" && classList[i].DeptName == "CPSC")
                            {
                                Classes tmp = classList[i];
                                classList[i] = classList[j];
                                classList[j] = tmp;
                            }
                        }
                        else if (classList[i].ClassNumber != classList[j].ClassNumber)
                        {
                            if (classList[j].ClassNumber > classList[i].ClassNumber)
                            {
                                Classes tmp = classList[i];
                                classList[i] = classList[j];
                                classList[j] = tmp;
                            }
                        }
                        else
                        {
                            if (classList[j].SectionNumber > classList[i].SectionNumber)
                            {
                                Classes tmp = classList[i];
                                classList[i] = classList[j];
                                classList[j] = tmp;
                            }
                        }
                    }
                }

                try
                {
                    CheckRoomInfo((columns - 1), file); // subtract 1 because we do not check notes
                    CheckProfInfo(columns, file);

                    string startRangeHeader = "A1"; // retrieve font of first header to apply to all headers
                    string endRangeHeader = "A1";

                    Excel.Range currentRangeHeader = (Excel.Range)worksheet.get_Range(startRangeHeader, endRangeHeader);

                    headerFont = currentRangeHeader.Font;
                    headerFontName = headerFont.Name;
                    headerFontStyle = headerFont.FontStyle;
                    headerFontSize = (int)headerFont.Size;
                    Excel.Range columnRow = range.Rows[2]; // retrieve font of first cell in each column



                    int nullFormatFix = 0;
                    for (int i = 1; i < excelHeaderList.Count; i++) // check to make sure we are getting cell format from a cell with values
                    {
                        if (columnRow.Cells[i].Value != null)
                        {
                            colFontName[i - 1] = columnRow.Cells[i].Font.Name;
                            colFontStyle[i - 1] = columnRow.Cells[i].Font.FontStyle;
                            colFontSize[i - 1] = (int)columnRow.Cells[i].Font.Size;
                            columnCount++;
                            nullFormatFix = i;
                        }
                        else
                        {
                            colFontName[i - 1] = columnRow.Cells[nullFormatFix].Font.Name;
                            colFontStyle[i - 1] = columnRow.Cells[nullFormatFix].Font.FontStyle;
                            colFontSize[i - 1] = (int)columnRow.Cells[nullFormatFix].Font.Size;
                            columnCount++;
                        }
                    }

                }
                catch (Exception ex)
                {
                    System.Windows.MessageBox.Show("Invalid Information contained within excel sheet. Not enough information is known to know where. Please refer to the user manual for an example excel file that is able to be accepted by this program.");
                    errorMSG = "Invalid Information contained within excel sheet.Not enough information is known to know where. Please refer to the user manual for an example excel file that is able to be accepted by this program.";

                    System.Windows.MessageBox.Show("Please select a new file or exit the program.");
                    WriteErrorLog(errorMSG);
                    SaveErrorHistory();
                    oExcel.Workbooks.Close();
                    System.Windows.Forms.Application.Restart();
                    //Helper.CloseUniqueWindow<FileSelect>(); // Close File Select window

                    System.Environment.Exit(0);


                }
                oExcel.Workbooks.Close();

            }
        }
        private void InitializeErrorLog()
        {
            string fileName = "ErrorLog.txt";
            FileInfo f = new FileInfo(fileName);
            string name = f.FullName;



            // write current date/time to Error Log and replace old information
            StreamWriter tw = new StreamWriter(name);
            tw.WriteLine(DateTime.Now);
            tw.Close();

        }

        private void ErrorLog()
        {
            String line;
            String errors;
            // create new Stack panel
            StackPanel sp = new StackPanel() { Orientation = System.Windows.Controls.Orientation.Vertical };
            string fileName = "ErrorLog.txt";
            FileInfo f = new FileInfo(fileName);
            string name = f.FullName;



            StreamReader sr = new StreamReader(name);

            //Create new Scroll viewer
            ScrollViewer sv = new ScrollViewer()
            {
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                Height = 700,
            };

            try
            {

                line = sr.ReadLine();
                errors = line;
                //Continue to read until you reach end of file
                while (line != null)
                {


                    //Read the next line
                    //Add the controls to stack panel

                    line = sr.ReadLine();
                    errors = errors + "@" + line;

                }
                errors = errors.Replace("@", System.Environment.NewLine);
                sv.Content = errors;

                theStackPanel.Children.Clear();

                sp.Children.Add(sv);

                theStackPanel.Children.Add(sp);
                //close the file
                sr.Close();

            }
            catch (Exception e)
            {
                Console.WriteLine("Exception: " + e.Message);
            }
            finally
            {
                Console.WriteLine("Executing finally block.");
            }


        }


        //saves error log to error log history
        public void SaveErrorHistory()
        {
            {
                String line;
                String errors;


                string fileName = "ErrorLog.txt";
                FileInfo f = new FileInfo(fileName);
                string name = f.FullName;

                string fileName2 = "ErrorLogHistory.txt";
                FileInfo f2 = new FileInfo(fileName2);
                string name2 = f2.FullName;


                StreamReader sr = new StreamReader(name);
                try
                {

                    line = sr.ReadLine();
                    errors = line;
                    //Continue to read until you reach end of file
                    while (line != null)
                    {


                        //Read the next line
                        //Add the controls to stack panel

                        line = sr.ReadLine();
                        errors = errors + "@" + line;

                    }
                    errors = errors.Replace("@", System.Environment.NewLine);

                    // write current date/time to Error Log
                    StreamWriter tw = File.AppendText(name2);
                    tw.WriteLine("------");
                    tw.WriteLine(errors);
                    tw.Close();

                    //close the file
                    sr.Close();

                }
                catch (Exception e)
                {
                    Console.WriteLine("Exception: " + e.Message);
                }
                finally
                {
                    Console.WriteLine("Executing finally block.");
                }


            }
        }

        public void WriteErrorLog(String error)
        {
            string fileName = "ErrorLog.txt";
            FileInfo f = new FileInfo(fileName);
            string name = f.FullName;

            StreamWriter tw = File.AppendText(name);
            tw.WriteLine(error);
            tw.Close();

            errorCount++;
            ErrorLog();
        }

        public void startupErrors()
        {
            if (errorCount > 0)
            {
                System.Windows.MessageBox.Show("There have been conflicts upon start up. Please refer to the error log for more information");
            }
        }


        public void DrawTimeTables() // Calls TimeTableSetup() for MWF and TR 
        {
            TimeTableSetup(MWF, times_MWF);
            TimeTableSetup(TR, times_TR);
        }
        public void TimeTableSetup(Grid parentGrid, Timeslot[] times) // Creates an empty GUI grid based on timeslots + classrooms, then calls PopulateTimeTable() 
        {
            String parentName = parentGrid.Name; // Used to uniquely identify the timeslots
            Grid timeTable = new Grid();
            string timeTableName = parentGrid.Name + "_";
            timeTable.Name = timeTableName;
            timeTable.SetValue(Grid.RowProperty, 1);
            timeTable.SetValue(Grid.ColumnProperty, 1);
            timeTable.MinHeight = 450;
            timeTable.MinWidth = 450;
            timeTable.VerticalAlignment = VerticalAlignment.Stretch;
            timeTable.HorizontalAlignment = System.Windows.HorizontalAlignment.Stretch;
            int rowLength = 0;

            //timeTable.ShowGridLines = true; // Uncomment for debugging (Shows gridlines)

            for (int j = 0; j < classrooms.Count; j++)  // Organize classrooms by room number lowest to highest
            {
                ClassRoom tmp;
                for (int k = j + 1; k < classrooms.Count; k++)
                {
                    if (classrooms[j].RoomNum > classrooms[k].RoomNum)
                    {
                        tmp = classrooms[k];
                        classrooms[k] = classrooms[j];
                        classrooms[j] = tmp;
                    }

                }

            }
            for (int j = 0; j < times_Default.Count; j++) // Organize default times from earliest to latest
            {
                Timeslot tmp;
                int tmpR;
                string tmpT;

                for (int k = j + 1; k < times_Default.Count; k++)
                {

                    string compareTime = times_Default[j].Time;
                    string compareTimes = times_Default[k].Time;
                    int compareTimeI = Int32.Parse(compareTime.Split(':')[0]);
                    int compareTimesI = Int32.Parse(compareTimes.Split(':')[0]);
                    if (times_Default[j].Meridian == "PM" && times_Default[k].Meridian == "AM")
                    {
                        tmp = times_Default[k];
                        times_Default[k] = times_Default[j];
                        times_Default[j] = tmp;

                        tmpR = times_Default_Room[k];
                        times_Default_Room[k] = times_Default_Room[j];
                        times_Default_Room[j] = tmpR;

                        tmpT = times_Default_Timetable[k];
                        times_Default_Timetable[k] = times_Default_Timetable[j];
                        times_Default_Timetable[j] = tmpT;
                    }
                    if (times_Default[j].Meridian == "AM" && times_Default[k].Meridian == "AM")
                    {
                        if (compareTimeI > compareTimesI)
                        {
                            tmp = times_Default[k];
                            times_Default[k] = times_Default[j];
                            times_Default[j] = tmp;

                            tmpR = times_Default_Room[k];
                            times_Default_Room[k] = times_Default_Room[j];
                            times_Default_Room[j] = tmpR;

                            tmpT = times_Default_Timetable[k];
                            times_Default_Timetable[k] = times_Default_Timetable[j];
                            times_Default_Timetable[j] = tmpT;
                        }
                    }
                    if (times_Default[j].Meridian == "PM" && times_Default[k].Meridian == "PM")
                    {
                        if (compareTimesI == 12 && compareTimeI != 12)
                        {
                            tmp = times_Default[k];
                            times_Default[k] = times_Default[j];
                            times_Default[j] = tmp;

                            tmpR = times_Default_Room[k];
                            times_Default_Room[k] = times_Default_Room[j];
                            times_Default_Room[j] = tmpR;

                            tmpT = times_Default_Timetable[k];
                            times_Default_Timetable[k] = times_Default_Timetable[j];
                            times_Default_Timetable[j] = tmpT;
                        }
                        if (compareTimeI != 12 && compareTimesI != 12)
                        {
                            if (compareTimeI > compareTimesI)
                            {
                                tmp = times_Default[k];
                                times_Default[k] = times_Default[j];
                                times_Default[j] = tmp;

                                tmpR = times_Default_Room[k];
                                times_Default_Room[k] = times_Default_Room[j];
                                times_Default_Room[j] = tmpR;

                                tmpT = times_Default_Timetable[k];
                                times_Default_Timetable[k] = times_Default_Timetable[j];
                                times_Default_Timetable[j] = tmpT;
                            }
                        }
                    }

                }
            }

            /*
            for (int q = 0; q < times_Default.Count; q++) // determine number of rows for each column
            {
                string time = times_Default[q].Time;
                int room = times_Default_Room[q];
                string table = times_Default_Timetable[q];
            }
            var dict = new Dictionary<int, int>();

            foreach (int roomNum in times_Default_Room)
            {
                dict.TryGetValue(roomNum, out int roomCount);
                dict[roomNum] = roomCount + 1;
            }
            foreach (var pair in dict)
            {
                defaultRowCountMWF = 1;
                defaultRowCountTR = 1;

                for (int g = 0; g < classrooms.Count; g++)
                {

                    if (classrooms[g].RoomNum == pair.Key)
                    {
                        for (int t = 0; t < times_Default.Count; t++)
                        {
                            if (times_Default_Room[t] == classrooms[g].RoomNum)
                            {
                                if (times_Default_Timetable[t] == "MWF")
                                {
                                    times_Default_Row[t] = defaultRowCountMWF;

                                    defaultRowCountMWF++;
                                }
                                else if (times_Default_Timetable[t] == "TR")
                                {
                                    times_Default_Row[t] = defaultRowCountTR;
                                    defaultRowCountTR++;
                                }
                            }
                        }
                    }
                }

            }
            */
            for (int i = 0; i < times_Default.Count(); i++)
            {

                for (int j = 0; j < times.Length; j++)
                {
                    if (times_Default[i].Start == times[j].Start)
                    {
                        times_Default_Row.Add(j + 1);
                    }
                }
            }

            DeleteDuplicateTimes();

            if (parentGrid.Name == "MWF")
            {
                rowLength = defaultRowCountMWF;
            }
            else
            {
                rowLength = defaultRowCountTR + 1;
            }

            // make a row for each timeslot
            for (int i = 0; i <= times.Length + 1; i++)
            {
                if (i == 0)
                {
                    timeTable.RowDefinitions.Add(new RowDefinition() { Height = new GridLength(1.2, GridUnitType.Star) });
                }
                else if (i <= times.Length)
                {
                    timeTable.RowDefinitions.Add(new RowDefinition() { Height = new GridLength(2, GridUnitType.Star) });
                }
                else if (i == times.Length + 4)
                {
                    timeTable.RowDefinitions.Add(new RowDefinition() { Height = new GridLength(2, GridUnitType.Star) });
                }
                else
                {
                    timeTable.RowDefinitions.Add(new RowDefinition() { Height = new GridLength(2, GridUnitType.Star) });
                }
            }
            // make a column for each classroom
            for (int i = 0; i < classrooms.Count * 2; i++)
            {
                if (i == 0)
                {
                    timeTable.ColumnDefinitions.Add(new ColumnDefinition() { Width = new GridLength(1.2, GridUnitType.Star) });
                }
                else
                {
                    timeTable.ColumnDefinitions.Add(new ColumnDefinition() { Width = new GridLength(2, GridUnitType.Star) });
                }
            }
            // fill the grid
            for (int i = 0; i <= times.Length; i++)
            {
                // Add row titles (Time Periods)
                if (i != 0)
                {


                    // Buttons with times for classes

                    for (int j = 0; j < classrooms.Count * 2; j++)
                    {
                        if (j == 0 || j % 2 == 0)
                        {

                            if (i <= times.Length)
                            {
                                System.Windows.Controls.Button timeLabelButton = new System.Windows.Controls.Button();
                                timeLabelButton.Content = times[i - 1].Time + " " + times[i - 1].Meridian;
                                timeLabelButton.SetValue(Grid.RowProperty, i);
                                timeLabelButton.SetValue(Grid.ColumnProperty, j);
                                timeLabelButton.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                timeLabelButton.VerticalAlignment = VerticalAlignment.Center;
                                timeLabelButton.Click += new RoutedEventHandler(ChangeTime_Click);
                                timeTable.Children.Add(timeLabelButton);

                                System.Windows.Controls.Button timeLabelButtonAdd = new System.Windows.Controls.Button();
                                timeLabelButtonAdd.Content = "Add Row";
                                timeLabelButtonAdd.SetValue(Grid.RowProperty, i + 1);
                                timeLabelButtonAdd.SetValue(Grid.ColumnProperty, j);
                                timeLabelButtonAdd.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                timeLabelButtonAdd.VerticalAlignment = VerticalAlignment.Center;
                                timeLabelButtonAdd.Background = new SolidColorBrush(Colors.Thistle);
                                timeLabelButtonAdd.Click += new RoutedEventHandler(AddTimeslotsClick);
                                timeTable.Children.Add(timeLabelButtonAdd);
                            }
                            else
                            {
                                /*
                                System.Windows.Controls.Button timeLabelButton = new System.Windows.Controls.Button();
                                timeLabelButton.Content = times[0].Time + " " + times[0].Meridian;
                                timeLabelButton.SetValue(Grid.RowProperty, i);
                                timeLabelButton.SetValue(Grid.ColumnProperty, j);
                                timeLabelButton.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                timeLabelButton.VerticalAlignment = VerticalAlignment.Center;
                                timeLabelButton.Click += new RoutedEventHandler(ChangeTime_Click);
                                timeTable.Children.Add(timeLabelButton);
                                
                                System.Windows.Controls.Button timeLabelButtonAdd = new System.Windows.Controls.Button();
                                timeLabelButtonAdd.Content = "Add Row";
                                timeLabelButtonAdd.SetValue(Grid.RowProperty, i + 1);
                                timeLabelButtonAdd.SetValue(Grid.ColumnProperty, j);
                                timeLabelButtonAdd.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                timeLabelButtonAdd.VerticalAlignment = VerticalAlignment.Center;
                                timeLabelButtonAdd.Background = new SolidColorBrush(Colors.Thistle);
                                timeLabelButtonAdd.Click += new RoutedEventHandler(AddTimeslotsClick);
                                timeTable.Children.Add(timeLabelButtonAdd);
                                */
                            }

                        }

                    }

                }
                for (int n = 1; n <= classrooms.Count; n++)
                {
                    if (i == 0) // Add column titles (Classroom Bldg-Number)
                    {


                        if (n == 1)
                        {
                            System.Windows.Controls.Label classLabel = new System.Windows.Controls.Label();
                            classLabel.Content = classrooms[n - 1].Location + "-" + classrooms[n - 1].RoomNum;
                            classLabel.SetValue(Grid.RowProperty, 0);
                            classLabel.SetValue(Grid.ColumnProperty, n);
                            classLabel.HorizontalAlignment = System.Windows.HorizontalAlignment.Center;
                            classLabel.VerticalAlignment = VerticalAlignment.Center;
                            timeTable.Children.Add(classLabel);

                            System.Windows.Controls.Button labelBtn_Click = new System.Windows.Controls.Button();
                            labelBtn_Click.Content = classLabel.Content;
                            labelBtn_Click.SetValue(Grid.RowProperty, 0);
                            labelBtn_Click.SetValue(Grid.ColumnProperty, n);
                            labelBtn_Click.HorizontalAlignment = System.Windows.HorizontalAlignment.Center;
                            labelBtn_Click.VerticalAlignment = VerticalAlignment.Center;
                            labelBtn_Click.Background = Brushes.DarkSeaGreen;
                            labelBtn_Click.BorderThickness = new Thickness(1, 1, 1, 1);
                            labelBtn_Click.BorderBrush = new SolidColorBrush(Color.FromRgb(0, 0, 0));
                            labelBtn_Click.Click += new RoutedEventHandler(ViewRoomInfo_Click);

                            timeTable.Children.Add(labelBtn_Click);
                        }
                        else if (n >= 2)
                        {
                            System.Windows.Controls.Label classLabel = new System.Windows.Controls.Label();
                            classLabel.Content = classrooms[n - 1].Location + "-" + classrooms[n - 1].RoomNum;
                            classLabel.SetValue(Grid.RowProperty, 0);
                            classLabel.SetValue(Grid.ColumnProperty, n + n - 1);
                            classLabel.HorizontalAlignment = System.Windows.HorizontalAlignment.Center;
                            classLabel.VerticalAlignment = VerticalAlignment.Center;
                            timeTable.Children.Add(classLabel);

                            System.Windows.Controls.Button labelBtn_Click = new System.Windows.Controls.Button();
                            labelBtn_Click.Content = classLabel.Content;
                            labelBtn_Click.SetValue(Grid.RowProperty, 0);
                            labelBtn_Click.SetValue(Grid.ColumnProperty, n + n - 1);
                            labelBtn_Click.HorizontalAlignment = System.Windows.HorizontalAlignment.Center;
                            labelBtn_Click.VerticalAlignment = VerticalAlignment.Center;
                            labelBtn_Click.Background = Brushes.DarkSeaGreen;
                            labelBtn_Click.BorderThickness = new Thickness(1, 1, 1, 1);
                            labelBtn_Click.BorderBrush = new SolidColorBrush(Color.FromRgb(0, 0, 0));
                            labelBtn_Click.Click += new RoutedEventHandler(ViewRoomInfo_Click);

                            timeTable.Children.Add(labelBtn_Click);
                        }

                    }
                    else // Add empty timeslots
                    {
                        if (i <= times.Length)
                        {

                            System.Windows.Controls.Label emptySlot = new System.Windows.Controls.Label();
                            string lbl_name = timeTable.Name + times[i - 1].TimeID + "_" + classrooms[n - 1].ClassID + times[i - 1].Meridian; // IMPORTANT: Empty timeslots naming convention = GridName_TimeID_ClassID_Meridian
                            emptySlot.Name = lbl_name;
                            //MessageBox.Show(emptySlot.Name); // DEBUG
                            emptySlot.Content = "";
                            emptySlot.AllowDrop = true;
                            emptySlot.Drop += new System.Windows.DragEventHandler(HandleDropToCell);
                            emptySlot.Style = Resources["DragLabel"] as System.Windows.Style;
                            emptySlot.SetValue(Grid.RowProperty, i);
                            emptySlot.SetValue(Grid.ColumnProperty, n + n - 1);
                            emptySlot.HorizontalContentAlignment = System.Windows.HorizontalAlignment.Center;
                            emptySlot.VerticalContentAlignment = VerticalAlignment.Center;
                            emptySlot.BorderThickness = new Thickness(1, 1, 1, 1);
                            emptySlot.BorderBrush = new SolidColorBrush(Color.FromRgb(0, 0, 0));
                            emptySlot.MinWidth = 75;
                            emptySlot.Background = new SolidColorBrush(Color.FromRgb(255, 255, 255));
                            emptySlot.Margin = new Thickness(5);
                            emptySlot.ContextMenu = null;
                            //emptySlot.ContextMenu = Resources["ClassContextMenu"] as System.Windows.Controls.ContextMenu;
                            object o = FindName(lbl_name);
                            if (o != null)
                            {
                                UnregisterName(lbl_name);
                            }
                            RegisterName(lbl_name, emptySlot);
                            timeTable.Children.Add(emptySlot);
                        }
                        else 
                        {
                            System.Windows.Controls.Label emptySlot = new System.Windows.Controls.Label();
                            string lbl_name = timeTable.Name + times[0].TimeID + "_" + classrooms[n - 1].ClassID + times[0].Meridian; // IMPORTANT: Empty timeslots naming convention = GridName_TimeID_ClassID
                            emptySlot.Name = lbl_name;
                            //MessageBox.Show(emptySlot.Name); // DEBUG
                            emptySlot.Content = "";
                            emptySlot.AllowDrop = true;
                            emptySlot.Drop += new System.Windows.DragEventHandler(HandleDropToCell);
                            emptySlot.Style = Resources["DragLabel"] as System.Windows.Style;
                            emptySlot.SetValue(Grid.RowProperty, i);
                            emptySlot.SetValue(Grid.ColumnProperty, n + n - 1);
                            emptySlot.HorizontalContentAlignment = System.Windows.HorizontalAlignment.Center;
                            emptySlot.VerticalContentAlignment = VerticalAlignment.Center;
                            emptySlot.BorderThickness = new Thickness(1, 1, 1, 1);
                            emptySlot.BorderBrush = new SolidColorBrush(Color.FromRgb(0, 0, 0));
                            emptySlot.MinWidth = 75;
                            emptySlot.Background = new SolidColorBrush(Color.FromRgb(255, 255, 255));
                            emptySlot.Margin = new Thickness(5);
                            emptySlot.ContextMenu = null;
                            //emptySlot.ContextMenu = Resources["ClassContextMenu"] as ContextMenu;
                            object o = FindName(lbl_name);
                            if (o != null)
                            {
                                UnregisterName(lbl_name);
                            }
                            RegisterName(lbl_name, emptySlot);
                            timeTable.Children.Add(emptySlot);
                        }

                    }
                }
            }

            // Add the grid to the MWF_Schedule Grid
            object x = FindName(timeTableName);
            if (x != null)
            {
                UnregisterName(timeTableName);
            }
            RegisterName(timeTableName, timeTable);
            parentGrid.Children.Add(timeTable);

            // Populate the empty timeslots with our available information
            PopulateTimeTable(timeTable, times);

        }

        public void AddTimeslotsClick(object sender, RoutedEventArgs e)
        {
            
            
            System.Windows.Application.Current.Resources["Set_ChangeTimeslots_Success"] = false;

            System.Windows.Application.Current.Resources["Set_min"] = times_MWF.Length;

            Unfocus_Overlay.Visibility = Visibility.Visible;
            AddRowDialog addRowsDialog = new AddRowDialog();
            addRowsDialog.Owner = this;
            addRowsDialog.ShowDialog();
            Unfocus_Overlay.Visibility = Visibility.Hidden;


            if (System.Windows.Application.Current.Resources["Set_ChangeTimeslots_Success"] != null && (bool)System.Windows.Application.Current.Resources["Set_ChangeTimeslots_Success"] == true)
            {


                int rowAdd = Int32.Parse(System.Windows.Application.Current.Resources["Set_rows"].ToString());
                int timeTable = Int32.Parse(System.Windows.Application.Current.Resources["Set_TimeTable"].ToString());

                
                string start = "09:00";
                string end = "10:00";
                string endTimePure = "09:50";
                string meridian = "PM";
                int row = Grid.GetRow((System.Windows.Controls.Button)sender);
                int column = Grid.GetColumn((System.Windows.Controls.Button)sender);
                int timeChange = 0;
                


                if (timeTable == 1)
                {
                    start = "09:00";
                    end = "10:30";
                    endTimePure = "10:15";
                    meridian = "PM";
                    row = Grid.GetRow((System.Windows.Controls.Button)sender);
                    column = Grid.GetColumn((System.Windows.Controls.Button)sender);
                    timeChange = 0;
                }
                else
                {
                    start = "09:00";
                    end = "10:00";
                    endTimePure = "09:50";
                    meridian = "PM";
                    row = Grid.GetRow((System.Windows.Controls.Button)sender);
                    column = Grid.GetColumn((System.Windows.Controls.Button)sender);
                    timeChange = 0;
                }




                // Remove old Grids
                Grid child = FindName("MWF_") as Grid;
                MWF.Children.Remove(child);
                Grid child2 = FindName("TR_") as Grid;
                TR.Children.Remove(child2);
                // Redraw Grids

                DrawTimeTablesDynamic(start, end, meridian, row, column, timeTable, timeChange, endTimePure, rowAdd);

            }
        }
        public void DrawTimeTablesDynamic(string start, string end, string meridian, int row, int column, int timeTable, int timeChange, string endTimePure, int rowAdd) // Calls TimeTableSetup() for MWF and TR 
        {
            TimeTableSetupDynamic(MWF, times_MWF, start, end, meridian, row, column, timeTable, timeChange, endTimePure, rowAdd);
            TimeTableSetupDynamic(TR, times_TR, start, end, meridian, row, column, timeTable, timeChange, endTimePure, rowAdd);
        }
        public void TimeTableSetupDynamic(Grid parentGrid, Timeslot[] times, string startTime, string endTime, string meridian, int row, int column, int timeTableD, int timeChange, string endTimePure, int rowAdd) // Creates an empty GUI grid based on timeslots + classrooms, then calls PopulateTimeTable() 
        {

            String parentName = parentGrid.Name; // Used to uniquely identify the timeslots
            Grid timeTable = new Grid();
            int c = 0; // creates loop to check all changed class times are in their positions
            int increment = INCREMENT; // applys default increment to automatically changing class times

            if (parentGrid.Name == "TR")
            {
                increment = INCREMENTTR; // applys default increment to automatically changing class times
            }
            
                
            
            string timeTableName = parentGrid.Name + "_";
            timeTable.Name = timeTableName;
            timeTable.SetValue(Grid.RowProperty, 1);
            timeTable.SetValue(Grid.ColumnProperty, 1);
            timeTable.MinHeight = 450;
            timeTable.MinWidth = 450;
            timeTable.VerticalAlignment = VerticalAlignment.Stretch;
            timeTable.HorizontalAlignment = System.Windows.HorizontalAlignment.Stretch;
            int rowLength = 0;
            bool meridianCheck = false;
            List<string> changedClasses = new List<string>();
            //timeTable.ShowGridLines = true; // Uncomment for debugging (Shows gridlines)
            int checkAddRowButton = 0;

            rowLength = times.Length;

            if (rowAdd != 0)
            {
                rowLength = rowAdd;
            }

            for (int j = 0; j < classrooms.Count; j++)  // Organize classrooms by room number lowest to highest
            {
                ClassRoom tmp;
                for (int k = j + 1; k < classrooms.Count; k++)
                {
                    if (classrooms[j].RoomNum > classrooms[k].RoomNum)
                    {
                        tmp = classrooms[k];
                        classrooms[k] = classrooms[j];
                        classrooms[j] = tmp;
                    }

                }

            }
            for (int p = 0; p < labelCount; p++)
            {

                for (int q = 0; q < p; q++)
                {
                    if (labelRow[p] == labelRow[q])
                    {
                        if (labelColumn[p] == labelColumn[q])
                        {
                            if (labelTimeTable[p] == labelTimeTable[q])
                            {
                               
                                labelRow.RemoveAt(q);
                                labelColumn.RemoveAt(q);
                                labelNames.RemoveAt(q);
                                labelTimeTable.RemoveAt(q);
                                labelCount--;

                            }
                        }
                    }

                }

            }

            for (int i = 0; i < times_Default.Count(); i++) // make sure default times have correct row for changing class times
            {

                for (int j = 0; j < times.Length; j++)
                {
                    if (times_Default[i].Start == times[j].Start)
                    {
                        times_Default_Row[i] = j + 1;
                    }
                }
            }

            // make a row for each timeslot
            for (int i = 0; i <= rowLength + 1; i++)
            {
                if (i == 0)
                {
                    timeTable.RowDefinitions.Add(new RowDefinition() { Height = new GridLength(1.2, GridUnitType.Star) });
                }
                else
                {
                    timeTable.RowDefinitions.Add(new RowDefinition() { Height = new GridLength(2, GridUnitType.Star) });
                }
            }
            // make a column for each classroom
            for (int i = 0; i < classrooms.Count * 2; i++)
            {
                if (i == 0)
                {
                    timeTable.ColumnDefinitions.Add(new ColumnDefinition() { Width = new GridLength(1.2, GridUnitType.Star) });
                }
                else
                {
                    timeTable.ColumnDefinitions.Add(new ColumnDefinition() { Width = new GridLength(2, GridUnitType.Star) });
                }
            }
            // fill the grid
            for (int i = 0; i <= rowLength; i++)
            {
                // Add row titles (Time Periods)
                if (i != 0)
                {

                    // Buttons with times for classes

                    for (int j = 0; j < classrooms.Count * 2; j++)
                    {
                        if (j == 0 || j % 2 == 0)
                        {
                            checkAddRowButton = 0; // makes sure add row button is always after the last row

                            if (count != 0)
                            {
                                if (row == -1 && column == -1)
                                {
                                    if (i <= rowLength && parentGrid.Name == "MWF")
                                    {
                                        if (i <= times.Length)
                                        {
                                            System.Windows.Controls.Button timeLabelButton = new System.Windows.Controls.Button();
                                            timeLabelButton.Content = times[i-1].Time + " " + times[i-1].Meridian;
                                            timeLabelButton.Name = timeTable.Name + times[i - 1].TimeID + "_" + classrooms[j / 2].ClassID + times[i - 1].Meridian;
                                            timeLabelButton.SetValue(Grid.RowProperty, i);
                                            timeLabelButton.SetValue(Grid.ColumnProperty, j);
                                            timeLabelButton.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                            timeLabelButton.VerticalAlignment = VerticalAlignment.Center;
                                            timeLabelButton.Click += new RoutedEventHandler(ChangeTime_Click);
                                            timeTable.Children.Add(timeLabelButton);
                                            labelNames.Add(timeLabelButton.Name);
                                            labelColumn.Add(j + 1);
                                            labelRow.Add(i);
                                            labelTimeTable.Add(parentGrid.Name);
                                            labelCount++;
                                        }

                                    }
                                    else if (i <= rowLength && parentGrid.Name == "TR")
                                    {
                                        if (i <= times.Length)
                                        {
                                            System.Windows.Controls.Button timeLabelButton = new System.Windows.Controls.Button();
                                            timeLabelButton.Content = times[i - 1].Time + " " + times[i - 1].Meridian;
                                            timeLabelButton.Name = timeTable.Name + times[i - 1].TimeID + "_" + classrooms[j / 2].ClassID + times[i - 1].Meridian;
                                            timeLabelButton.SetValue(Grid.RowProperty, i);
                                            timeLabelButton.SetValue(Grid.ColumnProperty, j);
                                            timeLabelButton.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                            timeLabelButton.VerticalAlignment = VerticalAlignment.Center;
                                            timeLabelButton.Click += new RoutedEventHandler(ChangeTime_Click);
                                            timeTable.Children.Add(timeLabelButton);
                                            labelNames.Add(timeLabelButton.Name);
                                            labelColumn.Add(j+1);
                                            labelRow.Add(i);
                                            labelTimeTable.Add(parentGrid.Name);
                                            labelCount++;
                                        }
                                    }
                                }

                                else if (i < row || j != column) //applies all defaults to grid
                                {
                                    if (parentGrid.Name == "MWF" && timeTableD == 0 || parentGrid.Name == "TR" && timeTableD == 1 || parentGrid.Name != "MWF" && timeTableD == 0 || parentGrid.Name != "TR" && timeTableD == 1)
                                    {
                                        if (i <= times.Length)
                                        {
                                            System.Windows.Controls.Button timeLabelButton = new System.Windows.Controls.Button();
                                            timeLabelButton.Name = timeTable.Name + times[i - 1].TimeID + "_" + classrooms[j / 2].ClassID + times[i - 1].Meridian;
                                            timeLabelButton.Content = times[i - 1].Time + " " + times[i - 1].Meridian;

                                            timeLabelButton.SetValue(Grid.RowProperty, i);
                                            timeLabelButton.SetValue(Grid.ColumnProperty, j);
                                            timeLabelButton.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                            timeLabelButton.VerticalAlignment = VerticalAlignment.Center;
                                            timeLabelButton.Click += new RoutedEventHandler(ChangeTime_Click);
                                            timeTable.Children.Add(timeLabelButton);
                                            bool labelCheck = false;
                                            for (int l = 0; l < labelNames.Count(); l++)
                                            {
                                                if (labelTimeTable[l] == parentGrid.Name)
                                                {
                                                    if (labelColumn[l] == j + 1)
                                                    {
                                                        if (labelRow[l] == i)
                                                        {
                                                            labelCheck = true;
                                                        }
                                                    }
                                                }
                                            }

                                            if (labelCheck == false)
                                            {
                                                labelNames.Add(timeLabelButton.Name);
                                                labelColumn.Add(j + 1);
                                                labelRow.Add(i);
                                                labelTimeTable.Add(parentGrid.Name);
                                                labelCount++;
                                            }
                                            System.Windows.Controls.Button timeLabelButtonAdd = new System.Windows.Controls.Button();
                                            timeLabelButtonAdd.Content = "Add Row";
                                            timeLabelButtonAdd.SetValue(Grid.RowProperty, i + 1);
                                            timeLabelButtonAdd.SetValue(Grid.ColumnProperty, j);
                                            timeLabelButtonAdd.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                            timeLabelButtonAdd.VerticalAlignment = VerticalAlignment.Center;
                                            timeLabelButtonAdd.Background = new SolidColorBrush(Colors.Thistle);
                                            timeLabelButtonAdd.Click += new RoutedEventHandler(AddTimeslotsClick);
                                            timeTable.Children.Add(timeLabelButtonAdd);
                                        

                                        }

                                    }
                                }

                                else if (i == row && j == column) // applys currently changed grid slot to grid
                                {

                                    if (parentGrid.Name == "MWF" && timeTableD == 0 || parentGrid.Name == "TR" && timeTableD == 1)
                                    {

                                        System.Windows.Controls.Button timeLabelButton = new System.Windows.Controls.Button();
                                        timeLabelButton.Content = startTime + " " + meridian;

                                        timeLabelButton.SetValue(Grid.RowProperty, i);
                                        timeLabelButton.SetValue(Grid.ColumnProperty, j);
                                        timeLabelButton.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                        timeLabelButton.VerticalAlignment = VerticalAlignment.Center;
                                        timeLabelButton.Click += new RoutedEventHandler(ChangeTime_Click);
                                        timeTable.Children.Add(timeLabelButton);
                                        changedRoomNum.Add(classrooms[column / 2].RoomNum);
                                        changedStartTime.Add(startTime);
                                        changedEndTime.Add(endTimePure);
                                        changedMeridian.Add(meridian);
                                        changedColumn.Add(column);
                                        changedRow.Add(row);
                                        changedTimeTable.Add(timeTableD);
                                        count++;

                                        for (int q = 0; q < classList.Count(); q++) // Changes class to have stating and ending time equal to user made changes on button label
                                        {
                                            
                                            bool breakLoop = false;
                                            bool changeClass = true;
                                            string time = "";
                                            if (classList[q].StartTime.Start.Length > 4)
                                            {
                                                time = classList[q].StartTime.Start.Substring(0, 5);
                                            }
                                            
                                            
                                            for (int d = 0; d < count; d++) //Checks changed times for current class time and changes it
                                                {
                                                if (classList[q].ClassDay == "MWF" && timeTableD == 0 && changedTimeTable[d] == 0 || classList[q].ClassDay == "TR" && timeTableD == 1 && changedTimeTable[d] == 1)
                                                {
                                                    if (classList[q].Classroom.RoomNum == classrooms[column / 2].RoomNum && changedRoomNum[d] == classList[q].Classroom.RoomNum && time != startTime)
                                                    {
                                                        if (changedRow[d] == i)
                                                        {
                                                            if (changedStartTime[d] == time && changedMeridian[d] == classList[q].StartTime.Meridian)
                                                            {
                                                                if (time != startTime || time == startTime && meridian != classList[q].StartTime.Meridian)
                                                                {
                                                                    for (int k = 0; k < changedClasses.Count(); k++)
                                                                    {
                                                                        if (changedClasses[k] == classList[q].TextBoxName)
                                                                        {
                                                                            changeClass = false;
                                                                        }
                                                                    }
                                                                    if (changeClass == true)
                                                                    {
                                                                        classList[q].StartTime = new Timeslot(startTime, endTimePure, meridian);
                                                                        changedClasses.Add(classList[q].TextBoxName);
                                                                        breakLoop = true;
                                                                        break;
                                                                    }
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                            
                                            for (int d = 0; d < autoCount; d++) // checks automatically changed times for current class time and changes it
                                            {
                                                if (classList[q].ClassDay == "MWF" && timeTableD == 0 && autoChangedTimeTable[d] == 0 || classList[q].ClassDay == "TR" && timeTableD == 1 && autoChangedTimeTable[d] == 1)
                                                {
                                                    if (classList[q].Classroom.RoomNum == classrooms[column / 2].RoomNum && autoChangedRoomNum[d] == classList[q].Classroom.RoomNum && time != startTime)
                                                    {
                                                        if (autoChangedRow[d] == i)
                                                        {
                                                            if (autoChangedStartTime[d] == time && autoChangedMeridian[d] == classList[q].StartTime.Meridian)
                                                            {
                                                                if (time != startTime || time == startTime && meridian != classList[q].StartTime.Meridian)
                                                                {
                                                                    for (int k = 0; k < changedClasses.Count(); k++)
                                                                    {
                                                                        if (changedClasses[k] == classList[q].TextBoxName)
                                                                        {
                                                                            changeClass = false;
                                                                        }
                                                                    }
                                                                    if (changeClass == true)
                                                                    {
                                                                        classList[q].StartTime = new Timeslot(startTime, endTimePure, meridian);
                                                                        changedClasses.Add(classList[q].TextBoxName);
                                                                        breakLoop = true;
                                                                        break;
                                                                    }
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                            for (int d = 0; d < times_Default.Count(); d++) // checks default times for current class time and changes it
                                            {
                                                if (classList[q].ClassDay == "MWF" && timeTableD == 0 && times_Default_Timetable[d] == "MWF" || classList[q].ClassDay == "TR" && timeTableD == 1 && times_Default_Timetable[d] == "TR")
                                                {
                                                    if (classList[q].Classroom.RoomNum == classrooms[column / 2].RoomNum && times_Default_Room[d] == classList[q].Classroom.RoomNum)
                                                    {
                                                        if (times_Default_Row[d] == i)
                                                        {
                                                            if (times_Default[d].Time == classList[q].StartTime.Time && times_Default[d].Meridian == classList[q].StartTime.Meridian)
                                                            {
                                                                if (time != startTime || time == startTime && meridian != classList[q].StartTime.Meridian)
                                                                {
                                                                    for (int k = 0; k < changedClasses.Count(); k++)
                                                                    {
                                                                        if (changedClasses[k] == classList[q].TextBoxName)
                                                                        {
                                                                            changeClass = false;
                                                                        }
                                                                    }
                                                                    if (changeClass == true)
                                                                    {
                                                                        classList[q].StartTime = new Timeslot(startTime, endTimePure, meridian);
                                                                        changedClasses.Add(classList[q].TextBoxName);
                                                                        breakLoop = true;
                                                                        break;
                                                                    }
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                            }

                                            if (breakLoop == true)
                                            {
                                                break;
                                            }

                                        }

                                        char removeChar = ':';
                                        string[] timeLabel = startTime.Split(removeChar);
                                        string startTimeLabel = timeLabel[0];
                                        timeLabelButton.Name = timeTable.Name + startTimeLabel + "_" + classrooms[j / 2].ClassID + meridian;
                                        labelNames.Add(timeLabelButton.Name);
                                        labelColumn.Add(j + 1);
                                        labelRow.Add(i);
                                        labelTimeTable.Add(parentGrid.Name);
                                        labelCount++;
                                    }
                                    else
                                    {
                                        if (i <= rowLength && parentGrid.Name == "MWF")
                                        {
                                            if (i <= times.Length)
                                            {
                                                System.Windows.Controls.Button timeLabelButton = new System.Windows.Controls.Button();
                                                timeLabelButton.Content = times[i - 1].Time + " " + times[i - 1].Meridian;
                                                timeLabelButton.Name = timeTable.Name + times[i - 1].TimeID + "_" + classrooms[j / 2].ClassID + times[i - 1].Meridian;
                                                timeLabelButton.SetValue(Grid.RowProperty, i);
                                                timeLabelButton.SetValue(Grid.ColumnProperty, j);
                                                timeLabelButton.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                                timeLabelButton.VerticalAlignment = VerticalAlignment.Center;
                                                timeLabelButton.Click += new RoutedEventHandler(ChangeTime_Click);
                                                timeTable.Children.Add(timeLabelButton);

                                                bool labelCheck = false;
                                                for (int l = 0; l < labelNames.Count(); l++)
                                                {
                                                    if (labelTimeTable[l] == parentGrid.Name)
                                                    {
                                                        if (labelColumn[l] == j + 1)
                                                        {
                                                            if (labelRow[l] == i)
                                                            {
                                                                labelCheck = true;
                                                            }
                                                        }
                                                    }
                                                }

                                                if (labelCheck == false)
                                                {
                                                    labelNames.Add(timeLabelButton.Name);
                                                    labelColumn.Add(j + 1);
                                                    labelRow.Add(i);
                                                    labelTimeTable.Add(parentGrid.Name);
                                                    labelCount++;
                                                }



                                                System.Windows.Controls.Button timeLabelButtonAdd = new System.Windows.Controls.Button();
                                                timeLabelButtonAdd.Content = "Add Row";
                                                timeLabelButtonAdd.SetValue(Grid.RowProperty, i + 1);
                                                timeLabelButtonAdd.SetValue(Grid.ColumnProperty, j);
                                                timeLabelButtonAdd.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                                timeLabelButtonAdd.VerticalAlignment = VerticalAlignment.Center;
                                                timeLabelButtonAdd.Background = new SolidColorBrush(Colors.Thistle);
                                                timeLabelButtonAdd.Click += new RoutedEventHandler(AddTimeslotsClick);
                                                timeTable.Children.Add(timeLabelButtonAdd);
                                            }

                                        }
                                        else if (i <= rowLength && parentGrid.Name == "TR")
                                        {
                                            if (i <= times.Length)
                                            {
                                                System.Windows.Controls.Button timeLabelButton = new System.Windows.Controls.Button();
                                                timeLabelButton.Content = times[i - 1].Time + " " + times[i - 1].Meridian;
                                                timeLabelButton.Name = timeTable.Name + times[i - 1].TimeID + "_" + classrooms[j / 2].ClassID + times[i - 1].Meridian;
                                                timeLabelButton.SetValue(Grid.RowProperty, i);
                                                timeLabelButton.SetValue(Grid.ColumnProperty, j);
                                                timeLabelButton.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                                timeLabelButton.VerticalAlignment = VerticalAlignment.Center;
                                                timeLabelButton.Click += new RoutedEventHandler(ChangeTime_Click);
                                                timeTable.Children.Add(timeLabelButton);

                                                bool labelCheck = false;
                                                for (int l = 0; l < labelNames.Count(); l++)
                                                {
                                                    if (labelTimeTable[l] == parentGrid.Name)
                                                    {
                                                        if (labelColumn[l] == j + 1)
                                                        {
                                                            if (labelRow[l] == i)
                                                            {
                                                                labelCheck = true;
                                                            }
                                                        }
                                                    }
                                                }

                                                if (labelCheck == false)
                                                {
                                                    labelNames.Add(timeLabelButton.Name);
                                                    labelColumn.Add(j + 1);
                                                    labelRow.Add(i);
                                                    labelTimeTable.Add(parentGrid.Name);
                                                    labelCount++;
                                                }

                                                System.Windows.Controls.Button timeLabelButtonAdd = new System.Windows.Controls.Button();
                                                timeLabelButtonAdd.Content = "Add Row";
                                                timeLabelButtonAdd.SetValue(Grid.RowProperty, i + 1);
                                                timeLabelButtonAdd.SetValue(Grid.ColumnProperty, j);
                                                timeLabelButtonAdd.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                                timeLabelButtonAdd.VerticalAlignment = VerticalAlignment.Center;
                                                timeLabelButtonAdd.Background = new SolidColorBrush(Colors.Thistle);
                                                timeLabelButtonAdd.Click += new RoutedEventHandler(AddTimeslotsClick);
                                                timeTable.Children.Add(timeLabelButtonAdd);
                                            }

                                        }
                                    }
                                }
                                else if (i > row && j == column) //changes column to have classes start after previous class ends
                                {
                                    if (timeChange == 0)
                                    {
                                        if (parentGrid.Name == "MWF" && timeTableD == 0 || parentGrid.Name == "TR" && timeTableD == 1)
                                        {
                                            if (i == row + 1)
                                            {

                                                char removeChar = ':';
                                                meridianCheck = true;

                                                string[] newEndTime = endTime.Split(removeChar);
                                                int frontTime = Int32.Parse(newEndTime[0]);
                                                int backTime = Int32.Parse(newEndTime[1]);
                                                string meridianPure = meridian;
                                                int frontTimeZero = 0;
                                                string frontTimePadded = "";
                                                int backTimeZero = 0;
                                                string backTimePadded = "";
                                                string frontTimeApply = "";
                                                string backTimeApply = "";
                                                
                                                string[] checkMeridianChange = startTime.Split(removeChar);
                                                int startTimeFrontCheck = Int32.Parse(checkMeridianChange[0]);
                                                if (startTimeFrontCheck < 12 && frontTime < startTimeFrontCheck && frontTime != 12)
                                                {
                                                    if (meridian == "PM")
                                                    {
                                                        meridian = "AM";
                                                    }
                                                    if (meridian == "AM")
                                                    {
                                                        meridian = "PM";
                                                    }
                                                }

                                                if (frontTime == 12)
                                                {
                                                    meridianCheck = false;
                                                }

                                                backTime = backTime + increment;

                                                while (backTime >= 60)
                                                {
                                                    frontTime = frontTime + 1;
                                                    backTime = backTime - 60;
                                                }

                                                if (frontTime >= 12 && meridianCheck == false && meridian == "PM")
                                                {
                                                    frontTime = frontTime - 12;
                                                    meridian = "AM";
                                                }
                                                else if (frontTime >= 12 && meridianCheck == false && meridian == "AM")
                                                {
                                                    frontTime = frontTime - 12;
                                                    meridian = "PM";
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
                                                        endTimePure = frontTimeApply + ":" + backTimeApply;
                                                    }
                                                    else
                                                        backTimeApply = backTime.ToString();
                                                    endTimePure = frontTimeApply + ":" + backTimeApply;
                                                }
                                                else if (backTimePadded != "")
                                                {

                                                    backTimeApply = backTimePadded.ToString();

                                                    if (frontTimePadded != "")
                                                    {
                                                        frontTimeApply = frontTimePadded.ToString();
                                                        endTimePure = frontTimeApply + ":" + backTimeApply;
                                                    }
                                                    else
                                                    {
                                                        frontTimeApply = frontTime.ToString();
                                                        endTimePure = frontTimeApply + ":" + backTimeApply;
                                                    }
                                                }
                                                else
                                                {

                                                    frontTimeApply = frontTime.ToString();
                                                    backTimeApply = backTime.ToString();
                                                    endTimePure = frontTimeApply + ":" + backTimeApply;
                                                }

                                                autoChangedEndTime.Add(endTimePure);
                                                
                                                

                                                System.Windows.Controls.Button timeLabelButton = new System.Windows.Controls.Button();
                                                timeLabelButton.Content = endTime + " " + meridian;
                                                timeLabelButton.SetValue(Grid.RowProperty, i);
                                                timeLabelButton.SetValue(Grid.ColumnProperty, j);
                                                timeLabelButton.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                                timeLabelButton.VerticalAlignment = VerticalAlignment.Center;
                                                timeLabelButton.Click += new RoutedEventHandler(ChangeTime_Click);
                                                timeTable.Children.Add(timeLabelButton);
                                                autoChangedRoomNum.Add(classrooms[j / 2].RoomNum);
                                                autoChangedStartTime.Add(endTime);
                                                autoChangedMeridian.Add(meridian);
                                                autoChangedColumn.Add(j);
                                                autoChangedRow.Add(i);
                                                autoChangedTimeTable.Add(timeTableD);
                                                autoCount++;

                                                for (int q = 0; q < classList.Count(); q++) // Changes class to have stating and ending time equal to user made changes on button label
                                                {

                                                    bool breakLoop = false;
                                                    bool changeClass = true;

                                                    string time = "";
                                                    if (classList[q].StartTime.Start.Length > 4)
                                                    {
                                                        time = classList[q].StartTime.Start.Substring(0, 5);
                                                    }
                                                    else

                                                        for (int d = 0; d < count; d++) //Checks changed times for current class time and changes it
                                                        {
                                                            if (classList[q].ClassDay == "MWF" && timeTableD == 0 && changedTimeTable[d] == 0 || classList[q].ClassDay == "TR" && timeTableD == 1 && changedTimeTable[d] == 1)
                                                            {
                                                                if (classList[q].Classroom.RoomNum == classrooms[column / 2].RoomNum && changedRoomNum[d] == classList[q].Classroom.RoomNum)
                                                                {
                                                                    if (changedRow[d] == i)
                                                                    {
                                                                        if (changedStartTime[d] == time && changedMeridian[d] == classList[q].StartTime.Meridian)
                                                                        {
                                                                            if (time != endTime || time == endTime && meridian != classList[q].StartTime.Meridian)
                                                                            {
                                                                                for (int k = 0; k < changedClasses.Count(); k++)
                                                                                {
                                                                                    if (changedClasses[k] == classList[q].TextBoxName)
                                                                                    {
                                                                                        changeClass = false;
                                                                                    }
                                                                                }
                                                                                if (changeClass == true)
                                                                                {
                                                                                    classList[q].StartTime = new Timeslot(endTime, endTimePure, meridian);
                                                                                    changedClasses.Add(classList[q].TextBoxName);
                                                                                    breakLoop = true;
                                                                                    break;
                                                                                }
                                                                            }
                                                                        }
                                                                    }
                                                                }
                                                            }
                                                        }

                                                    for (int d = 0; d < autoCount; d++) // checks automatically changed times for current class time and changes it
                                                    {
                                                        if (classList[q].ClassDay == "MWF" && timeTableD == 0 && autoChangedTimeTable[d] == 0 || classList[q].ClassDay == "TR" && timeTableD == 1 && autoChangedTimeTable[d] == 1)
                                                        {
                                                            if (classList[q].Classroom.RoomNum == classrooms[column / 2].RoomNum && autoChangedRoomNum[d] == classList[q].Classroom.RoomNum)
                                                            {
                                                                if (autoChangedRow[d] == i)
                                                                {
                                                                    if (autoChangedStartTime[d] == time && autoChangedMeridian[d] == classList[q].StartTime.Meridian)
                                                                    {
                                                                        if (time != endTime || time == endTime && meridian != classList[q].StartTime.Meridian)
                                                                        {
                                                                            for (int k = 0; k < changedClasses.Count(); k++)
                                                                            {
                                                                                if (changedClasses[k] == classList[q].TextBoxName)
                                                                                {
                                                                                    changeClass = false;
                                                                                }
                                                                            }
                                                                            if (changeClass == true)
                                                                            {
                                                                                classList[q].StartTime = new Timeslot(endTime, endTimePure, meridian);
                                                                                changedClasses.Add(classList[q].TextBoxName);
                                                                                breakLoop = true;
                                                                                break;
                                                                            }
                                                                        }
                                                                    }
                                                                }
                                                            }
                                                        }
                                                    }
                                                    for (int d = 0; d < times_Default.Count(); d++) // checks default times for current class time and changes it
                                                    {
                                                        if (classList[q].ClassDay == "MWF" && timeTableD == 0 && times_Default_Timetable[d] == "MWF" || classList[q].ClassDay == "TR" && timeTableD == 1 && times_Default_Timetable[d] == "TR")
                                                        {
                                                            if (classList[q].Classroom.RoomNum == classrooms[column / 2].RoomNum && times_Default_Room[d] == classList[q].Classroom.RoomNum)
                                                            {
                                                                if (times_Default_Row[d] == i)
                                                                {
                                                                    if (times_Default[d].Time == classList[q].StartTime.Time && times_Default[d].Meridian == classList[q].StartTime.Meridian)
                                                                    {
                                                                        if (time != endTime || time == endTime && meridian != classList[q].StartTime.Meridian)
                                                                        {
                                                                            for (int k = 0; k < changedClasses.Count(); k++)
                                                                            {
                                                                                if (changedClasses[k] == classList[q].TextBoxName)
                                                                                {
                                                                                    changeClass = false;
                                                                                }
                                                                            }
                                                                            if (changeClass == true)
                                                                            {
                                                                                classList[q].StartTime = new Timeslot(endTime, endTimePure, meridian);
                                                                                changedClasses.Add(classList[q].TextBoxName);
                                                                                breakLoop = true;
                                                                                break;
                                                                            }
                                                                        }
                                                                    }
                                                                }
                                                            }
                                                        }
                                                    }

                                                    if (breakLoop == true)
                                                    {
                                                        break;
                                                    }

                                                }

                                                removeChar = ':';
                                                string[] timeLabel = endTime.Split(removeChar);
                                                string endTimeLabel = timeLabel[0];
                                                timeLabelButton.Name = timeTable.Name + endTimeLabel + "_" + classrooms[j / 2].ClassID + meridian;
                                                labelNames.Add(timeLabelButton.Name);
                                                labelColumn.Add(j + 1);
                                                labelRow.Add(i);
                                                labelTimeTable.Add(parentGrid.Name);
                                                labelCount++;
                                            }
                                            else
                                            {
                                                int frontTimeZero = 0;
                                                string frontTimePadded = "";
                                                int backTimeZero = 0;
                                                string backTimePadded = "";
                                                string frontTimeApply = "";
                                                string backTimeApply = "";
                                                


                                                char removeChar = ':';
                                                string[] time = endTime.Split(removeChar);
                                                int frontTime = int.Parse(time[0]);
                                                int backTime = int.Parse(time[1]);
                                                meridianCheck = false;

                                                if (frontTime == 12)
                                                {
                                                    meridianCheck = true;
                                                }
                                                int breakTime = 10;
                                                if (parentGrid.Name == "TR")
                                                {
                                                    breakTime = 15;
                                                }
                                                backTime = backTime + increment + breakTime; // sets minute side of time

                                                while (backTime >= 60)
                                                {
                                                    frontTime = frontTime + 1;
                                                    backTime = backTime - 60;
                                                }

                                                if (frontTime > 12 && meridianCheck == false && meridian == "PM")
                                                {
                                                    frontTime = frontTime - 12;
                                                    meridian = "AM";
                                                }
                                                else if (frontTime > 12 && meridianCheck == false && meridian == "AM")
                                                {
                                                    frontTime = frontTime - 12;
                                                    meridian = "PM";
                                                }
                                                else if (frontTime == 12 && meridianCheck == false && meridian == "AM")
                                                {
                                                    meridian = "PM";
                                                }
                                                else if (frontTime == 12 && meridianCheck == false && meridian == "PM")
                                                {
                                                    meridian = "AM";
                                                }
                                                
                                                if (frontTime > 12)
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
                                                        backTimeApply = backTime.ToString();
                                                    endTime = frontTimeApply + ":" + backTimeApply;
                                                }
                                                else if (backTimePadded != "")
                                                {

                                                    backTimeApply = backTimePadded.ToString();

                                                    if (frontTimePadded != "")
                                                    {
                                                        frontTimeApply = frontTimePadded.ToString();
                                                        endTime = frontTimeApply + ":" + backTimeApply;
                                                    }
                                                    else
                                                    {
                                                        frontTimeApply = frontTime.ToString();
                                                        endTime = frontTimeApply + ":" + backTimeApply;
                                                    }
                                                }
                                                else
                                                {

                                                    frontTimeApply = frontTime.ToString();
                                                    backTimeApply = backTime.ToString();
                                                    endTime = frontTimeApply + ":" + backTimeApply;
                                                }

                                                

                                                System.Windows.Controls.Button timeLabelButton = new System.Windows.Controls.Button();
                                                timeLabelButton.Content = endTime + " " + meridian;
                                                timeLabelButton.SetValue(Grid.RowProperty, i);
                                                timeLabelButton.SetValue(Grid.ColumnProperty, j);
                                                timeLabelButton.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                                timeLabelButton.VerticalAlignment = VerticalAlignment.Center;
                                                timeLabelButton.Click += new RoutedEventHandler(ChangeTime_Click);
                                                timeTable.Children.Add(timeLabelButton);
                                                autoChangedRoomNum.Add(classrooms[j / 2].RoomNum);
                                                autoChangedStartTime.Add(endTime);
                                                autoChangedMeridian.Add(meridian);
                                                autoChangedColumn.Add(j);
                                                autoChangedRow.Add(i);
                                                autoChangedTimeTable.Add(timeTableD);

                                                


                                                string[] newEndTime = endTime.Split(removeChar);
                                                int frontTimePure = Int32.Parse(newEndTime[0]);
                                                int backTimePure = Int32.Parse(newEndTime[1]);
                                                string meridianPure = meridian;
                                                int frontTimeZeroPure = 0;
                                                string frontTimePaddedPure = "";
                                                int backTimeZeroPure = 0;
                                                string backTimePaddedPure = "";
                                                string frontTimeApplyPure = "";
                                                string backTimeApplyPure = "";
                                                
                                                meridianCheck = false;
                                                if (frontTimePure == 12)
                                                {
                                                    meridianCheck = true;
                                                }

                                                backTimePure = backTimePure + increment;


                                                while (backTimePure >= 60)
                                                {
                                                    if (meridianPure == "PM" && meridianCheck == false)
                                                    {
                                                        frontTimePure = frontTimePure + 1;
                                                        backTimePure = backTimePure - 60;
                                                        meridianPure = "AM";
                                                    }
                                                    else if (meridianPure == "AM" && meridianCheck == false)
                                                    {
                                                        frontTimePure = frontTimePure + 1;
                                                        backTimePure = backTimePure - 60;
                                                        meridianPure = "PM";
                                                    }
                                                    else
                                                    {
                                                        frontTimePure = frontTimePure + 1;
                                                        backTimePure = backTimePure - 60;
                                                    }
                                                }

                                                if (frontTimePure > 12)
                                                {
                                                    frontTimePure = frontTimePure - 12;
                                                }

                                                

                                                if (backTimePure >= 60)
                                                {
                                                    backTimePure = backTimePure - 60;
                                                    frontTimePure = frontTimePure + 1;
                                                }
                                                if (frontTimePure >= 12 && meridianCheck == false)
                                                {
                                                    if (meridianPure == "PM")
                                                    {
                                                        meridianPure = "AM";
                                                    }
                                                    else if (meridianPure == "AM")
                                                    {
                                                        meridianPure = "PM";
                                                    }
                                                }
                                                if (frontTimePure > 12)
                                                {
                                                    frontTimePure = frontTimePure - 12;
                                                }
                                                if (frontTimePure < 10 && frontTimePure.ToString().Length == 1)
                                                {
                                                    frontTimeZeroPure = frontTimePure.ToString("D").Length + 1;
                                                    frontTimePaddedPure = frontTimePure.ToString("D" + frontTimeZeroPure.ToString());
                                                }
                                                if (backTimePure < 10 && backTimePure.ToString().Length == 1)
                                                {
                                                    backTimeZeroPure = backTimePure.ToString("D").Length + 1;
                                                    backTimePaddedPure = backTimePure.ToString("D" + backTimeZeroPure.ToString());
                                                }

                                                if (frontTimePaddedPure != "")
                                                {
                                                    frontTimeApplyPure = frontTimePaddedPure.ToString();

                                                    if (backTimePaddedPure != "")
                                                    {
                                                        backTimeApplyPure = backTimePaddedPure.ToString();
                                                        endTimePure = frontTimeApplyPure + ":" + backTimeApplyPure;
                                                    }
                                                    else
                                                        backTimeApplyPure = backTimePure.ToString();
                                                    endTimePure = frontTimeApplyPure + ":" + backTimeApplyPure;
                                                }
                                                else if (backTimePaddedPure != "")
                                                {

                                                    backTimeApplyPure = backTimePaddedPure.ToString();

                                                    if (frontTimePaddedPure != "")
                                                    {
                                                        frontTimeApplyPure = frontTimePaddedPure.ToString();
                                                        endTimePure = frontTimeApplyPure + ":" + backTimeApplyPure;
                                                    }
                                                    else
                                                    {
                                                        frontTimeApplyPure = frontTimePure.ToString();
                                                        endTimePure = frontTimeApplyPure + ":" + backTimeApplyPure;
                                                    }
                                                }
                                                else
                                                {

                                                    frontTimeApplyPure = frontTimePure.ToString();
                                                    backTimeApplyPure = backTimePure.ToString();
                                                    endTimePure = frontTimeApplyPure + ":" + backTimeApplyPure;
                                                }

                                                autoChangedEndTime.Add(endTimePure);

                                                autoCount++;

                                                for (int q = 0; q < classList.Count(); q++) // Changes class to have stating and ending time equal to user made changes on button label
                                                {
                                                    bool breakLoop = false;
                                                    bool changeClass = true;
                                                    string timeSet = "";
                                                    if (classList[q].StartTime.Start.Length > 4)
                                                    {
                                                        timeSet = classList[q].StartTime.Start.Substring(0, 5);
                                                    }
                                                    
                                                    
                                                        for (int d = 0; d < count; d++) //Checks changed times for current class time and changes it
                                                        {
                                                            if (classList[q].ClassDay == "MWF" && timeTableD == 0 && changedTimeTable[d] == 0 || classList[q].ClassDay == "TR" && timeTableD == 1 && changedTimeTable[d] == 1)
                                                            {
                                                                if (classList[q].Classroom.RoomNum == classrooms[column / 2].RoomNum && changedRoomNum[d] == classList[q].Classroom.RoomNum)
                                                                {
                                                                    if (changedRow[d] == i)
                                                                    {
                                                                        if (changedStartTime[d] == timeSet && changedMeridian[d] == classList[q].StartTime.Meridian)
                                                                        {
                                                                            if (timeSet != endTime || timeSet == endTime && meridian != classList[q].StartTime.Meridian)
                                                                            {
                                                                                for (int k = 0; k < changedClasses.Count(); k++)
                                                                                {
                                                                                    if (changedClasses[k] == classList[q].TextBoxName)
                                                                                    {
                                                                                        changeClass = false;
                                                                                    }
                                                                                }
                                                                                if (changeClass == true)
                                                                                {
                                                                                    classList[q].StartTime = new Timeslot(endTime, endTimePure, meridian);
                                                                                    changedClasses.Add(classList[q].TextBoxName);
                                                                                    breakLoop = true;
                                                                                    break;
                                                                                }
                                                                            }
                                                                        }
                                                                    }
                                                                }
                                                            }
                                                        }

                                                    for (int d = 0; d < autoCount; d++) // checks automatically changed times for current class time and changes it
                                                    {
                                                        if (classList[q].ClassDay == "MWF" && timeTableD == 0 && autoChangedTimeTable[d] == 0 || classList[q].ClassDay == "TR" && timeTableD == 1 && autoChangedTimeTable[d] == 1)
                                                        {
                                                            if (classList[q].Classroom.RoomNum == classrooms[column / 2].RoomNum && autoChangedRoomNum[d] == classList[q].Classroom.RoomNum)
                                                            {
                                                                if (autoChangedRow[d] == i)
                                                                {
                                                                    if (autoChangedStartTime[d] == timeSet && autoChangedMeridian[d] == classList[q].StartTime.Meridian)
                                                                    {
                                                                        if (timeSet != endTime || timeSet == endTime && meridian != classList[q].StartTime.Meridian)
                                                                        {
                                                                            for (int k = 0; k < changedClasses.Count(); k++)
                                                                            {
                                                                                if (changedClasses[k] == classList[q].TextBoxName)
                                                                                {
                                                                                    changeClass = false;
                                                                                }
                                                                            }
                                                                            if (changeClass == true)
                                                                            {
                                                                                classList[q].StartTime = new Timeslot(endTime, endTimePure, meridian);
                                                                                changedClasses.Add(classList[q].TextBoxName);
                                                                                breakLoop = true;
                                                                                break;
                                                                            }
                                                                        }
                                                                    }
                                                                }
                                                            }
                                                        }
                                                    }
                                                    for (int d = 0; d < times_Default.Count(); d++) // checks default times for current class time and changes it
                                                    {
                                                        if (classList[q].ClassDay == "MWF" && timeTableD == 0 && times_Default_Timetable[d] == "MWF" || classList[q].ClassDay == "TR" && timeTableD == 1 && times_Default_Timetable[d] == "TR")
                                                        {
                                                            if (classList[q].Classroom.RoomNum == classrooms[column / 2].RoomNum && times_Default_Room[d] == classList[q].Classroom.RoomNum)
                                                            {
                                                                if (times_Default_Row[d] == i)
                                                                {
                                                                   
                                                                    if (times_Default[d].Start == classList[q].StartTime.Start && times_Default[d].Meridian == classList[q].StartTime.Meridian)
                                                                    {
                                                                        if (timeSet != endTime || timeSet == endTime && meridian != classList[q].StartTime.Meridian)
                                                                        {
                                                                            for (int k = 0; k < changedClasses.Count(); k++)
                                                                            {
                                                                                if (changedClasses[k] == classList[q].TextBoxName)
                                                                                {
                                                                                    changeClass = false;
                                                                                }
                                                                            }
                                                                            if (changeClass == true)
                                                                            {
                                                                                classList[q].StartTime = new Timeslot(endTime, endTimePure, meridian);
                                                                                changedClasses.Add(classList[q].TextBoxName);
                                                                                breakLoop = true;
                                                                                break;
                                                                            }
                                                                        }
                                                                    }
                                                                }
                                                            }
                                                        }
                                                    }

                                                    if (breakLoop == true)
                                                    {
                                                        break;
                                                    }
                                                }

                                                removeChar = ':';
                                                string[] timeLabel = endTime.Split(removeChar);
                                                string endTimeLabel = timeLabel[0];
                                                timeLabelButton.Name = timeTable.Name + endTimeLabel + "_" + classrooms[j / 2].ClassID + meridian;
                                                labelNames.Add(timeLabelButton.Name);
                                                labelColumn.Add(j + 1);
                                                labelRow.Add(i);
                                                labelTimeTable.Add(parentGrid.Name);
                                                labelCount++;
                                            }
                                        }
                                        else
                                        {
                                            if (i <= rowLength && parentGrid.Name == "MWF")
                                            {
                                                if (i <= times.Length)
                                                {
                                                    System.Windows.Controls.Button timeLabelButton = new System.Windows.Controls.Button();
                                                    timeLabelButton.Content = times[i - 1].Time + " " + times[i - 1].Meridian;
                                                    timeLabelButton.Name = timeTable.Name + times[i - 1].TimeID + "_" + classrooms[j / 2].ClassID + times[i - 1].Meridian;
                                                    timeLabelButton.SetValue(Grid.RowProperty, i);
                                                    timeLabelButton.SetValue(Grid.ColumnProperty, j);
                                                    timeLabelButton.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                                    timeLabelButton.VerticalAlignment = VerticalAlignment.Center;
                                                    timeLabelButton.Click += new RoutedEventHandler(ChangeTime_Click);
                                                    timeTable.Children.Add(timeLabelButton);


                                                    bool labelCheck = false;
                                                    for (int l = 0; l < labelNames.Count(); l++)
                                                    {
                                                        if (labelTimeTable[l] == parentGrid.Name)
                                                        {
                                                            if (labelColumn[l] == j + 1)
                                                            {
                                                                if (labelRow[l] == i)
                                                                {
                                                                    labelCheck = true;
                                                                }
                                                            }
                                                        }
                                                    }

                                                    if (labelCheck == false)
                                                    {
                                                        labelNames.Add(timeLabelButton.Name);
                                                        labelColumn.Add(j + 1);
                                                        labelRow.Add(i);
                                                        labelTimeTable.Add(parentGrid.Name);
                                                        labelCount++;
                                                    }

                                                    System.Windows.Controls.Button timeLabelButtonAdd = new System.Windows.Controls.Button();
                                                    timeLabelButtonAdd.Content = "Add Row";
                                                    timeLabelButtonAdd.SetValue(Grid.RowProperty, i + 1);
                                                    timeLabelButtonAdd.SetValue(Grid.ColumnProperty, j);
                                                    timeLabelButtonAdd.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                                    timeLabelButtonAdd.VerticalAlignment = VerticalAlignment.Center;
                                                    timeLabelButtonAdd.Background = new SolidColorBrush(Colors.Thistle);
                                                    timeLabelButtonAdd.Click += new RoutedEventHandler(AddTimeslotsClick);
                                                    timeTable.Children.Add(timeLabelButtonAdd);
                                                }

                                            }
                                            else if (i <= rowLength && parentGrid.Name == "TR")
                                            {
                                                if (i <= times.Length)
                                                {
                                                    System.Windows.Controls.Button timeLabelButton = new System.Windows.Controls.Button();
                                                    timeLabelButton.Content = times[i - 1].Time + " " + times[i - 1].Meridian;
                                                    timeLabelButton.Name = timeTable.Name + times[i - 1].TimeID + "_" + classrooms[j / 2].ClassID + times[i - 1].Meridian;
                                                    timeLabelButton.SetValue(Grid.RowProperty, i);
                                                    timeLabelButton.SetValue(Grid.ColumnProperty, j);
                                                    timeLabelButton.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                                    timeLabelButton.VerticalAlignment = VerticalAlignment.Center;
                                                    timeLabelButton.Click += new RoutedEventHandler(ChangeTime_Click);
                                                    timeTable.Children.Add(timeLabelButton);


                                                    bool labelCheck = false;
                                                    for (int l = 0; l < labelNames.Count(); l++)
                                                    {
                                                        if (labelTimeTable[l] == parentGrid.Name)
                                                        {
                                                            if (labelColumn[l] == j + 1)
                                                            {
                                                                if (labelRow[l] == i)
                                                                {
                                                                    labelCheck = true;
                                                                }
                                                            }
                                                        }
                                                    }

                                                    if (labelCheck == false)
                                                    {
                                                        labelNames.Add(timeLabelButton.Name);
                                                        labelColumn.Add(j + 1);
                                                        labelRow.Add(i);
                                                        labelTimeTable.Add(parentGrid.Name);
                                                        labelCount++;
                                                    }

                                                    System.Windows.Controls.Button timeLabelButtonAdd = new System.Windows.Controls.Button();
                                                    timeLabelButtonAdd.Content = "Add Row";
                                                    timeLabelButtonAdd.SetValue(Grid.RowProperty, i + 1);
                                                    timeLabelButtonAdd.SetValue(Grid.ColumnProperty, j);
                                                    timeLabelButtonAdd.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                                    timeLabelButtonAdd.VerticalAlignment = VerticalAlignment.Center;
                                                    timeLabelButtonAdd.Background = new SolidColorBrush(Colors.Thistle);
                                                    timeLabelButtonAdd.Click += new RoutedEventHandler(AddTimeslotsClick);
                                                    timeTable.Children.Add(timeLabelButtonAdd);
                                                }
                                            }
                                        }
                                    }
                                    else
                                    {
                                        if (i <= rowLength && parentGrid.Name == "MWF")
                                        {
                                            if (i <= times.Length)
                                            {
                                                System.Windows.Controls.Button timeLabelButton = new System.Windows.Controls.Button();
                                                timeLabelButton.Content = times[i - 1].Time + " " + times[i - 1].Meridian;
                                                timeLabelButton.Name = timeTable.Name + times[i - 1].TimeID + "_" + classrooms[j / 2].ClassID + times[i - 1].Meridian;
                                                timeLabelButton.SetValue(Grid.RowProperty, i);
                                                timeLabelButton.SetValue(Grid.ColumnProperty, j);
                                                timeLabelButton.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                                timeLabelButton.VerticalAlignment = VerticalAlignment.Center;
                                                timeLabelButton.Click += new RoutedEventHandler(ChangeTime_Click);
                                                timeTable.Children.Add(timeLabelButton);
                                                labelNames.Add(timeLabelButton.Name);
                                                labelColumn.Add(j + 1);
                                                labelRow.Add(i);
                                                labelTimeTable.Add(parentGrid.Name);
                                                labelCount++;

                                                System.Windows.Controls.Button timeLabelButtonAdd = new System.Windows.Controls.Button();
                                                timeLabelButtonAdd.Content = "Add Row";
                                                timeLabelButtonAdd.SetValue(Grid.RowProperty, i + 1);
                                                timeLabelButtonAdd.SetValue(Grid.ColumnProperty, j);
                                                timeLabelButtonAdd.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                                timeLabelButtonAdd.VerticalAlignment = VerticalAlignment.Center;
                                                timeLabelButtonAdd.Background = new SolidColorBrush(Colors.Thistle);
                                                timeLabelButtonAdd.Click += new RoutedEventHandler(AddTimeslotsClick);
                                                timeTable.Children.Add(timeLabelButtonAdd);
                                            }
                                        }
                                        else if (i <= rowLength && parentGrid.Name == "TR")
                                        {
                                            if (i <= times.Length)
                                            {
                                                System.Windows.Controls.Button timeLabelButton = new System.Windows.Controls.Button();
                                                timeLabelButton.Content = times[i - 1].Time + " " + times[i - 1].Meridian;
                                                timeLabelButton.Name = timeTable.Name + times[i - 1].TimeID + "_" + classrooms[j / 2].ClassID + times[i - 1].Meridian;
                                                timeLabelButton.SetValue(Grid.RowProperty, i);
                                                timeLabelButton.SetValue(Grid.ColumnProperty, j);
                                                timeLabelButton.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                                timeLabelButton.VerticalAlignment = VerticalAlignment.Center;
                                                timeLabelButton.Click += new RoutedEventHandler(ChangeTime_Click);
                                                timeTable.Children.Add(timeLabelButton);
                                                labelNames.Add(timeLabelButton.Name);
                                                labelColumn.Add(j + 1);
                                                labelRow.Add(i);
                                                labelTimeTable.Add(parentGrid.Name);
                                                labelCount++;

                                                System.Windows.Controls.Button timeLabelButtonAdd = new System.Windows.Controls.Button();
                                                timeLabelButtonAdd.Content = "Add Row";
                                                timeLabelButtonAdd.SetValue(Grid.RowProperty, i + 1);
                                                timeLabelButtonAdd.SetValue(Grid.ColumnProperty, j);
                                                timeLabelButtonAdd.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                                timeLabelButtonAdd.VerticalAlignment = VerticalAlignment.Center;
                                                timeLabelButtonAdd.Background = new SolidColorBrush(Colors.Thistle);
                                                timeLabelButtonAdd.Click += new RoutedEventHandler(AddTimeslotsClick);
                                                timeTable.Children.Add(timeLabelButtonAdd);
                                            }
                                        }
                                    }
                                }
                                else
                                {
                                    if (parentGrid.Name == "MWF" && timeTableD == 0 || parentGrid.Name == "TR" && timeTableD == 1 || parentGrid.Name != "MWF" && timeTableD == 0 || parentGrid.Name != "TR" && timeTableD == 1)
                                    {
                                        if (i <= rowLength && parentGrid.Name == "MWF")
                                        {
                                            if (i <= times.Length)
                                            {
                                                System.Windows.Controls.Button timeLabelButton = new System.Windows.Controls.Button();
                                                timeLabelButton.Content = times[i - 1].Time + " " + times[i - 1].Meridian;
                                                timeLabelButton.Name = timeTable.Name + times[i - 1].TimeID + "_" + classrooms[j / 2].ClassID + times[i - 1].Meridian;
                                                timeLabelButton.SetValue(Grid.RowProperty, i);
                                                timeLabelButton.SetValue(Grid.ColumnProperty, j);
                                                timeLabelButton.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                                timeLabelButton.VerticalAlignment = VerticalAlignment.Center;
                                                timeLabelButton.Click += new RoutedEventHandler(ChangeTime_Click);
                                                timeTable.Children.Add(timeLabelButton);
                                                labelNames.Add(timeLabelButton.Name);
                                                labelColumn.Add(j + 1);
                                                labelRow.Add(i);
                                                labelTimeTable.Add(parentGrid.Name);
                                                labelCount++;
                                            }
                                        }
                                        else if (i <= rowLength && parentGrid.Name == "TR")
                                        {
                                            if (i <= times.Length)
                                            {
                                                System.Windows.Controls.Button timeLabelButton = new System.Windows.Controls.Button();
                                                timeLabelButton.Content = times[i - 1].Time + " " + times[i - 1].Meridian;
                                                timeLabelButton.Name = timeTable.Name + times[i - 1].TimeID + "_" + classrooms[j / 2].ClassID + times[i - 1].Meridian;
                                                timeLabelButton.SetValue(Grid.RowProperty, i);
                                                timeLabelButton.SetValue(Grid.ColumnProperty, j);
                                                timeLabelButton.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                                timeLabelButton.VerticalAlignment = VerticalAlignment.Center;
                                                timeLabelButton.Click += new RoutedEventHandler(ChangeTime_Click);
                                                timeTable.Children.Add(timeLabelButton);
                                                labelNames.Add(timeLabelButton.Name);
                                                labelColumn.Add(j + 1);
                                                labelRow.Add(i);
                                                labelTimeTable.Add(parentGrid.Name);
                                                labelCount++;
                                            }
                                        }
                                    }
                                }
                                for (c = 0; c < autoCount; c++) // apply automatic saved column changes to grid
                                {
                                    for (int k = 0; k < classrooms.Count; k++)
                                    {
                                        if (autoChangedRoomNum[c] == classrooms[k].RoomNum)
                                        {
                                            autoChangedColumn[c] = k + k;
                                        }
                                    }
                                    if (j == autoChangedColumn[c]  || column != autoChangedColumn[c] && j == autoChangedColumn[c])
                                    {
                                        if (parentGrid.Name == "MWF" && autoChangedTimeTable[c] == 0 || parentGrid.Name == "TR" && autoChangedTimeTable[c] == 1)
                                        {
                                            for (int r = 0; r <= rowLength; r++)
                                            {
                                                if (r == autoChangedRow[c])
                                                {
                                                    if (checkAddRowButton < autoChangedRow[c])
                                                    {
                                                        checkAddRowButton = autoChangedRow[c];
                                                    }
                                                    System.Windows.Controls.Button timeLabelButton = new System.Windows.Controls.Button();
                                                    timeLabelButton.Content = autoChangedStartTime[c] + " " + autoChangedMeridian[c];
                                                    timeLabelButton.SetValue(Grid.RowProperty, autoChangedRow[c]);
                                                    timeLabelButton.SetValue(Grid.ColumnProperty, autoChangedColumn[c]);
                                                    timeLabelButton.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                                    timeLabelButton.VerticalAlignment = VerticalAlignment.Center;
                                                    timeLabelButton.Click += new RoutedEventHandler(ChangeTime_Click);
                                                    timeTable.Children.Add(timeLabelButton);

                                                    System.Windows.Controls.Button timeLabelButtonAdd = new System.Windows.Controls.Button();
                                                    timeLabelButtonAdd.Content = "Add Row";
                                                    timeLabelButtonAdd.SetValue(Grid.RowProperty, autoChangedRow[c] + 1);
                                                    timeLabelButtonAdd.SetValue(Grid.ColumnProperty, autoChangedColumn[c]);
                                                    timeLabelButtonAdd.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                                    timeLabelButtonAdd.VerticalAlignment = VerticalAlignment.Center;
                                                    timeLabelButtonAdd.Background = new SolidColorBrush(Colors.Thistle);
                                                    timeLabelButtonAdd.Click += new RoutedEventHandler(AddTimeslotsClick);
                                                    timeTable.Children.Add(timeLabelButtonAdd);

                                                    /*
                                                    char removeChar = ':';
                                                    string[] time = autoChangedStartTime[c].Split(removeChar);
                                                    string startTimeLabel = time[0];
                                                    timeLabelButton.Name = timeTable.Name + startTimeLabel + "_" + classrooms[j / 2].ClassID;
                                                    labelNames[labelCount] = timeLabelButton.Name;
                                                    labelColumn[labelCount] = autoChangedColumn[c] + 1;
                                                    labelRow[labelCount] = r;
                                                    labelTimeTable[labelCount] = parentGrid.Name;
                                                    labelCount++;
                                                    */
                                                }

                                            }
                                        }
                                    }
                                    else
                                    {
                                        /*
                                        if (parentGrid.Name == "MWF" && autoChangedTimeTable[c] == 0 || parentGrid.Name == "TR" && autoChangedTimeTable[c] == 1)
                                        {
                                            if (i <= defaultRowCountMWF && parentGrid.Name == "MWF")
                                            {
                                                for (int t = 0; t < times_Default.Count; t++)
                                                {
                                                    if (i == times_Default_Row[t] && times_Default_Room[t] == classrooms[j / 2].RoomNum && times_Default_Timetable[t] == parentGrid.Name)
                                                    {
                                                        System.Windows.Controls.Button timeLabelButton = new System.Windows.Controls.Button();
                                                        timeLabelButton.Content = times_Default[t].Time + " " + times_Default[t].Meridian;
                                                        timeLabelButton.Name = timeTable.Name + times_Default[t].TimeID + "_" + classrooms[j / 2].ClassID;
                                                        timeLabelButton.SetValue(Grid.RowProperty, i);
                                                        timeLabelButton.SetValue(Grid.ColumnProperty, j);
                                                        timeLabelButton.HorizontalAlignment = HorizontalAlignment.Left;
                                                        timeLabelButton.VerticalAlignment = VerticalAlignment.Center;
                                                        timeLabelButton.Click += new RoutedEventHandler(ChangeTime_Click);
                                                        timeTable.Children.Add(timeLabelButton);
                                                        labelNames[labelCount] = timeLabelButton.Name;
                                                        labelColumn[labelCount] = j + 1;
                                                        labelRow[labelCount] = i;
                                                        labelTimeTable[labelCount] = parentGrid.Name;
                                                        labelCount++;
                                                    }
                                                }
                                            }
                                            else if (i <= defaultRowCountTR && parentGrid.Name == "TR")
                                            {
                                                for (int t = 0; t < times_Default.Count; t++)
                                                {
                                                    if (i == times_Default_Row[t] && times_Default_Room[t] == classrooms[j / 2].RoomNum && times_Default_Timetable[t] == parentGrid.Name)
                                                    {
                                                        System.Windows.Controls.Button timeLabelButton = new System.Windows.Controls.Button();
                                                        timeLabelButton.Content = times_Default[t].Time + " " + times_Default[t].Meridian;
                                                        timeLabelButton.Name = timeTable.Name + times_Default[t].TimeID + "_" + classrooms[j / 2].ClassID;
                                                        timeLabelButton.SetValue(Grid.RowProperty, i);
                                                        timeLabelButton.SetValue(Grid.ColumnProperty, j);
                                                        timeLabelButton.HorizontalAlignment = HorizontalAlignment.Left;
                                                        timeLabelButton.VerticalAlignment = VerticalAlignment.Center;
                                                        timeLabelButton.Click += new RoutedEventHandler(ChangeTime_Click);
                                                        timeTable.Children.Add(timeLabelButton);
                                                        labelNames[labelCount] = timeLabelButton.Name;
                                                        labelColumn[labelCount] = j + 1;
                                                        labelRow[labelCount] = i;
                                                        labelTimeTable[labelCount] = parentGrid.Name;
                                                        labelCount++;
                                                    }
                                                }
                                            }
                                        
                                        }
                                        */
                                    }
                                }

                                for (c = 0; c < count; c++)  // applys saved grid values to grid
                                {
                                    for (int k = 0; k < classrooms.Count; k++)
                                    {
                                        if (changedRoomNum[c] == classrooms[k].RoomNum)
                                        {
                                            changedColumn[c] = k + k;
                                        }
                                    }

                                    if (j == changedColumn[c])
                                    {
                                        if (parentGrid.Name == "MWF" && changedTimeTable[c] == 0 || parentGrid.Name == "TR" && changedTimeTable[c] == 1)
                                        {
                                            int lastChangedRow = 0;
                                            for (int r = 0; r <= rowLength; r++)
                                            {
                                                if (r == changedRow[c])
                                                {
                                                    System.Windows.Controls.Button timeLabelButton = new System.Windows.Controls.Button();
                                                    timeLabelButton.Content = changedStartTime[c] + " " + changedMeridian[c];
                                                    timeLabelButton.SetValue(Grid.RowProperty, changedRow[c]);
                                                    timeLabelButton.SetValue(Grid.ColumnProperty, changedColumn[c]);
                                                    timeLabelButton.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                                    timeLabelButton.VerticalAlignment = VerticalAlignment.Center;
                                                    timeLabelButton.Click += new RoutedEventHandler(ChangeTime_Click);
                                                    timeTable.Children.Add(timeLabelButton);

                                                    
                                                    if (changedRow[c] > lastChangedRow)
                                                    {
                                                        lastChangedRow = changedRow[c];
                                                    }
                                                    if (checkAddRowButton <= lastChangedRow)
                                                    {
                                                        if (changedRow[c] == lastChangedRow && timeChange == 0 && lastChangedRow == rowLength || changedRow[c] == lastChangedRow && lastChangedRow == rowLength)
                                                        {
                                                            System.Windows.Controls.Button timeLabelButtonAdd = new System.Windows.Controls.Button();
                                                            timeLabelButtonAdd.Content = "Add Row";
                                                            timeLabelButtonAdd.SetValue(Grid.RowProperty, changedRow[c] + 1);
                                                            timeLabelButtonAdd.SetValue(Grid.ColumnProperty, changedColumn[c]);
                                                            timeLabelButtonAdd.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                                            timeLabelButtonAdd.VerticalAlignment = VerticalAlignment.Center;
                                                            timeLabelButtonAdd.Background = new SolidColorBrush(Colors.Thistle);
                                                            timeLabelButtonAdd.Click += new RoutedEventHandler(AddTimeslotsClick);
                                                            timeTable.Children.Add(timeLabelButtonAdd);
                                                        }
                                                    }
                                                    
                                                    /*
                                                    char removeChar = ':';
                                                    string[] time = changedStartTime[c].Split(removeChar);
                                                    string startTimeLabel = time[0];
                                                    timeLabelButton.Name = timeTable.Name + startTimeLabel + "_" + classrooms[j / 2].ClassID;
                                                    labelNames[labelCount] = timeLabelButton.Name;
                                                    labelColumn[labelCount] = changedColumn[c] + 1;
                                                    labelRow[labelCount] = r;
                                                    labelTimeTable[labelCount] = parentGrid.Name;
                                                    labelCount++;
                                                    */

                                                }
                                            }
                                        }
                                    }
                                }
                            }


                            else
                            {

                                if (i != row || j != column) // first change is being made loads default values
                                {
                                    if (parentGrid.Name == "MWF" && timeTableD == 0 || parentGrid.Name == "TR" && timeTableD == 1)
                                    {
                                        if (i <= rowLength && parentGrid.Name == "MWF")
                                        {
                                            if (i <= times.Length)
                                            {
                                                System.Windows.Controls.Button timeLabelButton = new System.Windows.Controls.Button();
                                                timeLabelButton.Content = times[i - 1].Time + " " + times[i - 1].Meridian;
                                                timeLabelButton.Name = timeTable.Name + times[i - 1].TimeID + "_" + classrooms[j / 2].ClassID + times[i - 1].Meridian;
                                                timeLabelButton.SetValue(Grid.RowProperty, i);
                                                timeLabelButton.SetValue(Grid.ColumnProperty, j);
                                                timeLabelButton.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                                timeLabelButton.VerticalAlignment = VerticalAlignment.Center;
                                                timeLabelButton.Click += new RoutedEventHandler(ChangeTime_Click);
                                                timeTable.Children.Add(timeLabelButton);
                                                labelNames.Add(timeLabelButton.Name);
                                                labelColumn.Add(j + 1);
                                                labelRow.Add(i);
                                                labelTimeTable.Add(parentGrid.Name);
                                                labelCount++;

                                                System.Windows.Controls.Button timeLabelButtonAdd = new System.Windows.Controls.Button();
                                                timeLabelButtonAdd.Content = "Add Row";
                                                timeLabelButtonAdd.SetValue(Grid.RowProperty, i + 1);
                                                timeLabelButtonAdd.SetValue(Grid.ColumnProperty, j);
                                                timeLabelButtonAdd.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                                timeLabelButtonAdd.VerticalAlignment = VerticalAlignment.Center;
                                                timeLabelButtonAdd.Background = new SolidColorBrush(Colors.Thistle);
                                                timeLabelButtonAdd.Click += new RoutedEventHandler(AddTimeslotsClick);
                                                timeTable.Children.Add(timeLabelButtonAdd);
                                            }

                                        }
                                        else if (i <= rowLength && parentGrid.Name == "TR")
                                        {
                                            if (i <= times.Length)
                                            {


                                                System.Windows.Controls.Button timeLabelButton = new System.Windows.Controls.Button();
                                                timeLabelButton.Content = times[i - 1].Time + " " + times[i - 1].Meridian;
                                                timeLabelButton.Name = timeTable.Name + times[i - 1].TimeID + "_" + classrooms[j / 2].ClassID + times[i - 1].Meridian;
                                                timeLabelButton.SetValue(Grid.RowProperty, i);
                                                timeLabelButton.SetValue(Grid.ColumnProperty, j);
                                                timeLabelButton.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                                timeLabelButton.VerticalAlignment = VerticalAlignment.Center;
                                                timeLabelButton.Click += new RoutedEventHandler(ChangeTime_Click);
                                                timeTable.Children.Add(timeLabelButton);
                                                labelNames.Add(timeLabelButton.Name);
                                                labelColumn.Add(j + 1);
                                                labelRow.Add(i);
                                                labelTimeTable.Add(parentGrid.Name);
                                                labelCount++;

                                                System.Windows.Controls.Button timeLabelButtonAdd = new System.Windows.Controls.Button();
                                                timeLabelButtonAdd.Content = "Add Row";
                                                timeLabelButtonAdd.SetValue(Grid.RowProperty, i + 1);
                                                timeLabelButtonAdd.SetValue(Grid.ColumnProperty, j);
                                                timeLabelButtonAdd.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                                timeLabelButtonAdd.VerticalAlignment = VerticalAlignment.Center;
                                                timeLabelButtonAdd.Background = new SolidColorBrush(Colors.Thistle);
                                                timeLabelButtonAdd.Click += new RoutedEventHandler(AddTimeslotsClick);
                                                timeTable.Children.Add(timeLabelButtonAdd);
                                            }

                                        }
                                    }
                                    else
                                    {
                                        if (i <= rowLength && parentGrid.Name == "MWF")
                                        {
                                            if (i <= times.Length)
                                            {
                                                System.Windows.Controls.Button timeLabelButton = new System.Windows.Controls.Button();
                                                timeLabelButton.Content = times[i - 1].Time + " " + times[i - 1].Meridian;
                                                timeLabelButton.Name = timeTable.Name + times[i - 1].TimeID + "_" + classrooms[j / 2].ClassID + times[i - 1].Meridian;
                                                timeLabelButton.SetValue(Grid.RowProperty, i);
                                                timeLabelButton.SetValue(Grid.ColumnProperty, j);
                                                timeLabelButton.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                                timeLabelButton.VerticalAlignment = VerticalAlignment.Center;
                                                timeLabelButton.Click += new RoutedEventHandler(ChangeTime_Click);
                                                timeTable.Children.Add(timeLabelButton);
                                                labelNames.Add(timeLabelButton.Name);
                                                labelColumn.Add(j + 1);
                                                labelRow.Add(i);
                                                labelTimeTable.Add(parentGrid.Name);
                                                labelCount++;

                                                System.Windows.Controls.Button timeLabelButtonAdd = new System.Windows.Controls.Button();
                                                timeLabelButtonAdd.Content = "Add Row";
                                                timeLabelButtonAdd.SetValue(Grid.RowProperty, i + 1);
                                                timeLabelButtonAdd.SetValue(Grid.ColumnProperty, j);
                                                timeLabelButtonAdd.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                                timeLabelButtonAdd.VerticalAlignment = VerticalAlignment.Center;
                                                timeLabelButtonAdd.Background = new SolidColorBrush(Colors.Thistle);
                                                timeLabelButtonAdd.Click += new RoutedEventHandler(AddTimeslotsClick);
                                                timeTable.Children.Add(timeLabelButtonAdd);
                                            }

                                        }
                                        else if (i <= rowLength && parentGrid.Name == "TR")
                                        {
                                            if (i <= times.Length)
                                            {
                                                System.Windows.Controls.Button timeLabelButton = new System.Windows.Controls.Button();
                                                timeLabelButton.Content = times[i - 1].Time + " " + times[i - 1].Meridian;
                                                timeLabelButton.Name = timeTable.Name + times[i - 1].TimeID + "_" + classrooms[j / 2].ClassID + times[i - 1].Meridian;
                                                timeLabelButton.SetValue(Grid.RowProperty, i);
                                                timeLabelButton.SetValue(Grid.ColumnProperty, j);
                                                timeLabelButton.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                                timeLabelButton.VerticalAlignment = VerticalAlignment.Center;
                                                timeLabelButton.Click += new RoutedEventHandler(ChangeTime_Click);
                                                timeTable.Children.Add(timeLabelButton);
                                                labelNames.Add(timeLabelButton.Name);
                                                labelColumn.Add(j + 1);
                                                labelRow.Add(i);
                                                labelTimeTable.Add(parentGrid.Name);
                                                labelCount++;

                                                System.Windows.Controls.Button timeLabelButtonAdd = new System.Windows.Controls.Button();
                                                timeLabelButtonAdd.Content = "Add Row";
                                                timeLabelButtonAdd.SetValue(Grid.RowProperty, i + 1);
                                                timeLabelButtonAdd.SetValue(Grid.ColumnProperty, j);
                                                timeLabelButtonAdd.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                                timeLabelButtonAdd.VerticalAlignment = VerticalAlignment.Center;
                                                timeLabelButtonAdd.Background = new SolidColorBrush(Colors.Thistle);
                                                timeLabelButtonAdd.Click += new RoutedEventHandler(AddTimeslotsClick);
                                                timeTable.Children.Add(timeLabelButtonAdd);
                                            }

                                        }
                                    }
                                }
                                else //first change being made to a time slot making a new timeslot
                                {
                                    if (parentGrid.Name == "MWF" && timeTableD == 0 || parentGrid.Name == "TR" && timeTableD == 1)
                                    {
                                        System.Windows.Controls.Button timeLabelButton = new System.Windows.Controls.Button();
                                        timeLabelButton.Content = startTime + " " + meridian;
                                        timeLabelButton.SetValue(Grid.RowProperty, i);
                                        timeLabelButton.SetValue(Grid.ColumnProperty, j);
                                        timeLabelButton.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                        timeLabelButton.VerticalAlignment = VerticalAlignment.Center;
                                        timeLabelButton.Click += new RoutedEventHandler(ChangeTime_Click);
                                        timeTable.Children.Add(timeLabelButton);
                                        changedRoomNum.Add(classrooms[column / 2].RoomNum);
                                        changedStartTime.Add(startTime);
                                        changedEndTime.Add(endTimePure);
                                        changedRow.Add(row);
                                        changedColumn.Add(column);
                                        changedMeridian.Add(meridian);
                                        changedTimeTable.Add(timeTableD);
                                        count++;


                                        for (int q = 0; q < classList.Count(); q++) // Changes class to have stating and ending time equal to user made changes on button label
                                        {
                                            bool breakLoop = false;
                                            bool changeClass = true;
                                            for (int d = 0; d < times_Default.Count(); d++)
                                            {
                                                if (classList[q].ClassDay == "MWF" && timeTableD == 0 && times_Default_Timetable[d] == "MWF" || classList[q].ClassDay == "TR" && timeTableD == 1 && times_Default_Timetable[d] == "TR")
                                                {
                                                    if (classList[q].Classroom.RoomNum == classrooms[column / 2].RoomNum && times_Default_Room[d] == classList[q].Classroom.RoomNum)
                                                    {
                                                        if (times_Default_Row[d] == i)
                                                        {
                                                            if (times_Default[d].Start == classList[q].StartTime.Start && times_Default[d].Meridian == classList[q].StartTime.Meridian)
                                                            {
                                                                if (classList[q].StartTime.Time != startTime || classList[q].StartTime.Time == startTime && meridian != classList[q].StartTime.Meridian)
                                                                {
                                                                    for (int k = 0; k < changedClasses.Count(); k++)
                                                                    {
                                                                        if (changedClasses[k] == classList[q].TextBoxName)
                                                                        {
                                                                            changeClass = false;
                                                                        }
                                                                    }
                                                                    if (changeClass == true)
                                                                    {
                                                                        classList[q].StartTime = new Timeslot(startTime, endTimePure, meridian);
                                                                        changedClasses.Add(classList[q].TextBoxName);
                                                                        breakLoop = true;
                                                                        break;
                                                                    }
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                            }

                                            if (breakLoop == true)
                                            {
                                                break;
                                            }

                                        }


                                        char removeChar = ':';
                                        string[] time = startTime.Split(removeChar);
                                        string startTimeLabel = time[0];
                                        timeLabelButton.Name = timeTable.Name + startTimeLabel + "_" + classrooms[j / 2].ClassID + meridian;
                                        labelNames.Add(timeLabelButton.Name);
                                        labelColumn.Add(j + 1);
                                        labelRow.Add(i);
                                        labelTimeTable.Add(parentGrid.Name);
                                        labelCount++;

                                        System.Windows.Controls.Button timeLabelButtonAdd = new System.Windows.Controls.Button();
                                        timeLabelButtonAdd.Content = "Add Row";
                                        timeLabelButtonAdd.SetValue(Grid.RowProperty, i + 1);
                                        timeLabelButtonAdd.SetValue(Grid.ColumnProperty, j);
                                        timeLabelButtonAdd.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                        timeLabelButtonAdd.VerticalAlignment = VerticalAlignment.Center;
                                        timeLabelButtonAdd.Background = new SolidColorBrush(Colors.Thistle);
                                        timeLabelButtonAdd.Click += new RoutedEventHandler(AddTimeslotsClick);
                                        timeTable.Children.Add(timeLabelButtonAdd);
                                    }
                                    else
                                    {
                                        if (i <= rowLength && parentGrid.Name == "MWF")
                                        {
                                            if (i <= times.Length)
                                            {
                                                System.Windows.Controls.Button timeLabelButton = new System.Windows.Controls.Button();
                                                timeLabelButton.Content = times[i - 1].Time + " " + times[i - 1].Meridian;
                                                timeLabelButton.Name = timeTable.Name + times[i - 1].TimeID + "_" + classrooms[j / 2].ClassID + times[i - 1].Meridian;
                                                timeLabelButton.SetValue(Grid.RowProperty, i);
                                                timeLabelButton.SetValue(Grid.ColumnProperty, j);
                                                timeLabelButton.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                                timeLabelButton.VerticalAlignment = VerticalAlignment.Center;
                                                timeLabelButton.Click += new RoutedEventHandler(ChangeTime_Click);
                                                timeTable.Children.Add(timeLabelButton);
                                                labelNames.Add(timeLabelButton.Name);
                                                labelColumn.Add(j + 1);
                                                labelRow.Add(i);
                                                labelTimeTable.Add(parentGrid.Name);
                                                labelCount++;

                                                System.Windows.Controls.Button timeLabelButtonAdd = new System.Windows.Controls.Button();
                                                timeLabelButtonAdd.Content = "Add Row";
                                                timeLabelButtonAdd.SetValue(Grid.RowProperty, i + 1);
                                                timeLabelButtonAdd.SetValue(Grid.ColumnProperty, j);
                                                timeLabelButtonAdd.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                                timeLabelButtonAdd.VerticalAlignment = VerticalAlignment.Center;
                                                timeLabelButtonAdd.Background = new SolidColorBrush(Colors.Thistle);
                                                timeLabelButtonAdd.Click += new RoutedEventHandler(AddTimeslotsClick);
                                                timeTable.Children.Add(timeLabelButtonAdd);
                                            }

                                        }
                                        else if (i <= rowLength && parentGrid.Name == "TR")
                                        {
                                            if (i <= times.Length)
                                            {
                                                System.Windows.Controls.Button timeLabelButton = new System.Windows.Controls.Button();
                                                timeLabelButton.Content = times[i - 1].Time + " " + times[i - 1].Meridian;
                                                timeLabelButton.Name = timeTable.Name + times[i - 1].TimeID + "_" + classrooms[j / 2].ClassID + times[i - 1].Meridian;
                                                timeLabelButton.SetValue(Grid.RowProperty, i);
                                                timeLabelButton.SetValue(Grid.ColumnProperty, j);
                                                timeLabelButton.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                                timeLabelButton.VerticalAlignment = VerticalAlignment.Center;
                                                timeLabelButton.Click += new RoutedEventHandler(ChangeTime_Click);
                                                timeTable.Children.Add(timeLabelButton);
                                                labelNames.Add(timeLabelButton.Name);
                                                labelColumn.Add(j + 1);
                                                labelRow.Add(i);
                                                labelTimeTable.Add(parentGrid.Name);
                                                labelCount++;

                                                System.Windows.Controls.Button timeLabelButtonAdd = new System.Windows.Controls.Button();
                                                timeLabelButtonAdd.Content = "Add Row";
                                                timeLabelButtonAdd.SetValue(Grid.RowProperty, i + 1);
                                                timeLabelButtonAdd.SetValue(Grid.ColumnProperty, j);
                                                timeLabelButtonAdd.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                                timeLabelButtonAdd.VerticalAlignment = VerticalAlignment.Center;
                                                timeLabelButtonAdd.Background = new SolidColorBrush(Colors.Thistle);
                                                timeLabelButtonAdd.Click += new RoutedEventHandler(AddTimeslotsClick);
                                                timeTable.Children.Add(timeLabelButtonAdd);
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }

                for (int n = 1; n <= classrooms.Count; n++)
                {
                    if (i == 0) // Add column titles (Classroom Bldg-Number)
                    {
                        if (n == 1)
                        {
                            System.Windows.Controls.Label classLabel = new System.Windows.Controls.Label();
                            classLabel.Content = classrooms[n - 1].Location + "-" + classrooms[n - 1].RoomNum;
                            classLabel.SetValue(Grid.RowProperty, 0);
                            classLabel.SetValue(Grid.ColumnProperty, n);
                            classLabel.HorizontalAlignment = System.Windows.HorizontalAlignment.Center;
                            classLabel.VerticalAlignment = VerticalAlignment.Center;
                            timeTable.Children.Add(classLabel);

                            System.Windows.Controls.Button labelBtn_Click = new System.Windows.Controls.Button();
                            labelBtn_Click.Content = classLabel.Content;
                            labelBtn_Click.SetValue(Grid.RowProperty, 0);
                            labelBtn_Click.SetValue(Grid.ColumnProperty, n);
                            labelBtn_Click.HorizontalAlignment = System.Windows.HorizontalAlignment.Center;
                            labelBtn_Click.VerticalAlignment = VerticalAlignment.Center;
                            labelBtn_Click.Background = Brushes.DarkSeaGreen;
                            labelBtn_Click.BorderThickness = new Thickness(1, 1, 1, 1);
                            labelBtn_Click.BorderBrush = new SolidColorBrush(Color.FromRgb(0, 0, 0));
                            labelBtn_Click.Click += new RoutedEventHandler(ViewRoomInfo_Click);

                            timeTable.Children.Add(labelBtn_Click);
                        }
                        else if (n >= 2)
                        {
                            System.Windows.Controls.Label classLabel = new System.Windows.Controls.Label();
                            classLabel.Content = classrooms[n - 1].Location + "-" + classrooms[n - 1].RoomNum;
                            classLabel.SetValue(Grid.RowProperty, 0);
                            classLabel.SetValue(Grid.ColumnProperty, n + n - 1);
                            classLabel.HorizontalAlignment = System.Windows.HorizontalAlignment.Center;
                            classLabel.VerticalAlignment = VerticalAlignment.Center;
                            timeTable.Children.Add(classLabel);

                            System.Windows.Controls.Button labelBtn_Click = new System.Windows.Controls.Button();
                            labelBtn_Click.Content = classLabel.Content;
                            labelBtn_Click.SetValue(Grid.RowProperty, 0);
                            labelBtn_Click.SetValue(Grid.ColumnProperty, n + n - 1);
                            labelBtn_Click.HorizontalAlignment = System.Windows.HorizontalAlignment.Center;
                            labelBtn_Click.VerticalAlignment = VerticalAlignment.Center;
                            labelBtn_Click.Background = Brushes.DarkSeaGreen;
                            labelBtn_Click.BorderThickness = new Thickness(1, 1, 1, 1);
                            labelBtn_Click.BorderBrush = new SolidColorBrush(Color.FromRgb(0, 0, 0));
                            labelBtn_Click.Click += new RoutedEventHandler(ViewRoomInfo_Click);

                            timeTable.Children.Add(labelBtn_Click);
                        }

                    }
                    else // Add empty timeslots
                    {
                        for (int p = 0; p < labelCount; p++)
                        {

                            for (int q = 0; q < p; q++)
                            {
                                if (p < labelCount)
                                {
                                    if (labelRow[p] == labelRow[q])
                                    {
                                        if (labelColumn[p] == labelColumn[q])
                                        {
                                            if (labelTimeTable[p] == labelTimeTable[q])
                                            {

                                                labelRow.RemoveAt(q);
                                                labelColumn.RemoveAt(q);
                                                labelNames.RemoveAt(q);
                                                labelTimeTable.RemoveAt(q);
                                                labelCount--;

                                            }
                                        }
                                    }
                                }
                                else 
                                {
                                    break;
                                }

                            }

                        }

                        for (int g = 0; g < labelCount; g++)
                        {
                            if (labelRow[g] == i)
                            {
                                if (labelColumn[g] == n + n - 1)
                                {
                                    if (parentGrid.Name == labelTimeTable[g])
                                    {
                                        System.Windows.Controls.Label emptySlot = new System.Windows.Controls.Label();
                                        string lbl_name = labelNames[g]; // IMPORTANT: Empty timeslots naming convention = GridName_TimeID_ClassID
                                        emptySlot.Name = lbl_name;
                                        //MessageBox.Show(emptySlot.Name); // DEBUG
                                        emptySlot.Content = "";
                                        emptySlot.AllowDrop = true;
                                        emptySlot.Drop += new System.Windows.DragEventHandler(HandleDropToCell);
                                        emptySlot.Style = Resources["DragLabel"] as System.Windows.Style;
                                        emptySlot.SetValue(Grid.RowProperty, i);
                                        emptySlot.SetValue(Grid.ColumnProperty, n + n - 1);
                                        emptySlot.HorizontalContentAlignment = System.Windows.HorizontalAlignment.Center;
                                        emptySlot.VerticalContentAlignment = VerticalAlignment.Center;
                                        emptySlot.BorderThickness = new Thickness(1, 1, 1, 1);
                                        emptySlot.BorderBrush = new SolidColorBrush(Color.FromRgb(0, 0, 0));
                                        emptySlot.MinWidth = 75;
                                        emptySlot.Background = new SolidColorBrush(Color.FromRgb(255, 255, 255));
                                        emptySlot.Margin = new Thickness(5);
                                        emptySlot.ContextMenu = null;
                                        //emptySlot.ContextMenu = Resources["ClassContextMenu"] as ContextMenu;
                                        object o = FindName(lbl_name);
                                        if (o != null)
                                        {
                                            UnregisterName(lbl_name);
                                        }
                                        RegisterName(lbl_name, emptySlot);
                                        timeTable.Children.Add(emptySlot);

                                    }
                                }
                            }
                        }
                    }
                }
            }

            // Add the grid to the MWF_Schedule Grid
            object x = FindName(timeTableName);
            if (x != null)
            {
                UnregisterName(timeTableName);
            }
            RegisterName(timeTableName, timeTable);
            parentGrid.Children.Add(timeTable);

            // Populate the empty timeslots with our available information
            DeleteDuplicateTimes();
            FillDerivedLists();
            PopulateTimeTable(timeTable, times);
            

        }
        public void DeleteDuplicateTimes() // Shortens saved value arrays by removing indexes that have been changed more than once
        {

            int removeIndex;
            for (int c = 0; c < count; c++)
            {
                for (int k = 1; k < count; k++)
                {
                    if (changedRoomNum[c] == changedRoomNum[k])
                    {

                        if (changedRow[c] == changedRow[k])
                        {
                            if (changedTimeTable[c] == changedTimeTable[k])
                            {
                                if (changedEndTime[c] == changedEndTime[k])
                                {
                                    if (k > c)
                                    {
                                        removeIndex = c;
                                        changedRoomNum.RemoveAt(c);
                                        changedStartTime.RemoveAt(c);
                                        changedEndTime.RemoveAt(c);
                                        changedRow.RemoveAt(c);
                                        changedColumn.RemoveAt(c);
                                        changedMeridian.RemoveAt(c);
                                        changedTimeTable.RemoveAt(c);
                                        count--;
                                    }
                                }
                            }
                        }
                    }
                }
            }

            for (int i = 0; i < times_Default.Count(); i++)
            {
                for (int k = 1; k < times_Default.Count(); k++)
                {
                    if (times_Default_Room[i] == times_Default_Room[k])
                    {
                        
                            if(times_Default_Timetable[i] == times_Default_Timetable[k])
                            {
                                if (times_Default[i].Start == times_Default[k].Start)
                                {
                                    if (k > i)
                                    {
                                        times_Default.RemoveAt(i);
                                        times_Default_Room.RemoveAt(i);
                                        times_Default_Timetable.RemoveAt(i);
                                        removeIndex = i;
                                        times_Default_Row.RemoveAt(i);
                                    }
                                }
                            }
                        
                    }
                }
            }

            for (int c = 0; c < autoCount; c++)
            {
                for (int k = 1; k < autoCount; k++)
                {
                    if (autoChangedRoomNum[c] == autoChangedRoomNum[k])
                    {
                        if (autoChangedRow[c] == autoChangedRow[k])
                        {
                            if (autoChangedTimeTable[c] == autoChangedTimeTable[k])
                            {
                                if (autoChangedEndTime[c] == autoChangedEndTime[k])
                                {
                                    if (k > c)
                                    {
                                        removeIndex = c;
                                        autoChangedRoomNum.RemoveAt(c);
                                        autoChangedStartTime.RemoveAt(c);
                                        autoChangedEndTime.RemoveAt(c);
                                        autoChangedRow.RemoveAt(c);
                                        autoChangedColumn.RemoveAt(c);
                                        autoChangedMeridian.RemoveAt(c);
                                        autoChangedTimeTable.RemoveAt(c);
                                        autoCount--;
                                    }
                                }
                            }
                        }
                    }
                }
            }
            for (int c = 0; c < labelCount; c++)
            {

                for (int k = 1; k < labelCount; k++)
                {
                    if (labelRow[c] == labelRow[k])
                    {
                        if (labelColumn[c] == labelColumn[k])
                        {
                            if (labelTimeTable[c] == labelTimeTable[k])
                            {
                                if (k > c)
                                {
                                    
                                    labelRow.RemoveAt(c);
                                    labelColumn.RemoveAt(c);
                                    labelNames.RemoveAt(c);
                                    labelTimeTable.RemoveAt(c);
                                    labelCount--;
                                }

                            }
                        }
                    }
                }

            }


            for (int i = 0; i < changedStartTime.Count(); i++)
            {
                if (changedStartTime[i] != null)
                {
                    masterTimeslotList.Add(new Timeslot(changedStartTime[i], changedEndTime[i], changedMeridian[i]));
                    masterTimetableList.Add(changedTimeTable[i]);
                    masterClassRoomList.Add(changedRoomNum[i]);
                }
            }

            for (int i = 0; i < autoChangedStartTime.Count(); i++)
            {
                if (autoChangedStartTime[i] != null)
                {
                    masterTimeslotList.Add(new Timeslot(autoChangedStartTime[i], autoChangedEndTime[i], autoChangedMeridian[i]));
                    masterTimetableList.Add(autoChangedTimeTable[i]);
                    masterClassRoomList.Add(autoChangedRoomNum[i]);
                }
            }

            for (int i = 0; i < masterTimeslotList.Count(); i++)
            {
                for (int k = 1; k < masterTimetableList.Count(); k++)
                {
                    if (masterTimetableList[i] == masterTimetableList[k])
                    {

                        if (masterClassRoomList[i] == masterClassRoomList[k])
                        {
                            if (masterTimeslotList[i].Start == masterTimeslotList[k].Start)
                            {
                                if (k > i)
                                {
                                    masterTimeslotList.RemoveAt(i);
                                    masterClassRoomList.RemoveAt(i);
                                    masterTimetableList.RemoveAt(i);
                                }
                            }
                        }
                    }
                }
            }


        }
        public void PopulateTimeTable(Grid timeTable, Timeslot[] times) // Populate a GUI grid based on classList 
        {
 
            string days = "";
            if (times.Length == times_MWF.Length)
            {
                days = "MWF";
            }
            else
            {
                days = "TR";
            }
            for (int i = 0; i < classList.Count; i++)
            {
                if (classList[i].ClassDay == days)
                {
                    if (classList[i].StartTime.TimeID != "--" && classList[i].Classroom.Location != "N/A" && !classList[i].Online)
                    {
                        
                        string targetBoxID = days + '_' + classList[i].StartTime.TimeID + '_' + classList[i].Classroom.ClassID + classList[i].StartTime.Meridian;
                        System.Windows.Controls.Label lbl = (System.Windows.Controls.Label)FindName(targetBoxID);

                      try
                      {


                          if (lbl != null)
                        {
                            if (lbl.Content.ToString() == "" || lbl.Content.ToString() == classList[i].TextBoxName)
                            {
                                if (!DetermineTimeConflict(classList[i], days, classList[i].StartTime.TimeID, classList[i].StartTime.Meridian))
                                {
                                    lbl.Content = classList[i].TextBoxName;
                                    if (classList[i].isHidden)
                                    {
                                        lbl.Background = stripedBackground(classList[i].Prof.profRGB.colorBrush);
                                    }
                                    else
                                    {
                                        lbl.Background = classList[i].Prof.Prof_Color;
                                    }
                                    lbl.Tag = classList[i].ClassID;
                                    lbl.ContextMenu = Resources["ClassContextMenuGUI"] as System.Windows.Controls.ContextMenu;
                                    lbl.ToolTip = classList[i].ToolTipText;
                                    classList[i].isAssigned = true;
                                }
                                else
                                { 
                                    errorMSG = "Conflict: " + classList[i].DeptName + " " + classList[i].ClassNumber + " At: " + classList[i].ClassDay + " " + classList[i].StartTime.FullTime + " Professor is already teaching at that time or Time interupts another class Time! Moving class to unassigned classes list";
                                    WriteErrorLog(errorMSG);
                                    classList[i].Classroom = new ClassRoom();
                                    classList[i].StartTime = new Timeslot();
                                    classList[i].isAssigned = false;
                                    
                                };
                            }
                            else
                            {
                                // MessageBoxButton button = MessageBoxButton.OK;
                                // MessageBoxImage icon = MessageBoxImage.Exclamation;
                                //  System.Windows.MessageBox.Show("Conflict: " + classList[i].DeptName + " " + classList[i].ClassNumber + "\nAt: " + classList[i].ClassDay + " " + classList[i].StartTime.FullTime + "\nTimeslot already taken!\n\nMoving class to unassigned classes list", "Timeslot Conflict", button, icon);
                                errorMSG = "Conflict: " + classList[i].DeptName + " " + classList[i].ClassNumber +
                                       " At: " + classList[i].ClassDay + " " + classList[i].StartTime.FullTime +
                                       " Timeslot already taken! Moving class to unassigned classes list";
                                WriteErrorLog(errorMSG);
                                classList[i].Classroom = new ClassRoom();
                                classList[i].StartTime = new Timeslot();
                                classList[i].isAssigned = false;
                            }
                        }
                      }
                      catch (Exception ex)
                      {
                            //Time slot not on grid so no label is added but the time is still changed
                            
                      }

                    }
                }
                else if (days == "MWF" && classList[i].ClassDay == "M" || days == "TR" && classList[i].ClassDay == "T" || days == "MWF" && classList[i].ClassDay == "W" || days == "TR" && classList[i].ClassDay == "R" || days == "MWF" && classList[i].ClassDay == "F")
                {
                    classList[i].isAssigned = true;

                    if (classList[i].StartTime.TimeID != "--" && classList[i].Classroom.Location != "N/A" && !classList[i].Online)
                    {

                        string timeLength = classList[i].StartTime.FullTime;
                        string[] bothTimes = timeLength.Split('-');
                        string startingTime = bothTimes[0];
                        startingTime = startingTime.Substring(0, 2);
                        string endingTime = bothTimes[1];
                        endingTime = endingTime.Substring(1, 2);

                        int startTime = int.Parse(startingTime);
                        int endTime = int.Parse(endingTime);

                        string meridian = classList[i].StartTime.Meridian;
                        if (startTime <= endTime)
                        {

                            for (int r = startTime; r <= endTime; r++)
                            {

                           
                                string TimeID = "";
                                if (r < 10)
                                {
                                    TimeID = "0" + r.ToString();
                                }
                                
                                else
                                {
                                    TimeID = r.ToString();
                                }

                                if (startTime < 12 && r >= 12)
                                {
                                    if (meridian == "AM")
                                    {
                                        meridian = "PM";
                                    }
                                    else if (meridian == "PM")
                                    {
                                            meridian = "AM";
                                    }
                                }

                                string targetBoxID = days + '_' + TimeID + '_' + classList[i].Classroom.ClassID + meridian;
                                System.Windows.Controls.Label lbl = (System.Windows.Controls.Label)FindName(targetBoxID);
                                if (lbl != null)
                                {
                                    if (lbl.Content.ToString() == "" || lbl.Content.ToString() == "Single Day Class")
                                    {
                                        if (!DetermineTimeConflict(classList[i], days, classList[i].StartTime.TimeID, classList[i].StartTime.Meridian))
                                        {
                                            lbl.Content = classList[i].TextBoxName;
                                            if (classList[i].isHidden)
                                            {
                                                lbl.Background = stripedBackground(classList[i].Prof.profRGB.colorBrush);
                                            }
                                            else
                                            {
                                                lbl.Background = classList[i].Prof.Prof_Color;
                                            }
                                            lbl.Tag = classList[i].ClassID;
                                            lbl.ContextMenu = Resources["ClassContextMenuGUI"] as System.Windows.Controls.ContextMenu;
                                            lbl.ToolTip = classList[i].ToolTipText;
                                            classList[i].isAssigned = true;
                                        }
                                        else
                                        {
                                            // MessageBoxButton button = MessageBoxButton.OK;
                                            //essageBoxImage icon = MessageBoxImage.Exclamation;
                                            //System.Windows.MessageBox.Show("Conflict: " + classList[i].DeptName + " " + classList[i].ClassNumber + "\nAt: " + classList[i].ClassDay + " " + classList[i].StartTime.FullTime +"\nProfessor is already teaching at that time!\n\nMoving class to unassigned classes list", "Timeslot Conflict", button, icon);
                                            errorMSG = "Conflict: " + classList[i].DeptName + " " + classList[i].ClassNumber + " At: " + classList[i].ClassDay + " " + classList[i].StartTime.FullTime + " Professor is already teaching at that time or Time interupts another class Time! Moving class to unassigned classes list";
                                            WriteErrorLog(errorMSG);
                                            classList[i].Classroom = new ClassRoom();
                                            classList[i].StartTime = new Timeslot();
                                            classList[i].isAssigned = false;
                                            
                                        }
                                    }
                                    else
                                    {
                                        // MessageBoxButton button = MessageBoxButton.OK;
                                        // MessageBoxImage icon = MessageBoxImage.Exclamation;
                                        //  System.Windows.MessageBox.Show("Conflict: " + classList[i].DeptName + " " + classList[i].ClassNumber + "\nAt: " + classList[i].ClassDay + " " + classList[i].StartTime.FullTime + "\nTimeslot already taken!\n\nMoving class to unassigned classes list", "Timeslot Conflict", button, icon);
                                        errorMSG = "Conflict: " + classList[i].DeptName + " " + classList[i].ClassNumber +
                                               " At: " + classList[i].ClassDay + " " + classList[i].StartTime.FullTime +
                                               " Timeslot already taken! Moving class to unassigned classes list";
                                        WriteErrorLog(errorMSG);
                                        classList[i].Classroom = new ClassRoom();
                                        classList[i].StartTime = new Timeslot();
                                        classList[i].isAssigned = false;
                                    }
                                }
                            }
                        }
                        else
                        {
                            for (int r = startTime; r < endTime + 12; r++)
                            {
                                string TimeID = "";
                                if (r < 10)
                                {
                                    TimeID = "0" + r.ToString();
                                }
                                else if (r > 12)
                                {
                                    TimeID = (r - 12).ToString();
                                }
                                else
                                {
                                    TimeID = r.ToString();
                                }
                                if (startTime < 12 && r >= 12)
                                {
                                    if (meridian == "AM")
                                    {
                                        meridian = "PM";
                                    }
                                    else if (meridian == "PM")
                                    {
                                        meridian = "AM";
                                    }
                                }
                                string targetBoxID = days + '_' + TimeID + '_' + classList[i].Classroom.ClassID + meridian;
                                System.Windows.Controls.Label lbl = (System.Windows.Controls.Label)FindName(targetBoxID);
                                if (lbl != null)
                                {
                                    if (lbl.Content.ToString() == "" || lbl.Content.ToString() == "Single Day Class")
                                    {
                                        if (!DetermineTimeConflict(classList[i], days, classList[i].StartTime.TimeID, classList[i].StartTime.Meridian))
                                        {
                                            lbl.Content = classList[i].TextBoxName;
                                            if (classList[i].isHidden)
                                            {
                                                lbl.Background = stripedBackground(classList[i].Prof.profRGB.colorBrush);
                                            }
                                            else
                                            {
                                                lbl.Background = classList[i].Prof.Prof_Color;
                                            }
                                            lbl.Tag = classList[i].ClassID;
                                            lbl.ContextMenu = Resources["ClassContextMenuGUI"] as System.Windows.Controls.ContextMenu;
                                            lbl.ToolTip = classList[i].ToolTipText;
                                            classList[i].isAssigned = true;

                                        }
                                        else
                                        {
                                            // MessageBoxButton button = MessageBoxButton.OK;
                                            //essageBoxImage icon = MessageBoxImage.Exclamation;
                                            //System.Windows.MessageBox.Show("Conflict: " + classList[i].DeptName + " " + classList[i].ClassNumber + "\nAt: " + classList[i].ClassDay + " " + classList[i].StartTime.FullTime +"\nProfessor is already teaching at that time!\n\nMoving class to unassigned classes list", "Timeslot Conflict", button, icon);
                                            errorMSG = "Conflict: " + classList[i].DeptName + " " + classList[i].ClassNumber + " At: " + classList[i].ClassDay + " " + classList[i].StartTime.FullTime + " Professor is already teaching at that time or Time interupts another class Time! Moving class to unassigned classes list";
                                            WriteErrorLog(errorMSG);
                                            classList[i].Classroom = new ClassRoom();
                                            classList[i].StartTime = new Timeslot();
                                            classList[i].isAssigned = false;
                                            
                                        }
                                    }
                                    else
                                    {
                                        // MessageBoxButton button = MessageBoxButton.OK;
                                        // MessageBoxImage icon = MessageBoxImage.Exclamation;
                                        //  System.Windows.MessageBox.Show("Conflict: " + classList[i].DeptName + " " + classList[i].ClassNumber + "\nAt: " + classList[i].ClassDay + " " + classList[i].StartTime.FullTime + "\nTimeslot already taken!\n\nMoving class to unassigned classes list", "Timeslot Conflict", button, icon);
                                        errorMSG = "Conflict: " + classList[i].DeptName + " " + classList[i].ClassNumber +
                                               " At: " + classList[i].ClassDay + " " + classList[i].StartTime.FullTime +
                                               " Timeslot already taken! Moving class to unassigned classes list";
                                        WriteErrorLog(errorMSG);
                                        classList[i].Classroom = new ClassRoom();
                                        classList[i].StartTime = new Timeslot();
                                        classList[i].isAssigned = false;
                                    }
                                }
                            }
                        }
                    }
                    
                }
            }
            FillDerivedLists();

        }

        public void EmptyGrid(Grid timetable)  // Empties all entries of a GUI grid 
        {
            UIElementCollection items = timetable.Children;
            for (int i = 0; i < items.Count; i++)
            {
                System.Windows.Controls.Label slot = items[i] as System.Windows.Controls.Label;
                if (slot != null && slot.Tag != null)
                {
                    slot.Content = "";
                    RGB_Color white_bg = new RGB_Color(255, 255, 255);
                    slot.Background = white_bg.colorBrush2;
                    slot.Tag = null;
                    slot.ContextMenu = null;
                }
            }
        }
        public void FillDerivedLists() // Fill Unassigned/Online/APPT/APPT2 lists. (They are subsets of classList) 
        {
            // empty online and unassigned class lists
            unassignedClasses.Clear();
            singleDayClasses.Clear();
            onlineClasses.Clear();
            appointmentClasses.Clear();
            appointment2Classes.Clear();
            // add from classList
            for (int i = 0; i < classList.Count; i++)
            {
                if (!classList[i].isAssigned)
                {
                    if (classList[i].Online)
                    {
                        onlineClasses.Add(classList[i]);
                    }
                    else if (classList[i].isAppointment)
                    {
                        if (classList[i].Classroom.Location == "APPT")
                        {
                            appointmentClasses.Add(classList[i]);
                        }
                        else if (classList[i].Classroom.Location == "APPT2")
                        {
                            appointment2Classes.Add(classList[i]);
                        }
                        else
                        {
                            //System.Windows.MessageBox.Show("DEBUG - ERROR: Couldnt assign appointed class to either APPT or APPT2");
                            errorMSG = "DEBUG - ERROR: Couldnt assign class " + classList[i] + " to either APPT or APPT2";
                            WriteErrorLog(errorMSG);
                        }
                    }
                    else
                    {
                        //MessageBox.Show("fillUnassigned() -> Adding " + classList[i].TextBoxName + " to unassigned list.");
                        unassignedClasses.Add(classList[i]);
                        
                    }
                }
                else if (classList[i].ClassDay == "M" || classList[i].ClassDay == "T" || classList[i].ClassDay == "W" || classList[i].ClassDay == "R" || classList[i].ClassDay == "F")
                {
                    singleDayClasses.Add(classList[i]);
                }
            }
        }
        public void AssignProfColors() // Give professors a color key based on the palette defined above + Save assigned colors to XML file 
        {
            //MessageBox.Show("ColorIndex is currently: " + Settings.Default.ColorIndex);
            // Read from Colors file to see which professors we have already assigned a color. Store in colorPairings List.
            string tempPath = System.IO.Path.GetTempPath();
            string filename = "ColorConfigurations22.xml";
            colorFilePath = System.IO.Path.Combine(tempPath, filename);
            XmlSerializer ser = new XmlSerializer(typeof(Pairs));
            if (!File.Exists(colorFilePath))
            {
                Settings.Default.Reset();
                colorPairs = new Pairs();
                colorPairs.ColorPairings = new List<ProfColors>();

                using (FileStream fs = new FileStream(colorFilePath, FileMode.OpenOrCreate))
                {
                    ser.Serialize(fs, colorPairs);
                }
            }
            using (FileStream fs = new FileStream(colorFilePath, FileMode.OpenOrCreate))
            {
                colorPairs = ser.Deserialize(fs) as Pairs;
            }
            // go through the professor array
            // if color not already set, add it based on next item on the palette (palette index is set at 0 the first time of execution on a user PC)
            for (int i = 0; i < professors.Count; i++)
            {
                bool found = false;
                for (int n = 0; n < colorPairs.ColorPairings.Count; n++)
                {
                    if (professors[i].FullName == colorPairs.ColorPairings[n].ProfName)
                    {
                        //MessageBox.Show("Found " + professors[i].FullName + "!");
                        found = true;
                        //MessageBox.Show("Reassigning " + colorPairs.ColorPairings[n].Color + " to " + professors[i].FullName + ".");
                        professors[i].profRGB = StringToRGB(colorPairs.ColorPairings[n].Color);
                        break;
                    }
                }
                if (!found)
                {
                    //MessageBox.Show("Adding " + professors[i].FullName + "!");
                    // Give professor a colour
                    int paletteIndex = Settings.Default.ColorIndex;
                    if (paletteIndex < colorPalette.Length)
                    {
                        professors[i].profRGB = colorPalette[paletteIndex];
                        //MessageBox.Show("Assigned: " + colorPalette[paletteIndex].colorString + "\nProfessor: " + professors[i].FullName);
                        paletteIndex++;
                        Settings.Default.ColorIndex = paletteIndex;
                    }
                    else
                    {
                        //MessageBox.Show("Random Color");
                        Random rand = new Random();
                        RGB_Color tempColor = new RGB_Color((byte)rand.Next(256), (byte)rand.Next(256), (byte)rand.Next(256));
                        while (isColorTaken(tempColor))
                        {
                            tempColor.R = (byte)rand.Next(256);
                            tempColor.G = (byte)rand.Next(256);
                            tempColor.B = (byte)rand.Next(256);
                        }
                        professors[i].profRGB = tempColor;
                    }
                    // Add it to pairings list
                    colorPairs.ColorPairings.Add(new ProfColors { ProfName = professors[i].FullName, Color = professors[i].profRGB.colorString });
                    //MessageBox.Show("Added " + professors[i].FullName + " + " + professors[i].profRGB.colorString);
                }
            }
            // Save changes to Colors.xml
            SerializePairs();
            // Save paletteIndex counter to application settings
            Settings.Default.Save();
            // Reassign colors to professors in classlist
            for (int i = 0; i < classList.Count; i++)
            {
                for (int n = 0; n < colorPairs.ColorPairings.Count; n++)
                {
                    if (classList[i].Prof.FullName == colorPairs.ColorPairings[n].ProfName)
                    {
                        classList[i].Prof.profRGB = StringToRGB(colorPairs.ColorPairings[n].Color);
                        break;
                    }
                }
            }
        }
        public void UpdatePairs() // Update colorPairs list to account for any additions
        {
            for (int i = 0; i < professors.Count; i++)
            {
                for (int n = 0; n < colorPairs.ColorPairings.Count; n++)
                {
                    if (colorPairs.ColorPairings[n].ProfName == professors[i].FullName)
                    {
                        break;
                    }
                    if (n == (colorPairs.ColorPairings.Count - 1))
                    {
                        //MessageBox.Show("Adding new color pair...\n\nProfessor: " + professors[i].FullName + "\nColor: " + professors[i].profRGB.colorString);
                        // Add prof + color pairing
                        colorPairs.ColorPairings.Add(new ProfColors { ProfName = professors[i].FullName, Color = professors[i].profRGB.colorString });
                    }
                }
            }
        }
        public void SerializePairs() // Save professor/color pairs to XML file 
        {
            // Update colorPairs to account for any new professors
            UpdatePairs();
            // Save changes to Colors.xml
            XmlSerializer ser = new XmlSerializer(typeof(Pairs));
            using (FileStream fs = new FileStream(colorFilePath, FileMode.OpenOrCreate))
            {
                ser.Serialize(fs, colorPairs);
            }
        }
        public void UpdateProfessorCapacity() // Update professor's numClasses and numPrep values
        {
            // for each professor
            for (int i = 0; i < professors.Count; i++) // Wont be a problem. professors wont + or -
            {
                // Reset class counters
                professors[i].NumClasses = 0;
                professors[i].NumPrep = 0;
                List<string> uniqueClasses = new List<string>();
                // check how many classes they are teaching
                for (int n = 0; n < classList.Count; n++) // Wont be a problem. class has already been added previously
                {
                    if (professors[i].SRUID == classList[n].Prof.SRUID)
                    {
                        bool unique = true;
                        
                        for (int j = 0; j < uniqueClasses.Count; j++)
                        {
                            if (uniqueClasses[j] == classList[n].ClassName)
                            {
                                unique = false;
                            }
                        }

                        if (unique && !classList[n].excludeCredits)
                        {
                            if (classList[n].isCrossListed)
                            {
                                if (classList[n].isCrossFirst == true)
                                {
                                    professors[i].NumPrep++;
                                    professors[i].NumClasses += classList[n].Credits;
                                }
                                uniqueClasses.Add(classList[n].ClassName);
                            }
                            else
                            {
                                professors[i].NumPrep++;
                                professors[i].NumClasses += classList[n].Credits;
                                uniqueClasses.Add(classList[n].ClassName);
                            }
                        }
                        else if (!unique && !classList[n].excludeCredits)
                        {
                            if (classList[n].isCrossListed)
                            {
                                if (classList[n].isCrossFirst == true)
                                {
                                    professors[i].NumClasses += classList[n].Credits;
                                }
                            }
                            else
                            {
                                professors[i].NumClasses += classList[n].Credits;
                            }
                        }
                    }
                }
            }
        }
        public void BindData() // Bind class/professor lists to GUI data tables 
        {
            Online_Classes_Grid.ItemsSource = onlineClasses; // Online classes GUI list
            Unassigned_Classes_Grid.ItemsSource = unassignedClasses; // Unassigned classes GUI list
            SingleDay_Classes_Grid.ItemsSource = singleDayClasses;
            Appointment_Classes_Grid.ItemsSource = appointmentClasses; // APPT classes GUI list
            Appointment2_Classes_Grid.ItemsSource = appointment2Classes; // APPT2 classes GUI list
            Professor_Key_List.ItemsSource = professors; // Professor Key GUI list
            Full_Classes_Grid.ItemsSource = classList;  // Classes GUI list (Classes tab)
            Full_Professors_Grid.ItemsSource = professors; // Professors GUI list (Professors tab)
        }
        public void RefreshGUI() // Empty GUI timetables, repopulate them and refresh derived lists 
        {
            Grid timetable_MWF = (Grid)FindName("MWF_");
            Grid timetable_TR = (Grid)FindName("TR_");
            EmptyGrid(timetable_MWF);
            EmptyGrid(timetable_TR);
            PopulateTimeTable(timetable_MWF, times_MWF);
            PopulateTimeTable(timetable_TR, times_TR);
            FillDerivedLists();
            UpdateProfessorCapacity();
            ProcessProfessorPreferences();
            ProcessClassGroupings();
        }


        public void SaveRoomHeaderFile()
        {
            //variables
            string bldg, notes, roomNum, capacity;

            //read in headers file
            string fileName = "Headers.xlsx";
            FileInfo f = new FileInfo(fileName);
            string name = f.FullName;
            var headerFile = new XLWorkbook(name);
            var classroomInfo = headerFile.Worksheet("ClassroomInfo");
            int numberRoomHeaders = classroomInfo.RowsUsed().Count();
            System.Data.DataTable dt = new System.Data.DataTable();

            //read headers from classroom headers sheet into list
            var headerList = new List<string>();
            for (int i = 0; i < numberRoomHeaders; i++)
            {
                headerList.Add(classroomInfo.Row(1).Cell(i + 1).GetValue<string>());
            }

            //Add Columns
            for (int i = 0; i < headerList.Count-1; i++)
            {
                dt.Columns.Add(headerList[i]);
            }

            //add classroom info to the datatable
            foreach (var room in classrooms)
            {
                bldg = room.Location.ToString();
                roomNum = room.RoomNum.ToString();
                capacity = room.AvailableSeats.ToString();
                notes = room.Notes.ToString();

                if (notes == null)
                {
                    notes = "";
                }


                dt.Rows.Add(capacity, bldg, roomNum, notes);
            }


            //Setting Table detail references 
            dt.TableName = "ClassroomInfo";
            XLColor headerColor = XLColor.FromHtml("#FF016648");
            XLColor defaultColor = XLColor.White;
            int headerFontSize = 11;
            string headerFontName = "Calibri";

            //output to excel file
            headerFile.Worksheet("ClassroomInfo").Delete();
            headerFile.Worksheets.Add(dataTable: dt);
            var newWorksheet = headerFile.Worksheet("ClassroomInfo");

            for (int i = 0; i < headerList.Count; i++)
            {
                newWorksheet.Row(1).Style.Fill.BackgroundColor = headerColor;
                newWorksheet.Row(1).Style.Font.SetFontName(headerFontName);
                newWorksheet.Row(1).Style.Font.FontSize = headerFontSize;
                newWorksheet.Row(1).Style.Font.FontColor = XLColor.White;
                newWorksheet.Columns().AdjustToContents();
            }
            for (int i = 2; i < newWorksheet.RowsUsed().Count(); i++)
            {
                newWorksheet.Row(i).Style.Fill.BackgroundColor = defaultColor;
            }

            //save classroom info to headers file
            headerFile.SaveAs(name);

        }




        public void SaveProfHeaderFile()
        {
            //read in headers file
            string fileName = "Headers.xlsx";
            FileInfo f = new FileInfo(fileName);
            string name = f.FullName;
            var headerFile = new XLWorkbook(name);
            var professorInfo = headerFile.Worksheet("ProfessorInfo");
            int numberHeaders = professorInfo.ColumnsUsed().Count();
            System.Data.DataTable dt = new System.Data.DataTable();

            //read headers from classroom headers sheet into list
            var headerList = new List<string>();
            for (int i = 0; i < numberHeaders; i++)
            {
                headerList.Add(professorInfo.Row(1).Cell(i + 1).GetValue<string>());
            }

            //Add Columns
            for (int i = 0; i < headerList.Count; i++)
            {
                dt.Columns.Add(headerList[i]);
            }

            //add prof info to the datatable
            foreach (var prof in professors)
            {
                string profName = prof.FullName.ToString();
                string profID = prof.SRUID.ToString();
                int profCredits = prof.MaxClasses;
                int profPrep = prof.MaxPrep;

                dt.Rows.Add(profName, profID, profCredits, profPrep);
            }

            //Setting Table detail references 
            dt.TableName = "ProfessorInfo";
            XLColor headerColor = XLColor.FromHtml("#FF016648");
            XLColor defaultColor = XLColor.White;
            int headerFontSize = 11;
            string headerFontName = "Calibri";

            //output to excel file
            headerFile.Worksheet("ProfessorInfo").Delete();
            headerFile.Worksheets.Add(dataTable: dt);
            var newWorksheet = headerFile.Worksheet("ProfessorInfo");

            for (int i = 0; i < headerList.Count; i++)
            {
                newWorksheet.Row(1).Style.Fill.BackgroundColor = headerColor;
                newWorksheet.Row(1).Style.Font.SetFontName(headerFontName);
                newWorksheet.Row(1).Style.Font.FontSize = headerFontSize;
                newWorksheet.Row(1).Style.Font.FontColor = XLColor.White;
                newWorksheet.Columns().AdjustToContents();
            }
            for (int i = 2; i < newWorksheet.RowsUsed().Count(); i++)
            {
                newWorksheet.Row(i).Style.Fill.BackgroundColor = defaultColor;
            }

            //save classroom info to headers file
            headerFile.SaveAs(name);
        }


        public void SaveChanges() // Writes classList to an excel file 
        {
            /*
            string path = Directory.GetCurrentDirectory() + "\\times.xml";
            XmlTextWriter xml = new XmlTextWriter(path, System.Text.Encoding.UTF8);
            xml.Formatting = Formatting.Indented;
            xml.WriteStartDocument();
            xml.WriteStartElement("ChangedClassRooms");
            for (int i = 0; i < count; i++)
            {
                xml.WriteStartElement("Room number: " + changedRoomNum[i]);
                xml.WriteElementString("StartTime", changedStartTime[i]);
                xml.WriteElementString("Row", changedRow[i].ToString());
                xml.WriteElementString("TimeTable", changedTimeTable[i].ToString());
            }
            xml.WriteEndElement();
            xml.WriteStartElement("autoChangedClassRooms");
            for (int i = 0; i < autoCount; i++)
            {
                xml.WriteStartElement("Room number: " + autoChangedRoomNum[i]);
                xml.WriteElementString("StartTime", autoChangedStartTime[i]);
                xml.WriteElementString("Row", autoChangedRow[i].ToString());
                xml.WriteElementString("TimeTable", autoChangedTimeTable[i].ToString());
            }
            xml.WriteEndElement();
            xml.WriteEndDocument();
            xml.Flush();
            xml.Close();
            */


            string fileDir = getFileDirectory(System.Windows.Application.Current.Resources["FilePath"].ToString());
            Microsoft.Win32.SaveFileDialog saveFileDialog = new Microsoft.Win32.SaveFileDialog();
            saveFileDialog.InitialDirectory = fileDir;
            saveFileDialog.Filter = "Excel File (*.xlsx)|*.xlsx";
            saveFileDialog.Title = "Save Excel File";
            saveFileDialog.ShowDialog();

            if (saveFileDialog.FileName != "")
            {
                savefileName = saveFileDialog.FileName;
                XLWorkbook wb = new XLWorkbook();
                System.Data.DataTable dt = getDataTableFromClasses();
                var ws = wb.Worksheets.Add(dt);

                // Colors
                XLColor empty = XLColor.NoColor;
                XLColor header = XLColor.FromHtml("#FF016648");
                XLColor edited = XLColor.FromHtml("#FFFFCFCF");
                XLColor added = XLColor.FromHtml("#FFD4FFC4");

                // Styling
                ws.Table(0).Theme = XLTableTheme.None;



                /*
                ws.Column(7).AdjustToContents();
                ws.Column(22).AdjustToContents();
                ws.Column(23).AdjustToContents();
                */
                ws.Columns().AdjustToContents();
                ws.Column(5).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Right);

                for (int i = 0; i < columnCount; i++)
                {
                    ws.Column(i + 1).Style.Font.SetFontName(colFontName[i]);

                    ws.Column(i + 1).Style.Font.FontSize = colFontSize[i];

                    if (colFontStyle[i] == "Bold")
                    {
                        ws.Column(i + 1).Style.Font.Bold = true;
                    }
                    else if (colFontStyle[i] == "Italic")
                    {
                        ws.Column(i + 1).Style.Font.Italic = true;
                    }
                    else if (colFontStyle[i] == "Bold Italic")
                    {
                        ws.Column(i + 1).Style.Font.Bold = true;
                        ws.Column(i + 1).Style.Font.Italic = true;
                    }
                    else
                    {
                        ws.Column(i + 1).Style.Font.Bold = false;
                        ws.Column(i + 1).Style.Font.Italic = false;
                    }
                }


                ws.Row(1).Style.Fill.BackgroundColor = header;


                ws.Row(1).Style.Font.SetFontName(headerFontName);

                if (headerFontStyle == "Bold")
                {
                    ws.Row(1).Style.Font.Bold = true;
                }
                else if (headerFontStyle == "Italic")
                {
                    ws.Row(1).Style.Font.Italic = true;
                }
                else if (headerFontStyle == "Bold Italic")
                {
                    ws.Row(1).Style.Font.Bold = true;
                    ws.Row(1).Style.Font.Italic = true;
                }
                else
                {
                    ws.Row(1).Style.Font.Italic = false;
                    ws.Row(1).Style.Font.Bold = false;
                }

                ws.Row(1).Style.Font.FontSize = headerFontSize;

                ws.Row(1).Style.Font.FontColor = XLColor.White;

                // Iterate over classList to format the background of each row appropriately
                for (int i = 0; i < classList.Count; i++)
                {
                    ws.Row(i + 2).Style.Fill.BackgroundColor = empty;
                    for (int j = 0; j < classList[i].ChangedData.Count; j++)
                    {
                        if (classList[i].ChangedData[j])
                        {
                            ws.Row(i + 2).Style.Fill.BackgroundColor = edited;
                        }
                    }
                    /*
                    // match ClassID
                    for (int n = 0; n < hashedClasses.Count; n++)
                    {
                        if (classList[i].ClassID == hashedClasses[n].ClassID)
                        {
                            // if hash is different change color to edited
                            if (hashedClasses[n].Hash != ComputeSha256Hash(classList[i].Serialize()))
                            {
                                for (int j = 0; j < classList[i].ChangedData.Count; j++)
                                {
                                    if (classList[i].ChangedData[j])
                                    {
                                        ws.Row(i + 2).Cell(j + 1).Style.Fill.BackgroundColor = edited;
                                    }
                                }
                                //ws.Row(i + 2).Cell().Style.Fill.BackgroundColor = empty;
                            }
                            break;
                        }
                    }
                    */
                }
                // Iterate over deletedclasses to format the background of each row appropriately
                for (int i = 0; i < deletedClasses.Count; i++)
                {
                    ws.Row(classList.Count + i + 2).Style.Fill.BackgroundColor = edited;
                    ws.Row(classList.Count + i + 2).Style.Font.Strikethrough = true;
                    ws.Row(classList.Count + i + 2).Style.Font.FontColor = XLColor.Red;
                }

                wb.SaveAs(savefileName);
            }
            SerializePairs();
        }
        public void Btn_SaveChanges_Click(object sender, RoutedEventArgs e) // Save changes button handler. Calls SaveChanges() 
        {
            SaveChanges();
            SaveRoomHeaderFile();
            SaveProfHeaderFile();
            latestHashDigest = ComputeSha256Hash(classList.Serialize());
        }

        private void CreateProfessorReport()
        {
            //read in headers file
            string fileName = "Headers.xlsx";
            FileInfo f = new FileInfo(fileName);
            string name = f.FullName;
            var headerFile = new XLWorkbook(name);
            var headerWorksheet = headerFile.Worksheet("Headers");
            int numberHeaders = headerWorksheet.RowsUsed().Count();
            var profInfo = headerFile.Worksheet("ProfessorInfo");
            int numberProfHeaders = profInfo.ColumnsUsed().Count();
            var Rows = profInfo.RangeUsed().RowsUsed().Skip(1);
            int numRows = profInfo.RowsUsed().Count();


            //read room headers from classroom worksheet into list
            var profHeaderList = new List<string>();
            for (int i = 0; i < numberProfHeaders; i++)
            {
                profHeaderList.Add(profInfo.Column(i + 1).Cell(1).GetValue<string>());
            }

            //indexes
            int indexProfName = profHeaderList.IndexOf("Name") + 1;//professor Worksheet
            int indexProfID = profHeaderList.IndexOf("ID") + 1;//professor Worksheet
            int indexMaxCredit = profHeaderList.IndexOf("Max Credit") + 1; //professor Worksheet
            int indexMaxPrep = profHeaderList.IndexOf("Max Prep") + 1; //professor Worksheet

            //professor name list
            var profList = new List<string>();
            for (int i = 1; i < numRows; i++)
            {
                profList.Add(profInfo.Row(i + 1).Cell(indexProfName).GetValue<string>());
            }

            //professor ID list
            var profIDList = new List<string>();
            for (int i = 1; i < numRows; i++)
            {
                profIDList.Add(profInfo.Row(i + 1).Cell(indexProfID).GetValue<string>());
            }



            try
            {
                //Create Word Document
                Microsoft.Office.Interop.Word.Application winword = new Microsoft.Office.Interop.Word.Application();
                winword.ShowAnimation = false;
                winword.Visible = false;
                object missing = System.Reflection.Missing.Value;
                Microsoft.Office.Interop.Word.Document document = winword.Documents.Add(ref missing, ref missing, ref missing, ref missing);



                //Title  

                Microsoft.Office.Interop.Word.Paragraph header = document.Content.Paragraphs.Add(ref missing);
                header.Range.Font.Bold = 1;
                header.Range.Font.Size = 25;
                header.Range.Font.ColorIndex = Microsoft.Office.Interop.Word.WdColorIndex.wdGreen;
                header.Range.Text = "                          Professor Report" + "\n\r";

                foreach (Microsoft.Office.Interop.Word.Section section in document.Sections)
                {

                   // Microsoft.Office.Interop.Word.Range headerRange = section.Headers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                   // headerRange.Fields.Add(headerRange, Microsoft.Office.Interop.Word.WdFieldType.wdFieldPage);
                   // headerRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
                   // headerRange.Font.ColorIndex = Microsoft.Office.Interop.Word.WdColorIndex.wdGreen;
                   // headerRange.Font.Size = 25;
                   // headerRange.Text = "Professor Report";
                }

                //document.Range.Text = profList[1] + " " + profIDList[1] + "\n\r";
                //List of all professors  
                Microsoft.Office.Interop.Word.Paragraph names = document.Content.Paragraphs.Add(ref missing);
                
                
                for (int i = 0; i < profList.Count; i++)
                {
                    names.Range.Font.Bold = 1;
                    names.Range.Font.Size = 15;
                    names.Range.Text = profList[i] + "     ID: " + profIDList[i] + "\r";

                    names.Range.Font.Bold = 0;
                    names.Range.Font.Size = 12;

                    names.Range.Text = "Classes: " + "\r";
                    
                    string sruID = profIDList[i];
                    int credits = 0;
                    //System.Windows.MessageBox.Show("in loop" + classList[i].Prof);
                    //doucument.Content.Environment.NewLine;
                    for (int j = 0; j < classList.Count; j++)
                    {
                        //System.Windows.MessageBox.Show("in for loop" + classList[j].Prof.FirstName);

                        if (classList[j].Prof.SRUID == sruID.ToString())
                        {
                            string classname = classList[j].ClassName.ToString();
                            names.Range.Text = "   Class: " + classname + " " + classList[j].ClassNumber + "\r";
                            names.Range.Text = "   CRN: " + classList[j].CRN + "\r";
                            credits = classList[j].Prof.NumClasses;
                            names.Range.Text = "   Credits: " + classList[j].Credits + "\n\r";

                            
                            //System.Windows.MessageBox.Show("in if loop" + classname);
                        }
                    }

                    names.Range.Text = "Total Credits : " + credits + "\n\r";

                }
                names.Range.InsertParagraphAfter();



                //Save the document  
                
                string fileDir = getFileDirectory(System.Windows.Application.Current.Resources["FilePath"].ToString());
                object filename = fileDir + @"\Professor_Report.docx";
                document.SaveAs2(ref filename);
                document.Close(ref missing, ref missing, ref missing);
                document = null;
                winword.Quit(ref missing, ref missing, ref missing);
                winword = null;
                
                //System.Windows.MessageBox.Show("Document created successfully !");
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show(ex.Message);
            }
        }
        public void Btn_Reports_Click(object sender, RoutedEventArgs e)
        {
            Unfocus_Overlay.Visibility = Visibility.Visible;
            CreateReportsDialog createReportsDialog = new CreateReportsDialog();
            createReportsDialog.Owner = this;
            createReportsDialog.ShowDialog();
            Unfocus_Overlay.Visibility = Visibility.Hidden;


            string newDigest = ComputeSha256Hash(classList.Serialize());
            bool cancelVal = false;

            if ((bool)System.Windows.Application.Current.Resources["Set_Report_Success"])
            {
                if (newDigest != latestHashDigest)
                {
                    string messageBoxText = "Professor report requires saved changes in order to proceed!\nWould you like to Save and Continue?";
                    string caption = "Unsaved changes";
                    MessageBoxButton button = MessageBoxButton.YesNo;
                    MessageBoxImage icon = MessageBoxImage.Question;
                    // Display + Process message box results
                    MessageBoxResult result = System.Windows.MessageBox.Show(messageBoxText, caption, button, icon);
                    switch (result)
                    {
                        case MessageBoxResult.Yes:
                            SaveChanges();
                            SaveErrorHistory();
                            SaveRoomHeaderFile();
                            SaveProfHeaderFile();
                            break;
                        case MessageBoxResult.No:
                            SaveErrorHistory();
                            cancelVal = true;
                            break;

                    }
                }
                if (cancelVal == false)
                {

                    int type = (int)System.Windows.Application.Current.Resources["Report_Type"];
                    if (type == 0)
                    {
                        CreateProfessorReport();
                    }

                    System.Windows.Application.Current.Resources["Set_Report_Success"] = false;


                }
            }



        }



        public void MainWindow_Closing(object sender, CancelEventArgs e) // Window close button handler. Prevents closing if user has unsaved changes 
        {
            string newDigest = ComputeSha256Hash(classList.Serialize());
            if (newDigest != latestHashDigest)
            {
                string messageBoxText = "You have unsaved changes!\nWould you like to Save and Exit?";
                string caption = "Unsaved changes";
                MessageBoxButton button = MessageBoxButton.YesNoCancel;
                MessageBoxImage icon = MessageBoxImage.Question;
                // Display + Process message box results
                MessageBoxResult result = System.Windows.MessageBox.Show(messageBoxText, caption, button, icon);
                switch (result)
                {
                    case MessageBoxResult.Yes:
                        SaveChanges();
                        SaveRoomHeaderFile();
                        SaveProfHeaderFile();
                        SaveErrorHistory();
                        break;
                    case MessageBoxResult.No:
                        SaveErrorHistory();
                        break;
                    case MessageBoxResult.Cancel:
                        e.Cancel = true;
                        break;
                }
            }
        }
        public void btn_MainWindow_Closing(object sender, RoutedEventArgs e) // Exit Program Button
        {
            string newDigest = ComputeSha256Hash(classList.Serialize());
            bool cancelVal = false;
            if (newDigest != latestHashDigest)
            {
                string messageBoxText = "You have unsaved changes!\nWould you like to Save and Exit?";
                string caption = "Unsaved changes";
                MessageBoxButton button = MessageBoxButton.YesNoCancel;
                MessageBoxImage icon = MessageBoxImage.Question;
                // Display + Process message box results
                MessageBoxResult result = System.Windows.MessageBox.Show(messageBoxText, caption, button, icon);
                switch (result)
                {
                    case MessageBoxResult.Yes:
                        SaveChanges();
                        SaveErrorHistory();
                        SaveRoomHeaderFile();
                        SaveProfHeaderFile();
                        break;
                    case MessageBoxResult.No:
                        SaveErrorHistory();
                        break;
                    case MessageBoxResult.Cancel:
                        cancelVal = true;
                        break;
                }
            }
            if (cancelVal == false)
            {

                this.Closing -= MainWindow_Closing;
                this.Close();
            }

        }
        private void CheckBox_Click(object sender, RoutedEventArgs e) // When excludeCredits checkbox is clicked
        {
            RefreshGUI();
        }
        private void Browse_Prof_Prefs_Click(object sender, RoutedEventArgs e) // Locate file where professor preferences is stored
        {
            Microsoft.Win32.OpenFileDialog openFileDialog = new Microsoft.Win32.OpenFileDialog();
            openFileDialog.Filter = "Excel File (*.xlsx)|*.xlsx";
            if (openFileDialog.ShowDialog() == true)
            {
                Professor_Preference_Source.Text = openFileDialog.FileName;
                try
                {
                    using (var excelWorkbook = new XLWorkbook(openFileDialog.FileName))
                    {
                    }
                }
                catch (IOException ex)
                {
                    System.Windows.MessageBox.Show("Excel file is currently open!\n\nPlease close it before proceeding...");
                }
            }
        }
        private void Submit_Prof_Prefs_Click(object sender, RoutedEventArgs e) // Get data from professor preference spreadsheet and create the preference list
        {
            int Pcount = 0;
            if (Professor_Preference_Source.Text == "")
            {
                System.Windows.MessageBox.Show("Please select a professor preference excel file first!");
            }
            else
            {
                //MessageBox.Show("Loading excel file into appropriate data structure...");
                try
                {
                    using (var excelWorkbook = new XLWorkbook(Professor_Preference_Source.Text))
                    {
                        // Select Worksheet
                        var worksheet = excelWorkbook.Worksheet("Headers");
                        var rows = worksheet.RangeUsed().RowsUsed().Skip(1);
                        // Determine number of columns
                        var headerRow = worksheet.Row(1);
                        int columns = 3;
                        string headerString;
                        while (true)
                        {
                            headerString = headerRow.Cell(columns + 1).GetValue<string>();
                            if (headerString == "")
                            {
                                //MessageBox.Show("Ended prof col counts. Total : " + columns);
                                break;
                            }
                            else
                            {
                                columns++;
                            }
                        }
                        // Create one professor preference object for each professor in the preference file
                        string profName;
                        for (int i = 3; i < columns; i++)
                        {
                            profName = headerRow.Cell(i + 1).GetValue<string>(); // Currently last name is the best ID we have
                            professorPreferences.Add(new ProfessorPreference(profName));
                            //MessageBox.Show("Added Preference for Professor: " + profName);
                        }
                        // Create preference object for each class and add it to its respective professor preference list
                        string dept, classString;
                        int classNum;
                        foreach (var row in rows)
                        {
                            if (row.Cell(1).GetValue<string>() != "")
                            {
                                string code;
                                dept = row.Cell(1).GetValue<string>();
                                classString = row.Cell(2).GetValue<string>();
                                if (classString.Contains("/"))
                                {
                                    string[] classes = classString.Split('/');
                                    for (int i = 0; i < classes.Length; i++)
                                    {
                                        classNum = Int32.Parse(classes[i]);
                                        for (int n = 4; n <= columns; n++)
                                        {
                                            code = row.Cell(n).GetValue<string>();
                                            professorPreferences[n - 4].PreferenceList.Add(new Preference(dept, classNum, code));
                                            //MessageBox.Show("Added Preference\nProf: " + professorPreferences[n - 4].ProfessorID + "\nPrefs: " + dept + " " + classNum + " : " + code);
                                        }
                                    }
                                }
                                else
                                {
                                    classNum = row.Cell(2).GetValue<int>();
                                    for (int i = 4; i <= columns; i++)
                                    {
                                        code = row.Cell(i).GetValue<string>();
                                        professorPreferences[i - 4].PreferenceList.Add(new Preference(dept, classNum, code));
                                        //MessageBox.Show("Added Preference\nProf: " + professorPreferences[i - 4].ProfessorID + "\nPrefs: " + dept + " " + classNum + " : " + code);
                                    }
                                }
                            }
                        }
                    }
                    Loaded_URL_Preferences.Text = "Loaded file: " + Professor_Preference_Source.Text;
                }
                catch (IOException ex)
                {
                    System.Windows.MessageBox.Show("Excel file is currently open!\n\nPlease close it before proceeding...");
                    Pcount = 1;
                }
                catch(Exception ex)
                {
                    System.Windows.MessageBox.Show("Unable to understand professor prefrences. Please select a different file.");
                    Pcount = 1;
                }
                if (Pcount == 0)
                {
                    System.Windows.MessageBox.Show("Preferences successfully submitted.");
                    ProcessProfessorPreferences();
                    RefreshGUI();
                }
                
            }
        }
        private void ProcessProfessorPreferences() // Update classes to reflect the preferences of professors, (if any) 
        {
            // Update preference level for each class in classList
            for (int i = 0; i < classList.Count; i++)
            {
                // Try and find a preference for this class+professor combo (professors are identified by last name in preference list)
                string profID = classList[i].Prof.LastName;
                string dept = classList[i].DeptName;
                int num = classList[i].ClassNumber;
                for (int n = 0; n < professorPreferences.Count; n++) // find the prof
                {
                    if (profID == professorPreferences[n].ProfessorID)
                    {
                        for (int x = 0; x < professorPreferences[n].PreferenceList.Count; x++) // find the class
                        {
                            if (professorPreferences[n].PreferenceList[x].Dept == dept && professorPreferences[n].PreferenceList[x].ClassNum == num)
                            {
                                classList[i].PreferenceLevel = professorPreferences[n].PreferenceList[x].Sentiment;
                                //MessageBox.Show("" + classList[i].ClassID + " : " + classList[i].PreferenceLevel);
                                classList[i].PreferenceMessage = professorPreferences[n].PreferenceList[x].Message;
                                //MessageBox.Show("Found preference for " + profID + " in " + dept + " " + num);
                                classList[i].PreferenceCode = professorPreferences[n].PreferenceList[x].Code;
                                if (professorPreferences[n].PreferenceList[x].Message == "Taught before but prefer to teach on-line" && !classList[i].Online)
                                {
                                    classList[i].PreferenceLevel = -1;
                                }
                                else if (professorPreferences[n].PreferenceList[x].Message == "Prefer not to teach this class in the Fall" && termString != "Fall")
                                {
                                    classList[i].PreferenceLevel = 0;
                                }
                                else if (professorPreferences[n].PreferenceList[x].Message == "Prefer not to teach this class in the Spring" && termString != "Spring")
                                {
                                    classList[i].PreferenceLevel = 0;
                                }
                                break;
                            }
                        }
                        break;
                    }
                }
            }
        }
        private void Browse_Class_Groups_Click(object sender, RoutedEventArgs e) // Locate file where professor preferences is stored
        {
            Microsoft.Win32.OpenFileDialog openFileDialog = new Microsoft.Win32.OpenFileDialog();
            openFileDialog.Filter = "Excel File (*.xlsx)|*.xlsx";
            if (openFileDialog.ShowDialog() == true)
            {
                Class_Groups_Source.Text = openFileDialog.FileName;
                try
                {
                    using (var excelWorkbook = new XLWorkbook(openFileDialog.FileName))
                    {
                    }
                }
                catch (IOException ex)
                {
                    System.Windows.MessageBox.Show("Excel file is currently open!\n\nPlease close it before proceeding...");
                }
            }
        }
        private void Submit_Class_Groups_Click(object sender, RoutedEventArgs e) // Get data from professor preference spreadsheet and create the preference list
        {
            int Gcount = 0;
            if (Class_Groups_Source.Text == "")
            {
                System.Windows.MessageBox.Show("Please select a class groupings excel file first!");
            }
            else
            {
                //MessageBox.Show("Loading excel file into appropriate data structure...");
                try
                {
                    using (var excelWorkbook = new XLWorkbook(Class_Groups_Source.Text))
                    {
                        // Select Worksheet
                        var worksheet = excelWorkbook.Worksheet("Headers");
                        var rows = worksheet.RangeUsed().RowsUsed();
                        //MessageBox.Show("" + rows.Count());

                        // Create a class group object for each row in grouping file
                        foreach (var row in rows)
                        {
                            if (row.Cell(1).GetValue<string>() != "")
                            {
                                ClassGroup group = new ClassGroup();
                                // Add the classes to the newly created group
                                bool rowEnd = false;
                                int i = 1;
                                while (!rowEnd)
                                {
                                    if (row.Cell(i).GetValue<string>() != "")
                                    {
                                        string classInfo = row.Cell(i).GetValue<string>();
                                        // Separte dept and classnum
                                        string dept = classInfo.Split(' ')[0].ToUpper();
                                        string num = classInfo.Split(' ')[1];
                                        // insert data into classgroup
                                        group.AddEntry(dept, num);
                                        i++;
                                    }
                                    else
                                    {
                                        rowEnd = true;
                                    }
                                }
                                classGroupings.Add(group);
                            }
                        }
                    }
                    Loaded_URL_Groups.Text = "Loaded file: " + Class_Groups_Source.Text;
                }
                catch (IOException ex)
                {
                    System.Windows.MessageBox.Show("Excel file is currently open!\n\nPlease close it before proceeding...");
                    Gcount = 1;
                }
                catch (Exception ex)
                {
                    System.Windows.MessageBox.Show("Unable to understand class groupings. Please select a different file.");
                    Gcount = 1;
                }
                if (Gcount == 0)
                {
                    System.Windows.MessageBox.Show("Class groupings successfully submitted.");
                    RefreshGUI();
                }
                
            }
        }
        private void ProcessClassGroupings() // Update classes to reflect the preferences of professors, (if any) 
        {
            string oldText = soft_constraint_log.Text;
            soft_constraint_log.Text = "";
            // Go through classlist and see if any class is scheduled at the same time as another in the same group
            for (int i = 0; i < classList.Count; i++)
            {
                for (int n = 0; n < classGroupings.Count; n++)
                {
                    if (isGroupMember(classList[i], n))
                    {
                        // class found inside this particular group -> check if scheduled at the same day / time as others
                        checkGroupConflict(classList[i], n);
                    }
                }
            }
            if (soft_constraint_log.Text == "")
            {
                soft_constraint_log.Text = "> None";
            }
            else
            {
                if (oldText != soft_constraint_log.Text)
                {
                    System.Windows.MessageBox.Show("Class group conflicts detected.\nPlease refer to the Conflicts tab for details.");
                }
            }
        }
        private void SubmitChangeTerm_Click(object sender, RoutedEventArgs e) // update term and year from user input
        {
            if (TermYearBox.Text.Length == 4)
            {
                termString = TermComboBox.Text;
                switch (termString)
                {
                    case "Spring":
                        term = "01";
                        break;
                    case "Summer":
                        term = "06";
                        break;
                    case "Fall":
                        term = "09";
                        break;
                    case "Winter":
                        term = "12";
                        break;
                    default:
                        System.Windows.MessageBox.Show("Unexpected term name!");
                        term = "00";
                        termString = "None";
                        break;
                }
                termYear = TermYearBox.Text;
                ProcessProfessorPreferences();
                System.Windows.MessageBox.Show("Changed Term to: " + termString + " " + termYear);
                RefreshGUI();
            }
            else
            {
                System.Windows.MessageBox.Show("Please enter a valid year (e.g. 2020)");
            }
        }

        // ADD / REMOVE / EDIT functionality (Professors, Classrooms, Classes)
        // Professors
        public void AddProfessor(Professors prof)
        {
            professors.Add(prof);
            colorPairs.ColorPairings.Add(new ProfColors { ProfName = prof.FullName, Color = prof.profRGB.colorString });
        }
        private void Btn_AddProfessor_Click(object sender, RoutedEventArgs e)
        {
            //MessageBox.Show(Application.Current.Resources["setProf"].ToString());

            Unfocus_Overlay.Visibility = Visibility.Visible;
            AddProfessorDialog addProfDialog = new AddProfessorDialog();
            addProfDialog.Owner = this;
            addProfDialog.ShowDialog();
            Unfocus_Overlay.Visibility = Visibility.Hidden;

            if ((bool)System.Windows.Application.Current.Resources["Set_Prof_Success"])
            {
                string fName = (string)System.Windows.Application.Current.Resources["Set_Prof_FN"];
                string lName = (string)System.Windows.Application.Current.Resources["Set_Prof_LN"];
                string id = (string)System.Windows.Application.Current.Resources["Set_Prof_ID"];
                string colorString = (string)System.Windows.Application.Current.Resources["Set_Prof_Color"];
                Professors tmpProf = new Professors(fName, lName, id);
                tmpProf.profRGB = new RGB_Color(colorString);
                AddProfessor(tmpProf);
                System.Windows.Application.Current.Resources["Set_Prof_Success"] = false;
            }
        }
        public void RemoveProfessor(string sruID)
        {
            for (int i = 0; i < professors.Count; i++)
            {
                if (professors[i].SRUID == sruID)
                {
                    professors.RemoveAt(i);
                    break;
                }
            }
            for (int i = 0; i < classList.Count; i++)
            {
                if (classList[i].Prof.SRUID == sruID)
                {
                    classList[i].Prof = new Professors();
                }
            }
            for (int i = 0; i < unassignedClasses.Count; i++)
            {
                if (unassignedClasses[i].Prof.SRUID == sruID)
                {
                    unassignedClasses[i].Prof = new Professors();
                }
            }
            for (int i = 0; i < onlineClasses.Count; i++)
            {
                if (onlineClasses[i].Prof.SRUID == sruID)
                {
                    onlineClasses[i].Prof = new Professors();
                }
            }
        }


        private void Btn_RemoveProfessor_Click(object sender, RoutedEventArgs e)
        {
            // find the professor
            string sruID = "";
            System.Windows.Controls.MenuItem mi = sender as System.Windows.Controls.MenuItem;
            if (mi != null)
            {
                System.Windows.Controls.ContextMenu cm = mi.CommandParameter as System.Windows.Controls.ContextMenu;
                if (cm != null)
                {
                    System.Windows.Controls.ListViewItem source = cm.PlacementTarget as System.Windows.Controls.ListViewItem;
                    if (source != null) // Being called from a Professor Color Key
                    {
                        sruID = source.Tag.ToString();
                    }
                    else // Being called from a GridRow
                    {
                        DataGridRow sourceRow = cm.PlacementTarget as DataGridRow;
                        System.Windows.Controls.DataGrid parentGrid = GetParent<System.Windows.Controls.DataGrid>(sourceRow as DependencyObject);
                        TextBlock prof_ID = parentGrid.Columns[2].GetCellContent(sourceRow) as TextBlock;
                        sruID = prof_ID.Text;
                    }
                    RemoveProfessor(sruID);
                    RefreshGUI();
                }
            }
        }
        public void EditProfessor(string sruID)
        {
            Unfocus_Overlay.Visibility = Visibility.Visible;
            Professors prof = DetermineProfessor(sruID);
            EditProfessorDialog editProfessorDialog = new EditProfessorDialog(prof);
            editProfessorDialog.Owner = this;
            editProfessorDialog.ShowDialog();
            // Edit ColorPairs entry
            for (int i = 0; i < colorPairs.ColorPairings.Count; i++)
            {
                if (colorPairs.ColorPairings[i].ProfName == prof.FullName)
                {
                    colorPairs.ColorPairings[i].Color = prof.profRGB.colorString;
                }
            }
            Unfocus_Overlay.Visibility = Visibility.Hidden;
        }
        private void Btn_EditProfessor_Click(object sender, RoutedEventArgs e)
        {
            // find the professor
            string sruID = "";
            System.Windows.Controls.MenuItem mi = sender as System.Windows.Controls.MenuItem;
            if (mi != null)
            {
                System.Windows.Controls.ContextMenu cm = mi.CommandParameter as System.Windows.Controls.ContextMenu;
                if (cm != null)
                {
                    System.Windows.Controls.ListViewItem source = cm.PlacementTarget as System.Windows.Controls.ListViewItem;
                    if (source != null) // Being called from a Professor Color Key item
                    {
                        sruID = source.Tag.ToString();
                    }
                    else // Being called from a GridRow
                    {
                        DataGridRow sourceRow = cm.PlacementTarget as DataGridRow;
                        System.Windows.Controls.DataGrid parentGrid = GetParent<System.Windows.Controls.DataGrid>(sourceRow as DependencyObject);
                        TextBlock prof_ID = parentGrid.Columns[2].GetCellContent(sourceRow) as TextBlock;
                        sruID = prof_ID.Text;
                    }
                    EditProfessor(sruID);
                    RefreshGUI();
                }
            }
        }
        public new object DataContext { get; set; }
        private void ViewRoomInfo_Click(object sender, RoutedEventArgs e) // When classroom button is clicked
        {
            string bldgID, building, tempRoomLabel, roomLabel, notes;
            int roomID, room, capacity;
            string start = "";
            string end = "";
            string endTimePure = "";
            string meridian = "";
            int row = -1;
            int column = -1;
            int timeTable = -1;
            int timeChange = 1;
            int rowAdd = 0;
            System.Windows.Controls.Button btn = sender as System.Windows.Controls.Button;
            roomLabel = (String)btn.Content;
            ClassRoomList classrooms = (ClassRoomList)System.Windows.Application.Current.FindResource("ClassRoom_List_View");


            for (int n = 0; n < classrooms.Count; n++)
            {
                bldgID = classrooms[n].Location;
                roomID = classrooms[n].RoomNum;
                tempRoomLabel = (bldgID + "-" + roomID);

                if (tempRoomLabel == roomLabel)
                {
                    building = classrooms[n].Location;
                    room = classrooms[n].RoomNum;
                    capacity = classrooms[n].AvailableSeats;
                    notes = classrooms[n].Notes;

                    //Open Window with classroom information
                    Unfocus_Overlay.Visibility = Visibility.Visible;
                    EditClassRoomInfo roomInfo = new EditClassRoomInfo(building, room, capacity, notes);
                    roomInfo.CurrentBuilding = building;
                    this.DataContext = roomInfo;
                    this.BindData();
                    roomInfo.ShowDialog();
                    Unfocus_Overlay.Visibility = Visibility.Hidden;

                    if(roomInfo.ChangeClasses() == true)
                    { 
                        if (roomInfo.NewBuilding == null)
                        {
                            break;
                        }
                        if (roomInfo.NewRoom.Equals(""))
                        {
                            break;
                        }
                        if (roomInfo.NewCapacity.Equals(0))
                        {
                            break;
                        }

                        var newBuilding = roomInfo.NewBuilding.ToUpper();
                        var newRoom = roomInfo.NewRoom;
                        var newCapacity = roomInfo.NewCapacity;
                        var newNotes = roomInfo.NewNotes;

                        ChangeClassLocation(newBuilding, newRoom, newCapacity, newNotes, building, room, capacity, notes);

                        //Make sure none of the rooms are duplicates. If so, delete both and re-add one
                        int count = 0;
                        for (int i = 0; i < classrooms.Count; i++)
                        {
                            if (classrooms[i].ClassID == newBuilding+newRoom)
                            {
                                count++;
                                if (count >= 2)
                                {
                                    if ((newBuilding != "WEB") && (newBuilding != "APPT") && (newBuilding != "APPT2"))
                                    {
                                        errorMSG = "Classroom " + classrooms[i].ClassID + " already exists. Classes have been moved to Unassigned tab.";
                                        WriteErrorLog(errorMSG);
                                        RemoveClassroom(newBuilding, newRoom, newCapacity, newNotes);
                                        break;
                                    }
                                }

                            }
                        }


                        // Remove old Grids
                        Grid child = FindName("MWF_") as Grid;
                        MWF.Children.Remove(child);
                        Grid child2 = FindName("TR_") as Grid;
                        TR.Children.Remove(child2);
                        // Redraw Grids
                        DrawTimeTablesDynamic(start, end, meridian, row, column, timeTable, timeChange, endTimePure, rowAdd);

                        break;
                    }
                }
            }
        }

        /// <summary>
        /// The purpose is to determine if the information in the master scheduler file matches what is found in the default classroom file.
        /// Users will be made aware of any errors via the error log
        /// </summary>

        private void CheckRoomInfo(int columns, string masterFileName) // Check the classroom headers worksheet against the master scheduling worksheet
        {
            //read in headers file
            string fileName = "Headers.xlsx";
            FileInfo f = new FileInfo(fileName);
            string name = f.FullName;
            var headerFile = new XLWorkbook(name);
            var headerWorksheet = headerFile.Worksheet("Headers");
            int numberHeaders = headerWorksheet.RowsUsed().Count();
            var classroomInfo = headerFile.Worksheet("ClassroomInfo");
            int numberRoomHeaders = classroomInfo.RowsUsed().Count();
            var Rows = classroomInfo.RangeUsed().RowsUsed().Skip(1);

            Excel.Application oExcel = new Microsoft.Office.Interop.Excel.Application();
            Excel.Workbook excelWorkbook = oExcel.Workbooks.Open(masterFileName);
            Excel.Worksheet worksheet = (Excel.Worksheet)excelWorkbook.Worksheets[1];
            Excel.Range range = worksheet.UsedRange;
            var masterRows = range.Rows.Count;


            //read headers from headers file into list
            var headerList = new List<string>();
            for (int i = 0; i < numberHeaders; i++)
            {
                headerList.Add(headerWorksheet.Row(i + 1).Cell(1).GetValue<string>());
            }

            //read room headers from classroom worksheet into list
            var roomHeaderList = new List<string>();
            for (int i = 0; i < numberRoomHeaders; i++)
            {
                roomHeaderList.Add(classroomInfo.Column(i + 1).Cell(1).GetValue<string>());
            }

            //identify where the necessary headers are in both files
            int indexMAXS = headerList.IndexOf("MAX SEATS") + 1;
            int indexProjSeats = headerList.IndexOf("ProjSeats") + 1;
            int indexRCAP = headerList.IndexOf("ROOM_CAP") + 1;
            int indexBLDG = headerList.IndexOf("BLDG") + 1;
            int indexROOM = headerList.IndexOf("ROOM") + 1;

            int indexDefaultSeats = roomHeaderList.IndexOf("ROOM_CAP") + 1;
            int indexDefaultBLDG = roomHeaderList.IndexOf("BLDG") + 1;
            int indexDefaultNUM = roomHeaderList.IndexOf("ROOM") + 1;

            int defaultRoom = 0, parseResult = 0;
            double room = 0.0, MAX_SEATS = 0.0, PROJ_SEATS = 0.0, ROOM_CAP = 0.0, defaultSeats = 0, outputRoom;
            string bldg = "", defaultBldg = "";

            for (int i = 2; i < (masterRows + 1); i++)
            {
                //capture the attributes of the indexed classroom
                bldg = (worksheet.Cells[i, indexBLDG] as Excel.Range).Value;

                if ((bldg == null) || (bldg == ""))
                {
                    errorMSG = "In row " + i + ", there is a missing value for the BLDG column. Please fix this error before next startup.";
                    WriteErrorLog(errorMSG);
                    break;
                }
                if (bldg != "WEB" && !bldg.Contains("APPT"))
                {
                    var strRoom = (worksheet.Cells[i, indexROOM] as Excel.Range).Value.ToString();
                    bool isDNumeric = double.TryParse(strRoom, out outputRoom);

                    if (isDNumeric)
                    {
                        room = Double.Parse(strRoom);
                    }

                    

                    if (!((worksheet.Cells[i, indexMAXS] as Excel.Range).Equals(null))) //make sure MAX_SEATS is not null
                    {
                        var strRoomMax = (worksheet.Cells[i, indexMAXS] as Excel.Range).Value.ToString();
                        bool isMNumeric = double.TryParse(strRoomMax, out outputRoom);
                        if (isMNumeric)
                        {
                            MAX_SEATS = Double.Parse(strRoomMax);
                        }
                    }

                    

                    if (!((worksheet.Cells[i, indexProjSeats] as Excel.Range).Equals(null))) //make sure PROJ_SEATS is not null
                    {
                        var strRoomProj = (worksheet.Cells[i, indexProjSeats] as Excel.Range).Value.ToString();
                        bool isPNumeric = double.TryParse(strRoomProj, out outputRoom);
                        if (isPNumeric)
                        {
                            PROJ_SEATS = Double.Parse(strRoomProj);
                        }
                    }

                    

                    if (!((worksheet.Cells[i, indexRCAP] as Excel.Range).Equals(null))) //make sure ROOM_CAP is not null
                    {
                        var strRoomCap = (worksheet.Cells[i, indexRCAP] as Excel.Range).Value.ToString();
                        bool isRNumeric = double.TryParse(strRoomCap, out outputRoom);
                        if (isRNumeric)
                        {
                            ROOM_CAP = Double.Parse(strRoomCap);
                        }
                    }
                }

                foreach (var rowD in Rows)
                {
                    //capture the default attributes of the indexed classroom
                    if (!rowD.Cell(indexDefaultBLDG).IsEmpty())
                    {
                        defaultBldg = rowD.Cell(indexDefaultBLDG).GetValue<string>().ToUpper(); //assign the building from the default sheet

                        if (!rowD.Cell(indexDefaultNUM).IsEmpty() && int.TryParse(rowD.Cell(indexDefaultNUM).GetValue<string>(), out parseResult))
                        {
                            defaultRoom = parseResult; //assign the room from the default sheet

                            if (!rowD.Cell(indexDefaultSeats).IsEmpty() && int.TryParse(rowD.Cell(indexDefaultSeats).GetValue<string>(), out parseResult))
                            {
                                defaultSeats = parseResult;
                            }

                        }
                    }
                    //compare the default seats with master scheduler attributes
                    if ((defaultBldg == bldg) && (defaultRoom == room))
                    {
                        if (!(defaultSeats >= MAX_SEATS))
                        {
                            //print to error log
                            errorMSG = "The value found for MAX_SEATS in row " + i + " of the schedule is larger than the ROOM_CAP for its assigned room. Please move the class to a new classroom or reduce the max seats.";
                            WriteErrorLog(errorMSG);
                        }
                        if (!(defaultSeats >= PROJ_SEATS))
                        {
                            //print to error log
                            errorMSG = "The value found for PROJ_SEATS in row " + i + " of the schedule is larger than the ROOM_CAP for its assigned room. Please move the class to a new classroom or reduce the projected seats.";
                            WriteErrorLog(errorMSG);
                        }
                        if (defaultSeats != ROOM_CAP)
                        {
                            //print to error log
                            errorMSG = "The value found for ROOM_CAP in row " + i + " of the schedule does not match the value in the Headers file regarding the following classroom: " + bldg + "-" + room;
                            WriteErrorLog(errorMSG);
                        }
                        if (!(ROOM_CAP >= PROJ_SEATS))
                        {
                            //print to error log
                            errorMSG = "The value found for PROJ_SEATS in row " + i + " of the schedule is higher than the ROOM_CAP for its assigned room. Please move the class to a new classroom or reduce the projected seats.";
                            WriteErrorLog(errorMSG);
                        }
                        if (!(ROOM_CAP >= MAX_SEATS))
                        {
                            //print to error log
                            errorMSG = "The value found for MAX_SEATS in row " + i + " of the schedule is higher than the ROOM_CAP for its assigned room. Please move the class to a new classroom or reduce the max seats.";
                            WriteErrorLog(errorMSG);
                        }
                    }
                }
            }

            oExcel.Workbooks.Close();
        }
        private void CheckProfInfo(int columns, string masterFileName) // Check the classroom headers worksheet against the master scheduling worksheet
        {
            //read in headers file
            string fileName = "Headers.xlsx";
            FileInfo f = new FileInfo(fileName);
            string name = f.FullName;
            var headerFile = new XLWorkbook(name);
            var headerWorksheet = headerFile.Worksheet("Headers");
            int numberHeaders = headerWorksheet.RowsUsed().Count();
            var profInfo = headerFile.Worksheet("ProfessorInfo");
            int numberProfHeaders = profInfo.ColumnsUsed().Count();
            var Rows = profInfo.RangeUsed().RowsUsed().Skip(1);
            int numRows = profInfo.RowsUsed().Count();

            Excel.Application oExcel = new Microsoft.Office.Interop.Excel.Application();
            Excel.Workbook excelWorkbook = oExcel.Workbooks.Open(masterFileName);
            Excel.Worksheet worksheet = (Excel.Worksheet)excelWorkbook.Worksheets[1];
            var excelprofessorInfo = excelWorkbook.Worksheets[1];
            Excel.Range range = worksheet.UsedRange;
            int masterRows = range.Rows.Count;
            


            //read headers from headers file into list
            var headerList = new List<string>();
            for (int i = 0; i < numberHeaders; i++)
            {
                headerList.Add(headerWorksheet.Row(i + 1).Cell(1).GetValue<string>());
            }

            //read room headers from classroom worksheet into list
            var profHeaderList = new List<string>();
            for (int i = 0; i < numberProfHeaders; i++)
            {
                profHeaderList.Add(profInfo.Column(i + 1).Cell(1).GetValue<string>());
            }

            int indexFaculty = headerList.IndexOf("Faculty Name") + 1;//22
            int indexFacultyID = headerList.IndexOf("Faculty ID") + 1;//23

            int indexProfName = profHeaderList.IndexOf("Name") + 1;//professor Worksheet
            int indexProfID = profHeaderList.IndexOf("ID") + 1;//professor Worksheet
            int indexMaxCredit = profHeaderList.IndexOf("Max Credit") + 1; //professor Worksheet
            int indexMaxPrep = profHeaderList.IndexOf("Max Prep") + 1; //professor Worksheet

            //known max credit to a list
            var maxCreditList = new List<string>();
            for (int i = 1; i < numRows; i++)
            {
                maxCreditList.Add(profInfo.Row(i + 1).Cell(indexMaxCredit).GetValue<string>());
            }

            //known professor names to a list
            var profList = new List<string>();
            for (int i = 1; i < numRows; i++)
            {
                profList.Add(profInfo.Row(i + 1).Cell(indexProfName).GetValue<string>());
            }

            //new professor names to a list
            var newProfList = new List<string>();
            for (int i = 1; i < masterRows; i++)
            {
                newProfList.Add((worksheet.Cells[i + 1, indexFaculty] as Excel.Range).Value);
            }

            //known professor IDs to a list
            var profIDList = new List<string>();
            for (int i = 1; i < numRows; i++)
            {
                profIDList.Add(profInfo.Row(i + 1).Cell(indexProfID).GetValue<string>());
            }

            //new professor IDs to a list
            var newProfIDList = new List<string>();
            for (int i = 1; i < masterRows; i++)
            {
                newProfIDList.Add((worksheet.Cells[i + 1, indexFacultyID] as Excel.Range).Value);
            }

            //compare known and new prof lists
            for (int i = 0; i < newProfList.Count; i++)
            {
                for (int j = 0; j < profList.Count; j++)
                {
                    if (newProfList.ElementAt(i) == profList.ElementAt(j))
                    {

                        
                        if(newProfIDList.ElementAt(i) == profIDList.ElementAt(j))
                        {
                            break;
                        }
                        else if(newProfIDList.ElementAt(i) == null || newProfIDList.ElementAt(i) == "")
                        {
                            //print to error log
                            errorMSG = "The faculty ID number for professor " + newProfList.ElementAt(i) + " is empty.";
                            WriteErrorLog(errorMSG);
                        }
                        else
                        {
                        
                            errorMSG = "The faculty ID number for professor " + newProfList.ElementAt(i) + " does not match the ID number listed in their profile." ;
                            WriteErrorLog(errorMSG);
                        }

                        break;
                    }
                    else if (j == profList.Count - 1 && newProfList.ElementAt(i) != profList.ElementAt(j))
                    {
                        
                        //professor not in list
                    }

                }

            }
            //check max credit and max prep
            int credits = 0;
            for (int i = 0; i < profList.Count; i++)
            {
                for (int j = 0; j < classList.Count; j++)
                {
                    //System.Windows.MessageBox.Show("in for loop" + classList[j].Prof.FirstName);
                    credits = 0;
                    if (classList[j].Prof.SRUID == profIDList.ElementAt(i))
                    {



                        credits = credits + classList[j].Credits;
                        //System.Windows.MessageBox.Show("in if loop" + classname);
                    }
                }
                 
                if (credits > Double.Parse(maxCreditList.ElementAt(i)))
                {
                    errorMSG = profList.ElementAt(i)  + " is assigned too many credits. ";
                    WriteErrorLog(errorMSG);
                }
            }

            oExcel.Workbooks.Close();
        }


        private void ChangeClassLocation(string newBldg, int newRoom, int newSeats, string newNotes, string oldBldg, int oldRoom, int oldSeats, string notes)
        {

            for (int i = 0; i < classList.Count; i++)
            {
                if ((classList[i].Classroom.Location == oldBldg) && (classList[i].Classroom.RoomNum == oldRoom))
                {
                    if (newBldg.Length == 3 && newRoom != 0)
                    {
                        classList[i].Classroom.Location = newBldg;
                        classList[i].Classroom.RoomNum = newRoom;
                        classList[i].Classroom.AvailableSeats = newSeats;
                        classList[i].Classroom.Notes = newNotes;
                        //RemoveClassroom(oldRoomID);
                    }
                }
            }


            for (int i = 0; i < times_Default.Count(); i++)
            {
                if (times_Default_Room[i] == oldRoom)
                {
                    times_Default_Room[i] = newRoom;
                }

            }
            for (int i = 0; i < masterTimeslotList.Count; i++)
            {
                if (masterClassRoomList[i] == oldRoom)
                {
                    masterClassRoomList[i] = newRoom;
                }
            }
            for (int i = 0; i < changedRoomNum.Count(); i++)
            {
                if (changedRoomNum[i] == oldRoom)
                {
                    changedRoomNum[i] = newRoom;
                }
            }
            for (int i = 0; i < autoChangedRoomNum.Count(); i++)
            {
                if (autoChangedRoomNum[i] == oldRoom)
                {
                    autoChangedRoomNum[i] = newRoom;
                }
            }
            return;
        }

        private void ChangeTime_Click(object sender, RoutedEventArgs e)
        {
            System.Windows.Application.Current.Resources["Set_ChangeTime_Success"] = false;
            int rows = Grid.GetRow((System.Windows.Controls.Button)sender);
            int cols = Grid.GetColumn((System.Windows.Controls.Button)sender);
            System.Windows.Application.Current.Resources["Set_thisTime"] = (sender as System.Windows.Controls.Button).Content;

            Unfocus_Overlay.Visibility = Visibility.Visible;
            ChangeTimeDialog changeTimeDialog = new ChangeTimeDialog();
            changeTimeDialog.Owner = this;
            changeTimeDialog.ShowDialog();
            Unfocus_Overlay.Visibility = Visibility.Hidden;

            if (System.Windows.Application.Current.Resources["Set_ChangeTime_Success"] != null && (bool)System.Windows.Application.Current.Resources["Set_ChangeTime_Success"] == true)
            {
                string start = System.Windows.Application.Current.Resources["Set_Start_Time"].ToString();
                string end = System.Windows.Application.Current.Resources["Set_End_Time"].ToString();
                string endTimePure = System.Windows.Application.Current.Resources["Set_End_Time_Pure"].ToString();
                string meridian = System.Windows.Application.Current.Resources["Set_Meridian"].ToString();
                int timeToChange = Int32.Parse(System.Windows.Application.Current.Resources["Set_TimeChange"].ToString());
                int row = rows;
                int rowAdd = 0;
                int column = cols;
                int timeTable = Int32.Parse(System.Windows.Application.Current.Resources["Set_TimeTable"].ToString());
                ChangeTimes(start, end, meridian, row, column, timeTable, timeToChange, endTimePure, rowAdd);
            }
        }
        public void ChangeTimes(string start, string end, string meridian, int row, int column, int timeTable, int timeChange, string endTimePure, int rowAdd)
        {
            // Remove old Grids
            Grid child = FindName("MWF_") as Grid;
            MWF.Children.Remove(child);
            Grid child2 = FindName("TR_") as Grid;
            TR.Children.Remove(child2);
            // Redraw Grids
            DrawTimeTablesDynamic(start, end, meridian, row, column, timeTable, timeChange, endTimePure, rowAdd);
        }
        // Classrooms
        public void AddClassroom(ClassRoom room)
        {
            string start = "";
            string end = "";
            string endTimePure = "";
            string meridian = "";
            int row = -1;
            int column = -1;
            int timeTable = -1;
            int timeChange = 1;
            int rowAdd = 0;
            // Add Classroom to classroom list
            classrooms.Add(room);
            // Remove old Grids
            Grid child = FindName("MWF_") as Grid;
            MWF.Children.Remove(child);
            Grid child2 = FindName("TR_") as Grid;
            TR.Children.Remove(child2);
            // Redraw Grids
            DrawTimeTablesDynamic(start, end, meridian, row, column, timeTable, timeChange, endTimePure, rowAdd);
        }
        private void Btn_AddClassRoom_Click(object sender, RoutedEventArgs e)
        {
            Unfocus_Overlay.Visibility = Visibility.Visible;
            AddClassRoomDialog addClassRoomDialog = new AddClassRoomDialog();
            addClassRoomDialog.Owner = this;
            addClassRoomDialog.ShowDialog();
            Unfocus_Overlay.Visibility = Visibility.Hidden;
            if (System.Windows.Application.Current.Resources["Set_ClassRoom_Success"] != null && (bool)System.Windows.Application.Current.Resources["Set_ClassRoom_Success"] == true)
            {
                string bldg = System.Windows.Application.Current.Resources["Set_ClassRoom_Bldg"].ToString();
                int roomNum = Int32.Parse(System.Windows.Application.Current.Resources["Set_ClassRoom_Num"].ToString());
                int capacity = Int32.Parse(System.Windows.Application.Current.Resources["Set_ClassRoom_Seats"].ToString());
                string notes = System.Windows.Application.Current.Resources["Set_ClassRoom_Notes"].ToString();
                AddClassroom(new ClassRoom(bldg, roomNum, capacity, notes));
                System.Windows.Application.Current.Resources["Set_ClassRoom_Success"] = false;
            }
        }
        public void RemoveClassroom(string bldg, int roomNum, int capacity, string notes)
        {
            //ClassRoomList classrooms = (ClassRoomList)System.Windows.Application.Current.FindResource("ClassRoom_List_View");
            for (int i = 0; i < classrooms.Count; i++)
            {
                if (classrooms[i].ClassID == bldg+roomNum)
                {
                    classrooms.RemoveAt(i);
                }
            }
            //move the classes that used to be there
            for (int i = 0; i < classList.Count; i++)
            {
                if ((classList[i].Classroom.Location == bldg) && (classList[i].Classroom.RoomNum == roomNum))
                {
                    if (bldg.Length == 3 && roomNum != 0)
                    {
                        unassignedClasses.Add(classList[i]);
                    }
                }
            }

        }
        private void Btn_RemoveClassroom_Click(object sender, RoutedEventArgs e)
        {
            System.Windows.MessageBox.Show("Yet to be implemented");
        }
        // Classes
        public void AddClass(Classes newClass)
        {
            classList.Add(newClass);
            if (newClass.Online)
            {
                onlineClasses.Add(newClass);
            }
            else
            {
                if (newClass.isAppointment)
                {
                    if (newClass.Classroom.Location == "APPT")
                    {
                        appointmentClasses.Add(newClass);
                    }
                    else if (newClass.Classroom.Location == "APPT2")
                    {
                        appointment2Classes.Add(newClass);
                    }
                }
                else
                {
                    unassignedClasses.Add(newClass);
                }
            }
        }
        private void Btn_AddClass_Click(object sender, RoutedEventArgs e)
        {
            Unfocus_Overlay.Visibility = Visibility.Visible;
            AddClassDialog addClassDialog = new AddClassDialog();
            addClassDialog.Owner = this;
            addClassDialog.ShowDialog();
            Unfocus_Overlay.Visibility = Visibility.Hidden;
            //MessageBox.Show("class_success: " + Application.Current.MainWindow.Resources["Set_Class_Success"].ToString());

            if (System.Windows.Application.Current.Resources["Set_Class_Success"] != null && (bool)System.Windows.Application.Current.Resources["Set_Class_Success"] == true)
            {
                int term = Int32.Parse(System.Windows.Application.Current.Resources["Set_Class_Term"].ToString());
                int session = Int32.Parse(System.Windows.Application.Current.Resources["Set_Class_Session"].ToString());
                string crn = System.Windows.Application.Current.Resources["Set_Class_CRN"].ToString();
                string dpt = System.Windows.Application.Current.Resources["Set_Class_Dept"].ToString();
                int number = Int32.Parse(System.Windows.Application.Current.Resources["Set_Class_Number"].ToString());
                int sect = Int32.Parse(System.Windows.Application.Current.Resources["Set_Class_Section"].ToString());
                string name = System.Windows.Application.Current.Resources["Set_Class_Name"].ToString();
                int credits = Int32.Parse(System.Windows.Application.Current.Resources["Set_Class_Credits"].ToString());
                string prof = System.Windows.Application.Current.Resources["Set_Class_Professor"].ToString();
                bool online = Boolean.Parse(System.Windows.Application.Current.Resources["Set_Class_Online"].ToString());
                bool appt = Boolean.Parse(System.Windows.Application.Current.Resources["Set_Class_Appointment"].ToString());
                bool appt2 = Boolean.Parse(System.Windows.Application.Current.Resources["Set_Class_Appointment2"].ToString());
                bool appointment = false;
                int projseats = Int32.Parse(System.Windows.Application.Current.Resources["Set_Class_projSeats"].ToString());
                int maxseats = Int32.Parse(System.Windows.Application.Current.Resources["Set_Class_maxSeats"].ToString());
                int enrolled = Int32.Parse(System.Windows.Application.Current.Resources["Set_Class_Enrolled"].ToString());
                int waitlist = Int32.Parse(System.Windows.Application.Current.Resources["Set_Class_Waitlist"].ToString());
                string crosslist = System.Windows.Application.Current.Resources["Set_Class_Crosslist"].ToString();
                string startDate = System.Windows.Application.Current.Resources["Set_Class_StartDate"].ToString();
                string endDate = System.Windows.Application.Current.Resources["Set_Class_EndDate"].ToString();
                //string building = System.Windows.Application.Current.Resources["Set_Class_Building"].ToString();
                //string room = System.Windows.Application.Current.Resources["Set_Class_Room"].ToString();
                //string roomCap = System.Windows.Application.Current.Resources["Set_Class_RoomCap"].ToString();
                string building = "";
                string room = "";
                string roomCap = "0";



                ClassRoom CRoom = null;
                if (appt || appt2)
                {
                    appointment = true;
                    if (appt)
                    {
                        CRoom = new ClassRoom("APPT", 0);
                    }
                    else
                    {
                        CRoom = new ClassRoom("APPT2", 0);
                    }
                }
                else if (online)
                {
                    CRoom = new ClassRoom("WEB", 999);
                }
                else
                {
                    CRoom = new ClassRoom();
                }
                AddClass(new Classes(term, session, crn, dpt, number, sect, name, credits, "", new Timeslot(), 0, CRoom, DetermineProfessor(prof), online, appointment, false, "", "", new List<string>(), maxseats, projseats, enrolled, waitlist, crosslist, startDate, endDate, building, room, roomCap));
                
                System.Windows.Application.Current.Resources["Set_Class_Success"] = false;
                RefreshGUI();
            }
        }
        public void RemoveClass(string ID)
        {
            Classes removalTarget;
            for (int i = 0; i < classList.Count; i++)
            {
                if (classList[i].ClassID == ID)
                {
                    removalTarget = classList[i];
                    deletedClasses.Add(removalTarget);
                    classList.RemoveAt(i);
                    break;
                }
            }
        }
        private void Btn_RemoveClass_Click(object sender, RoutedEventArgs e)
        {
            // find the class
            string classID = "";
            System.Windows.Controls.MenuItem mi = sender as System.Windows.Controls.MenuItem;
            if (mi != null)
            {
                System.Windows.Controls.ContextMenu cm = mi.CommandParameter as System.Windows.Controls.ContextMenu;
                if (cm != null)
                {
                    System.Windows.Controls.Label source = cm.PlacementTarget as System.Windows.Controls.Label;
                    if (source != null) // Being called from a Label
                    {
                        classID = source.Tag.ToString();
                    }
                    else // Being called from a GridRow
                    {
                        DataGridRow sourceRow = cm.PlacementTarget as DataGridRow;
                        System.Windows.Controls.DataGrid parentGrid = GetParent<System.Windows.Controls.DataGrid>(sourceRow as DependencyObject);
                        TextBlock classCRN = parentGrid.Columns[0].GetCellContent(sourceRow) as TextBlock;
                        TextBlock className = parentGrid.Columns[4].GetCellContent(sourceRow) as TextBlock;
                        TextBlock classSection = parentGrid.Columns[3].GetCellContent(sourceRow) as TextBlock;
                        TextBlock classNumber = parentGrid.Columns[2].GetCellContent(sourceRow) as TextBlock;
                        classID = classCRN.Text + className.Text + classSection.Text + classNumber.Text;
                    }
                    RemoveClass(classID);
                    RefreshGUI();
                }
            }
        }
        public void EditClass(string ID)
        {
            Unfocus_Overlay.Visibility = Visibility.Visible;
            Classes toEdit = DetermineClass(ID);
            EditClassDialog editClassDialog = new EditClassDialog(toEdit);
            editClassDialog.Owner = this;
            editClassDialog.ShowDialog();

            if ((bool)System.Windows.Application.Current.Resources["Set_Class_Success"])
            {
                bool conflict = false;
                bool check_conflicts = (bool)System.Windows.Application.Current.Resources["Edit_Class_Check"];
                if (check_conflicts)
                {
                    Classes temp = toEdit.DeepCopy();
                    Professors temp_Prof = DetermineProfessor((string)System.Windows.Application.Current.Resources["Set_Class_Professor"]);
                    temp.Prof = temp_Prof;
                    conflict = DetermineTimeConflict(temp, temp.ClassDay, temp.StartTime.TimeID, temp.StartTime.Meridian);
                    // flag down
                    System.Windows.Application.Current.Resources["Edit_Class_Check"] = false;
                }
                if (!conflict)
                {
                    bool originalOnline = toEdit.Online;
                    bool originalAssigned = toEdit.isAssigned;
                    string originalCRN = toEdit.CRN;
                    string originalBldg = toEdit.Classroom.Location;

                    if ((int)System.Windows.Application.Current.Resources["Set_Class_Term"] != toEdit.Term)
                    {
                        toEdit.Term = (int)System.Windows.Application.Current.Resources["Set_Class_Term"];
                    }
                    if ((int)System.Windows.Application.Current.Resources["Set_Class_Session"] != toEdit.Session)
                    {
                        toEdit.Session = (int)System.Windows.Application.Current.Resources["Set_Class_Session"];
                    }
                    if ((string)System.Windows.Application.Current.Resources["Set_Class_CRN"] != toEdit.CRN)
                    {
                        toEdit.CRN = (string)System.Windows.Application.Current.Resources["Set_Class_CRN"];
                    }
                    if ((string)System.Windows.Application.Current.Resources["Set_Class_Dept"] != toEdit.DeptName)
                    {
                        toEdit.DeptName = (string)System.Windows.Application.Current.Resources["Set_Class_Dept"];
                    }
                    if ((int)System.Windows.Application.Current.Resources["Set_Class_Number"] != toEdit.ClassNumber)
                    {
                        toEdit.ClassNumber = (int)System.Windows.Application.Current.Resources["Set_Class_Number"];
                    }
                    if ((int)System.Windows.Application.Current.Resources["Set_Class_Section"] != toEdit.SectionNumber)
                    {
                        toEdit.SectionNumber = (int)System.Windows.Application.Current.Resources["Set_Class_Section"];
                    }
                    if ((string)System.Windows.Application.Current.Resources["Set_Class_Name"] != toEdit.ClassName)
                    {
                        toEdit.ClassName = (string)System.Windows.Application.Current.Resources["Set_Class_Name"];
                    }
                    if ((int)System.Windows.Application.Current.Resources["Set_Class_Credits"] != toEdit.Credits)
                    {
                        toEdit.Credits = (int)System.Windows.Application.Current.Resources["Set_Class_Credits"];
                    }
                    if ((string)System.Windows.Application.Current.Resources["Set_Class_Professor"] != toEdit.Prof.SRUID)
                    {
                        toEdit.Prof = DetermineProfessor((string)System.Windows.Application.Current.Resources["Set_Class_Professor"]);
                    }
                    if ((bool)System.Windows.Application.Current.Resources["Set_Class_Online"] != toEdit.Online)
                    {
                        toEdit.Online = (bool)System.Windows.Application.Current.Resources["Set_Class_Online"];
                    }
                    if ((int)System.Windows.Application.Current.Resources["Set_Class_maxSeats"] != toEdit.MaxSeats)
                    {
                        toEdit.MaxSeats = (int)System.Windows.Application.Current.Resources["Set_Class_maxSeats"];
                    }
                    if ((int)System.Windows.Application.Current.Resources["Set_Class_projSeats"] != toEdit.ProjSeats)
                    {
                        toEdit.ProjSeats = (int)System.Windows.Application.Current.Resources["Set_Class_projSeats"];
                    }
                    if ((int)System.Windows.Application.Current.Resources["Set_Class_Enrolled"] != toEdit.Enrolled)
                    {
                        toEdit.Enrolled = (int)System.Windows.Application.Current.Resources["Set_Class_Enrolled"];
                    }
                    if ((int)System.Windows.Application.Current.Resources["Set_Class_Waitlist"] != toEdit.Waitlist)
                    {
                        toEdit.Waitlist = (int)System.Windows.Application.Current.Resources["Set_Class_Waitlist"];
                    }
                    if ((string)System.Windows.Application.Current.Resources["Set_Class_Crosslist"] != toEdit.Crosslist)
                    {
                        toEdit.Crosslist = (string)System.Windows.Application.Current.Resources["Set_Class_Crosslist"];

                        bool isCrossFirst = true;
                        for (int i = 0; i < classList.Count(); i++)
                        {
                            if (classList[i].Crosslist == toEdit.Crosslist && toEdit.TextBoxName != classList[i].TextBoxName)
                            {
                                isCrossFirst = false;
                            }
                        }
                        if (isCrossFirst == true)
                        {
                            toEdit.isCrossFirst = true;
                        }
                        else
                        {
                            toEdit.isCrossFirst = false;
                        }

                    }
                    if ((string)System.Windows.Application.Current.Resources["Set_Class_StartDate"] != toEdit.StartDate)
                    {
                        toEdit.StartDate = (string)System.Windows.Application.Current.Resources["Set_Class_StartDate"];
                    }
                    if ((string)System.Windows.Application.Current.Resources["Set_Class_EndDate"] != toEdit.EndDate)
                    {
                        toEdit.EndDate = (string)System.Windows.Application.Current.Resources["Set_Class_EndDate"];
                    }
                    bool appointment = false;
                    bool appt = (bool)System.Windows.Application.Current.Resources["Set_Class_Appointment"];
                    bool appt2 = (bool)System.Windows.Application.Current.Resources["Set_Class_Appointment2"];
                    if (appt || appt2)
                    {
                        appointment = true;
                    }
                    toEdit.isAppointment = appointment;

                    if (toEdit.Online)
                    {
                        toEdit.StartTime = new Timeslot();
                        toEdit.Classroom = new ClassRoom("WEB", 999);
                        toEdit.ClassDay = "";
                        toEdit.isAssigned = false;
                        toEdit.isAppointment = false;
                    }
                    else if (toEdit.isAppointment)
                    {
                        toEdit.StartTime = new Timeslot();
                        toEdit.ClassDay = "";
                        toEdit.isAssigned = false;
                        toEdit.Online = false;
                        if (appt)
                        {
                            toEdit.Classroom = new ClassRoom("APPT", 0);
                        }
                        else
                        {
                            toEdit.Classroom = new ClassRoom("APPT2", 0);
                        }
                    }
                    System.Windows.Application.Current.Resources["Set_Class_Success"] = false;
                }
                else
                {
                    MessageBoxButton button = MessageBoxButton.OK;
                    MessageBoxImage icon = MessageBoxImage.Exclamation;
                    System.Windows.MessageBox.Show("Professor is already teaching at that time  or Time interupts another class Time!\n\nReverting Changes...", "Invalid Edit", button, icon);
                }
            }
            Unfocus_Overlay.Visibility = Visibility.Hidden;
        }
        private void Btn_EditClass_Click(object sender, RoutedEventArgs e)
        {
            // find the class
            string ID = "";
            System.Windows.Controls.MenuItem mi = sender as System.Windows.Controls.MenuItem;
            if (mi != null)
            {
                System.Windows.Controls.ContextMenu cm = mi.CommandParameter as System.Windows.Controls.ContextMenu;
                if (cm != null)
                {
                    System.Windows.Controls.Label source = cm.PlacementTarget as System.Windows.Controls.Label;
                    if (source != null) // Being called from a Label
                    {
                        ID = source.Tag.ToString();
                    }
                    else // Being called from a GridRow
                    {
                        DataGridRow sourceRow = cm.PlacementTarget as DataGridRow;
                        System.Windows.Controls.DataGrid parentGrid = GetParent<System.Windows.Controls.DataGrid>(sourceRow as DependencyObject);

                        if (parentGrid.Columns[2].Header.Equals("CRN"))
                        {
                            TextBlock classCRN = parentGrid.Columns[2].GetCellContent(sourceRow) as TextBlock;
                            TextBlock className = parentGrid.Columns[6].GetCellContent(sourceRow) as TextBlock;
                            TextBlock classSection = parentGrid.Columns[5].GetCellContent(sourceRow) as TextBlock;
                            TextBlock classNumber = parentGrid.Columns[4].GetCellContent(sourceRow) as TextBlock;
                            ID = classCRN.Text + className.Text + classSection.Text + classNumber.Text;
                        }
                        else
                        {
                            TextBlock classCRN = parentGrid.Columns[0].GetCellContent(sourceRow) as TextBlock;
                            TextBlock className = parentGrid.Columns[4].GetCellContent(sourceRow) as TextBlock;
                            TextBlock classSection = parentGrid.Columns[3].GetCellContent(sourceRow) as TextBlock;
                            TextBlock classNumber = parentGrid.Columns[2].GetCellContent(sourceRow) as TextBlock;
                            ID = classCRN.Text + className.Text + classSection.Text + classNumber.Text;
                        }
                    }
                    EditClass(ID);
                    RefreshGUI();
                }
            }
        }
        public void EditClassTime(string ID)
        {
            Unfocus_Overlay.Visibility = Visibility.Visible;
            Classes toEdit = DetermineClass(ID);
            EditClassTimeDialog editClassDialog = new EditClassTimeDialog(toEdit);
            editClassDialog.Owner = this;
            editClassDialog.ShowDialog();
            Unfocus_Overlay.Visibility = Visibility.Hidden;
        }
        private void Btn_EditClassTime_Click(object sender, RoutedEventArgs e)
        {
            // find the class
            string ID = "";
            System.Windows.Controls.MenuItem mi = sender as System.Windows.Controls.MenuItem;
            if (mi != null)
            {
                System.Windows.Controls.ContextMenu cm = mi.CommandParameter as System.Windows.Controls.ContextMenu;
                if (cm != null)
                {
                    System.Windows.Controls.Label source = cm.PlacementTarget as System.Windows.Controls.Label;
                    if (source != null) // Being called from a Label
                    {
                        ID = source.Tag.ToString();
                    }
                    else // Being called from a GridRow
                    {
                        DataGridRow sourceRow = cm.PlacementTarget as DataGridRow;
                        System.Windows.Controls.DataGrid parentGrid = GetParent<System.Windows.Controls.DataGrid>(sourceRow as DependencyObject);
                        if (parentGrid.Columns[2].Equals("CRN"))
                        {                            
                            TextBlock classCRN = parentGrid.Columns[2].GetCellContent(sourceRow) as TextBlock;
                            TextBlock className = parentGrid.Columns[6].GetCellContent(sourceRow) as TextBlock;
                            TextBlock classSection = parentGrid.Columns[5].GetCellContent(sourceRow) as TextBlock;
                            TextBlock classNumber = parentGrid.Columns[4].GetCellContent(sourceRow) as TextBlock;
                            ID = classCRN.Text + className.Text + classSection.Text + classNumber.Text;
                        }
                        else
                        {
                            TextBlock classCRN = parentGrid.Columns[0].GetCellContent(sourceRow) as TextBlock;
                            TextBlock className = parentGrid.Columns[4].GetCellContent(sourceRow) as TextBlock;
                            TextBlock classSection = parentGrid.Columns[3].GetCellContent(sourceRow) as TextBlock;
                            TextBlock classNumber = parentGrid.Columns[2].GetCellContent(sourceRow) as TextBlock;
                            ID = classCRN.Text + className.Text + classSection.Text + classNumber.Text;
                        }

                    }

                    EditClassTime(ID);
                    RefreshGUI();
                }
            }
        }

        public void EditSingleDayClassTime(string ID)
        {
            Unfocus_Overlay.Visibility = Visibility.Visible;
            Classes toEdit = DetermineClass(ID);
            EditSingleDayClassTimeDialog editSingleDayClassDialog = new EditSingleDayClassTimeDialog(toEdit);
            editSingleDayClassDialog.Owner = this;
            editSingleDayClassDialog.ShowDialog();
            Unfocus_Overlay.Visibility = Visibility.Hidden;
        }

        private void Btn_EditSingleDayClassTime_Click(object sender, RoutedEventArgs e)
        {
            // find the class
            string ID = "";
            System.Windows.Controls.MenuItem mi = sender as System.Windows.Controls.MenuItem;
            if (mi != null)
            {
                System.Windows.Controls.ContextMenu cm = mi.CommandParameter as System.Windows.Controls.ContextMenu;
                if (cm != null)
                {
                    System.Windows.Controls.Label source = cm.PlacementTarget as System.Windows.Controls.Label;
                    if (source != null) // Being called from a Label
                    {
                        ID = source.Tag.ToString();
                    }
                    else // Being called from a GridRow
                    {
                        DataGridRow sourceRow = cm.PlacementTarget as DataGridRow;
                        System.Windows.Controls.DataGrid parentGrid = GetParent<System.Windows.Controls.DataGrid>(sourceRow as DependencyObject);
                        TextBlock classCRN = parentGrid.Columns[0].GetCellContent(sourceRow) as TextBlock;
                        TextBlock className = parentGrid.Columns[4].GetCellContent(sourceRow) as TextBlock;
                        TextBlock classSection = parentGrid.Columns[3].GetCellContent(sourceRow) as TextBlock;
                        TextBlock classNumber = parentGrid.Columns[2].GetCellContent(sourceRow) as TextBlock;
                        ID = classCRN.Text + className.Text + classSection.Text + classNumber.Text;
                    }
                    EditSingleDayClassTime(ID);
                    RefreshGUI();
                }
            }
        }


        private void Btn_MoveToUnassigned(object sender, RoutedEventArgs e)
        {
            string ID = "";
            System.Windows.Controls.MenuItem mi = sender as System.Windows.Controls.MenuItem;
            if (mi != null)
            {
                System.Windows.Controls.ContextMenu cm = mi.CommandParameter as System.Windows.Controls.ContextMenu;
                if (cm != null)
                {
                    System.Windows.Controls.Label source = cm.PlacementTarget as System.Windows.Controls.Label;
                    if (source != null) // Being called from a Label
                    {
                        ID = source.Tag.ToString();
                    }
                    else // Being called from a GridRow
                    {
                        DataGridRow sourceRow = cm.PlacementTarget as DataGridRow;
                        System.Windows.Controls.DataGrid parentGrid = GetParent<System.Windows.Controls.DataGrid>(sourceRow as DependencyObject);
                        TextBlock classCRN = parentGrid.Columns[0].GetCellContent(sourceRow) as TextBlock;
                        TextBlock className = parentGrid.Columns[4].GetCellContent(sourceRow) as TextBlock;
                        TextBlock classSection = parentGrid.Columns[3].GetCellContent(sourceRow) as TextBlock;
                        TextBlock classNumber = parentGrid.Columns[2].GetCellContent(sourceRow) as TextBlock;
                        ID = classCRN.Text + className.Text + classSection.Text + classNumber.Text;
                    }
                    MoveToUnassigned(ID);
                    RefreshGUI();
                }
            }

            
        }

        private void MoveToUnassigned(string ID)
        {
            Classes toEdit = DetermineClass(ID);

            toEdit.Classroom = new ClassRoom();
            toEdit.StartTime = new Timeslot();
            toEdit.ClassDay = "";
            toEdit.isAssigned = false;
        }
        public void CopyClass(string ID)
        {
            Classes copy;
            for (int i = 0; i < classList.Count; i++)
            {
                if (classList[i].ClassID == ID)
                {
                    copy = classList[i].DeepCopy();
                    // Change copy properties
                    copy.CRN = "NEW";
                    copy.isAssigned = false;
                    copy.ClassDay = "";
                    copy.StartTime = new Timeslot();
                    copy.Classroom = new ClassRoom();
                    // Add to classlist
                    classList.Add(copy);
                    break;
                }
            }
        }
        private void Btn_CopyClass_Click(object sender, RoutedEventArgs e)
        {
            // find the class
            string classID = "";
            System.Windows.Controls.MenuItem mi = sender as System.Windows.Controls.MenuItem;
            if (mi != null)
            {
                System.Windows.Controls.ContextMenu cm = mi.CommandParameter as System.Windows.Controls.ContextMenu;
                if (cm != null)
                {
                    System.Windows.Controls.Label source = cm.PlacementTarget as System.Windows.Controls.Label;
                    if (source != null) // Being called from a Label
                    {
                        classID = source.Tag.ToString();
                    }
                    else // Being called from a GridRow
                    {
                        DataGridRow sourceRow = cm.PlacementTarget as DataGridRow;
                        System.Windows.Controls.DataGrid parentGrid = GetParent<System.Windows.Controls.DataGrid>(sourceRow as DependencyObject);
                        TextBlock classCRN = parentGrid.Columns[0].GetCellContent(sourceRow) as TextBlock;
                        TextBlock className = parentGrid.Columns[4].GetCellContent(sourceRow) as TextBlock;
                        TextBlock classSection = parentGrid.Columns[3].GetCellContent(sourceRow) as TextBlock;
                        TextBlock classNumber = parentGrid.Columns[2].GetCellContent(sourceRow) as TextBlock;
                        classID = classCRN.Text + className.Text + classSection.Text + classNumber.Text;
                    }
                    CopyClass(classID);
                    RefreshGUI();
                }
            }
        }
        public void EditNotes(string ID)
        {
            Classes c1 = DetermineClass(ID);
            Unfocus_Overlay.Visibility = Visibility.Visible;
            EditNotesDialog editNotesDialog = new EditNotesDialog(c1);
            editNotesDialog.Owner = this;
            editNotesDialog.ShowDialog();
            Unfocus_Overlay.Visibility = Visibility.Hidden;
        }
        private void Btn_EditNotes_Click(object sender, RoutedEventArgs e)
        {
            // find the class
            string classID = "";
            System.Windows.Controls.MenuItem mi = sender as System.Windows.Controls.MenuItem;
            if (mi != null)
            {
                System.Windows.Controls.ContextMenu cm = mi.CommandParameter as System.Windows.Controls.ContextMenu;
                if (cm != null)
                {
                    System.Windows.Controls.Label source = cm.PlacementTarget as System.Windows.Controls.Label;
                    if (source != null) // Being called from a Label
                    {
                        classID = source.Tag.ToString();
                    }
                    else // Being called from a GridRow
                    {
                        DataGridRow sourceRow = cm.PlacementTarget as DataGridRow;
                        System.Windows.Controls.DataGrid parentGrid = GetParent<System.Windows.Controls.DataGrid>(sourceRow as DependencyObject);
                        TextBlock classCRN = parentGrid.Columns[0].GetCellContent(sourceRow) as TextBlock;
                        TextBlock className = parentGrid.Columns[4].GetCellContent(sourceRow) as TextBlock;
                        TextBlock classSection = parentGrid.Columns[3].GetCellContent(sourceRow) as TextBlock;
                        TextBlock classNumber = parentGrid.Columns[2].GetCellContent(sourceRow) as TextBlock;
                        classID = classCRN.Text + className.Text + classSection.Text + classNumber.Text;
                    }
                    EditNotes(classID);
                    RefreshGUI();
                }
            }
        }
        public void HideClass(string ID)
        {
            Classes c1 = DetermineClass(ID);
            // maxseats = ExtraData[2]
            // if hidden -> unhide / else hide
            if (c1.isHidden)
            {
                c1.MaxSeats = 1;
            }
            else
            {
                c1.MaxSeats = 0;
            }
        }
        private void Btn_HideClass_Click(object sender, RoutedEventArgs e)
        {
            // find the class
            string classID = "";
            System.Windows.Controls.MenuItem mi = sender as System.Windows.Controls.MenuItem;
            if (mi != null)
            {
                System.Windows.Controls.ContextMenu cm = mi.CommandParameter as System.Windows.Controls.ContextMenu;
                if (cm != null)
                {
                    System.Windows.Controls.Label source = cm.PlacementTarget as System.Windows.Controls.Label;
                    if (source != null) // Being called from a Label
                    {
                        classID = source.Tag.ToString();
                    }
                    else // Being called from a GridRow
                    {
                        DataGridRow sourceRow = cm.PlacementTarget as DataGridRow;
                        System.Windows.Controls.DataGrid parentGrid = GetParent<System.Windows.Controls.DataGrid>(sourceRow as DependencyObject);
                        TextBlock classCRN = parentGrid.Columns[0].GetCellContent(sourceRow) as TextBlock;
                        TextBlock className = parentGrid.Columns[4].GetCellContent(sourceRow) as TextBlock;
                        TextBlock classSection = parentGrid.Columns[3].GetCellContent(sourceRow) as TextBlock;
                        TextBlock classNumber = parentGrid.Columns[2].GetCellContent(sourceRow) as TextBlock;
                        classID = classCRN.Text + className.Text + classSection.Text + classNumber.Text;
                    }
                    HideClass(classID);
                    RefreshGUI();
                }
            }
        }
        private void EditSeats(string ID)
        {
            Unfocus_Overlay.Visibility = Visibility.Visible;
            Classes toEdit = DetermineClass(ID);
            EditClassSeating editClassSeating = new EditClassSeating(toEdit);
            editClassSeating.Owner = this;
            editClassSeating.ShowDialog();
            Unfocus_Overlay.Visibility = Visibility.Hidden;
        }
        private void Btn_EditSeats_Click(object sender, RoutedEventArgs e)
        {
            // find the class
            string classID = "";
            System.Windows.Controls.MenuItem mi = sender as System.Windows.Controls.MenuItem;
            if (mi != null)
            {
                System.Windows.Controls.ContextMenu cm = mi.CommandParameter as System.Windows.Controls.ContextMenu;
                if (cm != null)
                {
                    System.Windows.Controls.Label source = cm.PlacementTarget as System.Windows.Controls.Label;
                    if (source != null) // Being called from a Label
                    {
                        classID = source.Tag.ToString();
                    }
                    else // Being called from a GridRow
                    {
                        DataGridRow sourceRow = cm.PlacementTarget as DataGridRow;
                        System.Windows.Controls.DataGrid parentGrid = GetParent<System.Windows.Controls.DataGrid>(sourceRow as DependencyObject);
                        TextBlock classCRN = parentGrid.Columns[0].GetCellContent(sourceRow) as TextBlock;
                        TextBlock className = parentGrid.Columns[4].GetCellContent(sourceRow) as TextBlock;
                        TextBlock classSection = parentGrid.Columns[3].GetCellContent(sourceRow) as TextBlock;
                        TextBlock classNumber = parentGrid.Columns[2].GetCellContent(sourceRow) as TextBlock;
                        classID = classCRN.Text + className.Text + classSection.Text + classNumber.Text;
                    }
                    EditSeats(classID);
                    RefreshGUI();
                }
            }
        }

        // DRAG/DROP functionality
        private void MouseMoveOnGridRow(object sender, System.Windows.Input.MouseEventArgs e) // Handles DRAG operation on class list items 
        {
            TextBlock cellUnderMouse = sender as TextBlock;
            if (cellUnderMouse != null && e.LeftButton == MouseButtonState.Pressed)
            {
                DataGridRow row = DataGridRow.GetRowContainingElement(cellUnderMouse);
                DragDrop.DoDragDrop(Unassigned_Classes_Grid, row, System.Windows.DragDropEffects.Copy);
            }
        }
        private void MouseMoveOnAssignedClass(object sender, System.Windows.Input.MouseEventArgs e) // Handles DRAG operation on GUI classes box 
        {
            System.Windows.Controls.Label labelUnderMouse = sender as System.Windows.Controls.Label;
            int classIndex = -1;
            if ((labelUnderMouse != null) && (e.LeftButton == MouseButtonState.Pressed) && (labelUnderMouse.Tag != null) && (labelUnderMouse.Tag.ToString() != ""))
            {
                // find index of class being represented by the label
                for (int i = 0; i < classList.Count; i++)
                {
                    if (classList[i].ClassID == labelUnderMouse.Tag.ToString())
                    {
                        classIndex = i;
                        break;
                    }
                }
                // Package the data
                System.Windows.DataObject data = new System.Windows.DataObject();
                data.SetData(typeof(int), classIndex);
                data.SetData(typeof(object), labelUnderMouse);
                // send dataObject
                DragDrop.DoDragDrop(labelUnderMouse, data, System.Windows.DragDropEffects.Copy);
            }
        }
        private void HandleDropToList(Object sender, System.Windows.DragEventArgs e) // Handles DROP operation to unassigned classes list 
        {
            System.Windows.Controls.Label sourceLabel = (System.Windows.Controls.Label)e.Data.GetData(typeof(object));
            if (sourceLabel != null)
            {
                int classIndex = (int)e.Data.GetData(typeof(int));
                // clear the Label
                sourceLabel.Content = "";
                RGB_Color white_bg = new RGB_Color(255, 255, 255);
                sourceLabel.Background = white_bg.colorBrush2;
                sourceLabel.ContextMenu = null;
                // add the class to unassigned class list
                classList[classIndex].Classroom = new ClassRoom();
                classList[classIndex].ClassDay = "";
                classList[classIndex].StartTime = new Timeslot();
                classList[classIndex].isAssigned = false;

            }
            else
            {
                string ID = "";
                DataGridRow droppedRow = (DataGridRow)e.Data.GetData(typeof(DataGridRow));
                if (droppedRow != null)
                {
                    TextBlock classCRN = Unassigned_Classes_Grid.Columns[0].GetCellContent(droppedRow) as TextBlock;
                    TextBlock className = Unassigned_Classes_Grid.Columns[4].GetCellContent(droppedRow) as TextBlock;
                    TextBlock classSection = Unassigned_Classes_Grid.Columns[3].GetCellContent(droppedRow) as TextBlock;
                    TextBlock classNumber = Unassigned_Classes_Grid.Columns[2].GetCellContent(droppedRow) as TextBlock;
                    ID = classCRN.Text + className.Text + classSection.Text + classNumber.Text;
                    Classes theClass = DetermineClass(ID);
                    string classType = "";

                    if (theClass.Online || theClass.isAppointment)
                    {
                        if (theClass.Online)
                        {
                            classType = "Online";
                        }
                        else if (theClass.isAppointment)
                        {
                            classType = "Appointment";
                        }
                        string messageBoxText = "Are you sure you want to change this class\nfrom " + classType + " to In-Class?";
                        string caption = classType + " class alteration";
                        MessageBoxButton button = MessageBoxButton.YesNoCancel;
                        MessageBoxImage icon = MessageBoxImage.Question;
                        // Display + Process message box results
                        MessageBoxResult result = System.Windows.MessageBox.Show(messageBoxText, caption, button, icon);
                        switch (result)
                        {
                            case MessageBoxResult.Yes:
                                // Find the class
                                for (int i = 0; i < classList.Count; i++)
                                {
                                    if (classList[i].ClassID == ID)
                                    {
                                        if (classType == "Online")
                                        {
                                            classList[i].Online = false;
                                        }
                                        else if (classType == "Appointment")
                                        {
                                            classList[i].isAppointment = false;
                                        }
                                        classList[i].Classroom = new ClassRoom();
                                    }
                                }
                                break;
                            case MessageBoxResult.No:
                                break;
                            case MessageBoxResult.Cancel:
                                break;
                        }
                    }
                }
            }
            RefreshGUI();
        }
        private void HandleDropToCell(Object sender, System.Windows.DragEventArgs e) // Handles DROP operation to GUI classes box 
        {
            System.Windows.Controls.Label sourceLabel = (System.Windows.Controls.Label)e.Data.GetData(typeof(object));
            System.Windows.Controls.Label receiver = sender as System.Windows.Controls.Label;
            if (receiver.Content.ToString() == "")
            {
                if (sourceLabel != null)
                {
                    int classIndex = (int)e.Data.GetData(typeof(int));
                    // parse target slot info
                    string days = receiver.Name.Split('_')[0];
                    string start = receiver.Name.Split('_')[1];
                    string roomInfoFix = receiver.Name.Split('_')[2];
                    string roomInfo = roomInfoFix.Substring(0, roomInfoFix.Length - 2);
                    string bldg = roomInfo.Substring(0, 3);
                    int room = Int32.Parse(roomInfo.Substring(3, (roomInfo.Length - 3)));
                    string meridian = receiver.Name.Substring(receiver.Name.Length - 2);
                    bool isConflict = DetermineTimeConflict(classList[classIndex], days, start, meridian);
                    if (!isConflict)
                    {
                        if (classList[classIndex].ClassDay == "MWF" || classList[classIndex].ClassDay == "TR")
                        {
                            classList[classIndex].ClassDay = days;
                        }
                        else if (classList[classIndex].ClassDay == "M" || classList[classIndex].ClassDay == "T" || classList[classIndex].ClassDay == "W" || classList[classIndex].ClassDay == "R" || classList[classIndex].ClassDay == "F")
                        {
                            classList[classIndex].ClassDay = classList[classIndex].ClassDay;
                        }
                        classList[classIndex].StartTime = DetermineTime(start, days, room, meridian);
                        classList[classIndex].Classroom = DetermineClassroom(bldg, room);

                        // Give the newLabel the class information
                        receiver.Content = sourceLabel.Content;
                        receiver.Background = sourceLabel.Background;
                        receiver.Tag = sourceLabel.Tag;
                        receiver.ToolTip = sourceLabel.ToolTip;
                        receiver.ContextMenu = Resources["ClassContextMenu"] as System.Windows.Controls.ContextMenu;

                        // clear the sourceLabel
                        sourceLabel.Content = "";
                        RGB_Color white_bg = new RGB_Color(255, 255, 255);
                        sourceLabel.Background = white_bg.colorBrush2;
                        sourceLabel.Tag = null;
                        sourceLabel.ToolTip = null;
                        sourceLabel.ContextMenu = null;
                    }
                    else
                    {
                        MessageBoxButton button = MessageBoxButton.OK;
                        MessageBoxImage icon = MessageBoxImage.Exclamation;
                        System.Windows.MessageBox.Show("Professor is already teaching at that time!", "Invalid action", button, icon);
                    }
                }
                else
                {
                    string ID = "";
                    DataGridRow droppedRow = (DataGridRow)e.Data.GetData(typeof(DataGridRow));
                    if (droppedRow == null)
                    {
                        System.Windows.MessageBox.Show("dropped row was null");
                    }
                    else
                    {
                        TextBlock classCRN = Unassigned_Classes_Grid.Columns[0].GetCellContent(droppedRow) as TextBlock;
                        TextBlock className = Unassigned_Classes_Grid.Columns[4].GetCellContent(droppedRow) as TextBlock;
                        TextBlock classSection = Unassigned_Classes_Grid.Columns[3].GetCellContent(droppedRow) as TextBlock;
                        TextBlock classNumber = Unassigned_Classes_Grid.Columns[2].GetCellContent(droppedRow) as TextBlock;
                        ID = classCRN.Text + className.Text + classSection.Text + classNumber.Text;
                    }

                    /// VALIDATION CHECKS GO HERE ///
                    // check if its online class
                    bool validOperation = true;
                    int classIndex = -1;
                    for (int i = 0; i < classList.Count; i++)
                    {
                        if (classList[i].ClassID == ID)
                        {
                            if (classList[i].Online || classList[i].isAppointment)
                            {
                                string classType = "";
                                if (classList[i].Online)
                                {
                                    classType = "Online";
                                }
                                else if (classList[i].isAppointment)
                                {
                                    classType = "Appointment";
                                }
                                string messageBoxText = "Are you sure you want to change this class from " + classType + " to In-Class?";
                                string caption = classType + " class warning";
                                MessageBoxButton button = MessageBoxButton.YesNoCancel;
                                MessageBoxImage icon = MessageBoxImage.Question;
                                // Display + process message box results
                                MessageBoxResult result = System.Windows.MessageBox.Show(messageBoxText, caption, button, icon);
                                switch (result)
                                {
                                    case MessageBoxResult.Yes:
                                        break;
                                    case MessageBoxResult.No:
                                        validOperation = false;
                                        break;
                                    case MessageBoxResult.Cancel:
                                        validOperation = false;
                                        break;
                                }
                            }
                            classIndex = i;
                            break;
                        }
                    }
                    if (validOperation)
                    {
                        string days = receiver.Name.Split('_')[0];
                        string start = receiver.Name.Split('_')[1];
                        string roomInfo = receiver.Name.Split('_')[2];
                        string bldg = roomInfo.Substring(0, 3);
                        string meridian = roomInfo.Substring(roomInfo.Length - 2);
                        int room = Int32.Parse(roomInfo.Substring(3, (roomInfo.Length - 5)));
                        if (!DetermineTimeConflict(classList[classIndex], days, start, meridian))
                        {
                            if (!classList[classIndex].Online)
                            {
                                if (classList[classIndex].Classroom.Location.Contains("APPT")) // its by appointment
                                {
                                    if (classList[classIndex].Classroom.Location == "APPT")
                                    {
                                        classList[classIndex].isAssigned = true;
                                        classList[classIndex].isAppointment = false;
                                    }
                                    else if (classList[classIndex].Classroom.Location == "APPT2")
                                    {
                                        classList[classIndex].isAssigned = true;
                                        classList[classIndex].isAppointment = false;
                                    }
                                    else
                                    {
                                        System.Windows.MessageBox.Show("Couldnt remove appointed class from respective list");
                                    }
                                }
                                else // its unassigned
                                {
                                    classList[classIndex].isAssigned = true;
                                }
                            }
                            else // its online
                            {
                                classList[classIndex].Online = false;
                            }
                            // Update class in masterlist = give it a start time + classroom
                            classList[classIndex].ClassDay = days;
                            classList[classIndex].StartTime = DetermineTime(start, days, room, meridian);
                            classList[classIndex].Classroom = DetermineClassroom(bldg, room);
                            // Give the Label the class information
                            receiver.Content = classList[classIndex].TextBoxName;
                            if (classList[classIndex].isHidden)
                            {
                                receiver.Background = stripedBackground(classList[classIndex].Prof.profRGB.colorBrush);
                            }
                            else
                            {
                                receiver.Background = classList[classIndex].Prof.Prof_Color;
                            }
                            receiver.Tag = classList[classIndex].ClassID;
                            receiver.ToolTip = classList[classIndex].ToolTipText;
                            receiver.ContextMenu = Resources["ClassContextMenu"] as System.Windows.Controls.ContextMenu;
                        }
                        else
                        {
                            MessageBoxButton button = MessageBoxButton.OK;
                            MessageBoxImage icon = MessageBoxImage.Exclamation;
                            System.Windows.MessageBox.Show("Professor is already teaching at that time!", "Invalid action", button, icon);
                        }
                    }
                }
                RefreshGUI();
            }
            else
            {
                if (sourceLabel != null)
                {
                    if (sourceLabel.Content.ToString() != receiver.Content.ToString())
                    {
                        MessageBoxButton button = MessageBoxButton.OK;
                        MessageBoxImage icon = MessageBoxImage.Exclamation;
                        System.Windows.MessageBox.Show("Timeslot is already taken!", "Invalid action", button, icon);
                    }
                }
                else
                {
                    MessageBoxButton button = MessageBoxButton.OK;
                    MessageBoxImage icon = MessageBoxImage.Exclamation;
                    System.Windows.MessageBox.Show("Timeslot is already taken!", "Invalid action", button, icon);
                }
            }
        }
        private void HandleDropToOnlineList(Object sender, System.Windows.DragEventArgs e) // Handles DROP operation to online classes list 
        {
            System.Windows.Controls.Label sourceLabel = (System.Windows.Controls.Label)e.Data.GetData(typeof(object));
            if (sourceLabel != null)
            {
                string messageBoxText = "Are you sure you want to change this\nIn-Class course to Online format?";
                string caption = "Online class alteration";
                MessageBoxButton button = MessageBoxButton.YesNoCancel;
                MessageBoxImage icon = MessageBoxImage.Question;
                // Display + Process message box results
                MessageBoxResult result = System.Windows.MessageBox.Show(messageBoxText, caption, button, icon);
                switch (result)
                {
                    case MessageBoxResult.Yes:
                        // User pressed Yes button
                        int classIndex = (int)e.Data.GetData(typeof(int));
                        // clear the Label
                        sourceLabel.Content = "";
                        RGB_Color white_bg = new RGB_Color(255, 255, 255);
                        sourceLabel.Background = white_bg.colorBrush2;
                        sourceLabel.ContextMenu = Resources["ClassContextMenu"] as System.Windows.Controls.ContextMenu;
                        // add the class to online class list
                        classList[classIndex].Classroom = new ClassRoom("WEB", 999);
                        classList[classIndex].ClassDay = "";
                        classList[classIndex].StartTime = new Timeslot();
                        classList[classIndex].isAssigned = false;
                        classList[classIndex].Online = true;
                        break;
                    case MessageBoxResult.No:
                        break;
                    case MessageBoxResult.Cancel:
                        break;
                }
            }
            else // Comes from unassigned list
            {
                DataGridRow droppedRow = (DataGridRow)e.Data.GetData(typeof(DataGridRow));
                if (droppedRow != null)
                {
                    TextBlock classCRN = Online_Classes_Grid.Columns[0].GetCellContent(droppedRow) as TextBlock;
                    TextBlock className = Online_Classes_Grid.Columns[4].GetCellContent(droppedRow) as TextBlock;
                    TextBlock classSection = Online_Classes_Grid.Columns[3].GetCellContent(droppedRow) as TextBlock;
                    TextBlock classNumber = Online_Classes_Grid.Columns[2].GetCellContent(droppedRow) as TextBlock;
                    string classID = classCRN.Text + className.Text + classSection.Text + classNumber.Text;
                    Classes theClass = DetermineClass(classID);
                    if (!theClass.Online)
                    {
                        string messageBoxText = "Are you sure you want to change this\nCourse to Online format?";
                        string caption = "Online class alteration";
                        MessageBoxButton button = MessageBoxButton.YesNoCancel;
                        MessageBoxImage icon = MessageBoxImage.Question;
                        // Display + Process message box results
                        MessageBoxResult result = System.Windows.MessageBox.Show(messageBoxText, caption, button, icon);
                        switch (result)
                        {
                            case MessageBoxResult.Yes:
                                // User pressed Yes button
                                // Find the class
                                for (int i = 0; i < classList.Count; i++)
                                {
                                    if (classList[i].ClassID == classID)
                                    {
                                        classList[i].Online = true;
                                        classList[i].Classroom = new ClassRoom("WEB", 999);
                                    }
                                }
                                break;
                            case MessageBoxResult.No:
                                break;
                            case MessageBoxResult.Cancel:
                                break;
                        }
                    }
                }
            }
            RefreshGUI();
        }
        private void HandleDropToAppointmentList(Object sender, System.Windows.DragEventArgs e) // Handles DROP operation to appointment classes list 
        {
            System.Windows.Controls.Label sourceLabel = (System.Windows.Controls.Label)e.Data.GetData(typeof(object));
            if (sourceLabel != null)
            {
                string messageBoxText = "Are you sure you want to change this\nIn-Class course to 'By Appointment' format?";
                string caption = "By Appointment class alteration";
                MessageBoxButton button = MessageBoxButton.YesNoCancel;
                MessageBoxImage icon = MessageBoxImage.Question;
                // Display + Process message box results
                MessageBoxResult result = System.Windows.MessageBox.Show(messageBoxText, caption, button, icon);
                switch (result)
                {
                    case MessageBoxResult.Yes:
                        // User pressed Yes button
                        int classIndex = (int)e.Data.GetData(typeof(int));
                        // clear the Label
                        sourceLabel.Content = "";
                        RGB_Color white_bg = new RGB_Color(255, 255, 255);
                        sourceLabel.Background = white_bg.colorBrush2;
                        sourceLabel.ContextMenu = Resources["ClassContextMenu"] as System.Windows.Controls.ContextMenu;
                        // add the class to online class list
                        classList[classIndex].Classroom = new ClassRoom("APPT", 0);
                        classList[classIndex].ClassDay = "";
                        classList[classIndex].StartTime = new Timeslot();
                        classList[classIndex].isAssigned = false;
                        classList[classIndex].Online = false;
                        classList[classIndex].isAppointment = true;
                        break;
                    case MessageBoxResult.No:
                        break;
                    case MessageBoxResult.Cancel:
                        break;
                }
            }
            else // Its from unassigned list
            {
                DataGridRow droppedRow = (DataGridRow)e.Data.GetData(typeof(DataGridRow));
                if (droppedRow != null)
                {
                    TextBlock classCRN = Online_Classes_Grid.Columns[0].GetCellContent(droppedRow) as TextBlock;
                    TextBlock className = Online_Classes_Grid.Columns[4].GetCellContent(droppedRow) as TextBlock;
                    TextBlock classSection = Online_Classes_Grid.Columns[3].GetCellContent(droppedRow) as TextBlock;
                    TextBlock classNumber = Online_Classes_Grid.Columns[2].GetCellContent(droppedRow) as TextBlock;
                    string classID = classCRN.Text + className.Text + classSection.Text + classNumber.Text;
                    Classes theClass = DetermineClass(classID);
                    if (!theClass.isAppointment)
                    {
                        string messageBoxText = "Are you sure you want to change this\nCourse to 'By Appointment' format?";
                        string caption = "By Appointment class alteration";
                        MessageBoxButton button = MessageBoxButton.YesNoCancel;
                        MessageBoxImage icon = MessageBoxImage.Question;
                        // Display + Process message box results
                        MessageBoxResult result = System.Windows.MessageBox.Show(messageBoxText, caption, button, icon);
                        switch (result)
                        {
                            case MessageBoxResult.Yes:
                                // User pressed Yes button
                                // Find the class
                                for (int i = 0; i < classList.Count; i++)
                                {
                                    if (classList[i].ClassID == classID)
                                    {
                                        classList[i].isAppointment = true;
                                        classList[i].Classroom = new ClassRoom("APPT", 0);
                                    }
                                }
                                break;
                            case MessageBoxResult.No:
                                break;
                            case MessageBoxResult.Cancel:
                                break;
                        }
                    }
                }
            }
            RefreshGUI();
        }
        private void HandleDropToAppointment2List(Object sender, System.Windows.DragEventArgs e) // Handles DROP operation to appointment2 classes list 
        {
            System.Windows.Controls.Label sourceLabel = (System.Windows.Controls.Label)e.Data.GetData(typeof(object));
            if (sourceLabel != null)
            {
                string messageBoxText = "Are you sure you want to change this\nIn-Class course to 'By Appointment' format?";
                string caption = "Appointment class alteration";
                MessageBoxButton button = MessageBoxButton.YesNoCancel;
                MessageBoxImage icon = MessageBoxImage.Question;
                // Display + Process message box results
                MessageBoxResult result = System.Windows.MessageBox.Show(messageBoxText, caption, button, icon);
                switch (result)
                {
                    case MessageBoxResult.Yes:
                        // User pressed Yes button
                        int classIndex = (int)e.Data.GetData(typeof(int));
                        // clear the Label
                        sourceLabel.Content = "";
                        RGB_Color white_bg = new RGB_Color(255, 255, 255);
                        sourceLabel.Background = white_bg.colorBrush2;
                        sourceLabel.ContextMenu = Resources["ClassContextMenu"] as System.Windows.Controls.ContextMenu;
                        // add the class to online class list
                        classList[classIndex].Classroom = new ClassRoom("APPT2", 0);
                        classList[classIndex].ClassDay = "";
                        classList[classIndex].StartTime = new Timeslot();
                        classList[classIndex].isAssigned = false;
                        classList[classIndex].Online = false;
                        classList[classIndex].isAppointment = true;
                        break;
                    case MessageBoxResult.No:
                        break;
                    case MessageBoxResult.Cancel:
                        break;
                }
            }
            else // Its from unassigned list
            {
                DataGridRow droppedRow = (DataGridRow)e.Data.GetData(typeof(DataGridRow));
                if (droppedRow != null)
                {
                    TextBlock classCRN = Online_Classes_Grid.Columns[0].GetCellContent(droppedRow) as TextBlock;
                    TextBlock className = Online_Classes_Grid.Columns[4].GetCellContent(droppedRow) as TextBlock;
                    TextBlock classSection = Online_Classes_Grid.Columns[3].GetCellContent(droppedRow) as TextBlock;
                    TextBlock classNumber = Online_Classes_Grid.Columns[2].GetCellContent(droppedRow) as TextBlock;
                    string classID = classCRN.Text + className.Text + classSection.Text + classNumber.Text;
                    Classes theClass = DetermineClass(classID);
                    if (!theClass.isAppointment)
                    {
                        string messageBoxText = "Are you sure you want to change this\nCourse to 'By Appointment' format?";
                        string caption = "Appointment class alteration";
                        MessageBoxButton button = MessageBoxButton.YesNoCancel;
                        MessageBoxImage icon = MessageBoxImage.Question;
                        // Display + Process message box results
                        MessageBoxResult result = System.Windows.MessageBox.Show(messageBoxText, caption, button, icon);
                        switch (result)
                        {
                            case MessageBoxResult.Yes:
                                // User pressed Yes button
                                // Find the class
                                for (int i = 0; i < classList.Count; i++)
                                {
                                    if (classList[i].ClassID == classID)
                                    {
                                        classList[i].isAppointment = true;
                                        classList[i].Classroom = new ClassRoom("APPT2", 0);
                                    }
                                }
                                break;
                            case MessageBoxResult.No:
                                break;
                            case MessageBoxResult.Cancel:
                                break;
                        }
                    }
                }
            }
            RefreshGUI();
        }

        // Utility functions
        public RGB_Color StringToRGB(string s) // Converts rgb string to a RGB_Color object 
        {
            RGB_Color color;
            String[] parts = s.Split('.');
            color = new RGB_Color(Byte.Parse(parts[0]), Byte.Parse(parts[1]), Byte.Parse(parts[2]));
            return color;
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
        public Timeslot DetermineTime(string startTime, string classDay, int classroom, string meridian) // Finds corresponding Timeslot object based on start time and class day 
        {
            string id = startTime.Substring(0, 2);
            
            if (classroom == 0)
            {
                // Check for Night Classes first 2.5 hour classes
                if (classDay == "NM" || classDay == "NIGHT_MONDAY")
                {
                    for (int i = 0; i < times_Night_Monday.Length; i++)
                    {
                        if (times_Night_Monday[i].TimeID == id && times_Night_Monday[i].Meridian == meridian)
                        {
                            return times_Night_Monday[i];
                        }
                    }
                }
                else if (classDay == "NT" || classDay == "NIGHT_TUESDAY")
                {
                    for (int i = 0; i < times_Night_Tuesday.Length; i++)
                    {
                        if (times_Night_Tuesday[i].TimeID == id && times_Night_Tuesday[i].Meridian == meridian)
                        {
                            return times_Night_Tuesday[i];
                        }
                    }
                }
                else if (classDay == "NW" || classDay == "NIGHT_WEDNESDAY")
                {
                    for (int i = 0; i < times_Night_Wednesday.Length; i++)
                    {
                        if (times_Night_Wednesday[i].TimeID == id && times_Night_Wednesday[i].Meridian == meridian)
                        {
                            return times_Night_Wednesday[i];
                        }
                    }
                }
                else if (classDay == "NR" || classDay == "NIGHT_THURSDAY")
                {
                    for (int i = 0; i < times_Night_Thursday.Length; i++)
                    {
                        if (times_Night_Thursday[i].TimeID == id && times_Night_Thursday[i].Meridian == meridian)
                        {
                            return times_Night_Thursday[i];
                        }
                    }
                }
                else if (classDay == "NF" || classDay == "NIGHT_FRIDAY")
                {
                    for (int i = 0; i < times_Night_Friday.Length; i++)
                    {
                        if (times_Night_Friday[i].TimeID == id && times_Night_Friday[i].Meridian == meridian)
                        {
                            return times_Night_Friday[i];
                        }
                    }
                }
                // Regular day classes
                else if (classDay == "MWF" || classDay == "M" || classDay == "W" || classDay == "F")
                {

                    for (int i = 0; i < masterTimeslotList.Count; i++)
                    {
                       
                        if (masterTimetableList[i] == 0)
                        {
                            if (masterTimeslotList[i].TimeID == id && masterTimeslotList[i].Meridian == meridian)
                            {
                                return masterTimeslotList[i];
                            }
                        }
                    }
                    // Check times_Default for night class patterns
                    for (int i = 0; i < times_Default.Count(); i++)
                    {
                        if (times_Default_Timetable[i] == "NM" || times_Default_Timetable[i] == "NIGHT_MONDAY" ||
                            times_Default_Timetable[i] == "NT" || times_Default_Timetable[i] == "NIGHT_TUESDAY" ||
                            times_Default_Timetable[i] == "NW" || times_Default_Timetable[i] == "NIGHT_WEDNESDAY" ||
                            times_Default_Timetable[i] == "NR" || times_Default_Timetable[i] == "NIGHT_THURSDAY" ||
                            times_Default_Timetable[i] == "NF" || times_Default_Timetable[i] == "NIGHT_FRIDAY")
                        {
                            if (times_Default[i].TimeID == id && times_Default[i].Meridian == meridian)
                            {
                                return times_Default[i];
                            }
                        }
                    }
                    for (int i = 0; i < times_Default.Count(); i++)
                    {
                        if (times_Default_Timetable[i] == "MWF")
                        {
                            if (times_Default[i].TimeID == id && times_Default[i].Meridian == meridian)
                            {
                                return times_Default[i];
                            }
                        }
                    }
                    // Check if this  a night class on Monday/Wednesday/Friday 
                    if (meridian == "PM" && (id == "06" || id == "08"))
                    {
                        //  Monday night class first
                        for (int i = 0; i < times_Night_Monday.Length; i++)
                        {
                            if (times_Night_Monday[i].TimeID == id && times_Night_Monday[i].Meridian == meridian)
                            {
                                return times_Night_Monday[i];
                            }
                        }
                        //  Wednesday night class
                        for (int i = 0; i < times_Night_Wednesday.Length; i++)
                        {
                            if (times_Night_Wednesday[i].TimeID == id && times_Night_Wednesday[i].Meridian == meridian)
                            {
                                return times_Night_Wednesday[i];
                            }
                        }
                        //  Friday night class
                        for (int i = 0; i < times_Night_Friday.Length; i++)
                        {
                            if (times_Night_Friday[i].TimeID == id && times_Night_Friday[i].Meridian == meridian)
                            {
                                return times_Night_Friday[i];
                            }
                        }
                    }
                    for (int i = 0; i < times_MWF.Length; i++)
                    {
                        if (times_MWF[i].TimeID == id && times_MWF[i].Meridian == meridian)
                        {
                            return times_MWF[i];
                        }
                    }
                }
                else
                {

                    for (int i = 0; i < masterTimeslotList.Count; i++)
                    {
                        if (masterTimetableList[i] == 1)
                        {
                            if (masterTimeslotList[i].TimeID == id && masterTimeslotList[i].Meridian == meridian)
                            {
                                return masterTimeslotList[i];
                            }
                        }
                    }
                    for (int i = 0; i < times_Default.Count(); i++)
                    {
                        if (times_Default_Timetable[i] == "TR")
                        {
                            if (times_Default[i].TimeID == id && times_Default[i].Meridian == meridian)
                            {
                                return times_Default[i];
                            }
                        }
                    }
                    // Check if this a night class on Tuesday or Thursday 
                    if (meridian == "PM" && (id == "06" || id == "08"))
                    {
                        //  Tuesday night class first
                        for (int i = 0; i < times_Night_Tuesday.Length; i++)
                        {
                            if (times_Night_Tuesday[i].TimeID == id && times_Night_Tuesday[i].Meridian == meridian)
                            {
                                return times_Night_Tuesday[i];
                            }
                        }
                        //  Thursday night class
                        for (int i = 0; i < times_Night_Thursday.Length; i++)
                        {
                            if (times_Night_Thursday[i].TimeID == id && times_Night_Thursday[i].Meridian == meridian)
                            {
                                return times_Night_Thursday[i];
                            }
                        }
                    }
                    for (int i = 0; i < times_TR.Length; i++)
                    {
                        if (times_TR[i].TimeID == id && times_TR[i].Meridian == meridian)
                        {
                            return times_TR[i];
                        }
                    }

                }
            }
            else
            {
                // Check for Night Classes first with specific classroom
                if (classDay == "NM" || classDay == "NIGHT_MONDAY")
                {
                    for (int i = 0; i < times_Night_Monday.Length; i++)
                    {
                        if (times_Night_Monday[i].TimeID == id && times_Night_Monday[i].Meridian == meridian)
                        {
                            return times_Night_Monday[i];
                        }
                    }
                }
                else if (classDay == "NT" || classDay == "NIGHT_TUESDAY")
                {
                    for (int i = 0; i < times_Night_Tuesday.Length; i++)
                    {
                        if (times_Night_Tuesday[i].TimeID == id && times_Night_Tuesday[i].Meridian == meridian)
                        {
                            return times_Night_Tuesday[i];
                        }
                    }
                }
                else if (classDay == "NW" || classDay == "NIGHT_WEDNESDAY")
                {
                    for (int i = 0; i < times_Night_Wednesday.Length; i++)
                    {
                        if (times_Night_Wednesday[i].TimeID == id && times_Night_Wednesday[i].Meridian == meridian)
                        {
                            return times_Night_Wednesday[i];
                        }
                    }
                }
                else if (classDay == "NR" || classDay == "NIGHT_THURSDAY")
                {
                    for (int i = 0; i < times_Night_Thursday.Length; i++)
                    {
                        if (times_Night_Thursday[i].TimeID == id && times_Night_Thursday[i].Meridian == meridian)
                        {
                            return times_Night_Thursday[i];
                        }
                    }
                }
                else if (classDay == "NF" || classDay == "NIGHT_FRIDAY")
                {
                    for (int i = 0; i < times_Night_Friday.Length; i++)
                    {
                        if (times_Night_Friday[i].TimeID == id && times_Night_Friday[i].Meridian == meridian)
                        {
                            return times_Night_Friday[i];
                        }
                    }
                }
                // Regular day classes with specific classroom
                else if (classDay == "MWF" || classDay == "M" || classDay == "W" || classDay == "F")
                {

                    for (int i = 0; i < masterTimeslotList.Count; i++)
                    {
                        if (masterTimetableList[i] == 0)
                        {
                            if (masterClassRoomList[i] == classroom)
                            {
                                if (masterTimeslotList[i].TimeID == id && masterTimeslotList[i].Meridian == meridian)
                                {
                                    return masterTimeslotList[i];
                                }
                            }
                        }
                    }
                    for (int i = 0; i < times_Default.Count(); i++)
                    {
                        if (times_Default_Timetable[i] == "MWF" || times_Default_Timetable[i] == "M" || times_Default_Timetable[i] == "W" || times_Default_Timetable[i] == "F")
                        {
                            if (times_Default_Room[i] == classroom)
                            {
                                if (times_Default[i].TimeID == id && times_Default[i].Meridian == meridian)
                                {
                                    return times_Default[i];
                                }
                            }
                        }
                    }
                    for (int i = 0; i < times_MWF.Length; i++)
                    {
                        if (times_MWF[i].TimeID == id && times_MWF[i].Meridian == meridian)
                        {
                            return times_MWF[i];
                        }
                    }
                }
                else
                {

                    for (int i = 0; i < masterTimeslotList.Count; i++)
                    {
                        if (masterTimetableList[i] == 1)
                        {
                            if (masterClassRoomList[i] == classroom)
                            {
                                if (masterTimeslotList[i].TimeID == id && masterTimeslotList[i].Meridian == meridian)
                                {
                                    return masterTimeslotList[i];
                                }
                            }
                        }
                    }
                    for (int i = 0; i < times_Default.Count(); i++)
                    {
                        if (times_Default_Timetable[i] == "TR" || times_Default_Timetable[i] == "T" || times_Default_Timetable[i] == "R")
                        {
                            if (times_Default_Room[i] == classroom)
                            {
                                if (times_Default[i].TimeID == id && times_Default[i].Meridian == meridian)
                                {
                                    return times_Default[i];
                                }
                            }
                        }
                    }
                    for (int i = 0; i < times_TR.Length; i++)
                    {
                        if (times_TR[i].TimeID == id && times_TR[i].Meridian == meridian)
                        {
                            return times_TR[i];
                        }
                    }

                }
            }
            System.Windows.MessageBox.Show("DEBUG: Couldnt find the referenced time!");
            return new Timeslot();
        }
        public ClassRoom DetermineClassroom(string building, int roomNum) // Finds corresponding ClassRoom object based on building name and room number 
        {
            string id = building + roomNum;
            for (int i = 0; i < classrooms.Count; i++)
            {
                if (classrooms[i].ClassID == id)
                {
                    return classrooms[i];
                }
            }
            System.Windows.MessageBox.Show("DEBUG: Couldnt find the referenced classroom!");
            return new ClassRoom();
        }
        public Professors DetermineProfessor(string sruID) // Finds corresponding Professor object based on SRUID 
        {
            for (int i = 0; i < professors.Count; i++)
            {
                if (professors[i].SRUID == sruID)
                {
                    return professors[i];
                }
            }
            System.Windows.MessageBox.Show("DEBUG: Couldn't find the referenced professor!");
            return new Professors();
        }
        public Classes DetermineClass(string classID) // Finds corresponding Class object based on CRN 
        {
            for (int i = 0; i < classList.Count; i++)
            {
                if (classList[i].ClassID == classID)
                {
                    return classList[i];
                }
            }
            System.Windows.MessageBox.Show("DEBUG: Couldnt find the referenced class!");
            return new Classes();
        }
        private T GetParent<T>(DependencyObject d) where T : class
        {
            while (d != null && !(d is T))
            {
                d = VisualTreeHelper.GetParent(d);
            }
            return d as T;
        } // Finds the closest <T> type parent of the passed XAML element
        public bool DetermineTimeConflict(Classes _class, string days, string timeID, string meridian) // Determines if professor is already teaching at that time before he/she is asssigned to a timeslot 
        {
            if (_class.Prof.FirstName == "None")
            {
                return false;
            }
            else
            {
                bool isConflict = false;
                string profName = _class.Prof.FullName;
                string rowID = days + "_" + timeID;
                //MessageBox.Show("Checking against " + rowID + "\nProf: " + profName);
                string labelID = "";
                System.Windows.Controls.Label lbl = null;
                string classID = "";
                for (int i = 0; i < classrooms.Count; i++)
                {
                    labelID = rowID + "_" + classrooms[i].ClassID + classList[i].StartTime.Meridian;
                    lbl = (System.Windows.Controls.Label)FindName(labelID);
                    if (lbl != null)
                    {
                        if (lbl.Tag != null)
                        {
                            classID = lbl.Tag.ToString();
                            for (int n = 0; n < classList.Count; n++)
                            {
                                if (classList[n].ClassID == classID)
                                {
                                    if (classList[n].Prof.FullName == profName && _class.ClassID != classID)
                                    {
                                        isConflict = true;
                                        break;
                                    }
                                }
                            }
                        }
                    }
                    else
                    {
                        //System.Windows.MessageBox.Show("Label " + labelID + " wasnt found!");
                    }
                    if (isConflict)
                    {
                        break;
                    }
                }
                // go through classlist and see if the professor is assigned at the same time
                
                Timeslot targetTime = DetermineTime(timeID, days, _class.Classroom.RoomNum, meridian);
                for (int i = 0; i < classList.Count; i++)
                {
                    if (_class.ClassID != classList[i].ClassID)
                    {
                        if (classList[i].Prof.FullName == _class.Prof.FullName)
                        {
                            //MessageBox.Show("Prof Hit: " + _class.Prof.LastName);
                            if (classList[i].StartTime.FullTime == targetTime.FullTime)
                            {
                                if (classList[i].StartTime.Meridian == targetTime.Meridian)
                                {
                                    if (_class.ClassDay == "MWF" && classList[i].ClassDay == "MWF" || _class.ClassDay == "M" && classList[i].ClassDay == "MWF" || _class.ClassDay == "W" && classList[i].ClassDay == "MWF" || _class.ClassDay == "F" && classList[i].ClassDay == "MWF" || classList[i].ClassDay == "M" && _class.ClassDay == "MWF" || classList[i].ClassDay == "W" && _class.ClassDay == "MWF" || classList[i].ClassDay == "F" && _class.ClassDay == "MWF" || classList[i].ClassDay == "M" && _class.ClassDay == "M" || classList[i].ClassDay == "W" && _class.ClassDay == "W" || classList[i].ClassDay == "F" && _class.ClassDay == "F" || classList[i].ClassDay == "TR" && _class.ClassDay == "TR" || classList[i].ClassDay == "TR" && _class.ClassDay == "T" || classList[i].ClassDay == "TR" && _class.ClassDay == "R" || classList[i].ClassDay == "T" && _class.ClassDay == "TR" || classList[i].ClassDay == "R" && _class.ClassDay == "TR" || classList[i].ClassDay == "T" && _class.ClassDay == "T" || classList[i].ClassDay == "R" && _class.ClassDay == "R")
                                    {
                                        isConflict = true;
                                        if (classList[i].Online)
                                        {
                                            System.Windows.MessageBox.Show("Professor is teaching an ONLINE class at that time...");
                                        }
                                        break;
                                    }
                                }
                            }
                        }
                        if (_class.StartTime.TimeID == timeID)
                        {
                            if (_class.ClassDay == "MWF" && classList[i].ClassDay == "MWF" || _class.ClassDay == "M" && classList[i].ClassDay == "MWF" || _class.ClassDay == "W" && classList[i].ClassDay == "MWF" || _class.ClassDay == "F" && classList[i].ClassDay == "MWF" || classList[i].ClassDay == "M" && _class.ClassDay == "MWF" || classList[i].ClassDay == "W" && _class.ClassDay == "MWF" || classList[i].ClassDay == "F" && _class.ClassDay == "MWF" || classList[i].ClassDay == "M" && _class.ClassDay == "M" || classList[i].ClassDay == "W" && _class.ClassDay == "W" || classList[i].ClassDay == "F" && _class.ClassDay == "F" || classList[i].ClassDay == "TR" && _class.ClassDay == "TR" || classList[i].ClassDay == "TR" && _class.ClassDay == "T" || classList[i].ClassDay == "TR" && _class.ClassDay == "R" || classList[i].ClassDay == "T" && _class.ClassDay == "TR" || classList[i].ClassDay == "R" && _class.ClassDay == "TR" || classList[i].ClassDay == "T" && _class.ClassDay == "T" || classList[i].ClassDay == "R" && _class.ClassDay == "R")
                            {
                                if (_class.Classroom.ClassID == classList[i].Classroom.ClassID)
                                {
                                    try
                                    {
                                        string compareTime = classList[i].StartTime.FullTime; //compares entirety of class times to determine any conflicts based on when a class starts and ends
                                        string classTime = targetTime.FullTime;

                                        string[] compareTimeSplit = compareTime.Split(new char[] { '-' });
                                        string[] classTimeSplit = classTime.Split(new char[] { '-' });

                                        string compareTimeStarting = compareTimeSplit[0];
                                        string compareTimeEnding = compareTimeSplit[1];
                                        compareTimeStarting = compareTimeStarting.Trim();
                                        compareTimeEnding = compareTimeEnding.Trim();


                                        string classTimeStarting = classTimeSplit[0];
                                        string classTimeEnding = classTimeSplit[1];
                                        classTimeStarting = classTimeStarting.Trim();
                                        classTimeEnding = classTimeEnding.Trim();

                                        string compareMeridian = classList[i].StartTime.Meridian;
                                        string classMeridian = targetTime.Meridian;

                                        if (compareTimeStarting != "" && compareTimeEnding != "" && classTimeStarting != "" && classTimeEnding != "")
                                        {
                                            string[] compareTimeStartingSplit = compareTimeStarting.Split(new char[] { ':' }); //compare time starting time
                                            string compareTimeStartFront = compareTimeStartingSplit[0];
                                            string compareTimeStartBack = compareTimeStartingSplit[1];

                                            string[] compareTimeEndingSplit = compareTimeEnding.Split(new char[] { ':' }); //compare time ending time
                                            string compareTimeEndFront = compareTimeEndingSplit[0];
                                            string compareTimeEndBack = compareTimeEndingSplit[1];

                                            string[] classTimeStartingSplit = classTimeStarting.Split(new char[] { ':' }); // target time starting time
                                            string classTimeStartFront = classTimeStartingSplit[0];
                                            string classTimeStartBack = classTimeStartingSplit[1];

                                            string[] classTimeEndingSplit = classTimeEnding.Split(new char[] { ':' }); //target time ending time
                                            string classTimeEndFront = classTimeEndingSplit[0];
                                            string classTimeEndBack = classTimeEndingSplit[1];
                                            if (classMeridian == classList[i].StartTime.Meridian || Int32.Parse(compareTimeStartFront) == 11 && Int32.Parse(compareTimeEndFront) == 12 || Int32.Parse(classTimeStartFront) == 11 && Int32.Parse(classTimeEndFront) == 12)
                                            {
                                                for (int q = Int32.Parse(compareTimeStartFront); q <= Int32.Parse(compareTimeEndFront); q++)
                                                {
                                                    if (q == Int32.Parse(classTimeStartFront))
                                                    {
                                                        if (q == Int32.Parse(compareTimeEndFront))
                                                        {
                                                            if (Int32.Parse(compareTimeEndBack) >= Int32.Parse(classTimeStartBack))
                                                            {
                                                                isConflict = true;
                                                                System.Windows.MessageBox.Show("Time Conflict found between " + _class.TextBoxName + " and " + classList[i].TextBoxName + " Moving class to unassigned class List");
                                                                break;
                                                            }
                                                        }
                                                    }
                                                    else if (q == Int32.Parse(classTimeEndFront))
                                                    {
                                                        if (q == Int32.Parse(compareTimeStartFront))
                                                        {
                                                            if (Int32.Parse(classTimeEndBack) >= Int32.Parse(compareTimeStartBack))
                                                            {
                                                                isConflict = true;
                                                                System.Windows.MessageBox.Show("Time Conflict found between " + _class.TextBoxName + " and " + classList[i].TextBoxName + " Moving class to unassigned class List");
                                                                break;
                                                            }
                                                        }
                                                    }
                                                }
                                                if (isConflict == true)
                                                {
                                                    break;
                                                }
                                            }
                                        }
                                    }

                                    catch (Exception ex)
                                    {

                                    }
                                }
                            }
                        }
                    }
                }
                return isConflict;
            }
        }
        public string formatTime(string time) // Standardizes time format being read from excel file to prevent errors when creating the classes 
        {
            string formattedTime = "";
            if (time.Contains(":"))
            {
                string left = time.Split(':')[0];
                string right = time.Split(':')[1];
                if (left.Length == 1)
                {
                    left = "0" + left;
                }
                formattedTime = left + ":" + right;
            }
            else
            {
                if (time.Length == 3)
                {
                    time = "0" + time;
                }
                Console.WriteLine(time);
                string left = time.Substring(0, 2);
                string right = time.Substring(2, 2);
                formattedTime = left + ":" + right;
            }
            return formattedTime;
        }
        public string getFileDirectory(string filePath) // Extracts directory string from full filepath string 
        {
            string directory = "";
            for (int i = (filePath.Length - 1); i >= 0; i--)
            {
                if (filePath[i] == '\\')
                {
                    directory = filePath.Substring(0, (i + 1));
                    break;
                }
            }
            //MessageBox.Show("Directory: " + directory);
            return directory;
        }
        public System.Data.DataTable getDataTableFromClasses() // Creates a datatable based on classList 
        {
            //Creating DataTable  
            System.Data.DataTable dt = new System.Data.DataTable();
            //Setiing Table Name  
            dt.TableName = "Sheet 1";
            // Determine Types

            for (int i = 0; i < classList.Count(); i++)
            {
                for (int j = 0; j < i; j++)
                {
                    if (classList[i].DeptName != classList[j].DeptName)
                    {
                        if (classList[j].DeptName == "CYBR" && classList[i].DeptName == "CPSC")
                        {
                            Classes tmp = classList[i];
                            classList[i] = classList[j];
                            classList[j] = tmp;
                        }
                    }
                    else if (classList[i].ClassNumber != classList[j].ClassNumber)
                    {
                        if (classList[j].ClassNumber > classList[i].ClassNumber)
                        {
                            Classes tmp = classList[i];
                            classList[i] = classList[j];
                            classList[j] = tmp;
                        }
                    }
                    else
                    {
                        if (classList[j].SectionNumber > classList[i].SectionNumber)
                        {
                            Classes tmp = classList[i];
                            classList[i] = classList[j];
                            classList[j] = tmp;
                        }
                    }
                }
            }

            for (int i = 0; i < excelHeaders.Count; i++)
            {
                Type colType = typeof(string);
                if (i == 3 || i == 8 || i == 12 || i == 18 || i == 20)
                {
                    colType = typeof(int);
                }
                excelTypes.Add(colType);
            }
            //Add Columns
            for (int i = 0; i < excelHeaders.Count; i++)
            {
                dt.Columns.Add(excelHeaders[i], excelTypes[i]);
            }
            //Add Rows in DataTable
            for (int i = 0; i < classList.Count; i++)
            {
                if (classList[i].ExtraData.Count < 15)
                {
                    
                    int extraFields = 15; // number of extra fields in classes
                    for (int n = classList[i].ExtraData.Count; n < extraFields; n++)
                    {
                        classList[i].ExtraData.Add("");
                    }
                }
                string start = classList[i].StartTime.Start;
                string end = classList[i].StartTime.End;

                if (start != "-- ")
                {
                    string checkMeridian = start.Substring(0, 2);
                    string checkMeridianEnd = end.Substring(0, 2);
                    int starting = Int32.Parse(checkMeridian);
                    int ending = Int32.Parse(checkMeridianEnd);

                    if (starting < 12 && ending >= 12)
                    {
                        end = end.Substring(0, end.Length - 2);
                        end = end + "PM";
                    }
                }
                if (start == "-- ")
                {
                    start = "";
                    end = "";
                }

                string method;

                if (classList[i].Building == "WEB")
                {
                    method = "Online Only: 100 Pct OL";
                }
                else if(classList[i].Building == "APPT2")
                {
                    method = "Internship";
                }
                else if(classList[i].Building == "APPT")
                {
                    method = "Unknown"; //if further clarification is needed, consult the department chair
                }
                else
                {
                    method = "Traditional";
                }


                dt.Rows.Add(classList[i].Term, classList[i].Session, classList[i].DeptName, classList[i].ClassNumber,
                      classList[i].getSectionString(), classList[i].CRN, classList[i].ClassName, classList[i].Crosslist, classList[i].Credits,
                      classList[i].MaxSeats, classList[i].Waitlist, classList[i].ProjSeats, classList[i].SeatsTaken,
                      classList[i].StartDate, classList[i].EndDate, classList[i].ClassDay, start, end, classList[i].Classroom.AvailableSeats,
                      classList[i].Classroom.Location, classList[i].Classroom.RoomNum, classList[i].Prof.FullName, classList[i].Prof.SRUID, classList[i].ExtraData[0],
                       method, classList[i].ExtraData[1], classList[i].ExtraData[2], classList[i].ExtraData[3], classList[i].ExtraData[4], classList[i].ExtraData[5], classList[i].ExtraData[6], classList[i].SectionNotes, classList[i].Notes);
             
            }
            //Add Deleted classes to DataTable
            for (int i = 0; i < deletedClasses.Count; i++)
            {
                if (deletedClasses[i].ExtraData.Count == 0)
                {
                    int extraFields = 15; // number of extra fields in classes
                    for (int n = 0; n < extraFields; n++)
                    {
                        deletedClasses[i].ExtraData.Add("");
                    }
                }
                else
                {
                    int extraFields = 15; // number of extra fields in classes
                    for (int n = deletedClasses[i].ExtraData.Count + 1; n < extraFields; n++)
                    {
                        deletedClasses[i].ExtraData.Add("");
                    }
                }
                string start = deletedClasses[i].StartTime.Start;
                string end = deletedClasses[i].StartTime.End;
                if (start == "-- ")
                {
                    start = "";
                    end = "";
                }
                dt.Rows.Add(termYear + term, deletedClasses[i].Session, deletedClasses[i].DeptName, deletedClasses[i].ClassNumber,
                    deletedClasses[i].getSectionString(), deletedClasses[i].CRN, deletedClasses[i].ClassName, deletedClasses[i].ExtraData[0], deletedClasses[i].Credits,
                    deletedClasses[i].MaxSeats, deletedClasses[i].ExtraData[2], deletedClasses[i].ProjSeats, deletedClasses[i].SeatsTaken,
                    deletedClasses[i].ExtraData[4], deletedClasses[i].ExtraData[5], deletedClasses[i].ClassDay, start, end, deletedClasses[i].Classroom.AvailableSeats,
                    deletedClasses[i].Classroom.Location, deletedClasses[i].Classroom.RoomNum, deletedClasses[i].Prof.FullName, deletedClasses[i].Prof.SRUID,
                    deletedClasses[i].ExtraData[6], deletedClasses[i].ExtraData[7], deletedClasses[i].ExtraData[8], deletedClasses[i].ExtraData[9],
                    deletedClasses[i].ExtraData[10], deletedClasses[i].ExtraData[11], deletedClasses[i].ExtraData[12], deletedClasses[i].ExtraData[13],
                    deletedClasses[i].SectionNotes, deletedClasses[i].Notes);
            }
            dt.AcceptChanges();
            return dt;
        }
        public string ComputeSha256Hash(byte[] rawData) // Compute the SHA256 hash digest of the passed byte buffer. Then convert it to string format. 
        {
            // Create a SHA256   
            using (SHA256 sha256Hash = SHA256.Create())
            {
                // ComputeHash - returns byte array  
                byte[] bytes = sha256Hash.ComputeHash(rawData);

                // Convert byte array to a string   
                StringBuilder builder = new StringBuilder();
                for (int i = 0; i < bytes.Length; i++)
                {
                    builder.Append(bytes[i].ToString("x2"));
                }
                return builder.ToString();
            }
        }
        public void GenerateClassListHashes() // Generate initial hashes for class list read from excel file (for comparison when writing to new file)
        {
            string hash;
            for (int i = 0; i < classList.Count; i++)
            {
                hash = ComputeSha256Hash(classList[i].Serialize());
                hashedClasses.Add(new ClassesHash(classList[i].ClassID, hash));
            }
        }
        public void ScrollViewer_PreviewMouseWheel(object sender, MouseWheelEventArgs e) // Set the scrolling speed for the lists using mousewheel 
        {
            ScrollViewer scv = (ScrollViewer)sender;
            scv.ScrollToVerticalOffset(scv.VerticalOffset - e.Delta / 10);
            e.Handled = true;
        }
        public bool isGroupMember(Classes _class, int groupIndex)
        {
            for (int i = 0; i < classGroupings[groupIndex].ClassDept.Count; i++)
            {
                if (_class.DeptName == classGroupings[groupIndex].ClassDept[i] && _class.ClassNumber == Int32.Parse(classGroupings[groupIndex].ClassNum[i]))
                {
                    return true;
                }
            }
            return false;
        }
        public void checkGroupConflict(Classes _class, int groupIndex)
        {
            for (int i = 0; i < classList.Count; i++)
            {
                if (classList[i].DeptName != _class.DeptName || classList[i].ClassNumber != _class.ClassNumber)
                {
                    if (isGroupMember(classList[i], groupIndex))
                    {
                        if (classList[i].StartTime.FullTime == _class.StartTime.FullTime && classList[i].StartTime.Time != "--")
                        {
                            // check if there is another section at another time
                            if (!hasAlternativeSection(_class) && !hasAlternativeSection(classList[i]))
                            {
                                // Log conflict
                                soft_constraint_log.Text = soft_constraint_log.Text + "\n> " + "[Group Conflict] " + classList[i].DeptName + " " + classList[i].ClassNumber + " section " +
                                    classList[i].SectionNumber + "  <---> " + _class.DeptName + " " + _class.ClassNumber + " section " + _class.SectionNumber;
                                break;
                            }
                        }
                    }
                }
            }
        }
        public bool hasAlternativeSection(Classes _class)
        {
            for (int i = 0; i < classList.Count; i++)
            {
                if (classList[i].DeptName == _class.DeptName && classList[i].ClassNumber == _class.ClassNumber)
                {
                    if (classList[i].StartTime != null && _class.StartTime != null)
                    {
                        if (classList[i].StartTime.FullTime != _class.StartTime.FullTime)
                        {
                            return true;
                        }
                    }
                }
            }
            return false;
        }

        // Custom Brushes
        LinearGradientBrush stripedBackground(Color baseBackground)
        {
            LinearGradientBrush output = new LinearGradientBrush();

            output.MappingMode = BrushMappingMode.Absolute;
            output.SpreadMethod = GradientSpreadMethod.Repeat;
            output.StartPoint = new System.Windows.Point(0, 0);
            output.EndPoint = new System.Windows.Point(8, 8);

            output.GradientStops.Add(new GradientStop(baseBackground, 0.0));
            output.GradientStops.Add(new GradientStop(baseBackground, 0.5));
            output.GradientStops.Add(new GradientStop(Colors.LightGray, 0.5));
            output.GradientStops.Add(new GradientStop(Colors.LightGray, 1.0));
            return output;
        }
    }

    // XAML converters
    public class ColorConverter : IValueConverter
    {

        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {

            string input = value.ToString();

            // split string into numbers
            string[] nums = input.Split(new string[] { " / " }, StringSplitOptions.None);

            //custom condition is checked based on data.
            int current = Int32.Parse(nums[0].ToString());
            int max = Int32.Parse(nums[1].ToString());

            if (current > max)
            {
                //return "SeaGreen";
                return new SolidColorBrush(Colors.LightPink);
            }
            else if (current <= max)
                //return "LightGreen";
                return new SolidColorBrush(Colors.LightGreen);
            else
                return DependencyProperty.UnsetValue;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
    public class PreferenceConverter : IValueConverter
    {

        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            int input = Int32.Parse(value.ToString());

            if (input == -1)
            {
                return new SolidColorBrush(Colors.PaleGoldenrod);
            }
            else if (input == -2)
                return new SolidColorBrush(Colors.Pink);
            else
                return new SolidColorBrush(Colors.White);
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
    public class PreferenceMessageConverter : IValueConverter
    {

        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {

            string input = value.ToString();

            if (input != "")
            {
                return input;
            }
            else
                return DependencyProperty.UnsetValue;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }

}