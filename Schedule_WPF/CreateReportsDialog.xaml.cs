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
    public partial class CreateReportsDialog : Window
    {
        //ProfessorList professors = (ProfessorList)Application.Current.FindResource("Professor_List_View");

        public CreateReportsDialog()
        {
            InitializeComponent();
            System.Windows.Application.Current.Resources["Set_Report_Success"] = false;
        }

        

        private void Button_Click(object sender, RoutedEventArgs e)
        {

            int reportType = ReportType.SelectedIndex;

            Application.Current.Resources["Report_Type"] = reportType;

            System.Windows.Application.Current.Resources["Set_Report_Success"] = true;
            Close();
            
        }


        private void ComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

        }
    }
}
