using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
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
    /// Interaction logic for AddRowDialog.xaml
    /// </summary>
    public partial class AddRowDialog : Window
    {
        public ObservableCollection<ComboBoxItem> numTimeslots { get; set; }
        public ComboBoxItem selectedComboBoxItem { get; set; }
        
        public AddRowDialog()
        {
            InitializeComponent();
            DataContext = this;

            int min = Int32.Parse(System.Windows.Application.Current.Resources["Set_min"].ToString());
            

            numTimeslots = new ObservableCollection<ComboBoxItem>();
            

            for (int i = min; i <= 24; i++)
            {
                string cont = i.ToString();
                if (i == min)
                {
                    var itemOne = new ComboBoxItem { Content = cont };
                    selectedComboBoxItem = itemOne;
                    numTimeslots.Add(itemOne);
                }
                else
                {
                    var item = new ComboBoxItem { Content = cont };
                    numTimeslots.Add(item);
                }
            }

            
        }

        

        private void Submit_Click(object sender, RoutedEventArgs e)
        {
            Application.Current.Resources["Set_ChangeTimeslots_Success"] = false;

            if (allRequiredFields())
            {
                int checkMWF = Int32.Parse(System.Windows.Application.Current.Resources["Set_min"].ToString());

                Application.Current.Resources["Set_ChangeTimeslots_Success"] = true;
                int rows = RowNum.SelectedIndex;
                rows = rows + checkMWF;
                int timeTableNum = TimeTable.SelectedIndex;

                Application.Current.Resources["Set_rows"] = rows;
                Application.Current.Resources["Set_TimeTable"] = timeTableNum;
                this.Close();
            }

        }

        private bool allRequiredFields()
        {
            int checkMWF = Int32.Parse(System.Windows.Application.Current.Resources["Set_min"].ToString());
            int rows = RowNum.SelectedIndex;

            bool success = true;

            if (TimeTable.SelectedIndex == 0)
            {
                rows = rows + checkMWF;
                if (rows < checkMWF)
                {
                    success = false;
                }
            }
            if (TimeTable.SelectedIndex == 1)
            {
                rows = rows + checkMWF;
                if (rows < checkMWF)
                {
                    success = false;
                }
            }


            return success;
        }
    }
}
