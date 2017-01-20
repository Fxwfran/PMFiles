using System;
using System.Collections.Generic;
using System.Collections;
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
using System.Windows.Navigation;
using System.Windows.Shapes;
using MahApps.Metro.Controls;
using System.IO;
using System.Windows.Forms;
using OfficeOpenXml;
using MahApps.Metro.IconPacks;


namespace Excel_PM_helper
{
    using System.Windows.Forms;
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : MetroWindow
    {
        public MainWindow()
        {
            InitializeComponent();
        }
        private static OpenFileDialog openFileDialog1 = new OpenFileDialog();
        private static DialogResult returned;
        private static ExcelCellAddress start;
        private static ExcelCellAddress end;
        private static ExcelWorksheet mainsheet;
        private static float hours;
        public void Button_Click(object sender, RoutedEventArgs e)
        {
            openFileDialog1.Filter = "PM Worksheets|*.xlsx";
            openFileDialog1.Title = "Select a table to load";
            returned = openFileDialog1.ShowDialog();
            if(returned == System.Windows.Forms.DialogResult.OK)
            {
                Name_list.IsEnabled = true;
            FileInfo sr = new FileInfo(openFileDialog1.FileName);
                filename.Text = openFileDialog1.SafeFileName;
            List<string> listita = new List<string> { };
            ExcelPackage pck = new ExcelPackage(sr);
            foreach (ExcelWorksheet sheet in pck.Workbook.Worksheets)
            {
                for(int row = sheet.Dimension.Start.Row+1; row < sheet.Dimension.End.Row; row++)
                {
                    if(!(listita.Contains(sheet.Cells[row, 2].Text)))
                    {
                    listita.Add(sheet.Cells[row, 2].Text);
                    }
                }
                    mainsheet = sheet;
            }
                listita.Sort();
                Name_list.ItemsSource = listita;
            }

            return;
        }

        public void Button_Click_1(object sender, RoutedEventArgs e)
        {
            if (returned == System.Windows.Forms.DialogResult.OK)
            {
                FileInfo sr = new FileInfo(openFileDialog1.FileName);
                ExcelPackage pck = new ExcelPackage(sr);
                foreach (ExcelWorksheet sheet in pck.Workbook.Worksheets)
                {
                    start = sheet.Dimension.Start;
                    end = sheet.Dimension.End;
                    var val = sheet.Column(1);
                    float hours = 0;
                    calendar.FirstDayOfWeek = DayOfWeek.Monday;
                    calendar.SelectionMode = CalendarSelectionMode.MultipleRange;
                }
            }
        }

        private void calendar_selectionchanged(object sender, SelectionChangedEventArgs e)
        {
            for (int row = start.Row; row <= end.Row; row++)
            {
                string cellValue = mainsheet.Cells[row, 2].Text;
                if (cellValue == Name_list.Text)
                {
                    string uglydate = mainsheet.Cells[row, 3].Text;
                    double d = double.Parse(uglydate);
                    DateTime conv = DateTime.FromOADate(d);
                    //MessageBox.Show(conv.ToString());
                    //MessageBox.Show();
                    if (calendar.SelectedDates.Contains(conv))
                    {
                        hours += float.Parse(mainsheet.Cells[row, 6].Text);
                    }
                }
            }
            hoursworked.Content = hours.ToString();
            hours = 0;
        }
    }
}
