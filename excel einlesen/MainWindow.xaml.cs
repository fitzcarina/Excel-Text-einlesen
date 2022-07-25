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
using System.Windows.Navigation;
using System.Windows.Shapes;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Word;
using _Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
using System.IO;
using Microsoft.Win32;

using _Application = Microsoft.Office.Interop.Excel._Application;


namespace excel_einlesen
    
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow 
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void Excel(object sender, RoutedEventArgs e)
        {
            OpenFile();
        }
        private void Textdatei(object sender, RoutedEventArgs e)
        {
            StreamReader sr = new StreamReader(@"c:\test\test1.txt");
            MessageBox.Show(sr.ReadToEnd());
        }
        //public void Word(object sender, RoutedEventArgs e)
        //{

        //    OpenFileDialog fileDialog = new OpenFileDialog();
        //    fileDialog.ShowDialog();
        //    string filepath = fileDialog.FileName.ToString();
      
        //    Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application();
            


        //}
        public void OpenFile()
        {
            Excel excel = new Excel(@"c:\test\test1.xlsx", 1);
            MessageBox.Show(excel.ReadCell(0, 0));
        }

        //private void FindAndReplace(Word.Application wordApp, object ToFindText, object replaceWithText)
        //{

        //}


    }

    class Word
    {
        
    }
    class Excel
    {
        string path = "";
        _Application excel = new _Excel.Application();
        Workbook wb;
        Worksheet ws;


       public Excel (string path, int sheet)
        {
            this.path = path;
            wb = excel.Workbooks.Open(path);
            ws = wb.Worksheets[sheet];
        }
           
        public string  ReadCell ( int i , int j)
        {
            j++;
            i++;
            if (ws.Cells[i, j].Value2 != null)
            {
                return ws.Cells[i, j].Value2;

            }
            else
                return "";
        }
          
        
    }
}
