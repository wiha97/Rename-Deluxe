using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using MsgReader.Outlook;

namespace RenameDeluxe
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private string TxtIN { get { return txtIN.Text; } set { txtIN.Text = value; } }
        private string TxtOU { get { return txtOU.Text; } set { txtOU.Text = value; } }
        private DateTime? Date = DateTime.Now;
        private string sPath;
        private string tPath;
        private string fName;
        private List<string> fList = new List<string>();
        private string[] ArrFiles;
        private int Id = 0;
        //public string Name;
        //private DateTime euDate;
        //private string Chars = @"/?*\(){}#$@%";
        //private Regex IllChars = new Regex(@"/?*\(){}#$@%");

        #region Magic
        //TODO: Move outlook stuff to another method
        public MainWindow()
        {
            InitializeComponent();
            sPath = TxtIN;
            tPath = TxtOU;
            GetFiles();
        }

        public void DirCreate()
        {
            if (!Directory.Exists(sPath))
            {
                Directory.CreateDirectory(sPath);
            }
            if (!Directory.Exists(tPath))
            {
                Directory.CreateDirectory(tPath);
            }
        }

        /// <summary>
        /// Adds a date to the file name
        /// </summary>
        public void AddDate()
        {
            DirCreate();
            string NewName;
            lstITM.Items.Clear();
            ArrFiles = Directory.GetFiles(sPath);
            var LFiles = ArrFiles.ToList();
            foreach (string item in LFiles)
            {
                //Reads the email so it can get the proper subject and date
                Storage.Message message = new Storage.Message(item);
                string date = message.ReceivedOn.ToString();
                Name = message.Subject;

                //Renames the file so it won't have any special characters
                //DateTime euDate = DateTime.ParseExact(date, "HH:mm", CultureInfo.InvariantCulture);
                DateTime euDate = DateTime.Parse(date);
                fName = Path.GetFileName(item);
                Regex illegalInFileName = new Regex(@"[\\/;/:*?)!#//¤%/$!&•(""<>|]");
                string myString = illegalInFileName.Replace(euDate.ToString("HH:\\_mm_") + Name, "");
                NewName = myString + Name;
                //DelOld();
                try
                {
                    //Copies the files over to a new folder as it won't let you change anything when it has just been "opened"
                    File.Copy(sPath + "\\" + fName, tPath + "\\" + myString + ".msg");
                }
                catch (Exception e)
                {
                    Console.WriteLine(e);
                }
                lstITM.Items.Add(new Item() { ID = Id, Name = NewName, Date = Date, Message = date });
            }
        }

        /// <summary>
        /// Deletes the original files
        /// </summary>
        public void Delete()
        {
            var LFiles = ArrFiles.ToList();
            foreach (string item in LFiles)
            {
                try
                {
                    File.Delete(sPath + "\\" + Name);
                }
                catch (Exception e)
                {

                }
            }
        }

        public void GetFiles()
        {
            if (!Directory.Exists(sPath))
            {
                Directory.CreateDirectory(sPath);
            }
            OpenFolder();
            ArrFiles = Directory.GetFiles(sPath);
            var LFiles = ArrFiles.ToList();
            foreach (string item in LFiles)
            {
                string Name;
                fList.Add(item);

                Id++;
                Name = Path.GetFileName(item);
                //Regex illegalInFileName = new Regex(@"[\\/:*?)(""<>|]");
                //string myString = illegalInFileName.Replace(Name, "");
                Storage.Message message = new Storage.Message(sPath + "\\" + Name);
                lstITM.Items.Add(new Item() { ID = Id, Name = Name, Date = Date, Message = message.ReceivedOn.ToString() });
            }
        }

        public void Rename()
        {
            lstITM.Items.Clear();
            Id = 0;
            var LFiles = ArrFiles.ToList();
            foreach (string item in LFiles)
            {
                string Name;
                fList.Add(item);

                Id++;
                Name = Path.GetFileName(item);
                Regex illegalInFileName = new Regex(@"[\\/:*?)!#¤%$!&•(""+\-<>|]");
                string myString = illegalInFileName.Replace(Name, "_");
                try
                {
                    File.Move(sPath + "\\" + Name, sPath + "\\" + myString);
                }
                catch (Exception e)
                {
                    //File.CreateText(fPath + "\\oops.txt");
                }
                lstITM.Items.Add(new Item() { ID = Id, Name = myString, Date = Date });
            }
        }

        public void OpenFolder()
        {
            if (Directory.Exists(sPath))
            {
                Process.Start(sPath);
            }
        }

        public void ReadMail(Storage.Message message)
        {
            var LFiles = ArrFiles.ToList();
            foreach (string item in LFiles)
            {

            }
        }
        #endregion
        #region Buttons
        private void btnRN_Click(object sender, RoutedEventArgs e)
        {
            Rename();
        }

        private void btnDEL_Click(object sender, RoutedEventArgs e)
        {
            Delete();
        }

        private void btnDT_Click(object sender, RoutedEventArgs e)
        {
            AddDate();
        }

        private void btnRF_Click(object sender, RoutedEventArgs e)
        {
            sPath = TxtIN;
            tPath = TxtOU;
            lstITM.Items.Clear();
            OpenFolder();
            GetFiles();
        }
        #endregion
    }

    public class Item
    {
        public int ID { get; set; }
        public string Name { get; set; }
        public DateTime? Date { get; set; }
        public string Message { get; set; }
    }
}
