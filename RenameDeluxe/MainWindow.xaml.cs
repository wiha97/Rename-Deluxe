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
        private string TxtID { get { return txtID.Text; } set { txtID.Text = value; } }
        private DateTime? Date = DateTime.Now;
        private string sPath;
        private string tPath;
        private string sign;
        private string fName;
        private List<string> fList = new List<string>();
        private string[] ArrFiles;
        private int Id = 0;
        private string caption;
        private string messageBoxText;
        private string aName;
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
            sign = TxtID;
            GetFiles();
        }

        public void CheckDate(string attchs)
        {
            string[] isDate = new string[] { "yyyyMMdd yyyy-MM-dd yyyy:MM:dd yyyy_MM_dd" };
            DateTime dateTime = new DateTime();

            if (DateTime.TryParseExact(attchs, isDate,CultureInfo.InvariantCulture, DateTimeStyles.AssumeLocal, out dateTime))
            {
                aName = "test";
            }
        }

        public void ReadAtt(string item)
        {
            Storage.Message message = new Storage.Message(item);
            string attch = message.GetAttachmentNames();
            //CheckDate(attchs);
            aName = attch;
        }

        public void ErrMsg(Exception e)
        {
            caption = "Something went wrong";
            messageBoxText = e.ToString();
            MessageBoxButton button = MessageBoxButton.OK;
            MessageBoxImage icon = MessageBoxImage.Error;
            MessageBoxResult result = MessageBox.Show(messageBoxText, caption, button, icon);
        }

        public void DirCreate()
        {
            Directory.CreateDirectory(sPath);
            Directory.CreateDirectory(tPath);
        }

        /// <summary>
        /// Adds a date to the file name
        /// </summary>
        public void Rename()
        {
            DirCreate();
            string Name;
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
                if (Name.Length < 1)
                {
                    ReadAtt(item);
                    Name = aName;
                }

                //Renames the file so it won't have any special characters
                //DateTime euDate = DateTime.ParseExact(date, "HH:mm", CultureInfo.InvariantCulture);
                DateTime euDate = DateTime.Parse(date);
                fName = Path.GetFileName(item);
                Regex illegalInFileName = new Regex(@"[\\/;/:*?)!#//¤%/$!,.&•(""<>|]");
                string myString = illegalInFileName.Replace(euDate.ToString("HH:\\_mm_") + Name, "");
                NewName = myString + Name;
                string source = sPath + "\\" + fName;
                string target = tPath + "\\" + myString;
                try
                {
                    File.Copy(source, target + ".msg");
                }
                catch (Exception e)
                {
                    ErrMsg(e);
                }
                lstITM.Items.Add(new Item() { ID = Id, Name = NewName, Date = Date, Message = date });
                Process.Start(tPath);
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
                    File.Delete(item);
                }
                catch (Exception e)
                {
                    ErrMsg(e);
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

        //public void Rename()
        //{
        //    lstITM.Items.Clear();
        //    Id = 0;
        //    var LFiles = ArrFiles.ToList();
        //    foreach (string item in LFiles)
        //    {
        //        string Name;
        //        fList.Add(item);

        //        Id++;
        //        Name = Path.GetFileName(item);
        //        Regex illegalInFileName = new Regex(@"[\\/:*?)!#¤%$!&•(""+\-<>|]");
        //        string myString = illegalInFileName.Replace(Name, "_");
        //        try
        //        {
        //            File.Move(sPath + "\\" + Name, sPath + "\\" + myString);
        //        }
        //        catch (Exception e)
        //        {
        //            //File.CreateText(fPath + "\\oops.txt");
        //        }
        //        lstITM.Items.Add(new Item() { ID = Id, Name = myString, Date = Date });
        //    }
        //}

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

        private void btnDel_Click(object sender, RoutedEventArgs e)
        {
            Delete();
        }

        private void btnICN_Click(object sender, RoutedEventArgs e)
        {

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
