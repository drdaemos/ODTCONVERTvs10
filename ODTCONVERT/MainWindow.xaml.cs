﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Markup;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Windows.Threading;
using System.Xml;


namespace ODTCONVERT
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Elysium.Controls.Window
    {
        private int docinc = 1;
        private RecentHolder recents;
        public MainWindow()
        {
            InitializeComponent();
            sidebar.Width = 0;
            SelectWindowState(State.Start);
            PreviewMouseWheel += Zoom_MouseWheel;
            recents = new RecentHolder();
            lbRecents.ItemsSource = recents.recents;
        }

        private void Zoom_MouseWheel(object sender, MouseWheelEventArgs e)
        {
            bool handle = (Keyboard.Modifiers & ModifierKeys.Control) > 0;
            if (!handle)
                return;

            if (e.Delta > 0) scale.Value += 0.1; else scale.Value -= 0.1;
        }

        enum Message { WrongFormat, Unknown, NotFound };
        enum State { Text, Start, Error, Settings};
        
        private void btnRead_Click(object sender, RoutedEventArgs e)
        {
            OpenDialog();
        }        

        private void OpenDialog()
        {
            // Create OpenFileDialog 
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
            // Set filter for file extension and default file extension 
            dlg.DefaultExt = ".odt";
            dlg.Filter = "OpenDocument Text (*.odt)|*.odt";
            // Get the selected file name and display in a TextBox 
            if (dlg.ShowDialog() == true)
            {
                lbPath.Text = dlg.FileName;

                OpenDoc(dlg.FileName);
                recents.Update(new Recent() { Uri = dlg.FileName, Name = dlg.SafeFileName, Date = DateTime.Now});
            }
        }

        private void OpenDoc(string path)
        {
            pbParse.State = Elysium.Controls.ProgressState.Busy;
            lbText.Dispatcher.UnhandledException += Dispatcher_UnhandledException;
            System.Threading.Thread thread = new System.Threading.Thread(
            new System.Threading.ThreadStart(
               delegate()
               {
                   System.Windows.Threading.DispatcherOperation
                     dispatcherOp = lbText.Dispatcher.BeginInvoke(
                     System.Windows.Threading.DispatcherPriority.Normal,
                     new Action(
                       delegate()
                       {
                           lbText.Document = ParseODT(path);
                           SelectWindowState(State.Text);                           
                       }
                   ));
                   dispatcherOp.Completed += new EventHandler(read_Completed);
               }
          ));
            thread.Start();
        }
        private void NewDoc()
        {
             lbPath.Text = System.IO.Path.Combine(Environment.CurrentDirectory, "NewDocument"+Properties.Settings.Default.NewDocuments);                

                lbText.Dispatcher.UnhandledException += Dispatcher_UnhandledException;
                System.Threading.Thread thread = new System.Threading.Thread(
                new System.Threading.ThreadStart(
                   delegate()
                   {
                       System.Windows.Threading.DispatcherOperation
                         dispatcherOp = lbText.Dispatcher.BeginInvoke(
                         System.Windows.Threading.DispatcherPriority.Normal,
                         new Action(
                           delegate()
                           {
                               lbText.Document = new FlowDocument();
                               SelectWindowState(State.Text);
                           }
                       ));
                   }
              ));
              thread.Start();
              docinc++;
        }

        private void SaveDialog()
        {
            // Create OpenFileDialog 
            Microsoft.Win32.SaveFileDialog dlg = new Microsoft.Win32.SaveFileDialog();
            // Set filter for file extension and default file extension 
            dlg.DefaultExt = ".pdf";
            dlg.Filter = "Adobe PDF (*.pdf)|*.pdf";
            // Get the selected file name and display in a TextBox 
            if (dlg.ShowDialog() == true)
            {
                SaveDoc(dlg.FileName);
            }
        }

        private void SaveDoc(string path)
        {
            pbParse.State = Elysium.Controls.ProgressState.Indeterminate;

            lbText.Dispatcher.UnhandledException += Dispatcher_UnhandledException;
            System.Threading.Thread thread = new System.Threading.Thread(
            new System.Threading.ThreadStart(
               delegate()
               {
                   System.Windows.Threading.DispatcherOperation
                     dispatcherOp = btnSave.Dispatcher.BeginInvoke(
                     System.Windows.Threading.DispatcherPriority.Normal,
                     new Action(
                       delegate()
                       {
                           SavePDF(path);
                       }
                   ));
                   dispatcherOp.Completed += new EventHandler(save_Completed);
               }
          ));
            thread.Start();
        }
        private void SelectWindowState(State state)
        {
            switch (state)
            {
                case State.Text:
                    {
                        tabs.SelectedIndex = 1;
                        break;
                    }
                case State.Start:
                case State.Error:
                    {
                        tabs.SelectedIndex = 0;
                        break;
                    }
                case State.Settings:
                    {
                        OpenSettingsBar();
                        break;
                    }
            }
        }

        void Dispatcher_UnhandledException(object sender, DispatcherUnhandledExceptionEventArgs e)
        {
            SelectWindowState(State.Error);
            if (e.Exception is FileFormatException)
            {
                SelectWindowState(State.Text);
                lbStatus.Text = Properties.Resources.WrongFormatText;
            }
            else if (e.Exception is FileNotFoundException)
            {
                SelectWindowState(State.Text);
                lbStatus.Text = Properties.Resources.NotFoundText;
            }
            else
            {
                SelectWindowState(State.Text);
                lbStatus.Text = Properties.Resources.UnknownProblemText;
            }
            e.Handled = true;
        }


        private FlowDocument ParseODT(string path)
        {
            ODTReader reader = new ODTReader();
            return reader.Read(path);
        }

        void read_Completed(object sender, EventArgs e)
        {
            var dispatcherOp = pbParse.Dispatcher.BeginInvoke(
                     System.Windows.Threading.DispatcherPriority.Normal,
                     new Action(
                       delegate()
                       {
                           lbStatus.Text = Properties.Resources.DocumentReadyText;
                           pbParse.State = Elysium.Controls.ProgressState.Normal;
                       }
                   ));
        }

        private void ShowMessage(Message message)
        {
            //switch (message)
            //{
            //    case Message.WrongFormat:
            //        {
            //            imgOdt.Source = LoadImage(@"Resources\error_icon.png");
            //            lbMessage.Text = Properties.Resources.WrongFormatText;
            //            break;
            //        }
            //    case Message.NotFound:
            //        {
            //            imgOdt.Source = LoadImage(@"Resources\sad_face.png");
            //            lbMessage.Text = Properties.Resources.NotFoundText;
            //            break;
            //        }
            //    case Message.Unknown:
            //        {
            //            imgOdt.Source = LoadImage(@"Resources\question_mark.png");
            //            lbMessage.Text = Properties.Resources.UnknownProblemText;
            //            break;
            //        }
            //}
        }

        private static BitmapImage LoadImage(string path)
        {
            BitmapImage image = new BitmapImage();
            image.BeginInit();
            image.UriSource = new Uri(path, UriKind.RelativeOrAbsolute);
            image.DecodePixelWidth = 128;
            image.EndInit();
            return image;
        }

        private void btnSave_Click(object sender, RoutedEventArgs e)
        {
            SaveDialog();
        }

        

        private void scale_Click(object sender, EventArgs e)
        {
            scale.Value = 1;
        }

        private void save_Completed(object sender, EventArgs e)
        {
            lbStatus.Text = Properties.Resources.DocumentSavedText;
            pbParse.State = Elysium.Controls.ProgressState.Normal;
        }

        private void SavePDF(string outputPath)
        {
            try
            {
                PDFWriter pdf = new PDFWriter();
                pdf.Write(outputPath, lbText.Document);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message + e.StackTrace);
                SelectWindowState(State.Error);
                ShowMessage(Message.Unknown);
            }
        }

        private void btnSetting_Click(object sender, RoutedEventArgs e)
        {
            SelectWindowState(State.Settings);
        }

        private void scale_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            string msg = String.Format("{0}%", Convert.ToInt32(e.NewValue*100));
            if(scaleValue !=null) scaleValue.Text = msg;
        }


        private void OpenSettingsBar()
        {
            System.Windows.Media.Animation.DoubleAnimation da = new DoubleAnimation();
            da.From = sidebar.Width;
            da.To = this.Width / 3;
            da.Duration = TimeSpan.FromSeconds(0.3);
            sidebar.BeginAnimation(StackPanel.WidthProperty, da);
        }

        private void closeSettingsBar(object sender, RoutedEventArgs e)
        {
            System.Windows.Media.Animation.DoubleAnimation da = new DoubleAnimation();
            da.From = sidebar.Width;
            da.To = 0;
            da.Duration = TimeSpan.FromSeconds(0.3);
            sidebar.BeginAnimation(StackPanel.WidthProperty, da);
        }

        private void StartMenu_Click(object sender, MouseButtonEventArgs e)
        {
            if (sender.GetType() == typeof(StackPanel))
            {
                if ((sender as StackPanel).Name == "spOpen")
                {
                    OpenDialog();
                }
                if ((sender as StackPanel).Name == "spNew")
                {
                    NewDoc();
                }
            }
        }

        private void lbRecents_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (lbRecents.SelectedItem != null)
            {
                var item = (lbRecents.SelectedItem as Recent);
                OpenDoc(item.Uri);
                recents.Update(new Recent() {Uri = item.Uri, Name = item.Name, Date = DateTime.Now});
            }
        }

    }
}
