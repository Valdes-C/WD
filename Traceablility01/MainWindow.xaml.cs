using System;
using System.Collections.Generic;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using Traceablility01.DataReading;
using Microsoft.WindowsAPICodePack.Dialogs;
using System.Windows.Forms;
using System.Diagnostics;
using Application = System.Windows.Forms.Application;

namespace Traceablility01
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            DateTime start,end;
            start= DateTime.Parse("2020-01-01");
            end = DateTime.Today;
            foreach (DateTime day in EachDay(start, end))
            {
                startDate.Items.Add(day.ToString("yyyy-MM-dd"));
            } 
            startDate.SelectedIndex=startDate.Items.Count-200;
            endDate.SelectedItem = DateTime.Today.ToString("yyyy-MM-dd");
            ReportFile.Text = Environment.CurrentDirectory.ToString()+"\\";
        }
        public IEnumerable<DateTime> EachDay(DateTime from, DateTime thru)
        {
            for (var day = from.Date; day.Date <= thru.Date; day = day.AddDays(1))
                yield return day;
        }

        private void startDate_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            System.Windows.Controls.ComboBox comboBox = (System.Windows.Controls.ComboBox)sender;

            endDate.Items.Clear();
            if (comboBox.SelectedValue.ToString() != string.Empty)
            {
                foreach (DateTime day in EachDay(DateTime.Parse(comboBox.SelectedValue.ToString()), DateTime.Today))
                {
                    endDate.Items.Add(day.ToString("yyyy-MM-dd"));
                }
            }
        }

        private void AddBatch_Click(object sender, RoutedEventArgs e)
        {
            if (numberPZ.Text!=string.Empty)
            {
                var bChck = true;
                foreach (var item in BatchList.Items)
                {
                    if (item.Equals(numberPZ.Text.Trim().ToUpper()))
                    {
                        bChck = false;
                        break;
                    }
                }
                if (bChck)
                {
                    BatchList.Items.Add(numberPZ.Text.Trim().ToUpper());
                    numberPZ.Text = string.Empty;
                }
            }
        }

        private void RemoveBatch_Click(object sender, RoutedEventArgs e)
        {
            int iL = BatchList.SelectedIndex;
            if (iL > -1) 
            {
                BatchList.Items.Remove(BatchList.Items[iL]);
                BatchList.Items.Refresh();
            }
            
        }

        private void DoReport_Click(object sender, RoutedEventArgs e)
        { 
            if (BatchList.Items.Count==0 || ReportFile.Text==string.Empty){return;}

            var waitForm = new Window1();
            waitForm.infoLabel.Content= "....";
            waitForm.Show();

            foreach (var batch in BatchList.Items) 
            {
                var path = ReportFile.Text + $"Identyfikacja_{batch.ToString().Replace('/','_')}_{DateTime.Now.ToString("yyyyMMddhhmmss")}.xlsx";
                waitForm.infoLabel.Content = batch.ToString();
                var report=new Report(startDate.Text, endDate.Text,batch.ToString());
                SavingData.Program.SaveToXLSX(path , report, batch.ToString(), startDate.Text, endDate.Text);
                if (File.Exists(path))
                { 
                    var msg = System.Windows.MessageBox.Show($"Identyfikacja {batch.ToString()} została zapisana jako {@path}", "Sukces");
                }
                else
                {
                    var msg = System.Windows.MessageBox.Show($"Identyfikacja {batch.ToString()} nie powiodła się.", "Porażka");
                }
            }
            waitForm.Close();

            Process.Start("explorer.exe",ReportFile.Text);
            
            this.Close();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            CommonOpenFileDialog dialog = new CommonOpenFileDialog();
            dialog.InitialDirectory = "C:\\";
            dialog.IsFolderPicker = true;
            if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
            {
                ReportFile.Text = dialog.FileName.ToString()+ "\\";
            }
        }
    }
}
