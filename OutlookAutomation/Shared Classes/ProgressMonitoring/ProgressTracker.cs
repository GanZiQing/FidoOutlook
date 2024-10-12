﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OutlookAutomation
{
    public partial class ProgressTracker : Form
    {
        //public bool terminateProcess = false;
        public int progress;
        public int progressMax;
        public string labelText;
        public ProgressTracker()
        {
            InitializeComponent();
            progressMax = 100;
            ResetProgress();
        }

        public void ResetProgress()
        {
            progress = 0;
            RefreshProgress();
        }

        public void UpdateProgress(int progressValue)
        {
            // Use the following to update progress, do not call this function directly
            // worker.ReportProgress(ConvertToProgress(prog, maxprog));
            progress = progressValue;
            RefreshProgress();

            //public void UpdateProgressWithMax(int currentProgress, int maxProgress)
            //{
            //    double progressDouble = Convert.ToDouble(currentProgress) / Convert.ToDouble(maxProgress) * 100;
            //    progress = Convert.ToInt32(progressDouble);
            //    RefreshProgress();
            //}
        }


        private void RefreshProgress()
        {
            if (InvokeRequired)
            {
                Invoke(new Action(RefreshProgress));
                return;
            }
            ProgressBar1.Value = progress;
            ProgressLabel.Text = $"{progress}%";
            ProgressLabel.Update();
            ThreadLabel.Update();
        }

        public void UpdateStatus(string msg)
        {
            if (InvokeRequired)
            {
                Invoke(new Action<string>(UpdateStatus), msg);
                return;
            }
            ThreadLabel.Text = msg;
            ThreadLabel.Update();
        }

        private void CancelButton_Click(object sender, EventArgs e)
        {
            //    terminateProcess = true;
            UpdateStatus("Cancelling...");
            Close();
        }


        //BackgroundWorker WorkerThread;
        //public void ShowForm()
        //{
        //    WorkerThread = new BackgroundWorker();
        //    WorkerThread.WorkerReportsProgress = true;
        //    WorkerThread.RunWorkerAsync();
        //    WorkerThread.DoWork += WorkerThread_DoWork;
        //    //WorkerThread.ProgressChanged += WorkerThread_ProgressChanged;
        //}

        //void WorkerThread_DoWork(object sender, DoWorkEventArgs e)
        //{
        //    Show();
        //}
    }
}
