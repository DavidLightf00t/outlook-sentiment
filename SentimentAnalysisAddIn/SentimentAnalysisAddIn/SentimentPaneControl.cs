﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SentimentAnalysisAddIn
{
    public partial class SentimentPaneControl : UserControl
    {
        public SentimentPaneControl()
        {
            InitializeComponent();

            // Navigate to local UI
            webBrowser1.Navigate("https://localhost:3000/taskpane.html");
        }
    }
}
