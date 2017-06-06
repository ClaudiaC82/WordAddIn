﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace WordAddIn2
{
    public partial class ConfigurazioneMessaggio : Form
    {
        string messaggio = "";
        public ConfigurazioneMessaggio()
        {
            InitializeComponent();
        }

        private void ConfigurazioneMessaggio_Load(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Word.Document documento = Globals.ThisAddIn.Application.ActiveDocument;
            messaggio = documento.Content.Text;
        }
    }
}
