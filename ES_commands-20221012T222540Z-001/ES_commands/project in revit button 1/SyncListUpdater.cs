﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ES_commands
{
    public partial class SyncListUpdater : Form
    {
        public SyncListUpdater()
        {
            InitializeComponent();
        }

        public void textBox1_TextChanged(object sender, EventArgs e)
        {
            
        }

        private void SyncListUpdater_Load(object sender, EventArgs e)
        {
            timer1.Start();
           
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }
    }
}
