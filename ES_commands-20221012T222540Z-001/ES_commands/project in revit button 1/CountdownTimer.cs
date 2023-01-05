using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Timers;
using System.Windows.Forms;

namespace ES_commands
{
    public partial class CountdownTimer : Form
    {
        public System.Timers.Timer timer;
        public int h, m , s;

        public CountdownTimer()
        {
            InitializeComponent();
        }
        public void CountdownTimer_Load(object sender, EventArgs e)
        {
            timer = new System.Timers.Timer();
            timer.Interval = 100;
            timer.Elapsed += OnTimeEvent;
        }

        public void OnTimeEvent(object sender, ElapsedEventArgs e)
        {
            Invoke(new Action(() =>
            {
                s += 1;
                if (s == 60)
                {
                    s = 0;
                    m += 1;
                }
                if (m == 60)
                {
                    m = 0;
                    h += 1;
                }

                textBox1.Text = String.Format("{0}:{1}:{2}", h.ToString().PadLeft(2, '0'), m.ToString().PadLeft(2, '0'), s.ToString().PadLeft(2, '0'));

            }));
        }

        private void button1_Click(object sender, EventArgs e)
        {
            timer.Start();
        }

        public void start()
        {
            timer.Start();
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            timer.Stop();
        }

        private void CountdownTimer_FormClosed(object sender, FormClosedEventArgs e)
        {
            timer.Stop();
            Application.DoEvents();

        }
    }
}
