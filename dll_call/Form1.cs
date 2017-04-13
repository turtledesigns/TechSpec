using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using TechSpec;

namespace dll_call
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            TechSpec_Main.writeFileToJSON("results/specifications.json");
            this.label1.Text = TechSpec_Main.getJSONString();
        }
    }
}
