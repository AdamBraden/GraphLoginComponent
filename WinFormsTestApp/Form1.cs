using Microsoft.Toolkit.Services.MicrosoftGraph;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Toolkit.Services.Services.MicrosoftGraph;

namespace WinFormsTestApp
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

            // values to connect to Microsoft Graph
            graphLoginComponent1.ClientId = "f652df2c-f3f3-43b2-9db9-35af03051d74";
            graphLoginComponent1.Scopes = new string[] { MicrosoftGraphScope.UserRead };
        }

        private async void button1_Click(object sender, EventArgs e)
        {
            
            if (!await graphLoginComponent1.LoginAsync())
            {
                return;
            }

            //update the user's display fields
            label1.Text = graphLoginComponent1.DisplayName;
            label2.Text = graphLoginComponent1.JobTitle;

            pictureBox1.Image = graphLoginComponent1.Photo;

            // Do more things with the graph
            var graphClient = graphLoginComponent1.GraphServiceClient;
        }
    }
}
