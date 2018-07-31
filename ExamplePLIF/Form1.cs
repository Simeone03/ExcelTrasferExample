using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExamplePLIF
{
    public partial class Form1 : Form
    {
        Functions myFunctions;
        private static string outputFilePath = @"C:\Users\Simon\Documents\PLIF";
        public Form1()
        {
            InitializeComponent();
            myFunctions = new Functions();
        }

        private void OpenFIle_Click(object sender, EventArgs e)
        {
            DialogResult result = openFileDialog1.ShowDialog();
            if (result == DialogResult.OK) // Test result.
            {
                string file = openFileDialog1.FileName;
                //Console.WriteLine(file);
                myFunctions.insertPFA(file, outputFilePath);
            }
            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            DialogResult result = openFileDialog1.ShowDialog();
            if (result == DialogResult.OK) // Test result.
            {
                outputFilePath = openFileDialog1.FileName;
                
            }

        }
    }
}
