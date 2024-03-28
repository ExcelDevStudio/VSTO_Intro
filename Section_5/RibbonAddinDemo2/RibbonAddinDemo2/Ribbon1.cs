using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace RibbonAddinDemo2
{
    public partial class Ribbon1
    {
        //private Microsoft.Office.Tools.Ribbon.RibbonButton buttonForGroup1;
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            //buttonForGroup1 =   this.Factory.CreateRibbonButton();
          
            //this.dropDown1.Buttons.Append(buttonForGroup1);
            //this.dropDown1.ResumeLayout(false);
            //this.dropDown1.PerformLayout();
            //this.PerformLayout();

            //buttonForGroup1.Name = "buttonForGroup1";
            //buttonForGroup1.Label = "buttonForGroup1";
        }

        private void button4_Click(object sender, RibbonControlEventArgs e)
        {
            MessageBox.Show("Hello world!");
        }

        private void toggleButton1_Click(object sender, RibbonControlEventArgs e)
        {
            RibbonToggleButton tb = (RibbonToggleButton)sender;
            MessageBox.Show(string.Format("My state is {0}", tb.Checked));
        }

        private void editBox1_TextChanged(object sender, RibbonControlEventArgs e)
        {
            RibbonEditBox ebox =  (RibbonEditBox)sender;
            MessageBox.Show(string.Format("My text is {0}", ebox.Text));
        }
    }
}
