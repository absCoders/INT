using System;
using System.Collections.Generic;
using System.Text;
using System.ComponentModel;
using System.Windows.Forms;

namespace ABSCS
{
    public class ABSCheckBox : Infragistics.Win.UltraWinEditors.UltraCheckEditor 
    {
        [System.ComponentModel.Bindable(true),
         DisplayName("ABSChecked"),
         Category("ABS"),
         Description("Contains a 1 if checked and a 0 (or null) of not checked"),
         DefaultValue(typeof(String),"0")]
        public string ABSChecked
        {
            get { return this.Checked ? "1" : "0"; }
            set { this.Checked = (string)value == "1"; }
        }
    }
}