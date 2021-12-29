using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace HomeWork17
{
    /// <summary>
    /// Логика взаимодействия для FormAdd.xaml
    /// </summary>
    public partial class FormAdd : Window
    {
        public FormAdd()
        {
            InitializeComponent();
        }

        public FormAdd(DataRow row):this()
        {
            btCancel.Click += delegate { this.DialogResult = false; };
            btAdd.Click += delegate
            {
                row["LastName"] = tblastName.Text;
                row["FirstName"] = tbfirstName.Text;
                row["MiddleName"] = tbmmiddleName.Text;
                row["NumberTel"] = tbnumberTel.Text;
                row["Email"] = tbemail.Text;
                this.DialogResult = true;
            };
        }

    }
}
