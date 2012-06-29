using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace StandaloneForm
{
    /// <summary>
    /// Логика взаимодействия для EnterRegistrationNumber.xaml
    /// </summary>
    public partial class EnterRegistrationNumber : Window
    {
        public EnterRegistrationNumber()
        {
            InitializeComponent();
        }

        public string RegNumber{set;get;}

        private void button1_Click(object sender, RoutedEventArgs e)
        {
            if (textBox1.Text.ToString() != "")
            {
                RegNumber = textBox1.Text.ToString();
                this.Close();
            }
        }

        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
                button1_Click(null, null);
        }
    }
}
