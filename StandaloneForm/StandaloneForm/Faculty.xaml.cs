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
using System.Collections;

namespace StandaloneForm
{
    public partial class Faculty : Window
    {
        int Priority = 0;
        List<String> Faculties;
        public List<String> Output { set; get; }

        public Faculty(String spec)
        {
            Output = new List<String>();
            Dictionary<String, List<String>> faculties = MainWindow.GetFaculties();
            Faculties = faculties[spec];
            // Quit if there only one faculty
            if (Faculties.Count <= 1)
            {
                Output = Faculties;
                DialogResult = true;
            }

            InitializeComponent();
            // Set size according to number of controls
            FacultyWindow.Height = 25 * Faculties.Count + 75;
            FacultyWindow.Width = 375;

            GenerateControls(Faculties);            
        }

        /*
         * Support functions
         */

        private void GenerateControls(List<String> Faculties)
        {
            for (int i = 0; i < Faculties.Count; i++)
            {
                Label FacultyName = new Label();
                FacultyName.Content = Faculties[i];
                FacultyName.SetValue(Grid.RowProperty, i);
                FacultyName.SetValue(Grid.ColumnProperty, 0);
                Button FacultyChoice = new Button();
                FacultyChoice.Content = "+";
                FacultyChoice.Name = "F_" + i.ToString();
                FacultyChoice.SetValue(Grid.RowProperty, i);
                FacultyChoice.SetValue(Grid.ColumnProperty, 1);
                FacultyChoice.Click += ButtonFacultyChoice_Click;

                ListOfFaculties.Children.Add(FacultyName);
                ListOfFaculties.Children.Add(FacultyChoice);
            }
        }

        /*
         *  Event handlers
         */

        private void ButtonFacultyChoice_Click(object sender, RoutedEventArgs e)
        {
            var button = (Button)sender;

            int i = Int32.Parse(button.Name.Substring(2));

            Output.Add(Faculties[i]);
            button.Content = (++Priority).ToString();
            button.IsEnabled = false;
        }

        private void ButtonOk_Click(object sender, RoutedEventArgs e)
        {
            if (Priority == 0)
                MessageBox.Show("Вы должны выбрать хотя бы один факультет");
            else
                DialogResult = true;
        }

        private void ButtonReset_Click(object sender, RoutedEventArgs e)
        {
            Priority = 0;
            Output = new List<String>();
            ListOfFaculties.Children.Clear();
            GenerateControls(Faculties);

        }

        private void ButtonCancel_Click(object sender, RoutedEventArgs e)
        {
            if (Priority == 0)
                MessageBox.Show("Вы должны выбрать хотя бы один факультет");
            else
                DialogResult = false;
        }
    }
}
