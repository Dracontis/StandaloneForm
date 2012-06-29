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
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Collections;

namespace StandaloneForm
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            ControlSubject1.ItemsSource = LoadSubjects();
            ControlSubject2.ItemsSource = LoadSubjects();
            ControlSubject3.ItemsSource = LoadSubjects();
            ControlFirstPriority.ItemsSource = LoadSpecs();
            ControlSecondPriority.ItemsSource = LoadSpecs();
            ControlThirdPriority.ItemsSource = LoadSpecs();            
        }

        private void GenerateDocuments(object sender, RoutedEventArgs e)
        {
            
        }

        private String[] LoadSubjects()
        {
            String[] Subjects = 
            {
                "Математика",
                "Русский язык",
                "Физика",

            };
            return Subjects;
        }

        private String[] LoadSpecs()
        {
            String[] Specs = 
            {
                "230100 Информатика и выч. Техника",
                "160100 Авиа-ракетостроение",
                "Информатика и вычислительная техника",
            };
            return Specs;
        }

        public static Dictionary<String, List<String>> GetFaculties()
        {
            Dictionary<String, List<String>> faculties = new Dictionary<String, List<String>>()
            {
                { "Информатика и вычислительная техника", new List<String>() { "1", "2" } },
                { "160100 Авиа-ракетостроение", new List<String>() { "1" } }
            };

            return faculties;
        }

        /// <summary>
        /// Get faculty title by number
        /// </summary>
        /// <param name="num">Faculty number</param>
        /// <returns>String, that represent faculty by its number</returns>
        public static String GF(int num)
        {
            return "";
        }

        private void Propreties_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var PropCombo = (ComboBox)sender;
            String spec = PropCombo.SelectedValue as String;

            // If something choosen, select facs
            if (spec.Equals(""))
                return;
            else
            {
                Faculty FacultyDialog = new Faculty(spec);
                FacultyDialog.ShowDialog();
            }
            if (PropCombo.Name.Equals("ControlFirstPriority"))
            {

            }
            else if (PropCombo.Name.Equals("ControlSecondPriority"))
            {

            }
            else if (PropCombo.Name.Equals("ControlThirdPriority"))
            {

            }
        }
    }
}
