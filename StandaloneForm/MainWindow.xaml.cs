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
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;

namespace StandaloneForm
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private static ExcelFunctions ExcelForm, ExcelSummary;
        private static String _ProgramPath;
        private Specialization[] Specs = new Specialization[3];
        private static String ProgramPath
        {
            get
            {
                if (_ProgramPath != null) return _ProgramPath;
                _ProgramPath = new FileInfo(Assembly.GetExecutingAssembly().GetName().FullName).DirectoryName;
                if (!_ProgramPath.EndsWith("\\")) _ProgramPath += "\\";
                return _ProgramPath;
            }
        }
        public MainWindow()
        {
            InitializeComponent();
            ControlSubject1.ItemsSource = LoadSubjects();
            ControlSubject2.ItemsSource = LoadSubjects();
            ControlSubject3.ItemsSource = LoadSubjects();
            ControlSubject4.ItemsSource = LoadSubjects();
            ControlSubject5.ItemsSource = LoadSubjects();
            ControlFirstPriority.ItemsSource = LoadSpecs();
            ControlSecondPriority.ItemsSource = LoadSpecs();
            ControlThirdPriority.ItemsSource = LoadSpecs();
            ControlSchoolType.ItemsSource = LoadSchoolTypes();
            ControlMagistrSpec.ItemsSource = LoadSpecsForMasters();
            ControlFunding.ItemsSource = LoadFunding();
            ControlLearningForm.ItemsSource = LoadLearningForm();
            ControlSemesterNum.ItemsSource = LoadSemesters();
        }

        private String[] LoadSchoolTypes()
        {
            String[] SchoolTypes = 
            {
                "школа",
                "НПО",
                "СПО",
                "ВПО"
            };
            return SchoolTypes;
        }

        public int GetSchoolTypeNumber(String name)
        {
            List<String> SchoolTypeNames = new List<String>(LoadSchoolTypes());
            int SchoolTypeNumber = SchoolTypeNames.IndexOf(name) + 1;
            return SchoolTypeNumber;
        }

        private void GenerateDocuments(object sender, RoutedEventArgs e)
        {
            Address Addr = new Address();
            Addr.Index = ControlIndex.Text;
            Addr.Region = ControlRegion.Text;
            Addr.Town = ControlTown.Text;
            Addr.AppAddress = ControlAddress.Text;
            List<EnteranceExamination> LExams = new List<EnteranceExamination>();
            EnteranceExamination Exam = new EnteranceExamination();
            if (ControlSubject1.SelectedIndex != -1) Exam.Subject = ControlSubject1.SelectedItem.ToString();
            Exam.Points = ControlPoints1.Text;
            if (ControlEge1.IsChecked == true)
            {
                Exam.Ege = true;
                Exam.Olimp = false;
            }
            if (ControlOlimp1.IsChecked == true)
            {
                Exam.Ege = false;
                Exam.Olimp = true;
            }
            Exam.TitleAndNum = ControlTitleAndNum1.Text;
            Exam.DocumentIssuedDate = ControlDocumentIssuedDate1.Text;
            LExams.Add(Exam);
            Exam = new EnteranceExamination();
            if (ControlSubject2.SelectedIndex != -1) Exam.Subject = ControlSubject2.SelectedItem.ToString();
            Exam.Points = ControlPoints2.Text;
            if (ControlEge2.IsChecked == true)
            {
                Exam.Ege = true;
                Exam.Olimp = false;
            }
            if (ControlOlimp2.IsChecked == true)
            {
                Exam.Ege = false;
                Exam.Olimp = true;
            }
            Exam.TitleAndNum = ControlTitleAndNum2.Text;
            Exam.DocumentIssuedDate = ControlDocumentIssuedDate2.Text;
            LExams.Add(Exam);
            Exam = new EnteranceExamination();
            if (ControlSubject3.SelectedIndex != -1) Exam.Subject = ControlSubject3.SelectedItem.ToString();
            Exam.Points = ControlPoints3.Text;
            if (ControlEge3.IsChecked == true)
            {
                Exam.Ege = true;
                Exam.Olimp = false;
            }
            if (ControlOlimp3.IsChecked == true)
            {
                Exam.Ege = false;
                Exam.Olimp = true;
            }
            Exam.TitleAndNum = ControlTitleAndNum3.Text;
            Exam.DocumentIssuedDate = ControlDocumentIssuedDate3.Text;
            LExams.Add(Exam);
            Exam = new EnteranceExamination();
            if (ControlSubject4.SelectedIndex != -1) Exam.Subject = ControlSubject4.SelectedItem.ToString();
            Exam.Points = ControlPoints4.Text;
            if (ControlEge4.IsChecked == true)
            {
                Exam.Ege = true;
                Exam.Olimp = false;
            }
            if (ControlOlimp4.IsChecked == true)
            {
                Exam.Ege = false;
                Exam.Olimp = true;
            }
            Exam.TitleAndNum = ControlTitleAndNum4.Text;
            Exam.DocumentIssuedDate = ControlDocumentIssuedDate4.Text;
            LExams.Add(Exam);
            Exam = new EnteranceExamination();
            if (ControlSubject5.SelectedIndex != -1) Exam.Subject = ControlSubject5.SelectedItem.ToString();
            Exam.Points = ControlPoints5.Text;
            if (ControlEge5.IsChecked == true)
            {
                Exam.Ege = true;
                Exam.Olimp = false;
            }
            if (ControlOlimp5.IsChecked == true)
            {
                Exam.Ege = false;
                Exam.Olimp = true;
            }
            Exam.TitleAndNum = ControlTitleAndNum5.Text;
            Exam.DocumentIssuedDate = ControlDocumentIssuedDate5.Text;
            LExams.Add(Exam);
            EnterRegistrationNumber ERN = new EnterRegistrationNumber();
            ERN.ShowDialog();
            String RN = ERN.RegNumber;
            if(RN.Split(new Char[] { '-' }).Length != 3)
            {
                MessageBox.Show("Неправильный формат регистрационного номера.");
                return;
            }

            if (RN != "")
            {
                //OMFG begins
                Applicant NewApplicant = new Applicant(
                    RN,
                    ControlFirstName.Text,
                    ControlSecondName.Text,
                    ControlLastName.Text,
                    ControlBirthDate.Text,
                    ControlBirthPlace.Text,
                    Addr,
                    ControlCitizenship.Text,
                    ControlPassport.Text,
                    ControlSerial.Text,
                    ControlNumber.Text,
                    ControlPassportIssuedDate.Text,
                    ControlHomePhone.Text,
                    ControlLearningForm.Text,
                    ControlFunding.Text,
                    Specs,
                    ControlEducation.Text,
                    ControlTypeOfEducationDocument.Text,
                    ControlNumberOfEducationDocument.Text,
                    ControlEducationIssuedDate.Text,
                    LExams.ToArray(),
                    Convert.ToBoolean(ControlAllowUniversityExams.IsChecked),
                    ControlGrounds.Text,
                    ControlFacilities.Text,
                    ControlOlimpiads.Text,
                    Convert.ToBoolean(ControlNeedDorm.IsChecked),
                    ControlIssuedUniversityEducation.Text,
                    Convert.ToBoolean(ControlMATICourses.IsChecked),
                    Convert.ToBoolean(ControlMATISchool.IsChecked),
                    Convert.ToBoolean(ControlAttest.IsChecked),
                    ControlSex.Text,
                    ControlSchoolType.Text,
                    ControlSchoolName.Text
                    );
                //OMFG ends
                if (Convert.ToBoolean(ControlMagistrProof.IsChecked))
                {
                    Master NewMaster = new Master(
                        ControlMagistrUniversity.Text,
                        ControlMagistrDiploma.Text,
                        Specs
                    );
                    GenerateMasterDocuments(NewApplicant, NewMaster);
                    string[] split = NewApplicant.RegNumber.Split(new Char[] { '-' });
                    GenerateUSECheck(NewApplicant, GetFacultyNumber(split[0])); //USE means Unified State Exam, ЕГЭ короче
                }
                else
                {
                    if (NewApplicant.LearningForm.Equals("очной"))
                    {
                        File.Copy(ProgramPath + @"Шаблоны\MainFormTemplate.xls", ProgramPath + @"Документы (оч)\" + NewApplicant.SecondName + " " + NewApplicant.RegNumber + ".xls");
                        ExcelForm = new ExcelFunc();
                        ExcelForm.OpenDocument(ProgramPath + @"Документы (оч)\" + NewApplicant.SecondName + " " + NewApplicant.RegNumber + ".xls");
                    }
                    else if (NewApplicant.LearningForm.Equals("очно-заочной"))
                    {
                        File.Copy(ProgramPath + @"Шаблоны\MainFormTemplate.xls", ProgramPath + @"Документы (оч-заоч)\" + NewApplicant.SecondName + " " + NewApplicant.RegNumber + ".xls");
                        ExcelForm = new ExcelFunc();
                        ExcelForm.OpenDocument(ProgramPath + @"Документы (оч-заоч)\" + NewApplicant.SecondName + " " + NewApplicant.RegNumber + ".xls");
                    }
                    else
                    {
                        if (File.Exists(ProgramPath + @"Документы (др)\" + NewApplicant.SecondName + " " + NewApplicant.RegNumber + ".xls")) 
                        {
                            MessageBoxResult result = MessageBox.Show(this, "Документ с таким именем уже существует и будет безвозвратно утерян. Продолжить?", "Предупреждение!", MessageBoxButton.YesNo, MessageBoxImage.Warning, MessageBoxResult.No, MessageBoxOptions.None);
                            if(result == MessageBoxResult.No)
                                return;
                            else
                                File.Delete(ProgramPath + @"Документы (др)\" + NewApplicant.SecondName + " " + NewApplicant.RegNumber + ".xls");
                        }
                        File.Copy(ProgramPath + @"Шаблоны\MainFormTemplate.xls", ProgramPath + @"Документы (др)\" + NewApplicant.SecondName + " " + NewApplicant.RegNumber + ".xls");
                        ExcelForm = new ExcelFunc();
                        ExcelForm.OpenDocument(ProgramPath + @"Документы (др)\" + NewApplicant.SecondName + " " + NewApplicant.RegNumber + ".xls");
                    }
                    FillExcelForm(NewApplicant);
                    ExcelForm.OpenWorksheet(2);
                    FillExcelRequest(NewApplicant);
                    ExcelForm.OpenWorksheet(3);
                    FillExcelFacultiesList(NewApplicant);
                    ExcelForm.OpenWorksheet(4);
                    FillExcelContract(NewApplicant);
                    ExcelForm.OpenWorksheet(5);
                    FillExcelLKS(NewApplicant);
                    ExcelForm.OpenWorksheet(6);
                    FillExcelProfile(NewApplicant);
                    ExcelForm.OpenWorksheet(7);
                    FillExcelListOfExams(NewApplicant);
                    ExcelForm.OpenWorksheet(8);
                    FillExcelReceipt(NewApplicant);
                    ExcelForm.CloseDocument();
                    int Fac;
                    if (NewApplicant.Specs[0] != null)
                    {
                        String SFac = NewApplicant.RegNumber.Remove(2, (NewApplicant.RegNumber.Length - 2));
                        SFac = GetFacultyNumber(SFac);
                        Int32.TryParse(SFac, out Fac);
                        FillSummary(0, NewApplicant);
                    }
                }
            }

            try
            {
                if (ExcelForm != null) ExcelForm.Dispose();
                if (ExcelSummary != null) ExcelSummary.Dispose();
            }
            catch (COMException Excep)
            {
                // Тут нет костыля. Совсем. Иди отсюда мальчик. Или девочка. В общем, вали быстро!
                // It's a lion! Get in the car!
            }
        }
        private void FillExcelForm(Applicant NewApplicant)
        {
            ExcelForm.SetValue("C3", NewApplicant.RegNumber);
            ExcelForm.SetValue("C4", NewApplicant.SecondName);
            ExcelForm.SetValue("C5", NewApplicant.FirstName);
            ExcelForm.SetValue("C6", NewApplicant.LastName);
            ExcelForm.SetValue("C7", NewApplicant.BirthDate);
            ExcelForm.SetValue("C8", NewApplicant.BirthPlace);
            ExcelForm.SetValue("C9", NewApplicant.Address_.Index + ", " + NewApplicant.Address_.Region);
            ExcelForm.SetValue("C10", NewApplicant.Address_.Town);
            ExcelForm.SetValue("C11", NewApplicant.Address_.AppAddress);
            ExcelForm.SetValue("C12", NewApplicant.Citizenship);
            ExcelForm.SetValue("C13", NewApplicant.Passport);
            ExcelForm.SetValue("C14", NewApplicant.Serial);
            ExcelForm.SetValue("C15", NewApplicant.Number);
            ExcelForm.SetValue("C16", NewApplicant.PassportIssuedDate);
            ExcelForm.SetValue("C17", NewApplicant.HomePhone);
            ExcelForm.SetValue("C18", NewApplicant.LearningForm);
            ExcelForm.SetValue("C19", NewApplicant.Funding);
            if (NewApplicant.Specs[0] != null)
            {
                ExcelForm.SetValue("C21", NewApplicant.Specs[0].Spec);
                if (NewApplicant.Specs[0].Spec != "Неизвестно") ExcelForm.SetValue("C22", NewApplicant.Specs[0].Faculty.ToArray()[0]);
            }
            if (NewApplicant.Specs[1] != null)
            {
                ExcelForm.SetValue("C23", NewApplicant.Specs[1].Spec);
                if (NewApplicant.Specs[1].Spec != "Неизвестно") ExcelForm.SetValue("C24", NewApplicant.Specs[1].Faculty.ToArray()[0]);
            }
            if (NewApplicant.Specs[2] != null)
            {
                ExcelForm.SetValue("C25", NewApplicant.Specs[2].Spec);
                if (NewApplicant.Specs[2].Spec != "Неизвестно") ExcelForm.SetValue("C26", NewApplicant.Specs[2].Faculty.ToArray()[0]);
            }
            ExcelForm.SetValue("C27", NewApplicant.Education);
            ExcelForm.SetValue("C28", NewApplicant.EducationDocument);
            ExcelForm.SetValue("C29", NewApplicant.NumberOfEducationDocument);
            ExcelForm.SetValue("C30", NewApplicant.EducationIssuedDate);
            ExcelForm.SetValue("B33", NewApplicant.EnteranceExaminations[0].Subject);
            ExcelForm.SetValue("C33", NewApplicant.EnteranceExaminations[0].Points);
            if (NewApplicant.EnteranceExaminations[0].Ege == true) ExcelForm.SetValue("D33", "да");
            else if (NewApplicant.EnteranceExaminations[0].Olimp == true) ExcelForm.SetValue("E33", "да");
            ExcelForm.SetValue("F33", NewApplicant.EnteranceExaminations[0].TitleAndNum);
            ExcelForm.SetValue("G33", NewApplicant.EnteranceExaminations[0].DocumentIssuedDate);
            ExcelForm.SetValue("B34", NewApplicant.EnteranceExaminations[1].Subject);
            ExcelForm.SetValue("C34", NewApplicant.EnteranceExaminations[1].Points);
            if (NewApplicant.EnteranceExaminations[1].Ege == true) ExcelForm.SetValue("D34", "да");
            else if (NewApplicant.EnteranceExaminations[1].Olimp == true) ExcelForm.SetValue("E34", "да");
            ExcelForm.SetValue("F34", NewApplicant.EnteranceExaminations[1].TitleAndNum);
            ExcelForm.SetValue("G34", NewApplicant.EnteranceExaminations[1].DocumentIssuedDate);
            ExcelForm.SetValue("B35", NewApplicant.EnteranceExaminations[2].Subject);
            ExcelForm.SetValue("C35", NewApplicant.EnteranceExaminations[2].Points);
            if (NewApplicant.EnteranceExaminations[2].Ege == true) ExcelForm.SetValue("D35", "да");
            else if (NewApplicant.EnteranceExaminations[2].Olimp == true) ExcelForm.SetValue("E35", "да");
            ExcelForm.SetValue("F35", NewApplicant.EnteranceExaminations[2].TitleAndNum);
            ExcelForm.SetValue("G35", NewApplicant.EnteranceExaminations[2].DocumentIssuedDate);
            ExcelForm.SetValue("B36", NewApplicant.EnteranceExaminations[3].Subject);
            ExcelForm.SetValue("C36", NewApplicant.EnteranceExaminations[3].Points);
            if (NewApplicant.EnteranceExaminations[3].Ege == true) ExcelForm.SetValue("D36", "да");
            else if (NewApplicant.EnteranceExaminations[3].Olimp == true) ExcelForm.SetValue("E36", "да");
            ExcelForm.SetValue("F36", NewApplicant.EnteranceExaminations[3].TitleAndNum);
            ExcelForm.SetValue("G36", NewApplicant.EnteranceExaminations[3].DocumentIssuedDate);
            ExcelForm.SetValue("B37", NewApplicant.EnteranceExaminations[4].Subject);
            ExcelForm.SetValue("C37", NewApplicant.EnteranceExaminations[4].Points);
            if (NewApplicant.EnteranceExaminations[4].Ege == true) ExcelForm.SetValue("D37", "да");
            else if (NewApplicant.EnteranceExaminations[4].Olimp == true) ExcelForm.SetValue("E37", "да");
            ExcelForm.SetValue("F37", NewApplicant.EnteranceExaminations[4].TitleAndNum);
            ExcelForm.SetValue("G37", NewApplicant.EnteranceExaminations[4].DocumentIssuedDate);
            if (NewApplicant.AllowUniversityExams == true) ExcelForm.SetValue("C56", "да");
            ExcelForm.SetValue("C57", NewApplicant.Grounds);
            ExcelForm.SetValue("C59", NewApplicant.Facilities);
            ExcelForm.SetValue("C61", NewApplicant.Olimpiads);
            if (NewApplicant.NeedDorm == true) ExcelForm.SetValue("C63", "нуждаюсь");
            else ExcelForm.SetValue("C63", "не нуждаюсь");
            ExcelForm.SetValue("C64", NewApplicant.IssuedUniversityEducation);
            ExcelForm.SetValue("C65", DateTime.Today.ToString("dd.MM.yyyy"));
        }
        private void FillExcelRequest(Applicant NewApplicant)
        {
            ExcelForm.SetValue("AT1", NewApplicant.RegNumber);
            ExcelForm.SetValue("G5", NewApplicant.SecondName);
            ExcelForm.SetValue("G6", NewApplicant.FirstName);
            ExcelForm.SetValue("G7", NewApplicant.LastName);
            ExcelForm.SetValue("H8", NewApplicant.BirthDate);
            ExcelForm.SetValue("I9", NewApplicant.BirthPlace);
            ExcelForm.SetValue("P11", NewApplicant.Address_.Index + ", " + NewApplicant.Address_.Region + ", " + NewApplicant.Address_.Town);
            ExcelForm.SetValue("A12", NewApplicant.Address_.AppAddress);
            ExcelForm.SetValue("AF5", NewApplicant.Citizenship);
            ExcelForm.SetValue("AS6", NewApplicant.Passport);
            ExcelForm.SetValue("AM7", NewApplicant.Serial);
            ExcelForm.SetValue("AT7", NewApplicant.Number);
            ExcelForm.SetValue("Z9", NewApplicant.PassportIssuedDate);
            ExcelForm.SetValue("AL12", NewApplicant.HomePhone);
            ExcelForm.SetValue("D15", NewApplicant.LearningForm);
            ExcelForm.SetValue("Y15", NewApplicant.Funding);
            if (NewApplicant.Specs[0] != null)
            {
                ExcelForm.SetValue("U17", NewApplicant.Specs[0].Spec);
            }
            if (NewApplicant.Specs[1] != null)
            {
                ExcelForm.SetValue("U19", NewApplicant.Specs[1].Spec);
            }
            if (NewApplicant.Specs[2] != null)
            {
                ExcelForm.SetValue("U21", NewApplicant.Specs[2].Spec);
            }
            ExcelForm.SetValue("E23", NewApplicant.Education);
            ExcelForm.SetValue("N24", NewApplicant.EducationDocument);
            ExcelForm.SetValue("A27", NewApplicant.EnteranceExaminations[0].Subject);
            ExcelForm.SetValue("N27", NewApplicant.EnteranceExaminations[0].Points);
            if (NewApplicant.EnteranceExaminations[0].Ege == true) ExcelForm.SetValue("U27", "да");
            else if (NewApplicant.EnteranceExaminations[0].Olimp == true) ExcelForm.SetValue("X27", "да");
            ExcelForm.SetValue("AA27", NewApplicant.EnteranceExaminations[0].TitleAndNum);
            ExcelForm.SetValue("AV27", NewApplicant.EnteranceExaminations[0].DocumentIssuedDate);
            ExcelForm.SetValue("A28", NewApplicant.EnteranceExaminations[1].Subject);
            ExcelForm.SetValue("N28", NewApplicant.EnteranceExaminations[1].Points);
            if (NewApplicant.EnteranceExaminations[1].Ege == true) ExcelForm.SetValue("U28", "да");
            else if (NewApplicant.EnteranceExaminations[0].Olimp == true) ExcelForm.SetValue("X28", "да");
            ExcelForm.SetValue("AA28", NewApplicant.EnteranceExaminations[1].TitleAndNum);
            ExcelForm.SetValue("AV28", NewApplicant.EnteranceExaminations[1].DocumentIssuedDate);
            ExcelForm.SetValue("A29", NewApplicant.EnteranceExaminations[2].Subject);
            ExcelForm.SetValue("N29", NewApplicant.EnteranceExaminations[2].Points);
            if (NewApplicant.EnteranceExaminations[2].Ege == true) ExcelForm.SetValue("U29", "да");
            else if (NewApplicant.EnteranceExaminations[0].Olimp == true) ExcelForm.SetValue("X29", "да");
            ExcelForm.SetValue("AA29", NewApplicant.EnteranceExaminations[2].TitleAndNum);
            ExcelForm.SetValue("AV29", NewApplicant.EnteranceExaminations[2].DocumentIssuedDate);
            if (NewApplicant.AllowUniversityExams == true) ExcelForm.SetValue("I38", NewApplicant.Grounds);
            ExcelForm.SetValue("S42", NewApplicant.Facilities);
            ExcelForm.SetValue("E41", NewApplicant.Olimpiads);
            if (NewApplicant.NeedDorm == true) ExcelForm.SetValue("I62", "нуждаюсь");
            else ExcelForm.SetValue("I62", "не нуждаюсь");
            ExcelForm.SetValue("O64", NewApplicant.IssuedUniversityEducation);
            if (NewApplicant.Specs[0] != null)
            {
                ExcelForm.SetValue("AB71", NewApplicant.Specs[0].Spec);
            }
            ExcelForm.SetValue("A81", DateTime.Today.ToString("dd.MM.yyyy"));
            ExcelForm.SetValue("M81", NewApplicant.SecondName + " " + NewApplicant.FirstName + " " + NewApplicant.LastName);
            ExcelForm.SetValue("W34", NewApplicant.SecondName + " " + NewApplicant.FirstName + " " + NewApplicant.LastName);
        }
        private void FillExcelContract(Applicant NewApplicant)
        {
            if (NewApplicant.Specs[0] != null)
            {
                ExcelForm.SetValue("A18", NewApplicant.Specs[0].Spec);
                ExcelForm.SetValue("E20", "№" + GetFacultyNumber(NewApplicant.Specs[0].Faculty.ToArray()[0]));
            }
            ExcelForm.SetValue("A11", NewApplicant.SecondName + " " + NewApplicant.FirstName + " " + NewApplicant.LastName);
            if (NewApplicant.IssuedUniversityEducation.Equals("Впервые")) ExcelForm.SetValue("K16", "1-го");
            else ExcelForm.SetValue("K16", "2-го");
            ExcelForm.SetValue("P18", NewApplicant.LearningForm);
            ExcelForm.SetValue("A11", NewApplicant.SecondName + " " + NewApplicant.FirstName + " " + NewApplicant.LastName);
            ExcelForm.SetValue("F138", NewApplicant.Address_.Index + " " + NewApplicant.Address_.Region + ", " + NewApplicant.Address_.Town + ", " + NewApplicant.Address_.AppAddress);
            ExcelForm.SetValue("E139", NewApplicant.HomePhone);
            ExcelForm.SetValue("N139", NewApplicant.BirthDate);
            ExcelForm.SetValue("AA139", NewApplicant.Serial + " " + NewApplicant.Number);
            ExcelForm.SetValue("F140", NewApplicant.PassportIssuedDate);
            ExcelForm.SetValue("D137", NewApplicant.SecondName + " " + NewApplicant.FirstName + " " + NewApplicant.LastName);
            //ExcelForm.SetValue("B156", NewApplicant.SecondName + " " + NewApplicant.FirstName + " " + NewApplicant.LastName);
        }
        private void FillExcelFacultiesList(Applicant NewApplicant)
        {

            ExcelForm.SetValue("AR3", NewApplicant.RegNumber);
            ExcelForm.SetValue("E6", NewApplicant.SecondName + " " + NewApplicant.FirstName + " " + NewApplicant.LastName);
            ExcelForm.SetValue("L47", NewApplicant.SecondName + " " + NewApplicant.FirstName + " " + NewApplicant.LastName);
            ExcelForm.SetValue("B47", DateTime.Today.ToString("dd.MM.yyyy"));
            int cnt = 13;
            if (NewApplicant.Specs[0] != null)
            {
                ExcelForm.SetValue("Y10", NewApplicant.Specs[0].Spec);
                foreach (String Fc in NewApplicant.Specs[0].Faculty)
                {
                    ExcelForm.SetValue("AK" + cnt.ToString(), "№" + GetFacultyNumber(Fc));
                    ExcelForm.SetValue("C" + cnt.ToString(), Fc.Remove(0, 3));
                    cnt++;
                }
            }
            if (NewApplicant.Specs[1] != null)
            {
                ExcelForm.SetValue("Y22", NewApplicant.Specs[1].Spec);
                cnt = 25;
                foreach (String Fc in NewApplicant.Specs[1].Faculty)
                {
                    ExcelForm.SetValue("AK" + cnt.ToString(), "№" + GetFacultyNumber(Fc));
                    ExcelForm.SetValue("C" + cnt.ToString(), Fc.Remove(0, 3));
                    cnt++;
                }
            }
            if (NewApplicant.Specs[2] != null)
            {
                ExcelForm.SetValue("Y34", NewApplicant.Specs[2].Spec);
                cnt = 37;
                foreach (String Fc in NewApplicant.Specs[2].Faculty)
                {
                    ExcelForm.SetValue("AK" + cnt.ToString(), "№" + GetFacultyNumber(Fc));
                    ExcelForm.SetValue("C" + cnt.ToString(), Fc.Remove(0, 3));
                    cnt++;
                }
            }
        }
        private void FillExcelLKS(Applicant NewApplicant)
        {
            ExcelForm.SetValue("BI6", NewApplicant.RegNumber);
            ExcelForm.SetValue("G6", NewApplicant.SecondName);
            ExcelForm.SetValue("T6", NewApplicant.FirstName);
            ExcelForm.SetValue("AJ6", NewApplicant.LastName);
            ExcelForm.SetValue("K8", NewApplicant.BirthDate);
            ExcelForm.SetValue("B11", NewApplicant.BirthPlace);
            ExcelForm.SetValue("I12", NewApplicant.Citizenship);
            ExcelForm.SetValue("B15", NewApplicant.SchoolType + " " + NewApplicant.SchoolName);
            ExcelForm.SetValue("B16", NewApplicant.Address_.Town + ", " + NewApplicant.EducationIssuedDate);
            ExcelForm.SetValue("AL13", NewApplicant.Address_.Index + ", " + NewApplicant.Address_.Region);
            ExcelForm.SetValue("AD14", NewApplicant.Address_.Town + ", " + NewApplicant.Address_.AppAddress);
            ExcelForm.SetValue("AJ18", NewApplicant.HomePhone);
            ExcelForm.SetValue("AK20", DateTime.Today.ToString("dd.MM.yyyy"));
        }
        private void FillExcelProfile(Applicant NewApplicant)
        {
            ExcelForm.SetValue("E3", NewApplicant.RegNumber);
            ExcelForm.SetValue("L10", NewApplicant.SecondName);
            ExcelForm.SetValue("L11", NewApplicant.FirstName);
            ExcelForm.SetValue("L12", NewApplicant.LastName);
            ExcelForm.SetValue("L13", NewApplicant.BirthDate);
            ExcelForm.SetValue("L14", NewApplicant.Citizenship);
            ExcelForm.SetValue("AG10", NewApplicant.Education);
            ExcelForm.SetValue("L13", NewApplicant.BirthDate);
            ExcelForm.SetValue("L14", NewApplicant.Citizenship);
            ExcelForm.SetValue("AI16", NewApplicant.HomePhone);
            if (NewApplicant.Specs[0] != null)
            {
                ExcelForm.SetValue("U6", GetFacultyNumber(NewApplicant.Specs[0].Faculty.ToArray()[0]));
                ExcelForm.SetValue("U7", NewApplicant.Specs[0].Spec);
            }
            if (NewApplicant.NeedDorm == true) ExcelForm.SetValue("L20", "нуждаюсь");
            else ExcelForm.SetValue("L20", "не нуждаюсь");
        }
        private void FillExcelListOfExams(Applicant NewApplicant)
        {
            ExcelForm.SetValue("AL4", NewApplicant.RegNumber);
            if (NewApplicant.Specs[0] != null)
                ExcelForm.SetValue("H6", NewApplicant.Specs[0].Spec.Remove(6, (NewApplicant.Specs[0].Spec.Length - 6)));
            ExcelForm.SetValue("H7", NewApplicant.SecondName);
            ExcelForm.SetValue("H8", NewApplicant.FirstName);
            ExcelForm.SetValue("H9", NewApplicant.LastName);
            ExcelForm.SetValue("L14", DateTime.Today.ToString("dd.MM.yyyy"));
        }
        private void FillExcelReceipt(Applicant NewApplicant)
        {
            ExcelForm.SetValue("I4", NewApplicant.RegNumber);
            ExcelForm.SetValue("E5", NewApplicant.SecondName + " " + NewApplicant.FirstName + " " + NewApplicant.LastName);
            ExcelForm.SetValue("C11", DateTime.Today.ToString("dd.MM.yyyy"));
            ExcelForm.SetValue("B6", NewApplicant.EducationDocument + (NewApplicant.Attest == true ? " (подлинник)" : " (копия)"));
            for (int i =0; i < 5; i++)
            {
                if (NewApplicant.EnteranceExaminations[i]!=null)
                {
                    if (NewApplicant.EnteranceExaminations[i].Ege == true) ExcelForm.SetValue("B7", NewApplicant.EnteranceExaminations[i].TitleAndNum);
                    break;
                }
            }
            ExcelForm.SetValue("B9", "фотографии (6 штук)");
        }
        private void FillSummary(int Fac, Applicant NewApplicant)
        {
            ExcelSummary = new ExcelFunc();
            if (NewApplicant.LearningForm.Equals("очной"))
            {
                ExcelSummary.OpenDocument(ProgramPath + @"Документы (оч)\Сводки\Сводка.xls");
            }
            else if (NewApplicant.LearningForm.Equals("очно-заочной"))
            {
                ExcelSummary.OpenDocument(ProgramPath + @"Документы (оч-заоч)\Сводки\Сводка.xls");
            }
            else
            {
                ExcelSummary.OpenDocument(ProgramPath + @"Документы (др)\Сводки\Сводка.xls");
            }
            if (NewApplicant.Funding.Equals("финансируемые из госбюджета")) ExcelSummary.OpenWorksheet(1);
            else if (NewApplicant.Funding.Equals("с полным возмещением затрат")) ExcelSummary.OpenWorksheet(3);
            ExcelSummary.SetValue("E2", DateTime.Today.ToString("dd.MM.yyyy"));
            int Pos = 13;
            while (true)
            {
                try
                {
                    ExcelSummary.GetValue("B" + Pos.ToString());
                }
                catch (NullReferenceException nre)
                {
                    ExcelSummary.SetValue("B" + Pos.ToString(), NewApplicant.SecondName + " " + NewApplicant.FirstName + " " + NewApplicant.LastName);
                    ExcelSummary.SetValue("C" + Pos.ToString(), NewApplicant.Sex); //cause there are no sex in our program! //oh shi~, now there are some sex in our program...
                    try
                    {
                        ExcelSummary.SetValue("D" + Pos.ToString(), GetExamSummaryMark(NewApplicant).ToString());
                        //ExcelSummary.SetValue("D" + Pos.ToString(), (Convert.ToInt32(NewApplicant.EnteranceExaminations[0].Points) + Convert.ToInt32(NewApplicant.EnteranceExaminations[1].Points) + Convert.ToInt32(NewApplicant.EnteranceExaminations[2].Points)).ToString());
                    }
                    catch
                    {
                        ExcelSummary.SetValue("D" + Pos.ToString(), (0).ToString());
                    }
                    string[] split = NewApplicant.RegNumber.Split(new Char[] { '-' });
                                     
                    ExcelSummary.SetValue("E"/*pic*/ + Pos.ToString(), split[0]);
                    ExcelSummary.SetValue("F"/*ail*/ + Pos.ToString(), split[1]);
                    ExcelSummary.SetValue("G"/*uy*/ + Pos.ToString(), split[2]);
                    for (int i = 0; i < 5; i++)
                    {
                        String PointsCell = "", TypeCell = "";
                        switch (NewApplicant.EnteranceExaminations[i].Subject)
                        {
                            case "Русский язык":
                                PointsCell = "H";
                                TypeCell = "I";
                                break;
                            case "Математика":
                                PointsCell = "J";
                                TypeCell = "K";
                                break;
                            case "Физика":
                                PointsCell = "L";
                                TypeCell = "M";
                                break;
                            case "Иностранный язык":
                                PointsCell = "N";
                                TypeCell = "O";
                                break;
                            case "История":
                                PointsCell = "P";
                                TypeCell = "Q";
                                break;
                            case "Обществознание":
                                PointsCell = "R";
                                TypeCell = "S";
                                break;
                            default: //Эта ошибка появляется, если не выбраны экзамены
                                PointsCell = "Error";
                                TypeCell = "Error";
                                break;
                        }
                        if (PointsCell != "Error" && TypeCell != "Error")
                        {
                            ExcelSummary.SetValue(PointsCell + Pos.ToString(), NewApplicant.EnteranceExaminations[i].Points);
                            if (NewApplicant.EnteranceExaminations[i].Ege == true) ExcelSummary.SetValue(TypeCell + Pos.ToString(), "1");
                            else if (NewApplicant.EnteranceExaminations[i].Olimp == true) ExcelSummary.SetValue(TypeCell + Pos.ToString(), "2");
                        }
                    }
                    ExcelSummary.SetValue("T" + Pos.ToString(), GetSchoolTypeNumber(NewApplicant.SchoolType).ToString());
                    ExcelSummary.SetValue("U" + Pos.ToString(), NewApplicant.SchoolName);
                    ExcelSummary.SetValue("V" + Pos.ToString(), NewApplicant.Address_.Region);
                    ExcelSummary.SetValue("W" + Pos.ToString(), NewApplicant.Address_.Town);
                    DateTime EndDate = new DateTime();
                    if (DateTime.TryParse(NewApplicant.EducationIssuedDate, out EndDate) == true) ExcelSummary.SetValue("X" + Pos.ToString(), EndDate.ToString("yyyy"));
                    ExcelSummary.SetValue("X" + Pos.ToString(), DateTime.Today.ToString("yyyy"));
                    if (NewApplicant.MATICourses == true) ExcelSummary.SetValue("Z" + Pos.ToString(), "1");
                    else ExcelSummary.SetValue("Z" + Pos.ToString(), "0");
                    if (NewApplicant.MATISchool == true) ExcelSummary.SetValue("Y" + Pos.ToString(), "1");
                    else ExcelSummary.SetValue("Y" + Pos.ToString(), "0");
                    if (NewApplicant.Attest == true) ExcelSummary.SetValue("AA" + Pos.ToString(), "1");
                    else ExcelSummary.SetValue("AA" + Pos.ToString(), "0");
                    if (NewApplicant.NeedDorm == true) ExcelSummary.SetValue("AB" + Pos.ToString(), "1");
                    else ExcelSummary.SetValue("AB" + Pos.ToString(), "0");
                    if (NewApplicant.Specs[0] != null) ExcelSummary.SetValue("AC" + Pos.ToString(), NewApplicant.Specs[0].Spec.Remove(6, (NewApplicant.Specs[0].Spec.Length - 6)));
                    if (NewApplicant.Specs[1] != null) ExcelSummary.SetValue("AD" + Pos.ToString(), NewApplicant.Specs[1].Spec.Remove(6, (NewApplicant.Specs[1].Spec.Length - 6)));
                    if (NewApplicant.Specs[2] != null) ExcelSummary.SetValue("AE" + Pos.ToString(), NewApplicant.Specs[2].Spec.Remove(6, (NewApplicant.Specs[2].Spec.Length - 6)));
                    ExcelSummary.SetValue("AG" + Pos.ToString(), DateTime.Today.ToString("dd.MM.yyyy"));
                    if (NewApplicant.Funding.Equals("финансируемые из госбюджета")) ExcelSummary.OpenWorksheet(2);
                    else if (NewApplicant.Funding.Equals("с полным возмещением затрат")) ExcelSummary.OpenWorksheet(4);
                    ExcelSummary.SetValue("J3", DateTime.Today.ToString("dd.MM.yyyy"));
                    ExcelSummary.SetValue("B" + Pos.ToString(), NewApplicant.SecondName + " " + NewApplicant.FirstName + " " + NewApplicant.LastName);
                    if (NewApplicant.Specs[0] != null) ExcelSummary.SetValue("C" + Pos.ToString(), NewApplicant.Specs[0].Spec.Remove(6, (NewApplicant.Specs[0].Spec.Length - 6)));
                    if (NewApplicant.Specs[1] != null) ExcelSummary.SetValue("M" + Pos.ToString(), NewApplicant.Specs[1].Spec.Remove(6, (NewApplicant.Specs[1].Spec.Length - 6)));
                    if (NewApplicant.Specs[2] != null) ExcelSummary.SetValue("W" + Pos.ToString(), NewApplicant.Specs[2].Spec.Remove(6, (NewApplicant.Specs[2].Spec.Length - 6)));
                    if (NewApplicant.Specs[0] != null)
                    {
                        for (int i = 0; i < NewApplicant.Specs[0].Faculty.Count; i++)
                        {
                            switch (i)
                            {
                                case 0:
                                    ExcelSummary.SetValue("D" + Pos.ToString(), GetCathdraCode(NewApplicant.Specs[0].Spec, NewApplicant.Specs[0].Faculty[i]).ToString());
                                    break;
                                case 1:
                                    ExcelSummary.SetValue("E" + Pos.ToString(), GetCathdraCode(NewApplicant.Specs[0].Spec, NewApplicant.Specs[0].Faculty[i]).ToString());
                                    break;
                                case 2:
                                    ExcelSummary.SetValue("F" + Pos.ToString(), GetCathdraCode(NewApplicant.Specs[0].Spec, NewApplicant.Specs[0].Faculty[i]).ToString());
                                    break;
                                case 3:
                                    ExcelSummary.SetValue("G" + Pos.ToString(), GetCathdraCode(NewApplicant.Specs[0].Spec, NewApplicant.Specs[0].Faculty[i]).ToString());
                                    break;
                                case 4:
                                    ExcelSummary.SetValue("H" + Pos.ToString(), GetCathdraCode(NewApplicant.Specs[0].Spec, NewApplicant.Specs[0].Faculty[i]).ToString());
                                    break;
                                case 5:
                                    ExcelSummary.SetValue("I" + Pos.ToString(), GetCathdraCode(NewApplicant.Specs[0].Spec, NewApplicant.Specs[0].Faculty[i]).ToString());
                                    break;
                                case 6:
                                    ExcelSummary.SetValue("J" + Pos.ToString(), GetCathdraCode(NewApplicant.Specs[0].Spec, NewApplicant.Specs[0].Faculty[i]).ToString());
                                    break;
                                case 7:
                                    ExcelSummary.SetValue("K" + Pos.ToString(), GetCathdraCode(NewApplicant.Specs[0].Spec, NewApplicant.Specs[0].Faculty[i]).ToString());
                                    break;
                                case 8:
                                    ExcelSummary.SetValue("L" + Pos.ToString(), GetCathdraCode(NewApplicant.Specs[0].Spec, NewApplicant.Specs[0].Faculty[i]).ToString());
                                    break;
                            }
                        }
                    }
                    if (NewApplicant.Specs[1] != null)
                    {
                        for (int i = 0; i < NewApplicant.Specs[1].Faculty.Count; i++)
                        {
                            switch (i)
                            {
                                case 0:
                                    ExcelSummary.SetValue("N" + Pos.ToString(), GetCathdraCode(NewApplicant.Specs[1].Spec, NewApplicant.Specs[1].Faculty[i]).ToString());
                                    break;
                                case 1:
                                    ExcelSummary.SetValue("O" + Pos.ToString(), GetCathdraCode(NewApplicant.Specs[1].Spec, NewApplicant.Specs[1].Faculty[i]).ToString());
                                    break;
                                case 2:
                                    ExcelSummary.SetValue("P" + Pos.ToString(), GetCathdraCode(NewApplicant.Specs[1].Spec, NewApplicant.Specs[1].Faculty[i]).ToString());
                                    break;
                                case 3:
                                    ExcelSummary.SetValue("Q" + Pos.ToString(), GetCathdraCode(NewApplicant.Specs[1].Spec, NewApplicant.Specs[1].Faculty[i]).ToString());
                                    break;
                                case 4:
                                    ExcelSummary.SetValue("R" + Pos.ToString(), GetCathdraCode(NewApplicant.Specs[1].Spec, NewApplicant.Specs[1].Faculty[i]).ToString());
                                    break;
                                case 5:
                                    ExcelSummary.SetValue("S" + Pos.ToString(), GetCathdraCode(NewApplicant.Specs[1].Spec, NewApplicant.Specs[1].Faculty[i]).ToString());
                                    break;
                                case 6:
                                    ExcelSummary.SetValue("T" + Pos.ToString(), GetCathdraCode(NewApplicant.Specs[1].Spec, NewApplicant.Specs[1].Faculty[i]).ToString());
                                    break;
                                case 7:
                                    ExcelSummary.SetValue("U" + Pos.ToString(), GetCathdraCode(NewApplicant.Specs[1].Spec, NewApplicant.Specs[1].Faculty[i]).ToString());
                                    break;
                                case 8:
                                    ExcelSummary.SetValue("V" + Pos.ToString(), GetCathdraCode(NewApplicant.Specs[1].Spec, NewApplicant.Specs[1].Faculty[i]).ToString());
                                    break;
                            }
                        }
                    }
                    if (NewApplicant.Specs[2] != null)
                    {
                        for (int i = 0; i < NewApplicant.Specs[2].Faculty.Count; i++)
                        {
                            switch (i)
                            {
                                case 0:
                                    ExcelSummary.SetValue("X" + Pos.ToString(), GetCathdraCode(NewApplicant.Specs[2].Spec, NewApplicant.Specs[2].Faculty[i]).ToString());
                                    break;
                                case 1:
                                    ExcelSummary.SetValue("Y" + Pos.ToString(), GetCathdraCode(NewApplicant.Specs[2].Spec, NewApplicant.Specs[2].Faculty[i]).ToString());
                                    break;
                                case 2:
                                    ExcelSummary.SetValue("Z" + Pos.ToString(), GetCathdraCode(NewApplicant.Specs[2].Spec, NewApplicant.Specs[2].Faculty[i]).ToString());
                                    break;
                                case 3:
                                    ExcelSummary.SetValue("AA" + Pos.ToString(), GetCathdraCode(NewApplicant.Specs[2].Spec, NewApplicant.Specs[2].Faculty[i]).ToString());
                                    break;
                                case 4:
                                    ExcelSummary.SetValue("AB" + Pos.ToString(), GetCathdraCode(NewApplicant.Specs[2].Spec, NewApplicant.Specs[2].Faculty[i]).ToString());
                                    break;
                                case 5:
                                    ExcelSummary.SetValue("AC" + Pos.ToString(), GetCathdraCode(NewApplicant.Specs[2].Spec, NewApplicant.Specs[2].Faculty[i]).ToString());
                                    break;
                                case 6:
                                    ExcelSummary.SetValue("AD" + Pos.ToString(), GetCathdraCode(NewApplicant.Specs[2].Spec, NewApplicant.Specs[2].Faculty[i]).ToString());
                                    break;
                                case 7:
                                    ExcelSummary.SetValue("AE" + Pos.ToString(), GetCathdraCode(NewApplicant.Specs[2].Spec, NewApplicant.Specs[2].Faculty[i]).ToString());
                                    break;
                                case 8:
                                    ExcelSummary.SetValue("AF" + Pos.ToString(), GetCathdraCode(NewApplicant.Specs[2].Spec, NewApplicant.Specs[2].Faculty[i]).ToString());
                                    break;
                            }
                        }
                    }
                    break;
                }
                Pos++;
            }

            ExcelSummary.CloseDocument();
            MessageBox.Show("Генерация документов для абитуриента №" + NewApplicant.RegNumber + " завершена", "Анкета");
        }

        private void GenerateMasterDocuments(Applicant NewApplicant, Master NewMaster)
        {
            if (NewApplicant.LearningForm.Equals("очной"))
            {
                File.Copy(ProgramPath + @"Шаблоны\MasterTemplate.xls", ProgramPath + @"Документы (оч)\" + NewApplicant.SecondName + " " + NewApplicant.RegNumber + ".xls");
                ExcelForm = new ExcelFunc();
                ExcelForm.OpenDocument(ProgramPath + @"Документы (оч)\" + NewApplicant.SecondName + " " + NewApplicant.RegNumber + ".xls");
            }
            else if (NewApplicant.LearningForm.Equals("очно-заочной"))
            {
                File.Copy(ProgramPath + @"Шаблоны\MasterTemplate.xls", ProgramPath + @"Документы (оч-заоч)\" + NewApplicant.SecondName + " " + NewApplicant.RegNumber + ".xls");
                ExcelForm = new ExcelFunc();
                ExcelForm.OpenDocument(ProgramPath + @"Документы (оч-заоч)\" + NewApplicant.SecondName + " " + NewApplicant.RegNumber + ".xls");
            }
            else
            {
                File.Copy(ProgramPath + @"Шаблоны\MasterTemplate.xls", ProgramPath + @"Документы (др)\" + NewApplicant.SecondName + " " + NewApplicant.RegNumber + ".xls");
                ExcelForm = new ExcelFunc();
                ExcelForm.OpenDocument(ProgramPath + @"Документы (др)\" + NewApplicant.SecondName + " " + NewApplicant.RegNumber + ".xls");
            }
            ExcelForm.OpenWorksheet(1);
            FillMasterForm(NewApplicant, NewMaster);
            ExcelForm.OpenWorksheet(2);
            FillMasterRequest(NewApplicant, NewMaster);
            ExcelForm.OpenWorksheet(3);
            FillMasterListOfExams(NewApplicant, NewMaster);
            ExcelForm.OpenWorksheet(4);
            FillMasterProfile(NewApplicant, NewMaster);
            ExcelForm.OpenWorksheet(5);
            FillMasterReceipt(NewApplicant, NewMaster);
            int Fac;
            if (NewMaster.Specs[0] != null)
            {
                String SFac = GetFacultyNumber(NewMaster.Specs[0].Faculty.ToArray()[0]);
                Int32.TryParse(SFac, out Fac);
                FillSummary(Fac, NewApplicant);
            }
            ExcelForm.CloseDocument();
        }
        private void FillMasterForm(Applicant NewApplicant, Master NewMaster)
        {
            ExcelForm.SetValue("B2", NewApplicant.RegNumber);
            ExcelForm.SetValue("B3", NewApplicant.SecondName);
            ExcelForm.SetValue("B4", NewApplicant.FirstName);
            ExcelForm.SetValue("B5", NewApplicant.LastName);
            ExcelForm.SetValue("B7", NewApplicant.BirthDate);
            ExcelForm.SetValue("B8", NewApplicant.BirthPlace);
            ExcelForm.SetValue("B17", NewApplicant.Address_.Index + ", " + NewApplicant.Address_.Region + ", " + NewApplicant.Address_.Town + ", " + NewApplicant.Address_.AppAddress);
            ExcelForm.SetValue("B6", NewApplicant.Citizenship);
            ExcelForm.SetValue("B11", NewApplicant.Passport);
            ExcelForm.SetValue("B12", NewApplicant.Serial);
            ExcelForm.SetValue("B13", NewApplicant.Number);
            ExcelForm.SetValue("B14", NewApplicant.PassportIssuedDate);
            ExcelForm.SetValue("B19", NewApplicant.HomePhone);
            ExcelForm.SetValue("B22", NewApplicant.LearningForm);
            ExcelForm.SetValue("B23", NewApplicant.Funding);
            ExcelForm.SetValue("B25", NewMaster.University);
            ExcelForm.SetValue("B26", NewMaster.Diploma);
            if (NewApplicant.Specs[0] != null)
            {
                ExcelForm.SetValue("B20", NewMaster.Specs[0].Spec.Insert(6, ".068"));
                if (NewApplicant.Specs[0].Spec != "Неизвестно") ExcelForm.SetValue("B21", "№" + GetFacultyNumber(NewMaster.Specs[0].Faculty.ToArray()[0]));
            }
            ExcelForm.SetValue("B28", NewApplicant.Facilities);
            if (NewApplicant.NeedDorm == true) ExcelForm.SetValue("B30", "нуждаюсь");
            else ExcelForm.SetValue("B30", "не нуждаюсь");
            ExcelForm.SetValue("B38", NewApplicant.IssuedUniversityEducation);
            ExcelForm.SetValue("B37", DateTime.Today.ToString("dd.MM.yyyy"));
        }
        private void FillMasterRequest(Applicant NewApplicant, Master NewMaster)
        {
            ExcelForm.SetValue("AT1", NewApplicant.RegNumber);
            ExcelForm.SetValue("I5", NewApplicant.SecondName);
            ExcelForm.SetValue("I6", NewApplicant.FirstName);
            ExcelForm.SetValue("I7", NewApplicant.LastName);
            ExcelForm.SetValue("I8", NewApplicant.BirthDate);
            ExcelForm.SetValue("I9", NewApplicant.BirthPlace);
            ExcelForm.SetValue("P11", NewApplicant.Address_.Index + ", " + NewApplicant.Address_.Region + ", " + NewApplicant.Address_.Town);
            ExcelForm.SetValue("A12", NewApplicant.Address_.AppAddress);
            ExcelForm.SetValue("AL5", NewApplicant.Citizenship);
            ExcelForm.SetValue("AR6", NewApplicant.Passport);
            ExcelForm.SetValue("AK7", NewApplicant.Serial);
            ExcelForm.SetValue("AU7", NewApplicant.Number);
            ExcelForm.SetValue("Z9", NewApplicant.PassportIssuedDate);
            ExcelForm.SetValue("AL12", NewApplicant.HomePhone);
            ExcelForm.SetValue("C17", NewApplicant.LearningForm);
            ExcelForm.SetValue("Z17", NewApplicant.Funding);
            ExcelForm.SetValue("G23", NewMaster.University);
            ExcelForm.SetValue("G24", NewMaster.Diploma);
            if (NewApplicant.Specs[0] != null)
            {
                ExcelForm.SetValue("G15", NewMaster.Specs[0].Spec.Insert(6, ".068"));
                if (NewApplicant.Specs[0].Spec != "Неизвестно") ExcelForm.SetValue("G16", "№" + GetFacultyNumber(NewMaster.Specs[0].Faculty.ToArray()[0]));
            }
            ExcelForm.SetValue("V25", NewApplicant.Facilities);
            if (NewApplicant.NeedDorm == true) ExcelForm.SetValue("I27", "нуждаюсь");
            else ExcelForm.SetValue("I27", "не нуждаюсь");
            ExcelForm.SetValue("V20", NewApplicant.IssuedUniversityEducation);
            ExcelForm.SetValue("B40", DateTime.Today.ToString("dd.MM.yyyy"));
            ExcelForm.SetValue("M40", NewApplicant.SecondName + " " + NewApplicant.FirstName + " " + NewApplicant.LastName);
        }
        private void FillMasterListOfExams(Applicant NewApplicant, Master NewMaster)
        {
            ExcelForm.SetValue("AL4", NewApplicant.RegNumber);
            ExcelForm.SetValue("H8", NewApplicant.SecondName);
            ExcelForm.SetValue("H9", NewApplicant.FirstName);
            ExcelForm.SetValue("H10", NewApplicant.LastName);
            if (NewApplicant.Specs[0] != null)
            {
                ExcelForm.SetValue("H6", NewMaster.Specs[0].Spec.Insert(6, ".068"));
                if (NewApplicant.Specs[0].Spec != "Неизвестно") ExcelForm.SetValue("H5", "№" + GetFacultyNumber(NewMaster.Specs[0].Faculty.ToArray()[0]));
            }
            ExcelForm.SetValue("L15", DateTime.Today.ToString("dd.MM.yyyy"));
        }
        private void FillMasterProfile(Applicant NewApplicant, Master NewMaster)
        {
            ExcelForm.SetValue("E3", NewApplicant.RegNumber);
            ExcelForm.SetValue("L10", NewApplicant.SecondName);
            ExcelForm.SetValue("L11", NewApplicant.FirstName);
            ExcelForm.SetValue("L12", NewApplicant.LastName);
            ExcelForm.SetValue("L13", NewApplicant.BirthDate);
            ExcelForm.SetValue("AL13", NewApplicant.Address_.Index + ", " + NewApplicant.Address_.Region + ", " + NewApplicant.Address_.Town + ", " + NewApplicant.Address_.AppAddress);
            ExcelForm.SetValue("L14", NewApplicant.Citizenship);
            ExcelForm.SetValue("Z11", NewApplicant.Education);
            ExcelForm.SetValue("AL15", NewApplicant.HomePhone);
            if (NewApplicant.Specs[0] != null)
            {
                ExcelForm.SetValue("U7", NewMaster.Specs[0].Spec.Insert(6, ".068"));
                if (NewApplicant.Specs[0].Spec != "Неизвестно") ExcelForm.SetValue("U6", "№" + GetFacultyNumber(NewMaster.Specs[0].Faculty.ToArray()[0]));
            }
        }
        private void FillMasterReceipt(Applicant NewApplicant, Master NewMaster)
        {
            ExcelForm.SetValue("I4", NewApplicant.RegNumber);
            ExcelForm.SetValue("E5", NewApplicant.SecondName + " " + NewApplicant.FirstName + " " + NewApplicant.LastName);
            ExcelForm.SetValue("C11", DateTime.Today.ToString("dd.MM.yyyy"));
        }


        private String[] LoadSubjects()
        {
            String[] Subjects = 
            {
                
                "Русский язык",
                "Математика",
                "Физика",
                "Обществознание",
                "Иностранный язык",
                "История"

            };
            return Subjects;
        }
        private String[] LoadFunding()
        {
            String[] Subjects = 
            {
                
                "финансируемые из госбюджета",
                "с полным возмещением затрат"

            };
            return Subjects;
        }
        private String[] LoadLearningForm()
        {
            String[] Subjects = 
            {
                
                "очной",
                "очно-заочной",
                "заочной"
            };
            return Subjects;
        }
        private String[] LoadSemesters()
        {
            String[] Subjects = 
            {
                "второго",
                "третьего",
                "четвертого",
                "пятого",
                "шестого",
                "седьмого",
                "восьмого",
                "девятого",
                "десятого",
                "одиннадцатого",
                "двенадцатого"
            };
            return Subjects;
        }
        private String[] LoadSpecs()
        {
            String[] Specs = 
            {
                "010300 Фундаментальная информатика и информационные", 
                "010400 Прикладная математика и информатика", 
                "011200 Физика", 
                "031600 Реклама и связи с общественностью", 
                "034700 Документоведение и архивоведение", 
                "035700 Лингвистика", 
                "040700 Организация работы с молодежью",  
                "080100 Экономика", 
                "080200 Менеджмент", 
                "080400 Управление персоналом", 
                "080500 Бизнес информатика", 
                "081100 Государственное и муниципальное управление", 
                "150100 Материаловедение и технологии материалов",  
                "150400 Металургия",  
                "151600 Прикладная механика", 
                "160100 Авиастроение", 
                "160400 Ракетные комплексы и космонавтика", 
                "160700 Двигатели летательных аппаратов", 
                "162110 Испытание летательных аппаратов", 
                "200100 Приборостроение", 
                "200500 Лазерная техника и лазерные технологии",  
                "201000 Биотехнические системы и технологии", 
                "211000 Конструирование и технология электронных средств", 
                "220100 Системный анализ и управление", 
                "220700 Автоматизация технологических процессов и производств", 
                "221400 Управление качеством", 
                "221700 Стандартизация и метрология", 
                "222000 Инноватика", 
                "222900 Нанотехнологии и микросистемная техника", 
                "230100 Информатика и вычислительная техника", 
                "280700 Техносферная безопасность",
                ""//для отмены выбора направления
            };
            Array.Sort(Specs);
            return Specs;
        }
        public static Dictionary<String, List<String>> GetFaculties()
        {
            Dictionary<String, List<String>> faculties = new Dictionary<String, List<String>>()
            {
                { "010300 Фундаментальная информатика и информационные", new List<String>() {  "№5 Прикладная математика", "№1 Электроника и информатика" } },
                { "010400 Прикладная математика и информатика", new List<String>() { "№5 Прикладная математика" } },
                { "011200 Физика", new List<String>{ "№1 Физика" } },
                { "031600 Реклама и связи с общественностью", new List<String>{ "№7 Философия и соц. коммуникации", "№7 Культурология, ист., молод.политика и реклама"} },
                { "034700 Документоведение и архивоведение", new List<String>{ "№7 Документ. обесп. упр. и прик.лингвистика"} },
                { "035700 Лингвистика", new List<String>{ "№7 Проф.подготовка по ин.языкам" } },
                { "040700 Организация работы с молодежью", new List<String>{ "№7 Культурология, ист., молод.политика и реклама" } },
                { "080100 Экономика", new List<String>{ "№6 Экономика", "№6 Учет, анализ и аудит" } },
                { "080200 Менеджмент", new List<String>{ "№6 Производственный менеджмент", "№1 Технология ОМД", "№6 Финансовый менеджмент", "№6 Маркетинг", "№14 Экономика и управление"} },
                { "080400 Управление персоналом", new List<String>{ "№7 Социология и упр. персоналом" } },
                { "080500 Бизнес информатика", new List<String>{ "№6 Проектирование вычислительных комплексов" } },
                { "081100 Государственное и муниципальное управление", new List<String>{ "№7 ГМУ, правоведение и психология" } },
                { "150100 Материаловедение и технологии материалов", new List<String>{ "№4 Материаловедение и технол. обр-ки материалов", "№4 Технология переработки немет. материалов", "№4 Общая химия, физика и химия КМ", "№14 Технология и автом. обр-ки материалов" } },
                { "150400 Металургия", new List<String>{ "№1 Технология ОМД", "№1 САПР и технологии литейного производства", "№1 Технология сварочного производства", "№1 Порошковая металлургия, КМ и покрытия" } },
                { "151600 Прикладная механика", new List<String>{ "№5 Прикладная и вычислительная механика" } },
                { "160100 Авиастроение", new List<String>{ "№2 Технология проект.и эксплуатация ЛА" } },
                { "160400 Ракетные комплексы и космонавтика", new List<String>{ "№2 Технология производства ЛА", "№2 Испытания ЛА", "№2 Стартовые комплексы", "№2 Спутники и разгонные блоки" } },
                { "160700 Двигатели летательных аппаратов", new List<String>{ "№2 Технология производства двигателей ЛА", "№2 Двигатели ЛА и теплотехника", "№14 Технология производства авиационных двигателей" } },
                { "162110 Испытание летательных аппаратов", new List<String>{ "№2 Испытание ЛА" } },
                { "200100 Приборостроение", new List<String>{ "№3  Технология производства приборов и СУЛА" } },
                { "200500 Лазерная техника и лазерные технологии", new List<String>{ "№3 Технол.обр-ки матер. потоками высоких энергий" } },
                { "201000 Биотехнические системы и технологии", new List<String>{ "№4 Материаловедение и технол. обр-тки материалов" } },
                { "211000 Конструирование и технология электронных средств", new List<String>{ "№3 Радиоэлектроника, телекоммуникации и нанотехнологии", "  " } },
                { "220100 Системный анализ и управление", new List<String>{ "№3 Кибернетика", "№3 Эргономика и инф.-измерит. системы" } },
                { "220700 Автоматизация технологических процессов и производств", new List<String>{ "№1 Технология ОМД", "№14 Технология и автоматизация ОМ" } },
                { "221400 Управление качеством", new List<String>{ "№6 Управление качеством и сертификация" } },
                { "221700 Стандартизация и метрология", new List<String>{ "№5 Механика машин и механизмов" } },
                { "222000 Инноватика", new List<String>{ "№7 Управление инновациями" } },
                { "222900 Нанотехнологии и микросистемная техника", new List<String>{ "№4 Общая химия, физика и химия КМ", "№3 Технол. обр-ки матер. потоками высоких энергий" } },
                { "230100 Информатика и вычислительная техника", new List<String>{ "№2 Технологии интегрированных АС", "№2 Космические телекоммуникации", "№2 Испытания ЛА", "№3 Технология производства приборов и СУЛА", "№3 Технология производства приборов  и СУЛА (Раменское)", "№3 Информационные технологии", "№3 Интернет - технологии", "№5 Системное моделирование и инженерная графика", "№6 Проектирование вычислительных комплексов", "№14 Моделирование систем и информационные технологии" } },
                { "280700 Техносферная безопасность", new List<String>{ "№6 Пром.экология и безопасность производства", "№6 Природная и техногенная безопасность и УР", "№6 Космический мониторинг" } },
            };

            return faculties;
        }
        private String[] LoadSpecsForMasters()
        {
            String[] Specs = 
            {
                "010300 Фундаментальная информатика и информационные", 
                "010400 Прикладная математика и информатика",  
                "080100 Экономика", 
                "080200 Менеджмент", 
                "150100 Материаловедение и технологии материалов",    
                "151600 Прикладная механика", 
                "160100 Авиастроение", 
                "160400 Ракетные комплексы и космонавтика", 
                "160700 Двигатели летательных аппаратов", 
                "211000 Конструирование и технология электронных средств", 
                "220100 Системный анализ и управление", 
                "222000 Инноватика", 
                "230100 Информатика и вычислительная техника", 
                "280700 Техносферная безопасность",
                ""//для отмены выбора направления
            };
            Array.Sort(Specs);
            return Specs;
        }
        public static string GetFacultyNumber(String FullName)
        {
            Regex Reg = new Regex(@"\d+");
            Match matches = Reg.Match(FullName);
            return matches.Value;
        }
        public static int GetCathdraCode(String spec, String cathedra)
        {
            var Faculties = GetFaculties();
            List<String> Cathedras = Faculties[spec];
            int CathedraCode = Cathedras.IndexOf(cathedra) + 1;
            return CathedraCode;
        }
        public static int GetExamMark(Applicant NewApplicant, String ExamName)
        {
            int res = 0;
            bool tp_res = false;
            for (int i = 0; i < NewApplicant.EnteranceExaminations.Length; i++)
            {
                if (NewApplicant.EnteranceExaminations[i] != null)
                {
                    if (NewApplicant.EnteranceExaminations[i].Subject == ExamName)
                    {

                        tp_res = Int32.TryParse(NewApplicant.EnteranceExaminations[i].Points, out res);
                        if (tp_res == true)
                        {
                            //MessageBox.Show(ExamName, res.ToString());
                            return res;
                        }
                    }
                }
            }
            return 0;
        }
        public static int GetExamSummaryMark(Applicant NewApplicant)
        {
            String FirstPrior = null;
            if (NewApplicant.Specs[0] != null) FirstPrior = NewApplicant.Specs[0].Spec;
            else FirstPrior = "Error";
            switch (FirstPrior)
            {
                    //Nees optimization here! Right now!
                case "034700 Документоведение и архивоведение":
                case "035700 Лингвистика":
                case "040700 Организация работы с молодежью":
                    return (GetExamMark(NewApplicant, "Русский язык") + GetExamMark(NewApplicant, "Обществознание") + GetExamMark(NewApplicant, "История"));
                case "080100 Экономика":
                case "080200 Менеджмент":
                case "080400 Управление персоналом":
                case "080500 Бизнес информатика":
                case "081100 Государственное и муниципальное управление":
                case "031600 Реклама и связи с общественностью":
                    return (GetExamMark(NewApplicant, "Русский язык") + GetExamMark(NewApplicant, "История") + GetExamMark(NewApplicant, "Обществознание"));
                case "010300 Фундаментальная информатика и информационные":
                case "010400 Прикладная математика и информатика":
                case "011200 Физика":
                case "150100 Материаловедение и технологии материалов":
                case "150400 Металургия":
                case "151600 Прикладная механика":
                case "160100 Авиастроение":
                case "160400 Ракетные комплексы и космонавтика":
                case "160700 Двигатели летательных аппаратов":
                case "162110 Испытание летательных аппаратов":
                case "200100 Приборостроение":
                case "200500 Лазерная техника и лазерные технологии":
                case "201000 Биотехнические системы и технологии":
                case "211000 Конструирование и технология электронных средств":
                case "220100 Системный анализ и управление":
                case "220700 Автоматизация технологических процессов и производств":
                case "221400 Управление качеством":
                case "221700 Стандартизация и метрология":
                case "222000 Инноватика":
                case "222900 Нанотехнологии и микросистемная техника":
                case "230100 Информатика и вычислительная техника":
                case "280700 Техносферная безопасность":
                    return (GetExamMark(NewApplicant, "Русский язык") + GetExamMark(NewApplicant, "Математика") + GetExamMark(NewApplicant, "Физика"));
                default:
                    return 0;
            }
        }

        public static void GenerateUSECheck(Applicant NewApplicant, string FacultyNumber)
        {
            System.IO.StreamWriter file = new System.IO.StreamWriter(ProgramPath + FacultyNumber + @" form " + DateTime.Today.ToString("dd.MM.yyyy") + ".csv");
            file.Write(NewApplicant.Serial.Trim().Replace(" ", ""));
            file.Write("%");
            file.Write(NewApplicant.Number.Trim().Replace(" ", ""));
            file.Write("%");
            bool Rus = false, Mth = false, Phis = false, Obsh = false, Eng = false, Hist = false; //Экзамены
            for (int i = 0; i < 5; i++)
            {

                
            }
        }
        private void Propreties_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var PropCombo = (ComboBox)sender;
            String spec = PropCombo.SelectedValue as String;
            if (spec == null) spec = "";
            Faculty FacultyDialog;
            bool? FacultyDialogResult = false;

            // If this spec already choosen
            foreach(Specialization Spec in Specs)
            {
                if(Spec != null && Spec.Spec.Equals(spec))
                {
                    MessageBox.Show("Одну и ту же специальность нельзя выбирать дважды", "Предупреждение", MessageBoxButton.OK, MessageBoxImage.Warning);
                    PropCombo.SelectedItem = null;
                    return;
                }
            }

            // If something choosen, select facs
            if (spec.Equals(""))
                return;
            else
            {
                FacultyDialog = new Faculty(spec);
                FacultyDialogResult = FacultyDialog.ShowDialog();
            }
            if ((bool)FacultyDialogResult)
            {
                if (PropCombo.Name.Equals("ControlFirstPriority"))
                    Specs[0] = new Specialization(spec, FacultyDialog.Output);
                else if (PropCombo.Name.Equals("ControlSecondPriority"))
                    Specs[1] = new Specialization(spec, FacultyDialog.Output);
                else if (PropCombo.Name.Equals("ControlThirdPriority"))
                    Specs[2] = new Specialization(spec, FacultyDialog.Output);
            }
            else
                PropCombo.SelectedItem = null;

        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
        }

        private void ControlMagistrSpec_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var PropCombo = (ComboBox)sender;
            String spec = PropCombo.SelectedValue as String;
            Faculty FacultyDialog;

            // If something choosen, select facs
            if (spec.Equals(""))
                return;
            else
            {
                FacultyDialog = new Faculty(spec);
                FacultyDialog.ShowDialog();
            }
            if (PropCombo.Name.Equals("ControlMagistrSpec"))
            {
                Specs[0] = new Specialization(spec, FacultyDialog.Output);
                Specs[1] = null;
                Specs[2] = null;
            }
        }
        private void ControlMagistrProof_Checked(object sender, RoutedEventArgs e)
        {

        }
        private void ControlMagistrProof_Click(object sender, RoutedEventArgs e)
        {
            ControlMagistrSpec.IsEnabled = !ControlMagistrSpec.IsEnabled;
            ControlMagistrDiploma.IsEnabled = !ControlMagistrDiploma.IsEnabled;
            ControlMagistrUniversity.IsEnabled = !ControlMagistrUniversity.IsEnabled;
        }
        private void ControlContinueHigherEducation_Click(object sender, RoutedEventArgs e)
        {
            ControlSemesterNum.IsEnabled = !ControlSemesterNum.IsEnabled;
            ControlDismissalOrder.IsEnabled = !ControlDismissalOrder.IsEnabled;
            ControlDismissalOrderDate.IsEnabled = !ControlDismissalOrderDate.IsEnabled;
            ControlDismissalReason.IsEnabled = !ControlDismissalReason.IsEnabled;
        }
        private void ControlSemesterNum_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

        }
    }
}
