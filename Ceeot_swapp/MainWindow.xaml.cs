using System;
using System.Collections.Generic;
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
using System.Windows.Navigation;
using System.Windows.Shapes;
using MahApps.Metro.Controls;
using MahApps.Metro.Controls.Dialogs ;
using WinForms = System.Windows.Forms;
using SubBasin = Ceeot_swapp.SwattProject.SubBasin;
using HRU = Ceeot_swapp.SwattProject.HRU;

using System.Collections.ObjectModel;

namespace Ceeot_swapp
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : MetroWindow
    {
        // project manager
        ProjectManager projectManager;
        NewProjectDialog newProjectDialog;
 
        public MainWindow()
        {
            InitializeComponent();

            projectManager = new ProjectManager();
            //this.DataContext = this.projectManager.CurrentProject;
        }

        private void selectSubBasin(object sender, RoutedEventArgs e)
        {
            string basinFileName = ((CheckBox)sender).Content.ToString();

            bool IsActive = (bool)(sender as CheckBox).IsChecked;
            if (IsActive)
            {
                int n = this.projectManager.CurrentProject.SubBasins.Count;
                for (int i = 0; i < n; i++)
                {
                    var basin = projectManager.CurrentProject.SubBasins[i];
                    if (basin.Name == basinFileName)
                    {
                        basin.Selected = true;
                    }
                    projectManager.CurrentProject.SubBasins[i] = basin;
                }
                var hrus = projectManager.CurrentProject.SelectedSubBasinHrus;
                all_landuse_list.ItemsSource = new ObservableCollection<HRU>(hrus);
            } else {
                // If sub basin is unchecked remove its landuse from land use list
                foreach( var hru in projectManager.CurrentProject.SelectedSubBasinHrus) {
                    if (basinFileName == hru.SubBasin)
                    {
                        projectManager.CurrentProject.SelectedSubBasinHrus.Remove(hru);
                    }
                }
                all_landuse_list.ItemsSource = new ObservableCollection<HRU>(projectManager.CurrentProject.SelectedSubBasinHrus);
            }
        }

        public void openNewProjectDialog(object sender, RoutedEventArgs e)
        {
            this.newProjectDialog = new NewProjectDialog(this.projectManager);
            this.newProjectDialog.Closing += this.setupProjectUI;
            this.newProjectDialog.ShowDialog();
        }

        public void openExistingProject(object sender, RoutedEventArgs e)
        {
            using (var fbd = new WinForms.FolderBrowserDialog())
            {
                WinForms.DialogResult result = fbd.ShowDialog();

                if (result == WinForms.DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
                {
                    projectManager.readProject(fbd.SelectedPath);
                    all_sub_basins_list.ItemsSource 
                        = new ObservableCollection<SubBasin>(projectManager.CurrentProject.SubBasins);
                    all_landuse_list.ItemsSource 
                        = new ObservableCollection<HRU>(projectManager.CurrentProject.SelectedSubBasinHrus);
                }
            }
        }

        public void setupProjectUI(object sender, System.ComponentModel.CancelEventArgs e)
        {
            Console.WriteLine("Setting up ui");
            all_sub_basins_list.ItemsSource
                = new ObservableCollection<SubBasin>(projectManager.CurrentProject.SubBasins);
            all_landuse_list.ItemsSource
                = new ObservableCollection<HRU>(projectManager.CurrentProject.SelectedSubBasinHrus);
        }

        public void createApex(object sender, RoutedEventArgs e) {

        }

        public void createApexControl(object sender, RoutedEventArgs e)
        {

        }

        public void createApexOperations(object sender, RoutedEventArgs e)
        {

        }

        public void createApexSubAreas(object sender, RoutedEventArgs e)
        {

        }

        public void createApexSoils(object sender, RoutedEventArgs e)
        {

        }


        public void createApexSite(object sender, RoutedEventArgs e)
        {

        }


        public void createApexWeather(object sender, RoutedEventArgs e)
        {

        }

        public void createApexWeatherStation(object sender, RoutedEventArgs e)
        {

        }

        protected override void OnClosing(System.ComponentModel.CancelEventArgs e)
        {
            //do my stuff before closing
            base.OnClosing(e);
        }
    }
}
