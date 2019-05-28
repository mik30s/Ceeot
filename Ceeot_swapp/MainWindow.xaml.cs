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

using CEEOT_dll;

using System.Collections.ObjectModel;

namespace Ceeot_swapp
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : MetroWindow {
        // project manager
        ProjectManager projectManager;
        NewProjectDialog newProjectDialog;
 
        public MainWindow() {
            InitializeComponent();

            projectManager = new ProjectManager();
            //this.DataContext = this.projectManager.CurrentProject;
        }

        private void selectSubBasin(object sender, RoutedEventArgs e) {
            string basinFileName = ((CheckBox)sender).Content.ToString();

            bool IsActive = (bool)(sender as CheckBox).IsChecked;
            if (IsActive) {
                int n = this.projectManager.CurrentProject.SubBasins.Count;
                for (int i = 0; i < n; i++) {
                    var basin = projectManager.CurrentProject.SubBasins[i];
                    if (basin.Name == basinFileName) {
                        basin.Selected = true;
                    }
                    projectManager.CurrentProject.SubBasins[i] = basin;
                }
                var hrus = projectManager.CurrentProject.SelectedSubBasinHrus;
                all_landuse_list.ItemsSource = new ObservableCollection<HRU>(hrus);
            } else {
                // If sub basin is unchecked remove its landuse from land use list
                foreach( var hru in projectManager.CurrentProject.SelectedSubBasinHrus) {
                    if (basinFileName == hru.SubBasin) {
                        projectManager.CurrentProject.SelectedSubBasinHrus.Remove(hru);
                    }
                }
                all_landuse_list.ItemsSource = 
                    new ObservableCollection<HRU>(projectManager.CurrentProject.SelectedSubBasinHrus);
            }
        }

        private void selectHru(object sender, RoutedEventArgs e) {
            var runBlocks = ((sender as CheckBox).Content as TextBlock).Inlines;
            // extract hru code and hru description
            string[] hruText = new string[3];
            int i = 0, j = 0;
            foreach (var r in runBlocks) { 
                if (i % 4 == 0) { 
                    hruText[j] = (r as Run).Text.ToString();
                    j++;
                }
                i++;
            }
            bool IsActive = (bool)(sender as CheckBox).IsChecked;

            // extract hru and basin name
            var basins = this.projectManager.CurrentProject.SubBasins;
            // fill needed hru values for basin search
            var basinName = hruText[2];
            var hruDescription = hruText[1];
            var hruCropCodeString = hruText[0];
                
            // find the basin
            for (int n = 0; n < this.projectManager.CurrentProject.SubBasins.Count; n++) {
                // if the right basin is found
                var basin = this.projectManager.CurrentProject.SubBasins[n];
                if (basin.Name == basinName) {
                    // find the right hru
                    for (int m = 0; m < basin.Hrus.Count; m++) {    
                        var hru = basin.Hrus[m];
                        // Convert code string to code enum value
                        CropCodes.Code code = (CropCodes.Code)Enum.Parse(typeof(CropCodes.Code), hruCropCodeString);
                        if (hru.Code == code && hru.Description == hruDescription) {
                            if (IsActive) hru.Selected = true;
                            else hru.Selected = false;
                        }
                        basin.Hrus[m] = hru;
                    }
                }
                this.projectManager.CurrentProject.SubBasins[n] = basin;
            }
            all_landuse_list.ItemsSource = new ObservableCollection<HRU>(projectManager.CurrentProject.SelectedSubBasinHrus);
        }

        public void openNewProjectDialog(object sender, RoutedEventArgs e) {
            this.newProjectDialog = new NewProjectDialog(this.projectManager);
            this.newProjectDialog.Closing += this.setupProjectUI;
            this.newProjectDialog.ShowDialog();
        }

        public void openExistingProject(object sender, RoutedEventArgs e) {
            using (var fbd = new WinForms.FolderBrowserDialog()) {
                WinForms.DialogResult result = fbd.ShowDialog();

                if (result == WinForms.DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath)) {
                    projectManager.readProject(fbd.SelectedPath);
                    all_sub_basins_list.ItemsSource 
                        = new ObservableCollection<SubBasin>(projectManager.CurrentProject.SubBasins);
                    all_landuse_list.ItemsSource 
                        = new ObservableCollection<HRU>(projectManager.CurrentProject.SelectedSubBasinHrus);
                }
            }
        }

        public void setupProjectUI(object sender, System.ComponentModel.CancelEventArgs e) {
            Console.WriteLine("Setting up ui");
            var basinListSize = projectManager.CurrentProject.SubBasins.Capacity;
            if (basinListSize > 0) {
                all_sub_basins_list.ItemsSource
                    = new ObservableCollection<SubBasin>(projectManager.CurrentProject.SubBasins);
                all_landuse_list.ItemsSource
                    = new ObservableCollection<HRU>(projectManager.CurrentProject.SelectedSubBasinHrus);
            }
        }

        public void createApex(object sender, RoutedEventArgs e) {
            this.projectManager.createApexControlFiles();
            this.projectManager.createApexOperationsFiles();
            this.projectManager.createSubAreaFiles();
            this.projectManager.createSoilFiles();
            this.projectManager.createSiteFile();
            this.projectManager.createWeatherFiles(0);
            this.projectManager.createWmpFiles();
        }

        public void createApexControl(object sender, RoutedEventArgs e)
        {
            this.projectManager.createApexControlFiles();
        }

        public void createApexOperations(object sender, RoutedEventArgs e)
        {
            this.projectManager.createApexOperationsFiles();
        }

        public void createApexSubAreas(object sender, RoutedEventArgs e)
        {
            this.projectManager.createSubAreaFiles();
        }

        public void createApexSoils(object sender, RoutedEventArgs e)
        {
            this.projectManager.createSoilFiles();
        }


        public void createApexSite(object sender, RoutedEventArgs e)
        {
            this.projectManager.createSiteFile();
        }


        public void createApexWeather(object sender, RoutedEventArgs e)
        {
            this.projectManager.createWeatherFiles(0);
        }

        public void createApexWeatherStation(object sender, RoutedEventArgs e)
        {
            this.projectManager.createWmpFiles();
        }

        private void selectAllSubBasins(object sender, RoutedEventArgs e)
        {


        }

        private void deselectAllSubBasins(object sender, RoutedEventArgs e)
        {


        }

        private void selectAllHrus(object sender, RoutedEventArgs e)
        {


        }


        private void deselectAllHrus(object sender, RoutedEventArgs e)
        {


        }

        protected override void OnClosing(System.ComponentModel.CancelEventArgs e)
        {
            //do my stuff before closing
            base.OnClosing(e);
        }
    }
}
