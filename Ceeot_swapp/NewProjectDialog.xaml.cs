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
using System.Windows.Shapes;

namespace Ceeot_swapp
{
    /// <summary>
    /// Interaction logic for NewProjectDialog.xaml
    /// </summary>
    public partial class NewProjectDialog : Window
    {
        public ProjectManager projectManager;
        public NewProjectDialog()
        {
            InitializeComponent();
            // load 
        }

        private void okBtn_Click(object sender, RoutedEventArgs e)
        {
            if (this.createNewProject())
            {
                this.Close();
            }
        }

        private void cancelBtn_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = false;
            this.Close();
        }

        private void openFolderSelectionDialog(object sender, RoutedEventArgs e)
        {
            using (var fbd = new System.Windows.Forms.FolderBrowserDialog())
            {
                System.Windows.Forms.DialogResult result = fbd.ShowDialog();
                if (result == System.Windows.Forms.DialogResult.OK 
                    && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
                {
                    this.proj_loc_txt.Text = fbd.SelectedPath; 
                }
            }
        }

        public bool createNewProject()
        {
            ProjectManager.Version apexVersion = 0, swattVersion = 0;

            // get project name and location
            string name = proj_name_txt.Text;
            if (name == "")
            {
                MessageBox.Show("Project name cannot be empty!", "Project Creation Error");
                return false;
            }
            string location = proj_name_txt.Text;
            if (location == "")
            {
                MessageBox.Show("Project location cannot be empty!", "Project Creation Error");
                return false;
            }

            // select apex version
            if (apex_version_0406.IsChecked == true) apexVersion = ProjectManager.Version.APEX_0604;
            else if (apex_version_0406.IsChecked == true) apexVersion = ProjectManager.Version.APEX_0806;
            // select swatt version
            if (swatt_version_2005.IsChecked == true) swattVersion = ProjectManager.Version.SWATT_2005;
            else if (swatt_version_2009.IsChecked == true) swattVersion = ProjectManager.Version.SWATT_2009;
            else if (swatt_version_2012.IsChecked == true) swattVersion = ProjectManager.Version.SWATT_2012;

            Console.WriteLine(name + " " + location);
            // create project with project manager
            this.projectManager.createProject(name, location, apexVersion, swattVersion);

            return true;
        }
    }
 }
