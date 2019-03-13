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

        public SwattProject.SubBasin SelectedSubBasin { get; set; }

        private void getSelectedItem(object sender, MouseButtonEventArgs e)
        {
            string basinFileName = (string)all_sub_basins_list.SelectedItems[0];
            System.Windows.MessageBox.Show(basinFileName);

            projectManager.CurrentProject.SubBasins.ForEach(delegate(SubBasin b) {
                if (b.Name == basinFileName) {
                    b.Selected = true;
                }
            });
        }

        public void openNewProjectDialog(object sender, RoutedEventArgs e)
        {
            this.newProjectDialog = new NewProjectDialog();
            this.newProjectDialog.Closing += this.setupProjectUI;
            this.newProjectDialog.projectManager = this.projectManager;
            this.newProjectDialog.ShowDialog();
        }

        public void setupProjectUI(object sender, System.ComponentModel.CancelEventArgs e)
        {
            Console.WriteLine("Setting up ui");
            this.projectManager.loadSubBasins();
            this.DataContext = this.projectManager.CurrentProject;
        }

        protected override void OnClosing(System.ComponentModel.CancelEventArgs e)
        {
            //do my stuff before closing
            base.OnClosing(e);
        }
    }
}
