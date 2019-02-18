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
        //
        TabContent tabContent;
        ObservableCollection<TabContent> Tabs { get; set; }

        public class TabContent
        {
            public List<ProjectManager.Project.SubBasin> subBasins;
        }
        public MainWindow()
        {
            InitializeComponent();

            projectManager = new ProjectManager();
            this.tab_control.ItemsSource = projectManager.Projects;
            this.tab_control.SelectionChanged += this.updateCurrentProject;

            Tabs = new ObservableCollection<TabContent>();
            Tabs.Add(new TabContent {
                subBasins = {
                    //new ProjectManager.Project.SubBasin(),
                    //new ProjectManager.Project.SubBasin()
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
            this.tab_control.ItemsSource = projectManager.Projects;
        }

        public void updateCurrentProject(object sender, SelectionChangedEventArgs e)
        {

        }

        protected override void OnClosing(System.ComponentModel.CancelEventArgs e)
        {
            //do my stuff before closing
            base.OnClosing(e);
        }
    }
}
