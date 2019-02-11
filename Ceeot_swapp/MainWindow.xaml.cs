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

        public MainWindow()
        {
            InitializeComponent();
            projectManager = new ProjectManager();
        }

        public void openNewProjectDialog(object sender, RoutedEventArgs e)
        {
            this.newProjectDialog = new NewProjectDialog();
            this.newProjectDialog.projectManager = this.projectManager;
            this.newProjectDialog.ShowDialog();
            this.newProjectDialog.Closing += new System.ComponentModel.CancelEventHandler(this.setupProjectUI);
        }

        public void setupProjectUI(object sender, System.ComponentModel.CancelEventArgs e)
        {
            ;
        }

        protected override void OnClosing(System.ComponentModel.CancelEventArgs e)
        {
            //do my stuff before closing
            base.OnClosing(e);
        }
    }
}
