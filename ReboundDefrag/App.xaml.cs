using Microsoft.UI.Xaml;
using System;
using System.Linq;
using System.Threading.Tasks;

namespace ReboundDefrag
{
    public partial class App : Application
    {
        public App()
        {
            this.InitializeComponent();
        }

        protected override async void OnLaunched(LaunchActivatedEventArgs args)
        {
            m_window = new MainWindow();
            m_window.Activate();
            await (m_window as MainWindow).LoadAppAsync();

            string commandArgs = string.Join(" ", Environment.GetCommandLineArgs().Skip(1));

            if (commandArgs.Contains("SELECTED-SYSTEM"))
            {
                try
                {
                    // Extract the index after "SELECTED "
                    (m_window as MainWindow).AdvancedView.IsOn = true;
                    await (m_window as MainWindow).LoadData(true);
                    int selectedIndex = int.Parse(commandArgs[(commandArgs.IndexOf("SELECTED-SYSTEM") + 16)..].Trim());
                    (m_window as MainWindow).MyListView.SelectedIndex = selectedIndex;
                    await Task.Delay(2000);
                    (m_window as MainWindow).OptimizeSelected(true);
                }
                catch (Exception ex)
                {
                    await (m_window as MainWindow).ShowMessageDialogAsync(ex.Message);
                }
            }
            else if (commandArgs.Contains("SELECTED"))
            {
                try
                {
                    // Extract the index after "SELECTED "
                    await (m_window as MainWindow).LoadData(false);
                    int selectedIndex = int.Parse(commandArgs[(commandArgs.IndexOf("SELECTED") + 9)..].Trim());
                    (m_window as MainWindow).MyListView.SelectedIndex = selectedIndex;
                    await Task.Delay(1000);
                    (m_window as MainWindow).OptimizeSelected(false);
                }
                catch (Exception ex)
                {
                    await (m_window as MainWindow).ShowMessageDialogAsync(ex.Message);
                }
            }
            else if (commandArgs == "OPTIMIZEALL-SYSTEM")
            {
                try
                {
                    await (m_window as MainWindow).LoadData(true);
                    await Task.Delay(1000);
                    (m_window as MainWindow).OptimizeAll(false, true);
                }
                catch (Exception ex)
                {
                    await (m_window as MainWindow).ShowMessageDialogAsync(ex.Message);
                }
            }
            else if (commandArgs == "OPTIMIZEALL")
            {
                try
                {
                    await (m_window as MainWindow).LoadData(false);
                    await Task.Delay(1000);
                    (m_window as MainWindow).OptimizeAll(false, false);
                }
                catch (Exception ex)
                {
                    await (m_window as MainWindow).ShowMessageDialogAsync(ex.Message);
                }
            }
            else if (commandArgs == "OPTIMIZEALLANDCLOSE-SYSTEM")
            {
                try
                {
                    await (m_window as MainWindow).LoadData(true);
                    await Task.Delay(1000);
                    (m_window as MainWindow).OptimizeAll(true, true);
                }
                catch (Exception ex)
                {
                    await (m_window as MainWindow).ShowMessageDialogAsync(ex.Message);
                }
            }
            else if (commandArgs == "OPTIMIZEALLANDCLOSE")
            {
                try
                {
                    await (m_window as MainWindow).LoadData(false);
                    await Task.Delay(1000);
                    (m_window as MainWindow).OptimizeAll(true, false);
                }
                catch (Exception ex)
                {
                    await (m_window as MainWindow).ShowMessageDialogAsync(ex.Message);
                }
            }
        }

        private Window m_window;
    }
}
