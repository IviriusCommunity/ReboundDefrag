using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Controls.Primitives;
using Microsoft.UI.Xaml.Data;
using Microsoft.UI.Xaml.Input;
using Microsoft.UI.Xaml.Media;
using Microsoft.UI.Xaml.Navigation;
using Microsoft.UI.Xaml.Shapes;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices.WindowsRuntime;
using System.Security.Principal;
using System.Threading.Tasks;
using Windows.ApplicationModel;
using Windows.ApplicationModel.Activation;
using Windows.Foundation;
using Windows.Foundation.Collections;
using Windows.UI.Core;

// To learn more about WinUI, the WinUI project structure,
// and more about our project templates, see: http://aka.ms/winui-project-info.

namespace ReboundDefrag
{
    /// <summary>
    /// Provides application-specific behavior to supplement the default Application class.
    /// </summary>
    public partial class App : Application
    {
        /// <summary>
        /// Initializes the singleton application object.  This is the first line of authored code
        /// executed, and as such is the logical equivalent of main() or WinMain().
        /// </summary>
        public App()
        {
            this.InitializeComponent();
        }

        /// <summary>
        /// Invoked when the application is launched.
        /// </summary>
        /// <param name="args">Details about the launch request and process.</param>
        protected override async void OnLaunched(Microsoft.UI.Xaml.LaunchActivatedEventArgs args)
        {
            m_window = new MainWindow();
            IntPtr hWnd = WinRT.Interop.WindowNative.GetWindowHandle(m_window);
            var windowId = Microsoft.UI.Win32Interop.GetWindowIdFromWindow(hWnd);
            var appWindow = Microsoft.UI.Windowing.AppWindow.GetFromWindowId(windowId);
            m_window.Activate();
            await (m_window as MainWindow).LoadAppAsync();

            string commandArgs = string.Join(" ", Environment.GetCommandLineArgs().Skip(1));

            if (commandArgs.Contains("SELECTED-SYSTEM"))
            {
                try
                {
                    // Extract the index after "SELECTED "
                    int selectedIndex = Int32.Parse(commandArgs.Substring(commandArgs.IndexOf("SELECTED-SYSTEM") + 16).Trim());
                    await (m_window as MainWindow).LoadData(true);
                    (m_window as MainWindow).MyListView.SelectedIndex = selectedIndex;
                    await Task.Delay(50);
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
                    int selectedIndex = Int32.Parse(commandArgs.Substring(commandArgs.IndexOf("SELECTED") + 9).Trim());
                    (m_window as MainWindow).MyListView.SelectedIndex = selectedIndex;
                    await Task.Delay(50);
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
                    (m_window as MainWindow).OptimizeAll(true, false);
                }
                catch (Exception ex)
                {
                    await (m_window as MainWindow).ShowMessageDialogAsync(ex.Message);
                }
            }
        }

        private Window m_window;

        public static bool IsAdmin()
        {
            var identity = WindowsIdentity.GetCurrent();
            var principal = new WindowsPrincipal(identity);
            return principal.IsInRole(WindowsBuiltInRole.Administrator);
        }
    }
}
