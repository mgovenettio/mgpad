using System.Collections.Generic;
using WpfApplication = System.Windows.Application;
using WpfStartupEventArgs = System.Windows.StartupEventArgs;
using WpfWindowStartupLocation = System.Windows.WindowStartupLocation;

namespace MGPad;

public partial class App : WpfApplication
{
    protected override void OnStartup(WpfStartupEventArgs e)
    {
        base.OnStartup(e);

        var recoveryService = new RecoveryService();
        List<RecoveryItem> recoverableItems = recoveryService.GetRecoverableItems();
        RecoveryItem? selectedRecovery = null;

        if (recoverableItems.Count > 0)
        {
            var recoveryWindow = new RecoveryWindow(recoverableItems, recoveryService)
            {
                WindowStartupLocation = WpfWindowStartupLocation.CenterScreen
            };

            if (recoveryWindow.ShowDialog() == true)
            {
                selectedRecovery = recoveryWindow.SelectedRecovery;
            }
        }

        var mainWindow = new MainWindow(selectedRecovery);
        MainWindow = mainWindow;
        mainWindow.Show();
    }
}
