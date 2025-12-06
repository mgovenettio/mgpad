using System.Collections.Generic;
using System.Windows;

namespace MGPad;

public partial class App : Application
{
    protected override void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);

        var recoveryService = new RecoveryService();
        List<RecoveryItem> recoverableItems = recoveryService.GetRecoverableItems();
        RecoveryItem? selectedRecovery = null;

        if (recoverableItems.Count > 0)
        {
            var recoveryWindow = new RecoveryWindow(recoverableItems, recoveryService)
            {
                WindowStartupLocation = WindowStartupLocation.CenterScreen
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
