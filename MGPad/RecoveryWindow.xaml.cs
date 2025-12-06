using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;

namespace MGPad;

public partial class RecoveryWindow : Window
{
    private readonly RecoveryService _recoveryService;

    public ObservableCollection<RecoveryItem> RecoverableItems { get; }

    public RecoveryItem? SelectedRecovery { get; private set; }

    public RecoveryWindow(IEnumerable<RecoveryItem> items, RecoveryService recoveryService)
    {
        InitializeComponent();
        DataContext = this;
        _recoveryService = recoveryService;
        RecoverableItems = new ObservableCollection<RecoveryItem>(items);
        UpdateButtonStates();
    }

    private void RecoverButton_Click(object sender, RoutedEventArgs e)
    {
        if (RecoveryGrid.SelectedItem is RecoveryItem item)
        {
            SelectedRecovery = item;
            DialogResult = true;
            Close();
        }
    }

    private void DiscardSelectedButton_Click(object sender, RoutedEventArgs e)
    {
        if (RecoveryGrid.SelectedItem is not RecoveryItem item)
        {
            return;
        }

        _recoveryService.Discard(item);
        RecoverableItems.Remove(item);
        UpdateButtonStates();
    }

    private void DiscardAllButton_Click(object sender, RoutedEventArgs e)
    {
        _recoveryService.DiscardAll(RecoverableItems.ToList());
        RecoverableItems.Clear();
        SelectedRecovery = null;
        DialogResult = false;
        Close();
    }

    private void CloseButton_Click(object sender, RoutedEventArgs e)
    {
        DialogResult = false;
        Close();
    }

    private void UpdateButtonStates()
    {
        bool hasItems = RecoverableItems.Count > 0;
        RecoverButton.IsEnabled = hasItems;
        DiscardSelectedButton.IsEnabled = hasItems;
    }
}
