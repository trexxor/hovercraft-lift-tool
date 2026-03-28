using System.Windows;
using System.Windows.Controls;
using HovercraftLiftTool.ViewModels;

namespace HovercraftLiftTool;

public partial class MainWindow : Window
{
    private readonly MainViewModel _vm;

    public MainWindow()
    {
        InitializeComponent();
        _vm = new MainViewModel();
        DataContext = _vm;
    }

    // Called by every input TextBox's LostFocus event.
    // Mirrors the VBA _Exit event pattern — recalculate when the user leaves a field.
    private void InputField_LostFocus(object sender, RoutedEventArgs e)
    {
        // Skip if the field is read-only (solved-for) — user didn't change it.
        if (sender is TextBox { IsReadOnly: true }) return;
        _vm.RecalculateCommand.Execute(null);
    }

    // Called by every input TextBox's KeyDown event.
    // Triggers recalculation on Enter or Tab, matching the VBA clsTextBox KeyDown handler.
    private void InputField_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
    {
        if (e.Key == System.Windows.Input.Key.Return ||
            e.Key == System.Windows.Input.Key.Tab)
        {
            _vm.RecalculateCommand.Execute(null);
        }
    }
}
