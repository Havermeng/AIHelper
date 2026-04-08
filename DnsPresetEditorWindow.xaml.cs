using System.Net;
using System.Windows;
using LaptopSessionViewer.Models;
using LaptopSessionViewer.Services;

namespace LaptopSessionViewer;

public partial class DnsPresetEditorWindow : Window
{
    private readonly LocalizationService _strings;

    public DnsPresetEditorWindow(LocalizationService strings, DnsPreset? preset = null)
    {
        InitializeComponent();
        _strings = strings;

        Title = preset is null ? strings["DnsPresetDialogAddTitle"] : strings["DnsPresetDialogEditTitle"];
        HeaderTextBlock.Text = Title;
        NameLabel.Text = strings["DnsPresetNameLabel"];
        PrimaryLabel.Text = strings["DnsPrimaryLabel"];
        SecondaryLabel.Text = strings["DnsSecondaryLabel"];
        DescriptionLabel.Text = strings["DnsPresetDescriptionLabel"];
        DohCheckBox.Content = strings["DnsUseDohLabel"];
        DohTemplateLabel.Text = strings["DnsDohTemplateLabel"];
        SaveButton.Content = strings["Save"];
        CancelButton.Content = strings["Cancel"];

        if (preset is not null)
        {
            NameTextBox.Text = preset.Name;
            PrimaryTextBox.Text = preset.PrimaryDns;
            SecondaryTextBox.Text = preset.SecondaryDns;
            DescriptionTextBox.Text = preset.Description;
            DohCheckBox.IsChecked = preset.EnableDoh;
            DohTemplateTextBox.Text = preset.DohTemplate;
        }

        UpdateDohTemplateVisibility();
    }

    public DnsPreset? ResultPreset { get; private set; }

    private void SaveButton_Click(object sender, RoutedEventArgs e)
    {
        var name = NameTextBox.Text.Trim();
        var primaryDns = PrimaryTextBox.Text.Trim();
        var secondaryDns = SecondaryTextBox.Text.Trim();
        var description = DescriptionTextBox.Text.Trim();
        var enableDoh = DohCheckBox.IsChecked == true;
        var dohTemplate = DohTemplateTextBox.Text.Trim();

        if (string.IsNullOrWhiteSpace(name))
        {
            MessageBox.Show(
                _strings["DnsPresetNameRequired"],
                Title,
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
            return;
        }

        if (!IPAddress.TryParse(primaryDns, out _))
        {
            MessageBox.Show(
                _strings["DnsPresetPrimaryInvalid"],
                Title,
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
            return;
        }

        if (!string.IsNullOrWhiteSpace(secondaryDns) && !IPAddress.TryParse(secondaryDns, out _))
        {
            MessageBox.Show(
                _strings["DnsPresetSecondaryInvalid"],
                Title,
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
            return;
        }

        if (enableDoh &&
            (!Uri.TryCreate(dohTemplate, UriKind.Absolute, out var uri) ||
             !string.Equals(uri.Scheme, Uri.UriSchemeHttps, StringComparison.OrdinalIgnoreCase)))
        {
            MessageBox.Show(
                _strings["DnsDohTemplateInvalid"],
                Title,
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
            return;
        }

        ResultPreset = new DnsPreset
        {
            Name = name,
            PrimaryDns = primaryDns,
            SecondaryDns = secondaryDns,
            Description = description,
            EnableDoh = enableDoh,
            DohTemplate = enableDoh ? dohTemplate : string.Empty,
            IsCustom = true,
            IsAutomaticPreset = false
        };

        DialogResult = true;
        Close();
    }

    private void CancelButton_Click(object sender, RoutedEventArgs e)
    {
        DialogResult = false;
        Close();
    }

    private void DohCheckBox_Changed(object sender, RoutedEventArgs e)
    {
        UpdateDohTemplateVisibility();
    }

    private void UpdateDohTemplateVisibility()
    {
        DohTemplatePanel.Visibility = DohCheckBox.IsChecked == true ? Visibility.Visible : Visibility.Collapsed;
    }
}
