﻿namespace Morphic.ManualTester
{
    using System.Windows;
    using System.Windows.Controls;
    using System.Windows.Input;
    using System.Windows.Media;
    using Settings.SettingsHandlers;

    /// <summary>
    /// Interaction logic for ManualControlInteger.xaml
    /// </summary>
    public partial class ManualControlInteger : UserControl
    {
        public bool changed;
        private readonly Brush greenfield = new SolidColorBrush(Color.FromArgb(30, 0, 176, 0));
        private readonly Brush redfield = new SolidColorBrush(Color.FromArgb(30, 255, 0, 0));
        public Setting setting;
        private readonly Brush whitefield = new SolidColorBrush(Color.FromArgb(255, 255, 255, 255));
        private readonly MainWindow window;

        public ManualControlInteger(MainWindow window, Setting setting)
        {
            this.InitializeComponent();
            this.window = window;
            this.setting = setting;
            this.ControlName.Text = setting.Name;
            this.CaptureSetting();
        }

        private bool Validate()
        {
            try
            {
                long value = long.Parse(this.InputField.Text);
                return true;
            }
            catch
            {
                this.InputField.Background = this.redfield;
                return false;
            }
        }

        public async void CaptureSetting()
        {
            this.LoadingIcon.Visibility = Visibility.Visible;
            this.InputField.Text = "";
            this.InputField.Text = (await this.setting.GetValue<int>()).ToString();
            this.InputField.Background = this.whitefield;
            this.LoadingIcon.Visibility = Visibility.Hidden;
        }

        private void EnterCheck(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                this.changed = true;
                this.InputField.Background = this.greenfield;
                if (this.Validate() && this.window.AutoApply)
                {
                    this.ApplySetting();
                }
            }
        }

        private void ValueChanged(object sender, RoutedEventArgs e)
        {
            this.changed = true;
            this.InputField.Background = this.greenfield;
            if (this.Validate() && this.window.AutoApply)
            {
                this.ApplySetting();
            }
        }

        public async void ApplySetting()
        {
            if (!this.changed)
            {
                return;
            }

            this.changed = false;
            try
            {
                int value = int.Parse(this.InputField.Text);
                this.InputField.Background = this.whitefield;
                var result = await this.setting.SetValueAsync(value);
                if (result.IsError)
                {
                    this.CaptureSetting();
                }
            }
            catch
            {
                this.CaptureSetting();
            }
        }
    }
}
