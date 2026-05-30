using System;
using System.Collections.Generic;
using System.Linq;
using System.ComponentModel;
using System.Globalization;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using MahApps.Metro;
using System.Windows.Input;
using MahApps.Metro.Controls;
using MahApps.Metro.Controls.Dialogs;

namespace HISPharmacy
{
   

        public class AccentColorMenuData
        {
            public string Name { get; set; }
            public Brush BorderColorBrush { get; set; }
            public Brush ColorBrush { get; set; }

            private ICommand changeAccentCommand;

            public ICommand ChangeAccentCommand
            {
                get { return this.changeAccentCommand ?? (changeAccentCommand = new SimpleCommand { CanExecuteDelegate = x => true, ExecuteDelegate = x => this.DoChangeTheme(x) }); }
            }

            protected virtual void DoChangeTheme(object sender)
            {
                var theme = ThemeManager.DetectAppStyle(Application.Current);
                var accent = ThemeManager.GetAccent(this.Name);
                ThemeManager.ChangeAppStyle(Application.Current, accent, theme.Item1);
            }
        }

        public class AppThemeMenuData : AccentColorMenuData
        {
            protected override void DoChangeTheme(object sender)
            {
                var theme = ThemeManager.DetectAppStyle(Application.Current);
                var appTheme = ThemeManager.GetAppTheme(this.Name);
                ThemeManager.ChangeAppStyle(Application.Current, theme.Item2, appTheme);
            }
        }

        public class MainWindowViewModel 
        {
            private readonly IDialogCoordinator _dialogCoordinator;
            int? _integerGreater10Property;
            private bool _animateOnPositionChange = true;

            public MainWindowViewModel()
            {


                // create accent color menu items for the demo
                this.AccentColors = ThemeManager.Accents
                                                .Select(a => new AccentColorMenuData() { Name = a.Name, ColorBrush = a.Resources["AccentColorBrush"] as Brush })
                                                .ToList();

                // create metro theme color menu items for the demo
                this.AppThemes = ThemeManager.AppThemes
                                               .Select(a => new AppThemeMenuData() { Name = a.Name, BorderColorBrush = a.Resources["BlackColorBrush"] as Brush, ColorBrush = a.Resources["WhiteColorBrush"] as Brush })
                                               .ToList();


            }
            public List<AccentColorMenuData> AccentColors { get; set; }
            public List<AppThemeMenuData> AppThemes { get; set; }
        }
    
}
