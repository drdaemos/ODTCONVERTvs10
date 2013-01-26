using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Threading;

namespace ODTCONVERT
{
    /// <summary>
    /// Логика взаимодействия для App.xaml
    /// </summary>
    public sealed partial class App : System.Windows.Application
    {
        public void StartupHandler(object sender, System.Windows.StartupEventArgs e)    {
            Elysium.Manager.Apply(this, Elysium.Theme.Dark, Elysium.AccentBrushes.Blue, Elysium.AccentBrushes.Sky);
        }
    }
}
