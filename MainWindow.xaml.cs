using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Data; 
using System.Data.SqlClient;
using System.Security.Cryptography;
using HISPharmacy.GeneralModules;

namespace HISPharmacy
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        string strQry;
        DataHelper obj = new DataHelper();
       public GlobalVariables gVars;
        DataSet ds;
        private MainWindowViewModel _viewModel;
        public MainWindow()
        {
            _viewModel = new MainWindowViewModel();
            DataContext = _viewModel;
            InitializeComponent();
            
        }
         

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            tabUser.Content = gVars.gUserID + "  Login Time: " + DateTime.Now.ToString("dd-MMM-yyyy hh:mm:ss tt");
            //SqlParameter[] sqlParamSearch = new SqlParameter[]
            //{ 
            //    new SqlParameter("@ACTIVITY", "GetMenus"),
            //    new SqlParameter("@UserID", gVars.gUserID),
            //    new SqlParameter("@Name","Medical Stores Management"),
            //    new SqlParameter("@ModuleID", "MC0005"),
            //    new SqlParameter("@RoleID", gVars.gRoleId)
            //};
            //DataSet dsMenu = obj.getDataset("mModules", DataHelper.SqlCmdType.sqlStoredProcedure, sqlParamSearch);

            SqlParameter[] sqlParamSubMenu = new SqlParameter[]
            {
                new SqlParameter("@ACTIVITY", "GetSubMenusConsole"),
                new SqlParameter("@ModuleID",  "MC0005"),
                new SqlParameter("@Name","Medical Stores Management"),
                new SqlParameter("@UserID", gVars.gUserID),
                new SqlParameter("@RoleID", gVars.gRoleId)
            };
            DataSet dsSubMenu = obj.getDataset("mModules", DataHelper.SqlCmdType.sqlStoredProcedure, sqlParamSubMenu);
            //if (dsMenu.Tables[0].Rows.Count > 0)
            //{
            //    foreach (DataRow item in dsMenu.Tables[0].Rows)
            //    {
            //        MenuItem menuItem = new MenuItem { Header = item[0] };
                    if (dsSubMenu.Tables[0].Rows.Count > 0)
                    {
                      //  DataRow dtSubMenu = dsSubMenu.Tables[0].Select(" ='" + item[1]  + "' ");
                        foreach (DataRow Subitem in dsSubMenu.Tables[0].Select(" RootFormID ='127'"))
                        {
                            MenuItem subMenuItem = new MenuItem
                            {
                                Header = Subitem[0]
                            };
                             subMenuItem.Click += SubMenuItem_Click;

                    MainMenu.Items.Add(subMenuItem);
                        }
                    }
                    //MainMenu.Items.Add(menuItem);
               // }
                MenuItem menuItem1 = new MenuItem { Header = "Exit" };
                menuItem1.Click += (s, args) => {
                    var result = MessageBox.Show("Are you sure you want to exit?", "Exit", MessageBoxButton.YesNo, MessageBoxImage.Question);
                    if (result == MessageBoxResult.Yes)
                        Application.Current.Shutdown();
                };
                MainMenu.Items.Add(menuItem1);
            //}

        } 
        private void SubMenuItem_Click(object sender, RoutedEventArgs e)
        {
            var clicked = (MenuItem)sender;
            string menuHeader = clicked.Header.ToString();

            switch (menuHeader)
            {
                case "Sales":
                    var salesForm = new Pharmacy.Sales
                    {
                        gvars = gVars // Assign global variables if needed
                    };
                    OpenTab("Sales", salesForm);
                    break;
                case "IP Sales":
                    var ipsalesForm = new Pharmacy.IPSales
                    {
                        gvars = gVars // Assign global variables if needed
                    };
                    OpenTab("IPSales", ipsalesForm);
                    break;
                case "IP Sales Return":
                    var ipsalesReturnForm = new Pharmacy.IPSalesReturn
                    {
                        gvars = gVars // Assign global variables if needed
                    };
                    OpenTab("IPSalesReturn", ipsalesReturnForm);
                    break;
                   // _NavigationFrame.Navigate(salesForm);

                case "Sales Return":
                    var salesRetForm = new Pharmacy.OPSaleReturn
                    {
                        gvars = gVars // Assign global variables if needed
                    };
                    OpenTab("SaleReturn", salesRetForm);
                    break;
                    
                case "Exit":
                    Application.Current.Shutdown();
                    break;
                default:                     
                    break;
            }
        }
        private void OpenTab(string header, UserControl control)
        {
            if (string.IsNullOrWhiteSpace(header) || control == null)
                return;

            // Check if tab already exists
            foreach (TabItem tab in MainTab.Items)
            {
                if (tab.Header?.ToString() == header)
                {
                    MainTab.SelectedItem = tab;
                    return;
                }
            }
            var closeButton = new Button
            {
                Content = "×",
                Width = 16,
                Height = 16,
                Padding = new Thickness(0),
                Margin = new Thickness(5, 0, 0, 0),
                Background = null,
                BorderBrush = null,
                Cursor = System.Windows.Input.Cursors.Hand
            };

            // Header with title + close button
            var headerPanel = new StackPanel { Orientation = Orientation.Horizontal };
            headerPanel.Children.Add(new TextBlock { Text = header,FontWeight = FontWeights.Bold,FontSize = 16 });
            headerPanel.Children.Add(closeButton);

            // Create new tab
            var tabItem = new TabItem
            {
                Header = headerPanel,
                Content = control,
                // Optional: add close button support or style here
            };

            closeButton.Click += (s, args) => MainTab.Items.Remove(tabItem);

            MainTab.Items.Add(tabItem);
            MainTab.SelectedItem = tabItem;
        }
    }
}
