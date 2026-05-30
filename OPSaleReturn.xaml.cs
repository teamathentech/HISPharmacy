using HISPharmacy.GeneralModules;
using HISPharmacy.Models;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using static HISPharmacy.Pharmacy.IPSales;

namespace HISPharmacy.Pharmacy
{
    public partial class OPSaleReturn : UserControl
    {
        string strQry;
        DataHelper obj = new DataHelper();
        public GlobalVariables gvars;
        DataSet ds, dsBatch;
        decimal taxAmt, taxPer, Rate;
        string SupId, PBillNo;
        bool FLoad = false;
        string OrganisationID = string.Empty;
        ObservableCollection<AddItems> aItem { get; set; } = new ObservableCollection<AddItems>();
        ObservableCollection<AddPayDet> payDet { get; set; } = new ObservableCollection<AddPayDet>();
        public AutoCompleteCombobox ItemCB { get; set; }
        private List<Item> allItems; // Full item list
        private ICollectionView comboView;
        bool isHigh, isLook, isSound;

        private void btnFind_Click(object sender, RoutedEventArgs e)
        {

        }

        private void btnView_Click(object sender, RoutedEventArgs e)
        {

        }

        private void btnPrint_Click(object sender, RoutedEventArgs e)
        {

        }

        public OPSaleReturn()
        {
            InitializeComponent();
        }

        #region Page Loaded
        private void Page_Loaded(object sender, RoutedEventArgs e)
        {

        }
        #endregion

        #region GO
        private void btnGo_Click(object sender, RoutedEventArgs e)
        {
            SqlParameter[] sqlParamSearch = new SqlParameter[]
           {
                new SqlParameter("@ACTIVITY", "SalesData"),
                new SqlParameter("@FromDate", Convert.ToDateTime(Convert.ToDateTime(dtpFrom.Text).ToString("dd MMM yyyy"))),
                new SqlParameter("@ToDate", Convert.ToDateTime(Convert.ToDateTime(dtpTo.Text).ToString("dd MMM yyyy"))),
                new SqlParameter("@Searchtxt", txtFindBillNo.Text),
                new SqlParameter("@LocationID",  gvars.gLocationId),
                new SqlParameter("@DepartmentID", gvars.gDeptID),
           };


            DataSet dsSearch = obj.getDataset("tSales", DataHelper.SqlCmdType.sqlStoredProcedure, sqlParamSearch);
            if (dsSearch.Tables[0].Rows.Count > 0)
            {
                if (!dsSearch.Tables[0].Columns.Contains("Slno"))
                {
                    dsSearch.Tables[0].Columns.Add("Slno", typeof(int));
                }
                for (int i = 0; i < dsSearch.Tables[0].Rows.Count; i++)
                {
                    dsSearch.Tables[0].Rows[i]["Slno"] = i + 1;
                }
                dgvSalesRetFind.ItemsSource = dsSearch.Tables[0].DefaultView;
            }
            else
            {
                MessageBox.Show("No data found.", "Information", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }
        #endregion

        #region New
        private void btnNew_Click(object sender, RoutedEventArgs e)
        {
            if (btnNew.Content.ToString() == "+ Add New")
            {
                GridPanel.Visibility = Visibility.Collapsed;
                scrollFormPanel.Visibility = Visibility.Visible;
                FormPanel.Visibility = Visibility.Visible;
                btnNew.Content = "Back";
                btnNew.Background = (SolidColorBrush)(new BrushConverter().ConvertFrom("#dc3545"));
                //EnableControls();
                //ClearControls();
                txtFindBillNo.Focus();
            }
            else
            {
                GridPanel.Visibility = Visibility.Visible;
                FormPanel.Visibility = Visibility.Collapsed;
                btnNew.Background = (SolidColorBrush)(new BrushConverter().ConvertFrom("#17a2b8"));
                btnNew.Content = "+ Add New";
            }
        }

        #endregion
    }
}
