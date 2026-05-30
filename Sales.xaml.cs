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
using System.Collections.ObjectModel;
using System.Windows.Input;
using HISPharmacy.Models;
using HISPharmacy.Commands;
using System.Diagnostics;
using System.ComponentModel;
using System.Windows.Data;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System.IO;
using System.Windows.Media;
using System.Threading;
using System.Windows.Threading;

namespace HISPharmacy.Pharmacy
{
    /// <summary>
    /// Interaction logic for Sales.xaml
    /// </summary>
    public partial class Sales : UserControl
    {
        string strQry;
        DataHelper obj = new DataHelper();
        CoreHelper CH = new CoreHelper();
        public GlobalVariables gvars;
        DataSet ds,dsBatch;
        decimal taxAmt, taxPer,Rate;
        string SupId, PBillNo;
        bool FLoad = false;
        string OrganisationID = string.Empty;
        ObservableCollection<AddItems> aItem { get; set; } = new ObservableCollection<AddItems>();
        public ObservableCollection<Item> AllItems { get; set; } = new ObservableCollection<Item>();
        public ObservableCollection<Item> FilteredItems { get; set; } = new ObservableCollection<Item>();
        ObservableCollection<AddPayDet> payDet { get; set; } = new ObservableCollection<AddPayDet>();
        public AutoCompleteCombobox ItemCB { get; set; }
        private List<Item> allItems; // Full item list
        private ICollectionView comboView;
        bool isHigh, isLook, isSound;
        private bool isItemSelected = false;

        public Sales()
        {
            InitializeComponent();
            this.DataContext = this;
        }

        #region Page_Loaded
        private void Page_Loaded(object sender, RoutedEventArgs e)
        {
            dtpUHIDFrom.Text = dtpUHIDTo.Text = dtpBillDt.Text = dtpTo.Text = dtpFrom.Text = DateTime.Today.ToString();
            Helper.AttachDecimalInputHandlers(txtQty);
            Helper.AttachDecimalInputHandlers(txtDiscount);
            Helper.AttachDecimalInputHandlers(txtConcPer);
            Helper.AttachDecimalInputHandlers(txtPayAmount);
            Helper.AttachOnlyNumberInputHandlers(txtPhone);
            //cmbSaleType.SelectedIndex = 0;
            cmbDiscType.SelectedIndex = 0;         
            BindWallet();
            BindPayMode();
            BindDoctor();
            BindDueAuthorisation();
            GetUserDiscAuthorized();
            GetBank();
            BindItem();
           // allItems = LoadItemsFromDatabase(); 
            //cmbItemName.ItemsSource = allItems;
            //cmbItemName.DisplayMemberPath = "ItemName";
            //cmbItemName.SelectedValuePath = "ItemID";
            dgvItemDetails.ItemsSource = aItem;
            dgvPaymentDet.ItemsSource = payDet;
            rdoOP.IsChecked = true;
            this.PreviewKeyDown += Sales_PreviewKeyDown;
            txtUHIDPhone.Focus();
            //btnGo_Click(sender, e);
        }

        #endregion

        #region Sales_PreviewKeyDown
        private void Sales_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            UIElement element = Keyboard.FocusedElement as UIElement;
            if (element == null) return;
            if (IsInsideControl<Button>(element) || IsInsideControl<ListBox>(element) || IsInsideControl<ListView>(element) || IsInsideControl<DataGrid>(element))
            {
                return;
            }

            if (e.Key == Key.Enter)
            {
                TraversalRequest request = new TraversalRequest(FocusNavigationDirection.Next);

                element.MoveFocus(new TraversalRequest(FocusNavigationDirection.Next));

                e.Handled = true;
            }

            if (e.Key == Key.Tab && (Keyboard.Modifiers & ModifierKeys.Shift) == ModifierKeys.Shift)
            {
                TraversalRequest request = new TraversalRequest(FocusNavigationDirection.Previous); 

                if (element != null)
                {
                    element.MoveFocus(request);
                    e.Handled = true;
                }
            }
        }

        private bool IsInsideControl<T>(DependencyObject obj) where T : DependencyObject
        {
            while (obj != null)
            {
                if (obj is T)
                    return true;

                obj = VisualTreeHelper.GetParent(obj);
            }

            return false;
        }
        #endregion

        #region New
        private void btnNew_Click(object sender, RoutedEventArgs e)
        {
            if (btnNew.Content.ToString() == "+ Add New")
            {
                TitlePanel.Visibility = Visibility.Visible;
                GridPanel.Visibility = Visibility.Collapsed; 
                FormPanel.Visibility = Visibility.Visible;
                btnNew.Content = "Back";
                btnNew.Background = (SolidColorBrush)(new BrushConverter().ConvertFrom("#dc3545"));
                EnableControls();
                ClearControls();
                BindItem();
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
            if(dsSearch.Tables[0].Rows.Count > 0)
            {
                if (!dsSearch.Tables[0].Columns.Contains("Slno"))
                {
                    dsSearch.Tables[0].Columns.Add("Slno", typeof(int));
                }
                for (int i = 0; i < dsSearch.Tables[0].Rows.Count; i++)
                {
                    dsSearch.Tables[0].Rows[i]["Slno"] = i + 1;
                }
                dgvSalesFind.ItemsSource = dsSearch.Tables[0].DefaultView;
            }
            else
            {
                MessageBox.Show("No data found.", "Information", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }
        #endregion

        #region Print
        private void btnPrint_Click(object sender, RoutedEventArgs e)
        {
            var button = sender as Button;
            DataRowView pData = (DataRowView)button?.DataContext;
            if (pData != null)
            {
                PrintSlip(pData["BillNo"].ToString());
            }
      }
 
        private void PrintSlip( string strBillNo )
        {  
            // string spath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "PDFFiles\\OPSales.pdf");
            string spath = AppDomain.CurrentDomain.BaseDirectory;
            bool chkHeader = true;
           
            SqlParameter[] sqlParamPatDtls = new SqlParameter[]
            {
                    new SqlParameter("@BillNo", strBillNo),
                    new SqlParameter("@ACTIVITY", "GetPrint")
            };
            DataSet ds = obj.getDataset("tSales", DataHelper.SqlCmdType.sqlStoredProcedure, sqlParamPatDtls);

            if (ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
            {
                PharmacyReportPDF objPDfPrint = new PharmacyReportPDF();
                objPDfPrint.gvars = gvars;
                string pdfMem = objPDfPrint.ParamReport(ds, spath, chkHeader, "Pharmacy", "", "OPSales");


                //  MessageBox.Show("PDF report created successfully!", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo()
                {
                    FileName = pdfMem,
                    UseShellExecute = true
                });
            }
        }

        #endregion

        #region View
        private void btnView_Click(object sender, RoutedEventArgs e)
        {
            var button = sender as Button; 
            DataRowView pData=(DataRowView)button?.DataContext;
            if(pData != null)
            {
            SqlParameter[] sqlParamSearch = new SqlParameter[]
            {
                new SqlParameter("@ACTIVITY", "GetSingleData"),
                new SqlParameter("@BillNo", pData["BillNo"].ToString()),
                new SqlParameter("@LocationID", gvars.gLocationId),
                new SqlParameter("@DepartmentID", gvars.gDeptID)
            };
            DataSet dsSearch = obj.getDataset("tSales", DataHelper.SqlCmdType.sqlStoredProcedure, sqlParamSearch);
                if(dsSearch.Tables[0].Rows.Count>0 && dsSearch.Tables.Count > 0)
                {
                    txtUHIDPhone.Text = dsSearch.Tables[0].Rows[0]["UHID"].ToString();
                    txtUHID.Text = dsSearch.Tables[0].Rows[0]["UHID"].ToString();
                    if (dsSearch.Tables[0].Rows[0]["IPOPNo"].ToString() != "")
                    {
                        rdoOP.IsChecked = true;
                        //cmbSaleType.SelectedIndex = 1;
                        GetPatientDetails();
                    }
                    else
                    {
                        //cmbSaleType.SelectedIndex = 2;
                        rdoOthers.IsChecked = true;

                        if (txtUHIDPhone.Text != "")
                        {
                            GetPatientDetails();
                        }
                        else
                        {
                            txtName.Text = dsSearch.Tables[0].Rows[0]["PatientName"].ToString();
                            cmbDoctor.SelectedValue = dsSearch.Tables[0].Rows[0]["DocID"].ToString(); 
                        }
                    }
                    txtBillno.Text = dsSearch.Tables[0].Rows[0]["BillNo"].ToString();
                    dtpBillDt.Text = Convert.ToDateTime(dsSearch.Tables[0].Rows[0]["BillDate"]).ToString("dd-MMM-yyyy");
                    txtTotalAmt.Text = dsSearch.Tables[0].Rows[0]["TotalCharges"].ToString();
                    txtDiscount.Text = dsSearch.Tables[0].Rows[0]["Discount"].ToString();
                    txtPaidAmt.Text = dsSearch.Tables[0].Rows[0]["Paid"].ToString();
                    txtDueAmt.Text = dsSearch.Tables[0].Rows[0]["Due"].ToString();
                    txtConcAmt.Text = dsSearch.Tables[0].Rows[0]["Discount"].ToString();
                    txtDueReason.Text = dsSearch.Tables[0].Rows[0]["DueReason"].ToString();
                    cmbDueAuth.SelectedValue = dsSearch.Tables[0].Rows[0]["DueAuth"].ToString();
                    txtPayAmount.Text = dsSearch.Tables[0].Rows[0]["Paid"].ToString();
                    if (dsSearch.Tables[1].Rows.Count >0)
                    {
                        txtDiscReason.Text = dsSearch.Tables[1].Rows[0]["DiscReason"].ToString();
                        cmbDiscAuth.SelectedValue = dsSearch.Tables[1].Rows[0]["DiscAuth"].ToString();
                        txtConcPer.Text = dsSearch.Tables[1].Rows[0]["DiscPer"].ToString();
                        dgvItemDetails.ItemsSource = dsSearch.Tables[1].DefaultView;
                    }
                    if (dsSearch.Tables[2].Rows.Count > 0)
                    {
                        dgvPaymentDet.ItemsSource = dsSearch.Tables[2].DefaultView;
                    }
                    DisableControls();
                    GridPanel.Visibility = Visibility.Collapsed;
                    FormPanel.Visibility = Visibility.Visible;
                    btnNew.Content = "Back";
                }
            }
        }

        private void EnableControls()
        {
            rdoOP.IsEnabled = true;
            rdoOthers.IsEnabled = true;
            txtUHIDPhone.IsEnabled = true;
            grpPayDet.IsEnabled = true;
            grpPayDet.IsEnabled = true;
            grpItemDet.IsEnabled = true;
            btnSave.IsEnabled = true;
            txtTrnsCode.IsEnabled = true;
        }
        private void DisableControls()
        {
            rdoOP.IsEnabled = false;
            rdoOthers.IsEnabled = false;
            txtUHIDPhone.IsEnabled = false;
            grpPayDet.IsEnabled = false;
            grpPayDet.IsEnabled = false;
            grpItemDet.IsEnabled = false;
            btnSave.IsEnabled = false;
            txtTrnsCode.IsEnabled = false;
        }

        #endregion

        #region Find
        private void btnFind_Click(object sender, RoutedEventArgs e)
        {
            //if (cmbSaleType.SelectedIndex == 0)
            //{
            //    MessageBox.Show("Select SaleType","Information",MessageBoxButton.OK,MessageBoxImage.Information);
            //    cmbSaleType.Focus();
            //    return;
            //}
            if (txtUHIDPhone.Text != "")
            {
                GetPatientDetails();
            }
            else
            {
                UHIDPopUp.Visibility = Visibility.Visible;
                UHIDPopUp.IsOpen = true;
            }
        }

        private void GetPatientDetails()
        {
            if (txtUHIDPhone.Text != "")
            {
                string SaleType = string.Empty;
                if(rdoOP.IsChecked == true)
                {
                    SaleType = "OP";
                }
                else
                {
                    SaleType = "others";
                }
                       
                SqlParameter[] sqlParamSearch = new SqlParameter[]
                  {
                    new SqlParameter("@ACTIVITY", "GetPatientDtls"),
                    new SqlParameter("@Searchtxt", txtUHIDPhone.Text),
                    new SqlParameter("@Type", SaleType)
                  };
                ds = obj.getDataset("tSales", DataHelper.SqlCmdType.sqlStoredProcedure, sqlParamSearch);
                if (ds.Tables[0].Rows.Count > 0)
                {
                    txtName.Text = ds.Tables[0].Rows[0]["Name"].ToString();
                    txtAge.Text = ds.Tables[0].Rows[0]["Age"].ToString();
                    cmbDoctor.SelectedValue = ds.Tables[0].Rows[0]["DocId"];
                    txtPhone.Text = ds.Tables[0].Rows[0]["Phone"].ToString();
                    txtUHID.Text = ds.Tables[0].Rows[0]["UHID"].ToString();
                    txtOPDNO.Text = ds.Tables[0].Rows[0]["OPDNO"].ToString();
                }
            }

        }

        #endregion

        #region BindDoctor
        public void BindDoctor()
        {
            SqlParameter[] sqlParamLoc = new SqlParameter[]
            {
            new SqlParameter("@ACTIVITY", "GetDoc")
            };
            DataSet dsLoc = obj.getDataset("tSales", DataHelper.SqlCmdType.sqlStoredProcedure, sqlParamLoc);
            if (dsLoc.Tables.Count > 0 && dsLoc.Tables[0].Rows.Count > 0)
            {
                DataRow drow;
                drow = dsLoc.Tables[0].NewRow();
                drow["DocId"] = "--Select--";
                drow["DocName"] = "--Select--";
                dsLoc.Tables[0].Rows.InsertAt(drow, 0);

                cmbDoctor.SelectedValuePath = "DocId";
                cmbDoctor.DisplayMemberPath = "DocName";
                cmbDoctor.ItemsSource = dsLoc.Tables[0].DefaultView;
                cmbDoctor.SelectedIndex = 0;
            } 
        }
        #endregion

        #region Bind Item
        public void BindItem()
        {
            // Load items from the database
            LoadItemsFromDatabase();

            // Populate FilteredItems with AllItems
            foreach (var item in AllItems)
            {
                FilteredItems.Add(item);
            }
        }
        private void LoadItemsFromDatabase()
        {
            SqlParameter[] sqlParamLoc = new SqlParameter[]
            {
                new SqlParameter("@IsMedical", true),
                new SqlParameter("@ACTIVITY", "GetItem")
            };

            ds = obj.getDataset("tSales", DataHelper.SqlCmdType.sqlStoredProcedure, sqlParamLoc);

            if (ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
            {
                AllItems.Clear();
                for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                {
                    Item item = CH.ConvertDataTableToObjects<Item>(ds.Tables[0].Rows[i]);
                    AllItems.Add(item);
                }
            }
        }
      
        private void txtItem_GotFocus(object sender, RoutedEventArgs e)
        {
            if (FilteredItems.Count > 0)
            {
                popup.IsOpen = true;
            }
        }
        private void txtItem_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtItem.Text))
            {
                txtItem.Tag = "";
                ClearItems();
            }
            if (isItemSelected)
            {
                isItemSelected = false;
                popup.IsOpen = false;
                return;
            }

            string text = txtItem.Text.Trim().ToLower();

            var result = AllItems.Where(x => !string.IsNullOrEmpty(x.ItemName) && x.ItemName.ToLower().Contains(text)).ToList();
            FilteredItems.Clear();
            foreach (var item in result)
                FilteredItems.Add(item);

            popup.IsOpen = FilteredItems.Count > 0 && !string.IsNullOrWhiteSpace(text);
        }
      

        private void txtItem_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Down && FilteredItems.Count > 0)
            {
                if (!popup.IsOpen)
                    popup.IsOpen = true;

                lstItems.SelectedIndex = 0;
                lstItems.ScrollIntoView(lstItems.SelectedItem);

                Dispatcher.BeginInvoke(new Action(() =>
                {
                    lstItems.Focus();
                    ListBoxItem item = lstItems.ItemContainerGenerator.ContainerFromIndex(lstItems.SelectedIndex) as ListBoxItem;
                    if (item != null)
                        item.Focus();
                }), DispatcherPriority.Background);

                e.Handled = true;
            }
        }
        private void SetSelectedItem(Item selected)
        { 
            if (selected == null)
            {
                ClearItems();
                return;
            }
            isItemSelected = true;
            txtItem.Text = selected.ItemName;
            txtItem.Tag = selected.ItemID;
            txtItem.CaretIndex = txtItem.Text.Length;

            popup.IsOpen = false; 
            HandleSelectedMedicine();
        }
        private void lstItems_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                Item selected = lstItems.SelectedItem as Item;
                SetSelectedItem(selected); 

                popup.IsOpen = false;
                HandleSelectedMedicine();
                e.Handled = true;
            }
            else if (e.Key == Key.Escape)
            {
                popup.IsOpen = false;
                txtItem.Focus();
                e.Handled = true;
            }
        }
        private void lstItems_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            Item selected = lstItems.SelectedItem as Item;
            SetSelectedItem(selected);

            //bdList.Visibility = Visibility.Collapsed;
            popup.IsOpen = false;
            HandleSelectedMedicine();
        }
        #endregion
         
        #region HandleSelectedMedicine
        private void HandleSelectedMedicine()
        {
            if (txtItem.Tag != null)
            {
                string itemId = txtItem.Tag.ToString();

                if (!string.IsNullOrWhiteSpace(itemId) && itemId != "--Select--" && itemId != "0")
                {
                    GetBatch(itemId);
                    txtQty.Focus();
                }
                else
                {
                    ClearItems();
                }
            }
            else
            {
                ClearItems();
            }
        }
        #endregion
          

        #region txtQty_PreviewKeyDown
        private void txtQty_PreviewKeyDown(object sender, KeyEventArgs e)
            {
                if (e.Key == Key.Tab)
                {
                    // Move focus to btnAdd
                    btnAdd.Focus();
                    Keyboard.Focus(btnAdd);

                    e.Handled = true; // prevent default Tab navigation
                }
            }

        #endregion

        #region BindWallet
        public void BindWallet()
        {
            SqlParameter[] sqlParamLoc = new SqlParameter[]
            {
            new SqlParameter("@ACTIVITY", "GetWallet")
            };
            DataSet dsLoc = obj.getDataset("tConsultation", DataHelper.SqlCmdType.sqlStoredProcedure, sqlParamLoc);
            if (dsLoc.Tables.Count > 0 && dsLoc.Tables[0].Rows.Count > 0)
            {
                cmbWalletType.SelectedValuePath = "LookupId";
                cmbWalletType.DisplayMemberPath = "LookupText";
                cmbWalletType.ItemsSource = dsLoc.Tables[0].DefaultView;
                cmbWalletType.SelectedIndex = 0;
            } 
        }
        #endregion

        #region GetUserDiscAuthorized
        public void GetUserDiscAuthorized()
        {
            SqlParameter[] sqlParamDep = new SqlParameter[]
            {
            new SqlParameter("@ACTIVITY", "GetAuthorisation")
            };
            DataSet dsSpl = obj.getDataset("tSales", DataHelper.SqlCmdType.sqlStoredProcedure, sqlParamDep);
            if (dsSpl.Tables.Count > 0 && dsSpl.Tables[0].Rows.Count > 0)
            {
                DataRow drow;
                drow = dsSpl.Tables[0].NewRow();
                drow["UserID"] = "--Select--";
                drow["EmployeeName"] = "--Select--";
                dsSpl.Tables[0].Rows.InsertAt(drow, 0);
                cmbDiscAuth.SelectedValuePath = "UserID";
                cmbDiscAuth.DisplayMemberPath = "EmployeeName";
                cmbDiscAuth.ItemsSource = dsSpl.Tables[0].DefaultView;
                //cmbDiscAuth.Items.Insert(0, "--Select--");
            } 
        }
        #endregion

        #region BindDueAuthorisation
        public void BindDueAuthorisation()
        {
            SqlParameter[] sqlParamLoc = new SqlParameter[]
            {
            new SqlParameter("@ACTIVITY", "GetAuthorisation")
            };
            DataSet dsLoc = obj.getDataset("tSales", DataHelper.SqlCmdType.sqlStoredProcedure, sqlParamLoc);
            if (dsLoc.Tables.Count > 0 && dsLoc.Tables[0].Rows.Count > 0)
            {
                DataRow drow;
                drow = dsLoc.Tables[0].NewRow();
                drow["UserID"] = "--Select--";
                drow["EmployeeName"] = "--Select--";
                dsLoc.Tables[0].Rows.InsertAt(drow, 0);
                cmbDueAuth.SelectedValuePath = "UserID";
                cmbDueAuth.DisplayMemberPath = "EmployeeName";
                cmbDueAuth.ItemsSource = dsLoc.Tables[0].DefaultView;
               // cmbDueAuth.Items.Insert(0, "--Select--");
            } 
        }

        #endregion

        #region Bind Payment Mode
        public void BindPayMode()
        {
            SqlParameter[] sqlParamLoc = new SqlParameter[]
            {
            new SqlParameter("@ACTIVITY", "GetPayMode")
            };
            DataSet dsLoc = obj.getDataset("tConsultation", DataHelper.SqlCmdType.sqlStoredProcedure, sqlParamLoc);
            if (dsLoc.Tables.Count > 0 && dsLoc.Tables[0].Rows.Count > 0)
            {
                cmbPayMode.SelectedValuePath = "LookupId";
                cmbPayMode.DisplayMemberPath = "LookupText";
                cmbPayMode.ItemsSource = dsLoc.Tables[0].DefaultView; 
                cmbPayMode.SelectedIndex=0;
            } 
        }
        #endregion

        #region GetBank
        public void GetBank()
        {
            SqlParameter[] sqlParamDep = new SqlParameter[]
            {
            new SqlParameter("@ACTIVITY", "GetBank")
            };
            DataSet dsSpl = obj.getDataset("tRegistration", DataHelper.SqlCmdType.sqlStoredProcedure, sqlParamDep);
            if (dsSpl.Tables.Count > 0 && dsSpl.Tables[0].Rows.Count > 0)
            {
                cmbBank.SelectedValuePath = "BankID";
                cmbBank.DisplayMemberPath = "BankName";
                cmbBank.ItemsSource = dsSpl.Tables[0].DefaultView;
                cmbBank.SelectedIndex = 0;
            } 
        }
        #endregion

        #region Add Item
        private void btnAdd_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(txtItem.Text) || txtItem.Tag == null)
                {
                    MessageBox.Show("Select Item ", "Information", MessageBoxButton.OK, MessageBoxImage.Information);
                    txtItem.Focus();
                    return;
                } 
                decimal qty = 0; 
                if (!decimal.TryParse(txtQty.Text, out qty) || qty <= 0)
                {
                    MessageBox.Show("Required Quantity should be greater than 0 ", "Information", MessageBoxButton.OK, MessageBoxImage.Information);
                    txtQty.Focus();
                    return;
                }
                if (cmbBatchNo.SelectedItem == null)
                {
                    MessageBox.Show("Batch No Cannot Be Blank ", "Information", MessageBoxButton.OK, MessageBoxImage.Information);
                    cmbBatchNo.Focus();
                    return;
                }
              
                if (dgvItemDetails.Items.Count > 0)
                {
                    foreach (var item in dgvItemDetails.Items)
                    {
                        AddItems p = item as AddItems;
                        if (p == null) continue;
                        if (p.ItemId == txtItem.Tag.ToString() && p.BatchNo == cmbBatchNo.Text && p.ExpiryDate == dtpExpDt.Text && p.UnitRate == Rate && p.UnitMrp == Convert.ToDecimal(txtMRP.Text) && p.PurchBillNo == PBillNo && p.SupplierID == SupId)
                        {
                            MessageBox.Show("You have already added this Item ", "Information", MessageBoxButton.OK, MessageBoxImage.Information);
                            txtItem.Focus();
                            return;
                        }
                    }
                }
                 
                decimal SavedQty = 0, reqStock = Convert.ToDecimal(txtQty.Text);


                if (Convert.ToDecimal(txtBQty.Text) >= Convert.ToDecimal(txtQty.Text))
                {
                    strQry = "SELECT ISNULL(SUM(s.Qty),0) AS Qty FROM trnStock s " +
                         " WHERE S.LocationID = '" + gvars.gLocationId + "' AND UnitMrp=" + txtMRP.Text + " AND UnitRate=" + Rate + " AND BatchNo='" + cmbBatchNo.Text + "' " +
                         " AND ExpiryDate='" + dtpExpDt.SelectedDate.Value.ToString("yyyy-MM-dd") + "' AND SupplierID='" + SupId + "' AND PurchBillNo='" + PBillNo + "' AND ItemID='" + txtItem.Tag + "' " +
                         " AND DepartmentID='" + gvars.gDeptID + "'";
                    decimal cStock = Convert.ToDecimal(obj.ExecuteScalar(strQry, DataHelper.SqlCmdType.sqlText));
                    if (cStock >= Convert.ToInt32(txtQty.Text))
                    {
                        string Medid = txtItem.Tag.ToString();
                        aItem.Add(new AddItems
                        {
                            Slno = (dgvItemDetails.Items.Count + 1),
                            ItemId = Medid.ToString(),
                            ItemName = txtItem.Text.ToString(),
                            Rack = "",
                            Tray = "",
                            Qty = Convert.ToInt32(txtQty.Text),
                            UnitMrp = Convert.ToDecimal(txtMRP.Text),
                            UnitRate = Rate,
                            BatchNo = cmbBatchNo.Text.ToString(),
                            ExpiryDate = dtpExpDt.Text.ToString(),
                            Discount = 0,
                            DiscPer = 0,
                            Amount = Convert.ToDecimal(txtAmount.Text),
                            TaxPer = taxPer,
                            TaxAmount = Convert.ToDecimal(txtTaxAmt.Text),
                            SupplierID = SupId,
                            PurchBillNo = PBillNo,
                            BatchQty = Convert.ToDecimal(txtBQty.Text),
                            StockQty = Convert.ToDecimal(txtBQty.Text) - Convert.ToDecimal(txtQty.Text),
                            IsHigh = isHigh,
                            IsLook = isLook,
                            IsSound = isSound
                        });
                        dgvItemDetails.ItemsSource = aItem;
                        DataContext = this;
                        UpdatTotals();
                        ClearItems();
                        CalculateConc();
                        txtItem.Text = "";
                        txtItem.Tag = null;
                        txtItem.Focus();
                    }
                    else
                    {
                        MessageBox.Show("No Stock available", "Information", MessageBoxButton.OK, MessageBoxImage.Information);
                        txtItem.Focus();
                        return;
                    }
                }
                else
                {
                    for (int j = 0; j < dsBatch.Tables[0].Rows.Count - 1; j++)
                    {
                        strQry = "SELECT ISNULL(SUM(s.Qty),0) AS Qty FROM trnStock s " +
                          " WHERE S.LocationID = '" + gvars.gLocationId + "' AND UnitMrp=" + dsBatch.Tables[0].Rows[j]["UnitMrp"] + " AND UnitRate=" + dsBatch.Tables[0].Rows[j]["UnitRate"] + " AND BatchNo='" + dsBatch.Tables[0].Rows[j]["BatchNo"] + "' " +
                          " AND ExpiryDate='" + Convert.ToDateTime(dsBatch.Tables[0].Rows[j]["ExpiryDate"]).ToString("yyyy-MM-dd") + "' AND SupplierID='" + dsBatch.Tables[0].Rows[j]["SupplierID"] + "' " +
                          " AND PurchBillNo='" + dsBatch.Tables[0].Rows[j]["PurchBillNo"] + "' AND ItemID='" + txtItem.Tag + "' " +
                             " AND DepartmentID='" + gvars.gDeptID + "'";

                        decimal cStock = Convert.ToDecimal(obj.ExecuteScalar(strQry, DataHelper.SqlCmdType.sqlText));

                        if (SavedQty < Convert.ToDecimal(txtQty.Text))
                        {
                            if (reqStock >= cStock)
                            {
                                SavedQty = SavedQty + cStock;
                                reqStock = reqStock - cStock;

                                aItem.Add(new AddItems
                                {
                                    Slno = (dgvItemDetails.Items.Count + 1),
                                    ItemId = txtItem.Tag.ToString(),
                                    ItemName = txtItem.Text.ToString(),
                                    Rack = "",
                                    Tray = "",
                                    Qty = cStock,
                                    UnitMrp = Convert.ToDecimal(dsBatch.Tables[0].Rows[j]["UnitMrp"]),
                                    UnitRate = Convert.ToDecimal(dsBatch.Tables[0].Rows[j]["UnitRate"]),
                                    BatchNo = dsBatch.Tables[0].Rows[j]["BatchNo"].ToString(),
                                    ExpiryDate = Convert.ToDateTime(dsBatch.Tables[0].Rows[j]["ExpiryDate"]).ToString("dd-MMM-yyyy"),
                                    Discount = 0,
                                    DiscPer = 0,
                                    Amount = Math.Round(Convert.ToDecimal(dsBatch.Tables[0].Rows[j]["UnitMrp"]) * SavedQty),
                                    TaxPer = Convert.ToDecimal(dsBatch.Tables[0].Rows[j]["TaxPer"]),
                                    TaxAmount = Math.Round(((Convert.ToDecimal(dsBatch.Tables[0].Rows[j]["UnitMrp"]) * SavedQty) * taxPer) / 100),
                                    SupplierID = dsBatch.Tables[0].Rows[j]["SupplierID"].ToString(),
                                    PurchBillNo = dsBatch.Tables[0].Rows[j]["PurchBillNo"].ToString(),
                                    BatchQty = cStock,
                                    StockQty = cStock - SavedQty,
                                    IsHigh = isHigh,
                                    IsLook = isLook,
                                    IsSound = isSound
                                });
                                dgvItemDetails.ItemsSource = aItem;
                                DataContext = this;
                                SavedQty = 0;
                                UpdatTotals();
                                CalculateConc();
                                txtItem.Text = "";
                                txtItem.Tag = "";
                                txtItem.Focus();
                            }
                            else
                            {
                                SavedQty = SavedQty + reqStock;
                                aItem.Add(new AddItems
                                {
                                    Slno = (dgvItemDetails.Items.Count + 1),
                                    ItemId = txtItem.Tag.ToString(),
                                    ItemName = txtItem.Text.ToString(),
                                    Rack = "",
                                    Tray = "",
                                    Qty = reqStock,
                                    UnitMrp = Convert.ToDecimal(dsBatch.Tables[0].Rows[j]["UnitMrp"]),
                                    UnitRate = Convert.ToDecimal(dsBatch.Tables[0].Rows[j]["UnitRate"]),
                                    BatchNo = dsBatch.Tables[0].Rows[j]["BatchNo"].ToString(),
                                    ExpiryDate = Convert.ToDateTime(dsBatch.Tables[0].Rows[j]["ExpiryDate"]).ToString("dd-MMM-yyyy"),
                                    Discount = 0,
                                    DiscPer = 0,
                                    Amount = Math.Round(Convert.ToDecimal(dsBatch.Tables[0].Rows[j]["UnitMrp"]) * reqStock),
                                    TaxPer = Convert.ToDecimal(dsBatch.Tables[0].Rows[j]["TaxPer"]),
                                    TaxAmount = Math.Round(((Convert.ToDecimal(dsBatch.Tables[0].Rows[j]["UnitMrp"]) * reqStock) * taxPer) / 100),
                                    SupplierID = dsBatch.Tables[0].Rows[j]["SupplierID"].ToString(),
                                    PurchBillNo = dsBatch.Tables[0].Rows[j]["PurchBillNo"].ToString(),
                                    BatchQty = cStock,
                                    StockQty = cStock - reqStock,
                                    IsHigh = isHigh,
                                    IsLook = isLook,
                                    IsSound = isSound
                                });
                                dgvItemDetails.ItemsSource = aItem;
                                reqStock = 0;
                                SavedQty = 0;
                                DataContext = this;
                                UpdatTotals();
                                CalculateConc();
                                txtItem.Text = "";
                                txtItem.Tag = "";
                                txtItem.Focus();
                            }
                            if (reqStock == 0)
                            {
                                ClearItems();
                                return;
                            }
                        }
                        else
                        {
                            MessageBox.Show("No Stock available", "Information", MessageBoxButton.OK, MessageBoxImage.Information);
                            ClearItems();
                            txtItem.Focus();
                            return;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    ex.Message + "\n\n" + ex.StackTrace,
                    "Error",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
        }        
       
        private void ClearItems()
        {
            txtItem.Tag="";
            txtItem.Text = "";
            txtBQty.Text = "";
            txtAmount.Text=txtQty.Text=txtTaxAmt.Text=txtMRP.Text = txtTQty.Text= Convert.ToDecimal(0).ToString();
            cmbBatchNo.ItemsSource=null;
            PBillNo = "";
            SupId = "";
            Rate = 0;
            taxPer = 0;
        }

        #endregion

        #region DeleteItem
        private void btnDelete_Click(object sender, RoutedEventArgs e)
        {
                var button = sender as Button;
                var itm = button?.DataContext; 
                AddItems dr = (AddItems)itm;

            if (dr != null)
                {
                var result = MessageBox.Show(
                    "Are you sure you want to delete?",
                    "Confirm Delete",
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Warning);

                if (result == MessageBoxResult.Yes)
                {
                    if (aItem != null && aItem.Contains(dr))
                        aItem.Remove(dr);
                }
            }
        }
        #endregion

        //Add Discount in Grid

        #region Payment Add
        private void btnPayAdd_Click(object sender, RoutedEventArgs e)
        {
            if (PayAddValidation() == false)
            {
                return;
            }
            else
            {
                if (dgvPaymentDet.Items.Count > 0)
                {
                    foreach (var item in dgvPaymentDet.Items)
                    {
                        AddPayDet p = item as AddPayDet;
                        if (p.PayMode == cmbPayMode.Text)
                        {
                            MessageBox.Show("The Payment Mode Already Exist ", "Information", MessageBoxButton.OK, MessageBoxImage.Information);
                            cmbPayMode.Focus();
                            return;
                        }
                    }
                }
                if (cmbPayMode.Text == "Card")
                {
                    payDet.Add(new AddPayDet
                    {
                        Slno = (dgvPaymentDet.Items.Count + 1),
                        PayMode = cmbPayMode.Text,
                        Amount = Convert.ToDecimal(txtPayAmount.Text),
                        WalletAccount = "",
                        Bank = cmbBank.Text,
                        RefDt = "",
                        TransNo = txtRefNo.Text
                    });
                }
                else if (cmbPayMode.Text == "Cheque")
                {
                    payDet.Add(new AddPayDet
                    {
                        Slno = (dgvPaymentDet.Items.Count + 1),
                        PayMode = cmbPayMode.Text,
                        Amount = Convert.ToDecimal(txtPayAmount.Text),
                        WalletAccount = "",
                        Bank = cmbBank.Text,
                        RefDt = dtpCheqDt.Text,
                        TransNo = txtRefNo.Text
                    });
                }
                else if (cmbPayMode.Text == "UPI")
                {
                    payDet.Add(new AddPayDet
                    {
                        Slno = (dgvPaymentDet.Items.Count + 1),
                        PayMode = cmbPayMode.Text,
                        Amount = Convert.ToDecimal(txtPayAmount.Text),
                        WalletAccount = cmbWalletType.Text,
                        Bank = "",
                        RefDt = "",
                        TransNo = txtRefNo.Text
                    });
                }
                else
                {
                    payDet.Add(new AddPayDet
                    {
                        Slno = (dgvPaymentDet.Items.Count + 1),
                        PayMode = cmbPayMode.Text,
                        Amount = Convert.ToDecimal(txtPayAmount.Text),
                        WalletAccount = "",
                        Bank = "",
                        RefDt = "",
                        TransNo = ""
                    });
                }
                dgvPaymentDet.ItemsSource = payDet;
                txtPayAmount.Text = txtRefNo.Text = "";
                cmbBank.SelectedIndex = 0;
                cmbWalletType.SelectedIndex = 0;
                cmbPayMode.SelectedIndex = 0;
                txtConcPer.IsEnabled = false;
                txtConcAmt.IsEnabled = false;
                UpdateDue();
                lblCardNo.Visibility = Visibility.Collapsed;
                cmbBank.Visibility = Visibility.Collapsed;
                txtRefNo.Visibility = Visibility.Collapsed;
                lblBank.Visibility = Visibility.Collapsed;
                lblCheqDt.Visibility = Visibility.Collapsed;
                dtpCheqDt.Visibility = Visibility.Collapsed;
                lblCheqNo.Visibility = Visibility.Collapsed;
                lblReferenceNo.Visibility = Visibility.Collapsed;
                lblWallet.Visibility = Visibility.Collapsed;
                cmbWalletType.Visibility = Visibility.Collapsed;
            }
        }

        private bool PayAddValidation()
        {
            bool status = true;

            if (txtPayAmount.Text == "" || txtPayAmount.Text == "0")
            {
                MessageBox.Show("Enter Pay Amount", "Information", MessageBoxButton.OK, MessageBoxImage.Information);
                txtPayAmount.Focus();
                return false;
            }
            if (cmbPayMode.SelectedIndex == -1)
            {
                MessageBox.Show("Select Payment Mode", "Information", MessageBoxButton.OK, MessageBoxImage.Information);
                cmbPayMode.Focus();
                return false;
            }
            if (cmbPayMode.Text == "Card")
            {
                if (txtRefNo.Text == "")
                {
                    MessageBox.Show("Enter Reference No", "Information", MessageBoxButton.OK, MessageBoxImage.Information);
                    txtRefNo.Focus();
                    return false;
                }
                if (cmbBank.SelectedIndex == -1)
                {
                    MessageBox.Show("Select Reference bank", "Information", MessageBoxButton.OK, MessageBoxImage.Information);
                    cmbBank.Focus();
                    return false;
                }
            }
            if (cmbPayMode.Text == "Cheque")
            {
                if (txtRefNo.Text == "")
                {
                    MessageBox.Show("Enter Reference No", "Information", MessageBoxButton.OK, MessageBoxImage.Information);
                    txtRefNo.Focus();
                    return false;
                }
                if (cmbBank.SelectedIndex == -1)
                {
                    MessageBox.Show("Select Reference bank", "Information", MessageBoxButton.OK, MessageBoxImage.Information);
                    cmbBank.Focus();
                    return false;
                }
                if (dtpCheqDt.Text == "")
                {
                    MessageBox.Show("Select Cheque Date", "Information", MessageBoxButton.OK, MessageBoxImage.Information);
                    dtpCheqDt.Focus();
                    return false;
                }
            }
            if (cmbPayMode.Text == "UPI")
            {
                if (txtRefNo.Text == "")
                {
                    MessageBox.Show("Enter Reference No", "Information", MessageBoxButton.OK, MessageBoxImage.Information);
                    txtRefNo.Focus();
                    return false;
                }
                if (cmbWalletType.SelectedIndex == -1)
                {
                    MessageBox.Show("Select Wallet Type", "Information", MessageBoxButton.OK, MessageBoxImage.Information);
                    cmbWalletType.Focus();
                    return false;
                }
            }

            return status;
        }
       
        #endregion

        #region btnPayDelete_Click
        private void btnPayDelete_Click(object sender, RoutedEventArgs e)
        {
            var button = sender as Button;
            var p = button?.DataContext;
            AddPayDet dr = (AddPayDet)p;

            if (dr != null)
            {
                var result = MessageBox.Show(
                    "Are you sure you want to delete?",
                    "Confirm Delete",
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Warning);

                if (result == MessageBoxResult.Yes)
                {
                    if (payDet != null && payDet.Contains(dr))
                        payDet.Remove(dr);
                    UpdateDue();
                    if (payDet.Count == 0)
                    {
                        PayResetDetails();
                        txtConcPer.IsEnabled = true;
                        txtConcAmt.IsEnabled = true;
                    }
                }
            }
        }
        #endregion

        #region Launch Calculator
        private void btnCalc_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Process.Start("calc.exe");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to open calculator: " + ex.Message);
            }
        }
        #endregion

        #region Save
        private bool SaveValidation()
        {
            bool Valid = true;
            if (txtName.Text == "")
            {
                MessageBox.Show("Enter Name ", "Information", MessageBoxButton.OK, MessageBoxImage.Information);
                txtName.Focus();
                Valid = false;
            }
            //if (cmbSaleType.SelectedIndex == 1)
            if (rdoOP.IsChecked == true)
            {
                if (txtUHID.Text == "")
                {
                    MessageBox.Show("Enter UHID ", "Information", MessageBoxButton.OK, MessageBoxImage.Information);
                    txtUHID.Focus();
                    Valid = false;
                }
            }
            if (cmbDoctor.SelectedIndex == 0)
            {
                MessageBox.Show("Select Doctor ", "Information", MessageBoxButton.OK, MessageBoxImage.Information);
                cmbDoctor.Focus();
                Valid = false;
            }
            if (dgvItemDetails.Items.Count == 0)
            {
                MessageBox.Show("There is No Items to save ", "Information", MessageBoxButton.OK, MessageBoxImage.Information);
                txtItem.Focus();
                Valid = false;
            }
            if (Convert.ToDecimal(txtDueAmt.Text) > 0)
            {
                if (gvars.gDueType == "%")
                {
                    decimal Dper = (Convert.ToDecimal(txtDueAmt.Text) / Convert.ToDecimal(txtTotalAmt.Text)) * 100;
                    if (Dper > gvars.gDueLimit)
                    {
                        MessageBox.Show("You Are Not Allowed To Give Due More than " + gvars.gDueLimit.ToString() + " % ", "Information", MessageBoxButton.OK, MessageBoxImage.Information);
                        Valid = false;
                    }
                }
                {
                    var TotalDue = (((Convert.ToDecimal(txtDueAmt.Text)) * gvars.gDueLimit) / 100);
                    if (Convert.ToDecimal(txtDueAmt.Text) > TotalDue)
                    //    if (Convert.ToDecimal(txtDueAmt.Text) > gvars.gDueLimit)
                    {

                        MessageBox.Show("You Are Not Allowed To Give Due More than " + gvars.gDueLimit.ToString() + " Rs. ", "Information", MessageBoxButton.OK, MessageBoxImage.Information);
                        Valid = false;
                    }
                }
                if (cmbDueAuth.SelectedIndex == 0)
                {
                    MessageBox.Show("Select Due Auth ", "Information", MessageBoxButton.OK, MessageBoxImage.Information);
                    cmbDueAuth.Focus();
                    Valid = false;
                }
                if (txtDueReason.Text == "")
                {
                    MessageBox.Show("Enter Due Reason ", "Information", MessageBoxButton.OK, MessageBoxImage.Information);
                    txtDueReason.Focus();
                    Valid = false;
                }
            }
            if (Convert.ToDecimal(txtDiscount.Text) > 0)
            {
                if (cmbDiscAuth.SelectedIndex == 0)
                {
                    MessageBox.Show("Select Discount Auth ", "Information", MessageBoxButton.OK, MessageBoxImage.Information);
                    cmbDiscAuth.Focus();
                    Valid = false;
                }
                if (txtDiscReason.Text == "")
                {
                    MessageBox.Show("Enter Discount Reason ", "Information", MessageBoxButton.OK, MessageBoxImage.Information);
                    txtDiscReason.Focus();
                    Valid = false;
                }
            }
            if (txtTrnsCode.Text == "")
            {
                MessageBox.Show("Enter Transaction Code ", "Information", MessageBoxButton.OK, MessageBoxImage.Information);
                txtTrnsCode.Focus();
                Valid = false;
            }
            if (txtTrnsCode.Text != gvars.gUserCode)
            {
                MessageBox.Show("Invalid Transaction Code ", "Information", MessageBoxButton.OK, MessageBoxImage.Information);
                txtTrnsCode.Text = "";
                txtTrnsCode.Focus();
                Valid = false;
            }

            return Valid;
        }

        private void btnSave_Click(object sender, RoutedEventArgs e)
        {
            if (SaveValidation() == false)
            {
                return;
            }

            string BillNo = Helper.GenerateIPSaleBillNO(gvars.gLocationId).ToString();
            SqlConnection con = new SqlConnection();
            obj.OpenDBCon();
            con = obj.getConnection();
            SqlTransaction sqlTran = con.BeginTransaction(IsolationLevel.ReadCommitted);
            var PaymentID = Helper.GetPaymnetID();

            int i = 0;
            try
            {
                int OPrintCont = 0;
                int DPrintCount = 0;

                SqlParameter[] sqlParamInsert = new SqlParameter[]
                {
                    new SqlParameter("@UHID", txtUHID.Text),
                    new SqlParameter("@BillNo", BillNo),
                    new SqlParameter("@DepartmentID", gvars.gDeptID),
                    new SqlParameter("@IPOPNo", txtOPDNO.Text),
                    new SqlParameter("@DocID", cmbDoctor.SelectedValue),
                    new SqlParameter("@OrganisationID", OrganisationID),
                    new SqlParameter("@Remarks", txtDueReason.Text),
                    new SqlParameter("@IsIP", false),
                    new SqlParameter("@TotalCharges", Convert.ToDecimal(txtTotalAmt.Text)),
                    new SqlParameter("@Paid", Convert.ToDecimal(txtPaidAmt.Text)),
                    new SqlParameter("@Discount", Convert.ToDecimal(txtDiscount.Text)),
                    new SqlParameter("@PostDiscount",0),
                    new SqlParameter("@Due", Convert.ToDecimal(txtDueAmt.Text)),
                    new SqlParameter("@DueAuth", cmbDueAuth.SelectedValue),
                    new SqlParameter("@DueReason", txtDueReason.Text),
                    new SqlParameter("@OPrintCont", OPrintCont),
                    new SqlParameter("@DPrintCount", DPrintCount),
                    new SqlParameter("@PatientName", txtName.Text),
                    new SqlParameter("@Phone", txtPhone.Text),
                    new SqlParameter("@LocationID", gvars.gLocationId),
                    new SqlParameter("@TerminalID",gvars.gTermId),
                    new SqlParameter("@CreateUserID", gvars.gUserID),
                    new SqlParameter("@ACTIVITY", "Insert")
                };

                i = obj.ExecuteNonQuery("tSales", DataHelper.SqlCmdType.sqlStoredProcedure, sqlParamInsert, sqlTran);
                 //IN sales we added transaction in the Insert because it is not commited  At the time of details to check the 

                if (aItem.Count > 0 && aItem != null)
                {
                    decimal CGST = 0, SGST = 0, IGST = 0, Taxable = 0;

                    foreach (var Dtls in aItem)
                    {
                        decimal unitRate = Convert.ToDecimal(Dtls.UnitRate);
                        decimal unitMrp = Convert.ToDecimal(Dtls.UnitMrp);
                        DateTime expiry = Convert.ToDateTime(Dtls.ExpiryDate);

                        decimal stockQty = CommonService.CheckStockQty(Dtls.ItemId, unitRate, unitMrp, Dtls.BatchNo, expiry, gvars.gDeptID, Dtls.SupplierID, gvars.gLocationId, sqlTran);

                        if (Dtls.Qty > stockQty)
                        {
                            MessageBox.Show($"Your Enter Qty Is More Then Original Stock ", "warning", MessageBoxButton.OK, MessageBoxImage.Error);

                            foreach (var row in dgvItemDetails.Items)
                            {
                                if (row == null || row == CollectionView.NewItemPlaceholder)
                                    continue;


                                var item = row as AddItems;
                                if (item == null)
                                    continue;


                                string itemID = item.ItemId;
                                string batchNo = item.BatchNo; //System.Net.WebUtility.HtmlDecode(item.BatchNo) 
                                string supplierID = item.SupplierID;
                                string departmentID = gvars.gDeptID;

                                if (Dtls.ItemId == itemID && Dtls.BatchNo == batchNo && gvars.gDeptID == departmentID && Dtls.SupplierID == supplierID)
                                {

                                    item.Qty = stockQty;
                                    item.BatchQty = stockQty;
                                    item.IsQtyHighlighted = true;
                                    dgvItemDetails.UpdateLayout();
                                    dgvItemDetails.ScrollIntoView(item);
                                    int qtyColumnIndex = 4;
                                    dgvItemDetails.CurrentCell = new DataGridCellInfo(item, dgvItemDetails.Columns[qtyColumnIndex]);
                                    dgvItemDetails.BeginEdit();
                                    sqlTran.Rollback();
                                    return;
                                }
                            }

                            sqlTran.Rollback();
                            return;
                        }
                        else
                        {
                            CGST = SGST = Dtls.TaxAmount / 2;
                            Taxable = (Dtls.Amount / (1 + Dtls.TaxPer / 100));
                            SqlParameter[] sqlParamGrnDtls = new SqlParameter[]
                            {
                                new SqlParameter("@BillNo",BillNo),
                                new SqlParameter("@ItemID", Dtls.ItemId),
                                new SqlParameter("@Qty", Dtls.Qty),
                                new SqlParameter("@UnitRate", Dtls.UnitRate),
                                new SqlParameter("@UnitMrp", Dtls.UnitMrp),
                                new SqlParameter("@BatchNo", Dtls.BatchNo),
                                new SqlParameter("@ExpiryDate", Dtls.ExpiryDate),
                                new SqlParameter("@TaxPer", Dtls.TaxPer),
                                new SqlParameter("@TaxAmount",(CGST+SGST)),
                                new SqlParameter("@SupplierID", Dtls.SupplierID),
                                new SqlParameter("@PurchBillNo", Dtls.PurchBillNo),
                                new SqlParameter("@Amount", Dtls.Amount),
                                new SqlParameter("@Discount", Dtls.Discount),
                                new SqlParameter("@DiscPer", Dtls.DiscPer),
                                new SqlParameter("@DiscAuth", cmbDiscAuth.SelectedValue),
                                new SqlParameter("@DiscReason", txtDiscReason.Text),
                                new SqlParameter("@SGSTPer", Dtls.TaxPer/2),
                                new SqlParameter("@SGSTAmt", SGST),
                                new SqlParameter("@CGSTPer", Dtls.TaxPer/2),
                                new SqlParameter("@CGSTAmt", CGST),
                                new SqlParameter("@Taxable", Taxable),
                                new SqlParameter("@LocationID", gvars.gLocationId),
                                new SqlParameter("@CreateUserID", gvars.gUserID),
                                new SqlParameter("@TerminalID", gvars.gTermId),
                                new SqlParameter("@ACTIVITY", "Insert")
                            };

                            i = obj.ExecuteNonQuery("tSalesDtls", DataHelper.SqlCmdType.sqlStoredProcedure, sqlParamGrnDtls, sqlTran);
                        }
                    }

                }

                if (payDet.Count > 0 && payDet != null)
                {
                    foreach (var PayDetailsModel in payDet)
                    {
                        string Receipt = Convert.ToString(obj.ExecuteScalar("SELECT 'REC' + CONVERT(VARCHAR,ISNULL((MAX(CONVERT(INT,(SUBSTRING(ReceiptNo ,(3 + 1) ,len( ReceiptNo ))))) + 1),1)) AS ReceiptNo FROM trnPayDetails", DataHelper.SqlCmdType.sqlText, sqlTran));

                        SqlParameter[] sqlParamDocDtls = new SqlParameter[]
                        {
                            new SqlParameter("@BillNo", BillNo),
                            new SqlParameter("@Module","OPSales"),
                            new SqlParameter("@TransactionType", "Payment"),
                            new SqlParameter("@ReceiptNo",Receipt),
                            new SqlParameter("@PaymentID",PaymentID),
                            new SqlParameter("@Amount", PayDetailsModel.Amount),
                            new SqlParameter("@Authorisation", cmbDiscAuth.SelectedValue),
                            new SqlParameter("@Remarks", txtDiscReason.Text),
                            new SqlParameter("@PayMode", PayDetailsModel.PayMode),
                            new SqlParameter("@WalletAccount", PayDetailsModel.WalletAccount),
                            new SqlParameter("@Bank", PayDetailsModel.Bank),
                            new SqlParameter("@TransNo",PayDetailsModel.TransNo),
                            new SqlParameter("@EmpPayment", false),
                            new SqlParameter("@LocationID", gvars.gLocationId),
                            new SqlParameter("@UserID", gvars.gUserID),
                            new SqlParameter("@CreateUserID", gvars.gUserID),
                            new SqlParameter("@TerminalID", gvars.gTermId),
                            new SqlParameter("@ACTIVITY", "InsertPayDetails"),
                        };
                        i = obj.ExecuteNonQuery("tPayDetails", DataHelper.SqlCmdType.sqlStoredProcedure, sqlParamDocDtls, sqlTran);
                    }
                    if (Convert.ToDecimal(txtDiscount.Text) > 0)
                    {
                        string Receipt = Convert.ToString(obj.ExecuteScalar("SELECT 'REC' + CONVERT(VARCHAR,ISNULL((MAX(CONVERT(INT,(SUBSTRING(ReceiptNo ,(3 + 1) ,len( ReceiptNo ))))) + 1),1)) AS ReceiptNo FROM trnPayDetails", DataHelper.SqlCmdType.sqlText, sqlTran));

                        SqlParameter[] sqlParamDocDtls = new SqlParameter[]
                        {
                            new SqlParameter("@BillNo", BillNo),
                            new SqlParameter("@Module","OPSales"),
                            new SqlParameter("@TransactionType", "Discount"),
                            new SqlParameter("@ReceiptNo",Receipt),
                            new SqlParameter("@PaymentID",PaymentID),
                            new SqlParameter("@Amount", Convert.ToDecimal(txtDiscount.Text)),
                            new SqlParameter("@DiscPerc", Convert.ToDecimal(txtConcPer.Text)),
                            new SqlParameter("@Authorisation", cmbDiscAuth.SelectedValue),
                            new SqlParameter("@Remarks",txtDiscReason.Text),
                            new SqlParameter("@PayMode", "Cash"),
                            new SqlParameter("@WalletAccount", "0"),
                            new SqlParameter("@Bank", "0"),
                            new SqlParameter("@TransNo", "0"),
                            new SqlParameter("@EmpPayment", false),
                            new SqlParameter("@LocationID", gvars.gLocationId),
                            new SqlParameter("@UserID", gvars.gUserID),
                            new SqlParameter("@CreateUserID", gvars.gUserID),
                            new SqlParameter("@TerminalID", gvars.gTermId),
                            new SqlParameter("@ACTIVITY", "InsertPayDetails"),
                        };
                        i = obj.ExecuteNonQuery("tPayDetails", DataHelper.SqlCmdType.sqlStoredProcedure, sqlParamDocDtls, sqlTran);

                        #region Discount
                        string DiscID = Convert.ToString(obj.ExecuteScalar("SELECT 'DSC' + CONVERT(VARCHAR,ISNULL((MAX(CONVERT(INT,(SUBSTRING(DiscID,(3 + 1),LEN(DiscID))))) + 1),1)) AS DiscID FROM trnDiscountDtls", DataHelper.SqlCmdType.sqlText, sqlTran));

                        if (aItem != null)
                        {
                            foreach (var Dtls in aItem)
                            {
                                SqlParameter[] sqlParamDisc = new SqlParameter[]
                                   {
                                            new SqlParameter("@DiscID", DiscID),
                                            new SqlParameter("@Module", "OPSales"),
                                            new SqlParameter("@TransactionType", "Discount"),
                                            new SqlParameter("@BillNo", BillNo),
                                            new SqlParameter("@ServiceID", Dtls.ItemId),
                                            new SqlParameter("@DiscPerc", Dtls.DiscPer),
                                            new SqlParameter("@Amount", Dtls.Discount),
                                            new SqlParameter("@Authorisation", cmbDiscAuth.SelectedValue),
                                            new SqlParameter("@Remarks", txtDiscReason.Text),
                                            new SqlParameter("@PaymentID", PaymentID),
                                            new SqlParameter("@LocationID", gvars.gLocationId),
                                            new SqlParameter("@UserID", gvars.gUserID),
                                            new SqlParameter("@TerminalID", gvars.gTermId),
                                            new SqlParameter("@ACTIVITY", "InsertPharmacyDiscount")
                                   };

                                i = obj.ExecuteNonQuery("tPayDetails", DataHelper.SqlCmdType.sqlStoredProcedure, sqlParamDisc, sqlTran);
                            }
                        }

                        #endregion
                    }
                }

                if (aItem != null)
                {
                    foreach (var Dtls in aItem)
                    {
                        SqlParameter[] sqlParamGrnDtls = new SqlParameter[]
                        {
                            new SqlParameter("@DepartmentID",gvars.gDeptID),
                            new SqlParameter("@ItemID", Dtls.ItemId),
                            new SqlParameter("@Qty", Dtls.Qty),
                            new SqlParameter("@UnitRate", Dtls.UnitRate),
                            new SqlParameter("@UnitMrp", Dtls.UnitMrp),
                            new SqlParameter("@BatchNo", Dtls.BatchNo),
                            new SqlParameter("@ExpiryDate", Dtls.ExpiryDate),
                            new SqlParameter("@TaxPer", Dtls.TaxPer),
                            new SqlParameter("@SupplierID", Dtls.SupplierID),
                            new SqlParameter("@PurchBillNo", Dtls.PurchBillNo),
                            new SqlParameter("@LocationID", gvars.gLocationId),
                            new SqlParameter("@EditUserID",gvars.gUserID),
                            new SqlParameter("@EditTerminalID", gvars.gTermId),
                            new SqlParameter("@ACTIVITY", "UpdateStock")
                        };

                        i = obj.ExecuteNonQuery("tStock", DataHelper.SqlCmdType.sqlStoredProcedure, sqlParamGrnDtls, sqlTran);
                    }
                }

                sqlTran.Commit();

                if (i > 0)
                {
                    MessageBox.Show("Sales Details created Successfully", "Information", MessageBoxButton.OK, MessageBoxImage.Information);
                    PrintSlip(BillNo);
                }
                else
                {
                    MessageBox.Show("Sales Details Not Created", "Information", MessageBoxButton.OK, MessageBoxImage.Information);
                }

                ClearControls();
                GridPanel.Visibility = Visibility.Visible;
                FormPanel.Visibility = Visibility.Collapsed;
                btnNew.Content = "New";
            }
            catch (Exception ex)
            {
                sqlTran.Rollback();
                if (con.State == ConnectionState.Open)
                    con.Close();
            }
        }
        
        #endregion

        #region Reset
        private void btnReset_Click(object sender, RoutedEventArgs e)
        {
            EnableControls();
            ClearControls();
        }
        #endregion

        #region PayResetDetails
        private void PayResetDetails()
        {
            UpdatTotals();
            //txtDueAmt.Text = txtTotalAmt.Text =  txtPayAmount.Text = Convert.ToDecimal(0).ToString();
            txtDiscReason.Text = txtRefNo.Text = txtDueReason.Text = "";
            cmbWalletType.SelectedIndex = cmbDueAuth.SelectedIndex = cmbBank.SelectedIndex = -1;
            cmbDiscType.SelectedIndex = cmbPayMode.SelectedIndex = 0;
            payDet.Clear();
            dgvPaymentDet.ItemsSource = null;         
            txtConcPer.IsEnabled = true;
            txtConcAmt.IsEnabled = true;
            txtPaidAmt.Text = txtConcPer.Text = txtConcAmt.Text  = txtDiscount.Text = "0";
            lblCardNo.Visibility = Visibility.Collapsed;
            cmbBank.Visibility = Visibility.Collapsed;
            txtRefNo.Visibility = Visibility.Collapsed;
            lblBank.Visibility = Visibility.Collapsed;
            lblCheqDt.Visibility = Visibility.Collapsed;
            dtpCheqDt.Visibility = Visibility.Collapsed;
            lblCheqNo.Visibility = Visibility.Collapsed;
            lblReferenceNo.Visibility = Visibility.Collapsed;
            lblWallet.Visibility = Visibility.Collapsed;
            cmbWalletType.Visibility = Visibility.Collapsed;
            UpdateDue();
            decimal NetAmount = 0;
            if (dgvItemDetails.Items.Count > 0)
            {
                foreach (var item in aItem)
                {
                    item.DiscPer = Convert.ToDecimal(txtConcPer.Text);
                    item.Discount = Math.Round((item.Amount * Convert.ToDecimal(txtConcPer.Text)) / 100, 2);
                }
            }
        }
        #endregion

        #region Pay Reset
        private void btnPayReset_Click(object sender, RoutedEventArgs e)
        {
            PayResetDetails();
        }

        #endregion

        #region Delete Item
        private void dgvItemDetails_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            //if (e.Key == Key.Delete)
            //{
            //    var SelectRow = dgvItemDetails.SelectedItem as AddItems;
            //    if (SelectRow != null)
            //    {
            //        var viewModel = DataContext as AddItems;
            //        if (MessageBox.Show("Are you sure you want to delete ?", "Confirm Delete",
            //        MessageBoxButton.YesNo, MessageBoxImage.Warning) == MessageBoxResult.Yes)
            //        {
            //            //viewModel?.AddItems.Remove(SelectRow);
            //        }
            //    }
            //}
        }

        #endregion

        #region cmbDiscType_SelectionChanged
        private void cmbDiscType_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (cmbDiscType.SelectedIndex == 1)
            {
                txtConcAmt.Text = Convert.ToDecimal(0).ToString();
                txtConcPer.Text = Convert.ToDecimal(0).ToString();
                txtConcPer.IsEnabled = true;
                cmbDiscAuth.IsEnabled = true;
                txtDiscReason.IsEnabled = true;
                txtConcAmt.IsEnabled = false;
                txtConcPer.Focus();
            }
            else if (cmbDiscType.SelectedIndex == 2)
            {
                txtConcAmt.Text = Convert.ToDecimal(0).ToString();
                txtConcPer.Text = Convert.ToDecimal(0).ToString();
                txtConcPer.IsEnabled = false;
                txtConcAmt.IsEnabled = true;
                txtConcAmt.Focus();
                cmbDiscAuth.IsEnabled = true;
                txtDiscReason.IsEnabled = true;
            }
            else
            {
                cmbDiscAuth.IsEnabled = false;
                txtDiscReason.IsEnabled = false;
            }
        }
        #endregion

        #region cmbPayMode_DropDownClosed
        private void cmbPayMode_DropDownClosed(object sender, EventArgs e)
        {

            if (cmbPayMode.SelectedIndex > -1)
            {
                if (cmbPayMode.Text == "Card")
                {                     
                    lblCheqDt.Visibility = Visibility.Collapsed;
                    dtpCheqDt.Visibility = Visibility.Collapsed;
                    lblCheqNo.Visibility = Visibility.Collapsed;
                    lblReferenceNo.Visibility = Visibility.Collapsed;
                    lblWallet.Visibility = Visibility.Collapsed;
                    cmbWalletType.Visibility = Visibility.Collapsed;
                    lblCardNo.Visibility = Visibility.Visible;
                    cmbBank.Visibility = Visibility.Visible;
                    txtRefNo.Visibility = Visibility.Visible;
                    lblBank.Visibility = Visibility.Visible;
                }
                else if (cmbPayMode.Text == "Cheque")
                {
                    lblCardNo.Visibility = Visibility.Collapsed;
                    lblReferenceNo.Visibility = Visibility.Collapsed;
                    lblWallet.Visibility = Visibility.Collapsed;
                    cmbWalletType.Visibility = Visibility.Collapsed;
                    cmbBank.Visibility = Visibility.Visible;
                    txtRefNo.Visibility = Visibility.Visible;
                    lblBank.Visibility = Visibility.Visible;
                    lblCheqDt.Visibility = Visibility.Visible;
                    dtpCheqDt.Visibility = Visibility.Visible;
                    lblCheqNo.Visibility = Visibility.Visible;

                }
                else if (cmbPayMode.Text == "UPI")
                {
                    lblCardNo.Visibility = Visibility.Collapsed;
                    cmbBank.Visibility = Visibility.Collapsed;
                    lblBank.Visibility = Visibility.Collapsed;
                    lblCheqDt.Visibility = Visibility.Collapsed;
                    dtpCheqDt.Visibility = Visibility.Collapsed;
                    lblCheqNo.Visibility = Visibility.Collapsed;
                    lblReferenceNo.Visibility = Visibility.Visible;
                    lblWallet.Visibility = Visibility.Visible;
                    cmbWalletType.Visibility = Visibility.Visible;
                    txtRefNo.Visibility = Visibility.Visible;
                }
                else
                {
                    lblCardNo.Visibility = Visibility.Collapsed;
                    cmbBank.Visibility = Visibility.Collapsed;
                    txtRefNo.Visibility = Visibility.Collapsed;
                    lblBank.Visibility = Visibility.Collapsed;
                    lblCheqDt.Visibility = Visibility.Collapsed;
                    dtpCheqDt.Visibility = Visibility.Collapsed;
                    lblCheqNo.Visibility = Visibility.Collapsed;
                    lblReferenceNo.Visibility = Visibility.Collapsed;
                    lblWallet.Visibility = Visibility.Collapsed;
                    cmbWalletType.Visibility = Visibility.Collapsed;
                }
            }
        }
        #endregion

        #region Concession
        private void txtConcPer_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (txtBillno.Text == "" )
            {
                CalculateConc();
            }
        }

        private void txtConcAmt_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (txtBillno.Text == "" )
            {
                CalculateConc();
            }
        }

        private void CalculateConc()
        {

            if (txtDueAmt.Text != "" && txtDueAmt.Text != "0" && txtDueAmt.Text != Convert.ToDecimal(0).ToString())
            {
                if ((txtConcPer.Text != "" && txtConcPer.Text != "0")) // For PErcentage
                {
                    if (cmbDiscType.SelectedIndex == 1)
                    {
                        if (Convert.ToDecimal(txtConcPer.Text) > 100)
                        {
                            MessageBox.Show("You Are Not Allowed To Give Discount More than " + 100 + " % Discount", "Infomation", MessageBoxButton.OK, MessageBoxImage.Information);
                            txtConcPer.Text = "";
                            txtConcPer.Focus();
                            return;
                        }
                    }

                    if (cmbDiscType.SelectedIndex == 1)
                    {
                        txtConcAmt.Text = txtDiscount.Text = (Convert.ToDecimal(txtTotalAmt.Text) * Convert.ToDecimal(txtConcPer.Text) / 100).ToString();
                        if (gvars.gConcType == "%")
                        {
                            if (Convert.ToDecimal(txtConcPer.Text) > gvars.gConcLimit)
                            {
                                MessageBox.Show("You Are Not Allowed To Give Discount More than " + gvars.gConcLimit + " % Discount", "Infomation", MessageBoxButton.OK, MessageBoxImage.Information);
                                txtConcPer.Text = txtConcAmt.Text = Convert.ToDecimal(0).ToString();
                                txtConcPer.Focus();
                                return;
                            }
                        }
                        else
                        {
                            if (Convert.ToDecimal(txtConcAmt.Text) > gvars.gConcLimit)
                            {
                                MessageBox.Show("You Are Not Allowed To Give Discount More than " + gvars.gConcLimit + " Rs Discount", "Infomation", MessageBoxButton.OK, MessageBoxImage.Information);
                                txtConcPer.Text = txtConcAmt.Text = Convert.ToDecimal(0).ToString();
                                txtConcPer.Focus();
                                return;
                            }
                        }

                        txtDueAmt.Text = Math.Round((Convert.ToDecimal(txtTotalAmt.Text) - Convert.ToDecimal(txtConcAmt.Text)), 2).ToString();
                        txtPayAmount.Text = txtDueAmt.Text;
                    }

                    else if (cmbDiscType.SelectedIndex == 2)
                    {
                        if (txtConcAmt.Text != "")
                        {
                            txtConcPer.Text = Math.Round((Convert.ToDecimal(txtConcAmt.Text) / Convert.ToDecimal(txtTotalAmt.Text) * 100), 2).ToString();
                            if (gvars.gConcType == "%")
                            {
                                if (Convert.ToDecimal(txtConcPer.Text) > gvars.gConcLimit)
                                {
                                    MessageBox.Show("You Are Not Allowed To Give Discount More than " + gvars.gConcLimit + " % Discount", "Infomation", MessageBoxButton.OK, MessageBoxImage.Information);
                                    txtConcPer.Text = txtConcAmt.Text = Convert.ToDecimal(0).ToString();
                                    txtConcAmt.Focus();
                                    return;
                                }
                            }
                            else
                            {
                                if (Convert.ToDecimal(txtConcAmt.Text) > gvars.gConcLimit)
                                {
                                    MessageBox.Show("You Are Not Allowed To Give Discount More than " + gvars.gConcLimit + " Rs Discount", "Infomation", MessageBoxButton.OK, MessageBoxImage.Information);
                                    txtConcPer.Text = txtConcAmt.Text = Convert.ToDecimal(0).ToString();
                                    txtConcAmt.Focus();
                                    return;
                                }
                            }

                            txtDiscount.Text = txtConcAmt.Text;
                            txtDueAmt.Text = (Convert.ToDecimal(txtTotalAmt.Text) - Convert.ToDecimal(txtConcAmt.Text)).ToString();
                            txtPayAmount.Text = txtDueAmt.Text;
                        }
                    }

                    decimal NetAmount = 0;
                    string TaxCalculation = Convert.ToString(obj.ExecuteScalar("select TaxCalculation from mstDepartmentDtls where DepartmentID='" + gvars.gDeptID + "'", DataHelper.SqlCmdType.sqlText));
                    if (TaxCalculation == "BeforeDiscount")
                    {
                        if (dgvItemDetails.Items.Count > 0)
                        {
                            foreach (var item in aItem)
                            {
                                NetAmount = item.Amount;
                                item.DiscPer = Convert.ToDecimal(txtConcPer.Text);
                                item.Discount = Math.Round((item.Amount * Convert.ToDecimal(txtConcPer.Text)) / 100, 2);
                            }
                        }
                    }
                    else
                    {
                        if (dgvItemDetails.Items.Count > 0)
                        {
                            foreach (var item in aItem)
                            {
                                item.DiscPer = Convert.ToDecimal(txtConcPer.Text);
                                item.Discount = (item.Amount * Convert.ToDecimal(txtConcPer.Text)) / 100;
                                NetAmount = item.Amount - item.Discount;
                                item.TaxAmount = Math.Round((NetAmount * item.TaxPer) / 100, 2);
                            }
                        }
                    }
                }
                else
                {
                    if ((txtConcAmt.Text != "" && txtConcAmt.Text != "0"))
                    {
                        if (cmbDiscType.SelectedIndex == 2)
                        {

                            txtConcPer.Text = Math.Round((Convert.ToDecimal(txtConcAmt.Text) / Convert.ToDecimal(txtTotalAmt.Text) * 100), 2).ToString();
                            if (gvars.gConcType == "%")
                            {
                                if (Convert.ToDecimal(txtConcPer.Text) > gvars.gConcLimit)
                                {
                                    MessageBox.Show("You Are Not Allowed To Give Discount More than " + gvars.gConcLimit + " % Discount", "Infomation", MessageBoxButton.OK, MessageBoxImage.Information);
                                    txtConcPer.Text = txtConcAmt.Text = "";
                                    txtConcAmt.Focus();
                                    return;
                                }
                            }
                            else
                            {
                                if (Convert.ToDecimal(txtConcAmt.Text) > gvars.gConcLimit)
                                {
                                    MessageBox.Show("You Are Not Allowed To Give Discount More than " + gvars.gConcLimit + " Rs Discount", "Infomation", MessageBoxButton.OK, MessageBoxImage.Information);
                                    txtConcPer.Text = txtConcAmt.Text = "";
                                    txtConcAmt.Focus();
                                    return;
                                }
                            }
                            txtDiscount.Text = txtConcAmt.Text;
                            txtDueAmt.Text = (Convert.ToDecimal(txtTotalAmt.Text) - Convert.ToDecimal(txtConcAmt.Text)).ToString();
                            txtPayAmount.Text = txtDueAmt.Text;
                        }
                        else
                        {
                            txtConcPer.Text = Convert.ToDecimal(0).ToString();
                            txtConcAmt.Text = Convert.ToDecimal(0).ToString();
                            txtDiscount.Text = txtConcAmt.Text;
                            txtDueAmt.Text = (Convert.ToDecimal(txtTotalAmt.Text) - Convert.ToDecimal(txtConcAmt.Text)).ToString();
                            txtPayAmount.Text = txtDueAmt.Text;
                        }

                        decimal NetAmount = 0;
                        string TaxCalculation = Convert.ToString(obj.ExecuteScalar("select TaxCalculation from mstDepartmentDtls where DepartmentID='" + gvars.gDeptID + "'", DataHelper.SqlCmdType.sqlText));
                        if (TaxCalculation == "BeforeDiscount")
                        {
                            if (dgvItemDetails.Items.Count > 0)
                            {
                                foreach (var item in aItem)
                                {
                                    NetAmount = item.Amount;
                                    item.DiscPer = Convert.ToDecimal(txtConcPer.Text);
                                    item.Discount = Math.Round((item.Amount * Convert.ToDecimal(txtConcPer.Text)) / 100, 2);
                                }
                            }
                        }
                        else
                        {
                            if (dgvItemDetails.Items.Count > 0)
                            {
                                foreach (var item in aItem)
                                {
                                    item.DiscPer = Convert.ToDecimal(txtConcPer.Text);
                                    item.Discount = (item.Amount * Convert.ToDecimal(txtConcPer.Text)) / 100;
                                    NetAmount = item.Amount - item.Discount;
                                    item.TaxAmount = Math.Round((NetAmount * item.TaxPer) / 100, 2);
                                }
                            }
                        }
                    }
                }


            }
            else
            {
                MessageBox.Show("There is no Due to give Discount", "Infomation", MessageBoxButton.OK, MessageBoxImage.Information);

            }
        }           
        
        #endregion

        #region txtDueAmt_TextChanged
        private void txtDueAmt_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (txtBillno.Text == "" )
            {
                if (txtDueAmt.Text != "" || txtDueAmt.Text != "0")
                {
                    cmbDueAuth.IsEnabled = true;
                    txtDueReason.IsEnabled = true;
                }
                else
                {
                    cmbDueAuth.IsEnabled = false;
                    txtDueReason.IsEnabled = false;
                }
            }
        }
        #endregion

        #region Get Batch
        public void GetBatch(string ItemId)
        { 
            DataHelper objDataHelper1 = new DataHelper();
            SqlParameter[] sqlParamSearch = new SqlParameter[]
            {
               new SqlParameter("@ACTIVITY", "Getbatch"),
               new SqlParameter("@ItemID", ItemId),
               new SqlParameter("@LocationID", gvars.gLocationId),
               new SqlParameter("@DepartmentID", gvars.gDeptID),
            };

            ds = obj.getDataset("tSales", DataHelper.SqlCmdType.sqlStoredProcedure, sqlParamSearch);
            if (ds.Tables[0].Rows.Count > 0 && ds.Tables.Count > 0)
            {
                dsBatch = ds;
                if (Convert.ToInt32(ds.Tables[0].Rows[0]["SalesShowExpMsg"]) > Convert.ToInt32(ds.Tables[0].Rows[0]["NoOfDaysMsg"]))
                {
                    string itemName = txtItem.Tag.ToString();
                    string batchNo = ds.Tables[0].Rows[0]["ExpiryDate"].ToString();
                    DateTime expDate = Convert.ToDateTime(ds.Tables[0].Rows[0]["ExpiryDate"]);
                    string monthYear = expDate.ToString("MMM-yyyy");
                    int daysLeft = Convert.ToInt32(ds.Tables[0].Rows[0]["NoOfDaysMsg"]);

                    string msg = $"{itemName} with Batch No. {batchNo} is going to expire on {monthYear}, ({daysLeft} days left).";
                    MessageBox.Show(msg, "Expiry Alert", MessageBoxButton.OK, MessageBoxImage.Warning);

                    if (Convert.ToInt32(ds.Tables[0].Rows[0]["SaleStopExpMsg"]) > Convert.ToInt32(ds.Tables[0].Rows[0]["NoOfDaysMsg"]))
                    {
                        cmbBatchNo.SelectedIndex = 0;
                    }
                    else
                    {
                        cmbBatchNo.SelectedValuePath = "BatchNo";
                        cmbBatchNo.DisplayMemberPath = "BatchNo";
                        cmbBatchNo.ItemsSource = ds.Tables[0].DefaultView;
                        cmbBatchNo.SelectedIndex = 0;
                        txtBQty.Text = ds.Tables[0].Rows[0]["Qty"].ToString();
                        cmbBatchNo.Text = ds.Tables[0].Rows[0]["BatchNo"].ToString();
                        dtpExpDt.Text = Convert.ToDateTime(ds.Tables[0].Rows[0]["ExpiryDate"]).ToString("dd-MMM-yyyy");
                        txtMRP.Text = ds.Tables[0].Rows[0]["UnitMrp"].ToString();
                        Rate = Convert.ToDecimal(ds.Tables[0].Rows[0]["UnitRate"]);
                        taxPer = Convert.ToDecimal(ds.Tables[0].Rows[0]["TaxPer"]);
                        PBillNo = ds.Tables[0].Rows[0]["PurchBillNo"].ToString();
                        SupId = ds.Tables[0].Rows[0]["SupplierID"].ToString();
                        txtTQty.Text = ds.Tables[0].Compute("Sum(Qty)", "").ToString();
                        isLook = Convert.ToBoolean(ds.Tables[0].Rows[0]["IsLook"]);
                        isHigh = Convert.ToBoolean(ds.Tables[0].Rows[0]["IsHigh"]);
                        isSound = Convert.ToBoolean(ds.Tables[0].Rows[0]["IsSound"]);
                    }

                }
                else
                {
                    cmbBatchNo.SelectedValuePath = "BatchNo";
                    cmbBatchNo.DisplayMemberPath = "BatchNo";
                    cmbBatchNo.ItemsSource = ds.Tables[0].DefaultView;
                    cmbBatchNo.SelectedIndex = 0;
                    txtBQty.Text = ds.Tables[0].Rows[0]["Qty"].ToString();
                    cmbBatchNo.Text = ds.Tables[0].Rows[0]["BatchNo"].ToString();
                    dtpExpDt.Text = Convert.ToDateTime(ds.Tables[0].Rows[0]["ExpiryDate"]).ToString("dd-MMM-yyyy");
                    txtMRP.Text = ds.Tables[0].Rows[0]["UnitMrp"].ToString();
                    Rate = Convert.ToDecimal(ds.Tables[0].Rows[0]["UnitRate"]);
                    taxPer = Convert.ToDecimal(ds.Tables[0].Rows[0]["TaxPer"]);
                    PBillNo = ds.Tables[0].Rows[0]["PurchBillNo"].ToString();
                    SupId = ds.Tables[0].Rows[0]["SupplierID"].ToString();
                    txtTQty.Text = ds.Tables[0].Compute("Sum(Qty)", "").ToString();
                    isLook = Convert.ToBoolean(ds.Tables[0].Rows[0]["IsLook"]);
                    isHigh = Convert.ToBoolean(ds.Tables[0].Rows[0]["IsHigh"]);
                    isSound = Convert.ToBoolean(ds.Tables[0].Rows[0]["IsSound"]);
                }
            }
            else
            {
                MessageBox.Show("No Stock available", "NoStock", MessageBoxButton.OK, MessageBoxImage.Information); 
                txtItem.Focus();
                ClearItems();
            }
             
        }

        #endregion

        #region cmbSaleType_SelectionChanged
        private void cmbSaleType_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            //if (cmbSaleType.SelectedIndex == 1)
            //{
            if (rdoOP.IsChecked == true)
            {
                ClearControls();
                txtName.IsEnabled = false;
                txtAge.IsEnabled = false;
                cmbDoctor.IsEnabled = false;
                txtPhone.IsEnabled = false;
            }
            //else if (cmbSaleType.SelectedIndex == 2)
            else if (rdoOthers.IsChecked == true)
            {
                ClearControls();
                txtName.IsEnabled = true;
                txtAge.IsEnabled = true;
                cmbDoctor.IsEnabled = true;
                txtPhone.IsEnabled = true;
            }
        }
        #endregion

        #region txtPayAmount_TextChanged
        private void txtPayAmount_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (txtBillno.Text == "" )
            {
                if (txtDueAmt.Text != "" && txtPayAmount.Text != "")
                {
                    if (Convert.ToDecimal(txtPayAmount.Text) > Convert.ToDecimal(txtDueAmt.Text))
                    {
                        MessageBox.Show("Pay Amount should not be greater Than Due Amount!", "Information", MessageBoxButton.OK, MessageBoxImage.Information);
                        txtPayAmount.Text = "";
                        txtPayAmount.Focus();
                        return;
                    }
                }
            }
        }
        #endregion
        

        #region cmbBatchNo_SelectionChanged
        private void cmbBatchNo_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {           
            if (cmbBatchNo.SelectedIndex > -1 && dsBatch != null)
            {
                DataRow[] drow = dsBatch.Tables[0].Select("BatchNo = '" + cmbBatchNo.SelectedValue + "'");

                if (drow.Length > 0)
                {               
                    cmbBatchNo.Text = drow[0].ItemArray[4].ToString();
                    dtpExpDt.Text = Convert.ToDateTime(drow[0].ItemArray[5]).ToString("MMM-yyyy");
                    txtMRP.Text = drow[0].ItemArray[13].ToString();
                    Rate = Convert.ToDecimal(drow[0].ItemArray[14]);
                    taxPer = Convert.ToDecimal(drow[0].ItemArray[6]);
                    PBillNo = drow[0].ItemArray[10].ToString();
                    SupId = drow[0].ItemArray[12].ToString();
                    txtBQty.Text = drow[0].ItemArray[15].ToString();
                    txtTQty.Text = dsBatch.Tables[0].Compute("Sum(Qty)", "").ToString();
                }
            }                
        }
        #endregion

        #region Qty_TextChanged
        private void Qty_TextChanged(object sender, TextChangedEventArgs e)
        {
            var tb = sender as TextBox;
            if (tb == null) return;

            var item = tb.DataContext as AddItems;
            if (item == null) return;


            item.IsQtyHighlighted = false;
        }

        #endregion

        #region txtQty_TextChanged
        private void txtQty_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (txtQty.Text != "")
            {
                if (Convert.ToDecimal(txtQty.Text) > 0)
                {
                    if (Convert.ToDecimal(txtQty.Text) > Convert.ToDecimal(txtTQty.Text))
                    {
                        MessageBox.Show(" You Cannot Issue More than : " + txtTQty.Text, "Information", MessageBoxButton.OK, MessageBoxImage.Information);
                        txtQty.Text = "";
                        txtQty.Focus();
                        return;
                    }

                    txtAmount.Text = (Math.Round(Convert.ToDecimal(txtQty.Text) * Convert.ToDecimal(txtMRP.Text), 2)).ToString();
                    txtTaxAmt.Text = (Math.Round((Convert.ToDecimal(txtQty.Text) * Convert.ToDecimal(txtMRP.Text)) * (taxPer / 100), 2)).ToString();
                }
            }
        }

        #endregion

        #region UpdatTotals
        private void UpdatTotals()
        {
            if (txtBillno.Text == "" )
            { 
                if (dgvItemDetails.Items.Count > 0)
                {
                    decimal Tot = 0, Disc = 0;
                    foreach (var item in aItem)
                    {
                        Tot += item.Amount;
                        Disc += item.Discount;
                    }

                    txtPayAmount.Text = "0";
                    txtTotalAmt.Text = txtDueAmt.Text = Tot.ToString();
                    txtDiscount.Text = Disc.ToString();
                }
            }
        }
        
        private void UpdateDue()
        {
            if (txtBillno.Text == "" )
            {
                if (dgvPaymentDet.Items.Count > 0)
                {
                    decimal paid = 0;
                    foreach (var item in payDet)
                    {
                        paid += item.Amount;
                    }
                    txtPaidAmt.Text = paid.ToString();
                    txtPayAmount.Text = "0";
                    txtDueAmt.Text = (Convert.ToDecimal(txtDueAmt.Text) - Convert.ToDecimal(txtPaidAmt.Text)).ToString();
                }
            }
        }

        private void btnBackUHID_Click(object sender, RoutedEventArgs e)
        {
            UHIDPopUp.IsOpen = false;
        }

        #endregion

        #region ClearControls
        private void ClearControls()
        {
            txtUHIDPhone.Text = txtName.Text = txtAge.Text = txtPhone.Text=txtUHID.Text = txtOPDNO.Text=txtBillno.Text = "";
            txtTotalAmt.Text = txtDiscount.Text = txtDueAmt.Text = txtDueReason.Text = txtConcAmt.Text = txtConcPer.Text = txtDiscReason.Text = txtTrnsCode.Text = txtPayAmount.Text = ""; //Convert.ToDecimal(0).ToString();
            cmbDueAuth.SelectedIndex = 0;
            cmbDiscAuth.SelectedIndex = 0;
            cmbDiscType.SelectedIndex = 0;
            cmbDoctor.SelectedIndex = 0;
            //cmbSaleType.SelectedIndex = 0;
            //rdoOP.IsChecked = true;
            dgvItemDetails.ItemsSource=null;
            aItem.Clear();
            payDet.Clear();
            dgvPaymentDet.ItemsSource=null;
            popup.IsOpen = false;
        }
         
        #endregion

        #region btnUHIDView_Click
        private void btnUHIDView_Click(object sender, RoutedEventArgs e)
        {
              UHIDPopUp.IsOpen = false;

            if (dgvPatientDetails.SelectedItems.Count > 0)
            {
                var button = sender as Button;
                var rowData = button?.DataContext;
                    DataRowView dr = (DataRowView)rowData;
                
                if (dr != null)
                {
                    txtName.Text = dr["Name"].ToString();
                    txtAge.Text = dr["Age"].ToString();
                    txtUHID.Text = dr["UHID"].ToString();
                    txtOPDNO.Text = dr["OPDNO"].ToString();
                    txtPhone.Text = dr["Phone"].ToString();
                    cmbDoctor.SelectedValue= dr["DocId"].ToString();
                    UHIDPopUp.Visibility = Visibility.Collapsed;
                    if(dr["OrganisationName"].ToString() != "Cash")
                    {
                        OrganisationID = dr["OrganisationID"].ToString();
                    }
                    else
                    {
                        OrganisationID = "";
                    }
                }
            }
        }

        private void rdoOP_Checked(object sender, RoutedEventArgs e)
        {
            ClearControls();
            txtName.IsEnabled = false;
            txtAge.IsEnabled = false;
            cmbDoctor.IsEnabled = false;
            txtPhone.IsEnabled = false;
        }

        private void rdoOthers_Checked(object sender, RoutedEventArgs e)
        {
            ClearControls();
            txtName.IsEnabled = true;
            txtAge.IsEnabled = true;
            cmbDoctor.IsEnabled = true;
            txtPhone.IsEnabled = true;
            txtName.Focus();
        }

        #endregion

        #region txtItem_LostFocus
        private void txtItem_LostFocus(object sender, RoutedEventArgs e)
        {
            Dispatcher.BeginInvoke(new Action(() =>
            {
                if (!txtItem.IsKeyboardFocusWithin && !lstItems.IsKeyboardFocusWithin)
                {
                    popup.IsOpen = false;
                }
            }), DispatcherPriority.Background);
        }

        private void lstItems_LostFocus(object sender, RoutedEventArgs e)
        {
            Dispatcher.BeginInvoke(new Action(() =>
            {
                if (!txtItem.IsKeyboardFocusWithin && !lstItems.IsKeyboardFocusWithin)
                {
                    popup.IsOpen = false;
                }
            }), DispatcherPriority.Background);
        }
        #endregion

        #region btnUHIDGo_Click
        private void btnUHIDGo_Click(object sender, RoutedEventArgs e)
        {
            string SaleType = string.Empty;
            if (rdoOthers.IsChecked == true)
            {
                SaleType = "OP";
            }
            else
            {
                SaleType = "others";
            }

            SqlParameter[] sqlParamSearch = new SqlParameter[]
            {
                new SqlParameter("@ACTIVITY", "PatientDetailes"),
                new SqlParameter("@Type", SaleType),
                new SqlParameter("@FromDate", dtpUHIDFrom.Text),
                new SqlParameter("@ToDate", dtpUHIDTo.Text),
                new SqlParameter("@LocationID", gvars.gLocationId),
            };
            DataSet dsSearch = obj.getDataset("tSales", DataHelper.SqlCmdType.sqlStoredProcedure, sqlParamSearch);
            if (dsSearch.Tables.Count > 0 && dsSearch.Tables[0].Rows.Count > 0)
            {
                if (!dsSearch.Tables[0].Columns.Contains("Slno"))
                {
                    dsSearch.Tables[0].Columns.Add("Slno", typeof(int));
                }
                for (int i = 0; i < dsSearch.Tables[0].Rows.Count; i++)
                {
                    dsSearch.Tables[0].Rows[i]["Slno"] = i + 1;
                }
                dgvPatientDetails.ItemsSource = dsSearch.Tables[0].DefaultView;
            }
            else
            {
                MessageBox.Show("No Data Found", "Information", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }
         
        #endregion

        #region QTY
        private void Qty_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {

            e.Handled = !e.Text.All(char.IsDigit);
        }

        private void Qty_LostFocus(object sender, RoutedEventArgs e)
        {
            TextBox tb = sender as TextBox;
            if (tb == null) return;

            foreach (var prop in tb.DataContext.GetType().GetProperties())
            {
                var dc = tb.DataContext;
                var qtyProp = dc.GetType().GetProperty("Qty");
                var bqtyProp = dc.GetType().GetProperty("BatchQty");
                var unimrpProp = dc.GetType().GetProperty("UnitMrp");
                var amountProp = dc.GetType().GetProperty("Amount");

                decimal qty = Convert.ToDecimal(qtyProp?.GetValue(dc) ?? 0);
                decimal bqty = Convert.ToDecimal(bqtyProp?.GetValue(dc) ?? 0);
                decimal unimrp = Convert.ToDecimal(unimrpProp?.GetValue(dc) ?? 0);
                if (qty > bqty)
                {
                    MessageBox.Show(
                        $"Please Enter Below Batch Qty  {bqty}",
                        "warning",
                        MessageBoxButton.OK,
                        MessageBoxImage.Error
                    );

                    tb.Text = "0";
                    //  amountProp?.SetValue(dc, 0);
                    if (amountProp != null && dc != null)
                    {
                        Type propType = amountProp.PropertyType;

                        if (propType == typeof(decimal))
                            amountProp.SetValue(dc, 0m);
                        else if (propType == typeof(double))
                            amountProp.SetValue(dc, 0.0);
                        else if (propType == typeof(int))
                            amountProp.SetValue(dc, 0);
                        else
                            amountProp.SetValue(dc, null);
                    }

                    txtTotalAmt.Text = 0.00.ToString("0.00");
                    txtDueAmt.Text = 0.00.ToString("0.00");
                    payDet.Clear();
                    UpdatTotals();
                    CalculateConc();
                    dgvPaymentDet.ItemsSource = null;
                    return;
                }
                else
                {
                    decimal changeQty = qty * unimrp;
                    amountProp?.SetValue(dc, changeQty);

                    decimal totalAmount = 0m;
                    decimal amt = 0m;

                    foreach (var row in dgvItemDetails.Items)
                    {
                        if (row == null || row == CollectionView.NewItemPlaceholder)
                            continue;


                        var amountPropertyInfoLocal = row.GetType().GetProperty("Amount");
                        if (amountPropertyInfoLocal != null)
                        {
                            var val = amountPropertyInfoLocal.GetValue(row);
                            if (val != null && decimal.TryParse(val.ToString(), out amt))
                                totalAmount += amt;
                            continue;
                        }

                        System.Diagnostics.Debug.WriteLine($"No Amount property for type: {row.GetType().FullName}");

                    }

                    txtTotalAmt.Text = totalAmount.ToString("N2");
                    txtDueAmt.Text = totalAmount.ToString("N2");
                    UpdatTotals();
                    CalculateConc();
                    payDet.Clear();
                    dgvPaymentDet.ItemsSource = null;
                }
            }
        }

        #endregion
         
        #region UserControl_PreviewKeyDown
        private void UserControl_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                // Do not disturb multiline TextBox
                TextBox txt = Keyboard.FocusedElement as TextBox;
                if (txt != null && txt.AcceptsReturn)
                    return;

                // Do not disturb DataGrid enter key
                if (Keyboard.FocusedElement is DataGrid)
                    return;

                e.Handled = true;

                TraversalRequest request = new TraversalRequest(FocusNavigationDirection.Next);

                UIElement element = Keyboard.FocusedElement as UIElement;
                if (element != null)
                {
                    element.MoveFocus(request);
                }
            } 

        }
        #endregion

        #region Models
        public class AddPayDet
        {
            public int Slno { get; set; }
            public decimal Amount { get; set; }
            public string PayMode { get; set; }
            public string WalletAccount { get; set; }
            public string TransNo { get; set; }
            public string Bank { get; set; }
            public string RefDt { get; set; } 
        }
        public class Item
        {
            public string ItemID { get; set; }
            public string ItemName { get; set; }
            public override string ToString()
            {
                return ItemName;  // 👈 This ensures ComboBox.Text shows the actual item name
            }
        }
        
    }

    #endregion

}
