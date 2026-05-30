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
using System.Diagnostics;
using System.ComponentModel;
using System.Windows.Data;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System.IO;
using System.Windows.Media;
using System.Configuration;
using System.Data.SqlClient;

namespace HISPharmacy.Pharmacy
{
    /// <summary>
    /// Interaction logic for IPSalesReturn.xaml
    /// </summary>
    public partial class IPSalesReturn : UserControl
    {
        string strQry;
        DataHelper obj = new DataHelper();
        public GlobalVariables gvars;
        DataSet ds, dsBatch;
        decimal taxAmt, taxPer, Rate;
        string SupId, PBillNo;
        bool FLoad = false;
        ObservableCollection<AddItems> aItem { get; set; } = new ObservableCollection<AddItems>();
        ObservableCollection<AddPayDet> payDet { get; set; } = new ObservableCollection<AddPayDet>();
        public AutoCompleteCombobox ItemCB { get; set; }


        private List<Item> allItems; // Full item list
        private ICollectionView comboView;
        bool isHigh, isLook, isSound, IsReqPharmacyRoundOff;
        string PharamacyRoundOffType = "";
        public IPSalesReturn()
        {
            InitializeComponent();
        }

        #region Page_Loaded
        private void Page_Loaded(object sender, RoutedEventArgs e)
        {
            dtpUHIDFrom.Text = dtpUHIDTo.Text = dtpTo.Text = dtpFrom.Text = DateTime.Today.ToString(); // dtpBillDt.Text =
            Helper.AttachDecimalInputHandlers(txtQty);
            Helper.AttachDecimalInputHandlers(txtDiscount);
            Helper.AttachDecimalInputHandlers(txtConcPer);
            Helper.AttachDecimalInputHandlers(txtPayAmount);
            cmbSaleType.SelectedIndex = 0;
            cmbDiscType.SelectedIndex = 0;
            //BindItem();
            BindWallet();
            BindPayMode();
            BindDoctor();
            BindDueAuthorisation();
            GetUserDiscAuthorized();
            GetBank();
            allItems = LoadItemsFromDatabase();
            cmbItemName.ItemsSource = allItems;
            cmbItemName.DisplayMemberPath = "ItemName";
            cmbItemName.SelectedValuePath = "ItemID";
            cmbItemName.SelectedIndex = 0;
            dgvItemDetails.ItemsSource = aItem;
            dgvPaymentDet.ItemsSource = payDet;
            lblUHIDPhone.Visibility = Visibility.Visible;
            txtIndentPhone.Visibility = Visibility.Visible;
           
            txtConcAmt.Text = "0";
            txtConcPer.Text = "0";
            GetRoundOffTypeamt();
            GetPaymentRights();
            //btnGo_Click(sender, e);
        }
        #endregion

        #region GO
        private void btnGo_Click(object sender, RoutedEventArgs e)
        {
            SqlParameter[] sqlParamSearch = new SqlParameter[]
           {
            new SqlParameter("@ACTIVITY", "BindData1"),
            new SqlParameter("@FromDate", Convert.ToDateTime(Convert.ToDateTime(dtpFrom.Text).ToString("dd MMM yyyy"))),
            new SqlParameter("@ToDate", Convert.ToDateTime(Convert.ToDateTime(dtpTo.Text).ToString("dd MMM yyyy"))),
            new SqlParameter("@Searchtxt", txtFindBillNo.Text),
            new SqlParameter("@LocationID",  gvars.gLocationId),
            new SqlParameter("@DepartmentID", gvars.gDeptID),

           };
            DataSet dsSearch = obj.getDataset("tSalesRet", DataHelper.SqlCmdType.sqlStoredProcedure, sqlParamSearch);
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
                dgvIPSalesSalesFind.ItemsSource = dsSearch.Tables[0].DefaultView;
            }
            else
            {
                MessageBox.Show("No data found.", "Information", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }
        #endregion


        #region btnNew_Click
        private void btnNew_Click(object sender, RoutedEventArgs e)
        {
            if (btnNew.Content.ToString() == "+Add New")
            {
                GridPanel.Visibility = Visibility.Collapsed;
                scrollFormPanel.Visibility = Visibility.Visible;
                FormPanel.Visibility = Visibility.Visible;

                btnNew.Content = "Back";
                EnableControls();
                ClearControls();
                txtFindBillNo.Focus();

                lblUHIDPhone.Visibility = Visibility.Visible;
                txtUHIDPhone.Visibility = Visibility.Visible;


                lblIndentNo1.Visibility = Visibility.Collapsed;
                txtIndentPhone.Visibility = Visibility.Collapsed;
            }
            else
            {
                GridPanel.Visibility = Visibility.Visible;
                scrollFormPanel.Visibility = Visibility.Visible;
                FormPanel.Visibility = Visibility.Collapsed;

                btnNew.Content = "+Add New";
            }

        }
        #endregion

        #region ClearControls
        private void ClearControls()
        {
            txtTotalAmt.Text = txtDiscount.Text = txtDueAmt.Text = txtDueReason.Text = txtConcAmt.Text = txtConcPer.Text = txtDiscReason.Text = txtTrnsCode.Text = txtPayAmount.Text = "";
            cmbDueAuth.SelectedIndex = 0;
            cmbDiscAuth.SelectedIndex = 0;
            cmbDiscType.SelectedIndex = 0;
            cmbSaleType.SelectedIndex = 0;
            dgvItemDetails.ItemsSource = null;
            aItem.Clear();
            payDet.Clear();
            dgvPaymentDet.ItemsSource = null;
        }
        #endregion
        private void EnableControls()
        {
            cmbSaleType.IsEnabled = true;
            txtUHIDPhone.IsEnabled = true;
            grpPayDet.IsEnabled = true;
            grpPayDet.IsEnabled = true;
            grpItemDet.IsEnabled = true;
            btnSave.IsEnabled = true;
            txtTrnsCode.IsEnabled = true;
            btnReset.IsEnabled = true;
        }
        private void DisableControls()
        {
            cmbSaleType.IsEnabled = false;
            txtUHIDPhone.IsEnabled = false;
            grpPayDet.IsEnabled = false;
            grpPayDet.IsEnabled = false;
            grpItemDet.IsEnabled = false;
            btnSave.IsEnabled = false;
            txtTrnsCode.IsEnabled = false;
            btnReset.IsEnabled = false;
        }

        #region cmbSaleType_SelectionChanged
        private void cmbSaleType_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (cmbSaleType.SelectedIndex == 1)
            {
                lblUHID.Text = "";
                NameLabel.Text = "";
                AgeGenderLabel.Text = "";
                GuardianLabel.Text = "";
                PhoneLabel.Text = "";
                AddressLabel.Text = "";
                DOALabel.Text = "";
                PrimaryDrLabel.Text = "";
                AdmissionLabel.Text = "";
                BedRoomNoLabel.Text = "";
                RoomTypeLabel.Text = "";
                PayerLabel.Text = "";
               
                DocIDLabel.Text = "";
                BedIDLabel.Text = "";
                OrgIDLabel.Text = "";

                cmbBatchNo.Text = "";
                //txtBQty.Text = "";
                txtTQty.Text = "";
                dtpExpDt.Text = "";
                txtQty.Text = "";
                txtMRP.Text = "";
                txtTaxAmt.Text = "";
                txtAmount.Text = "";
                txtDueAmt.Text = txtTotalAmt.Text;
                txtDiscReason.Text = txtPayAmount.Text = txtRefNo.Text = txtDueReason.Text = "";
                cmbWalletType.SelectedIndex = cmbDueAuth.SelectedIndex = cmbBank.SelectedIndex = -1;
                cmbDiscType.SelectedIndex = cmbPayMode.SelectedIndex = 0;
                payDet.Clear();
                dgvPaymentDet.ItemsSource = null;
                txtConcPer.IsEnabled = true;
                txtConcAmt.IsEnabled = true;
                txtPaidAmt.Text = txtConcPer.Text = txtConcAmt.Text = txtDiscount.Text = txtTotalAmt.Text = txtDueAmt.Text = "0";
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
                dgvItemDetails.ItemsSource = null;

                lblIndentNo1.Visibility = Visibility.Collapsed;
                txtIndentPhone.Visibility = Visibility.Collapsed;

                lblUHIDPhone.Visibility = Visibility.Visible;
                txtUHIDPhone.Visibility = Visibility.Visible;
            }
            else if (cmbSaleType.SelectedIndex == 2)
            {
                lblUHID.Text = "";
                NameLabel.Text = "";
                AgeGenderLabel.Text = "";
                GuardianLabel.Text = "";
                PhoneLabel.Text = "";
                AddressLabel.Text = "";
                DOALabel.Text = "";
                PrimaryDrLabel.Text = "";
                AdmissionLabel.Text = "";
                BedRoomNoLabel.Text = "";
                RoomTypeLabel.Text = "";
                PayerLabel.Text = "";
               
                DocIDLabel.Text = "";
                BedIDLabel.Text = "";
                OrgIDLabel.Text = "";

                cmbBatchNo.Text = "";
                //txtBQty.Text = "";
                txtTQty.Text = "";
                dtpExpDt.Text = "";
                txtQty.Text = "";
                txtMRP.Text = "";
                txtTaxAmt.Text = "";
                txtAmount.Text = "";
                txtDueAmt.Text = txtTotalAmt.Text;
                txtDiscReason.Text = txtPayAmount.Text = txtRefNo.Text = txtDueReason.Text = "";
                cmbWalletType.SelectedIndex = cmbDueAuth.SelectedIndex = cmbBank.SelectedIndex = -1;
                cmbDiscType.SelectedIndex = cmbPayMode.SelectedIndex = 0;
                payDet.Clear();
                dgvPaymentDet.ItemsSource = null;
                txtConcPer.IsEnabled = true;
                txtConcAmt.IsEnabled = true;
                txtPaidAmt.Text = txtConcPer.Text = txtConcAmt.Text = txtDiscount.Text = txtTotalAmt.Text = txtDueAmt.Text = "0";
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
                dgvItemDetails.ItemsSource = null;

                lblIndentNo1.Visibility = Visibility.Visible;
                txtIndentPhone.Visibility = Visibility.Visible;

                lblUHIDPhone.Visibility = Visibility.Collapsed;
                txtUHIDPhone.Visibility = Visibility.Collapsed;

            }
            else
            {
                lblUHID.Text = "";
                NameLabel.Text = "";
                AgeGenderLabel.Text = "";
                GuardianLabel.Text = "";
                PhoneLabel.Text = "";
                AddressLabel.Text = "";
                DOALabel.Text = "";
                PrimaryDrLabel.Text = "";
                AdmissionLabel.Text = "";
                BedRoomNoLabel.Text = "";
                RoomTypeLabel.Text = "";
                PayerLabel.Text = "";
               
                DocIDLabel.Text = "";
                BedIDLabel.Text = "";
                OrgIDLabel.Text = "";

                cmbBatchNo.Text = "";
               // txtBQty.Text = "";
                txtTQty.Text = "";
                dtpExpDt.Text = "";
                txtQty.Text = "";
                txtMRP.Text = "";
                txtTaxAmt.Text = "";
                txtAmount.Text = "";
                txtDueAmt.Text = txtTotalAmt.Text;
                txtDiscReason.Text = txtPayAmount.Text = txtRefNo.Text = txtDueReason.Text = "";
                cmbWalletType.SelectedIndex = cmbDueAuth.SelectedIndex = cmbBank.SelectedIndex = -1;
                cmbDiscType.SelectedIndex = cmbPayMode.SelectedIndex = 0;
                payDet.Clear();
                dgvPaymentDet.ItemsSource = null;
                txtConcPer.IsEnabled = true;
                txtConcAmt.IsEnabled = true;
                txtPaidAmt.Text = txtConcPer.Text = txtConcAmt.Text = txtDiscount.Text = txtTotalAmt.Text = txtDueAmt.Text = "0";
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
                dgvItemDetails.ItemsSource = null;

                lblIndentNo1.Visibility = Visibility.Collapsed;
                txtIndentPhone.Visibility = Visibility.Collapsed;

                lblUHIDPhone.Visibility = Visibility.Visible;
                txtUHIDPhone.Visibility = Visibility.Visible;
            }
        }
        #endregion

        #region btnFind_Click
        private void btnFind_Click(object sender, RoutedEventArgs e)
        {
            if (cmbSaleType.SelectedIndex == 0)
            {
                MessageBox.Show("Select SaleType", "Information", MessageBoxButton.OK, MessageBoxImage.Information);
                cmbSaleType.Focus();
                return;
            }
            if (cmbSaleType.SelectedIndex == 1)
            {
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
            else if (cmbSaleType.SelectedIndex == 2)
            {
                if (txtIndentPhone.Text != "")
                {
                    CheckIndentData();
                }
                else
                {
                    UHIDPopUp.Visibility = Visibility.Visible;
                    UHIDPopUp.IsOpen = true;
                    //IndentPopUp.Visibility = Visibility.Visible;
                    //IndentPopUp.IsOpen = true;

                }
            }
        }
        #endregion

        #region Get Diretct Data GetPatientDetails
        private void GetPatientDetails()
        {
            if (txtUHIDPhone.Text != "")
            {
                SqlParameter[] sqlParamSearch = new SqlParameter[]
                  {
                    new SqlParameter("@ACTIVITY", "GetWPFSearchFind"),
                    new SqlParameter("@UHID", txtUHIDPhone.Text),
                     new SqlParameter("@LocationID", gvars.gLocationId),
                  };
                ds = obj.getDataset("tIPSales", DataHelper.SqlCmdType.sqlStoredProcedure, sqlParamSearch);
                if (ds.Tables[0].Rows.Count > 0)
                {
                    lblUHID.Text = ds.Tables[0].Rows[0]["UHID"].ToString() + "/" + ds.Tables[0].Rows[0]["IPDNo"].ToString();
                    NameLabel.Text = ds.Tables[0].Rows[0]["Name"].ToString();
                    AgeGenderLabel.Text = ds.Tables[0].Rows[0]["Age"].ToString() + "/" + ds.Tables[0].Rows[0]["Gender"].ToString();
                    GuardianLabel.Text = ds.Tables[0].Rows[0]["Guardian"].ToString();
                    PhoneLabel.Text = ds.Tables[0].Rows[0]["Phone"].ToString();
                    AddressLabel.Text = ds.Tables[0].Rows[0]["Address"].ToString();
                    DOALabel.Text = Convert.ToDateTime(ds.Tables[0].Rows[0]["AdmissionDate"]).ToString("dd-MMM-yyyy");
                    PrimaryDrLabel.Text = ds.Tables[0].Rows[0]["DocName"].ToString();
                    DocIDLabel.Text = ds.Tables[0].Rows[0]["DocId"].ToString();
                    AdmissionLabel.Text = ds.Tables[0].Rows[0]["AdmissonType"].ToString();
                    BedRoomNoLabel.Text = ds.Tables[0].Rows[0]["BedNo"].ToString();
                    BedIDLabel.Text = ds.Tables[0].Rows[0]["BedID"].ToString();
                    RoomTypeLabel.Text = ds.Tables[0].Rows[0]["RoomType"].ToString();
                    PayerLabel.Text = ds.Tables[0].Rows[0]["PayType"].ToString();
                    OrgIDLabel.Text = ds.Tables[0].Rows[0]["OrganisationID"].ToString();
                    UHIDPopUp.Visibility = Visibility.Collapsed;
                }
            }

        }
        #endregion

        #region btnUHIDGo_Click
        private void btnUHIDGo_Click(object sender, RoutedEventArgs e)
        {
            SqlParameter[] sqlParamSearch = new SqlParameter[]
         {
                new SqlParameter("@ACTIVITY", "GetWpfUHIDData"),
                new SqlParameter("@Searchtxt", txtSearch.Text),
                new SqlParameter("@FromDate", Convert.ToDateTime(dtpUHIDFrom.Text)),
                new SqlParameter("@ToDate", Convert.ToDateTime(dtpUHIDTo.Text)),
                new SqlParameter("@LocationID", gvars.gLocationId),
         };
            DataSet dsSearch = obj.getDataset("tSalesRet", DataHelper.SqlCmdType.sqlStoredProcedure, sqlParamSearch);
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
                    lblUHID.Text = dr["UHID"].ToString() + "/" + dr["IPDNo"].ToString();
                    NameLabel.Text = dr["Name"].ToString();
                    AgeGenderLabel.Text = dr["Age"].ToString() + "/" + dr["Gender"].ToString();
                    GuardianLabel.Text = dr["Guardian"].ToString();
                    PhoneLabel.Text = dr["Phone"].ToString();
                    AddressLabel.Text = dr["Address"].ToString();
                    DOALabel.Text = Convert.ToDateTime(dr["AdmDate"]).ToString("dd-MMM-yyyy");
                    PrimaryDrLabel.Text = dr["PrimDocName"].ToString();
                    DocIDLabel.Text = dr["DocId"].ToString();
                    AdmissionLabel.Text = dr["AdmissonType"].ToString();
                    BedRoomNoLabel.Text = dr["BedNo"].ToString();
                    BedIDLabel.Text = dr["BedID"].ToString();
                    RoomTypeLabel.Text = dr["RoomType"].ToString();
                    PayerLabel.Text = dr["PayType"].ToString();
                    OrgIDLabel.Text = dr["OrganisationID"].ToString();
                    UHIDPopUp.Visibility = Visibility.Collapsed;
                }
            }
        }
        #endregion

        #region btnPOPCalc_Click
        private void btnPOPCalc_Click(object sender, RoutedEventArgs e)
        {

            UHIDPopUp.IsOpen = false;
        }
        #endregion

        #region CheckIndentData
        public void CheckIndentData()
        {
            if (txtIndentPhone.Text != "")
            {
                SqlConnection con = new SqlConnection();
                obj.OpenDBCon();
                con = obj.getConnection();


                int iChkCount = 0;


                if (con.State != System.Data.ConnectionState.Open)
                    con.Open();

                using (SqlCommand cmd = new SqlCommand(
                    "SELECT COUNT(*) FROM trnSalesIndentMst WHERE IndentNo=@BillNo AND IssueStatus='Not Issue'", con))
                {
                    cmd.Parameters.AddWithValue("@BillNo", txtIndentPhone.Text.Trim());
                    iChkCount = Convert.ToInt32(cmd.ExecuteScalar());
                }


                if (iChkCount == 0)
                {
                    MessageBox.Show(
                       $"Their is no Indent No with this " + txtIndentPhone.Text + " Number !",
                       "Error",
                       MessageBoxButton.OK,
                       MessageBoxImage.Error
                   );
                }
                else
                {
                    GetIndentPatientDetails();
                }

                return;
            }
        }
        #endregion

        #region GetIndentPatientDetails
        private void GetIndentPatientDetails()
        {

            SqlParameter[] sqlParamSearch = new SqlParameter[]
              {
                     new SqlParameter("@ACTIVITY", "IndentData"),
                     new SqlParameter("@IndentNo", txtIndentPhone.Text),
                     new SqlParameter("@DepartmentID", gvars.gDeptID),
                     new SqlParameter("@LocationID", gvars.gLocationId),
              };
            ds = obj.getDataset("tIPSales", DataHelper.SqlCmdType.sqlStoredProcedure, sqlParamSearch);
            decimal TotalAmount = 0;
            if (ds.Tables[0].Rows.Count > 0)
            {
                //   TotalAmount =((TotalAmount) + Convert.ToDecimal(ds.Tables[0].Rows[0]["NetAmount"]));
                //if (!ds.Tables[0].Columns.Contains("Slno"))
                //    {
                //    ds.Tables[0].Columns.Add("Slno", typeof(int));
                //    }
                //    for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                //    {
                //    ds.Tables[0].Rows[i]["Slno"] = i + 1;
                //    }
                //   dgvItemDetails.ItemsSource = ds.Tables[0].DefaultView;

                //    txtTotalAmt.Text = Convert.ToString(TotalAmount);
                //    txtDueAmt.Text = Convert.ToString(TotalAmount);

                if (ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
                {
                    DataTable dt = ds.Tables[0];

                    // Create list to hold AddItems


                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        DataRow row = dt.Rows[i];

                        aItem.Add(new AddItems
                        {
                            Slno = i + 1,
                            ItemId = row["ItemID"].ToString(),
                            ItemName = row["ItemName"].ToString(),
                            Rack = "",
                            Tray = "",
                            Qty = Convert.ToInt32(row["Qty"]),
                            UnitMrp = Convert.ToDecimal(row["UnitMrp"]),
                            UnitRate = Convert.ToDecimal(row["UnitRate"]),
                            BatchNo = row["BatchNo"].ToString(),
                            ExpiryDate = row["ExpiryDate"].ToString(),
                            Discount = Convert.ToDecimal(row["Discount"]),
                            DiscPer = Convert.ToDecimal(row["DiscPer"]),
                            Amount = Convert.ToDecimal(row["Amount"]),
                            TaxPer = Convert.ToDecimal(row["TaxPer"]),
                            TaxAmount = Convert.ToDecimal(row["TaxAmount"]),
                            SupplierID = row["SupplierID"].ToString(),
                            PurchBillNo = row["PurchBillNo"].ToString(),
                            BatchQty = Convert.ToDecimal(row["BatchQty"]),
                            StockQty = Convert.ToDecimal(row["stockQty"]),
                            IsHigh = isHigh,
                            IsLook = isLook,
                            IsSound = isSound
                        });

                        // accumulate total amount
                        TotalAmount += Convert.ToDecimal(row["NetAmount"]);
                    }

                    dgvItemDetails.ItemsSource = aItem;
                    DataContext = this;

                    if (PharamacyRoundOffType == "RoundOff")
                    {
                        txtTotalAmt.Text = Math.Round(TotalAmount, 0).ToString("N2");
                        txtDueAmt.Text = Math.Round(TotalAmount, 0).ToString("N2");
                    }
                    if (PharamacyRoundOffType == "Ceil")
                    {
                        txtTotalAmt.Text = Math.Ceiling(TotalAmount).ToString("N2");
                        txtDueAmt.Text = Math.Ceiling(TotalAmount).ToString("N2");
                    }
                    if (PharamacyRoundOffType == "Floor")
                    {
                        txtTotalAmt.Text = Math.Floor(TotalAmount).ToString("N2");
                        txtDueAmt.Text = Math.Floor(TotalAmount).ToString("N2");
                    }
                    else
                    {
                        txtTotalAmt.Text = TotalAmount.ToString("0.00");
                        txtDueAmt.Text = TotalAmount.ToString("0.00");
                    }


                    decimal finalValue = 0;
                    decimal.TryParse(txtTotalAmt.Text, out finalValue);

                    decimal roundDiff = finalValue - TotalAmount;


                    txtRoundOFType.Text = $"{PharamacyRoundOffType} ({roundDiff:0.00})";
                    txtRoundOFType.Foreground = new SolidColorBrush(Colors.Red);
                }
                else
                {
                    MessageBox.Show("No records found for the given Indent No.");
                }
            }

            if (ds.Tables[1].Rows.Count > 0)
            {
                txtUHIDPhone.Text = ds.Tables[1].Rows[0]["OPIPDNo"].ToString();

                //lblIndentNo.Text = ds.Tables[1].Rows[0]["IndentNo"].ToString();
                //lblIndentDate.Text = Convert.ToDateTime(ds.Tables[1].Rows[0]["IndentDate"]).ToString("dd-MMM-yyyy");
                //IndentbyLabel.Text = ds.Tables[1].Rows[0]["CreateUserID"].ToString();
                //StatusLabel.Text = ds.Tables[1].Rows[0]["IssueStatus"].ToString();
                //RemarksLabel.Text = ds.Tables[1].Rows[0]["Remarks"].ToString();
                GetPatientDetails();
            }

          //  IndentPopUp.Visibility = Visibility.Collapsed;

        }
        #endregion

        #region btnIndentGo_Click
        //private void btnIndentGo_Click(object sender, RoutedEventArgs e)
        //{
        //    SqlParameter[] sqlParamSearch = new SqlParameter[]
        //    {
        //            new SqlParameter("@ACTIVITY", "GetIndentData"),
        //            new SqlParameter("@NursingStationID", cmbNursingstation.SelectedValue),
        //            new SqlParameter("@RoomTypeID", cmbWard.SelectedValue),
        //            new SqlParameter("@RoomID",cmbRooms.SelectedValue),
        //            new SqlParameter("@IndentNo",txtIndentSearch.Text),
        //            new SqlParameter("@IndentType",cmbType.SelectedValue),
        //            new SqlParameter("@DepartmentID", gvars.gDeptID),
        //            new SqlParameter("@LocationID", gvars.gLocationId),
        //    };
        //    DataSet dsSearch = obj.getDataset("tIPSales", DataHelper.SqlCmdType.sqlStoredProcedure, sqlParamSearch);
        //    if (dsSearch.Tables.Count > 0 && dsSearch.Tables[0].Rows.Count > 0)
        //    {
        //        if (!dsSearch.Tables[0].Columns.Contains("Slno"))
        //        {
        //            dsSearch.Tables[0].Columns.Add("Slno", typeof(int));
        //        }
        //        for (int i = 0; i < dsSearch.Tables[0].Rows.Count; i++)
        //        {
        //            dsSearch.Tables[0].Rows[i]["Slno"] = i + 1;
        //        }
        //      //  dgvIndentDetails.ItemsSource = dsSearch.Tables[0].DefaultView;
        //    }
        //    else
        //    {
        //        MessageBox.Show("No Data Found", "Information", MessageBoxButton.OK, MessageBoxImage.Information);
        //    }
        //}
        #endregion

        #region btnIndentView_Click
        private void btnIndentView_Click(object sender, RoutedEventArgs e)
        {
           // IndentPopUp.IsOpen = false;
           // if (dgvIndentDetails.SelectedItems.Count > 0)
           // {
                var button = sender as Button;
                var rowData = button?.DataContext;
                DataRowView dr = (DataRowView)rowData;

                if (dr != null)
                {
                    txtIndentPhone.Text = dr["IndentNo"].ToString();

                    SqlParameter[] sqlParamSearch = new SqlParameter[]
                    {
                            new SqlParameter("@ACTIVITY", "GetIndentViewData"),
                            new SqlParameter("@IndentNo", txtIndentPhone.Text),
                            new SqlParameter("@LocationID", gvars.gLocationId),
                    };
                    ds = obj.getDataset("tIPSales", DataHelper.SqlCmdType.sqlStoredProcedure, sqlParamSearch);
                    if (ds.Tables[0].Rows.Count > 0)
                    {
                        lblUHID.Text = ds.Tables[0].Rows[0]["UHID"].ToString() + "/" + ds.Tables[0].Rows[0]["IPDNo"].ToString();
                        NameLabel.Text = ds.Tables[0].Rows[0]["Name"].ToString();
                        AgeGenderLabel.Text = ds.Tables[0].Rows[0]["Age"].ToString() + "/" + ds.Tables[0].Rows[0]["Gender"].ToString();
                        GuardianLabel.Text = ds.Tables[0].Rows[0]["Guardian"].ToString();
                        PhoneLabel.Text = ds.Tables[0].Rows[0]["Phone"].ToString();
                        AddressLabel.Text = ds.Tables[0].Rows[0]["Address"].ToString();
                        DOALabel.Text = Convert.ToDateTime(ds.Tables[0].Rows[0]["AdmissionDate"]).ToString("dd-MMM-yyyy");
                        PrimaryDrLabel.Text = ds.Tables[0].Rows[0]["DocName"].ToString();
                        DocIDLabel.Text = ds.Tables[0].Rows[0]["DocId"].ToString();
                        AdmissionLabel.Text = ds.Tables[0].Rows[0]["AdmissonType"].ToString();
                        BedRoomNoLabel.Text = ds.Tables[0].Rows[0]["BedNo"].ToString();
                        BedIDLabel.Text = ds.Tables[0].Rows[0]["BedID"].ToString();
                        RoomTypeLabel.Text = ds.Tables[0].Rows[0]["RoomType"].ToString();
                        PayerLabel.Text = ds.Tables[0].Rows[0]["PayType"].ToString();
                        OrgIDLabel.Text = ds.Tables[0].Rows[0]["OrganisationID"].ToString();

                        //lblIndentNo.Text = ds.Tables[0].Rows[0]["IndentNo"].ToString();
                        //lblIndentDate.Text = Convert.ToDateTime(ds.Tables[0].Rows[0]["IndentDate"]).ToString("dd-MMM-yyyy");
                        //IndentbyLabel.Text = ds.Tables[0].Rows[0]["CreateUserID"].ToString();
                        //StatusLabel.Text = ds.Tables[0].Rows[0]["IssueStatus"].ToString();
                        //RemarksLabel.Text = ds.Tables[0].Rows[0]["Remarks"].ToString();
                       // IndentPopUp.Visibility = Visibility.Collapsed;

                        SqlParameter[] sqlParamSearch1 = new SqlParameter[]
                            {
                                new SqlParameter("@ACTIVITY", "IndentData"),
                                new SqlParameter("@IndentNo", txtIndentPhone.Text),
                                new SqlParameter("@DepartmentID", gvars.gDeptID),
                                new SqlParameter("@LocationID", gvars.gLocationId),
                            };

                        DataSet ds1 = obj.getDataset("tIPSales", DataHelper.SqlCmdType.sqlStoredProcedure, sqlParamSearch1);
                        decimal TotalAmount = 0;

                        if (ds1.Tables.Count > 0 && ds1.Tables[0].Rows.Count > 0)
                        {
                            DataTable dt = ds1.Tables[0];



                            for (int i = 0; i < dt.Rows.Count; i++)
                            {
                                DataRow row = dt.Rows[i];

                                aItem.Add(new AddItems
                                {
                                    Slno = i + 1,
                                    ItemId = row["ItemID"].ToString(),
                                    ItemName = row["ItemName"].ToString(),
                                    Rack = "",
                                    Tray = "",
                                    Qty = Convert.ToInt32(row["Qty"]),
                                    UnitMrp = Convert.ToDecimal(row["UnitMrp"]),
                                    UnitRate = Convert.ToDecimal(row["UnitRate"]),
                                    BatchNo = row["BatchNo"].ToString(),
                                    ExpiryDate = row["ExpiryDate"].ToString(),
                                    Discount = Convert.ToDecimal(row["Discount"]),
                                    DiscPer = Convert.ToDecimal(row["DiscPer"]),
                                    Amount = Convert.ToDecimal(row["Amount"]),
                                    TaxPer = Convert.ToDecimal(row["TaxPer"]),
                                    TaxAmount = Convert.ToDecimal(row["TaxAmount"]),
                                    SupplierID = row["SupplierID"].ToString(),
                                    PurchBillNo = row["PurchBillNo"].ToString(),
                                    BatchQty = Convert.ToDecimal(row["BatchQty"]),
                                    StockQty = Convert.ToDecimal(row["stockQty"]),
                                    IsHigh = isHigh,
                                    IsLook = isLook,
                                    IsSound = isSound
                                });

                                // accumulate total amount
                                TotalAmount += Convert.ToDecimal(row["NetAmount"]);
                            }

                            // Bind the list to DataGrid
                            // dgvItemDetails.ItemsSource = null;
                            dgvItemDetails.ItemsSource = aItem;
                            DataContext = this;

                            if (PharamacyRoundOffType == "RoundOff")
                            {
                                txtTotalAmt.Text = Math.Round(TotalAmount, 0).ToString("N2");
                                txtDueAmt.Text = Math.Round(TotalAmount, 0).ToString("N2");
                            }
                            if (PharamacyRoundOffType == "Ceil")
                            {
                                txtTotalAmt.Text = Math.Ceiling(TotalAmount).ToString("N2");
                                txtDueAmt.Text = Math.Ceiling(TotalAmount).ToString("N2");
                            }
                            if (PharamacyRoundOffType == "Floor")
                            {
                                txtTotalAmt.Text = Math.Floor(TotalAmount).ToString("N2");
                                txtDueAmt.Text = Math.Floor(TotalAmount).ToString("N2");
                            }
                            else
                            {
                                txtTotalAmt.Text = TotalAmount.ToString("0.00");
                                txtDueAmt.Text = TotalAmount.ToString("0.00");
                            }


                            decimal finalValue = 0;
                            decimal.TryParse(txtTotalAmt.Text, out finalValue);

                            decimal roundDiff = finalValue - TotalAmount;


                            txtRoundOFType.Text = $"{PharamacyRoundOffType} ({roundDiff:0.00})";
                            txtRoundOFType.Foreground = new SolidColorBrush(Colors.Red);


                        }
                        else
                        {
                            MessageBox.Show("No records found for the given Indent No.");
                        }

                    }
               // }
            }
        }
        #endregion

        #region btnPOPIndent_Click
        private void btnPOPIndent_Click(object sender, RoutedEventArgs e)
        {
           // IndentPopUp.IsOpen = false;
        }
        #endregion

        #region ItemSelectionChange
        private void cmbItemName_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (cmbItemName.SelectedIndex > 0)
            {
                GetBatch(cmbItemName.SelectedValue.ToString());
                txtQty.Focus();
            }
        }
        #endregion

        #region GetBatch
        public void GetBatch(string ItemId)
        {
            cmbBatchNo.Text = "";
           // txtBQty.Text = "";
            txtTQty.Text = "";
            dtpExpDt.Text = "";
            txtQty.Text = "";
            txtMRP.Text = "";
            txtTaxAmt.Text = "";
            txtAmount.Text = "";
            DataHelper objDataHelper1 = new DataHelper();
            SqlParameter[] sqlParamSearch = new SqlParameter[]
            {
               new SqlParameter("@ACTIVITY", "Getbatch"),
               new SqlParameter("@ItemID", ItemId),
               new SqlParameter("@LocationID", gvars.gLocationId),
               new SqlParameter("@DepartmentID", gvars.gDeptID),
            };
            ds = obj.getDataset("tIPSales", DataHelper.SqlCmdType.sqlStoredProcedure, sqlParamSearch);
            if (ds.Tables[0].Rows.Count > 0 && ds.Tables.Count > 0)
            {
                dsBatch = ds;
                cmbBatchNo.SelectedValuePath = "BatchNo";
                cmbBatchNo.DisplayMemberPath = "BatchNo";
                cmbBatchNo.SelectionChanged -= cmbBatchNo_SelectionChanged;

                cmbBatchNo.ItemsSource = ds.Tables[0].DefaultView;
                cmbBatchNo.SelectedIndex = 0;
                cmbBatchNo.SelectionChanged += cmbBatchNo_SelectionChanged;

                long daysLeft = Convert.ToInt64(ds.Tables[0].Rows[0]["NoOfDaysMsg"]);
                long SaleStopExpMsg = Convert.ToInt64(ds.Tables[0].Rows[0]["SaleStopExpMsg"]);
                long SalesShowExpMsg = Convert.ToInt64(ds.Tables[0].Rows[0]["SalesShowExpMsg"]);

                string itemName = ds.Tables[0].Rows[0]["ItemName"].ToString();
                string batchNo = ds.Tables[0].Rows[0]["BatchNo"].ToString();
                DateTime expiryDate = Convert.ToDateTime(ds.Tables[0].Rows[0]["ExpiryDate"]);

                if (daysLeft <= SaleStopExpMsg)
                {
                    MessageBox.Show(
                        $"{itemName} with Batch No. {batchNo} is going to expire on {expiryDate:dd MMM yyyy} " +
                        $"({daysLeft} days left). Item cannot be added.",
                        "Error",
                        MessageBoxButton.OK,
                        MessageBoxImage.Error
                    );
                  //  txtBQty.Text = "";
                    cmbBatchNo.Text = "";
                    dtpExpDt.Text = "";
                    txtMRP.Text = "";
                    txtTQty.Text = "";
                    return;
                }
                if (daysLeft <= SalesShowExpMsg)
                {
                    MessageBox.Show(
                        $"{itemName} with Batch No. {batchNo} is going to expire on {expiryDate:dd MMM yyyy} " +
                        $"({daysLeft} days left).",
                        "warning",
                        MessageBoxButton.OK,
                        MessageBoxImage.Error
                    );
                }



              //  txtBQty.Text = ds.Tables[0].Rows[0]["Qty"].ToString();
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
            else
            {
                MessageBox.Show("No Stock available", "NoStock", MessageBoxButton.OK, MessageBoxImage.Information);
                cmbItemName.SelectedIndex = -1;
                cmbItemName.Focus();
            }

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

        #region cmbBatchNo_SelectionChanged
        private void cmbBatchNo_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (cmbBatchNo.SelectedIndex > -1)
            {
                if (dsBatch != null)
                {
                    DataRow[] drow = dsBatch.Tables[0].Select(" BatchNo ='" + cmbBatchNo.SelectedValue + "'");
                    if (drow != null && drow.Length > 0)
                    {
                        long daysLeft = Convert.ToInt64(drow[0].ItemArray[31]);
                        long SaleStopExpMsg = Convert.ToInt64(drow[0].ItemArray[30]);
                        long SalesShowExpMsg = Convert.ToInt64(drow[0].ItemArray[29]);

                        string itemName = drow[0].ItemArray[1].ToString();
                        string batchNo = drow[0].ItemArray[4].ToString();
                        var expiryDate = Convert.ToDateTime(drow[0].ItemArray[5]).ToString("yyyy-MM-dd");

                        if (daysLeft <= SaleStopExpMsg)
                        {
                            MessageBox.Show(
                                $"{itemName} with Batch No. {batchNo} is going to expire on {expiryDate:dd MMM yyyy} " +
                                $"({daysLeft} days left). Item cannot be added.",
                                "Error",
                                MessageBoxButton.OK,
                                MessageBoxImage.Error
                            );
                           // txtBQty.Text = "";
                            cmbBatchNo.Text = "";
                            dtpExpDt.Text = "";
                            txtMRP.Text = "";
                            txtTQty.Text = "";
                            return;
                        }
                        if (daysLeft <= SalesShowExpMsg)
                        {
                            MessageBox.Show(
                                $"{itemName} with Batch No. {batchNo} is going to expire on {expiryDate:dd MMM yyyy} " +
                                $"({daysLeft} days left).",
                                "warning",
                                MessageBoxButton.OK,
                                MessageBoxImage.Error
                            );
                        }

                        cmbBatchNo.Text = drow[0].ItemArray[4].ToString();
                        dtpExpDt.Text = Convert.ToDateTime(drow[0].ItemArray[5]).ToString("yyyy-MM-dd");
                        txtMRP.Text = drow[0].ItemArray[13].ToString();
                        Rate = Convert.ToDecimal(drow[0].ItemArray[14]);
                        taxPer = Convert.ToDecimal(drow[0].ItemArray[6]);
                        PBillNo = drow[0].ItemArray[10].ToString();
                        SupId = drow[0].ItemArray[12].ToString();
                       // txtBQty.Text = drow[0].ItemArray[15].ToString();
                        txtTQty.Text = drow[0].ItemArray[15].ToString();
                        //txtTQty.Text = dsBatch.Tables[0].Compute("Sum(Qty)", "").ToString();
                    }
                }
            }
        }
        #endregion


        #region btnAdd_Click
        private void btnAdd_Click(object sender, RoutedEventArgs e)
        {
            if (cmbItemName.SelectedIndex == 0)
            {
                MessageBox.Show("Select Item ", "Information", MessageBoxButton.OK, MessageBoxImage.Information);
                cmbItemName.Focus();
                return;
            }
            if (txtQty.Text == "0" || txtQty.Text == "")
            {
                MessageBox.Show("Required Quantity should be greater than 0 ", "Information", MessageBoxButton.OK, MessageBoxImage.Information);
                txtQty.Focus();
                return;
            }
            if (cmbBatchNo.SelectedIndex == -1)
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
                    if (p.ItemId == cmbItemName.SelectedValue.ToString() && p.BatchNo == cmbBatchNo.Text && p.ExpiryDate == dtpExpDt.Text && p.UnitRate == Rate && p.UnitMrp == Convert.ToDecimal(txtMRP.Text) && p.PurchBillNo == PBillNo && p.SupplierID == SupId)
                    {
                        MessageBox.Show("You have already added this Item ", "Information", MessageBoxButton.OK, MessageBoxImage.Information);
                        cmbItemName.Focus();
                        return;
                    }
                }
            }
            decimal SavedQty = 0, reqStock = Convert.ToDecimal(txtQty.Text);


            if (Convert.ToDecimal(txtQty.Text) >= Convert.ToDecimal(txtQty.Text)) //Convert.ToDecimal(txtBQty.Text)
            {

                strQry = "SELECT ISNULL(SUM(s.Qty),0) AS Qty FROM trnStock s " +
                         " WHERE S.LocationID = '" + gvars.gLocationId + "' AND UnitMrp=" + txtMRP.Text + " AND UnitRate=" + Rate + " AND BatchNo='" + cmbBatchNo.Text + "' " +
                         " AND ExpiryDate='" + Convert.ToDateTime(dtpExpDt.Text).ToString("yyyy-MM-dd") + "' AND SupplierID='" + SupId + "' AND PurchBillNo='" + PBillNo + "' AND ItemID='" + cmbItemName.SelectedValue + "' " +
                         " AND DepartmentID='" + gvars.gDeptID + "'";
                decimal cStock = Convert.ToDecimal(obj.ExecuteScalar(strQry, DataHelper.SqlCmdType.sqlText));
                if (cStock >= Convert.ToInt32(txtQty.Text))
                {
                    aItem.Add(new AddItems
                    {
                        Slno = (dgvItemDetails.Items.Count + 1),
                        ItemId = cmbItemName.SelectedValue.ToString(),
                        ItemName = cmbItemName.Text.ToString(),
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
                       // BatchQty = Convert.ToDecimal(txtBQty.Text),
                      //  StockQty = Convert.ToDecimal(txtBQty.Text) - Convert.ToDecimal(txtQty.Text),
                        IsHigh = isHigh,
                        IsLook = isLook,
                        IsSound = isSound
                    });
                    dgvItemDetails.ItemsSource = aItem;
                    DataContext = this;
                    UpdatTotals();
                    ClearItems();
                }
                else
                {
                    MessageBox.Show("No Stock available", "Information", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }
            }
            else
            {
                for (int j = 0; j < dsBatch.Tables[0].Rows.Count - 1; j++)
                {
                    strQry = "SELECT ISNULL(SUM(s.Qty),0) AS Qty FROM trnStock s " +
                             " WHERE S.LocationID = '" + gvars.gLocationId + "' AND UnitMrp=" + dsBatch.Tables[0].Rows[j]["UnitMrp"] + " AND UnitRate=" + dsBatch.Tables[0].Rows[j]["UnitRate"] + " AND BatchNo='" + dsBatch.Tables[0].Rows[j]["BatchNo"] + "' " +
                             " AND ExpiryDate='" + Convert.ToDateTime(dsBatch.Tables[0].Rows[j]["ExpiryDate"]).ToString("dd-MMM-yyyy") + "' AND SupplierID='" + dsBatch.Tables[0].Rows[j]["SupplierID"] + "' " +
                             " AND PurchBillNo='" + dsBatch.Tables[0].Rows[j]["PurchBillNo"] + "' AND ItemID='" + cmbItemName.SelectedValue + "' " +
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
                                ItemId = cmbItemName.SelectedValue.ToString(),
                                ItemName = cmbItemName.Text.ToString(),
                                Rack = "",
                                Tray = "",
                                Qty = cStock,
                                UnitMrp = Convert.ToDecimal(dsBatch.Tables[0].Rows[j]["UnitMrp"]),
                                UnitRate = Convert.ToDecimal(dsBatch.Tables[0].Rows[j]["UnitRate"]),
                                BatchNo = dsBatch.Tables[0].Rows[j]["BatchNo"].ToString(),
                                ExpiryDate = Convert.ToDateTime(dsBatch.Tables[0].Rows[j]["ExpiryDate"]).ToString("dd-MMM-yyyy"),
                                Discount = 0,
                                DiscPer = 0,
                                Amount = (Convert.ToDecimal(dsBatch.Tables[0].Rows[j]["UnitMrp"]) * SavedQty),//Math.Round
                                TaxPer = Convert.ToDecimal(dsBatch.Tables[0].Rows[j]["TaxPer"]),
                                TaxAmount = (((Convert.ToDecimal(dsBatch.Tables[0].Rows[j]["UnitMrp"]) * SavedQty) * taxPer) / 100),//Math.Round
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
                        }
                        else
                        {
                            SavedQty = SavedQty + reqStock;
                            aItem.Add(new AddItems
                            {
                                Slno = (dgvItemDetails.Items.Count + 1),
                                ItemId = cmbItemName.SelectedValue.ToString(),
                                ItemName = cmbItemName.Text.ToString(),
                                Rack = "",
                                Tray = "",
                                Qty = reqStock,
                                UnitMrp = Convert.ToDecimal(dsBatch.Tables[0].Rows[j]["UnitMrp"]),
                                UnitRate = Convert.ToDecimal(dsBatch.Tables[0].Rows[j]["UnitRate"]),
                                BatchNo = dsBatch.Tables[0].Rows[j]["BatchNo"].ToString(),
                                ExpiryDate = Convert.ToDateTime(dsBatch.Tables[0].Rows[j]["ExpiryDate"]).ToString("dd-MMM-yyyy"),
                                Discount = 0,
                                DiscPer = 0,
                                Amount = (Convert.ToDecimal(dsBatch.Tables[0].Rows[j]["UnitMrp"]) * reqStock), //Math.Round
                                TaxPer = Convert.ToDecimal(dsBatch.Tables[0].Rows[j]["TaxPer"]),
                                TaxAmount = (((Convert.ToDecimal(dsBatch.Tables[0].Rows[j]["UnitMrp"]) * reqStock) * taxPer) / 100), //Math.Round
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
                        return;
                    }
                }

            }
            if (!string.IsNullOrEmpty(txtConcPer.Text) && Convert.ToDecimal(txtConcPer.Text) != 0)
            {
                var collection = dgvPaymentDet.ItemsSource as ObservableCollection<AddPayDet>;
                collection?.Clear();
                //  dgvPaymentDet.ItemsSource = null;


                cmbDiscType.SelectedIndex = 0;
                txtConcPer.Text = "0";
                txtConcAmt.Text = "0";
                txtDiscReason.Text = "";
                txtPayAmount.Text = "0";
                txtDiscount.Text = "0";
                txtPaidAmt.Text = "0";
                cmbDueAuth.SelectedIndex = 0;
                txtDueReason.Text = "";
                cmbDiscAuth.SelectedIndex = 0;
                txtDiscReason.Text = "";

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
                    if (payDet.Count == 0)
                    {
                        PayResetDetails();

                        txtConcPer.IsEnabled = true;
                        txtConcAmt.IsEnabled = true;
                    }
                }
                foreach (var row in dgvItemDetails.Items)
                {
                    if (row == null || row == CollectionView.NewItemPlaceholder)
                        continue;


                    var type = row.GetType();
                    var discPerProp = type.GetProperty("DiscPer");
                    var discountProp = type.GetProperty("Discount");

                    if (discPerProp != null)
                        discPerProp.SetValue(row, 0m, null);

                    if (discountProp != null)
                        discountProp.SetValue(row, 0m, null);
                }

                dgvItemDetails.Items.Refresh();
            }
        }
        #endregion
        private void ClearItems()
        {
            cmbItemName.SelectedIndex = 0;
             txtAmount.Text = txtQty.Text = txtTaxAmt.Text = txtMRP.Text = txtTQty.Text = ""; //txtBQty.Text =
            cmbBatchNo.ItemsSource = null;
            PBillNo = "";
            SupId = "";
            Rate = 0;
            taxPer = 0;
        }
        #endregion

        #region GetRoundOffTypeamt
        private void GetRoundOffTypeamt()
        {
            string strQry = "SELECT ISNULL(IsReqPharmacyRoundOff,0) AS IsReqPharmacyRoundOff, PharamacyRoundOffType " +
                            "FROM mstDepartmentDtls WHERE DepartmentID = '" + gvars.gDeptID + "'";

            DataSet ds = obj.getDataset(strQry, DataHelper.SqlCmdType.sqlText);
            if (ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
            {
                DataRow dr = ds.Tables[0].Rows[0];
                bool isReqPharmacyRoundOff = Convert.ToBoolean(dr["IsReqPharmacyRoundOff"]);

                if (isReqPharmacyRoundOff == true)
                {
                    string roundOffType = dr["PharamacyRoundOffType"].ToString();
                    PharamacyRoundOffType = roundOffType;
                }
            }
        }
        #endregion

        #region GetPaymentRights
        private void GetPaymentRights()
        {
            string strQry = "SELECT ISNULL(IsIPSalesPayment,0) as IsIPSalesPayment FROM mstLocation L WHERE L.LocationID= '" + gvars.gLocationId + "'";

            DataSet ds = obj.getDataset(strQry, DataHelper.SqlCmdType.sqlText);
            if (ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
            {
                DataRow dr = ds.Tables[0].Rows[0];
                bool IsIPSalesPayment = Convert.ToBoolean(dr["IsIPSalesPayment"]);

                //if (IsIPSalesPayment == true)
                //{
                //    spPayDetails.Visibility = Visibility.Visible;
                //    dgvPaymentDet.Visibility = Visibility.Visible;
                //    DIACPayDetails.Visibility = Visibility.Visible;
                //}
                //else
                //{
                //    spPayDetails.Visibility = Visibility.Collapsed;
                //    dgvPaymentDet.Visibility = Visibility.Collapsed;
                //    DIACPayDetails.Visibility = Visibility.Collapsed;
                //}
            }
        }
        #endregion

        #region QtyTextbox
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
                // PayResetDetails();

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
                    //if (PharamacyRoundOffType == "RoundOff")
                    //{
                    //    txtTotalAmt.Text = Math.Round(totalAmount, 0).ToString("N2");
                    //    txtDueAmt.Text = Math.Round(totalAmount, 0).ToString("N2");
                    //}
                    //else if (PharamacyRoundOffType == "Ceil")
                    //{
                    //    txtTotalAmt.Text = Math.Ceiling(totalAmount).ToString("N2");
                    //    txtDueAmt.Text = Math.Ceiling(totalAmount).ToString("N2");
                    //}
                    //else if (PharamacyRoundOffType == "Floor")
                    //{
                    //    txtTotalAmt.Text = Math.Floor(totalAmount).ToString("N2");
                    //    txtDueAmt.Text = Math.Floor(totalAmount).ToString("N2");
                    //}
                    //else
                    //{
                    //    txtTotalAmt.Text = totalAmount.ToString("N2");
                    //    txtDueAmt.Text = totalAmount.ToString("N2");
                    //}

                    //----------------
                    if (PharamacyRoundOffType == "RoundOff")
                    {
                        txtTotalAmt.Text = Math.Round(totalAmount, 0).ToString("N2");
                        txtDueAmt.Text = Math.Round(totalAmount, 0).ToString("N2");
                    }
                    if (PharamacyRoundOffType == "Ceil")
                    {
                        txtTotalAmt.Text = Math.Ceiling(totalAmount).ToString("N2");
                        txtDueAmt.Text = Math.Ceiling(totalAmount).ToString("N2");
                    }
                    if (PharamacyRoundOffType == "Floor")
                    {
                        txtTotalAmt.Text = Math.Floor(totalAmount).ToString("N2");
                        txtDueAmt.Text = Math.Floor(totalAmount).ToString("N2");
                    }
                    else
                    {
                        txtTotalAmt.Text = totalAmount.ToString("0.00");
                        txtDueAmt.Text = totalAmount.ToString("0.00");
                    }


                    decimal finalValue = 0;
                    decimal.TryParse(txtTotalAmt.Text, out finalValue);

                    decimal roundDiff = finalValue - totalAmount;


                    txtRoundOFType.Text = $"{PharamacyRoundOffType} ({roundDiff:0.00})";
                    txtRoundOFType.Foreground = new SolidColorBrush(Colors.Red);

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
                    if (PharamacyRoundOffType == "RoundOff")
                    {
                        txtTotalAmt.Text = Math.Round(totalAmount, 0).ToString("N2");
                        txtDueAmt.Text = Math.Round(totalAmount, 0).ToString("N2");
                    }
                    else if (PharamacyRoundOffType == "Ceil")
                    {
                        txtTotalAmt.Text = Math.Ceiling(totalAmount).ToString("N2");
                        txtDueAmt.Text = Math.Ceiling(totalAmount).ToString("N2");
                    }
                    else if (PharamacyRoundOffType == "Floor")
                    {
                        txtTotalAmt.Text = Math.Floor(totalAmount).ToString("N2");
                        txtDueAmt.Text = Math.Floor(totalAmount).ToString("N2");
                    }
                    else
                    {
                        txtTotalAmt.Text = totalAmount.ToString("N2");
                        txtDueAmt.Text = totalAmount.ToString("N2");
                    }
                    decimal finalValue = 0;
                    decimal.TryParse(txtTotalAmt.Text, out finalValue);

                    decimal roundDiff = finalValue - totalAmount;


                    txtRoundOFType.Text = $"{PharamacyRoundOffType} ({roundDiff:0.00})";
                    txtRoundOFType.Foreground = new SolidColorBrush(Colors.Red);

                    //
                    //if (Convert.ToDecimal(txtConcPer.Text) > 0 || Convert.ToDecimal(txtConcPer.Text) > 0)
                    if (!string.IsNullOrWhiteSpace(txtConcPer.Text) && Convert.ToDecimal(txtConcPer.Text) > 0)
                    {
                        if (cmbDiscType.SelectedIndex == 1)
                        {
                            if (PharamacyRoundOffType == "RoundOff")
                            {
                                txtConcAmt.Text = txtDiscount.Text = Math.Round(Convert.ToDecimal(txtTotalAmt.Text) * Convert.ToDecimal(txtConcPer.Text) / 100, 0).ToString();
                                txtConcPer.Text = Math.Round((Convert.ToDecimal(txtConcAmt.Text) / Convert.ToDecimal(txtTotalAmt.Text) * 100), 0).ToString();

                            }
                            else if (PharamacyRoundOffType == "Ceil")
                            {
                                txtConcAmt.Text = txtDiscount.Text = Math.Ceiling(Convert.ToDecimal(txtTotalAmt.Text) * Convert.ToDecimal(txtConcPer.Text) / 100).ToString();
                                txtConcPer.Text = Math.Ceiling((Convert.ToDecimal(txtConcAmt.Text) / Convert.ToDecimal(txtTotalAmt.Text) * 100)).ToString();
                            }
                            else if (PharamacyRoundOffType == "Floor")
                            {
                                txtConcAmt.Text = txtDiscount.Text = Math.Floor(Convert.ToDecimal(txtTotalAmt.Text) * Convert.ToDecimal(txtConcPer.Text) / 100).ToString();
                                txtConcPer.Text = Math.Floor((Convert.ToDecimal(txtConcAmt.Text) / Convert.ToDecimal(txtTotalAmt.Text) * 100)).ToString();
                            }
                            else
                            {
                                txtConcAmt.Text = txtDiscount.Text = (Convert.ToDecimal(txtTotalAmt.Text) * Convert.ToDecimal(txtConcPer.Text) / 100).ToString();
                                txtConcPer.Text = Math.Round((Convert.ToDecimal(txtConcAmt.Text) / Convert.ToDecimal(txtTotalAmt.Text) * 100)).ToString();
                            }
                            if (gvars.gConcType == "%")
                            {
                                if (Convert.ToDecimal(txtConcPer.Text) > gvars.gConcLimit)
                                {
                                    MessageBox.Show("You Are Not Allowed To Give Discount More than " + gvars.gConcLimit + " % Discount", "Infomation", MessageBoxButton.OK, MessageBoxImage.Information);
                                    txtConcPer.Text = txtConcAmt.Text = "0";
                                    txtConcPer.Focus();
                                    return;
                                }
                            }
                            else
                            {
                                if (Convert.ToDecimal(txtConcAmt.Text) > gvars.gConcLimit)
                                {
                                    MessageBox.Show("You Are Not Allowed To Give Discount More than " + gvars.gConcLimit + " Rs Discount", "Infomation", MessageBoxButton.OK, MessageBoxImage.Information);
                                    txtConcPer.Text = txtConcAmt.Text = "0";
                                    txtConcPer.Focus();
                                    return;
                                }
                            }
                            if (PharamacyRoundOffType == "RoundOff")
                            {
                                txtDueAmt.Text = Math.Round((Convert.ToDecimal(txtTotalAmt.Text) - Convert.ToDecimal(txtConcAmt.Text)), 0).ToString();
                            }
                            else if (PharamacyRoundOffType == "Ceil")
                            {
                                txtDueAmt.Text = Math.Ceiling((Convert.ToDecimal(txtTotalAmt.Text) - Convert.ToDecimal(txtConcAmt.Text))).ToString();
                            }
                            else if (PharamacyRoundOffType == "Floor")
                            {
                                txtDueAmt.Text = Math.Floor((Convert.ToDecimal(txtTotalAmt.Text) - Convert.ToDecimal(txtConcAmt.Text))).ToString();
                            }
                            else
                            {
                                txtDueAmt.Text = ((Convert.ToDecimal(txtTotalAmt.Text) - Convert.ToDecimal(txtConcAmt.Text))).ToString();
                            }
                            txtPayAmount.Text = txtDueAmt.Text;
                            decimal dueAmt;
                            if (decimal.TryParse(txtDueAmt.Text.Trim(), out dueAmt))
                            {
                                if (dueAmt > 0)
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
                        else
                        {
                            if (PharamacyRoundOffType == "RoundOff")
                            {
                                txtConcPer.Text = Math.Round((Convert.ToDecimal(txtConcAmt.Text) / Convert.ToDecimal(txtTotalAmt.Text) * 100), 0).ToString();
                                txtConcAmt.Text = txtDiscount.Text = Math.Round(Convert.ToDecimal(txtTotalAmt.Text) * Convert.ToDecimal(txtConcPer.Text) / 100, 2).ToString();
                            }
                            else if (PharamacyRoundOffType == "Ceil")
                            {
                                txtConcPer.Text = Math.Ceiling((Convert.ToDecimal(txtConcAmt.Text) / Convert.ToDecimal(txtTotalAmt.Text) * 100)).ToString();
                                txtConcAmt.Text = txtDiscount.Text = Math.Ceiling(Convert.ToDecimal(txtTotalAmt.Text) * Convert.ToDecimal(txtConcPer.Text) / 100).ToString();
                            }
                            else if (PharamacyRoundOffType == "Floor")
                            {
                                txtConcPer.Text = Math.Floor((Convert.ToDecimal(txtConcAmt.Text) / Convert.ToDecimal(txtTotalAmt.Text) * 100)).ToString();
                                txtConcAmt.Text = txtDiscount.Text = Math.Floor(Convert.ToDecimal(txtTotalAmt.Text) * Convert.ToDecimal(txtConcPer.Text) / 100).ToString();
                            }
                            else
                            {
                                txtConcPer.Text = ((Convert.ToDecimal(txtConcAmt.Text) / Convert.ToDecimal(txtTotalAmt.Text) * 100)).ToString();
                                txtConcAmt.Text = txtDiscount.Text = (Convert.ToDecimal(txtTotalAmt.Text) * Convert.ToDecimal(txtConcPer.Text) / 100).ToString();
                            }

                            if (gvars.gConcType == "%")
                            {
                                if (Convert.ToDecimal(txtConcPer.Text) > gvars.gConcLimit)
                                {
                                    MessageBox.Show("You Are Not Allowed To Give Discount More than " + gvars.gConcLimit + " % Discount", "Infomation", MessageBoxButton.OK, MessageBoxImage.Information);
                                    txtConcPer.Text = txtConcAmt.Text = "0";
                                    txtConcAmt.Focus();
                                    return;
                                }
                            }
                            else
                            {
                                if (Convert.ToDecimal(txtConcAmt.Text) > gvars.gConcLimit)
                                {
                                    MessageBox.Show("You Are Not Allowed To Give Discount More than " + gvars.gConcLimit + " Rs Discount", "Infomation", MessageBoxButton.OK, MessageBoxImage.Information);
                                    txtConcPer.Text = txtConcAmt.Text = "0";
                                    txtConcAmt.Focus();
                                    return;
                                }
                            }
                            txtDiscount.Text = txtConcAmt.Text;
                            if (PharamacyRoundOffType == "RoundOff")
                            {
                                txtDueAmt.Text = Math.Round(Convert.ToDecimal(txtTotalAmt.Text) - Convert.ToDecimal(txtConcAmt.Text)).ToString();
                            }
                            else if (PharamacyRoundOffType == "Ceil")
                            {
                                txtDueAmt.Text = Math.Ceiling(Convert.ToDecimal(txtTotalAmt.Text) - Convert.ToDecimal(txtConcAmt.Text)).ToString();
                            }
                            else if (PharamacyRoundOffType == "Floor")
                            {
                                txtDueAmt.Text = Math.Floor(Convert.ToDecimal(txtTotalAmt.Text) - Convert.ToDecimal(txtConcAmt.Text)).ToString();
                            }
                            else
                            {
                                txtDueAmt.Text = (Convert.ToDecimal(txtTotalAmt.Text) - Convert.ToDecimal(txtConcAmt.Text)).ToString();
                            }
                            txtPayAmount.Text = txtDueAmt.Text;
                            long dueAmt;
                            if (long.TryParse(txtDueAmt.Text, out dueAmt))
                            {
                                if (dueAmt > 0)
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
                                    if (PharamacyRoundOffType == "RoundOff")
                                    {
                                        item.Discount = Math.Round((item.Amount * Convert.ToDecimal(txtConcPer.Text)) / 100, 0);
                                    }
                                    else if (PharamacyRoundOffType == "Ceil")
                                    {
                                        item.Discount = Math.Ceiling((item.Amount * Convert.ToDecimal(txtConcPer.Text)) / 100);
                                    }
                                    else if (PharamacyRoundOffType == "Floor")
                                    {
                                        item.Discount = Math.Floor((item.Amount * Convert.ToDecimal(txtConcPer.Text)) / 100);
                                    }
                                    else
                                    {
                                        item.Discount = ((item.Amount * Convert.ToDecimal(txtConcPer.Text)) / 100);

                                    }

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
                                    if (PharamacyRoundOffType == "RoundOff")
                                    {
                                        item.Discount = Math.Round(item.Amount * Convert.ToDecimal(txtConcPer.Text)) / 100;
                                        NetAmount = item.Amount - item.Discount;
                                        item.TaxAmount = Math.Round((NetAmount * item.TaxPer) / 100, 0);
                                    }
                                    if (PharamacyRoundOffType == "Ceil")
                                    {
                                        item.Discount = Math.Ceiling(item.Amount * Convert.ToDecimal(txtConcPer.Text)) / 100;
                                        NetAmount = item.Amount - item.Discount;
                                        item.TaxAmount = Math.Ceiling((NetAmount * item.TaxPer) / 100);
                                    }
                                    else if (PharamacyRoundOffType == "Floor")
                                    {
                                        item.Discount = Math.Floor(item.Amount * Convert.ToDecimal(txtConcPer.Text)) / 100;
                                        NetAmount = item.Amount - item.Discount;
                                        item.TaxAmount = Math.Floor((NetAmount * item.TaxPer) / 100);
                                    }
                                    else
                                    {
                                        item.Discount = (item.Amount * Convert.ToDecimal(txtConcPer.Text)) / 100;
                                        NetAmount = item.Amount - item.Discount;
                                        item.TaxAmount = ((NetAmount * item.TaxPer) / 100);
                                    }
                                }
                            }
                        }
                    }
                    //
                    if (Convert.ToDecimal(txtPaidAmt.Text) > 0)
                    {
                        var collection = dgvPaymentDet.ItemsSource as ObservableCollection<AddPayDet>;
                        collection?.Clear();
                        txtPaidAmt.Text = "0";
                        txtPayAmount.Text = "0";
                    }

                }
            }

        }
        #endregion 

        #region UpdatTotals
        private void UpdatTotals()
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


                if (PharamacyRoundOffType == "RoundOff")
                {
                    txtTotalAmt.Text = Math.Round(Tot, 0).ToString("0.00");
                    txtDueAmt.Text = Math.Round(Tot, 0).ToString("0.00");
                    txtDiscount.Text = Math.Round(Disc, 0).ToString("0.00");
                }
                else if (PharamacyRoundOffType == "Ceil")
                {
                    txtTotalAmt.Text = Math.Ceiling(Tot).ToString("0.00");
                    txtDueAmt.Text = Math.Ceiling(Tot).ToString("0.00");
                    txtDiscount.Text = Math.Ceiling(Disc).ToString("0.00");
                }
                else if (PharamacyRoundOffType == "Floor")
                {
                    txtTotalAmt.Text = Math.Floor(Tot).ToString("0.00");
                    txtDueAmt.Text = Math.Floor(Tot).ToString("0.00");
                    txtDiscount.Text = Math.Floor(Disc).ToString("0.00");
                }
                else
                {
                    txtTotalAmt.Text = Tot.ToString("0.00");
                    txtDueAmt.Text = Tot.ToString("0.00");
                    txtDiscount.Text = Disc.ToString("0.00");
                }


                decimal finalValue = 0;
                decimal.TryParse(txtTotalAmt.Text, out finalValue);

                decimal roundDiff = finalValue - Tot;


                txtRoundOFType.Text = $"{PharamacyRoundOffType} ({roundDiff:0.00})";
                txtRoundOFType.Foreground = new SolidColorBrush(Colors.Red);
            }
        }


        private void UpdateDue()
        {
            // if (txtBillno.Text == "")
            // {
            if (dgvPaymentDet.Items.Count > 0)
            {
                decimal paid = 0;
                foreach (var item in payDet)
                {
                    paid += item.Amount;
                }
                if (PharamacyRoundOffType == "RoundOff")
                {
                    txtPaidAmt.Text = Math.Round(paid, 0).ToString("0.00");
                }
                else if (PharamacyRoundOffType == "Ceil")
                {
                    txtPaidAmt.Text = Math.Ceiling(paid).ToString();
                }
                else if (PharamacyRoundOffType == "Floor")
                {
                    txtPaidAmt.Text = Math.Floor(paid).ToString();
                }
                else
                {
                    txtPaidAmt.Text = paid.ToString();
                }
                txtPayAmount.Text = "0";
                txtDueAmt.Text = (Convert.ToDecimal(txtDueAmt.Text) - Convert.ToDecimal(txtPaidAmt.Text)).ToString();
            }
            // }
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
            //if (!string.IsNullOrEmpty(txtConcPer.Text) && Convert.ToDecimal(txtConcPer.Text) != 0)
            //{
            var collection = dgvPaymentDet.ItemsSource as ObservableCollection<AddPayDet>;
            collection?.Clear();
            //  dgvPaymentDet.ItemsSource = null;


            cmbDiscType.SelectedIndex = 0;
            txtConcPer.Text = "0";
            txtConcAmt.Text = "0";
            txtDiscReason.Text = "";
            txtPayAmount.Text = "0";
            txtDiscount.Text = "0";
            txtPaidAmt.Text = "0";
            cmbDueAuth.SelectedIndex = 0;
            txtDueReason.Text = "";
            cmbDiscAuth.SelectedIndex = 0;
            txtDiscReason.Text = "";

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
            // }

        }
        #endregion

        #region cmbDiscType_SelectionChanged
        private void cmbDiscType_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (cmbDiscType.SelectedIndex == 1)
            {
                txtConcPer.IsEnabled = true;
                cmbDiscAuth.IsEnabled = true;
                txtDiscReason.IsEnabled = true;
                txtConcAmt.IsEnabled = false;
                txtConcPer.Text = "0";
                txtConcAmt.Text = "0";
                txtDiscount.Text = "0";
                cmbDiscAuth.SelectedIndex = 0;
                txtDiscReason.Text = "0";
                txtPayAmount.Text = "0";
                txtPaidAmt.Text = "0";
                txtDueAmt.Text = txtTotalAmt.Text;
                foreach (var row in dgvItemDetails.Items)
                {
                    if (row == null || row == CollectionView.NewItemPlaceholder)
                        continue;
                    var type = row.GetType();
                    var discPerProp = type.GetProperty("DiscPer");
                    var discountProp = type.GetProperty("Discount");

                    if (discPerProp != null)
                        discPerProp.SetValue(row, 0m, null);

                    if (discountProp != null)
                        discountProp.SetValue(row, 0m, null);
                }

                dgvItemDetails.Items.Refresh();

                txtConcPer.Focus();
            }
            else if (cmbDiscType.SelectedIndex == 2)
            {
                txtConcAmt.IsEnabled = true;
                txtConcAmt.Focus();
                cmbDiscAuth.IsEnabled = true;
                txtDiscReason.IsEnabled = true;
                txtConcPer.IsEnabled = false;
                txtConcPer.Text = "0";
                txtConcAmt.Text = "0";
                txtDiscount.Text = "0";
                cmbDiscAuth.SelectedIndex = 0;
                txtDiscReason.Text = "0";
                txtPayAmount.Text = "0";
                txtPaidAmt.Text = "0";
                txtDueAmt.Text = txtTotalAmt.Text;
                foreach (var row in dgvItemDetails.Items)
                {
                    if (row == null || row == CollectionView.NewItemPlaceholder)
                        continue;

                    var type = row.GetType();
                    var discPerProp = type.GetProperty("DiscPer");
                    var discountProp = type.GetProperty("Discount");

                    if (discPerProp != null)
                        discPerProp.SetValue(row, 0m, null);

                    if (discountProp != null)
                        discountProp.SetValue(row, 0m, null);
                }

                dgvItemDetails.Items.Refresh();

            }
            else
            {
                txtConcPer.IsEnabled = false;
                txtConcAmt.IsEnabled = false;
                cmbDiscAuth.IsEnabled = false;
                txtDiscReason.IsEnabled = false;
                txtConcPer.Text = "0";
                txtConcAmt.Text = "0";
                txtDiscount.Text = "0";
                cmbDiscAuth.SelectedIndex = 0;
                txtDiscReason.Text = "0";
                txtPayAmount.Text = "0";
                txtPaidAmt.Text = "0";
                txtDueAmt.Text = txtTotalAmt.Text;
                foreach (var row in dgvItemDetails.Items)
                {
                    if (row == null || row == CollectionView.NewItemPlaceholder)
                        continue;


                    var type = row.GetType();
                    var discPerProp = type.GetProperty("DiscPer");
                    var discountProp = type.GetProperty("Discount");

                    if (discPerProp != null)
                        discPerProp.SetValue(row, 0m, null);

                    if (discountProp != null)
                        discountProp.SetValue(row, 0m, null);
                }

                dgvItemDetails.Items.Refresh();

            }
        }
        #endregion

        #region txtConcPer_TextChanged
        private void txtConcPer_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (cmbDiscType.SelectedIndex == 1)
            {
                if (txtConcPer.Text != "")
                {

                    if (!string.IsNullOrWhiteSpace(txtConcPer.Text))
                    {

                        CalculateConc();
                    }
                }
                else
                {
                    txtConcPer.Text = "0";

                }
            }
        }
        #endregion

        #region txtConcAmt_TextChanged
        private void txtConcAmt_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (cmbDiscType.SelectedIndex == 2)
            {
                if (txtConcAmt.Text != "")
                {
                    if (!string.IsNullOrWhiteSpace(txtConcAmt.Text))
                    {
                        CalculateConc();
                    }
                }
                else
                {
                    txtConcAmt.Text = "0";
                }
            }
        }
        #endregion

        #region CalculateConc
        private void CalculateConc()
        {

            if (txtDueAmt.Text != "" && txtDueAmt.Text != "0")
            {
                if (txtConcPer.Text != "" && txtConcPer.Text != "0" || txtConcAmt.Text != "" && txtConcAmt.Text != "0")
                {
                    if (cmbDiscType.SelectedIndex == 1)
                    {
                        if (Convert.ToDecimal(txtConcPer.Text) > 100)
                        {
                            MessageBox.Show("You Are Not Allowed To Give Discount More than " + 100 + " % Discount", "Infomation", MessageBoxButton.OK, MessageBoxImage.Information);
                            txtConcPer.Text = "0";
                            txtConcPer.Focus();
                            return;
                        }
                    }

                    if (cmbDiscType.SelectedIndex == 1)
                    {
                        if (PharamacyRoundOffType == "RoundOff")
                        {
                            txtConcAmt.Text = txtDiscount.Text = Math.Round(Convert.ToDecimal(txtTotalAmt.Text) * Convert.ToDecimal(txtConcPer.Text) / 100).ToString();
                        }
                        else if (PharamacyRoundOffType == "Ceil")
                        {
                            txtConcAmt.Text = txtDiscount.Text = Math.Ceiling(Convert.ToDecimal(txtTotalAmt.Text) * Convert.ToDecimal(txtConcPer.Text) / 100).ToString();
                        }
                        else if (PharamacyRoundOffType == "Floor")
                        {
                            txtConcAmt.Text = txtDiscount.Text = Math.Floor(Convert.ToDecimal(txtTotalAmt.Text) * Convert.ToDecimal(txtConcPer.Text) / 100).ToString();
                        }
                        else
                        {
                            txtConcAmt.Text = txtDiscount.Text = (Convert.ToDecimal(txtTotalAmt.Text) * Convert.ToDecimal(txtConcPer.Text) / 100).ToString();
                        }



                        if (gvars.gConcType == "%")
                        {
                            decimal totalAmt = 0;
                            decimal conctAmt = 0;

                            if (PharamacyRoundOffType == "RoundOff")
                            {
                                if (decimal.TryParse(txtTotalAmt.Text, out totalAmt))
                                {
                                    conctAmt = Math.Round((totalAmt * gvars.gConcLimit) / 100);
                                }
                            }
                            else if (PharamacyRoundOffType == "Ceil")
                            {
                                if (decimal.TryParse(txtTotalAmt.Text, out totalAmt))
                                {
                                    conctAmt = Math.Ceiling((totalAmt * gvars.gConcLimit) / 100);
                                }
                            }
                            else if (PharamacyRoundOffType == "Floor")
                            {
                                if (decimal.TryParse(txtTotalAmt.Text, out totalAmt))
                                {
                                    conctAmt = Math.Floor((totalAmt * gvars.gConcLimit) / 100);
                                }
                            }
                            else
                            {
                                if (decimal.TryParse(txtTotalAmt.Text, out totalAmt))
                                {
                                    conctAmt = ((totalAmt * gvars.gConcLimit) / 100);
                                }
                            }

                            if (Convert.ToDecimal(txtConcPer.Text) > gvars.gConcLimit)
                            {
                                MessageBox.Show("You Are Not Allowed To Give Discount More than " + gvars.gConcLimit + " % Discount & " + conctAmt + " Rs Discount", "Infomation", MessageBoxButton.OK, MessageBoxImage.Information);
                                txtConcPer.Text = txtConcAmt.Text = "0";
                                txtConcPer.Focus();
                                return;
                            }
                        }
                        else
                        {
                            decimal totalAmt = 0;
                            decimal conctAmt = 0;
                            if (PharamacyRoundOffType == "RoundOff")
                            {
                                if (decimal.TryParse(txtTotalAmt.Text, out totalAmt))
                                {
                                    conctAmt = Math.Round((totalAmt * gvars.gConcLimit) / 100);
                                }
                            }
                            else if (PharamacyRoundOffType == "Ceil")
                            {
                                if (decimal.TryParse(txtTotalAmt.Text, out totalAmt))
                                {
                                    conctAmt = Math.Ceiling((totalAmt * gvars.gConcLimit) / 100);
                                }

                            }
                            else if (PharamacyRoundOffType == "Floor")
                            {
                                if (decimal.TryParse(txtTotalAmt.Text, out totalAmt))
                                {
                                    conctAmt = Math.Floor((totalAmt * gvars.gConcLimit) / 100);
                                }
                            }
                            else
                            {
                                if (decimal.TryParse(txtTotalAmt.Text, out totalAmt))
                                {
                                    conctAmt = ((totalAmt * gvars.gConcLimit) / 100);
                                }
                            }


                            // txtConctAmt.Text = conctAmt.ToString("0.00");
                            if (Convert.ToDecimal(txtConcAmt.Text) > conctAmt)
                            {
                                MessageBox.Show("You Are Not Allowed To Give Discount More than " + gvars.gConcLimit + " % Discount & " + conctAmt + " Rs Discount", "Infomation", MessageBoxButton.OK, MessageBoxImage.Information);
                                txtConcPer.Text = txtConcAmt.Text = "0";
                                txtConcPer.Focus();
                                return;
                            }
                        }
                        if (PharamacyRoundOffType == "RoundOff")
                        {
                            txtDueAmt.Text = Math.Round((Convert.ToDecimal(txtTotalAmt.Text) - Convert.ToDecimal(txtConcAmt.Text)), 0).ToString();
                        }
                        else if (PharamacyRoundOffType == "Ceil")
                        {
                            txtDueAmt.Text = Math.Ceiling((Convert.ToDecimal(txtTotalAmt.Text) - Convert.ToDecimal(txtConcAmt.Text))).ToString();
                        }
                        else if (PharamacyRoundOffType == "Floor")
                        {
                            txtDueAmt.Text = Math.Floor((Convert.ToDecimal(txtTotalAmt.Text) - Convert.ToDecimal(txtConcAmt.Text))).ToString();
                        }
                        else
                        {
                            txtDueAmt.Text = ((Convert.ToDecimal(txtTotalAmt.Text) - Convert.ToDecimal(txtConcAmt.Text))).ToString();
                        }

                        txtPayAmount.Text = txtDueAmt.Text;
                        decimal dueAmt;
                        if (decimal.TryParse(txtDueAmt.Text.Trim(), out dueAmt))
                        {
                            if (dueAmt > 0)
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
                    else
                    {
                        if (PharamacyRoundOffType == "RoundOff")
                        {
                            txtConcPer.Text = Math.Round((Convert.ToDecimal(txtConcAmt.Text) / Convert.ToDecimal(txtTotalAmt.Text) * 100), 0).ToString();
                        }
                        else if (PharamacyRoundOffType == "Ceil")
                        {
                            txtConcPer.Text = Math.Ceiling((Convert.ToDecimal(txtConcAmt.Text) / Convert.ToDecimal(txtTotalAmt.Text) * 100)).ToString();
                        }
                        else if (PharamacyRoundOffType == "Floor")
                        {
                            txtConcPer.Text = Math.Floor((Convert.ToDecimal(txtConcAmt.Text) / Convert.ToDecimal(txtTotalAmt.Text) * 100)).ToString();
                        }
                        else
                        {
                            txtConcPer.Text = ((Convert.ToDecimal(txtConcAmt.Text) / Convert.ToDecimal(txtTotalAmt.Text) * 100)).ToString();
                        }
                        if (gvars.gConcType == "%")
                        {
                            if (Convert.ToDecimal(txtConcPer.Text) > gvars.gConcLimit)
                            {
                                MessageBox.Show("You Are Not Allowed To Give Discount More than " + gvars.gConcLimit + " % Discount", "Infomation", MessageBoxButton.OK, MessageBoxImage.Information);
                                txtConcPer.Text = txtConcAmt.Text = "0";
                                txtConcAmt.Focus();
                                return;
                            }
                        }
                        else
                        {
                            if (Convert.ToDecimal(txtConcAmt.Text) > gvars.gConcLimit)
                            {
                                MessageBox.Show("You Are Not Allowed To Give Discount More than " + gvars.gConcLimit + " Rs Discount", "Infomation", MessageBoxButton.OK, MessageBoxImage.Information);
                                txtConcPer.Text = txtConcAmt.Text = "0";
                                txtConcAmt.Focus();
                                return;
                            }
                        }
                        txtDiscount.Text = txtConcAmt.Text;
                        if (PharamacyRoundOffType == "RoundOff")
                        {
                            txtDueAmt.Text = Math.Round(Convert.ToDecimal(txtTotalAmt.Text) - Convert.ToDecimal(txtConcAmt.Text)).ToString();
                        }
                        else if (PharamacyRoundOffType == "Ceil")
                        {
                            txtDueAmt.Text = Math.Ceiling(Convert.ToDecimal(txtTotalAmt.Text) - Convert.ToDecimal(txtConcAmt.Text)).ToString();
                        }
                        else if (PharamacyRoundOffType == "Floor")
                        {
                            txtDueAmt.Text = Math.Floor(Convert.ToDecimal(txtTotalAmt.Text) - Convert.ToDecimal(txtConcAmt.Text)).ToString();
                        }
                        else
                        {
                            txtDueAmt.Text = (Convert.ToDecimal(txtTotalAmt.Text) - Convert.ToDecimal(txtConcAmt.Text)).ToString();
                        }
                        txtPayAmount.Text = txtDueAmt.Text;
                        long dueAmt;
                        if (long.TryParse(txtDueAmt.Text, out dueAmt))
                        {
                            if (dueAmt > 0)
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
                                if (PharamacyRoundOffType == "RoundOff")
                                {
                                    item.Discount = Math.Round((item.Amount * Convert.ToDecimal(txtConcPer.Text)) / 100, 0);
                                }
                                else if (PharamacyRoundOffType == "Ceil")
                                {
                                    item.Discount = Math.Ceiling((item.Amount * Convert.ToDecimal(txtConcPer.Text)) / 100);
                                }
                                else if (PharamacyRoundOffType == "Floor")
                                {
                                    item.Discount = Math.Floor((item.Amount * Convert.ToDecimal(txtConcPer.Text)) / 100);
                                }
                                else
                                {
                                    item.Discount = ((item.Amount * Convert.ToDecimal(txtConcPer.Text)) / 100);
                                }
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
                                if (PharamacyRoundOffType == "RoundOff")
                                {
                                    item.Discount = Math.Round((item.Amount * Convert.ToDecimal(txtConcPer.Text)) / 100);
                                    NetAmount = item.Amount - item.Discount;
                                    item.TaxAmount = Math.Round((NetAmount * item.TaxPer) / 100, 0);
                                }
                                else if (PharamacyRoundOffType == "Ceil")
                                {
                                    item.Discount = Math.Ceiling(item.Amount * Convert.ToDecimal(txtConcPer.Text)) / 100;
                                    NetAmount = item.Amount - item.Discount;
                                    item.TaxAmount = Math.Ceiling((NetAmount * item.TaxPer) / 100);
                                }
                                else if (PharamacyRoundOffType == "Floor")
                                {
                                    item.Discount = Math.Floor(item.Amount * Convert.ToDecimal(txtConcPer.Text)) / 100;
                                    NetAmount = item.Amount - item.Discount;
                                    item.TaxAmount = Math.Floor((NetAmount * item.TaxPer) / 100);
                                }
                                else
                                {
                                    item.Discount = (item.Amount * Convert.ToDecimal(txtConcPer.Text)) / 100;
                                    NetAmount = item.Amount - item.Discount;
                                    item.TaxAmount = ((NetAmount * item.TaxPer) / 100);
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

        #region btnPayAdd_Click
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
                UpdateDue();
                txtConcPer.IsEnabled = false;
                txtConcAmt.IsEnabled = false;
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

        #region btnPrint_Click
        private void btnPrint_Click(object sender, RoutedEventArgs e)
        {
            var button = sender as Button;
            DataRowView pData = (DataRowView)button?.DataContext;
            if (pData != null)
            {
                PrintSlip(pData["BillNo"].ToString());
            }
        }

        private void PrintSlip(string strBillNo)
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

        #region SaveValidation

        private bool SaveValidation()
        {
            bool Valid = true;
            //if (txtName.Text == "")
            //{
            //    MessageBox.Show("Enter Name ", "Information", MessageBoxButton.OK, MessageBoxImage.Information);
            //    txtName.Focus();
            //    Valid = false;
            //}
            if (cmbSaleType.SelectedIndex == 1)
            {
                if (lblUHID.Text == "")
                {
                    MessageBox.Show("Enter UHID ", "Information", MessageBoxButton.OK, MessageBoxImage.Information);
                    lblUHID.Focus();
                    Valid = false;
                }
            }
            //if (cmbDoctor.SelectedIndex == 0)
            //{
            //    MessageBox.Show("Select Doctor ", "Information", MessageBoxButton.OK, MessageBoxImage.Information);
            //    cmbDoctor.Focus();
            //    Valid = false;
            //}
            if (dgvItemDetails.Items.Count == 0)
            {
                MessageBox.Show("There is No Items to save ", "Information", MessageBoxButton.OK, MessageBoxImage.Information);
                cmbItemName.Focus();
                Valid = false;
            }
            decimal dueAmt = 0;
            if (decimal.TryParse(txtDueAmt.Text, out dueAmt))
            {
                if (dueAmt > 0)
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
            }
            decimal disAmt = 0;
            if (decimal.TryParse(txtDiscount.Text, out disAmt))
            {
                //    if (Convert.ToDecimal(txtDiscount.Text) > 0)
                //{
                if (disAmt > 0)
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

        #region btnSave_Click
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

                string[] parts = lblUHID.Text.Split('/');

                string uhid = "";
                string ipdNo = "";


                if (parts.Length == 2)
                {
                    uhid = parts[0];
                    ipdNo = parts[1];
                }



                string text = txtRoundOFType.Text.Trim();

                string roundof = "";
                decimal roundamt = 0;
                decimal amt = 0;
                if (text.Contains("(") && text.Contains(")"))
                {
                    int start = text.IndexOf("(") + 1;
                    int end = text.IndexOf(")");

                    roundof = text.Substring(0, start - 1).Trim();
                    string amtStr = text.Substring(start, end - start).Trim();

                    if (decimal.TryParse(amtStr, out amt))
                    {
                        roundamt = Math.Abs(amt);
                    }
                }
                else
                {
                    roundof = text;
                    roundamt = 0;
                }




                SqlParameter[] sqlParamInsert = new SqlParameter[]
                {
                        new SqlParameter("@UHID",uhid),
                        new SqlParameter("@BillNo", BillNo ),
                        new SqlParameter("@DepartmentID",gvars.gDeptID),
                        new SqlParameter("@IPOPNo", ipdNo),
                        new SqlParameter("@DocID", DocIDLabel.Text), //For Text DoctorNameRun.Text
                        new SqlParameter("@OrganisationID", OrgIDLabel.Text),
                        new SqlParameter("@BedID", BedIDLabel.Text),
                        new SqlParameter("@Remarks",  txtDueReason.Text),
                        new SqlParameter("@IsIP", true),
                        new SqlParameter("@TotalCharges", Convert.ToDecimal(txtTotalAmt.Text)),
                        new SqlParameter("@Paid", Convert.ToDecimal(txtPaidAmt.Text)),
                        new SqlParameter("@Discount", Convert.ToDecimal(txtDiscount.Text)),
                        new SqlParameter("@PostDiscount",0),
                        new SqlParameter("@Due", Convert.ToDecimal(txtDueAmt.Text)),
                        new SqlParameter("@DueAuth", cmbDueAuth.SelectedValue),
                        new SqlParameter("@DueReason", txtDueReason.Text),
                        new SqlParameter("@OPrintCont", OPrintCont),
                        new SqlParameter("@DPrintCount", DPrintCount),
                        new SqlParameter("@PatientName", lblUHID.Text),
                        new SqlParameter("@LocationID", gvars.gLocationId),
                        new SqlParameter("@TerminalID",gvars.gTermId),
                        new SqlParameter("@CreateUserID", gvars.gUserID),
                        new SqlParameter("@RoundOffAmt",roundamt),
                        new SqlParameter("@RoundOffType",roundof),
                        new SqlParameter("@ACTIVITY", "Insert")
                };
                i = obj.ExecuteNonQuery("tSales", DataHelper.SqlCmdType.sqlStoredProcedure, sqlParamInsert, sqlTran);


                if (aItem.Count > 0 && aItem != null)
                {
                    decimal CGST = 0, SGST = 0, IGST = 0;
                    foreach (var Dtls in aItem)
                    {
                        CGST = SGST = Dtls.TaxAmount / 2;

                        decimal unitRate = Convert.ToDecimal(Dtls.UnitRate);
                        decimal unitMrp = Convert.ToDecimal(Dtls.UnitMrp);
                        DateTime expiry = Convert.ToDateTime(Dtls.ExpiryDate);

                        decimal stockQty = CommonService.CheckStockQty(Dtls.ItemId, unitRate, unitMrp, Dtls.BatchNo, expiry, gvars.gDeptID, Dtls.SupplierID, gvars.gLocationId);

                        if (Dtls.Qty > stockQty)
                        {
                            MessageBox.Show(
                               $"Your Enter Qty Is More Then Original Stock ",
                               "warning",
                               MessageBoxButton.OK,
                               MessageBoxImage.Error
                           );

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
                            if (Dtls.Qty > 0)
                            {
                                SqlParameter[] sqlParamGrnDtls = new SqlParameter[]
                                {
                                new SqlParameter("@BillNo",BillNo),
                                new SqlParameter("@ItemID", Dtls.ItemId),
                                new SqlParameter("@Qty", Dtls.Qty),
                                new SqlParameter("@UnitRate", Dtls.UnitRate),
                                new SqlParameter("@UnitMrp", Dtls.UnitMrp),
                                new SqlParameter("@BatchNo", Dtls.BatchNo),
                                new SqlParameter("@ExpiryDate",Convert.ToDateTime(Dtls.ExpiryDate).ToString("dd-MMM-yyyy")),
                                new SqlParameter("@TaxPer", Dtls.TaxPer),
                                new SqlParameter("@TaxAmount",(CGST+SGST)),
                                new SqlParameter("@SupplierID", Dtls.SupplierID),
                                new SqlParameter("@PurchBillNo", Dtls.PurchBillNo),
                                new SqlParameter("@Amount", Dtls.Amount),
                                new SqlParameter("@Discount", Dtls.Discount),
                                new SqlParameter("@DiscPer", Dtls.DiscPer),
                                new SqlParameter("@DiscAuth", cmbDiscAuth.SelectedIndex),
                                new SqlParameter("@DiscReason", txtDiscReason.Text),
                                new SqlParameter("@SGSTPer", Dtls.TaxPer/2),
                                new SqlParameter("@SGSTAmt", SGST),
                                new SqlParameter("@CGSTPer", Dtls.TaxPer/2),
                                new SqlParameter("@CGSTAmt", CGST),
                                new SqlParameter("@LocationID", gvars.gLocationId),
                                new SqlParameter("@CreateUserID", gvars.gUserID),
                                new SqlParameter("@TerminalID", gvars.gTermId),
                                new SqlParameter("@ACTIVITY", "Insert")
                                };
                                i = obj.ExecuteNonQuery("tSalesDtls", DataHelper.SqlCmdType.sqlStoredProcedure, sqlParamGrnDtls, sqlTran);
                            }
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
                            new SqlParameter("@Module","IPSales"),
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
                            new SqlParameter("@Module","IPSales"),
                            new SqlParameter("@TransactionType", "Discount"),
                            new SqlParameter("@ReceiptNo",Receipt),
                            new SqlParameter("@PaymentID",PaymentID),
                            new SqlParameter("@Amount", Convert.ToDecimal(txtDiscount.Text)),
                            new SqlParameter("@DiscPerc", Convert.ToDecimal(txtConcPer.Text)),
                            new SqlParameter("@Authorisation", cmbDiscAuth.SelectedValue),
                            new SqlParameter("@Remarks","Discount" ),//txtDiscReason.Text
                            new SqlParameter("@PayMode","Cash" ),//payDet[0].PayMode
                            new SqlParameter("@WalletAccount","0"),// payDet[0].WalletAccount
                            new SqlParameter("@Bank", "0"), //payDet[0].Bank
                            new SqlParameter("@TransNo","0"), //payDet[0].TransNo
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
                                    new SqlParameter("@Module", "IPSales"),
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
                        if (Dtls.Qty > 0)
                        {
                            SqlParameter[] sqlParamGrnDtls = new SqlParameter[]
                            {
                                new SqlParameter("@DepartmentID",gvars.gDeptID),
                                new SqlParameter("@ItemID", Dtls.ItemId),
                                new SqlParameter("@Qty", Dtls.Qty),
                                new SqlParameter("@UnitRate", Dtls.UnitRate),
                                new SqlParameter("@UnitMrp", Dtls.UnitMrp),
                                new SqlParameter("@BatchNo", Dtls.BatchNo),
                                new SqlParameter("@ExpiryDate", Convert.ToDateTime(Dtls.ExpiryDate).ToString("dd-MMM-yyyy")),
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
                }
                if ((cmbSaleType.SelectedIndex == 0) && (i > 0))
                {
                    if (aItem != null)
                    {
                        foreach (var Dtls in aItem)
                        {
                            SqlParameter[] sqlParamGrnDtls = new SqlParameter[]
                        {
                            new SqlParameter("@ItemID", Dtls.ItemId),
                            new SqlParameter("@IssueQty", Dtls.Qty),
                            new SqlParameter("@IndentNo", Dtls.IndentNo),
                            new SqlParameter("@ACTIVITY", "UpdateIndent")
                        };
                            i = obj.ExecuteNonQuery("tIPSales", DataHelper.SqlCmdType.sqlStoredProcedure, sqlParamGrnDtls, sqlTran);
                        }
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
            }
        }
        #endregion

        #region btnView_Click
        private void btnView_Click(object sender, RoutedEventArgs e)
        {
            var button = sender as Button;
            DataRowView pData = (DataRowView)button?.DataContext;
            if (pData != null)
            {
                SqlParameter[] sqlParamSearch = new SqlParameter[]
                {
                new SqlParameter("@ACTIVITY", "GetSingleData"),
                new SqlParameter("@BillNo", pData["BillNo"].ToString()),
                new SqlParameter("@LocationID", gvars.gLocationId),
                new SqlParameter("@DepartmentID", gvars.gDeptID)
                };
                DataSet dsSearch = obj.getDataset("tIPSales", DataHelper.SqlCmdType.sqlStoredProcedure, sqlParamSearch);
                if (dsSearch.Tables[0].Rows.Count > 0 && dsSearch.Tables.Count > 0)
                {

                    //txtUHID.Text = dsSearch.Tables[0].Rows[0]["UHID"].ToString();
                    //if (dsSearch.Tables[0].Rows[0]["IPOPNo"].ToString() != "")
                    //{
                    //    cmbSaleType.SelectedIndex = 1;
                    //    GetPatientDetails();
                    //}
                    //else
                    //{
                    //    cmbSaleType.SelectedIndex = 2;
                    //    if (txtUHIDPhone.Text != "")
                    //    {
                    //        GetPatientDetails();
                    //    }
                    //    else
                    //    {
                    //        txtName.Text = dsSearch.Tables[0].Rows[0]["PatientName"].ToString();
                    //        cmbDoctor.SelectedValue = dsSearch.Tables[0].Rows[0]["DocID"].ToString();
                    //    }
                    //}
                    //txtBillno.Text = dsSearch.Tables[0].Rows[0]["BillNo"].ToString();
                    //dtpBillDt.Text = Convert.ToDateTime(dsSearch.Tables[0].Rows[0]["BillDate"]).ToString("dd-MMM-yyyy");
                    txtTotalAmt.Text = dsSearch.Tables[0].Rows[0]["TotalCharges"].ToString();
                    txtDiscount.Text = dsSearch.Tables[0].Rows[0]["Discount"].ToString();
                    txtPaidAmt.Text = dsSearch.Tables[0].Rows[0]["Paid"].ToString();
                    txtDueAmt.Text = dsSearch.Tables[0].Rows[0]["Due"].ToString();
                    txtConcAmt.Text = dsSearch.Tables[0].Rows[0]["Discount"].ToString();
                    txtDueReason.Text = dsSearch.Tables[0].Rows[0]["DueReason"].ToString();
                    cmbDueAuth.SelectedValue = dsSearch.Tables[0].Rows[0]["DueAuth"].ToString();
                    txtPayAmount.Text = dsSearch.Tables[0].Rows[0]["Paid"].ToString();
                    if (dsSearch.Tables[1].Rows.Count > 0)
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
                    if (!string.IsNullOrEmpty(dsSearch.Tables[0].Rows[0]["IndentNo"].ToString()))
                    {
                        txtIndentPhone.Text = dsSearch.Tables[0].Rows[0]["IndentNo"].ToString();
                        CheckIndentData();
                        txtIndentPhone.Visibility = Visibility.Visible;
                        txtUHIDPhone.Visibility = Visibility.Collapsed;
                    }
                    else
                    {
                        txtUHIDPhone.Text = dsSearch.Tables[0].Rows[0]["UHID"].ToString();
                        GetPatientDetails();
                        txtIndentPhone.Visibility = Visibility.Collapsed;
                        txtUHIDPhone.Visibility = Visibility.Visible;
                    }
                    // btnFind.Visibility = Visibility.Collapsed;
                    // btnReset.Visibility = Visibility.Collapsed;
                    btnFind.Visibility = Visibility.Collapsed;
                    GridPanel.Visibility = Visibility.Collapsed;
                    FormPanel.Visibility = Visibility.Visible;
                    btnNew.Content = "Back";
                }
            }
        }



        #endregion


        #region Launch Calculator btnCalc_Click
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

        #region btnReset_Click
        private void btnReset_Click(object sender, RoutedEventArgs e)
        {
            EnableControls();
            ClearControls();

        }
        #endregion

        #region PayResetDetails
        private void PayResetDetails()
        {
            txtDueAmt.Text = txtTotalAmt.Text;
            txtDiscReason.Text = txtPayAmount.Text = txtRefNo.Text = txtDueReason.Text = "";
            cmbWalletType.SelectedIndex = cmbDueAuth.SelectedIndex = cmbBank.SelectedIndex = -1;
            cmbDiscType.SelectedIndex = cmbPayMode.SelectedIndex = 0;
            payDet.Clear();
            dgvPaymentDet.ItemsSource = null;
            txtConcPer.IsEnabled = true;
            txtConcAmt.IsEnabled = true;
            txtPaidAmt.Text = txtConcPer.Text = txtConcAmt.Text = txtDiscount.Text = "0";
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

            foreach (var row in dgvItemDetails.Items)
            {
                if (row == null || row == CollectionView.NewItemPlaceholder)
                    continue;


                var type = row.GetType();
                var discPerProp = type.GetProperty("DiscPer");
                var discountProp = type.GetProperty("Discount");

                if (discPerProp != null)
                    discPerProp.SetValue(row, 0m, null);

                if (discountProp != null)
                    discountProp.SetValue(row, 0m, null);
            }

            dgvItemDetails.Items.Refresh();

        }
        #endregion

        #region btnPayReset_Click
        private void btnPayReset_Click(object sender, RoutedEventArgs e)
        {
            PayResetDetails();
        }
        #endregion

        #region dgvItemDetails_KeyDown
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

        #region txtPayAmount_TextChanged
        private void txtPayAmount_TextChanged(object sender, TextChangedEventArgs e)
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
        #endregion

        #region txtDueAmt_TextChanged
        private void txtDueAmt_TextChanged(object sender, TextChangedEventArgs e)
        {
            //if (txtBillno.Text == "")
            // {
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
            // }
        }
        #endregion

        #region cmbItemName_Loaded
        private TextBox cmbTextBox; // cache inner TextBox

        private void cmbItemName_Loaded(object sender, RoutedEventArgs e)
        {
            cmbTextBox = cmbItemName.Template.FindName("PART_EditableTextBox", cmbItemName) as TextBox;

            // Load data from DB
            Task.Run(() =>
            {
                var items = LoadItemsFromDatabase(); // Your DB fetch
                Dispatcher.Invoke(() =>
                {
                    allItems = items; // Keep in global list for filtering
                    cmbItemName.ItemsSource = allItems;
                    cmbItemName.DisplayMemberPath = "ItemName";
                    cmbItemName.SelectedValuePath = "ItemID";
                });
            });
        }


        private void cmbItemName_KeyUp(object sender, KeyEventArgs e)
        {
            string text = cmbItemName.Text;

            // Ignore navigation keys
            if (e.Key == Key.Up || e.Key == Key.Down || e.Key == Key.Enter || e.Key == Key.Tab)
                return;

            // Filter items
            var filtered = allItems
                .Where(x => !string.IsNullOrEmpty(x.ItemName) &&
                            x.ItemName.IndexOf(text, StringComparison.OrdinalIgnoreCase) >= 0)
                .ToList();

            // Rebind filtered items
            cmbItemName.ItemsSource = filtered;
            cmbItemName.IsDropDownOpen = true;

            // Now use Dispatcher to restore text after ItemsSource refresh
            Dispatcher.BeginInvoke(new Action(() =>
            {
                cmbItemName.Text = text; // restore what user typed

                if (cmbTextBox != null)
                {
                    cmbTextBox.Focus();
                    cmbTextBox.SelectionStart = cmbTextBox.Text.Length;
                    cmbTextBox.SelectionLength = 0;
                }

            }), System.Windows.Threading.DispatcherPriority.Background);
        }




        private void cmbItemName_TextChanged(object sender, TextChangedEventArgs e)
        {
            var textBox = cmbItemName.Template.FindName("PART_EditableTextBox", cmbItemName) as TextBox;
            if (textBox == null) return;

            string input = textBox.Text;

            if (string.IsNullOrWhiteSpace(input))
            {
                cmbItemName.ItemsSource = allItems;
            }
            else
            {
                var filtered = allItems
                    .Where(x => !string.IsNullOrEmpty(x.ItemName) &&
                                x.ItemName.StartsWith(input, StringComparison.OrdinalIgnoreCase))
                    .ToList();

                cmbItemName.ItemsSource = filtered;
            }

            cmbItemName.IsDropDownOpen = true;

            // Important: Preserve caret and text
            textBox.SelectionStart = textBox.Text.Length;
        }
        #endregion

        #region cmbItemName_PreviewTextInput
        private void cmbItemName_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            string input = cmbItemName.Text + e.Text;
            var filtered = allItems
                    .Where(x => !string.IsNullOrEmpty(x.ItemName) &&
                                x.ItemName.StartsWith(input, StringComparison.OrdinalIgnoreCase))
                  .ToList();
            cmbItemName.ItemsSource = filtered;
            cmbItemName.IsDropDownOpen = true;
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

                //cmbDoctor.SelectedValuePath = "DocId";
                //cmbDoctor.DisplayMemberPath = "DocName";
                //cmbDoctor.ItemsSource = dsLoc.Tables[0].DefaultView;
                //cmbDoctor.SelectedIndex = 0;
            }
        }
        #endregion

        #region BindPaymentMode
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
                cmbPayMode.SelectedIndex = 0;
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

        //#region BindItem
        //public void BindItem()
        //{
        //    SqlParameter[] sqlParamLoc = new SqlParameter[]
        //    {
        //       new SqlParameter("@IsMedical",true),
        //       new SqlParameter("@ACTIVITY", "GetItem")
        //    };
        //    DataSet dsLoc = obj.getDataset("tSales", DataHelper.SqlCmdType.sqlStoredProcedure, sqlParamLoc);
        //    if (dsLoc.Tables.Count > 0 && dsLoc.Tables[0].Rows.Count > 0)
        //    {
        //        cmbItemName.SelectedValuePath = "ItemID";
        //        cmbItemName.DisplayMemberPath = "ItemName";
        //        cmbItemName.ItemsSource = dsLoc.Tables[0].DefaultView;
        //        cmbItemName.SelectedIndex = 0;
        //    }
        //}
        //#endregion

        #region BindItem
        public void BindItem()
        {
            SqlParameter[] sqlParamLoc = new SqlParameter[]
            {
                new SqlParameter("@IsMedical",true),
                new SqlParameter("@ACTIVITY", "GetItem")
            };
            ds = obj.getDataset("tSales", DataHelper.SqlCmdType.sqlStoredProcedure, sqlParamLoc);
            if (ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
            {
                DataRow drow;
                drow = ds.Tables[0].NewRow();
                drow["ItemID"] = "--Select--";
                drow["ItemName"] = "--Select--";
                ds.Tables[0].Rows.InsertAt(drow, 0);
                //cmbItemName.SelectedValuePath = "ItemID";
                //cmbItemName.DisplayMemberPath = "ItemName";
                //cmbItemName.ItemsSource = ds.Tables[0].DefaultView;  
                //cmbItemName.SelectedIndex = 0;
                //ItemCB = new AutoCompleteCombobox(ds.Tables[0], "ItemName"); 
                //DataContext = ItemCB;

                allItems = ds.Tables[0]
        .AsEnumerable()
        .Select(row => new Item
        {
            ItemID = row["ItemID"].ToString(),
            ItemName = row["ItemName"].ToString()
        })
        .ToList();

                //cmbItemName.ItemsSource = allItems;
                //comboView = CollectionViewSource.GetDefaultView(allItems);
                //cmbItemName.ItemsSource = comboView;
                cmbItemName.DisplayMemberPath = "ItemName";
                cmbItemName.SelectedValuePath = "ItemID";
                cmbItemName.SelectedIndex = 0;

            }
            else
            {
                cmbItemName.ItemsSource = null;
            }
        }

        private List<Item> LoadItemsFromDatabase()
        {
            List<Item> itemList = new List<Item>();
            SqlParameter[] sqlParamLoc = new SqlParameter[]
            {
                new SqlParameter("@IsMedical",true),
                new SqlParameter("@ACTIVITY", "GetItem")
            };
            ds = obj.getDataset("tSales", DataHelper.SqlCmdType.sqlStoredProcedure, sqlParamLoc);
            if (ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
            {
                itemList = ConvertDataTableToObjects<Item>(ds.Tables[0]);
                allItems = ConvertDataTableToObjects<Item>(ds.Tables[0]);
            }
            Item defaultItem = new Item { ItemID = "0", ItemName = "--Select--" };
            itemList.Insert(0, defaultItem);
            allItems.Insert(0, defaultItem);

            return itemList;

        }
        #endregion

       

       

       

       

        #region AddPayDet
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
        #endregion

        #region Item
        public class Item
        {
            public string ItemID { get; set; }
            public string ItemName { get; set; }
            public int Qty { get; set; }
            public int BatchQty { get; set; }
            public decimal Amount { get; set; }
            public double DiscPer { get; set; }
            public decimal Discount { get; set; }
        }
        #endregion

        public static List<T> ConvertDataTableToObjects<T>(DataTable dataTable) where T : new()
        {
            List<T> list = new List<T>();

            foreach (DataRow row in dataTable.Rows)
            {
                T item = new T();
                foreach (DataColumn col in dataTable.Columns)
                {
                    var property = typeof(T).GetProperty(col.ColumnName);
                    if (property != null)
                    {
                        property.SetValue(item, row[col.ColumnName] is DBNull ? null : row[col.ColumnName]);
                    }
                }
                list.Add(item);
            }
            return list;
        }

        private T FindParent<T>(DependencyObject child) where T : DependencyObject
        {
            DependencyObject parentObject = VisualTreeHelper.GetParent(child);
            if (parentObject == null) return null;

            T parent = parentObject as T;
            return parent ?? FindParent<T>(parentObject);
        }



     }
}
