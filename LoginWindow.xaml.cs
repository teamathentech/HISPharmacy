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
using System.Windows.Input;

namespace HISPharmacy
{
    /// <summary>
    /// Interaction logic for LoginWindow.xaml
    /// </summary>
    public partial class LoginWindow : Window
    {
        string strQry;
        DataHelper obj = new DataHelper();
        GlobalVariables gbVars = new GlobalVariables();
        DataSet ds;

        public LoginWindow()
        {
            InitializeComponent();
            txtUserName.Focus();
        }

        #region txtUserName_TextChanged
        private void txtUserName_TextChanged(object sender, TextChangedEventArgs e)
        {
            strQry = " SELECT DISTINCT  U.UserID,UL.LocationID,L.LocationName from mstUser U LEFT JOIN mstUserLocations UL ON UL.LocationID=U.LocationID " +
               " INNER JOIN mstLocation L ON L.LocationID=U.LocationID WHERE U.UserID='" + txtUserName.Text + "'";
            ds = obj.getDataset(strQry, DataHelper.SqlCmdType.sqlText);
            if (ds.Tables[0].Rows.Count > 0)
            {
                cmbLocation.SelectedValuePath = "LocationID";
                cmbLocation.DisplayMemberPath = "LocationName";
                cmbLocation.ItemsSource = ds.Tables[0].DefaultView;
                cmbLocation.SelectedIndex = 0;
            }
            strQry = " SELECT UD.DepartmentID, DepartmentName FROM  mstUserDepartments UD INNER JOIN mstDepartment D ON D.DepartmentID = UD.DepartmentID " +
                    " WHERE UD.UserID = '" + txtUserName.Text + "'";
            ds = obj.getDataset(strQry, DataHelper.SqlCmdType.sqlText);
            if (ds.Tables[0].Rows.Count > 0)
            {
                cmbDepartment.SelectedValuePath = "DepartmentID";
                cmbDepartment.DisplayMemberPath = "DepartmentName";
                cmbDepartment.ItemsSource = ds.Tables[0].DefaultView;
                cmbDepartment.SelectedIndex = 0;
            }
        }
        #endregion

        #region btnLogin_Click
        private void btnLogin_Click(object sender, RoutedEventArgs e)
        {
            if (txtUserName.Text=="")
            {
                MessageBox.Show("Enter UserId", "Information", MessageBoxButton.OK, MessageBoxImage.Information);
                txtUserName.Focus();
                return;
            }
            if (txtPassword.Password == "")
            {
                MessageBox.Show("Enter Password", "Information", MessageBoxButton.OK, MessageBoxImage.Information);
                txtPassword.Focus();
                return;
            }
            string encryptedPWD = Helper.Encrypt(txtPassword.Password);
            strQry = " SELECT U.*,LOC.LookupText As ConcType,LOD.LookupText AS DueType from mstUser U left join mstLookupTable LOC on U.ConcessionType=LOC.LookupId " +
            " left join mstLookupTable LOD on U.DueType = LOD.LookupId WHERE U.UserID='" + txtUserName.Text + "' AND Password='" + encryptedPWD + "' ";
            ds = obj.getDataset(strQry, DataHelper.SqlCmdType.sqlText);
            if (ds.Tables[0].Rows.Count > 0)
            {
                //CheckPasswordExpiry();
                gbVars.gUID = ds.Tables[0].Rows[0]["ID"].ToString();
                gbVars.gUserID = ds.Tables[0].Rows[0]["UserID"].ToString();
                gbVars.gUserName = ds.Tables[0].Rows[0]["EmployeeName"].ToString();
                gbVars.gUserCode = ds.Tables[0].Rows[0]["UserCode"].ToString();
                gbVars.gLocationId = cmbLocation.SelectedValue.ToString();
                gbVars.gDeptID = cmbDepartment.SelectedValue.ToString(); 
                gbVars.gISSUPERUSER = ds.Tables[0].Rows[0]["IsSuperUser"].ToString();
                gbVars.gRoleId = ds.Tables[0].Rows[0]["RoleID"].ToString();
                gbVars.gConcLimit = Convert.ToDecimal(ds.Tables[0].Rows[0]["Concession"].ToString());
                gbVars.gDueLimit = Convert.ToDecimal(ds.Tables[0].Rows[0]["Due"].ToString());
                gbVars.gConcType = ds.Tables[0].Rows[0]["ConcType"].ToString();
                gbVars.gDueType = ds.Tables[0].Rows[0]["DueType1"].ToString();
                gbVars.gLocName = cmbLocation.Text.ToString();
                gbVars.gDeptName = cmbDepartment.Text.ToString();
                gbVars.gTermId = (Environment.MachineName).ToString();

                MainWindow main = new MainWindow(); // NavigationWindow 
                main.gVars = gbVars;
                main.Show();
                this.Close();
            }
            else
            {
                MessageBox.Show("Invalid UserId/Password", "Information", MessageBoxButton.OK,MessageBoxImage.Information);
                txtUserName.Focus();
                return;
            }

        }
        #endregion

        private void CheckPasswordExpiry()
        {
            SqlParameter[] sqlParamSearch = new SqlParameter[]
                           {
                        new SqlParameter("@UserID",gbVars.gUserID),
                        new SqlParameter("@ACTIVITY", "CheckChangePasword")
                           };
            DataSet dsSearch = obj.getDataset("tLoginHistory", DataHelper.SqlCmdType.sqlStoredProcedure, sqlParamSearch);
            if (dsSearch.Tables[0].Rows.Count != 0)
            {
                var row = dsSearch.Tables[0].Rows[0];

                string changePasswordDateString = row["ChangePasswordDate"].ToString();
                int passwordExpiry = Convert.ToInt32(row["PasswordExpiry"]);

                string sqlQuery = "SELECT CONVERT(date, GETDATE()) AS CurrentDate";
                object currentDateObj = obj.ExecuteScalar(sqlQuery, DataHelper.SqlCmdType.sqlText);

                DateTime currentDate = Convert.ToDateTime(currentDateObj);

                string formattedDate = currentDate.ToString("yyyy-MM-dd");
                Console.WriteLine(formattedDate);
                DateTime changePasswordDate = DateTime.Parse(changePasswordDateString);


                TimeSpan difference = DateTime.Parse(formattedDate).Date - changePasswordDate.Date;
                //----------------------For Login Details
                DateTime passwordExpiryDate = changePasswordDate.AddDays(passwordExpiry);
                int daysRemaining = (passwordExpiryDate - currentDate).Days;
 
                //--------------------------------------------

                if (difference.TotalDays > passwordExpiry)
                {
                    MessageBox.Show( "Your password has expired. For security reasons, you are required to change your password before you can log in again.","Information",MessageBoxButton.OK,MessageBoxImage.Information);


                  //  return JsonConvert.SerializeObject(new { message = response.Message, url = "/LoginChangePassword.aspx?UID=" + Model.UserID, sessionId = HttpContext.Current.Session.SessionID });
                }
                else
                {
                    //var sessionId = HttpContext.Current.Session.SessionID;
                    //string LoginHistoryId = Convert.ToString(objDataHelper1.ExecuteScalar("SELECT 'LOG' + CONVERT(VARCHAR,ISNULL((MAX(CONVERT(INT,(SUBSTRING(LoginHistoryId ,(3 + 1) ,len( LoginHistoryId ))))) + 1),1)) AS LoginHistoryId FROM trnLoginHistory", DataHelper.SqlCmdtype.sqlText));

                    //SqlParameter[] sqlParamInsert = new SqlParameter[]
                    //{
                    //                    new SqlParameter("@LoginHistoryId", LoginHistoryId),
                    //                    new SqlParameter("@UserId", Model.UserID),
                    //                    new SqlParameter("@SessionId",sessionId),
                    //                    new SqlParameter("@Browser", Model.Browser),
                    //                    new SqlParameter("@OS", Model.OS),
                    //                    new SqlParameter("@LocationId", location.LocationID),
                    //                    new SqlParameter("@CreateUserID",data.UserID),
                    //                    new SqlParameter("@CreateTerminalID",data.CreateTerminalID),
                    //                    new SqlParameter("@ACTIVITY", "InsertLoginHistory")
                    //};
                    //i = objDataHelper1.ExecuteNonQuery("tLoginHistory", DataHelper.SqlCmdtype.sqlStoredProcedure, sqlParamInsert);

                    //if (i > 0)
                    //{

                    //    return JsonConvert.SerializeObject(new { url = "/Home.aspx", sessionId = HttpContext.Current.Session.SessionID });
                    //}
                    //else
                    //{

                    //    return JsonConvert.SerializeObject(new { url = "/Login.aspx", sessionId = string.Empty });
                    //}


                }

            }
            else
            {
               // return JsonConvert.SerializeObject(new { message = response.Message, url = "/LoginChangePassword.aspx?UID=" + Model.UserID, sessionId = HttpContext.Current.Session.SessionID });
            }
        }
        private void Window_PreviewKeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                // If focus is on Button, do NOT move to next control
                if (Keyboard.FocusedElement is Button)
                {
                    return;
                }

                e.Handled = true;

                UIElement element = Keyboard.FocusedElement as UIElement;

                if (element != null)
                {
                    element.MoveFocus(
                        new TraversalRequest(FocusNavigationDirection.Next)
                    );
                }
            }
        }
    }
}
