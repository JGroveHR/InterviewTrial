using JG_Prospect.BLL;
using JG_Prospect.Common.modal;
//using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using Telerik.Web.UI;

namespace JG_Prospect.Sr_App
{
    public partial class GoogleCalendarView : System.Web.UI.Page
    {
        string strcon = ConfigurationManager.ConnectionStrings["JGPA"].ConnectionString;
        SqlConnection con = null;
        SqlCommand cmd = null;
        SqlDataAdapter da = null;
        DataSet ds = null;
        static string query = "";
        static string Admin = "Admin", usertType = "";
        protected void Page_Load(object sender, EventArgs e)
        {
            RadWindow2.VisibleOnPageLoad = false;
            if (!IsPostBack)
            {
                //Hide Insert,Edit Delete.....
                rsAppointments.AllowInsert = false;
                rsAppointments.AllowEdit = false;
                rsAppointments.AllowDelete = false;
                

                // BindGoogleMap();
                if (Session["usertype"] != null)
                {
                    usertType = Convert.ToString(Session["usertype"]);
                  
                    if (usertType == Admin)
                    {
                        btnAddEvent.Visible = true;
                        A4.Visible = false;
                        //  Response.Redirect("/home.aspx");
                    }
                    else if (Session["loginid"] != null)
                    {

                    }
                }
                Session["AppType"] = "SrApp";
                BindCalendar();
            }
            
        }

        public void BindCalendar()
        {
            DataSet ds = AdminBLL.Instance.GetAllAnnualEvent();
            if (ds.Tables[0].Rows.Count > 0)
            {
                rsAppointments.DataSource = ds.Tables[0];
                rsAppointments.DataBind();
            }
            RadWindow2.VisibleOnPageLoad = false;
        }
        public void BindHRCalendar()
        {
            DataSet ds = AdminBLL.Instance.GetHRCalendar();
            if (ds.Tables[0].Rows.Count > 0)
            {
                rsAppointments.DataSource = ds.Tables[0];
                rsAppointments.DataBind();
            }
            RadWindow2.VisibleOnPageLoad = false;
        }
        public void BindCompanyCalendar()
        {
            if (usertType == Admin)
            {
                DataSet ds = AdminBLL.Instance.GetCompanyCalendar();
                if (ds.Tables[0].Rows.Count > 0)
                {
                    rsAppointments.DataSource = ds.Tables[0];
                    rsAppointments.DataBind();
                }
            }
            RadWindow2.VisibleOnPageLoad = false;
        }
        public void BindEventCalendar()
        {
            DataSet ds = AdminBLL.Instance.GetEventCalendar();
            if (ds.Tables[0].Rows.Count > 0)
            {
                rsAppointments.DataSource = ds.Tables[0];
                rsAppointments.DataBind();
            }
            RadWindow2.VisibleOnPageLoad = false;
        }
        public void BindHRCompanyCalendar()
        {
            DataSet ds = AdminBLL.Instance.GetHRCompanyCalendar();
            if (ds.Tables[0].Rows.Count > 0)
            {
                rsAppointments.DataSource = ds.Tables[0];
                rsAppointments.DataBind();
            }
            RadWindow2.VisibleOnPageLoad = false;
        }

        public void BindHREventCalendar()
        {
            DataSet ds = AdminBLL.Instance.GetHRCompanyEventCalendar();
            if (ds.Tables[0].Rows.Count > 0)
            {
                rsAppointments.DataSource = ds.Tables[0];
                rsAppointments.DataBind();
            }
            RadWindow2.VisibleOnPageLoad = false;
        }
        public void BindCompanyEventCalendar()
        {
            DataSet ds = AdminBLL.Instance.GetEventCompanyCalendar();
            if (ds.Tables[0].Rows.Count > 0)
            {
                rsAppointments.DataSource = ds.Tables[0];
                rsAppointments.DataBind();
            }
            RadWindow2.VisibleOnPageLoad = false;
        }
        public void BindHRCompanyEventCalendar()
        {
            DataSet ds = AdminBLL.Instance.GetHRCompanyEventCalendar();
            if (ds.Tables[0].Rows.Count > 0)
            {
                rsAppointments.DataSource = ds.Tables[0];
                rsAppointments.DataBind();
            }
            RadWindow2.VisibleOnPageLoad = false;
        }
        protected void lbtCustomerID_Click(object sender, EventArgs e)
        {
            //Redirect to customer profile page....
            ScriptManager.RegisterStartupScript(Page, GetType(), "script1", "YetToDeveloped();", true);
            // Response.Redirect(lbtCustomerID.Text);
        }
        //Update Annual Event...............
        protected void btnsave_Click(object sender, EventArgs e)
        {
            AnnualEvent a = new AnnualEvent();
            a.EventName = txtEventName.Text;
            a.Eventdate = txtHolidayDate.Text;
            a.id =Convert.ToInt32(lbtCustomerID.Text);
            new_customerBLL.Instance.UpdateAnnualEvent(a);
            BindCalendar();
           
            int userId = Convert.ToInt16(Session[JG_Prospect.Common.SessionKey.Key.UserId.ToString()]);
            ScriptManager.RegisterStartupScript(this.Page, GetType(), "al", "alert('Event Updated Successfully');", true);
            RadWindow2.VisibleOnPageLoad = false;
/*
            try
            { 
                //Adding Record to Database through Stored Procedure
                con = new SqlConnection(strcon);
                cmd = new SqlCommand("UpdateAnnualEvent", con);
                cmd.CommandType = CommandType.StoredProcedure;


                cmd.Parameters.AddWithValue("@Eventname", txtEventName.Text);
                cmd.Parameters.AddWithValue("@EventDate", txtHolidayDate.Text);
                cmd.Parameters.AddWithValue("@ID", lbtCustomerID.Text);
                
                con.Open();
                cmd.ExecuteNonQuery();
                con.Close();

                BindCalendar();
                //Clear All Data after submitting.....               
                ScriptManager.RegisterStartupScript(this.Page, GetType(), "al", "alert('Event Updated Successfully');", true);
                RadWindow2.VisibleOnPageLoad = false;
            }

            catch
            {
                //return 0;
                //LogManager.Instance.WriteToFlatFile(ex);
            }
           */

        }

        protected void btnClose_Click(object sender, EventArgs e)
        {
            RadWindow2.VisibleOnPageLoad = false;
        }

        protected void rsAppointments_AppointmentClick(object sender, SchedulerEventArgs e)
        {
            if (usertType == Admin)
            {
                con = new SqlConnection(strcon);
                int ID = Convert.ToInt32(e.Appointment.ID);
                string Even = e.Appointment.Subject;
                string[] str = Even.Split(' ');
                string strResult = str[0];
                ViewState["ID"] = ID;
                string year = Convert.ToString(System.DateTime.Now.Year);

                if (strResult != "InterViewDetails")
                {
                    lblDesigna.Visible = false;
                    lblApplicant.Visible = false;
                    lblPhone.Visible = false;
                    lblDesigna.Visible = false;
                    lblAdded.Visible = false;
                    lblAplicantfirstName.Visible = false;
                    lblPhoneNo.Visible = false;
                    lblPhoneNo.Visible = false;
                    lblDesignation.Visible = false;
                    lblAddedBy.Visible = false;
                    lbtLastName.Visible = false;
                    Label2.Visible = true;
                    txtEventName.Visible = true;
                    Label2.Visible = true;
                    txtEventName.Visible = true;
                    Label3.Visible = true;
                    txtHolidayDate.Visible = true;
                    btnsave.Visible = true;
                    btnDelete.Visible = true;

                    //string query = "select * from tbl_AnnualEvents where DATEPART(yyyy,EventDate)='" + year + "'";// where id='" + ID + "'";
                    string query = "select * from tbl_AnnualEvents where id='" + ID + "'";
                    da = new SqlDataAdapter(query, con);
                    ds = new DataSet();
                    da.Fill(ds);
                    if (ds.Tables[0].Rows.Count > 0)
                    {
                        lbtCustomerID.Text = Convert.ToString(ds.Tables[0].Rows[0]["ID"]);
                        txtEventName.Text = Convert.ToString(ds.Tables[0].Rows[0]["EventName"]);
                        string EventDate = Convert.ToString(ds.Tables[0].Rows[0]["EventDate"]);
                        DateTime dat = Convert.ToDateTime(EventDate);
                        txtHolidayDate.Text = dat.ToString("MM/dd/yyyy"); ;
                        RadWindow2.VisibleOnPageLoad = true;
                    }
                }
                else
                {
                    lblAplicantfirstName.Visible = true;
                    lblPhoneNo.Visible = true;
                    lblPhoneNo.Visible = true;
                    lblDesignation.Visible = true;
                    lblAddedBy.Visible = true;
                    lbtLastName.Visible = true;
                    Label2.Visible = false;
                    txtEventName.Visible = false;
                    Label2.Visible = false;
                    txtEventName.Visible = false;
                    Label3.Visible = false;
                    txtHolidayDate.Visible = false;
                    lblDesigna.Visible = true;
                    lblApplicant.Visible = true;
                    lblPhone.Visible = true;
                    lblDesigna.Visible = true;
                    lblAdded.Visible = true;
                    btnsave.Visible = false;
                    btnDelete.Visible = false;

                    int id = Convert.ToInt32(ViewState["ID"]);
                    DataSet ds = AdminBLL.Instance.GetInterviewDetails(id);
                    int applicantId = 0;
                    if (ds.Tables[0].Rows.Count > 0)
                    {
                        applicantId = Convert.ToInt32(ds.Tables[0].Rows[0]["ApplicantId"]);
                        ViewState["ApplicantId"] = applicantId;
                        lbtCustomerID.Text = Convert.ToString(ds.Tables[0].Rows[0]["ID"]);                      
                        lblAplicantfirstName.Text = Convert.ToString(ds.Tables[0].Rows[0]["FristName"]);
                        lblPhoneNo.Text = Convert.ToString(ds.Tables[0].Rows[0]["Phone"]);
                        lblDesignation.Text = Convert.ToString(ds.Tables[0].Rows[0]["Designation"]);
                        int a = Convert.ToInt32(ds.Tables[0].Rows[0]["EventAddedBy"]);
                        string query = "select * from tblUsers where id='" + a + "'";
                        da = new SqlDataAdapter(query, con);
                        DataSet dsid = new DataSet();
                        da.Fill(dsid);

                        if (dsid.Tables[0].Rows.Count > 0)
                        {
                            lblAddedBy.Text = Convert.ToString(dsid.Tables[0].Rows[0]["Username"]);
                        }
                        lbtLastName.Text = Convert.ToString(ds.Tables[0].Rows[0]["LastName"]);
                        RadWindow2.VisibleOnPageLoad = true;
                    }
                }
            }
          
           else
           {
               RadWindow2.VisibleOnPageLoad = false;
           }
        }

        protected void btnDelete_Click(object sender, EventArgs e)
        {
            AnnualEvent a = new AnnualEvent();            
            a.id = Convert.ToInt32(lbtCustomerID.Text);
            new_customerBLL.Instance.DeleteAnnualEvent(a);
            BindCalendar();
           
            int userId = Convert.ToInt16(Session[JG_Prospect.Common.SessionKey.Key.UserId.ToString()]);
            ScriptManager.RegisterStartupScript(this.Page, GetType(), "al", "alert('Event Deleted Successfully');", true);
            RadWindow2.VisibleOnPageLoad = false;
            /*
            try
            {
                //Adding Record to Database through Stored Procedure
                con = new SqlConnection(strcon);
                cmd = new SqlCommand("DeleteAnnualEvent", con);
                cmd.CommandType = CommandType.StoredProcedure;

                cmd.Parameters.AddWithValue("@ID", lbtCustomerID.Text);

                con.Open();
                cmd.ExecuteNonQuery();
                con.Close();
                BindCalendar();

                //Clear All Data after submitting.....
                ScriptManager.RegisterStartupScript(this.Page, GetType(), "al", "alert('Event Deleted Successfully');", true);
                RadWindow2.VisibleOnPageLoad = false;
            }

            catch
            {
                //return 0;
                //LogManager.Instance.WriteToFlatFile(ex);
            }
           */
        }

        protected void btnAddEvent_Click(object sender, EventArgs e)
        {
            Response.Redirect("AddEvents.aspx");
        }

        protected void lbtLastName_Click(object sender, EventArgs e)
        {
            Response.Redirect("InstallCreateUser.aspx?ID=" + ViewState["ApplicantId"]);
        }

        protected void chkHR_CheckedChanged(object sender, EventArgs e)
        {

            //For HR Calendar.....
            if (chkHR.Checked == true && chkCompany.Checked == false && chkEvents.Checked == false)
            {
                BindHRCalendar();
            }
            //For Company Calendar.....
            else if (chkCompany.Checked == true && chkHR.Checked == false && chkEvents.Checked == false)
            {
                BindCompanyCalendar();
            }
            //For Event Calendar.....
            else if (chkEvents.Checked == true && chkHR.Checked == false && chkCompany.Checked == false)
            {
                BindEventCalendar();
            }
            //For HR Calendar AND Company Calendar.....
            else if (chkHR.Checked == true && chkCompany.Checked == true && chkEvents.Checked == false)
            {
                BindHRCompanyCalendar();
            }
            //For HR Calendar AND Event Calendar.....
            else if (chkHR.Checked == true && chkEvents.Checked == true && chkCompany.Checked == false)
            {
                BindHREventCalendar();
            }
            //For Company Calendar AND Event Calendar.....
            else if (chkCompany.Checked == true && chkEvents.Checked == true && chkHR.Checked == false)
            {
                BindCompanyEventCalendar();
            }
            //For Company Calendar AND Event Calendar AND HR Calendar.....
            else if (chkHR.Checked == true && chkCompany.Checked == true && chkEvents.Checked == true)
            {
                BindHRCompanyEventCalendar();
            }
            else if (chkHR.Checked == false && chkCompany.Checked == false && chkEvents.Checked == false)
            {
                BindCalendar();
            }
        }

        protected void chkCompany_CheckedChanged(object sender, EventArgs e)
        {

            //For HR Calendar.....
            if (chkHR.Checked == true && chkCompany.Checked == false && chkEvents.Checked == false)
            {
                BindHRCalendar();
            }
            //For Company Calendar.....
            else if (chkCompany.Checked == true && chkHR.Checked == false && chkEvents.Checked == false)
            {
                BindCompanyCalendar();
            }
            //For Event Calendar.....
            else if (chkEvents.Checked == true && chkHR.Checked == false && chkCompany.Checked == false)
            {
                BindEventCalendar();
            }
            //For HR Calendar AND Company Calendar.....
            else if (chkHR.Checked == true && chkCompany.Checked == true && chkEvents.Checked == false)
            {
                BindHRCompanyCalendar();
            }
            //For HR Calendar AND Event Calendar.....
            else if (chkHR.Checked == true && chkEvents.Checked == true && chkCompany.Checked == false)
            {
                BindHREventCalendar();
            }
            //For Company Calendar AND Event Calendar.....
            else if (chkCompany.Checked == true && chkEvents.Checked == true && chkHR.Checked == false)
            {
                BindCompanyEventCalendar();
            }
            //For Company Calendar AND Event Calendar AND HR Calendar.....
            else if (chkHR.Checked == true && chkCompany.Checked == true && chkEvents.Checked == true)
            {
                BindHRCompanyEventCalendar();
            }
            else if (chkHR.Checked == false && chkCompany.Checked == false && chkEvents.Checked == false)
            {
                BindCalendar();
            }
        }

        protected void chkEvents_CheckedChanged(object sender, EventArgs e)
        {

            //For HR Calendar.....
            if (chkHR.Checked == true && chkCompany.Checked == false && chkEvents.Checked == false)
            {
                BindHRCalendar();
            }
            //For Company Calendar.....
            else if (chkCompany.Checked == true && chkHR.Checked == false && chkEvents.Checked == false)
            {
                BindCompanyCalendar();
            }
            //For Event Calendar.....
            else if (chkEvents.Checked == true && chkHR.Checked == false && chkCompany.Checked == false)
            {
                BindEventCalendar();
            }
            //For HR Calendar AND Company Calendar.....
            else if (chkHR.Checked == true && chkCompany.Checked == true && chkEvents.Checked == false)
            {
                BindHRCompanyCalendar();
            }
            //For HR Calendar AND Event Calendar.....
            else if (chkHR.Checked == true && chkEvents.Checked == true && chkCompany.Checked == false)
            {
                BindHREventCalendar();
            }
            //For Company Calendar AND Event Calendar.....
            else if (chkCompany.Checked == true && chkEvents.Checked == true && chkHR.Checked == false)
            {
                BindCompanyEventCalendar();
            }
            //For Company Calendar AND Event Calendar AND HR Calendar.....
            else if (chkHR.Checked == true && chkCompany.Checked == true && chkEvents.Checked == true)
            {
                BindHRCompanyEventCalendar();
            }
            else if (chkHR.Checked == false && chkCompany.Checked == false && chkEvents.Checked == false)
            {
                BindCalendar();
            }
        }

        protected void lbtCustID_Click(object sender, EventArgs e)
        {
            LinkButton CustomerId = (LinkButton)sender;

            SchedulerAppointmentContainer appContainer = (SchedulerAppointmentContainer)CustomerId.Parent;
            Appointment appointment = appContainer.Appointment;
            int i = Convert.ToInt32(appointment.ID);
            DataSet ds = AdminBLL.Instance.GetInterviewDetails(i);
            int applicantId = 0;
            if (ds.Tables[0].Rows.Count > 0)
            {
                applicantId = Convert.ToInt32(ds.Tables[0].Rows[0]["ApplicantId"]);
                ViewState["ApplicantId"] = applicantId;
            }

            Response.Redirect("InstallCreateUser.aspx?ID=" + ViewState["ApplicantId"]);
        }

        protected void ddlStatus_SelectedIndexChanged(object sender, EventArgs e)
        {
            //Below 4 lines is to get that particular row control values
            DropDownList ddlNew = sender as DropDownList;

            SchedulerAppointmentContainer appContainer = (SchedulerAppointmentContainer)ddlNew.Parent;
            Appointment appointment = appContainer.Appointment;
            string Id = (string)appointment.Attributes["ApplicantId"];
            string Status_Old = (string)appointment.Attributes["Status"];
            string Designation = (string)appointment.Attributes["Designation"];
            string FirstName = "";
            string LastName = (string)appointment.Attributes["LastName"];

            string strddlNew = ddlNew.SelectedValue;            
            //Label lblDesignation = (Label)(grow.FindControl("lblDesignation"));
            //Label lblFirstName = (Label)(grow.FindControl("lblFirstName"));
            //Label lblLastName = (Label)(grow.FindControl("lblLastName"));
            //HiddenField lblStatus = (HiddenField)(grow.FindControl("lblStatus"));
            //Label Id = (Label)grow.FindControl("lblid");
            //DropDownList ddl = (DropDownList)grow.FindControl("ddlStatus");
            Session["EditId"] = Id;
            Session["EditStatus"] = ddlNew.SelectedValue;
            Session["DesignitionSC"] = lblDesignation.Text;
            //Session["FirstNameNewSC"] = lblFirstName.Text;
            Session["LastNameNewSC"] = LastName;
            if ((Status_Old == "Active") && (!(Convert.ToString(Session["usertype"]).Contains("Admin")) && !(Convert.ToString(Session["usertype"]).Contains("SM"))))
            {
                BindCalendar();
                ScriptManager.RegisterStartupScript(this, this.GetType(), "alert", "alert('You dont have rights change the status.')", true);
                return;
            }
            else if ((Status_Old == "Active" && ddlNew.SelectedValue != "Deactive") && ((Convert.ToString(Session["usertype"]).Contains("Admin")) || (Convert.ToString(Session["usertype"]).Contains("SM"))))
            {
                ScriptManager.RegisterStartupScript(this, this.GetType(), "Overlay", "overlayPassword();", true);
                return;
            }
            bool status = CheckRequiredFields(ddlNew.SelectedValue, Convert.ToInt32(Id));
            if (!status)
            {
                BindCalendar();
                ScriptManager.RegisterStartupScript(this, this.GetType(), "alert", "alert('Status cannot be changed as required field for selected status are not field')", true);
                return;
            }

            if ((ddlNew.SelectedValue == "Active" || ddlNew.SelectedValue == "Deactive") && (!(Convert.ToString(Session["usertype"]).Contains("Admin")) && !(Convert.ToString(Session["usertype"]).Contains("SM"))))
            {
                ddlNew.SelectedValue = Convert.ToString(Status_Old);
                ScriptManager.RegisterStartupScript(this, this.GetType(), "alert", "alert('You dont have permission to Activate or Deactivate user')", true);
                return;
            }
            else if (ddlNew.SelectedValue == "Rejected")
            {
                ScriptManager.RegisterStartupScript(this, this.GetType(), "Overlay", "overlay()", true);
                return;
            }
            else if (ddlNew.SelectedValue == "InterviewDate")
            {
                ddlInsteviewtime.DataSource = GetTimeIntervals();
                ddlInsteviewtime.DataBind();
                dtInterviewDate.Text = DateTime.Now.AddDays(1).ToShortDateString();
                ddlInsteviewtime.SelectedValue = "10:00 AM";
                ScriptManager.RegisterStartupScript(this, this.GetType(), "Overlay", "overlayInterviewDate()", true);
                return;
            }
            else if (ddlNew.SelectedValue == "Deactive" && ((Convert.ToString(Session["usertype"]).Contains("Admin")) && (Convert.ToString(Session["usertype"]).Contains("SM"))))
            {
                Session["DeactivationStatus"] = "Deactive";
                ScriptManager.RegisterStartupScript(this, this.GetType(), "Overlay", "overlay()", true);
                return;
            }
            else if (ddlNew.SelectedValue == "OfferMade")
            {
                DataSet ds = new DataSet();
                string email = "";
                string HireDate = "";
                string EmpType = "";
                string PayRates = "";
                ds = InstallUserBLL.Instance.ChangeStatus(Convert.ToString(Session["EditStatus"]), Convert.ToInt32(Session["EditId"]), DateTime.Today.ToString("yyyy-MM-dd"), DateTime.Now.ToShortTimeString(), Convert.ToInt32(Session[JG_Prospect.Common.SessionKey.Key.UserId.ToString()]), txtReason.Text);
                if (ds.Tables.Count > 0)
                {
                    if (ds.Tables[0].Rows.Count > 0)
                    {
                        if (Convert.ToString(ds.Tables[0].Rows[0][0]) != "")
                        {
                            email = Convert.ToString(ds.Tables[0].Rows[0][0]);
                        }
                        if (Convert.ToString(ds.Tables[0].Rows[0][1]) != "")
                        {
                            HireDate = Convert.ToString(ds.Tables[0].Rows[0][1]);
                        }
                        if (Convert.ToString(ds.Tables[0].Rows[0][2]) != "")
                        {
                            EmpType = Convert.ToString(ds.Tables[0].Rows[0][2]);
                        }
                        if (Convert.ToString(ds.Tables[0].Rows[0][3]) != "")
                        {
                            PayRates = Convert.ToString(ds.Tables[0].Rows[0][3]);
                        }
                    }
                }
                SendEmail(email, FirstName, LastName, "Offer Made", txtReason.Text, lblDesignation.Text, HireDate, EmpType, PayRates);
                BindCalendar();
                return;
            }

            if (Status_Old== "Active" && (!(Convert.ToString(Session["usertype"]).Contains("Admin"))))
            {
                ScriptManager.RegisterStartupScript(this, this.GetType(), "alert", "alert('Status cannot be changed to any other status other than Deactive once user is Active')", true);
                if (Convert.ToString(Session["PreviousStatusNew"]) != "")
                {
                    ddlNew.SelectedValue = Convert.ToString(Session["PreviousStatusNew"]);
                }
                return;
            }
            else
            {
                InstallUserBLL.Instance.ChangeStatus(Convert.ToString(Session["EditStatus"]), Convert.ToInt32(Session["EditId"]), Convert.ToString(DateTime.Today.ToShortDateString()), DateTime.Now.ToShortTimeString(), Convert.ToInt32(Session[JG_Prospect.Common.SessionKey.Key.UserId.ToString()]), txtReason.Text);
                BindCalendar();
                return;
            }

            if (ddlNew.SelectedValue == "Install Prospect")
            {
                if (Status_Old != "")
                {
                    ddlNew.SelectedValue = Status_Old;
                }
                ScriptManager.RegisterStartupScript(this, this.GetType(), "alert", "alert('Status cannot be changed to Install Prospect')", true);
                return;
            }

        }

        protected void rsAppointments_AppointmentCreated(object sender, AppointmentCreatedEventArgs e)
        {
            string status = e.Appointment.Attributes["Status"];

            DropDownList ddlStatus = (DropDownList)e.Container.FindControl("ddlStatus");
            if (ddlStatus != null)
            {
                ddlStatus.SelectedIndex = ddlStatus.Items.IndexOf(ddlStatus.Items.FindByValue(status.ToString()));
            }
        }

        private bool CheckRequiredFields(string SelectedStatus, int Id)
        {
            DataSet dsNew = new DataSet();
            dsNew = InstallUserBLL.Instance.getuserdetails(Id);
            if (dsNew.Tables.Count > 0)
            {
                if (dsNew.Tables[0].Rows.Count > 0)
                {
                    if (SelectedStatus == "Applicant")
                    {
                        if (Convert.ToString(dsNew.Tables[0].Rows[0][1]) == "" || Convert.ToString(dsNew.Tables[0].Rows[0][2]) == "" || Convert.ToString(dsNew.Tables[0].Rows[0][3]) == "" || Convert.ToString(dsNew.Tables[0].Rows[0][8]) == "" || Convert.ToString(dsNew.Tables[0].Rows[0][38]) == "")
                        {
                            //ScriptManager.RegisterStartupScript(this, this.GetType(), "alert", "alert('Status cannot be changed to Applicant as required fields for it are not filled.')", true);
                            return false;
                        }
                    }
                    else if (SelectedStatus == "OfferMade" || SelectedStatus == "Offer Made")
                    {
                        if (Convert.ToString(dsNew.Tables[0].Rows[0][1]) == "" || Convert.ToString(dsNew.Tables[0].Rows[0][2]) == "" || Convert.ToString(dsNew.Tables[0].Rows[0][4]) == "" || Convert.ToString(dsNew.Tables[0].Rows[0][5]) == "" || Convert.ToString(dsNew.Tables[0].Rows[0][11]) == "" || Convert.ToString(dsNew.Tables[0].Rows[0][12]) == "" || Convert.ToString(dsNew.Tables[0].Rows[0][13]) == "" || Convert.ToString(dsNew.Tables[0].Rows[0][3]) == "" || Convert.ToString(dsNew.Tables[0].Rows[0][8]) == "" || Convert.ToString(dsNew.Tables[0].Rows[0][38]) == "" || Convert.ToString(dsNew.Tables[0].Rows[0][44]) == "" || Convert.ToString(dsNew.Tables[0].Rows[0][46]) == "" || Convert.ToString(dsNew.Tables[0].Rows[0][48]) == "" || Convert.ToString(dsNew.Tables[0].Rows[0][50]) == "" || Convert.ToString(dsNew.Tables[0].Rows[0][100]) == "")
                        {
                            //ScriptManager.RegisterStartupScript(this, this.GetType(), "alert", "alert('Status cannot be changed to Offer Made as required fields for it are not filled.')", true);
                            return false;
                        }
                    }
                    else if (SelectedStatus == "Active")
                    {
                        if (Convert.ToString(dsNew.Tables[0].Rows[0][1]) == "" || Convert.ToString(dsNew.Tables[0].Rows[0][2]) == "" || Convert.ToString(dsNew.Tables[0].Rows[0][3]) == "" || Convert.ToString(dsNew.Tables[0].Rows[0][4]) == "" || Convert.ToString(dsNew.Tables[0].Rows[0][5]) == "" || Convert.ToString(dsNew.Tables[0].Rows[0][7]) == "" || Convert.ToString(dsNew.Tables[0].Rows[0][9]) == "" || Convert.ToString(dsNew.Tables[0].Rows[0][11]) == "" || Convert.ToString(dsNew.Tables[0].Rows[0][12]) == "" || Convert.ToString(dsNew.Tables[0].Rows[0][13]) == "" || Convert.ToString(dsNew.Tables[0].Rows[0][17]) == "" || Convert.ToString(dsNew.Tables[0].Rows[0][16]) == "" || Convert.ToString(dsNew.Tables[0].Rows[0][17]) == "" || Convert.ToString(dsNew.Tables[0].Rows[0][8]) == "" || Convert.ToString(dsNew.Tables[0].Rows[0][18]) == "" || Convert.ToString(dsNew.Tables[0].Rows[0][19]) == "" || Convert.ToString(dsNew.Tables[0].Rows[0][20]) == "" || Convert.ToString(dsNew.Tables[0].Rows[0][35]) == "" || Convert.ToString(dsNew.Tables[0].Rows[0][38]) == "" || Convert.ToString(dsNew.Tables[0].Rows[0][39]) == "" || Convert.ToString(dsNew.Tables[0].Rows[0][44]) == "" || Convert.ToString(dsNew.Tables[0].Rows[0][46]) == "" || Convert.ToString(dsNew.Tables[0].Rows[0][48]) == "" || Convert.ToString(dsNew.Tables[0].Rows[0][50]) == "" || Convert.ToString(dsNew.Tables[0].Rows[0][100]) == "")
                        {
                            //ScriptManager.RegisterStartupScript(this, this.GetType(), "alert", "alert('Status cannot be changed to Offer Made as required fields for it are not filled.')", true); 
                            return false;
                        }
                    }
                }
            }
            return true;
        }


        public List<string> GetTimeIntervals()
        {
            List<string> timeIntervals = new List<string>();
            TimeSpan startTime = new TimeSpan(0, 0, 0);
            DateTime startDate = new DateTime(DateTime.MinValue.Ticks); // Date to be used to get shortTime format.
            for (int i = 0; i < 48; i++)
            {
                int minutesToBeAdded = 30 * i;      // Increasing minutes by 30 minutes interval
                TimeSpan timeToBeAdded = new TimeSpan(0, minutesToBeAdded, 0);
                TimeSpan t = startTime.Add(timeToBeAdded);
                DateTime result = startDate + t;
                timeIntervals.Add(result.ToShortTimeString());      // Use Date.ToShortTimeString() method to get the desired format                
            }
            return timeIntervals;
        }

        private void SendEmail(string emailId, string FName, string LName, string status, string Reason, string Designition, string HireDate, string EmpType, string PayRates)
        {
            try
            {
                string fullname = FName + " " + LName;
                string HTML_TAG_PATTERN = "<.*?>";
                DataSet ds = AdminBLL.Instance.GetEmailTemplate(Designition);// AdminBLL.Instance.FetchContractTemplate(104);

                if (ds == null)
                {
                    ds = AdminBLL.Instance.GetEmailTemplate("Admin");
                }
                else if (ds.Tables[0].Rows.Count == 0)
                {
                    ds = AdminBLL.Instance.GetEmailTemplate("Admin");
                }
                string strHeader = ds.Tables[0].Rows[0]["HTMLHeader"].ToString(); //GetEmailHeader(status);
                string strBody = ds.Tables[0].Rows[0]["HTMLBody"].ToString(); //GetEmailBody(status);
                string strFooter = ds.Tables[0].Rows[0]["HTMLFooter"].ToString(); // GetFooter(status);
                string strsubject = ds.Tables[0].Rows[0]["HTMLSubject"].ToString();

                string userName = ConfigurationManager.AppSettings["VendorCategoryUserName"].ToString();
                string password = ConfigurationManager.AppSettings["VendorCategoryPassword"].ToString();

                strBody = strBody.Replace("#Name#", FName).Replace("#name#", FName);
                strBody = strBody.Replace("#Date#", dtInterviewDate.Text).Replace("#date#", dtInterviewDate.Text);
                strBody = strBody.Replace("#Time#", ddlInsteviewtime.SelectedValue).Replace("#time#", ddlInsteviewtime.SelectedValue);
                strBody = strBody.Replace("#Designation#", Designition).Replace("#designation#", Designition);

                strFooter = strFooter.Replace("#Name#", FName).Replace("#name#", FName);
                strFooter = strFooter.Replace("#Date#", dtInterviewDate.Text).Replace("#date#", dtInterviewDate.Text);
                strFooter = strFooter.Replace("#Time#", ddlInsteviewtime.SelectedValue).Replace("#time#", ddlInsteviewtime.SelectedValue);
                strFooter = strFooter.Replace("#Designation#", Designition).Replace("#designation#", Designition);

                strBody = strBody.Replace("Lbl Full name", fullname);
                strBody = strBody.Replace("LBL position", Designition);
                //strBody = strBody.Replace("lbl: start date", txtHireDate.Text);
                //strBody = strBody.Replace("($ rate","$"+ txtHireDate.Text);
                strBody = strBody.Replace("Reason", Reason);
                //Hi #lblFName#, <br/><br/>You are requested to appear for an interview on #lblDate# - #lblTime#.<br/><br/>Regards,<br/>
                StringBuilder Body = new StringBuilder();
                MailMessage Msg = new MailMessage();
                //Sender e-mail address.
                Msg.From = new MailAddress(userName, "JGrove Construction");
                // Recipient e-mail address.
                Msg.To.Add(emailId);
                Msg.Bcc.Add(new MailAddress("shabbir.kanchwala@straitapps.com", "Shabbir Kanchwala"));
                Msg.CC.Add(new MailAddress("jgrove.georgegrove@gmail.com", "Justin Grove"));

                Msg.Subject = strsubject;// "JG Prospect Notification";
                Body.Append(strHeader);
                Body.Append(strBody);
                Body.Append(strFooter);
                if (status == "OfferMade")
                {
                    createForeMenForJobAcceptance(Convert.ToString(Body), FName, LName, Designition, emailId, HireDate, EmpType, PayRates);
                }
                if (status == "Deactive")
                {
                    CreateDeactivationAttachment(Convert.ToString(Body), FName, LName, Designition, emailId, HireDate, EmpType, PayRates);
                }
                Msg.Body = Convert.ToString(Body);
                Msg.IsBodyHtml = true;
                // your remote SMTP server IP.

                for (int i = 0; i < ds.Tables[1].Rows.Count; i++)
                {
                    string sourceDir = Server.MapPath(ds.Tables[1].Rows[i]["DocumentPath"].ToString());
                    if (File.Exists(sourceDir))
                    {
                        Attachment attachment = new Attachment(sourceDir);
                        attachment.Name = Path.GetFileName(sourceDir);
                        Msg.Attachments.Add(attachment);
                    }
                }


                SmtpClient sc = new SmtpClient(ConfigurationManager.AppSettings["smtpHost"].ToString(), Convert.ToInt32(ConfigurationManager.AppSettings["smtpPort"].ToString()));


                NetworkCredential ntw = new System.Net.NetworkCredential(userName, password);
                sc.UseDefaultCredentials = false;
                sc.Credentials = ntw;

                sc.DeliveryMethod = SmtpDeliveryMethod.Network;
                sc.EnableSsl = Convert.ToBoolean(ConfigurationManager.AppSettings["enableSSL"].ToString()); // runtime encrypt the SMTP communications using SSL
                try
                {
                    sc.Send(Msg);
                }
                catch (Exception ex)
                {
                }

                Msg = null;
                sc.Dispose();
                sc = null;
                Page.RegisterStartupScript("UserMsg", "<script>alert('An email notification has sent on " + emailId + ".');}</script>");
            }
            catch (Exception ex)
            {
                Console.WriteLine("{0} Exception caught.", ex);
            }
            // //SmtpClient smtp = new SmtpClient();
            // //MailMessage email_msg = new MailMessage();
            // //email_msg.To.Add(emailId);
            // //email_msg.From = new MailAddress("customsoft.test@gmail.com", "Credit Chex");
            // StringBuilder Body = new StringBuilder();
            // Body.Append("Hello " + FName + " " + LName + ",");
            // Body.Append("<br>");
            // Body.Append("Your stattus for the JG Prospect is :" + status);
            // Body.Append("<br>");
            // //if (status == "Source" || status == "Rejected" || status == "Interview Date" || status == "Offer Made")
            // //{
            //     Body.Append(Reason);
            // //}
            // Body.Append("<br>");
            // Body.Append("Tanking you");
            //// AlternateView htmlView = AlternateView.CreateAlternateViewFromString(Body, null, "text/html");
            //// LinkedResource imagelink3 = new LinkedResource(HttpContext.Current.Server.MapPath("~/images/logo.png"), "image/png");
            // //imagelink3.ContentId = "imageId1";
            // //imagelink3.TransferEncoding = System.Net.Mime.TransferEncoding.Base64;
            // //htmlView.LinkedResources.Add(imagelink3);
            // var smtp = new System.Net.Mail.SmtpClient();
            // {
            //     smtp.Host = "smtp.gmail.com";
            //     smtp.Port = 587;
            //     smtp.EnableSsl = true;
            //     smtp.DeliveryMethod = System.Net.Mail.SmtpDeliveryMethod.Network;
            //     smtp.Credentials = new NetworkCredential("","q$7@wt%j*j*65ba#3M@9P6");
            //     smtp.Timeout = 20000;
            // }
            // // Passing values to smtp object
            // smtp.Send("", emailId, "JG Prospect Notification", Convert.ToString(Body));

        }

        public void createForeMenForJobAcceptance(string str_Body, string FName, string LName, string Designition, string emailId, string HireDate, string EmpType, string PayRates)
        {
            //copy sample file for Foreman Job Acceptance letter template
            string str_date = DateTime.Now.ToString().Replace("/", "");
            str_date = str_date.Replace(":", "");
            str_date = str_date.Replace("-", "");
            str_date = str_date.Replace(" ", "");
            string SourcePath = @"~/Sr_App/MailDocSample/ForemanJobAcceptancelettertemplate.docx";
            string TargetPath = @"~/Sr_App/MailDocument/" + str_date + FName + "ForemanJobAcceptanceletter.docx";
            System.IO.File.Copy(Server.MapPath(SourcePath), Server.MapPath(TargetPath), true);
            //modify word document
            object missing = System.Reflection.Missing.Value;
            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
            Microsoft.Office.Interop.Word.Document aDoc = null;
            object Target = Server.MapPath(TargetPath);
            if (File.Exists(Server.MapPath(TargetPath)))
            {
                DateTime today = DateTime.Now;
                object readonlyNew = false;
                object isVisible = false;
                wordApp.Visible = false;
                FileInfo objFInfo = new FileInfo(Server.MapPath(TargetPath));
                objFInfo.IsReadOnly = false;
                aDoc = wordApp.Documents.Open(ref Target, ref missing, ref readonlyNew, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref isVisible, ref missing, ref missing, ref missing, ref missing);
                aDoc.Activate();
                this.FindAndReplace(wordApp, "LBL Date", DateTime.Now.ToShortDateString());
                this.FindAndReplace(wordApp, "Lbl Full name", FName + " " + LName);
                this.FindAndReplace(wordApp, "LBL name", FName + " " + LName);
                this.FindAndReplace(wordApp, "LBL position", Designition);
                this.FindAndReplace(wordApp, "lbl fulltime", EmpType);
                this.FindAndReplace(wordApp, "lbl: start date", HireDate);
                this.FindAndReplace(wordApp, "$ rate", PayRates);
                this.FindAndReplace(wordApp, "lbl: next pay period", "");
                this.FindAndReplace(wordApp, "lbl: paycheck date", "");
                aDoc.SaveAs(ref Target, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);
                aDoc.Close(ref missing, ref missing, ref missing);
            }
            using (System.Net.Mail.MailMessage mm = new System.Net.Mail.MailMessage("qat2015team@gmail.com", emailId))
            {
                try
                {
                    mm.Subject = "Foreman Job Acceptance";
                    mm.Body = str_Body;
                    mm.Attachments.Add(new Attachment(Server.MapPath(TargetPath)));
                    mm.IsBodyHtml = true;
                    SmtpClient smtp = new SmtpClient();
                    smtp.Host = "smtp.gmail.com";
                    smtp.EnableSsl = true;
                    NetworkCredential NetworkCred = new NetworkCredential("qat2015team@gmail.com", "q$7@wt%j*65ba#3M@9P6");
                    smtp.UseDefaultCredentials = true;
                    smtp.Credentials = NetworkCred;
                    smtp.Port = 587;
                    smtp.Send(mm);
                }
                catch (Exception ex)
                {
                    ScriptManager.RegisterStartupScript(this, this.GetType(), "alert", "alert('" + ex.Message + "')", true);
                }
                //ClientScript.RegisterStartupScript(GetType(), "alert", "alert('Email sent.');", true);
            }
        }


        private void CreateDeactivationAttachment(string MailBody, string FName, string LName, string Designition, string emailId, string HireDate, string EmpType, string PayRates)
        {
            string str_date = DateTime.Now.ToString().Replace("/", "");
            str_date = str_date.Replace(":", "");
            str_date = str_date.Replace("-", "");
            str_date = str_date.Replace(" ", "");
            string SourcePath = @"~/Sr_App/MailDocSample/DeactivationMail.doc";
            string TargetPath = @"~/Sr_App/MailDocument/" + str_date + FName + "DeactivationMail.doc";
            System.IO.File.Copy(Server.MapPath(SourcePath), Server.MapPath(TargetPath), true);
            //modify word document
            object missing = System.Reflection.Missing.Value;
            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
            Microsoft.Office.Interop.Word.Document aDoc = null;
            object Target = Server.MapPath(TargetPath);
            if (File.Exists(Server.MapPath(TargetPath)))
            {
                DateTime today = DateTime.Now;
                object readonlyNew = false;
                object isVisible = false;
                wordApp.Visible = false;
                FileInfo objFInfo = new FileInfo(Server.MapPath(TargetPath));
                objFInfo.IsReadOnly = false;
                aDoc = wordApp.Documents.Open(ref Target, ref missing, ref readonlyNew, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref isVisible, ref missing, ref missing, ref missing, ref missing);
                aDoc.Activate();
                this.FindAndReplace(wordApp, "name", FName + " " + LName);
                this.FindAndReplace(wordApp, "HireDate", HireDate);
                this.FindAndReplace(wordApp, "full time or part  time", EmpType);
                this.FindAndReplace(wordApp, "HourlyRate", PayRates);
                this.FindAndReplace(wordApp, "WorkingStatus", "No");
                this.FindAndReplace(wordApp, "LastWorkingDay", DateTime.Now.ToShortDateString());
                aDoc.SaveAs(ref Target, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);
                aDoc.Close(ref missing, ref missing, ref missing);
            }
            using (MailMessage mm = new MailMessage("qat2015team@gmail.com", emailId))
            {
                try
                {
                    mm.Subject = "Deactivation";
                    mm.Body = MailBody;
                    mm.Attachments.Add(new Attachment(Server.MapPath(TargetPath)));
                    mm.IsBodyHtml = true;
                    SmtpClient smtp = new SmtpClient();
                    smtp.Host = "smtp.gmail.com";
                    smtp.EnableSsl = true;
                    NetworkCredential NetworkCred = new NetworkCredential("qat2015team@gmail.com", "q$7@wt%j*65ba#3M@9P6");
                    smtp.UseDefaultCredentials = true;
                    smtp.Credentials = NetworkCred;
                    smtp.Port = 587;
                    smtp.Send(mm);
                }
                catch (Exception ex)
                {
                    ScriptManager.RegisterStartupScript(this, this.GetType(), "alert", "alert('" + ex.Message + "')", true);
                }
            }
        }
        private void FindAndReplace(Microsoft.Office.Interop.Word.Application wordApp, object findText, object replaceText)
        {
            object matchCase = true;
            object matchWholeWord = true;
            object matchWildCards = false;
            object matchSoundsLike = false;
            object matchAllWordForms = false;
            object forward = true;
            object format = false;
            object matchKashida = false;
            object matchDiacritics = false;
            object matchAlefHamza = false;
            object matchControl = false;
            object read_only = false;
            object visible = true;
            object replace = 2;
            object wrap = 1;
            wordApp.Selection.Find.Execute(ref findText, ref matchCase,
                ref matchWholeWord, ref matchWildCards, ref matchSoundsLike,
                ref matchAllWordForms, ref forward, ref wrap, ref format,
                ref replaceText, ref replace, ref matchKashida,
                        ref matchDiacritics,
                ref matchAlefHamza, ref matchControl);
        }

        protected void btnSave_Click(object sender, EventArgs e)
        {
            if (Convert.ToString(Session["DeactivationStatus"]) == "Deactive")
            {
                DataSet ds = new DataSet();
                string email = "";
                string HireDate = "";
                string EmpType = "";
                string PayRates = "";
                ds = InstallUserBLL.Instance.ChangeStatus(Convert.ToString(Session["EditStatus"]), Convert.ToInt32(Session["EditId"]), DateTime.Today.ToString("yyyy-MM-dd"), DateTime.Now.ToShortTimeString(), Convert.ToInt32(Session[JG_Prospect.Common.SessionKey.Key.UserId.ToString()]), txtReason.Text);
                if (ds.Tables.Count > 0)
                {
                    if (ds.Tables[0].Rows.Count > 0)
                    {
                        if (Convert.ToString(ds.Tables[0].Rows[0][0]) != "")
                        {
                            email = Convert.ToString(ds.Tables[0].Rows[0][0]);
                        }
                        if (Convert.ToString(ds.Tables[0].Rows[0][1]) != "")
                        {
                            HireDate = Convert.ToString(ds.Tables[0].Rows[0][1]);
                        }
                        if (Convert.ToString(ds.Tables[0].Rows[0][2]) != "")
                        {
                            EmpType = Convert.ToString(ds.Tables[0].Rows[0][2]);
                        }
                        if (Convert.ToString(ds.Tables[0].Rows[0][3]) != "")
                        {
                            PayRates = Convert.ToString(ds.Tables[0].Rows[0][3]);
                        }
                    }
                }
                SendEmail(email, Convert.ToString(Session["FirstNameNewSC"]), Convert.ToString(Session["LastNameNewSC"]), "Deactivation", txtReason.Text, Convert.ToString(Session["DesignitionSC"]), HireDate, EmpType, PayRates);
            }
            else
            {
                InstallUserBLL.Instance.ChangeStatus(Convert.ToString(Session["EditStatus"]), Convert.ToInt32(Session["EditId"]), Convert.ToString(DateTime.Today.ToShortDateString()), DateTime.Now.ToShortTimeString(), Convert.ToInt32(Session[JG_Prospect.Common.SessionKey.Key.UserId.ToString()]), txtReason.Text);
                BindCalendar();
            }
            ScriptManager.RegisterStartupScript(this, this.GetType(), "Overlay", "ClosePopup()", true);
            return;
        }

        protected void btnSaveInterview_Click(object sender, EventArgs e)
        {
            DataSet ds = new DataSet();
            string email = "";
            string HireDate = "";
            string EmpType = "";
            string PayRates = "";


            //string InterviewDate = dtInterviewDate.Text;
            DateTime interviewDate;
            DateTime.TryParse(dtInterviewDate.Text, out interviewDate);
            if (interviewDate == null)
            {
                ScriptManager.RegisterStartupScript(this, this.GetType(), "Overlay", "alert('Invalid Interview Date, Please verify');", true);
                return;
            }
            ds = InstallUserBLL.Instance.ChangeStatus(Convert.ToString(Session["EditStatus"]), Convert.ToInt32(Session["EditId"]), interviewDate.ToString("yyyy-MM-dd"), ddlInsteviewtime.SelectedItem.Text, Convert.ToInt32(Session[JG_Prospect.Common.SessionKey.Key.UserId.ToString()]), txtReason.Text);
            if (ds.Tables.Count > 0)
            {
                if (ds.Tables[0].Rows.Count > 0)
                {
                    if (Convert.ToString(ds.Tables[0].Rows[0][0]) != "")
                    {
                        email = Convert.ToString(ds.Tables[0].Rows[0][0]);
                    }
                    if (Convert.ToString(ds.Tables[0].Rows[0][1]) != "")
                    {
                        HireDate = Convert.ToString(ds.Tables[0].Rows[0][1]);
                    }
                    if (Convert.ToString(ds.Tables[0].Rows[0][2]) != "")
                    {
                        EmpType = Convert.ToString(ds.Tables[0].Rows[0][2]);
                    }
                    if (Convert.ToString(ds.Tables[0].Rows[0][3]) != "")
                    {
                        PayRates = Convert.ToString(ds.Tables[0].Rows[0][3]);
                    }
                }
            }
            SendEmail(email, Convert.ToString(Session["FirstNameNewSC"]), Convert.ToString(Session["LastNameNewSC"]), "Interview Date Auto Email", txtReason.Text, Convert.ToString(Session["DesignitionSC"]), HireDate, EmpType, PayRates);
            BindCalendar();
            ScriptManager.RegisterStartupScript(this, this.GetType(), "Overlay", "ClosePopupInterviewDate()", true);
            return;
        }

    }
}