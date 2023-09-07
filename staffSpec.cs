using System;
using System.Collections.Generic;
using System.Data;
using System.Text;

namespace couplerWriter
{
    public class staffSpec
    {
        public String
            NDSName, EmailAddress, EmailAddressGoogle, Context, QLId, StaffID, Template, Profile,
            Forename, Surname, SQLSurname, /*Aka, */ 
            Site, DeptCode, Department, JobTitle, EmpType, Tel, LocID, VisitingLec, WCG_Ltd,
            HomeVol,SecondHomeVol, ThirdHomeVol, HomeVolRestrict, SharedVolRestrict, Vol1VolRestrict, PostOffice, ActionData,
            GroupMembership,EmpID
        ;
        public Boolean GWise = false;
        public int queueItem;
        public int attempts;
        public String action = "";
        public Boolean WhenRead = false;
        public String staffGroup = "";
        public String staffStartChar = "";
        public Boolean keepProxy;

        public staffSpec(DataRowView aDRV)
        {
            // SJL 20090427
            // changed to allow for a DELETED wcStaffIdentity Record
            // If the wcStaffIdentity Record is deleted then get the QLid from the actiondata
            // and set the NDSname to the NetworkName of the Coupler job
            if (aDRV["whenRead"].ToString().Trim().Length > 0 ) WhenRead = true;
            if (String.IsNullOrEmpty(aDRV["NDSName"].ToString()) | aDRV["NDSName"].ToString().Length < 1)
            {
                keepProxy = false;
                ActionData = aDRV["actionData"].ToString().Trim();
                String[] AD = aDRV["actionData"].ToString().Trim().Split('=');
                if (AD[0].ToString().ToLower() == "empnum")
                {
                    QLId = AD[1].ToString();
                }
                else
                {
                    QLId = aDRV["EmployeeNumber"].ToString().Trim();
                }
                NDSName = aDRV["NetworkName"].ToString().Trim();                
                EmailAddress = NDSName + "@warkscol.ac.uk";
                StaffID = "Deleted";
                //Forename = "Deleted";
                //Surname = "Deleted";
                Forename = NDSName;
                Surname = "warwickshire.ac.uk";
                SQLSurname = "Deleted";
                Site = "Deleted";
                DeptCode = "Deleted";
                Department = "Deleted";
                JobTitle = "Deleted";
                EmpType = "Deleted";
                Tel = "Deleted";
                VisitingLec = "False";
                if (!String.IsNullOrEmpty(aDRV["queueItem"].ToString())) queueItem = (int)aDRV["queueItem"];
                if (!String.IsNullOrEmpty(aDRV["attempts"].ToString())) attempts = (int)aDRV["attempts"];
                if (!String.IsNullOrEmpty(aDRV["action"].ToString())) action = aDRV["action"].ToString();

                //LocID = "Deleted";
                LocID = "LEAMINGTON SPA";

                GWise = false;

                Context = ".staff.lspa.wc";
                HomeVol = ".orion_home.lspa.wc";
                SecondHomeVol = ".orion_gwarc.lspa.wc";
                ThirdHomeVol = ".orion_home.lspa.wc";
                Template = "LSPA_AUTOTEMPLATE";
                Profile = ".LSPASTAFF.STAFF.LSPA.WC";
                PostOffice = "po-lspa.warkscollege";
                SharedVolRestrict = ".orion_shared.lspa.wc:250000";
                GroupMembership = ".STA_0ther.VDI_GROUPS.LSPA.WC";
                Vol1VolRestrict = ".orion_student_vol1.lspa.wc:20000";
            }
            else
            {
                if (aDRV["keepProxy"].ToString().Trim().ToLower() == "true") keepProxy = true;
                NDSName = aDRV["NDSName"].ToString().Trim();
                staffGroup = ".STA_0ther.VDI_GROUPS.LSPA.WC";//default group
                staffStartChar = NDSName.ToString().Trim().Substring(0, 1).ToUpper();
                if ("ABCDEFGHIJKLMNOPQRSTUVWXYZ".Contains(staffStartChar))
                {
                    staffGroup = ".STA_" + staffStartChar.ToString() + ".VDI_GROUPS.LSPA.WC";
                }
 
                EmailAddress = NDSName + "@warkscol.ac.uk";
                EmailAddressGoogle = NDSName + "@warwickshire.ac.uk";
                VisitingLec = aDRV["VisitingLec"].ToString().Trim();
                WCG_Ltd = aDRV["WCG_Ltd"].ToString().Trim();
                QLId = aDRV["EmployeeNumber"].ToString().Trim();
                EmpID = aDRV["EmployeeNumber"].ToString().Trim();
                StaffID = QLId.Trim().PadLeft(8, '0');
                Forename = aDRV["forename"].ToString().Trim();
                Surname = aDRV["surname"].ToString().Trim();
                SQLSurname = Surname.Replace("'", "''");
                Site = aDRV["loc_id"].ToString().Trim();
                DeptCode = aDRV["department"].ToString().Trim();
                Department = aDRV["department"].ToString().Trim();
                JobTitle = aDRV["job_title"].ToString().Trim();
                EmpType = aDRV["emp_type"].ToString().Trim();
                Tel = aDRV["Telephone"].ToString().Trim();
                ActionData = aDRV["actionData"].ToString().Trim();
                if (!String.IsNullOrEmpty(aDRV["queueItem"].ToString())) queueItem = (int)aDRV["queueItem"];
                if (!String.IsNullOrEmpty(aDRV["attempts"].ToString())) attempts = (int)aDRV["attempts"];
                if (!String.IsNullOrEmpty(aDRV["action"].ToString())) action = aDRV["action"].ToString();

                LocID = aDRV["loc_id"].ToString().Trim();

               // GWise = (Boolean)aDRV["statusGroupwise"];
                GWise = false;

                Context = ".staff.lspa.wc";
                HomeVol = ".orion_home.lspa.wc";
                SecondHomeVol = ".orion_gwarc.lspa.wc";
                ThirdHomeVol = ".orion_home.lspa.wc";
                Template = "LSPA_AUTOTEMPLATE";
                Profile = ".LSPASTAFF.STAFF.LSPA.WC";
                PostOffice = "po-lspa.warkscollege";
                SharedVolRestrict = ".orion_shared.lspa.wc:250000";
                Vol1VolRestrict = "";
                GroupMembership = staffGroup.ToString();
                switch (LocID.ToUpper())
                {
                    case "LEAMINGTON SPA":
                        Vol1VolRestrict = ".orion_student_vol1.lspa.wc:20000";
                        break;
                    case "RUGBY":
                        Context = ".staff.rug.wc";
                        Template = "RUG_AUTOTEMPLATE";
                        HomeVol = ".ka_staff_home.rug.wc";
                        SecondHomeVol = ".ka_gwarc.rug.wc";
                        ThirdHomeVol = ".ka_staff_home.rug.wc";
                        Profile = ".RUGSTAFF.STAFF.LSPA.WC";
                        PostOffice = "po-rugby.warkscollege";
                        SharedVolRestrict = ".ka_staff_shared.rug.wc:250000";
                        Vol1VolRestrict = ".ka_student_vol1.rug.wc:20000";
                        //GroupMembership = "._ALL-RUGBY-STAFF.GROUPWISE.WC";
                        break;
                    case "TRIDENT PARK":
                        Context = ".staff.trident.wc";
                        HomeVol = ".rover_staff_home.trident.wc";
                        SecondHomeVol = ".rover_gwarc.trident.wc";
                        ThirdHomeVol = ".rover_staff_home.trident.wc";
                        Template = "TRI_AUTOTEMPLATE";
                        Profile = ".TRISTAFF.STAFF.LSPA.WC";
                        PostOffice = "po-lspa.warkscollege";
                        SharedVolRestrict = ".orion_shared.lspa.wc:250000";
                        //GroupMembership = "._ALL-TRIDENT-STAFF.GROUPWISE.WC";
                        break;
                    case "MORETON MORRELL":
                        Context = ".staff.mm.wc";
                        HomeVol = ".sierra_staff_home.mm.wc";
                        SecondHomeVol = ".sierra_gwarc.mm.wc";
                        ThirdHomeVol = ".sierra_staff_home.mm.wc";
                        Template = "MM_AUTOTEMPLATE";
                        Profile = ".MMSTAFF.STAFF.LSPA.WC";
                        PostOffice = "po-mm.warkscollege";
                        SharedVolRestrict = ".sierra_staff_shared.mm.wc:250000";
                        //GroupMembership = "._ALL-MORETON-STAFF.GROUPWISE.WC";
                        break;
                    case "HENLEY IN ARDEN":
                        Context = ".staff.ardn.wc";
                        HomeVol = ".nova_staff_home.ardn.wc";
                        SecondHomeVol = ".nova_gwarc.ardn.wc";
                        ThirdHomeVol = ".nova_staff_home.ardn.wc";
                        Template = "ARDN_AUTOTEMPLATE";
                        Profile = ".ARDNSTAFF.STAFF.LSPA.WC";
                        PostOffice = "po-lspa.warkscollege";
                        SharedVolRestrict = ".orion_shared.lspa.wc:250000";
                        //GroupMembership = "._ALL-HENLEY-STAFF.GROUPWISE.WC";
                        break;
                    case "PERSHORE":
                        Context = ".staff.per.wc";
                        HomeVol = ".bravo_staff_home.per.wc";
                        SecondHomeVol = ".bravo_gwarc.per.wc";
                        ThirdHomeVol = ".bravo_staff_home.per.wc";
                        Template = "LSPA_AUTOTEMPLATE";
                        PostOffice = "po-per.warkscollege";
                        SharedVolRestrict = ".bravo_staff_shared.per.wc:250000";
                        //GroupMembership = "._ALL-PERSHORE-STAFF.GROUPWISE.WC";
                        break;
                    default:
                        break;
                }
            }
            HomeVolRestrict = HomeVol + ":20000";
        }
    }

    public class staffSpecSL : SortedList<String, staffSpec>
    {
        public staffSpecSL(DataView aDV, List<String> aSkipped)
            : base()
        {
            foreach (DataRowView drv in aDV)
            {
                if (this.ContainsKey(drv["NetworkName"].ToString()))
                    aSkipped.Add(drv["queueItem"].ToString());
                else
                    Add(drv["NetworkName"].ToString(), new staffSpec(drv));
            }
        }
    }
}
