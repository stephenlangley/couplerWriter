using System;
using System.IO;
using System.Collections.Generic;
using System.Text;
using System.DirectoryServices;
using System.Data;
using System.Data.SqlClient;
using System.Configuration;


namespace couplerWriter
{

    class jrbCoupler : coupler
    {
        protected String mJRBFilesPath = "";

        protected SortedList<String, List<String>> mAttributeValues;

        private void addAttributeValue(
            SortedList<String, List<String>> aAttributeSet,
            String aAttributeName, String aAttributeValue
        )
        {
            if (!aAttributeSet.ContainsKey(aAttributeName))
            {
                List<String> wAttributeList = new List<string>();
                aAttributeSet.Add(aAttributeName, wAttributeList);
            }
            aAttributeSet[aAttributeName].Add(aAttributeValue);
        }

        protected String attributeValue(String aAttributeName)
        {
            String wValue = "";
            try
            {
                if (mAttributeValues.ContainsKey(aAttributeName))
                    if (mAttributeValues[aAttributeName].Count > 0)
                        wValue = mAttributeValues[aAttributeName][0];
            }
            catch (Exception ue)
            {
                String wMsg = ue.Message;
            }
            return wValue;
        }



        protected String[] attributeArray(String aAttributeName)
        {
            String[] wValue = null;
            try
            {
                if (mAttributeValues.ContainsKey(aAttributeName))
                    if (mAttributeValues[aAttributeName].Count > 0)
                        wValue = mAttributeValues[aAttributeName].ToArray();
            }
            catch (Exception ue)
            {
                String wMsg = ue.Message;
            }
            return wValue;
        }

        protected SortedList<String, List<String>> GetLDAPInfo(String aFilter)
        {
            SortedList<String, List<String>> wSL = null;
            String domainAndUsername = @"LDAP://212.219.42.19/o=WC";
            string userName = string.Empty;
            string passWord = string.Empty;
            AuthenticationTypes at = AuthenticationTypes.Anonymous;

            //Create the object necessary to read the info from the LDAP directory
            DirectoryEntry entry = new DirectoryEntry(domainAndUsername, userName, passWord, at);
            DirectorySearcher mySearcher = new DirectorySearcher(entry);
            SearchResultCollection results;
            mySearcher.Filter = aFilter;

            try
            {
                results = mySearcher.FindAll();
                if (results.Count > 0)
                {
                    SearchResult resEnt = results[0];
                    {
                        wSL = new SortedList<String, List<String>>();
                        ResultPropertyCollection propcoll = resEnt.Properties;
                        String wKey = "";
                        foreach (string key in propcoll.PropertyNames)
                        {
                            wKey = key;
                            switch (key)
                            {
                                case "sn": wKey = "surname"; break;
                                case "l": wKey = "location"; break;
                                case "st": wKey = "state"; break;
                                case "ngwmailboxexpirationtime": wKey = "gwexpire"; break;
                                case "groupmembership": wKey = "grpmbr"; break;
                                case "uid": wKey = "userid"; break;
                                default: break;
                            }
                            if (key != "nsimhint")
                                foreach (object values in propcoll[key])
                                {
                                    //added 11/04/2011 SJL: needed to add the string for the ndshomedirectory
                                    //as part of the test for creating additional home directories.
                                    //originally the text 'system.byte[] was written to the wSL
                                    String sValue = "";
                                    if (values.ToString() == "System.Byte[]")
                                    {
                                        Byte[] x;
                                        x = (Byte[])values;
                                        Char v;
                                        for (int i = 0; i < x.Length; i++)
                                        {
                                            v = Convert.ToChar(x[i]);
                                            sValue = sValue + v.ToString();
                                        }
                                        addAttributeValue(wSL, wKey, sValue.ToString());
                                    }
                                    else
                                    {
                                        addAttributeValue(wSL, wKey, values.ToString());
                                    }
                                }
                        }
                        //mResult.Add(wSL["cn"][0], wSL);
                    }
                }
            }
            catch (Exception ex)
            {
                string msg = ex.Message;
            }
            return wSL;
        }

        public jrbCoupler() : base() { }

        public jrbCoupler(String aJRBFilesPath, String aActionName, int aTryCount)
        {
            mJRBFilesPath = aJRBFilesPath;
            mActionName = aActionName;
            mTryCount = aTryCount;
        }

        protected virtual String doSelectedItem(staffSpec aStaffSpec) { return ""; }

        protected override void doSelectedPhase(staffSpecSL aStaffSpecSL, DateTime runDateTime)
        {
            if (aStaffSpecSL.Count > 0)
            {
                StreamWriter wNewNDSStream = new StreamWriter(
                    mJRBFilesPath + runDateTime.ToString("yyyyMMddHHmmss") + ".txt"
                    );
                foreach (staffSpec ss in aStaffSpecSL.Values) wNewNDSStream.WriteLine(doSelectedItem(ss));
                wNewNDSStream.Close();
            }
        }

        protected override void testSelectedPhase(staffSpecSL aStaffSpecSL, DateTime runDateTime) { }

    }

    class jrbCouplerCreate : jrbCoupler
    {
        public jrbCouplerCreate
            (String aJRBFilesPath, String aActionName, int aTryCount)
            : base(aJRBFilesPath, aActionName, aTryCount) { }

        protected override String doSelectedItem(staffSpec aStaffSpec)
        {
            if (aStaffSpec.QLId != "")
            {
                //This is done in the staffSpec now; SJL 16/11/2011
                //String staffGroup = ".STA_0ther.VDI_GROUPS.LSPA.WC";//default group
                //String staffStartChar = aStaffSpec.NDSName.ToString().Trim().Substring(0, 1).ToUpper();
                //if ("ABCDEFGHIJKLMNOPQRSTUVWXYZ".Contains(staffStartChar))
                //{
                //    staffGroup = ".STA_" + staffStartChar.ToString() + ".VDI_GROUPS.LSPA.WC";
                //}
                String FullName = aStaffSpec.Forename + " " + aStaffSpec.Surname;
                String Password = "seaside";
                String ExpireDate = "none";
                String Description = "Created on " + DateTime.Now;
                String wNewStaffJRBText =
                    // "!Template=" + aStaffSpec.Template + "\n" +
                    "!Name context = " + "users" + aStaffSpec.Context + "\n"; 
                    //"!Home directory volume=" + aStaffSpec.HomeVol + "\n" +
                    //"!Home directory path=users\n";
                   // "!Second home directory volume=" + aStaffSpec.SecondHomeVol + "\n" +
                   // "!Second home directory path=users\n";
                if (aStaffSpec.GWise)
                    wNewStaffJRBText +=
                        "!Groupwise add users=y" + "\n" +
                        "!Groupwise domain object=.warkscollege.groupwise.wc" + "\n" +
                        "!Groupwise post office=" + aStaffSpec.PostOffice + "\n";
                wNewStaffJRBText +=
                    "\"" + aStaffSpec.NDSName + "\"" + "," + "\"" + aStaffSpec.Surname.ToUpper() + "\"" +
                    "," + "\"" + aStaffSpec.Forename.ToUpper() + "\"" + "," + "\"" + Password + "\"" +
                    "," + "\"" + FullName.ToUpper() + "\"" + "," + "\"" + aStaffSpec.Department.ToUpper() + "\"" +
                    //"," + "\"" + aStaffSpec.Site.ToUpper() + "\"" + 
                    "," + "\"" + aStaffSpec.JobTitle.ToUpper() + "\"" +
                    "," + "\"" + aStaffSpec.Tel + "\"" + 
                    "," + "\"" + aStaffSpec.EmailAddress.ToUpper() + "\"" +
                    "," + "\"" + aStaffSpec.EmailAddressGoogle.ToUpper() + "\"" +
                    "," + "\"" + Description + "\"" +
                    "," + "\"" + aStaffSpec.StaffID + "\"" +
                    "," + "\"" + ExpireDate + "\"" +
                    //"," + "\"" + aStaffSpec.HomeVolRestrict + "\"" +
                    //"," + "\"" + aStaffSpec.SharedVolRestrict + "\"" +
                    //"," + "\"" + aStaffSpec.Vol1VolRestrict + "\"" +
                    "," + "\"" + aStaffSpec.GroupMembership.ToString() + "\"";

                return wNewStaffJRBText;
            }
            else return "";
        }

        protected override void testSelectedPhase(staffSpecSL aStaffSpecSL, DateTime runDateTime)
        {

            DataView wDV;
            String sqlInsert;
            String LD = "FALSE";//set default FALSE for Attribute value LoginDisabled  does not exist, this means that the account is enabled.
            foreach (staffSpec ss in aStaffSpecSL.Values)
            {
                if ((mAttributeValues = GetLDAPInfo("cn=" + ss.NDSName)) != null)
                {

                    if (attributeValue("surname") != null)
                    {
                        String[] groups = attributeArray("grpmbr");
                        Boolean InGroups = false;
                        if (groups != null)
                        {
                            if (groups.Length > 1)
                            {
                                if ((groups[0].ToLower().Contains("allsites") || groups[0].ToLower().Contains("vdi_groups")) && ((groups[1].ToLower().Contains("allsites") || groups[1].ToLower().Contains("vdi_groups"))))
                                {
                                    InGroups = true;
                                }
                            }
                        }

                        if (InGroups == false)
                        {
                            // create a coupler job to updateNDS
                            staffUtility.writeCouplerMessageQueueV2(ss.NDSName, "", "UpdateNDS");

                        }
                    }

                    if (mAttributeValues.ContainsKey("logindisabled")) LD = "TRUE";//attribute value LoginDisabled exists
                    if (attributeValue("logindisabled").ToLower() == "false" || LD == "FALSE")
                    {
                        if (ss.EmpType.ToString().ToUpper()=="TEMP")mTestedOK.Add(ss.queueItem.ToString());//TEMP accounts do not have a PhoneList record.
                        staffUtility.updateWcStaffIdentity(ss.NDSName, "SET NDSdisabled = 0");

                        mTestedOK.Add(ss.queueItem.ToString());
                        //staffUtility.writeCouplerMessageQueueV2(ss.NDSName, "", "CreateHomeDirectory");// Write a coupler job to Enable an exchange account
                        staffUtility.writeCouplerMessageQueueV2(ss.NDSName, "", "ExchangeEnable");// Write a coupler job to Enable an exchange account
                        staffUtility.writeCouplerMessageQueueV2(ss.NDSName, "", "CreateLearningChannel");// Write a coupler job to create a learning channel account
                        staffUtility.writeCouplerMessageQueueV2(ss.NDSName, "", "CreateGoogleLogins");// Write a coupler job to create a Google account
                        bool wB = staffUtility.sendEmail(
                        "slangley@warkscol.ac.uk",
                        "idm@warkscol.ac.uk",
                        "Staff Creation : " + ss.NDSName,
                        "( " + ss.NDSName + " ) has been created in NDS. Please link user in to EXCHANGE.", null);

                        sqlInsert = "wcIDMPhoneListUpdateEmail " + ss.QLId.ToString();
                        wDV = staffUtility.readDataView(staffUtility.couplerDB, sqlInsert);
                        //if (wDV.Count > 0)
                        //{
                        //    if (wDV[0][0].ToString() == "1")
                        //    {
                        //    }
                        //}
                    }
                    else
                        if (ss.attempts > mTryCount) mFailed.Add(ss.queueItem.ToString());
                }
                else
                    if (ss.attempts > mTryCount) mFailed.Add(ss.queueItem.ToString());
            }
        }

    }

    class jrbCouplerUpdate : jrbCoupler
    {
        public jrbCouplerUpdate
            (String aJRBFilesPath, String aActionName, int aTryCount)
            : base(aJRBFilesPath, aActionName, aTryCount) { }

        protected override String doSelectedItem(staffSpec aStaffSpec)
        {
            if (aStaffSpec.QLId != "")
            {
                String staffGroup = ".STA_0ther.VDI_GROUPS.LSPA.WC";//default group
                String staffStartChar = aStaffSpec.NDSName.ToString().Trim().Substring(0, 1).ToUpper();
                if ("ABCDEFGHIJKLMNOPQRSTUVWXYZ".Contains(staffStartChar))
                {
                    staffGroup = ".STA_" + staffStartChar.ToString() + ".VDI_GROUPS.LSPA.WC";
                }
                String FullName = aStaffSpec.Forename + " " + aStaffSpec.Surname;
                String Password = "seaside";
                String ExpireDate = "none";
                String Description = "Created on " + DateTime.Now;
                String wNewStaffJRBText = "";
                wNewStaffJRBText +=
                    "\"" + aStaffSpec.NDSName + "\"" + "," + "\"" + staffGroup.ToString() + "\"";
                return wNewStaffJRBText;
            }
            else return "";
        }

        protected override void testSelectedPhase(staffSpecSL aStaffSpecSL, DateTime runDateTime)
        {

            DataView wDV;
            String sqlInsert;
            String LD = "FALSE";//set default FALSE for Attribute value LoginDisabled  does not exist, this means that the account is enabled.
            foreach (staffSpec ss in aStaffSpecSL.Values)
            {
                if ((mAttributeValues = GetLDAPInfo("cn=" + ss.NDSName.Trim())) != null)
                {
                    String x = attributeValue("gwexpire").ToString();
                    String[] groups = attributeArray("grpmbr");
                    Boolean InGroups = false;
                    Boolean allsites = false;
                    Boolean vdi = false;
                    if (groups != null)
                    {
                        if (groups.Length > 1)
                        {
                            for (int i = 0; i < groups.Length; i++)
                            {
                                if (groups[i].ToLower().Contains("allsites"))
                                {
                                    allsites = true;  
                                }
                                if (groups[i].ToLower().Contains("vdi_groups"))
                                {
                                    vdi = true;
                                }
                                if (allsites && vdi) InGroups = true;

                            }
                        }
                    }
                    if (attributeValue("logindisabled").ToLower() == "false" && InGroups == true)
                    {
                        //studentUtility.updateWcStaffIdentity(ss.NDSName, "SET NDSdisabled = 0");
                        mTestedOK.Add(ss.queueItem.ToString());
                    }
                    else
                        if (ss.attempts > mTryCount) mFailed.Add(ss.queueItem.ToString());
                }
                else
                    if (ss.attempts > mTryCount) mFailed.Add(ss.queueItem.ToString());
            }
        }

    }

    
    class jrbCouplerEnable : jrbCoupler
    {
        public jrbCouplerEnable
            (String aJRBFilesPath, String aActionName, int aTryCount)
            : base(aJRBFilesPath, aActionName, aTryCount) { }

        protected override String doSelectedItem(staffSpec aStaffSpec)
        {
            if (aStaffSpec.QLId != "")
            {
                String FullName = aStaffSpec.Forename.Trim() + " " + aStaffSpec.Surname.Trim();
                String ExpireDate = "none";
                String Password = ""; //default
                String[] AD = aStaffSpec.ActionData.Split('=');
                if (AD[0].ToString() == "Password")  // generated by StaffActivate
                {
                    Password = AD[1].ToString();
                }
                String Description = "Enabled=" + aStaffSpec.queueItem.ToString().Trim() + "@" + DateTime.Now;
                String wNewStaffJRBText =
                    // "!Template=" + aStaffSpec.Template + "\n" +
                    "!Name context = " + "users" + aStaffSpec.Context + "\n";
                wNewStaffJRBText +=
                    "\"" + aStaffSpec.NDSName + "\"" + "," + "\"" + aStaffSpec.Surname.ToUpper() + "\"" +
                    "," + "\"" + aStaffSpec.Forename.ToUpper() + "\"" +
                    "," + "\"" + Password + "\"" +
                    "," + "\"" + FullName.ToUpper() + "\"" +
                    "," + "\"" + aStaffSpec.Department.ToUpper() + "\"" +
                    //"," + "\"" + aStaffSpec.Site.ToUpper() + "\"" + 
                    "," + "\"" + aStaffSpec.JobTitle.ToUpper() + "\"" +
                    "," + "\"" + aStaffSpec.Tel + "\"" + "," + "\"" + aStaffSpec.EmailAddress.ToUpper() + "\"" +
                    "," + "\"" + Description + "\"" +
                    "," + "\"" + aStaffSpec.StaffID + "\"" +
                    "," + "\"" + aStaffSpec.GroupMembership + "\"" +
                    "," + "\"" + ExpireDate + "\"";
                //"," + "\"" + aStaffSpec.HomeVolRestrict + "\"" +
                //"," + "\"" + aStaffSpec.SharedVolRestrict + "\"" +
                //"," + "\"" + aStaffSpec.Vol1VolRestrict + "\"";

                return wNewStaffJRBText;
            }
            else return "";
        }

        protected override void testSelectedPhase(staffSpecSL aStaffSpecSL, DateTime runDateTime)
        {
            DataView wDV;
            String sqlInsert;
            foreach (staffSpec ss in aStaffSpecSL.Values)
            {
                mAttributeValues = GetLDAPInfo("cn=" + ss.NDSName);
                string logEx = attributeValue("loginexpirationtime").ToString();
                string description = "Enabled=" + ss.queueItem.ToString().Trim(); // set the test value for NDS
                // the decription from NDS should have the queueItem after a successful enable e.g. "Enabled=10817@07/11/2010 20:18".
                // split the NDS description to get the queuItem
                string[] descriptions = attributeValue("description").ToString().Trim().Split('@');

                String LD = "FALSE";//set default FALSE for Attribute value LoginDisabled  does not exist, this means that the account is enabled.
                if (mAttributeValues.ContainsKey("logindisabled")) LD = "TRUE";//attribute value LoginDisabled exists
                if (!ss.WhenRead) staffUtility.writeCouplerMessageQueueV2(ss.NDSName, "", "ExchangeConnect");// Create a coupler job for re-enabling Exchange.
                if (!ss.WhenRead) staffUtility.writeCouplerMessageQueueV2(ss.NDSName, "", "RestoreGoogle");// Write a coupler job to create a Google account

                if ((attributeValue("logindisabled").ToLower() == "false" || LD == "FALSE") && attributeValue("generationqualifier").ToString().Trim() == ss.StaffID.ToString().Trim() && descriptions[0].ToString().Trim() == description)
                {
                    staffUtility.updateWcStaffIdentity(ss.NDSName, "SET NDSdisabled = 0");

                    staffUtility.wcIDMReSetStaffIdentity(ss.NDSName);
                    sqlInsert = "wcIDMEnableDisableApplications " + ss.QLId.ToString() + ",'" + ss.NDSName.ToString() + "'," + "'T','T','Y','Y'";

                    wDV = staffUtility.readDataView(staffUtility.couplerDB, sqlInsert);
                    if (wDV.Count > 0)
                    {
                        if ((wDV[0][0].ToString() == "1" && wDV[0][1].ToString() == "1" && 
                             wDV[0][2].ToString() == "1" && wDV[0][3].ToString() == "1") || 
                            (wDV[0][0].ToString() == "1" && wDV[0][1].ToString() == "1" && 
                             wDV[0][2].ToString() == "1" && wDV[0][3].ToString() == "3" && 
                             ss.EmpType.ToString().ToUpper() == "TEMP"))
                        {
                            mTestedOK.Add(ss.queueItem.ToString());
                        }
                    }
                }
                else
                    if (ss.attempts > mTryCount) mFailed.Add(ss.queueItem.ToString());
            }
        }

    }

    class jrbCouplerDisable : jrbCoupler
    {
        public jrbCouplerDisable
            (String aJRBFilesPath, String aActionName, int aTryCount)
            : base(aJRBFilesPath, aActionName, aTryCount) { }

        protected override String doSelectedItem(staffSpec aStaffSpec)
        {
            if (! aStaffSpec.JobTitle.ToString().ToLower().Contains("governor"))
            {

                if (aStaffSpec.QLId != "")
                {
                    String Password = ""; //default
                    String Description = "Disabled on " + DateTime.Now;
                    String[] AD = aStaffSpec.ActionData.Split('=');
                    if (AD[0].ToString() == "Password")  // generated by StaffActivate
                    {
                        Password = AD[1].ToString();
                        if (Password != "") Description = "Trashed on " + DateTime.Now;
                        // write a coupler job to ClearNDSGroups
                        //write a clearNDS groups coupler job
                        //
                        staffUtility.writeCouplerMessageQueueV2(aStaffSpec.NDSName, "", "ClearNDSGroups");

                    }
                    String wNewStaffJRBText =
                        "\"" + aStaffSpec.NDSName + "\"" +
                        "," + "\"" + Password + "\"" +
                       // "," + "\"" + aStaffSpec.NDSName.Trim() + "\"" +
                        "," + "\"" + Description + "\"";
                    return wNewStaffJRBText;
                }

                else return "";
            }
            else return "";
        }

        protected override void testSelectedPhase(staffSpecSL aStaffSpecSL, DateTime runDateTime)
        {
            DataView wDV;
            String sqlInsert;
            foreach (staffSpec ss in aStaffSpecSL.Values)
            {
                mAttributeValues = GetLDAPInfo("cn=" + ss.NDSName);

                if (attributeValue("logindisabled").ToLower() == "true")
                {
                            staffUtility.updateWcStaffIdentity(ss.NDSName, "SET NDSdisabled = 1");
                            // update applications to disabled 
                            // SHD, EHD, QLR, phone_list
                            // Disable Lync first then the AD processing can create a coupler job to disable excahnge
                            if (ss.keepProxy == false) staffUtility.writeCouplerMessageQueueV2(ss.NDSName, "", "ExchangeLyncDisable");
                            if (ss.keepProxy == false) staffUtility.writeCouplerMessageQueueV2(ss.NDSName, "", "SuspendGoogle");
                            sqlInsert = "wcIDMEnableDisableApplications " + ss.QLId.ToString() + ",'" + ss.NDSName.ToString() + "'," + "'F','F','N','N'";

                            wDV = staffUtility.readDataView(staffUtility.couplerDB, sqlInsert);
                            if (wDV.Count > 0)
                            {
                                if (((wDV[0][0].ToString() == "1" || wDV[0][0].ToString() == "3") && 
                                     (wDV[0][1].ToString() == "1" || wDV[0][1].ToString() == "3") && 
                                     (wDV[0][2].ToString() == "1" || wDV[0][2].ToString() == "3") && 
                                     (wDV[0][3].ToString() == "1" || wDV[0][3].ToString() == "3")) ||
                                     ((wDV[0][0].ToString() == "1" || wDV[0][0].ToString() == "3") &&
                                     (wDV[0][1].ToString() == "1" || wDV[0][1].ToString() == "3") &&
                                     (wDV[0][2].ToString() == "1" || wDV[0][2].ToString() == "3") && 
                                      ss.EmpType.ToString().ToUpper() == "TEMP"))
                                {
                                    mTestedOK.Add(ss.queueItem.ToString());
                                    bool wB = staffUtility.sendEmail(
                                        "slangley@warkscol.ac.uk",
                                        "idm@warkscol.ac.uk",
                                        "Staff Disable : " + ss.NDSName,
                                        "( " + ss.NDSName + " ) has been disabled in NDS. Please disable user in EXCHANGE.", null);

                                }
                            }
                }
                else
                    if (ss.attempts > mTryCount) mFailed.Add(ss.queueItem.ToString());
            }
        }
    }

    class jrbCouplerHomeDir : jrbCoupler
    {
        // put on hold and not fully implemented yet 
        public jrbCouplerHomeDir
            (String aJRBFilesPath, String aActionName, int aTryCount)
            : base(aJRBFilesPath, aActionName, aTryCount) { }
        //this will call the sethome2 utility which will create a new directory
        //but will also set this to the home directory, will need to change this back later on once finished
        protected override String doSelectedItem(staffSpec aStaffSpec)
        {
            if (aStaffSpec.QLId != "")
            {
                String Password = ""; //default
                String Description = "Disabled on " + DateTime.Now;
                String[] AD = aStaffSpec.ActionData.Split('=');
                if (AD[0].ToString() == "Password")  // generated by StaffActivate
                {
                    Password = AD[1].ToString();
                    if (Password != "") Description = "Trashed on " + DateTime.Now;
                    // write a coupler job to ClearNDSGroups
                    //write a clearNDS groups coupler job
                    //
                    staffUtility.writeCouplerMessageQueueV2(aStaffSpec.NDSName, "", "ClearNDSGroups");

                }
                String wNewStaffJRBText =
                    "\"" + aStaffSpec.NDSName + "\"" +
                    "," + "\"" + Password + "\"" +
                    "," + "\"" + Description + "\"";
                return wNewStaffJRBText;
            }
            else return "";
        }

        protected override void testSelectedPhase(staffSpecSL aStaffSpecSL, DateTime runDateTime)
        {
            // put on hold and not fully implemented yet 
            DataView wDV;
            String sqlInsert;
            foreach (staffSpec ss in aStaffSpecSL.Values)
            {
                mAttributeValues = GetLDAPInfo("cn=" + ss.NDSName);
                String homeDir = attributeValue("ndshomedirectory").ToLower();
                String[] hd = homeDir.Split(',');
                //if (homeDir.Length > 0) hd = homeDir.Split(',');
                //test for example cn=orion_home
                if ( hd[0].ToString().Trim() == ss.ActionData.ToString().ToLower().Trim())
                {
                    //sqlInsert = "wcIDMEnableDisableApplications " + ss.QLId.ToString() + ",'" + ss.NDSName.ToString() + "'," + "'F','F','N','N'";

                    //wDV = staffUtility.readDataView(staffUtility.couplerDB, sqlInsert);
                    //if (wDV[0][0].ToString() == "1" && wDV[0][1].ToString() == "1" && wDV[0][2].ToString() == "1" && (wDV[0][3].ToString() == "1" || wDV[0][3].ToString() == "3"))
                   // {
                        mTestedOK.Add(ss.queueItem.ToString());
                    //}
                }
                else
                    if (ss.attempts > mTryCount) mFailed.Add(ss.queueItem.ToString());
            }
        }
    }

    class jrbCouplerReSetHomeDir : jrbCoupler
    {
        // put on hold and not fully implemented yet 
        public jrbCouplerReSetHomeDir
            (String aJRBFilesPath, String aActionName, int aTryCount)
            : base(aJRBFilesPath, aActionName, aTryCount) { }
        //ResEt the HOme directory back to original
        //when using the sethome2 utility the homedirectory is set
        //jrbCouplerHomeDir has set the home directory to something we dont't want so reset it back

        protected override String doSelectedItem(staffSpec aStaffSpec)
        {
            if (aStaffSpec.QLId != "")
            {
                String Password = ""; //default
                String Description = "Disabled on " + DateTime.Now;
                String[] AD = aStaffSpec.ActionData.Split('=');
                if (AD[0].ToString() == "Password")  // generated by StaffActivate
                {
                    Password = AD[1].ToString();
                    if (Password != "") Description = "Trashed on " + DateTime.Now;
                    // write a coupler job to ClearNDSGroups
                    //write a clearNDS groups coupler job
                    //
                    staffUtility.writeCouplerMessageQueueV2(aStaffSpec.NDSName, "", "ClearNDSGroups");

                }
                String wNewStaffJRBText =
                    "\"" + aStaffSpec.NDSName + "\"" +
                    "," + "\"" + Password + "\"" +
                    "," + "\"" + Description + "\"";
                return wNewStaffJRBText;
            }
            else return "";
        }

        protected override void testSelectedPhase(staffSpecSL aStaffSpecSL, DateTime runDateTime)
        {
            DataView wDV;
            String sqlInsert;
            foreach (staffSpec ss in aStaffSpecSL.Values)
            {
                mAttributeValues = GetLDAPInfo("cn=" + ss.NDSName);
                if (attributeValue("logindisabled").ToLower() == "true")
                {
                    staffUtility.updateWcStaffIdentity(ss.NDSName, "SET NDSdisabled = 1");
                    // update applications to disabled
                    // SHD, EHD, QLR, phone_list
                    sqlInsert = "wcIDMEnableDisableApplications " + ss.QLId.ToString() + ",'" + ss.NDSName.ToString() + "'," + "'F','F','N','N'";

                    wDV = staffUtility.readDataView(staffUtility.couplerDB, sqlInsert);
                    if (wDV[0][0].ToString() == "1" && wDV[0][1].ToString() == "1" && wDV[0][2].ToString() == "1" && (wDV[0][3].ToString() == "1" || wDV[0][3].ToString() == "3"))
                    {
                        mTestedOK.Add(ss.queueItem.ToString());
                    }
                }
                else
                    if (ss.attempts > mTryCount) mFailed.Add(ss.queueItem.ToString());
            }
        }
    }

    class jrbCouplerTrash : jrbCoupler
    {
        public jrbCouplerTrash
            (String aJRBFilesPath, String aActionName, int aTryCount)
            : base(aJRBFilesPath, aActionName, aTryCount) { }

        protected override String doSelectedItem(staffSpec aStaffSpec)
        {
            if (!aStaffSpec.JobTitle.ToString().ToLower().Contains("governor"))
            {
                if (aStaffSpec.QLId != "")
                {
                    String Password = ""; //default
                    String Description = "Disabled on " + DateTime.Now;
                    String[] AD = aStaffSpec.ActionData.Split('=');
                    if (AD[0].ToString() == "Password")  // generated by StaffActivate
                    {
                        Password = AD[1].ToString();
                        if (Password != "") Description = "Trashed on " + DateTime.Now;
                        // write a coupler job to ClearNDSGroups
                        //write a clearNDS groups coupler job
                        //
                        staffUtility.writeCouplerMessageQueueV2(aStaffSpec.NDSName, "", "ClearNDSGroups");

                    }
                    String wNewStaffJRBText =
                        "\"" + aStaffSpec.NDSName + "\"" +
                        "," + "\"" + Password + "\"" +
                        "," + "\"" + Description + "\"";
                    return wNewStaffJRBText;
                }
                else return "";
            }
            else return "";
        }

        protected override void testSelectedPhase(staffSpecSL aStaffSpecSL, DateTime runDateTime)
        {
            foreach (staffSpec ss in aStaffSpecSL.Values)
            {
                mAttributeValues = GetLDAPInfo("cn=" + ss.NDSName);
                if (attributeValue("logindisabled").ToLower() == "true" && attributeValue("description").ToLower().Contains( "trashed"))
                {
                    //staffUtility.updateWcStaffIdentity(ss.NDSName, "SET NDSdisabled = 1");
                    mTestedOK.Add(ss.queueItem.ToString());
                    bool wB = staffUtility.sendEmail(
                                    "slangley@warkscol.ac.uk",
                                    "idm@warkscol.ac.uk",
                                    "Staff Trashed : " + ss.NDSName,
                                    "( " + ss.NDSName + " ) has been trashed in NDS. Please ensure user is disabled in EXCHANGE.", null);

                }
                else
                    if (ss.attempts > mTryCount) mFailed.Add(ss.queueItem.ToString());
            }
        }

    }

    class jrbCouplerClearNDSGroups : jrbCoupler
    {
        public jrbCouplerClearNDSGroups
            (String aJRBFilesPath, String aActionName, int aTryCount)
            : base(aJRBFilesPath, aActionName, aTryCount) { }

        protected override String doSelectedItem(staffSpec aStaffSpec)
        {
            if (aStaffSpec.QLId != "")
            {
                //String Password = ""; //default
                //String Description = "Disabled on " + DateTime.Now;
                //String[] AD = aStaffSpec.ActionData.Split('=');
                //if (AD[0].ToString() == "Password")  // generated by StaffActivate
                //{
                //    Password = AD[1].ToString();
                //    if (Password != "") Description = "Trashed on " + DateTime.Now;
                //}
                //String sGroups = "\"" + "._ALL-LEAMINGTON-STAFF.GROUPWISE.WC" + "\""
                //         + "^" + "\"" + "._ALL-PERSHORE-STAFF.GROUPWISE.WC" + "\"" 
                //         + "^" + "\"" + "._ALL-MORETON-STAFF.GROUPWISE.WC" + "\"" 
                //         + "^" + "\"" + "._ALL-RUGBY-STAFF.GROUPWISE.WC" + "\"" 
                //         + "^" + "\"" + "._ALL-HENLEY-STAFF.GROUPWISE.WC" + "\"" 
                //         + "^" + "\"" + "._ALL-TRIDENT-STAFF.GROUPWISE.WC" + "\"";
                String wNewStaffJRBText =
                    "\"" + aStaffSpec.NDSName + "\"";
                    //+ "," + sGroups;
                return wNewStaffJRBText;
            }
            else return "";
        }

        protected override void testSelectedPhase(staffSpecSL aStaffSpecSL, DateTime runDateTime)
        {

            //
            foreach (staffSpec ss in aStaffSpecSL.Values)
            {
                bool groupsDeleted = true;
                mAttributeValues = GetLDAPInfo("cn=" + ss.NDSName);
                if (attributeArray("grpmbr") != null)
                {
                    if (attributeArray("grpmbr").Length > 0) // user is a member of a group
                    {
                        groupsDeleted = false;

                        // commented out because we are now interested in any group.
                        //string[] groups = attributeArray("grpmbr");
                        //foreach (string group in groups)
                        //{
                        //    if (group.ToLower().IndexOf("ou=groupwise,o=wc") > 1)
                        //    {
                        //        // this means that this user is still a member of an email group.
                        //        groupsDeleted = false;
                        //    }
                        //}
                    }
                }
                if (groupsDeleted) mTestedOK.Add(ss.queueItem.ToString());
                else
                    if (ss.attempts > mTryCount) mFailed.Add(ss.queueItem.ToString());
            }
        }

    }




    class jrbCouplerDelete : jrbCoupler
    {
        public jrbCouplerDelete
            (String aJRBFilesPath, String aActionName, int aTryCount)
            : base(aJRBFilesPath, aActionName, aTryCount) { }

        protected override String doSelectedItem(staffSpec aStaffSpec)
        {
            //DateTime theDate = DateTime.Parse("31 Dec 2006 21:00:00");
            //TimeSpan timeOnly = theDate.TimeOfDay;
            //DateTime theDateNow = DateTime.Now;
            //TimeSpan timeOnlyNow = theDateNow.TimeOfDay;
            //bool result;
            //result = timeOnlyNow > timeOnly;
            //if (result)
            //{

                if (aStaffSpec.QLId != "")
                {
                    //String Password = ""; //default
                    //String Description = "Disabled on " + DateTime.Now;
                    String[] AD = aStaffSpec.ActionData.Split('=');
                    //if (AD[0].ToString() == "Password")  // generated by StaffActivate
                    //{
                    //    Password = AD[1].ToString();
                    //    if (Password != "") Description = "Trashed on " + DateTime.Now;
                    //}
                    String wNewStaffJRBText =
                        "\"" + aStaffSpec.NDSName + "\"";
                    return wNewStaffJRBText;
                }
                else return "";

            //}
            //else return "";
        }

        protected override void testSelectedPhase(staffSpecSL aStaffSpecSL, DateTime runDateTime)
        {

            //if nds account not exist then delete is true
            //write a defaultDelete job in the coupler with updateCouplerMessageQueueSet() ?
            //
            //DateTime theDate = DateTime.Parse("31 Dec 2006 21:00:00");
            //TimeSpan timeOnly = theDate.TimeOfDay;
            //DateTime theDateNow = DateTime.Now;
            //TimeSpan timeOnlyNow = theDateNow.TimeOfDay;
            //bool result;
            //result = timeOnlyNow > timeOnly;
            //if (result)
            //{
            
                foreach (staffSpec ss in aStaffSpecSL.Values)
                {
                    mAttributeValues = GetLDAPInfo("cn=" + ss.NDSName);

                    if (mAttributeValues == null)
                    //if (attributeValue("logindisabled").ToLower() == "true")
                    {
                        staffUtility.writeCouplerMessageQueue(ss.NDSName, ss.QLId);
                        staffUtility.updateWcStaffIdentity(ss.NDSName, "SET job_Title = DELETED NDS record");
                        mTestedOK.Add(ss.queueItem.ToString());
                    }
                    else
                        if (ss.attempts > mTryCount) mFailed.Add(ss.queueItem.ToString());
                }
            //}
        }

    }
}



