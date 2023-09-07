using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
//using Google.GData.Apps;

using Google.Apis.Download;
using Google.Apis.Services;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Util.Store;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using Google.Apis.Util;
using Google.Apis.Discovery;
using Google.Apis.Json;
using Google.Apis.Admin.Directory.directory_v1;
using Google.Apis.Admin.Directory.directory_v1.Data;
using Google.GData.Apps;
using Google.GData.Apps.GoogleMailSettings;
using Google.GData.Client;
using Google.GData.Extensions;
using Google.GData.Contacts;
using Google.Contacts;

namespace couplerWriter
{
    class defaultCoupler : coupler
    {
        protected String mCSSHDDB, mCSEHDDB, mCSQLRDB;

        public defaultCoupler() : base() { }

        public DirectoryService wcService;
        public ContactsService wcCService;
        public static string redirectUri = "urn:ietf:wg:oauth:2.0:oob";
        public string scopes = "https://apps-apis.google.com/a/feeds/emailsettings/2.0/";

        //public AppsService service;
        //public AppsService serviceTEST;
        //public OrganizationService serviceORG;

        public defaultCoupler
            (String csEHDDB, String csSHDDB, String csQLRDB, String aActionName, int aTryCount)
        {
            mCSEHDDB = csEHDDB; mCSQLRDB = csQLRDB; mCSSHDDB = csSHDDB;
            mActionName = aActionName; mTryCount = aTryCount;
            try
            {

                //OAuth2Parameters parameters = new OAuth2Parameters()
                //{
                //    ClientId = "340413376411-qdqqhtcfcj7s6rm40qkerdkv3svnqcfu@developer.gserviceaccount.com",
                //    ClientSecret = "notasecret",
                //    RedirectUri = redirectUri,
                //    Scope = scopes
                //};

                //string url = OAuthUtil.CreateOAuth2AuthorizationUrl(parameters);
                //Console.WriteLine("Authorize URI: " + url);
                //parameters.AccessCode = Console.ReadLine();

                //OAuthUtil.GetAccessToken(parameters);
                //GOAuth2RequestFactory requestFactory = new GOAuth2RequestFactory("apps", "CouplerWriter", parameters);

                //GoogleMailSettingsService MISservice = new GoogleMailSettingsService("warwickshire.ac.uk", "CouplerWriter");
                //MISservice.RequestFactory = requestFactory;

                //MISservice.setUserCredentials("system@warwickshire.ac.ik", "backbeat");
                //AppsExtendedEntry entry = MISservice.CreateDelegate("slangley@warwickshire.ac.uk", "cbuck@warwickshire.ac.uk");

                // original email and security key - not working after InfraStructure deleted the account @20/03/2017
                //string SERVICE_ACCOUNT_EMAIL = "340413376411-qdqqhtcfcj7s6rm40qkerdkv3svnqcfu@developer.gserviceaccount.com";
                //string SERVICE_ACCOUNT_PKCS12_FILE_PATH = @"c:\wcIdentities\API Project-77490cc65f92.p12";

                string SERVICE_ACCOUNT_EMAIL = "mis-system@api-project-340413376411.warwickshire.ac.uk.iam.gserviceaccount.com";
                string SERVICE_ACCOUNT_PKCS12_FILE_PATH = @"c:\wcIdentities\API Project-38bf2cd104b0.p12";
                // Create the service.
                
                X509Certificate2 certificate = new X509Certificate2(SERVICE_ACCOUNT_PKCS12_FILE_PATH, "notasecret", X509KeyStorageFlags.Exportable);
                //DBETHELL foud that after a server update he had to use the pipe to | X509KeyStorageFlags.MachineKeySet | X509KeyStorageFlags.PersistKeySet SEE BELOW
                //X509Certificate2 certificate = new X509Certificate2(SERVICE_ACCOUNT_PKCS12_FILE_PATH, "notasecret", X509KeyStorageFlags.Exportable | X509KeyStorageFlags.MachineKeySet | X509KeyStorageFlags.PersistKeySet);

                ServiceAccountCredential credential = new ServiceAccountCredential(
                           new ServiceAccountCredential.Initializer(SERVICE_ACCOUNT_EMAIL)
                           {

                               Scopes = new[] { DirectoryService.Scope.AdminDirectoryUser, DirectoryService.Scope.AdminDirectoryGroupMember}
                              ,User = "system@warwickshire.ac.uk"
                            }.FromCertificate(certificate));
              

                wcService = new DirectoryService(new BaseClientService.Initializer()
                                {
                    HttpClientInitializer = credential,
                    ApplicationName = "User Provisioning"
                    
                });

               
             }
            catch (Exception gEx)
            {
                String gMsg = gEx.Message.ToString();
            }
        }

        protected void updateSelectedItem(staffSpec aStaffSpec, String aEHDDBActiveFlag, String aSHDDBActiveFlag, String aQLRActiveFlag)
        {
            String HDInsert =
                "wcIDMSTFCreateEHDAndSHD " +
                "'" + aStaffSpec.EmailAddress + "', " +
                "'" + aStaffSpec.Site + "', " +
                "'" + aStaffSpec.Forename.Replace("'", "''") + "', " +
                "'" + aStaffSpec.Surname.Replace("'", "''") + "', " +
                "'" + aStaffSpec.QLId + "', " +
                "'" + aStaffSpec.Department.Replace("'", "''") + "', " +
                "'" + aStaffSpec.Tel + "', " +
                "'" + aStaffSpec.EmpType + "', " +
                "'" + aEHDDBActiveFlag + "', " +
                "'" + aSHDDBActiveFlag + "'";
            DataView wDV;
            //EHDDB + SHDDB
            wDV = staffUtility.readDataView(staffUtility.XrayUtilDB, HDInsert);
            //QLR
            wDV = staffUtility.readDataView(staffUtility.couplerDB,
                "wcIDMSTFCreateQLR " +
                "'" + aStaffSpec.NDSName + "', " +
                "'" + aStaffSpec.Site + "', " +
                "'" + aStaffSpec.QLId + "', " +
                "'" + aStaffSpec.DeptCode + "', " +
                "'" + aStaffSpec.EmpType + "', " +
                "'" + aQLRActiveFlag + "'"
            );
            mWritten.Add(aStaffSpec.queueItem.ToString());
        }

        protected SortedList<String, String> testSelectedItemData(staffSpec aStaffSpec)
        {
            SortedList<String, String> wSL = new SortedList<String, String>();
            DataView wDV;
            //EHDDB + SHDDB
            wDV = staffUtility.readDataView(staffUtility.XrayUtilDB,
                "wcIDMSTFTestEHDAndSHD " +
                "'" + aStaffSpec.EmailAddress + "', " +
                "'" + aStaffSpec.Site + "', " +
                "'" + aStaffSpec.Forename.Replace("'", "''") + "', " +
                "'" + aStaffSpec.Surname.Replace("'", "''") + "', " +
                "'" + aStaffSpec.QLId + "', " +
                "'" + aStaffSpec.Department.Replace("'", "''") + "', " +
                "'" + aStaffSpec.Tel + "', " +
                "'" + aStaffSpec.EmpType + "'"
                );
            //QLR
            wSL.Add("EHDDBActive", wDV[0]["EHDDBActive"].ToString());
            wSL.Add("EHDDBData", wDV[0]["EHDDBData"].ToString());
            wSL.Add("SHDDBActive", wDV[0]["SHDDBActive"].ToString());
            wSL.Add("SHDDBData", wDV[0]["SHDDBData"].ToString());
            wDV = staffUtility.readDataView(staffUtility.couplerDB,
                "wcIDMSTFTestQLR " +
                "'" + aStaffSpec.NDSName + "', " +
                "'" + aStaffSpec.Site + "', " +
                "'" + aStaffSpec.QLId + "', " +
                "'" + aStaffSpec.DeptCode + "', " +
                "'" + aStaffSpec.EmpType + "'"
            );
            wSL.Add("QLRActive", wDV[0]["QLRActive"].ToString());
            wSL.Add("QLRData", wDV[0]["QLRData"].ToString());
            return (wSL);
        }

        protected virtual Boolean testSelectedItem(staffSpec aStaffSpec, String aEHDDBActiveFlag, String aSHDDBActiveFlag, String aQLRActiveFlag)
        {
            SortedList<String, String> wSL = testSelectedItemData(aStaffSpec);
            return (
                ((aEHDDBActiveFlag=="%")||(wSL["EHDDBActive"] == aEHDDBActiveFlag)) &&
                ((aSHDDBActiveFlag=="%")||(wSL["SHDDBActive"] == aSHDDBActiveFlag)) &&
                ((aQLRActiveFlag=="%")||(wSL["QLRActive"] == aQLRActiveFlag)) &&
                (wSL["EHDDBData"]=="OK") &&
                (wSL["SHDDBData"]=="OK") &&
                (wSL["QLRData"]=="OK")
              );
        }

        protected virtual void doSelectedItem(staffSpec aStaffSpec) { }

        protected override void doSelectedPhase(staffSpecSL aStaffSpecSL, DateTime runDateTime)
        {
            if (aStaffSpecSL.Count > 0)
                foreach (staffSpec aSS in aStaffSpecSL.Values)
                    doSelectedItem(aSS);
        }

        protected override void testSelectedPhase(staffSpecSL aStaffSpecSL, DateTime runDateTime) { }
    }

    
    class defaultCouplerCreate : defaultCoupler
    {
        public defaultCouplerCreate
            (String csSHDDB, String csEHDDB, String csQLRDB, String aActionName, int aTryCount)
            : base(csSHDDB, csEHDDB, csQLRDB, aActionName, aTryCount){}

        protected override void doSelectedItem(staffSpec aStaffSpec)
        {
            updateSelectedItem(aStaffSpec, "T", "T", "Y");
        }


        protected override void testSelectedPhase(staffSpecSL aStaffSpecSL, DateTime runDateTime)
        {
            foreach (staffSpec wSS in aStaffSpecSL.Values)
            {
                if (testSelectedItem(wSS, "T", "T", "Y"))
                    mTestedOK.Add(wSS.queueItem.ToString());
                else
                    if (wSS.attempts > mTryCount) mFailed.Add(wSS.queueItem.ToString());
            }
        }

    }

    class defaultCouplerEnable : defaultCoupler
    {
        public defaultCouplerEnable
            (String csSHDDB, String csEHDDB, String csQLRDB, String aActionName, int aTryCount)
            : base(csSHDDB, csEHDDB, csQLRDB, aActionName, aTryCount) { }

        protected override void doSelectedItem(staffSpec aStaffSpec)
        {
            updateSelectedItem(aStaffSpec, "T", "T", "Y");
            //String HDInsert =
            //    "st_eventSHDNewApproved " +
            //    "'" + aStaffSpec.EmailAddress + "', " +
            //    "'" + aStaffSpec.Site + "', " +
            //    "'" + aStaffSpec.Forename.Replace("'", "''") + "', " +
            //    "'" + aStaffSpec.Surname.Replace("'", "''") + "', " +
            //    "'" + aStaffSpec.QLId + "', " +
            //    "'" + aStaffSpec.Department.Replace("'", "''") + "', " +
            //    "'" + aStaffSpec.Tel + "', '" + aStaffSpec.EmpType + "'";
            //DataView wDV;
            ////EHDDB
            //wDV = staffUtility.readDataView(staffUtility.XrayUtilDB, HDInsert);
            ////SDDDB
            //wDV = staffUtility.readDataView(staffUtility.NovUtilDB, HDInsert);
            ////QLR
            //wDV = staffUtility.readDataView(staffUtility.couplerDB,
            //    "st_eventQLRNewApproved " +
            //    "'" + aStaffSpec.NDSName + "', '" + aStaffSpec.Site + "', '" + aStaffSpec.QLId + "', '" + aStaffSpec.DeptCode + "', '" + aStaffSpec.EmpType + "'"
            //);
            //mWritten.Add(aStaffSpec.queueItem.ToString());
        }

        protected override void testSelectedPhase(staffSpecSL aStaffSpecSL, DateTime runDateTime)
        {
            //DataView wDV;
            foreach (staffSpec wSS in aStaffSpecSL.Values)
            {
                if (testSelectedItem(wSS, "T", "T", "Y"))
                    mTestedOK.Add(wSS.queueItem.ToString());
                else
                    if (wSS.attempts > mTryCount) mFailed.Add(wSS.queueItem.ToString());

                //wDV = staffUtility.readDataView(mCSEHDDB,
                //    "SELECT COUNT(*) AS matches FROM people " +
                //    "WHERE user_email='" + wSS.NDSName + "@warkscol.ac.uk'"
                //);
                //if (wDV != null) if (wDV.Count > 0) if ((int)wDV[0]["matches"] > 0)
                //        {
                //            wDV = staffUtility.readDataView(mCSSHDDB,
                //                "SELECT COUNT(*) AS matches FROM people " +
                //                "WHERE user_email='" + wSS.NDSName + "@warkscol.ac.uk'"
                //            );
                //            if (wDV != null) if (wDV.Count > 0) if ((int)wDV[0]["matches"] > 0)
                //                    {
                //                        wDV = staffUtility.readDataView(staffUtility.couplerDB,
                //                            "SELECT COUNT(*) AS matches FROM wcPortal..wcQLRreportUsers " +
                //                            "WHERE NDSName='" + wSS.NDSName + "'"
                //                        );
                //                        if (wDV != null)
                //                            if (wDV.Count > 0)
                //                                if ((int)wDV[0]["matches"] > 0)
                //                                    mTestedOK.Add(wSS.queueItem.ToString());
                //                    }

                //        }
                //        else
                //            if (wSS.attempts > mTryCount) mFailed.Add(wSS.queueItem.ToString());
            }
        }

    }

    class defaultCouplerDisable : defaultCoupler
    {
        public defaultCouplerDisable
            (String csSHDDB, String csEHDDB, String csQLRDB, String aActionName, int aTryCount)
            : base(csSHDDB, csEHDDB, csQLRDB, aActionName, aTryCount) { }

        protected override void doSelectedItem(staffSpec aStaffSpec)
        {
            updateSelectedItem(aStaffSpec, "L", "T", "N");
        }

        protected override void testSelectedPhase(staffSpecSL aStaffSpecSL, DateTime runDateTime)
        {
            foreach (staffSpec wSS in aStaffSpecSL.Values)
            {

                if (testSelectedItem(wSS, "L", "%", "N"))
                    mTestedOK.Add(wSS.queueItem.ToString());
                else
                    if (wSS.attempts > mTryCount) mFailed.Add(wSS.queueItem.ToString());
            }

        }

    }

    class defaultCouplerTrash : defaultCoupler
    {
        public defaultCouplerTrash
            (String csSHDDB, String csEHDDB, String csQLRDB, String aActionName, int aTryCount)
            : base(csSHDDB, csEHDDB, csQLRDB, aActionName, aTryCount) { }

        protected override void doSelectedItem(staffSpec aStaffSpec)
        {
            updateSelectedItem(aStaffSpec, "L", "L", "N");
        }

        protected override void testSelectedPhase(staffSpecSL aStaffSpecSL, DateTime runDateTime)
        {
            //DataView wDV;
            foreach (staffSpec wSS in aStaffSpecSL.Values)
            {

                if (testSelectedItem(wSS, "L", "L", "N"))
                {
                    mTestedOK.Add(wSS.queueItem.ToString());
                }
                else
                    if (wSS.attempts > mTryCount) mFailed.Add(wSS.queueItem.ToString());
            }

        }

    }

    class defaultCouplerDelete : defaultCoupler
    {
        public defaultCouplerDelete
            (String csSHDDB, String csEHDDB, String csQLRDB, String aActionName, int aTryCount)
            : base(csSHDDB, csEHDDB, csQLRDB, aActionName, aTryCount) { }

        protected override void doSelectedItem(staffSpec aStaffSpec)
        {
            String HDdel = "wcIDMSTFDeleteEHDAndSHD " + "'" + aStaffSpec.QLId + "'";
            DataView wDV = staffUtility.readDataView(staffUtility.XrayUtilDB, HDdel); //from xray utils
            String QLRdel = "wcIDMSTFDeleteQLR '" + aStaffSpec.NDSName + "'";
            wDV = staffUtility.readDataView(staffUtility.NovUtilDB, QLRdel);
            // delete User apps (managed by Access Manager)here
            String UserAppsDel = "wcIDMSTFDeleteUserApps '" + aStaffSpec.NDSName + "'";
            wDV = staffUtility.readDataView(staffUtility.couplerDB, UserAppsDel);

            // delete staff identity record here
            String IDMdel = "wcIDMSTFDeleteIDM '" + aStaffSpec.NDSName + "'";
            wDV = staffUtility.readDataView(staffUtility.couplerDB, IDMdel);

        }

        protected override void testSelectedPhase(staffSpecSL aStaffSpecSL, DateTime runDateTime)
        {
            foreach (staffSpec wSS in aStaffSpecSL.Values)
            {

                if (testSelectedItem(wSS, "%", "%", "%"))
                {
                    String UserAppsDel = "wcIDMSTFDeleteUserApps '" + wSS.NDSName + "'";
                    DataView wDV = staffUtility.readDataView(staffUtility.couplerDB, UserAppsDel);
                    if (wDV.Count > 0)
                    {
                        if (wDV[0][0].ToString() == "N" && wDV[0][1].ToString() == "N" && wDV[0][2].ToString()== "N")
                        {
                            mTestedOK.Add(wSS.queueItem.ToString());
                        }
                    }
                }
                else
                    if (wSS.attempts > mTryCount) mFailed.Add(wSS.queueItem.ToString());
            }

        }

        protected override Boolean testSelectedItem(staffSpec aStaffSpec, String aEHDDBActiveFlag, String aSHDDBActiveFlag, String aQLRActiveFlag)
        {
            SortedList<String, String> wSL = testSelectedItemData(aStaffSpec);
            // test wcStaffIdentity to see if it has been deleted
            String sStaffTest;
            sStaffTest = "exists";
            DataView wDV = staffUtility.readDataView(staffUtility.couplerDB,
                "wcIDMSTFTestWcStaffIdentity " +
                "'" + aStaffSpec.NDSName + "'"
            );
            sStaffTest = wDV[0]["ReturnValue"].ToString().Trim().ToLower();
            return (
                (wSL["EHDDBActive"] == aEHDDBActiveFlag) &&
                (wSL["SHDDBActive"] == aSHDDBActiveFlag) &&
                (wSL["QLRActive"] == aQLRActiveFlag) &&
                (sStaffTest == "deleted")
              );
        }
    }

    class defaultCouplerUpdateUserName : defaultCoupler
    {
        public defaultCouplerUpdateUserName
            (String csSHDDB, String csEHDDB, String csQLRDB, String aActionName, int aTryCount)
            : base(csSHDDB, csEHDDB, csQLRDB, aActionName, aTryCount) { }

        protected override void doSelectedItem(staffSpec aStaffSpec)
        {
            String OldNDSName = "";
            // updateSelectedItem(aStaffSpec, "%", "%", "%");
            // update code here
            String[] AD = aStaffSpec.ActionData.Split('=');
            if (AD[0].ToString().Trim().ToLower() == "oldnds")
                OldNDSName = AD[1].ToString();

            String IDM = "wcIDMupdateApplications " + Convert.ToInt64(aStaffSpec.QLId) + ", '" + OldNDSName + "', '" + aStaffSpec.NDSName.ToString().Trim() + "', '" + aStaffSpec.Forename.ToString().Trim().Replace("'", "''") + "', '" + aStaffSpec.Surname.ToString().Trim().Replace("'", "''") + "'";
            DataView wDV = staffUtility.readDataView(staffUtility.couplerDB, IDM);

        }

        protected override void testSelectedPhase(staffSpecSL aStaffSpecSL, DateTime runDateTime)
        {
            //DataView wDV;
            String OldNDSName = "";
            String[] AD;
            String IDM;
            DataView wDV;
            foreach (staffSpec wSS in aStaffSpecSL.Values)
            {
                // updateSelectedItem(aStaffSpec, "%", "%", "%");
                // update code here
                AD = wSS.ActionData.Split('=');
                if (AD[0].ToString().Trim().ToLower() == "oldnds")
                    OldNDSName = AD[1].ToString();

                IDM = "wcIDMupdateApplications " + Convert.ToInt64(wSS.QLId) + ", '" + OldNDSName + "', '" + wSS.NDSName.ToString().Trim() + "', '" + wSS.Forename.ToString().Trim().Replace("'", "''") + "', '" + wSS.Surname.ToString().Trim().Replace("'", "''") + "'";
                wDV = staffUtility.readDataView(staffUtility.couplerDB, IDM);
                if (wDV.Count > 0)
                {
                    if ((wDV[0]["ReturnQLR"].ToString() == "4" || wDV[0]["ReturnQLR"].ToString() == "1") && wDV[0]["ReturnSHD"].ToString() == "1" && wDV[0]["ReturnEHD"].ToString() == "1")
                    {
                        mTestedOK.Add(wSS.queueItem.ToString());
                        // Success.
                    }
                    else
                    {
                        //Not Updated
                        if (wSS.attempts > mTryCount) mFailed.Add(wSS.queueItem.ToString());
                    }
                }
                else
                {
                    //maybe a problem with the return value
                    if (wSS.attempts > mTryCount) mFailed.Add(wSS.queueItem.ToString());
                }
            }
        }
    }

    //=================================================================================================================================

    class defaultCouplerRemoveFromGoogleGroup : defaultCoupler
    {
        public defaultCouplerRemoveFromGoogleGroup(String csSHDDB, String csEHDDB, String csQLRDB, String aActionName, int aTryCount)
            : base(csSHDDB, csEHDDB, csQLRDB, aActionName, aTryCount) { }

        public Boolean WcIsMember(String StaffEmail, String StaffGroup)
        {
            Boolean WcIs = false;
            String gc = "@warwickshire.ac.uk";// google context
            try
            {
                Member m = wcService.Members.Get(StaffGroup.Trim() + gc, StaffEmail.Trim() + gc).Execute();
                
                WcIs = true;
            }
            catch (Google.GoogleApiException ue)
            {
                String msg = ue.Message.ToString();
                staffUtility.putLog("RemoveFromGoogleGroup.WcIsMemberFailure: " + StaffEmail.Trim().ToString() + gc, msg.ToString());
            }

            return WcIs;
        }

        protected override void doSelectedItem(staffSpec aStaffSpec)
        {
            // DELETE user from groups 
            String gc = "@warwickshire.ac.uk";

            if (WcIsMember(aStaffSpec.NDSName.ToString(), aStaffSpec.ActionData.ToString().Trim()) == true)
            {
                try
                {
                    wcService.Members.Delete(aStaffSpec.ActionData.ToString().Trim() + gc, aStaffSpec.NDSName.ToString().Trim() + gc).Execute();
                    mWritten.Add(aStaffSpec.queueItem.ToString());
                }
                catch (Google.GoogleApiException uex)
                {
                    String msg = uex.Message;
                    staffUtility.putLog("DeleteGoogleGroup Members.Delete Failure :" + aStaffSpec.ActionData.ToString().Trim() + "; " + aStaffSpec.NDSName.ToString().Trim() + gc, msg.ToString());
                    //String sReason = uex.Reason.ToString();
                }
            }
        } 
        protected override void testSelectedPhase(staffSpecSL aStaffSpecSL, DateTime runDateTime)
        {
            //AppsService service = new AppsService("stu.warwickshire.ac.uk", "gadmin@stu.warwickshire.ac.uk", "thebigpic7ure");
            // this is for the GROUP test phase
            foreach (staffSpec wSS in aStaffSpecSL.Values)
            {
                try
                {    
                    String locID = wSS.LocID.ToString();
                    bool googleGroup = WcIsMember(wSS.NDSName.ToString().Trim(), wSS.ActionData.ToString().Trim());
                    
                    if (!googleGroup) mTestedOK.Add(wSS.queueItem.ToString());
               }
                catch (Google.GoogleApiException uex)
                {
                    String msg = uex.Message;
                    staffUtility.putLog("RemoveFromGoogleGroupFailure DoItem: " + wSS.NDSName.Trim().ToString(), msg.ToString());
                    //String sReason = uex.Reason.ToString();
                    if (wSS.attempts > mTryCount) mFailed.Add(wSS.queueItem.ToString());
                    if (msg.ToLower().Contains("userkey [404]"))
                    {
                        //User does not exist in Google Apps
                    }
                }
            }
        }


    }
    class defaultCouplerAddToGoogleGroup : defaultCoupler
    {
        public defaultCouplerAddToGoogleGroup(String csSHDDB, String csEHDDB, String csQLRDB, String aActionName, int aTryCount)
            : base(csSHDDB, csEHDDB, csQLRDB, aActionName, aTryCount) { }

        public Boolean WcIsMember(String StaffEmail, String StaffGroup)
        {
            Boolean WcIs = false;
            String gc = "@warwickshire.ac.uk";// google context
            try
            {
                Member m = wcService.Members.Get(StaffGroup.Trim() + gc, StaffEmail.Trim() + gc).Execute();

                WcIs = true;
            }
            catch (Google.GoogleApiException ue)
            {
                String msg = ue.Message.ToString();
                staffUtility.putLog("AddToGoogleGroup.WcIsMemberFailure: " + StaffEmail.Trim().ToString() + gc, msg.ToString());
            }
            return WcIs;
        }

        protected override void doSelectedItem(staffSpec aStaffSpec)
        {
            // DELETE user from groups 
            String gc = "@warwickshire.ac.uk";
            Member newMemberBody = new Member();
            newMemberBody.Email = aStaffSpec.NDSName.ToString().Trim() + gc;
            newMemberBody.Role = "MEMBER";
            newMemberBody.Type = "USER";

            if (WcIsMember(aStaffSpec.NDSName.ToString(), aStaffSpec.ActionData.ToString().Trim()) == false)
            {
                try
                {
                    wcService.Members.Insert(newMemberBody, aStaffSpec.ActionData.ToString().Trim() + gc).Execute();
                    mWritten.Add(aStaffSpec.queueItem.ToString());
                }
                catch (Google.GoogleApiException uex)
                {
                    String msg = uex.Message;
                    staffUtility.putLog("AddToGoogleGroup Members.Insert Failure " + aStaffSpec.ActionData.ToString().Trim() + ": " + aStaffSpec.NDSName.ToString().Trim() + gc, msg.ToString());
                    //String sReason = uex.Reason.ToString();
                }

            }
        }
        protected override void testSelectedPhase(staffSpecSL aStaffSpecSL, DateTime runDateTime)
        {
            //AppsService service = new AppsService("stu.warwickshire.ac.uk", "gadmin@stu.warwickshire.ac.uk", "thebigpic7ure");
            // this is for the GROUP test phase
            foreach (staffSpec wSS in aStaffSpecSL.Values)
            {
                try
                {
                    String locID = wSS.LocID.ToString();
                    bool googleGroup = WcIsMember(wSS.NDSName.ToString().Trim(), wSS.ActionData.ToString().Trim());
                    if (googleGroup) mTestedOK.Add(wSS.queueItem.ToString());
                }
                catch (Google.GoogleApiException uex)
                {
                    String msg = uex.Message;
                    staffUtility.putLog("RemoveFromGoogleGroupFailure DoItem: " + wSS.NDSName.Trim().ToString(), msg.ToString());
                    //String sReason = uex.Reason.ToString();
                    if (wSS.attempts > mTryCount) mFailed.Add(wSS.queueItem.ToString());
                    if (msg.ToLower().Contains("userkey [404]"))
                    {
                        //User does not exist in Google Apps
                    }
                }
            }
        }
    }


    //=================================================================================================================================


    class defaultCouplerDeleteGoogleGroup : defaultCoupler
    {
        public defaultCouplerDeleteGoogleGroup
            (String csSHDDB, String csEHDDB, String csQLRDB, String aActionName, int aTryCount)
            : base(csSHDDB, csEHDDB, csQLRDB, aActionName, aTryCount) { }
        //  create a SP for returning ALL sites for a studentID hat ^ delimitted
        // or add extra column in wcStudentPrimaryCourse ?? maybe not.

        public Boolean WcIsMember(String StaffEmail, String StaffGroup)
        {
            Boolean WcIs = false;
            String gc = "@warwickshire.ac.uk";// google context

            try
            {
                Member m = wcService.Members.Get(StaffGroup.Trim() + gc, StaffEmail.Trim() + gc).Execute();

                WcIs = true;
            }
            catch (Google.GoogleApiException ue)
            {
                String msg = ue.Message.ToString();
                staffUtility.putLog("DeleteGoogleGroup.WcIsMemberFailure: " + StaffEmail.Trim().ToString() + gc, msg.ToString());

            }

            return WcIs;
        }

        protected override void doSelectedItem(staffSpec aStaffSpec)
        {
            bool googleRls = false;
            bool googleMM = false;
            bool googleTrident = false;
            bool googleRugby = false;
            bool googleHenley = false;
            bool googlePershore = false;
            bool googleMalvern = false;
            bool googleEvesham = false;

            Member newMemberBody = new Member();
            String gc = "@warwickshire.ac.uk";
            newMemberBody.Email = aStaffSpec.NDSName.ToString().Trim() + gc;
            newMemberBody.Role = "MEMBER";
            newMemberBody.Type = "USER";

            try
            {
                // DELETE user from groups 

                if (WcIsMember(aStaffSpec.NDSName.ToString(), "STAFF-ALLSITES") == true)
                {
                    try
                    {
                        wcService.Members.Delete("STAFF-ALLSITES" + gc, aStaffSpec.NDSName.ToString().Trim() + gc).Execute();
                    }
                    catch (Google.GoogleApiException uex)
                    {
                        String msg = uex.Message;
                        staffUtility.putLog("DeleteGoogleGroup Members.Delete Failure STAFF-ALLSITES: " + aStaffSpec.NDSName.ToString().Trim() + gc, msg.ToString());
                        //String sReason = uex.Reason.ToString();
                    }

                }

                if (WcIsMember(aStaffSpec.NDSName.ToString(), "CLASSROOM_TEACHERS") == true)
                {
                    try
                    {
                        wcService.Members.Delete("CLASSROOM_TEACHERS" + gc, aStaffSpec.NDSName.ToString().Trim() + gc).Execute();
                    }
                    catch (Google.GoogleApiException uex)
                    {
                        String msg = uex.Message;
                        staffUtility.putLog("DeleteGoogleGroup Members.Delete Failure CLASSROOM_TEACHERS: " + aStaffSpec.NDSName.ToString().Trim() + gc, msg.ToString());
                        //String sReason = uex.Reason.ToString();
                    }

                }

                {
                    // start malvern
                    googleMalvern = WcIsMember(aStaffSpec.NDSName.ToString(), "STAFF-MALVERN");

                    if (googleMalvern)
                    {
                        try
                        {
                            wcService.Members.Delete("staff-malvern" + gc, aStaffSpec.NDSName.ToString().Trim() + gc).Execute();

                            //service.Groups.RemoveMemberFromGroup(aStaffSpec.NDSName.ToString(), "STAFF-PER");
                        }
                        catch (Google.GoogleApiException uex)
                        {
                            String msg = uex.Message;
                            staffUtility.putLog("CreateGoogleGroup Members.Delete Failure staff-malvern: " + aStaffSpec.NDSName.ToString().Trim() + gc, msg.ToString());
                            //String sReason = uex.Reason.ToString();
                        }
                    }
                    // end malvern

                    // start evesham
                    googleEvesham = WcIsMember(aStaffSpec.NDSName.ToString(), "STAFF-EVESHAM");

                    if (googleEvesham)
                    {
                        try
                        {
                            wcService.Members.Delete("staff-evesham" + gc, aStaffSpec.NDSName.ToString().Trim() + gc).Execute();

                            //service.Groups.RemoveMemberFromGroup(aStaffSpec.NDSName.ToString(), "STAFF-PER");
                        }
                        catch (Google.GoogleApiException uex)
                        {
                            String msg = uex.Message;
                            staffUtility.putLog("DeleteGoogleGroup Members.Delete Failure staff-evesham: " + aStaffSpec.NDSName.ToString().Trim() + gc, msg.ToString());
                            //String sReason = uex.Reason.ToString();
                        }
                    }
                    // end evesham
                    // start pershore
                    googlePershore = WcIsMember(aStaffSpec.NDSName.ToString(), "STAFF-PERSHORE");

                    if (googlePershore)
                    {
                        try
                        {
                            wcService.Members.Delete("staff-pershore" + gc, aStaffSpec.NDSName.ToString().Trim() + gc).Execute();

                            //service.Groups.RemoveMemberFromGroup(aStaffSpec.NDSName.ToString(), "STAFF-PER");
                        }
                        catch (Google.GoogleApiException uex)
                        {
                            String msg = uex.Message;
                            staffUtility.putLog("DeleteGoogleGroup Members.Delete Failure staff-pershore: " + aStaffSpec.NDSName.ToString().Trim() + gc, msg.ToString());
                            //String sReason = uex.Reason.ToString();
                        }
                    }
                    // end pershore

                    // start lspa
                    googleRls = WcIsMember(aStaffSpec.NDSName.ToString(), "STAFF-LSPA");
                    if (googleRls)
                    {
                        try
                        {
                            wcService.Members.Delete("staff-lspa" + gc, aStaffSpec.NDSName.ToString().Trim() + gc).Execute();

                            //service.Groups.RemoveMemberFromGroup(aStaffSpec.NDSName.ToString(), "STAFF-PER");
                        }
                        catch (Google.GoogleApiException uex)
                        {
                            String msg = uex.Message;
                            staffUtility.putLog("DeleteGoogleGroup Members.Delete Failure staff-lspa: " + aStaffSpec.NDSName.ToString().Trim() + gc, msg.ToString());
                            //String sReason = uex.Reason.ToString();
                        }

                    }
                    // end lspa

                    // start rugby
                    googleRugby = WcIsMember(aStaffSpec.NDSName.ToString(), "STAFF-RUGBY");
                    if (googleRugby)
                    {
                        try
                        {
                            wcService.Members.Delete("staff-rugby" + gc, aStaffSpec.NDSName.ToString().Trim() + gc).Execute();

                            //service.Groups.RemoveMemberFromGroup(aStaffSpec.NDSName.ToString(), "STAFF-PER");
                        }
                        catch (Google.GoogleApiException uex)
                        {
                            String msg = uex.Message;
                            staffUtility.putLog("DeleteGoogleGroup Members.Delete Failure staff-rugby: " + aStaffSpec.NDSName.ToString().Trim() + gc, msg.ToString());
                            //String sReason = uex.Reason.ToString();
                        }
                    }
                    // end rugby
                    // start moreton morrell

                    googleMM = WcIsMember(aStaffSpec.NDSName.ToString(), "STAFF-MORETON");
                    if (googleMM)
                    {
                        try
                        {
                            wcService.Members.Delete("staff-moreton" + gc, aStaffSpec.NDSName.ToString().Trim() + gc).Execute();

                            //service.Groups.RemoveMemberFromGroup(aStaffSpec.NDSName.ToString(), "STAFF-PER");
                        }
                        catch (Google.GoogleApiException uex)
                        {
                            String msg = uex.Message;
                            staffUtility.putLog("DeleteGoogleGroup Members.Delete Failure staff-moreton: " + aStaffSpec.NDSName.ToString().Trim() + gc, msg.ToString());
                            //String sReason = uex.Reason.ToString();
                        }
                    }
                    // end moreton morrell

                    // start trident
                    googleTrident = WcIsMember(aStaffSpec.NDSName.ToString(), "STAFF-TRIDENT");
                    if (googleTrident)
                    {
                        try
                        {
                            wcService.Members.Delete("staff-trident" + gc, aStaffSpec.NDSName.ToString().Trim() + gc).Execute();

                            //service.Groups.RemoveMemberFromGroup(aStaffSpec.NDSName.ToString(), "STAFF-PER");
                        }
                        catch (Google.GoogleApiException uex)
                        {
                            String msg = uex.Message;
                            staffUtility.putLog("DeleteGoogleGroup Members.Delete Failure staff-trident: " + aStaffSpec.NDSName.ToString().Trim() + gc, msg.ToString());
                            //String sReason = uex.Reason.ToString();
                        }
                    }
                    // end trident
                    // start henley
                    googleHenley = WcIsMember(aStaffSpec.NDSName.ToString(), "STAFF-HENLEY");
                    if (googleHenley)
                    {
                        try
                        {
                            wcService.Members.Delete("staff-henley" + gc, aStaffSpec.NDSName.ToString().Trim() + gc).Execute();

                            //service.Groups.RemoveMemberFromGroup(aStaffSpec.NDSName.ToString(), "STAFF-PER");
                        }
                        catch (Google.GoogleApiException uex)
                        {
                            String msg = uex.Message;
                            staffUtility.putLog("DelteGoogleGroup Members.Delete Failure staff-henley: " + aStaffSpec.NDSName.ToString().Trim() + gc, msg.ToString());
                            //String sReason = uex.Reason.ToString();
                        }

                    }
                    // end henley
                }

                mWritten.Add(aStaffSpec.queueItem.ToString());

            }
            catch (Google.GoogleApiException uexi)
            {
                String msgi = uexi.Message;
                //String sReasoni = uexi.Reason.ToString();

            }
            //updateSelectedItem(aStudentSpec, "T", "T", "Y");
        }
        // TEST select pahase of DELETE GOOGLE GROUP
        protected override void testSelectedPhase(staffSpecSL aStaffSpecSL, DateTime runDateTime)
        {
            //AppsService service = new AppsService("stu.warwickshire.ac.uk", "gadmin@stu.warwickshire.ac.uk", "thebigpic7ure");
            // this is for the GROUP test phase
            foreach (staffSpec wSS in aStaffSpecSL.Values)
            {

                try
                {
                        bool groupsDone = false;
                        String locID = wSS.LocID.ToString();
                        bool googleBroadcast = WcIsMember(wSS.NDSName.ToString().Trim(), "STAFF-ALLSITES");
                        bool googleClassroom = WcIsMember(wSS.NDSName.ToString().Trim(), "CLASSROOM_TEACHERS");
                        bool googleRLS = WcIsMember(wSS.NDSName.ToString(), "STAFF-LSPA");
                        bool googleMalvern = WcIsMember(wSS.NDSName.ToString(), "STAFF-MALVERN");
                        bool googleEvesham = WcIsMember(wSS.NDSName.ToString(), "STAFF-EVESHAM");
                        bool googlePershore = WcIsMember(wSS.NDSName.ToString(), "STAFF-PERSHORE");
                        bool googlemoreton = WcIsMember(wSS.NDSName.ToString(), "STAFF-MORETON");
                        bool googleTrident = WcIsMember(wSS.NDSName.ToString(), "STAFF-TRIDENT");
                        bool googleHenley = WcIsMember(wSS.NDSName.ToString(), "STAFF-HENLEY");
                        bool googleRugby = WcIsMember(wSS.NDSName.ToString(), "STAFF-RUGBY");
                        if (!googleBroadcast  && !googleClassroom && !googleRLS && !googleMalvern && !googleEvesham && !googlePershore
                         && !googlemoreton && !googleTrident && !googleHenley && !googleRugby) groupsDone = true;
                        if (groupsDone) mTestedOK.Add(wSS.queueItem.ToString());

                    //Write a coupler job for moveObject here
                    //studentUtility.writeCouplerMessageQueueStu(wSS.NDSName, "", "GoogleGroupAdd");

                }
                catch (Google.GoogleApiException uex)
                {
                    String msg = uex.Message;
                    staffUtility.putLog("DeleteGoogleGroupFailure DoItem: " + wSS.NDSName.Trim().ToString(), msg.ToString());
                    //String sReason = uex.Reason.ToString();
                    if (wSS.attempts > mTryCount) mFailed.Add(wSS.queueItem.ToString());
                    if (msg.ToLower().Contains("userkey [404]"))
                    {
                        //User does not exist in Google Apps
                    }
                }
            }
        }

    }
    //END OF DELETE GOOGLE GROUPS ===================================================================================================================


    //START of CreateGoogleGroup =================================================================================================================================
    class defaultCouplerCreateGoogleGroup : defaultCoupler
    {
        public defaultCouplerCreateGoogleGroup
            (String csSHDDB, String csEHDDB, String csQLRDB, String aActionName, int aTryCount)
            : base(csSHDDB, csEHDDB, csQLRDB, aActionName, aTryCount) { }
        //  create a SP for returning ALL sites for a studentID hat ^ delimitted
        // or add extra column in wcStudentPrimaryCourse ?? maybe not.

        public Boolean WcIsMember(String StaffEmail, String StaffGroup)
        {
            Boolean WcIs = false;
            String gc = "@warwickshire.ac.uk";// google context

            try
            {
                Member m = wcService.Members.Get(StaffGroup.Trim() + gc, StaffEmail.Trim() + gc).Execute();
                
                WcIs = true;
            }
            catch (Google.GoogleApiException ue)
            { 
                String msg = ue.Message.ToString();
                staffUtility.putLog("CreateGoogleGroup.WcIsMemberFailure: " + StaffEmail.Trim().ToString() + gc, msg.ToString());
  
            }

            return WcIs;
        }

        protected override void doSelectedItem(staffSpec aStaffSpec)
        {

             

            //user exists.
            bool rls = false;
            bool mm = false;
            bool trident = false;
            bool rugby = false;
            bool henley = false;
            bool pershore = false;
            bool malvern = false;
            bool evesham = false;


            bool googleRls = false;
            bool googleMM = false;
            bool googleTrident = false;
            bool googleRugby = false;
            bool googleHenley = false;
            bool googlePershore = false;
            bool googleMalvern = false;
            bool googleEvesham = false;

            Member newMemberBody = new Member();
            String gc = "@warwickshire.ac.uk";
            newMemberBody.Email = aStaffSpec.NDSName.ToString().Trim() + gc;
            newMemberBody.Role = "MEMBER";
            newMemberBody.Type = "USER";


            try
            {
                // Add user to group 

                if (WcIsMember(aStaffSpec.NDSName.ToString(), "CLASSROOM_TEACHERS") == false)
                {
                    try
                    {
                        wcService.Members.Insert(newMemberBody, "CLASSROOM_TEACHERS" + gc).Execute();
                    }
                    catch (Google.GoogleApiException uex)
                    {
                        String msg = uex.Message;
                        staffUtility.putLog("CreateGoogleGroup Members.Insert Failure CLASSROOM_TEACHERS: " + aStaffSpec.NDSName.ToString().Trim() + gc, msg.ToString());
                        //String sReason = uex.Reason.ToString();
                    }

                }

                //This group is for WCG Ltd. staff
                if (aStaffSpec.WCG_Ltd.ToString() == "Y")
                {
                    if (WcIsMember(aStaffSpec.NDSName.ToString(), "STAFF-WCG-LTD") == false)
                    {
                        try
                        {
                            wcService.Members.Insert(newMemberBody, "STAFF-WCG-LTD" + gc).Execute();
                        }
                        catch (Google.GoogleApiException uex)
                        {
                            String msg = uex.Message;
                            staffUtility.putLog("CreateGoogleGroup Members.Insert Failure STAFF-WCG-LTD: " + aStaffSpec.NDSName.ToString().Trim() + gc, msg.ToString());
                            //String sReason = uex.Reason.ToString();
                        }

                    }
                 }
                else
                {
                    //Delete from the group
                    try
                    {
                        if (WcIsMember(aStaffSpec.NDSName.ToString(), "STAFF-WCG-LTD") == true) wcService.Members.Delete("STAFF-WCG-LTD" + gc, aStaffSpec.NDSName.ToString().Trim() + gc).Execute();

                        //service.Groups.RemoveMemberFromGroup(aStaffSpec.NDSName.ToString(), "STAFF-PER");
                    }
                    catch (Google.GoogleApiException uex)
                    {
                        String msg = uex.Message;
                        staffUtility.putLog("CreateGoogleGroup Members.Delete Failure STAFF-WCG-LTD: " + aStaffSpec.NDSName.ToString().Trim() + gc, msg.ToString());
                        //String sReason = uex.Reason.ToString();
                    }

                }

                if (aStaffSpec.VisitingLec != "True")
                {

                    if (WcIsMember(aStaffSpec.NDSName.ToString(), "STAFF-ALLSITES") == false)
                    {
                        try
                        {
                            wcService.Members.Insert(newMemberBody, "STAFF-ALLSITES" + gc).Execute();
                        }
                        catch (Google.GoogleApiException uex)
                        {
                            String msg = uex.Message;
                            staffUtility.putLog("CreateGoogleGroup Members.Insert Failure STAFF-ALLSITES: " + aStaffSpec.NDSName.ToString().Trim() + gc, msg.ToString());
                            //String sReason = uex.Reason.ToString();
                        }

                    }


                //{
                    // start malvern
                    if (aStaffSpec.LocID.ToString().ToUpper().Trim().IndexOf("MALVERN") >= 0) malvern = true;
                    googleMalvern = WcIsMember(aStaffSpec.NDSName.ToString(), "STAFF-MALVERN");

                    if (googleMalvern && !malvern)
                    {
                        try
                        {
                            //wcService.Members.Delete("staff-malvern" + gc, aStaffSpec.NDSName.ToString().Trim() + gc).Execute();

                            //service.Groups.RemoveMemberFromGroup(aStaffSpec.NDSName.ToString(), "STAFF-PER");
                        }
                        catch (Google.GoogleApiException uex)
                        {
                            String msg = uex.Message;
                            staffUtility.putLog("CreateGoogleGroup Members.Delete Failure staff-malvern: " + aStaffSpec.NDSName.ToString().Trim() + gc, msg.ToString());
                            //String sReason = uex.Reason.ToString();
                        }
                    }
                    else
                    {
                        if (malvern && !googleMalvern)
                        {
                            try
                            {
                                wcService.Members.Insert(newMemberBody, "STAFF-MALVERN" + gc).Execute();
                                //service.Groups.AddMemberToGroup(aStaffSpec.NDSName.ToString(), "STAFF-PER");
                            }
                            catch (Google.GoogleApiException uex)
                            {
                                String msg = uex.Message;
                                staffUtility.putLog("CreateGoogleGroup Members.Insert Failure STAFF-MALVERN: " + aStaffSpec.NDSName.ToString().Trim() + gc, msg.ToString());
                                //String sReason = uex.Reason.ToString();
                            }

                        }
                    }
                    // end malvern

                    // start evesham
                    if (aStaffSpec.LocID.ToString().ToUpper().Trim().IndexOf("EVESHAM") >= 0) evesham = true;
                    googleEvesham = WcIsMember(aStaffSpec.NDSName.ToString(), "STAFF-EVESHAM");

                    if (googleEvesham && !evesham)
                    {
                        try
                        {
                            //wcService.Members.Delete("staff-evesham" + gc, aStaffSpec.NDSName.ToString().Trim() + gc).Execute();

                            //service.Groups.RemoveMemberFromGroup(aStaffSpec.NDSName.ToString(), "STAFF-PER");
                        }
                        catch (Google.GoogleApiException uex)
                        {
                            String msg = uex.Message;
                            staffUtility.putLog("CreateGoogleGroup Members.Delete Failure staff-evesham: " + aStaffSpec.NDSName.ToString().Trim() + gc, msg.ToString());
                            //String sReason = uex.Reason.ToString();
                        }
                    }
                    else
                    {
                        if (evesham && !googleEvesham)
                        {
                            try
                            {
                                wcService.Members.Insert(newMemberBody, "STAFF-EVESHAM" + gc).Execute();
                                //service.Groups.AddMemberToGroup(aStaffSpec.NDSName.ToString(), "STAFF-PER");
                            }
                            catch (Google.GoogleApiException uex)
                            {
                                String msg = uex.Message;
                                staffUtility.putLog("CreateGoogleGroup Members.Insert Failure STAFF-EVESHAM: " + aStaffSpec.NDSName.ToString().Trim() + gc, msg.ToString());
                                //String sReason = uex.Reason.ToString();
                            }

                        }
                    }
                    // end evesham


                    
                    // start pershore
                    if (aStaffSpec.LocID.ToString().ToUpper().Trim().IndexOf("PERSHORE") >= 0) pershore = true;
                    googlePershore = WcIsMember(aStaffSpec.NDSName.ToString(), "STAFF-PERSHORE");

                    if (googlePershore && !pershore)
                    {
                        try
                        {
                           //wcService.Members.Delete("staff-pershore" + gc, aStaffSpec.NDSName.ToString().Trim() + gc).Execute();

                            //service.Groups.RemoveMemberFromGroup(aStaffSpec.NDSName.ToString(), "STAFF-PER");
                        }
                        catch (Google.GoogleApiException uex)
                        {
                            String msg = uex.Message;
                            staffUtility.putLog("CreateGoogleGroup Members.Delete Failure staff-pershore: " + aStaffSpec.NDSName.ToString().Trim() + gc, msg.ToString());
                            //String sReason = uex.Reason.ToString();
                        }
                    }
                    else
                    {
                        if (pershore && !googlePershore)
                        {
                            try
                            {
                                wcService.Members.Insert(newMemberBody, "STAFF-PERSHORE" + gc).Execute();
                                //service.Groups.AddMemberToGroup(aStaffSpec.NDSName.ToString(), "STAFF-PER");
                            }
                            catch (Google.GoogleApiException uex)
                            {
                                String msg = uex.Message;
                                staffUtility.putLog("CreateGoogleGroup Members.Insert Failure STAFF-PERSHORE: " + aStaffSpec.NDSName.ToString().Trim() + gc, msg.ToString());
                                //String sReason = uex.Reason.ToString();
                            }

                        }
                    }
                    // end pershore

                    // start lspa
                    if (aStaffSpec.LocID.ToString().ToUpper().Trim().IndexOf("LEAMINGTON SPA") >= 0) rls = true;
                    googleRls = WcIsMember(aStaffSpec.NDSName.ToString(), "STAFF-LSPA");
                    if (googleRls && !rls)
                    {
                        try
                        {
                            //wcService.Members.Delete("staff-lspa" + gc, aStaffSpec.NDSName.ToString().Trim() + gc).Execute();

                            //service.Groups.RemoveMemberFromGroup(aStaffSpec.NDSName.ToString(), "STAFF-PER");
                        }
                        catch (Google.GoogleApiException uex)
                        {
                            String msg = uex.Message;
                            staffUtility.putLog("CreateGoogleGroup Members.Delete Failure staff-lspa: " + aStaffSpec.NDSName.ToString().Trim() + gc, msg.ToString());
                            //String sReason = uex.Reason.ToString();
                        }

                    }
                    else
                    {
                        if (rls && !googleRls)
                        {
                            try
                            {
                                wcService.Members.Insert(newMemberBody, "STAFF-LSPA" + gc).Execute();
                                //service.Groups.AddMemberToGroup(aStaffSpec.NDSName.ToString(), "STAFF-PER");
                            }
                            catch (Google.GoogleApiException uex)
                            {
                                String msg = uex.Message;
                                staffUtility.putLog("CreateGoogleGroup Members.Insert Failure STAFF-LSPA: " + aStaffSpec.NDSName.ToString().Trim() + gc, msg.ToString());
                                //String sReason = uex.Reason.ToString();
                            }
                        }
                    }
                    // end lspa

                    // start rugby
                    if (aStaffSpec.LocID.ToString().Trim().ToUpper().IndexOf("RUGBY") >= 0) rugby = true;

                    googleRugby = WcIsMember(aStaffSpec.NDSName.ToString(), "STAFF-RUGBY");
                    if (googleRugby && !rugby)
                    {
                        try
                        {
                            //wcService.Members.Delete("staff-rugby" + gc, aStaffSpec.NDSName.ToString().Trim() + gc).Execute();

                            //service.Groups.RemoveMemberFromGroup(aStaffSpec.NDSName.ToString(), "STAFF-PER");
                        }
                        catch (Google.GoogleApiException uex)
                        {
                            String msg = uex.Message;
                            staffUtility.putLog("CreateGoogleGroup Members.Delete Failure staff-rugby: " + aStaffSpec.NDSName.ToString().Trim() + gc, msg.ToString());
                            //String sReason = uex.Reason.ToString();
                        }
                    }
                    else
                    {
                        if (rugby && !googleRugby)
                        {
                            try
                            {
                                wcService.Members.Insert(newMemberBody, "STAFF-RUGBY" + gc).Execute();
                                //service.Groups.AddMemberToGroup(aStaffSpec.NDSName.ToString(), "STAFF-PER");
                            }
                            catch (Google.GoogleApiException uex)
                            {
                                String msg = uex.Message;
                                staffUtility.putLog("CreateGoogleGroup Members.Insert Failure STAFF-RUGBY: " + aStaffSpec.NDSName.ToString().Trim() + gc, msg.ToString());
                                //String sReason = uex.Reason.ToString();
                            }

                        }
                    }
                    // end rugby
                    // start moreton morrell
                    if (aStaffSpec.LocID.ToString().Trim().ToUpper().IndexOf("MORETON MORRELL") >= 0) mm = true;

                    googleMM = WcIsMember(aStaffSpec.NDSName.ToString(), "STAFF-MORETON");
                    if (googleMM && !mm)
                    {
                        try
                        {
                            //wcService.Members.Delete("staff-moreton" + gc, aStaffSpec.NDSName.ToString().Trim() + gc).Execute();

                            //service.Groups.RemoveMemberFromGroup(aStaffSpec.NDSName.ToString(), "STAFF-PER");
                        }
                        catch (Google.GoogleApiException uex)
                        {
                            String msg = uex.Message;
                            staffUtility.putLog("CreateGoogleGroup Members.Delete Failure staff-moreton: " + aStaffSpec.NDSName.ToString().Trim() + gc, msg.ToString());
                            //String sReason = uex.Reason.ToString();
                        } 
                    }
                    else
                    {
                        if (mm && !googleMM)
                        {
                            try
                            {
                                wcService.Members.Insert(newMemberBody, "STAFF-MORETON" + gc).Execute();
                                //service.Groups.AddMemberToGroup(aStaffSpec.NDSName.ToString(), "STAFF-PER");
                            }
                            catch (Google.GoogleApiException uex)
                            {
                                String msg = uex.Message;
                                staffUtility.putLog("CreateGoogleGroup Members.Insert Failure STAFF-MORETON: " + aStaffSpec.NDSName.ToString().Trim() + gc, msg.ToString());
                                //String sReason = uex.Reason.ToString();
                            }
 
                        }
                    }
                    // end moreton morrell

                    // start trident
                    if (aStaffSpec.LocID.ToString().ToUpper().Trim().IndexOf("TRIDENT PARK") >= 0) trident = true;

                    googleTrident = WcIsMember(aStaffSpec.NDSName.ToString(), "STAFF-TRIDENT");
                    if (googleTrident && !trident)
                    {
                        try
                        {
                            //wcService.Members.Delete("staff-trident" + gc, aStaffSpec.NDSName.ToString().Trim() + gc).Execute();

                            //service.Groups.RemoveMemberFromGroup(aStaffSpec.NDSName.ToString(), "STAFF-PER");
                        }
                        catch (Google.GoogleApiException uex)
                        {
                            String msg = uex.Message;
                            staffUtility.putLog("CreateGoogleGroup Members.Delete Failure staff-trident: " + aStaffSpec.NDSName.ToString().Trim() + gc, msg.ToString());
                            //String sReason = uex.Reason.ToString();
                        }
                    }
                    else
                    {
                        if (trident && !googleTrident)
                        {
                            try
                            {
                                wcService.Members.Insert(newMemberBody, "STAFF-TRIDENT" + gc).Execute();
                                //service.Groups.AddMemberToGroup(aStaffSpec.NDSName.ToString(), "STAFF-PER");
                            }
                            catch (Google.GoogleApiException uex)
                            {
                                String msg = uex.Message;
                                staffUtility.putLog("CreateGoogleGroup Members.Insert Failure STAFF-TRIDENT: " + aStaffSpec.NDSName.ToString().Trim() + gc, msg.ToString());
                                //String sReason = uex.Reason.ToString();
                            }
                        }
                    }
                    // end trident
                    // start henley
                    if (aStaffSpec.LocID.ToString().ToUpper().Trim().IndexOf("HENLEY IN ARDEN") >= 0) henley = true;

                    googleHenley = WcIsMember(aStaffSpec.NDSName.ToString(), "STAFF-HENLEY");
                    if (googleHenley && !henley)
                    {
                        try
                        {
                            //wcService.Members.Delete("staff-henley" + gc, aStaffSpec.NDSName.ToString().Trim() + gc).Execute();

                            //service.Groups.RemoveMemberFromGroup(aStaffSpec.NDSName.ToString(), "STAFF-PER");
                        }
                        catch (Google.GoogleApiException uex)
                        {
                            String msg = uex.Message;
                            staffUtility.putLog("CreateGoogleGroup Members.Delete Failure staff-henley: " + aStaffSpec.NDSName.ToString().Trim() + gc, msg.ToString());
                            //String sReason = uex.Reason.ToString();
                        }

                    }
                    else
                    {
                        if (henley && !googleHenley)
                        {
                            try
                            {
                                wcService.Members.Insert(newMemberBody, "STAFF-HENLEY" + gc).Execute();
                                //service.Groups.AddMemberToGroup(aStaffSpec.NDSName.ToString(), "STAFF-PER");
                            }
                            catch (Google.GoogleApiException uex)
                            {
                                String msg = uex.Message;
                                staffUtility.putLog("CreateGoogleGroup Members.Insert Failure STAFF-HENLEY: " + aStaffSpec.NDSName.ToString().Trim() + gc, msg.ToString());
                                //String sReason = uex.Reason.ToString();
                            }

                        }
                    }
                    // end henley

                }

                mWritten.Add(aStaffSpec.queueItem.ToString());

            }
            catch (Google.GoogleApiException uexi)
            {
                String msgi = uexi.Message;
            }

        }
        protected override void testSelectedPhase(staffSpecSL aStaffSpecSL, DateTime runDateTime)
        {
            //AppsService service = new AppsService("stu.warwickshire.ac.uk", "gadmin@stu.warwickshire.ac.uk", "thebigpic7ure");
            // this is for the GROUP test phase
            foreach (staffSpec wSS in aStaffSpecSL.Values)
            {

                try
                {

                    //UserEntry ue = service.RetrieveUser(wSS.NDSName.ToString());
                    //UserEntry ue = service.RetrieveUser("atestlan12347");
                    bool rls = false;
                    bool mm = false;
                    bool trident = false;
                    bool rugby = false;
                    bool henley = false;
                    bool pershore = false;
                    bool malvern = false;
                    bool evesham = false;

                    bool groupsDone = true;
                    //*jj
                    //DataView wCouplerDV = staffUtility.readDataView(staffUtility.couplerDB, "wcStudentSites '" + wSS.NDSName.ToString() + "'");
                    String locID = wSS.LocID.ToString();
                    //if (wCouplerDV.Count > 0)
                    {
                        if (wSS.VisitingLec != "True")
                        {
                            bool googleBroadcast = WcIsMember(wSS.NDSName.ToString().Trim(), "STAFF-ALLSITES");
                            if (googleBroadcast == false) groupsDone = false;

                            bool googleRLS = WcIsMember(wSS.NDSName.ToString(), "STAFF-LSPA");
                            if (locID.ToString().ToUpper().Trim().IndexOf("LEAMINGTON SPA") >= 0) rls = true;
                            //if (googleRLS != rls) groupsDone = false;

                            bool googleMalvern = WcIsMember(wSS.NDSName.ToString(), "STAFF-MALVERN");
                            if (locID.ToString().ToUpper().Trim().IndexOf("MALVERN") >= 0) malvern = true;
                            //if (googleMalvern != malvern) groupsDone = false;

                            bool googleEvesham = WcIsMember(wSS.NDSName.ToString(), "STAFF-EVESHAM");
                            if (locID.ToString().ToUpper().Trim().IndexOf("EVESHAM") >= 0) evesham = true;
                            //if (googleEvesham != evesham) groupsDone = false;

                            bool googlePershore = WcIsMember(wSS.NDSName.ToString(), "STAFF-PERSHORE");
                            if (locID.ToString().ToUpper().Trim().IndexOf("PERSHORE") >= 0) pershore = true;
                            //if (googlePershore != pershore) groupsDone = false;

                            bool googlemoreton = WcIsMember(wSS.NDSName.ToString(), "STAFF-MORETON");
                            if (locID.ToString().ToUpper().Trim().IndexOf("MORETON MORRELL") >= 0) mm = true;
                            //if (googlemoreton != mm) groupsDone = false;

                            bool googleTrident = WcIsMember(wSS.NDSName.ToString(), "STAFF-TRIDENT");
                            if (locID.ToString().ToUpper().Trim().IndexOf("TRIDENT PARK") >= 0) trident = true;
                            //if (googleTrident != trident) groupsDone = false;

                            bool googleHenley = WcIsMember(wSS.NDSName.ToString(), "STAFF-HENLEY");
                            if (locID.ToString().ToUpper().Trim().IndexOf("HENLEY IN ARDEN") >= 0) henley = true;
                            //if (googleHenley != henley) groupsDone = false;

                            bool googleRugby = WcIsMember(wSS.NDSName.ToString(), "STAFF-RUGBY");
                            if (locID.ToString().ToUpper().Trim().IndexOf("RUGBY") >= 0) rugby = true;
                            //if (googleRugby != rugby) groupsDone = false;
                        }// end test for visiting lecturer
                        // This test is for all staff including Visiting Lecturers
                        bool googleClassroom = WcIsMember(wSS.NDSName.ToString().Trim(), "CLASSROOM_TEACHERS");
                        if (googleClassroom == false) groupsDone = false;

                        // Test for WCG Ltd staff group
                        bool googleWCGLtd = WcIsMember(wSS.NDSName.ToString().Trim(), "STAFF-WCG-LTD");
                        if (wSS.WCG_Ltd.ToString() == "Y")
                        {
                            if (googleWCGLtd == false) groupsDone = false;
                        }
                        else
                        {
                            if (googleWCGLtd == true) groupsDone = false;

                        }

                    }
                    if (groupsDone && (rls || mm || trident || rugby || pershore || malvern || evesham)) mTestedOK.Add(wSS.queueItem.ToString());

                    //Write a coupler job for moveObject here
                    //studentUtility.writeCouplerMessageQueueStu(wSS.NDSName, "", "GoogleGroupAdd");

                }
                catch (Google.GoogleApiException uex)
                {
                    String msg = uex.Message;
                    staffUtility.putLog("AddToGoogleGroupFailure DoItem: " + wSS.NDSName.Trim().ToString(), msg.ToString());
                    //String sReason = uex.Reason.ToString();
                    if (wSS.attempts > mTryCount) mFailed.Add(wSS.queueItem.ToString());
                    if (msg.ToLower().Contains("userkey [404]"))
                    {
                        //User does not exist in Google Apps
                    }
                }
            }
        }

    }
    //===================================================================================================================
    class defaultCouplerCreateGoogle : defaultCoupler
    {
        public defaultCouplerCreateGoogle
            (String csSHDDB, String csEHDDB, String csQLRDB, String aActionName, int aTryCount)
            : base(csSHDDB, csEHDDB, csQLRDB, aActionName, aTryCount) { }

        protected override void doSelectedItem(staffSpec aStaffSpec)
        {
            // AppsService service = new AppsService("warwickshire.ac.uk", "gadmin@warwickshire.ac.uk", "adcv325jemin");
            try
            {
                User u = wcService.Users.Get(aStaffSpec.NDSName.Trim().ToString() + "@warwickshire.ac.uk").Execute();

                //UserEntry ue = service.RetrieveUser(aStaffSpec.NDSName.ToString());
            }
            //catch (Google.GoogleApiException ex) use this as the exception handling
           catch (Google.GoogleApiException uex)
            {
                String msg = uex.Message;
                //String sReason = uex.Reason.ToString();
                staffUtility.putLog("CreateFailure DoItem: " + aStaffSpec.NDSName.Trim().ToString(), msg.ToString());
                //bool sitePershore = false;
                if (msg.ToLower().Contains("userkey [404]"))
                {
                    //User does not exist in Google Apps so create the user
                    try
                    {
                        //UserEntry ueCreate = service.CreateUser("abc12345678", "David", "Brogue", "c0v3ntry");

                        String password = staffUtility.createGooglePassword(aStaffSpec.NDSName);
                        User newuserbody = new User();
                        UserName newusername = new UserName();
                        UserPhoto up = new UserPhoto();
                        

                        newuserbody.PrimaryEmail = aStaffSpec.NDSName.ToString().Trim() + "@warwickshire.ac.uk";

                        newusername.GivenName = aStaffSpec.Forename.ToString().Trim();
                        newusername.FamilyName = aStaffSpec.Surname.ToString().Trim();

                        UserOrganization[] userOrg = new UserOrganization[1];
                        userOrg[0] = new UserOrganization();
                        userOrg[0].Department = "STAFF";
                        newuserbody.Organizations = userOrg;

                        newuserbody.Name = newusername;
                        newuserbody.Password = password;
                        newuserbody.OrgUnitPath = "/STAFF";
                        User results = wcService.Users.Insert(newuserbody).Execute();

                        //UsersResource.PhotosResource.GetRequest ug = wcService.Users.Photos.Get(wSS.NDSName.Trim().ToString() + "@warwickshire.ac.uk");
                        //upZ.PhotoData = ug.Execute().PhotoData;
                        //upZ.MimeType = ug.Execute().MimeType;
                        //upZ.Width = ug.Execute().Width;
                        //upZ.Height = ug.Execute().Height;
                        //upZ.Kind = ug.Execute().Kind;

                        String sPhoto = "";
                        Int64 EmployeeNumber = 0;

                        Int64.TryParse(aStaffSpec.QLId, out EmployeeNumber);
                        String IDM = "wcIDMGetStaffPhoto " + EmployeeNumber;
                        DataView wDV = staffUtility.readDataView(staffUtility.couplerDB, IDM);
                        if (wDV.Count > 0)
                        {
                            sPhoto = wDV[0]["Base64Photo"].ToString();
                        }
                        if (sPhoto.Length > 100)// photo text less than 100 is probably not an image
                        {
                            //Base64UrlEncode takes aparameter of base64 encode (this is retrieved from selectHR via a funtion wcIdentities.dbo.ToBase64(varbinary(max))
                            String StaffImage = staffUtility.Base64UrlEncode(sPhoto);
                            UserPhoto uphoto = new UserPhoto();
                            uphoto.Kind = "admin#directory#user#photo";
                            uphoto.Height = 96;
                            uphoto.Width = 96;
                            uphoto.MimeType = "image/bmp";
                            uphoto.PhotoData = StaffImage;
                            UsersResource.PhotosResource.UpdateRequest upRequest = wcService.Users.Photos.Update(uphoto, aStaffSpec.NDSName.Trim().ToString() + "@warwickshire.ac.uk");
                            upRequest.Execute();
                        }


                        //UserEntry ueCreate = service.CreateUser(aStaffSpec.NDSName.ToString(), aStaffSpec.Forename.ToString(), aStaffSpec.Surname.ToString(), password);
                        mWritten.Add(aStaffSpec.queueItem.ToString());
                        String[] queueItem;
                        queueItem = new String[1];
                        queueItem[0] = aStaffSpec.queueItem.ToString();
                        int rowsChanged = staffUtility.updateCouplerMessageQueueSet(queueItem, " SET actionData='" + password + "'");
                    }
                    catch (Google.GoogleApiException uexi)
                    {
                        String msgi = uexi.Message;
                        staffUtility.putLog("CreateFailure DoItem: " + aStaffSpec.NDSName.Trim().ToString(), msgi.ToString());
                        //String sReasoni = uexi.Reason.ToString();
                        if (msg.ToLower().Contains("userkey [404]"))
                        {
                            //User deleted recently - wait at least 5 days ok
                        }
                    }
                }
            }


            //updateSelectedItem(aStudentSpec, "T", "T", "Y");
        }


        protected override void testSelectedPhase(staffSpecSL aStaffSpecSL, DateTime runDateTime)
        {
            // AppsService service = new AppsService("warwickshire.ac.uk", "gadmin@warwickshire.ac.uk", "adcv325jemin");
            String emailBody = "";
            emailBody = "If you ever forget your password, then <b>''Self Service Password Reset''</b> is a useful facility for changing your password without having to contact the HelpDesk.";
            emailBody = emailBody + "<br />";
            emailBody = emailBody + "<br /> <b>To set up  ''Self Service Password Reset'':</b>";
            emailBody = emailBody + "<br /> 1. Go to https://pwm.warkscol.ac.uk/sspr/private/updateprofile - (also found on the intranet main menu as ''Configure Self Service Password Reset'').";
            emailBody = emailBody + "<br /> 2. Enter your mobile phone number - this will be stored on your profile and used to verify your identity when resetting your own password.";
            emailBody = emailBody + "<br />";
            emailBody = emailBody + "<br /> <b>If you then forget your password:</b>";
            emailBody = emailBody + "<br /> 1. Go to the College website https://wcg.ac.uk, > Sign In, > Intranet > Need Help > Forgot Your Password.";
            emailBody = emailBody + "<br /> 2. Enter your account name and mobile phone number (phone number must match the one that you have stored previously).";
            emailBody = emailBody + "<br /> 3. A six character token will be texted to your phone.";
            emailBody = emailBody + "<br /> 4. Enter this six character token on the web page which will then verify your identity.";
            emailBody = emailBody + "<br /> 5. Once verified you will be presented with a web page to enter a new password.";
            
            foreach (staffSpec wSS in aStaffSpecSL.Values)
            {

                try
                {
                    // regular expression to catch only alphanumeric and underscore, hyphen and apostrophe
                    if (System.Text.RegularExpressions.Regex.IsMatch(wSS.Forename.ToString(), @"^[a-zA-Z0-9_\-']+$"))
                    {
                        Boolean x = true;
                    }

                    User u = wcService.Users.Get(wSS.NDSName.Trim().ToString() + "@warwickshire.ac.uk").Execute();

                   
                    
                    //SJL 14/08/2015 - users no longer being deleted so test for the suspended flag when a new account is attempted to be created.
                    if (u.Suspended == true)
                    {
                        u.Suspended = false;
                        wcService.Users.Update(u, wSS.NDSName.Trim().ToString() + "@warwickshire.ac.uk").Execute();
                    }
                    mTestedOK.Add(wSS.queueItem.ToString());
                    //Write a coupler job for moveObject here
                    staffUtility.writeCouplerMessageQueueV2(wSS.NDSName, "", "GoogleGroupAdd");
                    //staffUtility.writeCouplerMessageQueueV2(wSS.NDSName, "", "CreateGoogleOU");

                    //SEND an EMAIL to user
                    bool wB = staffUtility.sendEmail(
                                    wSS.NDSName.Trim() + "@warwickshire.ac.uk",
                                    "PasswordReset@warwickshire.ac.uk",
                                    "Self Service Password Reset",
                                    emailBody, null);

                }
                catch (Google.GoogleApiException uex)
                {
                    String msg = uex.Message;
                    staffUtility.putLog("CreateFailure TestItem: " + wSS.NDSName.Trim().ToString(), msg.ToString());
                    //String sReason = uex.Reason.ToString();
                    if (wSS.attempts > mTryCount) mFailed.Add(wSS.queueItem.ToString());
                    if (msg.ToLower().Contains("userkey [404]"))
                    {
                        //User does not exist in Google Apps
                    }
                }

                //if (testSelectedItem(wSS, "T", "T", "Y"))
                //    mTestedOK.Add(wSS.queueItem.ToString());
                //else
                //    if (wSS.attempts > mTryCount) mFailed.Add(wSS.queueItem.ToString());
            }
        }

    }
    //==================================================================================================

    class defaultCouplerUpdateGoogleAccount : defaultCoupler
    {
        public defaultCouplerUpdateGoogleAccount
            (String csSHDDB, String csEHDDB, String csQLRDB, String aActionName, int aTryCount)
            : base(csSHDDB, csEHDDB, csQLRDB, aActionName, aTryCount) { }

        protected override void doSelectedItem(staffSpec aStaffSpec)
        {
 

            //updateSelectedItem(aStudentSpec, "T", "T", "Y");
        }


        protected override void testSelectedPhase(staffSpecSL aStaffSpecSL, DateTime runDateTime)
        {
            // AppsService service = new AppsService("warwickshire.ac.uk", "gadmin@warwickshire.ac.uk", "adcv325jemin");

            foreach (staffSpec wSS in aStaffSpecSL.Values)
            {

                try
                {

                    User u = wcService.Users.Get(wSS.NDSName.Trim().ToString() + "@warwickshire.ac.uk").Execute();

                   //Setup the department field
                    UserOrganization[] userOrg = new UserOrganization[1];
                    userOrg[0] = new UserOrganization();
                    userOrg[0].Department = "STAFF";
                    u.Organizations = userOrg;

                    //setup the Phone field
                    //UserPhone[] userP = new UserPhone[1];
                    //userP[0] = new UserPhone();
                    //userP[0].Type = "home";
                    //userP[0].Value = "12345";
                    //u.Phones = userP;
         
                    wcService.Users.Update(u, wSS.NDSName.Trim().ToString() + "@warwickshire.ac.uk").Execute();

                    mTestedOK.Add(wSS.queueItem.ToString());

                }
                catch (Google.GoogleApiException uex)
                {
                    String msg = uex.Message;
                    staffUtility.putLog("AccountUpdateFailure TestItem: " + wSS.NDSName.Trim().ToString(), msg.ToString());
                    //String sReason = uex.Reason.ToString();
                    if (wSS.attempts > mTryCount) mFailed.Add(wSS.queueItem.ToString());
                    if (msg.ToLower().Contains("userkey [404]"))
                    {
                        //User does not exist in Google Apps
                    }
                }

            }
        }

    }
    //==================================================================================================

    
    //class defaultCouplerCreateGoogleOU : defaultCoupler
    //{
    //    public defaultCouplerCreateGoogleOU
    //        (String csSHDDB, String csEHDDB, String csQLRDB, String aActionName, int aTryCount)
    //        : base(csSHDDB, csEHDDB, csQLRDB, aActionName, aTryCount) { }

    //    protected override void doSelectedItem(staffSpec aStaffSpec)
    //    {
    //                try
    //                {
    //                    // UserEntry ueCreate = service.CreateUser("atestLAN12347", "Steve", "Langley", "c0v3ntry");

    //                    UserEntry ue = service.RetrieveUser(aStaffSpec.NDSName.ToString());
    //                    AppsExtendedEntry OrgUser = serviceORG.RetrieveOrganizationUser("C02u22n9o", aStaffSpec.NDSName.ToString().Trim().ToLower() + "@warwickshire.ac.uk");
    //                    String CurrentOrgUnit = OrgUser.getPropertyValueByName("orgUnitPath").ToUpper().Trim();
    //                    AppsExtendedEntry OrgUserCreate = serviceORG.UpdateOrganizationUser("C02u22n9o", aStaffSpec.NDSName.ToString().Trim().ToLower() + "@warwickshire.ac.uk", "STAFF", CurrentOrgUnit.ToString());
    //                    mWritten.Add(aStaffSpec.queueItem.ToString());
    //                    //String[] queueItem;
    //                    //queueItem = new String[1];
    //                    //queueItem[0] = aStaffSpec.queueItem.ToString();
    //                    //int rowsChanged = staffUtility.updateCouplerMessageQueueSet(queueItem, " ");
    //                }
    //                catch (AppsException uexi)
    //                {
    //                    String msgi = uexi.Message;
    //                    String sReasoni = uexi.Reason.ToString();
    //                    if (uexi.ErrorCode == AppsException.UserDeletedRecently)
    //                    {
    //                        //User deleted recently - wait at least 5 days ok
    //                    }
    //                }
    //    }
    
    //protected override void testSelectedPhase(staffSpecSL aStaffSpecSL, DateTime runDateTime)
    //    {
    //        // AppsService service = new AppsService("warwickshire.ac.uk", "gadmin@warwickshire.ac.uk", "adcv325jemin");

    //        foreach (staffSpec wSS in aStaffSpecSL.Values)
    //        {

    //            try
    //            {
    //                UserEntry ue = service.RetrieveUser(wSS.NDSName.ToString());

    //                AppsExtendedEntry OrgUser = serviceORG.RetrieveOrganizationUser("C02u22n9o", wSS.NDSName.ToString().Trim() + "@warwickshire.ac.uk");
    //                if (OrgUser.getPropertyValueByName("orgUnitPath").ToUpper().Trim() == "STAFF")
    //                {
    //                    mTestedOK.Add(wSS.queueItem.ToString());
    //                }

    //            }
    //            catch (AppsException uex)
    //            {
    //                String msg = uex.Message;
    //                String sReason = uex.Reason.ToString();
    //                if (wSS.attempts > mTryCount) mFailed.Add(wSS.queueItem.ToString());
    //                if (uex.ErrorCode == AppsException.EntityDoesNotExist)
    //                {
    //                    //User does not exist in Google Apps
    //                }
    //            }

    //            //if (testSelectedItem(wSS, "T", "T", "Y"))
    //            //    mTestedOK.Add(wSS.queueItem.ToString());
    //            //else
    //            //    if (wSS.attempts > mTryCount) mFailed.Add(wSS.queueItem.ToString());
    //        }
    //    }

    //}


    // -------------------------------------------------------------------------------------------------------------
    class defaultCouplerRestoreGoogle : defaultCoupler
    {
        public defaultCouplerRestoreGoogle
            (String csSHDDB, String csEHDDB, String csQLRDB, String aActionName, int aTryCount)
            : base(csSHDDB, csEHDDB, csQLRDB, aActionName, aTryCount) { }


        protected override void doSelectedItem(staffSpec aStaffSpec)
        {
            //AppsService service = new AppsService("stu.warwickshire.ac.uk", "gadmin@stu.warwickshire.ac.uk", "thebigpic7ure");
            try
            {
                User u = wcService.Users.Get(aStaffSpec.NDSName.Trim().ToString() + "@warwickshire.ac.uk").Execute();
                u.Suspended = false;
                wcService.Users.Update(u, aStaffSpec.NDSName.Trim().ToString() + "@warwickshire.ac.uk").Execute();
                //UsersResource.PhotosResource.GetRequest ug = wcService.Users.Photos.Get(wSS.NDSName.Trim().ToString() + "@warwickshire.ac.uk");
                //upZ.PhotoData = ug.Execute().PhotoData;
                //upZ.MimeType = ug.Execute().MimeType;
                //upZ.Width = ug.Execute().Width;
                //upZ.Height = ug.Execute().Height;
                //upZ.Kind = ug.Execute().Kind;

                String sPhoto = "";
                Int64 EmployeeNumber = 0;

                Int64.TryParse(aStaffSpec.QLId, out EmployeeNumber);
                String IDM = "wcIDMGetStaffPhoto " + EmployeeNumber;
                DataView wDV = staffUtility.readDataView(staffUtility.couplerDB, IDM);
                if (wDV.Count > 0)
                {
                    sPhoto = wDV[0]["Base64Photo"].ToString();
                }
                if (sPhoto.Length > 100)// photo text less than 100 is probably not an image
                {
                    //Base64UrlEncode takes aparameter of base64 encode (this is retrieved from selectHR via a funtion wcIdentities.dbo.ToBase64(varbinary(max))
                    String StaffImage = staffUtility.Base64UrlEncode(sPhoto);
                    UserPhoto uphoto = new UserPhoto();
                    uphoto.Kind = "admin#directory#user#photo";
                    uphoto.Height = 96;
                    uphoto.Width = 96;
                    uphoto.MimeType = "image/bmp";
                    uphoto.PhotoData = StaffImage;
                    UsersResource.PhotosResource.UpdateRequest up = wcService.Users.Photos.Update(uphoto, aStaffSpec.NDSName.Trim().ToString() + "@warwickshire.ac.uk");
                    up.Execute();
                }

                mWritten.Add(aStaffSpec.queueItem.ToString());

            }
            catch (Google.GoogleApiException uex)
            {
                String msg = uex.Message;
                staffUtility.putLog("RestoreFailure DoItem: " + aStaffSpec.NDSName.Trim().ToString(), msg.ToString());
                //String sReason = uex.Reason.ToString();
                //bool sitePershore = false;
                if (msg.ToLower().Contains("userkey [404]"))
                {
                    //studentUtility.updateCouplerMessageQueue(aStudentSpec.queueItem.ToString(), "SET actionData = 'CMQ.UserDoesNotExistInGoogle'");
                    //User does not exist in Google Apps so we are done.
                    //mWritten.Add(aStudentSpec.queueItem.ToString());
                }
            }


            //updateSelectedItem(aStudentSpec, "T", "T", "Y");
        }


        protected override void testSelectedPhase(staffSpecSL aStaffSpecSL, DateTime runDateTime)
        {
            //AppsService service = new AppsService("stu.warwickshire.ac.uk", "gadmin@stu.warwickshire.ac.uk", "thebigpic7ure");
            



            foreach (staffSpec wSS in aStaffSpecSL.Values)
            {

                try
                {
                    User ue = wcService.Users.Get(wSS.NDSName.Trim().ToString() + "@warwickshire.ac.uk").Execute();
                    
                    wcService.Users.Update(ue, wSS.NDSName.Trim().ToString() + "@warwickshire.ac.uk").Execute();
                    
                    //UserEntry ue = service.RetrieveUser(wSS.NDSName.ToString());
                    // if fore name or surname need updating then do it here.
                    //if (ue.Aliases != null)
                    //{
                    //    if (ue.Aliases.Count > 0 )   
                    //    {
                    //        foreach (String myAlias in ue.Aliases)
                    //        {
                    //            //myAlias.ToString();
                    //            //Alias mAlias = new Alias();
                    //            //wcService.Users.Aliases.Insert(mAlias, wSS.NDSName.Trim().ToString() + "@warwickshire.ac.uk").Execute();
                    //            wcService.Users.Aliases.Delete(wSS.NDSName.Trim().ToString() + "@warwickshire.ac.uk", myAlias.ToString()).Execute();
                                
                    //        }
                    //    }
                    //}
                        
                    ////ue.PrimaryEmail = "WTEMPLAR@warwickshire.ac.uk";
                    //wcService.Users.Update(ue, "WTEMPLAR" + "@warwickshire.ac.uk").Execute();
                    //wcService.Users.Update(ue, wSS.NDSName.Trim().ToString() + "@warwickshire.ac.uk").Execute();


                    // SJL 24/04/2019 DHARRIS1 for some reason did not have a Name and the code was bugging on the String.Compare below
                    // added in this test for NULL Name and to create one if NULL
                    if (ue.Name == null)
                    {
                        UserName newusername = new UserName();
                        newusername.GivenName = wSS.Forename.ToString().Trim();
                        newusername.FamilyName = wSS.Surname.ToString().Trim();
                        ue.Name = newusername;
                        wcService.Users.Update(ue, wSS.NDSName.Trim().ToString() + "@warwickshire.ac.uk").Execute();
                    }


                    if ((String.Compare(ue.Name.FamilyName.ToLower().Trim().ToString(), wSS.Surname.ToLower().Trim().ToString()) != 0) || (String.Compare(ue.Name.GivenName.ToLower().Trim().ToString(), wSS.Forename.ToLower().Trim().ToString()) != 0))
                    {
                        ue.Name.FamilyName = wSS.Surname.Trim().ToString();
                        ue.Name.GivenName = wSS.Forename.Trim().ToString();
                        wcService.Users.Update(ue, wSS.NDSName.Trim().ToString() + "@warwickshire.ac.uk").Execute();

                        //if (wSS.ActionData.Trim().Length < 1) { } //ue.Login.Password = password; int rowsChanged = staffUtility.updateCouplerMessageQueueSet(queueItem, " SET actionData='"+password+"'");
                        //int passlength = ue.Password.Length;
                        //ue.Update();
                    }
                    //String gn2 = ue.Name.FamilyName.ToString();

                    if (ue.Suspended == false)
                    {
                        //studentUtility.updateCouplerMessageQueue(wSS.queueItem.ToString(), "SET actionData = ''");
                        mTestedOK.Add(wSS.queueItem.ToString());
                        staffUtility.writeCouplerMessageQueueV2(wSS.NDSName, "", "GoogleGroupAdd");

                    }
                    //UserEntry ue = service.RetrieveUser("atestlan12347");
                    //Write a coupler job for moveObject here
                    //studentUtility.writeCouplerMessageQueueStu(wSS.NDSName, "", "GoogleGroupAdd");

                }
                catch (Google.GoogleApiException uex)
                {
                    String msg = uex.Message;
                    staffUtility.putLog("RestoreFailure TestItem: " + wSS.NDSName.Trim().ToString(), msg.ToString());
                    //String sReason = uex.Reason.ToString();
                    if (wSS.attempts > mTryCount) mFailed.Add(wSS.queueItem.ToString());
                    if (msg.ToLower().Contains("userkey [404]"))
                    {
                        //User does not exist in Google Apps
                        //mTestedOK.Add(wSS.queueItem.ToString());

                    }
                }

            }
        }

    }
    //============================================================================================================
    // Restore Photgrapsh from HR
    class defaultCouplerRestoreGooglePhoto : defaultCoupler
    {
        public defaultCouplerRestoreGooglePhoto
            (String csSHDDB, String csEHDDB, String csQLRDB, String aActionName, int aTryCount)
            : base(csSHDDB, csEHDDB, csQLRDB, aActionName, aTryCount) { }


        protected override void doSelectedItem(staffSpec aStaffSpec)
        {
            try
            {

                mWritten.Add(aStaffSpec.queueItem.ToString());

            }
            catch (Google.GoogleApiException uex)
            {
                String msg = uex.Message;
                staffUtility.putLog("RestorePhotoFailure DoItem: " + aStaffSpec.NDSName.Trim().ToString(), msg.ToString());
                //String sReason = uex.Reason.ToString();
                //bool sitePershore = false;
                if (msg.ToLower().Contains("userkey [404]"))
                {
                    //mWritten.Add(aStudentSpec.queueItem.ToString());
                }
            }
        }


        protected override void testSelectedPhase(staffSpecSL aStaffSpecSL, DateTime runDateTime)
        {
            Boolean NoPhoto = false;
            foreach (staffSpec wSS in aStaffSpecSL.Values)
            {

                try
                {
                    User ue = wcService.Users.Get(wSS.NDSName.Trim().ToString() + "@warwickshire.ac.uk").Execute();
                    
                    wcService.Users.Update(ue, wSS.NDSName.Trim().ToString() + "@warwickshire.ac.uk").Execute();
                    UsersResource.PhotosResource.GetRequest dpGet = wcService.Users.Photos.Get(wSS.NDSName.Trim().ToString() + "@warwickshire.ac.uk");
                    String xyy = dpGet.ToString();
                    UserPhoto up = new UserPhoto();
                    try
                    {
                        up.PhotoData = dpGet.Execute().PhotoData;//if there is no photo then it will throw an exception.
                        
                    }
                    catch (Google.GoogleApiException uex)
                    {
                        if (uex.ToString().Contains("Not Found: photo"))
                        {
                            NoPhoto = true;
                        }
                    }
                    String sPhoto = "";
                    String StaffImage = "";
                    Int64 EmployeeNumber = 0;

                    Int64.TryParse(wSS.QLId, out EmployeeNumber);
                    String IDM = "wcIDMGetStaffPhoto " + EmployeeNumber;
                    DataView wDV = staffUtility.readDataView(staffUtility.couplerDB, IDM);// Get the staff photo from HR
                    if (wDV.Count > 0)
                    {
                        sPhoto = wDV[0]["Base64Photo"].ToString();
                    }
                    if (sPhoto.Length > 100)// photo text less than 100 is probably not an image
                    {
                        //Base64UrlEncode takes aparameter of base64 encode (this is retrieved from selectHR via a funtion wcIdentities.dbo.ToBase64(varbinary(max))
                        //sPhoto = "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";
                        StaffImage = staffUtility.Base64UrlEncode(sPhoto);
                        UserPhoto uphoto = new UserPhoto();
                        uphoto.Kind = "admin#directory#user#photo";
                        uphoto.Height = 96;
                        uphoto.Width = 96;
                        uphoto.MimeType = "image/bmp";
                        uphoto.PhotoData = StaffImage;
                        if (!NoPhoto)//If there is a photo then delete it first 
                        {
                            UsersResource.PhotosResource.DeleteRequest dpDelete = wcService.Users.Photos.Delete(wSS.NDSName.Trim().ToString() + "@warwickshire.ac.uk");
                            dpDelete.Execute();
                        }

                        UsersResource.PhotosResource.UpdateRequest dpUpdate = wcService.Users.Photos.Update(uphoto, wSS.NDSName.Trim().ToString() + "@warwickshire.ac.uk");
                        dpUpdate.Execute();
                        mTestedOK.Add(wSS.queueItem.ToString());
                    }
                    //UsersResource.PhotosResource.GetRequest ug = wcService.Users.Photos.Get(wSS.NDSName.Trim().ToString() + "@warwickshire.ac.uk");
                    //UserPhoto up = new UserPhoto();
                    
                    //up.PhotoData = ug.Execute().PhotoData;
                    //String x = ug.Execute().ToString();
                    wcService.Users.Update(ue, wSS.NDSName.Trim().ToString() + "@warwickshire.ac.uk").Execute();

                    //if (StaffImage.ToString().Trim() == up.PhotoData.ToString().Trim())
                    //{
                        //studentUtility.updateCouplerMessageQueue(wSS.queueItem.ToString(), "SET actionData = ''");
                        //mTestedOK.Add(wSS.queueItem.ToString());
                        
                    //}

                }
                catch (Google.GoogleApiException uex)
                {
                    String msg = uex.Message;
                    staffUtility.putLog("RestorePhotoFailure TestItem: " + wSS.NDSName.Trim().ToString(), msg.ToString());
                    //String sReason = uex.Reason.ToString();
                    if (wSS.attempts > mTryCount) mFailed.Add(wSS.queueItem.ToString());
                    if (msg.ToLower().Contains("userkey [404]"))
                    {
                        //User does not exist in Google Apps
                        //mTestedOK.Add(wSS.queueItem.ToString());

                    }
                }

            }
        }

    }
    // END Restore Photo
    //=============================================================================================================
    // ----------------------------------------------------------------------------------------------------------- 

    class defaultCouplerRemoveGoogleAlias : defaultCoupler
    {
        public defaultCouplerRemoveGoogleAlias
            (String csSHDDB, String csEHDDB, String csQLRDB, String aActionName, int aTryCount)
            : base(csSHDDB, csEHDDB, csQLRDB, aActionName, aTryCount) { }

        protected override void doSelectedItem(staffSpec aStaffSpec)
        {
            //AppsService service = new AppsService("stu.warwickshire.ac.uk", "gadmin@stu.warwickshire.ac.uk", "thebigpic7ure");
            try
            {
                User ue = wcService.Users.Get(aStaffSpec.NDSName.Trim().ToString() + "@warwickshire.ac.uk").Execute();
                if (ue.Aliases != null)
                {
                    if (ue.Aliases.Count > 0)
                    {
                        foreach (String myAlias in ue.Aliases)
                        {
                            //myAlias.ToString();
                            //Alias mAlias = new Alias();
                            //wcService.Users.Aliases.Insert(mAlias, wSS.NDSName.Trim().ToString() + "@warwickshire.ac.uk").Execute();
                            wcService.Users.Aliases.Delete(aStaffSpec.NDSName.Trim().ToString() + "@warwickshire.ac.uk", myAlias.ToString()).Execute();

                        }
                    }
                }

                mWritten.Add(aStaffSpec.queueItem.ToString());

            }
            catch (Google.GoogleApiException uex)
            {
                String msg = uex.Message;
                staffUtility.putLog("RestoreFailure DoItem: " + aStaffSpec.NDSName.Trim().ToString(), msg.ToString());
                //String sReason = uex.Reason.ToString();
                //bool sitePershore = false;
                if (msg.ToLower().Contains("userkey [404]"))
                {
                    //studentUtility.updateCouplerMessageQueue(aStudentSpec.queueItem.ToString(), "SET actionData = 'CMQ.UserDoesNotExistInGoogle'");
                    //User does not exist in Google Apps so we are done.
                    //mWritten.Add(aStudentSpec.queueItem.ToString());
                }
            }


            //updateSelectedItem(aStudentSpec, "T", "T", "Y");
        }


        protected override void testSelectedPhase(staffSpecSL aStaffSpecSL, DateTime runDateTime)
        {
            //AppsService service = new AppsService("stu.warwickshire.ac.uk", "gadmin@stu.warwickshire.ac.uk", "thebigpic7ure");
            Boolean aliasExist = false;
            foreach (staffSpec wSS in aStaffSpecSL.Values)
            {

                try
                {
                    User ue = wcService.Users.Get(wSS.NDSName.Trim().ToString() + "@warwickshire.ac.uk").Execute();

                    //UserEntry ue = service.RetrieveUser(wSS.NDSName.ToString());
                    // if fore name or surname need updating then do it here.
                    if (ue.Aliases != null)
                    {
                        if (ue.Aliases.Count > 0 )
                        {

                            aliasExist = true;
                        }
                    }

                    if (aliasExist == false)
                    {
                        //studentUtility.updateCouplerMessageQueue(wSS.queueItem.ToString(), "SET actionData = ''");
                        mTestedOK.Add(wSS.queueItem.ToString());

                    }
                    //UserEntry ue = service.RetrieveUser("atestlan12347");
                    //Write a coupler job for moveObject here
                    //studentUtility.writeCouplerMessageQueueStu(wSS.NDSName, "", "GoogleGroupAdd");

                }
                catch (Google.GoogleApiException uex)
                {
                    String msg = uex.Message;
                    staffUtility.putLog("AliasExistFailure TestItem: " + wSS.NDSName.Trim().ToString(), msg.ToString());
                    //String sReason = uex.Reason.ToString();
                    if (wSS.attempts > mTryCount) mFailed.Add(wSS.queueItem.ToString());
                    if (msg.ToLower().Contains("userkey [404]"))
                    {
                        //User does not exist in Google Apps
                        //mTestedOK.Add(wSS.queueItem.ToString());

                    }
                }

            }
        }

    }
    // ----------------------------------------------------------------------------------------------------------- 

    //=============================================================================================================
    // ----------------------------------------------------------------------------------------------------------- 

    class defaultCouplerRenameGoogleEmail : defaultCoupler
    {
        public defaultCouplerRenameGoogleEmail
            (String csSHDDB, String csEHDDB, String csQLRDB, String aActionName, int aTryCount)
            : base(csSHDDB, csEHDDB, csQLRDB, aActionName, aTryCount) { }

        protected override void doSelectedItem(staffSpec aStaffSpec)
        {
            //AppsService service = new AppsService("stu.warwickshire.ac.uk", "gadmin@stu.warwickshire.ac.uk", "thebigpic7ure");
            try
            {
                User ue = wcService.Users.Get(aStaffSpec.NDSName.Trim().ToString() + "@warwickshire.ac.uk").Execute();
                wcService.Users.Update(ue, aStaffSpec.ActionData.Trim().ToString() + "@warwickshire.ac.uk").Execute();

                mWritten.Add(aStaffSpec.queueItem.ToString());

            }
            catch (Google.GoogleApiException uex)
            {
                String msg = uex.Message;
                staffUtility.putLog("RestoreFailure DoItem: " + aStaffSpec.NDSName.Trim().ToString(), msg.ToString());
                //String sReason = uex.Reason.ToString();
                //bool sitePershore = false;
                if (msg.ToLower().Contains("userkey [404]"))
                {
                    //studentUtility.updateCouplerMessageQueue(aStudentSpec.queueItem.ToString(), "SET actionData = 'CMQ.UserDoesNotExistInGoogle'");
                    //User does not exist in Google Apps so we are done.
                    //mWritten.Add(aStudentSpec.queueItem.ToString());
                }
            }


            //updateSelectedItem(aStudentSpec, "T", "T", "Y");
        }


        protected override void testSelectedPhase(staffSpecSL aStaffSpecSL, DateTime runDateTime)
        {
            //AppsService service = new AppsService("stu.warwickshire.ac.uk", "gadmin@stu.warwickshire.ac.uk", "thebigpic7ure");
            
            foreach (staffSpec wSS in aStaffSpecSL.Values)
            {

                try
                {
                    // test for the existence of the new email address fro the 'actiondata' 
                    User ue = wcService.Users.Get(wSS.ActionData.Trim().ToString() + "@warwickshire.ac.uk").Execute();

                    mTestedOK.Add(wSS.queueItem.ToString());

 
                }
                catch (Google.GoogleApiException uex)
                {
                    String msg = uex.Message;
                    staffUtility.putLog("RenameEmailExistFailure TestItem: " + wSS.NDSName.Trim().ToString(), msg.ToString());
                    //String sReason = uex.Reason.ToString();
                    if (wSS.attempts > mTryCount) mFailed.Add(wSS.queueItem.ToString());
                    if (msg.ToLower().Contains("userkey [404]"))
                    {
                        //User does not exist in Google Apps
                        //mTestedOK.Add(wSS.queueItem.ToString());

                    }
                }

            }
        }

    }
    // ----------------------------------------------------------------------------------------------------------- 

    //=============================================================================================================

    
    class defaultCouplerRestoreGoogleOrgUnit : defaultCoupler
    {
        public defaultCouplerRestoreGoogleOrgUnit
            (String csSHDDB, String csEHDDB, String csQLRDB, String aActionName, int aTryCount)
            : base(csSHDDB, csEHDDB, csQLRDB, aActionName, aTryCount) { }

        protected override void doSelectedItem(staffSpec aStaffSpec)
        {
            try
            {
                User u = wcService.Users.Get(aStaffSpec.NDSName.Trim().ToString() + "@warwickshire.ac.uk").Execute();
                u.OrgUnitPath = "/STAFF/MISC";
                wcService.Users.Update(u, aStaffSpec.NDSName.Trim().ToString() + "@warwickshire.ac.uk").Execute();

                mWritten.Add(aStaffSpec.queueItem.ToString());

            }
            catch (Google.GoogleApiException uex)
            {
                String msg = uex.Message;
                staffUtility.putLog("RestoreOrgUnitFailure DoItem: " + aStaffSpec.NDSName.Trim().ToString(), msg.ToString());
                //String sReason = uex.Reason.ToString();
                //bool sitePershore = false;
                if (msg.ToLower().Contains("userkey [404]"))
                {
                    //studentUtility.updateCouplerMessageQueue(aStudentSpec.queueItem.ToString(), "SET actionData = 'CMQ.UserDoesNotExistInGoogle'");
                    //User does not exist in Google Apps so we are done.
                    //mWritten.Add(aStudentSpec.queueItem.ToString());
                }
            }


            //updateSelectedItem(aStudentSpec, "T", "T", "Y");
        }


        protected override void testSelectedPhase(staffSpecSL aStaffSpecSL, DateTime runDateTime)
        {

            foreach (staffSpec wSS in aStaffSpecSL.Values)
            {

                try
                {
                    User ue = wcService.Users.Get(wSS.NDSName.Trim().ToString().ToLower() + "@warwickshire.ac.uk").Execute();

                    if ((String.Compare(ue.Name.FamilyName.ToLower().Trim().ToString(), wSS.Surname.ToLower().Trim().ToString()) != 0) || (String.Compare(ue.Name.GivenName.ToLower().Trim().ToString(), wSS.Forename.ToLower().Trim().ToString()) != 0))
                    {
                        ue.Name.FamilyName = wSS.Surname.Trim().ToString();
                        ue.Name.GivenName = wSS.Forename.Trim().ToString();
                        wcService.Users.Update(ue, wSS.NDSName.Trim().ToString() + "@warwickshire.ac.uk").Execute();
                    }

                    if (ue.OrgUnitPath == "/STAFF/MISC")
                    {
                        mTestedOK.Add(wSS.queueItem.ToString());

                    }

                }
                catch (Google.GoogleApiException uex)
                {
                    String msg = uex.Message;
                    staffUtility.putLog("RestoreOrgUnitFailure TestItem: " + wSS.NDSName.Trim().ToString(), msg.ToString());
                    //String sReason = uex.Reason.ToString();
                    if (wSS.attempts > mTryCount) mFailed.Add(wSS.queueItem.ToString());
                    if (msg.ToLower().Contains("userkey [404]"))
                    {
                        //User does not exist in Google Apps
                        //mTestedOK.Add(wSS.queueItem.ToString());

                    }
                }

            }
        }

    }
    // ----------------------------------------------------------------------------------------------------------- 
// ================================================================================================================
    class defaultCouplerMoveGoogleOrgUnit : defaultCoupler
    {
        // Accounts that are 'MOVED' are essentially now 'DEAD'
        // These accounts can never be used again.
        // Accounts are moved 90 days after their network accounts have been inactive.
        // After 7 Years we will then delete these accounts.
        // If the user of this account comes back then they have to get a brand new network account and Google email account.
        //
        public defaultCouplerMoveGoogleOrgUnit
            (String csSHDDB, String csEHDDB, String csQLRDB, String aActionName, int aTryCount)
            : base(csSHDDB, csEHDDB, csQLRDB, aActionName, aTryCount) { }

        protected override void doSelectedItem(staffSpec aStaffSpec)
        {
            try
            {
                User u = wcService.Users.Get(aStaffSpec.NDSName.Trim().ToString() + "@warwickshire.ac.uk").Execute();
                u.OrgUnitPath = "/STAFF_INACTIVE";
                wcService.Users.Update(u, aStaffSpec.NDSName.Trim().ToString() + "@warwickshire.ac.uk").Execute();

                mWritten.Add(aStaffSpec.queueItem.ToString());

            }
            catch (Google.GoogleApiException uex)
            {
                String msg = uex.Message;
                staffUtility.putLog("MoveOrgUnitFailure DoItem: " + aStaffSpec.NDSName.Trim().ToString(), msg.ToString());
                //String sReason = uex.Reason.ToString();
                //bool sitePershore = false;
                if (msg.ToLower().Contains("userkey [404]"))
                {
                    String[] qiSet  = new String[1];
                    qiSet[0] = aStaffSpec.queueItem.ToString();
                    staffUtility.updateCouplerMessageQueueSet(qiSet, "SET actionData = 'CMQ.UserDoesNotExistInGoogle'");
                    //User does not exist in Google Apps so we are done.
                    //mWritten.Add(aStudentSpec.queueItem.ToString());
                }
            }


            //updateSelectedItem(aStudentSpec, "T", "T", "Y");
        }


        protected override void testSelectedPhase(staffSpecSL aStaffSpecSL, DateTime runDateTime)
        {

            foreach (staffSpec wSS in aStaffSpecSL.Values)
            {

                try
                {
                    User ue = wcService.Users.Get(wSS.NDSName.Trim().ToString().ToLower() + "@warwickshire.ac.uk").Execute();

                    if ((String.Compare(ue.Name.FamilyName.ToLower().Trim().ToString(), wSS.Surname.ToLower().Trim().ToString()) != 0) || (String.Compare(ue.Name.GivenName.ToLower().Trim().ToString(), wSS.Forename.ToLower().Trim().ToString()) != 0))
                    {
                        ue.Name.FamilyName = wSS.Surname.Trim().ToString();
                        ue.Name.GivenName = wSS.Forename.Trim().ToString();
                        wcService.Users.Update(ue, wSS.NDSName.Trim().ToString() + "@warwickshire.ac.uk").Execute();
                    }

                    if (ue.OrgUnitPath == "/STAFF_INACTIVE")
                    {
                        //remove the email address from HR with procedure wcIDMRemoveStaffEmail employeeNumber

                        
                        String UserHREmailDel = "wcIDMRemoveStaffEmail '" + wSS.QLId.ToString() + "'";
                        DataView wDV = staffUtility.readDataView(staffUtility.couplerDB, UserHREmailDel);

                        mTestedOK.Add(wSS.queueItem.ToString());

                    }

                }
                catch (Google.GoogleApiException uex)
                {
                    String msg = uex.Message;
                    staffUtility.putLog("MoveOrgUnitFailure TestItem: " + wSS.NDSName.Trim().ToString(), msg.ToString());
                    //String sReason = uex.Reason.ToString();
                    if (wSS.attempts > mTryCount) mFailed.Add(wSS.queueItem.ToString());
                    if (msg.ToLower().Contains("userkey [404]"))
                    {
                        //User does not exist in Google Apps
                        //mTestedOK.Add(wSS.queueItem.ToString());

                    }
                }

            }
        }

    }

//-----------------------------------------------------------------------------------------------------------------------------
    // ================================================================================================================
    class defaultCouplerReInstateGoogleOrgUnit : defaultCoupler
    {
        // Accounts that are 'MOVED' are essentially now 'DEAD'
        // These accounts can never be used again.
        // Accounts are moved 90 days after their network accounts have been inactive.
        // After 7 Years we will then delete these accounts.
        // If the user of this account comes back then they have to get a brand new network account and Google email account.
        //
        public defaultCouplerReInstateGoogleOrgUnit
            (String csSHDDB, String csEHDDB, String csQLRDB, String aActionName, int aTryCount)
            : base(csSHDDB, csEHDDB, csQLRDB, aActionName, aTryCount) { }

        protected override void doSelectedItem(staffSpec aStaffSpec)
        {
            try
            {
                User u = wcService.Users.Get(aStaffSpec.NDSName.Trim().ToString() + "@warwickshire.ac.uk").Execute();
                u.OrgUnitPath = "/STAFF";
                wcService.Users.Update(u, aStaffSpec.NDSName.Trim().ToString() + "@warwickshire.ac.uk").Execute();

                mWritten.Add(aStaffSpec.queueItem.ToString());

            }
            catch (Google.GoogleApiException uex)
            {
                String msg = uex.Message;
                staffUtility.putLog("MoveOrgUnitFailure DoItem: " + aStaffSpec.NDSName.Trim().ToString(), msg.ToString());
                //String sReason = uex.Reason.ToString();
                //bool sitePershore = false;
                if (msg.ToLower().Contains("userkey [404]"))
                {
                    String[] qiSet = new String[1];
                    qiSet[0] = aStaffSpec.queueItem.ToString();
                    staffUtility.updateCouplerMessageQueueSet(qiSet, "SET actionData = 'CMQ.UserDoesNotExistInGoogle'");
                    //User does not exist in Google Apps so we are done.
                    //mWritten.Add(aStudentSpec.queueItem.ToString());
                }
            }


            //updateSelectedItem(aStudentSpec, "T", "T", "Y");
        }


        protected override void testSelectedPhase(staffSpecSL aStaffSpecSL, DateTime runDateTime)
        {

            foreach (staffSpec wSS in aStaffSpecSL.Values)
            {

                try
                {
                    User ue = wcService.Users.Get(wSS.NDSName.Trim().ToString().ToLower() + "@warwickshire.ac.uk").Execute();

                    if ((String.Compare(ue.Name.FamilyName.ToLower().Trim().ToString(), wSS.Surname.ToLower().Trim().ToString()) != 0) || (String.Compare(ue.Name.GivenName.ToLower().Trim().ToString(), wSS.Forename.ToLower().Trim().ToString()) != 0))
                    {
                        ue.Name.FamilyName = wSS.Surname.Trim().ToString();
                        ue.Name.GivenName = wSS.Forename.Trim().ToString();
                        wcService.Users.Update(ue, wSS.NDSName.Trim().ToString() + "@warwickshire.ac.uk").Execute();
                    }

                    if (ue.OrgUnitPath == "/STAFF")
                    {
                        //remove the email address from HR with procedure wcIDMRemoveStaffEmail employeeNumber


                        String UserHREmailDel = "wcIDMReInstateStaffEmail '" + wSS.QLId.ToString() + "'";
                        DataView wDV = staffUtility.readDataView(staffUtility.couplerDB, UserHREmailDel);

                        mTestedOK.Add(wSS.queueItem.ToString());

                    }

                }
                catch (Google.GoogleApiException uex)
                {
                    String msg = uex.Message;
                    staffUtility.putLog("MoveOrgUnitFailure TestItem: " + wSS.NDSName.Trim().ToString(), msg.ToString());
                    //String sReason = uex.Reason.ToString();
                    if (wSS.attempts > mTryCount) mFailed.Add(wSS.queueItem.ToString());
                    if (msg.ToLower().Contains("userkey [404]"))
                    {
                        //User does not exist in Google Apps
                        //mTestedOK.Add(wSS.queueItem.ToString());

                    }
                }

            }
        }

    }


    //--------------------------------------------------------------------------------------------------------------
    
    
    class defaultCouplerSuspendGoogle : defaultCoupler
    {
        public defaultCouplerSuspendGoogle
            (String csSHDDB, String csEHDDB, String csQLRDB, String aActionName, int aTryCount)
            : base(csSHDDB, csEHDDB, csQLRDB, aActionName, aTryCount) { }

        protected override void doSelectedItem(staffSpec aStaffSpec)
        {
            //AppsService service = new AppsService("stu.warwickshire.ac.uk", "gadmin@stu.warwickshire.ac.uk", "thebigpic7ure");
            try
            {

                User u = wcService.Users.Get(aStaffSpec.NDSName.Trim().ToString() + "@warwickshire.ac.uk").Execute();
                u.Suspended = true;
                wcService.Users.Update(u, aStaffSpec.NDSName.Trim().ToString() + "@warwickshire.ac.uk").Execute();
                mWritten.Add(aStaffSpec.queueItem.ToString());
            }
            catch (Google.GoogleApiException uex)
            {
                String msg = uex.Message;
                staffUtility.putLog("SuspendFailure DoItem: " + aStaffSpec.NDSName.Trim().ToString(), msg.ToString());
                //String sReason = uex.Reason.ToString();
                //bool sitePershore = false;
                if (msg.ToLower().Contains("userkey [404]"))
                {
                    //User does not exist in Google Apps so we are done.
                    //studentUtility.updateCouplerMessageQueue(aStudentSpec.queueItem.ToString(),"SET actionData = 'CMQ.UserDoesNotExistInGoogle'");
                    //mWritten.Add(aStudentSpec.queueItem.ToString());
                }
            }


            //updateSelectedItem(aStudentSpec, "T", "T", "Y");
        }


        protected override void testSelectedPhase(staffSpecSL aStaffSpecSL, DateTime runDateTime)
        {
            //AppsService service = new AppsService("stu.warwickshire.ac.uk", "gadmin@stu.warwickshire.ac.uk", "thebigpic7ure");

            foreach (staffSpec wSS in aStaffSpecSL.Values)
            {

                try
                {
                    User ue = wcService.Users.Get(wSS.NDSName.Trim().ToString() + "@warwickshire.ac.uk").Execute();
                    if (ue.Suspended == true)
                    {
                        //staffUtility.updateCouplerMessageQueue(wSS.queueItem.ToString(), "SET actionData = ''");
                        mTestedOK.Add(wSS.queueItem.ToString());
                        staffUtility.writeCouplerMessageQueueV2(wSS.NDSName, "", "DeleteGoogleGroups");
                    }
                    //UserEntry ue = service.RetrieveUser("atestlan12347");
                    //Write a coupler job for moveObject here
                    //studentUtility.writeCouplerMessageQueueStu(wSS.NDSName, "", "GoogleGroupAdd");

                }
                catch (Google.GoogleApiException uex)
                {
                    String msg = uex.Message;
                    staffUtility.putLog("SuspendFailure TestItem: " + wSS.NDSName.Trim().ToString(), msg.ToString());
                    //String sReason = uex.Reason.ToString();
                    if (wSS.attempts > mTryCount) mFailed.Add(wSS.queueItem.ToString());
                    if (msg.ToLower().Contains("userkey [404]"))
                    {
                        //User does not exist in Google Apps
                        //mTestedOK.Add(wSS.queueItem.ToString());

                    }
                }

            }//for each
        }
    }
    
    // -----------------------------------------------------------------------------------------------------------

    class defaultCouplerDeleteGoogle : defaultCoupler
    {
        public defaultCouplerDeleteGoogle
            (String csSHDDB, String csEHDDB, String csQLRDB, String aActionName, int aTryCount)
            : base(csSHDDB, csEHDDB, csQLRDB, aActionName, aTryCount) { }

        protected override void doSelectedItem(staffSpec aStaffSpec)
        {
            //AppsService service = new AppsService("stu.warwickshire.ac.uk", "gadmin@stu.warwickshire.ac.uk", "thebigpic7ure");
            try
            {
                wcService.Users.Delete(aStaffSpec.NDSName.ToString().Trim() + "@warwickshire.ac.uk").Execute();
                
                //UserEntry ue = service.RetrieveUser(aStaffSpec.NDSName.ToString());
                //service.DeleteUser(aStaffSpec.NDSName.ToString());
                mWritten.Add(aStaffSpec.queueItem.ToString());
            }
            catch (Google.GoogleApiException uex)
            {
                String msg = uex.Message;
                staffUtility.putLog("DeleteFailure Account does not exist. DoItem: " + aStaffSpec.NDSName.Trim().ToString(), msg.ToString());
                //String sReason = uex.Reason.ToString();
                //bool sitePershore = false;
                if (msg.ToLower().Contains("userkey [404]"))
                {
                    //User does not exist in Google Apps so we are done.
                    mWritten.Add(aStaffSpec.queueItem.ToString());
                }
            }


            //updateSelectedItem(aStudentSpec, "T", "T", "Y");
        }


        protected override void testSelectedPhase(staffSpecSL aStaffSpecSL, DateTime runDateTime)
        {
            //AppsService service = new AppsService("stu.warwickshire.ac.uk", "gadmin@stu.warwickshire.ac.uk", "thebigpic7ure");

            foreach (staffSpec wSS in aStaffSpecSL.Values)
            {

                try
                {
                    User u = wcService.Users.Get(wSS.NDSName.Trim().ToString() + "@warwickshire.ac.uk").Execute();
                    //UserEntry ue = service.RetrieveUser(wSS.NDSName.ToString());
                    //UserEntry ue = service.RetrieveUser("atestlan12347");
                    //Write a coupler job for moveObject here
                    //studentUtility.writeCouplerMessageQueueStu(wSS.NDSName, "", "GoogleGroupAdd");

                }
                catch (Google.GoogleApiException uex)
                {
                    String msg = uex.Message;
                    staffUtility.putLog("Delete process Account does not exist TestItem: " + wSS.NDSName.Trim().ToString(), msg.ToString());
                    //String sReason = uex.Reason.ToString();
                    if (wSS.attempts > mTryCount) mFailed.Add(wSS.queueItem.ToString());
                    if (msg.ToLower().Contains("userkey [404]"))
                    {
                        //User does not exist in Google Apps
                        //if (wSS.courseTitle.ToString().ToLower().Trim() == "deleted nds record") // this means that the NDS record has been deleted
                        //{
                            //create a coupler job to remove the wcStudentPrimaryCourse record
                            //studentUtility.writeCouplerMessageQueueStu(wSS.NDSName, "", "DeleteWcSPCourse");
                            mTestedOK.Add(wSS.queueItem.ToString());
                        //}

                    }
                }

            }
        }

    }


    // ============================================================================================================

    class defaultCouplerChangeDetails : defaultCoupler
    {
        public defaultCouplerChangeDetails
            (String csSHDDB, String csEHDDB, String csQLRDB, String aActionName, int aTryCount)
            : base(csSHDDB, csEHDDB, csQLRDB, aActionName, aTryCount) { }

        protected override void doSelectedItem(staffSpec aStaffSpec)
        {
            updateSelectedItem(aStaffSpec, "%", "%", "%");
        }

        protected override void testSelectedPhase(staffSpecSL aStaffSpecSL, DateTime runDateTime)
        {
            //DataView wDV;
            foreach (staffSpec wSS in aStaffSpecSL.Values)
            {

                if (testSelectedItem(wSS, "T", "T", "Y"))
                    mTestedOK.Add(wSS.queueItem.ToString());
                else
                    if (wSS.attempts > mTryCount) mFailed.Add(wSS.queueItem.ToString());
            }

        }
    }
}
