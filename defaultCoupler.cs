using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using Google.GData.Apps;


namespace couplerWriter
{
    class defaultCoupler : coupler
    {
        protected String mCSSHDDB, mCSEHDDB, mCSQLRDB;

        public defaultCoupler() : base() { }

        public AppsService service;
        public AppsService serviceTEST;

        public defaultCoupler
            (String csEHDDB, String csSHDDB, String csQLRDB, String aActionName, int aTryCount)
        {
            mCSEHDDB = csEHDDB; mCSQLRDB = csQLRDB; mCSSHDDB = csSHDDB;
            mActionName = aActionName; mTryCount = aTryCount;
            try
            {
                service = new AppsService("warwickshire.ac.uk", "system@warwickshire.ac.uk", "b4ckb34t");
                serviceTEST = new AppsService("test.warwickshire.ac.uk", "system@warwickshire.ac.uk", "b4ckb34t");

                //service = new AppsService("warwickshire.ac.uk", "gadmin@warwickshire.ac.uk", "adcv325jemin");
                //serviceTEST = new AppsService("test.warwickshire.ac.uk", "gadmin@warwickshire.ac.uk", "adcv325jemin");
            }
            catch (AppsException uex)
            {
                String msg = uex.Message;
                String sReason = uex.Reason.ToString();
            }
            catch (System.Net.WebException webEx)
            {
                String webMsg = webEx.Message.ToString();
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

    class defaultCouplerCreateGoogleGroup : defaultCoupler
    {
        public defaultCouplerCreateGoogleGroup
            (String csSHDDB, String csEHDDB, String csQLRDB, String aActionName, int aTryCount)
            : base(csSHDDB, csEHDDB, csQLRDB, aActionName, aTryCount) { }
        //  create a SP for returning ALL sites for a studentID hat ^ delimitted
        // or add extra column in wcStudentPrimaryCourse ?? maybe not.

        protected override void doSelectedItem(staffSpec aStaffSpec)
        {
            //try
            //{
            //    AppsService service = new AppsService("stu.warwickshire.ac.uk", "gadmin@stu.warwickshire.ac.uk", "thebigpic7ure");
            //}
            //catch (AppsException uex)
            //{
            //    String msg = uex.Message;
            //    String sReason = uex.Reason.ToString();
            //}

            try
            {
                UserEntry ue = service.RetrieveUser(aStaffSpec.NDSName.ToString());
                //service.

            }
            catch (AppsException uex)
            {
                String msg = uex.Message;
                String sReason = uex.Reason.ToString();
                if (uex.ErrorCode == AppsException.EntityDoesNotExist)
                {
                    //User does not exist in Google Apps so create the user ?
                }
            }
            catch (System.Net.WebException webEx)
            {
                String webMsg = webEx.Message.ToString();
            }
            catch (Exception gEx)
            {
                String gMsg = gEx.Message.ToString();
            }

            //user exists.
            bool rls = false;
            bool mm = false;
            bool trident = false;
            bool rugby = false;
            bool henley = false;
            bool pershore = false;
            bool googleRls = false;
            bool googleMM = false;
            bool googleTrident = false;
            bool googleRugby = false;
            bool googleHenley = false;
            bool googlePershore = false;
            try
            {
                // Add user to group 

                if (service.Groups.IsMember(aStaffSpec.NDSName.ToString(), "STAFF-ALLSITES") == false)
                {
                    try
                    {
                        //AppsExtendedFeed groups = service.Groups.RetrieveGroups(aStaffSpec.NDSName.ToString(), true);
                        service.Groups.AddMemberToGroup(aStaffSpec.NDSName.ToString(), "STAFF-ALLSITES");
                    }
                    catch (AppsException uex)
                    {
                        String msg = uex.Message;
                        String sReason = uex.Reason.ToString();
                    }
                    catch (System.Net.WebException webEx)
                    {
                        String webMsg = webEx.Message.ToString();
                    }
                    catch (Exception gEx)
                    {
                        String gMsg = gEx.Message.ToString();
                    }

                }
                // *jj
                //DataView wCouplerDV = staffUtility.readDataView(staffUtility.couplerDB, "wcStudentSites '" + aStaffSpec.NDSName.ToString() + "'");
                //if (wCouplerDV.Count > 0)
                {
                    // start pershore
                    if (aStaffSpec.LocID.ToString().ToUpper().Trim().IndexOf("PERSHORE") >= 0) pershore = true;
                    googlePershore = service.Groups.IsMember(aStaffSpec.NDSName.ToString(), "STAFF-PER");

                    if (googlePershore && !pershore)
                    {
                        try
                        {
                            service.Groups.RemoveMemberFromGroup(aStaffSpec.NDSName.ToString(), "STAFF-PER");
                        }
                        catch (AppsException uex)
                        {
                            String msg = uex.Message;
                            String sReason = uex.Reason.ToString();
                        }
                        catch (System.Net.WebException webEx)
                        {
                            String webMsg = webEx.Message.ToString();
                        }
                        catch (Exception gEx)
                        {
                            String gMsg = gEx.Message.ToString();
                        }

                    }
                    else
                    {
                        if (pershore && !googlePershore)
                        {
                            try
                            {
                                service.Groups.AddMemberToGroup(aStaffSpec.NDSName.ToString(), "STAFF-PER");
                            }
                            catch (AppsException uex)
                            {
                                String msg = uex.Message;
                                String sReason = uex.Reason.ToString();
                            }
                            catch (System.Net.WebException webEx)
                            {
                                String webMsg = webEx.Message.ToString();
                            }
                            catch (Exception gEx)
                            {
                                String gMsg = gEx.Message.ToString();
                            }

                        }
                    }
                    // end pershore

                    // start lspa
                    if (aStaffSpec.LocID.ToString().ToUpper().Trim().IndexOf("LEAMINGTON SPA") >= 0) rls = true;
                    //UserEntry steve = service.RetrieveUser("SLANGLEY");
                    //AppsExtendedFeed allmembers = service.Groups.RetrieveAllMembers("lspa");
                    //AppsExtendedFeed groups = service.Groups.RetrieveAllGroups();
                    googleRls = service.Groups.IsMember(aStaffSpec.NDSName.ToString(), "STAFF-LSPA");
                    if (googleRls && !rls)
                    {
                        try
                        {
                            service.Groups.RemoveMemberFromGroup(aStaffSpec.NDSName.ToString(), "STAFF-LSPA");
                        }
                        catch (AppsException uex)
                        {
                            String msg = uex.Message;
                            String sReason = uex.Reason.ToString();
                        }
                        catch (System.Net.WebException webEx)
                        {
                            String webMsg = webEx.Message.ToString();
                        }
                        catch (Exception gEx)
                        {
                            String gMsg = gEx.Message.ToString();
                        }

                    }
                    else
                    {
                        if (rls && !googleRls)
                        {
                            try
                            {
                                service.Groups.AddMemberToGroup(aStaffSpec.NDSName.ToString(), "STAFF-LSPA");
                            }
                            catch (AppsException uex)
                            {
                                String msg = uex.Message;
                                String sReason = uex.Reason.ToString();
                            }
                            catch (System.Net.WebException webEx)
                            {
                                String webMsg = webEx.Message.ToString();
                            }
                            catch (Exception gEx)
                            {
                                String gMsg = gEx.Message.ToString();
                            }

                        }
                    }
                    // end lspa

                    // start rugby
                    if (aStaffSpec.LocID.ToString().Trim().ToUpper().IndexOf("RUGBY") >= 0) rugby = true;

                    googleRugby = service.Groups.IsMember(aStaffSpec.NDSName.ToString(), "STAFF-RUG");
                    if (googleRugby && !rugby)
                    {
                        try
                        {
                            service.Groups.RemoveMemberFromGroup(aStaffSpec.NDSName.ToString(), "STAFF-RUG");
                        }
                        catch (AppsException uex)
                        {
                            String msg = uex.Message;
                            String sReason = uex.Reason.ToString();
                        }
                        catch (System.Net.WebException webEx)
                        {
                            String webMsg = webEx.Message.ToString();
                        }
                        catch (Exception gEx)
                        {
                            String gMsg = gEx.Message.ToString();
                        }

                    }
                    else
                    {
                        if (rugby && !googleRugby)
                        {
                            try
                            {
                                service.Groups.AddMemberToGroup(aStaffSpec.NDSName.ToString(), "STAFF-RUG");
                            }
                            catch (AppsException uex)
                            {
                                String msg = uex.Message;
                                String sReason = uex.Reason.ToString();
                            }
                            catch (System.Net.WebException webEx)
                            {
                                String webMsg = webEx.Message.ToString();
                            }
                            catch (Exception gEx)
                            {
                                String gMsg = gEx.Message.ToString();
                            }

                        }
                    }
                    // end rugby
                    // start moreton morrell
                    if (aStaffSpec.LocID.ToString().Trim().ToUpper().IndexOf("MORETON MORRELL") >= 0) mm = true;

                    googleMM = service.Groups.IsMember(aStaffSpec.NDSName.ToString(), "STAFF-MM");
                    if (googleMM && !mm)
                    {
                        try
                        {
                            service.Groups.RemoveMemberFromGroup(aStaffSpec.NDSName.ToString(), "STAFF-MM");
                        }
                        catch (AppsException uex)
                        {
                            String msg = uex.Message;
                            String sReason = uex.Reason.ToString();
                        }
                        catch (System.Net.WebException webEx)
                        {
                            String webMsg = webEx.Message.ToString();
                        }
                        catch (Exception gEx)
                        {
                            String gMsg = gEx.Message.ToString();
                        }

                    }
                    else
                    {
                        if (mm && !googleMM)
                        {
                            try
                            {
                                service.Groups.AddMemberToGroup(aStaffSpec.NDSName.ToString(), "STAFF-MM");
                            }
                            catch (AppsException uex)
                            {
                                String msg = uex.Message;
                                String sReason = uex.Reason.ToString();
                            }
                            catch (System.Net.WebException webEx)
                            {
                                String webMsg = webEx.Message.ToString();
                            }
                            catch (Exception gEx)
                            {
                                String gMsg = gEx.Message.ToString();
                            }

                        }
                    }
                    // end moreton morrell

                    // start trident
                    if (aStaffSpec.LocID.ToString().ToUpper().Trim().IndexOf("TRIDENT PARK") >= 0) trident = true;

                    googleTrident = service.Groups.IsMember(aStaffSpec.NDSName.ToString(), "STAFF-TRIDENT");
                    if (googleTrident && !trident)
                    {
                        try
                        {
                            service.Groups.RemoveMemberFromGroup(aStaffSpec.NDSName.ToString(), "STAFF-TRIDENT");
                        }
                        catch (AppsException uex)
                        {
                            String msg = uex.Message;
                            String sReason = uex.Reason.ToString();
                        }
                        catch (System.Net.WebException webEx)
                        {
                            String webMsg = webEx.Message.ToString();
                        }
                        catch (Exception gEx)
                        {
                            String gMsg = gEx.Message.ToString();
                        }

                    }
                    else
                    {
                        if (trident && !googleTrident)
                        {
                            try
                            {
                                service.Groups.AddMemberToGroup(aStaffSpec.NDSName.ToString(), "STAFF-TRIDENT");
                            }
                            catch (AppsException uex)
                            {
                                String msg = uex.Message;
                                String sReason = uex.Reason.ToString();
                            }
                            catch (System.Net.WebException webEx)
                            {
                                String webMsg = webEx.Message.ToString();
                            }
                            catch (Exception gEx)
                            {
                                String gMsg = gEx.Message.ToString();
                            }

                        }
                    }
                    // end trident
                    // start henley
                    if (aStaffSpec.LocID.ToString().ToUpper().Trim().IndexOf("HENLEY IN ARDEN") >= 0) henley = true;

                    googleHenley = service.Groups.IsMember(aStaffSpec.NDSName.ToString(), "STAFF-ARDN");
                    if (googleHenley && !henley)
                    {
                        try
                        {
                            service.Groups.RemoveMemberFromGroup(aStaffSpec.NDSName.ToString(), "STAFF-ARDN");
                        }
                        catch (AppsException uex)
                        {
                            String msg = uex.Message;
                            String sReason = uex.Reason.ToString();
                        }
                        catch (System.Net.WebException webEx)
                        {
                            String webMsg = webEx.Message.ToString();
                        }
                        catch (Exception gEx)
                        {
                            String gMsg = gEx.Message.ToString();
                        }

                    }
                    else
                    {
                        if (henley && !googleHenley)
                        {
                            try
                            {
                                service.Groups.AddMemberToGroup(aStaffSpec.NDSName.ToString(), "STAFF-ARDN");
                            }
                            catch (AppsException uex)
                            {
                                String msg = uex.Message;
                                String sReason = uex.Reason.ToString();
                            }
                            catch (System.Net.WebException webEx)
                            {
                                String webMsg = webEx.Message.ToString();
                            }
                            catch (Exception gEx)
                            {
                                String gMsg = gEx.Message.ToString();
                            }

                        }
                    }
                    // end henley

                }

                mWritten.Add(aStaffSpec.queueItem.ToString());

            }
            catch (AppsException uexi)
            {
                String msgi = uexi.Message;
                String sReasoni = uexi.Reason.ToString();
                if (uexi.ErrorCode == AppsException.UserDeletedRecently)
                {
                    //User deleted recently - wait at least 5 days
                }
            }
            catch (System.Net.WebException webEx)
            {
                String webMsg = webEx.Message.ToString();
            }
            catch (Exception gEx)
            {
                String gMsg = gEx.Message.ToString();
            }



            //updateSelectedItem(aStudentSpec, "T", "T", "Y");
        }



        protected override void testSelectedPhase(staffSpecSL aStaffSpecSL, DateTime runDateTime)
        {
            //AppsService service = new AppsService("stu.warwickshire.ac.uk", "gadmin@stu.warwickshire.ac.uk", "thebigpic7ure");
            // this is for the GROUP test phase
            foreach (staffSpec wSS in aStaffSpecSL.Values)
            {

                try
                {
                    UserEntry ue = service.RetrieveUser(wSS.NDSName.ToString());
                    //UserEntry ue = service.RetrieveUser("atestlan12347");
                    bool rls = false;
                    bool mm = false;
                    bool trident = false;
                    bool rugby = false;
                    bool henley = false;
                    bool pershore = false;
                    bool groupsDone = true;
                    //*jj
                    //DataView wCouplerDV = staffUtility.readDataView(staffUtility.couplerDB, "wcStudentSites '" + wSS.NDSName.ToString() + "'");
                    String locID = wSS.LocID.ToString();
                    //if (wCouplerDV.Count > 0)
                    {
                        bool googleBroadcast = service.Groups.IsMember(wSS.NDSName.ToString(), "STAFF-ALLSITES");
                        if (googleBroadcast == false) groupsDone = false;

                        bool googleRLS = service.Groups.IsMember(wSS.NDSName.ToString(), "STAFF-LSPA");
                        if (locID.ToString().ToUpper().Trim().IndexOf("LEAMINGTON SPA") >= 0) rls = true;
                        if (googleRLS != rls) groupsDone = false;

                        bool googlePershore = service.Groups.IsMember(wSS.NDSName.ToString(), "STAFF-PER");
                        if (locID.ToString().ToUpper().Trim().IndexOf("PERSHORE") >= 0) pershore = true;
                        if (googlePershore != pershore) groupsDone = false;

                        bool googlemoreton = service.Groups.IsMember(wSS.NDSName.ToString(), "STAFF-MM");
                        if (locID.ToString().ToUpper().Trim().IndexOf("MORETON MORRELL") >= 0) mm = true;
                        if (googlemoreton != mm) groupsDone = false;

                        bool googleTrident = service.Groups.IsMember(wSS.NDSName.ToString(), "STAFF-TRIDENT");
                        if (locID.ToString().ToUpper().Trim().IndexOf("TRIDENT PARK") >= 0) trident = true;
                        if (googleTrident != trident) groupsDone = false;

                        bool googleHenley = service.Groups.IsMember(wSS.NDSName.ToString(), "STAFF-ARDN");
                        if (locID.ToString().ToUpper().Trim().IndexOf("HENLEY IN ARDEN") >= 0) henley = true;
                        if (googleHenley != henley) groupsDone = false;

                        bool googleRugby = service.Groups.IsMember(wSS.NDSName.ToString(), "STAFF-RUG");
                        if (locID.ToString().ToUpper().Trim().IndexOf("RUGBY") >= 0) rugby = true;
                        if (googleRugby != rugby) groupsDone = false;
                    }
                    if (groupsDone) mTestedOK.Add(wSS.queueItem.ToString());

                    //Write a coupler job for moveObject here
                    //studentUtility.writeCouplerMessageQueueStu(wSS.NDSName, "", "GoogleGroupAdd");

                }
                catch (AppsException uex)
                {
                    String msg = uex.Message;
                    String sReason = uex.Reason.ToString();
                    if (wSS.attempts > mTryCount) mFailed.Add(wSS.queueItem.ToString());
                    if (uex.ErrorCode == AppsException.EntityDoesNotExist)
                    {
                        //User does not exist in Google Apps
                    }
                }
                catch (System.Net.WebException webEx)
                {
                    String webMsg = webEx.Message.ToString();
                }
                catch (Exception gEx)
                {
                    String gMsg = gEx.Message.ToString();
                }


                //if (testSelectedItem(wSS, "T", "T", "Y"))
                //    mTestedOK.Add(wSS.queueItem.ToString());
                //else
                //    if (wSS.attempts > mTryCount) mFailed.Add(wSS.queueItem.ToString());
            }
        }

    }
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
                UserEntry ue = service.RetrieveUser(aStaffSpec.NDSName.ToString());
            }
            catch (AppsException uex)
            {
                String msg = uex.Message;
                String sReason = uex.Reason.ToString();
                bool sitePershore = false;
                if (uex.ErrorCode == AppsException.EntityDoesNotExist)
                {
                    //User does not exist in Google Apps so create the user
                    try
                    {
                        // UserEntry ueCreate = service.CreateUser("atestLAN12347", "Steve", "Langley", "c0v3ntry");
                        
                        String password = staffUtility.createGooglePassword(aStaffSpec.NDSName);
                        UserEntry ueCreate = service.CreateUser(aStaffSpec.NDSName.ToString(), aStaffSpec.Forename.ToString(), aStaffSpec.Surname.ToString(), password);
                        mWritten.Add(aStaffSpec.queueItem.ToString());
                        String[] queueItem;
                        queueItem = new String[1];
                        queueItem[0] = aStaffSpec.queueItem.ToString();
                        int rowsChanged = staffUtility.updateCouplerMessageQueueSet(queueItem, " SET actionData='"+password+"'");
                    }
                    catch (AppsException uexi)
                    {
                        String msgi = uexi.Message;
                        String sReasoni = uexi.Reason.ToString();
                        if (uexi.ErrorCode == AppsException.UserDeletedRecently)
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

            foreach (staffSpec wSS in aStaffSpecSL.Values)
            {

                try
                {

                    
                    UserEntry ue = service.RetrieveUser(wSS.NDSName.ToString());
                    String password = staffUtility.createGooglePassword(wSS.NDSName);

                    //UserEntry ue = service.RetrieveUser("atestlan12347");
                    mTestedOK.Add(wSS.queueItem.ToString());
                    //Write a coupler job for moveObject here
                    staffUtility.writeCouplerMessageQueueV2(wSS.NDSName, "", "GoogleGroupAdd");

                }
                catch (AppsException uex)
                {
                    String msg = uex.Message;
                    String sReason = uex.Reason.ToString();
                    if (wSS.attempts > mTryCount) mFailed.Add(wSS.queueItem.ToString());
                    if (uex.ErrorCode == AppsException.EntityDoesNotExist)
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
                UserEntry ue = service.RetrieveUser(aStaffSpec.NDSName.ToString());
                service.RestoreUser(aStaffSpec.NDSName.ToString());
                mWritten.Add(aStaffSpec.queueItem.ToString());

            }
            catch (AppsException uex)
            {
                String msg = uex.Message;
                String sReason = uex.Reason.ToString();
                bool sitePershore = false;
                if (uex.ErrorCode == AppsException.EntityDoesNotExist)
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
                    UserEntry ue = service.RetrieveUser(wSS.NDSName.ToString());
                    // if fore name or surname need updating then do it here.

                    if ((String.Compare(ue.Name.FamilyName.ToLower().Trim().ToString(), wSS.Surname.ToLower().Trim().ToString()) != 0) || (String.Compare(ue.Name.GivenName.ToLower().Trim().ToString(), wSS.Forename.ToLower().Trim().ToString()) != 0))
                    {
                        ue.Name.FamilyName = wSS.Surname.Trim().ToString();
                        ue.Name.GivenName = wSS.Forename.Trim().ToString();

                        if (wSS.ActionData.Trim().Length < 1) { } //ue.Login.Password = password; int rowsChanged = staffUtility.updateCouplerMessageQueueSet(queueItem, " SET actionData='"+password+"'");
                        int passlength = ue.Login.Password.Length;
                        ue.Update();
                    }
                    //String gn2 = ue.Name.FamilyName.ToString();

                    if (ue.Login.Suspended == false)
                    {
                        //studentUtility.updateCouplerMessageQueue(wSS.queueItem.ToString(), "SET actionData = ''");
                        mTestedOK.Add(wSS.queueItem.ToString());

                    }
                    //UserEntry ue = service.RetrieveUser("atestlan12347");
                    //Write a coupler job for moveObject here
                    //studentUtility.writeCouplerMessageQueueStu(wSS.NDSName, "", "GoogleGroupAdd");

                }
                catch (AppsException uex)
                {
                    String msg = uex.Message;
                    String sReason = uex.Reason.ToString();
                    if (wSS.attempts > mTryCount) mFailed.Add(wSS.queueItem.ToString());
                    if (uex.ErrorCode == AppsException.EntityDoesNotExist)
                    {
                        //User does not exist in Google Apps
                        //mTestedOK.Add(wSS.queueItem.ToString());

                    }
                }

            }
        }

    }
    // ----------------------------------------------------------------------------------------------------------- 
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

                UserEntry ue = service.RetrieveUser(aStaffSpec.NDSName.ToString());
                service.SuspendUser(aStaffSpec.NDSName.ToString());
                mWritten.Add(aStaffSpec.queueItem.ToString());
            }
            catch (AppsException uex)
            {
                String msg = uex.Message;
                String sReason = uex.Reason.ToString();
                bool sitePershore = false;
                if (uex.ErrorCode == AppsException.EntityDoesNotExist)
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
                    UserEntry ue = service.RetrieveUser(wSS.NDSName.ToString());
                    if (ue.Login.Suspended == true)
                    {
                        //staffUtility.updateCouplerMessageQueue(wSS.queueItem.ToString(), "SET actionData = ''");
                        mTestedOK.Add(wSS.queueItem.ToString());

                    }
                    //UserEntry ue = service.RetrieveUser("atestlan12347");
                    //Write a coupler job for moveObject here
                    //studentUtility.writeCouplerMessageQueueStu(wSS.NDSName, "", "GoogleGroupAdd");

                }
                catch (AppsException uex)
                {
                    String msg = uex.Message;
                    String sReason = uex.Reason.ToString();
                    if (wSS.attempts > mTryCount) mFailed.Add(wSS.queueItem.ToString());
                    if (uex.ErrorCode == AppsException.EntityDoesNotExist)
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
                UserEntry ue = service.RetrieveUser(aStaffSpec.NDSName.ToString());
                service.DeleteUser(aStaffSpec.NDSName.ToString());
                mWritten.Add(aStaffSpec.queueItem.ToString());
            }
            catch (AppsException uex)
            {
                String msg = uex.Message;
                String sReason = uex.Reason.ToString();
                bool sitePershore = false;
                if (uex.ErrorCode == AppsException.EntityDoesNotExist)
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
                    UserEntry ue = service.RetrieveUser(wSS.NDSName.ToString());
                    //UserEntry ue = service.RetrieveUser("atestlan12347");
                    //Write a coupler job for moveObject here
                    //studentUtility.writeCouplerMessageQueueStu(wSS.NDSName, "", "GoogleGroupAdd");

                }
                catch (AppsException uex)
                {
                    String msg = uex.Message;
                    String sReason = uex.Reason.ToString();
                    if (wSS.attempts > mTryCount) mFailed.Add(wSS.queueItem.ToString());
                    if (uex.ErrorCode == AppsException.EntityDoesNotExist)
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
