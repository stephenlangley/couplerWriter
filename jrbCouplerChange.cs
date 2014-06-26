using System;
using System.IO;
using System.Collections.Generic;
using System.Text;
using System.Data;
using System.DirectoryServices;

namespace couplerWriter
{

    class jrbCouplerChangeDetails : jrbCoupler
    {
        public jrbCouplerChangeDetails
            (String aJRBFilesPath, String aActionName, int aTryCount)
            : base(aJRBFilesPath, aActionName, aTryCount) { }

        protected override String doSelectedItem(staffSpec aStaffSpec)
        {
            if (aStaffSpec.QLId != "")
            {
                String Description = "User details changed on " + DateTime.Now;
                String FullName = aStaffSpec.Forename + " " + aStaffSpec.Surname;
                String wNewStaffJRBText =
                    "\"" + aStaffSpec.NDSName + "\"" +
                    "," + "\"" + aStaffSpec.Surname.ToUpper() + "\"" +
                    "," + "\"" + aStaffSpec.Forename.ToUpper() + "\"" +
                    "," + "\"" + FullName.ToUpper() + "\"" +
                    "," + "\"" + aStaffSpec.JobTitle.ToUpper() + "\"" +
                    "," + "\"" + aStaffSpec.Tel + "\"" +
                     "," + "\"" + aStaffSpec.Department.ToUpper() + "\"" +
                    "," + "\"" + aStaffSpec.StaffID.ToString() + "\"" +
                   "," + "\"" + Description + "\"";

                return wNewStaffJRBText;
            }
            else return "";
        }

        protected override void testSelectedPhase(staffSpecSL aStaffSpecSL, DateTime runDateTime)
        {
            foreach (staffSpec ss in aStaffSpecSL.Values)
            {
                mAttributeValues = GetLDAPInfo("cn=" + ss.NDSName);
                // this problem is a synchronisation issue - nds may not match staffidentity for telephone number.
                //the staff maitenance form does not allow a blank telephone number so the staff identity telephone must be blank
                // because this class is shared with the Publishposts routine we need to ensure that the telephone number matches
                // with the NDS number so set them equal.
                if (ss.Tel.Length < 1) ss.Tel = attributeValue("telephonenumber").Trim().ToUpper().ToString();
                String y = attributeValue("telephonenumber").Trim().ToUpper().ToString();
                if (
                    (attributeValue("telephonenumber").Trim().ToUpper().ToString() == ss.Tel.Trim().ToUpper().ToString()) &&
                    (attributeValue("surname").ToUpper().Trim() == ss.Surname.ToUpper().Trim()) &&
                    (attributeValue("givenname").ToUpper().Trim() == ss.Forename.ToUpper().Trim()) &&
                    (attributeValue("title").ToUpper().Trim() == ss.JobTitle.ToUpper().Trim())
                )
                {
                    mTestedOK.Add(ss.queueItem.ToString());
                }
                else
                    if (ss.attempts > mTryCount) mFailed.Add(ss.queueItem.ToString());
            }
        }

    }

    class jrbCouplerChangeContact : jrbCoupler
    {
        public jrbCouplerChangeContact
            (String aJRBFilesPath, String aActionName, int aTryCount)
            : base(aJRBFilesPath, aActionName, aTryCount) { }

        protected override String doSelectedItem(staffSpec aStaffSpec)
        {
            if (aStaffSpec.QLId != "")
            {
                //String Description = "User contact details changed on " + DateTime.Now;
                String Description = "ChangeContactNDS=" + aStaffSpec.queueItem.ToString().Trim() + "@" + DateTime.Now;
                String FullName = aStaffSpec.Forename + " " + aStaffSpec.Surname;
                String wNewStaffJRBText =
                    "\"" + aStaffSpec.NDSName + "\"" +
                    "," + "\"" + aStaffSpec.Surname.ToUpper() + "\"" +
                    "," + "\"" + aStaffSpec.Forename.ToUpper() + "\"" +
                    "," + "\"" + FullName.ToUpper() + "\"" +
                    "," + "\"" + aStaffSpec.JobTitle.ToUpper() + "\"" +
                    "," + "\"" + aStaffSpec.Tel + "\"" +
                     "," + "\"" + aStaffSpec.Department.ToUpper() + "\"" +
                    "," + "\"" + aStaffSpec.StaffID.ToString() + "\"" +
                   "," + "\"" + Description + "\"";

                return wNewStaffJRBText;
            }
            else return "";
        }

        protected override void testSelectedPhase(staffSpecSL aStaffSpecSL, DateTime runDateTime)
        {
            foreach (staffSpec ss in aStaffSpecSL.Values)
            {
                mAttributeValues = GetLDAPInfo("cn=" + ss.NDSName);
                // TEST NDS for the queueItem, this shows us that the jrbimprt has taken place.
                // NOTE - jrbimprt has to overwrite the description field with the queueItem because it is a global setting within the control file
                // the reason is we don't want some of the multivalued fields to have multi values
                // so if more than one 
                // job has been sent for processing then only one job will test for complete and that will be the job
                // that was last processed by jrbimprt.
                // this is a problem because the system is 'disconnected' 
                String description = "ChangeContactNDS=" + ss.queueItem.ToString().Trim(); // set the test value for NDS
                String[] Alldescriptions = attributeArray("description");
                String isDone = "N";
                if (Alldescriptions != null )
                {
                    foreach (String NDSdescription in Alldescriptions)
                    {
                        String[] descriptions = NDSdescription.ToString().Trim().Split('@');

                        foreach (string NDSdesc in descriptions)
                        {
                            if (NDSdesc.ToString().Trim() == description)
                            {
                                isDone = "Y";
                            }
                        }
                    }
                }
                 if (isDone == "Y")
                 {// exec wcIDMupdateApplications - A call to upDateApplicationUserName with OLDNDS= current NDSname will update all relevant applications.
                        staffUtility.writeCouplerMessageQueueV2(ss.NDSName, "OldNDS=" + ss.NDSName, "updateApplicationUserName");

                        //Udate the 'OLD_' fields in the phone_list; NDS and Groupwise have been updated. 
                        String IDM = "wcIDMSynchPhoneListFields " + Convert.ToInt64(ss.QLId);
                        DataView wDV = staffUtility.readDataView(staffUtility.couplerDB, IDM);
                        mTestedOK.Add(ss.queueItem.ToString());
                    }
                    else
                        if (ss.attempts > mTryCount) mFailed.Add(ss.queueItem.ToString());
                
            }
        }

    }

    
    
    class jrbCouplerAddToGroup : jrbCoupler
    {
        public jrbCouplerAddToGroup
            (String aJRBFilesPath, String aActionName, int aTryCount)
            : base(aJRBFilesPath, aActionName, aTryCount) { }

        protected override String doSelectedItem(staffSpec aStaffSpec)
        {
            if (aStaffSpec.QLId != "")
            {
                String Description = "User details changed on " + DateTime.Now;
                String FullName = aStaffSpec.Forename + " " + aStaffSpec.Surname;
                String grpName = aStaffSpec.ActionData.Replace("Group=", "");
                String wNewStaffJRBText =
                    "\"" + aStaffSpec.NDSName + "\"" +
                   "," + "\"" + grpName.Trim() + "\"";

                return wNewStaffJRBText;
            }
            else return "";
        }

        protected String buildGroupContext(String grp)
        {
            string grpContext = "";
            String[] aItem = grp.Split(',');
            foreach (String aPart in aItem)
            {
                grpContext = grpContext + "." + aPart.Substring(aPart.IndexOf('=') + 1);
            }
            return (grpContext.ToString().ToUpper());

        }

        protected override void testSelectedPhase(staffSpecSL aStaffSpecSL, DateTime runDateTime)
        {
            foreach (staffSpec ss in aStaffSpecSL.Values)
            {
                string y = "";
                mAttributeValues = GetLDAPInfo("cn=" + ss.NDSName);
                foreach (String grp in attributeArray("grpmbr"))
                {
                    string x = buildGroupContext(grp);
                    if (buildGroupContext(grp) == ss.ActionData.Replace("Group=","").ToString().ToUpper())
                    { 
                        mTestedOK.Add(ss.queueItem.ToString());
                        break;
                    }
                }
                if (ss.attempts > mTryCount) mFailed.Add(ss.queueItem.ToString());                
            }
        }

    }

    class jrbCouplerRemoveFromGroup : jrbCoupler
    {
        public jrbCouplerRemoveFromGroup
            (String aJRBFilesPath, String aActionName, int aTryCount)
            : base(aJRBFilesPath, aActionName, aTryCount) { }

        protected override String doSelectedItem(staffSpec aStaffSpec)
        {
            if (aStaffSpec.QLId != "")
            {
                String Description = "User details changed on " + DateTime.Now;
                String FullName = aStaffSpec.Forename + " " + aStaffSpec.Surname;
                String grpName = aStaffSpec.ActionData.Replace("Group=", "");
                String wNewStaffJRBText =
                    "\"" + aStaffSpec.NDSName + "\"" +
                   "," + "\"" + grpName.Trim() + "\"";

                return wNewStaffJRBText;
            }
            else return "";
        }

        protected String buildGroupContext(String grp)
        {
            string grpContext = "";
            String[] aItem = grp.Split(',');
            foreach (String aPart in aItem)
            {
                grpContext = grpContext + "." + aPart.Substring(aPart.IndexOf('=') + 1);
            }
            return (grpContext.ToString().ToUpper());

        }

        protected override void testSelectedPhase(staffSpecSL aStaffSpecSL, DateTime runDateTime)
        {
            foreach (staffSpec ss in aStaffSpecSL.Values)
            {
                string y = "";
                mAttributeValues = GetLDAPInfo("cn=" + ss.NDSName);
                Boolean grpExists = false;
                foreach (String grp in attributeArray("grpmbr"))
                {
                    if (buildGroupContext(grp) == ss.ActionData.Replace("Group=", "").ToString().ToUpper())
                    {
                        grpExists = true;
                        break;
                    }
                }
                if (! grpExists) mTestedOK.Add(ss.queueItem.ToString());

                if (ss.attempts > mTryCount) mFailed.Add(ss.queueItem.ToString());
            }
        }

    }



    class jrbCouplerChangeSite : jrbCoupler
    {
        public jrbCouplerChangeSite
            (String aJRBFilesPath, String aActionName, int aTryCount)
            : base(aJRBFilesPath, aActionName, aTryCount) { }

        protected override String doSelectedItem(staffSpec aStaffSpec)
        {
            if (aStaffSpec.QLId != "")
            {
                /*
                 * Use site name, e.g. Rugby, Trident, henley in arden and check 
                 * coupler table for more recent un-processed ChangeSiteNDS messages
                 * for a move to any site. If there is one, ignore, as this is 
                 * probably a duplicate request.
                 */
                String wNewStaffJRBText = "Change site for " + aStaffSpec.NDSName;
                bool doChangeSite = true;
                DataView dv = staffUtility.couplerDV(
                    "WHERE " +
                    "(c.whenDone IS NOT NULL) AND " +
                    "(c.action='ChangeSiteNDS') AND " +
                    "(NDSName='" + aStaffSpec.NDSName + "') AND " +
                    "(c.queueItem=" + aStaffSpec.queueItem.ToString() + ")"
                    );
                if (dv != null) if (dv.Count > 0) doChangeSite = false;
                if (doChangeSite)
                {
                    String FullName = aStaffSpec.Forename + " " + aStaffSpec.Surname;
                    bool wB = staffUtility.sendEmail(
                        "itservices@warkscol.ac.uk;" + aStaffSpec.EmailAddress,
                        "idm@warkscol.ac.uk",
                        "Staff Site Change : " + aStaffSpec.NDSName,
                        FullName + "( " + aStaffSpec.NDSName + " ) has indicated a need " +
                        "to change site to " + aStaffSpec.ActionData.Replace("NewSite=", "") + ". " +
                        "Please log a sitehelpdesk job for the network team (infra_op) for this and notify the user when complete. " +
                        "These jobs are normally expected to complete within 2 working days.",
                         null
                    );
                    return wNewStaffJRBText + ".";
                }
                return wNewStaffJRBText + " ignored.";
            }
            else return "";
        }

        protected override void testSelectedPhase(staffSpecSL aStaffSpecSL, DateTime runDateTime)
        {
            foreach (staffSpec ss in aStaffSpecSL.Values)
            {
                bool testChangeSite = true;
                DataView dv = staffUtility.couplerDV(
                    "WHERE " +
                    "(c.whenDone IS NULL) AND " +
                    "(c.action='" + mActionName + "') AND " +
                    "(NDSName='" + ss.NDSName + "') AND " +
                    "(c.queueItem>" + ss.queueItem.ToString() + ")"
                    );
                if (dv != null) if (dv.Count > 0) testChangeSite = false;
                if (testChangeSite)
                {
                    mAttributeValues = GetLDAPInfo("cn=" + ss.NDSName);
                    String wADSPath = attributeValue("adspath");
                    String wTargetSite = ss.ActionData.Replace("NewSite=", "");
                    Boolean wDoneIt = false;

                    switch (wTargetSite.ToUpper())
                    {
                        case "LEAMINGTON SPA":
                            wDoneIt = (wADSPath.IndexOf("ou=LSPA,o=WC") > 0); break;
                        case "RUGBY":
                            wDoneIt = (wADSPath.IndexOf("ou=RUG,o=WC") > 0); break;
                        case "TRIDENT PARK":
                            wDoneIt = (wADSPath.IndexOf("ou=TRIDENT,o=WC") > 0); break;
                        case "MORETON MORRELL":
                            wDoneIt = (wADSPath.IndexOf("ou=MM,o=WC") > 0); break;
                        case "HENLEY IN ARDEN":
                            wDoneIt = (wADSPath.IndexOf("ou=ARDN,o=WC") > 0); break;
                        case "PERSHORE":
                            wDoneIt = (wADSPath.IndexOf("ou=PER,o=WC") > 0); break;
                        default:
                            break;
                    }
                    if (wDoneIt)
                    {
                        // site change has completed so update the wcStaffIdentity
                        // 
                        staffUtility.updateWcStaffIdentity(ss.NDSName, "SET loc_id = '" + wTargetSite.ToString().Trim() + "', location = '" + wTargetSite.ToString().Trim() + "'");
                        staffUtility.updatePhoneList(ss.EmpID, "SET site = '" + wTargetSite.ToString().Trim() + "'");

                        mTestedOK.Add(ss.queueItem.ToString());
                    }
                    else
                        if (ss.attempts > mTryCount) mFailed.Add(ss.queueItem.ToString());
                }
                else
                    mTestedOK.Add(ss.queueItem.ToString());
            }
        }

    }

    class jrbCouplerProxy : jrbCoupler
    {
        public jrbCouplerProxy
            (String aJRBFilesPath, String aActionName, int aTryCount)
            : base(aJRBFilesPath, aActionName, aTryCount) { }

        protected override String doSelectedItem(staffSpec aStaffSpec)
        {
            if (aStaffSpec.QLId != "")
            {
                staffUtility.updateWcStaffIdentity(aStaffSpec.NDSName, "SET keepProxy = 1, forDeletion = 0");

                //String wNewStaffJRBText =
                //    "\"" + aStaffSpec.NDSName + "\"" +
                //    "," + "\"" + Password + "\"" +
                //    "," + "\"" + Description + "\"";
                String wNewStaffJRBText = "keepProxy flag set to TRUE for " + aStaffSpec.NDSName.ToString();
                return wNewStaffJRBText;
            }
            else return "";
        }

        protected override void testSelectedPhase(staffSpecSL aStaffSpecSL, DateTime runDateTime)
        {
            foreach (staffSpec ss in aStaffSpecSL.Values)
            {
                mAttributeValues = GetLDAPInfo("cn=" + ss.NDSName);

                //DateTime gwexpire = DateTime.Parse(attributeValue("gwexpire"));
                DateTime cwDate = DateTime.Now;
                DateTime nullDate = new DateTime(1970, 1, 1, 0,0,0);// if groupwise date is equal to this then a date has not been set.
                DateTime gwexpire = nullDate;
                if (DateTime.TryParse(attributeValue("gwexpire"), out gwexpire))
                {
                    //gwexpire has a good value
                }
                else
                {
                    //set gwexpire again because the date returned by the tryparse was 01/01/0001 which means we have a problem with access manager.
                    gwexpire = nullDate;
                }
                int expiredCount = DateTime.Compare( cwDate,gwexpire);// >0 means that the expirydate has expired in groupwise.
                int expireDateSet = DateTime.Compare( nullDate,gwexpire); // 0 means that no expirydate has been set in groupwise
                if (expireDateSet != 0 && expiredCount > 0 ) // then the groupwise DateTime has expired - so remove the staffidentity keepProxy flag.
                //if (attributeValue("ngwmailboxexpirationtime").ToLower() == "true")
                {
                    staffUtility.updateWcStaffIdentity(ss.NDSName, "SET keepProxy = 0");

                    mTestedOK.Add(ss.queueItem.ToString());
                }
                else
                    if (ss.attempts > mTryCount) mFailed.Add(ss.queueItem.ToString());
            }
        }

    }



    class jrbCouplerChangeLogin : jrbCoupler
    {
        public jrbCouplerChangeLogin
            (String aJRBFilesPath, String aActionName, int aTryCount)
            : base(aJRBFilesPath, aActionName, aTryCount) { }

        protected override String doSelectedItem(staffSpec aStaffSpec)
        {
            if (aStaffSpec.QLId != "")
            {
                /* 
                  Currently a stub procedure

                  1. Send email to IT Services advising
                     of need to change login.
                  2. Create pending wcStaffIdentity record 
                     by cloning current one but set 
                     IsPending='Y' and NDSName to new value.
                */
                String FullName = aStaffSpec.Forename + " " + aStaffSpec.Surname;
                String NewNDSName = "";
                String[] AD = aStaffSpec.ActionData.Split('=');
                if(AD[0].ToString() == "NewNDS") 
                  NewNDSName = AD[1].ToString();

                bool wB = staffUtility.sendEmail(
                    "itservices@warkscol.ac.uk;mis@warkscol.ac.uk;" + aStaffSpec.EmailAddress,
                    "idm@warkscol.ac.uk",
                    "Staff Login Change : " + aStaffSpec.NDSName,
                    FullName + "( " + aStaffSpec.NDSName + " ) has indicated a need " +
                    "to change login to " + NewNDSName + ". " +
                    "Please log a sitehelpdesk job for the network team (infra_op) for this and notify the user when complete. " +
                    "These jobs are normally expected to complete within 4 working days.",
                     null
                );

                return
                  "Change login processed for " + aStaffSpec.NDSName;
            }
            else return "";
        }

        protected override void testSelectedPhase(staffSpecSL aStaffSpecSL, DateTime runDateTime)
        {
            foreach (staffSpec ss in aStaffSpecSL.Values)
            {
                Boolean wNewLoginOk=false;
                String NewNDSName = "";
                String[] AD = ss.ActionData.Split('=');
                if(AD[0].ToString() == "NewNDS") 
                  NewNDSName = AD[1].ToString();
                // chect to see if the old nds name has been deleted
                mAttributeValues=GetLDAPInfo("cn=" + ss.NDSName);
                if (mAttributeValues == null)
                {
                    // check to see if new nds name has been created
                    mAttributeValues=GetLDAPInfo("cn=" + NewNDSName);
                    if(mAttributeValues.Count>0) wNewLoginOk=true;
                }
                if (wNewLoginOk)
                {
                    // maybe put this in its own coupler class?l
                    // Call staff identity NEW NDS login stored procedure here
                    if (ss.QLId.ToString().Trim().ToLower() != "deleted")
                    {
                        String IDM = "wcIDMupdateNewNDSname " + Convert.ToInt64(ss.QLId);
                        DataView wDV = staffUtility.readDataView(staffUtility.couplerDB, IDM);

                        if (wDV.Count > 0)
                        {
                            if (wDV[0][0].ToString() == "1")
                            {
                                // Success
                                mTestedOK.Add(ss.queueItem.ToString());
                            }
                            else
                            {
                                if (ss.attempts > mTryCount) mFailed.Add(ss.queueItem.ToString());
                            }
                        }
                    }
                    // TEST the stored Procedure wcIDMupdateNewNDSname here
                    // then do the Application part in defaultCouplerUpdateUserName 
                    // maybe call the stored procedure wcIDMupdateNewNDSname to do this
                    /* Success processing 
                       1. Update NDSName on original
                          wcStaffIdentity record
                       2. Delete the IsPending='Y'    
                          wcStaffIdentity record
                       3. Add record to mTestedOk
                    */
                    
                }
                else
                    if (ss.attempts > mTryCount) mFailed.Add(ss.queueItem.ToString());
            }
        }

    }


}
