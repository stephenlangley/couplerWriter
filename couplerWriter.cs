using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using System.IO;
using System.Data.SqlClient; 

namespace couplerWriter
{
    public class couplerWriter
    {
        List<coupler> mCouplerList;
        List<String> mWritten, mSkipped, mTestedOK, mFailed;

        public couplerWriter(String aEHDDB, String aSHDDB, String aQLRDB, String aNDSTargetPath)
        {
            mCouplerList = new List<coupler>();
            mWritten = new List<String>();
            mTestedOK = new List<String>(); 
            mFailed = new List<String>();
            mSkipped = new List<String>();
            mCouplerList.Add(new jrbCouplerCreate(aNDSTargetPath + "NewStaffLogins", "CreateNDS", 9));
            mCouplerList.Add(new jrbCouplerUpdate(aNDSTargetPath + "UpdateStaffLogins", "UpdateNDS", 9));
            mCouplerList.Add(new jrbCouplerChangeContact(aNDSTargetPath + "ChangeContactNDS", "ChangeContactNDS", 9));
            mCouplerList.Add(new jrbCouplerProxy(aNDSTargetPath + "ProxyStaff", "ProxyNDS", 9999));
            mCouplerList.Add(new jrbCouplerEnable(aNDSTargetPath + "EnableStaffLogins", "EnableNDS", 9));
            mCouplerList.Add(new jrbCouplerDisable(aNDSTargetPath + "DisableStaffLogins", "DisableNDS", 9));
            mCouplerList.Add(new jrbCouplerTrash(aNDSTargetPath + "TrashStaffLogins", "TrashNDS", 9));
            mCouplerList.Add(new jrbCouplerDelete(aNDSTargetPath + "DeleteStaffLogins", "DeleteNDS", 9));
            mCouplerList.Add(new jrbCouplerChangeSite(aNDSTargetPath + "ChangeSiteStaffLogins", "ChangeSiteNDS", 144));
            mCouplerList.Add(new jrbCouplerClearNDSGroups(aNDSTargetPath + "ClearNDSGroups", "ClearNDSGroups", 144));
            mCouplerList.Add(new jrbCouplerChangeLogin(aNDSTargetPath + "RenameLoginStaffLogins", "changeUserNetworkName", 288));
            mCouplerList.Add(new defaultCouplerCreate(aEHDDB, aSHDDB, aQLRDB, "CreateDefault", 9));
            mCouplerList.Add(new defaultCouplerEnable(aEHDDB, aSHDDB, aQLRDB, "EnableDefault", 9));
            mCouplerList.Add(new defaultCouplerDelete(aEHDDB, aSHDDB, aQLRDB, "DeleteDefault", 9));
            mCouplerList.Add(new defaultCouplerUpdateUserName(aEHDDB, aSHDDB, aQLRDB, "updateApplicationUserName", 9));

            // Google Stuff below here
            mCouplerList.Add(new defaultCouplerCreateGoogle(aEHDDB, aSHDDB, aQLRDB, "CreateGoogleLogins", 9));
            mCouplerList.Add(new defaultCouplerCreateGoogleOU(aEHDDB, aSHDDB, aQLRDB, "CreateGoogleOU", 9));
            mCouplerList.Add(new defaultCouplerCreateGoogleGroup(aEHDDB, aSHDDB, aQLRDB, "GoogleGroupAdd", 9));
            mCouplerList.Add(new defaultCouplerSuspendGoogle(aEHDDB, aSHDDB, aQLRDB, "SuspendGoogle", 9));
            mCouplerList.Add(new defaultCouplerRestoreGoogle(aEHDDB, aSHDDB, aQLRDB, "RestoreGoogle", 9));
            mCouplerList.Add(new defaultCouplerDeleteGoogle(aEHDDB, aSHDDB, aQLRDB, "DeleteGoogle", 9));
//=================================================================================================================================
            //// put on hold mCouplerList.Add(new jrbCouplerHomeDir(aNDSTargetPath + "HomeDirStaffLogins", "HomeDirNDS", 9));
            //// put on hold mCouplerList.Add(new jrbCouplerReSetHomeDir(aNDSTargetPath + "ReSetHomeDirStaffLogins", "ReSetHomeDirNDS", 9));
            ////mCouplerList.Add(new jrbCouplerChangeDetails(aNDSTargetPath + "ChangeDetailsStaffLogins", "ChangeDetailsNDS", 9));
            ////not yet implemented
            ////mCouplerList.Add(new jrbCouplerAddToGroup(aNDSTargetPath + "GrpAddStaff", "AddToGroupNDS", 144));
            ////mCouplerList.Add(new jrbCouplerRemoveFromGroup(aNDSTargetPath + "GrpRemoveStaff", "RemoveFromGroupNDS", 144));
            //// not implemented mCouplerList.Add(new defaultCouplerDisable(aEHDDB, aSHDDB, aQLRDB, "DisableDefault", 9));
            //// not implemented mCouplerList.Add(new defaultCouplerTrash(aEHDDB, aSHDDB, aQLRDB, "TrashDefault", 9));
            ////mCouplerList.Add(new defaultCouplerChangeDetails(aEHDDB, aSHDDB, aQLRDB, "ChangeDetailsDefault", 9));
            //// not implemented mCouplerList.Add(new jrbCouplerChangeDetails(aNDSTargetPath + "PassExpChangeDetailsStaffLogins", "PassExpChangeDetailsNDS", 9));

            foreach (coupler c in mCouplerList) c.setResultLists(mWritten, mSkipped, mTestedOK, mFailed);
        }

        public String[] written() { return mWritten.ToArray(); }
        public String[] skipped() { return mSkipped.ToArray(); }
        public String[] testedOK() { return mTestedOK.ToArray(); }
        public String[] failed() { return mFailed.ToArray(); }

        public void doPhase(DataView aCouplerDV)
        {   foreach (coupler c in mCouplerList) c.doPhase(aCouplerDV, DateTime.Now);  }

        public void testPhase(DataView aCouplerDV)
        {   foreach (coupler c in mCouplerList) c.testPhase(aCouplerDV, DateTime.Now);  }

    }
}
