using System;
using System.Collections.Generic;
using System.Data;
using System.Text;

namespace couplerWriter
{
    class coupler
    {
        protected String mActionName = "";
        protected int mTryCount = 0;
        protected List<String> mWritten, mSkipped, mTestedOK, mFailed;

        public coupler() { }

        public void setResultLists(
            List<String> aWritten, List<String> aSkipped, 
            List<String> aTestedOK, List<String> aFailed)
        {
            mWritten = aWritten; mSkipped = aSkipped;
            mTestedOK = aTestedOK; mFailed = aFailed; 
        }

        protected virtual void doSelectedPhase(staffSpecSL aStaffSpecSL, DateTime runDateTime) { }

        protected virtual void testSelectedPhase(staffSpecSL aStaffSpecSL, DateTime runDateTime) { }

        public void doPhase(DataView aCouplerDV, DateTime runDateTime)
        {
            aCouplerDV.RowFilter = "action='" + mActionName + "'";
            aCouplerDV.Sort = "NDSName, queueItem desc";

            if (aCouplerDV.Count > 0)
            {
                staffSpecSL wSSL = new staffSpecSL(aCouplerDV, mSkipped);
                doSelectedPhase(wSSL, runDateTime);
                foreach (staffSpec aSS in wSSL.Values) 
                { 
                    mWritten.Add(aSS.queueItem.ToString()); 
                }
            }
        }

        public void testPhase(DataView aCouplerDV, DateTime runDateTime)
        {
            aCouplerDV.RowFilter = "action='" + mActionName + "'";
            if (aCouplerDV.Count > 0) testSelectedPhase(new staffSpecSL(aCouplerDV,mSkipped), runDateTime);
        }

    }
}
