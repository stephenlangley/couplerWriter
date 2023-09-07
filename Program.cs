using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Data;
using System.IO;
using System.Data.SqlClient;
using System.Configuration;
using System.Security.Cryptography;
using System.Text;

namespace couplerWriter
{

    class staffUtility
    {
        // Connection strings in .net4 framework64
         // change csIDUniformDB back to csIDNovDB when creating the executable after DEVELOPMENT/TESTING
        static public String couplerDB = "csIDNovDB";
        //static public String couplerDB = "csIDUniformDB";
// =======================================================================     
        static public String XrayUtilDB = "csUtilsXrayDB";
        static public String NovUtilDB = "csUtilityNovDB";
// =======================================================================
        static String mLogFile = "c:\\temp\\staffcoupler.log";
        static String mNetName = "idmCoupler";

        static public bool putLog(String aCSName, String aSqlQuery)
        {
            DateTime dt = DateTime.Now;
            try
            {
                StreamWriter sw = new StreamWriter(mLogFile, true);
                sw.WriteLine(dt.ToString("yyyy-MM-dd HH:mm") + " " + mNetName);
                sw.WriteLine(aCSName);
                sw.WriteLine(aSqlQuery);
                sw.Close();
                return true;
            }
            catch (Exception ue)
            {
                String ueDesc = ue.Message;
                StreamWriter sw = new StreamWriter(mLogFile, true);
                sw.WriteLine(dt.ToString("yyyy-MM-dd HH:mm") + " " + mNetName);
                sw.WriteLine(aCSName);
                sw.WriteLine(aSqlQuery);
                sw.WriteLine(ueDesc);
                sw.Close();
                return false;
            }
        }

       static public string Base64UrlEncode(string input)
        {
            //var inputBytes = System.Text.Encoding.UTF8.GetBytes(input);
            // Special "url-safe" base64 encode.
            //encoded = encoded.replace(/\//g,'_').replace(/\+/g,'-').replace(/\=/g,'*');
            return input.Replace('+', '-').Replace('/', '_').Replace("=", "");
        }


        static public String calculateMD5Hash(string input)
        {
            // step 1, calculate MD5 hash from input
            MD5 md5 = System.Security.Cryptography.MD5.Create();
            byte[] inputBytes = System.Text.Encoding.ASCII.GetBytes(input);
            byte[] hash = md5.ComputeHash(inputBytes);

            // step 2, convert byte array to hex string
            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < hash.Length; i++)
            {
                sb.Append(hash[i].ToString("X2"));
            }
            return sb.ToString();
        }

        static public String createGooglePassword(String ndsName)
        {
            String password;
            String passNdsname = ndsName.ToString().Trim() + "CreateMePassword";
            if (ndsName.Length > 4) passNdsname = ndsName.ToString();

            password = calculateMD5Hash(DateTime.Now.TimeOfDay.TotalMilliseconds.ToString().Substring(1,4) + passNdsname.Substring(passNdsname.Length-4,4));
            if (password.Length > 11) 
            {
                return password.Substring(0,12);
            }else return password;  
        }

        static public DataView readDataView(String aCSName, String aSqlQuery)
        {
            if (putLog(aCSName, aSqlQuery))
            {
                ConnectionStringSettingsCollection connections =
                    ConfigurationManager.ConnectionStrings;
                DataTable dt = new DataTable();
                SqlConnection connUtility = new SqlConnection(
                    connections[aCSName].ConnectionString
                    );
                try
                {
                    SqlDataAdapter adp = new SqlDataAdapter();
                    SqlCommand cmd = new SqlCommand(aSqlQuery, connUtility);
                    adp.SelectCommand = cmd;
                    try { adp.Fill(dt); }
                    catch (Exception ue) { putLog(aCSName + ":dataView Error", ue.Message); };
                }
                finally { connUtility.Close(); }
                return dt.DefaultView;
            }
            else
                return null;
        }

        static public void updateWcStaffIdentity(String aNDSName, String aSetString)
        {
            DataView wDV = readDataView(staffUtility.couplerDB,
                "UPDATE wcStaffIdentity " + aSetString + " " +
                "WHERE NDSName='" + aNDSName + "'"
            );
        }

        static public void updatePhoneList(String aEmpID, String aSetString)
        {
            DataView wDV = readDataView(staffUtility.couplerDB,
                "UPDATE phone_list " + aSetString + " " +
                "WHERE emp_id=" + aEmpID 
            );
        }
        static public DataView couplerDV(String whereClause)
        {
            String cDV = "SELECT DISTINCT w.NDSName, w.EmployeeNumber, w.loc_id, w.emp_type, w.statusGroupwise, w.keepProxy, isnull('0','0') as VisitingLec, w.WCG_Ltd, " +
                "isnull(p.forenames,w.forename) as forename, isnull(p.surname,w.surname) as surname, isnull(p.faculty,w.department) as department, isnull(p.job,w.job_Title) as job_title, isnull(p.phone_no,w.telephone) as telephone, " +
                "c.networkname, c.actionData, c.action, c.queueItem, c.attempts, c.whenRead " +
                "FROM couplerMessageQueue c LEFT JOIN wcStaffIdentity w ON " +
                "  w.NDSname=c.networkName " + "LEFT JOIN phone_list_newPhone p ON " +
                " w.EmployeeNumber=p.emp_id " + whereClause;  
            // "LEFT JOIN wcTempHR t ON " + " w.EmployeeNumber=t.EmployeeNumber " + whereClause;

            return readDataView(staffUtility.couplerDB, cDV);
        }

        static public int updateCouplerMessageQueueSet(String[] aQueueItemSet, String aSetString)
        {
            DataView wDV;
            // do we need to send e-mail to Line manager here?
            foreach (String qi in aQueueItemSet)
            {
                wDV = readDataView(staffUtility.couplerDB,
                    "UPDATE couplerMessageQueue " + aSetString + " " +
                    "WHERE queueItem=" + qi
                );
            }
            return aQueueItemSet.GetLength(0);
        }
        static public bool writeCouplerMessageQueueV2(String ndsName, String actionData, String action)
        {
            DataView wDV;
            String sqlInsert;
            sqlInsert = "wcIDMCMQinsert '" + ndsName.Trim().ToString() + "','" + action + "', '" + actionData + "'";

            wDV = readDataView(staffUtility.couplerDB, sqlInsert);
            return false;
        }

        static public bool writeCouplerMessageQueue(String ndsName,String qlid)
        {
            DataView wDV;
                String sqlInsert;
                sqlInsert = "wcIDMCMQinsert '" + ndsName.Trim().ToString() + "', 'DeleteDefault', 'EmpNum=" + qlid.ToString().Trim() + "'";

                wDV = readDataView(staffUtility.couplerDB, sqlInsert);
            return false;
        }

        static public bool wcIDMReSetStaffIdentity(String ndsName)
        {
            DataView wDV;
            String sqlInsert;
            sqlInsert = "wcIDMReSetStaffIdentity '" + ndsName.Trim().ToString() + "'";

            wDV = readDataView(staffUtility.couplerDB, sqlInsert);
            return false;
        }

        static public bool sendEmail(String aTo, String aFrom, String aSubject, String aText, String aCC)
        {
            bool wResult=true;
            try
            {
                DataView updateDV = staffUtility.readDataView(
                    "csEmailDB",
                    "email..wcIDMpostEmail '" + aTo +"'" +
                    ", '" + aFrom +"'" +
                    ", '" + aCC + "'" +
                    ", '" + aCC + "'" +
                    ", '" + aSubject + "'" +
                    ", '" + aText + "'" 
                );
                if (updateDV == null) wResult = false;
                if (updateDV.Count < 1) wResult = false;
            }
            catch (Exception ue) { String wM = ue.Message; wResult = false; }
            return wResult;
        }

    }

    class Program
    {

        // Connection strings in .net2 framework64
        static String EHDDB = "csEstatesDB";
        static String SHDDB = "csSiteHelpDeskDB";
        // change this back when to XRAY for DEVELOPMENT
        static String QLRDB = "csWcPortalNovDB";
        //static String QLRDB = "csWcPortalXRAYDB";
        static String NDSTargetPath = "c:\\wcIdentities\\staff\\todo\\";
        
        //static String NDSTargetPathStu = "c:\\wcIdentities\\student\todo\\";

        static void Main(string[] args)
        {

            couplerWriter cw = new couplerWriter(EHDDB,SHDDB,QLRDB,NDSTargetPath);

            DataView wCouplerDV = staffUtility.couplerDV("WHERE (c.whenDone IS NULL) AND (attempts IS NOT NULL)");

            if (wCouplerDV != null) if (wCouplerDV.Count > 0)
                {
                    cw.testPhase(wCouplerDV);
                    int testedOKCount =  staffUtility.updateCouplerMessageQueueSet(cw.testedOK(), "SET whenDone=getdate()");
                if (cw.failed().Length > 0)    
                {

                    bool wb = 
                            staffUtility.sendEmail(
                                "mis@warkscol.ac.uk",
                                "idm@warkscol.ac.uk",
                                "Staff ID Coupler Fails : " + DateTime.Now.ToString("yyyyMMddHHmmss"),
                                "Coupler job " + String.Join(", ", cw.failed()),
                                 null
                            );
                }
                    int failedCount =
                        staffUtility.updateCouplerMessageQueueSet(cw.failed(), "SET attempts=0");
                }

            wCouplerDV =  staffUtility.readDataView(staffUtility.couplerDB,
                    "UPDATE couplerMessageQueue SET attempts=COALESCE(attempts+1,1) " +
                    "WHERE whenDone IS NULL"
                );

            // Do Phase
            wCouplerDV =
                staffUtility.couplerDV("WHERE (c.whenRead IS NULL) AND (whenDone IS null)");
            if (wCouplerDV != null) if (wCouplerDV.Count > 0)
                {
                    cw.doPhase(wCouplerDV);
                    int skippedCount =
                        staffUtility.updateCouplerMessageQueueSet(
                            cw.skipped(), 
                            "SET whenRead=getdate(), whenDone=getDate()"
                        );

                    int writtenCount =
                        staffUtility.updateCouplerMessageQueueSet(
                            cw.written(), 
                            "SET whenRead=getdate()"
                        );
                }

            
        }
    
    }
}
