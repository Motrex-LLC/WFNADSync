using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.SqlClient;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;
using System.Collections;

namespace WFNADSync
{
    class ProcessJob
    {        
        private Hashtable hashADEmailToDirectoryEntry = new Hashtable();
        private Hashtable hashADEmpNoToDirectoryEntry = new Hashtable();

        private Hashtable hashADEmailToUserPrincipal = new Hashtable();
        private Hashtable hashADEmpNoToUserPrincipal = new Hashtable();

        public ProcessJob()
        {
            BuildADHashTable();
        }

        private void BuildADHashTable()
        {
            SqlCommand cmd = new SqlCommand("", AppDat.getWFNDBConn());
            string strSQL = "";
            int counter = 0;


            //DirectoryEntry de = new DirectoryEntry("LDAP://DC=Inside,DC=Exide,DC=ad", "inside\\weblogic_na", "webbea21", AuthenticationTypes.None);

            /*
            DirectoryEntry de = new DirectoryEntry();
            de.Path = "LDAP://10.138.212.222:389/DC=Inside,DC=Exide,DC=ad";
            de.Username = @"Inside\weblogic_na";
            de.Password = "webbea21";
                
            DirectorySearcher search = new DirectorySearcher(strFilter);                
            SearchResult adFind = search.FindOne();

            if (adFind == null) { return; }
                
            UserPrincipal u = new UserPrincipal(pContext);
            u.EmailAddress = "wendong.xu@exide.com";
            PrincipalSearcher pSearcher = new PrincipalSearcher(u);
            UserPrincipal uPrincipal = (UserPrincipal)pSearcher.FindOne();
            */



            if (Utility.getDumpADInfoToDB() == "Y")
            {
                cmd.CommandText = " Truncate Table wfnAdAttribute ";
                cmd.ExecuteNonQuery();
            }


            string TxDT = Utility.GetSessionIDFromSQL();

            PrincipalContext pContext = new PrincipalContext(ContextType.Domain, "10.138.212.222:389", "OU=US,OU=NA,DC=Inside,DC=Exide,DC=ad", "weblogic_na", "webbea21");
            //PrincipalContext pContext = new PrincipalContext(ContextType.Domain, "10.138.212.222:389", "DC=Inside,DC=Exide,DC=ad", "weblogic_na", "webbea21");

            UserPrincipal up = new UserPrincipal(pContext);
            PrincipalSearcher pSearcher = new PrincipalSearcher(up);
            List<UserPrincipal> allPrincipal = pSearcher.FindAll().Select(a => (UserPrincipal)a).ToList();

            foreach (UserPrincipal onePrl in allPrincipal)
            {
                try
                {
                    counter = counter + 1;
                    DirectoryEntry deWorking = (DirectoryEntry)onePrl.GetUnderlyingObject();
                    string extAttr2Email = deWorking.Properties["extensionAttribute2"]?.Value?.ToString()?.Trim()?.ToUpper();
                    string employeeNo = deWorking.Properties["employeeNumber"]?.Value?.ToString()?.Trim()?.ToUpper();
                    string sAMAccountName = deWorking.Properties["sAMAccountName"]?.Value?.ToString()?.Trim();
                    string displayName = deWorking.Properties["displayName"]?.Value?.ToString()?.Trim();
                    string title = deWorking.Properties["title"]?.Value?.ToString()?.Trim();
                    string mail = deWorking.Properties["mail"]?.Value?.ToString()?.Trim();
                    string userAccountControl = deWorking.Properties["userAccountControl"]?.Value?.ToString()?.Trim();

                    if (Utility.getDumpADInfoToDB() == "Y")
                    {
                        strSQL = " Insert Into wfnAdAttribute(sAMAccountName, employeeNumber, extensionAttribute2, displayName, title, mail, userAccountControl,TxDT) values('";
                        strSQL = strSQL + sAMAccountName + "','" + employeeNo + "','" + extAttr2Email + "','" + displayName?.Replace("'", "''") + "','" + title?.Replace("'", "''") + "','" + mail + "','" + userAccountControl + "','" + TxDT + "') ";
                        cmd.CommandText = strSQL;
                        cmd.ExecuteNonQuery();
                    }

                    if (String.IsNullOrEmpty(extAttr2Email) && String.IsNullOrEmpty(employeeNo)) { continue; }

                    //Utility.AppLog3Pass(employeeNo, onePrl.DisplayName, "", "EmpNo1_" + DateTime.Now.ToString("yyyyMMdd") + ".csv");
                    if (!String.IsNullOrEmpty(extAttr2Email) && !hashADEmailToDirectoryEntry.ContainsKey(extAttr2Email))
                    {
                        hashADEmailToDirectoryEntry.Add(extAttr2Email, deWorking);
                        hashADEmailToUserPrincipal.Add(extAttr2Email, onePrl);
                        //hashADEmailToDirectoryEntry.Add(extAttr2Email, onePrl);                        
                        //Utility.AppLog3Pass(extAttr2Email, onePrl.DisplayName, "", "EmailList_" + DateTime.Now.ToString("yyyyMMdd") + ".csv");
                    }

                    if (!String.IsNullOrEmpty(employeeNo) && !hashADEmpNoToDirectoryEntry.ContainsKey(employeeNo))
                    {
                        hashADEmpNoToDirectoryEntry.Add(employeeNo, deWorking);
                        hashADEmpNoToUserPrincipal.Add(employeeNo, onePrl);
                        //hashADEmpNoToDirectoryEntry.Add(employeeNo, onePrl);
                        //Utility.AppLog3Pass(employeeNo, onePrl.DisplayName, "", "EmpNot2_" + DateTime.Now.ToString("yyyyMMdd") + ".csv");
                    }

                }
                catch (Exception ex)
                {

                }
            }

            ////u.SamAccountName = "bautzda";
            //u.EmployeeId = strMeta4ID;
            //if (adFind.Properties["extensionAttribute2"].Count <= 0) {    return; } 
            //u.EmailAddress = adFind.Properties["mail"][0].ToString();

            //u.EmailAddress = "wendong.xu@exide.com";                
            //PrincipalSearcher pSearcher = new PrincipalSearcher(u);
            //UserPrincipal uPrincipal = (UserPrincipal)pSearcher.FindOne();
        }
        public void addNewEmpADAccount()
        {
            //https://docs.microsoft.com/en-us/previous-versions/bb384369(v=vs.90)?redirectedfrom=MSDN

            try
            {
                PrincipalContext pContext = new PrincipalContext(ContextType.Domain, "10.138.212.222:389", "OU=US,OU=NA,DC=Inside,DC=Exide,DC=ad", "weblogic_na", "webbea21");
                UserPrincipal up = new UserPrincipal(pContext);
                up.Surname = "TestAcct_Surname";
                up.GivenName = "TestAcct_GiveName";
                up.EmployeeId = "TestAcct_Surname";
                up.Enabled = true;
                up.ExpirePasswordNow();

                up.Save();
            }
            catch (Exception e)
            {

            }
        }

        //get employees on Web-API tables directly
        public void syncWFNExistingEmpToAD_ByEmployeeTable()
        {            
            SqlCommand cmd = new SqlCommand("", AppDat.getWFNDBConn());
            SqlCommand cmd2 = new SqlCommand("", AppDat.getWFNDBConn());
            SqlDataReader rdr = null;
            string strSQL = "";

            //strSQL = " Select Distinct homeOrgBU From wfnWorkAssignment ";
            //cmd.CommandText = strSQL;
            //rdr = cmd.ExecuteReader();
            //while (rdr.Read())
            //{
            //    string strLegalEntity= rdr["homeOrgBU"] ?.ToString();
            //    if (strLegalEntity == null || strLegalEntity == "") { continue; }

            //    processLegalEntityUsers(strLegalEntity);
            //}

            //BuildADHashTable();

            try
            {                
                strSQL = " Select B.associateOID, B.workerID As employeeNumber, B.givenName, B.middleName, B.familyName1, B.businessEmail, A.jobCode + ' - ' + A.jobTitle As positionCode, A.jobTitle, A.homeOrgBU + ' - ' + A.homeOrgBUName As legalEntity, A.homeOrgBUName As company, A.workerGroupCode + ' - ' + A.workerGroup As SBU, A.homeWorkLocCode As workLocationCode, A.homeOrgCost + ' - ' + A.homeOrgCostDept As USCostCenter, reportTo_workerID ";
                strSQL = strSQL + " From wfnWorkAssignment A Left Join wfnWorker B On A.associateOID=B.associateOID ";
                //strSQL = strSQL + " Where A.primaryIndicator = 'True' And B.workerStatus = 'Active'  ";
                strSQL = strSQL + " Where A.primaryIndicator = 'True' And B.workerStatus <> 'Terminated' And B.workerID='VINIMH2YE' ";

                cmd.CommandText = strSQL;
                rdr = cmd.ExecuteReader();
                while (rdr.Read())
                {
                    try
                    {
                        //string associateOID = rdr["associateOID"]?.ToString()?.Trim();
                        string employeeNumber = rdr["employeeNumber"]?.ToString()?.Trim()?.ToUpper();
                        string givenName = rdr["givenName"]?.ToString()?.Trim();
                        string middleName = rdr["middleName"]?.ToString()?.Trim();
                        string familyName1 = rdr["familyName1"]?.ToString()?.Trim();
                        string businessEmail = rdr["businessEmail"]?.ToString()?.Trim()?.ToUpper();
                        string positionCode = rdr["positionCode"]?.ToString()?.Trim().ToUpper();
                        string jobTitle = rdr["jobTitle"]?.ToString()?.Trim().ToUpper();
                        string legalEntity = rdr["legalEntity"]?.ToString()?.Trim().ToUpper();
                        string company = rdr["company"]?.ToString()?.Trim().ToUpper();
                        string SBU = rdr["SBU"]?.ToString()?.Trim();
                        string workLocationCode = rdr["workLocationCode"]?.ToString()?.Trim();
                        string USCostCenter = rdr["USCostCenter"]?.ToString()?.Trim();
                        string reportTo_workerID = rdr["reportTo_workerID"]?.ToString()?.Trim();

                        DirectoryEntry deWorking = (DirectoryEntry)hashADEmailToDirectoryEntry[businessEmail];
                        if (deWorking == null)
                        {
                            deWorking = (DirectoryEntry)hashADEmpNoToDirectoryEntry[employeeNumber];

                            if (deWorking == null) { continue; }
                            else
                            {
                                if (Utility.getDumpADInfoToDB() == "Y")
                                {
                                    strSQL = " Update wfnAdAttribute Set SyncedBy = 'EmpNo' Where employeeNumber='" + employeeNumber + "' ";
                                    cmd2.CommandText = strSQL;
                                    cmd2.ExecuteNonQuery();
                                }
                            }
                        }
                        else
                        {
                            if (Utility.getDumpADInfoToDB() == "Y")
                            {
                                strSQL = " Update wfnAdAttribute Set SyncedBy = 'Email' Where extensionAttribute2='" + businessEmail + "'  ";
                                cmd2.CommandText = strSQL;
                                cmd2.ExecuteNonQuery();
                            }
                        }


                        //employeeNumber,Title, JobTitleCode, LegalEntity, PositionCode, Company, SBU,  WorkLocationCode,USCostCenter,GlobalCostCenter,Manager

                        //if (!String.IsNullOrEmpty(givenName)){
                        //    deWorking.Properties["givenName"].Value = givenName;
                        //}

                        //if (!String.IsNullOrEmpty(familyName1))
                        //{
                        //    deWorking.Properties["sn"].Value = familyName1;
                        //}

                        //if (!String.IsNullOrEmpty(jobTitleCode))
                        //{
                        //    deWorking.Properties["JobTitleCode"].Value = jobTitleCode;
                        //}

                        if (!String.IsNullOrEmpty(employeeNumber))
                        {
                            deWorking.Properties["employeeNumber"].Value = employeeNumber;
                        }

                        if (!String.IsNullOrEmpty(jobTitle))
                        {
                            deWorking.Properties["title"].Value = jobTitle;                            
                        }
                        
                        if (!String.IsNullOrEmpty(legalEntity))
                        {
                            deWorking.Properties["legalEntity"].Value = legalEntity;
                        }

                        if (!String.IsNullOrEmpty(positionCode))
                        {
                            deWorking.Properties["positionCode"].Value = positionCode;
                        }

                        if (!String.IsNullOrEmpty(company))
                        {
                            deWorking.Properties["company"].Value = company;
                        }

                        if (!String.IsNullOrEmpty(SBU))
                        {
                            deWorking.Properties["SBU"].Value = SBU;
                        }

                        if (!String.IsNullOrEmpty(workLocationCode))
                        {
                            deWorking.Properties["workLocationCode"].Value = workLocationCode;
                        }

                        if (!String.IsNullOrEmpty(USCostCenter))
                        {
                            deWorking.Properties["USCostCenter"].Value = USCostCenter;
                        }

                        DirectoryEntry deWorking2 = (DirectoryEntry)hashADEmpNoToDirectoryEntry[reportTo_workerID];              
                        string managerDistinguishedName = deWorking2.Properties["DistinguishedName"]?.Value?.ToString()?.Trim();
                        if (!String.IsNullOrEmpty(managerDistinguishedName))
                        {
                            deWorking.Properties["Manager"].Value = managerDistinguishedName;
                        }

                        deWorking.CommitChanges();


                    }

                    catch (Exception e)
                    {
                    }
                }
            }

            catch (Exception e)
            {
            }

        }

        //get employees change from tabel WFNEmployeeADChanges refreshed by Renae
        public void syncWFNExistingEmpToAD_ByChangeTable()
        {
            SqlCommand cmd = new SqlCommand("", AppDat.getWFNDBConn());
            SqlDataReader rdr = null;
            string strSQL = "";

            //strSQL = " Select Distinct homeOrgBU From wfnWorkAssignment ";
            //cmd.CommandText = strSQL;
            //rdr = cmd.ExecuteReader();
            //while (rdr.Read())
            //{
            //    string strLegalEntity= rdr["homeOrgBU"] ?.ToString();
            //    if (strLegalEntity == null || strLegalEntity == "") { continue; }

            //    processLegalEntityUsers(strLegalEntity);
            //}

            //BuildADHashTable();

            try
            {

                strSQL = " Select InsertDate, WorkerID As employeeNumber, givenName, middleName, familyName, businessEmail, jobCode +' - ' + jobTitle As positionCode, jobTitle, homeOrgBU +' - ' + homeOrgBUName As legalEntity, homeOrgBUName As company, workerGroupCode +' - ' + workerGroup As SBU, homeWorkLocCode As workLocationCode, homeOrgCost +' - ' + homeOrgCostDept As USCostCenter, reportToWorkerID As reportTo_workerID ";
                strSQL = strSQL + " From WFNEmployeeADChanges ";
                strSQL = strSQL + " Where WorkerStatus = 'Active' And InsertDate > DateAdd(day, -5, Convert(Date, GetDate())) ";
                strSQL = strSQL + " Order By InsertDate ";

                //strSQL = " Select B.associateOID, B.workerID As employeeNumber, B.givenName, B.middleName, B.familyName1, B.businessEmail, A.jobCode + ' - ' + A.jobTitle As positionCode, A.jobTitle, A.homeOrgBU + ' - ' + A.homeOrgBUName As legalEntity, A.homeOrgBUName As company, A.workerGroupCode + ' - ' + A.workerGroup As SBU, A.homeWorkLocCode As workLocationCode, A.homeOrgCost + ' - ' + A.homeOrgCostDept As USCostCenter, reportTo_workerID ";
                //strSQL = strSQL + " From wfnWorkAssignment A Left Join wfnWorker B On A.associateOID=B.associateOID ";
                //strSQL = strSQL + " Where A.primaryIndicator = 'True' And workerID In ('M25851','M01801969') ";

                cmd.CommandText = strSQL;
                rdr = cmd.ExecuteReader();
                while (rdr.Read())
                {
                    try
                    {
                        //string associateOID = rdr["associateOID"]?.ToString()?.Trim();
                        string employeeNumber = rdr["employeeNumber"]?.ToString()?.Trim()?.ToUpper();
                        string givenName = rdr["givenName"]?.ToString()?.Trim();
                        string middleName = rdr["middleName"]?.ToString()?.Trim();
                        string familyName1 = rdr["familyName1"]?.ToString()?.Trim();
                        string businessEmail = rdr["businessEmail"]?.ToString()?.Trim()?.ToUpper();
                        string positionCode = rdr["positionCode"]?.ToString()?.Trim().ToUpper();
                        string jobTitle = rdr["jobTitle"]?.ToString()?.Trim().ToUpper();
                        string legalEntity = rdr["legalEntity"]?.ToString()?.Trim().ToUpper();
                        string company = rdr["company"]?.ToString()?.Trim().ToUpper();
                        string SBU = rdr["SBU"]?.ToString()?.Trim();
                        string workLocationCode = rdr["workLocationCode"]?.ToString()?.Trim();
                        string USCostCenter = rdr["USCostCenter"]?.ToString()?.Trim();
                        string reportTo_workerID = rdr["reportTo_workerID"]?.ToString()?.Trim();

                        DirectoryEntry deWorking = (DirectoryEntry)hashADEmailToDirectoryEntry[businessEmail];
                        if (deWorking == null) { deWorking = (DirectoryEntry)hashADEmpNoToDirectoryEntry[employeeNumber]; }
                        if (deWorking == null) { continue; }

                        if (!String.IsNullOrEmpty(employeeNumber))
                        {
                            deWorking.Properties["employeeNumber"].Value = employeeNumber;
                        }

                        if (!String.IsNullOrEmpty(jobTitle))
                        {
                            deWorking.Properties["title"].Value = jobTitle;
                        }

                        if (!String.IsNullOrEmpty(legalEntity))
                        {
                            deWorking.Properties["legalEntity"].Value = legalEntity;
                        }

                        if (!String.IsNullOrEmpty(positionCode))
                        {
                            deWorking.Properties["positionCode"].Value = positionCode;
                        }

                        if (!String.IsNullOrEmpty(company))
                        {
                            deWorking.Properties["company"].Value = company;
                        }

                        if (!String.IsNullOrEmpty(SBU))
                        {
                            deWorking.Properties["SBU"].Value = SBU;
                        }

                        if (!String.IsNullOrEmpty(workLocationCode))
                        {
                            deWorking.Properties["workLocationCode"].Value = workLocationCode;
                        }

                        if (!String.IsNullOrEmpty(USCostCenter))
                        {
                            deWorking.Properties["USCostCenter"].Value = USCostCenter;
                        }

                        DirectoryEntry deWorking2 = (DirectoryEntry)hashADEmpNoToDirectoryEntry[reportTo_workerID];
                        string managerDistinguishedName = deWorking2.Properties["DistinguishedName"]?.Value?.ToString()?.Trim();
                        if (!String.IsNullOrEmpty(managerDistinguishedName))
                        {
                            deWorking.Properties["Manager"].Value = managerDistinguishedName;
                        }

                        deWorking.CommitChanges();
                    }

                    catch (Exception e)
                    {
                    }
                }
            }

            catch (Exception e)
            {
            }

        }
        public void disableTerminatedEmpADAccount()
        {
            SqlCommand cmd = new SqlCommand("", AppDat.getWFNDBConn());
            SqlDataReader rdr = null;
            string strSQL = "";

            //strSQL = " Select Distinct homeOrgBU From wfnWorkAssignment ";
            //cmd.CommandText = strSQL;
            //rdr = cmd.ExecuteReader();
            //while (rdr.Read())
            //{
            //    string strLegalEntity= rdr["homeOrgBU"] ?.ToString();
            //    if (strLegalEntity == null || strLegalEntity == "") { continue; }

            //    processLegalEntityUsers(strLegalEntity);
            //}

            //BuildADHashTable();

            try
            {

                strSQL = " Select B.*, A.terminationDate ";
                strSQL = strSQL + " From wfnWorkAssignment A Left Join wfnWorker B On A.associateOID = B.associateOID ";                
                strSQL = strSQL + " Where primaryIndicator = 'True' And B.workerStatus = 'Terminated' And A.terminationDate > DateAdd(day, -30, Convert(Date, GETDATE())) And A.terminationDate < Convert(Date, GETDATE()) ";
                //strSQL = strSQL + " Where primaryIndicator = 'true' And B.workerStatus = 'Terminated' And A.terminationDate > DateAdd(day, -300, Convert(Date, GETDATE())) And workerID='M17070' ";
                strSQL = strSQL + " Order by A.terminationDate Desc ";

                cmd.CommandText = strSQL;
                rdr = cmd.ExecuteReader();
                while (rdr.Read())
                {
                    try
                    {
                        string associateOID = rdr["associateOID"]?.ToString()?.Trim();
                        string employeeNumber = rdr["workerID"]?.ToString()?.Trim();
                        string businessEmail = rdr["businessEmail"]?.ToString()?.Trim();
                        string terminationDate = rdr["terminationDate"]?.ToString()?.Trim();

                        ADTerminate(employeeNumber, businessEmail);
                    }

                    catch (Exception e)
                    {
                    }
                }
            }

            catch (Exception e)
            {
            }

        }

        private void ADTerminate(string strEmpID, string strEmail)
        {

            string strFilter = "";            

            if (!String.IsNullOrEmpty(strEmpID) && !String.IsNullOrEmpty(strEmail))
            {
                 strFilter = "(|(employeeNumber=" + strEmpID + ")(extensionAttribute2=" + strEmail + "))";
            }
            else if(!String.IsNullOrEmpty(strEmpID) && String.IsNullOrEmpty(strEmail))
            {
                 strFilter = "(employeeNumber=" + strEmpID + ")";
            }
            else if (String.IsNullOrEmpty(strEmpID) && !String.IsNullOrEmpty(strEmail))
            {
                 strFilter = "(extensionAttribute2=" + strEmail + ")";
            }
            else
            {
                return;
            }

            try
            {

                //DirectoryEntry de = new DirectoryEntry("LDAP://DC=Inside,DC=Exide,DC=ad", "inside\\weblogic_na", "webbea21", AuthenticationTypes.None);
                DirectoryEntry de = new DirectoryEntry();
                de.Path = "LDAP://10.138.212.222:389/DC=Inside,DC=Exide,DC=ad";
                de.Username = @"Inside\weblogic_na";
                de.Password = "webbea21";

                DirectorySearcher search = new DirectorySearcher(strFilter);
                SearchResult srEmpSingle = search.FindOne();
                
                if (srEmpSingle == null)    {   return; }

                if (srEmpSingle.Properties["SamAccountName"].Count > 0)
                {
                    //PrincipalContext pContext = new PrincipalContext(ContextType.Domain, "inside", "weblogic_na", "webbea21");
                    PrincipalContext pContext = new PrincipalContext(ContextType.Domain, "10.138.212.222:389", "DC=Inside,DC=Exide,DC=ad", "weblogic_na", "webbea21");                    
                    UserPrincipal up = new UserPrincipal(pContext);
                    //up.EmailAddress = srEmpSingle.Properties["mail"][0].ToString();
                    up.SamAccountName = srEmpSingle.Properties["SamAccountName"][0].ToString();

                    PrincipalSearcher psSearch = new PrincipalSearcher(up);
                    UserPrincipal result = (UserPrincipal)psSearch.FindOne();

                    if (result != null)
                    {
                        if ((bool)result.Enabled)
                        {
                            result.SamAccountName = "zz" + result.SamAccountName.ToString();
                            result.Description = "zz" + result.SamAccountName.ToString() + " - disabled by automated process on " + DateTime.Now.ToShortDateString() + " from WFN ";
                            result.DisplayName = "zz" + result.DisplayName.ToString();                            
                            result.UserPrincipalName = "zz" + result.UserPrincipalName.ToString();
                            result.Enabled = false;

                            result.Save();
                        }
                    }
                }               
            }
            catch (Exception ex)
            {
                
            }          
        }

        //old program source, do not use
        private string UpdateADfromWFN(string strEmail, string strMeta4ID, string strWork, string strLegalEntity, string strCompany, string strSBU, string strPositionCode, string strJob, string strJobTitleCode, string strWorkLocationCode, string strUSCostCenter, string strGlobalCostCenter, string strManager)
        {
            string strErrorLoc = "Begin";

            if (strEmail == "NO EMAIL" || strEmail == "NOEXIDEEMAIL" || strEmail == "NOEXIDEEMAIL@EXIDE.COM")
            {
                strEmail = string.Empty;
            }

            if (strSBU == "NO SBU")
            {
                strSBU = string.Empty;
            }

            if (strUSCostCenter == "NO US COST CENTER")
            {
                strUSCostCenter = string.Empty;
            }

            if (strManager == "NO MANAGER")
            {
                strManager = string.Empty;
            }

            strErrorLoc = "After update strings.";

            string strMessage = "Updated";
            string strFilter = string.Empty;

            if (strEmail == string.Empty)
            {

                strFilter = "((employeeNumber=" + strMeta4ID + "))";

            }
            else
            {

                strFilter = "(|(employeeNumber=" + strMeta4ID + ")(mail=" + strEmail + "))";

            }

            strErrorLoc = "Set Filter";

            try
            {
                strErrorLoc = "Set Directory entry";
                //DirectoryEntry de = new DirectoryEntry("LDAP://DC=Inside,DC=Exide,DC=ad", "inside\\weblogic_na", "webbea21", AuthenticationTypes.None);
                DirectoryEntry de = new DirectoryEntry();
                de.Path = "LDAP://10.138.212.222:389/DC=Inside,DC=Exide,DC=ad";
                de.Username = @"Inside\weblogic_na";
                de.Password = "webbea21";

                strErrorLoc = "Set Directory Searcher";
                DirectorySearcher search = new DirectorySearcher(strFilter);

                strErrorLoc = "Does a Search";
                SearchResult srMeta4 = search.FindOne();

                if (srMeta4 == null)
                {
                    strMessage = "Cannot find user.";
                }
                else
                {
                    strErrorLoc = "set Principal Context";
                    PrincipalContext AD = new PrincipalContext(ContextType.Domain, "inside", "weblogic_na", "webbea21");

                    strErrorLoc = "set User Principal";
                    UserPrincipal u = new UserPrincipal(AD);
                    ////u.SamAccountName = "bautzda";
                    //u.EmployeeId = strMeta4ID;

                    strErrorLoc = "check email";
                    if (srMeta4.Properties["mail"].Count > 0)
                    {
                        strErrorLoc = "set Email Address";

                        u.EmailAddress = srMeta4.Properties["mail"][0].ToString();

                        strErrorLoc = "set Principal Searcher";
                        PrincipalSearcher psSearch = new PrincipalSearcher(u);

                        strErrorLoc = "set User Principal";
                        UserPrincipal result = (UserPrincipal)psSearch.FindOne();

                        strErrorLoc = "Check result";
                        if (result != null)
                        {
                            try
                            {
                                strErrorLoc = "Set Directory Entry";
                                DirectoryEntry lowerldap = (DirectoryEntry)result.GetUnderlyingObject();

                                strErrorLoc = "Check Legal Entity";
                                if (strLegalEntity != string.Empty)
                                {
                                    lowerldap.Properties["LegalEntity"].Value = strLegalEntity;
                                }
                                strErrorLoc = "Check Company";
                                if (strCompany != string.Empty && (strLegalEntity.ToUpper().StartsWith("MTRX") || strLegalEntity.ToUpper().StartsWith("STTN") || strLegalEntity.ToUpper().StartsWith("ELRES")))
                                //if (strCompany != string.Empty)
                                {
                                    lowerldap.Properties["Company"].Value = strCompany;
                                }
                                strErrorLoc = "Check SBU";
                                if (strSBU != string.Empty)
                                {
                                    lowerldap.Properties["SBU"].Value = strSBU;
                                }

                                strErrorLoc = "check Position Code";
                                if (strPositionCode != string.Empty)
                                {
                                    lowerldap.Properties["PositionCode"].Value = strPositionCode;
                                }

                                strErrorLoc = "Check Job TItle Code";
                                if (strJobTitleCode != string.Empty)
                                {
                                    lowerldap.Properties["JobTitleCode"].Value = strJobTitleCode;
                                }

                                strErrorLoc = "Check work location code";
                                if (strWorkLocationCode != string.Empty)
                                {
                                    lowerldap.Properties["WorkLocationCode"].Value = strWorkLocationCode;
                                }

                                strErrorLoc = "check US Cost Center";
                                if (strUSCostCenter != string.Empty)
                                {
                                    lowerldap.Properties["USCostCenter"].Value = strUSCostCenter;
                                }

                                strErrorLoc = "Check global Cost Center";
                                if (strGlobalCostCenter != string.Empty)
                                {
                                    lowerldap.Properties["GlobalCostCenter"].Value = strGlobalCostCenter;
                                }

                                strErrorLoc = "update Meta4";
                                lowerldap.Properties["employeeNumber"].Value = strMeta4ID;

                                strErrorLoc = "Check Manager";
                                if (strManager != string.Empty)
                                {
                                    //lowerldap.Properties["Manager"].Value = GetManager(strManager);
                                }

                                strErrorLoc = "Update AD";
                                lowerldap.CommitChanges();

                                psSearch.Dispose();

                            }
                            catch (Exception ex)
                            {
                                //MessageBox.Show(ex.Message + " - " + strMeta4ID);
                                strMessage = strErrorLoc + " - " + ex.Message.ToString();
                            }
                        }

                    }

                }
            }
            catch (Exception ex)
            {
                strMessage = strErrorLoc + " - " + ex.Message.ToString();
            }

            return strMessage;

        }
        
    }
}
