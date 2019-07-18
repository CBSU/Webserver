using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using ClosedXML.Excel;
using Monthly_Reporting.Models;
using Monthly_Reporting.Properties;
using OfficeOpenXml;
namespace Monthly_Reporting.Controllers
{
    public class ReportsController : Controller
    {
        readonly SqlConnection sqlcon = new SqlConnection();
        Utility mutility = new Utility();
        readonly Reports rs = new Reports();
        ReportDateController CurrRunDate = new ReportDateController();
        DataTable downloadLoanStatement = new DataTable();
        string BrCode = "";
        // GET: Reports
        public ActionResult Index()
        {
            DataTable dt = new DataTable();
            DataTable dt_main_menu = new DataTable();
            DataTable dt_Report = new DataTable();
            DataTable dtRun = new DataTable();
            mutility.DbName = Settings.Default["DbName"].ToString();
            mutility.UserName = Settings.Default["UserName"].ToString();
            mutility.Password = Settings.Default["Password"].ToString();
            mutility.ServerName = Settings.Default["ServerName"].ToString();
            if (mutility.ServerName == "" || mutility.DbName == "" || mutility.UserName == "")
            {
                return RedirectToAction("Index", "Setting_Connection/Index");
            }
            else
            {
                if (Session["ID"] != null)
                {
                    UserAccessRightController UsAccRight = new UserAccessRightController();
                    string userkey = Convert.ToString(Session["user_key"]);
                    dt = UsAccRight.UserAccessRight(userkey);
                    dt_main_menu = UsAccRight.Main_Menu(userkey);
                    ViewBag.manin_menu = dt_main_menu;
                    ViewBag.sub_manin = dt;
                    string theString = Convert.ToString(Session["report_code"]);
                    var array = theString.Split(',');
                    string firstElem = array.First();
                    string restOfArray = string.Join("','", array.Skip(0));

                    string branchZone = Convert.ToString(mutility.dbSingleResult("select branch_zone from users where user_key='S001'"));

                    var zone_array = branchZone.Split(','); 
                    string zonetoarray = string.Join("','", zone_array.Skip(0));
                    string Sql = "select * from Table_Reports_temp where flag=1 and Report_Code in('"+ restOfArray + "')";
                    ViewBag.dt_Report = mutility.dbResult(Sql);
                    string brlist = "select * from BRANCH_LISTS where flag=1 and BrCode in('"+ zonetoarray + "');";
                    ViewBag.dt_ManageZone = mutility.dbResult(brlist);
                    ViewBag.dtRun = dtRun;
                    return View();
                }
                else
                {
                    return RedirectToAction("Index", "Login/Index");
                }
            }
        }
        [HttpPost]
        public ActionResult ExecuteScript(Reports rs)
        {
            DataTable dt = new DataTable();
            DataTable dt_main_menu = new DataTable();
            DataTable dt_Report = new DataTable();
            DataTable dtRun = new DataTable();
            mutility.DbName = Settings.Default["DbName"].ToString();
            mutility.UserName = Settings.Default["UserName"].ToString();
            mutility.Password = Settings.Default["Password"].ToString();
            mutility.ServerName = Settings.Default["ServerName"].ToString();
            if (mutility.ServerName == "" || mutility.DbName == "" || mutility.UserName == "")
            {
                return RedirectToAction("Index", "Setting_Connection/Index");
            }
            else
            {
                if (Session["ID"] != null)
                {
                    UserAccessRightController UsAccRight = new UserAccessRightController();
                    string userkey = Convert.ToString(Session["user_key"]);
                    dt = UsAccRight.UserAccessRight(userkey);
                    dt_main_menu = UsAccRight.Main_Menu(userkey);
                    ViewBag.manin_menu = dt_main_menu;
                    ViewBag.sub_manin = dt;
                    string theString = Convert.ToString(Session["report_code"]);
                    var array = theString.Split(',');
                    string firstElem = array.First();
                    string restOfArray = string.Join("','", array.Skip(0));

                    
                    string Sql = "select replace(Report_Code,' ','') as Report_Code, Name from Table_Reports_temp where flag=1 and Report_Code in('" + restOfArray + "')";
                    ViewBag.dt_Report = mutility.dbResult(Sql);

                    string branchZone = Convert.ToString(mutility.dbSingleResult("select branch_zone from users where user_key='S001'"));
                    var zone_array = branchZone.Split(',');
                    string zonetoarray = string.Join("','", zone_array.Skip(0));
                    string brlist = "select * from BRANCH_LISTS where flag=1 and BrCode in('" + zonetoarray + "');";

                    ViewBag.dt_ManageZone = mutility.dbResult(brlist);

                    string dtExcute = "";
                    DataTable dtAffterEx = new DataTable();              
                    
                  string ReExcuteName = "select StrSql from Table_Reports_temp where Report_Code='" + rs.ReportCode.Trim() + "'";                                          
                    dtExcute = Convert.ToString(mutility.dbSingleResult(ReExcuteName));                    
                    string Run = dtExcute;

                    //Get ReportName
                    string ValueString = "select Name from Table_Reports_temp where Report_Code = '" + rs.ReportCode.Replace(" ", "") + "'";
                    string ReportName = Convert.ToString(mutility.dbSingleResult(ValueString));
                    //Get Branch Name
               
                    string BrName = Convert.ToString(mutility.dbSingleResult("select BrLetter from BRANCH_LISTS where flag=1 and BrCode='" + rs.BrCode+"'"));


                    //Selected data 
                    ViewBag.BrCode = rs.BrCode;
                    ViewBag.ReportCode =rs.ReportCode.Replace(" ","");


                    DataTable dtstatus = new DataTable();
                    string BrZone_loop = Convert.ToString(mutility.dbSingleResult("select branch_zone from users where user_key='S001'"));
                    var zone_array_loop = BrZone_loop.Split(',');
                    string zonetoarray_loop = string.Join("','", zone_array_loop.Skip(0));
                    string SQl_loop = "";
                    if (rs.BrCode== "ALL")
                    {
                        SQl_loop = "select row_number() over(order by BL.BrCode) as id,*  ,1 as checked ,'0' as statusconnect from  BRANCH_LISTS BL inner join VPN V ON V.BrCode=BL.BrCode where BL.flag=1 and  BL.BrCode in('" + zonetoarray_loop + "')";
                    }
                    else
                    {
                         SQl_loop = "select row_number() over(order by BL.BrCode) as id,*  ,1 as checked ,'0' as statusconnect from  BRANCH_LISTS BL inner join VPN V ON V.BrCode=BL.BrCode where BL.flag=1 and BL.BrCode='" + rs.BrCode + "'";//BL.BrCode in('" + zonetoarray_loop + "')
                    }
                    
                    dtstatus = mutility.dbResult(SQl_loop);
                    SqlConnection sqlcon = new SqlConnection();                  
                    DataSet ds = new DataSet();
                    DataTable table = new DataTable();                    
                    foreach (DataRow dr in dtstatus.Rows)
                    {
                        sqlcon = mutility.BrConnection(dr["FullDbName"].ToString(), "sa", "Sa@#$Mbwin", dr["VPN"].ToString());
                        if (sqlcon.State != ConnectionState.Open)
                        {
                            try
                            {
                                dtRun = mutility.dbBranchResult(Run, sqlcon);
                                ds.Tables.Add(dtRun);
                            }
                            catch (Exception){
                            }                          
                        }                        
                    }
                    DataTable publicdt = new DataTable();
                    for(int i=0;i< ds.Tables.Count; i++)
                    {
                        publicdt.Merge(ds.Tables[i]);
                    }
                    if (rs.BrCode == "ALL")
                    {
                        BrName = "All_Branch";
                    }
                    
                    ExportDataToExcel(publicdt, BrName+"_"+ReportName);
                    TempData["sms"] = "You are already donwload reports";

                    return View("Index");                   
                }
                else
                {
                    return RedirectToAction("Index", "Login/Index");
                }
            }            
        }

        public void ExportDataToExcel(DataTable dt, string fileName)
        {
            using (XLWorkbook wb = new XLWorkbook())
            {
                wb.Worksheets.Add(dt,"sheet1");
                wb.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                wb.Style.Font.Bold = true;

                Response.Clear();
                Response.Buffer = true;
                Response.Charset = "";
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment;filename=" + fileName + ".xlsx");

                using (MemoryStream MyMemoryStream = new MemoryStream())
                {
                    wb.SaveAs(MyMemoryStream);
                    MyMemoryStream.WriteTo(Response.OutputStream);
                    Response.Flush();
                    Response.End();
                }
            }
        }
        public DataTable brlist(string userkey)
        {
            DataTable dt = new DataTable();
            string branchZone = Convert.ToString(mutility.dbSingleResult("select branch_zone from users where user_key='" + userkey + "'"));
            var zone_array = branchZone.Split(',');
            string zonetoarray = string.Join("','", zone_array.Skip(0));
            string brlist = "select * from BRANCH_LISTS where flag=1 and BrCode in('" + zonetoarray + "');";
            dt = mutility.dbResult(brlist);
            return dt;
        }

        [HttpGet]
        public ActionResult ExportExcel(string reportcode)
        {            
            DataTable dt = new DataTable();
            DataTable dt_main_menu = new DataTable();
            DataTable dtRun = new DataTable();
            DataTable dt_Report = new DataTable();
            mutility.DbName = Settings.Default["DbName"].ToString();
            mutility.UserName = Settings.Default["UserName"].ToString();
            mutility.Password = Settings.Default["Password"].ToString();
            mutility.ServerName = Settings.Default["ServerName"].ToString();
            if (mutility.ServerName == "" || mutility.DbName == "" || mutility.UserName == "")
            {
                return RedirectToAction("Index", "Setting_Connection/Index");
            }
            else
            {
                if (Session["ID"] != null)
                {
                    UserAccessRightController UsAccRight = new UserAccessRightController();
                    string userkey = Convert.ToString(Session["user_key"]);
                    dt = UsAccRight.UserAccessRight(userkey);
                    dt_main_menu = UsAccRight.Main_Menu(userkey);
                    ViewBag.manin_menu = dt_main_menu;
                    ViewBag.sub_manin = dt;
                    string theString = Convert.ToString(Session["report_code"]);
                    var array = theString.Split(',');
                    string firstElem = array.First();
                    string restOfArray = string.Join("','", array.Skip(0));
                    string branchZone = Convert.ToString(mutility.dbSingleResult("select branch_zone from users where user_key='S001'"));
                    var zone_array = branchZone.Split(',');
                    string zonetoarray = string.Join("','", zone_array.Skip(0));
                    string Sql = "select * from Table_Reports_temp where Report_Code in('" + restOfArray + "')";
                    ViewBag.dt_Report = mutility.dbResult(Sql);
                    string brlist = "select * from BRANCH_LISTS where flag=1 and BrCode in('" + zonetoarray + "');";
                    ViewBag.dt_ManageZone = mutility.dbResult(brlist);
                    string dtExcute = "";
                    DataTable dtAffterEx = new DataTable();
                    string ReExcuteName = "select StrSql from Table_Reports_temp where Report_Code='"+ reportcode + "'";
                    dtExcute = Convert.ToString(mutility.dbSingleResult(ReExcuteName));
                    string Run = dtExcute;
                   dtRun = mutility.dbResult(Run);
                    ViewBag.dtRun = dtRun;
                    XLWorkbook wb = new XLWorkbook();
                    wb.Worksheets.Add(dtRun, "Sheet1");
                    Response.Clear();
                    Response.Buffer = true;
                    Response.Charset = "";
                    Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                    Response.AddHeader("content-disposition", "attachment: filename=" + "ExcelReport.xlsx");
                    using (MemoryStream MyMemoryStream = new MemoryStream())
                    {
                        wb.SaveAs(MyMemoryStream);
                        MyMemoryStream.WriteTo(Response.OutputStream);
                        Response.Flush();
                        Response.End();
                    }
                    Response.End();
                    return View("index");
                }
                else
                {
                    return RedirectToAction("Index", "Login/Index");
                }

            }
        }
        
        public ActionResult LoanStatment()
        {
            DataTable dt = new DataTable();
            DataTable dt_main_menu = new DataTable();
            DataTable dtRun = new DataTable();
            DataTable dt_Report = new DataTable();
            mutility.DbName = Settings.Default["DbName"].ToString();
            mutility.UserName = Settings.Default["UserName"].ToString();
            mutility.Password = Settings.Default["Password"].ToString();
            mutility.ServerName = Settings.Default["ServerName"].ToString();
            if (mutility.ServerName == "" || mutility.DbName == "" || mutility.UserName == "")
            {
                return RedirectToAction("Index", "Setting_Connection/Index");
            }
            else
            {
                if (Session["ID"] != null)
                {
                    UserAccessRightController UsAccRight = new UserAccessRightController();
                    string userkey = Convert.ToString(Session["user_key"]);
                    dt = UsAccRight.UserAccessRight(userkey);
                    dt_main_menu = UsAccRight.Main_Menu(userkey);
                    ViewBag.manin_menu = dt_main_menu;
                    ViewBag.sub_manin = dt;
                    ViewBag.dt_ManageZone = this.brlist(userkey);

                    string Sql =@"";

                    ViewBag.CurrRunDate = CurrRunDate.CurrRunDate(userkey);
                    ViewBag.dt_Report = mutility.dtOnlyOneBranch(Sql,rs.BrCode);

                    return View();
                }
                else
                {
                    return RedirectToAction("Index", "Login/Index");
                }

            }
        }
     
        [HttpPost]
        //Action LoanStatment
        
        public ActionResult ExeLoanStatment(Reports re)
        {
            DataTable dt = new DataTable();
            DataTable dt_main_menu = new DataTable();
            DataTable dtRun = new DataTable();
            DataTable dt_Report = new DataTable();
            mutility.DbName = Settings.Default["DbName"].ToString();
            mutility.UserName = Settings.Default["UserName"].ToString();
            mutility.Password = Settings.Default["Password"].ToString();
            mutility.ServerName = Settings.Default["ServerName"].ToString();
            if (mutility.ServerName == "" || mutility.DbName == "" || mutility.UserName == "")
            {
                return RedirectToAction("Index", "Setting_Connection/Index");
            }
            else
            {
                if (Session["ID"] != null)
                {
                    UserAccessRightController UsAccRight = new UserAccessRightController();
                    string userkey = Convert.ToString(Session["user_key"]);
                    dt = UsAccRight.UserAccessRight(userkey);
                    dt_main_menu = UsAccRight.Main_Menu(userkey);
                    ViewBag.manin_menu = dt_main_menu;
                    ViewBag.sub_manin = dt;
                    ViewBag.dt_ManageZone = this.brlist(userkey);

                    DateTime datestart =re.datestart;
                    DateTime dateend =re.dateend;
                    string All = re.all;
                    string Sql = "";
                    if (All == "on")
                    {
                        Sql = @"
                    set dateformat DMY
                    declare @TotalAcc int
                    declare @Reportfrom datetime
                    declare @Reportto datetime
                    declare @i int
                    declare @BrCode nvarchar(10)
                    declare @BrShort nvarchar(10)
                    declare @BrName nvarchar(50)
                    declare @DbName nvarchar(50)

                    set @DbName = (SELECT DB_NAME())
                    select
                        @BrCode = SubBranchCode,
                        @BrShort = SubBranchID,
                        @BrName = SubBranchNameLatin
                    from skp_brlist
                    where DBName = (select left(@DbName, len(@DbName) - 4))
	
                    set @i = 1
                    declare @Acc nvarchar(15)

                    ----------------------------------------------
                    set @Reportfrom = '" + datestart + @"'
                    set @Reportto = '" + dateend + @"'



                    -- - prepared temptable to check account need to be loop and row number
                    if object_id('tempdb..#PreioueTrn') is not null

                        drop table #PreioueTrn

                    if object_id('tempdb..#LoopAcc') is not null

                        drop table #LoopAcc

                    Create table #PreioueTrn	
	                    (
                        name1 nvarchar(30),
	                    name2 nvarchar(30),
	                    DisplayName nvarchar(50),
	                    CID varchar(6),
	                    AccountNumber nvarchar(30),
	                    Trn nvarchar(50),
	                    TrnDate datetime,
                        ValueDate datetime,
	                    TrnType varchar(30),
	                    ShortDesc nvarchar(200),
	                    FullDesc nvarchar(max),
	                    TrnAmt numeric(18,3),
	                    TrnPenAmt numeric(18,3),
	                    TrnPrinAmt numeric(18,3),
	                    TrnIntAmt numeric(18,3),
	                    BalAmt numeric(18,3),
	                    TrnDesc nvarchar(max)
	                    );



                                        select

                        row_number() over(order by Acc) as RNo,
	                    Acc
                    into #LoopAcc 
                    from Lnacc
                    --where accstatus between '11' and '98'
                    ---------------------------------------------- -
                    set @TotalAcc = (select count(*) from #LoopAcc)
                    ------------------------------------------ -
                    --start Loop
                    while @i < @TotalAcc + 1
                    begin
                        -- - execute script...f
                          --set @Acc = (select left(Acc, len(Acc) - 1) from ##LoopAcc where RNo = @i )
	                    set @Acc = (select Acc from #LoopAcc where RNo = @i )
	                    Declare @CcyDiv Int

                        Set @CcyDiv = (Select CcyDiv From Ccy)


                    --inset into temTable...
                    insert into #PreioueTrn
                        (
                            name1,
                            name2,
                            DisplayName,
                            CID,
                            AccountNumber,
                            Trn,
                            TrnDate,
                            ValueDate,
                            TrnType,
                            ShortDesc,
                            FullDesc,
                            TrnAmt,
                            TrnPenAmt,
                            TrnPrinAmt,
                            TrnIntAmt,
                            BalAmt,
                            TrnDesc)

                    Select

                        c.name1,
	                    c.name2,
	                    C.DisplayName,
	                     Substring(t.Acc, 4, 6) as CID, (t.Acc + T.chd) as AA, t.Trn, t.TrnDate, t.ValueDate, t.TrnType, l.ShortDesc, l.FullDesc, 
	                    t.trnAmt / @CcyDiv as TrnAmt, 
	                    t.TrnPenAmt / @CcyDiv as TrnPenalty,
	                    t.TrnPriAmt / @CcyDiv as TrnPriAmt, 
	                    t.TrnIntAmt / @CcyDiv as TrnIntAmt, 
	                    t.BalAmt / @CcyDiv as BalAmt, t.TrnDesc
                      From TrnHist t
                    inner join[Lookup] l on t.TrnType = l.LookupCode and l.lookupid = 'TT' and l.LangType = '001'
                    LEFT JOIN CIF C ON Substring(T.Acc,4,6)= C.CID
                    where Acc = @Acc--505 10500153004 4
                    Order by t.TrnDate, T.Trn

                        ----incress to next record for execute
                        set @i = @i + 1
                    end

                    select

                         AccountNumber,
                         CID,
                         DisplayName,
                         Trn,
                         TrnDate,
                         ValueDate,
                         TrnType,
                         ShortDesc,
                         FullDesc,
                         TrnAmt,
                         TrnPenAmt,
                         TrnPrinAmt,
                         TrnIntAmt,
                         BalAmt,
                         TrnDesc,
                         @BrCode as BrCode,
                         @BrShort as BrName

                         from #PreioueTrn
	                     where ValueDate between @Reportfrom and @Reportto

                         order by ValueDate
                    truncate table #LoopAcc
                    drop table #LoopAcc

                    ";
                    }
                    else
                    {
                   Sql = @"
                    set dateformat DMY
                    declare @TotalAcc int
                    declare @Reportfrom datetime
                    declare @Reportto datetime
                    declare @i int
                    declare @BrCode nvarchar(10)
                    declare @BrShort nvarchar(10)
                    declare @BrName nvarchar(50)
                    declare @DbName nvarchar(50)

                    set @DbName = (SELECT DB_NAME())
                    select
                        @BrCode = SubBranchCode,
                        @BrShort = SubBranchID,
                        @BrName = SubBranchNameLatin
                    from skp_brlist
                    where DBName = (select left(@DbName, len(@DbName) - 4))
	
                    set @i = 1
                    declare @Acc nvarchar(15)

                    ----------------------------------------------
                    set @Reportfrom = '" + datestart + @"'
                    set @Reportto = '" + dateend + @"'



                    -- - prepared temptable to check account need to be loop and row number
                    if object_id('tempdb..#PreioueTrn') is not null

                        drop table #PreioueTrn

                    if object_id('tempdb..#LoopAcc') is not null

                        drop table #LoopAcc

                    Create table #PreioueTrn	
	                    (
                        name1 nvarchar(30),
	                    name2 nvarchar(30),
	                    DisplayName nvarchar(50),
	                    CID varchar(6),
	                    AccountNumber nvarchar(30),
	                    Trn nvarchar(50),
	                    TrnDate datetime,
                        ValueDate datetime,
	                    TrnType varchar(30),
	                    ShortDesc nvarchar(200),
	                    FullDesc nvarchar(max),
	                    TrnAmt numeric(18,3),
	                    TrnPenAmt numeric(18,3),
	                    TrnPrinAmt numeric(18,3),
	                    TrnIntAmt numeric(18,3),
	                    BalAmt numeric(18,3),
	                    TrnDesc nvarchar(max)
	                    );



                                        select

                        row_number() over(order by Acc) as RNo,
	                    Acc
                    into #LoopAcc 
                    from Lnacc where Acc+chd='"+re.accountnumber+@"'
                    --where accstatus between '11' and '98' and Acc+chd='" + re.accountnumber+@"'
                    ---------------------------------------------- -
                    set @TotalAcc = (select count(*) from #LoopAcc)
                    ------------------------------------------ -
                    --start Loop
                    while @i < @TotalAcc + 1
                    begin
                        -- - execute script...f
                          --set @Acc = (select left(Acc, len(Acc) - 1) from ##LoopAcc where RNo = @i )
	                    set @Acc = (select Acc from #LoopAcc where RNo = @i )
	                    Declare @CcyDiv Int

                        Set @CcyDiv = (Select CcyDiv From Ccy)


                    --inset into temTable...
                    insert into #PreioueTrn
                        (
                            name1,
                            name2,
                            DisplayName,
                            CID,
                            AccountNumber,
                            Trn,
                            TrnDate,
                            ValueDate,
                            TrnType,
                            ShortDesc,
                            FullDesc,
                            TrnAmt,
                            TrnPenAmt,
                            TrnPrinAmt,
                            TrnIntAmt,
                            BalAmt,
                            TrnDesc)

                    Select

                        c.name1,
	                    c.name2,
	                    C.DisplayName,
	                     Substring(t.Acc, 4, 6) as CID, (t.Acc + T.chd) as AA, t.Trn, t.TrnDate, t.ValueDate, t.TrnType, l.ShortDesc, l.FullDesc, 
	                    t.trnAmt / @CcyDiv as TrnAmt, 
	                    t.TrnPenAmt / @CcyDiv as TrnPenalty,
	                    t.TrnPriAmt / @CcyDiv as TrnPriAmt, 
	                    t.TrnIntAmt / @CcyDiv as TrnIntAmt, 
	                    t.BalAmt / @CcyDiv as BalAmt, t.TrnDesc
                      From TrnHist t
                    inner join[Lookup] l on t.TrnType = l.LookupCode and l.lookupid = 'TT' and l.LangType = '001'
                    LEFT JOIN CIF C ON Substring(T.Acc,4,6)= C.CID
                    where Acc = @Acc--505 10500153004 4
                    Order by t.TrnDate, T.Trn

                        ----incress to next record for execute
                        set @i = @i + 1
                    end

                    select

                         AccountNumber,
                         CID,
                         DisplayName,
                         Trn,
                         TrnDate,
                         ValueDate,
                         TrnType,
                         ShortDesc,
                         FullDesc,
                         TrnAmt,
                         TrnPenAmt,
                         TrnPrinAmt,
                         TrnIntAmt,
                         BalAmt,
                         TrnDesc,
                         @BrCode as BrCode,
                         @BrShort as BrName

                         from #PreioueTrn
	                     where ValueDate between @Reportfrom and @Reportto

                         order by ValueDate
                    truncate table #LoopAcc
                    drop table #LoopAcc

                    ";
                    }


                    
                    ViewBag.BrCode = re.BrCode;
                    ViewBag.datestart = re.datestart;
                    ViewBag.dateend = re.dateend;
                    ViewBag.Acc = re.accountnumber;

                    ViewBag.CurrRunDate = CurrRunDate.CurrRunDate(userkey);
                    ViewBag.dt_Report = mutility.dtOnlyOneBranch(Sql, re.BrCode);
                    downloadLoanStatement = mutility.dtOnlyOneBranch(Sql, re.BrCode);
                    BrCode = re.BrCode;
                    if(re.download== "download")
                    {
                        string ReportName = "Loan Statement Reports";
                        string BrName = Convert.ToString(mutility.dbSingleResult("select BrLetter from BRANCH_LISTS where flag=1 and BrCode='" + re.BrCode + "'"));
                        ExportDataToExcel(downloadLoanStatement, BrName + "_" + ReportName);
                    }
                    

                    return View("LoanStatment");
                }
                else
                {
                    return RedirectToAction("Index", "Login/Index");
                }
            }
            
        }
  
        public ActionResult LoanOverdue()
        {
            DataTable dt = new DataTable();
            DataTable dt_main_menu = new DataTable();
            DataTable dtRun = new DataTable();
            DataTable dt_Report = new DataTable();
            mutility.DbName = Settings.Default["DbName"].ToString();
            mutility.UserName = Settings.Default["UserName"].ToString();
            mutility.Password = Settings.Default["Password"].ToString();
            mutility.ServerName = Settings.Default["ServerName"].ToString();
            if (mutility.ServerName == "" || mutility.DbName == "" || mutility.UserName == "")
            {
                return RedirectToAction("Index", "Setting_Connection/Index");
            }
            else
            {
                if (Session["ID"] != null)
                {
                    UserAccessRightController UsAccRight = new UserAccessRightController();
                    string userkey = Convert.ToString(Session["user_key"]);
                    dt = UsAccRight.UserAccessRight(userkey);
                    dt_main_menu = UsAccRight.Main_Menu(userkey);
                    ViewBag.manin_menu = dt_main_menu;
                    ViewBag.sub_manin = dt;
                    ViewBag.dt_ManageZone = this.brlist(userkey);

                    string Sql = @"
                    
                    ";


                    ViewBag.CurrRunDate = CurrRunDate.CurrRunDate(userkey);
                    ViewBag.dt_Report = mutility.dtOnlyOneBranch(Sql, rs.BrCode);

                    return View();
                }
                else
                {
                    return RedirectToAction("Index", "Login/Index");
                }

            }
        }

        public ActionResult SummaryByCos()
        {
            DataTable dt = new DataTable();
            DataTable dt_main_menu = new DataTable();
            DataTable dtRun = new DataTable();
            DataTable dt_Report = new DataTable();
            mutility.DbName = Settings.Default["DbName"].ToString();
            mutility.UserName = Settings.Default["UserName"].ToString();
            mutility.Password = Settings.Default["Password"].ToString();
            mutility.ServerName = Settings.Default["ServerName"].ToString();
            if (mutility.ServerName == "" || mutility.DbName == "" || mutility.UserName == "")
            {
                return RedirectToAction("Index", "Setting_Connection/Index");
            }
            else
            {
                if (Session["ID"] != null)
                {
                    UserAccessRightController UsAccRight = new UserAccessRightController();
                    string userkey = Convert.ToString(Session["user_key"]);
                    dt = UsAccRight.UserAccessRight(userkey);
                    dt_main_menu = UsAccRight.Main_Menu(userkey);
                    ViewBag.manin_menu = dt_main_menu;
                    ViewBag.sub_manin = dt;
                    ViewBag.dt_ManageZone = this.brlist(userkey);

                    string Sql = @"
                    
                    ";


                    ViewBag.CurrRunDate = CurrRunDate.CurrRunDate(userkey);
                    ViewBag.dt_Report = mutility.dtOnlyOneBranch(Sql, rs.BrCode);

                    return View();
                }
                else
                {
                    return RedirectToAction("Index", "Login/Index");
                }

            }
        }

        public ActionResult SummaryByBranch()
        {
            DataTable dt = new DataTable();
            DataTable dt_main_menu = new DataTable();
            DataTable dtRun = new DataTable();
            DataTable dt_Report = new DataTable();
            mutility.DbName = Settings.Default["DbName"].ToString();
            mutility.UserName = Settings.Default["UserName"].ToString();
            mutility.Password = Settings.Default["Password"].ToString();
            mutility.ServerName = Settings.Default["ServerName"].ToString();
            if (mutility.ServerName == "" || mutility.DbName == "" || mutility.UserName == "")
            {
                return RedirectToAction("Index", "Setting_Connection/Index");
            }
            else
            {
                if (Session["ID"] != null)
                {
                    UserAccessRightController UsAccRight = new UserAccessRightController();
                    string userkey = Convert.ToString(Session["user_key"]);
                    dt = UsAccRight.UserAccessRight(userkey);
                    dt_main_menu = UsAccRight.Main_Menu(userkey);
                    ViewBag.manin_menu = dt_main_menu;
                    ViewBag.sub_manin = dt;
                    ViewBag.dt_ManageZone = this.brlist(userkey);

                    string Sql = @"
                    
                    ";


                    ViewBag.CurrRunDate = CurrRunDate.CurrRunDate(userkey);
                    ViewBag.dt_Report = mutility.dtOnlyOneBranch(Sql, rs.BrCode);

                    return View();
                }
                else
                {
                    return RedirectToAction("Index", "Login/Index");
                }

            }
        }

        public ActionResult fulltrialbalance()
        {
            DataTable dt = new DataTable();
            DataTable dt_main_menu = new DataTable();
            DataTable dtRun = new DataTable();
            DataTable dt_Report = new DataTable();
            mutility.DbName = Settings.Default["DbName"].ToString();
            mutility.UserName = Settings.Default["UserName"].ToString();
            mutility.Password = Settings.Default["Password"].ToString();
            mutility.ServerName = Settings.Default["ServerName"].ToString();
            if (mutility.ServerName == "" || mutility.DbName == "" || mutility.UserName == "")
            {
                return RedirectToAction("Index", "Setting_Connection/Index");
            }
            else
            {
                if (Session["ID"] != null)
                {
                    UserAccessRightController UsAccRight = new UserAccessRightController();
                    string userkey = Convert.ToString(Session["user_key"]);
                    dt = UsAccRight.UserAccessRight(userkey);
                    dt_main_menu = UsAccRight.Main_Menu(userkey);
                    ViewBag.manin_menu = dt_main_menu;
                    ViewBag.sub_manin = dt;
                    ViewBag.dt_ManageZone = this.brlist(userkey);

                    string Sql = @"
                    
                    ";


                    ViewBag.CurrRunDate = CurrRunDate.CurrRunDate(userkey);
                    ViewBag.dt_Report = mutility.dtOnlyOneBranch(Sql, rs.BrCode);

                    return View();
                }
                else
                {
                    return RedirectToAction("Index", "Login/Index");
                }

            }
        }

        public ActionResult GLTransactionbybatch()
        {
            DataTable dt = new DataTable();
            DataTable dt_main_menu = new DataTable();
            DataTable dtRun = new DataTable();
            DataTable dt_Report = new DataTable();
            mutility.DbName = Settings.Default["DbName"].ToString();
            mutility.UserName = Settings.Default["UserName"].ToString();
            mutility.Password = Settings.Default["Password"].ToString();
            mutility.ServerName = Settings.Default["ServerName"].ToString();
            if (mutility.ServerName == "" || mutility.DbName == "" || mutility.UserName == "")
            {
                return RedirectToAction("Index", "Setting_Connection/Index");
            }
            else
            {
                if (Session["ID"] != null)
                {
                    UserAccessRightController UsAccRight = new UserAccessRightController();
                    string userkey = Convert.ToString(Session["user_key"]);
                    dt = UsAccRight.UserAccessRight(userkey);
                    dt_main_menu = UsAccRight.Main_Menu(userkey);
                    ViewBag.manin_menu = dt_main_menu;
                    ViewBag.sub_manin = dt;
                    ViewBag.dt_ManageZone = this.brlist(userkey);

                    string Sql = @"
                    
                    ";


                    ViewBag.CurrRunDate = CurrRunDate.CurrRunDate(userkey);
                    ViewBag.dt_Report = mutility.dtOnlyOneBranch(Sql, rs.BrCode);

                    return View();
                }
                else
                {
                    return RedirectToAction("Index", "Login/Index");
                }

            }
        }

        public ActionResult CollateralInTool()
        {
            DataTable dt = new DataTable();
            DataTable dt_main_menu = new DataTable();
            DataTable dtRun = new DataTable();
            DataTable dt_Report = new DataTable();
            mutility.DbName = Settings.Default["DbName"].ToString();
            mutility.UserName = Settings.Default["UserName"].ToString();
            mutility.Password = Settings.Default["Password"].ToString();
            mutility.ServerName = Settings.Default["ServerName"].ToString();
            if (mutility.ServerName == "" || mutility.DbName == "" || mutility.UserName == "")
            {
                return RedirectToAction("Index", "Setting_Connection/Index");
            }
            else
            {
                if (Session["ID"] != null)
                {
                    UserAccessRightController UsAccRight = new UserAccessRightController();
                    string userkey = Convert.ToString(Session["user_key"]);
                    dt = UsAccRight.UserAccessRight(userkey);
                    dt_main_menu = UsAccRight.Main_Menu(userkey);
                    ViewBag.manin_menu = dt_main_menu;
                    ViewBag.sub_manin = dt;
                    ViewBag.dt_ManageZone = this.brlist(userkey);

                    string Sql = @"
                    
                    ";


                    ViewBag.CurrRunDate = CurrRunDate.CurrRunDate(userkey);
                    ViewBag.dt_Report = mutility.dtOnlyOneBranch(Sql, rs.BrCode);

                    return View();
                }
                else
                {
                    return RedirectToAction("Index", "Login/Index");
                }

            }
        }
        public ActionResult ListoffWrittenOff()
        {
            DataTable dt = new DataTable();
            DataTable dt_main_menu = new DataTable();
            DataTable dtRun = new DataTable();
            DataTable dt_Report = new DataTable();
            mutility.DbName = Settings.Default["DbName"].ToString();
            mutility.UserName = Settings.Default["UserName"].ToString();
            mutility.Password = Settings.Default["Password"].ToString();
            mutility.ServerName = Settings.Default["ServerName"].ToString();
            if (mutility.ServerName == "" || mutility.DbName == "" || mutility.UserName == "")
            {
                return RedirectToAction("Index", "Setting_Connection/Index");
            }
            else
            {
                if (Session["ID"] != null)
                {
                    UserAccessRightController UsAccRight = new UserAccessRightController();
                    string userkey = Convert.ToString(Session["user_key"]);
                    dt = UsAccRight.UserAccessRight(userkey);
                    dt_main_menu = UsAccRight.Main_Menu(userkey);
                    ViewBag.manin_menu = dt_main_menu;
                    ViewBag.sub_manin = dt;
                    ViewBag.dt_ManageZone = this.brlist(userkey);

                    string Sql = @"
                    
                    ";


                    ViewBag.CurrRunDate = CurrRunDate.CurrRunDate(userkey);
                    ViewBag.dt_Report = mutility.dtOnlyOneBranch(Sql, rs.BrCode);

                    return View();
                }
                else
                {
                    return RedirectToAction("Index", "Login/Index");
                }

            }
        }
        public ActionResult WrittenoffCollection()
        {
            DataTable dt = new DataTable();
            DataTable dt_main_menu = new DataTable();
            DataTable dtRun = new DataTable();
            DataTable dt_Report = new DataTable();
            mutility.DbName = Settings.Default["DbName"].ToString();
            mutility.UserName = Settings.Default["UserName"].ToString();
            mutility.Password = Settings.Default["Password"].ToString();
            mutility.ServerName = Settings.Default["ServerName"].ToString();
            if (mutility.ServerName == "" || mutility.DbName == "" || mutility.UserName == "")
            {
                return RedirectToAction("Index", "Setting_Connection/Index");
            }
            else
            {
                if (Session["ID"] != null)
                {
                    UserAccessRightController UsAccRight = new UserAccessRightController();
                    string userkey = Convert.ToString(Session["user_key"]);
                    dt = UsAccRight.UserAccessRight(userkey);
                    dt_main_menu = UsAccRight.Main_Menu(userkey);
                    ViewBag.manin_menu = dt_main_menu;
                    ViewBag.sub_manin = dt;
                    ViewBag.dt_ManageZone = this.brlist(userkey);

                    string Sql = @"
                    
                    ";


                    ViewBag.CurrRunDate = CurrRunDate.CurrRunDate(userkey);
                    ViewBag.dt_Report = mutility.dtOnlyOneBranch(Sql, rs.BrCode);

                    return View();
                }
                else
                {
                    return RedirectToAction("Index", "Login/Index");
                }

            }
        }
    }
}
        
    
