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
            string ReportName = "";
            string BrName = "";
            string Sql = "";
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


                    Sql = "select replace(Report_Code,' ','') as Report_Code, Name from Table_Reports_temp where flag=1 and Report_Code in('" + restOfArray + "')";
                    ViewBag.dt_Report = mutility.dbResult(Sql);

                    //Selected data 
                    ViewBag.BrCode = rs.BrCode;
                    ViewBag.ReportCode = rs.ReportCode.Replace(" ", "");

                    string branchZone = Convert.ToString(mutility.dbSingleResult("select branch_zone from users where user_key='S001'"));
                    var zone_array = branchZone.Split(',');
                    string zonetoarray = string.Join("','", zone_array.Skip(0));
                    string brlist = "select * from BRANCH_LISTS where flag=1 and BrCode in('" + zonetoarray + "');";

                    ViewBag.dt_ManageZone = mutility.dbResult(brlist);


                    if (rs.ReportCode == "1")
                    {
                        ViewBag.BrCode = rs.BrCode;
                        ViewBag.datestart = rs.datestart;
                        ViewBag.dateend = rs.dateend;
                        ViewBag.Acc = rs.accountnumber;
                        Sql = @"
            SET DATEFORMAT DMY
            DECLARE @REPORTDATE DATETIME
            DECLARE @BeginDateYear datetime
            DECLARE @BeginDateMonth datetime
            DECLARE @PreThreeMonthDate datetime
            DECLARE @ccydiv INT
            declare @MaxDayShort int
            declare @MaxDayLong int
            declare @TermDefined int
            -------------------------------------------------

                    --set @REPORTDATE = '30/06/2011';
                        --set @BeginDateYear = '01/01/2011';
                        --set @BeginDateMonth = '01/06/2011';
                        select
                            @REPORTDATE = CurrRunDate,
                            @BeginDateYear = convert(datetime, '01/01/' + cast(year(CurrRunDate) as varchar(4)), 103),
                            @BeginDateMonth = convert(datetime, '01/' + cast(month(CurrRunDate) as varchar(2)) + '/' + cast(year(CurrRunDate) as varchar(4)), 103),
                            @PreThreeMonthDate = dateadd(month, 3, CurrRunDate)

            from BRPARMS

            SET @ccydiv = (SELECT ccydiv FROM ccy)
            set @TermDefined = ((select top(1) max(M1.Term) from AmretLoanProvision M1))
            set @MaxDayShort = (select max(NumDaysNo) from AmretLoanProvision where Term = 0 )
            set @MaxDayLong = (select max(NumDaysNo) from AmretLoanProvision where Term = @TermDefined)
			declare @BrCode nvarchar(10)

            declare @BrShort nvarchar(10)

            declare @BrName nvarchar(50)

            declare @DbName nvarchar(50)

            set @DbName = (SELECT DB_NAME())

			select
                @BrCode = SubBranchCode,
                @BrShort = SubBranchID,
                @BrName = SubBranchNameLatin

            from SKP_Brlist

            where DBName = (select left(@DbName, len(@DbName) - 4))

            --print '01/' + cast(month(CurrRunDate) as varchar(2))
            print @BeginDateMonth
            print @BeginDateYear
            -------------------------------------------------
            /*
            CLEAN TEMP TABLE 
            */
            IF OBJECT_ID('tempdb..#LN_DETAIL') IS NOT NULL
                DROP TABLE #LN_DETAIL
            IF OBJECT_ID('tempdb..#PaidOff') is not null

                drop table #PaidOff
            IF OBJECT_ID('tempdb..#RepayMonth_Temp') IS NOT NULL
                DROP TABLE #RepayMonth_Temp
            IF OBJECT_ID('tempdb..#RepayMonth') IS NOT NULL
                DROP TABLE #RepayMonth
            IF OBJECT_ID('tempdb..#LNCalcInterest') IS NOT NULL
                DROP TABLE #LNCalcInterest

            SELECT

                    L.Acc,
		            L.Chd,
		            /*
		            CO information
		            */
		            co2.CID as IdCO,		
		            ltrim(rtrim(co2.DisplayName)) as CoName,
		            ltrim(rtrim(CO_Kh.Name1)) as CoNameKh,	
		            G.CID AS IdGroup,
		            ltrim(rtrim(g2.DisplayName)) as GroupName,
		            ----------------------------------------
                    c.CID as IdClient,
		            ltrim(rtrim(c.DisplayName)) as ClientName,
		            ltrim(rtrim(CKh.Name1)) as ClientNameKh,		
		            case 
			            when c.GenderType = '001' then 'M'

                        when c.GenderType = '002' then 'F'
			            else	''

                    end as Gender,
		            case 
			            when c.CivilStatusCode = '00D' then 'D'

                        when c.CivilStatusCode = '00M' then 'M'

                        when c.CivilStatusCode = '00S' then 'S'

                        when c.CivilStatusCode = '00W' then 'W'
			            else ''

                    end as [MaritalStatus],
		            case 
			            when left(c.Nid,1) = 'M'  then 'M'--M = Maried Letter
                          when left(c.Nid, 1) = 'C' then 'C'--C = Civil Servan
                           when left(c.Nid, 1) = 'F' then 'F'

                        when left(c.Nid,1) = 'B' then 'B'

                        when left(c.Nid,1) = 'N' and ltrim(rtrim(c.nid)) <> 'N/A' then 'N'

                        when left(c.Nid,1)= 'D' then 'D'

                        when left(c.Nid,1)= 'P' then 'P'

                        when left(c.Nid,1) = 'R' then 'R'
			            else 'Unknown'

                    end as [IDType_1],--N,F,D,B,G,R...columns 2 in CBC
		            case 
			            when left(c.Nid,1) in ('N', 'F', 'R', 'D', 'B', 'C', 'P', 'M') and ltrim(rtrim(c.nid)) <> 'N/A' then right(ltrim(rtrim(c.Nid)),len(c.Nid) - 1)
			            when left(c.Nid,1) in ('F', 'G', 'R') then
				            case when CAST(right(ltrim(rtrim(c.Nid)),len(c.Nid) - 1) AS int) = 0 then 'N/A'
					            else right(c.Nid, len(c.Nid) - 1)

                            end
			            else c.Nid

                    end as [IDNumber_1],--3

                    c.BirthDate as [DateofBirth],		
		            isnull(ltrim(rtrim(c.Mobile1)), '') as Mobile1,
		            isnull(ltrim(rtrim(c.Mobile2)), '') as Mobile2,		
		            ltrim(rtrim(c.CIFCode1)) as CIFCode1,
		            ltrim(rtrim(c.CIFCode2)) as CIFCode2,
		            ltrim(rtrim(c.CIFCode3)) as CIFCode3,
		            ltrim(rtrim(c.CIFCode4)) as CIFCode4,
		            ltrim(rtrim(c.CIFCode5)) as CIFCode5,			
		            ltrim(rtrim(c.CIFCode6)) as CIFCode6,
		            ltrim(rtrim(c.CIFCode7)) as CIFCode7,
		            ltrim(rtrim(c.CIFCode8)) as CIFCode8,
		            ltrim(rtrim(c.CIFCode9)) as CIFCode9,
		            CIF_FamilyMember.FullDesc as FamilyMemberDesc,
		            CIF_Income.FullDesc as CIF_IncomeDesc,
		            CIF_Occupation.FullDesc as CIF_OccupationDesc,		
		            CIF_TotalAsset.FullDesc as CIF_TotalAssetDesc,
		            c.LocationCode,
		            VIL.NameLatin AS Village,
		            COM.NameLatin as Commune,
		            Dis.NameLatin as District,
		            PRV.NameLatin as Province,
		            Vil.NameKhmer + ';' + Com.NameKhmer + ';' + dis.NameKhmer + ';' + PRV.NameKhmer as AdressKhmer,
		            -------------------------------
                    ---Co - Borrowere Information

                    Ckh.Name2 as CoBorrowerName,
		            case 
			            when CKh.Sex = 'Rbus' then 'M'

                        when CKh.Sex = 'RsI' then 'F'
			            else ''

                    end as CoBorrowerGender,
		            --'' as CoBorrowerDoBOld,
		            Ckh.Date1 as CoBorrowerDoB,
		            Ckh.CardType as CoBorrowerIDType,
		            Ckh.Name6 as CoBorrowerIDNum,
		            Ckh.RelatedName as CoBorrowerRelativeType,
		            ------------------------------
                    ---Loan Information

                    L.AppType,
		            L.BalAmt / @ccydiv AS BalAmt,

                    ------
                    L.AcrIntAmt / @ccydiv as AcrIntAmt_EndPreMonth,---Accrue interest Ammount until end of prevouse month

                    L.AcrChgAmt / @ccydiv as AcrChgAmt_EndPreMonth,
		            L.AcrPenAmt / @ccydiv as AcrPenAmt_EndPreMonth,
		            L.AcrIntODuePriAmt / @ccydiv as AcrIntODuePriAmt,
		            ------
                    L.IntBalAmt / @ccydiv as IntBalAmt,
		            L.PenBalAmt / @ccydiv as PenBalAmt,
		            L.StopIntTF ,
		            L.TrnSeq,
		            -------
                    Case

                        WHEN L.OduePriAmt < 0 THEN 0 ELSE L.OduePriAmt / @ccydiv

                    END AS OduePriAmt,
		            L.OdueIntAmt / @ccydiv AS OdueIntAmt,
                    L.StopAcrIntTF ,
		            L.ReschedSeq,
		            L.CcyType,
		            dbo.[GetFirstDisbDate](l.acc) as DisbDate,
		            L.GrantedAmt / @ccydiv AS GrantedAmt,
                     L.IntRate,
		            L.IntEffDate,
		            L.MatDate,	
		            L.GLCode,
		            L.GLCodeOrig,	
		            L.InstNo,
		            L.FreqType,
		            case 
			            when l.FreqType = '012' then 'Monthly'--monthly
                       when l.FreqType = '026' then '2 Week'--every 2 week
                      when l.FreqType = '052' then 'Weekly'--weekly
                     when l.FreqType = '013' then '4 Week'--every 4 week
			            else ''-- - ************protect in case we add some more product->then we can saw error.

                    end as [PaymentFrequency],--16

                    L.PrType,
		            P.FullDesc as PrName,
		            L.LNCode1,
		            L.LNCode2,
		            L.LNCode3,
		            L.LNCode4,
		            L.LNCode5,
		            L.LNCode6,
		            L.LNCode7,
		            L.LNCode8,	
		            Dbo.GetTermOfLoan(L.Acc) AS TERMS,

                    -----------
                    CAST('01/01/1900' AS DATETIME) as LastPaidDate,
		            CAST(0 AS NUMERIC(18, 3)) as LastPaidAmt,
		            cast('' as nvarchar(15)) as PreAcc,
		            CAST(0 AS NUMERIC(18, 3)) as PreDisbAmt,
		            Cast('01/01/1900' AS DATETIME) as NextPaidDate,
		            cast('01/01/1900' as datetime) as NextDueWorkDate,
		            Cast(0 as numeric(18, 3)) as NextPaidAmt,
		            Cast(0 as numeric(18, 3)) as IntRemainInSchedule,
		            cast(0 as numeric(18, 3)) as RemainLoanTerm,
		            cast('01/01/1900' as datetime) as LastPreDueDateOfReportDate,
		            cast(0 as numeric(18, 3)) as BalAmtLastDueOfReportDate,
		            cast(0 as numeric(18, 3)) as PrinAmtLastDueDateCollect,
		            --cast('01/01/1900' as datetime) as LastTransDateDueDate,
		            --cast(0 as numeric(18, 3)) as IntAmtAfterPreDueDateOfReportDate,
		            --cast(0 as numeric(18, 3)) as TotalAmtRequiredToClose,
		
		            -----------
                    dbo.GetAgeOfLoan(@REPORTDATE, l.Acc) as AgeofLoan,
		            cast(0 as int) as NoInstLate,
		            CAST('01/01/1900' AS DATETIME) AS FirstDayLate,

                    --CAST('' as nvarchar(10)) as StatusReNewLoan,
		            -------------------------------------------------------------------------
                    ---Collection--OK

                    CAST(0 AS NUMERIC(18, 2)) as TrnIntAmt_CollYear,--reserved for collection
                    CAST(0 AS NUMERIC(18, 2)) as TrnPriAmt_CollYear, --reserved for collection
 
                     CAST(0 AS NUMERIC(18, 2)) as TrnPenAmt_CollYear, --reserved for collection
  
                      CAST(0 AS NUMERIC(18, 2)) as TrnChgAmt_CollYear, --reserved for collection
  
                      CAST(0 AS NUMERIC(18, 2)) as TrnIntAmt_CollAsOfMonth, --reserved for collection
   
                       CAST(0 AS NUMERIC(18, 2)) as TrnPriAmt_CollAsOfMonth, --reserved for collection
   
                       CAST(0 AS NUMERIC(18, 2)) as TrnPenAmt_CollAsOfMonth, --reserved for collection
    
                        CAST(0 AS NUMERIC(18, 2)) as TrnChgAmt_CollAsOfMonth, --reserved for collection
    
                        CAST(0 AS NUMERIC(18, 2)) as TrnIntAmt_CollDaily, --reserved for collection
     
                         CAST(0 AS NUMERIC(18, 2)) as TrnPriAmt_CollDaily, --reserved for collection
      
                          CAST(0 AS NUMERIC(18, 2)) as TrnPenAmt_CollDaily, --reserved for collection
       
                           CAST(0 AS NUMERIC(18, 2)) as TrnChgAmt_CollDaily, --reserved for collection
                           ------------------------------------------------------------------------------
                           -- - Interesr Culculation Procedure--OK
       
                           CAST(0 AS NUMERIC(18, 2)) as IntOdueAmt_FromIntCalc_Proc,
                           CAST(0 AS NUMERIC(18, 2)) as IntToDateAmt_FromIntCalc_Proc, --= @IntToDate = @IntOdueAmt(ARI from begining month To Date) + @AcrIntAmt(ARI end of Previous month that get from LNACC )
       
                           CAST(0 AS NUMERIC(18, 2)) as PenaltyAmtDue_FromIntCalc_Proc,
                           CAST(0 AS NUMERIC(18, 2)) as TaxAmtDue_FromIntCalc_Proc,
                           CAST(0 AS NUMERIC(18, 2)) as TrnChagAmtDue_FromIntCalc_Proc,

                           ------------------------------------------------------------------------------
                           ----PaidOff, paid off can be normal paid off(loan Past Maturity then paid off) or loan that paid off before paturity
       
                           CAST(0 AS NUMERIC(18, 2)) as TrnIntAmt_PaidOff, --reserved for collection
        
                            CAST(0 AS NUMERIC(18, 2)) as TrnPriAmt_PaidOff, --reserved for collection
         
                             CAST(0 AS NUMERIC(18, 2)) as TrnPenAmt_PaidOff, --reserved for collection
          
                              CAST(0 AS NUMERIC(18, 2)) as TrnChgAmt_PaidOff, --reserved for collection
          
                              CAST('01/01/1900' AS DATETIME) as PaidOffDate,
                              CAST('' AS NVARCHAR(30)) AS STATUS_RENEW,
                              -----------------------------------------------------------------
                              --Provision code
          
                              ShortCode =
          		            case 

                                  when dbo.GetTermOfLoan(L.acc) >= (select max(TERM) from AmretLoanProvision) then 1
			            else 0

                    end,
		            MasterField = cast('00000' as nvarchar(10)),
		            ----------------------------------------------------------------
                    L.AccStatus,
		            L.AccStatusDate,
		            BusType.FullDesc BusType,
                    CASE

                        WHEN L.LNCode1 BETWEEN '100' AND '199' THEN 'Agricultural'

                        WHEN L.LNCode1 BETWEEN '200' AND '399' THEN 'Trading & Commerce'

                        WHEN L.LNCode1 BETWEEN '400' AND '499' THEN 'Service'

                        WHEN L.LNCode1 BETWEEN '500' AND '599' THEN 'Transportation'

                        WHEN L.LNCode1 BETWEEN '600' AND '699' THEN 'Construction'

                        WHEN L.LNCode1 > '700' THEN 'Consumption (Householg/Family)'

                    END AS BusSector,
		            Colateral.FullDesc as LoanColateral,
		            LNCycle.FullDesc as LoanCycle,
		            /*
		            guaranter information
		            */
		            R.CID as idGuarenter,
		            ltrim(rtrim(GC.DisplayName)) as GuarenterName,
		            ltrim(rtrim(GCKh.Name1)) as GuarenterNameKh,		
		            case 
			            when GC.GenderType = '001' then 'M'

                        when GC.GenderType = '002' then 'F'
			            else	''

                    end as GuarenterGender,
		            case 
			            when GC.CivilStatusCode = '00D' then 'D'

                        when GC.CivilStatusCode = '00M' then 'M'

                        when GC.CivilStatusCode = '00S' then 'S'

                        when GC.CivilStatusCode = '00W' then 'W'
			            else ''

                    end as GuarenterMaritalStatus,
		            case 
			            when left(GC.Nid,1) = 'M'  then 'M'--M = Maried Letter
                          when left(GC.Nid, 1) = 'C' then 'C'--C = Civil Servan
                           when left(GC.Nid, 1) = 'F' then 'F'

                        when left(GC.Nid,1) = 'B' then 'B'

                        when left(GC.Nid,1) = 'N' and ltrim(rtrim(c.nid)) <> 'N/A' then 'N'

                        when left(GC.Nid,1)= 'D' then 'D'

                        when left(GC.Nid,1)= 'P' then 'P'

                        when left(GC.Nid,1) = 'R' then 'R'
			            else 'Unknown'

                    end as Guarenter_IDType_1,--N,F,D,B,G,R...columns 2 in CBC
		            case 
			            when left(GC.Nid,1) in ('N', 'F', 'R', 'D', 'B', 'C', 'P', 'M') and ltrim(rtrim(GC.nid)) <> 'N/A' then right(ltrim(rtrim(GC.Nid)),len(GC.Nid) - 1)
			            when left(GC.Nid,1) in ('F', 'G', 'R') then
				            case when CAST(right(ltrim(rtrim(GC.Nid)),len(GC.Nid) - 1) AS int) = 0 then 'N/A'
					            else right(GC.Nid, len(GC.Nid) - 1)

                            end
			            else GC.Nid

                    end as GuarenterIDNumber_1,--3

                    GC.BirthDate as GuarenterDateofBirth,		
		            isnull(ltrim(rtrim(GC.Mobile1)), '') as GuarenterMobile1,
		            isnull(ltrim(rtrim(GC.Mobile2)), '') as GuarenterMobile2,
		            GC.locationCode as GuarenterLocationCode,
		            /*
		            Co Guarenter information
		            */
		            GCkh.Name2 as CoGuarenerName,
		            case 
			            when CKh.Sex = 'Rbus' then 'M'

                        when CKh.Sex = 'RsI' then 'F'
			            else ''

                    end as CoGuarenerGender,
		            GCkh.Date1 as CoGuarenerDoB,
		            GCkh.CardType as CoGuarenerIDType,
		            GCkh.Name6 as CoGuarenerIDNum,
		            GCkh.RelatedName as CoGuarenerRelativeType,

		            @REPORTDATE AS REPORTDATE
                INTO #LN_DETAIL 
	            FROM LNACC L
                    LEFT JOIN CIF C ON C.CID = substring(L.ACC, 4, 6) and c.type = '001'--Client
                      LEFT JOIN relcid G ON G.relatedcid = C.CID AND G.TYPE = '900'--MEMBER TO GROUP
                    LEFT JOIN CIF G2 ON G2.CID = G.CID--JOIN TO GET GROUP INFORMATION
                  LEFT JOIN RELCID CO ON CO.RelatedCID = G.CID    AND CO.TYPE = '499'--GROUP TO CO
                  LEFT JOIN CIF CO2 ON CO2.CID = CO.CID

                    LEFT JOIN PRPARMS P ON P.PrType = L.PrType and P.AppType = 4

                    LEFT JOIN VTSDCIF CKh on Ckh.CID = c.CID and ckh.Type = '001'--join to get client name & co - borrower in khmer
                      LEFT JOIN VTSDVillage VIL ON VIL.Code = C.LocationCode

                    LEFT JOIN VTSDCommune COM ON COM.CODE = VIL.CodeOfCommune

                    LEFT JOIN VTSDDistrict DIS ON DIS.Code = COM.CodeOfDistrict

                    LEFT JOIN VTSDProvince PRV ON PRV.Code = DIS.CodeOfProvince

                    LEFT JOIN VTSDCIF CO_Kh ON CO_Kh.CID = co2.CID and CO_Kh.Type = '499'

                    LEFT JOIN USERLOOKUP BusType ON BusType.LookUpCode = L.LNCode1 and BusType.LookUpId = '41'

                    LEFT JOIN USERLOOKUP Colateral ON Colateral.LookUpCode = L.LNCode2 and Colateral.LookUpId = '42'

                    LEFT JOIN USERLOOKUP LNCycle on LNCycle.LookUpCode = L.LNCode3 and LNCycle.LookUpId = '43'

                    LEFT JOIN USERLOOKUP CIF_FamilyMember ON CIF_FamilyMember.LookUpCode = C.CIFCode1 and CIF_FamilyMember.LookUpId = '61'--62

                    LEFT JOIN USERLOOKUP CIF_Income ON CIF_Income.LookUpCode = C.CIFCode2 and CIF_Income.LookUpId = '62'

                    LEFT JOIN USERLOOKUP CIF_Occupation ON CIF_Occupation.LookUpCode = C.CIFCode3 and CIF_Occupation.LookUpId = '63'

                    LEFT JOIN USERLOOKUP CIF_TotalAsset ON CIF_TotalAsset.LookUpCode = C.CIFCode4 and CIF_TotalAsset.LookUpId = '64'

                    LEFT JOIN relacc R on R.Acc + R.Chd = L.Acc + L.Chd and R.Type = '030' and R.AppType = '4'--type = 30 for guarentee
                    LEFT JOIN CIF GC ON GC.CID = R.CID--guaretee

                    LEFT JOIN VTSDCIF GCKh on GCKh.CID = GC.CID and gckh.Type = '001'--join to get client name & co - borrower in khmer
                    --INNER JOIN #LNProvisions PRO ON PRO.Acc  = L.Acc+L.Chd		
	            WHERE L.AccStatus BETWEEN '11' AND '98'
            -- -==================================================================================================================================================
            ----FOR CLOSE LOAN IN MONTH

            INSERT INTO #LN_DETAIL
            SELECT

                    L.Acc,
                    L.Chd,
                    co2.CID as IdCO,
                    ltrim(rtrim(co2.DisplayName)) as CoName,
                    ltrim(rtrim(CO_Kh.Name1)) as CoNameKh,
                    G.CID AS IdGroup,
                    ltrim(rtrim(g2.DisplayName)) as GroupName,
                    ----------------------------------------
                    c.CID as IdClient,
                    ltrim(rtrim(c.DisplayName)) as ClientName,
                    ltrim(rtrim(CKh.Name1)) as ClientNameKh,
		            case 

                        when c.GenderType = '001' then 'M'

                        when c.GenderType = '002' then 'F'
			            else	''

                    end as Gender,
		            case 

                        when c.CivilStatusCode = '00D' then 'D'

                        when c.CivilStatusCode = '00M' then 'M'

                        when c.CivilStatusCode = '00S' then 'S'

                        when c.CivilStatusCode = '00W' then 'W'
			            else ''

                    end as [MaritalStatus],
		            case 

                        when left(c.Nid, 1) = 'M'  then 'M'--M = Maried Letter

                        when left(c.Nid, 1) = 'C' then 'C'--C = Civil Servan

                        when left(c.Nid, 1) = 'F' then 'F'

                        when left(c.Nid, 1) = 'B' then 'B'

                        when left(c.Nid, 1) = 'N' and ltrim(rtrim(c.nid)) <> 'N/A' then 'N'

                        when left(c.Nid, 1) = 'D' then 'D'

                        when left(c.Nid, 1) = 'P' then 'P'

                        when left(c.Nid, 1) = 'R' then 'R'
			            else 'Unknown'

                    end as [IDType_1], --N, F, D, B, G, R...columns 2 in CBC
		            case 

                        when left(c.Nid, 1) in ('N', 'F', 'R', 'D', 'B', 'C', 'P', 'M') and ltrim(rtrim(c.nid)) <> 'N/A' then  right(ltrim(rtrim(c.Nid)), len(c.Nid) - 1)

                        when  left(c.Nid, 1) in ('F', 'G', 'R') then
				            case when CAST(right(ltrim(rtrim(c.Nid)), len(c.Nid) - 1) AS int) = 0 then 'N/A' 
					            else right(c.Nid, len(c.Nid) - 1)

                            end
			            else c.Nid

                    end as [IDNumber_1], --3

                    c.BirthDate as [DateofBirth],
                    isnull(ltrim(rtrim(c.Mobile1)), '') as Mobile1,
                    isnull(ltrim(rtrim(c.Mobile2)), '') as Mobile2,
                    ltrim(rtrim(c.CIFCode1)) as CIFCode1,
                    ltrim(rtrim(c.CIFCode2)) as CIFCode2,
                    ltrim(rtrim(c.CIFCode3)) as CIFCode3,
                    ltrim(rtrim(c.CIFCode4)) as CIFCode4,
                    ltrim(rtrim(c.CIFCode5)) as CIFCode5,
                    ltrim(rtrim(c.CIFCode6)) as CIFCode6,
                    ltrim(rtrim(c.CIFCode7)) as CIFCode7,
                    ltrim(rtrim(c.CIFCode8)) as CIFCode8,
                    ltrim(rtrim(c.CIFCode9)) as CIFCode9,
                    CIF_FamilyMember.FullDesc as FamilyMemberDesc,
                    CIF_Income.FullDesc as CIF_IncomeDesc,
                    CIF_Occupation.FullDesc as CIF_OccupationDesc,
                    CIF_TotalAsset.FullDesc as CIF_TotalAssetDesc,
                    c.LocationCode,
                    VIL.NameLatin AS Village,
                    COM.NameLatin as Commune,
                    Dis.NameLatin as District,
                    PRV.NameLatin as Province,
                    Vil.NameKhmer + ';' + Com.NameKhmer + ';' + dis.NameKhmer + ';' + PRV.NameKhmer as AdressKhmer,
                    -------------------------------
                    ---garanter Information

                    Ckh.Name2 as CoBorrowerName,
		            case 

                        when CKh.Sex = 'Rbus' then 'M'

                        when CKh.Sex = 'RsI' then 'F'
			            else ''

                    end as CoBorrowerGender,
                    --'' as CoBorrowerDoBOld,
                    Ckh.Date1 as CoBorrowerDoB,
                    Ckh.CardType as CoBorrowerIDType,
                    Ckh.Name6 as CoBorrowerIDNum,
                    Ckh.RelatedName as CoBorrowerRelativeType,
                    ------------------------------
                    ---Loan Information

                    L.AppType,
                    L.BalAmt / @ccydiv AS BalAmt,
                    ------
                    L.AcrIntAmt / @ccydiv as AcrIntAmt_EndPreMonth, ---Accrue interest Ammount until end of prevouse month

                    L.AcrChgAmt / @ccydiv as AcrChgAmt_EndPreMonth,
                    L.AcrPenAmt / @ccydiv as AcrPenAmt_EndPreMonth,
                    L.AcrIntODuePriAmt / @ccydiv as AcrIntODuePriAmt,
                    ---
                    L.IntBalAmt / @ccydiv as IntBalAmt,
                    L.PenBalAmt / @ccydiv as PenBalAmt,
                    L.StopIntTF,
                    L.TrnSeq,
                    ----
                    Case

                        WHEN L.OduePriAmt < 0 THEN 0 ELSE L.OduePriAmt / @ccydiv

                    END AS OduePriAmt,
                    L.OdueIntAmt / @ccydiv AS OdueIntAmt,
                    L.StopAcrIntTF,
                    L.ReschedSeq,
                    L.CcyType,
                    dbo.[GetFirstDisbDate](l.acc) as DisbDate,
                    L.GrantedAmt / @ccydiv AS GrantedAmt,
                    L.IntRate,
                    L.IntEffDate,
                    L.MatDate,
                    L.GLCode,
                    L.GLCodeOrig,
                    L.InstNo,
                    L.FreqType,
		            case 

                        when l.FreqType = '012' then 'Monthly'--monthly

                        when l.FreqType = '026' then '2 Week'--every 2 week

                        when l.FreqType = '052' then 'Weekly'--weekly

                        when l.FreqType = '013' then '4 Week'--every 4 week
			            else ''-- - ************protect in case we add some more product->then we can saw error.
                    end as [PaymentFrequency], --16

                    L.PrType,
                    P.FullDesc as PrName,
                    L.LNCode1,
                    L.LNCode2,
                    L.LNCode3,
                    L.LNCode4,
                    L.LNCode5,
                    L.LNCode6,
                    L.LNCode7,
                    L.LNCode8,
                    Dbo.GetTermOfLoan(L.Acc) AS TERMS,
                    -----------
                    CAST('01/01/1900' AS DATETIME) as LastPaidDate,
                    CAST(0 AS NUMERIC(18, 2)) as LastPaidAmt,
                    cast('' as nvarchar(15)) as PreAcc,
                    CAST(0 AS NUMERIC(18, 2)) as PreDisbAmt,
                    Cast('01/01/1900' AS DATETIME) as NextPaidDate,
                    cast('01/01/1900' as datetime) as NextDueWorkDate,
                    Cast(0 as numeric(18, 3)) as NextPaidAmt,
                    Cast(0 as numeric(18, 3)) as IntRemainInSchedule,
                    cast(0 as numeric(18, 3)) as RemainLoanTerm,
                    cast('01/01/1900' as datetime) as LastPreDueDateOfReportDate,
                    cast(0 as numeric(18, 3)) as BalAmtLastDueOfReportDate,
                    --cast(0 as numeric(18, 3)) as IntAmtAfterPreDueDateOfReportDate,
                    --cast(0 as numeric(18, 3)) as TotalAmtRequiredToClose,
                    cast(0 as numeric(18, 3)) as PrinAmtLastDueDateCollect,
                    --cast('01/01/1900' as datetime) as LastTransDateDueDate,
                    -----------
                    dbo.GetAgeOfLoan(@REPORTDATE, l.Acc) as AgeofLoan,
                    cast(0 as int) as NoInstLate,
                    CAST('01/01/1900' AS DATETIME) AS FirstDayLate,
                    --CAST('' as nvarchar(10)) as StatusReNewLoan,
                    -------------------------------------------------------------------------
                    ---Collection--OK

                    CAST(0 AS NUMERIC(18, 2)) as TrnIntAmt_CollYear, --reserved for collection
 
                     CAST(0 AS NUMERIC(18, 2)) as TrnPriAmt_CollYear, --reserved for collection
  
                      CAST(0 AS NUMERIC(18, 2)) as TrnPenAmt_CollYear, --reserved for collection
   
                       CAST(0 AS NUMERIC(18, 2)) as TrnChgAmt_CollYear, --reserved for collection
   
                       CAST(0 AS NUMERIC(18, 2)) as TrnIntAmt_CollAsOfMonth, --reserved for collection
    
                        CAST(0 AS NUMERIC(18, 2)) as TrnPriAmt_CollAsOfMonth, --reserved for collection
    
                        CAST(0 AS NUMERIC(18, 2)) as TrnPenAmt_CollAsOfMonth, --reserved for collection
     
                         CAST(0 AS NUMERIC(18, 2)) as TrnChgAmt_CollAsOfMonth, --reserved for collection
     
                         CAST(0 AS NUMERIC(18, 2)) as TrnIntAmt_CollDaily, --reserved for collection
      
                          CAST(0 AS NUMERIC(18, 2)) as TrnPriAmt_CollDaily, --reserved for collection
       
                           CAST(0 AS NUMERIC(18, 2)) as TrnPenAmt_CollDaily, --reserved for collection
        
                            CAST(0 AS NUMERIC(18, 2)) as TrnChgAmt_CollDaily, --reserved for collection
                            ------------------------------------------------------------------------------
                            -- - Interesr Culculation Procedure--OK
        
                            CAST(0 AS NUMERIC(18, 2)) as IntOdueAmt_FromIntCalc_Proc,
                            CAST(0 AS NUMERIC(18, 2)) as IntToDateAmt_FromIntCalc_Proc, --= @IntToDate = @IntOdueAmt(ARI from begining month To Date) + @AcrIntAmt(ARI end of Previous month that get from LNACC )
        
                            CAST(0 AS NUMERIC(18, 2)) as PenaltyAmtDue_FromIntCalc_Proc,
                            CAST(0 AS NUMERIC(18, 2)) as TaxAmtDue_FromIntCalc_Proc,
                            CAST(0 AS NUMERIC(18, 2)) as TrnChagAmtDue_FromIntCalc_Proc,
                            ------------------------------------------------------------------------------
                            ----PaidOff, paid off can be normal paid off(loan Past Maturity then paid off) or loan that paid off before paturity
        
                            CAST(0 AS NUMERIC(18, 2)) as TrnIntAmt_PaidOff, --reserved for collection
         
                             CAST(0 AS NUMERIC(18, 2)) as TrnPriAmt_PaidOff, --reserved for collection
          
                              CAST(0 AS NUMERIC(18, 2)) as TrnPenAmt_PaidOff, --reserved for collection
           
                               CAST(0 AS NUMERIC(18, 2)) as TrnChgAmt_PaidOff, --reserved for collection
           
                               CAST('01/01/1900' AS DATETIME) as PaidOffDate,
                               CAST('' AS NVARCHAR(30)) AS STATUS_RENEW,
                               ----
                               --Provision code
           
                               ShortCode =
           		            case 

                                   when dbo.GetTermOfLoan(L.acc) >= (select max(TERM) from AmretLoanProvision) then 1
			            else 0

                    end,
		            MasterField = cast('00000' as nvarchar(10)),
		            -----------------------------------------------------------------
                    L.AccStatus,
		            L.AccStatusDate,
		            BusType.FullDesc BusType,
                    CASE

                        WHEN L.LNCode1 BETWEEN '100' AND '199' THEN 'Agricultural'

                        WHEN L.LNCode1 BETWEEN '200' AND '399' THEN 'Trading & Commerce'

                        WHEN L.LNCode1 BETWEEN '400' AND '499' THEN 'Service'

                        WHEN L.LNCode1 BETWEEN '500' AND '599' THEN 'Transportation'

                        WHEN L.LNCode1 BETWEEN '600' AND '699' THEN 'Construction'

                        WHEN L.LNCode1 > '700' THEN 'Consumption (Householg/Family)'

                    END AS BusSector,
		            Colateral.FullDesc as LoanColateral,
		            LNCycle.FullDesc as LoanCycle,
		            /*
		            guaranter information
		            */
		            R.CID as idGuarenter,
		            ltrim(rtrim(GC.DisplayName)) as GuarenterName,
		            ltrim(rtrim(GCKh.Name1)) as GuarenterNameKh,		
		            case 
			            when GC.GenderType = '001' then 'M'

                        when GC.GenderType = '002' then 'F'
			            else	''

                    end as GuarenterGender,
		            case 
			            when GC.CivilStatusCode = '00D' then 'D'

                        when GC.CivilStatusCode = '00M' then 'M'

                        when GC.CivilStatusCode = '00S' then 'S'

                        when GC.CivilStatusCode = '00W' then 'W'
			            else ''

                    end as GuarenterMaritalStatus,
		            case 
			            when left(GC.Nid,1) = 'M'  then 'M'--M = Maried Letter
                          when left(GC.Nid, 1) = 'C' then 'C'--C = Civil Servan
                           when left(GC.Nid, 1) = 'F' then 'F'

                        when left(GC.Nid,1) = 'B' then 'B'

                        when left(GC.Nid,1) = 'N' and ltrim(rtrim(c.nid)) <> 'N/A' then 'N'

                        when left(GC.Nid,1)= 'D' then 'D'

                        when left(GC.Nid,1)= 'P' then 'P'

                        when left(GC.Nid,1) = 'R' then 'R'
			            else 'Unknown'

                    end as Guarenter_IDType_1,--N,F,D,B,G,R...columns 2 in CBC
		            case 
			            when left(GC.Nid,1) in ('N', 'F', 'R', 'D', 'B', 'C', 'P', 'M') and ltrim(rtrim(GC.nid)) <> 'N/A' then right(ltrim(rtrim(GC.Nid)),len(GC.Nid) - 1)
			            when left(GC.Nid,1) in ('F', 'G', 'R') then
				            case when CAST(right(ltrim(rtrim(GC.Nid)),len(GC.Nid) - 1) AS int) = 0 then 'N/A'
					            else right(GC.Nid, len(GC.Nid) - 1)

                            end
			            else GC.Nid

                    end as GuarenterIDNumber_1,--3

                    GC.BirthDate as GuarenterDateofBirth,		
		            isnull(ltrim(rtrim(GC.Mobile1)), '') as GuarenterMobile1,
		            isnull(ltrim(rtrim(GC.Mobile2)), '') as GuarenterMobile2,
		            GC.locationCode as GuarenterLocationCode,
		            /*
		            Co Guarenter information
		            */
		            GCkh.Name2 as CoGuarenerName,
		            case 
			            when CKh.Sex = 'Rbus' then 'M'

                        when CKh.Sex = 'RsI' then 'F'
			            else ''

                    end as CoGuarenerGender,
		            GCkh.Date1 as CoGuarenerDoB,
		            GCkh.CardType as CoGuarenerIDType,
		            GCkh.Name6 as CoGuarenerIDNum,
		            GCkh.RelatedName as CoGuarenerRelativeType,
		            @REPORTDATE AS REPORTDATE

                FROM LNACC L

                    LEFT JOIN CIF C ON C.CID = substring(L.ACC, 4, 6) and c.type = '001'--Client
                      LEFT JOIN relcid G ON G.relatedcid = C.CID AND G.TYPE = '900'--MEMBER TO GROUP
                    LEFT JOIN CIF G2 ON G2.CID = G.CID--JOIN TO GET GROUP INFORMATION
                  LEFT JOIN RELCID CO ON CO.RelatedCID = G.CID    AND CO.TYPE = '499'--GROUP TO CO
                  LEFT JOIN CIF CO2 ON CO2.CID = CO.CID

                    LEFT JOIN PRPARMS P ON P.PrType = L.PrType and P.AppType = 4

                    LEFT JOIN VTSDCIF CKh on Ckh.CID = c.CID and ckh.Type = '001'--join to get client name & co - borrower in khmer
                      LEFT JOIN VTSDVillage VIL ON VIL.Code = C.LocationCode

                    LEFT JOIN VTSDCommune COM ON COM.CODE = VIL.CodeOfCommune

                    LEFT JOIN VTSDDistrict DIS ON DIS.Code = COM.CodeOfDistrict

                    LEFT JOIN VTSDProvince PRV ON PRV.Code = DIS.CodeOfProvince

                    LEFT JOIN VTSDCIF CO_Kh ON CO_Kh.CID = co2.CID and CO_Kh.Type = '499'

                    inner join Trnhist T ON T.ACC = L.ACC

                    LEFT JOIN USERLOOKUP BusType ON BusType.LookUpCode = L.LNCode1 and BusType.LookUpId = '41'

                    LEFT JOIN USERLOOKUP Colateral ON Colateral.LookUpCode = L.LNCode2 and Colateral.LookUpId = '42'

                    LEFT JOIN USERLOOKUP LNCycle on LNCycle.LookUpCode = L.LNCode3 and LNCycle.LookUpId = '43'

                    LEFT JOIN USERLOOKUP CIF_FamilyMember ON CIF_FamilyMember.LookUpCode = C.CIFCode1 and CIF_FamilyMember.LookUpId = '61'--62

                    LEFT JOIN USERLOOKUP CIF_Income ON CIF_Income.LookUpCode = C.CIFCode2 and CIF_Income.LookUpId = '62'

                    LEFT JOIN USERLOOKUP CIF_Occupation ON CIF_Occupation.LookUpCode = C.CIFCode3 and CIF_Occupation.LookUpId = '63'

                    LEFT JOIN USERLOOKUP CIF_TotalAsset ON CIF_TotalAsset.LookUpCode = C.CIFCode4 and CIF_TotalAsset.LookUpId = '64'

                    LEFT JOIN relacc R on R.Acc + R.Chd = L.Acc + L.Chd and R.Type = '030' and R.AppType = '4'--type = 30 for guarentee
                    LEFT JOIN CIF GC ON GC.CID = R.CID--guaretee

                    LEFT JOIN VTSDCIF GCKh on GCKh.CID = GC.CID and gckh.Type = '001'--join to get client name & co - borrower in khmer

                WHERE

                    L.AccStatus = '99'--CLOSE LOAN IN MONTH ONLY

                    and T.ValueDate between @BeginDateMonth and @REPORTDATE
                    --and T.ValueDate between @BeginDateYear and @REPORTDATE

                    and ISNULL(T.TrnPriAmt, 0) + isnull(T.TrnIntAmt, 0) + isnull(T.TrnPenAmt, 0) + ISNULL(T.TrnChgAmt, 0) > 0

                    and T.BalAmt = 0-- - If loan pay many time i

            -- =============================================================================================================================================================
              --Update Client No of installment Late
               UPDATE #LN_DETAIL
		            SET NoInstLate = (SELECT COUNT(*) FROM LNINST I WHERE I.Status = '1' AND I.Acc = L.Acc),
		             FirstDayLate = (SELECT TOP 1 DueDate FROM LNINST I WHERE I.Status = '1' AND I.Acc = L.Acc ORDER BY DueDate)
               FROM #LN_DETAIL L
               WHERE L.AgeofLoan <> 0

               --Update Status Re - New Loan TO CLOSE LOAN IN MONTH
               UPDATE #LN_DETAIL
		            SET STATUS_RENEW =
                        CASE

                            WHEN(SELECT TOP 1 ISNULL(L1.IdClient, 0)  FROM #LN_DETAIL L1 WHERE L1.IdClient = L.IdClient AND L1.ACCSTATUS <>'99') > 0 then 'RENEW' 
					            ELSE 'NOTRENEW'

                            END
               FROM #LN_DETAIL  L
               WHERE L.accstatus = '99'
              -- ============================================================================================================================================================
            --/*
            --Get Repayment from Transaction history
            --*/
               SELECT

                      LTRIM(RTRIM(a.Acc)) AS Acc,
                      SUM(t.TrnIntAmt / @ccydiv)AS TrnIntAmt_CollAsOfMonth,
                      SUM(t.TrnPriAmt / @ccydiv) AS TrnPriAmt_CollAsOfMonth,
                      SUM(t.TrnPenAmt / @ccydiv) AS TrnPenAmt_CollAsOfMonth,
                      SUM(t.TrnChgAmt / @ccydiv) AS TrnChgAmt_CollAsOfMonth,
                      CAST(0 AS NUMERIC(18, 2)) as TrnIntAmt_CollDaily,
                      CAST(0 AS NUMERIC(18, 2)) as TrnPriAmt_CollDaily,
                      CAST(0 AS NUMERIC(18, 2)) as TrnPenAmt_CollDaily,
                      CAST(0 AS NUMERIC(18, 2)) as TrnChgAmt_CollDaily
               INTO #RepayMonth_Temp
               FROM TrnHist T,#LN_DETAIL a 
               WHERE t.acc = a.Acc AND t.TrnType in ('401', '403', '405', '411', '421', '431', '451', '453') AND
                     t.GLCode NOT LIKE 'W%' AND t.CancelledByTrn IS NULL AND
                     (t.ValueDate BETWEEN @BeginDateMonth AND @REPORTDATE)


               GROUP BY a.Acc

               INSERT INTO #RepayMonth_Temp
                SELECT

                      LTRIM(RTRIM(a.Acc)) AS Acc,
                      0 AS TrnIntAmt_CollAsOfMonth,
                      0 AS TrnPriAmt_CollAsOfMonth,
                      0 AS TrnPenAmt_CollAsOfMonth,
                      0 AS TrnChgAmt_CollAsOfMonth,
                      SUM(t.TrnIntAmt / @ccydiv) as TrnIntAmt_CollDaily,
                      SUM(t.TrnPriAmt / @ccydiv) as TrnPriAmt_CollDaily,
                      SUM(t.TrnPenAmt / @ccydiv) as TrnPenAmt_CollDaily,
                      SUM(t.TrnChgAmt / @ccydiv) as TrnChgAmt_CollDaily
               FROM TrnHist T,#LN_DETAIL a 
               WHERE t.acc = a.Acc AND t.TrnType in ('401', '403', '405', '411', '421', '431', '451', '453') AND
                     t.GLCode NOT LIKE 'W%' AND t.CancelledByTrn IS NULL AND
                     (t.ValueDate = @REPORTDATE)
               GROUP BY a.Acc
               ------------------------------------------------------------------------
               INSERT INTO #RepayMonth_Temp
               SELECT LTRIM(RTRIM(a.Acc)) AS Acc,
                      SUM(t.TrnIntAmt / @ccydiv)AS TrnIntAmt_CollAsOfMonth,
                      SUM(t.TrnPriAmt / @ccydiv) AS TrnPriAmt_CollAsOfMonth,
                      SUM(t.TrnPenAmt / @ccydiv) AS TrnPenAmt_CollAsOfMonth,
                      SUM(t.TrnChgAmt / @ccydiv) AS TrnChgAmt_CollAsOfMonth,
                      0 as TrnIntAmt_CollDaily,
                      0 as TrnPriAmt_CollDaily,
                      0 as TrnPenAmt_CollDaily,
                      0 as TrnChgAmt_CollDaily
               FROM TrnDaily T,#LN_DETAIL a 
               WHERE t.acc = a.Acc AND t.TrnType in ('401', '403', '405', '411', '421', '431', '451', '453') AND
                     t.GLCode NOT LIKE 'W%' AND t.CancelledByTrn IS NULL AND
                     (t.ValueDate BETWEEN @BeginDateMonth AND @REPORTDATE)
               GROUP BY a.Acc
               union-- - Current Day
                  SELECT LTRIM(RTRIM(a.Acc)) AS Acc,
                      0 AS TrnIntAmt_CollAsOfMonth,
                      0 AS TrnPriAmt_CollAsOfMonth,
                      0 AS TrnPenAmt_CollAsOfMonth,
                      0 AS TrnChgAmt_CollAsOfMonth,
                      SUM(t.TrnIntAmt / @ccydiv) as TrnIntAmt_CollDaily,
                      SUM(t.TrnPriAmt / @ccydiv) as TrnPriAmt_CollDaily,
                      SUM(t.TrnPenAmt / @ccydiv) as TrnPenAmt_CollDaily,
                      SUM(t.TrnChgAmt / @ccydiv) as TrnChgAmt_CollDaily
               FROM TrnDaily T,#LN_DETAIL a 
               WHERE t.acc = a.Acc AND t.TrnType in ('401', '403', '405', '411', '421', '431', '451', '453') AND
                     t.GLCode NOT LIKE 'W%' AND t.CancelledByTrn IS NULL AND
                     (t.ValueDate = @REPORTDATE)
               GROUP BY a.Acc
               --------Reverse Closing----------------------
               INSERT INTO #RepayMonth_Temp 
               SELECT LTRIM(RTRIM(a.Acc)) AS Acc,
                      (-1) * SUM(t.TrnIntAmt / @ccydiv)AS TrnIntAmt_CollAsOfMonth,
                      (-1) * SUM(t.TrnPriAmt / @ccydiv) AS TrnPriAmt_CollAsOfMonth,
                      (-1) * SUM(t.TrnPenAmt / @ccydiv) AS TrnPenAmt_CollAsOfMonth,
                      (-1) * SUM(t.TrnChgAmt / @ccydiv) AS TrnChgAmt_CollAsOfMonth,
                      0 as TrnIntAmt_CollDaily,
                      0 as TrnPriAmt_CollDaily,
                      0 as TrnPenAmt_CollDaily,
                      0 as TrnChgAmt_CollDaily
               FROM TrnHist T,#LN_DETAIL a 
               WHERE t.acc = a.Acc AND t.TrnType IN('105', '205', '305', '406', '705') AND t.GLCode NOT LIKE 'W%'
                     AND t.CancelledByTrn IS NULL

                     AND(t.ValueDate BETWEEN @BeginDateMonth AND @REPORTDATE)
               GROUP BY a.Acc
               UNION
               SELECT
               LTRIM(RTRIM(a.Acc)) AS Acc,
                      0 AS TrnIntAmt_CollAsOfMonth,
                      0 AS TrnPriAmt_CollAsOfMonth,
                      0 AS TrnPenAmt_CollAsOfMonth,
                      0 AS TrnChgAmt_CollAsOfMonth,
                      (-1) * SUM(t.TrnIntAmt / @ccydiv) as TrnIntAmt_CollDaily,
                      (-1) * SUM(t.TrnPriAmt / @ccydiv) as TrnPriAmt_CollDaily,
                      (-1) * SUM(t.TrnPenAmt / @ccydiv) as TrnPenAmt_CollDaily,
                      (-1) * SUM(t.TrnChgAmt / @ccydiv) as TrnChgAmt_CollDaily
               FROM TrnHist T,#LN_DETAIL a 
               WHERE t.acc = a.Acc AND t.TrnType IN('105', '205', '305', '406', '705') AND t.GLCode NOT LIKE 'W%'
                     AND t.CancelledByTrn IS NULL

                     AND(t.ValueDate = @REPORTDATE)
               GROUP BY a.Acc


               INSERT INTO #RepayMonth_Temp 
               SELECT LTRIM(RTRIM(a.Acc)) AS Acc,
                      (-1) * SUM(t.TrnIntAmt / @ccydiv)AS TrnIntAmt_CollAsOfMonth,
                      (-1) * SUM(t.TrnPriAmt / @ccydiv) AS TrnPriAmt_CollAsOfMonth,
                      (-1) * SUM(t.TrnPenAmt / @ccydiv) AS TrnPenAmt_CollAsOfMonth,
                      (-1) * SUM(t.TrnChgAmt / @ccydiv) AS TrnChgAmt_CollAsOfMonth,
                      0 as TrnIntAmt_CollDaily,
                      0 as TrnPriAmt_CollDaily,
                      0 as TrnPenAmt_CollDaily,
                      0 as TrnChgAmt_CollDaily
               FROM TrnDaily T,#LN_DETAIL a 
               WHERE t.acc = a.Acc AND t.TrnType IN('105', '205', '305', '406', '705') AND t.GLCode NOT LIKE 'W%'
                     AND t.CancelledByTrn IS NULL AND(t.ValueDate BETWEEN @BeginDateMonth AND @REPORTDATE)
               GROUP BY a.Acc
                UNION
               SELECT
               LTRIM(RTRIM(a.Acc)) AS Acc,
                      0 AS TrnIntAmt_CollAsOfMonth,
                      0 AS TrnPriAmt_CollAsOfMonth,
                      0 AS TrnPenAmt_CollAsOfMonth,
                      0 AS TrnChgAmt_CollAsOfMonth,
                      (-1) * SUM(t.TrnIntAmt / @ccydiv) as TrnIntAmt_CollDaily,
                      (-1) * SUM(t.TrnPriAmt / @ccydiv) as TrnPriAmt_CollDaily,
                      (-1) * SUM(t.TrnPenAmt / @ccydiv) as TrnPenAmt_CollDaily,
                      (-1) * SUM(t.TrnChgAmt / @ccydiv) as TrnChgAmt_CollDaily
               FROM TrnDaily T,#LN_DETAIL a 
               WHERE t.acc = a.Acc AND t.TrnType IN('105', '205', '305', '406', '705') AND t.GLCode NOT LIKE 'W%'
                     AND t.CancelledByTrn IS NULL AND(t.ValueDate = @REPORTDATE)
               GROUP BY a.Acc

               --Clear Reverse Closing----------------------
               SELECT Acc,
                      SUM(TrnIntAmt_CollAsOfMonth) AS TrnIntAmt_CollAsOfMonth,
                      SUM(TrnPriAmt_CollAsOfMonth) AS TrnPriAmt_CollAsOfMonth,
                      SUM(TrnPenAmt_CollAsOfMonth) AS TrnPenAmt_CollAsOfMonth,
                      SUM(TrnChgAmt_CollAsOfMonth) AS TrnChgAmt_CollAsOfMonth,
                      SUM(TrnIntAmt_CollDaily) AS TrnIntAmt_CollDaily,
                      SUM(TrnPriAmt_CollDaily) AS TrnPriAmt_CollDaily,
                      SUM(TrnPenAmt_CollDaily) AS TrnPenAmt_CollDaily,
                      SUM(TrnChgAmt_CollDaily) AS TrnChgAmt_CollDaily
               INTO #RepayMonth
               FROM #RepayMonth_Temp 
               GROUP BY Acc
            /*
           Part 2
           */
            ----Update all loan repayment----------------------
               UPDATE #LN_DETAIL 
	            SET TrnIntAmt_CollAsOfMonth = a.TrnIntAmt_CollAsOfMonth,
                   TrnPriAmt_CollAsOfMonth = a.TrnPriAmt_CollAsOfMonth,
                   TrnPenAmt_CollAsOfMonth = a.TrnPenAmt_CollAsOfMonth,
                   TrnChgAmt_CollAsOfMonth = a.TrnChgAmt_CollAsOfMonth,
                   TrnIntAmt_CollDaily = a.TrnIntAmt_CollDaily,
                   TrnPriAmt_CollDaily = a.TrnPriAmt_CollDaily,
                   TrnPenAmt_CollDaily = a.TrnPenAmt_CollDaily,
                   TrnChgAmt_CollDaily = a.TrnChgAmt_CollDaily
               FROM #LN_DETAIL l,#RepayMonth a 
               WHERE l.acc = a.acc
            -- =========================================================================================================================================================
            /*
            Culculate Interest, to get InterestToDate for only active loan
            */
            --Create other parameter--

                DECLARE @FutureDays         INT

                DECLARE @PenAmt             NUMERIC(18, 3)

                DECLARE @IntODueAmt         NUMERIC(18, 3)

                DECLARE @TaxAmt             NUMERIC(18, 3)

                DECLARE @TrnChgAmt          NUMERIC(18, 3)

                DECLARE @IntAmt             NUMERIC(18, 3)

                SET @FutureDays = 0

                SET @IntAmt = 0

                SET @PenAmt = 0

                SET @IntODueAmt = 0

                SET @TaxAmt = 0

                SET @TrnChgAmt = 0

                --Exec   sp_LNCalcInterest  @idAccount, @FutureDays, @IntAmt Output, @PenAmt Output, @IntODueAmt Output, @TaxAmt Output, @TrnChgAmt Output


                Create table #LNCalcInterest
	            (
                    Acc_Full nvarchar(15),
                    IntOdueAmt numeric(18, 3),
                    IntToDateAmt numeric(18, 3), --= @IntToDate = @IntOdueAmt(ARI from begining month To Date) + @AcrIntAmt(ARI end of Previous month that get from LNACC )

                    PenaltyAmt numeric(18, 3), --Penalty Amount Due

                    TaxAmt numeric(18, 3),
                    TrnChagAmt numeric(18, 3)
                )
                /*
	            loop from each row to culculate interest for each acc
	            */
                declare @AcrIntAmt_EndPreMonth numeric(18, 3)

                DECLARE @ACC_FULL NVARCHAR(15)

                DECLARE MyCursor CURSOR FOR

                    select

                        L.Acc + L.Chd,
                        L.AcrIntAmt_EndPreMonth

                    from #LN_DETAIL L
		            where L.AccStatus between '11' and '98'

                    OPEN    MyCursor

                    FETCH   MyCursor INTO @ACC_FULL, @AcrIntAmt_EndPreMonth

                    WHILE @@FETCH_STATUS <> -1

                    BEGIN

                        Exec   sp_LNCalcInterest  @ACC_FULL, @FutureDays, @IntAmt Output, @PenAmt Output, @IntODueAmt Output, @TaxAmt Output, @TrnChgAmt Output

                        INSERT INTO #LNCalcInterest (Acc_Full,IntOdueAmt,IntToDateAmt,PenaltyAmt,TaxAmt,TrnChagAmt)
			            SELECT

                            @ACC_FULL,
                            @IntODueAmt,
                            @AcrIntAmt_EndPreMonth + @IntAmt, --or @AcrIntAmt_EndPreMonth + @IntODueAmt,
                            @PenAmt,
                            @TaxAmt,
                            @TrnChgAmt

                        FETCH NEXT FROM MyCursor INTO @ACC_FULL, @AcrIntAmt_EndPreMonth

                    END

                CLOSE MyCursor

                DEALLOCATE  MyCursor
                -- - update ineterest calculation to temp table of laon(only active)

                UPDATE #LN_DETAIL 
	            SET

                    IntOdueAmt_FromIntCalc_Proc = c.IntOdueAmt,
                    IntToDateAmt_FromIntCalc_Proc = c.IntToDateAmt, --= @IntToDate = @IntOdueAmt(ARI from begining month To Date) + @AcrIntAmt(ARI end of Previous month that get from LNACC )

                    PenaltyAmtDue_FromIntCalc_Proc = c.PenaltyAmt,
                    TaxAmtDue_FromIntCalc_Proc = c.TaxAmt,
                    TrnChagAmtDue_FromIntCalc_Proc = c.TrnChagAmt
               FROM #LN_DETAIL L,#LNCalcInterest c
               WHERE l.acc + l.Chd = c.Acc_Full  and L.AccStatus between '11' and '98'

            -- =========================================================================================================================================================
            /*
            Paid Of Loan during this months.
            */
            select

                cast(co2.CID as int) as idCO,
                co2.DisplayName as CoName,
                L.acc,
                dbo.[GetFirstDisbDate](L.acc) as DisbDate,
                L.PrType,
                L.GrantedAmt,
                L.AccStatus,
                T.TrnPriAmt,
                T.TrnIntAmt,
                T.TrnPenAmt,
                T.TrnChgAmt,
                T.BalAmt,
                t.TrnDate,
                T.TrnType,
                T.TrnDesc,
                T.ValueDate,
                L.MatDate,
                @REPORTDATE as ReportDate
             into #PaidOff
             from LNACC L

                left
             join CIF c on c.CID = SUBSTRING(l.Acc, 4, 6)

        left
             join RELCID g on g.RelatedCID = c.CID and g.Type = '900'--MEMBER TO GROUP

                LEFT JOIN RELCID CO ON CO.RelatedCID = G.CID    AND CO.TYPE = '499'--GROUP TO CO

                LEFT JOIN CIF CO2 ON CO2.CID = CO.CID--Join to get CO Name

                inner join Trnhist T ON T.ACC = L.ACC

                where L.AccStatus = '99'
                    --and L.AccStatusDate between @BeginDateMonth and @REPORTDATE

                    and T.ValueDate between @BeginDateMonth and @REPORTDATE
                    --and L.MatDate > T.TrnDate

                    and ISNULL(T.TrnPriAmt, 0) + isnull(T.TrnIntAmt, 0) + isnull(T.TrnPenAmt, 0) + ISNULL(T.TrnChgAmt, 0) > 0

                    and T.BalAmt = 0-- - If loan pay many time in this month, only get the last one that paid until balance = 0

            /*
            Add more loan for current day in table TrnDaily, 
            then all transaction not yet moved to TrnHist Table
            */
            insert into #PaidOff
            select

                cast(co2.CID as int) as idCO,
                co2.DisplayName as CoName,
                L.acc,
                dbo.[GetFirstDisbDate](L.acc) as DisbDate,
                L.PrType,
                L.GrantedAmt,
                L.AccStatus,
                T.TrnPriAmt,
                T.TrnIntAmt,
                T.TrnPenAmt,
                T.TrnChgAmt,
                T.BalAmt,
                T.TrnDate,
                T.TrnType,
                T.TrnDesc,
                T.ValueDate,
                L.MatDate,
                @REPORTDATE as ReportDate
             from LNACC L

                left
             join CIF c on c.CID = SUBSTRING(l.Acc, 4, 6)

        left
             join RELCID g on g.RelatedCID = c.CID and g.Type = '900'--MEMBER TO GROUP

                LEFT JOIN RELCID CO ON CO.RelatedCID = G.CID    AND CO.TYPE = '499'--GROUP TO CO

                LEFT JOIN CIF CO2 ON CO2.CID = CO.CID--Join to get CO Name

                inner join TRNDAILY T ON T.ACC = L.ACC

                where L.AccStatus = '99'
                    --and L.AccStatusDate between @BeginDateMonth and @REPORTDATE

                    and T.ValueDate between @BeginDateMonth and @REPORTDATE
                    --and L.MatDate > T.TrnDate

                    and ISNULL(T.TrnPriAmt, 0) + isnull(T.TrnIntAmt, 0) + isnull(T.TrnPenAmt, 0) + ISNULL(T.TrnChgAmt, 0) > 0

                    and T.BalAmt = 0-- -

            update #LN_DETAIL 
	            set

                    TrnPriAmt_PaidOff = P.TrnPriAmt,
                    TrnIntAmt_PaidOff = P.TrnIntAmt,
                    TrnChgAmt_PaidOff = P.TrnChgAmt,
                    TrnPenAmt_PaidOff = P.TrnPenAmt,
                    PaidOffDate = P.ValueDate
            from #LN_DETAIL L1, #PaidOff P 
            WHERE P.Acc = L1.Acc AND L1.AccStatus = '99'
            ------------------------------------------ -
             ---Update provision
            update #LN_DETAIL 
	            set

                MasterField =
		            case 

                        when L1.ShortCode = 0  then
				            case 

                                when L1.AgeofLoan = 0 then '0000'--auto normal loan

                                when L1.AgeofLoan between 1 and  @MaxDayShort  then--@MaxDayShort = 90 days-- > lost
                                    (select

                                        R.code + '0'

                                        from AmretLoanProvision R

                                        where R.term = 0--0 for short term

                                       and R.NumDaysNo = (
                                                           select max(R2.NumDaysNo) from AmretLoanProvision R2

                                                               where R2.Term = 0

                                                               and  R2.NumDaysNo <= L1.AgeofLoan
                                                          )
						            )
					            when L1.AgeofLoan > @MaxDayShort then '0040'--auto loss	
					            else 'UnKnown'

                            end
                        when L1.ShortCode = 1 then
				            case 					
					            when L1.AgeofLoan = 0 then '0001'--normal
                                when L1.AgeofLoan between 1 and @MaxDayLong then--long terms
                                   (
                                   select
                                        R.code +'1'--1 for long term

                               from AmretLoanProvision R

                                   where R.term = @TermDefined-- - @TermDefined = 366 day and up is long term

                                       and R.NumDaysNo = (
                                                           select max(R2.NumDaysNo)

                                                           from AmretLoanProvision R2

                                                               where R2.Term = @TermDefined

                                                                   and R2.NumDaysNo <= L1.AgeofLoan
                                                           )
												            )
					            when L1.AgeofLoan > @MaxDayLong then '0041'--auto loss

                            end
                    end
              from #LN_DETAIL L1
            ----Update LastPaidDate, LastPaid Amt
            update #LN_DETAIL 
	            set LastPaidDate = (select top(1) max(t2.TrnDate) from TRNHIST T2 where t2.Acc = L.Acc and t2.TrnType in ('451', '431', '401')),--'401','403','405','411','421','431','451','453'

                LastPaidAmt = (
                            select isnull(sum(T1.TrnAmt), 0)

                            from TRNHIST T1

                                where
                                T1.Acc = L.Acc

                                and T1.TrnDate = (select top(1) max(T3.TrnDate) from TRNHIST T3 where t3.Acc = L.acc and t3.trnType in ('451', '431', '401'))
					            and T1.trnType in ('451', '431', '401')
				             ),
	            PreDisbAmt = (
                    Select L2.GrantedAmt
                    from LNACC L2

                    where substring(L2.acc,4,6) = L.IdClient

                          and L2.Opendate =
                            (
                                select max(Opendate) as Opendate from LNACC L3 where substring(L3.Acc, 4, 6) = L.IdClient and L3.AccStatus = '99' group by substring(L3.Acc, 4, 6)
				            ) 			  
			              and L2.AccStatus = '99'
			            ),
	            PreAcc = (
                        Select L4.Acc from LNACC L4

                        where L4.Opendate =
                        (
                            select max(L5.Opendate) as Opendate

                            from LNACC L5 where substring(L5.Acc, 4, 6) = L.IdClient and L5.AccStatus = '99' group by SUBSTRING(L5.acc, 4, 6)
			            ) and substring(L4.acc,4,6) = L.IdClient and L4.AccStatus = '99'
		            ),
	            NextPaidDate = (
                    select distinct(Min(I.DueDate)) from LNINST I where I.Acc = L.Acc and I.Chd = L.chd and I.Status = 0 and I.status <> 8
	            ),
	            IntRemainInSchedule = (
                    select

                        sum(I.IntAmt / @CCYDIV)

                    from LNINST I

                    inner
                    join LNACC L6 on L6.Acc = I.Acc and L6.Chd = I.Chd and I.Acc = L.acc and I.Chd = L.Chd

                    inner join RELACC R on R.ACC = L.Acc and R.Chd = L.Chd

                    where I.Status <> 8 and R.AppType = '4' and r.Type = '010'
		            ),
		            LastPreDueDateOfReportDate =
                    (
                        select distinct(max(T1.DueDate)) from LNINST T1 where t1.Acc = L.Acc and t1.Status <> 8

                                and T1.DueDate <= @ReportDate
		            )

            from #LN_DETAIL L 

            -- - Protect loan disburst but never paid
              update #LN_DETAIL 
	            set LastPreDueDateOfReportDate =
                    (
			            case 

                            when LastPreDueDateOfReportDate is null then L3.DisbDate
				            else LastPreDueDateOfReportDate
                        end
		            )
            from #LN_DETAIL L3

            update #LN_DETAIL 
	            set NextPaidAmt = (

                                select
						            case 
							            when L2.NextPaidDate is null then null
							            else isnull(I.PriAmt / @CCYDIV, 0) + isnull(I.IntAmt / @ccydiv, 0) + isnull(I.ChargesAmt / @ccydiv, 0)

                                    end
                                from LNINST I

                                    inner join LNACC L6 on L6.Acc = I.Acc and L6.Chd = I.Chd and I.Acc = L2.acc and I.Chd = L2.Chd

                                    inner join RELACC R on R.ACC = L2.Acc and R.Chd = L2.Chd

                                where--I.Status <> 8-- - reschedule

                                I.Status = 0

                                and R.AppType = '4' and r.Type = '010'

                                and i.DueDate = L2.NextPaidDate
						            ),
		            NextDueWorkDate = dbo.GetWorkingDueDate(NextPaidDate),	
	                BalAmtLastDueOfReportDate = (
                        select top(1)
				            case 
					            when(I3.BFBalAmt / @CcyDiv - I3.OrigPriAmt / @CcyDiv = 0) then I3.BFBalAmt / @CcyDiv

                                when(I3.BFBalAmt / @CcyDiv - I3.OrigPriAmt / @CcyDiv > 0) then(I3.BFBalAmt / @CcyDiv - I3.OrigPriAmt / @CcyDiv)

                            end as AA

                        from LNINST I3

                            where I3.Acc = L2.Acc  and I3.Status <> '8'

                            and I3.DueDate = LastPreDueDateOfReportDate
			            ),
		            PrinAmtLastDueDateCollect = (
                        select

                            top(1) I4.OrigPriAmt / @ccydiv

                        from LNINST I4

                            where acc = L2.Acc

                            and I4.DueDate = L2.LastPreDueDateOfReportDate

                            and I4.Status <> '8'
		            ),
		            RemainLoanTerm = (
                        select

                            count(distinct(I5.DueDate))

                        from LNINST I5

                            inner
                        join LNACC L6 on L6.Acc = I5.Acc and L6.Chd = I5.Chd and I5.Acc = L2.acc and I5.Chd = L2.Chd

                            inner join RELACC R on R.ACC = L2.Acc and R.Chd = L2.Chd

                                where--I.Status <> 8-- - reschedule

                                    I5.Status = 0

                                    and R.AppType = '4' and r.Type = '010'

                                    and I5.PaidDate is null
                                --and i.DueDate = L2.NextPaidDate
		            )
            from #LN_DETAIL L2 
	


            -- ==========================================================================================================================================================
            ---Replace string ',' with ';' for CSV
            declare @ColumnName nvarchar(50)
            declare xCursor cursor for

                SELECT

                    distinct(sc.name)
                --st.name as type_name

                --sc.max_length

                FROM tempdb.sys.columns sc inner join sys.types st on st.system_type_id = sc.system_type_id

                    WHERE[object_id] = OBJECT_ID('tempdb..#LN_DETAIL')

                    and st.name in ('char', 'nchar', 'ntext', 'nvarchar', 'text', 'varchar')
            OPEN    xCursor
            FETCH   xCursor INTO @ColumnName
            WHILE @@FETCH_STATUS <> -1
            BEGIN

                PRINT @ColumnName

                EXECUTE(' UPDATE #LN_DETAIL 

                SET  '  + @ColumnName + ' = REPLACE('  + @ColumnName + ', '', '', '+'''; '')' )

                FETCH NEXT FROM xCursor INTO @ColumnName
            END
            CLOSE xCursor
            DEALLOCATE xCursor


            --select* into[DailyReports].dbo.LN_DETAIL from #LN_DETAIL
            select *,@BrCode as BrCode,@BrShort as BrShort,@BrName as BrName from #LN_DETAIL where accstatus between '11' and '98'


            --select* into LN_DETAIL from #LN_DETAIL

            TRUNCATE TABLE #LN_DETAIL
            truncate table #PaidOff
            TRUNCATE TABLE #LNCalcInterest

            --DROP TABLE #LN_DETAIL
            drop table #PaidOff	
            drop table #RepayMonth_Temp
            drop table #RepayMonth
            DROP TABLE #LNCalcInterest";
                        DataTable dtreports = new DataTable();
                        dtreports = mutility.dtOnlyOneBranch(Sql, rs.BrCode);
                        ViewBag.CurrRunDate = CurrRunDate.CurrRunDate(userkey);                        
                        downloadLoanStatement = dtreports;
                        BrCode = rs.BrCode;
                        if (rs.download == "download")
                        {
                            ReportName = "GL Collateral Reports";
                            BrName = Convert.ToString(mutility.dbSingleResult("select BrLetter from BRANCH_LISTS where flag=1 and BrCode='" + rs.BrCode + "'"));
                            ExportDataToExcel(downloadLoanStatement, BrName + "_" + ReportName);
                            return View("Index");
                        }
                    }
                    else
                    {

                        
                     

                        string dtExcute = "";
                        DataTable dtAffterEx = new DataTable();

                        string ReExcuteName = "select StrSql from Table_Reports_temp where Report_Code='" + rs.ReportCode.Trim() + "'";
                        dtExcute = Convert.ToString(mutility.dbSingleResult(ReExcuteName));
                        string Run = dtExcute;

                        //Get ReportName
                        string ValueString = "select Name from Table_Reports_temp where Report_Code = '" + rs.ReportCode.Replace(" ", "") + "'";
                        ReportName = Convert.ToString(mutility.dbSingleResult(ValueString));
                        //Get Branch Name

                        BrName = Convert.ToString(mutility.dbSingleResult("select BrLetter from BRANCH_LISTS where flag=1 and BrCode='" + rs.BrCode + "'"));


                       


                        DataTable dtstatus = new DataTable();
                        string BrZone_loop = Convert.ToString(mutility.dbSingleResult("select branch_zone from users where user_key='S001'"));
                        var zone_array_loop = BrZone_loop.Split(',');
                        string zonetoarray_loop = string.Join("','", zone_array_loop.Skip(0));
                        string SQl_loop = "";
                        if (rs.BrCode == "ALL")
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
                                catch (Exception)
                                {
                                }
                            }
                        }
                        DataTable publicdt = new DataTable();
                        for (int i = 0; i < ds.Tables.Count; i++)
                        {
                            publicdt.Merge(ds.Tables[i]);
                        }
                        if (rs.BrCode == "ALL")
                        {
                            BrName = "All_Branch";
                        }

                        ExportDataToExcel(publicdt, BrName + "_" + ReportName);
                        TempData["sms"] = "You are already donwload reports";
                    }
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

        public ActionResult ExecLoanOverdue(Reports re)
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

                    DateTime datestart = re.datestart;
                    DateTime dateend = re.dateend;
                    string All = re.all;
                    string Sql = "";
                    if (All == "on")
                    {
                        Sql = @"
                            SET DATEFORMAT DMY
                            DECLARE @REPORTDATE DATETIME
                            DECLARE @BeginDateYear datetime
                            DECLARE @BeginDateMonth datetime
                            DECLARE @PreThreeMonthDate datetime
                            DECLARE @ccydiv INT
                            declare @MaxDayShort int
                            declare @MaxDayLong int 
                            declare @TermDefined int
                            -------------------------------------------------

            		         --set @REPORTDATE='30/06/2011';
			                --set @BeginDateYear='01/01/2011';
			                --set @BeginDateMonth='01/06/2011';
			                select 
				                @REPORTDATE = CurrRunDate,
				                @BeginDateYear = convert(datetime,'01/01/'+cast(year(CurrRunDate) as varchar(4)),103),
				                @BeginDateMonth = convert(datetime,'01/' + cast(month(CurrRunDate) as varchar(2))+'/' + cast(year(CurrRunDate) as varchar(4)),103),
				                @PreThreeMonthDate = dateadd(month,3,CurrRunDate)
			                from BRPARMS

                            SET @ccydiv=(SELECT ccydiv FROM ccy)
                            set @TermDefined = ((select top(1) max(M1.Term) from AmretLoanProvision M1))
                            set @MaxDayShort = (select max(NumDaysNo) from AmretLoanProvision where Term = 0 )
                            set @MaxDayLong = (select max(NumDaysNo) from AmretLoanProvision where Term = @TermDefined)
			                declare @BrCode nvarchar(10)
			                declare @BrShort nvarchar(10)
			                declare @BrName nvarchar(50)
			                declare @DbName nvarchar(50)
			                set @DbName = (SELECT DB_NAME())

			                select 
				                @BrCode = SubBranchCode,
				                @BrShort = SubBranchID,
				                @BrName = SubBranchNameLatin
			                from SKP_Brlist 
			                where DBName = (select left(@DbName,len(@DbName)-4))

                            --print '01/' + cast(month(CurrRunDate) as varchar(2))
                            print @BeginDateMonth
                            print @BeginDateYear
                            -------------------------------------------------
                            /*
                            CLEAN TEMP TABLE 
                            */
                            IF OBJECT_ID('tempdb..#LN_DETAIL') IS NOT NULL
                                DROP TABLE #LN_DETAIL
                            IF OBJECT_ID('tempdb..#PaidOff') is not null
	                            drop table #PaidOff
                            IF OBJECT_ID('tempdb..#RepayMonth_Temp') IS NOT NULL
	                            DROP TABLE #RepayMonth_Temp
                            IF OBJECT_ID('tempdb..#RepayMonth') IS NOT NULL
	                            DROP TABLE #RepayMonth
                            IF OBJECT_ID('tempdb..#LNCalcInterest') IS NOT NULL
	                            DROP TABLE #LNCalcInterest

                            SELECT  
		                            L.Acc,
		                            L.Chd,
		                            /*
		                            CO information
		                            */
		                            co2.CID as IdCO,		
		                            ltrim(rtrim(co2.DisplayName)) as CoName,
		                            ltrim(rtrim(CO_Kh.Name1)) as CoNameKh,	
		                            G.CID AS IdGroup,
		                            ltrim(rtrim(g2.DisplayName)) as GroupName,
		                            ----------------------------------------	
		                            c.CID as IdClient,
		                            ltrim(rtrim(c.DisplayName)) as ClientName,
		                            ltrim(rtrim(CKh.Name1)) as ClientNameKh,		
		                            case 
			                            when c.GenderType = '001' then 'M' 
			                            when c.GenderType = '002' then 'F'
			                            else	''
		                            end as Gender,
		                            case 
			                            when c.CivilStatusCode = '00D' then 'D'
			                            when c.CivilStatusCode = '00M' then 'M'
			                            when c.CivilStatusCode = '00S' then 'S'
			                            when c.CivilStatusCode = '00W' then 'W'
			                            else ''
		                            end as [MaritalStatus],
		                            case 
			                            when left(c.Nid,1) = 'M'  then 'M'--M=Maried Letter
			                            when left(c.Nid,1) = 'C' then 'C' --C= Civil Servan 
			                            when left(c.Nid,1) = 'F' then 'F' 
			                            when left(c.Nid,1) = 'B' then 'B' 
			                            when left(c.Nid,1) = 'N' and ltrim(rtrim(c.nid)) <> 'N/A' then 'N' 
			                            when left(c.Nid,1)= 'D' then 'D'
			                            when left(c.Nid,1)= 'P' then 'P'
			                            when left(c.Nid,1) = 'R' then 'R'
			                            else 'Unknown'
		                            end	 as [IDType_1],--N,F,D,B,G,R... columns 2 in CBC
		                            case 
			                            when left(c.Nid,1) in ('N','F','R','D','B','C','P','M') and ltrim(rtrim(c.nid)) <> 'N/A' then  right(ltrim(rtrim(c.Nid)),len(c.Nid)-1)
			                            when  left(c.Nid,1) in ('F','G','R') then
				                            case when CAST(right(ltrim(rtrim(c.Nid)),len(c.Nid)-1) AS int) = 0 then 'N/A' 
					                            else right(c.Nid,len(c.Nid)-1)
				                            end
			                            else c.Nid
		                            end  as [IDNumber_1],--3
		                            c.BirthDate	as [DateofBirth],		
		                            isnull(ltrim(rtrim(c.Mobile1)),'') as Mobile1,
		                            isnull(ltrim(rtrim(c.Mobile2)),'') as Mobile2,		
		                            ltrim(rtrim(c.CIFCode1)) as CIFCode1,
		                            ltrim(rtrim(c.CIFCode2)) as CIFCode2,
		                            ltrim(rtrim(c.CIFCode3)) as CIFCode3,
		                            ltrim(rtrim(c.CIFCode4)) as CIFCode4,
		                            ltrim(rtrim(c.CIFCode5)) as CIFCode5,			
		                            ltrim(rtrim(c.CIFCode6)) as CIFCode6,
		                            ltrim(rtrim(c.CIFCode7)) as CIFCode7,
		                            ltrim(rtrim(c.CIFCode8)) as CIFCode8,
		                            ltrim(rtrim(c.CIFCode9)) as CIFCode9,
		                            CIF_FamilyMember.FullDesc as FamilyMemberDesc,
		                            CIF_Income.FullDesc as CIF_IncomeDesc,
		                            CIF_Occupation.FullDesc as CIF_OccupationDesc,		
		                            CIF_TotalAsset.FullDesc as CIF_TotalAssetDesc,
		                            c.LocationCode,
		                            VIL.NameLatin AS Village,
		                            COM.NameLatin as Commune,
		                            Dis.NameLatin as District,
		                            PRV.NameLatin as Province,
		                            Vil.NameKhmer + ';' + Com.NameKhmer + ';' + dis.NameKhmer + ';' + PRV.NameKhmer as AdressKhmer,
		                            -------------------------------
		                            ---Co-Borrowere Information
		                            Ckh.Name2 as CoBorrowerName,
		                            case 
			                            when CKh.Sex = 'Rbus' then 'M' 
			                            when CKh.Sex ='RsI' then 'F'
			                            else ''
		                            end as CoBorrowerGender,
		                            --'' as CoBorrowerDoBOld,
		                            Ckh.Date1 as CoBorrowerDoB,
		                            Ckh.CardType as CoBorrowerIDType,
		                            Ckh.Name6 as CoBorrowerIDNum,
		                            Ckh.RelatedName as CoBorrowerRelativeType,
		                            ------------------------------
		                            ---Loan Information		
		                            L.AppType,
		                            L.BalAmt/ @ccydiv AS BalAmt,
		                            ------
		                            L.AcrIntAmt /@ccydiv as AcrIntAmt_EndPreMonth,---Accrue interest Ammount until end of prevouse month
		                            L.AcrChgAmt /@ccydiv as AcrChgAmt_EndPreMonth,
		                            L.AcrPenAmt /@ccydiv as AcrPenAmt_EndPreMonth,
		                            L.AcrIntODuePriAmt /@ccydiv as AcrIntODuePriAmt,
		                            ------
		                            L.IntBalAmt / @ccydiv as IntBalAmt,
		                            L.PenBalAmt / @ccydiv as PenBalAmt,
		                            L.StopIntTF ,
		                            L.TrnSeq,
		                            -------
		                            Case 
			                            WHEN L.OduePriAmt < 0 THEN 0 ELSE L.OduePriAmt / @ccydiv 
		                            END AS OduePriAmt,
		                            L.OdueIntAmt / @ccydiv AS OdueIntAmt,
		                            L.StopAcrIntTF ,
		                            L.ReschedSeq,
		                            L.CcyType,
		                            dbo.[GetFirstDisbDate](l.acc) as DisbDate,
		                            L.GrantedAmt/ @ccydiv AS GrantedAmt,
		                            L.IntRate,
		                            L.IntEffDate,
		                            L.MatDate,	
		                            L.GLCode,
		                            L.GLCodeOrig,	
		                            L.InstNo,
		                            L.FreqType,
		                            case 
			                            when l.FreqType = '012' then 'Monthly' --monthly
			                            when l.FreqType = '026' then '2 Week' --every 2 week
			                            when l.FreqType = '052' then 'Weekly' --weekly
			                            when l.FreqType = '013' then '4 Week' --every 4 week
			                            else '' ---************protect in case we add some more product->then we can saw error.
		                            end as [PaymentFrequency],--16
		                            L.PrType,
		                            P.FullDesc as PrName,
		                            L.LNCode1,
		                            L.LNCode2,
		                            L.LNCode3,
		                            L.LNCode4,
		                            L.LNCode5,
		                            L.LNCode6,
		                            L.LNCode7,
		                            L.LNCode8,	
		                            Dbo.GetTermOfLoan(L.Acc) AS TERMS,
		                            -----------
		                            CAST('01/01/1900' AS DATETIME) as LastPaidDate,
		                            CAST( 0 AS NUMERIC(18,3)) as LastPaidAmt,
		                            cast('' as nvarchar(15)) as PreAcc,
		                            CAST( 0 AS NUMERIC(18,3)) as PreDisbAmt,
		                            Cast('01/01/1900' AS DATETIME) as NextPaidDate,
		                            cast('01/01/1900' as datetime) as NextDueWorkDate,
		                            Cast(0 as numeric(18,3)) as NextPaidAmt,
		                            Cast(0 as numeric(18,3)) as IntRemainInSchedule,
		                            cast(0 as numeric(18,3)) as RemainLoanTerm,
		                            cast('01/01/1900' as datetime) as LastPreDueDateOfReportDate,
		                            cast(0 as numeric(18,3)) as BalAmtLastDueOfReportDate,
		                            cast(0 as numeric(18,3)) as PrinAmtLastDueDateCollect,
		                            --cast('01/01/1900' as datetime) as LastTransDateDueDate,
		                            --cast(0 as numeric(18,3)) as IntAmtAfterPreDueDateOfReportDate,
		                            --cast(0 as numeric(18,3)) as TotalAmtRequiredToClose,
		
		                            -----------
		                            dbo.GetAgeOfLoan(@REPORTDATE,l.Acc) as AgeofLoan,
		                            cast(0 as int) as NoInstLate,
		                            CAST('01/01/1900' AS DATETIME) AS FirstDayLate,
		                            --CAST('' as nvarchar(10) ) as StatusReNewLoan,
		                            -------------------------------------------------------------------------
		                            ---Collection --OK
		                            CAST( 0 AS NUMERIC(18,2)) as TrnIntAmt_CollYear,--reserved for collection
                                    CAST( 0 AS NUMERIC(18,2)) as TrnPriAmt_CollYear,--reserved for collection
                                    CAST( 0 AS NUMERIC(18,2)) as TrnPenAmt_CollYear,--reserved for collection
                                    CAST( 0 AS NUMERIC(18,2)) as TrnChgAmt_CollYear, --reserved for collection
		                            CAST( 0 AS NUMERIC(18,2)) as TrnIntAmt_CollAsOfMonth,--reserved for collection
                                    CAST( 0 AS NUMERIC(18,2)) as  TrnPriAmt_CollAsOfMonth,--reserved for collection
                                    CAST( 0 AS NUMERIC(18,2)) as TrnPenAmt_CollAsOfMonth,--reserved for collection
                                    CAST( 0 AS NUMERIC(18,2)) as TrnChgAmt_CollAsOfMonth, --reserved for collection
		                            CAST( 0 AS NUMERIC(18,2)) as TrnIntAmt_CollDaily,--reserved for collection
                                    CAST( 0 AS NUMERIC(18,2)) as TrnPriAmt_CollDaily,--reserved for collection
                                    CAST( 0 AS NUMERIC(18,2)) as TrnPenAmt_CollDaily,--reserved for collection
                                    CAST( 0 AS NUMERIC(18,2)) as TrnChgAmt_CollDaily, --reserved for collection
		                            ------------------------------------------------------------------------------
		                            ---Interesr Culculation Procedure--OK
		                            CAST( 0 AS NUMERIC(18,2)) as IntOdueAmt_FromIntCalc_Proc,
		                            CAST( 0 AS NUMERIC(18,2)) as IntToDateAmt_FromIntCalc_Proc,--= @IntToDate = @IntOdueAmt ( ARI from begining month To Date ) + @AcrIntAmt (ARI end of Previous month that get from LNACC )
		                            CAST( 0 AS NUMERIC(18,2)) as PenaltyAmtDue_FromIntCalc_Proc,
		                            CAST( 0 AS NUMERIC(18,2)) as TaxAmtDue_FromIntCalc_Proc,
		                            CAST( 0 AS NUMERIC(18,2)) as TrnChagAmtDue_FromIntCalc_Proc,

		                            ------------------------------------------------------------------------------
		                            ----PaidOff, paid off can be normal paid off ( loan Past Maturity then paid off ) or loan that paid off before paturity
		                            CAST( 0 AS NUMERIC(18,2)) as TrnIntAmt_PaidOff,--reserved for collection
                                    CAST( 0 AS NUMERIC(18,2)) as TrnPriAmt_PaidOff,--reserved for collection
                                    CAST( 0 AS NUMERIC(18,2)) as TrnPenAmt_PaidOff,--reserved for collection
                                    CAST( 0 AS NUMERIC(18,2)) as TrnChgAmt_PaidOff, --reserved for collection
		                            CAST('01/01/1900' AS DATETIME)  as PaidOffDate,
		                            CAST('' AS NVARCHAR(30) ) AS STATUS_RENEW,
		                            -----------------------------------------------------------------
		                            --Provision code
		                            ShortCode = 
		                            case 
			                            when dbo.GetTermOfLoan(L.acc) >= (select max(TERM) from AmretLoanProvision) then 1
			                            else 0 
		                            end,
		                            MasterField = cast('00000' as nvarchar(10)),
		                            ----------------------------------------------------------------
		                            L.AccStatus,
		                            L.AccStatusDate,
		                            BusType.FullDesc BusType,
		                            CASE 
			                            WHEN L.LNCode1 BETWEEN '100' AND '199' THEN 'Agricultural'
			                            WHEN L.LNCode1 BETWEEN '200' AND '399' THEN 'Trading & Commerce'
			                            WHEN L.LNCode1 BETWEEN '400' AND '499' THEN 'Service' 
			                            WHEN L.LNCode1 BETWEEN '500' AND '599' THEN 'Transportation' 
			                            WHEN L.LNCode1 BETWEEN '600' AND '699' THEN 'Construction' 
			                            WHEN L.LNCode1 >  '700' THEN 'Consumption (Householg/Family)'
		                            END AS BusSector,
		                            Colateral.FullDesc as LoanColateral,
		                            LNCycle.FullDesc as LoanCycle,
		                            /*
		                            guaranter information
		                            */
		                            R.CID as idGuarenter,
		                            ltrim(rtrim(GC.DisplayName)) as GuarenterName,
		                            ltrim(rtrim(GCKh.Name1)) as GuarenterNameKh,		
		                            case 
			                            when GC.GenderType = '001' then 'M' 
			                            when GC.GenderType = '002' then 'F'
			                            else	''
		                            end as GuarenterGender,
		                            case 
			                            when GC.CivilStatusCode = '00D' then 'D'
			                            when GC.CivilStatusCode = '00M' then 'M'
			                            when GC.CivilStatusCode = '00S' then 'S'
			                            when GC.CivilStatusCode = '00W' then 'W'
			                            else ''
		                            end as GuarenterMaritalStatus,
		                            case 
			                            when left(GC.Nid,1) = 'M'  then 'M'--M=Maried Letter
			                            when left(GC.Nid,1) = 'C' then 'C' --C= Civil Servan 
			                            when left(GC.Nid,1) = 'F' then 'F' 
			                            when left(GC.Nid,1) = 'B' then 'B' 
			                            when left(GC.Nid,1) = 'N' and ltrim(rtrim(c.nid)) <> 'N/A' then 'N' 
			                            when left(GC.Nid,1)= 'D' then 'D'
			                            when left(GC.Nid,1)= 'P' then 'P'
			                            when left(GC.Nid,1) = 'R' then 'R'
			                            else 'Unknown'
		                            end	 as Guarenter_IDType_1,--N,F,D,B,G,R... columns 2 in CBC
		                            case 
			                            when left(GC.Nid,1) in ('N','F','R','D','B','C','P','M') and ltrim(rtrim(GC.nid)) <> 'N/A' then  right(ltrim(rtrim(GC.Nid)),len(GC.Nid)-1)
			                            when  left(GC.Nid,1) in ('F','G','R') then
				                            case when CAST(right(ltrim(rtrim(GC.Nid)),len(GC.Nid)-1) AS int) = 0 then 'N/A' 
					                            else right(GC.Nid,len(GC.Nid)-1)
				                            end
			                            else GC.Nid
		                            end  as GuarenterIDNumber_1,--3
		                            GC.BirthDate	as GuarenterDateofBirth,		
		                            isnull(ltrim(rtrim(GC.Mobile1)),'') as GuarenterMobile1,
		                            isnull(ltrim(rtrim(GC.Mobile2)),'') as GuarenterMobile2,
		                            GC.locationCode as GuarenterLocationCode,
		                            /*
		                            Co Guarenter information
		                            */
		                            GCkh.Name2 as CoGuarenerName,
		                            case 
			                            when CKh.Sex = 'Rbus' then 'M' 
			                            when CKh.Sex ='RsI' then 'F'
			                            else ''
		                            end as CoGuarenerGender,
		                            GCkh.Date1 as CoGuarenerDoB,
		                            GCkh.CardType as CoGuarenerIDType,
		                            GCkh.Name6 as CoGuarenerIDNum,
		                            GCkh.RelatedName as CoGuarenerRelativeType,

		                            @REPORTDATE AS REPORTDATE
	                            INTO #LN_DETAIL 
	                            FROM LNACC L 
		                            LEFT JOIN CIF C ON C.CID = substring(L.ACC,4,6) and c.type = '001'--Client
		                            LEFT JOIN relcid G ON G.relatedcid = C.CID AND G.TYPE = '900'  --MEMBER TO GROUP
		                            LEFT JOIN CIF G2 ON G2.CID = G.CID  --JOIN TO GET GROUP INFORMATION
		                            LEFT JOIN RELCID CO ON CO.RelatedCID = G.CID  	AND CO.TYPE = '499'--GROUP TO CO
		                            LEFT JOIN CIF CO2 ON CO2.CID = CO.CID 
		                            LEFT JOIN PRPARMS P ON P.PrType = L.PrType and P.AppType = 4
		                            LEFT JOIN VTSDCIF CKh on Ckh.CID = c.CID and ckh.Type = '001'--join to get client name & co-borrower in khmer
		                            LEFT JOIN VTSDVillage VIL ON VIL.Code = C.LocationCode
		                            LEFT JOIN VTSDCommune COM ON COM.CODE = VIL.CodeOfCommune
		                            LEFT JOIN VTSDDistrict DIS ON DIS.Code = COM.CodeOfDistrict
		                            LEFT JOIN VTSDProvince PRV ON PRV.Code = DIS.CodeOfProvince
		                            LEFT JOIN VTSDCIF CO_Kh	ON CO_Kh.CID = co2.CID and CO_Kh.Type = '499'
		                            LEFT JOIN USERLOOKUP BusType ON BusType.LookUpCode = L.LNCode1 and BusType.LookUpId = '41'	
		                            LEFT JOIN USERLOOKUP Colateral ON Colateral.LookUpCode = L.LNCode2 and Colateral.LookUpId = '42'
		                            LEFT JOIN USERLOOKUP LNCycle on  LNCycle.LookUpCode = L.LNCode3 and LNCycle.LookUpId = '43'
		                            LEFT JOIN USERLOOKUP CIF_FamilyMember ON  CIF_FamilyMember.LookUpCode = C.CIFCode1 and CIF_FamilyMember.LookUpId = '61'--62
		                            LEFT JOIN USERLOOKUP CIF_Income ON  CIF_Income.LookUpCode = C.CIFCode2 and CIF_Income.LookUpId = '62'
		                            LEFT JOIN USERLOOKUP CIF_Occupation ON  CIF_Occupation.LookUpCode = C.CIFCode3 and CIF_Occupation.LookUpId = '63'
		                            LEFT JOIN USERLOOKUP CIF_TotalAsset ON  CIF_TotalAsset.LookUpCode = C.CIFCode4 and CIF_TotalAsset.LookUpId = '64'
		                            LEFT JOIN relacc R on R.Acc + R.Chd= L.Acc + L.Chd and R.Type  = '030' and R.AppType = '4'--type = 30 for guarentee
		                            LEFT JOIN CIF GC ON GC.CID = R.CID--guaretee
		                            LEFT JOIN VTSDCIF GCKh on GCKh.CID = GC.CID and gckh.Type = '001'--join to get client name & co-borrower in khmer
		                            --INNER JOIN #LNProvisions PRO ON PRO.Acc  = L.Acc+L.Chd		
	                            WHERE L.AccStatus BETWEEN '11' AND '98'
                            ---==================================================================================================================================================
                            ----FOR CLOSE LOAN IN MONTH

                            INSERT INTO #LN_DETAIL
                            SELECT  
		                            L.Acc,
		                            L.Chd,
		                            co2.CID as IdCO,		
		                            ltrim(rtrim(co2.DisplayName)) as CoName,
		                            ltrim(rtrim(CO_Kh.Name1)) as CoNameKh,	
		                            G.CID AS IdGroup,
		                            ltrim(rtrim(g2.DisplayName)) as GroupName,
		                            ----------------------------------------	
		                            c.CID as IdClient,
		                            ltrim(rtrim(c.DisplayName)) as ClientName,
		                            ltrim(rtrim(CKh.Name1)) as ClientNameKh,		
		                            case 
			                            when c.GenderType = '001' then 'M' 
			                            when c.GenderType = '002' then 'F'
			                            else	''
		                            end as Gender,
		                            case 
			                            when c.CivilStatusCode = '00D' then 'D'
			                            when c.CivilStatusCode = '00M' then 'M'
			                            when c.CivilStatusCode = '00S' then 'S'
			                            when c.CivilStatusCode = '00W' then 'W'
			                            else ''
		                            end as [MaritalStatus],
		                            case 
			                            when left(c.Nid,1) = 'M'  then 'M'--M=Maried Letter
			                            when left(c.Nid,1) = 'C' then 'C' --C= Civil Servan 
			                            when left(c.Nid,1) = 'F' then 'F' 
			                            when left(c.Nid,1) = 'B' then 'B' 
			                            when left(c.Nid,1) = 'N' and ltrim(rtrim(c.nid)) <> 'N/A' then 'N' 
			                            when left(c.Nid,1)= 'D' then 'D'
			                            when left(c.Nid,1)= 'P' then 'P'
			                            when left(c.Nid,1) = 'R' then 'R'
			                            else 'Unknown'
		                            end	 as [IDType_1],--N,F,D,B,G,R... columns 2 in CBC
		                            case 
			                            when left(c.Nid,1) in ('N','F','R','D','B','C','P','M') and ltrim(rtrim(c.nid)) <> 'N/A' then  right(ltrim(rtrim(c.Nid)),len(c.Nid)-1)
			                            when  left(c.Nid,1) in ('F','G','R') then
				                            case when CAST(right(ltrim(rtrim(c.Nid)),len(c.Nid)-1) AS int) = 0 then 'N/A' 
					                            else right(c.Nid,len(c.Nid)-1)
				                            end
			                            else c.Nid
		                            end  as [IDNumber_1],--3
		                            c.BirthDate	as [DateofBirth],		
		                            isnull(ltrim(rtrim(c.Mobile1)),'') as Mobile1,
		                            isnull(ltrim(rtrim(c.Mobile2)),'') as Mobile2,		
		                            ltrim(rtrim(c.CIFCode1)) as CIFCode1,
		                            ltrim(rtrim(c.CIFCode2)) as CIFCode2,
		                            ltrim(rtrim(c.CIFCode3)) as CIFCode3,
		                            ltrim(rtrim(c.CIFCode4)) as CIFCode4,
		                            ltrim(rtrim(c.CIFCode5)) as CIFCode5,			
		                            ltrim(rtrim(c.CIFCode6)) as CIFCode6,
		                            ltrim(rtrim(c.CIFCode7)) as CIFCode7,
		                            ltrim(rtrim(c.CIFCode8)) as CIFCode8,
		                            ltrim(rtrim(c.CIFCode9)) as CIFCode9,
		                            CIF_FamilyMember.FullDesc as FamilyMemberDesc,
		                            CIF_Income.FullDesc as CIF_IncomeDesc,
		                            CIF_Occupation.FullDesc as CIF_OccupationDesc,		
		                            CIF_TotalAsset.FullDesc as CIF_TotalAssetDesc,
		                            c.LocationCode,
		                            VIL.NameLatin AS Village,
		                            COM.NameLatin as Commune,
		                            Dis.NameLatin as District,
		                            PRV.NameLatin as Province,
		                            Vil.NameKhmer + ';' + Com.NameKhmer + ';' + dis.NameKhmer + ';' + PRV.NameKhmer as AdressKhmer,
		                            -------------------------------
		                            ---garanter Information
		                            Ckh.Name2 as CoBorrowerName,
		                            case 
			                            when CKh.Sex = 'Rbus' then 'M' 
			                            when CKh.Sex ='RsI' then 'F'
			                            else ''
		                            end as CoBorrowerGender,
		                            --'' as CoBorrowerDoBOld,
		                            Ckh.Date1 as CoBorrowerDoB,
		                            Ckh.CardType as CoBorrowerIDType,
		                            Ckh.Name6 as CoBorrowerIDNum,
		                            Ckh.RelatedName as CoBorrowerRelativeType,
		                            ------------------------------
		                            ---Loan Information	
		                            L.AppType,
		                            L.BalAmt/ @ccydiv AS BalAmt,
		                            ------
		                            L.AcrIntAmt /@ccydiv as AcrIntAmt_EndPreMonth,---Accrue interest Ammount until end of prevouse month
		                            L.AcrChgAmt /@ccydiv as AcrChgAmt_EndPreMonth,
		                            L.AcrPenAmt /@ccydiv as AcrPenAmt_EndPreMonth,
		                            L.AcrIntODuePriAmt /@ccydiv as AcrIntODuePriAmt,
		                            ---
		                            L.IntBalAmt / @ccydiv as IntBalAmt,
		                            L.PenBalAmt / @ccydiv as PenBalAmt,
		                            L.StopIntTF ,	
		                            L.TrnSeq,
		                            ----
		                            Case 
			                            WHEN L.OduePriAmt < 0 THEN 0 ELSE L.OduePriAmt / @ccydiv 
		                            END AS OduePriAmt,
		                            L.OdueIntAmt / @ccydiv AS OdueIntAmt,
		                            L.StopAcrIntTF ,
		                            L.ReschedSeq,
		                            L.CcyType,
		                            dbo.[GetFirstDisbDate](l.acc) as DisbDate,
		                            L.GrantedAmt/ @ccydiv AS GrantedAmt,
		                            L.IntRate,
		                            L.IntEffDate,
		                            L.MatDate,	
		                            L.GLCode,
		                            L.GLCodeOrig,	
		                            L.InstNo,
		                            L.FreqType,
		                            case 
			                            when l.FreqType = '012' then 'Monthly' --monthly
			                            when l.FreqType = '026' then '2 Week' --every 2 week
			                            when l.FreqType = '052' then 'Weekly' --weekly
			                            when l.FreqType = '013' then '4 Week' --every 4 week
			                            else '' ---************protect in case we add some more product->then we can saw error.
		                            end as [PaymentFrequency],--16
		                            L.PrType,
		                            P.FullDesc as PrName,
		                            L.LNCode1,
		                            L.LNCode2,
		                            L.LNCode3,
		                            L.LNCode4,
		                            L.LNCode5,
		                            L.LNCode6,
		                            L.LNCode7,
		                            L.LNCode8,	
		                            Dbo.GetTermOfLoan(L.Acc) AS TERMS,
		                            -----------
		                            CAST('01/01/1900' AS DATETIME) as LastPaidDate,
		                            CAST( 0 AS NUMERIC(18,2)) as LastPaidAmt,
		                            cast('' as nvarchar(15)) as PreAcc,
		                            CAST( 0 AS NUMERIC(18,2)) as PreDisbAmt,
		                            Cast('01/01/1900' AS DATETIME) as NextPaidDate,
		                            cast('01/01/1900' as datetime) as NextDueWorkDate,
		                            Cast(0 as numeric(18,3)) as NextPaidAmt,
		                            Cast(0 as numeric(18,3)) as IntRemainInSchedule,
		                            cast(0 as numeric(18,3)) as RemainLoanTerm,
		                            cast('01/01/1900' as datetime) as LastPreDueDateOfReportDate,
		                            cast(0 as numeric(18,3)) as BalAmtLastDueOfReportDate,
		                            --cast(0 as numeric(18,3)) as IntAmtAfterPreDueDateOfReportDate,
		                            --cast(0 as numeric(18,3)) as TotalAmtRequiredToClose,
		                            cast(0 as numeric(18,3)) as PrinAmtLastDueDateCollect,
		                            --cast('01/01/1900' as datetime) as LastTransDateDueDate,
		                            -----------
		                            dbo.GetAgeOfLoan(@REPORTDATE,l.Acc) as AgeofLoan,
		                            cast(0 as int) as NoInstLate,
		                            CAST('01/01/1900' AS DATETIME) AS FirstDayLate,
		                            --CAST('' as nvarchar(10) ) as StatusReNewLoan,
		                            -------------------------------------------------------------------------
		                            ---Collection --OK
		                            CAST( 0 AS NUMERIC(18,2)) as TrnIntAmt_CollYear,--reserved for collection
                                    CAST( 0 AS NUMERIC(18,2)) as TrnPriAmt_CollYear,--reserved for collection
                                    CAST( 0 AS NUMERIC(18,2)) as TrnPenAmt_CollYear,--reserved for collection
                                    CAST( 0 AS NUMERIC(18,2)) as TrnChgAmt_CollYear, --reserved for collection
		                            CAST( 0 AS NUMERIC(18,2)) as TrnIntAmt_CollAsOfMonth,--reserved for collection
                                    CAST( 0 AS NUMERIC(18,2)) as  TrnPriAmt_CollAsOfMonth,--reserved for collection
                                    CAST( 0 AS NUMERIC(18,2)) as TrnPenAmt_CollAsOfMonth,--reserved for collection
                                    CAST( 0 AS NUMERIC(18,2)) as TrnChgAmt_CollAsOfMonth, --reserved for collection
		                            CAST( 0 AS NUMERIC(18,2)) as TrnIntAmt_CollDaily,--reserved for collection
                                    CAST( 0 AS NUMERIC(18,2)) as TrnPriAmt_CollDaily,--reserved for collection
                                    CAST( 0 AS NUMERIC(18,2)) as TrnPenAmt_CollDaily,--reserved for collection
                                    CAST( 0 AS NUMERIC(18,2)) as TrnChgAmt_CollDaily, --reserved for collection
		                            ------------------------------------------------------------------------------
		                            ---Interesr Culculation Procedure--OK
		                            CAST( 0 AS NUMERIC(18,2)) as IntOdueAmt_FromIntCalc_Proc,
		                            CAST( 0 AS NUMERIC(18,2)) as IntToDateAmt_FromIntCalc_Proc,--= @IntToDate = @IntOdueAmt ( ARI from begining month To Date ) + @AcrIntAmt (ARI end of Previous month that get from LNACC )
		                            CAST( 0 AS NUMERIC(18,2)) as PenaltyAmtDue_FromIntCalc_Proc,
		                            CAST( 0 AS NUMERIC(18,2)) as TaxAmtDue_FromIntCalc_Proc,
		                            CAST( 0 AS NUMERIC(18,2)) as TrnChagAmtDue_FromIntCalc_Proc,
		                            ------------------------------------------------------------------------------
		                            ----PaidOff, paid off can be normal paid off ( loan Past Maturity then paid off ) or loan that paid off before paturity
		                            CAST( 0 AS NUMERIC(18,2)) as TrnIntAmt_PaidOff,--reserved for collection
                                    CAST( 0 AS NUMERIC(18,2)) as TrnPriAmt_PaidOff,--reserved for collection
                                    CAST( 0 AS NUMERIC(18,2)) as TrnPenAmt_PaidOff,--reserved for collection
                                    CAST( 0 AS NUMERIC(18,2)) as TrnChgAmt_PaidOff, --reserved for collection
		                            CAST('01/01/1900' AS DATETIME)  as PaidOffDate,
		                            CAST('' AS NVARCHAR(30) ) AS STATUS_RENEW,
		                            ----
		                            --Provision code
		                            ShortCode = 
		                            case 
			                            when dbo.GetTermOfLoan(L.acc) >= (select max(TERM) from AmretLoanProvision) then 1
			                            else 0 
		                            end,
		                            MasterField = cast('00000' as nvarchar(10)),
		                            -----------------------------------------------------------------
		                            L.AccStatus,
		                            L.AccStatusDate,
		                            BusType.FullDesc BusType,
		                            CASE 
			                            WHEN L.LNCode1 BETWEEN '100' AND '199' THEN 'Agricultural'
			                            WHEN L.LNCode1 BETWEEN '200' AND '399' THEN 'Trading & Commerce'
			                            WHEN L.LNCode1 BETWEEN '400' AND '499' THEN 'Service' 
			                            WHEN L.LNCode1 BETWEEN '500' AND '599' THEN 'Transportation' 
			                            WHEN L.LNCode1 BETWEEN '600' AND '699' THEN 'Construction' 
			                            WHEN L.LNCode1 >  '700' THEN 'Consumption (Householg/Family)'
		                            END AS BusSector,
		                            Colateral.FullDesc as LoanColateral,
		                            LNCycle.FullDesc as LoanCycle,
		                            /*
		                            guaranter information
		                            */
		                            R.CID as idGuarenter,
		                            ltrim(rtrim(GC.DisplayName)) as GuarenterName,
		                            ltrim(rtrim(GCKh.Name1)) as GuarenterNameKh,		
		                            case 
			                            when GC.GenderType = '001' then 'M' 
			                            when GC.GenderType = '002' then 'F'
			                            else	''
		                            end as GuarenterGender,
		                            case 
			                            when GC.CivilStatusCode = '00D' then 'D'
			                            when GC.CivilStatusCode = '00M' then 'M'
			                            when GC.CivilStatusCode = '00S' then 'S'
			                            when GC.CivilStatusCode = '00W' then 'W'
			                            else ''
		                            end as GuarenterMaritalStatus,
		                            case 
			                            when left(GC.Nid,1) = 'M'  then 'M'--M=Maried Letter
			                            when left(GC.Nid,1) = 'C' then 'C' --C= Civil Servan 
			                            when left(GC.Nid,1) = 'F' then 'F' 
			                            when left(GC.Nid,1) = 'B' then 'B' 
			                            when left(GC.Nid,1) = 'N' and ltrim(rtrim(c.nid)) <> 'N/A' then 'N' 
			                            when left(GC.Nid,1)= 'D' then 'D'
			                            when left(GC.Nid,1)= 'P' then 'P'
			                            when left(GC.Nid,1) = 'R' then 'R'
			                            else 'Unknown'
		                            end	 as Guarenter_IDType_1,--N,F,D,B,G,R... columns 2 in CBC
		                            case 
			                            when left(GC.Nid,1) in ('N','F','R','D','B','C','P','M') and ltrim(rtrim(GC.nid)) <> 'N/A' then  right(ltrim(rtrim(GC.Nid)),len(GC.Nid)-1)
			                            when  left(GC.Nid,1) in ('F','G','R') then
				                            case when CAST(right(ltrim(rtrim(GC.Nid)),len(GC.Nid)-1) AS int) = 0 then 'N/A' 
					                            else right(GC.Nid,len(GC.Nid)-1)
				                            end
			                            else GC.Nid
		                            end  as GuarenterIDNumber_1,--3
		                            GC.BirthDate	as GuarenterDateofBirth,		
		                            isnull(ltrim(rtrim(GC.Mobile1)),'') as GuarenterMobile1,
		                            isnull(ltrim(rtrim(GC.Mobile2)),'') as GuarenterMobile2,
		                            GC.locationCode as GuarenterLocationCode,
		                            /*
		                            Co Guarenter information
		                            */
		                            GCkh.Name2 as CoGuarenerName,
		                            case 
			                            when CKh.Sex = 'Rbus' then 'M' 
			                            when CKh.Sex ='RsI' then 'F'
			                            else ''
		                            end as CoGuarenerGender,
		                            GCkh.Date1 as CoGuarenerDoB,
		                            GCkh.CardType as CoGuarenerIDType,
		                            GCkh.Name6 as CoGuarenerIDNum,
		                            GCkh.RelatedName as CoGuarenerRelativeType,
		                            @REPORTDATE AS REPORTDATE

	                            FROM LNACC L 
		                            LEFT JOIN CIF C ON C.CID = substring(L.ACC,4,6) and c.type = '001'--Client
		                            LEFT JOIN relcid G ON G.relatedcid = C.CID AND G.TYPE = '900'  --MEMBER TO GROUP
		                            LEFT JOIN CIF G2 ON G2.CID = G.CID  --JOIN TO GET GROUP INFORMATION
		                            LEFT JOIN RELCID CO ON CO.RelatedCID = G.CID  	AND CO.TYPE = '499'--GROUP TO CO
		                            LEFT JOIN CIF CO2 ON CO2.CID = CO.CID 
		                            LEFT JOIN PRPARMS P ON P.PrType = L.PrType and P.AppType = 4
		                            LEFT JOIN VTSDCIF CKh on Ckh.CID = c.CID and ckh.Type = '001'--join to get client name & co-borrower in khmer
		                            LEFT JOIN VTSDVillage VIL ON VIL.Code = C.LocationCode
		                            LEFT JOIN VTSDCommune COM ON COM.CODE = VIL.CodeOfCommune
		                            LEFT JOIN VTSDDistrict DIS ON DIS.Code = COM.CodeOfDistrict
		                            LEFT JOIN VTSDProvince PRV ON PRV.Code = DIS.CodeOfProvince
		                            LEFT JOIN VTSDCIF CO_Kh	ON CO_Kh.CID = co2.CID and CO_Kh.Type = '499'
		                            inner join Trnhist T ON T.ACC = L.ACC 	
		                            LEFT JOIN USERLOOKUP BusType ON BusType.LookUpCode = L.LNCode1 and BusType.LookUpId = '41'	
		                            LEFT JOIN USERLOOKUP Colateral ON Colateral.LookUpCode = L.LNCode2 and Colateral.LookUpId = '42'
		                            LEFT JOIN USERLOOKUP LNCycle on  LNCycle.LookUpCode = L.LNCode3 and LNCycle.LookUpId = '43'
		                            LEFT JOIN USERLOOKUP CIF_FamilyMember ON  CIF_FamilyMember.LookUpCode = C.CIFCode1 and CIF_FamilyMember.LookUpId = '61'--62
		                            LEFT JOIN USERLOOKUP CIF_Income ON  CIF_Income.LookUpCode = C.CIFCode2 and CIF_Income.LookUpId = '62'
		                            LEFT JOIN USERLOOKUP CIF_Occupation ON  CIF_Occupation.LookUpCode = C.CIFCode3 and CIF_Occupation.LookUpId = '63'
		                            LEFT JOIN USERLOOKUP CIF_TotalAsset ON  CIF_TotalAsset.LookUpCode = C.CIFCode4 and CIF_TotalAsset.LookUpId = '64'
		                            LEFT JOIN relacc R on R.Acc + R.Chd= L.Acc + L.Chd and R.Type  = '030' and R.AppType = '4'--type = 30 for guarentee
		                            LEFT JOIN CIF GC ON GC.CID = R.CID--guaretee
		                            LEFT JOIN VTSDCIF GCKh on GCKh.CID = GC.CID and gckh.Type = '001'--join to get client name & co-borrower in khmer
	                            WHERE 
		                            L.AccStatus = '99' --CLOSE LOAN IN MONTH ONLY
		                            and T.ValueDate between @BeginDateMonth and @REPORTDATE
		                            --and T.ValueDate between @BeginDateYear and @REPORTDATE
		                            and ISNULL(T.TrnPriAmt,0) +	isnull(T.TrnIntAmt,0) + isnull(T.TrnPenAmt,0) + ISNULL(T.TrnChgAmt,0)  >0
		                            and T.BalAmt = 0 ---If loan pay many time i

                            --=============================================================================================================================================================
                              --Update Client No of installment Late
                               UPDATE #LN_DETAIL
		                            SET NoInstLate = (SELECT COUNT(*) FROM LNINST I WHERE I.Status ='1' AND I.Acc = L.Acc),
		                             FirstDayLate = (SELECT TOP 1 DueDate FROM LNINST I WHERE I.Status ='1' AND I.Acc = L.Acc ORDER BY DueDate)
                               FROM #LN_DETAIL L
                               WHERE L.AgeofLoan <>0  

                               --Update Status Re-New Loan TO CLOSE LOAN IN MONTH
                               UPDATE #LN_DETAIL
		                            SET STATUS_RENEW=
			                            CASE 
				                            WHEN (SELECT TOP 1 ISNULL(L1.IdClient,0)  FROM #LN_DETAIL L1 WHERE L1.IdClient = L.IdClient AND L1.ACCSTATUS <>'99') > 0 then 'RENEW' 
					                            ELSE 'NOTRENEW' 
				                            END
                               FROM #LN_DETAIL  L
                               WHERE L.accstatus='99'
                              --============================================================================================================================================================ 
                            --/*
                            --Get Repayment from Transaction history
                            --*/
                               SELECT 
		                              LTRIM(RTRIM(a.Acc)) AS Acc,
                                      SUM(t.TrnIntAmt/@ccydiv)AS TrnIntAmt_CollAsOfMonth,
                                      SUM(t.TrnPriAmt/@ccydiv) AS TrnPriAmt_CollAsOfMonth,
                                      SUM(t.TrnPenAmt/@ccydiv) AS TrnPenAmt_CollAsOfMonth,
                                      SUM(t.TrnChgAmt/@ccydiv) AS TrnChgAmt_CollAsOfMonth,
		                              CAST( 0 AS NUMERIC(18,2)) as TrnIntAmt_CollDaily,
		                              CAST( 0 AS NUMERIC(18,2)) as TrnPriAmt_CollDaily,
		                              CAST( 0 AS NUMERIC(18,2)) as TrnPenAmt_CollDaily,
		                              CAST( 0 AS NUMERIC(18,2)) as TrnChgAmt_CollDaily  		     
                               INTO #RepayMonth_Temp
                               FROM TrnHist T,#LN_DETAIL a 
                               WHERE t.acc=a.Acc AND t.TrnType in ('401','403','405','411','421','431','451','453') AND 
                                     t.GLCode NOT LIKE 'W%' AND t.CancelledByTrn IS NULL AND 
                                     (t.ValueDate BETWEEN @BeginDateMonth AND @REPORTDATE) 
		 
                               GROUP BY a.Acc

                               INSERT INTO #RepayMonth_Temp
                                SELECT 
		                              LTRIM(RTRIM(a.Acc)) AS Acc, 
                                      0 AS TrnIntAmt_CollAsOfMonth,
                                      0 AS TrnPriAmt_CollAsOfMonth,
                                      0 AS TrnPenAmt_CollAsOfMonth,
                                      0 AS TrnChgAmt_CollAsOfMonth,
		                              SUM(t.TrnIntAmt/@ccydiv) as TrnIntAmt_CollDaily,
		                              SUM(t.TrnPriAmt/@ccydiv) as TrnPriAmt_CollDaily,
		                              SUM(t.TrnPenAmt/@ccydiv) as TrnPenAmt_CollDaily,
		                              SUM(t.TrnChgAmt/@ccydiv) as TrnChgAmt_CollDaily
                               FROM TrnHist T,#LN_DETAIL a 
                               WHERE t.acc=a.Acc AND t.TrnType in ('401','403','405','411','421','431','451','453') AND 
                                     t.GLCode NOT LIKE 'W%' AND t.CancelledByTrn IS NULL AND 
                                     (t.ValueDate = @REPORTDATE) 
                               GROUP BY a.Acc
                               ------------------------------------------------------------------------   
                               INSERT INTO #RepayMonth_Temp
                               SELECT LTRIM(RTRIM(a.Acc)) AS Acc, 
                                      SUM(t.TrnIntAmt/@ccydiv)AS TrnIntAmt_CollAsOfMonth,
                                      SUM(t.TrnPriAmt/@ccydiv) AS TrnPriAmt_CollAsOfMonth,
                                      SUM(t.TrnPenAmt/@ccydiv) AS TrnPenAmt_CollAsOfMonth,
                                      SUM(t.TrnChgAmt/@ccydiv) AS TrnChgAmt_CollAsOfMonth,
		                              0 as TrnIntAmt_CollDaily,
		                              0 as TrnPriAmt_CollDaily,
		                              0 as TrnPenAmt_CollDaily,
		                              0 as TrnChgAmt_CollDaily
                               FROM TrnDaily T,#LN_DETAIL a 
                               WHERE t.acc=a.Acc AND t.TrnType in ('401','403','405','411','421','431','451','453') AND 
                                     t.GLCode NOT LIKE 'W%' AND t.CancelledByTrn IS NULL AND 
                                     (t.ValueDate BETWEEN @BeginDateMonth AND @REPORTDATE) 
                               GROUP BY a.Acc
                               union ---Current Day
                                  SELECT LTRIM(RTRIM(a.Acc)) AS Acc, 
                                      0 AS TrnIntAmt_CollAsOfMonth,
                                      0 AS TrnPriAmt_CollAsOfMonth,
                                      0 AS TrnPenAmt_CollAsOfMonth,
                                      0 AS TrnChgAmt_CollAsOfMonth,
		                              SUM(t.TrnIntAmt/@ccydiv) as TrnIntAmt_CollDaily,
		                              SUM(t.TrnPriAmt/@ccydiv) as TrnPriAmt_CollDaily,
		                              SUM(t.TrnPenAmt/@ccydiv) as TrnPenAmt_CollDaily,
		                              SUM(t.TrnChgAmt/@ccydiv) as TrnChgAmt_CollDaily
                               FROM TrnDaily T,#LN_DETAIL a 
                               WHERE t.acc=a.Acc AND t.TrnType in ('401','403','405','411','421','431','451','453') AND 
                                     t.GLCode NOT LIKE 'W%' AND t.CancelledByTrn IS NULL AND 
                                     (t.ValueDate = @REPORTDATE) 
                               GROUP BY a.Acc
                               --------Reverse Closing ----------------------
                               INSERT INTO #RepayMonth_Temp 
                               SELECT LTRIM(RTRIM(a.Acc)) AS Acc, 
                                      (-1)*SUM(t.TrnIntAmt/@ccydiv)AS TrnIntAmt_CollAsOfMonth,
                                      (-1)*SUM(t.TrnPriAmt/@ccydiv) AS TrnPriAmt_CollAsOfMonth,
                                      (-1)*SUM(t.TrnPenAmt/@ccydiv) AS TrnPenAmt_CollAsOfMonth,
                                      (-1)*SUM(t.TrnChgAmt/@ccydiv) AS TrnChgAmt_CollAsOfMonth,  
		                              0 as TrnIntAmt_CollDaily,
		                              0 as TrnPriAmt_CollDaily,
		                              0 as TrnPenAmt_CollDaily,
		                              0 as TrnChgAmt_CollDaily
                               FROM TrnHist T,#LN_DETAIL a 
                               WHERE t.acc=a.Acc AND t.TrnType IN('105','205','305','406','705') AND t.GLCode NOT LIKE 'W%'
                                     AND t.CancelledByTrn IS NULL 
		                             AND (t.ValueDate BETWEEN @BeginDateMonth AND @REPORTDATE) 
                               GROUP BY a.Acc     
                               UNION
                               SELECT
                               LTRIM(RTRIM(a.Acc)) AS Acc, 
                                      0 AS TrnIntAmt_CollAsOfMonth,
                                      0 AS TrnPriAmt_CollAsOfMonth,
                                      0 AS TrnPenAmt_CollAsOfMonth,
                                      0 AS TrnChgAmt_CollAsOfMonth,  
		                              (-1)*SUM(t.TrnIntAmt/@ccydiv) as TrnIntAmt_CollDaily,
		                              (-1)*SUM(t.TrnPriAmt/@ccydiv) as TrnPriAmt_CollDaily,
		                              (-1)*SUM(t.TrnPenAmt/@ccydiv) as TrnPenAmt_CollDaily,
		                              (-1)*SUM(t.TrnChgAmt/@ccydiv) as TrnChgAmt_CollDaily
                               FROM TrnHist T,#LN_DETAIL a 
                               WHERE t.acc=a.Acc AND t.TrnType IN('105','205','305','406','705') AND t.GLCode NOT LIKE 'W%'
                                     AND t.CancelledByTrn IS NULL 
		                             AND (t.ValueDate = @REPORTDATE) 
                               GROUP BY a.Acc   
      
                               INSERT INTO #RepayMonth_Temp 
                               SELECT LTRIM(RTRIM(a.Acc)) AS Acc, 
                                      (-1)*SUM(t.TrnIntAmt/@ccydiv)AS TrnIntAmt_CollAsOfMonth,
                                      (-1)*SUM(t.TrnPriAmt/@ccydiv) AS TrnPriAmt_CollAsOfMonth,
                                      (-1)*SUM(t.TrnPenAmt/@ccydiv) AS TrnPenAmt_CollAsOfMonth,
                                      (-1)*SUM(t.TrnChgAmt/@ccydiv) AS TrnChgAmt_CollAsOfMonth,
		                              0 as TrnIntAmt_CollDaily,
		                              0 as TrnPriAmt_CollDaily,
		                              0 as TrnPenAmt_CollDaily,
		                              0 as TrnChgAmt_CollDaily 
                               FROM TrnDaily T,#LN_DETAIL a 
                               WHERE t.acc=a.Acc AND t.TrnType IN('105','205','305','406','705') AND t.GLCode NOT LIKE 'W%'
                                     AND t.CancelledByTrn IS NULL AND (t.ValueDate BETWEEN @BeginDateMonth AND @REPORTDATE) 
                               GROUP BY a.Acc
                                UNION
                               SELECT
                               LTRIM(RTRIM(a.Acc)) AS Acc, 
                                      0 AS TrnIntAmt_CollAsOfMonth,
                                      0 AS TrnPriAmt_CollAsOfMonth,
                                      0 AS TrnPenAmt_CollAsOfMonth,
                                      0 AS TrnChgAmt_CollAsOfMonth,  
		                              (-1)*SUM(t.TrnIntAmt/@ccydiv) as TrnIntAmt_CollDaily,
		                              (-1)*SUM(t.TrnPriAmt/@ccydiv) as TrnPriAmt_CollDaily,
		                              (-1)*SUM(t.TrnPenAmt/@ccydiv) as TrnPenAmt_CollDaily,
		                              (-1)*SUM(t.TrnChgAmt/@ccydiv) as TrnChgAmt_CollDaily
                               FROM TrnDaily T,#LN_DETAIL a 
                               WHERE t.acc=a.Acc AND t.TrnType IN('105','205','305','406','705') AND t.GLCode NOT LIKE 'W%'
                                     AND t.CancelledByTrn IS NULL AND (t.ValueDate = @REPORTDATE) 
                               GROUP BY a.Acc   

                               --Clear Reverse Closing ----------------------
                               SELECT Acc, 
                                      SUM(TrnIntAmt_CollAsOfMonth) AS TrnIntAmt_CollAsOfMonth,
                                      SUM(TrnPriAmt_CollAsOfMonth) AS TrnPriAmt_CollAsOfMonth,
                                      SUM(TrnPenAmt_CollAsOfMonth) AS TrnPenAmt_CollAsOfMonth,
                                      SUM(TrnChgAmt_CollAsOfMonth) AS TrnChgAmt_CollAsOfMonth,
		                              SUM(TrnIntAmt_CollDaily) AS TrnIntAmt_CollDaily,
		                              SUM(TrnPriAmt_CollDaily) AS TrnPriAmt_CollDaily,
		                              SUM(TrnPenAmt_CollDaily) AS TrnPenAmt_CollDaily,
		                              SUM(TrnChgAmt_CollDaily) AS TrnChgAmt_CollDaily
                               INTO #RepayMonth
                               FROM #RepayMonth_Temp 
                               GROUP BY Acc
                                /*
                               Part 2
                               */
                            ----Update all loan repayment----------------------
                               UPDATE #LN_DETAIL 
	                            SET TrnIntAmt_CollAsOfMonth=a.TrnIntAmt_CollAsOfMonth,
                                   TrnPriAmt_CollAsOfMonth = a.TrnPriAmt_CollAsOfMonth,
                                   TrnPenAmt_CollAsOfMonth=a.TrnPenAmt_CollAsOfMonth,
                                   TrnChgAmt_CollAsOfMonth=a.TrnChgAmt_CollAsOfMonth,
	                               TrnIntAmt_CollDaily = a.TrnIntAmt_CollDaily,
	                               TrnPriAmt_CollDaily = a.TrnPriAmt_CollDaily,
	                               TrnPenAmt_CollDaily = a.TrnPenAmt_CollDaily,
	                               TrnChgAmt_CollDaily = a.TrnChgAmt_CollDaily
                               FROM #LN_DETAIL l,#RepayMonth a 
                               WHERE l.acc=a.acc    
                            --=========================================================================================================================================================
                            /*
                            Culculate Interest, to get InterestToDate for only active loan
                            */
                            --Create other parameter--
	                            DECLARE @FutureDays			INT
	                            DECLARE	@PenAmt				NUMERIC(18,3)     
	                            DECLARE	@IntODueAmt			NUMERIC(18,3)
	                            DECLARE	@TaxAmt				NUMERIC (18,3)
	                            DECLARE	@TrnChgAmt			NUMERIC(18,3)
	                            DECLARE @IntAmt				NUMERIC(18,3)
	                            SET @FutureDays		=0
	                            SET @IntAmt			=0
	                            SET @PenAmt			=0
	                            SET @IntODueAmt		=0
	                            SET @TaxAmt			=0
	                            SET @TrnChgAmt		=0

	                            --Exec   sp_LNCalcInterest  @idAccount, @FutureDays, @IntAmt Output, @PenAmt Output, @IntODueAmt Output, @TaxAmt Output, @TrnChgAmt Output

	                            Create table #LNCalcInterest
	                            ( 
		                            Acc_Full nvarchar(15),
		                            IntOdueAmt numeric(18,3),
		                            IntToDateAmt numeric(18,3),--= @IntToDate = @IntOdueAmt ( ARI from begining month To Date ) + @AcrIntAmt (ARI end of Previous month that get from LNACC )
		                            PenaltyAmt numeric(18,3),--Penalty Amount Due
		                            TaxAmt numeric(18,3),
		                            TrnChagAmt numeric(18,3)
	                            )
	                            /*
	                            loop from each row to culculate interest for each acc
	                            */	
	                            declare @AcrIntAmt_EndPreMonth numeric(18,3)
	                            DECLARE @ACC_FULL NVARCHAR(15)
	                            DECLARE MyCursor CURSOR FOR
		                            select 
			                            L.Acc+L.Chd,
			                            L.AcrIntAmt_EndPreMonth	
		                            from #LN_DETAIL L
		                            where L.AccStatus between '11' and '98'
		                            OPEN	MyCursor
		                            FETCH	MyCursor INTO @ACC_FULL,@AcrIntAmt_EndPreMonth
		                            WHILE @@FETCH_STATUS <> -1
		                            BEGIN
			                            Exec   sp_LNCalcInterest  @ACC_FULL, @FutureDays, @IntAmt Output, @PenAmt Output, @IntODueAmt Output, @TaxAmt Output, @TrnChgAmt Output
			                            INSERT INTO #LNCalcInterest (Acc_Full,IntOdueAmt,IntToDateAmt,PenaltyAmt,TaxAmt,TrnChagAmt)
			                            SELECT 
				                            @ACC_FULL,
				                            @IntODueAmt,
				                            @AcrIntAmt_EndPreMonth + @IntAmt,--	or @AcrIntAmt_EndPreMonth + @IntODueAmt,
				                            @PenAmt,
				                            @TaxAmt,
				                            @TrnChgAmt			
			                            FETCH NEXT FROM MyCursor INTO @ACC_FULL,@AcrIntAmt_EndPreMonth
		                            END
	                            CLOSE MyCursor
	                            DEALLOCATE  MyCursor
	                            ---update ineterest calculation to temp table of laon ( only active )
	                            UPDATE #LN_DETAIL 
	                            SET 
		                            IntOdueAmt_FromIntCalc_Proc = c.IntOdueAmt,
		                            IntToDateAmt_FromIntCalc_Proc = c.IntToDateAmt,--= @IntToDate = @IntOdueAmt ( ARI from begining month To Date ) + @AcrIntAmt (ARI end of Previous month that get from LNACC )
		                            PenaltyAmtDue_FromIntCalc_Proc = c.PenaltyAmt,
		                            TaxAmtDue_FromIntCalc_Proc = c.TaxAmt,
		                            TrnChagAmtDue_FromIntCalc_Proc = c.TrnChagAmt
                               FROM #LN_DETAIL L,#LNCalcInterest c
                               WHERE l.acc + l.Chd = c.Acc_Full  and L.AccStatus between '11' and '98'

                            --=========================================================================================================================================================
                            /*
                            Paid Of Loan during this months.
                            */
                            select 
	                            cast(co2.CID as int) as idCO,	
	                            co2.DisplayName as CoName,
	                            L.acc,
	                            dbo.[GetFirstDisbDate](L.acc) as DisbDate,
	                            L.PrType,
	                            L.GrantedAmt,
	                            L.AccStatus,
	                            T.TrnPriAmt,
	                            T.TrnIntAmt,
	                            T.TrnPenAmt,
	                            T.TrnChgAmt,
	                            T.BalAmt,
	                            t.TrnDate,
	                            T.TrnType,
	                            T.TrnDesc,
	                            T.ValueDate,
	                            L.MatDate,
	                            @REPORTDATE as ReportDate
                             into #PaidOff
                             from LNACC L
	                            left join CIF c on c.CID = SUBSTRING(l.Acc,4,6) 
	                            left join RELCID g on g.RelatedCID = c.CID and g.Type = '900' --MEMBER TO GROUP
	                            LEFT JOIN RELCID CO ON CO.RelatedCID = G.CID  	AND CO.TYPE = '499'--GROUP TO CO
	                            LEFT JOIN CIF CO2 ON CO2.CID = CO.CID --Join to get CO Name
	                            inner join Trnhist T ON T.ACC = L.ACC 
	                            where L.AccStatus = '99' 
		                            --and L.AccStatusDate between @BeginDateMonth and @REPORTDATE
		                            and T.ValueDate between @BeginDateMonth and @REPORTDATE
		                            --and L.MatDate > T.TrnDate
		                            and ISNULL(T.TrnPriAmt,0) +	isnull(T.TrnIntAmt,0) + isnull(T.TrnPenAmt,0) + ISNULL(T.TrnChgAmt,0)  >0
		                            and T.BalAmt = 0 ---If loan pay many time in this month, only get the last one that paid until balance = 0

                            /*
                            Add more loan for current day in table TrnDaily, 
                            then all transaction not yet moved to TrnHist Table
                            */
                            insert into #PaidOff
                            select 
	                            cast(co2.CID as int) as idCO,	
	                            co2.DisplayName as CoName,
	                            L.acc,
	                            dbo.[GetFirstDisbDate](L.acc) as DisbDate,
	                            L.PrType,
	                            L.GrantedAmt,
	                            L.AccStatus,
	                            T.TrnPriAmt,
	                            T.TrnIntAmt,
	                            T.TrnPenAmt,
	                            T.TrnChgAmt,
	                            T.BalAmt,
	                            T.TrnDate,
	                            T.TrnType,
	                            T.TrnDesc,
	                            T.ValueDate,
	                            L.MatDate,
	                            @REPORTDATE as ReportDate
                             from LNACC L
	                            left join CIF c on c.CID = SUBSTRING(l.Acc,4,6) 
	                            left join RELCID g on g.RelatedCID = c.CID and g.Type = '900' --MEMBER TO GROUP
	                            LEFT JOIN RELCID CO ON CO.RelatedCID = G.CID  	AND CO.TYPE = '499'--GROUP TO CO
	                            LEFT JOIN CIF CO2 ON CO2.CID = CO.CID --Join to get CO Name
	                            inner join TRNDAILY T ON T.ACC = L.ACC 
	                            where L.AccStatus = '99' 
		                            --and L.AccStatusDate between @BeginDateMonth and @REPORTDATE
		                            and T.ValueDate between @BeginDateMonth and @REPORTDATE
		                            --and L.MatDate > T.TrnDate
		                            and ISNULL(T.TrnPriAmt,0) +	isnull(T.TrnIntAmt,0) + isnull(T.TrnPenAmt,0) + ISNULL(T.TrnChgAmt,0)  >0
		                            and T.BalAmt = 0 ---

                            update #LN_DETAIL 
	                            set 
		                            TrnPriAmt_PaidOff = P.TrnPriAmt,
		                            TrnIntAmt_PaidOff = P.TrnIntAmt,
		                            TrnChgAmt_PaidOff = P.TrnChgAmt,
		                            TrnPenAmt_PaidOff = P.TrnPenAmt,
		                            PaidOffDate = P.ValueDate
                            from #LN_DETAIL L1, #PaidOff P 
                            WHERE P.Acc = L1.Acc AND L1.AccStatus = '99'
                            -------------------------------------------
                             ---Update provision
                            update #LN_DETAIL 
	                            set 
	                            MasterField = 
		                            case 
			                            when L1.ShortCode = 0  then
				                            case 
					                            when L1.AgeofLoan =0 then '0000' --auto normal loan
					                            when L1.AgeofLoan between 1 and  @MaxDayShort  then --@MaxDayShort = 90 days --> lost 
						                            (select 
							                            R.code	+'0'			
							                            from AmretLoanProvision R
							                            where R.term = 0 --0 for short term
							                            and R.NumDaysNo = (
												                            select max(R2.NumDaysNo) from AmretLoanProvision R2 
													                            where R2.Term  = 0 
													                            and  R2.NumDaysNo<=L1.AgeofLoan
											                               )
						                            )
					                            when L1.AgeofLoan > @MaxDayShort then '0040' --auto loss	
					                            else 'UnKnown'				
				                            end			
			                            when L1.ShortCode = 1 then
				                            case 					
					                            when L1.AgeofLoan =0 then '0001' --normal
					                            when L1.AgeofLoan between 1 and @MaxDayLong then --long terms
						                            (
						                            select 
							                             R.code +'1'	--1 for long term			
						                            from AmretLoanProvision R
							                            where R.term = @TermDefined---@TermDefined=366 day and up is long term 
								                            and R.NumDaysNo = (
													                            select max(R2.NumDaysNo) 
													                            from AmretLoanProvision R2 
														                            where R2.Term  = @TermDefined
															                            and R2.NumDaysNo<=L1.AgeofLoan
													                            )
												                            )
					                            when L1.AgeofLoan > @MaxDayLong then '0041' --auto loss
				                            end
		                            end
                              from #LN_DETAIL L1
                            ----Update LastPaidDate, LastPaid Amt
                            update #LN_DETAIL 
	                            set LastPaidDate = (select top(1) max(t2.TrnDate) from TRNHIST T2 where t2.Acc = L.Acc and t2.TrnType in ('451','431','401')),--'401','403','405','411','421','431','451','453'
	                            LastPaidAmt = (
				                            select isnull(sum(T1.TrnAmt) ,0)
				                            from TRNHIST T1
					                            where 
					                            T1.Acc = L.Acc 
					                            and T1.TrnDate = (select top(1) max(T3.TrnDate) from TRNHIST T3 where t3.Acc = L.acc and t3.trnType in ('451','431','401'))
					                            and T1.trnType in ('451','431','401')
				                             ),
	                            PreDisbAmt = (
		                            Select L2.GrantedAmt 
		                            from LNACC L2 
		                            where substring(L2.acc,4,6) = L.IdClient 
			                              and L2.Opendate = 
				                            (
					                            select max(Opendate) as Opendate from LNACC L3 where substring(L3.Acc,4,6)=L.IdClient and L3.AccStatus='99' group by substring(L3.Acc,4,6)
				                            ) 			  
			                              and L2.AccStatus='99'
			                            ),
	                            PreAcc = (
			                            Select L4.Acc from LNACC L4 
			                            where L4.Opendate = 
			                            (
				                            select max(L5.Opendate) as Opendate 
				                            from LNACC L5 where substring(L5.Acc,4,6)=L.IdClient and L5.AccStatus='99' group by SUBSTRING(L5.acc,4,6)
			                            ) and substring(L4.acc,4,6) = L.IdClient and L4.AccStatus='99'
		                            ),
	                            NextPaidDate = (
		                            select distinct(Min(I.DueDate)) from LNINST I where I.Acc = L.Acc and I.Chd = L.chd and  I.Status = 0 and I.status <> 8
	                            ),
	                            IntRemainInSchedule = (
		                            select 
			                            sum(I.IntAmt/@CCYDIV) 
		                            from LNINST I 
		                            inner join LNACC L6 on L6.Acc = I.Acc and L6.Chd = I.Chd and I.Acc = L.acc and I.Chd = L.Chd
		                            inner join RELACC R on R.ACC = L.Acc and R.Chd = L.Chd
		                            where I.Status <> 8 and R.AppType = '4' and r.Type = '010'
		                            ),
		                            LastPreDueDateOfReportDate = 
		                            (
			                            select distinct(max(T1.DueDate)) from LNINST T1 where t1.Acc = L.Acc and t1.Status <> 8
					                            and T1.DueDate <= @ReportDate		
		                            )

                            from #LN_DETAIL L 

                            ---Protect loan disburst but never paid
                            update #LN_DETAIL 
	                            set LastPreDueDateOfReportDate = 
		                            (
			                            case 
				                            when LastPreDueDateOfReportDate is null then L3.DisbDate
				                            else LastPreDueDateOfReportDate
			                            end
		                            )
                            from #LN_DETAIL L3

                            update #LN_DETAIL 
	                            set NextPaidAmt = (
					                            select 
						                            case 
							                            when L2.NextPaidDate is null then null
							                            else isnull(I.PriAmt/@CCYDIV,0) + isnull(I.IntAmt/@ccydiv,0) + isnull(I.ChargesAmt/@ccydiv,0) 
						                            end 
					                            from LNINST I 
						                            inner join LNACC L6 on L6.Acc = I.Acc and L6.Chd = I.Chd and I.Acc = L2.acc and I.Chd = L2.Chd
						                            inner join RELACC R on R.ACC = L2.Acc and R.Chd = L2.Chd
					                            where --I.Status <> 8  ---reschedule
					                            I.Status = 0
					                            and R.AppType = '4' and r.Type = '010'
					                            and i.DueDate = L2.NextPaidDate
						                            ),
		                            NextDueWorkDate = dbo.GetWorkingDueDate(NextPaidDate),	
	                                BalAmtLastDueOfReportDate = (
			                            select top(1)
				                            case 
					                            when (I3.BFBalAmt/@CcyDiv - I3.OrigPriAmt/@CcyDiv = 0) then I3.BFBalAmt/@CcyDiv
					                            when (I3.BFBalAmt/@CcyDiv - I3.OrigPriAmt/@CcyDiv >0 ) then (I3.BFBalAmt/@CcyDiv- I3.OrigPriAmt/@CcyDiv)
				                            end as AA		
			                            from LNINST I3
				                            where I3.Acc = L2.Acc  and I3.Status <> '8'
				                            and I3.DueDate = LastPreDueDateOfReportDate		
			                            ),
		                            PrinAmtLastDueDateCollect = (
			                            select 
				                            top(1) I4.OrigPriAmt/@ccydiv 
			                            from LNINST I4 
				                            where acc = L2.Acc 
				                            and I4.DueDate = L2.LastPreDueDateOfReportDate 
				                            and I4.Status <> '8'
		                            ),
		                            RemainLoanTerm = (
			                            select 
				                            count(distinct(I5.DueDate))							
			                            from LNINST I5
				                            inner join LNACC L6 on L6.Acc = I5.Acc and L6.Chd = I5.Chd and I5.Acc = L2.acc and I5.Chd = L2.Chd
				                            inner join RELACC R on R.ACC = L2.Acc and R.Chd = L2.Chd
					                            where --I.Status <> 8  ---reschedule
						                            I5.Status = 0
						                            and R.AppType = '4' and r.Type = '010'
						                            and I5.PaidDate is null
					                            --and i.DueDate = L2.NextPaidDate
		                            )
                            from #LN_DETAIL L2 
	


                            --==========================================================================================================================================================
                            ---Replace string ',' with ';' for CSV
                            declare @ColumnName nvarchar(50)
                            declare xCursor cursor for
	                            SELECT 
		                            distinct(sc.name)
	                            --st.name as type_name 
	                            --sc.max_length 
	                            FROM tempdb.sys.columns sc inner join sys.types st on st.system_type_id=sc.system_type_id 
		                            WHERE [object_id] = OBJECT_ID('tempdb..#LN_DETAIL')
		                            and st.name in ('char','nchar','ntext','nvarchar','text','varchar') 
                            OPEN	xCursor
                            FETCH	xCursor INTO @ColumnName
                            WHILE @@FETCH_STATUS <> -1
                            BEGIN
	                            PRINT @ColumnName
	                            EXECUTE (' UPDATE #LN_DETAIL 
	                            SET  '  + @ColumnName + ' = REPLACE( '  + @ColumnName + ','','','+''';'')' )
	                            FETCH NEXT FROM xCursor INTO @ColumnName
                            END
                            CLOSE xCursor
                            DEALLOCATE xCursor


                            --select * into [DailyReports].dbo.LN_DETAIL from #LN_DETAIL
                            select 
				                row_number() over(order by Acc) as ID,
				                @BrCode as BrCode,
				                @BrShort as BrShort,
				                Acc+chd as AccountNumber,
				                IdCo,
				                CoName,
				                CoNameKh,
				                ClientName,
				                ClientNameKh,
				                Gender,
				                CoBorrowerName,
				                CoBorrowerGender,
				                Mobile1,
				                AdressKhmer,
				                GrantedAmt,
				                Balamt,
				                DisbDate,
				                MatDate,
				                IntBalAmt,
				                PenBalAmt,
				                OduePriAmt,
				                OdueIntAmt,
				                PaymentFrequency,
				                LastPaidDate,
				                Ageofloan,
				                Reportdate

				                from #LN_DETAIL where accstatus between '11' and '98'


                            --select * into LN_DETAIL from #LN_DETAIL

                            TRUNCATE TABLE #LN_DETAIL
                            truncate table #PaidOff
                            TRUNCATE TABLE #LNCalcInterest

                            --DROP TABLE #LN_DETAIL
                            drop table #PaidOff	
                            drop table #RepayMonth_Temp
                            drop table #RepayMonth
                            DROP TABLE #LNCalcInterest

                    ";
                    }
                    else
                    {

                        Sql = @"
                            SET DATEFORMAT DMY
                            DECLARE @REPORTDATE DATETIME
                            DECLARE @BeginDateYear datetime
                            DECLARE @BeginDateMonth datetime
                            DECLARE @PreThreeMonthDate datetime
                            DECLARE @ccydiv INT
                            declare @MaxDayShort int
                            declare @MaxDayLong int 
                            declare @TermDefined int
                            -------------------------------------------------

            		         --set @REPORTDATE='30/06/2011';
			                --set @BeginDateYear='01/01/2011';
			                --set @BeginDateMonth='01/06/2011';
			                select 
				                @REPORTDATE = CurrRunDate,
				                @BeginDateYear = convert(datetime,'01/01/'+cast(year(CurrRunDate) as varchar(4)),103),
				                @BeginDateMonth = convert(datetime,'01/' + cast(month(CurrRunDate) as varchar(2))+'/' + cast(year(CurrRunDate) as varchar(4)),103),
				                @PreThreeMonthDate = dateadd(month,3,CurrRunDate)
			                from BRPARMS

                            SET @ccydiv=(SELECT ccydiv FROM ccy)
                            set @TermDefined = ((select top(1) max(M1.Term) from AmretLoanProvision M1))
                            set @MaxDayShort = (select max(NumDaysNo) from AmretLoanProvision where Term = 0 )
                            set @MaxDayLong = (select max(NumDaysNo) from AmretLoanProvision where Term = @TermDefined)
			                declare @BrCode nvarchar(10)
			                declare @BrShort nvarchar(10)
			                declare @BrName nvarchar(50)
			                declare @DbName nvarchar(50)
			                set @DbName = (SELECT DB_NAME())

			                select 
				                @BrCode = SubBranchCode,
				                @BrShort = SubBranchID,
				                @BrName = SubBranchNameLatin
			                from SKP_Brlist 
			                where DBName = (select left(@DbName,len(@DbName)-4))

                            --print '01/' + cast(month(CurrRunDate) as varchar(2))
                            print @BeginDateMonth
                            print @BeginDateYear
                            -------------------------------------------------
                            /*
                            CLEAN TEMP TABLE 
                            */
                            IF OBJECT_ID('tempdb..#LN_DETAIL') IS NOT NULL
                                DROP TABLE #LN_DETAIL
                            IF OBJECT_ID('tempdb..#PaidOff') is not null
	                            drop table #PaidOff
                            IF OBJECT_ID('tempdb..#RepayMonth_Temp') IS NOT NULL
	                            DROP TABLE #RepayMonth_Temp
                            IF OBJECT_ID('tempdb..#RepayMonth') IS NOT NULL
	                            DROP TABLE #RepayMonth
                            IF OBJECT_ID('tempdb..#LNCalcInterest') IS NOT NULL
	                            DROP TABLE #LNCalcInterest

                            SELECT  
		                            L.Acc,
		                            L.Chd,
		                            /*
		                            CO information
		                            */
		                            co2.CID as IdCO,		
		                            ltrim(rtrim(co2.DisplayName)) as CoName,
		                            ltrim(rtrim(CO_Kh.Name1)) as CoNameKh,	
		                            G.CID AS IdGroup,
		                            ltrim(rtrim(g2.DisplayName)) as GroupName,
		                            ----------------------------------------	
		                            c.CID as IdClient,
		                            ltrim(rtrim(c.DisplayName)) as ClientName,
		                            ltrim(rtrim(CKh.Name1)) as ClientNameKh,		
		                            case 
			                            when c.GenderType = '001' then 'M' 
			                            when c.GenderType = '002' then 'F'
			                            else	''
		                            end as Gender,
		                            case 
			                            when c.CivilStatusCode = '00D' then 'D'
			                            when c.CivilStatusCode = '00M' then 'M'
			                            when c.CivilStatusCode = '00S' then 'S'
			                            when c.CivilStatusCode = '00W' then 'W'
			                            else ''
		                            end as [MaritalStatus],
		                            case 
			                            when left(c.Nid,1) = 'M'  then 'M'--M=Maried Letter
			                            when left(c.Nid,1) = 'C' then 'C' --C= Civil Servan 
			                            when left(c.Nid,1) = 'F' then 'F' 
			                            when left(c.Nid,1) = 'B' then 'B' 
			                            when left(c.Nid,1) = 'N' and ltrim(rtrim(c.nid)) <> 'N/A' then 'N' 
			                            when left(c.Nid,1)= 'D' then 'D'
			                            when left(c.Nid,1)= 'P' then 'P'
			                            when left(c.Nid,1) = 'R' then 'R'
			                            else 'Unknown'
		                            end	 as [IDType_1],--N,F,D,B,G,R... columns 2 in CBC
		                            case 
			                            when left(c.Nid,1) in ('N','F','R','D','B','C','P','M') and ltrim(rtrim(c.nid)) <> 'N/A' then  right(ltrim(rtrim(c.Nid)),len(c.Nid)-1)
			                            when  left(c.Nid,1) in ('F','G','R') then
				                            case when CAST(right(ltrim(rtrim(c.Nid)),len(c.Nid)-1) AS int) = 0 then 'N/A' 
					                            else right(c.Nid,len(c.Nid)-1)
				                            end
			                            else c.Nid
		                            end  as [IDNumber_1],--3
		                            c.BirthDate	as [DateofBirth],		
		                            isnull(ltrim(rtrim(c.Mobile1)),'') as Mobile1,
		                            isnull(ltrim(rtrim(c.Mobile2)),'') as Mobile2,		
		                            ltrim(rtrim(c.CIFCode1)) as CIFCode1,
		                            ltrim(rtrim(c.CIFCode2)) as CIFCode2,
		                            ltrim(rtrim(c.CIFCode3)) as CIFCode3,
		                            ltrim(rtrim(c.CIFCode4)) as CIFCode4,
		                            ltrim(rtrim(c.CIFCode5)) as CIFCode5,			
		                            ltrim(rtrim(c.CIFCode6)) as CIFCode6,
		                            ltrim(rtrim(c.CIFCode7)) as CIFCode7,
		                            ltrim(rtrim(c.CIFCode8)) as CIFCode8,
		                            ltrim(rtrim(c.CIFCode9)) as CIFCode9,
		                            CIF_FamilyMember.FullDesc as FamilyMemberDesc,
		                            CIF_Income.FullDesc as CIF_IncomeDesc,
		                            CIF_Occupation.FullDesc as CIF_OccupationDesc,		
		                            CIF_TotalAsset.FullDesc as CIF_TotalAssetDesc,
		                            c.LocationCode,
		                            VIL.NameLatin AS Village,
		                            COM.NameLatin as Commune,
		                            Dis.NameLatin as District,
		                            PRV.NameLatin as Province,
		                            Vil.NameKhmer + ';' + Com.NameKhmer + ';' + dis.NameKhmer + ';' + PRV.NameKhmer as AdressKhmer,
		                            -------------------------------
		                            ---Co-Borrowere Information
		                            Ckh.Name2 as CoBorrowerName,
		                            case 
			                            when CKh.Sex = 'Rbus' then 'M' 
			                            when CKh.Sex ='RsI' then 'F'
			                            else ''
		                            end as CoBorrowerGender,
		                            --'' as CoBorrowerDoBOld,
		                            Ckh.Date1 as CoBorrowerDoB,
		                            Ckh.CardType as CoBorrowerIDType,
		                            Ckh.Name6 as CoBorrowerIDNum,
		                            Ckh.RelatedName as CoBorrowerRelativeType,
		                            ------------------------------
		                            ---Loan Information		
		                            L.AppType,
		                            L.BalAmt/ @ccydiv AS BalAmt,
		                            ------
		                            L.AcrIntAmt /@ccydiv as AcrIntAmt_EndPreMonth,---Accrue interest Ammount until end of prevouse month
		                            L.AcrChgAmt /@ccydiv as AcrChgAmt_EndPreMonth,
		                            L.AcrPenAmt /@ccydiv as AcrPenAmt_EndPreMonth,
		                            L.AcrIntODuePriAmt /@ccydiv as AcrIntODuePriAmt,
		                            ------
		                            L.IntBalAmt / @ccydiv as IntBalAmt,
		                            L.PenBalAmt / @ccydiv as PenBalAmt,
		                            L.StopIntTF ,
		                            L.TrnSeq,
		                            -------
		                            Case 
			                            WHEN L.OduePriAmt < 0 THEN 0 ELSE L.OduePriAmt / @ccydiv 
		                            END AS OduePriAmt,
		                            L.OdueIntAmt / @ccydiv AS OdueIntAmt,
		                            L.StopAcrIntTF ,
		                            L.ReschedSeq,
		                            L.CcyType,
		                            dbo.[GetFirstDisbDate](l.acc) as DisbDate,
		                            L.GrantedAmt/ @ccydiv AS GrantedAmt,
		                            L.IntRate,
		                            L.IntEffDate,
		                            L.MatDate,	
		                            L.GLCode,
		                            L.GLCodeOrig,	
		                            L.InstNo,
		                            L.FreqType,
		                            case 
			                            when l.FreqType = '012' then 'Monthly' --monthly
			                            when l.FreqType = '026' then '2 Week' --every 2 week
			                            when l.FreqType = '052' then 'Weekly' --weekly
			                            when l.FreqType = '013' then '4 Week' --every 4 week
			                            else '' ---************protect in case we add some more product->then we can saw error.
		                            end as [PaymentFrequency],--16
		                            L.PrType,
		                            P.FullDesc as PrName,
		                            L.LNCode1,
		                            L.LNCode2,
		                            L.LNCode3,
		                            L.LNCode4,
		                            L.LNCode5,
		                            L.LNCode6,
		                            L.LNCode7,
		                            L.LNCode8,	
		                            Dbo.GetTermOfLoan(L.Acc) AS TERMS,
		                            -----------
		                            CAST('01/01/1900' AS DATETIME) as LastPaidDate,
		                            CAST( 0 AS NUMERIC(18,3)) as LastPaidAmt,
		                            cast('' as nvarchar(15)) as PreAcc,
		                            CAST( 0 AS NUMERIC(18,3)) as PreDisbAmt,
		                            Cast('01/01/1900' AS DATETIME) as NextPaidDate,
		                            cast('01/01/1900' as datetime) as NextDueWorkDate,
		                            Cast(0 as numeric(18,3)) as NextPaidAmt,
		                            Cast(0 as numeric(18,3)) as IntRemainInSchedule,
		                            cast(0 as numeric(18,3)) as RemainLoanTerm,
		                            cast('01/01/1900' as datetime) as LastPreDueDateOfReportDate,
		                            cast(0 as numeric(18,3)) as BalAmtLastDueOfReportDate,
		                            cast(0 as numeric(18,3)) as PrinAmtLastDueDateCollect,
		                            --cast('01/01/1900' as datetime) as LastTransDateDueDate,
		                            --cast(0 as numeric(18,3)) as IntAmtAfterPreDueDateOfReportDate,
		                            --cast(0 as numeric(18,3)) as TotalAmtRequiredToClose,
		
		                            -----------
		                            dbo.GetAgeOfLoan(@REPORTDATE,l.Acc) as AgeofLoan,
		                            cast(0 as int) as NoInstLate,
		                            CAST('01/01/1900' AS DATETIME) AS FirstDayLate,
		                            --CAST('' as nvarchar(10) ) as StatusReNewLoan,
		                            -------------------------------------------------------------------------
		                            ---Collection --OK
		                            CAST( 0 AS NUMERIC(18,2)) as TrnIntAmt_CollYear,--reserved for collection
                                    CAST( 0 AS NUMERIC(18,2)) as TrnPriAmt_CollYear,--reserved for collection
                                    CAST( 0 AS NUMERIC(18,2)) as TrnPenAmt_CollYear,--reserved for collection
                                    CAST( 0 AS NUMERIC(18,2)) as TrnChgAmt_CollYear, --reserved for collection
		                            CAST( 0 AS NUMERIC(18,2)) as TrnIntAmt_CollAsOfMonth,--reserved for collection
                                    CAST( 0 AS NUMERIC(18,2)) as  TrnPriAmt_CollAsOfMonth,--reserved for collection
                                    CAST( 0 AS NUMERIC(18,2)) as TrnPenAmt_CollAsOfMonth,--reserved for collection
                                    CAST( 0 AS NUMERIC(18,2)) as TrnChgAmt_CollAsOfMonth, --reserved for collection
		                            CAST( 0 AS NUMERIC(18,2)) as TrnIntAmt_CollDaily,--reserved for collection
                                    CAST( 0 AS NUMERIC(18,2)) as TrnPriAmt_CollDaily,--reserved for collection
                                    CAST( 0 AS NUMERIC(18,2)) as TrnPenAmt_CollDaily,--reserved for collection
                                    CAST( 0 AS NUMERIC(18,2)) as TrnChgAmt_CollDaily, --reserved for collection
		                            ------------------------------------------------------------------------------
		                            ---Interesr Culculation Procedure--OK
		                            CAST( 0 AS NUMERIC(18,2)) as IntOdueAmt_FromIntCalc_Proc,
		                            CAST( 0 AS NUMERIC(18,2)) as IntToDateAmt_FromIntCalc_Proc,--= @IntToDate = @IntOdueAmt ( ARI from begining month To Date ) + @AcrIntAmt (ARI end of Previous month that get from LNACC )
		                            CAST( 0 AS NUMERIC(18,2)) as PenaltyAmtDue_FromIntCalc_Proc,
		                            CAST( 0 AS NUMERIC(18,2)) as TaxAmtDue_FromIntCalc_Proc,
		                            CAST( 0 AS NUMERIC(18,2)) as TrnChagAmtDue_FromIntCalc_Proc,

		                            ------------------------------------------------------------------------------
		                            ----PaidOff, paid off can be normal paid off ( loan Past Maturity then paid off ) or loan that paid off before paturity
		                            CAST( 0 AS NUMERIC(18,2)) as TrnIntAmt_PaidOff,--reserved for collection
                                    CAST( 0 AS NUMERIC(18,2)) as TrnPriAmt_PaidOff,--reserved for collection
                                    CAST( 0 AS NUMERIC(18,2)) as TrnPenAmt_PaidOff,--reserved for collection
                                    CAST( 0 AS NUMERIC(18,2)) as TrnChgAmt_PaidOff, --reserved for collection
		                            CAST('01/01/1900' AS DATETIME)  as PaidOffDate,
		                            CAST('' AS NVARCHAR(30) ) AS STATUS_RENEW,
		                            -----------------------------------------------------------------
		                            --Provision code
		                            ShortCode = 
		                            case 
			                            when dbo.GetTermOfLoan(L.acc) >= (select max(TERM) from AmretLoanProvision) then 1
			                            else 0 
		                            end,
		                            MasterField = cast('00000' as nvarchar(10)),
		                            ----------------------------------------------------------------
		                            L.AccStatus,
		                            L.AccStatusDate,
		                            BusType.FullDesc BusType,
		                            CASE 
			                            WHEN L.LNCode1 BETWEEN '100' AND '199' THEN 'Agricultural'
			                            WHEN L.LNCode1 BETWEEN '200' AND '399' THEN 'Trading & Commerce'
			                            WHEN L.LNCode1 BETWEEN '400' AND '499' THEN 'Service' 
			                            WHEN L.LNCode1 BETWEEN '500' AND '599' THEN 'Transportation' 
			                            WHEN L.LNCode1 BETWEEN '600' AND '699' THEN 'Construction' 
			                            WHEN L.LNCode1 >  '700' THEN 'Consumption (Householg/Family)'
		                            END AS BusSector,
		                            Colateral.FullDesc as LoanColateral,
		                            LNCycle.FullDesc as LoanCycle,
		                            /*
		                            guaranter information
		                            */
		                            R.CID as idGuarenter,
		                            ltrim(rtrim(GC.DisplayName)) as GuarenterName,
		                            ltrim(rtrim(GCKh.Name1)) as GuarenterNameKh,		
		                            case 
			                            when GC.GenderType = '001' then 'M' 
			                            when GC.GenderType = '002' then 'F'
			                            else	''
		                            end as GuarenterGender,
		                            case 
			                            when GC.CivilStatusCode = '00D' then 'D'
			                            when GC.CivilStatusCode = '00M' then 'M'
			                            when GC.CivilStatusCode = '00S' then 'S'
			                            when GC.CivilStatusCode = '00W' then 'W'
			                            else ''
		                            end as GuarenterMaritalStatus,
		                            case 
			                            when left(GC.Nid,1) = 'M'  then 'M'--M=Maried Letter
			                            when left(GC.Nid,1) = 'C' then 'C' --C= Civil Servan 
			                            when left(GC.Nid,1) = 'F' then 'F' 
			                            when left(GC.Nid,1) = 'B' then 'B' 
			                            when left(GC.Nid,1) = 'N' and ltrim(rtrim(c.nid)) <> 'N/A' then 'N' 
			                            when left(GC.Nid,1)= 'D' then 'D'
			                            when left(GC.Nid,1)= 'P' then 'P'
			                            when left(GC.Nid,1) = 'R' then 'R'
			                            else 'Unknown'
		                            end	 as Guarenter_IDType_1,--N,F,D,B,G,R... columns 2 in CBC
		                            case 
			                            when left(GC.Nid,1) in ('N','F','R','D','B','C','P','M') and ltrim(rtrim(GC.nid)) <> 'N/A' then  right(ltrim(rtrim(GC.Nid)),len(GC.Nid)-1)
			                            when  left(GC.Nid,1) in ('F','G','R') then
				                            case when CAST(right(ltrim(rtrim(GC.Nid)),len(GC.Nid)-1) AS int) = 0 then 'N/A' 
					                            else right(GC.Nid,len(GC.Nid)-1)
				                            end
			                            else GC.Nid
		                            end  as GuarenterIDNumber_1,--3
		                            GC.BirthDate	as GuarenterDateofBirth,		
		                            isnull(ltrim(rtrim(GC.Mobile1)),'') as GuarenterMobile1,
		                            isnull(ltrim(rtrim(GC.Mobile2)),'') as GuarenterMobile2,
		                            GC.locationCode as GuarenterLocationCode,
		                            /*
		                            Co Guarenter information
		                            */
		                            GCkh.Name2 as CoGuarenerName,
		                            case 
			                            when CKh.Sex = 'Rbus' then 'M' 
			                            when CKh.Sex ='RsI' then 'F'
			                            else ''
		                            end as CoGuarenerGender,
		                            GCkh.Date1 as CoGuarenerDoB,
		                            GCkh.CardType as CoGuarenerIDType,
		                            GCkh.Name6 as CoGuarenerIDNum,
		                            GCkh.RelatedName as CoGuarenerRelativeType,

		                            @REPORTDATE AS REPORTDATE
	                            INTO #LN_DETAIL 
	                            FROM LNACC L 
		                            LEFT JOIN CIF C ON C.CID = substring(L.ACC,4,6) and c.type = '001'--Client
		                            LEFT JOIN relcid G ON G.relatedcid = C.CID AND G.TYPE = '900'  --MEMBER TO GROUP
		                            LEFT JOIN CIF G2 ON G2.CID = G.CID  --JOIN TO GET GROUP INFORMATION
		                            LEFT JOIN RELCID CO ON CO.RelatedCID = G.CID  	AND CO.TYPE = '499'--GROUP TO CO
		                            LEFT JOIN CIF CO2 ON CO2.CID = CO.CID 
		                            LEFT JOIN PRPARMS P ON P.PrType = L.PrType and P.AppType = 4
		                            LEFT JOIN VTSDCIF CKh on Ckh.CID = c.CID and ckh.Type = '001'--join to get client name & co-borrower in khmer
		                            LEFT JOIN VTSDVillage VIL ON VIL.Code = C.LocationCode
		                            LEFT JOIN VTSDCommune COM ON COM.CODE = VIL.CodeOfCommune
		                            LEFT JOIN VTSDDistrict DIS ON DIS.Code = COM.CodeOfDistrict
		                            LEFT JOIN VTSDProvince PRV ON PRV.Code = DIS.CodeOfProvince
		                            LEFT JOIN VTSDCIF CO_Kh	ON CO_Kh.CID = co2.CID and CO_Kh.Type = '499'
		                            LEFT JOIN USERLOOKUP BusType ON BusType.LookUpCode = L.LNCode1 and BusType.LookUpId = '41'	
		                            LEFT JOIN USERLOOKUP Colateral ON Colateral.LookUpCode = L.LNCode2 and Colateral.LookUpId = '42'
		                            LEFT JOIN USERLOOKUP LNCycle on  LNCycle.LookUpCode = L.LNCode3 and LNCycle.LookUpId = '43'
		                            LEFT JOIN USERLOOKUP CIF_FamilyMember ON  CIF_FamilyMember.LookUpCode = C.CIFCode1 and CIF_FamilyMember.LookUpId = '61'--62
		                            LEFT JOIN USERLOOKUP CIF_Income ON  CIF_Income.LookUpCode = C.CIFCode2 and CIF_Income.LookUpId = '62'
		                            LEFT JOIN USERLOOKUP CIF_Occupation ON  CIF_Occupation.LookUpCode = C.CIFCode3 and CIF_Occupation.LookUpId = '63'
		                            LEFT JOIN USERLOOKUP CIF_TotalAsset ON  CIF_TotalAsset.LookUpCode = C.CIFCode4 and CIF_TotalAsset.LookUpId = '64'
		                            LEFT JOIN relacc R on R.Acc + R.Chd= L.Acc + L.Chd and R.Type  = '030' and R.AppType = '4'--type = 30 for guarentee
		                            LEFT JOIN CIF GC ON GC.CID = R.CID--guaretee
		                            LEFT JOIN VTSDCIF GCKh on GCKh.CID = GC.CID and gckh.Type = '001'--join to get client name & co-borrower in khmer
		                            --INNER JOIN #LNProvisions PRO ON PRO.Acc  = L.Acc+L.Chd		
	                            WHERE L.AccStatus BETWEEN '11' AND '98' and l.Acc+l.Chd='" + re.accountnumber + @"'
                            ---==================================================================================================================================================
                            ----FOR CLOSE LOAN IN MONTH

                            INSERT INTO #LN_DETAIL
                            SELECT  
		                            L.Acc,
		                            L.Chd,
		                            co2.CID as IdCO,		
		                            ltrim(rtrim(co2.DisplayName)) as CoName,
		                            ltrim(rtrim(CO_Kh.Name1)) as CoNameKh,	
		                            G.CID AS IdGroup,
		                            ltrim(rtrim(g2.DisplayName)) as GroupName,
		                            ----------------------------------------	
		                            c.CID as IdClient,
		                            ltrim(rtrim(c.DisplayName)) as ClientName,
		                            ltrim(rtrim(CKh.Name1)) as ClientNameKh,		
		                            case 
			                            when c.GenderType = '001' then 'M' 
			                            when c.GenderType = '002' then 'F'
			                            else	''
		                            end as Gender,
		                            case 
			                            when c.CivilStatusCode = '00D' then 'D'
			                            when c.CivilStatusCode = '00M' then 'M'
			                            when c.CivilStatusCode = '00S' then 'S'
			                            when c.CivilStatusCode = '00W' then 'W'
			                            else ''
		                            end as [MaritalStatus],
		                            case 
			                            when left(c.Nid,1) = 'M'  then 'M'--M=Maried Letter
			                            when left(c.Nid,1) = 'C' then 'C' --C= Civil Servan 
			                            when left(c.Nid,1) = 'F' then 'F' 
			                            when left(c.Nid,1) = 'B' then 'B' 
			                            when left(c.Nid,1) = 'N' and ltrim(rtrim(c.nid)) <> 'N/A' then 'N' 
			                            when left(c.Nid,1)= 'D' then 'D'
			                            when left(c.Nid,1)= 'P' then 'P'
			                            when left(c.Nid,1) = 'R' then 'R'
			                            else 'Unknown'
		                            end	 as [IDType_1],--N,F,D,B,G,R... columns 2 in CBC
		                            case 
			                            when left(c.Nid,1) in ('N','F','R','D','B','C','P','M') and ltrim(rtrim(c.nid)) <> 'N/A' then  right(ltrim(rtrim(c.Nid)),len(c.Nid)-1)
			                            when  left(c.Nid,1) in ('F','G','R') then
				                            case when CAST(right(ltrim(rtrim(c.Nid)),len(c.Nid)-1) AS int) = 0 then 'N/A' 
					                            else right(c.Nid,len(c.Nid)-1)
				                            end
			                            else c.Nid
		                            end  as [IDNumber_1],--3
		                            c.BirthDate	as [DateofBirth],		
		                            isnull(ltrim(rtrim(c.Mobile1)),'') as Mobile1,
		                            isnull(ltrim(rtrim(c.Mobile2)),'') as Mobile2,		
		                            ltrim(rtrim(c.CIFCode1)) as CIFCode1,
		                            ltrim(rtrim(c.CIFCode2)) as CIFCode2,
		                            ltrim(rtrim(c.CIFCode3)) as CIFCode3,
		                            ltrim(rtrim(c.CIFCode4)) as CIFCode4,
		                            ltrim(rtrim(c.CIFCode5)) as CIFCode5,			
		                            ltrim(rtrim(c.CIFCode6)) as CIFCode6,
		                            ltrim(rtrim(c.CIFCode7)) as CIFCode7,
		                            ltrim(rtrim(c.CIFCode8)) as CIFCode8,
		                            ltrim(rtrim(c.CIFCode9)) as CIFCode9,
		                            CIF_FamilyMember.FullDesc as FamilyMemberDesc,
		                            CIF_Income.FullDesc as CIF_IncomeDesc,
		                            CIF_Occupation.FullDesc as CIF_OccupationDesc,		
		                            CIF_TotalAsset.FullDesc as CIF_TotalAssetDesc,
		                            c.LocationCode,
		                            VIL.NameLatin AS Village,
		                            COM.NameLatin as Commune,
		                            Dis.NameLatin as District,
		                            PRV.NameLatin as Province,
		                            Vil.NameKhmer + ';' + Com.NameKhmer + ';' + dis.NameKhmer + ';' + PRV.NameKhmer as AdressKhmer,
		                            -------------------------------
		                            ---garanter Information
		                            Ckh.Name2 as CoBorrowerName,
		                            case 
			                            when CKh.Sex = 'Rbus' then 'M' 
			                            when CKh.Sex ='RsI' then 'F'
			                            else ''
		                            end as CoBorrowerGender,
		                            --'' as CoBorrowerDoBOld,
		                            Ckh.Date1 as CoBorrowerDoB,
		                            Ckh.CardType as CoBorrowerIDType,
		                            Ckh.Name6 as CoBorrowerIDNum,
		                            Ckh.RelatedName as CoBorrowerRelativeType,
		                            ------------------------------
		                            ---Loan Information	
		                            L.AppType,
		                            L.BalAmt/ @ccydiv AS BalAmt,
		                            ------
		                            L.AcrIntAmt /@ccydiv as AcrIntAmt_EndPreMonth,---Accrue interest Ammount until end of prevouse month
		                            L.AcrChgAmt /@ccydiv as AcrChgAmt_EndPreMonth,
		                            L.AcrPenAmt /@ccydiv as AcrPenAmt_EndPreMonth,
		                            L.AcrIntODuePriAmt /@ccydiv as AcrIntODuePriAmt,
		                            ---
		                            L.IntBalAmt / @ccydiv as IntBalAmt,
		                            L.PenBalAmt / @ccydiv as PenBalAmt,
		                            L.StopIntTF ,	
		                            L.TrnSeq,
		                            ----
		                            Case 
			                            WHEN L.OduePriAmt < 0 THEN 0 ELSE L.OduePriAmt / @ccydiv 
		                            END AS OduePriAmt,
		                            L.OdueIntAmt / @ccydiv AS OdueIntAmt,
		                            L.StopAcrIntTF ,
		                            L.ReschedSeq,
		                            L.CcyType,
		                            dbo.[GetFirstDisbDate](l.acc) as DisbDate,
		                            L.GrantedAmt/ @ccydiv AS GrantedAmt,
		                            L.IntRate,
		                            L.IntEffDate,
		                            L.MatDate,	
		                            L.GLCode,
		                            L.GLCodeOrig,	
		                            L.InstNo,
		                            L.FreqType,
		                            case 
			                            when l.FreqType = '012' then 'Monthly' --monthly
			                            when l.FreqType = '026' then '2 Week' --every 2 week
			                            when l.FreqType = '052' then 'Weekly' --weekly
			                            when l.FreqType = '013' then '4 Week' --every 4 week
			                            else '' ---************protect in case we add some more product->then we can saw error.
		                            end as [PaymentFrequency],--16
		                            L.PrType,
		                            P.FullDesc as PrName,
		                            L.LNCode1,
		                            L.LNCode2,
		                            L.LNCode3,
		                            L.LNCode4,
		                            L.LNCode5,
		                            L.LNCode6,
		                            L.LNCode7,
		                            L.LNCode8,	
		                            Dbo.GetTermOfLoan(L.Acc) AS TERMS,
		                            -----------
		                            CAST('01/01/1900' AS DATETIME) as LastPaidDate,
		                            CAST( 0 AS NUMERIC(18,2)) as LastPaidAmt,
		                            cast('' as nvarchar(15)) as PreAcc,
		                            CAST( 0 AS NUMERIC(18,2)) as PreDisbAmt,
		                            Cast('01/01/1900' AS DATETIME) as NextPaidDate,
		                            cast('01/01/1900' as datetime) as NextDueWorkDate,
		                            Cast(0 as numeric(18,3)) as NextPaidAmt,
		                            Cast(0 as numeric(18,3)) as IntRemainInSchedule,
		                            cast(0 as numeric(18,3)) as RemainLoanTerm,
		                            cast('01/01/1900' as datetime) as LastPreDueDateOfReportDate,
		                            cast(0 as numeric(18,3)) as BalAmtLastDueOfReportDate,
		                            --cast(0 as numeric(18,3)) as IntAmtAfterPreDueDateOfReportDate,
		                            --cast(0 as numeric(18,3)) as TotalAmtRequiredToClose,
		                            cast(0 as numeric(18,3)) as PrinAmtLastDueDateCollect,
		                            --cast('01/01/1900' as datetime) as LastTransDateDueDate,
		                            -----------
		                            dbo.GetAgeOfLoan(@REPORTDATE,l.Acc) as AgeofLoan,
		                            cast(0 as int) as NoInstLate,
		                            CAST('01/01/1900' AS DATETIME) AS FirstDayLate,
		                            --CAST('' as nvarchar(10) ) as StatusReNewLoan,
		                            -------------------------------------------------------------------------
		                            ---Collection --OK
		                            CAST( 0 AS NUMERIC(18,2)) as TrnIntAmt_CollYear,--reserved for collection
                                    CAST( 0 AS NUMERIC(18,2)) as TrnPriAmt_CollYear,--reserved for collection
                                    CAST( 0 AS NUMERIC(18,2)) as TrnPenAmt_CollYear,--reserved for collection
                                    CAST( 0 AS NUMERIC(18,2)) as TrnChgAmt_CollYear, --reserved for collection
		                            CAST( 0 AS NUMERIC(18,2)) as TrnIntAmt_CollAsOfMonth,--reserved for collection
                                    CAST( 0 AS NUMERIC(18,2)) as  TrnPriAmt_CollAsOfMonth,--reserved for collection
                                    CAST( 0 AS NUMERIC(18,2)) as TrnPenAmt_CollAsOfMonth,--reserved for collection
                                    CAST( 0 AS NUMERIC(18,2)) as TrnChgAmt_CollAsOfMonth, --reserved for collection
		                            CAST( 0 AS NUMERIC(18,2)) as TrnIntAmt_CollDaily,--reserved for collection
                                    CAST( 0 AS NUMERIC(18,2)) as TrnPriAmt_CollDaily,--reserved for collection
                                    CAST( 0 AS NUMERIC(18,2)) as TrnPenAmt_CollDaily,--reserved for collection
                                    CAST( 0 AS NUMERIC(18,2)) as TrnChgAmt_CollDaily, --reserved for collection
		                            ------------------------------------------------------------------------------
		                            ---Interesr Culculation Procedure--OK
		                            CAST( 0 AS NUMERIC(18,2)) as IntOdueAmt_FromIntCalc_Proc,
		                            CAST( 0 AS NUMERIC(18,2)) as IntToDateAmt_FromIntCalc_Proc,--= @IntToDate = @IntOdueAmt ( ARI from begining month To Date ) + @AcrIntAmt (ARI end of Previous month that get from LNACC )
		                            CAST( 0 AS NUMERIC(18,2)) as PenaltyAmtDue_FromIntCalc_Proc,
		                            CAST( 0 AS NUMERIC(18,2)) as TaxAmtDue_FromIntCalc_Proc,
		                            CAST( 0 AS NUMERIC(18,2)) as TrnChagAmtDue_FromIntCalc_Proc,
		                            ------------------------------------------------------------------------------
		                            ----PaidOff, paid off can be normal paid off ( loan Past Maturity then paid off ) or loan that paid off before paturity
		                            CAST( 0 AS NUMERIC(18,2)) as TrnIntAmt_PaidOff,--reserved for collection
                                    CAST( 0 AS NUMERIC(18,2)) as TrnPriAmt_PaidOff,--reserved for collection
                                    CAST( 0 AS NUMERIC(18,2)) as TrnPenAmt_PaidOff,--reserved for collection
                                    CAST( 0 AS NUMERIC(18,2)) as TrnChgAmt_PaidOff, --reserved for collection
		                            CAST('01/01/1900' AS DATETIME)  as PaidOffDate,
		                            CAST('' AS NVARCHAR(30) ) AS STATUS_RENEW,
		                            ----
		                            --Provision code
		                            ShortCode = 
		                            case 
			                            when dbo.GetTermOfLoan(L.acc) >= (select max(TERM) from AmretLoanProvision) then 1
			                            else 0 
		                            end,
		                            MasterField = cast('00000' as nvarchar(10)),
		                            -----------------------------------------------------------------
		                            L.AccStatus,
		                            L.AccStatusDate,
		                            BusType.FullDesc BusType,
		                            CASE 
			                            WHEN L.LNCode1 BETWEEN '100' AND '199' THEN 'Agricultural'
			                            WHEN L.LNCode1 BETWEEN '200' AND '399' THEN 'Trading & Commerce'
			                            WHEN L.LNCode1 BETWEEN '400' AND '499' THEN 'Service' 
			                            WHEN L.LNCode1 BETWEEN '500' AND '599' THEN 'Transportation' 
			                            WHEN L.LNCode1 BETWEEN '600' AND '699' THEN 'Construction' 
			                            WHEN L.LNCode1 >  '700' THEN 'Consumption (Householg/Family)'
		                            END AS BusSector,
		                            Colateral.FullDesc as LoanColateral,
		                            LNCycle.FullDesc as LoanCycle,
		                            /*
		                            guaranter information
		                            */
		                            R.CID as idGuarenter,
		                            ltrim(rtrim(GC.DisplayName)) as GuarenterName,
		                            ltrim(rtrim(GCKh.Name1)) as GuarenterNameKh,		
		                            case 
			                            when GC.GenderType = '001' then 'M' 
			                            when GC.GenderType = '002' then 'F'
			                            else	''
		                            end as GuarenterGender,
		                            case 
			                            when GC.CivilStatusCode = '00D' then 'D'
			                            when GC.CivilStatusCode = '00M' then 'M'
			                            when GC.CivilStatusCode = '00S' then 'S'
			                            when GC.CivilStatusCode = '00W' then 'W'
			                            else ''
		                            end as GuarenterMaritalStatus,
		                            case 
			                            when left(GC.Nid,1) = 'M'  then 'M'--M=Maried Letter
			                            when left(GC.Nid,1) = 'C' then 'C' --C= Civil Servan 
			                            when left(GC.Nid,1) = 'F' then 'F' 
			                            when left(GC.Nid,1) = 'B' then 'B' 
			                            when left(GC.Nid,1) = 'N' and ltrim(rtrim(c.nid)) <> 'N/A' then 'N' 
			                            when left(GC.Nid,1)= 'D' then 'D'
			                            when left(GC.Nid,1)= 'P' then 'P'
			                            when left(GC.Nid,1) = 'R' then 'R'
			                            else 'Unknown'
		                            end	 as Guarenter_IDType_1,--N,F,D,B,G,R... columns 2 in CBC
		                            case 
			                            when left(GC.Nid,1) in ('N','F','R','D','B','C','P','M') and ltrim(rtrim(GC.nid)) <> 'N/A' then  right(ltrim(rtrim(GC.Nid)),len(GC.Nid)-1)
			                            when  left(GC.Nid,1) in ('F','G','R') then
				                            case when CAST(right(ltrim(rtrim(GC.Nid)),len(GC.Nid)-1) AS int) = 0 then 'N/A' 
					                            else right(GC.Nid,len(GC.Nid)-1)
				                            end
			                            else GC.Nid
		                            end  as GuarenterIDNumber_1,--3
		                            GC.BirthDate	as GuarenterDateofBirth,		
		                            isnull(ltrim(rtrim(GC.Mobile1)),'') as GuarenterMobile1,
		                            isnull(ltrim(rtrim(GC.Mobile2)),'') as GuarenterMobile2,
		                            GC.locationCode as GuarenterLocationCode,
		                            /*
		                            Co Guarenter information
		                            */
		                            GCkh.Name2 as CoGuarenerName,
		                            case 
			                            when CKh.Sex = 'Rbus' then 'M' 
			                            when CKh.Sex ='RsI' then 'F'
			                            else ''
		                            end as CoGuarenerGender,
		                            GCkh.Date1 as CoGuarenerDoB,
		                            GCkh.CardType as CoGuarenerIDType,
		                            GCkh.Name6 as CoGuarenerIDNum,
		                            GCkh.RelatedName as CoGuarenerRelativeType,
		                            @REPORTDATE AS REPORTDATE

	                            FROM LNACC L 
		                            LEFT JOIN CIF C ON C.CID = substring(L.ACC,4,6) and c.type = '001'--Client
		                            LEFT JOIN relcid G ON G.relatedcid = C.CID AND G.TYPE = '900'  --MEMBER TO GROUP
		                            LEFT JOIN CIF G2 ON G2.CID = G.CID  --JOIN TO GET GROUP INFORMATION
		                            LEFT JOIN RELCID CO ON CO.RelatedCID = G.CID  	AND CO.TYPE = '499'--GROUP TO CO
		                            LEFT JOIN CIF CO2 ON CO2.CID = CO.CID 
		                            LEFT JOIN PRPARMS P ON P.PrType = L.PrType and P.AppType = 4
		                            LEFT JOIN VTSDCIF CKh on Ckh.CID = c.CID and ckh.Type = '001'--join to get client name & co-borrower in khmer
		                            LEFT JOIN VTSDVillage VIL ON VIL.Code = C.LocationCode
		                            LEFT JOIN VTSDCommune COM ON COM.CODE = VIL.CodeOfCommune
		                            LEFT JOIN VTSDDistrict DIS ON DIS.Code = COM.CodeOfDistrict
		                            LEFT JOIN VTSDProvince PRV ON PRV.Code = DIS.CodeOfProvince
		                            LEFT JOIN VTSDCIF CO_Kh	ON CO_Kh.CID = co2.CID and CO_Kh.Type = '499'
		                            inner join Trnhist T ON T.ACC = L.ACC 	
		                            LEFT JOIN USERLOOKUP BusType ON BusType.LookUpCode = L.LNCode1 and BusType.LookUpId = '41'	
		                            LEFT JOIN USERLOOKUP Colateral ON Colateral.LookUpCode = L.LNCode2 and Colateral.LookUpId = '42'
		                            LEFT JOIN USERLOOKUP LNCycle on  LNCycle.LookUpCode = L.LNCode3 and LNCycle.LookUpId = '43'
		                            LEFT JOIN USERLOOKUP CIF_FamilyMember ON  CIF_FamilyMember.LookUpCode = C.CIFCode1 and CIF_FamilyMember.LookUpId = '61'--62
		                            LEFT JOIN USERLOOKUP CIF_Income ON  CIF_Income.LookUpCode = C.CIFCode2 and CIF_Income.LookUpId = '62'
		                            LEFT JOIN USERLOOKUP CIF_Occupation ON  CIF_Occupation.LookUpCode = C.CIFCode3 and CIF_Occupation.LookUpId = '63'
		                            LEFT JOIN USERLOOKUP CIF_TotalAsset ON  CIF_TotalAsset.LookUpCode = C.CIFCode4 and CIF_TotalAsset.LookUpId = '64'
		                            LEFT JOIN relacc R on R.Acc + R.Chd= L.Acc + L.Chd and R.Type  = '030' and R.AppType = '4'--type = 30 for guarentee
		                            LEFT JOIN CIF GC ON GC.CID = R.CID--guaretee
		                            LEFT JOIN VTSDCIF GCKh on GCKh.CID = GC.CID and gckh.Type = '001'--join to get client name & co-borrower in khmer
	                            WHERE 
		                            L.AccStatus = '99' --CLOSE LOAN IN MONTH ONLY
		                            and T.ValueDate between @BeginDateMonth and @REPORTDATE
		                            --and T.ValueDate between @BeginDateYear and @REPORTDATE
		                            and ISNULL(T.TrnPriAmt,0) +	isnull(T.TrnIntAmt,0) + isnull(T.TrnPenAmt,0) + ISNULL(T.TrnChgAmt,0)  >0
		                            and T.BalAmt = 0 ---If loan pay many time i

                            --=============================================================================================================================================================
                              --Update Client No of installment Late
                               UPDATE #LN_DETAIL
		                            SET NoInstLate = (SELECT COUNT(*) FROM LNINST I WHERE I.Status ='1' AND I.Acc = L.Acc),
		                             FirstDayLate = (SELECT TOP 1 DueDate FROM LNINST I WHERE I.Status ='1' AND I.Acc = L.Acc ORDER BY DueDate)
                               FROM #LN_DETAIL L
                               WHERE L.AgeofLoan <>0  

                               --Update Status Re-New Loan TO CLOSE LOAN IN MONTH
                               UPDATE #LN_DETAIL
		                            SET STATUS_RENEW=
			                            CASE 
				                            WHEN (SELECT TOP 1 ISNULL(L1.IdClient,0)  FROM #LN_DETAIL L1 WHERE L1.IdClient = L.IdClient AND L1.ACCSTATUS <>'99') > 0 then 'RENEW' 
					                            ELSE 'NOTRENEW' 
				                            END
                               FROM #LN_DETAIL  L
                               WHERE L.accstatus='99'
                              --============================================================================================================================================================ 
                            --/*
                            --Get Repayment from Transaction history
                            --*/
                               SELECT 
		                              LTRIM(RTRIM(a.Acc)) AS Acc,
                                      SUM(t.TrnIntAmt/@ccydiv)AS TrnIntAmt_CollAsOfMonth,
                                      SUM(t.TrnPriAmt/@ccydiv) AS TrnPriAmt_CollAsOfMonth,
                                      SUM(t.TrnPenAmt/@ccydiv) AS TrnPenAmt_CollAsOfMonth,
                                      SUM(t.TrnChgAmt/@ccydiv) AS TrnChgAmt_CollAsOfMonth,
		                              CAST( 0 AS NUMERIC(18,2)) as TrnIntAmt_CollDaily,
		                              CAST( 0 AS NUMERIC(18,2)) as TrnPriAmt_CollDaily,
		                              CAST( 0 AS NUMERIC(18,2)) as TrnPenAmt_CollDaily,
		                              CAST( 0 AS NUMERIC(18,2)) as TrnChgAmt_CollDaily  		     
                               INTO #RepayMonth_Temp
                               FROM TrnHist T,#LN_DETAIL a 
                               WHERE t.acc=a.Acc AND t.TrnType in ('401','403','405','411','421','431','451','453') AND 
                                     t.GLCode NOT LIKE 'W%' AND t.CancelledByTrn IS NULL AND 
                                     (t.ValueDate BETWEEN @BeginDateMonth AND @REPORTDATE) 
		 
                               GROUP BY a.Acc

                               INSERT INTO #RepayMonth_Temp
                                SELECT 
		                              LTRIM(RTRIM(a.Acc)) AS Acc, 
                                      0 AS TrnIntAmt_CollAsOfMonth,
                                      0 AS TrnPriAmt_CollAsOfMonth,
                                      0 AS TrnPenAmt_CollAsOfMonth,
                                      0 AS TrnChgAmt_CollAsOfMonth,
		                              SUM(t.TrnIntAmt/@ccydiv) as TrnIntAmt_CollDaily,
		                              SUM(t.TrnPriAmt/@ccydiv) as TrnPriAmt_CollDaily,
		                              SUM(t.TrnPenAmt/@ccydiv) as TrnPenAmt_CollDaily,
		                              SUM(t.TrnChgAmt/@ccydiv) as TrnChgAmt_CollDaily
                               FROM TrnHist T,#LN_DETAIL a 
                               WHERE t.acc=a.Acc AND t.TrnType in ('401','403','405','411','421','431','451','453') AND 
                                     t.GLCode NOT LIKE 'W%' AND t.CancelledByTrn IS NULL AND 
                                     (t.ValueDate = @REPORTDATE) 
                               GROUP BY a.Acc
                               ------------------------------------------------------------------------   
                               INSERT INTO #RepayMonth_Temp
                               SELECT LTRIM(RTRIM(a.Acc)) AS Acc, 
                                      SUM(t.TrnIntAmt/@ccydiv)AS TrnIntAmt_CollAsOfMonth,
                                      SUM(t.TrnPriAmt/@ccydiv) AS TrnPriAmt_CollAsOfMonth,
                                      SUM(t.TrnPenAmt/@ccydiv) AS TrnPenAmt_CollAsOfMonth,
                                      SUM(t.TrnChgAmt/@ccydiv) AS TrnChgAmt_CollAsOfMonth,
		                              0 as TrnIntAmt_CollDaily,
		                              0 as TrnPriAmt_CollDaily,
		                              0 as TrnPenAmt_CollDaily,
		                              0 as TrnChgAmt_CollDaily
                               FROM TrnDaily T,#LN_DETAIL a 
                               WHERE t.acc=a.Acc AND t.TrnType in ('401','403','405','411','421','431','451','453') AND 
                                     t.GLCode NOT LIKE 'W%' AND t.CancelledByTrn IS NULL AND 
                                     (t.ValueDate BETWEEN @BeginDateMonth AND @REPORTDATE) 
                               GROUP BY a.Acc
                               union ---Current Day
                                  SELECT LTRIM(RTRIM(a.Acc)) AS Acc, 
                                      0 AS TrnIntAmt_CollAsOfMonth,
                                      0 AS TrnPriAmt_CollAsOfMonth,
                                      0 AS TrnPenAmt_CollAsOfMonth,
                                      0 AS TrnChgAmt_CollAsOfMonth,
		                              SUM(t.TrnIntAmt/@ccydiv) as TrnIntAmt_CollDaily,
		                              SUM(t.TrnPriAmt/@ccydiv) as TrnPriAmt_CollDaily,
		                              SUM(t.TrnPenAmt/@ccydiv) as TrnPenAmt_CollDaily,
		                              SUM(t.TrnChgAmt/@ccydiv) as TrnChgAmt_CollDaily
                               FROM TrnDaily T,#LN_DETAIL a 
                               WHERE t.acc=a.Acc AND t.TrnType in ('401','403','405','411','421','431','451','453') AND 
                                     t.GLCode NOT LIKE 'W%' AND t.CancelledByTrn IS NULL AND 
                                     (t.ValueDate = @REPORTDATE) 
                               GROUP BY a.Acc
                               --------Reverse Closing ----------------------
                               INSERT INTO #RepayMonth_Temp 
                               SELECT LTRIM(RTRIM(a.Acc)) AS Acc, 
                                      (-1)*SUM(t.TrnIntAmt/@ccydiv)AS TrnIntAmt_CollAsOfMonth,
                                      (-1)*SUM(t.TrnPriAmt/@ccydiv) AS TrnPriAmt_CollAsOfMonth,
                                      (-1)*SUM(t.TrnPenAmt/@ccydiv) AS TrnPenAmt_CollAsOfMonth,
                                      (-1)*SUM(t.TrnChgAmt/@ccydiv) AS TrnChgAmt_CollAsOfMonth,  
		                              0 as TrnIntAmt_CollDaily,
		                              0 as TrnPriAmt_CollDaily,
		                              0 as TrnPenAmt_CollDaily,
		                              0 as TrnChgAmt_CollDaily
                               FROM TrnHist T,#LN_DETAIL a 
                               WHERE t.acc=a.Acc AND t.TrnType IN('105','205','305','406','705') AND t.GLCode NOT LIKE 'W%'
                                     AND t.CancelledByTrn IS NULL 
		                             AND (t.ValueDate BETWEEN @BeginDateMonth AND @REPORTDATE) 
                               GROUP BY a.Acc     
                               UNION
                               SELECT
                               LTRIM(RTRIM(a.Acc)) AS Acc, 
                                      0 AS TrnIntAmt_CollAsOfMonth,
                                      0 AS TrnPriAmt_CollAsOfMonth,
                                      0 AS TrnPenAmt_CollAsOfMonth,
                                      0 AS TrnChgAmt_CollAsOfMonth,  
		                              (-1)*SUM(t.TrnIntAmt/@ccydiv) as TrnIntAmt_CollDaily,
		                              (-1)*SUM(t.TrnPriAmt/@ccydiv) as TrnPriAmt_CollDaily,
		                              (-1)*SUM(t.TrnPenAmt/@ccydiv) as TrnPenAmt_CollDaily,
		                              (-1)*SUM(t.TrnChgAmt/@ccydiv) as TrnChgAmt_CollDaily
                               FROM TrnHist T,#LN_DETAIL a 
                               WHERE t.acc=a.Acc AND t.TrnType IN('105','205','305','406','705') AND t.GLCode NOT LIKE 'W%'
                                     AND t.CancelledByTrn IS NULL 
		                             AND (t.ValueDate = @REPORTDATE) 
                               GROUP BY a.Acc   
      
                               INSERT INTO #RepayMonth_Temp 
                               SELECT LTRIM(RTRIM(a.Acc)) AS Acc, 
                                      (-1)*SUM(t.TrnIntAmt/@ccydiv)AS TrnIntAmt_CollAsOfMonth,
                                      (-1)*SUM(t.TrnPriAmt/@ccydiv) AS TrnPriAmt_CollAsOfMonth,
                                      (-1)*SUM(t.TrnPenAmt/@ccydiv) AS TrnPenAmt_CollAsOfMonth,
                                      (-1)*SUM(t.TrnChgAmt/@ccydiv) AS TrnChgAmt_CollAsOfMonth,
		                              0 as TrnIntAmt_CollDaily,
		                              0 as TrnPriAmt_CollDaily,
		                              0 as TrnPenAmt_CollDaily,
		                              0 as TrnChgAmt_CollDaily 
                               FROM TrnDaily T,#LN_DETAIL a 
                               WHERE t.acc=a.Acc AND t.TrnType IN('105','205','305','406','705') AND t.GLCode NOT LIKE 'W%'
                                     AND t.CancelledByTrn IS NULL AND (t.ValueDate BETWEEN @BeginDateMonth AND @REPORTDATE) 
                               GROUP BY a.Acc
                                UNION
                               SELECT
                               LTRIM(RTRIM(a.Acc)) AS Acc, 
                                      0 AS TrnIntAmt_CollAsOfMonth,
                                      0 AS TrnPriAmt_CollAsOfMonth,
                                      0 AS TrnPenAmt_CollAsOfMonth,
                                      0 AS TrnChgAmt_CollAsOfMonth,  
		                              (-1)*SUM(t.TrnIntAmt/@ccydiv) as TrnIntAmt_CollDaily,
		                              (-1)*SUM(t.TrnPriAmt/@ccydiv) as TrnPriAmt_CollDaily,
		                              (-1)*SUM(t.TrnPenAmt/@ccydiv) as TrnPenAmt_CollDaily,
		                              (-1)*SUM(t.TrnChgAmt/@ccydiv) as TrnChgAmt_CollDaily
                               FROM TrnDaily T,#LN_DETAIL a 
                               WHERE t.acc=a.Acc AND t.TrnType IN('105','205','305','406','705') AND t.GLCode NOT LIKE 'W%'
                                     AND t.CancelledByTrn IS NULL AND (t.ValueDate = @REPORTDATE) 
                               GROUP BY a.Acc   

                               --Clear Reverse Closing ----------------------
                               SELECT Acc, 
                                      SUM(TrnIntAmt_CollAsOfMonth) AS TrnIntAmt_CollAsOfMonth,
                                      SUM(TrnPriAmt_CollAsOfMonth) AS TrnPriAmt_CollAsOfMonth,
                                      SUM(TrnPenAmt_CollAsOfMonth) AS TrnPenAmt_CollAsOfMonth,
                                      SUM(TrnChgAmt_CollAsOfMonth) AS TrnChgAmt_CollAsOfMonth,
		                              SUM(TrnIntAmt_CollDaily) AS TrnIntAmt_CollDaily,
		                              SUM(TrnPriAmt_CollDaily) AS TrnPriAmt_CollDaily,
		                              SUM(TrnPenAmt_CollDaily) AS TrnPenAmt_CollDaily,
		                              SUM(TrnChgAmt_CollDaily) AS TrnChgAmt_CollDaily
                               INTO #RepayMonth
                               FROM #RepayMonth_Temp 
                               GROUP BY Acc
                                /*
                               Part 2
                               */
                            ----Update all loan repayment----------------------
                               UPDATE #LN_DETAIL 
	                            SET TrnIntAmt_CollAsOfMonth=a.TrnIntAmt_CollAsOfMonth,
                                   TrnPriAmt_CollAsOfMonth = a.TrnPriAmt_CollAsOfMonth,
                                   TrnPenAmt_CollAsOfMonth=a.TrnPenAmt_CollAsOfMonth,
                                   TrnChgAmt_CollAsOfMonth=a.TrnChgAmt_CollAsOfMonth,
	                               TrnIntAmt_CollDaily = a.TrnIntAmt_CollDaily,
	                               TrnPriAmt_CollDaily = a.TrnPriAmt_CollDaily,
	                               TrnPenAmt_CollDaily = a.TrnPenAmt_CollDaily,
	                               TrnChgAmt_CollDaily = a.TrnChgAmt_CollDaily
                               FROM #LN_DETAIL l,#RepayMonth a 
                               WHERE l.acc=a.acc    
                            --=========================================================================================================================================================
                            /*
                            Culculate Interest, to get InterestToDate for only active loan
                            */
                            --Create other parameter--
	                            DECLARE @FutureDays			INT
	                            DECLARE	@PenAmt				NUMERIC(18,3)     
	                            DECLARE	@IntODueAmt			NUMERIC(18,3)
	                            DECLARE	@TaxAmt				NUMERIC (18,3)
	                            DECLARE	@TrnChgAmt			NUMERIC(18,3)
	                            DECLARE @IntAmt				NUMERIC(18,3)
	                            SET @FutureDays		=0
	                            SET @IntAmt			=0
	                            SET @PenAmt			=0
	                            SET @IntODueAmt		=0
	                            SET @TaxAmt			=0
	                            SET @TrnChgAmt		=0

	                            --Exec   sp_LNCalcInterest  @idAccount, @FutureDays, @IntAmt Output, @PenAmt Output, @IntODueAmt Output, @TaxAmt Output, @TrnChgAmt Output

	                            Create table #LNCalcInterest
	                            ( 
		                            Acc_Full nvarchar(15),
		                            IntOdueAmt numeric(18,3),
		                            IntToDateAmt numeric(18,3),--= @IntToDate = @IntOdueAmt ( ARI from begining month To Date ) + @AcrIntAmt (ARI end of Previous month that get from LNACC )
		                            PenaltyAmt numeric(18,3),--Penalty Amount Due
		                            TaxAmt numeric(18,3),
		                            TrnChagAmt numeric(18,3)
	                            )
	                            /*
	                            loop from each row to culculate interest for each acc
	                            */	
	                            declare @AcrIntAmt_EndPreMonth numeric(18,3)
	                            DECLARE @ACC_FULL NVARCHAR(15)
	                            DECLARE MyCursor CURSOR FOR
		                            select 
			                            L.Acc+L.Chd,
			                            L.AcrIntAmt_EndPreMonth	
		                            from #LN_DETAIL L
		                            where L.AccStatus between '11' and '98'
		                            OPEN	MyCursor
		                            FETCH	MyCursor INTO @ACC_FULL,@AcrIntAmt_EndPreMonth
		                            WHILE @@FETCH_STATUS <> -1
		                            BEGIN
			                            Exec   sp_LNCalcInterest  @ACC_FULL, @FutureDays, @IntAmt Output, @PenAmt Output, @IntODueAmt Output, @TaxAmt Output, @TrnChgAmt Output
			                            INSERT INTO #LNCalcInterest (Acc_Full,IntOdueAmt,IntToDateAmt,PenaltyAmt,TaxAmt,TrnChagAmt)
			                            SELECT 
				                            @ACC_FULL,
				                            @IntODueAmt,
				                            @AcrIntAmt_EndPreMonth + @IntAmt,--	or @AcrIntAmt_EndPreMonth + @IntODueAmt,
				                            @PenAmt,
				                            @TaxAmt,
				                            @TrnChgAmt			
			                            FETCH NEXT FROM MyCursor INTO @ACC_FULL,@AcrIntAmt_EndPreMonth
		                            END
	                            CLOSE MyCursor
	                            DEALLOCATE  MyCursor
	                            ---update ineterest calculation to temp table of laon ( only active )
	                            UPDATE #LN_DETAIL 
	                            SET 
		                            IntOdueAmt_FromIntCalc_Proc = c.IntOdueAmt,
		                            IntToDateAmt_FromIntCalc_Proc = c.IntToDateAmt,--= @IntToDate = @IntOdueAmt ( ARI from begining month To Date ) + @AcrIntAmt (ARI end of Previous month that get from LNACC )
		                            PenaltyAmtDue_FromIntCalc_Proc = c.PenaltyAmt,
		                            TaxAmtDue_FromIntCalc_Proc = c.TaxAmt,
		                            TrnChagAmtDue_FromIntCalc_Proc = c.TrnChagAmt
                               FROM #LN_DETAIL L,#LNCalcInterest c
                               WHERE l.acc + l.Chd = c.Acc_Full  and L.AccStatus between '11' and '98'

                            --=========================================================================================================================================================
                            /*
                            Paid Of Loan during this months.
                            */
                            select 
	                            cast(co2.CID as int) as idCO,	
	                            co2.DisplayName as CoName,
	                            L.acc,
	                            dbo.[GetFirstDisbDate](L.acc) as DisbDate,
	                            L.PrType,
	                            L.GrantedAmt,
	                            L.AccStatus,
	                            T.TrnPriAmt,
	                            T.TrnIntAmt,
	                            T.TrnPenAmt,
	                            T.TrnChgAmt,
	                            T.BalAmt,
	                            t.TrnDate,
	                            T.TrnType,
	                            T.TrnDesc,
	                            T.ValueDate,
	                            L.MatDate,
	                            @REPORTDATE as ReportDate
                             into #PaidOff
                             from LNACC L
	                            left join CIF c on c.CID = SUBSTRING(l.Acc,4,6) 
	                            left join RELCID g on g.RelatedCID = c.CID and g.Type = '900' --MEMBER TO GROUP
	                            LEFT JOIN RELCID CO ON CO.RelatedCID = G.CID  	AND CO.TYPE = '499'--GROUP TO CO
	                            LEFT JOIN CIF CO2 ON CO2.CID = CO.CID --Join to get CO Name
	                            inner join Trnhist T ON T.ACC = L.ACC 
	                            where L.AccStatus = '99' 
		                            --and L.AccStatusDate between @BeginDateMonth and @REPORTDATE
		                            and T.ValueDate between @BeginDateMonth and @REPORTDATE
		                            --and L.MatDate > T.TrnDate
		                            and ISNULL(T.TrnPriAmt,0) +	isnull(T.TrnIntAmt,0) + isnull(T.TrnPenAmt,0) + ISNULL(T.TrnChgAmt,0)  >0
		                            and T.BalAmt = 0 ---If loan pay many time in this month, only get the last one that paid until balance = 0

                            /*
                            Add more loan for current day in table TrnDaily, 
                            then all transaction not yet moved to TrnHist Table
                            */
                            insert into #PaidOff
                            select 
	                            cast(co2.CID as int) as idCO,	
	                            co2.DisplayName as CoName,
	                            L.acc,
	                            dbo.[GetFirstDisbDate](L.acc) as DisbDate,
	                            L.PrType,
	                            L.GrantedAmt,
	                            L.AccStatus,
	                            T.TrnPriAmt,
	                            T.TrnIntAmt,
	                            T.TrnPenAmt,
	                            T.TrnChgAmt,
	                            T.BalAmt,
	                            T.TrnDate,
	                            T.TrnType,
	                            T.TrnDesc,
	                            T.ValueDate,
	                            L.MatDate,
	                            @REPORTDATE as ReportDate
                             from LNACC L
	                            left join CIF c on c.CID = SUBSTRING(l.Acc,4,6) 
	                            left join RELCID g on g.RelatedCID = c.CID and g.Type = '900' --MEMBER TO GROUP
	                            LEFT JOIN RELCID CO ON CO.RelatedCID = G.CID  	AND CO.TYPE = '499'--GROUP TO CO
	                            LEFT JOIN CIF CO2 ON CO2.CID = CO.CID --Join to get CO Name
	                            inner join TRNDAILY T ON T.ACC = L.ACC 
	                            where L.AccStatus = '99' 
		                            --and L.AccStatusDate between @BeginDateMonth and @REPORTDATE
		                            and T.ValueDate between @BeginDateMonth and @REPORTDATE
		                            --and L.MatDate > T.TrnDate
		                            and ISNULL(T.TrnPriAmt,0) +	isnull(T.TrnIntAmt,0) + isnull(T.TrnPenAmt,0) + ISNULL(T.TrnChgAmt,0)  >0
		                            and T.BalAmt = 0 ---

                            update #LN_DETAIL 
	                            set 
		                            TrnPriAmt_PaidOff = P.TrnPriAmt,
		                            TrnIntAmt_PaidOff = P.TrnIntAmt,
		                            TrnChgAmt_PaidOff = P.TrnChgAmt,
		                            TrnPenAmt_PaidOff = P.TrnPenAmt,
		                            PaidOffDate = P.ValueDate
                            from #LN_DETAIL L1, #PaidOff P 
                            WHERE P.Acc = L1.Acc AND L1.AccStatus = '99'
                            -------------------------------------------
                             ---Update provision
                            update #LN_DETAIL 
	                            set 
	                            MasterField = 
		                            case 
			                            when L1.ShortCode = 0  then
				                            case 
					                            when L1.AgeofLoan =0 then '0000' --auto normal loan
					                            when L1.AgeofLoan between 1 and  @MaxDayShort  then --@MaxDayShort = 90 days --> lost 
						                            (select 
							                            R.code	+'0'			
							                            from AmretLoanProvision R
							                            where R.term = 0 --0 for short term
							                            and R.NumDaysNo = (
												                            select max(R2.NumDaysNo) from AmretLoanProvision R2 
													                            where R2.Term  = 0 
													                            and  R2.NumDaysNo<=L1.AgeofLoan
											                               )
						                            )
					                            when L1.AgeofLoan > @MaxDayShort then '0040' --auto loss	
					                            else 'UnKnown'				
				                            end			
			                            when L1.ShortCode = 1 then
				                            case 					
					                            when L1.AgeofLoan =0 then '0001' --normal
					                            when L1.AgeofLoan between 1 and @MaxDayLong then --long terms
						                            (
						                            select 
							                             R.code +'1'	--1 for long term			
						                            from AmretLoanProvision R
							                            where R.term = @TermDefined---@TermDefined=366 day and up is long term 
								                            and R.NumDaysNo = (
													                            select max(R2.NumDaysNo) 
													                            from AmretLoanProvision R2 
														                            where R2.Term  = @TermDefined
															                            and R2.NumDaysNo<=L1.AgeofLoan
													                            )
												                            )
					                            when L1.AgeofLoan > @MaxDayLong then '0041' --auto loss
				                            end
		                            end
                              from #LN_DETAIL L1
                            ----Update LastPaidDate, LastPaid Amt
                            update #LN_DETAIL 
	                            set LastPaidDate = (select top(1) max(t2.TrnDate) from TRNHIST T2 where t2.Acc = L.Acc and t2.TrnType in ('451','431','401')),--'401','403','405','411','421','431','451','453'
	                            LastPaidAmt = (
				                            select isnull(sum(T1.TrnAmt) ,0)
				                            from TRNHIST T1
					                            where 
					                            T1.Acc = L.Acc 
					                            and T1.TrnDate = (select top(1) max(T3.TrnDate) from TRNHIST T3 where t3.Acc = L.acc and t3.trnType in ('451','431','401'))
					                            and T1.trnType in ('451','431','401')
				                             ),
	                            PreDisbAmt = (
		                            Select L2.GrantedAmt 
		                            from LNACC L2 
		                            where substring(L2.acc,4,6) = L.IdClient 
			                              and L2.Opendate = 
				                            (
					                            select max(Opendate) as Opendate from LNACC L3 where substring(L3.Acc,4,6)=L.IdClient and L3.AccStatus='99' group by substring(L3.Acc,4,6)
				                            ) 			  
			                              and L2.AccStatus='99'
			                            ),
	                            PreAcc = (
			                            Select L4.Acc from LNACC L4 
			                            where L4.Opendate = 
			                            (
				                            select max(L5.Opendate) as Opendate 
				                            from LNACC L5 where substring(L5.Acc,4,6)=L.IdClient and L5.AccStatus='99' group by SUBSTRING(L5.acc,4,6)
			                            ) and substring(L4.acc,4,6) = L.IdClient and L4.AccStatus='99'
		                            ),
	                            NextPaidDate = (
		                            select distinct(Min(I.DueDate)) from LNINST I where I.Acc = L.Acc and I.Chd = L.chd and  I.Status = 0 and I.status <> 8
	                            ),
	                            IntRemainInSchedule = (
		                            select 
			                            sum(I.IntAmt/@CCYDIV) 
		                            from LNINST I 
		                            inner join LNACC L6 on L6.Acc = I.Acc and L6.Chd = I.Chd and I.Acc = L.acc and I.Chd = L.Chd
		                            inner join RELACC R on R.ACC = L.Acc and R.Chd = L.Chd
		                            where I.Status <> 8 and R.AppType = '4' and r.Type = '010'
		                            ),
		                            LastPreDueDateOfReportDate = 
		                            (
			                            select distinct(max(T1.DueDate)) from LNINST T1 where t1.Acc = L.Acc and t1.Status <> 8
					                            and T1.DueDate <= @ReportDate		
		                            )

                            from #LN_DETAIL L 

                            ---Protect loan disburst but never paid
                            update #LN_DETAIL 
	                            set LastPreDueDateOfReportDate = 
		                            (
			                            case 
				                            when LastPreDueDateOfReportDate is null then L3.DisbDate
				                            else LastPreDueDateOfReportDate
			                            end
		                            )
                            from #LN_DETAIL L3

                            update #LN_DETAIL 
	                            set NextPaidAmt = (
					                            select 
						                            case 
							                            when L2.NextPaidDate is null then null
							                            else isnull(I.PriAmt/@CCYDIV,0) + isnull(I.IntAmt/@ccydiv,0) + isnull(I.ChargesAmt/@ccydiv,0) 
						                            end 
					                            from LNINST I 
						                            inner join LNACC L6 on L6.Acc = I.Acc and L6.Chd = I.Chd and I.Acc = L2.acc and I.Chd = L2.Chd
						                            inner join RELACC R on R.ACC = L2.Acc and R.Chd = L2.Chd
					                            where --I.Status <> 8  ---reschedule
					                            I.Status = 0
					                            and R.AppType = '4' and r.Type = '010'
					                            and i.DueDate = L2.NextPaidDate
						                            ),
		                            NextDueWorkDate = dbo.GetWorkingDueDate(NextPaidDate),	
	                                BalAmtLastDueOfReportDate = (
			                            select top(1)
				                            case 
					                            when (I3.BFBalAmt/@CcyDiv - I3.OrigPriAmt/@CcyDiv = 0) then I3.BFBalAmt/@CcyDiv
					                            when (I3.BFBalAmt/@CcyDiv - I3.OrigPriAmt/@CcyDiv >0 ) then (I3.BFBalAmt/@CcyDiv- I3.OrigPriAmt/@CcyDiv)
				                            end as AA		
			                            from LNINST I3
				                            where I3.Acc = L2.Acc  and I3.Status <> '8'
				                            and I3.DueDate = LastPreDueDateOfReportDate		
			                            ),
		                            PrinAmtLastDueDateCollect = (
			                            select 
				                            top(1) I4.OrigPriAmt/@ccydiv 
			                            from LNINST I4 
				                            where acc = L2.Acc 
				                            and I4.DueDate = L2.LastPreDueDateOfReportDate 
				                            and I4.Status <> '8'
		                            ),
		                            RemainLoanTerm = (
			                            select 
				                            count(distinct(I5.DueDate))							
			                            from LNINST I5
				                            inner join LNACC L6 on L6.Acc = I5.Acc and L6.Chd = I5.Chd and I5.Acc = L2.acc and I5.Chd = L2.Chd
				                            inner join RELACC R on R.ACC = L2.Acc and R.Chd = L2.Chd
					                            where --I.Status <> 8  ---reschedule
						                            I5.Status = 0
						                            and R.AppType = '4' and r.Type = '010'
						                            and I5.PaidDate is null
					                            --and i.DueDate = L2.NextPaidDate
		                            )
                            from #LN_DETAIL L2 
	


                            --==========================================================================================================================================================
                            ---Replace string ',' with ';' for CSV
                            declare @ColumnName nvarchar(50)
                            declare xCursor cursor for
	                            SELECT 
		                            distinct(sc.name)
	                            --st.name as type_name 
	                            --sc.max_length 
	                            FROM tempdb.sys.columns sc inner join sys.types st on st.system_type_id=sc.system_type_id 
		                            WHERE [object_id] = OBJECT_ID('tempdb..#LN_DETAIL')
		                            and st.name in ('char','nchar','ntext','nvarchar','text','varchar') 
                            OPEN	xCursor
                            FETCH	xCursor INTO @ColumnName
                            WHILE @@FETCH_STATUS <> -1
                            BEGIN
	                            PRINT @ColumnName
	                            EXECUTE (' UPDATE #LN_DETAIL 
	                            SET  '  + @ColumnName + ' = REPLACE( '  + @ColumnName + ','','','+''';'')' )
	                            FETCH NEXT FROM xCursor INTO @ColumnName
                            END
                            CLOSE xCursor
                            DEALLOCATE xCursor


                            --select * into [DailyReports].dbo.LN_DETAIL from #LN_DETAIL
                            select 
				                row_number() over(order by Acc) as ID,
				                @BrCode as BrCode,
				                @BrShort as BrShort,
				                Acc+chd as AccountNumber,
				                IdCo,
				                CoName,
				                CoNameKh,
				                ClientName,
				                ClientNameKh,
				                Gender,
				                CoBorrowerName,
				                CoBorrowerGender,
				                Mobile1,
				                AdressKhmer,
				                GrantedAmt,
				                Balamt,
				                DisbDate,
				                MatDate,
				                IntBalAmt,
				                PenBalAmt,
				                OduePriAmt,
				                OdueIntAmt,
				                PaymentFrequency,
				                LastPaidDate,
				                Ageofloan,
				                Reportdate

				                from #LN_DETAIL where accstatus between '11' and '98'


                            --select * into LN_DETAIL from #LN_DETAIL

                            TRUNCATE TABLE #LN_DETAIL
                            truncate table #PaidOff
                            TRUNCATE TABLE #LNCalcInterest

                            --DROP TABLE #LN_DETAIL
                            drop table #PaidOff	
                            drop table #RepayMonth_Temp
                            drop table #RepayMonth
                            DROP TABLE #LNCalcInterest

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
                    if (re.download == "download")
                    {
                        string ReportName = "Loan Overdue Reports";
                        string BrName = Convert.ToString(mutility.dbSingleResult("select BrLetter from BRANCH_LISTS where flag=1 and BrCode='" + re.BrCode + "'"));
                        ExportDataToExcel(downloadLoanStatement, BrName + "_" + ReportName);
                    }


                    return View("LoanOverdue");
                }
                else
                {
                    return RedirectToAction("Index", "Login/Index");
                }
            }

           
        }
        [HttpPost]
        //Exec LoanOverdue
        public ActionResult execfulltrialbalance(Reports re)
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

                    DateTime datestart = re.datestart;
                    DateTime dateend = re.dateend;

                    string Sql = @"---------------------------------------
                            declare @BrCode nvarchar(10)
                            declare @BrShort nvarchar(10)
                            select
	                            @BrCode = SubBranchCode,
	                            @BrShort = SubBranchID
                            from SKP_Brlist where DBName = (SELECT left(DB_NAME(),len(DB_NAME())-4))

                            Declare @CcyDiv int
                            set @CcyDiv = (Select CcyDiv from CCY)

                            Declare @Code char(3)
                            set @Code = (Select Code from CCY)
                            if @code='KHR' 
	                            set @code=1
                            else if @code='USD' 
	                            set @code=2
                            else if @code='THB'
	                            set @code=5
                            Declare @Code1 datetime
                            set @Code1 = (Select currrundate from BRPARMS)
                            Select @BrCode as BrCode,@BrShort as BrName,GLAcc as 'Acc Number', FullTitle as 'Acc Name',
                            CBalAmt/@CcyDiv as Balance,
                            --@BrCode as 'Br.Code', 
                            @Code as CCY ,@Code1 as Date 
                            from GLAC where
                            CBalAmt !=0 order by GLAcc
                            --------------------------------------------

                            select CurrRunDate,RunStatus from brparms
                            --cast(cast(@BrCode as int)as nvarchar(10))";


                    ViewBag.BrCode = re.BrCode;
                    ViewBag.datestart = re.datestart;
                    ViewBag.dateend = re.dateend;
                    ViewBag.Acc = re.accountnumber;

                    ViewBag.CurrRunDate = CurrRunDate.CurrRunDate(userkey);
                    ViewBag.dt_Report = mutility.dtOnlyOneBranch(Sql, re.BrCode);
                    downloadLoanStatement = mutility.dtOnlyOneBranch(Sql, re.BrCode);
                    BrCode = re.BrCode;
                    if (re.download == "download")
                    {
                        string ReportName = "Full Trail Balance Reports";
                        string BrName = Convert.ToString(mutility.dbSingleResult("select BrLetter from BRANCH_LISTS where flag=1 and BrCode='" + re.BrCode + "'"));
                        ExportDataToExcel(downloadLoanStatement, BrName + "_" + ReportName);
                    }


                    return View("fulltrialbalance");
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

        //Exec GLTransactionbybatch
        public ActionResult execGLTransactionbybatch(Reports re)
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

                    DateTime datestart = re.datestart;
                    DateTime dateend = re.dateend;

                    string Sql = @"	/*
	                        Generated by : HEM Loeurt
	                        Date : 8-Jan-2015
	                        To Get List of GL Transaction by batch, this script should run on database after EOS only
	                        */	
	                        set dateformat DMY	
	                        declare @DateFrom datetime
	                        declare @DateTo datetime
							declare @BrCode nvarchar(10)
							declare @BrShort nvarchar(10)
							select
								@BrCode = SubBranchCode,
								@BrShort = SubBranchID
							from SKP_Brlist where DBName = (SELECT left(DB_NAME(),len(DB_NAME())-4))
	                        declare @CcyDive int
	                        set @CcyDive = (select CcyDiv from CCY)

	                        ---------------------------------------------------------------------------------
	                        set @DateFrom = '"+ datestart + @"'
	                        set @DateTo = '"+ dateend + @"'
                            ---------------------------------------------------------------------------------	
	                        /*
	                        CLEAN TEMP TABLE
	                        */
	                        IF OBJECT_ID('tempdb..#GLTRNHIST') IS NOT NULL
		                        DROP TABLE #GLTRNHIST
	                        IF OBJECT_id('tempdb..#BATCH_REFF') IS NOT NULL
		                        DROP TABLE #BATCH_REFF
                           ---------------------------------------------------------------------------------- 	
	                        SELECT
		                        Distinct(Ref),
		                        JnlDate,
		                        JnlNumber
	                        INTO #BATCH_REFF
	                        FROM GLJLHist 
	                        WHERE JnlDate Between @DateFrom and @DateTo
		                        AND Status='0' AND JnlNumber Not In ('EOD', 'EOM')
		                        AND LineNumber='00001'
	                        SELECT		
		                        g.Type as TrnJnlType,
		                        h.jnlType as HISJnlType,
		                        JnlTrn=Case When g.Type='50' Then 'BTDR' ELSE l.ShortDesc End, 
		                        JnlDesc=Case When g.Type='50' Then 'Batch Debit' ELSE l.FullDesc End,
		                        g.JnlNumber,
		                        g.LineNumber,
		                        g.GLAcc,
		                        gl.FullTitle as GLAccName,
		                        g.PostDate,
		                        g.ValueDate,
		                        g.TrnAmt/@CcyDive as TrnAmt,
		                        g.FullDesc,
		                        g.CCyType,
		                        h.ApprovedByTlr as TlrPostedBy, 
		                        t1.TlrName as TlrPostedName, 
		                        t1.Designation as TlrPostedPosit,
		                        h.UpdatedByTlr as TlrAuthBy, 
		                        t2.TlrName as TlrAuthName, 
		                        t2.Designation as TlrAuthPosit,
		                        Case
			                        When g.JnlNumber in ('EOD', 'EOM') Then g.Ref
			                        Else ref.Ref
		                        End as Ref
		                        into #GLTRNHIST
		                        From	
		                        GLTrnHist g 
		                        Left join [Lookup] l on g.Type=l.LookupCode and l.lookupid='TT' and l.LangType='001' left Join--'TT' --> TranactionType, 001->English
		                        GLAC gl on g.GLAcc=gl.GLAcc LEFT JOIN
		                        GLJHHIST h on g.PostDate=h.JnlDate and g.JnlNumber=h.JnlNumber LEFT JOIN
		                        SAF t1 ON h.ApprovedByTlr=t1.Tlr LEFT JOIN
		                        SAF t2 ON h.UpdatedByTlr=t2.Tlr LEFT JOIN
		                        #BATCH_REFF ref ON g.PostDate=ref.JnlDate and g.JnlNumber=ref.JnlNumber		
		                        Where g.PostDate Between ''+ @DateFrom +'' and ''+ @DateTo +''--- and g.JnlNumber not in ('EOD','EOM','EOY') 

		                        /*
		                        =======================================================================================================
		                        Get Result....
		                        =======================================================================================================
		                        */

	                        Select
								@BrCode as BrCode,
		                        @BrShort as BrName,
		                        TrnJnlType			,
		                        CASE 
			                         WHEN HISJnlType='001' THEN 'General Batch'
			                         WHEN HISJnlType='003' THEN 'Debit List'
			                         WHEN HISJnlType='004' THEN 'Credit List'
			                         WHEN HISJnlType='005' THEN 'GL Batch'
			                         WHEN HISJnlType='006' THEN 'Inter-branch' 
			                         ELSE 'Other'
		                        End as HISJnlType,
		                        CASE 
			                         WHEN JnlDesc='GL Transfer Debit' THEN 'DR'--DR=Debit
			                         WHEN JnlDesc='GL Transfer Credit' THEN 'CR'--CR=Credit
			                         WHEN JnlDesc='GL Debit' THEN 'DR'
			                         WHEN JnlDesc='GL Credit' THEN 'CR'
			                         WHEN JnlDesc='Batch Credit' THEN 'CR' 
                        --			 ELSE 'Other'
		                        End as HISJnlType,
		                        JnlTrn			, 
		                        JnlDesc			,
		                        JnlNumber		,
		                        LineNumber		,
		                        GLAcc			,
		                        GLAccName		,
		                        PostDate		,
		                        ValueDate		,
		                        TrnAmt			,
		                        FullDesc		,
		                        CCyType			,
		                        TlrPostedBy		,
		                        TlrPostedName	,
		                        TlrPostedPosit	,
		                        TlrAuthBy		,
		                        TlrAuthName		,
		                        TlrAuthPosit	,
		                        Ref				
	                          From #GLTRNHIST
	                        Order By  PostDate, JnlNumber, LineNumber

	                        truncate table #GLTRNHIST
	                        truncate table #BATCH_REFF
	                        drop table #GLTRNHIST
	                        drop table #BATCH_REFF";


                    ViewBag.BrCode = re.BrCode;
                    ViewBag.datestart = re.datestart;
                    ViewBag.dateend = re.dateend;
                    ViewBag.Acc = re.accountnumber;

                    DataTable dtreports = new DataTable();
                    dtreports = mutility.dtOnlyOneBranch(Sql, re.BrCode);
                    ViewBag.CurrRunDate = CurrRunDate.CurrRunDate(userkey);
                    ViewBag.dt_Report = dtreports;
                    downloadLoanStatement = dtreports;
                    BrCode = re.BrCode;
                    if (re.download == "download")
                    {
                        string ReportName = "GL Transaction by batch Reports";
                        string BrName = Convert.ToString(mutility.dbSingleResult("select BrLetter from BRANCH_LISTS where flag=1 and BrCode='" + re.BrCode + "'"));
                        ExportDataToExcel(downloadLoanStatement, BrName + "_" + ReportName);
                    }


                    return View("GLTransactionbybatch");
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
        [HttpPost]
        //Exec GLTransactionbybatch
        public ActionResult execCollateralInTool(Reports re)
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

                    DateTime datestart = re.datestart;
                    DateTime dateend = re.dateend;
                    string All = re.all;
                    string Sql = "";
                    if (All == "on")
                    {                   
                    Sql = @"	
                          set dateformat DMY            
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
                        where DBName = (select left(@DbName,len(@DbName)-4))

                        select 
		                        @BrShort as BrName,
		                        @BrCode as BrCode,
		                        co.Accountnumber as AccNumber,
		                        co.CollateralName as CollName,
		                        co.CoWnership as CollClientName,
		                        co.CCopyright as Copyright,
		                        co.CDouctype as SupportDocument,
		                        co.CDescription as Description,
		                        co.COR as CollOReceipt,
		                        co.CDate as PostDate,
		                        co.CProvidedName as ProvidedName,
		                        co.CRecipientsName as ReceiptName,
		                        co.Createdate as CreateDate,
		                        l.OpenDate,
		                        l.PrType,
		                        l.GrantedAmt
		
		
		                        from dbo.SKP_Collateral co 
		                        inner join lnacc l on l.acc+l.chd=co.Accountnumber 
	
		                        where l.accstatus between '11' and '98' and co.Createdate between '" + datestart+"' and '"+dateend+"'";
                    }
                    else
                    {
                        Sql = @"	
                         set dateformat DMY 
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
                        where DBName = (select left(@DbName,len(@DbName)-4))

                        select 
		                        @BrShort as BrName,
		                        @BrCode as BrCode,
		                        co.Accountnumber as AccNumber,
		                        co.CollateralName as CollName,
		                        co.CoWnership as CollClientName,
		                        co.CCopyright as Copyright,
		                        co.CDouctype as SupportDocument,
		                        co.CDescription as Description,
		                        co.COR as CollOReceipt,
		                        co.CDate as PostDate,
		                        co.CProvidedName as ProvidedName,
		                        co.CRecipientsName as ReceiptName,
		                        co.Createdate as CreateDate,
		                        l.OpenDate,
		                        l.PrType,
		                        l.GrantedAmt
		
		
		                        from dbo.SKP_Collateral co 
		                        inner join lnacc l on l.acc+l.chd=co.Accountnumber 
	
		                        where l.accstatus between '11' and '98' and l.Acc+l.chd='"+re.accountnumber+"' and co.Createdate between '" + datestart + "' and '" + dateend + "'";
                                   
                    }

                    ViewBag.BrCode = re.BrCode;
                    ViewBag.datestart = re.datestart;
                    ViewBag.dateend = re.dateend;
                    ViewBag.Acc = re.accountnumber;

                    DataTable dtreports = new DataTable();
                    dtreports = mutility.dtOnlyOneBranch(Sql, re.BrCode);
                    ViewBag.CurrRunDate = CurrRunDate.CurrRunDate(userkey);
                    ViewBag.dt_Report = dtreports;
                    downloadLoanStatement = dtreports;
                    BrCode = re.BrCode;
                    if (re.download == "download")
                    {
                        string ReportName = "GL Collateral Reports";
                        string BrName = Convert.ToString(mutility.dbSingleResult("select BrLetter from BRANCH_LISTS where flag=1 and BrCode='" + re.BrCode + "'"));
                        ExportDataToExcel(downloadLoanStatement, BrName + "_" + ReportName);
                    }


                    return View("CollateralInTool");
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
        [HttpPost]
        public ActionResult execListoffWrittenOff(Reports re)
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

                    DateTime datestart = re.datestart;
                    DateTime dateend = re.dateend;

                    string Sql = @"/*
                1-Create by Tuy Ravy:
                2-Create for get List of loan Write Off Uptodate
                3-Create date:2017-12-31
                4-for this script for Branch Run Get Reports
                */
                set dateformat YMD


                declare @BrCode nvarchar(10)
                declare @BrShort nvarchar(10)
                declare @BrName nvarchar(50)
                declare @DbName nvarchar(50)
                declare @#DayAfterEoC nvarchar(50)
                declare @InstNo nvarchar(10)
                declare @Close_Date datetime
                set @Close_Date=''
                set @InstNo=0
                set @#DayAfterEoC=0
                set @DbName = (SELECT DB_NAME())

                select 
	                @BrCode = SubBranchCode,
	                @BrShort = SubBranchID,
	                @BrName = SubBranchNameLatin
                from SKP_Brlist 
                where DBName = (select left(@DbName,len(@DbName)-4))



                --Declare @DayAffEOC numeric(18,2)
                --Declare @LongTerms numeric(18,2)
                --Declare @FullDesc varchar(50)
                --Declare @FullTitle varchar(50)
                --set @DayAffEOC=0
                --set @LongTerms=0
                --set @FullDesc=0
                --set @FullTitle=0

	                select		
		                @BrCode as BrCode,
		                @BrShort as BrName,
		                wo.Acc,
		                wo.CoID,
		                wo.CoName,
		                wo.ClientName,		
		                wo.BalAmt-isnull([dbo].[fn_yGetPriCurrBalAmtWO](wo.BrCode,wo.Acc),0) as BalAmt,
		                wo.IntToDate-ISNULL([dbo].[fn_yGetIntCurrBalAmtWO](wo.BrCode,wo.Acc),0) as IntTodate,
		                case 
			                when isnull(([dbo].[fn_yGetTotalCurrBalAmtWO](wo.BrCode,wo.Acc)),0)<0 then 0
			                else isnull(([dbo].[fn_yGetTotalCurrBalAmtWO](wo.BrCode,wo.Acc)),0)
		                end as AmtToClose,		   
		                wo.DisbDate,
		                wo.DisbAmt as GrantedAmt,
		                wo.MatDate,
		                wo.NumDayAfterEOC as #DayAfterEoC,
		                wo.InstNum as InstNo,
		                wo.FreqType,
		                case 
			                when datediff(day,wo.DisbDate,wo.MatDate)<=365 then 0
			                when datediff(day,wo.DisbDate,wo.MatDate)>365 then 1
		                end as LongTerms,
		                wo.OduePriAmt,
		                wo.OdueIntAmt,
		                wo.ProvisionAmt,
		                wo.ProvisionPercentage,
		                wo.MasterField,
		                wo.ccyType,
		                wo.AgeOfLoan,
		                wo.PrType,
		                PriAccName ,
		                wo.locationCode,
		                wo.ClientAdress,
		                wo.Phone1,
		                wo.Phone2,
		                wo.ReportDate,
		                wo.BrCode,
		                @BrName as BrName,
		                @BrShort as BrShort,
		                ([dbo].[fn_yGetLastPiad_DateWO](wo.BrCode,wo.Acc)) as Last_Piad_Date,
		                ([dbo].[fn_yGetLastAmtPiad_WO](wo.BrCode,wo.Acc)) as Last_Amt_Paid,
		                @Close_Date as CloseDate,
		                case 
			                when wo.ReportDate='2016-11-28' then N'??????????????????????????????'
			                when wo.ReportDate='2015-10-30' then N'??????????????????????????????'
			                when wo.ReportDate='2015-01-01' then N'??????????????????????????????'
		                end as StepByWO,
		                wo.IntPastMaturity,
		                wo.IntFromInstallment
		
		                from [dbo].[VTSDxWriteOff] WO
		                where Wo.BrCode=@BrCode
		                ";

                    ViewBag.BrCode = re.BrCode;
                    ViewBag.datestart = re.datestart;
                    ViewBag.dateend = re.dateend;
                    ViewBag.Acc = re.accountnumber;

                    ViewBag.CurrRunDate = CurrRunDate.CurrRunDate(userkey);
                    ViewBag.dt_Report = mutility.dtOnlyOneBranch(Sql, re.BrCode);
                    downloadLoanStatement = mutility.dtOnlyOneBranch(Sql, re.BrCode);
                    BrCode = re.BrCode;
                    if (re.download == "download")
                    {
                        string ReportName = "Full Trail Balance Reports";
                        string BrName = Convert.ToString(mutility.dbSingleResult("select BrLetter from BRANCH_LISTS where flag=1 and BrCode='" + re.BrCode + "'"));
                        ExportDataToExcel(downloadLoanStatement, BrName + "_" + ReportName);
                    }


                    return View("ListoffWrittenOff");
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

        [HttpPost]
        //Exec GLTransactionbybatch
        public ActionResult execWrittenoffCollection(Reports re)
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

                    DateTime datestart = re.datestart;
                    DateTime dateend = re.dateend;
                    string All = re.all;
                    string Sql = "";
                    if (All == "on")
                    {
                        Sql = @"	
                          set dateformat DMY
                        select 
	                        s.SubBranchID,
	                        s.SubBranchCode,
	                        c.TrnNO,
	                        c.Acc,
	                        c.TotalCollectedAmt,
	                        c.TotalCollectedAmtOrigCcy,
	                        c.DeductedAtm,
	                        c.CollectPrinciple,
	                        c.CollectPenalty,
	                        c.CollectInterest,
	                        c.CollectDate,
	                        c.CollectRule,
	                        c.BalanceAmt,
	                        c.CollectedByCoID,
	                        c.PostedDate,
	                        c.Reff	
	                         from [dbo].[VTSDxWOCollection] c inner join [dbo].[SKP_Brlist] s on c.BrCode=s.SubBranchCode
	                         where CollectDate between '" + datestart + "' and '" + dateend + "'";
                    }
                    else
                    {
                        Sql = @"	
                         set dateformat DMY
                        select 
	                        s.SubBranchID,
	                        s.SubBranchCode,
	                        c.TrnNO,
	                        c.Acc,
	                        c.TotalCollectedAmt,
	                        c.TotalCollectedAmtOrigCcy,
	                        c.DeductedAtm,
	                        c.CollectPrinciple,
	                        c.CollectPenalty,
	                        c.CollectInterest,
	                        c.CollectDate,
	                        c.CollectRule,
	                        c.BalanceAmt,
	                        c.CollectedByCoID,
	                        c.PostedDate,
	                        c.Reff	
	                         from [dbo].[VTSDxWOCollection] c inner join [dbo].[SKP_Brlist] s on c.BrCode=s.SubBranchCode
	                         where c.Acc='"+re.accountnumber+"' and c.CollectDate between '" + datestart + "' and '" + dateend + "'";

                    }

                    ViewBag.BrCode = re.BrCode;
                    ViewBag.datestart = re.datestart;
                    ViewBag.dateend = re.dateend;
                    ViewBag.Acc = re.accountnumber;

                    DataTable dtreports = new DataTable();
                    dtreports = mutility.dtOnlyOneBranch(Sql, re.BrCode);
                    ViewBag.CurrRunDate = CurrRunDate.CurrRunDate(userkey);
                    ViewBag.dt_Report = dtreports;
                    downloadLoanStatement = dtreports;
                    BrCode = re.BrCode;
                    if (re.download == "download")
                    {
                        string ReportName = "GL Written off Collection Reports";
                        string BrName = Convert.ToString(mutility.dbSingleResult("select BrLetter from BRANCH_LISTS where flag=1 and BrCode='" + re.BrCode + "'"));
                        ExportDataToExcel(downloadLoanStatement, BrName + "_" + ReportName);
                    }


                    return View("WrittenoffCollection");
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
        
    
