#include "StdAfx.h"
#include ".\genorders.h"

CGenOrders::CGenOrders(void)
{
}

CGenOrders::~CGenOrders(void)
{
}

CString CGenOrders::GetSQL(void)
{
	//the sql that selects what campaigns need orders
	try
	{
		CString sql;
		sql = "SELECT TodaysCampaigns.campaignID \r\n";
		sql = sql + ", TodaysCampaignOrders.campaignID \r\n";
		sql = sql + ", t_Agents.Agent_Name \r\n";
		sql = sql + ", campaign.companyID \r\n";
		sql = sql + ", t_Customers.cust_CompanyName \r\n";
		sql = sql + ", campaign.contactName \r\n";
		sql = sql + ", campaign.makeGoodQuantityPerOrder \r\n";
		sql = sql + ", campaign.email \r\n";
		sql = sql + ", campaign.voiceNumber \r\n";
		sql = sql + ", campaign.faxNumber \r\n";
		sql = sql + ", campaign.checkNumber \r\n";
		sql = sql + ", campaign.costPerLead \r\n";
		sql = sql + ", campaign.rateFilter \r\n";
		sql = sql + ", campaign.areacodeFilter \r\n";
		sql = sql + ", campaign.zipcodeFilter \r\n";
		sql = sql + ", campaign.amountFilter \r\n";
		sql = sql + ", campaign.verified \r\n";
		sql = sql + ", campaign.onePass \r\n";
		sql = sql + ", campaign.lookingToRefinace \r\n";
		sql = sql + ", campaign.lookingToCashOut \r\n";
		sql = sql + ", campaign.lookingfor2ndMortgage \r\n";
		sql = sql + ", campaign.looking4DebtConsolidation \r\n";
		sql = sql + ", campaign.sendWord \r\n";
		sql = sql + ", campaign.sendExcel \r\n";
		sql = sql + ", campaign.sendEmailNote \r\n";
		sql = sql + ", campaign.printOrder \r\n";
		sql = sql + ", campaign.oneLeadPerPage \r\n";
		sql = sql + ", campaign.comment \r\n";
		sql = sql + ", campaign.states \r\n";
		sql = sql + ", campaign.sendFax \r\n";
		sql = sql + ", TodaysCampaigns.AvailCredits \r\n";
		sql = sql + ", returnTotals.SumOfreturned \r\n";
		sql = sql + ", [totalMakeGood]+[SumOfreturned] AS total \r\n";
		sql = sql + ", campaign.lookingToID \r\n";
		sql = sql + ", campaign.refiTypeInt \r\n";
		sql = sql + ", campaign.lang \r\n";
		sql = sql + ", campaign.creditscore \r\n";
		sql = sql + ", campaign.homeTypeInt \r\n";
		sql = sql + ", campaign.locationInt \r\n";
		sql = sql + ", campaign.priority \r\n";
		sql = sql + ", campaign.randOrderQuan \r\n";
		sql = sql + ", campaign.termDate \r\n";
		sql = sql + ", campaign.sendMini1003 \r\n";
		sql = sql + ", campaign.agentID \r\n";
		sql = sql + ", campaign.cityFilter \r\n";
		sql = sql + ", campaign.UseLimit \r\n";
		sql = sql + ", campaign.AgeLimit \r\n";
		sql = sql + ", campaign.LTV \r\n";
		sql = sql + ", campaign.LTVabove \r\n";
		sql = sql + ", campaign.cashOut \r\n";
		sql = sql + ", campaign.[2ndMortgage] \r\n";
		sql = sql + ", campaign.debtCon \r\n";
		sql = sql + ", campaign.subPrime \r\n";
		sql = sql + ", campaign.SFR \r\n";
		sql = sql + ", campaign.Commercial \r\n";
		sql = sql + ", campaign.Mobile \r\n";
		sql = sql + "FROM  \r\n";
		sql = sql + "(t_Agents  \r\n";
		sql = sql + "INNER JOIN  \r\n";
		sql = sql + "((campaign  \r\n";
		sql = sql + "INNER JOIN  \r\n";
		sql = sql + "(  \r\n";
		sql = sql + "( \r\n";
		sql = sql + "SELECT TodaysCampaignsShowCredits.campaignID \r\n";
		sql = sql + ", TodaysCampaignsShowCredits.AvailCredits \r\n";
		sql = sql + "FROM  \r\n";
		sql = sql + "( \r\n";
		sql = sql + "SELECT campaign.campaignID \r\n";
		sql = sql + ", IIf(IsNull([SumOfreturned]),0,[SumOfreturned]) AS returned \r\n";
		sql = sql + ", Weekday(Now()) AS DayoftheWeek \r\n";
		sql = sql + ", campaign.monday \r\n";
		sql = sql + ", campaign.tuesday \r\n";
		sql = sql + ", campaign.wednesday \r\n";
		sql = sql + ", campaign.thursday \r\n";
		sql = sql + ", campaign.friday \r\n";
		sql = sql + ", campaign.saturday \r\n";
		sql = sql + ", campaign.sunday \r\n";
		sql = sql + ", campaign.pause \r\n";
		sql = sql + ", campaign.startDate \r\n";
		sql = sql + ", Now() AS Today \r\n";
		sql = sql + ", IIf(IsNull([SumOfOrd_Leads_Sent]),0,[SumOfOrd_Leads_Sent]) AS SentSoFar \r\n";
		sql = sql + ", campaign.totalMakeGood \r\n";
		sql = sql + ", [totalMakeGood]-[SentSoFar]+ IIf(IsNull([SumOfreturned]),0,[SumOfreturned]) AS AvailCredits \r\n";
		sql = sql + "FROM \r\n";
		sql = sql + "(campaign \r\n";
		sql = sql + "LEFT JOIN  \r\n";
		sql = sql + "( \r\n";
		sql = sql + "SELECT t_Orders.campaignID \r\n";
		sql = sql + ", Sum(t_Orders.Ord_Leads_Sent \r\n";
		sql = sql + ") AS SumOfOrd_Leads_Sent \r\n";
		sql = sql + "FROM t_Orders \r\n";
		sql = sql + "GROUP BY t_Orders.campaignID \r\n";
		sql = sql + ") AS SumLeadsSent \r\n";
		sql = sql + "ON campaign.campaignID = SumLeadsSent.campaignID) \r\n";
		sql = sql + "LEFT JOIN  \r\n";
		sql = sql + "( \r\n";
		sql = sql + "SELECT return.campaignID \r\n";
		sql = sql + ", Sum(return.returned) AS SumOfreturned \r\n";
		sql = sql + "FROM return \r\n";
		sql = sql + "GROUP BY return.campaignID \r\n";
		sql = sql + ") AS returnTotals  \r\n";
		sql = sql + "ON campaign.campaignID = returnTotals.campaignID \r\n";
		sql = sql + "WHERE (((Weekday(Now()))=2) \r\n";
		sql = sql + "AND ((campaign.monday)=True) \r\n";
		sql = sql + "AND ((campaign.pause)=False)  \r\n";
		sql = sql + "AND ((campaign.startDate)<=Now()))  \r\n";
		sql = sql + "OR (((Weekday(Now()))=3)  \r\n";
		sql = sql + "AND ((campaign.tuesday)=True)  \r\n";
		sql = sql + "AND ((campaign.pause)=False)  \r\n";
		sql = sql + "AND ((campaign.startDate)<=Now()))  \r\n";
		sql = sql + "OR (((Weekday(Now()))=4)  \r\n";
		sql = sql + "AND ((campaign.wednesday)=True)  \r\n";
		sql = sql + "AND ((campaign.pause)=False)  \r\n";
		sql = sql + "AND ((campaign.startDate)<=Now()))  \r\n";
		sql = sql + "OR (((Weekday(Now()))=5)  \r\n";
		sql = sql + "AND ((campaign.thursday)=True)  \r\n";
		sql = sql + "AND ((campaign.pause)=False)  \r\n";
		sql = sql + "AND ((campaign.startDate)<=Now()))  \r\n";
		sql = sql + "OR (((Weekday(Now()))=6)  \r\n";
		sql = sql + "AND ((campaign.friday)=True)  \r\n";
		sql = sql + "AND ((campaign.pause)=False)  \r\n";
		sql = sql + "AND ((campaign.startDate)<=Now()))  \r\n";
		sql = sql + "OR (((Weekday(Now()))=7)  \r\n";
		sql = sql + "AND ((campaign.saturday)=True)  \r\n";
		sql = sql + "AND ((campaign.pause)=False)  \r\n";
		sql = sql + "AND ((campaign.startDate)<=Now()))  \r\n";
		sql = sql + "OR (((Weekday(Now()))=1)  \r\n";
		sql = sql + "AND ((campaign.sunday)=True)  \r\n";
		sql = sql + "AND ((campaign.pause)=False)  \r\n";
		sql = sql + "AND ((campaign.startDate)<=Now())) \r\n";
		sql = sql + ") AS TodaysCampaignsShowCredits \r\n";
		sql = sql + "WHERE (((TodaysCampaignsShowCredits.AvailCredits)>0)) \r\n";
		sql = sql + ")AS TodaysCampaigns  \r\n";
		sql = sql + "LEFT JOIN  \r\n";
		sql = sql + "( \r\n";
		sql = sql + "SELECT t_Orders.campaignID \r\n";
		sql = sql + ", t_Orders.Ord_Date, Year([Ord_Date]) AS ThisYear \r\n";
		sql = sql + ", Month([Ord_Date]) AS ThisMonth \r\n";
		sql = sql + ", Day([Ord_Date]) AS ThisDay \r\n";
		sql = sql + "FROM t_Orders \r\n";
		sql = sql + "WHERE (((t_Orders.campaignID)>0)  \r\n";
		sql = sql + "AND ((Year([Ord_Date]))=Year(Now()))  \r\n";
		sql = sql + "AND ((Month([Ord_Date]))=Month(Now()))  \r\n";
		sql = sql + "AND ((Day([Ord_Date]))=Day(Now()))) \r\n";
		sql = sql + ") AS TodaysCampaignOrders  \r\n";
		sql = sql + "ON TodaysCampaigns.campaignID = TodaysCampaignOrders.campaignID)  \r\n";
		sql = sql + "ON campaign.campaignID = TodaysCampaigns.campaignID)  \r\n";
		sql = sql + "INNER JOIN  \r\n";
		sql = sql + "t_Customers  \r\n";
		sql = sql + "ON campaign.companyID = t_Customers.cust_Id)  \r\n";
		sql = sql + "ON t_Agents.Agent_Num = campaign.agentID)  \r\n";
		sql = sql + "LEFT JOIN  \r\n";
		sql = sql + "( \r\n";
		sql = sql + "SELECT return.campaignID \r\n";
		sql = sql + ", Sum(return.returned) AS SumOfreturned \r\n";
		sql = sql + "FROM return \r\n";
		sql = sql + "GROUP BY return.campaignID \r\n";
		sql = sql + ") AS returnTotals  \r\n";
		sql = sql + "ON campaign.campaignID = returnTotals.campaignID \r\n";
		sql = sql + "WHERE (((TodaysCampaignOrders.campaignID) Is Null)) \r\n";

		return sql;
	

	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: CGenOrders::GetSQL";
		PrintOut(err);
		CWriteLog::WriteFail(err);
		return CString();
	}
}

bool CGenOrders::ReadCampaigns(void)
{
	//read campaigns that need orders
	bool success = false;	
    try
    {
		HRESULT hr = S_OK;
		_RecordsetPtr pRst;
		//PrintOut("Init database connection", 0);
         CoInitialize(NULL);
          // Define string variables.
		 
		 _bstr_t strCnn("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=z:\\Mortgage.mdb;Jet OLEDB:Database Password=cepher;");
		      

      // Call Create instance to instantiate the Record set
      hr = pRst.CreateInstance(__uuidof(Recordset));

      if(FAILED(hr))
      {            
		  PrintOut("Failed creating record set instance");		 
      }
	  else
	  {
		if(FAILED(hr))
		{            
			PrintOut("Failed creating record set instance");		 
		}
		else
		{
			srand( (unsigned)time( NULL ) );

			

			_RecordsetPtr pOrdersRst; 
			hr = pOrdersRst.CreateInstance(__uuidof(Recordset));
			
			CString sql, ordersSql;
			ordersSql = "SELECT * FROM t_Orders";

			sql = GetSQL();		  
			   
			//SendSQLString(IDC_EDIT_SQL, sql);
			PrintOut("Submit SQL: " + sql);
			PrintOut("Reading Campaigns");
			WriteLog("Reading Campaigns");
			Progress(1, 100);
			pRst->Open(sql.GetBuffer(), strCnn, adOpenStatic, adLockReadOnly, adCmdText);		
			pOrdersRst->Open(ordersSql.GetBuffer(), strCnn, adOpenForwardOnly, adLockOptimistic, adCmdText);

			if(!pRst->EndOfFile) pRst->MoveFirst(); 
			else 
			{
				PrintOut("No Campaigns Need Orders");
				Progress(0, 100);
				return false;
			}
			int max = pRst->GetRecordCount();
			int x = 0;
			CString sOrderCount = "Reading Campaigns:";
			sOrderCount = sOrderCount + GetString((long)max);
			PrintOut(sOrderCount);
			Progress(x, max);
			while(!pRst->EndOfFile)
			{	
				CString str;

				COleDateTime aDate;
				aDate = COleDateTime::GetCurrentTime();
				CString sDate = aDate.Format(_T("%m/%d/%y %I:%M:%S %p"));
				
				str = "Ord_Date (" + sDate +") ";
				str = str + "campaignID (" + GetString(GetFieldLong(pRst, "TodaysCampaigns.campaignID")) + ") ";
				str = str + "Ord_States (" + GetFieldString(pRst, "states") + ") ";
				str = str + "Ord_CompanyID (" + GetString(GetFieldLong(pRst, "companyID")) + ") ";
				str = str + "Ord_CompanyName (" + GetFieldString(pRst, "cust_CompanyName") + ") ";
				str = str + "Ord_Contact (" + GetFieldString(pRst, "contactName") + ") ";
				str = str + "Ord_Telephone (" + GetFieldString(pRst, "voiceNumber") + ") ";
				str = str + "Ord_Fax (" + GetFieldString(pRst, "faxNumber") + ") ";
				str = str + "randOrderQuan (" + GetBool_String(GetFieldBool(pRst, "randOrderQuan")) + ") ";
				str = str + "makeGoodQuantityPerOrder (" + GetFieldString(pRst, "makeGoodQuantityPerOrder") + ") ";
				str = str + "AvailCredits (" + GetString(GetFieldDouble(pRst, "AvailCredits")) + ") ";
				str = str + "Ord_Email (" + GetFieldString(pRst, "email") + ") ";
				str = str + "Ord_AgentName (" + GetFieldString(pRst, "Agent_Name") + ") ";
				str = str + "agentID (" + GetString(GetFieldLong(pRst, "agentID")) + ") ";
				str = str + "Ord_Rate_Filter (" + GetFieldString(pRst, "rateFilter") + ") ";
				str = str + "Ord_ZipCode_Filter (" + GetFieldString(pRst, "zipcodeFilter") + ") ";
				str = str + "Ord_DesiredLoanAmt_Filter (" + GetFieldString(pRst, "amountFilter") + ") ";
				str = str + "Ord_AreaCode_Filter (" + GetFieldString(pRst, "areacodeFilter") + ") ";
				str = str + "Ord_LookingTo_Refinance (" + GetBool_String(GetFieldBool(pRst, "lookingToRefinace")) + ") ";
				str = str + "Ord_LookingTo_CashOut (" + GetBool_String(GetFieldBool(pRst, "lookingToCashOut")) + ") ";
				str = str + "Ord_LookingTo_2ndMortgage ("  + GetBool_String(GetFieldBool(pRst, "lookingfor2ndMortgage")) + ") ";
				str = str + "Ord_LookingTo_DebtCons (" + GetBool_String(GetFieldBool(pRst, "looking4DebtConsolidation")) + ") ";
				str = str + "Ord_OnePass (" + GetBool_String(GetFieldBool(pRst, "onePass")) + ") ";
				str = str + "Ord_Excel (" + GetBool_String(GetFieldBool(pRst, "sendExcel")) + ") ";
				str = str + "Ord_Send_Fax (" + GetBool_String(GetFieldBool(pRst, "sendFax")) + ") ";
				str = str + "Ord_Verified (" + GetBool_String(GetFieldBool(pRst, "verified")) + ") ";
				str = str + "Ord_Word (" + GetBool_String(GetFieldBool(pRst, "sendWord")) + ") ";
				str = str + "Ord_1Lead (" + GetBool_String(GetFieldBool(pRst, "oneLeadPerPage")) + ") ";
				str = str + "Ord_EmailNote (" + GetBool_String(GetFieldBool(pRst, "sendEmailNote")) + ") ";
				str = str + "Ord_Print (" + GetBool_String(GetFieldBool(pRst, "printOrder")) + ") ";
				
				str = str + "lookingToID (" + GetString(GetFieldLong(pRst, "lookingToID")) + ") ";
				str = str + "refiTypeInt (" + GetString(GetFieldLong(pRst, "refiTypeInt")) + ") ";
				str = str + "lang (" + GetString(GetFieldLong(pRst, "lang")) + ") ";
				str = str + "creditscore (" + GetString(GetFieldLong(pRst, "creditscore")) + ") ";
				str = str + "homeTypeInt (" + GetString(GetFieldLong(pRst, "homeTypeInt")) + ") ";
				str = str + "locationInt (" + GetString(GetFieldLong(pRst, "locationInt")) + ") ";
				str = str + "priority (" + GetBool_String(GetFieldBool(pRst, "priority")) + ") ";
				str = str + "Ord_sendMini1003 (" + GetBool_String(GetFieldBool(pRst, "sendMini1003")) + ") ";
				str = str + "adCampaign (" + GetBool_String(GetFieldBool(pRst, "randOrderQuan")) + ") ";

				str = str + "City (" + GetFieldString(pRst, "cityFilter") + ") ";				
				str = str + "LTV (" + GetString(GetFieldDouble(pRst, "LTV")) + ") ";
				str = str + "LTVAbove (" + GetBool_String(GetFieldBool(pRst, "LTVabove")) + ") ";
				str = str + "UseLimit (" + GetString(GetFieldLong(pRst, "UseLimit")) + ") ";
				str = str + "AgeLimit (" + GetString(GetFieldLong(pRst, "AgeLimit")) + ") ";
				str = str + "cashOut (" + GetString(GetFieldLong(pRst, "cashOut")) + ") ";
				str = str + "2ndMortgage (" + GetString(GetFieldLong(pRst, "2ndMortgage")) + ") ";
				str = str + "debtCon (" + GetString(GetFieldLong(pRst, "debtCon")) + ") ";
				str = str + "subPrime (" + GetBool_String(GetFieldBool(pRst, "subPrime")) + ") ";
				str = str + "SFR (" + GetString(GetFieldLong(pRst, "SFR")) + ") ";
				str = str + "Commercial (" + GetString(GetFieldLong(pRst, "Commercial")) + ") ";
				str = str + "Mobile (" + GetString(GetFieldLong(pRst, "Mobile")) + ") ";

				PrintOut(str);
				////////////////////////////////////////////////////////////////////////////////
				////////////////////////////////insert order////////////////////////////////////
				pOrdersRst->AddNew();
				COleDateTime today = COleDateTime::GetCurrentTime();	
				pOrdersRst->Fields->Item["Ord_Date"]->Value = _variant_t(today);
				pOrdersRst->Fields->Item["Ord_State"]->Value = pRst->Fields->GetItem("states")->Value;
				pOrdersRst->Fields->Item["Ord_CompanyID"]->Value = pRst->Fields->GetItem("companyID")->Value;
				pOrdersRst->Fields->Item["Ord_CompanyName"]->Value = pRst->Fields->GetItem("cust_CompanyName")->Value;
				pOrdersRst->Fields->Item["Ord_Contact"]->Value = pRst->Fields->GetItem("contactName")->Value; 
				pOrdersRst->Fields->Item["Ord_Telephone"]->Value = pRst->Fields->GetItem("voiceNumber")->Value; 
				pOrdersRst->Fields->Item["Ord_Fax"]->Value = pRst->Fields->GetItem("faxNumber")->Value; 
				double fQuanitity = 0;
				CString sQuanitity = GetFieldString(pRst, "makeGoodQuantityPerOrder");
				fQuanitity = atof(sQuanitity.GetBuffer());
				if(GetFieldBool(pRst, "randOrderQuan"))
				{
					//random generated order quantity					
					double fLow = 0.75 * fQuanitity;
					double fHigh = 1.25 * fQuanitity;
					long iRange = (long)(fHigh - fLow);
					if(iRange < 1) iRange = 1;
					long aRand = ((rand()%iRange)+1);
					if(aRand < 1) aRand = 1;
					long aRandQuantity = ((long)fLow) + aRand;
					if(aRandQuantity < 1) aRandQuantity = 1;					

					pOrdersRst->Fields->Item["Ord_MakeGood"]->Value = aRandQuantity;						 
				}
				else
				{					
					double fAvailCredits = 0;
					fAvailCredits = GetFieldDouble(pRst, "AvailCredits");
					if((fAvailCredits - fQuanitity) < 0)
					{
						pOrdersRst->Fields->Item["Ord_MakeGood"]->Value = fAvailCredits;
					}
					else
					{
						pOrdersRst->Fields->Item["Ord_MakeGood"]->Value = fQuanitity;
					}
				}
				pOrdersRst->Fields->Item["Ord_Email"]->Value = pRst->Fields->GetItem("email")->Value;
				pOrdersRst->Fields->Item["Ord_AgentName"]->Value = pRst->Fields->GetItem("Agent_Name")->Value;
				pOrdersRst->Fields->Item["agentID"]->Value = pRst->Fields->GetItem("agentID")->Value;
				pOrdersRst->Fields->Item["Ord_Rate_Filter"]->Value = pRst->Fields->GetItem("rateFilter")->Value;
				pOrdersRst->Fields->Item["Ord_ZipCode_Filter"]->Value = pRst->Fields->GetItem("zipcodeFilter")->Value;
				pOrdersRst->Fields->Item["Ord_DesiredLoanAmt_Filter"]->Value = pRst->Fields->GetItem("amountFilter")->Value;
				pOrdersRst->Fields->Item["Ord_AreaCode_Filter"]->Value = pRst->Fields->GetItem("areacodeFilter")->Value;
				pOrdersRst->Fields->Item["Ord_LookingTo_Refinance"]->Value = pRst->Fields->GetItem("lookingToRefinace")->Value;
				pOrdersRst->Fields->Item["Ord_LookingTo_CashOut"]->Value = pRst->Fields->GetItem("lookingToCashOut")->Value;
				pOrdersRst->Fields->Item["Ord_LookingTo_2ndMortgage"]->Value = pRst->Fields->GetItem("lookingfor2ndMortgage")->Value;
				pOrdersRst->Fields->Item["Ord_LookingTo_DebtCons"]->Value = pRst->Fields->GetItem("looking4DebtConsolidation")->Value;
				pOrdersRst->Fields->Item["Ord_OnePass"]->Value = pRst->Fields->GetItem("onePass")->Value;
				pOrdersRst->Fields->Item["Ord_Excel"]->Value = pRst->Fields->GetItem("sendExcel")->Value;
				pOrdersRst->Fields->Item["Ord_Send_Fax"]->Value = pRst->Fields->GetItem("sendFax")->Value;
				pOrdersRst->Fields->Item["Ord_Verified"]->Value = pRst->Fields->GetItem("verified")->Value;
				pOrdersRst->Fields->Item["Ord_Word"]->Value = pRst->Fields->GetItem("sendWord")->Value;
				pOrdersRst->Fields->Item["Ord_1Lead"]->Value = pRst->Fields->GetItem("oneLeadPerPage")->Value;
				pOrdersRst->Fields->Item["Ord_EmailNote"]->Value = pRst->Fields->GetItem("sendEmailNote")->Value;
				pOrdersRst->Fields->Item["Ord_Print"]->Value = pRst->Fields->GetItem("printOrder")->Value;
				pOrdersRst->Fields->Item["lookingToID"]->Value = pRst->Fields->GetItem("lookingToID")->Value;
				if(GetFieldLong(pRst, "refiTypeInt") == -1)
				{
					pOrdersRst->Fields->Item["refiTypeInt"]->Value = 0;
				}
				else
				{
					pOrdersRst->Fields->Item["refiTypeInt"]->Value = pRst->Fields->GetItem("refiTypeInt")->Value;
				}

				pOrdersRst->Fields->Item["lang"]->Value = pRst->Fields->GetItem("lang")->Value;
				pOrdersRst->Fields->Item["creditscore"]->Value = pRst->Fields->GetItem("creditscore")->Value;
				pOrdersRst->Fields->Item["homeTypeInt"]->Value = pRst->Fields->GetItem("homeTypeInt")->Value;
				pOrdersRst->Fields->Item["locationInt"]->Value = pRst->Fields->GetItem("locationInt")->Value;
				pOrdersRst->Fields->Item["priority"]->Value = pRst->Fields->GetItem("priority")->Value;
				pOrdersRst->Fields->Item["Ord_sendMini1003"]->Value = pRst->Fields->GetItem("sendMini1003")->Value;
				pOrdersRst->Fields->Item["adCampaign"]->Value = pRst->Fields->GetItem("randOrderQuan")->Value;
				pOrdersRst->Fields->Item["City"]->Value = pRst->Fields->GetItem("cityFilter")->Value;
				if(GetFieldDouble(pRst, "LTV") == -1)
				{
					pOrdersRst->Fields->Item["LTV"]->Value = 0;
				}
				else
				{
					pOrdersRst->Fields->Item["LTV"]->Value = pRst->Fields->GetItem("LTV")->Value;
				}
				pOrdersRst->Fields->Item["LTVAbove"]->Value = pRst->Fields->GetItem("LTVabove")->Value;
				pOrdersRst->Fields->Item["campaignID"]->Value = pRst->Fields->GetItem("TodaysCampaigns.campaignID")->Value;
				if(GetFieldLong(pRst, "UseLimit") == -1)
				{
					pOrdersRst->Fields->Item["UseLimit"]->Value = 0;
				}
				else
				{
					pOrdersRst->Fields->Item["UseLimit"]->Value = pRst->Fields->GetItem("UseLimit")->Value;
				}
				if(GetFieldLong(pRst, "AgeLimit") == -1)
				{
					pOrdersRst->Fields->Item["AgeLimit"]->Value = 0;
				}
				else
				{
					pOrdersRst->Fields->Item["AgeLimit"]->Value = pRst->Fields->GetItem("AgeLimit")->Value;
				}
				
				pOrdersRst->Fields->Item["cashOut"]->Value = pRst->Fields->GetItem("cashOut")->Value;
				pOrdersRst->Fields->Item["2ndMortgage"]->Value = pRst->Fields->GetItem("2ndMortgage")->Value;
				pOrdersRst->Fields->Item["debtCon"]->Value = pRst->Fields->GetItem("debtCon")->Value;
				pOrdersRst->Fields->Item["subPrime"]->Value = pRst->Fields->GetItem("subPrime")->Value;
				if(GetFieldLong(pRst, "SFR") == -1)
				{
					pOrdersRst->Fields->Item["SFR"]->Value = 0;
				}
				else
				{
					pOrdersRst->Fields->Item["SFR"]->Value = pRst->Fields->GetItem("SFR")->Value;
				}
				if(GetFieldLong(pRst, "Commercial") == -1)
				{
					pOrdersRst->Fields->Item["Commercial"]->Value = 0;
				}
				else
				{
					pOrdersRst->Fields->Item["Commercial"]->Value = pRst->Fields->GetItem("Commercial")->Value;
				}
				if(GetFieldLong(pRst, "Mobile") == -1)
				{
					pOrdersRst->Fields->Item["Mobile"]->Value = 0;
				}
				else
				{
					pOrdersRst->Fields->Item["Mobile"]->Value = pRst->Fields->GetItem("Mobile")->Value;
				}				
				
				pOrdersRst->Update();

				////////////////////////////////////////////////////////////////////////////////

				x++;
				Progress(x, max);
				pRst->MoveNext();
				
			}		
			PrintOut(GetString((long)x));
			WriteLog(GetString((long)x));
			success = true;
			pRst->Close();
			pOrdersRst->Close();
			
		}
	  }	
	  CoUninitialize();
	}
	catch(_com_error & ce)
	{
		CString str;
		
		str = "Error Read Campaign: " + GetCString(ce.Description());
		PrintOut(str);
		CWriteLog::WriteFail(str);
		success = false;		
		
	}	
  
	return success;
}

void CGenOrders::SetParentWindow(HWND m_hwnd)
{	
	//set the window that will recieve messages
	try
	{
		mainWin = m_hwnd;
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: CGenOrders::SetParentWindow(HWND m_hwnd)";
		PrintOut(err);
		CWriteLog::WriteFail(err);
	}
}

void CGenOrders::PrintOut(CString str)
{
	//print out text to the feedback listbox
	try
	{
		if(mainWin)
		{
			SendMessage(mainWin, S_START_MESSAGE, 0, 0);
			int iLen = str.GetLength();
			for(int x = 0; x < iLen; x++)
			{
				int iLetter = (int) str.GetAt(x);
				SendMessage(mainWin, S_CHAR_MESSAGE, iLetter, 0);
			}

			SendMessage(mainWin, S_END_MESSAGE, 0, 0);
		}
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: CGenOrders::PrintOut(CString str)";
		
		CWriteLog::WriteFail(err);
	}
}

void CGenOrders::Progress(int iPos, int iMax)
{
	//change the progress bar position
	try
	{
		if(mainWin)
		{
			SendMessage(mainWin, ID_PROGRESS, iPos, iMax);
		}
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: CGenOrders::Progress(int iPos, int iMax)";
		PrintOut(err);
		CWriteLog::WriteFail(err);
	}
}

CString CGenOrders::GetString(long iNum)
{
	try
	{
		return CStringHelper::GetCString(iNum);
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: CGenOrders::GetString(long iNum)";
		PrintOut(err);
		CWriteLog::WriteFail(err);
	}
}

CString CGenOrders::GetBool_String(bool state)
{	
	try
	{
		return CStringHelper::GetCString(state);
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: CGenOrders::GetBool_String(bool state)";
		PrintOut(err);
		CWriteLog::WriteFail(err);
	}
}

CString CGenOrders::GetString(double fNum)
{	
	try
	{
		return CStringHelper::GetCString(fNum);
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: CGenOrders::GetString(double fNum)";
		PrintOut(err);
		CWriteLog::WriteFail(err);
		return CString();
	}

	
}

void CGenOrders::WriteLog(CString str)
{	
	try
	{
		CWriteLog::Write(str);
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: CGenOrders::WriteLog(CString str)";
		PrintOut(err);
		CWriteLog::WriteFail(err);
	}


	///////////////////////////////////////////////////////////////////
	
	
	/*try
    {
		HRESULT hr;
		hr = S_OK;	
		 CoInitialize(NULL);

		_RecordsetPtr pRst;		
		_bstr_t strCnn;
		pRst = NULL;

		strCnn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=BatcherLog.mdb;Jet OLEDB:Database Password=cepher;";
		 // Call Create instance to instantiate the Record set
		CString sql = "SELECT * FROM Feedback";
		hr = pRst.CreateInstance(__uuidof(Recordset));
		if(FAILED(hr))
		{            
			PrintOut("Failed creating batcher log record set instance");
		}
		else
		{			
			pRst->Open(sql.GetBuffer(), strCnn, adOpenForwardOnly, adLockOptimistic, adCmdText);
			COleDateTime aTime;
			aTime = COleDateTime::GetCurrentTime();
			
			pRst->AddNew();			
			pRst->Fields->Item["Date"]->Value = _variant_t(aTime);	
			pRst->Fields->Item["Message"]->Value = str.GetBuffer();	
			pRst->Update();
			
		}
		CoUninitialize();
	}
	catch(_com_error & ce)
	{
		CString str;		
		str = "Error WriteLog: " + GetCString(ce.Description());
		PrintOut(str);	
		
	}*/
	///////////////////////////////////////////////////////////////////
}
