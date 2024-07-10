#include "StdAfx.h"
#include ".\orderdb.h"
#include "resource.h"

COrderDB::COrderDB(void)
: stop(false)
, bOnePass(false)
, bVerified(false)
, bPriority(false)
, bZeros(false)
, bNewCustomer(false)
, testing(false)
, sState(_T(""))
{
	mySettings.ReadSettings();
	mySettings.ReadLocations();	
}

COrderDB::~COrderDB(void)
{
}

void COrderDB::Add(COrder aOrder)
{
	
	try
	{
		myOrderMap.SetAt(aOrder.GetOrderIDLong(), aOrder);
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: COrderDB::Add(COrder aOrder)";
		PrintOut(err);
		CWriteLog::WriteFail(err);
		//WriteLog(err);
	}
}

bool COrderDB::GetOrder(long iID, COrder &aOrder)
{	
	try
	{
		return myOrderMap.Lookup(iID, aOrder);
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: COrderDB::GetOrder(long iID, COrder &aOrder)";
		PrintOut(err);
		CWriteLog::WriteFail(err);
	}
}

void COrderDB::RemoveAll(void)
{
	
	try
	{
		myOrderMap.RemoveAll();
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: COrderDB::RemoveAll";
		PrintOut(err);
		CWriteLog::WriteFail(err);
	}
}

void COrderDB::PrintOrders(void)
{
	try
	{
		POSITION pos = myOrderMap.GetStartPosition();
		long    nKey;
		COrder aOrder;
		aOrder.SetStop(&stop);
		PrintOut("Open Orders: ", 0);
		while ((pos != NULL)&& (!stop))
		{	   
			myOrderMap.GetNextAssoc( pos, nKey, aOrder );
			PrintOut(aOrder.ToString(), aOrder.GetLocationLong());
			
			//aOrder.PrintMatchingLeads();
			aOrder.PrintAssignedLeads();
		}
		PrintOut("Complete Orders: ", 0);
		pos = myCompleteOrderMap.GetStartPosition();
		while ((pos != NULL) && (!stop))
		{	   
			myCompleteOrderMap.GetNextAssoc( pos, nKey, aOrder );
			PrintOut(aOrder.ToString(), aOrder.GetLocationLong());
			
			//aOrder.PrintMatchingLeads();
			aOrder.PrintAssignedLeads();
		}
	

	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: COrderDB::PrintOrders";
		PrintOut(err);
		CWriteLog::WriteFail(err);
	}
}

void COrderDB::PrintOut(CString str, int div)
{
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
		err = "Error: COrderDB::PrintOut(CString str, int div)";
		PrintOut(err);
		CWriteLog::WriteFail(err);
	}
}

void COrderDB::SetParentWindow(HWND m_wnd)
{
	try
	{
		mainWin = m_wnd;
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: COrderDB::SetParentWindow(HWND m_wnd)";
		PrintOut(err);
		CWriteLog::WriteFail(err);
	}
}

void COrderDB::SetStop(bool state)
{
	//stop = state;
}

bool COrderDB::AssignLeads(void)
{
	bool success = false;
	
	
    try
    {
		_RecordsetPtr pRst;
		HRESULT hr = S_OK;
		//PrintOut("Init database connection", 0);
         CoInitialize(NULL);
          // Define string variables.
		 
		 _bstr_t strCnn("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=z:\\Mortgage.mdb;Jet OLEDB:Database Password=cepher;");
      

		// Call Create instance to instantiate the Record set
		hr = pRst.CreateInstance(__uuidof(Recordset));

		if(FAILED(hr))
		{            
			PrintOut("Failed creating record set instance", 0);		 
		}
		else
		{
			
			CString sql = "SELECT * FROM t_Processed_Orders";
			PrintOut("Submit SQL: " + sql, 0);
			Progress(1, 100);
			int iPos = 0;
			int iMax = GetCountAssigned();			
			pRst->Open(sql.GetBuffer(), strCnn, adOpenStatic, adLockOptimistic, adCmdText);				
			if(myCompleteOrderMap.GetCount()>0)
			{
				POSITION pos = myCompleteOrderMap.GetStartPosition();
				long    nKey;
				COrder aOrder;
				aOrder.SetStop(&stop);
				PrintOut("Insert Processed Orders", 0);
				while ((pos != NULL)&& (!stop))
				{
					//PrintOut("Inside while loop", 0);
					myCompleteOrderMap.GetNextAssoc( pos, nKey, aOrder );
					
					long orderid = aOrder.GetOrderIDLong();
					long compid = aOrder.GetCompanyIDLong();
					CIDArray leadIDs = aOrder.GetAssignedLeadIDs();
					for(long x = 0; x < leadIDs.GetCount(); x++)
					{						
						pRst->AddNew();
						pRst->Fields->Item["PO_Ord_ID"]->Value = orderid;
						pRst->Fields->Item["PO_Company_ID"]->Value = compid;
						pRst->Fields->Item["PO_Lead_ID"]->Value = leadIDs[x];
						COleDateTime today = COleDateTime::GetCurrentTime();			
						pRst->Fields->Item["PO_Lead_SentOn"]->Value = _variant_t(today);						
						iPos++;						
						Progress(iPos, iMax);
						pRst->Update();						
					}
					MarkOrderReadyToSend(aOrder);
					CString sCount, sID, sQuantity, msg;
					sID = aOrder.GetOrderID();
					sCount = GetString(aOrder.GetCountAssigned());	
					sQuantity = aOrder.GetQuantity();
					msg = "Order#:(" + sID + ")";
					msg = msg + " Quantity Ordered: (" + sQuantity + ")";
					msg = msg + " Quantity Assigned:(" + sCount + ")";
					
					PrintOut(msg, 0);
					WriteLog(msg);
					if(aOrder.GetCountAssigned()> aOrder.GetQuanityLong())
					{
						PrintFail(msg);
					}
					
				}
			}
		}		
		CoUninitialize();
	}
	catch(_com_error & ce)
	{
		CString str;
		
		str = "Error COrderDB::AssignLeads: " + GetCString(ce.Description());
		PrintOut(str, 0);
		CWriteLog::WriteFail(str);
		success = false;		
		
	}  
  return success;	
}

long COrderDB::GetMaxNeeded(void)
{	
	try
	{
		POSITION pos = myOrderMap.GetStartPosition();
		long max = 0;
		long    nKey;
		COrder aOrder;
		while (pos != NULL)
		{	   
			myOrderMap.GetNextAssoc( pos, nKey, aOrder );
			max = max + aOrder.GetQuanityLong();
			
		}
		return max;
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: COrderDB::GetMaxNeeded";
		PrintOut(err);
		CWriteLog::WriteFail(err);
		
	}
}

bool COrderDB::ReadLeads(void)
{	
	bool success = false;
	try
	{
		
		DeleteAssignedLeads();//back out and delete any partially processed leads
		WriteLog("Orders Reading Leads");
		long max = GetMaxNeeded();
		POSITION pos = myOrderMap.GetStartPosition();
		int count = myOrderMap.GetCount();
		int x = 0;
		ProgressAll(x, count);
		long    nKey;
		COrder aOrder;
		while ((pos != NULL)&& (!stop))
		{	   
			myOrderMap.GetNextAssoc( pos, nKey, aOrder );
			//LoadScreenFromOrder(&aOrder);
			aOrder.SetMaxNeeded(max);
			aOrder.SetStop(&stop);// stop pointer is used to halt long processes when processing is canceled
			//aOrder.DeleteAssignedLeads();//back out and delete any partially processed leads
			aOrder.ReadLeads();			
			myOrderMap.SetAt(aOrder.GetOrderIDLong(), aOrder);
			x++;
			ProgressAll(x, count);
		}
		success = true;
	}
	catch(...)
	{
		success = false;
		//error trap
		CString err;
		err = "Error: COrderDB::ReadLeads";
		PrintOut(err);
		CWriteLog::WriteFail(err);
	}
	return success;	
}

void COrderDB::LoadScreenFromOrder(COrder* aOrder)
{
	try
	{
		SendCEditString(IDC_EDIT_COMPANYID, aOrder->GetCompanyID());
		SendCEditString(IDC_EDIT_STATES, aOrder->GetStates());//
		SendCEditString(IDC_EDIT_CONTACT, aOrder->GetContact());
		SendCEditString(IDC_EDIT_ORDERID, aOrder->GetOrderID());
		SendCEditString(IDC_EDIT_FAXNUMBER, aOrder->GetFaxNumber());
		SendCEditString(IDC_EDIT_QUANTITY, aOrder->GetQuantity());
		SendCEditString(IDC_EDIT_EMAIL, aOrder->GetEmail());
		SendCEditString(IDC_EDIT_AGENT, aOrder->GetAgent());
		SendCEditString(IDC_EDIT_RATE, aOrder->GetRateFilter());
		SendCEditString(IDC_EDIT_ZIP_FILTER, aOrder->GetZipFilter());
		SendCEditString(IDC_EDIT_DESIRED_LOAN_FILTER, aOrder->GetDesiredLoanFilter());
		SendCEditString(IDC_EDIT_AREACODE_FILTER, aOrder->GetAreacodeFilter());
		SendCEditString(IDC_EDIT_CASHOUT, aOrder->GetCashout());
		SendCEditString(IDC_EDIT_2NDMORTGAGE, aOrder->Get2ndMortgage());
		SendCEditString(IDC_EDIT_DEBTCON, aOrder->GetDebtCon());
		SendCEditString(IDC_EDIT_AGENTID, aOrder->GetAgentID());
		SendCEditString(IDC_EDIT_LOOKINGTO, aOrder->GetLookingto());
		SendCEditString(IDC_EDIT_LANG, aOrder->GetLang());
		SendCEditString(IDC_EDIT_CREDIT, aOrder->GetCreditscore());
		SendCEditString(IDC_EDIT_HOMETYPE, aOrder->GetHomeType());
		SendCEditString(IDC_EDIT_LOCATION, aOrder->GetLocation());
		SendCEditString(IDC_EDIT_SFR, aOrder->GetSFR());
		SendCEditString(IDC_EDIT_MOBILE, aOrder->GetMobile());
		SendCEditString(IDC_EDIT_COMMERCIAL, aOrder->GetCommercial());
		SendCEditString(IDC_EDIT_CITY, aOrder->GetCity());
		SendCEditString(IDC_EDIT_LTV, aOrder->GetLTV());
		SendSQLString(IDC_EDIT_SQL, aOrder->GetSQL());//
		SendCheckBoxState(IDC_CHECK_VERIFIED, aOrder->GetVerified());
		SendCheckBoxState(IDC_CHECK_ONEPASS, aOrder->GetOnePass());
		SendCheckBoxState(IDC_CHECK_WORD, aOrder->GetSendWord());
		SendCheckBoxState(IDC_CHECK_EXCEL, aOrder->GetExcel());
		SendCheckBoxState(IDC_CHECK_FAX, aOrder->GetSendFax());
		SendCheckBoxState(IDC_CHECK_ONELEAD, aOrder->GetOneLeadPerPage());
		SendCheckBoxState(IDC_CHECK_EMAILNOTE, aOrder->GetSendEmailNote());
		SendCheckBoxState(IDC_CHECK_PRINT, aOrder->GetPrint());//
		SendCheckBoxState(IDC_CHECK_PRIORITY, aOrder->GetPriority());
		SendCheckBoxState(IDC_CHECK_MINI1003, aOrder->GetSendMini1003());
		SendCheckBoxState(IDC_CHECK_ADCAMPAIGN, aOrder->GetAdCampaign());
		SendCheckBoxState(IDC_CHECK_SUBPRIME, aOrder->GetSubPrime());
		SendCheckBoxState(IDC_CHECK_ABOVE, aOrder->GetLTVAbove());
		SendCEditString(IDC_EDIT_LAST_ORDER_QUANTITY, aOrder->GetLastOrderQuantity());
		SendCEditString(IDC_EDIT_COMPANY_ORDERS_COUNT, aOrder->GetCompanyOrders());
		SendCEditString(IDC_EDIT_YIELD, aOrder->GetYield());
		SendCEditString(IDC_EDIT_CAMPAIGNID, aOrder->GetCampaignIDString());
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: COrderDB::LoadScreenFromOrder(COrder* aOrder)";
		PrintOut(err);
		WriteLog(err);
	}
}

void COrderDB::SendCEditString(int id, CString str)
{
	try
	{
		if(mainWin)
		{
			SendMessage(mainWin, S_ORDER_TRANS, id, 0);
			int iLen = str.GetLength();
			for(int x = 0; x < iLen; x++)
			{
				int iLetter = (int) str.GetAt(x);
				SendMessage(mainWin, S_ORDER_TRANS, id, iLetter);
			}		
		}
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: COrderDB::SendCEditString(int id, CString str)";
		PrintOut(err);
		WriteLog(err);
	}
}

void COrderDB::SendCheckBoxState(int id, bool state)
{	
	try
	{
		if(mainWin)
		{
			SendMessage(mainWin, S_ORDER_TRANS, id, (int)state);		
		}
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: COrderDB::SendCheckBoxState(int id, bool state)";
		PrintOut(err);
		WriteLog(err);
	}
}

void COrderDB::SendSQLString(int id, CString str)
{
	try
	{
		if(mainWin)
		{
			SendMessage(mainWin, S_ORDER_TRANS, id, 0);//start string message
			int iLen = str.GetLength();
			for(int x = 0; x < iLen; x++)
			{
				int iLetter = (int) str.GetAt(x);
				SendMessage(mainWin, S_ORDER_TRANS, id, iLetter);//send a letter
			}
			SendMessage(mainWin, S_ORDER_TRANS, id, -1);//end string message		
		}
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: COrderDB::SendSQLString(int id, CString str)";
		PrintOut(err);
		WriteLog(err);
	}
}

void COrderDB::SetOnePass(bool state)
{	
	try
	{
		bOnePass = state;
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: COrderDB::SetOnePass(bool state)";
		PrintOut(err);
		WriteLog(err);
	}
}

void COrderDB::SetVerified(bool state)
{	
	try
	{
		bVerified = state;
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: COrderDB::SetVerified(bool state)";
		PrintOut(err);
		WriteLog(err);
	}
}

CString COrderDB::GetOnePass(void)
{	
	try
	{
		return BoolToCString(bOnePass);
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: COrderDB::GetOnePass(void)";
		PrintOut(err);
		WriteLog(err);
		return "";
	}
}

CString COrderDB::GetVerified(void)
{	
	try
	{
		return BoolToCString(bVerified);
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: COrderDB::GetVerified";
		PrintOut(err);
		WriteLog(err);
	}
}

void COrderDB::SetPriority(bool state)
{	
	try
	{
		bPriority = state;
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: COrderDB::SetPriority(bool state)";
		PrintOut(err);
		WriteLog(err);
	}
}

CString COrderDB::GetPriority(void)
{	
	try
	{
		return BoolToCString(bPriority);
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: COrderDB::GetPriority";
		PrintOut(err);
		WriteLog(err);
	}
}

CString COrderDB::BoolToCString(bool state)
{	
	try
	{
		CString str;
		if(state)
		{
			str = "TRUE";
		}
		else
		{
			str = "FALSE";
		}
		return str;
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: COrderDB::BoolToCString(bool state)";
		PrintOut(err);
		WriteLog(err);
		return "";
	}
}

CString COrderDB::GetSQL(void)
{
	CString sql;	
	try
	{
		sql = "SELECT LastOrder.Ord_MakeGood AS LastOrderQuantity, LastOrder.CountOfPO_Lead_ID AS lastOrderSent, CompanyCountOrders.CountOfOrd_Id AS countOrders, *";
		sql = sql +	" FROM";
		sql = sql +	" (t_orders";
		sql = sql +	" LEFT JOIN";

		sql = sql +	" (SELECT CompanyLastOrder.LastOfOrd_Id, CompanyLastOrder.Ord_CompanyID, Count(t_Processed_Orders.PO_Lead_ID) AS CountOfPO_Lead_ID, t_Orders.Ord_MakeGood";
		sql = sql +	" FROM ([SELECT Last(t_Orders.Ord_Id) AS LastOfOrd_Id, t_Orders.Ord_CompanyID, t_Orders.Ord_Complete";
		sql = sql +	" FROM t_Orders";
		sql = sql +	" GROUP BY t_Orders.Ord_CompanyID, t_Orders.Ord_Complete";
		sql = sql +	" HAVING (((t_Orders.Ord_Complete)=True))]. AS CompanyLastOrder LEFT JOIN t_Processed_Orders ON CompanyLastOrder.LastOfOrd_Id = t_Processed_Orders.PO_Ord_ID) INNER JOIN t_Orders ON CompanyLastOrder.LastOfOrd_Id = t_Orders.Ord_Id";
		sql = sql +	" GROUP BY CompanyLastOrder.LastOfOrd_Id, CompanyLastOrder.Ord_CompanyID, t_Orders.Ord_MakeGood) AS LastOrder";

		sql = sql +	" ON t_orders.Ord_CompanyID = LastOrder.Ord_CompanyID) INNER JOIN ";
		sql = sql +	" (SELECT t_Orders.Ord_CompanyID, Count(t_Orders.Ord_Id) AS CountOfOrd_Id";
		sql = sql +	" FROM t_Orders";
		//sql = sql +	" GROUP BY t_Orders.Ord_CompanyID, t_Orders.Ord_Complete HAVING (((t_Orders.Ord_Complete)=True))) AS CompanyCountOrders";
		sql = sql +	" GROUP BY t_Orders.Ord_CompanyID) AS CompanyCountOrders";
	 
		sql = sql +	" ON t_orders.Ord_CompanyID = CompanyCountOrders.Ord_CompanyID";


		sql = sql +	" WHERE";
		sql = sql +	" (((t_orders.Ord_Complete)=False)";
		sql = sql +	" AND ((t_orders.ReadyToSend)=False)";
		sql = sql +	" AND ((t_orders.Ord_OnePass)=";	
		sql = sql +	GetOnePass() + ")";
		
		sql = sql +	" AND ((t_orders.Ord_Verified)=";
		sql = sql +	 GetVerified() + ") ";
		if(sState.GetLength() > 0)
		{
			sql = sql + " AND ((t_orders.Ord_State)=\"" + GetState() + "\")";
		}
		if(bPriority)
		{
			sql = sql + " AND (priority=";
			sql = sql + GetPriority() + ")";
		}
		if(bNewCustomer)
		{
			sql = sql + " AND (CountOfOrd_Id<4)";
		}
		else
		{
			//sql = sql + " AND (CountOfOrd_Id>3)";
		}
		if(bZeros)
		{
			sql = sql + " AND (CountOfPO_Lead_ID=0)";
		}
		else
		{
			//sql = sql + " AND NOT (CountOfPO_Lead_ID=0)";
		}
		sql = sql + ")";
		//sql = sql +	" ORDER BY t_orders.Ord_Id";
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: COrderDB::GetSQL";
		PrintOut(err);
		WriteLog(err);
	}
	return sql;
}

bool COrderDB::ReadOrders(bool bRetry)
{
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
		  PrintOut("Failed creating record set instance", 0);		 
      }
	  else
	  {		  
		  CString sql;
		  if(bRetry)
		  {
			  sql = GetSQLRetry();
		  }
		  else
		  {
			  sql = GetSQL();
		  }		   
		  SendSQLString(IDC_EDIT_SQL, sql);
		  PrintOut("Submit SQL: " + sql, 0);
		  PrintOut("Reading Orders", 0);
		  WriteLog("Reading Orders");
		  Progress(1, 100);
		pRst->Open(sql.GetBuffer(), strCnn, adOpenStatic, adLockReadOnly, adCmdText);
		
		if(pRst->GetRecordCount()>0)
		{
			pRst->MoveFirst(); 			
			int max = pRst->GetRecordCount();
			int x = 0;
			CString sOrderCount = "Reading Orders:";
			sOrderCount = sOrderCount + GetString(max);
			PrintOut(sOrderCount, 0);
			Progress(x, max);
			while((!pRst->EndOfFile)&& (!stop))
			{	
				//Progress(x, max);
				//PrintOut("Start order Read from file");			
				COrder aOrder;
				aOrder.SetStop(&stop);
				aOrder.SetParentWindow(mainWin);
				aOrder.SetLeadDBMap(&leadDBMap);				
				aOrder.ReadOrder(pRst);
				if(bRetry)
				{
					PrintOut("Read Assigned Leads", 0);
					aOrder.ReadAssignedLeads();
					myCompleteOrderMap.SetAt(aOrder.GetOrderIDLong(), aOrder);
				}
				CLocation myloc = mySettings.GetLocation(aOrder.GetLocationLong());
				if(aOrder.GetVerified())
				{
					aOrder.SetDefaultUseLimit(myloc.GetUseLimitVerified());
					aOrder.SetDefaultAgeLimit(myloc.GetAgeLimitVerified());
				}
				else if (aOrder.GetOnePass())
				{
					aOrder.SetDefaultUseLimit(myloc.GetUseLimitOnePass());
					aOrder.SetDefaultAgeLimit(myloc.GetAgeLimitOnePass());
				}
				else
				{
					aOrder.SetDefaultUseLimit(myloc.GetUseLimitInternet());
					aOrder.SetDefaultAgeLimit(myloc.GetAgeLimitInternet());
				}			
				PrintOut(aOrder.ToString(), 0);
				//LoadScreenFromOrder(&aOrder);			
				Add(aOrder);
				x++;
				Progress(x, max);
				pRst->MoveNext();				
			}			
			PrintOut(GetString(x),0);
			WriteLog(GetString(x) + " Orders Read");
			success = true;
		}
		else 
		{
			PrintOut("No Open Orders", 0);
			WriteLog("No Open Orders");
			Progress(0, 100);
			success = false;
		}
	  }	
	  //CoUninitialize();
	}
	catch(_com_error & ce)
	{
		CString str;		
		str = "Error: " + GetCString(ce.Description());
		PrintOut(str, 0);
		success = false;		
	}  
  return success;
}

CString COrderDB::GetCString(_bstr_t bstrStart)
{	
	return CStringHelper::GetCString(bstrStart);
	/*try
	{
		TCHAR szFinal[255];
		_stprintf(szFinal, _T("%s"), (LPCTSTR)bstrStart);
		CString str;
		str = szFinal;
		return str;
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: functionname";
		PrintOut(err);
		WriteLog(err);
		return "";
	}*/
}

void COrderDB::Reset(void)
{	
	try
	{
		leadDBMap.RemoveAll();
		leadDBMap.SetParentWindow(mainWin);
		leadDBMap.SetStop(false);
		this->SetOnePass(false);
		this->SetPriority(false);
		this->SetStop(false);
		this->SetVerified(false);
		this->SetState("");
		RemoveAll();
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: COrderDB::Reset";
		PrintOut(err);
		WriteLog(err);
	}
}

bool COrderDB::Process(void)
{
	try
	{
		CString sCategory, sGroup;
		if(bVerified) sCategory = "Verifed";
		else if(bOnePass)sCategory = "One-Pass";
		else sCategory = "Internet";
		if(bPriority) sGroup = "Priority";
		else if(bZeros) sGroup = "Sent Zero on last order";
		else if(bNewCustomer) sGroup = "New Customer";
		else sGroup = "Normal";	
		CString str;
		str = "Process (" + sCategory;
		str = str + ") (" + sGroup;
		str = str + ") States: " + sState;
		WriteLog(str);
		bool success = false;		
		if(ReadOrders())
		{
			if(ReadLeads())
			{
				//if(AssignLeads()) success = true;
				if(AssignOrders()) success = true;
			}
		}
		if(success) AssignLeads();
		return success;
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: COrderDB::Process";
		PrintOut(err);
		CWriteLog::WriteFail(err);
		return false;
	}
	
}

bool COrderDB::AssignOrders(void)
{
	bool success = false;	
	try
	{
		WriteLog("Assigning Leads to Orders");
		int max = GetMaxNeeded();
		int iPos = 0;
		Progress(iPos, max);
		
		int x = 0;	
		int count = myOrderMap.GetCount();
		ProgressAll(x, count);
		while((myOrderMap.GetCount() > 0) && (!stop))
		{
			CIDArray ids;
			POSITION pos = myOrderMap.GetStartPosition();
			long    nKey;
			COrder aOrder;
			while ((pos != NULL)&& (!stop))
			{	   
				myOrderMap.GetNextAssoc( pos, nKey, aOrder );		
				ids.Add(aOrder.GetOrderIDLong());
			}
			ids = FatQuickSort(ids);
			CIDArray low, high;
			SplitYield(ids, &low, &high);
			while((low.GetCount()>0)&& (!stop))
			{
				CIDArray temp;
				for(int x = 0; x < low.GetCount(); x++)
				{
					aOrder = GetOrder(low[x]);
					aOrder.SetStop(&stop);
					if(stop) return false;
					long count = aOrder.GetCountMatching();
					char buffer[200];
					_itoa(count, buffer, 10);
					CString str = "Order: " + aOrder.GetOrderID();
					str = str + " has ";
					str = str + buffer;
					str = str + " matching leads";
					//PrintOut(str);
					//LoadScreenFromOrder(&aOrder);
					if(aOrder.AssignLead())
					{
						iPos++;
						Progress(iPos, max);
						//PrintOut("Lead was assigned");
						myOrderMap.SetAt(aOrder.GetOrderIDLong(), aOrder);//update the order					
						temp.Add(low[x]);
					}
					else
					{
						int iCount = aOrder.GetCountAssigned();
						int iQuantity = aOrder.GetQuanityLong();
						Progress(iPos, max);
						PrintOut("Lead was not assigned, assume complete1", aOrder.GetLocationLong());					
						myCompleteOrderMap.SetAt(aOrder.GetOrderIDLong(), aOrder);
						myOrderMap.RemoveKey(aOrder.GetOrderIDLong());
					}
				}
				low = temp;
			}		
			while((high.GetCount()>0)&& (!stop))
			{
				CIDArray temp;
				for(int x = 0; x < high.GetCount(); x++)
				{
					aOrder = GetOrder(high[x]);
					aOrder.SetStop(&stop);
					if(stop) return false;
					long count = aOrder.GetCountMatching();
					char buffer[200];
					_itoa(count, buffer, 10);
					CString str = "Order: " + aOrder.GetOrderID();
					str = str + " has ";
					str = str + buffer;
					str = str + " matching leads";
					//PrintOut(str);
					LoadScreenFromOrder(&aOrder);
					if(aOrder.AssignLead())
					{
						iPos++;
						Progress(iPos, max);
						//PrintOut("Lead was assigned");
						myOrderMap.SetAt(aOrder.GetOrderIDLong(), aOrder);//update the order					
						temp.Add(high[x]);
					}
					else
					{
						int iCount = aOrder.GetCountAssigned();
						int iQuantity = aOrder.GetQuanityLong();
						Progress(iPos, max);
						iPos = iPos + (iQuantity - iCount);
						
						//add the completed order to the completed order map
						myCompleteOrderMap.SetAt(aOrder.GetOrderIDLong(), aOrder);
						myOrderMap.RemoveKey(aOrder.GetOrderIDLong());
					}
				}
				high = temp;
			}
			x = count - myOrderMap.GetCount();
			ProgressAll(x, count);
		}
		success = true;

	

	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: COrderDB::AssignOrders";
		PrintOut(err);
		CWriteLog::WriteFail(err);
		success = false;
	}
	return success;
}

CIDArray COrderDB::FatQuickSort(CIDArray ids)
{
	try
	{
		if(ids.GetCount()<= 1) return ids;
		else
		{
			CIDArray less, equal, greater, result;
			long pivot = ids.GetCount()/2;
			for(int x = 0; x < ids.GetCount(); x++)
			{
				if((GetOrder(ids[x]).GetCountMatching()) < (GetOrder(ids[pivot]).GetCountMatching()))
				{
					less.Add(ids[x]);
				}
				else if ((GetOrder(ids[x]).GetCountMatching()) == (GetOrder(ids[pivot]).GetCountMatching()))
				{
					equal.Add(ids[x]);
				}
				else if ((GetOrder(ids[x]).GetCountMatching()) > (GetOrder(ids[pivot]).GetCountMatching()))
				{
					greater.Add(ids[x]);
				}
			}
			result = FatQuickSort(less);
			result.Append(equal);
			result.Append(FatQuickSort(greater));
			return result;
		}
	

	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: COrderDB::FatQuickSort(CIDArray ids)";
		PrintOut(err);
		CWriteLog::WriteFail(err);
		return CIDArray();
	}
}

COrder COrderDB::GetOrder(long id)
{
	
	try
	{
		COrder aOrder;
		myOrderMap.Lookup(id, aOrder);
		return aOrder;
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: COrderDB::GetOrder(long id)";
		PrintOut(err);
		CWriteLog::WriteFail(err);
		return COrder();
	}	
	
}
COrderDB& COrderDB::operator=(const COrderDB& aCOrderDB)
{
	try
	{
		this->bOnePass = aCOrderDB.bOnePass;
		this->bPriority = aCOrderDB.bPriority;
		this->bVerified = aCOrderDB.bVerified;
		this->leadDBMap = aCOrderDB.leadDBMap;
		this->mainWin = aCOrderDB.mainWin;

		myOrderMap.RemoveAll();
		POSITION pos = aCOrderDB.myOrderMap.GetStartPosition();
		long    nKey;
		COrder aOrder;
		
		while (pos != NULL)
		{	   
			aCOrderDB.myOrderMap.GetNextAssoc( pos, nKey, aOrder );
			myOrderMap.SetAt(aOrder.GetOrderIDLong(), aOrder);
		}

		myCompleteOrderMap.RemoveAll();
		pos = aCOrderDB.myCompleteOrderMap.GetStartPosition();
		while (pos != NULL)
		{	   
			aCOrderDB.myCompleteOrderMap.GetNextAssoc( pos, nKey, aOrder );
			myCompleteOrderMap.SetAt(aOrder.GetOrderIDLong(), aOrder);
		}
		
		this->stop = aCOrderDB.stop;
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: COrderDB::operator=(const COrderDB& aCOrderDB)";
		PrintOut(err);
		CWriteLog::WriteFail(err);
	}
	return *this;
}
COrderDB::COrderDB(const COrderDB& aCOrderDB)
{
	try
	{
		this->bOnePass = aCOrderDB.bOnePass;
		this->bPriority = aCOrderDB.bPriority;
		this->bVerified = aCOrderDB.bVerified;
		this->leadDBMap = aCOrderDB.leadDBMap;
		myOrderMap.RemoveAll();
		POSITION pos = aCOrderDB.myOrderMap.GetStartPosition();
		long    nKey;
		COrder aOrder;		
		while (pos != NULL)
		{	   
			aCOrderDB.myOrderMap.GetNextAssoc( pos, nKey, aOrder );
			myOrderMap.SetAt(aOrder.GetOrderIDLong(), aOrder);
		}
		myCompleteOrderMap.RemoveAll();
		pos = aCOrderDB.myCompleteOrderMap.GetStartPosition();
		while (pos != NULL)
		{	   
			aCOrderDB.myCompleteOrderMap.GetNextAssoc( pos, nKey, aOrder );
			myCompleteOrderMap.SetAt(aOrder.GetOrderIDLong(), aOrder);
		}		
		this->stop = aCOrderDB.stop;
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: COrderDB::COrderDB(const COrderDB& aCOrderDB)";
		PrintOut(err);
		CWriteLog::WriteFail(err);
	}
}

COrderDB COrderDB::GetOrdersWithPastZeroSent(void)
{
	//pull out the orders that got zero leads on their last completed order
	return COrderDB();
}

COrderDB COrderDB::GetOrdersWithLowPartialsSent(void)
{
	//pull out orders that had a very low % of their order quantity sent on their last order
	return COrderDB();
}

double COrderDB::GetLastPercentSent(long iCompID)
{
	double ratio = -1;
	///////////////////////////////////////////////////////////////
	
	
    try
    {
		HRESULT hr = S_OK;
		char buffer[200];
		_RecordsetPtr pRst;
		//PrintOut("Init database connection", 0);
         CoInitialize(NULL);
          // Define string variables.
		 
		 _bstr_t strCnn("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=z:\\Mortgage.mdb;Jet OLEDB:Database Password=cepher;");

      // Call Create instance to instantiate the Record set
      hr = pRst.CreateInstance(__uuidof(Recordset));

      if(FAILED(hr))
      {            
		  PrintOut("Failed creating record set instance", 0);		 
      }
	  else
	  {
		  CString sql;
			sql = sql + "SELECT ";
			sql = sql + "CompanyLastOrder.LastOfOrd_Id, "; 
			sql = sql + "CompanyLastOrder.Ord_CompanyID, ";
			sql = sql + "Count(t_Processed_Orders.PO_Lead_ID) AS CountOfPO_Lead_ID, ";
			sql = sql + "t_Orders.Ord_MakeGood ";

			sql = sql + "FROM ";
			sql = sql + "( ";
			sql = sql + "(SELECT Last(t_Orders.Ord_Id) AS LastOfOrd_Id, t_Orders.Ord_CompanyID, t_Orders.Ord_Complete ";
			sql = sql + "FROM t_Orders ";
			sql = sql + "GROUP BY t_Orders.Ord_CompanyID, t_Orders.Ord_Complete ";
			sql = sql + "HAVING (((t_Orders.Ord_Complete)=True))) AS CompanyLastOrder ";

			sql = sql + "LEFT JOIN ";
			sql = sql + "t_Processed_Orders ON CompanyLastOrder.LastOfOrd_Id = t_Processed_Orders.PO_Ord_ID) ";
			sql = sql + "INNER JOIN t_Orders ON CompanyLastOrder.LastOfOrd_Id = t_Orders.Ord_Id ";
			sql = sql + "GROUP BY CompanyLastOrder.LastOfOrd_Id, CompanyLastOrder.Ord_CompanyID, t_Orders.Ord_MakeGood ";
			sql = sql + "HAVING (((CompanyLastOrder.Ord_CompanyID)=";
			CString id;
			sprintf(buffer,"%d",iCompID);
			id = buffer;
			sql = sql + id;
			sql = sql + ")) ";

		  PrintOut("Submit SQL: " + sql, 0);
		pRst->Open(sql.GetBuffer(), strCnn, adOpenStatic, adLockReadOnly, adCmdText);

		if(!pRst->EndOfFile) pRst->MoveFirst(); 
		else 
		{
			PrintOut("No Last Order", 0);
			return -1;
		}
		
		while((!pRst->EndOfFile)&& (!stop))
		{	
			
			PrintOut("Start last order Read from file", 0);			
			//get count here
			double ordered = GetFieldLong(pRst, "Ord_MakeGood");
			double sent = GetFieldLong(pRst, "CountOfPO_Lead_ID");
			
			sprintf(buffer, "Ordered: %f",ordered);
			PrintOut(buffer, 0);
			sprintf(buffer, "Sent: %f",sent);
			PrintOut(buffer, 0);
			if(ordered>0)
			{
				ratio = sent / ordered;
				sprintf(buffer, "Ratio: %f",ratio);
				PrintOut(buffer, 0);
			}
			else
			{
				ratio = -1;// return 0 if the order quantity is 0
			}
			
			pRst->MoveNext();
			
		}		
	  }	
	  CoUninitialize();
	}
	catch(_com_error & ce)
	{
		CString str;
		
		str = "Error COrderDB::GetLastPercentSent(long iCompID): " + GetCString(ce.Description());
		PrintOut(str, 0);
		WriteLog(str);
		CWriteLog::WriteFail(str);
		
	}  
  
	///////////////////////////////////////////////////////////////
	return ratio;
}

void COrderDB::SetZeros(bool state)
{	
	try
	{
		bZeros = state;
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: COrderDB::SetZeros(bool state)";
		PrintOut(err);
		CWriteLog::WriteFail(err);
	}
}

void COrderDB::SetNewCustomer(bool state)
{	
	try
	{
		bNewCustomer = state;
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: COrderDB::SetNewCustomer(bool state)";
		PrintOut(err);
		CWriteLog::WriteFail(err);
	}
}

void COrderDB::SplitYield(CIDArray ids, CIDArray *low, CIDArray *high)
{	
	try
	{
		long max = GetMaxNeeded();
		for(long x = 0; x < ids.GetCount(); x++)
		{
			COrder aOrder = GetOrder(ids[x]);
			
			double yield = aOrder.GetYieldDouble();
			if(yield <= 0.25)
			{
				low->Add(ids[x]);
			}
			else
			{
				high->Add(ids[x]);
			}
		}
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: COrderDB::SplitYield(CIDArray ids, CIDArray *low, CIDArray *high)";
		PrintOut(err);
		CWriteLog::WriteFail(err);
	}
}

void COrderDB::GenerateRTFs(void)
{
	try
	{
		WriteLog("Generate Documents");
		CLocation myloc;

		POSITION pos = myOrderMap.GetStartPosition();
		long    nKey;
		COrder aOrder;
		aOrder.SetStop(&stop);	

		PrintOut("Generate RTFs for Complete Orders: ", 0);
		pos = myCompleteOrderMap.GetStartPosition();
		int x = 0;
		int count = myCompleteOrderMap.GetCount();
		ProgressAll(x, count);
		while ((pos != NULL) && (!stop))
		{	   
			myCompleteOrderMap.GetNextAssoc( pos, nKey, aOrder );

			myloc = mySettings.GetLocation(aOrder.GetLocationLong());
			if(aOrder.GetCountAssigned()>0)
			{
				if(aOrder.GetSendWord() || aOrder.GetSendMini1003() || (!aOrder.GetSendEmailNote() && !aOrder.GetSendWord() && !aOrder.GetSendMini1003() && !aOrder.GetExcel()))
				{
					CRTFWriter aWriter(myloc.GetDesDir(), myloc.GetTemplateDir(), this->mainWin); 
					aWriter.Write(aOrder);
				}
				if(aOrder.GetExcel())
				{
					CCSVWriter aWriter(myloc.GetDesDir(),this->mainWin);
					aWriter.Write(aOrder, myloc);
				}
			}
			else
			{
				CTxtWriter aWriter(myloc.GetDesDir(), this->mainWin); 
				aWriter.Write(aOrder, myloc);
			}
			x++;
			ProgressAll(x, count);
		}
		CString msg = "Generated " + GetString(x);
		msg = msg + " Documents";
		WriteLog(msg);
	

	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: COrderDB::GenerateRTFs";
		PrintOut(err);
		CWriteLog::WriteFail(err);
	}
}

bool COrderDB::SendMail(CString brokerEmail, CString agentname, CString agentEmail, CString company, CString contact, int div, CString orderid)
{
	try
	{
		CLocation myLoc;	
		myLoc = mySettings.GetLocation(div);
		CString body;
		CString defaultReturnEmail;
		defaultReturnEmail = myLoc.GetReturnEmail();
		body = myLoc.GetBody(brokerEmail, agentname, agentEmail, company, contact, orderid, "");
		
		bool success = false;
		CString to;
		CString agent;
		CString subject;	

		subject = myLoc.GetSubject(brokerEmail, agentname, agentEmail, company, contact, orderid, "");	

		if(testing)
		{
			to = "test@country-loan.com";
			agent = "test@country-loan.com";
			
		}
		else
		{
			to = brokerEmail;
			if(to == "")		
			to = "kenco_computers@yahoo.com";
			agent = agentEmail;
		}
		ESMTP mail;
		mail.SetParentWindow(mainWin);
		//mail.SetName("Testing");
		mail.SetName(agentname);
		//mail.SetAddress("test@country-loan.com");
		if(agentEmail.GetLength()>0)
		{
			mail.SetAddress(agentEmail);
		}
		else
		{
			mail.SetAddress(defaultReturnEmail);
		}	
		mail.SetTo(to);
		//mail.SetBCC("test@country-loan.com");
		mail.SetSubject(subject);	
		//mail.m_sHost = "216.219.253.246";
		//mail.SetUserName("test@country-loan.com");
		//mail.SetPassword("test123");
		//PrintOut(mySettings.GetSmtp());
		mail.SetHost(mySettings.GetSmtp());
		mail.SetUserName(mySettings.GetUsername());
		mail.SetPassword(mySettings.GetPassword());
		mail.SetBody(body);
		CString filename = myLoc.GetDesDir() + "\\";
		filename = filename + orderid;
		filename = filename + ".rtf";
		mail.AddFilename(filename);

		PrintOut("Sending email", div);
		mail.SetDiv(div);
		if(mail.Send())
		{
			PrintOut("Email sent to: " + contact + " At: " + company + " Email: " + mail.GetTo() + " and BCC to: " + mail.GetBCC(), div);
			PrintOut(" ", div);
			success = true;
			//AfxMessageBox("Message send to: " + mail.m_sTo , MB_OK);
		}
		else
		{
			PrintOut("Error sending Email. Make sure you are connected to the internet.", div);
			//AfxMessageBox("Error sending Email. Make sure you are connected to the internet before hitting Send." , MB_OK);
			success = false;
		}
		return success;	

	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: COrderDB::SendMail(CString brokerEmail, CString agentname, CString agentEmail, CString company, CString contact, int div, CString orderid)";
		PrintOut(err);
		CWriteLog::WriteFail(err);
		return false;
	}
}

void COrderDB::SendOrders(void)
{
	try
	{
		WriteLog("Sending Orders to Clients");	
		long    nKey;
		COrder aOrder;
		aOrder.SetStop(&stop);
		
		PrintOut("Email Complete Orders: ", 0);
		POSITION pos = myCompleteOrderMap.GetStartPosition();
		int x = 0;
		while ((pos != NULL) && (!stop))
		{	   
			myCompleteOrderMap.GetNextAssoc( pos, nKey, aOrder );
			PrintOut("Send Order:" + aOrder.ToString(), aOrder.GetLocationLong());
			//MarkOrderReadyToSend(aOrder);
			if(aOrder.GetSendFax())
			{
				if(aOrder.GetCountAssigned()>0)
				{
					SendFax(aOrder);
				}
				else
				{
					PrintOut("Send 0 leads Fax", aOrder.GetLocationLong());
					SendFax(aOrder);
				}
			}
			if(aOrder.GetPrint())
			{
				Print(aOrder);
			}
			if(!aOrder.GetSendFax())// && !aOrder.GetSendEmailNote())
			{
				if(SendEmail(aOrder))
				{
					MarkOrderComplete(aOrder);
				}
				else
				{
					PrintOut("Order not marked complete. There was a problem sending the email", aOrder.GetLocationLong());
				}
			}
			x++;
		}
		CString msg = "Sent " + GetString(x);
		msg = msg + " Orders";
		WriteLog(msg);
	

	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: COrderDB::SendOrders";
		PrintOut(err);
		CWriteLog::WriteFail(err);
	}
}

bool COrderDB::SendEmail(COrder aOrder)
{
	try
	{
		CLocation myLoc;	
		myLoc = mySettings.GetLocation(aOrder.GetLocationLong());
		CString body;
		CString defaultReturnEmail;
		defaultReturnEmail = myLoc.GetReturnEmail();
		//aOrder.GetEmail(), aOrder.GetAgent(), aOrder.GetAgentEmail(), aOrder.GetCompanyName(), aOrder.GetContact(), aOrder.GetLocationLong(), aOrder.GetOrderID()
			
		bool success = false;
		CString to;
		CString agent;
		CString subject;	
		if(aOrder.GetAdCampaign()) subject = "Ad Campaign: ";

		subject = subject + myLoc.GetSubject(aOrder.GetEmail(), aOrder.GetAgent(), aOrder.GetAgentEmail(),  aOrder.GetCompanyName(), aOrder.GetContact(), aOrder.GetOrderID(), aOrder.GetStates());	

		if(testing)
		{
			to = "test@country-loan.com";
			agent = "test@country-loan.com";
			
		}
		else
		{
			to = aOrder.GetEmail();
			if(to == "")		
			to = "kenco_computers@yahoo.com";
			agent = aOrder.GetAgentEmail();
		}
		ESMTP mail;
		mail.SetParentWindow(mainWin);
		//mail.SetName("Testing");
		//mail.SetName(aOrder.GetAgent());
		mail.SetName("Your Order - DO NOT REPLY");
		//mail.SetAddress("test@country-loan.com");
		//if(aOrder.GetAgentEmail().GetLength()>0)
		//{
			//mail.SetAddress(aOrder.GetAgentEmail());
		//}
		//else
		//{
			mail.SetAddress(defaultReturnEmail);
		//}	
		mail.SetTo(to);
		//mail.SetBCC("test@country-loan.com");
		mail.SetSubject(subject);	
		//mail.m_sHost = "216.219.253.246";
		//mail.SetUserName("test@country-loan.com");
		//mail.SetPassword("test123");
		//PrintOut(mySettings.GetSmtp());
		mail.SetHost(mySettings.GetSmtp());
		mail.SetUserName(mySettings.GetUsername());
		mail.SetPassword(mySettings.GetPassword());
		if(aOrder.GetCountAssigned()>0)
		{
			body = myLoc.GetBody(aOrder.GetEmail(), aOrder.GetAgent(), aOrder.GetAgentEmail(),  aOrder.GetCompanyName(), aOrder.GetContact(), aOrder.GetOrderID(), aOrder.GetStates());
			if(aOrder.GetSendEmailNote())
			{
				body = body + aOrder.GetEmailNote(myLoc);			
			}
			if( (aOrder.GetSendWord())|| (aOrder.GetSendMini1003()) ||
				( (!aOrder.GetSendEmailNote()) 
				&& (!aOrder.GetSendFax()) 
				&& (!aOrder.GetSendMini1003()) 
				&& (!aOrder.GetExcel())) )
			{			
				CString filename = myLoc.GetDesDir() + "\\";
				filename = filename + aOrder.GetOrderID();
				filename = filename + ".rtf";
				if(!DoesFileExist(filename)) return false;
				mail.AddFilename(filename);
			}
			if(aOrder.GetExcel())
			{
				CString filename = myLoc.GetDesDir() + "\\";
				filename = filename + aOrder.GetOrderID();
				filename = filename + ".csv";
				if(!DoesFileExist(filename)) return false;
				mail.AddFilename(filename);
			}
			mail.SetBody(body);
		}
		else
		{
			body = aOrder.GetZeroBody(myLoc);
			mail.SetBody(body);
		}
		PrintOut("Sending email", aOrder.GetLocationLong());
		mail.SetDiv(aOrder.GetLocationLong());
		if(mail.Send())
		{
			WriteLog("Email sent to: " + aOrder.GetContact() + " At: " + aOrder.GetCompanyName() + " Email: " + mail.GetTo() + " and BCC to: " + mail.GetBCC());
			PrintOut("Email sent to: " + aOrder.GetContact() + " At: " + aOrder.GetCompanyName() + " Email: " + mail.GetTo() + " and BCC to: " + mail.GetBCC(), aOrder.GetLocationLong());
			PrintOut(" ", aOrder.GetLocationLong());
			success = true;
			//AfxMessageBox("Message send to: " + mail.m_sTo , MB_OK);
		}
		else
		{
			WriteLog("Error sending Email. Make sure you are connected to the internet");
			PrintOut("Error sending Email. Make sure you are connected to the internet.", aOrder.GetLocationLong());
			
			//AfxMessageBox("Error sending Email. Make sure you are connected to the internet before hitting Send." , MB_OK);
			success = false;
		}
		return success;	

	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: COrderDB::SendEmail(COrder aOrder)";
		PrintOut(err);
		CWriteLog::WriteFail(err);
		return false;
	}
	//return SendMail(aOrder.GetEmail(), aOrder.GetAgent(), aOrder.GetAgentEmail(), aOrder.GetCompanyName(), aOrder.GetContact(), aOrder.GetLocationLong(), aOrder.GetOrderID());;
}

void COrderDB::MarkOrderComplete(COrder aOrder)
{	
	HRESULT hr = S_OK;
    try
    {
		_RecordsetPtr pRst;
		//PrintOut("Init database connection", aOrder.GetLocationLong());
         CoInitialize(NULL);
          // Define string variables.
		 
		 _bstr_t strCnn("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=z:\\Mortgage.mdb;Jet OLEDB:Database Password=cepher;");

      // Call Create instance to instantiate the Record set
      hr = pRst.CreateInstance(__uuidof(Recordset));

      if(FAILED(hr))
      {            
		  PrintOut("Failed creating record set instance", aOrder.GetLocationLong());		 
      }
	  else
	  {
		  CString sql;
			sql = "SELECT * FROM t_Orders WHERE Ord_Id=" + aOrder.GetOrderID();			

		  PrintOut("Submit SQL: " + sql, aOrder.GetLocationLong());
		pRst->Open(sql.GetBuffer(), strCnn, adOpenStatic, adLockOptimistic, adCmdText);

		if(!pRst->EndOfFile) pRst->MoveFirst(); 
		else 
		{
			PrintOut("No Order Found", aOrder.GetLocationLong());			
		}		
		if((!pRst->EndOfFile)&& (!stop))
		{			
			PrintOut("Mark Order complete", aOrder.GetLocationLong());			
			pRst->Fields->Item["Ord_Complete"]->Value = true;
			pRst->Fields->Item["Ord_Leads_Sent"]->Value = aOrder.GetCountAssigned();
			COleDateTime today = COleDateTime::GetCurrentTime();			
			pRst->Fields->Item["Ord_Sent_On"]->Value = _variant_t(today);			
			pRst->Update();
						
		}		
	  }	
	  CoUninitialize();
	}
	catch(_com_error & ce)
	{
		CString str;
		
		str = "Mark Order Complete Error: " + GetCString(ce.Description());
		CWriteLog::WriteFail(str);
		
	}  
}

void COrderDB::MarkOrderReadyToSend(COrder aOrder)
{	
	
    try
    {
		HRESULT hr = S_OK;
		CoInitialize(NULL);
		_RecordsetPtr pRst;
		//PrintOut("Init database connection", aOrder.GetLocationLong());
         
          // Define string variables.
		 
		 _bstr_t strCnn("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=z:\\Mortgage.mdb;Jet OLEDB:Database Password=cepher;");
      

      // Call Create instance to instantiate the Record set
      hr = pRst.CreateInstance(__uuidof(Recordset));

      if(FAILED(hr))
      {            
		  PrintOut("Failed creating record set instance", aOrder.GetLocationLong());		 
      }
	  else
	  {
		  CString sql;
			sql = "SELECT * FROM t_Orders WHERE Ord_Id=" + aOrder.GetOrderID();			

		  PrintOut("Submit SQL: " + sql, aOrder.GetLocationLong());
		pRst->Open(sql.GetBuffer(), strCnn, adOpenStatic, adLockOptimistic, adCmdText);

		if(!pRst->EndOfFile) pRst->MoveFirst(); 
		else 
		{
			PrintOut("No Order Found", aOrder.GetLocationLong());			
		}		
		if((!pRst->EndOfFile)&& (!stop))
		{				
			PrintOut("Mark Order Ready to Send", aOrder.GetLocationLong());
			
			pRst->Fields->Item["ReadyToSend"]->Value = true;
			pRst->Fields->Item["Ord_Leads_Sent"]->Value = aOrder.GetCountAssigned();			
			pRst->Update();						
		}		
	  }	
	  CoUninitialize();  
	}
	catch(_com_error & ce)
	{
		CString str;		
		str = "Mark Order Complete Error: " + GetCString(ce.Description());
		PrintOut(str, 0);	
		CWriteLog::WriteFail(str);	
	}  
}

void COrderDB::SendFax(COrder aOrder)
{
	try
	{
		CLocation myloc = mySettings.GetLocation(aOrder.GetLocationLong());
		CFaxJob aJob = aOrder.GetFaxJob(myloc);
		aJob.SendTo(mainWin);	

	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: COrderDB::SendFax(COrder aOrder)";
		PrintOut(err);
		CWriteLog::WriteFail(err);
	}
}
CString COrderDB::GetIntheBody(COrder aOrder)
{
	//CIDArray leadIDs = aOrder.GetAssignedLeadIDs();
	//CLeadDB* pt_leadDB = aOrder.GetLeadDBMap();
	//int pagecount = 1;
	//for(int x = 0; x < leadIDs.GetCount(); x++)
	//{
		//CLead aLead = pt_leadDB->GetLead(leadIDs[x]);
	//}
	return CString();
}

void COrderDB::Print(COrder aOrder)
{
	try
	{
		CLocation loc = mySettings.GetLocation(aOrder.GetLocationLong());
		CString sPath = loc.GetDesDir() + "\\";
		sPath = sPath + aOrder.GetOrderID();
		if((aOrder.GetSendWord()) || ((!aOrder.GetSendWord())&&(!aOrder.GetExcel())))
		{
			CString rtfdoc;
			rtfdoc = sPath + ".rtf";
			ShellExecute(NULL, "print", rtfdoc, NULL, NULL, SW_SHOWNORMAL);
			if(DoesFileExist(rtfdoc)) MarkOrderComplete(aOrder);
		}
		if(aOrder.GetExcel())
		{
			CString csvdoc;
			csvdoc = sPath + ".csv";
			
			ShellExecute(NULL, "print", csvdoc, NULL, NULL, SW_SHOWNORMAL);
			if(DoesFileExist(csvdoc)) MarkOrderComplete(aOrder);
		}
		
	

	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: COrderDB::Print(COrder aOrder)";
		PrintOut(err);
		CWriteLog::WriteFail(err);
	}
	
}

void COrderDB::Progress(int iPos, int iMax)
{
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
		err = "Error: COrderDB::Progress(int iPos, int iMax)";
		PrintOut(err);
		CWriteLog::WriteFail(err);
	}
}

long COrderDB::GetCountAssigned(void)
{
	try
	{
		POSITION pos = myCompleteOrderMap.GetStartPosition();
		long total = 0;
		long    nKey;
		COrder aOrder;
		while (pos != NULL)
		{	   
			myCompleteOrderMap.GetNextAssoc( pos, nKey, aOrder );
			total = total + aOrder.GetCountAssigned();		
		}
		return total;
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: COrderDB::GetCountAssigned";
		PrintOut(err);
		CWriteLog::WriteFail(err);
		return 0;
	}
}

CString COrderDB::GetString(long iNum)
{
	try
	{
		char buffer[200];
		sprintf(buffer, "%d", iNum);
		return CString(buffer);	
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: COrderDB::GetString(long iNum)";
		PrintOut(err);
		CWriteLog::WriteFail(err);
		return CString();
	}
}

void COrderDB::PrintOrderTotals(void)
{
	try
	{
		POSITION pos = myCompleteOrderMap.GetStartPosition();	
		long    nKey;
		COrder aOrder;
		while (pos != NULL)
		{	   
			myCompleteOrderMap.GetNextAssoc( pos, nKey, aOrder );
			CString sCount, sID, sQuantity, msg;
			sID = aOrder.GetOrderID();
			sCount = GetString(aOrder.GetCountAssigned());	
			sQuantity = aOrder.GetQuantity();
			msg = "Order:(" + sID + ")";
			msg = msg + "Assigned:(" + sCount + ")";
			msg = msg + "Ordered: (" + sQuantity + ")";
			PrintOut(msg, 0);
			if(aOrder.GetCountAssigned()> aOrder.GetQuanityLong())
			{
				WriteLog(msg);
				//PrintFail(msg);
			}		
		}
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: COrderDB::PrintOrderTotals";
		PrintOut(err);
		CWriteLog::WriteFail(err);
	}
}

void COrderDB::PrintFail(CString str)
{
	try
	{
		if(mainWin)
		{
			SendMessage(mainWin, EMAIL_FAIL_START, 0, 0);
			int iLen = str.GetLength();
			for(int x = 0; x < iLen; x++)
			{
				int iLetter = (int) str.GetAt(x);
				SendMessage(mainWin, EMAIL_FAIL_CHAR, iLetter, 0);
			}
			SendMessage(mainWin, EMAIL_FAIL_END, 0, 0);
		}	

	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: COrderDB::PrintFail(CString str)";
		PrintOut(err);
		CWriteLog::WriteFail(err);
	}
}

bool COrderDB::ReadOrders(void)
{	
	try
	{
		return ReadOrders(false);
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: COrderDB::ReadOrders";
		PrintOut(err);
		CWriteLog::WriteFail(err);
		return false;
	}
}

CString COrderDB::GetSQLRetry(void)
{	
	try
	{
		CString sql;
		sql = "SELECT t_Orders.* ";
		sql = sql + "FROM t_Orders ";
		sql = sql + "WHERE (((t_Orders.ReadyToSend)=True) AND ((t_Orders.Ord_Complete)=False))";
		return sql;

	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: COrderDB::GetSQLRetry";
		PrintOut(err);
		CWriteLog::WriteFail(err);
		return CString();
	}
}
void COrderDB::DeleteAssignedLeads(void)
{
	
    try
    {
		HRESULT hr = S_OK;	
		WriteLog("Revoving any partially processed leads from orders");
		//PrintOut("Init database connection");		
         CoInitialize(NULL);
          // Define string variables.
		 
		 _bstr_t strCnn("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=z:\\Mortgage.mdb;Jet OLEDB:Database Password=cepher;");

       _RecordsetPtr pRst = NULL;

      // Call Create instance to instantiate the Record set
      hr = pRst.CreateInstance(__uuidof(Recordset));
      if(FAILED(hr))      {
            
		  PrintOut("Failed creating record set instance",0);		  
      }
	  else
	  {		
		CString sql = "DELETE t_Processed_Orders.PO_Ord_ID, t_Processed_Orders.PO_Lead_ID, t_Orders.Ord_Complete, t_Orders.ReadyToSend, t_Processed_Orders.*";
		sql = sql + " FROM t_Orders INNER JOIN t_Processed_Orders ON t_Orders.Ord_Id = t_Processed_Orders.PO_Ord_ID";
		sql = sql + " WHERE (((t_Orders.Ord_Complete)=False) AND ((t_Orders.ReadyToSend)=False))";
		
		PrintOut("Revoving any partially processed leads from orders. Submit SQL:" + sql,0);
		
		pRst->Open(sql.GetBuffer(), strCnn, adOpenStatic, adLockOptimistic, adCmdText);	
		PrintOut("Finished removing partials",0);
		
	  }
	  //pRst->Close();
	  CoUninitialize();
	  
	}
	catch(_com_error & ce)
	{
		CString str;		
		str = "Error: " + GetCString(ce.Description());
		PrintOut(str,0);
		CWriteLog::WriteFail(str);
		//WriteLog(str);
		
	}
}
void COrderDB::WriteLog(CString str)
{	
	CWriteLog::Write(str);
	///////////////////////////////////////////////////////////////////
	/*HRESULT hr;
	hr = S_OK;	
	try
    {
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
			PrintOut("Failed creating batcher log record set instance",0);
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
		PrintOut(str,0);
		AfxMessageBox(str);
		
	}*/
	///////////////////////////////////////////////////////////////////
	
}

bool COrderDB::ReadReadytoSend(bool bGenDoc)
{
	bool success = false;	
	
	try
    {
		HRESULT hr = S_OK;	
		_RecordsetPtr pRst;	
		CoInitialize(NULL);
          // Define string variables.		 
		_bstr_t strCnn("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=z:\\Mortgage.mdb;Jet OLEDB:Database Password=cepher;");
		
		// Call Create instance to instantiate the Record set
		hr = pRst.CreateInstance(__uuidof(Recordset));

		if(FAILED(hr))
		{            
			PrintOut("Failed creating record set instance", 0);		 
		}
		else
		{		  
			CString sql;		  
			sql = GetSQLRetry();		  
		   
			SendSQLString(IDC_EDIT_SQL, sql);
			PrintOut("Submit SQL: " + sql, 0);
			PrintOut("Reading Orders", 0);
			WriteLog("Reading Orders");
			Progress(1, 100);
			pRst->Open(sql.GetBuffer(), strCnn, adOpenStatic, adLockReadOnly, adCmdText);
		
			if(pRst->GetRecordCount()>0)
			{
				
				pRst->MoveFirst(); 
				
				int max = pRst->GetRecordCount();
				int x = 0;
				CString sOrderCount = "Reading Orders:";
				sOrderCount = sOrderCount + GetString(max);
				PrintOut(sOrderCount, 0);
				ProgressAll(x, max);
				while((!pRst->EndOfFile)&& (!stop))
				{	
					//Progress(x, max);
					//PrintOut("Start order Read from file");			
					COrder aOrder;
					aOrder.SetStop(&stop);
					aOrder.SetParentWindow(mainWin);
					aOrder.SetLeadDBMap(&leadDBMap);
					
					aOrder.ReadOrder(pRst);

					if(bGenDoc)
					{
						PrintOut("Read Assigned Leads", 0);
						aOrder.ReadAssignedLeads();
						CLocation myloc = mySettings.GetLocation(aOrder.GetLocationLong());
						if(aOrder.GetCountAssigned()>0)
						{
							if(aOrder.GetSendWord() || aOrder.GetSendMini1003() || (!aOrder.GetSendEmailNote() && !aOrder.GetSendWord() && !aOrder.GetSendMini1003() && !aOrder.GetExcel()))
							{
								CRTFWriter aWriter(myloc.GetDesDir(), myloc.GetTemplateDir(), this->mainWin); 
								aWriter.Write(aOrder);
							}
							if(aOrder.GetExcel())
							{
								CCSVWriter aWriter(myloc.GetDesDir(),this->mainWin);
								aWriter.Write(aOrder, myloc);
							}
						}
						else
						{
							CTxtWriter aWriter(myloc.GetDesDir(), this->mainWin); 
							aWriter.Write(aOrder, myloc);
						}

					}
	///////////////////////////////////////////////////////////////////////////////////////////////////
					if(aOrder.GetSendFax())
					{
						if(aOrder.GetCountAssigned()>0)
						{
							SendFax(aOrder);
						}
						else
						{
							PrintOut("Send 0 leads Fax", aOrder.GetLocationLong());
							SendFax(aOrder);
						}
					}
					if(aOrder.GetPrint())
					{
						Print(aOrder);
					}
					if(!aOrder.GetSendFax())// && !aOrder.GetSendEmailNote())
					{
						if(SendEmail(aOrder))
						{
							MarkOrderComplete(aOrder);
						}
						else
						{
							PrintOut("Order not marked complete. There was a problem sending the email", aOrder.GetLocationLong());
						}
					}
	///////////////////////////////////////////////////////////////////////////////////////////////////
					
					PrintOut(aOrder.ToString(), 0);					
					x++;
					ProgressAll(x, max);
					pRst->MoveNext();
					
				}
			
			PrintOut(GetString(x),0);
			WriteLog(GetString(x) + " Orders Read");
			success = true;
			}
			else 
			{
				PrintOut("No Open Orders", 0);
				WriteLog("No Open Orders");
				Progress(0, 100);
				success = false;
			}
		}	
		CoUninitialize();
	}
	catch(_com_error & ce)
	{
		CString str;
		
		str = "Error: " + GetCString(ce.Description());
		PrintOut(str, 0);
		success = false;
		CWriteLog::WriteFail(str);
		
	}  
	return success;
}

void COrderDB::ProgressAll(int iPos, int iMax)
{
	try
	{
		if(mainWin)
		{
			SendMessage(mainWin, ID_PROGRESS_OVERALL, iPos, iMax);
		}
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: COrderDB::ProgressAll(int iPos, int iMax)";
		PrintOut(err);
		CWriteLog::WriteFail(err);
	}	
}

void COrderDB::PrintOut(CString str)
{	
	try
	{
		PrintOut(str, 0);
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: COrderDB::PrintOut(CString str)";
		PrintOut(err);
		CWriteLog::WriteFail(err);
	}
}



void COrderDB::SetState(CString sStateAbr)
{
	sState = sStateAbr;
}
CString COrderDB::GetState()
{
	return sState;
}
bool COrderDB::DoesFileExist(CString sFilename)
{
	DWORD dwAttr = GetFileAttributes(sFilename);
	if(dwAttr == 0xffffffff)
	{
		return false;
	}
	else
	{
		return true;
	}

	
}
