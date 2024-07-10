#include "StdAfx.h"
#include ".\stats.h"

CStats::CStats(void)
: output()
, iVerified(0)
, iOnePass(0)
, iInternet(0)
, iInternetOrdered(0)
, iOnePassOrdered(0)
, iVerifiedOrdered(0)
, iSentZero(0)
, iSentAll(0)
, iSentPartial(0)
, iTotalCampaigns(0)
, fMean(0)
{
}

CStats::~CStats(void)
{
}

CString CStats::GetSQLLeadCounts(void)
{
	//get the SQL that count leads by groups
	//output.Print("CStats::GetSQLLeadCounts()");
	CString sql;
	try
	{
		sql = "SELECT Count(t_Processed_Orders.PO_Lead_ID) AS CountOfLeads \r\n";
		sql = sql + ", t_Orders.Ord_Verified \r\n";
		sql = sql + ", t_Orders.Ord_OnePass \r\n";
		
		sql = sql + "FROM t_Orders INNER JOIN t_Processed_Orders \r\n";
		sql = sql + "ON t_Orders.Ord_Id = t_Processed_Orders.PO_Ord_ID \r\n";
		sql = sql + "GROUP BY t_Orders.Ord_Complete \r\n";
		sql = sql + ", t_Orders.Ord_Verified \r\n";
		sql = sql + ", t_Orders.Ord_OnePass \r\n";
		sql = sql + ", Month([Ord_Date]) \r\n";
		sql = sql + ", Day([Ord_Date]) \r\n";
		sql = sql + ", Year([Ord_Date]) \r\n";
		sql = sql + "HAVING (((t_Orders.Ord_Complete)=True) \r\n";
		sql = sql + "AND ((Month([Ord_Date]))=Month(Now())) \r\n";
		sql = sql + "AND ((Day([Ord_Date]))=Day(Now())) \r\n";
		sql = sql + "AND ((Year([Ord_Date]))=Year(Now()))) \r\n";

	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: CStats::GetSQLLeadCounts";
		output.Print(err);
		//PrintOut(err);
		CWriteLog::WriteFail(err);
		
	}
	return sql;
}

void CStats::SetParentWindow(HWND m_wnd)
{
	
	try
	{
		output.SetParentWindow(m_wnd);
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: CStats::SetParentWindow(HWND m_wnd)";
		AfxMessageBox(err);
		CWriteLog::WriteFail(err);
	}
}

void CStats::PrintOut(CString str)
{
	try
	{
		output.Print(str);
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: PrintOut(CString str)";
		AfxMessageBox(err);
		CWriteLog::WriteFail(err);
	}
}

CString CStats::GetCountVerified(void)
{
	try
	{
		return CStringHelper::GetCString(iVerified);
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: CStats::GetCountVerifed";
		PrintOut(err);
		CWriteLog::WriteFail(err);
	}
}

CString CStats::GetCountOnePass(void)
{
	try
	{
		return CStringHelper::GetCString(iOnePass);
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: CStats::GetCountOnePass";
		PrintOut(err);
		CWriteLog::WriteFail(err);
		return CString();
	}
}

CString CStats::GetCountInternet(void)
{
	try
	{
		return CStringHelper::GetCString(iInternet);
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: CStats::GetCountInternet";
		PrintOut(err);
		CWriteLog::WriteFail(err);
		return CString();
	}
}

void CStats::ReadTodaysLeadsSent(void)
{
	try
    {
		HRESULT hr = S_OK;
		
        CoInitialize(NULL);
        // Define string variables.
		 
		_bstr_t strCnn("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=z:\\Mortgage.mdb;Jet OLEDB:Database Password=cepher;");

		_RecordsetPtr pRst = NULL;

		// Call Create instance to instantiate the Record set
		hr = pRst.CreateInstance(__uuidof(Recordset));

		if(FAILED(hr))
		{	            
			PrintOut("Failed creating settings record set instance");	            
		}
		else
		{
			
			CString sql;
			sql = GetSQLLeadCounts(); 	
			
			pRst->Open(sql.GetBuffer(), strCnn, adOpenStatic, adLockReadOnly, adCmdText);
			
			if (!pRst->EndOfFile)
			{
				pRst->MoveFirst();
				while(!pRst->EndOfFile)
				{
					long iCount = GetFieldLong(pRst, "CountOfLeads");
					bool bVerified = GetFieldBool(pRst, "Ord_Verified");
					bool bOnePass = GetFieldBool(pRst, "Ord_OnePass");
					if(bVerified && !bOnePass) SetCountVerified(iCount);
					if(!bVerified && !bOnePass) SetCountInternet(iCount);
					if(!bVerified && bOnePass) SetCountOnePass(iCount);

					pRst->MoveNext();
				}				
							
			}
			else
			{
				iInternet = 0;
				iOnePass = 0;
				iVerified = 0;
			}
		}
		CoUninitialize(); 	  
	}
	catch(_com_error & ce)
	{
		CString str;		
		str = "com Error: CStats::ReadTodaysLeadsSent " + GetCString(ce.Description());
		PrintOut(str);
		CWriteLog::WriteFail(str);		
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: CStats::ReadTodaysLeadsSent";
		PrintOut(err);
		CWriteLog::WriteFail(err);
	}
}

void CStats::SetCountInternet(long iNum)
{	
	try
	{
		iInternet = iNum;
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: CStats::SetCountInternet(long iNum)";
		PrintOut(err);
		CWriteLog::WriteFail(err);
	}
}

void CStats::SetCountOnePass(long iNum)
{
	try
	{
		iOnePass = iNum;
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: CStats::SetCountOnePass(long iNum)";
		PrintOut(err);
		CWriteLog::WriteFail(err);
	}
}

void CStats::SetCountVerified(long iNum)
{
	try
	{
		iVerified = iNum;
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: CStats::SetCountVerified(long iNum)";
		PrintOut(err);
		CWriteLog::WriteFail(err);
	}
}

CString CStats::GetCountTotal(void)
{
	try
	{
		long iTotal = iInternet + iOnePass + iVerified;
		return CStringHelper::GetCString(iTotal);
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: CStats::GetCountInternet";
		PrintOut(err);
		CWriteLog::WriteFail(err);
	}
}

CString CStats::GetSQLLeadOrderedCounts(void)
{
	CString sql;
	try
	{
		

		/*sql = "SELECT t_Orders.Ord_Verified \r\n";
		sql = sql + ", t_Orders.Ord_OnePass \r\n";
		sql = sql + ", Sum(Ord_MakeGood) AS SumOfOrd_MakeGood \r\n";
		sql = sql + "FROM t_Orders \r\n";
		sql = sql + "GROUP BY t_Orders.Ord_Verified \r\n";
		sql = sql + ", t_Orders.Ord_OnePass \r\n";
		sql = sql + ", Year([Ord_Date]) \r\n";
		sql = sql + ", Month([Ord_Date]) \r\n";
		sql = sql + ", Day([Ord_Date]) \r\n";
		sql = sql + "HAVING (((Year([Ord_Date]))=Year(Now())) \r\n";
		sql = sql + "AND ((Month([Ord_Date]))=Month(Now())) \r\n";
		sql = sql + "AND ((Day([Ord_Date]))=Day(Now()))) \r\n";*/

		sql = "SELECT t_Orders.Ord_Verified \r\n";
		sql = sql + ", t_Orders.Ord_OnePass \r\n";
		sql = sql + ", t_Orders.Ord_MakeGood \r\n";
		sql = sql + "FROM t_Orders \r\n";		
		sql = sql + "WHERE (((Year([Ord_Date]))=Year(Now())) \r\n";
		sql = sql + "AND ((Month([Ord_Date]))=Month(Now())) \r\n";
		sql = sql + "AND ((Day([Ord_Date]))=Day(Now()))) \r\n";




	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: CStats::GetSQLLeadOrderedCounts";
		output.Print(err);
		//PrintOut(err);
		CWriteLog::WriteFail(err);
		
	}
	return sql;
}

void CStats::ReadTodaysLeadsOrdered(void)
{
	try
    {
		HRESULT hr = S_OK;
		
        CoInitialize(NULL);
        // Define string variables.
		 
		_bstr_t strCnn("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=z:\\Mortgage.mdb;Jet OLEDB:Database Password=cepher;");

		_RecordsetPtr pRst = NULL;

		// Call Create instance to instantiate the Record set
		hr = pRst.CreateInstance(__uuidof(Recordset));

		if(FAILED(hr))
		{	            
			PrintOut("Failed creating settings record set instance");	            
		}
		else
		{
			
			CString sql;
			sql = GetSQLLeadOrderedCounts(); 	
			
			pRst->Open(sql.GetBuffer(), strCnn, adOpenStatic, adLockReadOnly, adCmdText);
			PrintOut(sql);
			if (!pRst->EndOfFile)
			{
				SetCountVerifiedOrdered(0);
				SetCountInternetOrdered(0);
				SetCountOnePassOrdered(0);

				pRst->MoveFirst();
				while(!pRst->EndOfFile)
				{
					long iMG = GetFieldLong(pRst, "Ord_MakeGood");
					//if(iMG ==0) PrintOut("Zero");
					//PrintOut(CStringHelper::GetCString(iMG));
					bool bVerified = GetFieldBool(pRst, "Ord_Verified");
					bool bOnePass = GetFieldBool(pRst, "Ord_OnePass");
					if(bVerified && !bOnePass) 
					{
						//PrintOut("Verified");
						SetCountVerifiedOrdered(iMG + iVerifiedOrdered);
					}
					if(!bVerified && !bOnePass) 
					{
						//PrintOut("Internet");
						SetCountInternetOrdered(iMG + iInternetOrdered);
					}
					if(!bVerified && bOnePass) 
					{
						//PrintOut("OnePass");
						SetCountOnePassOrdered(iMG + iOnePassOrdered);
					}

					pRst->MoveNext();
				}				
							
			}
			else
			{
				iInternet = -1;
				iOnePass = -1;
				iVerified = -1;
			}
		}
		CoUninitialize(); 	  
	}
	catch(_com_error & ce)
	{
		CString str;		
		str = "com Error: CStats::ReadTodaysLeadsOrdered " + GetCString(ce.Description());
		PrintOut(str);
		CWriteLog::WriteFail(str);		
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: CStats::ReadTodaysLeadsOrdered";
		PrintOut(err);
		CWriteLog::WriteFail(err);
	}
}

void CStats::SetCountVerifiedOrdered(long iNum)
{
	try
	{
		iVerifiedOrdered = iNum;
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: CStats::SetCountVerifiedOrdered(long iNum)";
		PrintOut(err);
		CWriteLog::WriteFail(err);
	}
}

void CStats::SetCountInternetOrdered(long iNum)
{
	try
	{
		iInternetOrdered = iNum;
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: CStats::SetCountInternetOrdered(long iNum)";
		PrintOut(err);
		CWriteLog::WriteFail(err);
	}
}

void CStats::SetCountOnePassOrdered(long iNum)
{
	try
	{
		iOnePassOrdered = iNum;
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: CStats::SetCountOnePassOrdered(long iNum)";
		PrintOut(err);
		CWriteLog::WriteFail(err);
	}
}

CString CStats::GetCountTotalOrdered(void)
{
	try
	{
		long iTotal = iInternetOrdered + iOnePassOrdered + iVerifiedOrdered;
		return CStringHelper::GetCString(iTotal);
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: CStats::GetCountTotalOrdered";
		PrintOut(err);
		CWriteLog::WriteFail(err);
	}
}

CString CStats::GetCountVerifiedOrdered(void)
{
	try
	{
		return CStringHelper::GetCString(iVerifiedOrdered);
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: CStats::GetCountVerifiedOrdered";
		PrintOut(err);
		CWriteLog::WriteFail(err);
	}
}

CString CStats::GetCountInternetOrdered(void)
{
	try
	{
		return CStringHelper::GetCString(iInternetOrdered);
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: CStats::GetCountInternetOrdered";
		PrintOut(err);
		CWriteLog::WriteFail(err);
	}
}

CString CStats::GetCountOnePassOrdered(void)
{
	try
	{
		return CStringHelper::GetCString(iOnePassOrdered);
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: CStats::GetCountOnePassOrdered";
		PrintOut(err);
		CWriteLog::WriteFail(err);
	}
}

CString CStats::GetYieldVerified(void)
{
	CString str;
	try
	{
		double fYield = (double)iVerified / (double)iVerifiedOrdered;
		fYield = fYield * 100;

		str = CStringHelper::GetCString(fYield, 4);
		if(str.Right(1) == ".")
		{
			str = str.Left(str.GetLength()-1);
		}
		str = str + "%";
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: GetYieldVerified";
		PrintOut(err);
		CWriteLog::WriteFail(err);
	}
	return str;
}

CString CStats::GetYieldInternet(void)
{
	CString str;
	try
	{
		double fYield = (double)iInternet / (double)iInternetOrdered;
		fYield = fYield * 100;

		str = CStringHelper::GetCString(fYield, 4);
		if(str.Right(1) == ".")
		{
			str = str.Left(str.GetLength()-1);
		}
		str = str + "%";
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: GetYieldInternet";
		PrintOut(err);
		CWriteLog::WriteFail(err);
	}
	return str;
}

CString CStats::GetYieldOnePass(void)
{
	CString str;
	try
	{
		double fYield = (double)iOnePass / (double)iOnePassOrdered;
		fYield = fYield * 100;

		str = CStringHelper::GetCString(fYield, 4);
		if(str.Right(1) == ".")
		{
			str = str.Left(str.GetLength()-1);
		}
		str = str + "%";
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: GetYieldOnePass";
		PrintOut(err);
		CWriteLog::WriteFail(err);
	}
	return str;
}



CString CStats::GetYieldTotal(void)
{
	CString str;
	try
	{
		double fTotalSent = iInternet + iOnePass + iVerified;
		double fTotalOrdered = iInternetOrdered + iOnePassOrdered + iVerifiedOrdered;
		double fYield = fTotalSent / fTotalOrdered;
		fYield = fYield * 100;

		str = CStringHelper::GetCString(fYield, 4);
		if(str.Right(1) == ".")
		{
			str = str.Left(str.GetLength()-1);
		}
		str = str + "%";
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: GetYieldTotal";
		PrintOut(err);
		CWriteLog::WriteFail(err);
	}
	return str;
}

CString CStats::GetSQLCampaigns(void)
{
	try
	{
		
		CString sql;
		return sql;
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: CStats::GetSQLCampaigns";
		PrintOut(err);
		CWriteLog::WriteFail(err);
		return CString();
	}
}

CString CStats::GetSQLZerosSent(void)
{
	//get the SQL that count leads by groups
	
	CString sql;
	try
	{
		sql = "SELECT CampaignSent.campaignID, CampaignOrdered.Ordered, CampaignSent.Sent";
		sql = sql + " FROM"; 
		sql = sql + " (";
		sql = sql + " SELECT campaign.campaignID, Count(t_Processed_Orders.PO_Lead_ID) AS Sent";
		sql = sql + " FROM (campaign INNER JOIN t_Orders ON campaign.campaignID = t_Orders.campaignID) LEFT JOIN t_Processed_Orders ON t_Orders.Ord_Id = t_Processed_Orders.PO_Ord_ID";
		sql = sql + " GROUP BY campaign.campaignID, t_Orders.Ord_Complete, campaign.verified, Year([Ord_Sent_On]), Month([Ord_Sent_On]), Day([Ord_Sent_On])";
		sql = sql + " HAVING (((t_Orders.Ord_Complete)=True) AND ((campaign.verified)=True) AND ((Year([Ord_Sent_On]))=Year(Now())) AND ((Month([Ord_Sent_On]))=Month(Now())) AND ((Day([Ord_Sent_On]))=Day(Now())))";

		sql = sql + " ) AS";
		sql = sql + " CampaignSent";
		sql = sql + " INNER JOIN ";

		sql = sql + " (";
		sql = sql + " SELECT campaign.campaignID, Sum(t_Orders.Ord_MakeGood) AS Ordered";
		sql = sql + " FROM campaign INNER JOIN t_Orders ON campaign.campaignID = t_Orders.campaignID";
		sql = sql + " GROUP BY campaign.campaignID, t_Orders.Ord_Complete, campaign.verified, Year([Ord_Sent_On]), Month([Ord_Sent_On]), Day([Ord_Sent_On])";
		sql = sql + " HAVING (((t_Orders.Ord_Complete)=True) AND ((campaign.verified)=True) AND ((Year([Ord_Sent_On]))=Year(Now())) AND ((Month([Ord_Sent_On]))=Month(Now())) AND ((Day([Ord_Sent_On]))=Day(Now())))";

		sql = sql + " ) AS";

		sql = sql + " CampaignOrdered";

		sql = sql + " ON CampaignSent.campaignID = CampaignOrdered.campaignID";

	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: CStats::GetSQLZerosSent";
		output.Print(err);
		//PrintOut(err);
		CWriteLog::WriteFail(err);
		
	}
	return sql;
}

void CStats::ReadTodaysZerosSent(void)
{
	try
    {
		HRESULT hr = S_OK;
		
        CoInitialize(NULL);
        // Define string variables.
		 
		_bstr_t strCnn("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=z:\\Mortgage.mdb;Jet OLEDB:Database Password=cepher;");

		_RecordsetPtr pRst = NULL;

		// Call Create instance to instantiate the Record set
		hr = pRst.CreateInstance(__uuidof(Recordset));

		if(FAILED(hr))
		{	            
			PrintOut("Failed creating settings record set instance");	            
		}
		else
		{
			
			CString sql;
			sql = GetSQLZerosSent(); 	
			
			pRst->Open(sql.GetBuffer(), strCnn, adOpenStatic, adLockReadOnly, adCmdText);
			 
			if (!pRst->EndOfFile)
			{
				iSentAll = 0;
				iSentPartial = 0;
				iSentZero = 0;

				iTotalCampaigns = 0;				
				fMean = 0;

				pRst->MoveFirst();

				double fYieldTotal = 0;

				while(!pRst->EndOfFile)
				{
					double iOrdered = GetFieldDouble(pRst, "Ordered");
					long iSent = GetFieldLong(pRst, "Sent");
					double fYield = 0;
					if(iOrdered > 0)
					{
						fYield = (double)iSent / iOrdered;
					}
					//PrintOut("Ordered:" + CStringHelper::GetCString(iOrdered));
					//PrintOut("Sent: " + CStringHelper::GetCString(iSent));
					//PrintOut("Yield: " + CStringHelper::GetCString(fYield));
					
					fYieldTotal = fYieldTotal + fYield;
					//PrintOut("fYieldTotal: " + CStringHelper::GetCString(fYieldTotal));


					
					if (iOrdered == iSent)
					{
						//they got all the leads the ordered
						iSentAll++;
					}
					else if(iSent == 0)  
					{
						//They didn't get any leads
						iSentZero++;
					}
					else
					{
						//their order was partially filled
						iSentPartial++;
					}
					iTotalCampaigns++; // count the total number of campaigns sent orders today
					pRst->MoveNext();
				}
				if(iTotalCampaigns > 0)
				{
					fMean = fYieldTotal / (double)iTotalCampaigns;
				}
			}
			else
			{
				iSentAll = 0;
				iSentPartial = 0;
				iSentZero = 0;
				iTotalCampaigns = 0;
				fMean = 0;

			}
		}
		CoUninitialize(); 	  
	}
	catch(_com_error & ce)
	{
		CString str;		
		str = "com Error: CStats::ReadTodaysZerosSent " + GetCString(ce.Description());
		PrintOut(str);
		CWriteLog::WriteFail(str);		
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: CStats::ReadTodaysZerosSent";
		PrintOut(err);
		CWriteLog::WriteFail(err);
	}
}

CString CStats::GetTotalCampaigns(void)
{
	try
	{
		return CStringHelper::GetCString((long)iTotalCampaigns);
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: CStats::GetTotalCampaigns";
		PrintOut(err);
		CWriteLog::WriteFail(err);
		return CString();
	}
}

CString CStats::GetSentAll(void)
{
	try
	{
		return CStringHelper::GetCString((long)iSentAll);
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: CStats::GetSentAll";
		PrintOut(err);
		CWriteLog::WriteFail(err);
		return CString();
	}
}

CString CStats::GetSentZero(void)
{
	try
	{
		return CStringHelper::GetCString((long)iSentZero);
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: CStats::GetSentZero";
		PrintOut(err);
		CWriteLog::WriteFail(err);
		return CString();
	}
}

CString CStats::GetSentPartials(void)
{
	try
	{
		return CStringHelper::GetCString((long)iSentPartial);
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: CStats::GetSentPartials";
		PrintOut(err);
		CWriteLog::WriteFail(err);
		return CString();
	}
}

CString CStats::GetPercentZero(void)
{
	CString str;
	try
	{
		if(iTotalCampaigns == 0) return "0.0%";//prevent divide by 0
		double fPercent = (double)iSentZero / (double)iTotalCampaigns;
		fPercent = fPercent * 100;

		str = CStringHelper::GetCString(fPercent, 4);
		if(str.Right(1) == ".")
		{
			str = str.Left(str.GetLength()-1);
		}
		str = str + "%";
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: CStats::GetPercentZero";
		PrintOut(err);
		CWriteLog::WriteFail(err);
	}
	return str;
}

CString CStats::GetPercentPartials(void)
{
	CString str;
	try
	{
		if(iTotalCampaigns == 0) return "0.0%";//prevent divide by 0
		double fPercent = (double)iSentPartial / (double)iTotalCampaigns;
		fPercent = fPercent * 100;

		str = CStringHelper::GetCString(fPercent, 4);
		if(str.Right(1) == ".")
		{
			str = str.Left(str.GetLength()-1);
		}
		str = str + "%";
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: CStats::GetPercentPartials";
		PrintOut(err);
		CWriteLog::WriteFail(err);
	}
	return str;
}

CString CStats::GetPercentAll(void)
{
	CString str;
	try
	{
		if(iTotalCampaigns == 0) return "0.0%";//prevent divide by 0
		double fPercent = (double)iSentAll / (double)iTotalCampaigns;
		fPercent = fPercent * 100;

		str = CStringHelper::GetCString(fPercent, 4);
		if(str.Right(1) == ".")
		{
			str = str.Left(str.GetLength()-1);
		}
		str = str + "%";
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: CStats::GetPercentAll";
		PrintOut(err);
		CWriteLog::WriteFail(err);
	}
	return str;
}

CString CStats::GetMean(void)
{
	CString str;
	try
	{
		double fPercent = fMean * 100;
		
		str = CStringHelper::GetCString(fPercent, 4);
		if(str.Right(1) == ".")
		{
			str = str.Left(str.GetLength()-1);
		}
		str = str + "%";
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: CStats::GetMean";
		PrintOut(err);
		CWriteLog::WriteFail(err);
	}
	return str;
}
