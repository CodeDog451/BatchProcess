#include "StdAfx.h"
#include ".\faxjob.h"

CFaxJob::CFaxJob(void)
: sFaxnum(_T(""))
, sAgentname(_T(""))
, sAgentEmail(_T(""))
, sCompany(_T(""))
, iDiv(0)
, sFilename(_T(""))
, sContact(_T(""))
, iOrderID(0)
, iCampaignID(0)
, sStates(_T(""))
, iSendCoverpage(0)
{
}

CFaxJob::~CFaxJob(void)
{
}

void CFaxJob::SetFaxnum(CString sNum)
{
	//set the fax number
	try
	{
		CString str = sNum;
		if(sNum.GetLength() > 0)
		{	
			str.Replace("(", "");
			str.Replace(")", "");
			str.Replace("-", "");
			str.Replace(" ", "");
			CString sFirst = str.Left(1);
			if(!(sFirst.Compare("1")==0))
			{
				str = "1" + str;
			}
		}

		sFaxnum = str;	

	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: CFaxJob::SetFaxnum(CString sNum)";
		PrintOut(err);
		CWriteLog::WriteFail(err);
	}
}

CString CFaxJob::GetFaxnum(void)
{
	// get the fax number
	try
	{
		return sFaxnum;
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: CFaxJob::GetFaxnum";
		PrintOut(err);
		CWriteLog::WriteFail(err);
		return CString();
	}
}

void CFaxJob::SetAgentname(CString sName)
{
	//set the agent name
	try
	{
		sAgentname = sName;
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: CFaxJob::SetAgentname(CString sName)";
		PrintOut(err);
		CWriteLog::WriteFail(err);
	}
}

CString CFaxJob::GetAgentname(void)
{
	// get the agent name
	try
	{
		return sAgentname;
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: CFaxJob::GetAgentname";
		PrintOut(err);
		CWriteLog::WriteFail(err);
		return CString();
	}
}

void CFaxJob::SetAgentEmail(CString sEmail)
{
	// set the agent email
	try
	{
		sAgentEmail = sEmail;
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: CFaxJob::SetAgentEmail(CString sEmail)";
		PrintOut(err);
		CWriteLog::WriteFail(err);
	}
}

CString CFaxJob::GetAgentEmail(void)
{
	//get the agent email
	try
	{
		return sAgentEmail;
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: CFaxJob::GetAgentEmail";
		PrintOut(err);
		CWriteLog::WriteFail(err);
	}
}

void CFaxJob::SetCompany(CString sName)
{
	// set the company name
	try
	{
		sCompany = sName;
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: CFaxJob::SetCompany(CString sName)";
		PrintOut(err);
		CWriteLog::WriteFail(err);
	}
}

CString CFaxJob::GetCompany(void)
{
	
	try
	{
		return sCompany;
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: CFaxJob::GetCompany";
		PrintOut(err);
		CWriteLog::WriteFail(err);
	}
}

int CFaxJob::GetDiv(void)
{
	
	try
	{
		return iDiv;
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: CFaxJob::GetDiv";
		PrintOut(err);
		CWriteLog::WriteFail(err);
	}
}

void CFaxJob::SetDiv(int iNum)
{
	try
	{
		iDiv = iNum;
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: CFaxJob::SetDiv(int iNum)";
		PrintOut(err);
		CWriteLog::WriteFail(err);
	}
}

void CFaxJob::SetFilename(CString sPath)
{
	
	try
	{
		sFilename = sPath;
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: CFaxJob::SetFilename(CString sPath)";
		PrintOut(err);
		CWriteLog::WriteFail(err);
	}
}

CString CFaxJob::GetFilename(void)
{
	
	try
	{
		return sFilename;
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: CFaxJob::GetFilename";
		PrintOut(err);
		CWriteLog::WriteFail(err);
	}
}

void CFaxJob::SetContact(CString sName)
{
	
	try
	{
		sContact = sName;
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: CFaxJob::SetContact(CString sName)";
		PrintOut(err);
		CWriteLog::WriteFail(err);
	}
}

CString CFaxJob::GetContact(void)
{
	
	try
	{
		return sContact;
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: CFaxJob::GetContact";
		PrintOut(err);
		CWriteLog::WriteFail(err);
		return CString();
	}
}

void CFaxJob::SetOrderID(long iID)
{
	
	try
	{
		iOrderID = iID;
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: CFaxJob::SetOrderID(long iID)";
		PrintOut(err);
		CWriteLog::WriteFail(err);
	}
}

CString CFaxJob::GetCampaignID(void)
{
	try
	{
		CString sNum;
		char buffer[200];
		sprintf( buffer, "%d", iCampaignID );
		sNum = buffer;
		return sNum;
	

	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: CFaxJob::GetCampaignID";
		PrintOut(err);
		CWriteLog::WriteFail(err);
		return CString();
	}
}

void CFaxJob::SetCampaignID(long iID)
{
	
	try
	{
		iCampaignID = iID;
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: CFaxJob::SetCampaignID(long iID)";
		PrintOut(err);
		CWriteLog::WriteFail(err);
	}
}

CString CFaxJob::GetOrderID(void)
{
	try
	{
		CString sNum;
		char buffer[200];
		sprintf( buffer, "%d", iOrderID );
		sNum = buffer;
		return sNum;
	

	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: CFaxJob::GetOrderID";
		PrintOut(err);
		CWriteLog::WriteFail(err);
		return CString();
	}
}

CString CFaxJob::GetIDMessage(void)
{
	try
	{
		CString str;
		str = "[Company: " + GetCompany() + "] ";
		str = str + "[Campaign ID: " + GetCampaignID() + "] ";
		str = str + "[Order: " + GetOrderID() + "] ";
		str = str + "[Fax Number:" + GetFaxnum() + "] ";
		return str;
	

	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: CFaxJob::GetIDMessage";
		PrintOut(err);
		CWriteLog::WriteFail(err);
		return CString();
	}
}

void CFaxJob::ClearAll(void)
{
	try
	{
		this->iCampaignID=0;
		this->iDiv=0;
		this->iOrderID=0;
		this->sAgentEmail="";
		this->sAgentname="";
		this->sCompany="";
		this->sContact="";
		this->sFaxnum="";
		this->sFilename="";
		this->sStates="";
		this->iSendCoverpage=1;//default is to send the cover page
	

	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: CFaxJob::ClearAll";
		PrintOut(err);
		CWriteLog::WriteFail(err);
	}
	
}

void CFaxJob::SendTo(HWND mainWin)
{
	try
	{
		if(mainWin)
		{
			SendMessage(mainWin, OB_FAXJOB_START, 0, 0);

			SendMessage(mainWin, OB_FAXJOB_CAMID, iCampaignID, 0);
			SendMessage(mainWin, OB_FAXJOB_DIV, iDiv, 0);
			SendMessage(mainWin, OB_FAXJOB_ORDERID, iOrderID, 0);
			SendMessage(mainWin, OB_FAXJOB_SENDCOVERPAGE, iSendCoverpage, 0);

			SendString(mainWin, OB_FAXJOB_AGENTEMAIL_CHAR, sAgentEmail);
			SendString(mainWin, OB_FAXJOB_AGENTNAME_CHAR, this->sAgentname);
			SendString(mainWin, OB_FAXJOB_COMPANY_CHAR, sCompany);
			SendString(mainWin, OB_FAXJOB_CONTACT_CHAR, sContact);
			SendString(mainWin, OB_FAXJOB_FAXNUM_CHAR, sFaxnum);
			SendString(mainWin, OB_FAXJOB_FILENAME_CHAR, sFilename);	
			SendString(mainWin, OB_FAXJOB_STATES_CHAR, sStates);
			

			SendMessage(mainWin, OB_FAXJOB_END, 0, 0);
		}
	

	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: CFaxJob::SendTo(HWND mainWin)";
		PrintOut(err);
		CWriteLog::WriteFail(err);
	}

}

void CFaxJob::SendString(HWND mainWin, int resid, CString str)
{
	try
	{
		int iLen = str.GetLength();
		for(int x = 0; x < iLen; x++)
		{
			int iLetter = (int) str.GetAt(x);
			SendMessage(mainWin, resid, iLetter, 0);
		}
	

	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: CFaxJob::SendString(HWND mainWin, int resid, CString str)";
		PrintOut(err);
		CWriteLog::WriteFail(err);
	}
}

void CFaxJob::SetStates(CString str)
{
	
	try
	{
		sStates = str;
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: CFaxJob::SetStates(CString str)";
		PrintOut(err);
		CWriteLog::WriteFail(err);
	}
}

CString CFaxJob::GetStates(void)
{
	
	try
	{
		return sStates;
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: CFaxJob::GetStates";
		PrintOut(err);
		CWriteLog::WriteFail(err);
	}
}

void CFaxJob::SetSendCoverpage(int iNum)
{
	
	try
	{
		iSendCoverpage = iNum;
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: CFaxJob::SetSendCoverpage(int iNum)";
		PrintOut(err);
		CWriteLog::WriteFail(err);
	}
}

int CFaxJob::GetSendCoverPage(void)
{
	
	try
	{
		return iSendCoverpage;
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: CFaxJob::GetSendCoverPage";
		PrintOut(err);
		CWriteLog::WriteFail(err);
		return -1;
	}
}

void CFaxJob::PrintOut(CString str)
{
	CWriteLog::Write(str);
	//AfxMessageBox(str);
}
