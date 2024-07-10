// ESMTP.cpp: implementation of the ESMTP class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
//#include "Worksheet.h"

#include "SMTP.h"
#include "ESMTP.h"
#include ".\esmtp.h"
#include "resource.h"
#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[]=__FILE__;
#define new DEBUG_NEW
#endif

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

ESMTP::ESMTP()
: iDiv(0)
{
	m_nPort = 25;
	m_Authenticate = (CPJNSMTPConnection::LoginMethod) 0;
	m_bMime = false;
	m_bHTML = false;
	m_sEncodingFriendly = _T("Western European (ISO)");
	m_sEncodingCharset = _T("iso-8859-1");
	m_Priority = (CSMTPMessage::PRIORITY) 0;

}

ESMTP::~ESMTP()
{

}
void ESMTP::SetAddress(CString address)
{
	
	try
	{
		m_sAddress = address;
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: ESMTP::SetAddress(CString address)";
		PrintOut(err);
		CWriteLog::WriteFail(err);
	}
}

void ESMTP::SetName(CString name)
{
	
	try
	{
		m_sName = name;
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: ESMTP::SetName(CString name)";
		PrintOut(err);
		CWriteLog::WriteFail(err);
	}

}

void ESMTP::SetAuthenticate(int authlevel)
{
	
	try
	{
		m_Authenticate = (CPJNSMTPConnection::LoginMethod) authlevel;
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: ESMTP::SetAuthenticate(int authlevel)";
		PrintOut(err);
		CWriteLog::WriteFail(err);
	}
}

void ESMTP::SetUserName(CString username)
{
	
	try
	{
		m_sUsername = username;
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: ESMTP::SetUserName(CString username)";
		PrintOut(err);
		CWriteLog::WriteFail(err);
	}

}

void ESMTP::SetPassword(CString password)
{
	
	try
	{
		m_sPassword = password;
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: ESMTP::SetPassword(CString password)";
		PrintOut(err);
		CWriteLog::WriteFail(err);
	}

}

bool ESMTP::Send()
{
	
	bool success = true;
	try
	{
		CSMTPMessage* pMessage = CreateMessage();
		
		//AfxMessageBox(m_sHost + " : " + m_sUsername + " : " + m_sPassword + " : " + m_sBoundIP, MB_ICONSTOP);
		//BOOL bConnect = connection.Connect(m_sHost, m_Authenticate, m_sUsername, m_sPassword, m_nPort, m_sBoundIP);
		BOOL bConnect = connection.Connect(m_sHost, m_Authenticate, m_sUsername, m_sPassword, m_nPort, m_sBoundIP);
		if (!bConnect) 
		{
			CString sMsg;
			sMsg.Format(_T("Couldn't connect to server!, Error:%d"), GetLastError());
			PrintOut(sMsg);
			PrintFail(sMsg);
			//AfxMessageBox(sMsg, MB_ICONSTOP);
			success = false;
		}
		else 
		{
			PrintOut("Connected Ok");
			success = true;
			//Send the message
			if (!connection.SendMessage(*pMessage)) 
			{
				CString sMsg;
				sMsg.Format(_T("Couldn't send message!\nResponse:%s"), connection.GetLastCommandResponse());
				PrintOut(sMsg);
				PrintFail(m_sSubject);
				PrintFail(m_sTo);
				PrintFail(sMsg);
				//AfxMessageBox(sMsg, MB_ICONSTOP);
				success = false;
			}
			else
			{
				PrintOut("Send success");
			}
		}
	      

		//Tidy up the heap memory
		delete pMessage;
	

	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: ESMTP::Send()";
		PrintOut(err);
		CWriteLog::WriteFail(err);
	}
	return success;	  

}

CSMTPMessage* ESMTP::CreateMessage()
{
	try
	{
				//Create the message
		CSMTPMessage* pMessage = new CSMTPMessage;
		CSMTPBodyPart attachment;

		//Set the mime flag
		pMessage->SetMime(m_bMime);

		//Set the charset of the message and all attachments
		pMessage->SetCharset(m_sEncodingCharset);
		attachment.SetCharset(m_sEncodingCharset);

		//Set the message priority
		pMessage->m_Priority = m_Priority;

		//Setup the all the recipient types for this message
			pMessage->AddMultipleRecipients(m_sTo, CSMTPMessage::TO);
			if (!m_sCC.IsEmpty()) 
				pMessage->AddMultipleRecipients(m_sCC, CSMTPMessage::CC);
			if (!m_sBCC.IsEmpty()) 
				pMessage->AddMultipleRecipients(m_sBCC, CSMTPMessage::BCC);
			if (!m_sSubject.IsEmpty()) 
				pMessage->m_sSubject = m_sSubject;
			if (!m_sBody.IsEmpty())
		{
			if (m_bHTML)
			pMessage->AddHTMLBody(m_sBody, _T(""));
			else
					pMessage->AddTextBody(m_sBody);
		}

		//Add the attachment(s) if necessary
			//if (!m_sFile.IsEmpty()) 
				//pMessage->AddMultipleAttachments(m_sFile);	

			if(myAttach.GetCount()>0)
			{
				for(int x = 0; x < myAttach.GetCount(); x++)
				{
					CString filename = myAttach.GetAt(x);
					pMessage->AddMultipleAttachments(filename);	
				}
			}

		//Setup the from address
			if (m_sName.IsEmpty()) 
		{
				pMessage->m_From = m_sAddress;
				//pMessage->m_ReplyTo = m_sAddress; uncomment this if you want to send a Reply-To header
			}
			else 
		{
			CSMTPAddress address(m_sName, m_sAddress);
				pMessage->m_From = address;
				//pMessage->m_ReplyTo = address; //uncomment this if you want to send a Reply-To header
			}

		pMessage->m_sXMailer = _T(""); //comment this line out if you want to send a X-Mailer header

		#ifdef _DEBUG
		//Add one custom header (for test purpose)
		pMessage->AddCustomHeader(_T("X-Program: CSTMPMessageTester"));

		//Try out the copy constructor and operator= methods
		CSMTPMessage copyOfMessage(*pMessage);
		#endif

		return pMessage;
	

	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: ESMTP::CreateMessage";
		PrintOut(err);
		CWriteLog::WriteFail(err);
		return new CSMTPMessage();
	}

}

void ESMTP::SetBody(CString body)
{
	
	try
	{
		m_sBody = body;
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: ESMTP::SetBody(CString body)";
		PrintOut(err);
		CWriteLog::WriteFail(err);
	}

}

void ESMTP::SetParentWindow(HWND m_hwnd)
{
	
	try
	{
		mainWin = m_hwnd;
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: ESMTP::SetParentWindow(HWND m_hwnd)";
		PrintOut(err);
		CWriteLog::WriteFail(err);
	}
}

void ESMTP::PrintOut(CString str)
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
		err = "Error: ESMTP::PrintOut(CString str)";
		PrintOut(err);
		CWriteLog::WriteFail(err);
	}
}

void ESMTP::AddFilename(CString sFilename)
{
	
	try
	{
		myAttach.Add(sFilename);
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: ESMTP::AddFilename(CString sFilename)";
		PrintOut(err);
		CWriteLog::WriteFail(err);
	}
}

void ESMTP::ClearAttach(void)
{
	
	try
	{
		myAttach.RemoveAll();
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: ESMTP::ClearAttach";
		PrintOut(err);
		CWriteLog::WriteFail(err);
	}
}

void ESMTP::PrintFail(CString str)
{
	WriteFailLog(str);
	/*if(mainWin)
	{
		SendMessage(mainWin, EMAIL_FAIL_START, 0, iDiv);
		int iLen = str.GetLength();
		for(int x = 0; x < iLen; x++)
		{
			int iLetter = (int) str.GetAt(x);
			SendMessage(mainWin, EMAIL_FAIL_CHAR, iLetter, iDiv);
		}
		SendMessage(mainWin, EMAIL_FAIL_END, 0, iDiv);
	}*/
}

void ESMTP::SetDiv(int iLoc)
{
	
	try
	{
		iDiv = iLoc;
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: ESMTP::SetDiv(int iLoc)";
		PrintOut(err);
		CWriteLog::WriteFail(err);
	}
}

void ESMTP::WriteFailLog(CString str)
{
	CWriteLog::WriteFail(str);
	/*
	///////////////////////////////////////////////////////////////////
	
	
	try
    {
		HRESULT hr;
		hr = S_OK;	
		 CoInitialize(NULL);

		_RecordsetPtr pRst;		
		_bstr_t strCnn;
		pRst = NULL;

		strCnn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=BatcherLog.mdb;Jet OLEDB:Database Password=cepher;";
		 // Call Create instance to instantiate the Record set
		CString sql = "SELECT * FROM FailLog";
		hr = pRst.CreateInstance(__uuidof(Recordset));
		if(FAILED(hr))
		{            
			PrintOut("Failed creating batcher Fail log record set instance");
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
		str = "Error WriteFailLog: " + GetCString(ce.Description());
		PrintOut(str);	
		
	}
	///////////////////////////////////////////////////////////////////
	*/
}

CString ESMTP::GetCString(_bstr_t bstrStart)
{
	return CStringHelper::GetCString(bstrStart);
	/*TCHAR szFinal[255];
	_stprintf(szFinal, _T("%s"), (LPCTSTR)bstrStart);
	CString str;
	str = szFinal;
	return str;*/
}

// set who the email is going to
void ESMTP::SetTo(CString str)
{
	
	try
	{
		m_sTo = str;
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: ESMTP::SetTo(CString str)";
		PrintOut(err);
		CWriteLog::WriteFail(err);
	}
}

// set the blind carbon copies
void ESMTP::SetBCC(CString str)
{
	try
	{
		m_sBCC = str;
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: ESMTP::SetBCC(CString str)";
		PrintOut(err);
		CWriteLog::WriteFail(err);
	}
}

// set the subject line of the email
void ESMTP::SetSubject(CString str)
{
	try
	{
		m_sSubject = str;
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: ESMTP::SetSubject(CString str)";
		PrintOut(err);
		CWriteLog::WriteFail(err);
	}
}

void ESMTP::SetHost(CString str)
{
	try
	{
		m_sHost = str;
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: ESMTP::SetHost(CString str)";
		PrintOut(err);
		CWriteLog::WriteFail(err);
	}
}

// get the email addres the email is being sent too
CString ESMTP::GetTo(void)
{
	
	try
	{
		return m_sTo;
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: ESMTP::GetTo";
		PrintOut(err);
		CWriteLog::WriteFail(err);
		return CString();
	}
}

// get the BCC
CString ESMTP::GetBCC(void)
{
	
	try
	{
		return m_sBCC;
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: ESMTP::GetBCC";
		PrintOut(err);
		CWriteLog::WriteFail(err);
	}
}
