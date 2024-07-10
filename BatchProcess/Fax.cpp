#include "StdAfx.h"
#include ".\fax.h"
#include "resource.h"
CFax::CFax(void)
: filepath(_T(""))
, coverpage(_T(""))
, coverpagenote(_T(""))
, faxnumber(_T(""))
, iSendCoverpage(1)
{
}

CFax::~CFax(void)
{
}

bool CFax::Send(void)
{
	bool success = false;

    try
    {

		//HRESULT sent;
        HRESULT hr;
        //
        // Define variables.
        //
		FAXCOMLib::IFaxServerPtr sipFaxServer;
		
		FAXCOMLib::IFaxDocPtr sipFaxDocument;
		long numPorts=-1;
		
		FAXCOMLib::IFaxPortsPtr sipFaxPorts;		
		
        //
        // Initialize the COM library on the current thread.
        //
        hr = CoInitialize(NULL); 
        if (FAILED(hr))
        {
            _com_issue_error(hr);
        }
        //                
        // Create the root objects.
        //
				
		hr = sipFaxServer.CreateInstance("FaxServer.FaxServer");		
        if (FAILED(hr))
        {
            _com_issue_error(hr);
        }
		PrintOut("Made a server instance");
		
        // Connect to the fax server. 
        // "" defaults to the machine on which the program is running.
        //
        hr = sipFaxServer->Connect("");
        if (FAILED(hr))
        {
            _com_issue_error(hr);
        }
		PrintOut("Connected to Fax");
		sipFaxPorts = sipFaxServer -> GetPorts();
		numPorts = sipFaxPorts->GetCount();
		if(numPorts < 1) 
		{
			PrintOut("No fax installed");
		}
		else
		{
			PrintOut("fax ports found");
			PrintOut(filepath);
			sipFaxDocument = sipFaxServer->CreateDocument(filepath.GetBuffer());
			sipFaxDocument->SendCoverpage = iSendCoverpage;
			sipFaxDocument->CoverpageName = _bstr_t(coverpage.GetBuffer());
			//sipFaxDocument->CoverpageName = "z:\\COVERPAGE.COV";
			//sipFaxDocument->CoverpageNote = coverpagenote.GetBuffer();
			sipFaxDocument->CoverpageNote = _bstr_t(coverpagenote.GetBuffer());
			sipFaxDocument->FaxNumber = faxnumber.GetBuffer();
			if(SUCCEEDED(sipFaxDocument->Send())) success = true;
			PrintOut(filepath + " faxed to" + faxnumber);
		}
        CoUninitialize();
    }
    catch (_com_error& e)
    {
		CString emsg;
		CString str = (char*)e.ErrorMessage();
		emsg = "Fax Error: ";
		emsg = emsg + str;
		char  buffer[200];
		sprintf( buffer,     "(Error CFax::Send: %d)", e.Error() );
		emsg = emsg + buffer;
		PrintOut(emsg);
		CWriteLog::WriteFail(emsg);
		
        if (e.ErrorInfo())
        {
           
			PrintOut((char *)e.Description());
			CWriteLog::WriteFail((char *)e.Description());
			
        }
    }
   
	//Sleep(65000); //time out test
    return success;


}

void CFax::SetParentWindow(HWND m_hwnd)
{
	try
	{
		mainWin = m_hwnd;
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: CFax::SetParentWindow(HWND m_hwnd)";
		PrintOut(err);
		CWriteLog::WriteFail(err);
	}
}

void CFax::PrintOut(CString str)
{
	//Print message to main GUI feedback listbox
	try
	{
		if(mainWin)
		{
			SendMessage(mainWin, S_START_FAX_STRING, 0, 0);
			int iLen = str.GetLength();
			for(int x = 0; x < iLen; x++)
			{
				int iLetter = (int) str.GetAt(x);
				SendMessage(mainWin, S_CHAR_FAX, iLetter, 0);
			}

			SendMessage(mainWin, S_END_FAX_STRING, 0, 0);
		}
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: CFax::PrintOut(CString str)";
		PrintOut(err);
		CWriteLog::WriteFail(err);
	}
}

void CFax::SetCoverPage(CString filename)
{
	//set the path to the coverpage file	
	try
	{
		coverpage = filename;
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: CFax::SetCoverPage(CString filename)";
		PrintOut(err);
		CWriteLog::WriteFail(err);
	}
}

void CFax::SetCoverPageNote(CString note)
{
	//set the contents of the coverpage note
	try
	{
		coverpagenote = note;
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: CFax::SetCoverPageNote(CString note)";
		PrintOut(err);
		CWriteLog::WriteFail(err);
	}
}

void CFax::SetFaxNumber(CString phone)
{	
	//set the fax number to call
	try
	{
		faxnumber = phone;
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: CFax::SetFaxNumber(CString phone)";
		PrintOut(err);
		CWriteLog::WriteFail(err);
	}
}

void CFax::SetBody(CString filename)
{
	//set the path to the body file
	try
	{
		filepath = filename;
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: CFax::SetBody(CString filename)";
		PrintOut(err);
		CWriteLog::WriteFail(err);
	}
}

CString CFax::GetFaxNumber(void)
{
	// get the fax number as a string
	try
	{
		return faxnumber;
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: CFax::GetFaxNumber";
		PrintOut(err);
		CWriteLog::WriteFail(err);
		return CString();
	}
}



void CFax::SetSendCoverpage(int iNum)
{
	// set whether to send a coversheet or not
	try
	{
		iSendCoverpage = iNum;
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: CFax::SetSendCoverpage(int iNum)";
		PrintOut(err);
		CWriteLog::WriteFail(err);
	}
}
