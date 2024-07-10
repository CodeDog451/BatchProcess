#include "StdAfx.h"
#include ".\printout.h"

CPrintOut::CPrintOut(void)
{
}

CPrintOut::~CPrintOut(void)
{
}

void CPrintOut::SetParentWindow(HWND m_wnd)
{
	try
	{
		mainWin = m_wnd;
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: CPrintOut::SetParentWindow(HWND m_wnd)";
		AfxMessageBox(err);
		CWriteLog::WriteFail(err);
	}
}

void CPrintOut::Print(CString str)
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
		err = "Error: CPrintOut::Print(CString str)";
		AfxMessageBox(err);
		CWriteLog::WriteFail(err);
	}
}
