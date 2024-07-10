#include "StdAfx.h"
#include ".\leaddb.h"
#include "resource.h"


CLeadDB::CLeadDB(void)
: stop(false)
{
}

CLeadDB::~CLeadDB(void)
{
}

void CLeadDB::Add(CLead aLead)
{
	myLeadMap.SetAt(aLead.GetLead_Id(), aLead);
}

void CLeadDB::SetParentWindow(HWND m_hwnd)
{
	mainWin = m_hwnd;
}

bool CLeadDB::GetLead(long iID, CLead &aLead)
{
	//CLead aLead;
	
	return myLeadMap.Lookup(iID, aLead);
}

void CLeadDB::PrintOut(CString str)
{
	//myLog.WriteLog(str, 0);
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

void CLeadDB::PrintLeads(void)
{
	POSITION pos = myLeadMap.GetStartPosition();
	long    nKey;
	CLead aLead;
	while ((pos != NULL)&& (!stop))
	{	   
		myLeadMap.GetNextAssoc( pos, nKey, aLead );
		PrintOut(aLead.ToString());      
	}

}



void CLeadDB::RemoveAll(void)
{
	myLeadMap.RemoveAll();
}

void CLeadDB::SetStop(bool state)
{
	stop = state;
}

CLead CLeadDB::GetLead(long id)
{
	CLead aLead;	
	myLeadMap.Lookup(id, aLead);

	return aLead;
}

void CLeadDB::Update(CLead aLead)
{
	myLeadMap.SetAt(aLead.GetLead_Id(), aLead);
}
CLeadDB& CLeadDB::operator=(const CLeadDB& aCLeadDB)
{
	this->mainWin = aCLeadDB.mainWin;
	this->stop = aCLeadDB.stop;
	POSITION pos = aCLeadDB.myLeadMap.GetStartPosition();
	long    nKey;
	CLead aLead;
	myLeadMap.RemoveAll();
	while (pos != NULL)
	{	   
		aCLeadDB.myLeadMap.GetNextAssoc( pos, nKey, aLead );
		myLeadMap.SetAt(aLead.GetLead_Id(), aLead);
	}	

	return *this;
}
CLeadDB::CLeadDB(const CLeadDB& aCLeadDB)
{
	
	this->mainWin = aCLeadDB.mainWin;
	this->stop = aCLeadDB.stop;
	POSITION pos = aCLeadDB.myLeadMap.GetStartPosition();
	long    nKey;
	CLead aLead;
	myLeadMap.RemoveAll();
	while (pos != NULL)
	{	   
		aCLeadDB.myLeadMap.GetNextAssoc( pos, nKey, aLead );
		myLeadMap.SetAt(aLead.GetLead_Id(), aLead);
	}
	

}
