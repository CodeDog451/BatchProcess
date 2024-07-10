#include "StdAfx.h"
#include ".\manualprocessthread.h"

CManualProcessThread::CManualProcessThread(void)
{
}

CManualProcessThread::~CManualProcessThread(void)
{
}

void CManualProcessThread::SetCategory(int iLeadType, int iOrderType)
{
	iCategory = iLeadType;// verifed, internet or one pass
	iGroup = iOrderType;// normal, priority, sent zero, new customer
	//iState = iStateType;
}

void CManualProcessThread::ProcessBatch(void)
{
	
	ProcessOrders(iCategory, iGroup);
	PrintOut("Done", 0);
	if(mainWin)
	{		
		SendMessage(mainWin, PO_DONE, 0, 0);		
	}
}
