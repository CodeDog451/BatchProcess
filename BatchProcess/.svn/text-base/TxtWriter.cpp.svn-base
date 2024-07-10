#include "StdAfx.h"
#include ".\txtwriter.h"

CTxtWriter::CTxtWriter(CString target, HWND m_hwnd)
	: CWriter(target, m_hwnd)	
{
	
}

CTxtWriter::~CTxtWriter(void)
{
}

void CTxtWriter::Write(COrder aOrder, CLocation myloc)
{
	CString target;
	
	target = desDir;
	target = target + "\\" + aOrder.GetOrderID() + ".txt";
	
	CFileException e;
	CFile *m_target = new CFile();
	

	if(!m_target->Open(target.GetBuffer(target.GetLength()), CFile::modeCreate | CFile::modeReadWrite, &e))
	{
		PrintOut("Could not open the target file for write at path: " + target);
		
	}
	else
	{
		PrintOut("Opened the target file for write at path: " + target);		
		
		CString body = myloc.GetZeroBody();
		body.Replace("<agentname>", aOrder.GetAgent());	
		body.Replace("<agentemail>", aOrder.GetAgentEmail());
		body.Replace("<orderid>", aOrder.GetOrderID());
		body.Replace("<company>", aOrder.GetCompanyName());
		body.Replace("<contact>", aOrder.GetContact());
		body.Replace("<state>", aOrder.GetStates());
		if(aOrder.GetFilterCount() > 0)
		{
			body.Replace("<filterstatement>", myloc.GetFilterStatement());
		}
		else
		{
			body.Replace("<filterstatement>", "");
		}
		
		WriteString(m_target, body);		
		
		m_target->Close();
	}
}


