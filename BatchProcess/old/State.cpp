#include "StdAfx.h"
#include ".\state.h"

State::State(void)
: st_Name(_T(""))
, st_Abr(_T(""))
, m_output(NULL)
{
}

State::~State(void)
{
}

void State::SetSt_Name(CString name)
{
	st_Name = name;
}

void State::SetSt_Abr(CString abr)
{
	st_Abr = abr;
}

CString State::GetSt_Name(void)
{
	return st_Name;
}

CString State::GetSt_Abr(void)
{
	return st_Abr;
}

void State::AddLead(CLead aLead)
{
	Output("got a add lead to state");
	if(aLead.GetLead_State().CompareNoCase(st_Abr) == 0)
	{
		//leads.add(aLead);
	}
}

LeadList State::GetLeads(void)
{
	return leads;
}

void State::ReadState(_RecordsetPtr rs)
{
	SetSt_Abr(GetFieldString(rs, "st_Abrv"));
	SetSt_Name(GetFieldString(rs, "st_Name"));
}

void State::PrintOut(CHListBox* m_output)
{
	CString str;
	str = GetSt_Abr();
	str = str + " - ";
	str = str + GetSt_Name();	

	m_output->AddString(str);	
	m_output->SetCurSel(m_output->GetCount()-1);		
	m_output->RedrawWindow();
}

void State::SetOutput(CHListBox* pointer)
{
	m_output = pointer;
}

void State::Output(CString str)
{
	if(m_output != NULL)
	{
		m_output->AddString(str);
		m_output->SetCurSel(m_output->GetCount()-1);		
		m_output->RedrawWindow();
	}
}
