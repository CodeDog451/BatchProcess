#include "StdAfx.h"
#include ".\leads.h"

Leads::Leads(void)
: m_output(NULL)
{
}

Leads::~Leads(void)
{
}

void Leads::AddLead(CLead aLead)
{
	//theLeads.add(aLead);
	//long index;
	//index = theLeads.getUpperBound();
	if(aLead.GetLead_Verified())
	{
		Output("got a verifed");
		verifiedLeads.AddLead(aLead);
		
	}
	else
	{
		
		Output("got an internet");
		nonVerifiedLeads.AddLead(aLead);
		
	}
}

LeadList Leads::GetLeads(bool verified)
{
		
	if(verified)
	{
		return verifiedLeads.GetLeads();
	}
	else
	{
		return nonVerifiedLeads.GetLeads();
	}
	
}

void Leads::PrintOut(CHListBox* m_output)
{
	m_output->AddString("verified states");
	verifiedLeads.PrintOut(m_output);
	m_output->AddString("internet states");
	nonVerifiedLeads.PrintOut(m_output);
}

void Leads::SetOutput(CHListBox* pointer)
{
	m_output = pointer;
	verifiedLeads.SetOutput(pointer);
	nonVerifiedLeads.SetOutput(pointer);
}

void Leads::Output(CString str)
{
	if(m_output != NULL)
	{
		m_output->AddString(str);
		m_output->SetCurSel(m_output->GetCount()-1);		
		m_output->RedrawWindow();
	}
}

LeadList Leads::GetLeads(bool verified, CString state)
{
	if(verified)
	{
		return verifiedLeads.GetLeads(state);
	}
	else
	{
		return nonVerifiedLeads.GetLeads(state);
	}
}
