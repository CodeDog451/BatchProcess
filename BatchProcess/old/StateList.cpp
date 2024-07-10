#include "StdAfx.h"
#include ".\statelist.h"

StateList::StateList(void)
: m_output(NULL)
{
	ReadStates();
}

StateList::~StateList(void)
{
}
StateList& StateList::operator=(const StateList& aStateList)
{
	//asignment
	
	theStates.Copy(aStateList.theStates);

	return *this;
}

StateList::StateList(const StateList& aStateList)
{
	//copy constructor
	
	theStates.Copy(aStateList.theStates);

}
void StateList::AddState(State aState)
{
	if(Contains(aState.GetSt_Abr()) == -1)
	{
		theStates.Add(aState);
	}
	
}

int StateList::Contains(CString state)
{
	int result;
	result = -1;
	for(int x = 0; x < theStates.GetCount(); x++)
	{
		if(((State)theStates[x]).GetSt_Abr().CompareNoCase(state)==0) 
		{
			result = x;
			
		}
	}
	return result;
}

void StateList::PrintOut(CListBox* m_output)
{
	for(int x = 0; x < theStates.GetCount(); x++)
	{
		((State)theStates[x]).PrintOut(m_output);
	}
}

void StateList::ReadStates(void)
{
	
	HRESULT hr = S_OK;
    try
    {
		//AfxMessageBox("Init database connection");
         CoInitialize(NULL);  
		 
		 _bstr_t strCnn("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=z:\\Mortgage.mdb;Jet OLEDB:Database Password=whompa;");

       _RecordsetPtr pRst = NULL;

      // Call Create instance to instantiate the Record set
      hr = pRst.CreateInstance(__uuidof(Recordset));
      if(FAILED(hr))
      {            
		  AfxMessageBox("Failed creating statelist record set instance");		  
      }
	  else
	  {		
		pRst->Open("SELECT * FROM t_States ORDER BY st_Abrv", strCnn, adOpenStatic, adLockReadOnly, adCmdText);
		pRst->MoveFirst();

		//Loop through the Record set
		if (!pRst->EndOfFile)
		{			
			State aState;
			while((!pRst->EndOfFile))
			{
				aState.ReadState(pRst);
				AddState(aState);					
				pRst->MoveNext();
			}			
			//AfxMessageBox("States Read");			
		}
	  }	  
	}
	catch(_com_error & ce)
	{			
		AfxMessageBox(ce.Description());		
	}	
	CoUninitialize();
	
}

void StateList::AddLead(CLead aLead)
{
	Output(aLead.GetLead_State());
	int index = Contains(aLead.GetLead_State());
	if(index > -1)
	{
		Output("Contains " + aLead.GetLead_State());
		State st = theStates[index];
		st.AddLead(aLead);
		theStates[index] = st;
		//((State)theStates[index]).AddLead(aLead);
	}

}

LeadList StateList::GetLeads(void)
{
	LeadList aLeadList;
	for(int x = 0; x < theStates.GetCount();x++)
	{
		aLeadList.Append(((State)theStates[x]).GetLeads());
	}
	
	return aLeadList;
}

void StateList::SetOutput(CHListBox* pointer)
{
	m_output = pointer;
	for(int x = 0; x < theStates.GetCount();x++)
	{
		State st = theStates[x];
		st.SetOutput(pointer);
		theStates[x] = st;		
		
	}
}

void StateList::Output(CString str)
{
	if(m_output != NULL)
	{
		m_output->AddString(str);
		m_output->SetCurSel(m_output->GetCount()-1);		
		m_output->RedrawWindow();
	}
}

LeadList StateList::GetLeads(CString st_Abr)
{
	LeadList aLeadList;

	for(int x = 0; x < theStates.GetCount(); x++)
	{
		State st;
		st = ((State)theStates.GetAt(x));
		if(st.GetSt_Abr().CompareNoCase(st_Abr)==0)
		{
			return st.GetLeads();
		}
	}
	return aLeadList;
}
