#include "StdAfx.h"
#include ".\leadlist.h"

LeadList::LeadList(void)
{
}

LeadList::~LeadList(void)
{
	//for(int x = 0; x< theLeads.GetCount(); x++)
	//{
		//delete theLeads.GetAt(x);
		//theLeads.GetAt(x) = NULL;
		//theLeads.RemoveAt(x);
	//}
	theLeads.RemoveAll();
}

LeadList& LeadList::operator=(const LeadList& aLeadList)
{
	//asignment
	
	theLeads.Copy(aLeadList.theLeads);

	return *this;
}

LeadList::LeadList(const LeadList& aLeadList)
{
	//copy constructor
	
	theLeads.Copy(aLeadList.theLeads);

}

CLead* LeadList::add(CLead* aLead)
{
	//Lead* mylead = new Lead(aLead);
	long index;
	index = Contains(aLead);
	if(index > -1)
	{
		//delete aLead;
		return theLeads.GetAt(index);
	}
	else
	{
		theLeads.Add(aLead);
		return aLead;
	}
}

CLead* LeadList::getAt(long index)
{

	return theLeads[index];
}

long LeadList::getCount(void)
{
	return theLeads.GetCount();
}

long LeadList::getUpperBound(void)
{
	return theLeads.GetUpperBound();
}

void LeadList::Append(LeadList aLeadList)
{
	theLeads.Append(aLeadList.theLeads);
}

long LeadList::Contains(CLead* aLead)
{
	long result = -1;
	long count = theLeads.GetCount();
	for(long x = 0; x < count; x++)
	{		
		if(aLead->GetLead_Id() == ((CLead*)theLeads.GetAt(x))->GetLead_Id())
		{
			result = x;
		}
	}
	return result;
}


