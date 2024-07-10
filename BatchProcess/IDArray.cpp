#include "StdAfx.h"
#include ".\idarray.h"


CIDArray::CIDArray(void)
{
	
}

CIDArray::~CIDArray(void)
{
}
CIDArray& CIDArray::operator=(const CIDArray& aCIDArray)
{
	//assignment operator
	try
	{
		this->myIDs.Copy(aCIDArray.myIDs);	
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: CIDArray::operator=(const CIDArray& aCIDArray)";
		AfxMessageBox(err);
		CWriteLog::WriteFail(err);
	}

	return *this;
}
CIDArray::CIDArray(const CIDArray& aCIDArray)
{
	//copy constructer
	try
	{
		this->myIDs.Copy(aCIDArray.myIDs);	
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: CIDArray::CIDArray(const CIDArray& aCIDArray)";
		AfxMessageBox(err);
		CWriteLog::WriteFail(err);
	}

}

long& CIDArray::operator[] (unsigned long i)
{	
	//array index operator
	try
	{
		return myIDs[i];
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: CIDArray::operator[] (unsigned long i)";
		AfxMessageBox(err);
		CWriteLog::WriteFail(err);
		
	}
}
long& CIDArray::GetAt(long index)
{	
	//get the long at the index
	try
	{
		return myIDs[index];
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: CIDArray::GetAt(long index)";
		AfxMessageBox(err);
		CWriteLog::WriteFail(err);
		
	}
}

INT_PTR CIDArray::Add(long id)
{	
	// add an id number to the array
	try
	{
		return myIDs.Add(id);
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: CIDArray::Add(long id)";
		AfxMessageBox(err);
		CWriteLog::WriteFail(err);
	}
}

INT_PTR CIDArray::Append(CIDArray ids)
{	
	try
	{
		return myIDs.Append(ids.myIDs);
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: CIDArray::Append(CIDArray ids)";
		AfxMessageBox(err);
		CWriteLog::WriteFail(err);
	}
}

INT_PTR CIDArray::GetCount(void)
{
	//count the number of ids in the array
	try
	{
		return myIDs.GetCount();
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: CIDArray::GetCount(void)";
		AfxMessageBox(err);
		CWriteLog::WriteFail(err);
		return -1;
	}
}

void CIDArray::RemoveAll(void)
{	
	try
	{
		myIDs.RemoveAll();
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: CIDArray::RemoveAll";
		AfxMessageBox(err);
		CWriteLog::WriteFail(err);
	}
}

void CIDArray::RemoveAt(INT_PTR index)
{	
	
	try
	{
		myIDs.RemoveAt(index);
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: CIDArray::RemoveAt(INT_PTR index)";
		AfxMessageBox(err);
		CWriteLog::WriteFail(err);
	}
}

void CIDArray::Copy(CIDArray ids)
{
	//copy the ids from one array to the other
	try
	{
		myIDs.Copy(ids.myIDs);
	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: CIDArray::Copy(CIDArray ids)";
		AfxMessageBox(err);
		CWriteLog::WriteFail(err);
	}
}
