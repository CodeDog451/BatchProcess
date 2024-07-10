#include "StdAfx.h"
#include ".\formatter.h"

CFormatter::CFormatter(void)
{
}

CFormatter::~CFormatter(void)
{
}

CString CFormatter::FormatPhone(CString str)
{
	try
	{
		CString phonenum;
		if(str.GetLength()==10)
		{
			phonenum = "(" + str;
			phonenum.Insert(4, ") ");
			
			phonenum.Insert(9, "-");
		}
		else
		{
			phonenum = str;
		}
		return phonenum;
	

	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: CFormatter::FormatPhone(CString str)";
		//PrintOut(err);
		CWriteLog::WriteFail(err);
	}

}

CString CFormatter::FormatMoney(CString str)
{
	try
	{
		CString money = str;
		money.Replace("$", "");
		money.Replace(",", "");
		money.Replace(" ", "");
		if((money=="0")||(money.GetLength()==0))
		{
			return "";
		}
		int y = 0;
		for(int x = money.GetLength(); x >0; x--)
		{
			if(((y % 3)==0)&& (x > 0) && (x < money.GetLength()))
			{
				money.Insert(x, ",");
			}
			y++;
		}
		money = "$" + money;
		money = money + ".00";

		return money;
	

	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: CFormatter::FormatMoney(CString str)";
		//PrintOut(err);
		CWriteLog::WriteFail(err);
		return CString();
	}
}
