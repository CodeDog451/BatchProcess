#pragma once
#include "dbitem.h"
#include "comutil.h"
#import "C:\Program Files\Common Files\System\ADO\msado15.dll" \
no_namespace rename("EOF", "EndOfFile")
class CLogDB :
	public DBItem
{
public:
	CLogDB(void);
	~CLogDB(void);
protected:
	_RecordsetPtr pRst;
	_RecordsetPtr pRstFail;
	//_RecordsetPtr pRstFax;
	//_RecordsetPtr pRstFaxFail;
	_bstr_t strCnn;
	//HRESULT hr;
	//HRESULT hrFail;
	//HRESULT hrFax;
	//HRESULT hrFaxFail;
public:
	void WriteLog(CString str, int div);
	void WriteFail(CString str, int div);
protected:
	bool Connect(_RecordsetPtr &pt_Rst, _bstr_t strCon, CString sql);
	void Write(_RecordsetPtr &pt_Rst, CString str, int div);
public:
	//void WriteFaxLog(CString str, int div);
	//void WriteFaxFail(CString str, int div);
	
};
