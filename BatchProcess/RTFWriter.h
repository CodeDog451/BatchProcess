#pragma once
#include ".\order.h"
#include "resource.h"
#include "afx.h"
#include "leaddb.h"
#include ".\idarray.h"
#include ".\settings.h"
#include "writer.h"

class CRTFWriter
	:
	public CWriter
{
public:
	CRTFWriter(CString target, CString rtfTemplate, HWND m_hwnd);
	~CRTFWriter(void);
protected:
	
	HWND mainWin;
public:
	void PrintOut(CString str);
protected:
	
public:
	void WriteOrderRTF(CString sHeader, CString sFooter, CString sPageBody, COrder aOrder, CString target);
	int WriteString(CFile* pt_target, CString str);
	CString FormatPhone(CString str);
	CString FormatMoney(CString str);
	void Write(COrder aOrder);
protected:
	CString desDir;
	CString templateDir;
public:
	bool IsPurchase(CString str);
};
