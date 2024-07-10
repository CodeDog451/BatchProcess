#pragma once
#include "writer.h"
//a csv document writter class
class CCSVWriter :
	public CWriter
{
public:
	CCSVWriter(CString target, HWND m_hwnd);
	~CCSVWriter(void);
	void Write(COrder aOrder, CLocation myloc);
	CString MakeCell(CString sContents);
};
