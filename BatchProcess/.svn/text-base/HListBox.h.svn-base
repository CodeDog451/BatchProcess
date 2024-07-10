#pragma once
#include "afxwin.h"

class CHListBox : public CListBox
{
// Construction
public:
	CHListBox();

// Attributes
public:

// Operations
public:
    int AddString(LPCTSTR s);
    int InsertString(int i, LPCTSTR s);
    void ResetContent();
    int DeleteString(int i);

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CHListBox)
	//}}AFX_VIRTUAL

// Implementation
public:
	virtual ~CHListBox();

	// Generated message map functions
protected:
        void updateWidth(LPCTSTR s);
	int width;
	//{{AFX_MSG(CHListBox)
		// NOTE - the ClassWizard will add and remove member functions here.
	//}}AFX_MSG

	DECLARE_MESSAGE_MAP()
};
