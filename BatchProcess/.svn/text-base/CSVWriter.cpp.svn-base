#include "StdAfx.h"
#include ".\csvwriter.h"

CCSVWriter::CCSVWriter(CString target, HWND m_hwnd)
	: CWriter(target, m_hwnd)
{
}

CCSVWriter::~CCSVWriter(void)
{
}

void CCSVWriter::Write(COrder aOrder, CLocation myloc)
{
	//write an order to file as a csv
	try
	{
		CString target;
		
		target = desDir;
		target = target + "\\" + aOrder.GetOrderID() + ".csv";
		
		CFileException e;
		CFile *m_target = new CFile();
		

		if(!m_target->Open(target.GetBuffer(target.GetLength()), CFile::modeCreate | CFile::modeReadWrite, &e))
		{
			PrintOut("Could not open the target file for write at path: " + target);
			
		}
		else
		{
			PrintOut("Opened the target file for write at path: " + target);		
			
			///////////////////////////////////////////////////////////////////////////////////////////////

			CIDArray leadIDs = aOrder.GetAssignedLeadIDs();
			CLeadDB* pt_leadDB = aOrder.GetLeadDBMap();
			CString title;
			if(aOrder.GetVerified())
			{
				title = "Verified Mortgage Leads,";			
			}
			else if(aOrder.GetOnePass())
			{
				title = "One-Pass Mortgage Leads,";			
			}
			else
			{
				title = "Internet Mortgage Leads,";			
			}
			COleDateTime aDate;
			aDate = COleDateTime::GetCurrentTime();
			CString sDate = aDate.Format(_T("%m/%d/%y"));
			title = title + sDate + "\n";

			WriteString(m_target, title);
			CString headings;
			headings = "First Name,Last Name,Co App,City,State,Address,Zip,Work Phone,Home Phone,House Type,Current Value,Desired Loan Amt,First Balance,Rate,Fixed / Adjust,Payment,Late,Credit,Work Info,Time Period,Salary,You Should Call,Looking To,Email,";
			if(aOrder.GetVerified())
			{
				headings = headings + "Verified by,";
				headings = headings + "Personal Touch,";
			}			
			headings = headings + "ID,cid\n";

			WriteString(m_target, headings);

			for(int x = 0; x < leadIDs.GetCount(); x++)
			{
				Progress(x, leadIDs.GetCount()-1);
				CLead aLead = pt_leadDB->GetLead(leadIDs[x]);			
				
				CString record;
				record = MakeCell(aLead.GetLead_FirstName());
				record = record + MakeCell(aLead.GetLead_LastName());
				record = record + MakeCell(aLead.GetLead_CoApp());
				record = record + MakeCell(aLead.GetLead_City());
				record = record + MakeCell(aLead.GetLead_State());
				record = record + MakeCell(aLead.GetLead_Address());
				record = record + MakeCell(aLead.GetLead_Zip());
				record = record + MakeCell(CFormatter::FormatPhone(aLead.GetLead_WorkPhone()));
				record = record + MakeCell(CFormatter::FormatPhone(aLead.GetLead_HomePhone()));
				record = record + MakeCell(aLead.GetLead_HouseType());			
				if(aLead.GetLead_CurrentValue()>0)
				{
					if(aLead.IsPurchase())
					{					
						record = record + ",";	
					}
					else
					{					
						record = record + MakeCell(CFormatter::FormatMoney(aLead.GetLead_CurrentValue_String()));					
					}
				}
				else
				{
					record = record + ",";
				}
				if(aLead.GetLead_DesiredLoan()>0)
				{				
					record = record + MakeCell(CFormatter::FormatMoney(aLead.GetLead_DesiredLoan_String()));				
				}
				else
				{
					record = record + ",";
				}
				if(aLead.GetLead_1stBalance()>0)
				{
					if(aLead.IsPurchase())
					{
						record = record + ",";
					}
					else
					{					
						record = record + MakeCell(CFormatter::FormatMoney(aLead.GetLead_1stBalance_String())); 					
					}
				}
				else
				{
					record = record + ",";
				}
				record = record + MakeCell(aLead.GetLead_Rate_String());	
				record = record + MakeCell(aLead.GetLead_FixedAdjust());			
				record = record + MakeCell(CFormatter::FormatMoney(aLead.GetLead_Payment()));			
				record = record + MakeCell(aLead.GetLead_Late());
				record = record + MakeCell(aLead.GetLead_Credit());	
				record = record + MakeCell(aLead.GetLead_WorkInfo());
				record = record + MakeCell(aLead.GetLead_TimePeriod());				
				if(aLead.GetLead_Salary()>0)
				{				
					record = record + MakeCell(CFormatter::FormatMoney(aLead.GetLead_Salary_String()));							
				}
				else
				{
					record = record + ",";
				}
				record = record + MakeCell(aLead.GetLead_YouShouldCall());
				record = record + MakeCell(aLead.GetLead_LookingTo());
				record = record + MakeCell(aLead.GetLead_Email());
				if(aOrder.GetVerified())
				{
					record = record + MakeCell(aLead.GetVerifiedBy());
					record = record + MakeCell(aLead.GetPersonalTouch());	
				}		
				record = record + MakeCell(aLead.GetLead_Id_String());
				record = record + MakeCell(aOrder.GetOrderID());
				record = record.Left(record.GetLength()-1) + "\n";//terminate record
				
				WriteString(m_target, record);
				

			}

			///////////////////////////////////////////////////////////////////////////////////////////////				
			
			m_target->Close();
		}
	

	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: CCSVWriter::Write(COrder aOrder, CLocation myloc)";
		PrintOut(err);
		CWriteLog::WriteFail(err);
		//WriteLog(err);
	}
}

CString CCSVWriter::MakeCell(CString sContents)
{
	try
	{
		//enclose values in quotes
		CString sCell;
		sCell = sCell + "\"";
		sCell = sCell + sContents; 
		sCell = sCell + "\"";	
		sCell = sCell + ",";

		return sCell;	

	}
	catch(...)
	{
		//error trap
		CString err;
		err = "Error: CCSVWriter::MakeCell(CString sContents)";
		PrintOut(err);
		CWriteLog::WriteFail(err);
		return CString();
	}
}
