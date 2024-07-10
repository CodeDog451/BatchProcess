#include "StdAfx.h"
#include ".\rtfwriter.h"
#include "ATLComTime.h"
#include <direct.h>
CRTFWriter::CRTFWriter(CString target, CString rtfTemplate, HWND m_hwnd)
: CWriter(target, m_hwnd) 
, desDir(_T(""))
, templateDir(_T(""))
{
	mainWin = m_hwnd;
	desDir = target;
	templateDir = rtfTemplate;
	if( _mkdir( desDir ) == 0 )
	{
		//PrintOut(desDir + " directory created");
	}
	else
	{
		//PrintOut(desDir + " directory already exists");
	}
	
	
}

CRTFWriter::~CRTFWriter(void)
{
}

void CRTFWriter::PrintOut(CString str)
{
	//myLog.WriteLog(str, 0);
	if(mainWin)
	{
		SendMessage(mainWin, S_START_MESSAGE, 0, 0);
		int iLen = str.GetLength();
		for(int x = 0; x < iLen; x++)
		{
			int iLetter = (int) str.GetAt(x);
			SendMessage(mainWin, S_CHAR_MESSAGE, iLetter, 0);
		}

		SendMessage(mainWin, S_END_MESSAGE, 0, 0);
	}
}

void CRTFWriter::WriteOrderRTF(CString sHeader, CString sFooter, CString sPageBody, COrder aOrder, CString target)
{
	CFileException e;
	CFile *m_target = new CFile();
	

	if(!m_target->Open(target.GetBuffer(target.GetLength()), CFile::modeCreate | CFile::modeReadWrite, &e))
	{
		PrintOut("Could not open the target file for write at path: " + target);
		
	}
	else
	{
		PrintOut("Opened the target file for write at path: " + target);	
		WriteString(m_target, sHeader);
		
		CIDArray leadIDs = aOrder.GetAssignedLeadIDs();
		CLeadDB* pt_leadDB = aOrder.GetLeadDBMap();
		int pagecount = 1;
		char buffer[200];			
		sprintf(buffer, "Count Lead IDs: %d", leadIDs.GetCount());
		PrintOut(buffer);
		for(int x = 0; x < leadIDs.GetCount(); x++)
		{
			Progress(x, leadIDs.GetCount()-1);
			CLead aLead = pt_leadDB->GetLead(leadIDs[x]);
			//PrintOut(aLead.ToString());
			CString aPage = sPageBody;
			COleDateTime aDate;
			aDate = COleDateTime::GetCurrentTime();
			CString sDate = aDate.Format(_T("%m/%d/%y"));

			aPage.Replace("<date>", sDate); 		
			

			if(aOrder.GetVerified())
			{
				aPage.Replace("<title>", "Verified Application"); 
			}
			else if(aOrder.GetOnePass())
			{
				aPage.Replace("<title>", "One-Pass Application"); 
			}
			else
			{
				aPage.Replace("<title>", "Internet Application"); 
			}
			if(aLead.GetLead_Id()> 0)
			{
				aPage.Replace("<id>", aLead.GetLead_Id_String()); 
			}
			else
			{
				aPage.Replace("<id>", ""); 
			}
			CString name = aLead.GetLead_FirstName() + " ";
			name = name + aLead.GetLead_LastName();
			aPage.Replace("<name>", name);
			aPage.Replace("<coapp>", aLead.GetLead_CoApp());
			aPage.Replace("<city>", aLead.GetLead_City());
			aPage.Replace("<state>", aLead.GetLead_State());
			aPage.Replace("<address>", aLead.GetLead_Address());
			aPage.Replace("<zip>", aLead.GetLead_Zip());
			
			if(IsPurchase(aLead.GetLead_LookingTo()))
			{
				aPage.Replace("<subcity>", "");
				aPage.Replace("<substate>", "");
				aPage.Replace("<subaddress>", "Purchase");
				aPage.Replace("<subzip>", "");
			}
			else
			{
				aPage.Replace("<subcity>", aLead.GetLead_City() + ",");
				aPage.Replace("<substate>", aLead.GetLead_State());
				aPage.Replace("<subaddress>", aLead.GetLead_Address() + ",");
				aPage.Replace("<subzip>", aLead.GetLead_Zip());

			}

			aPage.Replace("<workphone>", FormatPhone(aLead.GetLead_WorkPhone()));
			aPage.Replace("<homephone>", FormatPhone(aLead.GetLead_HomePhone()));
			aPage.Replace("<housetype>", aLead.GetLead_HouseType());
			if(aLead.GetLead_CurrentValue()>0)
			{
				if(IsPurchase(aLead.GetLead_LookingTo()))
				{
					//PrintOut("is a purchase");
					aPage.Replace("<currentvalue>", "");
				}
				else
				{
					aPage.Replace("<currentvalue>", FormatMoney(aLead.GetLead_CurrentValue_String()));
				}
			}
			else
			{
				aPage.Replace("<currentvalue>", "");
			}
			if(aLead.GetLead_DesiredLoan()>0)
			{
				aPage.Replace("<desiredloan>", FormatMoney(aLead.GetLead_DesiredLoan_String()));
			}
			else
			{
				aPage.Replace("<desiredloan>", "");
			}
			if(aLead.GetLead_1stBalance()>0)
			{
				if(IsPurchase(aLead.GetLead_LookingTo()))
				{
					aPage.Replace("<1stbal>", "");
				}
				else
				{
					aPage.Replace("<1stbal>", FormatMoney(aLead.GetLead_1stBalance_String()));
				}
			}
			else
			{
				aPage.Replace("<1stbal>", "");
			}
			aPage.Replace("<rate>", aLead.GetLead_Rate_String());
			aPage.Replace("<fixadj>", aLead.GetLead_FixedAdjust());			
			
			aPage.Replace("<payment>", FormatMoney(aLead.GetLead_Payment()));			
			
			aPage.Replace("<late>", aLead.GetLead_Late());
			aPage.Replace("<credit>", aLead.GetLead_Credit());
			aPage.Replace("<workinfo>", aLead.GetLead_WorkInfo());
			aPage.Replace("<timeperiod>", aLead.GetLead_TimePeriod());
			//if(aOrder.GetVerified())
			//{
				//aPage.Replace("<verifyby>", aLead.GetVerifiedBy());
			//}
			//else
			//{
				//aPage.Replace("<verifyby>", "");
			//}
			if(aLead.GetLead_Salary()>0)
			{
				aPage.Replace("<salary>", FormatMoney(aLead.GetLead_Salary_String()));
			}
			else
			{
				aPage.Replace("<salary>", "");
			}
			aPage.Replace("<youcall>", aLead.GetLead_YouShouldCall());
			aPage.Replace("<lookingto>", aLead.GetLead_LookingTo());
			aPage.Replace("<email>", aLead.GetLead_Email());
			aPage.Replace("<personaltouch>", aLead.GetPersonalTouch());
			
			if((!aOrder.GetVerified()) && !(aOrder.GetOneLeadPerPage()) && !(aOrder.GetSendMini1003()))
			{
				//one pass or internet leads go 2 per page
				CLead aLead2;				
				if((x+1) < leadIDs.GetCount())
				{
					x++;
					aLead2 = pt_leadDB->GetLead(leadIDs[x]);
				}
				//aPage.Replace("<id2>", aLead2.GetLead_Id_String()); 
				if(aLead2.GetLead_Id()> 0)
				{
					aPage.Replace("<id2>", aLead2.GetLead_Id_String()); 
				}
				else
				{
					aPage.Replace("<id2>", ""); 
				}
				name = aLead2.GetLead_FirstName() + " ";
				name = name + aLead2.GetLead_LastName();
				aPage.Replace("<name2>", name);
				aPage.Replace("<coapp2>", aLead2.GetLead_CoApp());
				aPage.Replace("<city2>", aLead2.GetLead_City());
				aPage.Replace("<state2>", aLead2.GetLead_State());
				aPage.Replace("<address2>", aLead2.GetLead_Address());
				aPage.Replace("<zip2>", aLead2.GetLead_Zip());
				aPage.Replace("<workphone2>", FormatPhone(aLead2.GetLead_WorkPhone()));
				aPage.Replace("<homephone2>", FormatPhone(aLead2.GetLead_HomePhone()));
				aPage.Replace("<housetype2>", aLead2.GetLead_HouseType());
				if(aLead2.GetLead_CurrentValue()>0)
				{
					if(IsPurchase(aLead2.GetLead_LookingTo()))
					{
						//PrintOut("is a purchase");
						aPage.Replace("<currentvalue2>", "");
					}
					else
					{
						aPage.Replace("<currentvalue2>", FormatMoney(aLead2.GetLead_CurrentValue_String()));
					}
					//aPage.Replace("<currentvalue2>", FormatMoney(aLead2.GetLead_CurrentValue_String()));
				}
				else
				{
					aPage.Replace("<currentvalue2>", "");
				}
				if(aLead2.GetLead_DesiredLoan()>0)
				{
					aPage.Replace("<desiredloan2>", FormatMoney(aLead2.GetLead_DesiredLoan_String()));
				}
				else
				{
					aPage.Replace("<desiredloan2>", "");
				}
				if(aLead2.GetLead_1stBalance()>0)
				{
					if(IsPurchase(aLead2.GetLead_LookingTo()))
					{
						aPage.Replace("<1stbal2>", "");
					}
					else
					{
						aPage.Replace("<1stbal2>", FormatMoney(aLead2.GetLead_1stBalance_String()));
					}
					//aPage.Replace("<1stbal2>", FormatMoney(aLead2.GetLead_1stBalance_String()));
				}
				else
				{
					aPage.Replace("<1stbal2>", "");
				}
				
				aPage.Replace("<rate2>", aLead2.GetLead_Rate_String());
				aPage.Replace("<fixadj2>", aLead2.GetLead_FixedAdjust());
				aPage.Replace("<payment2>", FormatMoney(aLead2.GetLead_Payment()));
				aPage.Replace("<late2>", aLead2.GetLead_Late());
				aPage.Replace("<credit2>", aLead2.GetLead_Credit());
				aPage.Replace("<workinfo2>", aLead2.GetLead_WorkInfo());
				aPage.Replace("<timeperiod2>", aLead2.GetLead_TimePeriod());
				//aPage.Replace("<verifyby2>", aLead2.GetVerifiedBy());
				if(aLead2.GetLead_Salary()>0)
				{
					aPage.Replace("<salary2>", FormatMoney(aLead2.GetLead_Salary_String()));
				}
				else
				{
					aPage.Replace("<salary2>", "");
				}
				
				aPage.Replace("<youcall2>", aLead2.GetLead_YouShouldCall());
				aPage.Replace("<lookingto2>", aLead2.GetLead_LookingTo());
				aPage.Replace("<email2>", aLead2.GetLead_Email());
							
				
			}
			aPage.Replace("<orderid>", aOrder.GetOrderID());
			char buffer[200];			
			sprintf(buffer, "%d", pagecount);
			CString pagenum = buffer;
			aPage.Replace("<pg#>", pagenum);
			aPage.Replace("<contact>", aOrder.GetContact());
			aPage.Replace("<agentname>", aOrder.GetAgent());	
			aPage.Replace("<company>", aOrder.GetCompanyName());
			pagecount++;
			WriteString(m_target, aPage);
			if(x<(leadIDs.GetCount()-1))
			{
				WriteString(m_target, "\\page");
			}

		}
		
		WriteString(m_target, sFooter);
		m_target->Close();
	}
}

int CRTFWriter::WriteString(CFile* pt_target, CString str)
{
	char letter;
	int byteswrite = 0;
	for(int x = 0; x < str.GetLength(); x++)
	{
		letter = str.GetAt(x);
		pt_target->Write(&letter, 1);
		byteswrite++;
	}
	return byteswrite;
}

CString CRTFWriter::FormatPhone(CString str)
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

CString CRTFWriter::FormatMoney(CString str)
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

void CRTFWriter::Write(COrder aOrder)
{
	CString sHead;
	CString sFoot;
	CString sPage;
	CString targetFile = desDir + "\\";
	targetFile = targetFile + aOrder.GetOrderID();
	targetFile = targetFile + ".rtf";

	CString templateFile = templateDir + "\\";
	if(aOrder.GetSendMini1003())
	{
		templateFile = templateFile + "mini1003.rtf";
	}
	else
	{
		if(aOrder.GetVerified() || aOrder.GetOneLeadPerPage())
		{
			templateFile = templateFile + "1perPage.rtf";
		}
		else
		{
			templateFile = templateFile + "2perPage.rtf";
		}
	}


	char letter;
	CString source;
	CFileException e;
	CFile *m_rtfTemplate = new CFile();
	if(!m_rtfTemplate->Open(templateFile.GetBuffer(templateFile.GetLength()), CFile::modeRead, &e))
	{
		PrintOut("Could not open the rtf Template file for reading at path: " + templateFile);
		
	}
	else
	{
		PrintOut("Opened the rtf Template file for reading at path: " + templateFile);		
	
		while(m_rtfTemplate->Read(&letter, 1)>0)
		{
			source = source + letter; 
		}
		if(source.GetLength()>0)
		{
			int iStart = 0;
			int iEnd = 0;
			iStart = source.Find("<startpage>");
			iEnd = source.Find("<endpage>");
			if((iStart > 0) && (iEnd > 0))
			{
				sHead = source.Left(iStart);
				sFoot = source.Right(source.GetLength() - iEnd);
				sFoot.Replace("<endpage>", "");
				sPage = source.Mid(iStart, iEnd - iStart);
				sPage.Replace("<startpage>", "");
				//PrintOut(sHead);
				//PrintOut(sPage);
				//PrintOut(sFoot);
				WriteOrderRTF(sHead, sFoot, sPage, aOrder, targetFile);
				

			}
			else
			{
				PrintOut("Error: <startpage> <endpage> block not found in RTF template");
			}
		}
		else
		{
			PrintOut("Error: RTF template file 0 length not allowed");
		}
		m_rtfTemplate->Close();
	}
	
	
}

bool CRTFWriter::IsPurchase(CString str)
{
	if((str.Find("Pur") > -1) || (str.Find("pur") > -1))
	{
		return true;
	}
	else
	{
		return false;
	}
}
