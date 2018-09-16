// RtfDoc.cpp : 实现文件
//

#include "stdafx.h"
#include "WDLED.h"
#include "RtfDoc.h"


// CRtfDoc

IMPLEMENT_DYNAMIC(CRtfDoc, CWnd)

CRtfDoc::CRtfDoc()
{
}

CRtfDoc::~CRtfDoc()
{
}


BEGIN_MESSAGE_MAP(CRtfDoc, CWnd)
END_MESSAGE_MAP()



// CRtfDoc 消息处理程序



void CRtfDoc::Serialize(CArchive& ar)
{
	if (ar.IsStoring())
	{
		ar<<m_RtfParagraph.GetCount();

		// storing code
		for(long i=0;i<m_RtfParagraph.GetCount();i++)
		{
			m_RtfParagraph[i]->Serialize(ar);
		}
	}
	else
	{	// loading code

		long RtfParagraphCount=0;
		ar>>RtfParagraphCount;

		for(long i=0;i<RtfParagraphCount;i++)
		{
			CRtfParagraph* rtfparagraph=new CRtfParagraph();
			rtfparagraph->Serialize(ar);
			m_RtfParagraph.InsertAt(m_RtfParagraph.GetCount(),rtfparagraph);
		}
	}
}
