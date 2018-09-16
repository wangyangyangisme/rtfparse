// RtfParagraph.cpp : 实现文件
//

#include "stdafx.h"
#include "WDLED.h"
#include "RtfParagraph.h"


// CRtfParagraph

IMPLEMENT_DYNAMIC(CRtfParagraph, CWnd)

CRtfParagraph::CRtfParagraph()
{
	//段落间距倍数
	m_Space="0.00";
	// 对齐方式  0--左对齐  1--居中  2--右对齐
	m_Alignment=0;
}

CRtfParagraph::~CRtfParagraph()
{
}


BEGIN_MESSAGE_MAP(CRtfParagraph, CWnd)
END_MESSAGE_MAP()



// CRtfParagraph 消息处理程序



void CRtfParagraph::Serialize(CArchive& ar)
{
	if (ar.IsStoring())
	{	// storing code
		ar<<m_Space<<m_Alignment<<m_RtfText.GetCount();

		for(long i=0;i<m_RtfText.GetCount();i++)
		{
			m_RtfText[i]->Serialize(ar);
		}
	}
	else
	{	// loading code
		long RtfTextCount=0;
		ar>>m_Space>>m_Alignment>>RtfTextCount;

		for(long i=0;i<RtfTextCount;i++)
		{
			CRtfText* rtftext=new CRtfText();
			rtftext->Serialize(ar);

			m_RtfText.InsertAt(m_RtfText.GetCount(),rtftext);
		}
	}
}
