// RtfParagraph.cpp : ʵ���ļ�
//

#include "stdafx.h"
#include "WDLED.h"
#include "RtfParagraph.h"


// CRtfParagraph

IMPLEMENT_DYNAMIC(CRtfParagraph, CWnd)

CRtfParagraph::CRtfParagraph()
{
	//�����౶��
	m_Space="0.00";
	// ���뷽ʽ  0--�����  1--����  2--�Ҷ���
	m_Alignment=0;
}

CRtfParagraph::~CRtfParagraph()
{
}


BEGIN_MESSAGE_MAP(CRtfParagraph, CWnd)
END_MESSAGE_MAP()



// CRtfParagraph ��Ϣ�������



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
