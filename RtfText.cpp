// RtfText.cpp : ʵ���ļ�
//

#include "stdafx.h"
#include "WDLED.h"
#include "RtfText.h"


// CRtfText

IMPLEMENT_DYNAMIC(CRtfText, CWnd)

CRtfText::CRtfText()
{
	m_FontName = m_Translation.Dt_RtfText.m_FontName;//��������
	m_Height=16;//�ָ�
	m_Width=0;//�ֿ�
	m_Bold=false;//����
	m_Italic=false;//б��
	m_Underline=false;//�»���
	m_color=RGB(255,0,0);//��ɫ
	m_Text="";//�ı�
	m_Size=0;//�ֺ�
}

CRtfText::~CRtfText()
{
}


BEGIN_MESSAGE_MAP(CRtfText, CWnd)
END_MESSAGE_MAP()



// CRtfText ��Ϣ�������



void CRtfText::Serialize(CArchive& ar)
{
	if (ar.IsStoring())
	{	// storing code
		ar<<m_FontName<<m_Height<<m_Width<<m_Bold<<m_Italic<<m_Underline<<m_color<<m_Text<<m_Size;
	}
	else
	{	// loading code
		ar>>m_FontName>>m_Height>>m_Width>>m_Bold>>m_Italic>>m_Underline>>m_color>>m_Text>>m_Size;
	}
}
