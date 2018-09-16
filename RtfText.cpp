// RtfText.cpp : 实现文件
//

#include "stdafx.h"
#include "WDLED.h"
#include "RtfText.h"


// CRtfText

IMPLEMENT_DYNAMIC(CRtfText, CWnd)

CRtfText::CRtfText()
{
	m_FontName = m_Translation.Dt_RtfText.m_FontName;//字体名称
	m_Height=16;//字高
	m_Width=0;//字宽
	m_Bold=false;//粗体
	m_Italic=false;//斜体
	m_Underline=false;//下划线
	m_color=RGB(255,0,0);//颜色
	m_Text="";//文本
	m_Size=0;//字号
}

CRtfText::~CRtfText()
{
}


BEGIN_MESSAGE_MAP(CRtfText, CWnd)
END_MESSAGE_MAP()



// CRtfText 消息处理程序



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
