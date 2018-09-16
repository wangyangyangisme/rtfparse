#pragma once
#include "RtfText.h"

// CRtfParagraph

class CRtfParagraph : public CWnd
{
	DECLARE_DYNAMIC(CRtfParagraph)

public:
	CRtfParagraph();
	virtual ~CRtfParagraph();

protected:
	DECLARE_MESSAGE_MAP()
public:
	virtual void Serialize(CArchive& ar);
public:
	//字类容器
	CTypedPtrArray<CPtrArray,CRtfText*> m_RtfText;

	//段落间距
	CString m_Space;
	// 对齐方式  0--左对齐  1--居中  2--右对齐
	int m_Alignment;
};


