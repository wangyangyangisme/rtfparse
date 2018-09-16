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
	//��������
	CTypedPtrArray<CPtrArray,CRtfText*> m_RtfText;

	//������
	CString m_Space;
	// ���뷽ʽ  0--�����  1--����  2--�Ҷ���
	int m_Alignment;
};


