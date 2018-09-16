#pragma once
#include "RtfParagraph.h"

// CRtfDoc

class CRtfDoc : public CWnd
{
	DECLARE_DYNAMIC(CRtfDoc)

public:
	CRtfDoc();
	virtual ~CRtfDoc();

protected:
	DECLARE_MESSAGE_MAP()
public:
	virtual void Serialize(CArchive& ar);
public:
	//∂Œ¬‰¿‡»›∆˜
	CTypedPtrArray<CPtrArray,CRtfParagraph*> m_RtfParagraph;
};


