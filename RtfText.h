#pragma once


// CRtfText

class CRtfText : public CWnd
{
	DECLARE_DYNAMIC(CRtfText)

public:
	CRtfText();
	virtual ~CRtfText();

protected:
	DECLARE_MESSAGE_MAP()
public:
	virtual void Serialize(CArchive& ar);
public:
	CString m_FontName;//��������
	int m_Size;//�ֺ�
	int m_Height;//�ָ�
	int m_Width;//�ֿ�
	bool m_Bold;//����
	bool m_Italic;//б��
	bool m_Underline;//�»���
	COLORREF m_color;//��ɫ
	CString m_Text;//�ı�
};


