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
	CString m_FontName;//字体名称
	int m_Size;//字号
	int m_Height;//字高
	int m_Width;//字宽
	bool m_Bold;//粗体
	bool m_Italic;//斜体
	bool m_Underline;//下划线
	COLORREF m_color;//颜色
	CString m_Text;//文本
};


