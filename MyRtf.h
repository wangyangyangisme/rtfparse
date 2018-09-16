#pragma once
#include "RtfDoc.h"

class CMyRtf
{
public:
	CMyRtf(void);
	~CMyRtf(void);
public:
	//��Ա����
	bool ReadRtf(CString filename);//��ȡ�ļ�
	void Serialize(CArchive& ar);//���л�
public:
	CRtfDoc m_RtfDoc;//�ĵ��洢��
	//�ĵ�����ߵ�����
	int m_MaxHeight;
private:
	CString LocationToChinase(CString str);//��λ���ַ�ת�����ַ�
	int HexsToDec(CString numstr,int bit);//16�����ַ�ת����
	COLORREF AsiiToColor(CString rgbstr);//�ַ�תRGB��ɫֵ
private:
	CFile m_File;//�ļ����
};
