#pragma once
#include "RtfDoc.h"

class CMyRtf
{
public:
	CMyRtf(void);
	~CMyRtf(void);
public:
	//成员函数
	bool ReadRtf(CString filename);//读取文件
	void Serialize(CArchive& ar);//序列化
public:
	CRtfDoc m_RtfDoc;//文档存储类
	//文档中最高的字体
	int m_MaxHeight;
private:
	CString LocationToChinase(CString str);//区位码字符转中文字符
	int HexsToDec(CString numstr,int bit);//16进制字符转整数
	COLORREF AsiiToColor(CString rgbstr);//字符转RGB颜色值
private:
	CFile m_File;//文件句柄
};
