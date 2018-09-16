#include "StdAfx.h"
#include "MyRtf.h"

CMyRtf::CMyRtf(void)
{
	m_MaxHeight=0;
}

CMyRtf::~CMyRtf(void)
{
}

//读取Rtf文件
bool CMyRtf::ReadRtf(CString filename)
{
	m_RtfDoc.m_RtfParagraph.RemoveAll();//数据释放

	try
	{
		if(m_File.Open(filename,CFile::modeRead|CFile::typeText))
		{
			char* buf=new char[m_File.GetLength()+2];
			m_File.Read(buf,m_File.GetLength());
			buf[m_File.GetLength()]='\0';
			CString data;
			data=buf;
			data.Replace("\\\\","\\xx");//在文本中"\\" 标示实际"\"
			data.Replace("\\{","\\[[");//在文本中"\\[[" 标示实际"{"
			data.Replace("\\}","\\]]");//在文本中"\\]]" 标示实际"}"

			CArray<CString,CString> m_Fonts;//字体数组列表
			CArray<COLORREF,COLORREF> m_Colors;//颜色数组列表

			int End=0;
			int Start=0;
			Start=data.Find("{\\rtf",End);

			int ShowStart=data.Find("\\viewkind4",End);//文本视图区开始位值
			if(Start!=-1 && Start<10)
			{
				//------字体表
				End=Start+5;
				Start=data.Find("{\\fonttbl",End);
				if(Start!=-1)
				{
					m_Fonts.RemoveAll();//字体列表数据清空

					End=Start+9;
					while(true)
					{
						Start=data.Find("{\\f",End);
						if(Start!=-1) 
						{
							End=Start+3;
							Start=data.Find(" ",End);
							if(Start!=-1)
							{
								End=Start+1;
								int start_old=Start+1;
								Start=data.Find(";}",End);
								if(Start!=-1)
								{
									CString strfont=data.Mid(start_old,Start-start_old);
									End=Start+2;

									m_Fonts.InsertAt(m_Fonts.GetCount(),this->LocationToChinase(strfont));
								}
							}
						}
						else
						{
							break;
						}
					}
				}
				//-----

				//-----颜色表
				Start=data.Find("{\\colortbl",End);

				if(Start!=-1 && Start<ShowStart)
				{
					End=Start+10;
					int Start_old=0;
					CString rgbstr;
					Start=data.Find(";",End);
					if (Start!=-1 && Start<ShowStart)
					{
						m_Colors.RemoveAll();//颜色列表数据清空

						Start_old=End=Start+1;
						while(true)
						{
							Start=data.Find(";",End);
							if (Start!=-1 && Start<ShowStart)
							{
								rgbstr=data.Mid(Start_old,Start-Start_old);
								m_Colors.InsertAt(m_Colors.GetCount(),this->AsiiToColor(rgbstr));
								Start_old=End=Start+1;
							}
							else
							{
								break;
							}
						}
					}
				}
				//-----

				//-----文本视图区解析
				CString textbuf="";//文本数据储存缓存
				long index=0;//数据区计数器
				Start=data.Find("\\viewkind4",End);
				End=Start+10;
				CString d_bufstr;//数据暂存字符串
				int start_old=0;
				Start=data.Find("\\",End);
				start_old=End=Start+1;

				//关键字段判断和数据缓存
				int _uc=0;//字节编码方式 1为单字节
				int _cf=0;//在颜色表中选取颜色
				bool _lang=false;
				int _f=0;//在字体表中选取字体
				int _fs=0;//文本字号 数字位字体字号
				bool _b=false;//粗体开始标志
				bool _i=false;//斜体开始标志
				bool _ul=false;//下划线开始标志
				int _horizontal=0;//水平对齐方式 0-左对齐 1-居中 2-右对齐
				double _sl=240;//段落行距sl加数值
				CString bufstr;
				if(Start!=-1)
				{
					CRtfParagraph *p_rtfparagraph;
					p_rtfparagraph=new CRtfParagraph();
					m_RtfDoc.m_RtfParagraph.InsertAt(m_RtfDoc.m_RtfParagraph.GetCount(),p_rtfparagraph);
					CRtfText *p_rtftext;
					while(true)
					{
						bufstr="";//缓存数据清空
						Start=data.Find("\\",End);
						if(Start!=-1)
						{
							d_bufstr=data.Mid(start_old,Start-start_old);
							start_old=End=Start+1;
						}
						else
						{
							d_bufstr=data.Mid(start_old,data.GetLength()-start_old);
							break;
						}

						if(d_bufstr.Left(2).Find("uc")!=-1)//字节编码方式 1为单字节
						{
							bufstr=d_bufstr;
							bufstr.Replace("uc","");
							_uc=::atoi(bufstr);

							//判断截取半角字符
							int index=bufstr.Find(" ");
							if(index != -1)
							{
								CString oldbuf=bufstr.Mid(index+1,bufstr.GetLength()-index+1);
								for(int k=0;k<oldbuf.GetLength();k++)
								{
									p_rtftext=new CRtfText();
									p_rtftext->m_Text=oldbuf.Mid(k,1);//文本
									if(m_Fonts.GetCount()>0)
										p_rtftext->m_FontName=m_Fonts.GetAt(_f);//字体
									if(m_Colors.GetCount()>0)
										p_rtftext->m_color=m_Colors.GetAt(_cf-1);//颜色
									p_rtftext->m_Size=_fs;//字号
									p_rtftext->m_Bold=_b;//粗体
									p_rtftext->m_Italic=_i;//斜体
									p_rtftext->m_Underline=_ul;//下划线

									int rtftextcount=m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.GetCount();
									m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.InsertAt(rtftextcount,p_rtftext);//加入段
								}
							}
						}
						else if(d_bufstr.Left(3).Find("tab")!=-1)//制造表符
						{
							bufstr=d_bufstr;
							for(int i=0;i<1;i++)
							{
								p_rtftext=new CRtfText();
								p_rtftext->m_Text="　　　　";//文本
								if(m_Fonts.GetCount()>0)
									p_rtftext->m_FontName=m_Fonts.GetAt(_f);//字体
								if(m_Colors.GetCount()>0)
									p_rtftext->m_color=m_Colors.GetAt(_cf-1);//颜色
								p_rtftext->m_Size=_fs;//字号
								p_rtftext->m_Bold=_b;//粗体
								p_rtftext->m_Italic=_i;//斜体
								p_rtftext->m_Underline=_ul;//下划线

								int rtftextcount=m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.GetCount();
								m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.InsertAt(rtftextcount,p_rtftext);//加入段
							}

							//判断截取半角字符
							int index=bufstr.Find(" ");
							if(index != -1)
							{
								CString oldbuf=bufstr.Mid(index+1,bufstr.GetLength()-index+1);
								for(int k=0;k<oldbuf.GetLength();k++)
								{
									p_rtftext=new CRtfText();
									p_rtftext->m_Text=oldbuf.Mid(k,1);//文本
									if(m_Fonts.GetCount()>0)
										p_rtftext->m_FontName=m_Fonts.GetAt(_f);//字体
									if(m_Colors.GetCount()>0)
										p_rtftext->m_color=m_Colors.GetAt(_cf-1);//颜色
									p_rtftext->m_Size=_fs;//字号
									p_rtftext->m_Bold=_b;//粗体
									p_rtftext->m_Italic=_i;//斜体
									p_rtftext->m_Underline=_ul;//下划线

									int rtftextcount=m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.GetCount();
									m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.InsertAt(rtftextcount,p_rtftext);//加入段
								}
							}
						}
						else if(d_bufstr.Left(2).Find("cf")!=-1)//在颜色表中选取颜色
						{
							bufstr=d_bufstr;
							bufstr.Replace("cf","");
							_cf=::atoi(bufstr);

							//判断截取半角字符
							int index=bufstr.Find(" ");
							if(index != -1)
							{
								CString oldbuf=bufstr.Mid(index+1,bufstr.GetLength()-index+1);
								for(int k=0;k<oldbuf.GetLength();k++)
								{
									p_rtftext=new CRtfText();
									p_rtftext->m_Text=oldbuf.Mid(k,1);//文本
									if(m_Fonts.GetCount()>0)
										p_rtftext->m_FontName=m_Fonts.GetAt(_f);//字体
									if(m_Colors.GetCount()>0)
										p_rtftext->m_color=m_Colors.GetAt(_cf-1);//颜色
									p_rtftext->m_Size=_fs;//字号
									p_rtftext->m_Bold=_b;//粗体
									p_rtftext->m_Italic=_i;//斜体
									p_rtftext->m_Underline=_ul;//下划线

									int rtftextcount=m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.GetCount();
									m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.InsertAt(rtftextcount,p_rtftext);//加入段
								}
							}
						}
						else if(d_bufstr.Left(8).Find("lang2052")!=-1)
						{
							bufstr=d_bufstr;
							//判断截取半角字符
							int index=bufstr.Find(" ");
							if(index != -1)
							{
								CString oldbuf=bufstr.Mid(index+1,bufstr.GetLength()-index+1);
								for(int k=0;k<oldbuf.GetLength();k++)
								{
									p_rtftext=new CRtfText();
									p_rtftext->m_Text=oldbuf.Mid(k,1);//文本
									if(m_Fonts.GetCount()>0)
										p_rtftext->m_FontName=m_Fonts.GetAt(_f);//字体
									if(m_Colors.GetCount()>0)
										p_rtftext->m_color=m_Colors.GetAt(_cf-1);//颜色
									p_rtftext->m_Size=_fs;//字号
									p_rtftext->m_Bold=_b;//粗体
									p_rtftext->m_Italic=_i;//斜体
									p_rtftext->m_Underline=_ul;//下划线

									int rtftextcount=m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.GetCount();
									m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.InsertAt(rtftextcount,p_rtftext);//加入段
								}
							}
						}
						else if(d_bufstr.Left(1).Find("f")!=-1)//在字体表中选取字体
						{
							if(d_bufstr.Left(2).Find("fs")!=-1)//文本字号 数字位字体字号
							{
								bufstr=d_bufstr.Mid(2,d_bufstr.GetLength()-2);
								_fs=::atoi(bufstr)/1.5;

								//冒泡排序排出最高字体
								if(_fs>m_MaxHeight)
									m_MaxHeight=_fs;

								//判断截取半角字符
								int index=bufstr.Find(" ");
								if(index != -1)
								{
									CString oldbuf=bufstr.Mid(index+1,bufstr.GetLength()-index+1);
									for(int k=0;k<oldbuf.GetLength();k++)
									{
										p_rtftext=new CRtfText();
										p_rtftext->m_Text=oldbuf.Mid(k,1);//文本
										if(m_Fonts.GetCount()>0)
											p_rtftext->m_FontName=m_Fonts.GetAt(_f);//字体
										if(m_Colors.GetCount()>0)
											p_rtftext->m_color=m_Colors.GetAt(_cf-1);//颜色
										p_rtftext->m_Size=_fs;//字号
										p_rtftext->m_Bold=_b;//粗体
										p_rtftext->m_Italic=_i;//斜体
										p_rtftext->m_Underline=_ul;//下划线

										int rtftextcount=m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.GetCount();
										m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.InsertAt(rtftextcount,p_rtftext);//加入段
									}
								}
							}
							else
							{
								bufstr=d_bufstr;
								bufstr.Replace("f","");
								_f=::atoi(bufstr);

								//判断截取半角字符
								int index=bufstr.Find(" ");
								if(index != -1)
								{
									CString oldbuf=bufstr.Mid(index+1,bufstr.GetLength()-index+1);
									for(int k=0;k<oldbuf.GetLength();k++)
									{
										p_rtftext=new CRtfText();
										p_rtftext->m_Text=oldbuf.Mid(k,1);//文本
										if(m_Fonts.GetCount()>0)
											p_rtftext->m_FontName=m_Fonts.GetAt(_f);//字体
										if(m_Colors.GetCount()>0)
											p_rtftext->m_color=m_Colors.GetAt(_cf-1);//颜色
										p_rtftext->m_Size=_fs;//字号
										p_rtftext->m_Bold=_b;//粗体
										p_rtftext->m_Italic=_i;//斜体
										p_rtftext->m_Underline=_ul;//下划线

										int rtftextcount=m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.GetCount();
										m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.InsertAt(rtftextcount,p_rtftext);//加入段
									}
								}
							}
						}
						else if(d_bufstr.Left(1).Find("b")!=-1)//粗体开始标志
						{
							if(d_bufstr.Left(2).Find("b0")!=-1)//粗体结束标志
							{
								bufstr=d_bufstr;
								_b=false;

								//判断截取半角字符
								int index=bufstr.Find(" ");
								if(index != -1)
								{
									CString oldbuf=bufstr.Mid(index+1,bufstr.GetLength()-index+1);
									for(int k=0;k<oldbuf.GetLength();k++)
									{
										p_rtftext=new CRtfText();
										p_rtftext->m_Text=oldbuf.Mid(k,1);//文本
										if(m_Fonts.GetCount()>0)
											p_rtftext->m_FontName=m_Fonts.GetAt(_f);//字体
										if(m_Colors.GetCount()>0)
											p_rtftext->m_color=m_Colors.GetAt(_cf-1);//颜色
										p_rtftext->m_Size=_fs;//字号
										p_rtftext->m_Bold=_b;//粗体
										p_rtftext->m_Italic=_i;//斜体
										p_rtftext->m_Underline=_ul;//下划线

										int rtftextcount=m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.GetCount();
										m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.InsertAt(rtftextcount,p_rtftext);//加入段
									}
								}

							}
							else
							{
								bufstr=d_bufstr;
								_b=true;

								//判断截取半角字符
								int index=bufstr.Find(" ");
								if(index != -1)
								{
									CString oldbuf=bufstr.Mid(index+1,bufstr.GetLength()-index+1);
									for(int k=0;k<oldbuf.GetLength();k++)
									{
										p_rtftext=new CRtfText();
										p_rtftext->m_Text=oldbuf.Mid(k,1);//文本
										if(m_Fonts.GetCount()>0)
											p_rtftext->m_FontName=m_Fonts.GetAt(_f);//字体
										if(m_Colors.GetCount()>0)
											p_rtftext->m_color=m_Colors.GetAt(_cf-1);//颜色
										p_rtftext->m_Size=_fs;//字号
										p_rtftext->m_Bold=_b;//粗体
										p_rtftext->m_Italic=_i;//斜体
										p_rtftext->m_Underline=_ul;//下划线

										int rtftextcount=m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.GetCount();
										m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.InsertAt(rtftextcount,p_rtftext);//加入段
									}
								}
							}
						}
						else if(d_bufstr.Left(1).Find("i")!=-1)//斜体开始标志
						{
							if(d_bufstr.Left(2).Find("i0")!=-1)//斜体结束标志
							{
								bufstr=d_bufstr;
								_i=false;

								//判断截取半角字符
								int index=bufstr.Find(" ");
								if(index != -1)
								{
									CString oldbuf=bufstr.Mid(index+1,bufstr.GetLength()-index+1);
									for(int k=0;k<oldbuf.GetLength();k++)
									{
										p_rtftext=new CRtfText();
										p_rtftext->m_Text=oldbuf.Mid(k,1);//文本
										if(m_Fonts.GetCount()>0)
											p_rtftext->m_FontName=m_Fonts.GetAt(_f);//字体
										if(m_Colors.GetCount()>0)
											p_rtftext->m_color=m_Colors.GetAt(_cf-1);//颜色
										p_rtftext->m_Size=_fs;//字号
										p_rtftext->m_Bold=_b;//粗体
										p_rtftext->m_Italic=_i;//斜体
										p_rtftext->m_Underline=_ul;//下划线

										int rtftextcount=m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.GetCount();
										m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.InsertAt(rtftextcount,p_rtftext);//加入段
									}
								}

							}
							else
							{
								bufstr=d_bufstr;
								_i=true;

								//判断截取半角字符
								int index=bufstr.Find(" ");
								if(index != -1)
								{
									CString oldbuf=bufstr.Mid(index+1,bufstr.GetLength()-index+1);
									for(int k=0;k<oldbuf.GetLength();k++)
									{
										p_rtftext=new CRtfText();
										p_rtftext->m_Text=oldbuf.Mid(k,1);//文本
										if(m_Fonts.GetCount()>0)
											p_rtftext->m_FontName=m_Fonts.GetAt(_f);//字体
										if(m_Colors.GetCount()>0)
											p_rtftext->m_color=m_Colors.GetAt(_cf-1);//颜色
										p_rtftext->m_Size=_fs;//字号
										p_rtftext->m_Bold=_b;//粗体
										p_rtftext->m_Italic=_i;//斜体
										p_rtftext->m_Underline=_ul;//下划线

										int rtftextcount=m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.GetCount();
										m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.InsertAt(rtftextcount,p_rtftext);//加入段
									}
								}
							}
						}
						else if(d_bufstr.Left(2).Find("ul")!=-1)//下划线开始标志
						{
							if(d_bufstr.Left(6).Find("ulnone")!=-1)//下划线结束标志
							{
								bufstr=d_bufstr;
								_ul=false;

								//判断截取半角字符
								int index=bufstr.Find(" ");
								if(index != -1)
								{
									CString oldbuf=bufstr.Mid(index+1,bufstr.GetLength()-index+1);
									for(int k=0;k<oldbuf.GetLength();k++)
									{
										p_rtftext=new CRtfText();
										p_rtftext->m_Text=oldbuf.Mid(k,1);//文本
										if(m_Fonts.GetCount()>0)
											p_rtftext->m_FontName=m_Fonts.GetAt(_f);//字体
										if(m_Colors.GetCount()>0)
											p_rtftext->m_color=m_Colors.GetAt(_cf-1);//颜色
										p_rtftext->m_Size=_fs;//字号
										p_rtftext->m_Bold=_b;//粗体
										p_rtftext->m_Italic=_i;//斜体
										p_rtftext->m_Underline=_ul;//下划线

										int rtftextcount=m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.GetCount();
										m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.InsertAt(rtftextcount,p_rtftext);//加入段
									}
								}

							}
							else
							{
								bufstr=d_bufstr;
								_ul=true;

								//判断截取半角字符
								int index=bufstr.Find(" ");
								if(index != -1)
								{
									CString oldbuf=bufstr.Mid(index+1,bufstr.GetLength()-index+1);
									for(int k=0;k<oldbuf.GetLength();k++)
									{
										p_rtftext=new CRtfText();
										p_rtftext->m_Text=oldbuf.Mid(k,1);//文本
										if(m_Fonts.GetCount()>0)
											p_rtftext->m_FontName=m_Fonts.GetAt(_f);//字体
										if(m_Colors.GetCount()>0)
											p_rtftext->m_color=m_Colors.GetAt(_cf-1);//颜色
										p_rtftext->m_Size=_fs;//字号
										p_rtftext->m_Bold=_b;//粗体
										p_rtftext->m_Italic=_i;//斜体
										p_rtftext->m_Underline=_ul;//下划线

										int rtftextcount=m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.GetCount();
										m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.InsertAt(rtftextcount,p_rtftext);//加入段
									}
								}
							}
						}
						else if(d_bufstr.Left(3).Find("par")!=-1)//段落标记
						{
							if(d_bufstr.Left(4).Find("pard")!=-1)//默认段落属性表示结束上一段的段落属性
							{
								bufstr=d_bufstr;
								_horizontal=0;//水平对齐方式 0-左对齐 1-居中 2-右对齐
								_sl=240;//段落行距sl加数值

								m_RtfDoc.m_RtfParagraph.GetAt(m_RtfDoc.m_RtfParagraph.GetCount()-1)->m_Alignment=_horizontal;//对齐方式
								//间距倍数计算
								CString space;
								space.Format("%0.2f",(_sl-240)/240);
								m_RtfDoc.m_RtfParagraph.GetAt(m_RtfDoc.m_RtfParagraph.GetCount()-1)->m_Space=space;//间距倍数

								//判断截取半角字符
								int index=bufstr.Find(" ");
								if(index != -1)
								{
									CString oldbuf=bufstr.Mid(index+1,bufstr.GetLength()-index+1);
									for(int k=0;k<oldbuf.GetLength();k++)
									{
										p_rtftext=new CRtfText();
										p_rtftext->m_Text=oldbuf.Mid(k,1);//文本
										if(m_Fonts.GetCount()>0)
											p_rtftext->m_FontName=m_Fonts.GetAt(_f);//字体
										if(m_Colors.GetCount()>0)
											p_rtftext->m_color=m_Colors.GetAt(_cf-1);//颜色
										p_rtftext->m_Size=_fs;//字号
										p_rtftext->m_Bold=_b;//粗体
										p_rtftext->m_Italic=_i;//斜体
										p_rtftext->m_Underline=_ul;//下划线

										int rtftextcount=m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.GetCount();
										m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.InsertAt(rtftextcount,p_rtftext);//加入段
									}
								}
							}
							else
							{
								bufstr=d_bufstr;
								p_rtfparagraph=new CRtfParagraph();
								m_RtfDoc.m_RtfParagraph.InsertAt(m_RtfDoc.m_RtfParagraph.GetCount(),p_rtfparagraph);
								m_RtfDoc.m_RtfParagraph.GetAt(m_RtfDoc.m_RtfParagraph.GetCount()-1)->m_Alignment=_horizontal;//对齐方式
								//间距倍数计算
								CString space;
								space.Format("%0.2f",(_sl-240)/240);
								m_RtfDoc.m_RtfParagraph.GetAt(m_RtfDoc.m_RtfParagraph.GetCount()-1)->m_Space=space;//间距倍数

								//判断截取半角字符
								if(bufstr.GetLength()>5)
								{
									CString oldbuf=bufstr.Mid(5,bufstr.GetLength()-5);
									for(int k=0;k<oldbuf.GetLength();k++)
									{
										p_rtftext=new CRtfText();
										p_rtftext->m_Text=oldbuf.Mid(k,1);//文本
										if(m_Fonts.GetCount()>0)
											p_rtftext->m_FontName=m_Fonts.GetAt(_f);//字体
										if(m_Colors.GetCount()>0)
											p_rtftext->m_color=m_Colors.GetAt(_cf-1);//颜色
										p_rtftext->m_Size=_fs;//字号
										p_rtftext->m_Bold=_b;//粗体
										p_rtftext->m_Italic=_i;//斜体
										p_rtftext->m_Underline=_ul;//下划线

										int rtftextcount=m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.GetCount();
										m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.InsertAt(rtftextcount,p_rtftext);//加入段
									}
								}

							}
						}
						else if(d_bufstr.Left(2).Find("qc")!=-1)//居中对齐
						{
							bufstr=d_bufstr;
							_horizontal=1;
							m_RtfDoc.m_RtfParagraph.GetAt(m_RtfDoc.m_RtfParagraph.GetCount()-1)->m_Alignment=_horizontal;//对齐方式

							//判断截取半角字符
							int index=bufstr.Find(" ");
							if(index != -1)
							{
								CString oldbuf=bufstr.Mid(index+1,bufstr.GetLength()-index+1);
								for(int k=0;k<oldbuf.GetLength();k++)
								{
									p_rtftext=new CRtfText();
									p_rtftext->m_Text=oldbuf.Mid(k,1);//文本
									if(m_Fonts.GetCount()>0)
										p_rtftext->m_FontName=m_Fonts.GetAt(_f);//字体
									if(m_Colors.GetCount()>0)
										p_rtftext->m_color=m_Colors.GetAt(_cf-1);//颜色
									p_rtftext->m_Size=_fs;//字号
									p_rtftext->m_Bold=_b;//粗体
									p_rtftext->m_Italic=_i;//斜体
									p_rtftext->m_Underline=_ul;//下划线

									int rtftextcount=m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.GetCount();
									m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.InsertAt(rtftextcount,p_rtftext);//加入段
								}
							}
						}
						else if(d_bufstr.Left(2).Find("qr")!=-1)//右对齐
						{
							bufstr=d_bufstr;
							_horizontal=2;
							m_RtfDoc.m_RtfParagraph.GetAt(m_RtfDoc.m_RtfParagraph.GetCount()-1)->m_Alignment=_horizontal;//对齐方式

							//判断截取半角字符
							int index=bufstr.Find(" ");
							if(index != -1)
							{
								CString oldbuf=bufstr.Mid(index+1,bufstr.GetLength()-index+1);
								for(int k=0;k<oldbuf.GetLength();k++)
								{
									p_rtftext=new CRtfText();
									p_rtftext->m_Text=oldbuf.Mid(k,1);//文本
									if(m_Fonts.GetCount()>0)
										p_rtftext->m_FontName=m_Fonts.GetAt(_f);//字体
									if(m_Colors.GetCount()>0)
										p_rtftext->m_color=m_Colors.GetAt(_cf-1);//颜色
									p_rtftext->m_Size=_fs;//字号
									p_rtftext->m_Bold=_b;//粗体
									p_rtftext->m_Italic=_i;//斜体
									p_rtftext->m_Underline=_ul;//下划线

									int rtftextcount=m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.GetCount();
									m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.InsertAt(rtftextcount,p_rtftext);//加入段
								}
							}
						}
						else if(d_bufstr.Left(2).Find("sl")!=-1)//段落行距sl加数值
						{
							if(d_bufstr.Left(6).Find("slmult")!=-1)
							{
								bufstr=d_bufstr;
								//判断截取半角字符
								int index=bufstr.Find(" ");
								if(index != -1)
								{
									CString oldbuf=bufstr.Mid(index+1,bufstr.GetLength()-index+1);
									for(int k=0;k<oldbuf.GetLength();k++)
									{
										p_rtftext=new CRtfText();
										p_rtftext->m_Text=oldbuf.Mid(k,1);//文本
										if(m_Fonts.GetCount()>0)
											p_rtftext->m_FontName=m_Fonts.GetAt(_f);//字体
										if(m_Colors.GetCount()>0)
											p_rtftext->m_color=m_Colors.GetAt(_cf-1);//颜色
										p_rtftext->m_Size=_fs;//字号
										p_rtftext->m_Bold=_b;//粗体
										p_rtftext->m_Italic=_i;//斜体
										p_rtftext->m_Underline=_ul;//下划线

										int rtftextcount=m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.GetCount();
										m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.InsertAt(rtftextcount,p_rtftext);//加入段
									}
								}

							}
							else
							{
								bufstr=d_bufstr;
								bufstr.Replace("sl","");
								_sl=::atoi(bufstr);

								//间距倍数计算
								CString space;
								space.Format("%0.2f",(_sl-240)/240);
								m_RtfDoc.m_RtfParagraph.GetAt(m_RtfDoc.m_RtfParagraph.GetCount()-1)->m_Space=space;//间距倍数

								//判断截取半角字符
								int index=bufstr.Find(" ");
								if(index != -1)
								{
									CString oldbuf=bufstr.Mid(index+1,bufstr.GetLength()-index+1);
									for(int k=0;k<oldbuf.GetLength();k++)
									{
										p_rtftext=new CRtfText();
										p_rtftext->m_Text=oldbuf.Mid(k,1);//文本
										if(m_Fonts.GetCount()>0)
											p_rtftext->m_FontName=m_Fonts.GetAt(_f);//字体
										if(m_Colors.GetCount()>0)
											p_rtftext->m_color=m_Colors.GetAt(_cf-1);//颜色
										p_rtftext->m_Size=_fs;//字号
										p_rtftext->m_Bold=_b;//粗体
										p_rtftext->m_Italic=_i;//斜体
										p_rtftext->m_Underline=_ul;//下划线

										int rtftextcount=m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.GetCount();
										m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.InsertAt(rtftextcount,p_rtftext);//加入段
									}
								}
							}
						}
						else if(d_bufstr.Left(2).Find("xx")!=-1)//半角字符"\" 在RTF文件中编码格式为"\\"
						{
							CString oldbuf="";
							if(d_bufstr.GetLength()>2)
							{
								oldbuf=d_bufstr.Mid(2,d_bufstr.GetLength()-2);
							}
							p_rtftext=new CRtfText();
							p_rtftext->m_Text="\\";//文本
							if(m_Fonts.GetCount()>0)
								p_rtftext->m_FontName=m_Fonts.GetAt(_f);//字体
							if(m_Colors.GetCount()>0)
								p_rtftext->m_color=m_Colors.GetAt(_cf-1);//颜色
							p_rtftext->m_Size=_fs;//字号
							p_rtftext->m_Bold=_b;//粗体
							p_rtftext->m_Italic=_i;//斜体
							p_rtftext->m_Underline=_ul;//下划线

							int rtftextcount=m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.GetCount();
							m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.InsertAt(rtftextcount,p_rtftext);//加入段

							if(oldbuf.GetLength()>0)
							{
								for(int k=0;k<oldbuf.GetLength();k++)
								{
									p_rtftext=new CRtfText();
									p_rtftext->m_Text=oldbuf.Mid(k,1);//文本
									if(m_Fonts.GetCount()>0)
										p_rtftext->m_FontName=m_Fonts.GetAt(_f);//字体
									if(m_Colors.GetCount()>0)
										p_rtftext->m_color=m_Colors.GetAt(_cf-1);//颜色
									p_rtftext->m_Size=_fs;//字号
									p_rtftext->m_Bold=_b;//粗体
									p_rtftext->m_Italic=_i;//斜体
									p_rtftext->m_Underline=_ul;//下划线

									int rtftextcount=m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.GetCount();
									m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.InsertAt(rtftextcount,p_rtftext);//加入段
								}
							}
						}
						else if(d_bufstr.Left(2).Find("[[")!=-1)//半角字符"{" 在RTF文件中编码格式为"[["
						{
							CString oldbuf="";
							if(d_bufstr.GetLength()>2)
							{
								oldbuf=d_bufstr.Mid(2,d_bufstr.GetLength()-2);
							}
							p_rtftext=new CRtfText();
							p_rtftext->m_Text="{";//文本
							if(m_Fonts.GetCount()>0)
								p_rtftext->m_FontName=m_Fonts.GetAt(_f);//字体
							if(m_Colors.GetCount()>0)
								p_rtftext->m_color=m_Colors.GetAt(_cf-1);//颜色
							p_rtftext->m_Size=_fs;//字号
							p_rtftext->m_Bold=_b;//粗体
							p_rtftext->m_Italic=_i;//斜体
							p_rtftext->m_Underline=_ul;//下划线

							int rtftextcount=m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.GetCount();
							m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.InsertAt(rtftextcount,p_rtftext);//加入段

							if(oldbuf.GetLength()>0)
							{
								for(int k=0;k<oldbuf.GetLength();k++)
								{
									p_rtftext=new CRtfText();
									p_rtftext->m_Text=oldbuf.Mid(k,1);//文本
									if(m_Fonts.GetCount()>0)
										p_rtftext->m_FontName=m_Fonts.GetAt(_f);//字体
									if(m_Colors.GetCount()>0)
										p_rtftext->m_color=m_Colors.GetAt(_cf-1);//颜色
									p_rtftext->m_Size=_fs;//字号
									p_rtftext->m_Bold=_b;//粗体
									p_rtftext->m_Italic=_i;//斜体
									p_rtftext->m_Underline=_ul;//下划线

									int rtftextcount=m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.GetCount();
									m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.InsertAt(rtftextcount,p_rtftext);//加入段
								}
							}
						}
						else if(d_bufstr.Left(2).Find("]]")!=-1)//半角字符"}" 在RTF文件中编码格式为"]]"
						{
							CString oldbuf="";
							if(d_bufstr.GetLength()>2)
							{
								oldbuf=d_bufstr.Mid(2,d_bufstr.GetLength()-2);
							}
							p_rtftext=new CRtfText();
							p_rtftext->m_Text="}";//文本
							if(m_Fonts.GetCount()>0)
								p_rtftext->m_FontName=m_Fonts.GetAt(_f);//字体
							if(m_Colors.GetCount()>0)
								p_rtftext->m_color=m_Colors.GetAt(_cf-1);//颜色
							p_rtftext->m_Size=_fs;//字号
							p_rtftext->m_Bold=_b;//粗体
							p_rtftext->m_Italic=_i;//斜体
							p_rtftext->m_Underline=_ul;//下划线

							int rtftextcount=m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.GetCount();
							m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.InsertAt(rtftextcount,p_rtftext);//加入段

							if(oldbuf.GetLength()>0)
							{
								for(int k=0;k<oldbuf.GetLength();k++)
								{
									p_rtftext=new CRtfText();
									p_rtftext->m_Text=oldbuf.Mid(k,1);//文本
									if(m_Fonts.GetCount()>0)
										p_rtftext->m_FontName=m_Fonts.GetAt(_f);//字体
									if(m_Colors.GetCount()>0)
										p_rtftext->m_color=m_Colors.GetAt(_cf-1);//颜色
									p_rtftext->m_Size=_fs;//字号
									p_rtftext->m_Bold=_b;//粗体
									p_rtftext->m_Italic=_i;//斜体
									p_rtftext->m_Underline=_ul;//下划线

									int rtftextcount=m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.GetCount();
									m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.InsertAt(rtftextcount,p_rtftext);//加入段
								}
							}
						}
						else if(d_bufstr.Left(9).Find("ldblquote")!=-1)//全角"“"字符
						{
							CString oldbuf="";
							if(d_bufstr.GetLength()>9)
							{
								oldbuf=d_bufstr.Mid(9,d_bufstr.GetLength()-9);
							}
							p_rtftext=new CRtfText();
							p_rtftext->m_Text="“";//文本
							if(m_Fonts.GetCount()>0)
								p_rtftext->m_FontName=m_Fonts.GetAt(_f);//字体
							if(m_Colors.GetCount()>0)
								p_rtftext->m_color=m_Colors.GetAt(_cf-1);//颜色
							p_rtftext->m_Size=_fs;//字号
							p_rtftext->m_Bold=_b;//粗体
							p_rtftext->m_Italic=_i;//斜体
							p_rtftext->m_Underline=_ul;//下划线

							int rtftextcount=m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.GetCount();
							m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.InsertAt(rtftextcount,p_rtftext);//加入段

							if(oldbuf.GetLength()>0)
							{
								for(int k=0;k<oldbuf.GetLength();k++)
								{
									p_rtftext=new CRtfText();
									p_rtftext->m_Text=oldbuf.Mid(k,1);//文本
									if(m_Fonts.GetCount()>0)
										p_rtftext->m_FontName=m_Fonts.GetAt(_f);//字体
									if(m_Colors.GetCount()>0)
										p_rtftext->m_color=m_Colors.GetAt(_cf-1);//颜色
									p_rtftext->m_Size=_fs;//字号
									p_rtftext->m_Bold=_b;//粗体
									p_rtftext->m_Italic=_i;//斜体
									p_rtftext->m_Underline=_ul;//下划线

									int rtftextcount=m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.GetCount();
									m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.InsertAt(rtftextcount,p_rtftext);//加入段
								}
							}
						}
						else if(d_bufstr.Left(9).Find("rdblquote")!=-1)//全角"”"字符
						{
							CString oldbuf="";
							if(d_bufstr.GetLength()>9)
							{
								oldbuf=d_bufstr.Mid(9,d_bufstr.GetLength()-9);
							}
							p_rtftext=new CRtfText();
							p_rtftext->m_Text="”";//文本
							if(m_Fonts.GetCount()>0)
								p_rtftext->m_FontName=m_Fonts.GetAt(_f);//字体
							if(m_Colors.GetCount()>0)
								p_rtftext->m_color=m_Colors.GetAt(_cf-1);//颜色
							p_rtftext->m_Size=_fs;//字号
							p_rtftext->m_Bold=_b;//粗体
							p_rtftext->m_Italic=_i;//斜体
							p_rtftext->m_Underline=_ul;//下划线

							int rtftextcount=m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.GetCount();
							m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.InsertAt(rtftextcount,p_rtftext);//加入段

							if(oldbuf.GetLength()>0)
							{
								for(int k=0;k<oldbuf.GetLength();k++)
								{
									p_rtftext=new CRtfText();
									p_rtftext->m_Text=oldbuf.Mid(k,1);//文本
									if(m_Fonts.GetCount()>0)
										p_rtftext->m_FontName=m_Fonts.GetAt(_f);//字体
									if(m_Colors.GetCount()>0)
										p_rtftext->m_color=m_Colors.GetAt(_cf-1);//颜色
									p_rtftext->m_Size=_fs;//字号
									p_rtftext->m_Bold=_b;//粗体
									p_rtftext->m_Italic=_i;//斜体
									p_rtftext->m_Underline=_ul;//下划线

									int rtftextcount=m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.GetCount();
									m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.InsertAt(rtftextcount,p_rtftext);//加入段
								}
							}
						}
						else if(d_bufstr.Left(6).Find("lquote")!=-1)//全角"‘"字符
						{
							CString oldbuf="";
							if(d_bufstr.GetLength()>6)
							{
								oldbuf=d_bufstr.Mid(6,d_bufstr.GetLength()-6);
							}
							p_rtftext=new CRtfText();
							p_rtftext->m_Text="‘";//文本
							if(m_Fonts.GetCount()>0)
								p_rtftext->m_FontName=m_Fonts.GetAt(_f);//字体
							if(m_Colors.GetCount()>0)
								p_rtftext->m_color=m_Colors.GetAt(_cf-1);//颜色
							p_rtftext->m_Size=_fs;//字号
							p_rtftext->m_Bold=_b;//粗体
							p_rtftext->m_Italic=_i;//斜体
							p_rtftext->m_Underline=_ul;//下划线

							int rtftextcount=m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.GetCount();
							m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.InsertAt(rtftextcount,p_rtftext);//加入段

							if(oldbuf.GetLength()>0)
							{
								for(int k=0;k<oldbuf.GetLength();k++)
								{
									p_rtftext=new CRtfText();
									p_rtftext->m_Text=oldbuf.Mid(k,1);//文本
									if(m_Fonts.GetCount()>0)
										p_rtftext->m_FontName=m_Fonts.GetAt(_f);//字体
									if(m_Colors.GetCount()>0)
										p_rtftext->m_color=m_Colors.GetAt(_cf-1);//颜色
									p_rtftext->m_Size=_fs;//字号
									p_rtftext->m_Bold=_b;//粗体
									p_rtftext->m_Italic=_i;//斜体
									p_rtftext->m_Underline=_ul;//下划线

									int rtftextcount=m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.GetCount();
									m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.InsertAt(rtftextcount,p_rtftext);//加入段
								}
							}
						}
						else if(d_bufstr.Left(6).Find("rquote")!=-1)//全角"’"字符
						{
							CString oldbuf="";
							if(d_bufstr.GetLength()>6)
							{
								oldbuf=d_bufstr.Mid(6,d_bufstr.GetLength()-6);
							}
							p_rtftext=new CRtfText();
							p_rtftext->m_Text="’";//文本
							if(m_Fonts.GetCount()>0)
								p_rtftext->m_FontName=m_Fonts.GetAt(_f);//字体
							if(m_Colors.GetCount()>0)
								p_rtftext->m_color=m_Colors.GetAt(_cf-1);//颜色
							p_rtftext->m_Size=_fs;//字号
							p_rtftext->m_Bold=_b;//粗体
							p_rtftext->m_Italic=_i;//斜体
							p_rtftext->m_Underline=_ul;//下划线

							int rtftextcount=m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.GetCount();
							m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.InsertAt(rtftextcount,p_rtftext);//加入段

							if(oldbuf.GetLength()>0)
							{
								for(int k=0;k<oldbuf.GetLength();k++)
								{
									p_rtftext=new CRtfText();
									p_rtftext->m_Text=oldbuf.Mid(k,1);//文本
									if(m_Fonts.GetCount()>0)
										p_rtftext->m_FontName=m_Fonts.GetAt(_f);//字体
									if(m_Colors.GetCount()>0)
										p_rtftext->m_color=m_Colors.GetAt(_cf-1);//颜色
									p_rtftext->m_Size=_fs;//字号
									p_rtftext->m_Bold=_b;//粗体
									p_rtftext->m_Italic=_i;//斜体
									p_rtftext->m_Underline=_ul;//下划线

									int rtftextcount=m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.GetCount();
									m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.InsertAt(rtftextcount,p_rtftext);//加入段
								}
							}
						}
						else if(d_bufstr.Left(6).Find("emdash")!=-1)//全角"—"字符
						{
							CString oldbuf="";
							if(d_bufstr.GetLength()>6)
							{
								oldbuf=d_bufstr.Mid(6,d_bufstr.GetLength()-6);
							}
							p_rtftext=new CRtfText();
							p_rtftext->m_Text="—";//文本
							if(m_Fonts.GetCount()>0)
								p_rtftext->m_FontName=m_Fonts.GetAt(_f);//字体
							if(m_Colors.GetCount()>0)
								p_rtftext->m_color=m_Colors.GetAt(_cf-1);//颜色
							p_rtftext->m_Size=_fs;//字号
							p_rtftext->m_Bold=_b;//粗体
							p_rtftext->m_Italic=_i;//斜体
							p_rtftext->m_Underline=_ul;//下划线

							int rtftextcount=m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.GetCount();
							m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.InsertAt(rtftextcount,p_rtftext);//加入段

							if(oldbuf.GetLength()>0)
							{
								for(int k=0;k<oldbuf.GetLength();k++)
								{
									p_rtftext=new CRtfText();
									p_rtftext->m_Text=oldbuf.Mid(k,1);//文本
									if(m_Fonts.GetCount()>0)
										p_rtftext->m_FontName=m_Fonts.GetAt(_f);//字体
									if(m_Colors.GetCount()>0)
										p_rtftext->m_color=m_Colors.GetAt(_cf-1);//颜色
									p_rtftext->m_Size=_fs;//字号
									p_rtftext->m_Bold=_b;//粗体
									p_rtftext->m_Italic=_i;//斜体
									p_rtftext->m_Underline=_ul;//下划线

									int rtftextcount=m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.GetCount();
									m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.InsertAt(rtftextcount,p_rtftext);//加入段
								}
							}
						}
						else if(d_bufstr.Left(1).Find("'")!=-1)//中文区位码数据
						{
							bufstr=d_bufstr;
							CString oldbuf="";
							bufstr.Replace("'","");
							int h_num=0,l_num=0;
							if(bufstr.GetLength()>0)
							{
								if(bufstr.GetLength()>2)
								{
									h_num=this->HexsToDec(bufstr.Mid(0,1),16);
									l_num=this->HexsToDec(bufstr.Mid(1,1),1);
									textbuf+=(char)(h_num+l_num);

									oldbuf=bufstr.Mid(2,bufstr.GetLength()-2);//存储半角字符
								}
								else
								{
									h_num=this->HexsToDec(bufstr.Mid(0,1),16);
									l_num=this->HexsToDec(bufstr.Mid(1,1),1);
									textbuf+=(char)(h_num+l_num);
								}
							}

							index++;
							if(index%2==0)
							{
								//数据入类库
								p_rtftext=new CRtfText();
								p_rtftext->m_Text=textbuf;//文本
								if(m_Fonts.GetCount()>0)
									p_rtftext->m_FontName=m_Fonts.GetAt(_f);//字体
								if(m_Colors.GetCount()>0)
									p_rtftext->m_color=m_Colors.GetAt(_cf-1);//颜色
								p_rtftext->m_Size=_fs;//字号
								p_rtftext->m_Bold=_b;//粗体
								p_rtftext->m_Italic=_i;//斜体
								p_rtftext->m_Underline=_ul;//下划线

								int rtftextcount=m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.GetCount();
								m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.InsertAt(rtftextcount,p_rtftext);//加入段

								if(oldbuf.GetLength()>0)
								{
									for(int k=0;k<oldbuf.GetLength();k++)
									{
										p_rtftext=new CRtfText();
										p_rtftext->m_Text=oldbuf.Mid(k,1);//文本
										if(m_Fonts.GetCount()>0)
											p_rtftext->m_FontName=m_Fonts.GetAt(_f);//字体
										if(m_Colors.GetCount()>0)
											p_rtftext->m_color=m_Colors.GetAt(_cf-1);//颜色
										p_rtftext->m_Size=_fs;//字号
										p_rtftext->m_Bold=_b;//粗体
										p_rtftext->m_Italic=_i;//斜体
										p_rtftext->m_Underline=_ul;//下划线

										int rtftextcount=m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.GetCount();
										m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.InsertAt(rtftextcount,p_rtftext);//加入段
									}
								}

								textbuf="";//清空文本缓存区数据
							}

						}

					}
				}
			}
			else
			{
				////内存释放
				//if(buf!=NULL)
				//	delete buf;

				m_File.Close();
				AfxMessageBox(m_Translation.Dt_MyRtf.m_Msgstr);
				return false;
			}
		
		}
		else
		{
			return false;
		}

		m_File.Close();

	}
	catch(CException* e)
	{
		char* errmsg=new char[2024];
		e->GetErrorMessage(errmsg,2024);
		AfxMessageBox(errmsg);
	}

	return true;
}

CString CMyRtf::LocationToChinase(CString location_str)//区位码字符转中文字符
{
	//断开每段区位码
	int Start=0,End=0,start_old=0;
	CString lstr,ret_str="";
	CArray<CString,CString> m_liststr;

	int _index=0;//记录器 判断是否需要解析区位码 _index<=0不需要解析区位码
	while(true)
	{
		Start=location_str.Find("\\'",End);
		if(Start>0 && m_liststr.GetCount()>0)
		{
			_index++;
			lstr=location_str.Mid(start_old,Start-start_old);
			start_old=End=Start+2;
			m_liststr.InsertAt(m_liststr.GetCount(),lstr);
		}
		else if(Start==0 && m_liststr.GetCount()==0)
		{
			_index++;
			lstr=location_str.Mid(start_old,Start-start_old);
			start_old=End=Start+2;
			m_liststr.InsertAt(m_liststr.GetCount(),lstr);
		}
		else if(Start>0 && m_liststr.GetCount()==0)
		{
			_index++;
			lstr=location_str.Mid(start_old,Start-start_old);
			start_old=End=Start+2;
			m_liststr.InsertAt(m_liststr.GetCount(),"|" + lstr);
		}
		else
		{
			if(_index<=0)
			{
				return location_str;
			}

			lstr=location_str.Mid(start_old,location_str.GetLength()-start_old);
			m_liststr.InsertAt(m_liststr.GetCount(),lstr);
			break;
		}
	}

	if(m_liststr.GetCount()<=0)
	{
		return "errNULL";
	}

	//解析区位码
	CString buf_str="";//没有编码字符数组缓存区
	CString head_buf;
	int ch_num=0;//字符十进制编码值
	CString asiistr;
	CString wcp_strOne;
	CString wcp_strTwo = "|";
	for(int i=0;i<m_liststr.GetCount();i++)
	{
		CString ch_str=m_liststr.GetAt(i);
		if(ch_str.GetLength()>0)
		{
			wcp_strOne = ch_str;
			if(wcp_strOne.Compare(wcp_strTwo) == 1)
			{
				asiistr = "";
				head_buf = ch_str.Mid(1, ch_str.GetLength()-1);
			}
			else
			{
				if(ch_str.GetLength() > 2)
				{
					asiistr=ch_str.Left(2);

					buf_str=ch_str.Mid(2,ch_str.GetLength()-2);
				}
				else
				{
					asiistr=ch_str;
				}


				CString numstr=asiistr.Mid(0,1);
				int h_num=this->HexsToDec(numstr,16);//16进制 十位字符
				numstr=asiistr.Mid(1,1);
				int l_num=this->HexsToDec(numstr,1);//16进制 个位字符
				ch_num=h_num+l_num;
			}

			
		}

		if(ch_num > 0)
		{
			ret_str+=(char)ch_num;

			if(buf_str != "")
			{
				ret_str+=buf_str;
			}

			if(head_buf != "")
			{
				ret_str = head_buf + ret_str;
				head_buf = "";
			}
		}
	}
	return ret_str;
}

//16进制字符转整数
int CMyRtf::HexsToDec(CString numstr,int bit)
{
	int num=0;
	if(numstr=="a")
		num=10*bit;
	else if(numstr=="b")
		num=11*bit;
	else if(numstr=="c")
		num=12*bit;
	else if(numstr=="d")
		num=13*bit;
	else if(numstr=="e")
		num=14*bit;
	else if(numstr=="f")
		num=15*bit;
	else
		num=::atoi(numstr)*bit;

	return num;
}

//字符转RGB颜色值
COLORREF CMyRtf::AsiiToColor(CString rgbstr)
{
	CString colstr;
	int red=255,green=255,blue=255;
	int Start=0,End=0,start_old;
	Start=rgbstr.Find("\\",End);
	start_old=End=Start+1;
	if(Start!=-1)
	{
		//--red 红色
		Start=rgbstr.Find("\\",End);
		if(Start!=-1)
		{
			colstr=rgbstr.Mid(start_old,Start-start_old);
			start_old=End=Start+1;

			colstr.Replace("red","");
			red=::atoi(colstr);
		}
		//--

		//--green 绿色
		Start=rgbstr.Find("\\",End);
		if(Start!=-1)
		{
			colstr=rgbstr.Mid(start_old,Start-start_old);
			start_old=End=Start+1;

			colstr.Replace("green","");
			green=::atoi(colstr);
		}
		//--

		//--blue 黄色
		colstr=rgbstr.Mid(start_old,rgbstr.GetLength()-start_old);
		start_old=End=Start+1;

		colstr.Replace("blue","");
		blue=::atoi(colstr);
		//--
	}

	return RGB(red,green,blue);
}

void CMyRtf::Serialize(CArchive &ar)
{
	if (ar.IsStoring())
	{// storing code
		m_RtfDoc.Serialize(ar);
		ar<<m_MaxHeight;
	}
	else
	{// loading code
		m_RtfDoc.Serialize(ar);
		ar>>m_MaxHeight;
	}
}
