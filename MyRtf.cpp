#include "StdAfx.h"
#include "MyRtf.h"

CMyRtf::CMyRtf(void)
{
	m_MaxHeight=0;
}

CMyRtf::~CMyRtf(void)
{
}

//��ȡRtf�ļ�
bool CMyRtf::ReadRtf(CString filename)
{
	m_RtfDoc.m_RtfParagraph.RemoveAll();//�����ͷ�

	try
	{
		if(m_File.Open(filename,CFile::modeRead|CFile::typeText))
		{
			char* buf=new char[m_File.GetLength()+2];
			m_File.Read(buf,m_File.GetLength());
			buf[m_File.GetLength()]='\0';
			CString data;
			data=buf;
			data.Replace("\\\\","\\xx");//���ı���"\\" ��ʾʵ��"\"
			data.Replace("\\{","\\[[");//���ı���"\\[[" ��ʾʵ��"{"
			data.Replace("\\}","\\]]");//���ı���"\\]]" ��ʾʵ��"}"

			CArray<CString,CString> m_Fonts;//���������б�
			CArray<COLORREF,COLORREF> m_Colors;//��ɫ�����б�

			int End=0;
			int Start=0;
			Start=data.Find("{\\rtf",End);

			int ShowStart=data.Find("\\viewkind4",End);//�ı���ͼ����ʼλֵ
			if(Start!=-1 && Start<10)
			{
				//------�����
				End=Start+5;
				Start=data.Find("{\\fonttbl",End);
				if(Start!=-1)
				{
					m_Fonts.RemoveAll();//�����б��������

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

				//-----��ɫ��
				Start=data.Find("{\\colortbl",End);

				if(Start!=-1 && Start<ShowStart)
				{
					End=Start+10;
					int Start_old=0;
					CString rgbstr;
					Start=data.Find(";",End);
					if (Start!=-1 && Start<ShowStart)
					{
						m_Colors.RemoveAll();//��ɫ�б��������

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

				//-----�ı���ͼ������
				CString textbuf="";//�ı����ݴ��滺��
				long index=0;//������������
				Start=data.Find("\\viewkind4",End);
				End=Start+10;
				CString d_bufstr;//�����ݴ��ַ���
				int start_old=0;
				Start=data.Find("\\",End);
				start_old=End=Start+1;

				//�ؼ��ֶ��жϺ����ݻ���
				int _uc=0;//�ֽڱ��뷽ʽ 1Ϊ���ֽ�
				int _cf=0;//����ɫ����ѡȡ��ɫ
				bool _lang=false;
				int _f=0;//���������ѡȡ����
				int _fs=0;//�ı��ֺ� ����λ�����ֺ�
				bool _b=false;//���忪ʼ��־
				bool _i=false;//б�忪ʼ��־
				bool _ul=false;//�»��߿�ʼ��־
				int _horizontal=0;//ˮƽ���뷽ʽ 0-����� 1-���� 2-�Ҷ���
				double _sl=240;//�����о�sl����ֵ
				CString bufstr;
				if(Start!=-1)
				{
					CRtfParagraph *p_rtfparagraph;
					p_rtfparagraph=new CRtfParagraph();
					m_RtfDoc.m_RtfParagraph.InsertAt(m_RtfDoc.m_RtfParagraph.GetCount(),p_rtfparagraph);
					CRtfText *p_rtftext;
					while(true)
					{
						bufstr="";//�����������
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

						if(d_bufstr.Left(2).Find("uc")!=-1)//�ֽڱ��뷽ʽ 1Ϊ���ֽ�
						{
							bufstr=d_bufstr;
							bufstr.Replace("uc","");
							_uc=::atoi(bufstr);

							//�жϽ�ȡ����ַ�
							int index=bufstr.Find(" ");
							if(index != -1)
							{
								CString oldbuf=bufstr.Mid(index+1,bufstr.GetLength()-index+1);
								for(int k=0;k<oldbuf.GetLength();k++)
								{
									p_rtftext=new CRtfText();
									p_rtftext->m_Text=oldbuf.Mid(k,1);//�ı�
									if(m_Fonts.GetCount()>0)
										p_rtftext->m_FontName=m_Fonts.GetAt(_f);//����
									if(m_Colors.GetCount()>0)
										p_rtftext->m_color=m_Colors.GetAt(_cf-1);//��ɫ
									p_rtftext->m_Size=_fs;//�ֺ�
									p_rtftext->m_Bold=_b;//����
									p_rtftext->m_Italic=_i;//б��
									p_rtftext->m_Underline=_ul;//�»���

									int rtftextcount=m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.GetCount();
									m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.InsertAt(rtftextcount,p_rtftext);//�����
								}
							}
						}
						else if(d_bufstr.Left(3).Find("tab")!=-1)//������
						{
							bufstr=d_bufstr;
							for(int i=0;i<1;i++)
							{
								p_rtftext=new CRtfText();
								p_rtftext->m_Text="��������";//�ı�
								if(m_Fonts.GetCount()>0)
									p_rtftext->m_FontName=m_Fonts.GetAt(_f);//����
								if(m_Colors.GetCount()>0)
									p_rtftext->m_color=m_Colors.GetAt(_cf-1);//��ɫ
								p_rtftext->m_Size=_fs;//�ֺ�
								p_rtftext->m_Bold=_b;//����
								p_rtftext->m_Italic=_i;//б��
								p_rtftext->m_Underline=_ul;//�»���

								int rtftextcount=m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.GetCount();
								m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.InsertAt(rtftextcount,p_rtftext);//�����
							}

							//�жϽ�ȡ����ַ�
							int index=bufstr.Find(" ");
							if(index != -1)
							{
								CString oldbuf=bufstr.Mid(index+1,bufstr.GetLength()-index+1);
								for(int k=0;k<oldbuf.GetLength();k++)
								{
									p_rtftext=new CRtfText();
									p_rtftext->m_Text=oldbuf.Mid(k,1);//�ı�
									if(m_Fonts.GetCount()>0)
										p_rtftext->m_FontName=m_Fonts.GetAt(_f);//����
									if(m_Colors.GetCount()>0)
										p_rtftext->m_color=m_Colors.GetAt(_cf-1);//��ɫ
									p_rtftext->m_Size=_fs;//�ֺ�
									p_rtftext->m_Bold=_b;//����
									p_rtftext->m_Italic=_i;//б��
									p_rtftext->m_Underline=_ul;//�»���

									int rtftextcount=m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.GetCount();
									m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.InsertAt(rtftextcount,p_rtftext);//�����
								}
							}
						}
						else if(d_bufstr.Left(2).Find("cf")!=-1)//����ɫ����ѡȡ��ɫ
						{
							bufstr=d_bufstr;
							bufstr.Replace("cf","");
							_cf=::atoi(bufstr);

							//�жϽ�ȡ����ַ�
							int index=bufstr.Find(" ");
							if(index != -1)
							{
								CString oldbuf=bufstr.Mid(index+1,bufstr.GetLength()-index+1);
								for(int k=0;k<oldbuf.GetLength();k++)
								{
									p_rtftext=new CRtfText();
									p_rtftext->m_Text=oldbuf.Mid(k,1);//�ı�
									if(m_Fonts.GetCount()>0)
										p_rtftext->m_FontName=m_Fonts.GetAt(_f);//����
									if(m_Colors.GetCount()>0)
										p_rtftext->m_color=m_Colors.GetAt(_cf-1);//��ɫ
									p_rtftext->m_Size=_fs;//�ֺ�
									p_rtftext->m_Bold=_b;//����
									p_rtftext->m_Italic=_i;//б��
									p_rtftext->m_Underline=_ul;//�»���

									int rtftextcount=m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.GetCount();
									m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.InsertAt(rtftextcount,p_rtftext);//�����
								}
							}
						}
						else if(d_bufstr.Left(8).Find("lang2052")!=-1)
						{
							bufstr=d_bufstr;
							//�жϽ�ȡ����ַ�
							int index=bufstr.Find(" ");
							if(index != -1)
							{
								CString oldbuf=bufstr.Mid(index+1,bufstr.GetLength()-index+1);
								for(int k=0;k<oldbuf.GetLength();k++)
								{
									p_rtftext=new CRtfText();
									p_rtftext->m_Text=oldbuf.Mid(k,1);//�ı�
									if(m_Fonts.GetCount()>0)
										p_rtftext->m_FontName=m_Fonts.GetAt(_f);//����
									if(m_Colors.GetCount()>0)
										p_rtftext->m_color=m_Colors.GetAt(_cf-1);//��ɫ
									p_rtftext->m_Size=_fs;//�ֺ�
									p_rtftext->m_Bold=_b;//����
									p_rtftext->m_Italic=_i;//б��
									p_rtftext->m_Underline=_ul;//�»���

									int rtftextcount=m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.GetCount();
									m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.InsertAt(rtftextcount,p_rtftext);//�����
								}
							}
						}
						else if(d_bufstr.Left(1).Find("f")!=-1)//���������ѡȡ����
						{
							if(d_bufstr.Left(2).Find("fs")!=-1)//�ı��ֺ� ����λ�����ֺ�
							{
								bufstr=d_bufstr.Mid(2,d_bufstr.GetLength()-2);
								_fs=::atoi(bufstr)/1.5;

								//ð�������ų��������
								if(_fs>m_MaxHeight)
									m_MaxHeight=_fs;

								//�жϽ�ȡ����ַ�
								int index=bufstr.Find(" ");
								if(index != -1)
								{
									CString oldbuf=bufstr.Mid(index+1,bufstr.GetLength()-index+1);
									for(int k=0;k<oldbuf.GetLength();k++)
									{
										p_rtftext=new CRtfText();
										p_rtftext->m_Text=oldbuf.Mid(k,1);//�ı�
										if(m_Fonts.GetCount()>0)
											p_rtftext->m_FontName=m_Fonts.GetAt(_f);//����
										if(m_Colors.GetCount()>0)
											p_rtftext->m_color=m_Colors.GetAt(_cf-1);//��ɫ
										p_rtftext->m_Size=_fs;//�ֺ�
										p_rtftext->m_Bold=_b;//����
										p_rtftext->m_Italic=_i;//б��
										p_rtftext->m_Underline=_ul;//�»���

										int rtftextcount=m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.GetCount();
										m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.InsertAt(rtftextcount,p_rtftext);//�����
									}
								}
							}
							else
							{
								bufstr=d_bufstr;
								bufstr.Replace("f","");
								_f=::atoi(bufstr);

								//�жϽ�ȡ����ַ�
								int index=bufstr.Find(" ");
								if(index != -1)
								{
									CString oldbuf=bufstr.Mid(index+1,bufstr.GetLength()-index+1);
									for(int k=0;k<oldbuf.GetLength();k++)
									{
										p_rtftext=new CRtfText();
										p_rtftext->m_Text=oldbuf.Mid(k,1);//�ı�
										if(m_Fonts.GetCount()>0)
											p_rtftext->m_FontName=m_Fonts.GetAt(_f);//����
										if(m_Colors.GetCount()>0)
											p_rtftext->m_color=m_Colors.GetAt(_cf-1);//��ɫ
										p_rtftext->m_Size=_fs;//�ֺ�
										p_rtftext->m_Bold=_b;//����
										p_rtftext->m_Italic=_i;//б��
										p_rtftext->m_Underline=_ul;//�»���

										int rtftextcount=m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.GetCount();
										m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.InsertAt(rtftextcount,p_rtftext);//�����
									}
								}
							}
						}
						else if(d_bufstr.Left(1).Find("b")!=-1)//���忪ʼ��־
						{
							if(d_bufstr.Left(2).Find("b0")!=-1)//���������־
							{
								bufstr=d_bufstr;
								_b=false;

								//�жϽ�ȡ����ַ�
								int index=bufstr.Find(" ");
								if(index != -1)
								{
									CString oldbuf=bufstr.Mid(index+1,bufstr.GetLength()-index+1);
									for(int k=0;k<oldbuf.GetLength();k++)
									{
										p_rtftext=new CRtfText();
										p_rtftext->m_Text=oldbuf.Mid(k,1);//�ı�
										if(m_Fonts.GetCount()>0)
											p_rtftext->m_FontName=m_Fonts.GetAt(_f);//����
										if(m_Colors.GetCount()>0)
											p_rtftext->m_color=m_Colors.GetAt(_cf-1);//��ɫ
										p_rtftext->m_Size=_fs;//�ֺ�
										p_rtftext->m_Bold=_b;//����
										p_rtftext->m_Italic=_i;//б��
										p_rtftext->m_Underline=_ul;//�»���

										int rtftextcount=m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.GetCount();
										m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.InsertAt(rtftextcount,p_rtftext);//�����
									}
								}

							}
							else
							{
								bufstr=d_bufstr;
								_b=true;

								//�жϽ�ȡ����ַ�
								int index=bufstr.Find(" ");
								if(index != -1)
								{
									CString oldbuf=bufstr.Mid(index+1,bufstr.GetLength()-index+1);
									for(int k=0;k<oldbuf.GetLength();k++)
									{
										p_rtftext=new CRtfText();
										p_rtftext->m_Text=oldbuf.Mid(k,1);//�ı�
										if(m_Fonts.GetCount()>0)
											p_rtftext->m_FontName=m_Fonts.GetAt(_f);//����
										if(m_Colors.GetCount()>0)
											p_rtftext->m_color=m_Colors.GetAt(_cf-1);//��ɫ
										p_rtftext->m_Size=_fs;//�ֺ�
										p_rtftext->m_Bold=_b;//����
										p_rtftext->m_Italic=_i;//б��
										p_rtftext->m_Underline=_ul;//�»���

										int rtftextcount=m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.GetCount();
										m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.InsertAt(rtftextcount,p_rtftext);//�����
									}
								}
							}
						}
						else if(d_bufstr.Left(1).Find("i")!=-1)//б�忪ʼ��־
						{
							if(d_bufstr.Left(2).Find("i0")!=-1)//б�������־
							{
								bufstr=d_bufstr;
								_i=false;

								//�жϽ�ȡ����ַ�
								int index=bufstr.Find(" ");
								if(index != -1)
								{
									CString oldbuf=bufstr.Mid(index+1,bufstr.GetLength()-index+1);
									for(int k=0;k<oldbuf.GetLength();k++)
									{
										p_rtftext=new CRtfText();
										p_rtftext->m_Text=oldbuf.Mid(k,1);//�ı�
										if(m_Fonts.GetCount()>0)
											p_rtftext->m_FontName=m_Fonts.GetAt(_f);//����
										if(m_Colors.GetCount()>0)
											p_rtftext->m_color=m_Colors.GetAt(_cf-1);//��ɫ
										p_rtftext->m_Size=_fs;//�ֺ�
										p_rtftext->m_Bold=_b;//����
										p_rtftext->m_Italic=_i;//б��
										p_rtftext->m_Underline=_ul;//�»���

										int rtftextcount=m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.GetCount();
										m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.InsertAt(rtftextcount,p_rtftext);//�����
									}
								}

							}
							else
							{
								bufstr=d_bufstr;
								_i=true;

								//�жϽ�ȡ����ַ�
								int index=bufstr.Find(" ");
								if(index != -1)
								{
									CString oldbuf=bufstr.Mid(index+1,bufstr.GetLength()-index+1);
									for(int k=0;k<oldbuf.GetLength();k++)
									{
										p_rtftext=new CRtfText();
										p_rtftext->m_Text=oldbuf.Mid(k,1);//�ı�
										if(m_Fonts.GetCount()>0)
											p_rtftext->m_FontName=m_Fonts.GetAt(_f);//����
										if(m_Colors.GetCount()>0)
											p_rtftext->m_color=m_Colors.GetAt(_cf-1);//��ɫ
										p_rtftext->m_Size=_fs;//�ֺ�
										p_rtftext->m_Bold=_b;//����
										p_rtftext->m_Italic=_i;//б��
										p_rtftext->m_Underline=_ul;//�»���

										int rtftextcount=m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.GetCount();
										m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.InsertAt(rtftextcount,p_rtftext);//�����
									}
								}
							}
						}
						else if(d_bufstr.Left(2).Find("ul")!=-1)//�»��߿�ʼ��־
						{
							if(d_bufstr.Left(6).Find("ulnone")!=-1)//�»��߽�����־
							{
								bufstr=d_bufstr;
								_ul=false;

								//�жϽ�ȡ����ַ�
								int index=bufstr.Find(" ");
								if(index != -1)
								{
									CString oldbuf=bufstr.Mid(index+1,bufstr.GetLength()-index+1);
									for(int k=0;k<oldbuf.GetLength();k++)
									{
										p_rtftext=new CRtfText();
										p_rtftext->m_Text=oldbuf.Mid(k,1);//�ı�
										if(m_Fonts.GetCount()>0)
											p_rtftext->m_FontName=m_Fonts.GetAt(_f);//����
										if(m_Colors.GetCount()>0)
											p_rtftext->m_color=m_Colors.GetAt(_cf-1);//��ɫ
										p_rtftext->m_Size=_fs;//�ֺ�
										p_rtftext->m_Bold=_b;//����
										p_rtftext->m_Italic=_i;//б��
										p_rtftext->m_Underline=_ul;//�»���

										int rtftextcount=m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.GetCount();
										m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.InsertAt(rtftextcount,p_rtftext);//�����
									}
								}

							}
							else
							{
								bufstr=d_bufstr;
								_ul=true;

								//�жϽ�ȡ����ַ�
								int index=bufstr.Find(" ");
								if(index != -1)
								{
									CString oldbuf=bufstr.Mid(index+1,bufstr.GetLength()-index+1);
									for(int k=0;k<oldbuf.GetLength();k++)
									{
										p_rtftext=new CRtfText();
										p_rtftext->m_Text=oldbuf.Mid(k,1);//�ı�
										if(m_Fonts.GetCount()>0)
											p_rtftext->m_FontName=m_Fonts.GetAt(_f);//����
										if(m_Colors.GetCount()>0)
											p_rtftext->m_color=m_Colors.GetAt(_cf-1);//��ɫ
										p_rtftext->m_Size=_fs;//�ֺ�
										p_rtftext->m_Bold=_b;//����
										p_rtftext->m_Italic=_i;//б��
										p_rtftext->m_Underline=_ul;//�»���

										int rtftextcount=m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.GetCount();
										m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.InsertAt(rtftextcount,p_rtftext);//�����
									}
								}
							}
						}
						else if(d_bufstr.Left(3).Find("par")!=-1)//������
						{
							if(d_bufstr.Left(4).Find("pard")!=-1)//Ĭ�϶������Ա�ʾ������һ�εĶ�������
							{
								bufstr=d_bufstr;
								_horizontal=0;//ˮƽ���뷽ʽ 0-����� 1-���� 2-�Ҷ���
								_sl=240;//�����о�sl����ֵ

								m_RtfDoc.m_RtfParagraph.GetAt(m_RtfDoc.m_RtfParagraph.GetCount()-1)->m_Alignment=_horizontal;//���뷽ʽ
								//��౶������
								CString space;
								space.Format("%0.2f",(_sl-240)/240);
								m_RtfDoc.m_RtfParagraph.GetAt(m_RtfDoc.m_RtfParagraph.GetCount()-1)->m_Space=space;//��౶��

								//�жϽ�ȡ����ַ�
								int index=bufstr.Find(" ");
								if(index != -1)
								{
									CString oldbuf=bufstr.Mid(index+1,bufstr.GetLength()-index+1);
									for(int k=0;k<oldbuf.GetLength();k++)
									{
										p_rtftext=new CRtfText();
										p_rtftext->m_Text=oldbuf.Mid(k,1);//�ı�
										if(m_Fonts.GetCount()>0)
											p_rtftext->m_FontName=m_Fonts.GetAt(_f);//����
										if(m_Colors.GetCount()>0)
											p_rtftext->m_color=m_Colors.GetAt(_cf-1);//��ɫ
										p_rtftext->m_Size=_fs;//�ֺ�
										p_rtftext->m_Bold=_b;//����
										p_rtftext->m_Italic=_i;//б��
										p_rtftext->m_Underline=_ul;//�»���

										int rtftextcount=m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.GetCount();
										m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.InsertAt(rtftextcount,p_rtftext);//�����
									}
								}
							}
							else
							{
								bufstr=d_bufstr;
								p_rtfparagraph=new CRtfParagraph();
								m_RtfDoc.m_RtfParagraph.InsertAt(m_RtfDoc.m_RtfParagraph.GetCount(),p_rtfparagraph);
								m_RtfDoc.m_RtfParagraph.GetAt(m_RtfDoc.m_RtfParagraph.GetCount()-1)->m_Alignment=_horizontal;//���뷽ʽ
								//��౶������
								CString space;
								space.Format("%0.2f",(_sl-240)/240);
								m_RtfDoc.m_RtfParagraph.GetAt(m_RtfDoc.m_RtfParagraph.GetCount()-1)->m_Space=space;//��౶��

								//�жϽ�ȡ����ַ�
								if(bufstr.GetLength()>5)
								{
									CString oldbuf=bufstr.Mid(5,bufstr.GetLength()-5);
									for(int k=0;k<oldbuf.GetLength();k++)
									{
										p_rtftext=new CRtfText();
										p_rtftext->m_Text=oldbuf.Mid(k,1);//�ı�
										if(m_Fonts.GetCount()>0)
											p_rtftext->m_FontName=m_Fonts.GetAt(_f);//����
										if(m_Colors.GetCount()>0)
											p_rtftext->m_color=m_Colors.GetAt(_cf-1);//��ɫ
										p_rtftext->m_Size=_fs;//�ֺ�
										p_rtftext->m_Bold=_b;//����
										p_rtftext->m_Italic=_i;//б��
										p_rtftext->m_Underline=_ul;//�»���

										int rtftextcount=m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.GetCount();
										m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.InsertAt(rtftextcount,p_rtftext);//�����
									}
								}

							}
						}
						else if(d_bufstr.Left(2).Find("qc")!=-1)//���ж���
						{
							bufstr=d_bufstr;
							_horizontal=1;
							m_RtfDoc.m_RtfParagraph.GetAt(m_RtfDoc.m_RtfParagraph.GetCount()-1)->m_Alignment=_horizontal;//���뷽ʽ

							//�жϽ�ȡ����ַ�
							int index=bufstr.Find(" ");
							if(index != -1)
							{
								CString oldbuf=bufstr.Mid(index+1,bufstr.GetLength()-index+1);
								for(int k=0;k<oldbuf.GetLength();k++)
								{
									p_rtftext=new CRtfText();
									p_rtftext->m_Text=oldbuf.Mid(k,1);//�ı�
									if(m_Fonts.GetCount()>0)
										p_rtftext->m_FontName=m_Fonts.GetAt(_f);//����
									if(m_Colors.GetCount()>0)
										p_rtftext->m_color=m_Colors.GetAt(_cf-1);//��ɫ
									p_rtftext->m_Size=_fs;//�ֺ�
									p_rtftext->m_Bold=_b;//����
									p_rtftext->m_Italic=_i;//б��
									p_rtftext->m_Underline=_ul;//�»���

									int rtftextcount=m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.GetCount();
									m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.InsertAt(rtftextcount,p_rtftext);//�����
								}
							}
						}
						else if(d_bufstr.Left(2).Find("qr")!=-1)//�Ҷ���
						{
							bufstr=d_bufstr;
							_horizontal=2;
							m_RtfDoc.m_RtfParagraph.GetAt(m_RtfDoc.m_RtfParagraph.GetCount()-1)->m_Alignment=_horizontal;//���뷽ʽ

							//�жϽ�ȡ����ַ�
							int index=bufstr.Find(" ");
							if(index != -1)
							{
								CString oldbuf=bufstr.Mid(index+1,bufstr.GetLength()-index+1);
								for(int k=0;k<oldbuf.GetLength();k++)
								{
									p_rtftext=new CRtfText();
									p_rtftext->m_Text=oldbuf.Mid(k,1);//�ı�
									if(m_Fonts.GetCount()>0)
										p_rtftext->m_FontName=m_Fonts.GetAt(_f);//����
									if(m_Colors.GetCount()>0)
										p_rtftext->m_color=m_Colors.GetAt(_cf-1);//��ɫ
									p_rtftext->m_Size=_fs;//�ֺ�
									p_rtftext->m_Bold=_b;//����
									p_rtftext->m_Italic=_i;//б��
									p_rtftext->m_Underline=_ul;//�»���

									int rtftextcount=m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.GetCount();
									m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.InsertAt(rtftextcount,p_rtftext);//�����
								}
							}
						}
						else if(d_bufstr.Left(2).Find("sl")!=-1)//�����о�sl����ֵ
						{
							if(d_bufstr.Left(6).Find("slmult")!=-1)
							{
								bufstr=d_bufstr;
								//�жϽ�ȡ����ַ�
								int index=bufstr.Find(" ");
								if(index != -1)
								{
									CString oldbuf=bufstr.Mid(index+1,bufstr.GetLength()-index+1);
									for(int k=0;k<oldbuf.GetLength();k++)
									{
										p_rtftext=new CRtfText();
										p_rtftext->m_Text=oldbuf.Mid(k,1);//�ı�
										if(m_Fonts.GetCount()>0)
											p_rtftext->m_FontName=m_Fonts.GetAt(_f);//����
										if(m_Colors.GetCount()>0)
											p_rtftext->m_color=m_Colors.GetAt(_cf-1);//��ɫ
										p_rtftext->m_Size=_fs;//�ֺ�
										p_rtftext->m_Bold=_b;//����
										p_rtftext->m_Italic=_i;//б��
										p_rtftext->m_Underline=_ul;//�»���

										int rtftextcount=m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.GetCount();
										m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.InsertAt(rtftextcount,p_rtftext);//�����
									}
								}

							}
							else
							{
								bufstr=d_bufstr;
								bufstr.Replace("sl","");
								_sl=::atoi(bufstr);

								//��౶������
								CString space;
								space.Format("%0.2f",(_sl-240)/240);
								m_RtfDoc.m_RtfParagraph.GetAt(m_RtfDoc.m_RtfParagraph.GetCount()-1)->m_Space=space;//��౶��

								//�жϽ�ȡ����ַ�
								int index=bufstr.Find(" ");
								if(index != -1)
								{
									CString oldbuf=bufstr.Mid(index+1,bufstr.GetLength()-index+1);
									for(int k=0;k<oldbuf.GetLength();k++)
									{
										p_rtftext=new CRtfText();
										p_rtftext->m_Text=oldbuf.Mid(k,1);//�ı�
										if(m_Fonts.GetCount()>0)
											p_rtftext->m_FontName=m_Fonts.GetAt(_f);//����
										if(m_Colors.GetCount()>0)
											p_rtftext->m_color=m_Colors.GetAt(_cf-1);//��ɫ
										p_rtftext->m_Size=_fs;//�ֺ�
										p_rtftext->m_Bold=_b;//����
										p_rtftext->m_Italic=_i;//б��
										p_rtftext->m_Underline=_ul;//�»���

										int rtftextcount=m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.GetCount();
										m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.InsertAt(rtftextcount,p_rtftext);//�����
									}
								}
							}
						}
						else if(d_bufstr.Left(2).Find("xx")!=-1)//����ַ�"\" ��RTF�ļ��б����ʽΪ"\\"
						{
							CString oldbuf="";
							if(d_bufstr.GetLength()>2)
							{
								oldbuf=d_bufstr.Mid(2,d_bufstr.GetLength()-2);
							}
							p_rtftext=new CRtfText();
							p_rtftext->m_Text="\\";//�ı�
							if(m_Fonts.GetCount()>0)
								p_rtftext->m_FontName=m_Fonts.GetAt(_f);//����
							if(m_Colors.GetCount()>0)
								p_rtftext->m_color=m_Colors.GetAt(_cf-1);//��ɫ
							p_rtftext->m_Size=_fs;//�ֺ�
							p_rtftext->m_Bold=_b;//����
							p_rtftext->m_Italic=_i;//б��
							p_rtftext->m_Underline=_ul;//�»���

							int rtftextcount=m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.GetCount();
							m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.InsertAt(rtftextcount,p_rtftext);//�����

							if(oldbuf.GetLength()>0)
							{
								for(int k=0;k<oldbuf.GetLength();k++)
								{
									p_rtftext=new CRtfText();
									p_rtftext->m_Text=oldbuf.Mid(k,1);//�ı�
									if(m_Fonts.GetCount()>0)
										p_rtftext->m_FontName=m_Fonts.GetAt(_f);//����
									if(m_Colors.GetCount()>0)
										p_rtftext->m_color=m_Colors.GetAt(_cf-1);//��ɫ
									p_rtftext->m_Size=_fs;//�ֺ�
									p_rtftext->m_Bold=_b;//����
									p_rtftext->m_Italic=_i;//б��
									p_rtftext->m_Underline=_ul;//�»���

									int rtftextcount=m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.GetCount();
									m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.InsertAt(rtftextcount,p_rtftext);//�����
								}
							}
						}
						else if(d_bufstr.Left(2).Find("[[")!=-1)//����ַ�"{" ��RTF�ļ��б����ʽΪ"[["
						{
							CString oldbuf="";
							if(d_bufstr.GetLength()>2)
							{
								oldbuf=d_bufstr.Mid(2,d_bufstr.GetLength()-2);
							}
							p_rtftext=new CRtfText();
							p_rtftext->m_Text="{";//�ı�
							if(m_Fonts.GetCount()>0)
								p_rtftext->m_FontName=m_Fonts.GetAt(_f);//����
							if(m_Colors.GetCount()>0)
								p_rtftext->m_color=m_Colors.GetAt(_cf-1);//��ɫ
							p_rtftext->m_Size=_fs;//�ֺ�
							p_rtftext->m_Bold=_b;//����
							p_rtftext->m_Italic=_i;//б��
							p_rtftext->m_Underline=_ul;//�»���

							int rtftextcount=m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.GetCount();
							m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.InsertAt(rtftextcount,p_rtftext);//�����

							if(oldbuf.GetLength()>0)
							{
								for(int k=0;k<oldbuf.GetLength();k++)
								{
									p_rtftext=new CRtfText();
									p_rtftext->m_Text=oldbuf.Mid(k,1);//�ı�
									if(m_Fonts.GetCount()>0)
										p_rtftext->m_FontName=m_Fonts.GetAt(_f);//����
									if(m_Colors.GetCount()>0)
										p_rtftext->m_color=m_Colors.GetAt(_cf-1);//��ɫ
									p_rtftext->m_Size=_fs;//�ֺ�
									p_rtftext->m_Bold=_b;//����
									p_rtftext->m_Italic=_i;//б��
									p_rtftext->m_Underline=_ul;//�»���

									int rtftextcount=m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.GetCount();
									m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.InsertAt(rtftextcount,p_rtftext);//�����
								}
							}
						}
						else if(d_bufstr.Left(2).Find("]]")!=-1)//����ַ�"}" ��RTF�ļ��б����ʽΪ"]]"
						{
							CString oldbuf="";
							if(d_bufstr.GetLength()>2)
							{
								oldbuf=d_bufstr.Mid(2,d_bufstr.GetLength()-2);
							}
							p_rtftext=new CRtfText();
							p_rtftext->m_Text="}";//�ı�
							if(m_Fonts.GetCount()>0)
								p_rtftext->m_FontName=m_Fonts.GetAt(_f);//����
							if(m_Colors.GetCount()>0)
								p_rtftext->m_color=m_Colors.GetAt(_cf-1);//��ɫ
							p_rtftext->m_Size=_fs;//�ֺ�
							p_rtftext->m_Bold=_b;//����
							p_rtftext->m_Italic=_i;//б��
							p_rtftext->m_Underline=_ul;//�»���

							int rtftextcount=m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.GetCount();
							m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.InsertAt(rtftextcount,p_rtftext);//�����

							if(oldbuf.GetLength()>0)
							{
								for(int k=0;k<oldbuf.GetLength();k++)
								{
									p_rtftext=new CRtfText();
									p_rtftext->m_Text=oldbuf.Mid(k,1);//�ı�
									if(m_Fonts.GetCount()>0)
										p_rtftext->m_FontName=m_Fonts.GetAt(_f);//����
									if(m_Colors.GetCount()>0)
										p_rtftext->m_color=m_Colors.GetAt(_cf-1);//��ɫ
									p_rtftext->m_Size=_fs;//�ֺ�
									p_rtftext->m_Bold=_b;//����
									p_rtftext->m_Italic=_i;//б��
									p_rtftext->m_Underline=_ul;//�»���

									int rtftextcount=m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.GetCount();
									m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.InsertAt(rtftextcount,p_rtftext);//�����
								}
							}
						}
						else if(d_bufstr.Left(9).Find("ldblquote")!=-1)//ȫ��"��"�ַ�
						{
							CString oldbuf="";
							if(d_bufstr.GetLength()>9)
							{
								oldbuf=d_bufstr.Mid(9,d_bufstr.GetLength()-9);
							}
							p_rtftext=new CRtfText();
							p_rtftext->m_Text="��";//�ı�
							if(m_Fonts.GetCount()>0)
								p_rtftext->m_FontName=m_Fonts.GetAt(_f);//����
							if(m_Colors.GetCount()>0)
								p_rtftext->m_color=m_Colors.GetAt(_cf-1);//��ɫ
							p_rtftext->m_Size=_fs;//�ֺ�
							p_rtftext->m_Bold=_b;//����
							p_rtftext->m_Italic=_i;//б��
							p_rtftext->m_Underline=_ul;//�»���

							int rtftextcount=m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.GetCount();
							m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.InsertAt(rtftextcount,p_rtftext);//�����

							if(oldbuf.GetLength()>0)
							{
								for(int k=0;k<oldbuf.GetLength();k++)
								{
									p_rtftext=new CRtfText();
									p_rtftext->m_Text=oldbuf.Mid(k,1);//�ı�
									if(m_Fonts.GetCount()>0)
										p_rtftext->m_FontName=m_Fonts.GetAt(_f);//����
									if(m_Colors.GetCount()>0)
										p_rtftext->m_color=m_Colors.GetAt(_cf-1);//��ɫ
									p_rtftext->m_Size=_fs;//�ֺ�
									p_rtftext->m_Bold=_b;//����
									p_rtftext->m_Italic=_i;//б��
									p_rtftext->m_Underline=_ul;//�»���

									int rtftextcount=m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.GetCount();
									m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.InsertAt(rtftextcount,p_rtftext);//�����
								}
							}
						}
						else if(d_bufstr.Left(9).Find("rdblquote")!=-1)//ȫ��"��"�ַ�
						{
							CString oldbuf="";
							if(d_bufstr.GetLength()>9)
							{
								oldbuf=d_bufstr.Mid(9,d_bufstr.GetLength()-9);
							}
							p_rtftext=new CRtfText();
							p_rtftext->m_Text="��";//�ı�
							if(m_Fonts.GetCount()>0)
								p_rtftext->m_FontName=m_Fonts.GetAt(_f);//����
							if(m_Colors.GetCount()>0)
								p_rtftext->m_color=m_Colors.GetAt(_cf-1);//��ɫ
							p_rtftext->m_Size=_fs;//�ֺ�
							p_rtftext->m_Bold=_b;//����
							p_rtftext->m_Italic=_i;//б��
							p_rtftext->m_Underline=_ul;//�»���

							int rtftextcount=m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.GetCount();
							m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.InsertAt(rtftextcount,p_rtftext);//�����

							if(oldbuf.GetLength()>0)
							{
								for(int k=0;k<oldbuf.GetLength();k++)
								{
									p_rtftext=new CRtfText();
									p_rtftext->m_Text=oldbuf.Mid(k,1);//�ı�
									if(m_Fonts.GetCount()>0)
										p_rtftext->m_FontName=m_Fonts.GetAt(_f);//����
									if(m_Colors.GetCount()>0)
										p_rtftext->m_color=m_Colors.GetAt(_cf-1);//��ɫ
									p_rtftext->m_Size=_fs;//�ֺ�
									p_rtftext->m_Bold=_b;//����
									p_rtftext->m_Italic=_i;//б��
									p_rtftext->m_Underline=_ul;//�»���

									int rtftextcount=m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.GetCount();
									m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.InsertAt(rtftextcount,p_rtftext);//�����
								}
							}
						}
						else if(d_bufstr.Left(6).Find("lquote")!=-1)//ȫ��"��"�ַ�
						{
							CString oldbuf="";
							if(d_bufstr.GetLength()>6)
							{
								oldbuf=d_bufstr.Mid(6,d_bufstr.GetLength()-6);
							}
							p_rtftext=new CRtfText();
							p_rtftext->m_Text="��";//�ı�
							if(m_Fonts.GetCount()>0)
								p_rtftext->m_FontName=m_Fonts.GetAt(_f);//����
							if(m_Colors.GetCount()>0)
								p_rtftext->m_color=m_Colors.GetAt(_cf-1);//��ɫ
							p_rtftext->m_Size=_fs;//�ֺ�
							p_rtftext->m_Bold=_b;//����
							p_rtftext->m_Italic=_i;//б��
							p_rtftext->m_Underline=_ul;//�»���

							int rtftextcount=m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.GetCount();
							m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.InsertAt(rtftextcount,p_rtftext);//�����

							if(oldbuf.GetLength()>0)
							{
								for(int k=0;k<oldbuf.GetLength();k++)
								{
									p_rtftext=new CRtfText();
									p_rtftext->m_Text=oldbuf.Mid(k,1);//�ı�
									if(m_Fonts.GetCount()>0)
										p_rtftext->m_FontName=m_Fonts.GetAt(_f);//����
									if(m_Colors.GetCount()>0)
										p_rtftext->m_color=m_Colors.GetAt(_cf-1);//��ɫ
									p_rtftext->m_Size=_fs;//�ֺ�
									p_rtftext->m_Bold=_b;//����
									p_rtftext->m_Italic=_i;//б��
									p_rtftext->m_Underline=_ul;//�»���

									int rtftextcount=m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.GetCount();
									m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.InsertAt(rtftextcount,p_rtftext);//�����
								}
							}
						}
						else if(d_bufstr.Left(6).Find("rquote")!=-1)//ȫ��"��"�ַ�
						{
							CString oldbuf="";
							if(d_bufstr.GetLength()>6)
							{
								oldbuf=d_bufstr.Mid(6,d_bufstr.GetLength()-6);
							}
							p_rtftext=new CRtfText();
							p_rtftext->m_Text="��";//�ı�
							if(m_Fonts.GetCount()>0)
								p_rtftext->m_FontName=m_Fonts.GetAt(_f);//����
							if(m_Colors.GetCount()>0)
								p_rtftext->m_color=m_Colors.GetAt(_cf-1);//��ɫ
							p_rtftext->m_Size=_fs;//�ֺ�
							p_rtftext->m_Bold=_b;//����
							p_rtftext->m_Italic=_i;//б��
							p_rtftext->m_Underline=_ul;//�»���

							int rtftextcount=m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.GetCount();
							m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.InsertAt(rtftextcount,p_rtftext);//�����

							if(oldbuf.GetLength()>0)
							{
								for(int k=0;k<oldbuf.GetLength();k++)
								{
									p_rtftext=new CRtfText();
									p_rtftext->m_Text=oldbuf.Mid(k,1);//�ı�
									if(m_Fonts.GetCount()>0)
										p_rtftext->m_FontName=m_Fonts.GetAt(_f);//����
									if(m_Colors.GetCount()>0)
										p_rtftext->m_color=m_Colors.GetAt(_cf-1);//��ɫ
									p_rtftext->m_Size=_fs;//�ֺ�
									p_rtftext->m_Bold=_b;//����
									p_rtftext->m_Italic=_i;//б��
									p_rtftext->m_Underline=_ul;//�»���

									int rtftextcount=m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.GetCount();
									m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.InsertAt(rtftextcount,p_rtftext);//�����
								}
							}
						}
						else if(d_bufstr.Left(6).Find("emdash")!=-1)//ȫ��"��"�ַ�
						{
							CString oldbuf="";
							if(d_bufstr.GetLength()>6)
							{
								oldbuf=d_bufstr.Mid(6,d_bufstr.GetLength()-6);
							}
							p_rtftext=new CRtfText();
							p_rtftext->m_Text="��";//�ı�
							if(m_Fonts.GetCount()>0)
								p_rtftext->m_FontName=m_Fonts.GetAt(_f);//����
							if(m_Colors.GetCount()>0)
								p_rtftext->m_color=m_Colors.GetAt(_cf-1);//��ɫ
							p_rtftext->m_Size=_fs;//�ֺ�
							p_rtftext->m_Bold=_b;//����
							p_rtftext->m_Italic=_i;//б��
							p_rtftext->m_Underline=_ul;//�»���

							int rtftextcount=m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.GetCount();
							m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.InsertAt(rtftextcount,p_rtftext);//�����

							if(oldbuf.GetLength()>0)
							{
								for(int k=0;k<oldbuf.GetLength();k++)
								{
									p_rtftext=new CRtfText();
									p_rtftext->m_Text=oldbuf.Mid(k,1);//�ı�
									if(m_Fonts.GetCount()>0)
										p_rtftext->m_FontName=m_Fonts.GetAt(_f);//����
									if(m_Colors.GetCount()>0)
										p_rtftext->m_color=m_Colors.GetAt(_cf-1);//��ɫ
									p_rtftext->m_Size=_fs;//�ֺ�
									p_rtftext->m_Bold=_b;//����
									p_rtftext->m_Italic=_i;//б��
									p_rtftext->m_Underline=_ul;//�»���

									int rtftextcount=m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.GetCount();
									m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.InsertAt(rtftextcount,p_rtftext);//�����
								}
							}
						}
						else if(d_bufstr.Left(1).Find("'")!=-1)//������λ������
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

									oldbuf=bufstr.Mid(2,bufstr.GetLength()-2);//�洢����ַ�
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
								//���������
								p_rtftext=new CRtfText();
								p_rtftext->m_Text=textbuf;//�ı�
								if(m_Fonts.GetCount()>0)
									p_rtftext->m_FontName=m_Fonts.GetAt(_f);//����
								if(m_Colors.GetCount()>0)
									p_rtftext->m_color=m_Colors.GetAt(_cf-1);//��ɫ
								p_rtftext->m_Size=_fs;//�ֺ�
								p_rtftext->m_Bold=_b;//����
								p_rtftext->m_Italic=_i;//б��
								p_rtftext->m_Underline=_ul;//�»���

								int rtftextcount=m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.GetCount();
								m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.InsertAt(rtftextcount,p_rtftext);//�����

								if(oldbuf.GetLength()>0)
								{
									for(int k=0;k<oldbuf.GetLength();k++)
									{
										p_rtftext=new CRtfText();
										p_rtftext->m_Text=oldbuf.Mid(k,1);//�ı�
										if(m_Fonts.GetCount()>0)
											p_rtftext->m_FontName=m_Fonts.GetAt(_f);//����
										if(m_Colors.GetCount()>0)
											p_rtftext->m_color=m_Colors.GetAt(_cf-1);//��ɫ
										p_rtftext->m_Size=_fs;//�ֺ�
										p_rtftext->m_Bold=_b;//����
										p_rtftext->m_Italic=_i;//б��
										p_rtftext->m_Underline=_ul;//�»���

										int rtftextcount=m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.GetCount();
										m_RtfDoc.m_RtfParagraph[m_RtfDoc.m_RtfParagraph.GetCount()-1]->m_RtfText.InsertAt(rtftextcount,p_rtftext);//�����
									}
								}

								textbuf="";//����ı�����������
							}

						}

					}
				}
			}
			else
			{
				////�ڴ��ͷ�
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

CString CMyRtf::LocationToChinase(CString location_str)//��λ���ַ�ת�����ַ�
{
	//�Ͽ�ÿ����λ��
	int Start=0,End=0,start_old=0;
	CString lstr,ret_str="";
	CArray<CString,CString> m_liststr;

	int _index=0;//��¼�� �ж��Ƿ���Ҫ������λ�� _index<=0����Ҫ������λ��
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

	//������λ��
	CString buf_str="";//û�б����ַ����黺����
	CString head_buf;
	int ch_num=0;//�ַ�ʮ���Ʊ���ֵ
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
				int h_num=this->HexsToDec(numstr,16);//16���� ʮλ�ַ�
				numstr=asiistr.Mid(1,1);
				int l_num=this->HexsToDec(numstr,1);//16���� ��λ�ַ�
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

//16�����ַ�ת����
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

//�ַ�תRGB��ɫֵ
COLORREF CMyRtf::AsiiToColor(CString rgbstr)
{
	CString colstr;
	int red=255,green=255,blue=255;
	int Start=0,End=0,start_old;
	Start=rgbstr.Find("\\",End);
	start_old=End=Start+1;
	if(Start!=-1)
	{
		//--red ��ɫ
		Start=rgbstr.Find("\\",End);
		if(Start!=-1)
		{
			colstr=rgbstr.Mid(start_old,Start-start_old);
			start_old=End=Start+1;

			colstr.Replace("red","");
			red=::atoi(colstr);
		}
		//--

		//--green ��ɫ
		Start=rgbstr.Find("\\",End);
		if(Start!=-1)
		{
			colstr=rgbstr.Mid(start_old,Start-start_old);
			start_old=End=Start+1;

			colstr.Replace("green","");
			green=::atoi(colstr);
		}
		//--

		//--blue ��ɫ
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
