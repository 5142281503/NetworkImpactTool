// NIPDlg.cpp : implementation file
//

#include "stdafx.h"
#include "NIP.h"
#include "NIPDlg.h"
#include <shlwapi.h>	// PathFileExists
#include <fstream>		// ifstream
#include <sstream>		// wostringstream
#include <wininet.h>	// INTERNET_CACHE_ENTRY_INFO
#include <atlbase.h>	// AtlWaitWithMessageLoop

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

namespace
{
	void MsgWait( unsigned int in_msTime )
	{
		DWORD msEnter = GetTickCount();
		long msRest = in_msTime;
		while( msRest > 0 )
		{
			DWORD dwRet = ::MsgWaitForMultipleObjectsEx( 0, 0, msRest, 
				QS_ALLINPUT | QS_ALLPOSTMESSAGE,
				MWMO_INPUTAVAILABLE );

			if( dwRet == WAIT_OBJECT_0 )
			{
				MSG msg;
				// There is one or more window message available. Dispatch them
				while(PeekMessage(&msg,NULL,NULL,NULL,PM_REMOVE))
				{
					TranslateMessage(&msg);
					DispatchMessage(&msg);
				}

				msRest = in_msTime - (GetTickCount() - msEnter);
			}
			else
			{
				break;
			}
		}
	}
}


#define WM_PAGE_LOADED WM_APP+69

class CAboutDlg : public CDialog
{
public:
	CAboutDlg();

// Dialog Data
	enum { IDD = IDD_ABOUTBOX };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

// Implementation
protected:
	DECLARE_MESSAGE_MAP()
};

CAboutDlg::CAboutDlg() : CDialog(CAboutDlg::IDD)
{
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialog)
END_MESSAGE_MAP()

//http://support.microsoft.com/kb/q180366/
BEGIN_EVENTSINK_MAP(CNIPDlg, CDialog)
   ON_EVENT(CNIPDlg, IDC_EXPLORER1, 259 /* DocumentComplete */, OnDocumentComplete, VTS_DISPATCH VTS_PVARIANT)
END_EVENTSINK_MAP()




CNIPDlg::CNIPDlg(CWnd* pParent /*=NULL*/)
	: CDialog(CNIPDlg::IDD, pParent)
	, m_eventHandle ( CreateEvent( NULL, false /*manual reset*/, false /*initial state*/, NULL ) )
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CNIPDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_EDIT1, m_CEditFileName);
	DDX_Control(pDX, IDC_EDIT2, m_CEditInfo);
	DDX_Control(pDX, IDC_EXPLORER1, m_browser);
}

BEGIN_MESSAGE_MAP(CNIPDlg, CDialog)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	//}}AFX_MSG_MAP
	ON_BN_CLICKED(IDOK, &CNIPDlg::OnBnClickedOk)
	ON_MESSAGE( WM_PAGE_LOADED, OnPageLoaded )
	ON_BN_CLICKED(IDCANCEL, &CNIPDlg::OnBnClickedCancel)
END_MESSAGE_MAP()


// CNIPDlg message handlers

BOOL CNIPDlg::OnInitDialog()
{
	CDialog::OnInitDialog();

	// Add "About..." menu item to system menu.

	// IDM_ABOUTBOX must be in the system command range.
	ASSERT((IDM_ABOUTBOX & 0xFFF0) == IDM_ABOUTBOX);
	ASSERT(IDM_ABOUTBOX < 0xF000);

	CMenu* pSysMenu = GetSystemMenu(FALSE);
	if (pSysMenu != NULL)
	{
		CString strAboutMenu;
		strAboutMenu.LoadString(IDS_ABOUTBOX);
		if (!strAboutMenu.IsEmpty())
		{
			pSysMenu->AppendMenu(MF_SEPARATOR);
			pSysMenu->AppendMenu(MF_STRING, IDM_ABOUTBOX, strAboutMenu);
		}
	}

	// Set the icon for this dialog.  The framework does this automatically
	//  when the application's main window is not a dialog
	SetIcon(m_hIcon, TRUE);			// Set big icon
	SetIcon(m_hIcon, FALSE);		// Set small icon

	// TODO: Add extra initialization here
	GetLastErrorMessage(L"This is free only for non-commercial use only.\r\nIf you are at a company or using this tool for a company please do not do so.\r\n\r\nDeveloped and tested under XPSP3 and VistaSp1.\r\nI cannot be held liable or responsible for anything breaking on a machine running this app.");

	m_CEditFileName.SetWindowTextW(L"domains.txt");

	return TRUE;  // return TRUE  unless you set the focus to a control
}

void CNIPDlg::OnSysCommand(UINT nID, LPARAM lParam)
{
	if ((nID & 0xFFF0) == IDM_ABOUTBOX)
	{
		CAboutDlg dlgAbout;
		dlgAbout.DoModal();
	}
	else
	{
		CDialog::OnSysCommand(nID, lParam);
	}
}

// If you add a minimize button to your dialog, you will need the code below
//  to draw the icon.  For MFC applications using the document/view model,
//  this is automatically done for you by the framework.

void CNIPDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // device context for painting

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// Center icon in client rectangle
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// Draw the icon
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialog::OnPaint();
	}
}

// The system calls this function to obtain the cursor to display while the user drags
//  the minimized window.
HCURSOR CNIPDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}


void CNIPDlg::OnBnClickedOk()
{
	CString cstr;
	m_CEditFileName.GetWindowTextW(cstr);

	if(cstr.IsEmpty())
	{
		AfxMessageBox(L"Please specify filename (domains.txt can be used as an example) containing domains", MB_ICONERROR);
		GetLastErrorMessage(L"This is the file I will use as input to load domains");
		return;
	}	

	 if(!PathFileExists(cstr))
	 {
		AfxMessageBox(L"Sorry can't find this file", MB_ICONERROR);
		GetLastErrorMessage(L"Make sure your path is valid");
		return;
	 }

	 CWaitCursor c;
	 GetLastErrorMessage(L"Reading file contents...");

	std::wifstream inputFile(cstr);
	std::wstring fileData((std::istreambuf_iterator<wchar_t>(inputFile)), std::istreambuf_iterator<wchar_t>());

	GetLastErrorMessage(L"Parsing file contents...");

	std::vector<std::wstring> vec;
	SplitOnString(fileData, vec, L"\n");

	WSADATA wsaData = { 0 };;
	if(0 != WSAStartup(MAKEWORD(2, 2), &wsaData))
    {        
		GetLastErrorMessage(L"Unable to initialize winsock...");
        return ;
    }

	std::vector<std::wstring> good_vec;

	for each (std::wstring s in vec)
	{
		TRACE1("we have read [%s] (ignoring lines starting with ':', those are comments)\r\n", s.c_str());		

		//////////////// Skip comments ///////////////////////////////////////////////////////

		const std::wstring::size_type i = s.find_first_of(L':');

		if( (i != std::string::npos) && (0 == i))
		{
			CString Msg(L"Skipping comment ");
			Msg += s.c_str();
			GetLastErrorMessage(Msg);
			continue;
		}

		//////////////// Resolve DNS ///////////////////////////////////////////////////////

		CString Msg(L"Resolving DNS for ...");
		Msg += s.c_str();
		GetLastErrorMessage(Msg);		

		if(!gethostbyname(WStringToString(s).c_str()))
		{
			CString Msg(L"Skipping, Failed DNS lookup for ");
			Msg += s.c_str();
			GetLastErrorMessage(Msg);
			continue;
		}
		else
		{
			GetLastErrorMessage(L"ok");
		}

		good_vec.push_back(s);
	}


	std::wostringstream oss;
	oss << L"We are left with " <<  good_vec.size() << L" domains to download";
	GetLastErrorMessage(oss.str().c_str());

	if(!good_vec.empty())
	{
		GetLastErrorMessage(L"Clearing IE cache including cookies for consistent results");
		//////////////// Clear IE cache including cookies ///////////////////////////////////////////////////////
		CleearIECache();		

		//////////////// Begin timer ///////////////////////////////////////////////////////
		CTime m_startTime ( CTime::GetCurrentTime() );

		for each (std::wstring s in good_vec)
		{
			oss.rdbuf()->str(L"");
			oss << L"Navigating to " << s << L" writing to local cache, but not reading local cache. Waiting for it to load as this better reflects typical user habits.";
			GetLastErrorMessage(oss.str().c_str());

			//////////////// IE Navigate to the site ///////////////////////////////////////////////////////
			VARIANTARG vWorkaround;
			VariantInit(&vWorkaround);
			vWorkaround.vt = VT_I4;
			vWorkaround.lVal = navNoReadFromCache;			

			m_browser.Navigate(s.c_str(), &vWorkaround, NULL, NULL, NULL);

			ATL::AtlWaitWithMessageLoop(m_eventHandle);
			
		}


		//////////////// End timer ///////////////////////////////////////////////////////
		CTimeSpan cts ( CTime::GetCurrentTime() - m_startTime );

		oss.rdbuf()->str(L"");
		oss << L"Network Impact Results are (LOWER is BETTER): hours: " << cts.GetHours() << L" Minutes: " << cts.GetMinutes() << L" Secs: " << cts.GetSeconds();
		
		AfxMessageBox(oss.str().c_str());

	}
	
	WSACleanup();

	//OnOK();
}


void 
CNIPDlg::GetLastErrorMessage(const CString& in_Msg)
{
	LPVOID lpMsgBuf = NULL;	

	if( FormatMessage(  FORMAT_MESSAGE_ALLOCATE_BUFFER |
						FORMAT_MESSAGE_FROM_SYSTEM |
						FORMAT_MESSAGE_IGNORE_INSERTS,
						NULL,
						GetLastError(),
						MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT),
						(LPTSTR) &lpMsgBuf,
						0,
						NULL ))
	{
		CString x(in_Msg);
		x += L"\r\n\r\n";
		x += (LPWSTR)lpMsgBuf;

		MsgWait(1000);

		m_CEditInfo.SetWindowTextW(x);		
		UpdateWindow();

		if(lpMsgBuf) LocalFree( lpMsgBuf );
	}
}


void
CNIPDlg::SplitOnString(const std::wstring& in_str, std::vector<std::wstring>& out_vector,const std::wstring& in_strSplit/* = ";"*/)
{
	if (in_strSplit.empty())
	{
		out_vector.push_back(in_str);
		return;
	}

	size_t i = 0;
	size_t j = 0;
	do{
		j = in_str.find(in_strSplit, i);
		if(j == std::wstring::npos)
		{
			j = in_str.length();
		}
		else if (j == i)
		{
			j++;
			i++;
			continue;
		}

		const std::wstring strTemp(in_str.substr(i, j - i));
		if(strTemp.length()){
			out_vector.push_back(strTemp);
		}
		i = j + in_strSplit.length();
	}while(j < in_str.length());
}


std::string 
CNIPDlg::WStringToString( const std::wstring& in_strW )
{
	USES_CONVERSION;
	
	//return W2CT( in_strW.c_str() );
	return W2A( in_strW.c_str() );
}



bool CNIPDlg::CleearIECache( )
{
	bool bResult = false;
    bool bDone = false;
    LPINTERNET_CACHE_ENTRY_INFO lpCacheEntry = NULL;  
 
    DWORD  dwEntrySize = 4096; // start buffer size    
    HANDLE hCacheDir = NULL;    
    DWORD  dwError = ERROR_INSUFFICIENT_BUFFER;
	BOOL bSuccess = FALSE;
    
    do 
    {                               
        switch (dwError)
        {
            // need a bigger buffer
            case ERROR_INSUFFICIENT_BUFFER: 
                if(lpCacheEntry) delete [] lpCacheEntry ;
                lpCacheEntry = (LPINTERNET_CACHE_ENTRY_INFO) new char[dwEntrySize];
                lpCacheEntry->dwStructSize = dwEntrySize;
                bSuccess = FALSE;
                
				if (hCacheDir == NULL)                
				{
                    bSuccess = (hCacheDir = FindFirstUrlCacheEntry(NULL, lpCacheEntry,&dwEntrySize)) != NULL;
				}
                else
                    bSuccess = FindNextUrlCacheEntry(hCacheDir, lpCacheEntry, &dwEntrySize);

                if (bSuccess)
				{
                    dwError = ERROR_SUCCESS;    
				}
                else
                {
                    dwError = GetLastError();
                }
                break;

             // we are done
            case ERROR_NO_MORE_ITEMS:
                bDone = true;
                bResult = true;                
                break;

             // we have got an entry
            case ERROR_SUCCESS:                       
        
                DeleteUrlCacheEntry(lpCacheEntry->lpszSourceUrlName);
                    
                // get ready for next entry
                if (FindNextUrlCacheEntry(hCacheDir, lpCacheEntry, &dwEntrySize))
				{
                    dwError = ERROR_SUCCESS;          
				}     
                else
                {
                    dwError = GetLastError();
                }                    
                break;

            // unknown error
            default:
                bDone = true;                
                break;
        }

        if (bDone)
        {   
            if(lpCacheEntry) delete [] lpCacheEntry ;

            if (hCacheDir)
                FindCloseUrlCache(hCacheDir);         
                                  
        }
    } while (!bDone);
    return bResult;
}

//http://support.microsoft.com/kb/q180366/
void CNIPDlg::OnDocumentComplete(LPDISPATCH lpDisp, VARIANT FAR* URL)
{
   IUnknown*  pUnk;
   LPDISPATCH lpWBDisp;
   HRESULT    hr;

   pUnk = m_browser.GetControlUnknown();
   ASSERT(pUnk);

   hr = pUnk->QueryInterface(IID_IDispatch, (void**)&lpWBDisp);
   ASSERT(SUCCEEDED(hr));

   if (lpDisp == lpWBDisp )
   {
      // Top-level Window object, so document has been loaded
      TRACE("Web document is finished downloading\n");
	  GetLastErrorMessage(L"site loaded");

	  PostMessage( WM_PAGE_LOADED, reinterpret_cast<WPARAM>(m_eventHandle), 0 );
   }

  lpWBDisp->Release();
}

LRESULT CNIPDlg::OnPageLoaded(WPARAM in_eventHandle, LPARAM lParam )
{
	SetEvent(reinterpret_cast<HANDLE*>(in_eventHandle));
	return 0;
}
void CNIPDlg::OnBnClickedCancel()
{
	SetEvent(m_eventHandle); // just in case...

	OnCancel();
}

