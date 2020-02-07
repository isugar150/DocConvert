// FilePathCheckerModuleExample.cpp : Defines the entry point for the DLL application.
//

// Registry의 HKEY_CURRENT_USER\\Software\\HNC\\HwpCtrl\\Modules 키에
// HwpCtrl.RegisterModule("FilePathCheckDLL", "[모듈이름]") 으로 등록한 [모듈이름]의 값으로 
// 현재 DLL의 Full Path가 지정되어 있어야 동작함.


#include "stdafx.h"
#include <TCHAR.h>
#include <string.h>
#include <stdio.h>
#include <stdlib.h>
#include <assert.h>
#include "resource.h"
#include "FilePathCheckerModuleExample.h"

HINSTANCE THIS_MODULE_INSTANCE = NULL; // resource instance

BOOL g_bAskDlg = TRUE; // 대화상자를 띄울 지를 결정
INT_PTR g_bAskDlgResult = IDNO; // 대화상자의 결과를 저장

BOOL APIENTRY DllMain( HANDLE hModule, 
                       DWORD  ul_reason_for_call, 
                       LPVOID lpReserved
					 )
{
    switch (ul_reason_for_call)
	{
		case DLL_PROCESS_ATTACH:
			THIS_MODULE_INSTANCE = (HINSTANCE)hModule;
		case DLL_THREAD_ATTACH:
		case DLL_THREAD_DETACH:
		case DLL_PROCESS_DETACH:
			break;
    }
    return TRUE;
}

// 임시 폴더인지 검사.
BOOL IsTempPath(LPCTSTR szFilePath)
{
	static TCHAR szTempPath[MAX_PATH] = {0};
	static nTempPathLen = 0;

	if (!szTempPath[0]) // 한번만 실행한다.
	{
		GetTempPath(sizeof(szTempPath), szTempPath);
		nTempPathLen = _tcslen(szTempPath);
	}

	if (_tcsnicmp(szFilePath, szTempPath, nTempPathLen) == 0) {
		return TRUE;
	}
	return FALSE;		
}

// 인터넷 주소인지 검사
BOOL IsURLPath(LPCTSTR szFilePath)
{
	static LPCTSTR szURLPrefix = _T("HTTP://");
	static nURLPrefixLen = _tcslen(szURLPrefix);
		

	if (_tcsnicmp(szFilePath, szURLPrefix, nURLPrefixLen) == 0) {
		return TRUE;
	}
	return FALSE;		
}

BOOL CheckSiteInfo(LPCTSTR szSiteInfo)
{
	// SiteInfo 검사
	// 특정 사이트나 어플리케이션에서만 동작하도록 할 수 있다.
	// 일단은 무조건 통과 시킨다.
//	return TRUE;

	BOOL bRet = FALSE;
	BOOL bStopParsing = FALSE;
	if (szSiteInfo && szSiteInfo[0]) {
		LPTSTR szSiteInfoBuf = new TCHAR[_tcslen(szSiteInfo) + 8];
		LPTSTR szToken;
		LPTSTR szData;

		_tcscpy(szSiteInfoBuf, szSiteInfo);

		szToken = _tcstok(szSiteInfoBuf, _T("\r\n"));
		while(szToken) {
			if ((szData = _tcsstr(szToken, _T("URL:"))) == szToken) { // URL 검사
				// http://www.haansoft.com 로 시작하는 URL만 통과 시킨다.
				LPCTSTR szURL = szToken + _tcslen(_T("URL:"));
				LPCTSTR szAcceptedURL = _T("http://www.haansoft.com"); 

				if (_tcsnicmp(szURL, szAcceptedURL, _tcslen(szAcceptedURL)) == 0) {				
					bRet = TRUE;
				} else {
					bRet = FALSE;
				}
				bStopParsing = TRUE;
			}
			 if (bStopParsing)
				 break;// 한번이라도 실패하면 문자열 분석을 멈추고 실패로 간주한다.
			szToken = _tcstok(NULL, _T("\r\n"));
		}

		delete[] szSiteInfoBuf;
	}

	return bRet;
}
// modified from atlwin.h
BOOL CenterWindow(HWND hWnd, HWND hWndCenter)
{
	assert(::IsWindow(hWnd));
	
	// determine owner window to center against
	DWORD dwStyle = ::GetWindowLong(hWnd, GWL_STYLE);
	if(hWndCenter == NULL)
	{
		if(dwStyle & WS_CHILD)
			hWndCenter = ::GetParent(hWnd);
		else
			hWndCenter = ::GetWindow(hWnd, GW_OWNER);
	}
	
	// get coordinates of the window relative to its parent
	RECT rcDlg;
	::GetWindowRect(hWnd, &rcDlg);
	RECT rcArea;
	RECT rcCenter;
	HWND hWndParent;
	if(!(dwStyle & WS_CHILD))
	{
		// don't center against invisible or minimized windows
		if(hWndCenter != NULL)
		{
			DWORD dwStyleCenter = ::GetWindowLong(hWndCenter, GWL_STYLE);
			if(!(dwStyleCenter & WS_VISIBLE) || (dwStyleCenter & WS_MINIMIZE))
				hWndCenter = NULL;
		}
		
		// center within screen coordinates
		::SystemParametersInfo(SPI_GETWORKAREA, NULL, &rcArea, NULL);
		if(hWndCenter == NULL)
			rcCenter = rcArea;
		else
			::GetWindowRect(hWndCenter, &rcCenter);
	}
	else
	{
		// center within parent client coordinates
		hWndParent = ::GetParent(hWnd);
		assert(::IsWindow(hWndParent));
		
		::GetClientRect(hWndParent, &rcArea);
		assert(::IsWindow(hWndCenter));
		::GetClientRect(hWndCenter, &rcCenter);
		::MapWindowPoints(hWndCenter, hWndParent, (POINT*)&rcCenter, 2);
	}
	
	int DlgWidth = rcDlg.right - rcDlg.left;
	int DlgHeight = rcDlg.bottom - rcDlg.top;
	
	// find dialog's upper left based on rcCenter
	int xLeft = (rcCenter.left + rcCenter.right) / 2 - DlgWidth / 2;
	int yTop = (rcCenter.top + rcCenter.bottom) / 2 - DlgHeight / 2;
	
	// if the dialog is outside the screen, move it inside
	if(xLeft < rcArea.left)
		xLeft = rcArea.left;
	else if(xLeft + DlgWidth > rcArea.right)
		xLeft = rcArea.right - DlgWidth;
	
	if(yTop < rcArea.top)
		yTop = rcArea.top;
	else if(yTop + DlgHeight > rcArea.bottom)
		yTop = rcArea.bottom - DlgHeight;
	
	// map screen coordinates to child coordinates
	return ::SetWindowPos(hWnd, NULL, xLeft, yTop, -1, -1,
		SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE);
}


typedef struct tagDlgParam
{
	LONG nID;
	LPCTSTR szFilePath;
	HWND hParnetWnd;
}DLGPARAM, *LPDLGPARAM;

BOOL CALLBACK AskDialogProc(HWND hwndDlg, UINT message, WPARAM wParam, LPARAM lParam) 
{ 
    switch (message) 
    { 
        case WM_INITDIALOG:
			{
				LPDLGPARAM pDlgParam = (LPDLGPARAM)lParam;
				HWND hMsgHwnd = ::GetDlgItem(hwndDlg, IDC_STATIC_MSG);
				if (hMsgHwnd && pDlgParam && pDlgParam->szFilePath && pDlgParam->szFilePath[0]) {
					TCHAR szMsg[_MAX_PATH * 2];
					_stprintf(szMsg, _T("FilePathCheckerModuleExample가 %s 파일을 접근하려고 합니다.\n\n접근을 허용하시겠습니까?"), pDlgParam->szFilePath);
					::SetWindowText(hMsgHwnd, szMsg);
				}
				if (pDlgParam)
					CenterWindow(hwndDlg, pDlgParam->hParnetWnd);
			}
			break;
        case WM_COMMAND: 
            switch (LOWORD(wParam)) 
            { 
                case IDYES: 
                case IDYESALL: 
                case IDNO: 
                case IDNOALL: 
                    EndDialog(hwndDlg, wParam); 
                    return TRUE; 
            } 
			break;
    } 
    return FALSE; 
} 

// IsAccessiblePath function
// HWND hWnd : 메시지 박스의 부모가 될 윈도우 핸들
// LONG nID : HwpCtrl.RegisterModule(FilePathCheckHandle, nID) 를 통해 넘겨준 nID
// LPCTSTR szFilePath : 검사할 파일의 위치
// LPCTSTR szSiteInfo : 부가적으로 보내진 정보 (IE 브라우져 내의 컨트롤로 실행되면 URL:http://???.???.??? 등과같이 url이 보내지며 각 아이템은 \r\n으로 구분한다.
// 리턴값 : 허용 여부 TRUE/FALSE
FILEPATHCHECKERMODULEEXAMPLE_API BOOL __stdcall IsAccessiblePath(HWND hWnd, LONG nID, LPCTSTR szFilePath, LPCTSTR szSiteInfo)
{
	return TRUE;	// 무조건 팝업창 나타나지 않게..
	// 임시 폴더나 공개된 인터넷 주소를 접근하는 것은 허용한다.
	if (IsURLPath(szFilePath)) {
		MessageBox(NULL, szFilePath, _T("URL"), MB_OK);
		return TRUE;
	}
	if (IsTempPath(szFilePath)) {
		MessageBox(NULL, szFilePath, _T("TempPath"), MB_OK);
		return TRUE;
	}

	if (!CheckSiteInfo(szSiteInfo)) {
		return FALSE;
	}

	if (g_bAskDlg) {
		DLGPARAM dlgParam;
		dlgParam.nID = nID;
		dlgParam.szFilePath = szFilePath;
		dlgParam.hParnetWnd = hWnd;
		g_bAskDlgResult = DialogBoxParam(THIS_MODULE_INSTANCE, MAKEINTRESOURCE(IDD_ASK_DIALOG), hWnd, AskDialogProc, (LPARAM)&dlgParam);
	}

	switch (g_bAskDlgResult)
	{
	case IDYES:
		return TRUE;
	case IDYESALL: 
		g_bAskDlg = FALSE;
		return TRUE;
	case IDNO: 
		return FALSE;
	case IDNOALL: 
		g_bAskDlg = FALSE;
		return FALSE;
	}

	return FALSE;
}
