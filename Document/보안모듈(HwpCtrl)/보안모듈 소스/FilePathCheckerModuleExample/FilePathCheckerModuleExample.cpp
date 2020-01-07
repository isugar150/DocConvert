// FilePathCheckerModuleExample.cpp : Defines the entry point for the DLL application.
//

// Registry�� HKEY_CURRENT_USER\\Software\\HNC\\HwpCtrl\\Modules Ű��
// HwpCtrl.RegisterModule("FilePathCheckDLL", "[����̸�]") ���� ����� [����̸�]�� ������ 
// ���� DLL�� Full Path�� �����Ǿ� �־�� ������.


#include "stdafx.h"
#include <TCHAR.h>
#include <string.h>
#include <stdio.h>
#include <stdlib.h>
#include <assert.h>
#include "resource.h"
#include "FilePathCheckerModuleExample.h"

HINSTANCE THIS_MODULE_INSTANCE = NULL; // resource instance

BOOL g_bAskDlg = TRUE; // ��ȭ���ڸ� ��� ���� ����
INT_PTR g_bAskDlgResult = IDNO; // ��ȭ������ ����� ����

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

// �ӽ� �������� �˻�.
BOOL IsTempPath(LPCTSTR szFilePath)
{
	static TCHAR szTempPath[MAX_PATH] = {0};
	static nTempPathLen = 0;

	if (!szTempPath[0]) // �ѹ��� �����Ѵ�.
	{
		GetTempPath(sizeof(szTempPath), szTempPath);
		nTempPathLen = _tcslen(szTempPath);
	}

	if (_tcsnicmp(szFilePath, szTempPath, nTempPathLen) == 0) {
		return TRUE;
	}
	return FALSE;		
}

// ���ͳ� �ּ����� �˻�
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
	// SiteInfo �˻�
	// Ư�� ����Ʈ�� ���ø����̼ǿ����� �����ϵ��� �� �� �ִ�.
	// �ϴ��� ������ ��� ��Ų��.
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
			if ((szData = _tcsstr(szToken, _T("URL:"))) == szToken) { // URL �˻�
				// http://www.haansoft.com �� �����ϴ� URL�� ��� ��Ų��.
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
				 break;// �ѹ��̶� �����ϸ� ���ڿ� �м��� ���߰� ���з� �����Ѵ�.
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
					_stprintf(szMsg, _T("FilePathCheckerModuleExample�� %s ������ �����Ϸ��� �մϴ�.\n\n������ ����Ͻðڽ��ϱ�?"), pDlgParam->szFilePath);
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
// HWND hWnd : �޽��� �ڽ��� �θ� �� ������ �ڵ�
// LONG nID : HwpCtrl.RegisterModule(FilePathCheckHandle, nID) �� ���� �Ѱ��� nID
// LPCTSTR szFilePath : �˻��� ������ ��ġ
// LPCTSTR szSiteInfo : �ΰ������� ������ ���� (IE ������ ���� ��Ʈ�ѷ� ����Ǹ� URL:http://???.???.??? ������� url�� �������� �� �������� \r\n���� �����Ѵ�.
// ���ϰ� : ��� ���� TRUE/FALSE
FILEPATHCHECKERMODULEEXAMPLE_API BOOL __stdcall IsAccessiblePath(HWND hWnd, LONG nID, LPCTSTR szFilePath, LPCTSTR szSiteInfo)
{
	return TRUE;	// ������ �˾�â ��Ÿ���� �ʰ�..
	// �ӽ� ������ ������ ���ͳ� �ּҸ� �����ϴ� ���� ����Ѵ�.
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
