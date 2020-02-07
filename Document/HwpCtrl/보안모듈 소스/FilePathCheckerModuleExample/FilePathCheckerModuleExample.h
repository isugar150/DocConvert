
// The following ifdef block is the standard way of creating macros which make exporting 
// from a DLL simpler. All files within this DLL are compiled with the FILEPATHCHECKERMODULEEXAMPLE_EXPORTS
// symbol defined on the command line. this symbol should not be defined on any project
// that uses this DLL. This way any other project whose source files include this file see 
// FILEPATHCHECKERMODULEEXAMPLE_API functions as being imported from a DLL, wheras this DLL sees symbols
// defined with this macro as being exported.
#ifdef FILEPATHCHECKERMODULEEXAMPLE_EXPORTS
#define FILEPATHCHECKERMODULEEXAMPLE_API __declspec(dllexport)
#else
#define FILEPATHCHECKERMODULEEXAMPLE_API __declspec(dllimport)
#endif

FILEPATHCHECKERMODULEEXAMPLE_API BOOL __stdcall IsAccessiblePath(HWND hWnd, LONG nID, LPCTSTR szFilePath);
