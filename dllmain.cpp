// dllmain.cpp: реализация DllMain.

#include "stdafx.h"
#include "resource.h"
#include "Wnd1S_i.h"
#include "dllmain.h"

CWnd1SModule _AtlModule;

class CWnd1SApp : public CWinApp
{
public:

// Переопределение
	virtual BOOL InitInstance();
	virtual int ExitInstance();

	DECLARE_MESSAGE_MAP()
};

BEGIN_MESSAGE_MAP(CWnd1SApp, CWinApp)
END_MESSAGE_MAP()

CWnd1SApp theApp;

BOOL CWnd1SApp::InitInstance()
{
	return CWinApp::InitInstance();
}

int CWnd1SApp::ExitInstance()
{
	return CWinApp::ExitInstance();
}
