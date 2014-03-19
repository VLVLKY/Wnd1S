// stdafx.h: включаемый файл дл€ стандартных системных включаемых файлов
// или включаемых файлов дл€ конкретного проекта, которые часто используютс€,
// но не часто измен€ютс€

#pragma once

#ifndef STRICT
#define STRICT
#endif

#include "targetver.h"

#define _ATL_APARTMENT_THREADED

#define _ATL_NO_AUTOMATIC_NAMESPACE

#define _ATL_CSTRING_EXPLICIT_CONSTRUCTORS	// некоторые конструкторы CString будут €вными

#include <afxwin.h>
#include <afxext.h>
#include <afxole.h>
#include <afxodlgs.h>
#include <afxrich.h>
#include <afxhtml.h>
#include <afxcview.h>
#include <afxwinappex.h>
#include <afxframewndex.h>
#include <afxmdiframewndex.h>

#ifndef _AFX_NO_OLE_SUPPORT
#include <afxdisp.h>        // классы автоматизации MFC
#endif // _AFX_NO_OLE_SUPPORT

#define ATL_NO_ASSERT_ON_DESTROY_NONEXISTENT_WINDOW

#include "resource.h"
#include <atlbase.h>
#include <atlcom.h>
#include <atlctl.h>

extern "C" const GUID ;
extern "C" const GUID ;
extern "C" const GUID ;

#define TRACING

#if defined(TRACING) \

#include "string"
#include "fstream"
#include <ctime>
#include <iomanip>

#define TRC(fileName, outString) \
{\
	using namespace std; \
	time_t     rawtime; \
	struct tm* timeinfo; \
	ofstream file(fileName, ofstream::app); \
	file.setf(ios::unitbuf); \
	\
	time(&rawtime); \
	timeinfo = localtime(&rawtime); \
	\
	file \
	<< setfill('0') << setw(2) << timeinfo->tm_mday << "/" \
	<< setfill('0') << setw(2) << (timeinfo->tm_mon + 1)  << "/" \
	<< (timeinfo->tm_year + 1900) << " " \
	<< setfill('0') << setw(2) << timeinfo->tm_hour << ":" \
	<< setfill('0') << setw(2) << timeinfo->tm_min << ":" \
	<< setfill('0') << setw(2) << timeinfo->tm_sec \
	<< ": " \
	<< outString << endl; \
	file.close();\
}

#else

#define TRC(fileName, outString)((void) 0)

#endif

#include "string"
 
// ƒл€ преобразований BSTR:
#include <comutil.h>
#include <comdef.h>

//#include <afxcmn.h>