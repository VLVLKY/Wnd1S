// Juncter.h: ���������� CJuncter

#pragma once
#include "resource.h"       // �������� �������



#include "Wnd1S_i.h"

#include "atlctl.h"

#if defined(_WIN32_WCE) && !defined(_CE_DCOM) && !defined(_CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA)
#error "������������� COM-������� �� �������������� ������� ������� ���������� Windows CE, �������� ����������� Windows Mobile, � ������� �� ������������� ������ ��������� DCOM. ���������� _CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA, ����� ��������� ATL ������������ �������� ������������� COM-�������� � ��������� ������������� ��� ���������� ������������� COM-��������. ��� ��������� ������ � ����� rgs-����� ������ �������� 'Free', ��������� ��� ������������ ��������� ������, �������������� ��-DCOM ����������� Windows CE."
#endif

using namespace ATL;

// ��������� ���� ���������.
enum AddInErrors {

	// ��������� ������ � ���� ���������:
	ADDIN_E_NONE = 1000,
	ADDIN_E_ORDINARY = 1001,
	ADDIN_E_ATTENTION = 1002,
	ADDIN_E_IMPORTANT = 1003,
	ADDIN_E_VERY_IMPORTANT = 1004,
	ADDIN_E_INFO = 1005,
	ADDIN_E_FAIL = 1006,

	// ��������� ���� �������������� � ������� ��:
	ADDIN_E_MSGBOX_ATTENTION = 1007,
	ADDIN_E_MSGBOX_INFO = 1008,
	ADDIN_E_MSGBOX_FAIL = 1009
};

// CJuncter

class ATL_NO_VTABLE CJuncter :
	public CComObjectRootEx<CComSingleThreadModel>,//+
	public CComCoClass<CJuncter, &CLSID_Juncter>,//+
	public IInitDone,//+
	public IDispatchImpl<ILanguageExtender, &__uuidof(ILanguageExtender), &LIBID_Wnd1SLib, /* wMajor = */ 1, /* wMinor = */ 0>,//+
	public IDispatchImpl<IPropertyLink, &__uuidof(IPropertyLink), &LIBID_Wnd1SLib, /* wMajor = */ 1, /* wMinor = */ 0>//+
{

	// �������� �������� ��������.
	BOOL m_someBooleanPropertyValue;

	// ��������.
	enum
	{
		someBooleanProperty = 0, // ��������� ��������.
		//propIsTimerPresent,
		LastProp                 // ���� ������� ������ ���� ��������� � ����������� ��� ��������� ���������� ������������ �������.
	};

	// ������
	enum 
	{
		someMethod = 0, // ��������� �����.
		longFunc,       // ��������� �������.
		doubleFunc,     // �������.
		stringFunc,     // ���������.
		dateFunc,       // ����.
		dateTimeFunc,   // ���������.
		arrayFunc,      // ������.
		twoParamsFunc,  // ��� ���������.
		callBackFunc,   // �������, ���������� ����� 1�.
		LastMethod      // ���� ������� ������ ���� ��������� � ����������� ��� ��������� ���������� ������������ �������.
	};

public:

	// ��������� �� ��������� 1�.
	CComPtr<IDispatch> m_iConnect;

	// ��������� ��� ������ � ���������� ���������� ����������.
	// IPropertyProfile ����������� �� ������������ IPropertyBag.
	// ���������� ��������� � ������ JuncterConnection.
	//IPropertyProfile *m_iProfile;

	// ���������, ������������ ��� �������� �������� ������� � 1�.
	// ���������� ��������� � ������ JuncterConnection.
	IAsyncEvent *m_iAsyncEvent;

	// ���������, ������������ ��� �������� �������� ������� � 1�.
	// ���������� ��������� � ������ JuncterConnection.
	//IStatusLine *m_iStatusLine;

	// ����������� COM-���������.
	// ������������ ��� ��������� ������������ ���������� � ����� ������.
	// ����������� ��������� �������������� ��� � ������� ������ ��������� (��� ����������� ���������� �� � �������),
	// ��� � ���:
	// - �������� �� ������ ������������� Init;
	// - �������� �� ������ ����������.
	//IErrorLog *m_iErrorLog;


	// These two methods is convenient way to access function 
	// parameters from SAFEARRAY vector of variants
	VARIANT GetNParam(SAFEARRAY *pArray,long lIndex);
	void PutNParam(SAFEARRAY *pArray,long lIndex,VARIANT vt);

	CJuncter()
	{
		// m_dwTitleID = IDS_PROPPAGE_CAPTION; 
	}

	DECLARE_REGISTRY_RESOURCEID(IDR_JUNCTER)


	BEGIN_COM_MAP(CJuncter)
		COM_INTERFACE_ENTRY(IInitDone)//+
		COM_INTERFACE_ENTRY(ILanguageExtender)//+
		COM_INTERFACE_ENTRY(IPropertyLink)//+
	END_COM_MAP()



	DECLARE_PROTECT_FINAL_CONSTRUCT()

	HRESULT FinalConstruct()
	{
		return S_OK;
	}

	void FinalRelease()
	{
	}

public:

	// ������ ������
	CString TermString(UINT uiResID,long nAlias);

	// ������ �����������

	// IInitDone
	STDMETHOD(Init)(IDispatch* pConnection);
	STDMETHOD(Done)(void);
	STDMETHOD(GetInfo)(SAFEARRAY ** pInfo);

	// ILanguageExtender Methods
	STDMETHOD(RegisterExtensionAs)(BSTR * bstrExtensionName);

	// ������ �� ����������.
	STDMETHOD(GetNProps)(long *plProps);
	STDMETHOD(FindProp)(BSTR bstrPropName,long *plPropNum);
	STDMETHOD(GetPropName)(long lPropNum,long lPropAlias,BSTR *pbstrPropName);
	STDMETHOD(GetPropVal)(long lPropNum,VARIANT *pvarPropVal);
	STDMETHOD(SetPropVal)(long lPropNum,VARIANT *pvarPropVal);
	STDMETHOD(IsPropReadable)(long lPropNum,BOOL *pboolPropRead);
	STDMETHOD(IsPropWritable)(long lPropNum,BOOL *pboolPropWrite);

	// ������ � ��������.
	STDMETHOD(GetNMethods)(long *plMethods);
	STDMETHOD(FindMethod)(BSTR bstrMethodName,long *plMethodNum);
	STDMETHOD(GetMethodName)(long lMethodNum,long lMethodAlias,BSTR *pbstrMethodName);
	STDMETHOD(HasRetVal)(long lMethodNum,BOOL *pboolRetValue);

	// ������ � �����������.
	STDMETHOD(GetNParams)(long lMethodNum, long *plParams);
	STDMETHOD(GetParamDefValue)(long lMethodNum,long lParamNum,VARIANT *pvarParamDefValue);

	STDMETHOD(CallAsProc)(long lMethodNum,SAFEARRAY **paParams);
	STDMETHOD(CallAsFunc)(long lMethodNum,VARIANT *pvarRetValue,SAFEARRAY **paParams);

	// IPropertyLink Methods
	STDMETHOD(get_Enabled)(BOOL * boolEnabled);
	STDMETHOD(put_Enabled)(BOOL boolEnabled);
};

OBJECT_ENTRY_AUTO(__uuidof(Juncter), CJuncter)
