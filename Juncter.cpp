// Juncter.cpp: ���������� CJuncter

#include "stdafx.h"
#include "Juncter.h"

//#include <Rpcdce.h>
//#include "comutil.h"

// ������� ��������� ������ 1� �� ����������� �����
//
HRESULT ExecuteMethod_1C(IDispatch* pDisp, LPTSTR NameMethod, VARIANT* vRetVal, VARIANTARG* vParams, int ParmCount, WORD Flag)
{
	if(pDisp)
	{
		try
		{
			_bstr_t str(NameMethod);
			
			OLECHAR FAR* szMember = str.GetBSTR(); // �������� ��� ������ � ������� ����
			DISPID dispid = NULL; // ��������� �� ����� ������� 1�
			HRESULT res = pDisp->GetIDsOfNames(IID_NULL, &szMember, 1, LOCALE_SYSTEM_DEFAULT, &dispid);
			if(SUCCEEDED(res)) // ���� ����� ����������
			{
				DISPPARAMS dParams = {vParams, NULL, ParmCount, 0}; // �������� ��������� ����������
				
				// ������� ���
				return pDisp->Invoke(dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT, Flag, &dParams, vRetVal, NULL, 0);
			}
			return E_NOTIMPL;
		}
		catch(_com_error er)
		{
			//ErrorMessage("������ ���������� ������ 1�",(LPCSTR)er.Description());
			return S_FALSE;
		}
	}
	return S_FALSE;
}

// ��������� �������, ������� ���������� ������� � ����.
//
STDMETHODIMP ShowSafeArray(const char * fileName, SAFEARRAY ** sArr) {

	//typedef struct tagSAFEARRAYBOUND
	//{
	//	unsigned long cElements; // ����� ��������� � �������.
	//	long lLbound;            // ������ ������� �������.
	//} SAFEARRAYBOUND;

	//typedef struct tagSAFEARRAY
	//{
	//	unsigned short cDims;        // ����� ������������ � �������.
	//	unsigned fFeatures;          // ����� ������, ������������ ��� ������������ ������ �������.
	//	unsigned long cbElements;    // ������ ������� �� ��������� �������.
	//	unsigned long cLocks;        // ����� ���������� �������.
	//	unsigned short handle;       // �� ������������, ��������� ��� �������������.
	//	void* pvData;                // ��������� �� ������ �������.
	//	SAFEARRAYBOUND rgsabound[1]; // ���������, ����������� ������ ������� ������� � ����� ���������.
	//} SAFEARRAY;
	//SAFEARRAY *pSA = NULL;

	long lBound, uBound;
	SafeArrayGetLBound(*sArr, 1, &lBound);
	SafeArrayGetUBound(*sArr, 1, &uBound);

	VARIANT item;
	V_VT(&item) = VT_I4;

	std::string ss("ShowSafeArray:");
	TRC(fileName, ss);

	for (long j = lBound; j <= uBound; j++) {

		SafeArrayGetElement(*sArr, &j,  &item);
		TRC(fileName, V_I4(&item));
	}

	return S_OK;
}

// CJuncter

// ������ ������
/* 
These two methods is convenient way to access function 
parameters from SAFEARRAY vector of variants
*/
VARIANT CJuncter::GetNParam(SAFEARRAY *pArray, long lIndex)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState())

		std::string ss ("CJuncter::GetNParam");
	TRC("1.txt", ss);

	ASSERT(pArray);
	ASSERT(pArray->fFeatures | FADF_VARIANT); // �������� ����, ��� � ������� ���������� ������ VARIANT?

	VARIANT vt;
	HRESULT hRes = ::SafeArrayGetElement(pArray, &lIndex, &vt);
	ASSERT(hRes == S_OK);

	ShowSafeArray("1.txt", &pArray);

	return vt;
}

void CJuncter::PutNParam(SAFEARRAY *pArray,long lIndex,VARIANT vt)
{

	AFX_MANAGE_STATE(AfxGetStaticModuleState())

		std::string ss ("CJuncter::PutNParam");
	TRC("1.txt", ss);

	ASSERT(pArray);
	ASSERT(pArray->fFeatures | FADF_VARIANT);

	HRESULT hRes = ::SafeArrayPutElement(pArray,&lIndex,&vt);
	ASSERT(hRes == S_OK);

	ShowSafeArray("1.txt", &pArray);
}

// ���������� ������ � ������ �����
//.
CString CJuncter::TermString(UINT uiResID, long nAlias)
{

	AFX_MANAGE_STATE(AfxGetStaticModuleState())

		//std::string ss ("CJuncter::TermString");
		//TRC("1.txt", ss);

		CString cs;
	cs.LoadString(uiResID);
	int iInd = cs.Find(',');
	if (iInd == -1)
		return cs;
	switch(nAlias)
	{
	case 0: // First language
		return cs.Left(iInd);
	case 1: // Second language
		return cs.Right(cs.GetLength()-iInd-1);
	default:
		return CString("");
	};

}


// ������ �����������

STDMETHODIMP CJuncter::Init(IDispatch* pConnection)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState())

	std::string ss ("CJuncter::Init ");
	
	TRC("1.txt", ss);
	TRC("1.txt", pConnection);

	// ��������� �� ��������� 1�.
	m_iConnect = pConnection;

	//m_iErrorLog = NULL;
	//pConnection->QueryInterface(IID_IErrorLog,(void **)&m_iErrorLog);
	//TRC("1.txt", m_iErrorLog);

	m_iAsyncEvent = NULL;
	pConnection->QueryInterface(IID_IAsyncEvent,(void **)&m_iAsyncEvent);
	TRC("1.txt", m_iAsyncEvent);

	//m_iStatusLine = NULL;
	//pConnection->QueryInterface(IID_IStatusLine,(void **)&m_iStatusLine);
	//TRC("1.txt", m_iStatusLine);

	//CString csProfileName = "Sample AddIn Profile Name";
	//m_iProfile = NULL;
	//pConnection->QueryInterface(IID_IPropertyProfile,(void **)&m_iProfile);
	//if (m_iProfile) 
	//{
	//    m_iProfile->RegisterProfileAs(csProfileName.AllocSysString());
	//    if (LoadProperties() == FALSE)
	//        return E_FAIL;
	//}

	return S_OK;
}


STDMETHODIMP CJuncter::Done(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState())

	std::string ss ("CJuncter::Done");
	TRC("1.txt", ss);

	//if (m_iStatusLine)
	//	m_iStatusLine->Release();

	//TRC("1.txt", m_iStatusLine);

	//if (m_iProfile)
	//	m_iProfile->Release();


	if (m_iAsyncEvent)
		m_iAsyncEvent->Release();

	TRC("1.txt", m_iAsyncEvent);


	//if (m_iErrorLog)
	//	m_iErrorLog->Release();

	//TRC("1.txt", m_iErrorLog);

	return S_OK;
}


STDMETHODIMP CJuncter::GetInfo(SAFEARRAY ** pInfo)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState())

		std::string ss ("CJuncter::GetInfo");
	TRC("1.txt", ss);

	TRC("1.txt", pInfo);
	TRC("1.txt", *pInfo);

	// ��������� ������ ���������� ����� ������ �������������� ����������� ����������
	// � ������� VARIANT �� ������� 0.    

	// ������, ��� ������� ����� ������� ������� � ������.
	long lInd = 0;

	// ����������, � ������� ����� ��������� �������� ������.
	VARIANT varVersion;

	// ��������� ���� ��������, ������� ����� ��������� � ���� ���� VARIANT.
	V_VT(&varVersion) = VT_I4;

	// ��������� ��������, ������� ����� ��������� � ���� ���� VARIANT.
	// ��������� ������������ ������ 2.0, ����������� ��������� ��������� �������� ������ ����
	// AddIn.<�������������>.
	// ���� ������� ������ 1.0, �� ����������� �������� ������ ������ �������.
	V_I4(&varVersion) = 2000;

	// ��������� ���������� �������� � ������.
	SafeArrayPutElement(
		*pInfo,     // [in] Pointer to an array descriptor created by SafeArrayCreate.
		&lInd,      // [in] Pointer to a vector of indexes for each dimension of the array. The right-most (least significant) dimension is rgIndices[0]. The left-most dimension is stored at rgIndices[psa->cDims �1].
		&varVersion // [out] Void pointer to the data to assign to the array. The variant types VT_DISPATCH, VT_UNKNOWN, and VT_BSTR are pointers, and do not require another level of indirection.
		);

	ShowSafeArray("1.txt", pInfo);

	return S_OK;
}

STDMETHODIMP CJuncter::RegisterExtensionAs(BSTR * bstrExtensionName)
{

	AFX_MANAGE_STATE(AfxGetStaticModuleState())

		USES_CONVERSION;

	std::string ss ("CJuncter::RegisterExtensionAs");
	TRC("1.txt", ss);
	TRC("1.txt", bstrExtensionName);
	TRC("1.txt", *bstrExtensionName);
	//TRC("1.txt", OLE2T(*bstrExtensionName)); - ������, �.�. ������ ��� �� ��������!

	CString csExtenderName = "Win1SConnector";
	*bstrExtensionName = csExtenderName.AllocSysString();

	TRC("1.txt", *bstrExtensionName);
	TRC("1.txt", OLE2T(*bstrExtensionName));

	return NULL;
}

// ���������� ���������� ������� ����������.
//
STDMETHODIMP CJuncter::GetNProps(long *plProps)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState())

		std::string ss ("CJuncter::GetNProps");
	TRC("1.txt", ss);
	TRC("1.txt", plProps);
	TRC("1.txt", *plProps);

	// LastProp - ��� ������ ��������� ������� ������������, ������������� � ������ CJuncter.
	*plProps = LastProp;
	TRC("1.txt", *plProps);

	return S_OK;
}

// ���������� ���������� ����� �������� � ������ bstrPropName.
//
STDMETHODIMP CJuncter::FindProp(BSTR bstrPropName, long *plPropNum)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState())

		USES_CONVERSION;

	std::string ss ("CJuncter::FindProp");
	TRC("1.txt", ss);
	TRC("1.txt", bstrPropName);
	TRC("1.txt", OLE2T(bstrPropName));
	TRC("1.txt", plPropNum);
	TRC("1.txt", *plPropNum);

	*plPropNum = -1;

	// �������������� � CString.
	CString csPropName = OLE2T(bstrPropName);

	// ��������� ��������� ������� (����������� � ��������� �������) � ���������� ������.
	// ������ � ������ �������� ����� �������� �� ���� ����, ����������� �������.
	// ���� ������ TermString �������� ������, ���������� �������, � ������ ���������� ��� �������� 0,
	// �� ��������� ����� ������������� � ������ ������ �� ������� (��������� � ������ ������);
	// � ���� ������ ���������� �������� 1,
	// �� ��������� ����� ������������� � ������ ������ ����� ������� (��������� �� ������ ������).
	if (TermString(IDS_SOME_BOOL_PROP, 0) == csPropName) // ��������� � ������ ������ (���������� ������������).
		*plPropNum = 0;

	if (TermString(IDS_SOME_BOOL_PROP, 1) == csPropName) // ��������� �� ������ ������ (��������� ������������).
		*plPropNum = 0;

	TRC("1.txt", plPropNum);
	TRC("1.txt", *plPropNum);

	// ���� �������� ���� �������, �� ������������ S_OK.
	return (*plPropNum == -1) ? S_FALSE : S_OK;
}

// ���������� ��� �������� � ���������� ������� lPropNum.
// lPropAlias: 0 - ���������� ������������;
//             1 - ��������� ������������.
//
STDMETHODIMP CJuncter::GetPropName(long lPropNum, long lPropAlias, BSTR *pbstrPropName)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState())

		USES_CONVERSION;

	std::string ss ("CJuncter::GetPropName");
	TRC("1.txt", ss);
	TRC("1.txt", lPropNum);
	TRC("1.txt", lPropAlias);
	TRC("1.txt", pbstrPropName);
	TRC("1.txt", *pbstrPropName);
	//TRC("1.txt", OLE2T(*pbstrPropName));

	CString csPropName = "";

	// 
	switch(lPropNum)
	{
	case someBooleanProperty: // �� �� �����, ��� 0, �.�. someBooleanProperty ���� � ������������ ������.
		// �.�. ���� �������� 0, �� ����� ���������� ��������� ������������� ������ ������� ��������, ��� ��� ������ � ������� �����.
		csPropName = TermString(IDS_SOME_BOOL_PROP, lPropAlias);
		*pbstrPropName = csPropName.AllocSysString();
		break;
	default:

		*pbstrPropName = csPropName.AllocSysString();

		return S_FALSE;
	}

	TRC("1.txt", *pbstrPropName);

	return S_OK;
}

// ���������� �������� �������� � ���������� ������� lPropNum.
//
STDMETHODIMP CJuncter::GetPropVal(long lPropNum, VARIANT *pvarPropVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState())

		std::string ss ("CJuncter::GetPropVal");

	TRC("1.txt", ss);
	TRC("1.txt", lPropNum);
	TRC("1.txt", pvarPropVal);
	TRC("1.txt", V_I4(pvarPropVal));


	// ������������� �������� ���� Variant.
	// ���������, ����� � ������ ���������� �������� � ����� ������� � ���������� ��� �������� ��� VT_EMPTY.
	// ������� ��������, ��� ������� VariantInit ��� ��� ��������� ����� ���.
	::VariantInit(pvarPropVal);

	TRC("1.txt", V_I4(pvarPropVal));

	switch(lPropNum)
	{
	case someBooleanProperty: // �� �� �����, ��� 0, �.�. someBooleanProperty ���� � ������������ ������.

		// ���������� ����, ������� ����� ��������� � ���������� pvarPropVal.
		V_VT(pvarPropVal) = VT_I4;

		// ���� ���� ����������, �� �������� 1, ����� 0.
		V_I4(pvarPropVal) = m_someBooleanPropertyValue ? 1:0;

		break;

	default:

		ss = "CJuncter::GetPropVal - default";
		TRC("1.txt", pvarPropVal);

		return S_FALSE;
	}

	TRC("1.txt", pvarPropVal);
	TRC("1.txt", V_I4(pvarPropVal));

	return S_OK;
}

// ������������� �������� �������� � ���������� ������� lPropNum.
//
STDMETHODIMP CJuncter::SetPropVal(long lPropNum, VARIANT *pvarPropVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState())

		std::string ss ("CJuncter::SetPropVal");
	TRC("1.txt", ss);
	TRC("1.txt", lPropNum);
	TRC("1.txt", pvarPropVal);
	TRC("1.txt", V_I4(pvarPropVal));

	switch(lPropNum)
	{ 
	case someBooleanProperty: // �� �� �����, ��� 0, �.�. someBooleanProperty ���� � ������������ ������.

		// �������� ����.
		if (V_VT(pvarPropVal) == VT_I4)
		{
			ss = "CJuncter::SetPropVal - type VT_I4";

		} else if (V_VT(pvarPropVal) == VT_R8)
		{
			ss = "CJuncter::SetPropVal - type VT_R8";
		}
		else if (V_VT(pvarPropVal) == VT_DATE)
		{
			ss = "CJuncter::SetPropVal - type VT_DATE";
		}
		else if (V_VT(pvarPropVal) == VT_BSTR)
		{
			ss = "CJuncter::SetPropVal - type VT_BSTR";
		}
		else if (V_VT(pvarPropVal) == VT_DISPATCH)
		{
			ss = "CJuncter::SetPropVal - type VT_DISPATCH";
		}
		else if (V_VT(pvarPropVal) == VT_ARRAY)
		{
			ss = "CJuncter::SetPropVal - type VT_ARRAY";
		}
		else
		{
			ss = "CJuncter::SetPropVal - other type!";
		}

		TRC("1.txt", ss);

		// ���� ��������������� �������� �� ����� ������ ���, �� ��� ����������� ���� �� �����.
		if (V_VT(pvarPropVal) != VT_I4)
		{
			ss = "CJuncter::SetPropVal - wrong type";
			TRC("1.txt", ss);

			return S_FALSE;
		}

		// ��������� �������� ��������.
		m_someBooleanPropertyValue = V_I4(pvarPropVal)?1:0;

		break;

	default:

		ss = "CJuncter::SetPropVal - default";
		TRC("1.txt", ss);
		TRC("1.txt", m_someBooleanPropertyValue);

		return S_FALSE;
	}

	TRC("1.txt", m_someBooleanPropertyValue);

	TRC("1.txt", pvarPropVal);
	TRC("1.txt", V_I4(pvarPropVal));

	return S_OK;
}

// ���������� ���� ����������� ������ �������� � ���������� ������� lPropNum.
//
STDMETHODIMP CJuncter::IsPropReadable(long lPropNum, BOOL *pboolPropRead)
{
	// Call this macro to protect an exported function in a DLL.
	AFX_MANAGE_STATE(AfxGetStaticModuleState())

		std::string ss ("CJuncter::IsPropReadable");
	TRC("1.txt", ss);
	TRC("1.txt", lPropNum);
	TRC("1.txt", pboolPropRead);
	TRC("1.txt", *pboolPropRead);

	switch(lPropNum)
	{ 
	case someBooleanProperty: // �� �� �����, ��� 0, �.�. someBooleanProperty ���� � ������������ ������.

		*pboolPropRead = TRUE;

		break;

	default:

		ss = "CJuncter::IsPropReadable - default";
		TRC("1.txt", *pboolPropRead);

		return S_FALSE;
	}

	TRC("1.txt", pboolPropRead);
	TRC("1.txt", *pboolPropRead);

	return S_OK;
}

// ���������� ���� ����������� ������ �������� � ���������� ������� lPropNum.
//
STDMETHODIMP CJuncter::IsPropWritable(long lPropNum, BOOL *pboolPropWrite)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState())

		std::string ss ("CJuncter::IsPropWritable");
	TRC("1.txt", ss);
	TRC("1.txt", lPropNum);
	TRC("1.txt", pboolPropWrite);
	TRC("1.txt", *pboolPropWrite);

	switch(lPropNum)
	{ 
	case someBooleanProperty: // �� �� �����, ��� 0, �.�. someBooleanProperty ���� � ������������ ������.

		*pboolPropWrite = TRUE;

		break;

	default:

		ss = "CJuncter::IsPropWritable - default";
		TRC("1.txt", ss);

		return S_FALSE;
	}

	TRC("1.txt", pboolPropWrite);
	TRC("1.txt", *pboolPropWrite);

	return S_OK;
}

// ���������� ���������� ������� ����������.
//
STDMETHODIMP CJuncter::GetNMethods(long *plMethods)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState())

		std::string ss ("CJuncter::GetNMethods");
	TRC("1.txt", ss);

	// LastMethod - ��� ������ ��������� ������� ������������, ������������� � ������ CJuncter.

	*plMethods = LastMethod;

	return S_OK;
}

// ���������� ���������� ����� ������ � ������ bstrMethodName.
//
STDMETHODIMP CJuncter::FindMethod(BSTR bstrMethodName, long *plMethodNum)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState())

		// USES_CONVERSION is a macro that allows to use BSTR string conversion functions like OLE2T and T2OLE.
		USES_CONVERSION;

	std::string ss ("CJuncter::FindMethod");
	TRC("1.txt", ss);
	TRC("1.txt", bstrMethodName);
	TRC("1.txt", OLE2T(bstrMethodName));
	TRC("1.txt", plMethodNum);
	TRC("1.txt", *plMethodNum);

	*plMethodNum = -1;

	// �������������� � CString.
	CString csMethodName = OLE2T(bstrMethodName);

	// ��������� ��������� ������� (����������� � ��������� �������) � ���������� ������.
	// ������ � ������ �������� ����� �������� �� ���� ����, ����������� �������.
	// ���� ������ TermString �������� ������, ���������� �������, � ������ ���������� ��� �������� 0,
	// �� ��������� ����� ������������� � ������ ������ �� ������� (��������� � ������ ������);
	// � ���� ������ ���������� �������� 1,
	// �� ��������� ����� ������������� � ������ ������ ����� ������� (��������� �� ������ ������).

	// �����.
	if (csMethodName == TermString(IDS_SOME_METHOD, 0)) // ��������� � ������ ������ (���������� ������������).
		*plMethodNum = 0;

	if (csMethodName == TermString(IDS_SOME_METHOD, 1)) // ��������� �� ������ ������ (��������� ������������).
		*plMethodNum = 0;

	// ������������� �������.
	if (csMethodName == TermString(IDS_LONG_FUNC, 0))
		*plMethodNum = 1;

	if (csMethodName == TermString(IDS_LONG_FUNC, 1))
		*plMethodNum = 1;

	// �������.
	if (csMethodName == TermString(IDS_DOUBLE_FUNC, 0))
		*plMethodNum = 2;

	if (csMethodName == TermString(IDS_DOUBLE_FUNC, 1))
		*plMethodNum = 2;

	// ���������.
	if (csMethodName == TermString(IDS_STRING_FUNC, 0))
		*plMethodNum = 3;

	if (csMethodName == TermString(IDS_STRING_FUNC, 1))
		*plMethodNum = 3;

	// ����.
	if (csMethodName == TermString(IDS_DATE_FUNC, 0))
		*plMethodNum = 4;

	if (csMethodName == TermString(IDS_DATE_FUNC, 1))
		*plMethodNum = 4;

	// ���������.
	if (csMethodName == TermString(IDS_DATETIME_FUNC, 0))
		*plMethodNum = 5;

	if (csMethodName == TermString(IDS_DATETIME_FUNC, 1))
		*plMethodNum = 5;

	// ������.
	if (csMethodName == TermString(IDS_ARRAY_FUNC, 0))
		*plMethodNum = 6;

	if (csMethodName == TermString(IDS_ARRAY_FUNC, 1))
		*plMethodNum = 6;

	if (csMethodName == TermString(IDS_TWO_PARAMS_FUNC, 0))
		*plMethodNum = 7;

	if (csMethodName == TermString(IDS_TWO_PARAMS_FUNC, 1))
		*plMethodNum = 7;

	if (csMethodName == TermString(IDS_CALLBACK, 0))
		*plMethodNum = 8;

	if (csMethodName == TermString(IDS_CALLBACK, 1))
		*plMethodNum = 8;

	TRC("1.txt", plMethodNum);
	TRC("1.txt", *plMethodNum);
	
	// ���� ����� ��� ������, �� ������������ S_OK.
	return (*plMethodNum == -1) ? S_FALSE : S_OK;
}

// ���������� ��� ������ � ���������� ������� lMethodNum.
// lMethodAlias: 0 - ���������� ������������;
//               1 - ��������� ������������.
//
STDMETHODIMP CJuncter::GetMethodName(long lMethodNum, long lMethodAlias, BSTR *pbstrMethodName)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState())

		USES_CONVERSION;

	std::string ss ("CJuncter::GetMethodName");
	TRC("1.txt", ss);
	TRC("1.txt", lMethodNum);
	TRC("1.txt", lMethodAlias);
	TRC("1.txt", pbstrMethodName);
	TRC("1.txt", *pbstrMethodName);
	//TRC("1.txt", OLE2T(*pbstrMethodName));

	CString csMethodName = "";     

	//
	switch(lMethodNum)
	{ 
	case someMethod: // �� �� �����, ��� 0, �.�. someMethod ���� � ������������ ������.

		// �.�. ���� �������� 0, �� ����� ���������� ��������� ������������� ������ ������� ��������, ��� ��� ������ � ������� �����.
		csMethodName = TermString(IDS_SOME_METHOD, lMethodAlias);

		break;

	case longFunc:

		csMethodName = TermString(IDS_LONG_FUNC, lMethodAlias);

		break;

	case doubleFunc:

		csMethodName = TermString(IDS_DOUBLE_FUNC, lMethodAlias);

		break;

	case stringFunc:

		csMethodName = TermString(IDS_STRING_FUNC, lMethodAlias);

		break;

	case dateFunc:

		csMethodName = TermString(IDS_DATE_FUNC, lMethodAlias);

		break;

	case dateTimeFunc:

		csMethodName = TermString(IDS_DATETIME_FUNC, lMethodAlias);

		break;

	case arrayFunc:

		csMethodName = TermString(IDS_ARRAY_FUNC, lMethodAlias);

		break;

	case twoParamsFunc:

		csMethodName = TermString(IDS_TWO_PARAMS_FUNC, lMethodAlias);

		break;

	case callBackFunc:

		csMethodName = TermString(IDS_CALLBACK, lMethodAlias);

		break;

	default:

		*pbstrMethodName = csMethodName.AllocSysString();

		return S_FALSE;
	}

	*pbstrMethodName = csMethodName.AllocSysString();

	TRC("1.txt", OLE2T(*pbstrMethodName));

	return S_OK;
}

// ���������� ���������� ���������� ������ � ���������� ������� lMethodNum.
//
STDMETHODIMP CJuncter::GetNParams(long lMethodNum, long *plParams)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState())

		std::string ss ("CJuncter::GetNParams");
	TRC("1.txt", ss);
	TRC("1.txt", plParams);
	TRC("1.txt", *plParams);

	*plParams = 0;

	switch(lMethodNum)
	{ 
	case someMethod:

		*plParams = 0; // � ������ someMethod ��� ����������.

		break;

	case longFunc:

		*plParams = 1; // ������� someFunction ���������� 1 ��������.

		break;

	case doubleFunc:

		*plParams = 1;

		break;

	case stringFunc:

		*plParams = 1;

		break;

	case dateFunc:

		*plParams = 1;

		break;

	case dateTimeFunc:

		*plParams = 1;

		break;

	case arrayFunc:

		*plParams = 1;

		break;

	case twoParamsFunc:

		*plParams = 2; // ��� ���������!

		break;

	case callBackFunc:

		*plParams = 0;

		break;

	default:

		ss = "CJuncter::GetNParams - default";
		TRC("1.txt", ss);

		return S_FALSE;
	}

	TRC("1.txt", *plParams);

	return S_OK;
}

// ���������� �������� �� ��������� ��� ��� ��������� � ������� lParamNum ������ � ������� lMethodNum.
//
STDMETHODIMP CJuncter::GetParamDefValue(long lMethodNum, long lParamNum, VARIANT *pvarParamDefValue)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState())

		std::string ss ("CJuncter::GetParamDefValue");
	TRC("1.txt", ss);
	TRC("1.txt", lMethodNum);
	TRC("1.txt", lParamNum);
	TRC("1.txt", pvarParamDefValue);
	TRC("1.txt", V_I4(pvarParamDefValue));

	// ������������� ������������� �������� ��������� VT_EMPTY.
	::VariantInit(pvarParamDefValue);

	TRC("1.txt", pvarParamDefValue);
	TRC("1.txt", V_I4(pvarParamDefValue));

	switch(lMethodNum)
	{ 
	case someMethod:

		if (someMethod == 0) {
			ss = "CJuncter::GetParamDefValue, someMethod = 0";
		} else {
			ss = "CJuncter::GetParamDefValue, someMethod = " + std::to_string(lMethodNum);
		}
		TRC("1.txt", ss);

		if (lParamNum == 0) {
			ss = "CJuncter::GetParamDefValue, someMethod, lParamNum = 0";
		} else {
			ss = "CJuncter::GetParamDefValue, someMethod, lParamNum = " + std::to_string(lParamNum);
		}
		TRC("1.txt", ss);

		break; // ���������� � ������ ���.

	case longFunc:

		ss = "CJuncter::GetParamDefValue, longFunc = " + std::to_string(lMethodNum);
		TRC("1.txt", ss);

		if (lParamNum == 0) {
			ss = "CJuncter::GetParamDefValue, longFunc, lParamNum = 0";
		} else {
			ss = "CJuncter::GetParamDefValue, longFunc, lParamNum = " + std::to_string(lParamNum);
		}

		// � ������ 1 ��������, ������� �������� �� ��������� 0.
		// ���������� ����, ������� ����� ��������� � VARIANT.
		V_VT(pvarParamDefValue) = VT_I4;

		// ��������� �������� VARIANT.
		V_I4(pvarParamDefValue) = 0;

		break;

	case doubleFunc:

		ss = "CJuncter::GetParamDefValue, doubleFunc = " + std::to_string(lMethodNum);
		TRC("1.txt", ss);

		if (lParamNum == 0) {
			ss = "CJuncter::GetParamDefValue, doubleFunc, lParamNum = 0";
		} else {
			ss = "CJuncter::GetParamDefValue, doubleFunc, lParamNum = " + std::to_string(lParamNum);
		}

		V_VT(pvarParamDefValue) = VT_R8;
		V_R8(pvarParamDefValue) = 0;

		break;

	case stringFunc:

		ss = "CJuncter::GetParamDefValue, stringFunc = " + std::to_string(lMethodNum);
		TRC("1.txt", ss);

		if (lParamNum == 0) {
			ss = "CJuncter::GetParamDefValue, stringFunc, lParamNum = 0";
		} else {
			ss = "CJuncter::GetParamDefValue, stringFunc, lParamNum = " + std::to_string(lParamNum);
		}

		V_VT(pvarParamDefValue) = VT_BSTR;
		V_BSTR(pvarParamDefValue) = SysAllocString(L"");

		break;

	case dateFunc:

		ss = "CJuncter::GetParamDefValue, dateFunc = " + std::to_string(lMethodNum);
		TRC("1.txt", ss);

		if (lParamNum == 0) {
			ss = "CJuncter::GetParamDefValue, dateFunc, lParamNum = 0";
		} else {
			ss = "CJuncter::GetParamDefValue, dateFunc, lParamNum = " + std::to_string(lParamNum);
		}


		V_VT(pvarParamDefValue) = VT_DATE;
		V_DATE(pvarParamDefValue) = 0; // ��� ������������ ����?

		break;

	case dateTimeFunc:

		ss = "CJuncter::GetParamDefValue, dateTimeFunc = " + std::to_string(lMethodNum);
		TRC("1.txt", ss);

		if (lParamNum == 0) {
			ss = "CJuncter::GetParamDefValue, dateTimeFunc, lParamNum = 0";
		} else {
			ss = "CJuncter::GetParamDefValue, dateTimeFunc, lParamNum = " + std::to_string(lParamNum);
		}

		V_VT(pvarParamDefValue) = VT_DATE;
		V_DATE(pvarParamDefValue) = 0;

		break;

	case arrayFunc:

		ss = "CJuncter::GetParamDefValue, arrayFunc = " + std::to_string(lMethodNum);
		TRC("1.txt", ss);

		if (lParamNum == 0) {
			ss = "CJuncter::GetParamDefValue, arrayFunc, lParamNum = 0";
		} else {
			ss = "CJuncter::GetParamDefValue, arrayFunc, lParamNum = " + std::to_string(lParamNum);
		}

		V_VT(pvarParamDefValue) = VT_ARRAY;
		V_ARRAY(pvarParamDefValue) = 0;

		break;

	case twoParamsFunc:

		ss = "CJuncter::GetParamDefValue, twoParamsFunc = " + std::to_string(lMethodNum);
		TRC("1.txt", ss);

		// ������ �������� - ������, ������ - ����.
		if (lParamNum == 0) {

			ss = "CJuncter::GetParamDefValue, twoParamsFunc, lParamNum = 0";

			V_VT(pvarParamDefValue) = VT_BSTR;
			V_BSTR(pvarParamDefValue) = SysAllocString(L"");

		} else if (lParamNum == 1) {

			ss = "CJuncter::GetParamDefValue, twoParamsFunc, lParamNum = " + std::to_string(lParamNum);

			V_VT(pvarParamDefValue) = VT_DATE;
			V_DATE(pvarParamDefValue) = 0;

		} else {

			ss = "CJuncter::GetParamDefValue, twoParamsFunc, unknown lParamNum = " + std::to_string(lParamNum);
			return S_FALSE; // ����� ��� �������� ������ �����������.
		}

		TRC("1.txt", ss);

		break;
	
	case callBackFunc:
		
		if (someMethod == 0) {
			ss = "CJuncter::GetParamDefValue, callBackFunc = 0";
		} else {
			ss = "CJuncter::GetParamDefValue, callBackFunc = " + std::to_string(lMethodNum);
		}
		TRC("1.txt", ss);

		if (lParamNum == 0) {
			ss = "CJuncter::GetParamDefValue, callBackFunc, lParamNum = 0";
		} else {
			ss = "CJuncter::GetParamDefValue, callBackFunc, lParamNum = " + std::to_string(lParamNum);
		}
		TRC("1.txt", ss);

		break; // ���������� � ������ ���.

	default:

		ss = "CJuncter::GetParamDefValue, default someMethod = " + std::to_string(lMethodNum);
		TRC("1.txt", ss);

		return S_FALSE; // ����� ��� �������� ������ �����������.
	}

	// ���������� ����, ������� ����� ��������� � ���������� pvarPropVal.
	//V_VT(pvarParamDefValue) = VT_I4;
	//V_I4(pvarParamDefValue) = 0;

	ss = "CJuncter::GetParamDefValue, before return S_OK;";
	TRC("1.txt", ss);
	TRC("1.txt", pvarParamDefValue);
	TRC("1.txt", V_I4(pvarParamDefValue));

	return S_OK; // �������� ��������� �������.
}

// ���������� ���� ������� ������������� �������� � ������ � ���������� ������� lMethodNum.
//
STDMETHODIMP CJuncter::HasRetVal(long lMethodNum, BOOL *pboolRetValue)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState())

		std::string ss ("CJuncter::HasRetVal");
	TRC("1.txt", ss);
	TRC("1.txt", lMethodNum);
	TRC("1.txt", pboolRetValue);
	TRC("1.txt", *pboolRetValue);

	// �� ������ ������!!!
	*pboolRetValue = -1;

	switch(lMethodNum)
	{ 
	case someMethod:

		if (someMethod == 0) {
			ss = "CJuncter::HasRetVal, someMethod = 0";
		} else {
			ss = "CJuncter::HasRetVal, someMethod = " + std::to_string(lMethodNum);
		}
		TRC("1.txt", ss);


		*pboolRetValue = FALSE; // ����� �������� �� ���������� (�.�. �������� ����������).

		break;

	case longFunc:

		if (lMethodNum == 0) {
			ss = "CJuncter::HasRetVal, longFunc = 0";
		} else {
			ss = "CJuncter::HasRetVal, longFunc = " + std::to_string(lMethodNum);
		}
		TRC("1.txt", ss);

		*pboolRetValue = TRUE; // ����� ���������� �������� (�.�. �������� ��������).

		break;

	case doubleFunc:

		if (lMethodNum == 0) {
			ss = "CJuncter::HasRetVal, doubleFunc = 0";
		} else {
			ss = "CJuncter::HasRetVal, doubleFunc = " + std::to_string(lMethodNum);
		}
		TRC("1.txt", ss);

		*pboolRetValue = TRUE;

		break;

	case stringFunc:

		if (lMethodNum == 0) {
			ss = "CJuncter::HasRetVal, stringFunc = 0";
		} else {
			ss = "CJuncter::HasRetVal, stringFunc = " + std::to_string(lMethodNum);
		}
		TRC("1.txt", ss);

		*pboolRetValue = TRUE;

		break;

	case dateFunc:

		if (lMethodNum == 0) {
			ss = "CJuncter::HasRetVal, dateFunc = 0";
		} else {
			ss = "CJuncter::HasRetVal, dateFunc = " + std::to_string(lMethodNum);
		}
		TRC("1.txt", ss);

		*pboolRetValue = TRUE;

		break;

	case dateTimeFunc:

		if (lMethodNum == 0) {
			ss = "CJuncter::HasRetVal, dateTimeFunc = 0";
		} else {
			ss = "CJuncter::HasRetVal, dateTimeFunc = " + std::to_string(lMethodNum);
		}
		TRC("1.txt", ss);

		*pboolRetValue = TRUE;

		break;

	case arrayFunc:

		if (lMethodNum == 0) {
			ss = "CJuncter::HasRetVal, arrayFunc = 0";
		} else {
			ss = "CJuncter::HasRetVal, arrayFunc = " + std::to_string(lMethodNum);
		}
		TRC("1.txt", ss);

		*pboolRetValue = TRUE;

		break;

	case twoParamsFunc:

		if (lMethodNum == 0) {
			ss = "CJuncter::HasRetVal, twoParamsFunc = 0";
		} else {
			ss = "CJuncter::HasRetVal, twoParamsFunc = " + std::to_string(lMethodNum);
		}
		TRC("1.txt", ss);

		*pboolRetValue = TRUE;

		break;

	case callBackFunc:

		if (someMethod == 0) {
			ss = "CJuncter::HasRetVal, callBackFunc = 0";
		} else {
			ss = "CJuncter::HasRetVal, callBackFunc = " + std::to_string(lMethodNum);
		}
		TRC("1.txt", ss);

		*pboolRetValue = TRUE;

		break;


	default:

		ss = "CJuncter::HasRetVal, default someMethod/someFunction = " + std::to_string(lMethodNum);
		TRC("1.txt", ss);

		return S_FALSE;
	}

	ss = "CJuncter::HasRetVal, before return S_OK;";
	TRC("1.txt", ss);
	TRC("1.txt", pboolRetValue);
	TRC("1.txt", *pboolRetValue);

	return S_OK;
}

// ��������� ����� � ���������� ������� lMethodNum.
// paParams - ������� ��������� �� ������ �������� VARIANT, ���������� �������� ���������� ������.
//
STDMETHODIMP CJuncter::CallAsProc(long lMethodNum, SAFEARRAY **paParams)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState())

		USES_CONVERSION;

	std::string ss ("CJuncter::CallAsProc");
	TRC("1.txt", ss);

	switch(lMethodNum)
	{ 
	case someMethod: // = 0

		// TODO: ����� ��� ������ someMethod.
		ss = "CJuncter::CallAsProc - someMethod!!!";
		TRC("1.txt", ss);

		m_someBooleanPropertyValue = TRUE; // ����� ������������� ������ �������� � TRUE.

		break;

	default:
		return S_FALSE;
	}

	ShowSafeArray("1.txt", paParams);

	return S_OK; // ����� ������ ��� ������.
	//return E_FAIL; // ����� ������, ��������� ������ ������� ���������� (���������� ������ 1� ������������).
	//return S_FALSE; // ����������� �����, ��������������� lMethodNum.
}

// ����� ���������������� �������, ������������ � 1�.
//
STDMETHODIMP CallUserFunc(IDispatch* iDisp, wchar_t* methodName)
{

	//wchar_t* methodName1 = L"myFunc1"; // ��������� � ������ ����� ���������.
	DISPID dispid = NULL;
	HRESULT hr = iDisp->GetIDsOfNames(
		IID_NULL,              // ��������������, ������ IID_NULL.
		&methodName,           // ��� (�����) ������ ��� ��������.
		1,                     // ���������� ������������� ����.
		LOCALE_SYSTEM_DEFAULT, // ������������� �����.
		&dispid               // �������� ��� �������� DISPID ������.
		);

	if (SUCCEEDED(hr)) {

		VARIANT vRetVal;
		VariantInit(&vRetVal);
		int ParmCount = 0;
		VARIANTARG vParams;
		VariantInit(&vParams);

		DISPPARAMS dParams = {&vParams, NULL, ParmCount, 0}; // �������� ��������� ����������

		// ������� ���
		return iDisp->Invoke(dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT, DISPATCH_METHOD, &dParams, &vRetVal, NULL, 0);

	} else {
		return S_FALSE;
	}
}

// ��������� ����� � ���������� ������� lMethodNum.
// pvarRetValue - ��������� �� ��������� VARIANT, ���������� ������������ ��������.
// paParams - ������� ��������� �� ������ �������� VARIANT, ���������� �������� ���������� ������.
//
STDMETHODIMP CJuncter::CallAsFunc(long lMethodNum, VARIANT *pvarRetValue, SAFEARRAY **paParams)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState())

		USES_CONVERSION;

	std::string ss ("CJuncter::CallAsFunc");
	TRC("1.txt", ss);

	switch(lMethodNum)
	{ 
	case longFunc: // = 1
		{
			// TODO: ����� ��� ������� someFunction.
			ss = "CJuncter::CallAsFunc - longFunc!!!";
			TRC("1.txt", ss);

			// � ������� ���� ������������ ��������.
			// ���������� ����, ������� ����� ��������� � VARIANT.
			V_VT(pvarRetValue) = VT_I4;

			// ��������� �������� VARIANT.
			VARIANT var = GetNParam(*paParams, 0);
			V_VT(&var) = VT_I4;
			V_I4(pvarRetValue) = V_I4(&var);

			::VariantClear(&var);

			break;
		}
	case doubleFunc:
		{
			ss = "CJuncter::CallAsFunc - doubleFunc!!!";
			TRC("1.txt", ss);

			// ���������� ����.
			V_VT(pvarRetValue) = VT_R8;

			// ��������� �������� VARIANT.

			long iLBound; // ������ �������.
			HRESULT hr = SafeArrayGetLBound(*paParams, 1, &iLBound);
			if(FAILED(hr))
				return S_FALSE;

			VARIANT iVal; // ��������.

			VariantInit(&iVal); // ������������ �������������.

			// ��������� ��������, ������������ � ������ �������� �������.
			hr = SafeArrayGetElement(*paParams, &iLBound, &iVal); // ��� ����������� �������������!
			if(FAILED(hr))
				return S_FALSE;

			// ������������ ��������.
			V_R8(pvarRetValue) = V_R8(&iVal);

			VariantClear(&iVal); // ������������ �������.

			break;
		}
	case stringFunc:
		{
			ss = "CJuncter::CallAsFunc - stringFunc!!!";
			TRC("1.txt", ss);

			// ��������� �������� �� ����������� �������.
			VARIANT var = GetNParam(*paParams, 0); // ��� ����������� �������������!

			// ���������� ���� ������������� ��������.
			V_VT(pvarRetValue) = VT_BSTR;

			// ��������� �������� ������������� ��������.
			V_BSTR(pvarRetValue) = V_BSTR(&var);

			::VariantClear(&var);

			break;
		}
	case dateFunc:

		{
			ss = "CJuncter::CallAsFunc - dateFunc!!!";
			TRC("1.txt", ss);

			VARIANT var = GetNParam(*paParams, 0);
			V_VT(pvarRetValue) = VT_DATE;
			V_DATE(pvarRetValue) = V_DATE(&var);

			break;
		}

	case dateTimeFunc:

		{
			ss = "CJuncter::CallAsFunc - dateTimeFunc!!!";
			TRC("1.txt", ss);

			VARIANT var = GetNParam(*paParams, 0);
			V_VT(pvarRetValue) = VT_DATE;
			V_DATE(pvarRetValue) = V_DATE(&var);

			break;
		}

	case arrayFunc:

		{
			ss = "CJuncter::CallAsFunc - arrayFunc!!!";
			TRC("1.txt", ss);

			// **********************************
			// *** ������ � �������� �������� ***
			// **********************************
			// ��������� ��������� �� ������� ������� ������� SAFEARRAY
			// (�������� ������ �������� ������� 1�, ������� ���� �������� ��������).
			long inputIdx = 0;
			void* inputVoidPtr = NULL;
			HRESULT hr = SafeArrayPtrOfIndex(*paParams, &inputIdx, &inputVoidPtr);

			// ���������� ����������� ��������� void* � ���� VARIANT*
			// (�.�. �� ����� �����, ��� ��� ������ VARIANT).
			VARIANT* inputVarPtr = (VARIANT*)inputVoidPtr;

			///////////// ����������� ������ �������.
			long iLowerBound;
			hr = SafeArrayGetLBound(inputVarPtr->parray, 1, &iLowerBound);
			if(FAILED(hr))
				return S_FALSE;

			long iUpperBound;
			hr = SafeArrayGetUBound(inputVarPtr->parray, 1, &iUpperBound);
			if(FAILED(hr))
				return S_FALSE;

			//////////////////////////////////////////////

			void* sourcePtr  = NULL; // ��������� �� �������, ����������� �������� SafeArrayPtrOfIndex.
			double* sourceValPtr = NULL;
			// ������� �������� ������� �� ������ ������� � �������.
			for(long l = iLowerBound; l <= iUpperBound; l++)
			{
				// ���������� ��������� �� �������.
				hr = SafeArrayPtrOfIndex(inputVarPtr->parray, &l, &sourcePtr);


				// ���������� ���������� void* � ���������� ����.
				sourceValPtr = (double*)sourcePtr;

				// ��������� �������� �� ������� �������.
				++(*sourceValPtr);
			}


			// *******************************
			// *** ������ � ������ ������� ***
			// *******************************

			// ��������� � ���������� ����� �������� �������
			// (�.�. � ������� ���������� ������� SafeArrayGetElement,
			// ������� ��������� ����������� �������� ������� SAFEARRAY).
			VARIANT var = GetNParam(*paParams, 0);

			// ���������, ����� ������� �������������� ������ � ��������
			// (������������� ��� ��������).
			SAFEARRAY* sArr = V_ARRAY(&var); // �� �� �����, ��� �: SAFEARRAY* sArr = var.parray

			/////////// ����������� ������ �������.
			long iLBound; // ������ �������.
			hr = SafeArrayGetLBound(sArr, 1, &iLBound);
			if(FAILED(hr))
				return S_FALSE;

			long iUBound; // ������� �������.
			hr = SafeArrayGetUBound(sArr, 1, &iUBound);
			if(FAILED(hr))
				return S_FALSE;
			////////////////////////////////////////////


			///////////// �������� ������ ������� (� ����� �� ������������).
			// ������� ������ �������
			SAFEARRAYBOUND sBound[1];
			sBound[0].cElements = iUBound - iLBound + 1;
			sBound[0].lLbound = 0;

			// ��������� ����, ����������� � �������� �������.
			VARTYPE varType;
			hr = SafeArrayGetVartype(sArr, &varType);
			if(FAILED(hr))
				return S_FALSE;

			// �������� ������ �������.
			SAFEARRAY* sArrNew = SafeArrayCreate(varType, 1, sBound);
			//////////////////////////////////////////////


			/////////// ������� �������.
			void* sourPtr  = NULL; // ��������� �� �������, ����������� �������� SafeArrayPtrOfIndex.
			void* destPtr  = NULL; // ��������� �� �������, ����������� �������� SafeArrayPtrOfIndex.

			// ������� ��������� ������� �� ������ ������� � �������.
			for(long l = iLBound; l <= iUBound; l++)
			{
				// ���������� ��������� �� �������.
				hr = SafeArrayPtrOfIndex(sArr, &l, &sourPtr);

				hr = SafeArrayPtrOfIndex(sArrNew, &l, &destPtr);

				// ���������� ���������� void* � ���������� ����
				// � ���������� �������� �������� ������� (��������, ���������� �� 2).
				*((double*)destPtr) = *((double*)sourPtr) * 2;
			}
			////////////////////////////////////////////

			///////////// ���������� ������������� �������� �������� �������.
			V_VT(pvarRetValue) = VT_ARRAY;
			V_ARRAY(pvarRetValue) = sArrNew;

			break;
		}

	case twoParamsFunc:

		{
			ss = "CJuncter::CallAsFunc - twoParamsFunc!!!";
			TRC("1.txt", ss);

			// �������� ����� ���������� �������� � ����������� ���������� ����,
			// �������� � ���������� � ���� ��������������� ��������.

			// ������ �������� - ������.
			VARIANT var1 = GetNParam(*paParams, 0);

			///////////////////////////////////////////
			//// ����������� CString � BSTR:
			//CString cs("Hello");
			//BSTR bstr = cs.AllocSysString();

			///////////////////////////////////////////
			//// ����������� BSTR � CString:
			//BSTR bstr = ::SysAllocString(L"Hello");
			//_bstr_t tmp(bstr, FALSE);   //wrap the BSTR
			//CString cs(static_cast<const char*>(tmp));  //convert it
			//AfxMessageBox(cs, MB_OK, 0);

			//// ���:
			//BSTR bstr1;
			//bstr1 = ::SysAllocString(L"Hello1");
			//CString cs1(_com_util::ConvertBSTRToString(bstr1));
			//AfxMessageBox(cs1, MB_OK, 0);
			///////////////////////////////////////////

			// The _bstr_t class encapsulates and manages the BSTR data type.
			_bstr_t tmp1(V_BSTR(&var1), FALSE);
			CString cs1(static_cast<const char*>(tmp1));
			
			std::string stringParam1(cs1);

			// ������ �������� - ����.
			VARIANT var2 = GetNParam(*paParams, 1);

			///////////////////////////////////////////
			// http://msdn.microsoft.com/en-us/library/windows/desktop/ms221508%28v=vs.85%29.aspx
			//Converts a date value to a BSTR value.
			//HRESULT VarBstrFromDate(
			//	_In_   DATE dateIn,
			//	_In_   LCID lcid,
			//	_In_   ULONG dwFlags,
			//	_Out_  BSTR *pbstrOut
			//	);
			BSTR pbstrOut;
			HRESULT hr = VarBstrFromDate(V_DATE(&var2), LOCALE_CUSTOM_DEFAULT, VAR_FOURDIGITYEARS, &pbstrOut);

			_bstr_t tmp2(pbstrOut, FALSE);
			CString cs2(static_cast<const char*>(tmp2));
			
			std::string stringParam2(cs2);

			// ��������.
			cs1 = cs1 + ", " + cs2 + ".";
			
			// ����������.
			V_VT(pvarRetValue) = VT_BSTR;
			V_BSTR(pvarRetValue) = cs1.AllocSysString();

			// ������������ ������.
			::VariantClear(&var1);
			::VariantClear(&var2);
			
			::SysFreeString(pbstrOut);

			break;
		}

	case callBackFunc:

		{
			ss = "CJuncter::CallAsFunc - callBackFunc!!!";
			TRC("1.txt", ss);

			///////////// �������� ������ ������� �� 3 ���������.
			// ������� ������ �������
			SAFEARRAYBOUND sBound[1];
			sBound[0].cElements = 3;
			sBound[0].lLbound = 0;

			// �������� ������ ������� c� ���������� ���������� ����.
			SAFEARRAY* sArrNew = SafeArrayCreate(VT_BSTR, 1, sBound);
			//SAFEARRAY* sArrNew = SafeArrayCreate(VT_I4, 1, sBound);
			//////////////////////////////////////////////

			// ���������� ������� ����������.
			void* sourPtr  = NULL;

			LONG l = 0;
			HRESULT hr = SafeArrayPtrOfIndex(sArrNew, &l, &sourPtr);
			
			// ���������� ���������� void* � ���������� ����
			// � ���������� �������� �������� �������.
			*((BSTR*)sourPtr) = SysAllocString(L"TICKER");
			//*((LONG*)sourPtr) = 10;

			l = 1;
			hr = SafeArrayPtrOfIndex(sArrNew, &l, &sourPtr);
			*((BSTR*)sourPtr) = SysAllocString(L"DATA");
			//*((LONG*)sourPtr) = 20;

			l = 2;
			hr = SafeArrayPtrOfIndex(sArrNew, &l, &sourPtr);
			*((BSTR*)sourPtr) = SysAllocString(L"VALUE");
			//*((LONG*)sourPtr) = 30;

			///////////// ���������� ������������� �������� �������� �������.
			V_VT(pvarRetValue) = VT_ARRAY;
			V_ARRAY(pvarRetValue) = sArrNew;


			//// ����� �������� ������� � 1�.

			// � ����� ���:
			// ���������� � ������ Init ��������� �� IDispatch ��������� �������� ������ � 1�:�����������
			// ����� �������� OLE Automation. �� ����������� ��������� ����� �������� �������� AppDispatch,
			// ��������� ������ ��� ������.
			// ��� �������� �������� ��������� �� IDispatch 1�:����������� (�� ������� � ���������� � Init)!
			// �������� AppDispatch ���������� �������� ������ ����� ������ ������������� ���� ������� 1�:�����������,
			// ������� � ������ �������� ������� ���������� � ������ ������ Init ��� �������� ������������
			// ������ �� �� ���� ������������ 1�:�����������.

			// � ����� ���:
			// ��� ������� � 1�:����������� � ������ 2.0 ��������� ����������� �������� AppDispatch. ���������� AppDispatch ������������ ����� �������,
			// ���������� ���������� ����� OLE Automation, �������, ��������� AppDispatch, ����� �������� ���������
			// � ������� ����������� ������ 1�:�����������. ���������� ���� �����������, ��������� � ���� ����������:
			// ����� ��������� AppDispatch ���������� ������������� �������� AddRef � ����������� IDispatch - � ���������
			// ������ ����� ���������� ������ �������� "���������" 1�:����������� � ������. ��� ����������� �.�. ������
			// ������������� ������� ���������� ��� ������ � 1�:������������, ������� � ������ 7.70.006.

			// � ������:
			// https://partners.v8.1c.ru/forum/message/240854#m_240854
			// ��������� � ������� ����� ������� �������� ��� ������� ����� ������� AppDispatch,
			// ���� � �� �������� ������������ �������� ����� ������� � ������, � ������� ��� ���������, ������� ��� ���������� ��� ������.

			// �.�. � ��� ���� ��� ��������� IDispatch:
			// 1. ������������ �������� � �������� ��������� ������ Init. ����� ���� ����� �������� ����������
			// 1�:�����������, �� ������ �������� ��� ������;
			// 2. ����������� � �������� AppDispach. �� ����� �������������� ��� ������ ������� 1�:�����������,
			// ��� ���� ������������ �������� OLE Automation.

			// ��������� ���� ��������� IDispatch, ������� ��������� �������� ������ 1�:
			 
			// ����������, ����������� ��������� ��������� � ������� ������� GetPropertyByName.
			_variant_t var;
			
			// ��������� �� IDispatch, ����������� ���������� � 1�:����������� � �����.
			// ��������, ����� ���� ����� ������� ���������� ���������, ����������� � ����� ������, ����������� �� �������.
			CComPtr<IDispatch> appConnect = NULL;

			// ��������� �� IDispatch, ����������� ���������� � ���������� � �����.
			// ������������� ��������, ����� ��������� ��������� IDispatch ���������� ���������, �� �����.
			CComPtr<IDispatch> execConnect = NULL;
									
			// IDispatch ��� ���������� ��������� � ������ "COM".
			CComPtr<IDispatch> myExecConnect = NULL;
										
			// IDispatch ��� ���������� ����� "�����" ���������� ��������� "COM".
			CComPtr<IDispatch> myExecFormConnect = NULL;
			
			hr = m_iConnect.GetPropertyByName(L"AppDispatch", var.GetAddress());
			if (SUCCEEDED(hr)) {

				// �������� ������������ ����.
				ATLASSERT(V_VT(&var) == VT_DISPATCH);

				appConnect = V_DISPATCH(&var);

				hr = appConnect.GetPropertyByName(L"���������", var.GetAddress());
				if (SUCCEEDED(hr)) {

					execConnect = V_DISPATCH(&var);

					hr = execConnect.GetPropertyByName(L"COM", var.GetAddress());
					if (SUCCEEDED(hr)) {

						myExecConnect = V_DISPATCH(&var);

						// � 1� ��� �������� ������������ ����� ������������ �������
						// �������������(<�����>, <��������>, <���� ������������>).
						// ��������������, ��� ����, ����� ������� ��� ������� ������� ���������
						// � ���������.
						// ��� ���� � ������� ���������� ������� ��������� �.�. ��������.

						// ������ ����������.
						_variant_t vars[3];

						// ��� �����.
						BSTR formName = SysAllocString(L"�����1");
						V_VT(&vars[2]) = VT_BSTR;
						V_BSTR(&vars[2]) = formName;

						// �������� �����.
						V_VT(&vars[1]) = VT_ERROR;

						// �������� ����� ������������.
						UUID id;
						::UuidCreate(&id);
						
						BYTE* strId = NULL;
						::UuidToString(&id, &strId);

						// ���� ������������.
						vars[0] = T2OLE((LPCTSTR)strId);

						// ������������ ������.
						::SysFreeString(formName);
						::RpcStringFree(&strId);

						// ����� ������� "�������������" ������� ���������.
						hr = myExecConnect.InvokeN(L"�������������", vars, 3, var.GetAddress());
						if (SUCCEEDED(hr))
						{
							myExecFormConnect = V_DISPATCH(&var);

							// ����� ������� "�������" �����.
							hr = myExecFormConnect.Invoke0(L"�������");
						}
					}
				}

				// ������� �������� �������, ��������� ������ ��������� IDispatch.
				/////////////////////////////////////////////////
				// ��������� �� 1�.
				BOOL success = FALSE;
				hr = CallUserFunc(appConnect, L"myFunc1");  // S_FALSE
				hr = CallUserFunc(appConnect, L"myFunc11"); // S_FALSE
				hr = CallUserFunc(appConnect, L"myFunc2");  // S_FALSE
				hr = CallUserFunc(appConnect, L"myFunc22"); // S_FALSE
				hr = CallUserFunc(appConnect, L"myFunc3");  // S_OK
				hr = CallUserFunc(appConnect, L"myFunc33"); // S_FALSE
				hr = CallUserFunc(appConnect, L"myFunc4");  // S_FALSE
				hr = CallUserFunc(appConnect, L"myFunc44"); // S_FALSE
				hr = CallUserFunc(appConnect, L"myFunc5");  // S_FALSE
				hr = CallUserFunc(appConnect, L"myFunc55"); // S_FALSE
				
				// VARIANT vRetVal;
				// ���� ������� ���������� �������� ���� ���:
				// VARIANT vMem;
				// vMem.vt = VT_DISPATCH;
				// ���� �� ���������� �������� � �������� ��������� NULL.
				
				// VARIANTARG vParams;
				// ���� ������� �� ��������� ��������� �� ���� ������� NULL.
				// ��� ���������� �������� ������ ���������� � NULL �� ������ ������������.
				
				// � �������� ����� ���� ���������� ���� DISPATCH_METHOD, ���� DISPATCH_PROPERTYGET L
				hr = ExecuteMethod_1C(appConnect, "myFunc1", NULL, NULL, 0, DISPATCH_METHOD);  // E_NOTIMP
				hr = ExecuteMethod_1C(appConnect, "myFunc11", NULL, NULL, 0, DISPATCH_METHOD); // E_NOTIMP
				hr = ExecuteMethod_1C(appConnect, "myFunc2", NULL, NULL, 0, DISPATCH_METHOD);  // E_NOTIMP
				hr = ExecuteMethod_1C(appConnect, "myFunc22", NULL, NULL, 0, DISPATCH_METHOD); // E_NOTIMP
				hr = ExecuteMethod_1C(appConnect, "myFunc3", NULL, NULL, 0, DISPATCH_METHOD);  // S_OK
				hr = ExecuteMethod_1C(appConnect, "myFunc33", NULL, NULL, 0, DISPATCH_METHOD); // E_NOTIMP
				hr = ExecuteMethod_1C(appConnect, "myFunc4", NULL, NULL, 0, DISPATCH_METHOD);  // E_NOTIMP
				hr = ExecuteMethod_1C(appConnect, "myFunc44", NULL, NULL, 0, DISPATCH_METHOD); // E_NOTIMP
				hr = ExecuteMethod_1C(appConnect, "myFunc5", NULL, NULL, 0, DISPATCH_METHOD);  // E_NOTIMP
				hr = ExecuteMethod_1C(appConnect, "myFunc55", NULL, NULL, 0, DISPATCH_METHOD); // E_NOTIMP

				/////////////////////////////////////////////////
				// ��������� �� ���������.
				hr = CallUserFunc(execConnect, L"myFunc1");  // S_FALSE
				hr = CallUserFunc(execConnect, L"myFunc11"); // S_FALSE
				hr = CallUserFunc(execConnect, L"myFunc2");  // S_FALSE
				hr = CallUserFunc(execConnect, L"myFunc22"); // S_FALSE
				hr = CallUserFunc(execConnect, L"myFunc3");  // S_FALSE
				hr = CallUserFunc(execConnect, L"myFunc33"); // S_FALSE
				hr = CallUserFunc(execConnect, L"myFunc4");  // S_FALSE
				hr = CallUserFunc(execConnect, L"myFunc44"); // S_FALSE
				hr = CallUserFunc(execConnect, L"myFunc5");  // S_FALSE
				hr = CallUserFunc(execConnect, L"myFunc55"); // S_FALSE
				
				hr = ExecuteMethod_1C(execConnect, "myFunc1", NULL, NULL, 0, DISPATCH_METHOD);  // E_NOTIMP
				hr = ExecuteMethod_1C(execConnect, "myFunc11", NULL, NULL, 0, DISPATCH_METHOD); // E_NOTIMP
				hr = ExecuteMethod_1C(execConnect, "myFunc2", NULL, NULL, 0, DISPATCH_METHOD);  // E_NOTIMP
				hr = ExecuteMethod_1C(execConnect, "myFunc22", NULL, NULL, 0, DISPATCH_METHOD); // E_NOTIMP
				hr = ExecuteMethod_1C(execConnect, "myFunc3", NULL, NULL, 0, DISPATCH_METHOD);  // E_NOTIMP
				hr = ExecuteMethod_1C(execConnect, "myFunc33", NULL, NULL, 0, DISPATCH_METHOD); // E_NOTIMP
				hr = ExecuteMethod_1C(execConnect, "myFunc4", NULL, NULL, 0, DISPATCH_METHOD);  // E_NOTIMP
				hr = ExecuteMethod_1C(execConnect, "myFunc44", NULL, NULL, 0, DISPATCH_METHOD); // E_NOTIMP
				hr = ExecuteMethod_1C(execConnect, "myFunc5", NULL, NULL, 0, DISPATCH_METHOD);  // E_NOTIMP
				hr = ExecuteMethod_1C(execConnect, "myFunc55", NULL, NULL, 0, DISPATCH_METHOD); // E_NOTIMP

				
				/////////////////////////////////////////////////
				// ��������� �� ���������� ���������.
				hr = CallUserFunc(myExecConnect, L"myFunc1");  // S_FALSE
				hr = CallUserFunc(myExecConnect, L"myFunc11"); // S_FALSE
				hr = CallUserFunc(myExecConnect, L"myFunc2");  // S_FALSE
				hr = CallUserFunc(myExecConnect, L"myFunc22"); // S_FALSE
				hr = CallUserFunc(myExecConnect, L"myFunc3");  // S_FALSE
				hr = CallUserFunc(myExecConnect, L"myFunc33"); // S_FALSE
				hr = CallUserFunc(myExecConnect, L"myFunc4");  // S_FALSE
				hr = CallUserFunc(myExecConnect, L"myFunc44"); // S_FALSE
				hr = CallUserFunc(myExecConnect, L"myFunc5");  // S_FALSE
				hr = CallUserFunc(myExecConnect, L"myFunc55"); // S_FALSE
				
				hr = ExecuteMethod_1C(myExecConnect, "myFunc1", NULL, NULL, 0, DISPATCH_METHOD);  // E_NOTIMP
				hr = ExecuteMethod_1C(myExecConnect, "myFunc11", NULL, NULL, 0, DISPATCH_METHOD); // E_NOTIMP
				hr = ExecuteMethod_1C(myExecConnect, "myFunc2", NULL, NULL, 0, DISPATCH_METHOD);  // E_NOTIMP
				hr = ExecuteMethod_1C(myExecConnect, "myFunc22", NULL, NULL, 0, DISPATCH_METHOD); // E_NOTIMP
				hr = ExecuteMethod_1C(myExecConnect, "myFunc3", NULL, NULL, 0, DISPATCH_METHOD);  // E_NOTIMP
				hr = ExecuteMethod_1C(myExecConnect, "myFunc33", NULL, NULL, 0, DISPATCH_METHOD); // E_NOTIMP
				hr = ExecuteMethod_1C(myExecConnect, "myFunc4", NULL, NULL, 0, DISPATCH_METHOD);  // E_NOTIMP
				hr = ExecuteMethod_1C(myExecConnect, "myFunc44", NULL, NULL, 0, DISPATCH_METHOD); // E_NOTIMP
				hr = ExecuteMethod_1C(myExecConnect, "myFunc5", NULL, NULL, 0, DISPATCH_METHOD);  // E_NOTIMP
				hr = ExecuteMethod_1C(myExecConnect, "myFunc55", NULL, NULL, 0, DISPATCH_METHOD); // E_NOTIMP

				/////////////////////////////////////////////////
				// ��������� �� ���������� ����� ���������� ���������.
				hr = CallUserFunc(myExecFormConnect, L"myFunc1");  // S_FALSE
				hr = CallUserFunc(myExecFormConnect, L"myFunc11"); // S_FALSE
				hr = CallUserFunc(myExecFormConnect, L"myFunc2");  // S_OK
				hr = CallUserFunc(myExecFormConnect, L"myFunc22"); // S_FALSE
				hr = CallUserFunc(myExecFormConnect, L"myFunc3");  // S_FALSE
				hr = CallUserFunc(myExecFormConnect, L"myFunc33"); // S_FALSE
				hr = CallUserFunc(myExecFormConnect, L"myFunc4");  // S_FALSE
				hr = CallUserFunc(myExecFormConnect, L"myFunc44"); // S_FALSE
				hr = CallUserFunc(myExecFormConnect, L"myFunc5");  // S_FALSE
				hr = CallUserFunc(myExecFormConnect, L"myFunc55"); // S_FALSE
				
				hr = ExecuteMethod_1C(myExecFormConnect, "myFunc1", NULL, NULL, 0, DISPATCH_METHOD);  // E_NOTIMP
				hr = ExecuteMethod_1C(myExecFormConnect, "myFunc11", NULL, NULL, 0, DISPATCH_METHOD); // E_NOTIMP
				hr = ExecuteMethod_1C(myExecFormConnect, "myFunc2", NULL, NULL, 0, DISPATCH_METHOD);  // S_OK
				hr = ExecuteMethod_1C(myExecFormConnect, "myFunc22", NULL, NULL, 0, DISPATCH_METHOD); // E_NOTIMP
				hr = ExecuteMethod_1C(myExecFormConnect, "myFunc3", NULL, NULL, 0, DISPATCH_METHOD);  // E_NOTIMP
				hr = ExecuteMethod_1C(myExecFormConnect, "myFunc33", NULL, NULL, 0, DISPATCH_METHOD); // E_NOTIMP
				hr = ExecuteMethod_1C(myExecFormConnect, "myFunc4", NULL, NULL, 0, DISPATCH_METHOD);  // E_NOTIMP
				hr = ExecuteMethod_1C(myExecFormConnect, "myFunc44", NULL, NULL, 0, DISPATCH_METHOD); // E_NOTIMP
				hr = ExecuteMethod_1C(myExecFormConnect, "myFunc5", NULL, NULL, 0, DISPATCH_METHOD);  // E_NOTIMP
				hr = ExecuteMethod_1C(myExecFormConnect, "myFunc55", NULL, NULL, 0, DISPATCH_METHOD); // E_NOTIMP
			}

			break;
		}

	default:
		return S_FALSE;
	}

	ShowSafeArray("1.txt", paParams);
	TRC("1.txt", pvarRetValue);
	TRC("1.txt", V_I4(pvarRetValue));

	return S_OK; // ����� ������ ��� ������.
}

// IPropertyLink Methods
STDMETHODIMP CJuncter::get_Enabled(BOOL * boolEnabled)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState())

		std::string ss ("CJuncter::get_Enabled");
	TRC("1.txt", ss);

	*boolEnabled = m_someBooleanPropertyValue;

	return S_OK;
}

STDMETHODIMP CJuncter::put_Enabled(BOOL boolEnabled)
{

	AFX_MANAGE_STATE(AfxGetStaticModuleState())

		std::string ss ("CJuncter::put_Enabled");
	TRC("1.txt", ss);

	CString csSource("����������"),csMessage("���������"),csData;

	m_someBooleanPropertyValue = boolEnabled;

	csData = boolEnabled?"1":"0";

	m_iAsyncEvent->ExternalEvent(csSource.AllocSysString(),
		csMessage.AllocSysString(),
		csData.AllocSysString());

	return S_OK;
}


