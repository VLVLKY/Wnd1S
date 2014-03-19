// Juncter.cpp: реализация CJuncter

#include "stdafx.h"
#include "Juncter.h"

//#include <Rpcdce.h>
//#include "comutil.h"

// Функция выполняет методы 1С по переданному имени
//
HRESULT ExecuteMethod_1C(IDispatch* pDisp, LPTSTR NameMethod, VARIANT* vRetVal, VARIANTARG* vParams, int ParmCount, WORD Flag)
{
	if(pDisp)
	{
		try
		{
			_bstr_t str(NameMethod);
			
			OLECHAR FAR* szMember = str.GetBSTR(); // приведем имя метода к нужному типу
			DISPID dispid = NULL; // Указатель на метод объекта 1С
			HRESULT res = pDisp->GetIDsOfNames(IID_NULL, &szMember, 1, LOCALE_SYSTEM_DEFAULT, &dispid);
			if(SUCCEEDED(res)) // если метод существует
			{
				DISPPARAMS dParams = {vParams, NULL, ParmCount, 0}; // заполним структуру параметров
				
				// вызовем его
				return pDisp->Invoke(dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT, Flag, &dParams, vRetVal, NULL, 0);
			}
			return E_NOTIMPL;
		}
		catch(_com_error er)
		{
			//ErrorMessage("Ошибка выполнения метода 1С",(LPCSTR)er.Description());
			return S_FALSE;
		}
	}
	return S_FALSE;
}

// Служебная функция, выводит содержимое массива в файл.
//
STDMETHODIMP ShowSafeArray(const char * fileName, SAFEARRAY ** sArr) {

	//typedef struct tagSAFEARRAYBOUND
	//{
	//	unsigned long cElements; // Число элементов в массиве.
	//	long lLbound;            // Нижняя граница массива.
	//} SAFEARRAYBOUND;

	//typedef struct tagSAFEARRAY
	//{
	//	unsigned short cDims;        // Число размерностей в массиве.
	//	unsigned fFeatures;          // Набор флагов, используемый для освобождения памяти массива.
	//	unsigned long cbElements;    // Размер каждого из элементов массива.
	//	unsigned long cLocks;        // Число блокировок массива.
	//	unsigned short handle;       // Не используется, оставлено для совместимости.
	//	void* pvData;                // Указатель на данные массива.
	//	SAFEARRAYBOUND rgsabound[1]; // Структура, описывающая нижняя границу массива и число элементов.
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

// МЕТОДЫ КЛАССА
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
	ASSERT(pArray->fFeatures | FADF_VARIANT); // Проверка того, что в массиве содержится именно VARIANT?

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

// Возвращает строку с учетом языка
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


// МЕТОДЫ ИНТЕРФЕЙСОВ

STDMETHODIMP CJuncter::Init(IDispatch* pConnection)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState())

	std::string ss ("CJuncter::Init ");
	
	TRC("1.txt", ss);
	TRC("1.txt", pConnection);

	// Указатель на интерфейс 1С.
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

	// Компонент должен установить номер версии поддерживаемой компонентой технологии
	// в формате VARIANT по индексу 0.    

	// Индекс, под которым будет занесен элемент в массив.
	long lInd = 0;

	// Переменная, в которой будет храниться значение версии.
	VARIANT varVersion;

	// Установка типа значения, которое будет храниться в поле типа VARIANT.
	V_VT(&varVersion) = VT_I4;

	// Установка значения, которое будет храниться в поле типа VARIANT.
	// Компонент поддерживает версию 2.0, позволяющую создавать несколько объектов одного типа
	// AddIn.<ИмяРасширения>.
	// Если указать версию 1.0, то допускается создание только одного объекта.
	V_I4(&varVersion) = 2000;

	// Установка одиночного элемента в массив.
	SafeArrayPutElement(
		*pInfo,     // [in] Pointer to an array descriptor created by SafeArrayCreate.
		&lInd,      // [in] Pointer to a vector of indexes for each dimension of the array. The right-most (least significant) dimension is rgIndices[0]. The left-most dimension is stored at rgIndices[psa->cDims –1].
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
	//TRC("1.txt", OLE2T(*bstrExtensionName)); - падает, т.к. память еще не выделена!

	CString csExtenderName = "Win1SConnector";
	*bstrExtensionName = csExtenderName.AllocSysString();

	TRC("1.txt", *bstrExtensionName);
	TRC("1.txt", OLE2T(*bstrExtensionName));

	return NULL;
}

// Возвращает количество свойств расширения.
//
STDMETHODIMP CJuncter::GetNProps(long *plProps)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState())

		std::string ss ("CJuncter::GetNProps");
	TRC("1.txt", ss);
	TRC("1.txt", plProps);
	TRC("1.txt", *plProps);

	// LastProp - это всегда последний элемент перечисления, определенного в классе CJuncter.
	*plProps = LastProp;
	TRC("1.txt", *plProps);

	return S_OK;
}

// Возвращает порядковый номер свойства с именем bstrPropName.
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

	// Преобразование в CString.
	CString csPropName = OLE2T(bstrPropName);

	// Сравнение имеющихся свойств (перечислены в строковой таблице) с переданным именем.
	// Строка с именем свойства может состоять из двух слов, разделенных запятой.
	// Если методу TermString передать строку, содержащую запятую, и вторым параметром ему передать 0,
	// то сравнение будет производиться с частью строки ДО запятой (сравнение с первым языком);
	// а если вторым параметром передать 1,
	// то сравнение будет производиться с частью строки ПОСЛЕ запятой (сравнение со вторым языком).
	if (TermString(IDS_SOME_BOOL_PROP, 0) == csPropName) // Сравнение с первым языком (английское наименование).
		*plPropNum = 0;

	if (TermString(IDS_SOME_BOOL_PROP, 1) == csPropName) // Сравнение со вторым языком (локальное наименование).
		*plPropNum = 0;

	TRC("1.txt", plPropNum);
	TRC("1.txt", *plPropNum);

	// Если свойство было найдено, то возвращается S_OK.
	return (*plPropNum == -1) ? S_FALSE : S_OK;
}

// Возвращает имя свойства с порядковым номером lPropNum.
// lPropAlias: 0 - английское наименование;
//             1 - локальное наименование.
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
	case someBooleanProperty: // То же самое, что 0, т.к. someBooleanProperty идет в перечислении первым.
		// Т.е. если передать 0, то будет возвращено строковое представление самого первого свойства, как оно задано в таблице строк.
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

// Возвращает значение свойства с порядковым номером lPropNum.
//
STDMETHODIMP CJuncter::GetPropVal(long lPropNum, VARIANT *pvarPropVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState())

		std::string ss ("CJuncter::GetPropVal");

	TRC("1.txt", ss);
	TRC("1.txt", lPropNum);
	TRC("1.txt", pvarPropVal);
	TRC("1.txt", V_I4(pvarPropVal));


	// Инициализация значения типа Variant.
	// Требуется, чтобы в случае отсутствия свойства с таким номером в переменной был размещен тип VT_EMPTY.
	// Поэтому возможно, что функция VariantInit как раз назначает такой тип.
	::VariantInit(pvarPropVal);

	TRC("1.txt", V_I4(pvarPropVal));

	switch(lPropNum)
	{
	case someBooleanProperty: // То же самое, что 0, т.к. someBooleanProperty идет в перечислении первым.

		// Назначение типа, который будет храниться в переменной pvarPropVal.
		V_VT(pvarPropVal) = VT_I4;

		// Если флаг установлен, то значение 1, иначе 0.
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

// Устанавливает значение свойства с порядковым номером lPropNum.
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
	case someBooleanProperty: // То же самое, что 0, т.к. someBooleanProperty идет в перечислении первым.

		// Проверка типа.
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

		// Если устанавливаемое значение не имеет нужный тип, то оно установлено быть не может.
		if (V_VT(pvarPropVal) != VT_I4)
		{
			ss = "CJuncter::SetPropVal - wrong type";
			TRC("1.txt", ss);

			return S_FALSE;
		}

		// Установка значения свойства.
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

// Возвращает флаг возможности чтения свойства с порядковым номером lPropNum.
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
	case someBooleanProperty: // То же самое, что 0, т.к. someBooleanProperty идет в перечислении первым.

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

// Возвращает флаг возможности записи свойства с порядкомым номером lPropNum.
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
	case someBooleanProperty: // То же самое, что 0, т.к. someBooleanProperty идет в перечислении первым.

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

// Возвращает количество методов расширения.
//
STDMETHODIMP CJuncter::GetNMethods(long *plMethods)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState())

		std::string ss ("CJuncter::GetNMethods");
	TRC("1.txt", ss);

	// LastMethod - это всегда последний элемент перечисления, определенного в классе CJuncter.

	*plMethods = LastMethod;

	return S_OK;
}

// Возвращает порядковый номер метода с именем bstrMethodName.
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

	// Преобразование в CString.
	CString csMethodName = OLE2T(bstrMethodName);

	// Сравнение имеющихся свойств (перечислены в строковой таблице) с переданным именем.
	// Строка с именем свойства может состоять из двух слов, разделенных запятой.
	// Если методу TermString передать строку, содержащую запятую, и вторым параметром ему передать 0,
	// то сравнение будет производиться с частью строки ДО запятой (сравнение с первым языком);
	// а если вторым параметром передать 1,
	// то сравнение будет производиться с частью строки ПОСЛЕ запятой (сравнение со вторым языком).

	// Метод.
	if (csMethodName == TermString(IDS_SOME_METHOD, 0)) // Сравнение с первым языком (английское наименование).
		*plMethodNum = 0;

	if (csMethodName == TermString(IDS_SOME_METHOD, 1)) // Сравнение со вторым языком (локальное наименование).
		*plMethodNum = 0;

	// Целочисленная функция.
	if (csMethodName == TermString(IDS_LONG_FUNC, 0))
		*plMethodNum = 1;

	if (csMethodName == TermString(IDS_LONG_FUNC, 1))
		*plMethodNum = 1;

	// Дробная.
	if (csMethodName == TermString(IDS_DOUBLE_FUNC, 0))
		*plMethodNum = 2;

	if (csMethodName == TermString(IDS_DOUBLE_FUNC, 1))
		*plMethodNum = 2;

	// Строковая.
	if (csMethodName == TermString(IDS_STRING_FUNC, 0))
		*plMethodNum = 3;

	if (csMethodName == TermString(IDS_STRING_FUNC, 1))
		*plMethodNum = 3;

	// Дата.
	if (csMethodName == TermString(IDS_DATE_FUNC, 0))
		*plMethodNum = 4;

	if (csMethodName == TermString(IDS_DATE_FUNC, 1))
		*plMethodNum = 4;

	// ДатаВремя.
	if (csMethodName == TermString(IDS_DATETIME_FUNC, 0))
		*plMethodNum = 5;

	if (csMethodName == TermString(IDS_DATETIME_FUNC, 1))
		*plMethodNum = 5;

	// Массив.
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
	
	// Если метод был найден, то возвращается S_OK.
	return (*plMethodNum == -1) ? S_FALSE : S_OK;
}

// Возвращает имя метода с порядковым номером lMethodNum.
// lMethodAlias: 0 - английское наименование;
//               1 - локальное наименование.
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
	case someMethod: // То же самое, что 0, т.к. someMethod идет в перечислении первым.

		// Т.е. если передать 0, то будет возвращено строковое представление самого первого свойства, как оно задано в таблице строк.
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

// Возвращает количество параметров метода с порядковым номером lMethodNum.
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

		*plParams = 0; // У метода someMethod нет параметров.

		break;

	case longFunc:

		*plParams = 1; // Функция someFunction использует 1 параметр.

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

		*plParams = 2; // Два параметра!

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

// Возвращает значение по умолчанию для для параметра с номером lParamNum метода с номером lMethodNum.
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

	// Инициализация возвращаемого значения значением VT_EMPTY.
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

		break; // Параметров у метода нет.

	case longFunc:

		ss = "CJuncter::GetParamDefValue, longFunc = " + std::to_string(lMethodNum);
		TRC("1.txt", ss);

		if (lParamNum == 0) {
			ss = "CJuncter::GetParamDefValue, longFunc, lParamNum = 0";
		} else {
			ss = "CJuncter::GetParamDefValue, longFunc, lParamNum = " + std::to_string(lParamNum);
		}

		// У метода 1 параметр, имеющий значение по умолчанию 0.
		// Назначение типа, который будет храниться в VARIANT.
		V_VT(pvarParamDefValue) = VT_I4;

		// Установка значения VARIANT.
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
		V_DATE(pvarParamDefValue) = 0; // Как представлять дату?

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

		// Первый параметр - строка, второй - дата.
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
			return S_FALSE; // Метод или параметр метода отсутствует.
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

		break; // Параметров у метода нет.

	default:

		ss = "CJuncter::GetParamDefValue, default someMethod = " + std::to_string(lMethodNum);
		TRC("1.txt", ss);

		return S_FALSE; // Метод или параметр метода отсутствует.
	}

	// Назначение типа, который будет храниться в переменной pvarPropVal.
	//V_VT(pvarParamDefValue) = VT_I4;
	//V_I4(pvarParamDefValue) = 0;

	ss = "CJuncter::GetParamDefValue, before return S_OK;";
	TRC("1.txt", ss);
	TRC("1.txt", pvarParamDefValue);
	TRC("1.txt", V_I4(pvarParamDefValue));

	return S_OK; // Операция завершена успешно.
}

// Возвращает флаг наличия возвращаемого значения у метода с порядкомым номером lMethodNum.
//
STDMETHODIMP CJuncter::HasRetVal(long lMethodNum, BOOL *pboolRetValue)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState())

		std::string ss ("CJuncter::HasRetVal");
	TRC("1.txt", ss);
	TRC("1.txt", lMethodNum);
	TRC("1.txt", pboolRetValue);
	TRC("1.txt", *pboolRetValue);

	// На всякий случай!!!
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


		*pboolRetValue = FALSE; // Метод значений не возвращает (т.е. является процедурой).

		break;

	case longFunc:

		if (lMethodNum == 0) {
			ss = "CJuncter::HasRetVal, longFunc = 0";
		} else {
			ss = "CJuncter::HasRetVal, longFunc = " + std::to_string(lMethodNum);
		}
		TRC("1.txt", ss);

		*pboolRetValue = TRUE; // Метод возвращает значение (т.е. является функцией).

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

// Выполняет метод с порядковым номером lMethodNum.
// paParams - двойной указатель на массив структур VARIANT, содержащий значения параметров метода.
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

		// TODO: Здесь код метода someMethod.
		ss = "CJuncter::CallAsProc - someMethod!!!";
		TRC("1.txt", ss);

		m_someBooleanPropertyValue = TRUE; // Метод устанавливает булево значение в TRUE.

		break;

	default:
		return S_FALSE;
	}

	ShowSafeArray("1.txt", paParams);

	return S_OK; // Метод вызван без ошибок.
	//return E_FAIL; // Метод вызван, произошла ошибка времени исполнения (выполнение модуля 1С прекращается).
	//return S_FALSE; // Отсутствует метод, соответствующий lMethodNum.
}

// Вызов пользовательской команды, определенной в 1С.
//
STDMETHODIMP CallUserFunc(IDispatch* iDisp, wchar_t* methodName)
{

	//wchar_t* methodName1 = L"myFunc1"; // Процедура в модуле формы обработки.
	DISPID dispid = NULL;
	HRESULT hr = iDisp->GetIDsOfNames(
		IID_NULL,              // Зарезервирован, всегда IID_NULL.
		&methodName,           // Имя (текст) метода или свойства.
		1,                     // Количество запрашиваемых имен.
		LOCALE_SYSTEM_DEFAULT, // Идентификатор языка.
		&dispid               // Параметр для хранения DISPID метода.
		);

	if (SUCCEEDED(hr)) {

		VARIANT vRetVal;
		VariantInit(&vRetVal);
		int ParmCount = 0;
		VARIANTARG vParams;
		VariantInit(&vParams);

		DISPPARAMS dParams = {&vParams, NULL, ParmCount, 0}; // заполним структуру параметров

		// вызовем его
		return iDisp->Invoke(dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT, DISPATCH_METHOD, &dParams, &vRetVal, NULL, 0);

	} else {
		return S_FALSE;
	}
}

// Выполняет метод с порядковым номером lMethodNum.
// pvarRetValue - указатель на структуру VARIANT, содержащую возвращаемое значение.
// paParams - двойной указатель на массив структур VARIANT, содержащий значения параметров метода.
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
			// TODO: Здесь код функции someFunction.
			ss = "CJuncter::CallAsFunc - longFunc!!!";
			TRC("1.txt", ss);

			// У функции есть возвращаемый параметр.
			// Назначение типа, который будет храниться в VARIANT.
			V_VT(pvarRetValue) = VT_I4;

			// Установка значения VARIANT.
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

			// Назначение типа.
			V_VT(pvarRetValue) = VT_R8;

			// Установка значения VARIANT.

			long iLBound; // Нижняя граница.
			HRESULT hr = SafeArrayGetLBound(*paParams, 1, &iLBound);
			if(FAILED(hr))
				return S_FALSE;

			VARIANT iVal; // Значение.

			VariantInit(&iVal); // Обязательная инициализация.

			// Получение значения, находящегося в первом элементе массива.
			hr = SafeArrayGetElement(*paParams, &iLBound, &iVal); // Тип назначается автоматически!
			if(FAILED(hr))
				return S_FALSE;

			// Присваивание значения.
			V_R8(pvarRetValue) = V_R8(&iVal);

			VariantClear(&iVal); // Обязательная очистка.

			break;
		}
	case stringFunc:
		{
			ss = "CJuncter::CallAsFunc - stringFunc!!!";
			TRC("1.txt", ss);

			// Получение значения из переданного массива.
			VARIANT var = GetNParam(*paParams, 0); // Тип назначается автоматически!

			// Назначение типа возвращаемого значения.
			V_VT(pvarRetValue) = VT_BSTR;

			// Установка значения возвращаемого значения.
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
			// *** РАБОТА С ИСХОДНЫМ МАССИВОМ ***
			// **********************************
			// Получение указателя на нулевой элемент массива SAFEARRAY
			// (содержит первый параметр функции 1С, который тоже является массивом).
			long inputIdx = 0;
			void* inputVoidPtr = NULL;
			HRESULT hr = SafeArrayPtrOfIndex(*paParams, &inputIdx, &inputVoidPtr);

			// Приведение полученного указателя void* к типу VARIANT*
			// (т.к. мы точно знаем, что там именно VARIANT).
			VARIANT* inputVarPtr = (VARIANT*)inputVoidPtr;

			///////////// ОПРЕДЕЛЕНИЕ ГРАНИЦ МАССИВА.
			long iLowerBound;
			hr = SafeArrayGetLBound(inputVarPtr->parray, 1, &iLowerBound);
			if(FAILED(hr))
				return S_FALSE;

			long iUpperBound;
			hr = SafeArrayGetUBound(inputVarPtr->parray, 1, &iUpperBound);
			if(FAILED(hr))
				return S_FALSE;

			//////////////////////////////////////////////

			void* sourcePtr  = NULL; // Указатель на элемент, заполняемый функцией SafeArrayPtrOfIndex.
			double* sourceValPtr = NULL;
			// Перебор входного массива от нижней границы к верхней.
			for(long l = iLowerBound; l <= iUpperBound; l++)
			{
				// Заполнение указателя на элемент.
				hr = SafeArrayPtrOfIndex(inputVarPtr->parray, &l, &sourcePtr);


				// Приведение указателей void* к требуемому типу.
				sourceValPtr = (double*)sourcePtr;

				// Изменение значения во входном массиве.
				++(*sourceValPtr);
			}


			// *******************************
			// *** РАБОТА С КОПИЕЙ МАССИВА ***
			// *******************************

			// Получение в переменную копии значений массива
			// (т.к. в функции содержится команда SafeArrayGetElement,
			// которая выполняет копирование элемента массива SAFEARRAY).
			VARIANT var = GetNParam(*paParams, 0);

			// Указатель, через который осуществляется работа с массивом
			// (исключительно для удобства).
			SAFEARRAY* sArr = V_ARRAY(&var); // то же самое, что и: SAFEARRAY* sArr = var.parray

			/////////// ОПРЕДЕЛЕНИЕ ГРАНИЦ МАССИВА.
			long iLBound; // Нижняя граница.
			hr = SafeArrayGetLBound(sArr, 1, &iLBound);
			if(FAILED(hr))
				return S_FALSE;

			long iUBound; // Верхняя граница.
			hr = SafeArrayGetUBound(sArr, 1, &iUBound);
			if(FAILED(hr))
				return S_FALSE;
			////////////////////////////////////////////


			///////////// СОЗДАНИЕ НОВОГО МАССИВА (С ТАКОЙ ЖЕ РАЗМЕРНОСТЬЮ).
			// Границы нового массива
			SAFEARRAYBOUND sBound[1];
			sBound[0].cElements = iUBound - iLBound + 1;
			sBound[0].lLbound = 0;

			// Получение типа, хранящегося в исходном массиве.
			VARTYPE varType;
			hr = SafeArrayGetVartype(sArr, &varType);
			if(FAILED(hr))
				return S_FALSE;

			// Создание нового массива.
			SAFEARRAY* sArrNew = SafeArrayCreate(varType, 1, sBound);
			//////////////////////////////////////////////


			/////////// ПЕРЕБОР МАССИВА.
			void* sourPtr  = NULL; // Указатель на элемент, заполняемый функцией SafeArrayPtrOfIndex.
			void* destPtr  = NULL; // Указатель на элемент, заполняемый функцией SafeArrayPtrOfIndex.

			// Перебор исходного массива от нижней границы к верхней.
			for(long l = iLBound; l <= iUBound; l++)
			{
				// Заполнение указателя на элемент.
				hr = SafeArrayPtrOfIndex(sArr, &l, &sourPtr);

				hr = SafeArrayPtrOfIndex(sArrNew, &l, &destPtr);

				// Приведение указателей void* к требуемому типу
				// и присвоение значений целевому массиву (исходный, умноженный на 2).
				*((double*)destPtr) = *((double*)sourPtr) * 2;
			}
			////////////////////////////////////////////

			///////////// ПРИСВОЕНИЕ ВОЗВРАЩАЕМОМУ ЗНАЧЕНИЮ ЗНАЧЕНИЯ МАССИВА.
			V_VT(pvarRetValue) = VT_ARRAY;
			V_ARRAY(pvarRetValue) = sArrNew;

			break;
		}

	case twoParamsFunc:

		{
			ss = "CJuncter::CallAsFunc - twoParamsFunc!!!";
			TRC("1.txt", ss);

			// Значения обоих параметров приводим к приемлемому строковому виду,
			// сцепляем и возвращаем в виде результирующего значения.

			// Первый параметр - строка.
			VARIANT var1 = GetNParam(*paParams, 0);

			///////////////////////////////////////////
			//// КОНВЕРТАЦИЯ CString в BSTR:
			//CString cs("Hello");
			//BSTR bstr = cs.AllocSysString();

			///////////////////////////////////////////
			//// КОНВЕРТАЦИЯ BSTR в CString:
			//BSTR bstr = ::SysAllocString(L"Hello");
			//_bstr_t tmp(bstr, FALSE);   //wrap the BSTR
			//CString cs(static_cast<const char*>(tmp));  //convert it
			//AfxMessageBox(cs, MB_OK, 0);

			//// ИЛИ:
			//BSTR bstr1;
			//bstr1 = ::SysAllocString(L"Hello1");
			//CString cs1(_com_util::ConvertBSTRToString(bstr1));
			//AfxMessageBox(cs1, MB_OK, 0);
			///////////////////////////////////////////

			// The _bstr_t class encapsulates and manages the BSTR data type.
			_bstr_t tmp1(V_BSTR(&var1), FALSE);
			CString cs1(static_cast<const char*>(tmp1));
			
			std::string stringParam1(cs1);

			// Второй параметр - дата.
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

			// Сцепляем.
			cs1 = cs1 + ", " + cs2 + ".";
			
			// Возвращаем.
			V_VT(pvarRetValue) = VT_BSTR;
			V_BSTR(pvarRetValue) = cs1.AllocSysString();

			// Освобождение памяти.
			::VariantClear(&var1);
			::VariantClear(&var2);
			
			::SysFreeString(pbstrOut);

			break;
		}

	case callBackFunc:

		{
			ss = "CJuncter::CallAsFunc - callBackFunc!!!";
			TRC("1.txt", ss);

			///////////// СОЗДАНИЕ НОВОГО МАССИВА ИЗ 3 элементов.
			// Границы нового массива
			SAFEARRAYBOUND sBound[1];
			sBound[0].cElements = 3;
			sBound[0].lLbound = 0;

			// Создание нового массива cо значениями строкового типа.
			SAFEARRAY* sArrNew = SafeArrayCreate(VT_BSTR, 1, sBound);
			//SAFEARRAY* sArrNew = SafeArrayCreate(VT_I4, 1, sBound);
			//////////////////////////////////////////////

			// Заполнение массива значениями.
			void* sourPtr  = NULL;

			LONG l = 0;
			HRESULT hr = SafeArrayPtrOfIndex(sArrNew, &l, &sourPtr);
			
			// Приведение указателей void* к требуемому типу
			// и присвоение значений целевому массиву.
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

			///////////// ПРИСВОЕНИЕ ВОЗВРАЩАЕМОМУ ЗНАЧЕНИЮ ЗНАЧЕНИЯ МАССИВА.
			V_VT(pvarRetValue) = VT_ARRAY;
			V_ARRAY(pvarRetValue) = sArrNew;


			//// Вызов внешнего события в 1С.

			// С диска ИТС:
			// Переданный в методе Init указатель на IDispatch позволяет получить доступ к 1С:Предприятию
			// через механизм OLE Automation. Из полученного указателя можно получить свойство AppDispatch,
			// доступное только для чтения.
			// Это свойство содержит указатель на IDispatch 1С:Предприятия (не путайте с переданным в Init)!
			// Свойство AppDispatch становится доступно только после полной инициализации всей системы 1С:Предприятия,
			// поэтому в момент загрузки внешней компоненты и вызова метода Init это свойство обеспечивает
			// доступ не ко всем возможностям 1С:Предприятия.

			// С диска ИТС:
			// Для доступа к 1С:Предприятию в версии 2.0 появилось специальное свойство AppDispatch. Фактически AppDispatch представляет собой объекта,
			// идентичный созданному через OLE Automation, поэтому, используя AppDispatch, можно вызывать процедуры
			// и функции глобального модуля 1С:Предприятия. Существует одна зависимость, связанная с этим механизмом:
			// после получения AppDispatch необходимо дополнительно вызывать AddRef у полученного IDispatch - в противном
			// случае после завершения работы возможно "зависание" 1С:Предприятия в памяти. Эта особонность д.б. учтена
			// разработчиком внешней компоненты при работе с 1С:Предприятием, начиная с релиза 7.70.006.

			// С форума:
			// https://partners.v8.1c.ru/forum/message/240854#m_240854
			// Процедуры и функции общих модулей доступны как обычный метод объекта AppDispatch,
			// если в их описании присутствует ключевое слово Экспорт и модуль, в котором они находятся, помечен для исполнения как Клиент.

			// Т.е. у нас есть два указателя IDispatch:
			// 1. Передаваемый системой в качестве параметра метода Init. Через него можно получать интерфейсы
			// 1С:Предприятия, но нельзя вызывать его методы;
			// 2. Находящийся в свойстве AppDispach. Он может использоваться для вызова методов 1С:Предприятия,
			// при этом используется механизм OLE Automation.

			// Получение того указателя IDispatch, который позволяет вызывать методы 1С:
			 
			// Переменная, заполняемая значением указателя с помощью вызовов GetPropertyByName.
			_variant_t var;
			
			// Указатель на IDispatch, позволяющий обращаться к 1С:Предприятию в целом.
			// Например, через него можно вызвать экспортную процедуру, находящуюся в общем модуле, исполняемом на клиенте.
			CComPtr<IDispatch> appConnect = NULL;

			// Указатель на IDispatch, позволяющий обращаться к обработкам в целом.
			// Практического значения, кроме получения указателя IDispatch конкретной обработки, не имеет.
			CComPtr<IDispatch> execConnect = NULL;
									
			// IDispatch для конкретной обработки с именем "COM".
			CComPtr<IDispatch> myExecConnect = NULL;
										
			// IDispatch для конкретной формы "Форма" конкретной обработки "COM".
			CComPtr<IDispatch> myExecFormConnect = NULL;
			
			hr = m_iConnect.GetPropertyByName(L"AppDispatch", var.GetAddress());
			if (SUCCEEDED(hr)) {

				// Проверка соответствия типа.
				ATLASSERT(V_VT(&var) == VT_DISPATCH);

				appConnect = V_DISPATCH(&var);

				hr = appConnect.GetPropertyByName(L"Обработки", var.GetAddress());
				if (SUCCEEDED(hr)) {

					execConnect = V_DISPATCH(&var);

					hr = execConnect.GetPropertyByName(L"COM", var.GetAddress());
					if (SUCCEEDED(hr)) {

						myExecConnect = V_DISPATCH(&var);

						// В 1С для открытия определенной формы используется команда
						// ПолучитьФорму(<Форма>, <Владелец>, <Ключ уникальности>).
						// Соответственно, для того, чтобы вызвать эту команду следует заполнить
						// её параметры.
						// При этом в массиве параметров порядок элементов д.б. обратным.

						// Массив параметров.
						_variant_t vars[3];

						// Имя формы.
						BSTR formName = SysAllocString(L"Форма1");
						V_VT(&vars[2]) = VT_BSTR;
						V_BSTR(&vars[2]) = formName;

						// Владелец формы.
						V_VT(&vars[1]) = VT_ERROR;

						// Создание ключа уникальности.
						UUID id;
						::UuidCreate(&id);
						
						BYTE* strId = NULL;
						::UuidToString(&id, &strId);

						// Ключ уникальности.
						vars[0] = T2OLE((LPCTSTR)strId);

						// Освобождение памяти.
						::SysFreeString(formName);
						::RpcStringFree(&strId);

						// Вызов команды "ПолучитьФорму" объекта обработки.
						hr = myExecConnect.InvokeN(L"ПолучитьФорму", vars, 3, var.GetAddress());
						if (SUCCEEDED(hr))
						{
							myExecFormConnect = V_DISPATCH(&var);

							// Вызов команды "Открыть" формы.
							hr = myExecFormConnect.Invoke0(L"Открыть");
						}
					}
				}

				// Попытки вызывать функции, используя разные указатели IDispatch.
				/////////////////////////////////////////////////
				// Указатель на 1С.
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
				// Если функция возвращает значение надо так:
				// VARIANT vMem;
				// vMem.vt = VT_DISPATCH;
				// Если не возвращает передаем в качестве параметра NULL.
				
				// VARIANTARG vParams;
				// Если функция не принимает аргументы то надо указать NULL.
				// Для вариантных значений пустая переменная и NULL не всегда тождественны.
				
				// В качестве флага надо передавать либо DISPATCH_METHOD, либо DISPATCH_PROPERTYGET L
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
				// Указатель на обработки.
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
				// Указатель на конкретную обработку.
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
				// Указатель на конкретную форму конкретной обработку.
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

	return S_OK; // Метод вызван без ошибок.
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

	CString csSource("Компонента"),csMessage("Включение"),csData;

	m_someBooleanPropertyValue = boolEnabled;

	csData = boolEnabled?"1":"0";

	m_iAsyncEvent->ExternalEvent(csSource.AllocSysString(),
		csMessage.AllocSysString(),
		csData.AllocSysString());

	return S_OK;
}


