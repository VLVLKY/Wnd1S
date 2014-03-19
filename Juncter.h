// Juncter.h: объявление CJuncter

#pragma once
#include "resource.h"       // основные символы



#include "Wnd1S_i.h"

#include "atlctl.h"

#if defined(_WIN32_WCE) && !defined(_CE_DCOM) && !defined(_CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA)
#error "Однопотоковые COM-объекты не поддерживаются должным образом платформой Windows CE, например платформами Windows Mobile, в которых не предусмотрена полная поддержка DCOM. Определите _CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA, чтобы принудить ATL поддерживать создание однопотоковых COM-объектов и разрешить использование его реализаций однопотоковых COM-объектов. Для потоковой модели в вашем rgs-файле задано значение 'Free', поскольку это единственная потоковая модель, поддерживаемая не-DCOM платформами Windows CE."
#endif

using namespace ATL;

// Возможные коды сообщений.
enum AddInErrors {

	// Выводится строка в окне сообщений:
	ADDIN_E_NONE = 1000,
	ADDIN_E_ORDINARY = 1001,
	ADDIN_E_ATTENTION = 1002,
	ADDIN_E_IMPORTANT = 1003,
	ADDIN_E_VERY_IMPORTANT = 1004,
	ADDIN_E_INFO = 1005,
	ADDIN_E_FAIL = 1006,

	// Выводится окно предупреждения с кнопкой ОК:
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

	// Значение булевого свойства.
	BOOL m_someBooleanPropertyValue;

	// Свойства.
	enum
	{
		someBooleanProperty = 0, // Некоторое свойство.
		//propIsTimerPresent,
		LastProp                 // Этот элемент всегда идет последним и использутся для получения количества используемых свойств.
	};

	// Методы
	enum 
	{
		someMethod = 0, // Некоторый метод.
		longFunc,       // Некоторая функция.
		doubleFunc,     // Дробная.
		stringFunc,     // Строковая.
		dateFunc,       // Дата.
		dateTimeFunc,   // ДатаВремя.
		arrayFunc,      // Массив.
		twoParamsFunc,  // Два параметра.
		callBackFunc,   // Функция, вызывающая метод 1С.
		LastMethod      // Этот элемент всегда идет последним и использутся для получения количества используемых методов.
	};

public:

	// Указатель на интерфейс 1С.
	CComPtr<IDispatch> m_iConnect;

	// Указатель для работы с механизмом сохранения параметров.
	// IPropertyProfile унаследован от стандартного IPropertyBag.
	// Реализован интерфейс в классе JuncterConnection.
	//IPropertyProfile *m_iProfile;

	// Интерфейс, используемый для создания внешнего события в 1С.
	// Реализован интерфейс в классе JuncterConnection.
	IAsyncEvent *m_iAsyncEvent;

	// Интерфейс, используемый для создания внешнего события в 1С.
	// Реализован интерфейс в классе JuncterConnection.
	//IStatusLine *m_iStatusLine;

	// Стандартный COM-интерфейс.
	// Используется для сообщения пользователю информации о своей работе.
	// Возникающие сообщения обрабатываются как в течение работы программы (при асинхронном помещениее их в очередь),
	// так и при:
	// - возврате из метода инициализации Init;
	// - возврате из метода расширения.
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

	// МЕТОДЫ КЛАССА
	CString TermString(UINT uiResID,long nAlias);

	// МЕТОДЫ ИНТЕРФЕЙСОВ

	// IInitDone
	STDMETHOD(Init)(IDispatch* pConnection);
	STDMETHOD(Done)(void);
	STDMETHOD(GetInfo)(SAFEARRAY ** pInfo);

	// ILanguageExtender Methods
	STDMETHOD(RegisterExtensionAs)(BSTR * bstrExtensionName);

	// Работа со свойствами.
	STDMETHOD(GetNProps)(long *plProps);
	STDMETHOD(FindProp)(BSTR bstrPropName,long *plPropNum);
	STDMETHOD(GetPropName)(long lPropNum,long lPropAlias,BSTR *pbstrPropName);
	STDMETHOD(GetPropVal)(long lPropNum,VARIANT *pvarPropVal);
	STDMETHOD(SetPropVal)(long lPropNum,VARIANT *pvarPropVal);
	STDMETHOD(IsPropReadable)(long lPropNum,BOOL *pboolPropRead);
	STDMETHOD(IsPropWritable)(long lPropNum,BOOL *pboolPropWrite);

	// Работа с методами.
	STDMETHOD(GetNMethods)(long *plMethods);
	STDMETHOD(FindMethod)(BSTR bstrMethodName,long *plMethodNum);
	STDMETHOD(GetMethodName)(long lMethodNum,long lMethodAlias,BSTR *pbstrMethodName);
	STDMETHOD(HasRetVal)(long lMethodNum,BOOL *pboolRetValue);

	// Работа с параметрами.
	STDMETHOD(GetNParams)(long lMethodNum, long *plParams);
	STDMETHOD(GetParamDefValue)(long lMethodNum,long lParamNum,VARIANT *pvarParamDefValue);

	STDMETHOD(CallAsProc)(long lMethodNum,SAFEARRAY **paParams);
	STDMETHOD(CallAsFunc)(long lMethodNum,VARIANT *pvarRetValue,SAFEARRAY **paParams);

	// IPropertyLink Methods
	STDMETHOD(get_Enabled)(BOOL * boolEnabled);
	STDMETHOD(put_Enabled)(BOOL boolEnabled);
};

OBJECT_ENTRY_AUTO(__uuidof(Juncter), CJuncter)
