// dllmain.h: объ€вление класса модул€.

class CWnd1SModule : public ATL::CAtlDllModuleT< CWnd1SModule >
{
public :
	DECLARE_LIBID(LIBID_Wnd1SLib)
	//DECLARE_REGISTRY_APPID_RESOURCEID(IDR_WND1S, "{0D51EABD-0943-4491-BAE3-97753A31797D}")
};

extern class CWnd1SModule _AtlModule;
