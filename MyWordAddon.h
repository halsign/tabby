#include <objbase.h>
#include "IFace.h"

#import "libid:AC0714F2-3D04-11D1-AE7D-00A0C90F26F4" auto_rename auto_search raw_interfaces_only rename_namespace("ADO")
#import "libid:2DF8D04C-5BFA-101B-BDE5-00AA0044DE52" version("2.5") auto_rename auto_search raw_interfaces_only rename_namespace("Office")

// Office 插件必须实现 ADO::_IDTExtensibility2
// 实现 IRibbonExtensibility 以加载自定义Ribbon界面
class MyWordAddin : public ADO::_IDTExtensibility2, public Office::IRibbonExtensibility, public IConnect
{
public:
	static ULONG s_instanceCount;  // 组件的个数

	// 构造、析构
public:
	MyWordAddin();
	virtual ~MyWordAddin();

	// 初始化函数，获取ITypeInfo接口，在需要时注册类型库到注册表，见实现
	STDMETHOD(Init)();

	// IUnknown 实现
public:
	STDMETHOD(QueryInterface)(REFIID riid, void **ppv) override;
	STDMETHOD_(ULONG, AddRef)() override;
	STDMETHOD_(ULONG, Release)() override;

	// IDispatch 实现
public:
	STDMETHOD(GetTypeInfoCount)(UINT *pctinfo) override;
	STDMETHOD(GetTypeInfo)(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo) override;
	STDMETHOD(GetIDsOfNames)(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId) override;
	STDMETHOD(Invoke)(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags,
		DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr) override;

	// _IDTExtensibility2 实现
public:
	STDMETHOD(OnConnection)(IDispatch *Application, ADO::ext_ConnectMode ConnectMode, IDispatch * AddInInst, SAFEARRAY **custom) override;
	STDMETHOD(OnDisconnection)(ADO::ext_DisconnectMode RemoveMode, SAFEARRAY **custom) override;
	STDMETHOD(OnAddInsUpdate)(SAFEARRAY  **custom) override;
	STDMETHOD(OnStartupComplete)(SAFEARRAY **custom) override;
	STDMETHOD(OnBeginShutdown)(SAFEARRAY **custom) override;

	// IRibbonExtensibility 实现
public:
	STDMETHOD(GetCustomUI)(BSTR RibbonID, BSTR *RibbonXml) override;

	// IConnect 实现
public:
	STDMETHOD(ButtonClicked)(IDispatch *ribbonControl) override;

private:
	ULONG m_ref = 0UL;  // 引用计数
	ITypeInfo *m_pITypeInfo = nullptr;  // 类型信息，用来实现 IDispatch
};
