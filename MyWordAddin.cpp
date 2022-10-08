#include "MyWordAddin.h"
#include <memory>
#include <sstream>
#include "resource.h"

ULONG MyWordAddin::s_instanceCount = 0UL;

MyWordAddin::MyWordAddin()
{
	InterlockedIncrement(&s_instanceCount);  // 增加类实例个数
}

MyWordAddin::~MyWordAddin()
{
	InterlockedDecrement(&s_instanceCount);  // 减少类实例个数
}

STDMETHODIMP MyWordAddin::QueryInterface(REFIID riid, void **ppvObject)
{
	//
	// 根据不同的riid获取不同的接口指针并增加引用计数
	//

	if (riid == IID_IUnknown)
	{
		// 由于多个基类都继承自IUnknown，因此有多个IUnknown指针，
		// 所以这里先转换成IConnect基类指针，再转换为IUnknown指针
		IUnknown *pIUnknown = static_cast<IUnknown*>(static_cast<IConnect*>(this));
		*ppvObject = pIUnknown;
		pIUnknown->AddRef();
		return S_OK;
	}
	else if (riid == IID_IDispatch)
	{
		// 和IUnknown一样，需要先转换为IConnect指针，在转换为IDispatch指针
		IDispatch *pIDispatch = static_cast<IDispatch*>(static_cast<IConnect*>(this));
		*ppvObject = pIDispatch;
		pIDispatch->AddRef();
		return S_OK;
	}
	else if (riid == __uuidof(ADO::_IDTExtensibility2))
	{
		ADO::_IDTExtensibility2 *pIDTExt2 = static_cast<ADO::_IDTExtensibility2*>(this);
		*ppvObject = pIDTExt2;
		pIDTExt2->AddRef();
		return S_OK;
	}
	else if (riid == __uuidof(Office::IRibbonExtensibility))
	{
		Office::IRibbonExtensibility *pIRibbonExt = static_cast<Office::IRibbonExtensibility*>(this);
		*ppvObject = pIRibbonExt;
		pIRibbonExt->AddRef();
		return S_OK;
	}
	else if (riid == IID_IConnect)
	{
		IConnect *pIConnect = static_cast<IConnect*>(this);
		*ppvObject = pIConnect;
		pIConnect->AddRef();
		return S_OK;
	}
	else
	{
		*ppvObject = nullptr;
		return E_NOINTERFACE;
	}
}

STDMETHODIMP_(ULONG) MyWordAddin::AddRef()
{
	InterlockedIncrement(&m_ref);  // 增加引用计数
	return m_ref;
}

STDMETHODIMP_(ULONG) MyWordAddin::Release()
{
	InterlockedDecrement(&m_ref);  // 减少引用计数
	if (m_ref == 0UL)  // 如果引用计数为0
	{
		delete this;  // 销毁组件对象
		return 0;
	}
	return m_ref;
}

STDMETHODIMP MyWordAddin::Init()
{
	if (m_pITypeInfo == nullptr)
	{
		// 从注册表中读取类型库接口
		ITypeLib *pITypeLib = nullptr;
		HRESULT hr = LoadRegTypeLib(LIBID_MyWordAddinLib, 1, 0, 0x00, &pITypeLib);
		if (FAILED(hr))  // 如果失败，表示类型库没有被注册
		{
			// 从文件中读取类型库接口
			// 读取资源中的类型库时，需要指定的路径为：G:\projects\vs\MyWordAddin\Debug\MyWordAddin.dll\资源ID
			extern HMODULE g_hModule;
			std::unique_ptr<wchar_t[]> filename(new wchar_t[MAX_PATH + 1]());
			GetModuleFileName(g_hModule, filename.get(), MAX_PATH);
			*(filename.get() + MAX_PATH) = L'\0';
			std::wstringstream oss;
			oss << filename.get() << L"\\" << IDR_TYPELIB1 << std::ends;
			hr = LoadTypeLib(oss.str().c_str(), &pITypeLib);
			if (FAILED(hr))
			{
				return hr;
			}

			// 注册类型库到注册表
			RegisterTypeLib(pITypeLib, filename.get(), NULL);
		}

		// 获取类型信息接口
		hr = pITypeLib->GetTypeInfoOfGuid(IID_IConnect, &m_pITypeInfo);
		pITypeLib->Release();
		if (FAILED(hr))
		{
			return hr;
		}
	}

	return S_OK;
}

STDMETHODIMP MyWordAddin::GetTypeInfoCount(UINT *pctinfo)
{
	*pctinfo = 1;  // 类型信息的数量为1
	return S_OK;
}

STDMETHODIMP MyWordAddin::GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo)
{
	*ppTInfo = nullptr;

	// 因为只有一个类型信息，因此索引必须为0
	if (iTInfo != 0)
	{
		return DISP_E_BADINDEX;
	}

	*ppTInfo = m_pITypeInfo;

	(*ppTInfo)->AddRef();  // 增加引用计数

	return S_OK;
}

STDMETHODIMP MyWordAddin::GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId)
{
	if (riid != IID_NULL)
	{
		return DISP_E_UNKNOWNINTERFACE;
	}

	// 利用 ITypeInfo 实现 GetIDsOfNames
	return m_pITypeInfo->GetIDsOfNames(rgszNames, cNames, rgDispId);
}

STDMETHODIMP MyWordAddin::Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags,
	DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
{
	if (riid != IID_NULL)
	{
		return DISP_E_UNKNOWNINTERFACE;
	}

	SetErrorInfo(0, NULL);

	// 利用 ITypeInfo 实现 Invoke
	return m_pITypeInfo->Invoke(static_cast<IDispatch*>(static_cast<IConnect*>(this)),
		dispIdMember, wFlags, pDispParams, pVarResult, pExcepInfo, puArgErr);
}

STDMETHODIMP MyWordAddin::OnConnection(IDispatch *Application, ADO::ext_ConnectMode ConnectMode, IDispatch * AddInInst, SAFEARRAY **custom)
{
	// 在加载成功时弹出消息框
	MessageBox(NULL, TEXT("测试Word插件"), TEXT(""), MB_OK | MB_TOPMOST);
	return S_OK;
}

STDMETHODIMP MyWordAddin::OnDisconnection(ADO::ext_DisconnectMode RemoveMode, SAFEARRAY **custom)
{
	return S_OK;
}

STDMETHODIMP MyWordAddin::OnAddInsUpdate(SAFEARRAY  **custom)
{
	return S_OK;
}

STDMETHODIMP MyWordAddin::OnStartupComplete(SAFEARRAY **custom)
{
	return S_OK;
}

STDMETHODIMP MyWordAddin::OnBeginShutdown(SAFEARRAY **custom)
{
	return S_OK;
}

STDMETHODIMP MyWordAddin::GetCustomUI(BSTR RibbonID, BSTR *RibbonXml)
{
	extern HMODULE g_hModule;
	HRSRC hRsrc = FindResource(g_hModule, MAKEINTRESOURCE(IDR_XML1), L"XML");
	if (hRsrc == NULL)
	{
		return HRESULT_FROM_WIN32(GetLastError());
	}

	HGLOBAL hGlobal = LoadResource(g_hModule, hRsrc);
	if (hGlobal == NULL)
	{
		return HRESULT_FROM_WIN32(GetLastError());
	}

	DWORD size = SizeofResource(g_hModule, hRsrc);
	char *data = static_cast<char*>(LockResource(hGlobal));

	int wcSize = MultiByteToWideChar(CP_UTF8, 0, data, size, NULL, NULL);
	std::unique_ptr<wchar_t[]> wcData(new wchar_t[wcSize + 1]());
	MultiByteToWideChar(CP_UTF8, 0, data, size, wcData.get(), wcSize);
	*(wcData.get() + wcSize) = L'\0';

	*RibbonXml = SysAllocString(wcData.get());

	return S_OK;
}

STDMETHODIMP MyWordAddin::ButtonClicked(IDispatch *ribbonControl)
{
	MessageBox(NULL, TEXT("测试按钮被按下"), TEXT(""), MB_OK | MB_TOPMOST);

	return S_OK;
}
