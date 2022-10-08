#include "Factory.h"
#include "MyWordAddin.h"

ULONG Factory::s_lockCount = 0UL;

Factory::Factory()
{
}

Factory::~Factory()
{
}

STDMETHODIMP Factory::QueryInterface(REFIID riid, void **ppvObject)
{
	//
	// 根据不同的riid获取不同的接口指针并增加引用计数
	//

	if (riid == IID_IUnknown)
	{
		IUnknown *pIUnknown = static_cast<IUnknown*>(this);
		*ppvObject = pIUnknown;
		pIUnknown->AddRef();
		return S_OK;
	}
	else if (riid == IID_IClassFactory)
	{
		IClassFactory *pIClassFactory = static_cast<IClassFactory*>(this);
		*ppvObject = pIClassFactory;
		pIClassFactory->AddRef();
		return S_OK;
	}
	else
	{
		*ppvObject = nullptr;
		return E_NOINTERFACE;
	}
}

STDMETHODIMP_(ULONG) Factory::AddRef()
{
	InterlockedIncrement(&m_ref);  // 增加引用计数
	return m_ref;
}

STDMETHODIMP_(ULONG) Factory::Release()
{
	InterlockedDecrement(&m_ref);  // 减少引用计数
	if (m_ref == 0UL)  // 如果引用计数为0
	{
		delete this;  // 销毁工厂对象
		return 0;
	}
	return m_ref;
}

STDMETHODIMP Factory::CreateInstance(IUnknown *pUnkOuter, REFIID riid, void **ppvObject)
{
	// 不支持聚合
	if (pUnkOuter != nullptr)
	{
		return CLASS_E_NOAGGREGATION;
	}

	MyWordAddin *addin = new MyWordAddin();  // 创建组件对象

	HRESULT hr = addin->Init();  // 初始化组件
	if (FAILED(hr))
	{
		return hr;
	}

	// 获取所请求的接口，这里先AddRef()再Release()的目的
	// 是当QueryInterface失败时，自动销毁组件对象
	addin->AddRef();  
	hr = addin->QueryInterface(riid, ppvObject);
	addin->Release();

	return hr;
}

STDMETHODIMP Factory::LockServer(BOOL fLock)
{
	if (fLock)
	{
		InterlockedIncrement(&s_lockCount);  // 增加锁定计数
	}
	else
	{
		InterlockedDecrement(&s_lockCount);  // 减少锁定计数
	}
	return S_OK;
}
