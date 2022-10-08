#include <objbase.h>

// COM规定工厂类必须实现IClassFactory
class Factory : public IClassFactory
{
public:
	static ULONG s_lockCount;  // 加锁次数

	// 构造、析构
public:
	Factory();
	virtual ~Factory();

	// IUnknown 实现
public:
	STDMETHOD(QueryInterface)(REFIID riid, void **ppvObject) override;
	STDMETHOD_(ULONG, AddRef)() override;
	STDMETHOD_(ULONG, Release)() override;

	// IClassFactory 实现
public:
	STDMETHOD(CreateInstance)(IUnknown *pUnkOuter, REFIID riid, void **ppvObject) override;
	STDMETHOD(LockServer)(BOOL fLock) override;

private:
	ULONG m_ref = 0UL;  // 引用计数
};
