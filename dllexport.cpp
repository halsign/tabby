#include <objbase.h>
#include <string>
#include "MyWordAddin.h"
#include "Factory.h"
#include "IFace.h"
#include <memory>

HMODULE g_hModule = NULL;  // DLL句柄

static const std::wstring ProgID = L"MyWordAddin.Component.1";
static const std::wstring VersionIndependentProgID = L"MyWordAddin.Component";
static const std::wstring FriendlyName = L"My Word Addin";

// 辅助函数，将键值写入注册表
static LONG SetKeyValue(HKEY hKey, const std::wstring &subKey, const std::wstring &valueName, const std::wstring &value);

// 辅助函数，获取GUID字符串
static std::wstring GetGUIDString(REFGUID refGUID);

BOOL WINAPI DllMain(HINSTANCE hinstDLL, DWORD fdwReason, LPVOID lpvReserved)
{
	// 在DLL加载时，获取句柄
	if (fdwReason == DLL_PROCESS_ATTACH)
	{
		g_hModule = hinstDLL;
	}

	return TRUE;
}

STDAPI DllCanUnloadNow()
{
	// 如果组件的个数为0，并且类厂的锁定计数为0，则可以卸载该DLL
	if (MyWordAddin::s_instanceCount == 0 && Factory::s_lockCount == 0)
	{
		return S_OK;
	}
	else
	{
		return S_FALSE;
	}
}

STDAPI DllGetClassObject(REFCLSID rclsid, REFIID riid, void **ppv)
{
	if (rclsid != CLSID_MyWordAddin)
	{
		return CLASS_E_CLASSNOTAVAILABLE;
	}

	Factory *factory = new Factory();  // 创建工厂对象

	// 获取所请求的接口，这里先AddRef()再Release()的目的
	// 是当QueryInterface失败时，自动销毁工厂对象
	factory->AddRef();
	HRESULT hr = factory->QueryInterface(riid, ppv);
	factory->Release();

	return hr;
}

STDAPI DllRegisterServer()
{
	//
	// 登记组件信息
	//

	// HKEY_CLASSES_ROOT
	//   |-CLSID
	//   |   |-{91A7F0E2-1235-45FE-AEB1-F3C5D0C20DB1} - My Word Addin
	//   |     |-InprocServer32 - G:\projects\vs\MyWordAddin\Debug\MyWordAddin.dll
	//   |     |-ProgID - MyWordAddin.Component.1
	//	 |     |-VersionIndependentProgID - MyWordAddin.Component
	//   |     |-TypeLib - {9CEA336F-7263-4844-BFA2-E6AA3CD316DC}
	//   |-MyWordAddin.Component - My Word Addin
	//   |   |-CLSID - {91A7F0E2-1235-45FE-AEB1-F3C5D0C20DB1}
	//   |   |-CurVer - MyWordAddin.Component.1
	//   |-MyWordAddin.Component.1 - My Word Addin
	//       |-CLSID - {91A7F0E2-1235-45FE-AEB1-F3C5D0C20DB1}

	// 将组件ID作为 HKEY_CLASSES_ROOT\CLSID 的子健写入注册表
	const std::wstring clsid = GetGUIDString(CLSID_MyWordAddin);
	std::wstring key = L"CLSID\\" + clsid;
	SetKeyValue(HKEY_CLASSES_ROOT, key, L"", FriendlyName);

	// 将 InprocServer32 作为组件ID的子健写入注册表
	std::unique_ptr<wchar_t[]> filename(new wchar_t[MAX_PATH + 1]());
	GetModuleFileName(g_hModule, filename.get(), MAX_PATH);
	*(filename.get() + MAX_PATH) = L'\0';
	SetKeyValue(HKEY_CLASSES_ROOT, key + L"\\InprocServer32", L"", filename.get());

	// 将 ProgID 作为组件ID的子健写入注册表
	SetKeyValue(HKEY_CLASSES_ROOT, key + L"\\ProgID", L"", ProgID);

	// 将 VersionIndependentProgID 作为组件ID的子健写入注册表
	SetKeyValue(HKEY_CLASSES_ROOT, key + L"\\VersionIndependentProgID", L"", VersionIndependentProgID);

	// 将ProgID作为 HKEY_CLASSES_ROOT 的子健写入注册表
	SetKeyValue(HKEY_CLASSES_ROOT, ProgID, L"", FriendlyName);
	SetKeyValue(HKEY_CLASSES_ROOT, ProgID + L"\\CLSID", L"", clsid);

	// 将版本无关的 ProgID 作为 HKEY_CLASSES_ROOT 的子健写入注册表
	SetKeyValue(HKEY_CLASSES_ROOT, VersionIndependentProgID, L"", FriendlyName);
	SetKeyValue(HKEY_CLASSES_ROOT, VersionIndependentProgID + L"\\CLSID", L"", clsid);
	SetKeyValue(HKEY_CLASSES_ROOT, VersionIndependentProgID + L"\\CurVer", L"", ProgID);

	//
	// 登记 Office 插件信息
	//

	// HKEY_CURRENT_USER\Software\Microsoft\Office\Word\Addins
	//   |-MyWordAddin.Component
	//								Description			My Word Addin
	//								FriendlyName		My Word Addin
	//								LoadBehavior		3

	key = L"Software\\Microsoft\\Office\\Word\\Addins\\" + VersionIndependentProgID;
	SetKeyValue(HKEY_CURRENT_USER, key, L"Description", FriendlyName);
	SetKeyValue(HKEY_CURRENT_USER, key, L"FriendlyName", FriendlyName);

	HKEY hSubKey = NULL;
	RegCreateKeyEx(HKEY_CURRENT_USER, key.c_str(), 0, NULL,
		REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, NULL, &hSubKey, NULL);
	DWORD dwValue = 3;
	RegSetValueEx(hSubKey, L"LoadBehavior", 0, REG_DWORD,
		reinterpret_cast<const BYTE*>(&dwValue), sizeof (dwValue));

	return S_OK;
}

STDAPI DllUnregisterServer()
{
	// 删除 HKEY_CLASSES_ROOT\CLSID 下的组件ID
	std::wstring subKey = L"CLSID\\" + GetGUIDString(CLSID_MyWordAddin);
	RegDeleteTree(HKEY_CLASSES_ROOT, subKey.c_str());

	// 删除 HKEY_CLASSES_ROOT 下的 ProgID
	RegDeleteTree(HKEY_CLASSES_ROOT, ProgID.c_str());

	// 删除 HKEY_CLASSES_ROOT 下的版本无关的　ProgID
	RegDeleteTree(HKEY_CLASSES_ROOT, VersionIndependentProgID.c_str());

	// 删除 Office 插件信息
	subKey = L"Software\\Microsoft\\Office\\Word\\Addins\\" + VersionIndependentProgID;
	RegDeleteTree(HKEY_CURRENT_USER, subKey.c_str());

	// 删除类型库信息
	subKey = L"TypeLib\\" + GetGUIDString(LIBID_MyWordAddinLib);
	RegDeleteTree(HKEY_CLASSES_ROOT, subKey.c_str());

	return S_OK;
}

LONG SetKeyValue(HKEY hKey, const std::wstring &subKey, const std::wstring &valueName, const std::wstring &value)
{
	HKEY hSubKey = NULL;
	LONG res = RegCreateKeyEx(hKey, subKey.c_str(), 0, NULL,
		REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, NULL, &hSubKey, NULL);
	if (res != ERROR_SUCCESS)
	{
		return res;
	}

	return RegSetValueEx(hSubKey, valueName.c_str(), 0, REG_SZ,
		reinterpret_cast<const BYTE*>(value.c_str()),
		(value.length() + 1) * sizeof(std::wstring::value_type));
}

std::wstring GetGUIDString(REFGUID refGUID)
{
	LPOLESTR lpGUID = nullptr;
	StringFromCLSID(refGUID, &lpGUID);
	std::wstring strGUID(lpGUID);
	CoTaskMemFree(lpGUID);
	return strGUID;
}
