// ComTest2.cpp : 定义应用程序的入口点。
//

#include "stdafx.h"
#include "CppActiveXEventDemo.h"

#define MAX_LOADSTRING 100

#include <iostream>
#include <sstream>
#include <string>

#include <stdio.h>
#include <tchar.h>
#include <ocidl.h>
#include <ctype.h>

#include <assert.h>

HRESULT OLEMethod(int nType, VARIANT *pvResult, IDispatch *pDisp, LPCOLESTR ptName, int cArgs...)
{
	if (!pDisp) return E_FAIL;

	va_list marker;
	va_start(marker, cArgs);

	DISPPARAMS dp = { NULL, NULL, 0, 0 };
	DISPID dispidNamed = DISPID_PROPERTYPUT;
	DISPID dispID;
	char szName[200];


	// Convert down to ANSI
	WideCharToMultiByte(CP_ACP, 0, ptName, -1, szName, 256, NULL, NULL);

	// Get DISPID for name passed...
	HRESULT hr = pDisp->GetIDsOfNames(IID_NULL, (LPOLESTR*)&ptName, 1, LOCALE_USER_DEFAULT, &dispID);
	if (FAILED(hr)) {
		return hr;
	}
	// Allocate memory for arguments...
	VARIANT *pArgs = new VARIANT[cArgs + 1];
	// Extract arguments...
	for (int i = 0; i < cArgs; i++) {
		pArgs[i] = va_arg(marker, VARIANT);
	}

	// Build DISPPARAMS
	dp.cArgs = cArgs;
	dp.rgvarg = pArgs;

	// Handle special-case for property-puts!
	if (nType & DISPATCH_PROPERTYPUT) {
		dp.cNamedArgs = 1;
		dp.rgdispidNamedArgs = &dispidNamed;
	}

	// Make the call!
	hr = pDisp->Invoke(dispID, IID_NULL, LOCALE_SYSTEM_DEFAULT, nType, &dp, pvResult, NULL, NULL);
	if (FAILED(hr)) {
		return hr;
	}
	// End variable-argument section...
	va_end(marker);

	delete[] pArgs;
	return hr;
}

ITypeInfo *g_eventinfo;

std::wstring qaxTypeInfoNames(ITypeInfo *typeInfo, FUNCDESC* funcdesc) {
	HRESULT hr;

	BSTR bstrNames[256];
	UINT maxNames = 255;
	UINT maxNamesOut = 0;
	hr = typeInfo->GetNames(funcdesc->memid, reinterpret_cast<BSTR *>(&bstrNames), maxNames, &maxNamesOut);
	assert(hr == S_OK);

	std::wstring funcname(bstrNames[0]);

	std::wstringstream ss;
	ss << _T("# ") << funcdesc->memid << _T(" ") << funcname << _T("(");
	for (UINT p = 1; p < maxNamesOut; ++p) {
		ss << bstrNames[p];
		SysFreeString(bstrNames[p]);

		if (p < maxNamesOut - 1) {
			ss << _T(",");
		}
	}
	ss << _T(")");

	return ss.str();
}


class Handler : public IDispatch {
public:
	unsigned long __stdcall AddRef() override
	{
		return InterlockedIncrement(&ref);
	}
	unsigned long __stdcall Release() override
	{
		LONG refCount = InterlockedDecrement(&ref);
		if (!refCount)
			delete this;

		return refCount;
	}

	HRESULT __stdcall QueryInterface(REFIID riid, void **ppvObject) override
	{
		*ppvObject = 0;
		if (riid == IID_IUnknown)
			*ppvObject = static_cast<IUnknown *>(static_cast<IDispatch *>(this));
		else if (riid == IID_IDispatch)
			*ppvObject = static_cast<IDispatch *>(this);
		else
			return E_NOINTERFACE;

		AddRef();
		return S_OK;
	}

	// IDispatch
	HRESULT __stdcall GetTypeInfoCount(unsigned int *) override
	{
		return E_NOTIMPL;
	}
	HRESULT __stdcall GetTypeInfo(UINT, LCID, ITypeInfo **) override
	{
		return E_NOTIMPL;
	}
	HRESULT __stdcall GetIDsOfNames(const _GUID &, wchar_t **, unsigned int,
		unsigned long, long *) override
	{
		return E_NOTIMPL;
	}

	HRESULT __stdcall Invoke(DISPID dispIdMember, REFIID riid, LCID lcid,
		WORD wFlags, DISPPARAMS *pDispParams,
		VARIANT *var, EXCEPINFO *exception, UINT * i) override
	{
		HRESULT hr;
		TYPEATTR *eventattr;
		ITypeInfo *eventinfo = g_eventinfo;
		hr = g_eventinfo->GetTypeAttr(&eventattr);
		assert(hr == S_OK);

		for (UINT fd = 0; fd < eventattr->cFuncs; ++fd) {
			FUNCDESC *funcdesc;
			hr = eventinfo->GetFuncDesc(fd, &funcdesc);
			assert(hr == S_OK);
			if (!funcdesc)
				break;
			if (funcdesc->memid != dispIdMember) {
				continue;
			}
			if (funcdesc->invkind != INVOKE_FUNC ||
				funcdesc->funckind != FUNC_DISPATCH) {
				eventinfo->ReleaseFuncDesc(funcdesc);
				continue;
			}

			std::wcout << "Event triggered: " << qaxTypeInfoNames(eventinfo, funcdesc) << std::endl;

			eventinfo->ReleaseFuncDesc(funcdesc);
		}
		return S_OK;
	}

private:
	long ref = 0;
};

Handler g_handler;
DWORD g_cookie;

void readEventInterface(ITypeInfo *eventinfo, IConnectionPoint *cpoint) {
	HRESULT hr;

	g_eventinfo = eventinfo;
	TYPEATTR *eventattr;
	hr = eventinfo->GetTypeAttr(&eventattr);
	assert(hr == S_OK);
	// if (!eventattr) return;
	if (eventattr->typekind != TKIND_DISPATCH) {
		eventinfo->ReleaseTypeAttr(eventattr);
		return;
	}

	IID conniid;
	hr = cpoint->GetConnectionInterface(&conniid);
	assert(hr == S_OK);

	// get information about all event functions
	for (UINT fd = 0; fd < eventattr->cFuncs; ++fd) {
		FUNCDESC *funcdesc;
		eventinfo->GetFuncDesc(fd, &funcdesc);
		if (!funcdesc)
			break;
		if (funcdesc->invkind != INVOKE_FUNC ||
			funcdesc->funckind != FUNC_DISPATCH) {
			eventinfo->ReleaseFuncDesc(funcdesc);
			continue;
		}

		qaxTypeInfoNames(eventinfo, funcdesc);
		// list all availabe events
		std::wcout << "Event available: " << qaxTypeInfoNames(eventinfo, funcdesc) << std::endl;

		eventinfo->ReleaseFuncDesc(funcdesc);
	}
	eventinfo->ReleaseTypeAttr(eventattr);

	hr = cpoint->Advise(&g_handler, &g_cookie);
	assert(hr == S_OK);

	std::wcout << std::endl;

	std::wcout << "@@@@@@@@@@Waiting events, please create a new Excel workbook@@@@@@@@@@" << std::endl;
}

void testExcel() {
	CoInitialize(nullptr);
	CLSID clsid;
	HRESULT hr;
	hr = CLSIDFromProgID(_T("Excel.Application"), &clsid);
	assert(hr == S_OK);

	IDispatch *application;
	hr = CoCreateInstance(clsid, 0, CLSCTX_SERVER, IID_IDispatch, reinterpret_cast<LPVOID*>(&application));
	assert(hr == S_OK);
	assert(application);

	VARIANT var;
	var.vt = VT_I4;
	var.lVal = 1; // 1=visible; 0=invisible;
	hr = OLEMethod(DISPATCH_PROPERTYPUT, NULL, application, L"Visible", 1, var);
	assert(hr == S_OK);

	UINT count;
	hr = application->GetTypeInfoCount(&count);
	assert(hr == S_OK);
	// std::wcout << _T("count: ") << count << std::endl;

	UINT index = 0;
	ITypeInfo *dispInfo = nullptr;
	hr = application->GetTypeInfo(index, LOCALE_USER_DEFAULT, &dispInfo);
	assert(hr == S_OK);

	ITypeLib *typelib = nullptr;
	hr = dispInfo->GetContainingTypeLib(&typelib, &index);
	assert(hr == S_OK);


	IConnectionPointContainer *cpoints = nullptr;
	hr = application->QueryInterface(IID_IConnectionPointContainer, reinterpret_cast<void **>(&cpoints));
	assert(hr == S_OK);
	assert(cpoints);

	IEnumConnectionPoints *epoints = nullptr;
	hr = cpoints->EnumConnectionPoints(&epoints);
	assert(hr == S_OK);
	assert(epoints);

	epoints->Reset();

	ULONG c = 1;

	while (true) {
		IConnectionPoint *cpoint = nullptr;
		hr = epoints->Next(c, &cpoint, &c);
		if (hr != S_OK) {
			break;
		}
		assert(cpoint);

		IID conniid;
		hr = cpoint->GetConnectionInterface(&conniid);
		assert(hr == S_OK);
		OLECHAR* str;
		StringFromCLSID(conniid, &str);
		// std::wcout << str << std::endl;
		StringFromCLSID(IID_IPropertyNotifySink, &str);
		// std::wcout << _T("???") << str << std::endl;
		// IID_IPropertyNotifySink;
		// assert(conniid == IID_IPropertyNotifySink);
		if (conniid == IID_IPropertyNotifySink) {
			break;
		}

		ITypeInfo *eventinfo = nullptr;
		hr = typelib->GetTypeInfoOfGuid(conniid, &eventinfo);
		assert(hr == S_OK);

		readEventInterface(eventinfo, cpoint);
	}

	// CoUninitialize();
}


// 全局变量:
HINSTANCE hInst;                                // 当前实例
WCHAR szTitle[MAX_LOADSTRING];                  // 标题栏文本
WCHAR szWindowClass[MAX_LOADSTRING];            // 主窗口类名

// 此代码模块中包含的函数的前向声明:
ATOM                MyRegisterClass(HINSTANCE hInstance);
BOOL                InitInstance(HINSTANCE, int);
LRESULT CALLBACK    WndProc(HWND, UINT, WPARAM, LPARAM);
INT_PTR CALLBACK    About(HWND, UINT, WPARAM, LPARAM);

int APIENTRY wWinMain(_In_ HINSTANCE hInstance,
                     _In_opt_ HINSTANCE hPrevInstance,
                     _In_ LPWSTR    lpCmdLine,
                     _In_ int       nCmdShow)
{
	// alloc console and redirect
	AllocConsole();
	FILE* file;
	freopen_s(&file, "CONOUT$", "w", stdout);
	freopen_s(&file, "CONOUT$", "w", stderr);

	testExcel();

    UNREFERENCED_PARAMETER(hPrevInstance);
    UNREFERENCED_PARAMETER(lpCmdLine);

    // TODO: 在此处放置代码。

    // 初始化全局字符串
    LoadStringW(hInstance, IDS_APP_TITLE, szTitle, MAX_LOADSTRING);
    LoadStringW(hInstance, IDC_COMTEST2, szWindowClass, MAX_LOADSTRING);
    MyRegisterClass(hInstance);

    // 执行应用程序初始化:
    if (!InitInstance (hInstance, nCmdShow))
    {
        return FALSE;
    }

    HACCEL hAccelTable = LoadAccelerators(hInstance, MAKEINTRESOURCE(IDC_COMTEST2));

    MSG msg;

    // 主消息循环:
    while (GetMessage(&msg, nullptr, 0, 0))
    {
        if (!TranslateAccelerator(msg.hwnd, hAccelTable, &msg))
        {
            TranslateMessage(&msg);
            DispatchMessage(&msg);
        }
    }

    return (int) msg.wParam;
}



//
//  函数: MyRegisterClass()
//
//  目标: 注册窗口类。
//
ATOM MyRegisterClass(HINSTANCE hInstance)
{
    WNDCLASSEXW wcex;

    wcex.cbSize = sizeof(WNDCLASSEX);

    wcex.style          = CS_HREDRAW | CS_VREDRAW;
    wcex.lpfnWndProc    = WndProc;
    wcex.cbClsExtra     = 0;
    wcex.cbWndExtra     = 0;
    wcex.hInstance      = hInstance;
    wcex.hIcon          = LoadIcon(hInstance, MAKEINTRESOURCE(IDI_COMTEST2));
    wcex.hCursor        = LoadCursor(nullptr, IDC_ARROW);
    wcex.hbrBackground  = (HBRUSH)(COLOR_WINDOW+1);
    wcex.lpszMenuName   = MAKEINTRESOURCEW(IDC_COMTEST2);
    wcex.lpszClassName  = szWindowClass;
    wcex.hIconSm        = LoadIcon(wcex.hInstance, MAKEINTRESOURCE(IDI_SMALL));

    return RegisterClassExW(&wcex);
}

//
//   函数: InitInstance(HINSTANCE, int)
//
//   目标: 保存实例句柄并创建主窗口
//
//   注释:
//
//        在此函数中，我们在全局变量中保存实例句柄并
//        创建和显示主程序窗口。
//
BOOL InitInstance(HINSTANCE hInstance, int nCmdShow)
{
   hInst = hInstance; // 将实例句柄存储在全局变量中

   HWND hWnd = CreateWindowW(szWindowClass, szTitle, WS_OVERLAPPEDWINDOW,
      CW_USEDEFAULT, 0, CW_USEDEFAULT, 0, nullptr, nullptr, hInstance, nullptr);

   if (!hWnd)
   {
      return FALSE;
   }

   ShowWindow(hWnd, nCmdShow);
   UpdateWindow(hWnd);

   return TRUE;
}

//
//  函数: WndProc(HWND, UINT, WPARAM, LPARAM)
//
//  目标: 处理主窗口的消息。
//
//  WM_COMMAND  - 处理应用程序菜单
//  WM_PAINT    - 绘制主窗口
//  WM_DESTROY  - 发送退出消息并返回
//
//
LRESULT CALLBACK WndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam)
{
    switch (message)
    {
    case WM_COMMAND:
        {
            int wmId = LOWORD(wParam);
            // 分析菜单选择:
            switch (wmId)
            {
            case IDM_ABOUT:
                DialogBox(hInst, MAKEINTRESOURCE(IDD_ABOUTBOX), hWnd, About);
                break;
            case IDM_EXIT:
                DestroyWindow(hWnd);
                break;
            default:
                return DefWindowProc(hWnd, message, wParam, lParam);
            }
        }
        break;
    case WM_PAINT:
        {
            PAINTSTRUCT ps;
            HDC hdc = BeginPaint(hWnd, &ps);
            // TODO: 在此处添加使用 hdc 的任何绘图代码...
            EndPaint(hWnd, &ps);
        }
        break;
    case WM_DESTROY:
        PostQuitMessage(0);
        break;
    default:
        return DefWindowProc(hWnd, message, wParam, lParam);
    }
    return 0;
}

// “关于”框的消息处理程序。
INT_PTR CALLBACK About(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam)
{
    UNREFERENCED_PARAMETER(lParam);
    switch (message)
    {
    case WM_INITDIALOG:
        return (INT_PTR)TRUE;

    case WM_COMMAND:
        if (LOWORD(wParam) == IDOK || LOWORD(wParam) == IDCANCEL)
        {
            EndDialog(hDlg, LOWORD(wParam));
            return (INT_PTR)TRUE;
        }
        break;
    }
    return (INT_PTR)FALSE;
}
