#pragma comment(linker,"\"/manifestdependency:type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='*' publicKeyToken='6595b64144ccf1df' language='*'\"")

#pragma comment(lib,"wbemuuid")

#include <windows.h>
#include <comdef.h>
#include <wbemidl.h>

TCHAR szClassName[] = TEXT("Window");

VOID GetAntiVirusProduct(HWND hEdit)
{
	HRESULT hResult = CoInitializeEx(0, COINIT_MULTITHREADED);
	if (FAILED(hResult))
	{
		return;
	}
	hResult = CoInitializeSecurity(NULL, -1, NULL, NULL, RPC_C_AUTHN_LEVEL_DEFAULT,
		RPC_C_IMP_LEVEL_IMPERSONATE, NULL, EOAC_NONE, NULL);
	if (FAILED(hResult))
	{
		CoUninitialize();
		return;
	}
	IWbemLocator *pLoc = NULL;
	hResult = CoCreateInstance(CLSID_WbemLocator, 0, CLSCTX_INPROC_SERVER, IID_IWbemLocator, (LPVOID*)&pLoc);
	if (FAILED(hResult))
	{
		CoUninitialize();
		return;
	}
	IWbemServices *pSvc = NULL;
	hResult = pLoc->ConnectServer(BSTR(L"Root\\SecurityCenter2"), NULL, NULL, 0, NULL, 0, 0, &pSvc);
	if (FAILED(hResult))
	{
		hResult = pLoc->ConnectServer(BSTR(L"Root\\SecurityCenter"), NULL, NULL, 0, NULL, 0, 0, &pSvc);
		if (FAILED(hResult))
		{
			pLoc->Release();
			CoUninitialize();
			return;
		}
	}
	hResult = CoSetProxyBlanket(pSvc, RPC_C_AUTHN_WINNT, RPC_C_AUTHZ_NONE, NULL, RPC_C_AUTHN_LEVEL_CALL,
		RPC_C_IMP_LEVEL_IMPERSONATE, NULL, EOAC_NONE);
	if (FAILED(hResult))
	{
		pSvc->Release();
		pLoc->Release();
		CoUninitialize();
		return;
	}
	IEnumWbemClassObject *pEnumerator = NULL;
	hResult = pSvc->ExecQuery(BSTR(L"WQL"), BSTR(L"SELECT * FROM AntiVirusProduct"),
		WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY, NULL, &pEnumerator);
	if (FAILED(hResult))
	{
		pSvc->Release();
		pLoc->Release();
		CoUninitialize();
		return;
	}
	IWbemClassObject *pclsObj;
	SetWindowText(hEdit, 0);
	while (pEnumerator)
	{
		ULONG uReturn = 0;
		hResult = pEnumerator->Next(WBEM_INFINITE, 1, &pclsObj, &uReturn);
		if (FAILED(hResult) || uReturn == 0)
			break;
		VARIANT vtDisplayName;
		if (SUCCEEDED(pclsObj->Get(L"DisplayName", 0, &vtDisplayName, 0, 0)))
		{
			SendMessageW(hEdit, EM_REPLACESEL, 0, (LPARAM)(LPWSTR)vtDisplayName.bstrVal);
			SendMessageW(hEdit, EM_REPLACESEL, 0, (LPARAM)L"\r\n");
			VariantClear(&vtDisplayName);
		}
		pclsObj->Release();
	}
	SendMessage(hEdit, EM_SETSEL, 0, -1);
	pEnumerator->Release();
	pSvc->Release();
	pLoc->Release();
	CoUninitialize();
}

LRESULT CALLBACK WndProc(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
	static HWND hEdit;
	switch (msg)
	{
	case WM_CREATE:
		hEdit = CreateWindow(TEXT("EDIT"), 0, WS_VISIBLE | WS_CHILD | ES_MULTILINE | ES_AUTOVSCROLL | ES_AUTOHSCROLL | ES_READONLY, 0, 0, 0, 0, hWnd, 0, ((LPCREATESTRUCT)lParam)->hInstance, 0);
		GetAntiVirusProduct(hEdit);
		break;
	case WM_SETFOCUS:
		SetFocus(hEdit);
		break;
	case WM_SIZE:
		MoveWindow(hEdit, 0, 0, LOWORD(lParam), HIWORD(lParam), TRUE);
		break;
	case WM_DESTROY:
		PostQuitMessage(0);
		break;
	default:
		return DefWindowProc(hWnd, msg, wParam, lParam);
	}
	return 0;
}

int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPreInst, LPSTR pCmdLine, int nCmdShow)
{
	MSG msg;
	WNDCLASS wndclass = {
		0,
		WndProc,
		0,
		0,
		hInstance,
		0,
		LoadCursor(0,IDC_ARROW),
		0,
		0,
		szClassName
	};
	RegisterClass(&wndclass);
	HWND hWnd = CreateWindow(
		szClassName,
		TEXT("セキュリティソフトの名前を取得"),
		WS_OVERLAPPEDWINDOW | WS_CLIPCHILDREN,
		CW_USEDEFAULT,
		0,
		CW_USEDEFAULT,
		0,
		0,
		0,
		hInstance,
		0
	);
	ShowWindow(hWnd, SW_SHOWDEFAULT);
	UpdateWindow(hWnd);
	while (GetMessage(&msg, 0, 0, 0))
	{
		TranslateMessage(&msg);
		DispatchMessage(&msg);
	}
	return (int)msg.wParam;
}
