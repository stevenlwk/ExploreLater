#include "stdafx.h"
#include "commctrl.h"
#include "comutil.h"
#include "windows.h"
#include "Windowsx.h"
#include "shlobj.h"
#include "shobjidl.h"
#include "shlguid.h"
#include "shlwapi.h"
#include "strsafe.h"
#include "ExploreLater.h"

#define MAX_LOADSTRING 100

HINSTANCE hInst;
char szStartupPath[MAX_PATH];

int APIENTRY wWinMain(_In_ HINSTANCE hInstance, _In_opt_ HINSTANCE hPrevInstance, _In_ LPWSTR lpCmdLine, _In_ int nCmdShow) {
	HWND hDialog;
	
	UNREFERENCED_PARAMETER(hPrevInstance);
	UNREFERENCED_PARAMETER(lpCmdLine);

	CoInitialize(nullptr);
	// Construct the path to the startup directory
	if(SUCCEEDED(SHGetFolderPathA(NULL, CSIDL_APPDATA, NULL, 0, szStartupPath))) {
		StringCchCatA(szStartupPath, MAX_PATH, "\\Microsoft\\Windows\\Start Menu\\Programs\\Startup\\");
	}

	hDialog = CreateDialog(hInst, MAKEINTRESOURCE(IDD_EXPLORELATER), 0, DialogProc);
	if(!hDialog) {
		return FALSE;
	}

	HACCEL hAccelTable = LoadAccelerators(hInstance, MAKEINTRESOURCE(IDC_EXPLORELATER));
	MSG msg;

	while(GetMessage(&msg, nullptr, 0, 0)) {
		if(!IsDialogMessage(hDialog, &msg)) {
			TranslateMessage(&msg);
			DispatchMessage(&msg);
		}
	}
	return (int)msg.wParam;
}

BOOL CALLBACK DialogProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam) {
	switch(message) {
		case WM_INITDIALOG:
		{
			HWND hCheck = GetDlgItem(hWnd, IDC_CHECK);

			InitList(hWnd);
			PostMessage(hCheck, BM_SETCHECK, BST_CHECKED, 0);
			LoadStartupDirectories(hWnd, true);
		}
		return TRUE;
		case WM_COMMAND:
			switch(HIWORD(wParam)) {
				case BN_CLICKED:
					switch(LOWORD(wParam)) {
						case IDADD:
							AddAllOpenExplorers(hWnd);
							LoadStartupDirectories(hWnd);
							break;
						case IDREMOVE:
							RemoveAllSelectedLinks(hWnd);
							LoadStartupDirectories(hWnd);
							break;
						case IDQUIT:
							SendMessage(hWnd, WM_DESTROY, NULL, NULL);
							break;
					}
					break;
			}
			return TRUE;
		/*case WM_NOTIFY:
		{
			NMHDR* header = (NMHDR*)lParam;
			NMLISTVIEW* nmlist = (NMLISTVIEW*)lParam;

			if(header->code == LVN_ITEMCHANGED && nmlist->uNewState & LVIS_STATEIMAGEMASK) {
			}
		}
		return TRUE;*/
		case WM_DESTROY:
			PostQuitMessage(0);
			return TRUE;
		case WM_CLOSE:
			DestroyWindow(hWnd);
			return TRUE;
	}
	return FALSE;
}

void InitList(HWND hWnd) {
	HWND hListView;
	LVCOLUMN lvc;
	RECT rect;
	int cx, numCols;
	WCHAR label[2][7] = {L"Name", L"Target"};

	hListView = GetDlgItem(hWnd, IDC_LIST);
	numCols = sizeof(label) / sizeof(label[0]);
	GetClientRect(hListView, &rect);
	cx = (rect.right - rect.left) / numCols;
	ListView_SetExtendedListViewStyle(hListView, LVS_EX_AUTOCHECKSELECT | LVS_EX_CHECKBOXES);

	for(int i = 0; i < numCols; i++) {
		lvc.mask = LVCF_WIDTH | LVCF_TEXT;
		lvc.iSubItem = i;
		lvc.pszText = label[i];
		lvc.cx = cx;

		ListView_InsertColumn(hListView, i, &lvc);
	}
}

void LoadStartupDirectories(HWND hWnd, BOOL selectAll) {
	HWND hListView;
	char szSearchPath[MAX_PATH];
	WIN32_FIND_DATAA ffd;
	HANDLE hFind = INVALID_HANDLE_VALUE;

	hListView = GetDlgItem(hWnd, IDC_LIST);
	ListView_DeleteAllItems(hListView);
	// Construct the path for searching by appending "*" to the startup directory path
	StringCchCopyA(szSearchPath, MAX_PATH, szStartupPath);
	StringCchCatA(szSearchPath, MAX_PATH, "*");

	hFind = FindFirstFileA(szSearchPath, &ffd);
	do {
		if(hFind != INVALID_HANDLE_VALUE) {
			char szFilePath[MAX_PATH];
			char szExtension[MAX_PATH];
			WCHAR fileName[MAX_PATH];
			WCHAR targetPath[MAX_PATH];

			StringCchCopyA(szFilePath, MAX_PATH, szStartupPath);
			StringCchCatA(szFilePath, MAX_PATH, ffd.cFileName);

			mbstowcs_s(nullptr, fileName, ffd.cFileName, MAX_PATH);
			PathRemoveExtension(fileName);

			strcpy_s(szExtension, PathFindExtensionA(szFilePath));
			if(!StrCmpIA(szExtension, ".lnk")) {
				// Get the target of the shortcut
				ResolveLink(hWnd, szFilePath, targetPath, MAX_PATH);
				// Proceed only if the shortcut points to a directory
				if(GetFileAttributes(targetPath) ^ FILE_ATTRIBUTE_DIRECTORY) continue;

				// Add information of the shortcut to the list
				LVITEM lvI;

				lvI.mask = LVIF_TEXT;
				lvI.pszText = fileName;
				lvI.iItem = 0;
				lvI.iSubItem = 0;

				ListView_InsertItem(hListView, &lvI);
				ListView_SetItemText(hListView, 0, 1, targetPath);
				if(selectAll) ListView_SetItemState(hListView, 0, LVIS_SELECTED, LVIS_SELECTED);
			}
		}
	} while(FindNextFileA(hFind, &ffd) != 0);
}

HRESULT CreateLink(LPCWSTR lpszPathObj, LPCSTR lpszPathLink, LPCWSTR lpszDesc) {
	HRESULT hres;
	IShellLink* psl;

	// Get a pointer to the IShellLink interface
	// It is assumed that CoInitialize has already been called
	hres = CoCreateInstance(CLSID_ShellLink, NULL, CLSCTX_INPROC_SERVER, IID_IShellLink, (LPVOID*)&psl);
	if(SUCCEEDED(hres)) {
		IPersistFile* ppf;

		// Set the path to the shortcut target and add the description
		psl->SetPath(lpszPathObj);

		// Query IShellLink for the IPersistFile interface, used for saving the shortcut in persistent storage
		hres = psl->QueryInterface(IID_IPersistFile, (LPVOID*)&ppf);

		if(SUCCEEDED(hres)) {
			WCHAR wsz[MAX_PATH];

			// Ensure that the string is Unicode
			MultiByteToWideChar(CP_ACP, 0, lpszPathLink, -1, wsz, MAX_PATH);

			// Save the link by calling IPersistFile::Save
			hres = ppf->Save(wsz, TRUE);
			ppf->Release();
		}
		psl->Release();
	}
	return hres;
}

HRESULT ResolveLink(HWND hwnd, LPCSTR lpszLinkFile, LPWSTR lpszPath, int iPathBufferSize) {
	HRESULT hres;
	IShellLink* psl;
	WCHAR szGotPath[MAX_PATH];
	WCHAR szDescription[MAX_PATH];
	WIN32_FIND_DATA wfd;

	*lpszPath = 0;
	
	// Get a pointer to the IShellLink interface
	// It is assumed that CoInitialize has already been called
	hres = CoCreateInstance(CLSID_ShellLink, NULL, CLSCTX_INPROC_SERVER, IID_IShellLink, (LPVOID*)&psl);
	if(SUCCEEDED(hres)) {
		IPersistFile* ppf;

		// Get a pointer to the IPersistFile interface
		hres = psl->QueryInterface(IID_IPersistFile, (void**)&ppf);
		if(SUCCEEDED(hres)) {
			WCHAR wsz[MAX_PATH];

			// Ensure that the string is Unicode
			MultiByteToWideChar(CP_ACP, 0, lpszLinkFile, -1, wsz, MAX_PATH);
			// Add code here to check return value from MultiByteWideChar for success
			// Load the shortcut
			hres = ppf->Load(wsz, STGM_READ);
			if(SUCCEEDED(hres)) {
				// Resolve the link.
				hres = psl->Resolve(hwnd, 0);
				if(SUCCEEDED(hres)) {
					// Get the path to the link target
					hres = psl->GetPath(szGotPath, MAX_PATH, (WIN32_FIND_DATA*)&wfd, NULL);
					if(SUCCEEDED(hres)) {
						// Get the description of the target
						hres = psl->GetDescription(szDescription, MAX_PATH);
						if(SUCCEEDED(hres)) {
							hres = StringCbCopy(lpszPath, iPathBufferSize, szGotPath);
							if(SUCCEEDED(hres)) {
								// Handle success
							} else {
								// Handle the error
							}
						}
					}
				}
			}
			// Release the pointer to the IPersistFile interface
			ppf->Release();
		}
		// Release the pointer to the IShellLink interface
		psl->Release();
	}
	return hres;
}

void AddAllOpenExplorers(HWND hWnd) {
	IShellWindows* psw;

	if(SUCCEEDED(CoCreateInstance(CLSID_ShellWindows, NULL, CLSCTX_LOCAL_SERVER, IID_PPV_ARGS(&psw)))) {
		VARIANT v = {VT_I4};

		if(SUCCEEDED(psw->get_Count(&v.lVal))) {
			HWND hCheck = GetDlgItem(hWnd, IDC_CHECK);
			LRESULT checkState = Button_GetState(hCheck);

			// Walk backward to make sure that closing windows will not reorder the array
			while(--v.lVal >= 0) {
				IDispatch *pdisp;

				if(S_OK == psw->Item(v, &pdisp)) {
					IWebBrowser2 *pwb;
					BSTR locationName;
					BSTR locationURL;

					if(SUCCEEDED(pdisp->QueryInterface(IID_PPV_ARGS(&pwb)))) {
						char szLinkPath[MAX_PATH];

						pwb->get_LocationName(&locationName);
						// Prevent the startup directory from being added to itself
						if(!StrCmpI(locationName, L"Startup")) continue;
						StringCchCopyA(szLinkPath, MAX_PATH, szStartupPath);
						StringCchCatA(szLinkPath, MAX_PATH, _com_util::ConvertBSTRToString(locationName));
						StringCchCatA(szLinkPath, MAX_PATH, ".lnk");
						pwb->get_LocationURL(&locationURL);
						
						CreateLink(locationURL, szLinkPath, NULL);
						if(checkState == BST_CHECKED) pwb->Quit();
						pwb->Release();
					}
					pdisp->Release();
				}
			}
		}
		psw->Release();
	}
}

void RemoveAllSelectedLinks(HWND hWnd) {
	HWND hListView = GetDlgItem(hWnd, IDC_LIST);
	int iPos = ListView_GetNextItem(hListView, -1, LVIS_SELECTED);

	do {
		LVITEM lvi;
		WCHAR linkName[MAX_PATH];
		WCHAR linkPath[MAX_PATH];

		lvi.mask = LVIF_TEXT | LVIF_STATE;
		lvi.pszText = linkName;
		lvi.cchTextMax = MAX_PATH;
		lvi.iItem = iPos;
		lvi.stateMask = LVIS_SELECTED;

		ListView_GetItem(hListView, &lvi);

		// Construct the path of the shortcut file
		mbstowcs_s(nullptr, linkPath, szStartupPath, MAX_PATH);
		StringCchCat(linkPath, MAX_PATH, linkName);
		StringCchCat(linkPath, MAX_PATH, L".lnk");

		// Remove the shortcut file
		DeleteFile(linkPath);

		iPos = ListView_GetNextItem(hListView, iPos, LVIS_SELECTED);
	} while(iPos != -1);
}