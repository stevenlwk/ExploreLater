#pragma once

#include "resource.h"

BOOL CALLBACK DialogProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam);

void InitList(HWND hWnd);

void LoadStartupDirectories(HWND hWnd, BOOL selectAll = false);

HRESULT CreateLink(LPCWSTR lpszPathObj, LPCSTR lpszPathLink, LPCWSTR lpszDesc);
HRESULT ResolveLink(HWND hwnd, LPCSTR lpszLinkFile, LPWSTR lpszPath, int iPathBufferSize);

void AddAllOpenExplorers(HWND hWnd);
void RemoveAllSelectedLinks(HWND hWnd);