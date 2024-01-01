////////////////////////////////////////////////////////////////
// Copyright 2018 Richard V. Day
// Email: richardvday@gmail.com
//
// As one of my favorite programmers used to say
// If this code works, it was written by me.
// If not, I don't know who wrote it.
//
// rest in peace Paul
//
//
//
//			THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, 
//			EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED
//			WARRANTIES OF MERCHANTABILITY AND / OR FITNESS FOR A PARTICULAR PURPOSE.
//                 
//////////////////////////////////////////////////////////////////////

#ifndef MsiFileDialogEvents_h_included
#define MsiFileDialogEvents_h_included
#pragma once

class MsiFileDialogEvents;
class MsiFileDialogEventsHelper;
class MsiFileDialogEventsNotify;

typedef enum _MsiFDEvent_ {
	None = 0,
	ItemSelected = (1 << 1),
	ButtonClicked = (1 << 2),
	ButtonToggled = (1 << 3),
	ControlActivating = (1 << 4),
	FileOk = (1 << 5),
	FolderChanging = (1 << 6),
	FolderChange = (1 << 7),
	SelectionChange = (1 << 8),
	ShareViolation = (1 << 9),
	TypeChange = (1 << 10),
	OverWrite = (1 << 11),
	OnExit = (1 << 12) // only using this for now
} MsiFDEvent;

/*
				you might say hey why not just merge these classes already right ?
				well theres a reason.... com can just disconnect when it feels like it for one thing.
				Then that class above is kaput no references it deletes good bye
				now it tells the Helper class below hey ime GONE and we can handle that.
*/
//template< class T = MsiFileDialogEvents>
class MsiFileDialogEventsNotify {
private:
	MsiFileDialogEventsNotify(); // basic init
public:
	MsiFileDialogEventsNotify(MsiFileDialogEvents*); // basic init
	~MsiFileDialogEventsNotify();
	void Create(void); // real work done here may throw std:exception
	bool Advise(void); // we want feedback for events
	HRESULT Unadvise(); // now we dont
	HRESULT GetLastError() { return hrLastError; }
	bool StatusGood() { return active; }
	virtual bool NotifyMe(MsiFDEvent EventCode); // return false if you no longer wish to be notified

	 // treat like IFileDialog* if you want to trivial to add
	IFileDialog* operator->() { return  dialog_ptr; }
	operator IFileDialog*() { return  dialog_ptr; }

	/// IFileDialog* wrapper functions with error checking/reporting
	HRESULT AddPlace(IShellItem *psi, FDAP fdap);
	HRESULT ClearClientData();
	HRESULT Close(HRESULT hres);
	HRESULT GetCurrentSelection(IShellItem **ppsi);
	HRESULT GetFileName(LPWSTR *pszName);
	HRESULT GetFileTypeIndex(UINT *piFileType);
	HRESULT GetFolder(IShellItem **ppsi);
	HRESULT GetOptions(FILEOPENDIALOGOPTIONS *pfos);
	HRESULT GetResult(IShellItem **ppsi);
	HRESULT SetClientGuid(REFGUID guid);
	HRESULT SetDefaultExtension(LPCWSTR pszDefaultExtension);
	HRESULT SetDefaultFolder(IShellItem *psi);
	HRESULT SetFileName(LPCWSTR pszName);
	HRESULT SetFileNameLabel(LPCWSTR pszLabel);
	HRESULT SetFileTypeIndex(UINT iFileType);
	HRESULT SetFileTypes(UINT cFileTypes, const COMDLG_FILTERSPEC *rgFilterSpec);
	HRESULT SetFilter(IShellItemFilter *pFilter);
	HRESULT SetFolder(IShellItem *psi);
	HRESULT SetOkButtonLabel(LPCWSTR pszText);
	HRESULT Show(HWND hwndOwner);
	HRESULT SetOptions(FILEOPENDIALOGOPTIONS fos);
	HRESULT SetTitle(LPCWSTR pszTitle);

protected:
	HRESULT hrLastError; // last error if any
	MsiFileDialogEvents*  pSink;       /// our event sink
	CComPtr<IFileDialog>  dialog_ptr;
	DWORD rCookie;
	bool active;
};


// this is a base class it does nothing by itself, derive from it DONT modify this one for your usage ! Just implement what you need that way :)
class MsiFileDialogEvents : public IFileDialogEvents, public IFileDialogControlEvents
{
protected:
	MsiFileDialogEvents() : uiEventCode(MsiFDEvent::None), _cRef(1), ptrCallMe(nullptr) { };
public:
	MsiFileDialogEvents* NotifyMeOn( MsiFDEvent EventCode, MsiFileDialogEventsNotify* CallMe)	{
		assert(nullptr != CallMe);
		uiEventCode = EventCode;
		ptrCallMe = CallMe;
		return this;
	}

	// IUnknown
	IFACEMETHODIMP QueryInterface(REFIID riid, void** ppv) {
		assert(nullptr != ppv);
		static const QITAB qit[] = {
		   QITABENT(MsiFileDialogEvents, IFileDialogEvents),
		   QITABENT(MsiFileDialogEvents, IFileDialogControlEvents),
		   { 0 },
		};
		return QISearch(this, qit, riid, ppv);
	}
	IFACEMETHODIMP_(ULONG) AddRef()
	{
		return InterlockedIncrement(&_cRef);
	}

	IFACEMETHODIMP_(ULONG) Release()
	{
		ULONG cRef = InterlockedDecrement(&_cRef);
		if ( 0 == cRef)
		{
			delete this;
		}
		return cRef;
	}
	// IFileDialogEvents

	IFACEMETHODIMP OnFileOk(IFileDialog *pfd) { return E_NOTIMPL; }
	IFACEMETHODIMP OnFolderChanging(IFileDialog *, IShellItem *) { return E_NOTIMPL; }
	IFACEMETHODIMP OnFolderChange(IFileDialog *) { return E_NOTIMPL; }
	IFACEMETHODIMP OnSelectionChange(IFileDialog *) { return E_NOTIMPL; }
	IFACEMETHODIMP OnShareViolation(IFileDialog *, IShellItem *, FDE_SHAREVIOLATION_RESPONSE *) { return E_NOTIMPL; }
	IFACEMETHODIMP OnTypeChange(IFileDialog *) { return E_NOTIMPL; }
	IFACEMETHODIMP OnOverwrite(IFileDialog *, IShellItem *, FDE_OVERWRITE_RESPONSE *) { return E_NOTIMPL; }

	// IFileDialogControlEvents methods
	IFACEMETHODIMP OnItemSelected(IFileDialogCustomize *pfdc, DWORD dwIDCtl, DWORD dwIDItem) { return E_NOTIMPL; };
	IFACEMETHODIMP OnButtonClicked(IFileDialogCustomize *, DWORD) { return E_NOTIMPL; };
	IFACEMETHODIMP OnCheckButtonToggled(IFileDialogCustomize *, DWORD, BOOL) { return E_NOTIMPL; };
	IFACEMETHODIMP OnControlActivating(IFileDialogCustomize *, DWORD) { return E_NOTIMPL; };
	
protected: // only supporting ONE callback for now, can add more later with a pair or tuple or something... to clarify.. One sink but multiple events possible using bitmask
	MsiFDEvent uiEventCode;
	MsiFileDialogEventsNotify* ptrCallMe;
	~MsiFileDialogEvents();
	ULONG _cRef;
};


#endif //#ifndef MsiFileDialogEvents_h_included
