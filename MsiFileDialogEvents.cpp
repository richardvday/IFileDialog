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

#include "stdafx.h"
#include "MsiFileDialogEvents.h"

namespace msi = MithrilSoftware;

MsiFileDialogEvents::~MsiFileDialogEvents() {
	if (ptrCallMe && Bit(uiEventCode, MsiFDEvent::OnExit)) {
		ptrCallMe->NotifyMe(MsiFDEvent::OnExit);
	}
	if (0 != _cRef) {
		msi::DebugOutput(std::source_location::current(), "dangling reference! (0 != _cRef)");
		EXCEPTION_RUNTIME_ERROR("dangling reference! (0 != _cRef)");
	}
	assert(_cRef == 0);
	delete this;
};

MsiFileDialogEventsNotify::MsiFileDialogEventsNotify() : dialog_ptr(nullptr),pSink(nullptr),active(false), rCookie(0),hrLastError(S_OK)
{
	OleInitialize(nullptr);
}

MsiFileDialogEventsNotify::MsiFileDialogEventsNotify(MsiFileDialogEvents* pEventSink ) : dialog_ptr(nullptr), pSink(pEventSink), active(false), rCookie(0), hrLastError(S_OK) {
}

MsiFileDialogEventsNotify::~MsiFileDialogEventsNotify() {
	dialog_ptr->Unadvise(rCookie);
	if (pSink) {
		pSink->Release();
	}
	OleUninitialize(); // we called OleInitialize(0) so call OleUninitialize();
}

void MsiFileDialogEventsNotify::Create(void) {
	assert(nullptr == dialog_ptr);
	if (nullptr != dialog_ptr) {
		msi::DebugOutput(std::source_location::current(), "IFileDialog* already exist! Cant make twice");
		return; // we can continue just a bug getting here
	}
	else {
		HRESULT hr = CoCreateInstance(CLSID_FileOpenDialog, NULL, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&dialog_ptr));
		if (FAILED(hr)) {
			active = false;
			hrLastError = hr;
			msi::DebugOutput(std::source_location::current(), "CoCreateInstance failed result is ", msi::GetResultString(hr));
			throw std::exception("CoCreateInstance failed"); // we cannot continue!
		}
		active = true;
		Advise();
	}
	return;
}

bool MsiFileDialogEventsNotify::Advise(void) {
	if (active) {
		assert(nullptr != pSink);
		if (nullptr == pSink) {
			msi::DebugOutput(std::source_location::current(), "IFileDialog* is nullptr must create first");
			EXCEPTION_RUNTIME_ERROR("IFileDialog* is nullptr must create first");// we cannot continue!
		}
		pSink->NotifyMeOn(MsiFDEvent::OnExit, this);
		dialog_ptr->Advise(pSink, &rCookie);
	}
	return active;
}
     
bool MsiFileDialogEventsNotify::NotifyMe( MsiFDEvent EventCode) {
	// a switch would be better if your doing more then one or two things
	if (BIT(EventCode, MsiFDEvent::OnExit)) { // only one we care about for now
		dialog_ptr = nullptr;
		active = false;
	}
	/*
	if (BIT(EventCode, MsiFDEvent::ItemSelected)) {}
	if (BIT(EventCode, MsiFDEvent::ButtonClicked)) {}
	if (BIT(EventCode, MsiFDEvent::ButtonToggled) ){}
	if (BIT(EventCode, MsiFDEvent::ControlActivating)) {}
	if (BIT(EventCode, MsiFDEvent::FileOk)) {}
	if (BIT(EventCode, MsiFDEvent::FolderChanging)) {}
	if (BIT(EventCode, MsiFDEvent::FolderChange)) {}
	if (BIT(EventCode, MsiFDEvent::SelectionChange)) {}
	if (BIT(EventCode, MsiFDEvent::ShareViolation)) {}
	if (BIT(EventCode, MsiFDEvent::TypeChange)) {}
	if (BIT(EventCode, MsiFDEvent::OverWrite)) {}
	*/
	return true; // always want to be notifed
}

HRESULT MsiFileDialogEventsNotify::AddPlace(IShellItem *psi, FDAP fdap) {
	if (active) {
		HRESULT hr = dialog_ptr->AddPlace(psi, fdap);
		if (FAILED(hr)) {
			hrLastError = hr;
			msi::DebugOutput(std::source_location::current(), " failed result is ", msi::GetResultString(hr));
			return E_FAIL;
		}
		return hr;
	}
	return E_FAIL;
}

HRESULT MsiFileDialogEventsNotify::ClearClientData() {
	if (active) {
		HRESULT hr = dialog_ptr->ClearClientData();
		if (FAILED(hr)) {
			hrLastError = hr;
			msi::DebugOutput(std::source_location::current(), " failed result is ", msi::GetResultString(hr));
			return E_FAIL;
		}
		return hr;
	}
	return E_FAIL;
}
HRESULT MsiFileDialogEventsNotify::Close(HRESULT hres) {
	if (active) {
		HRESULT hr = dialog_ptr->Close(hres);
		if (FAILED(hr)) {
			hrLastError = hr;
			msi::DebugOutput(std::source_location::current(), " failed result is ", msi::GetResultString(hr));
			return E_FAIL;
		}
		return hr;
	}
	return E_FAIL;
}
HRESULT MsiFileDialogEventsNotify::GetCurrentSelection(IShellItem **ppsi) {
	if (active) {
		HRESULT hr = dialog_ptr->GetCurrentSelection(ppsi);
		if (FAILED(hr)) {
			hrLastError = hr;
			msi::DebugOutput(std::source_location::current(), " failed result is ", msi::GetResultString(hr));
			return E_FAIL;
		}
		return hr;
	}
	return E_FAIL;

}
HRESULT MsiFileDialogEventsNotify::GetFileName(LPWSTR *pszName) {
	if (active) {
		HRESULT hr = dialog_ptr->GetFileName(pszName);
		if (FAILED(hr)) {
			hrLastError = hr;
			msi::DebugOutput(std::source_location::current(), " failed result is ", msi::GetResultString(hr));
			return E_FAIL;
		}
		return hr;
	}
	return E_FAIL;
}
HRESULT MsiFileDialogEventsNotify::GetFileTypeIndex(UINT *piFileType) {
	if (active) {
		HRESULT hr = dialog_ptr->GetFileTypeIndex(piFileType);
		if (FAILED(hr)) {
			hrLastError = hr;
			msi::DebugOutput(std::source_location::current(), " failed result is ", msi::GetResultString(hr));
			return E_FAIL;
		}
		return hr;
	}
	return E_FAIL;
}
HRESULT MsiFileDialogEventsNotify::GetFolder(IShellItem **ppsi) {
	if (active) {
		HRESULT hr = dialog_ptr->GetFolder(ppsi);
		if (FAILED(hr)) {
			hrLastError = hr;
			msi::DebugOutput(std::source_location::current(), " failed result is ", msi::GetResultString(hr));
			return E_FAIL;
		}
		return hr;
	}
	return E_FAIL;
}
HRESULT MsiFileDialogEventsNotify::GetOptions(FILEOPENDIALOGOPTIONS *pfos) {
	if (active) {
		HRESULT hr = dialog_ptr->GetOptions(pfos);
		if (FAILED(hr)) {
			hrLastError = hr;
			msi::DebugOutput(std::source_location::current(), " failed result is ", msi::GetResultString(hr));
			return E_FAIL;
		}
		return hr;
	}
	return E_FAIL;
}

HRESULT MsiFileDialogEventsNotify::GetResult(IShellItem **ppsi) {
	if (active) {
		HRESULT hr = dialog_ptr->GetResult(ppsi);
		if (FAILED(hr)) {
			hrLastError = hr;
			msi::DebugOutput(std::source_location::current(), " failed result is ", msi::GetResultString(hr));
			return E_FAIL;
		}
		return hr;
	}
	return E_FAIL;
}

HRESULT MsiFileDialogEventsNotify::SetClientGuid(REFGUID guid) {
	if (active) {
		HRESULT hr = dialog_ptr->SetClientGuid(guid);
		if (FAILED(hr)) {
			hrLastError = hr;
			msi::DebugOutput(std::source_location::current(), " failed result is ", msi::GetResultString(hr));
			return E_FAIL;
		}
		return hr;
	}
	return E_FAIL;
}

HRESULT MsiFileDialogEventsNotify::SetDefaultExtension(LPCWSTR pszDefaultExtension) {
	if (active) {
		HRESULT hr = dialog_ptr->SetDefaultExtension(pszDefaultExtension);
		if (FAILED(hr)) {
			hrLastError = hr;
			msi::DebugOutput(std::source_location::current(), " failed result is ", msi::GetResultString(hr));
			return E_FAIL;
		}
		return hr;
	}
	return E_FAIL;
}

HRESULT MsiFileDialogEventsNotify::SetDefaultFolder(IShellItem *psi) {
	if (active) {
		HRESULT hr = dialog_ptr->SetDefaultFolder(psi);
		if (FAILED(hr)) {
			hrLastError = hr;
			msi::DebugOutput(std::source_location::current(), " failed result is ", msi::GetResultString(hr));
			return E_FAIL;
		}
		return hr;
	}
	return E_FAIL;
}

HRESULT MsiFileDialogEventsNotify::SetFileName(LPCWSTR pszName) {
	if (active) {
		HRESULT hr = dialog_ptr->SetFileName(pszName);
		if (FAILED(hr)) {
			hrLastError = hr;
			msi::DebugOutput(std::source_location::current(), " failed result is ", msi::GetResultString(hr));
			return E_FAIL;
		}
		return hr;
	}
	return E_FAIL;
}

HRESULT MsiFileDialogEventsNotify::SetFileNameLabel(LPCWSTR pszLabel) {
	if (active) {
		HRESULT hr = dialog_ptr->SetFileNameLabel(pszLabel);
		if (FAILED(hr)) {
			hrLastError = hr;
			msi::DebugOutput(std::source_location::current(), " failed result is ", msi::GetResultString(hr));
			return E_FAIL;
		}
		return hr;
	}
	return E_FAIL;
}

HRESULT MsiFileDialogEventsNotify::SetFileTypeIndex(UINT iFileType) {
	if (active) {
		HRESULT hr = dialog_ptr->SetFileTypeIndex(iFileType);
		if (FAILED(hr)) {
			hrLastError = hr;
			msi::DebugOutput(std::source_location::current(), " failed result is ", msi::GetResultString(hr));
			return E_FAIL;
		}
		return hr;
	}
	return E_FAIL;
}

HRESULT MsiFileDialogEventsNotify::SetFileTypes(UINT cFileTypes, const COMDLG_FILTERSPEC *rgFilterSpec) {
	if (active) {
		HRESULT hr = dialog_ptr->SetFileTypes(cFileTypes, rgFilterSpec);
		if (FAILED(hr)) {
			hrLastError = hr;
			msi::DebugOutput(std::source_location::current(), " failed result is ", msi::GetResultString(hr));
			return E_FAIL;
		}
		return hr;
	}
	return E_FAIL;
}

HRESULT MsiFileDialogEventsNotify::SetFilter(IShellItemFilter *pFilter) {
	if (active) {
		HRESULT hr = dialog_ptr->SetFilter(pFilter);
		if (FAILED(hr)) {
			hrLastError = hr;
			msi::DebugOutput(std::source_location::current(), " failed result is ", msi::GetResultString(hr));
			return E_FAIL;
		}
		return hr;
	}
	return E_FAIL;
}

HRESULT MsiFileDialogEventsNotify::SetFolder(IShellItem *psi) {
	if (active) {
		HRESULT hr = dialog_ptr->SetFolder(psi);
		if (FAILED(hr)) {
			hrLastError = hr;
			msi::DebugOutput(std::source_location::current(), " failed result is ", msi::GetResultString(hr));
			return E_FAIL;
		}
		return hr;
	}
	return E_FAIL;
}

HRESULT MsiFileDialogEventsNotify::SetOkButtonLabel(LPCWSTR pszText) {
	if (active) {
		HRESULT hr = dialog_ptr->SetOkButtonLabel(pszText);
		if (FAILED(hr)) {
			hrLastError = hr;
			msi::DebugOutput(std::source_location::current(), " failed result is ", msi::GetResultString(hr));
			return E_FAIL;
		}
		return hr;
	}
	return E_FAIL;
}

HRESULT MsiFileDialogEventsNotify::Show(HWND hwndOwner) {
	if (active) {
		HRESULT hr = dialog_ptr->Show(hwndOwner);
		if (FAILED(hr)) {
			hrLastError = hr;
			msi::DebugOutput(std::source_location::current(), " failed result is ", msi::GetResultString(hr));
			return E_FAIL;
		}
		return hr;
	}
	return E_FAIL;
}

HRESULT MsiFileDialogEventsNotify::SetOptions(FILEOPENDIALOGOPTIONS fos) {
	if (active) {
		HRESULT hr = dialog_ptr->SetOptions(fos);
		if (FAILED(hr)) {
			hrLastError = hr;
			msi::DebugOutput(std::source_location::current(), " failed result is ", msi::GetResultString(hr));
			return E_FAIL;
		}
		return hr;
	}
	return E_FAIL;
}

HRESULT MsiFileDialogEventsNotify::SetTitle(LPCWSTR pszTitle) {
	if (active) {
		HRESULT hr = dialog_ptr->SetTitle(pszTitle);
		if (FAILED(hr)) {
			hrLastError = hr;
			msi::DebugOutput(std::source_location::current(), " failed result is ", msi::GetResultString(hr));
			return E_FAIL;
		}
		return hr;
	}
	return E_FAIL;
}

HRESULT MsiFileDialogEventsNotify::Unadvise() {

	if (dialog_ptr && active) {
		HRESULT hr = dialog_ptr->Unadvise(rCookie);
		if (FAILED(hr)) {
			hrLastError = hr;
			msi::DebugOutput(std::source_location::current(), " failed result is ", msi::GetResultString(hr));
			return E_FAIL;
		}
		else {
			pSink = nullptr; // self deletes
		}
		return hr;
	}
	return E_FAIL;
}
