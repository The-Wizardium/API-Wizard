/////////////////////////////////////////////////////////////////////////////////
// * FB2K Component: COM Automation and ActiveX Interface                    * //
// * Description:    MyCOM Impl Source File                                  * //
// * Author:         TT                                                      * //
// * Website:        https://github.com/The-Wizardium/API-Wizard             * //
// * Version:        0.1.0                                                   * //
// * Dev. started:   12-12-2024                                              * //
// * Last change:    01-09-2025                                              * //
/////////////////////////////////////////////////////////////////////////////////


#include "APIW_PCH.h"
#include "MyCOM.h"
#include "APIW.h"
#include "APIW_Main.h"


//////////////////////////////////////////
// * MyCOM - CONSTRUCTOR & DESTRUCTOR * //
//////////////////////////////////////////
#pragma region MyCOM - Constructor & Destructor
MyCOM::MyCOM() {
	InterlockedIncrement(&activeObjects);
}

MyCOM::~MyCOM() {
	InterlockedDecrement(&activeObjects);
}
#pragma endregion


///////////////////////////////////////
// * MyCOM - PUBLIC STATIC METHODS * //
///////////////////////////////////////
#pragma region MyCOM - Public Static Methods
HRESULT MyCOM::InitMyCOM() {
	HMODULE hModule = nullptr;
	HRESULT hr = OleInitialize(nullptr);

	if (FAILED(hr)) {
		return LogError(hr, L"API Wizard => MyCOM::InitMyCOM", L"OleInitialize failed", true);
	}

	if (!GetModuleHandleEx(GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS, reinterpret_cast<LPCTSTR>(static_cast<void*>(&InitMyCOM)), &hModule)) {
		OleUninitialize();
		return LogError(E_FAIL, L"API Wizard => MyCOM::InitMyCOM", L"Failed to get module handle", true);
	}

	hr = InitTypeLibrary(hModule);
	if (FAILED(hr)) {
		OleUninitialize();
		return LogError(hr, L"API Wizard => MyCOM::InitMyCOM", L"InitTypeLibrary failed", true);
	}

	hr = RegisterMyCOM();
	if (FAILED(hr)) {
		OleUninitialize();
		return LogError(hr, L"API Wizard => MyCOM::InitMyCOM", L"RegisterMyCOM failed", true);
	}

	return S_OK;
}

HRESULT MyCOM::InitTypeLibrary(HMODULE hModule) {
	std::wstring module_path(MAX_PATH, L'\0');

	DWORD len = GetModuleFileNameW(hModule, &module_path[0], MAX_PATH);
	if (len == 0 || len >= MAX_PATH) {
		return LogError(E_FAIL, L"API Wizard => MyCOM::InitTypeLibrary", L"Failed to retrieve module path", true);
	}

	module_path.resize(len);

	HRESULT hr = LoadTypeLibEx(module_path.c_str(), REGKIND_NONE, &MyCOM::typeLib);
	if (FAILED(hr)) {
		return LogError(hr, L"API Wizard => MyCOM::InitTypeLibrary", L"LoadTypeLibEx failed", true);
	}

	return S_OK;
}

HRESULT MyCOM::RegisterMyCOM(std::wstring_view regMethod) {
	auto hr = S_OK;
	CComPtr<MyCOMClassFactory> pClassFactory = new MyCOMClassFactory();

	if (!pClassFactory) {
		return LogError(E_OUTOFMEMORY, L"API Wizard => MyCOM::RegisterMyCOM", L"Failed to create class factory", true);
	}

	hr = CoRegisterClassObject(MyCOM_CLSID, pClassFactory, CLSCTX_INPROC_SERVER, REGCLS_MULTIPLEUSE, &dwRegister);
	if (FAILED(hr)) {
		return LogError(hr, L"API Wizard => MyCOM::RegisterMyCOM", L"CoRegisterClassObject failed", true);
	}

	if (regMethod == L"RegEntry") {
		hr = RegisterCLSID();
		if (FAILED(hr)) {
			CoRevokeClassObject(dwRegister);
			return LogError(hr, L"API Wizard => MyCOM::RegisterMyCOM", L"RegisterCLSID failed", true);
		}
	}
	else {
		hr = HookCLSIDFromProgID();
		if (FAILED(hr)) {
			CoRevokeClassObject(dwRegister);
			return LogError(hr, L"API Wizard => MyCOM::RegisterMyCOM", L"HookCLSIDFromProgID failed", true);
		}
	}

	return S_OK;
}

HRESULT MyCOM::UnregisterMyCOM(std::wstring_view regMethod) {
	if (dwRegister != 0) {
		HRESULT hr = CoRevokeClassObject(dwRegister);

		if (FAILED(hr)) {
			return LogError(hr, L"API Wizard => MyCOM::UnregisterMyCOM", L"CoRevokeClassObject failed", true);
		}

		if (regMethod == L"RegFree" && original_CLSIDFromProgID) {
			MH_DisableHook(&CLSIDFromProgID);
			MH_Uninitialize();
			original_CLSIDFromProgID = nullptr;
		}
		else if (regMethod == L"RegEntry") {
			UnregisterCLSID();
		}

		dwRegister = 0;
	}

	return S_OK;
}

HRESULT MyCOM::QuitMyCOM(std::wstring_view regMethod) {
	HRESULT hr = UnregisterMyCOM(regMethod);
	OleUninitialize();
	return hr;
}

HRESULT MyCOM::LogError(HRESULT errorCode, const std::wstring& source, const std::wstring& description, bool setErrorInfo) {
	pfc::string8 logMessage;

	logMessage << pfc::stringcvt::string_utf8_from_wide(source.c_str())
		<< ": " << pfc::stringcvt::string_utf8_from_wide(description.c_str())
		<< ", HRESULT: 0x" << pfc::format_hex(errorCode);

	FB2K_console_formatter() << logMessage;

	if (setErrorInfo && FAILED(errorCode)) {
		CComPtr<ICreateErrorInfo> createErrorInfo;

		if (SUCCEEDED(CreateErrorInfo(&createErrorInfo))) {
			createErrorInfo->SetSource(CComBSTR(source.c_str()));
			createErrorInfo->SetDescription(CComBSTR(description.c_str()));

			CComPtr<IErrorInfo> errorInfo;
			if (SUCCEEDED(createErrorInfo->QueryInterface(IID_IErrorInfo, reinterpret_cast<void**>(&errorInfo)))) {
				SetErrorInfo(0, errorInfo);
			}
		}
	}

	return errorCode;
}
#pragma endregion


////////////////////////////////////////
// * MyCOM - PRIVATE STATIC METHODS * //
////////////////////////////////////////
#pragma region MyCOM - COM Registration-Free with MinHook
HRESULT WINAPI MyCOM::MyCLSIDFromProgID(LPCOLESTR lpszProgID, LPCLSID lpclsid) {
	static std::map<std::wstring, CLSID> progidToClsidMap = {
		{ MyCOM_PROGID, MyCOM_CLSID }
	};

	auto it = progidToClsidMap.find(lpszProgID);
	if (it != progidToClsidMap.end()) {
		*lpclsid = it->second;
		return S_OK;
	}

	// Call the original function for other ProgIDs
	if (original_CLSIDFromProgID != nullptr) {
		return original_CLSIDFromProgID(lpszProgID, lpclsid);
	}

	return LogError(REGDB_E_CLASSNOTREG, L"API Wizard => MyCOM::MyCLSIDFromProgID", L"Class not registered", true);
}

HRESULT MyCOM::HookCLSIDFromProgID() {
	if (MH_Initialize() != MH_OK) {
		return LogError(E_FAIL, L"API Wizard => MyCOM::HookCLSIDFromProgID", L"MinHook initialization failed", true);
	}

	if (MH_CreateHook(&CLSIDFromProgID, &MyCLSIDFromProgID, reinterpret_cast<void**>(&original_CLSIDFromProgID)) != MH_OK) {
		MH_Uninitialize();
		return LogError(E_FAIL, L"API Wizard => MyCOM::HookCLSIDFromProgID", L"MinHook creation failed", true);
	}

	if (MH_EnableHook(&CLSIDFromProgID) != MH_OK) {
		MH_Uninitialize();
		return LogError(E_FAIL, L"API Wizard => MyCOM::HookCLSIDFromProgID", L"MinHook enable failed", true);
	}

	return S_OK;
}
#pragma endregion

#pragma region MyCOM - COM Registration with Reg Entry
HRESULT MyCOM::RegisterCLSID() {
	HRESULT hr;
	LPOLESTR clsidString = nullptr;
	hr = StringFromCLSID(MyCOM_CLSID, &clsidString);

	if (FAILED(hr)) {
		return LogError(hr, L"API Wizard => MyCOM::RegisterCLSID", L"StringFromCLSID failed", true);
	}

	CComBSTR clsidBSTR(clsidString);
	CoTaskMemFree(clsidString);

	// Register ProgID
	hr = SetRegistry(HKEY_CLASSES_ROOT, L"MyCOM\\CLSID", nullptr, clsidBSTR);
	if (FAILED(hr)) {
		return LogError(hr, L"API Wizard => MyCOM::RegisterCLSID", L"Failed to set registry for ProgID", true);
	}

	// Register CLSID key
	std::wstring clsidKeyPath = L"CLSID\\" + std::wstring(clsidBSTR);
	std::wstring modulePath(MAX_PATH, L'\0');
	DWORD len = GetModuleFileNameW(core_api::get_my_instance(), &modulePath[0], MAX_PATH);

	if (len == 0) {
		return LogError(E_FAIL, L"API Wizard => MyCOM::RegisterCLSID", L"GetModuleFileNameW failed", true);
	}
	modulePath.resize(len);

	// Set InProcServer32
	hr = SetRegistry(HKEY_CLASSES_ROOT, clsidKeyPath.c_str(), L"InProcServer32", modulePath.c_str());
	if (FAILED(hr)) {
		return LogError(hr, L"API Wizard => MyCOM::RegisterCLSID", L"Failed to set InProcServer32 registry", true);
	}

	// Set ThreadingModel
	hr = SetRegistry(HKEY_CLASSES_ROOT, clsidKeyPath.c_str(), L"ThreadingModel", L"Apartment");
	if (FAILED(hr)) {
		return LogError(hr, L"API Wizard => MyCOM::RegisterCLSID", L"Failed to set ThreadingModel registry", true);
	}

	return S_OK;
}

HRESULT MyCOM::UnregisterCLSID() {
	LPOLESTR clsidString = nullptr;
	HRESULT hr = StringFromCLSID(MyCOM_CLSID, &clsidString);

	if (FAILED(hr)) {
		return LogError(hr, L"API Wizard => MyCOM::UnregisterCLSID", L"StringFromCLSID failed");
	}

	std::wstring clsidKeyPath = L"CLSID\\" + std::wstring(clsidString);
	CoTaskMemFree(clsidString);

	LONG result = RegDeleteTree(HKEY_CLASSES_ROOT, clsidKeyPath.c_str());
	if (result != ERROR_SUCCESS) {
		LogError(HRESULT_FROM_WIN32(result), L"API Wizard => MyCOM::UnregisterCLSID", L"Failed to delete CLSID registry key");
	}

	result = RegDeleteTree(HKEY_CLASSES_ROOT, L"APIWizard");
	if (result != ERROR_SUCCESS) {
		LogError(HRESULT_FROM_WIN32(result), L"API Wizard => MyCOM::UnregisterCLSID", L"Failed to delete MyCOM registry key");
	}

	return S_OK;
}

HRESULT MyCOM::SetRegistry(HKEY rootKey, const wchar_t* subKey, const wchar_t* valueName, const wchar_t* value) {
	CRegKey hKey;
	LONG lResult = hKey.Create(rootKey, subKey);

	if (lResult != ERROR_SUCCESS) {
		return LogError(HRESULT_FROM_WIN32(lResult), L"API Wizard => MyCOM::SetRegistry", L"Failed to create/open registry key", true);
	}

	lResult = hKey.SetStringValue(valueName, value);

	if (lResult != ERROR_SUCCESS) {
		return LogError(HRESULT_FROM_WIN32(lResult), L"API Wizard => MyCOM::SetRegistry", L"Failed to set registry value", true);
	}

	return S_OK;
}
#pragma endregion


/////////////////////////////////////////
// * MyCOM - IUNKNOWN PUBLIC METHODS * //
/////////////////////////////////////////
#pragma region MyCOM - IUnknown Public Methods
STDMETHODIMP MyCOM::QueryInterface(REFIID riid, void** ppvObject) {
	if (riid == IID_IUnknown || riid == IID_IDispatch || riid == MyCOMAPI_IID) {
		*ppvObject = static_cast<IDispatch*>(this);
		AddRef();
		return S_OK;
	}

	*ppvObject = nullptr;
	return LogError(E_NOINTERFACE, L"API Wizard => MyCOM::QueryInterface", L"Interface not supported", true);
}

STDMETHODIMP_(ULONG) MyCOM::AddRef() {
	return InterlockedIncrement(&refCount);
}

STDMETHODIMP_(ULONG) MyCOM::Release() {
	LONG count = InterlockedDecrement(&refCount);

	if (count == 0) {
		delete this;
	}

	return count;
}
#pragma endregion


///////////////////////////////////////////
// * MyCOM - IDISPATCH PRIVATE METHODS * //
///////////////////////////////////////////
#pragma region MyCOM - IDispatch Private Methods
HRESULT MyCOM::InitTypeInfo() {
	if (MyCOM::typeInfo == nullptr && typeLib != nullptr) {
		HRESULT hr = typeLib->GetTypeInfoOfGuid(MyCOMAPI_IID, &MyCOM::typeInfo);

		if (FAILED(hr)) {
			return LogError(hr, L"API Wizard => MyCOM::InitTypeInfo", L"GetTypeInfoOfGuid failed", true);
		}
	}

	return S_OK;
}

STDMETHODIMP MyCOM::GetTypeInfoCount(UINT* pctinfo) {
	*pctinfo = (typeInfo != nullptr) ? 1 : 0;
	return S_OK;
}

STDMETHODIMP MyCOM::GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo** ppTInfo) {
	if (iTInfo != 0) {
		return LogError(DISP_E_BADINDEX, L"API Wizard => MyCOM::GetTypeInfo", L"Invalid type info index", true);
	}

	if (typeInfo == nullptr) {
		HRESULT hr = InitTypeInfo();
		if (FAILED(hr)) {
			return LogError(hr, L"API Wizard => MyCOM::GetTypeInfo", L"InitTypeInfo failed", true);
		}
	}

	*ppTInfo = typeInfo;
	if (*ppTInfo) {
		(*ppTInfo)->AddRef();
	}

	return S_OK;
}

STDMETHODIMP MyCOM::GetIDsOfNames(REFIID riid, LPOLESTR* rgszNames, UINT cNames, LCID lcid, DISPID* rgDispId) {
	if (typeInfo == nullptr) {
		return LogError(E_NOTIMPL, L"API Wizard => MyCOM::GetIDsOfNames", L"Type info not initialized", true);
	}

	return typeInfo->GetIDsOfNames(rgszNames, cNames, rgDispId);
}

STDMETHODIMP MyCOM::Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS* pDispParams, VARIANT* pVarResult, EXCEPINFO* pExcepInfo, UINT* puArgErr) {
	if (typeInfo == nullptr) {
		return LogError(E_NOTIMPL, L"API Wizard => MyCOM::Invoke", L"Type info not initialized", true);
	}

	return typeInfo->Invoke(static_cast<IDispatch*>(this), dispIdMember, wFlags, pDispParams, pVarResult, pExcepInfo, puArgErr);
}
#pragma endregion


/////////////////////////////////////////
// * MyCOM - PUBLIC API - PROPERTIES * //
/////////////////////////////////////////
#pragma region MyCOM - Public API - Properties
STDMETHODIMP MyCOM::get_WindowWidth(long* pValue) {
	if (!pValue) {
		return LogError(E_POINTER, L"API Wizard => MyCOM::get_WindowWidth", L"Invalid pointer");
	}
	if (!APIWizard::Main()) {
		return LogError(E_UNEXPECTED, L"API Wizard => MyCOM::get_WindowWidth", L"Window not initialized");
	}

	*pValue = APIWizard::Main()->GetWindowWidth();
	return S_OK;
}

STDMETHODIMP MyCOM::put_WindowWidth(const VARIANT& newValue) {
	if (!APIWizard::Main()) {
		return LogError(E_UNEXPECTED, L"API Wizard => MyCOM::get_WindowWidth", L"Window not initialized");
	}

	long convertedValue = 1200; // Start with default value
	if (newValue.vt == VT_I4 && newValue.intVal > 0) { // Check passed argument value
		convertedValue = newValue.intVal; // No invalid value, assign from passed argument
	}

	APIWizard::Main()->SetWindowWidth(convertedValue);
	return S_OK;
}

STDMETHODIMP MyCOM::get_WindowHeight(long* pValue) {
	if (!pValue) {
		return LogError(E_POINTER, L"API Wizard => MyCOM::get_WindowHeight", L"Invalid pointer");
	}
	if (!APIWizard::Main()) {
		return LogError(E_UNEXPECTED, L"API Wizard => MyCOM::get_WindowHeight", L"Window not initialized");
	}

	*pValue = APIWizard::Main()->GetWindowHeight();
	return S_OK;
}

STDMETHODIMP MyCOM::put_WindowHeight(const VARIANT& newValue) {
	if (!APIWizard::Main()) {
		return LogError(E_UNEXPECTED, L"API Wizard => MyCOM::get_WindowWidth", L"Window not initialized");
	}

	long convertedValue = 800; // Start with default value
	if (newValue.vt == VT_I4 && newValue.intVal > 0) { // Check passed argument value
		convertedValue = newValue.intVal; // No invalid value, assign from passed argument
	}

	APIWizard::Main()->SetWindowHeight(convertedValue);
	return S_OK;
}
#pragma endregion


//////////////////////////////////////
// * MyCOM - PUBLIC API - METHODS * //
//////////////////////////////////////
#pragma region MyCOM - Public API - Methods
STDMETHODIMP MyCOM::SetWindowSize(int width, int height) {
	if (!APIWizard::Main()) {
		return LogError(E_UNEXPECTED, L"API Wizard => MyCOM::SetWindowSize", L"Window not initialized");
	}

	// Invalid values passed from arguments, set to defaults
	if (width <= 0 || height <= 0) {
		width = 1000;
		height = 800;
	}

	APIWizard::Main()->SetWindowSize(width, height);
	return S_OK;
}

STDMETHODIMP MyCOM::PrintMessage() {
	MessageBox(nullptr, L"COM/ActiveX interface is working!", L"Message", MB_OK);
	return S_OK;
}
#pragma endregion


///////////////////////////////////////////////////////
// * MyCOM CLASS FACTORY - IUNKNOWN PUBLIC METHODS * //
///////////////////////////////////////////////////////
#pragma region MyCOM Class Factory - IUnknown Public Methods
STDMETHODIMP MyCOMClassFactory::QueryInterface(REFIID riid, void** ppvObject) {
	if (riid == IID_IUnknown || riid == IID_IClassFactory) {
		*ppvObject = static_cast<IClassFactory*>(this);
		AddRef();
		return S_OK;
	}

	*ppvObject = nullptr;
	return MyCOM::LogError(E_NOINTERFACE, L"API Wizard => MyCOMClassFactory::QueryInterface", L"Interface not supported", true);
}

STDMETHODIMP_(ULONG) MyCOMClassFactory::AddRef() {
	return InterlockedIncrement(&refCount);
}

STDMETHODIMP_(ULONG) MyCOMClassFactory::Release() {
	LONG count = InterlockedDecrement(&refCount);

	if (count == 0) {
		delete this;
	}

	return count;
}
#pragma endregion


/////////////////////////////////////////////////////////////
// * MyCOM CLASS FACTORY - ICLASSFACTORY PRIVATE METHODS * //
/////////////////////////////////////////////////////////////
#pragma region MyCOM Class Factory - IClassFactory Private Methods
STDMETHODIMP MyCOMClassFactory::CreateInstance(IUnknown* pUnkOuter, REFIID riid, void** ppvObject) {
	if (pUnkOuter != nullptr) {
		return MyCOM::LogError(CLASS_E_NOAGGREGATION, L"API Wizard => MyCOMClassFactory::CreateInstance", L"Aggregation not supported", true);
	}

	CComPtr<MyCOM> pMyCOM = new MyCOM();
	return pMyCOM->QueryInterface(riid, ppvObject);
}

STDMETHODIMP MyCOMClassFactory::LockServer(BOOL fLock) {
	fLock ? InterlockedIncrement(&MyCOM::serverLocks) :
		InterlockedDecrement(&MyCOM::serverLocks);

	return S_OK;
}
#pragma endregion
