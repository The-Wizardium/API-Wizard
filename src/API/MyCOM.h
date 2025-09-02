/////////////////////////////////////////////////////////////////////////////////
// * FB2K Component: COM Automation and ActiveX Interface                    * //
// * Description:    MyCOM Impl Header File                                  * //
// * Author:         TT                                                      * //
// * Website:        https://github.com/The-Wizardium/API-Wizard             * //
// * Version:        0.1.0                                                   * //
// * Dev. started:   12-12-2024                                              * //
// * Last change:    01-09-2025                                              * //
/////////////////////////////////////////////////////////////////////////////////


#pragma once


///////////////
// * MyCOM * //
///////////////
#pragma region MyCOM
class MyCOM : public IDispatch {
public:
	MyCOM();
	~MyCOM();

	using CLSIDFromProgID_t = HRESULT(WINAPI*)(LPCOLESTR, LPCLSID);

	// * STATIC MEMBERS * //
	static inline IID MyCOMAPI_IID = { 0xB2C3D4E5, 0xF678, 0x40A1, { 0x82, 0x34, 0x67, 0x89, 0xAB, 0xCD, 0xEF, 0x01 } };
	static inline IID MyCOMLIB_IID = { 0xA1B2C3D4, 0xE5F6, 0x4780, { 0x92, 0x34, 0x56, 0x78, 0x9A, 0xBC, 0xDE, 0xF1 } };
	static inline CLSID MyCOM_CLSID = { 0xC3D4E5F6, 0x4780, 0x91B2, { 0x34, 0x56, 0x78, 0x9A, 0xBC, 0xDE, 0xFF, 0x02 } };
	static inline LPCOLESTR MyCOM_PROGID = OLESTR("APIWizard");

	static inline LONG activeObjects = 0;
	static inline DWORD dwRegister = 0;
	static inline LONG serverLocks = 0;
	static inline CLSIDFromProgID_t original_CLSIDFromProgID = nullptr;
	static inline CComPtr<ITypeLib> typeLib = nullptr;
	static inline CComPtr<ITypeInfo> typeInfo = nullptr;

	// * STATIC METHODS * //
	static HRESULT InitMyCOM();
	static HRESULT InitTypeLibrary(HMODULE hModule);
	static HRESULT RegisterMyCOM(std::wstring_view regMethod = L"RegFree"); // "RegFree" or "RegEntry"
	static HRESULT UnregisterMyCOM(std::wstring_view regMethod = L"RegFree"); // "RegFree" or "RegEntry"
	static HRESULT QuitMyCOM(std::wstring_view regMethod = L"RegFree"); // "RegFree" or "RegEntry"
	static HRESULT LogError(HRESULT errorCode, const std::wstring& source, const std::wstring& description, bool setErrorInfo = false);

	// * IUNKNOWN METHODS * //
	STDMETHODIMP QueryInterface(REFIID riid, void** ppvObject) override;
	STDMETHODIMP_(ULONG) AddRef() override;
	STDMETHODIMP_(ULONG) Release() override;

	// * PUBLIC API - PROPERTIES * //
	STDMETHOD(get_WindowWidth)(long* pValue);
	STDMETHOD(put_WindowWidth)(const VARIANT& newValue);
	STDMETHOD(get_WindowHeight)(long* pValue);
	STDMETHOD(put_WindowHeight)(const VARIANT& newValue);

	// * PUBLIC API - METHODS * //
	STDMETHOD(SetWindowSize)(int width, int height);
	STDMETHOD(PrintMessage)();

private:
	LONG refCount = 0;

	// * STATIC METHODS * //
	static HRESULT WINAPI MyCLSIDFromProgID(LPCOLESTR lpszProgID, LPCLSID lpclsid); // "RegFree" with MinHook method
	static HRESULT HookCLSIDFromProgID(); // "RegFree" with MinHook method
	static HRESULT RegisterCLSID(); // "RegEntry" method
	static HRESULT UnregisterCLSID(); // "RegEntry" method
	static HRESULT SetRegistry(HKEY rootKey, const wchar_t* subKey, const wchar_t* valueName, const wchar_t* value); // "RegEntry" method

	// * IDISPATCH METHODS * //
	static HRESULT InitTypeInfo();
	STDMETHODIMP GetTypeInfoCount(UINT* pctinfo) override;
	STDMETHODIMP GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo** ppTInfo) override;
	STDMETHODIMP GetIDsOfNames(REFIID riid, LPOLESTR* rgszNames, UINT cNames, LCID lcid, DISPID* rgDispId) override;
	STDMETHODIMP Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS* pDispParams, VARIANT* pVarResult, EXCEPINFO* pExcepInfo, UINT* puArgErr) override;
};
#pragma endregion


/////////////////////////////
// * MyCOM CLASS FACTORY * //
/////////////////////////////
#pragma region MyCOM Class Factory
class MyCOMClassFactory : public IClassFactory {
public:
	MyCOMClassFactory() = default;
	~MyCOMClassFactory() = default;

	// * IUNKNOWN METHODS * //
	STDMETHODIMP QueryInterface(REFIID riid, void** ppvObject) override;
	STDMETHODIMP_(ULONG) AddRef() override;
	STDMETHODIMP_(ULONG) Release() override;

private:
	LONG refCount = 0;

	// * ICLASSFACTORY METHODS * //
	STDMETHODIMP CreateInstance(IUnknown* pUnkOuter, REFIID riid, void** ppvObject) override;
	STDMETHODIMP LockServer(BOOL fLock) override;
};
#pragma endregion
