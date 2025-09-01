/////////////////////////////////////////////////////////////////////////////////
// * FB2K Component: COM Automation and ActiveX Interface                    * //
// * Description:    API Wizard Header File                                  * //
// * Author:         TT                                                      * //
// * Website:        https://github.com/The-Wizardium/API-Wizard             * //
// * Version:        0.1.0                                                   * //
// * Dev. started:   12-12-2024                                              * //
// * Last change:    01-09-2025                                              * //
/////////////////////////////////////////////////////////////////////////////////


#pragma once
#include "APIW_Main.h"


////////////////////
// * API WIZARD * //
////////////////////
#pragma region API Wizard
class APIWizard {
public:
	APIWizard() = default;
	~APIWizard() = default;

	static void InitAPIWizard();
	static void QuitAPIWizard();

	// Public getters for external access
	static HWND MainWindow() { return mainHwnd; }
	static APIWizardMain* Main() { return apiWizardMain.get(); }

private:
	static inline HWND mainHwnd = nullptr;
	static inline std::unique_ptr<APIWizardMain> apiWizardMain = nullptr;
};
#pragma endregion
