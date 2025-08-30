/////////////////////////////////////////////////////////////////////////////////
// * FB2K Component: COM Automation and ActiveX Interface                    * //
// * Description:    API Wizard Main Source File                             * //
// * Author:         TT                                                      * //
// * Website:        https://github.com/The-Wizardium/API-Wizard             * //
// * Version:        0.2                                                     * //
// * Dev. started:   12-12-2024                                              * //
// * Last change:    22-12-2024                                              * //
/////////////////////////////////////////////////////////////////////////////////


#include "APIW_PCH.h"
#include "APIW_Main.h"


//////////////////////////////////
// * CONSTRUCTOR & DESTRUCTOR * //
//////////////////////////////////
#pragma region Constructor & Destructor
APIWizardMain::APIWizardMain(HWND hWnd) :
	mainHwnd(hWnd) {
}

APIWizardMain::~APIWizardMain() = default;
#pragma endregion


////////////////////////
// * PUBLIC METHODS * //
////////////////////////
#pragma region Public Methods
int APIWizardMain::GetWindowWidth() {
	RECT rect;
	return GetWindowRect(mainHwnd, &rect) ? rect.right - rect.left : -1;
}

void APIWizardMain::SetWindowWidth(int width) {
	const int height = GetWindowHeight();
	SetWindowPos(mainHwnd, nullptr, 0, 0, width, height, SWP_NOMOVE | SWP_NOZORDER);
}

int APIWizardMain::GetWindowHeight() {
	RECT rect;
	return GetWindowRect(mainHwnd, &rect) ? rect.bottom - rect.top : -1;
}

void APIWizardMain::SetWindowHeight(int height) {
	const int width = GetWindowWidth();
	SetWindowPos(mainHwnd, nullptr, 0, 0, width, height, SWP_NOMOVE | SWP_NOZORDER);
}

void APIWizardMain::SetWindowSize(int width, int height) {
	SetWindowPos(mainHwnd, nullptr, 0, 0, width, height, SWP_NOMOVE | SWP_NOZORDER);
}
#pragma endregion
