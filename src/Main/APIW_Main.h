/////////////////////////////////////////////////////////////////////////////////
// * FB2K Component: COM Automation and ActiveX Interface                    * //
// * Description:    API Wizard Main Header File                             * //
// * Author:         TT                                                      * //
// * Website:        https://github.com/The-Wizardium/API-Wizard             * //
// * Version:        0.2                                                     * //
// * Dev. started:   12-12-2024                                              * //
// * Last change:    22-12-2024                                              * //
/////////////////////////////////////////////////////////////////////////////////


#pragma once


/////////////////////////
// * API WIZARD MAIN * //
/////////////////////////
#pragma region API Wizard Main
class APIWizardMain {
public:
	explicit APIWizardMain(HWND hWnd);
	virtual ~APIWizardMain();

	HWND mainHwnd = nullptr;

	int GetWindowWidth();
	void SetWindowWidth(int width);
	int GetWindowHeight();
	void SetWindowHeight(int height);
	void SetWindowSize(int width, int height);
};
#pragma endregion
