/////////////////////////////////////////////////////////////////////////////////
// * FB2K Component: COM Automation and ActiveX Interface                    * //
// * Description:    API Wizard Source File                                  * //
// * Author:         TT                                                      * //
// * Website:        https://github.com/The-Wizardium/API-Wizard             * //
// * Version:        0.2                                                     * //
// * Dev. started:   12-12-2024                                              * //
// * Last change:    22-12-2024                                              * //
/////////////////////////////////////////////////////////////////////////////////


#include "APIW_PCH.h"
#include "APIW.h"
#include "MyCOM.h"


/////////////////////////
// * COMPONENT SETUP * //
/////////////////////////
#pragma region Component Setup
// Declaration of your component's version information
DECLARE_COMPONENT_VERSION(
	"API Wizard", "0.1",
	"API Wizard \n"
	"The Golden Spell of Automation Rituals \n"
	"A Sacred Chapter of The Wizardium \n"
	"https://github.com/The-Wizardium \n\n"
	"Forged in radiant gold within the Holy Foobar Land, API Wizard unlocks the ActiveX realm.\n"
	"Cast automation rituals through JavaScript incantations, choosing between:\n"
	"- **Registry-Free Communion**: MinHook's relic binds the magic without earthly registries\n"
	"- **Traditional Registry Rites**: Classic ceremonial bindings for ancient systems\n\n"
	"Summon the golden key with this sacred invocation:\n"
	"const apiWizard = new ActiveXObject('APIWizard'); \n"
	"apiWizard.PrintMessage(); \n"
	"console.log('apiWizard', apiWizard);"
);

// Define the GUID for your component
// {152C13C3-5F06-4314-B1F7-78ED35C62DAD}
DEFINE_GUID(APIWizard, 0xd4e8f8d6, 0x5b1f, 0x4f4d, 0x8a, 0x8f, 0x11, 0x6c, 0x3b, 0x8b, 0x1e, 0x7e);

// Implement the class_guid for component_installation_validator
// This will prevent users from renaming your component around (important for proper troubleshooter behaviors) or loading multiple instances of it.
VALIDATE_COMPONENT_FILENAME("foo_api_wizard.dll");

// Activate cfg_var downgrade functionality if enabled. Relevant only when cycling from newer FOOBAR2000_TARGET_VERSION to older.
FOOBAR2000_IMPLEMENT_CFG_VAR_DOWNGRADE;
#pragma endregion


/////////////////////////////
// * MAIN INITIALIZATION * //
/////////////////////////////
#pragma region Main Initialization
void APIWizard::InitAPIWizard() {
	mainHwnd = core_api::get_main_window();

	if (!mainHwnd) return;

	apiWizardMain = std::make_unique<APIWizardMain>(mainHwnd);
	// APIWizardMain init implementation
}

void APIWizard::QuitAPIWizard() {
	// APIWizardMain cleanup implementation
	apiWizardMain.reset();
	MyCOM::QuitMyCOM();
}

namespace {
	FB2K_ON_INIT_STAGE(MyCOM::InitMyCOM, init_stages::after_config_read);
	FB2K_ON_INIT_STAGE(APIWizard::InitAPIWizard, init_stages::after_ui_init);
	FB2K_RUN_ON_QUIT(APIWizard::QuitAPIWizard);
}
#pragma endregion
