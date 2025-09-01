/////////////////////////////////////////////////////////////////////////////////
// * FB2K Component: COM Automation and ActiveX Interface                    * //
// * Description:    API Wizard Precompiled Header File                      * //
// * Author:         TT                                                      * //
// * Website:        https://github.com/The-Wizardium/API-Wizard             * //
// * Version:        0.1.0                                                   * //
// * Dev. started:   12-12-2024                                              * //
// * Last change:    01-09-2025                                              * //
/////////////////////////////////////////////////////////////////////////////////


#define NOMINMAX
#define STRICT
#define WIN32_LEAN_AND_MEAN


// * System headers * //
#include <Windows.h> // WIN32_LEAN_AND_MEAN must precede Windows.h
#include <mmsystem.h> // When WIN32_LEAN_AND_MEAN macro is used, mmsystem.h is needed for timeGetTime in pfc/timers.h

// * Standard C++ libraries * //
#include <map>
#include <memory>
#include <string>
#include <string_view>

// * External libraries * //
#include <MinHook.h>

// * foobar2000 SDK * //
#ifdef __cplusplus
#include <helpers/foobar2000+atl.h>
#endif

// * macOS-specific includes * //
#ifdef __OBJC__
#include <Cocoa/Cocoa.h>
#endif
