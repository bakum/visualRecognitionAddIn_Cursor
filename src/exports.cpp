/*
 *  Modern Native AddIn
 *  Copyright (C) 2018  Infactum
 *
 *  This program is free software: you can redistribute it and/or modify
 *  it under the terms of the GNU Affero General Public License as
 *  published by the Free Software Foundation, either version 3 of the
 *  License, or (at your option) any later version.
 *
 *  This program is distributed in the hope that it will be useful,
 *  but WITHOUT ANY WARRANTY; without even the implied warranty of
 *  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 *  GNU Affero General Public License for more details.
 *
 *  You should have received a copy of the GNU Affero General Public License
 *  along with this program.  If not, see <https://www.gnu.org/licenses/>.
 *
 */

#include <string>

#include <ComponentBase.h>
#include <types.h>

#include "VisualAddIn.h"

#ifdef _WINDOWS
#pragma warning (disable : 4311 4302)
#endif

#if defined(__GNUC__) && !defined(_WIN32)
#define BARCODEADDIN_C_EXPORT __attribute__((visibility("default")))
#else
#define BARCODEADDIN_C_EXPORT
#endif

extern "C" {

BARCODEADDIN_C_EXPORT const WCHAR_T *GetClassNames() {
    // Might contain multiple class names seperated by |
    static char16_t cls_names[] = u"VisualRecognition";
    return reinterpret_cast<WCHAR_T *>(cls_names);
}

BARCODEADDIN_C_EXPORT long GetClassObject(const WCHAR_T *clsName, IComponentBase **pInterface) {
    if (!*pInterface) {
        auto cls_name = std::u16string(reinterpret_cast<const char16_t *>(clsName));
        if (cls_name == u"VisualRecognition") {
            *pInterface = new VisualAddIn();
        }
        return (long) *pInterface;
    }
    return 0;
}

BARCODEADDIN_C_EXPORT long DestroyObject(IComponentBase **pInterface) {
    if (!*pInterface) {
        return -1;
    }

    delete *pInterface;
    *pInterface = nullptr;
    return 0;
}

BARCODEADDIN_C_EXPORT AppCapabilities SetPlatformCapabilities(const AppCapabilities capabilities) {
    (void)capabilities;
    return eAppCapabilitiesLast;
}

} // extern "C"
