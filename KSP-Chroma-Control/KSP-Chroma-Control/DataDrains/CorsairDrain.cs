// ReSharper disable CompareOfFloatsByEqualityOperator

namespace KspChromaControl.DataDrains
{
    using System;
    using System.Collections.Generic;
    using System.Runtime.InteropServices;
    using KspChromaControl.ColorSchemes;
    using UnityEngine;
    using Corsair.CUE.SDK;

    /// <summary>
    ///     Data drain that colors razer devices.
    /// </summary>
    internal class CorsairDrain : IDataDrain
    {
        // TODO - should get rid of the hasValidDevices flag and just throw an exception if we can't conenct to 
        // a keyboard at startup.
        public CorsairDrain()
        {
            LoadLibrary(@"GameData\KSPChromaControl\x64\CUESDK.dll");

            init();
            buildLedList();
        }

        private List<CorsairLedId> keyboardLeds = new List<CorsairLedId>();

        private Boolean hasValidDevices = false;
        private int keyboardIndex;

        /*
         * So the Colore Drain does this the other way, mapping from Unity KeyCodes to the
         * keyboard LED identifer.  This makes it possible for map two KeyCodes to the same LED
         * (for example, the number 2 and the "@" symbol on US keyboards).  This can cause 
         * flickering.  Doing the mapping the other way makes it impossible to do that.  This isn't
         * a problem because KSP only worries about the non-shifted KeyCode.
         * 
         * Doing the mapping this way also makes it easier to make sure that every writeable LED
         * is at least colored with the scheme base color, even if the LED doesn't map to a Unity
         * KeyCode (this happens with stuff like media control buttons and vendor-specific buttons).
         * 
         * This mapping is specific to US keyboards.  It should be possible to detect the keyboard
         * type and choose an appropriate mapping dynamically.
         */
        private static readonly Dictionary<CorsairLedId, KeyCode> keyMapping = new Dictionary<CorsairLedId, KeyCode>
        {
            {CorsairLedId.CLK_A, KeyCode.A},
            {CorsairLedId.CLK_0, KeyCode.Alpha0},
            {CorsairLedId.CLK_1, KeyCode.Alpha1},
            {CorsairLedId.CLK_2, KeyCode.Alpha2},
            {CorsairLedId.CLK_3, KeyCode.Alpha3},
            {CorsairLedId.CLK_4, KeyCode.Alpha4},
            {CorsairLedId.CLK_5, KeyCode.Alpha5},
            {CorsairLedId.CLK_6, KeyCode.Alpha6},
            {CorsairLedId.CLK_7, KeyCode.Alpha7},
            {CorsairLedId.CLK_8, KeyCode.Alpha8},
            {CorsairLedId.CLK_9, KeyCode.Alpha9},
            {CorsairLedId.CLK_B, KeyCode.B},
            {CorsairLedId.CLK_GraveAccentAndTilde, KeyCode.BackQuote},
            {CorsairLedId.CLK_Backslash, KeyCode.Backslash},
            {CorsairLedId.CLK_Backspace, KeyCode.Backspace},
            {CorsairLedId.CLK_C, KeyCode.C},
            {CorsairLedId.CLK_CapsLock, KeyCode.CapsLock},
            {CorsairLedId.CLK_CommaAndLessThan, KeyCode.Comma},
            {CorsairLedId.CLK_D, KeyCode.D},
            {CorsairLedId.CLK_Delete, KeyCode.Delete},
            {CorsairLedId.CLK_DownArrow, KeyCode.DownArrow},
            {CorsairLedId.CLK_E, KeyCode.E},
            {CorsairLedId.CLK_End, KeyCode.End},
            {CorsairLedId.CLK_EqualsAndPlus, KeyCode.Equals},
            {CorsairLedId.CLK_Escape, KeyCode.Escape},
            {CorsairLedId.CLK_F, KeyCode.F},
            {CorsairLedId.CLK_F1, KeyCode.F1},
            {CorsairLedId.CLK_F2, KeyCode.F2},
            {CorsairLedId.CLK_F3, KeyCode.F3},
            {CorsairLedId.CLK_F4, KeyCode.F4},
            {CorsairLedId.CLK_F5, KeyCode.F5},
            {CorsairLedId.CLK_F6, KeyCode.F6},
            {CorsairLedId.CLK_F7, KeyCode.F7},
            {CorsairLedId.CLK_F8, KeyCode.F8},
            {CorsairLedId.CLK_F9, KeyCode.F9},
            {CorsairLedId.CLK_F10, KeyCode.F10},
            {CorsairLedId.CLK_F11, KeyCode.F11},
            {CorsairLedId.CLK_F12, KeyCode.F12},
            {CorsairLedId.CLK_G, KeyCode.G},
            {CorsairLedId.CLK_H, KeyCode.H},
            {CorsairLedId.CLK_Home, KeyCode.Home},
            {CorsairLedId.CLK_I, KeyCode.I},
            {CorsairLedId.CLK_Insert, KeyCode.Insert},
            {CorsairLedId.CLK_J, KeyCode.J},
            {CorsairLedId.CLK_K, KeyCode.K},
            {CorsairLedId.CLK_Keypad0, KeyCode.Keypad0},
            {CorsairLedId.CLK_Keypad1, KeyCode.Keypad1},
            {CorsairLedId.CLK_Keypad2, KeyCode.Keypad2},
            {CorsairLedId.CLK_Keypad3, KeyCode.Keypad3},
            {CorsairLedId.CLK_Keypad4, KeyCode.Keypad4},
            {CorsairLedId.CLK_Keypad5, KeyCode.Keypad5},
            {CorsairLedId.CLK_Keypad6, KeyCode.Keypad6},
            {CorsairLedId.CLK_Keypad7, KeyCode.Keypad7},
            {CorsairLedId.CLK_Keypad8, KeyCode.Keypad8},
            {CorsairLedId.CLK_Keypad9, KeyCode.Keypad9},
            {CorsairLedId.CLK_KeypadSlash, KeyCode.KeypadDivide},
            {CorsairLedId.CLK_KeypadEnter, KeyCode.KeypadEnter},
            {CorsairLedId.CLK_KeypadMinus, KeyCode.KeypadMinus},
            {CorsairLedId.CLK_KeypadAsterisk, KeyCode.KeypadMultiply},
            {CorsairLedId.CLK_KeypadPeriodAndDelete, KeyCode.KeypadPeriod},
            {CorsairLedId.CLK_KeypadPlus, KeyCode.KeypadPlus},
            {CorsairLedId.CLK_L, KeyCode.L},
            {CorsairLedId.CLK_LeftAlt, KeyCode.LeftAlt},
            {CorsairLedId.CLK_LeftArrow, KeyCode.LeftArrow},
            {CorsairLedId.CLK_BracketLeft, KeyCode.LeftBracket},
            {CorsairLedId.CLK_LeftCtrl, KeyCode.LeftControl},
            {CorsairLedId.CLK_LeftShift, KeyCode.LeftShift},
            {CorsairLedId.CLK_LeftGui, KeyCode.LeftWindows},
            {CorsairLedId.CLK_M, KeyCode.M},
            {CorsairLedId.CLK_MR, KeyCode.Menu},
            {CorsairLedId.CLK_MinusAndUnderscore, KeyCode.Minus},
            {CorsairLedId.CLK_N, KeyCode.N},
            {CorsairLedId.CLK_NumLock, KeyCode.Numlock},
            {CorsairLedId.CLK_O, KeyCode.O},
            {CorsairLedId.CLK_P, KeyCode.P},
            {CorsairLedId.CLK_PageDown, KeyCode.PageDown},
            {CorsairLedId.CLK_PageUp, KeyCode.PageUp},
            {CorsairLedId.CLK_PauseBreak, KeyCode.Pause},
            {CorsairLedId.CLK_PeriodAndBiggerThan, KeyCode.Period},
            {CorsairLedId.CLK_PrintScreen, KeyCode.Print},
            {CorsairLedId.CLK_Q, KeyCode.Q},
            {CorsairLedId.CLK_ApostropheAndDoubleQuote, KeyCode.Quote},
            {CorsairLedId.CLK_R, KeyCode.R},
            {CorsairLedId.CLK_Enter, KeyCode.Return},
            {CorsairLedId.CLK_RightAlt, KeyCode.RightAlt},
            {CorsairLedId.CLK_RightArrow, KeyCode.RightArrow},
            {CorsairLedId.CLK_BracketRight, KeyCode.RightBracket},
            {CorsairLedId.CLK_RightCtrl, KeyCode.RightControl},
            {CorsairLedId.CLK_RightGui, KeyCode.RightWindows},
            {CorsairLedId.CLK_RightShift, KeyCode.RightShift},
            {CorsairLedId.CLK_S, KeyCode.S},
            {CorsairLedId.CLK_ScrollLock, KeyCode.ScrollLock},
            {CorsairLedId.CLK_SemicolonAndColon, KeyCode.Semicolon},
            {CorsairLedId.CLK_SlashAndQuestionMark, KeyCode.Slash},
            {CorsairLedId.CLK_Space, KeyCode.Space},
            {CorsairLedId.CLK_T, KeyCode.T},
            {CorsairLedId.CLK_Tab, KeyCode.Tab},
            {CorsairLedId.CLK_U, KeyCode.U},
            {CorsairLedId.CLK_UpArrow, KeyCode.UpArrow},
            {CorsairLedId.CLK_V, KeyCode.V},
            {CorsairLedId.CLK_W, KeyCode.W},
            {CorsairLedId.CLK_X, KeyCode.X},
            {CorsairLedId.CLK_Y, KeyCode.Y},
            {CorsairLedId.CLK_Z, KeyCode.Z}
        };

        // Connects to iCUE and searches for a valid device
        private void init()
        {
            CUESDK.CorsairPerformProtocolHandshake();

            var deviceCount = CUESDK.CorsairGetDeviceCount();
            UnityEngine.Debug.Log("[CHROMA] Device count " + deviceCount);

            if (deviceCount < 1)
            {
                hasValidDevices = false;
                return;
            }

            // The keyboard detection code makes an assumption that there is only one
            // keyboard plugged in.
            for (int deviceIndex = 0; deviceIndex < deviceCount; deviceIndex++)
            {
                var deviceInfo = CUESDK.CorsairGetDeviceInfo(deviceIndex);
                UnityEngine.Debug.Log("[CHROMA] Found device type " + deviceInfo.type);
                if (deviceInfo.type == CorsairDeviceType.CDT_Keyboard)
                {
                    keyboardIndex = deviceIndex;
                    hasValidDevices = true;
                    return;
                }
            }

            // Didn't find a keyboard.  Some keyboards show up as CDT_Unknown, so lets
            // fall back to looking for one of those
            for (int deviceIndex = 0; deviceIndex < deviceCount; deviceIndex++)
            {
                var deviceInfo = CUESDK.CorsairGetDeviceInfo(deviceIndex);
                UnityEngine.Debug.Log("[CHROMA] Found device type " + deviceInfo.type);
                if (deviceInfo.type == CorsairDeviceType.CDT_Unknown)
                {
                    UnityEngine.Debug.Log("[CHROMA] Trying an unknown device.");
                    keyboardIndex = deviceIndex;
                    hasValidDevices = true;
                    return;
                }
            }
        }

        // Get a list of all leds from the keyboard.  This helps ensure that we 
        // don't leave any unknown leds unset to the base color.
        private void buildLedList()
        {
            if (!hasValidDevices)
            {
                // No keyboard
                return;
            }

            var positions = CUESDK.CorsairGetLedPositionsByDeviceIndex(keyboardIndex);
            foreach (CorsairLedPosition position in positions.pLedPosition)
            {
                var ledId = position.ledId;
                keyboardLeds.Add(ledId);

                // Log the unmapped keys.  There will always be some, but we might
                // want to clean things up later.
                if (!keyMapping.ContainsKey(ledId))
                {
                    UnityEngine.Debug.Log("[CHROMA] Led \"" + ledId + "\" is not mapped.");
                }
            }

        }

        public bool Available()
        {
            return hasValidDevices;
        }

        public void Send(ColorScheme scheme)
        {
            /* There is opprotunity to increase effiency here.  The whole scheme is updated
             * every game frame and sent to the drain.  It should be possible to track which LEDs 
             * changed (if any) and only write those.  I suspect it would reduce the number of writes
             * by more than 90% during normal play.
             */

            if (hasValidDevices)
            {
                ApplyToKeyboard(scheme);
            }         
        }

        private void ApplyToKeyboard(ColorScheme colorScheme)
        {
        
            if (colorScheme == null)
            {
                return;
            }

            List<CorsairLedColor> ledColors = new List<CorsairLedColor>(keyboardLeds.Count);

            foreach (CorsairLedId led in keyboardLeds)
            {
                // If the key mapping isn't in the color scheme, default
                // to the scheme base color
                Color32 unityColor = colorScheme.BaseColor;

                if (keyMapping.ContainsKey(led))
                {
                    var unityKey = keyMapping[led];
                    unityColor = colorScheme[unityKey];
                }

                ledColors.Add(unityToCorsair(led, unityColor));
            }

            CUESDK.CorsairSetLedsColorsBufferByDeviceIndex(keyboardIndex, ledColors.Count, ledColors.ToArray());
            CUESDK.CorsairSetLedsColorsFlushBuffer();
        }

        private CorsairLedColor unityToCorsair(CorsairLedId ledId, Color32 unityColor)
        {
            CorsairLedColor color = new CorsairLedColor();
            color.ledId = ledId;
            color.r = unityColor.r;
            color.g = unityColor.g;
            color.b = unityColor.b;

            return color;
        }

        // Helper method to load the CUESDK native (C++) dll.  These will not be loaded
        // automatically by Unity/KSP, so we have to do it ourselves.
        [DllImport("kernel32", SetLastError = true, CharSet = CharSet.Ansi)]
        private static extern IntPtr LoadLibrary(
            [MarshalAs(UnmanagedType.LPStr)] string lpFileName);
    }
}