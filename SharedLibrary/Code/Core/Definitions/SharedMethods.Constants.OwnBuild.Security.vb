' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: SharedMethods.Constants.OwnBuild.Security.vb
' Purpose: Provides build-specific (fork/own-build) security-related constants for
'          the `SharedMethods` partial class.
'
' Architecture / How it works:
'  - This file is intentionally separated to simplify maintenance of private or
'    forked builds without touching the main constants definitions.
'  - It contains optional hard-coded values used by the library; when left empty,
'    the effective values are expected to be read elsewhere from the Windows
'    registry.
' =============================================================================

Option Strict On
Option Explicit On

Namespace SharedLibrary
    Partial Public Class SharedMethods

        ' Amend the following two values to hard code the encryption key and permitted domains (otherwise the values are taken from the registry at the path below)

        Private Const Int_CodeBasis As String = ""
        Public Const allowedDomains As String = ""
        Public Const noSilentIniUpdatesWithoutRegistryFlag As Boolean = False    ' If set to True, SharedMethods.UpdateIni.vb will disable silent INI unless explicitly enabled via registry flag

    End Class
End Namespace