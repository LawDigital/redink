Option Explicit On
Option Strict On
Option Infer On

Namespace Agents

    ''' <summary>
    ''' Parses the packed INI_PythonAgentPath value into a RedInkPythonAgentConfiguration.
    ''' Format: executablePath[;signer=Organization][;sha256=Hex]
    ''' Fields are order-independent and optional; only the executable path is required.
    ''' </summary>
    Public NotInheritable Class PythonExecuteToolConfig

        Private Sub New()
        End Sub

        Public Shared Function Parse(rawPythonAgentPath As System.String) As RedInkPythonAgentConfiguration
            If System.String.IsNullOrWhiteSpace(rawPythonAgentPath) Then
                Return Nothing
            End If

            Dim segments As System.String() = rawPythonAgentPath.Split(";"c)
            Dim executablePath As System.String = segments(0).Trim()
            If System.String.IsNullOrWhiteSpace(executablePath) Then
                Return Nothing
            End If

            Dim signer As System.String = Nothing
            Dim sha256 As System.String = Nothing

            For index As System.Int32 = 1 To segments.Length - 1
                Dim segment As System.String = segments(index).Trim()
                If segment.Length = 0 Then
                    Continue For
                End If

                Dim separatorIndex As System.Int32 = segment.IndexOf("="c)
                If separatorIndex <= 0 Then
                    Continue For
                End If

                Dim key As System.String = segment.Substring(0, separatorIndex).Trim().ToLowerInvariant()
                Dim value As System.String = segment.Substring(separatorIndex + 1).Trim()
                If value.Length = 0 Then
                    Continue For
                End If

                Select Case key
                    Case "signer", "signerorganization", "signer_organization"
                        signer = value
                    Case "sha256", "hash"
                        sha256 = value
                End Select
            Next

            Return New RedInkPythonAgentConfiguration() With {
                .ExecutablePath = executablePath,
                .ExpectedSignerOrganization = signer,
                .ExpectedSha256 = sha256
            }
        End Function

    End Class

End Namespace
