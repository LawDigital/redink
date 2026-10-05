' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved.

' =============================================================================
' File: SemanticArchiveIndexPolicy.vb
' Purpose:
'   Document-section indexing thresholds and effective semantic representation
'   signatures.
'
' Architecture / Function:
'   Separates policy-derived index signatures from stable document identity.
' =============================================================================

Option Strict On
Option Explicit On

Namespace SharedLibrary
    ''' <summary>Versioned document-index policy; sizes refer only to verified original source bytes.</summary>
    Public NotInheritable Class SemanticArchiveIndexPolicy
        Public Const SemanticPolicyVersion As System.String = "source-bytes-v1;document-card-reduction-v1"

        Private Sub New()
        End Sub

        Public Shared Function RequiresDocumentSectionIndex(originalByteLength As System.Int64, thresholdBytes As System.Int64) As System.Boolean
            If originalByteLength < 0 Then Throw New System.ArgumentOutOfRangeException(NameOf(originalByteLength))
            If thresholdBytes < 0 Then Throw New System.ArgumentOutOfRangeException(NameOf(thresholdBytes))
            Return thresholdBytes > 0 AndAlso originalByteLength >= thresholdBytes
        End Function

        Public Shared Function CreateSemanticSignature(modelSignature As System.String, generatorVersion As System.String,
                                                       semanticProfileVersion As System.String, thresholdBytes As System.Int64,
                                                       allowPartialSearch As System.Boolean) As System.String
            If thresholdBytes < 0 Then Throw New System.ArgumentOutOfRangeException(NameOf(thresholdBytes))
            Dim policy As New Newtonsoft.Json.Linq.JObject From {
                {"policy_version", SemanticPolicyVersion}, {"threshold_byte_domain", "original_source"},
                {"model", modelSignature}, {"generator", generatorVersion}, {"profile", semanticProfileVersion},
                {"source_index_threshold_bytes", thresholdBytes}, {"allow_partial_search", allowPartialSearch}
            }
            Return SemanticArchiveIdentity.StableId("semantic", policy.ToString(Newtonsoft.Json.Formatting.None))
        End Function
    End Class
End Namespace
