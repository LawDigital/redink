' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
' Semantic Archive control-plane, immutable generation and provenance contracts.

' =============================================================================
' File: SemanticArchive.Schema.vb
' Purpose:
'   Catalog, archive, binding, generation, card and semantic-routing data contracts.
'
' Architecture / Function:
'   Provides shared serialization models and configuration conversion; source identity
'   and processing/storage policy remain separate.
' =============================================================================

Option Strict On
Option Explicit On

Namespace SharedLibrary

    <Newtonsoft.Json.JsonConverterAttribute(GetType(SemanticArchiveConfigurationJsonConverter))>
    Public NotInheritable Class SemanticArchiveCatalog
        Public Property SchemaVersion As System.Int32 = 1
        Public Property Revision As System.Int64
        Public Property Archives As New System.Collections.Generic.List(Of SemanticArchiveDefinition)()
        <Newtonsoft.Json.JsonPropertyAttribute("SemanticArchiveDefaultArchiveIds")>
        Public Property DefaultArchiveIds As New System.Collections.Generic.List(Of System.String)()

        ' Preserve unrecognized configuration during conservative catalog round trips.
        <Newtonsoft.Json.JsonExtensionDataAttribute(), System.ComponentModel.BrowsableAttribute(False)>
        Public Property AdditionalConfiguration As New System.Collections.Generic.Dictionary(Of System.String, Newtonsoft.Json.Linq.JToken)(System.StringComparer.Ordinal)
    End Class

    <Newtonsoft.Json.JsonConverterAttribute(GetType(SemanticArchiveConfigurationJsonConverter))>
    Public NotInheritable Class SemanticArchiveDefinition
        <Newtonsoft.Json.JsonPropertyAttribute("SemanticArchiveLibrary"), System.ComponentModel.BrowsableAttribute(False)>
        Public Property Library As SemanticArchiveLibraryRegistration
        Public Property ArchiveId As System.String = ""
        <Newtonsoft.Json.JsonPropertyAttribute("SemanticArchiveName")>
        Public Property Name As System.String = SharedMethods.DEFAULT_SEMANTICARCHIVE_NAME
        <Newtonsoft.Json.JsonPropertyAttribute("SemanticArchiveDescription")>
        Public Property Description As System.String = SharedMethods.DEFAULT_SEMANTICARCHIVE_DESCRIPTION
        <Newtonsoft.Json.JsonPropertyAttribute("SemanticArchiveEnabled")>
        Public Property Enabled As System.Boolean = SharedMethods.DEFAULT_SEMANTICARCHIVE_ENABLED
        <Newtonsoft.Json.JsonPropertyAttribute("SemanticArchiveRoots")>
        Public Property Roots As New System.Collections.Generic.List(Of SemanticArchiveSourceBinding)()
        <Newtonsoft.Json.JsonPropertyAttribute("SemanticArchiveBackgroundEnabled")>
        Public Property BackgroundEnabled As System.Boolean = SharedMethods.DEFAULT_SEMANTICARCHIVE_BACKGROUND_ENABLED
        <Newtonsoft.Json.JsonPropertyAttribute("SemanticArchiveBackgroundWindow")>
        Public Property BackgroundWindow As System.String = SharedMethods.DEFAULT_SEMANTICARCHIVE_BACKGROUND_WINDOW
        ' Threshold uses original source-file bytes; zero disables source index generation.
        <Newtonsoft.Json.JsonPropertyAttribute("SemanticArchiveSourceIndexThresholdBytes")>
        Public Property SectionIndexThresholdBytes As System.Int64 = SharedMethods.DEFAULT_SEMANTICARCHIVE_SOURCE_INDEX_THRESHOLD_BYTES
        <Newtonsoft.Json.JsonPropertyAttribute("SemanticArchiveMaxChildrenPerNode")>
        Public Property MaxChildrenPerNode As System.Int32 = SharedMethods.DEFAULT_SEMANTICARCHIVE_MAX_CHILDREN_PER_NODE
        <Newtonsoft.Json.JsonPropertyAttribute("SemanticArchiveMaxRoutingCharacters")>
        Public Property MaxRoutingCharacters As System.Int32 = SharedMethods.DEFAULT_SEMANTICARCHIVE_MAX_ROUTING_CHARACTERS
        <Newtonsoft.Json.JsonPropertyAttribute("SemanticArchiveExtractionProfileVersion")>
        Public Property ExtractionProfileVersion As System.String = SharedMethods.DEFAULT_SEMANTICARCHIVE_EXTRACTION_PROFILE_VERSION
        <Newtonsoft.Json.JsonPropertyAttribute("SemanticArchiveSemanticProfileVersion")>
        Public Property SemanticProfileVersion As System.String = SharedMethods.DEFAULT_SEMANTICARCHIVE_SEMANTIC_PROFILE_VERSION
        <Newtonsoft.Json.JsonPropertyAttribute("SemanticArchiveAllowPartialSearch")>
        Public Property AllowPartialSearch As System.Boolean = SharedMethods.DEFAULT_SEMANTICARCHIVE_ALLOW_PARTIAL_SEARCH
        <Newtonsoft.Json.JsonPropertyAttribute("SemanticArchiveRetrievalBudgets")>
        Public Property RetrievalBudgets As New SemanticArchiveRetrievalBudgets()

        ' Preserve unrecognized configuration during conservative catalog round trips.
        <Newtonsoft.Json.JsonExtensionDataAttribute(), System.ComponentModel.BrowsableAttribute(False)>
        Public Property AdditionalConfiguration As New System.Collections.Generic.Dictionary(Of System.String, Newtonsoft.Json.Linq.JToken)(System.StringComparer.Ordinal)
    End Class

    <Newtonsoft.Json.JsonConverterAttribute(GetType(SemanticArchiveConfigurationJsonConverter))>
    Public NotInheritable Class SemanticArchiveSourceBinding
        Public Property BindingId As System.String = ""
        <Newtonsoft.Json.JsonPropertyAttribute("SemanticArchiveRootPath")>
        Public Property RootPath As System.String = SharedMethods.DEFAULT_SEMANTICARCHIVE_ROOT_PATH
        <Newtonsoft.Json.JsonPropertyAttribute("SemanticArchiveRecursive")>
        Public Property Recursive As System.Boolean = SharedMethods.DEFAULT_SEMANTICARCHIVE_RECURSIVE
        <Newtonsoft.Json.JsonPropertyAttribute("SemanticArchiveExclusions")>
        Public Property Exclusions As New System.Collections.Generic.List(Of System.String)()
        <Newtonsoft.Json.JsonPropertyAttribute("SemanticArchiveFileTypeFilterEnabled")>
        Public Property FileTypeFilterEnabled As System.Boolean = SharedMethods.DEFAULT_SEMANTICARCHIVE_FILE_TYPE_FILTER_ENABLED
        <Newtonsoft.Json.JsonPropertyAttribute("SemanticArchiveSupportedExtensions", ObjectCreationHandling:=Newtonsoft.Json.ObjectCreationHandling.Replace)>
        Public Property SupportedExtensions As New System.Collections.Generic.List(Of System.String)(SharedMethods.DEFAULT_SEMANTICARCHIVE_SUPPORTED_EXTENSIONS.Split(";"c))
        ' auto shares per-source artifacts only when source protection can be projected.
        <Newtonsoft.Json.JsonPropertyAttribute("SemanticArchiveArtifactPlacementMode")>
        Public Property ArtifactPlacementMode As System.String = SharedMethods.DEFAULT_SEMANTICARCHIVE_ARTIFACT_PLACEMENT_MODE
        <Newtonsoft.Json.JsonPropertyAttribute("SemanticArchiveSharedArtifactRoot")>
        Public Property SharedArtifactRoot As System.String = SharedMethods.DEFAULT_SEMANTICARCHIVE_SHARED_ARTIFACT_ROOT
        <Newtonsoft.Json.JsonPropertyAttribute("SemanticArchiveShadowArtifactRoot")>
        Public Property ShadowArtifactRoot As System.String = SharedMethods.DEFAULT_SEMANTICARCHIVE_SHADOW_ARTIFACT_ROOT
        <Newtonsoft.Json.JsonPropertyAttribute("SemanticArchiveEnableOcr")>
        Public Property EnableOcr As System.Boolean = SharedMethods.DEFAULT_SEMANTICARCHIVE_ENABLE_OCR
        <Newtonsoft.Json.JsonPropertyAttribute("SemanticArchiveOcrBatchPages")>
        Public Property OcrBatchPages As System.Int32 = SharedMethods.DEFAULT_SEMANTICARCHIVE_OCR_BATCH_PAGES
        <Newtonsoft.Json.JsonPropertyAttribute("SemanticArchiveExtractionOptionsSignature")>
        Public Property ExtractionOptionsSignature As System.String = SharedMethods.DEFAULT_SEMANTICARCHIVE_EXTRACTION_OPTIONS_SIGNATURE
        <Newtonsoft.Json.JsonPropertyAttribute("SemanticArchiveScopeTags")>
        Public Property ScopeTags As New System.Collections.Generic.List(Of System.String)()

        ' Preserve unrecognized configuration during conservative catalog round trips.
        <Newtonsoft.Json.JsonExtensionDataAttribute(), System.ComponentModel.BrowsableAttribute(False)>
        Public Property AdditionalConfiguration As New System.Collections.Generic.Dictionary(Of System.String, Newtonsoft.Json.Linq.JToken)(System.StringComparer.Ordinal)
    End Class

    <Newtonsoft.Json.JsonConverterAttribute(GetType(SemanticArchiveConfigurationJsonConverter))>
    Public NotInheritable Class SemanticArchiveRetrievalBudgets
        <Newtonsoft.Json.JsonPropertyAttribute("SemanticArchiveMaxNodesVisited")>
        <System.ComponentModel.DisplayNameAttribute("Navigation nodes per search")>
        <System.ComponentModel.DescriptionAttribute("Maximum archive navigation nodes inspected in one search call. Higher values can improve coverage but require more work. Must be positive; the host caps this at 256. The lowest limit of all selected archives applies. Parameter: SemanticArchiveMaxNodesVisited.")>
        Public Property MaxNodesVisited As System.Int32 = SharedMethods.DEFAULT_SEMANTICARCHIVE_MAX_NODES_VISITED
        <Newtonsoft.Json.JsonPropertyAttribute("SemanticArchiveMaxModelCalls")>
        <System.ComponentModel.DisplayNameAttribute("Model calls per operation")>
        <System.ComponentModel.DescriptionAttribute("Maximum model selection calls in one search or evidence-read operation. Must be positive; the host caps this at 64. Lower limits can leave work for a continuation. The lowest limit of all selected archives applies. Parameter: SemanticArchiveMaxModelCalls.")>
        Public Property MaxModelCalls As System.Int32 = SharedMethods.DEFAULT_SEMANTICARCHIVE_MAX_MODEL_CALLS
        <Newtonsoft.Json.JsonPropertyAttribute("SemanticArchiveMaxElapsedSeconds")>
        <System.ComponentModel.DisplayNameAttribute("Time limit (seconds)")>
        <System.ComponentModel.DescriptionAttribute("Time budget for one search or evidence-read operation, in seconds. Must be positive; the host caps this at 180. Work stops at cancellation/checkpoint boundaries, so this is not a hard wall-clock guarantee. The lowest selected-archive limit applies. Parameter: SemanticArchiveMaxElapsedSeconds.")>
        Public Property MaxElapsedSeconds As System.Int32 = SharedMethods.DEFAULT_SEMANTICARCHIVE_MAX_ELAPSED_SECONDS
        <Newtonsoft.Json.JsonPropertyAttribute("SemanticArchiveMaxCandidateFiles")>
        <System.ComponentModel.DisplayNameAttribute("Candidate file limit")>
        <System.ComponentModel.DescriptionAttribute("Maximum candidate files retained for selection during search. This is not the number of files in the archive or the final result count. Must be positive; the host caps this at 512. Semantic, metadata and explicit literal searches share this capacity. Parameter: SemanticArchiveMaxCandidateFiles.")>
        Public Property MaxCandidateFiles As System.Int32 = SharedMethods.DEFAULT_SEMANTICARCHIVE_MAX_CANDIDATE_FILES
        <Newtonsoft.Json.JsonPropertyAttribute("SemanticArchiveMaxEvidenceBytes")>
        <System.ComponentModel.DisplayNameAttribute("Evidence text per read (bytes)")>
        <System.ComponentModel.DescriptionAttribute("Maximum extracted UTF-8 text bytes returned by one evidence-read call. Use at least 128; smaller positive values prevent evidence reads. The host caps this at 262,144 bytes (256 KiB). The lowest selected-archive limit and the remaining run budget also apply. Parameter: SemanticArchiveMaxEvidenceBytes.")>
        Public Property MaxEvidenceBytes As System.Int32 = SharedMethods.DEFAULT_SEMANTICARCHIVE_MAX_EVIDENCE_BYTES
        <Newtonsoft.Json.JsonPropertyAttribute("SemanticArchiveInitialBranches")>
        <System.ComponentModel.DisplayNameAttribute("Branches selected at a time")>
        <System.ComponentModel.DescriptionAttribute("Maximum navigation branches initially selected from one metadata batch. Must be positive; the host caps this at 16. Higher values broaden the search. Cards per selection batch must be at least this value and at least 8; remaining branches can be explored by continuing the search. Parameter: SemanticArchiveInitialBranches.")>
        Public Property InitialBranches As System.Int32 = SharedMethods.DEFAULT_SEMANTICARCHIVE_INITIAL_BRANCHES
        <Newtonsoft.Json.JsonPropertyAttribute("SemanticArchiveMaxExactLookupDocuments")>
        <System.ComponentModel.DisplayNameAttribute("Metadata records per search")>
        <System.ComponentModel.DescriptionAttribute("Maximum document metadata records inspected by the independent exact/lexical metadata channel in one search call. 0 disables that channel; semantic search remains available. This does not scan extracted document text. The host caps this at 10,000 records. Parameter: SemanticArchiveMaxExactLookupDocuments.")>
        Public Property MaxExactLookupDocuments As System.Int32 = SharedMethods.DEFAULT_SEMANTICARCHIVE_MAX_EXACT_LOOKUP_DOCUMENTS
        <Newtonsoft.Json.JsonPropertyAttribute("SemanticArchiveMaxSectionCandidates")>
        <System.ComponentModel.DisplayNameAttribute("Cards per selection batch")>
        <System.ComponentModel.DescriptionAttribute("Maximum navigation, document or section metadata cards passed to one model selection batch. Use at least 8 and at least Branches selected at a time. The host caps this at 64. Character and token limits can make batches smaller; this is not a limit on all sections in a document. Parameter: SemanticArchiveMaxSectionCandidates.")>
        Public Property MaxSectionCandidates As System.Int32 = SharedMethods.DEFAULT_SEMANTICARCHIVE_MAX_SECTION_CANDIDATES
        <Newtonsoft.Json.JsonPropertyAttribute("SemanticArchiveMaxPromptCharacters")>
        <System.ComponentModel.DisplayNameAttribute("Selection metadata limit (characters)")>
        <System.ComponentModel.DescriptionAttribute("Maximum compact metadata characters in one model selection batch; excludes the query and prompt framing. Must be at least 1,024; the host caps this at 64,000. Complete cards must fit without truncation. The full model request is also checked against the token limit. Parameter: SemanticArchiveMaxPromptCharacters.")>
        Public Property MaxPromptCharacters As System.Int32 = SharedMethods.DEFAULT_SEMANTICARCHIVE_MAX_PROMPT_CHARACTERS
        <Newtonsoft.Json.JsonPropertyAttribute("SemanticArchiveMaxRequestTokens")>
        <System.ComponentModel.DisplayNameAttribute("Model request limit (tokens)")>
        <System.ComponentModel.DescriptionAttribute("Token budget for a complete model selection request, including its instructions, query, metadata and reserved response space. Valid range: 1 to 262,144. Allow more than the 4,096-token response reservation plus the input; very small values prevent requests. Model-specific and selected-archive limits also apply. Parameter: SemanticArchiveMaxRequestTokens.")>
        Public Property MaxRequestTokens As System.Int32 = SharedMethods.DEFAULT_SEMANTICARCHIVE_MAX_REQUEST_TOKENS
        <Newtonsoft.Json.JsonPropertyAttribute("SemanticArchiveMaxLiteralScanBytes")>
        <System.ComponentModel.DisplayNameAttribute("Literal-search work limit (bytes)")>
        <System.ComponentModel.DescriptionAttribute("Byte budget for an explicitly requested case-sensitive search of extracted text, per call. Includes validation work: original source bytes plus twice the extracted UTF-8 text bytes. Range: 1 to 67,108,864 (64 MiB). A document whose full validation exceeds this budget is skipped and reported; this does not enable a text scan by itself. Parameter: SemanticArchiveMaxLiteralScanBytes.")>
        Public Property MaxLiteralScanBytes As System.Int32 = SharedMethods.DEFAULT_SEMANTICARCHIVE_MAX_LITERAL_SCAN_BYTES
        <Newtonsoft.Json.JsonPropertyAttribute("SemanticArchiveMaxLiteralScanDocuments")>
        <System.ComponentModel.DisplayNameAttribute("Literal-search documents per call")>
        <System.ComponentModel.DescriptionAttribute("Maximum document records inspected in one explicitly requested literal-text search call. Range: 1 to 1,024. Authorization checks, the byte budget, available candidate slots and the time budget can stop the scan earlier. This does not limit ordinary semantic search. Parameter: SemanticArchiveMaxLiteralScanDocuments.")>
        Public Property MaxLiteralScanDocuments As System.Int32 = SharedMethods.DEFAULT_SEMANTICARCHIVE_MAX_LITERAL_SCAN_DOCUMENTS

        ' Preserve unrecognized configuration during conservative catalog round trips.
        <Newtonsoft.Json.JsonExtensionDataAttribute(), System.ComponentModel.BrowsableAttribute(False)>
        Public Property AdditionalConfiguration As New System.Collections.Generic.Dictionary(Of System.String, Newtonsoft.Json.Linq.JToken)(System.StringComparer.Ordinal)
    End Class

    ''' <summary>Read-only migration for the four mutable configuration contracts. Canonical
    ''' names win over their old CLR-name aliases regardless of JSON property order.
    ''' Immutable generation/provenance objects deliberately do not use this converter.</summary>
    Public NotInheritable Class SemanticArchiveConfigurationJsonConverter
        Inherits Newtonsoft.Json.JsonConverter

        Public Overrides ReadOnly Property CanWrite As System.Boolean
            Get
                Return False
            End Get
        End Property

        Public Overrides Function CanConvert(objectType As System.Type) As System.Boolean
            Return objectType Is GetType(SemanticArchiveCatalog) OrElse objectType Is GetType(SemanticArchiveDefinition) OrElse
                objectType Is GetType(SemanticArchiveSourceBinding) OrElse objectType Is GetType(SemanticArchiveRetrievalBudgets)
        End Function

        Public Overrides Function ReadJson(reader As Newtonsoft.Json.JsonReader, objectType As System.Type,
                                           existingValue As System.Object, serializer As Newtonsoft.Json.JsonSerializer) As System.Object
            If Not CanConvert(objectType) Then Throw New Newtonsoft.Json.JsonSerializationException("Unsupported Semantic Archive configuration contract.")
            If reader.TokenType = Newtonsoft.Json.JsonToken.Null Then Return Nothing
            If reader.TokenType <> Newtonsoft.Json.JsonToken.StartObject Then Throw New Newtonsoft.Json.JsonSerializationException("Semantic Archive configuration must be a JSON object.")
            Dim configuration As Newtonsoft.Json.Linq.JObject = Newtonsoft.Json.Linq.JObject.Load(reader,
                New Newtonsoft.Json.Linq.JsonLoadSettings() With {.DuplicatePropertyNameHandling = Newtonsoft.Json.Linq.DuplicatePropertyNameHandling.Error})
            For Each member As System.Reflection.PropertyInfo In objectType.GetProperties(System.Reflection.BindingFlags.Public Or System.Reflection.BindingFlags.Instance)
                Dim attribute As Newtonsoft.Json.JsonPropertyAttribute = TryCast(System.Attribute.GetCustomAttribute(member, GetType(Newtonsoft.Json.JsonPropertyAttribute)), Newtonsoft.Json.JsonPropertyAttribute)
                If attribute Is Nothing OrElse System.String.IsNullOrWhiteSpace(attribute.PropertyName) Then
                    ' Structural names and stable IDs do not migrate, but ambiguous
                    ' case variants must not silently choose a different identity.
                    If Not System.Attribute.IsDefined(member, GetType(Newtonsoft.Json.JsonExtensionDataAttribute)) Then FindUnambiguousProperty(configuration, member.Name)
                    Continue For
                End If
                Dim canonical As Newtonsoft.Json.Linq.JProperty = FindUnambiguousProperty(configuration, attribute.PropertyName)
                Dim legacy As Newtonsoft.Json.Linq.JProperty = FindUnambiguousProperty(configuration, member.Name)
                Dim selected As Newtonsoft.Json.Linq.JProperty = If(canonical, legacy)
                If selected Is Nothing Then Continue For
                Dim value As Newtonsoft.Json.Linq.JToken = selected.Value.DeepClone()
                If canonical IsNot Nothing Then canonical.Remove()
                If legacy IsNot Nothing Then legacy.Remove()
                configuration.Add(attribute.PropertyName, value)
            Next
            ' Populate avoids re-entering this type's converter while nested configuration
            ' contracts still use their own migration. A fresh object preserves defaults
            ' and avoids appending source roots/budgets to an existing deserialization value.
            Dim result As System.Object = System.Activator.CreateInstance(objectType)
            Using normalized As Newtonsoft.Json.JsonReader = configuration.CreateReader()
                serializer.Populate(normalized, result)
            End Using
            Return result
        End Function

        Private Shared Function FindUnambiguousProperty(configuration As Newtonsoft.Json.Linq.JObject,
                                                        name As System.String) As Newtonsoft.Json.Linq.JProperty
            Dim found As Newtonsoft.Json.Linq.JProperty = Nothing
            For Each item As Newtonsoft.Json.Linq.JProperty In configuration.Properties()
                If Not System.String.Equals(item.Name, name, System.StringComparison.OrdinalIgnoreCase) Then Continue For
                If found IsNot Nothing Then Throw New Newtonsoft.Json.JsonSerializationException("Ambiguous Semantic Archive configuration property: " & name & ".")
                found = item
            Next
            Return found
        End Function

        Public Overrides Sub WriteJson(writer As Newtonsoft.Json.JsonWriter, value As System.Object, serializer As Newtonsoft.Json.JsonSerializer)
            Throw New System.NotSupportedException("Canonical Semantic Archive configuration uses the normal JsonProperty writer.")
        End Sub
    End Class

    Public NotInheritable Class SemanticArchiveSourceFingerprint
        Public Property Length As System.Int64
        Public Property LastWriteUtcTicks As System.Int64
        Public Property Sha256 As System.String = ""
    End Class

    Public NotInheritable Class SemanticArchiveRepresentation
        Public Property RepresentationId As System.String = ""
        Public Property SourceHash As System.String = ""
        Public Property ExtractorVersion As System.String = ""
        Public Property OptionsSignature As System.String = ""
        Public Property TextPath As System.String = ""
        Public Property TextFileHash As System.String = ""
        Public Property TextByteLength As System.Int64
        Public Property EncodingName As System.String = "utf-8"
        ' complete, empty, incomplete, unknown. Unknown is deliberately not complete.
        Public Property Completeness As System.String = "unknown"
        Public Property SourceMapJson As System.String = ""
        Public Property ExtractedUtc As System.DateTimeOffset = System.DateTimeOffset.UtcNow
    End Class

    Public NotInheritable Class SemanticArchiveIndexDescriptor
        Public Property IndexId As System.String = ""
        Public Property RepresentationId As System.String = ""
        Public Property Path As System.String = ""
        Public Property FileHash As System.String = ""
        Public Property PayloadHash As System.String = ""
        Public Property FormatVersion As System.Int32 = 1
        Public Property GeneratorVersion As System.String = ""
        Public Property ProfileVersion As System.String = ""
        Public Property ModelIdentity As System.String = ""
        Public Property EntryCount As System.Int32
    End Class

    Public NotInheritable Class SemanticArchiveDocumentRecord
        Public Property DocumentId As System.String = ""
        Public Property SourceItemId As System.String = ""
        Public Property CanonicalSourceKey As System.String = ""
        Public Property PartitionKey As System.String = ""
        Public Property SourcePath As System.String = ""
        Public Property RelativePath As System.String = ""
        Public Property DisplayName As System.String = ""
        Public Property BindingIds As New System.Collections.Generic.List(Of System.String)()
        Public Property Fingerprint As New SemanticArchiveSourceFingerprint()
        Public Property Active As System.Boolean = True
        Public Property ProcessingStatus As System.String = "pending"
        Public Property Diagnostic As System.String = ""
        Public Property SemanticSignature As System.String = ""
        Public Property SemanticModelIdentity As System.String = ""
        Public Property ExtractionSignature As System.String = ""
        Public Property CooperativeState As System.String = ""
        Public Property CooperativeDiagnostic As System.String = ""
        Public Property Representation As SemanticArchiveRepresentation
        Public Property Index As SemanticArchiveIndexDescriptor
        Public Property Card As SemanticArchiveCard
    End Class

    Public NotInheritable Class SemanticArchiveCard
        Public Property CardId As System.String = ""
        ' CONTAINER, DOCUMENT, SECTION. These values carry no provider semantics.
        Public Property Level As System.String = "DOCUMENT"
        Public Property TargetId As System.String = ""
        Public Property PartitionKey As System.String = ""
        Public Property SourceVersion As System.String = ""
        Public Property RepresentationId As System.String = ""
        Public Property Title As System.String = ""
        Public Property Summary As System.String = ""
        Public Property Topics As New System.Collections.Generic.List(Of System.String)()
        Public Property UserIntents As New System.Collections.Generic.List(Of System.String)()
        Public Property Identifiers As New System.Collections.Generic.List(Of System.String)()
        Public Property ExactTerms As New System.Collections.Generic.List(Of System.String)()
        Public Property Metadata As SharedMethods.SemanticSearchIndexEntry
        Public Property RetrievalText As System.String = ""
        Public Property RoutingReduced As System.Boolean
        Public Property FullMetadataReference As System.String = ""
        Public Property StartByte As System.Int64
        Public Property LengthBytes As System.Int64
        Public Property IsPermissionNeutral As System.Boolean
    End Class

    Public NotInheritable Class SemanticArchiveArtifactReference
        ' Relative to the configured archive directory, never a model-generated path.
        Public Property RelativePath As System.String = ""
        Public Property Sha256 As System.String = ""
        Public Property Length As System.Int64
    End Class

    Public NotInheritable Class SemanticArchiveNode
        Public Property NodeId As System.String = ""
        Public Property GenerationId As System.String = ""
        Public Property ParentNodeId As System.String = ""
        Public Property IsLeaf As System.Boolean
        Public Property Level As System.Int32
        Public Property MinKey As System.String = ""
        Public Property MaxKey As System.String = ""
        Public Property IndexPath As System.String = ""
        Public Property Cards As New System.Collections.Generic.List(Of SemanticArchiveCard)()
        Public Property IsPermissionNeutral As System.Boolean
    End Class

    Public NotInheritable Class SemanticArchiveNodeDescriptor
        Public Property NodeId As System.String = ""
        Public Property ArtifactPath As System.String = ""
        Public Property GenerationId As System.String = ""
        Public Property ParentNodeId As System.String = ""
        Public Property MinKey As System.String = ""
        Public Property MaxKey As System.String = ""
        Public Property IsLeaf As System.Boolean
        Public Property ChildCount As System.Int32
        Public Property ChildNodeIds As New System.Collections.Generic.List(Of System.String)()
        Public Property Sha256 As System.String = ""
        Public Property Artifact As SemanticArchiveArtifactReference
        Public Property IndexArtifact As SemanticArchiveArtifactReference
    End Class

    Public NotInheritable Class SemanticArchiveDocumentShard
        Public Property ShardId As System.String = ""
        Public Property MinKey As System.String = ""
        Public Property MaxKey As System.String = ""
        Public Property Documents As New System.Collections.Generic.List(Of SemanticArchiveDocumentRecord)()
    End Class

    Public NotInheritable Class SemanticArchiveDocumentShardDescriptor
        Public Property Inventory As New SemanticArchiveInventory()
        Public Property ShardId As System.String = ""
        Public Property ArtifactPath As System.String = ""
        Public Property GenerationId As System.String = ""
        Public Property MinKey As System.String = ""
        Public Property MaxKey As System.String = ""
        Public Property DocumentIds As New System.Collections.Generic.List(Of System.String)()
        Public Property RecordCount As System.Int32
        Public Property Sha256 As System.String = ""
        Public Property Artifact As SemanticArchiveArtifactReference
    End Class


    Public NotInheritable Class SemanticArchiveRoutingGroup
        Public Property GroupId As System.String = ""
        Public Property RouteKind As System.String = ""
        Public Property Level As System.Int32
        Public Property Card As SemanticArchiveCard
        Public Property ChildGroupIds As New System.Collections.Generic.List(Of System.String)()
        Public Property DocumentIds As New System.Collections.Generic.List(Of System.String)()
        Public Property RepresentativeDocumentIds As New System.Collections.Generic.List(Of System.String)()
        Public Property MembershipSignature As System.String = ""
        Public Property ContentSignature As System.String = ""
    End Class

    Public NotInheritable Class SemanticArchiveRoutingGraph
        ' Zero is invalid: a missing persisted version must not become current by default.
        Public Property SchemaVersion As System.Int32
        Public Property ProfileSignature As System.String = ""
        Public Property MaxChildrenPerGroup As System.Int32 = 48
        Public Property RootGroupIds As New System.Collections.Generic.List(Of System.String)()
        Public Property Groups As New System.Collections.Generic.Dictionary(Of System.String, SemanticArchiveRoutingGroup)(System.StringComparer.Ordinal)
        Public Property DocumentParentGroupIds As New System.Collections.Generic.Dictionary(Of System.String, System.Collections.Generic.List(Of System.String))(System.StringComparer.Ordinal)
        ' Required current-format signatures; old routing formats are not migrated.
        Public Property DocumentCardSignatures As New System.Collections.Generic.Dictionary(Of System.String, System.String)(System.StringComparer.Ordinal)
        Public Property DocumentCount As System.Int32
    End Class

    Public NotInheritable Class SemanticArchiveGenerationManifest
        Public Property Inventory As New SemanticArchiveInventory()
        Public Property SchemaVersion As System.Int32 = SemanticArchiveStore.GenerationSchemaVersion
        Public Property ArchiveId As System.String = ""
        Public Property GenerationId As System.String = ""
        Public Property PreviousGenerationId As System.String = ""
        Public Property ConfigurationSignature As System.String = ""
        Public Property CreatedUtc As System.DateTimeOffset = System.DateTimeOffset.UtcNow
        Public Property FenceToken As System.Int64
        Public Property RootNodeId As System.String = ""
        Public Property MaxChildrenPerNode As System.Int32 = 48
        Public Property MaxRoutingCharacters As System.Int32 = 32000
        Public Property Nodes As New System.Collections.Generic.Dictionary(Of System.String, SemanticArchiveNodeDescriptor)(System.StringComparer.Ordinal)
        Public Property DocumentShards As New System.Collections.Generic.List(Of SemanticArchiveDocumentShardDescriptor)()
        Public Property TermShards As New System.Collections.Generic.List(Of SemanticArchiveArtifactReference)()
        Public Property RoutingGraphArtifact As SemanticArchiveArtifactReference
        Public Property DocumentCount As System.Int32
        Public Property TotalDocumentCount As System.Int32
        Public Property FailureCount As System.Int32
        Public Property ValidationStatus As System.String = "pending"
        Public Property Diagnostics As New System.Collections.Generic.List(Of System.String)()
    End Class

    Public NotInheritable Class SemanticArchiveGenerationPointer
        Public Property SchemaVersion As System.Int32 = SemanticArchiveStore.GenerationSchemaVersion
        Public Property ArchiveId As System.String = ""
        Public Property GenerationId As System.String = ""
        Public Property Manifest As SemanticArchiveArtifactReference
        Public Property FenceToken As System.Int64
        Public Property ActivatedUtc As System.DateTimeOffset = System.DateTimeOffset.UtcNow
    End Class

    Public NotInheritable Class SemanticArchiveValidityRecord
        Public Property SchemaVersion As System.Int32 = 1
        Public Property ArchiveId As System.String = ""
        Public Property DocumentId As System.String = ""
        Public Property SourcePath As System.String = ""
        Public Property SourceHash As System.String = ""
        Public Property RepresentationId As System.String = ""
        Public Property Valid As System.Boolean
        Public Property Reason As System.String = ""
        Public Property UpdatedUtc As System.DateTimeOffset = System.DateTimeOffset.UtcNow
        Public Property AncestorNodeIds As New System.Collections.Generic.List(Of System.String)()
        Public Property FenceToken As System.Int64
    End Class

    Public NotInheritable Class SemanticArchiveSuppressionRecord
        Public Property SchemaVersion As System.Int32 = 1
        Public Property ArchiveId As System.String = ""
        Public Property NodeId As System.String = ""
        Public Property Suppressed As System.Boolean = True
        ' Only this exact immutable artifact may clear a previous suppression.
        Public Property AllowedArtifactHash As System.String = ""
        Public Property Reason As System.String = ""
        Public Property UpdatedUtc As System.DateTimeOffset = System.DateTimeOffset.UtcNow
    End Class

    Public Enum SemanticArchiveAccessDecision
        Unknown = 0
        Allowed = 1
        Denied = 2
    End Enum

    Public NotInheritable Class SemanticArchiveIdentity
        Private Sub New()
        End Sub

        Public Shared Function NewId() As System.String
            Return System.Guid.NewGuid().ToString("N")
        End Function

        Public Shared Function StableId(prefix As System.String, identity As System.String) As System.String
            If System.String.IsNullOrWhiteSpace(prefix) Then Throw New System.ArgumentException("An identity prefix is required.", NameOf(prefix))
            Using hasher As System.Security.Cryptography.SHA256 = System.Security.Cryptography.SHA256.Create()
                Return prefix & "_" & Hex(hasher.ComputeHash(System.Text.Encoding.UTF8.GetBytes(If(identity, ""))))
            End Using
        End Function

        Public Shared Function ComputeFileHash(path As System.String) As System.String
            Using stream As New System.IO.FileStream(path, System.IO.FileMode.Open, System.IO.FileAccess.Read, System.IO.FileShare.Read)
                Using hasher As System.Security.Cryptography.SHA256 = System.Security.Cryptography.SHA256.Create()
                    Return Hex(hasher.ComputeHash(stream))
                End Using
            End Using
        End Function

        Friend Shared Function HashBytes(bytes As System.Byte()) As System.String
            Using hasher As System.Security.Cryptography.SHA256 = System.Security.Cryptography.SHA256.Create()
                Return Hex(hasher.ComputeHash(bytes))
            End Using
        End Function

        Private Shared Function Hex(bytes As System.Byte()) As System.String
            Return System.BitConverter.ToString(bytes).Replace("-", "").ToLowerInvariant()
        End Function

        Public Shared Function ValidateId(value As System.String, argumentName As System.String) As System.String
            If System.String.IsNullOrWhiteSpace(value) OrElse value.Length > 128 Then Throw New System.ArgumentException("A bounded stable identifier is required.", argumentName)
            For Each character As System.Char In value
                If Not ((character >= "a"c AndAlso character <= "z"c) OrElse (character >= "A"c AndAlso character <= "Z"c) OrElse (character >= "0"c AndAlso character <= "9"c) OrElse character = "_"c OrElse character = "-"c) Then
                    Throw New System.ArgumentException("Invalid stable identifier.", argumentName)
                End If
            Next
            If System.Text.RegularExpressions.Regex.IsMatch(value, "^(con|prn|aux|nul|com[0-9]|lpt[0-9])$", System.Text.RegularExpressions.RegexOptions.IgnoreCase Or System.Text.RegularExpressions.RegexOptions.CultureInvariant) Then Throw New System.ArgumentException("A stable identifier may not be a reserved Windows device name.", argumentName)
            Return value
        End Function
    End Class
End Namespace
