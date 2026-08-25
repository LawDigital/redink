' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved.
' For license to use see https://redink.ai.
'
' =============================================================================
' File: ArtifactDelivery.vb
'
' Purpose:
'   Provides the shared, host-agnostic artifact registry and delivery-selection
'   semantics used by Word, Outlook Local Agent, and Outlook AutoPilot.
'
'   The class separates three independent concerns:
'
'     1. Storage location
'        Where an artifact physically exists, for example:
'          - session / staging area
'          - connected persistent workspace
'          - another host-managed location
'
'     2. Artifact lifecycle
'        What role the artifact currently has:
'          - Working
'          - Intermediate
'          - Final
'          - Superseded
'
'     3. Delivery intent
'        Whether the artifact is intended to be surfaced to the user:
'          - None
'          - DeliverToUser
'          - PersistOnly
'          - DeliverAndPersist
'
' Design rules:
'   - Storage location does NOT determine whether a file is a deliverable.
'   - Files may be created temporarily in both staging and a connected workspace.
'   - Multiple files may legitimately belong to one logical deliverable.
'   - Logical deliverables therefore support multiple opaque output slots.
'   - Supersession is evaluated per explicitly identified output slot, not per
'     filename or per logical deliverable as a whole.
'   - Different output slots remain simultaneously deliverable.
'   - Legacy path-only outputs may be tracked as Intermediate state, but they
'     can never establish user-final delivery without explicit opaque IDs.
'
' No-heuristics contract:
'   This implementation MUST NOT infer logical identity, output roles, finality,
'   supersession, or equivalence from:
'
'     - filenames
'     - extensions
'     - directory names
'     - "(2)" or other collision suffixes
'     - "_final", "_markup", "_v2", etc.
'     - source-path similarity
'     - tool names
'     - prompt similarity
'     - semantic similarity of file contents
'
'   Supersession is allowed only from explicit machine-readable identity:
'
'     - LogicalDeliverableId + OutputSlotId
'       and/or
'     - SupersedesArtifactId
'
' Explicit artifact contract for tool authors:
'   File-producing tools SHOULD return an optional `artifacts` array when they
'   can describe their outputs precisely.
'
'   Example:
'
'     {
'       "artifacts": [
'         {
'           "artifact_id": "review-clean-2",
'           "logical_deliverable_id": "review-1",
'           "output_slot_id": "clean",
'           "path": "...",
'           "state": "final",
'           "delivery_intent": "deliver_to_user",
'           "storage_kind": "connected_workspace",
'           "supersedes_artifact_id": "review-clean-1"
'         },
'         {
'           "artifact_id": "review-markup-1",
'           "logical_deliverable_id": "review-1",
'           "output_slot_id": "markup",
'           "path": "...",
'           "state": "final",
'           "delivery_intent": "deliver_to_user",
'           "storage_kind": "session_staging"
'         }
'       ]
'     }
'
' Tool-author guidance:
'   - artifact_id:
'       Opaque identifier for one physical artifact/version.
'       Do not derive it from the filename unless that is already the tool's
'       explicit identity model.
'
'   - logical_deliverable_id:
'       Opaque identifier for the user-level logical result.
'       Several output slots may belong to the same logical deliverable.
'
'   - output_slot_id:
'       Opaque identifier for one independently deliverable output position.
'       Examples such as "clean" or "markup" are allowed, but SharedLibrary
'       treats the value as opaque and never interprets its meaning.
'
'   - supersedes_artifact_id:
'       Optional exact artifact id replaced by the new artifact.
'
'   - state:
'       working | intermediate | final | superseded
'
'   - delivery_intent:
'       none | deliver_to_user | persist_only | deliver_and_persist
'
'   - storage_kind:
'       session_staging | connected_workspace | host_managed | unknown
'
'   A tool that creates several legitimate outputs SHOULD emit one artifact
'   entry per output. It MUST NOT combine them into one slot merely because they
'   share a source or basename.
'
'   A tool that creates temporary files SHOULD register them as Working or
'   Intermediate with no delivery intent if explicit registration is useful.
'   Temporary files are permitted in both staging and connected workspaces.
'
' Legacy compatibility:
'   Tools that do not yet implement the explicit `artifacts` contract may
'   continue returning existing fields such as:
'
'       output_file_path
'       output_files
'       created
'       saved
'       exported
'
'   Legacy outputs may be preserved conservatively as non-deliverable
'   Intermediate state. SharedLibrary never promotes them to user Finals and
'   never infers identity from their paths.
'
' Delivery boundary:
'   ArtifactDelivery decides WHICH artifacts are current user deliverables.
'   It does not perform host-specific transport.
'
'   - Word may copy selected artifacts to Desktop.
'   - Outlook Local Agent may surface/copy selected artifacts.
'   - AutoPilot may materialize selected workspace artifacts into its secure
'     per-mail staging directory and attach them to the reply.
'
'   Host-specific copying, attachment creation, security boundaries, and UI
'   presentation remain outside this class.
'
' Security:
'   Artifact registration does not weaken host path/security policies.
'   A registered workspace artifact may still require host-side materialization
'   into a permitted delivery area before transport.
'
' =============================================================================

Option Strict On
Option Explicit On

Imports System.IO
Imports System.Collections.Generic
Imports System.Linq
Imports Newtonsoft.Json.Linq

Namespace Agents

    Public Enum ArtifactLifecycleState
        Working = 0
        Intermediate = 1
        Final = 2
        Superseded = 3
    End Enum

    Public Enum ArtifactDeliveryIntent
        None = 0
        DeliverToUser = 1
        PersistOnly = 2
        DeliverAndPersist = 3
    End Enum

    Public Enum ArtifactStorageKind
        Unknown = 0
        SessionStaging = 1
        ConnectedWorkspace = 2
        HostManaged = 3
    End Enum

    Public NotInheritable Class ArtifactRegistration
        Public Property ArtifactId As String = ""
        Public Property LogicalDeliverableId As String = ""
        Public Property OutputSlotId As String = ""
        Public Property Path As String = ""
        Public Property LifecycleState As ArtifactLifecycleState = ArtifactLifecycleState.Intermediate
        Public Property DeliveryIntent As ArtifactDeliveryIntent = ArtifactDeliveryIntent.None
        Public Property StorageKind As ArtifactStorageKind = ArtifactStorageKind.Unknown
        Public Property SupersedesArtifactId As String = ""
        Public Property IsExplicitContract As Boolean = True
    End Class


    ' <summary>
    ' Validated optional explicit artifact metadata supplied to a legacy-compatible
    ' single-file producer. A Nothing instance means the call intentionally uses the
    ' existing legacy protocol. No identity is ever inferred from the output path.
    ' </summary>
    Public NotInheritable Class OptionalToolArtifactMetadata
        Public Property ArtifactId As String = ""
        Public Property LogicalDeliverableId As String = ""
        Public Property OutputSlotId As String = ""
        Public Property SupersedesArtifactId As String = ""
        Public Property State As String = ""
        Public Property DeliveryIntent As String = ""
        Public Property StorageKind As String = ""

        Public Function BuildArtifact(outputPath As String) As JObject
            Return New JObject From {
                {"artifact_id", ArtifactId},
                {"logical_deliverable_id", LogicalDeliverableId},
                {"output_slot_id", OutputSlotId},
                {"path", If(outputPath, "")},
                {"state", State},
                {"delivery_intent", DeliveryIntent},
                {"storage_kind", StorageKind},
                {"supersedes_artifact_id", SupersedesArtifactId}
            }
        End Function

        Public ReadOnly Property ProducesUserDeliverable As Boolean
            Get
                Return System.String.Equals(State, "final", System.StringComparison.Ordinal) AndAlso
                       (System.String.Equals(DeliveryIntent, "deliver_to_user", System.StringComparison.Ordinal) OrElse
                        System.String.Equals(DeliveryIntent, "deliver_and_persist", System.StringComparison.Ordinal))
            End Get
        End Property

        Public ReadOnly Property ProducesIntermediateData As Boolean
            Get
                Return System.String.Equals(State, "working", System.StringComparison.Ordinal) OrElse
                       System.String.Equals(State, "intermediate", System.StringComparison.Ordinal)
            End Get
        End Property
    End Class

    Public NotInheritable Class ArtifactDelivery
        Private Sub New()
        End Sub


        ' <summary>
        ' Returns True when a structured tool result explicitly declares an artifacts
        ' property, even when that property is malformed or empty. Once declared, the
        ' explicit protocol is authoritative and callers must not downgrade the same
        ' physical side effect to legacy path-only delivery.
        ' </summary>
        Public Shared Function DeclaresExplicitArtifacts(rootObject As JObject,
                                                         resultObject As JObject) As Boolean
            Return ObjectDeclaresArtifacts(rootObject) OrElse ObjectDeclaresArtifacts(resultObject)
        End Function

        Public Shared Function ResponseDeclaresExplicitArtifacts(responseText As String) As Boolean
            If String.IsNullOrWhiteSpace(responseText) Then Return False

            Try
                Dim token As JToken = JToken.Parse(responseText)
                Dim rootObject As JObject = TryCast(token, JObject)
                If rootObject Is Nothing Then Return False

                Dim resultObject As JObject = TryCast(rootObject("result"), JObject)
                Return DeclaresExplicitArtifacts(rootObject, resultObject)
            Catch ex As System.Exception
                Return False
            End Try
        End Function

        ' <summary>
        ' Validates optional explicit artifact metadata for a single physical output.
        ' Calls that provide no artifact metadata remain fully legacy-compatible. When
        ' any explicit artifact field is present, all opaque identity/finality metadata
        ' must be self-consistent before the producer performs its file side effect.
        ' </summary>

        Public Shared Function EnableOptionalSingleFileArtifactProtocol(tool As SharedLibrary.ModelConfig) As SharedLibrary.ModelConfig
            If tool Is Nothing OrElse System.String.IsNullOrWhiteSpace(tool.ToolDefinition) Then Return tool

            Try
                Dim root As Newtonsoft.Json.Linq.JObject = Newtonsoft.Json.Linq.JObject.Parse(tool.ToolDefinition)
                Dim parameters As Newtonsoft.Json.Linq.JObject = TryCast(root("parameters"), Newtonsoft.Json.Linq.JObject)
                If parameters Is Nothing Then
                    parameters = New Newtonsoft.Json.Linq.JObject()
                    root("parameters") = parameters
                End If

                Dim properties As Newtonsoft.Json.Linq.JObject = TryCast(parameters("properties"), Newtonsoft.Json.Linq.JObject)
                If properties Is Nothing Then
                    properties = New Newtonsoft.Json.Linq.JObject()
                    parameters("properties") = properties
                End If

                If properties("artifact_id") Is Nothing Then properties("artifact_id") = New Newtonsoft.Json.Linq.JObject(New Newtonsoft.Json.Linq.JProperty("type", "string"))
                If properties("logical_deliverable_id") Is Nothing Then properties("logical_deliverable_id") = New Newtonsoft.Json.Linq.JObject(New Newtonsoft.Json.Linq.JProperty("type", "string"))
                If properties("output_slot_id") Is Nothing Then properties("output_slot_id") = New Newtonsoft.Json.Linq.JObject(New Newtonsoft.Json.Linq.JProperty("type", "string"))
                If properties("supersedes_artifact_id") Is Nothing Then properties("supersedes_artifact_id") = New Newtonsoft.Json.Linq.JObject(New Newtonsoft.Json.Linq.JProperty("type", "string"))
                If properties("artifact_state") Is Nothing Then properties("artifact_state") = New Newtonsoft.Json.Linq.JObject(New Newtonsoft.Json.Linq.JProperty("type", "string"), New Newtonsoft.Json.Linq.JProperty("enum", New Newtonsoft.Json.Linq.JArray("working", "intermediate", "final")))
                If properties("artifact_delivery_intent") Is Nothing Then properties("artifact_delivery_intent") = New Newtonsoft.Json.Linq.JObject(New Newtonsoft.Json.Linq.JProperty("type", "string"), New Newtonsoft.Json.Linq.JProperty("enum", New Newtonsoft.Json.Linq.JArray("none", "deliver_to_user", "persist_only", "deliver_and_persist")))
                If properties("storage_kind") Is Nothing Then properties("storage_kind") = New Newtonsoft.Json.Linq.JObject(New Newtonsoft.Json.Linq.JProperty("type", "string"), New Newtonsoft.Json.Linq.JProperty("enum", New Newtonsoft.Json.Linq.JArray("session_staging", "connected_workspace", "host_managed", "unknown")))
                If properties("expected_artifacts") Is Nothing Then
                    properties("expected_artifacts") = New Newtonsoft.Json.Linq.JObject(
                        New Newtonsoft.Json.Linq.JProperty("type", "array"),
                        New Newtonsoft.Json.Linq.JProperty("items", New Newtonsoft.Json.Linq.JObject(
                            New Newtonsoft.Json.Linq.JProperty("type", "object"),
                            New Newtonsoft.Json.Linq.JProperty("properties", New Newtonsoft.Json.Linq.JObject(
                                New Newtonsoft.Json.Linq.JProperty("logical_deliverable_id", New Newtonsoft.Json.Linq.JObject(New Newtonsoft.Json.Linq.JProperty("type", "string"))),
                                New Newtonsoft.Json.Linq.JProperty("output_slot_id", New Newtonsoft.Json.Linq.JObject(New Newtonsoft.Json.Linq.JProperty("type", "string"))))),
                            New Newtonsoft.Json.Linq.JProperty("required", New Newtonsoft.Json.Linq.JArray("logical_deliverable_id", "output_slot_id")))))
                End If

                tool.ToolDefinition = root.ToString(Newtonsoft.Json.Formatting.None)

                Dim artifactProtocolGuidance As String =
                    " EXPLICIT ARTIFACT RULES: If you supply artifact metadata, artifact_id, logical_deliverable_id, and output_slot_id must be opaque stable IDs and must be supplied together. " &
                    "Use artifact_delivery_intent='deliver_to_user' or 'deliver_and_persist' ONLY with artifact_state='final'. " &
                    "A user-facing Final requires expected_artifacts containing the COMPLETE expected output-slot set for the current run/delegation. " &
                    "For mutation tools that can return status='partial' or status='none', do not treat that physical file as complete: retry only the unresolved operations. " &
                    "The runtime defensively downgrades an incomplete mutation result to an intermediate, non-user-deliverable artifact even if the call requested Final."

                If System.String.IsNullOrWhiteSpace(tool.ToolInstructionsPrompt) Then
                    tool.ToolInstructionsPrompt = If(tool.ToolName, "") & ":" & artifactProtocolGuidance
                ElseIf tool.ToolInstructionsPrompt.IndexOf("EXPLICIT ARTIFACT RULES:", System.StringComparison.Ordinal) < 0 Then
                    tool.ToolInstructionsPrompt &= artifactProtocolGuidance
                End If

                If System.String.IsNullOrWhiteSpace(tool.CapabilityTags) Then
                    tool.CapabilityTags = "artifact_generation"
                ElseIf tool.CapabilityTags.IndexOf("artifact_generation", System.StringComparison.OrdinalIgnoreCase) < 0 Then
                    tool.CapabilityTags &= ",artifact_generation"
                End If
            Catch
                ' Schema augmentation must never make an otherwise valid legacy tool unavailable.
            End Try

            Return tool
        End Function

        Public Shared Function AttachOptionalSingleFileArtifactToResult(resultJson As String,
                                                                         metadata As OptionalToolArtifactMetadata,
                                                                         physicalPath As String) As String
            If metadata Is Nothing Then Return resultJson
            If System.String.IsNullOrWhiteSpace(physicalPath) OrElse Not System.IO.File.Exists(physicalPath) Then Return resultJson

            Dim obj As Newtonsoft.Json.Linq.JObject = Nothing
            Try
                obj = TryCast(Newtonsoft.Json.Linq.JToken.Parse(If(resultJson, "")), Newtonsoft.Json.Linq.JObject)
            Catch
            End Try
            If obj Is Nothing Then obj = New Newtonsoft.Json.Linq.JObject(New Newtonsoft.Json.Linq.JProperty("summary", If(resultJson, "")))

            Dim effectiveMetadata As OptionalToolArtifactMetadata = metadata
            Dim incompleteMutationResult As Boolean = IsIncompleteMutationResult(obj)

            If incompleteMutationResult AndAlso
               System.String.Equals(metadata.State, "final", System.StringComparison.Ordinal) Then

                Dim downgradedIntent As String = "none"

                If System.String.Equals(metadata.DeliveryIntent, "persist_only", System.StringComparison.Ordinal) OrElse
                   System.String.Equals(metadata.DeliveryIntent, "deliver_and_persist", System.StringComparison.Ordinal) Then
                    downgradedIntent = "persist_only"
                End If

                effectiveMetadata = New OptionalToolArtifactMetadata With {
                    .ArtifactId = metadata.ArtifactId,
                    .LogicalDeliverableId = metadata.LogicalDeliverableId,
                    .OutputSlotId = metadata.OutputSlotId,
                    .SupersedesArtifactId = metadata.SupersedesArtifactId,
                    .State = "intermediate",
                    .DeliveryIntent = downgradedIntent,
                    .StorageKind = metadata.StorageKind
                }

                obj("artifact_finalization_deferred") = True
                obj("artifact_finalization_reason") =
                    "The mutation result is incomplete (status=partial/none or failed_count>0); the artifact was returned as intermediate and will not be delivered as a Final until the unresolved operations are completed."
            End If

            obj("produces_user_deliverable") = effectiveMetadata.ProducesUserDeliverable
            obj("produces_intermediate_data") = effectiveMetadata.ProducesIntermediateData
            obj("artifacts") = New Newtonsoft.Json.Linq.JArray(Newtonsoft.Json.Linq.JToken.FromObject(effectiveMetadata.BuildArtifact(physicalPath)))
            Return obj.ToString(Newtonsoft.Json.Formatting.None)
        End Function

        Private Shared Function IsIncompleteMutationResult(obj As Newtonsoft.Json.Linq.JObject) As Boolean
            If obj Is Nothing Then Return False

            Dim status As String = If(obj.Value(Of String)("status"), "").Trim()
            If System.String.Equals(status, "partial", System.StringComparison.OrdinalIgnoreCase) OrElse
               System.String.Equals(status, "none", System.StringComparison.OrdinalIgnoreCase) Then
                Return True
            End If

            Dim failedCountToken As Newtonsoft.Json.Linq.JToken = obj("failed_count")
            If failedCountToken IsNot Nothing Then
                Dim failedCount As Integer
                If System.Int32.TryParse(
                    failedCountToken.ToString(),
                    System.Globalization.NumberStyles.Integer,
                    System.Globalization.CultureInfo.InvariantCulture,
                    failedCount) AndAlso failedCount > 0 Then

                    Return True
                End If
            End If

            Return False
        End Function
        Public Shared Function TryPrepareOptionalToolArtifactMetadata(
            arguments As System.Collections.Generic.IDictionary(Of String, System.Object),
            defaultStorageKind As ArtifactStorageKind,
            ByRef metadata As OptionalToolArtifactMetadata,
            ByRef failureCode As String,
            ByRef failureMessage As String) As Boolean

            metadata = Nothing
            failureCode = ""
            failureMessage = ""

            If arguments Is Nothing Then Return True

            Dim artifactId As String = GetArgumentString(arguments, "artifact_id")
            Dim logicalId As String = GetArgumentString(arguments, "logical_deliverable_id")
            Dim slotId As String = GetArgumentString(arguments, "output_slot_id")
            Dim supersedesId As String = GetArgumentString(arguments, "supersedes_artifact_id")
            Dim stateText As String = GetArgumentString(arguments, "artifact_state")
            Dim intentText As String = GetArgumentString(arguments, "artifact_delivery_intent")
            Dim storageText As String = GetArgumentString(arguments, "storage_kind")

            Dim expectedRaw As System.Object = Nothing
            Dim hasExpectedArtifacts As Boolean = arguments.TryGetValue("expected_artifacts", expectedRaw)

            Dim hasAnyExplicitIdentity As Boolean =
                artifactId <> "" OrElse logicalId <> "" OrElse slotId <> "" OrElse supersedesId <> ""

            ' Lifecycle/delivery hints without an explicit identity remain legacy-compatible.
            ' The sequencing validator already treats only the opaque identity fields as the
            ' opt-in boundary. Keeping the delivery preparation aligned prevents a harmless
            ' artifact_state/artifact_delivery_intent hint from causing a retry-loop failure.
            ' expected_artifacts by itself is orchestration context, not artifact identity.
            If Not hasAnyExplicitIdentity Then Return True

            If artifactId = "" OrElse logicalId = "" OrElse slotId = "" Then
                failureCode = "explicit_artifact_identity_incomplete"
                failureMessage = "artifact_id, logical_deliverable_id, and output_slot_id are required together when explicit artifact metadata is used."
                Return False
            End If

            If supersedesId <> "" AndAlso System.String.Equals(artifactId, supersedesId, System.StringComparison.Ordinal) Then
                failureCode = "explicit_artifact_self_supersession"
                failureMessage = "supersedes_artifact_id must not equal artifact_id."
                Return False
            End If

            Dim parsedState As ArtifactLifecycleState
            If stateText = "" OrElse Not TryParseLifecycleState(stateText, parsedState) OrElse
               parsedState = ArtifactLifecycleState.Superseded Then

                failureCode = "invalid_artifact_state"
                failureMessage = "artifact_state must be working, intermediate, or final when this tool produces a file."
                Return False
            End If

            Dim parsedIntent As ArtifactDeliveryIntent
            If intentText = "" OrElse Not TryParseDeliveryIntent(intentText, parsedIntent) Then
                failureCode = "invalid_artifact_delivery_intent"
                failureMessage = "artifact_delivery_intent must be none, deliver_to_user, persist_only, or deliver_and_persist."
                Return False
            End If

            Dim hasUserDeliveryIntent As Boolean =
                parsedIntent = ArtifactDeliveryIntent.DeliverToUser OrElse
                parsedIntent = ArtifactDeliveryIntent.DeliverAndPersist

            If hasUserDeliveryIntent AndAlso parsedState <> ArtifactLifecycleState.Final Then
                failureCode = "invalid_artifact_delivery_state"
                failureMessage = "Only artifact_state=final may use a user-delivery intent."
                Return False
            End If

            Dim parsedStorage As ArtifactStorageKind = defaultStorageKind
            If storageText <> "" AndAlso Not TryParseStorageKind(storageText, parsedStorage) Then
                failureCode = "invalid_artifact_storage_kind"
                failureMessage = "storage_kind must be session_staging, connected_workspace, host_managed, or unknown."
                Return False
            End If

            storageText = StorageKindToProtocolValue(parsedStorage)

            Dim expectedToken As JToken = Nothing
            If hasExpectedArtifacts Then
                If expectedRaw Is Nothing Then
                    failureCode = "invalid_expected_artifacts"
                    failureMessage = "expected_artifacts must be a JSON array when supplied."
                    Return False
                End If

                Try
                    expectedToken = JToken.FromObject(expectedRaw)
                Catch ex As System.Exception
                    failureCode = "invalid_expected_artifacts"
                    failureMessage = "expected_artifacts could not be parsed as a JSON array."
                    Return False
                End Try

                If expectedToken Is Nothing OrElse expectedToken.Type <> JTokenType.Array Then
                    failureCode = "invalid_expected_artifacts"
                    failureMessage = "expected_artifacts must be a JSON array when supplied."
                    Return False
                End If
            End If

            Dim currentSlotDeclared As Boolean = False
            Dim expectedCount As Integer = 0

            If expectedToken IsNot Nothing Then
                For Each expectedItem As JToken In DirectCast(expectedToken, JArray)
                    Dim expectedObject As JObject = TryCast(expectedItem, JObject)
                    If expectedObject Is Nothing Then
                        failureCode = "invalid_expected_artifacts"
                        failureMessage = "Every expected_artifacts item must contain logical_deliverable_id and output_slot_id."
                        Return False
                    End If

                    Dim expectedLogicalId As String = If(expectedObject.Value(Of String)("logical_deliverable_id"), "").Trim()
                    Dim expectedSlotId As String = If(expectedObject.Value(Of String)("output_slot_id"), "").Trim()
                    If expectedLogicalId = "" OrElse expectedSlotId = "" Then
                        failureCode = "invalid_expected_artifacts"
                        failureMessage = "Every expected_artifacts item must contain non-empty logical_deliverable_id and output_slot_id."
                        Return False
                    End If

                    expectedCount += 1
                    If System.String.Equals(expectedLogicalId, logicalId, System.StringComparison.Ordinal) AndAlso
                       System.String.Equals(expectedSlotId, slotId, System.StringComparison.Ordinal) Then
                        currentSlotDeclared = True
                    End If
                Next
            End If

            If parsedState = ArtifactLifecycleState.Final AndAlso hasUserDeliveryIntent Then
                If expectedToken Is Nothing OrElse expectedCount = 0 Then
                    failureCode = "missing_expected_artifacts"
                    failureMessage = "A user-deliverable Final requires expected_artifacts containing the complete expected output-slot set."
                    Return False
                End If

                If Not currentSlotDeclared Then
                    failureCode = "current_output_slot_not_expected"
                    failureMessage = "The current logical_deliverable_id/output_slot_id pair must appear in expected_artifacts."
                    Return False
                End If
            End If

            metadata = New OptionalToolArtifactMetadata With {
                .ArtifactId = artifactId,
                .LogicalDeliverableId = logicalId,
                .OutputSlotId = slotId,
                .SupersedesArtifactId = supersedesId,
                .State = LifecycleStateToProtocolValue(parsedState),
                .DeliveryIntent = DeliveryIntentToProtocolValue(parsedIntent),
                .StorageKind = storageText
            }

            Return True
        End Function

        Public Shared Function Register(runState As ToolCallSequencing.ToolingRunState,
                                        registration As ArtifactRegistration) As ToolCallSequencing.DeliverableArtifact
            If runState Is Nothing OrElse registration Is Nothing Then Return Nothing

            Dim rawPath As String = If(registration.Path, "").Trim()
            If rawPath = "" Then Return Nothing

            Dim fullPath As String
            Try
                If Not System.IO.File.Exists(rawPath) Then Return Nothing
                fullPath = System.IO.Path.GetFullPath(rawPath)
            Catch ex As System.Exception
                Return Nothing
            End Try

            If runState.RegisteredDeliverableArtifacts Is Nothing Then
                runState.RegisteredDeliverableArtifacts = New System.Collections.Generic.List(Of ToolCallSequencing.DeliverableArtifact)()
            End If

            Dim artifactId As String = If(registration.ArtifactId, "").Trim()
            Dim logicalId As String = If(registration.LogicalDeliverableId, "").Trim()
            Dim slotId As String = If(registration.OutputSlotId, "").Trim()
            Dim supersedesId As String = If(registration.SupersedesArtifactId, "").Trim()

            Dim hasAnyExplicitIdentity As Boolean =
                artifactId <> "" OrElse
                logicalId <> "" OrElse
                slotId <> "" OrElse
                supersedesId <> ""

            If hasAnyExplicitIdentity AndAlso
               (artifactId = "" OrElse logicalId = "" OrElse slotId = "") Then
                Return Nothing
            End If

            Dim hasUserDeliveryIntent As Boolean =
                registration.DeliveryIntent = ArtifactDeliveryIntent.DeliverToUser OrElse
                registration.DeliveryIntent = ArtifactDeliveryIntent.DeliverAndPersist

            ' A user-delivery intent is meaningful only for an explicit current Final.
            ' Working/Intermediate/Superseded artifacts may be persisted, but they may
            ' never carry an instruction that would make them user-deliverable.
            If hasUserDeliveryIntent AndAlso
               registration.LifecycleState <> ArtifactLifecycleState.Final Then
                Return Nothing
            End If

            Dim isUserFacingFinal As Boolean =
                registration.LifecycleState = ArtifactLifecycleState.Final AndAlso
                hasUserDeliveryIntent

            ' User-facing finals require complete explicit opaque identity and must
            ' enter through an explicit artifact contract, never a legacy/path-only path.
            If isUserFacingFinal AndAlso
               (Not registration.IsExplicitContract OrElse
                artifactId = "" OrElse
                logicalId = "" OrElse
                slotId = "") Then
                Return Nothing
            End If

            If runState.ExpectedDeliverableContractLocked AndAlso
               isUserFacingFinal AndAlso
               Not runState.IsExpectedDeliverableSlot(logicalId, slotId) Then
                Return Nothing
            End If

            Dim existing As ToolCallSequencing.DeliverableArtifact = Nothing

            ' Artifact identity comes ONLY from the explicit opaque artifact_id.
            ' Never merge registrations merely because their paths are equal.
            If artifactId <> "" Then
                existing = runState.RegisteredDeliverableArtifacts.FirstOrDefault(
                    Function(a)
                        Return a IsNot Nothing AndAlso
                               System.String.Equals(If(a.ArtifactId, ""), artifactId, System.StringComparison.Ordinal)
                    End Function)
            End If

            ' Once an artifact_id has been registered, its explicit identity and physical
            ' version binding are immutable. Path equality is used only as a consistency
            ' check for the SAME artifact_id; it never establishes artifact identity.
            If existing IsNot Nothing Then
                Dim existingLogicalId As String = If(existing.LogicalDeliverableId, "").Trim()
                Dim existingSlotId As String = If(existing.OutputSlotId, "").Trim()
                Dim existingSupersedesId As String = If(existing.SupersedesArtifactId, "").Trim()

                If Not System.String.Equals(
                    existingLogicalId,
                    logicalId,
                    System.StringComparison.Ordinal) Then
                    Return Nothing
                End If

                If Not System.String.Equals(
                    existingSlotId,
                    slotId,
                    System.StringComparison.Ordinal) Then
                    Return Nothing
                End If

                Dim existingPath As String = If(existing.SessionPath, "").Trim()
                If existingPath <> "" Then
                    Try
                        If Not System.String.Equals(
                            System.IO.Path.GetFullPath(existingPath),
                            fullPath,
                            System.StringComparison.OrdinalIgnoreCase) Then
                            Return Nothing
                        End If
                    Catch ex As System.Exception
                        Return Nothing
                    End Try
                End If

                ' The explicit revision edge is part of the immutable identity of this
                ' physical artifact/version. Replaying the same artifact_id may never add,
                ' remove, or redirect supersedes_artifact_id at a later lifecycle stage.
                If Not System.String.Equals(
                    existingSupersedesId,
                    supersedesId,
                    System.StringComparison.Ordinal) Then
                    Return Nothing
                End If

                ' Superseded versions are terminal. Re-registering an old artifact_id must
                ' never resurrect it or alter its delivery semantics.
                If existing.LifecycleState = ArtifactLifecycleState.Superseded Then
                    If registration.LifecycleState <> ArtifactLifecycleState.Superseded OrElse
                       existing.DeliveryIntent <> registration.DeliveryIntent Then
                        Return Nothing
                    End If
                End If

                ' Lifecycle is monotonic. Intermediate may never regress to Working.
                If existing.LifecycleState = ArtifactLifecycleState.Intermediate AndAlso
                   registration.LifecycleState = ArtifactLifecycleState.Working Then
                    Return Nothing
                End If

                ' A current Final is terminal except for an explicit transition to Superseded.
                ' This keeps retries idempotent and prevents a later registration from
                ' downgrading or changing the delivery semantics of the same physical version.
                If existing.LifecycleState = ArtifactLifecycleState.Final Then
                    If registration.LifecycleState <> ArtifactLifecycleState.Final AndAlso
                       registration.LifecycleState <> ArtifactLifecycleState.Superseded Then
                        Return Nothing
                    End If

                    If registration.LifecycleState = ArtifactLifecycleState.Final AndAlso
                       existing.DeliveryIntent <> registration.DeliveryIntent Then
                        Return Nothing
                    End If
                End If

                If existing.StorageKind <> ArtifactStorageKind.Unknown AndAlso
                   registration.StorageKind <> ArtifactStorageKind.Unknown AndAlso
                   existing.StorageKind <> registration.StorageKind Then
                    Return Nothing
                End If
            End If

            Dim supersededArtifact As ToolCallSequencing.DeliverableArtifact = Nothing
            If supersedesId <> "" Then
                supersededArtifact = runState.RegisteredDeliverableArtifacts.FirstOrDefault(
                    Function(a)
                        Return a IsNot Nothing AndAlso
                               System.String.Equals(If(a.ArtifactId, ""), supersedesId, System.StringComparison.Ordinal)
                    End Function)

                ' An explicit supersession reference must resolve to another artifact in the
                ' exact same opaque logical-deliverable/output-slot pair.
                If supersededArtifact Is Nothing OrElse
                   Object.ReferenceEquals(supersededArtifact, existing) OrElse
                   logicalId = "" OrElse slotId = "" OrElse
                   Not System.String.Equals(If(supersededArtifact.LogicalDeliverableId, ""), logicalId, System.StringComparison.Ordinal) OrElse
                   Not System.String.Equals(If(supersededArtifact.OutputSlotId, ""), slotId, System.StringComparison.Ordinal) Then
                    Return Nothing
                End If
            End If

            If supersededArtifact IsNot Nothing AndAlso
               registration.LifecycleState = ArtifactLifecycleState.Final Then

                supersededArtifact.LifecycleState = ArtifactLifecycleState.Superseded
                supersededArtifact.IsFinalDeliverable = False
                supersededArtifact.DeliveryIntent = ArtifactDeliveryIntent.None
            End If

            ' A new final for an explicitly identified slot supersedes older artifacts only
            ' within that exact same opaque slot. Sibling slots remain current.
            If registration.LifecycleState = ArtifactLifecycleState.Final AndAlso
               logicalId <> "" AndAlso slotId <> "" Then

                For Each prior As ToolCallSequencing.DeliverableArtifact In runState.RegisteredDeliverableArtifacts
                    If prior Is Nothing OrElse Object.ReferenceEquals(prior, existing) Then Continue For

                    If System.String.Equals(If(prior.LogicalDeliverableId, ""), logicalId, System.StringComparison.Ordinal) AndAlso
                       System.String.Equals(If(prior.OutputSlotId, ""), slotId, System.StringComparison.Ordinal) Then
                        prior.LifecycleState = ArtifactLifecycleState.Superseded
                        prior.IsFinalDeliverable = False
                        prior.DeliveryIntent = ArtifactDeliveryIntent.None
                    End If
                Next
            End If

            Dim isNewArtifact As Boolean = (existing Is Nothing)

            If existing Is Nothing Then
                existing = New ToolCallSequencing.DeliverableArtifact()
                runState.RegisteredDeliverableArtifacts.Add(existing)
            End If

            existing.ArtifactId = artifactId
            existing.LogicalDeliverableId = logicalId
            existing.OutputSlotId = slotId
            existing.SessionPath = fullPath
            existing.LifecycleState = registration.LifecycleState
            existing.DeliveryIntent = registration.DeliveryIntent

            ' Unknown is absence of storage metadata, not an instruction to erase a
            ' previously established storage kind. A known value may enrich Unknown,
            ' while conflicting known values were rejected above.
            If existing.StorageKind = ArtifactStorageKind.Unknown OrElse
               registration.StorageKind <> ArtifactStorageKind.Unknown Then
                existing.StorageKind = registration.StorageKind
            End If

            existing.SupersedesArtifactId = supersedesId

            If isNewArtifact Then
                existing.IsExplicitContract = registration.IsExplicitContract
                existing.RegisteredUtc = System.DateTime.UtcNow
            ElseIf registration.IsExplicitContract Then
                existing.IsExplicitContract = True
            End If

            existing.IsFinalDeliverable = isUserFacingFinal

            Return existing
        End Function

        Public Shared Function RegisterLegacyPath(runState As ToolCallSequencing.ToolingRunState,
                                                  candidatePath As String,
                                                  sourceTool As String,
                                                  isFinalDeliverable As Boolean) As ToolCallSequencing.DeliverableArtifact
            ' Legacy path-only output has no explicit opaque artifact/slot identity and
            ' therefore can never become a user-deliverable Final. Keep it only as
            ' non-deliverable intermediate telemetry/state.
            Dim reg As New ArtifactRegistration With {
                .Path = candidatePath,
                .LifecycleState = ArtifactLifecycleState.Intermediate,
                .DeliveryIntent = ArtifactDeliveryIntent.None,
                .StorageKind = ArtifactStorageKind.Unknown,
                .IsExplicitContract = False
            }

            Dim artifact As ToolCallSequencing.DeliverableArtifact = Register(runState, reg)
            If artifact IsNot Nothing Then
                artifact.SourceTool = If(sourceTool, "")
                artifact.LegacyCompatibilityEligible =
                    Not runState.HasExpectedDeliverableContract
            End If
            Return artifact
        End Function

        Public Shared Function RegisterExplicitArtifacts(runState As ToolCallSequencing.ToolingRunState,
                                                         rootObject As JObject,
                                                         resultObject As JObject,
                                                         sourceTool As String) As Boolean
            If runState Is Nothing Then Return False

            Dim arrays As New List(Of JArray)()
            AddArtifactArray(arrays, rootObject)
            AddArtifactArray(arrays, resultObject)

            If arrays.Count = 0 Then Return False

            If runState.RegisteredDeliverableArtifacts Is Nothing Then
                runState.RegisteredDeliverableArtifacts =
                    New System.Collections.Generic.List(Of ToolCallSequencing.DeliverableArtifact)()
            End If

            ' Explicit artifacts[] is registered transactionally at the logical-state layer.
            ' A malformed/conflicting sibling must not leave earlier siblings registered or
            ' supersede prior versions while the overall explicit artifact result is rejected.
            ' The physical tool side effect cannot be undone here; such files remain unknown
            ' to delivery unless a later valid explicit registration claims them.
            Dim snapshot As New System.Collections.Generic.List(Of ToolCallSequencing.DeliverableArtifact)()
            For Each artifact As ToolCallSequencing.DeliverableArtifact In runState.RegisteredDeliverableArtifacts
                snapshot.Add(CloneArtifact(artifact))
            Next

            For Each arr In arrays
                For Each token In arr

                    Dim obj As JObject = TryCast(token, JObject)
                    If obj Is Nothing Then
                        RestoreArtifactSnapshot(runState.RegisteredDeliverableArtifacts, snapshot)
                        Return False
                    End If

                    Dim path As String = GetString(obj, "path")
                    If String.IsNullOrWhiteSpace(path) Then
                        RestoreArtifactSnapshot(runState.RegisteredDeliverableArtifacts, snapshot)
                        Return False
                    End If

                    Dim stateText As String = GetString(obj, "state")
                    Dim state As ArtifactLifecycleState = ArtifactLifecycleState.Intermediate

                    If stateText <> "" AndAlso Not TryParseLifecycleState(stateText, state) Then
                        RestoreArtifactSnapshot(runState.RegisteredDeliverableArtifacts, snapshot)
                        Return False
                    End If

                    ' Delivery intent is always explicit. Lifecycle=Final alone never
                    ' implies that the artifact should be delivered to the user.
                    Dim intentText As String = GetString(obj, "delivery_intent")
                    Dim intent As ArtifactDeliveryIntent = ArtifactDeliveryIntent.None

                    If intentText <> "" AndAlso Not TryParseDeliveryIntent(intentText, intent) Then
                        RestoreArtifactSnapshot(runState.RegisteredDeliverableArtifacts, snapshot)
                        Return False
                    End If

                    Dim storageText As String = GetString(obj, "storage_kind")
                    Dim storageKind As ArtifactStorageKind = ArtifactStorageKind.Unknown

                    If storageText <> "" AndAlso Not TryParseStorageKind(storageText, storageKind) Then
                        RestoreArtifactSnapshot(runState.RegisteredDeliverableArtifacts, snapshot)
                        Return False
                    End If

                    Dim reg As New ArtifactRegistration With {
                        .ArtifactId = GetString(obj, "artifact_id"),
                        .LogicalDeliverableId = GetString(obj, "logical_deliverable_id"),
                        .OutputSlotId = GetString(obj, "output_slot_id"),
                        .Path = path,
                        .LifecycleState = state,
                        .DeliveryIntent = intent,
                        .StorageKind = storageKind,
                        .SupersedesArtifactId = GetString(obj, "supersedes_artifact_id"),
                        .IsExplicitContract = True
                    }

                    Dim registeredArtifact As ToolCallSequencing.DeliverableArtifact =
                        Register(runState, reg)

                    If registeredArtifact Is Nothing Then
                        RestoreArtifactSnapshot(runState.RegisteredDeliverableArtifacts, snapshot)
                        Return False
                    End If

                    registeredArtifact.SourceTool = If(sourceTool, "")
                Next
            Next

            ' Presence of an explicit empty artifacts=[] array is a valid authoritative
            ' declaration of zero artifacts. It suppresses legacy fallback but registers
            ' no artifact identity. Any malformed/conflicting non-empty entry returned
            ' False above after restoring the transactional snapshot.
            Return True
        End Function

        Public Shared Function TryRegisterExplicitArtifactsFromResponse(
            runState As ToolCallSequencing.ToolingRunState,
            responseText As String,
            sourceTool As String,
            ByRef failureCode As String,
            ByRef failureMessage As String) As Boolean

            failureCode = ""
            failureMessage = ""

            If runState Is Nothing Then
                failureCode = "artifact_registry_unavailable"
                failureMessage = "The explicit artifact result could not be validated because the run registry is unavailable."
                Return False
            End If

            If System.String.IsNullOrWhiteSpace(responseText) Then
                failureCode = "invalid_explicit_artifact_result"
                failureMessage = "The tool declared an explicit artifact protocol but returned an empty result."
                Return False
            End If

            Try
                Dim token As Newtonsoft.Json.Linq.JToken = Newtonsoft.Json.Linq.JToken.Parse(responseText)
                Dim rootObject As Newtonsoft.Json.Linq.JObject = TryCast(token, Newtonsoft.Json.Linq.JObject)
                If rootObject Is Nothing Then
                    failureCode = "invalid_explicit_artifact_result"
                    failureMessage = "A tool result containing artifacts must be a JSON object."
                    Return False
                End If

                Dim resultObject As Newtonsoft.Json.Linq.JObject = TryCast(rootObject("result"), Newtonsoft.Json.Linq.JObject)
                If Not DeclaresExplicitArtifacts(rootObject, resultObject) Then Return True

                RegisterExplicitArtifactProtocolOwnedPathsFromResponse(runState, rootObject, resultObject)

                If Not RegisterExplicitArtifacts(runState, rootObject, resultObject, sourceTool) Then
                    failureCode = "invalid_explicit_artifact_result"
                    failureMessage = "The tool returned a malformed or conflicting artifacts[] payload. Explicit artifact results are authoritative and cannot fall back to legacy path delivery."
                    Return False
                End If

                Return True
            Catch ex As System.Exception
                failureCode = "invalid_explicit_artifact_result"
                failureMessage = "The explicit artifact result could not be parsed: " & ex.Message
                Return False
            End Try
        End Function

        Private Shared Sub RegisterExplicitArtifactProtocolOwnedPathsFromResponse(
            runState As ToolCallSequencing.ToolingRunState,
            rootObject As Newtonsoft.Json.Linq.JObject,
            resultObject As Newtonsoft.Json.Linq.JObject)

            If runState Is Nothing Then Return

            For Each obj As Newtonsoft.Json.Linq.JObject In New Newtonsoft.Json.Linq.JObject() {rootObject, resultObject}
                If obj Is Nothing Then Continue For

                Dim artifactsToken As Newtonsoft.Json.Linq.JToken = obj("artifacts")
                If artifactsToken IsNot Nothing Then
                    RegisterProtocolOwnedPathsFromArtifactToken(runState, artifactsToken)
                End If

                For Each key As String In New String() {"outputFilePath", "output_file_path", "output_path", "file_path"}
                    Dim valueToken As Newtonsoft.Json.Linq.JToken = obj(key)
                    If valueToken Is Nothing OrElse valueToken.Type = Newtonsoft.Json.Linq.JTokenType.Null Then Continue For
                    runState.RegisterExplicitArtifactProtocolOwnedPath(valueToken.ToString())
                Next

                For Each key As String In New String() {"outputFiles", "output_files"}
                    Dim listToken As Newtonsoft.Json.Linq.JToken = obj(key)
                    If listToken Is Nothing OrElse listToken.Type <> Newtonsoft.Json.Linq.JTokenType.Array Then Continue For
                    For Each item As Newtonsoft.Json.Linq.JToken In DirectCast(listToken, Newtonsoft.Json.Linq.JArray)
                        If item Is Nothing OrElse item.Type = Newtonsoft.Json.Linq.JTokenType.Null Then Continue For
                        runState.RegisterExplicitArtifactProtocolOwnedPath(item.ToString())
                    Next
                Next
            Next
        End Sub

        Private Shared Sub RegisterProtocolOwnedPathsFromArtifactToken(
            runState As ToolCallSequencing.ToolingRunState,
            token As Newtonsoft.Json.Linq.JToken)

            If runState Is Nothing OrElse token Is Nothing Then Return

            If token.Type = Newtonsoft.Json.Linq.JTokenType.Array Then
                For Each item As Newtonsoft.Json.Linq.JToken In DirectCast(token, Newtonsoft.Json.Linq.JArray)
                    RegisterProtocolOwnedPathsFromArtifactToken(runState, item)
                Next
                Return
            End If

            Dim obj As Newtonsoft.Json.Linq.JObject = TryCast(token, Newtonsoft.Json.Linq.JObject)
            If obj Is Nothing Then Return

            Dim pathToken As Newtonsoft.Json.Linq.JToken = obj("path")
            If pathToken IsNot Nothing AndAlso pathToken.Type <> Newtonsoft.Json.Linq.JTokenType.Null Then
                runState.RegisterExplicitArtifactProtocolOwnedPath(pathToken.ToString())
            End If
        End Sub

        Public Shared Function ResolvePathsForDelivery(
            runState As ToolCallSequencing.ToolingRunState,
            Optional legacyCandidates As IEnumerable(Of String) = Nothing,
            Optional allowIncompleteExpectedContract As Boolean = False) As List(Of String)

            ' Preserve one transport entry per explicit current Final. Distinct sibling
            ' artifacts are never collapsed merely because they currently point at the
            ' same physical path.
            Dim resolvedPaths As New List(Of String)()
            Dim resolutionErrors As New List(Of String)()

            If runState Is Nothing Then
                Return resolvedPaths
            End If

            Dim hasExpectedContract As Boolean =
                runState.HasExpectedDeliverableContract

            ' Expected-artifact completeness is authoritative even when the artifact
            ' registry itself is missing. A missing registry must never turn an
            ' incomplete requested sibling set into an empty-but-successful delivery.
            If Not allowIncompleteExpectedContract AndAlso
               hasExpectedContract AndAlso
               runState.ExpectedDeliverableSlots IsNot Nothing AndAlso
               runState.ExpectedDeliverableSlots.Count > 0 AndAlso
               Not runState.HasAllExpectedDeliverableSlots Then

                Throw New System.IO.IOException(
                    "The requested expected-artifact contract is incomplete; automatic delivery of a partial sibling set is blocked.")
            End If

            If runState.RegisteredDeliverableArtifacts Is Nothing Then
                Return resolvedPaths
            End If

            ' Defensive invariant: every exact logical-deliverable/output-slot pair
            ' may have at most one current user-facing Final, regardless of whether this
            ' particular run declared expected_artifacts. Register() normally enforces
            ' this by exact-slot supersession; detecting corruption here prevents
            ' ambiguous delivery from a single logical slot.
            For i As Integer = 0 To runState.RegisteredDeliverableArtifacts.Count - 1
                Dim left As ToolCallSequencing.DeliverableArtifact =
                    runState.RegisteredDeliverableArtifacts(i)

                If left Is Nothing OrElse
                   left.LifecycleState <> ArtifactLifecycleState.Final OrElse
                   Not left.IsFinalDeliverable OrElse
                   Not left.IsExplicitContract OrElse
                   (left.DeliveryIntent <> ArtifactDeliveryIntent.DeliverToUser AndAlso
                    left.DeliveryIntent <> ArtifactDeliveryIntent.DeliverAndPersist) Then

                    Continue For
                End If

                For j As Integer = i + 1 To runState.RegisteredDeliverableArtifacts.Count - 1
                    Dim right As ToolCallSequencing.DeliverableArtifact =
                        runState.RegisteredDeliverableArtifacts(j)

                    If right Is Nothing OrElse
                       right.LifecycleState <> ArtifactLifecycleState.Final OrElse
                       Not right.IsFinalDeliverable OrElse
                       Not right.IsExplicitContract OrElse
                       (right.DeliveryIntent <> ArtifactDeliveryIntent.DeliverToUser AndAlso
                        right.DeliveryIntent <> ArtifactDeliveryIntent.DeliverAndPersist) Then

                        Continue For
                    End If

                    If System.String.Equals(
                        If(left.LogicalDeliverableId, "").Trim(),
                        If(right.LogicalDeliverableId, "").Trim(),
                        System.StringComparison.Ordinal) AndAlso
                       System.String.Equals(
                        If(left.OutputSlotId, "").Trim(),
                        If(right.OutputSlotId, "").Trim(),
                        System.StringComparison.Ordinal) Then

                        Throw New System.IO.IOException(
                            "More than one current user-facing Final is registered for logical slot '" &
                            If(left.LogicalDeliverableId, "").Trim() &
                            "' / '" &
                            If(left.OutputSlotId, "").Trim() &
                            "'.")
                    End If
                Next
            Next

            For Each artifact In runState.RegisteredDeliverableArtifacts
                If artifact Is Nothing Then Continue For

                If artifact.LifecycleState <> ArtifactLifecycleState.Final Then
                    Continue For
                End If

                If Not artifact.IsFinalDeliverable Then
                    Continue For
                End If

                If Not artifact.IsExplicitContract OrElse
                   System.String.IsNullOrWhiteSpace(artifact.ArtifactId) OrElse
                   System.String.IsNullOrWhiteSpace(artifact.LogicalDeliverableId) OrElse
                   System.String.IsNullOrWhiteSpace(artifact.OutputSlotId) Then

                    resolutionErrors.Add(
                        "A user-facing Final is missing its explicit artifact identity/contract metadata.")
                    Continue For
                End If

                If artifact.DeliveryIntent <> ArtifactDeliveryIntent.DeliverToUser AndAlso
                   artifact.DeliveryIntent <> ArtifactDeliveryIntent.DeliverAndPersist Then

                    Continue For
                End If

                ' When the run has an explicit expected-artifact set, delivery is scoped
                ' to that exact set. Shared parent/nested registries may contain Finals
                ' from other delegated tasks; those must never ride along accidentally.
                If hasExpectedContract AndAlso
                   Not runState.IsExpectedDeliverableSlot(
                       If(artifact.LogicalDeliverableId, "").Trim(),
                       If(artifact.OutputSlotId, "").Trim()) Then

                    Continue For
                End If

                Dim artifactId As String = If(artifact.ArtifactId, "").Trim()
                Dim rawPath As String = If(artifact.SessionPath, "").Trim()

                If rawPath = "" Then
                    resolutionErrors.Add(
                        "Registered final artifact '" & artifactId & "' has no delivery path.")
                    Continue For
                End If

                Try
                    Dim fullPath As String = System.IO.Path.GetFullPath(rawPath)

                    If Not System.IO.File.Exists(fullPath) Then
                        resolutionErrors.Add(
                            "Registered final artifact '" & artifactId & "' no longer exists at '" & fullPath & "'.")
                        Continue For
                    End If

                    resolvedPaths.Add(fullPath)
                Catch ex As System.Exception
                    resolutionErrors.Add(
                        "Registered final artifact '" & artifactId & "' has an invalid delivery path: " & ex.Message)
                End Try
            Next

            ' IMPORTANT:
            ' legacyCandidates is intentionally NOT used as an implicit filesystem
            ' discovery mechanism. Storage location alone must never establish
            ' deliverability. Legacy path-only outputs remain Intermediate telemetry
            ' and cannot become user-deliverable Finals.

            If resolutionErrors.Count > 0 Then
                Throw New System.IO.IOException(
                    "One or more registered final deliverables cannot be resolved for delivery: " &
                    System.String.Join(" | ", resolutionErrors))
            End If

            Return resolvedPaths
        End Function

        ''' <summary>
        ''' Captures a physical-file snapshot for the bounded legacy-compatibility path.
        ''' This is NOT artifact identity and never establishes Final state. It is used only
        ''' to attribute newly created/changed files to one exact, successfully executed,
        ''' explicitly deliverable-capable legacy tool call.
        ''' </summary>
        Public Shared Function CaptureLegacyCompatibilityFileSnapshot(
            runState As ToolCallSequencing.ToolingRunState,
            sourceTool As String,
            roots As System.Collections.Generic.IEnumerable(Of String)) As System.Collections.Generic.Dictionary(Of String, String)

            Dim snapshot As New System.Collections.Generic.Dictionary(Of String, String)(
                System.StringComparer.OrdinalIgnoreCase)

            If Not IsLegacyCompatibilityAllowed(runState, sourceTool) OrElse roots Is Nothing Then
                Return snapshot
            End If

            For Each rawRoot As String In roots
                If System.String.IsNullOrWhiteSpace(rawRoot) Then Continue For

                Try
                    Dim fullRoot As String = System.IO.Path.GetFullPath(rawRoot)
                    If Not System.IO.Directory.Exists(fullRoot) Then Continue For

                    For Each filePath As String In System.IO.Directory.EnumerateFiles(
                        fullRoot,
                        "*",
                        System.IO.SearchOption.AllDirectories)

                        Try
                            Dim fullPath As String = System.IO.Path.GetFullPath(filePath)
                            snapshot(fullPath) = GetLegacyCompatibilityFileFingerprint(fullPath)
                        Catch ex As System.Exception
                        End Try
                    Next
                Catch ex As System.Exception
                End Try
            Next

            Return snapshot
        End Function

        ''' <summary>
        ''' Registers only files that appeared or changed during one successful legacy
        ''' deliverable-capable tool call. Registrations remain path-only Intermediate
        ''' telemetry; they are never promoted to explicit Finals.
        ''' </summary>
        Public Shared Sub RegisterLegacyCompatibilityFileDelta(
            runState As ToolCallSequencing.ToolingRunState,
            sourceTool As String,
            roots As System.Collections.Generic.IEnumerable(Of String),
            beforeSnapshot As System.Collections.Generic.IDictionary(Of String, String),
            Optional excludedPaths As System.Collections.Generic.IEnumerable(Of String) = Nothing)

            If Not IsLegacyCompatibilityAllowed(runState, sourceTool) OrElse roots Is Nothing Then
                Return
            End If

            Dim excluded As New System.Collections.Generic.HashSet(Of String)(
                System.StringComparer.OrdinalIgnoreCase)

            If excludedPaths IsNot Nothing Then
                For Each rawExcludedPath As String In excludedPaths
                    If System.String.IsNullOrWhiteSpace(rawExcludedPath) Then Continue For

                    Try
                        excluded.Add(System.IO.Path.GetFullPath(rawExcludedPath))
                    Catch ex As System.Exception
                    End Try
                Next
            End If

            Dim seen As New System.Collections.Generic.HashSet(Of String)(
                System.StringComparer.OrdinalIgnoreCase)

            For Each rawRoot As String In roots
                If System.String.IsNullOrWhiteSpace(rawRoot) Then Continue For

                Try
                    Dim fullRoot As String = System.IO.Path.GetFullPath(rawRoot)
                    If Not System.IO.Directory.Exists(fullRoot) Then Continue For

                    For Each filePath As String In System.IO.Directory.EnumerateFiles(
                        fullRoot,
                        "*",
                        System.IO.SearchOption.AllDirectories)

                        Try
                            Dim fullPath As String = System.IO.Path.GetFullPath(filePath)
                            If Not seen.Add(fullPath) Then Continue For
                            If excluded.Contains(fullPath) Then Continue For
                            If Not System.IO.File.Exists(fullPath) Then Continue For

                            Dim afterFingerprint As String =
                                GetLegacyCompatibilityFileFingerprint(fullPath)

                            Dim beforeFingerprint As String = Nothing
                            Dim existedBefore As Boolean =
                                beforeSnapshot IsNot Nothing AndAlso
                                beforeSnapshot.TryGetValue(fullPath, beforeFingerprint)

                            If existedBefore AndAlso
                               System.String.Equals(
                                   If(beforeFingerprint, ""),
                                   If(afterFingerprint, ""),
                                   System.StringComparison.Ordinal) Then

                                Continue For
                            End If

                            Dim registeredLegacyArtifact As ToolCallSequencing.DeliverableArtifact =
                                RegisterLegacyPath(
                                    runState,
                                    fullPath,
                                    sourceTool,
                                    isFinalDeliverable:=False)

                            If registeredLegacyArtifact IsNot Nothing Then
                                registeredLegacyArtifact.WasObservedLegacyFileDelta = True
                            End If
                        Catch ex As System.Exception
                        End Try
                    Next
                Catch ex As System.Exception
                End Try
            Next
        End Sub

        ''' <summary>
        ''' Resolves only bounded path-only compatibility outputs already attributed to
        ''' explicitly deliverable-capable legacy tools. The result is path-deduplicated
        ''' because legacy entries have no opaque artifact identity. An expected-artifact
        ''' contract, including a locked empty contract, disables this compatibility path.
        ''' </summary>
        Public Shared Function ResolveLegacyCompatibilityPaths(
            runState As ToolCallSequencing.ToolingRunState,
            Optional excludedPaths As System.Collections.Generic.IEnumerable(Of String) = Nothing) As System.Collections.Generic.List(Of String)

            Dim resolved As New System.Collections.Generic.List(Of String)()

            If runState Is Nothing OrElse
               runState.HasExpectedDeliverableContract OrElse
               runState.RegisteredDeliverableArtifacts Is Nothing OrElse
               runState.DeliverableCapableToolNames Is Nothing OrElse
               runState.DeliverableCapableToolNames.Count = 0 Then

                Return resolved
            End If

            Dim excluded As New System.Collections.Generic.HashSet(Of String)(
                System.StringComparer.OrdinalIgnoreCase)

            If excludedPaths IsNot Nothing Then
                For Each rawExcludedPath As String In excludedPaths
                    If System.String.IsNullOrWhiteSpace(rawExcludedPath) Then Continue For

                    Try
                        excluded.Add(System.IO.Path.GetFullPath(rawExcludedPath))
                    Catch ex As System.Exception
                    End Try
                Next
            End If

            Dim seen As New System.Collections.Generic.HashSet(Of String)(
                System.StringComparer.OrdinalIgnoreCase)

            For Each artifact As ToolCallSequencing.DeliverableArtifact In runState.RegisteredDeliverableArtifacts
                If artifact Is Nothing OrElse
                   artifact.IsExplicitContract OrElse
                   Not artifact.LegacyCompatibilityEligible Then

                    Continue For
                End If

                If artifact.LifecycleState <> ArtifactLifecycleState.Intermediate Then Continue For

                Dim sourceTool As String = If(artifact.SourceTool, "").Trim()
                If sourceTool = "" OrElse
                   Not runState.DeliverableCapableToolNames.Contains(sourceTool) Then

                    Continue For
                End If

                Dim rawPath As String = If(artifact.SessionPath, "").Trim()
                If rawPath = "" Then Continue For

                Try
                    Dim fullPath As String = System.IO.Path.GetFullPath(rawPath)

                    If excluded.Contains(fullPath) OrElse
                       Not System.IO.File.Exists(fullPath) OrElse
                       Not seen.Add(fullPath) Then

                        Continue For
                    End If

                    resolved.Add(fullPath)
                Catch ex As System.Exception
                End Try
            Next

            Return resolved
        End Function

        ''' <summary>
        ''' Returns True only when the exact physical path was observed as newly created or
        ''' changed during a successful deliverable-capable legacy tool call in this run.
        ''' Used by hosts to distinguish an unchanged original input path from an in-place
        ''' legacy output without inventing artifact identity.
        ''' </summary>
        Public Shared Function WasObservedLegacyCompatibilityFileDelta(
            runState As ToolCallSequencing.ToolingRunState,
            candidatePath As String) As Boolean

            If runState Is Nothing OrElse
               runState.RegisteredDeliverableArtifacts Is Nothing OrElse
               runState.DeliverableCapableToolNames Is Nothing OrElse
               runState.DeliverableCapableToolNames.Count = 0 OrElse
               System.String.IsNullOrWhiteSpace(candidatePath) Then

                Return False
            End If

            Dim fullCandidatePath As String

            Try
                fullCandidatePath = System.IO.Path.GetFullPath(candidatePath)
            Catch ex As System.Exception
                Return False
            End Try

            For Each artifact As ToolCallSequencing.DeliverableArtifact In runState.RegisteredDeliverableArtifacts
                If artifact Is Nothing OrElse
                   artifact.IsExplicitContract OrElse
                   Not artifact.LegacyCompatibilityEligible OrElse
                   Not artifact.WasObservedLegacyFileDelta OrElse
                   artifact.LifecycleState <> ArtifactLifecycleState.Intermediate Then

                    Continue For
                End If

                Dim sourceTool As String = If(artifact.SourceTool, "").Trim()
                If sourceTool = "" OrElse
                   Not runState.DeliverableCapableToolNames.Contains(sourceTool) Then

                    Continue For
                End If

                Try
                    If System.String.Equals(
                        System.IO.Path.GetFullPath(If(artifact.SessionPath, "")),
                        fullCandidatePath,
                        System.StringComparison.OrdinalIgnoreCase) Then

                        Return True
                    End If
                Catch ex As System.Exception
                End Try
            Next

            Return False
        End Function

        Private Shared Function IsLegacyCompatibilityAllowed(
            runState As ToolCallSequencing.ToolingRunState,
            sourceTool As String) As Boolean

            If runState Is Nothing OrElse
               runState.HasExpectedDeliverableContract OrElse
               runState.DeliverableCapableToolNames Is Nothing OrElse
               runState.DeliverableCapableToolNames.Count = 0 Then

                Return False
            End If

            Dim normalizedToolName As String = If(sourceTool, "").Trim()
            Return normalizedToolName <> "" AndAlso
                   runState.DeliverableCapableToolNames.Contains(normalizedToolName)
        End Function

        Private Shared Function GetLegacyCompatibilityFileFingerprint(filePath As String) As String
            Try
                Dim info As New System.IO.FileInfo(filePath)
                Dim metadataPrefix As String =
                    info.Length.ToString(System.Globalization.CultureInfo.InvariantCulture) &
                    "|" &
                    info.LastWriteTimeUtc.Ticks.ToString(System.Globalization.CultureInfo.InvariantCulture)

                ' Metadata alone can miss an in-place rewrite that deliberately preserves
                ' length/timestamps. Hash the bytes as well so legacy delta attribution does
                ' not silently lose that edge case. If a producer still holds the file open,
                ' keep the metadata fingerprint as a best-effort fallback rather than failing
                ' the successful tool call.
                Try
                    Using sha256 As System.Security.Cryptography.SHA256 =
                        System.Security.Cryptography.SHA256.Create()

                        Using stream As New System.IO.FileStream(
                            filePath,
                            System.IO.FileMode.Open,
                            System.IO.FileAccess.Read,
                            System.IO.FileShare.ReadWrite Or System.IO.FileShare.Delete)

                            Dim hashBytes As Byte() = sha256.ComputeHash(stream)
                            Return metadataPrefix & "|" & System.BitConverter.ToString(hashBytes).Replace("-", "")
                        End Using
                    End Using
                Catch ex As System.Exception
                    Return metadataPrefix & "|<hash-unavailable>"
                End Try
            Catch ex As System.Exception
                Return ""
            End Try
        End Function

        Public Shared Function HasValidatedFinalDeliverable(runState As ToolCallSequencing.ToolingRunState) As Boolean
            If runState Is Nothing OrElse runState.RegisteredDeliverableArtifacts Is Nothing Then Return False

            For Each artifact In runState.RegisteredDeliverableArtifacts
                If artifact Is Nothing Then Continue For
                If artifact.LifecycleState <> ArtifactLifecycleState.Final Then Continue For
                If Not artifact.IsFinalDeliverable Then Continue For
                If Not artifact.IsExplicitContract Then Continue For
                If String.IsNullOrWhiteSpace(artifact.ArtifactId) Then Continue For
                If String.IsNullOrWhiteSpace(artifact.LogicalDeliverableId) Then Continue For
                If String.IsNullOrWhiteSpace(artifact.OutputSlotId) Then Continue For
                If artifact.DeliveryIntent <> ArtifactDeliveryIntent.DeliverToUser AndAlso
                   artifact.DeliveryIntent <> ArtifactDeliveryIntent.DeliverAndPersist Then Continue For
                Try
                    If Not String.IsNullOrWhiteSpace(artifact.SessionPath) AndAlso File.Exists(artifact.SessionPath) Then Return True
                Catch
                End Try
            Next

            Return False
        End Function

        Private Shared Function CloneArtifact(
            source As ToolCallSequencing.DeliverableArtifact) As ToolCallSequencing.DeliverableArtifact

            If source Is Nothing Then Return Nothing

            Return New ToolCallSequencing.DeliverableArtifact With {
                .ArtifactId = If(source.ArtifactId, ""),
                .LogicalDeliverableId = If(source.LogicalDeliverableId, ""),
                .OutputSlotId = If(source.OutputSlotId, ""),
                .SessionPath = If(source.SessionPath, ""),
                .SourceTool = If(source.SourceTool, ""),
                .LegacyCompatibilityEligible = source.LegacyCompatibilityEligible,
                .WasObservedLegacyFileDelta = source.WasObservedLegacyFileDelta,
                .IsFinalDeliverable = source.IsFinalDeliverable,
                .LifecycleState = source.LifecycleState,
                .DeliveryIntent = source.DeliveryIntent,
                .StorageKind = source.StorageKind,
                .SupersedesArtifactId = If(source.SupersedesArtifactId, ""),
                .IsExplicitContract = source.IsExplicitContract,
                .RegisteredUtc = source.RegisteredUtc
            }
        End Function

        Private Shared Sub RestoreArtifactSnapshot(
            target As System.Collections.Generic.List(Of ToolCallSequencing.DeliverableArtifact),
            snapshot As System.Collections.Generic.IEnumerable(Of ToolCallSequencing.DeliverableArtifact))

            If target Is Nothing Then Return

            target.Clear()

            If snapshot Is Nothing Then Return

            For Each artifact As ToolCallSequencing.DeliverableArtifact In snapshot
                target.Add(CloneArtifact(artifact))
            Next
        End Sub

        Private Shared Function ObjectDeclaresArtifacts(obj As JObject) As Boolean
            If obj Is Nothing Then Return False

            For Each propertyToken As JProperty In obj.Properties()
                If System.String.Equals(propertyToken.Name, "artifacts", System.StringComparison.Ordinal) Then
                    Return True
                End If
            Next

            Return False
        End Function

        Private Shared Function GetArgumentString(arguments As System.Collections.Generic.IDictionary(Of String, System.Object),
                                                   key As String) As String
            If arguments Is Nothing OrElse String.IsNullOrWhiteSpace(key) Then Return ""

            Dim raw As System.Object = Nothing
            If Not arguments.TryGetValue(key, raw) OrElse raw Is Nothing Then Return ""
            Return If(System.Convert.ToString(raw), "").Trim()
        End Function

        Private Shared Function LifecycleStateToProtocolValue(value As ArtifactLifecycleState) As String
            Select Case value
                Case ArtifactLifecycleState.Working
                    Return "working"
                Case ArtifactLifecycleState.Intermediate
                    Return "intermediate"
                Case ArtifactLifecycleState.Final
                    Return "final"
                Case ArtifactLifecycleState.Superseded
                    Return "superseded"
                Case Else
                    Return "intermediate"
            End Select
        End Function

        Private Shared Function DeliveryIntentToProtocolValue(value As ArtifactDeliveryIntent) As String
            Select Case value
                Case ArtifactDeliveryIntent.DeliverToUser
                    Return "deliver_to_user"
                Case ArtifactDeliveryIntent.PersistOnly
                    Return "persist_only"
                Case ArtifactDeliveryIntent.DeliverAndPersist
                    Return "deliver_and_persist"
                Case Else
                    Return "none"
            End Select
        End Function

        Private Shared Function StorageKindToProtocolValue(value As ArtifactStorageKind) As String
            Select Case value
                Case ArtifactStorageKind.SessionStaging
                    Return "session_staging"
                Case ArtifactStorageKind.ConnectedWorkspace
                    Return "connected_workspace"
                Case ArtifactStorageKind.HostManaged
                    Return "host_managed"
                Case Else
                    Return "unknown"
            End Select
        End Function

        Private Shared Sub AddArtifactArray(target As List(Of JArray), obj As JObject)
            If obj Is Nothing Then Return
            Dim token As JToken = obj("artifacts")
            If token IsNot Nothing AndAlso token.Type = JTokenType.Array Then target.Add(DirectCast(token, JArray))
        End Sub

        Private Shared Function GetString(obj As JObject, key As String) As String
            If obj Is Nothing OrElse String.IsNullOrWhiteSpace(key) Then Return ""
            Dim token As JToken = obj(key)
            If token Is Nothing OrElse token.Type = JTokenType.Null Then Return ""
            Return token.ToString().Trim()
        End Function

        Private Shared Function TryParseLifecycleState(
            value As String,
            ByRef result As ArtifactLifecycleState) As Boolean

            Select Case If(value, "").Trim().ToLowerInvariant()
                Case "working"
                    result = ArtifactLifecycleState.Working
                    Return True
                Case "intermediate"
                    result = ArtifactLifecycleState.Intermediate
                    Return True
                Case "final"
                    result = ArtifactLifecycleState.Final
                    Return True
                Case "superseded"
                    result = ArtifactLifecycleState.Superseded
                    Return True
                Case Else
                    Return False
            End Select
        End Function

        Private Shared Function TryParseDeliveryIntent(
            value As String,
            ByRef result As ArtifactDeliveryIntent) As Boolean

            Select Case If(value, "").Trim().ToLowerInvariant()
                Case "none"
                    result = ArtifactDeliveryIntent.None
                    Return True
                Case "deliver", "deliver_to_user"
                    result = ArtifactDeliveryIntent.DeliverToUser
                    Return True
                Case "persist", "persist_only"
                    result = ArtifactDeliveryIntent.PersistOnly
                    Return True
                Case "deliver_and_persist"
                    result = ArtifactDeliveryIntent.DeliverAndPersist
                    Return True
                Case Else
                    Return False
            End Select
        End Function

        Private Shared Function TryParseStorageKind(
            value As String,
            ByRef result As ArtifactStorageKind) As Boolean

            Select Case If(value, "").Trim().ToLowerInvariant()
                Case "unknown"
                    result = ArtifactStorageKind.Unknown
                    Return True
                Case "staging", "session_staging"
                    result = ArtifactStorageKind.SessionStaging
                    Return True
                Case "workspace", "connected_workspace"
                    result = ArtifactStorageKind.ConnectedWorkspace
                    Return True
                Case "host_managed"
                    result = ArtifactStorageKind.HostManaged
                    Return True
                Case Else
                    Return False
            End Select
        End Function
    End Class
End Namespace
