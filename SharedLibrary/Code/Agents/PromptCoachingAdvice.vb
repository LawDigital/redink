' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' Evidence-bound request coaching. The model supplies structured, plain-text
' explanations; the host validates provenance/quality and owns all HTML/layout.
' Only short original-request excerpts (never raw runtime logs) enter the outbox.

Option Explicit On
Option Strict On

Namespace Agents

    Public NotInheritable Class PromptCoachingAdvice

        Public Const CurrentSchema As System.Int32 = 2
        Public Const MinimumConfidence As System.Int32 = 95
        Private Shared ReadOnly AllowedTopics As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.Ordinal) From {
            "goal", "sources", "output", "scope", "sequence", "constraints", "verification", "ambiguity"
        }
        Private Shared ReadOnly AllowedEffects As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.Ordinal) From {
            "none", "clarification", "rework", "tool_failure", "extra_steps"
        }
        Private Shared ReadOnly AllowedFormats As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.Ordinal) From {
            "text", "table", "document", "spreadsheet", "presentation", "unspecified"
        }
        Private Shared ReadOnly AllowedImpacts As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.Ordinal) From {
            "blocked", "likely_wrong_result", "substantial_rework", "minor_style"
        }
        Private Shared ReadOnly AllowedSeverities As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.Ordinal) From {
            "low", "medium", "high"
        }

        ''' <summary>Validate durable receipt topics without constructing display advice.</summary>
        Public Shared Function IsCanonicalTopic(value As System.String) As System.Boolean
            Return value IsNot Nothing AndAlso AllowedTopics.Contains(value)
        End Function

        Private Shared Function ControlCommands() As System.String()
            Return New System.String() {"OFF", "ON", "DAILY", "IMMEDIATE", "SYSTEM", "STATUS", "DEFAULT"}
        End Function

        Private Sub New()
        End Sub

        Public NotInheritable Class Recommendation
            Public ReadOnly Property Topic As System.String
            Public ReadOnly Property ObservedEffect As System.String
            Public ReadOnly Property OutputFormat As System.String
            Public ReadOnly Property Confidence As System.Int32
            Public ReadOnly Property Impact As System.String
            Public ReadOnly Property Title As System.String
            Public ReadOnly Property RequestExcerpt As System.String
            Public ReadOnly Property Explanation As System.String
            Public ReadOnly Property WhyItHelps As System.String
            Public ReadOnly Property ImprovedRequest As System.String
            Public ReadOnly Property EffectExplanation As System.String
            Public ReadOnly Property Reviewed As System.Boolean

            Private Sub New(topic As System.String, effect As System.String, outputFormat As System.String,
                            confidence As System.Int32, impact As System.String, title As System.String,
                            excerpt As System.String, explanation As System.String, why As System.String,
                            improved As System.String, effectExplanation As System.String, reviewed As System.Boolean)
                If Not AllowedTopics.Contains(topic) OrElse Not AllowedEffects.Contains(effect) OrElse
                   Not AllowedFormats.Contains(outputFormat) OrElse Not AllowedImpacts.Contains(impact) OrElse
                   impact = "minor_style" OrElse confidence < MinimumConfidence OrElse confidence > 100 Then
                    Throw New System.IO.InvalidDataException("Invalid coaching classification.")
                End If
                Me.Topic = topic
                Me.ObservedEffect = effect
                Me.OutputFormat = outputFormat
                Me.Confidence = confidence
                Me.Impact = impact
                Me.Title = CheckedText(title, 8, 100)
                If System.Text.RegularExpressions.Regex.IsMatch(Me.Title, "^\d+[.)]\s", System.Text.RegularExpressions.RegexOptions.CultureInvariant,
                                                               System.TimeSpan.FromSeconds(1)) Then
                    Throw New System.IO.InvalidDataException("Coaching block headings must not be numbered.")
                End If
                Me.RequestExcerpt = CheckedText(excerpt, 8, 360)
                Me.Explanation = CheckedText(explanation, 35, 650)
                Me.WhyItHelps = CheckedText(why, 30, 550)
                Me.ImprovedRequest = CheckedText(improved, 25, 1200)
                Me.EffectExplanation = If(effect = "none", System.String.Empty, CheckedText(effectExplanation, 35, 550))
                Me.Reviewed = reviewed
            End Sub

            Friend Shared Function Create(candidate As Newtonsoft.Json.Linq.JObject, effect As System.String,
                                          reviewed As System.Boolean, Optional stored As System.Boolean = False) As Recommendation
                Return New Recommendation(ReadEnum(candidate, "topic", AllowedTopics), effect,
                    ReadEnum(candidate, "output_format", AllowedFormats), ReadInteger(candidate, "confidence", MinimumConfidence, 100),
                    ReadEnum(candidate, "impact", AllowedImpacts), ReadString(candidate, "title", False),
                    ReadString(candidate, If(stored, "request_excerpt", "request_evidence"), False),
                    ReadString(candidate, "explanation", False), ReadString(candidate, "why_it_helps", False),
                    ReadString(candidate, "improved_request", False), ReadString(candidate, "effect_explanation", True), reviewed)
            End Function

            Friend Function WithReview() As Recommendation
                Return New Recommendation(Topic, ObservedEffect, OutputFormat, Confidence, Impact, Title, RequestExcerpt,
                    Explanation, WhyItHelps, ImprovedRequest, EffectExplanation, True)
            End Function

            Public Function ToJson() As Newtonsoft.Json.Linq.JObject
                Return New Newtonsoft.Json.Linq.JObject From {
                    {"topic", Topic}, {"observed_effect", ObservedEffect}, {"output_format", OutputFormat},
                    {"confidence", Confidence}, {"impact", Impact}, {"title", Title}, {"request_excerpt", RequestExcerpt},
                    {"explanation", Explanation}, {"why_it_helps", WhyItHelps}, {"improved_request", ImprovedRequest},
                    {"effect_explanation", EffectExplanation}, {"reviewed", Reviewed}
                }
            End Function

            Public Shared Function FromJson(value As Newtonsoft.Json.Linq.JObject) As Recommendation
                If value Is Nothing OrElse Not ReadBoolean(value, "reviewed") Then
                    Throw New System.IO.InvalidDataException("Stored coaching advice has not passed its second review.")
                End If
                Return Create(value, ReadEnum(value, "observed_effect", AllowedEffects), True, stored:=True)
            End Function
        End Class

        Public NotInheritable Class Result
            Public ReadOnly Property Language As System.String
            Public ReadOnly Property Recommendations As System.Collections.Generic.IReadOnlyList(Of Recommendation)
            Public ReadOnly Property ValidationNotes As System.Collections.Generic.IReadOnlyList(Of System.String)

            Friend Sub New(language As System.String,
                           recommendations As System.Collections.Generic.IEnumerable(Of Recommendation),
                           notes As System.Collections.Generic.IEnumerable(Of System.String))
                Me.Language = NormalizeLanguage(language)
                Me.Recommendations = New System.Collections.Generic.List(Of Recommendation)(recommendations).AsReadOnly()
                Me.ValidationNotes = New System.Collections.Generic.List(Of System.String)(notes).AsReadOnly()
            End Sub

            Public Function ToJson() As Newtonsoft.Json.Linq.JObject
                Dim items As New Newtonsoft.Json.Linq.JArray()
                For Each item As Recommendation In Recommendations
                    items.Add(item.ToJson())
                Next
                Return New Newtonsoft.Json.Linq.JObject From {{"schema", CurrentSchema}, {"language", Language}, {"recommendations", items}}
            End Function

            Public Shared Function FromJson(value As Newtonsoft.Json.Linq.JObject) As Result
                If value Is Nothing Then Throw New System.IO.InvalidDataException("Missing coaching result.")
                Dim version As System.Int32 = ReadInteger(value, "schema", 1, CurrentSchema)
                Dim language As System.String = NormalizeLanguage(ReadString(value, "language", False))
                If version = 1 Then
                    ' Old analysed jobs no longer have their original input. Do not invent
                    ' personalised prose or mail their now-obsolete generic templates.
                    Return New Result(language, New Recommendation() {}, New System.String() {"suppressed_legacy_generic_advice"})
                End If
                ValidateResultEnvelope(value)
                Dim result As New System.Collections.Generic.List(Of Recommendation)()
                Dim seen As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.Ordinal)
                For Each token As Newtonsoft.Json.Linq.JToken In DirectCast(value("recommendations"), Newtonsoft.Json.Linq.JArray)
                    Dim item As Recommendation = Recommendation.FromJson(TryCast(token, Newtonsoft.Json.Linq.JObject))
                    If Not seen.Add(item.Topic) Then Throw New System.IO.InvalidDataException("Duplicate stored coaching topic.")
                    result.Add(item)
                Next
                Return New Result(language, result, New System.String() {})
            End Function
        End Class

        Public Shared Function Parse(raw As System.String, originalRequest As System.String,
                                     logText As System.String, allowObservedEffects As System.Boolean) As Result
            Dim payload As Newtonsoft.Json.Linq.JObject = ParseJsonObject(raw)
            ValidateResultEnvelope(payload)
            Dim language As System.String = NormalizeLanguage(ReadString(payload, "language", False))
            Dim recommendations As New System.Collections.Generic.List(Of Recommendation)()
            Dim notes As New System.Collections.Generic.List(Of System.String)()
            Dim seen As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.Ordinal)
            Dim request As System.String = If(originalRequest, System.String.Empty)
            Dim log As System.String = If(logText, System.String.Empty)
            For Each token As Newtonsoft.Json.Linq.JToken In DirectCast(payload("recommendations"), Newtonsoft.Json.Linq.JArray)
                Dim candidate As Newtonsoft.Json.Linq.JObject = TryCast(token, Newtonsoft.Json.Linq.JObject)
                If candidate Is Nothing Then Throw New System.IO.InvalidDataException("Coaching advice must be an object.")
                Dim topic As System.String = ReadEnum(candidate, "topic", AllowedTopics)
                Dim severity As System.String = ReadEnum(candidate, "severity", AllowedSeverities)
                Dim impact As System.String = ReadEnum(candidate, "impact", AllowedImpacts)
                Dim confidence As System.Int32 = ReadInteger(candidate, "confidence", 0, 100)
                If Not ReadBoolean(candidate, "material") OrElse severity <> "high" OrElse
                   impact = "minor_style" OrElse confidence < MinimumConfidence Then
                    notes.Add("suppressed_below_high_materiality_threshold:" & topic)
                    Continue For
                End If
                Dim evidence As System.String = ReadString(candidate, "request_evidence", False).Trim()
                If evidence.Length < 8 OrElse evidence.Length > 360 OrElse request.IndexOf(evidence, System.StringComparison.Ordinal) < 0 Then
                    notes.Add("suppressed_missing_or_nonliteral_request_excerpt:" & topic)
                    Continue For
                End If
                Dim effect As System.String = ReadEnum(candidate, "observed_effect", AllowedEffects)
                Dim logEvidence As System.String = ReadString(candidate, "log_evidence", True).Trim()
                If effect <> "none" Then
                    If Not allowObservedEffects OrElse logEvidence.Length < 15 OrElse logEvidence.Length > 500 OrElse
                       log.IndexOf(logEvidence, System.StringComparison.Ordinal) < 0 OrElse
                       request.IndexOf(logEvidence, System.StringComparison.Ordinal) >= 0 OrElse
                       ReadInteger(candidate, "effect_confidence", 0, 100) < MinimumConfidence Then
                        effect = "none"
                        notes.Add("effect_omitted_without_sufficient_runtime_evidence:" & topic)
                    End If
                End If
                Try
                    Dim item As Recommendation = Recommendation.Create(candidate, effect, False)
                    If Not seen.Add(topic) Then
                        notes.Add("suppressed_duplicate_topic:" & topic)
                        Continue For
                    End If
                    recommendations.Add(item)
                Catch ex As System.IO.InvalidDataException
                    ' Weak/unsafe display text is not silently replaced by boilerplate.
                    notes.Add("suppressed_incomplete_or_unsafe_contextual_advice:" & topic)
                End Try
            Next
            Return New Result(language, recommendations, notes)
        End Function

        ''' <summary>Independent conservative review; every candidate must receive one explicit decision.</summary>
        Public Shared Function ApplyReview(raw As System.String, proposals As Result) As Result
            If proposals Is Nothing Then Throw New System.ArgumentNullException(NameOf(proposals))
            Dim payload As Newtonsoft.Json.Linq.JObject = ParseJsonObject(raw)
            If ReadInteger(payload, "schema", 1, 1) <> 1 Then Throw New System.IO.InvalidDataException("Invalid coaching review schema.")
            Dim decisions As Newtonsoft.Json.Linq.JArray = TryCast(payload("decisions"), Newtonsoft.Json.Linq.JArray)
            If decisions Is Nothing OrElse decisions.Count <> proposals.Recommendations.Count Then
                Throw New System.IO.InvalidDataException("Incomplete coaching review decisions.")
            End If
            Dim accepted As New System.Collections.Generic.List(Of Recommendation)()
            Dim notes As New System.Collections.Generic.List(Of System.String)(proposals.ValidationNotes)
            Dim seen As New System.Collections.Generic.HashSet(Of System.Int32)()
            For Each token As Newtonsoft.Json.Linq.JToken In decisions
                Dim decision As Newtonsoft.Json.Linq.JObject = TryCast(token, Newtonsoft.Json.Linq.JObject)
                If decision Is Nothing Then Throw New System.IO.InvalidDataException("Invalid coaching review decision.")
                Dim id As System.Int32 = ReadInteger(decision, "candidate", 1, proposals.Recommendations.Count)
                If Not seen.Add(id) Then Throw New System.IO.InvalidDataException("Duplicate coaching review decision.")
                Dim approve As System.Boolean = ReadBoolean(decision, "approve")
                For Each field As System.String In New System.String() {"substantial", "specific", "faithful", "safe", "context_sufficient"}
                    Dim passed As System.Boolean = ReadBoolean(decision, field)
                    approve = approve AndAlso passed
                Next
                Dim confidence As System.Int32 = ReadInteger(decision, "confidence", 0, 100)
                If approve AndAlso confidence >= MinimumConfidence Then
                    accepted.Add(proposals.Recommendations(id - 1).WithReview())
                Else
                    notes.Add("suppressed_by_independent_quality_review:" & proposals.Recommendations(id - 1).Topic)
                End If
            Next
            Return New Result(proposals.Language, accepted, notes)
        End Function

        Private Shared Function ParseJsonObject(raw As System.String) As Newtonsoft.Json.Linq.JObject
            Dim text As System.String = If(raw, System.String.Empty).Trim()
            If text.Length = 0 OrElse text.Length > 65536 Then Throw New System.IO.InvalidDataException("Invalid coaching response size.")
            If text.StartsWith("```", System.StringComparison.Ordinal) Then
                Dim newline As System.Int32 = text.IndexOf(Microsoft.VisualBasic.vbLf, System.StringComparison.Ordinal)
                If newline < 0 OrElse Not text.EndsWith("```", System.StringComparison.Ordinal) Then Throw New System.IO.InvalidDataException("Incomplete coaching JSON fence.")
                Dim kind As System.String = text.Substring(3, newline - 3).Trim()
                If kind.Length > 0 AndAlso Not kind.Equals("json", System.StringComparison.OrdinalIgnoreCase) Then Throw New System.IO.InvalidDataException("Unsupported coaching JSON fence.")
                text = text.Substring(newline + 1, text.Length - newline - 4).Trim()
            End If
            Try
                Using input As New System.IO.StringReader(text)
                    Using reader As New Newtonsoft.Json.JsonTextReader(input)
                        reader.MaxDepth = 6
                        reader.DateParseHandling = Newtonsoft.Json.DateParseHandling.None
                        Dim settings As New Newtonsoft.Json.Linq.JsonLoadSettings With {
                            .DuplicatePropertyNameHandling = Newtonsoft.Json.Linq.DuplicatePropertyNameHandling.Error
                        }
                        Dim result As Newtonsoft.Json.Linq.JObject = Newtonsoft.Json.Linq.JObject.Load(reader, settings)
                        If reader.Read() Then Throw New System.IO.InvalidDataException("Unexpected data after coaching JSON.")
                        Return result
                    End Using
                End Using
            Catch ex As Newtonsoft.Json.JsonException
                Throw New System.IO.InvalidDataException("Invalid coaching JSON.", ex)
            End Try
        End Function

        Private Shared Sub ValidateResultEnvelope(value As Newtonsoft.Json.Linq.JObject)
            If value Is Nothing OrElse ReadInteger(value, "schema", CurrentSchema, CurrentSchema) <> CurrentSchema Then
                Throw New System.IO.InvalidDataException("Unsupported coaching result schema.")
            End If
            Dim items As Newtonsoft.Json.Linq.JArray = TryCast(value("recommendations"), Newtonsoft.Json.Linq.JArray)
            If items Is Nothing OrElse items.Count > 3 Then Throw New System.IO.InvalidDataException("Invalid coaching recommendation array.")
        End Sub

        Private Shared Function ReadString(value As Newtonsoft.Json.Linq.JObject, name As System.String, optionalField As System.Boolean) As System.String
            Dim token As Newtonsoft.Json.Linq.JToken = value(name)
            If token Is Nothing AndAlso optionalField Then Return System.String.Empty
            If token Is Nothing OrElse token.Type <> Newtonsoft.Json.Linq.JTokenType.String Then Throw New System.IO.InvalidDataException("Invalid coaching text field: " & name)
            Return token.ToObject(Of System.String)()
        End Function

        Private Shared Function ReadBoolean(value As Newtonsoft.Json.Linq.JObject, name As System.String) As System.Boolean
            Dim token As Newtonsoft.Json.Linq.JToken = value(name)
            If token Is Nothing OrElse token.Type <> Newtonsoft.Json.Linq.JTokenType.Boolean Then Throw New System.IO.InvalidDataException("Invalid coaching boolean field: " & name)
            Return token.ToObject(Of System.Boolean)()
        End Function

        Private Shared Function ReadInteger(value As Newtonsoft.Json.Linq.JObject, name As System.String,
                                           minimum As System.Int32, maximum As System.Int32) As System.Int32
            Dim token As Newtonsoft.Json.Linq.JToken = value(name)
            If token Is Nothing OrElse token.Type <> Newtonsoft.Json.Linq.JTokenType.Integer Then Throw New System.IO.InvalidDataException("Invalid coaching integer field: " & name)
            Dim number As System.Int64
            If Not System.Int64.TryParse(token.ToString(), System.Globalization.NumberStyles.Integer,
                                        System.Globalization.CultureInfo.InvariantCulture, number) OrElse
               number < minimum OrElse number > maximum Then Throw New System.IO.InvalidDataException("Coaching integer outside allowed range: " & name)
            Return System.Convert.ToInt32(number)
        End Function

        Private Shared Function ReadEnum(value As Newtonsoft.Json.Linq.JObject, name As System.String,
                                        allowed As System.Collections.Generic.HashSet(Of System.String)) As System.String
            Dim normalized As System.String = ReadString(value, name, False).Trim().ToLowerInvariant()
            If Not allowed.Contains(normalized) Then Throw New System.IO.InvalidDataException("Unknown coaching classification: " & name)
            Return normalized
        End Function

        Private Shared Function NormalizeLanguage(value As System.String) As System.String
            Dim normalized As System.String = If(value, System.String.Empty).Trim().ToLowerInvariant()
            Select Case normalized
                Case "de", "en", "fr", "it" : Return normalized
                Case Else : Throw New System.IO.InvalidDataException("Unsupported coaching language.")
            End Select
        End Function

        Private Shared Function CheckedText(value As System.String, minimum As System.Int32, maximum As System.Int32) As System.String
            Dim text As System.String = If(value, System.String.Empty).Trim()
            If text.Length < minimum OrElse text.Length > maximum Then Throw New System.IO.InvalidDataException("Coaching display text outside its limits.")
            For Each character As System.Char In text
                Dim category As System.Globalization.UnicodeCategory = System.Char.GetUnicodeCategory(character)
                If (System.Char.IsControl(character) AndAlso character <> Microsoft.VisualBasic.ChrW(9) AndAlso
                    character <> Microsoft.VisualBasic.ChrW(10) AndAlso character <> Microsoft.VisualBasic.ChrW(13)) OrElse
                   category = System.Globalization.UnicodeCategory.Format Then
                    Throw New System.IO.InvalidDataException("Coaching display text contains hidden controls.")
                End If
            Next
            ' Conservative credential protection, not a promise of general-purpose DLP.
            ' No raw logs/attachments are copied. Unsafe excerpts are suppressed rather
            ' than silently rewritten into a quote which the user never actually sent.
            If System.Text.RegularExpressions.Regex.IsMatch(text,
                "(?i)\b(password|passwort|secret|api[ _-]?key|access[ _-]?token|authorization|bearer)\s*[:=]|-----BEGIN[^\r\n]*PRIVATE KEY-----|\b[A-Za-z0-9_+/=-]{64,}\b",
                System.Text.RegularExpressions.RegexOptions.CultureInvariant, System.TimeSpan.FromSeconds(1)) Then
                Throw New System.IO.InvalidDataException("Coaching display text may contain a credential.")
            End If
            Return text
        End Function

        Private Shared Function SelectForRendering(recommendations As System.Collections.Generic.IEnumerable(Of Recommendation)) As System.Collections.Generic.List(Of Recommendation)
            Dim items As New System.Collections.Generic.List(Of Recommendation)()
            If recommendations IsNot Nothing Then
                For Each item As Recommendation In recommendations
                    If item Is Nothing Then Continue For
                    If Not item.Reviewed Then Throw New System.IO.InvalidDataException("Unreviewed coaching advice cannot be rendered.")
                    items.Add(item)
                Next
            End If
            items.Sort(Function(left As Recommendation, right As Recommendation)
                           Dim rank As System.Int32 = right.Confidence.CompareTo(left.Confidence)
                           Return If(rank <> 0, rank, System.StringComparer.Ordinal.Compare(left.Topic, right.Topic))
                       End Function)
            Dim selected As New System.Collections.Generic.List(Of Recommendation)()
            Dim seen As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.Ordinal)
            For Each item As Recommendation In items
                If seen.Add(item.Topic) Then selected.Add(item)
                If selected.Count = 3 Then Exit For
            Next
            Return selected
        End Function

        Private Shared Function Labels(language As System.String) As System.String()
            Select Case NormalizeLanguage(language)
                Case "de"
                    Return {"Ein konkreter Tipp zu Ihrem Auftrag an Inky", "Ihre gesammelten Inky-Tipps", "Anhand Ihres Auftrags lässt sich eine wichtige Unklarheit vermeiden. Die folgende Anregung ist eine mögliche Präzisierung, keine Bewertung Ihrer Arbeit.",
                        "Aus Ihrem Auftrag", "Was noch offenbleibt", "Warum diese Präzisierung hilft", "So könnte Ihr Auftrag lauten", "Bezug zum Ablauf",
                        "Der Ablauf liefert dazu einen Hinweis, aber keinen Beweis, dass Ihre Formulierung ein technisches Problem verursacht hat.",
                        "Sie entscheiden, ob und wann Sie Inky-Tipps erhalten. Senden Sie eine neue Nachricht an Inky und schreiben Sie nur eines der folgenden Kommandos in den Nachrichtentext, ohne Signatur oder zitierten Verlauf.",
                        "Abstellen", "Anstellen", "Einmal täglich (ab 20 Uhr)", "Nach einer auswertbaren Anfrage", "Versandzeit gemäss Systemeinstellung", "Status", "Alle persönlichen Einstellungen zurücksetzen", "Notwendiges Wort im Betreff", "Auch eine Bitte in natürlicher Sprache ist möglich, zum Beispiel: «Sende mir Inky-Tipps nur einmal täglich.»", "Die Tipps werden nur bei erheblichem Verbesserungspotenzial verschickt; nicht nach jeder Nachricht. Platzhalter im Vorschlag ergänzen Sie mit Ihren Angaben."}
                Case "fr"
                    Return {"Un conseil concret pour votre demande à Inky", "Vos conseils Inky regroupés", "Votre demande permet d'illustrer une ambiguïté importante. La suggestion suivante est une précision possible, pas une évaluation de votre travail.",
                        "Extrait de votre demande", "Ce qui reste à préciser", "Pourquoi cette précision aide", "Une formulation possible", "Lien avec le déroulement",
                        "Le déroulement apporte un indice, mais ne prouve pas que votre formulation a causé un problème technique.",
                        "Vous choisissez si et quand recevoir les conseils Inky. Envoyez un nouveau message à Inky contenant uniquement une commande ci-dessous, sans signature ni texte cité.",
                        "Désactiver", "Activer", "Une fois par jour (dès 20 h)", "Après une demande analysée", "Horaire selon le système", "État", "Réinitialiser tous les choix personnels", "Mot nécessaire dans l'objet", "Une demande en langage naturel est aussi possible : «Envoie-moi les conseils Inky une seule fois par jour.»", "Les conseils sont envoyés uniquement en cas d'amélioration importante, pas après chaque message. Complétez les champs entre crochets avec vos indications."}
                Case "it"
                    Return {"Un consiglio concreto per la vostra richiesta a Inky", "I vostri consigli Inky raccolti", "La vostra richiesta permette di chiarire un'ambiguità importante. La proposta seguente è una possibile precisazione, non una valutazione del vostro lavoro.",
                        "Dalla vostra richiesta", "Che cosa resta da chiarire", "Perché questa precisazione aiuta", "Una possibile formulazione", "Collegamento con l'esecuzione",
                        "L'esecuzione offre un indizio, ma non dimostra che la formulazione abbia causato un problema tecnico.",
                        "Scegliete se e quando ricevere i consigli Inky. Inviate un nuovo messaggio a Inky con un solo comando riportato sotto, senza firma né testo citato.",
                        "Disattivare", "Attivare", "Una volta al giorno (dalle 20)", "Dopo una richiesta analizzata", "Orario secondo il sistema", "Stato", "Ripristinare tutte le scelte personali", "Parola necessaria nell'oggetto", "È possibile anche una richiesta in linguaggio naturale: «Inviami i consigli Inky solo una volta al giorno.»", "I consigli vengono inviati solo per miglioramenti importanti, non dopo ogni messaggio. Completate i segnaposto con le vostre indicazioni."}
                Case Else
                    Return {"A concrete suggestion for your request to Inky", "Your collected Inky tips", "Your request provides an example of an important ambiguity that can be avoided. The following is a possible clarification, not an assessment of your work.",
                        "From your request", "What remains unclear", "Why this clarification helps", "A possible version of your request", "Connection to the run",
                        "The run provides an indication, not proof that your wording caused a technical problem.",
                        "You choose whether and when to receive Inky tips. Send a new message to Inky with just one command below in the body, without a signature or quoted text.",
                        "Turn off", "Turn on", "Once daily (from 20:00)", "After an analysed request", "Use the system delivery schedule", "Status", "Reset all personal choices", "Required word in the subject", "You can also ask in natural language, for example: 'Send me Inky tips only once a day.'", "Tips are sent only for substantial improvements, not after every message. Fill in the placeholders with your own details."}
            End Select
        End Function

        Private Shared Function H(value As System.String) As System.String
            Dim encoded As System.String = System.Net.WebUtility.HtmlEncode(If(value, System.String.Empty))
            Return encoded.Replace(Microsoft.VisualBasic.vbCrLf, Microsoft.VisualBasic.vbLf).Replace(Microsoft.VisualBasic.vbCr, Microsoft.VisualBasic.vbLf).Replace(Microsoft.VisualBasic.vbLf, "<br/>")
        End Function

        ''' <summary>HTML fragment only; the host appends its normal product footer and document wrapper.</summary>
        Public Shared Function RenderHtml(language As System.String,
                                          recommendations As System.Collections.Generic.IEnumerable(Of Recommendation),
                                          isDaily As System.Boolean, dateLabel As System.String,
                                          Optional subjectTriggerWord As System.String = Nothing) As System.String
            Dim items As System.Collections.Generic.List(Of Recommendation) = SelectForRendering(recommendations)
            If items.Count = 0 Then Return System.String.Empty
            Dim text As System.String() = Labels(language)
            Dim body As New System.Text.StringBuilder()
            body.Append("<div style='font-family:Arial,sans-serif;font-size:11pt;line-height:1.45;color:#222222;'>")
            body.Append("<h1 style='font-size:17pt;margin:0 0 12px 0;'>").Append(H(If(isDaily, text(1), text(0)))).Append("</h1>")
            If isDaily AndAlso Not System.String.IsNullOrWhiteSpace(dateLabel) Then body.Append("<p style='color:#666666;'>").Append(H(CheckedText(dateLabel, 1, 60))).Append("</p>")
            body.Append("<p>").Append(H(text(2))).Append("</p>")
            For Each item As Recommendation In items
                body.Append("<table role='presentation' cellpadding='0' cellspacing='0' width='100%' style='margin:18px 0;border:1px solid #dddddd;'><tr><td style='padding:16px;'>")
                body.Append("<h2 style='font-size:13pt;margin:0 0 12px 0;'>").Append(H(item.Title)).Append("</h2>")
                AppendHtmlSection(body, text(3), item.RequestExcerpt, True)
                AppendHtmlSection(body, text(4), item.Explanation, False)
                AppendHtmlSection(body, text(5), item.WhyItHelps, False)
                AppendHtmlSection(body, text(6), item.ImprovedRequest, True)
                If item.ObservedEffect <> "none" Then AppendHtmlSection(body, text(7), item.EffectExplanation & System.Environment.NewLine & text(8), False)
                body.Append("</td></tr></table>")
            Next
            body.Append("<p>").Append(H(text(19))).Append("</p>")
            body.Append("<div style='border-top:1px solid #dddddd;margin-top:24px;padding-top:12px;font-size:9pt;line-height:1.4;color:#666666;'>")
            body.Append("<p style='font-size:9pt;'>").Append(H(text(9))).Append("</p>")
            Dim commands As System.String() = ControlCommands()
            body.Append("<table role='presentation' cellpadding='0' cellspacing='0' style='font-size:9pt;color:#666666;'>")
            For index As System.Int32 = 0 To commands.Length - 1
                body.Append("<tr><td style='font-size:9pt;padding:2px 14px 2px 0;'>").Append(H(text(10 + index))).Append("</td><td style='font-size:9pt;padding:2px 0;'><code>INKY_TIPS ").Append(commands(index)).Append("</code></td></tr>")
            Next
            body.Append("</table>")
            If Not System.String.IsNullOrWhiteSpace(subjectTriggerWord) Then
                body.Append("<p style='font-size:9pt;'>").Append(H(text(17))).Append(": <strong>").Append(H(subjectTriggerWord.Trim())).Append("</strong></p>")
            End If
            body.Append("<p style='font-size:9pt;'>").Append(H(text(18))).Append("</p>")
            body.Append("</div></div>")
            Return body.ToString()
        End Function

        Private Shared Sub AppendHtmlSection(body As System.Text.StringBuilder, label As System.String,
                                             content As System.String, boxed As System.Boolean)
            body.Append("<p style='margin:12px 0 4px 0;'><strong>").Append(H(label)).Append("</strong></p>")
            body.Append(If(boxed, "<p style='margin:0;padding:10px;background-color:#f5f5f5;border-left:3px solid #bbbbbb;'>", "<p style='margin:0;'>"))
            body.Append(H(content)).Append("</p>")
        End Sub

        ''' <summary>Plain-text representation for non-HTML consumers; no ordinal numbering.</summary>
        Public Shared Function Render(language As System.String,
                                      recommendations As System.Collections.Generic.IEnumerable(Of Recommendation),
                                      isDaily As System.Boolean, dateLabel As System.String,
                                      Optional subjectTriggerWord As System.String = Nothing) As System.String
            Dim items As System.Collections.Generic.List(Of Recommendation) = SelectForRendering(recommendations)
            If items.Count = 0 Then Return System.String.Empty
            Dim text As System.String() = Labels(language)
            Dim body As New System.Text.StringBuilder()
            body.AppendLine(If(isDaily, text(1), text(0)))
            If isDaily AndAlso Not System.String.IsNullOrWhiteSpace(dateLabel) Then body.AppendLine(CheckedText(dateLabel, 1, 60))
            body.AppendLine(text(2))
            For Each item As Recommendation In items
                body.AppendLine().AppendLine(item.Title)
                body.AppendLine(text(3) & ": " & item.RequestExcerpt)
                body.AppendLine(text(4) & ": " & item.Explanation)
                body.AppendLine(text(5) & ": " & item.WhyItHelps)
                body.AppendLine(text(6) & ": " & item.ImprovedRequest)
                If item.ObservedEffect <> "none" Then body.AppendLine(text(7) & ": " & item.EffectExplanation & " " & text(8))
            Next
            body.AppendLine().AppendLine(text(19)).AppendLine().AppendLine(text(9))
            Dim commands As System.String() = ControlCommands()
            For index As System.Int32 = 0 To commands.Length - 1
                body.AppendLine(text(10 + index) & ": INKY_TIPS " & commands(index))
            Next
            If Not System.String.IsNullOrWhiteSpace(subjectTriggerWord) Then body.AppendLine(text(17) & ": " & subjectTriggerWord.Trim())
            body.AppendLine(text(18))
            Return body.ToString().TrimEnd()
        End Function

    End Class
End Namespace
