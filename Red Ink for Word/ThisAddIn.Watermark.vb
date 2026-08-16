' Part of "Red Ink for Word"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: ThisAddIn.Watermark.vb
' Purpose: Detects watermark-like signals in text using Unicode/steganography
'          analysis, token statistics, and optional user-supplied source profiles.
' Target: .NET Framework 4.8 / VB.NET / VSTO
' Dependency: Microsoft.ML.Tokenizers
' =============================================================================

Option Explicit On
Option Strict On
Option Infer On

Imports SharedLibrary.SharedLibrary
Imports SharedLibrary.SharedLibrary.SharedMethods
Imports Whisper.net.LibraryLoader
Imports SLib = SharedLibrary.SharedLibrary.SharedMethods

Partial Public Class ThisAddIn

    Private Sub ShowWatermarkAnalysisWindow(
    text As System.String
)

        Try

            If System.String.IsNullOrWhiteSpace(text) Then
                ShowCustomMessageBox("No text was supplied to the watermark analyzer.")
                Return
            End If

            Dim tokenizer As Microsoft.ML.Tokenizers.Tokenizer =
            Microsoft.ML.Tokenizers.TiktokenTokenizer.CreateForModel(
                "gpt-4"
            )

            Dim detector As New UniversalTextWatermarkDetector(
                tokenizer:=tokenizer,
                profiles:=Nothing
            )

            Dim watermarkResult As UniversalTextWatermarkResult =
            detector.Analyze(text)

            Dim sourceText As System.String =
            "No source profile configured."

            If watermarkResult.BestSourceMatch IsNot Nothing Then

                sourceText =
                watermarkResult.BestSourceMatch.SourceName &
                " (" &
                watermarkResult.BestSourceMatch.Score.ToString(
                    "0.0",
                    System.Globalization.CultureInfo.InvariantCulture
                ) &
                ")"

            End If

            Dim interpretation As System.String

            If watermarkResult.WatermarkDetected Then

                interpretation =
                "Verified watermark evidence detected."

            ElseIf watermarkResult.WatermarkSuspected Then

                interpretation =
                "Watermark-like statistical evidence found, but no watermark was verified."

            Else

                interpretation =
                "No verified watermark evidence detected."

            End If

            Dim details As New System.Text.StringBuilder()

            details.AppendLine(interpretation)
            details.AppendLine()

            details.AppendLine(
            "Overall score: " &
            watermarkResult.OverallScore.ToString(
                "0.0",
                System.Globalization.CultureInfo.InvariantCulture
            )
        )

            details.AppendLine(
            "Statistical score: " &
            watermarkResult.StatisticalScore.ToString(
                "0.0",
                System.Globalization.CultureInfo.InvariantCulture
            )
        )

            details.AppendLine(
            "Unicode / steganography score: " &
            watermarkResult.UnicodeScore.ToString(
                "0.0",
                System.Globalization.CultureInfo.InvariantCulture
            )
        )

            details.AppendLine()
            details.AppendLine(
            "Token count: " &
            watermarkResult.TokenCount.ToString(
                System.Globalization.CultureInfo.InvariantCulture
            )
        )

            details.AppendLine(
            "Invisible characters: " &
            If(
                watermarkResult.HasInvisibleCharacters,
                "Yes",
                "No"
            )
        )

            details.AppendLine(
            "Encoded Unicode pattern: " &
            If(
                watermarkResult.PossibleEncodedUnicodePattern,
                "Possible",
                "No"
            )
        )

            details.AppendLine()

            details.AppendLine(
            "Zero-width characters: " &
            watermarkResult.ZeroWidthCount.ToString(
                System.Globalization.CultureInfo.InvariantCulture
            )
        )

            details.AppendLine(
            "Variation selectors: " &
            watermarkResult.VariationSelectorCount.ToString(
                System.Globalization.CultureInfo.InvariantCulture
            )
        )

            details.AppendLine(
            "Bidirectional controls: " &
            watermarkResult.BidirectionalControlCount.ToString(
                System.Globalization.CultureInfo.InvariantCulture
            )
        )

            details.AppendLine(
            "Other format characters: " &
            watermarkResult.FormatCharacterCount.ToString(
                System.Globalization.CultureInfo.InvariantCulture
            )
        )

            details.AppendLine()
            details.AppendLine(
            "Best source profile: " &
            sourceText
        )

            Dim features() As System.Double =
            watermarkResult.GetFeatures()

            details.AppendLine()
            details.AppendLine("Statistical feature vector:")

            For index As System.Int32 = 0 To features.Length - 1

                details.AppendLine(
                "[" &
                index.ToString(
                    System.Globalization.CultureInfo.InvariantCulture
                ) &
                "] = " &
                features(index).ToString(
                    "0.000000",
                    System.Globalization.CultureInfo.InvariantCulture
                )
            )

            Next

            ShowCustomWindow("These are the results of the watermark analysis:", details.ToString(), "You can copy them to the clipboard (and edit them beforehand).", AN & " Watermark Detector")

        Catch ex As System.Exception

            ShowCustomMessageBox("Watermark analysis failed:" & ex.ToString())

        End Try

    End Sub

    Public NotInheritable Class UniversalTextWatermarkDetector

        Public Const FeatureCount As Integer = 10

        Private ReadOnly _tokenizer As Microsoft.ML.Tokenizers.Tokenizer
        Private ReadOnly _profiles As System.Collections.Generic.List(Of UniversalTextSourceProfile)

        Public Sub New(
            tokenizer As Microsoft.ML.Tokenizers.Tokenizer,
            Optional profiles As System.Collections.Generic.IEnumerable(Of UniversalTextSourceProfile) = Nothing
        )

            If tokenizer Is Nothing Then
                Throw New System.ArgumentNullException(NameOf(tokenizer))
            End If

            _tokenizer = tokenizer
            _profiles = New System.Collections.Generic.List(Of UniversalTextSourceProfile)()

            If profiles IsNot Nothing Then
                For Each profile As UniversalTextSourceProfile In profiles
                    If profile IsNot Nothing Then
                        _profiles.Add(profile)
                    End If
                Next
            End If

        End Sub

        Public Function Analyze(text As String) As UniversalTextWatermarkResult

            If text Is Nothing Then
                Throw New System.ArgumentNullException(NameOf(text))
            End If

            Dim unicodeResult As UnicodeWatermarkResult = AnalyzeUnicode(text)

            Dim tokenIds As System.Collections.Generic.IReadOnlyList(Of Integer) =
                _tokenizer.EncodeToIds(text)

            Dim features() As Double = CalculateFeatures(tokenIds)

            Dim statisticalScore As Double =
                CalculateStatisticalScore(features, tokenIds.Count)

            Dim bestMatch As UniversalTextSourceMatch = Nothing

            For Each profile As UniversalTextSourceProfile In _profiles

                Dim currentMatch As UniversalTextSourceMatch = profile.Compare(features)

                If bestMatch Is Nothing OrElse currentMatch.Score > bestMatch.Score Then
                    bestMatch = currentMatch
                End If

            Next

            Dim profileScore As Double = 0.0

            If bestMatch IsNot Nothing Then
                profileScore = bestMatch.Score
            End If

            Dim overallScore As Double =
                System.Math.Max(
                    unicodeResult.Score,
                    System.Math.Max(statisticalScore, profileScore)
                )

            ' A verified watermark decision is deliberately restricted to direct
            ' structural encoding evidence. Generic statistics and provenance profiles
            ' are reported as suspicion only.
            Dim detected As Boolean = unicodeResult.StrongEvidence

            Dim suspected As Boolean =
                Not detected AndAlso
                (statisticalScore >= 60.0 OrElse
                 (bestMatch IsNot Nothing AndAlso bestMatch.Score >= 75.0))

            Return New UniversalTextWatermarkResult(
                tokenCount:=tokenIds.Count,
                watermarkDetected:=detected,
                watermarkSuspected:=suspected,
                overallScore:=overallScore,
                unicodeScore:=unicodeResult.Score,
                statisticalScore:=statisticalScore,
                hasInvisibleCharacters:=unicodeResult.HasInvisibleCharacters,
                possibleEncodedUnicodePattern:=unicodeResult.PossibleEncodedPattern,
                zeroWidthCount:=unicodeResult.ZeroWidthCount,
                variationSelectorCount:=unicodeResult.VariationSelectorCount,
                bidirectionalControlCount:=unicodeResult.BidirectionalControlCount,
                formatCharacterCount:=unicodeResult.FormatCharacterCount,
                bestSourceMatch:=bestMatch,
                features:=features
            )

        End Function

        Private Shared Function AnalyzeUnicode(text As String) As UnicodeWatermarkResult

            Dim suspiciousCount As Integer = 0
            Dim zeroWidthCount As Integer = 0
            Dim variationSelectorCount As Integer = 0
            Dim bidiCount As Integer = 0
            Dim formatCount As Integer = 0

            Dim distinctInvisible As New System.Collections.Generic.HashSet(Of Integer)()

            Dim index As Integer = 0

            While index < text.Length

                Dim codePoint As Integer = 0
                Dim characterLength As Integer = 0

                GetCodePoint(text, index, codePoint, characterLength)

                Dim suspicious As Boolean = False

                Select Case codePoint

                    Case &H200B, &H200C, &H200D, &H2060, &HFEFF
                        zeroWidthCount += 1
                        suspicious = True
                        distinctInvisible.Add(codePoint)

                    Case &HFE00 To &HFE0F, &HE0100 To &HE01EF
                        variationSelectorCount += 1
                        suspicious = True
                        distinctInvisible.Add(codePoint)

                    Case &H202A To &H202E, &H2066 To &H2069
                        bidiCount += 1
                        suspicious = True

                    Case Else
                        Dim category As System.Globalization.UnicodeCategory =
                            GetUnicodeCategory(codePoint)

                        If category = System.Globalization.UnicodeCategory.Format Then
                            formatCount += 1
                            suspicious = True
                        End If

                End Select

                If suspicious Then
                    suspiciousCount += 1
                End If

                index += characterLength

            End While

            Dim score As Double = 0.0

            If zeroWidthCount > 0 Then
                score += 20.0 + System.Math.Min(30.0, CDbl(zeroWidthCount) * 2.0)
            End If

            If variationSelectorCount > 0 Then
                score += 20.0 + System.Math.Min(30.0, CDbl(variationSelectorCount) * 2.0)
            End If

            If bidiCount > 0 Then
                score += System.Math.Min(10.0, CDbl(bidiCount) * 2.0)
            End If

            If formatCount > 0 Then
                score += System.Math.Min(10.0, CDbl(formatCount))
            End If

            Dim possibleEncodedPattern As Boolean =
                distinctInvisible.Count = 2 AndAlso suspiciousCount >= 8

            If possibleEncodedPattern Then
                score += 25.0
            End If

            score = System.Math.Min(100.0, score)

            Return New UnicodeWatermarkResult(
                score:=score,
                hasInvisibleCharacters:=(suspiciousCount > 0),
                possibleEncodedPattern:=possibleEncodedPattern,
                strongEvidence:=(score >= 50.0),
                zeroWidthCount:=zeroWidthCount,
                variationSelectorCount:=variationSelectorCount,
                bidirectionalControlCount:=bidiCount,
                formatCharacterCount:=formatCount
            )

        End Function

        Private Shared Function CalculateFeatures(
            tokenIds As System.Collections.Generic.IReadOnlyList(Of Integer)
        ) As Double()

            Dim result(FeatureCount - 1) As Double

            If tokenIds Is Nothing OrElse tokenIds.Count = 0 Then
                Return result
            End If

            result(0) = CalculateUniqueTokenRatio(tokenIds)
            result(1) = CalculateNGramRepeatRate(tokenIds, 2)
            result(2) = CalculateNGramRepeatRate(tokenIds, 3)
            result(3) = CalculateNGramRepeatRate(tokenIds, 4)
            result(4) = CalculateModuloDeviation(tokenIds, 2)
            result(5) = CalculateModuloDeviation(tokenIds, 4)
            result(6) = CalculateModuloDeviation(tokenIds, 8)
            result(7) = CalculateHashDeviation(tokenIds, 3, 2)
            result(8) = CalculateHashDeviation(tokenIds, 4, 4)
            result(9) = CalculateHashDeviation(tokenIds, 5, 8)

            Return result

        End Function

        Private Shared Function CalculateStatisticalScore(
            features() As Double,
            tokenCount As Integer
        ) As Double

            If tokenCount < 40 Then
                Return 0.0
            End If

            Dim value As Double =
                (features(4) +
                 features(5) +
                 features(6) +
                 features(7) +
                 features(8) +
                 features(9)) / 6.0

            Dim sampleFactor As Double =
                System.Math.Min(1.0, CDbl(tokenCount) / 400.0)

            Return System.Math.Min(100.0, value * 100.0 * sampleFactor)

        End Function

        Private Shared Function CalculateUniqueTokenRatio(
            tokenIds As System.Collections.Generic.IReadOnlyList(Of Integer)
        ) As Double

            Dim values As New System.Collections.Generic.HashSet(Of Integer)()

            For Each tokenId As Integer In tokenIds
                values.Add(tokenId)
            Next

            Return CDbl(values.Count) / CDbl(tokenIds.Count)

        End Function

        Private Shared Function CalculateNGramRepeatRate(
            tokenIds As System.Collections.Generic.IReadOnlyList(Of Integer),
            ngramLength As Integer
        ) As Double

            If tokenIds.Count < ngramLength Then
                Return 0.0
            End If

            Dim total As Integer = tokenIds.Count - ngramLength + 1
            Dim uniqueValues As New System.Collections.Generic.HashSet(Of Long)()

            For i As Integer = 0 To total - 1
                uniqueValues.Add(CalculateHash(tokenIds, i, ngramLength))
            Next

            Return 1.0 - (CDbl(uniqueValues.Count) / CDbl(total))

        End Function

        Private Shared Function CalculateModuloDeviation(
            tokenIds As System.Collections.Generic.IReadOnlyList(Of Integer),
            bucketCount As Integer
        ) As Double

            Dim buckets(bucketCount - 1) As Integer

            For Each tokenId As Integer In tokenIds

                Dim bucket As Integer = tokenId Mod bucketCount

                If bucket < 0 Then
                    bucket += bucketCount
                End If

                buckets(bucket) += 1

            Next

            Return CalculateBucketDeviation(buckets, tokenIds.Count)

        End Function

        Private Shared Function CalculateHashDeviation(
            tokenIds As System.Collections.Generic.IReadOnlyList(Of Integer),
            ngramLength As Integer,
            bucketCount As Integer
        ) As Double

            If tokenIds.Count < ngramLength Then
                Return 0.0
            End If

            Dim total As Integer = tokenIds.Count - ngramLength + 1
            Dim buckets(bucketCount - 1) As Integer

            For i As Integer = 0 To total - 1

                Dim hashValue As Long = CalculateHash(tokenIds, i, ngramLength)
                Dim bucket As Integer = CInt(hashValue Mod CLng(bucketCount))

                If bucket < 0 Then
                    bucket += bucketCount
                End If

                buckets(bucket) += 1

            Next

            Return CalculateBucketDeviation(buckets, total)

        End Function

        Private Shared Function CalculateBucketDeviation(
            buckets() As Integer,
            total As Integer
        ) As Double

            If buckets Is Nothing OrElse buckets.Length < 2 OrElse total <= 0 Then
                Return 0.0
            End If

            Dim expected As Double = 1.0 / CDbl(buckets.Length)
            Dim squared As Double = 0.0

            For Each count As Integer In buckets

                Dim actual As Double = CDbl(count) / CDbl(total)
                Dim difference As Double = actual - expected

                squared += difference * difference

            Next

            Dim maxDeviation As Double = 1.0 - (1.0 / CDbl(buckets.Length))

            If maxDeviation <= 0.0 Then
                Return 0.0
            End If

            Return System.Math.Min(
                1.0,
                System.Math.Sqrt(squared / maxDeviation)
            )

        End Function

        Private Shared Function CalculateHash(
            tokenIds As System.Collections.Generic.IReadOnlyList(Of Integer),
            startIndex As Integer,
            ngramLength As Integer
        ) As Long

            Const Prime As Long = 2147483647L
            Const Multiplier As Long = 1000003L

            Dim value As Long = 216613626L

            For offset As Integer = 0 To ngramLength - 1

                Dim token As Long = CLng(tokenIds(startIndex + offset)) Mod Prime

                If token < 0L Then
                    token += Prime
                End If

                value =
                    ((value * Multiplier) +
                     token +
                     (CLng(offset + 1) * 97L)) Mod Prime

            Next

            Return value

        End Function

        Private Shared Sub GetCodePoint(
            text As String,
            index As Integer,
            ByRef codePoint As Integer,
            ByRef characterLength As Integer
        )

            Dim current As Char = text.Chars(index)

            If System.Char.IsHighSurrogate(current) AndAlso
               index + 1 < text.Length AndAlso
               System.Char.IsLowSurrogate(text.Chars(index + 1)) Then

                codePoint =
                    System.Char.ConvertToUtf32(
                        current,
                        text.Chars(index + 1)
                    )

                characterLength = 2

            Else

                codePoint = System.Convert.ToInt32(current)
                characterLength = 1

            End If

        End Sub

        Private Shared Function GetUnicodeCategory(
            codePoint As Integer
        ) As System.Globalization.UnicodeCategory

            Dim value As String = System.Char.ConvertFromUtf32(codePoint)

            Return System.Globalization.CharUnicodeInfo.GetUnicodeCategory(value, 0)

        End Function

        Private NotInheritable Class UnicodeWatermarkResult

            Public Sub New(
                score As Double,
                hasInvisibleCharacters As Boolean,
                possibleEncodedPattern As Boolean,
                strongEvidence As Boolean,
                zeroWidthCount As Integer,
                variationSelectorCount As Integer,
                bidirectionalControlCount As Integer,
                formatCharacterCount As Integer
            )

                Me.Score = score
                Me.HasInvisibleCharacters = hasInvisibleCharacters
                Me.PossibleEncodedPattern = possibleEncodedPattern
                Me.StrongEvidence = strongEvidence
                Me.ZeroWidthCount = zeroWidthCount
                Me.VariationSelectorCount = variationSelectorCount
                Me.BidirectionalControlCount = bidirectionalControlCount
                Me.FormatCharacterCount = formatCharacterCount

            End Sub

            Public ReadOnly Property Score As Double
            Public ReadOnly Property HasInvisibleCharacters As Boolean
            Public ReadOnly Property PossibleEncodedPattern As Boolean
            Public ReadOnly Property StrongEvidence As Boolean
            Public ReadOnly Property ZeroWidthCount As Integer
            Public ReadOnly Property VariationSelectorCount As Integer
            Public ReadOnly Property BidirectionalControlCount As Integer
            Public ReadOnly Property FormatCharacterCount As Integer

        End Class

    End Class

    Public NotInheritable Class UniversalTextWatermarkResult

        Private ReadOnly _features() As Double

        Public Sub New(
            tokenCount As Integer,
            watermarkDetected As Boolean,
            watermarkSuspected As Boolean,
            overallScore As Double,
            unicodeScore As Double,
            statisticalScore As Double,
            hasInvisibleCharacters As Boolean,
            possibleEncodedUnicodePattern As Boolean,
            zeroWidthCount As Integer,
            variationSelectorCount As Integer,
            bidirectionalControlCount As Integer,
            formatCharacterCount As Integer,
            bestSourceMatch As UniversalTextSourceMatch,
            features() As Double
        )

            If features Is Nothing Then
                Throw New System.ArgumentNullException(NameOf(features))
            End If

            Me.TokenCount = tokenCount
            Me.WatermarkDetected = watermarkDetected
            Me.WatermarkSuspected = watermarkSuspected
            Me.OverallScore = overallScore
            Me.UnicodeScore = unicodeScore
            Me.StatisticalScore = statisticalScore
            Me.HasInvisibleCharacters = hasInvisibleCharacters
            Me.PossibleEncodedUnicodePattern = possibleEncodedUnicodePattern
            Me.ZeroWidthCount = zeroWidthCount
            Me.VariationSelectorCount = variationSelectorCount
            Me.BidirectionalControlCount = bidirectionalControlCount
            Me.FormatCharacterCount = formatCharacterCount
            Me.BestSourceMatch = bestSourceMatch

            _features = CType(features.Clone(), Double())

        End Sub

        Public ReadOnly Property TokenCount As Integer
        Public ReadOnly Property WatermarkDetected As Boolean
        Public ReadOnly Property WatermarkSuspected As Boolean
        Public ReadOnly Property OverallScore As Double
        Public ReadOnly Property UnicodeScore As Double
        Public ReadOnly Property StatisticalScore As Double
        Public ReadOnly Property HasInvisibleCharacters As Boolean
        Public ReadOnly Property PossibleEncodedUnicodePattern As Boolean
        Public ReadOnly Property ZeroWidthCount As Integer
        Public ReadOnly Property VariationSelectorCount As Integer
        Public ReadOnly Property BidirectionalControlCount As Integer
        Public ReadOnly Property FormatCharacterCount As Integer
        Public ReadOnly Property BestSourceMatch As UniversalTextSourceMatch

        Public ReadOnly Property HasEnoughTokensForStatisticalAnalysis As Boolean
            Get
                Return TokenCount >= 40
            End Get
        End Property

        Public Function GetFeatures() As Double()
            Return CType(_features.Clone(), Double())
        End Function

    End Class

    Public NotInheritable Class UniversalTextSourceProfile

        Private ReadOnly _expected() As Double
        Private ReadOnly _tolerances() As Double

        Public Sub New(
            sourceName As String,
            expectedValues() As Double,
            Optional tolerances() As Double = Nothing
        )

            If System.String.IsNullOrWhiteSpace(sourceName) Then
                Throw New System.ArgumentException(
                    "Source name is required.",
                    NameOf(sourceName)
                )
            End If

            ValidateFeatureArray(expectedValues, NameOf(expectedValues), allowZero:=True)

            If tolerances IsNot Nothing Then
                ValidateFeatureArray(tolerances, NameOf(tolerances), allowZero:=False)
            End If

            Me.SourceName = sourceName
            _expected = CType(expectedValues.Clone(), Double())

            If tolerances Is Nothing Then

                ReDim _tolerances(UniversalTextWatermarkDetector.FeatureCount - 1)

                For i As Integer = 0 To _tolerances.Length - 1
                    _tolerances(i) = 0.05
                Next

            Else

                _tolerances = CType(tolerances.Clone(), Double())

            End If

        End Sub

        Public ReadOnly Property SourceName As String

        Public Function Compare(values() As Double) As UniversalTextSourceMatch

            ValidateFeatureArray(values, NameOf(values), allowZero:=True)

            Dim squaredDistance As Double = 0.0

            For i As Integer = 0 To UniversalTextWatermarkDetector.FeatureCount - 1

                Dim normalizedDifference As Double =
                    (values(i) - _expected(i)) / _tolerances(i)

                squaredDistance += normalizedDifference * normalizedDifference

            Next

            Dim distance As Double =
                System.Math.Sqrt(
                    squaredDistance /
                    CDbl(UniversalTextWatermarkDetector.FeatureCount)
                )

            Dim score As Double = 100.0 / (1.0 + distance)

            Return New UniversalTextSourceMatch(
                sourceName:=SourceName,
                score:=score,
                distance:=distance
            )

        End Function

        Public Function GetExpectedValues() As Double()
            Return CType(_expected.Clone(), Double())
        End Function

        Public Function GetTolerances() As Double()
            Return CType(_tolerances.Clone(), Double())
        End Function

        Private Shared Sub ValidateFeatureArray(
            values() As Double,
            parameterName As String,
            allowZero As Boolean
        )

            If values Is Nothing Then
                Throw New System.ArgumentNullException(parameterName)
            End If

            If values.Length <> UniversalTextWatermarkDetector.FeatureCount Then
                Throw New System.ArgumentException(
                    "Exactly " &
                    UniversalTextWatermarkDetector.FeatureCount.ToString(
                        System.Globalization.CultureInfo.InvariantCulture
                    ) &
                    " values are required.",
                    parameterName
                )
            End If

            For i As Integer = 0 To values.Length - 1

                If System.Double.IsNaN(values(i)) OrElse
                   System.Double.IsInfinity(values(i)) Then

                    Throw New System.ArgumentException(
                        "The array may not contain NaN or Infinity.",
                        parameterName
                    )
                End If

                If Not allowZero AndAlso values(i) <= 0.0 Then
                    Throw New System.ArgumentException(
                        "All tolerance values must be greater than zero.",
                        parameterName
                    )
                End If

            Next

        End Sub

    End Class

    Public NotInheritable Class UniversalTextSourceMatch

        Public Sub New(
            sourceName As String,
            score As Double,
            distance As Double
        )

            Me.SourceName = sourceName
            Me.Score = score
            Me.Distance = distance

        End Sub

        Public ReadOnly Property SourceName As String
        Public ReadOnly Property Score As Double
        Public ReadOnly Property Distance As Double

    End Class

End Class
