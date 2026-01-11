' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: SharedMethods.Diff.vb
' Purpose: Defines a minimal diff token structure used to represent text differences
'          as a sequence of operations (equal/insert/delete) with associated text.
'
' Architecture:
'  - Container Type: `SharedMethods.Diff` represents one diff segment.
'  - Operation Kind: `Diff.Operation` indicates how `Text` should be interpreted
'    (unchanged, inserted, or deleted).
'  - Segment Payload: `Diff.Text` carries the segment content for the given operation.
' =============================================================================

Option Strict On
Option Explicit On

Namespace SharedLibrary

    Partial Public Class SharedMethods

        ''' <summary>
        ''' Represents a single diff segment consisting of an operation and its associated text.
        ''' </summary>
        Public Class Diff

            ''' <summary>
            ''' Specifies the kind of change represented by a diff segment.
            ''' </summary>
            Public Enum Operation
                ''' <summary>
                ''' Indicates that the text is unchanged between compared inputs.
                ''' </summary>
                Equal

                ''' <summary>
                ''' Indicates that the text was inserted.
                ''' </summary>
                Insert

                ''' <summary>
                ''' Indicates that the text was deleted.
                ''' </summary>
                Delete
            End Enum

            ''' <summary>
            ''' Gets or sets the operation associated with this diff segment.
            ''' </summary>
            Public Property Op As Operation

            ''' <summary>
            ''' Gets or sets the text associated with this diff segment.
            ''' </summary>
            Public Property Text As String

            ''' <summary>
            ''' Initializes a new instance of the <see cref="Diff"/> class.
            ''' </summary>
            ''' <param name="op">The operation associated with this diff segment.</param>
            ''' <param name="text">The text associated with this diff segment.</param>
            Public Sub New(op As Operation, text As String)
                Me.Op = op
                Me.Text = text
            End Sub
        End Class

    End Class

End Namespace