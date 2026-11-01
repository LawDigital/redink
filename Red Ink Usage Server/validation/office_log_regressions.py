"""Inspect actual add-in sources and optionally compile/run focused VB tests with .NET 8.

From the server directory: python validation/office_log_regressions.py .. --run-vb
The generated probe imports the production functions verbatim; its Office objects
are test doubles. This does not replace a VSTO build or a live Word/Outlook test.
"""
import argparse
from pathlib import Path
import re
import shutil
import subprocess
import tempfile


def extract(text, name, kind='Function', count=1):
    rx = r'^ *Private (?:Shared )?'+kind+' '+name+r'\b[^\n]*(?:\n(?! *End '+kind+r'\b)[^\n]*)*\n *End '+kind+r'\b'
    found = re.findall(rx, text, re.M | re.I)
    if len(found) != count:
        raise ValueError(f'Expected {count} definitions of {name}, found {len(found)}.')
    return '\n'.join(found)


def generate(root):
    def read(path): return (root / path).read_text('utf-8-sig')
    word='Red Ink for Word/'; outlook='Red Ink for Outlook/'; core='SharedLibrary/Code/Core/'
    w=read(word+'ThisAddIn.CommentCommands.vb'); s=read(word+'ThisAddIn.Processing.SurgicalInsert.vb'); o=read(outlook+'ThisAddIn.Commands.vb'); ocr=read(core+'Various/SharedMethods.OcrMarkdownCleaner.vb');t=read(word+'WordCommentSelectionTracker.vb')
    def require(title,condition):
        if not condition:raise AssertionError(title)
        print('OK - '+title)
    require('One top-level marker per host, outside the subagent/resume path',all(x.count('"AgentRun_"')==1 and re.search(r'If Not subAgentMode AndAlso Not resumeWorkflowFromCheckpoint Then\s*SharedLogger\.Log',x) for x in (read(word+'ThisAddIn.Processing.Tooling.vb'),read(outlook+'ThisAddin.Tooling.vb'))))
    require('Both dispatchers use the same named skill callback and real registry',all('onSkillLoaded:=' in x and 'registry.Contains(toolCall.ToolName)' in x for x in (read(word+'ThisAddIn.Processing.Tooling.ToolExecution.vb'),read(outlook+'ThisAddin.Tooling.vb'))))
    require('Both store tools log only bound catalogue name', all(any('LogKnowledgeStoreInvocation' in line and 'boundStore.Name)' in line for line in x.splitlines()) for x in (read(word+'ThisAddIn.Processing.Tooling.Sources.vb'),read(outlook+'ThisAddIn.Tooling.Sources.vb'))))
    ribbon=read(word+'Ribbon1.vb'); require('Corrected log label and designer-backed button remain wired','"BalloonMergePart_Word invoked"' in ribbon and 'Handles RI_BalloonMergePart.Click' in ribbon and 'Me.RI_BalloonMergePart = ' in read(word+'Ribbon1.Designer.vb'))
    require('Comment window guard, reset and post-await identity check', 'Application.ActiveWindow.Hwnd' not in w and 'tracker.WatchWindow(System.IntPtr.Zero)' in w and 'after.Hwnd <> hwnd.ToInt32()' in w)
    require('Expired UIA work not started; no synchronous wait added','If completion.Task.IsCompleted Then Return' in t and 'System.Threading.Tasks.Task.Delay(1200)' in t and 'SetApartmentState(System.Threading.ApartmentState.MTA)' in t)
    require('Expected Outlook cancellation no longer serializes a stack', 'If(normalCancellation, "cancelled_by_user", ex.ToString())' in o and 'If cancelled Then Return' in o)
    require('Genuine LLM timeout and cancellation remain distinct', 'Throw New System.TimeoutException(' in read(core+'LLM/SharedMethods.MainLLM.vb') and 'Throw New System.OperationCanceledException("LLM request canceled by caller."' in read(core+'LLM/SharedMethods.MainLLM.vb'))
    require('Incomplete field maps fail closed and cleanup will not target another window/story', 'Cannot safely map all fields' in s and 'window.Hwnd = entryHandle' in s and 'selected.Range.InStory(targetRange)' in s)
    require('Bad OCR reference offsets are never silently relocated', 'start > line.Text.Length - length' in ocr and 'span overlaps an accepted reference' in ocr and 'Reference record rejected; source marker retained' in ocr)
    funcs=[extract(w,'GetActiveWordDocumentWindow'),extract(o,'IsOutlookCommandCancellation'),extract(ocr,'SourceReferenceMarkerMatches')]
    for name in ('SurgicalFieldAtomSpan','SurgicalLiveCharAtom','SurgicalPositionSpan','SurgicalTextAtom'):funcs.append(extract(s,name,'Structure'))
    for name in ('BuildSurgicalFieldAtomSpans','BuildCanonicalWfldPlaceholderKey','NormalizeWordFieldCodeForDiff','BuildSurgicalVisibleAtomMap','SurgicalReadVisibleRunAtom','SurgicalIsHorizontalWhitespaceChar','SurgicalIsParagraphBreakChar','SurgicalIsWordChar','SurgicalCreateFieldAtom'):funcs.append(extract(s,name))
    funcs.extend([extract(s,'SurgicalCreateTextAtom',count=2),extract(s,'SurgicalAppendSpan','Sub')]); return '\n'.join(funcs)

STUBS='''Option Strict On
Imports System
Imports System.Collections.Generic
Imports System.Diagnostics
Imports Microsoft.VisualBasic
Namespace Microsoft.Office.Interop.Word
 Public Class Document
 End Class
 Public Class Documents
  Public Property Count As Integer
 End Class
 Public Class Windows
  Public Property Count As Integer
 End Class
 Public Class Window
 End Class
 Public Class Application
  Public Property Documents As Documents
  Public Property Windows As Windows
  Public Property Current As Window
  Public Property Reads As Integer
  Public Property Failure As System.Exception
  Public ReadOnly Property ActiveWindow As Window
   Get
    Reads += 1
    If Failure IsNot Nothing Then Throw Failure
    Return Current
   End Get
  End Property
 End Class
 Public Class Range
  Public Property Document As Document
  Public Property Start As Integer
  Public Property [End] As Integer
  Public Property Text As String
  Public Fields As New List(Of Field)()
 End Class
 Public Class Field
  Public Property Code As Range
  Public Property Result As Range
 End Class
End Namespace
Public Class Probe
'''
TEST='''
 Private Shared Function BuildSurgicalLiveCharAtoms(r As Microsoft.Office.Interop.Word.Range) As List(Of SurgicalLiveCharAtom)
  Dim atoms As New List(Of SurgicalLiveCharAtom)
  For i As Integer=0 To r.Text.Length-1
   atoms.Add(New SurgicalLiveCharAtom With {.LiveStart=r.Start+i,.LiveEnd=r.Start+i+1,.Text=r.Text.Substring(i,1)})
  Next
  Return atoms
 End Function
 Private Shared Sub Verify(ok As Boolean,label As String)
  If Not ok Then Throw New System.Exception(label)
  System.Console.WriteLine("OK - " & label)
 End Sub
 Public Shared Sub Main()
  Dim a As New Microsoft.Office.Interop.Word.Application With {.Documents=New Microsoft.Office.Interop.Word.Documents(), .Windows=New Microsoft.Office.Interop.Word.Windows()}
  Verify(GetActiveWordDocumentWindow(Nothing) Is Nothing,"no application")
  Verify(GetActiveWordDocumentWindow(a) Is Nothing AndAlso a.Reads=0,"no documents: do not touch ActiveWindow")
  a.Documents.Count=1
  Verify(GetActiveWordDocumentWindow(a) Is Nothing AndAlso a.Reads=0,"no open window")
  a.Windows.Count=1
  Verify(GetActiveWordDocumentWindow(a) Is Nothing,"null ActiveWindow")
  a.Failure=New System.Runtime.InteropServices.COMException("race", &H800A1098I)
  Verify(GetActiveWordDocumentWindow(a) Is Nothing,"last document closed between calls")
  a.Failure=New System.Runtime.InteropServices.COMException("other fault", &H80004005I)
  Dim propagated As Boolean=False
  Try
   GetActiveWordDocumentWindow(a)
  Catch ex As System.Runtime.InteropServices.COMException
   propagated=ex.ErrorCode=&H80004005I
  End Try
  Verify(propagated,"unrelated COM fault not masked")
  a.Failure=Nothing : a.Current=New Microsoft.Office.Interop.Word.Window()
  Verify(GetActiveWordDocumentWindow(a) Is a.Current,"normal active window")
  Dim c As New System.OperationCanceledException()
  Verify(IsOutlookCommandCancellation(New System.Reflection.TargetInvocationException(c)),"wrapped cancellation")
  Verify(Not IsOutlookCommandCancellation(New System.OperationCanceledException("stopped by a fault",New System.IO.IOException()),True),"fault wrapped as cancellation retains diagnostic stack")
  Verify(Not IsOutlookCommandCancellation(New System.TimeoutException("timeout",c)),"timeout is not user cancellation")
  Verify(Not IsOutlookCommandCancellation(New System.AggregateException(c,New System.IO.IOException())),"mixed aggregate is a failure")
  Verify(SourceReferenceMarkerMatches("[¹]","1") AndAlso SourceReferenceMarkerMatches("<sup>²</sup>","2"),"validated legacy superscript wrappers")
  Verify(Not SourceReferenceMarkerMatches("[12]","2") AndAlso Not SourceReferenceMarkerMatches("<sup id='x'>1</sup>","1"),"no guessed source matches")
  Dim r As New Microsoft.Office.Interop.Word.Range With {.Document=New Microsoft.Office.Interop.Word.Document(),.Start=0,.End=8,.Text="aXXXXb" & vbCrLf}
  r.Fields.Add(New Microsoft.Office.Interop.Word.Field With {.Code=New Microsoft.Office.Interop.Word.Range With {.Start=2,.Text="TEST"},.Result=New Microsoft.Office.Interop.Word.Range With {.Start=4,.End=4}})
  Dim atoms As List(Of SurgicalTextAtom)=BuildSurgicalVisibleAtomMap(r)
  Verify(atoms.Count=4,"four atoms for adjacent word, empty field, word, CRLF")
  Verify(atoms(0).KeyText="a" AndAlso atoms(1).KeyText="WFLD:TEST" AndAlso atoms(1).LiveEnd=5 AndAlso atoms(2).KeyText="b","empty field does not absorb the adjacent character")
  Verify(atoms(3).KeyText=vbCr AndAlso atoms(3).LiveStart=6 AndAlso atoms(3).LiveEnd=8,"CRLF atom covers both physical positions")
  r.Fields(0).Result=Nothing: propagated=False
  Try
   BuildSurgicalVisibleAtomMap(r)
  Catch ex As System.IO.InvalidDataException
   propagated=True
  End Try
  Verify(propagated,"incomplete field map is not used for edits")
 End Sub
End Class
'''

def main():
    p=argparse.ArgumentParser(description=__doc__);p.add_argument('solution',type=Path);p.add_argument('--run-vb',action='store_true');args=p.parse_args()
    vb=STUBS+generate(args.solution.resolve())+TEST
    if args.run_vb:
        if not shutil.which('dotnet'): p.exit(2,'Install a .NET 8 SDK to run the isolated VB fixture.\n')
        with tempfile.TemporaryDirectory() as d:
            out=Path(d);(out/'Program.vb').write_text(vb,'utf-8-sig');(out/'Probe.vbproj').write_text('<Project Sdk="Microsoft.NET.Sdk"><PropertyGroup><OutputType>Exe</OutputType><TargetFramework>net8.0</TargetFramework><RootNamespace></RootNamespace><OptionStrict>On</OptionStrict></PropertyGroup></Project>')
            subprocess.run(['dotnet','run','--project',str(out/'Probe.vbproj')],check=True)
    else:print('Production snippets extracted. VB runtime assertions NOT executed (pass --run-vb).')
if __name__=='__main__':main()
