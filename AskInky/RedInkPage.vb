' Part of "Red Ink" (Ask Inky)
' Copyright (c) LawDigital AG, Switzerland. All rights reserved. For license to use see https://redink.ai.
' =============================================================================
' File: RedInkPage.vb
' Purpose: Browser pages Ask Inky serves itself (Microsoft sign-in return page on 127.0.0.1) in the Red Ink design.
' Architecture: Windows native host with embedded Office Web-based interface.
' Notes: Red Ink License and Ask Inky third-party attributions are embedded in LICENSE.txt.
' =============================================================================
' Part of Red Ink (AskInky). P18 W367 2.0.4 (DARO UID 577: "alle Seiten im Red Ink Design"): same tokens, logo
' header, light/dark and phone layout as the Helper's runtime/redink_page.py (W380). No external resource: the
' logo is an inline data URL and the page has no script.
Friend NotInheritable Class RedInkPage
    Private Sub New()
    End Sub

    Private Const Dark As String = "--bg:#0b0f14;--card:#11161d;--fg:#e8eef6;--muted:#9aa8b7;--line:#273241;--redink-red:#cf2e2e;" &
        "--redink-red-fg:#fff;--redink-field:#0b1016;--redink-field-fg:#e8eef6;--redink-field-line:#4a5768;" &
        "--link:#ff7a7a;--shadow:0 18px 48px rgba(0,0,0,.35);--ok:#5fbf7f;color-scheme:dark"
    Private Const Light As String = "--bg:#f6f7f9;--card:#ffffff;--fg:#0e1116;--muted:#555d66;--line:#d9dde3;--redink-red:#b4232d;" &
        "--redink-red-fg:#fff;--redink-field:#ffffff;--redink-field-fg:#0e1116;--redink-field-line:#8a96a3;" &
        "--link:#b4232d;--shadow:0 16px 42px rgba(26,32,44,.09);--ok:#1f7a3f;color-scheme:light"

    Friend Shared ReadOnly Css As String =
        ":root{" & Dark & "}@media (prefers-color-scheme:light){:root{" & Light & "}}" &
        "*{box-sizing:border-box}html{background:var(--bg)}" &
        "body{margin:0;background:var(--bg);color:var(--fg);font:1rem/1.55 ""Segoe UI"",system-ui,Arial,sans-serif;" &
        "min-height:100vh;min-height:100svh;display:grid;place-items:center;padding:16px}" &
        "main{width:100%;max-width:560px;margin:0;padding:24px;background:var(--card);border:1px solid var(--line);" &
        "border-radius:18px;box-shadow:var(--shadow);overflow-wrap:anywhere}" &
        ".ri-brand{display:flex;align-items:center;gap:10px;margin:-24px -24px 20px;padding:14px 24px;border-bottom:1px solid var(--line)}" &
        ".ri-brand img{height:30px;width:auto;display:block;flex:none}.ri-brand span{font-weight:700;font-size:1.1rem;letter-spacing:.01em}" &
        "h1{font-size:1.45rem;line-height:1.3;margin:0 0 12px}p{margin:12px 0}.muted{color:var(--muted);font-size:.92rem}" &
        ".ri-next{display:flex;align-items:center;gap:12px;margin:18px 0 6px;padding:14px 16px;border-radius:10px;" &
        "border:1px solid var(--redink-red);background:var(--redink-red);color:var(--redink-red-fg);font-weight:600}" &
        ".ri-next .arrow{font-size:1.4rem;line-height:1}.ri-state{display:inline-block;width:12px;height:12px;border-radius:50%;" &
        "background:var(--ok);margin-right:8px;vertical-align:middle}.ri-state.fail{background:var(--redink-red)}" &
        "@media (max-width:420px){body{padding:8px}main{padding:18px}.ri-brand{margin:-18px -18px 16px;padding:12px 18px}h1{font-size:1.3rem}}"

    Private Shared logoCache As String

    Friend Shared Function LogoDataUrl() As String
        If logoCache IsNot Nothing Then Return logoCache
        Dim value As String = ""
        Try
            Using stream As System.IO.Stream = GetType(RedInkPage).Assembly.GetManifestResourceStream("AskInky.Assets.Red_Ink_Logo_Large.png")
                If stream IsNot Nothing Then
                    Using buffer As New System.IO.MemoryStream()
                        stream.CopyTo(buffer)
                        value = "data:image/png;base64," & System.Convert.ToBase64String(buffer.ToArray())
                    End Using
                End If
            End Using
        Catch ex As System.Exception
            ClientLog.Write("page.logo.unavailable", ex)
        End Try
        logoCache = value
        Return value
    End Function

    ' Result page after the Microsoft sign-in in the browser. All texts are escaped here.
    Friend Shared Function SignInResult(language As String, success As Boolean, Optional detail As String = "") As String
        Dim enc As System.Func(Of String, String) = Function(s) System.Net.WebUtility.HtmlEncode(If(s, ""))
        Dim title As String = UiStrings.GetText(If(success, "desktop.callback.doneTitle", "desktop.callback.failedTitle"), language)
        Dim body As New System.Text.StringBuilder()
        body.Append("<h1><span class=""ri-state").Append(If(success, "", " fail")).Append(""" aria-hidden=""true""></span>").Append(enc(title)).Append("</h1>")
        If success Then
            body.Append("<p>").Append(enc(UiStrings.GetText("desktop.callback.doneText", language))).Append("</p>")
            body.Append("<div class=""ri-next"" role=""note""><span class=""arrow"" aria-hidden=""true"">&#8592;</span><span>").Append(enc(UiStrings.GetText("desktop.callback.back", language))).Append("</span></div>")
            body.Append("<p class=""muted"">").Append(enc(UiStrings.GetText("desktop.callback.close", language))).Append("</p>")
        Else
            body.Append("<p>").Append(enc(If(detail.Length > 0, detail, UiStrings.GetText("desktop.microsoftFailed", language)))).Append("</p>")
            body.Append("<p class=""muted"">").Append(enc(UiStrings.GetText("desktop.callback.retry", language))).Append("</p>")
        End If
        Dim logo As String = LogoDataUrl()
        Dim header As String = "<header class=""ri-brand"">" & If(logo.Length > 0, "<img src=""" & logo & """ alt="""">", "") & "<span>Ask Inky · Red Ink</span></header>"
        Dim lang As String = If(String.IsNullOrWhiteSpace(language), "de", language)
        Return "<!doctype html><html lang=""" & enc(lang) & """><head><meta charset=""utf-8""><meta name=""viewport"" content=""width=device-width,initial-scale=1"">" &
            "<meta name=""referrer"" content=""no-referrer""><meta name=""color-scheme"" content=""dark light""><title>" & enc(title) & " – Ask Inky</title>" &
            "<style>" & Css & "</style></head><body><main>" & header & body.ToString() & "</main></body></html>"
    End Function
End Class
