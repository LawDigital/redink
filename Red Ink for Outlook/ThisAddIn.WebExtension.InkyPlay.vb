' Part of "Red Ink for Outlook"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved.
' For license to use see https://redink.ai.
'
' =============================================================================
' File: ThisAddIn.WebExtension.InkyPlay.vb
' Purpose:
'   Hosts the InkyPlay browser mini-game suite under `/inky/play`, including
'   AI-generated phishing/safe email scenarios and persisted high scores.
'
' Features:
'  - Four arcade-style awareness games:
'      1) Inbox Invaders
'      2) Phish Pong
'      3) Data Defender
'      4) Risk Stack
'  - Optional runtime scenario generation via LLM (`inkyplay_generate`).
'  - Local high-score persistence via `My.Settings.InkyPlay_HighScores`.
'
' Routes / Commands:
'  - GET  `/inky/play` → full single-file HTML/CSS/JS game UI
'  - API commands routed via the authenticated `/inky/api` surface with CSRF protection:
'      * `inkyplay_generate`
'      * `inkyplay_gethighscores`
'      * `inkyplay_savescore`
'      * `inkyplay_clearhighscores`
'
' Architecture:
'  - `BuildInkyPlayHtmlPage` returns a self-contained game page with all assets
'    and logic embedded inline.
'  - `ProcessInkyPlayCommand` acts as command dispatcher and JSON response producer.
'  - Scoreboard model is serialized/deserialized as JSON and capped during save.
'  - AI model selection prefers alternate "Play" model when configured, then
'    falls back to primary model and finally to internal fallback email dataset.
' =============================================================================

Option Explicit On
Option Strict Off

Imports System.Threading
Imports System.Threading.Tasks
Imports SharedLibrary.SharedLibrary.SharedMethods

Partial Public Class ThisAddIn

    ' ─── Routes ───────────────────────────────────────────────────────────────
    Private Const InkyPlayRoute As String = "/inky/play"   ' GET → game HTML

    ' ─── High-Score Persistence ───────────────────────────────────────────────

    ''' <summary>
    ''' Represents a single high-score entry for one game run.
    ''' </summary>
    <Serializable>
    Private Class InkyPlayHighScore
        Public PlayerName As String = ""
        Public Game As String = ""
        Public Score As Integer = 0
        Public Wave As Integer = 0
        Public Utc As DateTime = DateTime.UtcNow
    End Class

    ''' <summary>
    ''' Container for persisted InkyPlay high-score entries.
    ''' </summary>
    <Serializable>
    Private Class InkyPlayScoreBoard
        Public Scores As System.Collections.Generic.List(Of InkyPlayHighScore) =
            New System.Collections.Generic.List(Of InkyPlayHighScore)()
    End Class


    ''' <summary>
    ''' Loads persisted InkyPlay high scores from settings.
    ''' </summary>
    ''' <returns>
    ''' A valid scoreboard object; returns an empty scoreboard on missing/invalid data.
    ''' </returns>
    Private Function LoadPlayScores() As InkyPlayScoreBoard
        Try
            Dim raw As String = ""
            Try
                raw = DirectCast(My.Settings.[GetType]().GetProperty("InkyPlay_HighScores").GetValue(My.Settings, Nothing), String)
            Catch
                raw = ""
            End Try
            If String.IsNullOrWhiteSpace(raw) Then Return New InkyPlayScoreBoard()
            Dim board = Newtonsoft.Json.JsonConvert.DeserializeObject(Of InkyPlayScoreBoard)(raw)
            If board Is Nothing Then board = New InkyPlayScoreBoard()
            Return board
        Catch
            Return New InkyPlayScoreBoard()
        End Try
    End Function

    ''' <summary>
    ''' Persists the InkyPlay high-score board to settings as JSON.
    ''' </summary>
    Private Sub SavePlayScores(board As InkyPlayScoreBoard)
        Try
            Dim json = Newtonsoft.Json.JsonConvert.SerializeObject(board)
            Try
                My.Settings.[GetType]().GetProperty("InkyPlay_HighScores").SetValue(My.Settings, json, Nothing)
                My.Settings.Save()
            Catch
            End Try
        Catch
        End Try
    End Sub

    ' ─── API Command Dispatcher ───────────────────────────────────────────────

    ''' <summary>
    ''' Handles all `inkyplay_*` commands routed from ProcessRequestInAddIn.
    ''' Returns a JSON response string (prefixed with CT:json), or Nothing if
    ''' the command is not recognized.
    ''' </summary>
    Friend Async Function ProcessInkyPlayCommand(
        cmd As String,
        j As Newtonsoft.Json.Linq.JObject
    ) As Task(Of String)

        Select Case cmd

            Case "inkyplay_generate"
                ' Ask LLM to generate a batch of emails for the games
                Try
                    Dim count As Integer = 10
                    Try : count = CInt(j("Count")) : Catch : count = 10 : End Try
                    If count < 5 Then count = 5
                    If count > 20 Then count = 20

                    Dim difficulty As String = If(j("Difficulty")?.ToString(), "medium")
                    Dim seed As String = System.Guid.NewGuid().ToString("N")

                    Dim userDomain As String = Environment.GetEnvironmentVariable("USERDNSDOMAIN")
                    If String.IsNullOrWhiteSpace(userDomain) Then userDomain = "yourcompany.com"

                    Dim sysPrompt As String =
                        "You are a security awareness training content generator. " &
                        "Create a JSON array of email objects for a phishing awareness game. " &
                        "Each object MUST have exactly these fields: " &
                        """from"" (sender email), ""subject"" (subject line), ""preview"" (first 1-2 lines of body), " &
                        """isPhishing"" (boolean), ""redFlag"" (short explanation why it is or isn't phishing). " &
                        "Ensure every email is distinct. Vary names, domains, industries, regions, brands, and writing styles. " &
                        "Do not reuse examples or patterns across the array. " &
                        "Phishing should feel realistic; legitimate should feel routine and credible. " &
                        $"For legitimate (non-phishing) emails, use the domain ""{userDomain}"" as the sender domain to make them appear as internal company emails. " &
                        "Keep subject under 60 characters and preview under 90 characters. " &
                        $"Difficulty level: {difficulty}. " &
                        "Return ONLY the JSON array, no markdown fences, no explanation."

                    Dim userPrompt As String =
                        $"Generate exactly {count} email objects as a JSON array. Seed: {seed}."

                    ' Dynamic model: try to use a "Play" model from alternate models
                    Dim useSecondForCall As Boolean = False
                    Dim playModelSwitched As Boolean = False
                    Try
                        If Not String.IsNullOrWhiteSpace(_context.INI_AlternateModelPath) Then
                            If GetSpecialTaskModel(_context, _context.INI_AlternateModelPath, "Play") Then
                                useSecondForCall = True
                                playModelSwitched = True
                            End If
                        End If
                    Catch
                        ' Ignore model selection errors — fall back to primary
                    End Try

                    Dim llmOut As String
                    Try
                        llmOut = Await RunLlmAsync(sysPrompt, userPrompt, useSecondForCall, False).ConfigureAwait(False)
                    Finally
                        ' Restore model if temporarily switched
                        If playModelSwitched AndAlso originalConfigLoaded Then
                            Try
                                RestoreDefaults(_context, originalConfig)
                                originalConfigLoaded = False
                            Catch
                            End Try
                        End If
                    End Try

                    ' Try to extract JSON array from response
                    If llmOut IsNot Nothing Then
                        llmOut = llmOut.Trim()
                        ' Strip markdown fences if present
                        If llmOut.StartsWith("```") Then
                            Dim firstNl = llmOut.IndexOf(vbLf)
                            If firstNl > 0 Then llmOut = llmOut.Substring(firstNl + 1)
                            If llmOut.EndsWith("```") Then llmOut = llmOut.Substring(0, llmOut.Length - 3).Trim()
                        End If
                    End If

                    Return JsonOk(New With {.ok = True, .emails = Newtonsoft.Json.Linq.JToken.Parse(If(llmOut, "[]"))})
                Catch ex As Exception
                    Return JsonErr("Failed to generate emails: " & ex.Message)
                End Try

            Case "inkyplay_gethighscores"
                Try
                    Dim board = LoadPlayScores()
                    Dim game As String = If(j("Game")?.ToString(), "")
                    Dim filtered = board.Scores
                    If Not String.IsNullOrWhiteSpace(game) Then
                        filtered = board.Scores.Where(
                            Function(s) String.Equals(s.Game, game, StringComparison.OrdinalIgnoreCase)).ToList()
                    End If
                    filtered = filtered.OrderByDescending(Function(s) s.Score).Take(10).ToList()
                    Return JsonOk(New With {.ok = True, .scores = filtered})
                Catch ex As Exception
                    Return JsonErr("Failed to load scores: " & ex.Message)
                End Try

            Case "inkyplay_savescore"
                Try
                    Dim board = LoadPlayScores()
                    Dim entry As New InkyPlayHighScore With {
                        .PlayerName = If(j("PlayerName")?.ToString(), "Player"),
                        .Game = If(j("Game")?.ToString(), "unknown"),
                        .Score = CInt(If(j("Score"), 0)),
                        .Wave = CInt(If(j("Wave"), 0)),
                        .Utc = DateTime.UtcNow
                    }
                    board.Scores.Add(entry)
                    ' Keep only top 50 scores total
                    If board.Scores.Count > 50 Then
                        board.Scores = board.Scores.OrderByDescending(Function(s) s.Score).Take(50).ToList()
                    End If
                    SavePlayScores(board)
                    Return JsonOk(New With {.ok = True, .saved = True})
                Catch ex As Exception
                    Return JsonErr("Failed to save score: " & ex.Message)
                End Try

            Case "inkyplay_clearhighscores"
                Try
                    SavePlayScores(New InkyPlayScoreBoard())
                    Return JsonOk(New With {.ok = True, .cleared = True})
                Catch ex As Exception
                    Return JsonErr("Failed to clear scores: " & ex.Message)
                End Try

            Case Else
                Return Nothing  ' Not an InkyPlay command
        End Select
    End Function

    ' ─── HTML Page Builder ────────────────────────────────────────────────────

    ''' <summary>
    ''' Generates the full single-file HTML page for the InkyPlay game suite.
    ''' All CSS, JS, and game logic are inline (no external assets).
    ''' </summary>
    Friend Function BuildInkyPlayHtmlPage(ByVal csrfToken As System.String) As System.String
        Dim logoUrl As String = GetLogoDataUrl()
        Dim brandName As String = If(Not String.IsNullOrWhiteSpace(AN), AN, "Red Ink")

        Dim sb As New System.Text.StringBuilder(262144) ' pre-allocate ~256KB

        sb.AppendLine("<!doctype html>")
        sb.AppendLine("<html lang=""en""><head><meta charset=""utf-8"">")
        sb.AppendLine("<meta name=""viewport"" content=""width=device-width,initial-scale=1"">")
        sb.AppendLine("<title>" & System.Net.WebUtility.HtmlEncode(brandName) & " — InkyPlay</title>")
        If Not String.IsNullOrWhiteSpace(logoUrl) Then
            sb.AppendLine("<link rel=""icon"" type=""image/png"" href=""" & System.Net.WebUtility.HtmlEncode(logoUrl) & """>")
        End If

        ' ── CSS ───────────────────────────────────────────────────────────
        sb.AppendLine("<style>")
        sb.AppendLine("*{margin:0;padding:0;box-sizing:border-box}")
        sb.AppendLine(":root{--bg:#0b0f14;--card:#11161d;--fg:#e8eef6;--muted:#9aa8b7;--accent:#3b82f6;--danger:#ef4444;--success:#22c55e;--warning:#eab308;--border:#1b2430}")
        sb.AppendLine("html,body{height:100%;font-family:system-ui,Segoe UI,sans-serif;background:var(--bg);color:var(--fg)}")
        sb.AppendLine("a{color:var(--accent)}")
        sb.AppendLine("canvas{display:block;background:#0a0e13;border:1px solid var(--border);border-radius:8px;width:800px;height:500px}")
        sb.AppendLine(".app{display:flex;flex-direction:column;height:100%;overflow:hidden}")
        sb.AppendLine(".topbar{display:flex;align-items:center;gap:.6rem;padding:.6rem 1rem;border-bottom:1px solid var(--border);background:var(--card);flex-shrink:0}")
        sb.AppendLine(".topbar img{width:22px;height:22px;border-radius:5px}")
        sb.AppendLine(".topbar .brand{font-weight:700}")
        sb.AppendLine(".topbar .sub{color:var(--muted);font-size:.85rem}")
        sb.AppendLine(".topbar .spacer{flex:1}")
        sb.AppendLine("button{background:var(--card);color:var(--fg);border:1px solid var(--border);border-radius:.5rem;padding:.45rem .8rem;cursor:pointer;font:inherit;transition:filter .15s}")
        sb.AppendLine("button:hover{filter:brightness(1.1)}")
        sb.AppendLine("button:disabled{opacity:.4;cursor:not-allowed}")
        sb.AppendLine("button.primary{background:var(--accent);border-color:var(--accent);color:#fff}")
        sb.AppendLine("button.danger{background:var(--danger);border-color:var(--danger);color:#fff}")
        sb.AppendLine(".main{flex:1;overflow:auto;display:flex;flex-direction:column;align-items:center;padding:1.5rem}")

        ' Menu screen
        sb.AppendLine("#menuScreen{max-width:700px;width:100%;text-align:center}")
        sb.AppendLine("#menuScreen h1{font-size:2rem;margin-bottom:.3rem}")
        sb.AppendLine("#menuScreen .tagline{color:var(--muted);margin-bottom:2rem;font-size:.95rem}")
        sb.AppendLine(".gameCards{display:grid;grid-template-columns:repeat(auto-fit,minmax(280px,1fr));gap:1rem;margin-bottom:2rem}")
        sb.AppendLine(".gcard{background:var(--card);border:1px solid var(--border);border-radius:1rem;padding:1.5rem;text-align:left;cursor:pointer;transition:border-color .2s,transform .15s}")
        sb.AppendLine(".gcard:hover{border-color:var(--accent);transform:translateY(-2px)}")
        sb.AppendLine(".gcard h3{margin-bottom:.3rem}")
        sb.AppendLine(".gcard .gdesc{color:var(--muted);font-size:.85rem;line-height:1.4}")
        sb.AppendLine(".gcard .gicon{font-size:2rem;margin-bottom:.5rem}")

        ' Game screen
        sb.AppendLine("#gameScreen{display:none;width:100%;max-width:900px;text-align:center}")
        sb.AppendLine("#gameScreen h2{margin-bottom:.5rem}")
        sb.AppendLine("#gameScreen .gameInfo{display:flex;justify-content:center;gap:2rem;margin-bottom:.8rem;font-size:.9rem}")
        sb.AppendLine("#gameScreen .gameInfo span{color:var(--muted)}")
        sb.AppendLine("#gameScreen .gameInfo b{color:var(--fg)}")
        sb.AppendLine(".canvasWrap{display:flex;justify-content:center;margin-bottom:.8rem}")
        sb.AppendLine(".controls{display:flex;justify-content:center;gap:.5rem;flex-wrap:wrap;margin-bottom:.5rem}")
        sb.AppendLine("#redFlagBox{position:absolute;left:10px;right:10px;top:110px;background:var(--card);border:1px solid var(--border);border-radius:.6rem;padding:.5rem .8rem;font-size:.85rem;color:var(--muted);min-height:2rem;max-width:780px;text-align:left;pointer-events:none;box-shadow:0 6px 14px rgba(0,0,0,.35)}")

        ' Scores screen
        sb.AppendLine("#scoresScreen{display:none;max-width:600px;width:100%}")
        sb.AppendLine("#scoresScreen h2{margin-bottom:1rem;text-align:center}")
        sb.AppendLine(".scoreTable{width:100%;border-collapse:collapse}")
        sb.AppendLine(".scoreTable th,.scoreTable td{padding:.5rem .7rem;border-bottom:1px solid var(--border);text-align:left}")
        sb.AppendLine(".scoreTable th{color:var(--muted);font-size:.8rem;text-transform:uppercase;letter-spacing:.5px}")

        ' Loading overlay
        sb.AppendLine(".loadingOverlay{position:fixed;inset:0;background:rgba(0,0,0,.7);display:flex;align-items:center;justify-content:center;z-index:100;font-size:1.2rem}")
        sb.AppendLine(".loadingOverlay .spinner{width:32px;height:32px;border:3px solid var(--border);border-top-color:var(--accent);border-radius:50%;animation:spin .8s linear infinite;margin-right:1rem}")
        sb.AppendLine("@keyframes spin{to{transform:rotate(360deg)}}")

        ' Game over overlay
        sb.AppendLine(".goOverlay{position:absolute;inset:0;background:rgba(0,0,0,.8);display:flex;flex-direction:column;align-items:center;justify-content:center;z-index:10;border-radius:8px}")
        sb.AppendLine(".goOverlay h2{font-size:2rem;margin-bottom:.5rem}")
        sb.AppendLine(".goOverlay .finalScore{font-size:1.3rem;color:var(--accent);margin-bottom:1rem}")
        sb.AppendLine(".goOverlay input{background:var(--card);color:var(--fg);border:1px solid var(--border);border-radius:.4rem;padding:.4rem .6rem;font:inherit;width:200px;margin-bottom:.8rem;text-align:center}")
        sb.AppendLine(".goOverlay .goButtons{display:flex;gap:.5rem}")

        sb.AppendLine("</style></head><body>")

        ' ── App Shell ─────────────────────────────────────────────────────
        sb.AppendLine("<div class=""app"">")
        sb.AppendLine("<div class=""topbar"">")
        If Not String.IsNullOrWhiteSpace(logoUrl) Then
            sb.AppendLine("<img src=""" & System.Net.WebUtility.HtmlEncode(logoUrl) & """ alt=""logo"">")
        End If
        sb.AppendLine("<span class=""brand"">" & System.Net.WebUtility.HtmlEncode(brandName) & "</span>")
        sb.AppendLine("<span class=""sub"">InkyPlay — Security Awareness Arcade</span>")
        sb.AppendLine("<span class=""spacer""></span>")
        sb.AppendLine("<button id=""btnScores"">🏆 High Scores</button>")
        sb.AppendLine("<button id=""btnMenu"">☰ Menu</button>")
        sb.AppendLine("<button onclick=""window.open('/inky','_self')"">💬 Chat</button>")
        sb.AppendLine("</div>")

        sb.AppendLine("<div class=""main"">")

        ' ── Menu Screen ───────────────────────────────────────────────────
        sb.AppendLine("<div id=""menuScreen"">")
        sb.AppendLine("<h1>🎮 InkyPlay</h1>")
        sb.AppendLine("<p class=""tagline"">Learn to spot phishing — the fun way. 2–5 minute sessions powered by AI.</p>")
        sb.AppendLine("<div class=""gameCards"">")

        sb.AppendLine("<div class=""gcard"" onclick=""startGame('invaders')"">")
        sb.AppendLine("<div class=""gicon"">👾</div><h3>Inbox Invaders</h3>")
        sb.AppendLine("<p class=""gdesc"">Read the email at the top. Shoot the invader if it's phishing, let it pass if it's safe!</p></div>")

        sb.AppendLine("<div class=""gcard"" onclick=""startGame('pong')"">")
        sb.AppendLine("<div class=""gicon"">🏓</div><h3>Phish Pong</h3>")
        sb.AppendLine("<p class=""gdesc"">Bounce the falling email to the Left (Inbox) if safe, or Right (Sandbox) if phishing.</p></div>")

        sb.AppendLine("<div class=""gcard"" onclick=""startGame('defender')"">")
        sb.AppendLine("<div class=""gicon"">🟡</div><h3>Data Defender</h3>")
        sb.AppendLine("<p class=""gdesc"">Navigate the maze step-by-step. Collect emails and classify them correctly to survive.</p></div>")

        sb.AppendLine("<div class=""gcard"" onclick=""startGame('stack')"">")
        sb.AppendLine("<div class=""gicon"">🧱</div><h3>Risk Stack</h3>")
        sb.AppendLine("<p class=""gdesc"">Sort the current email into Safe, Verify, or Block using your arrow keys.</p></div>")

        sb.AppendLine("</div>") ' gameCards

        sb.AppendLine("<p style=""color:var(--muted);font-size:.8rem"">Each game uses AI to generate fresh phishing scenarios. Tip: Look for urgency, domain mismatches, and suspicious links.</p>")
        sb.AppendLine("</div>") ' menuScreen

        ' ── Game Screen ───────────────────────────────────────────────────
        sb.AppendLine("<div id=""gameScreen"">")
        sb.AppendLine("<h2 id=""gameTitle""></h2>")
        sb.AppendLine("<div class=""gameInfo""><span>Score: <b id=""scoreDisplay"">0</b></span><span>Wave: <b id=""waveDisplay"">1</b></span><span>Lives: <b id=""livesDisplay"">3</b></span><span>Email Source: <b id=""emailSource"">AI</b></span></div>")
        sb.AppendLine("<div class=""canvasWrap"" style=""position:relative"">")
        sb.AppendLine("<canvas id=""gc"" width=""800"" height=""500""></canvas>")
        sb.AppendLine("<div id=""redFlagBox""></div>")
        sb.AppendLine("</div>")
        sb.AppendLine("<div class=""controls"">")
        sb.AppendLine("<button onclick=""pauseGame()"">⏸ Pause</button>")
        sb.AppendLine("<button onclick=""showMenu()"">🔙 Back to Menu</button>")
        sb.AppendLine("</div>")
        sb.AppendLine("</div>") ' gameScreen

        ' ── Scores Screen ─────────────────────────────────────────────────
        sb.AppendLine("<div id=""scoresScreen"">")
        sb.AppendLine("<h2>🏆 High Scores</h2>")
        sb.AppendLine("<div style=""display:flex;justify-content:center;gap:.5rem;margin-bottom:1rem"">")
        sb.AppendLine("<button class=""primary"" onclick=""loadScores()"">Refresh</button>")
        sb.AppendLine("<button class=""danger"" onclick=""clearScores()"">Clear All</button>")
        sb.AppendLine("<button onclick=""showMenu()"">🔙 Menu</button>")
        sb.AppendLine("</div>")
        sb.AppendLine("<table class=""scoreTable""><thead><tr><th>#</th><th>Player</th><th>Game</th><th>Score</th><th>Wave</th><th>Date</th></tr></thead><tbody id=""scoreBody""></tbody></table>")
        sb.AppendLine("</div>") ' scoresScreen

        sb.AppendLine("</div>") ' main
        sb.AppendLine("</div>") ' app

        ' Loading overlay (hidden by default)
        sb.AppendLine("<div id=""loadingOverlay"" class=""loadingOverlay"" style=""display:none""><div class=""spinner""></div><span id=""loadingText"">Generating emails with AI…</span></div>")

        ' ── JavaScript ────────────────────────────────────────────────────
        sb.AppendLine("<script>")

        ' ═══════════════════════════════════════════════════════════════════
        ' CORE API & STATE
        ' ═══════════════════════════════════════════════════════════════════
        sb.AppendLine("'use strict';")
        sb.AppendLine("const __redInkCsrf=" & Newtonsoft.Json.JsonConvert.SerializeObject(If(csrfToken, System.String.Empty)) & ";")
        sb.AppendLine("const api=async(cmd,data={})=>{try{const r=await fetch('/inky/api',{method:'POST',headers:{'Content-Type':'application/json','X-RedInk-CSRF':__redInkCsrf},body:JSON.stringify(Object.assign({Command:cmd},data))});const t=await r.text();try{return JSON.parse(t)}catch{return{ok:false,error:t}}}catch(e){return{ok:false,error:e.message}}};")

        sb.AppendLine("const canvas=document.getElementById('gc');const ctx=canvas.getContext('2d');")
        sb.AppendLine("var dpr=window.devicePixelRatio||1;canvas.width=800*dpr;canvas.height=500*dpr;ctx.scale(dpr,dpr);")
        sb.AppendLine("const menuScreen=document.getElementById('menuScreen');const gameScreen=document.getElementById('gameScreen');const scoresScreen=document.getElementById('scoresScreen');")
        sb.AppendLine("const scoreDisp=document.getElementById('scoreDisplay');const waveDisp=document.getElementById('waveDisplay');const livesDisp=document.getElementById('livesDisplay');")
        sb.AppendLine("const emailSourceDisp=document.getElementById('emailSource');")
        sb.AppendLine("const redFlagBox=document.getElementById('redFlagBox');const loadingOverlay=document.getElementById('loadingOverlay');const loadingText=document.getElementById('loadingText');")

        sb.AppendLine("let currentGame=null,animFrame=0,paused=false,gameOver=false;")
        sb.AppendLine("let score=0,wave=1,lives=3;")
        sb.AppendLine("let emails=[],keys={};")
        sb.AppendLine("let emailSource='Fallback',emailRefreshInProgress=false;")

        ' Fallback emails (used if LLM call fails)
        sb.AppendLine("const fallbackEmails=[")
        sb.AppendLine("{from:'it-support@c0mpany.net',subject:'Urgent: Password Reset Required',preview:'Your password will expire in 1 hour. Click here to reset immediately.',isPhishing:true,redFlag:'Misspelled domain (c0mpany with zero), artificial urgency'},")
        sb.AppendLine("{from:'hr@company.com',subject:'Q1 All-Hands Meeting',preview:'Please join us Thursday at 2pm in the main conference room for our quarterly update.',isPhishing:false,redFlag:'Legitimate internal meeting invite with no suspicious links'},")
        sb.AppendLine("{from:'ceo@cornpany.com',subject:'Wire Transfer Needed ASAP',preview:'I need you to process an urgent wire transfer. This is confidential.',isPhishing:true,redFlag:'CEO impersonation (cornpany vs company), urgency, confidentiality pressure'},")
        sb.AppendLine("{from:'noreply@github.com',subject:'[GitHub] Security alert',preview:'We noticed a new sign-in to your account from Chrome on Windows.',isPhishing:false,redFlag:'Legitimate GitHub security notification'},")
        sb.AppendLine("{from:'shipping@fedex-delivery.info',subject:'Package Delivery Failed',preview:'Your package could not be delivered. Click to reschedule.',isPhishing:true,redFlag:'Suspicious domain (fedex-delivery.info instead of fedex.com)'},")
        sb.AppendLine("{from:'newsletter@medium.com',subject:'Your Daily Digest',preview:'Top stories in Technology, Science, and Programming today.',isPhishing:false,redFlag:'Legitimate newsletter from known platform'},")
        sb.AppendLine("{from:'billing@micros0ft-online.com',subject:'Invoice #INV-2847 Payment Due',preview:'Your subscription payment of $299.99 has failed. Update payment method now.',isPhishing:true,redFlag:'Homoglyph domain (micros0ft), unexpected invoice amount'},")
        sb.AppendLine("{from:'calendar@google.com',subject:'Invitation: Project Sync',preview:'You have been invited to a meeting on March 15 at 10:00 AM.',isPhishing:false,redFlag:'Standard Google Calendar invite'},")
        sb.AppendLine("{from:'security@paypa1.com',subject:'Unusual Activity Detected',preview:'We detected unusual activity on your account. Verify your identity immediately.',isPhishing:true,redFlag:'Homoglyph in domain (paypa1 with number 1), urgency'},")
        sb.AppendLine("{from:'team@slack.com',subject:'New message in #general',preview:'John posted: ""Has anyone reviewed the latest PR?""',isPhishing:false,redFlag:'Legitimate Slack notification'},")
        sb.AppendLine("{from:'admin@lT-helpdesk.org',subject:'MFA Token Expiring',preview:'Your multi-factor authentication token expires today. Re-enroll now.',isPhishing:true,redFlag:'Suspicious domain with capital I/lowercase L trick, urgency'},")
        sb.AppendLine("{from:'no-reply@linkedin.com',subject:'You appeared in 5 searches',preview:'See who is looking at your profile this week.',isPhishing:false,redFlag:'Standard LinkedIn engagement email'},")
        sb.AppendLine("{from:'helpdesk@company-vpn.xyz',subject:'VPN Certificate Renewal',preview:'Your VPN certificate will expire tomorrow. Download the new certificate.',isPhishing:true,redFlag:'Unusual TLD (.xyz), asks to download files'},")
        sb.AppendLine("{from:'noreply@zoom.us',subject:'Cloud Recording Available',preview:'Your recorded meeting ""Weekly Standup"" is now available.',isPhishing:false,redFlag:'Legitimate Zoom recording notification'},")
        sb.AppendLine("{from:'prize@w1nner-notification.com',subject:'Congratulations! You Won!',preview:'You have been selected as the lucky winner of $10,000!',isPhishing:true,redFlag:'Too-good-to-be-true prize, suspicious domain'},")
        sb.AppendLine("{from:'jira@atlassian.net',subject:'[PROJ-142] Bug fix merged',preview:'The pull request for PROJ-142 has been merged to main.',isPhishing:false,redFlag:'Legitimate Jira/Atlassian notification'},")
        sb.AppendLine("{from:'tax@irs-refund.us',subject:'Tax Refund Notification',preview:'You are eligible for a tax refund of $3,247.00. Submit your details.',isPhishing:true,redFlag:'IRS does not contact via email, suspicious domain'},")
        sb.AppendLine("{from:'donotreply@amazon.com',subject:'Your order has shipped',preview:'Your order #112-4839271 is on its way. Expected delivery: Friday.',isPhishing:false,redFlag:'Standard Amazon shipping notification'},")
        sb.AppendLine("{from:'scan@office-printer.net',subject:'Scanned Document',preview:'The attached document was scanned and sent from Printer-03.',isPhishing:true,redFlag:'Generic printer scan email often carries malware attachments'},")
        sb.AppendLine("{from:'updates@notion.so',subject:'What is new in Notion',preview:'Check out our latest features: databases, API improvements, and more.',isPhishing:false,redFlag:'Legitimate product update from Notion'},")
        sb.AppendLine("];")

        ' ═══════════════════════════════════════════════════════════════════
        ' SCREEN NAVIGATION
        ' ═══════════════════════════════════════════════════════════════════
        sb.AppendLine("function showMenu(){cancelAnimationFrame(animFrame);gameOver=true;menuScreen.style.display='block';gameScreen.style.display='none';scoresScreen.style.display='none';currentGame=null;}")
        sb.AppendLine("function showGame(){menuScreen.style.display='none';gameScreen.style.display='block';scoresScreen.style.display='none';}")
        sb.AppendLine("function showScores(){menuScreen.style.display='none';gameScreen.style.display='none';scoresScreen.style.display='block';loadScores();}")
        sb.AppendLine("document.getElementById('btnMenu').onclick=showMenu;")
        sb.AppendLine("document.getElementById('btnScores').onclick=showScores;")
        sb.AppendLine("function pauseGame(){paused=!paused;}")

        ' ═══════════════════════════════════════════════════════════════════
        ' EMAIL LOADING (LLM or fallback)
        ' ═══════════════════════════════════════════════════════════════════
        sb.AppendLine("function updateEmailSource(){if(emailSourceDisp)emailSourceDisp.textContent=emailSource;}")
        sb.AppendLine("async function loadEmails(difficulty,count){")
        sb.AppendLine("  if(!difficulty)difficulty='medium';if(!count)count=10;")
        sb.AppendLine("  loadingOverlay.style.display='flex';loadingText.textContent='Generating emails with AI\u2026';")
        sb.AppendLine("  try{")
        sb.AppendLine("    const r=await api('inkyplay_generate',{Difficulty:difficulty,Count:count});")
        sb.AppendLine("    if(r&&r.ok===true&&Array.isArray(r.emails)&&r.emails.length>=5){emails=r.emails;emailSource='AI';}")
        sb.AppendLine("    else{emails=fallbackEmails.slice();emailSource='Fallback';}")
        sb.AppendLine("  }catch(e){emails=fallbackEmails.slice();emailSource='Fallback';}")
        sb.AppendLine("  finally{loadingOverlay.style.display='none';updateEmailSource();}")
        sb.AppendLine("  if(!emails||!emails.length){emails=fallbackEmails.slice();emailSource='Fallback';updateEmailSource();}")
        sb.AppendLine("  for(let i=emails.length-1;i>0;i--){const j=Math.floor(Math.random()*(i+1));[emails[i],emails[j]]=[emails[j],emails[i]];}")
        sb.AppendLine("}")

        sb.AppendLine("async function refreshEmailsForWave(){")
        sb.AppendLine("  if(emailRefreshInProgress) return;")
        sb.AppendLine("  emailRefreshInProgress=true;")
        sb.AppendLine("  const diff=wave<=2?'easy':wave<=4?'medium':'hard';")
        sb.AppendLine("  await loadEmails(diff);")
        sb.AppendLine("  emails=normalizeEmails(emails);")
        sb.AppendLine("  if(!emails.length){emails=normalizeEmails(fallbackEmails);emailSource='Fallback';updateEmailSource();}")
        sb.AppendLine("  emailRefreshInProgress=false;")
        sb.AppendLine("}")

        sb.AppendLine("function queueEmailRefresh(onReady){")
        sb.AppendLine("  if(emailRefreshInProgress) return;")
        sb.AppendLine("  paused=true;")
        sb.AppendLine("  refreshEmailsForWave().then(()=>{if(typeof onReady==='function')onReady();paused=false;});")
        sb.AppendLine("}")

        ' ═══════════════════════════════════════════════════════════════════
        ' HIGH SCORES
        ' ═══════════════════════════════════════════════════════════════════
        sb.AppendLine("async function loadScores(){const r=await api('inkyplay_gethighscores',{});const tb=document.getElementById('scoreBody');tb.innerHTML='';if(!r.ok||!r.scores||!r.scores.length){tb.innerHTML='<tr><td colspan=""6"" style=""text-align:center;color:var(--muted)"">No scores yet</td></tr>';return;}r.scores.forEach((s,i)=>{const d=s.utc?new Date(s.utc).toLocaleDateString():'';tb.innerHTML+=`<tr><td>${i+1}</td><td>${esc(s.playerName)}</td><td>${esc(s.game)}</td><td>${s.score}</td><td>${s.wave||'-'}</td><td>${d}</td></tr>`;});}")
        sb.AppendLine("async function clearScores(){if(!confirm('Clear all high scores?'))return;await api('inkyplay_clearhighscores');loadScores();}")
        sb.AppendLine("async function saveScore(game,sc,w){const name=prompt('Enter your name:','Player');if(!name)return;await api('inkyplay_savescore',{PlayerName:name,Game:game,Score:sc,Wave:w});}")

        ' HTML-escape helper
        sb.AppendLine("function esc(s){if(!s)return'';const d=document.createElement('div');d.textContent=s;return d.innerHTML;}")

        ' ═══════════════════════════════════════════════════════════════════
        ' KEY HANDLING
        ' ═══════════════════════════════════════════════════════════════════
        sb.AppendLine("document.addEventListener('keydown',e=>{keys[e.key]=true;if(['ArrowLeft','ArrowRight','ArrowUp','ArrowDown',' '].includes(e.key))e.preventDefault();});")
        sb.AppendLine("document.addEventListener('keyup',e=>{keys[e.key]=false;});")

        ' ═══════════════════════════════════════════════════════════════════
        ' GAME OVER OVERLAY
        ' ═══════════════════════════════════════════════════════════════════
        sb.AppendLine("function showGameOver(gameName,finalScore,finalWave){")
        sb.AppendLine("  const wrap=canvas.parentElement;")
        sb.AppendLine("  let ov=wrap.querySelector('.goOverlay');if(ov)ov.remove();")
        sb.AppendLine("  ov=document.createElement('div');ov.className='goOverlay';")
        sb.AppendLine("  ov.innerHTML=`<h2>Game Over!</h2><div class=""finalScore"">Score: ${finalScore}  |  Wave: ${finalWave}</div><input id=""goName"" placeholder=""Your name"" value=""Player"" maxlength=""20""><div class=""goButtons""><button class=""primary"" id=""goSave"">Save Score</button><button id=""goRetry"">Play Again</button><button id=""goMenu"">Menu</button></div>`;")
        sb.AppendLine("  wrap.appendChild(ov);")
        sb.AppendLine("  ov.querySelector('#goSave').onclick=async()=>{const n=ov.querySelector('#goName').value||'Player';await api('inkyplay_savescore',{PlayerName:n,Game:gameName,Score:finalScore,Wave:finalWave});ov.querySelector('#goSave').disabled=true;ov.querySelector('#goSave').textContent='Saved!';};")
        sb.AppendLine("  ov.querySelector('#goRetry').onclick=()=>{ov.remove();startGame(currentGame);};")
        sb.AppendLine("  ov.querySelector('#goMenu').onclick=()=>{ov.remove();showMenu();};")
        sb.AppendLine("}")

        ' Show red flag explanation
        sb.AppendLine("function showRedFlag(email,correct){const pre=correct?'✅ Correct!':'❌ Wrong!';redFlagBox.innerHTML=`<b>${pre}</b> <b>${esc(email.subject)}</b> from ${esc(email.from)}<br>${email.isPhishing?'🚩 Phishing':'✅ Legitimate'}: ${esc(email.redFlag)}`;}")

        ' ═══════════════════════════════════════════════════════════════════
        ' POLYFILL: roundRect for older browser engines (IE/Edge legacy)
        ' ═══════════════════════════════════════════════════════════════════
        sb.AppendLine("if(!CanvasRenderingContext2D.prototype.roundRect){CanvasRenderingContext2D.prototype.roundRect=function(x,y,w,h,r){var tl,tr,br,bl;if(typeof r==='number'){tl=r;tr=r;br=r;bl=r;}else if(Array.isArray(r)){tl=r[0]||0;tr=r[1]||0;br=r[2]||0;bl=r[3]||0;}else{tl=0;tr=0;br=0;bl=0;}this.moveTo(x+tl,y);this.lineTo(x+w-tr,y);this.quadraticCurveTo(x+w,y,x+w,y+tr);this.lineTo(x+w,y+h-br);this.quadraticCurveTo(x+w,y+h,x+w-br,y+h);this.lineTo(x+bl,y+h);this.quadraticCurveTo(x,y+h,x,y+h-bl);this.lineTo(x,y+tl);this.quadraticCurveTo(x,y,x+tl,y);this.closePath();return this;};}")

        ' ═══════════════════════════════════════════════════════════════════
        ' EMAIL NORMALIZER — fix casing from LLM output
        ' ═══════════════════════════════════════════════════════════════════
        sb.AppendLine("function normalizeEmail(e){if(!e||typeof e!=='object')return{from:'unknown@example.com',subject:'(no subject)',preview:'',isPhishing:false,redFlag:''};var o={};var keys=Object.keys(e);for(var i=0;i<keys.length;i++){o[keys[i].toLowerCase()]=e[keys[i]];}return{from:String(o.from||o.sender||o.email||'unknown@example.com'),subject:String(o.subject||o.title||'(no subject)'),preview:String(o.preview||o.body||o.snippet||''),isPhishing:!!(o.isphishing===true||o.isphishing==='true'||o.phishing===true||o.phishing==='true'),redFlag:String(o.redflag||o.red_flag||o.explanation||o.reason||'')};}")
        sb.AppendLine("function normalizeEmails(arr){if(!Array.isArray(arr))return[];var out=[];for(var i=0;i<arr.length;i++){out.push(normalizeEmail(arr[i]));}return out;}")

        ' ═══════════════════════════════════════════════════════════════════
        ' COMMON HUD
        ' ═══════════════════════════════════════════════════════════════════
        sb.AppendLine("function drawEmailHUD(email){")
        sb.AppendLine("  ctx.fillStyle='#1e293b';ctx.fillRect(0,0,800,100);")
        sb.AppendLine("  ctx.fillStyle='#e2e8f0';ctx.font='bold 16px sans-serif';ctx.textAlign='left';")
        sb.AppendLine("  ctx.fillText('Subject: '+(email.subject.length>80?email.subject.slice(0,79)+'…':email.subject),20,25);")
        sb.AppendLine("  ctx.fillStyle='#94a3b8';ctx.font='14px sans-serif';")
        sb.AppendLine("  ctx.fillText('From: '+(email.from.length>90?email.from.slice(0,89)+'…':email.from),20,45);")
        sb.AppendLine("  ctx.fillStyle='#cbd5e1';ctx.font='italic 14px sans-serif';")
        sb.AppendLine("  let words = email.preview.split(' '); let line = ''; let y = 65;")
        sb.AppendLine("  for(let n=0; n<words.length; n++){")
        sb.AppendLine("    let testLine = line + words[n] + ' ';")
        sb.AppendLine("    if(ctx.measureText(testLine).width > 760 && n > 0){")
        sb.AppendLine("      ctx.fillText(line, 20, y); line = words[n] + ' '; y += 18;")
        sb.AppendLine("    }else{ line = testLine; }")
        sb.AppendLine("  }")
        sb.AppendLine("  ctx.fillText(line, 20, y);")
        sb.AppendLine("  ctx.strokeStyle='#334155';ctx.beginPath();ctx.moveTo(0,100);ctx.lineTo(800,100);ctx.stroke();")
        sb.AppendLine("}")

        ' ═══════════════════════════════════════════════════════════════════
        ' GAME START DISPATCHER
        ' ═══════════════════════════════════════════════════════════════════
        sb.AppendLine("async function startGame(name){")
        sb.AppendLine("  cancelAnimationFrame(animFrame);gameOver=false;paused=false;score=0;wave=1;lives=3;")
        sb.AppendLine("  currentGame=name;scoreDisp.textContent='0';waveDisp.textContent='1';livesDisp.textContent='3';redFlagBox.innerHTML='';redFlagBox.style.display=(name==='defender'?'none':'block');")
        sb.AppendLine("  var titles={invaders:'👾 Inbox Invaders',pong:'🏓 Phish Pong',defender:'🟡 Data Defender',stack:'🧱 Risk Stack'};")
        sb.AppendLine("  document.getElementById('gameTitle').textContent=titles[name]||name;")
        sb.AppendLine("  showGame();")
        sb.AppendLine("  var diff=wave<=2?'easy':wave<=4?'medium':'hard';")
        sb.AppendLine("  await loadEmails(diff);")
        sb.AppendLine("  emails=normalizeEmails(emails);")
        sb.AppendLine("  if(!emails.length){emails=normalizeEmails(fallbackEmails);emailSource='Fallback';updateEmailSource();}")
        sb.AppendLine("  if(gameOver)return;")
        sb.AppendLine("  if(name==='invaders')initInvaders();")
        sb.AppendLine("  else if(name==='pong')initPong();")
        sb.AppendLine("  else if(name==='defender')initDefender();")
        sb.AppendLine("  else if(name==='stack')initStack();")
        sb.AppendLine("}")

        ' ═══════════════════════════════════════════════════════════════════
        ' GAME 1: INBOX INVADERS
        ' ═══════════════════════════════════════════════════════════════════
        sb.AppendLine("function initInvaders(){")
        sb.AppendLine("  const W=800,H=500;")
        sb.AppendLine("  let player={x:W/2,y:H-30,w:40,h:20,speed:6};")
        sb.AppendLine("  let bullets=[];")
        sb.AppendLine("  let emailIdx=0;")
        sb.AppendLine("  let currentEmail=emails[emailIdx];")
        sb.AppendLine("  let invader={x:W/2,y:100,w:60,h:40,speed:0.5+wave*0.2};")
        sb.AppendLine("  let cooldown=0;")
        sb.AppendLine("  function resetInvader(){invader={x:100+Math.random()*(W-200),y:100,w:60,h:40,speed:0.5+wave*0.2};bullets=[];}")
        sb.AppendLine("  function nextEmail(wasShot){")
        sb.AppendLine("    if(wasShot){")
        sb.AppendLine("      if(currentEmail.isPhishing){ score+=100; showRedFlag(currentEmail,true); }")
        sb.AppendLine("      else{ score-=50; lives--; showRedFlag(currentEmail,false); }")
        sb.AppendLine("    }else{")
        sb.AppendLine("      if(currentEmail.isPhishing){ score-=50; lives--; showRedFlag(currentEmail,false); }")
        sb.AppendLine("      else{ score+=100; showRedFlag(currentEmail,true); }")
        sb.AppendLine("    }")
        sb.AppendLine("    scoreDisp.textContent=score; livesDisp.textContent=lives;")
        sb.AppendLine("    emailIdx++;")
        sb.AppendLine("    if(emailIdx>=emails.length){")
        sb.AppendLine("      wave++; waveDisp.textContent=wave;")
        sb.AppendLine("      queueEmailRefresh(()=>{emailIdx=0;currentEmail=emails[emailIdx];resetInvader();});")
        sb.AppendLine("      if(lives<=0){ gameOver=true; showGameOver('Inbox Invaders',score,wave); }")
        sb.AppendLine("      return;")
        sb.AppendLine("    }")
        sb.AppendLine("    currentEmail=emails[emailIdx];")
        sb.AppendLine("    resetInvader();")
        sb.AppendLine("    if(lives<=0){ gameOver=true; showGameOver('Inbox Invaders',score,wave); }")
        sb.AppendLine("  }")
        sb.AppendLine("  function update(){")
        sb.AppendLine("    if(keys['ArrowLeft']||keys['a'])player.x=Math.max(player.w/2,player.x-player.speed);")
        sb.AppendLine("    if(keys['ArrowRight']||keys['d'])player.x=Math.min(W-player.w/2,player.x+player.speed);")
        sb.AppendLine("    if((keys[' ']||keys['ArrowUp'])&&cooldown<=0){bullets.push({x:player.x,y:player.y-10,speed:8});cooldown=20;}")
        sb.AppendLine("    if(cooldown>0)cooldown--;")
        sb.AppendLine("    bullets.forEach(b=>b.y-=b.speed);")
        sb.AppendLine("    bullets=bullets.filter(b=>b.y>100);")
        sb.AppendLine("    invader.y+=invader.speed;")
        sb.AppendLine("    for(let i=0;i<bullets.length;i++){")
        sb.AppendLine("      let b=bullets[i];")
        sb.AppendLine("      if(b.x>invader.x-invader.w/2 && b.x<invader.x+invader.w/2 && b.y>invader.y-invader.h/2 && b.y<invader.y+invader.h/2){")
        sb.AppendLine("        nextEmail(true); break;")
        sb.AppendLine("      }")
        sb.AppendLine("    }")
        sb.AppendLine("    if(invader.y>H){ nextEmail(false); }")
        sb.AppendLine("  }")
        sb.AppendLine("  function draw(){")
        sb.AppendLine("    ctx.clearRect(0,0,W,H);")
        sb.AppendLine("    drawEmailHUD(currentEmail);")
        sb.AppendLine("    ctx.fillStyle='#3b82f6';ctx.fillRect(player.x-player.w/2,player.y-player.h/2,player.w,player.h);")
        sb.AppendLine("    ctx.fillStyle='#fff';ctx.font='10px sans-serif';ctx.textAlign='center';ctx.fillText('SHOOT',player.x,player.y+4);")
        sb.AppendLine("    ctx.fillStyle='#fbbf24';bullets.forEach(b=>{ctx.fillRect(b.x-2,b.y-6,4,12);});")
        sb.AppendLine("    ctx.fillStyle='#ef4444';")
        sb.AppendLine("    ctx.beginPath();ctx.roundRect(invader.x-invader.w/2,invader.y-invader.h/2,invader.w,invader.h,4);ctx.fill();")
        sb.AppendLine("    ctx.fillStyle='#fff';ctx.font='20px sans-serif';ctx.fillText('👾',invader.x,invader.y+7);")
        sb.AppendLine("    ctx.fillStyle='rgba(255,255,255,.3)';ctx.font='12px sans-serif';")
        sb.AppendLine("    ctx.fillText('Shoot if Phishing. Let pass if Safe.',W/2,H-20);")
        sb.AppendLine("  }")
        sb.AppendLine("  function loop(){if(gameOver)return;animFrame=requestAnimationFrame(loop);if(!paused){update();draw();}}")
        sb.AppendLine("  loop();")
        sb.AppendLine("}")

        ' ═══════════════════════════════════════════════════════════════════
        ' GAME 2: PHISH PONG
        ' ═══════════════════════════════════════════════════════════════════
        sb.AppendLine("function initPong(){")
        sb.AppendLine("  const W=800,H=500;")
        sb.AppendLine("  let paddle={x:W/2,y:H-20,w:120,h:15,speed:8};")
        sb.AppendLine("  let emailIdx=0;")
        sb.AppendLine("  let currentEmail=emails[emailIdx];")
        sb.AppendLine("  let ball={x:W/2,y:120,vx:(Math.random()>0.5?3:-3),vy:3+wave*0.5,r:10};")
        sb.AppendLine("  function resetBall(){ball={x:W/2,y:120,vx:(Math.random()>0.5?3:-3),vy:3+wave*0.5,r:10};paddle.x=W/2;}")
        sb.AppendLine("  function nextEmail(zone){")
        sb.AppendLine("    if(zone===-1){")
        sb.AppendLine("      score-=50; lives--; showRedFlag(currentEmail,false);")
        sb.AppendLine("    }else if(zone===0){")
        sb.AppendLine("      if(!currentEmail.isPhishing){ score+=100; showRedFlag(currentEmail,true); }")
        sb.AppendLine("      else{ score-=50; lives--; showRedFlag(currentEmail,false); }")
        sb.AppendLine("    }else if(zone===1){")
        sb.AppendLine("      if(currentEmail.isPhishing){ score+=100; showRedFlag(currentEmail,true); }")
        sb.AppendLine("      else{ score-=50; lives--; showRedFlag(currentEmail,false); }")
        sb.AppendLine("    }")
        sb.AppendLine("    scoreDisp.textContent=score; livesDisp.textContent=lives;")
        sb.AppendLine("    emailIdx++;")
        sb.AppendLine("    if(emailIdx>=emails.length){")
        sb.AppendLine("      wave++; waveDisp.textContent=wave;")
        sb.AppendLine("      queueEmailRefresh(()=>{emailIdx=0;currentEmail=emails[emailIdx];resetBall();});")
        sb.AppendLine("      if(lives<=0){ gameOver=true; showGameOver('Phish Pong',score,wave); }")
        sb.AppendLine("      return;")
        sb.AppendLine("    }")
        sb.AppendLine("    currentEmail=emails[emailIdx];")
        sb.AppendLine("    resetBall();")
        sb.AppendLine("    if(lives<=0){ gameOver=true; showGameOver('Phish Pong',score,wave); }")
        sb.AppendLine("  }")
        sb.AppendLine("  function update(){")
        sb.AppendLine("    if(keys['ArrowLeft']||keys['a'])paddle.x=Math.max(paddle.w/2+20,paddle.x-paddle.speed);")
        sb.AppendLine("    if(keys['ArrowRight']||keys['d'])paddle.x=Math.min(W-paddle.w/2-20,paddle.x+paddle.speed);")
        sb.AppendLine("    ball.x+=ball.vx; ball.y+=ball.vy;")
        sb.AppendLine("    if(ball.y-ball.r<100){ ball.y=100+ball.r; ball.vy*=-1; }")
        sb.AppendLine("    if(ball.x-ball.r<20){ nextEmail(0); return; }")
        sb.AppendLine("    if(ball.x+ball.r>W-20){ nextEmail(1); return; }")
        sb.AppendLine("    if(ball.vy>0 && ball.y+ball.r>paddle.y-paddle.h/2 && ball.y-ball.r<paddle.y+paddle.h/2 && ball.x>paddle.x-paddle.w/2 && ball.x<paddle.x+paddle.w/2){")
        sb.AppendLine("      ball.y=paddle.y-paddle.h/2-ball.r;")
        sb.AppendLine("      ball.vy*=-1;")
        sb.AppendLine("      let hitPos = (ball.x - paddle.x) / (paddle.w/2);")
        sb.AppendLine("      ball.vx = hitPos * 6;")
        sb.AppendLine("    }")
        sb.AppendLine("    if(ball.y>H){ nextEmail(-1); }")
        sb.AppendLine("  }")
        sb.AppendLine("  function draw(){")
        sb.AppendLine("    ctx.clearRect(0,0,W,H);")
        sb.AppendLine("    drawEmailHUD(currentEmail);")
        sb.AppendLine("    ctx.fillStyle='rgba(34,197,94,.2)'; ctx.fillRect(0,100,20,H-100);")
        sb.AppendLine("    ctx.fillStyle='#22c55e'; ctx.font='bold 14px sans-serif'; ctx.textAlign='left';")
        sb.AppendLine("    ctx.save(); ctx.translate(15,H/2+40); ctx.rotate(-Math.PI/2); ctx.fillText('INBOX (Safe)',0,0); ctx.restore();")
        sb.AppendLine("    ctx.fillStyle='rgba(239,68,68,.2)'; ctx.fillRect(W-20,100,20,H-100);")
        sb.AppendLine("    ctx.fillStyle='#ef4444'; ctx.font='bold 14px sans-serif'; ctx.textAlign='left';")
        sb.AppendLine("    ctx.save(); ctx.translate(W-5,H/2+50); ctx.rotate(-Math.PI/2); ctx.fillText('SANDBOX (Phish)',0,0); ctx.restore();")
        sb.AppendLine("    ctx.fillStyle='#3b82f6'; ctx.fillRect(paddle.x-paddle.w/2,paddle.y-paddle.h/2,paddle.w,paddle.h);")
        sb.AppendLine("    ctx.fillStyle='#fbbf24'; ctx.beginPath(); ctx.arc(ball.x,ball.y,ball.r,0,Math.PI*2); ctx.fill();")
        sb.AppendLine("    ctx.fillStyle='rgba(255,255,255,.3)'; ctx.font='12px sans-serif'; ctx.textAlign='center';")
        sb.AppendLine("    ctx.fillText(""Bounce ball to Left (Safe) or Right (Phishing). Don't let it drop!"",W/2,H-40);")
        sb.AppendLine("  }")
        sb.AppendLine("  function loop(){if(gameOver)return;animFrame=requestAnimationFrame(loop);if(!paused){update();draw();}}")
        sb.AppendLine("  loop();")
        sb.AppendLine("}")

        ' ═══════════════════════════════════════════════════════════════════
        ' GAME 3: DATA DEFENDER (Pac-Man style)
        ' ═══════════════════════════════════════════════════════════════════
        sb.AppendLine("function initDefender(){")
        sb.AppendLine("  const W=800,H=500,CS=40;")
        sb.AppendLine("  const cols=Math.floor(W/CS),rows=Math.floor(H/CS);")
        sb.AppendLine("  let maze=[];")
        sb.AppendLine("  function genMaze(){for(let r=0;r<rows;r++){maze[r]=[];for(let c=0;c<cols;c++){maze[r][c]=(r===0||r===rows-1||c===0||c===cols-1)?1:(Math.random()<0.22?1:0);}}maze[1][1]=0;maze[1][2]=0;maze[2][1]=0;maze[2][2]=0;}")
        sb.AppendLine("  function floodFill(){const vis=[];for(let r=0;r<rows;r++){vis[r]=[];for(let c=0;c<cols;c++)vis[r][c]=false;}const q=[[1,1]];vis[1][1]=true;let count=0;while(q.length>0){const[cr,cc]=q.shift();count++;const dirs=[[0,1],[0,-1],[1,0],[-1,0]];for(const[dr,dc] of dirs){const nr=cr+dr,nc=cc+dc;if(nr>=0&&nr<rows&&nc>=0&&nc<cols&&!vis[nr][nc]&&maze[nr][nc]===0){vis[nr][nc]=true;q.push([nr,nc]);}}}return count;}")
        sb.AppendLine("  function totalOpen(){let c=0;for(let r=0;r<rows;r++)for(let k=0;k<cols;k++)if(maze[r][k]===0)c++;return c;}")
        sb.AppendLine("  for(let attempt=0;attempt<50;attempt++){genMaze();const reach=floodFill();const open=totalOpen();if(reach>=open*0.9)break;}")
        sb.AppendLine("  let px=1,py=1; let dir={x:0,y:0}, nextDir={x:0,y:0}; let moveTimer=0;")
        sb.AppendLine("  let collectibles=[]; let emailIdx=0;")
        sb.AppendLine("  for(let r=0;r<rows;r++)for(let c=0;c<cols;c++){if(maze[r][c]===0&&!(r<=2&&c<=2)&&Math.random()<0.15){const em=emails[emailIdx%emails.length];emailIdx++;collectibles.push({x:c,y:r,email:em,collected:false});}}")
        sb.AppendLine("  if(collectibles.length<3){for(let r=1;r<rows-1;r++)for(let c=1;c<cols-1;c++){if(maze[r][c]===0&&!(r<=2&&c<=2)&&collectibles.length<8){const em=emails[emailIdx%emails.length];emailIdx++;collectibles.push({x:c,y:r,email:em,collected:false});}}}")
        sb.AppendLine("  let ghosts=[];")
        sb.AppendLine("  for(let i=0;i<Math.min(1+wave,4);i++){let gx,gy;do{gx=Math.floor(Math.random()*(cols-2))+1;gy=Math.floor(Math.random()*(rows-2))+1;}while(maze[gy][gx]!==0||(Math.abs(gx-1)+Math.abs(gy-1)<5));ghosts.push({x:gx,y:gy,moveTimer:0});}")
        sb.AppendLine("  let activeEmail=null;")
        sb.AppendLine("  let answerFeedback=null;")
        sb.AppendLine("  let feedbackTimer=0;")
        sb.AppendLine("  function handleAnswer(isPhishingChoice){")
        sb.AppendLine("    const correct=isPhishingChoice===activeEmail.isPhishing;")
        sb.AppendLine("    if(correct){score+=100;}else{score-=50;lives--;}")
        sb.AppendLine("    scoreDisp.textContent=score;livesDisp.textContent=lives;")
        sb.AppendLine("    answerFeedback={correct:correct,email:activeEmail};")
        sb.AppendLine("    feedbackTimer=45;")
        sb.AppendLine("    if(lives<=0){ gameOver=true; showGameOver('Data Defender',score,wave); }")
        sb.AppendLine("  }")
        sb.AppendLine("  function drawCenteredWrap(text,x,y,maxWidth,lineHeight){")
        sb.AppendLine("    const words=text.split(' '); let line=''; let yy=y;")
        sb.AppendLine("    for(let n=0;n<words.length;n++){")
        sb.AppendLine("      let testLine=line+words[n]+' ';")
        sb.AppendLine("      if(ctx.measureText(testLine).width>maxWidth && n>0){ctx.fillText(line,x,yy);line=words[n]+' ';yy+=lineHeight;}else{line=testLine;}")
        sb.AppendLine("    }")
        sb.AppendLine("    ctx.fillText(line,x,yy);")
        sb.AppendLine("  }")
        sb.AppendLine("  function update(){")
        sb.AppendLine("    if(activeEmail){")
        sb.AppendLine("      if(answerFeedback){")
        sb.AppendLine("        feedbackTimer--;")
        sb.AppendLine("        if(feedbackTimer<=0){activeEmail=null;answerFeedback=null;}")
        sb.AppendLine("        return;")
        sb.AppendLine("      }")
        sb.AppendLine("      if(keys['y']||keys['Y']){")
        sb.AppendLine("        handleAnswer(true); keys['y']=false; keys['Y']=false;")
        sb.AppendLine("      }else if(keys['n']||keys['N']){")
        sb.AppendLine("        handleAnswer(false); keys['n']=false; keys['N']=false;")
        sb.AppendLine("      }")
        sb.AppendLine("      return;")
        sb.AppendLine("    }")
        sb.AppendLine("    if(keys['ArrowLeft']||keys['a']) nextDir={x:-1,y:0};")
        sb.AppendLine("    if(keys['ArrowRight']||keys['d']) nextDir={x:1,y:0};")
        sb.AppendLine("    if(keys['ArrowUp']||keys['w']) nextDir={x:0,y:-1};")
        sb.AppendLine("    if(keys['ArrowDown']||keys['s']) nextDir={x:0,y:1};")
        sb.AppendLine("    moveTimer++;")
        sb.AppendLine("    if(moveTimer>=8){")
        sb.AppendLine("      moveTimer=0;")
        sb.AppendLine("      let nx=px+nextDir.x, ny=py+nextDir.y;")
        sb.AppendLine("      if(nx>=0&&nx<cols&&ny>=0&&ny<rows&&maze[ny][nx]===0){ dir=nextDir; px=nx; py=ny; }")
        sb.AppendLine("      else{")
        sb.AppendLine("        nx=px+dir.x; ny=py+dir.y;")
        sb.AppendLine("        if(nx>=0&&nx<cols&&ny>=0&&ny<rows&&maze[ny][nx]===0){ px=nx; py=ny; }")
        sb.AppendLine("      }")
        sb.AppendLine("      ghosts.forEach(g=>{")
        sb.AppendLine("        g.moveTimer++;")
        sb.AppendLine("        if(g.moveTimer>=2){")
        sb.AppendLine("          g.moveTimer=0;")
        sb.AppendLine("          const dirs=[{x:1,y:0},{x:-1,y:0},{x:0,y:1},{x:0,y:-1}];")
        sb.AppendLine("          const valid=dirs.filter(d=>{const nx=g.x+d.x,ny=g.y+d.y;return nx>=0&&nx<cols&&ny>=0&&ny<rows&&maze[ny][nx]===0;});")
        sb.AppendLine("          if(valid.length>0){")
        sb.AppendLine("            const biased=valid.filter(d=>{const nx=g.x+d.x,ny=g.y+d.y;return(Math.abs(nx-px)+Math.abs(ny-py))<(Math.abs(g.x-px)+Math.abs(g.y-py));});")
        sb.AppendLine("            const pick=biased.length>0&&Math.random()>0.2?biased:valid;")
        sb.AppendLine("            const chosen=pick[Math.floor(Math.random()*pick.length)];")
        sb.AppendLine("            g.x+=chosen.x; g.y+=chosen.y;")
        sb.AppendLine("          }")
        sb.AppendLine("        }")
        sb.AppendLine("      });")
        sb.AppendLine("    }")
        sb.AppendLine("    collectibles.forEach(c=>{")
        sb.AppendLine("      if(!c.collected && c.x===px && c.y===py){")
        sb.AppendLine("        c.collected=true; activeEmail=c.email; dir={x:0,y:0}; nextDir={x:0,y:0};")
        sb.AppendLine("      }")
        sb.AppendLine("    });")
        sb.AppendLine("    ghosts.forEach(g=>{")
        sb.AppendLine("      if(g.x===px && g.y===py){")
        sb.AppendLine("        lives--; livesDisp.textContent=lives; px=1; py=1; dir={x:0,y:0}; nextDir={x:0,y:0};")
        sb.AppendLine("        if(lives<=0){ gameOver=true; showGameOver('Data Defender',score,wave); }")
        sb.AppendLine("      }")
        sb.AppendLine("    });")
        sb.AppendLine("    if(collectibles.every(c=>c.collected)){ wave++; waveDisp.textContent=wave; showMenu(); startGame('defender'); return; }")
        sb.AppendLine("  }")
        sb.AppendLine("  function draw(){")
        sb.AppendLine("    ctx.clearRect(0,0,W,H);")
        sb.AppendLine("    for(let r=0;r<rows;r++)for(let c=0;c<cols;c++){if(maze[r][c]===1){ctx.fillStyle='#1e293b';ctx.fillRect(c*CS,r*CS,CS,CS);}}")
        sb.AppendLine("    collectibles.forEach(c=>{if(!c.collected){ctx.fillStyle='#22c55e';ctx.beginPath();ctx.arc(c.x*CS+CS/2,c.y*CS+CS/2,6,0,Math.PI*2);ctx.fill();}});")
        sb.AppendLine("    ghosts.forEach(g=>{ctx.fillStyle='#ef4444';ctx.beginPath();ctx.arc(g.x*CS+CS/2,g.y*CS+CS/2,CS/2-4,0,Math.PI*2);ctx.fill();ctx.fillStyle='#fff';ctx.font='12px sans-serif';ctx.textAlign='center';ctx.fillText('👻',g.x*CS+CS/2,g.y*CS+CS/2+4);});")
        sb.AppendLine("    ctx.fillStyle='#3b82f6';ctx.beginPath();ctx.arc(px*CS+CS/2,py*CS+CS/2,CS/2-4,0,Math.PI*2);ctx.fill();")
        sb.AppendLine("    if(activeEmail){")
        sb.AppendLine("      ctx.fillStyle='rgba(0,0,0,0.8)'; ctx.fillRect(0,0,W,H);")
        sb.AppendLine("      ctx.fillStyle='#1e293b'; ctx.strokeStyle='#3b82f6'; ctx.lineWidth=2;")
        sb.AppendLine("      ctx.beginPath(); ctx.roundRect(100,100,600,300,8); ctx.fill(); ctx.stroke();")
        sb.AppendLine("      ctx.fillStyle='#e2e8f0'; ctx.font='bold 20px sans-serif'; ctx.textAlign='center';")
        sb.AppendLine("      ctx.fillText('Analyze Email', W/2, 140);")
        sb.AppendLine("      ctx.textAlign='left'; ctx.font='bold 16px sans-serif';")
        sb.AppendLine("      ctx.fillText('Subject: '+activeEmail.subject, 130, 180);")
        sb.AppendLine("      ctx.fillStyle='#94a3b8'; ctx.font='14px sans-serif';")
        sb.AppendLine("      ctx.fillText('From: '+activeEmail.from, 130, 210);")
        sb.AppendLine("      ctx.fillStyle='#cbd5e1'; ctx.font='italic 14px sans-serif';")
        sb.AppendLine("      let words = activeEmail.preview.split(' ');")
        sb.AppendLine("      let line = ''; let y = 250;")
        sb.AppendLine("      for(let n=0; n<words.length; n++){")
        sb.AppendLine("        let testLine = line + words[n] + ' ';")
        sb.AppendLine("        let metrics = ctx.measureText(testLine);")
        sb.AppendLine("        if(metrics.width > 540 && n > 0){")
        sb.AppendLine("          ctx.fillText(line, 130, y);")
        sb.AppendLine("          line = words[n] + ' '; y += 20;")
        sb.AppendLine("        }else{ line = testLine; }")
        sb.AppendLine("      }")
        sb.AppendLine("      ctx.fillText(line, 130, y);")
        sb.AppendLine("      ctx.textAlign='center';")
        sb.AppendLine("      if(answerFeedback){")
        sb.AppendLine("        ctx.fillStyle=answerFeedback.correct?'#22c55e':'#ef4444'; ctx.font='bold 18px sans-serif';")
        sb.AppendLine("        ctx.fillText(answerFeedback.correct?'Result: Correct':'Result: Incorrect', W/2, 345);")
        sb.AppendLine("        ctx.fillStyle='#cbd5e1'; ctx.font='13px sans-serif';")
        sb.AppendLine("        drawCenteredWrap('Reason: '+(answerFeedback.email.isPhishing?'Phishing':'Legitimate')+'. '+answerFeedback.email.redFlag, W/2, 370, 520, 16);")
        sb.AppendLine("      }else{")
        sb.AppendLine("        ctx.fillStyle='#fbbf24'; ctx.font='bold 18px sans-serif';")
        sb.AppendLine("        ctx.fillText('Is this Phishing? Press [Y] for Yes, [N] for No', W/2, 360);")
        sb.AppendLine("      }")
        sb.AppendLine("    }")
        sb.AppendLine("  }")
        sb.AppendLine("  function loop(){if(gameOver)return;animFrame=requestAnimationFrame(loop);if(!paused){update();draw();}}")
        sb.AppendLine("  loop();")
        sb.AppendLine("}")

        ' ═══════════════════════════════════════════════════════════════════
        ' GAME 4: RISK STACK (Tetris-style)
        ' ═══════════════════════════════════════════════════════════════════
        sb.AppendLine("function initStack(){")
        sb.AppendLine("  const W=800,H=500;")
        sb.AppendLine("  let emailIdx=0;")
        sb.AppendLine("  let currentEmail=emails[emailIdx];")
        sb.AppendLine("  let riskBar=0; const maxRisk=100;")
        sb.AppendLine("  const zones=[{label:'✅ Safe',x:0,w:W/3,color:'rgba(34,197,94,.12)'},{label:'🔍 Verify',x:W/3,w:W/3,color:'rgba(234,179,8,.12)'},{label:'🚫 Block',x:W*2/3,w:W/3,color:'rgba(239,68,68,.12)'}];")
        sb.AppendLine("  let blockX=W/3, blockY=100;")
        sb.AppendLine("  let fallSpeed=1+wave*0.3;")
        sb.AppendLine("  let keyCooldown=0;")
        sb.AppendLine("  let playerMoved=false;")
        sb.AppendLine("  function resetBlock(){blockX=W/3; blockY=100; fallSpeed=1+wave*0.3; playerMoved=false;}")
        sb.AppendLine("  function judge(zoneIdx){")
        sb.AppendLine("    if(!playerMoved){")
        sb.AppendLine("      score-=50;riskBar+=20;showRedFlag(currentEmail,false);")
        sb.AppendLine("      scoreDisp.textContent=score;")
        sb.AppendLine("      if(riskBar>=maxRisk){lives--;livesDisp.textContent=lives;riskBar=0;}")
        sb.AppendLine("      emailIdx++;")
        sb.AppendLine("      if(emailIdx>=emails.length){")
        sb.AppendLine("        wave++;waveDisp.textContent=wave;")
        sb.AppendLine("        queueEmailRefresh(()=>{emailIdx=0;currentEmail=emails[emailIdx];resetBlock();});")
        sb.AppendLine("        if(lives<=0){gameOver=true;showGameOver('Risk Stack',score,wave);}")
        sb.AppendLine("        return;")
        sb.AppendLine("      }")
        sb.AppendLine("      currentEmail=emails[emailIdx];")
        sb.AppendLine("      resetBlock();")
        sb.AppendLine("      if(lives<=0){gameOver=true;showGameOver('Risk Stack',score,wave);}")
        sb.AppendLine("      return;")
        sb.AppendLine("    }")
        sb.AppendLine("    const em=currentEmail;")
        sb.AppendLine("    if(em.isPhishing){")
        sb.AppendLine("      if(zoneIdx===2){score+=100;showRedFlag(em,true);}")
        sb.AppendLine("      else if(zoneIdx===1){score+=30;showRedFlag(em,true);}")
        sb.AppendLine("      else{score-=50;riskBar+=20;showRedFlag(em,false);}")
        sb.AppendLine("    }else{")
        sb.AppendLine("      if(zoneIdx===0){score+=100;showRedFlag(em,true);}")
        sb.AppendLine("      else if(zoneIdx===1){score+=30;showRedFlag(em,true);}")
        sb.AppendLine("      else{score-=50;riskBar+=20;showRedFlag(em,false);}")
        sb.AppendLine("    }")
        sb.AppendLine("    scoreDisp.textContent=score;")
        sb.AppendLine("    if(riskBar>=maxRisk){lives--;livesDisp.textContent=lives;riskBar=0;}")
        sb.AppendLine("    emailIdx++;")
        sb.AppendLine("    if(emailIdx>=emails.length){")
        sb.AppendLine("      wave++;waveDisp.textContent=wave;")
        sb.AppendLine("      queueEmailRefresh(()=>{emailIdx=0;currentEmail=emails[emailIdx];resetBlock();});")
        sb.AppendLine("      if(lives<=0){gameOver=true;showGameOver('Risk Stack',score,wave);}")
        sb.AppendLine("      return;")
        sb.AppendLine("    }")
        sb.AppendLine("    currentEmail=emails[emailIdx];")
        sb.AppendLine("    resetBlock();")
        sb.AppendLine("    if(lives<=0){gameOver=true;showGameOver('Risk Stack',score,wave);}")
        sb.AppendLine("  }")
        sb.AppendLine("  function update(){")
        sb.AppendLine("    if(keyCooldown>0) keyCooldown--;")
        sb.AppendLine("    if(keyCooldown<=0){")
        sb.AppendLine("      if(keys['ArrowLeft']||keys['a']){ blockX=Math.max(0, blockX-W/3); keyCooldown=15; playerMoved=true; }")
        sb.AppendLine("      else if(keys['ArrowRight']||keys['d']){ blockX=Math.min(W*2/3, blockX+W/3); keyCooldown=15; playerMoved=true; }")
        sb.AppendLine("      else if(keys['ArrowDown']||keys['s']){ blockY+=10; playerMoved=true; }")
        sb.AppendLine("    }")
        sb.AppendLine("    if(!keys['ArrowLeft']&&!keys['a']&&!keys['ArrowRight']&&!keys['d']){")
        sb.AppendLine("      if(keyCooldown>0 && !keys['ArrowDown'] && !keys['s']) keyCooldown=0;")
        sb.AppendLine("    }")
        sb.AppendLine("    blockY+=fallSpeed;")
        sb.AppendLine("    if(blockY>H-40){")
        sb.AppendLine("      let zoneIdx = blockX<W/3 ? 0 : (blockX<W*2/3 ? 1 : 2);")
        sb.AppendLine("      judge(zoneIdx);")
        sb.AppendLine("    }")
        sb.AppendLine("  }")
        sb.AppendLine("  function draw(){")
        sb.AppendLine("    ctx.clearRect(0,0,W,H);")
        sb.AppendLine("    drawEmailHUD(currentEmail);")
        sb.AppendLine("    zones.forEach(z=>{ctx.fillStyle=z.color;ctx.fillRect(z.x,100,z.w,H-100);ctx.fillStyle='rgba(255,255,255,.3)';ctx.font='bold 16px sans-serif';ctx.textAlign='center';ctx.fillText(z.label,z.x+z.w/2,H-20);});")
        sb.AppendLine("    ctx.setLineDash([4,6]);ctx.strokeStyle='#334155';ctx.beginPath();ctx.moveTo(W/3,100);ctx.lineTo(W/3,H);ctx.moveTo(W*2/3,100);ctx.lineTo(W*2/3,H);ctx.stroke();ctx.setLineDash([]);")
        sb.AppendLine("    ctx.fillStyle='#1e293b';ctx.fillRect(W-20,110,10,H-120);")
        sb.AppendLine("    const rh=(riskBar/maxRisk)*(H-120);ctx.fillStyle=riskBar>70?'#ef4444':riskBar>40?'#eab308':'#22c55e';ctx.fillRect(W-20,H-10-rh,10,rh);")
        sb.AppendLine("    ctx.fillStyle='#3b82f6'; ctx.beginPath(); ctx.roundRect(blockX+20, blockY, W/3-40, 40, 8); ctx.fill();")
        sb.AppendLine("    ctx.fillStyle='#fff'; ctx.font='bold 14px sans-serif'; ctx.textAlign='center';")
        sb.AppendLine("    ctx.fillText('Drop Here', blockX+W/6, blockY+25);")
        sb.AppendLine("  }")
        sb.AppendLine("  function loop(){if(gameOver)return;animFrame=requestAnimationFrame(loop);if(!paused){update();draw();}}")
        sb.AppendLine("  loop();")
        sb.AppendLine("}")

        sb.AppendLine("</script>")
        sb.AppendLine("</body></html>")

        Return sb.ToString()
    End Function

End Class