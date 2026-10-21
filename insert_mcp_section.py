import sys
sys.stdout.reconfigure(encoding='utf-8')

fpath = r'C:\Users\hodowns\AppData\Local\Temp\USPS-PMO-Training\index.html'
content = open(fpath, 'r', encoding='utf-8').read()

new_box = (
    '\n'
    '  <!-- Outlook MCP Deep Dive -->\n'
    '  <div class="box">\n'
    '    <h3>&#128231; Outlook MCP &mdash; What PMs Need to Know</h3>\n'
    '    <p style="margin-bottom:16px;">The Outlook MCP connects Claude directly to your Cisco M365 inbox, calendar, and contacts. Every prompt in the use-cases section that says &ldquo;Using the Outlook MCP&hellip;&rdquo; runs through this server. This section covers what is unique to Outlook: data policy, how authentication works, and read-only vs. write mode.</p>\n'
    '\n'
    '    <div class="callout c-danger" style="margin-bottom:20px;">\n'
    '      <div class="ci">&#x1F6A8;</div>\n'
    '      <div>\n'
    '        <strong>Data Classification &mdash; Required Reading</strong>\n'
    '        <p style="margin-top:6px;margin-bottom:8px;">Outlook email is classified up to <strong>Restricted</strong> &mdash; a higher classification than most Cisco AI tools are approved to handle. The Cisco-hosted Outlook MCP in Claude Code is approved, but you are responsible for where that data goes.</p>\n'
    '        <ul style="margin:0 0 0 18px;line-height:1.9;">\n'
    '          <li><strong>Do not route email content to unapproved tools</strong> &mdash; Circuit, personal AI assistants, external services, or any tool not in Cisco&rsquo;s approved catalog.</li>\n'
    '          <li><strong>Not sure?</strong> Use the <a href="https://dataadvisor.cloudapps.cisco.com/da/cwizard" target="_blank">Cisco Data Advisor &rarr;</a> to verify your workflow handles data at the right classification level.</li>\n'
    '          <li><strong>Accidental exposure must be reported.</strong> If Restricted email content may have reached an unapproved system, report it immediately to your manager and CISO.</li>\n'
    '        </ul>\n'
    '      </div>\n'
    '    </div>\n'
    '\n'
    '    <p style="font-weight:700;font-size:.9rem;margin-bottom:8px;">&#x1F510; Authentication &mdash; What to Expect</p>\n'
    '    <p style="margin-bottom:10px;">The Outlook MCP authenticates through your browser the first time you use it in a session. After that it stores your session automatically &mdash; no repeated logins unless the token expires.</p>\n'
    '    <table style="margin-bottom:16px;">\n'
    '      <tr><th>Situation</th><th>What to do</th></tr>\n'
    '      <tr><td><strong>First use:</strong> a browser window opens to Microsoft login</td><td>Sign in with your Cisco credentials. Complete 2FA if prompted. Close the browser &mdash; Claude picks up automatically.</td></tr>\n'
    '      <tr><td><strong>Token expires</strong> (~45&ndash;60 min of inactivity)</td><td>Claude pauses and asks you to re-authenticate. Same browser flow &mdash; Claude resumes when done.</td></tr>\n'
    '      <tr><td><strong>Already authenticated:</strong> session stored</td><td>Nothing to do &mdash; Claude reuses your session automatically.</td></tr>\n'
    '    </table>\n'
    '    <div class="callout c-tip" style="margin-bottom:20px;"><div class="ci">&#x1F4A1;</div><div>Token expiry is expected &mdash; the Outlook OAuth flow does not support refresh tokens. If Claude stops mid-task and asks to re-authenticate, complete the browser login and it will pick up where it left off.</div></div>\n'
    '\n'
    '    <p style="font-weight:700;font-size:.9rem;margin-bottom:8px;">&#x1F4D6; Read Mode vs. Write Mode</p>\n'
    '    <p style="margin-bottom:10px;">By default the Outlook MCP is <strong>read-only</strong> &mdash; it can read email, calendar, and contacts, but cannot send, reply, delete, or create events. Write tools must be explicitly turned on.</p>\n'
    '    <table style="margin-bottom:16px;">\n'
    '      <tr><th>Mode</th><th>What it can do</th><th>Requires</th></tr>\n'
    '      <tr>\n'
    '        <td><strong>Read-only</strong><br><span style="font-size:.75rem;font-weight:400;color:var(--muted);">Default &mdash; no setup needed</span></td>\n'
    '        <td>List &amp; search emails &middot; Read message threads &middot; List calendar events &middot; Look up contacts &middot; List folders &amp; inbox rules</td>\n'
    '        <td>Nothing extra &mdash; works after first authentication</td>\n'
    '      </tr>\n'
    '      <tr>\n'
    '        <td><strong>Write mode</strong><br><span style="font-size:.75rem;font-weight:400;color:var(--muted);">Optional &mdash; must enable</span></td>\n'
    '        <td>Send / reply / forward emails &middot; Create &amp; manage drafts &middot; Create, update &amp; delete calendar events &middot; Move or delete messages &middot; Manage inbox rules</td>\n'
    '        <td><code style="background:#f1f5f9;padding:1px 5px;border-radius:4px;">OUTLOOK_MCP_ENABLE_WRITE_TOOLS=true</code> in your MCP config <em>plus</em> <code style="background:#f1f5f9;padding:1px 5px;border-radius:4px;">confirm=true</code> in each write prompt</td>\n'
    '      </tr>\n'
    '    </table>\n'
    '    <div class="callout c-warn" style="margin-bottom:20px;"><div class="ci">&#x26A0;&#xFE0F;</div><div><strong>Write tools require two steps by design &mdash; this is intentional.</strong> You must (1) enable write tools in your config, and (2) include <code style="background:#f1f5f9;padding:1px 5px;border-radius:4px;">confirm=true</code> in each prompt that writes. If either is missing, Claude previews the action and waits. This prevents accidental sends and deletes.</div></div>\n'
    '\n'
    '    <p style="font-weight:700;font-size:.9rem;margin-bottom:8px;">&#x2699;&#xFE0F; How to Enable Write Tools</p>\n'
    '    <p style="margin-bottom:10px;">If you want to draft, send, reply, or manage calendar events through Claude, paste this after the Outlook MCP is installed:</p>\n'
    '    <div class="prompt" style="margin-bottom:6px;"><div class="lbl sample">PROMPT &mdash; ENABLE WRITE TOOLS</div><span class="hi">I want to enable write tools for the Outlook MCP so I can send emails and manage calendar events from Claude. Open my MCP settings file, add OUTLOOK_MCP_ENABLE_WRITE_TOOLS=true to the Outlook MCP environment variables, and show me what you are changing before you save it.</span></div>\n'
    '    <p style="font-size:.82rem;color:var(--muted);margin-bottom:20px;">Claude will find the config file, show you the change, and apply it. Restart Claude after the change for it to take effect.</p>\n'
    '\n'
    '    <p style="font-weight:700;font-size:.9rem;margin-bottom:10px;">&#x2705; PM Use Cases &mdash; What to Try First</p>\n'
    '    <table style="margin-bottom:14px;">\n'
    '      <tr><th>Task</th><th>Mode</th><th>How to phrase it</th></tr>\n'
    '      <tr><td>Read a client&rsquo;s recent emails</td><td>Read</td><td>&ldquo;Using the Outlook MCP, read my emails from [client name] in the last 7 days.&rdquo;</td></tr>\n'
    '      <tr><td>Search by topic</td><td>Read</td><td>&ldquo;Using the Outlook MCP, search for emails about [topic] in the last 30 days.&rdquo;</td></tr>\n'
    '      <tr><td>Check your calendar for the week</td><td>Read</td><td>&ldquo;Using the Outlook MCP, list my calendar events for this week with attendees and times.&rdquo;</td></tr>\n'
    '      <tr><td>Look up a contact</td><td>Read</td><td>&ldquo;Using the Outlook MCP, look up [name]&rsquo;s email address.&rdquo;</td></tr>\n'
    '      <tr><td>Create a draft reply (review before sending)</td><td>Write</td><td>&ldquo;Using the Outlook MCP, create a draft reply to the latest email from [client]. [instructions]. Do not send &mdash; save as draft only.&rdquo;</td></tr>\n'
    '      <tr><td>Send a draft after reviewing it</td><td>Write</td><td>&ldquo;Using the Outlook MCP, send the draft I just reviewed. confirm=true.&rdquo;</td></tr>\n'
    '      <tr><td>Create a calendar event</td><td>Write</td><td>&ldquo;Using the Outlook MCP, create a 1-hour event called [name] on [date] at [time] with [attendees]. confirm=true.&rdquo;</td></tr>\n'
    '      <tr><td>Triage unread email by priority</td><td>Read</td><td>&ldquo;Using the Outlook MCP, read my unread emails from the last 24 hours. Group by priority: needs reply today, FYI only, can wait.&rdquo;</td></tr>\n'
    '    </table>\n'
    '    <div class="callout c-info" style="margin-bottom:0;"><div class="ci">&#x1F517;</div><div>Source code and full tool reference (all 25 tools, environment variable options, authentication details): <a href="https://wwwin-github.cisco.com/msagadev/outlook-mcp-server" target="_blank"><strong>wwwin-github.cisco.com/msagadev/outlook-mcp-server &rarr;</strong></a>&nbsp; (VPN required)</div></div>\n'
    '  </div>\n'
)

marker = 'someone will help you get the MCP set up correctly for your environment.</p>\n  </div>\n</div>\n\n<!-- \u2550\u2550 SKILLS \u2550\u2550 -->'
replacement = 'someone will help you get the MCP set up correctly for your environment.</p>\n  </div>\n' + new_box + '\n</div>\n\n<!-- \u2550\u2550 SKILLS \u2550\u2550 -->'

if marker in content:
    new_content = content.replace(marker, replacement, 1)
    open(fpath, 'w', encoding='utf-8').write(new_content)
    lines_before = content.count('\n')
    lines_after = new_content.count('\n')
    print('SUCCESS: %d -> %d lines (+%d)' % (lines_before, lines_after, lines_after - lines_before))
else:
    print('ERROR: marker not found')
