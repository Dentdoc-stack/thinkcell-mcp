# Use it as a connector (Claude in PowerPoint, claude.ai)

Connectors such as Canva are hosted services. This one cannot be: think-cell
only runs on a Windows PC with PowerPoint and a think-cell licence. So **your
own PC is the server**. `connector.py` runs it in the background and gives it a
fixed public HTTPS address through [ngrok](https://ngrok.com) (free). Claude
then adds it like any other connector: paste the URL, press **Connect**, and
type a code.

```
Claude (cloud) ──HTTPS──▶ ngrok ──▶ your PC: server.py --http ──▶ ppttc.exe + PowerPoint
```

Claude can only use it while your PC is on and the connector is running.

## One-time setup (about 10 minutes)

1. **Install the server** (you need Python 3.10+):

   ```powershell
   git clone https://github.com/Dentdoc-stack/thinkcell-mcp.git
   cd thinkcell-mcp
   python -m pip install -r requirements.txt
   ```

2. **Get a free ngrok account.** At [dashboard.ngrok.com](https://dashboard.ngrok.com)
   you need two things: your **authtoken** (*Getting Started → Your Authtoken*)
   and your free **static domain** (*Domains*), something like
   `quiet-otter-123.ngrok-free.app`.

3. **Run the setup** and answer its questions:

   ```powershell
   python connector.py setup
   ```

   It checks that think-cell is installed and finds ngrok, offering to download
   it if needed. It then asks for your authtoken, which is not shown as you
   type, and your domain. It also asks for any folders whose Excel/CSV files or
   templates Claude may read. Then it starts the connector, offers to start it
   with Windows, and prints two things:

   - the **connector URL**: `https://<your-domain>/mcp`
   - your **connect code**, like `K7MP-Q2XD`

4. **Add the connector in Claude.** On claude.ai go to *Settings → Connectors →
   Add custom connector*. Paste the connector URL and press **Add**, then
   **Connect**. A page from your PC opens. Type the connect code there. On the
   free ngrok plan a "You are about to visit" page from ngrok comes first:
   press **Visit Site**.

   The connector now shows up in Claude in PowerPoint and in Claude chats on the
   same account. It stays signed in. Claude renews its own token for up to
   90 days of inactivity.

## Use it from think-cell Assist (PowerPoint and Excel)

think-cell Assist can call MCP servers too. It needs a think-cell AI Suite
licence and a think-cell version with Assist (Early Access). Older versions,
such as 14, don't have the *MCP Servers* policy. It connects from a browser
page (`https://app.prod.ai.think-cell.com`), and it signs in only with fixed
HTTP headers, not OAuth. The connector handles both:

- It answers Assist's CORS preflight and adds the CORS headers its manual
  lists.
- It accepts an **Assist key** as `Authorization: Bearer <key>`.

Requests from any other browser origin get 403.

1. Run `python connector.py assist`. It prints the name, URL and headers to
   enter. Use `--new` to replace the key.
2. Enter them as an administrator:
   - **Windows:** in the group policy *thinkcell Assist: MCP Servers*, select
     **Enabled**, then add the Name, URL and Headers. Headers are a JSON
     object, for example `{"Authorization": "Bearer <key>"}`.
   - **macOS:** run the `sudo defaults write … com.think-cell.settings.assist …`
     command that `assist` prints. For a second server, use `MCP2_…` keys.
3. In PowerPoint or Excel, go to *Insert → thinkcell → Assist → Recent
   History*. Under **MCP Servers**, a green dot means connected; red means
   a connection error.

On the free ngrok plan, `assist` also adds the header
`ngrok-skip-browser-warning`. Without it, ngrok answers browsers with its
warning page instead of the server. To allow another browser-based client,
add its origin to `"origins"` in `config.json`.

## Everyday use

| Command | What it does |
| --- | --- |
| `python connector.py status` | Is it running and reachable from the internet? |
| `python connector.py start` / `stop` | Start or stop it in the background. |
| `python connector.py code` | Show the connect code. |
| `python connector.py code --new` | Make a new code; then `stop` and `start`. |
| `python connector.py signout` | Sign every Claude connection out. They must Connect again. |
| `python connector.py assist` | Show the think-cell Assist settings; `--new` makes a new key. |
| `python connector.py autostart on` / `off` | Start with Windows, or not. |

The log is in `%APPDATA%\thinkcell-connector\connector.log`.

In PowerPoint, ask Claude to build the chart. The `insert_into_powerpoint`
tool then puts the slides straight into the deck you have open. The files
are also saved under `output\` in the project folder.

For a chart type the auto deck doesn't draw, ask Claude to pick one from
think-cell's chart library (`list_chart_templates`). To reuse a slide of your
own, open the deck in PowerPoint and ask Claude to save slide N as a template
(`template_from_slide`). Then ask for that slide with new data
(`fill_chart_template`).

## Security

- **Who can use it:** only someone who knows your connect code, or the
  Assist key. Without a
  token the server answers 401. Wrong codes are limited: 5 tries per sign-in
  and 10 per 10 minutes overall. Keep the code private, like a password.
- **What it can touch:** Claude can read and save files only in `output\`,
  `templates\` and the folders you listed in setup. Any other path is refused.
  To change the folders, run `setup` again.
- **Where tokens live:** tokens are stored only as SHA-256 hashes in
  `%APPDATA%\thinkcell-connector\oauth_state.json`. Your ngrok authtoken stays
  in ngrok's own config.
- **How it is reached:** the server listens on `127.0.0.1` only. The internet
  reaches it through the ngrok tunnel, nothing else.
- **If you think the code leaked:** run `code --new`, `signout`, `stop`, then
  `start`.

## Troubleshooting

- **The ngrok download fails.** Some antivirus tools block it. Install ngrok
  yourself from the Microsoft Store ("ngrok") or
  [ngrok.com/download](https://ngrok.com/download), then run `setup` again.
- **Connect works, then every call fails.** Antivirus web shields, such as
  Norton or Avast, can break local HTTP replies. The server already sends
  them in a form these shields leave alone. If calls still fail, add an
  exception for `python.exe` and `ngrok.exe` in the shield.
- **`status` says not reachable and the log shows `x509: certificate signed
  by unknown authority`.** An antivirus web shield (Norton, Avast) is scanning
  ngrok's own connection to its servers. Add an exclusion for `ngrok.exe` (or
  `connect.ngrok-agent.com`) in the shield's settings, then `stop` and
  `start`. ngrok's `connect_cas: host` option does not help: the shield marks
  ngrok's certificate as untrusted.
- **Pasting the authtoken does nothing in setup.** Some terminals block paste
  into hidden input. Press Enter there and run
  `ngrok config add-authtoken <token>` in another terminal.
- **`status` says not reachable.** Check that the PC is online and that the
  domain in setup matches your ngrok dashboard. Then look at the end of
  `connector.log`.
- **think-cell errors.** Ask Claude to run `diagnose_thinkcell`. JSON
  automation needs a licence that includes it (think-cell AI Suite in the
  current manual).

## Without connector.py

`python server.py --http [--port 8765] [--public-url https://your.domain]` runs
the server alone, so you can put your own tunnel or reverse proxy in front of
it. It reads the connect code and folders from the same config, so run
`connector.py setup` once anyway, or write
`%APPDATA%\thinkcell-connector\config.json` yourself with these keys:
`public_url`, `connect_code`, `port` and `folders`. Change the location with
the `THINKCELL_CONNECTOR_HOME` environment variable.

## MCP protocol versions

The `/mcp` endpoint serves both MCP eras. Clients on 2024-11-05 to 2025-11-25
start with `initialize`. Clients on 2026-07-28 skip it and are answered
statelessly (`server/discover`, `tools/list`, `tools/call`). The local stdio
server (`python server.py`, the `.mcpb` extension) speaks only the older
versions. Clients that support both fall back to them.
