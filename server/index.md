wordbridge {{version}} — Office live-editing bridge

Endpoints:
  GET  /status        bridge health + connected add-in clients
  GET  /tools         full tool catalog (JSON Schema) for LLM callers
  GET  /tools/<name>  one tool
  POST /op            execute one op          body: { kind, ... }
  POST /ops           execute a batch of ops  body: [ {...}, ... ]
  GET  /addin/...     Office.js task pane assets
  WS   /ws            task-pane connection (internal)

Add-in Manifests:
  Word:       {{origin}}/addin/manifest.xml
  Excel:      {{origin}}/addin/excel/manifest.xml
  PowerPoint: {{origin}}/addin/powerpoint/manifest.xml

Sideload an add-in:

  Option A — Upload via Office UI:
    1. Open the app > Insert > Get Add-ins (or Add-ins > My Add-ins)
    2. Click "Upload My Add-in" (under Manage My Add-ins or via the dropdown)
    3. Browse and provide the manifest URL for the corresponding app
    4. The Bridge task pane will appear on the right

  Option B — Install via manifest folder (no UI needed):

    macOS:
      # Word
      mkdir -p ~/Library/Containers/com.microsoft.Word/Data/Documents/wef
      curl -o ~/Library/Containers/com.microsoft.Word/Data/Documents/wef/manifest.xml \
           {{origin}}/addin/manifest.xml

      # Excel
      mkdir -p ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef
      curl -o ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/manifest.xml \
           {{origin}}/addin/excel/manifest.xml

      # PowerPoint
      mkdir -p ~/Library/Containers/com.microsoft.PowerPoint/Data/Documents/wef
      curl -o ~/Library/Containers/com.microsoft.PowerPoint/Data/Documents/wef/manifest.xml \
           {{origin}}/addin/powerpoint/manifest.xml

    Windows (all apps share one WEF folder):
      mkdir "%LOCALAPPDATA%\Microsoft\Office\16.0\WEF" 2>nul
      curl -o "%LOCALAPPDATA%\Microsoft\Office\16.0\WEF\wordbridge.xml" ^
           {{origin}}/addin/manifest.xml
      curl -o "%LOCALAPPDATA%\Microsoft\Office\16.0\WEF\excelbridge.xml" ^
           {{origin}}/addin/excel/manifest.xml
      curl -o "%LOCALAPPDATA%\Microsoft\Office\16.0\WEF\pptbridge.xml" ^
           {{origin}}/addin/powerpoint/manifest.xml

    Restart the app after placing the manifest. The add-in appears in Insert > My Add-ins.
    The add-in connects to this server via WebSocket automatically.

Start with: curl {{origin}}/tools | jq .
LLM callers: read /tools, pick a tool, POST its example to /op.
