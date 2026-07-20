# Document

The system office-document handler for the Cinatra artifact library. It recognizes uploaded office documents — Word (`.docx`), Excel (`.xlsx`), PowerPoint (`.pptx`), and OpenDocument Text (`.odt`) — and files them in the library under a dedicated document type, so an office file you attach in chat or upload to `/artifacts` lands correctly typed instead of being refused.

Install from the Cinatra marketplace by searching for "Document" and clicking **Add**. No credentials or configuration are required; the type is active immediately for all workspace members. Opening a document names its concrete format and size with a download affordance — a browser cannot render these formats inline, so the detail view is a clear typed download shell, never a blank panel.

## Works with

- Cinatra chat — attach a Word, Excel, PowerPoint, or OpenDocument file directly in any thread
- The artifact library — open any document item to download it

## Capabilities

- Accept Word, Excel, PowerPoint, and OpenDocument Text uploads as a dedicated artifact type
- Name the concrete office format and size with a download affordance
- Never a blank view — the shell always renders the document identity
