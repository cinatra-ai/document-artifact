// Document detail renderer (slot `detail`).
//
// The office formats this base accepts (Word .docx, Excel .xlsx, PowerPoint
// .pptx, OpenDocument Text .odt) cannot be rendered inline by a browser, and the
// v1 renderer snapshot carries only host-authorized URLs — never the document
// bytes. So the faithful minimal renderer is a typed download SHELL: it names
// the concrete office format (derived from the declared media type), shows the
// size, and offers a download affordance. This matches the sibling bases'
// passive-URL depth — no client-side document conversion / preview pipeline
// (that would be gold-plating well beyond a system base).
//
// v1 renderer: requests NO host ports; renders ONLY from the host-supplied
// authorized snapshot (`ArtifactRendererProps`).
//
// NEVER-BLANK: the shell always renders the document identity; the panel is
// never empty even when no download URL was authorized.

import type { ReactElement } from "react";

import type { ArtifactRendererProps } from "../artifact-renderer-props";

/** Map a declared office media type to a human format label (pure; exported for
 *  tests). Unknown / absent media types fall back to the generic "Document". */
export function officeFormatLabel(mime: string | null | undefined): string {
  switch ((mime ?? "").split(";", 1)[0]?.trim().toLowerCase()) {
    case "application/vnd.openxmlformats-officedocument.wordprocessingml.document":
      return "Word document";
    case "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet":
      return "Excel spreadsheet";
    case "application/vnd.openxmlformats-officedocument.presentationml.presentation":
      return "PowerPoint presentation";
    case "application/vnd.oasis.opendocument.text":
      return "OpenDocument text";
    default:
      return "Document";
  }
}

/** Human-readable byte size for the shell (pure; exported for tests). */
export function formatBytes(size: number | null | undefined): string | null {
  if (typeof size !== "number" || !Number.isFinite(size) || size < 0) return null;
  if (size < 1024) return `${size} B`;
  const units = ["KB", "MB", "GB", "TB"];
  let n = size / 1024;
  let i = 0;
  while (n >= 1024 && i < units.length - 1) {
    n /= 1024;
    i += 1;
  }
  return `${n.toFixed(1)} ${units[i]}`;
}

export default function DocumentArtifactDetail(props: ArtifactRendererProps): ReactElement {
  const downloadHref = props.actions?.download ?? props.urls?.download ?? null;
  const title = props.artifact?.title ?? null;
  const format = officeFormatLabel(props.representation?.mime ?? props.artifact?.mime);
  const size = formatBytes(props.artifact?.size);
  const heading = title ?? format;

  return (
    <article
      className="soft-panel rounded-card overflow-hidden p-6"
      data-document-artifact="shell"
    >
      <p className="text-sm font-medium">{heading}</p>
      <p className="text-sm text-muted-foreground">
        {format}{size ? ` · ${size}` : ""}. Download to open in its native application.
      </p>
      {downloadHref ? (
        <a href={downloadHref} className="text-sm underline" download>
          Download the document
        </a>
      ) : null}
    </article>
  );
}
