// Document detail renderer (slot `detail`).
//
// This base accepts four office formats — Word .docx, Excel .xlsx, PowerPoint
// .pptx and OpenDocument Text .odt. Three of them a reader takes away rather
// than reads in place, and for those the faithful minimal renderer is a typed
// download SHELL: it names the concrete office format (derived from the
// declared media type), shows the size, and offers a download affordance.
//
// THE PRESENTATION FORM IS THE EXCEPTION. It trades that shell for an embedded
// OpenXML viewer pointed at the address the byte road resolved, with the same
// typed panel as the viewer's own fallback content — so the reader reads the
// deck in place where the browser can present it, and lands on exactly the old
// shell where it cannot.
//
// THE ADDRESS COMES FROM THE BYTE ROAD. Inside a third-party application the
// host's session route carries no cookie, so a shell offering it hands the
// reader a dead link and a viewer pointed at it draws a blank plate. At props
// version 2 the snapshot carries the byte reference the reader may actually
// fetch on the surface they are on; a snapshot built at the older version has
// no reference and falls back to the session href. The renderer requests NO
// host ports, builds no address of its own, and fetches nothing.
//
// NEVER-BLANK: the typed panel always renders the document identity, whether on
// its own or as the viewer's fallback — the panel is never empty even when no
// road carried an address.

import type { ReactElement } from "react";

import type { ArtifactRendererProps } from "../artifact-renderer-props";
import { resolveByteRoad, type ByteRoadName } from "./byte-road";
import {
  OpenXmlPresentationViewer,
  isOpenXmlPresentation,
} from "./openxml-presentation-viewer";

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

/**
 * The typed download shell — the panel three of the four forms draw on their
 * own, and the panel the presentation viewer falls back to. Exported so the two
 * uses cannot drift apart.
 *
 * IT DRAWS THE FOUR PARTS THE DRAWING GIVES A DOWNLOAD CARD, AND NO FIFTH (the
 * review drawing §V.2): "the file's name, its form, its size, and the download".
 * It used to append a sentence telling the reader what to do with the file, and
 * its control said what it was downloading; a proof round graded both as
 * deviations from the drawn card, and neither is a part of it.
 */
export function DocumentDownloadShell({
  heading,
  format,
  size,
  downloadHref,
  road,
  embedded = false,
}: {
  readonly heading: string;
  readonly format: string;
  readonly size: string | null;
  readonly downloadHref: string | null;
  readonly road: ByteRoadName;
  readonly embedded?: boolean;
}): ReactElement {
  return (
    <article
      className="soft-panel rounded-card overflow-hidden p-6"
      data-document-artifact={embedded ? "shell-fallback" : "shell"}
      data-byte-road={road}
    >
      <p className="text-sm font-medium">{heading}</p>
      <p className="text-sm text-muted-foreground">
        {format}{size ? ` · ${size}` : ""}
      </p>
      {downloadHref ? (
        <a href={downloadHref} className="text-sm underline" download>
          Download
        </a>
      ) : null}
    </article>
  );
}

export default function DocumentArtifactDetail(props: ArtifactRendererProps): ReactElement {
  const bytes = resolveByteRoad(props);
  const mime = props.representation?.mime ?? props.artifact?.mime;
  const title = props.artifact?.title ?? null;
  const format = officeFormatLabel(mime);
  const size = formatBytes(props.artifact?.size);
  const heading = title ?? format;

  const shell = (
    <DocumentDownloadShell
      heading={heading}
      format={format}
      size={size}
      downloadHref={bytes.download}
      road={bytes.road}
    />
  );

  // The presentation form reads in place wherever the browser can present it.
  // An attachment-disposition address would prompt a download instead of
  // presenting, so the viewer mounts only over a preview address.
  if (isOpenXmlPresentation(mime) && bytes.preview) {
    return (
      <OpenXmlPresentationViewer
        src={bytes.preview}
        road={bytes.road}
        label={`Presentation preview: ${heading}`}
        fallback={
          <DocumentDownloadShell
            heading={heading}
            format={format}
            size={size}
            downloadHref={bytes.download}
            road={bytes.road}
            embedded
          />
        }
      />
    );
  }

  return shell;
}
