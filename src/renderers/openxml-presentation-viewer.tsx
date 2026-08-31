// THE EMBEDDED OPENXML VIEWER for the presentation form.
//
// The presentation form is the one office form this base accepts that a reader
// expects to READ in place rather than to take away, and the wave the byte road
// belongs to trades this form's typed download shell for an embedded viewer.
// Every other office form this base accepts keeps the shell.
//
// THE SHAPE IS THE ONE THE PDF DISPLAY ALREADY ESTABLISHED, deliberately: an
// embedding element pointed at the address the byte road resolved, with the
// typed download panel as its OWN fallback content — the element's built-in
// fallback, which the browser draws whenever it cannot present the object
// itself. So the never-blank floor is not a second code path that has to be
// kept in step; it is the same panel, in the same place, always rendered.
//
// NO CONVERSION PIPELINE AND NO THIRD-PARTY VIEWER SERVICE. The renderer hands
// the reader's own browser an address and a media type and lets whatever handler
// it has present them; it fetches nothing, parses nothing, and sends the work's
// bytes nowhere. A client-side OpenXML rendering dependency would be a far
// larger commitment than a system base carries, and a hosted viewer would send
// the reviewed work to a third party — neither is what "embedded" asks for.

import type { ReactElement, ReactNode } from "react";

import type { ByteRoadName } from "./byte-road";

/** The presentation form — the one office form that gets the viewer. */
export const OPENXML_PRESENTATION_MIME =
  "application/vnd.openxmlformats-officedocument.presentationml.presentation";

/** Is this the presentation form? Pure; tolerant of a parameterized media type. */
export function isOpenXmlPresentation(mime: string | null | undefined): boolean {
  return (mime ?? "").split(";", 1)[0]?.trim().toLowerCase() === OPENXML_PRESENTATION_MIME;
}

export function OpenXmlPresentationViewer({
  src,
  road,
  label,
  fallback,
}: {
  readonly src: string;
  readonly road: ByteRoadName;
  readonly label: string;
  readonly fallback: ReactNode;
}): ReactElement {
  return (
    <article
      className="soft-panel rounded-card overflow-hidden p-0"
      data-document-artifact="presentation-viewer"
      data-byte-road={road}
    >
      <object
        data={src}
        type={OPENXML_PRESENTATION_MIME}
        // 75vh so the slides fill most of the viewport without forcing the page
        // to scroll; the presented object handles its own paging.
        className="h-[75vh] w-full"
        aria-label={label}
      >
        {fallback}
      </object>
    </article>
  );
}
