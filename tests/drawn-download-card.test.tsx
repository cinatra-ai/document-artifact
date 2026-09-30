/**
 * THE DOWNLOAD CARD, AS THE RATIFIED DRAWING DRAWS IT (the review drawing §V.2).
 *
 * VERBATIM: "it belongs to a base of its own, and that base's display is the
 * download card — the file's name, its form, its size, and the download — and no
 * reading of the bytes, because there is none to make." And: "It has no tabs and
 * nothing else to put in a header, so it carries no header strip at all: the
 * file, and the download."
 *
 * FOUR PARTS, AND NOTHING APPENDED TO THEM. The card the drawing draws reads
 *
 *   upgrade-road-notes.bin
 *   application/octet-stream · 2.4 MB
 *   Download
 *
 * — a name, a form beside a size, and a control that says one word. This display
 * carried a fifth part the drawing does not give: a sentence telling the reader
 * what to do with the file, and a control that said what it was downloading. A
 * proof round graded both of them as deviations from the drawn card.
 */
import { afterEach, describe, expect, it } from "vitest";
import { cleanup, render } from "@testing-library/react";

import Detail from "../src/renderers/detail";
import type { ArtifactRendererProps } from "../src/artifact-renderer-props";

afterEach(cleanup);

const PROPS = {
  propsApiVersion: 2,
  artifact: {
    id: "art_1",
    title: "onboarding-checklist.docx",
    objectType: "@cinatra-ai/document-artifact:artifact",
    mime: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    size: 951,
    createdAt: "2026-01-01T00:00:00.000Z",
    updatedAt: "2026-01-01T00:00:00.000Z",
    ownerLevel: "workspace",
    visibility: "organization",
    sourceUrl: null,
  },
  representation: { revisionId: "rev_1", mime: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" },
  urls: { preview: null, download: "/d" },
  identity: { kind: "no-primary", extension: null },
  actions: { download: "/d", openInSource: null },
} as unknown as ArtifactRendererProps;

describe("the card reads the four parts the drawing gives it", () => {
  it("says Download, and says nothing else on the control", () => {
    const { container } = render(<Detail {...PROPS} />);
    const link = container.querySelector("a[download]");
    expect(link).not.toBeNull();
    expect(link?.textContent?.trim()).toBe("Download");
  });

  it("appends no instruction to the form and the size", () => {
    const { container } = render(<Detail {...PROPS} />);
    const text = container.textContent ?? "";
    expect(text).not.toMatch(/Download to open/);
    expect(text).not.toMatch(/native application/);
    expect(text).not.toMatch(/open its contents/);
  });

  it("still draws the file, its form and its size", () => {
    const { container } = render(<Detail {...PROPS} />);
    const text = container.textContent ?? "";
    expect(text).toContain("onboarding-checklist.docx");
    expect(text).toContain("Word document");
    expect(text).toContain("951 B");
  });
});
