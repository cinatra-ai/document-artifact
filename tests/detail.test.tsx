import { afterEach, describe, expect, it } from "vitest";
import { cleanup, render } from "@testing-library/react";

import DocumentArtifactDetail, {
  formatBytes,
  officeFormatLabel,
} from "../src/renderers/detail";
import {
  OPENXML_PRESENTATION_MIME,
  isOpenXmlPresentation,
} from "../src/renderers/openxml-presentation-viewer";
import type { ArtifactRendererProps } from "../src/artifact-renderer-props";

afterEach(cleanup);

function props(overrides: {
  download?: string | null;
  title?: string | null;
  mime?: string;
  size?: number;
}): ArtifactRendererProps {
  const mime =
    overrides.mime ??
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
  return {
    propsApiVersion: 2,
    artifact: {
      id: "art_1",
      title: overrides.title === undefined ? "report.docx" : overrides.title,
      objectType: "@cinatra-ai/document-artifact:artifact",
      mime,
      size: overrides.size === undefined ? 32_768 : overrides.size,
      createdAt: "2026-01-01T00:00:00.000Z",
      updatedAt: "2026-01-01T00:00:00.000Z",
      ownerLevel: "workspace",
      visibility: "organization",
      sourceUrl: null,
    },
    representation: { revisionId: "rev_1", mime },
    urls: {
      preview: null,
      download: overrides.download === undefined ? "/api/artifacts/art_1/download" : overrides.download,
    },
    identity: { kind: "no-primary", extension: null },
    actions: {
      download: overrides.download === undefined ? "/api/artifacts/art_1/download" : overrides.download,
      openInSource: null,
    },
  };
}

describe("officeFormatLabel", () => {
  it("maps each accepted office media type to a human label", () => {
    expect(
      officeFormatLabel("application/vnd.openxmlformats-officedocument.wordprocessingml.document"),
    ).toBe("Word document");
    expect(
      officeFormatLabel("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
    ).toBe("Excel spreadsheet");
    expect(
      officeFormatLabel("application/vnd.openxmlformats-officedocument.presentationml.presentation"),
    ).toBe("PowerPoint presentation");
    expect(officeFormatLabel("application/vnd.oasis.opendocument.text")).toBe("OpenDocument text");
  });

  it("tolerates a media-type parameter and falls back to a generic label", () => {
    expect(
      officeFormatLabel(
        "application/vnd.oasis.opendocument.text; charset=utf-8",
      ),
    ).toBe("OpenDocument text");
    expect(officeFormatLabel(null)).toBe("Document");
    expect(officeFormatLabel("application/octet-stream")).toBe("Document");
  });
});

describe("formatBytes", () => {
  it("formats byte sizes and rejects invalid input", () => {
    expect(formatBytes(32_768)).toBe("32.0 KB");
    expect(formatBytes(null)).toBeNull();
    expect(formatBytes(-5)).toBeNull();
  });
});

describe("DocumentArtifactDetail — the office-document download shell", () => {
  it("names the concrete office format and offers a download", () => {
    const { container } = render(<DocumentArtifactDetail {...props({})} />);
    const shell = container.querySelector('[data-document-artifact="shell"]');
    expect(shell?.tagName.toLowerCase()).toBe("article");
    expect(shell?.getAttribute("class")).toContain("soft-panel rounded-card");
    expect(shell?.textContent).toContain("Word document");
    expect(container.querySelector("a")?.getAttribute("href")).toBe("/api/artifacts/art_1/download");
  });

  it("derives the format label from the representation media type", () => {
    const { container } = render(
      <DocumentArtifactDetail
        {...props({
          title: null,
          mime: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        })}
      />,
    );
    expect(container.textContent).toContain("Excel spreadsheet");
  });

  it("embeds nothing for the take-away forms (a browser cannot present them)", () => {
    for (const mime of [
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      "application/vnd.oasis.opendocument.text",
    ]) {
      cleanup();
      const { container } = render(
        <DocumentArtifactDetail
          {...props({ mime })}
          bytes={{ road: "island", preview: "/island-preview", download: "/island-dl" }}
        />,
      );
      expect(container.querySelector("iframe")).toBeNull();
      expect(container.querySelector("embed")).toBeNull();
      expect(container.querySelector("object")).toBeNull();
      expect(container.querySelector('[data-document-artifact="shell"]')).not.toBeNull();
    }
  });

  it("NEVER-BLANK: a null download URL still renders the shell without a link", () => {
    const { container } = render(<DocumentArtifactDetail {...props({ download: null })} />);
    expect(container.querySelector("a")).toBeNull();
    expect(container.querySelector('[data-document-artifact="shell"]')).not.toBeNull();
    expect((container.textContent ?? "").trim().length).toBeGreaterThan(0);
  });

  it("tolerates a malformed snapshot missing urls/actions (never throws, never blank)", () => {
    const malformed = {
      propsApiVersion: 2,
      artifact: { title: null },
    } as unknown as ArtifactRendererProps;
    const { container } = render(<DocumentArtifactDetail {...malformed} />);
    const shell = container.querySelector('[data-document-artifact="shell"]');
    expect(shell).not.toBeNull();
    expect(shell?.textContent).toContain("Document");
    expect((container.textContent ?? "").trim().length).toBeGreaterThan(0);
  });
});

describe("DocumentArtifactDetail — the embedded OpenXML viewer on the presentation form", () => {
  const ISLAND_PREVIEW = "/api/lifecycle-views/artifact-bytes?bc=sealed-preview";
  const ISLAND_DOWNLOAD = "/api/lifecycle-views/artifact-bytes?bc=sealed-download";

  it("recognises the presentation form, and only it", () => {
    expect(isOpenXmlPresentation(OPENXML_PRESENTATION_MIME)).toBe(true);
    expect(isOpenXmlPresentation(`${OPENXML_PRESENTATION_MIME}; charset=utf-8`)).toBe(true);
    expect(
      isOpenXmlPresentation(
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      ),
    ).toBe(false);
    expect(isOpenXmlPresentation(null)).toBe(false);
  });

  it("embeds the deck instead of the typed download shell", () => {
    const { container } = render(
      <DocumentArtifactDetail
        {...props({ mime: OPENXML_PRESENTATION_MIME, title: "Q3 review.pptx" })}
        bytes={{ road: "island", preview: ISLAND_PREVIEW, download: ISLAND_DOWNLOAD }}
      />,
    );
    const viewer = container.querySelector('[data-document-artifact="presentation-viewer"]');
    expect(viewer).not.toBeNull();
    expect(viewer?.getAttribute("data-byte-road")).toBe("island");
    const object = container.querySelector("object");
    expect(object?.getAttribute("data")).toBe(ISLAND_PREVIEW);
    expect(object?.getAttribute("type")).toBe(OPENXML_PRESENTATION_MIME);
    // The viewer replaces the standalone shell, and never the other way round.
    expect(container.querySelector('[data-document-artifact="shell"]')).toBeNull();
  });

  it("carries the typed panel as the viewer's own fallback — never a blank frame", () => {
    const { container } = render(
      <DocumentArtifactDetail
        {...props({ mime: OPENXML_PRESENTATION_MIME, title: "Q3 review.pptx" })}
        bytes={{ road: "island", preview: ISLAND_PREVIEW, download: ISLAND_DOWNLOAD }}
      />,
    );
    const fallback = container.querySelector('[data-document-artifact="shell-fallback"]');
    expect(fallback).not.toBeNull();
    expect(fallback?.textContent).toContain("PowerPoint presentation");
    expect(fallback?.querySelector("a")?.getAttribute("href")).toBe(ISLAND_DOWNLOAD);
  });

  it("embeds the byte reference and never the cookie-gated session route", () => {
    const { container } = render(
      <DocumentArtifactDetail
        {...props({ mime: OPENXML_PRESENTATION_MIME })}
        urls={{ preview: "/session-preview", download: "/session-dl" }}
        actions={{ download: "/session-dl", openInSource: null }}
        bytes={{ road: "island", preview: ISLAND_PREVIEW, download: ISLAND_DOWNLOAD }}
      />,
    );
    expect(container.innerHTML).not.toContain("/session-preview");
    expect(container.innerHTML).not.toContain("/session-dl");
  });

  it("floors to the typed shell when no road carries a presentable address", () => {
    const { container } = render(
      <DocumentArtifactDetail {...props({ mime: OPENXML_PRESENTATION_MIME })} propsApiVersion={1} />,
    );
    expect(container.querySelector("object")).toBeNull();
    const shell = container.querySelector('[data-document-artifact="shell"]');
    expect(shell).not.toBeNull();
    expect(shell?.getAttribute("data-byte-road")).toBe("session");
    expect(shell?.textContent).toContain("PowerPoint presentation");
  });
});

describe("DocumentArtifactDetail — the byte road (props version 2)", () => {
  it("offers the reference's download address on the shell forms", () => {
    const { container } = render(
      <DocumentArtifactDetail
        {...props({})}
        bytes={{ road: "island", preview: null, download: "/island-dl" }}
      />,
    );
    expect(container.querySelector("a")?.getAttribute("href")).toBe("/island-dl");
    expect(
      container.querySelector('[data-document-artifact="shell"]')?.getAttribute("data-byte-road"),
    ).toBe("island");
  });

  it("falls back to the session href on an older snapshot that carries no reference", () => {
    const { container } = render(<DocumentArtifactDetail {...props({})} propsApiVersion={1} />);
    expect(container.querySelector("a")?.getAttribute("href")).toBe(
      "/api/artifacts/art_1/download",
    );
    expect(
      container.querySelector('[data-document-artifact="shell"]')?.getAttribute("data-byte-road"),
    ).toBe("session");
  });

  it("stays a typed, never-blank shell when no road carries an address", () => {
    const { container } = render(
      <DocumentArtifactDetail {...props({ download: null })} propsApiVersion={1} />,
    );
    expect(container.querySelector("a")).toBeNull();
    expect(
      container.querySelector('[data-document-artifact="shell"]')?.getAttribute("data-byte-road"),
    ).toBe("none");
  });
});
