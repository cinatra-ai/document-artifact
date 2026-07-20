import { afterEach, describe, expect, it } from "vitest";
import { cleanup, render } from "@testing-library/react";

import DocumentArtifactDetail, {
  formatBytes,
  officeFormatLabel,
} from "../src/renderers/detail";
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
    propsApiVersion: 1,
    artifact: {
      id: "art_1",
      title: overrides.title === undefined ? "report.docx" : overrides.title,
      objectType: "@cinatra-ai/document-artifact:artifact",
      mime,
      size: overrides.size === undefined ? 32_768 : overrides.size,
      createdAt: "2026-01-01T00:00:00.000Z",
      updatedAt: "2026-01-01T00:00:00.000Z",
      ownerLevel: "workspace",
      visibility: "workspace",
      sourceUrl: null,
    },
    representation: { revisionId: "rev_1", mime },
    urls: {
      preview: null,
      download: overrides.download === undefined ? "/api/artifacts/art_1/download" : overrides.download,
    },
    identity: { kind: "mime", extension: null, basis: null, selectable: false },
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

  it("does not render an inline frame (browsers cannot render office formats)", () => {
    const { container } = render(<DocumentArtifactDetail {...props({})} />);
    expect(container.querySelector("iframe")).toBeNull();
    expect(container.querySelector("embed")).toBeNull();
  });

  it("NEVER-BLANK: a null download URL still renders the shell without a link", () => {
    const { container } = render(<DocumentArtifactDetail {...props({ download: null })} />);
    expect(container.querySelector("a")).toBeNull();
    expect(container.querySelector('[data-document-artifact="shell"]')).not.toBeNull();
    expect((container.textContent ?? "").trim().length).toBeGreaterThan(0);
  });

  it("tolerates a malformed snapshot missing urls/actions (never throws, never blank)", () => {
    const malformed = {
      propsApiVersion: 1,
      artifact: { title: null },
    } as unknown as ArtifactRendererProps;
    const { container } = render(<DocumentArtifactDetail {...malformed} />);
    const shell = container.querySelector('[data-document-artifact="shell"]');
    expect(shell).not.toBeNull();
    expect(shell?.textContent).toContain("Document");
    expect((container.textContent ?? "").trim().length).toBeGreaterThan(0);
  });
});
