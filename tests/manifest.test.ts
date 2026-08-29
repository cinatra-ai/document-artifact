// @vitest-environment node
import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";

import { describe, expect, it } from "vitest";

import { documentArtifactManifest } from "../src/index";

const pkg = JSON.parse(
  readFileSync(fileURLToPath(new URL("../package.json", import.meta.url)), "utf8"),
) as {
  name: string;
  cinatra: {
    apiVersion: string;
    kind: string;
    displayName: string;
    vendor: { key: string; name: string };
    dependencies: unknown[];
    artifact: {
      accepts: { file: { mimeTypes: string[] } };
      ui: {
        abiVersion: number;
        sdkAbiRange: string;
        renderers: Record<string, { entry: string; propsApiVersion: number; representations?: string[] }>;
      };
      objectTypes: Array<{
        type: string;
        claim: string;
        dispositions: Record<string, unknown>;
        schema: Record<string, unknown>;
      }>;
    };
  };
};

const MIMES = ["application/vnd.openxmlformats-officedocument.wordprocessingml.document","application/vnd.openxmlformats-officedocument.spreadsheetml.sheet","application/vnd.openxmlformats-officedocument.presentationml.presentation","application/vnd.oasis.opendocument.text"];

const ARTIFACT_ALLOWED_CINATRA_KEYS = new Set([
  "kind",
  "apiVersion",
  "artifact",
  "dependencies",
  "roles",
  "displayName",
  "vendor",
]);
const ARTIFACT_UI_RENDERER_ALLOWED_KEYS = new Set(["entry", "propsApiVersion", "representations"]);

describe("package.json manifest — the system-base document identity", () => {
  it("names the package per the @cinatra-ai/<slug>-artifact convention", () => {
    expect(pkg.name).toBe("@cinatra-ai/document-artifact");
  });

  it("declares the first-party artifact identity", () => {
    expect(pkg.cinatra.kind).toBe("artifact");
    expect(pkg.cinatra.apiVersion).toBe("cinatra.ai/v1");
    expect(pkg.cinatra.displayName).toBe("Document");
    expect(pkg.cinatra.vendor).toEqual({ key: "cinatra-ai", name: "Cinatra" });
  });

  it("omits dependency edges (a system base is platform-guaranteed)", () => {
    expect(pkg.cinatra.dependencies).toEqual([]);
  });

  it("declares only the allowed top-level cinatra.* keys", () => {
    for (const k of Object.keys(pkg.cinatra)) {
      expect(ARTIFACT_ALLOWED_CINATRA_KEYS.has(k)).toBe(true);
    }
    // A renderer base ships no matcher skill bundle.
    expect("skills" in pkg.cinatra.artifact).toBe(false);
  });

  it("CLAIMS EXACTLY the declared MIME roster and nothing else (no wildcards)", () => {
    expect(pkg.cinatra.artifact.accepts.file.mimeTypes).toEqual(MIMES);
    expect(pkg.cinatra.artifact.ui.renderers.detail.representations).toEqual(MIMES);
    // No wildcard entry — a required base is a dedicated, non-universal MIME home.
    for (const m of pkg.cinatra.artifact.accepts.file.mimeTypes) {
      expect(m.includes("*")).toBe(false);
    }
  });

  it("declares a strict v1 ui block bound to the host SDK ABI", () => {
    const ui = pkg.cinatra.artifact.ui;
    expect(ui.abiVersion).toBe(1);
    expect(ui.sdkAbiRange).toBe("^2.5.0");
    expect(Object.keys(ui.renderers)).toEqual(["detail"]);
  });

  it("the detail renderer requests NO host ports (v1 no-ports contract)", () => {
    const detail = pkg.cinatra.artifact.ui.renderers.detail;
    for (const k of Object.keys(detail)) {
      expect(ARTIFACT_UI_RENDERER_ALLOWED_KEYS.has(k)).toBe(true);
    }
    expect(detail.propsApiVersion).toBe(1);
  });

  it("points the detail entry at a package-contained subpath that exists", () => {
    const entry = pkg.cinatra.artifact.ui.renderers.detail.entry;
    expect(entry).toBe("./src/renderers/detail.tsx");
    expect(entry.startsWith("./")).toBe(true);
    expect(entry.includes("..")).toBe(false);
    const resolved = fileURLToPath(new URL(`../${entry.slice(2)}`, import.meta.url));
    expect(() => readFileSync(resolved, "utf8")).not.toThrow();
  });

  it("declares exactly one dedicated objectTypes claim for the upload type map", () => {
    const claims = pkg.cinatra.artifact.objectTypes;
    expect(Array.isArray(claims)).toBe(true);
    expect(claims).toHaveLength(1);
    const claim = claims[0];
    expect(claim.type).toBe("@cinatra-ai/document-artifact:artifact");
    expect(claim.claim).toBe("dedicated");
    expect(claim.dispositions).toEqual({
      projection: "artifact-safe",
      pinnable: false,
      snapshotPolicy: "none",
      sensitivity: "normal",
    });
    expect(claim.schema).toEqual({ type: "object" });
  });

  it("keeps the typed src manifest in agreement with package.json", () => {
    expect(documentArtifactManifest.accepts).toEqual(pkg.cinatra.artifact.accepts);
    expect(documentArtifactManifest.ui).toEqual(pkg.cinatra.artifact.ui);
  });
});
