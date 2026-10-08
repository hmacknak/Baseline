import { describe, expect, it } from "vitest";
import { isUpdate } from "../src/update";

describe("isUpdate", () => {
  it("offers an update when a different build is live", () => {
    expect(isUpdate("9ac3b7a", "1b2c3d4")).toBe(true);
  });

  it("does nothing when already on the latest build", () => {
    expect(isUpdate("9ac3b7a", "9ac3b7a")).toBe(false);
  });

  it("never updates development builds or on a missing version", () => {
    expect(isUpdate("dev", "1b2c3d4")).toBe(false);
    expect(isUpdate("9ac3b7a", "dev")).toBe(false);
    expect(isUpdate("9ac3b7a", null)).toBe(false);
    expect(isUpdate("9ac3b7a", "")).toBe(false);
  });
});
