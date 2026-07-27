import { describe, it } from "node:test";
import assert from "node:assert/strict";
import {
  isSharedMicrosoftClient,
  dataverseScope,
  noteSharedClient,
  resolveClientIdChain,
  shouldTryNextClient,
  aadstsCode,
} from "./auth.js";

const SHARED = "51f81489-12ee-4a9e-aaae-a2591f45987d";
const DVLOAD = "e6828b0f-9fde-43f8-85d0-602660d498bb";

/** Run `fn` with the given env vars applied, then restore them. */
function withEnv(vars: Record<string, string | undefined>, fn: () => void): void {
  const saved: Record<string, string | undefined> = {};
  for (const k of Object.keys(vars)) {
    saved[k] = process.env[k];
    if (vars[k] === undefined) delete process.env[k];
    else process.env[k] = vars[k]!;
  }
  try {
    fn();
  } finally {
    for (const k of Object.keys(saved)) {
      if (saved[k] === undefined) delete process.env[k];
      else process.env[k] = saved[k]!;
    }
  }
}

describe("isSharedMicrosoftClient", () => {
  it("returns true for the PowerApps client id", () => {
    assert.equal(isSharedMicrosoftClient("2ad88395-b77d-4561-9441-d0e40824f9bc"), true);
  });

  it("returns true for the Dynamics CRM / XRM Tooling client id", () => {
    assert.equal(isSharedMicrosoftClient(SHARED), true);
  });

  it("returns false for dvload's own app", () => {
    assert.equal(isSharedMicrosoftClient(DVLOAD), false);
  });

  it("returns false for an unknown client id", () => {
    assert.equal(isSharedMicrosoftClient("00000000-0000-0000-0000-000000000000"), false);
  });

  it("returns false for an empty string", () => {
    assert.equal(isSharedMicrosoftClient(""), false);
  });
});

describe("resolveClientIdChain", () => {
  it("defaults to the shared Microsoft client, then dvload's own app", () => {
    withEnv({ DATAVERSE_LOAD_CLIENT_ID: undefined, DVLOAD_NO_SHARED_CLIENT: undefined }, () => {
      assert.deepEqual(resolveClientIdChain(), [SHARED, DVLOAD]);
    });
  });

  it("an explicit --client-id wins and disables the fallback", () => {
    withEnv({ DATAVERSE_LOAD_CLIENT_ID: undefined, DVLOAD_NO_SHARED_CLIENT: undefined }, () => {
      assert.deepEqual(resolveClientIdChain("abc"), ["abc"]);
    });
  });

  it("a pinned env var wins and disables the fallback", () => {
    withEnv({ DATAVERSE_LOAD_CLIENT_ID: "pinned", DVLOAD_NO_SHARED_CLIENT: undefined }, () => {
      assert.deepEqual(resolveClientIdChain(), ["pinned"]);
    });
  });

  it("--client-id beats the env var", () => {
    withEnv({ DATAVERSE_LOAD_CLIENT_ID: "pinned" }, () => {
      assert.deepEqual(resolveClientIdChain("explicit"), ["explicit"]);
    });
  });

  it("DVLOAD_NO_SHARED_CLIENT=1 drops the shared client", () => {
    withEnv({ DATAVERSE_LOAD_CLIENT_ID: undefined, DVLOAD_NO_SHARED_CLIENT: "1" }, () => {
      assert.deepEqual(resolveClientIdChain(), [DVLOAD]);
    });
  });
});

describe("aadstsCode", () => {
  it("extracts the code from an MSAL errorMessage", () => {
    assert.equal(
      aadstsCode({ errorMessage: "AADSTS65001: The user or administrator has not consented" }),
      "AADSTS65001"
    );
  });

  it("falls back to the Error message", () => {
    assert.equal(aadstsCode(new Error("boom AADSTS53003 blocked")), "AADSTS53003");
  });

  it("returns null when there is no code", () => {
    assert.equal(aadstsCode(new Error("network unreachable")), null);
  });
});

describe("shouldTryNextClient", () => {
  it("retries when consent is missing for this client", () => {
    assert.equal(shouldTryNextClient({ errorMessage: "AADSTS65001: no consent" }), true);
  });

  it("retries when admin approval is required", () => {
    assert.equal(shouldTryNextClient({ errorMessage: "AADSTS90094: admin consent" }), true);
  });

  it("retries when public client flows are disabled", () => {
    assert.equal(shouldTryNextClient({ errorMessage: "AADSTS7000218" }), true);
  });

  it("does NOT retry when Conditional Access blocked the sign-in", () => {
    assert.equal(shouldTryNextClient({ errorMessage: "AADSTS53003: blocked by CA" }), false);
  });

  it("does NOT retry when the user declined consent", () => {
    assert.equal(shouldTryNextClient({ errorMessage: "AADSTS65004: user declined" }), false);
  });

  it("does NOT retry when MFA is required", () => {
    assert.equal(shouldTryNextClient({ errorMessage: "AADSTS50076" }), false);
  });

  it("does NOT retry when the device code expired", () => {
    assert.equal(shouldTryNextClient({ errorCode: "device_code_expired" }), false);
  });

  it("does NOT retry on an unrecognised failure", () => {
    assert.equal(shouldTryNextClient(new Error("socket hang up")), false);
  });

  it("retries on a bare invalid_client error code", () => {
    assert.equal(shouldTryNextClient({ errorCode: "invalid_client" }), true);
  });
});

describe("dataverseScope", () => {
  it("appends /.default to the origin", () => {
    assert.equal(
      dataverseScope("https://org.crm.dynamics.com"),
      "https://org.crm.dynamics.com/.default"
    );
  });

  it("uses origin only, ignoring any path segment", () => {
    assert.equal(
      dataverseScope("https://org.crm.dynamics.com/some/path"),
      "https://org.crm.dynamics.com/.default"
    );
  });
});

/** Capture everything written to stderr while `fn` runs. */
function captureStderr(fn: () => void): string {
  const captured: string[] = [];
  const origWrite = process.stderr.write.bind(process.stderr);
  (process.stderr as any).write = (msg: string | Uint8Array) => {
    captured.push(String(msg));
    return true;
  };
  try {
    fn();
  } finally {
    process.stderr.write = origWrite;
  }
  return captured.join("");
}

describe("noteSharedClient", () => {
  it("names the Microsoft app the user will actually see", () => {
    const output = captureStderr(() => noteSharedClient("2ad88395-b77d-4561-9441-d0e40824f9bc"));
    assert.ok(output.includes("Microsoft PowerApps"), "should name the shared client");
    assert.ok(output.includes("DVLOAD_NO_SHARED_CLIENT"), "should say how to opt out");
  });

  it("is silent for dvload's own app", () => {
    assert.equal(captureStderr(() => noteSharedClient(DVLOAD)), "");
  });

  it("is silent for an unknown client id", () => {
    assert.equal(
      captureStderr(() => noteSharedClient("00000000-0000-0000-0000-000000000000")),
      ""
    );
  });
});
