import { describe, it } from "vitest";
import assert from "node:assert/strict";
import {
  isSharedMicrosoftClient,
  dataverseScope,
  noteSharedClient,
  resolveClientIdChain,
  shouldTryNextClient,
  aadstsCode,
  defaultLoginFlow,
  shouldFallBackToDeviceCode,
  isConditionalAccessBlock,
  conditionalAccessHint,
  assertUsableAuthorizeUrl,
  announceLoginAttempt,
  probeLoopbackRedirect,
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
  const clean = { DATAVERSE_LOAD_CLIENT_ID: undefined, DVLOAD_NO_SHARED_CLIENT: undefined };

  it("tries the shared Microsoft client first, then dvload's own app", () => {
    withEnv(clean, () => {
      assert.deepEqual(resolveClientIdChain(), [SHARED, DVLOAD]);
    });
  });

  // The shared client leads for BOTH flows. Interactive is protected by the
  // loopback preflight probe instead of a different chain order, so the
  // no-admin-consent client keeps first shot everywhere it can work.

  it("an explicit --client-id wins and disables the fallback", () => {
    withEnv(clean, () => {
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

  it("retries on an interactive timeout — the answer is on an Entra page this client caused", () => {
    assert.equal(shouldTryNextClient({ errorCode: "interactive_timeout" }), true);
  });

  it("retries when the loopback preflight says this client can't do browser sign-in", () => {
    assert.equal(shouldTryNextClient({ errorCode: "no_loopback_redirect" }), true);
  });

  it("retries when the app has no reply address registered", () => {
    assert.equal(shouldTryNextClient({ errorMessage: "AADSTS500113: no reply address" }), true);
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

  it("retries on a redirect-URI mismatch, so loopback support is discovered", () => {
    assert.equal(shouldTryNextClient({ errorMessage: "AADSTS50011: redirect URI mismatch" }), true);
  });
});

describe("assertUsableAuthorizeUrl", () => {
  const base = "https://login.microsoftonline.com/organizations/oauth2/v2.0/authorize";

  it("accepts a URL carrying a redirect_uri", () => {
    assert.doesNotThrow(() =>
      assertUsableAuthorizeUrl(`${base}?client_id=x&redirect_uri=http%3A%2F%2Flocalhost%3A5000`)
    );
  });

  it("rejects a URL with no redirect_uri, naming the params present", () => {
    assert.throws(() => assertUsableAuthorizeUrl(`${base}?client_id=x&scope=y`), /AADSTS900971/);
    assert.throws(() => assertUsableAuthorizeUrl(`${base}?client_id=x&scope=y`), /client_id, scope/);
  });

  it("rejects an empty redirect_uri", () => {
    assert.throws(() => assertUsableAuthorizeUrl(`${base}?client_id=x&redirect_uri=`), /no redirect_uri/);
  });

  it("rejects an unparseable URL", () => {
    assert.throws(() => assertUsableAuthorizeUrl("not a url"), /unparseable/);
  });
});

describe("defaultLoginFlow", () => {
  it("prefers the browser on Windows", () => {
    assert.equal(defaultLoginFlow({ platform: "win32" }), "interactive");
  });

  it("prefers the browser on macOS", () => {
    assert.equal(defaultLoginFlow({ platform: "darwin" }), "interactive");
  });

  it("uses device code on Linux with no display server", () => {
    assert.equal(defaultLoginFlow({ platform: "linux" }), "deviceCode");
  });

  it("uses the browser on Linux when X11 is present", () => {
    assert.equal(defaultLoginFlow({ platform: "linux", DISPLAY: ":0" }), "interactive");
  });

  it("uses the browser on Linux under Wayland", () => {
    assert.equal(defaultLoginFlow({ platform: "linux", WAYLAND_DISPLAY: "wayland-0" }), "interactive");
  });

  it("uses device code over SSH, where the browser would open on the wrong machine", () => {
    assert.equal(
      defaultLoginFlow({ platform: "win32", SSH_CONNECTION: "10.0.0.1 22 10.0.0.2 22" }),
      "deviceCode"
    );
  });

  it("honours DVLOAD_AUTH_FLOW=device-code", () => {
    assert.equal(defaultLoginFlow({ platform: "win32", DVLOAD_AUTH_FLOW: "device-code" }), "deviceCode");
  });

  it("honours DVLOAD_AUTH_FLOW=interactive even over SSH", () => {
    assert.equal(
      defaultLoginFlow({
        platform: "linux",
        DVLOAD_AUTH_FLOW: "interactive",
        SSH_CONNECTION: "x",
      }),
      "interactive"
    );
  });
});

describe("isConditionalAccessBlock", () => {
  it("recognises AADSTS53003", () => {
    assert.equal(isConditionalAccessBlock({ errorMessage: "AADSTS53003: blocked" }), true);
  });

  it("does not fire on a plain consent error", () => {
    assert.equal(isConditionalAccessBlock({ errorMessage: "AADSTS65001" }), false);
  });

  it("does not fire when there is no AADSTS code", () => {
    assert.equal(isConditionalAccessBlock(new Error("socket hang up")), false);
  });
});

describe("shouldFallBackToDeviceCode", () => {
  it("falls back when no browser could be launched", () => {
    assert.equal(shouldFallBackToDeviceCode({ errorCode: "browser_launch_failed" }), true);
  });

  it("falls back when the loopback port is taken", () => {
    assert.equal(shouldFallBackToDeviceCode({ code: "EADDRINUSE" }), true);
  });

  it("does NOT fall back on a Conditional Access block", () => {
    // Device code is the *more* restricted flow — retrying with it would
    // replace a clear policy error with a confusing one.
    assert.equal(shouldFallBackToDeviceCode({ errorMessage: "AADSTS53003" }), false);
  });

  it("does NOT fall back when the user declined consent", () => {
    assert.equal(shouldFallBackToDeviceCode({ errorMessage: "AADSTS65004" }), false);
  });
});

describe("conditionalAccessHint", () => {
  it("suggests --interactive when device code was the blocked flow", () => {
    const hint = conditionalAccessHint({ errorMessage: "AADSTS53003" }, "deviceCode");
    assert.match(hint, /--interactive/);
    assert.match(hint, /Authentication Flows/);
    assert.match(hint, /Sign-in logs/);
  });

  it("omits the --interactive suggestion when already interactive", () => {
    const hint = conditionalAccessHint({ errorMessage: "AADSTS53003" }, "interactive");
    assert.doesNotMatch(hint, /--interactive/);
    // Still explains the device-compliance case, which is the likely cause here.
    assert.match(hint, /compliant device/);
  });

  it("names the actual error code", () => {
    assert.match(conditionalAccessHint({ errorMessage: "AADSTS53001" }, "interactive"), /AADSTS53001/);
  });
});

describe("dataverseScope", () => {
  it("appends /.default to the origin", () => {
    assert.equal(dataverseScope("https://org.crm.dynamics.com"), "https://org.crm.dynamics.com/.default");
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
    assert.equal(
      captureStderr(() => noteSharedClient(DVLOAD)),
      ""
    );
  });

  it("is silent for an unknown client id", () => {
    assert.equal(
      captureStderr(() => noteSharedClient("00000000-0000-0000-0000-000000000000")),
      ""
    );
  });
});

describe("announceLoginAttempt", () => {
  const clean = { DATAVERSE_LOAD_CLIENT_ID: undefined, DVLOAD_NO_SHARED_CLIENT: undefined };

  it("always names the client id, even for an unrecognised one", () => {
    withEnv(clean, () => {
      const out = captureStderr(() =>
        announceLoginAttempt("00000000-0000-0000-0000-000000000000", "interactive")
      );
      assert.match(out, /00000000-0000-0000-0000-000000000000/);
      assert.match(out, /browser/);
    });
  });

  it("reveals when DATAVERSE_LOAD_CLIENT_ID pinned the choice", () => {
    // This is the whole point: a stray env var silently disabling the
    // fallback chain must be visible in the output, not inferred.
    withEnv({ DATAVERSE_LOAD_CLIENT_ID: DVLOAD, DVLOAD_NO_SHARED_CLIENT: undefined }, () => {
      const out = captureStderr(() => announceLoginAttempt(DVLOAD, "interactive"));
      assert.match(out, /DATAVERSE_LOAD_CLIENT_ID/);
    });
  });

  it("reveals when DVLOAD_NO_SHARED_CLIENT disabled the shared client", () => {
    withEnv({ DATAVERSE_LOAD_CLIENT_ID: undefined, DVLOAD_NO_SHARED_CLIENT: "1" }, () => {
      const out = captureStderr(() => announceLoginAttempt(DVLOAD, "deviceCode"));
      assert.match(out, /DVLOAD_NO_SHARED_CLIENT/);
      assert.match(out, /device code/);
    });
  });

  it("names the Microsoft app when using the shared client", () => {
    withEnv(clean, () => {
      const out = captureStderr(() => announceLoginAttempt(SHARED, "interactive"));
      assert.match(out, /Microsoft Dynamics CRM/);
    });
  });
});

describe("probeLoopbackRedirect", () => {
  // Minimal fetch stubs. The probe caches definitive answers per client id
  // for the process lifetime, so every test uses its own fake client id.
  const stubResponse = (status: number, headers: Record<string, string>, body = ""): Response =>
    ({
      status,
      headers: { get: (k: string) => headers[k.toLowerCase()] ?? null },
      text: async () => body,
    }) as unknown as Response;

  const fetchReturning = (res: Response): typeof fetch => (async () => res) as unknown as typeof fetch;

  it("a redirect back to localhost means the redirect URI is registered", async () => {
    const res = stubResponse(302, {
      location: "http://localhost:53682/?error=login_required&error_description=...",
    });
    assert.equal(await probeLoopbackRedirect("probe-usable", "organizations", fetchReturning(res)), "usable");
  });

  it("an error page naming AADSTS500113 means it is not registered", async () => {
    const res = stubResponse(200, {}, "<html>AADSTS500113: No reply address is registered</html>");
    assert.equal(
      await probeLoopbackRedirect("probe-500113", "organizations", fetchReturning(res)),
      "unusable"
    );
  });

  it("an error page naming AADSTS900971 means it is not registered", async () => {
    const res = stubResponse(200, {}, "AADSTS900971: No reply address provided.");
    assert.equal(
      await probeLoopbackRedirect("probe-900971", "organizations", fetchReturning(res)),
      "unusable"
    );
  });

  it("a redirect to somewhere other than localhost is inconclusive", async () => {
    const res = stubResponse(302, { location: "https://login.live.com/elsewhere" });
    assert.equal(
      await probeLoopbackRedirect("probe-elsewhere", "organizations", fetchReturning(res)),
      "inconclusive"
    );
  });

  it("a network failure is inconclusive — never veto a sign-in over a flaky proxy", async () => {
    const failing = (async () => {
      throw new Error("getaddrinfo ENOTFOUND login.microsoftonline.com");
    }) as unknown as typeof fetch;
    assert.equal(await probeLoopbackRedirect("probe-offline", "organizations", failing), "inconclusive");
  });

  it("an unrecognisable error page is inconclusive, not unusable", async () => {
    const res = stubResponse(400, {}, "<html>Something unrelated went wrong</html>");
    assert.equal(
      await probeLoopbackRedirect("probe-unknown", "organizations", fetchReturning(res)),
      "inconclusive"
    );
  });

  it("caches a definitive answer per client id", async () => {
    const res = stubResponse(200, {}, "AADSTS500113");
    assert.equal(await probeLoopbackRedirect("probe-cached", "organizations", fetchReturning(res)), "unusable");
    const explode = (async () => {
      throw new Error("should not be called — answer was cached");
    }) as unknown as typeof fetch;
    assert.equal(await probeLoopbackRedirect("probe-cached", "organizations", explode), "unusable");
  });
});
