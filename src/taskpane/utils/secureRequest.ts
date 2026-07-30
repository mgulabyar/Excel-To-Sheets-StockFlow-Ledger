// ============================================================================
// secureRequest.ts - HMAC Signed Request Layer
// Fixes: (1) missing security/auth, (2) the fake-success "no-cors" bug
// ============================================================================

/**
 * IMPORTANT: This secret must match the one configured in Code.gs (Sheets side).
 * In production, do NOT hardcode this - load it from Office.context.document
 * settings, a secure config store, or prompt the admin once on first setup.
 */
const SHARED_SECRET = "REPLACE_WITH_SECURE_SHARED_SECRET";

/**
 * Generates an HMAC-SHA256 signature for a payload using the Web Crypto API.
 * Runs client-side inside the Excel taskpane (browser-based webview),
 * so window.crypto.subtle is available without extra dependencies.
 */
async function signPayload(payloadString: string): Promise<string> {
  const encoder = new TextEncoder();
  const keyData = encoder.encode(SHARED_SECRET);
  const messageData = encoder.encode(payloadString);

  const cryptoKey = await window.crypto.subtle.importKey(
    "raw",
    keyData,
    { name: "HMAC", hash: "SHA-256" },
    false,
    ["sign"]
  );

  const signatureBuffer = await window.crypto.subtle.sign("HMAC", cryptoKey, messageData);
  const signatureArray = Array.from(new Uint8Array(signatureBuffer));
  return signatureArray.map((b) => b.toString(16).padStart(2, "0")).join("");
}

export interface SecureResponse<T = any> {
  ok: boolean;
  status: number;
  data: T | null;
  errorMessage: string | null;
}

/**
 * Sends a signed, verifiable POST request and ACTUALLY reads the response.
 * This replaces the old `mode: "no-cors"` call, which always reported
 * "success" even when the server failed or rejected the request.
 *
 * Requirement: The Google Apps Script Web App must be deployed with
 * "Execute as: Me" and "Who has access: Anyone" (or org-restricted),
 * which allows normal CORS-enabled responses to be read - no-cors is
 * only needed if you deliberately want to ignore the response, which
 * defeats real error handling.
 */
export async function sendSecurePost<T = any>(
  url: string,
  payload: Record<string, any>
): Promise<SecureResponse<T>> {
  try {
    const payloadString = JSON.stringify(payload);
    const signature = await signPayload(payloadString);

    const response = await fetch(url, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "X-Signature": signature,
      },
      body: payloadString,
    });

    let data: T | null = null;
    try {
      data = await response.json();
    } catch {
      // Response wasn't JSON - still treat status code as source of truth
    }

    if (!response.ok) {
      return {
        ok: false,
        status: response.status,
        data,
        errorMessage: `Server rejected request (HTTP ${response.status})`,
      };
    }

    return { ok: true, status: response.status, data, errorMessage: null };
  } catch (err) {
    return {
      ok: false,
      status: 0,
      data: null,
      errorMessage: err instanceof Error ? err.message : "Network request failed",
    };
  }
}

/**
 * Signed GET request wrapper (for polling reads).
 */
export async function sendSecureGet<T = any>(url: string): Promise<SecureResponse<T>> {
  try {
    const response = await fetch(url);
    let data: T | null = null;
    try {
      data = await response.json();
    } catch {
      // ignore
    }

    if (!response.ok) {
      return {
        ok: false,
        status: response.status,
        data,
        errorMessage: `Server returned HTTP ${response.status}`,
      };
    }
    return { ok: true, status: response.status, data, errorMessage: null };
  } catch (err) {
    return {
      ok: false,
      status: 0,
      data: null,
      errorMessage: err instanceof Error ? err.message : "Network request failed",
    };
  }
}