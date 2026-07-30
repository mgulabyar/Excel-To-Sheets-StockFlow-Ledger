
const SHARED_SECRET = "a3f9c2e8b1d4f6a7c9e2b5d8f1a4c7e9b2d5f8a1c4e7b9d2f5a8c1e4b7d9f2a5";

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


export async function sendSecureGet<T = any>(url: string): Promise<SecureResponse<T>> {
  try {
    const response = await fetch(url);
    let data: T | null = null;
    try {
      data = await response.json();
    } catch {

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