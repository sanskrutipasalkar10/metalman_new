/**
 * Central API configuration.
 * In production (Vercel), VITE_API_URL is set to the Render backend URL.
 * In a single-VM Docker deployment (Nginx reverse-proxying /api to the
 * backend on the same origin), VITE_API_URL is built as an explicit empty
 * string so requests stay relative - `??` (not `||`) preserves that empty
 * string instead of falling back to the dev default below.
 * In development, it falls back to localhost:8001.
 */
export const API_BASE = import.meta.env.VITE_API_URL ?? "http://127.0.0.1:8001";

/**
 * Retry options for fetchWithRetry.
 */
interface RetryOptions {
  maxRetries?: number;      // default: 5
  initialDelay?: number;    // ms, default: 2000
  maxDelay?: number;        // ms, default: 30000
  onRetry?: (attempt: number, delay: number) => void;
}

/**
 * A fetch wrapper with exponential backoff retry logic.
 * Retries on network errors (cold-start / CORS-due-to-502) and on
 * HTTP 502 / 503 / 504 gateway errors from Render's load balancer.
 *
 * Usage:  const resp = await fetchWithRetry(url, fetchOptions, retryOptions);
 */
export async function fetchWithRetry(
  url: string,
  options?: RequestInit,
  retryOptions?: RetryOptions
): Promise<Response> {
  const {
    maxRetries = 5,
    initialDelay = 2000,
    maxDelay = 30_000,
    onRetry,
  } = retryOptions ?? {};

  let lastError: unknown;

  for (let attempt = 0; attempt <= maxRetries; attempt++) {
    try {
      const resp = await fetch(url, options);

      // Retry on gateway/server-unavailable responses
      if ([502, 503, 504].includes(resp.status) && attempt < maxRetries) {
        const delay = Math.min(initialDelay * 2 ** attempt, maxDelay);
        onRetry?.(attempt + 1, delay);
        await sleep(delay);
        continue;
      }

      return resp;
    } catch (err) {
      // Network error (CORS-due-to-502, ERR_FAILED, etc.) — retry
      lastError = err;
      if (attempt < maxRetries) {
        const delay = Math.min(initialDelay * 2 ** attempt, maxDelay);
        onRetry?.(attempt + 1, delay);
        await sleep(delay);
      }
    }
  }

  throw lastError;
}

/** Utility: resolves after `ms` milliseconds. */
function sleep(ms: number) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

