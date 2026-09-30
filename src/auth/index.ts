import type { AuthProvider } from '../contracts.js';

export { createAuthProvider, describeAccount } from './provider.js';
export { cachePaths, clearCache, createCachePlugin } from './token-cache.js';

/**
 * An `AuthProvider` that refuses to do anything.
 *
 * Two CLI paths — `status` and `permissions` — build the tool catalogue only to
 * read the scopes off it, and have no session to hand over. `ToolDeps.auth` is
 * deliberately required rather than optional (a missing provider used to remove
 * the sign-in tools silently), so those paths need something to pass. Throwing
 * is the honest thing for it to do: if a call ever reaches here, a code path
 * that only meant to enumerate scopes has started using a session it does not
 * have, and a thrown error says so where a null would not.
 */
export function scopeReadingOnlyAuth(): AuthProvider {
  // Rejected promises, not synchronous throws. Every method on AuthProvider is
  // declared to return one, and a caller that writes `auth.getToken(x).catch(…)`
  // would otherwise get an exception thrown at it before a promise ever exists.
  const refuse = async (what: string): Promise<never> => {
    throw new Error(
      `${what} is not available here: this catalogue was built to read tool scopes, not to use a session.`,
    );
  };
  return {
    getToken: async () => refuse('Acquiring a token'),
    getAccount: async () => refuse('Reading the account'),
    login: async () => refuse('Signing in'),
    logout: async () => refuse('Signing out'),
    beginDeviceLogin: async () => refuse('Starting a sign-in'),
    deviceLoginStatus: async () => refuse('Reading sign-in status'),
  };
}
