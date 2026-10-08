export const DEFAULT_AUTHENTICATED_REDIRECT = '/ops/overview';

export const AUTHENTICATED_ROUTE_PATHS = [
  '/ops/overview',
  '/ops/mailbox',
  '/ops/accounts',
  '/ops/groups',
  '/ops/temp-emails',
  '/ops/pool-admin',
  '/ops/plugins',
  '/ops/refresh-log',
  '/ops/settings',
  '/ops/audit',
  '/ops/token-tool',
] as const;

const AUTHENTICATED_ROUTE_PATH_SET = new Set<string>(
  AUTHENTICATED_ROUTE_PATHS,
);

export const getSafeRedirectUrl = (
  redirect: string | null,
  origin: string,
): string => {
  if (!redirect?.startsWith('/') || redirect.startsWith('//')) {
    return DEFAULT_AUTHENTICATED_REDIRECT;
  }

  try {
    const parsed = new URL(redirect, origin);
    if (parsed.origin !== origin) return DEFAULT_AUTHENTICATED_REDIRECT;

    const pathname =
      parsed.pathname.length > 1
        ? parsed.pathname.replace(/\/+$/, '')
        : parsed.pathname;
    if (!AUTHENTICATED_ROUTE_PATH_SET.has(pathname)) {
      return DEFAULT_AUTHENTICATED_REDIRECT;
    }

    return `${pathname}${parsed.search}${parsed.hash}`;
  } catch {
    return DEFAULT_AUTHENTICATED_REDIRECT;
  }
};
