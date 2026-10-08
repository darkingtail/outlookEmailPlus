import routes from '../../config/routes';
import {
  AUTHENTICATED_ROUTE_PATHS,
  DEFAULT_AUTHENTICATED_REDIRECT,
  getSafeRedirectUrl,
} from './authRedirect';

const ORIGIN = 'https://mail.example.com';

describe('getSafeRedirectUrl', () => {
  it.each([
    null,
    '',
    'https://evil.example.com/accounts',
    '//evil.example.com/accounts',
    '/\\evil.example.com/accounts',
    '/',
    '/dashboard',
    '/user/login',
    '/accounts-extra',
  ])('falls back for an invalid or removed route: %s', (redirect) => {
    expect(getSafeRedirectUrl(redirect, ORIGIN)).toBe(
      DEFAULT_AUTHENTICATED_REDIRECT,
    );
  });

  it.each(AUTHENTICATED_ROUTE_PATHS)(
    'allows the authenticated route %s',
    (route) => {
      expect(getSafeRedirectUrl(route, ORIGIN)).toBe(route);
    },
  );

  it('preserves query parameters and hash for a valid route', () => {
    expect(
      getSafeRedirectUrl('/accounts/?filter=active#account-42', ORIGIN),
    ).toBe('/accounts?filter=active#account-42');
  });

  it('stays in sync with the configured authenticated routes', () => {
    const configuredPaths = routes
      .map((route) => route.path)
      .filter(
        (path): path is string =>
          typeof path === 'string' &&
          path !== '/' &&
          path !== '/user' &&
          path !== '*',
      );

    expect(configuredPaths.map((path) => `/ops${path}`)).toEqual(
      AUTHENTICATED_ROUTE_PATHS,
    );
  });
});
