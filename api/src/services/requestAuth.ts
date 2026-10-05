import type { HttpRequest } from '@azure/functions';
import { ConfigurationError } from '../config.js';

export type UserPrincipal = {
  userId: string;
  userDetails: string;
  roles: string[];
};

export class HttpError extends Error {
  constructor(public readonly status: number, message: string) {
    super(message);
  }
}

export function getPrincipal(request: HttpRequest): UserPrincipal {
  const encoded = request.headers.get('x-ms-client-principal');
  if (!encoded) {
    if (process.env.AZURE_FUNCTIONS_ENVIRONMENT?.toLowerCase() === 'development' && process.env.NODE_ENV !== 'production') {
      return { userId: 'local-dev', userDetails: 'Local Dev', roles: ['admin'] };
    }
    throw new HttpError(401, 'Sign in is required.');
  }
  try {
    const value = JSON.parse(Buffer.from(encoded, 'base64').toString('utf8')) as {
      userId?: string;
      userDetails?: string;
      userRoles?: string[];
      claims?: Array<{ typ?: string; val?: string }>;
    };
    const claims = value.claims ?? [];
    const userDetails = value.userDetails
      ?? claims.find((claim) => /email|preferred_username|upn/i.test(claim.typ ?? ''))?.val
      ?? '';
    if (!userDetails) throw new Error('The authenticated principal has no user identity.');
    return {
      userId: value.userId ?? userDetails,
      userDetails,
      roles: [...new Set([...(value.userRoles ?? []), ...claims.filter((claim) => /role/i.test(claim.typ ?? '')).map((claim) => claim.val ?? '')])],
    };
  } catch {
    throw new HttpError(401, 'The authenticated user context is invalid.');
  }
}

export function requireAdmin(request: HttpRequest): UserPrincipal {
  const principal = getPrincipal(request);
  if (!principal.roles.some((role) => role.toLowerCase() === 'admin')) {
    throw new HttpError(403, 'Administrator access is required.');
  }
  return principal;
}

export function errorResponse(error: unknown) {
  const status = error instanceof HttpError ? error.status : error instanceof ConfigurationError ? 503 : 500;
  const message = status === 500 ? 'The request could not be completed.' : (error as Error).message;
  const code = error instanceof ConfigurationError ? 'SERVICE_NOT_CONFIGURED' : status === 500 ? 'INTERNAL_ERROR' : 'REQUEST_REJECTED';
  return { status, jsonBody: { error: { code, message } } };
}