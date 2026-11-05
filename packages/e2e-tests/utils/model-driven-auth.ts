// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

export function getModelDrivenSignInState(expectedOrigin: string): 'ready' | 'error' | false {
  if (window.location.origin !== expectedOrigin) {
    return false;
  }

  if (
    window.location.pathname.toLowerCase().endsWith('/errorhandler.aspx') ||
    document.body?.innerText.includes('An error has occurred')
  ) {
    return 'error';
  }

  const appWindow = window as Window & {
    Xrm?: {
      Utility?: {
        getGlobalContext?: () => { userSettings?: { userId?: string } };
      };
    };
  };
  const userId = appWindow.Xrm?.Utility?.getGlobalContext?.()?.userSettings?.userId;
  return typeof userId === 'string' && userId.length > 0 ? 'ready' : false;
}

export function isModelDrivenCookie(cookie: unknown): boolean {
  if (
    !cookie ||
    typeof cookie !== 'object' ||
    !('domain' in cookie) ||
    typeof cookie.domain !== 'string'
  ) {
    return false;
  }

  return /^(?:[a-z0-9-]+\.)*crm\d*\.dynamics\.com$/i.test(cookie.domain.replace(/^\./, ''));
}
