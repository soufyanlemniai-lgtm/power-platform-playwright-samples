// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { setTimeout as delay } from 'node:timers/promises';
import { chromium } from '@playwright/test';
import { getModelDrivenSignInState, isModelDrivenCookie } from '../utils/model-driven-auth';

function withBrowserState(
  origin: string,
  pathname: string,
  text: string,
  userId: string | undefined,
  action: () => void
) {
  const windowDescriptor = Object.getOwnPropertyDescriptor(globalThis, 'window');
  const documentDescriptor = Object.getOwnPropertyDescriptor(globalThis, 'document');
  Object.defineProperty(globalThis, 'window', {
    configurable: true,
    value: {
      location: { origin, pathname },
      Xrm:
        userId === undefined
          ? undefined
          : {
              Utility: { getGlobalContext: () => ({ userSettings: { userId } }) },
            },
    },
  });
  Object.defineProperty(globalThis, 'document', {
    configurable: true,
    value: { body: { innerText: text } },
  });
  try {
    action();
  } finally {
    if (windowDescriptor) {
      Object.defineProperty(globalThis, 'window', windowDescriptor);
    } else {
      Reflect.deleteProperty(globalThis, 'window');
    }
    if (documentDescriptor) {
      Object.defineProperty(globalThis, 'document', documentDescriptor);
    } else {
      Reflect.deleteProperty(globalThis, 'document');
    }
  }
}

const origin = 'https://example.crm4.dynamics.com';

test('Microsoft sign-in and MFA are not treated as completed authentication', () => {
  withBrowserState('https://login.microsoftonline.com', '/login', '', undefined, () => {
    assert.equal(getModelDrivenSignInState(origin), false);
  });
});

test('the CRM redirect is not sufficient while the app context is still missing', () => {
  withBrowserState(origin, '/main.aspx', '', undefined, () => {
    assert.equal(getModelDrivenSignInState(origin), false);
  });
});

test('an empty user context is not treated as signed in', () => {
  withBrowserState(origin, '/main.aspx', '', '', () => {
    assert.equal(getModelDrivenSignInState(origin), false);
  });
});

test('the authenticated user context on the expected origin completes sign-in', () => {
  withBrowserState(origin, '/main.aspx', '', 'test-user', () => {
    assert.equal(getModelDrivenSignInState(origin), 'ready');
  });
});

test('an authenticated context on a different CRM origin does not complete sign-in', () => {
  withBrowserState('https://other.crm4.dynamics.com', '/main.aspx', '', 'test-user', () => {
    assert.equal(getModelDrivenSignInState(origin), false);
  });
});

test('the CRM error route is reported even without an English error message', () => {
  withBrowserState(origin, '/errorhandler.aspx', '', undefined, () => {
    assert.equal(getModelDrivenSignInState(origin), 'error');
  });
});

test('the existing English app error is reported instead of saving a session', () => {
  withBrowserState(origin, '/main.aspx', 'An error has occurred', undefined, () => {
    assert.equal(getModelDrivenSignInState(origin), 'error');
  });
});

test('CRM session cookies are recognized in default and regional environments', () => {
  for (const domain of [
    'example.crm.dynamics.com',
    '.example.crm4.dynamics.com',
    'example.crm11.dynamics.com',
    '.crm4.dynamics.com',
    'EXAMPLE.CRM4.DYNAMICS.COM',
  ]) {
    assert.equal(isModelDrivenCookie({ domain }), true, domain);
  }
});

test('login, unrelated and misleading cookie domains do not establish a CRM session', () => {
  for (const domain of [
    'login.microsoftonline.com',
    'make.powerapps.com',
    'example.crm4.dynamics.com.invalid',
    'notcrm4.dynamics.com',
    '',
  ]) {
    assert.equal(isModelDrivenCookie({ domain }), false, domain);
  }
});

test('malformed cookie records cannot establish a CRM session', () => {
  for (const value of [undefined, null, {}, { domain: 42 }, 'example.crm4.dynamics.com']) {
    assert.equal(isModelDrivenCookie(value), false);
  }
});

test(
  'browser polling stays pending past five seconds until the app user context is ready',
  { timeout: 25000 },
  async () => {
    const browser = await chromium.launch({ channel: 'msedge', headless: true });
    try {
      const page = await browser.newPage();
      await page.setContent('<p>Sign-in pending</p>');
      const pending = page.waitForFunction(getModelDrivenSignInState, 'null', {
        timeout: 15000,
        polling: 50,
      });
      const completedEarly = await Promise.race([pending.then(() => true), delay(5500, false)]);
      assert.equal(completedEarly, false);
      assert.equal(browser.isConnected(), true);
      await page.addScriptTag({
        content:
          "window.Xrm = { Utility: { getGlobalContext: () => ({ userSettings: { userId: 'test-user' } }) } };",
      });
      const result = await pending;
      assert.equal(await result.jsonValue(), 'ready');
      await result.dispose();
    } finally {
      await browser.close();
    }
  }
);
