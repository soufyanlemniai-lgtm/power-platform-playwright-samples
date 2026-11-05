// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { executeInFormContext, getFormContext } from 'power-platform-playwright-toolkit';
import { test, expect } from '../../../fixtures/mda.fixtures';
import { getModelDrivenSignInState } from '../../../utils/model-driven-auth';

const CONTACT_FORM_ID = 'c3d4439c-0be4-4231-8ca3-11444e4cf585';

interface ContactLabelContext {
  Page: {
    getControl(name: string): { getLabel(): string } | null;
    ui: { getFormType(): number };
  };
  Utility: {
    getGlobalContext(): { userSettings: { languageId: number } };
  };
}

test.describe('INES - model-driven smoke tests', () => {
  test.use({ modelDrivenAppName: 'INES 2.0' });

  test('opens the configured INES app with an authenticated user context', async ({
    page,
    modelDrivenApp,
  }) => {
    const configuredUrl = new URL(modelDrivenApp.getBaseAppUrl());
    const appId = configuredUrl.searchParams.get('appid');
    expect(appId, 'The configured model-driven URL must include appid').toBeTruthy();
    await expect(page).toHaveURL(
      (url) =>
        url.origin === configuredUrl.origin &&
        url.searchParams.get('appid')?.toLowerCase() === appId?.toLowerCase()
    );
    const state = await page.waitForFunction(getModelDrivenSignInState, configuredUrl.origin, {
      timeout: 60000,
    });
    const status = await state.jsonValue();
    await state.dispose();
    expect(status, 'The INES app must expose an authenticated user context').toBe('ready');
  });

  test('opens the Contacts grid without changing records', async ({ page, modelDrivenApp }) => {
    await modelDrivenApp.navigateToGridView('contact');
    await expect(page).toHaveURL(
      (url) =>
        url.searchParams.get('pagetype') === 'entitylist' &&
        url.searchParams.get('etn') === 'contact'
    );
  });

  test('shows the expected fields and Salutations label on an unsaved Contact form', async ({
    page,
    modelDrivenApp,
  }) => {
    await modelDrivenApp.navigateToFormView('contact', { formId: CONTACT_FORM_ID });
    const form = await getFormContext(page);
    expect(form.entityName).toBe('contact');
    expect(form.entityId, 'The smoke test must remain on an unsaved Contact draft').toBe('');
    expect(form.attributeNames).toEqual(
      expect.arrayContaining([
        'gendercode',
        'salutation',
        'lastname',
        'firstname',
        'birthdate',
        'etn_niss',
      ])
    );

    const caption = await executeInFormContext(page, (Xrm: ContactLabelContext) => {
      const control = Xrm.Page.getControl('gendercode');
      if (!control) {
        throw new Error('The INES Contact form is missing the gendercode control.');
      }
      return {
        label: control.getLabel(),
        languageId: Xrm.Utility.getGlobalContext().userSettings.languageId,
        formType: Xrm.Page.ui.getFormType(),
      };
    });
    expect(caption.formType, 'The Contact form must be in create mode without saving').toBe(1);
    expect(caption.label).toBe(caption.languageId === 1033 ? 'Salutation' : 'Salutations');
  });
});
