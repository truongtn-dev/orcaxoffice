import 'reflect-metadata';
import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { MODULE_METADATA } from '@nestjs/common/constants';
import { AppModule } from '../src/app.module';
import { DatabaseModule } from '../src/database/database.module';
import { HealthModule } from '../src/health/health.module';
import { AuditModule } from '../src/modules/audit/audit.module';
import { MODULE_CATALOG } from '../src/modules/module-catalog';
import { SHELL_MODULES } from '../src/modules/register-shells';
import { ObservabilityModule } from '../src/observability/observability.module';

const IMPLEMENTED = new Set([
  'IdentityModule',
  'OrganizationModule',
  'AuthorizationModule',
  'LicensingModule',
  'ProfileModule',
  'TemplateModule',
  'WordAutomationModule',
  'ExcelAutomationModule',
  'ConfigurationModule',
  'CommerceModule',
  'SupportModule',
]);

describe('module boundaries', () => {
  it('matches the catalog to registered shells', () => {
    assert.deepEqual(
      SHELL_MODULES.map((item) => item.name).sort(),
      MODULE_CATALOG.map((item) => item.className).sort(),
    );
  });

  it('keeps future modules free of providers and controllers', () => {
    for (const shell of SHELL_MODULES) {
      if (IMPLEMENTED.has(shell.name)) {
        continue;
      }
      const controllers = Reflect.getMetadata(MODULE_METADATA.CONTROLLERS, shell) ?? [];
      const providers = Reflect.getMetadata(MODULE_METADATA.PROVIDERS, shell) ?? [];
      assert.deepEqual(controllers, []);
      assert.deepEqual(providers, []);
    }
  });

  it('registers implemented modules with providers', () => {
    for (const shell of SHELL_MODULES) {
      if (!IMPLEMENTED.has(shell.name)) {
        continue;
      }
      const providers = Reflect.getMetadata(MODULE_METADATA.PROVIDERS, shell) ?? [];
      assert.ok(providers.length > 0, `${shell.name} has no providers`);
    }
  });

  it('composes foundation and shell modules in one application', () => {
    const imports = Reflect.getMetadata(MODULE_METADATA.IMPORTS, AppModule) as unknown[];
    for (const expected of [
      DatabaseModule,
      ObservabilityModule,
      AuditModule,
      HealthModule,
      ...SHELL_MODULES,
    ]) {
      assert.equal(imports.includes(expected), true, `${expected.name} is not composed`);
    }
  });
});
