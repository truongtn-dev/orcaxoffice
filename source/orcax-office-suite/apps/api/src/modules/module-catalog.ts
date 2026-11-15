export interface ModuleBoundary {
  key: string;
  className: string;
  epic: string;
  responsibility: string;
}

export const MODULE_CATALOG: readonly ModuleBoundary[] = [
  {
    key: 'identity',
    className: 'IdentityModule',
    epic: 'E02',
    responsibility: 'Registration, login, email verification, password reset, and sessions.',
  },
  {
    key: 'organization',
    className: 'OrganizationModule',
    epic: 'E02',
    responsibility: 'Personal and organization workspaces, memberships, and invitations.',
  },
  {
    key: 'authorization',
    className: 'AuthorizationModule',
    epic: 'E03',
    responsibility:
      'RBAC, ABAC policy checks, permission guards, and workspace-scoped audit access.',
  },
  {
    key: 'licensing',
    className: 'LicensingModule',
    epic: 'E04',
    responsibility: 'Plans, entitlements, activation, and device limits.',
  },
  {
    key: 'commerce',
    className: 'CommerceModule',
    epic: 'E09',
    responsibility: 'Orders, payments, refunds, and reconciliation.',
  },
  {
    key: 'profile',
    className: 'ProfileModule',
    epic: 'E05',
    responsibility: 'Governed profiles, versions, and legal references.',
  },
  {
    key: 'template',
    className: 'TemplateModule',
    epic: 'E05',
    responsibility: 'Template lifecycle and deterministic field mapping.',
  },
  {
    key: 'word-automation',
    className: 'WordAutomationModule',
    epic: 'E06',
    responsibility:
      'Word bootstrap, capability matrix, and validation-run metadata. Document execution stays in the add-in.',
  },
  {
    key: 'excel-automation',
    className: 'ExcelAutomationModule',
    epic: 'E07',
    responsibility:
      'Excel bootstrap, print-run metadata, and PDF capability notes. Workbook execution stays in the add-in.',
  },
  {
    key: 'document',
    className: 'DocumentModule',
    epic: 'later',
    responsibility: 'Document metadata and explicit file references. No default body upload.',
  },
  {
    key: 'asset',
    className: 'AssetModule',
    epic: 'later',
    responsibility: 'Digital asset metadata and usage rights.',
  },
  {
    key: 'workflow',
    className: 'WorkflowModule',
    epic: 'later',
    responsibility: 'Deterministic workflow definitions and runs.',
  },
  {
    key: 'form',
    className: 'FormModule',
    epic: 'later',
    responsibility: 'Dynamic form definitions and submissions.',
  },
  {
    key: 'task',
    className: 'TaskModule',
    epic: 'later',
    responsibility: 'Tasks linked to records the caller is allowed to see.',
  },
  {
    key: 'collaboration',
    className: 'CollaborationModule',
    epic: 'later',
    responsibility: 'Comments, attachments, and activity history.',
  },
  {
    key: 'knowledge',
    className: 'KnowledgeModule',
    epic: 'later',
    responsibility: 'Knowledge items and taxonomy.',
  },
  {
    key: 'search',
    className: 'SearchModule',
    epic: 'later',
    responsibility: 'Permission-filtered metadata and full-text search. No semantic AI search.',
  },
  {
    key: 'reporting',
    className: 'ReportingModule',
    epic: 'later',
    responsibility: 'Deterministic reports scoped by permission.',
  },
  {
    key: 'analytics',
    className: 'AnalyticsModule',
    epic: 'later',
    responsibility: 'Deterministic workspace metrics. No AI personalization.',
  },
  {
    key: 'notification',
    className: 'NotificationModule',
    epic: 'later',
    responsibility: 'In-app and email notification records.',
  },
  {
    key: 'support',
    className: 'SupportModule',
    epic: 'E09',
    responsibility: 'Support tickets and help content.',
  },
  {
    key: 'integration',
    className: 'IntegrationModule',
    epic: 'later',
    responsibility: 'Adapters for email, payment, and signature providers.',
  },
  {
    key: 'configuration',
    className: 'ConfigurationModule',
    epic: 'E08',
    responsibility:
      'Workspace settings with version history and permission-aware admin overview widgets.',
  },
];
