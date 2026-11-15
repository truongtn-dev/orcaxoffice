import { AnalyticsModule } from './analytics/analytics.module';
import { AssetModule } from './asset/asset.module';
import { AuthorizationModule } from './authorization/authorization.module';
import { CollaborationModule } from './collaboration/collaboration.module';
import { CommerceModule } from './commerce/commerce.module';
import { ConfigurationModule } from './configuration/configuration.module';
import { DocumentModule } from './document/document.module';
import { ExcelAutomationModule } from './excel-automation/excel-automation.module';
import { FormModule } from './form/form.module';
import { IdentityModule } from './identity/identity.module';
import { IntegrationModule } from './integration/integration.module';
import { KnowledgeModule } from './knowledge/knowledge.module';
import { LicensingModule } from './licensing/licensing.module';
import { NotificationModule } from './notification/notification.module';
import { OrganizationModule } from './organization/organization.module';
import { ProfileModule } from './profile/profile.module';
import { ReportingModule } from './reporting/reporting.module';
import { SearchModule } from './search/search.module';
import { SupportModule } from './support/support.module';
import { TaskModule } from './task/task.module';
import { TemplateModule } from './template/template.module';
import { WordAutomationModule } from './word-automation/word-automation.module';
import { WorkflowModule } from './workflow/workflow.module';

export const SHELL_MODULES = [
  IdentityModule,
  OrganizationModule,
  AuthorizationModule,
  LicensingModule,
  CommerceModule,
  ProfileModule,
  TemplateModule,
  WordAutomationModule,
  ExcelAutomationModule,
  DocumentModule,
  AssetModule,
  WorkflowModule,
  FormModule,
  TaskModule,
  CollaborationModule,
  KnowledgeModule,
  SearchModule,
  ReportingModule,
  AnalyticsModule,
  NotificationModule,
  SupportModule,
  IntegrationModule,
  ConfigurationModule,
];
