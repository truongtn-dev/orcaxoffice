import type { RuleResult } from '@orcax/api-contracts';

export const RULE_PACK_VERSION: string;

export interface RuleDefinition {
  code: string;
  version: string;
  host: 'word' | 'excel';
  title: string;
  description: string;
  checkKind: string;
  defaultSeverity: 'fail' | 'warning';
  checkConfig: Record<string, unknown>;
}

export const CURATED_RULES: readonly RuleDefinition[];

export function assertRuleDefinition(
  input: Partial<RuleDefinition> & {
    code: string;
    version: string;
    host: 'word' | 'excel';
    checkKind: string;
    defaultSeverity: 'fail' | 'warning';
  },
): RuleDefinition;

export function ruleId(code: string, version: string): string;

export interface RuleObservation {
  ruleCode: string;
  ruleVersion?: string;
  passed?: boolean;
  severity?: 'pass' | 'fail' | 'warning';
  defaultSeverity?: 'fail' | 'warning';
  location?: string | null;
  explanation?: string | null;
  correction?: string | null;
}

export interface ValidationReportSummary {
  pass: number;
  fail: number;
  warning: number;
  corrected: number;
}

export interface ValidationReport {
  profileId: string;
  profileVersionId: string | null;
  profileVersion: number;
  profileCode: string | null;
  generatedAt: string;
  summary: ValidationReportSummary;
  results: readonly RuleResult[];
}

export function observationToResult(observation: RuleObservation): RuleResult;

export function buildValidationReport(input: {
  profileId: string;
  profileVersion: number;
  profileVersionId?: string | null;
  profileCode?: string | null;
  generatedAt?: string;
  results: Array<RuleResult | RuleObservation>;
}): ValidationReport;
