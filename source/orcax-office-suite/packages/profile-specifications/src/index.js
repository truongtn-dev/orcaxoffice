const PROFILE_PACKAGE_FORMAT = 'orcax.profile.v1';

const PROFILE_STATUSES = Object.freeze(['draft', 'in_review', 'approved', 'deprecated']);
const PROFILE_HOSTS = Object.freeze(['word', 'excel']);
const PROFILE_SCOPES = Object.freeze(['system', 'personal', 'organization']);

function isNonEmptyString(value, max) {
  return typeof value === 'string' && value.trim().length > 0 && value.length <= max;
}

function validateRule(rule, index) {
  if (!rule || typeof rule !== 'object') {
    return `rules[${index}] phải là object.`;
  }
  if (!isNonEmptyString(rule.code, 64)) {
    return `rules[${index}].code không hợp lệ.`;
  }
  if (!isNonEmptyString(rule.version, 32)) {
    return `rules[${index}].version không hợp lệ.`;
  }
  if (!isNonEmptyString(rule.checkKind, 64)) {
    return `rules[${index}].checkKind không hợp lệ.`;
  }
  if (rule.defaultSeverity !== 'fail' && rule.defaultSeverity !== 'warning') {
    return `rules[${index}].defaultSeverity phải là fail hoặc warning.`;
  }
  return null;
}

function validateLegalReference(item, index) {
  if (!item || typeof item !== 'object') {
    return `legalReferences[${index}] phải là object.`;
  }
  if (!isNonEmptyString(item.citation, 500)) {
    return `legalReferences[${index}].citation không hợp lệ.`;
  }
  if (item.verified !== undefined && typeof item.verified !== 'boolean') {
    return `legalReferences[${index}].verified phải là boolean.`;
  }
  return null;
}

function validateProfilePackage(input) {
  const errors = [];
  if (!input || typeof input !== 'object') {
    return { ok: false, errors: ['Gói hồ sơ phải là object JSON.'] };
  }
  if (input.format !== PROFILE_PACKAGE_FORMAT) {
    errors.push(`format phải là ${PROFILE_PACKAGE_FORMAT}.`);
  }
  const profile = input.profile;
  if (!profile || typeof profile !== 'object') {
    errors.push('profile là bắt buộc.');
  } else {
    if (!isNonEmptyString(profile.code, 64)) {
      errors.push('profile.code không hợp lệ.');
    }
    if (!isNonEmptyString(profile.name, 160)) {
      errors.push('profile.name không hợp lệ.');
    }
    if (!PROFILE_HOSTS.includes(profile.host)) {
      errors.push('profile.host phải là word hoặc excel.');
    }
    if (profile.scope && !PROFILE_SCOPES.includes(profile.scope)) {
      errors.push('profile.scope không hợp lệ.');
    }
    if (
      profile.source !== undefined &&
      profile.source !== null &&
      typeof profile.source !== 'string'
    ) {
      errors.push('profile.source phải là chuỗi.');
    }
  }
  const rules = Array.isArray(input.rules) ? input.rules : null;
  if (!rules) {
    errors.push('rules phải là mảng.');
  } else {
    rules.forEach((rule, index) => {
      const message = validateRule(rule, index);
      if (message) {
        errors.push(message);
      }
    });
  }
  const legalReferences = input.legalReferences === undefined ? [] : input.legalReferences;
  if (!Array.isArray(legalReferences)) {
    errors.push('legalReferences phải là mảng.');
  } else {
    legalReferences.forEach((item, index) => {
      const message = validateLegalReference(item, index);
      if (message) {
        errors.push(message);
      }
    });
  }
  return { ok: errors.length === 0, errors };
}

function createProfilePackage(input) {
  const candidate = {
    format: PROFILE_PACKAGE_FORMAT,
    profile: {
      code: input.profile.code,
      name: input.profile.name,
      host: input.profile.host,
      scope: input.profile.scope || undefined,
      source: input.profile.source || null,
      effectiveAt: input.profile.effectiveAt || null,
    },
    rules: (input.rules || []).map((rule) => ({
      code: rule.code,
      version: rule.version,
      title: rule.title || rule.code,
      description: rule.description || '',
      checkKind: rule.checkKind,
      defaultSeverity: rule.defaultSeverity,
      checkConfig: rule.checkConfig || {},
      sortOrder: rule.sortOrder || 0,
    })),
    legalReferences: (input.legalReferences || []).map((item) => ({
      citation: item.citation,
      sourceUrl: item.sourceUrl || null,
      verified: Boolean(item.verified),
      notes: item.notes || null,
    })),
  };
  const validation = validateProfilePackage(candidate);
  if (!validation.ok) {
    const error = new Error(validation.errors.join(' '));
    error.details = validation.errors;
    throw error;
  }
  return candidate;
}

module.exports = {
  PROFILE_PACKAGE_FORMAT,
  PROFILE_STATUSES,
  PROFILE_HOSTS,
  PROFILE_SCOPES,
  validateProfilePackage,
  createProfilePackage,
};
