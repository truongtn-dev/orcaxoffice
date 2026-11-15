# Rule definitions

Versioned deterministic rule packs for Word and Excel validation. Rules are explicit data: `code`, `version`, and a `checkKind` with config. They do not call generative models.

`buildValidationReport` stamps the profile version used for a run so results stay reproducible. Office Object Model execution stays in the add-in; this package only shapes contracts and curated packs.
