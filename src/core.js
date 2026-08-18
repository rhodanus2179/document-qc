import { RULES, CATEGORY_LABELS } from './rules-polished.js';
import { runRules as runRuleEngine } from './engine-polished.js';
import { decorateFindings, groupFindingsByRule } from './finding-meta.js';

export { RULES, CATEGORY_LABELS, groupFindingsByRule };

export function runDocumentQc(model, enabledIds = new Set(RULES.map(rule => rule.id)), onProgress = () => {}) {
  const findings = runRuleEngine(model, enabledIds, onProgress);
  return decorateFindings(model, findings);
}
