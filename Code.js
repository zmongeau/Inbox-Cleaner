/**
 * ============================================================================
 * INBOX CLEANER — Gmail Add-on
 * ============================================================================
 *
 * Automatically organizes incoming email by matching senders (or domains)
 * to Gmail labels. Rules are stored in a JSON config file (`mail-rules.json`)
 * that lives in the same Drive folder as this script.
 *
 * Workflow:
 *   1. DISCOVER  — `discoverRules()` scans existing labels to build rules.
 *   2. CLEAN     — `runGlobalCleanup()` sweeps the inbox and files messages.
 *   3. MANUAL    — `addNewRule()` lets you add one-off rules from the UI.
 *
 * Extended Features:
 *   - Quick File   — Pre-populates sender from the currently viewed email.
 *   - Dry Run      — Preview cleanup results without moving anything.
 *   - Exclusions   — Whitelist senders/domains to never auto-file.
 *   - Keywords     — Match subject-line keywords in addition to senders.
 *   - Filing Stats — Track how many threads each rule has filed.
 *   - Triggers     — Manage automated schedules from the UI.
 *   - Import/Export — Backup and restore rules via Drive files.
 *
 * The add-on UI (`buildAddOn`) exposes these actions plus label migration
 * and color-coding utilities.
 * ============================================================================
 */

// ── Configuration ───────────────────────────────────────────────────────────

/** Name of the JSON rules file stored alongside this script in Drive. */
const CONFIG_FILENAME = 'mail-rules.json';

/** Max threads to inspect per label during discovery. */
const THREAD_LIMIT_PER_LABEL = 30;

/** Maximum execution time (ms) before discovery saves progress and exits. */
const MAX_DISCOVERY_TIME_MS = 5 * 60 * 1000; // 5 minutes

/** Auto-consolidate rules after adding new ones (can be disabled). */
const AUTO_CONSOLIDATE = false;

/*
 * Exclusions, keyword rules, and filing stats are stored in UserProperties
 * (via PropertiesService) to keep the main config file clean and
 * backward-compatible. No additional OAuth scopes required.
 */


// ── Private Helpers ─────────────────────────────────────────────────────────

/**
 * Extracts a bare email address from a formatted sender string.
 *
 * @param {string} senderString - e.g. "Sender Name <sender@email.com>"
 * @returns {string} The bare email, e.g. "sender@email.com". Returns the
 *   original string unchanged if no angle-bracket format is detected.
 */
function extractEmail_(senderString) {
  const match = senderString.match(/<(.+)>/);
  return match ? match[1] : senderString;
}

/**
 * Returns the Gmail label with the given name, creating it if necessary.
 *
 * @param {string} labelName - Label name (may include "/" for sub-labels).
 * @returns {GmailApp.GmailLabel} The existing or newly created label.
 */
function getOrCreateLabel_(labelName) {
  return GmailApp.getUserLabelByName(labelName) || GmailApp.createLabel(labelName);
}

/**
 * Applies a label to a single thread and archives it out of the inbox.
 * Optionally tracks filing stats when a ruleKey is provided.
 *
 * @param {GmailApp.GmailThread} thread - The thread to file.
 * @param {string} labelName - Destination label name.
 * @param {string} [ruleKey] - The rule that triggered this filing (for stats).
 */
function applyLabelToThread_(thread, labelName, ruleKey) {
  const label = getOrCreateLabel_(labelName);
  label.addToThread(thread);
  GmailApp.moveThreadToArchive(thread);
  if (ruleKey) incrementFilingStat_(ruleKey, 1);
  Logger.log('Filed thread to: ' + labelName);
}

/**
 * Applies a label to multiple threads in bulk and archives them.
 * Optionally tracks filing stats when a ruleKey is provided.
 *
 * @param {GmailApp.GmailThread[]} threads - Array of threads to file.
 * @param {string} labelName - Destination label name.
 * @param {string} [ruleKey] - The rule that triggered this filing (for stats).
 */
function applyLabelToThreads_(threads, labelName, ruleKey) {
  const label = getOrCreateLabel_(labelName);
  label.addToThreads(threads);
  GmailApp.moveThreadsToArchive(threads);
  if (ruleKey) incrementFilingStat_(ruleKey, threads.length);
  Logger.log('Filed ' + threads.length + ' threads to: ' + labelName);
}

/**
 * Builds a CardService notification response (convenience wrapper).
 *
 * @param {string} message - Text shown in the Gmail notification toast.
 * @returns {CardService.ActionResponse}
 */
function notify_(message) {
  return CardService.newActionResponseBuilder()
    .setNotification(CardService.newNotification().setText(message))
    .build();
}

/**
 * Resolves which rule (if any) matches a given email address and subject.
 * Checks exclusions first, then exact sender match, domain wildcard, and
 * finally keyword rules on the subject line.
 *
 * @param {Object<string,string>} config - Sender rules map.
 * @param {string[]} exclusions - Array of excluded senders/domains.
 * @param {Object<string,string>} keywordRules - Keyword → label map.
 * @param {string} email - Sender email address.
 * @param {string} subject - Message subject line.
 * @returns {Object|null} { label, ruleKey } or null if no match.
 */
function resolveRule_(config, exclusions, keywordRules, email, subject) {
  // Check exclusions first — both exact email and domain
  if (isExcludedFromList_(exclusions, email)) return null;

  // Priority 1: exact sender match (e.g. "john@acme.com")
  if (config[email]) {
    return { label: config[email], ruleKey: email };
  }

  // Priority 2: domain wildcard match (e.g. "@acme.com")
  if (email.includes('@')) {
    const domain = '@' + email.split('@')[1];
    if (config[domain]) {
      return { label: config[domain], ruleKey: domain };
    }
  }

  // Priority 3: keyword match on subject line
  if (subject && keywordRules) {
    const subjectLower = subject.toLowerCase();
    for (const keyword in keywordRules) {
      if (subjectLower.includes(keyword.toLowerCase())) {
        return { label: keywordRules[keyword], ruleKey: 'keyword:' + keyword };
      }
    }
  }

  return null;
}

/**
 * Checks if an email is excluded based on a pre-loaded exclusion list.
 * Matches both exact email and domain-level patterns.
 *
 * @param {string[]} exclusionList - Array of exclusion patterns.
 * @param {string} email - Email address to check.
 * @returns {boolean} True if excluded.
 */
function isExcludedFromList_(exclusionList, email) {
  if (!exclusionList || exclusionList.length === 0) return false;
  const emailLower = email.toLowerCase();
  const domain = emailLower.includes('@') ? '@' + emailLower.split('@')[1] : '';
  for (let i = 0; i < exclusionList.length; i++) {
    const pattern = exclusionList[i].toLowerCase();
    if (pattern === emailLower || (domain && pattern === domain)) return true;
  }
  return false;
}


// ── Config File I/O ─────────────────────────────────────────────────────────

/**
 * Locates (or creates) the JSON rules file in the same Drive folder as this
 * Apps Script project.
 *
 * @returns {DriveApp.File} The config file handle.
 */
function getConfigFile_() {
  const scriptFile = DriveApp.getFileById(ScriptApp.getScriptId());
  const folder = scriptFile.getParents().next();
  const files = folder.getFilesByName(CONFIG_FILENAME);

  if (files.hasNext()) {
    return files.next();
  }

  Logger.log('Config file not found — creating ' + CONFIG_FILENAME);
  return folder.createFile(CONFIG_FILENAME, '{}', MimeType.PLAIN_TEXT);
}

/**
 * Reads and parses the JSON rules file from Drive.
 *
 * @returns {Object<string, string>} Map of sender/domain → label name.
 */
function loadConfig_() {
  const content = getConfigFile_().getBlob().getDataAsString();

  if (!content) return {};

  try {
    return JSON.parse(content);
  } catch (e) {
    Logger.log('Error parsing config JSON — starting fresh.');
    return {};
  }
}

/**
 * Serializes the rules object back to the JSON file in Drive.
 *
 * @param {Object<string, string>} data - The complete rules map to persist.
 */
function saveConfig_(data) {
  getConfigFile_().setContent(JSON.stringify(data, null, 2));
}


// ── PropertiesService Helpers ───────────────────────────────────────────────
// Exclusions, keyword rules, and filing stats are stored per-user in
// UserProperties to keep the main Drive config file clean.

/** @returns {string[]} Array of excluded sender/domain patterns. */
function loadExclusions_() {
  const raw = PropertiesService.getUserProperties().getProperty('exclusions');
  return raw ? JSON.parse(raw) : [];
}

/** @param {string[]} list - Array of exclusion patterns to persist. */
function saveExclusions_(list) {
  PropertiesService.getUserProperties().setProperty('exclusions', JSON.stringify(list));
}

/**
 * Checks if an email address is in the exclusion list.
 *
 * @param {string} email - Email address to check.
 * @returns {boolean} True if excluded.
 */
function isExcluded_(email) {
  return isExcludedFromList_(loadExclusions_(), email);
}

/** @returns {Object<string,string>} Keyword → label map. */
function loadKeywordRules_() {
  const raw = PropertiesService.getUserProperties().getProperty('keyword_rules');
  return raw ? JSON.parse(raw) : {};
}

/** @param {Object<string,string>} rules - Keyword → label map to persist. */
function saveKeywordRules_(rules) {
  PropertiesService.getUserProperties().setProperty('keyword_rules', JSON.stringify(rules));
}

/** @returns {Object<string,number>} Rule key → lifetime filing count map. */
function loadFilingStats_() {
  const raw = PropertiesService.getUserProperties().getProperty('filing_stats');
  return raw ? JSON.parse(raw) : {};
}

/** @param {Object<string,number>} stats - Stats map to persist. */
function saveFilingStats_(stats) {
  PropertiesService.getUserProperties().setProperty('filing_stats', JSON.stringify(stats));
}

/**
 * Increments the filing counter for a single rule.
 *
 * @param {string} ruleKey - Rule identifier (email, domain, or "keyword:term").
 * @param {number} count - Number of threads filed.
 */
function incrementFilingStat_(ruleKey, count) {
  const stats = loadFilingStats_();
  stats[ruleKey] = (stats[ruleKey] || 0) + count;
  saveFilingStats_(stats);
}

/**
 * Increments filing counters for multiple rules in a single write.
 * Used by runGlobalCleanup to batch stat updates for efficiency.
 *
 * @param {Object<string,number>} updates - Map of ruleKey → count to add.
 */
function batchIncrementFilingStats_(updates) {
  if (!updates || Object.keys(updates).length === 0) return;
  const stats = loadFilingStats_();
  for (const key in updates) {
    stats[key] = (stats[key] || 0) + updates[key];
  }
  saveFilingStats_(stats);
}


// ── Phase 1: Rule Discovery ────────────────────────────────────────────────

/**
 * Scans every user label and records sender → label mappings in the config
 * file. Uses batch message fetching for speed and enforces a time limit to
 * avoid Apps Script execution timeouts.
 */
function discoverRules() {
  const userLabels = GmailApp.getUserLabels();
  const config = loadConfig_();
  const startTime = Date.now();
  let newRulesCount = 0;

  Logger.log('Starting rule discovery across ' + userLabels.length + ' labels…');

  for (let i = 0; i < userLabels.length; i++) {
    // Bail out early if we're approaching the execution time limit
    if (Date.now() - startTime > MAX_DISCOVERY_TIME_MS) {
      Logger.log('Time limit reached after ' + i + ' labels — saving progress.');
      break;
    }

    const label = userLabels[i];
    const labelName = label.getName();

    // Skip labels that shadow system folders
    if (labelName.includes('Trash') || labelName.includes('Spam')) continue;

    const threads = label.getThreads(0, THREAD_LIMIT_PER_LABEL);
    if (threads.length === 0) continue;

    // Batch-fetch the first message of every thread in a single API call
    const messagesRows = GmailApp.getMessagesForThreads(threads);

    for (let j = 0; j < messagesRows.length; j++) {
      const cleanEmail = extractEmail_(messagesRows[j][0].getFrom());

      if (cleanEmail && config[cleanEmail] !== labelName) {
        config[cleanEmail] = labelName;
        newRulesCount++;
      }
    }

    if (i % 10 === 0) {
      Logger.log('Checkpoint: processed ' + (i + 1) + '/' + userLabels.length + ' labels');
    }
  }

  saveConfig_(config);
  Logger.log(
    'Discovery complete. ' + newRulesCount + ' new rules added. '
    + Object.keys(config).length + ' total rules.'
  );
}


// ── Phase 2: Inbox Cleanup ─────────────────────────────────────────────────

/**
 * Sweeps the entire inbox and files every thread whose sender (or sender's
 * domain) matches a rule in the config. Supports dry-run mode for previewing
 * results without moving anything. Respects exclusions and keyword rules.
 * Exact-email rules take priority over domain wildcards, which take priority
 * over keyword rules.
 *
 * @param {boolean} [dryRun=false] - When true, returns results without filing.
 * @returns {Array<Object>|null} Array of match objects in dry-run mode, null otherwise.
 */
function runGlobalCleanup(dryRun) {
  const config = loadConfig_();
  const exclusions = loadExclusions_();
  const keywordRules = loadKeywordRules_();
  const threads = GmailApp.getInboxThreads();
  const results = [];
  const statUpdates = {};

  for (let i = 0; i < threads.length; i++) {
    const firstMsg = threads[i].getMessages()[0];
    const cleanEmail = extractEmail_(firstMsg.getFrom());
    const subject = firstMsg.getSubject() || '';

    const match = resolveRule_(config, exclusions, keywordRules, cleanEmail, subject);

    if (match) {
      if (dryRun) {
        results.push({
          sender: cleanEmail,
          subject: subject,
          label: match.label,
          ruleKey: match.ruleKey
        });
      } else {
        applyLabelToThread_(threads[i], match.label);
        statUpdates[match.ruleKey] = (statUpdates[match.ruleKey] || 0) + 1;
      }
    }
  }

  if (!dryRun) {
    batchIncrementFilingStats_(statUpdates);
  }

  return dryRun ? results : null;
}


// ── Manual Rule Management ─────────────────────────────────────────────────

/**
 * Adds (or updates) a rule in the config and immediately files any matching
 * inbox threads. Supports both full addresses and @domain wildcards.
 * Respects exclusions — if the sender is excluded, the rule is saved but
 * immediate filing is skipped.
 *
 * @param {string} email - Sender email or "@domain.com" pattern.
 * @param {string} labelName - Destination label name.
 */
function addNewRule(email, labelName) {
  const config = loadConfig_();
  config[email] = labelName;
  saveConfig_(config);
  Logger.log('Rule saved: ' + email + ' → ' + labelName);

  // Check exclusion list before filing
  if (isExcluded_(email)) {
    Logger.log('Note: ' + email + ' is in exclusion list — skipping immediate filing.');
    return;
  }

  // Immediately clean matching inbox threads
  const threads = GmailApp.search('from:' + email + ' is:inbox');
  if (threads.length > 0) {
    applyLabelToThreads_(threads, labelName, email);
    Logger.log('Cleaned ' + threads.length + ' existing threads.');
  }

  // Auto-consolidate if enabled
  if (AUTO_CONSOLIDATE) {
    const opportunities = findConsolidationOpportunities_();
    if (opportunities.length > 0) {
      consolidateRules(opportunities);
    }
  }
}

/**
 * Deletes a rule from the config file.
 *
 * @param {string} email - The sender email or domain pattern to remove.
 * @returns {boolean} True if the rule existed and was deleted, false otherwise.
 */
function deleteRule(email) {
  const config = loadConfig_();

  if (config.hasOwnProperty(email)) {
    delete config[email];
    saveConfig_(config);
    Logger.log('Rule deleted: ' + email);
    return true;
  }

  Logger.log('Rule not found: ' + email);
  return false;
}

/**
 * Retrieves all rules from the config, optionally filtered by search term.
 *
 * @param {string} [searchTerm] - Optional filter for sender email or label (case-insensitive).
 * @returns {Array<{sender: string, label: string}>} Array of rule objects.
 */
function getRules(searchTerm) {
  const config = loadConfig_();
  const rules = [];

  for (const email in config) {
    if (searchTerm) {
      const term = searchTerm.toLowerCase();
      if (email.toLowerCase().includes(term) || config[email].toLowerCase().includes(term)) {
        rules.push({ sender: email, label: config[email] });
      }
    } else {
      rules.push({ sender: email, label: config[email] });
    }
  }

  // Sort by sender email for consistent display
  rules.sort((a, b) => a.sender.localeCompare(b.sender));
  return rules;
}

/**
 * Gets basic statistics about the current rule set including exclusion
 * and keyword rule counts.
 *
 * @returns {Object} Stats object with totalRules, domainRules, exactRules,
 *   exclusionCount, and keywordCount.
 */
function getRuleStats() {
  const config = loadConfig_();
  const exclusions = loadExclusions_();
  const keywordRules = loadKeywordRules_();
  const stats = {
    totalRules: 0,
    domainRules: 0,
    exactRules: 0,
    exclusionCount: exclusions.length,
    keywordCount: Object.keys(keywordRules).length
  };

  for (const email in config) {
    stats.totalRules++;
    if (email.startsWith('@')) {
      stats.domainRules++;
    } else {
      stats.exactRules++;
    }
  }

  return stats;
}

/**
 * Analyzes rules and finds consolidation opportunities where multiple rules
 * could be merged into a single broader rule.
 *
 * @returns {Array<Object>} Array of consolidation opportunities, each containing:
 *   - type: 'domain-merge' or 'redundant'
 *   - rulesToRemove: Array of sender patterns to delete
 *   - ruleToAdd: New sender pattern (null for redundant rules)
 *   - label: Target label name
 *   - description: Human-readable explanation
 */
function findConsolidationOpportunities_() {
  const config = loadConfig_();
  const opportunities = [];

  // Group rules by target label
  const labelGroups = {};
  for (const email in config) {
    const label = config[email];
    if (!labelGroups[label]) {
      labelGroups[label] = [];
    }
    labelGroups[label].push(email);
  }

  // Analyze each label group for consolidation opportunities
  for (const label in labelGroups) {
    const senders = labelGroups[label];

    // Skip if only one rule for this label
    if (senders.length < 2) continue;

    const domainRules = senders.filter(s => s.startsWith('@'));
    const exactRules = senders.filter(s => !s.startsWith('@'));

    // Type 1: Find redundant exact rules covered by domain rules
    for (const exact of exactRules) {
      const domain = '@' + exact.split('@')[1];
      if (domainRules.includes(domain)) {
        opportunities.push({
          type: 'redundant',
          rulesToRemove: [exact],
          ruleToAdd: null,
          label: label,
          description: exact + ' is redundant (covered by ' + domain + ')'
        });
      }
    }

    // Type 2: Find multiple subdomain rules that could merge to parent
    const domainHierarchy = {};
    for (const domain of domainRules) {
      const parts = domain.substring(1).split('.');
      if (parts.length >= 2) {
        const parent = '@' + parts.slice(-2).join('.');
        if (!domainHierarchy[parent]) {
          domainHierarchy[parent] = [];
        }
        domainHierarchy[parent].push(domain);
      }
    }

    for (const parent in domainHierarchy) {
      const subdomains = domainHierarchy[parent];
      if (subdomains.length >= 2 && !domainRules.includes(parent)) {
        opportunities.push({
          type: 'domain-merge',
          rulesToRemove: subdomains,
          ruleToAdd: parent,
          label: label,
          description: 'Merge ' + subdomains.join(', ') + ' → ' + parent
        });
      }
    }

    // Type 3: Find multiple exact emails from same domain (3+ emails)
    const emailsByDomain = {};
    for (const exact of exactRules) {
      if (exact.includes('@')) {
        const domain = '@' + exact.split('@')[1];
        if (!emailsByDomain[domain]) {
          emailsByDomain[domain] = [];
        }
        emailsByDomain[domain].push(exact);
      }
    }

    for (const domain in emailsByDomain) {
      const emails = emailsByDomain[domain];
      if (emails.length >= 3 && !domainRules.includes(domain)) {
        opportunities.push({
          type: 'domain-merge',
          rulesToRemove: emails,
          ruleToAdd: domain,
          label: label,
          description: 'Merge ' + emails.length + ' emails → ' + domain
        });
      }
    }
  }

  return opportunities;
}

/**
 * Applies a set of consolidation opportunities to the config.
 *
 * @param {Array<Object>} opportunities - Array of opportunities from findConsolidationOpportunities_()
 * @returns {number} Number of rules removed/simplified
 */
function consolidateRules(opportunities) {
  if (!opportunities || opportunities.length === 0) {
    Logger.log('No consolidation opportunities found.');
    return 0;
  }

  const config = loadConfig_();
  let changesCount = 0;

  for (const opp of opportunities) {
    // Remove the old rules
    for (const sender of opp.rulesToRemove) {
      if (config[sender]) {
        delete config[sender];
        changesCount++;
        Logger.log('Removed: ' + sender);
      }
    }

    // Add the new consolidated rule (if applicable)
    if (opp.ruleToAdd) {
      config[opp.ruleToAdd] = opp.label;
      Logger.log('Added: ' + opp.ruleToAdd + ' → ' + opp.label);
    }
  }

  saveConfig_(config);
  Logger.log('Consolidation complete. ' + changesCount + ' rules simplified.');
  return changesCount;
}

/**
 * Quick-run wrapper for adding a rule from the script editor.
 * Edit the variables below, then execute this function manually.
 */
function manualInput() {
  const emailToAdd = 'noreply@example.com';
  const labelTarget = 'Example';
  addNewRule(emailToAdd, labelTarget);
}


// ── Exclusion Management ────────────────────────────────────────────────────

/**
 * Adds a sender or domain to the exclusion list. Excluded senders are
 * never auto-filed during cleanup or rule-based filing, even if a
 * matching rule exists.
 *
 * @param {string} pattern - Email address or "@domain.com" wildcard.
 */
function addExclusion(pattern) {
  const list = loadExclusions_();
  const normalized = pattern.toLowerCase().trim();
  if (list.indexOf(normalized) === -1) {
    list.push(normalized);
    saveExclusions_(list);
    Logger.log('Exclusion added: ' + normalized);
  } else {
    Logger.log('Already excluded: ' + normalized);
  }
}

/**
 * Removes a sender or domain from the exclusion list.
 *
 * @param {string} pattern - The exclusion pattern to remove.
 * @returns {boolean} True if removed, false if not found.
 */
function removeExclusion(pattern) {
  const list = loadExclusions_();
  const normalized = pattern.toLowerCase().trim();
  const idx = list.indexOf(normalized);
  if (idx !== -1) {
    list.splice(idx, 1);
    saveExclusions_(list);
    Logger.log('Exclusion removed: ' + normalized);
    return true;
  }
  Logger.log('Exclusion not found: ' + normalized);
  return false;
}

/**
 * Returns all current exclusion patterns, optionally filtered.
 *
 * @param {string} [searchTerm] - Optional case-insensitive filter.
 * @returns {string[]} Array of exclusion patterns.
 */
function getExclusions(searchTerm) {
  const list = loadExclusions_();
  if (!searchTerm) return list.sort();
  const term = searchTerm.toLowerCase();
  return list.filter(p => p.includes(term)).sort();
}


// ── Keyword Rule Management ────────────────────────────────────────────────

/**
 * Adds a keyword rule that matches on message subject lines. When a
 * subject contains the keyword (case-insensitive), the thread is filed
 * to the specified label. Keyword rules have the lowest priority (after
 * exact sender and domain wildcard rules).
 *
 * @param {string} keyword - Substring to match in subject lines.
 * @param {string} labelName - Destination label name.
 */
function addKeywordRule(keyword, labelName) {
  const rules = loadKeywordRules_();
  rules[keyword.trim()] = labelName;
  saveKeywordRules_(rules);
  Logger.log('Keyword rule saved: "' + keyword + '" → ' + labelName);
}

/**
 * Deletes a keyword rule.
 *
 * @param {string} keyword - The keyword to remove.
 * @returns {boolean} True if removed, false if not found.
 */
function deleteKeywordRule(keyword) {
  const rules = loadKeywordRules_();
  if (rules.hasOwnProperty(keyword)) {
    delete rules[keyword];
    saveKeywordRules_(rules);
    Logger.log('Keyword rule deleted: "' + keyword + '"');
    return true;
  }
  return false;
}

/**
 * Returns all keyword rules, optionally filtered.
 *
 * @param {string} [searchTerm] - Optional case-insensitive filter.
 * @returns {Array<{keyword: string, label: string}>} Array of keyword rule objects.
 */
function getKeywordRules(searchTerm) {
  const rules = loadKeywordRules_();
  const result = [];
  for (const kw in rules) {
    if (searchTerm) {
      const term = searchTerm.toLowerCase();
      if (kw.toLowerCase().includes(term) || rules[kw].toLowerCase().includes(term)) {
        result.push({ keyword: kw, label: rules[kw] });
      }
    } else {
      result.push({ keyword: kw, label: rules[kw] });
    }
  }
  result.sort((a, b) => a.keyword.localeCompare(b.keyword));
  return result;
}


// ── Filing Stats ────────────────────────────────────────────────────────────

/**
 * Returns filing statistics for all rules. Each entry records how many
 * threads a given rule has filed over its lifetime.
 *
 * @returns {Object<string, number>} Map of rule key → filing count.
 */
function getFilingStats() {
  return loadFilingStats_();
}

/**
 * Resets all filing statistics to zero.
 */
function resetFilingStats() {
  saveFilingStats_({});
  Logger.log('Filing stats reset.');
}


// ── Rule Import / Export ────────────────────────────────────────────────────

/**
 * Exports the current rules to a timestamped JSON backup file in the same
 * Drive folder as the script. Returns the filename for confirmation.
 *
 * @returns {string} The name of the created backup file.
 */
function exportRulesToDrive() {
  const config = loadConfig_();
  const json = JSON.stringify(config, null, 2);
  const timestamp = Utilities.formatDate(
    new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HHmm'
  );
  const filename = 'mail-rules-backup-' + timestamp + '.json';

  const scriptFile = DriveApp.getFileById(ScriptApp.getScriptId());
  const folder = scriptFile.getParents().next();
  folder.createFile(filename, json, MimeType.PLAIN_TEXT);

  Logger.log('Rules exported to: ' + filename);
  return filename;
}

/**
 * Imports rules from a named JSON file in the script's Drive folder.
 * Supports merge (add new, update existing) or full replacement.
 *
 * @param {string} filename - Name of the JSON file to import from.
 * @param {boolean} [replaceAll=false] - When true, replaces all rules.
 *   When false, merges imported rules with existing ones.
 * @returns {Object} Result with success boolean, message, and count.
 */
function importRulesFromFile(filename, replaceAll) {
  const scriptFile = DriveApp.getFileById(ScriptApp.getScriptId());
  const folder = scriptFile.getParents().next();
  const files = folder.getFilesByName(filename);

  if (!files.hasNext()) {
    return { success: false, message: 'File not found: ' + filename, count: 0 };
  }

  const content = files.next().getBlob().getDataAsString();
  let imported;
  try {
    imported = JSON.parse(content);
  } catch (e) {
    return { success: false, message: 'Invalid JSON in file.', count: 0 };
  }

  if (typeof imported !== 'object' || Array.isArray(imported)) {
    return { success: false, message: 'Expected a JSON object of rules.', count: 0 };
  }

  const importedCount = Object.keys(imported).length;

  if (replaceAll) {
    saveConfig_(imported);
    Logger.log('Replaced all rules with ' + importedCount + ' imported rules.');
  } else {
    const config = loadConfig_();
    let newCount = 0;
    for (const key in imported) {
      if (!config.hasOwnProperty(key)) newCount++;
      config[key] = imported[key];
    }
    saveConfig_(config);
    Logger.log('Merged ' + importedCount + ' rules (' + newCount + ' new).');
  }

  return { success: true, message: 'Imported ' + importedCount + ' rules.', count: importedCount };
}


// ── Trigger Management ──────────────────────────────────────────────────────

/**
 * Returns the current status of automated triggers for cleanup and
 * discovery functions.
 *
 * @returns {Object} Status object with cleanup and discovery booleans.
 */
function getTriggerStatus_() {
  const triggers = ScriptApp.getProjectTriggers();
  const status = { cleanup: false, discovery: false };

  for (const trigger of triggers) {
    const fn = trigger.getHandlerFunction();
    if (fn === 'runGlobalCleanup') status.cleanup = true;
    if (fn === 'discoverRules') status.discovery = true;
  }

  return status;
}

/**
 * Creates an hourly time-based trigger for `runGlobalCleanup`.
 * Removes any existing cleanup trigger first to avoid duplicates.
 */
function enableCleanupTrigger() {
  disableCleanupTrigger();
  ScriptApp.newTrigger('runGlobalCleanup')
    .timeBased()
    .everyHours(1)
    .create();
  Logger.log('Cleanup trigger enabled: every 1 hour.');
}

/**
 * Removes all time-based triggers for `runGlobalCleanup`.
 */
function disableCleanupTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === 'runGlobalCleanup') {
      ScriptApp.deleteTrigger(trigger);
    }
  }
  Logger.log('Cleanup trigger disabled.');
}

/**
 * Creates a daily time-based trigger for `discoverRules`.
 * Removes any existing discovery trigger first to avoid duplicates.
 */
function enableDiscoveryTrigger() {
  disableDiscoveryTrigger();
  ScriptApp.newTrigger('discoverRules')
    .timeBased()
    .everyDays(1)
    .create();
  Logger.log('Discovery trigger enabled: every 24 hours.');
}

/**
 * Removes all time-based triggers for `discoverRules`.
 */
function disableDiscoveryTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === 'discoverRules') {
      ScriptApp.deleteTrigger(trigger);
    }
  }
  Logger.log('Discovery trigger disabled.');
}


// ── Label Reorganization Utilities ──────────────────────────────────────────

/**
 * Updates every config rule that points to `oldName` so it points to
 * `newName` instead. Use this **after** manually renaming a large label
 * in the Gmail sidebar (avoids moving hundreds of threads via API).
 *
 * @param {string} oldName - The previous label name.
 * @param {string} newName - The new label name already applied in Gmail.
 */
function syncJsonAfterManualRename(oldName, newName) {
  const config = loadConfig_();
  let updatedCount = 0;

  for (const email in config) {
    if (config[email] === oldName) {
      config[email] = newName;
      updatedCount++;
    }
  }

  saveConfig_(config);
  Logger.log(updatedCount + ' rules now point to "' + newName + '".');
}

/**
 * Fully migrates a label: moves all threads from `oldName` to `newName`,
 * deletes the old label, and updates every matching config rule. Best
 * suited for small/medium labels (< 200 threads). For large labels, use
 * `syncJsonAfterManualRename` instead.
 *
 * @param {string} oldName - Current label name.
 * @param {string} newName - Desired new label name / path.
 */
function migrateLabel(oldName, newName) {
  const config = loadConfig_();
  const oldLabel = GmailApp.getUserLabelByName(oldName);
  const newLabel = getOrCreateLabel_(newName);

  if (oldLabel) {
    const threads = oldLabel.getThreads();
    if (threads.length > 0) {
      newLabel.addToThreads(threads);
      oldLabel.removeFromThreads(threads);
    }
    oldLabel.deleteLabel();
    Logger.log('Migrated Gmail label: ' + oldName + ' → ' + newName);
  }

  // Update config rules
  let updatedCount = 0;
  for (const email in config) {
    if (config[email] === oldName) {
      config[email] = newName;
      updatedCount++;
    }
  }

  saveConfig_(config);
  Logger.log('Updated ' + updatedCount + ' rules in config.');
}


// ── Label Color Theming ────────────────────────────────────────────────────

/**
 * Applies color themes to top-level label groups using the Gmail Advanced
 * Service (REST API). Colors are drawn from the official Gmail palette to
 * avoid "Invalid Color" errors. Sub-labels inherit their parent's color.
 */
function autoColorLabels() {
  const labels = Gmail.Users.Labels.list('me').labels;

  // Official Gmail API hex codes — using values outside this palette will fail.
  const THEMES = {
    '@Shopping':      { bg: '#fb4c2f', txt: '#ffffff' }, // Red
    '@Work':          { bg: '#4a86e8', txt: '#ffffff' }, // Blue
    '@Finance':       { bg: '#16a766', txt: '#ffffff' }, // Green
    '@Notifications': { bg: '#ffad47', txt: '#ffffff' }, // Orange
    '@Social':        { bg: '#999999', txt: '#ffffff' }, // Grey
    '@Projects':      { bg: '#a479e2', txt: '#ffffff' }, // Purple
    '@Updates':       { bg: '#fad165', txt: '#ffffff' }, // Yellow
    '@Travel':        { bg: '#a4c2f4', txt: '#ffffff' }, // Sky Blue
  };

  Logger.log('Applying label colors via Gmail API…');

  labels.forEach(function (label) {
    if (label.type !== 'user') return; // skip system labels

    for (const parent in THEMES) {
      if (label.name === parent || label.name.startsWith(parent + '/')) {
        try {
          Gmail.Users.Labels.patch(
            { color: { backgroundColor: THEMES[parent].bg, textColor: THEMES[parent].txt } },
            'me',
            label.id
          );
          Logger.log('Colored: ' + label.name);
        } catch (err) {
          Logger.log('Error coloring ' + label.name + ': ' + err.message);
        }
        break; // matched this label — no need to check remaining parents
      }
    }
  });
}


// ── Add-on UI (Gmail Side Panel) ────────────────────────────────────────────

/**
 * Builds the Card-based UI shown in the Gmail side panel. Sections:
 *   1. Quick File – pre-populated sender from current message (contextual).
 *   2. Add New Rule – map a sender/domain to a label.
 *   3. Manage Rules – view, edit, delete rules, consolidate.
 *   4. Reorganize Labels – rename or migrate labels.
 *   5. Settings – sync rules, run sweeps, refresh colors.
 *   6. Advanced – triggers, keywords, exclusions, import/export, dry run.
 *
 * @param {Object} [e] - Gmail add-on event (contains messageId in contextual mode).
 * @returns {CardService.Card}
 */
function buildAddOn(e) {
  const card = CardService.newCardBuilder();
  card.setHeader(CardService.newCardHeader().setTitle('Inbox Cleaner'));

  // ── Section 0 — Quick File (contextual: only when viewing a message) ──
  if (e && e.gmail && e.gmail.messageId) {
    try {
      const msg = GmailApp.getMessageById(e.gmail.messageId);
      const senderEmail = extractEmail_(msg.getFrom());
      const domain = '@' + senderEmail.split('@')[1];
      const config = loadConfig_();
      const existingRule = config[senderEmail] || config[domain];

      const quickSection = CardService.newCardSection().setHeader('⚡ Quick File');
      quickSection.addWidget(CardService.newTextParagraph()
        .setText('Sender: <b>' + senderEmail + '</b>'
          + (existingRule
            ? '<br>Current rule: → ' + existingRule
            : '<br>No rule for this sender')));
      quickSection.addWidget(CardService.newTextInput()
        .setFieldName('quick_label')
        .setTitle('Target Label')
        .setValue(existingRule || ''));
      quickSection.addWidget(CardService.newButtonSet()
        .addButton(CardService.newTextButton()
          .setText('File Sender')
          .setOnClickAction(CardService.newAction()
            .setFunctionName('handleQuickFile')
            .setParameters({ sender: senderEmail })))
        .addButton(CardService.newTextButton()
          .setText('File ' + domain)
          .setOnClickAction(CardService.newAction()
            .setFunctionName('handleQuickFile')
            .setParameters({ sender: domain }))));
      card.addSection(quickSection);
    } catch (err) {
      Logger.log('Quick File error: ' + err.message);
    }
  }

  // ── Section 1 — Add New Rule ──
  const addSection = CardService.newCardSection().setHeader('Add New Rule');
  addSection.addWidget(CardService.newTextInput()
    .setFieldName('sender_email')
    .setTitle('Sender Email or @domain.com'));
  addSection.addWidget(CardService.newTextInput()
    .setFieldName('label_name')
    .setTitle('Target Label (use / for sublabels)'));
  addSection.addWidget(CardService.newTextButton()
    .setText('Add & Clean Now')
    .setOnClickAction(CardService.newAction().setFunctionName('handleUiSubmit')));

  // ── Section 2 — Manage Existing Rules ──
  const manageSection = CardService.newCardSection().setHeader('Manage Rules');
  const stats = getRuleStats();
  manageSection.addWidget(CardService.newTextParagraph()
    .setText('<b>' + stats.totalRules + '</b> sender rules (' + stats.exactRules
      + ' exact, ' + stats.domainRules + ' domains)'
      + '<br>' + stats.keywordCount + ' keyword rules · '
      + stats.exclusionCount + ' exclusions'));
  manageSection.addWidget(CardService.newTextButton()
    .setText('View/Edit/Delete Rules')
    .setOnClickAction(CardService.newAction().setFunctionName('showRuleEditor')));
  manageSection.addWidget(CardService.newTextButton()
    .setText('🔄 Consolidate Similar Rules')
    .setOnClickAction(CardService.newAction().setFunctionName('showConsolidationPreview')));

  // ── Section 3 — Reorganize Labels ──
  const migrateSection = CardService.newCardSection().setHeader('Reorganize Labels');
  migrateSection.addWidget(CardService.newTextInput()
    .setFieldName('old_label')
    .setTitle('Old Label Name'));
  migrateSection.addWidget(CardService.newTextInput()
    .setFieldName('new_label')
    .setTitle('New Path (e.g. Work/Invoices)'));
  migrateSection.addWidget(CardService.newTextButton()
    .setText('Auto Move (Small Labels)')
    .setOnClickAction(CardService.newAction().setFunctionName('handleUiMigrate')));
  migrateSection.addWidget(CardService.newTextParagraph()
    .setText('<b>For large labels:</b> Rename in Gmail manually first, then click below.'));
  migrateSection.addWidget(CardService.newTextButton()
    .setText('Sync Rules Only (Big Labels)')
    .setOnClickAction(CardService.newAction().setFunctionName('handleUiSyncOnly')));

  // ── Section 4 — Settings / Maintenance ──
  const syncSection = CardService.newCardSection().setHeader('Settings');
  syncSection.addWidget(CardService.newTextButton()
    .setText('Sync Rules from Labels')
    .setOnClickAction(CardService.newAction().setFunctionName('handleUiSync')));
  syncSection.addWidget(CardService.newTextButton()
    .setText('Run Full Inbox Sweep')
    .setOnClickAction(CardService.newAction().setFunctionName('handleUiCleanup')));
  syncSection.addWidget(CardService.newTextButton()
    .setText('Refresh Label Colors')
    .setOnClickAction(CardService.newAction().setFunctionName('handleUiColor')));

  // ── Section 5 — Advanced Tools ──
  const advancedSection = CardService.newCardSection().setHeader('Advanced');
  advancedSection.addWidget(CardService.newTextButton()
    .setText('🔍 Dry Run Preview')
    .setOnClickAction(CardService.newAction().setFunctionName('handleUiDryRun')));
  advancedSection.addWidget(CardService.newTextButton()
    .setText('🚫 Manage Exclusions')
    .setOnClickAction(CardService.newAction().setFunctionName('showExclusionEditor')));
  advancedSection.addWidget(CardService.newTextButton()
    .setText('🔑 Keyword Rules')
    .setOnClickAction(CardService.newAction().setFunctionName('showKeywordEditor')));
  advancedSection.addWidget(CardService.newTextButton()
    .setText('⏰ Trigger Settings')
    .setOnClickAction(CardService.newAction().setFunctionName('showTriggerSettings')));
  advancedSection.addWidget(CardService.newTextButton()
    .setText('📦 Import / Export')
    .setOnClickAction(CardService.newAction().setFunctionName('showImportExport')));

  card.addSection(addSection);
  card.addSection(manageSection);
  card.addSection(migrateSection);
  card.addSection(syncSection);
  card.addSection(advancedSection);
  return card.build();
}


// ── UI View Cards ───────────────────────────────────────────────────────────

/**
 * Builds a card displaying existing rules with search/filter, edit/delete
 * actions, and filing stats. Uses ButtonSet for Edit/Delete to avoid the
 * single-button limitation of DecoratedText.
 *
 * @param {Object} [e] - Card action event with optional formInput and parameters.
 * @returns {CardService.Card}
 */
function showRuleEditor(e) {
  const searchTerm = e && e.formInput && e.formInput.search_rules ? e.formInput.search_rules : '';
  const rules = getRules(searchTerm);
  const filingStats = loadFilingStats_();
  const RULES_PER_PAGE = 10;
  const page = e && e.parameters && e.parameters.page ? parseInt(e.parameters.page) : 0;
  const startIdx = page * RULES_PER_PAGE;
  const endIdx = Math.min(startIdx + RULES_PER_PAGE, rules.length);

  const card = CardService.newCardBuilder();
  card.setHeader(CardService.newCardHeader()
    .setTitle('Rule Editor')
    .setSubtitle(rules.length + ' rules found'));

  // Search section
  const searchSection = CardService.newCardSection();
  searchSection.addWidget(CardService.newTextInput()
    .setFieldName('search_rules')
    .setTitle('Search by sender or label')
    .setValue(searchTerm || ''));
  searchSection.addWidget(CardService.newTextButton()
    .setText('Search')
    .setOnClickAction(CardService.newAction().setFunctionName('showRuleEditor')));
  searchSection.addWidget(CardService.newTextButton()
    .setText('← Back to Main')
    .setOnClickAction(CardService.newAction().setFunctionName('buildAddOn')));

  // Rules list section
  const rulesSection = CardService.newCardSection()
    .setHeader('Rules ' + (startIdx + 1) + '-' + endIdx + ' of ' + rules.length);

  if (rules.length === 0) {
    rulesSection.addWidget(CardService.newTextParagraph()
      .setText('No rules found. '
        + (searchTerm ? 'Try a different search term.' : 'Add some rules to get started!')));
  } else {
    for (let i = startIdx; i < endIdx; i++) {
      const rule = rules[i];
      const isDomain = rule.sender.startsWith('@');
      const icon = isDomain ? '🌐' : '📧';
      const hitCount = filingStats[rule.sender] || 0;
      const statsText = hitCount > 0 ? ' · ' + hitCount + ' filed' : '';

      // Display rule info
      rulesSection.addWidget(CardService.newDecoratedText()
        .setText(icon + ' <b>' + rule.sender + '</b>')
        .setBottomLabel('→ ' + rule.label + statsText));

      // Edit + Delete buttons (ButtonSet avoids DecoratedText single-button limit)
      rulesSection.addWidget(CardService.newButtonSet()
        .addButton(CardService.newTextButton()
          .setText('Edit')
          .setOnClickAction(CardService.newAction()
            .setFunctionName('showEditRule')
            .setParameters({ sender: rule.sender, label: rule.label })))
        .addButton(CardService.newTextButton()
          .setText('Delete')
          .setOnClickAction(CardService.newAction()
            .setFunctionName('handleDeleteRule')
            .setParameters({ sender: rule.sender }))));
    }

    // Pagination buttons
    if (rules.length > RULES_PER_PAGE) {
      const navSection = CardService.newCardSection();
      const buttonSet = CardService.newButtonSet();

      if (page > 0) {
        buttonSet.addButton(CardService.newTextButton()
          .setText('◀ Previous')
          .setOnClickAction(CardService.newAction()
            .setFunctionName('showRuleEditor')
            .setParameters({ page: (page - 1).toString(), search: searchTerm })));
      }

      if (endIdx < rules.length) {
        buttonSet.addButton(CardService.newTextButton()
          .setText('Next ▶')
          .setOnClickAction(CardService.newAction()
            .setFunctionName('showRuleEditor')
            .setParameters({ page: (page + 1).toString(), search: searchTerm })));
      }

      navSection.addWidget(buttonSet);
      card.addSection(navSection);
    }
  }

  card.addSection(searchSection);
  card.addSection(rulesSection);

  return card.build();
}

/**
 * Shows the edit form for a specific rule.
 *
 * @param {Object} e - Event containing parameters.sender and parameters.label.
 * @returns {CardService.Card}
 */
function showEditRule(e) {
  const sender = e.parameters.sender;
  const currentLabel = e.parameters.label;

  const card = CardService.newCardBuilder();
  card.setHeader(CardService.newCardHeader().setTitle('Edit Rule'));

  const editSection = CardService.newCardSection();
  editSection.addWidget(CardService.newTextParagraph()
    .setText('<b>Sender:</b> ' + sender));
  editSection.addWidget(CardService.newTextInput()
    .setFieldName('new_label')
    .setTitle('New Label')
    .setValue(currentLabel));
  editSection.addWidget(CardService.newButtonSet()
    .addButton(CardService.newTextButton()
      .setText('Save Changes')
      .setOnClickAction(CardService.newAction()
        .setFunctionName('handleEditRule')
        .setParameters({ sender: sender })))
    .addButton(CardService.newTextButton()
      .setText('Cancel')
      .setOnClickAction(CardService.newAction().setFunctionName('showRuleEditor'))));

  card.addSection(editSection);
  return card.build();
}

/**
 * Shows consolidation opportunities for user review and approval.
 *
 * @returns {CardService.Card}
 */
function showConsolidationPreview(e) {
  const opportunities = findConsolidationOpportunities_();

  const card = CardService.newCardBuilder();
  card.setHeader(CardService.newCardHeader()
    .setTitle('Rule Consolidation')
    .setSubtitle(opportunities.length + ' opportunities found'));

  const infoSection = CardService.newCardSection();
  infoSection.addWidget(CardService.newTextButton()
    .setText('← Back to Main')
    .setOnClickAction(CardService.newAction().setFunctionName('buildAddOn')));

  if (opportunities.length === 0) {
    infoSection.addWidget(CardService.newTextParagraph()
      .setText('✓ Your rules are already optimized! No consolidations needed.'));
  } else {
    infoSection.addWidget(CardService.newTextParagraph()
      .setText('The following rules can be simplified. Review and apply to optimize your config.'));

    // Show each opportunity
    const oppSection = CardService.newCardSection().setHeader('Opportunities');
    for (let i = 0; i < Math.min(opportunities.length, 10); i++) {
      const opp = opportunities[i];
      const icon = opp.type === 'redundant' ? '🗑️' : '🔀';
      oppSection.addWidget(CardService.newTextParagraph()
        .setText(icon + ' ' + opp.description));
    }

    if (opportunities.length > 10) {
      oppSection.addWidget(CardService.newTextParagraph()
        .setText('...and ' + (opportunities.length - 10) + ' more'));
    }

    // Apply button
    const actionSection = CardService.newCardSection();
    actionSection.addWidget(CardService.newTextButton()
      .setText('✓ Apply All Consolidations')
      .setOnClickAction(CardService.newAction().setFunctionName('handleConsolidate')));

    card.addSection(oppSection);
    card.addSection(actionSection);
  }

  card.addSection(infoSection);
  return card.build();
}

/**
 * Shows the dry-run results card with a summary of what cleanup would file,
 * grouped by target label.
 *
 * @param {Array<Object>} results - Array of match objects from runGlobalCleanup(true).
 * @returns {CardService.Card}
 */
function showDryRunResults(results) {
  const card = CardService.newCardBuilder();
  card.setHeader(CardService.newCardHeader()
    .setTitle('Dry Run Preview')
    .setSubtitle(results.length + ' threads would be filed'));

  const navSection = CardService.newCardSection();
  navSection.addWidget(CardService.newTextButton()
    .setText('← Back to Main')
    .setOnClickAction(CardService.newAction().setFunctionName('buildAddOn')));

  if (results.length === 0) {
    navSection.addWidget(CardService.newTextParagraph()
      .setText('✓ No inbox threads match any rules. Nothing to file.'));
  } else {
    // Group results by label
    const byLabel = {};
    for (const r of results) {
      if (!byLabel[r.label]) byLabel[r.label] = [];
      byLabel[r.label].push(r);
    }

    const detailSection = CardService.newCardSection().setHeader('Results by Label');
    const sortedLabels = Object.keys(byLabel).sort();
    for (const label of sortedLabels) {
      const items = byLabel[label];
      const senders = items.map(i => i.sender).filter((v, idx, a) => a.indexOf(v) === idx);
      detailSection.addWidget(CardService.newDecoratedText()
        .setText('<b>' + label + '</b>')
        .setBottomLabel(items.length + ' threads from ' + senders.length + ' senders'));
    }

    // Proceed button
    detailSection.addWidget(CardService.newTextButton()
      .setText('▶ Run Full Sweep Now')
      .setOnClickAction(CardService.newAction().setFunctionName('handleUiCleanup')));

    card.addSection(detailSection);
  }

  card.addSection(navSection);
  return card.build();
}

/**
 * Shows the exclusion list editor with add/remove functionality.
 *
 * @param {Object} [e] - Card action event.
 * @returns {CardService.Card}
 */
function showExclusionEditor(e) {
  const exclusions = getExclusions();

  const card = CardService.newCardBuilder();
  card.setHeader(CardService.newCardHeader()
    .setTitle('Exclusion List')
    .setSubtitle(exclusions.length + ' exclusions'));

  // Add new exclusion
  const addSection = CardService.newCardSection().setHeader('Add Exclusion');
  addSection.addWidget(CardService.newTextInput()
    .setFieldName('exclusion_pattern')
    .setTitle('Email or @domain.com'));
  addSection.addWidget(CardService.newButtonSet()
    .addButton(CardService.newTextButton()
      .setText('Add Exclusion')
      .setOnClickAction(CardService.newAction().setFunctionName('handleAddExclusion')))
    .addButton(CardService.newTextButton()
      .setText('← Back')
      .setOnClickAction(CardService.newAction().setFunctionName('buildAddOn'))));

  // List current exclusions
  const listSection = CardService.newCardSection().setHeader('Current Exclusions');
  if (exclusions.length === 0) {
    listSection.addWidget(CardService.newTextParagraph()
      .setText('No exclusions set. Excluded senders are never auto-filed.'));
  } else {
    for (const pattern of exclusions) {
      const icon = pattern.startsWith('@') ? '🌐' : '📧';
      listSection.addWidget(CardService.newDecoratedText()
        .setText(icon + ' ' + pattern)
        .setButton(CardService.newTextButton()
          .setText('Remove')
          .setOnClickAction(CardService.newAction()
            .setFunctionName('handleDeleteExclusion')
            .setParameters({ pattern: pattern }))));
    }
  }

  card.addSection(addSection);
  card.addSection(listSection);
  return card.build();
}

/**
 * Shows the keyword rules editor with add/remove functionality.
 *
 * @param {Object} [e] - Card action event.
 * @returns {CardService.Card}
 */
function showKeywordEditor(e) {
  const keywords = getKeywordRules();
  const filingStats = loadFilingStats_();

  const card = CardService.newCardBuilder();
  card.setHeader(CardService.newCardHeader()
    .setTitle('Keyword Rules')
    .setSubtitle(keywords.length + ' keyword rules'));

  // Add new keyword rule
  const addSection = CardService.newCardSection().setHeader('Add Keyword Rule');
  addSection.addWidget(CardService.newTextParagraph()
    .setText('Keywords match against message subject lines (case-insensitive). '
      + 'They have lowest priority: sender and domain rules are checked first.'));
  addSection.addWidget(CardService.newTextInput()
    .setFieldName('keyword_text')
    .setTitle('Keyword (substring match)'));
  addSection.addWidget(CardService.newTextInput()
    .setFieldName('keyword_label')
    .setTitle('Target Label'));
  addSection.addWidget(CardService.newButtonSet()
    .addButton(CardService.newTextButton()
      .setText('Add Keyword Rule')
      .setOnClickAction(CardService.newAction().setFunctionName('handleAddKeyword')))
    .addButton(CardService.newTextButton()
      .setText('← Back')
      .setOnClickAction(CardService.newAction().setFunctionName('buildAddOn'))));

  // List current keyword rules
  const listSection = CardService.newCardSection().setHeader('Current Keywords');
  if (keywords.length === 0) {
    listSection.addWidget(CardService.newTextParagraph()
      .setText('No keyword rules set. Add one above to match by subject line.'));
  } else {
    for (const kw of keywords) {
      const statKey = 'keyword:' + kw.keyword;
      const hitCount = filingStats[statKey] || 0;
      const statsText = hitCount > 0 ? ' · ' + hitCount + ' filed' : '';
      listSection.addWidget(CardService.newDecoratedText()
        .setText('🔑 <b>"' + kw.keyword + '"</b>')
        .setBottomLabel('→ ' + kw.label + statsText)
        .setButton(CardService.newTextButton()
          .setText('Delete')
          .setOnClickAction(CardService.newAction()
            .setFunctionName('handleDeleteKeyword')
            .setParameters({ keyword: kw.keyword }))));
    }
  }

  card.addSection(addSection);
  card.addSection(listSection);
  return card.build();
}

/**
 * Shows the trigger management card with current status and toggle buttons.
 *
 * @param {Object} [e] - Card action event.
 * @returns {CardService.Card}
 */
function showTriggerSettings(e) {
  const status = getTriggerStatus_();

  const card = CardService.newCardBuilder();
  card.setHeader(CardService.newCardHeader().setTitle('Trigger Settings'));

  const section = CardService.newCardSection();
  section.addWidget(CardService.newTextButton()
    .setText('← Back to Main')
    .setOnClickAction(CardService.newAction().setFunctionName('buildAddOn')));

  // Cleanup trigger
  section.addWidget(CardService.newDecoratedText()
    .setText('<b>Inbox Cleanup</b>')
    .setBottomLabel(status.cleanup ? '✓ Active — runs every 1 hour' : '✗ Not scheduled'));
  section.addWidget(CardService.newTextButton()
    .setText(status.cleanup ? 'Disable Hourly Cleanup' : 'Enable Hourly Cleanup')
    .setOnClickAction(CardService.newAction().setFunctionName('handleToggleCleanupTrigger')));

  // Discovery trigger
  section.addWidget(CardService.newDecoratedText()
    .setText('<b>Rule Discovery</b>')
    .setBottomLabel(status.discovery ? '✓ Active — runs daily' : '✗ Not scheduled'));
  section.addWidget(CardService.newTextButton()
    .setText(status.discovery ? 'Disable Daily Discovery' : 'Enable Daily Discovery')
    .setOnClickAction(CardService.newAction().setFunctionName('handleToggleDiscoveryTrigger')));

  // Info
  section.addWidget(CardService.newTextParagraph()
    .setText('<i>Triggers run automatically in the background. '
      + 'Hourly cleanup sweeps your inbox; daily discovery updates rules from label contents.</i>'));

  card.addSection(section);
  return card.build();
}

/**
 * Shows the import/export card for backing up and restoring rules.
 *
 * @param {Object} [e] - Card action event.
 * @returns {CardService.Card}
 */
function showImportExport(e) {
  const card = CardService.newCardBuilder();
  card.setHeader(CardService.newCardHeader().setTitle('Import / Export'));

  // Export section
  const exportSection = CardService.newCardSection().setHeader('Export Rules');
  exportSection.addWidget(CardService.newTextParagraph()
    .setText('Creates a timestamped JSON backup file in the same Drive folder as this script.'));
  exportSection.addWidget(CardService.newTextButton()
    .setText('📤 Export to Drive')
    .setOnClickAction(CardService.newAction().setFunctionName('handleExportRules')));

  // Import section
  const importSection = CardService.newCardSection().setHeader('Import Rules');
  importSection.addWidget(CardService.newTextParagraph()
    .setText('Enter the filename of a JSON backup in the script\'s Drive folder.'));
  importSection.addWidget(CardService.newTextInput()
    .setFieldName('import_filename')
    .setTitle('Filename (e.g. mail-rules-backup-2026-02-28_1430.json)'));
  importSection.addWidget(CardService.newButtonSet()
    .addButton(CardService.newTextButton()
      .setText('Merge Import')
      .setOnClickAction(CardService.newAction().setFunctionName('handleImportRules')))
    .addButton(CardService.newTextButton()
      .setText('Replace All')
      .setOnClickAction(CardService.newAction().setFunctionName('handleReplaceRules'))));

  // Stats section
  const statsSection = CardService.newCardSection().setHeader('Filing Stats');
  const filingStats = loadFilingStats_();
  const totalFiled = Object.values(filingStats).reduce((sum, n) => sum + n, 0);
  statsSection.addWidget(CardService.newTextParagraph()
    .setText('<b>' + totalFiled + '</b> threads filed across <b>'
      + Object.keys(filingStats).length + '</b> rules (lifetime).'));
  statsSection.addWidget(CardService.newTextButton()
    .setText('🗑️ Reset Filing Stats')
    .setOnClickAction(CardService.newAction().setFunctionName('handleResetStats')));

  // Nav
  const navSection = CardService.newCardSection();
  navSection.addWidget(CardService.newTextButton()
    .setText('← Back to Main')
    .setOnClickAction(CardService.newAction().setFunctionName('buildAddOn')));

  card.addSection(exportSection);
  card.addSection(importSection);
  card.addSection(statsSection);
  card.addSection(navSection);
  return card.build();
}


// ── UI Event Handlers ───────────────────────────────────────────────────────

/**
 * Handles the "Add & Clean Now" form submission.
 *
 * @param {Object} e - Card action event containing `formInput`.
 * @returns {CardService.ActionResponse}
 */
function handleUiSubmit(e) {
  const email = e.formInput.sender_email;
  const label = e.formInput.label_name;

  if (!email || !label) {
    return notify_('Please fill in both fields.');
  }

  addNewRule(email, label);
  return notify_('Rule added! Inbox cleaned for ' + email);
}

/**
 * Handles the "Sync Rules from Labels" button.
 * @returns {CardService.ActionResponse}
 */
function handleUiSync() {
  discoverRules();
  return notify_('Rules updated from existing labels!');
}

/**
 * Handles the "Run Full Inbox Sweep" button.
 * @returns {CardService.ActionResponse}
 */
function handleUiCleanup() {
  runGlobalCleanup();
  return notify_('Full inbox sweep complete.');
}

/**
 * Handles the "Auto Move (Small Labels)" button.
 *
 * @param {Object} e - Card action event containing `formInput`.
 * @returns {CardService.ActionResponse}
 */
function handleUiMigrate(e) {
  const oldName = e.formInput.old_label;
  const newName = e.formInput.new_label;

  if (!oldName || !newName) {
    return notify_('Please fill in both label names.');
  }

  migrateLabel(oldName, newName);
  return notify_('Renamed to ' + newName + ' and updated rules!');
}

/**
 * Handles the "Sync Rules Only (Big Labels)" button.
 *
 * @param {Object} e - Card action event containing `formInput`.
 * @returns {CardService.ActionResponse}
 */
function handleUiSyncOnly(e) {
  const oldName = e.formInput.old_label;
  const newName = e.formInput.new_label;

  syncJsonAfterManualRename(oldName, newName);
  return notify_('JSON rules updated to ' + newName);
}

/**
 * Handles the "Refresh Label Colors" button.
 * @returns {CardService.ActionResponse}
 */
function handleUiColor() {
  autoColorLabels();
  return notify_('Colors updated! Refresh Gmail to see changes.');
}

/**
 * Handles deleting a rule.
 *
 * @param {Object} e - Event containing parameters.sender.
 * @returns {CardService.ActionResponse}
 */
function handleDeleteRule(e) {
  const sender = e.parameters.sender;
  const success = deleteRule(sender);

  if (success) {
    // Return to the rule editor view
    return CardService.newActionResponseBuilder()
      .setNavigation(CardService.newNavigation().updateCard(showRuleEditor(e)))
      .setNotification(CardService.newNotification().setText('Rule deleted: ' + sender))
      .build();
  } else {
    return notify_('Error: Rule not found.');
  }
}

/**
 * Handles editing a rule (updating the label for an existing sender).
 *
 * @param {Object} e - Event containing parameters.sender and formInput.new_label.
 * @returns {CardService.ActionResponse}
 */
function handleEditRule(e) {
  const sender = e.parameters.sender;
  const newLabel = e.formInput.new_label;

  if (!newLabel) {
    return notify_('Please enter a label name.');
  }

  // Use addNewRule since it handles both add and update
  addNewRule(sender, newLabel);

  // Navigate back to rule editor
  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().popToRoot().updateCard(showRuleEditor({})))
    .setNotification(CardService.newNotification().setText('Rule updated: ' + sender + ' → ' + newLabel))
    .build();
}

/**
 * Handles applying consolidation opportunities.
 *
 * @returns {CardService.ActionResponse}
 */
function handleConsolidate(e) {
  const opportunities = findConsolidationOpportunities_();
  const changesCount = consolidateRules(opportunities);

  if (changesCount > 0) {
    return CardService.newActionResponseBuilder()
      .setNavigation(CardService.newNavigation().popToRoot())
      .setNotification(CardService.newNotification()
        .setText('✓ Consolidated ' + changesCount + ' rules successfully!'))
      .build();
  } else {
    return notify_('No changes made.');
  }
}

/**
 * Handles the Quick File action from the contextual message view.
 * Creates a rule for the sender/domain and files matching threads.
 *
 * @param {Object} e - Event containing parameters.sender and formInput.quick_label.
 * @returns {CardService.ActionResponse}
 */
function handleQuickFile(e) {
  const sender = e.parameters.sender;
  const label = e.formInput.quick_label;

  if (!label) {
    return notify_('Please enter a target label.');
  }

  addNewRule(sender, label);
  return notify_('Rule added: ' + sender + ' → ' + label);
}

/**
 * Handles the Dry Run Preview button. Runs cleanup in dry-run mode
 * and displays the results on a new card.
 *
 * @returns {CardService.ActionResponse}
 */
function handleUiDryRun(e) {
  const results = runGlobalCleanup(true);
  const card = showDryRunResults(results);

  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().pushCard(card))
    .build();
}

/**
 * Handles adding a new exclusion from the exclusion editor.
 *
 * @param {Object} e - Event containing formInput.exclusion_pattern.
 * @returns {CardService.ActionResponse}
 */
function handleAddExclusion(e) {
  const pattern = e.formInput.exclusion_pattern;
  if (!pattern) return notify_('Please enter an email or domain.');

  addExclusion(pattern);

  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().updateCard(showExclusionEditor(e)))
    .setNotification(CardService.newNotification().setText('Exclusion added: ' + pattern))
    .build();
}

/**
 * Handles removing an exclusion from the exclusion editor.
 *
 * @param {Object} e - Event containing parameters.pattern.
 * @returns {CardService.ActionResponse}
 */
function handleDeleteExclusion(e) {
  const pattern = e.parameters.pattern;
  removeExclusion(pattern);

  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().updateCard(showExclusionEditor(e)))
    .setNotification(CardService.newNotification().setText('Exclusion removed: ' + pattern))
    .build();
}

/**
 * Handles adding a new keyword rule from the keyword editor.
 *
 * @param {Object} e - Event containing formInput.keyword_text and formInput.keyword_label.
 * @returns {CardService.ActionResponse}
 */
function handleAddKeyword(e) {
  const keyword = e.formInput.keyword_text;
  const label = e.formInput.keyword_label;

  if (!keyword || !label) return notify_('Please fill in both fields.');

  addKeywordRule(keyword, label);

  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().updateCard(showKeywordEditor(e)))
    .setNotification(CardService.newNotification().setText('Keyword rule added: "' + keyword + '" → ' + label))
    .build();
}

/**
 * Handles deleting a keyword rule from the keyword editor.
 *
 * @param {Object} e - Event containing parameters.keyword.
 * @returns {CardService.ActionResponse}
 */
function handleDeleteKeyword(e) {
  const keyword = e.parameters.keyword;
  deleteKeywordRule(keyword);

  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().updateCard(showKeywordEditor(e)))
    .setNotification(CardService.newNotification().setText('Keyword rule deleted: "' + keyword + '"'))
    .build();
}

/**
 * Toggles the hourly cleanup trigger on or off.
 *
 * @returns {CardService.ActionResponse}
 */
function handleToggleCleanupTrigger(e) {
  const status = getTriggerStatus_();
  if (status.cleanup) {
    disableCleanupTrigger();
  } else {
    enableCleanupTrigger();
  }

  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().updateCard(showTriggerSettings(e)))
    .setNotification(CardService.newNotification()
      .setText(status.cleanup ? 'Cleanup trigger disabled.' : 'Cleanup trigger enabled (every 1 hour).'))
    .build();
}

/**
 * Toggles the daily discovery trigger on or off.
 *
 * @returns {CardService.ActionResponse}
 */
function handleToggleDiscoveryTrigger(e) {
  const status = getTriggerStatus_();
  if (status.discovery) {
    disableDiscoveryTrigger();
  } else {
    enableDiscoveryTrigger();
  }

  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().updateCard(showTriggerSettings(e)))
    .setNotification(CardService.newNotification()
      .setText(status.discovery ? 'Discovery trigger disabled.' : 'Discovery trigger enabled (daily).'))
    .build();
}

/**
 * Handles exporting rules to a Drive backup file.
 *
 * @returns {CardService.ActionResponse}
 */
function handleExportRules(e) {
  const filename = exportRulesToDrive();
  return notify_('Exported to ' + filename);
}

/**
 * Handles merging rules from an imported file (adds new, updates existing).
 *
 * @param {Object} e - Event containing formInput.import_filename.
 * @returns {CardService.ActionResponse}
 */
function handleImportRules(e) {
  const filename = e.formInput.import_filename;
  if (!filename) return notify_('Please enter a filename.');

  const result = importRulesFromFile(filename, false);
  return notify_(result.message);
}

/**
 * Handles replacing all rules from an imported file.
 *
 * @param {Object} e - Event containing formInput.import_filename.
 * @returns {CardService.ActionResponse}
 */
function handleReplaceRules(e) {
  const filename = e.formInput.import_filename;
  if (!filename) return notify_('Please enter a filename.');

  const result = importRulesFromFile(filename, true);
  return notify_(result.message);
}

/**
 * Handles resetting all filing statistics.
 *
 * @returns {CardService.ActionResponse}
 */
function handleResetStats(e) {
  resetFilingStats();

  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().updateCard(showImportExport(e)))
    .setNotification(CardService.newNotification().setText('Filing stats reset.'))
    .build();
}
