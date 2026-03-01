# 📧 Inbox Cleaner – Gmail Add-on

Automatically organize your Gmail inbox by filing messages from specific senders (or entire domains) into labeled folders. Perfect for managing **newsletters**, **notifications**, **marketing emails**, and **bulk email organization**. Rules are stored in a JSON config file and can be **discovered automatically** from your existing label structure or added manually via a Gmail side-panel UI.

---

## ❓ Why Inbox Cleaner?

Gmail's native filters are powerful but clunky — they require manual rule-by-rule configuration and lack automation for common patterns. **Inbox Cleaner** provides a simpler, faster alternative:

| Feature | Gmail Filters | Inbox Cleaner |
| --------- | --- | --- |
| Auto-discover rules from existing labels | ❌ | ✅ |
| Visual rule editor & management | ❌ | ✅ |
| Bulk cleanup in one click | ❌ | ✅ |
| Test changes before applying (dry run) | ❌ | ✅ |
| Track filing stats over time | ❌ | ✅ |

Built on **Google Apps Script**, this add-on integrates seamlessly with Gmail without plugins or third-party services.

---

## ✨ Features

- **🔍 Auto-Discovery** – Email automation that scans your existing labels and learns sender → label mappings automatically
- **🧹 Bulk Cleanup** – Perform bulk email organization by sweeping your inbox and filing all matching threads in one pass
- **🌐 Domain Wildcards** – Use `@company.com` patterns for rule-based filing that matches everyone from a domain
- **📋 Rule Editor** – View, search, edit, and delete rules from the UI with pagination
- **🔄 Smart Consolidation** – Automatically detects and merges similar rules to keep config lean
- **⚡ Quick File** – One-click rule creation from the currently viewed email
- **🔍 Dry Run** – Preview what cleanup would do before moving any threads
- **🚫 Exclusion List** – Whitelist senders/domains to never auto-file, even if rules match
- **🔑 Keyword Rules** – Match on subject-line keywords in addition to sender rules
- **📊 Filing Stats** – Track how many threads each rule has filed over its lifetime
- **⏰ Trigger Management** – Enable/disable hourly inbox automation and daily discovery from the UI
- **📦 Import/Export** – Backup and restore your email organization rules as timestamped Drive files
- **🎨 Label Color Themes** – Apply consistent colors to label hierarchies for better email management
- **♻️ Label Migration** – Rename/reorganize labels and automatically update your email filtering rules
- **⚡ Optimized** – Batch API calls and timeout protection prevent script failures

---

## � Common Use Cases

**For newsletter junkies:**

- Automatically file newsletters from Substack, Medium, Patreon, and Medium into a `Newsletters` label
- Use keyword rules to catch digest emails that don't always come from the same sender

**For notification management:**

- Route GitHub notifications, CI/CD alerts, and monitoring emails into separate labels
- Enable automated hourly cleanup to keep your inbox free of operational clutter

**For teams managing multiple email accounts:**

- Set up rules that match entire domains (`@company.com`, `@partner.org`)
- Use bulk cleanup for periodic inbox maintenance across different label hierarchies

**For label power users:**

- Migrate and reorganize your label structure without losing email organization rules
- Use the dry run feature to preview changes before committing to bulk operations

---

## �🚀 Setup

### 1. Prerequisites

- A Google account with Gmail
- [Google Apps Script CLI (clasp)](https://github.com/google/clasp) installed (optional, for local development)

### 2. Installation

#### Option A: Via Apps Script Editor

1. Go to [script.google.com](https://script.google.com)
2. Create a new project
3. Copy the contents of `Code.js` and `appsscript.json` into your project
4. Enable the **Gmail API** advanced service:
   - Click ⚙️ **Project Settings** → **Services**
   - Add **Gmail API v1**

#### Option B: Via clasp (Local)

```bash
# Clone this repo
git clone https://github.com/zmongeau/Inbox-Cleaner.git
cd Inbox-Cleaner

# Log in to Google
clasp login

# Create a new Apps Script project
clasp create --title "Inbox Cleaner" --type standalone

# Push the code
clasp push
```

### 3. Deploy as Add-on

1. In the Apps Script editor, click **Deploy** → **Test deployments**
2. Select **Install** to enable the add-on in your Gmail
3. Open Gmail and look for the **Inbox Cleaner** icon in the right sidebar

---

## 📖 How It Works

Built with **Google Apps Script** and the **Gmail API**, Inbox Cleaner automates email filtering through three core phases:

### Core Workflow

```md
┌─────────────┐       ┌──────────────┐       ┌─────────────┐
│  Discovery  │ ----> │ Config File  │ ----> │   Cleanup   │
│             │       │ (JSON rules) │       │             │
└─────────────┘       └──────────────┘       └─────────────┘
```

1. **Discovery Phase** (`discoverRules()`)
   - Scans up to 30 recent threads per label
   - Extracts sender email from the first message
   - Builds a JSON map: `{ "sender@domain.com": "LabelName" }`

2. **Cleanup Phase** (`runGlobalCleanup()`)
   - Fetches all inbox threads
   - Checks each sender against rules (exact match → domain wildcard → keyword)
   - Respects exclusion list — excluded senders are never filed
   - Applies label + archives matching threads
   - Tracks filing stats per rule

3. **Manual Additions** (`addNewRule()`)
   - Add one-off rules from the UI
   - Immediately files existing inbox threads for that sender

### Rule Priority

When multiple rules could match a thread, the following priority applies:

1. **Exclusions** — If sender/domain is excluded, skip entirely
2. **Exact email** — `john@acme.com` matches first
3. **Domain wildcard** — `@acme.com` matches any sender from that domain
4. **Keyword rules** — Subject-line keyword matches (lowest priority)

---

## 🎮 Usage

### Using the Gmail Side Panel

Open Gmail → Click the **Inbox Cleaner** icon in the right sidebar.

#### **Quick File (Contextual)**

When viewing an email, the add-on shows the sender's email and whether a rule already exists:

- **File Sender** — Create a rule for the exact sender email
- **File @domain** — Create a rule for the entire domain
- Both buttons use the target label you enter

#### **Section 1: Add New Rule**

- **Sender Email**: Enter `john@acme.com` or `@acme.com` for a domain wildcard
- **Target Label**: Enter `Work/Invoices` (use `/` for sub-labels)
- Click **Add & Clean Now** to save the rule and file existing threads

#### **Section 2: Manage Rules**

- **View Rule Stats** – See count of exact email rules, domain rules, keyword rules, and exclusions
- **View/Edit/Delete Rules** – Browse all rules with search, pagination, and filing stats:
  - 📧 = exact sender email
  - 🌐 = domain wildcard rule
  - Filing count shown next to each rule
  - Edit button to change the target label
  - Delete button to remove rules
- **Search** – Filter rules by sender email or label name
- **Consolidate Similar Rules** – Merge redundant or similar rules automatically

#### **Section 3: Reorganize Labels**

- **Small Labels** (< 200 threads): Use **Auto Move** to rename and update config
- **Large Labels**:
  1. Manually rename the label in Gmail sidebar
  2. Use **Sync Rules Only** to update the JSON config

#### **Section 4: Settings**

- **Sync Rules from Labels** – Runs discovery across all your labels
- **Run Full Inbox Sweep** – Processes your entire inbox
- **Refresh Label Colors** – Applies color themes to top-level labels

#### **Section 5: Advanced**

- **🔍 Dry Run Preview** – See what cleanup would do without moving anything, then optionally proceed
- **🚫 Manage Exclusions** – Add/remove senders and domains that should never be auto-filed
- **🔑 Keyword Rules** – Add/remove subject-line keyword matching rules
- **⏰ Trigger Settings** – Enable/disable hourly cleanup and daily discovery automation
- **📦 Import / Export** – Backup rules to Drive, import from backup files, view/reset filing stats

---

### Dry Run Preview

Before running a full sweep, preview what would happen:

1. Click **🔍 Dry Run Preview** in the Advanced section
2. See a summary grouped by target label (thread count + sender count)
3. Click **▶ Run Full Sweep Now** to proceed, or go back to cancel

---

### Exclusion List

Prevent specific senders or domains from ever being auto-filed:

1. Click **🚫 Manage Exclusions**
2. Enter an email (`boss@company.com`) or domain (`@company.com`)
3. Click **Add Exclusion**

Excluded senders are skipped during:

- Full inbox sweeps (`runGlobalCleanup`)
- Immediate filing when adding new rules (`addNewRule`)

Rules for excluded senders are preserved — remove the exclusion to resume filing.

---

### Keyword Rules

Match emails by subject line instead of sender:

1. Click **🔑 Keyword Rules**
2. Enter a keyword (case-insensitive substring match)
3. Enter a target label
4. Click **Add Keyword Rule**

**Priority:** Keyword rules have the lowest priority. Sender and domain rules are always checked first. An email is only matched by keyword if no sender/domain rule applies.

**Example:**

- Keyword: `invoice` → Label: `Finance/Invoices`
- Any email with "invoice" in the subject gets filed (unless a sender rule matches first)

---

### Trigger Management

Automate cleanup and discovery without leaving Gmail:

1. Click **⏰ Trigger Settings**
2. **Inbox Cleanup** — Toggle hourly automatic sweeps
3. **Rule Discovery** — Toggle daily automatic label scanning

Status shows whether each trigger is active or inactive. Triggers run in the background using Apps Script's time-based trigger system.

---

### Import / Export

Back up your rules and restore them:

#### Export

1. Click **📦 Import / Export**
2. Click **📤 Export to Drive**
3. A timestamped JSON backup file is created in your script's Drive folder

#### Import

1. Enter the backup filename (e.g., `mail-rules-backup-2026-02-28_1430.json`)
2. **Merge Import** — Adds new rules, updates existing ones (doesn't delete anything)
3. **Replace All** — Overwrites all rules with the imported file

#### Filing Stats

- View lifetime filing counts across all rules
- **Reset Filing Stats** to clear counters

---

### Manage Rules from the UI

Click **View/Edit/Delete Rules** to access the rule editor:

#### **Viewing Rules**

- Displays rules in pages of 10 (with Previous/Next navigation)
- Shows sender → label mappings with filing counts
- Icons: 📧 for exact emails, 🌐 for domain rules
- Search box to filter by sender email or label name

#### **Editing a Rule**

1. Click **Edit** next to any rule
2. Update the target label name
3. Click **Save Changes** to apply
4. The rule updates immediately and files matching inbox threads

#### **Deleting a Rule**

1. Click **Delete** next to any rule
2. Rule is removed from config and logs
3. Future emails from that sender will no longer be filed

### Rule Consolidation

Clean up your rules by merging similar patterns. Click **🔄 Consolidate Similar Rules** to:

#### **What Gets Consolidated**

**Type 1: Redundant Rules** 🗑️

- Exact emails already covered by domain rules
- Example: `alice@company.com` is removed if `@company.com` → same label exists

**Type 2: Domain Merging** 🔀

- Multiple subdomains that can merge to parent
- Example: `@sub1.company.com`, `@sub2.company.com` → `@company.com`
- Also: 3+ exact emails from same domain → single domain rule
- Example: `alice@co.com`, `bob@co.com`, `charlie@co.com` → `@co.com`

#### **How to Use**

1. Click **🔄 Consolidate Similar Rules**
2. Review the consolidation opportunities shown
3. Click **✓ Apply All Consolidations** to merge rules
4. Config is updated and logs show each change

#### **Example: Before & After**

**Before:**

```json
{
  "alice@one.gmail.com": "Personal",
  "bob@two.gmail.com": "Personal",
  "charlie@three.gmail.com": "Personal",
  "@one.gmail.com": "Personal",
  "@two.gmail.com": "Personal"
}
```

**After Consolidation:**

```json
{
  "@gmail.com": "Personal"
}
```

**Changes Made:**

- 🗑️ Removed exact emails (redundant)
- 🔀 Merged subdomains to parent domain
- **Result:** 5 rules → 1 rule

### Using the Script Editor

Run functions directly from the Apps Script editor for advanced workflows:

```javascript
// Discover rules from existing labels
discoverRules();

// Clean the inbox based on current rules
runGlobalCleanup();

// Dry-run preview (returns array of matches without filing)
const preview = runGlobalCleanup(true);

// Add a single rule manually
addNewRule('newsletter@example.com', 'Newsletters');

// Add a domain wildcard rule
addNewRule('@github.com', 'Dev/GitHub');

// Get all rules (optionally filtered)
const allRules = getRules();           // all rules
const filtered = getRules('github');   // search for 'github'

// Get rule statistics
const stats = getRuleStats();
// { totalRules: 47, exactRules: 32, domainRules: 15, exclusionCount: 3, keywordCount: 5 }

// Find and apply consolidation opportunities
const opportunities = findConsolidationOpportunities_();
consolidateRules(opportunities);

// Manage exclusions
addExclusion('boss@company.com');      // never file this sender
addExclusion('@personal.com');         // never file this domain
removeExclusion('boss@company.com');
const excluded = getExclusions();

// Manage keyword rules
addKeywordRule('invoice', 'Finance/Invoices');
addKeywordRule('shipping confirmation', 'Orders');
deleteKeywordRule('invoice');
const keywords = getKeywordRules();

// Filing stats
const filingStats = getFilingStats();  // { "user@co.com": 42, "@domain.com": 128 }
resetFilingStats();

// Import/export
const backupFile = exportRulesToDrive();
importRulesFromFile('backup.json', false);  // merge
importRulesFromFile('backup.json', true);   // replace all

// Trigger management
enableCleanupTrigger();    // hourly cleanup
disableCleanupTrigger();
enableDiscoveryTrigger();  // daily discovery
disableDiscoveryTrigger();

// Migrate a label (small)
migrateLabel('Old Label', 'Work/New Label');

// Update rules after manual rename (large labels)
syncJsonAfterManualRename('Old Label', 'Work/New Label');

// Apply color themes
autoColorLabels();

// Delete a specific rule
deleteRule('old@email.com');
```

---

## ⚙️ Configuration

### Constants (in `Code.js`)

```javascript
const CONFIG_FILENAME = 'mail-rules.json';     // Rules file name
const THREAD_LIMIT_PER_LABEL = 30;              // Threads inspected per label
const MAX_DISCOVERY_TIME_MS = 5 * 60 * 1000;   // 5-minute timeout
const AUTO_CONSOLIDATE = false;                 // Auto-consolidate after adding rules
```

**Note on AUTO_CONSOLIDATE:**

- Set to `true` to automatically consolidate rules after each `addNewRule()` call
- Set to `false` (default) to review consolidation opportunities before applying them
- Recommended: Keep disabled and use the **Consolidate Similar Rules** button to review changes first

### Rules File Structure (`mail-rules.json`)

Stored in the same Drive folder as your Apps Script project:

```json
{
  "alice@example.com": "Work/Clients",
  "bob@company.org": "Projects/Alpha",
  "@newsletter.com": "Subscriptions",
  "@github.com": "Dev/GitHub"
}
```

- **Exact matches** take priority over domain wildcards
- Domain wildcards start with `@`

### Extended Data Storage

Exclusions, keyword rules, and filing stats are stored in **UserProperties** (via `PropertiesService`), keeping the main config file clean and backward-compatible:

| Property Key | Format | Description |
| --- | --- | --- |
| `exclusions` | JSON array | `["boss@co.com", "@personal.com"]` |
| `keyword_rules` | JSON object | `{"invoice": "Finance", "shipping": "Orders"}` |
| `filing_stats` | JSON object | `{"@github.com": 142, "keyword:invoice": 28}` |

### Color Themes

Edit the `THEMES` object in `autoColorLabels()` to customize:

```javascript
const THEMES = {
  '@Shopping':      { bg: '#fb4c2f', txt: '#ffffff' }, // Red
  '@Work':          { bg: '#4a86e8', txt: '#ffffff' }, // Blue
  // ... add your own parent labels here
};
```

⚠️ **Only use official Gmail API color codes** to avoid errors.

---

## 🛠️ Triggers

### Manual Setup (Apps Script Editor)

1. In Apps Script editor: **Triggers** (clock icon)
2. Add trigger:
   - **Function**: `runGlobalCleanup`
   - **Event**: `Time-driven` → `Hour timer` → `Every hour`
3. Repeat for `discoverRules` (run daily or weekly)

### UI-Managed Triggers

Manage triggers directly from the Gmail side panel:

1. Click **⏰ Trigger Settings** in the Advanced section
2. Toggle **Hourly Cleanup** — runs `runGlobalCleanup` every hour
3. Toggle **Daily Discovery** — runs `discoverRules` every 24 hours

Status indicators show whether each trigger is active or inactive.

---

## 📚 API Permissions

The add-on requests these OAuth scopes:

| Scope | Purpose |
| ------- | --------- |
| `gmail.addons.execute` | Run as a Gmail Add-on |
| `gmail.readonly` | Read message metadata |
| `gmail.modify` | Apply labels and archive |
| `gmail.addons.current.message.metadata` | Access current message sender for Quick File |
| `drive` | Access the config file and import/export backups |
| `script.container.ui` | Display the side panel |
| `script.locale` | Detect user locale for formatting |

---

## 🐛 Troubleshooting

### "Invalid Color" Error

- Use only the hex codes from the official Gmail palette (see `THEMES` in `autoColorLabels()`)

### Discovery Times Out

- Reduce `THREAD_LIMIT_PER_LABEL` to `15` or lower
- The script saves progress when hitting `MAX_DISCOVERY_TIME_MS`

### Rules Not Applying

1. Check the config file exists in Drive (same folder as the script)
2. Run `loadConfig_()` and check logs for JSON parse errors
3. Verify sender email format (check logs after `extractEmail_()`)
4. Check if the sender is in the exclusion list (`getExclusions()`)

### Quick File Not Showing

- Quick File only appears when viewing an individual email
- Check that `appsscript.json` has both `homepageTrigger` and `contextualTriggers`

### Add-on Not Showing in Gmail

- Go to **Deploy** → **Test deployments** → **Install**
- Check that `appsscript.json` has the `addOns` section correctly configured

### Keyword Rules Not Matching

- Keywords match as case-insensitive substrings on the subject line only
- Sender and domain rules always take priority over keyword rules
- Check that the sender isn't in the exclusion list

---

## 🤝 Contributing

Contributions welcome! Areas for improvement:

- Batch processing for extremely large inboxes
- Regex-based sender and keyword matching
- Undo log for recent filing actions
- Gmail filter import (convert native Gmail filters to rules)
- Multi-action rules (mark read, star, forward)
- Integration with Google Workspace admin features

---

## 📄 License

MIT License – See [LICENSE](LICENSE) file for details.

---

## 👤 Author

Community-maintained open-source project.

---

## 🤖 AI Disclosure

This project was conceived and directed by a human, with the implementation written entirely using AI assistance (Claude). The codebase reflects iterative refinement through AI-guided development, testing, and documentation.
