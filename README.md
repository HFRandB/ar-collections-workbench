# AR Collections Workbench

A browser-based workbench for Greater China AR collections — CC0038 · CC0073 · CC0074.

**Live site:** https://hfrandb.github.io/ar-collections-workbench/

---

## Pages

| Page | Description |
|---|---|
| 📊 **Summary** | KPI totals, EHF exposure, risk overview, FC Month EUR breakdown, income collected |
| 📈 **Activity** | Daily edit chart, full edit log with filters and CSV export |
| 📋 **Worklist** | Account cards with inline editing, dunning, FDL, unblock request |
| 🎯 **Action Plan** | Expert-generated collection recommendations per account with LOB breakdown |

---

## Getting Started

1. Open the launcher at the live site URL above
2. Go to **Worklist → Data Source panel**
3. Drop your **Master Report_GC_YYYYMM.xlsx** file
4. Also drop **dunning contacts** file to load contact emails
5. Click **⟳ Refresh Worklist** — all 4 pages update automatically

---

## Key Features

**Worklist**
- 1,000+ accounts grouped by Sold-to Party, sorted by AR amount
- Risk tags: ⚠ Termination Risk · EHF Impact · ⛔ Order Block Risk
- Filters: Entity / Priority / Risk / Collector / Customer / Search (all stackable)
- ▸ Details — per-doc inline editing: FC Month, Customer Payment Process, Current Status
- Mark Paid — strikethrough + income tracking
- CDC A/R badge on docs billed via China Datacom

**Actions per account**
- ✉ **Dunning** — pre-filled email in ZH-CN / ZH-TW / EN by entity
- ⚑ **FDL** — Final Demand Letter matching Word templates, export to PDF
- 🔓 **Unblock** — Unblock request with AR table (local amt, EUR, OD days ME/QE, EHF ME/QE), export to Excel

**Quick Links** (on Worklist)
- 📄 Contracts (CMS)
- 👤 Update Contacts (CRM)
- 🤖 Ask CC AI

**Data**
- Empty by default — all data loaded from user-uploaded files only
- No personal contact data stored in the repository
- Edits (FC Month, status, paid marks) saved to browser localStorage per user
- Export: Full Report CSV · EHF Assessment CSV · Edit History CSV · Contacts CSV

---

## Data Sources

| File | Content |
|---|---|
| `Master Report_GC_YYYYMM.xlsx` | AR data — all 3 CCs (required) |
| `dunning contacts in FICA.xlsx` | Dunning contact emails |
| `2026 Direct Contact list.xlsx` | Additional CC0038 contacts |

---

## Repository

https://github.com/HFRandB/ar-collections-workbench
