# evictionlawtracker

## Goal
Build a **longitudinal dataset of state and jurisdiction eviction laws**. Today, LSC's Eviction Laws Database (built with Temple's Center for Public Health Law Research) is a single snapshot of laws in effect on **January 1, 2021**. This project extends coverage **back to the 2000s and forward to the present**, then keeps it updated.

## Approach
1. **Source text.** Downloaded state statutes and codes are stored as `.xml` files, organized as `Jurisdiction / Year-of-change / *.xml`. A year folder exists only for years in which the law changed.
2. **Anchor workbook.** A workbook we maintain (`EvictionLawTracker.xlsx`) lists high priority eviction-related law questions for each jurisdiction, along with the controlling citation and its description as of January 2021, taken from the LSC data.
3. **Linking.** Each workbook citation is matched to the corresponding section in the XML for every earlier and later year. Citations can be renumbered or renamed over time, so the matching process uses the **citation name/number** and also the **text of the law**.
4. **Output.** For each jurisdiction × question × year: the controlling section, and from that a view of how the law changed over time.

## Setup
Requires Python 3.9 or newer. From the repo folder:

```bash
python3 -m venv .venv                      # create a virtual environment (once)
source .venv/bin/activate                  # activate it (each new terminal)
python3 -m pip install --upgrade pip
python3 -m pip install -r requirements.txt
python3 match_statutes.py --self-test      # should print "Self-test passed."
```

Then run the matcher with `./run_match_statutes.sh`.

## Current state (Oct 2026)
- **Workbook:** 306 rows = 51 jurisdictions (50 states + DC) × 6 high-priority questions:
  - Is there a law?
  - Eviction cause
  - Just cause
  - Late fee regulation
  - Min notice, nonpayment
  - Min notice, other
- **Columns:** `Jurisdiction`, `Question`, `Citation/Statute` (filled in from LSC 2021), and `XML Title`, `Chapter`, `Section` (link-out columns, still empty).
- **Sample XML:** Alabama 2009, Title 1 (*Code of Alabama*, edition date 01/01/2009). The file is a nested hierarchy: Title → Chapter → Section. Each section has:
  - `Caption`
  - `Description` (heading)
  - `ShortName` (e.g., `Ala. Code 1-1-1`)
  - `Content` (HTML-escaped statute text)
  - `RevisionHistory` (session-law history)

## Matching notes
- `ShortName` is the natural join key. Workbook citations need to be normalized to it first. For example, `Ala. Code § 35-9A-141. Definitions.` becomes `Ala. Code 35-9A-141`, plus the heading `Definitions`.
- A section heading (`Description`) together with text similarity can catch renumbered or renamed sections. `RevisionHistory` can confirm that two versions are the same provision across years.
- Note: Alabama's landlord-tenant law sits in Title 35, so the Title 1 sample shows the format but doesn't hold any target sections.