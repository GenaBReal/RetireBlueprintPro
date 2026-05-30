# RetireBlueprint Pro — Solver Correctness Fixes (Session 2)

All values verified against current IRS / CMS sources (Rev. Proc. 2025-32, CMS 2026 Medicare). Every page re-rendered with zero JavaScript errors.

## Files changed (upload these)
1. `RetireBlueprintPro_TaxEngine.html`
2. `RetireBlueprintPro_RothLadder.html`
3. `RetireBlueprintPro_Inputs.html`
4. `RetireBlueprintPro_Medicare.html`
5. `RetireBlueprintPro_SetupGuide.html`

---

## Tax tables brought current to 2026

**Federal income tax brackets (MFJ).** Were 2025 values displayed under a "2026" label. Now the actual 2026 thresholds: 10% to $24,800 · 12% to $100,800 · 22% to $211,400 · 24% to $403,550 · 32% to $512,450 · 35% to $768,700 · 37% above. Updated in the Tax Engine bracket chart, the tax calculation, the subtitle, and the Inputs bracket projector (rebased to 2026, then projected forward at 3%/yr). *Verified: $114,523 taxable → $14,619 federal tax, 12.77% effective.*

**Standard deduction (MFJ).** $30,000 → **$32,200**.

**Long-term capital gains thresholds (MFJ).** Were 2024 values ($94,050 / $583,750). Now 2026: 0% up to **$98,900**, 15% to **$613,700**, 20% above.

**Medicare IRMAA + Part B premium.** Medicare.html carried its own 2025 IRMAA table. Updated to 2026: base Part B premium **$202.90** (was $185), couple thresholds $218K / $274K / $342K / $410K / $750K, with correct per-tier surcharges. *Verified every tier: $202.90 → $284.10 → $405.80 → $527.50 → $649.20 → $689.90.* Part B premium references throughout the page updated to $202.90.

**Roth Ladder bracket-fill estimate.** Used the 2024 12%-bracket top ($94,300); now the 2026 figure ($100,800).

---

## Logic bug fixed: RMD age (SECURE 2.0)

This was a genuine correctness error, not just stale data. Required Minimum Distributions begin at **age 73 for those born 1951–1959, and age 75 for those born 1960 or later**. The code hardcoded **73** everywhere.

- In the Tax Engine, the dynamic RMD note targeted an element ID that didn't exist (`rmd-partner-note`), so it silently never ran and the static "75" text showed. Rewired to the real elements and made it compute each partner's age from birth year.
- Roth Ladder (conversion planner **and** ladder builder), Inputs (Roth solver, RMD-year auto-populate), and the Medicare timeline all now compute RMD age from birth year.
- The Inputs RMD estimate also used the wrong divisor — the IRS Uniform Lifetime Table first-year factor is **26.5 at age 73 but 24.6 at age 75**; now selected by the correct start age.
- Educational copy in the Tax Engine and Setup Guide updated to state "73 or 75 by birth year."

For your sample couple (born ~1963 and ~1971), the correct RMD age is **75** — previously shown/calculated as 73.

---

## Verified correct, left unchanged
- The RMD Uniform Lifetime Table distribution periods (26.5 at 73, 24.6 at 75, … 6.4 at 100).
- The Tax Engine's IRMAA reference table (was already on 2026 values).

## Scope note
These pages use **Married Filing Jointly** brackets. A single filer would see MFJ thresholds applied to their income — correct for the sample couple, but worth knowing if you market to single retirees. Adding filing-status awareness would be a future enhancement.
