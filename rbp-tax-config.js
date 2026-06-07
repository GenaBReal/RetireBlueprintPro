/* ============================================================================
   RetireBlueprint Pro — TAX CONSTANTS  (the one file to update each year)
   ----------------------------------------------------------------------------
   HOW TO USE THIS FILE:
   Once a year, when the IRS and Medicare publish new numbers (usually October–
   November), change the values in the RBP_TAX block below, save, and re-upload
   THIS ONE FILE to your website. Every page updates automatically.

   You only ever need to touch the numbers between the lines marked
   "EDIT THESE NUMBERS". Leave everything below "DO NOT EDIT BELOW" alone.

   Currently set for the 2026 tax year.
   ============================================================================ */

/* ----------------------------- EDIT THESE NUMBERS ----------------------------- */
window.RBP_TAX = {

  /* The tax year these numbers are for (shown on some pages) */
  year: 2026,

  /* FEDERAL INCOME TAX BRACKETS — Married Filing Jointly.
     "max" is the top dollar of each bracket. The last (37%) bracket has no top,
     so it stays as Infinity. (Source: IRS Revenue Procedure, published each fall.) */
  mfjBrackets: [
    { rate: 0.10, max: 24800 },
    { rate: 0.12, max: 100800 },
    { rate: 0.22, max: 211400 },
    { rate: 0.24, max: 403550 },
    { rate: 0.32, max: 512450 },
    { rate: 0.35, max: 768700 },
    { rate: 0.37, max: Infinity }
  ],

  /* STANDARD DEDUCTION by filing status */
  stdDed: {
    'Married Filing Jointly':    32200,
    'Married Filing Separately': 16100,
    'Single':                    16100,
    'Head of Household':         24150
  },

  /* MEDICARE IRMAA — base Part B premium plus the income tiers.
     partBBase     = the standard monthly Part B premium everyone pays.
     singleMax     = top income for that tier, Individual filing.
     mfjMax        = top income for that tier, Married-Filing-Jointly (couple).
     partBSurcharge= extra $/month ADDED to the base premium in that tier.
     partD         = extra Part D $/month surcharge in that tier.
     (Source: CMS, published each fall.) */
  irmaa: {
    partBBase: 202.90,
    partBDeductible: 283,   /* annual Part B deductible (CMS, published each fall) */
    tiers: [
      { label: 'Standard', singleMax: 109000,   mfjMax: 218000,   partBSurcharge: 0,      partD: 0     },
      { label: 'Tier 1',   singleMax: 137000,   mfjMax: 274000,   partBSurcharge: 81.20,  partD: 14.50 },
      { label: 'Tier 2',   singleMax: 171000,   mfjMax: 342000,   partBSurcharge: 202.90, partD: 37.50 },
      { label: 'Tier 3',   singleMax: 205000,   mfjMax: 410000,   partBSurcharge: 324.60, partD: 60.40 },
      { label: 'Tier 4',   singleMax: 500000,   mfjMax: 750000,   partBSurcharge: 446.30, partD: 83.30 },
      { label: 'Tier 5',   singleMax: Infinity, mfjMax: Infinity, partBSurcharge: 487.00, partD: 91.00 }
    ]
  },

  /* ACA / HEALTH-INSURANCE subsidy estimate */
  acaFpl2Person: 21150   /* 2-person federal poverty level used for the subsidy estimate */

};
/* --------------------------- END OF NUMBERS TO EDIT --------------------------- */


/* ============================ DO NOT EDIT BELOW ==============================
   Helper that fills any plain-text spots on a page (like a sentence that says
   "the premium is $202.90/month") straight from the numbers above, so those
   read-outs stay in sync too. You don't need to change anything here. */
window.RBP_TAX.applyText = function () {
  var C = window.RBP_TAX; if (!C || !C.irmaa) return;
  var base = C.irmaa.partBBase;
  var ded = C.irmaa.partBDeductible || 283;
  var firstMfj = C.irmaa.tiers[0].mfjMax;
  var map = {
    'rbp-pb-base':     '$' + base.toFixed(2),
    'rbp-pb-base-mo':  '$' + base.toFixed(2) + '/mo',
    'rbp-pb-deductible': '$' + ded.toLocaleString(),
    'rbp-year':        '' + C.year,
    'rbp-irmaa-mfj1':  '$' + firstMfj.toLocaleString(),
    'rbp-irmaa-mfj1k': '$' + Math.round(firstMfj / 1000) + 'K',
    'rbp-current-year': '' + (new Date().getFullYear())
  };
  Object.keys(map).forEach(function (cls) {
    var els = document.querySelectorAll('.' + cls);
    for (var i = 0; i < els.length; i++) { els[i].textContent = map[cls]; }
  });
};

(function () {
  var run = function () { try { window.RBP_TAX.applyText(); } catch (e) {} };
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', run);
  } else {
    run();
  }
})();
/* ============================================================================ */
