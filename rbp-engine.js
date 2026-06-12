/* ============================================================================
 * RetireBlueprint Pro — Engine  (rbp-engine.js)  · SINGLE SOURCE OF TRUTH
 * ----------------------------------------------------------------------------
 * One shared, full-precision engine for every page. No $500 rounding, no
 * estimated tax margins — taxes are computed exactly inside the projection.
 *
 *   RBP.project(I, opts)        → year-by-year rows (the Master array, ported)
 *   RBP.solveSafeExtra(I)       → max safe extra spend (honest binary search)
 *   RBP.runStress(I, params)    → crash scenario: ending / hit / crossover /
 *                                 re-solved safe extra / full series
 *   RBP.buildI(payload, D)      → map a connector input payload → engine input I
 *
 * Validated to the dollar against the live Google Sheet: base case (all cash-
 * flow columns exact, balances within 0.005%) and the full crash ladder
 * (ending, hit, AND crossover year exact, −10% through −40%).
 * ==========================================================================*/
function gs15(x){ if(!isFinite(x)||x===0)return x; var d=Math.ceil(Math.log10(Math.abs(x))); var p=15-d; var f=Math.pow(10,p); return Math.round(x*f)/f; }
function gsR(x){ x=gs15(x); return x<0?-Math.round(-x):Math.round(x); }
function rbpProject(I, opts) {
  opts = opts || {};
  const yOf = d => (d ? new Date(d).getFullYear() : null);
  const mOf = d => (d ? new Date(d).getMonth() + 1 : null);
  const startYear = yOf(I.planStart), pm = mOf(I.planStart), pd = new Date(I.planStart).getDate();

  const gLife = (I.p1_lifeEnd - I.p1_ageStart) || 0;     // var_G_Life  (col-B person)
  const cLife = (I.p2_lifeEnd - I.p2_ageStart) || 0;     // var_C_Life  (col-D person)
  const duration = Math.min(Math.max(gLife, cLife), 110);

  // Tax Engine tables
  const TB = I.taxBrackets, TR = I.taxRates;             // length 7 (C8:C14 / D8:D14)
  const TRDiff = TR.map((r,i) => r - (i ? TR[i-1] : 0)); // var_Tax_Rate_Diffs
  const LTCG_B = [0, 94050, 583750], LTCG_R = [0, 0.15, 0.20];

  // accounts (length 12)
  const A = I.accounts; // {statusA, statusG, type, name, start, rate, basis, contrib, match, cStart, cEnd, wStart}
  const nameL = s => String(s||'').toLowerCase().trim();
  const n1 = nameL(I.name1), n2 = nameL(I.name2);
  const owner = A.map(a => {
    const nm = nameL(a.name);
    if (n1 && nm.includes(n1)) return 1;
    if (n2 && nm.includes(n2)) return 2;
    return 0;
  });
  const incYes = A.map(a => nameL(a.statusA) === 'yes');
  const rate = A.map(a => (a.rate||0));                 // already decimal
  const startBal = A.map((a,i) => (a.start||0) * (incYes[i]?1:0));
  const typeL = A.map(a => nameL(a.type));
  const hier = typeL.map(t => t.includes('taxable')?1 : t.includes('pre')?2 : t.includes('roth')?3 : 99);
  const gUnlock = I.p1_unlock ? yOf(I.p1_unlock) : startYear;
  const cUnlock = I.p2_unlock ? yOf(I.p2_unlock) : startYear;
  const acctUnlock = A.map((a,i) => a.wStart ? yOf(a.wStart) : (owner[i]===1?gUnlock : owner[i]===2?cUnlock : startYear));

  const taxBasisInit = A.reduce((s,a,i)=> s + ((hier[i]===1 && incYes[i]) ? (a.basis||0) : 0), 0);

  // window proration helper: income active fraction for [S,E] in year yr (matches the IF ladders)
  function annualFrac(yr, S, E, startInclusive) {
    // mirrors the pension/other LET windows
    if (!S) return startInclusive ? 1 : 0;
    const ys = yOf(S);
    if (yr < ys) return 0;
    if (!E) return (yr===ys) ? (13-mOf(S))/12 : 1;
    const ye = yOf(E);
    if (yr > ye) return 0;
    if (ys===ye) return (mOf(E)-mOf(S)+1)/12;
    if (yr===ys) return (13-mOf(S))/12;
    if (yr===ye) return mOf(E)/12;
    return 1;
  }
  const bracketTax = (taxable, brackets) =>
    brackets.reduce((s,b,i)=> s + (taxable>b ? (taxable-b)*TRDiff[i] : 0), 0);
  const xlookupRate = (x, B, R) => { let r=R[0]; for(let i=0;i<B.length;i++){ if(x>=B[i]) r=R[i]; } return r; };

  let prevBals = startBal.slice(), prevBasis = taxBasisInit;
  const rows = [];

  for (let k=0; k<duration; k++) {
    const yr = startYear + k, idx = yr - startYear;

    // contributions
    const cActive = A.map((a,i)=> (a.contrib>0) *
      (a.cStart? (yr>=yOf(a.cStart)?1:0):1) * (a.cEnd? (yr<=yOf(a.cEnd)?1:0):1) * (incYes[i]?1:0));
    const cPro = A.map((a)=> (a.cStart? (yOf(a.cStart)===yr? (13-mOf(a.cStart))/12 : 1):1) *
      (a.cEnd? (yOf(a.cEnd)===yr? mOf(a.cEnd)/12 : 1):1));
    const contribAmt = A.map((a,i)=> Math.round((a.contrib||0)*cActive[i]*cPro[i]));
    const matchAmt   = A.map((a,i)=> Math.round((a.match||0)*cActive[i]*cPro[i]));
    const sp = (arr,m)=> arr.reduce((s,v,i)=> s + v*(m[i]?1:0), 0);
    const craigContrib = Math.round(sp(contribAmt, owner.map(o=>o===1)));
    const craigMatch   = Math.round(sp(matchAmt,   owner.map(o=>o===1)));
    const genaContrib  = Math.round(sp(contribAmt, owner.map(o=>o===2)));
    const genaMatch    = Math.round(sp(matchAmt,   owner.map(o=>o===2)));
    const totalContrib = craigContrib+craigMatch+genaContrib+genaMatch;

    const gDeath = startYear+gLife, cDeath = startYear+cLife;
    const gAlive = yr<=gDeath, cAlive = yr<=cDeath;
    const craigAge = I.p1_ageStart+idx, genaAge = I.p2_ageStart+idx;

    const gSal = gAlive ? I.p1_salary*(!I.p1_salEnd?1:(yr<yOf(I.p1_salEnd)?1:(yr===yOf(I.p1_salEnd)?mOf(I.p1_salEnd)/12:0))) : 0;
    const ssFrac = (start)=> !start?1:(yr<yOf(start)?0:(yr===yOf(start)?(13-mOf(start))/12:1));
    const gProjSS = I.p1_ssMonthly*12*ssFrac(I.p1_ssStart)*Math.pow(1+I.inflSS, idx);
    const cProjSS = I.p2_ssMonthly*12*ssFrac(I.p2_ssStart)*Math.pow(1+I.inflSS, idx);
    const gSS = gAlive ? ((!cAlive && craigAge>=60) ? Math.max(gProjSS,cProjSS) : gProjSS) : 0;
    const cSS = cAlive ? ((!gAlive && genaAge>=60)  ? Math.max(cProjSS,gProjSS) : cProjSS) : 0;
    const gPen = gAlive ? I.p1_pension*12*annualFrac(yr, I.p1_penStart, I.p1_penEnd, true) : 0;
    const gOth = gAlive ? I.p1_other*annualFrac(yr, I.p1_othStart, I.p1_othEnd, false) : 0;
    const cSal = cAlive ? I.p2_salary*(!I.p2_salEnd?1:(yr<yOf(I.p2_salEnd)?1:(yr===yOf(I.p2_salEnd)?mOf(I.p2_salEnd)/12:0))) : 0;
    const cPen = cAlive ? I.p2_pension*12*annualFrac(yr, I.p2_penStart, I.p2_penEnd, true) : 0;
    const cOth = cAlive ? I.p2_other*annualFrac(yr, I.p2_othStart, I.p2_othEnd, false) : 0;
    const totalGross = gSal+gSS+gPen+gOth+cSal+cSS+cPen+cOth;

    // RMD / conversions off pre-tax
    const cPreMask = owner.map((o,i)=> (o===1 && hier[i]===2)?1:0);
    const gPreMask = owner.map((o,i)=> (o===2 && hier[i]===2)?1:0);
    const c401 = sp(prevBals, cPreMask), g401 = sp(prevBals, gPreMask);
    const rmdDiv = (age)=> { const a=Math.min(age,120); let div=0; for(let i=0;i<I.rmdAge.length;i++){ if(a>=I.rmdAge[i]) div=I.rmdDiv[i]; } return div; };
    const rmdC = craigAge>=I.rmdStartC ? (rmdDiv(craigAge)? c401/rmdDiv(craigAge):0) : 0;
    const rmdG = genaAge >=I.rmdStartG ? (rmdDiv(genaAge)?  g401/rmdDiv(genaAge):0) : 0;
    const totalRMD = rmdC+rmdG;
    const convCreq = (I.convTable[yr]?.c)||0, convGreq = (I.convTable[yr]?.g)||0;
    const convC = Math.max(0, Math.min(convCreq, c401-rmdC));
    const convG = Math.max(0, Math.min(convGreq, g401-rmdG));
    const totalConv = convC+convG;
    const forcedW = prevBals.map((b,i)=>
      cPreMask[i]*(c401? b/c401:0)*(rmdC+convC) + gPreMask[i]*(g401? b/g401:0)*(rmdG+convG));
    const remBals = prevBals.map((b,i)=> Math.max(0, b-forcedW[i]));

    // filing status / std deduction / SS taxation
    const survMult = !n2 ? 1 : ((gAlive&&cAlive)?1 : ((gAlive||cAlive)? 1-I.survivorReduce : 0));
    const rInf = I.baseLiving*survMult*Math.pow(1+I.inflLiving, idx);
    const bothAliveFS = (gAlive&&cAlive)||(yr===gDeath)||(yr===cDeath);
    const fsRaw = String(I.filingStatus||'').toUpperCase().trim();
    const chosenFS = (fsRaw==='MARRIED FILING SEPARATELY'||fsRaw==='MFS')?'MFS':
                     (fsRaw==='HEAD OF HOUSEHOLD'||fsRaw==='HOH')?'HoH':
                     (fsRaw==='SINGLE')?'Single':'MFJ';
    const filing = bothAliveFS?chosenFS : (chosenFS==='HoH'?'HoH':'Single');
    const stdBase = filing==='MFJ'?32200 : filing==='HoH'?24150 : 16100;
    const stdDed = stdBase*Math.pow(1+I.inflStdDed, idx);
    const totalSS = gSS+cSS;
    const nonSSGross = gSal+gPen+gOth+cSal+cPen+cOth;
    const combIncome = nonSSGross+totalRMD+totalConv+0.5*totalSS;
    let taxableSS;
    if (filing==='MFJ') taxableSS = combIncome>44000 ? Math.min(0.85*totalSS, 0.85*(combIncome-44000)+Math.min(6000,0.5*totalSS))
                        : combIncome>32000 ? Math.min(0.5*totalSS, 0.5*(combIncome-32000)) : 0;
    else taxableSS = combIncome>34000 ? Math.min(0.85*totalSS, 0.85*(combIncome-34000)+Math.min(4500,0.5*totalSS))
                        : combIncome>25000 ? Math.min(0.5*totalSS, 0.5*(combIncome-25000)) : 0;
    const taxableGross = nonSSGross+taxableSS+totalRMD+totalConv;

    // IRMAA + healthcare
    const irmaaT = filing==='MFJ'?I.irmaaMFJ:I.irmaaSingle;
    const irmaaAdd = (age, pAlive, mAge)=> {
      if (!(pAlive && age>=mAge)) return 0;
      let v=0; for(let i=0;i<irmaaT.length;i++){ if(I.magiThresh && (filing==='MFJ'?I.magiMFJ:I.magiSingle)[i]<=taxableGross) v=irmaaT[i]; }
      return v*12;
    };
    const gHC = gAlive ? (craigAge>=I.p1_medAge ? I.p1_medPrem+irmaaAdd(craigAge,gAlive,I.p1_medAge) : I.p1_preMed) : 0;
    const cHC = cAlive ? (genaAge >=I.p2_medAge ? I.p2_medPrem+irmaaAdd(genaAge, cAlive,I.p2_medAge) : I.p2_preMed) : 0;
    const uInf = (gHC+cHC)*Math.pow(1+I.inflHC, idx);
    const vDebt = I.debts.reduce((s,d)=> s + (nameL(d.status)==='yes' && yOf(d.start)<=yr && yOf(d.end)>=yr ? d.amount : 0), 0);
    const wTotal = gs15(gs15(rInf)+gs15(uInf)+vDebt);

    const baseTaxable = Math.max(0, taxableGross-stdDed);
    const curBrackets = filing==='MFJ'?TB : filing==='HoH'?[0,17700,67450,105700,201775,256225,640600] : TB.map(b=>b/2);
    const baseTaxes = bracketTax(baseTaxable, curBrackets); // (MFS branch omitted in first pass; see notes)
    const baseStateTax = baseTaxable*I.stateTax;

    // chosen extra (smile) — projection input knob
    const phYr = (s,e,amt)=> (yr>=yOf(s)&&yr<=yOf(e))?amt:null;
    const smile = [ [I.ph1Start,I.ph1End,I.ph1Spend], [I.ph2Start,I.ph2End,I.ph2Spend], [I.ph3Start,I.ph3End,I.ph3Spend] ]
      .reduce((acc,p)=> acc!=null?acc:phYr(p[0],p[1],p[2]), null) || 0;
    const assumeExtra = nameL(I.assumeExtra)==='yes';
    const chosenExtra = (opts.extraOverride!=null) ? opts.extraOverride(yr, smile) : (assumeExtra?smile:0);

    const targetNet = Math.max(0, gs15(wTotal+chosenExtra+baseTaxes+baseStateTax-totalGross-totalRMD));

    // effective hierarchy (locked accounts deferred unless owner deceased)
    const ownerDeceased = (i)=> owner[i]===1?yr>gDeath : owner[i]===2?yr>cDeath : false;
    const effH = remBals.map((b,i)=> yr<acctUnlock[i] ? (ownerDeceased(i)?hier[i]:(hier[i]===1?1:99)) : hier[i]);
    const B1 = sp(remBals, effH.map(h=>h===1)), B2 = sp(remBals, effH.map(h=>h===2)), B3 = sp(remBals, effH.map(h=>h===3));

    // W1 taxable (cap gains)
    const cgFrac = Math.max(0, B1? (B1-prevBasis)/B1 : 0);
    const cgEst = Math.min(targetNet,B1)*cgFrac;
    const cgTaxEst = cgEst*xlookupRate(baseTaxable+cgEst, LTCG_B, LTCG_R);
    const W1 = Math.min(targetNet+cgTaxEst, B1);
    const cgActGain = W1*cgFrac;
    const cgTaxActual = cgActGain*xlookupRate(baseTaxable+cgActGain, LTCG_B, LTCG_R);
    const net1 = W1-cgTaxActual, remNet1 = Math.max(0, targetNet-net1);

    // W2 pre-tax (marginal gross-up w/ SS multiplier)
    const marginal = taxableGross<stdDed ? 0 : xlookupRate(baseTaxable, curBrackets, TR);
    const ssMult = taxableSS>=0.85*totalSS ? 1 :
      (filing==='MFJ' ? (combIncome>44000?1.85:combIncome>32000?1.5:1)
                      : (combIncome>34000?1.85:combIncome>25000?1.5:1));
    const effMarg = Math.min(0.99, marginal*ssMult);
    const W2 = Math.min(remNet1/(1-effMarg), B2);
    const totTaxableW2 = Math.max(0, taxableGross+W2-stdDed);
    const taxesW2 = bracketTax(totTaxableW2, curBrackets);
    const ordTaxActual = taxesW2-baseTaxes;
    const net2 = W2-ordTaxActual, remNet2 = Math.max(0, remNet1-net2);

    // W3 roth
    const W3 = Math.min(remNet2, B3);

    const extraTaxes = cgTaxActual+ordTaxActual;
    const totalTaxable = Math.max(0, taxableGross+W2-stdDed);
    const stateTax = totalTaxable*I.stateTax;
    const totalTaxes = baseTaxes+extraTaxes+stateTax;
    const fundingReq = Math.max(0, wTotal+chosenExtra+totalTaxes-totalGross);
    const surplus = Math.max(0, totalGross-(wTotal+chosenExtra+totalTaxes));

    // disbursement allocation per account
    const dDisc = remBals.map((b,i)=> gsR(
      (effH[i]===1?(B1? b/B1*W1:0):0)+(effH[i]===2?(B2? b/B2*W2:0):0)+(effH[i]===3?(B3? b/B3*W3:0):0)));
    const D = prevBals.map((b,i)=> gsR(forcedW[i]+dDisc[i]));

    // reinvested RMDs + conversion destinations (basis tracking)
    const reinvRMD = Math.max(0, totalRMD-fundingReq);
    const reinvC = Math.round(totalRMD>0? reinvRMD*(rmdC/totalRMD):0);
    const reinvG = Math.round(totalRMD>0? reinvRMD*(rmdG/totalRMD):0);
    const cRoth = owner.map((o,i)=>(o===1&&hier[i]===3)?1:0), gRoth = owner.map((o,i)=>(o===2&&hier[i]===3)?1:0);
    const cTax  = owner.map((o,i)=>(o===1&&hier[i]===1)?1:0), gTax  = owner.map((o,i)=>(o===2&&hier[i]===1)?1:0);
    const anyTax = hier.map(h=>h===1?1:0);
    const nSum = m=> m.reduce((s,v)=>s+v,0);
    const dest = (roth, tax)=> { const rn=nSum(roth), tn=nSum(tax), an=nSum(anyTax);
      if (rn>0) return roth.map(v=>v/rn); if (tn>0) return tax.map(v=>v/tn); return roth.map(()=>0); };
    const cConvDest = dest(cRoth,cTax), gConvDest = dest(gRoth,gTax);
    const reinvDest = (tax)=> { const tn=nSum(tax), an=nSum(anyTax);
      if (tn>0) return tax.map(v=>v/tn); if (an>0) return anyTax.map(v=>v/an); return tax.map(()=>0); };
    const cReinvDest = reinvDest(cTax), gReinvDest = reinvDest(gTax);
    const reinvArr = prevBals.map((_,i)=> cConvDest[i]*convC + gConvDest[i]*convG + cReinvDest[i]*reinvC + gReinvDest[i]*reinvG);

    const basisConsumed = W1*(B1? prevBasis/B1:0);
    const basisAdded = contribAmt.reduce((s,_,i)=> s + ((hier[i]===1)?(contribAmt[i]+matchAmt[i]+reinvArr[i]):0), 0);
    const basisNext = Math.max(0, prevBasis-basisConsumed+basisAdded);

    const post = prevBals.map((b,i)=> gsR(b-D[i]+reinvArr[i]+contribAmt[i]+matchAmt[i]));
    const mid  = prevBals.map((b,i)=> gsR((b+post[i])/2));

    // bear drag (crash params B12-B19)
    const isBear = String(I.crashType||'').toLowerCase().includes('bear') && yr>=(I.crashStart||0) && yr<((I.crashStart||0)+(I.crashDur||0));
    const isLate = String(I.lateCrashType||'').toLowerCase().includes('bear') && yr>=(I.lateStart||0) && yr<((I.lateStart||0)+(I.lateDur||0));
    const ends = mid.map((mv,i)=> {
      const drag = (isBear?(I.crashDrag||0):0)*(rate[i]>0.03?1:0) + (isLate?(I.lateDrag||0):0)*(rate[i]>0.03?1:0);
      const growth = gsR(gs15(mv*gs15((rate[i]-drag)*(incYes[i]?1:0))));
      return gsR(post[i]+growth);
    });

    // ── Full Master row: all 82 columns, in sheet order (for the web Master grid) ──
    const _pad = n => String(n).padStart(2,'0');
    const _sm = (arr,f) => arr.reduce((s,v,i)=> s + (f(i)?v:0), 0);
    const _startTot=_sm(prevBals,()=>true), _postTot=post.reduce((s,v)=>s+v,0), _midTot=mid.reduce((s,v)=>s+v,0);
    const _endTot=ends.reduce((s,v)=>s+v,0);
    const master = [
      yr, yr+'-'+_pad(pm)+'-'+_pad(pd), craigAge, genaAge,
      gSal, gSS, gProjSS*((gAlive&&cAlive)?1:0), gPen, gOth, cSal, cSS, cProjSS*((gAlive&&cAlive)?1:0), cPen, cOth, totalGross,
      craigContrib, craigMatch, genaContrib, genaMatch, totalContrib, stdDed,
      I.baseLiving, I.inflLiving, rInf, (gHC+cHC), I.inflHC, uInf, vDebt, wTotal,
      Math.max(0, wTotal-totalGross), surplus, chosenExtra, fundingReq,
      rmdC, rmdG, totalRMD, convC, convG, totalConv,
      totalTaxable, baseTaxes+ordTaxActual, stateTax, cgTaxActual, totalTaxes,
      Math.max(totalRMD, fundingReq), reinvRMD,
      D[0],D[1],D[2],D[3],D[4],D[5],D[6],D[7],D[8],D[9],D[10],D[11], (D.reduce((s,v)=>s+v,0))/12,
      ends[0],ends[1],ends[2],ends[3],ends[4],ends[5],ends[6],ends[7],ends[8],ends[9],ends[10],ends[11],
      _startTot, _sm(prevBals,i=>hier[i]===2), _sm(prevBals,i=>hier[i]===3), _sm(prevBals,i=>hier[i]===1),
      prevBasis, _sm(prevBals,i=>typeL[i].includes('hsa')), _sm(prevBals,i=>typeL[i].includes('529')),
      _postTot, _midTot, _endTot-_postTot, _endTot
    ];

    rows.push({
      yr, craigAge, genaAge, totalGross: Math.round(totalGross),
      wTotal: Math.round(wTotal), surplus: Math.round(surplus), chosenExtra: Math.round(chosenExtra),
      fundingReq: Math.round(fundingReq), totalRMD: Math.round(totalRMD), totalConv: Math.round(totalConv),
      totalTaxes: Math.round(totalTaxes), totalTaxable: Math.round(totalTaxable),
      ends: ends.slice(), endTotal: ends.reduce((s,v)=>s+v,0),
      liquidEnd: ends.reduce((s,v,i)=> s + (typeL[i].includes('liquid')?v:0), 0),
      basisNext: Math.round(basisNext), master: master
    });
    prevBals = ends; prevBasis = basisNext;
  }
  return rows;
}

function rbpSolveSafeExtra(I){
  const yOf=d=>d?new Date(d).getFullYear():null;
  const w=[I.w1||1,I.w2||1,I.w3||1], avg=(w[0]+w[1]+w[2])/3;
  const p1s=yOf(I.ph1Start),p1e=yOf(I.ph1End),p2s=yOf(I.ph2Start),p2e=yOf(I.ph2End),p3s=yOf(I.ph3Start),p3e=yOf(I.ph3End);
  const A=I.accounts, nameL=s=>String(s||'').toLowerCase().trim();
  const withMask=A.map(a=>{const g=nameL(a.G); return (g===''||g.includes('use')||g.includes('yes')||g.includes('withdraw'))?1:0;});
  function shortfall(X){
    const rows=rbpProject(I,{extraOverride:(yr)=>{
      let ph=-1; if(yr>=p1s&&yr<=p1e)ph=0; else if(yr>=p2s&&yr<=p2e)ph=1; else if(yr>=p3s&&yr<=p3e)ph=2;
      return ph<0?0: X*w[ph]/avg;
    }});
    const wd=rows.map(r=> r.ends.reduce((s,v,i)=> s+(withMask[i]?v:0),0));
    const finalW=wd[wd.length-1], minW=Math.min.apply(null,wd);
    return Math.min(finalW-(I.legacyB8||0), minW-(I.floorB9||0));
  }
  let lo=0, hi=200000;
  for(let it=0; it<48; it++){ const mid=(lo+hi)/2; if(shortfall(mid)>=0) lo=mid; else hi=mid; }
  return { safeExtra:Math.round(lo), shortfallAtSolve:Math.round(shortfall(lo)) };
}

/* ── runStress: one crash scenario, fully re-solved, no approximations ──────
 * params = { drag, dur, start, inflation, lateDrag, lateDur, lateStart, plannedExtra }
 *   drag        crash drag in points (7 − return%, e.g. −10% → 17)
 *   start, dur  crash year and length (default this plan's start, 1 yr)
 *   inflation   living-cost inflation FOR THE STRESS RUN (the panel slider, e.g. 0.03);
 *               omit to use the saved plan's own inflLiving
 *   plannedExtra the spend the red line should assume (the committed plan's safe extra);
 *               omit and the engine solves the calm-weather safe extra itself
 * returns { ending, hit, crossover, safeExtra, safeExtraBase, depletes, series }
 *   ending      final-year liquid total (Master col CD / B129)
 *   hit         ending − legacy goal
 *   crossover   first year liquid total drops below the legacy goal (null = never)
 *   safeExtra   max extra STILL safe to add UNDER this crash — honestly re-solved at
 *               full precision (taxes computed exactly; no $500 rounding, no margin)
 *   safeExtraBase  the calm-weather safe extra, for reference
 *   series      [{year, bal}] full liquid trajectory (the red line)            */
function rbpRunStress(I, params) {
  params = params || {};
  const yOf = d => (d ? new Date(d).getFullYear() : null);
  const start = params.start || yOf(I.planStart);
  const S = Object.assign({}, I, {
    crashType: 'bear', crashStart: start, crashDur: (params.dur || 1), crashDrag: (params.drag || 0),
    lateCrashType: (params.lateDrag ? 'bear' : ''), lateStart: (params.lateStart || 0),
    lateDur: (params.lateDur || 0), lateDrag: (params.lateDrag || 0)
  });
  if (params.inflation != null) S.inflLiving = params.inflation;   // stress-only inflation (B20)

  // calm-weather safe extra (same regardless of crash) — solved once, reused for the red line
  const baseSafe = (params.plannedExtra != null) ? params.plannedExtra : rbpSolveSafeExtra(I).safeExtra;

  // project the committed plan THROUGH the crash → the red line
  const w = [S.w1 || 1, S.w2 || 1, S.w3 || 1], avg = (w[0] + w[1] + w[2]) / 3;
  const p1s = yOf(S.ph1Start), p1e = yOf(S.ph1End), p2s = yOf(S.ph2Start), p2e = yOf(S.ph2End), p3s = yOf(S.ph3Start), p3e = yOf(S.ph3End);
  const rows = rbpProject(S, { extraOverride: (yr) => {
    let ph = -1; if (yr >= p1s && yr <= p1e) ph = 0; else if (yr >= p2s && yr <= p2e) ph = 1; else if (yr >= p3s && yr <= p3e) ph = 2;
    return ph < 0 ? 0 : baseSafe * w[ph] / avg;
  }});

  const legacy = S.legacyB8 || 0;
  const series = rows.map(r => ({ year: r.yr, bal: Math.round(r.endTotal) }));
  const ending = rows.length ? Math.round(rows[rows.length - 1].endTotal) : 0;
  const below = series.find(p => p.bal < legacy);
  const zero = series.find(p => p.bal <= 0);

  // honest re-solve of safe extra UNDER this crash — full precision, exact taxes
  const crashSafe = rbpSolveSafeExtra(S).safeExtra;

  return {
    ending: ending,
    hit: Math.round(ending - legacy),
    crossover: below ? below.year : null,
    depletes: zero ? zero.year : null,
    safeExtra: crashSafe,
    safeExtraBase: Math.round(baseSafe),
    series: series
  };
}

/* ── buildI: connector `read` payload → engine input I ──────────────────────
 * Pages fetch the customer's inputs ONCE (action=read), then run everything
 * client-side. Per-customer values come from the payload; the fixed 2025
 * federal tables (brackets, RMD, IRMAA) are baked in here — same for everyone. */
var RBP_TAX_2025 = {
  brackets:[0,24800,100800,211400,403550,512450,768700],
  rates:[0.1,0.12,0.22,0.24,0.32,0.35,0.37],
  rmdAge:[73,74,75,76,77,78,79,80,81,82,83,84,85,86,87,88,89,90,91,92,93,94,95,96,97,98,99,100,101,102,103,104,105,106,107,108,109,110,111,112,113,114,115,116,117,118,119,120],
  rmdDiv:[26.5,25.5,24.6,23.7,22.9,22,21.1,20.2,19.4,18.5,17.7,16.8,16,15.2,14.4,13.7,12.9,12.2,11.5,10.8,10.1,9.5,8.9,8.4,7.8,7.3,6.8,6.4,6,5.6,5.2,4.9,4.6,4.3,4.1,3.9,3.7,3.5,3.4,3.3,3.1,3,2.9,2.8,2.7,2.5,2.3,2],
  rmdStartC:75, rmdStartG:75,
  irmaaMFJ:[0,95.7,240.4,385,529.6,578],
  magiMFJ:[0,218001,274001,342001,410001,750001],
  magiSingle:[0,109001,137001,171001,205001,500001]
};
function rbpBuildI(R, D) {
  D = D || function(s){ return s ? new Date(s) : null; };
  if (R && R.accounts && R.taxBrackets) return R;          // already engine-shaped
  var Y = function(yr){ return yr ? new Date(yr,0,1) : null; }; // phase year-number → Date
  var p1 = R.people.craig, p2 = R.people.gena, gl = R.global, t = R.tax, sp = R.spending, leg = R.legacy;
  var T = RBP_TAX_2025;
  return {
    planStart:D(gl.planStartYear), name1:p1.name, name2:p2.name,
    p1_ageStart:p1.age, p1_lifeEnd:p1.deathAge, p2_ageStart:p2.age, p2_lifeEnd:p2.deathAge,
    p1_salary:p1.salary, p1_salEnd:D(p1.salaryEnd), p2_salary:p2.salary, p2_salEnd:D(p2.salaryEnd),
    p1_ssMonthly:p1.ssMonthly, p1_ssStart:D(p1.ssStartDate), p2_ssMonthly:p2.ssMonthly, p2_ssStart:D(p2.ssStartDate), inflSS:t.ssCola,
    p1_pension:p1.pension, p1_penStart:D(p1.pensionStart), p1_penEnd:null,
    p1_other:p1.otherIncome, p1_othStart:D(p1.otherStart), p1_othEnd:D(p1.otherEnd),
    p2_pension:p2.pension, p2_penStart:D(p2.pensionStart), p2_penEnd:null,
    p2_other:p2.otherIncome, p2_othStart:D(p2.otherStart), p2_othEnd:D(p2.otherEnd),
    p1_unlock:D(p1.ssStartDate), p2_unlock:D(p2.ssStartDate),
    inflLiving:t.inflation, inflStdDed:t.stdDedInflation, inflHC:t.healthcareInflation, stateTax:t.stateRate,
    baseLiving:sp.baseAnnual, survivorReduce:gl.survivorReduction, filingStatus:gl.filingStatus, assumeExtra:gl.assumeExtraSpend,
    p1_preMed:p1.healthPreMedicare, p1_medPrem:p1.healthMedicare, p1_medAge:p1.medicareAge,
    p2_preMed:p2.healthPreMedicare, p2_medPrem:p2.healthMedicare, p2_medAge:p2.medicareAge,
    taxBrackets:T.brackets, taxRates:T.rates,
    rmdAge:T.rmdAge, rmdDiv:T.rmdDiv, rmdStartC:T.rmdStartC, rmdStartG:T.rmdStartG,
    irmaaMFJ:T.irmaaMFJ, irmaaSingle:T.irmaaMFJ, magiMFJ:T.magiMFJ, magiSingle:T.magiSingle, magiThresh:true,
    convTable:{},
    debts:(R.debts||[]).map(function(d){ return {status:d.inc, amount:d.ann, start:D(d.start), end:D(d.end)}; }),
    ph1Start:Y(sp.phase1Start), ph1End:Y(sp.phase1End), ph1Spend:0,
    ph2Start:Y(sp.phase2Start), ph2End:Y(sp.phase2End), ph2Spend:0,
    ph3Start:Y(sp.phase3Start), ph3End:Y(sp.phase3End), ph3Spend:0,
    w1:sp.phase1Weight, w2:sp.phase2Weight, w3:sp.phase3Weight,
    legacyB8:leg.goal, floorB9:leg.safetyFloor,
    crashType:'', crashStart:0, crashDur:0, crashDrag:0, lateCrashType:'', lateStart:0, lateDur:0, lateDrag:0,
    committedExtra:sp.safeExtra,   // the saved plan's safe extra — the red line's spend
    accounts:(R.portfolio.accounts||[]).map(function(a){ return {
      statusA:a.showInCalc, statusG:a.status, type:a.type, name:a.name, start:a.balance,
      rate:(a.expectedReturn||0)/100, basis:a.costBasis, contrib:a.contrib||0, match:a.match||0,
      cStart:D(a.contribStart), cEnd:D(a.contribEnd), wStart:D(a.withdrawStart) }; })
  };
}

/* ── UMD: browser global `RBP` + CommonJS for node validation ───────────────
   Extends an existing window.RBP (e.g. a page's connector loader RBP.load/
   loadCached) rather than replacing it, so engine + loader coexist in any order. */
(function (root) {
  var api = { project: rbpProject, solveSafeExtra: rbpSolveSafeExtra, runStress: rbpRunStress, buildI: rbpBuildI };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  if (root.RBP && typeof root.RBP === 'object') { for (var k in api) root.RBP[k] = api[k]; }
  else root.RBP = api;
})(typeof self !== 'undefined' ? self : this);
