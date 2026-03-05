import React, { useState } from 'react';
import { Save, RefreshCw, FileSpreadsheet, Calculator, Wind, Droplets, Gauge, Scale, AlertTriangle, CheckCircle2, Plus, Trash2, ListOrdered, Target, ChevronDown, Download, Zap } from 'lucide-react';

const NOZZLE_SET = [
  { num: 4, d: 3.21 }, { num: 5, d: 3.97 }, { num: 6, d: 4.79 },
  { num: 7, d: 5.54 }, { num: 8, d: 6.28 }, { num: 9, d: 7.10 },
  { num: 10, d: 7.86 }, { num: 11, d: 8.40 }, { num: 12, d: 9.24 },
  { num: 13, d: 10.18 }, { num: 14, d: 10.87 }, { num: 15, d: 11.70 },
  { num: 16, d: 12.56 }, { num: 17, d: 13.48 }, { num: 18, d: 14.30 },
  { num: 20, d: 15.95 }, { num: 22, d: 17.51 }, { num: 26, d: 20.72 }
];

const SAMPLERS = [
  { id: 1, name: '샘플러 1호기', yd: '0.991', dHAt: '47.6' },
  { id: 2, name: '샘플러 2호기', yd: '1.005', dHAt: '40.2' },
  { id: 3, name: '샘플러 3호기', yd: '0.988', dHAt: '39.5' },
  { id: 4, name: '샘플러 4호기', yd: '1.012', dHAt: '42.1' },
  { id: 5, name: '샘플러 5호기', yd: '0.998', dHAt: '41.5' }
];

export default function App() {
  const [formData, setFormData] = useState({
    date: '', company: '', location: '', sampler: '',
    weather: '맑음', atmTemp: '20', atmPressure: '760',
    totalStackDepth: '', flangeLength: '', stackDiameter: '', pitotFactor: '0.84',
    standardO2: '',
    gasAnalyzer: [
      { o2: '21.00', co2: '0.00', co: '0.00', sox: '0.00', nox: '0.00' },
      { o2: '21.00', co2: '0.00', co: '0.00', sox: '0.00', nox: '0.00' },
      { o2: '21.00', co2: '0.00', co: '0.00', sox: '0.00', nox: '0.00' }
    ],
    moistureValues: ['', '', '', '', ''],
    impingers: [
      { id: 1, initial: '', final: '' }, { id: 2, initial: '', final: '' },
      { id: 3, initial: '', final: '' }, { id: 4, initial: '', final: '' },
    ],
    points: [
      { id: 1, tp: '', sp: '', dp: '', ts: '' }, { id: 2, tp: '', sp: '', dp: '', ts: '' },
      { id: 3, tp: '', sp: '', dp: '', ts: '' }, { id: 4, tp: '', sp: '', dp: '', ts: '' },
      { id: 5, tp: '', sp: '', dp: '', ts: '' },
    ],
    samplerId: '', targetVolume: '1000', planSamplingTime: '65',
    gasMeterFactor: '0.991', deltaHAt: '47.6', expectedTmIn: '', expectedTmOut: '',
    recommendedNozzleNum: '', recommendedNozzleDia: '', usedNozzleNum: '',
    nozzleDiameter: '', kFactor: '',
    gasMeters: Array.from({ length: 6 }, (_, i) => ({
      id: i, pointNum: i === 0 ? '시작' : '', time: i === 0 ? '0' : '', stackTemp: '', 
      dp: '', pressure: '', volume: '', tmIn: '', tmOut: '', vacuum: '', impingerTemp: '' 
    })),
    filterId: '', filterInitial: '', filterFinal: '', remarks: ''
  });

  const [savedData, setSavedData] = useState([]);
  const [recommendations, setRecommendations] = useState(null);

  const handleChange = (e) => {
    const { name, value } = e.target;
    setFormData(prev => {
      const newData = { ...prev, [name]: value };
      if (name === 'gasMeterFactor' || name === 'deltaHAt') newData.samplerId = '';
      if (name === 'totalStackDepth' || name === 'flangeLength') {
        const total = parseFloat(newData.totalStackDepth);
        const flange = parseFloat(newData.flangeLength) || 0; 
        if (!isNaN(total)) newData.stackDiameter = Math.max(0, total - flange).toFixed(3);
      }
      return newData;
    });
  };

  const handleSamplerChange = (e) => {
    const id = e.target.value;
    if (id === '') { setFormData(prev => ({ ...prev, samplerId: '' })); return; }
    const sampler = SAMPLERS.find(s => s.id === parseInt(id));
    if (sampler) setFormData(prev => ({ ...prev, samplerId: id, gasMeterFactor: sampler.yd, deltaHAt: sampler.dHAt }));
  };

  const handleGasAnalyzerChange = (index, field, value) => {
    setFormData(prev => {
      const newAnalyzer = prev.gasAnalyzer.map((item, i) => 
        i === index ? { ...item, [field]: value } : item
      );
      return { ...prev, gasAnalyzer: newAnalyzer };
    });
  };

  const handleMoistureChange = (index, value) => {
    setFormData(prev => {
      const newMoistures = [...prev.moistureValues];
      newMoistures[index] = value;
      return { ...prev, moistureValues: newMoistures };
    });
  };

  const handleImpingerChange = (index, field, value) => {
    setFormData(prev => {
      const newImpingers = prev.impingers.map((item, i) => 
        i === index ? { ...item, [field]: value } : item
      );
      return { ...prev, impingers: newImpingers };
    });
  };

  const handlePointChange = (index, field, value) => {
    setFormData(prev => {
      const newPoints = prev.points.map((item, i) => {
        if (i !== index) return item;
        const newItem = { ...item, [field]: value };
        if (field === 'tp' || field === 'sp') {
          const tp = parseFloat(newItem.tp);
          const sp = parseFloat(newItem.sp);
          if (!isNaN(tp) && !isNaN(sp)) newItem.dp = (tp - sp).toFixed(2);
        }
        return newItem;
      });
      return { ...prev, points: newPoints };
    });
  };

  const addPoint = () => setFormData(prev => ({ ...prev, points: [...prev.points, { id: prev.points.length + 1, tp: '', sp: '', dp: '', ts: '' }] }));
  const removePoint = (index) => {
    setFormData(prev => {
      const newPoints = [...prev.points];
      newPoints.splice(index, 1);
      return { ...prev, points: newPoints };
    });
  };

  const handleGasMeterChange = (index, field, value) => {
    setFormData(prev => {
      const newMeters = prev.gasMeters.map((item, i) => {
        if (i !== index) return item;
        const newItem = { ...item, [field]: value };
        if (field === 'dp') {
           const dpValue = parseFloat(value);
           const kVal = parseFloat(prev.kFactor);
           if (!isNaN(dpValue) && !isNaN(kVal)) newItem.pressure = (kVal * dpValue).toFixed(2);
           else if (value === '') newItem.pressure = '';
        }
        return newItem;
      });
      return { ...prev, gasMeters: newMeters };
    });
  };
  
  const handleKFactorChange = (e) => {
      const kValStr = e.target.value, kVal = parseFloat(kValStr);
      setFormData(prev => {
          const newMeters = prev.gasMeters.map((meter, idx) => {
              if (idx === 0) return meter;
              const dpValue = parseFloat(meter.dp);
              if (!isNaN(kVal) && !isNaN(dpValue)) return { ...meter, pressure: (kVal * dpValue).toFixed(2) };
              return meter;
          });
          return { ...prev, kFactor: kValStr, gasMeters: newMeters };
      });
  };

  const addGasMeter = () => setFormData(prev => ({ ...prev, gasMeters: [...prev.gasMeters, { id: prev.gasMeters.length, pointNum: '', time: '', stackTemp: '', dp: '', pressure: '', volume: '', tmIn: '', tmOut: '', vacuum: '', impingerTemp: '' }] }));
  const removeGasMeter = (index) => {
    if (index === 0) return;
    setFormData(prev => {
      const newMeters = [...prev.gasMeters];
      newMeters.splice(index, 1);
      newMeters.forEach((m, i) => m.id = i);
      return { ...prev, gasMeters: newMeters };
    });
  };

  const getSamplingPoints = () => {
    const D = parseFloat(formData.stackDiameter), flange = parseFloat(formData.flangeLength) || 0;
    if (isNaN(D) || D <= 0) return null;
    const R = D / 2, area = Math.PI * Math.pow(R, 2);
    let rnCoeffs = [], isCenterOnly = false;

    if (area <= 0.25) isCenterOnly = true;
    else if (D <= 1) rnCoeffs = [0.707]; 
    else if (D <= 2) rnCoeffs = [0.500, 0.866]; 
    else if (D <= 4) rnCoeffs = [0.408, 0.707, 0.913]; 
    else if (D <= 4.5) rnCoeffs = [0.354, 0.612, 0.791, 0.935];
    else rnCoeffs = [0.316, 0.548, 0.707, 0.837, 0.949]; 
    
    let nearDistances = [], farDistances = [];
    if (isCenterOnly) {
      nearDistances.push(R); farDistances.push(R);
    } else {
      for (let i = rnCoeffs.length - 1; i >= 0; i--) {
        nearDistances.push(R - (rnCoeffs[i] * R));
        farDistances.push(R + (rnCoeffs[i] * R));
      }
    }
    return { perRadius: isCenterOnly ? 1 : rnCoeffs.length, isCenterOnly, area: area.toFixed(3), nearInsertion: nearDistances.map(d => (d + flange).toFixed(3)), farInsertion: farDistances.map(d => (d + flange).toFixed(3)) };
  };

  const applySamplingPointsToTable = () => {
    const spData = getSamplingPoints();
    if (!spData) return;
    const pointCount = spData.perRadius; 
    setFormData(prev => {
      const newPoints = Array.from({ length: pointCount }, (_, i) => prev.points[i] || { id: i + 1, tp: '', sp: '', dp: '', ts: '' });
      const newMeters = Array.from({ length: pointCount + 1 }, (_, i) => {
        if (i === 0) return prev.gasMeters[0] || { id: 0, pointNum: '시작', time: '0', stackTemp: '', dp: '', pressure: '', volume: '', tmIn: '', tmOut: '', vacuum: '', impingerTemp: '' };
        return prev.gasMeters[i] || { id: i, pointNum: '', time: '', stackTemp: '', dp: '', pressure: '', volume: '', tmIn: '', tmOut: '', vacuum: '', impingerTemp: '' };
      });
      return { ...prev, points: newPoints, gasMeters: newMeters };
    });
    alert(`✅ 0분 행을 포함하여 측정 포인트 [총 ${pointCount}개]에 맞춰 하단 표가 생성되었습니다.`);
  };

  const getGasComposition = () => {
    const o2 = formData.gasAnalyzer.reduce((acc, curr) => acc + (parseFloat(curr.o2) || 0), 0) / 3;
    const co2 = formData.gasAnalyzer.reduce((acc, curr) => acc + (parseFloat(curr.co2) || 0), 0) / 3;
    const co = formData.gasAnalyzer.reduce((acc, curr) => acc + (parseFloat(curr.co) || 0), 0) / 3;
    const sox = formData.gasAnalyzer.reduce((acc, curr) => acc + (parseFloat(curr.sox) || 0), 0) / 3;
    const nox = formData.gasAnalyzer.reduce((acc, curr) => acc + (parseFloat(curr.nox) || 0), 0) / 3;
    
    const coPercent = co / 10000; 
    const n2 = 100 - co2 - o2 - coPercent; 
    
    const M_O2 = 32, M_CO2 = 44, M_CO = 28, M_N2 = 28;
    const sumMx = (M_O2 * o2) + (M_CO2 * co2) + (M_CO * coPercent) + (M_N2 * n2);

    const Md = sumMx / 100;
    const validVals = formData.moistureValues.map(v => parseFloat(v)).filter(v => !isNaN(v));
    const Xw = validVals.length > 0 ? (validVals.reduce((a, b) => a + b, 0) / validVals.length) : 0;
    const Ms = Md * ((100 - Xw) / 100) + 18 * (Xw / 100);

    const r0 = (1 / (22.4 * 100)) * ( sumMx * ((100 - Xw) / 100) + 18 * Xw );

    return { o2, co2, co, sox, nox, n2, Md, Ms, Xw, r0 };
  };

  const getRawAvgDp = () => {
    const validDps = formData.points.map(p => parseFloat(p.dp)).filter(v => !isNaN(v));
    return validDps.length === 0 ? 0 : validDps.reduce((a, b) => a + b, 0) / validDps.length;
  };

  const getRawAvgTs = () => {
    const validTs = formData.points.map(p => parseFloat(p.ts)).filter(v => !isNaN(v));
    return validTs.length === 0 ? 0 : validTs.reduce((a, b) => a + b, 0) / validTs.length;
  };

  const getRawStackPressure = () => {
    const Pa = parseFloat(formData.atmPressure);
    const validSps = formData.points.map(p => parseFloat(p.sp)).filter(v => !isNaN(v));
    const avgPs = validSps.length === 0 ? 0 : validSps.reduce((a, b) => a + b, 0) / validSps.length;
    return !isNaN(Pa) ? Pa + (avgPs / 13.6) : NaN;
  };

  const getRawGasVelocity = () => {
    const avgTs = getRawAvgTs(), C = parseFloat(formData.pitotFactor) || 0.84, P = getRawStackPressure(); 
    const { Ms } = getGasComposition();
    const validDps = formData.points.map(p => parseFloat(p.dp)).filter(v => !isNaN(v) && v >= 0);
    if (validDps.length === 0 || isNaN(avgTs) || isNaN(P) || isNaN(Ms)) return 0;
    
    const sqrtDpAvg = validDps.reduce((a, b) => a + Math.sqrt(b), 0) / validDps.length;
    const velocityConstant = Math.sqrt((2 * 9.81 * 760 * 22.4) / 273); 
    return velocityConstant * C * sqrtDpAvg * Math.sqrt((273 + avgTs) / (P * Ms));
  };

  const getRawGasMeterVolDiff = () => {
    const validVols = formData.gasMeters.map(g => g.volume).filter(v => v !== '');
    if (validVols.length > 1) {
      const initial = parseFloat(validVols[0]), final = parseFloat(validVols[validVols.length - 1]);
      if (!isNaN(initial) && !isNaN(final)) return final - initial;
    }
    return 0;
  };

  const getRawAvgTm = () => {
    const validTemps = formData.gasMeters.filter((_, i) => i !== 0).map(g => {
      const inT = parseFloat(g.tmIn), outT = parseFloat(g.tmOut);
      if (!isNaN(inT) && !isNaN(outT)) return (inT + outT) / 2;
      if (!isNaN(inT)) return inT;
      if (!isNaN(outT)) return outT;
      return NaN;
    }).filter(v => !isNaN(v));
    return validTemps.length === 0 ? NaN : validTemps.reduce((a, b) => a + b, 0) / validTemps.length;
  };

  const getRawAvgOrifice = () => {
    const validPressures = formData.gasMeters.filter((_, i) => i !== 0).map(g => parseFloat(g.pressure)).filter(v => !isNaN(v));
    return validPressures.length === 0 ? NaN : validPressures.reduce((a, b) => a + b, 0) / validPressures.length;
  };

  const getRawTotalMoistureWeight = () => {
    let total = 0;
    formData.impingers.forEach(imp => {
      const init = parseFloat(imp.initial), fin = parseFloat(imp.final);
      if (!isNaN(init) && !isNaN(fin)) total += (fin - init);
    });
    return total;
  };

  const getRawPostMoisture = () => {
    const Wm = getRawTotalMoistureWeight(), Vm = getRawGasMeterVolDiff(), Tm = getRawAvgTm();
    const Pa = parseFloat(formData.atmPressure), Pm = getRawAvgOrifice() || 0;
    const Y = parseFloat(formData.gasMeterFactor) || 1.0;
    if (Wm > 0 && Vm > 0 && !isNaN(Tm) && !isNaN(Pa)) {
      const Vm_std = Vm * Y * (273 / (273 + Tm)) * ((Pa + Pm / 13.6) / 760);
      const Vw_std = Wm * (22.4 / 18.015);
      return (Vw_std / (Vm_std + Vw_std)) * 100;
    }
    return 0;
  };

  const getRawVmStd = () => {
    const Vm = getRawGasMeterVolDiff(), Tm = getRawAvgTm();
    const Pa = parseFloat(formData.atmPressure), Pm = getRawAvgOrifice() || 0;
    const Y = parseFloat(formData.gasMeterFactor) || 1.0;
    if (Vm > 0 && !isNaN(Tm) && !isNaN(Pa)) {
      return Vm * Y * (273 / (273 + Tm)) * ((Pa + Pm / 13.6) / 760);
    }
    return 0;
  };

  const getRawDustWeightDiff = () => {
    const initial = parseFloat(formData.filterInitial), final = parseFloat(formData.filterFinal);
    if (!isNaN(initial) && !isNaN(final)) return final - initial;
    return 0;
  };

  const calcAvgDp = () => getRawAvgDp(); 
  const calcAvgTs = () => getRawAvgTs();
  const calcStackPressure = () => getRawStackPressure();
  const calcMoisture = () => getGasComposition().Xw.toFixed(2);
  const calcPostMoisture = () => getRawPostMoisture().toFixed(2);
  const calcGasVelocity = () => { const v = getRawGasVelocity(); return v === 0 ? '0.00' : v.toFixed(2); };
  const calcGasMeterVolDiff = () => { const v = getRawGasMeterVolDiff(); return v > 0 ? v.toFixed(2) : 0; };
  const calcAvgTm = () => { const v = getRawAvgTm(); return isNaN(v) ? '-' : v.toFixed(1); };
  const calcAvgOrifice = () => { const v = getRawAvgOrifice(); return isNaN(v) ? '-' : v.toFixed(1); };
  const calcTotalMoistureWeight = () => getRawTotalMoistureWeight();
  const getVmStd = () => { const v = getRawVmStd(); return v > 0 ? v.toFixed(3) : '0.000'; };
  const calcDustWeightDiff = () => { const v = getRawDustWeightDiff(); return v.toFixed(4); };

  const calcO2CorrectionFactor = () => {
    const o2_ref = parseFloat(formData.standardO2);
    const { o2: o2_act } = getGasComposition();
    if (isNaN(o2_ref) || isNaN(o2_act) || o2_act >= 21) return 1.0;
    return (21 - o2_ref) / (21 - o2_act);
  };

  const getSamplingMinutes = () => formData.gasMeters.reduce((sum, meter, idx) => idx === 0 ? sum : sum + (isNaN(parseFloat(meter.time)) ? 0 : parseFloat(meter.time)), 0);

  const calcGasFlowRates = () => {
    const Vs = getRawGasVelocity(); 
    const Ts = getRawAvgTs();
    const Ps = getRawStackPressure();
    const D = parseFloat(formData.stackDiameter);
    
    if (Vs === 0 || isNaN(Ts) || isNaN(Ps) || isNaN(D)) return { dry: '-', wet: '-' };

    const A = Math.PI * Math.pow(D / 2, 2);
    const { Xw } = getGasComposition(); 

    const Q_wet = Vs * A * (273 / (273 + Ts)) * (Ps / 760) * 3600;
    const Q_dry = Q_wet * (1 - Xw / 100);

    return { wet: Q_wet.toFixed(0), dry: Q_dry.toFixed(0) };
  };

  const generateRecommendation = (mode) => {
    const Vs = getRawGasVelocity(), dpAvg = getRawAvgDp();
    const targetVol = parseFloat(formData.targetVolume) || 1000;
    const tmIn = parseFloat(formData.expectedTmIn) || parseFloat(formData.atmTemp) || 25;
    const tmOut = parseFloat(formData.expectedTmOut) || parseFloat(formData.atmTemp) || 25;
    const Tm = (tmIn + tmOut) / 2; 

    const Ts = getRawAvgTs();
    const Pm = parseFloat(formData.atmPressure); 
    
    const validSps = formData.points.map(p => parseFloat(p.sp)).filter(v => !isNaN(v));
    const avgSp = validSps.length === 0 ? 0 : validSps.reduce((a, b) => a + b, 0) / validSps.length;
    const Ps_ratio = (Pm + avgSp / 13.6) / Pm; 

    const { Md, Ms, Xw } = getGasComposition();
    const Cp = parseFloat(formData.pitotFactor) || 0.84;
    const Yd = parseFloat(formData.gasMeterFactor) || 1.0;
    const dHAt = parseFloat(formData.deltaHAt);

    let bestNozzle = null, maxDH = 0, closestDiffStandard = Infinity, closestDiffStable = Infinity;

    for (let n of NOZZLE_SET) {
        const Dn = n.d;
        const tempK = 8.001e-5 * Math.pow(Dn, 4) * dHAt * Math.pow(Cp, 2) * Math.pow(1 - Xw / 100, 2) * (Md / Ms) * ((Tm + 273) / (Ts + 273)) * Ps_ratio;
        const expectedDHCheck = tempK * dpAvg;

        if (expectedDHCheck <= 50) {
            if (mode === 'fast') {
                if (expectedDHCheck > maxDH) { maxDH = expectedDHCheck; bestNozzle = n; }
            } else if (mode === 'standard') {
                const diff = Math.abs(expectedDHCheck - 25);
                if (diff < closestDiffStandard) { closestDiffStandard = diff; bestNozzle = n; }
            } else if (mode === 'stable') {
                const diff = Math.abs(expectedDHCheck - 10);
                if (diff < closestDiffStable) { closestDiffStable = diff; bestNozzle = n; }
            }
        }
    }

    if (!bestNozzle) bestNozzle = NOZZLE_SET[0];

    const Dn_best = bestNozzle.d;
    const An = Math.PI * Math.pow(Dn_best / 2000, 2);
    const Ps_abs = Pm + avgSp / 13.6;
    const bestQm = (Vs * An * 60 * 1000 * ((Tm + 273) / (Ts + 273)) * (Ps_abs / Pm) * (1 - Xw / 100)) / Yd;
    const Q_sl = bestQm * Yd * (273 / (Tm + 273)) * (Pm / 760);
    
    let fastestTime = Math.ceil(targetVol / Q_sl);
    if (fastestTime <= 0 || !isFinite(fastestTime)) fastestTime = 65;

    const K = 8.001e-5 * Math.pow(Dn_best, 4) * dHAt * Math.pow(Cp, 2) * Math.pow(1 - Xw / 100, 2) * (Md / Ms) * ((Tm + 273) / (Ts + 273)) * Ps_ratio;

    return { bestNozzleNum: bestNozzle.num, finalDn: Dn_best.toFixed(2), fastestTime, calculatedK: K.toFixed(3), expectedDH: (K * dpAvg).toFixed(2) };
  };

  const handleCalculateOptions = () => {
    const Vs = getRawGasVelocity(), Ts = getRawAvgTs(), Ps = getRawStackPressure(), dpAvg = getRawAvgDp();
    const Y = parseFloat(formData.gasMeterFactor), dHAt = parseFloat(formData.deltaHAt);

    if (isNaN(Y) || Y <= 0 || isNaN(dHAt) || dHAt <= 0) { alert('장비 교정 성적서에 명시된 [보정계수(Yd)]와 [오리피스 계수(ΔH@)]를 먼저 정확히 입력해주세요.'); return; }
    if (Vs === 0 || isNaN(Vs) || isNaN(Ts) || isNaN(Ps) || isNaN(dpAvg) || dpAvg === 0) { alert('오류: 4번 항목의 유속(동압 및 온도) 기초 데이터가 먼저 입력되어야 산정이 가능합니다.'); return; }

    setRecommendations({ fast: generateRecommendation('fast'), standard: generateRecommendation('standard'), stable: generateRecommendation('stable') });
  };

  const applyRecommendation = (opt) => {
    setFormData(prev => {
        const newMeters = prev.gasMeters.map((meter, idx) => {
            if (idx === 0) return meter;
            const dpValue = parseFloat(meter.dp);
            if (!isNaN(parseFloat(opt.calculatedK)) && !isNaN(dpValue)) return { ...meter, pressure: (parseFloat(opt.calculatedK) * dpValue).toFixed(2) };
            return meter;
        });

        return { 
          ...prev, 
          usedNozzleNum: String(opt.bestNozzleNum), 
          nozzleDiameter: opt.finalDn, 
          kFactor: opt.calculatedK, 
          planSamplingTime: String(opt.fastestTime), 
          recommendedNozzleNum: String(opt.bestNozzleNum), 
          recommendedNozzleDia: opt.finalDn, 
          gasMeters: newMeters 
        };
    });
  };

  const getExpectedValues = () => {
    const d = parseFloat(formData.nozzleDiameter), k = parseFloat(formData.kFactor), dp = getRawAvgDp(), Vs = getRawGasVelocity();

    let dH = '-', L = '-', SL = '-', Sm3 = '-', reqTime = '-';
    if (!isNaN(k) && !isNaN(dp)) dH = (k * dp).toFixed(2);

    const Ts = getRawAvgTs(), Ps = getRawStackPressure(), Pm = parseFloat(formData.atmPressure);
    const tmIn = parseFloat(formData.expectedTmIn) || parseFloat(formData.atmTemp) || 25;
    const tmOut = parseFloat(formData.expectedTmOut) || parseFloat(formData.atmTemp) || 25;
    const Tm = (tmIn + tmOut) / 2;

    const { Xw } = getGasComposition();
    const Yd = parseFloat(formData.gasMeterFactor) || 1.0;
    const targetVol = parseFloat(formData.targetVolume) || 1000;

    if (!isNaN(d) && !isNaN(Vs) && Vs > 0 && !isNaN(Ts) && !isNaN(Ps) && Pm > 0 && !isNaN(Xw)) {
        const An = Math.PI * Math.pow(d / 2000, 2);
        const Q_m = (Vs * An * 60 * 1000 * ((Tm + 273) / (Ts + 273)) * (Ps / Pm) * (1 - Xw / 100)) / Yd;
        const Q_sl = Q_m * Yd * (273 / (273 + Tm)) * (Pm / 760);
        
        const calcTime = Math.ceil(targetVol / Q_sl);
        if (isFinite(calcTime) && calcTime > 0) {
            reqTime = calcTime;
            const Vm_val = Q_m * reqTime;
            L = Vm_val.toFixed(1);
            const Vsl_val = Q_sl * reqTime;
            SL = Vsl_val.toFixed(1);
            Sm3 = (Vsl_val / 1000).toFixed(3);
        }
    }
    return { dH, L, SL, Sm3, reqTime };
  };

  const calcRowIsokineticRate = (idx) => {
    if (idx === 0) return '-';
    const meter = formData.gasMeters[idx];
    
    let prevVolStr = '';
    for (let i = idx - 1; i >= 0; i--) {
      if (formData.gasMeters[i].volume) { prevVolStr = formData.gasMeters[i].volume; break; }
    }
    
    const currVolStr = meter.volume, pressureStr = meter.pressure, inT = meter.tmIn, outT = meter.tmOut, timeStr = meter.time; 
    if (!prevVolStr || !currVolStr || !timeStr || !pressureStr || (inT === '' && outT === '')) return '-';
    
    const prevVol = parseFloat(prevVolStr), currVol = parseFloat(currVolStr), Vm = currVol - prevVol; 
    if (Vm <= 0 || isNaN(Vm)) return '-';
    
    let theta = parseFloat(timeStr);
    if (isNaN(theta) || theta <= 0) return '-';
    
    let Tm;
    if (inT !== '' && outT !== '') Tm = (parseFloat(inT) + parseFloat(outT)) / 2;
    else if (inT !== '') Tm = parseFloat(inT);
    else Tm = parseFloat(outT);
    if (isNaN(Tm)) return '-';
    
    let Ts = parseFloat(meter.stackTemp);
    if (isNaN(Ts)) Ts = getRawAvgTs();
    if (isNaN(Ts)) return '-';
    
    const Pa = parseFloat(formData.atmPressure);
    const Pm_val = parseFloat(pressureStr);
    const Ps = getRawStackPressure(), Cp = parseFloat(formData.pitotFactor) || 0.84;
    const Y = parseFloat(formData.gasMeterFactor) || 1.0;
    if (isNaN(Ps) || isNaN(Pa) || isNaN(Pm_val)) return '-';
    
    const Pm_abs = Pa + Pm_val / 13.6;
    const { Ms, Xw } = getGasComposition();
    
    const dpStr = meter.dp;
    if (!dpStr) return '-';
    const dp = parseFloat(dpStr);
    if (dp < 0 || isNaN(dp)) return '-';

    const velocityConstant = Math.sqrt((2 * 9.81 * 760 * 22.4) / 273);
    const Vs = velocityConstant * Cp * Math.sqrt(dp) * Math.sqrt((273 + Ts) / (Ps * Ms));
    if (Vs <= 0 || isNaN(Vs)) return '-';
    
    const Dn = parseFloat(formData.nozzleDiameter);
    if (isNaN(Dn) || Dn <= 0) return '-';
    const An = Math.PI * Math.pow(Dn / 2000, 2);
    
    const rate = (Vm * Y * (Ts + 273) * Pm_abs * 100) / ((Tm + 273) * Vs * An * 60 * 1000 * theta * Ps * (1 - Xw / 100));
    return rate.toFixed(1);
  };

  const calcIsokineticRate = (isPost = false) => {
    const Vm = getRawGasMeterVolDiff();
    const Ts = getRawAvgTs();
    const rawTm = getRawAvgTm();
    const Tm = isNaN(rawTm) ? Ts : rawTm; 
    const Dn = parseFloat(formData.nozzleDiameter);
    const Vs = getRawGasVelocity();
    const theta = getSamplingMinutes();
    const Y = parseFloat(formData.gasMeterFactor) || 1.0;
    const Pa = parseFloat(formData.atmPressure);
    const dH_avg = getRawAvgOrifice();
    const Ps = getRawStackPressure();

    if (Vm > 0 && Dn > 0 && Vs > 0 && theta > 0 && !isNaN(Ts) && !isNaN(Pa) && !isNaN(dH_avg) && !isNaN(Ps)) {
      const An = Math.PI * Math.pow(Dn / 2000, 2);
      const Pm_abs = Pa + dH_avg / 13.6;

      if (isPost) {
        const Vic = getRawTotalMoistureWeight(); 
        const Vm_m3 = Vm / 1000;
        const part1 = 0.00346 * Vic;
        const part2 = (Vm_m3 * Y / (Tm + 273)) * Pm_abs;
        const I = (Ts + 273) * (part1 + part2) * 1.667 / (theta * Vs * Ps * An);
        return I.toFixed(1);
      } else {
        const { Xw } = getGasComposition();
        const I = (Vm * Y * (Ts + 273) * Pm_abs * 100) / ((Tm + 273) * Vs * An * 60 * 1000 * theta * Ps * (1 - Xw / 100));
        return I.toFixed(1);
      }
    }
    return '-';
  };

  const calcDustConcentrations = () => {
    const Vm_std = getRawVmStd();
    const Wm = getRawDustWeightDiff();
    let actualC = '-', correctedC = '-';

    if (Vm_std > 0 && Wm > 0) {
      const C = Wm / Vm_std; 
      actualC = C.toFixed(2);
      const K = calcO2CorrectionFactor();
      correctedC = (C * K).toFixed(2);
    }
    return { actualC, correctedC };
  };

  const handleSave = (e) => {
    e.preventDefault();
    const { actualC, correctedC } = calcDustConcentrations();
    const result = {
      ...formData, moisturePercent: calcTotalMoistureWeight() > 0 ? calcPostMoisture() : calcMoisture(),
      avgVelocity: calcGasVelocity(), isokineticRate: calcIsokineticRate(true), dustWeight: calcDustWeightDiff(),
      actualConcentration: actualC, correctedConcentration: correctedC
    };
    setSavedData([...savedData, result]);
    alert('삐까삐까! 작성하신 데이터가 임시 저장되었습니다. ⚡');
  };

  const exportToCSV = () => {
    if (savedData.length === 0) {
      alert('추출할 데이터가 없습니다.');
      return;
    }

    const BOM = '\uFEFF';
    const headers = ['측정일자', '사업장명', '배출구명', '수분량(%)', '등속흡인율(%)', '먼지무게(mg)', '실측농도', '보정농도'].join(',');
    
    const rows = savedData.map(data => 
      [
        data.date,
        data.company,
        data.location,
        data.moisturePercent,
        data.isokineticRate,
        data.dustWeight,
        data.actualConcentration,
        data.correctedConcentration
      ].map(val => `"${String(val).replace(/"/g, '""')}"`).join(',')
    );

    const csvContent = BOM + [headers, ...rows].join('\n');
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    
    const link = document.createElement('a');
    link.href = url;
    link.setAttribute('download', `먼지시료채취_결과리포트_${new Date().toISOString().slice(0, 10)}.csv`);
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  };

  const { o2, co2, co, sox, nox, n2, Md, Ms, r0 } = getGasComposition();
  const currentTs = getRawAvgTs();
  const currentPs = getRawStackPressure();
  const r = (!isNaN(r0) && !isNaN(currentTs) && !isNaN(currentPs)) ? r0 * (273 / (273 + currentTs)) * (currentPs / 760) : NaN;
  const samplingPointsData = getSamplingPoints();
  const expData = getExpectedValues();

  return (
    <div className="min-h-screen bg-yellow-50 p-4 md:p-8 font-sans text-slate-800 selection:bg-yellow-200">
      <div className="max-w-7xl mx-auto space-y-6">
        
        {/* 헤더 섹션 */}
        <div className="bg-white p-6 rounded-xl shadow-md border-t-4 border-t-yellow-400 flex items-center justify-between">
          <div>
            <h1 className="text-2xl font-black text-slate-900 flex items-center gap-2">
              <Zap className="text-yellow-500 w-8 h-8 fill-yellow-400" />
              먼지 시료채취 종합 기록부 ⚡
            </h1>
            <p className="text-slate-600 mt-2 text-sm font-medium">대기오염공정시험기준에 맞춘 전 항목 종합 연산 및 자동 기록 시트</p>
          </div>
        </div>

        <form onSubmit={handleSave} className="space-y-6">
          
          {/* 섹션 1: 일반 사항 및 배출구 제원 */}
          <div className="bg-white p-6 rounded-xl shadow-sm border border-yellow-200">
            <h2 className="text-lg font-bold text-slate-800 mb-4 border-b border-yellow-100 pb-2 flex items-center gap-2">
              <span className="bg-yellow-400 text-slate-900 font-black w-6 h-6 rounded-full flex items-center justify-center text-sm">1</span>
              일반 사항 및 배출구 제원
            </h2>
            <div className="grid grid-cols-1 md:grid-cols-4 gap-4 mb-4">
              <div><label className="block text-xs font-bold text-slate-600 mb-1">측정 일자</label>
                <input type="date" name="date" value={formData.date} onChange={handleChange} className="w-full p-2 bg-yellow-50/50 border border-yellow-200 rounded focus:ring-2 focus:ring-yellow-400 outline-none transition-colors" /></div>
              <div><label className="block text-xs font-bold text-slate-600 mb-1">사업장명</label>
                <input type="text" name="company" value={formData.company} onChange={handleChange} className="w-full p-2 bg-yellow-50/50 border border-yellow-200 rounded focus:ring-2 focus:ring-yellow-400 outline-none transition-colors" /></div>
              <div><label className="block text-xs font-bold text-slate-600 mb-1">배출구명 (위치)</label>
                <input type="text" name="location" value={formData.location} onChange={handleChange} className="w-full p-2 bg-yellow-50/50 border border-yellow-200 rounded focus:ring-2 focus:ring-yellow-400 outline-none transition-colors" /></div>
              <div><label className="block text-xs font-bold text-slate-600 mb-1">측정자</label>
                <input type="text" name="sampler" value={formData.sampler} onChange={handleChange} className="w-full p-2 bg-yellow-50/50 border border-yellow-200 rounded focus:ring-2 focus:ring-yellow-400 outline-none transition-colors" /></div>
            </div>
            <div className="grid grid-cols-1 md:grid-cols-4 gap-4 bg-yellow-100/50 p-4 rounded-lg border border-yellow-200">
              <div><label className="block text-xs font-bold text-slate-600 mb-1">날씨</label>
                <input type="text" name="weather" value={formData.weather} onChange={handleChange} className="w-full p-2 border border-yellow-300 rounded focus:ring-2 focus:ring-yellow-400 outline-none" /></div>
              <div><label className="block text-xs font-bold text-slate-600 mb-1">대기 온도 (℃)</label>
                <input type="number" step="0.1" name="atmTemp" value={formData.atmTemp} onChange={handleChange} className="w-full p-2 border border-yellow-300 rounded focus:ring-2 focus:ring-yellow-400 outline-none" /></div>
              <div className="md:col-span-2"><label className="block text-xs font-bold text-slate-600 mb-1">대기압 (측정공 주변, mmHg)</label>
                <input type="number" step="1" name="atmPressure" value={formData.atmPressure} onChange={handleChange} className="w-full p-2 border border-yellow-300 rounded focus:ring-2 focus:ring-yellow-400 outline-none" /></div>
            </div>

            {/* 굴뚝 제원 (직경 산출) */}
            <div className="mt-4 grid grid-cols-1 md:grid-cols-4 gap-4 bg-yellow-50 p-4 rounded-lg border border-yellow-200">
              <div>
                <label className="block text-xs font-bold text-yellow-900 mb-1">총 측정 깊이 (m)</label>
                <input type="number" step="0.001" name="totalStackDepth" value={formData.totalStackDepth} onChange={handleChange} placeholder="끝벽까지의 총 길이" className="w-full p-2 border border-yellow-300 rounded focus:ring-2 focus:ring-yellow-400 bg-white outline-none" />
              </div>
              <div>
                <label className="block text-xs font-bold text-yellow-900 mb-1">플랜지 길이 (m)</label>
                <input type="number" step="0.001" name="flangeLength" value={formData.flangeLength} onChange={handleChange} placeholder="외벽 돌출부" className="w-full p-2 border border-yellow-300 rounded focus:ring-2 focus:ring-yellow-400 bg-white outline-none" />
              </div>
              <div>
                <label className="block text-xs font-bold text-slate-900 mb-1">순수 굴뚝 내경 (D, m)</label>
                <input type="number" step="0.001" name="stackDiameter" value={formData.stackDiameter} onChange={handleChange} className="w-full p-2 border border-yellow-400 rounded focus:ring-2 focus:ring-yellow-500 bg-yellow-100 font-black text-slate-900 outline-none" placeholder="자동계산 (총 깊이 - 플랜지)" />
              </div>
              <div>
                <label className="block text-xs font-bold text-slate-900 mb-1">단면적 (A, ㎡)</label>
                <input type="text" readOnly value={samplingPointsData ? samplingPointsData.area : ''} className="w-full p-2 border border-yellow-400 rounded bg-yellow-100 font-black text-slate-900 outline-none" placeholder="자동계산" />
              </div>
            </div>

            {/* 굴뚝 내경에 따른 측정점 산출 패널 */}
            {samplingPointsData && (
              <div className="mt-4 bg-white p-4 rounded-lg border border-yellow-200 shadow-sm">
                <div className="flex justify-between items-start mb-3">
                  <h3 className="text-sm font-bold text-slate-800 flex items-center gap-1">
                    <Target className="w-4 h-4 text-red-500"/> 굴뚝 내경({formData.stackDiameter}m) 연동 측정점 산출
                  </h3>
                  {samplingPointsData.isCenterOnly && (
                    <span className="bg-red-100 text-red-700 px-2 py-1 rounded text-xs font-bold animate-pulse border border-red-200">
                      ⚠️ 단면적 0.25㎡ 이하 (단면중심 1점 측정)
                    </span>
                  )}
                </div>
                
                <div className="flex flex-col xl:flex-row gap-4 items-start">
                  <div className="flex flex-col gap-2 shrink-0">
                    <div className="bg-yellow-100 px-3 py-2 rounded border border-yellow-300 text-center shadow-sm">
                      <span className="block text-[10px] text-slate-600 font-bold">1개 측정공(Port) 기준</span>
                      <span className="font-black text-slate-900 text-lg">총 {samplingPointsData.perRadius}개</span>
                    </div>
                  </div>
                  
                  <div className="flex-1 w-full overflow-x-auto bg-white rounded border border-yellow-200">
                    <table className="w-full text-xs text-center border-collapse">
                      <thead>
                        <tr className="bg-yellow-50 text-slate-700">
                          <th className="border-r border-b border-yellow-200 p-2 whitespace-nowrap" rowSpan="2">삽입 깊이 (마킹 위치)</th>
                          <th className="border-b border-yellow-200 p-2 whitespace-nowrap" colSpan={samplingPointsData.perRadius}>측정 포인트 (1~5)</th>
                        </tr>
                        <tr className="bg-yellow-50 text-slate-700">
                          {samplingPointsData.nearInsertion.map((_, i) => (
                            <th key={i} className="border-b border-l border-yellow-200 p-1">{i + 1}번</th>
                          ))}
                        </tr>
                      </thead>
                      <tbody>
                        <tr>
                          <td className="border-r border-b border-yellow-200 p-2 font-bold text-slate-800 whitespace-nowrap bg-yellow-100/50">
                            가까운 쪽 (정방향)<br/><span className="text-[9px] font-normal text-slate-600">기본 측정</span>
                          </td>
                          {samplingPointsData.nearInsertion.map((d, i) => (
                            <td key={i} className="border-l border-b border-yellow-200 p-2 font-black text-slate-800 bg-white">
                              {d} m
                            </td>
                          ))}
                        </tr>
                        {!samplingPointsData.isCenterOnly && (
                          <tr>
                            <td className="border-r border-yellow-200 p-2 font-bold text-red-700 whitespace-nowrap bg-red-50">
                              깊은 쪽 (역방향)<br/><span className="text-[9px] font-normal text-red-500">반대편 측정공 대체시</span>
                            </td>
                            {samplingPointsData.farInsertion.map((d, i) => (
                              <td key={i} className="border-l border-yellow-200 p-2 font-black text-red-600 bg-red-50/50">
                                {d} m
                              </td>
                            ))}
                          </tr>
                        )}
                      </tbody>
                    </table>
                  </div>
                </div>
                
                <div className="mt-4 flex flex-col md:flex-row justify-between items-center gap-3 border-t border-yellow-100 pt-4">
                  <div className="text-[10px] text-slate-500 text-left w-full space-y-1">
                    <p>※ 반대편 측정공이 없어 깊게 찔러야 할 경우에도 <strong>기록표의 칸수는 변하지 않으며</strong>, 위 표의 <strong>[깊은 쪽] 삽입 깊이 수치</strong>만 보시고 1~5번에 맞춰 찌르시면 됩니다.</p>
                  </div>
                  <button type="button" onClick={applySamplingPointsToTable} className="w-full md:w-auto py-2 px-5 bg-yellow-400 text-slate-900 rounded-lg font-black text-sm hover:bg-yellow-500 shadow-sm flex items-center justify-center gap-2 transition-colors shrink-0 border border-yellow-500">
                    <ListOrdered className="w-4 h-4" /> 하단 표 {samplingPointsData.perRadius}칸 자동 생성
                  </button>
                </div>
              </div>
            )}
          </div>

          {/* 섹션 2: 배출가스 조성 및 분자량 (가스분석기 자료) */}
          <div className="bg-white p-6 rounded-xl shadow-sm border border-yellow-200">
            <div className="flex flex-col md:flex-row md:items-center justify-between mb-4 border-b border-yellow-100 pb-2">
              <h2 className="text-lg font-bold text-slate-800 flex items-center gap-2">
                <span className="bg-red-500 text-white w-6 h-6 rounded-full flex items-center justify-center text-sm font-black">2</span>
                배출가스 조성 및 밀도 (가스분석기 5분 간격 3회 측정)
              </h2>
            </div>
            
            <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
               <div className="lg:col-span-2 overflow-x-auto">
                  <table className="w-full text-sm text-center border-collapse">
                    <thead>
                      <tr className="bg-red-50 text-red-800 border-b border-red-200">
                        <th className="p-2 border-r border-red-100 font-bold whitespace-nowrap">회차 (시간)</th>
                        <th className="p-2 border-r border-red-100 font-bold">O₂ (%)</th>
                        <th className="p-2 border-r border-red-100 font-bold">CO₂ (%)</th>
                        <th className="p-2 border-r border-red-100 font-bold">CO (ppm)</th>
                        <th className="p-2 border-r border-red-100 font-bold">SOx (ppm)</th>
                        <th className="p-2 font-bold">NOx (ppm)</th>
                      </tr>
                    </thead>
                    <tbody>
                      {formData.gasAnalyzer.map((gas, idx) => (
                        <tr key={idx} className="hover:bg-yellow-50 border-b border-yellow-100">
                          <td className="p-2 border-r border-yellow-100 font-medium text-slate-700">{idx + 1}회차 ({(idx + 1) * 5}분)</td>
                          <td className="p-1 border-r border-yellow-100">
                            <input type="number" step="0.01" value={gas.o2} onChange={(e) => handleGasAnalyzerChange(idx, 'o2', e.target.value)} className="w-full p-1 border border-yellow-300 rounded text-center outline-none focus:ring-2 focus:ring-yellow-400" />
                          </td>
                          <td className="p-1 border-r border-yellow-100">
                            <input type="number" step="0.01" value={gas.co2} onChange={(e) => handleGasAnalyzerChange(idx, 'co2', e.target.value)} className="w-full p-1 border border-yellow-300 rounded text-center outline-none focus:ring-2 focus:ring-yellow-400" />
                          </td>
                          <td className="p-1 border-r border-yellow-100">
                            <input type="number" step="0.01" value={gas.co} onChange={(e) => handleGasAnalyzerChange(idx, 'co', e.target.value)} className="w-full p-1 border border-yellow-300 rounded text-center outline-none focus:ring-2 focus:ring-yellow-400" />
                          </td>
                          <td className="p-1 border-r border-yellow-100">
                            <input type="number" step="0.01" value={gas.sox} onChange={(e) => handleGasAnalyzerChange(idx, 'sox', e.target.value)} className="w-full p-1 border border-yellow-300 rounded text-center outline-none focus:ring-2 focus:ring-yellow-400" />
                          </td>
                          <td className="p-1">
                            <input type="number" step="0.01" value={gas.nox} onChange={(e) => handleGasAnalyzerChange(idx, 'nox', e.target.value)} className="w-full p-1 border border-yellow-300 rounded text-center outline-none focus:ring-2 focus:ring-yellow-400" />
                          </td>
                        </tr>
                      ))}
                      <tr className="bg-yellow-100/50 font-bold text-slate-800">
                        <td className="p-2 border-r border-yellow-200">평균</td>
                        <td className="p-2 border-r border-yellow-200 text-red-600">{o2.toFixed(2)}</td>
                        <td className="p-2 border-r border-yellow-200">{co2.toFixed(2)}</td>
                        <td className="p-2 border-r border-yellow-200">{co.toFixed(2)}</td>
                        <td className="p-2 border-r border-yellow-200">{sox.toFixed(2)}</td>
                        <td className="p-2">{nox.toFixed(2)}</td>
                      </tr>
                    </tbody>
                  </table>
               </div>
               
               <div className="lg:col-span-1 flex flex-col gap-4">
                 {/* O2 보정 계수 패널 */}
                 <div className="bg-white p-4 rounded-lg border border-red-300 shadow-sm relative overflow-hidden">
                    <div className="absolute top-0 left-0 w-1.5 h-full bg-red-500"></div>
                    <h3 className="text-xs font-bold text-red-800 mb-3 flex items-center gap-1 border-b border-red-100 pb-2">
                      <Target className="w-4 h-4"/> 산소(O₂) 보정 계수 산출
                    </h3>
                    <div className="flex justify-between items-center mb-2">
                       <div className="flex flex-col">
                         <span className="text-xs font-bold text-slate-700">표준 산소 농도 (O<sub>s</sub>, %)</span>
                         <span className="text-[9px] text-slate-500">보정 대상이 아니면 비워두세요</span>
                       </div>
                       <input type="number" step="0.1" name="standardO2" value={formData.standardO2} onChange={handleChange} className="w-20 p-1 border border-red-300 rounded text-center text-sm font-black text-red-700 bg-red-50 focus:ring-2 focus:ring-red-400 outline-none" placeholder="공란" />
                    </div>
                    <div className="flex justify-between items-center mb-2">
                       <span className="text-xs font-bold text-slate-700">실측 산소 농도 (O<sub>a</sub>, %)</span>
                       <span className="font-bold text-slate-800">{o2.toFixed(2)} %</span>
                    </div>
                    <div className="flex justify-between items-center pt-2 mt-1 border-t border-red-100">
                       <span className="text-xs font-bold text-red-700" title="(21 - 표준O2) / (21 - 실측O2)">산소 보정 계수 (K)</span>
                       <span className="font-black text-lg text-red-600">{calcO2CorrectionFactor().toFixed(3)}</span>
                    </div>
                 </div>

                 {/* 분자량/밀도 패널 */}
                 <div className="bg-orange-50 p-4 rounded-lg border border-orange-200">
                    <div className="flex justify-between items-center mb-1">
                       <span className="text-[11px] font-bold text-orange-800" title="100 - O₂(%) - CO₂(%) - CO(%)">평균 N₂ (%)</span>
                       <span className="text-xs font-bold text-orange-700">{n2.toFixed(2)}</span>
                    </div>
                    <div className="flex justify-between items-center mb-1">
                       <span className="text-[11px] font-bold text-orange-800" title="Σ(M_i * x_i) / 100">건조가스분자량 (M<sub>d</sub>)</span>
                       <span className="text-xs font-black text-orange-900">{Md.toFixed(3)}</span>
                    </div>
                    <div className="flex justify-between items-center mb-1 pt-1 border-t border-orange-200/50">
                       <span className="text-[11px] font-bold text-orange-800" title="사전 수분량 반영됨">습가스분자량 (M<sub>s</sub>)</span>
                       <span className="text-sm font-black text-orange-900">{Ms.toFixed(3)}</span>
                    </div>
                    <div className="flex justify-between items-center pt-1 mt-1 border-t border-orange-200/50">
                       <span className="text-[11px] font-bold text-orange-800" title="1 / (22.4 * 100) * [ Σ(M_i*x_i) * (100 - Xw)/100 + 18 * Xw ]">배출가스밀도 (γ<sub>a</sub>)</span>
                       <span className="text-sm font-black text-red-600">{r0.toFixed(3)}</span>
                    </div>
                    <div className="flex justify-between items-center pt-1 mt-1 border-t border-orange-200/50">
                       <span className="text-[11px] font-bold text-orange-800" title="온도 및 압력 보정된 실제 밀도">배출가스밀도 (r)</span>
                       <span className="text-sm font-black text-red-600">{isNaN(r) ? '-' : r.toFixed(3)}</span>
                    </div>
                 </div>
               </div>
            </div>
          </div>

          <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
            {/* 섹션 3: 배출가스 수분량 */}
            <div className="bg-white p-6 rounded-xl shadow-sm border border-yellow-200">
              <h2 className="text-lg font-bold text-slate-800 mb-4 border-b border-yellow-100 pb-2 flex items-center gap-2">
                <span className="bg-yellow-500 text-slate-900 font-black w-6 h-6 rounded-full flex items-center justify-center text-sm">3</span>
                배출가스 수분량 (자동측정기법)
              </h2>
              <div className="space-y-4">
                <div>
                  <label className="block text-xs font-bold text-slate-600 mb-2">사전 수분량 (%) - 5회 측정값 입력</label>
                  <div className="grid grid-cols-5 gap-2">
                    {formData.moistureValues.map((val, idx) => (
                      <div key={idx}>
                        <label className="block text-[10px] text-slate-500 text-center mb-1">{idx + 1}회차</label>
                        <input 
                          type="number" step="0.1" value={val} 
                          onChange={(e) => handleMoistureChange(idx, e.target.value)} 
                          className="w-full p-2 border border-yellow-300 rounded focus:ring-2 focus:ring-yellow-400 bg-yellow-50 text-center text-sm outline-none" 
                        />
                      </div>
                    ))}
                  </div>
                </div>
                
                <div className="bg-yellow-100 p-4 rounded-lg flex justify-between items-center border border-yellow-300 mt-4 shadow-sm">
                  <div className="flex items-center gap-2 text-slate-800">
                    <Droplets className="w-5 h-5 text-blue-500" />
                    <span className="font-bold">사전 적용 수분량 (평균)</span>
                  </div>
                  <div className="text-2xl font-black text-slate-900">{calcMoisture()}%</div>
                </div>
                
                <div className="mt-6 pt-4 border-t border-yellow-200">
                  <label className="block text-sm font-bold text-slate-800 mb-2 flex items-center gap-1">
                    사후 수분량 (임핀저 무게법) <span className="text-slate-500 font-normal text-xs">- 최종 등속흡인계수 계산용</span>
                  </label>
                  <div className="overflow-x-auto">
                    <table className="w-full text-sm text-center border-collapse">
                      <thead>
                        <tr className="bg-yellow-200 text-slate-800">
                          <th className="p-2 border border-yellow-300 font-bold rounded-tl-lg">임핀저</th>
                          <th className="p-2 border border-yellow-300 font-bold">채취 전 (g)</th>
                          <th className="p-2 border border-yellow-300 font-bold">채취 후 (g)</th>
                          <th className="p-2 border border-yellow-300 font-bold rounded-tr-lg">증가량 (g)</th>
                        </tr>
                      </thead>
                      <tbody>
                        {formData.impingers.map((imp, idx) => (
                          <tr key={imp.id} className="hover:bg-yellow-50">
                            <td className="p-2 border border-yellow-200 text-slate-700 font-bold">{imp.id}단{imp.id === 4 ? '(실리카겔)' : ''}</td>
                            <td className="p-1 border border-yellow-200">
                              <input type="number" step="0.01" value={imp.initial} onChange={(e) => handleImpingerChange(idx, 'initial', e.target.value)} className="w-full p-1 border border-yellow-300 rounded text-center outline-none focus:ring-2 focus:ring-yellow-400" />
                            </td>
                            <td className="p-1 border border-yellow-200">
                              <input type="number" step="0.01" value={imp.final} onChange={(e) => handleImpingerChange(idx, 'final', e.target.value)} className="w-full p-1 border border-yellow-300 rounded text-center outline-none focus:ring-2 focus:ring-yellow-400" />
                            </td>
                            <td className="p-2 border border-yellow-200 font-black text-slate-900 bg-yellow-100/50">
                              {!isNaN(parseFloat(imp.initial)) && !isNaN(parseFloat(imp.final)) ? (parseFloat(imp.final) - parseFloat(imp.initial)).toFixed(2) : '-'}
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                  <div className="flex flex-col gap-2 mt-3">
                    <div className="flex justify-between items-center bg-yellow-50 p-3 rounded-lg border border-yellow-200">
                      <span className="text-sm font-bold text-slate-700">포집된 수분 총량 (V<sub>lc</sub>)</span>
                      <span className="font-black text-slate-900">{calcTotalMoistureWeight().toFixed(2)} ml</span>
                    </div>
                    <div className="flex justify-between items-center bg-red-500 p-3 rounded-lg shadow-sm text-white">
                      <span className="text-sm font-bold">사후 수분량 (X<sub>w</sub>)</span>
                      <span className="font-black text-xl text-yellow-300">{calcPostMoisture()}%</span>
                    </div>
                  </div>
                </div>
              </div>
            </div>

            {/* 섹션 4: 동압 및 유속 */}
            <div className="bg-white p-6 rounded-xl shadow-sm border border-yellow-200 flex flex-col h-full">
              <div className="flex justify-between items-center mb-4 border-b border-yellow-100 pb-2">
                <h2 className="text-lg font-bold text-slate-800 flex items-center gap-2">
                  <span className="bg-red-500 text-white w-6 h-6 rounded-full flex items-center justify-center text-sm font-black">4</span>
                  측정점별 동압 및 온도
                </h2>
                <div className="flex items-center gap-2">
                  <label className="text-xs font-bold text-slate-700">피토관 계수</label>
                  <input type="number" step="0.01" name="pitotFactor" value={formData.pitotFactor} onChange={handleChange} className="w-20 p-1 border border-yellow-300 rounded text-sm text-center outline-none focus:ring-2 focus:ring-red-400" />
                </div>
              </div>
              
              <div className="flex-1 overflow-auto bg-yellow-50/50 p-2 rounded border border-yellow-200 mb-4">
                <table className="w-full text-sm text-center">
                  <thead>
                    <tr className="text-slate-700 border-b border-yellow-300 bg-yellow-100/50">
                      <th className="pb-2 pt-2 rounded-tl-lg font-bold">측정점</th>
                      <th className="pb-2 pt-2 text-xs font-bold">전압<br/>(mmH2O)</th>
                      <th className="pb-2 pt-2 text-xs font-bold">정압<br/>(mmH2O)</th>
                      <th className="pb-2 pt-2 text-xs font-bold">동압<br/>(mmH2O)</th>
                      <th className="pb-2 pt-2 text-xs font-bold">온도(℃)</th>
                      <th className="pb-2 pt-2 rounded-tr-lg"></th>
                    </tr>
                  </thead>
                  <tbody>
                    {formData.points.map((point, idx) => (
                      <tr key={point.id} className="hover:bg-yellow-100/30">
                        <td className="py-1 font-bold text-slate-700">{point.id}</td>
                        <td className="py-1"><input type="number" step="0.1" value={point.tp !== undefined ? point.tp : ''} onChange={(e) => handlePointChange(idx, 'tp', e.target.value)} className="w-16 p-1 border border-yellow-300 rounded text-center text-sm outline-none focus:ring-2 focus:ring-yellow-400" /></td>
                        <td className="py-1"><input type="number" step="0.1" value={point.sp !== undefined ? point.sp : ''} onChange={(e) => handlePointChange(idx, 'sp', e.target.value)} className="w-16 p-1 border border-yellow-300 rounded text-center text-sm outline-none focus:ring-2 focus:ring-yellow-400" /></td>
                        <td className="py-1"><input type="number" step="0.1" value={point.dp} onChange={(e) => handlePointChange(idx, 'dp', e.target.value)} className="w-16 p-1 border border-red-300 bg-red-50 rounded text-center font-black text-red-600 text-sm outline-none focus:ring-2 focus:ring-red-400" /></td>
                        <td className="py-1"><input type="number" step="0.1" value={point.ts} onChange={(e) => handlePointChange(idx, 'ts', e.target.value)} className="w-16 p-1 border border-yellow-300 rounded text-center text-sm outline-none focus:ring-2 focus:ring-yellow-400" /></td>
                        <td className="py-1">
                          <button type="button" onClick={() => removePoint(idx)} className="text-red-400 hover:text-red-600 p-1 transition-colors"><Trash2 className="w-4 h-4" /></button>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
                <button type="button" onClick={addPoint} className="w-full mt-2 py-1.5 flex items-center justify-center gap-1 text-slate-800 bg-yellow-200 hover:bg-yellow-300 rounded text-xs font-black border border-yellow-400 transition-colors shadow-sm">
                  <Plus className="w-3 h-3" /> 측정점 추가
                </button>
              </div>

              <div className="bg-yellow-100 p-4 rounded-lg flex justify-between items-center border border-yellow-300 shadow-sm">
                <div className="flex items-center gap-2 text-slate-800">
                  <Wind className="w-5 h-5 text-slate-700" />
                  <span className="font-bold">평균 유속 (Vs)</span>
                </div>
                <div className="text-2xl font-black text-red-600">{calcGasVelocity()} <span className="text-sm font-bold text-slate-700">m/s</span></div>
              </div>
            </div>
          </div>

          {/* 섹션 5: 시료채취 및 등속흡인 */}
          <div className="bg-white p-6 rounded-xl shadow-sm border border-yellow-200">
            <h2 className="text-lg font-bold text-slate-800 mb-4 border-b border-yellow-100 pb-2 flex items-center gap-2">
              <span className="bg-yellow-400 text-slate-900 font-black w-6 h-6 rounded-full flex items-center justify-center text-sm">5</span>
              시료 채취 기록 (적산유량계 및 등속흡인)
            </h2>

            {/* 1단계: 자동 산정 및 예측 결과 통합 패널 */}
            <div className="bg-yellow-50 border border-yellow-300 p-5 rounded-lg mb-6 shadow-sm">
               <div className="flex flex-col md:flex-row md:justify-between md:items-end mb-4 gap-2">
                 <h3 className="font-bold text-slate-800 flex items-center gap-2">
                    <Target className="w-5 h-5 text-red-500" /> 1단계: 채취조건 정밀 산출 (기기 고유 계수 연동)
                 </h3>
                 <div className="flex items-center gap-2 bg-white px-3 py-1.5 rounded-lg border border-yellow-300 shadow-sm">
                    <label className="text-xs font-bold text-slate-700">장비 프리셋 선택</label>
                    <select 
                       value={formData.samplerId} 
                       onChange={handleSamplerChange}
                       className="p-1 border border-yellow-400 rounded text-sm font-bold text-slate-900 bg-yellow-100 focus:ring-2 focus:ring-yellow-500 outline-none cursor-pointer"
                    >
                       <option value="">직접 입력</option>
                       {SAMPLERS.map(s => (
                          <option key={s.id} value={s.id}>{s.name}</option>
                       ))}
                    </select>
                 </div>
               </div>
               
               <div className="grid grid-cols-2 md:grid-cols-5 gap-4 mb-3 items-end">
                  <div>
                    <label className="block text-xs font-bold text-slate-700 mb-1">목표 채취량 (SL)</label>
                    <input type="number" step="1" name="targetVolume" value={formData.targetVolume} onChange={handleChange} className="w-full p-2 border border-yellow-300 rounded text-sm bg-white font-bold outline-none focus:ring-2 focus:ring-yellow-400" />
                  </div>
                  <div>
                    <label className="block text-xs font-bold text-slate-700 mb-1" title="가스미터 보정계수">보정계수 (Y<sub>d</sub>)</label>
                    <input type="number" step="0.001" name="gasMeterFactor" value={formData.gasMeterFactor} onChange={handleChange} placeholder="Yd" className="w-full p-2 border border-yellow-300 rounded text-sm bg-white outline-none focus:ring-2 focus:ring-yellow-400 font-bold text-red-600" />
                  </div>
                  <div>
                    <label className="block text-xs font-bold text-slate-700 mb-1" title="오리피스 계수">오리피스 계수 (ΔH<sub>@</sub>)</label>
                    <input type="number" step="0.1" name="deltaHAt" value={formData.deltaHAt} onChange={handleChange} placeholder="ΔH@" className="w-full p-2 border border-yellow-300 rounded text-sm bg-white outline-none focus:ring-2 focus:ring-yellow-400 font-bold text-red-600" />
                  </div>
                  <div>
                    <label className="block text-[11px] font-bold text-slate-700 mb-1" title="채취 중 예상되는 입구 온도">예상 T<sub>m</sub> (입구, ℃)</label>
                    <input type="number" step="0.1" name="expectedTmIn" value={formData.expectedTmIn} onChange={handleChange} placeholder={formData.atmTemp ? formData.atmTemp : "25"} className="w-full p-2 border border-yellow-300 rounded text-sm bg-white outline-none focus:ring-2 focus:ring-yellow-400 font-bold text-slate-800" />
                  </div>
                  <div>
                    <label className="block text-[11px] font-bold text-slate-700 mb-1" title="채취 중 예상되는 출구 온도">예상 T<sub>m</sub> (출구, ℃)</label>
                    <input type="number" step="0.1" name="expectedTmOut" value={formData.expectedTmOut} onChange={handleChange} placeholder={formData.atmTemp ? formData.atmTemp : "25"} className="w-full p-2 border border-yellow-300 rounded text-sm bg-white outline-none focus:ring-2 focus:ring-yellow-400 font-bold text-slate-800" />
                  </div>
               </div>

               <div className="mb-4">
                  <button type="button" onClick={handleCalculateOptions} className="w-full py-3 bg-yellow-400 text-slate-900 rounded-lg font-black text-sm hover:bg-yellow-500 transition shadow-md flex justify-center items-center gap-2 border border-yellow-500">
                    <Calculator className="w-5 h-5" /> K-Factor 및 예상 채취시간 자동 산출
                  </button>
               </div>
               
               {recommendations && (
                 <div className="grid grid-cols-1 md:grid-cols-3 gap-4 mb-5">
                    <div className="bg-red-50 border-2 border-red-300 p-4 rounded-lg shadow-sm hover:border-red-500 transition-colors">
                       <h4 className="font-black text-red-800 mb-3 flex items-center gap-1 border-b border-red-200 pb-2">⚡ 빠른 채취 <span className="text-[10px] font-bold text-red-500 ml-auto">(최대치, 50 이하)</span></h4>
                       <ul className="text-sm text-slate-700 space-y-2 mb-4">
                          <li className="flex justify-between font-bold"><span>추천 노즐:</span> <span className="font-black text-lg text-slate-900">No.{recommendations.fast.bestNozzleNum} <span className="text-sm font-medium">({recommendations.fast.finalDn}mm)</span></span></li>
                          <li className="flex justify-between font-bold"><span>예상 시간:</span> <span className="font-black text-red-600">{recommendations.fast.fastestTime} 분</span></li>
                          <li className="flex justify-between font-bold"><span>K-Factor:</span> <span className="font-black">{recommendations.fast.calculatedK}</span></li>
                          <li className="flex justify-between items-center bg-white p-1 rounded border border-red-200 font-bold"><span>예상 ΔH:</span> <span className="font-black text-red-600 text-lg">{recommendations.fast.expectedDH}</span></li>
                       </ul>
                       <button type="button" onClick={() => applyRecommendation(recommendations.fast)} className="w-full py-2 bg-red-500 text-white rounded-lg text-sm font-bold hover:bg-red-600 shadow-sm flex justify-center items-center gap-1"><CheckCircle2 className="w-4 h-4"/> 적용하기</button>
                    </div>
                    <div className="bg-yellow-100 border-2 border-yellow-400 p-4 rounded-lg shadow-sm hover:border-yellow-500 transition-colors">
                       <h4 className="font-black text-slate-800 mb-3 flex items-center gap-1 border-b border-yellow-300 pb-2">🎯 표준 채취 <span className="text-[10px] font-bold text-slate-600 ml-auto">(ΔH 25 내외)</span></h4>
                       <ul className="text-sm text-slate-700 space-y-2 mb-4">
                          <li className="flex justify-between font-bold"><span>추천 노즐:</span> <span className="font-black text-lg text-slate-900">No.{recommendations.standard.bestNozzleNum} <span className="text-sm font-medium">({recommendations.standard.finalDn}mm)</span></span></li>
                          <li className="flex justify-between font-bold"><span>예상 시간:</span> <span className="font-black text-red-600">{recommendations.standard.fastestTime} 분</span></li>
                          <li className="flex justify-between font-bold"><span>K-Factor:</span> <span className="font-black">{recommendations.standard.calculatedK}</span></li>
                          <li className="flex justify-between items-center bg-white p-1 rounded border border-yellow-300 font-bold"><span>예상 ΔH:</span> <span className="font-black text-yellow-600 text-lg">{recommendations.standard.expectedDH}</span></li>
                       </ul>
                       <button type="button" onClick={() => applyRecommendation(recommendations.standard)} className="w-full py-2 bg-yellow-400 text-slate-900 rounded-lg text-sm font-bold hover:bg-yellow-500 shadow-sm flex justify-center items-center gap-1"><CheckCircle2 className="w-4 h-4"/> 적용하기</button>
                    </div>
                    <div className="bg-orange-50 border-2 border-orange-300 p-4 rounded-lg shadow-sm hover:border-orange-500 transition-colors">
                       <h4 className="font-black text-orange-900 mb-3 flex items-center gap-1 border-b border-orange-200 pb-2">🛡️ 안정/장시간 <span className="text-[10px] font-bold text-orange-600 ml-auto">(ΔH 10 내외)</span></h4>
                       <ul className="text-sm text-slate-700 space-y-2 mb-4">
                          <li className="flex justify-between font-bold"><span>추천 노즐:</span> <span className="font-black text-lg text-slate-900">No.{recommendations.stable.bestNozzleNum} <span className="text-sm font-medium">({recommendations.stable.finalDn}mm)</span></span></li>
                          <li className="flex justify-between font-bold"><span>예상 시간:</span> <span className="font-black text-red-600">{recommendations.stable.fastestTime} 분</span></li>
                          <li className="flex justify-between font-bold"><span>K-Factor:</span> <span className="font-black">{recommendations.stable.calculatedK}</span></li>
                          <li className="flex justify-between items-center bg-white p-1 rounded border border-orange-200 font-bold"><span>예상 ΔH:</span> <span className="font-black text-orange-600 text-lg">{recommendations.stable.expectedDH}</span></li>
                       </ul>
                       <button type="button" onClick={() => applyRecommendation(recommendations.stable)} className="w-full py-2 bg-orange-500 text-white rounded-lg text-sm font-bold hover:bg-orange-600 shadow-sm flex justify-center items-center gap-1"><CheckCircle2 className="w-4 h-4"/> 적용하기</button>
                    </div>
                 </div>
               )}
               
               <p className="text-[11px] text-slate-600 mb-4 font-bold">※ K-Factor 및 예상 오리피스압(ΔH) 산출 후, 해당 노즐의 유량(Q<sub>sl</sub>)을 바탕으로 <span className="text-red-600">목표 채취량({formData.targetVolume || 1000}SL)에 도달하기 위한 필요 채취시간</span>을 자동 역산합니다.</p>

               {/* 접기/펴기 엑셀 전체 리스트 */}
               <details className="mt-4 mb-4">
                  <summary className="cursor-pointer text-xs font-black text-slate-700 hover:text-slate-900 flex items-center gap-1">
                     <ChevronDown className="w-4 h-4"/> 전체 노즐 예측표 열어보기 (참고용)
                  </summary>
                  <div className="mt-3 overflow-x-auto">
                     <table className="w-full text-xs text-center border-collapse border border-yellow-300">
                       <thead className="bg-yellow-200 text-slate-800">
                          <tr>
                             <th className="border border-yellow-300 p-1 font-bold">노즐번호</th>
                             <th className="border border-yellow-300 p-1 font-bold">노즐직경(mm)</th>
                             <th className="border border-yellow-300 p-1 font-bold">K-factor</th>
                             <th className="border border-yellow-300 p-1 font-bold">ΔH</th>
                             <th className="border border-yellow-300 p-1 font-bold text-red-600">예상시간(분)</th>
                             <th className="border border-yellow-300 p-1 font-bold">(L)</th>
                             <th className="border border-yellow-300 p-1 font-bold">(SL)</th>
                             <th className="border border-yellow-300 p-1 font-bold">(Sm³)</th>
                          </tr>
                       </thead>
                       <tbody className="bg-white">
                          {NOZZLE_SET.map((n) => {
                             let k='-', dh='-', vl='-', vsl='-', vsm='-', reqTime='-';
                             const vs = getRawGasVelocity(), cP = parseFloat(formData.pitotFactor)||0.84, ts = getRawAvgTs(), pm = parseFloat(formData.atmPressure), ps = getRawStackPressure();
                             
                             const tmIn = parseFloat(formData.expectedTmIn) || parseFloat(formData.atmTemp) || 25;
                             const tmOut = parseFloat(formData.expectedTmOut) || parseFloat(formData.atmTemp) || 25;
                             const tm = (tmIn + tmOut) / 2;
                             
                             const validSps = formData.points.map(p => parseFloat(p.sp)).filter(v => !isNaN(v));
                             const avgSp = validSps.length === 0 ? 0 : validSps.reduce((a, b) => a + b, 0) / validSps.length;
                             const ps_ratio = (pm + avgSp / 13.6) / pm;

                             const { Md, Ms, Xw } = getGasComposition();
                             const y = parseFloat(formData.gasMeterFactor)||1.0;
                             const dHAt = parseFloat(formData.deltaHAt);
                             const Dn = n.d;

                             if(vs>0 && ps && pm>0) {
                                const k_val = 8.001e-5 * Math.pow(Dn, 4) * dHAt * Math.pow(cP, 2) * Math.pow(1 - Xw / 100, 2) * (Md / Ms) * ((tm + 273) / (ts + 273)) * ps_ratio;
                                k = k_val.toFixed(3);
                                dh = (k_val * getRawAvgDp()).toFixed(2);
                                const an = Math.PI * Math.pow(Dn/2000, 2);
                                const q_m = (vs * an * 60000 * ((tm+273)/(ts+273)) * (ps/pm) * (1-Xw/100)) / y;
                                const q_sl = q_m * y * (273/(tm+273)) * (pm/760);
                                
                                let time = Math.ceil((parseFloat(formData.targetVolume) || 1000) / q_sl);
                                if (!isFinite(time) || time <= 0) time = 0;
                                reqTime = time;
                                
                                const v_val = q_m * time;
                                vl = v_val.toFixed(1);
                                vsl = (q_sl * time).toFixed(1);
                                vsm = ((q_sl * time) / 1000).toFixed(3);
                             }
                             const isSelected = formData.usedNozzleNum === String(n.num);
                             return (
                               <tr key={n.num} className={isSelected ? "bg-yellow-100 font-bold" : "hover:bg-yellow-50"}>
                                 <td className="border border-yellow-200 p-1 font-bold text-slate-800">
                                   {n.num}
                                   {formData.recommendedNozzleNum === String(n.num) && (
                                     <span className="ml-1 text-[10px] bg-red-500 text-white px-1.5 py-0.5 rounded-full shadow-sm animate-pulse">추천</span>
                                   )}
                                 </td>
                                 <td className="border border-yellow-200 p-1 font-bold">{n.d.toFixed(2)}</td>
                                 <td className="border border-yellow-200 p-1 font-black text-slate-700">{k}</td>
                                 <td className="border border-yellow-200 p-1 font-black text-red-600">{dh}</td>
                                 <td className="border border-yellow-200 p-1 font-black text-red-600 bg-red-50/50">{reqTime}</td>
                                 <td className="border border-yellow-200 p-1">{vl}</td>
                                 <td className="border border-yellow-200 p-1">{vsl}</td>
                                 <td className="border border-yellow-200 p-1">{vsm}</td>
                               </tr>
                             )
                          })}
                       </tbody>
                     </table>
                  </div>
               </details>

               {/* 산정 결과 요약창 */}
               <div className="bg-white p-4 border border-yellow-400 rounded-lg shadow-inner">
                  {formData.recommendedNozzleNum && (
                    <div className="mb-4 p-3 bg-red-50 border border-red-200 rounded-lg flex items-center justify-between shadow-sm">
                      <div className="flex items-center gap-2">
                        <span className="text-xl">⚡</span>
                        <span className="font-black text-red-900">시스템 추천 적정 노즐:</span>
                        <span className="text-lg font-black text-slate-900 bg-yellow-300 px-3 py-1 border border-yellow-400 rounded-lg">
                          No.{formData.recommendedNozzleNum} ({formData.recommendedNozzleDia} mm)
                        </span>
                      </div>
                    </div>
                  )}
                  
                  <h4 className="text-sm font-black text-slate-800 mb-3 border-b border-yellow-200 pb-2">📊 노즐 선정 및 채취량 예측 정보</h4>
                  <div className="flex flex-wrap justify-center gap-2 text-center">
                      <div className="bg-slate-100 py-2 px-1 rounded-lg border border-slate-200 flex-1 min-w-[90px]">
                          <div className="text-[10px] text-slate-500 font-bold">가스유속(V<sub>s</sub>)</div>
                          <div className="font-black text-slate-800 mt-1">{calcGasVelocity()}</div>
                      </div>
                      <div className="bg-yellow-100 py-2 px-1 rounded-lg border border-yellow-300 flex-1 min-w-[90px]">
                          <div className="text-[10px] text-slate-700 font-bold">노즐경(D<sub>n</sub>)</div>
                          <div className="font-black text-slate-900 mt-1">{formData.nozzleDiameter || '-'}</div>
                      </div>
                      <div className="bg-yellow-100 py-2 px-1 rounded-lg border border-yellow-300 flex-1 min-w-[90px]">
                          <div className="text-[10px] text-slate-700 font-bold">K-Factor</div>
                          <div className="font-black text-slate-900 mt-1">{formData.kFactor || '-'}</div>
                      </div>
                      <div className="bg-red-100 py-2 px-1 rounded-lg border border-red-300 flex-1 min-w-[90px]">
                          <div className="text-[10px] text-red-700 font-bold">예상 ΔH</div>
                          <div className="font-black text-red-700 mt-1">{expData.dH}</div>
                      </div>
                      <div className="bg-red-500 py-2 px-1 rounded-lg border border-red-600 flex-1 min-w-[90px] shadow-sm">
                          <div className="text-[10px] text-red-100 font-bold">예상시간(분)</div>
                          <div className="font-black text-white mt-1">{expData.reqTime}</div>
                      </div>
                      <div className="bg-orange-100 py-2 px-1 rounded-lg border border-orange-300 flex-1 min-w-[90px]">
                          <div className="text-[10px] text-orange-800 font-bold">건조채취량(V<sub>m</sub>)</div>
                          <div className="font-black text-orange-900 mt-1">{expData.L}</div>
                      </div>
                      <div className="bg-orange-100 py-2 px-1 rounded-lg border border-orange-300 flex-1 min-w-[90px]">
                          <div className="text-[10px] text-orange-800 font-bold">표준채취량(SL)</div>
                          <div className="font-black text-orange-900 mt-1">{expData.SL}</div>
                      </div>
                      <div className="bg-orange-100 py-2 px-1 rounded-lg border border-orange-300 flex-1 min-w-[90px]">
                          <div className="text-[10px] text-orange-800 font-bold">표준채취량(Sm³)</div>
                          <div className="font-black text-orange-900 mt-1">{expData.Sm3}</div>
                      </div>
                  </div>
               </div>
            </div>
            
            {/* 기본 채취 조건 */}
            <div className="grid grid-cols-1 md:grid-cols-2 gap-4 mb-6 bg-yellow-50 p-4 rounded-xl border border-yellow-200">
              <div>
                <label className="block text-xs font-bold text-slate-700 mb-1">사용 노즐경 (mm)</label>
                <input 
                  type="number" step="0.01" list="nozzle-list" name="nozzleDiameter" 
                  value={formData.nozzleDiameter} onChange={handleChange} 
                  placeholder="목록 선택 또는 입력"
                  className="w-full p-2 border border-yellow-300 rounded-lg focus:ring-2 focus:ring-yellow-400 bg-white outline-none" 
                />
                <datalist id="nozzle-list">
                  {NOZZLE_SET.map(n => <option key={n.num} value={n.d.toFixed(2)}>No.{n.num} ({n.d}mm)</option>)}
                </datalist>
              </div>
              <div><label className="block text-xs font-bold text-slate-700 mb-1">적용 K-Factor</label>
                <input type="number" step="0.001" name="kFactor" value={formData.kFactor} onChange={handleKFactorChange} className="w-full p-2 border border-yellow-400 bg-white rounded-lg focus:ring-2 focus:ring-yellow-500 font-black outline-none text-slate-800" /></div>
            </div>

            {/* 적산유량계 기록표 */}
            <div className="mb-6">
              <div className="flex justify-between items-end mb-2">
                <h3 className="text-sm font-bold text-slate-800 flex items-center gap-1"><ListOrdered className="w-4 h-4"/> 2단계: 적산유량계 기록표 (대기오염공정시험기준 통합 양식)</h3>
              </div>
              
              <div className="overflow-x-auto border border-yellow-300 rounded-xl shadow-sm">
                <table className="w-full text-xs text-center min-w-[900px]">
                  <thead className="bg-yellow-200 text-slate-800 border-b border-yellow-300">
                    <tr>
                      <th className="p-1 font-bold whitespace-nowrap">순번</th>
                      <th className="p-1 font-bold whitespace-nowrap">채취점</th>
                      <th className="p-1 font-bold whitespace-nowrap">시간<br/>(분)</th>
                      <th className="p-1 font-bold whitespace-nowrap">배출가스<br/>온도(T<sub>s</sub>, ℃)</th>
                      <th className="p-1 font-black text-red-600 whitespace-nowrap">동압<br/>(Δp)</th>
                      <th className="p-1 font-black text-slate-900 whitespace-nowrap">오리피스압<br/>(ΔH)</th>
                      <th className="p-1 font-bold whitespace-nowrap">적산유량<br/>(V<sub>m</sub>, L)</th>
                      <th className="p-1 font-bold whitespace-nowrap">미터온도(T<sub>m</sub>, ℃)<br/><span className="text-[10px] text-slate-600">입구 | 출구</span></th>
                      <th className="p-1 font-bold whitespace-nowrap">진공압<br/>(mmHg)</th>
                      <th className="p-1 font-bold whitespace-nowrap">임핀저<br/>온도(℃)</th>
                      <th className="p-1 font-black text-red-600 border-l border-yellow-300 whitespace-nowrap">순간 등속<br/>(I%)</th>
                      <th className="p-1"></th>
                    </tr>
                  </thead>
                  <tbody className="bg-white">
                    {formData.gasMeters.map((meter, idx) => {
                      const isStartRow = idx === 0;
                      const rowRate = calcRowIsokineticRate(idx);
                      const isRateValid = rowRate !== '-' && parseFloat(rowRate) >= 95 && parseFloat(rowRate) <= 105;
                      
                      return (
                        <tr key={meter.id} className={`border-b border-yellow-100 last:border-0 hover:bg-yellow-50 transition-colors`}>
                          <td className="p-1 font-bold text-slate-500 w-8">{isStartRow ? '초기' : meter.id}</td>
                          <td className="p-1 w-12">
                            {isStartRow ? (
                              <span className="font-bold text-slate-400">시작</span>
                            ) : (
                              <input type="text" value={meter.pointNum} onChange={(e) => handleGasMeterChange(idx, 'pointNum', e.target.value)} placeholder="입력" className="w-full p-1 border border-yellow-200 rounded text-center outline-none focus:ring-2 focus:ring-yellow-400" />
                            )}
                          </td>
                          <td className="p-1 w-12">
                            {isStartRow ? (
                              <span className="font-bold text-slate-400">0</span>
                            ) : (
                              <input type="number" step="0.1" value={meter.time} onChange={(e) => handleGasMeterChange(idx, 'time', e.target.value)} placeholder="5" className="w-full p-1 border border-yellow-200 rounded text-center outline-none focus:ring-2 focus:ring-yellow-400" />
                            )}
                          </td>
                          <td className="p-1 w-14">
                            <input type="number" step="0.1" value={meter.stackTemp} onChange={(e) => handleGasMeterChange(idx, 'stackTemp', e.target.value)} placeholder="Ts" className={`w-full p-1 border border-yellow-200 rounded text-center outline-none focus:ring-2 focus:ring-yellow-400 ${isStartRow ? 'bg-slate-50' : ''}`} />
                          </td>
                          
                          <td className="p-1 w-14">
                            <input type="number" step="0.1" value={meter.dp} onChange={(e) => handleGasMeterChange(idx, 'dp', e.target.value)} className="w-full p-1 border border-red-300 bg-red-50 rounded text-center font-bold outline-none focus:ring-2 focus:ring-red-400" />
                          </td>
                          <td className="p-1 w-14">
                            <input type="number" step="0.1" value={meter.pressure} onChange={(e) => handleGasMeterChange(idx, 'pressure', e.target.value)} className="w-full p-1 border border-slate-300 bg-slate-100 rounded text-center font-black text-slate-800 outline-none focus:ring-2 focus:ring-slate-400" />
                          </td>
                          <td className="p-1 w-20">
                            <input type="number" step="0.1" value={meter.volume} onChange={(e) => handleGasMeterChange(idx, 'volume', e.target.value)} className={`w-full p-1 border ${isStartRow ? 'border-yellow-400 bg-yellow-100 text-slate-900 font-black' : 'border-yellow-200 bg-white font-bold text-slate-700'} rounded text-center outline-none focus:ring-2 focus:ring-yellow-400`} placeholder={isStartRow ? '초기유량' : ''} />
                          </td>
                          
                          <td className="p-1 w-28">
                            <div className="flex items-center gap-1">
                              <input type="number" step="0.1" value={meter.tmIn} onChange={(e) => handleGasMeterChange(idx, 'tmIn', e.target.value)} placeholder="입구" className="w-1/2 p-1 border border-orange-200 bg-orange-50 rounded text-center outline-none focus:ring-2 focus:ring-orange-400" />
                              <input type="number" step="0.1" value={meter.tmOut} onChange={(e) => handleGasMeterChange(idx, 'tmOut', e.target.value)} placeholder="출구" className="w-1/2 p-1 border border-orange-200 bg-orange-50 rounded text-center outline-none focus:ring-2 focus:ring-orange-400" />
                            </div>
                          </td>
                          
                          <td className="p-1 w-14">
                            <input type="number" step="0.1" value={meter.vacuum} onChange={(e) => handleGasMeterChange(idx, 'vacuum', e.target.value)} placeholder="압" className={`w-full p-1 border border-yellow-200 rounded text-center outline-none focus:ring-2 focus:ring-yellow-400 ${isStartRow ? 'bg-slate-50' : ''}`} />
                          </td>
                          <td className="p-1 w-14">
                            <input type="number" step="0.1" value={meter.impingerTemp} onChange={(e) => handleGasMeterChange(idx, 'impingerTemp', e.target.value)} placeholder="온도" className={`w-full p-1 border border-yellow-200 rounded text-center outline-none focus:ring-2 focus:ring-yellow-400 ${isStartRow ? 'bg-slate-50' : ''}`} />
                          </td>
                          
                          <td className="p-1 border-l border-yellow-200 w-16">
                            <span className={`inline-block w-full py-1 rounded-lg font-black ${rowRate === '-' ? 'text-slate-400 bg-slate-100' : isRateValid ? 'text-slate-900 bg-yellow-400' : 'text-white bg-red-500'}`}>
                              {rowRate === '-' ? '-' : `${rowRate}%`}
                            </span>
                          </td>
                          <td className="p-1 w-8">
                            {!isStartRow && <button type="button" onClick={() => removeGasMeter(idx)} className="text-red-400 hover:text-red-600 p-1 transition-colors"><Trash2 className="w-4 h-4" /></button>}
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
              <button type="button" onClick={addGasMeter} className="w-full mt-2 py-2 flex items-center justify-center gap-1 text-slate-800 bg-yellow-200 hover:bg-yellow-300 rounded-lg text-sm font-black border border-yellow-400 transition-colors shadow-sm">
                <Plus className="w-4 h-4" /> 측정 기록 칸 추가 ⚡
              </button>
            </div>

            {/* 유량계 기록 요약표 */}
            <div className="grid grid-cols-1 md:grid-cols-4 gap-4 mb-6">
              <div className="bg-slate-100 border border-slate-300 p-3 rounded-xl flex flex-col items-center justify-center">
                <span className="text-xs text-slate-600 font-bold mb-1">총 채취 시간 (분)</span>
                <span className="text-lg font-black text-slate-900">{getSamplingMinutes()}</span>
              </div>
              <div className="bg-yellow-100 border border-yellow-300 p-3 rounded-xl flex flex-col items-center justify-center">
                <span className="text-xs text-slate-700 font-bold mb-1">총 채취 가스량 (V<sub>m</sub>)</span>
                <span className="text-xl font-black text-slate-900">{calcGasMeterVolDiff() > 0 ? calcGasMeterVolDiff() : '-'}</span>
              </div>
              <div className="bg-orange-100 border border-orange-300 p-3 rounded-xl flex flex-col items-center justify-center">
                <span className="text-xs text-orange-800 font-bold mb-1">평균 온도 (T<sub>m</sub>, ℃)</span>
                <span className="text-lg font-black text-orange-900">{calcAvgTm()}</span>
              </div>
              <div className="bg-red-50 border border-red-200 p-3 rounded-xl flex flex-col items-center justify-center">
                <span className="text-xs text-red-700 font-bold mb-1">평균 오리피스압 (ΔH)</span>
                <span className="text-lg font-black text-red-700">{calcAvgOrifice()}</span>
              </div>
            </div>

            <div className="bg-slate-900 p-5 rounded-2xl flex flex-col justify-between text-white shadow-xl border-2 border-yellow-400">
              <div className="flex items-center gap-3 mb-4">
                <Gauge className="w-8 h-8 text-yellow-400" />
                <div>
                  <h3 className="font-black text-xl text-yellow-400 tracking-tight">전체 평균 등속흡인율 (Isokinetic Rate)</h3>
                  <p className="text-xs text-slate-300 mt-1">대기오염공정시험기준 유효범위 : 95% ~ 105% (수식 0.00346 적용됨)</p>
                </div>
              </div>
              
              <div className="grid grid-cols-2 gap-4 mt-2">
                <div className="bg-slate-800 p-4 rounded-xl border border-slate-700 flex flex-col items-center justify-center">
                  <span className="text-xs text-slate-400 mb-1 font-bold">사전 수분량 기준 (전체)</span>
                  <div className="text-2xl font-black text-white">{calcIsokineticRate(false)}<span className="text-sm ml-1 text-slate-500">%</span></div>
                </div>
                <div className="bg-red-500 p-4 rounded-xl border-2 border-red-400 flex flex-col items-center justify-center shadow-lg relative overflow-hidden">
                  <div className="absolute top-0 left-0 w-full h-1.5 bg-yellow-400"></div>
                  <span className="text-xs text-red-100 mb-1 font-black">사후 수분량 기준 (최종)</span>
                  <div className="text-4xl font-black text-yellow-300">{calcIsokineticRate(true)}<span className="text-xl ml-1 text-white">%</span></div>
                </div>
              </div>

              <div className="mt-5 flex justify-end">
                {calcIsokineticRate(true) !== '-' && (
                  <div>
                    {(parseFloat(calcIsokineticRate(true)) >= 95 && parseFloat(calcIsokineticRate(true)) <= 105) ? 
                      <span className="flex items-center text-slate-900 text-sm font-black bg-yellow-400 px-4 py-2 rounded-full shadow-md"><Zap className="w-4 h-4 mr-1 fill-slate-900"/> 적합 (최종 유효 데이터)</span> : 
                      <span className="flex items-center text-white text-sm font-black bg-red-600 px-4 py-2 rounded-full shadow-md"><AlertTriangle className="w-4 h-4 mr-1"/> 부적합 (재측정 요망)</span>
                    }
                  </div>
                )}
              </div>
            </div>
          </div>

          {/* 섹션 6: 시료(여과지) 분석 및 비고 */}
          <div className="bg-white p-6 rounded-xl shadow-sm border border-yellow-200">
            <h2 className="text-lg font-bold text-slate-800 mb-4 border-b border-yellow-100 pb-2 flex items-center gap-2">
              <span className="bg-slate-800 text-yellow-400 w-6 h-6 rounded-full flex items-center justify-center text-sm font-black">6</span>
              먼지 시료 무게 및 최종 결과
            </h2>
            
            <div className="grid grid-cols-1 lg:grid-cols-3 gap-6 mb-6">
              <div className="lg:col-span-2 grid grid-cols-3 gap-4 border border-yellow-200 p-4 rounded-xl bg-yellow-50/50 h-fit">
                <div><label className="block text-xs font-bold text-slate-700 mb-1">여과지 번호</label>
                  <input type="text" name="filterId" value={formData.filterId} onChange={handleChange} className="w-full p-2 border border-yellow-300 rounded-lg focus:ring-2 focus:ring-yellow-400 outline-none" /></div>
                <div><label className="block text-xs font-bold text-slate-700 mb-1">채취 전 무게 (mg)</label>
                  <input type="number" step="0.0001" name="filterInitial" value={formData.filterInitial} onChange={handleChange} placeholder="0.0000" className="w-full p-2 border border-yellow-300 rounded-lg focus:ring-2 focus:ring-yellow-400 outline-none" /></div>
                <div><label className="block text-xs font-bold text-slate-700 mb-1">채취 후 무게 (mg)</label>
                  <input type="number" step="0.0001" name="filterFinal" value={formData.filterFinal} onChange={handleChange} placeholder="0.0000" className="w-full p-2 border border-yellow-300 rounded-lg focus:ring-2 focus:ring-yellow-400 outline-none" /></div>
              </div>
              <div className="lg:col-span-1 bg-yellow-100 p-4 rounded-xl border border-yellow-300 flex flex-col justify-center items-center text-slate-900 shadow-sm relative overflow-hidden">
                <div className="absolute -right-4 -bottom-4 opacity-10">
                  <Zap className="w-24 h-24 fill-yellow-500" />
                </div>
                <Scale className="w-8 h-8 mb-2 text-yellow-600" />
                <span className="text-xs font-black mb-1 z-10">포집된 먼지 무게 (W<sub>d</sub>)</span>
                <span className="text-3xl font-black text-red-600 z-10">{calcDustWeightDiff()} <span className="text-sm font-bold text-slate-700">mg</span></span>
              </div>
            </div>

            {/* 최종 농도 산출 결과 패널 */}
            <div className="bg-slate-900 p-6 rounded-2xl border-2 border-yellow-400 shadow-xl text-white mb-6 relative overflow-hidden">
               <div className="absolute top-0 left-0 w-full h-2 bg-red-500"></div>
               <h3 className="text-base font-black text-yellow-400 mb-4 flex items-center gap-2 border-b border-slate-700 pb-3">
                 <FileSpreadsheet className="w-5 h-5"/> 최종 산출 결과 요약 ⚡
               </h3>
               <div className="grid grid-cols-2 md:grid-cols-4 gap-6 divide-y md:divide-y-0 md:divide-x divide-slate-700 text-center mt-2">
                  <div className="flex flex-col items-center justify-center p-2">
                      <span className="text-xs text-slate-400 mb-2 font-bold">표준 습가스 유량 (Q<sub>sw</sub>)</span>
                      <span className="text-2xl font-black text-white">{calcGasFlowRates().wet} <span className="text-sm font-medium text-slate-500">Sm³/hr</span></span>
                  </div>
                  <div className="flex flex-col items-center justify-center p-2">
                      <span className="text-xs text-slate-400 mb-2 font-bold">표준 건조가스 유량 (Q<sub>s</sub>)</span>
                      <span className="text-2xl font-black text-white">{calcGasFlowRates().dry} <span className="text-sm font-medium text-slate-500">Sm³/hr</span></span>
                  </div>
                  <div className="flex flex-col items-center justify-center p-2">
                      <span className="text-xs text-slate-400 mb-2 font-bold">실측 먼지 농도 (C)</span>
                      <span className="text-2xl font-black text-white">{calcDustConcentrations().actualC} <span className="text-sm font-medium text-slate-500">mg/Sm³</span></span>
                  </div>
                  <div className="flex flex-col items-center justify-center p-4 bg-slate-800 rounded-xl border border-slate-700 md:border-none md:rounded-none md:bg-transparent">
                      <span className="text-xs text-yellow-400 font-black mb-2">
                        O₂ 보정 농도 (C<sub>c</sub>)
                        {!formData.standardO2 && <span className="text-[10px] font-medium text-slate-500 ml-1">(제외)</span>}
                      </span>
                      <span className="text-3xl font-black text-red-400">{calcDustConcentrations().correctedC} <span className="text-sm font-medium text-slate-400">mg/Sm³</span></span>
                  </div>
               </div>
            </div>

            <div className="mt-2">
              <label className="block text-xs font-bold text-slate-700 mb-1">현장 특이사항 및 비고</label>
              <textarea name="remarks" value={formData.remarks} onChange={handleChange} rows="2" className="w-full p-3 border border-yellow-300 rounded-xl focus:ring-2 focus:ring-yellow-400 outline-none bg-yellow-50/30" placeholder="측정공 상태, 장비 특이점 등"></textarea>
            </div>
          </div>

          <div className="flex items-center justify-center gap-4 mt-8 pb-8 border-b-2 border-yellow-300 border-dashed">
            <button type="button" onClick={() => window.location.reload()} className="px-8 py-3.5 bg-white border-2 border-slate-300 text-slate-700 font-black rounded-xl hover:bg-slate-50 transition-colors flex items-center gap-2 shadow-sm">
              <RefreshCw className="w-5 h-5" /> 새 기록지 작성
            </button>
            <button type="submit" className="px-8 py-3.5 bg-red-500 text-white font-black rounded-xl hover:bg-red-600 transition-colors flex items-center gap-2 shadow-lg hover:-translate-y-0.5 transform">
              <Save className="w-5 h-5" /> 기록부 데이터 저장
            </button>
          </div>
        </form>

        {savedData.length > 0 && (
          <div className="bg-slate-900 p-6 rounded-2xl shadow-2xl mt-8 text-white border-2 border-slate-700">
            <div className="flex flex-col md:flex-row justify-between items-center mb-6 gap-4 border-b border-slate-700 pb-4">
              <h2 className="text-xl font-black flex items-center gap-2 text-yellow-400"><FileSpreadsheet className="w-6 h-6"/> 저장된 종합 리포트 ⚡</h2>
              <button 
                onClick={exportToCSV} 
                className="flex items-center gap-2 px-5 py-2.5 bg-yellow-400 hover:bg-yellow-500 text-slate-900 font-black rounded-xl transition-transform shadow-md text-sm hover:scale-105"
              >
                <Download className="w-4 h-4" /> CSV 파일로 추출
              </button>
            </div>
            <div className="overflow-x-auto rounded-xl border border-slate-700">
              <table className="w-full text-sm text-center">
                <thead className="bg-slate-800 text-slate-300">
                  <tr>
                    <th className="p-3 font-bold">배출구명</th>
                    <th className="p-3 font-bold">수분량(%)</th>
                    <th className="p-3 font-bold">등속흡인율(%)</th>
                    <th className="p-3 font-bold">먼지무게(mg)</th>
                    <th className="p-3 font-bold text-white">실측농도</th>
                    <th className="p-3 text-yellow-400 font-black">보정농도</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-700 bg-slate-900">
                  {savedData.map((data, index) => (
                    <tr key={index} className="hover:bg-slate-800/80 transition-colors">
                      <td className="p-3 font-bold text-slate-100">{data.location}</td>
                      <td className="p-3 text-slate-300">{data.moisturePercent}</td>
                      <td className="p-3">
                          <span className={parseFloat(data.isokineticRate) >= 95 && parseFloat(data.isokineticRate) <= 105 ? "text-slate-900 font-black bg-yellow-400 px-2.5 py-1 rounded-md" : "text-white font-black bg-red-500 px-2.5 py-1 rounded-md"}>
                           {data.isokineticRate}
                          </span>
                      </td>
                      <td className="p-3 text-slate-300">{data.dustWeight}</td>
                      <td className="p-3 font-bold text-white">{data.actualConcentration}</td>
                      <td className="p-3 text-yellow-400 font-black text-lg">{data.correctedConcentration}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        )}

      </div>
    </div>
  );
}