import React, { useEffect, useRef, useState, Fragment } from 'react';
import { createPortal } from 'react-dom';
import mainLogoImage from './assets/mainlogo.svg';
import mainLogoClayImage from './assets/mainlogo-clay.svg';
import JSZip from 'jszip';
import { recalcWorkbookCaches } from './xlsxRecalc.js';
import { tidyWorkbook, rowsAt } from './xlsxTidy.js';
import { fetchCityObservations as webFetchCityObservations, searchPlaces as webSearchPlaces, fetchLocalWeather as webFetchLocalWeather } from '../electron/weather.mjs';
import { buildRecordSheets, buildSamplingLogSheets, buildPresetSheets, buildWorkbook, buildStackSectionSvg, renderSvgToPng } from './openWorkbook.js';

// 배포용(데스크톱 exe) 빌드 여부. `vite build --mode desktop` 에서 .env.desktop 으로 켜지며,
// 빌드 시점 상수라서 아래 분기의 회사 장비 값은 배포용 번들에 포함되지 않습니다.
const IS_DESKTOP_BUILD = import.meta.env.VITE_DESKTOP_BUILD === '1';

const DEFAULT_NOZZLE_SET = IS_DESKTOP_BUILD ? [] : [
  { num: 4, d: 3.21 }, { num: 5, d: 3.97 }, { num: 6, d: 4.79 },
  { num: 7, d: 5.54 }, { num: 8, d: 6.28 }, { num: 9, d: 7.10 },
  { num: 10, d: 7.86 }, { num: 11, d: 8.40 }, { num: 12, d: 9.24 },
  { num: 13, d: 10.18 }, { num: 14, d: 10.87 }, { num: 15, d: 11.70 },
  { num: 16, d: 12.56 }, { num: 17, d: 13.48 }, { num: 18, d: 14.30 },
  { num: 20, d: 15.95 }, { num: 22, d: 17.51 }, { num: 26, d: 20.72 }
];

// 수은·PAHs는 노즐 세트가 따로 있음 (각 양식 노즐산정 B39:C 목록)
const DEFAULT_GROUP_NOZZLE_SETS = IS_DESKTOP_BUILD ? { mercury: [], pahs: [] } : {
  mercury: [
    { num: 1, d: 4.10 }, { num: 2, d: 4.69 }, { num: 3, d: 5.10 }, { num: 4, d: 5.69 }, { num: 5, d: 6.05 },
    { num: 6, d: 6.72 }, { num: 7, d: 7.14 }, { num: 8, d: 7.69 }, { num: 9, d: 8.08 }, { num: 10, d: 8.68 },
    { num: 13, d: 10.00 }, { num: 15, d: 11.03 }, { num: 17, d: 12.12 }, { num: 18, d: 12.51 }, { num: 19, d: 13.10 },
  ],
  pahs: [
    { num: 1, d: 4.10 }, { num: 2, d: 4.39 }, { num: 3, d: 5.09 }, { num: 4, d: 5.68 }, { num: 5, d: 6.09 },
    { num: 6, d: 6.74 }, { num: 7, d: 7.08 }, { num: 8, d: 7.61 }, { num: 9, d: 8.13 }, { num: 10, d: 8.65 },
    { num: 11, d: 9.06 }, { num: 12, d: 9.70 }, { num: 13, d: 10.09 }, { num: 14, d: 10.68 }, { num: 15, d: 11.05 },
    { num: 16, d: 11.44 }, { num: 17, d: 12.16 }, { num: 18, d: 12.47 }, { num: 19, d: 13.12 },
  ],
};
// 기록부 → 노즐 세트 (먼지·금속·유량은 같은 세트)
const getNozzleGroupForSheet = (sheetId) => (sheetId === 'mercury' || sheetId === 'pahs' ? sheetId : 'dust');
const NOZZLE_GROUP_LABELS = { dust: '먼지·금속·유량', mercury: '수은', pahs: 'PAHs' };
const cloneDefaultGroupNozzleSets = () => ({
  mercury: DEFAULT_GROUP_NOZZLE_SETS.mercury.map((item) => ({ ...item })),
  pahs: DEFAULT_GROUP_NOZZLE_SETS.pahs.map((item) => ({ ...item })),
});

const DEFAULT_SAMPLERS = IS_DESKTOP_BUILD ? [] : [
  { id: 1, name: '샘플러 1호기 (1910162K)', yd: '0.991', dHAt: '47.6079' },
  { id: 2, name: '샘플러 2호기 (1910163K)', yd: '0.994', dHAt: '47.4726' },
  { id: 3, name: '샘플러 3호기 (1910164K)', yd: '1.0078', dHAt: '47.557' },
  { id: 4, name: '샘플러 4호기 (1910165K)', yd: '0.9976', dHAt: '47.6512' },
  { id: 5, name: '샘플러 5호기 (A2312850)', yd: '1', dHAt: '49.1' },
];

// 피토관 교정 성적서 (기록부 양식의 '피토' 시트 기준, 교정일 2026-01): 유속(m/s)별 Cp
const DEFAULT_PITOT_PRESETS = IS_DESKTOP_BUILD ? [] : [
  { id: 1, sn: 'YNJ0412', usage: '입자', length: '65', calDate: '2026-01-20', factor: '0.822', cpByVelocity: [{ v: '4.95', cp: '0.816' }, { v: '6.92', cp: '0.821' }, { v: '8.91', cp: '0.822' }, { v: '11.92', cp: '0.822' }, { v: '14.93', cp: '0.823' }, { v: '19.95', cp: '0.824' }, { v: '29.94', cp: '0.823' }, { v: '37.39', cp: '0.824' }] },
  { id: 2, sn: 'GNT-TECH 525', usage: '입자', length: '125', calDate: '2026-01-20', factor: '0.822', cpByVelocity: [{ v: '5.13', cp: '0.825' }, { v: '7.04', cp: '0.824' }, { v: '8.96', cp: '0.823' }, { v: '11.88', cp: '0.822' }, { v: '14.84', cp: '0.821' }, { v: '19.82', cp: '0.822' }, { v: '29.86', cp: '0.823' }, { v: '37.49', cp: '0.823' }] },
  { id: 3, sn: 'GNT-TECH 528', usage: '입자', length: '125', calDate: '2026-01-20', factor: '0.824', cpByVelocity: [{ v: '5.13', cp: '0.827' }, { v: '7.04', cp: '0.827' }, { v: '8.96', cp: '0.826' }, { v: '11.88', cp: '0.824' }, { v: '14.84', cp: '0.824' }, { v: '19.82', cp: '0.823' }, { v: '29.86', cp: '0.822' }, { v: '37.49', cp: '0.821' }] },
  { id: 4, sn: 'GNT-TECH 621', usage: '입자', length: '125', calDate: '2026-01-20', factor: '0.823', cpByVelocity: [{ v: '5.13', cp: '0.827' }, { v: '7.03', cp: '0.825' }, { v: '8.96', cp: '0.824' }, { v: '11.88', cp: '0.823' }, { v: '14.85', cp: '0.822' }, { v: '19.81', cp: '0.822' }, { v: '29.86', cp: '0.823' }, { v: '37.51', cp: '0.823' }] },
  { id: 5, sn: 'YNJ0022', usage: '입자', length: '230', calDate: '2026-01-20', factor: '0.811', cpByVelocity: [{ v: '4.96', cp: '0.807' }, { v: '6.91', cp: '0.809' }, { v: '8.89', cp: '0.810' }, { v: '11.9', cp: '0.811' }, { v: '14.9', cp: '0.813' }, { v: '19.92', cp: '0.814' }, { v: '29.91', cp: '0.814' }, { v: '37.39', cp: '0.814' }] },
  { id: 6, sn: 'YNJ0854', usage: '입자', length: '235', calDate: '2026-01-21', factor: '0.820', cpByVelocity: [{ v: '4.96', cp: '0.814' }, { v: '6.92', cp: '0.818' }, { v: '8.92', cp: '0.820' }, { v: '11.92', cp: '0.820' }, { v: '14.92', cp: '0.820' }, { v: '19.94', cp: '0.820' }, { v: '29.92', cp: '0.819' }, { v: '37.41', cp: '0.819' }] },
  { id: 7, sn: 'YNJ0046', usage: '입자', length: '295', calDate: '2026-01-20', factor: '0.826', cpByVelocity: [{ v: '4.96', cp: '0.815' }, { v: '6.92', cp: '0.822' }, { v: '8.9', cp: '0.825' }, { v: '11.91', cp: '0.826' }, { v: '14.92', cp: '0.825' }, { v: '19.92', cp: '0.825' }, { v: '29.94', cp: '0.828' }, { v: '37.42', cp: '0.828' }] },
  { id: 8, sn: 'ENK2404', usage: '수은', length: '55', calDate: '2026-01-20', factor: '0.826', cpByVelocity: [{ v: '5.13', cp: '0.826' }, { v: '7.03', cp: '0.826' }, { v: '8.96', cp: '0.826' }, { v: '11.88', cp: '0.826' }, { v: '14.85', cp: '0.825' }, { v: '19.82', cp: '0.825' }, { v: '29.87', cp: '0.824' }, { v: '37.51', cp: '0.823' }] },
  { id: 9, sn: 'GNT-TECH 1440', usage: '수은', length: '115', calDate: '2026-01-20', factor: '0.819', cpByVelocity: [{ v: '5.14', cp: '0.826' }, { v: '7.04', cp: '0.824' }, { v: '8.96', cp: '0.822' }, { v: '11.89', cp: '0.819' }, { v: '14.86', cp: '0.818' }, { v: '19.84', cp: '0.818' }, { v: '29.9', cp: '0.818' }, { v: '37.53', cp: '0.817' }] },
  { id: 10, sn: 'GNT-TECH 1444', usage: '수은', length: '115', calDate: '2026-01-20', factor: '0.822', cpByVelocity: [{ v: '5.14', cp: '0.825' }, { v: '7.03', cp: '0.824' }, { v: '8.96', cp: '0.823' }, { v: '11.88', cp: '0.822' }, { v: '14.85', cp: '0.822' }, { v: '19.83', cp: '0.821' }, { v: '29.87', cp: '0.820' }, { v: '37.52', cp: '0.820' }] },
  { id: 11, sn: 'ENK2407', usage: '수은', length: '215', calDate: '2026-01-20', factor: '0.827', cpByVelocity: [{ v: '5.13', cp: '0.828' }, { v: '7.03', cp: '0.828' }, { v: '8.95', cp: '0.827' }, { v: '11.88', cp: '0.827' }, { v: '14.84', cp: '0.826' }, { v: '19.82', cp: '0.825' }, { v: '29.87', cp: '0.825' }, { v: '37.5', cp: '0.824' }] },
  { id: 12, sn: 'YNJ0557', usage: '수은', length: '315', calDate: '2026-01-20', factor: '0.823', cpByVelocity: [{ v: '4.96', cp: '0.812' }, { v: '6.93', cp: '0.820' }, { v: '8.9', cp: '0.822' }, { v: '11.9', cp: '0.823' }, { v: '14.92', cp: '0.825' }, { v: '19.95', cp: '0.825' }, { v: '29.92', cp: '0.827' }, { v: '37.41', cp: '0.829' }] },
  { id: 13, sn: 'YNJ1336', usage: 'PAHs', length: '120', calDate: '2026-01-21', factor: '0.817', cpByVelocity: [{ v: '4.96', cp: '0.810' }, { v: '6.93', cp: '0.815' }, { v: '8.92', cp: '0.817' }, { v: '11.92', cp: '0.817' }, { v: '14.91', cp: '0.816' }, { v: '19.94', cp: '0.817' }, { v: '29.92', cp: '0.816' }, { v: '37.43', cp: '0.817' }] },
  { id: 14, sn: 'YNJ1319', usage: 'PAHs', length: '170', calDate: '2026-01-21', factor: '0.820', cpByVelocity: [{ v: '4.96', cp: '0.813' }, { v: '6.92', cp: '0.818' }, { v: '8.92', cp: '0.820' }, { v: '11.91', cp: '0.820' }, { v: '14.92', cp: '0.821' }, { v: '19.95', cp: '0.821' }, { v: '29.95', cp: '0.821' }, { v: '37.42', cp: '0.821' }] },
  { id: 15, sn: 'YNJ1374', usage: 'PAHs', length: '315', calDate: '2026-01-21', factor: '0.818', cpByVelocity: [{ v: '4.96', cp: '0.808' }, { v: '6.91', cp: '0.814' }, { v: '8.91', cp: '0.817' }, { v: '11.9', cp: '0.818' }, { v: '14.92', cp: '0.818' }, { v: '19.93', cp: '0.818' }, { v: '29.89', cp: '0.817' }, { v: '37.41', cp: '0.819' }] },
];

// 기록부 종류별 엑셀 양식 레이아웃: 기록지 측정행 범위, 임핀저 단수, 노즐 목록 행수, 측정번호 접두어
const EXPORT_TEMPLATE_LAYOUTS = {
  dust: { file: 'dust-template.xlsm', meterStartRow: 17, meterTemplateEndRow: 27, impingerRows: 4, nozzleRows: 18, filterPrefix: '먼지', leak: { firstRow: 14, hasMeter: true } },
  metal: { file: 'dust-template.xlsm', meterStartRow: 17, meterTemplateEndRow: 27, impingerRows: 4, nozzleRows: 18, filterPrefix: '금속', leak: { firstRow: 14, hasMeter: true } },
  mercury: { file: 'mercury-template.xlsm', meterStartRow: 17, meterTemplateEndRow: 27, impingerRows: 4, nozzleRows: 15, filterPrefix: '수은', leak: { firstRow: 14, hasMeter: true } },
  pahs: { file: 'pahs-template.xlsm', meterStartRow: 17, meterTemplateEndRow: 24, impingerRows: 6, nozzleRows: 19, filterPrefix: 'PAHs', leak: { firstRow: 14, hasMeter: true } },
  flow: { file: 'flow-template.xlsm', meterStartRow: 17, meterTemplateEndRow: 25, impingerRows: 4, nozzleRows: 18, filterPrefix: '유량', leak: null, hideSheets: ['누출시험'] }, // 유량 측정에는 누출시험이 없다 → 화면에서 빼고 엑셀 시트는 숨김
};
// 양식에서 주변 칸은 가운데 정렬인데 혼자 왼쪽·오른쪽으로 붙어 있던 표 칸들. 내보낼 때 가운데로 맞춘다
const EXPORT_CENTER_CELLS = (layout) => ({
  수분량: ['A5:B7', 'A10:B15', 'C19', 'C32'],
  노즐산정: ['D', 'F', 'H', 'I', 'J'].map((col) => `${col}39:${col}${38 + (layout.nozzleRows || 18)}`),
});

// 누출시험 기록 (양식 '누출시험' 시트): 구분별 유량계 멈춤 여부(O/X)와 가스미터 전·후 지시값
const LEAK_TEST_PARTS = [
  { key: 'device', label: '시료채취장치' },
  { key: 'impinger', label: '임핀저 트레인' },
  { key: 'samplingPart', label: '시료채취부' },
];
const createEmptyLeakTest = () => ({ device: 'O', impinger: 'O', samplingPart: 'O', meterBefore: '', meterAfter: '', note: '' });
const sanitizeLeakTest = (value) => {
  const base = createEmptyLeakTest();
  if (!value || typeof value !== 'object') return base;
  const ox = (v) => (String(v || '').trim().toUpperCase() === 'X' ? 'X' : 'O');
  const text = (v) => (v === undefined || v === null ? '' : String(v));
  return { device: ox(value.device), impinger: ox(value.impinger), samplingPart: ox(value.samplingPart), meterBefore: text(value.meterBefore), meterAfter: text(value.meterAfter), note: text(value.note) };

};
const SAMPLING_LOG_TEMPLATE_FILE = 'samplinglog-template.xlsx';
const SHEET_SHORT_LABELS = { dust: '먼지', metal: '금속', mercury: '수은', pahs: 'PAHs', flow: '유량', samplinglog: '종합' };
const SHEET_BADGE_COLORS = { dust: '#16a34a', metal: '#ca8a04', mercury: '#2563eb', pahs: '#9333ea', flow: '#0891b2', samplinglog: '#ea580c' };
// 기록부 종류 아이콘 배지: 종류별 색 + 짧은 이름
const SheetBadge = ({ sheetId, title }) => (
  <span
    className="inline-flex items-center justify-center min-w-[2.6rem] px-1.5 py-0.5 rounded-md text-[11px] font-bold text-white whitespace-nowrap shadow-sm"
    style={{ backgroundColor: SHEET_BADGE_COLORS[sheetId] || '#475569', boxShadow: 'inset 0 1px 0 rgba(255,255,255,0.25)' }}
    title={title || ''}
  >
    {SHEET_SHORT_LABELS[sheetId] || title || '-'}
  </span>
);

const SHEET_MENU = [
  {
    id: 'dust',
    title: '먼지시료채취기록부',
    desc: '먼지시료채취 및 등속흡인 기록',
    toneClass: 'from-emerald-200/60 to-teal-200/20',
    swatchClass: 'bg-emerald-500',
    hoverTextClass: 'group-hover:text-emerald-700',
  },
  {
    id: 'metal',
    title: '금속시료채취기록부',
    desc: '금속시료채취 및 등속흡인 기록',
    toneClass: 'from-yellow-200/70 to-amber-200/20',
    swatchClass: 'bg-yellow-400',
    hoverTextClass: 'group-hover:text-yellow-700',
  },
  {
    id: 'mercury',
    title: '수은시료채취기록부',
    desc: '수은시료채취 및 등속흡인 기록',
    toneClass: 'from-blue-200/70 to-sky-200/20',
    swatchClass: 'bg-blue-500',
    hoverTextClass: 'group-hover:text-blue-700',
  },
  {
    id: 'pahs',
    title: 'PAHs시료채취기록부',
    desc: 'PAHs시료채취 및 등속흡인 기록',
    toneClass: 'from-purple-200/70 to-violet-200/20',
    swatchClass: 'bg-purple-500',
    hoverTextClass: 'group-hover:text-purple-700',
  },
  {
    id: 'flow',
    title: '유량측정기록부',
    desc: '배출가스 유속·유량 측정 기록',
    toneClass: 'from-cyan-200/70 to-sky-200/20',
    swatchClass: 'bg-cyan-500',
    hoverTextClass: 'group-hover:text-cyan-700',
  },
  {
    id: 'samplinglog',
    title: '통합시료채취기록부',
    desc: '사업장별 물질 채취 현황 (가스·입자·VOCs)',
    toneClass: 'from-orange-200/70 to-amber-200/20',
    swatchClass: 'bg-orange-500',
    hoverTextClass: 'group-hover:text-orange-700',
  },
];

// 통합시료채취기록부: 가스 샘플러 목록 (양식 '시료채취 작성' S6:T10 기준)
const DEFAULT_GAS_SAMPLERS = IS_DESKTOP_BUILD ? [] : [
  { id: 1, sn: '1910122', yd: '0.9857' },
  { id: 2, sn: '1910123', yd: '1.0036' },
  { id: 3, sn: '1910124', yd: '0.9985' },
  { id: 4, sn: '1910125', yd: '0.9892' },
  { id: 5, sn: 'A2304264', yd: '1' },
];

// 통합시료채취기록부 물질 목록: 양식 '시료채취 작성' 13~57행 순서 그대로
// kind: particle(입자 샘플러) / gas(가스 샘플러) / analyzer(가스분석기) / thc(총탄화수소측정기) / box(음압상자) / plain(시작·종료값 차이)
const SAMPLING_LOG_SUBSTANCES = [
  { row: 13, code: 'P', name: '먼지', kind: 'particle', targetVolume: '1000', testPeriod: '2' },
  { row: 14, code: 'G', name: '황산화물', kind: 'analyzer', duration: '15', testPeriod: '현장측정' },
  { row: 15, code: 'G', name: '질소산화물', kind: 'analyzer', duration: '15', testPeriod: '현장측정' },
  { row: 16, code: 'G', name: '일산화탄소', kind: 'analyzer', duration: '15', testPeriod: '현장측정' },
  { row: 17, code: 'NH3', name: '암모니아-1(BA)', species: '암모니아', part: 'BA 흡수액', kind: 'gas', suffix: 'BA-1', suctionRate: '1~2', targetVolume: '20', duration: '20', testPeriod: '3' },
  { row: 18, code: 'NH3', name: '암모니아-2(H2O2)', species: '암모니아', part: 'H₂O₂ 흡수액', kind: 'gas', suffix: 'H2O2-1', suctionRate: '1~2', targetVolume: '20', duration: '20', testPeriod: '3' },
  { row: 19, code: 'F', name: '불소', kind: 'gas', suctionRate: '1', targetVolume: '40', duration: '40', testPeriod: '4' },
  { row: 20, code: 'HCl', name: '염화수소', kind: 'gas', suctionRate: '1', targetVolume: '20', duration: '20', testPeriod: '3' },
  { row: 21, code: 'H2S', name: '황화수소', kind: 'gas', suctionRate: '1', targetVolume: '20', duration: '20', testPeriod: '4' },
  { row: 22, code: 'HCN', name: '시안화수소', kind: 'gas', suctionRate: '1', targetVolume: '10', duration: '10', testPeriod: '3' },
  { row: 23, code: 'HDZ', name: '하이드라진', kind: 'gas', suctionRate: '1', targetVolume: '10', duration: '10', testPeriod: '3' },
  { row: 24, code: 'CS2', name: '이황화탄소', kind: 'gas', testPeriod: '9' },
  { row: 25, code: 'M', name: '구리', kind: 'particle', targetVolume: '1000', testPeriod: '6' },
  { row: 26, code: 'M', name: '카드뮴', kind: 'particle', testPeriod: '6' },
  { row: 27, code: 'M', name: '납', kind: 'particle', testPeriod: '6' },
  { row: 28, code: 'M', name: '크롬', kind: 'particle', testPeriod: '6' },
  { row: 29, code: 'M', name: '니켈', kind: 'particle', testPeriod: '6' },
  { row: 30, code: 'M', name: '아연', kind: 'particle', testPeriod: '6' },
  { row: 31, code: 'M', name: '베릴륨', kind: 'particle', testPeriod: '6' },
  { row: 32, code: 'M', name: '비소(입자상)', species: '비소', part: '입자상', kind: 'particle', targetVolume: '1000', testPeriod: '9' },
  { row: 33, code: 'As-G', name: '비소(가스상)', species: '비소', part: '가스상', kind: 'particle', suctionRate: '1', duration: '입자상 동일' },
  { row: 34, code: 'Hg', name: '수은', kind: 'particle', targetVolume: '1000', testPeriod: '9' },
  { row: 35, code: 'ALD-1', name: '알데히드-1', species: '포름알데하이드 · 아세트알데하이드', options: ['포름알데하이드', '아세트알데하이드'], part: '1회차', kind: 'gas', suctionRate: '1', targetVolume: '10', testPeriod: '7' },
  { row: 36, code: 'ALD-2', name: '알데히드-2', species: '포름알데하이드 · 아세트알데하이드', options: ['포름알데하이드', '아세트알데하이드'], part: '2회차', kind: 'gas', suctionRate: '1', targetVolume: '10' },
  { row: 37, code: 'ALD-3', name: '알데히드-3', species: '포름알데하이드 · 아세트알데하이드', options: ['포름알데하이드', '아세트알데하이드'], part: '3회차', kind: 'gas', suctionRate: '1', targetVolume: '10' },
  { row: 38, code: 'PN', name: '페놀', kind: 'gas', suctionRate: '1', targetVolume: '20', testPeriod: '4' },
  { row: 39, code: 'Br', name: '브롬', kind: 'gas', suctionRate: '1', targetVolume: '40', testPeriod: '4' },
  { row: 40, code: 'THC', name: '총탄화수소', kind: 'thc', duration: '30', testPeriod: '현장측정' },
  { row: 41, code: 'BZ', name: '벤조(a)피렌', kind: 'plain', testPeriod: '9' },
  { row: 42, code: 'VOCs', name: '염화비닐', kind: 'box', testPeriod: '9' },
  { row: 43, code: 'VOCs', name: '벤젠', kind: 'box', duration: '15', testPeriod: '9' },
  { row: 44, code: 'VOCs', name: '디클로로메탄', kind: 'box', testPeriod: '9' },
  { row: 45, code: 'VOCs', name: '트리클로로에틸렌', kind: 'box', testPeriod: '9' },
  { row: 46, code: 'VOCs', name: '1,3-부타디엔', kind: 'box', testPeriod: '9' },
  { row: 47, code: 'VOCs', name: '테트라클로로에틸렌', kind: 'box', testPeriod: '9' },
  { row: 48, code: 'VOCs', name: '1,2-디클로로에탄', kind: 'box', testPeriod: '9' },
  { row: 49, code: 'VOCs', name: '클로로포름', kind: 'box', testPeriod: '9' },
  { row: 50, code: 'VOCs', name: '아크릴로니트릴', kind: 'box', testPeriod: '9' },
  { row: 51, code: 'VOCs', name: '스틸렌', kind: 'box', testPeriod: '9' },
  { row: 52, code: 'VOCs', name: '에틸벤젠', kind: 'box', testPeriod: '9' },
  { row: 53, code: 'VOCs', name: '사염화탄소', kind: 'box', testPeriod: '9' },
  { row: 54, code: 'BEN', name: '벤지딘', kind: 'plain' },
  { row: 55, code: 'AN', name: '아닐린', kind: 'plain' },
  { row: 56, code: 'DMS', name: '이황화메틸', kind: 'box' },
  { row: 57, code: 'OX', name: '옥사이드', species: '에틸렌옥사이드 · 프로필렌옥사이드', options: ['에틸렌옥사이드', '프로필렌옥사이드'], kind: 'box', testPeriod: '9' },
];

// 기록부(먼지·금속·수은·PAHs) 저장본을 통합시료채취기록부 행에 옮길 때의 대응표
const SHEET_TO_LOG_ROWS = { dust: [13], metal: [25, 26, 27, 28, 29, 30, 31, 32], mercury: [34], pahs: [41] };
const LOG_SOURCE_SHEETS = ['dust', 'metal', 'mercury', 'pahs', 'flow'];

const addMinutesToClock = (hhmm, minutes) => {
  const m = /^(\d{1,2}):(\d{2})$/.exec(String(hhmm || '').trim());
  if (!m || !Number.isFinite(minutes)) return '';
  const total = ((parseInt(m[1], 10) * 60 + parseInt(m[2], 10) + Math.round(minutes)) % (24 * 60) + 24 * 60) % (24 * 60);
  return `${String(Math.floor(total / 60)).padStart(2, '0')}:${String(total % 60).padStart(2, '0')}`;
};

// 저장된 기록부 한 건에서 통합시료채취기록부 행에 넣을 값(시각, 가스미터 시작·종료, 미터 온도, 오리피스압)을 뽑는다
const buildSamplingLogValuesFromReport = (report) => {
  const meters = Array.isArray(report?.gasMeters) ? report.gasMeters : [];
  const volumes = meters.map((m) => parseFloat(m?.volume)).filter((v) => Number.isFinite(v));
  const fmt = (v) => (Number.isFinite(v) ? String(Math.round(v * 100) / 100) : '');
  const clean = (v) => (v === undefined || v === null || v === '-' ? '' : String(v));
  const totalMinutes = parseFloat(report?.totalSamplingMinutes);
  const start = clean(report?.samplingStartTime);
  const end = start && Number.isFinite(totalMinutes) && totalMinutes > 0 ? addMinutesToClock(start, totalMinutes) : '';
  return {
    start,
    end,
    meterStart: volumes.length > 0 ? fmt(volumes[0]) : '',
    meterEnd: volumes.length > 1 ? fmt(volumes[volumes.length - 1]) : '',
    temp: clean(report?.avgTm),
    pressureH2O: clean(report?.avgOrifice),
  };
};

// 같은 사업장·배출구(스택)의 기록부를 한 묶음으로 묶되, 날짜가 이어지는 것(1·2·3일차처럼 사이가 3일 이내)만 같은 측정 건으로 본다.
// 날짜가 그보다 떨어져 있으면 같은 스택이라도 별도 묶음. (기록부 종류별로는 최신 저장본 하나를 대표로 사용)
const CAMPAIGN_MAX_GAP_DAYS = 3;
const formatDateRange = (dates) => {
  const list = (dates || []).filter(Boolean);
  if (list.length === 0) return '-';
  if (list.length === 1) return list[0];
  return `${list[0]} ~ ${list[list.length - 1]} (${list.length}일)`;
};
const daysBetween = (a, b) => {
  const t1 = Date.parse(a);
  const t2 = Date.parse(b);
  if (!Number.isFinite(t1) || !Number.isFinite(t2)) return Infinity;
  return Math.abs(t2 - t1) / 86400000;
};
// 사업장·배출구 이름 비교용: 앞뒤·중간 공백과 대소문자 차이는 같은 곳으로 본다
const normalizeSiteName = (value) => String(value || '').trim().replace(/\s+/g, ' ').toLowerCase();
// 통합시료채취기록부 일차 칸 수
const SAMPLING_LOG_DAY_SLOTS = 4;
// 통합시료채취기록부 엑셀: 날짜별 시트 열 → 입력 시트('시료채취 작성') 열
const DAY_SHEET_INPUT_COLUMNS = { B: 'D', C: 'E', D: 'G', E: 'H', F: 'I', G: 'J', H: 'K', I: 'L', J: 'M', K: 'N' };
// includeBlank: 사업장·배출구·날짜를 하나도 안 적은 기록도 "정보 없음" 묶음으로 보여 준다 (저장 목록용)
const groupSheetReportsForLog = (reports, resolveSheetId, sheetFilter = LOG_SOURCE_SHEETS, { includeBlank = false } = {}) => {
  const sites = new Map();
  (reports || []).forEach((report) => {
    const sheetId = resolveSheetId(report);
    if (sheetFilter && !sheetFilter.includes(sheetId)) return;
    const company = String(report?.company || '').trim();
    const location = String(report?.location || report?.stackName || '').trim();
    const date = String(report?.date || '').trim();
    if (!company && !location && !date && !includeBlank) return;
    const key = `${normalizeSiteName(company)}|${normalizeSiteName(location)}`;
    if (!sites.has(key)) sites.set(key, { company, location, items: [] });
    sites.get(key).items.push({ report, sheetId, date });
  });

  const groups = [];
  sites.forEach(({ company, location, items }) => {
    // 날짜순으로 늘어놓고, 사이가 CAMPAIGN_MAX_GAP_DAYS보다 벌어지면 새 묶음
    const dated = items.filter((it) => it.date).sort((x, y) => x.date.localeCompare(y.date));
    const undated = items.filter((it) => !it.date);
    const clusters = [];
    dated.forEach((it) => {
      const last = clusters[clusters.length - 1];
      const fitsLast = last
        && daysBetween(last.dates[last.dates.length - 1], it.date) <= CAMPAIGN_MAX_GAP_DAYS
        && (last.dates.includes(it.date) || last.dates.length < SAMPLING_LOG_DAY_SLOTS);
      if (fitsLast) {
        last.items.push(it);
        if (!last.dates.includes(it.date)) last.dates.push(it.date);
      } else {
        clusters.push({ items: [it], dates: [it.date] });
      }
    });
    if (undated.length > 0) clusters.push({ items: undated, dates: [] });
    clusters.forEach((cluster) => {
      const group = { key: `${company}|${location}|${cluster.dates[0] || 'nodate'}`, company, location, dates: cluster.dates, bySheet: {}, reports: [] };
      cluster.items.forEach(({ report, sheetId }) => {
        group.reports.push(report);
        const current = group.bySheet[sheetId];
        if (!current || String(report.savedAt || '') > String(current.savedAt || '')) group.bySheet[sheetId] = report;
      });
      group.date = group.dates[group.dates.length - 1] || '';
      group.dateLabel = formatDateRange(group.dates);
      group.reports.sort((x, y) => String(y.savedAt || '').localeCompare(String(x.savedAt || '')));
      groups.push(group);
    });
  });
  return groups.sort((x, y) => (y.date || '').localeCompare(x.date || '') || x.company.localeCompare(y.company));
};

// 양식에서 의뢰 칸(C열)이 윗행과 병합된 행: 15, 18, 33, 36, 37
// 양식에서 의뢰 칸(C열)이 윗행과 병합된 행: 18(암모니아), 33(비소), 36·37(알데히드). 황산화물·질소산화물(14·15)은 별개 종이라 추출 시 병합을 풀고 각각 O 표시
const SAMPLING_LOG_MERGED_CHECK_ROWS = new Set([18, 33, 36, 37]);
const SAMPLING_LOG_UNMERGE_ON_EXPORT = ['C14:C15'];

// 물질 표 그룹 (표시용): 입자상 / 가스상 / 현장측정 / VOCs·기타
const SAMPLING_LOG_GROUPS = [
  { id: 'particle', title: '입자상 물질', note: '여과지 포집 · 입자 샘플러', kinds: ['particle'], color: 'emerald' },
  { id: 'gas', title: '가스상 물질', note: '흡수액 포집 · 가스 샘플러', kinds: ['gas'], color: 'sky' },
  { id: 'field', title: '현장측정', note: '가스분석기 · 총탄화수소측정기', kinds: ['analyzer', 'thc'], color: 'slate' },
  { id: 'vocs', title: 'VOCs · 기타', note: '흡착관 · 음압상자', kinds: ['box', 'plain'], color: 'violet' },
];
const SAMPLING_LOG_KIND_BADGE = {
  particle: 'bg-emerald-100 text-emerald-800 border-emerald-200',
  gas: 'bg-sky-100 text-sky-800 border-sky-200',
  analyzer: 'bg-slate-100 text-slate-700 border-slate-200',
  thc: 'bg-slate-100 text-slate-700 border-slate-200',
  box: 'bg-violet-100 text-violet-800 border-violet-200',
  plain: 'bg-amber-100 text-amber-800 border-amber-200',
};
const SAMPLING_LOG_GROUP_HEADER = {
  emerald: 'bg-emerald-50 text-emerald-900 border-emerald-200',
  sky: 'bg-sky-50 text-sky-900 border-sky-200',
  slate: 'bg-slate-100 text-slate-800 border-slate-200',
  violet: 'bg-violet-50 text-violet-900 border-violet-200',
};
// 종(種) 단위: 같은 측정 건을 여러 행으로 적는 것(암모니아 흡수액 2종, 비소 입자상·가스상, 알데히드 3회차)은 한 종으로 센다.
// 옥사이드 한 행은 에틸렌옥사이드·프로필렌옥사이드 2종, 알데히드는 포름·아세트 2종 → 전체 43종
const getSubstanceSpeciesKey = (sub) => sub.species || sub.name;
const getSubstanceSpeciesCount = (sub, row = null) => {
  if (!Array.isArray(sub.options)) return 1;
  if (row && Array.isArray(row.speciesSel)) return Math.max(1, row.speciesSel.length);
  return sub.options.length;
};
const countSpecies = (subs, log = null) => {
  const seen = new Set();
  let total = 0;
  subs.forEach((sub) => {
    const key = getSubstanceSpeciesKey(sub);
    if (seen.has(key)) return;
    seen.add(key);
    total += getSubstanceSpeciesCount(sub, log ? log.rows[String(sub.row)] : null);
  });
  return total;
};
const SAMPLING_LOG_SPECIES_TOTAL = countSpecies(SAMPLING_LOG_SUBSTANCES);
const getSamplingLogRowsOfSpecies = (sub) => SAMPLING_LOG_SUBSTANCES.filter((item) => getSubstanceSpeciesKey(item) === getSubstanceSpeciesKey(sub));

const getSamplingLogStats = (log) => {
  const requested = SAMPLING_LOG_SUBSTANCES.filter((sub) => log.rows[String(sub.row)]?.requested);
  const filled = requested.filter((sub) => {
    const row = log.rows[String(sub.row)];
    return row && row.start && row.end;
  });
  return { requestedCount: countSpecies(requested, log), requestedRows: requested.length, filledCount: filled.length, speciesTotal: SAMPLING_LOG_SPECIES_TOTAL };
};

const SAMPLING_LOG_KIND_LABEL = {
  particle: '입자 샘플러', gas: '가스 샘플러', analyzer: '가스분석기', thc: '총탄화수소측정기', box: '음압상자', plain: '-',
};

const createEmptySamplingLogRow = (substance) => ({
  requested: false,
  dayIndex: 0,
  start: '', end: '', meterStart: '', meterEnd: '', temp: '', pressureH2O: '',
  suctionRate: substance.suctionRate || '',
  targetVolume: substance.targetVolume || '',
  duration: substance.duration || '',
  testPeriod: substance.testPeriod || '',
  // 한 행에 두 종이 묶인 경우(알데히드, 옥사이드) 어느 종을 의뢰하는지
  speciesSel: Array.isArray(substance.options) ? [...substance.options] : undefined,
});

const createEmptySamplingLog = () => ({
  company: '', industry: '', facility: '', stackName: '', siteAbbr: '',
  particleSamplerId: '', particleSamplerYd: '', gasSamplerId: '', gasSamplerYd: '',
  // Yd 칸을 직접 고쳤는지: true면 프리셋을 고쳐도 그대로 둔다
  particleSamplerYdManual: false, gasSamplerYdManual: false,
  days: Array.from({ length: 4 }, () => ({ date: '', weather: '', temp: '', humidity: '', pressure: '', windDir: '', windSpeed: '' })),
  rows: Object.fromEntries(SAMPLING_LOG_SUBSTANCES.map((item) => [String(item.row), createEmptySamplingLogRow(item)])),
});

const sanitizeSamplingLog = (value) => {
  const base = createEmptySamplingLog();
  if (!value || typeof value !== 'object') return base;
  const text = (v) => (v === undefined || v === null ? '' : String(v));
  const next = { ...base };
  ['company', 'industry', 'facility', 'stackName', 'siteAbbr', 'particleSamplerId', 'particleSamplerYd', 'gasSamplerId', 'gasSamplerYd']
    .forEach((key) => { next[key] = text(value[key]); });
  next.particleSamplerYdManual = value.particleSamplerYdManual === true;
  next.gasSamplerYdManual = value.gasSamplerYdManual === true;
  next.days = base.days.map((day, idx) => {
    const src = Array.isArray(value.days) ? value.days[idx] : null;
    if (!src || typeof src !== 'object') return day;
    return Object.fromEntries(Object.keys(day).map((key) => [key, text(src[key])]));
  });
  next.rows = Object.fromEntries(SAMPLING_LOG_SUBSTANCES.map((item) => {
    const key = String(item.row);
    const src = value.rows && typeof value.rows === 'object' ? value.rows[key] : null;
    const row = createEmptySamplingLogRow(item);
    if (src && typeof src === 'object') {
      row.requested = Boolean(src.requested);
      const di = parseInt(src.dayIndex, 10);
      row.dayIndex = Number.isFinite(di) && di >= 0 && di < 4 ? di : 0;
      Object.keys(row).forEach((field) => {
        if (field === 'requested' || field === 'dayIndex' || field === 'speciesSel') return;
        if (src[field] !== undefined) row[field] = text(src[field]);
      });
      if (Array.isArray(item.options) && Array.isArray(src.speciesSel)) {
        const picked = item.options.filter((opt) => src.speciesSel.includes(opt));
        row.speciesSel = picked.length > 0 ? picked : [...item.options];
      }
    }
    return [key, row];
  }));
  return next;
};

// 시료번호: 날짜(mmdd)-시설약자-굴뚝명-물질코드-접미어 (양식 D열 수식과 동일)
const buildSamplingLogSampleNo = (log, substance) => {
  const row = log.rows[String(substance.row)];
  if (!row || !row.requested) return '';
  const date = log.days[row.dayIndex]?.date || '';
  const dt = date ? new Date(date) : null;
  const mmdd = dt && !Number.isNaN(dt.getTime())
    ? `${String(dt.getMonth() + 1).padStart(2, '0')}${String(dt.getDate()).padStart(2, '0')}`
    : '';
  return [mmdd, log.siteAbbr, log.stackName, substance.code, substance.suffix || '1'].map((v) => String(v || '').trim()).join('-');
};

// 채취유량(L): 입자/가스 샘플러는 (종료값-시작값)×Yd, 음압상자 등은 종료값-시작값
const calcSamplingLogVolume = (log, substance) => {
  const row = log.rows[String(substance.row)];
  if (!row) return NaN;
  const a = parseFloat(row.meterStart), b = parseFloat(row.meterEnd);
  if (!Number.isFinite(a) || !Number.isFinite(b)) return NaN;
  if (substance.kind === 'particle') {
    const yd = parseFloat(log.particleSamplerYd);
    return (b - a) * (Number.isFinite(yd) && yd > 0 ? yd : 1);
  }
  if (substance.kind === 'gas') {
    const yd = parseFloat(log.gasSamplerYd);
    return (b - a) * (Number.isFinite(yd) && yd > 0 ? yd : 1);
  }
  if (substance.kind === 'box' || substance.kind === 'plain') return b - a;
  return NaN;
};

const getSkyPhaseByHour = (date = new Date()) => {
  const hour = date.getHours();
  if (hour >= 6 && hour < 11) return 'morning';
  if (hour >= 11 && hour < 16) return 'day';
  return 'night';
};

const seededNoise = (seed) => {
  const x = Math.sin(seed * 12.9898) * 43758.5453;
  return x - Math.floor(x);
};

const SKYLINE_STARS = Array.from({ length: 46 }, (_, idx) => ({
  id: `star-${idx}`,
  left: `${2 + ((idx * 17) % 96)}%`,
  top: `${4 + ((idx * 13) % 44)}%`,
  size: 1.2 + (idx % 4) * 0.55,
  duration: 2.4 + (idx % 5) * 0.65,
  delay: (idx % 7) * 0.35,
}));

// 느리게 흘러가는 구름 (top: 화면 상단 기준 %, scale: 크기, duration: 한 번 가로지르는 시간)
const SKYLINE_CLOUDS = [
  { id: 'cl-1', top: '6%', scale: 1.25, duration: 170, delay: -20, opacity: 0.95 },
  { id: 'cl-2', top: '14%', scale: 0.8, duration: 130, delay: -75, opacity: 0.8 },
  { id: 'cl-3', top: '22%', scale: 1.05, duration: 200, delay: -140, opacity: 0.9 },
  { id: 'cl-4', top: '10%', scale: 0.6, duration: 115, delay: -40, opacity: 0.7 },
  { id: 'cl-5', top: '30%', scale: 0.9, duration: 185, delay: -110, opacity: 0.75 },
  { id: 'cl-6', top: '18%', scale: 1.4, duration: 230, delay: -180, opacity: 0.85 },
  { id: 'cl-7', top: '26%', scale: 0.7, duration: 150, delay: -10, opacity: 0.65 },
];

const SKY_PREVIEW_OPTIONS = [
  { id: 'auto', label: '자동' },
  { id: 'morning', label: '아침' },
  { id: 'day', label: '낮' },
  { id: 'night', label: '밤' },
];

const sanitizeSkyPreviewMode = (value) => {
  const mode = String(value || '').trim();
  return SKY_PREVIEW_OPTIONS.some((opt) => opt.id === mode) ? mode : 'auto';
};

const MOBILE_LIGHT_MAX_WIDTH = 768;
const DEFAULT_MEASURE_TOTAL_MINUTES = 60;
const DEFAULT_MEASURE_ROW_COUNT = 12;
const DEFAULT_SLOT_MINUTES = DEFAULT_MEASURE_TOTAL_MINUTES / DEFAULT_MEASURE_ROW_COUNT;
const PITOT_BASE_CP = 0.84; // 피토계수 자동 선정 시 유속 가계산에 쓰는 기준계수
const TABLE_PLAN_MAX_POINTS = 5; // 반경당 측정점 최대 개수
// 지점 수별 권장 지점당 행 수: 지점이 적으면 촘촘히(1점 6행), 많으면 지점당 3행
const RECOMMENDED_ROWS_PER_POINT = { 1: 6, 2: 4, 3: 4, 4: 3, 5: 3 };

// ── 사전 수분량 측정 방법 ──
// 사전 수분량 측정방법. 엑셀 수분량!C5 선택번호는 이 목록의 순서(1·2·3)를 그대로 쓴다
// 사후 수분량(채취 후 임핀저 전후 무게)은 방법이 아니라 아래쪽 칸에 따로 적는다
// 이름·기구·기록 항목은 「대기오염공정시험기준」 ES 01321(배출가스 중 수분량)을 따른다
//  - 흡습관법 ES 01321.1 (시료채취장치 1형): 흡습관 2개에 입자상 무수염화칼슘, 채취 전후 무게 차가 흡습 수분의 질량 ma
//  - 임핀저법 ES 01321.2 (시료채취장치 2형): 임핀저 4개(1·2번 물 100 g, 3번 빈 것, 4번 실리카젤 200 g), 전후 총무게 차가 ma
// 계산식(두 방법 공통): Xw = (22.4/18·ma) ÷ [Vm × 273/(273+θm) × (Pa+Pm)/760 + 22.4/18·ma] × 100
const DH_TARGETS_STORAGE_KEY = 'dust-sampling.dh-targets.v1';
const DEFAULT_DH_TARGETS = { max: 50, standard: 25, stable: 10 };
const MOISTURE_METHODS = [
  {
    id: 'imp2', label: '흡습관법 (2구)', stages: 2, standard: 'ES 01321.1',
    items: ['흡습관 1', '흡습관 2'], unitTitle: '흡습관',
    guide: '흡습관(U자관) 2개에 입자상 무수염화칼슘을 넣고, 채취 전(ma1)·후(ma2) 무게 차를 흡습 수분의 질량 ma로 씁니다. 흡습된 수분이 0.1 g~1.0 g이 되도록 흡입합니다.',
    maRange: [0.1, 1.0],
  },
  {
    id: 'impinger', label: '임핀저법 (4구)', stages: 4, standard: 'ES 01321.2',
    items: ['임핀저 1 (물 100 g)', '임핀저 2 (물 100 g)', '임핀저 3 (빈 것)', '임핀저 4 (실리카젤 200 g)'], unitTitle: '임핀저',
    guide: '임핀저 4개 중 1·2번에 물 100 g씩, 4번에 실리카젤 200 g을 넣고, 채취 전(ma1)·후(ma2) 총무게 차를 흡습 수분의 질량 ma로 씁니다.',
  },
  { id: 'auto', label: '자동측정법', standard: 'ES 01321.3', guide: '수분량 자동측정기로 5회 측정한 값의 평균을 씁니다.' },
];
const isImpingerMoistureMethod = (method) => Boolean(method && method.stages);
const createEmptyPreImpinger = () => ({
  impingers: [1, 2, 3, 4].map((id) => ({ id, initial: '', final: '' })),
  meterStart: '', meterEnd: '', meterTemp: '', pressure: '', orifice: '',
  // 흡습관법은 가스 샘플러로 흡입한다: 그 샘플러의 보정계수, 습식 가스미터 여부
  gasSamplerId: '', yd: '', wetMeter: false,
});
// 사전 수분량 계산에 쓰는 가스미터 보정계수: 흡습관법은 가스 샘플러의 Yd(적어 둔 값), 임핀저법은 입자 샘플러의 Yd
const getPreMoistureYd = (method, data) => {
  const num = (v) => { const n = parseFloat(v); return Number.isFinite(n) && n > 0 ? n : NaN; };
  if (method && method.id === 'imp2') {
    const own = num(data?.preImpinger?.yd);
    return Number.isFinite(own) ? own : 1;
  }
  const main = num(data?.gasMeterFactor);
  return Number.isFinite(main) ? main : 1;
};
// θ(℃)에서 물의 포화수증기압(mmHg). 습식 가스미터를 쓸 때 빼 주는 값 (ES 01114.c 식 9의 Pv)
const saturationVaporMmHg = (tempC) => (Number.isFinite(tempC) ? 4.5812 * Math.pow(10, (7.5 * tempC) / (237.3 + tempC)) : NaN);
const getPreMoisturePv = (method, data) => {
  if (!method || method.id !== 'imp2' || !data?.preImpinger?.wetMeter) return 0;
  const pv = saturationVaporMmHg(parseFloat(data.preImpinger.meterTemp));
  return Number.isFinite(pv) ? pv : 0;
};
// 실제로 측정값이 들어간 마지막 기록표 행 (표 자동 생성으로 미리 채워지는 채취점·시간은 측정값으로 보지 않음)
const METER_MEASURED_FIELDS = ['stackTemp', 'dp', 'sp', 'pressure', 'volume', 'tmIn', 'tmOut', 'vacuum', 'impingerTemp'];
const getLastMeasuredMeterIdx = (rows) => (rows || []).reduce((lastIdx, row, idx) => (
  idx > 0 && METER_MEASURED_FIELDS.some((field) => {
    const raw = row?.[field];
    return raw !== null && raw !== undefined && String(raw).trim() !== '';
  }) ? idx : lastIdx
), 0);

const isMobileViewportWidth = () => {
  if (typeof window === 'undefined') return false;
  return window.innerWidth <= MOBILE_LIGHT_MAX_WIDTH;
};

const makeSparseLayout = (items, step = 2) =>
  items.filter((_, idx) => idx % step === 0 || idx === items.length - 1);

const buildPointSlotSequence = (pointCount, slotCount = DEFAULT_MEASURE_ROW_COUNT) => {
  const safePointCount = Math.max(1, Math.floor(Number(pointCount) || 1));
  const safeSlotCount = Math.max(1, Math.floor(Number(slotCount) || 1));
  const baseCount = Math.floor(safeSlotCount / safePointCount);
  const counts = Array.from({ length: safePointCount }, () => baseCount);
  let remainder = safeSlotCount - (baseCount * safePointCount);

  let left = 0;
  let right = safePointCount - 1;
  let applyFromStart = true;
  while (remainder > 0) {
    if (left > right) {
      left = 0;
      right = safePointCount - 1;
    }
    if (left === right) {
      counts[left] += 1;
      remainder -= 1;
      continue;
    }
    if (applyFromStart) {
      counts[left] += 1;
      left += 1;
    } else {
      counts[right] += 1;
      right -= 1;
    }
    applyFromStart = !applyFromStart;
    remainder -= 1;
  }

  const sequence = [];
  counts.forEach((count, idx) => {
    for (let i = 0; i < count; i += 1) {
      sequence.push(String(idx + 1));
    }
  });
  return sequence.slice(0, safeSlotCount);
};

const createGasMeterRowsByPoints = (pointCount, slotCount = DEFAULT_MEASURE_ROW_COUNT, slotMinutes = DEFAULT_SLOT_MINUTES, totalMinutes = null) => {
  const sequence = buildPointSlotSequence(pointCount, slotCount);
  const safeSlot = Number.isFinite(slotMinutes) && slotMinutes > 0 ? slotMinutes : DEFAULT_SLOT_MINUTES;
  const safeTotal = Number.isFinite(totalMinutes) && totalMinutes > 0 ? totalMinutes : safeSlot * sequence.length;
  // 시간 열은 누적 분(5, 10, 15 …). 마지막 행은 총 채취시간으로 맞춤
  const minutesForRow = (idx) => {
    const cumulative = Math.round(safeSlot * (idx + 1) * 100) / 100;
    if (idx >= sequence.length - 1) return Math.round(safeTotal * 100) / 100;
    return Math.min(cumulative, safeTotal);
  };
  const rows = [{
    id: 0,
    pointNum: '시작',
    time: '0',
    stackTemp: '',
    dp: '',
    sp: '',
    pressure: '',
    volume: '',
    tmIn: '',
    tmOut: '',
    vacuum: '',
    impingerTemp: '',
  }];

  sequence.forEach((pointNum, idx) => {
    rows.push({
      id: idx + 1,
      pointNum,
      time: String(minutesForRow(idx)),
      stackTemp: '',
      dp: '',
      sp: '',
      pressure: '',
      volume: '',
      tmIn: '',
      tmOut: '',
      vacuum: '',
      impingerTemp: '',
    });
  });

  return rows;
};

const normalizeGasMeterRows = (rows, fallbackPointCount = 5) => {
  const sourceRows = Array.isArray(rows)
    ? rows.filter((row) => row && typeof row === 'object')
    : [];

  const safeRows = sourceRows.length > 0
    ? sourceRows
    : createGasMeterRowsByPoints(Math.max(1, Number(fallbackPointCount) || 1));

  const normalizedRows = safeRows.map((row, idx) => ({
    id: idx,
    pointNum: idx === 0
      ? '시작'
      : (row.pointNum === undefined || row.pointNum === null ? '' : String(row.pointNum)),
    time: idx === 0
      ? '0'
      : (row.time === undefined || row.time === null ? '' : String(row.time)),
    stackTemp: row.stackTemp === undefined || row.stackTemp === null ? '' : String(row.stackTemp),
    dp: row.dp === undefined || row.dp === null ? '' : String(row.dp),
    sp: row.sp === undefined || row.sp === null ? '' : String(row.sp),
    pressure: row.pressure === undefined || row.pressure === null ? '' : String(row.pressure),
    volume: row.volume === undefined || row.volume === null ? '' : String(row.volume),
    tmIn: row.tmIn === undefined || row.tmIn === null ? '' : String(row.tmIn),
    tmOut: row.tmOut === undefined || row.tmOut === null ? '' : String(row.tmOut),
    vacuum: row.vacuum === undefined || row.vacuum === null ? '' : String(row.vacuum),
    impingerTemp: row.impingerTemp === undefined || row.impingerTemp === null ? '' : String(row.impingerTemp),
  }));

  if (normalizedRows.length === 1) {
    normalizedRows.push({
      id: 1,
      pointNum: '',
      time: String(DEFAULT_SLOT_MINUTES),
      stackTemp: '',
      dp: '',
      sp: '',
      pressure: '',
      volume: '',
      tmIn: '',
      tmOut: '',
      vacuum: '',
      impingerTemp: '',
    });
  }

  return normalizedRows;
};

const SKYLINE_FLYER_CSS = `
@keyframes skylineTwinkle {
  0%, 100% { opacity: 0.26; transform: scale(0.9); }
  50% { opacity: 0.95; transform: scale(1.08); }
}
.skyline-star {
  position: absolute;
  border-radius: 9999px;
  background: rgba(241, 245, 249, 0.95);
  box-shadow: 0 0 6px rgba(224, 231, 255, 0.8), 0 0 14px rgba(191, 219, 254, 0.45);
  animation: skylineTwinkle var(--twinkle-duration, 3.2s) ease-in-out infinite;
  animation-delay: var(--twinkle-delay, 0s);
  will-change: opacity, transform;
}
.skyline-moon {
  position: absolute;
  left: 8%;
  top: 8%;
  width: clamp(72px, 9vw, 108px);
  height: clamp(72px, 9vw, 108px);
  border-radius: 9999px;
  background: radial-gradient(circle at 35% 35%, rgba(254, 249, 195, 1) 0 44%, rgba(253, 230, 138, 0.96) 68%, rgba(250, 204, 21, 0.92) 100%);
  -webkit-mask: radial-gradient(circle at 70% 50%, transparent 46%, #000 47%);
  mask: radial-gradient(circle at 70% 50%, transparent 46%, #000 47%);
  box-shadow: 0 0 24px rgba(250, 204, 21, 0.35), 0 0 54px rgba(165, 180, 252, 0.18);
  transform: rotate(-12deg);
}
.skyline-moon::before {
  content: none;
}
.skyline-moon::after {
  content: none;
}
.skyline-sun {
  position: absolute;
  left: 8%;
  top: 8%;
  width: clamp(74px, 9vw, 112px);
  height: clamp(74px, 9vw, 112px);
  border-radius: 9999px;
  background: radial-gradient(circle at 36% 34%, rgba(254, 249, 195, 1), rgba(253, 230, 138, 0.95) 48%, rgba(251, 191, 36, 0.88));
  box-shadow: 0 0 0 12px rgba(254, 240, 138, 0.25), 0 0 36px rgba(250, 204, 21, 0.35), 0 0 70px rgba(249, 115, 22, 0.2);
}
.skyline-sunrise-mountain {
  position: absolute;
  left: 80%;
  top: 33%;
  width: clamp(92px, 10vw, 136px);
  height: clamp(92px, 10vw, 136px);
  border-radius: 9999px;
  background: radial-gradient(circle at 50% 38%, rgba(254, 249, 195, 1), rgba(253, 224, 71, 0.96) 44%, rgba(251, 146, 60, 0.9));
  box-shadow: 0 0 0 11px rgba(254, 240, 138, 0.2), 0 0 36px rgba(251, 191, 36, 0.32), 0 0 66px rgba(251, 146, 60, 0.2);
}
.skyline-sunrise-mountain-sheet {
  position: absolute;
  left: 78%;
  top: 54%;
  width: clamp(56px, 7vw, 84px);
  height: clamp(56px, 7vw, 84px);
  border-radius: 9999px;
  background: radial-gradient(circle at 50% 38%, rgba(254, 249, 195, 0.98), rgba(253, 224, 71, 0.92) 44%, rgba(251, 146, 60, 0.86));
  box-shadow: 0 0 0 8px rgba(254, 240, 138, 0.18), 0 0 28px rgba(251, 191, 36, 0.28), 0 0 48px rgba(251, 146, 60, 0.18);
}
@keyframes skylineCloudDrift {
  from { transform: translateX(-30vw) scale(var(--cloud-scale, 1)); }
  to { transform: translateX(125vw) scale(var(--cloud-scale, 1)); }
}
.skyline-cloud, .skyline-cloud::before, .skyline-cloud::after {
  background: var(--cloud-color, rgba(255, 255, 255, 0.92));
  border-radius: 9999px;
}
.skyline-cloud {
  position: absolute;
  left: 0;
  width: 190px;
  height: 54px;
  transform-origin: left center;
  filter: blur(0.6px) drop-shadow(0 6px 10px var(--cloud-shadow, rgba(120, 140, 170, 0.18)));
  animation: skylineCloudDrift var(--cloud-duration, 160s) linear infinite;
  animation-delay: var(--cloud-delay, 0s);
  will-change: transform;
}
.skyline-cloud::before {
  content: '';
  position: absolute;
  left: 18%;
  bottom: 30%;
  width: 56%;
  height: 140%;
}
.skyline-cloud::after {
  content: '';
  position: absolute;
  left: 52%;
  bottom: 24%;
  width: 40%;
  height: 120%;
}
@keyframes skylineSunRays {
  from { transform: rotate(0deg); }
  to { transform: rotate(360deg); }
}
.skyline-sun::after {
  content: '';
  position: absolute;
  inset: -90%;
  border-radius: 9999px;
  background: repeating-conic-gradient(from 0deg, rgba(255, 236, 170, 0.22) 0deg 7deg, rgba(255, 236, 170, 0) 7deg 24deg);
  -webkit-mask: radial-gradient(circle, rgba(0, 0, 0, 0.9) 20%, rgba(0, 0, 0, 0) 68%);
  mask: radial-gradient(circle, rgba(0, 0, 0, 0.9) 20%, rgba(0, 0, 0, 0) 68%);
  animation: skylineSunRays 90s linear infinite;
}
.skyline-moon-halo {
  position: absolute;
  left: calc(8% - 40px);
  top: calc(8% - 40px);
  width: clamp(152px, 17vw, 190px);
  height: clamp(152px, 17vw, 190px);
  border-radius: 9999px;
  background: radial-gradient(circle, rgba(250, 240, 200, 0.22), rgba(199, 210, 254, 0.1) 45%, rgba(199, 210, 254, 0) 72%);
}
@keyframes skylineShootingStar {
  0% { opacity: 0; transform: translate(0, 0) rotate(-28deg); }
  3% { opacity: 1; }
  10% { opacity: 0; transform: translate(-300px, 160px) rotate(-28deg); }
  100% { opacity: 0; transform: translate(-300px, 160px) rotate(-28deg); }
}
.skyline-shooting-star {
  position: absolute;
  width: 150px;
  height: 2px;
  border-radius: 9999px;
  background: linear-gradient(90deg, rgba(255, 255, 255, 0), rgba(255, 255, 255, 0.95));
  box-shadow: 0 0 8px rgba(255, 255, 255, 0.6);
  opacity: 0;
  animation: skylineShootingStar var(--shoot-duration, 13s) ease-out infinite;
  animation-delay: var(--shoot-delay, 0s);
}
@keyframes skylineWindowGlow {
  0%, 100% { opacity: 0.55; }
  50% { opacity: 1; }
}
.skyline-window-night {
  animation: skylineWindowGlow var(--glow-duration, 5s) ease-in-out infinite;
  animation-delay: var(--glow-delay, 0s);
}
@keyframes skylineSvgSmoke {
  0% { transform: translate(0px, 0px) scale(0.45); opacity: 0; }
  12% { opacity: var(--max-opacity, 0.5); }
  100% { transform: translate(var(--drift, 80px), calc(-1 * var(--rise, 260px))) scale(var(--end-scale, 3)); opacity: 0; }
}
.skyline-svg-smoke {
  transform-box: fill-box;
  transform-origin: center;
  opacity: 0;
  animation: skylineSvgSmoke var(--rise-duration, 16s) ease-out infinite;
  animation-delay: var(--rise-delay, 0s);
}
@keyframes skylineBeacon {
  0%, 100% { opacity: 0.15; }
  50% { opacity: 1; }
}
.skyline-beacon {
  animation: skylineBeacon 2.4s ease-in-out infinite;
  animation-delay: var(--beacon-delay, 0s);
}
@keyframes skylineFlap {
  from { transform: scaleY(1); }
  to { transform: scaleY(0.3); }
}
.skyline-bird {
  transform-box: fill-box;
  transform-origin: center;
  animation: skylineFlap var(--flap-duration, 0.8s) ease-in-out infinite alternate;
  animation-delay: var(--flap-delay, 0s);
}
@keyframes skylineFog {
  from { transform: translateX(-70px); }
  to { transform: translateX(70px); }
}
.skyline-fog {
  animation: skylineFog var(--fog-duration, 36s) ease-in-out infinite alternate;
  animation-delay: var(--fog-delay, 0s);
}
.skyline-svg {
  position: absolute;
  left: 0;
  right: 0;
  bottom: 0;
  width: 100%;
  height: 100%;
  display: block;
  /* 연기가 그림 영역 위로 올라가도 잘리지 않고 하늘로 이어지게 */
  overflow: visible;
}
`;

// ── 겹겹이 오린 종이 디오라마: 깊이별 한 가지 색 실루엣으로 그린 공단 풍경 ──
// 빛은 왼쪽 위. 각 층(먼 산 → 도시 → 공장 → 앞 땅)은 같은 색조의 명암 두 단계만 쓴다.
const SKYLINE_PALETTES = {
  day: {
    far: '#cfdceb', farDetail: '#bccbdd', city: '#a7bbd1', cityWin: '#bfd0e2', haze1: 'rgba(222,236,250,0.65)', haze2: 'rgba(214,230,246,0.5)',
    plant: '#73899f', plantShade: '#5e7389', plantLight: '#8ea2b6', plantBand: '#9fb1c3', plantWin: '#95a9bd',
    near: '#4b7362', nearShade: '#3a5c4e', nearLight: '#64917b', road: '#587566', roadLine: 'rgba(235,245,238,0.75)', rail: '#34514a',
    smoke: ['rgba(255,255,255,0.85)', 'rgba(255,255,255,0)'], windowLit: null, beacon: null, lamp: null, headlight: null,
    bird: '#48607a', plane: '#f2f6fa', contrail: 'rgba(255,255,255,0.7)', cirrus: 'rgba(255,255,255,0.55)', sunHaze: 'rgba(255,238,200,0.22)',
  },
  morning: {
    far: '#e3cfdd', farDetail: '#d4bccd', city: '#c0a3b8', cityWin: '#d2bacb', haze1: 'rgba(255,228,214,0.6)', haze2: 'rgba(250,220,210,0.45)',
    plant: '#8f7489', plantShade: '#776076', plantLight: '#a58ca0', plantBand: '#b59fb0', plantWin: '#ad95a8',
    near: '#5a4f6a', nearShade: '#473d57', nearLight: '#6e6380', road: '#66597a', roadLine: 'rgba(255,235,225,0.65)', rail: '#3d3450',
    smoke: ['rgba(255,246,238,0.85)', 'rgba(255,246,238,0)'], windowLit: null, beacon: null, lamp: null, headlight: null,
    bird: '#5a4a5c', plane: '#fff6ee', contrail: 'rgba(255,236,220,0.7)', cirrus: 'rgba(255,240,228,0.5)', sunHaze: 'rgba(255,190,150,0.3)',
  },
  night: {
    far: '#1a2850', farDetail: '#162243', city: '#111c3a', cityWin: '#ffd47a', haze1: 'rgba(30,45,90,0.5)', haze2: 'rgba(20,32,70,0.5)',
    plant: '#0d1630', plantShade: '#080f22', plantLight: '#16223f', plantBand: '#5a2d35', plantWin: '#ffd47a',
    near: '#060a18', nearShade: '#04060f', nearLight: '#0b1224', road: '#0e1526', roadLine: 'rgba(255,230,180,0.3)', rail: '#03050c',
    smoke: ['rgba(200,210,235,0.5)', 'rgba(200,210,235,0)'], windowLit: '#ffd47a', beacon: '#ff5a4a', lamp: 'rgba(255,214,130,0.9)', headlight: 'rgba(255,240,200,0.9)',
    bird: null, plane: '#8e9ab5', contrail: null, cirrus: 'rgba(190,205,240,0.07)', sunHaze: 'rgba(255,170,90,0.14)',
  },
};

// 바닥선. 앞 땅 층이 420 근처부터 덮으므로 물체 밑동은 자연스럽게 땅에 묻힌다
const SCENE_GROUND = 432;
// 도시 층: 촘촘한 건물 열 (같은 색이라 겹치면 한 덩어리로 보인다)
const SCENE_CITY = Array.from({ length: 64 }, (_, i) => ({
  x: -12 + i * 26 + ((i * 7) % 9), w: 14 + ((i * 5) % 20), h: 26 + ((i * 13) % 74), antenna: i % 9 === 4,
}));
// 공장 층
const SCENE_HALLS = [{ x: 40, w: 280, h: 72, teeth: 5 }, { x: 880, w: 320, h: 92, teeth: 6 }, { x: 1320, w: 220, h: 62, teeth: 4 }];
const SCENE_BLOCKS = [{ x: 600, w: 84, h: 196 }, { x: 1090, w: 72, h: 166 }, { x: 0, w: 56, h: 140 }, { x: 760, w: 60, h: 120 }, { x: 1560, w: 60, h: 150 }];
const SCENE_STACKS = [
  { x: 150, w: 34, h: 296, smoke: 7 }, { x: 362, w: 48, h: 392, smoke: 9 }, { x: 556, w: 40, h: 336, smoke: 8 },
  { x: 1004, w: 52, h: 380, smoke: 9 }, { x: 1186, w: 30, h: 256, smoke: 6 }, { x: 1486, w: 36, h: 300, smoke: 7 },
];
const SCENE_COOLING = [{ x: 690, w: 150, h: 208 }, { x: 1236, w: 118, h: 178 }];
const SCENE_TANKS = [{ x: 420, w: 72, h: 60 }, { x: 504, w: 72, h: 82 }, { x: 1398, w: 62, h: 50 }, { x: 1468, w: 62, h: 50 }];
const SCENE_SPHERE = { cx: 862, cy: 388, r: 40 };
const SCENE_CRANE = { x: 262, top: 146, jib: 230, counter: 72, hookX: 140 };
const SCENE_PYLONS = [{ x: 470, h: 150 }, { x: 940, h: 165 }, { x: 1300, h: 150 }];
// 앞 땅 층: 나무는 크기가 다른 무리로 (같은 색 실루엣이라 겹치며 숲 덩어리가 된다)
const SCENE_TREE_CLUSTERS = [
  [40, 62, 88], [230, 250], [330, 352, 372, 390], [560, 580], [700, 724, 744], [880, 902], [1010, 1036, 1052, 1076], [1210, 1230], [1360, 1384, 1402], [1520, 1544, 1566, 1590],
];
const SCENE_TREES = SCENE_TREE_CLUSTERS.flatMap((cluster, ci) => cluster.map((x, k) => ({
  x, r: 15 + ((ci * 5 + k * 3) % 11), h: 20 + ((ci * 7 + k * 5) % 16),
})));
const SCENE_LAMPS = [120, 400, 680, 960, 1240, 1500];
const SCENE_CARS = [
  { dir: 1, dur: 36, delay: -6, len: 26 }, { dir: -1, dur: 44, delay: -21, len: 24 }, { dir: 1, dur: 52, delay: -33, len: 36 }, { dir: -1, dur: 40, delay: -13, len: 22 },
];
const SCENE_TRAIN = { dur: 100, delay: -42, cars: [50, 42, 46, 46, 42, 46, 42, 46] };
const SCENE_BIRD_FLOCKS = [
  { path: 'M-80,150 C300,120 700,170 1700,130', dur: 52, delay: -8, birds: [[0, 0], [-16, 9], [-32, 18], [16, 9], [32, 18], [-48, 26]] },
  { path: 'M1700,200 C1200,170 600,230 -120,190', dur: 64, delay: -30, birds: [[0, 0], [-14, 8], [14, 8], [-28, 16]] },
];
const SCENE_PLANE = { path: 'M-160,58 L1760,92', dur: 78, delay: -26 };
const SCENE_ROAD_Y = 478;
const SCENE_RAIL_Y = 500;

const renderSkylineSvg = ({ mode, scope, compact = false }) => {
  const pal = SKYLINE_PALETTES[mode] || SKYLINE_PALETTES.day;
  const id = (name) => `${scope}-sky-${name}`;
  const G = SCENE_GROUND;
  const stackScale = compact ? 0.66 : 1;
  const shadeW = (w) => Math.max(4, Math.round(w * 0.34));
  const night = Boolean(pal.windowLit);

  // 공장 층의 상자형 물체: 본체 + 오른쪽 그늘면 (같은 층 색조)
  const box = (key, x, y, w, h, extra = null) => (
    <g key={key}>
      <rect x={x} y={y} width={w} height={h} fill={pal.plant} />
      <rect x={x + w - shadeW(w)} y={y} width={shadeW(w)} height={h} fill={pal.plantShade} />
      {extra}
    </g>
  );
  const windowsFor = (x, y, w, h, seedBase) => {
    const cols = Math.max(2, Math.floor(w / 14));
    const rows = Math.max(2, Math.floor((h - 16) / (compact ? 26 : 18)));
    const out = [];
    for (let r = 0; r < rows; r += 1) for (let c = 0; c < cols; c += 1) {
      const seed = seedBase + r * 17 + c * 7;
      const lit = night && seededNoise(seed) > 0.45;
      if (night && !lit) continue;
      out.push(<rect key={`${r}-${c}`} x={x + 6 + c * ((w - 12) / cols) + 1} y={y + 10 + r * ((h - 16) / rows)} width={Math.max(3, (w - 12) / cols - 4)} height="6" fill={pal.plantWin} className={lit ? 'skyline-window-night' : undefined} style={lit ? { '--glow-duration': `${4 + (seed % 5)}s`, '--glow-delay': `-${seed % 7}s` } : undefined} />);
    }
    return out;
  };
  const coolingPath = (t) => {
    const top = G - t.h, x0 = t.x, x1 = t.x + t.w, inset = t.w * 0.17, waist = G - t.h * 0.42;
    return `M${x0},${G} C${x0 + inset * 0.3},${waist} ${x0 + inset},${waist} ${x0 + inset},${top + 18} L${x1 - inset},${top + 18} C${x1 - inset},${waist} ${x1 - inset * 0.3},${waist} ${x1},${G} Z`;
  };
  const smokeFor = (cx, cy, baseR, count, seedBase) => Array.from({ length: compact ? Math.ceil(count * 0.5) : count }).map((_, i) => {
    const seed = seedBase + i * 31;
    return (
      <circle key={i} className="skyline-svg-smoke" cx={cx} cy={cy} r={baseR + (i % 3) * 3} fill={`url(#${id('smoke')})`} filter={`url(#${id('blur')})`}
        style={{ '--drift': `${(seededNoise(seed + 2) > 0.2 ? 1 : -1) * (60 + seededNoise(seed + 3) * 150)}px`, '--rise': `${170 + seededNoise(seed + 5) * 170}px`, '--end-scale': `${2.6 + seededNoise(seed + 8) * 1.6}`, '--max-opacity': `${0.35 + seededNoise(seed + 6) * 0.3}`, '--rise-duration': `${14 + seededNoise(seed + 4) * 9}s`, '--rise-delay': `-${seededNoise(seed + 11) * 23}s` }} />
    );
  });
  const trees = compact ? SCENE_TREES.filter((_, i) => i % 2 === 0) : SCENE_TREES;
  // 나무 무리 실루엣: 둥근 머리들을 한 경로로 이어 숲 덩어리로
  const treeBlob = (list, yBase) => list.map((t) => `M${t.x - t.r},${yBase} A${t.r},${t.r} 0 0 1 ${t.x + t.r},${yBase} Z`).join(' ');

  return (
    <svg className="skyline-svg" viewBox="0 0 1600 520" preserveAspectRatio="xMidYMax slice" aria-hidden="true">
      <defs>
        <radialGradient id={id('smoke')} cx="0.45" cy="0.4" r="0.55">
          <stop offset="0" stopColor={pal.smoke[0]} />
          <stop offset="1" stopColor={pal.smoke[1]} />
        </radialGradient>
        <linearGradient id={id('haze1')} x1="0" y1="0" x2="0" y2="1">
          <stop offset="0" stopColor={pal.haze1} stopOpacity="0" />
          <stop offset="1" stopColor={pal.haze1} />
        </linearGradient>
        <linearGradient id={id('haze2')} x1="0" y1="0" x2="0" y2="1">
          <stop offset="0" stopColor={pal.haze2} stopOpacity="0" />
          <stop offset="1" stopColor={pal.haze2} />
        </linearGradient>
        <linearGradient id={id('sunhaze')} x1="0" y1="0" x2="0" y2="1">
          <stop offset="0" stopColor={pal.sunHaze} stopOpacity="0" />
          <stop offset="0.8" stopColor={pal.sunHaze} />
          <stop offset="1" stopColor={pal.sunHaze} stopOpacity="0" />
        </linearGradient>
        <linearGradient id={id('ground')} x1="0" y1="0" x2="0" y2="1">
          <stop offset="0" stopColor={pal.near} />
          <stop offset="1" stopColor={pal.nearShade} />
        </linearGradient>
        {pal.contrail && (
          <linearGradient id={id('contrail')} x1="0" y1="0" x2="1" y2="0">
            <stop offset="0" stopColor={pal.contrail} stopOpacity="0" />
            <stop offset="1" stopColor={pal.contrail} />
          </linearGradient>
        )}
        {pal.lamp && (
          <radialGradient id={id('lampglow')} cx="0.5" cy="0" r="0.75">
            <stop offset="0" stopColor={pal.lamp} stopOpacity="0.55" />
            <stop offset="1" stopColor={pal.lamp} stopOpacity="0" />
          </radialGradient>
        )}
        <filter id={id('blur')} x="-50%" y="-50%" width="200%" height="200%"><feGaussianBlur stdDeviation="3" /></filter>
        <filter id={id('soft')} x="-20%" y="-20%" width="140%" height="140%"><feGaussianBlur stdDeviation="1.2" /></filter>
        <filter id={id('wide')} x="-20%" y="-100%" width="140%" height="300%"><feGaussianBlur stdDeviation="9" /></filter>
      </defs>

      {/* 하늘 아래 빛 띠와 권운 */}
      <rect x="0" y="140" width="1600" height="190" fill={`url(#${id('sunhaze')})`} />
      {[[120, 48, 420, 7], [640, 92, 300, 5], [980, 36, 520, 8], [1380, 118, 260, 5]].map(([x, y, w, h], i) => (
        <ellipse key={`cirrus-${i}`} cx={x + w / 2} cy={y} rx={w / 2} ry={h / 2} fill={pal.cirrus} filter={`url(#${id('wide')})`} />
      ))}

      {/* 1층: 먼 산 */}
      <path d="M0,296 L70,262 L130,280 L210,226 L290,268 L350,246 L430,286 L520,214 L600,262 L660,244 L740,282 L830,228 L900,268 L960,250 L1040,292 L1130,230 L1200,266 L1260,238 L1340,288 L1420,224 L1500,264 L1560,250 L1600,270 L1600,520 L0,520 Z" fill={pal.far} />
      <path d="M520,214 L498,236 L545,234 Z M830,228 L808,248 L852,246 Z M1420,224 L1398,246 L1442,244 Z M210,226 L190,246 L232,244 Z" fill={pal.farDetail} opacity="0.9" />
      <rect x="0" y="230" width="1600" height="120" fill={`url(#${id('haze1')})`} />

      {/* 2층: 도시 실루엣 */}
      <g>
        {SCENE_CITY.map((b, i) => (
          <g key={`city-${i}`}>
            <rect x={b.x} y={318 - b.h} width={b.w} height={b.h + 60} fill={pal.city} />
            {b.antenna && <rect x={b.x + b.w / 2 - 1} y={318 - b.h - 18} width="2" height="18" fill={pal.city} />}
            {(night ? i % 2 === 0 : i % 3 === 0) && [0, 1, 2].map((k) => (
              <rect key={k} x={b.x + 3} y={318 - b.h + 6 + k * 12} width={Math.max(4, b.w - 6)} height="2" fill={pal.cityWin} opacity={night ? 0.55 : 0.8} />
            ))}
          </g>
        ))}
      </g>
      <rect x="0" y="300" width="1600" height="150" fill={`url(#${id('haze2')})`} />

      {/* 3층: 공단. 한 색조의 실루엣, 오른쪽 그늘면만 어둡게 */}
      {!compact && (() => {
        const c = SCENE_CRANE;
        return (
          <g stroke={pal.plantShade} strokeWidth="2" fill="none" strokeLinecap="round">
            <path d={`M${c.x - 7},${G} L${c.x - 7},${c.top} M${c.x + 7},${G} L${c.x + 7},${c.top}`} />
            <path d={Array.from({ length: 11 }).map((_, i) => `M${c.x - 7},${c.top + 14 + i * 26} L${c.x + 7},${c.top + i * 26}`).join(' ')} strokeWidth="1.2" />
            <path d={`M${c.x - c.counter},${c.top - 10} L${c.x + c.jib},${c.top - 10}`} strokeWidth="3" />
            <path d={`M${c.x - c.counter},${c.top - 10} L${c.x},${c.top - 36} L${c.x + c.jib},${c.top - 10} M${c.x},${c.top - 36} L${c.x},${c.top - 18}`} strokeWidth="1.2" />
            <rect x={c.x - c.counter + 6} y={c.top - 10} width="24" height="14" fill={pal.plantShade} stroke="none" />
            <path d={`M${c.x + c.hookX},${c.top - 10} L${c.x + c.hookX},${c.top + 150}`} strokeWidth="1" />
            <rect x={c.x + c.hookX - 10} y={c.top + 150} width="20" height="16" fill={pal.plantShade} stroke="none" />
            {pal.beacon && <circle className="skyline-beacon" cx={c.x} cy={c.top - 38} r="3" fill={pal.beacon} stroke="none" style={{ '--beacon-delay': '-1.6s' }} />}
          </g>
        );
      })()}
      {SCENE_BLOCKS.map((b, i) => box(`block-${i}`, b.x, G - b.h, b.w, b.h, windowsFor(b.x, G - b.h, b.w, b.h, i * 131)))}
      {SCENE_HALLS.map((f, i) => {
        const tw = f.w / f.teeth;
        const roof = Array.from({ length: f.teeth }).map((_, k) => `L${f.x + k * tw},${G - f.h} L${f.x + k * tw + tw * 0.62},${G - f.h - 22} L${f.x + (k + 1) * tw},${G - f.h}`).join(' ');
        return box(`hall-${i}`, f.x, G - f.h, f.w, f.h, (
          <>
            <path d={`M${f.x},${G - f.h} ${roof} Z`} fill={pal.plantLight} />
            <path d={Array.from({ length: f.teeth }).map((_, k) => `M${f.x + k * tw + tw * 0.62},${G - f.h - 22} L${f.x + (k + 1) * tw},${G - f.h} L${f.x + k * tw + tw * 0.62},${G - f.h} Z`).join(' ')} fill={pal.plantShade} />
            <rect x={f.x + 14} y={G - f.h * 0.62} width={f.w - 28} height="6" fill={pal.plantWin} opacity={night ? 0.9 : 1} className={night ? 'skyline-window-night' : undefined} style={night ? { '--glow-duration': `${5 + i}s`, '--glow-delay': `-${i * 2}s` } : undefined} />
          </>
        ));
      })}
      {!compact && (() => {
        const s = SCENE_SPHERE;
        return (
          <g key="sphere">
            <path d={`M${s.cx - s.r * 0.6},${G} L${s.cx - s.r * 0.45},${s.cy + s.r * 0.5} M${s.cx + s.r * 0.6},${G} L${s.cx + s.r * 0.45},${s.cy + s.r * 0.5} M${s.cx - s.r * 0.18},${G} L${s.cx - s.r * 0.14},${s.cy + s.r * 0.92} M${s.cx + s.r * 0.18},${G} L${s.cx + s.r * 0.14},${s.cy + s.r * 0.92}`} stroke={pal.plantShade} strokeWidth="2.5" fill="none" />
            <circle cx={s.cx} cy={s.cy} r={s.r} fill={pal.plant} />
            <path d={`M${s.cx + s.r * 0.25},${s.cy - s.r * 0.97} A${s.r},${s.r} 0 0 1 ${s.cx + s.r * 0.25},${s.cy + s.r * 0.97} A${s.r * 0.55},${s.r} 0 0 0 ${s.cx + s.r * 0.25},${s.cy - s.r * 0.97} Z`} fill={pal.plantShade} />
            <ellipse cx={s.cx} cy={s.cy} rx={s.r} ry={s.r * 0.16} fill="none" stroke={pal.plantLight} strokeWidth="1.5" opacity="0.8" />
          </g>
        );
      })()}
      {SCENE_TANKS.map((t, i) => box(`tank-${i}`, t.x, G - t.h, t.w, t.h, (
        <>
          <ellipse cx={t.x + t.w / 2} cy={G - t.h} rx={t.w / 2} ry="5" fill={pal.plantLight} />
          <rect x={t.x} y={G - t.h * 0.5} width={t.w} height="3" fill={pal.plantBand} opacity="0.7" />
        </>
      )))}
      {SCENE_COOLING.map((t, i) => (
        <g key={`cooling-${i}`}>
          <path d={coolingPath(t)} fill={pal.plant} />
          <path d={`M${t.x + t.w * 0.62},${G} C${t.x + t.w * 0.66},${G - t.h * 0.42} ${t.x + t.w * 0.72},${G - t.h * 0.42} ${t.x + t.w * 0.72},${G - t.h + 18} L${t.x + t.w - t.w * 0.17},${G - t.h + 18} C${t.x + t.w - t.w * 0.17},${G - t.h * 0.42} ${t.x + t.w - t.w * 0.05},${G - t.h * 0.42} ${t.x + t.w},${G} Z`} fill={pal.plantShade} />
          <ellipse cx={t.x + t.w / 2} cy={G - t.h + 18} rx={t.w * 0.33} ry="5" fill={pal.plantShade} />
          {smokeFor(t.x + t.w / 2, G - t.h + 14, 16, 5, i * 53)}
        </g>
      ))}
      {SCENE_STACKS.map((st, i) => {
        const h = Math.round(st.h * stackScale), top = G - h;
        return (
          <g key={`stack-${i}`}>
            <rect x={st.x} y={top} width={st.w} height={h} fill={pal.plant} />
            <rect x={st.x + st.w - shadeW(st.w)} y={top} width={shadeW(st.w)} height={h} fill={pal.plantShade} />
            <rect x={st.x - 3} y={top} width={st.w + 6} height="7" fill={pal.plantShade} />
            <rect x={st.x} y={top + 24} width={st.w} height="8" fill={pal.plantBand} />
            <rect x={st.x} y={top + 48} width={st.w} height="8" fill={pal.plantBand} />
            {pal.beacon && <circle className="skyline-beacon" cx={st.x + st.w / 2} cy={top - 4} r="3.2" fill={pal.beacon} style={{ '--beacon-delay': `-${i * 0.9}s` }} />}
            {smokeFor(st.x + st.w / 2, top - 2, st.w * 0.32, st.smoke, i * 97)}
          </g>
        );
      })}
      {SCENE_PYLONS.map((py, i) => {
        const top = G - py.h, next = SCENE_PYLONS[i + 1];
        return (
          <g key={`pylon-${i}`} stroke={pal.plantShade} strokeWidth="2" fill="none" strokeLinecap="round">
            <path d={`M${py.x - 16},${G} L${py.x - 4},${top} L${py.x + 4},${top} L${py.x + 16},${G}`} />
            <path d={`M${py.x - 12},${G - 30} L${py.x + 12},${G - 30} M${py.x - 9},${G - 62} L${py.x + 9},${G - 62} M${py.x - 13},${top + 22} L${py.x + 13},${top + 22} M${py.x - 22},${top + 46} L${py.x + 22},${top + 46}`} />
            {next && [[22, 46], [13, 22]].map(([dx, dy], k) => (
              <path key={k} d={`M${py.x + dx},${top + dy} Q${(py.x + next.x) / 2},${Math.max(top, G - next.h) + dy + 46} ${next.x - dx},${G - next.h + dy} M${py.x - dx},${top + dy} Q${(py.x + next.x) / 2},${Math.max(top, G - next.h) + dy + 46} ${next.x + dx},${G - next.h + dy}`} strokeWidth="1" opacity="0.7" />
            ))}
          </g>
        );
      })}
      {/* 공단 밑동을 땅에 묻는 접지 안개 */}
      <rect x="0" y={G - 40} width="1600" height="44" fill={`url(#${id('haze2')})`} />

      {/* 4층: 앞 땅. 언덕 + 숲 덩어리 + 도로 + 철길, 전부 한 색조 */}
      <path d={`M0,${G - 14} C180,${G - 36} 300,${G} 480,${G - 24} C660,${G - 48} 820,${G - 8} 1000,${G - 30} C1180,${G - 52} 1320,${G - 10} 1600,${G - 36} L1600,520 L0,520 Z`} fill={`url(#${id('ground')})`} />
      <path d={treeBlob(trees, G + 6)} fill={pal.nearLight} transform={`translate(0 ${-26})`} />
      <path d={treeBlob(trees.map((t) => ({ ...t, r: t.r * 0.72, x: t.x + 3 })), G + 6)} fill={pal.nearShade} transform={`translate(0 ${-22})`} opacity="0.55" />
      <rect x="0" y={G + 6} width="1600" height="100" fill={pal.near} />
      <rect x="0" y={G + 6} width="1600" height="18" fill={pal.nearLight} opacity="0.35" />
      {/* 도로 */}
      <rect x="0" y={SCENE_ROAD_Y - 8} width="1600" height="16" fill={pal.road} />
      <path d={`M0,${SCENE_ROAD_Y} L1600,${SCENE_ROAD_Y}`} stroke={pal.roadLine} strokeWidth="1.3" strokeDasharray="14 12" />
      {SCENE_LAMPS.map((x, i) => (
        <g key={`lamp-${i}`}>
          {pal.lamp && <ellipse cx={x + 10} cy={SCENE_ROAD_Y - 34} rx="36" ry="46" fill={`url(#${id('lampglow')})`} />}
          <path d={`M${x},${SCENE_ROAD_Y - 10} L${x},${SCENE_ROAD_Y - 40} Q${x},${SCENE_ROAD_Y - 46} ${x + 6},${SCENE_ROAD_Y - 46} L${x + 12},${SCENE_ROAD_Y - 46}`} fill="none" stroke={pal.nearShade} strokeWidth="2" strokeLinecap="round" />
          <rect x={x + 8} y={SCENE_ROAD_Y - 48} width="8" height="4" fill={pal.lamp || pal.nearLight} />
        </g>
      ))}
      {/* 뒤 차선(왼쪽으로 가는 차)을 먼저 그려, 엇갈릴 때 앞 차선 차가 위에 오게 한다 */}
      {[...SCENE_CARS.map((car, i) => ({ car, i }))].sort((a, b) => a.car.dir - b.car.dir).map(({ car, i }) => {
        const lane = car.dir > 0 ? 3 : -4;
        return (
          <g key={`car-${i}`}>
            <g transform={`translate(0 ${SCENE_ROAD_Y + lane}) scale(${car.dir} 1)`}>
              <rect x={-car.len / 2} y={-5} width={car.len} height="6" rx="2" fill={pal.nearShade} />
              <rect x={-car.len / 2 + 5} y={-9} width={car.len * 0.55} height="5" rx="2" fill={pal.nearShade} />
              {pal.headlight && <path d={`M${car.len / 2},-3 L${car.len / 2 + 34},-9 L${car.len / 2 + 34},4 Z`} fill={pal.headlight} opacity="0.2" />}
              {pal.headlight && <rect x={car.len / 2 - 1.5} y={-3.5} width="2.5" height="2.5" fill={pal.headlight} />}
              {pal.headlight && <rect x={-car.len / 2 - 1} y={-3.5} width="2.5" height="2.5" fill="#ff4d3d" />}
            </g>
            <animateMotion dur={`${car.dur}s`} begin={`${car.delay}s`} repeatCount="indefinite" path="M-60,0 L1660,0" keyPoints={car.dir > 0 ? '0;1' : '1;0'} keyTimes="0;1" calcMode="linear" />
          </g>
        );
      })}
      {/* 철길과 화물열차 실루엣 */}
      <path d={`M0,${SCENE_RAIL_Y + 6} L1600,${SCENE_RAIL_Y + 6}`} stroke={pal.rail} strokeWidth="5" strokeDasharray="3 9" />
      <path d={`M0,${SCENE_RAIL_Y + 3.5} L1600,${SCENE_RAIL_Y + 3.5} M0,${SCENE_RAIL_Y + 8.5} L1600,${SCENE_RAIL_Y + 8.5}`} stroke={pal.rail} strokeWidth="1.4" />
      {(() => {
        let offset = 0;
        const trainLength = SCENE_TRAIN.cars.reduce((sum, w) => sum + w + 5, 0);
        const cars = SCENE_TRAIN.cars.map((w, i) => {
          // 기관차(i=0)를 진행 방향(오른쪽) 맨 앞에 두고 화차를 그 뒤로 잇는다 → 전조등이 앞을 비춘다
          const x = trainLength - offset - w; offset += w + 5;
          const h = i === 0 ? 16 : 13;
          const y = SCENE_RAIL_Y + 3 - h;
          return (
            <g key={`train-${i}`}>
              <rect x={x} y={y} width={w} height={h} rx={i % 3 === 1 ? 6 : 1.5} fill={pal.nearShade} />
              {i === 0 && <rect x={x + w - 16} y={y - 6} width="14" height="7" fill={pal.nearShade} />}
              {i === 0 && pal.headlight && <path d={`M${x + w},${y + 4} L${x + w + 46},${y - 6} L${x + w + 46},${y + 14} Z`} fill={pal.headlight} opacity="0.22" />}
              <circle cx={x + 7} cy={SCENE_RAIL_Y + 3.5} r="2.4" fill={pal.rail} />
              <circle cx={x + w - 7} cy={SCENE_RAIL_Y + 3.5} r="2.4" fill={pal.rail} />
            </g>
          );
        });
        return (
          <g>
            <g transform={`translate(${-offset} 0)`}>{cars}</g>
            <animateMotion dur={`${SCENE_TRAIN.dur}s`} begin={`${SCENE_TRAIN.delay}s`} repeatCount="indefinite" path={`M-60,0 L${1600 + offset + 60},0`} />
          </g>
        );
      })()}

      {/* 새 떼 (낮·아침)와 비행기 */}
      {pal.bird && SCENE_BIRD_FLOCKS.slice(0, compact ? 1 : 2).map((flock, fi) => (
        <g key={`flock-${fi}`} fill="none" stroke={pal.bird} strokeWidth="1.6" strokeLinecap="round">
          {flock.birds.map(([bx, by], bi) => (
            <path key={bi} className="skyline-bird" d={`M${bx - 6},${by} Q${bx - 3},${by - 5} ${bx},${by} Q${bx + 3},${by - 5} ${bx + 6},${by}`} style={{ '--flap-duration': `${0.7 + (bi % 3) * 0.12}s`, '--flap-delay': `-${(bi * 0.23) % 0.8}s` }} />
          ))}
          <animateMotion dur={`${flock.dur}s`} begin={`${flock.delay}s`} repeatCount="indefinite" path={flock.path} />
        </g>
      ))}
      {!compact && (
        <g>
          {pal.contrail && <rect x="-230" y="-1.2" width="212" height="2.4" rx="1.2" fill={`url(#${id('contrail')})`} />}
          <path d="M-14,0 L8,0 Q14,0 14,2.5 L-14,2.5 Z M-8,0 L-6,-7 L-2,-7 L2,0 Z M-10,2.5 L-14,8 L-9,8 L-3,2.5 Z" fill={pal.plane} />
          {!pal.contrail && <circle className="skyline-beacon" cx="-12" cy="0" r="1.8" fill="#ff5a4a" />}
          {!pal.contrail && <circle className="skyline-beacon" cx="-4" cy="-7" r="1.8" fill="#8bf0a0" style={{ '--beacon-delay': '-1.2s' }} />}
          <animateMotion dur={`${SCENE_PLANE.dur}s`} begin={`${SCENE_PLANE.delay}s`} repeatCount="indefinite" path={SCENE_PLANE.path} rotate="auto" />
        </g>
      )}
    </svg>
  );
};

const SHEET_THEMES = {
  flow: {
    selection: 'selection:bg-cyan-100',
    pageBg: 'bg-cyan-50',
    pageTint: 'bg-cyan-100/40',
    headerBorder: 'border-t-cyan-600',
    headerIcon: 'text-cyan-600',
    focusRing: 'focus:ring-cyan-500',
    primaryButton: 'bg-cyan-600 hover:bg-cyan-700',
    primaryButtonSoft: 'bg-cyan-600 hover:bg-cyan-500',
    accentText: 'text-cyan-400',
  },
  samplinglog: {
    selection: 'selection:bg-orange-100',
    pageBg: 'bg-orange-50',
    pageTint: 'bg-orange-100/40',
    headerBorder: 'border-t-orange-600',
    headerIcon: 'text-orange-600',
    focusRing: 'focus:ring-orange-500',
    primaryButton: 'bg-orange-600 hover:bg-orange-700',
    primaryButtonSoft: 'bg-orange-600 hover:bg-orange-500',
    accentText: 'text-orange-400',
  },
  dust: {
    selection: 'selection:bg-emerald-100',
    pageBg: 'bg-emerald-50',
    pageTint: 'bg-emerald-100/40',
    headerBorder: 'border-t-emerald-600',
    headerIcon: 'text-emerald-600',
    focusRing: 'focus:ring-emerald-500',
    primaryButton: 'bg-emerald-600 hover:bg-emerald-700',
    primaryButtonSoft: 'bg-emerald-600 hover:bg-emerald-500',
    accentText: 'text-emerald-400',
  },
  metal: {
    selection: 'selection:bg-yellow-100',
    pageBg: 'bg-yellow-50',
    pageTint: 'bg-yellow-100/50',
    headerBorder: 'border-t-yellow-500',
    headerIcon: 'text-yellow-600',
    focusRing: 'focus:ring-yellow-500',
    primaryButton: 'bg-yellow-600 hover:bg-yellow-700',
    primaryButtonSoft: 'bg-yellow-600 hover:bg-yellow-500',
    accentText: 'text-yellow-400',
  },
  mercury: {
    selection: 'selection:bg-blue-100',
    pageBg: 'bg-blue-50',
    pageTint: 'bg-blue-100/50',
    headerBorder: 'border-t-blue-600',
    headerIcon: 'text-blue-600',
    focusRing: 'focus:ring-blue-500',
    primaryButton: 'bg-blue-600 hover:bg-blue-700',
    primaryButtonSoft: 'bg-blue-600 hover:bg-blue-500',
    accentText: 'text-blue-300',
  },
  pahs: {
    selection: 'selection:bg-purple-100',
    pageBg: 'bg-purple-50',
    pageTint: 'bg-purple-100/50',
    headerBorder: 'border-t-purple-600',
    headerIcon: 'text-purple-600',
    focusRing: 'focus:ring-purple-500',
    primaryButton: 'bg-purple-600 hover:bg-purple-700',
    primaryButtonSoft: 'bg-purple-600 hover:bg-purple-500',
    accentText: 'text-purple-300',
  },
};

const THEME_OVERRIDE_CSS = `
.theme-metal [class*="bg-emerald-50"], .theme-metal [class*="bg-teal-50"], .theme-metal [class*="bg-cyan-50"] { background-color: #fbfaf6 !important; }
.theme-metal [class*="bg-emerald-100"], .theme-metal [class*="bg-teal-100"], .theme-metal [class*="bg-cyan-100"] { background-color: #f6f1df !important; }
.theme-metal [class*="bg-emerald-200"], .theme-metal [class*="bg-teal-200"], .theme-metal [class*="bg-cyan-200"] { background-color: #fde68a !important; }
.theme-metal [class*="bg-emerald-600"], .theme-metal [class*="bg-teal-600"], .theme-metal [class*="bg-cyan-600"] { background-color: #ca8a04 !important; }
.theme-metal [class*="bg-emerald-700"], .theme-metal [class*="bg-teal-700"], .theme-metal [class*="bg-cyan-700"] { background-color: #a16207 !important; }
.theme-metal [class*="text-emerald-"], .theme-metal [class*="text-teal-"], .theme-metal [class*="text-cyan-"] { color: #a16207 !important; }
.theme-metal [class*="border-emerald-"], .theme-metal [class*="border-teal-"], .theme-metal [class*="border-cyan-"] { border-color: #facc15 !important; }
.theme-metal [class*="focus:ring-emerald-500"]:focus { --tw-ring-color: rgb(234 179 8 / 0.55) !important; }

.theme-mercury [class*="bg-emerald-50"], .theme-mercury [class*="bg-teal-50"], .theme-mercury [class*="bg-cyan-50"] { background-color: #f7f9fc !important; }
.theme-mercury [class*="bg-emerald-100"], .theme-mercury [class*="bg-teal-100"], .theme-mercury [class*="bg-cyan-100"] { background-color: #edf2f9 !important; }
.theme-mercury [class*="bg-emerald-200"], .theme-mercury [class*="bg-teal-200"], .theme-mercury [class*="bg-cyan-200"] { background-color: #bfdbfe !important; }
.theme-mercury [class*="bg-emerald-600"], .theme-mercury [class*="bg-teal-600"], .theme-mercury [class*="bg-cyan-600"] { background-color: #2563eb !important; }
.theme-mercury [class*="bg-emerald-700"], .theme-mercury [class*="bg-teal-700"], .theme-mercury [class*="bg-cyan-700"] { background-color: #1d4ed8 !important; }
.theme-mercury [class*="text-emerald-"], .theme-mercury [class*="text-teal-"], .theme-mercury [class*="text-cyan-"] { color: #1d4ed8 !important; }
.theme-mercury [class*="border-emerald-"], .theme-mercury [class*="border-teal-"], .theme-mercury [class*="border-cyan-"] { border-color: #60a5fa !important; }
.theme-mercury [class*="focus:ring-emerald-500"]:focus { --tw-ring-color: rgb(59 130 246 / 0.55) !important; }

.theme-pahs [class*="bg-emerald-50"], .theme-pahs [class*="bg-teal-50"], .theme-pahs [class*="bg-cyan-50"] { background-color: #faf8fc !important; }
.theme-pahs [class*="bg-emerald-100"], .theme-pahs [class*="bg-teal-100"], .theme-pahs [class*="bg-cyan-100"] { background-color: #f1ebf8 !important; }
.theme-pahs [class*="bg-emerald-200"], .theme-pahs [class*="bg-teal-200"], .theme-pahs [class*="bg-cyan-200"] { background-color: #e9d5ff !important; }
.theme-pahs [class*="bg-emerald-600"], .theme-pahs [class*="bg-teal-600"], .theme-pahs [class*="bg-cyan-600"] { background-color: #9333ea !important; }
.theme-pahs [class*="bg-emerald-700"], .theme-pahs [class*="bg-teal-700"], .theme-pahs [class*="bg-cyan-700"] { background-color: #7e22ce !important; }
.theme-pahs [class*="text-emerald-"], .theme-pahs [class*="text-teal-"], .theme-pahs [class*="text-cyan-"] { color: #7e22ce !important; }
.theme-pahs [class*="border-emerald-"], .theme-pahs [class*="border-teal-"], .theme-pahs [class*="border-cyan-"] { border-color: #c084fc !important; }
.theme-pahs [class*="focus:ring-emerald-500"]:focus { --tw-ring-color: rgb(168 85 247 / 0.55) !important; }
`;

const STORAGE_KEYS = {
  nozzles: 'dust-sampling.nozzles.v1',
  samplers: 'dust-sampling.samplers.v1',
  pitotPresets: 'dust-sampling.pitot-presets.v1',
  nozzlesUserPrefix: 'dust-sampling.nozzles.user.v1.',
  groupNozzles: 'dust-sampling.group-nozzles.v1',
  groupNozzlesUserPrefix: 'dust-sampling.group-nozzles.user.v1.',
  samplersUserPrefix: 'dust-sampling.samplers.user.v1.',
  gasSamplers: 'dust-sampling.gas-samplers.v1',
  gasSamplersUserPrefix: 'dust-sampling.gas-samplers.user.v1.',
  pitotUserPrefix: 'dust-sampling.pitot-presets.user.v1.',
  secureUsers: 'dust-sampling.secure-users.v2',
  activeUser: 'dust-sampling.secure-active-user.v2',
  lastLoginId: 'dust-sampling.last-login-id.v1',
  sessionAuth: 'dust-sampling.session-auth.v1',
  skyPreviewMode: 'dust-sampling.sky-preview-mode.v1',
  cloudSyncPrefix: 'dust-sampling.cloud-sync.v1.',
  samplingLogDraftPrefix: 'dust-sampling.samplinglog-draft.v1.',
  sheetContext: 'dust-sampling.sheet-context.v1',
  pendingReport: 'dust-sampling.pending-report.v1',
};

// 통합 시료채취기록부를 별도 창(window.open)으로 띄웠을 때 true (#sheet=samplinglog&popup=1)
const IS_POPUP_WINDOW = typeof window !== 'undefined'
  && new URLSearchParams(window.location.hash.replace(/^#/, '')).get('popup') === '1';

// 메뉴에서 새로 연 종합 기록부 창 (불러오기 아님): 다른 기록부처럼 빈 기록지로 시작한다
const OPENED_SAMPLING_LOG_FRESH = IS_POPUP_WINDOW && (() => {
  const params = new URLSearchParams(window.location.hash.replace(/^#/, ''));
  return params.get('sheet') === 'samplinglog' && !params.get('load');
})();

const getSamplingLogDraftStorageKey = (userId) => (
  `${STORAGE_KEYS.samplingLogDraftPrefix}${encodeURIComponent(String(userId || '').trim().toLowerCase())}`
);

const getUserNozzlesStorageKey = (userId) => (
  `${STORAGE_KEYS.nozzlesUserPrefix}${encodeURIComponent(String(userId || '').trim().toLowerCase())}`
);

const getUserSamplersStorageKey = (userId) => (
  `${STORAGE_KEYS.samplersUserPrefix}${encodeURIComponent(String(userId || '').trim().toLowerCase())}`
);

const getUserGasSamplersStorageKey = (userId) => (
  `${STORAGE_KEYS.gasSamplersUserPrefix}${encodeURIComponent(String(userId || '').trim().toLowerCase())}`
);

const getUserPitotStorageKey = (userId) => (
  `${STORAGE_KEYS.pitotUserPrefix}${encodeURIComponent(String(userId || '').trim().toLowerCase())}`
);

const getUserCloudSyncStorageKey = (userId) => (
  `${STORAGE_KEYS.cloudSyncPrefix}${encodeURIComponent(String(userId || '').trim().toLowerCase())}`
);

const sanitizeCloudSyncConfig = (value) => {
  if (!value || typeof value !== 'object' || Array.isArray(value)) return null;
  return {
    token: typeof value.token === 'string' ? value.token.trim() : '',
    gistId: typeof value.gistId === 'string' ? value.gistId.trim() : '',
  };
};

const persistSessionAuth = (userId, password) => {
  try {
    localStorage.setItem(
      STORAGE_KEYS.sessionAuth,
      JSON.stringify({ userId: String(userId || ''), password: String(password || '') })
    );
  } catch (error) {
    console.error('로그인 세션 저장 실패:', error);
  }
};

const readSessionAuth = () => {
  try {
    const raw = localStorage.getItem(STORAGE_KEYS.sessionAuth);
    if (!raw) return null;
    const parsed = JSON.parse(raw);
    const userId = String(parsed?.userId || '').trim();
    const password = typeof parsed?.password === 'string' ? parsed.password : '';
    if (!userId || !password) return null;
    return { userId, password };
  } catch (error) {
    console.error('로그인 세션 로딩 실패:', error);
    return null;
  }
};

const clearSessionAuth = () => {
  try {
    localStorage.removeItem(STORAGE_KEYS.sessionAuth);
  } catch (error) {
    console.error('로그인 세션 삭제 실패:', error);
  }
};

const cloneDefaultNozzles = () => DEFAULT_NOZZLE_SET.map(item => ({ ...item }));
const cloneDefaultSamplers = () => DEFAULT_SAMPLERS.map(item => ({ ...item }));
const cloneDefaultGasSamplers = () => DEFAULT_GAS_SAMPLERS.map(item => ({ ...item }));

// 통합 시료채취기록부의 가스 샘플러 프리셋 (번호, S/N, Yd)
const sanitizeGasSamplers = (value) => {
  if (!Array.isArray(value)) return null;
  return value
    .filter(item => item && typeof item === 'object')
    .map((item, idx) => {
      const parsedId = parseInt(item.id, 10);
      return {
        id: Number.isFinite(parsedId) ? parsedId : idx + 1,
        sn: item.sn === undefined || item.sn === null ? '' : String(item.sn),
        yd: item.yd === undefined || item.yd === null ? '' : String(item.yd),
        calDate: item.calDate === undefined || item.calDate === null ? '' : String(item.calDate).trim(),
      };
    });
};

const getStoredGasSamplersForUser = (userId) => {
  const normalizedId = String(userId || '').trim().toLowerCase();
  const canUseScoped = normalizedId && normalizedId !== 'user' && normalizedId !== 'user1';
  const tryKey = (key) => {
    try {
      const raw = localStorage.getItem(key);
      if (!raw) return null;
      const parsed = sanitizeGasSamplers(JSON.parse(raw));
      return parsed && parsed.length > 0 ? parsed : null;
    } catch (error) {
      console.error('가스 샘플러 프리셋 로딩 실패:', error);
      return null;
    }
  };
  if (canUseScoped) {
    const scoped = tryKey(getUserGasSamplersStorageKey(userId));
    if (scoped) return scoped;
  }
  return tryKey(STORAGE_KEYS.gasSamplers) || cloneDefaultGasSamplers();
};
const cloneDefaultPitotPresets = () => DEFAULT_PITOT_PRESETS.map(item => ({ ...item }));

// ── 교정 이력: 샘플러·가스 샘플러·피토관의 교정값이 바뀌어 저장될 때마다 한 줄씩 쌓는다 (지우지 않는다) ──
const PRESET_HISTORY_STORAGE_KEY = 'dust-sampling.preset-history.v1';
const PRESET_HISTORY_KINDS = {
  sampler: { label: '샘플러', name: (item) => item.name || `샘플러 ${item.id}`, detail: (item) => `Yd ${item.yd || '-'} / ΔH@ ${item.dHAt || '-'}`, filled: (item) => !!(item.yd || item.dHAt) },
  gas: { label: '가스 샘플러', name: (item) => `가스 샘플러 ${item.id}${item.sn ? ` (${item.sn})` : ''}`, detail: (item) => `Yd ${item.yd || '-'}`, filled: (item) => !!item.yd },
  pitot: {
    label: '피토관', name: (item) => item.sn || `피토관 ${item.id}`,
    detail: (item) => `Cp ${item.factor || '-'}${(item.cpByVelocity || []).length ? ` / 유속별 ${(item.cpByVelocity || []).map((row) => `${row.v}:${row.cp}`).join(', ')}` : ''}`,
    filled: (item) => !!(item.sn && item.factor),
  },
};
const readPresetHistory = () => {
  try {
    const list = JSON.parse(localStorage.getItem(PRESET_HISTORY_STORAGE_KEY));
    return Array.isArray(list) ? list.filter((row) => row && typeof row === 'object') : [];
  } catch { return []; }
};
const presetHistoryRow = (kind, item, at, note = '') => {
  const def = PRESET_HISTORY_KINDS[kind];
  return { at, kind, id: item.id, name: def.name(item), calDate: item.calDate || '', detail: def.detail(item), note };
};
// 장비마다 이력의 마지막 줄과 견줘, 처음 보는 장비이거나 교정일·교정값이 달라진 장비만 이력에 더한다
const appendPresetHistory = (kind, _prevList, nextList, user = '', note = '') => {
  const def = PRESET_HISTORY_KINDS[kind];
  const history = readPresetHistory();
  const now = new Date().toISOString();
  const added = [];
  (nextList || []).forEach((item) => {
    if (!def.filled(item)) return;
    const last = history.filter((row) => row.kind === kind && String(row.id) === String(item.id)).pop();
    const row = presetHistoryRow(kind, item, now, last ? '' : note);
    if (last && last.name === row.name && last.calDate === row.calDate && last.detail === row.detail) return;
    added.push({ ...row, user });
  });
  if (added.length === 0) return history;
  const next = [...history, ...added];
  try { localStorage.setItem(PRESET_HISTORY_STORAGE_KEY, JSON.stringify(next)); } catch { /* 이력 저장 실패는 프리셋 저장을 막지 않는다 */ }
  return next;
};
const readStoredList = (key) => { try { const v = JSON.parse(localStorage.getItem(key)); return Array.isArray(v) ? v : []; } catch { return []; } };

const sanitizeNozzleSet = (value) => {
  if (!Array.isArray(value)) return null;
  return value
    .filter(item => item && typeof item === 'object')
    .map(item => ({
      num: item.num === undefined || item.num === null ? '' : String(item.num),
      d: item.d === undefined || item.d === null ? '' : String(item.d),
    }));
};

const sanitizeSamplers = (value) => {
  if (!Array.isArray(value)) return null;
  return value
    .filter(item => item && typeof item === 'object')
    .map((item, idx) => {
      const parsedId = parseInt(item.id, 10);
      return {
        id: Number.isFinite(parsedId) ? parsedId : idx + 1,
        name: typeof item.name === 'string' && item.name.trim() ? item.name : `샘플러 ${idx + 1}호기`,
        yd: item.yd === undefined || item.yd === null ? '' : String(item.yd),
        dHAt: item.dHAt === undefined || item.dHAt === null ? '' : String(item.dHAt),
        calDate: item.calDate === undefined || item.calDate === null ? '' : String(item.calDate).trim(),
      };
    });
};

const sanitizePitotPresets = (value) => {
  if (!Array.isArray(value)) return null;
  return value
    .filter(item => item && typeof item === 'object')
    .map((item, idx) => {
      const parsedId = parseInt(item.id, 10);
      const legacyName = typeof item.name === 'string' ? item.name.trim() : '';
      const snRaw = item.sn === undefined || item.sn === null ? '' : String(item.sn).trim();
      const cpByVelocity = Array.isArray(item.cpByVelocity)
        ? item.cpByVelocity
            .filter(row => row && typeof row === 'object')
            .map(row => ({
              v: row.v === undefined || row.v === null ? '' : String(row.v),
              cp: row.cp === undefined || row.cp === null ? '' : String(row.cp),
            }))
            .filter(row => row.v !== '' || row.cp !== '')
        : [];
      const text = (v) => (v === undefined || v === null ? '' : String(v).trim());
      return {
        id: Number.isFinite(parsedId) ? parsedId : idx + 1,
        sn: snRaw || legacyName || '',
        usage: text(item.usage),
        length: text(item.length),
        calDate: text(item.calDate),
        factor: item.factor === undefined || item.factor === null ? '' : String(item.factor),
        cpByVelocity,
      };
    });
};

// 유속별 피토계수 표: 각 행의 유속은 그 구간의 상한 (예: 4.95 m/s → 0~4.95 m/s 구간에 0.816 적용)
// 표의 마지막 유속을 넘으면 마지막 행의 Cp를 적용
const getPitotVelocityRows = (preset) => (
  Array.isArray(preset?.cpByVelocity)
    ? preset.cpByVelocity
        .map(row => ({ v: parseFloat(row.v), cp: parseFloat(row.cp) }))
        .filter(row => Number.isFinite(row.v) && row.v > 0 && Number.isFinite(row.cp) && row.cp > 0)
        .sort((a, b) => a.v - b.v)
    : []
);

const resolvePitotCpByVelocity = (preset, velocity) => {
  const nominal = parseFloat(preset?.factor);
  const rows = getPitotVelocityRows(preset);
  if (rows.length === 0) return Number.isFinite(nominal) ? nominal : NaN;
  if (!Number.isFinite(velocity) || velocity <= 0) return Number.isFinite(nominal) ? nominal : rows[0].cp;
  const band = rows.find((row) => velocity <= row.v + 1e-9);
  return band ? band.cp : rows[rows.length - 1].cp;
};

const formatPitotCp = (value) => (Number.isFinite(value) ? String(Math.round(value * 10000) / 10000) : '');

const getGroupNozzlesStorageKey = (userId) => {
  const normalizedId = String(userId || '').trim().toLowerCase();
  return normalizedId && normalizedId !== 'user' && normalizedId !== 'user1'
    ? `${STORAGE_KEYS.groupNozzlesUserPrefix}${encodeURIComponent(normalizedId)}`
    : STORAGE_KEYS.groupNozzles;
};
const sanitizeGroupNozzleSets = (value) => {
  const base = cloneDefaultGroupNozzleSets();
  if (!value || typeof value !== 'object') return base;
  ['mercury', 'pahs'].forEach((group) => {
    const parsed = sanitizeNozzleSet(value[group]);
    if (parsed) base[group] = parsed;
  });
  return base;
};
const getStoredGroupNozzlesForUser = (userId) => {
  try {
    const raw = localStorage.getItem(getGroupNozzlesStorageKey(userId));
    if (raw) return sanitizeGroupNozzleSets(JSON.parse(raw));
  } catch (error) {
    console.error('수은·PAHs 노즐 프리셋 로딩 실패:', error);
  }
  return cloneDefaultGroupNozzleSets();
};

const getStoredNozzlesForUser = (userId) => {
  const normalizedId = String(userId || '').trim().toLowerCase();
  const canUseScoped = normalizedId && normalizedId !== 'user' && normalizedId !== 'user1';
  if (!canUseScoped) {
    try {
      const legacyRaw = localStorage.getItem(STORAGE_KEYS.nozzles);
      if (legacyRaw) {
        const legacyParsed = sanitizeNozzleSet(JSON.parse(legacyRaw));
        if (legacyParsed && legacyParsed.length > 0) return legacyParsed;
      }
    } catch (error) {
      console.error('기존 노즐 프리셋 로딩 실패:', error);
    }
    return cloneDefaultNozzles();
  }

  const scopedKey = getUserNozzlesStorageKey(userId);
  try {
    const scopedRaw = localStorage.getItem(scopedKey);
    if (scopedRaw) {
      const scopedParsed = sanitizeNozzleSet(JSON.parse(scopedRaw));
      if (scopedParsed && scopedParsed.length > 0) return scopedParsed;
    }
  } catch (error) {
    console.error('사용자 노즐 프리셋 로딩 실패:', error);
  }

  try {
    const legacyRaw = localStorage.getItem(STORAGE_KEYS.nozzles);
    if (legacyRaw) {
      const legacyParsed = sanitizeNozzleSet(JSON.parse(legacyRaw));
      if (legacyParsed && legacyParsed.length > 0) return legacyParsed;
    }
  } catch (error) {
    console.error('기존 노즐 프리셋 로딩 실패:', error);
  }

  return cloneDefaultNozzles();
};

const getStoredSamplersForUser = (userId) => {
  const normalizedId = String(userId || '').trim().toLowerCase();
  const canUseScoped = normalizedId && normalizedId !== 'user' && normalizedId !== 'user1';
  if (!canUseScoped) {
    try {
      const legacyRaw = localStorage.getItem(STORAGE_KEYS.samplers);
      if (legacyRaw) {
        const legacyParsed = sanitizeSamplers(JSON.parse(legacyRaw));
        if (legacyParsed && legacyParsed.length > 0) return legacyParsed;
      }
    } catch (error) {
      console.error('기존 샘플러 프리셋 로딩 실패:', error);
    }
    return cloneDefaultSamplers();
  }

  const scopedKey = getUserSamplersStorageKey(userId);
  try {
    const scopedRaw = localStorage.getItem(scopedKey);
    if (scopedRaw) {
      const scopedParsed = sanitizeSamplers(JSON.parse(scopedRaw));
      if (scopedParsed && scopedParsed.length > 0) return scopedParsed;
    }
  } catch (error) {
    console.error('사용자 샘플러 프리셋 로딩 실패:', error);
  }

  try {
    const legacyRaw = localStorage.getItem(STORAGE_KEYS.samplers);
    if (legacyRaw) {
      const legacyParsed = sanitizeSamplers(JSON.parse(legacyRaw));
      if (legacyParsed && legacyParsed.length > 0) return legacyParsed;
    }
  } catch (error) {
    console.error('기존 샘플러 프리셋 로딩 실패:', error);
  }

  return cloneDefaultSamplers();
};

const getStoredPitotPresetsForUser = (userId) => {
  const normalizedId = String(userId || '').trim().toLowerCase();
  const canUseScoped = normalizedId && normalizedId !== 'user' && normalizedId !== 'user1';
  if (!canUseScoped) {
    try {
      const legacyRaw = localStorage.getItem(STORAGE_KEYS.pitotPresets);
      if (legacyRaw) {
        const legacyParsed = sanitizePitotPresets(JSON.parse(legacyRaw));
        if (legacyParsed && legacyParsed.length > 0) return legacyParsed;
      }
    } catch (error) {
      console.error('기존 피토 프리셋 로딩 실패:', error);
    }
    return cloneDefaultPitotPresets();
  }

  const scopedKey = getUserPitotStorageKey(userId);
  try {
    const scopedRaw = localStorage.getItem(scopedKey);
    if (scopedRaw) {
      const scopedParsed = sanitizePitotPresets(JSON.parse(scopedRaw));
      if (scopedParsed && scopedParsed.length > 0) return scopedParsed;
    }
  } catch (error) {
    console.error('사용자 피토 프리셋 로딩 실패:', error);
  }

  try {
    const legacyRaw = localStorage.getItem(STORAGE_KEYS.pitotPresets);
    if (legacyRaw) {
      const legacyParsed = sanitizePitotPresets(JSON.parse(legacyRaw));
      if (legacyParsed && legacyParsed.length > 0) return legacyParsed;
    }
  } catch (error) {
    console.error('기존 피토 프리셋 로딩 실패:', error);
  }

  return cloneDefaultPitotPresets();
};

const sanitizeSecureUsers = (value) => {
  if (!value || typeof value !== 'object' || Array.isArray(value)) return null;
  const entries = Object.entries(value);
  if (entries.length === 0) return {};

  const result = {};
  entries.forEach(([user, record]) => {
    const userName = String(user || '').trim();
    if (!userName) return;
    if (!record || typeof record !== 'object' || Array.isArray(record)) return;

    const normalized = {
      salt: typeof record.salt === 'string' ? record.salt : '',
      verifierIv: typeof record.verifierIv === 'string' ? record.verifierIv : '',
      verifierCipher: typeof record.verifierCipher === 'string' ? record.verifierCipher : '',
      dataIv: typeof record.dataIv === 'string' ? record.dataIv : '',
      dataCipher: typeof record.dataCipher === 'string' ? record.dataCipher : '',
      updatedAt: typeof record.updatedAt === 'string' ? record.updatedAt : '',
      hintQuestion: typeof record.hintQuestion === 'string' ? record.hintQuestion : '',
      hintAnswerHash: typeof record.hintAnswerHash === 'string' ? record.hintAnswerHash : '',
      avatarUrl: sanitizeAvatarUrl(record.avatarUrl),
      nickname: typeof record.nickname === 'string' ? record.nickname : '',
    };

    if (!normalized.salt || !normalized.verifierIv || !normalized.verifierCipher || !normalized.dataIv || !normalized.dataCipher) return;
    result[userName] = normalized;
  });

  return result;
};

const PASSWORD_RULE = /^(?=.*[A-Za-z])(?=.*\d)(?=.*[^A-Za-z0-9\s])\S{8,64}$/;
const normalizeHintText = (value) => String(value || '').trim().toLowerCase();
const ADMIN_USER_ID = 'dightkdtjr';
const ADMIN_PASSWORD_SHA256 = 'd3a8457d2972ea9b7c9973753a47336eec669a2479d3286a39b622dc53659bdb';
const BETA_USER_ID = 'user';
const BETA_USER_PASSWORD = 'beta-user';
const LEGACY_BETA_USER_ID = 'user1';
const LEGACY_BETA_USER_PASSWORD = 'beta-user1';

const getImpingerStageCount = (sheetId) => (sheetId === 'pahs' ? 6 : 4);

const normalizeImpingers = (value, stageCount) => {
  const targetCount = Number.isFinite(stageCount) && stageCount > 0 ? stageCount : 4;
  const source = Array.isArray(value) ? value : [];
  return Array.from({ length: targetCount }, (_, idx) => {
    const current = source[idx] || {};
    return {
      id: idx + 1,
      initial: current.initial === undefined || current.initial === null ? '' : String(current.initial),
      final: current.final === undefined || current.final === null ? '' : String(current.final),
    };
  });
};
const CLOUD_SYNC_FILE_NAME = 'stackpilot-cloud.json';
const CLOUD_SYNC_VERSION = 1;

const textEncoder = new TextEncoder();
const textDecoder = new TextDecoder();

const sha256Hex = async (text) => {
  const digest = await crypto.subtle.digest('SHA-256', textEncoder.encode(String(text || '')));
  return Array.from(new Uint8Array(digest)).map((b) => b.toString(16).padStart(2, '0')).join('');
};
const isAdminPassword = async (password) => Boolean(password) && (await sha256Hex(password)) === ADMIN_PASSWORD_SHA256;

const bytesToBase64 = (bytes) => {
  const arr = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
  let binary = '';
  const chunkSize = 0x8000;
  for (let i = 0; i < arr.length; i += chunkSize) {
    binary += String.fromCharCode(...arr.subarray(i, i + chunkSize));
  }
  return btoa(binary);
};

const base64ToBytes = (b64) => {
  const binary = atob(b64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) {
    bytes[i] = binary.charCodeAt(i);
  }
  return bytes;
};

const deriveVaultKey = async (password, saltB64) => {
  const baseKey = await crypto.subtle.importKey('raw', textEncoder.encode(password), 'PBKDF2', false, ['deriveKey']);
  return crypto.subtle.deriveKey(
    {
      name: 'PBKDF2',
      salt: base64ToBytes(saltB64),
      iterations: 120000,
      hash: 'SHA-256',
    },
    baseKey,
    { name: 'AES-GCM', length: 256 },
    false,
    ['encrypt', 'decrypt']
  );
};

const encryptTextWithKey = async (key, plainText) => {
  const iv = crypto.getRandomValues(new Uint8Array(12));
  const encrypted = await crypto.subtle.encrypt(
    { name: 'AES-GCM', iv },
    key,
    textEncoder.encode(plainText)
  );
  return {
    iv: bytesToBase64(iv),
    cipher: bytesToBase64(new Uint8Array(encrypted)),
  };
};

const decryptTextWithKey = async (key, ivB64, cipherB64) => {
  const decrypted = await crypto.subtle.decrypt(
    { name: 'AES-GCM', iv: base64ToBytes(ivB64) },
    key,
    base64ToBytes(cipherB64)
  );
  return textDecoder.decode(decrypted);
};

const encryptJsonWithKey = async (key, value) => {
  return encryptTextWithKey(key, JSON.stringify(value));
};

const decryptJsonWithKey = async (key, ivB64, cipherB64) => {
  const text = await decryptTextWithKey(key, ivB64, cipherB64);
  return JSON.parse(text);
};

const hashHintAnswer = async (answerText) => {
  const normalized = normalizeHintText(answerText);
  if (!normalized) return '';
  const digest = await crypto.subtle.digest('SHA-256', textEncoder.encode(normalized));
  return bytesToBase64(new Uint8Array(digest));
};

// 프로필 이미지는 사용하지 않음 (이름 첫 글자로 표시)
// 표 채우기 공통 도구: 한 칸(td) 안의 입력들, 값 넣기(React가 알아채도록 input 이벤트 발생)
const fillCellFields = (cell) => (cell ? [...cell.querySelectorAll('input, select')].filter((f) => !f.readOnly && !f.disabled && !['checkbox', 'radio', 'file', 'button'].includes(f.type)) : []);
const fillSetFieldValue = (field, value) => {
  const proto = field instanceof HTMLSelectElement ? HTMLSelectElement.prototype : HTMLInputElement.prototype;
  Object.getOwnPropertyDescriptor(proto, 'value').set.call(field, value);
  field.dispatchEvent(new Event(field instanceof HTMLSelectElement ? 'change' : 'input', { bubbles: true }));
};
// 같은 열·같은 자리의 입력을 다른 행에서 찾는다
const fillPeerFinder = (field) => {
  const cell = field.closest('td');
  const row = cell && cell.parentElement;
  const body = row && row.parentElement;
  if (!cell || !body || body.tagName !== 'TBODY') return null;
  const colIdx = [...row.children].indexOf(cell);
  const fieldIdx = fillCellFields(cell).indexOf(field);
  if (fieldIdx < 0) return null;
  const rows = [...body.children];
  const peerOf = (r) => {
    const peer = fillCellFields(r.children[colIdx])[fieldIdx];
    return peer && peer.tagName === field.tagName && peer.type === field.type ? peer : null;
  };
  return { rows, rowIdx: rows.indexOf(row), peerOf };
};
const fillValueInto = async (targets, value) => {
  for (const target of targets) {
    if (!target.isConnected) break;
    if (target.value !== value) fillSetFieldValue(target, value);
    // 한 칸씩 화면에 반영된 뒤 다음 칸을 채운다
    await new Promise((resolve) => setTimeout(resolve, 0));
  }
};

const PALETTE_STORAGE_KEY = 'dust-sampling.palette.v1';

// 기상청 종관(ASOS) 관측 지점: [지점번호, 이름] (가나다순). 기상청 날씨누리 "도시별관측"에 나오는 지점
const KMA_STATIONS = [['105', '강릉'], ['259', '강진군'], ['201', '강화'], ['294', '거제'], ['284', '거창'], ['283', '경주시'], ['185', '고산'], ['172', '고창'], ['251', '고창군'], ['262', '고흥'], ['266', '광양시'], ['156', '광주'], ['279', '구미'], ['140', '군산'], ['238', '금산'], ['253', '김해시'], ['247', '남원'], ['295', '남해'], ['100', '대관령'], ['143', '대구'], ['133', '대전'], ['98', '동두천'], ['106', '동해'], ['165', '목포'], ['273', '문경'], ['288', '밀양'], ['102', '백령도'], ['235', '보령'], ['258', '보성군'], ['226', '보은'], ['271', '봉화'], ['159', '부산'], ['243', '부안'], ['236', '부여'], ['104', '북강릉'], ['296', '북부산'], ['255', '북창원'], ['93', '북춘천'], ['289', '산청'], ['137', '상주'], ['189', '서귀포'], ['129', '서산'], ['108', '서울'], ['181', '서청주'], ['188', '성산'], ['239', '세종'], ['90', '속초'], ['119', '수원'], ['254', '순창군'], ['174', '순천'], ['136', '안동'], ['257', '양산시'], ['202', '양평'], ['168', '여수'], ['252', '영광군'], ['277', '영덕'], ['121', '영월'], ['272', '영주'], ['281', '영천'], ['170', '완도'], ['115', '울릉도'], ['152', '울산'], ['130', '울진'], ['114', '원주'], ['263', '의령군'], ['278', '의성'], ['203', '이천'], ['211', '인제'], ['112', '인천'], ['244', '임실'], ['248', '장수'], ['260', '장흥'], ['146', '전주'], ['217', '정선군'], ['245', '정읍'], ['184', '제주'], ['221', '제천'], ['268', '진도군'], ['192', '진주'], ['155', '창원'], ['232', '천안'], ['95', '철원'], ['276', '청송군'], ['131', '청주'], ['135', '추풍령'], ['101', '춘천'], ['127', '충주'], ['216', '태백'], ['162', '통영'], ['99', '파주'], ['138', '포항'], ['264', '함양군'], ['285', '합천'], ['261', '해남'], ['177', '홍성'], ['212', '홍천'], ['169', '흑산도']];
// 관측 지점의 대략 위치 [위도, 경도]: 검색한 장소에서 가장 가까운 지점을 고르는 데 쓴다
const KMA_STATION_COORDS = { 90: [38.25, 128.56], 93: [37.95, 127.75], 95: [38.15, 127.30], 98: [37.90, 127.06], 99: [37.89, 126.77], 100: [37.68, 128.72], 101: [37.90, 127.74], 102: [37.97, 124.71], 104: [37.80, 128.86], 105: [37.75, 128.89], 106: [37.51, 129.12], 108: [37.57, 126.97], 112: [37.48, 126.62], 114: [37.34, 127.95], 115: [37.48, 130.90], 119: [37.27, 126.99], 121: [37.18, 128.46], 127: [36.97, 127.95], 129: [36.78, 126.49], 130: [36.99, 129.41], 131: [36.64, 127.44], 133: [36.37, 127.37], 135: [36.22, 127.99], 136: [36.57, 128.71], 137: [36.41, 128.16], 138: [36.03, 129.38], 140: [36.00, 126.76], 143: [35.88, 128.65], 146: [35.84, 127.12], 152: [35.58, 129.33], 155: [35.17, 128.57], 156: [35.17, 126.89], 159: [35.10, 129.03], 162: [34.85, 128.44], 165: [34.82, 126.38], 168: [34.74, 127.74], 169: [34.69, 125.45], 170: [34.40, 126.70], 172: [35.35, 126.60], 174: [35.02, 127.37], 177: [36.66, 126.69], 181: [36.64, 127.38], 184: [33.51, 126.53], 185: [33.29, 126.16], 188: [33.39, 126.88], 189: [33.25, 126.57], 192: [35.16, 128.04], 201: [37.71, 126.45], 202: [37.49, 127.49], 203: [37.26, 127.48], 211: [38.06, 128.17], 212: [37.68, 127.88], 216: [37.17, 128.99], 217: [37.38, 128.65], 221: [37.16, 128.19], 226: [36.49, 127.73], 232: [36.76, 127.29], 235: [36.33, 126.56], 236: [36.27, 126.92], 238: [36.11, 127.48], 239: [36.49, 127.24], 243: [35.73, 126.72], 244: [35.61, 127.29], 245: [35.56, 126.87], 247: [35.41, 127.33], 248: [35.66, 127.52], 251: [35.43, 126.70], 252: [35.28, 126.48], 253: [35.23, 128.89], 254: [35.37, 127.13], 255: [35.23, 128.67], 257: [35.31, 129.02], 258: [34.76, 127.21], 259: [34.63, 126.77], 260: [34.69, 126.92], 261: [34.55, 126.57], 262: [34.62, 127.28], 263: [35.32, 128.29], 264: [35.51, 127.75], 266: [34.94, 127.69], 268: [34.47, 126.26], 271: [36.94, 128.91], 272: [36.87, 128.52], 273: [36.63, 128.15], 276: [36.44, 129.04], 277: [36.53, 129.41], 278: [36.36, 128.69], 279: [36.13, 128.32], 281: [35.98, 128.95], 283: [35.82, 129.20], 284: [35.67, 127.91], 285: [35.57, 128.17], 288: [35.49, 128.74], 289: [35.41, 127.88], 294: [34.89, 128.60], 295: [34.82, 127.93], 296: [35.22, 128.96] };
const findNearestKmaStation = (lat, lon) => {
  let best = null;
  Object.entries(KMA_STATION_COORDS).forEach(([id, [sLat, sLon]]) => {
    const km = Math.hypot((lat - sLat) * 111, (lon - sLon) * 111 * Math.cos((lat * Math.PI) / 180));
    if (!best || km < best.km) best = { id, km };
  });
  return best;
};
const WEATHER_STATION_STORAGE_KEY = 'dust-sampling.weather-station.v1';
const WEATHER_PLACE_STORAGE_KEY = 'dust-sampling.weather-place.v1';
// 모바일(웹)에서 기상청 자료를 받아 줄 중계 서버 주소 (relay/kma-relay.js 를 올린 Cloudflare Workers 주소)
const WEATHER_RELAY_STORAGE_KEY = 'dust-sampling.weather-relay.v1';
// 날짜(YYYY-MM-DD)와 시각(HH:MM)으로 조회할 정시를 정한다. 시각이 없으면 오늘은 지금 시각, 지난 날은 12시
const buildKmaObservationTime = (date, timeText) => {
  const now = new Date();
  const pad = (n) => String(n).padStart(2, '0');
  const today = `${now.getFullYear()}-${pad(now.getMonth() + 1)}-${pad(now.getDate())}`;
  const day = /^\d{4}-\d{2}-\d{2}$/.test(String(date || '')) ? String(date) : today;
  if (day > today) return null;
  const m = /^(\d{1,2}):(\d{2})$/.exec(String(timeText || '').trim());
  let hour = m ? Math.min(23, parseInt(m[1], 10)) : (day === today ? now.getHours() : 12);
  if (day === today) hour = Math.min(hour, now.getHours());
  return { day, hour, tm: `${day.replace(/-/g, '.')}.${pad(hour)}:00` };
};
// 현재일기 글자가 없으면 전운량(10분위)으로 하늘 상태를 정한다
const describeKmaWeather = (obs) => {
  if (obs.weather) return obs.weather;
  if (obs.cloud === null || obs.cloud === undefined) return '';
  if (obs.cloud <= 5) return '맑음';
  return obs.cloud <= 8 ? '구름많음' : '흐림';
};

// 엑셀 저장 이름: 폴더 = 날짜_사업장명_#배출구, 파일 = 날짜_기록부명_사업장명_#배출구
const buildExportNames = ({ date, title, company, location, ext }) => {
  const clean = (v) => String(v || '').replace(/[\\/:*?"<>|]/g, '').replace(/\s+/g, ' ').trim();
  // 날짜는 20261029처럼 붙여 쓴다
  const day = (clean(date) || new Date().toISOString().slice(0, 10)).replace(/[-./\s]/g, '');
  const site = clean(company);
  const stackRaw = clean(location);
  const stack = stackRaw ? (stackRaw.startsWith('#') ? stackRaw : `#${stackRaw}`) : '';
  const join = (parts) => parts.filter(Boolean).join('_').replace(/[. ]+$/, '');
  return {
    folder: join([day, site, stack]),
    file: `${join([day, clean(title), site, stack])}.${ext}`,
  };
};

// 프로필 사진: 이 기기에 저장하는 작은 이미지(data URL)만 허용
const AVATAR_MAX_LENGTH = 300000;
const sanitizeAvatarUrl = (value) => (
  typeof value === 'string' && /^data:image\/(png|jpeg|webp);base64,/.test(value) && value.length <= AVATAR_MAX_LENGTH ? value : ''
);
// 고른 사진을 가운데 기준 정사각형으로 잘라 256px JPEG로 줄인다
const resizeImageFileToAvatar = (file) => new Promise((resolve, reject) => {
  const url = URL.createObjectURL(file);
  const img = new Image();
  img.onload = () => {
    try {
      const size = 256;
      const side = Math.min(img.naturalWidth, img.naturalHeight);
      const canvas = document.createElement('canvas');
      canvas.width = size; canvas.height = size;
      const ctx = canvas.getContext('2d');
      ctx.fillStyle = '#ffffff';
      ctx.fillRect(0, 0, size, size);
      ctx.drawImage(img, (img.naturalWidth - side) / 2, (img.naturalHeight - side) / 2, side, side, 0, 0, size, size);
      resolve(canvas.toDataURL('image/jpeg', 0.85));
    } catch (error) { reject(error); } finally { URL.revokeObjectURL(url); }
  };
  img.onerror = () => { URL.revokeObjectURL(url); reject(new Error('image load failed')); };
  img.src = url;
});

// 자바스크립트 부동소수점 오차 방지용 안전한 반올림 함수 (K * dp 오차 보정용)
const roundDH = (val) => {
  return (Math.round((val + Number.EPSILON) * 100) / 100).toFixed(2);
};

const roundFixed = (val, digits = 2) => {
  const num = Number(val);
  if (!Number.isFinite(num)) return '-';
  const factor = Math.pow(10, digits);
  return (Math.round((num + Number.EPSILON) * factor) / factor).toFixed(digits);
};

const formatFlowDisplay = (val) => {
  const num = Number(val);
  if (!Number.isFinite(num)) return '-';
  return num.toFixed(2);
};

const calcExpectedOrificeDH = (kVal, dpVal) => {
  const kNum = parseFloat(kVal);
  const dpNum = parseFloat(dpVal);
  if (!Number.isFinite(kNum) || !Number.isFinite(dpNum)) return '-';
  // 엑셀과 동일하게 원값으로 곱한 뒤 최종값만 소수 둘째자리로 반올림
  return roundDH(kNum * dpNum);
};

const formatSavedDateTime = (value) => {
  const dt = new Date(value);
  if (!Number.isFinite(dt.getTime())) return '-';
  const yyyy = dt.getFullYear();
  const mm = String(dt.getMonth() + 1).padStart(2, '0');
  const dd = String(dt.getDate()).padStart(2, '0');
  const hh = String(dt.getHours()).padStart(2, '0');
  const min = String(dt.getMinutes()).padStart(2, '0');
  const ss = String(dt.getSeconds()).padStart(2, '0');
  return `${yyyy}-${mm}-${dd} ${hh}:${min}:${ss}`;
};

const formatKFactorDisplay = (value) => {
  const num = parseFloat(value);
  return Number.isFinite(num) ? num.toFixed(2) : '-';
};

const PresetSaveIcon = ({ className = 'w-3.5 h-3.5' }) => (
  <svg viewBox="0 0 24 24" fill="none" className={className} aria-hidden="true">
    <path d="M5 4.75h11.2l2.05 2.05V19.25H5V4.75Z" stroke="currentColor" strokeWidth="1.6" />
    <path d="M8 4.75V9h7V4.75" stroke="currentColor" strokeWidth="1.6" />
    <path d="M8 14.25h7v5H8v-5Z" stroke="currentColor" strokeWidth="1.6" />
  </svg>
);

// 기록지 섹션 번호 줄: 머리글 아래에 붙어 스크롤해도 따라오고, 보고 있는 섹션을 표시
const SheetSectionNav = ({ scope }) => {
  const [items, setItems] = useState([]);
  const [active, setActive] = useState('');
  useEffect(() => {
    const heads = Array.from(document.querySelectorAll('h2.ds-sec-title'));
    const next = heads.map((h, i) => {
      const num = h.querySelector('.ds-sec-num')?.textContent?.trim() || String(i + 1);
      const title = (h.textContent || '').replace(num, '').replace(/\s*\(.*$/, '').trim();
      if (!h.id) h.id = `ds-sec-${scope}-${num}`;
      return { id: h.id, num, title };
    });
    const frame = requestAnimationFrame(() => setItems(next));
    if (!heads.length) return () => cancelAnimationFrame(frame);
    // 보고 있는 섹션 = 화면 위쪽(번호 줄 아래)을 지나간 마지막 제목
    let ticking = false;
    const update = () => {
      ticking = false;
      const line = 72;
      let current = heads[0];
      heads.forEach((h) => { if (h.getBoundingClientRect().top <= line) current = h; });
      setActive(current.id);
    };
    const onScroll = () => { if (!ticking) { ticking = true; requestAnimationFrame(update); } };
    window.addEventListener('scroll', onScroll, { passive: true });
    window.addEventListener('resize', onScroll);
    onScroll();
    return () => { cancelAnimationFrame(frame); window.removeEventListener('scroll', onScroll); window.removeEventListener('resize', onScroll); };
  }, [scope]);
  if (items.length < 2) return null;
  return (
    <nav className="ds-secnav" aria-label="기록지 섹션">
      {items.map((item) => (
        <a
          key={item.id}
          href={`#${item.id}`}
          className={`ds-secnav-item${active === item.id ? ' is-active' : ''}`}
          onClick={(e) => { e.preventDefault(); document.getElementById(item.id)?.scrollIntoView({ behavior: 'smooth', block: 'start' }); setActive(item.id); }}
        >
          <span className="ds-secnav-num">{item.num}</span>
          <span className="ds-secnav-title">{item.title}</span>
        </a>
      ))}
    </nav>
  );
};

const SignatureBadge = () => (
  <div className="fixed right-4 bottom-4 z-30 px-3 py-1.5 rounded-full border border-slate-300 bg-white/95 backdrop-blur shadow-sm text-xs font-bold text-slate-700">
    sangseok
  </div>
);

export default function App() {
  // 색 테마: 'default'(기본) | 'earth'(세이지·베이지·테라코타). 이 기기에 저장하고 모든 창에 같이 적용
  const [palette, setPalette] = useState(() => {
    try { return localStorage.getItem(PALETTE_STORAGE_KEY) === 'earth' ? 'earth' : 'default'; } catch { return 'default'; }
  });
  useEffect(() => {
    if (palette === 'earth') document.documentElement.dataset.palette = 'earth';
    else delete document.documentElement.dataset.palette;
  }, [palette]);
  useEffect(() => {
    const onPaletteStorage = (event) => {
      if (event.key === PALETTE_STORAGE_KEY) setPalette(event.newValue === 'earth' ? 'earth' : 'default');
    };
    window.addEventListener('storage', onPaletteStorage);
    return () => window.removeEventListener('storage', onPaletteStorage);
  }, []);
  // ── 기상청 자료 가져오기 ──
  const [weatherStation, setWeatherStation] = useState(() => {
    try { const saved = localStorage.getItem(WEATHER_STATION_STORAGE_KEY); return KMA_STATIONS.some(([id]) => id === saved) ? saved : '108'; } catch { return '108'; }
  });
  const [weatherBusy, setWeatherBusy] = useState(false);
  // 기상청 자료를 받는 길: 설치형은 프로그램이 직접, 웹(모바일)은 설정한 중계 서버를 거친다
  const [weatherRelay, setWeatherRelay] = useState(() => { try { return localStorage.getItem(WEATHER_RELAY_STORAGE_KEY) || ''; } catch { return ''; } });
  const [weatherRelayDraft, setWeatherRelayDraft] = useState(null);
  const saveWeatherRelay = (value) => {
    const text = String(value || '').trim().replace(/\/+$/, '');
    if (text && !/^https:\/\/[^\s/]+(\/\S*)?$/.test(text)) { alert('https:// 로 시작하는 주소를 넣어 주세요.'); return; }
    setWeatherRelay(text);
    setWeatherRelayDraft(null);
    try { if (text) localStorage.setItem(WEATHER_RELAY_STORAGE_KEY, text); else localStorage.removeItem(WEATHER_RELAY_STORAGE_KEY); } catch (error) { console.error(error); }
  };
  const isDesktopApp = typeof window !== 'undefined' && Boolean(window.stackpilotDesktop?.fetchWeather);
  const weatherApi = (() => {
    if (isDesktopApp) return window.stackpilotDesktop;
    if (!weatherRelay) return null;
    // 사용자 정의 헤더 없이 GET만 보내면 브라우저가 사전 확인(preflight) 없이 바로 보낸다
    const relayFetch = (url) => fetch(`${weatherRelay}?u=${encodeURIComponent(url)}`);
    return {
      fetchWeather: (tm) => webFetchCityObservations(tm, relayFetch),
      searchPlaces: (query) => webSearchPlaces(query, relayFetch),
      fetchLocalWeather: (dongCode) => webFetchLocalWeather(dongCode, relayFetch),
    };
  })();
  const weatherUnavailableMessage = '모바일에서 기상청 자료를 쓰려면 ⚙ 설정의 "날씨 중계 주소"를 먼저 넣어 주세요.';
  // 검색해서 고른 장소(동 단위). 있으면 오늘 날씨는 그 동네 값으로 가져온다: { label, dongCode, lat, lon }
  const [weatherPlace, setWeatherPlace] = useState(() => {
    try {
      const saved = JSON.parse(localStorage.getItem(WEATHER_PLACE_STORAGE_KEY) || 'null');
      return saved && /^\d{10}$/.test(String(saved.dongCode || '')) ? saved : null;
    } catch { return null; }
  });
  const [weatherPlaceResults, setWeatherPlaceResults] = useState(null); // null: 검색 전, []: 결과 없음
  const saveWeatherPlace = (place) => {
    setWeatherPlace(place);
    try {
      if (place) localStorage.setItem(WEATHER_PLACE_STORAGE_KEY, JSON.stringify(place));
      else localStorage.removeItem(WEATHER_PLACE_STORAGE_KEY);
    } catch (error) { console.error(error); }
  };
  const changeWeatherStation = (stn, { keepPlace = false } = {}) => {
    setWeatherStation(stn);
    // 관측 지점을 직접 고르면 검색해 둔 장소는 지운다
    if (!keepPlace) { saveWeatherPlace(null); setWeatherPlaceResults(null); }
    try { localStorage.setItem(WEATHER_STATION_STORAGE_KEY, stn); } catch (error) { console.error(error); }
  };
  // 고른 지점의 정시 관측값을 받아 온다. 실패하면 알리고 null
  const fetchKmaObservation = async (date, timeText) => {
    if (!weatherApi) { alert(weatherUnavailableMessage); return null; }
    const when = buildKmaObservationTime(date, timeText);
    if (!when) { alert('측정 일자가 오늘보다 뒤라서 관측 자료가 없습니다.'); return null; }
    setWeatherBusy(true);
    try {
      // 장소를 검색해 두었고 측정일이 오늘이면: 그 동네의 현재 날씨(가까운 관측값 + 동네예보의 지금 하늘 상태)
      const now = new Date();
      const today = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}-${String(now.getDate()).padStart(2, '0')}`;
      if (weatherPlace && when.day === today && weatherApi.fetchLocalWeather) {
        try {
          const local = await weatherApi.fetchLocalWeather(weatherPlace.dongCode);
          if (local && (local.sky || local.humidity !== null || local.windDir)) {
            return { name: weatherPlace.label, stn: '', weather: local.sky || '', weatherText: local.sky || '', cloud: null, temp: local.temp, humidity: local.humidity, windDir: local.windDir, windSpeed: local.windSpeed, shownTime: `${local.updated} 현재`, local: true };
          }
        } catch (error) {
          console.error('동네 날씨 가져오기 실패, 관측 지점 자료로 대신합니다:', error);
        }
      }
      const result = await weatherApi.fetchWeather(when.tm);
      const obs = (result?.stations || []).find((item) => item.stn === weatherStation);
      if (!obs) {
        alert(`기상청에 ${when.day} ${when.hour}시 관측 자료가 없습니다. 날짜와 시각을 확인해 주세요.`);
        return null;
      }
      return { ...obs, shownTime: String(result.tm || when.tm), weatherText: describeKmaWeather(obs) };
    } catch (error) {
      console.error('기상청 자료 가져오기 실패:', error);
      alert('기상청 자료를 가져오지 못했습니다. 인터넷 연결을 확인하고 다시 시도해 주세요.');
      return null;
    } finally {
      setWeatherBusy(false);
    }
  };
  // 기온·기압은 현장에서 직접 잰 값을 적으므로 기상청 값으로 채우지 않는다
  const describeKmaObservation = (obs) => [
    obs.local ? `기상청 ${obs.name} · ${obs.shownTime}` : (weatherPlace ? `기상청 ${weatherPlace.label} 부근 · ${obs.shownTime} 관측` : `기상청 ${obs.name} 지점 · ${obs.shownTime} 관측`),
    `날씨 ${obs.weatherText || '(관측값 없음)'} · 습도 ${obs.humidity ?? '-'} % · 풍향 ${obs.windDir || '-'} · 풍속 ${obs.windSpeed ?? '-'} m/s`,
  ].join('\n');
  // 기록부: 날씨만 채운다
  const applyKmaWeatherToSheet = async () => {
    const obs = await fetchKmaObservation(formData.date, formData.samplingStartTime);
    if (!obs) return;
    if (!obs.weatherText) {
      alert(`${describeKmaObservation(obs)}\n\n이 지점·시각에는 날씨 관측값이 없어 날씨 칸을 그대로 두었습니다. 가까운 다른 지점을 골라 보세요.`);
      return;
    }
    setFormData((prev) => ({ ...prev, weather: obs.weatherText }));
    unsavedRef.current = true;
    alert(`${describeKmaObservation(obs)}\n\n날씨 칸을 채웠습니다. 대기 온도와 대기압은 현장에서 잰 값을 적어 주세요.`);
  };
  // 통합시료채취기록부: 그 일차의 기상·습도·풍향·풍속을 채운다 (시각은 그날 가장 이른 채취 시작 시각, 없으면 12시)
  const applyKmaWeatherToLogDay = async (dayIdx) => {
    const day = samplingLogData.days[dayIdx];
    if (!day || !day.date) { alert('먼저 채취일을 적어 주세요.'); return; }
    const starts = Object.values(samplingLogData.rows || {}).filter((row) => row && row.requested && row.dayIndex === dayIdx && /^\d{1,2}:\d{2}$/.test(String(row.start || ''))).map((row) => String(row.start).padStart(5, '0')).sort();
    // 기온·기압은 그날 저장한 기록부에 적어 둔 현장 측정값을 가져온다
    const fromSheets = findSiteWeatherFromSavedReports(day.date);
    const obs = await fetchKmaObservation(day.date, starts[0] || '');
    if (!obs && !fromSheets.temp && !fromSheets.pressure) return;
    if (!obs) {
      setSamplingLogData((prev) => ({ ...prev, days: prev.days.map((item, i) => (i === dayIdx ? { ...item, ...(fromSheets.temp ? { temp: fromSheets.temp } : {}), ...(fromSheets.pressure ? { pressure: fromSheets.pressure } : {}) } : item)) }));
      unsavedRef.current = true;
      alert(`${dayIdx + 1}일차의 기온·기압을 그날 저장한 기록부 값으로 채웠습니다.`);
      return;
    }
    setSamplingLogData((prev) => ({
      ...prev,
      days: prev.days.map((item, i) => (i === dayIdx ? {
        ...item,
        ...(fromSheets.temp ? { temp: fromSheets.temp } : {}),
        ...(fromSheets.pressure ? { pressure: fromSheets.pressure } : {}),
        weather: obs.weatherText || item.weather,
        humidity: obs.humidity === null ? item.humidity : String(obs.humidity),
        windDir: obs.windDir || item.windDir,
        windSpeed: obs.windSpeed === null ? item.windSpeed : String(obs.windSpeed),
      } : item)),
    }));
    unsavedRef.current = true;
    const sheetNote = fromSheets.temp || fromSheets.pressure
      ? `기온·기압은 그날 저장한 기록부 값(${[fromSheets.temp ? `${fromSheets.temp} ℃` : '', fromSheets.pressure ? `${fromSheets.pressure} mmHg` : ''].filter(Boolean).join(', ')})으로 채웠습니다.`
      : '기온과 기압은 사업장·배출구·날짜가 같은 저장 기록부가 없어 그대로 두었습니다.';
    alert(`${describeKmaObservation(obs)}\n\n${dayIdx + 1}일차의 기상·습도·풍향·풍속을 채웠습니다.\n${sheetNote}`);
  };
  // ── 날씨 위치 고르기: 지역 이름을 적어 기상청에서 찾고, 아래로 펼쳐지는 목록에서 고른다 ──
  const [weatherQuery, setWeatherQuery] = useState(null); // null이면 고른 위치 이름을 보여 준다
  const weatherSearchInputRef = useRef(null);
  const weatherPickerRef = useRef(null);
  const searchWeatherPlaces = async () => {
    const text = String(weatherSearchInputRef.current?.value || '').trim();
    if (!text) { alert('찾을 곳의 이름을 적어 주세요. (예: 안산, 고잔동, 온산읍)'); return; }
    if (!weatherApi) { alert(weatherUnavailableMessage); return; }
    setWeatherBusy(true);
    try {
      const list = await weatherApi.searchPlaces(text);
      setWeatherPlaceResults(Array.isArray(list) ? list : []);
    } catch (error) {
      console.error('장소 검색 실패:', error);
      alert('장소를 찾지 못했습니다. 인터넷 연결을 확인하고 다시 시도해 주세요.');
    } finally {
      setWeatherBusy(false);
    }
  };
  const chooseWeatherPlace = (item) => {
    const nearest = findNearestKmaStation(item.lat, item.lon);
    if (nearest) changeWeatherStation(nearest.id, { keepPlace: true });
    saveWeatherPlace({ label: item.address || item.title, dongCode: item.dongCode, lat: item.lat, lon: item.lon });
    setWeatherPlaceResults(null);
    setWeatherQuery(null);
  };
  // 목록 바깥을 누르면 목록을 닫는다
  useEffect(() => {
    if (weatherPlaceResults === null) return undefined;
    const onDown = (event) => {
      if (weatherPickerRef.current && !weatherPickerRef.current.contains(event.target)) { setWeatherPlaceResults(null); setWeatherQuery(null); }
    };
    document.addEventListener('mousedown', onDown);
    return () => document.removeEventListener('mousedown', onDown);
  }, [weatherPlaceResults]);
  const weatherStationSelect = (
    <span ref={weatherPickerRef} className="weather-picker">
      <input
        ref={weatherSearchInputRef}
        type="text"
        value={weatherQuery === null ? (weatherPlace ? weatherPlace.label : '') : weatherQuery}
        onFocus={(e) => e.target.select()}
        onChange={(e) => setWeatherQuery(e.target.value)}
        onKeyDown={(e) => {
          if (e.key === 'Enter') { e.preventDefault(); searchWeatherPlaces(); }
          if (e.key === 'Escape') { setWeatherPlaceResults(null); setWeatherQuery(null); }
        }}
        placeholder="지역 이름 (예: 안산)"
        title="날씨를 가져올 위치. 지역 이름을 적고 Enter를 누르면 찾습니다"
        aria-label="날씨 위치"
        className="weather-picker-input"
      />
      <button type="button" onClick={searchWeatherPlaces} disabled={weatherBusy} className="px-2 py-1.5 rounded border border-slate-300 bg-white text-slate-700 hover:bg-slate-100 font-bold text-xs disabled:opacity-50">{weatherBusy ? '찾는 중…' : '검색'}</button>
      {!weatherPlace && weatherPlaceResults === null && <span className="text-slate-500">위치를 정하지 않으면 서울 값을 가져옵니다</span>}
      {weatherPlaceResults !== null && (
        <div className="weather-picker-pop" role="listbox" aria-label="장소 검색 결과">
          <div className="weather-picker-head">
            <span>{weatherPlaceResults.length > 0 ? `검색 결과 ${weatherPlaceResults.length}곳` : '검색 결과 없음'}</span>
            <span role="button" tabIndex={0} className="weather-picker-close" onClick={() => { setWeatherPlaceResults(null); setWeatherQuery(null); }} onKeyDown={(e) => { if (e.key === 'Enter') { setWeatherPlaceResults(null); setWeatherQuery(null); } }}>닫기</span>
          </div>
          {weatherPlaceResults.length === 0 && <div className="weather-picker-empty">동 이름이나 시·군 이름으로 다시 찾아 보세요.</div>}
          {weatherPlaceResults.map((item) => {
            const title = String(item.title || '').trim();
            const address = String(item.address || '').trim();
            return (
              <div key={`${title}-${address}-${item.dongCode}`} role="option" aria-selected="false" tabIndex={0} className="weather-picker-item" onClick={() => chooseWeatherPlace(item)} onKeyDown={(e) => { if (e.key === 'Enter') chooseWeatherPlace(item); }}>
                <span className="weather-picker-title">{title || address}</span>
                {address && address !== title && <span className="weather-picker-address">{address}</span>}
              </div>
            );
          })}
        </div>
      )}
    </span>
  );

  const changePalette = (next) => {
    const value = next === 'earth' ? 'earth' : 'default';
    setPalette(value);
    try { localStorage.setItem(PALETTE_STORAGE_KEY, value); } catch (error) { console.error(error); }
  };
  const logoCandidates = [
    palette === 'earth' ? mainLogoClayImage : mainLogoImage,
    `${import.meta.env.BASE_URL}mainlogo.png`,
    `${import.meta.env.BASE_URL}app-logo.png`,
    `${import.meta.env.BASE_URL}mainlogo.jpg`,
    `${import.meta.env.BASE_URL}app-logo.jpg`,
  ];
  const [logoIndex, setLogoIndex] = useState(0);
  const mainLogoSrc = logoCandidates[Math.min(logoIndex, logoCandidates.length - 1)];
  // 기록부 종류별 엑셀 양식 (먼지와 금속은 같은 양식 사용)
  const getTemplateUrlForSheet = (sheetId) => {
    const layout = EXPORT_TEMPLATE_LAYOUTS[sheetId] || EXPORT_TEMPLATE_LAYOUTS.dust;
    return `${import.meta.env.BASE_URL}templates/${layout.file}`;
  };
  const [selectedSheet, setSelectedSheet] = useState('');
  const [nozzleSet, setNozzleSet] = useState(cloneDefaultNozzles);
  const [groupNozzleSets, setGroupNozzleSets] = useState(() => getStoredGroupNozzlesForUser(''));
  // 지금 열린 기록부가 쓰는 노즐 세트 (편집·추천 모두 이 세트 기준)
  const activeNozzleGroup = getNozzleGroupForSheet(selectedSheet);
  const activeNozzleSet = activeNozzleGroup === 'dust' ? nozzleSet : (groupNozzleSets[activeNozzleGroup] || []);
  const setActiveNozzleSet = (updater) => {
    if (activeNozzleGroup === 'dust') { setNozzleSet(updater); return; }
    setGroupNozzleSets((prev) => ({
      ...prev,
      [activeNozzleGroup]: typeof updater === 'function' ? updater(prev[activeNozzleGroup] || []) : updater,
    }));
  };
  const getNozzleSetForSheet = (sheetId) => {
    const group = getNozzleGroupForSheet(sheetId);
    return group === 'dust' ? nozzleSet : (groupNozzleSets[group] || []);
  };
  const [samplers, setSamplers] = useState(cloneDefaultSamplers);
  const [gasSamplers, setGasSamplers] = useState(cloneDefaultGasSamplers);
  const [pitotPresets, setPitotPresets] = useState(cloneDefaultPitotPresets);

  // 빈 기록지 (처음 열 때·로그아웃·불러오기 전 초기화에 공통 사용)
  const createInitialFormData = () => ({
    date: '', company: '', location: '', sampler: '',
    weather: '', atmTemp: '', atmPressure: '',
    totalStackDepth: '', flangeLength: '', stackDiameter: '', pitotInfo: '', pitotFactor: '0.84',
    pitotPresetId: '',
    standardO2: '',
    gasAnalyzer: [
      { time: '', o2: '', co2: '', co: '', nox: '', sox: '' },
      { time: '', o2: '', co2: '', co: '', nox: '', sox: '' },
      { time: '', o2: '', co2: '', co: '', nox: '', sox: '' }
    ],
    moistureValues: ['', '', '', '', ''],
    moistureMethod: 'auto',
    preImpinger: createEmptyPreImpinger(),
    impingers: normalizeImpingers([], getImpingerStageCount('dust')),
    // 동정압 측정 시 가스미터 온도 (예비조사 Tm)
    traverseTmIn: '', traverseTmOut: '',
    points: [
      { id: 1, tp: '', sp: '', dp: '', ts: '' }, { id: 2, tp: '', sp: '', dp: '', ts: '' },
      { id: 3, tp: '', sp: '', dp: '', ts: '' }, { id: 4, tp: '', sp: '', dp: '', ts: '' },
      { id: 5, tp: '', sp: '', dp: '', ts: '' },
    ],
    samplerId: '', targetVolume: '1000', planSamplingTime: '65', samplingStartTime: '',
    leakTest: createEmptyLeakTest(),
    gasMeterFactor: IS_DESKTOP_BUILD ? '' : '0.991', deltaHAt: IS_DESKTOP_BUILD ? '' : '47.6079',
    recommendedNozzleNum: '', recommendedNozzleDia: '', usedNozzleNum: '',
    nozzleDiameter: '', kFactor: '',
    gasMeters: createGasMeterRowsByPoints(5),
    filterId: '', filterInitial: '', filterFinal: '', remarks: ''
  });
  const [formData, setFormData] = useState(createInitialFormData);

  const [secureUsers, setSecureUsers] = useState({});
  const [samplingLogData, setSamplingLogData] = useState(createEmptySamplingLog);
  // 기록부별 별도 창 핸들 (이미 열려 있으면 다시 띄우지 않고 앞으로 가져옴)
  const sheetWindowsRef = useRef({});
  // 통합 기록부 임시 저장본이 어느 사용자 키로 로드됐는지 (로드 전에 빈 내용으로 덮어쓰지 않도록)
  const [samplingLogDraftReadyKey, setSamplingLogDraftReadyKey] = useState('');
  const [samplingLogRequestedOnly, setSamplingLogRequestedOnly] = useState(false);
  // 측정점 표 자동 생성 옵션: 총 채취시간(분, 비우면 1단계 예상 채취시간 또는 60분), 행 간격(분)
  const [tableTotalMinutes, setTableTotalMinutes] = useState('');
  const [activeUser, setActiveUser] = useState('');
  const [activeUserReports, setActiveUserReports] = useState([]);
  const [isUserUnlocked, setIsUserUnlocked] = useState(false);
  const [loginUserId, setLoginUserId] = useState(BETA_USER_ID);
  const [loginPassword, setLoginPassword] = useState('');
  const [newUserName, setNewUserName] = useState('');
  const [newUserPassword, setNewUserPassword] = useState('');
  const [newUserPasswordConfirm, setNewUserPasswordConfirm] = useState('');
  const [newUserHintQuestion, setNewUserHintQuestion] = useState('');
  const [newUserHintAnswer, setNewUserHintAnswer] = useState('');
  const [recoverHintAnswer, setRecoverHintAnswer] = useState('');
  const [recoveredUserIds, setRecoveredUserIds] = useState([]);
  const [recoverHintUserId, setRecoverHintUserId] = useState('');
  const [recoveredHintQuestion, setRecoveredHintQuestion] = useState('');
  const [profileAvatarUrl, setProfileAvatarUrl] = useState('');
  const [profileNickname, setProfileNickname] = useState('');
  const [authModal, setAuthModal] = useState('');
  // ── 엑셀 저장 위치 (설치형 프로그램에서만): 이 폴더 아래에 "날짜_사업장명_#배출구" 폴더를 만들어 저장한다 ──
  const [exportBase, setExportBase] = useState('');
  const canSetExportBase = typeof window !== 'undefined' && Boolean(window.stackpilotDesktop?.getExportBase);
  useEffect(() => {
    if (authModal !== 'settings' || !canSetExportBase) return;
    window.stackpilotDesktop.getExportBase().then((dir) => setExportBase(String(dir || ''))).catch((error) => console.error(error));
  }, [authModal, canSetExportBase]);
  const chooseExportBase = async () => {
    try { setExportBase(String(await window.stackpilotDesktop.chooseExportBase() || '')); } catch (error) { console.error(error); alert('저장 위치를 바꾸지 못했습니다.'); }
  };
  const resetExportBase = async () => {
    try { setExportBase(String(await window.stackpilotDesktop.resetExportBase() || '')); } catch (error) { console.error(error); }
  };

  const [cloudSyncConfig, setCloudSyncConfig] = useState({ token: '', gistId: '' });
  const [cloudSyncBusy, setCloudSyncBusy] = useState(false);
  const [isStorageHydrated, setIsStorageHydrated] = useState(false);
  const isPresetHydratingRef = useRef(false);
  const vaultKeyRef = useRef({});
  const cloudSyncBusyRef = useRef(false);
  const autoCloudRestoreDoneRef = useRef({});
  const autoCloudRestoreInFlightRef = useRef(false);
  const autoCloudUploadTimerRef = useRef(null);
  const autoCloudUploadInFlightRef = useRef(false);
  const autoCloudUploadFailRef = useRef(null);
  const unsavedRef = useRef(false);

  // ── 표 채우기(필다운): 표 안의 칸에서 Ctrl+D = 바로 윗칸 값 복사, Ctrl+Shift+D = 지금 값을 아래 끝까지 채우기 ──
  // 화면의 모든 입력 표(측정 기록, 측정점, 가스 분석 등)에 공통으로 적용
  useEffect(() => {
    const onKeyDown = async (event) => {
      if (!(event.ctrlKey || event.metaKey) || event.altKey || String(event.key).toLowerCase() !== 'd') return;
      const field = event.target;
      if (!(field instanceof HTMLInputElement || field instanceof HTMLSelectElement)) return;
      const found = fillPeerFinder(field);
      if (!found) return;
      event.preventDefault();
      const { rows, rowIdx, peerOf } = found;
      if (event.shiftKey) {
        await fillValueInto(rows.slice(rowIdx + 1).map(peerOf).filter(Boolean), field.value);
      } else {
        for (let i = rowIdx - 1; i >= 0; i--) {
          const above = peerOf(rows[i]);
          if (!above) continue;
          if (above.value !== field.value) fillSetFieldValue(field, above.value);
          break;
        }
      }
    };
    document.addEventListener('keydown', onKeyDown);
    return () => document.removeEventListener('keydown', onKeyDown);
  }, []);

  // ── 끌어서 채우기: 측정점 표·측정 기록표(.fill-drag-table)에서 칸을 고르면 오른쪽 아래에 작은 손잡이가 생기고,
  //    그 손잡이를 아래(또는 위)로 끌면 지나간 칸에 같은 값이 들어간다 (엑셀의 채우기 핸들) ──
  useEffect(() => {
    const handle = document.createElement('div');
    handle.className = 'fill-handle';
    handle.title = '끌어서 같은 값 채우기';
    handle.style.display = 'none';
    document.body.appendChild(handle);
    let source = null;
    let drag = null;
    const place = () => {
      if (!source || !source.isConnected) { handle.style.display = 'none'; return; }
      const rect = source.getBoundingClientRect();
      if (rect.width === 0 || rect.bottom < 70 || rect.bottom > window.innerHeight) { handle.style.display = 'none'; return; }
      handle.style.display = 'block';
      handle.style.left = `${rect.right - 5}px`;
      handle.style.top = `${rect.bottom - 5}px`;
    };
    const clearMarks = () => document.querySelectorAll('.fill-target').forEach((el) => el.classList.remove('fill-target'));
    const onFocusIn = (event) => {
      if (drag) return;
      const field = event.target;
      const ok = (field instanceof HTMLInputElement || field instanceof HTMLSelectElement) && field.closest('.fill-drag-table tbody td') && fillCellFields(field.closest('td')).includes(field);
      source = ok ? field : null;
      place();
    };
    const onFocusOut = () => { if (!drag) { source = null; setTimeout(() => { if (!source && !drag) handle.style.display = 'none'; }, 0); } };
    const targetsAt = (clientY) => {
      const { rows, rowIdx, peerOf } = drag.found;
      // 포인터가 놓인 행 찾기 (표 밖으로 나가면 가장 가까운 끝 행)
      let hit = rowIdx;
      rows.forEach((r, i) => {
        const rect = r.getBoundingClientRect();
        if (clientY >= rect.top && clientY < rect.bottom) hit = i;
      });
      const first = rows[0].getBoundingClientRect();
      const last = rows[rows.length - 1].getBoundingClientRect();
      if (clientY < first.top) hit = 0;
      if (clientY >= last.bottom) hit = rows.length - 1;
      const [from, to] = hit >= rowIdx ? [rowIdx + 1, hit] : [hit, rowIdx - 1];
      const list = [];
      for (let i = from; i <= to; i++) { const peer = peerOf(rows[i]); if (peer) list.push(peer); }
      return list;
    };
    const onPointerDown = (event) => {
      if (!source) return;
      const found = fillPeerFinder(source);
      if (!found) return;
      event.preventDefault();
      drag = { found, value: source.value, targets: [] };
      handle.setPointerCapture(event.pointerId);
      document.body.classList.add('fill-dragging');
    };
    const onPointerMove = (event) => {
      if (!drag) return;
      clearMarks();
      drag.targets = targetsAt(event.clientY);
      drag.targets.forEach((el) => el.classList.add('fill-target'));
    };
    const onPointerUp = async (event) => {
      if (!drag) return;
      const { targets, value } = drag;
      try { handle.releasePointerCapture(event.pointerId); } catch { /* 이미 풀림 */ }
      drag = null;
      document.body.classList.remove('fill-dragging');
      clearMarks();
      await fillValueInto(targets, value);
      place();
    };
    document.addEventListener('focusin', onFocusIn);
    document.addEventListener('focusout', onFocusOut);
    handle.addEventListener('pointerdown', onPointerDown);
    handle.addEventListener('pointermove', onPointerMove);
    handle.addEventListener('pointerup', onPointerUp);
    handle.addEventListener('pointercancel', onPointerUp);
    window.addEventListener('scroll', place, true);
    window.addEventListener('resize', place);
    return () => {
      document.removeEventListener('focusin', onFocusIn);
      document.removeEventListener('focusout', onFocusOut);
      window.removeEventListener('scroll', place, true);
      window.removeEventListener('resize', place);
      handle.remove();
    };
  }, []);
  const [cloudRetryTick, setCloudRetryTick] = useState(0);
  const lastAutoCloudUploadSignatureRef = useRef('');
  const csvImportInputRef = useRef(null);
  const [menuCheckedReportKeys, setMenuCheckedReportKeys] = useState([]);
  const [expandedReportGroups, setExpandedReportGroups] = useState([]);
  const [sheetCheckedReportKeys, setSheetCheckedReportKeys] = useState([]);
  const [recommendations, setRecommendations] = useState(null);
  const [presetHistory, setPresetHistory] = useState(readPresetHistory);
  // 노즐 추천에 쓰는 오리피스압(ΔH) 기준: 빠른 채취의 상한, 표준·안정 채취의 목표값. 이 기기에 기억한다
  const [dhTargets, setDhTargets] = useState(() => {
    try { return { ...DEFAULT_DH_TARGETS, ...(JSON.parse(localStorage.getItem(DH_TARGETS_STORAGE_KEY)) || {}) }; } catch { return { ...DEFAULT_DH_TARGETS }; }
  });
  const getDhTargets = () => {
    const pick = (key) => { const v = parseFloat(dhTargets[key]); return Number.isFinite(v) && v > 0 ? v : DEFAULT_DH_TARGETS[key]; };
    return { max: pick('max'), standard: pick('standard'), stable: pick('stable') };
  };
  const changeDhTarget = (key, value) => {
    setDhTargets((prev) => {
      const next = { ...prev, [key]: value };
      try { localStorage.setItem(DH_TARGETS_STORAGE_KEY, JSON.stringify(next)); } catch { /* 저장 못 해도 이번 계산엔 쓴다 */ }
      return next;
    });
    setRecommendations(null);
  };
  const [skyPhase, setSkyPhase] = useState(() => getSkyPhaseByHour());
  const [skyPreviewMode, setSkyPreviewMode] = useState(() => {
    try {
      return sanitizeSkyPreviewMode(localStorage.getItem(STORAGE_KEYS.skyPreviewMode));
    } catch (error) {
      console.error('하늘 시간대 설정 로딩 실패:', error);
      return 'auto';
    }
  });
  const [isMobileViewport, setIsMobileViewport] = useState(() => isMobileViewportWidth());
  const [isKFactorEditing, setIsKFactorEditing] = useState(false);
  const displayUserName = String(profileNickname || '').trim() || activeUser || '';
  const displayUserNameWithFallback = displayUserName || '미선택';

  useEffect(() => {
    cloudSyncBusyRef.current = cloudSyncBusy;
  }, [cloudSyncBusy]);

  useEffect(() => {
    try {
      localStorage.setItem(STORAGE_KEYS.skyPreviewMode, sanitizeSkyPreviewMode(skyPreviewMode));
    } catch (error) {
      console.error('하늘 시간대 설정 저장 실패:', error);
    }
  }, [skyPreviewMode]);

  // 로그인하면 지금 저장돼 있는 교정값을 이력의 첫 줄로 남긴다 (이미 이력이 있는 장비는 건너뛴다)
  useEffect(() => {
    if (!isUserUnlocked || !activeUser) return undefined;
    const timer = setTimeout(() => {
      const pick = (userKey, globalKey, fallback, sanitize) => {
        const own = readStoredList(userKey);
        const list = own.length > 0 ? own : readStoredList(globalKey);
        return sanitize(list.length > 0 ? list : fallback) || [];
      };
      appendPresetHistory('sampler', null, pick(getUserSamplersStorageKey(activeUser), STORAGE_KEYS.samplers, cloneDefaultSamplers(), sanitizeSamplers), activeUser, '처음 등록값');
      appendPresetHistory('gas', null, pick(getUserGasSamplersStorageKey(activeUser), STORAGE_KEYS.gasSamplers, cloneDefaultGasSamplers(), sanitizeGasSamplers), activeUser, '처음 등록값');
      setPresetHistory(appendPresetHistory('pitot', null, pick(getUserPitotStorageKey(activeUser), STORAGE_KEYS.pitotPresets, cloneDefaultPitotPresets(), sanitizePitotPresets), activeUser, '처음 등록값'));
    }, 1200);
    return () => clearTimeout(timer);
  }, [isUserUnlocked, activeUser]);

  const canUseUserScopedPresets = (userId) => {
    const normalizedId = String(userId || '').trim().toLowerCase();
    return Boolean(normalizedId) && normalizedId !== BETA_USER_ID && normalizedId !== LEGACY_BETA_USER_ID;
  };

  useEffect(() => {
    if (!activeUser) {
      setCloudSyncConfig({ token: '', gistId: '' });
      return;
    }
    try {
      const raw = localStorage.getItem(getUserCloudSyncStorageKey(activeUser));
      if (!raw) {
        setCloudSyncConfig({ token: '', gistId: '' });
        return;
      }
      const parsed = sanitizeCloudSyncConfig(JSON.parse(raw));
      setCloudSyncConfig(parsed || { token: '', gistId: '' });
    } catch (error) {
      console.error('클라우드 동기화 설정 로딩 실패:', error);
      setCloudSyncConfig({ token: '', gistId: '' });
    }
  }, [activeUser]);

  useEffect(() => {
    let cancelled = false;

    const bootstrap = async () => {
      try {
        const rawNozzles = localStorage.getItem(STORAGE_KEYS.nozzles);
        const rawSamplers = localStorage.getItem(STORAGE_KEYS.samplers);
        const rawPitotPresets = localStorage.getItem(STORAGE_KEYS.pitotPresets);
        const rawSecureUsers = localStorage.getItem(STORAGE_KEYS.secureUsers);
        const rawLastLoginId = localStorage.getItem(STORAGE_KEYS.lastLoginId);

        if (rawNozzles) {
          const parsedNozzles = sanitizeNozzleSet(JSON.parse(rawNozzles));
          if (!cancelled && parsedNozzles) setNozzleSet(parsedNozzles);
        }

        if (rawSamplers) {
          const parsedSamplers = sanitizeSamplers(JSON.parse(rawSamplers));
          if (!cancelled && parsedSamplers && parsedSamplers.length > 0) setSamplers(parsedSamplers);
        }

        if (rawPitotPresets) {
          const parsedPitotPresets = sanitizePitotPresets(JSON.parse(rawPitotPresets));
          if (!cancelled && parsedPitotPresets && parsedPitotPresets.length > 0) setPitotPresets(parsedPitotPresets);
        }

        const parsedUsers = rawSecureUsers ? sanitizeSecureUsers(JSON.parse(rawSecureUsers)) : null;
        let normalizedUsers = parsedUsers || {};
        if (!cancelled) setSecureUsers(normalizedUsers);

        let restored = false;
        const savedSession = readSessionAuth();
        if (savedSession) {
          const { userId, password } = savedSession;
          const targetUser = normalizedUsers[userId];

          if (targetUser) {
            try {
              const key = await deriveVaultKey(password, targetUser.salt);
              const verifierText = await decryptTextWithKey(key, targetUser.verifierIv, targetUser.verifierCipher);
              if (verifierText === `vault:${userId}:verified`) {
                const decrypted = await decryptJsonWithKey(key, targetUser.dataIv, targetUser.dataCipher);
                const reports = Array.isArray(decrypted) ? decrypted.filter(item => item && typeof item === 'object') : [];
                const fixedAvatarUrl = sanitizeAvatarUrl(targetUser.avatarUrl);

                if (targetUser.avatarUrl !== fixedAvatarUrl) {
                  normalizedUsers = {
                    ...normalizedUsers,
                    [userId]: {
                      ...targetUser,
                      avatarUrl: fixedAvatarUrl,
                      updatedAt: new Date().toISOString(),
                    },
                  };
                  if (!cancelled) setSecureUsers(normalizedUsers);
                }

                if (!cancelled) {
                  openUserSession(userId, key, reports, normalizedUsers[userId]);
                  restored = true;
                }
              }
            } catch (error) {
              console.error('자동 로그인 복원 실패:', error);
            }
          }
        }

        if (!restored && !cancelled) {
          clearSessionAuth();
          setActiveUser('');
          setLoginUserId(BETA_USER_ID);
          setIsUserUnlocked(false);
          setActiveUserReports([]);
          setLoginPassword('');
          setProfileAvatarUrl('');
          setProfileNickname('');
          setAuthModal('');
          setRecoveredUserIds([]);
          setRecoveredHintQuestion('');
        }
      } catch (error) {
        console.error('로컬 저장값 로딩 실패:', error);
      } finally {
        if (!cancelled) setIsStorageHydrated(true);
      }
    };

    bootstrap();
    return () => {
      cancelled = true;
    };
  }, []);

  useEffect(() => {
    if (!isStorageHydrated || isPresetHydratingRef.current) return;
    try {
      const targetKey = canUseUserScopedPresets(activeUser)
        ? getUserNozzlesStorageKey(activeUser)
        : STORAGE_KEYS.nozzles;
      localStorage.setItem(targetKey, JSON.stringify(nozzleSet));
    } catch (error) {
      console.error('노즐 설정 저장 실패:', error);
    }
  }, [nozzleSet, isStorageHydrated, activeUser]);

  useEffect(() => {
    if (!isStorageHydrated || isPresetHydratingRef.current) return;
    try {
      localStorage.setItem(getGroupNozzlesStorageKey(activeUser), JSON.stringify(groupNozzleSets));
    } catch (error) {
      console.error('수은·PAHs 노즐 설정 저장 실패:', error);
    }
  }, [groupNozzleSets, isStorageHydrated, activeUser]);

  useEffect(() => {
    if (!isStorageHydrated || isPresetHydratingRef.current) return;
    try {
      const targetKey = canUseUserScopedPresets(activeUser)
        ? getUserSamplersStorageKey(activeUser)
        : STORAGE_KEYS.samplers;
      localStorage.setItem(targetKey, JSON.stringify(samplers));
    } catch (error) {
      console.error('샘플러 설정 저장 실패:', error);
    }
  }, [samplers, isStorageHydrated, activeUser]);

  useEffect(() => {
    if (!isStorageHydrated || isPresetHydratingRef.current) return;
    try {
      const targetKey = canUseUserScopedPresets(activeUser)
        ? getUserGasSamplersStorageKey(activeUser)
        : STORAGE_KEYS.gasSamplers;
      localStorage.setItem(targetKey, JSON.stringify(gasSamplers));
    } catch (error) {
      console.error('가스 샘플러 설정 저장 실패:', error);
    }
  }, [gasSamplers, isStorageHydrated, activeUser]);

  useEffect(() => {
    if (!isStorageHydrated || isPresetHydratingRef.current) return;
    try {
      const targetKey = canUseUserScopedPresets(activeUser)
        ? getUserPitotStorageKey(activeUser)
        : STORAGE_KEYS.pitotPresets;
      localStorage.setItem(targetKey, JSON.stringify(pitotPresets));
    } catch (error) {
      console.error('피토 프리셋 저장 실패:', error);
    }
  }, [pitotPresets, isStorageHydrated, activeUser]);

  useEffect(() => {
    if (!isStorageHydrated) return;
    try {
      localStorage.setItem(STORAGE_KEYS.secureUsers, JSON.stringify(secureUsers));
    } catch (error) {
      console.error('보안 사용자 저장 실패:', error);
      // 저장 공간이 가득 차면 "저장했습니다"만 보고 재시작 후 사라지는 일이 없도록 알린다
      alert('저장 공간이 가득 차서 방금 저장한 기록이 이 PC에 보관되지 않았습니다.\n오래된 리포트를 엑셀로 추출한 뒤 삭제하고 다시 저장해주세요.');
    }
  }, [secureUsers, isStorageHydrated]);

  useEffect(() => {
    if (!isStorageHydrated) return;
    try {
      localStorage.setItem(STORAGE_KEYS.activeUser, activeUser);
    } catch (error) {
      console.error('활성 사용자 저장 실패:', error);
    }
  }, [activeUser, isStorageHydrated]);

  // ── 통합 시료채취기록부: 창 사이 공유 ──
  // 입력 중인 내용을 사용자별 임시 저장본으로 보관하고, 다른 창(별도 창/메인 창)의 변경을 storage 이벤트로 받아 동기화
  const samplingLogDraftKey = activeUser && isUserUnlocked ? getSamplingLogDraftStorageKey(activeUser) : '';
  // 다른 창에서 받아 온 임시본(문자열). 받은 그대로라면 다시 써 보내지 않는다 —
  // 늦게 도착한 옛 이벤트를 받아 놓고 되쓰면 더 새 값을 덮어쓰기 때문
  const samplingLogReceivedRef = useRef('');
  const samplingLogFreshRef = useRef(OPENED_SAMPLING_LOG_FRESH);
  useEffect(() => {
    if (!samplingLogDraftKey) {
      setSamplingLogDraftReadyKey('');
      return;
    }
    // 새로 연 창은 지난 임시본을 불러오지 않고 빈 기록지로 시작한다 (빈 임시본이 다른 창에도 전해진다)
    if (samplingLogFreshRef.current) {
      samplingLogFreshRef.current = false;
      setSamplingLogData(createEmptySamplingLog());
      setSamplingLogDraftReadyKey(samplingLogDraftKey);
      return;
    }
    try {
      const raw = localStorage.getItem(samplingLogDraftKey);
      if (raw) setSamplingLogData(sanitizeSamplingLog(JSON.parse(raw)));
    } catch (error) {
      console.error('통합 기록부 임시 저장본 로딩 실패:', error);
    }
    setSamplingLogDraftReadyKey(samplingLogDraftKey);
  }, [samplingLogDraftKey]);

  useEffect(() => {
    if (!samplingLogDraftKey || samplingLogDraftReadyKey !== samplingLogDraftKey) return;
    try {
      // 정리된 형태로 저장해야 다른 창이 받은 값을 다시 써 보내는(옛 값으로 되돌리는) 일이 없다
      const next = JSON.stringify(sanitizeSamplingLog(samplingLogData));
      if (next === samplingLogReceivedRef.current) return;
      if (localStorage.getItem(samplingLogDraftKey) !== next) localStorage.setItem(samplingLogDraftKey, next);
    } catch (error) {
      console.error('통합 기록부 임시 저장 실패:', error);
    }
  }, [samplingLogData, samplingLogDraftKey, samplingLogDraftReadyKey]);

  useEffect(() => {
    const onStorage = async (event) => {
      if (!event || !event.key) return;
      if (samplingLogDraftKey && event.key === samplingLogDraftKey && event.newValue) {
        try {
          const received = sanitizeSamplingLog(JSON.parse(event.newValue));
          samplingLogReceivedRef.current = JSON.stringify(received);
          setSamplingLogData(received);
        } catch (error) {
          console.error('다른 창의 통합 기록부 동기화 실패:', error);
        }
        return;
      }
      if (event.newValue && event.key === getGroupNozzlesStorageKey(activeUser)) {
        try { setGroupNozzleSets(sanitizeGroupNozzleSets(JSON.parse(event.newValue))); } catch (error) { console.error(error); }
        return;
      }
      if (activeUser && event.newValue && (event.key === getUserNozzlesStorageKey(activeUser) || event.key === STORAGE_KEYS.nozzles)) {
        try { const parsed = sanitizeNozzleSet(JSON.parse(event.newValue)); if (parsed) setNozzleSet(parsed); } catch (error) { console.error(error); }
        return;
      }
      if (activeUser && event.newValue && (event.key === getUserSamplersStorageKey(activeUser) || event.key === STORAGE_KEYS.samplers)) {
        try { const parsed = sanitizeSamplers(JSON.parse(event.newValue)); if (parsed) setSamplers(parsed); } catch (error) { console.error(error); }
        return;
      }
      if (activeUser && event.newValue && (event.key === getUserGasSamplersStorageKey(activeUser) || event.key === STORAGE_KEYS.gasSamplers)) {
        try { const parsed = sanitizeGasSamplers(JSON.parse(event.newValue)); if (parsed) setGasSamplers(parsed); } catch (error) { console.error(error); }
        return;
      }
      if (activeUser && event.newValue && (event.key === getUserPitotStorageKey(activeUser) || event.key === STORAGE_KEYS.pitotPresets)) {
        try { const parsed = sanitizePitotPresets(JSON.parse(event.newValue)); if (parsed) setPitotPresets(parsed); } catch (error) { console.error(error); }
        return;
      }
      if (event.key === STORAGE_KEYS.sessionAuth && !event.newValue && activeUser) {
        // 다른 창에서 로그아웃하면 이 창도 잠근다 (기록부 창은 닫음)
        handleLockActiveUser();
        if (IS_POPUP_WINDOW) window.close();
        return;
      }
      if (event.key === STORAGE_KEYS.secureUsers && event.newValue) {
        // 다른 창에서 리포트를 저장하면 이 창의 저장 목록도 갱신 (덮어쓰기 방지)
        try {
          const parsed = sanitizeSecureUsers(JSON.parse(event.newValue));
          if (!parsed) return;
          setSecureUsers(parsed);
          const record = activeUser ? parsed[activeUser] : null;
          const key = activeUser ? vaultKeyRef.current[activeUser] : null;
          if (record && key) {
            const decrypted = await decryptJsonWithKey(key, record.dataIv, record.dataCipher);
            setActiveUserReports(Array.isArray(decrypted) ? decrypted.filter((item) => item && typeof item === 'object') : []);
          }
        } catch (error) {
          console.error('다른 창의 저장 목록 동기화 실패:', error);
        }
      }
    };
    window.addEventListener('storage', onStorage);
    return () => window.removeEventListener('storage', onStorage);
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [samplingLogDraftKey, activeUser]);

  // 메인 창의 현재 기록부 정보(사업장·배출구·측정일)를 별도 창에서 가져갈 수 있도록 기록
  useEffect(() => {
    if (!selectedSheet || selectedSheet === 'samplinglog') return;
    try {
      localStorage.setItem(STORAGE_KEYS.sheetContext, JSON.stringify({
        company: formData.company || '',
        location: formData.location || '',
        date: formData.date || '',
        sheetTitle: SHEET_MENU.find((item) => item.id === selectedSheet)?.title || '',
      }));
    } catch (error) {
      console.error('기록부 정보 공유 실패:', error);
    }
  }, [formData.company, formData.location, formData.date, selectedSheet]);

  useEffect(() => {
    if (!IS_POPUP_WINDOW) return;
    const title = SHEET_MENU.find((item) => item.id === selectedSheet)?.title;
    document.title = title ? `StackPilot - ${title}` : 'StackPilot';
  }, [selectedSheet]);

  // 기록부를 별도 창으로 연다. report가 있으면 그 창이 열리자마자 불러오도록 임시 저장소로 넘긴다.
  const openSheetWindow = (sheetId, report = null) => {
    const existing = sheetWindowsRef.current[sheetId];
    let loadToken = '';
    if (report) {
      loadToken = `${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
      try {
        localStorage.setItem(STORAGE_KEYS.pendingReport, JSON.stringify({ token: loadToken, sheetId, report }));
      } catch (error) {
        console.error('리포트 전달 실패:', error);
      }
    }
    if (existing && !existing.closed) {
      try {
        existing.location.hash = `sheet=${sheetId}&popup=1${loadToken ? `&load=${loadToken}` : ''}`;
        existing.focus();
      } catch (error) {
        console.error(error);
      }
      return;
    }
    const url = `${window.location.pathname}${window.location.search}#sheet=${sheetId}&popup=1${loadToken ? `&load=${loadToken}` : ''}`;
    const popup = window.open(url, `stackpilot-${sheetId}`, 'popup=yes,width=1280,height=900');
    if (!popup) {
      alert('새 창을 열 수 없습니다. 브라우저의 팝업 차단을 해제한 뒤 다시 시도해주세요.');
      return;
    }
    sheetWindowsRef.current[sheetId] = popup;
  };
  // 이미 열린 기록부 창에 불러오기를 보내면 주소(#...&load=)만 바뀌므로, 주소 변경도 감지한다
  const [hashLoadTick, setHashLoadTick] = useState(0);
  useEffect(() => {
    if (!IS_POPUP_WINDOW) return undefined;
    const onHash = () => { if (window.location.hash.includes('load=')) setHashLoadTick((t) => t + 1); };
    window.addEventListener('hashchange', onHash);
    return () => window.removeEventListener('hashchange', onHash);
  }, []);
  // 별도 창: 주소의 load 토큰과 맞는 리포트가 전달돼 있으면 열리자마자 불러오기
  useEffect(() => {
    if (!IS_POPUP_WINDOW || !activeUser || !isUserUnlocked) return;
    const params = new URLSearchParams(window.location.hash.replace(/^#/, ''));
    const token = params.get('load');
    if (!token) return;
    try {
      const raw = localStorage.getItem(STORAGE_KEYS.pendingReport);
      const parsed = raw ? JSON.parse(raw) : null;
      if (parsed && parsed.token === token && parsed.report) {
        localStorage.removeItem(STORAGE_KEYS.pendingReport);
        handleLoadSavedReport(parsed.report, true);
      }
    } catch (error) {
      console.error('전달된 리포트 불러오기 실패:', error);
    }
    params.delete('load');
    window.history.replaceState(null, '', `${window.location.pathname}${window.location.search}#${params.toString()}`);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [activeUser, isUserUnlocked, selectedSheet, hashLoadTick]);

  const buildReportKey = (id, savedAt) => String(id || savedAt || '');
  const getReportKey = (report) => buildReportKey(report?.id, report?.savedAt);

  const savedData = [...activeUserReports].sort((a, b) => {
    // 가장 최근 저장 파일이 항상 맨 위로 오도록 저장시각 기준 내림차순 정렬
    const aTime = new Date(a.savedAt || 0).getTime();
    const bTime = new Date(b.savedAt || 0).getTime();
    if (Number.isFinite(aTime) && Number.isFinite(bTime)) return bTime - aTime;
    if (Number.isFinite(bTime)) return 1;
    if (Number.isFinite(aTime)) return -1;
    return 0;
  });
  const resolveReportSheetId = (report) => {
    const byId = String(report?.sheetId || '').trim();
    if (byId && SHEET_MENU.some(item => item.id === byId)) return byId;

    const byTitle = String(report?.sheetTitle || '').trim();
    const matched = SHEET_MENU.find(item => item.title === byTitle);
    if (matched) return matched.id;
    if (byTitle === '시료채취기록부(종합)') return 'samplinglog'; // 예전 이름으로 저장된 리포트

    return 'dust';
  };
  const profileCombinedReports = savedData.map((report) => {
    const sheetId = resolveReportSheetId(report);
    const sheetMeta = SHEET_MENU.find(item => item.id === sheetId);
    return {
      ...report,
      _sheetId: sheetId,
      _sheetTitle: sheetMeta ? sheetMeta.title : (report.sheetTitle || '먼지시료채취기록부'),
    };
  });
  // 메뉴 목록: 사업장·스택명이 같은 리포트를 한 묶음으로 (날짜는 기간으로 표시)
  const profileReportGroups = groupSheetReportsForLog(profileCombinedReports, resolveReportSheetId, null, { includeBlank: true });
  const toggleReportGroup = (key) => setExpandedReportGroups((prev) => (prev.includes(key) ? prev.filter((k) => k !== key) : [...prev, key]));

  // 저장 리포트 묶음 표 (메뉴와 기록부 창이 같은 양식 사용)
  // 펼치면 기록부 종류별로 최신 저장본 한 줄씩, 그 아래에 이전 저장본은 회색 줄로
  const renderReportGroupTable = ({ groups, checkedSet, toggleCheck, allChecked, toggleAll, allowCopy = false, emptyText = '저장된 리포트가 없습니다.' }) => {
    if (!groups || groups.length === 0) {
      return (
        <div className="p-5 text-center text-xs text-slate-500 border border-dashed border-slate-300 rounded-lg">{emptyText}</div>
      );
    }
    const summaryOf = (data) => {
      const sheetId = resolveReportSheetId(data);
      if (sheetId === 'samplinglog') return '통합시료채취기록부';
      if (sheetId === 'flow') return `유속 ${data.avgVelocity || '-'} m/s · 건조유량 ${data.dryFlowRate || '-'} Sm³/hr`;
      return `등속 ${data.isokineticRate || '-'}% · 실측 ${data.actualConcentration || '-'} · 보정 ${data.correctedConcentration || '-'}`;
    };
    const exportOne = async (data) => {
      if (resolveReportSheetId(data) === 'samplinglog') await exportSamplingLogExcel(data);
      else await exportToTemplateExcel(data);
    };
    const actionButtons = (data) => (
      <span className="inline-flex items-center gap-1">
        <button type="button" onClick={(e) => { e.stopPropagation(); handleLoadSavedReport(data); }} className="px-2 py-1 rounded text-[11px] border border-slate-300 bg-white text-slate-800 hover:bg-slate-50 font-bold whitespace-nowrap">불러오기</button>
        <button type="button" onClick={(e) => { e.stopPropagation(); exportOne(data); }} className="px-2 py-1 rounded text-[11px] border border-emerald-300 bg-white text-emerald-700 hover:bg-emerald-50 font-bold whitespace-nowrap">엑셀</button>
        {allowCopy && (
          <button type="button" onClick={(e) => { e.stopPropagation(); handleCopySavedReportToCurrentSheet(data); }} className={`px-2 py-1 rounded text-[11px] text-white font-bold whitespace-nowrap ${activeTheme.primaryButton}`}>이 기록부로 복사</button>
        )}
        <button type="button" onClick={(e) => { e.stopPropagation(); removeSavedReport(data.id, data.savedAt); }} className="px-2 py-1 rounded text-[11px] border border-red-200 bg-white text-red-600 hover:bg-red-50 font-bold whitespace-nowrap">삭제</button>
      </span>
    );
    return (
      <div className="overflow-x-auto rounded-lg border border-slate-200">
        <table className="w-full text-xs min-w-[1000px]">
          <thead className="bg-slate-100 text-slate-700">
            <tr>
              <th className="p-2 font-bold text-center w-8">
                <input type="checkbox" checked={allChecked} onChange={(e) => toggleAll(e.target.checked)} title="전체 선택" />
              </th>
              <th className="p-2 font-bold text-left">사업장</th>
              <th className="p-2 font-bold text-left">배출구(스택)</th>
              <th className="p-2 font-bold text-left">측정일자</th>
              <th className="p-2 font-bold text-left">저장된 기록부</th>
              <th className="p-2 font-bold text-center">건수</th>
              <th className="p-2 font-bold text-center w-24">상세</th>
            </tr>
          </thead>
          <tbody className="divide-y divide-slate-100">
            {groups.map((group) => {
              const expanded = expandedReportGroups.includes(group.key);
              const groupKeys = group.reports.map(getReportKey).filter(Boolean);
              const groupAllChecked = groupKeys.length > 0 && groupKeys.every((k) => checkedSet.has(k));
              // 종류별 최신 저장본 + 이전 저장본
              const bySheet = SHEET_MENU
                .map((sheet) => ({ sheet, list: group.reports.filter((r) => resolveReportSheetId(r) === sheet.id) }))
                .filter(({ list }) => list.length > 0);
              return (
                <Fragment key={group.key}>
                  <tr className="hover:bg-slate-50 cursor-pointer" onClick={() => toggleReportGroup(group.key)}>
                    <td className="p-2 text-center">
                      <input type="checkbox" checked={groupAllChecked} onClick={(e) => e.stopPropagation()} onChange={(e) => groupKeys.forEach((k) => toggleCheck(k, e.target.checked))} title="이 묶음 전체 선택" />
                    </td>
                    <td className="p-2 font-bold text-slate-900 whitespace-nowrap">{group.company || (group.location || group.dates.length ? '-' : '(정보 없음)')}</td>
                    <td className="p-2 whitespace-nowrap">{group.location || '-'}</td>
                    <td className="p-2 whitespace-nowrap">{group.dateLabel}</td>
                    <td className="p-2">
                      <div className="flex flex-wrap gap-1">
                        {bySheet.map(({ sheet, list }) => (
                          <button key={sheet.id} type="button" onClick={(e) => { e.stopPropagation(); handleLoadSavedReport(list[0]); }} title={`${sheet.title} 최신 저장본 열기`} className="hover:opacity-80">
                            <SheetBadge sheetId={sheet.id} title={sheet.title} />
                          </button>
                        ))}
                      </div>
                    </td>
                    <td className="p-2 text-center font-bold">{group.reports.length}</td>
                    <td className="p-2 text-center"><span className="text-[11px] font-bold text-slate-600">{expanded ? '접기 ▴' : '펼치기 ▾'}</span></td>
                  </tr>
                  {expanded && bySheet.map(({ sheet, list }) => list.map((data, idx) => {
                    const reportKey = getReportKey(data);
                    const isLatest = idx === 0;
                    return (
                      <tr key={reportKey || data.savedAt} className={`${isLatest ? 'bg-slate-50/60' : 'bg-slate-50/30 text-slate-500'} hover:bg-slate-100 cursor-pointer`} onClick={() => handleLoadSavedReport(data)} title="클릭하면 불러옵니다">
                        <td className="p-2 text-center">
                          <input type="checkbox" checked={checkedSet.has(reportKey)} onClick={(e) => e.stopPropagation()} onChange={(e) => toggleCheck(reportKey, e.target.checked)} />
                        </td>
                        <td className="p-2 pl-6" colSpan="2">
                          <span className="inline-flex items-center gap-2">
                            {isLatest ? <SheetBadge sheetId={sheet.id} title={sheet.title} /> : <span className="inline-block min-w-[2.6rem] text-center text-[11px] text-slate-400">└ 이전</span>}
                            <span className={isLatest ? 'text-slate-700' : 'text-slate-500'}>저장 {formatSavedDateTime(data.savedAt)}</span>
                          </span>
                        </td>
                        <td className="p-2 whitespace-nowrap">{data.date || '-'}</td>
                        <td className="p-2 text-slate-600 whitespace-nowrap">{summaryOf(data)}</td>
                        <td className="p-2 text-center"></td>
                        <td className="p-2 text-center">{actionButtons(data)}</td>
                      </tr>
                    );
                  }))}
                </Fragment>
              );
            })}
          </tbody>
        </table>
      </div>
    );
  };
  const menuSelectableReportKeys = profileCombinedReports.map(getReportKey).filter(Boolean);
  const menuSelectableReportKeySet = new Set(menuSelectableReportKeys);
  const menuCheckedReportKeysValid = menuCheckedReportKeys.filter((key) => menuSelectableReportKeySet.has(key));
  const menuCheckedReportKeySet = new Set(menuCheckedReportKeysValid);
  const isAllMenuReportsChecked =
    menuSelectableReportKeys.length > 0 && menuCheckedReportKeysValid.length === menuSelectableReportKeys.length;
  const profileSheetCounts = SHEET_MENU.reduce((acc, item) => {
    acc[item.id] = profileCombinedReports.filter(report => report._sheetId === item.id).length;
    return acc;
  }, {});
  const savedFitCount = savedData.filter(item => item.isokineticStatus === '적합').length;
  const savedFailCount = savedData.filter(item => item.isokineticStatus === '부적합').length;
  const savedAvgActualC = (() => {
    const nums = savedData
      .map(item => parseFloat(item.actualConcentration))
      .filter(v => Number.isFinite(v));
    if (nums.length === 0) return '-';
    return (nums.reduce((acc, cur) => acc + cur, 0) / nums.length).toFixed(2);
  })();
  const sheetSelectableReportKeys = savedData.map(getReportKey).filter(Boolean);
  const sheetSelectableReportKeySet = new Set(sheetSelectableReportKeys);
  const sheetCheckedReportKeysValid = sheetCheckedReportKeys.filter((key) => sheetSelectableReportKeySet.has(key));
  const sheetCheckedReportKeySet = new Set(sheetCheckedReportKeysValid);
  const isAllSheetReportsChecked =
    sheetSelectableReportKeys.length > 0 && sheetCheckedReportKeysValid.length === sheetSelectableReportKeys.length;

  const getConfiguredNozzles = () => {
    return activeNozzleSet
      .map((n) => ({ num: parseInt(n.num, 10), d: parseFloat(n.d) }))
      .filter((n) => !isNaN(n.num) && !isNaN(n.d) && n.d > 0)
      .sort((a, b) => a.num - b.num);
  };

  const saveNozzlePresets = () => {
    const sanitized = sanitizeNozzleSet(activeNozzleSet) || [];
    try {
      if (activeNozzleGroup === 'dust') {
        const targetKey = canUseUserScopedPresets(activeUser)
          ? getUserNozzlesStorageKey(activeUser)
          : STORAGE_KEYS.nozzles;
        localStorage.setItem(targetKey, JSON.stringify(sanitized));
      } else {
        localStorage.setItem(getGroupNozzlesStorageKey(activeUser), JSON.stringify({ ...groupNozzleSets, [activeNozzleGroup]: sanitized }));
      }
      setActiveNozzleSet(sanitized);
      alert(`${NOZZLE_GROUP_LABELS[activeNozzleGroup]} 노즐 프리셋을 저장했습니다.`);
    } catch (error) {
      console.error('노즐 프리셋 수동 저장 실패:', error);
      alert('노즐 프리셋 저장에 실패했습니다.');
    }
  };

  const handleNozzleConfigChange = (index, field, value) => {
    setActiveNozzleSet(prev => prev.map((item, i) => (
      i === index ? { ...item, [field]: value } : item
    )));
  };

  const addNozzleConfig = () => {
    setActiveNozzleSet(prev => [...prev, { num: '', d: '' }]);
  };

  const removeNozzleConfig = (index) => {
    setActiveNozzleSet(prev => prev.filter((_, i) => i !== index));
  };

  const resetNozzleDefaults = () => {
    setActiveNozzleSet(activeNozzleGroup === 'dust' ? cloneDefaultNozzles() : cloneDefaultGroupNozzleSets()[activeNozzleGroup]);
  };

  const handleSamplerPresetChange = (index, field, value) => {
    setSamplers(prev => {
      const next = prev.map((item, i) => (i === index ? { ...item, [field]: value } : item));
      const selected = next.find(s => String(s.id) === formData.samplerId);
      if (selected) {
        setFormData(current => ({
          ...current,
          gasMeterFactor: selected.yd,
          deltaHAt: selected.dHAt
        }));
      }
      return next;
    });
  };

  const saveSamplerPresets = () => {
    const sanitized = sanitizeSamplers(samplers) || [];
    try {
      const targetKey = canUseUserScopedPresets(activeUser)
        ? getUserSamplersStorageKey(activeUser)
        : STORAGE_KEYS.samplers;
      setPresetHistory(appendPresetHistory('sampler', readStoredList(targetKey), sanitized, activeUser));
      localStorage.setItem(targetKey, JSON.stringify(sanitized));
      setSamplers(sanitized);
      alert('샘플러 프리셋을 저장했습니다.');
    } catch (error) {
      console.error('샘플러 프리셋 수동 저장 실패:', error);
      alert('샘플러 프리셋 저장에 실패했습니다.');
    }
  };

  const saveGasSamplerPresets = () => {
    const sanitized = sanitizeGasSamplers(gasSamplers) || [];
    try {
      const targetKey = canUseUserScopedPresets(activeUser)
        ? getUserGasSamplersStorageKey(activeUser)
        : STORAGE_KEYS.gasSamplers;
      setPresetHistory(appendPresetHistory('gas', readStoredList(targetKey), sanitized, activeUser));
      localStorage.setItem(targetKey, JSON.stringify(sanitized));
      setGasSamplers(sanitized);
      alert('가스 샘플러 프리셋을 저장했습니다.');
    } catch (error) {
      console.error('가스 샘플러 프리셋 저장 실패:', error);
      alert('가스 샘플러 프리셋 저장에 실패했습니다.');
    }
  };
  const addGasSamplerPreset = () => {
    setGasSamplers(prev => {
      const nextId = prev.reduce((max, item) => Math.max(max, Number(item.id) || 0), 0) + 1;
      return [...prev, { id: nextId, sn: '', yd: '', calDate: '' }];
    });
  };
  const removeGasSamplerPreset = (index) => {
    setGasSamplers(prev => prev.filter((_, idx) => idx !== index));
  };
  const handleGasSamplerPresetChange = (index, field, value) => {
    setGasSamplers(prev => prev.map((item, idx) => (idx === index ? { ...item, [field]: value } : item)));
  };
  const resetGasSamplerDefaults = () => {
    setGasSamplers(cloneDefaultGasSamplers());
  };

  const addSamplerPreset = () => {
    setSamplers(prev => {
      const maxId = prev.reduce((max, item) => Math.max(max, Number(item.id) || 0), 0);
      return [...prev, { id: maxId + 1, name: `샘플러 ${maxId + 1}호기`, yd: '1.000', dHAt: '40.0' }];
    });
  };

  const removeSamplerPreset = (index) => {
    setSamplers(prev => {
      const removed = prev[index];
      const next = prev.filter((_, i) => i !== index);
      if (removed && String(removed.id) === formData.samplerId) {
        setFormData(current => ({ ...current, samplerId: '' }));
      }
      return next;
    });
  };

  const resetSamplerDefaults = () => {
    setSamplers(cloneDefaultSamplers());
    setFormData(prev => ({ ...prev, samplerId: '' }));
  };

  const handlePitotPresetConfigChange = (index, field, value) => {
    setPitotPresets(prev => {
      const next = prev.map((item, i) => (i === index ? { ...item, [field]: value } : item));
      const selected = next.find(item => String(item.id) === formData.pitotPresetId);
      if (selected) {
        const selectedSn = String(selected.sn || '').trim();
        setFormData(current => ({
          ...current,
          pitotFactor: selected.factor,
          pitotInfo: selectedSn || current.pitotInfo,
        }));
      }
      return next;
    });
  };

  const handlePitotCpRowChange = (presetIndex, rowIndex, field, value) => {
    setPitotPresets(prev => prev.map((item, i) => {
      if (i !== presetIndex) return item;
      const rows = Array.isArray(item.cpByVelocity) ? item.cpByVelocity : [];
      return { ...item, cpByVelocity: rows.map((row, r) => (r === rowIndex ? { ...row, [field]: value } : row)) };
    }));
  };

  const addPitotCpRow = (presetIndex) => {
    setPitotPresets(prev => prev.map((item, i) => (
      i === presetIndex
        ? { ...item, cpByVelocity: [...(Array.isArray(item.cpByVelocity) ? item.cpByVelocity : []), { v: '', cp: '' }] }
        : item
    )));
  };

  const removePitotCpRow = (presetIndex, rowIndex) => {
    setPitotPresets(prev => prev.map((item, i) => (
      i === presetIndex
        ? { ...item, cpByVelocity: (Array.isArray(item.cpByVelocity) ? item.cpByVelocity : []).filter((_, r) => r !== rowIndex) }
        : item
    )));
  };

  const savePitotPresets = () => {
    const sanitized = sanitizePitotPresets(pitotPresets) || [];
    try {
      const targetKey = canUseUserScopedPresets(activeUser)
        ? getUserPitotStorageKey(activeUser)
        : STORAGE_KEYS.pitotPresets;
      setPresetHistory(appendPresetHistory('pitot', readStoredList(targetKey), sanitized, activeUser));
      localStorage.setItem(targetKey, JSON.stringify(sanitized));
      setPitotPresets(sanitized);
      alert('피토 프리셋을 저장했습니다.');
    } catch (error) {
      console.error('피토 프리셋 수동 저장 실패:', error);
      alert('피토 프리셋 저장에 실패했습니다.');
    }
  };

  const addPitotPreset = () => {
    setPitotPresets(prev => {
      const maxId = prev.reduce((max, item) => Math.max(max, Number(item.id) || 0), 0);
      return [...prev, { id: maxId + 1, sn: '', usage: '', length: '', calDate: '', factor: '0.84', cpByVelocity: [] }];
    });
  };

  const removePitotPreset = (index) => {
    setPitotPresets(prev => {
      const removed = prev[index];
      const next = prev.filter((_, i) => i !== index);
      if (removed && String(removed.id) === formData.pitotPresetId) {
        setFormData(current => ({ ...current, pitotPresetId: '' }));
      }
      return next;
    });
  };

  const resetPitotDefaults = () => {
    setPitotPresets(cloneDefaultPitotPresets());
    setFormData(prev => ({ ...prev, pitotPresetId: '' }));
  };

  // 노즐/샘플러/피토 프리셋을 한 번에 저장
  const persistAllEquipmentPresets = (nozzles, samplerList, pitotList, gasSamplerList = null) => {
    const scoped = canUseUserScopedPresets(activeUser);
    localStorage.setItem(scoped ? getUserNozzlesStorageKey(activeUser) : STORAGE_KEYS.nozzles, JSON.stringify(nozzles));
    localStorage.setItem(scoped ? getUserSamplersStorageKey(activeUser) : STORAGE_KEYS.samplers, JSON.stringify(samplerList));
    localStorage.setItem(scoped ? getUserPitotStorageKey(activeUser) : STORAGE_KEYS.pitotPresets, JSON.stringify(pitotList));
    localStorage.setItem(scoped ? getUserGasSamplersStorageKey(activeUser) : STORAGE_KEYS.gasSamplers, JSON.stringify(gasSamplerList || sanitizeGasSamplers(gasSamplers) || []));
  };

  const saveAllEquipmentPresets = () => {
    const nozzles = sanitizeNozzleSet(nozzleSet) || [];
    const samplerList = sanitizeSamplers(samplers) || [];
    const pitotList = sanitizePitotPresets(pitotPresets) || [];
    const gasList = sanitizeGasSamplers(gasSamplers) || [];
    try {
      {
        const scoped = canUseUserScopedPresets(activeUser);
        appendPresetHistory('sampler', readStoredList(scoped ? getUserSamplersStorageKey(activeUser) : STORAGE_KEYS.samplers), samplerList, activeUser);
        appendPresetHistory('gas', readStoredList(scoped ? getUserGasSamplersStorageKey(activeUser) : STORAGE_KEYS.gasSamplers), gasList, activeUser);
        setPresetHistory(appendPresetHistory('pitot', readStoredList(scoped ? getUserPitotStorageKey(activeUser) : STORAGE_KEYS.pitotPresets), pitotList, activeUser));
      }
      persistAllEquipmentPresets(nozzles, samplerList, pitotList, gasList);
      const groupSets = sanitizeGroupNozzleSets(groupNozzleSets);
      localStorage.setItem(getGroupNozzlesStorageKey(activeUser), JSON.stringify(groupSets));
      setNozzleSet(nozzles);
      setGroupNozzleSets(groupSets);
      setSamplers(samplerList);
      setPitotPresets(pitotList);
      setGasSamplers(gasList);
      alert(`장비 프리셋을 저장했습니다.\n노즐 ${nozzles.length}개 (수은 ${groupSets.mercury.length} / PAHs ${groupSets.pahs.length}) / 샘플러 ${samplerList.length}개 / 가스 샘플러 ${gasList.length}개 / 피토관 ${pitotList.length}개`);
    } catch (error) {
      console.error('장비 프리셋 전체 저장 실패:', error);
      alert('장비 프리셋 저장에 실패했습니다.');
    }
  };

  // 설정 창의 "장비 프리셋 초기화": 이 PC에 저장된 노즐/샘플러/피토관 프리셋을 모두 지우고 기본 상태로 되돌림
  const resetAllEquipmentPresets = () => {
    const message = IS_DESKTOP_BUILD
      ? '이 PC에 저장된 노즐/샘플러/피토관 프리셋을 모두 지울까요?\n(지운 뒤에는 처음부터 다시 입력합니다.)'
      : '이 PC에 저장된 노즐/샘플러/피토관 프리셋을 모두 지우고 초기값으로 되돌릴까요?';
    if (!window.confirm(message)) return;
    const presetPrefixes = [
      STORAGE_KEYS.nozzles, STORAGE_KEYS.samplers, STORAGE_KEYS.pitotPresets, STORAGE_KEYS.gasSamplers,
      STORAGE_KEYS.nozzlesUserPrefix, STORAGE_KEYS.samplersUserPrefix, STORAGE_KEYS.pitotUserPrefix, STORAGE_KEYS.gasSamplersUserPrefix,
      STORAGE_KEYS.groupNozzles, STORAGE_KEYS.groupNozzlesUserPrefix,
    ];
    try {
      const staleKeys = [];
      for (let i = 0; i < localStorage.length; i++) {
        const key = localStorage.key(i);
        if (key && presetPrefixes.some((prefix) => key.startsWith(prefix))) staleKeys.push(key);
      }
      staleKeys.forEach((key) => localStorage.removeItem(key));
    } catch (error) {
      console.error('장비 프리셋 초기화 실패:', error);
    }
    const nozzles = cloneDefaultNozzles();
    const samplerList = cloneDefaultSamplers();
    const pitotList = cloneDefaultPitotPresets();
    const gasList = cloneDefaultGasSamplers();
    try {
      persistAllEquipmentPresets(nozzles, samplerList, pitotList, gasList);
    } catch (error) {
      console.error('장비 프리셋 초기화 저장 실패:', error);
    }
    setNozzleSet(nozzles);
    setGroupNozzleSets(cloneDefaultGroupNozzleSets());
    setSamplers(samplerList);
    setPitotPresets(pitotList);
    setGasSamplers(gasList);
    setFormData(prev => ({ ...prev, samplerId: '', pitotPresetId: '' }));
    alert('장비 프리셋을 초기화했습니다.');
  };

  // 설정 창의 "전체 초기화": 이 PC에 저장된 앱 데이터를 모두 지우고 처음 상태로 재시작
  const resetAllAppData = () => {
    const first = window.confirm(
      '전체 초기화를 진행할까요?\n\n이 PC에 저장된 모든 데이터가 삭제됩니다.\n- 모든 계정과 저장된 리포트\n- 노즐 / 샘플러 / 피토관 프리셋\n- 로그인 상태, 클라우드 동기화 설정, 앱 설정'
    );
    if (!first) return;
    const second = window.confirm('정말 삭제할까요? 삭제한 데이터는 복구할 수 없습니다.');
    if (!second) return;
    try {
      localStorage.clear();
    } catch (error) {
      console.error('전체 초기화 실패:', error);
      alert('전체 초기화에 실패했습니다.');
      return;
    }
    window.location.reload();
  };

  const persistReportsEncrypted = async (userName, reports, keyOverride = null) => {
    const targetUser = secureUsers[userName];
    if (!targetUser) throw new Error('사용자 정보가 없습니다.');
    const key = keyOverride || vaultKeyRef.current[userName];
    if (!key) throw new Error('잠금 해제가 필요합니다.');

    const encryptedData = await encryptJsonWithKey(key, reports);
    setSecureUsers(prev => {
      const current = prev[userName];
      if (!current) return prev;
      return {
        ...prev,
        [userName]: {
          ...current,
          dataIv: encryptedData.iv,
          dataCipher: encryptedData.cipher,
          updatedAt: new Date().toISOString(),
        },
      };
    });
  };

  const createVaultUserRecord = async (userId, password, reports = [], hintQuestion = '', hintAnswer = '') => {
    const salt = bytesToBase64(crypto.getRandomValues(new Uint8Array(16)));
    const key = await deriveVaultKey(password, salt);
    const verifierText = `vault:${userId}:verified`;
    const verifier = await encryptTextWithKey(key, verifierText);
    const encryptedData = await encryptJsonWithKey(key, reports);
    const hintAnswerHash = hintAnswer ? await hashHintAnswer(hintAnswer) : '';

    return {
      key,
      record: {
        salt,
        verifierIv: verifier.iv,
        verifierCipher: verifier.cipher,
        dataIv: encryptedData.iv,
        dataCipher: encryptedData.cipher,
        hintQuestion: hintQuestion || '',
        hintAnswerHash,
        avatarUrl: '',
        nickname: '',
        updatedAt: new Date().toISOString(),
      },
    };
  };

  const openUserSession = (userId, key, reports, userRecord = null) => {
    const fixedAvatarUrl = sanitizeAvatarUrl(userRecord && userRecord.avatarUrl);
    const userNozzles = getStoredNozzlesForUser(userId);
    const userSamplers = getStoredSamplersForUser(userId);
    const userGasSamplers = getStoredGasSamplersForUser(userId);
    const userPitotPresets = getStoredPitotPresetsForUser(userId);

    isPresetHydratingRef.current = true;
    vaultKeyRef.current[userId] = key;
    setActiveUser(userId);
    setLoginUserId(userId);
    localStorage.setItem(STORAGE_KEYS.lastLoginId, userId);
    setNozzleSet(userNozzles);
    setGroupNozzleSets(getStoredGroupNozzlesForUser(userId));
    setSamplers(userSamplers);
    setGasSamplers(userGasSamplers);
    setPitotPresets(userPitotPresets);
    setActiveUserReports(Array.isArray(reports) ? reports : []);
    setIsUserUnlocked(true);
    setLoginPassword('');
    setProfileAvatarUrl(fixedAvatarUrl);
    setProfileNickname((userRecord && userRecord.nickname) || '');

    setTimeout(() => {
      isPresetHydratingRef.current = false;
    }, 0);
  };

  const updateActiveUserProfile = (profilePatch) => {
    if (!activeUser || !isUserUnlocked) return;
    setSecureUsers(prev => {
      const current = prev[activeUser];
      if (!current) return prev;
      return {
        ...prev,
        [activeUser]: {
          ...current,
          ...profilePatch,
          updatedAt: new Date().toISOString(),
        },
      };
    });
  };

  const avatarFileInputRef = useRef(null);
  const handleProfileAvatarFile = async (event) => {
    const file = event.target.files && event.target.files[0];
    event.target.value = '';
    if (!file) return;
    if (!activeUser || !isUserUnlocked) { alert('로그인 후 프로필 사진을 바꿀 수 있습니다.'); return; }
    if (!/^image\//.test(file.type)) { alert('사진 파일(JPG, PNG 등)을 골라 주세요.'); return; }
    try {
      const dataUrl = sanitizeAvatarUrl(await resizeImageFileToAvatar(file));
      if (!dataUrl) throw new Error('avatar too large');
      updateActiveUserProfile({ avatarUrl: dataUrl });
      setProfileAvatarUrl(dataUrl);
    } catch (error) {
      console.error('프로필 사진 변경 실패:', error);
      alert('이 사진은 사용할 수 없습니다. 다른 사진을 골라 주세요.');
    }
  };
  const handleRemoveProfileAvatar = () => {
    if (!activeUser || !isUserUnlocked) return;
    if (!window.confirm('프로필 사진을 지우고 기본 모양으로 돌릴까요?')) return;
    updateActiveUserProfile({ avatarUrl: '' });
    setProfileAvatarUrl('');
  };

  const handleSaveProfileNickname = () => {
    if (!activeUser || !isUserUnlocked) {
      alert('로그인 후 별명을 저장할 수 있습니다.');
      return false;
    }

    const nextNickname = profileNickname.trim().slice(0, 24);
    updateActiveUserProfile({ nickname: nextNickname });
    setProfileNickname(nextNickname);
    alert('별명이 저장되었습니다.');
    return true;
  };

  const handleFindUserIdsByHint = async () => {
    const answer = recoverHintAnswer.trim();
    if (!answer) {
      alert('힌트 정답을 입력해주세요.');
      return;
    }
    const answerHash = await hashHintAnswer(answer);
    const matched = Object.entries(secureUsers)
      .filter(([, user]) => user && user.hintAnswerHash && user.hintAnswerHash === answerHash)
      .map(([userId]) => userId);

    setRecoveredUserIds(matched);
    if (matched.length === 0) {
      alert('일치하는 아이디를 찾지 못했습니다.');
    }
  };

  const handleShowPasswordHint = () => {
    const userId = recoverHintUserId.trim();
    if (!userId) {
      alert('아이디를 입력해주세요.');
      return;
    }
    const targetUser = secureUsers[userId];
    if (!targetUser) {
      alert('해당 아이디를 찾지 못했습니다.');
      setRecoveredHintQuestion('');
      return;
    }
    if (!targetUser.hintQuestion) {
      alert('이 계정은 힌트가 설정되어 있지 않습니다.');
      setRecoveredHintQuestion('');
      return;
    }
    setRecoveredHintQuestion(targetUser.hintQuestion);
  };

  const handleBetaQuickLogin = async () => {
    const userId = BETA_USER_ID;
    const password = BETA_USER_PASSWORD;

    try {
      let targetUser = secureUsers[userId];

      if (!targetUser) {
        const legacyUser = secureUsers[LEGACY_BETA_USER_ID];
        if (legacyUser) {
          try {
            const legacyKey = await deriveVaultKey(LEGACY_BETA_USER_PASSWORD, legacyUser.salt);
            const legacyVerifier = await decryptTextWithKey(legacyKey, legacyUser.verifierIv, legacyUser.verifierCipher);
            if (legacyVerifier === `vault:${LEGACY_BETA_USER_ID}:verified`) {
              const legacyDecrypted = await decryptJsonWithKey(legacyKey, legacyUser.dataIv, legacyUser.dataCipher);
              const legacyReports = Array.isArray(legacyDecrypted) ? legacyDecrypted.filter(item => item && typeof item === 'object') : [];
              const migrated = await createVaultUserRecord(userId, password, legacyReports);
              setSecureUsers(prev => ({ ...prev, [userId]: migrated.record }));
              openUserSession(userId, migrated.key, legacyReports, migrated.record);
              persistSessionAuth(userId, password);
              setAuthModal('');
              return;
            }
          } catch (legacyError) {
            console.error('legacy beta migration failed:', legacyError);
          }
        }

        const created = await createVaultUserRecord(userId, password, []);
        targetUser = created.record;
        setSecureUsers(prev => ({ ...prev, [userId]: created.record }));
        openUserSession(userId, created.key, [], created.record);
        persistSessionAuth(userId, password);
        setAuthModal('');
        return;
      }

      const key = await deriveVaultKey(password, targetUser.salt);
      const verifierText = await decryptTextWithKey(key, targetUser.verifierIv, targetUser.verifierCipher);
      if (verifierText !== `vault:${userId}:verified`) throw new Error('Invalid beta credentials');

      const decrypted = await decryptJsonWithKey(key, targetUser.dataIv, targetUser.dataCipher);
      const reports = Array.isArray(decrypted) ? decrypted.filter(item => item && typeof item === 'object') : [];
      const fixedAvatarUrl = sanitizeAvatarUrl(targetUser.avatarUrl);

      if (targetUser.avatarUrl !== fixedAvatarUrl) {
        setSecureUsers(prev => {
          const current = prev[userId];
          if (!current) return prev;
          return {
            ...prev,
            [userId]: {
              ...current,
              avatarUrl: fixedAvatarUrl,
              updatedAt: new Date().toISOString(),
            },
          };
        });
        targetUser = { ...targetUser, avatarUrl: fixedAvatarUrl };
      }

      openUserSession(userId, key, reports, targetUser);
      persistSessionAuth(userId, password);
      setAuthModal('');
    } catch (error) {
      console.error(error);
      alert('user 로그인 중 오류가 발생했습니다.');
    }
  };

  const openAdminLoginModal = () => {
    setLoginUserId('');
    setLoginPassword('');
    setAuthModal('admin');
  };

  const handleCreateUser = async () => {
    const trimmed = newUserName.trim();
    const normalizedId = trimmed.toLowerCase();
    const hintQuestion = newUserHintQuestion.trim();
    const hintAnswer = newUserHintAnswer.trim();
    if (!trimmed) {
      alert('사용자 이름을 입력해주세요.');
      return;
    }
    if (secureUsers[trimmed]) {
      alert('이미 존재하는 사용자 이름입니다.');
      return;
    }
    if (normalizedId === ADMIN_USER_ID && !(await isAdminPassword(newUserPassword))) {
      alert('관리자 계정 비밀번호는 지정된 값으로만 생성할 수 있습니다.');
      return;
    }
    if (newUserPassword !== newUserPasswordConfirm) {
      alert('비밀번호 확인이 일치하지 않습니다.');
      return;
    }
    if (!PASSWORD_RULE.test(newUserPassword)) {
      alert('비밀번호는 8자 이상, 영문/숫자/특수문자를 각각 1개 이상 포함해야 합니다.');
      return;
    }
    if (!hintQuestion || !hintAnswer) {
      alert('아이디/비밀번호 찾기용 힌트 질문과 정답을 입력해주세요.');
      return;
    }

    try {
      const { key, record } = await createVaultUserRecord(trimmed, newUserPassword, [], hintQuestion, hintAnswer);

      setSecureUsers(prev => ({
        ...prev,
        [trimmed]: record,
      }));
      openUserSession(trimmed, key, [], record);
      persistSessionAuth(trimmed, newUserPassword);
      setNewUserName('');
      setNewUserPassword('');
      setNewUserPasswordConfirm('');
      setNewUserHintQuestion('');
      setNewUserHintAnswer('');
      setAuthModal('');
      alert(`[${trimmed}] 계정을 만들고 로그인했습니다.`);
    } catch (error) {
      console.error(error);
      alert('사용자 생성 중 오류가 발생했습니다.');
    }
  };

  const handleLoginUser = async () => {
    const userId = loginUserId.trim();
    if (!userId) {
      alert('아이디를 입력해주세요.');
      return;
    }
    if (!loginPassword) {
      alert('비밀번호를 입력해주세요.');
      return;
    }

    const isAdminLogin = userId.toLowerCase() === ADMIN_USER_ID && (await isAdminPassword(loginPassword));

    try {
      let targetUser = secureUsers[userId];
      if (!targetUser && isAdminLogin) {
        const created = await createVaultUserRecord(userId, loginPassword, []);
        targetUser = created.record;
        setSecureUsers(prev => ({ ...prev, [userId]: created.record }));
        openUserSession(userId, created.key, [], created.record);
        persistSessionAuth(userId, loginPassword);
        setAuthModal('');
        return;
      }

      if (!targetUser) throw new Error('Invalid credentials');
      const key = await deriveVaultKey(loginPassword, targetUser.salt);
      const verifierText = await decryptTextWithKey(key, targetUser.verifierIv, targetUser.verifierCipher);
      if (verifierText !== `vault:${userId}:verified`) throw new Error('Invalid credentials');
      const decrypted = await decryptJsonWithKey(key, targetUser.dataIv, targetUser.dataCipher);
      const reports = Array.isArray(decrypted) ? decrypted.filter(item => item && typeof item === 'object') : [];

      const fixedAvatarUrl = sanitizeAvatarUrl(targetUser.avatarUrl);
      if (targetUser.avatarUrl !== fixedAvatarUrl) {
        setSecureUsers(prev => {
          const current = prev[userId];
          if (!current) return prev;
          return {
            ...prev,
            [userId]: {
              ...current,
              avatarUrl: fixedAvatarUrl,
              updatedAt: new Date().toISOString(),
            },
          };
        });
        targetUser = { ...targetUser, avatarUrl: fixedAvatarUrl };
      }

      openUserSession(userId, key, reports, targetUser);
      persistSessionAuth(userId, loginPassword);
      setAuthModal('');
    } catch (error) {
      console.error(error);

      if (isAdminLogin) {
        try {
          const recreated = await createVaultUserRecord(userId, loginPassword, []);
          setSecureUsers(prev => ({ ...prev, [userId]: recreated.record }));
          openUserSession(userId, recreated.key, [], recreated.record);
          persistSessionAuth(userId, loginPassword);
          setAuthModal('');
          alert('관리자 계정을 표준 비밀번호로 초기화하고 로그인했습니다.');
          return;
        } catch (adminError) {
          console.error(adminError);
        }
      }

      clearSessionAuth();
      setIsUserUnlocked(false);
      setActiveUserReports([]);
      setProfileAvatarUrl('');
      setProfileNickname('');
      alert('아이디 또는 비밀번호가 올바르지 않습니다.');
    }
  };

  const handleDeleteActiveUser = async () => {
    if (!activeUser || !isUserUnlocked) {
      alert('로그인 후 계정 삭제가 가능합니다.');
      return;
    }

    const willDelete = window.confirm(`[${displayUserNameWithFallback}] 계정을 삭제할까요?\n저장된 모든 리포트가 함께 삭제됩니다.`);
    if (!willDelete) return;

    const verifyPassword = window.prompt('삭제 확인을 위해 비밀번호를 다시 입력하세요.');
    if (verifyPassword === null) return;
    if (!verifyPassword) {
      alert('비밀번호를 입력해주세요.');
      return;
    }

    try {
      const targetUser = secureUsers[activeUser];
      if (!targetUser) throw new Error('Missing user');
      const key = await deriveVaultKey(verifyPassword, targetUser.salt);
      const verifierText = await decryptTextWithKey(key, targetUser.verifierIv, targetUser.verifierCipher);
      if (verifierText !== `vault:${activeUser}:verified`) throw new Error('Invalid password');

      setSecureUsers(prev => {
        const next = { ...prev };
        delete next[activeUser];
        return next;
      });
      delete vaultKeyRef.current[activeUser];
      clearSessionAuth();

      setActiveUser('');
      setIsUserUnlocked(false);
      setActiveUserReports([]);
      setLoginUserId(BETA_USER_ID);
      setLoginPassword('');
      setProfileAvatarUrl('');
      setProfileNickname('');
      setMenuCheckedReportKeys([]);
      setSheetCheckedReportKeys([]);
      setSelectedSheet('');
      if (window.location.hash) {
        window.history.pushState(null, '', window.location.pathname + window.location.search);
      }

      alert('계정이 삭제되었습니다.');
    } catch (error) {
      console.error(error);
      alert('비밀번호가 올바르지 않아 계정을 삭제할 수 없습니다.');
    }
  };

  const handleLockActiveUser = () => {
    if (activeUser) {
      delete vaultKeyRef.current[activeUser];
    }
    clearSessionAuth();
    setActiveUser('');
    setIsUserUnlocked(false);
    setActiveUserReports([]);
    // 다음에 로그인하는 사람에게 이전 사용자의 입력 내용이 넘어가지 않도록 화면 데이터도 비움
    setFormData(createInitialFormData());
    setSamplingLogData(createEmptySamplingLog());
    setLoginUserId(BETA_USER_ID);
    setLoginPassword('');
    setProfileAvatarUrl('');
    setProfileNickname('');
    setMenuCheckedReportKeys([]);
    setSheetCheckedReportKeys([]);
    setAuthModal('');
    setSelectedSheet('');
    if (window.location.hash) {
      window.history.pushState(null, '', window.location.pathname + window.location.search);
    }
  };

  const persistCloudSyncConfig = (nextConfig) => {
    if (!activeUser) return;
    const sanitized = sanitizeCloudSyncConfig(nextConfig) || { token: '', gistId: '' };
    setCloudSyncConfig(sanitized);
    try {
      localStorage.setItem(getUserCloudSyncStorageKey(activeUser), JSON.stringify(sanitized));
    } catch (error) {
      console.error('클라우드 동기화 설정 저장 실패:', error);
    }
  };

  const parseGithubApiError = async (response) => {
    try {
      const parsed = await response.json();
      if (typeof parsed?.message === 'string' && parsed.message.trim()) {
        return parsed.message.trim();
      }
    } catch (error) {
      // ignore json parse failure
    }
    return `HTTP ${response.status}`;
  };

  const buildCloudSyncPayload = () => {
    const userRecord = secureUsers[activeUser];
    if (!userRecord) throw new Error('현재 계정 데이터가 없습니다.');

    return {
      app: 'StackPilot',
      version: CLOUD_SYNC_VERSION,
      exportedAt: new Date().toISOString(),
      userId: activeUser,
      secureUser: userRecord,
      nozzles: sanitizeNozzleSet(nozzleSet) || [],
      groupNozzles: sanitizeGroupNozzleSets(groupNozzleSets),
      samplers: sanitizeSamplers(samplers) || [],
      gasSamplers: sanitizeGasSamplers(gasSamplers) || [],
      pitotPresets: sanitizePitotPresets(pitotPresets) || [],
      presetHistory: readPresetHistory(),
    };
  };

  const buildCloudSyncPayloadSignature = () => {
    const userRecord = secureUsers[activeUser];
    if (!userRecord || !activeUser) return '';
    try {
      return JSON.stringify({
        userId: activeUser,
        secureUser: userRecord,
        nozzles: sanitizeNozzleSet(nozzleSet) || [],
        groupNozzles: sanitizeGroupNozzleSets(groupNozzleSets),
        samplers: sanitizeSamplers(samplers) || [],
        gasSamplers: sanitizeGasSamplers(gasSamplers) || [],
        pitotPresets: sanitizePitotPresets(pitotPresets) || [],
        presetHistory: presetHistory.length,
      });
    } catch (error) {
      console.error('클라우드 시그니처 생성 실패:', error);
      return '';
    }
  };

  const handleSaveCloudSyncConfig = () => {
    if (!activeUser || !isUserUnlocked) {
      alert('로그인 후 설정 가능합니다.');
      return;
    }
    persistCloudSyncConfig(cloudSyncConfig);
    alert('클라우드 동기화 설정을 저장했습니다.');
  };

  const uploadCloudSyncToGist = async ({ silent = false } = {}) => {
    if (!activeUser || !isUserUnlocked) {
      if (!silent) alert('로그인 후 업로드 가능합니다.');
      return { ok: false, reason: 'not-authenticated' };
    }
    if (cloudSyncBusyRef.current) {
      if (!silent) alert('클라우드 동기화가 진행 중입니다.');
      return { ok: false, reason: 'busy' };
    }
    const token = String(cloudSyncConfig.token || '').trim();
    if (!token) {
      if (!silent) alert('GitHub PAT를 입력해주세요.');
      return { ok: false, reason: 'missing-token' };
    }

    try {
      cloudSyncBusyRef.current = true;
      setCloudSyncBusy(true);
      const payload = buildCloudSyncPayload();
      const content = JSON.stringify(payload, null, 2);
      const gistId = String(cloudSyncConfig.gistId || '').trim();
      const headers = {
        Accept: 'application/vnd.github+json',
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/json',
      };
      const requestBody = {
        description: `StackPilot Cloud Sync (${activeUser})`,
        files: {
          [CLOUD_SYNC_FILE_NAME]: { content },
        },
      };

      let response;
      if (gistId) {
        response = await fetch(`https://api.github.com/gists/${encodeURIComponent(gistId)}`, {
          method: 'PATCH',
          headers,
          body: JSON.stringify(requestBody),
        });
      } else {
        response = await fetch('https://api.github.com/gists', {
          method: 'POST',
          headers,
          body: JSON.stringify({ ...requestBody, public: false }),
        });
      }

      if (!response.ok) {
        const reason = await parseGithubApiError(response);
        throw new Error(reason);
      }

      const result = await response.json();
      const nextGistId = typeof result?.id === 'string' ? result.id : gistId;
      if (!nextGistId) throw new Error('Gist ID를 확인할 수 없습니다.');

      persistCloudSyncConfig({ token, gistId: nextGistId });
      if (!silent) alert('클라우드 백업 업로드가 완료되었습니다.');
      return { ok: true, gistId: nextGistId };
    } catch (error) {
      console.error('클라우드 업로드 실패:', error);
      if (!silent) alert(`클라우드 업로드 실패: ${error?.message || '알 수 없는 오류'}`);
      return { ok: false, reason: error?.message || 'upload-failed' };
    } finally {
      cloudSyncBusyRef.current = false;
      setCloudSyncBusy(false);
    }
  };

  // 백업 내용(클라우드 또는 백업 파일)을 이 기기에 적용한다. 이 기기에만 있는 기록은 합쳐서 남긴다
  const applySyncPayload = async (payload, { silent = false } = {}) => {
    if (!payload || typeof payload !== 'object' || payload.app !== 'StackPilot') throw new Error('StackPilot 백업 형식이 아닙니다.');
    if (String(payload.userId || '') !== activeUser) {
      throw new Error(`현재 로그인(${activeUser})과 백업 계정(${payload.userId || '알수없음'})이 다릅니다.`);
    }

    const sanitizedUserMap = sanitizeSecureUsers({ [activeUser]: payload.secureUser });
    const cloudUser = sanitizedUserMap ? sanitizedUserMap[activeUser] : null;
    if (!cloudUser) throw new Error('클라우드 사용자 데이터가 올바르지 않습니다.');

    const localKey = vaultKeyRef.current[activeUser];
    if (!localKey) throw new Error('복원을 위해 다시 로그인해주세요.');
    // 이 기기의 계정과 백업의 계정은 암호화 소금(salt)이 다를 수 있다 (캐시를 지우고 다시 만든 경우).
    // 기본 계정은 비밀번호가 정해져 있으므로 백업의 salt로 열쇠를 다시 만들어 연다
    let key = localKey;
    if (cloudUser.salt !== secureUsers[activeUser]?.salt) {
      if (activeUser === BETA_USER_ID) key = await deriveVaultKey(BETA_USER_PASSWORD, cloudUser.salt);
      else throw new Error('이 기기의 계정과 백업의 계정 암호가 달라 열 수 없습니다. 백업을 만든 기기에서 같은 계정으로 로그인해 주세요.');
    }
    const decrypted = await decryptJsonWithKey(key, cloudUser.dataIv, cloudUser.dataCipher);
    const cloudReports = Array.isArray(decrypted) ? decrypted.filter(item => item && typeof item === 'object') : [];

    // 이 PC에만 있는 리포트(업로드 실패 중 저장한 것 등)를 잃지 않도록 클라우드 것과 합친다.
    // 같은 리포트면 더 최근에 저장된 쪽을 남긴다.
    let localReports = [];
    try {
      const storedUsers = JSON.parse(localStorage.getItem(STORAGE_KEYS.secureUsers) || '{}');
      const localUser = storedUsers?.[activeUser];
      if (localUser?.dataIv && localUser?.dataCipher) {
        const localDecrypted = await decryptJsonWithKey(localKey, localUser.dataIv, localUser.dataCipher);
        if (Array.isArray(localDecrypted)) localReports = localDecrypted.filter(item => item && typeof item === 'object');
      }
    } catch (localError) {
      console.warn('로컬 리포트 읽기 실패, 클라우드 데이터만 사용:', localError);
    }
    const mergedByKey = new Map();
    [...cloudReports, ...localReports].forEach((report) => {
      const reportKey = getReportKey(report);
      if (!reportKey) return;
      const existing = mergedByKey.get(reportKey);
      const time = (item) => new Date(item?.savedAt || 0).getTime() || 0;
      if (!existing || time(report) >= time(existing)) mergedByKey.set(reportKey, report);
    });
    const restoredReports = [...mergedByKey.values()];
    const hasLocalOnly = restoredReports.length !== cloudReports.length
      || localReports.some((report) => !cloudReports.some((c) => getReportKey(c) === getReportKey(report) && c.savedAt === report.savedAt));

    const restoredNozzles = sanitizeNozzleSet(payload.nozzles);
    const groupNozzleKey = getGroupNozzlesStorageKey(activeUser);
    if (payload.groupNozzles && !(silent && localStorage.getItem(groupNozzleKey))) {
      const restoredGroups = sanitizeGroupNozzleSets(payload.groupNozzles);
      setGroupNozzleSets(restoredGroups);
      localStorage.setItem(groupNozzleKey, JSON.stringify(restoredGroups));
    }
    const restoredSamplers = sanitizeSamplers(payload.samplers);
    const restoredPitot = sanitizePitotPresets(payload.pitotPresets);
    const restoredGasSamplers = Array.isArray(payload.gasSamplers) ? sanitizeGasSamplers(payload.gasSamplers) : null;
    // 교정 이력은 백업의 것과 이 기기의 것을 합친다 (같은 줄은 한 번만)
    if (Array.isArray(payload.presetHistory)) {
      const seen = new Set();
      const merged = [...payload.presetHistory, ...readPresetHistory()]
        .filter((row) => row && typeof row === 'object')
        .filter((row) => { const key = `${row.at}|${row.kind}|${row.id}|${row.calDate}|${row.detail}|${row.note || ''}`; if (seen.has(key)) return false; seen.add(key); return true; })
        .sort((a, b) => String(a.at).localeCompare(String(b.at)));
      try { localStorage.setItem(PRESET_HISTORY_STORAGE_KEY, JSON.stringify(merged)); } catch { /* 이력만 못 살려도 나머지는 복원한다 */ }
      setPresetHistory(merged);
    }

    if (hasLocalOnly) {
      const mergedCipher = await encryptJsonWithKey(key, restoredReports);
      setSecureUsers(prev => ({
        ...prev,
        [activeUser]: { ...cloudUser, dataIv: mergedCipher.iv, dataCipher: mergedCipher.cipher, updatedAt: new Date().toISOString() },
      }));
    } else {
      setSecureUsers(prev => ({
        ...prev,
        [activeUser]: cloudUser,
      }));
    }
    setActiveUserReports(restoredReports);
    // 계정 기록을 백업 쪽으로 바꿨으므로 열쇠도 그에 맞춘다 (이후 저장이 같은 열쇠로 암호화되도록)
    vaultKeyRef.current[activeUser] = key;

    const nozzleTargetKey = canUseUserScopedPresets(activeUser)
      ? getUserNozzlesStorageKey(activeUser)
      : STORAGE_KEYS.nozzles;
    const samplerTargetKey = canUseUserScopedPresets(activeUser)
      ? getUserSamplersStorageKey(activeUser)
      : STORAGE_KEYS.samplers;
    const pitotTargetKey = canUseUserScopedPresets(activeUser)
      ? getUserPitotStorageKey(activeUser)
      : STORAGE_KEYS.pitotPresets;

    const hasLocalPreset = (storageKey) => {
      try { const v = JSON.parse(localStorage.getItem(storageKey) || 'null'); return Array.isArray(v) && v.length > 0; } catch { return false; }
    };
    if (restoredNozzles && restoredNozzles.length > 0 && !(silent && hasLocalPreset(nozzleTargetKey))) {
      setNozzleSet(restoredNozzles);
      localStorage.setItem(nozzleTargetKey, JSON.stringify(restoredNozzles));
    }
    if (restoredSamplers && restoredSamplers.length > 0 && !(silent && hasLocalPreset(samplerTargetKey))) {
      setSamplers(restoredSamplers);
      localStorage.setItem(samplerTargetKey, JSON.stringify(restoredSamplers));
    }
    const gasSamplerTargetKey = canUseUserScopedPresets(activeUser) ? getUserGasSamplersStorageKey(activeUser) : STORAGE_KEYS.gasSamplers;
    if (restoredGasSamplers && restoredGasSamplers.length > 0 && !(silent && hasLocalPreset(gasSamplerTargetKey))) {
      setGasSamplers(restoredGasSamplers);
      localStorage.setItem(gasSamplerTargetKey, JSON.stringify(restoredGasSamplers));
    }
    if (restoredPitot && restoredPitot.length > 0 && !(silent && hasLocalPreset(pitotTargetKey))) {
      setPitotPresets(restoredPitot);
      localStorage.setItem(pitotTargetKey, JSON.stringify(restoredPitot));
    }

    return restoredReports;
  };

  // 장비 프리셋(노즐·샘플러·가스 샘플러·피토관)을 엑셀 한 파일로 내보낸다
  const exportPresetsExcel = async () => {
    try {
      const sheets = buildPresetSheets({
        nozzles: { '먼지·금속': nozzleSet || [], 수은: groupNozzleSets?.mercury || [], PAHs: groupNozzleSets?.pahs || [] },
        samplers, gasSamplers, pitotPresets,
        history: presetHistory.map((row) => ({ ...row, kindLabel: PRESET_HISTORY_KINDS[row.kind]?.label || row.kind })),
      });
      const buffer = await buildWorkbook(sheets);
      const stamp = new Date().toISOString().slice(0, 10).replace(/-/g, '');
      const url = URL.createObjectURL(new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
      const link = document.createElement('a');
      link.href = url;
      link.setAttribute('download', `${stamp}_장비프리셋.xlsx`);
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      setTimeout(() => URL.revokeObjectURL(url), 1000);
    } catch (error) {
      console.error('프리셋 엑셀 만들기 실패:', error);
      alert('프리셋 엑셀을 만들지 못했습니다.');
    }
  };

  // ── 백업 파일: 설정 없이 쓰는 백업. 휴대폰이면 '파일' 앱(iCloud Drive 등)에 저장해 두면 캐시를 지워도 되살릴 수 있다 ──
  const backupFileInputRef = useRef(null);
  const exportBackupFile = () => {
    try {
      const payload = buildCloudSyncPayload();
      const stamp = new Date().toISOString().slice(0, 10).replace(/-/g, '');
      const blob = new Blob([JSON.stringify(payload)], { type: 'application/json' });
      const url = URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = url;
      link.setAttribute('download', `StackPilot_백업_${activeUser}_${stamp}.json`);
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      setTimeout(() => URL.revokeObjectURL(url), 1000);
      try { localStorage.setItem(`dust-sampling.last-backup.${activeUser}`, new Date().toISOString()); } catch { /* 무시 */ }
    } catch (error) {
      console.error('백업 파일 만들기 실패:', error);
      alert('백업 파일을 만들지 못했습니다.');
    }
  };
  const importBackupFile = async (event) => {
    const file = event.target.files && event.target.files[0];
    event.target.value = '';
    if (!file) return;
    if (!activeUser || !isUserUnlocked) { alert('로그인 후 불러올 수 있습니다.'); return; }
    try {
      const payload = JSON.parse(await file.text());
      const restored = await applySyncPayload(payload, { silent: false });
      alert(`백업을 불러왔습니다. 기록 ${restored.length}건`);
    } catch (error) {
      console.error('백업 파일 불러오기 실패:', error);
      alert(`백업 파일을 불러오지 못했습니다: ${error?.message || '알 수 없는 오류'}`);
    }
  };

  const restoreCloudSyncFromGist = async ({ silent = false } = {}) => {
    if (!activeUser || !isUserUnlocked) {
      if (!silent) alert('로그인 후 복원 가능합니다.');
      return { ok: false, reason: 'not-authenticated' };
    }
    if (cloudSyncBusyRef.current) {
      if (!silent) alert('클라우드 동기화가 진행 중입니다.');
      return { ok: false, reason: 'busy' };
    }
    const token = String(cloudSyncConfig.token || '').trim();
    const gistId = String(cloudSyncConfig.gistId || '').trim();
    if (!token) {
      if (!silent) alert('GitHub PAT를 입력해주세요.');
      return { ok: false, reason: 'missing-token' };
    }
    if (!gistId) {
      if (!silent) alert('Gist ID를 입력해주세요.');
      return { ok: false, reason: 'missing-gist-id' };
    }

    try {
      cloudSyncBusyRef.current = true;
      setCloudSyncBusy(true);
      const headers = {
        Accept: 'application/vnd.github+json',
        Authorization: `Bearer ${token}`,
      };
      const gistResponse = await fetch(`https://api.github.com/gists/${encodeURIComponent(gistId)}`, {
        method: 'GET',
        headers,
      });
      if (!gistResponse.ok) {
        const reason = await parseGithubApiError(gistResponse);
        throw new Error(reason);
      }

      const gistData = await gistResponse.json();
      const fileMap = gistData?.files || {};
      const targetFile = fileMap[CLOUD_SYNC_FILE_NAME] || Object.values(fileMap)[0];
      if (!targetFile) throw new Error('Gist 파일을 찾지 못했습니다.');

      let rawContent = typeof targetFile.content === 'string' ? targetFile.content : '';
      if ((!rawContent || targetFile.truncated) && targetFile.raw_url) {
        const rawResponse = await fetch(targetFile.raw_url, { headers });
        if (!rawResponse.ok) {
          const reason = await parseGithubApiError(rawResponse);
          throw new Error(reason);
        }
        rawContent = await rawResponse.text();
      }
      if (!rawContent) throw new Error('클라우드 파일 내용이 비어 있습니다.');

      const payload = JSON.parse(rawContent);
      const restoredReports = await applySyncPayload(payload, { silent });

      persistCloudSyncConfig({ token, gistId });
      setMenuCheckedReportKeys([]);
      setSheetCheckedReportKeys([]);
      if (!silent) alert(`클라우드 복원 완료: 리포트 ${restoredReports.length}건`);
      return { ok: true, restoredCount: restoredReports.length };
    } catch (error) {
      console.error('클라우드 복원 실패:', error);
      if (!silent) alert(`클라우드 복원 실패: ${error?.message || '알 수 없는 오류'}`);
      return { ok: false, reason: error?.message || 'restore-failed' };
    } finally {
      cloudSyncBusyRef.current = false;
      setCloudSyncBusy(false);
    }
  };

  const handleCloudUploadToGist = async () => {
    await uploadCloudSyncToGist({ silent: false });
  };

  const handleCloudRestoreFromGist = async () => {
    await restoreCloudSyncFromGist({ silent: false });
  };

  useEffect(() => {
    return () => {
      if (autoCloudUploadTimerRef.current) {
        clearTimeout(autoCloudUploadTimerRef.current);
      }
    };
  }, []);

  useEffect(() => {
    lastAutoCloudUploadSignatureRef.current = '';
    if (autoCloudUploadTimerRef.current) {
      clearTimeout(autoCloudUploadTimerRef.current);
      autoCloudUploadTimerRef.current = null;
    }
  }, [activeUser, cloudSyncConfig.token, cloudSyncConfig.gistId]);

  useEffect(() => {
    if (!isStorageHydrated || !activeUser || !isUserUnlocked) return;
    // 자동 복원은 메인 창에서만 (기록부 창을 열 때마다 복원·업로드가 겹치지 않게)
    if (IS_POPUP_WINDOW) return;
    const token = String(cloudSyncConfig.token || '').trim();
    const gistId = String(cloudSyncConfig.gistId || '').trim();
    if (!token || !gistId) return;
    const restoreKey = `${activeUser}|${token}|${gistId}`;
    if (autoCloudRestoreDoneRef.current[restoreKey]) return;

    autoCloudRestoreDoneRef.current[restoreKey] = true;
    autoCloudRestoreInFlightRef.current = true;
    void (async () => {
      await restoreCloudSyncFromGist({ silent: true });
      autoCloudRestoreInFlightRef.current = false;
    })();
  }, [isStorageHydrated, activeUser, isUserUnlocked, cloudSyncConfig.token, cloudSyncConfig.gistId]);

  useEffect(() => {
    if (!isStorageHydrated || !activeUser || !isUserUnlocked) return;
    if (cloudSyncBusyRef.current || autoCloudRestoreInFlightRef.current || autoCloudUploadInFlightRef.current) return;

    const token = String(cloudSyncConfig.token || '').trim();
    const gistId = String(cloudSyncConfig.gistId || '').trim();
    if (!token) return;
    if (gistId) {
      const restoreKey = `${activeUser}|${token}|${gistId}`;
      if (!autoCloudRestoreDoneRef.current[restoreKey]) return;
    }

    const nextSignature = buildCloudSyncPayloadSignature();
    if (!nextSignature || nextSignature === lastAutoCloudUploadSignatureRef.current) return;

    if (autoCloudUploadTimerRef.current) {
      clearTimeout(autoCloudUploadTimerRef.current);
      autoCloudUploadTimerRef.current = null;
    }

    // 업로드 실패(현장 인터넷 끊김 등) 뒤에는 1분 쉬었다가 다시 시도
    const failed = autoCloudUploadFailRef.current;
    const retryWait = failed && failed.signature === nextSignature
      ? Math.max(1200, 60000 - (Date.now() - failed.at))
      : 1200;
    autoCloudUploadTimerRef.current = setTimeout(async () => {
      if (autoCloudUploadInFlightRef.current || cloudSyncBusyRef.current || autoCloudRestoreInFlightRef.current) return;
      autoCloudUploadInFlightRef.current = true;
      const result = await uploadCloudSyncToGist({ silent: true });
      if (result?.ok) {
        lastAutoCloudUploadSignatureRef.current = nextSignature;
        autoCloudUploadFailRef.current = null;
      } else {
        autoCloudUploadFailRef.current = { signature: nextSignature, at: Date.now() };
        setCloudRetryTick((tick) => tick + 1);
      }
      autoCloudUploadInFlightRef.current = false;
      autoCloudUploadTimerRef.current = null;
    }, retryWait);
  }, [
    isStorageHydrated,
    activeUser,
    isUserUnlocked,
    cloudSyncConfig.token,
    cloudSyncConfig.gistId,
    cloudSyncBusy,
    activeUserReports,
    nozzleSet,
    groupNozzleSets,
    samplers,
    pitotPresets,
    secureUsers,
    cloudRetryTick,
  ]);

  const clearActiveUserReports = async () => {
    if (!activeUser) return;
    if (!isUserUnlocked) {
      alert('먼저 로그인해주세요.');
      return;
    }
    if (!window.confirm(`'${displayUserNameWithFallback}' 사용자의 저장 리포트를 모두 삭제할까요?`)) return;

    try {
      await persistReportsEncrypted(activeUser, []);
      setActiveUserReports([]);
      setMenuCheckedReportKeys([]);
      setSheetCheckedReportKeys([]);
    } catch (error) {
      console.error(error);
      alert('리포트 삭제 중 오류가 발생했습니다.');
    }
  };

  const removeSavedReport = async (id, savedAt) => {
    const targetKey = buildReportKey(id, savedAt);
    if (!targetKey) {
      alert('삭제할 리포트 키를 찾지 못했습니다.');
      return;
    }
    if (!activeUser || !isUserUnlocked) {
      alert('로그인 후 삭제할 수 있습니다.');
      return;
    }

    const nextReports = (activeUserReports || []).filter((item) => buildReportKey(item.id, item.savedAt) !== targetKey);

    try {
      await persistReportsEncrypted(activeUser, nextReports);
      setActiveUserReports(nextReports);
      setMenuCheckedReportKeys((prev) => prev.filter((key) => key !== targetKey));
      setSheetCheckedReportKeys((prev) => prev.filter((key) => key !== targetKey));
    } catch (error) {
      console.error(error);
      alert('리포트 삭제 중 오류가 발생했습니다.');
    }
  };

  const toggleMenuReportCheck = (reportKey, checked) => {
    if (!reportKey) return;
    setMenuCheckedReportKeys((prev) => {
      if (checked) {
        if (prev.includes(reportKey)) return prev;
        return [...prev, reportKey];
      }
      return prev.filter((key) => key !== reportKey);
    });
  };

  const toggleMenuAllReportChecks = (checked) => {
    setMenuCheckedReportKeys(checked ? menuSelectableReportKeys : []);
  };

  const handleDeleteCheckedMenuReports = async () => {
    if (!activeUser || !isUserUnlocked) {
      alert('로그인 후 선택 삭제가 가능합니다.');
      return;
    }
    if (menuCheckedReportKeysValid.length === 0) {
      alert('삭제할 리포트를 먼저 선택해주세요.');
      return;
    }
    if (!window.confirm(`선택한 ${menuCheckedReportKeysValid.length}건 리포트를 삭제할까요?`)) return;

    const selectedSet = new Set(menuCheckedReportKeysValid);
    const nextReports = (activeUserReports || []).filter(
      (item) => !selectedSet.has(buildReportKey(item.id, item.savedAt))
    );
    try {
      await persistReportsEncrypted(activeUser, nextReports);
      setActiveUserReports(nextReports);
      setMenuCheckedReportKeys([]);
      setSheetCheckedReportKeys((prev) => prev.filter((key) => !selectedSet.has(key)));
    } catch (error) {
      console.error(error);
      alert('선택 삭제 중 오류가 발생했습니다.');
    }
  };

  const toggleSheetReportCheck = (reportKey, checked) => {
    if (!reportKey) return;
    setSheetCheckedReportKeys((prev) => {
      if (checked) {
        if (prev.includes(reportKey)) return prev;
        return [...prev, reportKey];
      }
      return prev.filter((key) => key !== reportKey);
    });
  };

  const toggleSheetAllReportChecks = (checked) => {
    setSheetCheckedReportKeys(checked ? sheetSelectableReportKeys : []);
  };

  const handleDeleteCheckedSheetReports = async () => {
    if (!activeUser || !isUserUnlocked) {
      alert('로그인 후 선택 삭제가 가능합니다.');
      return;
    }
    if (sheetCheckedReportKeysValid.length === 0) {
      alert('삭제할 리포트를 먼저 선택해주세요.');
      return;
    }
    if (!window.confirm(`선택한 ${sheetCheckedReportKeysValid.length}건 리포트를 삭제할까요?`)) return;

    const selectedSet = new Set(sheetCheckedReportKeysValid);
    const nextReports = (activeUserReports || []).filter(
      (item) => !selectedSet.has(buildReportKey(item.id, item.savedAt))
    );

    try {
      await persistReportsEncrypted(activeUser, nextReports);
      setActiveUserReports(nextReports);
      setSheetCheckedReportKeys([]);
      setMenuCheckedReportKeys((prev) => prev.filter((key) => !selectedSet.has(key)));
    } catch (error) {
      console.error(error);
      alert('선택 삭제 중 오류가 발생했습니다.');
    }
  };

  const parseCsvLine = (line) => {
    const result = [];
    let curr = '';
    let inQuotes = false;
    for (let i = 0; i < line.length; i++) {
      const ch = line[i];
      if (ch === '"') {
        if (inQuotes && line[i + 1] === '"') {
          curr += '"';
          i += 1;
        } else {
          inQuotes = !inQuotes;
        }
      } else if (ch === ',' && !inQuotes) {
        result.push(curr);
        curr = '';
      } else {
        curr += ch;
      }
    }
    result.push(curr);
    return result.map((val) => val.trim());
  };

  const handleImportCsvReports = async (event) => {
    const file = event?.target?.files?.[0];
    if (!file) return;
    if (!activeUser || !isUserUnlocked) {
      alert('로그인 후 CSV 임포트가 가능합니다.');
      event.target.value = '';
      return;
    }

    try {
      const rawText = await file.text();
      const normalized = rawText.replace(/^\uFEFF/, '').replace(/\r\n/g, '\n').replace(/\r/g, '\n');
      const lines = normalized.split('\n').filter((line) => line.trim() !== '');
      if (lines.length < 2) {
        alert('CSV 내용이 비어 있거나 형식이 올바르지 않습니다.');
        event.target.value = '';
        return;
      }

      const headers = parseCsvLine(lines[0]);
      const importedRows = lines.slice(1).map((line) => {
        const cols = parseCsvLine(line);
        const row = {};
        headers.forEach((header, idx) => {
          row[header] = cols[idx] ?? '';
        });
        return row;
      });

      const importedReports = importedRows.map((row, idx) => {
        const sheetTitle = row['구분'] || '먼지시료채취기록부';
        const matchedSheet = SHEET_MENU.find((item) => item.title === sheetTitle);
        return {
          id: `csv-${Date.now()}-${idx}-${Math.random().toString(36).slice(2, 6)}`,
          savedAt: new Date().toISOString(),
          user: activeUser,
          sheetId: matchedSheet?.id || 'dust',
          sheetTitle: matchedSheet?.title || sheetTitle,
          date: row['측정일자'] || row['측정일시'] || '',
          company: row['사업장명'] || row['사업장'] || '',
          location: row['배출구명'] || row['배출구'] || '',
          moisturePercent: row['수분량(%)'] || '',
          isokineticRate: row['등속흡인율(%)'] || '',
          dustWeight: row['먼지무게(mg)'] || '',
          actualConcentration: row['실측농도'] || '',
          correctedConcentration: row['보정농도'] || '',
          remarks: row['비고'] || '',
          gasAnalyzer: [
            { time: '', o2: '', co2: '', co: '', nox: '', sox: '' },
            { time: '', o2: '', co2: '', co: '', nox: '', sox: '' },
            { time: '', o2: '', co2: '', co: '', nox: '', sox: '' },
          ],
        };
      }).filter((item) => item.date || item.company || item.location);

      if (importedReports.length === 0) {
        alert('임포트할 유효 리포트가 없습니다. CSV 헤더를 확인해주세요.');
        event.target.value = '';
        return;
      }

      const nextReports = [...activeUserReports, ...importedReports];
      await persistReportsEncrypted(activeUser, nextReports);
      setActiveUserReports(nextReports);
      alert(`${importedReports.length}건 CSV 리포트를 임포트했습니다.`);
    } catch (error) {
      console.error(error);
      alert('CSV 임포트 중 오류가 발생했습니다.');
    } finally {
      event.target.value = '';
    }
  };

  const handleLoadSavedReport = (report, silent = false) => {
    if (!report || typeof report !== 'object') return;
    if (!activeUser || !isUserUnlocked) {
      alert('로그인 후 불러올 수 있습니다.');
      return;
    }

    const targetSheet = resolveReportSheetId(report);
    // 메뉴(메인) 창에서는 해당 기록부 창을 열어 거기서 불러온다
    if (!IS_POPUP_WINDOW && !selectedSheet) {
      openSheetWindow(targetSheet, report);
      return;
    }
    if (targetSheet === 'samplinglog') {
      setSamplingLogData(sanitizeSamplingLog(report.samplingLog));
      unsavedRef.current = false;
      if (selectedSheet !== 'samplinglog') window.location.hash = 'sheet=samplinglog&popup=1';
      if (!silent) alert('통합시료채취기록부 리포트를 불러왔습니다.');
      return;
    }
    const targetImpingerCount = getImpingerStageCount(targetSheet);

    setFormData(() => {
      // 이전 기록지 내용이 섞이지 않도록 빈 기록지에서 시작해 리포트 값을 채운다
      const nextForm = { ...createInitialFormData() };
      Object.keys(nextForm).forEach((key) => {
        if (Object.prototype.hasOwnProperty.call(report, key)) {
          if (key === 'gasMeters') {
            const pointCount = Array.isArray(report.points) && report.points.length > 0
              ? report.points.length
              : nextForm.points.length;
            nextForm.gasMeters = normalizeGasMeterRows(report.gasMeters, pointCount);
          } else {
            nextForm[key] = report[key];
          }
        }
      });
      nextForm.impingers = normalizeImpingers(nextForm.impingers, targetImpingerCount);
      return nextForm;
    });

    setRecommendations(null);
    if (selectedSheet !== targetSheet) {
      window.location.hash = `sheet=${targetSheet}${IS_POPUP_WINDOW ? '&popup=1' : ''}`;
    }

    unsavedRef.current = false;
    const targetTitle = SHEET_MENU.find(item => item.id === targetSheet)?.title || report.sheetTitle || '기록부';
    if (!silent) alert(`${targetTitle} 리포트를 불러왔습니다.`);
  };

  const handleChange = (e) => {
    const { name, value } = e.target;
    setFormData(prev => {
      const newData = { ...prev, [name]: value };
      if (name === 'gasMeterFactor' || name === 'deltaHAt') newData.samplerId = '';
      if (name === 'pitotFactor') newData.pitotPresetId = '';
      if (name === 'totalStackDepth' || name === 'flangeLength') {
        const total = parseFloat(newData.totalStackDepth);
        const flange = parseFloat(newData.flangeLength) || 0; 
        if (!isNaN(total)) newData.stackDiameter = String(Math.max(0, total - flange));
      }
      return newData;
    });
  };

  const handleSamplerChange = (e) => {
    const id = e.target.value;
    if (id === '') { setFormData(prev => ({ ...prev, samplerId: '' })); return; }
    const sampler = samplers.find(s => s.id === parseInt(id, 10));
    if (sampler) setFormData(prev => ({ ...prev, samplerId: id, gasMeterFactor: sampler.yd, deltaHAt: sampler.dHAt }));
  };

  const applyPitotPresetById = (id) => {
    if (id === '' || id === null || id === undefined) {
      setFormData(prev => ({ ...prev, pitotPresetId: '' }));
      return;
    }
    const preset = pitotPresets.find(item => item.id === parseInt(id, 10));
    if (preset) {
      const nextInfo = String(preset.sn || '').trim();
      setFormData(prev => ({ ...prev, pitotPresetId: String(id), pitotFactor: preset.factor, pitotInfo: nextInfo || prev.pitotInfo }));
    }
  };

  const handlePitotPresetChange = (e) => applyPitotPresetById(e.target.value);

  // 선택된 피토관 프리셋에 유속별 Cp 표가 있으면, 예비조사 유속에 맞는 Cp를 자동 적용
  // 현재 기록부에 맞는 피토관 용도 (먼지/금속=입자, 수은, PAHs) 를 목록 앞쪽에 표시
  const isPitotUsageForSheet = (usage) => {
    const u = String(usage || '').trim();
    if (!u) return false;
    const sheetId = selectedSheet; // (activeSheet는 아래에서 정의되므로 state를 직접 사용)
    if (sheetId === 'mercury') return u.includes('수은');
    if (sheetId === 'pahs') return u.toUpperCase().includes('PAH');
    return u.includes('입자') || u.includes('먼지') || u.includes('금속');
  };
  // 이 기록부에 고를 수 있는 피토관: 용도가 맞는 것만 (수은=수은, PAHs=PAHs, 먼지·금속·유량=입자).
  // 용도를 적어 둔 피토관이 하나도 없으면(직접 등록한 경우 등) 전체를 보여 준다. 지금 골라 둔 것은 용도가 달라도 목록에 남긴다
  const sheetPitotPresets = (() => {
    const matched = pitotPresets.filter((preset) => isPitotUsageForSheet(preset.usage));
    if (matched.length === 0) return pitotPresets;
    const current = pitotPresets.find((preset) => String(preset.id) === String(formData.pitotPresetId));
    return current && !matched.includes(current) ? [...matched, current] : matched;
  })();
  const activePitotPreset = pitotPresets.find(item => String(item.id) === String(formData.pitotPresetId)) || null;
  const activePitotVelocityRows = activePitotPreset ? getPitotVelocityRows(activePitotPreset) : [];
  const isPitotCpAutoApplied = activePitotVelocityRows.length > 0;

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
      const stageCount = Math.max(getImpingerStageCount(selectedSheet), index + 1);
      const baseImpingers = normalizeImpingers(prev.impingers, stageCount);
      const newImpingers = baseImpingers.map((item, i) => (i === index ? { ...item, [field]: value } : item));
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
          // 동압은 소수 첫째 자리까지 (기준: 0.1 mmH₂O 단위로 읽는다)
          if (!isNaN(tp) && !isNaN(sp)) newItem.dp = (tp - sp).toFixed(1);
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
           if (!isNaN(dpValue) && !isNaN(kVal)) newItem.pressure = roundDH(kVal * dpValue);
           else if (value === '') newItem.pressure = '';
        }
        return newItem;
      });
      return { ...prev, gasMeters: newMeters };
    });
  };
  
  const handleKFactorChange = (e) => {
      const kValStr = e.target.value;
      setFormData(prev => {
          const kNum = parseFloat(kValStr);
          const newMeters = prev.gasMeters.map((meter, idx) => {
              if (idx === 0) return meter;
              const dpValue = parseFloat(meter.dp);
              if (Number.isFinite(kNum) && !isNaN(dpValue)) return { ...meter, pressure: roundDH(kNum * dpValue) };
              return meter;
          });
          return { ...prev, kFactor: kValStr, gasMeters: newMeters };
      });
  };

  const addGasMeter = () => setFormData(prev => ({ ...prev, gasMeters: [...prev.gasMeters, { id: prev.gasMeters.length, pointNum: '', time: '', stackTemp: '', dp: '', sp: '', pressure: '', volume: '', tmIn: '', tmOut: '', vacuum: '', impingerTemp: '' }] }));
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
    return {
      perRadius: isCenterOnly ? 1 : rnCoeffs.length,
      isCenterOnly,
      coeffs: rnCoeffs,
      diameterLabel: `${D} m`,
      area: area.toFixed(2),
      nearInsertion: nearDistances.map(d => ((d + flange) * 100).toFixed(2)),
      farInsertion: farDistances.map(d => ((d + flange) * 100).toFixed(2)),
    };
  };

  // 굴뚝 단면 측정점 그림 (논문 삽화식 흑백 선화): 등면적 원환, 지름선, 정방향 ●, 역방향 ○
  const renderSamplingPointDiagram = (sp) => {
    const width = 240;
    const height = 232;
    const cx = 134;
    const cy = 108;
    const R = 70;
    const ink = '#1d1d1f';
    const coeffs = sp.isCenterOnly ? [] : [...sp.coeffs].sort((a, b) => b - a); // 벽에 가까운 점부터 1번
    const nearPts = sp.isCenterOnly ? [{ x: cx, y: cy, n: 1 }] : coeffs.map((c, i) => ({ x: cx - c * R, y: cy, n: i + 1 }));
    const farPts = sp.isCenterOnly ? [] : coeffs.map((c, i) => ({ x: cx + c * R, y: cy, n: i + 1 }));
    return (
      <svg width={width} height={height} viewBox={`0 0 ${width} ${height}`} className="max-w-full h-auto" aria-label={`반경당 ${sp.perRadius}점 측정점 위치`} style={{ fontFamily: 'inherit' }}>
        <circle cx={cx} cy={cy} r={R + 6} fill="none" stroke={ink} strokeWidth="1.1" />
        <circle cx={cx} cy={cy} r={R} fill="none" stroke={ink} strokeWidth="0.8" />
        {coeffs.map((c, i) => (
          <circle key={i} cx={cx} cy={cy} r={c * R} fill="none" stroke="#8e8e93" strokeWidth="0.7" strokeDasharray="2.5 3" />
        ))}
        <line x1={cx - R} y1={cy} x2={cx + R} y2={cy} stroke={ink} strokeWidth="0.8" />
        <line x1={cx} y1={cy - 5} x2={cx} y2={cy + 5} stroke={ink} strokeWidth="0.8" />
        <rect x={cx - R - 22} y={cy - 4} width="14" height="8" fill="#fff" stroke={ink} strokeWidth="0.9" />
        <text x={cx - R - 26} y={cy + 3.5} fontSize="9" fill={ink} textAnchor="end">측정공</text>
        {farPts.map((pt) => (
          <g key={`far-${pt.n}`}>
            <circle cx={pt.x} cy={pt.y} r="3.4" fill="#fff" stroke="#d70015" strokeWidth="1.1" />
            <text x={pt.x} y={pt.y + 15} fontSize="9" fill="#d70015" textAnchor="middle">{pt.n}′</text>
          </g>
        ))}
        {nearPts.map((pt) => (
          <g key={`near-${pt.n}`}>
            <circle cx={pt.x} cy={pt.y} r="3.6" fill="#0a84ff" />
            <text x={pt.x} y={pt.y - 9} fontSize="9" fill="#0a84ff" textAnchor="middle">{pt.n}</text>
          </g>
        ))}
        <line x1={cx - R - 6} y1={cy + R + 20} x2={cx + R + 6} y2={cy + R + 20} stroke={ink} strokeWidth="0.6" />
        <line x1={cx - R - 6} y1={cy + R + 16} x2={cx - R - 6} y2={cy + R + 24} stroke={ink} strokeWidth="0.6" />
        <line x1={cx + R + 6} y1={cy + R + 16} x2={cx + R + 6} y2={cy + R + 24} stroke={ink} strokeWidth="0.6" />
        <text x={cx} y={cy + R + 36} fontSize="9" fill={ink} textAnchor="middle">D = {sp.diameterLabel || ''}</text>
      </svg>
    );
  };

  // 표 생성 계획: 지점 수에 맞는 권장 행 수(지점당)를 정하고, 총 시간을 그 행 수로 나눠 0.5분 단위 간격을 정한다
  // 총 시간 = 지점 수 × 지점당 행 수 × 행 간격 (모든 지점·행이 같은 시간)
  const getTableGenerationPlan = () => {
    const spData = getSamplingPoints();
    if (!spData) return null;
    const pointCount = Math.min(TABLE_PLAN_MAX_POINTS, Math.max(1, spData.perRadius));
    const requestedRaw = parseFloat(tableTotalMinutes);
    const planTotal = parseFloat(formData.planSamplingTime);
    const requestedTotal = Number.isFinite(requestedRaw) && requestedRaw > 0
      ? requestedRaw
      : (Number.isFinite(planTotal) && planTotal > 0 ? planTotal : DEFAULT_MEASURE_TOTAL_MINUTES);
    const rowsPerPoint = RECOMMENDED_ROWS_PER_POINT[pointCount] || 3;
    const measureRows = pointCount * rowsPerPoint;
    // 0.5분 단위로 올림: 총 채취시간이 목표보다 짧아지지 않게
    const intervalMinutes = Math.max(0.5, Math.ceil((requestedTotal / measureRows) * 2 - 1e-9) / 2);
    const adjustedTotal = Math.round(measureRows * intervalMinutes * 100) / 100;
    return {
      pointCount,
      requestedTotal,
      intervalMinutes,
      rowsPerPoint,
      measureRows,
      adjustedTotal,
      perPointMinutes: Math.round(rowsPerPoint * intervalMinutes * 100) / 100,
    };
  };

  const applySamplingPointsToTable = ({ silent = false } = {}) => {
    const spData = getSamplingPoints();
    if (!spData) {
      if (!silent) alert('1번 섹션에서 총 측정 깊이와 플랜지 길이를 입력해 굴뚝 내경이 계산되어야 측정점 표를 만들 수 있습니다.');
      return;
    }
    const plan = getTableGenerationPlan();
    if (!plan) return;
    const { pointCount, requestedTotal, intervalMinutes, rowsPerPoint, measureRows, adjustedTotal } = plan;
    setFormData(prev => {
      const newPoints = Array.from({ length: pointCount }, (_, i) => prev.points[i] || { id: i + 1, tp: '', sp: '', dp: '', ts: '' });
      const newMeters = createGasMeterRowsByPoints(pointCount, measureRows, intervalMinutes, adjustedTotal);
      return { ...prev, points: newPoints, gasMeters: newMeters };
    });
    if (silent) return;
    const adjustedNote = adjustedTotal !== requestedTotal
      ? `\n${requestedTotal}분은 ${intervalMinutes}분 간격으로 나누어떨어지지 않아 ${adjustedTotal}분으로 조정했습니다.`
      : '';
    alert(`측정점 ${pointCount}개 × 지점당 ${rowsPerPoint}행 × ${intervalMinutes}분 = 총 ${adjustedTotal}분 (시작행 포함 ${measureRows + 1}행)${adjustedNote}`);
  };

  const getGasComposition = () => {
    const getAnalyzerAvg = (field) => {
      const values = formData.gasAnalyzer
        .map((item) => parseFloat(item[field]))
        .filter((v) => Number.isFinite(v));
      if (values.length === 0) return 0;
      return values.reduce((acc, cur) => acc + cur, 0) / values.length;
    };

    const o2 = getAnalyzerAvg('o2');
    const co2 = getAnalyzerAvg('co2');
    const co = getAnalyzerAvg('co');
    const sox = getAnalyzerAvg('sox');
    const nox = getAnalyzerAvg('nox');
    
    // 엑셀(노즐산정!E28) 호환: CO = AVERAGE(D8:D10) / 1000
    const coPercent = co / 1000; 
    const n2 = 100 - co2 - o2 - coPercent; 
    
    // 엑셀(노즐산정!D28) 호환: CO 분자량 38 사용
    const M_O2 = 32, M_CO2 = 44, M_CO = 38, M_N2 = 28;
    const sumMx = (M_O2 * o2) + (M_CO2 * co2) + (M_CO * coPercent) + (M_N2 * n2);

    const Md = sumMx / 100;
    const Xw = getRawPreMoisture();
    const Ms = Md * ((100 - Xw) / 100) + 18 * (Xw / 100);

    const r0 = (1 / (22.4 * 100)) * ( sumMx * ((100 - Xw) / 100) + 18 * Xw );

    return { o2, co2, co, sox, nox, n2, Md, Ms, Xw, r0 };
  };

  // 방법별 사전 수분량(%) 계산. ready=false면 입력이 부족함
  const getMoistureByMethod = (methodId, data = formData) => {
    const num = (v) => { const n = parseFloat(v); return Number.isFinite(n) ? n : NaN; };
    const Pa = num(data.atmPressure);
    if (methodId === 'auto') {
      const vals = (data.moistureValues || []).map(num).filter(Number.isFinite);
      if (vals.length === 0) return { ready: false, value: NaN, detail: '5회 측정값을 입력하세요' };
      const value = vals.reduce((a, b) => a + b, 0) / vals.length;
      return { ready: true, value, detail: `${vals.length}회 평균` };
    }
    const impMethod = MOISTURE_METHODS.find((m) => m.id === methodId);
    if (isImpingerMoistureMethod(impMethod)) {
      const pre = data.preImpinger || createEmptyPreImpinger();
      const vlc = (pre.impingers || []).slice(0, impMethod.stages).reduce((sum, imp) => {
        const a = num(imp?.initial), b = num(imp?.final);
        return Number.isFinite(a) && Number.isFinite(b) ? sum + (b - a) : sum;
      }, 0);
      const vm = num(pre.meterEnd) - num(pre.meterStart);
      const yd = getPreMoistureYd(impMethod, data);
      const pv = getPreMoisturePv(impMethod, data);
      const tm = num(pre.meterTemp);
      // 측정공에서 기압(E28), 비어 있으면 대기압
      const holeP = Number.isFinite(num(pre.pressure)) ? num(pre.pressure) : Pa;
      const orifice = Number.isFinite(num(pre.orifice)) ? num(pre.orifice) : 0;
      if (!(vlc > 0) || !(vm > 0) || !Number.isFinite(tm) || !Number.isFinite(holeP)) {
        return { ready: false, value: NaN, detail: '측정전·후 무게, 가스미터 측정전·후, 가스미터온도, 기압이 필요합니다' };
      }
      // 엑셀 수분량!H28 = 22.4/18*E24/(C28*(273/(273+D28))*((E28+G28)/760)+22.4/18*E24)*100, C28=(B28-A28)*Yd, G28=F28/13.6
      const vw = (22.4 / 18) * vlc;
      // 건식 가스미터: (Pa+Pm)/760, 습식 가스미터: (Pa+Pm−Pv)/760 (ES 01114.c 식 9·10)
      const vmStd = vm * yd * (273 / (273 + tm)) * ((holeP + orifice / 13.6 - pv) / 760);
      const value = (vw / (vmStd + vw)) * 100;
      return { ready: Number.isFinite(value), value, detail: `흡습 수분 ${vlc.toFixed(2)} g · 흡입 ${vm.toFixed(2)} L` };
    }
    return { ready: false, value: NaN, detail: '' };
  };
  const activeMoistureMethod = MOISTURE_METHODS.find((m) => m.id === formData.moistureMethod) || MOISTURE_METHODS.find((m) => m.id === 'auto');

  const getRawPreMoisture = () => {
    const result = getMoistureByMethod(activeMoistureMethod.id);
    return result.ready ? result.value : 0;
  };

  const getRawAvgTp = () => {
    const validTps = formData.points.map(p => parseFloat(p.tp)).filter(v => !isNaN(v));
    return validTps.length === 0 ? 0 : validTps.reduce((a, b) => a + b, 0) / validTps.length;
  };

  const getRawAvgSp = () => {
    const validSps = formData.points.map(p => parseFloat(p.sp)).filter(v => !isNaN(v));
    return validSps.length === 0 ? 0 : validSps.reduce((a, b) => a + b, 0) / validSps.length;
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

  const getRawGasVelocity = (cpOverride = null) => {
    const avgTs = getRawAvgTs(), P = getRawStackPressure();
    const C = Number.isFinite(cpOverride) && cpOverride > 0 ? cpOverride : (parseFloat(formData.pitotFactor) || 0.84);
    const dpAvg = getRawAvgDp();
    if (dpAvg === 0 || isNaN(avgTs) || isNaN(P)) return 0;
    
    const gasComp = getGasComposition();
    if (isNaN(gasComp.r0) || gasComp.r0 === 0) return 0;
    
    const r = gasComp.r0 * (273 / (273 + avgTs)) * (P / 760);
    if (!Number.isFinite(r) || r <= 0) return 0;
    
    // 엑셀 수식: =D9*(2*9.81*J15/K15)^0.5
    // 각 점을 따로 계산하지 않고, '평균 동압(dpAvg)'을 식에 한 번에 넣어서 엑셀과 완전히 똑같이 계산합니다.
    return C * Math.pow((2 * 9.81 * dpAvg) / r, 0.5);
  };


  const isActiveMeterRow = (row, idx) => (
    idx === 0 || String(row?.pointNum || '').trim() !== ''
  );

  const getActiveMeterRows = () => (
    formData.gasMeters.filter((row, idx) => isActiveMeterRow(row, idx))
  );

  const getActiveMeterRowsUntil = (endIdx) => (
    formData.gasMeters
      .slice(0, endIdx + 1)
      .filter((row, idx) => isActiveMeterRow(row, idx))
  );

  // 엑셀 양식과 같이 등속·수분량·평균 계산은 시작행(0분, 17행)을 빼고 18행부터 평균
  const getSamplingRowsUntil = (endIdx) => getActiveMeterRowsUntil(endIdx).slice(1);

  const getRawOrificePressureFromRow = (row) => {
    const kVal = parseFloat(formData.kFactor);
    const dpVal = parseFloat(row?.dp);
    if (Number.isFinite(kVal) && Number.isFinite(dpVal)) return kVal * dpVal;
    const manualPressure = parseFloat(row?.pressure);
    return Number.isFinite(manualPressure) ? manualPressure : NaN;
  };

  const getRawCorrectedDensityFromAvgTs = (avgTs) => {
    const gasComp = getGasComposition();
    const Ps = getRawStackPressure();
    if (!Number.isFinite(avgTs) || !Number.isFinite(Ps) || Ps <= 0 || !Number.isFinite(gasComp.r0) || gasComp.r0 <= 0) return NaN;
    const density = gasComp.r0 * (273 / (273 + avgTs)) * (Ps / 760);
    return Number.isFinite(density) && density > 0 ? density : NaN;
  };

  const getRawCorrectedDensityFromRows = (rows) => {
    const tsValues = rows.map((row) => parseFloat(row?.stackTemp)).filter((v) => Number.isFinite(v));
    if (tsValues.length === 0) return NaN;
    const avgTs = tsValues.reduce((a, b) => a + b, 0) / tsValues.length;
    return getRawCorrectedDensityFromAvgTs(avgTs);
  };

  const getRawAvgCorrectedDensityFromMeterRows = () => {
    // 적산유량계 기록표에서 마지막 유효행의 누적 보정밀도(N행 방식) 사용
    // = r0 * 273/(273+AVERAGE(Ts_start:Ts_last)) * (Ps/760)
    // 엑셀 계산식!K24: 평균 배출가스 온도 D28(18행부터)
    const meterRows = getMeterRowsWithVolume().filter((row) => row !== formData.gasMeters[0]);
    if (meterRows.length === 0) return NaN;
    return getRawCorrectedDensityFromRows(meterRows);
  };

  const getRawGasVelocityFactorsFromMeterRows = (cpOverride = null) => {
    // 엑셀 계산식!G35: 평균 동압 E28·평균 온도 D28 모두 18행부터 (시작행 제외)
    const meterRows = getMeterRowsWithVolume().filter((row) => row !== formData.gasMeters[0]);
    const dpValues = meterRows
      .map((row) => parseFloat(row.dp))
      .filter((v) => Number.isFinite(v) && v >= 0);
    const tsValues = meterRows
      .map((row) => parseFloat(row.stackTemp))
      .filter((v) => Number.isFinite(v));

    if (dpValues.length === 0 || tsValues.length === 0) return null;

    const avgDp = dpValues.reduce((a, b) => a + b, 0) / dpValues.length;
    const avgTs = tsValues.reduce((a, b) => a + b, 0) / tsValues.length;
    const C = Number.isFinite(cpOverride) && cpOverride > 0 ? cpOverride : (parseFloat(formData.pitotFactor) || 0.84);
    const gasComp = getGasComposition();

    if (!Number.isFinite(avgDp) || avgDp <= 0 || !Number.isFinite(avgTs) || !Number.isFinite(C) || C <= 0) return null;
    if (!Number.isFinite(gasComp.r0) || gasComp.r0 <= 0) return null;

    // Vs = Cp * sqrt((2*9.81*평균동압) / (평균 보정밀도))
    const r = getRawAvgCorrectedDensityFromMeterRows();
    if (!Number.isFinite(r) || r <= 0) return null;

    const velocity = C * Math.pow((2 * 9.81 * avgDp) / r, 0.5);
    if (!Number.isFinite(velocity) || velocity <= 0) return null;
    return { velocity, avgDp, avgTs, C, r0: gasComp.r0, r, avgDensity: r };
  };

  const getRawGasVelocityFromMeterRows = (cpOverride = null) => {
    const factors = getRawGasVelocityFactorsFromMeterRows(cpOverride);
    return factors ? factors.velocity : NaN;
  };


  const getRawGasMeterVolDiff = () => {
    const initial = parseFloat(formData.gasMeters[0]?.volume);
    if (!Number.isFinite(initial)) return 0;

    const activeSamplingRows = formData.gasMeters.filter((row, idx) => idx > 0 && String(row?.pointNum || '').trim() !== '');
    const lastValidRow = [...activeSamplingRows].reverse().find((row) => {
      const volume = parseFloat(row?.volume);
      return Number.isFinite(volume);
    });
    if (!lastValidRow) return 0;

    const final = parseFloat(lastValidRow.volume);
    if (Number.isFinite(final) && final > initial) return final - initial;
    return 0;
  };

  const getRawAvgTm = () => {
    // 엑셀 H28 = AVERAGE(H18:I27): 시작행(0분)은 평균에서 제외
    const validTemps = getActiveMeterRows().slice(1).flatMap(g => {
      const inT = parseFloat(g.tmIn);
      const outT = parseFloat(g.tmOut);
      const values = [];
      if (Number.isFinite(inT)) values.push(inT);
      if (Number.isFinite(outT)) values.push(outT);
      return values;
    });
    return validTemps.length === 0 ? NaN : validTemps.reduce((a, b) => a + b, 0) / validTemps.length;
  };

  const getRawAvgSamplingTs = () => {
    const validTemps = getMeterRowsWithVolume().filter((row) => row !== formData.gasMeters[0]).map(g => parseFloat(g.stackTemp)).filter(v => Number.isFinite(v));
    return validTemps.length === 0 ? NaN : validTemps.reduce((a, b) => a + b, 0) / validTemps.length;
  };

  const getRawAvgOrifice = () => {
    // 엑셀 F28 = AVERAGE(F18:F27): 시작행(0분)은 평균에서 제외
    const validPressures = getActiveMeterRows().slice(1).map(g => getRawOrificePressureFromRow(g)).filter(v => Number.isFinite(v));
    return validPressures.length === 0 ? NaN : validPressures.reduce((a, b) => a + b, 0) / validPressures.length;
  };

  const getTraverseMeterTempAvg = () => {
    const values = [parseFloat(formData.traverseTmIn), parseFloat(formData.traverseTmOut)].filter(v => Number.isFinite(v));
    if (values.length === 0) return NaN;
    return values.reduce((acc, cur) => acc + cur, 0) / values.length;
  };

  const getRawTotalMoistureWeight = () => {
    let total = 0;
    formData.impingers.forEach(imp => {
      const init = parseFloat(imp.initial), fin = parseFloat(imp.final);
      if (!isNaN(init) && !isNaN(fin)) total += (fin - init);
    });
    return total;
  };

  // 유량측정기록부는 적산유량계가 없으므로 동압(Δp)이 입력된 행을 "측정 행"으로 본다
  const isFlowRecordSheet = selectedSheet === 'flow';
  const isMeterRowValid = (row) => (isFlowRecordSheet
    ? Number.isFinite(parseFloat(row?.dp))
    : Number.isFinite(parseFloat(row?.volume)));

  const getMeterRowsUntilLastVolume = () => {
    let lastValidIdx = -1;
    formData.gasMeters.forEach((row, idx) => {
      if (isMeterRowValid(row)) lastValidIdx = idx;
    });
    if (lastValidIdx < 0) return [];
    return formData.gasMeters.slice(0, lastValidIdx + 1);
  };

  const getMeterRowsWithVolume = () => (
    getMeterRowsUntilLastVolume().filter((row) => isMeterRowValid(row))
  );

  const getRawPostMoistureFactors = () => {
    // 엑셀 수분량!I44 계산 경로와 동일:
    // Xw = (22.4/18*F38) / (D44*(273/(273+E44))*((F44+H44)/760) + 22.4/18*F38) * 100
    const Wm = getRawTotalMoistureWeight(); // F38
    const meterRows = getMeterRowsUntilLastVolume();
    if (meterRows.length === 0) return null;

    const startVolume = parseFloat(meterRows[0]?.volume); // B44
    const endVolume = Math.max(...meterRows
      .map((row) => parseFloat(row?.volume))
      .filter((v) => Number.isFinite(v))); // C44(MAX)
    if (!Number.isFinite(startVolume) || !Number.isFinite(endVolume) || endVolume <= startVolume) return null;

    const deltaVmL = endVolume - startVolume;
    const yd = parseFloat(formData.gasMeterFactor) || 1.0;
    const pa = parseFloat(formData.atmPressure);
    if (!Number.isFinite(yd) || yd <= 0 || !Number.isFinite(pa)) return null;

    const samplingRows = meterRows.slice(1); // 엑셀 H28/F28: 18행부터
    const tmValues = samplingRows.flatMap((row) => {
      const inT = parseFloat(row.tmIn);
      const outT = parseFloat(row.tmOut);
      const values = [];
      if (Number.isFinite(inT)) values.push(inT);
      if (Number.isFinite(outT)) values.push(outT);
      return values;
    });
    if (tmValues.length === 0) return null;
    const tmAvg = tmValues.reduce((a, b) => a + b, 0) / tmValues.length; // E44

    const pmValues = samplingRows
      .map((row) => getRawOrificePressureFromRow(row))
      .filter((v) => Number.isFinite(v));
    if (pmValues.length === 0) return null;
    const pmAvg = pmValues.reduce((a, b) => a + b, 0) / pmValues.length; // G44

    const pmAbsTerm = pa + (pmAvg / 13.6); // F44 + H44
    if (!Number.isFinite(pmAbsTerm) || pmAbsTerm <= 0) return null;

    const vmStdL = deltaVmL * yd * (273 / (273 + tmAvg)) * (pmAbsTerm / 760);
    const vwStdL = Wm * (22.4 / 18);
    if (!Number.isFinite(vmStdL) || vmStdL <= 0) return null;

    let Xw = NaN;
    const denominator = vmStdL + vwStdL;
    if (Wm > 0 && Number.isFinite(vwStdL) && denominator > 0) {
      const rawXw = (vwStdL / denominator) * 100;
      if (Number.isFinite(rawXw) && rawXw >= 0) Xw = rawXw;
    }

    return { Xw, vmStdL, vwStdL, deltaVmL, tmAvg, pmAvg, pa, yd, Wm };
  };

  const getRawPostMoisture = () => {
    const factors = getRawPostMoistureFactors();
    return factors && Number.isFinite(factors.Xw) ? factors.Xw : 0;
  };

  const getRawVmStd = () => {
    const factors = getRawPostMoistureFactors();
    return factors ? factors.vmStdL : 0;
  };

  const getRawDustWeightDiff = () => {
    const initial = parseFloat(formData.filterInitial), final = parseFloat(formData.filterFinal);
    if (!isNaN(initial) && !isNaN(final)) return final - initial;
    return 0;
  };

  // -------------------------------------------------------------
  // 화면 표기(UI 표시) 전용 함수 모음
  // -------------------------------------------------------------
  const calcAvgTp = () => getRawAvgTp();
  const calcAvgSp = () => getRawAvgSp();
  const calcAvgDp = () => getRawAvgDp(); 
  const calcAvgTs = () => getRawAvgTs();
  const calcStackPressure = () => getRawStackPressure();
  const calcMoisture = () => getRawPreMoisture().toFixed(2);
  const calcPostMoisture = () => getRawPostMoisture().toFixed(2);
  const calcGasVelocity = () => { const v = getRawGasVelocity(); return v === 0 ? '0.00' : v.toFixed(2); };
  // 유속이 0으로 나올 때 어떤 입력이 빠졌는지 안내
  const calcFinalGasVelocity = () => {
    // 유량측정기록부: 4번 섹션 측정점(동압·온도) 기준 유속
    const v = isFlowRecordSheet ? getRawGasVelocity() : getRawGasVelocityFromMeterRows();
    if (!Number.isFinite(v) || v <= 0) return '-';
    return v.toFixed(2);
  };
  // 측정점별 유속(m/s): v_i = Cp × √(2g·Δp_i / r_i), r_i = r0 × 273/(273+Ts_i) × (Pa + Ps_i/13.6)/760
  const getRawPointGasVelocity = (point) => {
    const dp = parseFloat(point?.dp);
    const ts = parseFloat(point?.ts);
    const Pa = parseFloat(formData.atmPressure);
    const sp = parseFloat(point?.sp);
    const C = parseFloat(formData.pitotFactor) || 0.84;
    const gasComp = getGasComposition();
    if (!Number.isFinite(dp) || dp <= 0 || !Number.isFinite(ts) || !Number.isFinite(Pa)) return NaN;
    if (!Number.isFinite(gasComp.r0) || gasComp.r0 <= 0) return NaN;
    const absP = Pa + (Number.isFinite(sp) ? sp / 13.6 : 0);
    const r = gasComp.r0 * (273 / (273 + ts)) * (absP / 760);
    if (!Number.isFinite(r) || r <= 0) return NaN;
    return C * Math.sqrt((2 * 9.81 * dp) / r);
  };
  const getStackAreaM2 = () => {
    const D = parseFloat(formData.stackDiameter);
    return Number.isFinite(D) && D > 0 ? Math.PI * Math.pow(D / 2, 2) : NaN;
  };
  // 측정점별 실측 유량(㎥/min) = v_i × A × 60
  const getRawPointFlowRate = (point) => {
    const v = getRawPointGasVelocity(point);
    const A = getStackAreaM2();
    return Number.isFinite(v) && Number.isFinite(A) ? v * A * 60 : NaN;
  };
  const fmtOrDash = (v, digits = 2) => (Number.isFinite(v) ? v.toFixed(digits) : '-');
  const calcGasMeterVolDiff = () => { const v = getRawGasMeterVolDiff(); return v > 0 ? v.toFixed(2) : 0; };
  const getRawGasMeterVolDiffYd = () => {
    const vm = getRawGasMeterVolDiff();
    const y = parseFloat(formData.gasMeterFactor);
    if (!Number.isFinite(vm) || vm <= 0) return 0;
    if (!Number.isFinite(y) || y <= 0) return vm;
    return vm * y;
  };
  const calcGasMeterVolDiffYdL = () => {
    const v = getRawGasMeterVolDiffYd();
    return v > 0 ? v.toFixed(2) : '-';
  };
  const calcGasMeterVolDiffM3 = () => {
    const v = getRawGasMeterVolDiffYd();
    return v > 0 ? (v / 1000).toFixed(3) : '-';
  };
  const calcAvgTm = () => { const v = getRawAvgTm(); return isNaN(v) ? '-' : v.toFixed(2); };
  const calcAvgOrifice = () => { const v = getRawAvgOrifice(); return isNaN(v) ? '-' : v.toFixed(2); };
  const calcTotalMoistureWeight = () => getRawTotalMoistureWeight();
  const getVmStd = () => {
    const v = getRawVmStd();
    return v > 0 ? (v / 1000).toFixed(3) : '0.000';
  };
  const calcVmStdSL = () => {
    const v = getRawVmStd();
    return v > 0 ? v.toFixed(2) : '-';
  };
  const calcVmStdSm3 = () => {
    const v = getRawVmStd();
    return v > 0 ? (v / 1000).toFixed(3) : '-';
  };
  const calcDustWeightDiff = () => { const v = getRawDustWeightDiff(); return v.toFixed(4); };

  const hasMeasuredO2 = () => (formData.gasAnalyzer || []).some((item) => Number.isFinite(parseFloat(item?.o2)));
  const calcO2CorrectionFactor = () => {
    const o2_ref = parseFloat(formData.standardO2);
    const gasComp = getGasComposition();
    const o2_act = gasComp.o2;
    if (isNaN(o2_ref)) return 1.0;
    // 표준산소농도는 있는데 실측 O2가 없으면 0%로 보정하지 않고 계산 불가로 둔다
    if (!hasMeasuredO2() || isNaN(o2_act) || o2_act >= 21) return NaN;
    return (21 - o2_ref) / (21 - o2_act);
  };

  const getSamplingMinutes = () => {
    // 엑셀 B28 = MAX(B17:B27): 누적 시간의 최댓값. 미리 채워진 계획 행은 빼고 실제 측정한 행까지만 본다
    const lastMeasured = getLastMeasuredMeterIdx(formData.gasMeters);
    const times = formData.gasMeters
      .filter((meter, idx) => idx > 0 && idx <= lastMeasured && String(meter?.pointNum || '').trim() !== '')
      .map((meter) => parseFloat(meter.time))
      .filter((v) => Number.isFinite(v) && v > 0);

    if (times.length === 0) return 0;
    return Math.max(...times);
  };

  const getRawGasFlowFactors = () => {
    // 표준가스 유량:
    // 정압은 본 측정 전(측정점) 평균 정압,
    // 유속은 적산유량계 기록표 평균 동압으로 구한 값 사용
    // 건조 유량의 수분량(Xw)은 사후 수분량(임핀저 무게법) 기준
    const pointSp = formData.points.map((p) => parseFloat(p.sp)).filter((v) => Number.isFinite(v));
    const flowRows = getMeterRowsUntilLastVolume();
    const Ts = isFlowRecordSheet ? getRawAvgTs() : getRawAvgSamplingTs();
    const Vs = isFlowRecordSheet ? getRawGasVelocity() : getRawGasVelocityFromMeterRows();
    const Pa = parseFloat(formData.atmPressure);
    const D = parseFloat(formData.stackDiameter);

    if (!Number.isFinite(Ts) || !Number.isFinite(Vs) || Vs <= 0 || pointSp.length === 0 || !Number.isFinite(Pa) || !Number.isFinite(D) || D <= 0) {
      return null;
    }

    const spAvg = pointSp.reduce((a, b) => a + b, 0) / pointSp.length;
    const Ps = spAvg / 13.6;
    const absPressure = Pa + Ps;

    if (!Number.isFinite(absPressure) || absPressure <= 0) return null;

    // 중간 계산은 반올림 없이 원시값 사용 (엑셀과 동일한 계산 경로)
    const A = Math.PI * Math.pow(D / 2, 2);
    // 유량측정기록부는 임핀저 무게법이 없으므로 자동측정기 수분량(사전)을 사용
    const postFactors = isFlowRecordSheet ? null : getRawPostMoistureFactors();
    const postMoisture = isFlowRecordSheet
      ? getRawPreMoisture()
      : (postFactors && Number.isFinite(postFactors.Xw) ? postFactors.Xw : NaN);
    const moistureRatio = Math.min(1, Math.max(0, (Number.isFinite(postMoisture) ? postMoisture : 0) / 100));
    const tempTerm = 273 / (273 + Ts);
    const pressureTerm = absPressure / 760;
    const moistureTerm = 1 - moistureRatio;

    if (!Number.isFinite(tempTerm) || !Number.isFinite(pressureTerm) || !Number.isFinite(moistureTerm)) return null;

    return {
      Vs,
      A,
      Ts,
      Pa,
      Ps,
      XwPercent: postMoisture,
      flowRowCount: flowRows.length,
      moistureRatio,
      tempTerm,
      pressureTerm,
      moistureTerm,
      constant3600: 3600,
    };
  };

  const calcGasFlowRates = () => {
    const factors = getRawGasFlowFactors();
    if (!factors) return { dry: '-', wet: '-', dryRaw: NaN, wetRaw: NaN, factors: null };

    // Qn(건조) = v * A * 273/(273+Ts) * (Pa+Ps)/760 * (1 - Xw/100) * 3600
    const Q_wet = factors.Vs * factors.A * factors.tempTerm * factors.pressureTerm * factors.constant3600;
    // 사후 수분량(임핀저 무게)이 없으면 건조유량은 계산하지 않는다 (습식=건조로 보이지 않게)
    const Q_dry = Number.isFinite(factors.XwPercent) ? Q_wet * factors.moistureTerm : NaN;

    return {
      wet: formatFlowDisplay(Q_wet),
      dry: formatFlowDisplay(Q_dry),
      wetRaw: Q_wet,
      dryRaw: Q_dry,
      factors: {
        ...factors,
        Q_wet,
        Q_dry,
      },
    };
  };

  const generateRecommendation = (mode) => {
    const Vs = getRawGasVelocity(), dpAvg = getRawAvgDp();
    const targetVol = parseFloat(formData.targetVolume) || 1000;
    const configuredNozzles = getConfiguredNozzles();
    if (configuredNozzles.length === 0) return null;

    // 엑셀 수식의 AVERAGE(F15:G16)과 동일하게, 입력된 입/출구 Tm만 평균
    const Tm = getTraverseMeterTempAvg();

    const Ts = getRawAvgTs();
    const Pm = parseFloat(formData.atmPressure); 
    if (!Number.isFinite(Tm) || !Number.isFinite(Ts) || !Number.isFinite(Pm) || Pm <= 0) return null;
    
    const validSps = formData.points.map(p => parseFloat(p.sp)).filter(v => !isNaN(v));
    const avgSp = validSps.length === 0 ? 0 : validSps.reduce((a, b) => a + b, 0) / validSps.length;
    const Ps_ratio = (Pm + avgSp / 13.6) / Pm; 

    const gasComp = getGasComposition();
    const Md = gasComp.Md, Ms = gasComp.Ms, Xw = gasComp.Xw;
    if (!Number.isFinite(Md) || !Number.isFinite(Ms) || Ms === 0 || !Number.isFinite(Xw)) return null;

    const Cp = parseFloat(formData.pitotFactor) || 0.84;
    const Yd = parseFloat(formData.gasMeterFactor) || 1.0;
    const dHAt = parseFloat(formData.deltaHAt);

    const dhLimit = getDhTargets();
    let bestNozzle = null, maxDH = 0, closestDiffStandard = Infinity, closestDiffStable = Infinity;

    for (let n of configuredNozzles) {
        const Dn = n.d;
        const tempK = 8.001e-5 * Math.pow(Dn, 4) * dHAt * Math.pow(Cp, 2) * Math.pow(1 - Xw / 100, 2) * (Md / Ms) * ((Tm + 273) / (Ts + 273)) * Ps_ratio;
        const expectedDHRaw = tempK * dpAvg;
        if (!Number.isFinite(expectedDHRaw)) continue;

        if (expectedDHRaw <= dhLimit.max) {
            if (mode === 'fast') {
                if (expectedDHRaw > maxDH) { maxDH = expectedDHRaw; bestNozzle = n; }
            } else if (mode === 'standard') {
                const diff = Math.abs(expectedDHRaw - dhLimit.standard);
                if (diff < closestDiffStandard) { closestDiffStandard = diff; bestNozzle = n; }
            } else if (mode === 'stable') {
                const diff = Math.abs(expectedDHRaw - dhLimit.stable);
                if (diff < closestDiffStable) { closestDiffStable = diff; bestNozzle = n; }
            }
        }
    }

    if (!bestNozzle) bestNozzle = configuredNozzles[0];

    const Dn_best = bestNozzle.d;
    const An = Math.PI * Math.pow(Dn_best / 2000, 2);
    const Ps_abs = Pm + avgSp / 13.6;
    const bestQm = (Vs * An * 60 * 1000 * ((Tm + 273) / (Ts + 273)) * (Ps_abs / Pm) * (1 - Xw / 100)) / Yd;
    const Q_sl = bestQm * Yd * (273 / (Tm + 273)) * (Pm / 760);
    
    let fastestTime = Math.ceil(targetVol / Q_sl);
    if (fastestTime <= 0 || !isFinite(fastestTime)) fastestTime = 65;

    const K = 8.001e-5 * Math.pow(Dn_best, 4) * dHAt * Math.pow(Cp, 2) * Math.pow(1 - Xw / 100, 2) * (Md / Ms) * ((Tm + 273) / (Ts + 273)) * Ps_ratio;

    return { bestNozzleNum: bestNozzle.num, finalDn: Dn_best.toFixed(2), fastestTime, calculatedK: K, expectedDH: calcExpectedOrificeDH(K, dpAvg) };
  };

  const handleCalculateOptions = () => {
    const Vs = getRawGasVelocity(), Ts = getRawAvgTs(), Ps = getRawStackPressure(), dpAvg = getRawAvgDp();
    const Y = parseFloat(formData.gasMeterFactor), dHAt = parseFloat(formData.deltaHAt);
    const configuredNozzles = getConfiguredNozzles();
    const traverseTm = getTraverseMeterTempAvg();

    if (isNaN(Y) || Y <= 0 || isNaN(dHAt) || dHAt <= 0) { alert('교정 성적서의 보정계수(Yd)와 오리피스 계수(ΔH@)를 먼저 입력해주세요.'); return; }
    if (!Number.isFinite(traverseTm)) { alert('K-Factor 계산에는 미터온도(Tm) 입구/출구 값이 필요합니다.'); return; }
    if (Vs === 0 || isNaN(Vs) || isNaN(Ts) || isNaN(Ps) || isNaN(dpAvg) || dpAvg === 0) { alert('4번 섹션의 동압과 온도를 먼저 입력해야 산정할 수 있습니다.'); return; }
    if (configuredNozzles.length === 0) { alert('노즐 목록이 비어 있습니다. 노즐 설정에서 최소 1개 이상 입력해주세요.'); return; }

    const fast = generateRecommendation('fast');
    const standard = generateRecommendation('standard');
    const stable = generateRecommendation('stable');
    if (!fast || !standard || !stable) {
      alert('K-Factor 계산 입력값을 다시 확인해주세요. (Tm, Ts, 압력, 가스조성)');
      return;
    }

    setRecommendations({ fast, standard, stable });
  };

  // manual: 예측표에서 직접 고른 노즐 — 추천 표시는 그대로 두고 사용 노즐만 바꾼다
  const applyRecommendation = (opt, manual = false) => {
    const appliedK = parseFloat(opt.calculatedK);
    setFormData(prev => {
        const newMeters = prev.gasMeters.map((meter, idx) => {
            if (idx === 0) return meter;
            const dpValue = parseFloat(meter.dp);
            if (Number.isFinite(appliedK) && !isNaN(dpValue)) return { ...meter, pressure: roundDH(appliedK * dpValue) };
            return meter;
        });

        return { 
          ...prev, 
          usedNozzleNum: String(opt.bestNozzleNum), 
          nozzleDiameter: opt.finalDn, 
          kFactor: Number.isFinite(appliedK) ? String(opt.calculatedK) : '', 
          planSamplingTime: String(opt.fastestTime), 
          ...(manual ? {} : { recommendedNozzleNum: String(opt.bestNozzleNum), recommendedNozzleDia: opt.finalDn }),
          gasMeters: newMeters 
        };
    });
  };

  const calcRowIsokineticRate = (idx) => {
    if (idx === 0) return '-';
    if (String(formData.gasMeters[idx]?.pointNum || '').trim() === '') return '-';
    const startRow = formData.gasMeters[0];
    const currentRow = formData.gasMeters[idx];
    if (!startRow || !currentRow) return '-';

    const startVol = parseFloat(startRow.volume);
    const currentVol = parseFloat(currentRow.volume);
    if (!Number.isFinite(startVol) || !Number.isFinite(currentVol)) return '-';
    const VmCum = currentVol - startVol;
    if (!Number.isFinite(VmCum) || VmCum <= 0) return '-';

    const elapsedRows = getActiveMeterRowsUntil(idx).slice(1);
    const timeValues = elapsedRows
      .map((row) => parseFloat(row.time))
      .filter((v) => Number.isFinite(v) && v > 0);
    if (timeValues.length === 0) return '-';

    // 시료채취 시간은 누적 시간(엑셀 B열): 지금까지의 최댓값
    const thetaCum = Math.max(...timeValues);
    if (!Number.isFinite(thetaCum) || thetaCum <= 0) return '-';

    const rangeRows = getSamplingRowsUntil(idx);
    const tsValues = rangeRows.map(row => parseFloat(row.stackTemp)).filter(v => Number.isFinite(v));
    if (tsValues.length === 0) return '-';
    const TsAvg = tsValues.reduce((a, b) => a + b, 0) / tsValues.length;

    // 엑셀 AVERAGE($H$18:I19)처럼 입구/출구 모두 누적 평균 (시작행 제외)
    const tmValues = rangeRows.flatMap((row) => {
      const inT = parseFloat(row.tmIn);
      const outT = parseFloat(row.tmOut);
      const vals = [];
      if (Number.isFinite(inT)) vals.push(inT);
      if (Number.isFinite(outT)) vals.push(outT);
      return vals;
    });
    if (tmValues.length === 0) return '-';
    const TmAvg = tmValues.reduce((a, b) => a + b, 0) / tmValues.length;

    // 엑셀 AVERAGE($F$18:F19)처럼 오리피스압 누적 평균 (시작행 제외)
    const pmValues = rangeRows.map(row => getRawOrificePressureFromRow(row)).filter(v => Number.isFinite(v));
    if (pmValues.length === 0) return '-';
    const PmAvg = pmValues.reduce((a, b) => a + b, 0) / pmValues.length;

    const Pa = parseFloat(formData.atmPressure);
    const Ps = getRawStackPressure();
    const Y = parseFloat(formData.gasMeterFactor) || 1.0;
    if (!Number.isFinite(Pa) || !Number.isFinite(Ps) || Ps <= 0) return '-';
    const PmAbs = Pa + (PmAvg / 13.6);

    // 엑셀 M열과 동일하게 누적 평균 동압 기반의 배출가스속도를 사용
    const Vs = getRowGasVelocityRaw(idx);
    if (!Number.isFinite(Vs) || Vs <= 0) return '-';

    const Dn = parseFloat(formData.nozzleDiameter);
    if (!Number.isFinite(Dn) || Dn <= 0) return '-';
    const DnCm = Dn / 10;

    const vicCum = getRowMoistureCaptureRaw(idx);
    if (!Number.isFinite(vicCum) || vicCum < 0) return '-';

    // 엑셀 L열 수식과 동일한 누적 순간등속 (온도·오리피스 평균은 18행부터)
    const postTerm = (0.00346 * vicCum)
      + (((VmCum * Y) / 1000) / (273 + TmAvg)) * PmAbs;
    const rate = ((273 + TsAvg) * postTerm)
      / (Ps * thetaCum * Vs * (Math.PI / 4) * Math.pow(DnCm, 2))
      * 16670;

    if (!Number.isFinite(rate) || rate <= 0) return '-';
    return rate.toFixed(2);
  };

  const getRowCorrectedGasDensityRaw = (idx) => {
    if (idx < 0) return NaN;
    if (idx > 0 && String(formData.gasMeters[idx]?.pointNum || '').trim() === '') return NaN;
    const rangeRows = getActiveMeterRowsUntil(idx);
    return getRawCorrectedDensityFromRows(rangeRows);
  };

  const calcRowCorrectedGasDensity = (idx) => {
    const correctedDensity = getRowCorrectedGasDensityRaw(idx);
    if (!Number.isFinite(correctedDensity) || correctedDensity <= 0) return '-';
    return correctedDensity.toFixed(2);
  };

  const getRowGasVelocityRaw = (idx) => {
    if (idx <= 0) return NaN;
    if (String(formData.gasMeters[idx]?.pointNum || '').trim() === '') return NaN;
    const rangeRows = getActiveMeterRowsUntil(idx);
    const dpValues = rangeRows.map(row => parseFloat(row.dp)).filter(v => Number.isFinite(v) && v >= 0);
    if (dpValues.length === 0) return NaN;
    const dpAvg = dpValues.reduce((a, b) => a + b, 0) / dpValues.length;
    if (!Number.isFinite(dpAvg) || dpAvg <= 0) return NaN;

    const density = getRowCorrectedGasDensityRaw(idx);
    const C = parseFloat(formData.pitotFactor) || 0.84;
    if (!Number.isFinite(density) || density <= 0 || !Number.isFinite(C) || C <= 0) return NaN;

    const velocity = C * Math.pow((2 * 9.81 * dpAvg) / density, 0.5);
    if (!Number.isFinite(velocity) || velocity <= 0) return NaN;
    return velocity;
  };

  const calcRowGasVelocity = (idx) => {
    const velocity = getRowGasVelocityRaw(idx);
    if (!Number.isFinite(velocity) || velocity <= 0) return '-';
    return velocity.toFixed(2);
  };

  const getRowMoistureCaptureRaw = (idx) => {
    if (idx <= 0) return NaN;
    if (String(formData.gasMeters[idx]?.pointNum || '').trim() === '') return NaN;
    const startRow = formData.gasMeters[0];
    const currentRow = formData.gasMeters[idx];
    if (!startRow || !currentRow) return NaN;

    const startVol = parseFloat(startRow.volume);
    const currentVol = parseFloat(currentRow.volume);
    if (!Number.isFinite(startVol) || !Number.isFinite(currentVol)) return NaN;
    const VmCum = currentVol - startVol;
    if (!Number.isFinite(VmCum) || VmCum <= 0) return NaN;

    const rangeRows = getSamplingRowsUntil(idx);
    const tmValues = rangeRows.flatMap((row) => {
      const inT = parseFloat(row.tmIn);
      const outT = parseFloat(row.tmOut);
      const vals = [];
      if (Number.isFinite(inT)) vals.push(inT);
      if (Number.isFinite(outT)) vals.push(outT);
      return vals;
    });
    if (tmValues.length === 0) return NaN;
    const TmAvg = tmValues.reduce((a, b) => a + b, 0) / tmValues.length;

    const pmValues = rangeRows.map(row => getRawOrificePressureFromRow(row)).filter(v => Number.isFinite(v));
    if (pmValues.length === 0) return NaN;
    const PmAvg = pmValues.reduce((a, b) => a + b, 0) / pmValues.length;

    const Pa = parseFloat(formData.atmPressure);
    const Y = parseFloat(formData.gasMeterFactor) || 1.0;
    const Xw = getGasComposition().Xw;
    if (!Number.isFinite(Pa) || !Number.isFinite(TmAvg) || !Number.isFinite(PmAvg) || !Number.isFinite(Xw) || Xw <= 0 || Xw >= 100) return NaN;

    // 엑셀 O열 수식 동일: 누적 수분량(ml)
    return (((VmCum * Y * (273 / (273 + TmAvg)) * ((Pa + PmAvg / 13.6) / 760) * Xw) / (100 - Xw)) * (18 / 22.4));
  };

  const calcRowMoistureCapture = (idx) => {
    const raw = getRowMoistureCaptureRaw(idx);
    if (!Number.isFinite(raw) || raw < 0) return '-';
    return raw.toFixed(2);
  };

  const calcIsokineticRate = (isPost = false) => {
    const Vm = getRawGasMeterVolDiff();
    const traverseTs = getRawAvgTs();
    const samplingTs = getRawAvgSamplingTs();
    const Ts = Number.isFinite(samplingTs) ? samplingTs : traverseTs;
    const rawTm = getRawAvgTm();
    const Tm = isNaN(rawTm) ? Ts : rawTm; 
    const Dn = parseFloat(formData.nozzleDiameter);
    // 최종 등속흡인율은 적산유량계 기록표의 평균 동압 기반 유속을 우선 사용
    const meterBasedVs = getRawGasVelocityFromMeterRows();
    const Vs = Number.isFinite(meterBasedVs) && meterBasedVs > 0 ? meterBasedVs : getRawGasVelocity();
    const theta = getSamplingMinutes();
    const Y = parseFloat(formData.gasMeterFactor) || 1.0;
    const Pa = parseFloat(formData.atmPressure);
    const dH_avg = getRawAvgOrifice();
    const Ps = getRawStackPressure();

    if (Vm > 0 && Dn > 0 && Vs > 0 && theta > 0 && !isNaN(Ts) && !isNaN(Pa) && !isNaN(dH_avg) && !isNaN(Ps)) {
      const DnCm = Dn / 10;
      const AnCm2 = (Math.PI / 4) * Math.pow(DnCm, 2);
      const Pm_abs = Pa + dH_avg / 13.6;

      if (isPost) {
        // 엑셀식 대응:
        // =(((Ts+273)*(0.00346*Vic + (Vm*Y/1000)*(Pa+Pm)/(Tm+273))) / ((Pa+Ps)*theta*Vs*(PI()/4*Dn^2))) * 16670
        const Vic = getRawTotalMoistureWeight();
        const Vm_m3 = Vm / 1000;
        const part1 = 0.00346 * Vic;
        const part2 = ((Vm_m3 * Y) * Pm_abs) / (Tm + 273);
        const I = ((Ts + 273) * (part1 + part2)) / (Ps * theta * Vs * AnCm2) * 16670;
        return I.toFixed(2);
      } else {
        const gasComp = getGasComposition();
        const Xw = gasComp.Xw;
        if (!Number.isFinite(Xw) || Xw < 0 || Xw >= 100) return '-';

        // 엑셀 사전수분량 기준 대응:
        // Vic_pre = ((Vm*Y*(273/(273+Tm))*((Pa+Pm)/760))*Xw/(100-Xw))*(18/22.4)
        // I_pre = (((Ts+273)*(0.00346*Vic_pre + (Vm*Y/1000)*(Pa+Pm)/(Tm+273))) / (Ps*theta*Vs*(PI()/4*Dn^2))) * 16670
        const VicPre = ((Vm * Y * (273 / (273 + Tm)) * (Pm_abs / 760) * Xw) / (100 - Xw)) * (18 / 22.4);
        const Vm_m3 = Vm / 1000;
        const part1 = 0.00346 * VicPre;
        const part2 = ((Vm_m3 * Y) * Pm_abs) / (Tm + 273);
        const I = ((Ts + 273) * (part1 + part2)) / (Ps * theta * Vs * AnCm2) * 16670;
        return I.toFixed(2);
      }
    }
    return '-';
  };

  const calcDustConcentrations = () => {
    const Vm_std_sm3 = getRawVmStd() / 1000;
    const Wm = getRawDustWeightDiff();
    let actualC = '-', correctedC = '-';

    // 실측농도(mg/Sm3) = 채취무게(mg) / 표준가스량(Sm3)
    if (Vm_std_sm3 > 0 && Wm > 0) {
      const C = Wm / Vm_std_sm3;
      actualC = C.toFixed(2);
      const K = calcO2CorrectionFactor();
      correctedC = Number.isFinite(K) ? (C * K).toFixed(2) : '-';
    }
    return { actualC, correctedC };
  };

  const buildReportSnapshot = (targetSheetId = (selectedSheet || 'dust')) => {
    const { actualC, correctedC } = calcDustConcentrations();
    const flowRates = calcGasFlowRates();
    const isokineticRate = calcIsokineticRate(true);
    const isokineticNum = parseFloat(isokineticRate);
    const isokineticStatus = !isNaN(isokineticNum) && isokineticNum >= 90 && isokineticNum <= 110 ? '적합' : '부적합';
    const normalizedGasMeters = normalizeGasMeterRows(formData.gasMeters, formData.points.length);
    const sheetMeta = SHEET_MENU.find((item) => item.id === targetSheetId);
    const targetImpingers = normalizeImpingers(formData.impingers, getImpingerStageCount(targetSheetId));

    return {
      id: `${Date.now()}-${Math.random().toString(36).slice(2, 8)}`,
      savedAt: new Date().toISOString(),
      user: activeUser,
      sheetId: targetSheetId,
      sheetTitle: sheetMeta?.title || activeSheet?.title || '먼지시료채취기록부',
      ...formData,
      impingers: targetImpingers,
      gasMeters: normalizedGasMeters,
      sampler: formData.sampler || '',
      moisturePre: calcMoisture(),
      moisturePost: calcPostMoisture(),
      moisturePercent: calcTotalMoistureWeight() > 0 ? calcPostMoisture() : calcMoisture(),
      avgVelocity: calcFinalGasVelocity(),
      isokineticRate,
      isokineticStatus,
      dustWeight: calcDustWeightDiff(),
      actualConcentration: actualC,
      correctedConcentration: correctedC,
      wetFlowRate: flowRates.wet,
      dryFlowRate: flowRates.dry,
      gasMeterVolumeRaw: calcGasMeterVolDiff(),
      gasMeterVolume: calcGasMeterVolDiffYdL(),
      gasMeterVolumeM3: calcGasMeterVolDiffM3(),
      vmStd: getVmStd(),
      vmStdSL: calcVmStdSL(),
      vmStdSm3: calcVmStdSm3(),
      totalSamplingMinutes: String(getSamplingMinutes()),
      avgTm: calcAvgTm(),
      avgOrifice: calcAvgOrifice(),
      o2CorrectionFactor: Number.isFinite(calcO2CorrectionFactor()) ? calcO2CorrectionFactor().toFixed(3) : '-',
      nozzleUsed: formData.usedNozzleNum || formData.recommendedNozzleNum || '-',
      nozzleDiameterUsed: formData.nozzleDiameter || '-',
      kFactorApplied: formData.kFactor || '-',
    };
  };

  const handleSave = async (e, { silent = false } = {}) => {
    if (e && e.preventDefault) e.preventDefault();
    if (!activeUser) {
      alert('먼저 아이디/비밀번호로 로그인해주세요.');
      return false;
    }
    if (!isUserUnlocked) {
      alert('저장하려면 먼저 로그인해주세요.');
      return false;
    }
    const result = buildReportSnapshot(selectedSheet || 'dust');

    const nextReports = [...activeUserReports, result];
    try {
      await persistReportsEncrypted(activeUser, nextReports);
      setActiveUserReports(nextReports);
      const synced = syncSamplingLogFromSavedReport(result);
      unsavedRef.current = false;
      if (!silent) alert(`[${displayUserNameWithFallback}] 기록부를 저장했습니다.${synced ? '\n통합시료채취기록부에도 반영했습니다.' : ''}`);
      return true;
    } catch (error) {
      console.error(error);
      alert('암호화 저장 중 오류가 발생했습니다.');
      return false;
    }
  };

  const buildCopiedReportForSheet = (sourceReport, targetSheetId) => {
    const targetMeta = SHEET_MENU.find((item) => item.id === targetSheetId);
    const pointCount = Array.isArray(sourceReport?.points) && sourceReport.points.length > 0
      ? sourceReport.points.length
      : formData.points.length;
    const { _sheetId, _sheetTitle, ...baseReport } = sourceReport || {};

    return {
      ...baseReport,
      id: `${Date.now()}-${Math.random().toString(36).slice(2, 8)}`,
      savedAt: new Date().toISOString(),
      user: activeUser,
      sheetId: targetSheetId,
      sheetTitle: targetMeta?.title || baseReport.sheetTitle || '먼지시료채취기록부',
      impingers: normalizeImpingers(baseReport.impingers, getImpingerStageCount(targetSheetId)),
      gasMeters: normalizeGasMeterRows(baseReport.gasMeters, pointCount),
    };
  };

  const handleCopySavedReportToCurrentSheet = async (sourceReport) => {
    if (!activeUser) {
      alert('먼저 아이디/비밀번호로 로그인해주세요.');
      return;
    }
    if (!isUserUnlocked) {
      alert('로그인 후 복사할 수 있습니다.');
      return;
    }
    const targetSheetId = selectedSheet || 'dust';
    const sourceSheetId = resolveReportSheetId(sourceReport);
    if (sourceSheetId === targetSheetId) {
      alert('이미 현재 기록부와 같은 종류의 리포트입니다.');
      return;
    }

    const copiedReport = buildCopiedReportForSheet(sourceReport, targetSheetId);
    const nextReports = [...activeUserReports, copiedReport];
    const targetTitle = SHEET_MENU.find((item) => item.id === targetSheetId)?.title || targetSheetId;
    const sourceTitle = SHEET_MENU.find((item) => item.id === sourceSheetId)?.title || sourceSheetId;

    try {
      await persistReportsEncrypted(activeUser, nextReports);
      setActiveUserReports(nextReports);
      alert(`${sourceTitle} 리포트를 ${targetTitle}로 복사 저장했습니다.`);
    } catch (error) {
      console.error(error);
      alert('리포트 복사 저장 중 오류가 발생했습니다.');
    }
  };

  const handleCopyCheckedReportsToCurrentSheet = async () => {
    if (!activeUser) {
      alert('먼저 아이디/비밀번호로 로그인해주세요.');
      return;
    }
    if (!isUserUnlocked) {
      alert('로그인 후 복사할 수 있습니다.');
      return;
    }
    if (sheetCheckedReportKeysValid.length === 0) {
      alert('복사할 리포트를 먼저 선택해주세요.');
      return;
    }

    const targetSheetId = selectedSheet || 'dust';
    const selectedSet = new Set(sheetCheckedReportKeysValid);
    const selectedReports = savedData.filter((item) => selectedSet.has(buildReportKey(item.id, item.savedAt)));
    if (selectedReports.length === 0) {
      alert('선택된 리포트를 찾지 못했습니다.');
      return;
    }

    const copyTargets = selectedReports.filter((report) => resolveReportSheetId(report) !== targetSheetId);
    if (copyTargets.length === 0) {
      alert('선택된 리포트가 모두 현재 기록부와 같은 종류입니다.');
      return;
    }

    const copiedReports = copyTargets.map((report) => buildCopiedReportForSheet(report, targetSheetId));
    const nextReports = [...activeUserReports, ...copiedReports];
    const targetTitle = SHEET_MENU.find((item) => item.id === targetSheetId)?.title || targetSheetId;

    try {
      await persistReportsEncrypted(activeUser, nextReports);
      setActiveUserReports(nextReports);
      setSheetCheckedReportKeys([]);
      setMenuCheckedReportKeys((prev) => prev.filter((key) => !selectedSet.has(key)));
      alert(`선택 리포트 ${copiedReports.length}건을 ${targetTitle}로 복사 저장했습니다.`);
    } catch (error) {
      console.error(error);
      alert('선택 리포트 복사 저장 중 오류가 발생했습니다.');
    }
  };

  // Enter 키를 제출 대신 다음 입력 칸 이동(Tab 동작)으로 사용
  const handleFormKeyDown = (e) => {
    if (e.key !== 'Enter') return;

    const target = e.target;
    if (!target || !(target instanceof HTMLElement)) return;

    const tagName = target.tagName.toLowerCase();
    if (tagName !== 'input' && tagName !== 'select') return;
    if (target instanceof HTMLInputElement) {
      const inputType = (target.type || '').toLowerCase();
      if (inputType === 'submit' || inputType === 'button' || inputType === 'checkbox' || inputType === 'radio') {
        return;
      }
    }

    e.preventDefault();

    const form = e.currentTarget;
    if (!(form instanceof HTMLFormElement)) return;

    const focusable = Array.from(
      form.querySelectorAll('input:not([disabled]):not([type="hidden"]), select:not([disabled]), textarea:not([disabled]), button:not([disabled]), [tabindex]:not([tabindex="-1"])')
    ).filter((el) => el instanceof HTMLElement && (el.offsetParent !== null || el === document.activeElement));

    const currentIndex = focusable.indexOf(target);
    if (currentIndex === -1) return;

    const nextEl = focusable[currentIndex + 1] || focusable[0];
    if (nextEl && nextEl instanceof HTMLElement) {
      nextEl.focus();
      if (nextEl instanceof HTMLInputElement && ['text', 'number', 'search', 'email', 'tel', 'url', 'password', 'time', 'date'].includes(nextEl.type)) {
        nextEl.select?.();
      }
    }
  };

  const toNumber = (val) => {
    const n = typeof val === 'number' ? val : parseFloat(val);
    return Number.isFinite(n) ? n : null;
  };

  const toExcelDateSerial = (val) => {
    if (!val) return null;
    const dt = new Date(val);
    if (Number.isNaN(dt.getTime())) return null;
    return (Date.UTC(dt.getFullYear(), dt.getMonth(), dt.getDate()) - Date.UTC(1899, 11, 30)) / 86400000;
  };

  const toExcelTimeFraction = (val) => {
    if (!val || typeof val !== 'string' || !val.includes(':')) return null;
    const [h, m] = val.split(':').map(Number);
    if (!Number.isFinite(h) || !Number.isFinite(m)) return null;
    return ((h * 60) + m) / (24 * 60);
  };

  const parseCellRef = (ref) => {
    const match = /^([A-Z]+)(\d+)$/.exec(String(ref).toUpperCase());
    if (!match) return null;

    const [, colText, rowText] = match;
    let col = 0;
    for (let i = 0; i < colText.length; i++) {
      col = (col * 26) + (colText.charCodeAt(i) - 64);
    }

    const row = parseInt(rowText, 10);
    if (!Number.isFinite(row) || row <= 0) return null;
    return { ref: `${colText}${row}`, row, col };
  };

  const getChildElementsByLocalName = (node, localName) => {
    return Array.from(node.childNodes).filter(
      child => child.nodeType === 1 && child.localName === localName
    );
  };

  const findExistingRow = (sheetData, rowNum) => {
    const rows = getChildElementsByLocalName(sheetData, 'row');
    const rowKey = String(rowNum);
    return rows.find(row => row.getAttribute('r') === rowKey) || null;
  };

  const findExistingCell = (rowEl, cellRef) => {
    const cells = getChildElementsByLocalName(rowEl, 'c');
    return cells.find(cell => (cell.getAttribute('r') || '').toUpperCase() === cellRef) || null;
  };

  const splitCellRef = (cellRef) => {
    const match = /^([A-Z]+)(\d+)$/i.exec(String(cellRef || '').toUpperCase());
    if (!match) return null;
    return { col: match[1], row: parseInt(match[2], 10) };
  };

  const shiftFormulaRowRefs = (formula, delta) => {
    if (!formula || !Number.isFinite(delta) || delta === 0) return formula;
    return String(formula).replace(/(^|[^A-Z0-9_])(\$?[A-Z]{1,3})(\$?)(\d+)/g, (full, prefix, colPart, rowAbs, rowText) => {
      if (rowAbs === '$' || prefix === '!') return full; // 절대행, 다른 시트 참조는 그대로
      const row = parseInt(rowText, 10);
      if (!Number.isFinite(row)) return full;
      const shifted = Math.max(1, row + delta);
      return `${prefix}${colPart}${shifted}`;
    });
  };

  const shiftFormulaRowRefsFromRow = (formula, fromRow, delta) => {
    if (!formula || !Number.isFinite(delta) || delta === 0) return formula;
    // 다른 시트 참조(노즐산정!C39:C56 등)는 범위의 끝까지 통째로 보호한 뒤 이동하고 되돌린다
    const protectedRefs = [];
    const masked = String(formula).replace(
      /('(?:[^']|'')+'|[^\s!(),=+\-*/&<>:^;"{}]+)!\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?/g,
      (ref) => { protectedRefs.push(ref); return `\uE000${protectedRefs.length - 1}\uE000`; },
    );
    const shiftedFormula = masked.replace(/(^|[^A-Z0-9_\uE000])(\$?[A-Z]{1,3})(\$?)(\d+)/g, (full, prefix, colPart, rowAbs, rowText) => {
      if (rowAbs === '$' || prefix === '!') return full; // 절대행, 다른 시트 참조는 그대로
      const row = parseInt(rowText, 10);
      if (!Number.isFinite(row) || row < fromRow) return full;
      const shifted = Math.max(1, row + delta);
      return `${prefix}${colPart}${shifted}`;
    });
    return shiftedFormula.replace(/\uE000(\d+)\uE000/g, (_, i) => protectedRefs[Number(i)]);
  };

  const shiftCellRefRowsFrom = (refText, fromRow, delta) => {
    const cellMatch = /^([A-Z]+)(\d+)$/i.exec(String(refText || '').toUpperCase());
    if (cellMatch) {
      const rowNum = parseInt(cellMatch[2], 10);
      if (!Number.isFinite(rowNum) || rowNum < fromRow) return refText;
      return `${cellMatch[1]}${rowNum + delta}`;
    }

    const rangeMatch = /^([A-Z]+)(\d+):([A-Z]+)(\d+)$/i.exec(String(refText || '').toUpperCase());
    if (rangeMatch) {
      const startRow = parseInt(rangeMatch[2], 10);
      const endRow = parseInt(rangeMatch[4], 10);
      const nextStart = Number.isFinite(startRow) && startRow >= fromRow ? startRow + delta : startRow;
      const nextEnd = Number.isFinite(endRow) && endRow >= fromRow ? endRow + delta : endRow;
      return `${rangeMatch[1]}${nextStart}:${rangeMatch[3]}${nextEnd}`;
    }

    return refText;
  };

  const shiftSheetRowsDown = (sheetData, fromRow, delta) => {
    const rows = getChildElementsByLocalName(sheetData, 'row')
      .map((rowEl) => {
        const rowNum = parseInt(rowEl.getAttribute('r'), 10);
        return { rowEl, rowNum };
      })
      .filter((item) => Number.isFinite(item.rowNum) && item.rowNum >= fromRow)
      .sort((a, b) => b.rowNum - a.rowNum);

    rows.forEach(({ rowEl, rowNum }) => {
      const nextRowNum = rowNum + delta;
      rowEl.setAttribute('r', String(nextRowNum));
      getChildElementsByLocalName(rowEl, 'c').forEach((cellEl) => {
        const parsed = splitCellRef(cellEl.getAttribute('r'));
        if (!parsed) return;
        cellEl.setAttribute('r', `${parsed.col}${nextRowNum}`);
        const fNode = getChildElementsByLocalName(cellEl, 'f')[0];
        if (!fNode) return;
        const formulaText = String(fNode.textContent || '').trim();
        if (!formulaText) return;
        fNode.textContent = shiftFormulaRowRefsFromRow(formulaText, fromRow, delta);
      });
    });
  };

  const shiftMergedCellsRowsDown = (worksheet, fromRow, delta) => {
    const mergeCellsNode = worksheet.getElementsByTagNameNS('*', 'mergeCells')[0];
    if (!mergeCellsNode) return;
    const mergeCells = getChildElementsByLocalName(mergeCellsNode, 'mergeCell');
    mergeCells.forEach((mergeCell) => {
      const ref = mergeCell.getAttribute('ref');
      if (!ref) return;
      mergeCell.setAttribute('ref', shiftCellRefRowsFrom(ref, fromRow, delta));
    });
  };

  const buildSharedFormulaInfoMap = (sheetData) => {
    const map = new Map();
    const rows = getChildElementsByLocalName(sheetData, 'row');
    rows.forEach((rowEl) => {
      const rowNum = parseInt(rowEl.getAttribute('r'), 10);
      if (!Number.isFinite(rowNum)) return;
      getChildElementsByLocalName(rowEl, 'c').forEach((cellEl) => {
        const fNode = getChildElementsByLocalName(cellEl, 'f')[0];
        if (!fNode) return;
        if (fNode.getAttribute('t') !== 'shared') return;
        const si = fNode.getAttribute('si');
        const text = String(fNode.textContent || '').trim();
        if (!si || !text) return;
        if (!map.has(si)) map.set(si, { formula: text, row: rowNum });
      });
    });
    return map;
  };

  const insertClonedMeterRow = (sheetData, sourceRowNum, targetRowNum, sharedFormulaMap) => {
    const sourceRow = findExistingRow(sheetData, sourceRowNum);
    if (!sourceRow) return;

    const newRow = sourceRow.cloneNode(true);
    newRow.setAttribute('r', String(targetRowNum));

    getChildElementsByLocalName(newRow, 'c').forEach((cellEl) => {
      const parsed = splitCellRef(cellEl.getAttribute('r'));
      if (parsed) {
        cellEl.setAttribute('r', `${parsed.col}${targetRowNum}`);
      }

      const fNode = getChildElementsByLocalName(cellEl, 'f')[0];
      if (fNode) {
        const formulaType = fNode.getAttribute('t');
        const si = fNode.getAttribute('si');
        let sourceFormula = String(fNode.textContent || '').trim();

        if (formulaType === 'shared' && si && sharedFormulaMap.has(si)) {
          const master = sharedFormulaMap.get(si);
          sourceFormula = shiftFormulaRowRefs(master.formula, sourceRowNum - master.row);
        }

        if (sourceFormula) {
          const shiftedFormula = shiftFormulaRowRefs(sourceFormula, targetRowNum - sourceRowNum);
          fNode.removeAttribute('t');
          fNode.removeAttribute('si');
          fNode.removeAttribute('ref');
          fNode.textContent = shiftedFormula;
        } else {
          cellEl.removeChild(fNode);
        }

        getChildElementsByLocalName(cellEl, 'v').forEach((vNode) => cellEl.removeChild(vNode));
      }
    });

    const nextRow = getChildElementsByLocalName(sheetData, 'row').find((rowEl) => {
      const rowNum = parseInt(rowEl.getAttribute('r'), 10);
      return Number.isFinite(rowNum) && rowNum > targetRowNum;
    });
    if (nextRow) sheetData.insertBefore(newRow, nextRow);
    else sheetData.appendChild(newRow);
  };

  const expandRecordSheetMeterRows = (worksheet, requiredRowCount, layout = {}) => {
    const sheetData = worksheet.getElementsByTagNameNS('*', 'sheetData')[0];
    if (!sheetData) return;

    const meterStartRow = Number.isFinite(layout.meterStartRow) ? layout.meterStartRow : 17;
    const meterTemplateEndRow = Number.isFinite(layout.meterTemplateEndRow) ? layout.meterTemplateEndRow : 26;
    const templateCapacity = meterTemplateEndRow - meterStartRow + 1;
    if (!Number.isFinite(requiredRowCount) || requiredRowCount <= templateCapacity) return;

    const extraRows = requiredRowCount - templateCapacity;
    const sharedFormulaMap = buildSharedFormulaInfoMap(sheetData);

    for (let i = 0; i < extraRows; i += 1) {
      const insertRow = meterTemplateEndRow + 1 + i;
      const sourceRowNum = insertRow - 1;
      shiftSheetRowsDown(sheetData, insertRow, 1);
      shiftMergedCellsRowsDown(worksheet, insertRow, 1);
      insertClonedMeterRow(sheetData, sourceRowNum, insertRow, sharedFormulaMap);
    }

    const summaryRowNum = meterTemplateEndRow + 1 + extraRows;
    const meterEndRow = meterTemplateEndRow + extraRows;
    const summaryRow = findExistingRow(sheetData, summaryRowNum);
    if (summaryRow) {
      getChildElementsByLocalName(summaryRow, 'c').forEach((cellEl) => {
        const fNode = getChildElementsByLocalName(cellEl, 'f')[0];
        if (!fNode) return;
        const formulaText = String(fNode.textContent || '').trim();
        if (!formulaText) return;
        const extended = formulaText.replace(
          new RegExp(`(\\$?[A-Z]+\\$?(?:${meterStartRow}|${meterStartRow + 1}):\\$?[A-Z]+\\$?)(\\d+)`, 'g'),
          (full, head) => `${head}${meterEndRow}`
        ).replace(
          new RegExp(`(^|[^A-Z0-9_:$])([A-Z]{1,3})${meterTemplateEndRow}(?![0-9:])`, 'g'),
          (full, prefix, col) => `${prefix}${col}${meterEndRow}`
        );
        fNode.textContent = extended;
      });
    }

    // 측정행 위쪽(헤더 영역)의 수식이 평균행 등 아래 행을 가리키면 늘어난 만큼 행 번호를 밀어줌
    getChildElementsByLocalName(sheetData, 'row').forEach((rowEl) => {
      const rowNum = parseInt(rowEl.getAttribute('r'), 10);
      if (!Number.isFinite(rowNum) || rowNum >= meterStartRow) return;
      getChildElementsByLocalName(rowEl, 'c').forEach((cellEl) => {
        const fNode = getChildElementsByLocalName(cellEl, 'f')[0];
        if (!fNode) return;
        const formulaText = String(fNode.textContent || '').trim();
        if (!formulaText) return;
        fNode.textContent = shiftFormulaRowRefsFromRow(formulaText, meterTemplateEndRow + 1, extraRows);
      });
    });

    const dimensionNode = worksheet.getElementsByTagNameNS('*', 'dimension')[0];
    if (dimensionNode) {
      const currentRef = dimensionNode.getAttribute('ref');
      if (currentRef && currentRef.includes(':')) {
        const [startRef, endRef] = currentRef.split(':');
        const nextEndRef = shiftCellRefRowsFrom(endRef, meterTemplateEndRow + 1, extraRows);
        dimensionNode.setAttribute('ref', `${startRef}:${nextEndRef}`);
      }
    }
  };

  // 측정행이 양식보다 적으면 남는 양식 행을 지우고 아래 행(평균 등)을 끌어올림
  const shrinkRecordSheetMeterRows = (worksheet, usedRowCount, layout = {}) => {
    const sheetData = worksheet.getElementsByTagNameNS('*', 'sheetData')[0];
    if (!sheetData) return;
    const meterStartRow = Number.isFinite(layout.meterStartRow) ? layout.meterStartRow : 17;
    const meterTemplateEndRow = Number.isFinite(layout.meterTemplateEndRow) ? layout.meterTemplateEndRow : 26;
    const templateCapacity = meterTemplateEndRow - meterStartRow + 1;
    if (!Number.isFinite(usedRowCount) || usedRowCount < 1 || usedRowCount >= templateCapacity) return;
    const removeCount = templateCapacity - usedRowCount;
    const firstRemove = meterStartRow + usedRowCount;

    getChildElementsByLocalName(sheetData, 'row').forEach((rowEl) => {
      const rowNum = parseInt(rowEl.getAttribute('r'), 10);
      if (Number.isFinite(rowNum) && rowNum >= firstRemove && rowNum <= meterTemplateEndRow) sheetData.removeChild(rowEl);
    });
    const mergeCellsNode = worksheet.getElementsByTagNameNS('*', 'mergeCells')[0];
    if (mergeCellsNode) {
      getChildElementsByLocalName(mergeCellsNode, 'mergeCell').forEach((mergeCell) => {
        const m = /^[A-Z]+(\d+):[A-Z]+(\d+)$/.exec(mergeCell.getAttribute('ref') || '');
        if (!m) return;
        const r1 = parseInt(m[1], 10);
        const r2 = parseInt(m[2], 10);
        if (r1 >= firstRemove && r2 <= meterTemplateEndRow) mergeCellsNode.removeChild(mergeCell);
      });
      mergeCellsNode.setAttribute('count', String(getChildElementsByLocalName(mergeCellsNode, 'mergeCell').length));
    }
    shiftSheetRowsDown(sheetData, meterTemplateEndRow + 1, -removeCount);
    shiftMergedCellsRowsDown(worksheet, meterTemplateEndRow + 1, -removeCount);

    const meterEndRow = meterTemplateEndRow - removeCount;
    const summaryRow = findExistingRow(sheetData, meterEndRow + 1);
    if (summaryRow) {
      getChildElementsByLocalName(summaryRow, 'c').forEach((cellEl) => {
        const fNode = getChildElementsByLocalName(cellEl, 'f')[0];
        if (!fNode) return;
        const formulaText = String(fNode.textContent || '').trim();
        if (!formulaText) return;
        fNode.textContent = formulaText.replace(
          new RegExp(`(\\$?[A-Z]+\\$?(?:${meterStartRow}|${meterStartRow + 1}):\\$?[A-Z]+\\$?)(\\d+)`, 'g'),
          (full, head, tail) => (parseInt(tail, 10) > meterEndRow ? `${head}${meterEndRow}` : full)
        ).replace(
          new RegExp(`(^|[^A-Z0-9_:$])([A-Z]{1,3})${meterTemplateEndRow}(?![0-9:])`, 'g'),
          (full, prefix, col) => `${prefix}${col}${meterEndRow}`
        );
      });
    }
    getChildElementsByLocalName(sheetData, 'row').forEach((rowEl) => {
      const rowNum = parseInt(rowEl.getAttribute('r'), 10);
      if (!Number.isFinite(rowNum) || rowNum >= meterStartRow) return;
      getChildElementsByLocalName(rowEl, 'c').forEach((cellEl) => {
        const fNode = getChildElementsByLocalName(cellEl, 'f')[0];
        if (!fNode) return;
        const formulaText = String(fNode.textContent || '').trim();
        if (!formulaText) return;
        fNode.textContent = shiftFormulaRowRefsFromRow(formulaText, meterTemplateEndRow + 1, -removeCount);
      });
    });
    const dimensionNode = worksheet.getElementsByTagNameNS('*', 'dimension')[0];
    if (dimensionNode) {
      const currentRef = dimensionNode.getAttribute('ref');
      if (currentRef && currentRef.includes(':')) {
        const [startRef, endRef] = currentRef.split(':');
        dimensionNode.setAttribute('ref', `${startRef}:${shiftCellRefRowsFrom(endRef, meterTemplateEndRow + 1, -removeCount)}`);
      }
    }
  };

  // 측정점별 행 배경색: styles.xml에 채우기 + 셀 서식을 추가해 새 서식 번호를 돌려줌
  const METER_ROW_POINT_COLORS = ['FFF2CC', 'DDEBF7', 'E2EFDA', 'FCE4D6', 'EDE7F6', 'FFE8F0'];
  const createMeterRowStyler = (stylesXml) => {
    const parser = new DOMParser();
    const doc = parser.parseFromString(stylesXml, 'application/xml');
    const root = doc.documentElement;
    const ns = root.namespaceURI;
    const fillsNode = root.getElementsByTagNameNS('*', 'fills')[0];
    const cellXfsNode = root.getElementsByTagNameNS('*', 'cellXfs')[0];
    if (!fillsNode || !cellXfsNode) return null;
    const fillIdByColor = new Map();
    const xfCache = new Map();
    const getFillId = (color) => {
      if (fillIdByColor.has(color)) return fillIdByColor.get(color);
      const fill = doc.createElementNS(ns, 'fill');
      const pattern = doc.createElementNS(ns, 'patternFill');
      pattern.setAttribute('patternType', 'solid');
      const fg = doc.createElementNS(ns, 'fgColor');
      fg.setAttribute('rgb', `FF${color}`);
      const bg = doc.createElementNS(ns, 'bgColor');
      bg.setAttribute('indexed', '64');
      pattern.appendChild(fg);
      pattern.appendChild(bg);
      fill.appendChild(pattern);
      fillsNode.appendChild(fill);
      const id = getChildElementsByLocalName(fillsNode, 'fill').length - 1;
      fillsNode.setAttribute('count', String(id + 1));
      fillIdByColor.set(color, id);
      return id;
    };
    const styleFor = (originalIndex, colorIdx) => {
      const color = METER_ROW_POINT_COLORS[colorIdx % METER_ROW_POINT_COLORS.length];
      const key = `${originalIndex}|${color}`;
      if (xfCache.has(key)) return xfCache.get(key);
      const xfs = getChildElementsByLocalName(cellXfsNode, 'xf');
      const source = xfs[originalIndex] || xfs[0];
      if (!source) return originalIndex;
      const clone = source.cloneNode(true);
      clone.setAttribute('fillId', String(getFillId(color)));
      clone.setAttribute('applyFill', '1');
      cellXfsNode.appendChild(clone);
      const newIndex = getChildElementsByLocalName(cellXfsNode, 'xf').length - 1;
      cellXfsNode.setAttribute('count', String(newIndex + 1));
      xfCache.set(key, newIndex);
      return newIndex;
    };
    // 같은 서식에서 가로 정렬만 가운데로 바꾼 서식 번호 (글꼴·테두리·바탕·표시 형식은 그대로)
    const centeredStyle = (originalIndex) => {
      const key = `center|${originalIndex}`;
      if (xfCache.has(key)) return xfCache.get(key);
      const xfs = getChildElementsByLocalName(cellXfsNode, 'xf');
      const source = xfs[originalIndex] || xfs[0];
      if (!source) return originalIndex;
      const clone = source.cloneNode(true);
      let alignment = getChildElementsByLocalName(clone, 'alignment')[0];
      if (!alignment) {
        alignment = doc.createElementNS(ns, 'alignment');
        alignment.setAttribute('vertical', 'center');
        clone.insertBefore(alignment, clone.firstChild);
      }
      alignment.setAttribute('horizontal', 'center');
      alignment.removeAttribute('indent');
      clone.setAttribute('applyAlignment', '1');
      cellXfsNode.appendChild(clone);
      const newIndex = getChildElementsByLocalName(cellXfsNode, 'xf').length - 1;
      cellXfsNode.setAttribute('count', String(newIndex + 1));
      xfCache.set(key, newIndex);
      return newIndex;
    };
    // 같은 서식에서 세로 정렬만 아래로 바꾼 서식 번호 (줄을 높여 제목 위에 여백을 만들 때)
    const bottomStyle = (originalIndex) => {
      const key = `bottom|${originalIndex}`;
      if (xfCache.has(key)) return xfCache.get(key);
      const xfs = getChildElementsByLocalName(cellXfsNode, 'xf');
      const source = xfs[originalIndex] || xfs[0];
      if (!source) return originalIndex;
      const clone = source.cloneNode(true);
      let alignment = getChildElementsByLocalName(clone, 'alignment')[0];
      if (!alignment) { alignment = doc.createElementNS(ns, 'alignment'); clone.insertBefore(alignment, clone.firstChild); }
      alignment.setAttribute('vertical', 'bottom');
      clone.setAttribute('applyAlignment', '1');
      cellXfsNode.appendChild(clone);
      const newIndex = getChildElementsByLocalName(cellXfsNode, 'xf').length - 1;
      cellXfsNode.setAttribute('count', String(newIndex + 1));
      xfCache.set(key, newIndex);
      return newIndex;
    };
    // 같은 서식에서 표시 형식만 바꾼 서식 번호 (예: 동압을 소수 첫째 자리로)
    const numFmtsNode = (() => {
      let node = root.getElementsByTagNameNS('*', 'numFmts')[0];
      if (!node) { node = doc.createElementNS(ns, 'numFmts'); node.setAttribute('count', '0'); root.insertBefore(node, root.firstChild); }
      return node;
    })();
    const numFmtIdFor = (code) => {
      const existing = getChildElementsByLocalName(numFmtsNode, 'numFmt').find((n) => n.getAttribute('formatCode') === code);
      if (existing) return existing.getAttribute('numFmtId');
      const used = getChildElementsByLocalName(numFmtsNode, 'numFmt').map((n) => parseInt(n.getAttribute('numFmtId'), 10));
      const id = String(Math.max(300, ...used) + 1);
      const node = doc.createElementNS(ns, 'numFmt');
      node.setAttribute('numFmtId', id); node.setAttribute('formatCode', code);
      numFmtsNode.appendChild(node);
      numFmtsNode.setAttribute('count', String(getChildElementsByLocalName(numFmtsNode, 'numFmt').length));
      return id;
    };
    const formatStyle = (originalIndex, code) => {
      const key = `fmt|${code}|${originalIndex}`;
      if (xfCache.has(key)) return xfCache.get(key);
      const xfs = getChildElementsByLocalName(cellXfsNode, 'xf');
      const source = xfs[originalIndex] || xfs[0];
      if (!source) return originalIndex;
      const clone = source.cloneNode(true);
      clone.setAttribute('numFmtId', numFmtIdFor(code));
      clone.setAttribute('applyNumberFormat', '1');
      cellXfsNode.appendChild(clone);
      const newIndex = getChildElementsByLocalName(cellXfsNode, 'xf').length - 1;
      cellXfsNode.setAttribute('count', String(newIndex + 1));
      xfCache.set(key, newIndex);
      return newIndex;
    };
    return { styleFor, centeredStyle, bottomStyle, formatStyle, serialize: () => new XMLSerializer().serializeToString(doc) };
  };

  const applyMeterRowColors = (worksheet, rowColorList, styler) => {
    if (!styler || !Array.isArray(rowColorList)) return;
    const sheetData = worksheet.getElementsByTagNameNS('*', 'sheetData')[0];
    if (!sheetData) return;
    const ns = worksheet.namespaceURI || 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
    rowColorList.forEach(({ row, colorIdx }) => {
      const rowEl = ensureRow(sheetData, row, ns);
      // 양식에서 원래 색이 칠해진 칸만 (F·J·L~O 수식·여백 칸은 제외)
      ['A', 'B', 'C', 'D', 'E', 'G', 'H', 'I', 'K', 'P'].forEach((col) => {
        const cell = ensureCell(rowEl, parseCellRef(`${col}${row}`), ns);
        const original = parseInt(cell.getAttribute('s') || '0', 10);
        cell.setAttribute('s', String(styler.styleFor(Number.isFinite(original) ? original : 0, colorIdx)));
      });
    });
  };

  const removeCalcChainFromZip = async (zip) => {
    if (!zip.file('xl/calcChain.xml')) return;
    zip.remove('xl/calcChain.xml');
    const relsFile = zip.file('xl/_rels/workbook.xml.rels');
    if (relsFile) {
      const relsXmlNow = await relsFile.async('string');
      zip.file('xl/_rels/workbook.xml.rels', relsXmlNow.replace(/<Relationship[^>]*calcChain\.xml"[^>]*\/>/g, ''));
    }
    const ctFile = zip.file('[Content_Types].xml');
    if (ctFile) {
      const ctXml = await ctFile.async('string');
      zip.file('[Content_Types].xml', ctXml.replace(/<Override[^>]*calcChain\.xml"[^>]*\/>/g, ''));
    }
  };

  const setXmlCellValue = (cell, value, ns) => {
    getChildElementsByLocalName(cell, 'f').forEach(node => cell.removeChild(node));
    getChildElementsByLocalName(cell, 'v').forEach(node => cell.removeChild(node));
    getChildElementsByLocalName(cell, 'is').forEach(node => cell.removeChild(node));

    if (value === null || value === undefined || value === '') {
      cell.removeAttribute('t');
      return;
    }

    if (typeof value === 'number' && Number.isFinite(value)) {
      cell.removeAttribute('t');
      const v = cell.ownerDocument.createElementNS(ns, 'v');
      v.textContent = String(value);
      cell.appendChild(v);
      return;
    }

    if (typeof value === 'object' && typeof value.formula === 'string') {
      cell.removeAttribute('t');
      const f = cell.ownerDocument.createElementNS(ns, 'f');
      f.textContent = value.formula;
      cell.appendChild(f);
      return;
    }

    // XML에 쓸 수 없는 제어문자(붙여넣기로 섞여 들어온 것)는 빼고, 엑셀 셀 한도(32,767자)로 자른다
    // eslint-disable-next-line no-control-regex
    const text = String(value).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F\uFFFE\uFFFF]/g, '').slice(0, 32767);
    cell.setAttribute('t', 'inlineStr');

    const isNode = cell.ownerDocument.createElementNS(ns, 'is');
    const tNode = cell.ownerDocument.createElementNS(ns, 't');
    if (/^\s|\s$/.test(text) || /\s{2,}/.test(text)) {
      tNode.setAttribute('xml:space', 'preserve');
    }
    tNode.textContent = text;
    isNode.appendChild(tNode);
    cell.appendChild(isNode);
  };

  const patchWorksheetXml = (sheetXml, updates, options = {}) => {
    const parser = new DOMParser();
    const doc = parser.parseFromString(sheetXml, 'application/xml');
    if (doc.getElementsByTagName('parsererror').length > 0) {
      throw new Error('시트 XML 파싱에 실패했습니다.');
    }

    const worksheet = doc.documentElement;
    const ns = worksheet.namespaceURI || 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
    const sheetData = worksheet.getElementsByTagNameNS('*', 'sheetData')[0];
    if (!sheetData) {
      throw new Error('시트 데이터(sheetData)를 찾지 못했습니다.');
    }

    if (Number.isFinite(options.expandMeterRowsTo)) {
      expandRecordSheetMeterRows(worksheet, options.expandMeterRowsTo, options.layout || {});
    }
    if (Number.isFinite(options.shrinkMeterRowsTo)) {
      shrinkRecordSheetMeterRows(worksheet, options.shrinkMeterRowsTo, options.layout || {});
    }
    if (Array.isArray(options.removeMerges) && options.removeMerges.length > 0) {
      const mergeCellsNode = worksheet.getElementsByTagNameNS('*', 'mergeCells')[0];
      if (mergeCellsNode) {
        getChildElementsByLocalName(mergeCellsNode, 'mergeCell').forEach((mergeCell) => {
          if (options.removeMerges.includes(mergeCell.getAttribute('ref'))) mergeCellsNode.removeChild(mergeCell);
        });
        mergeCellsNode.setAttribute('count', String(getChildElementsByLocalName(mergeCellsNode, 'mergeCell').length));
      }
    }
    // 새 병합 범위 추가 (예: 선택번호 칸을 세 줄로)
    if (Array.isArray(options.addMerges) && options.addMerges.length > 0) {
      let mergeCellsNode = worksheet.getElementsByTagNameNS('*', 'mergeCells')[0];
      if (!mergeCellsNode) {
        mergeCellsNode = doc.createElementNS(ns, 'mergeCells');
        if (sheetData.nextSibling) worksheet.insertBefore(mergeCellsNode, sheetData.nextSibling);
        else worksheet.appendChild(mergeCellsNode);
      }
      const existing = getChildElementsByLocalName(mergeCellsNode, 'mergeCell').map((node) => node.getAttribute('ref'));
      options.addMerges.forEach((ref) => {
        if (existing.includes(ref)) return;
        const node = doc.createElementNS(ns, 'mergeCell');
        node.setAttribute('ref', ref);
        mergeCellsNode.appendChild(node);
      });
      mergeCellsNode.setAttribute('count', String(getChildElementsByLocalName(mergeCellsNode, 'mergeCell').length));
    }

    if (options.shiftRefsToSheet && Number.isFinite(options.shiftRefsFromRow) && options.shiftRefsDelta) {
      shiftCrossSheetRowRefs(sheetData, options.shiftRefsToSheet, options.shiftRefsFromRow, options.shiftRefsDelta, options.extendRangeEndFrom);
    }

    updates.forEach(({ ref, value, styleFrom }) => {
      const parsed = parseCellRef(ref);
      if (!parsed) return;
      const row = ensureRow(sheetData, parsed.row, ns);
      const cell = ensureCell(row, parsed, ns);
      setXmlCellValue(cell, value, ns);
      // 새로 만든 칸에 다른 칸의 서식(테두리·글꼴)을 입힌다
      if (styleFrom) {
        const src = parseCellRef(styleFrom);
        const srcRow = src ? findExistingRow(sheetData, src.row) : null;
        const srcCell = srcRow ? findExistingCell(srcRow, src.ref) : null;
        const style = srcCell ? srcCell.getAttribute('s') : null;
        if (style !== null) cell.setAttribute('s', style);
        if (srcRow && srcRow.getAttribute('ht') && !row.getAttribute('ht')) { row.setAttribute('ht', srcRow.getAttribute('ht')); row.setAttribute('customHeight', '1'); }
      }
    });

    if (options.meterRowColors && options.rowStyler) {
      applyMeterRowColors(worksheet, options.meterRowColors, options.rowStyler);
    }
    // 줄 높이를 바꾸고(제목 위 여백), 그 줄의 칸은 아래로 붙인다: { 8: 34 }
    if (options.rowGaps && options.rowStyler?.bottomStyle) {
      Object.entries(options.rowGaps).forEach(([rowText, height]) => {
        const rowEl = findExistingRow(sheetData, parseInt(rowText, 10));
        if (!rowEl) return;
        rowEl.setAttribute('ht', String(height));
        rowEl.setAttribute('customHeight', '1');
        getChildElementsByLocalName(rowEl, 'c').forEach((cell) => {
          const original = parseInt(cell.getAttribute('s') || '0', 10);
          cell.setAttribute('s', String(options.rowStyler.bottomStyle(Number.isFinite(original) ? original : 0)));
        });
      });
    }
    // 표시 형식을 바꿀 칸들: [{ range: 'E17:E30', code: '0.0' }]
    if (Array.isArray(options.numberFormats) && options.rowStyler?.formatStyle) {
      options.numberFormats.forEach(({ range, code }) => {
        const [fromRef, toRef] = String(range).split(':');
        const from = parseCellRef(fromRef); const to = parseCellRef(toRef || fromRef);
        if (!from || !to) return;
        for (let r = from.row; r <= to.row; r++) {
          const rowEl = findExistingRow(sheetData, r);
          if (!rowEl) continue;
          for (let c = from.col; c <= to.col; c++) {
            const cell = getChildElementsByLocalName(rowEl, 'c').find((el) => { const p = parseCellRef(el.getAttribute('r')); return p && p.col === c; });
            if (!cell) continue;
            const original = parseInt(cell.getAttribute('s') || '0', 10);
            cell.setAttribute('s', String(options.rowStyler.formatStyle(Number.isFinite(original) ? original : 0, code)));
          }
        }
      });
    }
    // 가운데 정렬로 맞출 칸들 (예: 'A5:B7', 'C19'): 양식에서 주변 칸과 정렬이 어긋나 있던 곳
    if (Array.isArray(options.centerCells) && options.centerCells.length > 0 && options.rowStyler?.centeredStyle) {
      options.centerCells.forEach((range) => {
        const [fromRef, toRef] = String(range).split(':');
        const from = parseCellRef(fromRef);
        const to = parseCellRef(toRef || fromRef);
        if (!from || !to) return;
        for (let r = Math.min(from.row, to.row); r <= Math.max(from.row, to.row); r++) {
          const rowEl = findExistingRow(sheetData, r);
          if (!rowEl) continue;
          for (let c = Math.min(from.col, to.col); c <= Math.max(from.col, to.col); c++) {
            const cell = getChildElementsByLocalName(rowEl, 'c').find((el) => { const p = parseCellRef(el.getAttribute('r')); return p && p.col === c; });
            if (!cell) continue;
            const original = parseInt(cell.getAttribute('s') || '0', 10);
            cell.setAttribute('s', String(options.rowStyler.centeredStyle(Number.isFinite(original) ? original : 0)));
          }
        }
      });
    }

    return new XMLSerializer().serializeToString(doc);
  };

  // 행/셀이 XML에 없으면 (빈 칸) 순서에 맞게 만들어 줌
  const ensureRow = (sheetData, rowNum, ns) => {
    const existing = findExistingRow(sheetData, rowNum);
    if (existing) return existing;
    const newRow = sheetData.ownerDocument.createElementNS(ns, 'row');
    newRow.setAttribute('r', String(rowNum));
    const nextRow = getChildElementsByLocalName(sheetData, 'row').find((rowEl) => {
      const n = parseInt(rowEl.getAttribute('r'), 10);
      return Number.isFinite(n) && n > rowNum;
    });
    if (nextRow) sheetData.insertBefore(newRow, nextRow);
    else sheetData.appendChild(newRow);
    return newRow;
  };

  const ensureCell = (rowEl, parsedRef, ns) => {
    const existing = findExistingCell(rowEl, parsedRef.ref);
    if (existing) return existing;
    const newCell = rowEl.ownerDocument.createElementNS(ns, 'c');
    newCell.setAttribute('r', parsedRef.ref);
    const nextCell = getChildElementsByLocalName(rowEl, 'c').find((cellEl) => {
      const p = parseCellRef(cellEl.getAttribute('r'));
      return p && p.col > parsedRef.col;
    });
    if (nextCell) rowEl.insertBefore(newCell, nextCell);
    else rowEl.appendChild(newCell);
    return newCell;
  };

  // 다른 시트가 기록지 행을 참조하는 수식: 늘어난 행만큼 밀고, 측정행 범위 끝은 확장
  const shiftCrossSheetRowRefs = (sheetData, targetSheetName, fromRow, delta, extendRangeEndFrom) => {
    const escaped = targetSheetName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    const refPattern = new RegExp(`('${escaped}'|${escaped})!(\\$?[A-Z]{1,3}\\$?)(\\d+)(?::(\\$?[A-Z]{1,3}\\$?)(\\d+))?`, 'g');
    getChildElementsByLocalName(sheetData, 'row').forEach((rowEl) => {
      getChildElementsByLocalName(rowEl, 'c').forEach((cellEl) => {
        const fNode = getChildElementsByLocalName(cellEl, 'f')[0];
        if (!fNode) return;
        const formulaText = String(fNode.textContent || '');
        if (!formulaText.includes(targetSheetName)) return;
        fNode.textContent = formulaText.replace(refPattern, (full, sheetPart, col1, row1, col2, row2) => {
          const r1 = parseInt(row1, 10);
          const shift = (r) => (r >= fromRow ? r + delta : r);
          if (row2 === undefined) {
            const single = (Number.isFinite(extendRangeEndFrom) && r1 === extendRangeEndFrom) ? r1 + delta : shift(r1);
            return `${sheetPart}!${col1}${single}`;
          }
          const r2 = parseInt(row2, 10);
          const nextEnd = (Number.isFinite(extendRangeEndFrom) && r2 === extendRangeEndFrom) ? r2 + delta : shift(r2);
          return `${sheetPart}!${col1}${shift(r1)}:${col2}${nextEnd}`;
        });
      });
    });
  };

  const normalizeWorksheetPath = (targetPath) => {
    const cleaned = String(targetPath || '').trim();
    if (!cleaned) return null;
    if (cleaned.startsWith('/')) return cleaned.replace(/^\/+/, '');
    if (cleaned.startsWith('xl/')) return cleaned;
    return `xl/${cleaned.replace(/^\.?\//, '')}`;
  };

  const buildSheetPathMap = (workbookXml, relsXml) => {
    const parser = new DOMParser();
    const workbookDoc = parser.parseFromString(workbookXml, 'application/xml');
    const relsDoc = parser.parseFromString(relsXml, 'application/xml');

    if (workbookDoc.getElementsByTagName('parsererror').length > 0 || relsDoc.getElementsByTagName('parsererror').length > 0) {
      throw new Error('템플릿 메타데이터 파싱에 실패했습니다.');
    }

    const relMap = new Map();
    Array.from(relsDoc.getElementsByTagNameNS('*', 'Relationship')).forEach((rel) => {
      const id = rel.getAttribute('Id');
      const target = rel.getAttribute('Target');
      const normalized = normalizeWorksheetPath(target);
      if (id && normalized) relMap.set(id, normalized);
    });

    const sheetMap = new Map();
    Array.from(workbookDoc.getElementsByTagNameNS('*', 'sheet')).forEach((sheet) => {
      const name = sheet.getAttribute('name');
      const relId = sheet.getAttribute('r:id') || sheet.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id');
      if (!name || !relId) return;
      const path = relMap.get(relId);
      if (path) sheetMap.set(name, path);
    });

    return sheetMap;
  };

  // 내려받기 공통: 폴더(날짜_사업장명_#배출구)와 파일 이름을 알려 주고 저장 창을 띄운다
  const downloadExportFile = (buffer, mime, exportNames) => {
    const url = URL.createObjectURL(new Blob([buffer], { type: mime }));
    const link = document.createElement('a');
    link.href = url;
    try { if (window.stackpilotDesktop?.noteDownloadFolder) window.stackpilotDesktop.noteDownloadFolder(exportNames.folder); } catch (error) { console.error(error); }
    link.setAttribute('download', exportNames.file);
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
  };

  // 공개 양식(회사 양식 파일 없이 프로그램이 직접 그리는 엑셀)으로 기록부 내보내기 — 배포용 빌드에서 사용
  const exportOpenRecordExcel = async (sourceReport = null) => {
    try {
      const report = sourceReport && typeof sourceReport === 'object' ? sourceReport : buildReportSnapshot(selectedSheet || 'dust');
      const sheetId = resolveReportSheetId(report);
      const title = report.sheetTitle || SHEET_MENU.find((item) => item.id === sheetId)?.title || '시료채취기록부';
      const method = MOISTURE_METHODS.find((item) => item.id === (report.moistureMethod || 'auto'));
      const rows = Array.isArray(report.gasMeters) ? report.gasMeters : [];
      // 측정값이 들어간 행까지만 (계획만 채워 둔 빈 행은 뺀다)
      const meterRows = rows.slice(0, Math.max(1, getLastMeasuredMeterIdx(rows) + 1));
      // 흡입관(피토관) 길이: 고른 피토관 프리셋의 유효 길이(cm) → m
      const pitot = pitotPresets.find((item) => String(item.id) === String(report.pitotPresetId));
      const pitotLength = pitot ? parseFloat(pitot.length) : NaN;
      // 굴뚝단면 및 측정점 배열 그림: 굴뚝 안지름으로 그려 표 3의 가운데 칸에 넣는다
      let sectionPng = null;
      try {
        const svg = buildStackSectionSvg(parseFloat(report.stackDiameter), 520, 460);
        if (svg) sectionPng = await renderSvgToPng(svg, 1040, 920);
      } catch (error) {
        console.error('측정점 그림 만들기 실패:', error);
      }
      const sheets = buildRecordSheets(report, {
        title, isFlow: sheetId === 'flow', moistureLabel: method ? `${method.label.replace(/\s*\(.*\)$/, '')} (${method.standard})` : '', meterRows, sectionPng,
        // 사전 수분량을 흡습관법·임핀저법으로 잰 경우의 기록 (기준의 기구 이름으로)
        preMoisture: isImpingerMoistureMethod(method) ? {
          unitTitle: method.unitTitle,
          rows: method.items.map((label, i) => ({ label, initial: report.preImpinger?.impingers?.[i]?.initial, final: report.preImpinger?.impingers?.[i]?.final })),
          meterStart: report.preImpinger?.meterStart, meterEnd: report.preImpinger?.meterEnd, meterTemp: report.preImpinger?.meterTemp,
          pressure: String(report.preImpinger?.pressure ?? '').trim() !== '' ? report.preImpinger.pressure : report.atmPressure,
          orifice: report.preImpinger?.orifice, yd: getPreMoistureYd(method, report), pv: getPreMoisturePv(method, report),
          pressureLabel: method.id === 'impinger' ? '오리피스압 ΔH ÷ 13.6\nPm (mmHg)' : '가스미터압 Pm (mmHg)',
        } : null,
        probeLengthM: Number.isFinite(pitotLength) && pitotLength > 0 ? Math.round(pitotLength) / 100 : null,
      });
      const buffer = await buildWorkbook(sheets);
      downloadExportFile(buffer, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        buildExportNames({ date: report.date, title, company: report.company, location: report.location, ext: 'xlsx' }));
    } catch (error) {
      console.error(error);
      alert('엑셀 저장 중 오류가 발생했습니다. 입력값을 확인해주세요.');
    }
  };

  const exportToTemplateExcel = async (sourceReport = null) => {
    if (IS_DESKTOP_BUILD) { await exportOpenRecordExcel(sourceReport); return; }
    try {
      const sourceData = sourceReport && typeof sourceReport === 'object' ? sourceReport : formData;
      const sourcePoints = Array.isArray(sourceData.points) ? sourceData.points : [];
      const sourceGasAnalyzer = Array.isArray(sourceData.gasAnalyzer) ? sourceData.gasAnalyzer : [];
      const sourceGasMeters = Array.isArray(sourceData.gasMeters) ? sourceData.gasMeters : [];
      const sourceMoistures = Array.isArray(sourceData.moistureValues) ? sourceData.moistureValues : [];
      const sourceImpingers = Array.isArray(sourceData.impingers) ? sourceData.impingers : [];
      const avg = (vals) => {
        const valid = vals.filter((v) => Number.isFinite(v));
        if (valid.length === 0) return NaN;
        return valid.reduce((acc, cur) => acc + cur, 0) / valid.length;
      };
      const toNumStrict = (v) => {
        const n = parseFloat(v);
        return Number.isFinite(n) ? n : NaN;
      };
      const avgSp = avg(sourcePoints.map((p) => toNumStrict(p?.sp)));
      const avgTs = avg(sourcePoints.map((p) => toNumStrict(p?.ts)));
      const avgTm = avg(sourceGasMeters
        .filter((_, idx) => idx !== 0)
        .map((m) => {
          const inT = toNumStrict(m?.tmIn);
          const outT = toNumStrict(m?.tmOut);
          if (Number.isFinite(inT) && Number.isFinite(outT)) return (inT + outT) / 2;
          if (Number.isFinite(inT)) return inT;
          if (Number.isFinite(outT)) return outT;
          return NaN;
        }));
      const avgOrifice = avg(sourceGasMeters
        .filter((_, idx) => idx !== 0)
        .map((m) => toNumStrict(m?.pressure)));

      const hasO2 = sourceGasAnalyzer.some((g) => Number.isFinite(parseFloat(g?.o2)));
      if (!hasO2) {
        alert('연소가스(가스분석기) O2 값이 없습니다. 양식의 노즐산정·기록지 계산은 O2/CO2 값이 있어야 하며, 없으면 Excel에서 #VALUE!/#DIV/0!로 표시됩니다. 추출은 계속 진행합니다.');
      }
      const exportSheetId = (sourceReport && typeof sourceReport === 'object')
        ? resolveReportSheetId(sourceReport)
        : (activeSheet?.id || 'dust');
      const layout = EXPORT_TEMPLATE_LAYOUTS[exportSheetId] || EXPORT_TEMPLATE_LAYOUTS.dust;
      const response = await fetch(getTemplateUrlForSheet(exportSheetId));
      if (!response.ok) throw new Error('템플릿 파일을 불러오지 못했습니다.');

      const templateBuffer = await response.arrayBuffer();
      const zip = await JSZip.loadAsync(templateBuffer);

      const workbookFile = zip.file('xl/workbook.xml');
      const relsFile = zip.file('xl/_rels/workbook.xml.rels');
      if (!workbookFile || !relsFile) {
        throw new Error('템플릿 workbook 구조를 찾지 못했습니다.');
      }

      const workbookXml = await workbookFile.async('string');
      const workbookRelsXml = await relsFile.async('string');
      const sheetPathMap = buildSheetPathMap(workbookXml, workbookRelsXml);

      const requiredSheets = ['표지', '연소가스자료', '노즐산정', '수분량', '기록지(수분량자동측정기)'];
      requiredSheets.forEach((sheetName) => {
        if (!sheetPathMap.has(sheetName)) {
          throw new Error(`템플릿 시트 누락: ${sheetName}`);
        }
      });

      const updatesBySheet = new Map();
      const addUpdate = (sheetName, ref, value, styleFrom = null) => {
        if (!updatesBySheet.has(sheetName)) updatesBySheet.set(sheetName, []);
        updatesBySheet.get(sheetName).push({ ref, value, styleFrom });
      };
      // 측정값이 들어간 행까지만 내보낸다 (조기 종료 시 계획만 채워진 빈 행이 엑셀 시간 MAX·평균에 섞이지 않게)
      const lastMeasuredIdx = getLastMeasuredMeterIdx(sourceGasMeters);
      const lastMeaningfulIdx = lastMeasuredIdx > 0
        ? lastMeasuredIdx
        : sourceGasMeters.reduce((lastIdx, meter, idx) => (
          String(meter?.pointNum || '').trim() !== '' || String(meter?.time || '').trim() !== '' ? idx : lastIdx
        ), 0);
      const gasMeters = sourceGasMeters.slice(0, Math.max(1, lastMeaningfulIdx + 1));
      const meterTemplateCapacity = layout.meterTemplateEndRow - layout.meterStartRow + 1;
      // 측정행 수는 입력한 행 수 그대로: 양식보다 많으면 행 추가, 적으면 남는 양식 행 삭제
      // 시작행만 남기면 평균 수식이 자기 자신을 참조하므로 측정행을 최소 1줄 둔다
      const meterRowsForTemplate = Math.max(2, gasMeters.length);
      const meterExtraRows = meterRowsForTemplate - meterTemplateCapacity; // 음수면 삭제된 행 수
      const firstVolume = gasMeters.find(m => m.volume !== '' && m.volume !== undefined)?.volume ?? '';
      const lastVolume = [...gasMeters].reverse().find(m => m.volume !== '' && m.volume !== undefined)?.volume ?? '';

      const recordSheet = '기록지(수분량자동측정기)';
      const filterNumber = String(sourceData.filterId || '').trim() || `${layout.filterPrefix}-1`;
      const pitotSerial = String(sourceData.pitotInfo || '').trim();

      // 표지
      addUpdate('표지', 'D19', sourceData.company);
      addUpdate('표지', 'D20', sourceData.location);
      addUpdate('표지', 'D21', toExcelDateSerial(sourceData.date));
      // 누출시험 시트: 구분별 O/X, 가스미터 전·후(또는 비고)
      if (layout.leak) {
        const leak = sanitizeLeakTest(sourceData.leakTest);
        LEAK_TEST_PARTS.forEach((part, idx) => addUpdate('누출시험', `C${layout.leak.firstRow + idx}`, leak[part.key]));
        if (layout.leak.hasMeter) {
          addUpdate('누출시험', `D${layout.leak.firstRow}`, toNumber(leak.meterBefore));
          addUpdate('누출시험', `E${layout.leak.firstRow}`, toNumber(leak.meterAfter));
        } else if (layout.leak.noteCol && leak.note) {
          addUpdate('누출시험', `${layout.leak.noteCol}${layout.leak.firstRow}`, leak.note);
        }
      }

      // 연소가스자료 (가스분석기 3회)
      for (let i = 0; i < 3; i++) {
        const gas = sourceGasAnalyzer[i] || {};
        const row = 8 + i;
        addUpdate('연소가스자료', `A${row}`, toExcelTimeFraction(gas.time));
        addUpdate('연소가스자료', `B${row}`, toNumber(gas.o2));
        addUpdate('연소가스자료', `C${row}`, toNumber(gas.co2));
        addUpdate('연소가스자료', `D${row}`, toNumber(gas.co));
        addUpdate('연소가스자료', `E${row}`, toNumber(gas.nox));
        addUpdate('연소가스자료', `F${row}`, toNumber(gas.sox));
      }

      // 노즐산정: 대기 조건, 피토관, 샘플러 계수
      addUpdate('노즐산정', 'B5', toNumber(sourceData.samplerId));
      addUpdate('노즐산정', 'A8', toNumber(sourceData.atmTemp));
      addUpdate('노즐산정', 'B8', toNumber(sourceData.atmPressure));
      // 피토관을 고르지 않았으면 양식에 들어 있던 예시 피토관 번호도 지운다
      addUpdate('노즐산정', 'C8', pitotSerial || null);
      addUpdate('노즐산정', 'C10', toNumber(sourceData.pitotFactor));
      addUpdate('노즐산정', 'F8', toNumber(sourceData.deltaHAt));
      addUpdate('노즐산정', 'G8', toNumber(sourceData.gasMeterFactor));
      addUpdate('노즐산정', 'D15', toNumber(sourceData.traverseTmIn));
      addUpdate('노즐산정', 'F15', toNumber(sourceData.traverseTmOut));
      addUpdate('노즐산정', 'G15', Number.isFinite(avgTs) ? avgTs : null);
      for (let i = 0; i < 5; i++) {
        const point = sourcePoints[i] || {};
        const row = 17 + i;
        addUpdate('노즐산정', `G${row}`, toNumber(point.ts));
        addUpdate('노즐산정', `H${row}`, toNumber(point.tp));
        addUpdate('노즐산정', `I${row}`, toNumber(point.sp));
        addUpdate('노즐산정', `J${row}`, toNumber(point.dp));
      }
      addUpdate('노즐산정', 'A39', toNumber(sourceData.planSamplingTime));
      addUpdate('노즐산정', 'K39', toNumber(sourceData.usedNozzleNum) || toNumber(sourceData.recommendedNozzleNum));
      const exportNozzles = (sanitizeNozzleSet(getNozzleSetForSheet(exportSheetId)) || []).slice().sort((a, b) => Number(a.num) - Number(b.num));
      for (let i = 0; i < layout.nozzleRows; i++) {
        const row = 39 + i;
        // 수은·PAHs 양식에는 그 기록부의 노즐 세트를 쓴다
        const nozzle = exportNozzles[i];
        addUpdate('노즐산정', `B${row}`, nozzle ? toNumber(nozzle.num) : null);
        addUpdate('노즐산정', `C${row}`, nozzle ? toNumber(nozzle.d) : null);
      }

      // 시료측정위치: 전체 길이(m), 플랜지 길이(cm)
      addUpdate('시료측정위치', 'B34', toNumber(sourceData.totalStackDepth));
      const flangeCm = toNumber(sourceData.flangeLength);
      addUpdate('시료측정위치', 'B37', flangeCm === null ? null : flangeCm * 100);
      // 굴뚝 내경: 유량 양식은 A37이 고정값(1.1)이라 반드시 써 넣고, 직접 고친 내경도 그대로 반영
      const stackDiameterM = toNumber(sourceData.stackDiameter);
      if (stackDiameterM !== null && stackDiameterM > 0) addUpdate('시료측정위치', 'A37', stackDiameterM);

      // 수분량: 사전 수분량은 고른 방법대로, 사후는 임핀저 무게
      //  - 자동측정기: C5=1, B10~B14에 5회값
      //  - 임핀저법: C5=2, 예비 임핀저(C20~D23)·가스미터(A28~G28)
      const exportMoistureMethod = MOISTURE_METHODS.find((m) => m.id === sourceData.moistureMethod) || MOISTURE_METHODS.find((m) => m.id === 'auto');
      // 측정방법 표를 화면과 같은 1·2·3번으로 적고, 선택번호(C5)도 화면의 번호를 쓴다
      //   1 흡습관법 / 2 임핀저법 / 3 자동측정법 (기준 ES 01321.1~.3 순서). 양식 수식(노즐산정 C15)은 자동측정법일 때 자동측정기 평균, 그 밖엔 무게법 결과(H28)를 보게 바꾼다
      MOISTURE_METHODS.forEach((item, i) => {
        addUpdate('수분량', `A${5 + i}`, i + 1, 'A6');
        // 칸이 좁아 방법 이름만 적는다 (조항 번호는 아래 표 제목에)
        addUpdate('수분량', `B${5 + i}`, item.label.replace(/\s*\(.*\)$/, ''), 'B6');
      });
      addUpdate('수분량', 'C7', null, 'C5');
      addUpdate('수분량', 'A8', `▶ 자동측정법 (${MOISTURE_METHODS.find((item) => item.id === 'auto').standard})`);
      addUpdate('수분량', 'C5', MOISTURE_METHODS.indexOf(exportMoistureMethod) + 1);
      addUpdate('노즐산정', 'C15', { formula: `IF(수분량!C5=${MOISTURE_METHODS.findIndex((item) => item.id === 'auto') + 1},수분량!B15,수분량!H28)` });
      // 무게법 표의 제목·항목 이름·압력 칸 이름은 선택번호(C5)를 보는 수식으로 적는다
      // → 엑셀에서 번호를 바꿔도 이름이 같이 바뀐다 (1 흡습관법 / 2 임핀저법 / 3 자동측정법)
      {
        const tubeNo = MOISTURE_METHODS.findIndex((m) => m.id === 'imp2') + 1;
        const impNo = MOISTURE_METHODS.findIndex((m) => m.id === 'impinger') + 1;
        const tube = MOISTURE_METHODS[tubeNo - 1];
        const imp = MOISTURE_METHODS[impNo - 1];
        const q = (text) => `"${String(text).replace(/"/g, '""')}"`;
        const pick = (tubeText, impText, otherText) => ({ formula: `IF($C$5=${tubeNo},${q(tubeText)},IF($C$5=${impNo},${q(impText)},${typeof otherText === 'number' ? otherText : q(otherText)}))` });
        addUpdate('수분량', 'A18', pick(`▶ 흡습관법 (${tube.standard})`, `▶ 임핀저법 (${imp.standard})`, '▶ 흡습관법 · 임핀저법'));
        addUpdate('수분량', 'E19', '흡습 수분량(전후차)');
        for (let i = 0; i < 4; i++) addUpdate('수분량', `A${20 + i}`, pick(tube.items[i] || '-', imp.items[i] || '-', i + 1));
        addUpdate('수분량', 'A24', '흡습 수분의 질량 ( ma )');
        // 임핀저법은 입자 샘플러로 흡입하므로 가스미터 게이지압 = 오리피스압. 흡습관법은 가스 샘플러의 가스미터압
        addUpdate('수분량', 'F26', pick('가스미터압', '오리피스압', '가스미터압'));
        addUpdate('수분량', 'G26', pick('가스미터압', '오리피스압', '가스미터압'));
      }
      // 사후 수분량(본 측정)은 항상 입자 샘플러: 가스미터압 칸은 평균 오리피스압이다
      // 사후 수분량 표의 이름도 위 표와 같은 말로
      for (let i = 0; i < layout.impingerRows; i++) addUpdate('수분량', `A${33 + i}`, `임핀저 ${i + 1}`);
      addUpdate('수분량', 'E32', '흡습 수분량(전후차)');
      // 본 측정 가스흡인량 표의 머리줄: 임핀저가 4개인 양식은 41행, 6개인 PAHs 양식은 43행
      const postMeterHeadRow = 41 + (layout.impingerRows - 4);
      addUpdate('수분량', `F${postMeterHeadRow}`, '오리피스압');
      addUpdate('수분량', `G${postMeterHeadRow}`, '오리피스압');
      if (exportMoistureMethod.id === 'auto') {
        for (let i = 0; i < 5; i++) addUpdate('수분량', `B${10 + i}`, toNumber(sourceMoistures[i]));
      } else {
        // 자동측정기 값이 있으면 참고로 남긴다 (선택은 C5=2라 임핀저 값이 적용됨)
        for (let i = 0; i < 5; i++) addUpdate('수분량', `B${10 + i}`, toNumber(sourceMoistures[i]));
      }
      // 양식의 가스 밀도(노즐산정 F26), 습윤 밀도(노즐산정 A34), 유량 양식의 수분량(계산식 O75)이
      // 자동측정기 평균(B15)이나 빈 칸(H29)만 보던 것을, 선택번호(C5)를 따르는 사전 수분량(노즐산정 C15)을 보게 한다
      // → 엑셀에서 1·2·3을 바꿔도 밀도·유속·노즐 계산이 같이 바뀐다
      addUpdate('노즐산정', 'F26', { formula: 'C15' });
      addUpdate('노즐산정', 'A34', { formula: '(C34*(100-$C$15)/100)+18/(22.4*100)*$C$15' });
      if (exportSheetId === 'flow') addUpdate('계산식', 'O75', { formula: '노즐산정!C15' });
      const preImp = sourceData.preImpinger || createEmptyPreImpinger();
      // 무게법(흡습관·임핀저) 입력이 있으면 지금 고른 방법과 상관없이 적어 둔다 → 엑셀에서 1·2로 바꿔도 값이 있다
      const preImpHasData = [preImp.meterStart, preImp.meterEnd, preImp.meterTemp, preImp.orifice]
        .concat((preImp.impingers || []).flatMap((imp) => [imp?.initial, imp?.final]))
        .some((value) => String(value ?? '').trim() !== '');
      {
        // 흡습관법(1)은 흡습관 2개와 가스 샘플러(보정계수·습식이면 포화수증기압), 임핀저법(2)은 임핀저 4개와 입자 샘플러(노즐산정 G8)
        const tubeMethod = MOISTURE_METHODS.find((m) => m.id === 'imp2');
        const tubeNo = MOISTURE_METHODS.indexOf(tubeMethod) + 1;
        const tubeYd = getPreMoistureYd(tubeMethod, sourceData);
        const tubePv = Math.round(getPreMoisturePv(tubeMethod, sourceData) * 100) / 100;
        addUpdate('수분량', 'E24', { formula: `IF($C$5=${tubeNo},SUM(E20:F21),SUM(E20:F23))` });
        addUpdate('수분량', 'C28', { formula: `IF($C$5=${tubeNo},(B28-A28)*${Number.isFinite(tubeYd) ? tubeYd : 1},(B28-A28)*노즐산정!G8)` });
        // 자료가 없으면 0으로 (#DIV/0!이 노즐산정·계산식 전체로 번지지 않게)
        addUpdate('수분량', 'H28', { formula: `IF(OR(N(C28)=0,N(E24)=0),0,22.4/18*E24/(C28*(273/(273+D28))*((E28+G28-IF($C$5=${tubeNo},${tubePv},0))/760)+22.4/18*E24)*100)` });
      }
      if (preImpHasData) {
        for (let i = 0; i < 4; i++) {
          const imp = (preImp.impingers || [])[i] || {};
          // 아래 사후 수분량 표의 무게 칸과 같은 서식(가운데 정렬)으로 적는다
          addUpdate('수분량', `C${20 + i}`, toNumber(imp.initial), 'C33');
          addUpdate('수분량', `D${20 + i}`, toNumber(imp.final), 'D33');
        }
        const preOrifice = toNumber(preImp.orifice);
        addUpdate('수분량', 'A28', toNumber(preImp.meterStart));
        addUpdate('수분량', 'B28', toNumber(preImp.meterEnd));
        addUpdate('수분량', 'D28', toNumber(preImp.meterTemp));
        addUpdate('수분량', 'E28', toNumber(preImp.pressure) ?? toNumber(sourceData.atmPressure));
        addUpdate('수분량', 'F28', preOrifice);
        addUpdate('수분량', 'G28', preOrifice !== null ? preOrifice / 13.6 : 0);
      } else {
        addUpdate('수분량', 'A28', toNumber(firstVolume));
        addUpdate('수분량', 'B28', toNumber(lastVolume));
        addUpdate('수분량', 'D28', Number.isFinite(avgTm) ? avgTm : null);
        addUpdate('수분량', 'E28', toNumber(sourceData.atmPressure));
        addUpdate('수분량', 'F28', Number.isFinite(avgOrifice) ? avgOrifice : null);
        addUpdate('수분량', 'G28', Number.isFinite(avgOrifice) ? avgOrifice / 13.6 : null);
      }
      for (let i = 0; i < layout.impingerRows; i++) {
        const imp = sourceImpingers[i] || {};
        addUpdate('수분량', `C${33 + i}`, toNumber(imp.initial));
        addUpdate('수분량', `D${33 + i}`, toNumber(imp.final));
        addUpdate('수분량', `E${33 + i}`, { formula: `D${33 + i}-C${33 + i}` });
      }

      // 기록지: 머리글 (계산 셀은 양식 수식을 그대로 두고 입력 셀만 기록)
      addUpdate(recordSheet, 'C5', sourceData.company);
      addUpdate(recordSheet, 'C6', sourceData.location);
      addUpdate(recordSheet, 'C7', String(sourceData.sampler || '').trim() ? `${String(sourceData.sampler).trim()} (인)` : null);
      addUpdate(recordSheet, 'C8', toExcelDateSerial(sourceData.date));
      addUpdate(recordSheet, 'C9', filterNumber);
      addUpdate(recordSheet, 'J13', filterNumber);
      addUpdate(recordSheet, 'B14', sourceData.standardO2 ? toNumber(sourceData.standardO2) : null);
      if (toNumber(sourceData.filterInitial) !== null) addUpdate(recordSheet, 'D14', toNumber(sourceData.filterInitial));
      if (toNumber(sourceData.filterFinal) !== null) addUpdate(recordSheet, 'F14', toNumber(sourceData.filterFinal));

      // 기록지: 측정행 (양식 예시값을 지우고 기록)
      const meterColumns = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'K', 'P'];
      for (let i = 0; i < meterRowsForTemplate; i++) {
        const row = layout.meterStartRow + i;
        meterColumns.forEach((col) => addUpdate(recordSheet, `${col}${row}`, null));
      }
      // 기록지 B열(시료채취 시간)은 누적 분. 앱 입력이 구간(5,5,5…)이면 누적으로 변환, 이미 누적(5,10,15…)이면 그대로
      const meterTimeValues = gasMeters.slice(1).map((m) => parseFloat(m?.time)).filter((v) => Number.isFinite(v) && v > 0);
      // 같은 시간이 연달아 있어도(5,10,10,15) 누적으로 본다
      const isCumulativeMeterTime = meterTimeValues.length >= 2 && meterTimeValues.every((v, i) => i === 0 || v >= meterTimeValues[i - 1]);
      let runningMinutes = 0;
      gasMeters.forEach((meter, idx) => {
        const row = layout.meterStartRow + idx;
        const rawMinutes = toNumber(meter.time);
        let cellMinutes = rawMinutes;
        if (idx > 0 && rawMinutes !== null && !isCumulativeMeterTime) {
          runningMinutes = Math.round((runningMinutes + rawMinutes) * 100) / 100;
          cellMinutes = runningMinutes;
        }
        // 채취점 번호는 숫자로 (글자로 넣으면 엑셀에서 '텍스트 형식의 숫자'로 표시된다)
        addUpdate(recordSheet, `A${row}`, /^\d+$/.test(String(meter.pointNum || '').trim()) ? Number(meter.pointNum) : (meter.pointNum || (idx === 0 ? '시작' : '')));
        addUpdate(recordSheet, `B${row}`, idx === 0 ? 0 : cellMinutes);
        addUpdate(recordSheet, `C${row}`, toNumber(meter.vacuum));
        addUpdate(recordSheet, `D${row}`, toNumber(meter.stackTemp));
        addUpdate(recordSheet, `E${row}`, toNumber(meter.dp));
        // 오리피스 압차: 앱 화면과 같은 값(K계수 × 동압, 없으면 직접 입력값)을 쓴다.
        // 빈칸으로 두면 양식의 F열 수식까지 지워져 등속흡인율이 #DIV/0!가 된다
        const exportK = toNumber(sourceData.kFactor);
        const exportDp = toNumber(meter.dp);
        const orifice = exportK !== null && exportDp !== null
          ? Math.round(exportK * exportDp * 10000) / 10000
          : toNumber(meter.pressure);
        addUpdate(recordSheet, `F${row}`, orifice);
        addUpdate(recordSheet, `G${row}`, toNumber(meter.volume));
        addUpdate(recordSheet, `H${row}`, toNumber(meter.tmIn));
        addUpdate(recordSheet, `I${row}`, toNumber(meter.tmOut));
        addUpdate(recordSheet, `K${row}`, toNumber(meter.impingerTemp));
        const rowSp = toNumber(meter.sp);
        addUpdate(recordSheet, `P${row}`, rowSp !== null ? rowSp : (Number.isFinite(avgSp) ? avgSp : null));
      });
      // ── 양식 자체의 수식 오류 보정 ──
      // 1) 표준산소 환산 농도(계산식 DC): 양식에는 설명과 달리 (21-4)/(21-5)로 숫자가 박혀 있어
      //    입력한 표준산소(Os)·실측산소(Oa)와 무관한 값이 나온다 → Os·Oa를 쓰도록 바꾼다 (표준산소가 없으면 보정 제외)
      addUpdate('계산식', exportSheetId === 'flow' ? 'M115' : 'M107', { formula: 'IF(Os="",Ca,Ca*((21-Os)/(21-Oa)))' });
      // 2) 여과지 홀더 온도(J열)는 입력 칸이 없어 평균이 #DIV/0!로 찍힌다 → 값이 없으면 빈칸
      {
        const firstRow = layout.meterStartRow;
        const lastRow = layout.meterStartRow + Math.max(1, gasMeters.length - 1);
        addUpdate(recordSheet, `J${lastRow + 1}`, { formula: `IFERROR(AVERAGE(J${firstRow}:J${lastRow}),"")` });
      }
      // 3) 피토관을 고르지 않으면 흡인관 길이(J10)가 #VALUE!로 찍힌다 → 빈칸
      addUpdate(recordSheet, 'J10', { formula: 'IFERROR(노즐산정!D9/100,"")' });

      // 측정점(채취점)이 바뀔 때마다 행 배경색을 바꿔 구분
      const meterRowColors = [];
      {
        let colorIdx = -1;
        let lastPoint = null;
        gasMeters.forEach((meter, idx) => {
          if (idx === 0) return;
          const point = String(meter.pointNum || '').trim();
          if (point !== lastPoint) { colorIdx += 1; lastPoint = point; }
          meterRowColors.push({ row: layout.meterStartRow + idx, colorIdx: Math.max(0, colorIdx) });
        });
      }
      const stylesFile = zip.file('xl/styles.xml');
      const rowStyler = stylesFile ? createMeterRowStyler(await stylesFile.async('string')) : null;

      if (sheetPathMap.has('먼지측정결과보고서')) {
        addUpdate('먼지측정결과보고서', 'C5', sourceData.company);
        addUpdate('먼지측정결과보고서', 'C6', sourceData.location);
        addUpdate('먼지측정결과보고서', 'A17', sourceData.remarks);
      }

      for (const [sheetName, updates] of updatesBySheet.entries()) {
        const sheetPath = sheetPathMap.get(sheetName);
        if (!sheetPath) continue;
        const file = zip.file(sheetPath);
        if (!file) {
          throw new Error(`시트 파일 누락: ${sheetName} (${sheetPath})`);
        }
        const xml = await file.async('string');
        const patched = patchWorksheetXml(
          xml,
          updates,
          sheetName === recordSheet
            ? {
                expandMeterRowsTo: meterExtraRows > 0 ? meterRowsForTemplate : undefined,
                shrinkMeterRowsTo: meterExtraRows < 0 ? meterRowsForTemplate : undefined,
                layout,
                meterRowColors,
                rowStyler,
                // 동압은 소수 첫째 자리로 (측정 행과 평균)
                numberFormats: [{ range: `E${layout.meterStartRow}:E${layout.meterStartRow + meterRowsForTemplate}`, code: '0.0' }],
              }
            : {
                ...(meterExtraRows !== 0
                  ? {
                      shiftRefsToSheet: recordSheet,
                      shiftRefsFromRow: layout.meterTemplateEndRow + 1,
                      shiftRefsDelta: meterExtraRows,
                      extendRangeEndFrom: layout.meterTemplateEndRow,
                    }
                  : {}),
                // 수분량 시트: 측정방법이 세 줄이 되었으므로 선택번호 칸도 세 줄로 합친다
                // 측정방법 표가 한 줄 늘어 바로 아래 제목과 붙었으므로, 제목 줄을 높여 다른 제목들처럼 위에 여백을 둔다
                ...(sheetName === '수분량' ? { removeMerges: ['C5:C6'], addMerges: ['C5:C7'], rowGaps: { 8: 38 } } : {}),
                rowStyler,
                centerCells: EXPORT_CENTER_CELLS(layout)[sheetName] || [],
                ...(sheetName === '노즐산정' ? { numberFormats: [{ range: `J15:J${16 + Math.max(5, sourcePoints.length)}`, code: '0.0' }, { range: 'A34:A34', code: '0.000' }] } : {}),
              }
        );
        zip.file(sheetPath, patched);
      }

      if (rowStyler) zip.file('xl/styles.xml', rowStyler.serialize());

      // 기록지 행 수가 바뀐 경우, 갱신되지 않은 다른 시트(계산식 등)의 기록지 참조도 함께 보정
      if (meterExtraRows !== 0) {
        for (const [sheetName, sheetPath] of sheetPathMap.entries()) {
          if (updatesBySheet.has(sheetName) || sheetName === recordSheet) continue;
          const file = zip.file(sheetPath);
          if (!file) continue;
          const xml = await file.async('string');
          if (!xml.includes(recordSheet)) continue;
          zip.file(sheetPath, patchWorksheetXml(xml, [], {
            shiftRefsToSheet: recordSheet,
            shiftRefsFromRow: layout.meterTemplateEndRow + 1,
            shiftRefsDelta: meterExtraRows,
            extendRangeEndFrom: layout.meterTemplateEndRow,
          }));
        }
      }
      // 계산 체인은 Excel이 다시 만들므로 항상 제거 (수식 칸에 값을 써도 복구 메시지가 뜨지 않게)
      await removeCalcChainFromZip(zip);

      // 열 때 전체 재계산되도록 설정 (양식의 예시 계산값이 남지 않게)
      {
        let wbXml = workbookXml;
        // 기록지 행 수가 바뀌면 인쇄 영역 끝 행도 같이 옮긴다
        if (meterExtraRows !== 0) {
          wbXml = wbXml.replace(/(<definedName\b[^>]*name="_xlnm\.Print_Area"[^>]*>)([^<]*)(<\/definedName>)/g, (full, open, body, close) => {
            if (!body.includes(recordSheet)) return full;
            const nextBody = body.replace(/(\$?[A-Z]{1,3}\$?)(\d+)(?=\s*(?:,|$))/g, (m, colPart, rowText) => {
              const rowNum = parseInt(rowText, 10);
              return rowNum > layout.meterTemplateEndRow ? `${colPart}${Math.max(layout.meterStartRow, rowNum + meterExtraRows)}` : m;
            });
            return `${open}${nextBody}${close}`;
          });
        }
        // 이 기록부에 해당 없는 시트(유량의 누출시험)는 숨긴다
        (layout.hideSheets || []).forEach((name) => {
          wbXml = wbXml.replace(new RegExp(`<sheet\\b([^>]*\\bname="${name}"[^>]*?)\\s*/>`), (full, attrs) => (
            /\bstate="/.test(attrs) ? full : `<sheet${attrs} state="hidden"/>`
          ));
        });
        if (/<calcPr\b[^>]*\/>/.test(wbXml)) {
          wbXml = wbXml.replace(/<calcPr\b([^>]*?)(\s*fullCalcOnLoad="[^"]*")?([^>]*)\/>/, '<calcPr$1$3 fullCalcOnLoad="1"/>');
        } else {
          wbXml = wbXml.replace('</workbook>', '<calcPr fullCalcOnLoad="1"/></workbook>');
        }
        zip.file('xl/workbook.xml', wbXml);
      }

      // 서식 정리: 같은 표는 같은 행 높이, 같은 열은 같은 소수 자릿수, 숫자 서식의 꼬리 공백 제거, 한 장 너비 인쇄
      try {
        const meterFirst = layout.meterStartRow;
        const meterLast = layout.meterStartRow + meterRowsForTemplate; // 평균 행까지
        const postHead = 41 + (layout.impingerRows - 4);
        await tidyWorkbook(zip, {
          // 양식의 시트는 모두 한 장짜리 서식이므로 한 장에 맞춰 인쇄한다
          onePageSheets: 'all',
          rules: {
            연소가스자료: {
              rowHeights: rowsAt(4, 41, 16.5),
              // % 성분은 소수 둘째 자리, ppm 성분은 소수 첫째 자리
              numberFormats: [{ range: 'B5:C41', code: '0.00' }, { range: 'D5:F41', code: '0.0' }],
            },
            시료측정위치: {
              rowHeights: { ...rowsAt(34, 38, 18), ...rowsAt(42, 48, 17.25) },
              numberFormats: [{ range: 'B34', code: '0.00' }, { range: 'A37:C37', code: '0.00' }, { range: 'B44:B48', code: '0.00' }, { range: 'F44:H48', code: '0.00' }],
            },
            수분량: {
              rowHeights: { ...rowsAt(26, 28, 20.1), ...rowsAt(postHead, postHead + 2, 20.1) },
            },
            노즐산정: {
              rowHeights: { ...rowsAt(37, 38, 15), ...rowsAt(39, 56, 18) },
              // 밀도는 모두 소수 셋째 자리, 샘플러 ΔH@·Yd는 넷째 자리
              numberFormats: [{ range: 'H26', code: '0.000' }, { range: 'C34', code: '0.000' }, { range: 'A34', code: '0.000' }, { range: 'O39:P43', code: '0.0000' }],
            },
            [recordSheet]: {
              // 측정 표: 시간·진공은 정수, 온도는 모두 소수 첫째 자리 (측정 행과 평균 행을 같게)
              numberFormats: [
                { range: `B${meterFirst}:C${meterLast}`, code: '0' },
                { range: `D${meterFirst}:D${meterLast}`, code: '0.0' },
                { range: `K${meterFirst}:K${meterLast}`, code: '0.0' },
              ],
            },
            계산식: { numberFormats: [{ range: 'G24:G25', code: '0.0' }, { range: 'J75:J76', code: '0.0' }] },
            // 누출시험 표는 A~F열 18행으로 그리므로 인쇄 범위도 맞춘다 (수은·PAHs 양식은 A1:D17로 잘려 있었다)
            누출시험: { printArea: 'A1:F18' },
            먼지측정결과보고서: { numberFormats: [{ range: 'G10', code: '0.0' }] },
          },
        });
      } catch (error) { console.error('서식 정리 실패:', error); }
      // 계산식 결과를 파일에 미리 채워 둔다: 휴대폰 미리보기·엑셀 보호된 보기처럼 다시 계산하지 않는 화면에서도
      // 밀도·수분량 등 계산값이 바로 보이게 (엑셀은 열 때 다시 계산한다)
      try { await recalcWorkbookCaches(zip); } catch (error) { console.error('계산값 채우기 실패:', error); }
      const output = await zip.generateAsync({ type: 'arraybuffer', compression: 'DEFLATE' });
      const blob = new Blob([output], { type: 'application/vnd.ms-excel.sheet.macroEnabled.12' });
      const url = URL.createObjectURL(blob);

      const link = document.createElement('a');
      link.href = url;
      const targetSheetTitle = sourceData.sheetTitle || activeSheet?.title || '먼지시료채취기록부';
      const exportNames = buildExportNames({ date: sourceData.date, title: targetSheetTitle, company: sourceData.company, location: sourceData.location, ext: 'xlsm' });
      try { if (window.stackpilotDesktop?.noteDownloadFolder) window.stackpilotDesktop.noteDownloadFolder(exportNames.folder); } catch (error) { console.error(error); }
      link.setAttribute('download', exportNames.file);
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      URL.revokeObjectURL(url);
    } catch (error) {
      console.error(error);
      alert('템플릿 엑셀 저장 중 오류가 발생했습니다. 템플릿 파일 또는 입력값을 확인해주세요.');
    }
  };

  // ===== 통합시료채취기록부 =====
  const updateSamplingLogField = (key, value) => setSamplingLogData(prev => ({ ...prev, [key]: value }));

  // ── 샘플러 프리셋의 Yd·ΔH@를 고치면, 그 샘플러를 골라 둔 열린 기록부와 통합시료채취기록부 칸도 따라 바꾼다 ──
  // (다른 창에서 고친 프리셋도 storage 동기화로 들어와 같은 규칙이 적용됨. 저장된 기록은 그대로)
  const prevSamplersRef = useRef(null);
  useEffect(() => {
    const prev = prevSamplersRef.current;
    prevSamplersRef.current = samplers;
    if (!prev || prev === samplers) return;
    const changed = samplers
      .map((cur) => ({ cur, old: prev.find((p) => String(p.id) === String(cur.id)) }))
      .filter(({ cur, old }) => old && (String(old.yd) !== String(cur.yd) || String(old.dHAt) !== String(cur.dHAt)));
    if (changed.length === 0) return;
    setFormData((form) => {
      const hit = changed.find(({ cur }) => String(cur.id) === String(form.samplerId));
      if (!hit) return form;
      return { ...form, gasMeterFactor: String(hit.cur.yd ?? ''), deltaHAt: String(hit.cur.dHAt ?? '') };
    });
  }, [samplers]);
  // 통합시료채취기록부 Yd는 통합시료채취기록부 창에서만 맞춘다 (여러 창이 임시본을 동시에 쓰면 서로 덮어쓰므로).
  // 직접 고친 칸(…YdManual)이 아니면 고른 샘플러 프리셋의 Yd를 늘 따라간다.
  const isSamplingLogWindow = selectedSheet === 'samplinglog';
  useEffect(() => {
    if (!isSamplingLogWindow) return;
    setSamplingLogData((log) => {
      let next = log;
      const pick = (list, idKey, ydKey, manualKey) => {
        if (log[manualKey] === true) return;
        const found = list.find((item) => String(item.id) === String(log[idKey]));
        if (!found) return;
        const yd = String(found.yd ?? '');
        if (String(log[ydKey] ?? '') === yd) return;
        next = { ...next, [ydKey]: yd };
      };
      pick(samplers, 'particleSamplerId', 'particleSamplerYd', 'particleSamplerYdManual');
      pick(gasSamplers, 'gasSamplerId', 'gasSamplerYd', 'gasSamplerYdManual');
      return next;
    });
  }, [isSamplingLogWindow, samplers, gasSamplers, samplingLogData.particleSamplerId, samplingLogData.gasSamplerId, samplingLogData.particleSamplerYd, samplingLogData.gasSamplerYd, samplingLogData.particleSamplerYdManual, samplingLogData.gasSamplerYdManual]);

  // 기록부 한 건에서 통합시료채취기록부 일차 칸에 옮길 현장 값: 날씨, 기온(대기 온도), 기압(대기압). 빈 값은 빼서 통합시료채취기록부 값을 지우지 않는다
  const pickSiteWeatherFromReport = (report) => {
    const patch = {};
    const text = (v) => String(v ?? '').trim();
    if (text(report?.weather)) patch.weather = text(report.weather);
    if (text(report?.atmTemp)) patch.temp = text(report.atmTemp);
    if (text(report?.atmPressure)) patch.pressure = text(report.atmPressure);
    return patch;
  };
  // 사업장·배출구(스택)·날짜가 모두 같은 저장 기록부 가운데 가장 최근 것의 현장 값
  const findSiteWeatherFromSavedReports = (date) => {
    const logCompany = normalizeSiteName(samplingLogData.company);
    const logStack = normalizeSiteName(samplingLogData.stackName);
    if (!logCompany || !logStack || !date) return {};
    const candidates = savedData.filter((report) => LOG_SOURCE_SHEETS.includes(resolveReportSheetId(report))
      && String(report?.date || '').trim() === date
      && normalizeSiteName(report?.company) === logCompany
      && normalizeSiteName(report?.location) === logStack);
    const merged = {};
    // savedData는 최근 저장 순이므로, 앞의 것(최근)이 우선
    candidates.forEach((report) => { Object.entries(pickSiteWeatherFromReport(report)).forEach(([key, value]) => { if (!(key in merged)) merged[key] = value; }); });
    return merged;
  };

  // 기록부 묶음(같은 사업장·배출구·날짜)을 통합시료채취기록부에 옮겨 적기: 사업장 정보, 입자 샘플러, 물질별 시각·가스미터 값
  const applySheetGroupToSamplingLog = (group, { silent = false } = {}) => {
    if (!group) return;
    const reports = Object.entries(group.bySheet);
    if (reports.length === 0) return;
    // 일차 칸이 넘치는 날짜는 미리 계산 (상태 갱신 함수 안에서 모으면 실행 시점이 보장되지 않음)
    const overflowDates = (() => {
      const known = (samplingLogData.days || []).map((d) => d.date).filter(Boolean);
      const freeSlots = (samplingLogData.days || []).filter((d) => !d.date).length;
      const incoming = [...new Set([...(group.dates || [group.date]), ...reports.map(([, r]) => String(r?.date || '').trim())].filter(Boolean))];
      const fresh = incoming.filter((d) => !known.includes(d));
      return fresh.slice(freeSlots);
    })();
    setSamplingLogData(prev => {
      const next = { ...prev, rows: { ...prev.rows }, days: prev.days.map((d) => ({ ...d })) };
      if (group.company) next.company = group.company;
      if (group.location) next.stackName = group.location;
      // 측정일자별 일차 배정: 이미 같은 날짜의 일차가 있으면 그 칸, 없으면 빈 일차에 날짜를 채움 (최대 4일)
      const dayIndexFor = (date) => {
        if (!date) return 0;
        const found = next.days.findIndex((d) => d.date === date);
        if (found >= 0) return found;
        const empty = next.days.findIndex((d) => !d.date);
        if (empty >= 0) { next.days[empty] = { ...next.days[empty], date }; return empty; }
        return -1; // 일차 칸이 모두 찼음
      };
      (group.dates || [group.date]).filter(Boolean).forEach((d) => dayIndexFor(d));
      const samplerSource = reports.map(([, r]) => r).find((r) => String(r?.samplerId || '').trim() !== '');
      if (samplerSource) {
        const found = samplers.find((item) => String(item.id) === String(samplerSource.samplerId));
        if (found) {
          next.particleSamplerId = String(found.id);
          next.particleSamplerYd = String(samplerSource.gasMeterFactor || found.yd || '');
          // 기록부에서 Yd를 프리셋과 다르게 적었으면 그 값을 직접 고친 값으로 지킨다
          next.particleSamplerYdManual = String(next.particleSamplerYd) !== String(found.yd ?? '');
        } else if (samplerSource.gasMeterFactor) {
          next.particleSamplerYd = String(samplerSource.gasMeterFactor);
          next.particleSamplerYdManual = true;
        }
      }
      reports.forEach(([sheetId, report]) => {
        const values = buildSamplingLogValuesFromReport(report);
        // 기록부에 없는 값(빈칸)으로 통합시료채취기록부에 직접 적은 값을 지우지 않는다
        const filledValues = Object.fromEntries(Object.entries(values).filter(([, v]) => v !== null && v !== undefined && String(v).trim() !== ''));
        const dayIndex = dayIndexFor(String(report?.date || '').trim());
        if (dayIndex < 0) return;
        // 기록부에 적은 날씨·대기 온도·대기압을 그 일차의 현장기상에 옮긴다 (기록부에 값이 있는 것만)
        next.days[dayIndex] = { ...next.days[dayIndex], ...pickSiteWeatherFromReport(report) };
        const targetRows = new Set();
        (SHEET_TO_LOG_ROWS[sheetId] || []).forEach((rowNum) => {
          targetRows.add(String(rowNum));
          // 같은 종으로 묶인 행(비소 입자상·가스상 등)은 의뢰 표시만 함께
          const sub = SAMPLING_LOG_SUBSTANCES.find((item) => String(item.row) === String(rowNum));
          if (sub) getSamplingLogRowsOfSpecies(sub).forEach((item) => {
            const pairKey = String(item.row);
            if (!targetRows.has(pairKey)) next.rows[pairKey] = { ...next.rows[pairKey], requested: true };
          });
        });
        targetRows.forEach((key) => {
          next.rows[key] = { ...next.rows[key], requested: true, dayIndex, ...filledValues };
        });
      });
      return next;
    });
    if (overflowDates.length > 0) {
      alert(`통합시료채취기록부 일차 칸(${SAMPLING_LOG_DAY_SLOTS}일)이 모두 차서 ${overflowDates.join(', ')} 기록은 옮기지 못했습니다.\n새 통합시료채취기록부에서 옮겨 주세요.`);
    }
    if (!silent && overflowDates.length === 0) {
      const names = reports.map(([sheetId]) => SHEET_MENU.find((m) => m.id === sheetId)?.title || sheetId).join(', ');
      alert(`${group.company || '사업장'} 기록(${names})을 통합시료채취기록부에 옮겨 적었습니다.`);
    }
  };

  // 기록부 저장 시: 통합시료채취기록부가 비어 있거나 같은 사업장·배출구면 그 물질 행을 자동으로 채운다
  const syncSamplingLogFromSavedReport = (report) => {
    const sheetId = resolveReportSheetId(report);
    if (!LOG_SOURCE_SHEETS.includes(sheetId)) return false;
    const company = String(report?.company || '').trim();
    const location = String(report?.location || '').trim();
    const logCompany = String(samplingLogData.company || '').trim();
    const logStack = String(samplingLogData.stackName || '').trim();
    const sameSite = normalizeSiteName(logCompany) === normalizeSiteName(company)
      && (!logStack || !location || normalizeSiteName(logStack) === normalizeSiteName(location));
    if (logCompany && !sameSite) return false;
    const date = String(report?.date || '').trim();
    // 같은 사업장이라도 지금 통합시료채취기록부의 측정일과 멀리 떨어진 날(다른 회차)이면 섞지 않는다
    const logDates = (samplingLogData.days || []).map((d) => d.date).filter(Boolean);
    if (date && logDates.length > 0 && !logDates.includes(date)) {
      const nearest = Math.min(...logDates.map((d) => Math.abs(daysBetween(d, date))));
      if (!Number.isFinite(nearest) || nearest > CAMPAIGN_MAX_GAP_DAYS) return false;
      if (logDates.length >= SAMPLING_LOG_DAY_SLOTS) return false;
    }
    applySheetGroupToSamplingLog({ company, location, date, dates: date ? [date] : [], bySheet: { [sheetId]: report } }, { silent: true });
    return true;
  };
  const updateSamplingLogDay = (idx, key, value) => setSamplingLogData(prev => ({
    ...prev,
    days: prev.days.map((day, i) => (i === idx ? { ...day, [key]: value } : day)),
  }));
  const updateSamplingLogRow = (rowKey, key, value) => setSamplingLogData(prev => {
    if (key === 'requested' || key === 'speciesSel') {
      // 같은 종의 행(암모니아 흡수액 2종, 비소 입자·가스상, 알데히드 3회차)은 의뢰 여부·종 선택을 함께 바꾼다
      const sub = SAMPLING_LOG_SUBSTANCES.find((item) => String(item.row) === String(rowKey));
      const keys = sub ? getSamplingLogRowsOfSpecies(sub).map((item) => String(item.row)) : [String(rowKey)];
      const rows = { ...prev.rows };
      keys.forEach((k) => { rows[k] = { ...rows[k], [key]: value }; });
      return { ...prev, rows };
    }
    return { ...prev, rows: { ...prev.rows, [rowKey]: { ...prev.rows[rowKey], [key]: value } } };
  });
  const handleSamplingLogParticleSampler = (id) => {
    const found = samplers.find((item) => String(item.id) === String(id));
    setSamplingLogData(prev => ({ ...prev, particleSamplerId: id, particleSamplerYd: found ? String(found.yd) : prev.particleSamplerYd, particleSamplerYdManual: false }));
  };
  const handleSamplingLogGasSampler = (id) => {
    const found = gasSamplers.find((item) => String(item.id) === String(id));
    setSamplingLogData(prev => ({ ...prev, gasSamplerId: id, gasSamplerYd: found ? String(found.yd) : prev.gasSamplerYd, gasSamplerYdManual: false }));
  };
  const resetSamplingLogForm = () => {
    if (!window.confirm('입력 중인 통합시료채취기록부 내용을 모두 지울까요?')) return;
    setSamplingLogData(createEmptySamplingLog());
    unsavedRef.current = false;
  };

  // Yd 칸이 비어 있으면 선택한 샘플러 프리셋의 Yd를 사용
  const resolveSamplingLogYds = (log) => {
    const next = { ...log };
    if (!String(next.particleSamplerYd || '').trim()) {
      const found = samplers.find((item) => String(item.id) === String(next.particleSamplerId));
      if (found) next.particleSamplerYd = String(found.yd);
    }
    if (!String(next.gasSamplerYd || '').trim()) {
      const found = gasSamplers.find((item) => String(item.id) === String(next.gasSamplerId));
      if (found) next.gasSamplerYd = String(found.yd);
    }
    return next;
  };

  const buildSamplingLogReport = () => {
    const log = resolveSamplingLogYds(sanitizeSamplingLog(samplingLogData));
    const firstDay = log.days.find((day) => day.date)?.date || '';
    return {
      id: `${Date.now()}-${Math.random().toString(36).slice(2, 8)}`,
      savedAt: new Date().toISOString(),
      user: activeUser,
      sheetId: 'samplinglog',
      sheetTitle: '통합시료채취기록부',
      date: firstDay,
      company: log.company,
      location: log.stackName,
      sampler: '',
      isokineticStatus: '-',
      actualConcentration: '-',
      samplingLog: log,
    };
  };

  const handleSaveSamplingLog = async ({ silent = false } = {}) => {
    if (!activeUser || !isUserUnlocked) {
      alert('저장하려면 먼저 로그인해주세요.');
      return false;
    }
    const nextReports = [...activeUserReports, buildSamplingLogReport()];
    try {
      await persistReportsEncrypted(activeUser, nextReports);
      setActiveUserReports(nextReports);
      unsavedRef.current = false;
      if (!silent) alert(`[${displayUserNameWithFallback}] 통합시료채취기록부가 저장되었습니다.`);
      return true;
    } catch (error) {
      console.error(error);
      alert('암호화 저장 중 오류가 발생했습니다.');
      return false;
    }
  };

  // ── 저장하지 않은 입력 확인: 창을 닫거나 새 기록지를 열기 전에 묻는다 ──
  useEffect(() => {
    const markDirty = (event) => {
      const target = event.target;
      if (target && target.closest && target.closest('.ds-page')) unsavedRef.current = true;
    };
    document.addEventListener('input', markDirty, true);
    document.addEventListener('change', markDirty, true);
    // 웹 브라우저에서는 기본 확인창으로 막는다 (데스크톱 앱은 메인 프로세스가 저장 여부를 묻는다)
    const onBeforeUnload = (event) => {
      if (window.stackpilotDesktop || !unsavedRef.current) return;
      event.preventDefault();
      event.returnValue = '';
    };
    window.addEventListener('beforeunload', onBeforeUnload);
    return () => {
      document.removeEventListener('input', markDirty, true);
      document.removeEventListener('change', markDirty, true);
      window.removeEventListener('beforeunload', onBeforeUnload);
    };
  }, []);
  useEffect(() => {
    window.__stackpilotHasUnsaved = () => Boolean(unsavedRef.current && selectedSheet && activeUser && isUserUnlocked);
    window.__stackpilotSaveBeforeClose = () => (selectedSheet === 'samplinglog'
      ? handleSaveSamplingLog({ silent: true })
      : handleSave(null, { silent: true }));
  });
  const startNewRecord = () => {
    if (unsavedRef.current && !window.confirm('저장하지 않은 내용이 있습니다.\n저장하지 않고 새 기록지를 열까요?')) return;
    unsavedRef.current = false;
    window.location.reload();
  };

  const extractSerialFromName = (name) => {
    const text = String(name || '').trim();
    const match = /\(([^)]+)\)\s*$/.exec(text);
    return match ? match[1].trim() : text;
  };

  // '시료채취 작성' 시트를 채우면 날짜별 시트는 양식 수식으로 자동 완성됨
  // 공개 양식으로 통합시료채취기록부 내보내기 — 배포용 빌드에서 사용. 의뢰 표시한 항목만 적는다
  const exportOpenSamplingLogExcel = async (sourceReport = null) => {
    try {
      const log = resolveSamplingLogYds(sanitizeSamplingLog(sourceReport && sourceReport.samplingLog ? sourceReport.samplingLog : samplingLogData));
      const samplerName = (list, id) => {
        const found = list.find((item) => String(item.id) === String(id));
        return found ? [found.name, found.sn].filter(Boolean).join(' ') : '';
      };
      const rows = SAMPLING_LOG_SUBSTANCES.filter((sub) => log.rows[String(sub.row)]?.requested).map((sub) => {
        const row = log.rows[String(sub.row)];
        const volume = calcSamplingLogVolume(log, sub);
        const species = Array.isArray(sub.options) && Array.isArray(row.speciesSel) && row.speciesSel.length > 0 ? ` (${row.speciesSel.join(', ')})` : '';
        const notes = [];
        if (String(row.targetVolume || '').trim()) notes.push(`목표 ${row.targetVolume} L`);
        if (String(row.testPeriod || '').trim()) notes.push(/^\d+$/.test(String(row.testPeriod).trim()) ? `시험기한 ${row.testPeriod}일` : String(row.testPeriod));
        return {
          sampleNo: buildSamplingLogSampleNo(log, sub),
          name: `${sub.name}${species}`,
          date: log.days[row.dayIndex]?.date || '',
          start: row.start, end: row.end, meterStart: row.meterStart, meterEnd: row.meterEnd,
          volume: Number.isFinite(volume) ? Math.round(volume * 100) / 100 : null,
          temp: row.temp, pressure: row.pressureH2O, suctionRate: row.suctionRate,
          note: notes.join(' · '),
        };
      });
      const sheets = buildSamplingLogSheets({
        company: log.company, industry: log.industry, facility: log.facility, stackName: log.stackName,
        requestedCount: getSamplingLogStats(log).requestedCount,
        samplers: [
          { label: '입자 시료채취기', name: samplerName(samplers, log.particleSamplerId), yd: log.particleSamplerYd },
          { label: '가스 시료채취기', name: samplerName(gasSamplers, log.gasSamplerId), yd: log.gasSamplerYd },
        ],
        days: log.days, rows,
      });
      const buffer = await buildWorkbook(sheets);
      const firstDay = log.days.find((day) => day.date)?.date || new Date().toISOString().slice(0, 10);
      downloadExportFile(buffer, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        buildExportNames({ date: firstDay, title: '통합시료채취기록부', company: log.company, location: log.stackName, ext: 'xlsx' }));
    } catch (error) {
      console.error(error);
      alert('통합시료채취기록부 엑셀 저장 중 오류가 발생했습니다. 입력값을 확인해주세요.');
    }
  };

  const exportSamplingLogExcel = async (sourceReport = null) => {
    if (IS_DESKTOP_BUILD) { await exportOpenSamplingLogExcel(sourceReport); return; }
    try {
      const log = resolveSamplingLogYds(sanitizeSamplingLog(sourceReport && sourceReport.samplingLog ? sourceReport.samplingLog : samplingLogData));
      const response = await fetch(`${import.meta.env.BASE_URL}templates/${SAMPLING_LOG_TEMPLATE_FILE}`);
      if (!response.ok) throw new Error('템플릿 파일을 불러오지 못했습니다.');
      const zip = await JSZip.loadAsync(await response.arrayBuffer());
      const workbookFile = zip.file('xl/workbook.xml');
      const relsFile = zip.file('xl/_rels/workbook.xml.rels');
      if (!workbookFile || !relsFile) throw new Error('템플릿 workbook 구조를 찾지 못했습니다.');
      let workbookXml = await workbookFile.async('string');
      const sheetPathMap = buildSheetPathMap(workbookXml, await relsFile.async('string'));
      const inputSheet = '시료채취 작성';
      if (!sheetPathMap.has(inputSheet)) throw new Error(`템플릿 시트 누락: ${inputSheet}`);

      const updates = [];
      const add = (ref, value) => updates.push({ ref, value });
      const textOrNull = (v) => { const t = String(v ?? '').trim(); return t ? t : null; };
      const numOrText = (v) => {
        const t = String(v ?? '').trim();
        if (!t) return null;
        return /^-?\d+(\.\d+)?$/.test(t) ? parseFloat(t) : t;
      };

      const gasSampler = gasSamplers.find((item) => String(item.id) === String(log.gasSamplerId));
      add('P2', textOrNull(log.siteAbbr));
      const particleIndex = samplers.findIndex((item) => String(item.id) === String(log.particleSamplerId));
      add('P3', particleIndex >= 0 ? `입자샘플러${particleIndex + 1}` : null);
      add('R3', gasSampler ? `가스샘플러${gasSampler.id}` : null);
      add('C4', textOrNull(log.company));
      add('F4', textOrNull(log.industry));
      add('J4', textOrNull(log.facility));
      add('M4', textOrNull(log.stackName));

      log.days.forEach((day, idx) => {
        const r = 6 + idx;
        add(`C${r}`, toExcelDateSerial(day.date));
        add(`D${r}`, textOrNull(day.weather));
        add(`E${r}`, toNumber(day.temp));
        add(`G${r}`, toNumber(day.humidity));
        add(`I${r}`, toNumber(day.pressure));
        add(`K${r}`, textOrNull(day.windDir));
        add(`M${r}`, toNumber(day.windSpeed));
      });
      for (let i = 0; i < 5; i++) {
        const r = 6 + i;
        const ps = samplers[i];
        add(`O${r}`, ps ? `#${i + 1}` : null);
        add(`P${r}`, ps ? extractSerialFromName(ps.name) : null);
        add(`Q${r}`, ps ? toNumber(ps.yd) : null);
        const gs = gasSamplers[i];
        add(`R${r}`, gs ? `#${i + 1}` : null);
        // 가스 샘플러 일련번호가 숫자뿐이면 숫자로 (엑셀의 '텍스트 형식 숫자' 표시 방지)
        add(`S${r}`, gs ? (/^\d+$/.test(String(gs.sn || '').trim()) ? Number(gs.sn) : gs.sn) : null);
        add(`T${r}`, gs ? toNumber(gs.yd) : null);
      }

      SAMPLING_LOG_SUBSTANCES.forEach((sub) => {
        const r = sub.row;
        const row = log.rows[String(r)];
        const on = Boolean(row && row.requested);
        const day = on ? log.days[row.dayIndex] : null;
        if (!SAMPLING_LOG_MERGED_CHECK_ROWS.has(r)) add(`C${r}`, on ? 'O' : null);
        if (Array.isArray(sub.options)) {
          // 알데히드·옥사이드: 의뢰한 종을 물질명 칸에 함께 적음 (예: 옥사이드 (에틸렌옥사이드))
          const picked = on && Array.isArray(row.speciesSel) && row.speciesSel.length > 0 ? row.speciesSel : sub.options;
          add(`B${r}`, on ? `${sub.name} (${picked.join('·')})` : sub.name);
        }
        add(`D${r}`, on ? buildSamplingLogSampleNo(log, sub) : null);
        add(`F${r}`, on ? toExcelDateSerial(day?.date) : null);
        add(`G${r}`, on ? toExcelTimeFraction(row.start) : null);
        add(`H${r}`, on ? toExcelTimeFraction(row.end) : null);
        add(`I${r}`, on ? toNumber(row.meterStart) : null);
        add(`J${r}`, on ? toNumber(row.meterEnd) : null);
        const vol = on ? calcSamplingLogVolume(log, sub) : NaN;
        add(`K${r}`, Number.isFinite(vol) ? Math.round(vol * 100) / 100 : null);
        add(`L${r}`, on ? toNumber(row.temp) : null);
        add(`M${r}`, on ? toNumber(row.pressureH2O) : null);
        add(`O${r}`, numOrText(row.suctionRate));
        add(`P${r}`, numOrText(row.targetVolume));
        add(`Q${r}`, numOrText(row.duration));
        add(`S${r}`, numOrText(row.testPeriod));
        // 표준 채취량(V): 게이지압(N)이 비면 양식 수식이 #VALUE!를 내므로 빈 압력은 0으로 본다
        if (r >= 14) add(`V${r}`, { formula: `IF(K${r}="","",K${r}*(273/(273+$L${r}))*($I$6+N($N${r}))/760)` });
      });

      const inputPath = sheetPathMap.get(inputSheet);
      const inputXml = await zip.file(inputPath).async('string');
      zip.file(inputPath, patchWorksheetXml(inputXml, updates, { removeMerges: SAMPLING_LOG_UNMERGE_ON_EXPORT }));

      // 날짜별 시트: 양식에는 1·2일차 시트만 있으므로, 완전한 1일차 시트를 복사해
      // 각 일차가 입력 시트의 자기 날짜 행(6행=1일차 … 9행=4일차)을 읽도록 만든다.
      // 2일차 시트도 다시 만든다 (양식의 2일차 시트는 B15 수식이 빠져 있음)
      {
        const templateDaySheets = Array.from(workbookXml.matchAll(/<sheet\b[^>]*\bname="([^"]+)"/g))
          .map((m) => m[1]).filter((name) => name !== inputSheet);
        const firstDayPath = sheetPathMap.get(templateDaySheets[0]);
        const firstDayXml = firstDayPath ? await zip.file(firstDayPath).async('string') : '';
        const inputRef = inputSheet.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
        const buildDaySheetXml = (dayIdx) => firstDayXml
          // 입력 시트의 날짜 행 참조 6 → 6+일차
          .replace(new RegExp(`('${inputRef}'!\\$?[A-Z]{1,3}\\$?)6(?!\\d)`, 'g'), `$1${6 + dayIdx}`)
          // 분석 기한(N열)은 같은 행의 M열을 보도록 (양식 N16이 M17을 보던 것 바로잡음)
          .replace(/(<c r="N(\d+)"[^>]*><f>[^<]*?DAY\(M)\d+(\))/g, '$1$2$3')
          // 복사본에 남은 1일차 계산값은 지운다 (열 때 다시 계산)
          .replace(/(<c\b[^>]*?)\s+t="(?:str|e|n|b)"([^>]*>(?:<f\b[^>]*>[^<]*<\/f>|<f\b[^>]*\/>))<v>[^<]*<\/v>/g, '$1$2')
          .replace(/(<c\b[^>]*>(?:<f\b[^>]*>[^<]*<\/f>|<f\b[^>]*\/>))<v>[^<]*<\/v>/g, '$1')
          .replace(/\s+tabSelected="1"/g, '')
          .replace(/\s+xr:uid="[^"]*"/g, '')
          // 양식의 깨진 참조 바로잡기: 날짜 시트 각 칸은 입력 시트의 같은 물질 행(F열 날짜와 같은 행)을 읽어야 함
          // (양식 C40이 E42를, C41이 #REF!를 보던 것)
          .replace(
            new RegExp(`(<c r="([B-K])\\d+"[^>]*><f>IF\\(\\$J\\$7='${inputRef}'!\\$?F(\\d+),'${inputRef}'!)(?:\\$?[A-Z]{1,3}\\$?\\d+|#REF!)(,"")`, 'g'),
            (full, head, col, srcRow, tail) => `${head}${DAY_SHEET_INPUT_COLUMNS[col]}${srcRow}${tail}`,
          );
        const dayInUse = (dayIdx) => Boolean(log.days[dayIdx]?.date)
          || SAMPLING_LOG_SUBSTANCES.some((sub) => {
            const row = log.rows[String(sub.row)];
            return row && row.requested && Number(row.dayIndex) === dayIdx;
          });
        if (firstDayXml) {
          // 1·2일차: 기존 시트 파일도 1일차 기반으로 다시 만든다 (양식의 깨진 참조 수정 포함)
          zip.file(firstDayPath, buildDaySheetXml(0));
          const secondDayPath = sheetPathMap.get(templateDaySheets[1]);
          if (secondDayPath) zip.file(secondDayPath, buildDaySheetXml(1));
          let relsXml = await zip.file('xl/_rels/workbook.xml.rels').async('string');
          const ctFile = zip.file('[Content_Types].xml');
          let ctXml = ctFile ? await ctFile.async('string') : '';
          const firstDayRelsPath = firstDayPath.replace(/worksheets\/([^/]+)$/, 'worksheets/_rels/$1.rels');
          const firstDayRels = zip.file(firstDayRelsPath) ? await zip.file(firstDayRelsPath).async('string') : null;
          for (let dayIdx = templateDaySheets.length; dayIdx < log.days.length; dayIdx += 1) {
            if (!dayInUse(dayIdx)) continue;
            const existingFiles = Object.keys(zip.files).filter((f) => /^xl\/worksheets\/sheet\d+\.xml$/.test(f));
            const nextNum = Math.max(0, ...existingFiles.map((f) => parseInt(f.match(/(\d+)\.xml$/)[1], 10))) + 1;
            const partPath = `xl/worksheets/sheet${nextNum}.xml`;
            zip.file(partPath, buildDaySheetXml(dayIdx));
            if (firstDayRels) zip.file(`xl/worksheets/_rels/sheet${nextNum}.xml.rels`, firstDayRels);
            const relIds = Array.from(relsXml.matchAll(/Id="rId(\d+)"/g)).map((m) => parseInt(m[1], 10));
            const relId = `rId${Math.max(0, ...relIds) + 1}`;
            relsXml = relsXml.replace('</Relationships>', `<Relationship Id="${relId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${nextNum}.xml"/></Relationships>`);
            if (ctXml) ctXml = ctXml.replace('</Types>', `<Override PartName="/${partPath}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>`);
            const sheetIds = Array.from(workbookXml.matchAll(/<sheet\b[^>]*\bsheetId="(\d+)"/g)).map((m) => parseInt(m[1], 10));
            const position = Array.from(workbookXml.matchAll(/<sheet\b/g)).length;
            const tempName = `${dayIdx + 1}일차`;
            workbookXml = workbookXml.replace('</sheets>', `<sheet name="${tempName}" sheetId="${Math.max(0, ...sheetIds) + 1}" r:id="${relId}"/></sheets>`);
            const printArea = `<definedName name="_xlnm.Print_Area" localSheetId="${position}">'${tempName}'!$A$2:$K$59</definedName>`;
            workbookXml = workbookXml.includes('</definedNames>')
              ? workbookXml.replace('</definedNames>', `${printArea}</definedNames>`)
              : workbookXml.replace('</sheets>', `</sheets><definedNames>${printArea}</definedNames>`);
          }
          zip.file('xl/_rels/workbook.xml.rels', relsXml);
          if (ctFile) zip.file('[Content_Types].xml', ctXml);
        }
      }

      // 쓰지 않는 2일차 시트는 지운다 (양식에 들어 있던 예시 날짜 이름의 시트가 그대로 남지 않게)
      const templateDayNames = Array.from(workbookXml.matchAll(/<sheet\b[^>]*\bname="([^"]+)"/g)).map((m) => m[1]).filter((name) => name !== inputSheet);
      const isDayUsed = (dayIdx) => Boolean(log.days[dayIdx]?.date)
        || SAMPLING_LOG_SUBSTANCES.some((sub) => { const row = log.rows[String(sub.row)]; return row && row.requested && Number(row.dayIndex) === dayIdx; });
      // 시트 이름 → 일차 번호: 양식의 두 시트는 1·2일차, 새로 만든 시트는 "N일차"
      const dayIndexOfSheet = (name) => {
        const added = /^(\d+)일차$/.exec(name);
        if (added) return parseInt(added[1], 10) - 1;
        return templateDayNames.indexOf(name);
      };
      if (templateDayNames.length >= 2 && !isDayUsed(1)) {
        const unusedName = templateDayNames[1];
        const sheetTags = Array.from(workbookXml.matchAll(/<sheet\b[^>]*\/>/g)).map((m) => m[0]);
        const position = sheetTags.findIndex((tag) => tag.includes(`name="${unusedName}"`));
        const relId = position >= 0 ? (/r:id="([^"]+)"/.exec(sheetTags[position]) || [])[1] : '';
        if (position >= 0 && relId) {
          workbookXml = workbookXml.replace(sheetTags[position], '');
          // 그 시트의 인쇄 영역 등 이름 정의를 지우고, 뒤 시트들의 번호를 하나씩 당긴다
          workbookXml = workbookXml.replace(/<definedName\b[^>]*\blocalSheetId="(\d+)"[^>]*>[^<]*<\/definedName>/g, (full, idText) => {
            const id = parseInt(idText, 10);
            if (id === position) return '';
            return id > position ? full.replace(`localSheetId="${id}"`, `localSheetId="${id - 1}"`) : full;
          });
          workbookXml = workbookXml.replace(/(<workbookView\b[^>]*\bactiveTab=")(\d+)(")/, (full, head, tab, tail) => `${head}${parseInt(tab, 10) >= position ? 0 : tab}${tail}`);
          let relsNow = await zip.file('xl/_rels/workbook.xml.rels').async('string');
          const relTag = new RegExp(`<Relationship\\b[^>]*\\bId="${relId}"[^>]*/>`).exec(relsNow);
          const target = relTag ? (/Target="([^"]+)"/.exec(relTag[0]) || [])[1] : '';
          if (relTag) relsNow = relsNow.replace(relTag[0], '');
          zip.file('xl/_rels/workbook.xml.rels', relsNow);
          if (target) {
            const partPath = `xl/${target.replace(/^\//, '').replace(/^xl\//, '')}`;
            zip.remove(partPath);
            zip.remove(partPath.replace(/worksheets\/([^/]+)$/, 'worksheets/_rels/$1.rels'));
            const ctNow = zip.file('[Content_Types].xml');
            if (ctNow) zip.file('[Content_Types].xml', (await ctNow.async('string')).replace(new RegExp(`<Override\\b[^>]*PartName="/${partPath.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}"[^>]*/>`), ''));
          }
        }
      }

      // 날짜별 시트 이름을 실제 채취일(yyyymmdd)로 변경 (인쇄 영역 이름도 함께)
      const sheetNames = Array.from(workbookXml.matchAll(/<sheet\b[^>]*\bname="([^"]+)"/g)).map((m) => m[1]);
      const dayNames = log.days.map((day) => {
        const dt = day.date ? new Date(day.date) : null;
        return dt && !Number.isNaN(dt.getTime())
          ? `${dt.getFullYear()}${String(dt.getMonth() + 1).padStart(2, '0')}${String(dt.getDate()).padStart(2, '0')}`
          : '';
      });
      const used = new Set(sheetNames);
      sheetNames.filter((name) => name !== inputSheet).forEach((oldName) => {
        // 순서가 아니라 그 시트가 맡은 일차의 날짜로 이름을 붙인다 (중간 일차를 건너뛴 경우에도 맞게)
        const nextName = dayNames[dayIndexOfSheet(oldName)];
        if (!nextName || nextName === oldName || used.has(nextName)) return;
        used.add(nextName);
        workbookXml = workbookXml
          .split(`name="${oldName}"`).join(`name="${nextName}"`)
          .split(`'${oldName}'!`).join(`'${nextName}'!`);
      });
      if (/<calcPr\b[^>]*\/>/.test(workbookXml)) {
        workbookXml = workbookXml.replace(/<calcPr\b([^>]*?)(\s*fullCalcOnLoad="[^"]*")?([^>]*)\/>/, '<calcPr$1$3 fullCalcOnLoad="1"/>');
      } else {
        workbookXml = workbookXml.replace('</workbook>', '<calcPr fullCalcOnLoad="1"/></workbook>');
      }
      zip.file('xl/workbook.xml', workbookXml);
      await removeCalcChainFromZip(zip);

      // 서식 정리: 날짜별 시트는 한 장에, 입력 시트의 기상 숫자는 가운데, 샘플러 Yd 자릿수 통일, 물질 행 높이 통일
      try {
        await tidyWorkbook(zip, {
          onePageSheets: 'all',
          // 숫자 열은 자릿수를 하나로: 가스미터·채취유량 소수 둘째, 온도·게이지압(mmH2O) 소수 첫째, 압력(mmHg) 소수 둘째
          rules: (name) => (name === inputSheet
            ? {
                rowHeights: rowsAt(13, 57, 17.25),
                numberFormats: [
                  { range: 'Q6:Q10', code: '0.0000' }, { range: 'T6:T10', code: '0.0000' },
                  { range: 'I13:K57', code: '0.00' }, { range: 'L13:M57', code: '0.0' }, { range: 'N13:N57', code: '0.00' },
                ],
                center: ['E6:M9'],
              }
            : {
                numberFormats: [{ range: 'F12:H56', code: '0.00' }, { range: 'I12:J56', code: '0.0' }, { range: 'K12:K56', code: '0.00' }],
              }),
        });
      } catch (error) { console.error('서식 정리 실패:', error); }
      // 계산식 결과를 파일에 미리 채워 둔다: 휴대폰 미리보기·엑셀 보호된 보기처럼 다시 계산하지 않는 화면에서도
      // 밀도·수분량 등 계산값이 바로 보이게 (엑셀은 열 때 다시 계산한다)
      try { await recalcWorkbookCaches(zip); } catch (error) { console.error('계산값 채우기 실패:', error); }
      const output = await zip.generateAsync({ type: 'arraybuffer', compression: 'DEFLATE' });
      const blob = new Blob([output], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const url = URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = url;
      const firstDay = log.days.find((day) => day.date)?.date || new Date().toISOString().slice(0, 10);
      const exportNames = buildExportNames({ date: firstDay, title: '통합시료채취기록부', company: log.company, location: log.stackName, ext: 'xlsx' });
      try { if (window.stackpilotDesktop?.noteDownloadFolder) window.stackpilotDesktop.noteDownloadFolder(exportNames.folder); } catch (error) { console.error(error); }
      link.setAttribute('download', exportNames.file);
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      URL.revokeObjectURL(url);
    } catch (error) {
      console.error(error);
      alert('통합시료채취기록부 엑셀 저장 중 오류가 발생했습니다. 템플릿 파일 또는 입력값을 확인해주세요.');
    }
  };

  if (import.meta.env.VITE_E2E === '1' && typeof window !== 'undefined') {
    window.__stackpilotExportForTest = exportToTemplateExcel;
    window.__stackpilotExportOpenForTest = exportOpenRecordExcel;
    window.__stackpilotExportOpenSamplingLogForTest = exportOpenSamplingLogExcel;
    window.__stackpilotBuildReportForTest = () => buildReportSnapshot(selectedSheet || 'dust');
    window.__stackpilotExportSamplingLogForTest = exportSamplingLogExcel;
    window.__stackpilotSetSamplingLogForTest = (value) => setSamplingLogData(sanitizeSamplingLog(value));
  }

  const handleExportSelectedTemplateExcel = async () => {
    if (!activeUser || !isUserUnlocked) {
      alert('로그인 후 추출할 수 있습니다.');
      return;
    }
    if (sheetCheckedReportKeysValid.length === 0) {
      alert('템플릿으로 추출할 리포트를 먼저 선택해주세요.');
      return;
    }

    const selectedSet = new Set(sheetCheckedReportKeysValid);
    const targetReports = savedData.filter((item) => selectedSet.has(buildReportKey(item.id, item.savedAt)));
    if (targetReports.length === 0) {
      alert('선택된 리포트를 찾지 못했습니다.');
      return;
    }

    for (const report of targetReports) {
      // 선택된 리포트 내용을 기반으로 템플릿 파일을 순차 추출합니다.
      // 파일명은 각 리포트의 기록부명+측정일자로 생성됩니다.
      // eslint-disable-next-line no-await-in-loop
      if (resolveReportSheetId(report) === 'samplinglog') await exportSamplingLogExcel(report);
      else await exportToTemplateExcel(report);
    }
  };

  // ── 피토계수 AI 자동 선정 (모든 계산 함수 정의 이후) ──
  // ① 기준계수 0.84로 유속을 가계산 → ② 피토관의 유속별 Cp 표에서 그 유속이 속한 구간(행의 유속 = 구간 상한)의 Cp를 선정
  // → ③ 선정한 Cp로 유속을 다시 계산하고, Cp가 더 이상 바뀌지 않을 때까지 반복(수렴)
  // 유속 출처: 기록표(적산유량 행)에 동압이 있으면 기록표, 없으면 예비조사(4번 섹션)
  const buildPitotAutoSelection = () => {
    if (!isPitotCpAutoApplied || !activePitotPreset) return null;
    const meterV = getRawGasVelocityFromMeterRows(PITOT_BASE_CP);
    const useMeter = Number.isFinite(meterV) && meterV > 0;
    const baseVelocity = useMeter ? meterV : getRawGasVelocity(PITOT_BASE_CP);
    const source = useMeter ? '기록표' : '예비조사';
    if (!Number.isFinite(baseVelocity) || baseVelocity <= 0) return { ready: false, source, steps: [] };
    const kFactor = baseVelocity / PITOT_BASE_CP; // 유속 = Cp × K (K는 동압·밀도만의 함수)
    const steps = [{ cp: PITOT_BASE_CP, velocity: baseVelocity, note: '기준계수 가계산' }];
    let cp = resolvePitotCpByVelocity(activePitotPreset, baseVelocity);
    let velocity = baseVelocity;
    let converged = false;
    const seenCps = [];
    for (let i = 0; i < 6; i += 1) {
      velocity = cp * kFactor;
      steps.push({ cp, velocity, note: i === 0 ? '표에서 선정한' : '재계산' });
      seenCps.push(cp);
      const nextCp = resolvePitotCpByVelocity(activePitotPreset, velocity);
      if (!Number.isFinite(nextCp) || Math.abs(nextCp - cp) < 0.00005) { converged = true; break; }
      // 구간 경계에서 두 Cp를 오가면 지금 Cp(와 그 유속)로 멈춘다
      if (seenCps.some((seen) => Math.abs(seen - nextCp) < 0.00005)) break;
      cp = nextCp;
    }
    velocity = cp * kFactor;
    const rows = activePitotVelocityRows;
    const minV = 0;
    const maxV = rows[rows.length - 1].v;
    const inRange = velocity <= maxV;
    const coversVelocity = (preset) => {
      const presetRows = getPitotVelocityRows(preset);
      return presetRows.length > 0 && velocity <= presetRows[presetRows.length - 1].v;
    };
    const alternative = inRange
      ? null
      // 다른 피토관을 권할 때도 이 기록부 용도의 피토관 중에서만
      : (sheetPitotPresets.find((preset) => preset.id !== activePitotPreset.id && coversVelocity(preset)) || null);
    return { ready: true, source, steps, finalCp: cp, finalVelocity: velocity, converged, inRange, minV, maxV, alternative };
  };
  const pitotAutoSelection = buildPitotAutoSelection();
  const appliedPitotVelocity = pitotAutoSelection?.ready ? pitotAutoSelection.finalVelocity : NaN;
  const appliedPitotVelocitySource = pitotAutoSelection?.source || '예비조사';
  const autoSelectedPitotCp = pitotAutoSelection?.ready
    ? pitotAutoSelection.finalCp
    : (isPitotCpAutoApplied ? resolvePitotCpByVelocity(activePitotPreset, NaN) : NaN);
  useEffect(() => {
    if (!isPitotCpAutoApplied || !Number.isFinite(autoSelectedPitotCp)) return;
    const current = parseFloat(formData.pitotFactor);
    if (Number.isFinite(current) && Math.abs(current - autoSelectedPitotCp) < 0.00005) return;
    setFormData(prev => ({ ...prev, pitotFactor: formatPitotCp(autoSelectedPitotCp) }));
  }, [isPitotCpAutoApplied, autoSelectedPitotCp, formData.pitotFactor]);

  const appGasComp = getGasComposition();
  const o2 = appGasComp.o2, co2 = appGasComp.co2, co = appGasComp.co, sox = appGasComp.sox, nox = appGasComp.nox, n2 = appGasComp.n2, Md = appGasComp.Md, Ms = appGasComp.Ms, r0 = appGasComp.r0;
  
  const currentTs = getRawAvgTs();
  const currentPs = getRawStackPressure();
  const r = (!isNaN(r0) && !isNaN(currentTs) && !isNaN(currentPs)) ? r0 * (273 / (273 + currentTs)) * (currentPs / 760) : NaN;
  const samplingPointsData = getSamplingPoints();
  // 4번 섹션 측정점 행 수를 굴뚝 내경으로 정해진 측정점 수(반경당 1~5점)에 맞춤
  const samplingPointCount = samplingPointsData ? samplingPointsData.perRadius : 0;
  useEffect(() => {
    if (!samplingPointCount) return;
    setFormData(prev => {
      if (prev.points.length === samplingPointCount) return prev;
      const nextPoints = Array.from({ length: samplingPointCount }, (_, i) => ({
        ...(prev.points[i] || { tp: '', sp: '', dp: '', ts: '' }),
        id: i + 1,
      }));
      return { ...prev, points: nextPoints };
    });
  }, [samplingPointCount]);
  // 기록표에 아직 입력한 값이 없으면 측정점 수·채취시간에 맞춰 표를 자동으로 만들어 둔다
  const meterTableHasEntries = formData.gasMeters.some((row, idx) => (
    [row?.dp, row?.sp, row?.pressure, row?.volume, row?.stackTemp, row?.tmIn, row?.tmOut, row?.vacuum, row?.impingerTemp]
      .some((v) => String(v ?? '').trim() !== '') && (idx > 0 || String(row?.volume ?? '').trim() !== '')
  ));
  const autoTablePlanKey = samplingPointCount
    ? `${samplingPointCount}|${tableTotalMinutes}|${formData.planSamplingTime}`
    : '';
  useEffect(() => {
    if (!autoTablePlanKey || meterTableHasEntries || isFlowSheet) return;
    applySamplingPointsToTable({ silent: true });
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [autoTablePlanKey]);
  const configuredNozzles = getConfiguredNozzles();
  const gasFlowRates = calcGasFlowRates();
  
  const getExpectedValues = () => {
    const d = parseFloat(formData.nozzleDiameter), k = parseFloat(formData.kFactor), dp = getRawAvgDp(), Vs = getRawGasVelocity();

    let dH = '-', L = '-', SL = '-', Sm3 = '-', reqTime = '-';
    if (!isNaN(k) && !isNaN(dp)) dH = calcExpectedOrificeDH(k, dp);

    const Ts = getRawAvgTs(), Ps = getRawStackPressure(), Pm = parseFloat(formData.atmPressure);
    const Tm = getTraverseMeterTempAvg();

    const gasComp = getGasComposition();
    const Xw = gasComp.Xw;
    const Yd = parseFloat(formData.gasMeterFactor) || 1.0;
    const targetVol = parseFloat(formData.targetVolume) || 1000;

    if (!isNaN(d) && !isNaN(Vs) && Vs > 0 && !isNaN(Ts) && !isNaN(Tm) && !isNaN(Ps) && Pm > 0 && !isNaN(Xw)) {
        const An = Math.PI * Math.pow(d / 2000, 2);
        const Q_m = (Vs * An * 60 * 1000 * ((Tm + 273) / (Ts + 273)) * (Ps / Pm) * (1 - Xw / 100)) / Yd;
        const Q_sl = Q_m * Yd * (273 / (273 + Tm)) * (Pm / 760);
        
        const calcTime = Math.ceil(targetVol / Q_sl);
        if (isFinite(calcTime) && calcTime > 0) {
            reqTime = calcTime;
            const Vm_val = Q_m * reqTime;
            L = Vm_val.toFixed(2);
            const Vsl_val = Q_sl * reqTime;
            SL = Vsl_val.toFixed(2);
            Sm3 = (Vsl_val / 1000).toFixed(3);
        }
    }
    return { dH, L, SL, Sm3, reqTime };
  };

  const getMeterTableAverages = () => {
    const activeRowsWithIndex = formData.gasMeters
      .map((row, idx) => ({ row, idx }))
      .filter(({ row, idx }) => isActiveMeterRow(row, idx));
    const samplingRowsWithIndex = activeRowsWithIndex.filter(({ idx }) => idx > 0);

    const avgOf = (values) => {
      const valid = values.filter((v) => Number.isFinite(v));
      if (valid.length === 0) return NaN;
      return valid.reduce((acc, cur) => acc + cur, 0) / valid.length;
    };

    const fmt = (value, digits) => (Number.isFinite(value) ? roundFixed(value, digits) : '-');

    // 엑셀 평균 행(C·D·E·F·H·K·P 28행)과 같이 시작행(0분)은 빼고 평균
    const avgRows = isFlowRecordSheet ? activeRowsWithIndex : samplingRowsWithIndex;
    const tsAvg = avgOf(avgRows.map(({ row }) => parseFloat(row.stackTemp)));
    const dpAvg = avgOf(avgRows.map(({ row }) => parseFloat(row.dp)));
    const spAvg = avgOf(avgRows.map(({ row }) => parseFloat(row.sp)));
    const pressureAvg = avgOf(avgRows.map(({ row }) => getRawOrificePressureFromRow(row)));
    const tmInAvg = avgOf(avgRows.map(({ row }) => parseFloat(row.tmIn)));
    const tmOutAvg = avgOf(avgRows.map(({ row }) => parseFloat(row.tmOut)));
    const vacuumAvg = avgOf(avgRows.map(({ row }) => parseFloat(row.vacuum)));
    const impingerTempAvg = avgOf(avgRows.map(({ row }) => parseFloat(row.impingerTemp)));
    const velocityAvg = avgOf(samplingRowsWithIndex.map(({ idx }) => getRowGasVelocityRaw(idx)));
    const densityAvg = avgOf(samplingRowsWithIndex.map(({ idx }) => getRowCorrectedGasDensityRaw(idx)));
    const lastSamplingIdx = samplingRowsWithIndex.length > 0
      ? samplingRowsWithIndex[samplingRowsWithIndex.length - 1].idx
      : -1;
    const totalTime = getSamplingMinutes();
    const totalVolume = getRawGasMeterVolDiff();
    const finalMoisture = lastSamplingIdx > 0 ? getRowMoistureCaptureRaw(lastSamplingIdx) : NaN;
    const finalRate = lastSamplingIdx > 0 ? parseFloat(calcRowIsokineticRate(lastSamplingIdx)) : NaN;

    const tmPair = (Number.isFinite(tmInAvg) || Number.isFinite(tmOutAvg))
      ? `${fmt(tmInAvg, 1)} / ${fmt(tmOutAvg, 1)}`
      : '-';

    return {
      time: fmt(totalTime, 1),
      stackTemp: fmt(tsAvg, 1),
      dp: fmt(dpAvg, 1),
      sp: fmt(spAvg, 2),
      pressure: fmt(pressureAvg, 2),
      volume: fmt(totalVolume, 2),
      tmPair,
      vacuum: fmt(vacuumAvg, 1),
      impingerTemp: fmt(impingerTempAvg, 1),
      velocity: fmt(velocityAvg, 2),
      density: fmt(densityAvg, 2),
      moisture: fmt(finalMoisture, 2),
      rate: fmt(finalRate, 1),
    };
  };

  const expData = getExpectedValues();
  const meterTableAverages = getMeterTableAverages();
  const activeSheet = SHEET_MENU.find(item => item.id === selectedSheet);
  const activeTheme = SHEET_THEMES[selectedSheet] || SHEET_THEMES.dust;
  const isDustSheet = selectedSheet === 'dust';
  const isFlowSheet = selectedSheet === 'flow';
  const impingerStageCount = getImpingerStageCount(selectedSheet || 'dust');
  const visibleImpingers = normalizeImpingers(formData.impingers, impingerStageCount);
  const resolvedSkyMode = skyPreviewMode === 'auto' ? skyPhase : skyPreviewMode;
  const isNightSky = resolvedSkyMode === 'night';
  const isMorningSky = resolvedSkyMode === 'morning';
  const skyOrbClass = isNightSky ? 'skyline-moon' : (isMorningSky ? '' : 'skyline-sun');
  const menuSkyBgClass = isNightSky
    ? 'bg-gradient-to-b from-[#050914] via-[#264284] to-[#7a5a8c]'
    : (isMorningSky
      ? 'bg-gradient-to-b from-[#99ceff] via-[#8ab8ef] to-[#f1b7c8]'
      : 'bg-gradient-to-b from-[#dff3ff] via-[#bee5ff] to-[#ffe8be]');
  const menuSkyOverlayClass = isNightSky
    ? 'bg-gradient-to-b from-[#040915]/45 via-[#1d2f64]/15 to-transparent'
    : (isMorningSky
      ? 'bg-gradient-to-b from-[#fff2cc]/38 via-[#ffdca5]/20 to-transparent'
      : 'bg-gradient-to-b from-white/42 via-sky-100/12 to-transparent');
  const menuSkyGlowLeftClass = isNightSky
    ? 'bg-indigo-300/20'
    : (isMorningSky ? 'bg-amber-100/70' : 'bg-white/85');
  const menuSkyGlowRightClass = isNightSky
    ? 'bg-violet-400/15'
    : (isMorningSky ? 'bg-orange-200/45' : 'bg-sky-200/70');
  const sheetSkyBgClass = isNightSky
    ? 'bg-gradient-to-b from-[#050914] via-[#264284] to-[#7a5a8c]'
    : (isMorningSky
      ? 'bg-gradient-to-b from-[#99ceff] via-[#8ab8ef] to-[#f1b7c8]'
      : 'bg-gradient-to-b from-[#dff3ff] via-[#bee5ff] to-[#ffe8be]');
  const sheetSkyOverlayClass = isNightSky
    ? 'bg-gradient-to-b from-[#040915]/42 via-[#1d2f64]/16 to-transparent'
    : (isMorningSky
      ? 'bg-gradient-to-b from-[#fff3d6]/34 via-[#ffe4b8]/16 to-transparent'
      : 'bg-gradient-to-b from-white/34 via-sky-100/10 to-transparent');
  const sceneCloudColor = isNightSky ? 'rgba(190, 200, 235, 0.09)' : (isMorningSky ? 'rgba(255, 243, 230, 0.9)' : 'rgba(255, 255, 255, 0.93)');
  const sceneCloudShadow = isNightSky ? 'rgba(0, 0, 0, 0)' : (isMorningSky ? 'rgba(240, 150, 120, 0.2)' : 'rgba(120, 140, 170, 0.2)');
  const sceneClouds = isMobileViewport ? SKYLINE_CLOUDS.slice(0, 4) : SKYLINE_CLOUDS;
  // 하늘 장식: 구름, (밤) 달무리·별똥별 — 메뉴와 기록부 화면이 공용으로 사용
  const renderSkyDecor = (scope, cloudCount = sceneClouds.length) => (
    <>
      {isNightSky && <div className="skyline-moon-halo" />}
      {isMorningSky && <div className={scope === 'menu' ? 'skyline-sunrise-mountain' : 'skyline-sunrise-mountain-sheet'} />}
      {sceneClouds.slice(0, cloudCount).map((cloud) => (
        <div
          key={`${scope}-${cloud.id}`}
          className="skyline-cloud"
          style={{
            top: cloud.top,
            opacity: cloud.opacity,
            '--cloud-scale': cloud.scale,
            '--cloud-duration': `${cloud.duration}s`,
            '--cloud-delay': `${cloud.delay}s`,
            '--cloud-color': sceneCloudColor,
            '--cloud-shadow': sceneCloudShadow,
          }}
        />
      ))}
      {isNightSky && (
        <>
          <div className="skyline-shooting-star" style={{ top: '9%', left: '72%', '--shoot-duration': '14s', '--shoot-delay': '-3s' }} />
          <div className="skyline-shooting-star" style={{ top: '18%', left: '38%', '--shoot-duration': '23s', '--shoot-delay': '-15s' }} />
        </>
      )}
    </>
  );
  const isMobileLightMode = isMobileViewport;
  const sceneStarsMenu = isMobileLightMode ? makeSparseLayout(SKYLINE_STARS, 3) : SKYLINE_STARS;
  const sceneStarsSheet = isMobileLightMode ? makeSparseLayout(SKYLINE_STARS.slice(0, 30), 3) : SKYLINE_STARS.slice(0, 30);
  const navigateToMenu = () => {
    if (IS_POPUP_WINDOW) {
      // 데스크톱 앱: 창의 닫기 경로로 보내 저장 확인을 받는다. 웹: 저장 안 한 내용이 있으면 먼저 묻는다
      if (window.stackpilotDesktop && window.stackpilotDesktop.closeWindow) {
        window.stackpilotDesktop.closeWindow();
        return;
      }
      if (unsavedRef.current && !window.confirm('저장하지 않은 내용이 있습니다.\n저장하지 않고 창을 닫을까요?')) return;
      window.close();
      return;
    }
    setSelectedSheet('');
    if (window.location.hash) {
      window.location.hash = '';
    }
  };
  const navigateToSheet = (sheetId) => {
    if (!activeUser || !isUserUnlocked) {
      alert('첫 화면에서 로그인한 뒤 이용해주세요.');
      return;
    }
    if (selectedSheet === sheetId) return;
    if (!IS_POPUP_WINDOW) {
      openSheetWindow(sheetId);
      return;
    }
    window.location.hash = `sheet=${sheetId}`;
  };

  useEffect(() => {
    const applyHashRoute = () => {
      const rawHash = window.location.hash.startsWith('#') ? window.location.hash.slice(1) : '';
      const params = new URLSearchParams(rawHash);
      const hashSheet = params.get('sheet');
      if (!hashSheet || !SHEET_MENU.some(sheet => sheet.id === hashSheet)) {
        setSelectedSheet('');
        return;
      }
      if (!activeUser || !isUserUnlocked) {
        setSelectedSheet('');
        return;
      }
      setSelectedSheet(hashSheet);
    };

    applyHashRoute();
    window.addEventListener('hashchange', applyHashRoute);
    return () => window.removeEventListener('hashchange', applyHashRoute);
  }, [activeUser, isUserUnlocked]);

  useEffect(() => {
    if (!selectedSheet) return;
    window.requestAnimationFrame(() => {
      window.scrollTo({ top: 0, left: 0, behavior: 'auto' });
    });
  }, [selectedSheet]);

  useEffect(() => {
    if (!selectedSheet) return;
    setFormData((prev) => {
      const nextImpingers = normalizeImpingers(prev.impingers, getImpingerStageCount(selectedSheet));
      return { ...prev, impingers: nextImpingers };
    });
  }, [selectedSheet]);

  useEffect(() => {
    const refreshSkyPhase = () => {
      setSkyPhase(getSkyPhaseByHour());
    };
    refreshSkyPhase();
    const timer = window.setInterval(refreshSkyPhase, 60 * 1000);
    return () => window.clearInterval(timer);
  }, []);

  useEffect(() => {
    const updateViewportState = () => {
      setIsMobileViewport(isMobileViewportWidth());
    };
    updateViewportState();
    window.addEventListener('resize', updateViewportState);
    window.addEventListener('orientationchange', updateViewportState);
    return () => {
      window.removeEventListener('resize', updateViewportState);
      window.removeEventListener('orientationchange', updateViewportState);
    };
  }, []);

  if (!selectedSheet) {
    return (
      <div className={`relative min-h-screen overflow-hidden p-6 md:p-10 font-sans text-slate-800 ${menuSkyBgClass}`}>
        <style>{SKYLINE_FLYER_CSS}</style>
      <div className="pointer-events-none absolute inset-0 z-0">
        <div className={`absolute inset-0 ${menuSkyOverlayClass}`} />
        <div className={`absolute -left-20 -top-24 h-80 w-80 rounded-full blur-3xl ${menuSkyGlowLeftClass}`} />
        <div className={`absolute right-4 top-2 h-72 w-72 rounded-full blur-3xl ${menuSkyGlowRightClass}`} />
        {skyOrbClass && <div className={skyOrbClass} />}
        {renderSkyDecor('menu')}
          {isNightSky && (
            <div className="absolute inset-0">
              {sceneStarsMenu.map((star) => (
                <span
                  key={star.id}
                  className="skyline-star"
                  style={{
                    left: star.left,
                    top: star.top,
                    width: `${star.size}px`,
                    height: `${star.size}px`,
                    '--twinkle-duration': `${star.duration}s`,
                    '--twinkle-delay': `-${star.delay}s`,
                  }}
                />
              ))}
            </div>
          )}
          <div className="absolute bottom-0 left-0 right-0 h-[22rem] md:h-[32rem]">
            {renderSkylineSvg({ mode: resolvedSkyMode, scope: 'menu' })}
          </div>
        </div>

        <div className="relative z-30 max-w-6xl mx-auto">
          <div className="mb-6 p-2 text-center">
            <h1 className="text-slate-900 tracking-tight flex justify-center">
              <img
                src={mainLogoSrc}
                alt="STACKPILOT"
                onError={() => setLogoIndex((prev) => Math.min(prev + 1, logoCandidates.length - 1))}
                className="h-[2.6rem] md:h-[4rem] w-auto select-none"
              />
            </h1>
          </div>
          <div className="bg-white/85 backdrop-blur p-4 rounded-2xl shadow-lg border border-white/70 mb-6 relative">
            {!isUserUnlocked ? (
              <div className="relative overflow-hidden border border-emerald-200 rounded-2xl p-5 bg-gradient-to-br from-emerald-50 via-white to-emerald-100/70 shadow-inner text-center">
                <div className="pointer-events-none absolute -right-10 -top-10 h-28 w-28 rounded-full bg-emerald-200/70 blur-2xl" />
                <p className="text-sm font-black text-emerald-900">베타 접속</p>
                <p className="text-xs text-emerald-800 mt-1">회원가입 없이 기본 계정(user)으로 바로 사용합니다.</p>
                <div className="mt-3 flex flex-wrap items-center justify-center gap-2">
                  <button
                    type="button"
                    onClick={handleBetaQuickLogin}
                    className="px-4 py-2 bg-emerald-600 text-white rounded-lg font-bold text-sm hover:bg-emerald-700 shadow-sm"
                  >
                    user 로그인
                  </button>
                  <button
                    type="button"
                    onClick={openAdminLoginModal}
                    className="px-4 py-2 bg-slate-700 text-white rounded-lg font-bold text-sm hover:bg-slate-800 shadow-sm"
                  >
                    관리자 로그인
                  </button>
                </div>

                {authModal === 'admin' && typeof document !== 'undefined' && createPortal(
                  <div className="fixed inset-0 z-[9999] bg-black/40 p-4 flex items-center justify-center" onClick={() => setAuthModal('')}>
                    <div className="w-full max-w-md max-h-[85vh] overflow-y-auto bg-white rounded-2xl border border-slate-200 shadow-xl p-5" onClick={(e) => e.stopPropagation()}>
                      <div className="flex items-center justify-between mb-4">
                        <h3 className="text-base font-black text-slate-900">관리자 로그인</h3>
                        <button type="button" onClick={() => setAuthModal('')} className="px-2 py-1 rounded hover:bg-slate-100 text-slate-600 text-xs font-bold">
                          닫기
                        </button>
                      </div>
                      <div className="space-y-2">
                        <input
                          type="text"
                          value={loginUserId}
                          onChange={(e) => setLoginUserId(e.target.value)}
                          className="w-full p-2 border border-slate-300 rounded-lg bg-white outline-none focus:ring-2 focus:ring-emerald-500"
                          placeholder="아이디"
                        />
                        <input
                          type="password"
                          value={loginPassword}
                          onChange={(e) => setLoginPassword(e.target.value)}
                          onKeyDown={(e) => {
                            if (e.key === 'Enter') {
                              e.preventDefault();
                              handleLoginUser();
                            }
                          }}
                          className="w-full p-2 border border-slate-300 rounded-lg bg-white outline-none focus:ring-2 focus:ring-emerald-500"
                          placeholder="비밀번호"
                        />
                        <button
                          type="button"
                          onClick={handleLoginUser}
                          className="w-full py-2 bg-slate-800 text-white rounded-lg font-bold text-sm hover:bg-slate-900"
                        >
                          로그인
                        </button>
                      </div>
                    </div>
                  </div>,
                  document.body
                )}
              </div>
            ) : (
              <div className="space-y-4 flex flex-col items-center">
                <div className="w-full max-w-3xl border border-slate-200 rounded-xl p-4 bg-white relative shadow-sm">
                  <button
                    type="button"
                    aria-label="설정 열기"
                    onClick={() => setAuthModal('settings')}
                    className="absolute right-3 top-3 px-2 py-1 rounded border border-slate-300 bg-white text-slate-700 hover:bg-slate-100 font-bold text-xs z-10"
                  >
                    ⚙ 설정
                  </button>
                  <div className="flex flex-col items-center md:flex-row md:items-center gap-4 justify-center">
                    <div className="shrink-0 flex flex-col items-center gap-2">
                      <input ref={avatarFileInputRef} type="file" accept="image/*" className="hidden" onChange={handleProfileAvatarFile} />
                      <button
                        type="button"
                        onClick={() => avatarFileInputRef.current && avatarFileInputRef.current.click()}
                        title="프로필 사진 바꾸기"
                        aria-label="프로필 사진 바꾸기"
                        className="group relative w-24 h-24 rounded-xl border border-slate-300 bg-white shadow-sm overflow-hidden flex items-center justify-center cursor-pointer p-0"
                      >
                        {profileAvatarUrl ? (
                          <img src={profileAvatarUrl} alt={`${displayUserNameWithFallback} 프로필`} className="w-full h-full object-cover" />
                        ) : (
                          <span className="menu-avatar-fallback w-full h-full flex items-center justify-center bg-emerald-600 text-white text-4xl font-black select-none">
                            {String(displayUserNameWithFallback || '').trim().charAt(0).toUpperCase() || '?'}
                          </span>
                        )}
                        <span className="absolute inset-x-0 bottom-0 bg-black/55 text-white text-[11px] font-bold py-1 opacity-0 group-hover:opacity-100 group-focus-visible:opacity-100 transition-opacity">사진 바꾸기</span>
                      </button>
                    </div>
                    <div className="flex-1 text-center md:text-left">
                      <p className="text-sm text-slate-600">현재 사용자</p>
                      <p className="text-xl font-black text-slate-900">{displayUserNameWithFallback}</p>
                      <p className="text-xs text-slate-600 mt-1">
                        상태: <span className="font-black text-emerald-700">로그인됨</span>
                        {' '}| 저장 건수: <span className="font-black text-emerald-700">{savedData.length}</span>건
                      </p>
                      <div className="mt-3 flex flex-wrap items-center justify-center md:justify-start gap-2">
                        <button
                          type="button"
                          onClick={() => setAuthModal('nickname')}
                          className="px-3 py-1.5 rounded border border-slate-300 bg-white text-slate-700 hover:bg-slate-100 font-bold text-xs"
                        >
                          별명 설정
                        </button>
                        <button
                          type="button"
                          onClick={() => avatarFileInputRef.current && avatarFileInputRef.current.click()}
                          className="px-3 py-1.5 rounded border border-slate-300 bg-white text-slate-700 hover:bg-slate-100 font-bold text-xs"
                        >
                          사진 변경
                        </button>
                        {profileAvatarUrl && (
                          <button
                            type="button"
                            onClick={handleRemoveProfileAvatar}
                            className="px-3 py-1.5 rounded border border-slate-300 bg-white text-slate-700 hover:bg-slate-100 font-bold text-xs"
                          >
                            사진 지우기
                          </button>
                        )}
                        <button
                          type="button"
                          onClick={handleLockActiveUser}
                          className="px-3 py-1.5 rounded bg-slate-700 text-white hover:bg-slate-800 font-bold text-xs"
                        >
                          로그아웃
                        </button>
                      </div>
                    </div>
                  </div>
                </div>

                {authModal === 'nickname' && typeof document !== 'undefined' && createPortal(
                  <div className="fixed inset-0 z-[9999] bg-black/40 p-4 flex items-center justify-center" onClick={() => setAuthModal('')}>
                    <div className="w-full max-w-lg max-h-[85vh] overflow-y-auto bg-white rounded-2xl border border-slate-200 shadow-xl p-5" onClick={(e) => e.stopPropagation()}>
                      <div className="flex items-center justify-between mb-4">
                        <h3 className="text-base font-black text-slate-900">별명 설정</h3>
                        <button type="button" onClick={() => setAuthModal('')} className="px-2 py-1 rounded hover:bg-slate-100 text-slate-600 text-xs font-bold">
                          닫기
                        </button>
                      </div>
                      <div className="flex gap-2">
                        <input
                          type="text"
                          value={profileNickname}
                          onChange={(e) => setProfileNickname(e.target.value)}
                          className="flex-1 p-2 border border-slate-300 rounded-lg bg-white outline-none focus:ring-2 focus:ring-emerald-500"
                          placeholder="별명 입력 (최대 24자)"
                          maxLength={24}
                        />
                        <button
                          type="button"
                          onClick={() => {
                            const saved = handleSaveProfileNickname();
                            if (saved) setAuthModal('');
                          }}
                          className="px-4 py-2 bg-emerald-600 text-white rounded-lg font-bold text-sm hover:bg-emerald-700"
                        >
                          저장
                        </button>
                      </div>
                    </div>
                  </div>,
                  document.body
                )}

                {authModal === 'settings' && typeof document !== 'undefined' && createPortal(
                  <div className="fixed inset-0 z-[9999] bg-black/40 p-4 flex items-center justify-center" onClick={() => setAuthModal('')}>
                    <div className="w-full max-w-lg max-h-[85vh] overflow-y-auto bg-white rounded-2xl border border-slate-200 shadow-xl p-5" onClick={(e) => e.stopPropagation()}>
                      <div className="flex items-center justify-between mb-4">
                        <h3 className="text-base font-black text-slate-900">설정</h3>
                        <button type="button" onClick={() => setAuthModal('')} className="px-2 py-1 rounded hover:bg-slate-100 text-slate-600 text-xs font-bold">
                          닫기
                        </button>
                      </div>
                      <div className="mb-3 p-3 rounded-lg border border-slate-200 bg-slate-50">
                        <p className="text-xs font-black text-slate-800">색 테마</p>
                        <p className="text-[11px] text-slate-600 mt-1">열려 있는 모든 창에 바로 적용됩니다.</p>
                        <div className="mt-2 grid grid-cols-2 gap-2">
                          {[
                            { id: 'default', label: '기본', swatches: ['#059669', '#f8fafc', '#1e293b'] },
                            { id: 'earth', label: '어스 (세이지·베이지·테라코타)', swatches: ['#9caf88', '#e3d9c6', '#c8643c'] },
                          ].map((item) => (
                            <button
                              key={item.id}
                              type="button"
                              aria-pressed={palette === item.id}
                              onClick={() => changePalette(item.id)}
                              className={`px-3 py-2 rounded border text-left text-xs font-bold ${palette === item.id ? 'border-slate-800 bg-white text-slate-900 ring-2 ring-slate-800' : 'border-slate-300 bg-white text-slate-700 hover:bg-slate-100'}`}
                            >
                              <span className="flex items-center gap-1 mb-1">
                                {item.swatches.map((color) => (
                                  <span key={color} className="palette-swatch inline-block w-4 h-4 rounded-full border border-black/10" style={{ backgroundColor: color }} />
                                ))}
                              </span>
                              {item.label}{palette === item.id ? ' ✓' : ''}
                            </button>
                          ))}
                        </div>
                      </div>
                      {canSetExportBase && (
                        <div className="mb-3 p-3 rounded-lg border border-slate-200 bg-slate-50">
                          <p className="text-xs font-black text-slate-800">엑셀 저장 위치</p>
                          <p className="text-[11px] text-slate-600 mt-1">이 폴더 아래에 "날짜_사업장명_#배출구" 폴더를 만들어 엑셀을 저장합니다.</p>
                          <p className="mt-2 px-2 py-1.5 rounded border border-slate-300 bg-white text-[11px] text-slate-800 break-all" title={exportBase}>{exportBase || '불러오는 중…'}</p>
                          <div className="mt-2 flex gap-2">
                            <button type="button" onClick={chooseExportBase} className="flex-1 px-3 py-2 rounded border border-slate-300 bg-white text-slate-700 hover:bg-slate-100 font-bold text-xs">폴더 선택</button>
                            <button type="button" onClick={resetExportBase} className="flex-1 px-3 py-2 rounded border border-slate-300 bg-white text-slate-700 hover:bg-slate-100 font-bold text-xs">다운로드 폴더로 되돌리기</button>
                          </div>
                        </div>
                      )}
                      <div className="space-y-2">
                        <button
                          type="button"
                          onClick={clearActiveUserReports}
                          className="w-full px-3 py-2 rounded border border-red-300 text-red-700 hover:bg-red-50 font-bold text-sm"
                        >
                          리포트 전체 삭제
                        </button>
                        <button
                          type="button"
                          onClick={handleDeleteActiveUser}
                          className="w-full px-3 py-2 rounded bg-red-600 text-white hover:bg-red-700 font-bold text-sm"
                        >
                          현재 계정 삭제
                        </button>
                        <div className="mt-3 p-3 rounded-lg border border-slate-200 bg-slate-50">
                          <p className="text-xs font-black text-slate-800">장비 프리셋</p>
                          <p className="text-[11px] text-slate-600 mt-1">노즐·샘플러·가스 샘플러·피토관 프리셋을 엑셀 파일로 뽑거나, 이 PC에 저장된 프리셋을 모두 지웁니다.</p>
                          <button
                            type="button"
                            onClick={resetAllEquipmentPresets}
                            className="mt-2 w-full px-3 py-2 rounded border border-amber-400 bg-amber-50 text-amber-800 hover:bg-amber-100 font-bold text-sm"
                          >
                            장비 프리셋 초기화
                          </button>
                          <button type="button" onClick={exportPresetsExcel} className="mt-2 w-full px-3 py-2 rounded border border-emerald-300 bg-white text-emerald-700 hover:bg-emerald-50 font-bold text-sm">프리셋 엑셀 추출</button>
                          <details className="mt-2">
                            <summary className="cursor-pointer text-[11px] font-bold text-slate-700">교정 이력 {presetHistory.length}건</summary>
                            <p className="text-[11px] text-slate-500 mt-1">샘플러·가스 샘플러·피토관의 교정일이나 교정값을 바꿔 저장할 때마다 한 줄씩 남습니다. 엑셀 추출 파일의 "교정 이력" 시트에도 들어갑니다.</p>
                            {presetHistory.length > 0 && (
                              <div className="mt-2 max-h-56 overflow-auto border border-slate-200 rounded bg-white">
                                <table className="w-full text-[11px] text-center">
                                  <thead className="bg-slate-100 text-slate-700 sticky top-0"><tr><th className="p-1">저장일</th><th className="p-1">구분</th><th className="p-1">장비</th><th className="p-1">교정일</th><th className="p-1">교정값</th></tr></thead>
                                  <tbody>
                                    {[...presetHistory].reverse().map((row, i) => (
                                      <tr key={i} className="border-t border-slate-100">
                                        <td className="p-1 whitespace-nowrap">{String(row.at || '').slice(0, 10)}</td>
                                        <td className="p-1 whitespace-nowrap">{PRESET_HISTORY_KINDS[row.kind]?.label || row.kind}</td>
                                        <td className="p-1">{row.name}</td>
                                        <td className="p-1 whitespace-nowrap">{row.calDate || '-'}</td>
                                        <td className="p-1 text-left">{row.detail}{row.note ? ` (${row.note})` : ''}</td>
                                      </tr>
                                    ))}
                                  </tbody>
                                </table>
                              </div>
                            )}
                          </details>
                        </div>
                        <div className="mt-3 p-3 rounded-lg border border-red-200 bg-red-50">
                          <p className="text-xs font-black text-red-800">전체 초기화</p>
                          <p className="text-[11px] text-red-700 mt-1">이 PC에 저장된 모든 계정, 리포트, 프리셋, 설정을 지우고 처음 상태로 되돌립니다.</p>
                          <button
                            type="button"
                            onClick={resetAllAppData}
                            className="mt-2 w-full px-3 py-2 rounded bg-red-700 text-white hover:bg-red-800 font-bold text-sm"
                          >
                            전체 초기화 (모든 데이터 삭제)
                          </button>
                        </div>
                        <div className="mt-3 p-3 rounded-lg border border-slate-200 bg-slate-50">
                          <p className="text-xs font-black text-slate-800">앱 설정</p>
                          <p className="text-[11px] text-slate-600 mt-1">배경 시간대(아침/낮/밤) 표시를 선택합니다.</p>
                          <div className="mt-2 flex flex-wrap items-center gap-2">
                            {SKY_PREVIEW_OPTIONS.map((opt) => {
                              const selected = skyPreviewMode === opt.id;
                              return (
                                <button
                                  key={`settings-sky-${opt.id}`}
                                  type="button"
                                  onClick={() => setSkyPreviewMode(opt.id)}
                                  className={`px-2.5 py-1 rounded border text-[11px] font-bold transition-colors ${
                                    selected
                                      ? 'border-emerald-600 bg-emerald-600 text-white'
                                      : 'border-slate-300 bg-white text-slate-700 hover:bg-slate-100'
                                  }`}
                                >
                                  {opt.label}
                                </button>
                              );
                            })}
                          </div>
                        </div>
                        {!isDesktopApp && (
                          <div className="mt-3 p-3 rounded-lg border border-slate-200 bg-slate-50">
                            <p className="text-xs font-black text-slate-800">날씨 중계 주소 (모바일)</p>
                            <p className="text-[11px] text-slate-600 mt-1">모바일에서 기상청 자료 가져오기·장소 검색을 쓰려면 중계 주소가 필요합니다. 만드는 방법은 저장소의 relay/README.md에 있습니다.</p>
                            <div className="mt-2 flex gap-2">
                              <input
                                type="url"
                                value={weatherRelayDraft === null ? weatherRelay : weatherRelayDraft}
                                onChange={(e) => setWeatherRelayDraft(e.target.value)}
                                placeholder="https://stackpilot-weather.○○○.workers.dev"
                                className="flex-1 min-w-0 p-2 border border-slate-300 rounded-lg bg-white outline-none text-xs"
                              />
                              <button type="button" onClick={() => saveWeatherRelay(weatherRelayDraft === null ? weatherRelay : weatherRelayDraft)} className="px-3 py-2 rounded border border-slate-300 bg-white text-slate-700 hover:bg-slate-100 font-bold text-[11px]">저장</button>
                            </div>
                            <p className="text-[11px] mt-1 text-slate-500">{weatherRelay ? '저장됨 · 날씨 가져오기를 쓸 수 있습니다' : '아직 설정하지 않았습니다'}</p>
                          </div>
                        )}
                        <div className="mt-3 p-3 rounded-lg border border-slate-200 bg-slate-50">
                          <p className="text-xs font-black text-slate-800">백업 파일</p>
                          <p className="text-[11px] text-slate-600 mt-1">기록과 장비 프리셋을 파일 하나로 저장합니다. 휴대폰에서는 '파일' 앱(iCloud Drive)에 저장해 두면, 사파리 데이터를 지워도 불러와서 되살릴 수 있습니다.</p>
                          <input ref={backupFileInputRef} type="file" accept="application/json,.json" className="hidden" onChange={importBackupFile} />
                          <div className="mt-2 grid grid-cols-2 gap-2">
                            <button type="button" onClick={exportBackupFile} className="px-2 py-2 rounded border border-emerald-300 text-emerald-700 hover:bg-emerald-50 font-bold text-[11px]">백업 파일 저장</button>
                            <button type="button" onClick={() => backupFileInputRef.current && backupFileInputRef.current.click()} className="px-2 py-2 rounded border border-blue-300 text-blue-700 hover:bg-blue-50 font-bold text-[11px]">백업 파일 불러오기</button>
                          </div>
                        </div>
                        <div className="mt-3 p-3 rounded-lg border border-slate-200 bg-slate-50">
                          <p className="text-xs font-black text-slate-800">클라우드 동기화 (GitHub Gist) · 자동 백업</p>
                          <p className="text-[11px] text-slate-600 mt-1">한 번 설정하면 저장할 때마다 자동으로 깃허브에 백업됩니다. 데이터를 지운 뒤에는 같은 PAT와 Gist ID를 다시 넣고 "클라우드 복원"을 누르세요. Gist ID는 따로 적어 두세요.</p>
                          <div className="mt-2 space-y-2">
                            <input
                              type="password"
                              value={cloudSyncConfig.token}
                              onChange={(e) => setCloudSyncConfig((prev) => ({ ...prev, token: e.target.value }))}
                              className="w-full p-2 border border-slate-300 rounded-lg bg-white outline-none focus:ring-2 focus:ring-emerald-500 text-xs"
                              placeholder="GitHub PAT (gist 권한)"
                            />
                            <input
                              type="text"
                              value={cloudSyncConfig.gistId}
                              onChange={(e) => setCloudSyncConfig((prev) => ({ ...prev, gistId: e.target.value }))}
                              className="w-full p-2 border border-slate-300 rounded-lg bg-white outline-none focus:ring-2 focus:ring-emerald-500 text-xs"
                              placeholder="Gist ID (처음 업로드 시 자동 생성)"
                            />
                            <div className="grid grid-cols-3 gap-2">
                              <button
                                type="button"
                                onClick={handleSaveCloudSyncConfig}
                                disabled={cloudSyncBusy}
                                className="px-2 py-2 rounded border border-slate-300 text-slate-700 hover:bg-slate-100 font-bold text-[11px] disabled:opacity-60"
                              >
                                설정 저장
                              </button>
                              <button
                                type="button"
                                onClick={handleCloudUploadToGist}
                                disabled={cloudSyncBusy}
                                className="px-2 py-2 rounded border border-emerald-300 text-emerald-700 hover:bg-emerald-50 font-bold text-[11px] disabled:opacity-60"
                              >
                                {cloudSyncBusy ? '진행 중' : '클라우드 업로드'}
                              </button>
                              <button
                                type="button"
                                onClick={handleCloudRestoreFromGist}
                                disabled={cloudSyncBusy}
                                className="px-2 py-2 rounded border border-blue-300 text-blue-700 hover:bg-blue-50 font-bold text-[11px] disabled:opacity-60"
                              >
                                {cloudSyncBusy ? '진행 중' : '클라우드 복원'}
                              </button>
                            </div>
                          </div>
                        </div>
                      </div>
                    </div>
                  </div>,
                  document.body
                )}
              </div>
            )}
          </div>
          {isUserUnlocked && (
            <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-5">
              {SHEET_MENU.map((sheet) => {
                return (
                  <button
                    key={sheet.id}
                    type="button"
                    onClick={() => navigateToSheet(sheet.id)}
                    className="group relative overflow-hidden text-left bg-white border border-slate-200 rounded-2xl p-5 shadow-sm transition-all hover:shadow-lg hover:-translate-y-0.5"
                  >
                    <div className={`absolute -right-8 -top-8 w-28 h-28 rounded-full bg-gradient-to-br ${sheet.toneClass} blur-sm group-hover:scale-110 transition-transform`} />
                    <h2 className="text-base font-black text-slate-900">{sheet.title}</h2>
                    <p className="text-xs text-slate-600 mt-1">{sheet.desc}</p>
                    <div className={`mt-4 flex items-center text-[11px] font-bold text-slate-500 transition-colors ${sheet.hoverTextClass}`}>
                      새 창으로 열기
                    </div>
                  </button>
                );
              })}
            </div>
          )}

          {isUserUnlocked && (
            <div className="mt-6 bg-white p-4 rounded-xl shadow-sm border border-slate-200">
              <div className="flex items-center justify-between gap-2 mb-3">
                <h3 className="text-sm font-black text-slate-900">저장된 리포트 종합 (먼지/금속/수은/PAHs/유량/통합시료채취)</h3>
                <span className="text-xs font-bold text-slate-600">총 {profileCombinedReports.length}건</span>
              </div>

              <div className="grid grid-cols-2 md:grid-cols-4 gap-2 mb-3">
                {SHEET_MENU.map((sheet) => (
                  <div key={sheet.id} className="p-2 rounded-lg border border-slate-200 bg-slate-50 text-center">
                    <p className="text-[11px] font-bold text-slate-600">{sheet.title}</p>
                    <p className="text-lg font-black text-slate-900">{profileSheetCounts[sheet.id] || 0}</p>
                  </div>
                ))}
              </div>

              {profileCombinedReports.length === 0 ? (
                <div className="p-5 text-center text-xs text-slate-500 border border-dashed border-slate-300 rounded-lg">
                  저장된 통합 리포트가 없습니다.
                </div>
              ) : (
                <div>
                  <div className="flex flex-wrap items-center justify-between gap-2 mb-3 p-2 rounded-lg border border-slate-200 bg-slate-50">
                    <p className="text-[11px] font-bold text-slate-700">
                      체크 선택: {menuCheckedReportKeysValid.length}건
                    </p>
                    <div className="flex flex-wrap gap-2">
                      <label
                        htmlFor="profile-csv-import"
                        className="px-2 py-1 rounded border border-emerald-300 text-emerald-700 hover:bg-emerald-50 font-bold text-[11px] cursor-pointer"
                      >
                        선택 항목 엑셀로 추출
                      </label>
                      <input
                        id="profile-csv-import"
                        ref={csvImportInputRef}
                        type="file"
                        accept=".csv,text/csv"
                        onChange={handleImportCsvReports}
                        className="hidden"
                      />
                      <button
                        type="button"
                        onClick={handleDeleteCheckedMenuReports}
                        className="px-2 py-1 rounded border border-red-300 text-red-700 hover:bg-red-50 font-bold text-[11px]"
                      >
                        선택 삭제
                      </button>
                      <button
                        type="button"
                        onClick={() => setMenuCheckedReportKeys([])}
                        className="px-2 py-1 rounded border border-slate-300 text-slate-700 hover:bg-slate-100 font-bold text-[11px]"
                      >
                        선택 해제
                      </button>
                    </div>
                  </div>
                  {renderReportGroupTable({
                    groups: profileReportGroups,
                    checkedSet: menuCheckedReportKeySet,
                    toggleCheck: toggleMenuReportCheck,
                    allChecked: isAllMenuReportsChecked,
                    toggleAll: toggleMenuAllReportChecks,
                  })}
                </div>
              )}
            </div>
          )}
        </div>
        <SignatureBadge />
      </div>
    );
  }

  // 통합 시료채취기록부 본문 (전체 화면과 기록부 위 패널에서 공용)
  const renderSamplingLogBody = () => {
    const log = samplingLogData;
    const logForCalc = resolveSamplingLogYds(log);
    const logStats = getSamplingLogStats(log);
    const savedLogs = savedData.filter((item) => resolveReportSheetId(item) === 'samplinglog');
    const setGroupRequested = (group, requested) => setSamplingLogData(prev => ({
      ...prev,
      rows: Object.fromEntries(Object.entries(prev.rows).map(([key, row]) => {
        const sub = SAMPLING_LOG_SUBSTANCES.find((item) => String(item.row) === key);
        return [key, sub && group.kinds.includes(sub.kind) ? { ...row, requested } : row];
      })),
    }));
    const logInput = 'p-1.5 border border-slate-300 rounded-lg bg-white text-sm outline-none focus:ring-2 focus:ring-orange-500';
    const cellInput = 'w-full p-1 border border-slate-200 rounded text-xs text-center outline-none focus:ring-2 focus:ring-orange-400 bg-white disabled:bg-slate-50 disabled:text-slate-400';
    return (
      <>

          <div className="bg-white p-3 rounded-xl shadow-sm border border-slate-200 flex flex-col md:flex-row md:items-center md:justify-between gap-3">
            <div className="flex flex-wrap items-center gap-2 text-xs text-slate-600">
              <span>사용자 <span className="font-black text-slate-900">{displayUserNameWithFallback}</span></span>
              <span className={`px-2 py-0.5 rounded-full border text-[11px] font-bold ${isUserUnlocked ? 'bg-orange-50 border-orange-200 text-orange-700' : 'bg-red-50 border-red-200 text-red-700'}`}>{isUserUnlocked ? '로그인됨' : '로그인 필요'}</span>
              <span className="px-2 py-0.5 rounded-full bg-slate-100 border border-slate-200 text-[11px] font-bold text-slate-700">의뢰 {logStats.requestedCount}종 / {logStats.speciesTotal}종</span>
              <span className={`px-2 py-0.5 rounded-full border text-[11px] font-bold ${logStats.requestedRows > 0 && logStats.filledCount === logStats.requestedRows ? 'bg-emerald-50 border-emerald-200 text-emerald-700' : 'bg-amber-50 border-amber-200 text-amber-800'}`}>시각 입력 {logStats.filledCount}/{logStats.requestedRows}</span>
              {log.company && <span className="px-2 py-0.5 rounded-full bg-slate-100 border border-slate-200 text-[11px] font-bold text-slate-700">{log.company}{log.stackName ? ` · ${log.stackName}` : ''}</span>}
            </div>
          </div>
          <SheetSectionNav scope="samplinglog" />

          <div className="bg-white p-6 rounded-xl shadow-sm border border-slate-200">
            <h2 className="ds-sec-title">
              <span className="ds-sec-num">1</span>
              사업장 및 장비
            </h2>
            <div className="grid grid-cols-1 md:grid-cols-5 gap-3">
              <label className="text-xs font-bold text-slate-700 md:col-span-2">사업장명
                <input type="text" name="log-company" value={log.company} onChange={(e) => updateSamplingLogField('company', e.target.value)} className={`${logInput} w-full mt-1`} />
              </label>
              <label className="text-xs font-bold text-slate-700">업종
                <input type="text" name="log-industry" value={log.industry} onChange={(e) => updateSamplingLogField('industry', e.target.value)} className={`${logInput} w-full mt-1`} placeholder="예: 화력발전업(35113)" />
              </label>
              <label className="text-xs font-bold text-slate-700">배출시설
                <input type="text" name="log-facility" value={log.facility} onChange={(e) => updateSamplingLogField('facility', e.target.value)} className={`${logInput} w-full mt-1`} />
              </label>
              <label className="text-xs font-bold text-slate-700">굴뚝명
                <input type="text" name="log-stack" value={log.stackName} onChange={(e) => updateSamplingLogField('stackName', e.target.value)} className={`${logInput} w-full mt-1`} placeholder="예: #A-1" />
              </label>
              <label className="text-xs font-bold text-slate-700">시설약자 (시료번호용)
                <input type="text" name="log-abbr" value={log.siteAbbr} onChange={(e) => updateSamplingLogField('siteAbbr', e.target.value)} className={`${logInput} w-full mt-1`} placeholder="예: GSP" />
              </label>
              <label className="text-xs font-bold text-slate-700 md:col-span-2">입자 샘플러
                <div className="flex gap-2 mt-1">
                  <select name="log-particle-sampler" value={log.particleSamplerId} onChange={(e) => handleSamplingLogParticleSampler(e.target.value)} className={`${logInput} flex-1`}>
                    <option value="">선택</option>
                    {samplers.map((item) => <option key={item.id} value={item.id}>{item.name}</option>)}
                  </select>
                  <input type="number" step="0.0001" value={log.particleSamplerYd} onChange={(e) => setSamplingLogData(prev => ({ ...prev, particleSamplerYd: e.target.value, particleSamplerYdManual: e.target.value.trim() !== '' }))} className={`${logInput} w-24`} placeholder="Yd" title="입자 샘플러 Yd" />
                </div>
              </label>
              <label className="text-xs font-bold text-slate-700 md:col-span-2">가스 샘플러
                <div className="flex gap-2 mt-1">
                  <select name="log-gas-sampler" value={log.gasSamplerId} onChange={(e) => handleSamplingLogGasSampler(e.target.value)} className={`${logInput} flex-1`}>
                    <option value="">선택</option>
                    {gasSamplers.map((item) => <option key={item.id} value={item.id}>{`가스샘플러${item.id}${item.sn ? ` (S/N ${item.sn})` : ''}`}</option>)}
                  </select>
                  <input type="number" step="0.0001" value={log.gasSamplerYd} onChange={(e) => setSamplingLogData(prev => ({ ...prev, gasSamplerYd: e.target.value, gasSamplerYdManual: e.target.value.trim() !== '' }))} className={`${logInput} w-24`} placeholder="Yd" title="가스 샘플러 Yd" />
                </div>
              </label>
              <details className="md:col-span-4 border border-slate-200 rounded-lg p-2 bg-slate-50/60">
                <summary className="cursor-pointer text-[11px] font-bold text-slate-700">
                  가스 샘플러 프리셋 편집 (번호 · S/N · Yd)
                </summary>
                <div className="mt-2 flex items-center gap-2">
                  <button type="button" onClick={saveGasSamplerPresets} className="px-2 py-1 text-[11px] font-bold bg-orange-600 text-white rounded border border-orange-700 inline-flex items-center gap-1">
                    <PresetSaveIcon />
                    가스 샘플러 저장
                  </button>
                  <button type="button" onClick={addGasSamplerPreset} className="px-2 py-1 text-[11px] font-bold bg-white text-slate-700 rounded border border-slate-300">행 추가</button>
                  <button type="button" onClick={resetGasSamplerDefaults} className="px-2 py-1 text-[11px] font-bold bg-slate-100 text-slate-700 rounded border border-slate-300">{IS_DESKTOP_BUILD ? '비우기' : '초기값 복원'}</button>
                  <span className="text-[10px] text-slate-500">양식에는 앞의 5대가 들어갑니다</span>
                </div>
                <div className="mt-2 overflow-x-auto border border-slate-200 rounded-lg bg-white">
                  <table className="w-full text-[11px] text-center">
                    <thead className="bg-slate-100 text-slate-700">
                      <tr>
                        <th className="p-1 w-16">번호</th>
                        <th className="p-1">S/N</th>
                        <th className="p-1 w-32">Yd</th>
                        <th className="p-1 w-36">교정일</th>
                        <th className="p-1 w-14">삭제</th>
                      </tr>
                    </thead>
                    <tbody>
                      {gasSamplers.length === 0 && (
                        <tr><td colSpan="5" className="p-3 text-slate-400">가스 샘플러가 없습니다. 행 추가로 입력하세요.</td></tr>
                      )}
                      {gasSamplers.map((item, idx) => (
                        <tr key={`${item.id}-${idx}`}>
                          <td className="p-1 font-bold text-slate-700">#{item.id}</td>
                          <td className="p-1"><input type="text" value={item.sn} onChange={(e) => handleGasSamplerPresetChange(idx, 'sn', e.target.value)} className="w-full p-1 border border-slate-300 rounded text-center" placeholder="S/N" /></td>
                          <td className="p-1"><input type="number" step="0.0001" value={item.yd} onChange={(e) => handleGasSamplerPresetChange(idx, 'yd', e.target.value)} className="w-full p-1 border border-slate-300 rounded text-center" placeholder="Yd" /></td>
                          <td className="p-1"><input type="date" value={item.calDate || ''} onChange={(e) => handleGasSamplerPresetChange(idx, 'calDate', e.target.value)} className="w-full p-1 border border-slate-300 rounded text-center" title="검교정일" /></td>
                          <td className="p-1"><button type="button" onClick={() => removeGasSamplerPreset(idx)} className="text-red-600 font-bold text-[11px]">삭제</button></td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </details>
            </div>
          </div>

          {(() => {
            const groups = groupSheetReportsForLog(activeUserReports, resolveReportSheetId);
            if (groups.length === 0) return null;
            return (
              <div className="bg-white p-4 rounded-xl shadow-sm border border-slate-200">
                <div className="flex flex-wrap items-center justify-between gap-2 mb-2">
                  <h3 className="text-sm font-bold text-slate-800">저장된 기록부에서 가져오기 <span className="text-xs font-normal text-slate-500">같은 사업장·스택, 3일 이내 날짜끼리 한 묶음</span></h3>
                  
                </div>
                <div className="overflow-x-auto border border-slate-200 rounded-lg">
                  <table className="w-full text-xs text-left">
                    <thead className="bg-slate-100 text-slate-700">
                      <tr>
                        <th className="p-2">측정일자</th>
                        <th className="p-2">사업장</th>
                        <th className="p-2">배출구(스택)</th>
                        <th className="p-2">저장된 기록부</th>
                        <th className="p-2 w-28"></th>
                      </tr>
                    </thead>
                    <tbody>
                      {groups.slice(0, 12).map((group) => (
                        <tr key={group.key}>
                          <td className="p-2 whitespace-nowrap">{group.dateLabel}</td>
                          <td className="p-2 font-bold text-slate-900">{group.company || '-'}</td>
                          <td className="p-2">{group.location || '-'}</td>
                          <td className="p-2">
                            {LOG_SOURCE_SHEETS.map((sheetId) => (
                              <span key={sheetId} className={`inline-block mr-1 px-1.5 py-0.5 rounded border text-[10px] font-bold ${group.bySheet[sheetId] ? 'border-orange-300 bg-orange-50 text-orange-800' : 'border-slate-200 text-slate-300'}`}>
                                {{ dust: '먼지', metal: '중금속', mercury: '수은', pahs: 'PAHs', flow: '유량' }[sheetId]}
                              </span>
                            ))}
                          </td>
                          <td className="p-2 text-right">
                            <button type="button" onClick={() => applySheetGroupToSamplingLog(group)} className="px-2.5 py-1 text-[11px] font-bold rounded bg-orange-600 text-white border border-orange-700 hover:bg-orange-700">통합시료채취기록부에 옮기기</button>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            );
          })()}

          <div className="bg-white p-6 rounded-xl shadow-sm border border-slate-200">
            <h2 className="ds-sec-title">
              <span className="ds-sec-num">2</span>
              채취일 및 현장기상 <span className="text-xs font-normal text-slate-500">(최대 4일)</span>
            </h2>
            <div className="mb-2 flex flex-wrap items-center gap-2 text-xs text-slate-600">
              <span className="font-bold">날씨 위치</span>
              {weatherStationSelect}
              <span className="text-slate-500">채취일을 적고 오른쪽 "가져오기"를 누르면 기상·습도·풍향·풍속을 채웁니다 (기온·기압은 현장 측정값)</span>
            </div>
            <div className="overflow-x-auto">
              <table className="w-full text-xs text-center border-collapse">
                <thead>
                  <tr className="bg-orange-50 text-slate-700">
                    <th className="p-2 border border-slate-200">일차</th><th className="p-2 border border-slate-200">채취일</th><th className="p-2 border border-slate-200">기상</th>
                    <th className="p-2 border border-slate-200">기온(℃)</th><th className="p-2 border border-slate-200">습도(%)</th><th className="p-2 border border-slate-200">기압(mmHg)</th>
                    <th className="p-2 border border-slate-200">풍향</th><th className="p-2 border border-slate-200">풍속(m/s)</th>
                    <th className="p-2 border border-slate-200">기상청</th>
                  </tr>
                </thead>
                <tbody>
                  {log.days.map((day, idx) => (
                    <tr key={`log-day-${idx}`}>
                      <td className="p-1 border border-slate-200 font-bold">{idx + 1}일차</td>
                      <td className="p-1 border border-slate-200"><input type="date" name={`log-day-${idx}-date`} value={day.date} onChange={(e) => updateSamplingLogDay(idx, 'date', e.target.value)} className={cellInput} /></td>
                      <td className="p-1 border border-slate-200"><input type="text" value={day.weather} onChange={(e) => updateSamplingLogDay(idx, 'weather', e.target.value)} className={cellInput} placeholder="맑음" /></td>
                      <td className="p-1 border border-slate-200"><input type="number" step="0.1" value={day.temp} onChange={(e) => updateSamplingLogDay(idx, 'temp', e.target.value)} className={cellInput} /></td>
                      <td className="p-1 border border-slate-200"><input type="number" step="1" value={day.humidity} onChange={(e) => updateSamplingLogDay(idx, 'humidity', e.target.value)} className={cellInput} /></td>
                      <td className="p-1 border border-slate-200"><input type="number" step="1" value={day.pressure} onChange={(e) => updateSamplingLogDay(idx, 'pressure', e.target.value)} className={cellInput} /></td>
                      <td className="p-1 border border-slate-200"><input type="text" value={day.windDir} onChange={(e) => updateSamplingLogDay(idx, 'windDir', e.target.value)} className={cellInput} placeholder="북서" /></td>
                      <td className="p-1 border border-slate-200"><input type="number" step="0.1" value={day.windSpeed} onChange={(e) => updateSamplingLogDay(idx, 'windSpeed', e.target.value)} className={cellInput} /></td>
                      <td className="p-1 border border-slate-200">
                        <button type="button" onClick={() => applyKmaWeatherToLogDay(idx)} disabled={weatherBusy || !day.date} className="px-2 py-1 rounded border border-slate-300 bg-white text-slate-700 hover:bg-slate-100 font-bold text-[11px] whitespace-nowrap disabled:opacity-40">가져오기</button>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>

          <div className="bg-white p-6 rounded-xl shadow-sm border border-slate-200">
            <h2 className="ds-sec-title">
              <span className="ds-sec-num">3</span>
              물질별 채취 기록 <span className="text-xs font-normal text-slate-500">(의뢰 표시한 물질만 추출)</span>
              <span className="ml-auto flex items-center gap-1">
                <button type="button" onClick={() => setSamplingLogRequestedOnly(false)} className={`px-2.5 py-1 rounded-full text-[11px] font-bold border ${!samplingLogRequestedOnly ? 'bg-orange-600 text-white border-orange-700' : 'bg-white text-slate-600 border-slate-300 hover:bg-slate-50'}`}>전체</button>
                <button type="button" onClick={() => setSamplingLogRequestedOnly(true)} className={`px-2.5 py-1 rounded-full text-[11px] font-bold border ${samplingLogRequestedOnly ? 'bg-orange-600 text-white border-orange-700' : 'bg-white text-slate-600 border-slate-300 hover:bg-slate-50'}`}>의뢰만 보기 ({logStats.requestedCount})</button>
              </span>
            </h2>
            <div className="overflow-x-auto max-h-[70vh] overflow-y-auto rounded-lg border border-slate-200">
              <table className="w-full text-[11px] text-center border-collapse min-w-[1180px]">
                <thead className="sticky top-0 z-10">
                  <tr className="bg-slate-800 text-white">
                    <th className="p-2 border border-slate-700 w-10">의뢰</th>
                    <th className="p-2 border border-slate-700 text-left min-w-[190px]">물질명</th>
                    <th className="p-2 border border-slate-700">일차</th>
                    <th className="p-2 border border-slate-700">시작</th>
                    <th className="p-2 border border-slate-700">종료</th>
                    <th className="p-2 border border-slate-700">시작값</th>
                    <th className="p-2 border border-slate-700">종료값</th>
                    <th className="p-2 border border-slate-700 bg-orange-700">채취유량(L)</th>
                    <th className="p-2 border border-slate-700">온도(℃)</th>
                    <th className="p-2 border border-slate-700">압력(mmH2O)</th>
                    <th className="p-2 border border-slate-700">흡입속도</th>
                    <th className="p-2 border border-slate-700">목표(L)</th>
                    <th className="p-2 border border-slate-700">소요(min)</th>
                    <th className="p-2 border border-slate-700">검사기간</th>
                    <th className="p-2 border border-slate-700">시료번호</th>
                  </tr>
                </thead>
                <tbody>
                  {SAMPLING_LOG_GROUPS.map((group) => {
                    const members = SAMPLING_LOG_SUBSTANCES.filter((sub) => group.kinds.includes(sub.kind));
                    const visible = samplingLogRequestedOnly ? members.filter((sub) => log.rows[String(sub.row)]?.requested) : members;
                    const requestedInGroup = members.filter((sub) => log.rows[String(sub.row)]?.requested).length;
                    if (samplingLogRequestedOnly && visible.length === 0) return null;
                    return (
                      <React.Fragment key={`log-group-${group.id}`}>
                        <tr className={`${SAMPLING_LOG_GROUP_HEADER[group.color]} border-y`}>
                          <td colSpan={15} className="p-1.5 text-left">
                            <div className="flex flex-wrap items-center gap-2">
                              <span className="font-black text-xs">{group.title}</span>
                              <span className="text-[10px] opacity-70">{group.note}</span>
                              <span className="px-1.5 py-0.5 rounded-full bg-white/80 border border-current/20 text-[10px] font-bold">{requestedInGroup}/{members.length} 의뢰</span>
                              <span className="flex gap-1">
                                <button type="button" onClick={() => setGroupRequested(group, true)} className="px-2 py-0.5 rounded border border-current/30 bg-white/70 text-[10px] font-bold hover:bg-white">전체 의뢰</button>
                                <button type="button" onClick={() => setGroupRequested(group, false)} className="px-2 py-0.5 rounded border border-current/30 bg-white/70 text-[10px] font-bold hover:bg-white">해제</button>
                              </span>
                            </div>
                          </td>
                        </tr>
                        {visible.map((sub) => {
                          const row = log.rows[String(sub.row)];
                          const on = row.requested;
                          const volume = calcSamplingLogVolume(logForCalc, sub);
                          const isMeter = sub.kind === 'particle' || sub.kind === 'gas' || sub.kind === 'box' || sub.kind === 'plain';
                          const timeDone = on && row.start && row.end;
                          return (
                            <tr key={`log-row-${sub.row}`} className={on ? 'bg-orange-50/60' : 'text-slate-500'}>
                              <td className="p-1 border border-slate-200">
                                <input type="checkbox" name={`log-row-${sub.row}-requested`} checked={on} onChange={(e) => updateSamplingLogRow(String(sub.row), 'requested', e.target.checked)} className="w-4 h-4" />
                              </td>
                              <td className="p-1 border border-slate-200 text-left whitespace-nowrap">
                                <div className="flex items-center gap-1.5">
                                  <span className={`font-bold ${on ? 'text-slate-900' : 'text-slate-500'}`}>{sub.species || sub.name}</span>
                                  {sub.part && <span className="text-[10px] text-slate-500">{sub.part}</span>}
                                  {Array.isArray(sub.options) && getSamplingLogRowsOfSpecies(sub)[0].row === sub.row && (
                                    <span className="inline-flex items-center gap-2 ml-1">
                                      {sub.options.map((opt) => {
                                        const picked = Array.isArray(row.speciesSel) ? row.speciesSel.includes(opt) : true;
                                        return (
                                          <label key={opt} className={`inline-flex items-center gap-1 text-[10px] font-semibold ${on ? 'text-slate-700' : 'text-slate-400'}`}>
                                            <input
                                              type="checkbox"
                                              checked={on && picked}
                                              disabled={!on}
                                              onChange={(e) => {
                                                const current = Array.isArray(row.speciesSel) ? row.speciesSel : [...sub.options];
                                                const next = e.target.checked ? sub.options.filter((o) => o === opt || current.includes(o)) : current.filter((o) => o !== opt);
                                                updateSamplingLogRow(String(sub.row), 'speciesSel', next.length > 0 ? next : [opt]);
                                              }}
                                              className="w-3.5 h-3.5"
                                            />
                                            {opt}
                                          </label>
                                        );
                                      })}
                                    </span>
                                  )}
                                  <span className={`px-1.5 py-0.5 rounded border text-[10px] font-bold ${SAMPLING_LOG_KIND_BADGE[sub.kind]}`}>{sub.code}</span>
                                  {on && (
                                    <span className={`w-2 h-2 rounded-full ${timeDone ? 'bg-emerald-500' : 'bg-amber-400'}`} title={timeDone ? '시각 입력 완료' : '시작/종료 시각 미입력'} />
                                  )}
                                </div>
                              </td>
                              <td className="p-1 border border-slate-200">
                                <select value={row.dayIndex} disabled={!on} onChange={(e) => updateSamplingLogRow(String(sub.row), 'dayIndex', parseInt(e.target.value, 10))} className={cellInput}>
                                  {log.days.map((day, idx) => <option key={idx} value={idx}>{idx + 1}일차{day.date ? ` (${day.date.slice(5)})` : ''}</option>)}
                                </select>
                              </td>
                              <td className="p-1 border border-slate-200"><input type="time" value={row.start} disabled={!on} onChange={(e) => updateSamplingLogRow(String(sub.row), 'start', e.target.value)} className={cellInput} /></td>
                              <td className="p-1 border border-slate-200"><input type="time" value={row.end} disabled={!on} onChange={(e) => updateSamplingLogRow(String(sub.row), 'end', e.target.value)} className={cellInput} /></td>
                              <td className="p-1 border border-slate-200"><input type="number" step="0.01" value={row.meterStart} disabled={!on || !isMeter} onChange={(e) => updateSamplingLogRow(String(sub.row), 'meterStart', e.target.value)} className={cellInput} /></td>
                              <td className="p-1 border border-slate-200"><input type="number" step="0.01" value={row.meterEnd} disabled={!on || !isMeter} onChange={(e) => updateSamplingLogRow(String(sub.row), 'meterEnd', e.target.value)} className={cellInput} /></td>
                              <td className={`p-1 border border-slate-200 font-black ${on && Number.isFinite(volume) ? 'text-orange-800 bg-orange-100/70' : ''}`}>{on && Number.isFinite(volume) ? volume.toFixed(2) : '-'}</td>
                              <td className="p-1 border border-slate-200"><input type="number" step="0.1" value={row.temp} disabled={!on} onChange={(e) => updateSamplingLogRow(String(sub.row), 'temp', e.target.value)} className={cellInput} /></td>
                              <td className="p-1 border border-slate-200"><input type="number" step="0.1" value={row.pressureH2O} disabled={!on} onChange={(e) => updateSamplingLogRow(String(sub.row), 'pressureH2O', e.target.value)} className={cellInput} /></td>
                              <td className="p-1 border border-slate-200"><input type="text" value={row.suctionRate} disabled={!on} onChange={(e) => updateSamplingLogRow(String(sub.row), 'suctionRate', e.target.value)} className={cellInput} /></td>
                              <td className="p-1 border border-slate-200"><input type="text" value={row.targetVolume} disabled={!on} onChange={(e) => updateSamplingLogRow(String(sub.row), 'targetVolume', e.target.value)} className={cellInput} /></td>
                              <td className="p-1 border border-slate-200"><input type="text" value={row.duration} disabled={!on} onChange={(e) => updateSamplingLogRow(String(sub.row), 'duration', e.target.value)} className={cellInput} /></td>
                              <td className="p-1 border border-slate-200"><input type="text" value={row.testPeriod} disabled={!on} onChange={(e) => updateSamplingLogRow(String(sub.row), 'testPeriod', e.target.value)} className={cellInput} /></td>
                              <td className="p-1 border border-slate-200 text-[10px] font-mono text-slate-600 whitespace-nowrap">{on ? buildSamplingLogSampleNo(log, sub) : ''}</td>
                            </tr>
                          );
                        })}
                      </React.Fragment>
                    );
                  })}
                </tbody>
              </table>
            </div>
          </div>

          <div className="flex items-center justify-center gap-4 mt-8 pb-8 border-b-2 border-slate-200 border-dashed">
            <button type="button" onClick={resetSamplingLogForm} className="px-8 py-3.5 bg-white border border-slate-300 text-slate-700 font-bold rounded-xl hover:bg-slate-50 transition-colors shadow-sm">
              새 기록지 작성
            </button>
            <button type="button" onClick={() => exportSamplingLogExcel(null)} className="px-8 py-3.5 bg-white border border-orange-300 text-orange-800 font-bold rounded-xl hover:bg-orange-50 transition-colors shadow-sm">
              현재 내용 엑셀로 추출
            </button>
            <button type="button" onClick={handleSaveSamplingLog} className="px-8 py-3.5 text-white font-bold rounded-xl transition-colors shadow-md bg-orange-600 hover:bg-orange-700">
              기록부 데이터 저장
            </button>
          </div>

          <div className="bg-white p-6 rounded-2xl shadow-sm mt-8 text-slate-800 border border-slate-200">
            <div className="flex flex-col md:flex-row justify-between items-center mb-6 gap-4 border-b border-slate-200 pb-4">
              <h2 className="text-lg font-bold text-slate-900">저장된 종합 리포트 ({displayUserNameWithFallback})</h2>
              <div className="flex items-center gap-2">
                <span className="text-[11px] text-slate-500 font-bold">선택 {sheetCheckedReportKeysValid.length}건</span>
                <button onClick={handleDeleteCheckedSheetReports} className="px-3 py-1.5 border border-red-300 text-red-600 bg-white hover:bg-red-50 font-bold rounded-lg text-xs">선택 삭제</button>
                <button onClick={handleExportSelectedTemplateExcel} className="px-3 py-1.5 text-white font-bold rounded-lg text-xs bg-orange-600 hover:bg-orange-700">선택 항목 엑셀로 추출</button>
              </div>
            </div>
            {renderReportGroupTable({
              groups: groupSheetReportsForLog(savedLogs, resolveReportSheetId, null, { includeBlank: true }),
              checkedSet: sheetCheckedReportKeySet,
              toggleCheck: toggleSheetReportCheck,
              allChecked: savedLogs.length > 0 && savedLogs.every((item) => sheetCheckedReportKeySet.has(getReportKey(item))),
              toggleAll: (checked) => savedLogs.forEach((item) => toggleSheetReportCheck(getReportKey(item), checked)),
              emptyText: '저장된 통합시료채취기록부가 없습니다.',
            })}
          </div>
      </>
    );
  };

  const importSheetInfoToSamplingLog = () => {
    let source = { company: formData.company || '', location: formData.location || '', date: formData.date || '' };
    if (IS_POPUP_WINDOW || selectedSheet === 'samplinglog') {
      try {
        const raw = localStorage.getItem(STORAGE_KEYS.sheetContext);
        const parsed = raw ? JSON.parse(raw) : null;
        if (parsed && typeof parsed === 'object') {
          source = { company: String(parsed.company || ''), location: String(parsed.location || ''), date: String(parsed.date || '') };
        }
      } catch (error) {
        console.error('기록부 정보 읽기 실패:', error);
      }
    }
    if (!source.company && !source.location && !source.date) {
      alert('가져올 기록부 정보가 없습니다. 다른 창에서 기록부를 열고 사업장명·배출구·측정일자를 먼저 입력해주세요.');
      return;
    }
    setSamplingLogData(prev => ({
      ...prev,
      company: prev.company || source.company,
      stackName: prev.stackName || source.location,
      days: prev.days.map((day, idx) => (idx === 0 && !day.date && source.date ? { ...day, date: source.date } : day)),
    }));
  };

  // 기록부 화면 공통 배경(하늘·굴뚝 풍경): 화면에 고정되어 스크롤해도 끊기지 않음
  const renderSheetSkyScene = () => (
      <div className="pointer-events-none fixed inset-0 z-0 overflow-hidden opacity-[0.45]">
        <div className={`absolute inset-0 ${sheetSkyOverlayClass}`} />
        {skyOrbClass && <div className={skyOrbClass} />}
        {renderSkyDecor('sheet', 5)}
        {isNightSky && sceneStarsSheet.map((star) => (
          <span
            key={`sheet-${star.id}`}
            className="skyline-star"
            style={{
              left: star.left,
              top: star.top,
              width: `${star.size}px`,
              height: `${star.size}px`,
              '--twinkle-duration': `${star.duration}s`,
              '--twinkle-delay': `-${star.delay}s`,
            }}
          />
        ))}
        <div className="absolute bottom-0 left-0 right-0 h-[16rem] md:h-[19rem]">
          {renderSkylineSvg({ mode: resolvedSkyMode, scope: 'sheet', compact: true })}
        </div>
      </div>
  );

  if (selectedSheet === 'samplinglog') {
    return (
      <div className={`theme-samplinglog relative min-h-screen p-4 md:p-8 font-sans text-slate-800 ${activeTheme.selection} ${sheetSkyBgClass}`}>
        <style>{SKYLINE_FLYER_CSS}</style>
        {activeTheme.pageTint && <div className={`pointer-events-none fixed inset-0 z-0 ${activeTheme.pageTint}`} />}
        {renderSheetSkyScene()}
        <div className="relative z-10 max-w-[1400px] mx-auto space-y-6 ds-page">
          <div className={`relative bg-white p-6 rounded-xl shadow-sm border-t-4 ${activeTheme.headerBorder} text-center`}>
            <h1 className="text-2xl font-black text-slate-900">
              <img src={mainLogoSrc} alt="STACKPILOT" className="h-[2.2rem] md:h-[3.2rem] w-auto select-none mx-auto" />
            </h1>
            <p className="ds-sheet-title">
              <span>통합시료채취기록부</span>
              <span className="ds-sheet-title-sub">사업장별 채취 물질과 가스미터 지시값 · 날짜별 시트로 추출</span>
            </p>
            <div className="mt-3 md:mt-0 md:absolute md:right-4 md:top-4 flex flex-wrap justify-center gap-2">
              <button type="button" onClick={importSheetInfoToSamplingLog} className="px-3 py-2 text-xs font-bold rounded-lg border border-orange-300 bg-orange-50 text-orange-800 hover:bg-orange-100" title="다른 창에서 열어 둔 기록부의 사업장명·배출구·측정일자를 비어 있는 칸에 채웁니다">
                현재 기록부 정보 가져오기
              </button>
              <button type="button" onClick={navigateToMenu} className="px-3 py-2 text-xs font-bold rounded-lg border border-slate-300 bg-white hover:bg-slate-50">
                {IS_POPUP_WINDOW ? '창 닫기' : '기록부 선택'}
              </button>
            </div>
          </div>
          {renderSamplingLogBody()}
        </div>
      </div>
    );
  }

  return (
    <div className={`theme-${selectedSheet} relative min-h-screen p-4 md:p-8 font-sans text-slate-800 ${activeTheme.selection} ${sheetSkyBgClass}`}>
      <style>{SKYLINE_FLYER_CSS}</style>
      {selectedSheet !== 'dust' && <style>{THEME_OVERRIDE_CSS}</style>}
      {activeTheme.pageTint && <div className={`pointer-events-none fixed inset-0 z-0 ${activeTheme.pageTint}`} />}
      {renderSheetSkyScene()}
      <div className="max-w-7xl mx-auto space-y-6 relative z-10 ds-page">
        
        {/* 헤더 섹션 */}
        <div className={`bg-white p-6 rounded-xl shadow-md border-t-4 relative ${activeTheme.headerBorder}`}>
          <div className="flex flex-col items-center text-center">
            <h1 className="text-2xl font-black text-slate-900">
              <img
                src={mainLogoSrc}
                alt="STACKPILOT"
                onError={() => setLogoIndex((prev) => Math.min(prev + 1, logoCandidates.length - 1))}
                className="h-[2.2rem] md:h-[3.2rem] w-auto select-none"
              />
            </h1>
            <p className="ds-sheet-title">
              <span>{activeSheet?.title || '먼지시료채취기록부'}</span>
              <span className="ds-sheet-title-sub">대기오염공정시험기준</span>
            </p>
          </div>
          <button
            type="button"
            onClick={navigateToMenu}
            className="mt-3 md:mt-0 md:absolute md:right-4 md:top-4 px-3 py-2 text-xs font-bold rounded-lg border border-slate-300 bg-white hover:bg-slate-50"
          >
            {IS_POPUP_WINDOW ? '창 닫기' : '기록부 선택'}
          </button>
        </div>

        <div className="bg-white p-3 rounded-xl shadow-sm border border-slate-200 flex flex-col md:flex-row md:items-center md:justify-between gap-2">
          <p className="text-xs text-slate-600 ds-meta">
            <span className="ds-meta-user">{displayUserNameWithFallback}</span>
            <span className={isUserUnlocked ? 'ds-meta-ok' : 'ds-meta-warn'}>{isUserUnlocked ? '로그인됨' : '로그인 필요'}</span>
            <span className="ds-meta-dim">계정 설정은 첫 화면에서</span>
          </p>
          <div className="flex flex-wrap items-center gap-2 shrink-0">
            <span className="text-[11px] font-bold text-slate-500">장비 프리셋 (노즐/샘플러/피토관)</span>
            <button
              type="button"
              onClick={saveAllEquipmentPresets}
              className="px-3 py-1.5 text-xs font-bold bg-emerald-600 text-white rounded-lg border border-emerald-700 hover:bg-emerald-700 inline-flex items-center gap-1"
            >
              <PresetSaveIcon />
              프리셋 저장하기
            </button>
          </div>
        </div>
        <SheetSectionNav scope={selectedSheet || 'sheet'} />

        <form onSubmit={handleSave} onKeyDown={handleFormKeyDown} className="space-y-6 ds-paper">
          
          {/* 섹션 1: 일반 사항 및 배출구 제원 */}
          <div className="bg-white p-6 rounded-xl shadow-sm border border-slate-200">
            <h2 className="ds-sec-title">
              <span className="ds-sec-num">1</span>
              일반 사항 및 배출구 제원
            </h2>
            <div className="grid grid-cols-1 md:grid-cols-4 gap-4 mb-4">
              <div><label className="block text-xs font-bold text-slate-600 mb-1">측정 일자</label>
                <input type="date" name="date" value={formData.date} onChange={handleChange} className="date-input-fixed w-full h-10 p-2 bg-emerald-50/50 border border-emerald-200 rounded focus:ring-2 focus:ring-emerald-500 outline-none transition-colors" /></div>
              <div><label className="block text-xs font-bold text-slate-600 mb-1">사업장명</label>
                <input type="text" name="company" value={formData.company} onChange={handleChange} className="w-full h-10 p-2 bg-emerald-50/50 border border-emerald-200 rounded focus:ring-2 focus:ring-emerald-500 outline-none transition-colors" /></div>
              <div><label className="block text-xs font-bold text-slate-600 mb-1">배출구명 (위치)</label>
                <input type="text" name="location" value={formData.location} onChange={handleChange} className="w-full h-10 p-2 bg-emerald-50/50 border border-emerald-200 rounded focus:ring-2 focus:ring-emerald-500 outline-none transition-colors" /></div>
              <div><label className="block text-xs font-bold text-slate-600 mb-1">측정자</label>
                <input type="text" name="sampler" value={formData.sampler} onChange={handleChange} className="w-full h-10 p-2 bg-emerald-50/50 border border-emerald-200 rounded focus:ring-2 focus:ring-emerald-500 outline-none transition-colors" /></div>
            </div>
            <div className="grid grid-cols-1 md:grid-cols-4 gap-4 bg-emerald-50/30 p-4 rounded-lg border border-emerald-100">
              <div><label className="block text-xs font-bold text-slate-600 mb-1">날씨</label>
                <input type="text" name="weather" value={formData.weather} onChange={handleChange} className="w-full p-2 border border-slate-300 rounded focus:ring-2 focus:ring-emerald-500 outline-none" /></div>
              <div><label className="block text-xs font-bold text-slate-600 mb-1">대기 온도 (℃)</label>
                <input type="number" step="0.1" name="atmTemp" value={formData.atmTemp} onChange={handleChange} className="w-full p-2 border border-slate-300 rounded focus:ring-2 focus:ring-emerald-500 outline-none" /></div>
              <div className="md:col-span-2"><label className="block text-xs font-bold text-slate-600 mb-1">대기압 (측정공 주변, mmHg)</label>
                <input type="number" step="1" name="atmPressure" value={formData.atmPressure} onChange={handleChange} className="w-full p-2 border border-slate-300 rounded focus:ring-2 focus:ring-emerald-500 outline-none" /></div>
            </div>
            <div className="mt-2 flex flex-wrap items-center gap-2 text-xs text-slate-600">
              <span className="font-bold">날씨 위치</span>
              {weatherStationSelect}
              <button type="button" onClick={applyKmaWeatherToSheet} disabled={weatherBusy} className="px-3 py-1.5 rounded border border-slate-300 bg-white text-slate-700 hover:bg-slate-100 font-bold disabled:opacity-50">
                {weatherBusy ? '가져오는 중…' : '기상청 자료 가져오기'}
              </button>
              <span className="text-slate-500">측정 일자와 채취 시작 시각의 관측값으로 날씨 칸을 채웁니다 (기온·기압은 현장 측정값)</span>
            </div>

            {/* 굴뚝 제원 (직경 산출) */}
            <div className="mt-4 grid grid-cols-1 md:grid-cols-4 gap-4 bg-emerald-50 p-4 rounded-lg border border-emerald-200">
              <div>
                <label className="block text-xs font-bold text-emerald-800 mb-1">총 측정 깊이 (m)</label>
                <input type="number" step="0.0001" name="totalStackDepth" value={formData.totalStackDepth} onChange={handleChange} placeholder="끝벽까지의 총 길이" className="w-full p-2 border border-emerald-300 rounded focus:ring-2 focus:ring-emerald-500 bg-white outline-none" />
              </div>
              <div>
                <label className="block text-xs font-bold text-emerald-800 mb-1">플랜지 길이 (m)</label>
                <input type="number" step="0.0001" name="flangeLength" value={formData.flangeLength} onChange={handleChange} placeholder="외벽 돌출부" className="w-full p-2 border border-emerald-300 rounded focus:ring-2 focus:ring-emerald-500 bg-white outline-none" />
              </div>
              <div>
                <label className="block text-xs font-bold text-emerald-900 mb-1">순수 굴뚝 내경 (D, m)</label>
                <input type="number" step="0.0001" name="stackDiameter" value={formData.stackDiameter} onChange={handleChange} className="w-full p-2 border border-emerald-400 rounded focus:ring-2 focus:ring-emerald-500 bg-emerald-100 font-black text-emerald-900 outline-none" placeholder="자동계산 (총 깊이 - 플랜지)" />
              </div>
              <div>
                <label className="block text-xs font-bold text-emerald-900 mb-1">단면적 (A, ㎡)</label>
                <input type="text" readOnly value={samplingPointsData ? samplingPointsData.area : ''} className="w-full p-2 border border-emerald-400 rounded bg-emerald-100 font-black text-emerald-900 outline-none" placeholder="자동계산" />
              </div>
            </div>

            {/* 굴뚝 내경에 따른 측정점 산출 패널 */}
            {samplingPointsData && (
              <div className="mt-4 bg-white p-4 rounded-lg border border-slate-200 shadow-sm">
                <div className="flex justify-between items-start mb-3">
                  <h3 className="text-sm font-bold text-slate-800">
                    측정점 위치 (내경 {formData.stackDiameter} m)
                  </h3>
                  {samplingPointsData.isCenterOnly && (
                    <span className="bg-red-50 text-red-600 px-2 py-1 rounded text-xs font-bold animate-pulse border border-red-200">
                      단면적 0.25㎡ 이하 (단면중심 1점 측정)
                    </span>
                  )}
                </div>
                
                <div className="flex flex-col md:flex-row gap-4 items-start">
                  <figure className="shrink-0 flex flex-col items-center gap-1 m-0 px-1 md:w-60">
                    {renderSamplingPointDiagram(samplingPointsData)}
                    <figcaption className="mt-1 text-[11px] text-slate-700 text-center leading-snug">
                      그림. 측정공 1개당 측정점 위치 (반경당 {samplingPointsData.perRadius}점)<br />
                      <span className="text-slate-500"><span className="text-[#0a84ff]">●</span> 정방향 측정점 1~{samplingPointsData.perRadius} <span className="text-[#d70015]">○</span> 역방향(반대편) 1′~{samplingPointsData.perRadius}′</span>
                    </figcaption>
                  </figure>
                  <div className="flex-1 w-full overflow-x-auto bg-white rounded border border-slate-200">
                    <table className="w-full text-xs text-center border-collapse">
                      <thead>
                        <tr className="bg-slate-50 text-slate-700">
                          <th className="border-r border-b border-slate-200 p-2 whitespace-nowrap" rowSpan="2">삽입 깊이 (마킹 위치, cm)</th>
                          <th className="border-b border-slate-200 p-2 whitespace-nowrap" colSpan={samplingPointsData.perRadius}>측정 포인트 (1~5)</th>
                        </tr>
                        <tr className="bg-slate-50 text-slate-700">
                          {samplingPointsData.nearInsertion.map((_, i) => (
                            <th key={i} className="border-b border-l border-slate-200 p-1">{i + 1}번</th>
                          ))}
                        </tr>
                      </thead>
                      <tbody>
                        <tr>
                          <td className="border-r border-b border-slate-200 p-2 font-bold text-slate-800 whitespace-nowrap bg-slate-50/50">
                            가까운 쪽 (정방향)<br/><span className="text-[9px] font-normal text-slate-600">기본 측정</span>
                          </td>
                          {samplingPointsData.nearInsertion.map((d, i) => (
                            <td key={i} className="border-l border-b border-slate-200 p-2 font-black text-emerald-700 bg-white">
                              {d} cm
                            </td>
                          ))}
                        </tr>
                        {!samplingPointsData.isCenterOnly && (
                          <tr>
                            <td className="border-r border-slate-200 p-2 font-bold text-red-700 whitespace-nowrap bg-red-50">
                              깊은 쪽 (역방향)<br/><span className="text-[9px] font-normal text-red-500">반대편 측정공 대신 쓸 때</span>
                            </td>
                            {samplingPointsData.farInsertion.map((d, i) => (
                              <td key={i} className="border-l border-slate-200 p-2 font-black text-red-600 bg-red-50/50">
                                {d} cm
                              </td>
                            ))}
                          </tr>
                        )}
                      </tbody>
                    </table>
                  </div>
                </div>
                
              </div>
            )}

            {/* 누출시험 (양식 '누출시험' 시트). 유량측정기록부에는 없음 */}
            {!isFlowSheet && (() => {
              const leak = sanitizeLeakTest(formData.leakTest);
              const setLeak = (key, value) => setFormData(prev => ({ ...prev, leakTest: { ...sanitizeLeakTest(prev.leakTest), [key]: value } }));
              const diff = parseFloat(leak.meterAfter) - parseFloat(leak.meterBefore);
              return (
                <div className="mt-4 border border-slate-200 rounded-lg p-3">
                  <div className="flex flex-wrap items-center gap-x-4 gap-y-2">
                    <span className="text-sm font-bold text-slate-800 shrink-0">누출시험</span>
                    <span className="text-[11px] text-slate-500">진공 380 mmHg에서 5초간 유량계가 멈추면 O</span>
                    {LEAK_TEST_PARTS.map((part) => (
                      <label key={part.key} className="flex items-center gap-1.5 text-xs font-semibold text-slate-700">
                        {part.label}
                        <select value={leak[part.key]} onChange={(e) => setLeak(part.key, e.target.value)} className={`p-1 border rounded text-xs font-bold outline-none ${leak[part.key] === 'O' ? 'border-slate-300 text-emerald-700' : 'border-red-300 text-red-600'}`}>
                          <option value="O">O</option>
                          <option value="X">X</option>
                        </select>
                      </label>
                    ))}
                    {!isFlowSheet ? (
                      <label className="flex items-center gap-1.5 text-xs font-semibold text-slate-700">
                        가스미터(L) 전
                        <input type="number" step="0.1" value={leak.meterBefore} onChange={(e) => setLeak('meterBefore', e.target.value)} className="w-24 p-1 border border-slate-300 rounded text-center text-xs outline-none" />
                        후
                        <input type="number" step="0.1" value={leak.meterAfter} onChange={(e) => setLeak('meterAfter', e.target.value)} className="w-24 p-1 border border-slate-300 rounded text-center text-xs outline-none" />
                        <span className="text-slate-500 font-normal">차 {Number.isFinite(diff) ? diff.toFixed(2) : '-'}</span>
                      </label>
                    ) : (
                      <label className="flex items-center gap-1.5 text-xs font-semibold text-slate-700 flex-1 min-w-[200px]">
                        비고
                        <input type="text" value={leak.note} onChange={(e) => setLeak('note', e.target.value)} className="flex-1 p-1 border border-slate-300 rounded text-xs outline-none" />
                      </label>
                    )}
                  </div>
                </div>
              );
            })()}
          </div>

          {/* 섹션 2: 배출가스 조성 및 분자량 (가스분석기 자료) */}
          <div className="bg-white p-6 rounded-xl shadow-sm border border-slate-200">
            <div className="flex flex-col md:flex-row md:items-center justify-between mb-4 border-b border-slate-200 pb-2">
              <h2 className="ds-sec-title is-inline">
                <span className="ds-sec-num">2</span>
                배출가스 조성 및 밀도
              </h2>
            </div>
            
            <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
               <div className="lg:col-span-2 overflow-x-auto">
                  <table className="w-full text-sm text-center border-collapse">
                    <thead>
                      <tr className="bg-teal-50 text-teal-800 border-b border-teal-200">
                        <th className="p-2 border-r border-teal-100 font-bold whitespace-nowrap">회차 (시간)</th>
                        <th className="p-2 border-r border-teal-100 font-bold">O₂ (%)</th>
                        <th className="p-2 border-r border-teal-100 font-bold">CO₂ (%)</th>
                        <th className="p-2 border-r border-teal-100 font-bold">CO (ppm)</th>
                        {/* NOx, SOx 위치 변경 적용됨 */}
                        <th className="p-2 border-r border-teal-100 font-bold">NOx (ppm)</th>
                        <th className="p-2 font-bold">SOx (ppm)</th>
                      </tr>
                    </thead>
                    <tbody>
                      {formData.gasAnalyzer.map((gas, idx) => (
                        <tr key={idx} className="hover:bg-slate-50 border-b border-slate-100">
                          <td className="p-2 border-r border-slate-100 font-medium text-slate-700 flex justify-center items-center gap-2">
                            <span className="whitespace-nowrap">{idx + 1}회차</span>
                            <input
                              type="time"
                              name={`gas-time-${idx}`}
                              autoComplete="off"
                              value={gas.time || ''}
                              onChange={(e) => handleGasAnalyzerChange(idx, 'time', e.target.value)}
                              className="w-24 p-1 border border-slate-300 rounded text-center text-xs outline-none focus:ring-2 focus:ring-teal-500"
                            />
                          </td>
                          <td className="p-1 border-r border-slate-100">
                            <input type="number" step="0.01" value={gas.o2} onChange={(e) => handleGasAnalyzerChange(idx, 'o2', e.target.value)} className="w-full p-1 border border-slate-300 rounded text-center outline-none focus:ring-2 focus:ring-teal-500" />
                          </td>
                          <td className="p-1 border-r border-slate-100">
                            <input type="number" step="0.01" value={gas.co2} onChange={(e) => handleGasAnalyzerChange(idx, 'co2', e.target.value)} className="w-full p-1 border border-slate-300 rounded text-center outline-none focus:ring-2 focus:ring-teal-500" />
                          </td>
                          <td className="p-1 border-r border-slate-100">
                            <input type="number" step="0.01" value={gas.co} onChange={(e) => handleGasAnalyzerChange(idx, 'co', e.target.value)} className="w-full p-1 border border-slate-300 rounded text-center outline-none focus:ring-2 focus:ring-teal-500" />
                          </td>
                          <td className="p-1 border-r border-slate-100">
                            <input type="number" step="0.01" value={gas.nox} onChange={(e) => handleGasAnalyzerChange(idx, 'nox', e.target.value)} className="w-full p-1 border border-slate-300 rounded text-center outline-none focus:ring-2 focus:ring-teal-500" />
                          </td>
                          <td className="p-1">
                            <input type="number" step="0.01" value={gas.sox} onChange={(e) => handleGasAnalyzerChange(idx, 'sox', e.target.value)} className="w-full p-1 border border-slate-300 rounded text-center outline-none focus:ring-2 focus:ring-teal-500" />
                          </td>
                        </tr>
                      ))}
                      <tr className="bg-slate-50 font-bold text-slate-800">
                        <td className="p-2 border-r border-slate-200">평균</td>
                        <td className="p-2 border-r border-slate-200 text-teal-600">{o2.toFixed(2)}</td>
                        <td className="p-2 border-r border-slate-200">{co2.toFixed(2)}</td>
                        <td className="p-2 border-r border-slate-200">{co.toFixed(2)}</td>
                        <td className="p-2 border-r border-slate-200">{nox.toFixed(2)}</td>
                        <td className="p-2">{sox.toFixed(2)}</td>
                      </tr>
                    </tbody>
                  </table>
               </div>
               
               <div className="lg:col-span-1 flex flex-col gap-4">
                 {/* O2 보정 계수 패널 */}
                 <div className="bg-white p-4 rounded-lg border border-teal-200 shadow-sm relative overflow-hidden">
                    <div className="absolute top-0 left-0 w-1.5 h-full bg-teal-500"></div>
                    <h3 className="text-xs font-bold text-teal-800 mb-3 border-b border-teal-100 pb-2">
                      산소(O₂) 보정 계수 산출
                    </h3>
                    <div className="flex justify-between items-center mb-2">
                       <div className="flex flex-col">
                         <span className="text-xs font-bold text-slate-700">표준 산소 농도 (O<sub>s</sub>, %)</span>
                         <span className="text-[9px] text-slate-500">보정 대상이 아니면 비워두세요</span>
                       </div>
                       <input type="number" step="0.1" name="standardO2" value={formData.standardO2} onChange={handleChange} className="w-20 p-1 border border-teal-300 rounded text-center text-sm font-black text-teal-700 bg-teal-50 focus:ring-2 focus:ring-teal-400 outline-none" placeholder="공란" />
                    </div>
                    <div className="flex justify-between items-center mb-2">
                       <span className="text-xs font-bold text-slate-700">실측 산소 농도 (O<sub>a</sub>, %)</span>
                       <span className="font-bold text-slate-800">{o2.toFixed(2)} %</span>
                    </div>
                    <div className="flex justify-between items-center pt-2 mt-1 border-t border-teal-100">
                       <span className="text-xs font-bold text-teal-700" title="(21 - 표준O2) / (21 - 실측O2)">산소 보정 계수 (K)</span>
                       <span className="font-black text-lg text-teal-600">{Number.isFinite(calcO2CorrectionFactor()) ? calcO2CorrectionFactor().toFixed(2) : '-'}</span>
                    </div>
                 </div>

                 {/* 분자량/밀도 패널 */}
                 <div className="bg-slate-50 p-4 rounded-lg border border-slate-200">
                    <div className="flex justify-between items-center mb-1">
                       <span className="text-[11px] font-bold text-slate-600" title="100 - O₂(%) - CO₂(%) - CO(%)">평균 N₂ (%)</span>
                       <span className="text-xs font-bold text-slate-700">{n2.toFixed(2)}</span>
                    </div>
                    <div className="flex justify-between items-center mb-1">
                       <span className="text-[11px] font-bold text-slate-600" title="Σ(M_i * x_i) / 100">건조가스분자량 (M<sub>d</sub>)</span>
                       <span className="text-xs font-black text-slate-700">{Md.toFixed(2)}</span>
                    </div>
                    <div className="flex justify-between items-center mb-1 pt-1 border-t border-slate-200/50">
                       <span className="text-[11px] font-bold text-slate-600" title="사전 수분량 반영됨">습가스분자량 (M<sub>s</sub>)</span>
                       <span className="text-sm font-black text-slate-700">{Ms.toFixed(2)}</span>
                    </div>
                    <div className="flex justify-between items-center pt-1 mt-1 border-t border-slate-200/50">
                       <span className="text-[11px] font-bold text-slate-600" title="1 / (22.4 * 100) * [ Σ(M_i*x_i) * (100 - Xw)/100 + 18 * Xw ]">배출가스밀도 (γ<sub>a</sub>)</span>
                       <span className="text-sm font-black text-teal-600">{r0.toFixed(2)}</span>
                    </div>
                    <div className="flex justify-between items-center pt-1 mt-1 border-t border-slate-200/50">
                       <span className="text-[11px] font-bold text-slate-600" title="온도·압력 보정 밀도">배출가스밀도 (r)</span>
                       <span className="text-sm font-black text-teal-600">{isNaN(r) ? '-' : r.toFixed(2)}</span>
                    </div>
                 </div>
               </div>
            </div>
          </div>

          <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
            {/* 섹션 3: 배출가스 수분량 */}
            <div className="bg-white p-6 rounded-xl shadow-sm border border-slate-200">
              <h2 className="ds-sec-title">
                <span className="ds-sec-num">3</span>
                배출가스 수분량
              </h2>
              <div className="space-y-4">
                <div>
                  <label className="block text-xs font-bold text-slate-600 mb-2">측정방법</label>
                  <div className="ds-segment" role="tablist">
                    {MOISTURE_METHODS.map((method) => (
                      <button
                        key={method.id}
                        type="button"
                        role="tab"
                        aria-selected={activeMoistureMethod.id === method.id}
                        onClick={() => setFormData((prev) => ({ ...prev, moistureMethod: method.id }))}
                        className={`ds-segment-item${activeMoistureMethod.id === method.id ? ' is-on' : ''}`}
                      >
                        {MOISTURE_METHODS.indexOf(method) + 1}. {method.label}
                      </button>
                    ))}
                  </div>
                </div>

                {activeMoistureMethod.id === 'auto' && (
                  <div>
                    <p className="text-[11px] text-slate-600 mb-3"><b>{activeMoistureMethod.standard} {activeMoistureMethod.label}</b> · {activeMoistureMethod.guide}</p>
                    <label className="block text-xs font-bold text-slate-600 mb-2">수분량 자동측정기 (%)</label>
                    <div className="grid grid-cols-5 gap-2">
                      {formData.moistureValues.map((val, idx) => (
                        <div key={idx}>
                          <label className="block text-[10px] text-slate-500 text-center mb-1">{idx + 1}회</label>
                          <input
                            type="number" step="0.01" value={val}
                            onChange={(e) => handleMoistureChange(idx, e.target.value)}
                            className="w-full p-2 border border-slate-300 rounded focus:ring-2 focus:ring-cyan-500 bg-slate-50 text-center text-sm outline-none"
                          />
                        </div>
                      ))}
                    </div>
                  </div>
                )}

                {isImpingerMoistureMethod(activeMoistureMethod) && (() => {
                  const pre = formData.preImpinger || createEmptyPreImpinger();
                  const setPre = (patch) => setFormData((prev) => ({ ...prev, preImpinger: { ...(prev.preImpinger || createEmptyPreImpinger()), ...patch } }));
                  const setImp = (idx, key, value) => setPre({ impingers: (pre.impingers || []).map((imp, i) => (i === idx ? { ...imp, [key]: value } : imp)) });
                  const n = (v) => { const x = parseFloat(v); return Number.isFinite(x) ? x : NaN; };
                  const stages = activeMoistureMethod.stages;
                  const vlc = (pre.impingers || []).slice(0, stages).reduce((sum, imp) => { const d = n(imp.final) - n(imp.initial); return Number.isFinite(d) ? sum + d : sum; }, 0);
                  const yd = getPreMoistureYd(activeMoistureMethod, formData);
                  const pv = getPreMoisturePv(activeMoistureMethod, formData);
                  const vol = (n(pre.meterEnd) - n(pre.meterStart)) * yd;
                  const gaugeMmHg = n(pre.orifice) / 13.6;
                  const result = getMoistureByMethod(activeMoistureMethod.id);
                  const cellInput = (value, onChange, placeholder = '') => (
                    <input type="number" step="0.01" value={value} placeholder={placeholder} onChange={(e) => onChange(e.target.value)} className="w-full p-1.5 text-center outline-none" />
                  );
                  return (
                    <div className="space-y-3">
                      <p className="text-[11px] text-slate-600"><b>{activeMoistureMethod.standard} {activeMoistureMethod.label.replace(/\s*\(.*\)$/, '')}</b> · {activeMoistureMethod.guide}</p>
                      {activeMoistureMethod.id === 'imp2' && (
                        <div className="flex flex-wrap items-center gap-2 text-xs text-slate-700">
                          <span className="font-bold">가스 샘플러</span>
                          <select
                            value={pre.gasSamplerId || ''}
                            onChange={(e) => {
                              const found = gasSamplers.find((item) => String(item.id) === String(e.target.value));
                              setPre({ gasSamplerId: e.target.value, ...(found ? { yd: String(found.yd) } : {}) });
                            }}
                            className="p-1.5 border border-slate-300 rounded bg-white outline-none"
                            title="흡습관법에 쓴 가스 샘플러"
                          >
                            <option value="">직접 입력</option>
                            {gasSamplers.map((item) => <option key={item.id} value={item.id}>{`가스샘플러 ${item.id}${item.sn ? ` (${item.sn})` : ''}`}</option>)}
                          </select>
                          <span className="font-bold">보정계수 Y<sub>d</sub></span>
                          <input type="number" step="0.0001" value={pre.yd || ''} onChange={(e) => setPre({ yd: e.target.value, gasSamplerId: '' })} placeholder="1" title="가스 샘플러 보정계수" className="w-24 p-1.5 border border-slate-300 rounded text-center outline-none" />
                          <label className="inline-flex items-center gap-1">
                            <input type="checkbox" checked={Boolean(pre.wetMeter)} onChange={(e) => setPre({ wetMeter: e.target.checked })} className="w-3.5 h-3.5" />
                            습식 가스미터 (포화수증기압 P<sub>v</sub> 차감)
                          </label>
                          {pv > 0 && <span className="text-slate-500">P<sub>v</sub> = {pv.toFixed(1)} mmHg</span>}
                        </div>
                      )}
                      <table className="w-full text-sm text-center border-collapse">
                        <thead>
                          <tr className="bg-slate-100 text-slate-700">
                            <th className="p-2 border border-slate-200 font-bold">{activeMoistureMethod.unitTitle}</th>
                            <th className="p-2 border border-slate-200 font-bold">측정전 무게 m<sub>a1</sub> (g)</th>
                            <th className="p-2 border border-slate-200 font-bold">측정후 무게 m<sub>a2</sub> (g)</th>
                            <th className="p-2 border border-slate-200 font-bold">흡습 수분량 (g)</th>
                          </tr>
                        </thead>
                        <tbody>
                          {(pre.impingers || []).slice(0, stages).map((imp, idx) => {
                            const gain = n(imp.final) - n(imp.initial);
                            return (
                              <tr key={imp.id || idx}>
                                <td className="p-1.5 border border-slate-200 font-bold text-slate-600 whitespace-nowrap">{activeMoistureMethod.items[idx]}</td>
                                <td className="p-1 border border-slate-200">{cellInput(imp.initial, (v) => setImp(idx, 'initial', v))}</td>
                                <td className="p-1 border border-slate-200">{cellInput(imp.final, (v) => setImp(idx, 'final', v))}</td>
                                <td className="p-1.5 border border-slate-200 font-bold text-slate-700">{Number.isFinite(gain) ? gain.toFixed(2) : '-'}</td>
                              </tr>
                            );
                          })}
                          <tr className="bg-slate-100 font-bold">
                            <td className="p-1.5 border border-slate-200">흡습 수분의 질량 m<sub>a</sub></td>
                            <td className="p-1.5 border border-slate-200" colSpan="2"></td>
                            <td className="p-1.5 border border-slate-200">{vlc > 0 ? vlc.toFixed(2) : '-'}</td>
                          </tr>
                        </tbody>
                      </table>
                      <div className="overflow-x-auto">
                        <table className="w-full text-sm text-center border-collapse min-w-[640px]">
                          <thead>
                            <tr className="bg-slate-100 text-slate-700">
                              <th className="p-2 border border-slate-200 font-bold" colSpan="2">가스미터(L)</th>
                              <th className="p-2 border border-slate-200 font-bold">가스채취량<br /><span className="text-[10px] font-normal">(L)</span></th>
                              <th className="p-2 border border-slate-200 font-bold">가스미터온도<br /><span className="text-[10px] font-normal">(℃)</span></th>
                              <th className="p-2 border border-slate-200 font-bold">측정공에서 기압<br /><span className="text-[10px] font-normal">(mm Hg)</span></th>
                              {/* 임핀저법은 입자 샘플러로 흡입하므로 가스미터 게이지압 = 오리피스압(ΔH). 흡습관법은 가스 샘플러의 가스미터압 */}
                              <th className="p-2 border border-slate-200 font-bold">{activeMoistureMethod.id === 'impinger' ? '오리피스압 ΔH' : '가스미터압'}<br /><span className="text-[10px] font-normal">(mm H₂O)</span></th>
                              <th className="p-2 border border-slate-200 font-bold">{activeMoistureMethod.id === 'impinger' ? 'ΔH ÷ 13.6' : '가스미터압'}<br /><span className="text-[10px] font-normal">P<sub>m</sub> (mm Hg)</span></th>
                              <th className="p-2 border border-slate-200 font-bold">수분량<br /><span className="text-[10px] font-normal">(%)</span></th>
                            </tr>
                            <tr className="bg-slate-50 text-slate-600">
                              <th className="p-1 border border-slate-200 font-bold text-xs">측정전</th>
                              <th className="p-1 border border-slate-200 font-bold text-xs">측정후</th>
                              <th className="border border-slate-200" colSpan="6"></th>
                            </tr>
                          </thead>
                          <tbody>
                            <tr>
                              <td className="p-1 border border-slate-200">{cellInput(pre.meterStart, (v) => setPre({ meterStart: v }))}</td>
                              <td className="p-1 border border-slate-200">{cellInput(pre.meterEnd, (v) => setPre({ meterEnd: v }))}</td>
                              <td className="p-1.5 border border-slate-200 font-bold">{Number.isFinite(vol) && vol > 0 ? vol.toFixed(2) : '-'}</td>
                              <td className="p-1 border border-slate-200">{cellInput(pre.meterTemp, (v) => setPre({ meterTemp: v }))}</td>
                              <td className="p-1 border border-slate-200">{cellInput(pre.pressure, (v) => setPre({ pressure: v }), formData.atmPressure || '')}</td>
                              <td className="p-1 border border-slate-200">{cellInput(pre.orifice, (v) => setPre({ orifice: v }))}</td>
                              <td className="p-1.5 border border-slate-200 font-bold">{Number.isFinite(gaugeMmHg) ? gaugeMmHg.toFixed(2) : '-'}</td>
                              <td className="p-1.5 border border-slate-200 font-bold">{result.ready ? result.value.toFixed(2) : '-'}</td>
                            </tr>
                          </tbody>
                        </table>
                      </div>
                      {activeMoistureMethod.maRange && vlc > 0 && (vlc < activeMoistureMethod.maRange[0] || vlc > activeMoistureMethod.maRange[1]) && (
                        <p className="text-[11px] font-bold text-red-600">흡습 수분의 질량이 {vlc.toFixed(2)} g입니다. 기준은 {activeMoistureMethod.maRange[0]} g~{activeMoistureMethod.maRange[1].toFixed(1)} g이 되도록 흡입하라고 정하고 있습니다.</p>
                      )}
                      <p className="text-[11px] text-slate-500">X<sub>w</sub> = (22.4/18 × m<sub>a</sub>) ÷ [V<sub>m</sub> × 273/(273+θ<sub>m</sub>) × (P<sub>a</sub>+P<sub>m</sub>)/760 + 22.4/18 × m<sub>a</sub>] × 100. 가스채취량 V<sub>m</sub> = (측정후 − 측정전) × {activeMoistureMethod.id === 'imp2' ? '가스 샘플러 보정계수' : '입자 샘플러 보정계수'}(Y<sub>d</sub> = {yd}), P<sub>m</sub> = {activeMoistureMethod.id === 'impinger' ? '오리피스압 ΔH' : '가스미터압'}(mmH₂O) ÷ 13.6.{pv > 0 ? ' 습식 가스미터라 (Pa+Pm−Pv)로 계산합니다.' : ''} 측정공에서 기압을 비워 두면 대기압을 씁니다.{result.ready ? ` 결과 표기(소수 첫째 자리): ${result.value.toFixed(1)} %` : ''}</p>
                    </div>
                  );
                })()}

                <div className="bg-cyan-50 p-4 rounded-lg flex justify-between items-center border border-cyan-100 mt-4 shadow-sm">
                  <div className="text-slate-800">
                    <span className="font-bold">사전 적용 수분량</span>
                    <span className="block text-[11px] text-slate-500">{activeMoistureMethod.label}</span>
                  </div>
                  <div className="text-2xl font-black text-cyan-700">{getMoistureByMethod(activeMoistureMethod.id).ready ? `${calcMoisture()}%` : '-'}</div>
                </div>
                
                {!isFlowSheet && (
                <div className="mt-6 pt-4 border-t border-slate-200">
                  <label className="block text-sm font-bold text-slate-800 mb-2 flex items-center gap-1">
                    사후 수분량 (임핀저 무게법)
                  </label>
                  <div className="overflow-x-auto">
                    <table className="w-full text-sm text-center border-collapse">
                      <thead>
                        <tr className="bg-slate-100 text-slate-700">
                          <th className="p-2 border border-slate-200 font-bold rounded-tl-lg">임핀저</th>
                          <th className="p-2 border border-slate-200 font-bold">채취 전 (g)</th>
                          <th className="p-2 border border-slate-200 font-bold">채취 후 (g)</th>
                          <th className="p-2 border border-slate-200 font-bold rounded-tr-lg">증가량 (g)</th>
                        </tr>
                      </thead>
                      <tbody>
                        {visibleImpingers.map((imp, idx) => (
                          <tr key={imp.id} className="hover:bg-slate-50">
                            <td className="p-2 border border-slate-200 text-slate-700 font-bold">{imp.id}단{imp.id === impingerStageCount ? '(실리카겔)' : ''}</td>
                            <td className="p-1 border border-slate-200">
                              <input type="number" step="0.01" value={imp.initial} onChange={(e) => handleImpingerChange(idx, 'initial', e.target.value)} className="w-full p-1 border border-slate-300 rounded text-center outline-none focus:ring-2 focus:ring-emerald-500" />
                            </td>
                            <td className="p-1 border border-slate-200">
                              <input type="number" step="0.01" value={imp.final} onChange={(e) => handleImpingerChange(idx, 'final', e.target.value)} className="w-full p-1 border border-slate-300 rounded text-center outline-none focus:ring-2 focus:ring-emerald-500" />
                            </td>
                            <td className="p-2 border border-slate-200 font-black text-slate-800 bg-slate-50/50">
                              {!isNaN(parseFloat(imp.initial)) && !isNaN(parseFloat(imp.final)) ? (parseFloat(imp.final) - parseFloat(imp.initial)).toFixed(2) : '-'}
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                  <div className="flex flex-col gap-2 mt-3">
                    <div className="flex justify-between items-center bg-slate-50 p-3 rounded-lg border border-slate-200">
                      <span className="text-sm font-bold text-slate-700">포집된 수분 총량 (V<sub>lc</sub>)</span>
                      <span className="font-black text-slate-800">{calcTotalMoistureWeight().toFixed(2)} ml</span>
                    </div>
                    <div className="flex justify-between items-center bg-emerald-600 p-3 rounded-lg shadow-sm text-white">
                      <span className="text-sm font-bold">사후 수분량 (X<sub>w</sub>)</span>
                      <span className="font-black text-xl text-emerald-100">{calcPostMoisture()}%</span>
                    </div>
                  </div>
                </div>
                )}
              </div>
            </div>

            {/* ✅ 4번 섹션 동정압 측정 시의 미터온도(Tm) 입력칸 추가 */}
            <div className="bg-white p-6 rounded-xl shadow-sm border border-slate-200 flex flex-col h-full">
              <div className="flex flex-col mb-4 border-b border-slate-200 pb-3 gap-2">
                <h2 className="ds-sec-title is-inline whitespace-nowrap">
                  <span className="ds-sec-num">4</span>
                  측정점별 동압 및 온도
                </h2>
                <div className="flex flex-wrap items-center gap-2">
                  <div className="flex items-center gap-1.5 bg-emerald-50 px-2 py-1.5 rounded-lg border border-emerald-200 shadow-sm">
                    <label className="text-[11px] font-bold text-emerald-800 whitespace-nowrap">피토계수</label>
                    <input
                      type="number"
                      step="0.01"
                      name="pitotFactor"
                      value={formData.pitotFactor}
                      onChange={handleChange}
                      readOnly={isPitotCpAutoApplied}
                      title={isPitotCpAutoApplied ? '유속별 Cp 표에서 자동으로 정합니다. 직접 입력하려면 피토 프리셋을 해제하세요.' : undefined}
                      className={`w-20 p-1 border border-emerald-300 rounded text-xs text-center outline-none focus:ring-2 focus:ring-emerald-500 ${isPitotCpAutoApplied ? 'bg-emerald-100 font-black text-emerald-900' : ''}`}
                    />
                    {isPitotCpAutoApplied && (
                      <span className="text-[10px] text-emerald-700 font-bold whitespace-nowrap">
                        {Number.isFinite(appliedPitotVelocity) && appliedPitotVelocity > 0
                          ? `${appliedPitotVelocitySource} 유속 ${appliedPitotVelocity.toFixed(2)} m/s 기준으로 자동 선정`
                          : '동압 입력 시 자동 선정'}
                      </span>
                    )}
                    <select
                      value={formData.pitotPresetId}
                      onChange={handlePitotPresetChange}
                      className="p-1 border border-emerald-300 rounded text-[11px] font-bold text-emerald-800 bg-white outline-none focus:ring-2 focus:ring-emerald-500 cursor-pointer"
                      title="피토 프리셋 선택"
                    >
                      <option value="">피토 프리셋</option>
                      {sheetPitotPresets
                        .map((preset) => (
                          <option key={preset.id} value={preset.id}>
                            {(preset.sn ? `S/N ${preset.sn}` : `S/N 미입력 (${preset.id})`)
                              + (preset.usage || preset.length ? ` · ${[preset.usage, preset.length ? `${preset.length}cm` : ''].filter(Boolean).join(' ')}` : '')}
                          </option>
                        ))}
                    </select>
                  </div>
                  <div className="flex items-center gap-1.5 bg-emerald-50 px-2 py-1.5 rounded-lg border border-emerald-200 shadow-sm">
                    <label className="text-[11px] font-bold text-emerald-800">미터온도(T<sub>m</sub>)</label>
                    <input type="number" step="0.1" name="traverseTmIn" value={formData.traverseTmIn} onChange={handleChange} placeholder="입구" className="w-14 p-1 border border-emerald-300 rounded text-xs text-center outline-none focus:ring-2 focus:ring-emerald-500" />
                    <span className="text-slate-400 font-black">/</span>
                    <input type="number" step="0.1" name="traverseTmOut" value={formData.traverseTmOut} onChange={handleChange} placeholder="출구" className="w-14 p-1 border border-emerald-300 rounded text-xs text-center outline-none focus:ring-2 focus:ring-emerald-500" />
                  </div>
                </div>
              </div>

              {isPitotCpAutoApplied && (() => {
                const sel = pitotAutoSelection;
                const fmtV = (v) => `${v.toFixed(2)} m/s`;
                return (
                  <div className="mb-3 border-l-2 border-emerald-600 pl-3 text-[11px] text-slate-700 leading-5">
                    <div>
                      <span className="font-black text-slate-900">피토계수 자동 선정</span>
                      <span className="text-slate-500"> · {activePitotPreset?.sn ? `S/N ${activePitotPreset.sn}` : '선택 피토관'}</span>
                      {sel?.ready && (
                        <span className="text-slate-500"> · {sel.source} 유속 기준</span>
                      )}
                    </div>
                    {sel?.ready ? (
                      <div>
                        {sel.steps.map((step, idx) => (
                          <span key={idx}>
                            {idx > 0 && <span className="text-slate-400"> → </span>}
                            {idx === 0
                              ? <>기준계수 {PITOT_BASE_CP} → 유속 {fmtV(step.velocity)}</>
                              : <>{step.note} C<sub>p</sub> <b className="text-slate-900">{formatPitotCp(step.cp)}</b> → 유속 {fmtV(step.velocity)}</>}
                          </span>
                        ))}
                        <span className="font-black text-emerald-800"> · {sel.converged ? '수렴' : '반복 한도 도달'}, C<sub>p</sub> {formatPitotCp(sel.finalCp)} 적용</span>
                      </div>
                    ) : (
                      <div className="text-slate-600">동압·온도·대기압·가스 조성을 입력하면 기준계수 {PITOT_BASE_CP}로 유속을 먼저 구하고, 그 유속 구간의 C<sub>p</sub>를 고릅니다. 지금은 공칭값 {formData.pitotFactor}을 씁니다.</div>
                    )}
                    {sel?.ready && !sel.inRange && (
                      <div className="font-bold text-amber-800">
                        유속 {fmtV(sel.finalVelocity)}는 이 피토관 표의 최대 유속({sel.maxV} m/s)을 넘어 마지막 구간의 값을 적용했습니다.
                        {sel.alternative ? (
                          <>
                            {' '}범위에 맞는 피토관: S/N {sel.alternative.sn || sel.alternative.id}{sel.alternative.usage ? ` (${sel.alternative.usage})` : ''}
                            {' '}<button type="button" onClick={() => applyPitotPresetById(sel.alternative.id)} className="underline text-amber-900 hover:text-amber-700">이 피토관으로 변경</button>
                          </>
                        ) : ' 범위에 맞는 다른 피토관 프리셋이 없습니다.'}
                      </div>
                    )}
                  </div>
                );
              })()}
              <details className="mb-3 bg-emerald-50 border border-emerald-200 rounded-lg p-2 shadow-sm">
                <summary className="cursor-pointer text-[11px] font-bold text-emerald-800">
                  피토 프리셋 직접 편집 (C<sub>p</sub>)
                </summary>
                <div className="mt-2 flex items-center gap-2">
                  <button type="button" onClick={savePitotPresets} className="px-2 py-1 text-[11px] font-bold bg-emerald-600 text-white rounded border border-emerald-700 inline-flex items-center gap-1">
                    <PresetSaveIcon />
                    저장
                  </button>
                  <button type="button" onClick={addPitotPreset} className="px-2 py-1 text-[11px] font-bold bg-emerald-100 text-emerald-800 rounded border border-emerald-300">행 추가</button>
                  <button type="button" onClick={resetPitotDefaults} className="px-2 py-1 text-[11px] font-bold bg-slate-100 text-slate-700 rounded border border-slate-300">{IS_DESKTOP_BUILD ? '비우기' : '초기값 복원'}</button>
                </div>
                <div className="mt-2 overflow-x-auto">
                  <table className="w-full text-[11px] text-center border border-emerald-200">
                    <thead className="bg-emerald-100 text-slate-800">
                      <tr>
                        <th className="border border-emerald-200 p-1">S/N</th>
                        <th className="border border-emerald-200 p-1">기본 C<sub>p</sub></th>
                        <th className="border border-emerald-200 p-1">유속별 C<sub>p</sub> (교정 성적서)</th>
                        <th className="border border-emerald-200 p-1">삭제</th>
                      </tr>
                    </thead>
                    <tbody>
                      {pitotPresets.map((preset, idx) => (
                        <tr key={preset.id} className="bg-white">
                          <td className="border border-emerald-200 p-1">
                            <input
                              type="text"
                              value={preset.sn}
                              onChange={(e) => handlePitotPresetConfigChange(idx, 'sn', e.target.value)}
                              className="w-full p-1 border border-emerald-300 rounded text-center"
                              placeholder="피토관 S/N"
                            />
                            <div className="mt-1 flex items-center gap-1">
                              <input
                                type="text"
                                value={preset.usage || ''}
                                onChange={(e) => handlePitotPresetConfigChange(idx, 'usage', e.target.value)}
                                className="w-14 p-1 border border-emerald-200 rounded text-center text-[10px]"
                                placeholder="용도"
                                title="용도 (입자 / 수은 / PAHs)"
                              />
                              <input
                                type="number"
                                step="1"
                                value={preset.length || ''}
                                onChange={(e) => handlePitotPresetConfigChange(idx, 'length', e.target.value)}
                                className="w-14 p-1 border border-emerald-200 rounded text-center text-[10px]"
                                placeholder="cm"
                                title="유효 길이 (cm)"
                              />
                              <input
                                type="date"
                                value={preset.calDate || ''}
                                onChange={(e) => handlePitotPresetConfigChange(idx, 'calDate', e.target.value)}
                                className="p-1 border border-emerald-200 rounded text-center text-[10px]"
                                title="검교정일"
                              />
                            </div>
                          </td>
                          <td className="border border-emerald-200 p-1">
                            <input
                              type="number"
                              step="0.01"
                              value={preset.factor}
                              onChange={(e) => handlePitotPresetConfigChange(idx, 'factor', e.target.value)}
                              className="w-16 min-w-[4.5rem] p-1 border border-emerald-300 rounded text-center"
                            />
                          </td>
                          <td className="border border-emerald-200 p-1 text-left">
                            <div className="flex flex-col gap-1">
                              {(Array.isArray(preset.cpByVelocity) ? preset.cpByVelocity : []).map((row, rowIdx) => (
                                <div key={`${preset.id}-cp-${rowIdx}`} className="flex items-center gap-1">
                                  <input
                                    type="number"
                                    step="0.1"
                                    value={row.v}
                                    onChange={(e) => handlePitotCpRowChange(idx, rowIdx, 'v', e.target.value)}
                                    className="w-16 p-1 border border-emerald-300 rounded text-center"
                                    placeholder="m/s"
                                  />
                                  <span className="text-slate-500">m/s →</span>
                                  <input
                                    type="number"
                                    step="0.001"
                                    value={row.cp}
                                    onChange={(e) => handlePitotCpRowChange(idx, rowIdx, 'cp', e.target.value)}
                                    className="w-16 p-1 border border-emerald-300 rounded text-center"
                                    placeholder="Cp"
                                  />
                                  <button type="button" onClick={() => removePitotCpRow(idx, rowIdx)} className="text-red-500 hover:text-red-700 font-bold px-1" title="구간 삭제">×</button>
                                </div>
                              ))}
                              <button type="button" onClick={() => addPitotCpRow(idx)} className="self-start px-2 py-0.5 bg-emerald-100 text-emerald-800 rounded border border-emerald-300 font-bold">
                                + 유속 구간 추가
                              </button>
                              <span className="text-[10px] text-slate-500">유속 값은 구간 상한 (4.95 → 0~4.95 m/s). 표 범위를 넘으면 마지막 Cp, 표가 비면 기본 Cp.</span>
                            </div>
                          </td>
                          <td className="border border-emerald-200 p-1">
                            <button type="button" onClick={() => removePitotPreset(idx)} className="text-red-500 hover:text-red-700 text-[11px] font-bold p-1">
                              삭제
                            </button>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </details>
              
              <div className="flex-1 overflow-auto bg-slate-50/50 p-2 rounded border border-slate-200 mb-4">
                <table className="fill-drag-table w-full text-sm text-center">
                  <thead>
                    <tr className="text-slate-700 border-b border-slate-300 bg-slate-100/50">
                      <th className="pb-2 pt-2 rounded-tl-lg font-bold">측정점</th>
                      <th className="pb-2 pt-2 text-xs font-bold">전압<br/>(mmH2O)</th>
                      <th className="pb-2 pt-2 text-xs font-bold">정압<br/>(mmH2O)</th>
                      <th className="pb-2 pt-2 text-xs font-bold">동압<br/>(mmH2O)</th>
                      <th className="pb-2 pt-2 text-xs font-bold">온도(℃)</th>
                      {isFlowSheet && <th className="pb-2 pt-2 text-xs font-bold">유속<br/>(m/s)</th>}
                      {isFlowSheet && <th className="pb-2 pt-2 text-xs font-bold">유량<br/>(㎥/min, 실측)</th>}
                      <th className="pb-2 pt-2 rounded-tr-lg"></th>
                    </tr>
                  </thead>
                  <tbody>
                    {formData.points.map((point, idx) => (
                      <tr key={point.id} className="hover:bg-slate-100/30">
                        <td className="py-1 font-bold text-slate-700">{point.id}</td>
                        <td className="py-1"><input type="number" step="0.1" value={point.tp !== undefined ? point.tp : ''} onChange={(e) => handlePointChange(idx, 'tp', e.target.value)} className="w-16 p-1 border border-slate-300 rounded text-center text-sm outline-none focus:ring-2 focus:ring-emerald-500" /></td>
                        <td className="py-1"><input type="number" step="0.1" value={point.sp !== undefined ? point.sp : ''} onChange={(e) => handlePointChange(idx, 'sp', e.target.value)} className="w-16 p-1 border border-slate-300 rounded text-center text-sm outline-none focus:ring-2 focus:ring-emerald-500" /></td>
                        <td className="py-1"><input type="number" step="0.1" value={point.dp} onChange={(e) => handlePointChange(idx, 'dp', e.target.value)} className="w-16 p-1 border border-emerald-300 bg-emerald-50 rounded text-center font-black text-emerald-700 text-sm outline-none focus:ring-2 focus:ring-emerald-500" /></td>
                        <td className="py-1"><input type="number" step="0.1" value={point.ts} onChange={(e) => handlePointChange(idx, 'ts', e.target.value)} className="w-16 p-1 border border-slate-300 rounded text-center text-sm outline-none focus:ring-2 focus:ring-emerald-500" /></td>
                        {isFlowSheet && <td className="py-1 font-bold text-slate-900">{fmtOrDash(getRawPointGasVelocity(point))}</td>}
                        {isFlowSheet && <td className="py-1 font-bold text-slate-900">{fmtOrDash(getRawPointFlowRate(point))}</td>}
                        <td className="py-1">
                          <button type="button" onClick={() => removePoint(idx)} className="text-red-500 hover:text-red-700 text-xs font-bold p-1 transition-colors">삭제</button>
                        </td>
                      </tr>
                    ))}
                    {/* 전압/정압/동압/온도 평균 행 표시 */}
                    <tr className="bg-emerald-50/50 border-t-2 border-emerald-200 text-emerald-900 font-bold">
                        <td className="py-1.5">평균</td>
                        <td className="py-1.5">{calcAvgTp() !== 0 ? calcAvgTp().toFixed(2) : '-'}</td>
                        <td className="py-1.5">{calcAvgSp() !== 0 ? calcAvgSp().toFixed(2) : '-'}</td>
                        <td className="py-1.5 text-emerald-700">{calcAvgDp() !== 0 ? calcAvgDp().toFixed(1) : '-'}</td>
                        <td className="py-1.5">{calcAvgTs() !== 0 ? calcAvgTs().toFixed(2) : '-'}</td>
                        {isFlowSheet && <td className="py-1.5 text-emerald-700">{getRawGasVelocity() > 0 ? getRawGasVelocity().toFixed(2) : '-'}</td>}
                        {isFlowSheet && <td className="py-1.5 text-emerald-700">{getRawGasVelocity() > 0 && Number.isFinite(getStackAreaM2()) ? (getRawGasVelocity() * getStackAreaM2() * 60).toFixed(2) : '-'}</td>}
                        <td className="py-1.5"></td>
                    </tr>
                  </tbody>
                </table>
                <button type="button" onClick={addPoint} className="w-full mt-2 py-1.5 text-emerald-700 bg-emerald-100 hover:bg-emerald-200 rounded text-xs font-bold border border-emerald-300 transition-colors shadow-sm">
                  측정점 추가
                </button>
              </div>

              <div className="bg-emerald-50 p-4 rounded-lg flex justify-between items-center border border-emerald-200 shadow-sm">
                <div className="text-slate-800">
                  <span className="font-bold">평균 유속 (Vs)</span>
                </div>
                <div className="text-right">
                  <div className="text-2xl font-black text-emerald-700">{calcGasVelocity()} <span className="text-sm font-bold text-slate-700">m/s</span></div>
                </div>
              </div>
            </div>
          </div>

          {/* 섹션 5: 시료채취 및 등속흡인 (유량측정기록부에서는 사용하지 않음) */}
          {!isFlowSheet && (
          <div className="bg-white p-6 rounded-xl shadow-sm border border-slate-200">
            <h2 className="ds-sec-title">
              <span className="ds-sec-num">5</span>
              시료 채취 기록
            </h2>

            {!isFlowSheet && (<>
            {/* 1단계: 자동 산정 및 예측 결과 통합 패널 */}
            <div className="bg-emerald-50 border border-emerald-200 p-5 rounded-lg mb-6 shadow-sm">
               <div className="flex flex-col md:flex-row md:justify-between md:items-end mb-4 gap-2">
                 <h3 className="font-bold text-emerald-800">
                    1단계 채취 조건
                 </h3>
                 <div className="flex items-center gap-2 bg-white px-3 py-1.5 rounded-lg border border-emerald-200 shadow-sm">
                    <label className="text-xs font-bold text-slate-700">장비 프리셋 선택</label>
                    <select 
                       value={formData.samplerId} 
                       onChange={handleSamplerChange}
                       className="p-1 border border-emerald-300 rounded text-sm font-bold text-emerald-800 bg-emerald-50 focus:ring-2 focus:ring-emerald-500 outline-none cursor-pointer"
                    >
                       <option value="">직접 입력</option>
                       {samplers.map(s => (
                          <option key={s.id} value={s.id}>{s.name}</option>
                       ))}
                    </select>
                 </div>
               </div>

               <details className="mb-3 bg-emerald-50 border border-emerald-200 rounded-lg p-2 shadow-sm">
                 <summary className="cursor-pointer text-[11px] font-bold text-emerald-800">
                   샘플러 프리셋 직접 편집 (Yd / ΔH@)
                 </summary>
                 <div className="mt-2 flex items-center gap-2">
                   <button type="button" onClick={saveSamplerPresets} className="px-2 py-1 text-[11px] font-bold bg-emerald-600 text-white rounded border border-emerald-700 inline-flex items-center gap-1">
                     <PresetSaveIcon />
                     저장
                   </button>
                   <button type="button" onClick={addSamplerPreset} className="px-2 py-1 text-[11px] font-bold bg-emerald-100 text-emerald-800 rounded border border-emerald-300">행 추가</button>
                   <button type="button" onClick={resetSamplerDefaults} className="px-2 py-1 text-[11px] font-bold bg-slate-100 text-slate-700 rounded border border-slate-300">{IS_DESKTOP_BUILD ? '비우기' : '초기값 복원'}</button>
                 </div>
                 <div className="mt-2 overflow-x-auto">
                   <table className="w-full text-[11px] text-center border border-emerald-200">
                     <thead className="bg-emerald-100 text-slate-800">
                       <tr>
                         <th className="border border-emerald-200 p-1">샘플러명</th>
                         <th className="border border-emerald-200 p-1">Yd</th>
                         <th className="border border-emerald-200 p-1">ΔH@</th>
                         <th className="border border-emerald-200 p-1">교정일</th>
                         <th className="border border-emerald-200 p-1">삭제</th>
                       </tr>
                     </thead>
                     <tbody>
                       {samplers.map((s, idx) => (
                         <tr key={s.id} className="bg-white">
                           <td className="border border-emerald-200 p-1">
                             <input
                               type="text"
                               value={s.name}
                               onChange={(e) => handleSamplerPresetChange(idx, 'name', e.target.value)}
                               className="w-full p-1 border border-emerald-300 rounded text-center"
                             />
                           </td>
                           <td className="border border-emerald-200 p-1">
                             <input
                               type="number"
                               step="0.001"
                               value={s.yd}
                               onChange={(e) => handleSamplerPresetChange(idx, 'yd', e.target.value)}
                               className="w-full p-1 border border-emerald-300 rounded text-center"
                             />
                           </td>
                           <td className="border border-emerald-200 p-1">
                             <input
                               type="number"
                               step="0.0001"
                               value={s.dHAt}
                               onChange={(e) => handleSamplerPresetChange(idx, 'dHAt', e.target.value)}
                               className="w-full p-1 border border-emerald-300 rounded text-center"
                             />
                           </td>
                           <td className="border border-emerald-200 p-1">
                             <input
                               type="date"
                               value={s.calDate || ''}
                               onChange={(e) => handleSamplerPresetChange(idx, 'calDate', e.target.value)}
                               className="w-full p-1 border border-emerald-300 rounded text-center"
                               title="검교정일"
                             />
                           </td>
                           <td className="border border-emerald-200 p-1">
                             <button type="button" onClick={() => removeSamplerPreset(idx)} className="text-red-500 hover:text-red-700 text-[11px] font-bold p-1">
                               삭제
                             </button>
                           </td>
                         </tr>
                       ))}
                     </tbody>
                   </table>
                 </div>
               </details>
               
               <div className="grid grid-cols-1 md:grid-cols-3 gap-4 mb-3 items-end">
                 <div>
                    <label className="block text-xs font-bold text-emerald-800 mb-1">목표 채취량 (SL)</label>
                    <input type="number" step="1" name="targetVolume" value={formData.targetVolume} onChange={handleChange} className="w-full p-2 border border-emerald-300 rounded text-sm bg-white font-bold outline-none focus:ring-2 focus:ring-emerald-500" />
                    {IS_DESKTOP_BUILD && selectedSheet === 'dust' && (
                      <div className="mt-1 flex items-center gap-1" role="group" aria-label="목표 채취량 고르기">
                        {['400', '1000'].map((volume) => (
                          <button
                            key={volume}
                            type="button"
                            aria-pressed={String(formData.targetVolume) === volume}
                            onClick={() => { setFormData((prev) => ({ ...prev, targetVolume: volume })); unsavedRef.current = true; }}
                            className={`flex-1 px-2 py-1 rounded border text-xs font-bold ${String(formData.targetVolume) === volume ? 'bg-emerald-600 text-white border-emerald-700' : 'bg-white text-slate-700 border-slate-300 hover:bg-slate-100'}`}
                          >
                            {volume} L
                          </button>
                        ))}
                      </div>
                    )}
                 </div>
                 <div>
                    <label className="block text-xs font-bold text-emerald-800 mb-1" title="가스미터 보정계수">보정계수 (Y<sub>d</sub>)</label>
                    <input type="number" step="0.001" name="gasMeterFactor" value={formData.gasMeterFactor} onChange={handleChange} placeholder="Yd" className="w-full p-2 border border-emerald-300 rounded text-sm bg-white outline-none focus:ring-2 focus:ring-emerald-500 font-bold text-emerald-700" />
                 </div>
                 <div>
                    <label className="block text-xs font-bold text-emerald-800 mb-1" title="오리피스 계수">오리피스 계수 (ΔH<sub>@</sub>)</label>
                    <input type="number" step="0.0001" name="deltaHAt" value={formData.deltaHAt} onChange={handleChange} placeholder="ΔH@" className="w-full p-2 border border-emerald-300 rounded text-sm bg-white outline-none focus:ring-2 focus:ring-emerald-500 font-bold text-emerald-700" />
                 </div>
               </div>

               <div className="mb-4">
                  <button type="button" onClick={handleCalculateOptions} className="ds-btn-primary w-full text-center">
                    K-Factor와 예상 채취시간 계산
                  </button>
                  <div className="mt-2 flex flex-wrap items-center gap-x-4 gap-y-2 text-[11px] font-bold text-slate-600">
                    <span>추천 기준 오리피스압 ΔH (mmH₂O)</span>
                    {[['max', '빠른 채취 상한'], ['standard', '표준 채취 목표'], ['stable', '안정/장시간 목표']].map(([key, label]) => (
                      <label key={key} className="inline-flex items-center gap-1">
                        {label}
                        <input type="number" step="1" min="1" value={dhTargets[key]} onChange={(e) => changeDhTarget(key, e.target.value)} placeholder={String(DEFAULT_DH_TARGETS[key])} className="w-16 p-1 border border-emerald-300 rounded text-center text-xs bg-white outline-none focus:ring-2 focus:ring-emerald-500" />
                      </label>
                    ))}
                  </div>
               </div>
               
               {recommendations && (
                 <div className="grid grid-cols-1 md:grid-cols-3 gap-4 mb-5">
                    <div className="bg-white border-2 border-emerald-300 p-4 rounded-lg shadow-sm hover:border-emerald-500 transition-colors">
                       <h4 className="font-bold text-emerald-800 mb-3 flex items-center gap-1 border-b border-emerald-100 pb-2">빠른 채취 <span className="text-[10px] font-bold text-emerald-600 ml-auto">(최대치, {getDhTargets().max} 이하)</span></h4>
                       <ul className="text-sm text-slate-700 space-y-2 mb-4">
                          <li className="flex justify-between font-bold"><span>추천 노즐:</span> <span className="font-black text-lg text-slate-900">No.{recommendations.fast.bestNozzleNum} <span className="text-sm font-medium">({recommendations.fast.finalDn}mm)</span></span></li>
                          <li className="flex justify-between font-bold"><span>예상 시간:</span> <span className="font-black text-emerald-700">{recommendations.fast.fastestTime} 분</span></li>
                          <li className="flex justify-between font-bold"><span>K-Factor:</span> <span className="font-black">{formatKFactorDisplay(recommendations.fast.calculatedK)}</span></li>
                          <li className="flex justify-between items-center bg-emerald-50 p-1 rounded border border-emerald-100 font-bold"><span>예상 ΔH:</span> <span className="font-black text-emerald-700 text-lg">{recommendations.fast.expectedDH}</span></li>
                       </ul>
                       <button type="button" onClick={() => applyRecommendation(recommendations.fast)} className="w-full py-2 bg-emerald-600 text-white rounded-lg text-sm font-bold hover:bg-emerald-700 shadow-sm">적용하기</button>
                    </div>
                    <div className="bg-teal-50 border-2 border-teal-400 p-4 rounded-lg shadow-sm hover:border-teal-500 transition-colors">
                       <h4 className="font-bold text-teal-900 mb-3 flex items-center gap-1 border-b border-teal-200 pb-2">표준 채취 <span className="text-[10px] font-bold text-teal-600 ml-auto">(ΔH {getDhTargets().standard} 내외)</span></h4>
                       <ul className="text-sm text-slate-700 space-y-2 mb-4">
                          <li className="flex justify-between font-bold"><span>추천 노즐:</span> <span className="font-black text-lg text-slate-900">No.{recommendations.standard.bestNozzleNum} <span className="text-sm font-medium">({recommendations.standard.finalDn}mm)</span></span></li>
                          <li className="flex justify-between font-bold"><span>예상 시간:</span> <span className="font-black text-teal-700">{recommendations.standard.fastestTime} 분</span></li>
                          <li className="flex justify-between font-bold"><span>K-Factor:</span> <span className="font-black">{formatKFactorDisplay(recommendations.standard.calculatedK)}</span></li>
                          <li className="flex justify-between items-center bg-white p-1 rounded border border-teal-200 font-bold"><span>예상 ΔH:</span> <span className="font-black text-teal-700 text-lg">{recommendations.standard.expectedDH}</span></li>
                       </ul>
                       <button type="button" onClick={() => applyRecommendation(recommendations.standard)} className="w-full py-2 bg-teal-600 text-white rounded-lg text-sm font-bold hover:bg-teal-700 shadow-sm">적용하기</button>
                    </div>
                    <div className="bg-cyan-50 border-2 border-cyan-300 p-4 rounded-lg shadow-sm hover:border-cyan-500 transition-colors">
                       <h4 className="font-bold text-cyan-900 mb-3 flex items-center gap-1 border-b border-cyan-200 pb-2">안정/장시간 <span className="text-[10px] font-bold text-cyan-600 ml-auto">(ΔH {getDhTargets().stable} 내외)</span></h4>
                       <ul className="text-sm text-slate-700 space-y-2 mb-4">
                          <li className="flex justify-between font-bold"><span>추천 노즐:</span> <span className="font-black text-lg text-slate-900">No.{recommendations.stable.bestNozzleNum} <span className="text-sm font-medium">({recommendations.stable.finalDn}mm)</span></span></li>
                          <li className="flex justify-between font-bold"><span>예상 시간:</span> <span className="font-black text-cyan-700">{recommendations.stable.fastestTime} 분</span></li>
                          <li className="flex justify-between font-bold"><span>K-Factor:</span> <span className="font-black">{formatKFactorDisplay(recommendations.stable.calculatedK)}</span></li>
                          <li className="flex justify-between items-center bg-white p-1 rounded border border-cyan-200 font-bold"><span>예상 ΔH:</span> <span className="font-black text-cyan-700 text-lg">{recommendations.stable.expectedDH}</span></li>
                       </ul>
                       <button type="button" onClick={() => applyRecommendation(recommendations.stable)} className="w-full py-2 bg-cyan-600 text-white rounded-lg text-sm font-bold hover:bg-cyan-700 shadow-sm">적용하기</button>
                    </div>
                 </div>
               )}
               
               

               {/* 접기/펴기 엑셀 전체 리스트 */}
               <details className="mt-4 mb-4 bg-emerald-50 border border-emerald-200 rounded-lg p-2 shadow-sm">
                  <summary className="cursor-pointer text-xs font-bold text-emerald-700 hover:text-emerald-900 flex items-center justify-between gap-2">
                     <span>노즐 예측표</span>
                     <span className="text-[10px] text-slate-500">원하는 노즐 직접 선택 · 번호와 직경 편집</span>
                  </summary>
                  <div className="mt-2 flex items-center gap-2">
                    <span className="text-[11px] font-bold text-emerald-800 bg-emerald-50 border border-emerald-200 rounded px-2 py-1">{NOZZLE_GROUP_LABELS[activeNozzleGroup]} 노즐 세트</span>
                    <button type="button" onClick={saveNozzlePresets} className="px-2 py-1 text-[11px] font-bold bg-emerald-600 text-white rounded border border-emerald-700 inline-flex items-center gap-1">
                      <PresetSaveIcon />
                      저장
                    </button>
                    <button type="button" onClick={addNozzleConfig} className="px-2 py-1 text-[11px] font-bold bg-emerald-100 text-emerald-800 rounded border border-emerald-300">행 추가</button>
                    <button type="button" onClick={resetNozzleDefaults} className="px-2 py-1 text-[11px] font-bold bg-slate-100 text-slate-700 rounded border border-slate-300">{IS_DESKTOP_BUILD ? '비우기' : '초기값 복원'}</button>
                  </div>
                  <div className="mt-3 overflow-x-auto">
                     <table className="w-full text-xs text-center border-collapse border border-emerald-300">
                       <thead className="bg-emerald-100 text-slate-800">
                          <tr>
                             <th className="border border-emerald-200 p-1 font-bold">노즐번호</th>
                             <th className="border border-emerald-200 p-1 font-bold">노즐직경(mm)</th>
                             <th className="border border-emerald-200 p-1 font-bold">K-factor</th>
                             <th className="border border-emerald-200 p-1 font-bold">ΔH</th>
                             <th className="border border-emerald-200 p-1 font-bold text-emerald-800">예상시간(분)</th>
                             <th className="border border-emerald-200 p-1 font-bold">(L)</th>
                             <th className="border border-emerald-200 p-1 font-bold">(SL)</th>
                             <th className="border border-emerald-200 p-1 font-bold">(Sm³)</th>
                             <th className="border border-emerald-200 p-1 font-bold">사용</th>
                             <th className="border border-emerald-200 p-1 font-bold">삭제</th>
                          </tr>
                       </thead>
                       <tbody className="bg-white">
                          {activeNozzleSet.map((n, idx) => {
                             let k='-', dh='-', vl='-', vsl='-', vsm='-', reqTime='-', choice=null;
                             const nozzleNum = parseInt(n.num, 10);
                             const Dn = parseFloat(n.d);
                             const vs = getRawGasVelocity(), cP = parseFloat(formData.pitotFactor)||0.84, ts = getRawAvgTs(), pm = parseFloat(formData.atmPressure), ps = getRawStackPressure();
                             const tm = getTraverseMeterTempAvg();
                             
                             const validSps = formData.points.map(p => parseFloat(p.sp)).filter(v => !isNaN(v));
                             const avgSp = validSps.length === 0 ? 0 : validSps.reduce((a, b) => a + b, 0) / validSps.length;
                             const ps_ratio = (pm + avgSp / 13.6) / pm;

                             const gasComp = getGasComposition();
                             const Md = gasComp.Md, Ms = gasComp.Ms, Xw = gasComp.Xw;
                             const y = parseFloat(formData.gasMeterFactor)||1.0;
                             const dHAt = parseFloat(formData.deltaHAt);

                             if(!isNaN(nozzleNum) && !isNaN(Dn) && Dn > 0 && vs > 0 && Number.isFinite(ts) && Number.isFinite(tm) && Number.isFinite(ps) && Number.isFinite(pm) && pm > 0 && Number.isFinite(Md) && Number.isFinite(Ms) && Ms !== 0 && Number.isFinite(Xw)) {
                                const k_val = 8.001e-5 * Math.pow(Dn, 4) * dHAt * Math.pow(cP, 2) * Math.pow(1 - Xw / 100, 2) * (Md / Ms) * ((tm + 273) / (ts + 273)) * ps_ratio;
                                k = formatKFactorDisplay(k_val);
                                dh = calcExpectedOrificeDH(k_val, getRawAvgDp());
                                const an = Math.PI * Math.pow(Dn/2000, 2);
                                const q_m = (vs * an * 60000 * ((tm+273)/(ts+273)) * (ps/pm) * (1-Xw/100)) / y;
                                const q_sl = q_m * y * (273/(tm+273)) * (pm/760);
                                
                                let time = Math.ceil((parseFloat(formData.targetVolume) || 1000) / q_sl);
                                if (!isFinite(time) || time <= 0) time = 0;
                                reqTime = time;
                                if (Number.isFinite(k_val) && time > 0) choice = { bestNozzleNum: nozzleNum, finalDn: Dn.toFixed(2), fastestTime: time, calculatedK: k_val };
                                
                                const v_val = q_m * time;
                                vl = v_val.toFixed(2);
                                vsl = (q_sl * time).toFixed(2);
                                vsm = ((q_sl * time) / 1000).toFixed(3);
                             }
                             const isSelected = !isNaN(nozzleNum) && formData.usedNozzleNum === String(nozzleNum);
                             return (
                               <tr key={`${idx}-${n.num}-${n.d}`} className={isSelected ? "bg-emerald-100 font-bold" : "hover:bg-emerald-50"}>
                                 <td className="border border-emerald-200 p-1 font-bold text-slate-800">
                                   <div className="flex items-center justify-center gap-1">
                                     <input
                                       type="number"
                                       step="1"
                                       value={n.num}
                                       onChange={(e) => handleNozzleConfigChange(idx, 'num', e.target.value)}
                                       className="w-16 p-1 border border-emerald-300 rounded text-center font-bold"
                                       placeholder="번호"
                                     />
                                   </div>
                                   {!isNaN(nozzleNum) && formData.recommendedNozzleNum === String(nozzleNum) && (
                                     <span className="ml-1 text-[10px] bg-teal-600 text-white px-1.5 py-0.5 rounded-full shadow-sm">추천</span>
                                   )}
                                 </td>
                                 <td className="border border-emerald-200 p-1 font-bold">
                                   <input
                                     type="number"
                                     step="0.01"
                                     value={n.d}
                                     onChange={(e) => handleNozzleConfigChange(idx, 'd', e.target.value)}
                                     className="w-20 p-1 border border-emerald-300 rounded text-center font-bold"
                                     placeholder="직경"
                                   />
                                 </td>
                                 <td className="border border-emerald-200 p-1 font-black text-slate-700">{k}</td>
                                 <td className="border border-emerald-200 p-1 font-black text-emerald-700">{dh}</td>
                                 <td className="border border-emerald-200 p-1 font-black text-teal-800 bg-teal-50/50">{reqTime}</td>
                                 <td className="border border-emerald-200 p-1">{vl}</td>
                                 <td className="border border-emerald-200 p-1">{vsl}</td>
                                 <td className="border border-emerald-200 p-1">{vsm}</td>
                                 <td className="border border-emerald-200 p-1">
                                   {isSelected
                                     ? <span className="text-[11px] font-black text-emerald-700">사용 중</span>
                                     : <button type="button" disabled={!choice} onClick={() => applyRecommendation(choice, true)} className="px-2 py-1 text-[11px] font-bold rounded border border-emerald-600 bg-emerald-600 text-white hover:bg-emerald-700 disabled:opacity-40 disabled:cursor-not-allowed">선택</button>}
                                 </td>
                                 <td className="border border-emerald-200 p-1">
                                   <button type="button" onClick={() => removeNozzleConfig(idx)} className="text-red-500 hover:text-red-700 text-[11px] font-bold p-1">
                                     삭제
                                   </button>
                                 </td>
                               </tr>
                             )
                          })}
                          {activeNozzleSet.length === 0 && (
                            <tr>
                              <td className="border border-emerald-200 p-2 text-slate-500" colSpan={10}>
                                노즐이 없습니다. "행 추가"로 입력하세요.
                              </td>
                            </tr>
                          )}
                       </tbody>
                     </table>
                  </div>
               </details>

               {/* 산정 결과 요약창 */}
               <div className="bg-white p-4 border border-emerald-200 rounded-lg shadow-inner">
                  {formData.recommendedNozzleNum && (
                    <div className="mb-4 p-3 bg-teal-50 border border-teal-200 rounded-lg flex items-center justify-between shadow-sm">
                      <div className="flex items-center gap-2">
                        <span className="font-black text-teal-900">추천 노즐:</span>
                        <span className="text-lg font-black text-emerald-900 bg-emerald-200 px-3 py-1 border border-emerald-300 rounded-lg">
                          No.{formData.recommendedNozzleNum}{formData.recommendedNozzleDia ? ` (${formData.recommendedNozzleDia} mm)` : ''}
                        </span>
                      </div>
                    </div>
                  )}
                  
                  <h4 className="text-sm font-bold text-slate-800 mb-3 border-b border-slate-200 pb-2">노즐 선정 결과</h4>
                  <div className="ds-stats">
                      <div className="ds-stat">
                          <div className="ds-stat-label">가스유속(V<sub>s</sub>)</div>
                          <div className="ds-stat-value">{calcGasVelocity()}</div>
                      </div>
                      <div className="ds-stat">
                          <div className="ds-stat-label">노즐경(D<sub>n</sub>)</div>
                          <div className="ds-stat-value">{formData.nozzleDiameter || '-'}</div>
                      </div>
                      <div className="ds-stat">
                          <div className="ds-stat-label">K-Factor</div>
                          <div className="ds-stat-value">{formatKFactorDisplay(formData.kFactor)}</div>
                      </div>
                      <div className="ds-stat">
                          <div className="ds-stat-label">예상 ΔH</div>
                          <div className="ds-stat-value">{expData.dH}</div>
                      </div>
                      <div className="ds-stat">
                          <div className="ds-stat-label">예상시간(분)</div>
                          <div className="ds-stat-value">{expData.reqTime}</div>
                      </div>
                      <div className="ds-stat">
                          <div className="ds-stat-label">건조채취량(V<sub>m</sub>)</div>
                          <div className="ds-stat-value">{expData.L}</div>
                      </div>
                      <div className="ds-stat">
                          <div className="ds-stat-label">표준채취량(SL)</div>
                          <div className="ds-stat-value">{expData.SL}</div>
                      </div>
                      <div className="ds-stat">
                          <div className="ds-stat-label">표준채취량(Sm³)</div>
                          <div className="ds-stat-value">{expData.Sm3}</div>
                      </div>
                  </div>
               </div>
            </div>
            
            {/* 기본 채취 조건 */}
            <div className="grid grid-cols-1 md:grid-cols-2 gap-4 mb-6 bg-slate-50 p-4 rounded-xl border border-slate-200">
              <div>
                <label className="block text-xs font-bold text-slate-700 mb-1">사용 노즐경 (mm)</label>
                <input 
                  type="number" step="0.01" list="nozzle-list" name="nozzleDiameter" 
                  value={formData.nozzleDiameter} onChange={handleChange} 
                  placeholder="목록 선택 또는 입력"
                  className="w-full p-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-emerald-500 bg-white outline-none" 
                />
                <datalist id="nozzle-list">
                  {configuredNozzles.map(n => <option key={n.num} value={n.d.toFixed(2)}>No.{n.num} ({n.d}mm)</option>)}
                </datalist>
              </div>
              <div><label className="block text-xs font-bold text-slate-700 mb-1">적용 K-Factor</label>
                <input
                  type="number"
                  step="any"
                  name="kFactor"
                  value={isKFactorEditing ? formData.kFactor : (formData.kFactor === '' ? '' : formatKFactorDisplay(formData.kFactor))}
                  onFocus={() => setIsKFactorEditing(true)}
                  onBlur={() => setIsKFactorEditing(false)}
                  onChange={handleKFactorChange}
                  className="w-full p-2 border border-slate-300 bg-white rounded-lg focus:ring-2 focus:ring-emerald-500 font-black outline-none text-slate-800"
                />
              </div>
            </div>

            </>)}

            {/* 적산유량계 기록표 */}
            <div className="mb-6">
              <div className="flex flex-col md:flex-row md:justify-between md:items-center gap-2 mb-2">
                <h3 className="text-sm font-bold text-slate-800 whitespace-nowrap shrink-0">{isFlowSheet ? '측정점별 동압·온도 기록표' : '2단계 적산유량계 기록표'}</h3>
                <div className="flex flex-wrap items-center gap-2 justify-end">
                    <label className="flex items-center gap-1 text-[11px] font-bold text-slate-700 bg-slate-50 border border-slate-200 rounded-lg px-2 py-1" title="통합시료채취기록부의 시작·종료 시각에 쓰입니다 (종료 = 시작 + 총 시간)">
                      채취 시작
                      <input
                        type="time"
                        name="samplingStartTime"
                        value={formData.samplingStartTime || ''}
                        onChange={handleChange}
                        className="p-1 border border-slate-300 rounded text-center text-xs outline-none focus:ring-2 focus:ring-emerald-500 bg-white"
                      />
                      {formData.samplingStartTime && getSamplingMinutes() > 0 && (
                        <span className="text-slate-500 font-normal">~ {addMinutesToClock(formData.samplingStartTime, getSamplingMinutes())}</span>
                      )}
                    </label>
                    <label className="flex items-center gap-1 text-[11px] font-bold text-slate-700 bg-slate-50 border border-slate-200 rounded-lg px-2 py-1">
                      {isFlowSheet ? '총 측정시간' : '총 채취시간'}
                      <input
                        type="number"
                        step="1"
                        min="1"
                        value={tableTotalMinutes}
                        onChange={(e) => setTableTotalMinutes(e.target.value)}
                        placeholder={String(parseFloat(formData.planSamplingTime) > 0 ? formData.planSamplingTime : DEFAULT_MEASURE_TOTAL_MINUTES)}
                        className="w-16 p-1 border border-slate-300 rounded text-center text-xs outline-none focus:ring-2 focus:ring-emerald-500 bg-white"
                        title="비워 두면 1단계의 예상 채취시간을 쓰고, 없으면 60분으로 봅니다"
                      />
                      분
                    </label>
                    <button type="button" onClick={applySamplingPointsToTable} className="py-2 px-5 bg-emerald-600 text-white rounded-lg font-bold text-sm hover:bg-emerald-700 shadow-sm transition-colors border border-emerald-700">
                      표 다시 만들기
                    </button>
                    {(() => {
                      const plan = getTableGenerationPlan();
                      if (!plan) return <span className="w-full text-right text-[11px] text-amber-700 font-bold">굴뚝 내경(1번 섹션)을 입력하면 표가 만들어집니다</span>;
                      const changed = plan.adjustedTotal !== plan.requestedTotal;
                      return (
                        <span className={`w-full text-right text-[11px] font-bold ${changed ? 'text-amber-700' : 'text-emerald-700'}`}>
                          {plan.pointCount}지점 × 지점당 {plan.rowsPerPoint}행 × {plan.intervalMinutes}분 = 총 {plan.adjustedTotal}분 (지점당 {plan.perPointMinutes}분)
                          {changed ? ` · ${plan.requestedTotal}분은 ${plan.intervalMinutes}분 간격으로 나누어떨어지지 않아 ${plan.adjustedTotal}분으로 조정` : ''}
                        </span>
                      );
                    })()}
                  </div>
              </div>
              
              <div className="overflow-x-auto border border-slate-300 rounded-xl shadow-sm">
                <table className={`fill-drag-table w-full text-xs text-center ${isFlowSheet ? 'min-w-[560px]' : 'min-w-[900px]'}`}>
                  <thead className="bg-slate-100 text-slate-700 border-b border-slate-300">
                    <tr>
                      <th className="p-1 font-bold whitespace-nowrap">순번</th>
                      <th className="p-1 font-bold whitespace-nowrap">채취점</th>
                      <th className="p-1 font-bold whitespace-nowrap">누적 시간<br/>(분)</th>
                      <th className="p-1 font-bold whitespace-nowrap">배출가스<br/>온도(T<sub>s</sub>, ℃)</th>
                      <th className="p-1 font-bold text-emerald-700 whitespace-nowrap">동압<br/>(Δp)</th>
                      <th className="p-1 font-bold whitespace-nowrap">정압<br/>(mmH₂O)</th>
                      {!isFlowSheet && <th className="p-1 font-bold text-teal-700 whitespace-nowrap">오리피스압<br/>(ΔH)</th>}
                      {!isFlowSheet && <th className="p-1 font-bold whitespace-nowrap">적산유량<br/>(V<sub>m</sub>, L)</th>}
                      {!isFlowSheet && <th className="p-1 font-bold whitespace-nowrap">미터온도(T<sub>m</sub>, ℃)<br/><span className="text-[10px] text-slate-500">입구 | 출구</span></th>}
                      {!isFlowSheet && <th className="p-1 font-bold whitespace-nowrap">진공압<br/>(mmHg)</th>}
                      {!isFlowSheet && <th className="p-1 font-bold whitespace-nowrap">임핀저<br/>온도(℃)</th>}
                      <th className="p-1 font-bold whitespace-nowrap">배출가스<br/>속도(m/s)</th>
                      <th className="p-1 font-bold whitespace-nowrap">배출가스보정<br/>밀도(kg/m³)</th>
                      {!isFlowSheet && <th className="p-1 font-bold whitespace-nowrap">수분량<br/>(ml)</th>}
                      {!isFlowSheet && <th className="p-1 font-black text-emerald-600 border-l border-slate-300 whitespace-nowrap">순간 등속<br/>(I%)</th>}
                      <th className="p-1"></th>
                    </tr>
                  </thead>
                  <tbody className="bg-white">
                    {formData.gasMeters.map((meter, idx) => {
                      const isStartRow = idx === 0;
                      const rowRate = calcRowIsokineticRate(idx);
                      const rowVelocity = calcRowGasVelocity(idx);
                      const rowDensity = calcRowCorrectedGasDensity(idx);
                      const rowMoisture = calcRowMoistureCapture(idx);
                      const isRateValid = rowRate !== '-' && parseFloat(rowRate) >= 90 && parseFloat(rowRate) <= 110;
                      
                      return (
                        <tr key={`meter-row-${idx}`} className={`border-b border-slate-100 last:border-0 hover:bg-slate-50 transition-colors`}>
                          <td className="p-1 font-bold text-slate-500 w-8">{isStartRow ? '초기' : idx}</td>
                          <td className="p-1 w-12">
                            {isStartRow ? (
                              <span className="font-bold text-slate-400">시작</span>
                            ) : (
                              <input type="text" value={meter.pointNum} onChange={(e) => handleGasMeterChange(idx, 'pointNum', e.target.value)} placeholder="입력" className="w-full p-1 border border-slate-200 rounded text-center outline-none focus:ring-2 focus:ring-emerald-500" />
                            )}
                          </td>
                          <td className="p-1 w-12">
                            {isStartRow ? (
                              <span className="font-bold text-slate-400">0</span>
                            ) : (
                              <input type="number" step="0.1" value={meter.time} onChange={(e) => handleGasMeterChange(idx, 'time', e.target.value)} placeholder="5" className="w-full p-1 border border-slate-200 rounded text-center outline-none focus:ring-2 focus:ring-emerald-500" />
                            )}
                          </td>
                          <td className="p-1 w-14">
                            <input type="number" step="0.1" value={meter.stackTemp} onChange={(e) => handleGasMeterChange(idx, 'stackTemp', e.target.value)} placeholder="Ts" className={`w-full p-1 border border-slate-200 rounded text-center outline-none focus:ring-2 focus:ring-emerald-500 ${isStartRow ? 'bg-slate-50' : ''}`} />
                          </td>
                          
                          <td className="p-1 w-14">
                            <input type="number" step="0.1" value={meter.dp} onChange={(e) => handleGasMeterChange(idx, 'dp', e.target.value)} placeholder="Δp" className="w-full p-1 border border-emerald-300 bg-emerald-50 rounded text-center font-bold outline-none focus:ring-2 focus:ring-emerald-500" />
                          </td>
                          <td className="p-1 w-14">
                            <input type="number" step="0.1" value={meter.sp || ''} onChange={(e) => handleGasMeterChange(idx, 'sp', e.target.value)} placeholder="Ps" className={`w-full p-1 border border-slate-200 rounded text-center outline-none focus:ring-2 focus:ring-emerald-500 ${isStartRow ? 'bg-slate-50' : ''}`} />
                          </td>
                          {!isFlowSheet && (
                          <td className="p-1 w-14">
                            <input type="number" step="0.1" value={meter.pressure} onChange={(e) => handleGasMeterChange(idx, 'pressure', e.target.value)} placeholder="ΔH" className="w-full p-1 border border-teal-300 bg-teal-50 rounded text-center font-black text-teal-800 outline-none focus:ring-2 focus:ring-teal-500" />
                          </td>
                          )}
                          {!isFlowSheet && (
                          <td className="p-1 w-20">
                            <input type="number" step="0.1" value={meter.volume} onChange={(e) => handleGasMeterChange(idx, 'volume', e.target.value)} className={`w-full p-1 border ${isStartRow ? 'border-slate-300 bg-slate-100 text-slate-900 font-bold' : 'border-slate-200 bg-white font-medium'} rounded text-center outline-none focus:ring-2 focus:ring-emerald-500`} placeholder={isStartRow ? '초기유량' : 'Vm'} />
                          </td>
                          )}
                          
                          {!isFlowSheet && (
                          <td className="p-1 w-28">
                            <div className="flex items-center gap-1">
                              <input type="number" step="0.1" value={meter.tmIn} onChange={(e) => handleGasMeterChange(idx, 'tmIn', e.target.value)} placeholder="입구" className="w-1/2 p-1 border border-slate-200 rounded text-center outline-none focus:ring-2 focus:ring-emerald-500" />
                              <input type="number" step="0.1" value={meter.tmOut} onChange={(e) => handleGasMeterChange(idx, 'tmOut', e.target.value)} placeholder="출구" className="w-1/2 p-1 border border-slate-200 rounded text-center outline-none focus:ring-2 focus:ring-emerald-500" />
                            </div>
                          </td>
                          )}
                          
                          {!isFlowSheet && (
                          <td className="p-1 w-14">
                            <input type="number" step="0.1" value={meter.vacuum} onChange={(e) => handleGasMeterChange(idx, 'vacuum', e.target.value)} placeholder="압" className={`w-full p-1 border border-slate-200 rounded text-center outline-none focus:ring-2 focus:ring-emerald-500 ${isStartRow ? 'bg-slate-50' : ''}`} />
                          </td>
                          )}
                          {!isFlowSheet && (
                          <td className="p-1 w-14">
                            <input type="number" step="0.1" value={meter.impingerTemp} onChange={(e) => handleGasMeterChange(idx, 'impingerTemp', e.target.value)} placeholder="온도" className={`w-full p-1 border border-slate-200 rounded text-center outline-none focus:ring-2 focus:ring-emerald-500 ${isStartRow ? 'bg-slate-50' : ''}`} />
                          </td>
                          )}
                          <td className="p-1 w-16">
                            <span className={`inline-block w-full py-1 rounded-lg font-bold ${rowVelocity === '-' ? 'text-slate-400 bg-slate-100' : 'text-slate-900 bg-slate-50 border border-slate-200'}`}>
                              {rowVelocity}
                            </span>
                          </td>
                          <td className="p-1 w-16">
                            <span className={`inline-block w-full py-1 rounded-lg font-bold ${rowDensity === '-' ? 'text-slate-400 bg-slate-100' : 'text-slate-900 bg-slate-50 border border-slate-200'}`}>
                              {rowDensity}
                            </span>
                          </td>
                          {!isFlowSheet && (
                          <td className="p-1 w-20">
                            <span className={`inline-block w-full py-1 rounded-lg font-bold ${rowMoisture === '-' ? 'text-slate-400 bg-slate-100' : 'text-slate-900 bg-slate-50 border border-slate-200'}`}>
                              {rowMoisture}
                            </span>
                          </td>
                          )}
                          
                          {!isFlowSheet && (
                          <td className="p-1 border-l border-slate-200 w-20 ds-iso-cell">
                            {rowRate === '-' ? (
                              <span className="ds-iso-empty">-</span>
                            ) : (
                              <span className={`ds-iso-value ${isRateValid ? 'is-ok' : 'is-out'}`} title={isRateValid ? '등속 범위(90~110%) 안' : '등속 범위(90~110%) 밖'}>{rowRate}%</span>
                            )}
                          </td>
                          )}
                          <td className="p-1 w-8">
                            {!isStartRow && <button type="button" onClick={() => removeGasMeter(idx)} className="text-red-500 hover:text-red-700 text-xs font-bold p-1 transition-colors">삭제</button>}
                          </td>
                        </tr>
                      );
                    })}
                    <tr className="bg-slate-100 border-t-2 border-slate-300 font-bold text-slate-800">
                      <td className="p-1">평균</td>
                      <td className="p-1">-</td>
                      <td className="p-1">{meterTableAverages.time}</td>
                      <td className="p-1">{meterTableAverages.stackTemp}</td>
                      <td className="p-1">{meterTableAverages.dp}</td>
                      <td className="p-1">{meterTableAverages.sp}</td>
                      {!isFlowSheet && <td className="p-1">{meterTableAverages.pressure}</td>}
                      {!isFlowSheet && <td className="p-1">{meterTableAverages.volume}</td>}
                      {!isFlowSheet && <td className="p-1">{meterTableAverages.tmPair}</td>}
                      {!isFlowSheet && <td className="p-1">{meterTableAverages.vacuum}</td>}
                      {!isFlowSheet && <td className="p-1">{meterTableAverages.impingerTemp}</td>}
                      <td className="p-1">{meterTableAverages.velocity}</td>
                      <td className="p-1">{meterTableAverages.density}</td>
                      {!isFlowSheet && <td className="p-1">{meterTableAverages.moisture}</td>}
                      {!isFlowSheet && <td className="p-1 text-emerald-700">{meterTableAverages.rate === '-' ? '-' : `${meterTableAverages.rate}%`}</td>}
                      <td className="p-1"></td>
                    </tr>
                  </tbody>
                </table>
              </div>
              <p className="mt-2 text-[11px] text-slate-500">표 채우기: 칸에서 <b>Ctrl+D</b> 윗칸 값 복사 · <b>Ctrl+Shift+D</b> 지금 값을 아래 끝까지 채우기 · 칸 오른쪽 아래 <b>작은 네모</b>를 끌면 지나간 칸에 같은 값</p>
              <button type="button" onClick={addGasMeter} className="w-full mt-2 py-2 text-emerald-700 bg-emerald-100 hover:bg-emerald-200 rounded-lg text-sm font-bold border border-emerald-300 transition-colors shadow-sm">
                측정 기록 칸 추가
              </button>
            </div>

            {/* 유량계 기록 요약표 */}
            <div className={`grid grid-cols-1 ${isFlowSheet ? 'md:grid-cols-3' : 'md:grid-cols-6'} gap-4 mb-6`}>
              <div className="ds-stat">
                <span className="ds-stat-label">{isFlowSheet ? '총 측정 시간 (분)' : '총 채취 시간 (분)'}</span>
                <span className="ds-stat-value">{getSamplingMinutes()}</span>
              </div>
              {isFlowSheet && (
                <>
                  <div className="ds-stat">
                    <span className="ds-stat-label">평균 유속 (V<sub>s</sub>, m/s)</span>
                    <span className="ds-stat-value">{calcFinalGasVelocity()}</span>
                  </div>
                  <div className="ds-stat">
                    <span className="ds-stat-label">측정 행 수</span>
                    <span className="ds-stat-value">{getMeterRowsWithVolume().filter((_, i) => i > 0 || !isFlowSheet).length}</span>
                  </div>
                </>
              )}
              {!isFlowSheet && (<>
              <div className="ds-stat">
                <span className="ds-stat-label">총 시료채취량 (m³, Yd 적용)</span>
                <span className="ds-stat-value">{calcGasMeterVolDiffM3()}</span>
              </div>
              <div className="ds-stat">
                <span className="ds-stat-label">표준 채취량 (SL)</span>
                <span className="ds-stat-value">{calcVmStdSL()}</span>
              </div>
              <div className="ds-stat">
                <span className="ds-stat-label">표준 채취량 (Sm³)</span>
                <span className="ds-stat-value">{calcVmStdSm3()}</span>
              </div>
              <div className="ds-stat">
                <span className="ds-stat-label">평균 온도 (T<sub>m</sub>, ℃)</span>
                <span className="ds-stat-value">{calcAvgTm()}</span>
              </div>
              <div className="ds-stat">
                <span className="ds-stat-label">평균 오리피스압 (ΔH)</span>
                <span className="ds-stat-value">{calcAvgOrifice()}</span>
              </div>
              </>)}
            </div>

            {!isFlowSheet && (
            <div className="bg-white p-5 rounded-2xl flex flex-col justify-between shadow-lg border-2 border-slate-300">
              <div className="mb-4">
                <div>
                  <h3 className="font-black text-lg text-slate-900 tracking-tight">등속흡인율</h3>
                  <p className="text-xs text-slate-600 mt-1">적합 범위 90~110%</p>
                </div>
              </div>
              
              <div className="grid grid-cols-1 md:grid-cols-2 gap-4 mt-2">
                <div className="bg-slate-100 p-4 rounded-xl border border-slate-300 flex flex-col items-center justify-center">
                  <span className="text-xs text-slate-700 mb-1 font-bold">사전 수분량 기준 (전체)</span>
                  <div className="text-3xl font-black text-slate-900">{(() => { const v = calcIsokineticRate(false); return Number.isFinite(parseFloat(v)) ? <>{v}<span className="text-base ml-1 text-slate-600">%</span></> : '-'; })()}</div>
                </div>
                <div className="bg-slate-900 p-4 rounded-xl border border-slate-800 flex flex-col items-center justify-center shadow-md relative overflow-hidden">
                  <div className="absolute top-0 left-0 w-full h-1 bg-white/50"></div>
                  <span className="text-xs text-slate-200 mb-1 font-bold">사후 수분량 기준 (최종)</span>
                  <div className="text-4xl font-black text-white">{(() => { const v = calcIsokineticRate(true); return Number.isFinite(parseFloat(v)) ? <>{v}<span className="text-xl ml-1 text-slate-200">%</span></> : <span className="text-2xl text-slate-200">사후 수분량 입력 전</span>; })()}</div>
                </div>
              </div>

              <div className="mt-5 flex justify-end">
                {calcIsokineticRate(true) !== '-' && (
                  <div>
                    {(parseFloat(calcIsokineticRate(true)) >= 90 && parseFloat(calcIsokineticRate(true)) <= 110) ? 
                      <span className="flex items-center text-white text-sm font-bold bg-green-700 px-4 py-2 rounded-full shadow-sm">적합</span> : 
                      <span className="flex items-center text-white text-sm font-bold bg-red-600 px-4 py-2 rounded-full shadow-sm">부적합 (재측정 필요)</span>
                    }
                  </div>
                )}
              </div>
            </div>
            )}
          </div>
          )}

          {/* 섹션 6: 시료 분석 및 비고 */}
          <div className="bg-white p-6 rounded-xl shadow-sm border border-slate-200">
            <h2 className="ds-sec-title">
              <span className="ds-sec-num">{isFlowSheet ? 5 : 6}</span>
              {isDustSheet ? '먼지 시료 무게 및 최종 결과' : (isFlowSheet ? '유량 산출 결과 및 비고' : '최종 결과 및 비고')}
            </h2>
            
            {isDustSheet && (
              <div className="grid grid-cols-1 lg:grid-cols-3 gap-6 mb-6">
                <div className="lg:col-span-2 grid grid-cols-3 gap-4 border border-slate-200 p-4 rounded-xl bg-slate-50 h-fit">
                  <div><label className="block text-xs font-bold text-slate-700 mb-1">여과지 번호</label>
                    <input type="text" name="filterId" value={formData.filterId} onChange={handleChange} className="w-full p-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-teal-500 outline-none" /></div>
                  <div><label className="block text-xs font-bold text-slate-700 mb-1">채취 전 무게 (mg)</label>
                    <input type="number" step="0.0001" name="filterInitial" value={formData.filterInitial} onChange={handleChange} placeholder="0.0000" className="w-full p-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-teal-500 outline-none" /></div>
                  <div><label className="block text-xs font-bold text-slate-700 mb-1">채취 후 무게 (mg)</label>
                    <input type="number" step="0.0001" name="filterFinal" value={formData.filterFinal} onChange={handleChange} placeholder="0.0000" className="w-full p-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-teal-500 outline-none" /></div>
                </div>
                <div className="lg:col-span-1 bg-teal-50 p-4 rounded-xl border border-teal-200 flex flex-col justify-center items-center text-teal-900 shadow-inner">
                  <span className="text-xs font-bold mb-1">포집된 먼지 무게 (W<sub>d</sub>)</span>
                  <span className="text-3xl font-black text-teal-700">{calcDustWeightDiff()} <span className="text-sm font-normal">mg</span></span>
                </div>
              </div>
            )}

            {/* 최종 농도 산출 결과 패널 */}
            {(() => {
              const finalRate = calcIsokineticRate(true);
              const finalRateNum = parseFloat(finalRate);
              const rateState = !Number.isFinite(finalRateNum) ? 'none' : (finalRateNum >= 90 && finalRateNum <= 110 ? 'ok' : 'out');
              const items = [
                { label: '표준 습가스 유량', sym: 'Qsw', value: gasFlowRates.wet, unit: 'Sm³/hr' },
                { label: '표준 건조가스 유량', sym: 'Qs', value: gasFlowRates.dry, unit: 'Sm³/hr' },
                { label: '배출가스 유속', sym: 'Vs', value: calcFinalGasVelocity(), unit: 'm/s' },
                ...(isDustSheet ? [
                  { label: '실측 먼지 농도', sym: 'C', value: calcDustConcentrations().actualC, unit: 'mg/Sm³' },
                  { label: formData.standardO2 ? 'O₂ 보정 농도' : 'O₂ 보정 농도 (보정 제외)', sym: 'Cc', value: calcDustConcentrations().correctedC, unit: 'mg/Sm³', strong: true },
                ] : []),
              ];
              return (
                <section className="ds-ledger mb-6" aria-label="최종 산출 결과">
                  <div className="ds-ledger-head">
                    <span className="ds-ledger-title">최종 산출 결과</span>
                    {!isFlowSheet && (
                      <span className={`ds-ledger-status is-${rateState}`}>
                        등속흡인율 <strong>{Number.isFinite(finalRateNum) ? `${finalRate}%` : '-'}</strong>
                        <span className="ds-ledger-status-note">{rateState === 'ok' ? '90~110% 적합' : rateState === 'out' ? '90~110% 밖' : '기록표 입력 전'}</span>
                      </span>
                    )}
                  </div>
                  <dl className="ds-ledger-row">
                    {items.map((item) => (
                      <div key={item.sym} className={`ds-ledger-item${item.strong ? ' is-strong' : ''}`}>
                        <dt>{item.label} <span className="ds-sym">{item.sym}</span></dt>
                        <dd><span className="ds-num">{item.value}</span><span className="ds-unit">{item.unit}</span></dd>
                      </div>
                    ))}
                  </dl>
                </section>
              );
            })()}

            <div className="mt-2">
              <label className="block text-xs font-bold text-slate-700 mb-1">현장 특이사항 및 비고</label>
              <textarea name="remarks" value={formData.remarks} onChange={handleChange} rows="2" className="w-full p-3 border border-slate-300 rounded-xl focus:ring-2 focus:ring-emerald-500 outline-none bg-slate-50/50" placeholder="측정공 상태, 장비 특이점 등"></textarea>
            </div>
          </div>

          <div className="flex items-center justify-center gap-4 mt-8 pb-8 border-b-2 border-slate-200 border-dashed">
            <button type="button" onClick={startNewRecord} className="px-8 py-3.5 bg-white border border-slate-300 text-slate-700 font-bold rounded-xl hover:bg-slate-50 transition-colors shadow-sm">
              새 기록지 작성
            </button>
            <button type="submit" className={`px-8 py-3.5 text-white font-bold rounded-xl transition-colors shadow-md ${activeTheme.primaryButton}`}>
              기록부 데이터 저장
            </button>
          </div>
        </form>

        <div className="bg-white p-6 rounded-2xl shadow-sm mt-8 text-slate-800 border border-slate-200">
          <div className="flex flex-col md:flex-row justify-between items-center mb-6 gap-4 border-b border-slate-200 pb-4">
            <h2 className="text-lg font-bold text-slate-900">
              저장된 종합 리포트 ({displayUserNameWithFallback})
            </h2>
            <div className="flex items-center gap-2">
              <span className="text-[11px] text-slate-500 font-bold">선택 {sheetCheckedReportKeysValid.length}건</span>
              <button
                onClick={handleDeleteCheckedSheetReports}
                className="px-3 py-1.5 border border-red-300 text-red-600 bg-white hover:bg-red-50 font-bold rounded-lg text-xs"
              >
                선택 삭제
              </button>
              <button
                onClick={handleCopyCheckedReportsToCurrentSheet}
                className={`px-3 py-1.5 text-white font-bold rounded-lg text-xs ${activeTheme.primaryButton}`}
              >
                선택 항목을 이 기록부로 복사
              </button>
              <button
                onClick={handleExportSelectedTemplateExcel}
                className={`px-3 py-1.5 text-white font-bold rounded-lg text-xs ${activeTheme.primaryButton}`}
              >
                선택 항목 엑셀로 추출
              </button>
            </div>
          </div>

          <div className={`grid grid-cols-1 ${isFlowSheet ? 'md:grid-cols-1' : 'md:grid-cols-4'} gap-3 mb-4`}>
            <div className="bg-slate-50 border border-slate-200 rounded-lg p-3 text-center">
              <p className="text-[11px] text-slate-500 font-bold">저장 건수</p>
              <p className="text-2xl font-black text-slate-900">{savedData.length}</p>
            </div>
            {!isFlowSheet && (<>
            <div className="bg-slate-50 border border-slate-200 rounded-lg p-3 text-center">
              <p className={`text-[11px] font-bold ${activeTheme.accentText}`}>등속 적합</p>
              <p className={`text-2xl font-black ${activeTheme.accentText}`}>{savedFitCount}</p>
            </div>
            <div className="bg-red-50 border border-red-200 rounded-lg p-3 text-center">
              <p className="text-[11px] text-red-600 font-bold">등속 부적합</p>
              <p className="text-2xl font-black text-red-600">{savedFailCount}</p>
            </div>
            <div className="bg-slate-50 border border-slate-200 rounded-lg p-3 text-center">
              <p className="text-[11px] text-slate-500 font-bold">평균 실측농도</p>
              <p className="text-2xl font-black text-slate-900">{savedAvgActualC}</p>
            </div>
            </>)}
          </div>

          {renderReportGroupTable({
            groups: groupSheetReportsForLog(savedData, resolveReportSheetId, null, { includeBlank: true }),
            checkedSet: sheetCheckedReportKeySet,
            toggleCheck: toggleSheetReportCheck,
            allChecked: isAllSheetReportsChecked,
            toggleAll: toggleSheetAllReportChecks,
            allowCopy: true,
            emptyText: activeUser
              ? (isUserUnlocked ? '현재 사용자의 저장 리포트가 없습니다.' : '로그인하면 저장 리포트가 표시됩니다.')
              : '로그인하거나 새 계정을 만들어주세요.',
          })}
        </div>

      </div>
    </div>
  );
}
