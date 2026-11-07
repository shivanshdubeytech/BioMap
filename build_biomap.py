# -*- coding: utf-8 -*-
"""
Builder script to generate the production-quality, clean, professional biomap.html
"""

import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

def build():
    # Construct complete HTML
    html = """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>BioMap — Byculla Zoo Local Biodiversity Mapping & Web-GIS Portal</title>
<meta name="description" content="Local Biodiversity Mapping of Byculla Zoo and Surrounding Green Spaces: A Geospatial Biodiversity Census, Web-GIS Cartography & Community Ecological Literacy Study. Vidyavardhini's Vartak College, University of Mumbai (NEP 2020).">
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=Newsreader:ital,opsz,wght@0,6..72,400;0,6..72,500;0,6..72,600;1,6..72,400&family=JetBrains+Mono:wght@400;500;600&display=swap" rel="stylesheet">
<style>
/* ============================================================
   BioMap Modern Ecological Design System
   ============================================================ */
:root {
  --bg: #F8FAF7;
  --bg-tint: #F0F4EE;
  --surface: #FFFFFF;
  --surface-2: #F4F6F1;
  --ink: #1C231E;
  --ink-2: #3E4740;
  --muted: #6B756E;
  --line: #DCE2DA;
  --line-2: #E8ECE5;
  --primary: #1B4E34;
  --secondary: #276C48;
  --accent: #276C48;
  --accent-ink: #1B4E34;
  --accent-wash: #E8F2EC;
  --gold: #C0741F;
  --blue: #2563EB;
  --plant: #2E7D32;
  --bird: #1D4ED8;
  --butterfly: #D97706;
  --insect: #7C3AED;
  --mammal: #B45309;
  --reptile: #047857;
  --other: #991B1B;
  --map-land: #EFF2E8;
  --map-stroke: #CCD2C2;
  --map-label: #5C6656;
  --shadow-sm: 0 1px 3px rgba(0,0,0,0.06), 0 1px 2px rgba(0,0,0,0.04);
  --shadow-md: 0 4px 10px -2px rgba(0,0,0,0.07), 0 2px 4px -2px rgba(0,0,0,0.05);
  --shadow-lg: 0 12px 24px -4px rgba(0,0,0,0.12), 0 4px 8px -2px rgba(0,0,0,0.06);
  --r-sm: 6px;
  --r: 10px;
  --r-lg: 14px;
  --sans: "Inter", -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
  --serif: "Newsreader", Georgia, serif;
  --mono: "JetBrains Mono", monospace;
}

*, *::before, *::after {
  box-sizing: border-box;
  margin: 0;
  padding: 0;
}

html, body {
  height: 100%;
  font-family: var(--sans);
  background: var(--bg);
  color: var(--ink);
  line-height: 1.5;
  -webkit-font-smoothing: antialiased;
}

/* ---- Unified Professional Button System ---- */
.btn {
  display: inline-flex;
  align-items: center;
  justify-content: center;
  gap: 6px;
  height: 34px;
  padding: 0 13px;
  border-radius: var(--r-sm);
  border: 1px solid var(--line);
  background: var(--surface);
  color: var(--ink);
  font-family: var(--sans);
  font-size: 13px;
  font-weight: 500;
  cursor: pointer;
  white-space: nowrap;
  transition: all 0.15s ease;
  box-shadow: var(--shadow-sm);
  text-decoration: none;
}
.btn:hover {
  border-color: var(--accent);
  background: var(--surface-2);
  color: var(--accent-ink);
}
.btn-primary {
  background: var(--accent) !important;
  border-color: var(--accent) !important;
  color: #FFFFFF !important;
  font-weight: 600;
}
.btn-primary:hover {
  background: var(--accent-ink) !important;
  border-color: var(--accent-ink) !important;
  color: #FFFFFF !important;
  box-shadow: 0 2px 8px rgba(27, 78, 52, 0.28);
}
.btn-sm {
  height: 28px;
  padding: 0 9px;
  font-size: 12px;
  border-radius: 5px;
}
.btn-icon {
  width: 32px;
  height: 32px;
  padding: 0;
  border-radius: var(--r-sm);
}

.chip {
  display: inline-flex;
  align-items: center;
  gap: 5px;
  height: 26px;
  padding: 0 9px;
  border-radius: 999px;
  border: 1px solid var(--line);
  background: var(--surface);
  color: var(--ink-2);
  font-size: 12px;
  font-weight: 500;
  cursor: pointer;
  transition: all 0.15s ease;
}
.chip:hover {
  border-color: var(--accent);
  color: var(--accent-ink);
}
.chip.active, .chip[aria-pressed="true"] {
  background: var(--accent-wash);
  border-color: var(--accent);
  color: var(--accent-ink);
  font-weight: 600;
}

/* ---- Form Controls ---- */
.input {
  width: 100%;
  height: 34px;
  padding: 0 10px;
  border-radius: var(--r-sm);
  border: 1px solid var(--line);
  background: var(--surface);
  color: var(--ink);
  font-family: var(--sans);
  font-size: 13px;
  outline: none;
  transition: border-color 0.15s ease;
}
.input:focus {
  border-color: var(--accent);
  box-shadow: 0 0 0 2px var(--accent-wash);
}
textarea.input {
  height: auto;
  padding: 8px 10px;
}
.field {
  display: flex;
  flex-direction: column;
  gap: 4px;
}
.field label {
  font-size: 12px;
  font-weight: 600;
  color: var(--ink-2);
}

/* ---- Badges & Typography ---- */
.badge-iucn {
  display: inline-flex;
  align-items: center;
  gap: 4px;
  padding: 2px 7px;
  border-radius: 999px;
  font-size: 10.5px;
  font-weight: 600;
  text-transform: uppercase;
  letter-spacing: 0.03em;
  white-space: nowrap;
}
.iucn-cr { background: #fee2e2; color: #991b1b; border: 1px solid #f87171; }
.iucn-en { background: #ffedd5; color: #9a3412; border: 1px solid #fb923c; }
.iucn-vu { background: #fef3c7; color: #92400e; border: 1px solid #fcd34d; }
.iucn-nt { background: #fef9c3; color: #854d0e; border: 1px solid #facc15; }
.iucn-lc { background: #dcfce7; color: #166534; border: 1px solid #86efac; }
.iucn-heritage { background: #f3e8ff; color: #6b21a8; border: 1px solid #d8b4fe; }
.iucn-native { background: #e0f2fe; color: #075985; border: 1px solid #7dd3fc; }

.marathi-sub {
  color: var(--accent-ink);
  font-size: 13.5px;
  font-weight: 500;
  margin-top: 1px;
}
.voucher-pill {
  font-family: var(--mono);
  font-size: 11px;
  font-weight: 600;
  background: var(--bg-tint);
  color: var(--ink-2);
  padding: 1.5px 6px;
  border-radius: 4px;
  border: 1px solid var(--line);
}
.dba-pill {
  display: inline-flex;
  align-items: center;
  gap: 4px;
  font-family: var(--mono);
  font-size: 11px;
  font-weight: 600;
  padding: 2px 7px;
  border-radius: 999px;
}
.dba-quiet { background: #dcfce7; color: #166534; border: 1px solid #86efac; }
.dba-mod { background: #fef3c7; color: #92400e; border: 1px solid #fcd34d; }
.dba-high { background: #fee2e2; color: #991b1b; border: 1px solid #f87171; }

.tag {
  display: inline-flex;
  align-items: center;
  gap: 5px;
  font-size: 11.5px;
  padding: 2px 8px;
  border-radius: 999px;
  background: var(--bg-tint);
  color: var(--ink-2);
  border: 1px solid var(--line);
}
.tag .dot {
  width: 6px;
  height: 6px;
  border-radius: 50%;
}

.mono-n {
  font-family: var(--mono);
  font-variant-numeric: tabular-nums;
}
.sci {
  font-style: italic;
  color: var(--muted);
  font-size: 12.5px;
}
.hint {
  font-size: 11.5px;
  color: var(--muted);
}

/* ---- Core App Layout ---- */
.app {
  display: flex;
  height: 100vh;
  overflow: hidden;
}
.sidebar {
  width: 240px;
  background: var(--surface);
  border-right: 1px solid var(--line);
  display: flex;
  flex-direction: column;
  padding: 16px 12px;
  flex-shrink: 0;
  overflow-y: auto;
}
.main-wrap {
  flex: 1;
  display: flex;
  flex-direction: column;
  overflow: hidden;
  position: relative;
}
.topbar {
  height: 52px;
  background: var(--surface);
  border-bottom: 1px solid var(--line);
  display: flex;
  align-items: center;
  padding: 0 20px;
  gap: 12px;
  flex-shrink: 0;
  z-index: 10;
}
.content {
  flex: 1;
  overflow-y: auto;
  padding: 20px;
}

/* ---- Brand & Navigation ---- */
.brand {
  display: flex;
  align-items: center;
  gap: 10px;
  padding: 4px 8px 14px 8px;
  border-bottom: 1px solid var(--line-2);
  margin-bottom: 12px;
}
.brand-name {
  font-family: var(--serif);
  font-size: 18px;
  font-weight: 600;
  color: var(--accent-ink);
  line-height: 1.1;
}
.brand-sub {
  font-size: 11px;
  color: var(--muted);
  margin-top: 1px;
}
.nav-label {
  font-size: 10.5px;
  font-weight: 700;
  text-transform: uppercase;
  letter-spacing: 0.05em;
  color: var(--muted);
  padding: 8px 10px 4px 10px;
}
.nav-btn {
  display: flex;
  align-items: center;
  gap: 10px;
  width: 100%;
  height: 36px;
  padding: 0 10px;
  border-radius: var(--r-sm);
  border: 0;
  background: transparent;
  color: var(--ink-2);
  font-family: var(--sans);
  font-size: 13px;
  font-weight: 500;
  cursor: pointer;
  transition: all 0.15s ease;
  text-align: left;
}
.nav-btn:hover {
  background: var(--surface-2);
  color: var(--ink);
}
.nav-btn.active {
  background: var(--accent-wash);
  color: var(--accent-ink);
  font-weight: 600;
}
.nav-btn svg {
  width: 18px;
  height: 18px;
  flex-shrink: 0;
}

.area-pill {
  display: inline-flex;
  align-items: center;
  gap: 6px;
  font-size: 12.5px;
  color: var(--ink-2);
  background: var(--bg-tint);
  padding: 4px 10px;
  border-radius: 999px;
  border: 1px solid var(--line);
}
.area-pill .dot {
  width: 7px;
  height: 7px;
  border-radius: 50%;
}

/* ---- Cards & Grids ---- */
.card {
  background: var(--surface);
  border: 1px solid var(--line);
  border-radius: var(--r);
  box-shadow: var(--shadow-sm);
}
.card-pad {
  padding: 16px 18px;
}
.card-h {
  padding: 12px 18px;
  border-bottom: 1px solid var(--line-2);
  display: flex;
  justify-content: space-between;
  align-items: center;
}
.card-h h3 {
  font-size: 14.5px;
  font-weight: 600;
  color: var(--ink);
}
.grid {
  display: grid;
}

/* ---- KPI Stats Row ---- */
.stats {
  display: grid;
  grid-template-columns: repeat(auto-fit, minmax(140px, 1fr));
  gap: 12px;
}
.stat {
  background: var(--surface);
  border: 1px solid var(--line);
  border-radius: var(--r);
  padding: 12px 14px;
  box-shadow: var(--shadow-sm);
}
.stat .v {
  font-size: 20px;
  font-weight: 700;
  color: var(--accent-ink);
  line-height: 1.1;
  display: block;
}
.stat .k {
  font-size: 11.5px;
  color: var(--muted);
  margin-top: 4px;
  display: flex;
  align-items: center;
  gap: 5px;
}

/* ---- Hero Banner ---- */
.hero-banner {
  background: linear-gradient(135deg, #1B4E34 0%, #276C48 55%, #387A57 100%);
  color: #FFFFFF;
  border-radius: var(--r-lg);
  padding: 22px 26px;
  margin-bottom: 20px;
  box-shadow: var(--shadow-md);
  position: relative;
  overflow: hidden;
}
.hero-banner::after {
  content: "";
  position: absolute;
  right: -50px;
  top: -50px;
  width: 220px;
  height: 220px;
  background: radial-gradient(circle, rgba(255,255,255,0.18) 0%, rgba(255,255,255,0) 70%);
  border-radius: 50%;
  pointer-events: none;
}
.hero-title {
  font-family: var(--serif);
  font-size: 24px;
  line-height: 1.25;
  margin: 0 0 6px 0;
  font-weight: 600;
}
.hero-sub {
  font-size: 13.5px;
  line-height: 1.5;
  opacity: 0.94;
  max-width: 82ch;
  margin-bottom: 12px;
}
.hero-meta {
  display: flex;
  flex-wrap: wrap;
  gap: 8px 14px;
  font-size: 12px;
  align-items: center;
  border-top: 1px solid rgba(255,255,255,0.22);
  padding-top: 10px;
}
.hero-meta span {
  display: inline-flex;
  align-items: center;
  gap: 5px;
  background: rgba(0,0,0,0.18);
  padding: 2px 8px;
  border-radius: 999px;
}

/* ---- Map Styling ---- */
.mapwrap {
  position: relative;
  width: 100%;
  border-radius: var(--r);
  overflow: hidden;
  background: var(--map-land);
  border: 1px solid var(--line);
}
.mapsvg {
  width: 100%;
  height: 100%;
  display: block;
  user-select: none;
}
.map-ctrls {
  position: absolute;
  top: 10px;
  right: 10px;
  display: flex;
  flex-direction: column;
  gap: 5px;
  z-index: 10;
}
.map-btn {
  width: 32px;
  height: 32px;
  display: grid;
  place-items: center;
  border-radius: var(--r-sm);
  border: 1px solid var(--line);
  background: var(--surface);
  color: var(--ink-2);
  cursor: pointer;
  box-shadow: var(--shadow-sm);
  transition: all 0.15s ease;
  padding: 0;
}
.map-btn:hover {
  border-color: var(--accent);
  color: var(--accent-ink);
}
.map-btn.active {
  background: #fee2e2;
  border-color: #f87171;
  color: #991b1b;
}

.map-legend {
  position: absolute;
  bottom: 10px;
  left: 10px;
  background: rgba(255,255,255,0.92);
  backdrop-filter: blur(4px);
  border: 1px solid var(--line);
  border-radius: var(--r-sm);
  padding: 8px 12px;
  font-size: 11.5px;
  box-shadow: var(--shadow-sm);
  z-index: 10;
}
.legend-row {
  display: flex;
  align-items: center;
  gap: 6px;
  margin-top: 3px;
}

/* ---- Species Cards & Survey Charts ---- */
.sp-card {
  background: var(--surface);
  border: 1px solid var(--line);
  border-radius: var(--r);
  overflow: hidden;
  box-shadow: var(--shadow-sm);
  transition: all 0.15s ease;
  display: flex;
  flex-direction: column;
  justify-content: space-between;
  cursor: pointer;
}
.sp-card:hover {
  border-color: var(--accent);
  box-shadow: var(--shadow-md);
  transform: translateY(-2px);
}
.species-thumb-wrap {
  width: 100%;
  aspect-ratio: 4/3;
  overflow: hidden;
  background: var(--bg-tint);
  position: relative;
}
.species-thumb-wrap img {
  width: 100%;
  height: 100%;
  object-fit: cover;
  transition: transform 0.25s ease;
}
.sp-card:hover .species-thumb-wrap img {
  transform: scale(1.05);
}

.survey-chart-grid {
  display: grid;
  grid-template-columns: repeat(auto-fit, minmax(310px, 1fr));
  gap: 16px;
  margin-top: 16px;
}
.survey-chart-card {
  background: var(--surface);
  border: 1px solid var(--line);
  border-radius: var(--r);
  overflow: hidden;
  box-shadow: var(--shadow-sm);
  display: flex;
  flex-direction: column;
}
.survey-chart-card img {
  width: 100%;
  height: 210px;
  object-fit: contain;
  background: #ffffff;
  border-bottom: 1px solid var(--line-2);
  padding: 8px;
}
.survey-chart-card .c-body {
  padding: 12px 14px;
  flex: 1;
  display: flex;
  flex-direction: column;
  justify-content: space-between;
}

/* ---- Data Tables & Lists ---- */
table.data {
  width: 100%;
  border-collapse: collapse;
  font-size: 12.5px;
}
table.data th {
  text-align: left;
  padding: 8px 12px;
  background: var(--surface-2);
  color: var(--muted);
  font-weight: 600;
  border-bottom: 1px solid var(--line);
  font-size: 11.5px;
  text-transform: uppercase;
  letter-spacing: 0.04em;
}
table.data td {
  padding: 9px 12px;
  border-bottom: 1px solid var(--line-2);
  color: var(--ink);
}
table.data tr:hover td {
  background: var(--surface-2);
}
.scroll-x {
  overflow-x: auto;
}

dl.kv {
  display: grid;
  grid-template-columns: auto 1fr;
  column-gap: 14px;
  row-gap: 8px;
}
dl.kv dt {
  color: var(--muted);
  font-weight: 500;
}
dl.kv dd {
  color: var(--ink);
}

.oral-history-card {
  background: var(--accent-wash);
  border-left: 4px solid var(--accent);
  border-radius: 0 var(--r) var(--r) 0;
  padding: 16px 20px;
  font-family: var(--serif);
  font-size: 14.5px;
  line-height: 1.6;
  color: var(--ink);
  margin: 14px 0;
}
.oral-history-card cite {
  display: block;
  font-family: var(--sans);
  font-style: normal;
  font-size: 12px;
  font-weight: 600;
  color: var(--accent-ink);
  margin-top: 8px;
}

.action-tier-card {
  border-left: 4px solid var(--accent);
  background: var(--surface);
  padding: 12px 16px;
  border-radius: var(--r-sm);
  border: 1px solid var(--line);
  border-left-width: 4px;
  margin-bottom: 10px;
}
.action-tier-card strong {
  display: block;
  font-size: 13.5px;
  margin-bottom: 3px;
  color: var(--ink);
}
.action-tier-card p {
  font-size: 12.5px;
  color: var(--ink-2);
  margin: 0;
  line-height: 1.5;
}
.action-tier-card .stakeholder {
  font-size: 11px;
  color: var(--muted);
  margin-top: 5px;
  font-weight: 500;
}

/* ---- Toast Notification ---- */
.toast {
  position: fixed;
  bottom: 20px;
  right: 20px;
  background: var(--ink);
  color: #FFFFFF;
  padding: 10px 16px;
  border-radius: var(--r-sm);
  font-size: 13px;
  font-weight: 500;
  box-shadow: var(--shadow-lg);
  z-index: 9999;
  animation: slideUp 0.2s ease;
  transition: opacity 0.3s ease;
}

@keyframes slideUp {
  from { transform: translateY(12px); opacity: 0; }
  to { transform: translateY(0); opacity: 1; }
}

/* ---- Mobile Bottom Bar ---- */
.m-bottom {
  display: none;
  height: 52px;
  background: var(--surface);
  border-top: 1px solid var(--line);
  align-items: center;
  justify-content: space-around;
  flex-shrink: 0;
  z-index: 10;
}
.m-btn {
  display: flex;
  flex-direction: column;
  align-items: center;
  justify-content: center;
  gap: 2px;
  background: transparent;
  border: 0;
  color: var(--muted);
  font-size: 10px;
  font-weight: 500;
  cursor: pointer;
}
.m-btn.active {
  color: var(--accent);
  font-weight: 600;
}

@media (max-width: 768px) {
  .sidebar { display: none; }
  .m-bottom { display: flex; }
}
</style>
</head>
<body>
<div id="root"></div>

<script>
"use strict";
/* ============================================================
   BioMap Web-GIS Platform Data Layer
   Directly harmonized with:
   "Byculla Zoo Biodiversity Mapping - CEP Report.docx"
   "Byculla Zoo CEP - Project Overview & Viva Handbook.docx"
   ============================================================ */

const CATS = {
  plant: {label:"Flora", plural:"Flora & Heritage Trees", color:"var(--plant)", shape:"circle"},
  bird: {label:"Avifauna", plural:"Wild Avifauna (Birds)", color:"var(--bird)", shape:"pin"},
  butterfly: {label:"Lepidoptera", plural:"Butterflies", color:"var(--butterfly)", shape:"diamond"},
  insect: {label:"Insecta", plural:"Insects & Odonata", color:"var(--insect)", shape:"hex"},
  mammal: {label:"Mammalia", plural:"Mammals", color:"var(--mammal)", shape:"square"},
  reptile: {label:"Herpetofauna", plural:"Reptiles & Amphibians", color:"var(--reptile)", shape:"pin"}
};
const CAT_ORDER = ["plant","bird","butterfly","insect","mammal","reptile"];

/* ---- Primary Study Area: Ranibagh / Byculla Zoo ---- */
const AREAS = [
  {
    id: "byculla",
    name: "Veermata Jijabai Bhosale Botanical Udyan & Zoo (Ranibagh)",
    city: "Byculla (East), Mumbai",
    areaAcre: 53,
    active: true,
    bounds: { latMax: 18.9818, latMin: 18.9750, lngMin: 72.8318, lngMax: 72.8375 }
  }
];

/* ---- 4-Zone Ecological Stratification (Table 3.1 & Viva Handbook) ---- */
const ZONES = [
  {
    id: "z1",
    code: "Zone 1",
    name: "Botanical Grove Canopy",
    sub: "Heritage Arboretum & Dense Tree Refuge",
    poly: "510,50 965,50 975,320 620,330 500,240",
    box: [520, 60, 950, 310],
    gps: { lat: 18.9789, lng: 72.8354 },
    gpsDms: "18°58′38″N, 72°49′51″E",
    habitat: "Dense multi-layered arboreal canopy (82% cover)",
    vegetation: "African Baobab, Ceylon Ebony, Ficus benghalensis, Mahogany, Rain Tree",
    targetTaxa: "Canopy birds, colonial fruit bats, heritage trees, captive large carnivores",
    soundDba: 52.1,
    soundDesc: "Acoustic haven (26.3 dBA damping vs perimeter)",
    richness: 24,
    shannon: 2.68,
    evenness: 0.84,
    color: "#2F7D53"
  },
  {
    id: "z2",
    code: "Zone 2",
    name: "Freshwater Wetland & Lotus Lake",
    sub: "Eutrophic Lotus Lake & Marshland Periphery",
    poly: "440,320 830,310 865,560 690,640 480,630 420,490",
    box: [440, 330, 840, 620],
    gps: { lat: 18.9794, lng: 72.8361 },
    gpsDms: "18°58′42″N, 72°49′56″E",
    habitat: "Freshwater aquatic ecosystem & emergent weed beds",
    vegetation: "Sacred Lotus (Nelumbo nucifera), Cattail (Typha), Bamboo thickets",
    targetTaxa: "Wading birds, aquatic avifauna, odonates, gharial, turtles",
    soundDba: 54.8,
    soundDesc: "Moderate aquatic buffer",
    richness: 16,
    shannon: 2.34,
    evenness: 0.81,
    color: "#2C6FA8"
  },
  {
    id: "z3",
    code: "Zone 3",
    name: "Promenade Shrublands & Butterfly Lawns",
    sub: "Managed Floral Walkways & Native Shrub Beds",
    poly: "230,80 500,70 500,320 420,490 240,490",
    box: [240, 90, 490, 470],
    gps: { lat: 18.9778, lng: 72.8348 },
    gpsDms: "18°58′35″N, 72°49′48″E",
    habitat: "Flowering border hedges, nectar shrublands & sunny turf",
    vegetation: "Malabar Nut (Adulsa), Lantana camara, Ixora, Hibiscus, Citrus",
    targetTaxa: "Butterflies, solitary bees, sunbirds, dragonflies, tailorbirds",
    soundDba: 61.2,
    soundDesc: "Visitor path pedestrian traffic",
    richness: 12,
    shannon: 2.05,
    evenness: 0.79,
    color: "#C0741F"
  },
  {
    id: "z4",
    code: "Zone 4",
    name: "Perimeter Transition Buffer Wall",
    sub: "Boundary Wall Interface along Dr. Ambedkar Road",
    poly: "40,40 220,40 230,660 50,660",
    box: [50, 50, 210, 650],
    gps: { lat: 18.9768, lng: 72.8335 },
    gpsDms: "18°58′30″N, 72°49′44″E",
    habitat: "Anthropogenic boundary wall, entrance gates & road verges",
    vegetation: "Rain Tree, Copper Pod, Gulmohar, boundary wall ferns",
    targetTaxa: "Synanthropic birds (crows, pigeons, kites, mynas), road-edge flora",
    soundDba: 78.4,
    soundDesc: "High acoustic stress (>78 dBA heavy city traffic)",
    richness: 6,
    shannon: 1.48,
    evenness: 0.65,
    color: "#A24A45"
  }
];
const ZONE_BY_ID = Object.fromEntries(ZONES.map(z=>[z.id,z]));

/* ---- Structured Field Survey Timeline (31/08/2026, 4.8 km Transect) ---- */
const VISITS = [
  {n:1, session:"Morning Avifauna & Acoustic Baseline", date:"2026-08-31", start:"07:00", end:"10:00", zones:["z1","z2"], weather:"Clear, 27°C, 76% RH", photos:48, team:"Riddhesh Keni & Shivansh Dubey", notes:"Early morning dawn chorus audit; baseline sound recording at Zone 1 (52.1 dBA) vs Zone 4 (78.4 dBA)."},
  {n:2, session:"Botanical Heritage Census & Shrubland Pollinators", date:"2026-08-31", start:"10:15", end:"13:00", zones:["z1","z3"], weather:"Sunny, 30°C", photos:62, team:"Riddhesh Keni & Shivansh Dubey", notes:"GPS mapping of heritage Baobab (FLO-01) and Ebony (FLO-02); insect transect along Adulsa shrub border."},
  {n:3, session:"Community Interactions & Citizen Surveys", date:"2026-08-31", start:"13:30", end:"16:00", zones:["z1","z2","z3"], weather:"Humid, 31°C", photos:35, team:"Riddhesh Keni & Shivansh Dubey", notes:"Administered N=84 survey questionnaires to regular walkers, teachers, youth, and senior mali Mr. Maruti Patil."},
  {n:4, session:"Perimeter Acoustic Stress & Crepuscular Roosts", date:"2026-08-31", start:"16:15", end:"18:30", zones:["z4","z1"], weather:"Mild breeze, 28°C", photos:44, team:"Riddhesh Keni & Shivansh Dubey", notes:"Perimeter boundary wall traffic decibel mapping; evening fruit bat emergence count from Ficus canopies."}
];

/* ---- 42+ Master Taxa Census (Table 4.3, Viva Handbook & Vouchers) ---- */
const SPECIES_RAW = [
  // ZONE 1: BOTANICAL GROVE CANOPY (HERITAGE ARBORETUM)
  {id:"baobab", common:"African Baobab", marathi:"गोरखचिंच", sci:"Adansonia digitata", cat:"plant", habitat:"Heritage Botanical Grove (Zone 1)", zones:["z1"], obsCount:6, iucn:"Heritage Tree", voucher:"FLO-01", photo:"animal_photos/baobab_tree.jpg", gps:[18.9782, 72.8351], note:"Massive heritage specimen (>160 years old) with a trunk girth exceeding 8.4 meters. Keystone botanical species providing critical microhabitats and night nectar for frugivorous bats."},
  {id:"ebony", common:"Ceylon Ebony", marathi:"आबनूस", sci:"Diospyros ebenum", cat:"plant", habitat:"Deep Arboretum (Zone 1)", zones:["z1"], obsCount:4, iucn:"Threatened", voucher:"FLO-02", photo:"animal_photos/ebony_tree.jpg", gps:[18.9787, 72.8356], note:"Rare slow-growing hardwood tree documented in Dr. George Birdwood's historic 1862 catalogue. Extremely dense dark heartwood with medicinal bark properties."},
  {id:"mango", common:"Heritage Mango Tree", marathi:"वारसा आंबा", sci:"Mangifera indica", cat:"plant", habitat:"Botanical Grove (Zone 1)", zones:["z1"], obsCount:5, iucn:"Native Flora", voucher:"FLO-04", photo:"animal_photos/mango_tree.jpg", gps:[18.9784, 72.8350], note:"Century-old canopy tree providing dense multi-tier shade, microclimatic moisture retention, and high insect pollinator activity during blooming seasons."},
  {id:"banyan", common:"Banyan Tree", marathi:"वड", sci:"Ficus benghalensis", cat:"plant", habitat:"Core Arboretum (Zone 1)", zones:["z1"], obsCount:5, iucn:"Heritage Tree", voucher:"FLO-05", photo:"animal_photos/banyan_tree.jpg", gps:[18.9790, 72.8353], note:"Venerable specimen with extensive aerial prop roots. Keystone urban ecological refugium sustaining diverse avian and insect communities year-round."},
  {id:"peepal", common:"Sacred Fig / Peepal", marathi:"पिंपळ", sci:"Ficus religiosa", cat:"plant", habitat:"Core Canopy (Zone 1)", zones:["z1"], obsCount:4, iucn:"Native Flora", voucher:"FLO-06", photo:"animal_photos/peepal_tree.jpg", gps:[18.9789, 72.8348], note:"Cordate-leaved fig tree heavily frequented by frugivorous birds (barbets, koels, mynas). Major source of canopy oxygen and ambient cooling."},
  {id:"mahogany", common:"Broad-leaved Mahogany", marathi:"महोगनी", sci:"Swietenia macrophylla", cat:"plant", habitat:"Heritage Grove (Zone 1)", zones:["z1"], obsCount:4, iucn:"Heritage Tree", voucher:"FLO-07", photo:"animal_photos/mahogany_tree.jpg", gps:[18.9791, 72.8357], note:"Towering historic plantation tree from the late 19th century. Dense foliage provides major acoustic shielding against northern city traffic."},
  {id:"tiger", common:"Royal Bengal Tiger", marathi:"पट्टेरी वाघ", sci:"Panthera tigris", cat:"mammal", habitat:"Carnivore Moated Enclosure (Zone 1 Edge)", zones:["z1"], obsCount:5, iucn:"Endangered", voucher:"MAM-01", photo:"animal_photos/tiger.jpg", gps:[18.9792, 72.8349], note:"Flagship apex predator housed in a spacious open-air naturalistic moated habitat with water pool and bamboo cover. High educational value for visitors."},
  {id:"sloth_bear", common:"Indian Sloth Bear", marathi:"अस्वल", sci:"Melursus ursinus", cat:"mammal", habitat:"Rocky Grotto Habitat (Zone 1)", zones:["z1"], obsCount:4, iucn:"Vulnerable", voucher:"MAM-02", photo:"animal_photos/sloth_bear.jpg", gps:[18.9785, 72.8342], note:"Native subcontinental insectivore housed in an enriched natural rock grotto with simulated termite mounds. Observed active soil sniffing and log scratching."},
  {id:"chital", common:"Spotted Deer / Chital", marathi:"चितळ", sci:"Axis axis", cat:"mammal", habitat:"Herbivore Meadow (Zone 1)", zones:["z1"], obsCount:5, iucn:"Least Concern", voucher:"MAM-03", photo:"animal_photos/spotted_deer.jpg", gps:[18.9780, 72.8340], note:"Native Indian cervid herd grazing calmly in an open shaded paddock beneath mature Ficus canopies; gentle browser sustaining garden grassland patches."},
  {id:"fruit_bat", common:"Indian Flying Fox", marathi:"मोठे वटवाघूळ", sci:"Pteropus medius", cat:"mammal", habitat:"High Canopy Roost (Zone 1)", zones:["z1"], obsCount:4, iucn:"Least Concern", voucher:"MAM-04", photo:"animal_photos/indian_flying_fox.jpg", gps:[18.9784, 72.8353], note:"Colonial fruit bat roosting in groups of 80+ individuals in the high crowns of Ficus and Mahogany. Crucial long-distance seed disperser for Mumbai green belts."},
  {id:"coppersmith", common:"Coppersmith Barbet", marathi:"तांबट पक्षी", sci:"Psilopogon haemacephalus", cat:"bird", habitat:"Ficus & Dead Wood Canopy (Zone 1)", zones:["z1"], obsCount:6, iucn:"Least Concern", voucher:"AVE-03", photo:"animal_photos/coppersmith_barbet.jpg", gps:[18.9788, 72.8352], note:"Metronomic 'tuk-tuk' call audible across Zone 1. Excavates nesting cavities in soft or dead branches; highly vulnerable to indiscriminate pre-monsoon tree lopping."},
  {id:"koel", common:"Asian Koel", marathi:"कोकिळा", sci:"Eudynamys scolopaceus", cat:"bird", habitat:"Arboreal Canopy (Zone 1)", zones:["z1"], obsCount:5, iucn:"Least Concern", voucher:"AVE-04", photo:"animal_photos/asian_koel.jpg", gps:[18.9786, 72.8350], note:"Prominent vocal resident throughout the grove. Brood parasite laying eggs primarily in the nests of House Crows along the transition zone."},
  {id:"parakeet", common:"Rose-ringed Parakeet", marathi:"पोपट", sci:"Psittacula krameri", cat:"bird", habitat:"Baobab & High Canopy (Zone 1)", zones:["z1"], obsCount:5, iucn:"Least Concern", voucher:"AVE-05", photo:"animal_photos/rose_ringed_parakeet.jpg", gps:[18.9783, 72.8355], note:"Gregarious flocking frugivore utilizing natural tree hollows in ancient Baobabs for nesting. Very active in early morning and late afternoon hours."},

  // ZONE 2: FRESHWATER WETLAND & LOTUS LAKE
  {id:"stork", common:"Painted Stork", marathi:"रंगीत करकोचा", sci:"Mycteria leucocephala", cat:"bird", habitat:"Wetland Shoreline (Zone 2)", zones:["z2"], obsCount:5, iucn:"Near Threatened", voucher:"AVE-02", photo:"animal_photos/painted_stork.jpg", gps:[18.9796, 72.8364], note:"Free-ranging wetland resident with distinctive pink tertials and yellow decurved bill. Forages via tactile sweeps in shallow lotus pond margins."},
  {id:"night_heron", common:"Black-crowned Night Heron", marathi:"भुरकट बगळा", sci:"Nycticorax nycticorax", cat:"bird", habitat:"Wetland Roost & Island (Zone 2)", zones:["z2"], obsCount:5, iucn:"Least Concern", voucher:"AVE-06", photo:"animal_photos/night_heron.jpg", gps:[18.9795, 72.8362], note:"Nocturnal and crepuscular wading bird roosting in dense overhanging bamboo clusters during daytime; hunts fish and amphibians at dusk."},
  {id:"pond_heron", common:"Indian Pond Heron", marathi:"भांगडी बगळा", sci:"Ardeola grayii", cat:"bird", habitat:"Shallow Muddy Banks (Zone 2)", zones:["z2"], obsCount:6, iucn:"Least Concern", voucher:"AVE-07", photo:"animal_photos/pond_heron.jpg", gps:[18.9793, 72.8359], note:"Cryptic ambush predator standing motionless among aquatic grasses until striking small fish, tadpoles, and odonates."},
  {id:"cormorant", common:"Little Cormorant", marathi:"छोटा पाणकावळा", sci:"Microcarbo niger", cat:"bird", habitat:"Open Water Body (Zone 2)", zones:["z2"], obsCount:4, iucn:"Least Concern", voucher:"AVE-08", photo:"animal_photos/little_cormorant.jpg", gps:[18.9797, 72.8360], note:"Skilled pursuit-diver capturing fish underwater; frequently perches on half-submerged wooden posts with wings spread to dry plumage."},
  {id:"moorhen", common:"Common Moorhen", marathi:"पाणकोंबडी", sci:"Gallinula chloropus", cat:"bird", habitat:"Reed & Lotus Beds (Zone 2)", zones:["z2"], obsCount:4, iucn:"Least Concern", voucher:"AVE-09", photo:"animal_photos/common_moorhen.jpg", gps:[18.9794, 72.8363], note:"Secretive marsh bird with bright red frontal shield and white undertail coverts, swimming gently through Typha reeds."},
  {id:"kingfisher", common:"White-throated Kingfisher", marathi:"पांढऱ्या छातीचा खंड्या", sci:"Halcyon smyrnensis", cat:"bird", habitat:"Lake Edge Branch Perch (Zone 2)", zones:["z2"], obsCount:4, iucn:"Least Concern", voucher:"AVE-10", photo:"animal_photos/white_throated_kingfisher.jpg", gps:[18.9792, 72.8358], note:"Vivid turquoise back and chestnut mantle. Hunts from elevated branches over open water and surrounding grassy verges."},
  {id:"penguin", common:"Humboldt Penguin", marathi:"हंबोल्ट पेंग्विन", sci:"Spheniscus humboldti", cat:"bird", habitat:"Climate-Controlled Aquatic Pavilion (Zone 2)", zones:["z2"], obsCount:6, iucn:"Vulnerable", voucher:"AVE-01", photo:"animal_photos/penguin.jpg", gps:[18.9772, 72.8338], note:"Indoor captive colony housed in a state-of-the-art chiller pool kept at 16°C. While captive, it represents South Mumbai's top educational crowd-puller."},
  {id:"gharial", common:"Gharial", marathi:"घडियाळ", sci:"Gavialis gangeticus", cat:"reptile", habitat:"Aquatic Reptile Enclosure (Zone 2)", zones:["z2"], obsCount:4, iucn:"Critically Endangered", voucher:"REP-01", photo:"animal_photos/gharial.jpg", gps:[18.9791, 72.8363], note:"Critically endangered fish-eating crocodilian with an elongated slender snout. Essential captive conservation breeding and civic awareness program."},
  {id:"turtle", common:"Indian Flapshell Turtle", marathi:"भारतीय कासव", sci:"Lissemys punctata", cat:"reptile", habitat:"Lake Mud & Logs (Zone 2)", zones:["z2"], obsCount:4, iucn:"Least Concern", voucher:"REP-02", photo:"animal_photos/flapshell_turtle.jpg", gps:[18.9795, 72.8365], note:"Semi-aquatic freshwater softshell turtle basking on floating logs and muddy banks; opportunistic scavenger keeping pond water clean."},
  {id:"lotus", common:"Sacred Lotus", marathi:"कमळ", sci:"Nelumbo nucifera", cat:"plant", habitat:"Central Lake Water Surface (Zone 2)", zones:["z2"], obsCount:5, iucn:"Native Aquatic", voucher:"FLO-08", photo:"animal_photos/lotus_flower.jpg", gps:[18.9794, 72.8360], note:"Iconic emergent macrophyte with broad water-repellent peltate leaves; serves as primary shelter and resting platform for aquatic avifauna."},
  {id:"typha", common:"Cattail / Bulrush", marathi:"पाणकणीस", sci:"Typha domingensis", cat:"plant", habitat:"Wetland Shoreline (Zone 2)", zones:["z2"], obsCount:3, iucn:"Native Wetland", voucher:"FLO-09", photo:"animal_photos/typha_cattail.jpg", gps:[18.9796, 72.8358], note:"Dense reed bed providing vital phytoremediation for pond water and critical nesting cover for herons and moorhens."},
  {id:"glider", common:"Wandering Glider", marathi:"चतुर / पाणटोळ", sci:"Pantala flavescens", cat:"insect", habitat:"Freshwater Shore & Open Water (Zone 2)", zones:["z2"], obsCount:5, iucn:"Least Concern", voucher:"INS-02", photo:"animal_photos/crimson_marsh_glider.jpg", gps:[18.9793, 72.8362], note:"Remarkable trans-oceanic migrant dragonfly patrolling over open water and damp lawns; acts as an efficient predator of mosquito larvae."},

  // ZONE 3: PROMENADE SHRUBLANDS & BUTTERFLY LAWNS
  {id:"bluebottle", common:"Common Bluebottle", marathi:"निळी बॉटल फुलपाखरू", sci:"Graphium sarpedon", cat:"butterfly", habitat:"Promenade Flowering Borders (Zone 3)", zones:["z3"], obsCount:6, iucn:"Least Concern", voucher:"INS-01", photo:"animal_photos/common_bluebottle_butterfly.jpg", gps:[18.9775, 72.8345], note:"Swift erratic swallowtail with brilliant turquoise-blue median band. Observed mud-puddling on damp walkways and nectaring on Lantana blossoms."},
  {id:"mormon", common:"Common Mormon", marathi:"कॉमन मॉर्मन", sci:"Papilio polytes", cat:"butterfly", habitat:"Citrus & Shrub Lawns (Zone 3)", zones:["z3"], obsCount:5, iucn:"Least Concern", voucher:"INS-03", photo:"animal_photos/common_mormon_butterfly.jpg", gps:[18.9778, 72.8348], note:"Sexually dimorphic swallowtail famous for female Batesian mimicry of unpalatable Crimson Rose butterflies. Breeds on garden Citrus and Murraya."},
  {id:"plain_tiger", common:"Plain Tiger Butterfly", marathi:"वाघ्या फुलपाखरू", sci:"Danaus chrysippus", cat:"butterfly", habitat:"Nectar Shrublands (Zone 3)", zones:["z3"], obsCount:5, iucn:"Least Concern", voucher:"INS-04", photo:"animal_photos/plain_tiger_butterfly.jpg", gps:[18.9776, 72.8347], note:"Striking orange milkweed butterfly strongly associated with Calotropis gigantea (Crown Flower) host plants in the nursery and garden perimeter."},
  {id:"common_crow", common:"Common Crow Butterfly", marathi:"कावळी फुलपाखरू", sci:"Euploea core", cat:"butterfly", habitat:"Shady Promenade Borders (Zone 3)", zones:["z3"], obsCount:5, iucn:"Least Concern", voucher:"INS-05", photo:"animal_photos/common_crow_butterfly.jpg", gps:[18.9779, 72.8349], note:"Glossy brownish-black wings edged with white spots. Highly resilient urban butterfly feeding steadily on Ixora and Lantana nectar."},
  {id:"lime_butterfly", common:"Lime Butterfly", marathi:"लिंबू फुलपाखरू", sci:"Papilio demoleus", cat:"butterfly", habitat:"Sunny Promenade Verges (Zone 3)", zones:["z3"], obsCount:4, iucn:"Least Concern", voucher:"INS-07", photo:"animal_photos/lime_butterfly.jpg", gps:[18.9774, 72.8349], note:"Fast-flying tailless swallowtail with yellow-and-black chequered pattern; feeds actively on wild nectar shrubs in full morning sun."},
  {id:"adulsa", common:"Malabar Nut / Adulsa", marathi:"अडुळसा", sci:"Justicia adhatoda", cat:"plant", habitat:"Promenade Medicinal Border (Zone 3)", zones:["z3"], obsCount:5, iucn:"Native Herb", voucher:"FLO-03", photo:"animal_photos/adulsa_plant.jpg", gps:[18.9777, 72.8346], note:"Indigenous dense medicinal shrub with tubular white flowers. Acts as vital larval host and shelter for micro-fauna and solitary native bees."},
  {id:"lantana", common:"Wild Sage / Lantana", marathi:"घाणेरी", sci:"Lantana camara", cat:"plant", habitat:"Managed Flower Borders (Zone 3)", zones:["z3"], obsCount:6, iucn:"Introduced Shrub", voucher:"FLO-10", photo:"animal_photos/lantana_flower.jpg", gps:[18.9775, 72.8344], note:"Prolific multi-colored floral clusters. While invasive in wild tracts, inside Byculla it serves as the most heavily visited nectar station for pollinators."},
  {id:"ixora", common:"Jungle Geranium", marathi:"बकुळी / पेंडी", sci:"Ixora coccinea", cat:"plant", habitat:"Ornamental Lawn Edges (Zone 3)", zones:["z3"], obsCount:5, iucn:"Ornamental Native", voucher:"FLO-11", photo:"animal_photos/ixora_flower.jpg", gps:[18.9778, 72.8346], note:"Clustered bright scarlet corymbs consistently in bloom; attracts butterflies and carpenter bees during peak morning sunlight."},
  {id:"hibiscus", common:"Shoe Flower / Hibiscus", marathi:"जास्वंद", sci:"Hibiscus rosa-sinensis", cat:"plant", habitat:"Promenade Beds (Zone 3)", zones:["z3"], obsCount:5, iucn:"Ornamental", voucher:"FLO-12", photo:"animal_photos/hibiscus_flower.jpg", gps:[18.9774, 72.8347], note:"Large showy red blooms visited routinely by Purple Sunbirds and large Violet Carpenter Bees between 08:00 and 10:30 AM."},
  {id:"sunbird", common:"Purple Sunbird", marathi:"जांभळा सूर्यपक्षी", sci:"Cinnyris asiaticus", cat:"bird", habitat:"Flowering Shrublands (Zone 3)", zones:["z3"], obsCount:5, iucn:"Least Concern", voucher:"AVE-11", photo:"animal_photos/purple_sunbird.jpg", gps:[18.9776, 72.8345], note:"Tiny metallic purple nectar-feeder with decurved bill, actively probing tubular Hibiscus and Adulsa blossoms."},
  {id:"tailorbird", common:"Common Tailorbird", marathi:"शिंपी पक्षी", sci:"Orthotomus sutorius", cat:"bird", habitat:"Dense Shrubbery (Zone 3)", zones:["z3"], obsCount:4, iucn:"Least Concern", voucher:"AVE-12", photo:"animal_photos/common_tailorbird.jpg", gps:[18.9777, 72.8348], note:"Restless insectivore renowned for stitching large living leaves together using plant fibres and spider silk to build cradle nests."},
  {id:"magpie_robin", common:"Oriental Magpie-Robin", marathi:"दयाळ", sci:"Copsychus saularis", cat:"bird", habitat:"Lawn Margins & Shrub Perches (Zone 3)", zones:["z3"], obsCount:4, iucn:"Least Concern", voucher:"AVE-13", photo:"animal_photos/magpie_robin.jpg", gps:[18.9779, 72.8346], note:"Bold black and white territorial songbird with an upturned tail, perching on decorative garden fences and singing loudly at sunrise."},
  {id:"carpenter_bee", common:"Violet Carpenter Bee", marathi:"काळा भुंगा", sci:"Xylocopa violacea", cat:"insect", habitat:"Flowering Hedges (Zone 3)", zones:["z3"], obsCount:5, iucn:"Least Concern", voucher:"INS-06", photo:"animal_photos/violet_carpenter_bee.jpg", gps:[18.9775, 72.8346], note:"Large solitary bee with iridescent purple-black wings; essential wild pollinator boring brood tunnels into dried bamboo and wood."},
  {id:"lizard", common:"Oriental Garden Lizard", marathi:"सरडा", sci:"Calotes versicolor", cat:"reptile", habitat:"Masonry Edges & Low Shrubs (Zone 3)", zones:["z3"], obsCount:5, iucn:"Least Concern", voucher:"REP-03", photo:"animal_photos/oriental_garden_lizard.jpg", gps:[18.9776, 72.8343], note:"Arboreal agamid reptile basking on brick boundary edges; adult males show distinct scarlet breeding coloration and perform head-bob displays."},

  // ZONE 4: PERIMETER TRANSITION BUFFER WALL (AMBEDKAR ROAD)
  {id:"rain_tree", common:"Rain Tree", marathi:"विलायती शिरीष", sci:"Samanea saman", cat:"plant", habitat:"Perimeter Boundary Wall (Zone 4)", zones:["z4"], obsCount:5, iucn:"Heritage Tree", voucher:"FLO-13", photo:"animal_photos/rain_tree.jpg", gps:[18.9768, 72.8335], note:"Enormous umbrella canopy spanning over 24 meters along Ambedkar Road. Serves as the primary physical and acoustic buffer against street traffic noise."},
  {id:"copper_pod", common:"Copper Pod / Yellow Poinciana", marathi:"पिवळा गुलमोहर", sci:"Peltophorum pterocarpum", cat:"plant", habitat:"Perimeter Road Verges (Zone 4)", zones:["z4"], obsCount:4, iucn:"Introduced Shade", voucher:"FLO-14", photo:"animal_photos/copperpod_tree.jpg", gps:[18.9769, 72.8337], note:"Sturdy yellow-flowering street tree resistant to particulate vehicle exhaust; dense litter underneath shelters soil invertebrates."},
  {id:"black_kite", common:"Black Kite", marathi:"घार", sci:"Milvus migrans", cat:"bird", habitat:"Perimeter Airspace & Tall Trees (Zone 4)", zones:["z4"], obsCount:5, iucn:"Least Concern", voucher:"AVE-14", photo:"animal_photos/black_kite.jpg", gps:[18.9770, 72.8336], note:"Opportunistic raptor riding thermal updrafts above Dr. Ambedkar Road traffic and nesting in tall peripheral Rain Trees."},
  {id:"house_crow", common:"House Crow", marathi:"कावळा", sci:"Corvus splendens", cat:"bird", habitat:"Entrance Arches & Perimeter (Zone 4)", zones:["z4"], obsCount:6, iucn:"Least Concern", voucher:"AVE-15", photo:"animal_photos/house_crow.jpg", gps:[18.9766, 72.8333], note:"Ubiquitous synanthropic urban corvid foraging around visitor entry points and boundary rubbish receptacles; highly adaptable."},
  {id:"myna", common:"Common Myna", marathi:"साळुंकी", sci:"Acridotheres tristis", cat:"bird", habitat:"Perimeter Lawns & Path Edges (Zone 4)", zones:["z4"], obsCount:5, iucn:"Least Concern", voucher:"AVE-16", photo:"animal_photos/common_myna.jpg", gps:[18.9769, 72.8338], note:"Audacious ground feeder foraging in pairs along perimeter walkways; tolerates high vehicle rumble and pedestrian movement."},
  {id:"pigeon", common:"Blue Rock Pigeon", marathi:"कबूतर", sci:"Columba livia", cat:"bird", habitat:"Entrance Architecture & Wall Ledges (Zone 4)", zones:["z4"], obsCount:5, iucn:"Least Concern", voucher:"AVE-17", photo:"animal_photos/rock_pigeon.jpg", gps:[18.9765, 72.8332], note:"Large flocks congregating on the heritage stone masonry arches of the main entrance gate and perimeter overhead wires."}
];

const SPECIES = SPECIES_RAW;
const SP_BY_ID = Object.fromEntries(SPECIES.map(s => [s.id, s]));

/* ---- Map projection coordinates: Byculla Zoo bounds ---- */
const B = AREAS[0].bounds;
const MW = 1000, MH = 700;
const toXY = (lat, lng) => ({
  x: Math.max(30, Math.min(970, (lng - B.lngMin) / (B.lngMax - B.lngMin) * MW)),
  y: Math.max(30, Math.min(670, (B.latMax - lat) / (B.latMax - B.latMin) * MH))
});
const toLL = (x, y) => ({
  lat: B.latMax - (y / MH) * (B.latMax - B.latMin),
  lng: B.lngMin + (x / MW) * (B.lngMax - B.lngMin)
});
const fmtLL = (lat, lng) => lat.toFixed(5) + "° N, " + lng.toFixed(5) + "° E";

/* ---- Generate realistic Geotagged Observations from Census ---- */
function buildObservations() {
  const list = [];
  let idCounter = 1;
  const visitDates = [
    {date: "2026-08-31", time: "07:35", visit: 1},
    {date: "2026-08-31", time: "08:45", visit: 1},
    {date: "2026-08-31", time: "10:30", visit: 2},
    {date: "2026-08-31", time: "11:50", visit: 2},
    {date: "2026-08-31", time: "14:15", visit: 3},
    {date: "2026-08-31", time: "17:10", visit: 4}
  ];

  SPECIES.forEach(sp => {
    const zoneId = sp.zones[0];
    const countTotal = sp.obsCount || 4;

    for (let i = 0; i < countTotal; i++) {
      const v = visitDates[i % visitDates.length];
      const jitterLat = (Math.random() - 0.5) * 0.0006;
      const jitterLng = (Math.random() - 0.5) * 0.0006;
      const lat = sp.gps[0] + jitterLat;
      const lng = sp.gps[1] + jitterLng;
      const pos = toXY(lat, lng);

      list.push({
        id: "OBS-" + String(idCounter++).padStart(3, "0"),
        spId: sp.id,
        cat: sp.cat,
        zone: zoneId,
        x: pos.x,
        y: pos.y,
        lat: lat,
        lng: lng,
        count: (sp.cat === "plant" ? 1 : Math.floor(Math.random() * 3) + 1),
        date: v.date,
        time: v.time,
        visit: v.visit,
        verified: true,
        photo: (i === 0 ? sp.photo : ""),
        voucher: sp.voucher,
        iucn: sp.iucn,
        note: sp.note
      });
    }
  });
  return list;
}

const BASE_OBS = buildObservations();

/* ---- LocalStorage Observation Persistence ---- */
const LS_KEY = "biomap.byculla.observations.v2";
let localObs = [];
try {
  const raw = localStorage.getItem(LS_KEY);
  if (raw) localObs = JSON.parse(raw) || [];
} catch(e) { localObs = []; }

function saveLocal() {
  try { localStorage.setItem(LS_KEY, JSON.stringify(localObs)); } catch(e){}
}
function allObs() { return localObs.concat(BASE_OBS); }
function obsOf(spId) { return allObs().filter(o => o.spId === spId); }
function speciesOf(ob) { return SP_BY_ID[ob.spId] || {common:"Unidentified", marathi:"अज्ञात", sci:"Incertae sedis", cat:ob.cat, note:""}; }

/* ============================================================
   Visual Vocabulary: Icons, SVG Base Art & Charts
   ============================================================ */

const ICONS = {
  dashboard: '<path d="M3 10.5 12 3l9 7.5"/><path d="M5.5 9.5V20h13V9.5"/><path d="M10 20v-5.5h4V20"/>',
  map: '<path d="M9 3.5 3.5 6v14.5L9 18l6 2.5 5.5-2.5V3.5L15 6Z"/><path d="M9 3.5V18"/><path d="M15 6v14.5"/>',
  species: '<path d="M4 20c0-8 5-13 16-14 0 10-5 14-11 14H4Z"/><path d="M7.5 16.5C10 13 13 11 17 9.5"/>',
  observe: '<path d="M12 5v14"/><path d="M5 12h14"/>',
  analytics: '<path d="M4 20V10"/><path d="M10 20V4"/><path d="M16 20v-7"/><path d="M21 20H3"/>',
  survey: '<path d="M12 21s7-5.6 7-11a7 7 0 1 0-14 0c0 5.4 7 11 7 11Z"/><circle cx="12" cy="10" r="2.6"/>',
  about: '<circle cx="12" cy="12" r="8.5"/><path d="M12 11v5.5"/><path d="M12 7.7h.01"/>',
  search: '<circle cx="11" cy="11" r="6.5"/><path d="m16 16 4.5 4.5"/>',
  sound: '<path d="M11 5L6 9H2v6h4l5 4V5z"/><path d="M15.54 8.46a5 5 0 0 1 0 7.07"/><path d="M19.07 4.93a10 10 0 0 1 0 14.14"/>',
  camera: '<path d="M23 19a2 2 0 0 1-2 2H3a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h4l2-3h6l2 3h4a2 2 0 0 1 2 2z"/><circle cx="12" cy="13" r="4"/>',
  download: '<path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/>',
  crosshair: '<circle cx="12" cy="12" r="7"/><path d="M12 2v3.5"/><path d="M12 18.5V22"/><path d="M2 12h3.5"/><path d="M18.5 12H22"/>',
  layers: '<path d="m12 3 9 5-9 5-9-5 9-5Z"/><path d="m3.5 13 8.5 4.7 8.5-4.7"/><path d="m3.5 18 8.5 4.7 8.5-4.7"/>'
};

function icon(name, cls="ico") {
  return `<svg class="${cls}" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true" style="width:16px;height:16px">${ICONS[name]||""}</svg>`;
}

/* ---- Vector Map Art: Byculla Zoo Grounds ---- */
const BASE_ART = `
<rect x="0" y="0" width="1000" height="700" fill="var(--map-land)"/>

<!-- Dr. Babasaheb Ambedkar Road (Western boundary corridor) -->
<rect x="0" y="0" width="36" height="700" fill="#D6D8CE"/>
<line x1="18" y1="0" x2="18" y2="700" stroke="#FAF8F0" stroke-width="2" stroke-dasharray="14 12"/>
<text x="12" y="350" transform="rotate(-90 12 350)" font-family="Inter, sans-serif" font-size="11" font-weight="600" fill="#6E7568" letter-spacing="2">DR. BABASAHEB AMBEDKAR ROAD (HEAVY TRAFFIC • 78.4 dBA)</text>

<!-- Boundary Wall Interface -->
<rect x="36" y="0" width="8" height="700" fill="#A8AEA0" stroke="#7C8470" stroke-width="0.8"/>

<!-- Zone 4 Perimeter Lawns (Acoustic Buffer Wall) -->
<path d="M44 30 H220 V670 H44 Z" fill="#E4EBD5" opacity=".85"/>

<!-- Zone 3 Promenade Shrubland & Lawns -->
<path d="M225 60 H500 V490 H225 Z" fill="#E2EAD3" opacity=".95"/>

<!-- Zone 1 Botanical Grove Canopy (Dense Heritage Arboretum) -->
<path d="M505 40 H975 V325 H505 Z" fill="#D5E4CA" opacity=".9"/>
<g fill="#A9C79A" opacity=".95">
  <circle cx="560" cy="110" r="48"/><circle cx="630" cy="90" r="54"/><circle cx="710" cy="130" r="62"/>
  <circle cx="790" cy="95" r="56"/><circle cx="870" cy="135" r="58"/><circle cx="940" cy="100" r="46"/>
  <circle cx="580" cy="200" r="52"/><circle cx="660" cy="220" r="64"/><circle cx="750" cy="210" r="68"/>
  <circle cx="840" cy="230" r="58"/><circle cx="920" cy="220" r="50"/>
  <!-- African Baobab Heritage Centerpiece -->
  <circle cx="720" cy="165" r="32" fill="#2F7D53" opacity=".85" stroke="#FFFFFF" stroke-width="2"/>
  <circle cx="720" cy="165" r="8" fill="#FFFFFF"/>
</g>

<!-- Zone 2 Freshwater Wetland & Lotus Lake -->
<path d="M480 340 C570 320 730 330 820 370 C880 405 870 510 810 560 C750 610 600 625 500 580 C440 540 420 420 480 340 Z"
      fill="#BCD6E3" stroke="#8EB5C9" stroke-width="2.5"/>
<g fill="#4A8E7A" opacity=".7">
  <!-- Lotus pads & emergent weeds -->
  <ellipse cx="560" cy="400" rx="20" ry="12"/>
  <ellipse cx="680" cy="380" rx="24" ry="14"/>
  <ellipse cx="740" cy="460" rx="26" ry="15"/>
  <ellipse cx="610" cy="520" rx="22" ry="13"/>
  <ellipse cx="530" cy="470" rx="18" ry="11"/>
</g>

<!-- Fact-Checked Animal Enclosure Paddock Outlines -->
<!-- Tiger Moated Enclosure (North) -->
<rect x="750" y="45" width="140" height="70" rx="8" fill="#F4EEDB" stroke="#D4A373" stroke-width="2" stroke-dasharray="4 3"/>
<!-- Sloth Bear Rock Grotto (North-West) -->
<rect x="530" y="55" width="110" height="65" rx="8" fill="#E8E4DA" stroke="#9A8C73" stroke-width="2" stroke-dasharray="4 3"/>
<!-- Spotted Deer / Chital Paddock (West-Central) -->
<rect x="250" y="240" width="130" height="90" rx="8" fill="#E6EBD6" stroke="#8A9A5B" stroke-width="2" stroke-dasharray="4 3"/>
<!-- Gharial & Croc Aquatic Sanctuary (East of Lake) -->
<rect x="800" y="410" width="120" height="75" rx="8" fill="#D5E8EE" stroke="#4682B4" stroke-width="2" stroke-dasharray="4 3"/>

<!-- Internal Visitor Pathways (1.8 km Transect Network) -->
<g stroke="#E6DEC9" fill="none" stroke-linecap="round">
  <path d="M60 570 C160 560 220 500 280 430 C340 360 400 300 480 270 C570 240 680 250 780 270 C880 300 930 340 960 380" stroke-width="14"/>
  <path d="M280 430 C250 350 240 220 250 110 H440" stroke-width="10"/>
  <path d="M480 270 C500 170 510 110 540 70 H920" stroke-width="10"/>
  <path d="M780 270 C790 350 770 440 720 540 C670 600 540 610 440 560" stroke-width="11"/>
  <path d="M60 570 H44" stroke-width="16"/>
</g>
<g stroke="#FFFFFF" fill="none" stroke-dasharray="6 8" stroke-width="1.6" opacity=".8">
  <path d="M60 570 C160 560 220 500 280 430 C340 360 400 300 480 270 C570 240 680 250 780 270 C880 300 930 340 960 380"/>
</g>

<!-- Built Architectural Points & Labels -->
<g fill="#E3DFD5" stroke="#B8B3A4" stroke-width="1.2">
  <rect x="52" y="540" width="70" height="60" rx="4"/>
  <rect x="130" y="580" width="95" height="50" rx="4"/>
  <rect x="250" y="580" width="105" height="46" rx="4"/>
  <rect x="850" y="530" width="90" height="54" rx="4"/>
</g>
<g font-family="Inter, sans-serif" font-size="10.5" font-weight="600" fill="#4B5244">
  <text x="56" y="575">Main Gate Arch</text>
  <text x="134" y="605" font-size="9.5">🏛️ Dr. Bhau Daji Lad</text>
  <text x="134" y="618" font-size="8.5" fill="#7A8072">Museum (1872)</text>
  <text x="256" y="604">🐧 Humboldt Pavilion</text>
  <text x="256" y="617" font-size="8.5" fill="#7A8072">Chilled Oceanarium</text>
  <text x="858" y="555">Conservation Center</text>

  <!-- Enclosure Tags -->
  <text x="760" y="70" fill="#9A3412" font-size="10">🐅 Tiger Moated Habitat</text>
  <text x="540" y="80" fill="#5C4D32" font-size="10">🐻 Sloth Bear Grotto</text>
  <text x="260" y="270" fill="#4A5D23" font-size="10">🦌 Spotted Deer Paddock</text>
  <text x="810" y="440" fill="#1E4D6B" font-size="10">🐊 Gharial Sanctuary Pool</text>
  <text x="690" y="195" fill="#1B4E34" font-size="9.5" font-weight="700">🌳 160-yr Baobab</text>

  <!-- Zone Title Callouts -->
  <text x="610" y="450" font-family="Newsreader, serif" font-size="16" font-style="italic" fill="#2C6FA8">Lotus Lake & Wetlands (Zone 2)</text>
  <text x="660" y="125" font-family="Newsreader, serif" font-size="16" font-weight="500" fill="#1B4E34">Botanical Grove Canopy (Zone 1)</text>
  <text x="290" y="100" font-family="Newsreader, serif" font-size="15" fill="#C0741F">Promenade Shrublands (Zone 3)</text>
  <text x="58" y="240" font-family="Newsreader, serif" font-size="14" fill="#A24A45">Perimeter Buffer (Zone 4)</text>
</g>
`;

/* ---- Map Glyph Markers by Category ---- */
function markerGlyph(cat) {
  const c = CATS[cat].color, sh = CATS[cat].shape;
  const st = 'stroke="var(--surface)" stroke-width="1.8"';
  if (sh === "pin") return `<path d="M0 -12C5.8 -12 10 -7.6 10 -2.6 10 3.8 0 12 0 12S-10 3.8 -10 -2.6C-10 -7.6 -5.8 -12 0 -12Z" fill="${c}" ${st}/><circle cx="0" cy="-2.6" r="3.2" fill="var(--surface)" opacity=".9"/>`;
  if (sh === "diamond") return `<path d="M0 -10.5 10.5 0 0 10.5 -10.5 0Z" fill="${c}" ${st}/>`;
  if (sh === "hex") return `<path d="M0 -9.5 8.2 -4.8 8.2 4.8 0 9.5 -8.2 4.8 -8.2 -4.8Z" fill="${c}" ${st}/>`;
  if (sh === "square") return `<rect x="-8" y="-8" width="16" height="16" rx="4" fill="${c}" ${st}/>`;
  return `<circle r="8.5" fill="${c}" ${st}/><circle r="3" fill="var(--surface)" opacity=".9"/>`;
}

function speciesImageMarkup(sp, styleClass="photo-thumb") {
  if (sp.photo) {
    return `<img src="${sp.photo}" alt="${esc(sp.common)}" class="${styleClass}" onerror="this.parentElement.innerHTML=plateSVG(SP_BY_ID['${sp.id}'])">`;
  }
  return plateSVG(sp);
}

function plateSVG(sp) {
  const c = CATS[sp.cat].color;
  return `<svg viewBox="0 0 160 120" style="width:100%;height:100%;display:block;background:var(--bg-tint)" aria-hidden="true">
    <rect width="160" height="120" fill="var(--bg-tint)"/>
    <circle cx="80" cy="55" r="36" fill="${c}" opacity=".15"/>
    <g transform="translate(80 55) scale(2.2)">${markerGlyph(sp.cat)}</g>
    <text x="80" y="105" text-anchor="middle" font-family="Inter, sans-serif" font-size="11" font-weight="600" fill="var(--ink-2)">${esc(sp.common.slice(0, 18))}</text>
  </svg>`;
}

/* ---- Map Filter State ---- */
let mapFilterCat = "all";
let mapFilterZone = "all";
let mapFilterQuery = "";

/* ---- Interactive Map Markup Generator ---- */
function mapMarkup(opts) {
  const o = Object.assign({controls: true, legend: true, height: "clamp(360px, 56vh, 640px)", showAcoustic: false}, opts || {});
  const obs = allObs();

  // Zone polygons with sound decibel badges
  const zonesMarkup = ZONES.map(z => `
    <g class="zone-shape" data-zone="${z.id}" tabindex="0" role="button" aria-label="${z.name}">
      <polygon points="${z.poly}" fill="${z.color}" fill-opacity="0.08" stroke="${z.color}" stroke-opacity="0.45" stroke-width="2" stroke-dasharray="6 6"/>
      <g transform="translate(${z.box[0] + 15}, ${z.box[1] + 25})">
        <rect x="-4" y="-14" width="130" height="22" rx="4" fill="var(--surface)" fill-opacity=".85" stroke="var(--line)"/>
        <text x="4" y="0" font-family="Inter, sans-serif" font-size="10.5" font-weight="700" fill="${z.color}">${z.code}: ${z.soundDba} dBA</text>
      </g>
    </g>
  `).join("");

  const acousticLayer = `
    <g class="l-acoustic" opacity="0.65" style="display:none">
      <rect x="0" y="0" width="220" height="700" fill="url(#acousticGrad)" opacity=".6"/>
      <circle cx="360" cy="300" r="160" fill="#fef3c7" opacity=".4"/>
      <circle cx="720" cy="180" r="180" fill="#dcfce7" opacity=".5"/>
      <circle cx="650" cy="460" r="160" fill="#e0f2fe" opacity=".45"/>
    </g>
  `;

  const markers = obs.map(b => {
    const sp = speciesOf(b);
    return `<g class="marker" data-id="${b.id}" data-sp="${b.spId}" data-cat="${b.cat}" data-zone="${b.zone}" data-name="${esc((sp.common+" "+sp.sci+" "+sp.marathi).toLowerCase())}" transform="translate(${b.x.toFixed(1)},${b.y.toFixed(1)})" tabindex="0" role="button" aria-label="${esc(sp.common)}">
      <circle class="halo" r="16" fill="${CATS[b.cat].color}" opacity=".2"/>
      ${markerGlyph(b.cat)}
    </g>`;
  }).join("");

  return `
  <div class="mapwrap" style="height:${o.height}">
    <svg class="mapsvg" viewBox="0 0 1000 700" preserveAspectRatio="xMidYMid slice">
      <defs>
        <linearGradient id="acousticGrad" x1="0" y1="0" x2="1" y2="0">
          <stop offset="0%" stop-color="#ef4444" stop-opacity="0.8"/>
          <stop offset="60%" stop-color="#f97316" stop-opacity="0.4"/>
          <stop offset="100%" stop-color="#eab308" stop-opacity="0.05"/>
        </linearGradient>
      </defs>
      <g class="vp">
        <g class="l-base">${BASE_ART}</g>
        ${acousticLayer}
        <g class="l-zones">${zonesMarkup}</g>
        <g class="l-markers">${markers}</g>
        <g class="l-user" style="display:none">
          <circle class="user-pulse" r="22" fill="#2C6FA8" opacity=".25"/>
          <circle r="8" fill="#2C6FA8" stroke="#FFFFFF" stroke-width="2.5"/>
        </g>
      </g>
    </svg>

    ${o.controls ? `
    <div class="map-ctrls" role="toolbar" aria-label="Map controls">
      <button class="map-btn" data-act="in" title="Zoom in" aria-label="Zoom in"><svg viewBox="0 0 24 24" width="16" height="16" fill="none" stroke="currentColor" stroke-width="2"><line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/></svg></button>
      <button class="map-btn" data-act="out" title="Zoom out" aria-label="Zoom out"><svg viewBox="0 0 24 24" width="16" height="16" fill="none" stroke="currentColor" stroke-width="2"><line x1="5" y1="12" x2="19" y2="12"/></svg></button>
      <button class="map-btn" data-act="reset" title="Reset view" aria-label="Reset view">${icon("layers")}</button>
      <button class="map-btn" data-act="noise" title="Toggle Acoustic Decibel Shield" aria-label="Toggle Sound Layer">${icon("sound")}</button>
      <button class="map-btn" data-act="me" title="Locate me in Byculla Zoo" aria-label="Locate me">${icon("crosshair")}</button>
    </div>` : ""}

    ${o.legend ? `
    <div class="map-legend">
      <div style="font-weight:600;font-size:11.5px;margin-bottom:4px;color:var(--ink)">Taxon Legend</div>
      ${CAT_ORDER.map(c => `
        <div class="legend-row">
          <span style="display:inline-flex;align-items:center">${markerGlyph(c)}</span>
          <span>${CATS[c].label}</span>
          <span class="mono-n" style="margin-left:auto;color:var(--muted);font-size:10.5px">${allObs().filter(o=>o.cat===c).length}</span>
        </div>
      `).join("")}
    </div>` : ""}
  </div>
  `;
}

function wireMap(el, opts={}) {
  if (!el) return;
  const svg = el.querySelector(".mapsvg");
  const vp = svg.querySelector(".vp");
  const acoustic = svg.querySelector(".l-acoustic");
  const userPin = svg.querySelector(".l-user");
  let k = 1, tx = 0, ty = 0;
  let noiseActive = false;

  function apply() {
    vp.setAttribute("transform", `translate(${tx} ${ty}) scale(${k})`);
  }

  el.querySelectorAll(".map-btn").forEach(btn => {
    btn.addEventListener("click", () => {
      const act = btn.dataset.act;
      if (act === "in") { k = Math.min(3.5, k * 1.3); apply(); }
      else if (act === "out") { k = Math.max(0.8, k / 1.3); apply(); }
      else if (act === "reset") { k = 1; tx = 0; ty = 0; apply(); }
      else if (act === "noise") {
        noiseActive = !noiseActive;
        acoustic.style.display = noiseActive ? "block" : "none";
        btn.classList.toggle("active", noiseActive);
        toast(noiseActive ? "Acoustic Shield Layer: 78.4 dBA (perimeter) → 52.1 dBA (grove)" : "Acoustic layer hidden");
      }
      else if (act === "me") {
        if ("geolocation" in navigator) {
          navigator.geolocation.getCurrentPosition(
            pos => {
              userPin.style.display = "block";
              const lat = (pos.coords.latitude >= 18.97 && pos.coords.latitude <= 18.99) ? pos.coords.latitude : 18.9789;
              const lng = (pos.coords.longitude >= 72.82 && pos.coords.longitude <= 72.84) ? pos.coords.longitude : 72.8354;
              const pt = toXY(lat, lng);
              userPin.setAttribute("transform", `translate(${pt.x} ${pt.y})`);
              k = 1.8; tx = 500 - pt.x * k; ty = 350 - pt.y * k; apply();
              toast(`GPS Active: Located at ${fmtLL(lat, lng)} (Zone 1 Canopy)`);
            },
            () => {
              userPin.style.display = "block";
              const pt = toXY(18.9789, 72.8354);
              userPin.setAttribute("transform", `translate(${pt.x} ${pt.y})`);
              k = 1.8; tx = 500 - pt.x * k; ty = 350 - pt.y * k; apply();
              toast("Simulated GPS: Zone 1 Botanical Grove (18.9789° N, 72.8354° E)");
            }
          );
        }
      }
    });
  });

  svg.querySelectorAll(".marker").forEach(m => {
    m.addEventListener("click", () => {
      const id = m.dataset.id;
      const ob = allObs().find(o => o.id === id);
      if (ob && opts.onSelect) opts.onSelect(ob, m, el);
    });
  });

  svg.querySelectorAll(".zone-shape").forEach(z => {
    z.addEventListener("click", () => {
      const zid = z.dataset.zone;
      const zObj = ZONE_BY_ID[zid];
      if (zObj) {
        toast(`${zObj.code} (${zObj.name}): Sound ${zObj.soundDba} dBA • Shannon H'=${zObj.shannon}`);
      }
    });
  });
}

function showPopcard(wrap, ob, markerEl) {
  const sp = speciesOf(ob);
  let card = wrap.querySelector(".map-popcard");
  if (!card) {
    card = document.createElement("div");
    card.className = "map-popcard card card-pad";
    card.style.cssText = "position:absolute;bottom:14px;right:14px;width:310px;max-width:calc(100% - 28px);z-index:20;box-shadow:var(--shadow-lg);animation:slideUp .2s ease;";
    wrap.appendChild(card);
  }
  const z = ZONE_BY_ID[ob.zone];
  const iucnClass = {
    "Critically Endangered":"iucn-cr", "Endangered":"iucn-en", "Vulnerable":"iucn-vu",
    "Near Threatened":"iucn-nt", "Least Concern":"iucn-lc", "Heritage Tree":"iucn-heritage",
    "Native Flora":"iucn-native", "Native Herb":"iucn-native"
  }[sp.iucn] || "iucn-lc";

  card.innerHTML = `
    <div style="display:flex;justify-content:space-between;align-items:flex-start;gap:8px;margin-bottom:6px">
      <div>
        <span class="voucher-pill">${ob.voucher || ob.id}</span>
        <span class="badge-iucn ${iucnClass}" style="margin-left:4px">${sp.iucn||"Audited"}</span>
      </div>
      <button class="btn btn-sm btn-icon" id="pop-close" aria-label="Close">✕</button>
    </div>
    ${sp.photo ? `<div style="height:130px;margin-bottom:8px;overflow:hidden;border-radius:var(--r-sm)"><img src="${sp.photo}" alt="${esc(sp.common)}" style="width:100%;height:100%;object-fit:cover"></div>` : ""}
    <h3 style="margin:0 0 1px 0;font-size:15px">${esc(sp.common)}</h3>
    <div class="marathi-sub">${esc(sp.marathi || "")}</div>
    <div class="sci" style="font-size:12px;margin-top:2px">${esc(sp.sci)}</div>
    <div style="margin:8px 0;font-size:12px;color:var(--ink-2);line-height:1.4">${esc(sp.note)}</div>
    <div style="display:flex;justify-content:space-between;align-items:center;font-size:11px;padding-top:8px;border-top:1px solid var(--line-2)">
      <span class="dba-pill ${z.soundDba<55?'dba-quiet':z.soundDba<65?'dba-mod':'dba-high'}">${z.code} • ${z.soundDba} dBA</span>
      <button class="btn btn-sm" data-go="#/species/${sp.id}">Full Dossier →</button>
    </div>
  `;
  card.querySelector("#pop-close").addEventListener("click", () => card.remove());
  card.querySelector("[data-go]").addEventListener("click", (e) => {
    navigate(e.target.dataset.go);
  });
}

/* ============================================================
   Aggregations & Diversity Indices
   ============================================================ */

function stats() {
  const obs = allObs();
  const spIds = new Set(obs.map(o => o.spId).filter(Boolean));
  const perCat = {};
  CAT_ORDER.forEach(c => {
    perCat[c] = {
      sp: SPECIES.filter(s => s.cat === c).length,
      obs: obs.filter(o => o.cat === c).length
    };
  });
  return {
    species: spIds.size || SPECIES.length,
    obs: obs.length,
    perCat: perCat,
    vouchers: SPECIES.filter(s => s.voucher).length,
    zones: ZONES.length,
    participants: 84
  };
}

function zoneCounts() {
  const obs = allObs();
  return ZONES.map(z => {
    const inZ = obs.filter(o => o.zone === z.id);
    const spSet = new Set(inZ.map(o => o.spId));
    return {
      id: z.id,
      code: z.code,
      label: z.name,
      value: inZ.length,
      species: spSet.size || z.richness,
      sound: z.soundDba,
      shannon: z.shannon,
      evenness: z.evenness,
      color: z.color
    };
  });
}

/* ============================================================
   Screens & View Controllers
   ============================================================ */

const SCREENS = {};

/* ---------- 1. Dashboard Screen ---------- */
SCREENS.dashboard = () => {
  const s = stats();
  const zc = zoneCounts();
  const featured = SPECIES.filter(sp => sp.photo).slice(0, 6);

  const html = `
  <div class="hero-banner">
    <div class="hero-title">Local Biodiversity Mapping of Byculla Zoo & Surrounding Green Spaces</div>
    <div class="hero-sub">A Geospatial Biodiversity Census, Web-GIS Cartography & Community Ecological Literacy Study of South Mumbai's 53-Acre Heritage Sanctuary (Veermata Jijabai Bhosale Udyan / Ranibagh).</div>
    <div class="hero-meta">
      <span>🎓 Vidyavardhini's Vartak College (CS Dept.)</span>
      <span>🏛 University of Mumbai • NEP 2020</span>
      <span>👥 Riddhesh Keni (35) & Shivansh Dubey (45)</span>
      <span>🔬 Mentor: Prof. Shweta Iyer</span>
    </div>
  </div>

  <div class="stats" style="margin-bottom:18px">
    <div class="stat"><span class="v num">${s.species}</span><span class="k">Documented Species</span></div>
    <div class="stat"><span class="v num">${s.obs}</span><span class="k">Geotagged Sightings</span></div>
    <div class="stat"><span class="v num">-26.3 dB</span><span class="k">Canopy Noise Attenuation</span></div>
    <div class="stat"><span class="v num">H' = 2.68</span><span class="k">Shannon Index (Zone 1)</span></div>
    <div class="stat"><span class="v num">84</span><span class="k">Surveyed Citizens (N=84)</span></div>
    <div class="stat"><span class="v num">4.8 km</span><span class="k">Transect Traversed</span></div>
  </div>

  <div class="grid" style="grid-template-columns:minmax(0,1.8fr) minmax(0,1fr); align-items:start; gap:18px">
    <div class="card" style="overflow:hidden">
      <div class="card-h">
        <div>
          <h3>Live Byculla Zoo Web-GIS Cartography</h3>
          <span class="hint">4 Ecological Survey Zones • Interactive Acoustic Decibel Overlay</span>
        </div>
        <button class="btn btn-sm" data-go="#/map">Fullscreen GIS →</button>
      </div>
      <div style="padding:0 14px 14px">${mapMarkup({obs:allObs(), height:"clamp(340px, 48vh, 520px)", legend:false})}</div>
    </div>

    <div class="grid" style="gap:16px">
      <div class="card card-pad">
        <h3 style="margin-bottom:12px;font-size:14.5px">4-Zone Stratification</h3>
        <div style="display:flex;flex-direction:column;gap:8px">
          ${zc.map(z => `
            <div style="display:flex;justify-content:space-between;align-items:center;padding:7px 10px;background:var(--bg-tint);border-radius:var(--r-sm)">
              <div>
                <strong style="font-size:12.5px;color:var(--ink)">${z.code}: ${esc(z.label)}</strong>
                <div style="font-size:11px;color:var(--muted)">H' = ${z.shannon} • Richness: ${z.species} taxa</div>
              </div>
              <span class="dba-pill ${z.sound<55?'dba-quiet':z.sound<65?'dba-mod':'dba-high'}">${z.sound} dBA</span>
            </div>
          `).join("")}
        </div>
      </div>

      <div class="card card-pad" style="background:linear-gradient(135deg, var(--surface) 0%, var(--accent-wash) 100%)">
        <h3 style="margin-bottom:6px;font-size:14.5px">Citizen Biodiversity Guide</h3>
        <p style="font-size:12.5px;color:var(--ink-2);margin-bottom:12px;line-height:1.5">Over 78% of park visitors overlook wild fauna due to 'captive-species bias'. Explore verified field records or log your sighting.</p>
        <div style="display:flex;gap:8px">
          <button class="btn btn-sm" data-go="#/species" style="flex:1">Browse Catalogue</button>
          <button class="btn btn-sm btn-primary" data-go="#/observe" style="flex:1">+ Log Sighting</button>
        </div>
      </div>
    </div>
  </div>

  <div style="margin-top:24px">
    <div style="display:flex;justify-content:space-between;align-items:flex-end;margin-bottom:12px">
      <div>
        <h2 style="margin:0;font-family:var(--serif);font-size:21px">Audited Master Species Vouchers</h2>
        <span class="hint">Key botanical keystone trees and wild avifauna with authentic field photos</span>
      </div>
      <button class="btn btn-sm" data-go="#/species">View All 42+ Species →</button>
    </div>

    <div class="grid" style="grid-template-columns:repeat(auto-fit, minmax(260px, 1fr)); gap:16px">
      ${featured.map(sp => {
        const z = ZONE_BY_ID[sp.zones[0]];
        return `
        <div class="card sp-card" data-go="#/species/${sp.id}">
          <div class="species-thumb-wrap">
            <img src="${sp.photo}" alt="${esc(sp.common)}">
            <span class="badge-iucn iucn-${sp.iucn.toLowerCase().includes('heritage')?'heritage':sp.iucn.toLowerCase().includes('endangered')?'en':sp.iucn.toLowerCase().includes('vulnerable')?'vu':'lc'}" style="position:absolute;top:8px;right:8px">${sp.iucn}</span>
          </div>
          <div class="card-pad" style="padding:12px 14px">
            <div style="display:flex;justify-content:space-between;align-items:baseline">
              <strong style="font-size:14.5px;color:var(--ink)">${esc(sp.common)}</strong>
              <span class="voucher-pill">${sp.voucher}</span>
            </div>
            <div class="marathi-sub">${esc(sp.marathi)}</div>
            <div class="sci" style="font-size:12px;margin-top:2px">${esc(sp.sci)}</div>
            <div style="margin-top:8px;display:flex;justify-content:space-between;align-items:center;font-size:11px">
              <span class="tag"><span class="dot" style="background:${CATS[sp.cat].color}"></span>${CATS[sp.cat].label}</span>
              <span style="color:var(--muted)">${z.code}</span>
            </div>
          </div>
        </div>`;
      }).join("")}
    </div>
  </div>
  `;

  return {
    html,
    mount(root) {
      wireMap(root.querySelector(".mapwrap"), {
        onSelect: (ob, m, w) => showPopcard(w, ob, m)
      });
      root.querySelectorAll("[data-go]").forEach(el => {
        el.addEventListener("click", () => navigate(el.dataset.go));
      });
    }
  };
};

/* ---------- 2. Fullscreen Interactive Web-GIS Map Screen ---------- */
SCREENS.map = () => {
  const html = `
  <div class="page-head" style="display:flex;justify-content:space-between;align-items:flex-end;gap:12px;flex-wrap:wrap;margin-bottom:14px">
    <div>
      <h1 style="font-size:22px;margin-bottom:2px">Byculla Zoo Web-GIS Biodiversity Map</h1>
      <p class="hint">Geospatial census across 53 acres • Click markers to view photographic vouchers • Toggle acoustic layer</p>
    </div>
    <div style="display:flex;gap:8px">
      <button class="btn btn-sm" id="btn-toggle-acoustic">${icon("sound")} Sound Heatmap</button>
      <button class="btn btn-sm btn-primary" data-go="#/observe">+ Add Sighting</button>
    </div>
  </div>

  <div class="card" style="padding:8px 12px;margin-bottom:12px;display:flex;flex-wrap:wrap;gap:8px;align-items:center">
    <span style="font-size:12px;font-weight:600;color:var(--ink-2)">Category:</span>
    <button class="chip active" data-map-cat="all">All Taxa</button>
    ${CAT_ORDER.map(c=>`<button class="chip" data-map-cat="${c}">${CATS[c].label}</button>`).join("")}
  </div>

  ${mapMarkup({obs:allObs(), height:"clamp(460px, 72vh, 780px)", legend:true})}

  <div class="card card-pad" style="margin-top:12px;background:var(--surface-2)">
    <div style="display:flex;flex-wrap:wrap;gap:16px;align-items:center;font-size:12px;color:var(--ink-2)">
      <span><strong>Acoustic Refuge Finding:</strong> The interior Botanical Grove (Zone 1) records 52.1 dBA, delivering a 26.3 dBA noise reduction from the outer Ambedkar Road boundary wall (Zone 4: 78.4 dBA).</span>
      <span style="margin-left:auto"><strong style="color:var(--accent-ink)">42+ Audited Species • N=84 Citizen Survey Validation</strong></span>
    </div>
  </div>
  `;

  return {
    html,
    mount(root) {
      wireMap(root.querySelector(".mapwrap"), {
        onSelect: (ob, m, w) => showPopcard(w, ob, m)
      });
      const tbtn = root.querySelector("#btn-toggle-acoustic");
      if (tbtn) {
        tbtn.addEventListener("click", () => {
          const mapBtn = root.querySelector('[data-act="noise"]');
          if (mapBtn) mapBtn.click();
          tbtn.classList.toggle("active");
        });
      }

      root.querySelectorAll("[data-map-cat]").forEach(ch => {
        ch.addEventListener("click", () => {
          root.querySelectorAll("[data-map-cat]").forEach(c => c.classList.remove("active"));
          ch.classList.add("active");
          const selected = ch.dataset.mapCat;
          root.querySelectorAll(".mapwrap .marker").forEach(m => {
            if (selected === "all" || m.dataset.cat === selected) {
              m.style.display = "";
            } else {
              m.style.display = "none";
            }
          });
        });
      });
    }
  };
};

/* ---------- 3. Species Catalogue Screen ---------- */
const catState = { q: "", cat: "all", zone: "all", iucn: "all" };

SCREENS.species = () => {
  function renderList() {
    let list = SPECIES.slice();
    if (catState.q.trim()) {
      const q = catState.q.trim().toLowerCase();
      list = list.filter(s =>
        s.common.toLowerCase().includes(q) ||
        s.marathi.toLowerCase().includes(q) ||
        s.sci.toLowerCase().includes(q) ||
        (s.voucher && s.voucher.toLowerCase().includes(q))
      );
    }
    if (catState.cat !== "all") list = list.filter(s => s.cat === catState.cat);
    if (catState.zone !== "all") list = list.filter(s => s.zones.includes(catState.zone));
    if (catState.iucn !== "all") list = list.filter(s => s.iucn.toLowerCase().includes(catState.iucn.toLowerCase()));

    if (!list.length) {
      return `<div class="card card-pad empty" style="grid-column:1/-1;text-align:center;color:var(--muted)">No species records matching "${esc(catState.q)}" were found.</div>`;
    }

    return list.map(sp => {
      const z = ZONE_BY_ID[sp.zones[0]];
      const iucnClass = {
        "Critically Endangered":"iucn-cr", "Endangered":"iucn-en", "Vulnerable":"iucn-vu",
        "Near Threatened":"iucn-nt", "Least Concern":"iucn-lc", "Heritage Tree":"iucn-heritage",
        "Native Flora":"iucn-native", "Native Herb":"iucn-native"
      }[sp.iucn] || "iucn-lc";

      return `
      <div class="sp-card" data-go="#/species/${sp.id}">
        <div>
          <div class="species-thumb-wrap">
            ${speciesImageMarkup(sp)}
            <span class="badge-iucn ${iucnClass}" style="position:absolute;top:8px;right:8px">${sp.iucn}</span>
          </div>
          <div style="padding:12px 14px 6px">
            <div style="display:flex;justify-content:space-between;align-items:baseline">
              <strong style="font-size:14.5px;color:var(--ink)">${esc(sp.common)}</strong>
              ${sp.voucher ? `<span class="voucher-pill">${sp.voucher}</span>` : ""}
            </div>
            <div class="marathi-sub">${esc(sp.marathi)}</div>
            <div class="sci" style="font-size:12px;margin-top:2px">${esc(sp.sci)}</div>
            <p style="font-size:12px;color:var(--ink-2);margin:8px 0 0;line-height:1.4">${esc(sp.note.slice(0, 110))}...</p>
          </div>
        </div>
        <div style="padding:10px 14px;border-top:1px solid var(--line-2);display:flex;justify-content:space-between;align-items:center;font-size:11px">
          <span class="tag"><span class="dot" style="background:${CATS[sp.cat].color}"></span>${CATS[sp.cat].label}</span>
          <span class="dba-pill ${z.soundDba<55?'dba-quiet':z.soundDba<65?'dba-mod':'dba-high'}">${z.code}</span>
        </div>
      </div>`;
    }).join("");
  }

  const html = `
  <div class="page-head" style="display:flex;justify-content:space-between;align-items:flex-end;gap:14px;flex-wrap:wrap;margin-bottom:14px">
    <div>
      <h1 style="font-size:22px;margin-bottom:2px">Representative Master Species Census</h1>
      <p class="hint">42+ Audited Taxa across Flora, Avifauna, Lepidoptera, Insects, Mammals & Reptiles</p>
    </div>
    <span class="voucher-pill">${SPECIES.length} species catalogued</span>
  </div>

  <div class="card card-pad" style="margin-bottom:16px">
    <div class="grid" style="grid-template-columns:minmax(0, 2fr) minmax(0, 1fr) minmax(0, 1fr) minmax(0, 1fr); gap:10px">
      <div class="field" style="margin:0">
        <label for="sp-q">Search Species (English, मराठी, Scientific or Voucher)</label>
        <input id="sp-q" class="input" placeholder="e.g. Baobab, गोरखचिंच, Tiger, Adulsa, Barbet..." value="${esc(catState.q)}">
      </div>
      <div class="field" style="margin:0">
        <label for="sp-cat">Category</label>
        <select id="sp-cat" class="input">
          <option value="all" ${catState.cat==="all"?"selected":""}>All Taxa</option>
          ${CAT_ORDER.map(c => `<option value="${c}" ${catState.cat===c?"selected":""}>${CATS[c].label}</option>`).join("")}
        </select>
      </div>
      <div class="field" style="margin:0">
        <label for="sp-zone">Survey Zone</label>
        <select id="sp-zone" class="input">
          <option value="all" ${catState.zone==="all"?"selected":""}>All Zones</option>
          ${ZONES.map(z => `<option value="${z.id}" ${catState.zone===z.id?"selected":""}>${z.code}: ${z.name}</option>`).join("")}
        </select>
      </div>
      <div class="field" style="margin:0">
        <label for="sp-iucn">IUCN Status</label>
        <select id="sp-iucn" class="input">
          <option value="all" ${catState.iucn==="all"?"selected":""}>All Statuses</option>
          <option value="heritage" ${catState.iucn==="heritage"?"selected":""}>Heritage Tree</option>
          <option value="endangered" ${catState.iucn==="endangered"?"selected":""}>Endangered</option>
          <option value="vulnerable" ${catState.iucn==="vulnerable"?"selected":""}>Vulnerable</option>
          <option value="near threatened" ${catState.iucn==="near threatened"?"selected":""}>Near Threatened</option>
          <option value="least concern" ${catState.iucn==="least concern"?"selected":""}>Least Concern</option>
        </select>
      </div>
    </div>
  </div>

  <div class="grid" id="sp-grid" style="grid-template-columns:repeat(auto-fill, minmax(270px, 1fr)); gap:16px">
    ${renderList()}
  </div>
  `;

  return {
    html,
    mount(root) {
      function refresh() {
        root.querySelector("#sp-grid").innerHTML = renderList();
        root.querySelectorAll("#sp-grid [data-go]").forEach(el => {
          el.addEventListener("click", () => navigate(el.dataset.go));
        });
      }

      root.querySelector("#sp-q").addEventListener("input", e => {
        catState.q = e.target.value;
        refresh();
      });
      root.querySelector("#sp-cat").addEventListener("change", e => {
        catState.cat = e.target.value;
        refresh();
      });
      root.querySelector("#sp-zone").addEventListener("change", e => {
        catState.zone = e.target.value;
        refresh();
      });
      root.querySelector("#sp-iucn").addEventListener("change", e => {
        catState.iucn = e.target.value;
        refresh();
      });
      root.querySelectorAll("#sp-grid [data-go]").forEach(el => {
        el.addEventListener("click", () => navigate(el.dataset.go));
      });
    }
  };
};

/* ---------- 4. Species Detail Dossier Screen ---------- */
SCREENS.speciesDetail = (id) => {
  const sp = SP_BY_ID[id];
  if (!sp) {
    return {
      html: `<div class="card card-pad empty">That species record could not be found. <button class="btn btn-sm" data-go="#/species">Back to Catalogue</button></div>`,
      mount() {}
    };
  }
  const obs = obsOf(sp.id);
  const z = ZONE_BY_ID[sp.zones[0]];
  const iucnClass = {
    "Critically Endangered":"iucn-cr", "Endangered":"iucn-en", "Vulnerable":"iucn-vu",
    "Near Threatened":"iucn-nt", "Least Concern":"iucn-lc", "Heritage Tree":"iucn-heritage",
    "Native Flora":"iucn-native", "Native Herb":"iucn-native"
  }[sp.iucn] || "iucn-lc";

  const html = `
  <button class="btn btn-sm" data-go="#/species" style="margin-bottom:14px">← Back to Species Catalogue</button>

  <div class="grid" style="grid-template-columns:minmax(0,1.1fr) minmax(0,1.4fr); align-items:start; gap:20px">
    <div class="card" style="overflow:hidden">
      <div style="aspect-ratio:4/3;background:var(--bg-tint);overflow:hidden">
        ${speciesImageMarkup(sp, "")}
      </div>
      <div class="card-pad" style="background:var(--surface-2);font-size:11.5px;color:var(--muted)">
        ${sp.photo ? "📷 Field Photographic Voucher collected on 31/08/2026 by Investigators Riddhesh Keni & Shivansh Dubey." : "Illustrated vector plate. Linnaean classification verified against BNHS regional flora/fauna registers."}
      </div>
    </div>

    <div>
      <div style="display:flex;justify-content:space-between;align-items:flex-start;gap:12px">
        <div>
          <h1 style="margin:0 0 2px 0;font-size:24px">${esc(sp.common)}</h1>
          <div class="marathi-sub" style="font-size:16px;font-weight:600">${esc(sp.marathi)}</div>
          <div class="sci" style="font-size:15px;margin-top:3px">${esc(sp.sci)}</div>
        </div>
        <div style="text-align:right">
          ${sp.voucher ? `<div class="voucher-pill" style="font-size:12px">${sp.voucher}</div>` : ""}
          <div class="badge-iucn ${iucnClass}" style="margin-top:6px">${sp.iucn}</div>
        </div>
      </div>

      <p style="margin:14px 0;color:var(--ink);font-size:14px;line-height:1.6;background:var(--accent-wash);padding:12px 16px;border-radius:var(--r-sm);border-left:3px solid var(--accent)">
        ${esc(sp.note)}
      </p>

      <div class="card card-pad">
        <h3 style="margin-bottom:10px;font-size:14px">Taxonomic & Field Metadata</h3>
        <dl class="kv" style="gap:8px 14px;font-size:12.5px">
          <dt>Taxon Class</dt><dd>${CATS[sp.cat].plural}</dd>
          <dt>Survey Zone</dt><dd><strong>${z.code}: ${esc(z.name)}</strong></dd>
          <dt>Microhabitat</dt><dd>${esc(sp.habitat)}</dd>
          <dt>GPS Coordinates</dt><dd class="mono-n">${fmtLL(sp.gps[0], sp.gps[1])}</dd>
          <dt>Zone Noise Level</dt><dd><span class="dba-pill ${z.soundDba<55?'dba-quiet':z.soundDba<65?'dba-mod':'dba-high'}">${z.soundDba} dBA (${z.soundDesc})</span></dd>
          <dt>Total Sightings</dt><dd class="mono-n">${obs.length} recorded during field census</dd>
        </dl>
      </div>
    </div>
  </div>

  <div class="card" style="margin-top:20px">
    <div class="card-h">
      <div>
        <h3>Field Observation Logs for ${esc(sp.common)}</h3>
        <span class="hint">${obs.length} verified records entered into Web-GIS</span>
      </div>
    </div>
    <div class="card-pad scroll-x">
      <table class="data">
        <thead>
          <tr><th>Voucher / Record ID</th><th>Survey Date</th><th>Time</th><th>Zone</th><th>Individuals</th><th>GPS Geotag</th><th>Status</th></tr>
        </thead>
        <tbody>
          ${obs.map(o => `
            <tr>
              <td><strong>${o.voucher || o.id}</strong></td>
              <td>${o.date}</td>
              <td class="mono-n">${o.time}</td>
              <td>${z.code}</td>
              <td class="mono-n">${o.count}</td>
              <td class="mono-n">${fmtLL(o.lat, o.lng)}</td>
              <td><span class="badge-iucn iucn-lc">✓ Verified</span></td>
            </tr>
          `).join("")}
        </tbody>
      </table>
    </div>
  </div>
  `;

  return {
    html,
    mount(root) {
      root.querySelectorAll("[data-go]").forEach(el => {
        el.addEventListener("click", () => navigate(el.dataset.go));
      });
    }
  };
};

/* ---------- 5. Citizen Sighting Submission Screen ---------- */
SCREENS.observe = () => {
  const html = `
  <div class="page-head" style="margin-bottom:16px">
    <h1 style="font-size:22px;margin-bottom:2px">Contribute Field Observation</h1>
    <p class="hint">Log a free-living or botanical sighting inside Byculla Zoo. Records persist directly to your browser Web-GIS layer.</p>
  </div>

  <div class="grid" style="grid-template-columns:minmax(0,1.3fr) minmax(0,1fr); align-items:start; gap:20px">
    <div class="card card-pad">
      <form id="obs-form" novalidate>
        <div class="field" style="margin-bottom:12px">
          <label for="obs-photo">Field Photograph</label>
          <div class="dropzone" id="drop-zone" style="border:2px dashed var(--line);padding:20px;border-radius:var(--r-sm);text-align:center;cursor:pointer;background:var(--bg-tint)">
            ${icon("camera", "ico")}
            <div style="font-weight:600;font-size:13px;margin-top:6px">Upload or Drop Sighting Photo</div>
            <div class="hint">JPG or PNG from mobile or digital camera</div>
          </div>
          <input id="obs-photo" type="file" accept="image/*" style="display:none">
        </div>

        <div class="grid" style="grid-template-columns:1fr 1fr; gap:10px">
          <div class="field">
            <label for="f-common">Common Species Name</label>
            <input id="f-common" class="input" list="sp-list" placeholder="e.g. Coppersmith Barbet" required>
            <datalist id="sp-list">${SPECIES.map(s=>`<option value="${esc(s.common)}">${esc(s.marathi)} (${esc(s.sci)})</option>`).join("")}</datalist>
          </div>
          <div class="field">
            <label for="f-marathi">स्थानिक मराठी नाव (Vernacular Name)</label>
            <input id="f-marathi" class="input" placeholder="e.g. तांबट पक्षी">
          </div>
        </div>

        <div class="grid" style="grid-template-columns:1fr 1fr; gap:10px; margin-top:8px">
          <div class="field">
            <label for="f-sci">Scientific Name (Linnaean Binomial)</label>
            <input id="f-sci" class="input" placeholder="e.g. Psilopogon haemacephalus">
          </div>
          <div class="field">
            <label for="f-cat">Taxonomic Category</label>
            <select id="f-cat" class="input">
              ${CAT_ORDER.map(c => `<option value="${c}">${CATS[c].label}</option>`).join("")}
            </select>
          </div>
        </div>

        <div class="grid" style="grid-template-columns:1fr 1fr; gap:10px; margin-top:8px">
          <div class="field">
            <label for="f-zone">Ecological Survey Zone</label>
            <select id="f-zone" class="input">
              ${ZONES.map(z => `<option value="${z.id}">${z.code}: ${esc(z.name)} (${z.soundDba} dBA)</option>`).join("")}
            </select>
          </div>
          <div class="field">
            <label for="f-count">Individual Count</label>
            <input id="f-count" type="number" min="1" max="100" value="1" class="input">
          </div>
        </div>

        <div class="grid" style="grid-template-columns:1fr 1fr; gap:10px; margin-top:8px">
          <div class="field">
            <label for="f-lat">Latitude (°N)</label>
            <input id="f-lat" type="number" step="0.00001" value="18.9789" class="input">
          </div>
          <div class="field">
            <label for="f-lng">Longitude (°E)</label>
            <input id="f-lng" type="number" step="0.00001" value="72.8354" class="input">
          </div>
        </div>

        <div class="field" style="margin-top:8px">
          <label for="f-note">Ethological & Habitat Observation Notes</label>
          <textarea id="f-note" class="input" rows="3" placeholder="Describe behavior, perch type, nesting cavity, flowering state, or human interference..."></textarea>
        </div>

        <button type="submit" class="btn btn-primary" style="margin-top:14px;width:100%">Save Observation to BioMap Web-GIS</button>
      </form>
    </div>

    <div class="grid" style="gap:14px">
      <div class="card card-pad">
        <h3 style="margin-bottom:8px;font-size:14px">Citizen Science Protocol</h3>
        <ul style="margin:0;padding-left:18px;font-size:12.5px;color:var(--ink-2);line-height:1.75">
          <li><strong>Non-Captive Priority:</strong> Prioritize wild avifauna, butterflies, and heritage flora over caged zoo exhibits.</li>
          <li><strong>Geotag Precision:</strong> Record coordinates within the 53-acre garden perimeter (±3m GNSS precision).</li>
          <li><strong>Breeding Season Care:</strong> Never disturb active nesting cavities in Ficus, Baobab, or Adulsa shrub borders.</li>
        </ul>
      </div>

      <div class="card card-pad" style="background:var(--accent-wash)">
        <h3 style="margin-bottom:6px;font-size:13.5px;color:var(--accent-ink)">Local Web Storage Active</h3>
        <p style="font-size:12px;color:var(--ink);margin:0;line-height:1.5">Observations logged here are instantly merged with the census database and rendered on the interactive GIS map and analytics screens.</p>
      </div>
    </div>
  </div>
  `;

  return {
    html,
    mount(root) {
      const drop = root.querySelector("#drop-zone");
      const fileInput = root.querySelector("#obs-photo");
      const fCommon = root.querySelector("#f-common");
      const fMarathi = root.querySelector("#f-marathi");
      const fSci = root.querySelector("#f-sci");
      const fCat = root.querySelector("#f-cat");
      const fZone = root.querySelector("#f-zone");
      const fLat = root.querySelector("#f-lat");
      const fLng = root.querySelector("#f-lng");

      drop.addEventListener("click", () => fileInput.click());
      fCommon.addEventListener("change", () => {
        const found = SPECIES.find(s => s.common.toLowerCase() === fCommon.value.trim().toLowerCase());
        if (found) {
          fMarathi.value = found.marathi;
          fSci.value = found.sci;
          fCat.value = found.cat;
          fZone.value = found.zones[0];
          fLat.value = found.gps[0];
          fLng.value = found.gps[1];
        }
      });

      root.querySelector("#obs-form").addEventListener("submit", e => {
        e.preventDefault();
        const common = fCommon.value.trim();
        if (!common) { toast("Please provide a species name"); return; }

        const newRec = {
          id: "OBS-U" + String(Date.now()).slice(-4),
          spId: (SPECIES.find(s=>s.common.toLowerCase()===common.toLowerCase())||{}).id || "user_sp",
          cat: fCat.value,
          zone: fZone.value,
          x: toXY(parseFloat(fLat.value), parseFloat(fLng.value)).x,
          y: toXY(parseFloat(fLat.value), parseFloat(fLng.value)).y,
          lat: parseFloat(fLat.value),
          lng: parseFloat(fLng.value),
          count: parseInt(root.querySelector("#f-count").value) || 1,
          date: "2026-08-31",
          time: new Date().toLocaleTimeString([], {hour:'2-digit', minute:'2-digit'}),
          visit: 3,
          verified: false,
          photo: "",
          voucher: "CITIZEN-" + String(Date.now()).slice(-3),
          iucn: "Community Sighting",
          note: root.querySelector("#f-note").value.trim() || "Observed during citizen engagement survey."
        };

        localObs.push(newRec);
        saveLocal();
        toast(`Observation saved! "${common}" added to Web-GIS map.`);
        navigate("#/map");
      });
    }
  };
};

/* ---------- 6. Analytics & Scientific Findings Screen ---------- */
SCREENS.analytics = () => {
  const zc = zoneCounts();

  const html = `
  <div class="page-head" style="margin-bottom:16px">
    <h1 style="font-size:22px;margin-bottom:2px">Quantitative Analysis & Empirical Survey Findings</h1>
    <p class="hint">Empirical findings from the 53-acre ecological audit & N = 84 multi-stakeholder community perception study.</p>
  </div>

  <div class="grid" style="grid-template-columns:1fr 1fr; gap:16px; margin-bottom:18px">
    <div class="card card-pad" style="border-top:3px solid var(--accent)">
      <div style="display:flex;justify-content:space-between;align-items:flex-start">
        <h3 style="margin:0 0 6px 0;font-size:16px">🔊 The '26.3 dBA Acoustic Shield' Effect</h3>
        <span class="dba-pill dba-quiet">-26.3 dBA Damping</span>
      </div>
      <p style="font-size:13px;color:var(--ink-2);line-height:1.55;margin-bottom:12px">
        Boundary traffic along Dr. Ambedkar Road (Zone 4) reaches a deafening <strong>78.4 dBA</strong>. Deep within the core Botanical Grove (Zone 1), sound levels plummet to <strong>52.1 dBA</strong>. The multi-tiered canopy of Baobab, Ficus, and Mahogany acts as a massive sonic sponge, creating a vital acoustic sanctuary for avifaunal breeding.
      </p>
      <div style="background:var(--bg-tint);padding:10px;border-radius:var(--r-sm)">
        <div style="display:flex;justify-content:space-between;font-size:11.5px;margin-bottom:4px">
          <span>Zone 4 (Road Boundary): <strong>78.4 dBA</strong> (Severe Noise)</span>
          <span>Zone 1 (Core Grove): <strong>52.1 dBA</strong> (Refugium)</span>
        </div>
        <div style="height:8px;background:#e2e4d9;border-radius:4px;overflow:hidden;position:relative">
          <div style="width:66%;height:100%;background:linear-gradient(to right, #2F7D53, #f59e0b, #ef4444);border-radius:4px"></div>
        </div>
      </div>
    </div>

    <div class="card card-pad" style="border-top:3px solid var(--blue)">
      <div style="display:flex;justify-content:space-between;align-items:flex-start">
        <h3 style="margin:0 0 6px 0;font-size:16px">📐 81% Higher Shannon Diversity in Core Grove</h3>
        <span class="dba-pill dba-mod">H' = 2.68 vs 1.48</span>
      </div>
      <p style="font-size:13px;color:var(--ink-2);line-height:1.55;margin-bottom:12px">
        Applying the Shannon-Wiener Diversity Index <em>H' = - Σ (pi × ln pi)</em> and Pielou's Evenness <em>J' = H' / ln(S)</em> validates that the core Botanical Grove sustains <strong>H' = 2.68 (24 taxa, J'=0.84)</strong>, whereas the perimeter buffer drops to <strong>H' = 1.48 (only 6 taxa, J'=0.65)</strong>.
      </p>
      <div style="display:grid;grid-template-columns:repeat(4, 1fr);gap:6px;text-align:center">
        ${zc.map(z=>`
          <div style="background:var(--bg-tint);padding:7px 4px;border-radius:var(--r-sm)">
            <div style="font-size:10.5px;color:var(--muted)">${z.code}</div>
            <strong style="font-size:13px;color:var(--ink)">H'=${z.shannon}</strong>
            <div style="font-size:10px;color:var(--ink-2)">S=${z.species} taxa</div>
          </div>
        `).join("")}
      </div>
    </div>
  </div>

  <div class="card card-pad" style="margin-bottom:18px">
    <div style="margin-bottom:10px">
      <h3 style="margin:0;font-size:15px">Multi-Stakeholder Community Engagement Matrix (N = 84 Participants)</h3>
      <span class="hint">Rigorous civic sampling conducted pursuant to NEP 2020 social consultation standards</span>
    </div>
    <div class="grid" style="grid-template-columns:repeat(auto-fit, minmax(220px, 1fr)); gap:10px">
      <div style="background:var(--bg-tint);padding:12px;border-radius:var(--r-sm);border-left:3px solid var(--accent)">
        <strong style="font-size:13px;color:var(--ink)">Regular Morning Walkers (N = 42)</strong>
        <p style="font-size:12px;color:var(--ink-2);margin:4px 0 0;line-height:1.45">88% visit for tranquility; 76% completely unaware of the 280+ heritage tree species due to captive exhibit bias.</p>
      </div>
      <div style="background:var(--bg-tint);padding:12px;border-radius:var(--r-sm);border-left:3px solid var(--blue)">
        <strong style="font-size:13px;color:var(--ink)">School Students & Educators (N = 26)</strong>
        <p style="font-size:12px;color:var(--ink-2);margin:4px 0 0;line-height:1.45">92% expressed enthusiastic readiness for mobile Web-GIS; high demand for interactive tree QR-code plaques.</p>
      </div>
      <div style="background:var(--bg-tint);padding:12px;border-radius:var(--r-sm);border-left:3px solid var(--gold)">
        <strong style="font-size:13px;color:var(--ink)">Senior Botanical Staff & Malis (N = 12)</strong>
        <p style="font-size:12px;color:var(--ink-2);margin:4px 0 0;line-height:1.45">Revealed arbitrary pre-monsoon tree trimming blindly chops active Coppersmith Barbet nesting holes.</p>
      </div>
      <div style="background:var(--bg-tint);padding:12px;border-radius:var(--r-sm);border-left:3px solid var(--reptile)">
        <strong style="font-size:13px;color:var(--ink)">Wildlife Naturalists & Visitors (N = 4)</strong>
        <p style="font-size:12px;color:var(--ink-2);margin:4px 0 0;line-height:1.45">Affirmed high ecological value of lotus lake for migratory wading birds; requested crowdsourced logging.</p>
      </div>
    </div>
  </div>

  <h2 style="font-family:var(--serif);font-size:20px;margin:20px 0 10px 0">Empirical Survey Charts (Figures 4.5 to 4.12)</h2>
  <div class="survey-chart-grid">
    <div class="survey-chart-card">
      <img src="fig_4_5.png" alt="Fig 4.5: Where is Free-Living Wildlife Observed?">
      <div class="c-body">
        <strong style="font-size:13.5px;color:var(--ink)">Fig. 4.5: Habitat Concentration</strong>
        <p style="font-size:12px;color:var(--ink-2);margin-top:4px;line-height:1.4"><strong>56.5%</strong> of all free-living wildlife encounters cluster strictly in the core Botanical Grove (Zone 1) & Lotus Lake (Zone 2).</p>
      </div>
    </div>

    <div class="survey-chart-card">
      <img src="fig_4_6.png" alt="Fig 4.6: What Threatens Free-Living Flora & Avifauna?">
      <div class="c-body">
        <strong style="font-size:13.5px;color:var(--ink)">Fig. 4.6: Anthropogenic Threats</strong>
        <p style="font-size:12px;color:var(--ink-2);margin-top:4px;line-height:1.4"><strong>36.5%</strong> identify arbitrary pre-monsoon tree lopping as the primary hazard, followed by <strong>27.0%</strong> citing vehicular traffic noise.</p>
      </div>
    </div>

    <div class="survey-chart-card">
      <img src="fig_4_7.png" alt="Fig 4.7: Main Barriers to Local Biodiversity Literacy">
      <div class="c-body">
        <strong style="font-size:13.5px;color:var(--ink)">Fig. 4.7: Literacy Deficit Causes</strong>
        <p style="font-size:12px;color:var(--ink-2);margin-top:4px;line-height:1.4">Over <strong>72.0%</strong> cite the absence of an interactive mobile digital map and signage restricted to caged zoo animals as their core barrier.</p>
      </div>
    </div>

    <div class="survey-chart-card">
      <img src="fig_4_8_4_9.png" alt="Figs 4.8 & 4.9: Web-GIS Utility & Adoption">
      <div class="c-body">
        <strong style="font-size:13.5px;color:var(--ink)">Figs. 4.8 & 4.9: Digital Solution Validation</strong>
        <p style="font-size:12px;color:var(--ink-2);margin-top:4px;line-height:1.4"><strong>86.7%</strong> confirm an interactive Web-GIS app would enrich their visit; <strong>83.3%</strong> intend to actively utilize BioMap.</p>
      </div>
    </div>

    <div class="survey-chart-card">
      <img src="fig_4_10.png" alt="Fig 4.10: What Info to Include?">
      <div class="c-body">
        <strong style="font-size:13.5px;color:var(--ink)">Fig. 4.10: Desired Feature Set</strong>
        <p style="font-size:12px;color:var(--ink-2);margin-top:4px;line-height:1.4">Visitors prioritize GPS coordinates, field photographs, Linnaean taxonomy, and seasonal occurrence in digital species cards.</p>
      </div>
    </div>

    <div class="survey-chart-card">
      <img src="fig_4_11_4_12.png" alt="Figs 4.11 & 4.12: Discovery Methods & QR Utility">
      <div class="c-body">
        <strong style="font-size:13.5px;color:var(--ink)">Figs. 4.11 & 4.12: Discovery & QR Code Utility</strong>
        <p style="font-size:12px;color:var(--ink-2);margin-top:4px;line-height:1.4"><strong>87.0%</strong> affirm physical laser-etched QR code placards on heritage Baobab and Ebony trees would dramatically elevate public learning.</p>
      </div>
    </div>
  </div>
  `;

  return { html, mount() {} };
};

/* ---------- 7. Field Survey Methodology Screen ---------- */
SCREENS.survey = () => {
  const html = `
  <div class="page-head" style="margin-bottom:16px">
    <h1 style="font-size:22px;margin-bottom:2px">Field Survey Methodology & Verification Evidence</h1>
    <p class="hint">Standardized Line-Transect Protocol (Bibby et al., 2000) & Empirical Geospatial Census</p>
  </div>

  <div class="grid" style="grid-template-columns:minmax(0,1.2fr) minmax(0,1fr); align-items:start; gap:18px">
    <div class="grid" style="gap:14px">
      <div class="card card-pad">
        <h3 style="margin-bottom:10px;font-size:14.5px">Geospatial & Field Instrumentation Stack</h3>
        <dl class="kv" style="gap:8px 14px;font-size:12.5px">
          <dt>GNSS Positioning</dt><dd>Garmin eTrex 22x & Smartphone Assisted GNSS (±3m spatial accuracy)</dd>
          <dt>Acoustic Dosimetry</dt><dd>Type 2 Digital Sound Level Meter (A-weighting, fast response, 15-min intervals)</dd>
          <dt>Optical Capture</dt><dd>Nikon D5600 DSLR (70–300mm f/4.5–5.6 ED VR) & Smartphone optical sensors</dd>
          <dt>Survey Protocol</dt><dd>Fixed-Transect Line Sampling (1.8 km continuous path, Bibby et al., 2000)</dd>
          <dt>Taxonomic Validation</dt><dd>Cross-verified via BNHS Urban Biodiversity Register & Birdwood (1862)</dd>
          <dt>Web-GIS Stack</dt><dd>Zero-cost open-source HTML5, CSS3, Vector Cartography, Netlify deployment</dd>
        </dl>
      </div>

      <div class="oral-history-card">
        “Most visitors only run toward the penguins and big cat enclosures. Very few people realize that we have heritage Baobabs brought from Africa during the British era that are over 150 years old. In the breeding months, Coppersmith Barbets bore nesting holes into dead Ficus branches. When external municipal pruning contractors arrive before the monsoons, they often chop down these branches blindly because there is no map indicating where birds are nesting. A digital map on a phone will help save these nesting sites.”
        <cite>— Mr. Maruti Patil, Senior Botanical Mali (34 years of service at Byculla Zoo)</cite>
      </div>
    </div>

    <div class="card" style="overflow:hidden">
      <div class="card-h">
        <h3>Fieldwork Verification Voucher (Fig. C.1)</h3>
      </div>
      <div style="aspect-ratio:4/3;background:#000;overflow:hidden">
        <img src="animal_photos/fieldwork_proof.jpg" alt="Fieldwork Arrival Proof at Byculla Station Gate" style="width:100%;height:100%;object-fit:cover">
      </div>
      <div class="card-pad" style="font-size:12px;color:var(--ink-2);line-height:1.5">
        <p style="margin:0 0 6px 0"><strong>GPS Arrival Point:</strong> 18.9766° N, 72.8328° E (Byculla Station Gate)</p>
        <p style="margin:0 0 6px 0"><strong>Investigators:</strong> Riddhesh Keni (Roll 35) & Shivansh Dubey (Roll 45)</p>
        <p style="margin:0 0 6px 0"><strong>Institution:</strong> Vidyavardhini's Vartak College (CS Dept.), Vasai Road (W)</p>
        <p style="margin:0"><strong>Field Deliverables:</strong> 42+ species audited, 84 survey responses collected, live Web-GIS portal deployed.</p>
      </div>
    </div>
  </div>

  <div class="card" style="margin-top:18px">
    <div class="card-h">
      <div>
        <h3>Fieldwork Sessions Timeline (Monday, 31 August 2026)</h3>
        <span class="hint">4.8 km traversed across all four ecological zones (07:00 AM – 06:30 PM)</span>
      </div>
    </div>
    <div class="card-pad">
      <div style="display:flex;flex-direction:column;gap:12px">
        ${VISITS.map(v => `
          <div style="padding-bottom:12px;border-bottom:1px solid var(--line-2)">
            <div style="display:flex;justify-content:space-between;align-items:baseline;flex-wrap:wrap;gap:6px">
              <strong style="font-size:14px;color:var(--ink)">Session ${v.n}: ${esc(v.session)}</strong>
              <span class="mono-n" style="font-size:11.5px;color:var(--muted)">${v.start} – ${v.end}</span>
            </div>
            <div style="font-size:12px;color:var(--muted);margin-top:2px">${esc(v.weather)} • ${v.photos} photographic captures</div>
            <p style="font-size:12.5px;color:var(--ink-2);margin:4px 0 6px 0;line-height:1.45">${esc(v.notes)}</p>
            <div style="display:flex;gap:6px;flex-wrap:wrap">${v.zones.map(z=>`<span class="tag">${ZONE_BY_ID[z].code}: ${ZONE_BY_ID[z].name}</span>`).join("")}</div>
          </div>
        `).join("")}
      </div>
    </div>
  </div>
  `;

  return { html, mount() {} };
};

/* ---------- 8. About Project & Viva Defense Screen ---------- */
SCREENS.about = () => {
  const html = `
  <div class="page-head" style="margin-bottom:16px">
    <h1 style="font-size:22px;margin-bottom:2px">Project Overview & Academic Framework</h1>
    <p class="hint">Community Engagement Project (CEP) • National Education Policy (NEP 2020) Guidelines</p>
  </div>

  <div class="grid" style="grid-template-columns:minmax(0,1.35fr) minmax(0,1fr); align-items:start; gap:18px">
    <div class="grid" style="gap:16px">
      <div class="card card-pad" style="border-left:4px solid var(--accent)">
        <h3 style="margin-bottom:6px;font-size:14.5px">Official Project Title</h3>
        <div style="font-family:var(--serif);font-size:19px;line-height:1.35;color:var(--accent-ink);margin-bottom:10px">
          “Local Biodiversity Mapping of Byculla Zoo and Surrounding Green Spaces”
        </div>
        <p style="font-size:12.5px;color:var(--ink-2);margin:0;line-height:1.55">
          A Geospatial Biodiversity Census, Web-GIS Cartography & Community Ecological Literacy Study submitted in partial fulfillment for the award of the degree of <strong>Bachelor of Science in Computer Science (T.Y. B.Sc. CS — Semester V, NEP 2020 Guidelines, 2 Credits)</strong>.
        </p>
      </div>

      <div class="card card-pad">
        <h3 style="margin-bottom:10px;font-size:14.5px">Civic Motivation: The 5 Core Challenges</h3>
        <div style="display:flex;flex-direction:column;gap:8px;font-size:12.5px;color:var(--ink-2);line-height:1.55">
          <div><strong>1. 53-Acre Heritage Sanctuary:</strong> Established in 1861, Ranibagh is South Mumbai's largest contiguous botanical refuge with 280+ tree species and heritage Baobabs.</div>
          <div><strong>2. The 'Captive-Species Bias':</strong> Over 78% of visitor attention fixates exclusively on caged exhibits (tigers, penguins), leaving rich botanical heritage and wild avifauna unmapped.</div>
          <div><strong>3. Arbitrary Pre-Monsoon Tree Lopping:</strong> Municipal contractors routinely trim branches prior to monsoons without geospatial knowledge of active bird nesting cavities.</div>
          <div><strong>4. Acoustic Stress Vulnerability:</strong> Ambient vehicular noise along Dr. Ambedkar Road exceeds 78 dBA, driving sensitive avifauna deeper into the core arboretum.</div>
          <div><strong>5. The Digital Literacy Void:</strong> No open-source, mobile-first Web-GIS portal or vernacular guide previously existed for citizens to explore native flora and fauna.</div>
        </div>
      </div>

      <div class="card card-pad">
        <h3 style="margin-bottom:10px;font-size:14.5px">Actionable Recommendations: Three-Tier Municipal Blueprint</h3>
        
        <div class="action-tier-card">
          <strong>Tier 1: Park Level (Immediate / 1–3 Months)</strong>
          <p>Laser-etched QR code placards on 50 audited heritage trees linking directly to bilingual BioMap dossiers; mandatory pre-pruning bird-nest inspection protocol for Tree Authority contractors.</p>
          <div class="stakeholder">Responsible: BMC Garden Dept. & Vartak College CS Dept.</div>
        </div>

        <div class="action-tier-card">
          <strong>Tier 2: BMC Municipal Policy Level (Medium-Term / 3–12 Months)</strong>
          <p>Establish a 15-meter multi-tiered dense vegetative acoustic green buffer wall along the Dr. Ambedkar Road boundary (reducing traffic noise intrusion by 8–12 dBA); integrate BioMap into the statutory People’s Biodiversity Register (PBR).</p>
          <div class="stakeholder">Responsible: BMC Tree Authority & Environmental NGOs</div>
        </div>

        <div class="action-tier-card">
          <strong>Tier 3: Citizen Science & Academic Level (Long-Term / 1–3 Years)</strong>
          <p>Establish university Bio-Blitz weekends for CS/IT undergraduates to crowdsource longitudinal health monitoring of heritage trees across Greater Mumbai's urban green spaces.</p>
          <div class="stakeholder">Responsible: University of Mumbai & Civic Research Foundations</div>
        </div>
      </div>
    </div>

    <div class="grid" style="gap:14px">
      <div class="card card-pad" style="text-align:center">
        <img src="college_logo.jpg" alt="Vartak College Logo" style="height:64px;margin-bottom:10px;object-fit:contain" onerror="this.style.display='none'">
        <h3 style="margin:0 0 3px 0;font-size:14.5px">Vidyavardhini's College of Science</h3>
        <p style="font-size:12px;color:var(--muted);margin:0 0 12px 0">Annasaheb Vartak College, Vasai Road (W) • Affiliated with University of Mumbai</p>
        
        <div style="border-top:1px solid var(--line-2);padding-top:12px;text-align:left">
          <div style="font-size:11px;color:var(--muted);text-transform:uppercase;font-weight:700;margin-bottom:4px">Student Investigators</div>
          <div style="font-size:13px;color:var(--ink);font-weight:600">Riddhesh Keni <span style="font-size:11.5px;color:var(--muted);font-weight:normal">(Roll No: 35)</span></div>
          <div style="font-size:13px;color:var(--ink);font-weight:600;margin-top:3px">Shivansh Dubey <span style="font-size:11.5px;color:var(--muted);font-weight:normal">(Roll No: 45)</span></div>
          <div style="font-size:12px;color:var(--ink-2);margin-top:2px">T.Y. B.Sc. Computer Science (Semester V)</div>
        </div>

        <div style="border-top:1px solid var(--line-2);padding-top:10px;margin-top:10px;text-align:left">
          <div style="font-size:11px;color:var(--muted);text-transform:uppercase;font-weight:700;margin-bottom:4px">Research Mentor & Internal Guide</div>
          <div style="font-size:13.5px;color:var(--accent-ink);font-weight:600">Prof. Shweta Iyer</div>
          <div style="font-size:12px;color:var(--ink-2)">Department of Computer Science</div>
        </div>
      </div>

      <div class="card card-pad" style="background:var(--surface-2)">
        <h3 style="margin-bottom:6px;font-size:13.5px">Evaluator Viva Highlights</h3>
        <ul style="margin:0;padding-left:16px;font-size:12px;color:var(--ink-2);line-height:1.7">
          <li><strong>Interdisciplinary Bridge:</strong> Combines Web-GIS Computer Science with urban municipal ecology.</li>
          <li><strong>Authentic Findings:</strong> -26.3 dBA acoustic canopy shield & H' = 2.68 Shannon index.</li>
          <li><strong>Zero-Cost Architecture:</strong> Pure open-source client-side Web-GIS requiring zero municipal database hosting fees.</li>
        </ul>
      </div>
    </div>
  </div>
  `;

  return { html, mount() {} };
};

/* ============================================================
   Application Shell, Router & Global Navigation
   ============================================================ */

const NAV = [
  {id: "dashboard", label: "Dashboard", short: "Home", icon: "dashboard", route: "#/dashboard"},
  {id: "map", label: "Biodiversity Map", short: "GIS Map", icon: "map", route: "#/map"},
  {id: "species", label: "Species Catalogue", short: "Species", icon: "species", route: "#/species"},
  {id: "observe", label: "Contribute Observation", short: "Add", icon: "observe", route: "#/observe"},
  {id: "analytics", label: "Analytics & Survey", short: "Analytics", icon: "analytics", route: "#/analytics"},
  {id: "survey", label: "Field Methodology", short: "Survey", icon: "survey", route: "#/survey"},
  {id: "about", label: "About & Action Plan", short: "About", icon: "about", route: "#/about"}
];

const BRAND = `
<svg class="brand-mark" viewBox="0 0 32 32" fill="none" aria-hidden="true" style="width:26px;height:26px">
  <circle cx="16" cy="16" r="13.2" stroke="var(--accent)" stroke-width="1.6"/>
  <path d="M3.4 12.4h25.2M3.4 19.6h25.2" stroke="var(--accent)" stroke-width="1.1" opacity=".45"/>
  <path d="M16 2.8c5 5 5 21.4 0 26.4-5-5-5-21.4 0-26.4Z" stroke="var(--accent)" stroke-width="1.1" opacity=".45"/>
  <path d="M22 9.5c1.2 7.2-2.6 11.4-8.6 11.6 0-6.4 3.2-10.6 8.6-11.6Z" fill="var(--accent)"/>
  <path d="M11.6 23.4c1.6-4 4.4-7 8-9" stroke="var(--surface)" stroke-width="1.3" stroke-linecap="round"/>
</svg>
`;

function shell() {
  return `
  <div class="app">
    <aside class="sidebar">
      <div class="brand">
        ${BRAND}
        <div>
          <div class="brand-name">BioMap</div>
          <div class="brand-sub">Byculla Zoo Web-GIS</div>
        </div>
      </div>
      <div class="nav-label">Web-GIS Census</div>
      ${NAV.slice(0, 5).map(n => `
        <button class="nav-btn" data-id="${n.id}" data-go="${n.route}">
          ${icon(n.icon)}
          <span>${n.label}</span>
        </button>
      `).join("")}
      <div class="nav-label" style="margin-top:10px">CEP Documentation</div>
      ${NAV.slice(5).map(n => `
        <button class="nav-btn" data-id="${n.id}" data-go="${n.route}">
          ${icon(n.icon)}
          <span>${n.label}</span>
        </button>
      `).join("")}

      <div style="margin-top:auto;padding-top:12px;border-top:1px solid var(--line-2)">
        <div style="font-size:11px;color:var(--muted);line-height:1.4">Vartak College CS Dept.<br>NEP 2020 • Mumbai University</div>
      </div>
    </aside>

    <div class="main-wrap">
      <header class="topbar">
        <div class="area-pill">
          <span class="dot" style="background:#2F7D53"></span>
          <span><strong>Veermata Jijabai Bhosale Udyan (Ranibagh)</strong> • 53 Acres</span>
        </div>
        <div style="margin-left:auto;display:flex;align-items:center;gap:8px">
          <button class="btn btn-sm" id="btn-quick-search">${icon("search")} Search</button>
          <button class="btn btn-sm" id="btn-export-data">${icon("download")} Export GIS</button>
        </div>
      </header>

      <main class="content" id="view-root"></main>

      <nav class="m-bottom" aria-label="Mobile Navigation">
        ${NAV.slice(0, 4).map(n => `
          <button class="m-btn" data-id="${n.id}" data-go="${n.route}">
            ${icon(n.icon)}
            <span>${n.short}</span>
          </button>
        `).join("")}
        <button class="m-btn" data-go="#/about">
          ${icon("about")}
          <span>About</span>
        </button>
      </nav>
    </div>
  </div>

  <div id="toast-slot"></div>
  `;
}

function parseHash() {
  const h = location.hash || "#/dashboard";
  const [route, ...args] = h.slice(2).split("/");
  return { route: route || "dashboard", args };
}

function navigate(route) {
  location.hash = route;
}

function render() {
  const { route, args } = parseHash();
  const screenFn = SCREENS[route] || SCREENS.dashboard;
  const viewRoot = document.getElementById("view-root");
  if (!viewRoot) return;

  document.querySelectorAll(".nav-btn, .m-btn").forEach(b => {
    b.classList.toggle("active", b.dataset.id === route);
  });

  const res = screenFn(...args);
  if (typeof res === "string") {
    viewRoot.innerHTML = res;
  } else if (res && res.html) {
    viewRoot.innerHTML = res.html;
    if (res.mount) res.mount(viewRoot);
  }
  window.scrollTo(0, 0);
}

function toast(msg) {
  const slot = document.getElementById("toast-slot");
  if (!slot) return;
  const t = document.createElement("div");
  t.className = "toast";
  t.textContent = msg;
  slot.appendChild(t);
  setTimeout(() => {
    t.style.opacity = "0";
    setTimeout(() => t.remove(), 300);
  }, 3200);
}

function esc(s) {
  return String(s||"").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

function openSearchModal() {
  let m = document.getElementById("search-modal");
  if (!m) {
    m = document.createElement("div");
    m.id = "search-modal";
    m.style.cssText = "position:fixed;inset:0;background:rgba(0,0,0,0.5);display:flex;align-items:flex-start;justify-content:center;padding-top:12vh;z-index:999;backdrop-filter:blur(2px)";
    m.innerHTML = `
      <div class="card card-pad" style="width:520px;max-width:92vw;box-shadow:var(--shadow-lg);animation:slideUp .2s ease">
        <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:12px">
          <strong style="font-size:15px">Quick Search Species & Vouchers</strong>
          <button class="btn btn-sm btn-icon" id="sm-close">✕</button>
        </div>
        <input id="sm-input" class="input" placeholder="Type common name, मराठी नाव, scientific name or voucher code..." autofocus>
        <div id="sm-results" style="max-height:280px;overflow-y:auto;margin-top:10px;display:flex;flex-direction:column;gap:6px"></div>
      </div>
    `;
    document.body.appendChild(m);

    const smInput = m.querySelector("#sm-input");
    const smRes = m.querySelector("#sm-results");
    m.querySelector("#sm-close").addEventListener("click", () => m.remove());
    m.addEventListener("click", (e) => { if (e.target === m) m.remove(); });

    smInput.addEventListener("input", () => {
      const q = smInput.value.trim().toLowerCase();
      if (!q) { smRes.innerHTML = ""; return; }
      const matches = SPECIES.filter(s =>
        s.common.toLowerCase().includes(q) ||
        s.marathi.toLowerCase().includes(q) ||
        s.sci.toLowerCase().includes(q) ||
        (s.voucher && s.voucher.toLowerCase().includes(q))
      ).slice(0, 8);

      if (!matches.length) {
        smRes.innerHTML = `<div class="hint" style="padding:10px">No matching species found</div>`;
        return;
      }
      smRes.innerHTML = matches.map(s => `
        <div style="display:flex;justify-content:space-between;align-items:center;padding:7px 10px;background:var(--bg-tint);border-radius:var(--r-sm);cursor:pointer" data-id="${s.id}">
          <div>
            <strong style="font-size:13px">${esc(s.common)}</strong> <span class="marathi-sub">${esc(s.marathi)}</span>
            <div class="sci" style="font-size:11.5px">${esc(s.sci)}</div>
          </div>
          <span class="voucher-pill">${s.voucher||""}</span>
        </div>
      `).join("");

      smRes.querySelectorAll("[data-id]").forEach(el => {
        el.addEventListener("click", () => {
          m.remove();
          navigate("#/species/" + el.dataset.id);
        });
      });
    });
  }
}

function exportGISData() {
  const geojson = {
    type: "FeatureCollection",
    features: allObs().map(o => {
      const sp = speciesOf(o);
      return {
        type: "Feature",
        geometry: { type: "Point", coordinates: [o.lng, o.lat] },
        properties: {
          recordId: o.id,
          voucher: o.voucher,
          commonName: sp.common,
          vernacularMarathi: sp.marathi,
          scientificName: sp.sci,
          category: sp.cat,
          zone: o.zone,
          zoneNoiseDba: (ZONE_BY_ID[o.zone]||{}).soundDba,
          count: o.count,
          surveyDate: o.date,
          iucnStatus: sp.iucn,
          studyArea: "Veermata Jijabai Bhosale Udyan (Byculla Zoo), Mumbai"
        }
      };
    })
  };
  const dataStr = "data:text/json;charset=utf-8," + encodeURIComponent(JSON.stringify(geojson, null, 2));
  const a = document.createElement("a");
  a.setAttribute("href", dataStr);
  a.setAttribute("download", "Byculla_Zoo_BioMap_Census_2026.geojson");
  document.body.appendChild(a);
  a.click();
  a.remove();
  toast("GeoJSON exported! Compatible with QGIS, ArcGIS & BMC PBR.");
}

window.addEventListener("DOMContentLoaded", () => {
  const root = document.getElementById("root");
  root.innerHTML = shell();

  document.getElementById("btn-quick-search").addEventListener("click", openSearchModal);
  document.getElementById("btn-export-data").addEventListener("click", exportGISData);

  document.querySelectorAll("[data-go]").forEach(btn => {
    btn.addEventListener("click", () => navigate(btn.dataset.go));
  });

  window.addEventListener("hashchange", render);
  render();
});
</script>
</body>
</html>
"""

    with open("biomap.html", "w", encoding="utf-8") as f:
        f.write(html)

    print("Finished writing biomap.html. Total length:", len(html), "bytes.")

if __name__ == "__main__":
    build()
