import { useState, useCallback, useEffect, useMemo } from "react";
import { BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer, ReferenceLine, Cell, RadarChart, PolarGrid, PolarAngleAxis, PolarRadiusAxis, Radar } from "recharts";
import * as XLSX from "xlsx";
import { extractPdfText, extractFieldsFromText } from "./pdfParser.js";

// ─── SECTOR DEFINITIONS ───
const SECTORS = [
  { id: "bancos", label: "Bancos", icon: "\u{1F3E6}", type: "acao" },
  { id: "transmissoras", label: "Transmissoras", icon: "\u26A1", type: "acao" },
  { id: "geradoras", label: "Geradoras", icon: "\u{1F4A1}", type: "acao" },
  { id: "seguros", label: "Seguros", icon: "\u{1F6E1}\uFE0F", type: "acao" },
  { id: "saneamento", label: "Saneamento", icon: "\u{1F6B0}", type: "acao" },
  { id: "holdings", label: "Holdings", icon: "\u{1F3E2}", type: "acao" },
  { id: "telecom", label: "Telecomunica\u00E7\u00F5es", icon: "\u{1F4E1}", type: "acao" },
  { id: "fii_logistico", label: "FIIs Log\u00EDsticos", icon: "\u{1F69B}", type: "fii" },
  { id: "fii_shopping", label: "FIIs Shoppings", icon: "\u{1F6CD}\uFE0F", type: "fii" },
  { id: "fii_lajes", label: "FIIs Lajes", icon: "\u{1F3D7}\uFE0F", type: "fii" },
  { id: "fii_papel", label: "FIIs Papel", icon: "\u{1F4C4}", type: "fii" },
];

const DEFAULT_MARGINS = {
  bancos: 20,
  transmissoras: 10,
  geradoras: 25,
  seguros: 25,
  saneamento: 15,
  holdings: 15,
  telecom: 20,
  fii_logistico: 10,
  fii_shopping: 15,
  fii_lajes: 20,
  fii_papel: 15,
};

const MARGIN_RATIONALE = {
  bancos: "Risco de cr\u00E9dito c\u00EDclico, inadimpl\u00EAncia pode saltar",
  transmissoras: "Fluxo previs\u00EDvel (RAP), risco regulat\u00F3rio baixo",
  geradoras: "Risco hidrol\u00F3gico, exposi\u00E7\u00E3o ao spot, GSF",
  seguros: "Provis\u00F5es t\u00E9cnicas complexas, risco de cauda",
  saneamento: "Risco pol\u00EDtico/regulat\u00F3rio, capex de universaliza\u00E7\u00E3o",
  holdings: "Desconto de holding j\u00E1 embutido; margem adicional por governan\u00E7a",
  telecom: "Ciclo tecnol\u00F3gico, capex pesado (5G/fibra)",
  fii_logistico: "Contratos longos, demanda estrutural e-commerce",
  fii_shopping: "Sensibilidade ao ciclo econ\u00F4mico e varejo",
  fii_lajes: "Vac\u00E2ncia estrutural, home office, ciclo imobili\u00E1rio",
  fii_papel: "Risco de cr\u00E9dito, sensibilidade a juros e infla\u00E7\u00E3o",
};

const SECTOR_METHODOLOGY = {
  bancos: { name: "Excess Return / Gordon P/B", formula: "P/B = (ROE \u2212 g) / (Ke \u2212 g)" },
  transmissoras: { name: "DCF Regulat\u00F3rio (FCFE)", formula: "\u03A3 (RAP \u2212 OPEX \u2212 Capex) / (1+Ke)^t" },
  geradoras: { name: "EV/EBITDA + EV/MW", formula: "Fair EV = EBITDA \u00D7 m\u00FAltiplo justo" },
  seguros: { name: "Embedded Value Simplif. + P/B", formula: "P/B = (ROE \u2212 g) / (Ke \u2212 g)" },
  saneamento: { name: "DCF Regulat\u00F3rio (RAB)", formula: "EV/RAB target \u00D7 RAB" },
  holdings: { name: "Sum-of-the-Parts (SOTP)", formula: "NAV = \u03A3 Participa\u00E7\u00F5es \u2212 D\u00EDvida" },
  telecom: { name: "DCF (FCFF) / EV/EBITDA", formula: "Fair EV = EBITDA \u00D7 m\u00FAltiplo justo" },
  fii_logistico: { name: "NAV (Cap Rate)", formula: "NAV = (NOI / Cap Rate) \u2212 D\u00EDvida" },
  fii_shopping: { name: "NAV (Cap Rate) + NOI/ABL", formula: "NAV = (NOI / Cap Rate) \u2212 D\u00EDvida" },
  fii_lajes: { name: "NAV Ajustado por Vac\u00E2ncia", formula: "NAV = (NOI\u00D7(1\u2212vac_norm) / Cap Rate) \u2212 D\u00EDvida" },
  fii_papel: { name: "An\u00E1lise de Cr\u00E9dito / P/VP", formula: "VP justo \u2248 VP \u00D7 (spread_adj / risco)" },
};

// ─── INPUT FIELD DEFINITIONS PER SECTOR ───
const SECTOR_FIELDS = {
  bancos: [
    { key: "ticker", label: "Ticker", type: "text", placeholder: "Ex: ITUB4" },
    { key: "currentPrice", label: "Pre\u00E7o Atual (R$)", type: "number", step: "0.01" },
    { key: "bookValuePerShare", label: "VPA - Valor Patrimonial/A\u00E7\u00E3o (R$)", type: "number", step: "0.01" },
    { key: "roe", label: "ROE Recorrente (%)", type: "number", step: "0.1", hint: "Ex: 20 para 20%" },
    { key: "ke", label: "Custo de Equity - Ke (%)", type: "number", step: "0.1", hint: "Ex: 14 para 14%" },
    { key: "g", label: "Crescimento Sustent\u00E1vel - g (%)", type: "number", step: "0.1", hint: "Ex: 5 para 5%" },
    { key: "earningsPerShare", label: "LPA - Lucro por A\u00E7\u00E3o (R$)", type: "number", step: "0.01" },
  ],
  transmissoras: [
    { key: "ticker", label: "Ticker", type: "text", placeholder: "Ex: TAEE11" },
    { key: "currentPrice", label: "Pre\u00E7o Atual (R$)", type: "number", step: "0.01" },
    { key: "rapAnual", label: "RAP Anual (R$ milh\u00F5es)", type: "number", step: "1" },
    { key: "opex", label: "OPEX Anual (R$ milh\u00F5es)", type: "number", step: "1" },
    { key: "capex", label: "Capex Manuten\u00E7\u00E3o Anual (R$ milh\u00F5es)", type: "number", step: "1" },
    { key: "ke", label: "Custo de Equity - Ke (%)", type: "number", step: "0.1" },
    { key: "g", label: "Crescimento da RAP (%)", type: "number", step: "0.1" },
    { key: "payout", label: "Payout (%)", type: "number", step: "1" },
    { key: "sharesOutstanding", label: "A\u00E7\u00F5es em Circula\u00E7\u00E3o (milh\u00F5es)", type: "number", step: "0.1" },
    { key: "concessionYears", label: "Anos Restantes Concess\u00E3o", type: "number", step: "1" },
  ],
  geradoras: [
    { key: "ticker", label: "Ticker", type: "text", placeholder: "Ex: EGIE3" },
    { key: "currentPrice", label: "Pre\u00E7o Atual (R$)", type: "number", step: "0.01" },
    { key: "ebitda", label: "EBITDA (R$ milh\u00F5es)", type: "number", step: "1" },
    { key: "netDebt", label: "D\u00EDvida L\u00EDquida (R$ milh\u00F5es)", type: "number", step: "1" },
    { key: "mwInstalled", label: "MW Instalado", type: "number", step: "1" },
    { key: "sharesOutstanding", label: "A\u00E7\u00F5es em Circula\u00E7\u00E3o (milh\u00F5es)", type: "number", step: "0.1" },
    { key: "evEbitdaTarget", label: "EV/EBITDA Alvo (m\u00FAltiplo)", type: "number", step: "0.1", hint: "Geradoras: 6-9x" },
  ],
  seguros: [
    { key: "ticker", label: "Ticker", type: "text", placeholder: "Ex: BBSE3" },
    { key: "currentPrice", label: "Pre\u00E7o Atual (R$)", type: "number", step: "0.01" },
    { key: "bookValuePerShare", label: "VPA (R$)", type: "number", step: "0.01" },
    { key: "roe", label: "ROE Recorrente (%)", type: "number", step: "0.1" },
    { key: "ke", label: "Custo de Equity - Ke (%)", type: "number", step: "0.1" },
    { key: "g", label: "Crescimento - g (%)", type: "number", step: "0.1" },
    { key: "combinedRatio", label: "Combined Ratio (%)", type: "number", step: "0.1" },
  ],
  saneamento: [
    { key: "ticker", label: "Ticker", type: "text", placeholder: "Ex: SAPR11" },
    { key: "currentPrice", label: "Pre\u00E7o Atual (R$)", type: "number", step: "0.01" },
    { key: "rab", label: "RAB (R$ milh\u00F5es)", type: "number", step: "1" },
    { key: "ebitda", label: "EBITDA (R$ milh\u00F5es)", type: "number", step: "1" },
    { key: "netDebt", label: "D\u00EDvida L\u00EDquida (R$ milh\u00F5es)", type: "number", step: "1" },
    { key: "sharesOutstanding", label: "A\u00E7\u00F5es em Circula\u00E7\u00E3o (milh\u00F5es)", type: "number", step: "0.1" },
    { key: "evRabTarget", label: "EV/RAB Alvo", type: "number", step: "0.1", hint: "Refer\u00EAncia: 1.0-1.5x" },
  ],
  holdings: [
    { key: "ticker", label: "Ticker", type: "text", placeholder: "Ex: ITSA4" },
    { key: "currentPrice", label: "Pre\u00E7o Atual (R$)", type: "number", step: "0.01" },
    { key: "totalInvestments", label: "Valor Total Participa\u00E7\u00F5es (R$ milh\u00F5es)", type: "number", step: "1" },
    { key: "holdingDebt", label: "D\u00EDvida L\u00EDquida da Holding (R$ milh\u00F5es)", type: "number", step: "1" },
    { key: "sharesOutstanding", label: "A\u00E7\u00F5es em Circula\u00E7\u00E3o (milh\u00F5es)", type: "number", step: "0.1" },
    { key: "holdingDiscount", label: "Desconto de Holding (%)", type: "number", step: "1", hint: "T\u00EDpico Brasil: 15-30%" },
  ],
  telecom: [
    { key: "ticker", label: "Ticker", type: "text", placeholder: "Ex: VIVT3" },
    { key: "currentPrice", label: "Pre\u00E7o Atual (R$)", type: "number", step: "0.01" },
    { key: "ebitda", label: "EBITDA (R$ milh\u00F5es)", type: "number", step: "1" },
    { key: "netDebt", label: "D\u00EDvida L\u00EDquida (R$ milh\u00F5es)", type: "number", step: "1" },
    { key: "sharesOutstanding", label: "A\u00E7\u00F5es em Circula\u00E7\u00E3o (milh\u00F5es)", type: "number", step: "0.1" },
    { key: "evEbitdaTarget", label: "EV/EBITDA Alvo (m\u00FAltiplo)", type: "number", step: "0.1", hint: "Telecom: 5-7x" },
  ],
  fii_logistico: [
    { key: "ticker", label: "Ticker", type: "text", placeholder: "Ex: HGLG11" },
    { key: "currentPrice", label: "Pre\u00E7o Atual (R$)", type: "number", step: "0.01" },
    { key: "noiAnual", label: "NOI Anual (R$ milh\u00F5es)", type: "number", step: "0.1" },
    { key: "capRate", label: "Cap Rate de Mercado (%)", type: "number", step: "0.1", hint: "Log\u00EDstico: 7-9%" },
    { key: "totalDebt", label: "D\u00EDvida Total (R$ milh\u00F5es)", type: "number", step: "0.1" },
    { key: "totalShares", label: "Total de Cotas (milh\u00F5es)", type: "number", step: "0.1" },
    { key: "vpPerShare", label: "VP por Cota (R$)", type: "number", step: "0.01" },
  ],
  fii_shopping: [
    { key: "ticker", label: "Ticker", type: "text", placeholder: "Ex: VISC11" },
    { key: "currentPrice", label: "Pre\u00E7o Atual (R$)", type: "number", step: "0.01" },
    { key: "noiAnual", label: "NOI Anual (R$ milh\u00F5es)", type: "number", step: "0.1" },
    { key: "capRate", label: "Cap Rate de Mercado (%)", type: "number", step: "0.1", hint: "Shoppings: 7-10%" },
    { key: "totalDebt", label: "D\u00EDvida Total (R$ milh\u00F5es)", type: "number", step: "0.1" },
    { key: "totalShares", label: "Total de Cotas (milh\u00F5es)", type: "number", step: "0.1" },
    { key: "vpPerShare", label: "VP por Cota (R$)", type: "number", step: "0.01" },
    { key: "vendasM2", label: "Vendas/m\u00B2 (R$/m\u00EAs)", type: "number", step: "1" },
    { key: "ablTotal", label: "ABL Total (m\u00B2)", type: "number", step: "1" },
  ],
  fii_lajes: [
    { key: "ticker", label: "Ticker", type: "text", placeholder: "Ex: BRCR11" },
    { key: "currentPrice", label: "Pre\u00E7o Atual (R$)", type: "number", step: "0.01" },
    { key: "noiAnual", label: "NOI Anual (R$ milh\u00F5es)", type: "number", step: "0.1" },
    { key: "capRate", label: "Cap Rate de Mercado (%)", type: "number", step: "0.1", hint: "Lajes: 8-11%" },
    { key: "vacancyActual", label: "Vac\u00E2ncia Atual (%)", type: "number", step: "0.1" },
    { key: "vacancyNormalized", label: "Vac\u00E2ncia Normalizada (%)", type: "number", step: "0.1", hint: "M\u00E9dia hist\u00F3rica da regi\u00E3o" },
    { key: "totalDebt", label: "D\u00EDvida Total (R$ milh\u00F5es)", type: "number", step: "0.1" },
    { key: "totalShares", label: "Total de Cotas (milh\u00F5es)", type: "number", step: "0.1" },
    { key: "vpPerShare", label: "VP por Cota (R$)", type: "number", step: "0.01" },
  ],
  fii_papel: [
    { key: "ticker", label: "Ticker", type: "text", placeholder: "Ex: KNCR11" },
    { key: "currentPrice", label: "Pre\u00E7o Atual (R$)", type: "number", step: "0.01" },
    { key: "vpPerShare", label: "VP por Cota (R$)", type: "number", step: "0.01" },
    { key: "spreadMedio", label: "Spread M\u00E9dio Carteira (%)", type: "number", step: "0.1" },
    { key: "duration", label: "Duration M\u00E9dia (anos)", type: "number", step: "0.1" },
    { key: "ltvMedio", label: "LTV M\u00E9dio (%)", type: "number", step: "0.1" },
    { key: "inadimplencia", label: "Inadimpl\u00EAncia (%)", type: "number", step: "0.1" },
    { key: "pctIPCA", label: "% Carteira IPCA+", type: "number", step: "1" },
    { key: "pctCDI", label: "% Carteira CDI+", type: "number", step: "1" },
    { key: "dividendYield", label: "Dividend Yield 12m (%)", type: "number", step: "0.1" },
    { key: "ntnbRate", label: "NTN-B Longa (%)", type: "number", step: "0.1" },
  ],
};

// ─── VALUATION ENGINES ───
function calculateValuation(sector, inputs) {
  const p = {};
  Object.keys(inputs).forEach((k) => {
    p[k] = inputs[k] === "" || inputs[k] === undefined ? 0 : parseFloat(inputs[k]) || 0;
  });

  let fairPrice = 0;
  let details = {};

  switch (sector) {
    case "bancos": {
      const roe = p.roe / 100;
      const ke = p.ke / 100;
      const g = p.g / 100;
      if (ke > g && ke > 0) {
        const fairPB = (roe - g) / (ke - g);
        fairPrice = fairPB * p.bookValuePerShare;
        const currentPB = p.currentPrice / (p.bookValuePerShare || 1);
        details = {
          "P/B Justo": fairPB.toFixed(2) + "x",
          "P/B Atual": currentPB.toFixed(2) + "x",
          "ROE": (roe * 100).toFixed(1) + "%",
          "Ke": (ke * 100).toFixed(1) + "%",
          "g": (g * 100).toFixed(1) + "%",
          "ROE > Ke": roe > ke ? "\u2705 Cria valor" : "\u26A0\uFE0F Destr\u00F3i valor",
          "P/E impl\u00EDcito": p.earningsPerShare > 0 ? (fairPrice / p.earningsPerShare).toFixed(1) + "x" : "N/A",
        };
      }
      break;
    }
    case "transmissoras": {
      const ke = p.ke / 100;
      const g = p.g / 100;
      const fcfAnual = p.rapAnual - p.opex - p.capex;
      if (ke > g && p.sharesOutstanding > 0 && p.concessionYears > 0) {
        let totalPV = 0;
        for (let t = 1; t <= p.concessionYears; t++) {
          const fcf_t = fcfAnual * Math.pow(1 + g, t);
          totalPV += fcf_t / Math.pow(1 + ke, t);
        }
        fairPrice = totalPV / p.sharesOutstanding;
        const divPerShare = (fcfAnual * (p.payout / 100)) / p.sharesOutstanding;
        details = {
          "FCF Anual (R$ mi)": fcfAnual.toFixed(0),
          "VP Total (R$ mi)": totalPV.toFixed(0),
          "DY Impl\u00EDcito": p.currentPrice > 0 ? ((divPerShare / p.currentPrice) * 100).toFixed(1) + "%" : "N/A",
          "Anos Restantes": p.concessionYears,
          "Payout": p.payout + "%",
        };
      }
      break;
    }
    case "geradoras": {
      if (p.evEbitdaTarget > 0 && p.sharesOutstanding > 0) {
        const fairEV = p.ebitda * p.evEbitdaTarget;
        const fairEquity = fairEV - p.netDebt;
        fairPrice = fairEquity / p.sharesOutstanding;
        const evMw = fairEV / (p.mwInstalled || 1);
        const currentEV = p.currentPrice * p.sharesOutstanding + p.netDebt;
        details = {
          "EV Justo (R$ mi)": fairEV.toFixed(0),
          "EV Atual (R$ mi)": currentEV.toFixed(0),
          "EV/EBITDA Alvo": p.evEbitdaTarget.toFixed(1) + "x",
          "EV/EBITDA Atual": (currentEV / (p.ebitda || 1)).toFixed(1) + "x",
          "EV/MW Justo (R$ mi)": (evMw).toFixed(2),
          "MW Instalado": p.mwInstalled,
        };
      }
      break;
    }
    case "seguros": {
      const roe = p.roe / 100;
      const ke = p.ke / 100;
      const g = p.g / 100;
      if (ke > g && ke > 0) {
        const fairPB = (roe - g) / (ke - g);
        const crAdj = p.combinedRatio < 100 ? 1 + (100 - p.combinedRatio) / 200 : 1 - (p.combinedRatio - 100) / 200;
        fairPrice = fairPB * p.bookValuePerShare * Math.max(0.5, Math.min(1.5, crAdj));
        details = {
          "P/B Justo (base)": fairPB.toFixed(2) + "x",
          "Ajuste Combined Ratio": (crAdj * 100).toFixed(0) + "%",
          "P/B Justo Ajustado": (fairPrice / (p.bookValuePerShare || 1)).toFixed(2) + "x",
          "Combined Ratio": p.combinedRatio + "%",
          "Resultado T\u00E9cnico": p.combinedRatio < 100 ? "\u2705 Lucro t\u00E9cnico" : "\u26A0\uFE0F Preju\u00EDzo t\u00E9cnico",
        };
      }
      break;
    }
    case "saneamento": {
      if (p.evRabTarget > 0 && p.sharesOutstanding > 0) {
        const fairEV = p.rab * p.evRabTarget;
        const fairEquity = fairEV - p.netDebt;
        fairPrice = fairEquity / p.sharesOutstanding;
        const currentEV = p.currentPrice * p.sharesOutstanding + p.netDebt;
        details = {
          "EV Justo (R$ mi)": fairEV.toFixed(0),
          "EV/RAB Alvo": p.evRabTarget.toFixed(2) + "x",
          "EV/RAB Atual": (currentEV / (p.rab || 1)).toFixed(2) + "x",
          "EV/EBITDA Atual": (currentEV / (p.ebitda || 1)).toFixed(1) + "x",
          "RAB (R$ mi)": p.rab.toFixed(0),
        };
      }
      break;
    }
    case "holdings": {
      if (p.sharesOutstanding > 0) {
        const nav = p.totalInvestments - p.holdingDebt;
        const discount = p.holdingDiscount / 100;
        const fairEquity = nav * (1 - discount);
        fairPrice = fairEquity / p.sharesOutstanding;
        const navPerShare = nav / p.sharesOutstanding;
        details = {
          "NAV Total (R$ mi)": nav.toFixed(0),
          "NAV/A\u00E7\u00E3o": navPerShare.toFixed(2),
          "Desconto Aplicado": p.holdingDiscount + "%",
          "P/NAV Atual": (p.currentPrice / (navPerShare || 1)).toFixed(2) + "x",
          "Desconto Atual vs NAV": ((1 - p.currentPrice / (navPerShare || 1)) * 100).toFixed(1) + "%",
        };
      }
      break;
    }
    case "telecom": {
      if (p.evEbitdaTarget > 0 && p.sharesOutstanding > 0) {
        const fairEV = p.ebitda * p.evEbitdaTarget;
        const fairEquity = fairEV - p.netDebt;
        fairPrice = fairEquity / p.sharesOutstanding;
        const currentEV = p.currentPrice * p.sharesOutstanding + p.netDebt;
        details = {
          "EV Justo (R$ mi)": fairEV.toFixed(0),
          "EV/EBITDA Alvo": p.evEbitdaTarget.toFixed(1) + "x",
          "EV/EBITDA Atual": (currentEV / (p.ebitda || 1)).toFixed(1) + "x",
        };
      }
      break;
    }
    case "fii_logistico":
    case "fii_shopping": {
      const capRate = p.capRate / 100;
      if (capRate > 0 && p.totalShares > 0) {
        const propertyValue = p.noiAnual / capRate;
        const nav = propertyValue - p.totalDebt;
        fairPrice = nav / p.totalShares;
        details = {
          "Valor Im\u00F3veis (R$ mi)": propertyValue.toFixed(1),
          "NAV (R$ mi)": nav.toFixed(1),
          "NAV/Cota": fairPrice.toFixed(2),
          "P/VP Atual": (p.currentPrice / (p.vpPerShare || 1)).toFixed(2) + "x",
          "Cap Rate": p.capRate + "%",
          "FFO Yield Impl\u00EDcito": p.currentPrice > 0 ? ((p.noiAnual / p.totalShares / p.currentPrice) * 100).toFixed(1) + "%" : "N/A",
        };
        if (sector === "fii_shopping" && p.vendasM2 > 0) {
          details["Vendas/m\u00B2"] = "R$ " + p.vendasM2.toFixed(0);
          details["NOI/ABL (R$/m\u00B2/m\u00EAs)"] = p.ablTotal > 0 ? "R$ " + ((p.noiAnual * 1e6) / p.ablTotal / 12).toFixed(0) : "N/A";
        }
      }
      break;
    }
    case "fii_lajes": {
      const capRate = p.capRate / 100;
      const vacNorm = p.vacancyNormalized / 100;
      if (capRate > 0 && p.totalShares > 0) {
        const noiAdjusted = p.noiAnual * (1 - vacNorm);
        const propertyValue = noiAdjusted / capRate;
        const nav = propertyValue - p.totalDebt;
        fairPrice = nav / p.totalShares;
        details = {
          "NOI Ajustado (R$ mi)": noiAdjusted.toFixed(1),
          "Valor Im\u00F3veis Ajust. (R$ mi)": propertyValue.toFixed(1),
          "NAV/Cota": fairPrice.toFixed(2),
          "P/VP Atual": (p.currentPrice / (p.vpPerShare || 1)).toFixed(2) + "x",
          "Vac\u00E2ncia Atual": p.vacancyActual + "%",
          "Vac\u00E2ncia Normalizada": p.vacancyNormalized + "%",
          "Cap Rate": p.capRate + "%",
        };
      }
      break;
    }
    case "fii_papel": {
      if (p.vpPerShare > 0) {
        const spreadQuality = p.spreadMedio / (p.ntnbRate || 10);
        const riskPenalty = (p.inadimplencia / 100) * 2 + (p.ltvMedio > 70 ? 0.05 : 0);
        const fairPVP = Math.max(0.7, Math.min(1.3, 0.9 + spreadQuality * 0.15 - riskPenalty));
        fairPrice = p.vpPerShare * fairPVP;
        details = {
          "P/VP Justo": fairPVP.toFixed(2) + "x",
          "P/VP Atual": (p.currentPrice / p.vpPerShare).toFixed(2) + "x",
          "Spread M\u00E9dio": p.spreadMedio + "%",
          "Duration": p.duration + " anos",
          "LTV M\u00E9dio": p.ltvMedio + "%",
          "Inadimpl\u00EAncia": p.inadimplencia + "%",
          "Mix IPCA/CDI": p.pctIPCA + "% / " + p.pctCDI + "%",
          "DY 12m": p.dividendYield + "%",
          "Yield vs NTN-B": ((p.dividendYield - p.ntnbRate) > 0 ? "+" : "") + (p.dividendYield - p.ntnbRate).toFixed(1) + "pp",
        };
      }
      break;
    }
  }

  return { fairPrice: Math.max(0, fairPrice), details };
}

// ─── EXCEL PARSER ───
function parseExcelFile(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: "array" });
        const sheets = {};
        workbook.SheetNames.forEach((name) => {
          sheets[name] = XLSX.utils.sheet_to_json(workbook.Sheets[name], { header: 1 });
        });
        resolve(sheets);
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}

// ─── STYLES ───
const colors = {
  bg: "#0a0f1a",
  surface: "#111827",
  surfaceAlt: "#1a2236",
  border: "#1e2a3f",
  borderLight: "#2a3a52",
  accent: "#00d4aa",
  accentDim: "#00d4aa33",
  warning: "#f59e0b",
  danger: "#ef4444",
  dangerDim: "#ef444433",
  text: "#e2e8f0",
  textMuted: "#8896ab",
  textDim: "#4a5568",
  green: "#10b981",
  greenDim: "#10b98133",
  red: "#ef4444",
  blue: "#3b82f6",
  purple: "#8b5cf6",
  chart1: "#00d4aa",
  chart2: "#3b82f6",
  chart3: "#f59e0b",
};

// ─── MAIN COMPONENT ───
export default function ValuationApp() {
  const [selectedSector, setSelectedSector] = useState(null);
  const [inputs, setInputs] = useState({});
  const [margins, setMargins] = useState({ ...DEFAULT_MARGINS });
  const [showMarginPanel, setShowMarginPanel] = useState(false);
  const [result, setResult] = useState(null);
  const [uploadedData, setUploadedData] = useState(null);
  const [uploadFileName, setUploadFileName] = useState("");
  const [pdfData, setPdfData] = useState(null);
  const [pdfLoading, setPdfLoading] = useState(false);
  const [history, setHistory] = useState([]);
  const [activeTab, setActiveTab] = useState("calc");
  const [showMethodology, setShowMethodology] = useState(false);

  const sectorInfo = SECTORS.find((s) => s.id === selectedSector);
  const methodology = selectedSector ? SECTOR_METHODOLOGY[selectedSector] : null;

  const handleSectorSelect = (sectorId) => {
    setSelectedSector(sectorId);
    setInputs({});
    setResult(null);
    setUploadedData(null);
    setUploadFileName("");
    setPdfData(null);
  };

  const handleInputChange = (key, value) => {
    setInputs((prev) => ({ ...prev, [key]: value }));
  };

  const handleFileUpload = async (e) => {
    const file = e.target.files[0];
    if (!file) return;
    setUploadFileName(file.name);
    setPdfData(null);
    try {
      if (file.name.endsWith(".xlsx") || file.name.endsWith(".xls") || file.name.endsWith(".csv")) {
        const sheets = await parseExcelFile(file);
        setUploadedData(sheets);
      } else if (file.name.toLowerCase().endsWith(".pdf")) {
        setPdfLoading(true);
        setUploadedData(null);
        const arrayBuffer = await file.arrayBuffer();
        const { pages, fullText, numPages } = await extractPdfText(arrayBuffer);
        const extracted = selectedSector
          ? extractFieldsFromText(fullText, selectedSector)
          : { extractedFields: {}, confidence: {} };
        setPdfData({ pages, fullText, numPages, ...extracted });
        setUploadedData({ info: `PDF processado: ${numPages} p\u00E1gina(s). ${Object.keys(extracted.extractedFields).length} campo(s) detectado(s).` });
        setPdfLoading(false);
      } else {
        setUploadedData({ info: "Arquivo carregado: " + file.name + ". Para extra\u00E7\u00E3o autom\u00E1tica, use .xlsx, .csv ou .pdf." });
      }
    } catch (err) {
      setPdfLoading(false);
      setUploadedData({ error: "Erro ao processar arquivo: " + err.message });
    }
  };

  const handleApplyPdfFields = () => {
    if (!pdfData?.extractedFields) return;
    setInputs((prev) => {
      const merged = { ...prev };
      for (const [key, value] of Object.entries(pdfData.extractedFields)) {
        if (!merged[key] || merged[key] === "") {
          merged[key] = String(value);
        }
      }
      return merged;
    });
  };

  const handleApplySingleField = (key, value) => {
    setInputs((prev) => ({ ...prev, [key]: String(value) }));
  };

  const handleCalculate = () => {
    if (!selectedSector) return;
    const { fairPrice, details } = calculateValuation(selectedSector, inputs);
    const margin = margins[selectedSector] || 0;
    const fairWithMargin = fairPrice * (1 - margin / 100);
    const currentPrice = parseFloat(inputs.currentPrice) || 0;
    const upside = currentPrice > 0 ? ((fairPrice / currentPrice - 1) * 100) : 0;
    const upsideWithMargin = currentPrice > 0 ? ((fairWithMargin / currentPrice - 1) * 100) : 0;

    const newResult = {
      ticker: inputs.ticker || "N/A",
      sector: selectedSector,
      sectorLabel: sectorInfo?.label || "",
      currentPrice,
      fairPrice,
      fairWithMargin,
      margin,
      upside,
      upsideWithMargin,
      verdict: upsideWithMargin > 0 ? "COMPRA" : upsideWithMargin > -10 ? "NEUTRO" : "CARO",
      details,
      timestamp: new Date().toLocaleString("pt-BR"),
    };
    setResult(newResult);
    setHistory((prev) => [newResult, ...prev.filter((h) => h.ticker !== newResult.ticker)].slice(0, 20));
  };

  const chartData = useMemo(() => {
    if (!result || result.fairPrice === 0) return [];
    return [
      { name: "Pre\u00E7o Atual", value: result.currentPrice, fill: colors.blue },
      { name: "Pre\u00E7o Justo", value: result.fairPrice, fill: colors.accent },
      { name: "Pre\u00E7o c/ Margem", value: result.fairWithMargin, fill: colors.warning },
    ];
  }, [result]);

  const historyChartData = useMemo(() => {
    return history.map((h) => ({
      name: h.ticker,
      "Pre\u00E7o Atual": h.currentPrice,
      "Pre\u00E7o Justo": h.fairPrice,
      "Com Margem": h.fairWithMargin,
    }));
  }, [history]);

  const verdictColor = result?.verdict === "COMPRA" ? colors.green : result?.verdict === "NEUTRO" ? colors.warning : colors.red;
  const verdictBg = result?.verdict === "COMPRA" ? colors.greenDim : result?.verdict === "NEUTRO" ? colors.warning + "22" : colors.dangerDim;

  return (
    <div style={{ minHeight: "100vh", background: `linear-gradient(170deg, ${colors.bg} 0%, #0d1424 50%, #0a1018 100%)`, color: colors.text, fontFamily: "'DM Sans', 'Segoe UI', sans-serif" }}>
      {/* ─── HEADER ─── */}
      <header style={{ borderBottom: `1px solid ${colors.border}`, padding: "16px 24px", display: "flex", alignItems: "center", justifyContent: "space-between", backdropFilter: "blur(20px)", background: colors.bg + "dd", position: "sticky", top: 0, zIndex: 100 }}>
        <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
          <div style={{ width: 36, height: 36, borderRadius: 8, background: `linear-gradient(135deg, ${colors.accent}, ${colors.blue})`, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 18, fontWeight: 700 }}>V</div>
          <div>
            <h1 style={{ fontSize: 18, fontWeight: 700, margin: 0, letterSpacing: "-0.02em" }}>ValuationPro</h1>
            <p style={{ fontSize: 11, color: colors.textMuted, margin: 0 }}>Framework Multi-Setor</p>
          </div>
        </div>
        <div style={{ display: "flex", gap: 8 }}>
          {["calc", "historico"].map((tab) => (
            <button key={tab} onClick={() => setActiveTab(tab)} style={{ padding: "6px 14px", borderRadius: 6, border: `1px solid ${activeTab === tab ? colors.accent : colors.border}`, background: activeTab === tab ? colors.accentDim : "transparent", color: activeTab === tab ? colors.accent : colors.textMuted, fontSize: 13, fontWeight: 500, cursor: "pointer", transition: "all 0.2s" }}>
              {tab === "calc" ? "Calculadora" : "Hist\u00F3rico"}
            </button>
          ))}
          <button onClick={() => setShowMarginPanel(!showMarginPanel)} style={{ padding: "6px 14px", borderRadius: 6, border: `1px solid ${showMarginPanel ? colors.warning : colors.border}`, background: showMarginPanel ? colors.warning + "22" : "transparent", color: showMarginPanel ? colors.warning : colors.textMuted, fontSize: 13, fontWeight: 500, cursor: "pointer" }}>
            Margens
          </button>
        </div>
      </header>

      <div style={{ maxWidth: 1280, margin: "0 auto", padding: "24px 20px" }}>
        {/* ─── MARGIN CONFIG PANEL ─── */}
        {showMarginPanel && (
          <div style={{ background: colors.surface, border: `1px solid ${colors.border}`, borderRadius: 12, padding: 20, marginBottom: 24, animation: "fadeIn 0.3s ease" }}>
            <h3 style={{ fontSize: 15, fontWeight: 600, marginBottom: 4, color: colors.warning }}>Configura\u00E7\u00E3o de Margem de Seguran\u00E7a</h3>
            <p style={{ fontSize: 12, color: colors.textMuted, marginBottom: 16 }}>Ajuste a margem de seguran\u00E7a para cada setor. Valores mais altos = mais conservador.</p>
            <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(280px, 1fr))", gap: 12 }}>
              {SECTORS.map((s) => (
                <div key={s.id} style={{ display: "flex", alignItems: "center", gap: 10, padding: "8px 12px", background: colors.surfaceAlt, borderRadius: 8, border: `1px solid ${colors.border}` }}>
                  <span style={{ fontSize: 16 }}>{s.icon}</span>
                  <div style={{ flex: 1, minWidth: 0 }}>
                    <div style={{ fontSize: 12, fontWeight: 500, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>{s.label}</div>
                    <div style={{ fontSize: 10, color: colors.textMuted, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>{MARGIN_RATIONALE[s.id]}</div>
                  </div>
                  <div style={{ display: "flex", alignItems: "center", gap: 4 }}>
                    <input type="range" min="0" max="50" value={margins[s.id]} onChange={(e) => setMargins((prev) => ({ ...prev, [s.id]: parseInt(e.target.value) }))} style={{ width: 60, accentColor: colors.warning }} />
                    <span style={{ fontSize: 13, fontWeight: 600, color: colors.warning, minWidth: 32, textAlign: "right", fontFamily: "'JetBrains Mono', monospace" }}>{margins[s.id]}%</span>
                  </div>
                </div>
              ))}
            </div>
            <button onClick={() => setMargins({ ...DEFAULT_MARGINS })} style={{ marginTop: 12, padding: "6px 12px", borderRadius: 6, border: `1px solid ${colors.border}`, background: "transparent", color: colors.textMuted, fontSize: 12, cursor: "pointer" }}>
              Restaurar padr\u00F5es
            </button>
          </div>
        )}

        {activeTab === "calc" ? (
          <div style={{ display: "grid", gridTemplateColumns: result ? "1fr 1fr" : "1fr", gap: 24, alignItems: "start" }}>
            {/* ─── LEFT: INPUTS ─── */}
            <div>
              {/* Sector Selector */}
              <div style={{ background: colors.surface, border: `1px solid ${colors.border}`, borderRadius: 12, padding: 20, marginBottom: 20 }}>
                <h3 style={{ fontSize: 14, fontWeight: 600, marginBottom: 14, color: colors.textMuted, textTransform: "uppercase", letterSpacing: "0.05em" }}>Selecione o Setor</h3>
                <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(140px, 1fr))", gap: 8 }}>
                  {SECTORS.map((s) => (
                    <button key={s.id} onClick={() => handleSectorSelect(s.id)} style={{ padding: "10px 8px", borderRadius: 8, border: `1px solid ${selectedSector === s.id ? colors.accent : colors.border}`, background: selectedSector === s.id ? colors.accentDim : colors.surfaceAlt, color: selectedSector === s.id ? colors.accent : colors.text, fontSize: 12, fontWeight: 500, cursor: "pointer", transition: "all 0.2s", display: "flex", flexDirection: "column", alignItems: "center", gap: 4 }}>
                      <span style={{ fontSize: 20 }}>{s.icon}</span>
                      <span>{s.label}</span>
                      <span style={{ fontSize: 10, color: colors.textDim, textTransform: "uppercase" }}>{s.type}</span>
                    </button>
                  ))}
                </div>
              </div>

              {/* Input Form */}
              {selectedSector && (
                <div style={{ background: colors.surface, border: `1px solid ${colors.border}`, borderRadius: 12, padding: 20, marginBottom: 20 }}>
                  <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16 }}>
                    <div>
                      <h3 style={{ fontSize: 15, fontWeight: 600, margin: 0 }}>
                        {sectorInfo?.icon} {sectorInfo?.label}
                      </h3>
                      {methodology && (
                        <p style={{ fontSize: 11, color: colors.accent, margin: "4px 0 0 0", fontFamily: "'JetBrains Mono', monospace" }}>
                          {methodology.name}
                        </p>
                      )}
                    </div>
                    <button onClick={() => setShowMethodology(!showMethodology)} style={{ padding: "4px 10px", borderRadius: 4, border: `1px solid ${colors.border}`, background: "transparent", color: colors.textMuted, fontSize: 11, cursor: "pointer" }}>
                      {showMethodology ? "Ocultar" : "Ver"} Metodologia
                    </button>
                  </div>

                  {showMethodology && methodology && (
                    <div style={{ background: colors.surfaceAlt, borderRadius: 8, padding: 12, marginBottom: 16, borderLeft: `3px solid ${colors.accent}` }}>
                      <p style={{ fontSize: 12, fontWeight: 600, margin: "0 0 4px 0", color: colors.accent }}>{methodology.name}</p>
                      <p style={{ fontSize: 12, margin: "0 0 4px 0", fontFamily: "'JetBrains Mono', monospace", color: colors.text }}>{methodology.formula}</p>
                      <p style={{ fontSize: 11, margin: 0, color: colors.textMuted }}>Margem de seguran\u00E7a: {margins[selectedSector]}% — {MARGIN_RATIONALE[selectedSector]}</p>
                    </div>
                  )}

                  <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
                    {SECTOR_FIELDS[selectedSector]?.map((field) => (
                      <div key={field.key} style={{ gridColumn: field.key === "ticker" ? "1 / -1" : undefined }}>
                        <label style={{ fontSize: 11, fontWeight: 500, color: colors.textMuted, display: "block", marginBottom: 4 }}>
                          {field.label}
                        </label>
                        <input
                          type={field.type}
                          step={field.step}
                          placeholder={field.placeholder || ""}
                          value={inputs[field.key] || ""}
                          onChange={(e) => handleInputChange(field.key, e.target.value)}
                          style={{ width: "100%", padding: "8px 10px", borderRadius: 6, border: `1px solid ${colors.border}`, background: colors.surfaceAlt, color: colors.text, fontSize: 13, fontFamily: field.type === "number" ? "'JetBrains Mono', monospace" : "inherit", outline: "none", boxSizing: "border-box", transition: "border-color 0.2s" }}
                          onFocus={(e) => (e.target.style.borderColor = colors.accent)}
                          onBlur={(e) => (e.target.style.borderColor = colors.border)}
                        />
                        {field.hint && <p style={{ fontSize: 10, color: colors.textDim, margin: "2px 0 0 0" }}>{field.hint}</p>}
                      </div>
                    ))}
                  </div>

                  {/* File Upload */}
                  <div style={{ marginTop: 16, padding: 12, borderRadius: 8, border: `1px dashed ${colors.borderLight}`, background: colors.surfaceAlt, textAlign: "center" }}>
                    <label style={{ cursor: "pointer", display: "block" }}>
                      <input type="file" accept=".xlsx,.xls,.csv,.pdf" onChange={handleFileUpload} style={{ display: "none" }} />
                      <div style={{ fontSize: 12, color: colors.textMuted }}>
                        {pdfLoading ? "Processando PDF..." : uploadFileName ? uploadFileName : "Upload relat\u00F3rio trimestral (.xlsx, .csv, .pdf)"}
                      </div>
                    </label>
                    {uploadedData && !uploadedData.error && (
                      <p style={{ fontSize: 10, color: colors.green, marginTop: 4, marginBottom: 0 }}>
                        Arquivo carregado — {typeof uploadedData === "object" && !uploadedData.info ? Object.keys(uploadedData).length + " aba(s) encontrada(s)" : uploadedData.info || ""}
                      </p>
                    )}
                    {uploadedData?.error && (
                      <p style={{ fontSize: 10, color: colors.red, marginTop: 4, marginBottom: 0 }}>{uploadedData.error}</p>
                    )}
                  </div>

                  {/* PDF Extracted Fields */}
                  {pdfData && pdfData.extractedFields && Object.keys(pdfData.extractedFields).length > 0 && (
                    <div style={{ marginTop: 12, borderRadius: 8, border: `1px solid ${colors.accent}44`, background: colors.surfaceAlt, overflow: "hidden" }}>
                      <div style={{ padding: "8px 12px", background: colors.accentDim, display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                        <span style={{ fontSize: 12, fontWeight: 600, color: colors.accent }}>Campos detectados no PDF</span>
                        <button onClick={handleApplyPdfFields} style={{ padding: "4px 10px", borderRadius: 4, border: `1px solid ${colors.accent}`, background: colors.accent + "22", color: colors.accent, fontSize: 11, fontWeight: 600, cursor: "pointer" }}>
                          Preencher todos
                        </button>
                      </div>
                      <div style={{ padding: 8 }}>
                        {Object.entries(pdfData.extractedFields).map(([key, value]) => {
                          const fieldDef = SECTOR_FIELDS[selectedSector]?.find((f) => f.key === key);
                          const label = fieldDef ? fieldDef.label : key;
                          const alreadyFilled = inputs[key] && inputs[key] !== "";
                          return (
                            <div key={key} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "4px 8px", borderBottom: `1px solid ${colors.border}`, fontSize: 11 }}>
                              <span style={{ color: colors.textMuted }}>{label}</span>
                              <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
                                <span style={{ fontWeight: 600, fontFamily: "'JetBrains Mono', monospace", color: colors.text }}>{value}</span>
                                <button onClick={() => handleApplySingleField(key, value)} style={{ padding: "2px 6px", borderRadius: 3, border: `1px solid ${alreadyFilled ? colors.warning : colors.accent}`, background: "transparent", color: alreadyFilled ? colors.warning : colors.accent, fontSize: 10, cursor: "pointer" }}>
                                  {alreadyFilled ? "Substituir" : "Usar"}
                                </button>
                              </div>
                            </div>
                          );
                        })}
                      </div>
                    </div>
                  )}

                  {/* PDF Text Preview */}
                  {pdfData && pdfData.pages && (
                    <div style={{ marginTop: 12, maxHeight: 180, overflow: "auto", borderRadius: 8, border: `1px solid ${colors.border}` }}>
                      <div style={{ padding: "6px 10px", background: colors.surfaceAlt, fontSize: 11, fontWeight: 600, color: colors.purple, position: "sticky", top: 0, zIndex: 1, borderBottom: `1px solid ${colors.border}` }}>
                        Texto extra\u00EDdo ({pdfData.numPages} p\u00E1gina{pdfData.numPages > 1 ? "s" : ""})
                      </div>
                      {pdfData.pages.map((pg) => (
                        <div key={pg.pageNum} style={{ padding: "6px 10px", borderBottom: `1px solid ${colors.border}` }}>
                          <div style={{ fontSize: 10, color: colors.textDim, marginBottom: 2 }}>P\u00E1gina {pg.pageNum}</div>
                          <div style={{ fontSize: 10, color: colors.text, fontFamily: "'JetBrains Mono', monospace", whiteSpace: "pre-wrap", wordBreak: "break-word", lineHeight: 1.5 }}>
                            {pg.text.length > 800 ? pg.text.slice(0, 800) + "..." : pg.text}
                          </div>
                        </div>
                      ))}
                    </div>
                  )}

                  {/* Uploaded Data Preview (Excel/CSV) */}
                  {uploadedData && typeof uploadedData === "object" && !uploadedData.error && !uploadedData.info && (
                    <div style={{ marginTop: 12, maxHeight: 200, overflow: "auto", borderRadius: 8, border: `1px solid ${colors.border}` }}>
                      {Object.entries(uploadedData).map(([sheetName, rows]) => (
                        <div key={sheetName}>
                          <div style={{ padding: "6px 10px", background: colors.surfaceAlt, fontSize: 11, fontWeight: 600, color: colors.accent, position: "sticky", top: 0 }}>
                            {sheetName}
                          </div>
                          <table style={{ width: "100%", fontSize: 10, borderCollapse: "collapse" }}>
                            <tbody>
                              {rows.slice(0, 10).map((row, i) => (
                                <tr key={i} style={{ borderBottom: `1px solid ${colors.border}` }}>
                                  {row.map((cell, j) => (
                                    <td key={j} style={{ padding: "3px 6px", color: i === 0 ? colors.textMuted : colors.text, fontWeight: i === 0 ? 600 : 400, fontFamily: "'JetBrains Mono', monospace", whiteSpace: "nowrap" }}>
                                      {cell}
                                    </td>
                                  ))}
                                </tr>
                              ))}
                            </tbody>
                          </table>
                          {rows.length > 10 && <div style={{ padding: "4px 10px", fontSize: 10, color: colors.textDim }}>... +{rows.length - 10} linhas</div>}
                        </div>
                      ))}
                    </div>
                  )}

                  {/* Calculate Button */}
                  <button onClick={handleCalculate} style={{ width: "100%", marginTop: 16, padding: "12px", borderRadius: 8, border: "none", background: `linear-gradient(135deg, ${colors.accent}, ${colors.blue})`, color: "#fff", fontSize: 14, fontWeight: 700, cursor: "pointer", letterSpacing: "0.02em", transition: "opacity 0.2s" }} onMouseEnter={(e) => (e.target.style.opacity = 0.9)} onMouseLeave={(e) => (e.target.style.opacity = 1)}>
                    CALCULAR VALUATION
                  </button>
                </div>
              )}
            </div>

            {/* ─── RIGHT: RESULTS ─── */}
            {result && (
              <div>
                {/* Verdict Card */}
                <div style={{ background: colors.surface, border: `1px solid ${colors.border}`, borderRadius: 12, padding: 20, marginBottom: 20, position: "relative", overflow: "hidden" }}>
                  <div style={{ position: "absolute", top: 0, right: 0, width: 120, height: 120, borderRadius: "0 0 0 120px", background: verdictBg, opacity: 0.5 }} />
                  <div style={{ position: "relative" }}>
                    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 16 }}>
                      <div>
                        <h2 style={{ fontSize: 24, fontWeight: 700, margin: 0, fontFamily: "'JetBrains Mono', monospace" }}>{result.ticker}</h2>
                        <p style={{ fontSize: 12, color: colors.textMuted, margin: "2px 0 0 0" }}>{result.sectorLabel} - {result.timestamp}</p>
                      </div>
                      <div style={{ padding: "6px 16px", borderRadius: 6, background: verdictBg, color: verdictColor, fontWeight: 700, fontSize: 14, border: `1px solid ${verdictColor}44` }}>
                        {result.verdict}
                      </div>
                    </div>

                    <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 12, marginBottom: 16 }}>
                      <div style={{ padding: 12, background: colors.surfaceAlt, borderRadius: 8, textAlign: "center" }}>
                        <div style={{ fontSize: 10, color: colors.textMuted, textTransform: "uppercase", letterSpacing: "0.05em" }}>Pre\u00E7o Atual</div>
                        <div style={{ fontSize: 20, fontWeight: 700, color: colors.blue, fontFamily: "'JetBrains Mono', monospace" }}>R$ {result.currentPrice.toFixed(2)}</div>
                      </div>
                      <div style={{ padding: 12, background: colors.surfaceAlt, borderRadius: 8, textAlign: "center" }}>
                        <div style={{ fontSize: 10, color: colors.textMuted, textTransform: "uppercase", letterSpacing: "0.05em" }}>Pre\u00E7o Justo</div>
                        <div style={{ fontSize: 20, fontWeight: 700, color: colors.accent, fontFamily: "'JetBrains Mono', monospace" }}>R$ {result.fairPrice.toFixed(2)}</div>
                        <div style={{ fontSize: 11, color: result.upside >= 0 ? colors.green : colors.red }}>
                          {result.upside >= 0 ? "\u25B2" : "\u25BC"} {result.upside.toFixed(1)}%
                        </div>
                      </div>
                      <div style={{ padding: 12, background: colors.surfaceAlt, borderRadius: 8, textAlign: "center" }}>
                        <div style={{ fontSize: 10, color: colors.textMuted, textTransform: "uppercase", letterSpacing: "0.05em" }}>C/ Margem ({result.margin}%)</div>
                        <div style={{ fontSize: 20, fontWeight: 700, color: colors.warning, fontFamily: "'JetBrains Mono', monospace" }}>R$ {result.fairWithMargin.toFixed(2)}</div>
                        <div style={{ fontSize: 11, color: result.upsideWithMargin >= 0 ? colors.green : colors.red }}>
                          {result.upsideWithMargin >= 0 ? "\u25B2" : "\u25BC"} {result.upsideWithMargin.toFixed(1)}%
                        </div>
                      </div>
                    </div>

                    {/* Price Comparison Chart */}
                    <div style={{ height: 220, marginBottom: 16 }}>
                      <ResponsiveContainer width="100%" height="100%">
                        <BarChart data={chartData} barCategoryGap="25%">
                          <CartesianGrid strokeDasharray="3 3" stroke={colors.border} />
                          <XAxis dataKey="name" tick={{ fill: colors.textMuted, fontSize: 11 }} axisLine={{ stroke: colors.border }} />
                          <YAxis tick={{ fill: colors.textMuted, fontSize: 11, fontFamily: "'JetBrains Mono', monospace" }} axisLine={{ stroke: colors.border }} tickFormatter={(v) => `R$${v}`} />
                          <Tooltip contentStyle={{ background: colors.surface, border: `1px solid ${colors.border}`, borderRadius: 8, fontSize: 12 }} formatter={(v) => [`R$ ${v.toFixed(2)}`, ""]} />
                          <Bar dataKey="value" radius={[6, 6, 0, 0]}>
                            {chartData.map((entry, i) => (
                              <Cell key={i} fill={entry.fill} />
                            ))}
                          </Bar>
                          <ReferenceLine y={result.currentPrice} stroke={colors.blue} strokeDasharray="4 4" strokeWidth={1.5} />
                        </BarChart>
                      </ResponsiveContainer>
                    </div>

                    {/* Upside Visual Gauge */}
                    <div style={{ padding: 12, background: colors.surfaceAlt, borderRadius: 8, marginBottom: 16 }}>
                      <div style={{ fontSize: 11, color: colors.textMuted, marginBottom: 8, fontWeight: 500 }}>UPSIDE / DOWNSIDE COM MARGEM</div>
                      <div style={{ position: "relative", height: 28, background: colors.bg, borderRadius: 14, overflow: "hidden" }}>
                        <div style={{ position: "absolute", left: "50%", top: 0, bottom: 0, width: 2, background: colors.textDim, zIndex: 2 }} />
                        <div style={{
                          position: "absolute",
                          top: 2, bottom: 2,
                          borderRadius: 12,
                          background: result.upsideWithMargin >= 0
                            ? `linear-gradient(90deg, transparent, ${colors.green})`
                            : `linear-gradient(90deg, ${colors.red}, transparent)`,
                          ...(result.upsideWithMargin >= 0
                            ? { left: "50%", width: `${Math.min(50, Math.abs(result.upsideWithMargin) / 2)}%` }
                            : { right: "50%", width: `${Math.min(50, Math.abs(result.upsideWithMargin) / 2)}%` }),
                          transition: "width 0.5s ease",
                        }} />
                        <div style={{ position: "absolute", top: "50%", left: "50%", transform: "translate(-50%, -50%)", fontSize: 12, fontWeight: 700, color: colors.text, fontFamily: "'JetBrains Mono', monospace", zIndex: 3 }}>
                          {result.upsideWithMargin >= 0 ? "+" : ""}{result.upsideWithMargin.toFixed(1)}%
                        </div>
                      </div>
                      <div style={{ display: "flex", justifyContent: "space-between", fontSize: 10, color: colors.textDim, marginTop: 4 }}>
                        <span>-50%</span><span>0%</span><span>+50%</span>
                      </div>
                    </div>

                    {/* Details */}
                    <div style={{ borderTop: `1px solid ${colors.border}`, paddingTop: 12 }}>
                      <h4 style={{ fontSize: 12, color: colors.textMuted, textTransform: "uppercase", letterSpacing: "0.05em", marginBottom: 8 }}>Detalhamento</h4>
                      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 6 }}>
                        {Object.entries(result.details).map(([key, value]) => (
                          <div key={key} style={{ display: "flex", justifyContent: "space-between", padding: "4px 8px", background: colors.surfaceAlt, borderRadius: 4, fontSize: 11 }}>
                            <span style={{ color: colors.textMuted }}>{key}</span>
                            <span style={{ fontWeight: 600, fontFamily: "'JetBrains Mono', monospace", color: String(value).includes("\u2705") ? colors.green : String(value).includes("\u26A0") ? colors.warning : colors.text }}>{value}</span>
                          </div>
                        ))}
                      </div>
                    </div>
                  </div>
                </div>
              </div>
            )}
          </div>
        ) : (
          /* ─── HISTORY TAB ─── */
          <div>
            {history.length === 0 ? (
              <div style={{ textAlign: "center", padding: 60, color: colors.textMuted }}>
                <div style={{ fontSize: 48, marginBottom: 12 }}>📊</div>
                <p style={{ fontSize: 14 }}>Nenhum c\u00E1lculo realizado ainda.</p>
                <p style={{ fontSize: 12 }}>Use a calculadora para come\u00E7ar a avaliar ativos.</p>
              </div>
            ) : (
              <>
                {/* History Chart */}
                {historyChartData.length > 1 && (
                  <div style={{ background: colors.surface, border: `1px solid ${colors.border}`, borderRadius: 12, padding: 20, marginBottom: 20 }}>
                    <h3 style={{ fontSize: 14, fontWeight: 600, marginBottom: 16 }}>Comparativo de Ativos</h3>
                    <div style={{ height: 300 }}>
                      <ResponsiveContainer width="100%" height="100%">
                        <BarChart data={historyChartData} barCategoryGap="20%">
                          <CartesianGrid strokeDasharray="3 3" stroke={colors.border} />
                          <XAxis dataKey="name" tick={{ fill: colors.textMuted, fontSize: 11 }} axisLine={{ stroke: colors.border }} />
                          <YAxis tick={{ fill: colors.textMuted, fontSize: 11, fontFamily: "'JetBrains Mono', monospace" }} axisLine={{ stroke: colors.border }} tickFormatter={(v) => `R$${v}`} />
                          <Tooltip contentStyle={{ background: colors.surface, border: `1px solid ${colors.border}`, borderRadius: 8, fontSize: 12 }} formatter={(v) => [`R$ ${v.toFixed(2)}`, ""]} />
                          <Legend wrapperStyle={{ fontSize: 11 }} />
                          <Bar dataKey="Pre\u00E7o Atual" fill={colors.blue} radius={[4, 4, 0, 0]} />
                          <Bar dataKey="Pre\u00E7o Justo" fill={colors.accent} radius={[4, 4, 0, 0]} />
                          <Bar dataKey="Com Margem" fill={colors.warning} radius={[4, 4, 0, 0]} />
                        </BarChart>
                      </ResponsiveContainer>
                    </div>
                  </div>
                )}

                {/* History Cards */}
                <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(300px, 1fr))", gap: 12 }}>
                  {history.map((h, i) => {
                    const vc = h.verdict === "COMPRA" ? colors.green : h.verdict === "NEUTRO" ? colors.warning : colors.red;
                    return (
                      <div key={i} style={{ background: colors.surface, border: `1px solid ${colors.border}`, borderRadius: 10, padding: 16 }}>
                        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8 }}>
                          <div>
                            <span style={{ fontSize: 16, fontWeight: 700, fontFamily: "'JetBrains Mono', monospace" }}>{h.ticker}</span>
                            <span style={{ fontSize: 11, color: colors.textMuted, marginLeft: 8 }}>{h.sectorLabel}</span>
                          </div>
                          <span style={{ padding: "2px 8px", borderRadius: 4, background: vc + "22", color: vc, fontSize: 11, fontWeight: 700 }}>{h.verdict}</span>
                        </div>
                        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 8, fontSize: 11 }}>
                          <div><span style={{ color: colors.textMuted }}>Atual</span><div style={{ fontWeight: 600, fontFamily: "'JetBrains Mono', monospace" }}>R$ {h.currentPrice.toFixed(2)}</div></div>
                          <div><span style={{ color: colors.textMuted }}>Justo</span><div style={{ fontWeight: 600, fontFamily: "'JetBrains Mono', monospace", color: colors.accent }}>R$ {h.fairPrice.toFixed(2)}</div></div>
                          <div><span style={{ color: colors.textMuted }}>Upside</span><div style={{ fontWeight: 600, fontFamily: "'JetBrains Mono', monospace", color: h.upsideWithMargin >= 0 ? colors.green : colors.red }}>{h.upsideWithMargin >= 0 ? "+" : ""}{h.upsideWithMargin.toFixed(1)}%</div></div>
                        </div>
                        <div style={{ fontSize: 10, color: colors.textDim, marginTop: 6 }}>{h.timestamp}</div>
                      </div>
                    );
                  })}
                </div>
              </>
            )}
          </div>
        )}
      </div>

      {/* Footer */}
      <footer style={{ padding: "20px 24px", borderTop: `1px solid ${colors.border}`, textAlign: "center", fontSize: 11, color: colors.textDim, marginTop: 40 }}>
        ValuationPro - Framework Multi-Setor para Valuation de A\u00E7\u00F5es e FIIs brasileiros<br />
        Metodologias: Excess Return, DCF Regulat\u00F3rio, SOTP, NAV/Cap Rate, An\u00E1lise de Cr\u00E9dito<br />
        Ferramenta educacional — n\u00E3o constitui recomenda\u00E7\u00E3o de investimento
      </footer>

      <style>{`
        * { box-sizing: border-box; }
        input[type="range"] { height: 4px; }
        input[type="number"]::-webkit-inner-spin-button { opacity: 0.3; }
        @keyframes fadeIn { from { opacity: 0; transform: translateY(-8px); } to { opacity: 1; transform: translateY(0); } }
        ::-webkit-scrollbar { width: 6px; height: 6px; }
        ::-webkit-scrollbar-track { background: ${colors.bg}; }
        ::-webkit-scrollbar-thumb { background: ${colors.borderLight}; border-radius: 3px; }
      `}</style>
    </div>
  );
}
