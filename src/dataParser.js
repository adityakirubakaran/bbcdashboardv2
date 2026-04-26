import * as XLSX from 'xlsx';

export async function parseExcel(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();

    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        
        const sheetName = "Q1,Q2,Q3";
        const targetSheet = workbook.SheetNames.includes(sheetName) ? sheetName : workbook.SheetNames[0];
        
        const sheet = workbook.Sheets[targetSheet];
        // We inspect the raw grid first because incoming spreadsheets do not always place headers
        // on the first row or use a stable schema between waste and sales exports.
        const rawGrid = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false });
        
        let headerRowIdx = -1;
        let maxScore = 0;
        let finalHeaders = [];
        let isSalesData = false;
        
        for (let i = 0; i < Math.min(30, rawGrid.length); i++) {
            if (!rawGrid[i]) continue;
            const rowStrs = rawGrid[i].map(c => typeof c === 'string' ? c.toUpperCase() : String(c).toUpperCase());
            
            let wasteScore = 0;
            if (rowStrs.some(c => c.includes('DATE'))) wasteScore++;
            if (rowStrs.some(c => c.includes('SITE'))) wasteScore++;
            if (rowStrs.some(c => c.includes('CONTRACT'))) wasteScore++;
            if (rowStrs.some(c => c.includes('LANDFILL'))) wasteScore++;
            
            let salesScore = 0;
            if (rowStrs.some(c => c.includes('SALE DATE'))) salesScore++;
            if (rowStrs.some(c => c.includes('PRICE TOTAL'))) salesScore++;
            if (rowStrs.some(c => c.includes('DESCRIPTION'))) salesScore++;
            if (rowStrs.some(c => c.includes('CLIENT SHARE'))) salesScore++;
            
            let score = Math.max(wasteScore, salesScore);
            
            if (score > maxScore && score >= 2) {
                maxScore = score;
                headerRowIdx = i;
                isSalesData = salesScore > wasteScore;
                finalHeaders = rawGrid[i].map(c => typeof c === 'string' ? c.replace(/\n/g, ' ').trim() : String(c).trim());
            }
        }

        let rawJson = [];
        if (headerRowIdx !== -1) {
            for (let i = headerRowIdx + 1; i < rawGrid.length; i++) {
                const rowArr = rawGrid[i];
                if (!rowArr || rowArr.length === 0) continue;
                
                let rowObj = {};
                for (let j = 0; j < finalHeaders.length; j++) {
                    if (finalHeaders[j]) {
                        rowObj[finalHeaders[j]] = rowArr[j];
                    }
                }
                rawJson.push(rowObj);
            }
        } else {
            rawJson = XLSX.utils.sheet_to_json(sheet, { raw: false });
        }
        let aggregatedData;
        if (isSalesData) {
            aggregatedData = {
                 type: 'sales',
                 metrics: { totalWaste: 0, totalLandfill: 0, diversionRate: 0 },
                 uniqueSites: [],
                 uniqueCategories: [],
                 rawRows: rawJson,
                 salesData: processSalesData(rawJson, finalHeaders)
            };
        } else {
            aggregatedData = processData(rawJson, finalHeaders);
            aggregatedData.type = 'waste';
        }
        aggregatedData.rawHeaders = finalHeaders;
        resolve(aggregatedData);
      } catch (err) {
        reject(err);
      }
    };

    reader.onerror = (err) => reject(err);
    reader.readAsArrayBuffer(file);
  });
}

function extractCategories(headers) {
    const upper = headers.map(h => h.toUpperCase());
    
    // Waste stream columns sit between site metadata and disposal summary columns.
    let startIdx = upper.findIndex(h => h.includes('WASTE STREAM') || h.includes('EWC'));
    if (startIdx === -1) {
        startIdx = upper.findIndex(h => h.includes('BUILDING') || h.includes('POST CODE') || h.includes('SITE'));
    }
    
    let endIdx = upper.findIndex((h, idx) => idx > startIdx && (h.includes('TOTAL') || h.includes('AD ') || h.includes('INCIN') || h.includes('RECYCLE') || h.includes('LANDFILL')));
    if (endIdx === -1) endIdx = headers.length;
    
    let cats = [];
    if (startIdx !== -1 && endIdx > startIdx) {
        cats = headers.slice(startIdx + 1, endIdx).filter(h => {
            if (!h || h.trim().length === 0) return false;
            let upperH = h.toUpperCase();
            if (upperH.includes('WASTE STREAM') || upperH.includes('EWC') || upperH.includes('SITE') || upperH.includes('DATE') || upperH.includes('CONTRACT')) return false;
            return true;
        });
    }
    return cats.length > 0 ? cats : ['General Waste'];
}

export function processData(rows, headers = [], filters = null) {
  // Metadata is collected before filtering so the UI can still render dropdown options even when
  // the current selection produces an empty result set.
  const uniqueCategories = extractCategories(headers);
  const uniqueSites = new Set();
  
  rows.forEach(row => {
    let rawDate = row['DATE'] || row['Date'] || row['date'];
    if (!rawDate) return;
    const site = row['SITE'] || row['Site'] || row['site'] || 'Unknown';
    uniqueSites.add(site);
  });
  const siteList = Array.from(uniqueSites).filter(s => s && s.trim().length > 0);

  if (filters !== null && (filters.sites.length === 0 || filters.categories.length === 0)) {
       return {
        timeline: [],
        metrics: { totalWaste: 0, totalLandfill: 0, diversionRate: 0 },
        uniqueSites: siteList,
        uniqueCategories: uniqueCategories,
        topLocations: [],
        rawRows: rows,
        insights: ["Please select at least one Location and one Category filter."]
      };
  }

  const activeFilters = filters || { sites: siteList, categories: uniqueCategories };

  const monthlyData = {};
  const siteData = {};
  let totalWasteAllTime = 0;
  let totalLandfillAllTime = 0;
  

  rows.forEach(row => {
    let rawDate = row['DATE'] || row['Date'] || row['date'];
    if (!rawDate) return;
    
    let monthStr = String(rawDate).substring(0, 3);
    const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
    if (!months.includes(monthStr)) {
        if (typeof rawDate === 'number') {
            const dateObj = new Date(Math.round((rawDate - 25569)*86400*1000));
            monthStr = months[dateObj.getMonth()];
        } else {
            return;
        }
    }

    const site = row['SITE'] || row['Site'] || row['site'] || 'Unknown';
    uniqueSites.add(site);

    if (activeFilters.sites.length > 0 && !activeFilters.sites.includes(site)) return;
    
    // Disposal columns are totals for the whole row, not per-category values. When the user filters
    // down to specific categories, we scale landfill/diversion totals by that category share so the
    // charts stay directionally accurate instead of dropping the row entirely.
    let rowCategoryTotalSum = 0;
    let rowCategoryFilteredSum = 0;
    
    uniqueCategories.forEach(cat => {
        let v = parseFloat(row[cat]) || 0;
        rowCategoryTotalSum += v;
        if (activeFilters.categories.length === 0 || activeFilters.categories.includes(cat)) {
            rowCategoryFilteredSum += v;
        }
    });

    if (activeFilters.categories.length > 0 && rowCategoryFilteredSum === 0) return;

    let ratio = rowCategoryTotalSum > 0 ? (rowCategoryFilteredSum / rowCategoryTotalSum) : 0;
    if (activeFilters.categories.length === 0) ratio = 1;

    const landfill = (parseFloat(row['Landfill']) || parseFloat(row['LANDFILL']) || 0) * ratio;
    const recycled = (parseFloat(row['Recycled']) || parseFloat(row['RECYCLED']) || 0) * ratio;
    const incineration = (parseFloat(row['Incineration']) || parseFloat(row['INCINERATION']) || 0) * ratio;
    const ad = (parseFloat(row['AD ']) || parseFloat(row['AD']) || 0) * ratio;
    
    const diverted = recycled + incineration + ad;
    const totalGen = landfill + diverted;

    totalWasteAllTime += totalGen;
    totalLandfillAllTime += landfill;

    if (!siteData[site]) {
        siteData[site] = { site, total: 0, diverted: 0, landfill: 0 };
    }
    siteData[site].total += totalGen;
    siteData[site].landfill += landfill;
    siteData[site].diverted += diverted;

    if (!monthlyData[monthStr]) {
      monthlyData[monthStr] = { 
          month: monthStr, 
          landfill: 0, 
          diverted: 0, 
          total: 0,
          rawAd: 0,
          rawIncineration: 0,
          rawRecycled: 0,
          rawLandfill: 0,
          categoryVolumes: {}
      };
    }

    monthlyData[monthStr].landfill += landfill;
    monthlyData[monthStr].diverted += diverted;
    monthlyData[monthStr].total += totalGen;
    
    monthlyData[monthStr].rawAd += ad;
    monthlyData[monthStr].rawIncineration += incineration;
    monthlyData[monthStr].rawRecycled += recycled;
    monthlyData[monthStr].rawLandfill += landfill;

    uniqueCategories.forEach(cat => {
        if (activeFilters.categories.length === 0 || activeFilters.categories.includes(cat)) {
            if (!monthlyData[monthStr].categoryVolumes[cat]) monthlyData[monthStr].categoryVolumes[cat] = 0;
            monthlyData[monthStr].categoryVolumes[cat] += (parseFloat(row[cat]) || 0);
        }
    });
  });

  const monthOrder = { 'Jan': 1, 'Feb': 2, 'Mar': 3, 'Apr': 4, 'May': 5, 'Jun': 6, 'Jul': 7, 'Aug': 8, 'Sep': 9, 'Oct': 10, 'Nov': 11, 'Dec': 12 };
  const sortedMonths = Object.values(monthlyData).sort((a, b) => monthOrder[a.month] - monthOrder[b.month]);
  const insights = generateInsights(sortedMonths, totalWasteAllTime, totalLandfillAllTime, siteData);
  
  const topLocations = Object.values(siteData).sort((a, b) => b.total - a.total).slice(0, 20);

  return {
    timeline: sortedMonths,
    metrics: {
      totalWaste: totalWasteAllTime,
      totalLandfill: totalLandfillAllTime,
      diversionRate: totalWasteAllTime > 0 ? ((totalWasteAllTime - totalLandfillAllTime) / totalWasteAllTime) * 100 : 0
    },
    uniqueSites: Array.from(uniqueSites).filter(s => s && s.trim().length > 0),
    uniqueCategories: uniqueCategories,
    topLocations: topLocations,
    rawRows: rows,
    insights: insights
  };
}

function generateInsights(timeline, totalWaste, totalLandfill, siteData) {
  const insightsMap = {
      'chart-disposition-detailed': [],
      'chart-disposition-diverted': [],
      'chart-locations': [],
      'chart-category': []
  };
  
  if (totalWaste === 0) {
      insightsMap['chart-disposition-detailed'].push("No waste generation data detected across the provided timeframe.");
      return insightsMap;
  }

  const divRate = ((totalWaste - totalLandfill) / totalWaste) * 100;
  if (divRate > 80) {
      insightsMap['chart-disposition-diverted'].push(`Excellent diversion rate of ${divRate.toFixed(1)}%. The majority of waste is successfully bypassing landfill deposition.`);
      insightsMap['chart-category'].push(`Excellent diversion rate of ${divRate.toFixed(1)}%.`);
  } else if (divRate < 50) {
      insightsMap['chart-disposition-diverted'].push(`Diversion rate is tracking low at ${divRate.toFixed(1)}%. Consider evaluating sorting protocols at primary generation sites.`);
      insightsMap['chart-category'].push(`Diversion rate is low at ${divRate.toFixed(1)}%. Check primary streams.`);
  } else {
      insightsMap['chart-disposition-diverted'].push(`Average performance with a diversion rate of ${divRate.toFixed(1)}%.`);
  }

  if (timeline.length >= 2) {
      const lastMonth = timeline[timeline.length - 1];
      const prevMonth = timeline[timeline.length - 2];
      
      if (lastMonth.total > 0 && prevMonth.total > 0) {
          const change = ((lastMonth.total - prevMonth.total) / prevMonth.total) * 100;
          if (change > 15) {
              insightsMap['chart-disposition-detailed'].push(`Alert: Total waste volume surged by ${change.toFixed(1)}% in ${lastMonth.month} compared to ${prevMonth.month}.`);
          } else if (change < -15) {
              insightsMap['chart-disposition-detailed'].push(`Positive Trend: Waste volume decreased by ${Math.abs(change).toFixed(1)}% in ${lastMonth.month}.`);
          }
      }
      
      const lastMonthLandfillProp = lastMonth.landfill / lastMonth.total;
      const prevMonthLandfillProp = prevMonth.landfill / prevMonth.total;
      if (lastMonthLandfillProp > prevMonthLandfillProp + 0.1) {
          insightsMap['chart-disposition-detailed'].push(`Landfill dependency increased sharply in ${lastMonth.month}. This typically correlates with non-recyclable batch purges.`);
      }
  }

  if (siteData) {
      const topSites = Object.values(siteData).sort((a, b) => b.total - a.total);
      if (topSites.length > 0) {
          const worstSite = topSites[0];
          insightsMap['chart-locations'].push(`<strong>${worstSite.site}</strong> is the highest volume location, contributing ${Math.round(worstSite.total).toLocaleString()} tonnes to the footprint.`);
      }
      if (topSites.length > 1) {
          const secondSite = topSites[1];
          insightsMap['chart-locations'].push(`<strong>${secondSite.site}</strong> is the second highest contributor, producing ${Math.round(secondSite.total).toLocaleString()} tonnes.`);
      }
  }

  return insightsMap;
}

export function processSalesData(rows, headers = []) {
    const salesByCategory = {};
    const salesByRoute = {};
    let totalRevenue = 0;
    
    rows.forEach(row => {
        const descMatch = (row['Description'] || row['DESCRIPTION'] || '').toUpperCase();
        let cat = 'Other';
        if (descMatch.includes('AUDIO') || descMatch.includes('RYCOTE') || descMatch.includes('WINDSHIELD') || descMatch.includes('MICROPHONE')) cat = 'Audio Equipment';
        else if (descMatch.includes('CAMERA') || descMatch.includes('VIDEO') || descMatch.includes('FILM') || descMatch.includes('GOPRO') || descMatch.includes('OSMO') || descMatch.includes('CAMBLOCK') || descMatch.includes('CAMCORDER') || descMatch.includes('PHOTOGRAPHY')) cat = 'Camera & Video';
        else if (descMatch.includes('ENDOSCOPE') || descMatch.includes('LENS')) cat = 'Specialist Lenses';
        else if (descMatch.includes('LAMP') || descMatch.includes('LIGHT') || descMatch.includes('ILLUMINATOR')) cat = 'Lighting';
        else if (descMatch.includes('SCRAP') || descMatch.includes('METAL')) cat = 'Scrap & Metals';
        else if (descMatch.includes('CASE') || descMatch.includes('PELI') || descMatch.includes('PORTABRACE')) cat = 'Cases & Bags';
        
        const rawPrice = row['Total Due'] || row['TOTAL DUE'] || row['Price Total'] || row['PRICE TOTAL'] || '0';
        const priceStr = String(rawPrice).replace(/[^0-9.-]+/g, "");
        const price = parseFloat(priceStr) || 0;
        const route = row['Sales Route'] || row['SALES ROUTE'] || 'Unknown';
        
        if (!salesByCategory[cat]) salesByCategory[cat] = 0;
        salesByCategory[cat] += price;
        totalRevenue += price;
        
        if (!salesByRoute[route]) salesByRoute[route] = 0;
        salesByRoute[route] += price;
    });

    const categories = Object.keys(salesByCategory).map(name => ({
        category: name,
        revenue: salesByCategory[name]
    })).sort((a,b) => b.revenue - a.revenue);
    
    const topCategory = categories.length > 0 ? categories[0] : null;
    let topRoute = '';
    let maxRouteRev = 0;
    for (const [r, rev] of Object.entries(salesByRoute)) {
        if (rev > maxRouteRev) {
            maxRouteRev = rev;
            topRoute = r;
        }
    }
    
    const insights = [];
    if (totalRevenue > 0) {
        insights.push(`Total recorded sales generated: <strong>£${totalRevenue.toLocaleString()}</strong>.`);
    }
    if (topCategory && topCategory.revenue > 0) {
        const perc = ((topCategory.revenue / totalRevenue) * 100).toFixed(1);
        insights.push(`<strong>${topCategory.category}</strong> drives the majority of revenue, bringing in <strong>£${topCategory.revenue.toLocaleString()}</strong> (${perc}% of total sales).`);
    }
    if (topRoute) {
        insights.push(`The most profitable sales route was <strong>${topRoute}</strong>, capturing <strong>£${maxRouteRev.toLocaleString()}</strong> in value.`);
    }
    
    return {
        categories,
        totalRevenue,
        insights: { 'chart-sales': insights }
    };
}
