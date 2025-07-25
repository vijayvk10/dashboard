import React, { useState, useRef } from "react";
import PptxGenJS from "pptxgenjs";
import Plot from "react-plotly.js";
import * as XLSX from "xlsx";
import "./App.css";

// Helper for CSV download
function downloadCSV(csv, filename) {
  const blob = new Blob([csv], { type: "text/csv" });
  const url = window.URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  a.click();
  window.URL.revokeObjectURL(url);
}

// Helper: convert array of objects to CSV string
function arrayToCSV(data, columns) {
  if (!data || !data.length) return "";
  const header = columns.join(",");
  const rows = data.map((row) =>
    columns.map((col) => `"${row[col] ?? ""}"`).join(",")
  );
  return [header, ...rows].join("\r\n");
}

// Helper: get unique values from a column
function uniqueValues(data, col) {
  return Array.from(new Set(data.map((row) => row[col]).filter(Boolean)));
}

// Helper: extract month/year from column names
function extractMonthYear(col) {
  const m = String(col).match(/(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[-–\s]?(\d{2,4})?/i);
  if (m) return m[0];
  return "";
}

function App() {
  // --- Master PPT Generation ---
  // Helper to collect all chart data for master PPT
  function getAllTabChartData() {
    const tabLabels = [
      "Budget vs Actual", "Budget", "LY", "Act", "Gr", "Ach",
      "YTD Budget", "YTD LY", "YTD Act", "YTD Gr", "YTD Ach",
      "Branch Performance", "Branch Monthwise",
      "Product Performance", "Product Monthwise"
    ];
    return tabLabels.map(tab => ({
      label: tab,
      chart: getTabChartData(tab, dataTable, columns, filters, visualType)
    }));
  }

  // Handler to generate and download master PPT
  async function handleGenerateMasterPPT() {
    const pptx = new PptxGenJS();
    const allCharts = getAllTabChartData();

    allCharts.forEach(({ label, chart }) => {
      if (chart.data && chart.data.length > 0) {
        const slide = pptx.addSlide();
        slide.addText(label, { x: 0.5, y: 0.2, fontSize: 18, bold: true });
        // Only add chart if it's a bar/line/pie
        const chartType = chart.data[0]?.type;
        if (["bar", "scatter", "pie"].includes(chartType)) {
          // For scatter, treat as line chart
          let pptxType = chartType === "scatter" ? pptx.ChartType.line : chartType;
          // Prepare chart data for pptxgenjs
          let pptxData = [];
          if (chartType === "pie") {
            pptxData = [{
              name: label,
              labels: chart.data[0].labels,
              values: chart.data[0].values
            }];
          } else if (chartType === "bar" || chartType === "scatter") {
            // Support grouped bar/line
            if (Array.isArray(chart.data)) {
              pptxData = chart.data.map(d => ({
                name: d.name || label,
                labels: d.x,
                values: d.y
              }));
            }
          }
          slide.addChart(pptxType, pptxData, { x: 0.5, y: 1, w: 8, h: 4 });
        }
      }
    });

    await pptx.writeFile({ fileName: "charts_only_master.pptx" });
  }

  // Handler to generate and download a single PPT for the active tab
  async function handleDownloadTabPPT() {
    const pptx = new PptxGenJS();
    const chart = getTabChartData(activeTab, dataTable, columns, filters, visualType);
    if (chart.data && chart.data.length > 0) {
      const slide = pptx.addSlide();
      slide.addText(activeTab, { x: 0.5, y: 0.2, fontSize: 18, bold: true });
      const chartType = chart.data[0]?.type;
      if (["bar", "scatter", "pie"].includes(chartType)) {
        let pptxType = chartType === "scatter" ? pptx.ChartType.line : chartType;
        let pptxData = [];
        if (chartType === "pie") {
          pptxData = [{
            name: activeTab,
            labels: chart.data[0].labels,
            values: chart.data[0].values
          }];
        } else if (chartType === "bar" || chartType === "scatter") {
          if (Array.isArray(chart.data)) {
            pptxData = chart.data.map(d => ({
              name: d.name || activeTab,
              labels: d.x,
              values: d.y
            }));
          }
        }
        slide.addChart(pptxType, pptxData, { x: 0.5, y: 1, w: 8, h: 4 });
      }
      await pptx.writeFile({ fileName: `${activeTab.replace(/\s+/g, "_").toLowerCase()}_chart.pptx` });
    }
  }
  // State
  const [excelFile, setExcelFile] = useState(null);
  const [sheetNames, setSheetNames] = useState([]);
  const [selectedSheet, setSelectedSheet] = useState("");
  const [tableOptions, setTableOptions] = useState([]);
  const [selectedTable, setSelectedTable] = useState("");
  const [filters, setFilters] = useState({
    month: "Select All",
    year: "Select All",
    branch: "Select All",
    product: "Select All",
  });
  const [filterOptions, setFilterOptions] = useState({
    months: ["Select All"],
    years: ["Select All"],
    branches: ["Select All"],
    products: ["Select All"],
  });
  const [dataTable, setDataTable] = useState([]);
  const [columns, setColumns] = useState([]);
  const [visualType, setVisualType] = useState("Bar Chart");
  const [chartData, setChartData] = useState(null);
  const [chartLayout, setChartLayout] = useState({});
  const [csvData, setCsvData] = useState("");
  const [loading, setLoading] = useState(false);

  const [workbook, setWorkbook] = useState(null);
  const [rawTables, setRawTables] = useState({}); // {sheet: {tableName: [rows]}}

  const fileInputRef = useRef();

  // 1. Upload Excel file and parse
  const handleFileChange = async (e) => {
    const file = e.target.files[0];
    setExcelFile(file);
    setLoading(true);
    const reader = new FileReader();
    reader.onload = (evt) => {
      const data = evt.target.result;
      const wb = XLSX.read(data, { type: "array" });
      setWorkbook(wb);
      setSheetNames(wb.SheetNames);
      setSelectedSheet("");
      setTableOptions([]);
      setSelectedTable("");
      setDataTable([]);
      setColumns([]);
      setFilterOptions({
        months: ["Select All"],
        years: ["Select All"],
        branches: ["Select All"],
        products: ["Select All"],
      });
      setFilters({
        month: "Select All",
        year: "Select All",
        branch: "Select All",
        product: "Select All",
      });
      setRawTables({});
      setLoading(false);
    };
    reader.readAsArrayBuffer(file);
  };

  // 2. Select sheet and detect tables
  const handleSheetSelect = (e) => {
    const sheet = e.target.value;
    setSelectedSheet(sheet);
    setLoading(true);
    setTimeout(() => {
      const ws = workbook.Sheets[sheet];
      // Parse all rows as arrays
      const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
      // Heuristic: find table header rows (look for "sales in" or "budget/act/ly" etc)
      let tableStarts = [];
      rows.forEach((row, idx) => {
        const rowText = row.join(" ").toLowerCase();
        if (
          /sales\s*in\s*(mt|value|tonnage|tonage)/.test(rowText) ||
          /\bbudget\b|\bact\b|\bly\b|\bgr\b|\bach\b/.test(rowText)
        ) {
          tableStarts.push(idx);
        }
      });
      // Build table options
      let tables = [];
      let tableMap = {};
      for (let i = 0; i < tableStarts.length; ++i) {
        const start = tableStarts[i];
        const end = tableStarts[i + 1] || rows.length;
        // Find header row (first non-empty row after start)
        let headerRowIdx = start;
        while (
          headerRowIdx < end &&
          rows[headerRowIdx].filter((x) => x && String(x).trim()).length < 2
        )
          headerRowIdx++;
        if (headerRowIdx >= end) continue;
        const header = rows[headerRowIdx];
        // Data rows
        const dataRows = [];
        for (let j = headerRowIdx + 1; j < end; ++j) {
          const row = rows[j];
          if (row.filter((x) => x && String(x).trim()).length === 0) continue;
          let obj = {};
          header.forEach((col, k) => {
            obj[String(col).trim() || `Col${k + 1}`] = row[k];
          });
          dataRows.push(obj);
        }
        const tableName =
          "Table " +
          (i + 1) +
          ": " +
          (header.join(" ").slice(0, 30) || "Unnamed Table");
        tables.push(tableName);
        tableMap[tableName] = dataRows;
      }
      setTableOptions(tables);
      setRawTables((prev) => ({ ...prev, [sheet]: tableMap }));
      setSelectedTable("");
      setDataTable([]);
      setColumns([]);
      setLoading(false);
    }, 100);
  };

  // 3. Select table and extract filter options
  const handleTableSelect = (e) => {
    const table = e.target.value;
    setSelectedTable(table);
    setLoading(true);
    setTimeout(() => {
      const tableData = rawTables[selectedSheet][table] || [];
      const cols = tableData.length ? Object.keys(tableData[0]) : [];
      // Detect filter options
      let months = [];
      let years = [];
      let branches = [];
      let products = [];
      cols.forEach((col) => {
        const m = extractMonthYear(col);
        if (m) months.push(m);
        const y = String(col).match(/[-–](\d{2,4})/);
        if (y) years.push(y[1]);
      });
      // Try to guess branch/product columns
      if (cols.length) {
        const firstCol = cols[0];
        const values = uniqueValues(tableData, firstCol);
        if (
          values.some((v) =>
            String(v).toLowerCase().includes("branch") ||
            String(v).toLowerCase().includes("region")
          )
        ) {
          branches = values;
        } else if (
          values.some((v) =>
            String(v).toLowerCase().includes("product")
          )
        ) {
          products = values;
        } else if (values.length < 30) {
          // Heuristic: if <30 unique, treat as branch/product
          branches = values;
        }
      }
      months = Array.from(new Set(months));
      years = Array.from(new Set(years));
      branches = Array.from(new Set(branches));
      products = Array.from(new Set(products));
      setFilterOptions({
        months: ["Select All", ...months],
        years: ["Select All", ...years],
        branches: ["Select All", ...branches.filter((x) => x)],
        products: ["Select All", ...products.filter((x) => x)],
      });
      setFilters({
        month: "Select All",
        year: "Select All",
        branch: "Select All",
        product: "Select All",
      });
      setColumns(cols);
      setDataTable(tableData);
      setCsvData(arrayToCSV(tableData, cols));
      setLoading(false);
      // Initial chart
      fetchChartData(tableData, cols, {
        month: "Select All",
        year: "Select All",
        branch: "Select All",
        product: "Select All",
      }, visualType);
    }, 100);
  };

  // 4. Change filters
  const handleFilterChange = (e) => {
    const { name, value } = e.target;
    const newFilters = { ...filters, [name]: value };
    setFilters(newFilters);
    // Filter data
    const tableData = rawTables[selectedSheet][selectedTable] || [];
    let filtered = [...tableData];
    const cols = columns;
    // Filter by branch/product
    if (newFilters.branch !== "Select All" && cols.length) {
      filtered = filtered.filter((row) => row[cols[0]] === newFilters.branch);
    }
    if (newFilters.product !== "Select All" && cols.length) {
      filtered = filtered.filter((row) => row[cols[0]] === newFilters.product);
    }
    setDataTable(filtered);
    setCsvData(arrayToCSV(filtered, cols));
    fetchChartData(filtered, cols, newFilters, visualType);
  };

  // 5. Change visualization type
  const handleVisualTypeChange = (e) => {
    setVisualType(e.target.value);
    fetchChartData(dataTable, columns, filters, e.target.value);
  };

  // --- Chart Data Generation ---
  function fetchChartData(tableData, cols, appliedFilters, visType) {
    // Heuristic: use first column as category, next as value, or melt columns for months
    if (!tableData || !cols.length) {
      setChartData(null);
      return;
    }
    let x = [];
    let y = [];
    let chartType = visType === "Pie Chart" ? "pie" : visType === "Line Chart" ? "scatter" : "bar";
    let plotData = [];
    // Try to find month columns
    const monthCols = cols.filter((col) => extractMonthYear(col));
    if (monthCols.length) {
      // Melt data for months
      let melted = [];
      tableData.forEach((row) => {
        monthCols.forEach((col) => {
          melted.push({
            category: row[cols[0]],
            month: extractMonthYear(col),
            value: Number(row[col]) || 0,
          });
        });
      });
      // Filter by month/year if needed
      let filtered = melted;
      if (appliedFilters.month !== "Select All") {
        filtered = filtered.filter((r) => r.month === appliedFilters.month);
      }
      // Aggregate by category
      let grouped = {};
      filtered.forEach((r) => {
        if (!grouped[r.category]) grouped[r.category] = 0;
        grouped[r.category] += r.value;
      });
      x = Object.keys(grouped);
      y = Object.values(grouped);
      if (chartType === "pie") {
        plotData = [
          {
            type: "pie",
            labels: x,
            values: y,
            textinfo: "percent+label",
            hoverinfo: "label+value+percent",
          },
        ];
      } else {
        plotData = [
          {
            type: chartType,
            x,
            y,
            marker: { color: "#2E86AB" },
            ...(chartType === "scatter" ? { mode: "lines+markers" } : {}),
          },
        ];
      }
      setChartData(plotData);
      setChartLayout({
        title: `${visType} of ${cols[0]}${appliedFilters.month !== "Select All" ? " - " + appliedFilters.month : ""}`,
        xaxis: { title: cols[0] },
        yaxis: { title: "Value" },
        autosize: true,
      });
    } else if (cols.length >= 2) {
      // Use first col as x, second as y
      x = tableData.map((row) => row[cols[0]]);
      y = tableData.map((row) => Number(row[cols[1]]) || 0);
      if (chartType === "pie") {
        plotData = [
          {
            type: "pie",
            labels: x,
            values: y,
            textinfo: "percent+label",
            hoverinfo: "label+value+percent",
          },
        ];
      } else {
        plotData = [
          {
            type: chartType,
            x,
            y,
            marker: { color: "#2E86AB" },
            ...(chartType === "scatter" ? { mode: "lines+markers" } : {}),
          },
        ];
      }
      setChartData(plotData);
      setChartLayout({
        title: `${visType} of ${cols[1]} by ${cols[0]}`,
        xaxis: { title: cols[0] },
        yaxis: { title: cols[1] },
        autosize: true,
      });
    } else {
      setChartData(null);
    }
  }

  // --- Download CSV ---
  const handleDownloadCSV = () => {
    if (csvData) downloadCSV(csvData, "filtered_data.csv");
  };

  // Add tab names as in the Python dashboard
  const tabNames = [
    "📊 Budget vs Actual", "📊 Budget", "📊 LY", "📊 Act", "📊 Gr", "📊 Ach",
    "📈 YTD Budget", "📈 YTD LY", "📈 YTD Act", "📈 YTD Gr", "📈 YTD Ach",
    "🌍 Branch Performance", "🌍 Branch Monthwise",
    "📦 Product Performance", "📦 Product Monthwise"
  ];
  const [activeTab, setActiveTab] = useState(tabNames[0]);

  // Add this state for the Budget vs Actual data table expander
  const [showBvATable, setShowBvATable] = useState(false);

  // --- Chart Data Generation for each tab ---
  // Helper to strip emoji and whitespace from tab label (safe for JS regex)
  function stripEmoji(label) {
    // Remove leading emoji and whitespace
    return label.replace(/^[^\p{L}\p{N}]+/u, '').trim();
  }
  function getTabChartData(tabLabel, tableData, cols, appliedFilters, visType) {
    // Always strip emoji for logic
    const cleanTabLabel = stripEmoji(tabLabel);
    // Helper for melting data
    const melt = (data, idVar, valueVars, varName, valueName) => {
      let out = [];
      data.forEach(row => {
        valueVars.forEach(col => {
          out.push({
            [idVar]: row[idVar],
            [varName]: col,
            [valueName]: Number(row[col]) || 0
          });
        });
      });
      return out;
    };

    // Budget vs Actual
    if (cleanTabLabel === "Budget vs Actual") {
      const budgetCols = cols.filter(col => /^budget(?!.*ytd)/i.test(col));
      const actCols = cols.filter(col => /^act(?!.*ytd)/i.test(col));
      if (budgetCols.length && actCols.length) {
        let melted = [];
        tableData.forEach(row => {
          budgetCols.forEach(col => {
            const val = Number(row[col]);
            if (!isNaN(val) && row[col] !== "" && row[col] !== null) {
              melted.push({ Month: extractMonthYear(col), Metric: "Budget", Value: val });
            }
          });
          actCols.forEach(col => {
            const val = Number(row[col]);
            if (!isNaN(val) && row[col] !== "" && row[col] !== null) {
              melted.push({ Month: extractMonthYear(col), Metric: "Act", Value: val });
            }
          });
        });
        melted = melted.filter(r => r.Month && r.Metric && !isNaN(r.Value));
        let grouped = {};
        melted.forEach(r => {
          const key = r.Month + "|" + r.Metric;
          if (!grouped[key]) grouped[key] = { Month: r.Month, Metric: r.Metric, Value: 0 };
          grouped[key].Value += r.Value;
        });
        const chartRows = Object.values(grouped);
        const tableRows = chartRows
          .filter(r => r.Month && r.Metric)
          .sort((a, b) => {
            const months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
            const [aMonth, aYear] = (a.Month || "").split("-");
            const [bMonth, bYear] = (b.Month || "").split("-");
            if (aYear !== bYear) return (aYear || "").localeCompare(bYear || "");
            return months.indexOf(aMonth) - months.indexOf(bMonth);
          });
        if (visType === "Pie Chart") {
          const totalBudget = chartRows.filter(r => r.Metric === "Budget").reduce((a, b) => a + b.Value, 0);
          const totalAct = chartRows.filter(r => r.Metric === "Act").reduce((a, b) => a + b.Value, 0);
          return {
            data: [{
              type: "pie",
              labels: ["Budget", "Act"],
              values: [totalBudget, totalAct],
              textinfo: "percent+label",
              hoverinfo: "label+value+percent"
            }],
            layout: { title: "Budget vs Actual" },
            table: tableRows
          };
        } else {
          const months = [...new Set(tableRows.map(r => r.Month))];
          const budgetY = months.map(m => (tableRows.find(r => r.Month === m && r.Metric === "Budget") || {}).Value || 0);
          const actY = months.map(m => (tableRows.find(r => r.Month === m && r.Metric === "Act") || {}).Value || 0);
          return {
            data: [
              {
                type: visType === "Line Chart" ? "scatter" : "bar",
                x: months,
                y: budgetY,
                name: "Budget",
                marker: { color: "#2E86AB" },
                ...(visType === "Line Chart" ? { mode: "lines+markers" } : {})
              },
              {
                type: visType === "Line Chart" ? "scatter" : "bar",
                x: months,
                y: actY,
                name: "Act",
                marker: { color: "#FF8C00" },
                ...(visType === "Line Chart" ? { mode: "lines+markers" } : {})
              }
            ],
            layout: { title: "Budget vs Actual", barmode: "group", xaxis: { title: "Month" }, yaxis: { title: "Value" } },
            table: tableRows
          };
        }
      }
      return { data: [], layout: {}, table: [] };
    }

    // Budget, LY, Act, Gr, Ach (monthly)
    if (["Budget", "LY", "Act", "Gr", "Ach"].includes(cleanTabLabel)) {
      const label = cleanTabLabel;
      const valueCols = cols.filter(col =>
        new RegExp(`^${label}(?!.*ytd)`, "i").test(col)
      );
      if (valueCols.length) {
        let melted = melt(tableData, cols[0], valueCols, "Month", label);
        melted = melted.filter(r => r.Month && !isNaN(r[label]));
        if (appliedFilters.month !== "Select All") {
          melted = melted.filter(r => extractMonthYear(r.Month) === appliedFilters.month);
        }
        let grouped = {};
        melted.forEach(r => {
          const key = r.Month;
          if (!grouped[key]) grouped[key] = { Month: extractMonthYear(r.Month), [label]: 0 };
          grouped[key][label] += r[label];
        });
        const chartRows = Object.values(grouped);
        const tableRows = chartRows
          .filter(r => r.Month)
          .sort((a, b) => {
            const months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
            const [aMonth, aYear] = (a.Month || "").split("-");
            const [bMonth, bYear] = (b.Month || "").split("-");
            if (aYear !== bYear) return (aYear || "").localeCompare(bYear || "");
            return months.indexOf(aMonth) - months.indexOf(bMonth);
          });
        if (visType === "Pie Chart") {
          return {
            data: [{
              type: "pie",
              labels: tableRows.map(r => r.Month),
              values: tableRows.map(r => r[label]),
              textinfo: "percent+label",
              hoverinfo: "label+value+percent"
            }],
            layout: { title: `${label} Distribution` },
            table: tableRows
          };
        } else {
          return {
            data: [{
              type: visType === "Line Chart" ? "scatter" : "bar",
              x: tableRows.map(r => r.Month),
              y: tableRows.map(r => r[label]),
              marker: { color: label === "Act" ? "#FF8C00" : "#2E86AB" },
              ...(visType === "Line Chart" ? { mode: "lines+markers" } : {}),
              name: label
            }],
            layout: { title: `${label} by Month`, xaxis: { title: "Month" }, yaxis: { title: "Value" } },
            table: tableRows
          };
        }
      }
      return { data: [], layout: {}, table: [] };
    }

    // YTD Budget, YTD LY, YTD Act, YTD Gr, YTD Ach
    if (cleanTabLabel.startsWith("YTD")) {
      const label = cleanTabLabel.replace("YTD ", "");
      const ytdCols = cols.filter(col =>
        new RegExp(`ytd.*${label}|${label}.*ytd`, "i").test(col)
      );
      if (ytdCols.length) {
        let melted = melt(tableData, cols[0], ytdCols, "Period", label);
        melted = melted.filter(r => r.Period && !isNaN(r[label]));
        let grouped = {};
        melted.forEach(r => {
          const key = r.Period;
          if (!grouped[key]) grouped[key] = { Period: extractMonthYear(r.Period), [label]: 0 };
          grouped[key][label] += r[label];
        });
        const chartRows = Object.values(grouped);
        const tableRows = chartRows
          .filter(r => r.Period)
          .sort((a, b) => {
            const months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
            const [aMonth, aYear] = (a.Period || "").split("-");
            const [bMonth, bYear] = (b.Period || "").split("-");
            if (aYear !== bYear) return (aYear || "").localeCompare(bYear || "");
            return months.indexOf(aMonth) - months.indexOf(bMonth);
          });
        if (visType === "Pie Chart") {
          return {
            data: [{
              type: "pie",
              labels: tableRows.map(r => r.Period),
              values: tableRows.map(r => r[label]),
              textinfo: "percent+label",
              hoverinfo: "label+value+percent"
            }],
            layout: { title: `YTD ${label} Distribution` },
            table: tableRows
          };
        } else {
          return {
            data: [{
              type: visType === "Line Chart" ? "scatter" : "bar",
              x: tableRows.map(r => r.Period),
              y: tableRows.map(r => r[label]),
              marker: { color: label === "Act" ? "#FF8C00" : "#2E86AB" },
              ...(visType === "Line Chart" ? { mode: "lines+markers" } : {}),
              name: `YTD ${label}`
            }],
            layout: { title: `YTD ${label} by Period`, xaxis: { title: "Period" }, yaxis: { title: "Value" } },
            table: tableRows
          };
        }
      }
      return { data: [], layout: {}, table: [] };
    }

    // Branch Performance
    if (cleanTabLabel === "Branch Performance") {
      const ytdActCol = cols.find(col => /ytd.*act|act.*ytd/i.test(col));
      if (ytdActCol) {
        let filtered = tableData.filter(row => row[cols[0]] && row[ytdActCol]);
        // Exclude 'north total' and 'grand total' (case-insensitive) from branch names
        filtered = filtered.filter(row => {
          const val = row[cols[0]];
          if (!val || typeof val !== "string") return true;
          const lower = val.toLowerCase();
          return !(
            lower.includes("north total") ||
            lower.includes("grand total")
          );
        });
        let x = filtered.map(row => row[cols[0]]);
        let y = filtered.map(row => Number(row[ytdActCol]) || 0);
        const tableRows = x.map((branch, i) => ({
          Branch: branch,
          Performance: y[i]
        }));
        if (visType === "Pie Chart") {
          return {
            data: [{
              type: "pie",
              labels: x,
              values: y,
              textinfo: "percent+label",
              hoverinfo: "label+value+percent"
            }],
            layout: { title: "Branch Performance" },
            table: tableRows
          };
        } else {
          return {
            data: [{
              type: visType === "Line Chart" ? "scatter" : "bar",
              x,
              y,
              marker: { color: "#2E86AB" },
              ...(visType === "Line Chart" ? { mode: "lines+markers" } : {}),
              name: "Branch"
            }],
            layout: { title: "Branch Performance", xaxis: { title: "Branch" }, yaxis: { title: "Performance" } },
            table: tableRows
          };
        }
      }
      return { data: [], layout: {}, table: [] };
    }

    // Branch Monthwise
    if (cleanTabLabel === "Branch Monthwise") {
      const actCols = cols.filter(col => /^act(?!.*ytd)/i.test(col));
      if (actCols.length) {
        let melted = melt(tableData, cols[0], actCols, "Month", "Value");
        melted = melted.filter(r => r.Month && !isNaN(r.Value));
        let grouped = {};
        melted.forEach(r => {
          const key = r.Month;
          if (!grouped[key]) grouped[key] = { Month: extractMonthYear(r.Month), Value: 0 };
          grouped[key].Value += r.Value;
        });
        const chartRows = Object.values(grouped);
        const tableRows = chartRows
          .filter(r => r.Month)
          .sort((a, b) => {
            const months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
            const [aMonth, aYear] = (a.Month || "").split("-");
            const [bMonth, bYear] = (b.Month || "").split("-");
            if (aYear !== bYear) return (aYear || "").localeCompare(bYear || "");
            return months.indexOf(aMonth) - months.indexOf(bMonth);
          });
        if (visType === "Pie Chart") {
          return {
            data: [{
              type: "pie",
              labels: tableRows.map(r => r.Month),
              values: tableRows.map(r => r.Value),
              textinfo: "percent+label",
              hoverinfo: "label+value+percent"
            }],
            layout: { title: "Branch Monthwise" },
            table: tableRows
          };
        } else {
          return {
            data: [{
              type: visType === "Line Chart" ? "scatter" : "bar",
              x: tableRows.map(r => r.Month),
              y: tableRows.map(r => r.Value),
              marker: { color: "#2E86AB" },
              ...(visType === "Line Chart" ? { mode: "lines+markers" } : {}),
              name: "Branch Monthwise"
            }],
            layout: { title: "Branch Monthwise", xaxis: { title: "Month" }, yaxis: { title: "Value" } },
            table: tableRows
          };
        }
      }
      return { data: [], layout: {}, table: [] };
    }

    // Product Performance
    if (cleanTabLabel === "Product Performance") {
      const ytdActCol = cols.find(col => /ytd.*act|act.*ytd/i.test(col));
      if (ytdActCol) {
        let filtered = tableData.filter(row => row[cols[0]] && row[ytdActCol]);
        // Exclude 'north total' and 'grand total' (case-insensitive) from product names
        filtered = filtered.filter(row => {
          const val = row[cols[0]];
          if (!val || typeof val !== "string") return true;
          const lower = val.toLowerCase();
          return !(
            lower.includes("north total") ||
            lower.includes("grand total")
          );
        });
        let x = filtered.map(row => row[cols[0]]);
        let y = filtered.map(row => Number(row[ytdActCol]) || 0);
        const tableRows = x.map((product, i) => ({
          Product: product,
          Performance: y[i]
        }));
        if (visType === "Pie Chart") {
          return {
            data: [{
              type: "pie",
              labels: x,
              values: y,
              textinfo: "percent+label",
              hoverinfo: "label+value+percent"
            }],
            layout: { title: "Product Performance" },
            table: tableRows
          };
        } else {
          return {
            data: [{
              type: visType === "Line Chart" ? "scatter" : "bar",
              x,
              y,
              marker: { color: "#2E86AB" },
              ...(visType === "Line Chart" ? { mode: "lines+markers" } : {}),
              name: "Product"
            }],
            layout: { title: "Product Performance", xaxis: { title: "Product" }, yaxis: { title: "Performance" } },
            table: tableRows
          };
        }
      }
      return { data: [], layout: {}, table: [] };
    }

    // Product Monthwise
    if (cleanTabLabel === "Product Monthwise") {
      const actCols = cols.filter(col => /^act(?!.*ytd)/i.test(col));
      if (actCols.length) {
        let melted = melt(tableData, cols[0], actCols, "Month", "Value");
        melted = melted.filter(r => r.Month && !isNaN(r.Value));
        let grouped = {};
        melted.forEach(r => {
          const key = r.Month;
          if (!grouped[key]) grouped[key] = { Month: extractMonthYear(r.Month), Value: 0 };
          grouped[key].Value += r.Value;
        });
        const chartRows = Object.values(grouped);
        const tableRows = chartRows
          .filter(r => r.Month)
          .sort((a, b) => {
            const months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
            const [aMonth, aYear] = (a.Month || "").split("-");
            const [bMonth, bYear] = (b.Month || "").split("-");
            if (aYear !== bYear) return (aYear || "").localeCompare(bYear || "");
            return months.indexOf(aMonth) - months.indexOf(bMonth);
          });
        if (visType === "Pie Chart") {
          return {
            data: [{
              type: "pie",
              labels: tableRows.map(r => r.Month),
              values: tableRows.map(r => r.Value),
              textinfo: "percent+label",
              hoverinfo: "label+value+percent"
            }],
            layout: { title: "Product Monthwise" },
            table: tableRows
          };
        } else {
          return {
            data: [{
              type: visType === "Line Chart" ? "scatter" : "bar",
              x: tableRows.map(r => r.Month),
              y: tableRows.map(r => r.Value),
              marker: { color: "#2E86AB" },
              ...(visType === "Line Chart" ? { mode: "lines+markers" } : {}),
              name: "Product Monthwise"
            }],
            layout: { title: "Product Monthwise", xaxis: { title: "Month" }, yaxis: { title: "Value" } },
            table: tableRows
          };
        }
      }
      return { data: [], layout: {}, table: [] };
    }

    // Default fallback
    return { data: [], layout: {}, table: [] };
  }

  // --- UI ---
  return (
    <div className="w-screen min-h-screen min-w-screen bg-gray-50 py-0 px-0" style={{ overflowX: "hidden" }}>
      <div className="w-full min-h-screen min-w-screen flex flex-col md:flex-row bg-white rounded-none md:rounded-xl shadow-lg" style={{ minHeight: "100vh" }}>
        {/* Sidebar */}
        <aside
          className="w-full md:w-[23rem] bg-gradient-to-b from-blue-700 to-blue-500 text-white md:rounded-l-xl p-7 flex-shrink-0 md:mb-0 md:mr-10 shadow-lg"
          style={{
            position: "fixed",
            top: 0,
            left: 0,
            width: "23rem",
            height: "100vh",
            zIndex: 20,
            minHeight: "100vh",
            overflowY: "auto"
          }}
        >
          <h2 className="text-2xl font-bold mb-6 flex items-center gap-2">
            <span role="img" aria-label="chart">📊</span>
            Dashboard
          </h2>
          <div className="mb-6">
            <label className="block text-base font-medium mb-2">Upload Excel File</label>
            <div className="w-full bg-white border-2 border-dashed border-blue-400 rounded-lg p-4 flex flex-col items-center justify-center shadow-sm hover:border-blue-600 transition-all" style={{ minHeight: '110px' }}>
              <input
                type="file"
                accept=".xlsx"
                onChange={handleFileChange}
                ref={fileInputRef}
                className="block w-full text-base font-semibold text-blue-900 file:mr-4 file:py-2 file:px-4 file:rounded-full file:border-0 file:text-base file:font-semibold file:bg-white file:text-blue-700 hover:file:bg-blue-100 focus:outline-none py-2 px-3"
                style={{ cursor: 'pointer', minHeight: '44px' }}
              />
              <span className="mt-2 text-xs text-gray-500">Drag & drop or click to select an Excel (.xlsx) file</span>
            </div>
          </div>
          {sheetNames.length > 0 && (
            <div className="mb-4">
              <label className="block text-sm font-normal mb-2">📄 Select a Sheet</label>
              <div className="relative w-full">
                <select
                  value={selectedSheet}
                  onChange={handleSheetSelect}
                  className="w-full rounded border-gray-300 text-blue-900 focus:ring-blue-300 focus:border-blue-300 bg-white text-base font-semibold py-2 px-3"
                  style={{ minHeight: '44px', maxHeight: '44px', overflowY: 'auto' }}
                >
                  <option value="">📄 Select a Sheet</option>
                  {sheetNames.map((s) => (
                    <option key={s} value={s}>
                      {s}
                    </option>
                  ))}
                </select>
                <style>{`
                  select::-webkit-scrollbar {
                    width: 8px;
                  }
                  select option {
                    max-height: 160px;
                    overflow-y: auto;
                  }
                `}</style>
              </div>
            </div>
          )}
          {tableOptions.length > 0 && (
            <div className="mb-4">
              <label className="block text-sm font-normal mb-2">📌 Select Table</label>
              <div className="flex flex-col gap-2">
                {tableOptions.map((t) => {
                  let firstCol = t;
                  if (rawTables[selectedSheet] && rawTables[selectedSheet][t]) {
                    const tableData = rawTables[selectedSheet][t];
                    if (tableData.length > 0) {
                      const cols = Object.keys(tableData[0]);
                      if (cols.length > 0) firstCol = cols[0];
                    }
                  }
                  return (
                    <label key={t} className="flex items-center gap-2 cursor-pointer text-sm font-normal">
                      <input
                        type="radio"
                        name="table"
                        value={t}
                        checked={selectedTable === t}
                        onChange={handleTableSelect}
                        className="form-radio text-blue-600 focus:ring-blue-500 scale-125"
                        style={{ marginRight: '8px' }}
                      />
                      <span className="truncate" title={t}>{firstCol}</span>
                    </label>
                  );
                })}
              </div>
            </div>
          )}
          {selectedTable && (
            <div className="mb-4 space-y-3">
              <label className="block text-sm font-normal">📅 Filter by Month</label>
              <div className="relative w-full">
                <select
                  name="month"
                  value={filters.month}
                  onChange={handleFilterChange}
                  className="w-full rounded border-gray-300 text-blue-900 focus:ring-blue-300 focus:border-blue-300 bg-white text-base font-semibold py-2 px-3"
                  style={{ minHeight: '44px', maxHeight: '44px', overflowY: 'auto' }}
                >
                  {filterOptions.months.map((m) => (
                    <option key={m} value={m}>
                      {m}
                    </option>
                  ))}
                </select>
                <style>{`
                  select::-webkit-scrollbar {
                    width: 8px;
                  }
                  select option {
                    max-height: 160px;
                    overflow-y: auto;
                  }
                `}</style>
              </div>
              <label className="block text-sm font-normal">📆 Filter by Year</label>
              <div className="relative w-full">
                <select
                  name="year"
                  value={filters.year}
                  onChange={handleFilterChange}
                  className="w-full rounded border-gray-300 text-blue-900 focus:ring-blue-300 focus:border-blue-300 bg-white text-base font-semibold py-2 px-3"
                  style={{ minHeight: '44px', maxHeight: '44px', overflowY: 'auto' }}
                >
                  {filterOptions.years.map((y) => (
                    <option key={y} value={y}>
                      {y}
                    </option>
                  ))}
                </select>
                <style>{`
                  select::-webkit-scrollbar {
                    width: 8px;
                  }
                  select option {
                    max-height: 160px;
                    overflow-y: auto;
                  }
                `}</style>
              </div>
              {filterOptions.branches.length > 1 && (
                <>
                  <label className="block text-sm font-normal">🌍 Filter by Branch</label>
                  <div className="relative w-full">
                    <select
                      name="branch"
                      value={filters.branch}
                      onChange={handleFilterChange}
                      className="w-full rounded border-gray-300 text-blue-900 focus:ring-blue-300 focus:border-blue-300 bg-white text-base font-semibold py-2 px-3"
                      style={{ minHeight: '44px', maxHeight: '44px', overflowY: 'auto' }}
                    >
                      {filterOptions.branches.map((b) => (
                        <option key={b} value={b}>
                          {b}
                        </option>
                      ))}
                    </select>
                    <style>{`
                      select::-webkit-scrollbar {
                        width: 8px;
                      }
                      select option {
                        max-height: 160px;
                        overflow-y: auto;
                      }
                    `}</style>
                  </div>
                </>
              )}
              {filterOptions.products.length > 1 && (
                <>
                  <label className="block text-sm font-normal">📦 Filter by Product</label>
                  <div className="relative w-full">
                    <select
                      name="product"
                      value={filters.product}
                      onChange={handleFilterChange}
                      className="w-full rounded border-gray-300 text-blue-900 focus:ring-blue-300 focus:border-blue-300 bg-white text-base font-semibold py-2 px-3"
                      style={{ minHeight: '44px', maxHeight: '44px', overflowY: 'auto' }}
                    >
                      {filterOptions.products.map((p) => (
                        <option key={p} value={p}>
                          {p}
                        </option>
                      ))}
                    </select>
                    <style>{`
                      select::-webkit-scrollbar {
                        width: 8px;
                      }
                      select option {
                        max-height: 160px;
                        overflow-y: auto;
                      }
                    `}</style>
                  </div>
                </>
              )}
              <hr className="my-3 border-blue-200" />
              <label className="block text-sm font-normal mt-2">📊 Visualization Options</label>
              <div style={{ height: 12 }}></div>
              <label className="block text-xs font-medium mb-1 text-white">Select Visualization Type</label>
              <div style={{ marginBottom: '0.75rem' }}></div>
              <div className="relative w-full">
                <select
                  value={visualType}
                  onChange={handleVisualTypeChange}
                  className="w-full rounded border-gray-300 text-blue-900 focus:ring-blue-300 focus:border-blue-300 bg-white text-base font-semibold py-2 px-3"
                  style={{ minHeight: '44px', maxHeight: '44px', overflowY: 'auto' }}
                >
                  <option>Bar Chart</option>
                  <option>Pie Chart</option>
                  <option>Line Chart</option>
                </select>
                <style>{`
                  select::-webkit-scrollbar {
                    width: 8px;
                  }
                  select option {
                    max-height: 160px;
                    overflow-y: auto;
                  }
                `}</style>
              </div>
              {/* Master PPT Download Section */}
              <div className="mt-6 border-t border-blue-200 pt-4">
                <div className="font-medium text-base mb-2 flex items-center gap-2">
                  <span role="img" aria-label="ppt">📊</span> Download All Visuals
                </div>
                <button
                  onClick={handleGenerateMasterPPT}
                  className="w-full px-4 py-2 bg-orange-600 text-white rounded font-semibold hover:bg-orange-700 transition-all"
                  style={{ marginBottom: 8 }}
                >
                  🔄 Generate Master PPT (Charts Only)
                </button>
                <div className="text-xs text-white/80 mt-1">
                  All PPT downloads contain ONLY clean charts (no tables, straight x-axis labels)
                </div>
              </div>
            </div>
          )}
        </aside>
        {/* Main Content */}
        <div className="flex-1 min-h-screen bg-white" style={{ marginLeft: "23rem", minHeight: "100vh", height: "100vh", overflowY: "auto" }}>
          <div className="p-2 md:p-6" style={{ minHeight: "100vh" }}>
            {/* Heading above data table, always visible */}
            <div className="mb-6 flex items-center gap-2">
              <span role="img" aria-label="chart" className="text-2xl">📊</span>
              <span className="text-2xl font-bold">Excel Dashboard - Data Table & Visualizations</span>
            </div>
            {loading && (
              <div className="flex items-center gap-2 text-blue-600 font-medium mb-4">
                <svg className="animate-spin h-5 w-5 text-blue-600" viewBox="0 0 24 24">
                  <circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4" fill="none"/>
                  <path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8v8z"/>
                </svg>
                Loading...
              </div>
            )}
            {selectedTable && (
              <>
                {/* Only Data Table for 'sales analysis' sheet */}
                {selectedSheet.toLowerCase().includes("sales analysis") ? (
                  <div className="mb-8">
                    <h3 className="text-lg font-semibold mb-2">📋 Filtered Table View</h3>
                    <div className="overflow-x-auto rounded border border-gray-200 bg-gray-50" style={{ maxHeight: "400px", minHeight: "200px", overflowY: "auto", height: "400px" }}>
                      <table className="min-w-full text-sm text-gray-800">
                        <thead className="bg-blue-100 sticky top-0 z-10">
                          <tr>
                            {columns.map((col) => (
                              <th key={col} className="px-3 py-2 font-semibold text-left">
                                {col}
                              </th>
                            ))}
                          </tr>
                        </thead>
                        <tbody>
                          {dataTable.map((row, idx) => (
                            <tr key={idx} className={idx % 2 === 0 ? "bg-white" : "bg-gray-100"}>
                              {columns.map((col) => (
                                <td key={col} className="px-3 py-2">
                                  {row[col]}
                                </td>
                              ))}
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                    <button
                      onClick={handleDownloadCSV}
                      disabled={!csvData}
                      className="mt-3 px-3 py-1 text-sm bg-blue-600 text-white rounded hover:bg-blue-700 disabled:bg-gray-300 disabled:cursor-not-allowed font-semibold"
                      style={{ width: "auto", minWidth: "0" }}
                    >
                      ⬇️ Download CSV
                    </button>
                  </div>
                ) : (
                  <>
                    <div className="mb-8">
                      <h3 className="text-lg font-semibold mb-2">📋 Filtered Table View</h3>
                    <div>
                      <div
                        className="overflow-x-auto rounded border border-gray-200 bg-gray-50"
                        style={{
                          maxHeight: "400px",
                          minHeight: "200px",
                          overflowY: "auto",
                          width: "100%",
                          height: "400px"
                        }}
                      >
                        <table className="min-w-full text-sm text-gray-800" style={{ minWidth: "700px" }}>
                          <thead className="bg-blue-100 sticky top-0 z-10">
                            <tr>
                              {columns.map((col) => (
                                <th key={col} className="px-3 py-2 font-semibold text-left">
                                  {col}
                                </th>
                              ))}
                            </tr>
                          </thead>
                          <tbody>
                            {dataTable.map((row, idx) => (
                              <tr key={idx} className={idx % 2 === 0 ? "bg-white" : "bg-gray-100"}>
                                {columns.map((col) => (
                                  <td key={col} className="px-3 py-2">
                                    {typeof row[col] === "number" && !Number.isInteger(row[col])
                                      ? row[col].toFixed(2)
                                      : row[col]}
                                  </td>
                                ))}
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      </div>
                      <button
                        onClick={handleDownloadCSV}
                        disabled={!csvData}
                        className="mt-3 px-3 py-1 text-sm bg-blue-600 text-white rounded hover:bg-blue-700 disabled:bg-gray-300 disabled:cursor-not-allowed font-semibold"
                        style={{ width: "auto", minWidth: "0" }}
                      >
                        ⬇️ Download CSV
                      </button>
                    </div>
                    </div>
                    {/* Restore chart and view data table section */}
                    <div className="mb-8">
                      <div className="mb-4 overflow-x-auto custom-scrollbar" style={{ whiteSpace: 'nowrap' }}>
                        <div className="flex gap-3" style={{ minWidth: 'max-content' }}>
                          {tabNames.map((tab) => (
                            <button
                              key={tab}
                              className={`px-4 py-2 rounded-md text-base ${activeTab === tab ? "bg-blue-600 text-white" : "bg-gray-200 text-gray-800"} font-medium shadow-sm transition-all duration-150`}
                              onClick={() => {
                                setActiveTab(tab);
                                setShowBvATable(false);
                              }}
                              style={{ whiteSpace: 'nowrap', minWidth: '110px' }}
                            >
                              {tab}
                            </button>
                          ))}
                        </div>
                        <style>{`
                          .custom-scrollbar::-webkit-scrollbar {
                            height: 6px;
                          }
                          .custom-scrollbar::-webkit-scrollbar-thumb {
                            background: #b3c6e0;
                            border-radius: 4px;
                          }
                          .custom-scrollbar::-webkit-scrollbar-track {
                            background: #f0f4fa;
                          }
                        `}</style>
                      </div>
                      <h3 className="text-lg font-semibold mb-2">{activeTab} Visualization</h3>
                      <div className="bg-white rounded shadow p-2" style={{ width: "100%", minHeight: "320px" }}>
                        <Plot
                          data={getTabChartData(activeTab, dataTable, columns, filters, visualType).data}
                          layout={{ ...getTabChartData(activeTab, dataTable, columns, filters, visualType).layout, autosize: true }}
                          useResizeHandler
                          style={{ width: "100%", height: "min(60vw, 500px)", minHeight: "320px" }}
                          config={{
                            displayModeBar: true,
                            displaylogo: false,
                            responsive: true,
                          }}
                        />
                      </div>
                      {/* Data Table for all tabs */}
                      <div className="mt-4 flex flex-wrap gap-2 items-center">
                        <button
                          className="px-4 py-2 rounded bg-gray-200 text-gray-800 font-semibold hover:bg-blue-100"
                          onClick={() => setShowBvATable((v) => !v)}
                        >
                          {showBvATable ? "Hide" : "📊 View Data Table"}
                        </button>
                      </div>
                      {showBvATable && (
                        <div className="mt-2">
                          <div className="overflow-x-auto border rounded bg-gray-50" style={{ maxHeight: "300px", minHeight: "120px", overflowY: "auto", height: "300px" }}>
                            <table className="min-w-full text-sm text-gray-800">
                              <thead className="bg-blue-100 sticky top-0 z-10">
                                <tr>
                                  {getTabChartData(activeTab, dataTable, columns, filters, visualType).table &&
                                    getTabChartData(activeTab, dataTable, columns, filters, visualType).table.length > 0 &&
                                    Object.keys(getTabChartData(activeTab, dataTable, columns, filters, visualType).table[0]).map((col) => (
                                      <th key={col} className="px-3 py-2 font-semibold text-left">{col}</th>
                                    ))}
                                </tr>
                              </thead>
                              <tbody>
                                {getTabChartData(activeTab, dataTable, columns, filters, visualType).table &&
                                  getTabChartData(activeTab, dataTable, columns, filters, visualType).table.map(
                                    (row, idx) => (
                                      <tr key={idx} className={idx % 2 === 0 ? "bg-white" : "bg-gray-100"}>
                                        {Object.keys(row).map((col) => (
                                          <td key={col} className="px-3 py-2">
                                            {typeof row[col] === "number" && !Number.isInteger(row[col])
                                              ? row[col].toFixed(2)
                                              : row[col]}
                                          </td>
                                        ))}
                                      </tr>
                                    )
                                  )}
                              </tbody>
                            </table>
                          </div>
                        </div>
                      )}
                      {/* Download buttons below the table, with space */}
                      <div className="mt-6 flex flex-col gap-3 items-start">
                        <button
                          onClick={handleDownloadTabPPT}
                          className="px-4 py-2 text-sm bg-blue-600 text-white rounded-md hover:bg-blue-700 disabled:bg-gray-300 disabled:cursor-not-allowed font-semibold shadow"
                          style={{ width: "auto", minWidth: "0" }}
                          disabled={
                            !getTabChartData(activeTab, dataTable, columns, filters, visualType).data ||
                            getTabChartData(activeTab, dataTable, columns, filters, visualType).data.length === 0
                          }
                        >
                          ⬇️ Download {activeTab} PPT
                        </button>
                      </div>
                      {/* Metrics for Branch/Product Performance - now truly below the data table */}
                      <div style={{ marginTop: 32 }}></div>
                      {(activeTab === "Branch Performance" || activeTab === "Product Performance") && (() => {
                        const tabData = getTabChartData(activeTab, dataTable, columns, filters, visualType).table || [];
                        if (!tabData.length) return null;
                        // Use correct key for name
                        const nameKey = activeTab === "Branch Performance" ? "Branch" : "Product";
                        const perfKey = "Performance";
                        // Sort descending for top, ascending for bottom
                        const sorted = [...tabData].sort((a, b) => (b[perfKey] || 0) - (a[perfKey] || 0));
                        const topPerformer = sorted[0];
                        const totalPerformance = tabData.reduce((sum, r) => sum + (Number(r[perfKey]) || 0), 0);
                        const avgPerformance = tabData.length ? totalPerformance / tabData.length : 0;
                        const top5 = sorted.slice(0, 5);
                        const bottom5 = [...sorted].reverse().slice(0, 5);
                        return (
                          <div className="mb-6 flex flex-col items-center justify-center">
                            <div className="flex flex-wrap gap-6 mb-6 justify-center">
                              <div className="bg-green-100 rounded-xl p-4 min-w-[180px] flex flex-col items-center">
                                <div className="font-semibold text-green-800 text-xl mb-1">Top Performer</div>
                                <div className="text-xl font-bold mb-1">{topPerformer?.[nameKey]}</div>
                                <div className="text-green-700 text-lg">{topPerformer?.[perfKey]?.toLocaleString()}</div>
                              </div>
                              <div className="bg-blue-100 rounded-xl p-4 min-w-[180px] flex flex-col items-center">
                                <div className="font-semibold text-blue-800 text-xl mb-1">Total Performance</div>
                                <div className="text-xl font-bold text-blue-700 mb-1">{totalPerformance.toLocaleString()}</div>
                              </div>
                              <div className="bg-yellow-100 rounded-xl p-4 min-w-[180px] flex flex-col items-center">
                                <div className="font-semibold text-yellow-800 text-xl mb-1">Average Performance</div>
                                <div className="text-xl font-bold text-yellow-700 mb-1">{avgPerformance.toLocaleString(undefined, {maximumFractionDigits: 2})}</div>
                              </div>
                            </div>
                            <div className="flex flex-wrap gap-6 justify-center w-full">
                              <div className="flex-1 min-w-[260px] max-w-[340px] mx-auto">
                                <div className="font-semibold text-lg text-center mb-2">🏆 Top 5 {activeTab === "Branch Performance" ? "Regions" : "Products"}</div>
                                <table className="min-w-full text-base text-gray-800 border rounded-xl bg-white mx-auto">
                                  <thead className="bg-blue-50">
                                    <tr>
                                      <th className="px-3 py-2 text-center">{nameKey}</th>
                                      <th className="px-3 py-2 text-center">{perfKey}</th>
                                    </tr>
                                  </thead>
                                  <tbody>
                                    {top5.map((row, idx) => (
                                      <tr key={idx} className={idx % 2 === 0 ? "bg-white" : "bg-gray-50"}>
                                        <td className="px-3 py-2 text-center font-semibold">{row[nameKey]}</td>
                                        <td className="px-3 py-2 text-center">{row[perfKey]?.toLocaleString()}</td>
                                      </tr>
                                    ))}
                                  </tbody>
                                </table>
                              </div>
                              <div className="flex-1 min-w-[260px] max-w-[340px] mx-auto">
                                <div className="font-semibold text-lg text-center mb-2">📉 Bottom 5 {activeTab === "Branch Performance" ? "Regions" : "Products"}</div>
                                <table className="min-w-full text-base text-gray-800 border rounded-xl bg-white mx-auto">
                                  <thead className="bg-blue-50">
                                    <tr>
                                      <th className="px-3 py-2 text-center">{nameKey}</th>
                                      <th className="px-3 py-2 text-center">{perfKey}</th>
                                    </tr>
                                  </thead>
                                  <tbody>
                                    {bottom5.map((row, idx) => (
                                      <tr key={idx} className={idx % 2 === 0 ? "bg-white" : "bg-gray-50"}>
                                        <td className="px-3 py-2 text-center font-semibold">{row[nameKey]}</td>
                                        <td className="px-3 py-2 text-center">{row[perfKey]?.toLocaleString()}</td>
                                      </tr>
                                    ))}
                                  </tbody>
                                </table>
                              </div>
                            </div>
                          </div>
                        );
                      })()}

                      {/* Metrics for Branch/Product Monthwise */}
                      {(activeTab === "Branch Monthwise" || activeTab === "Product Monthwise") && (() => {
                        // Get the tab data (should be in long format: {Branch/Product, Month, Value})
                        const tabData = getTabChartData(activeTab, dataTable, columns, filters, visualType).table || [];
                        if (!tabData.length) return null;
                        // Group by month, sum values
                        const monthTotals = {};
                        tabData.forEach(row => {
                          const month = row["Month"];
                          const value = Number(row["Value"]) || 0;
                          if (!monthTotals[month]) monthTotals[month] = 0;
                          monthTotals[month] += value;
                        });
                        const months = Object.keys(monthTotals);
                        if (!months.length) return null;
                        // Best Month
                        let bestMonth = months[0];
                        let bestMonthValue = monthTotals[bestMonth];
                        months.forEach(m => {
                          if (monthTotals[m] > bestMonthValue) {
                            bestMonth = m;
                            bestMonthValue = monthTotals[m];
                          }
                        });
                        // Monthly Average
                        const monthlyAvg = months.length ? (Object.values(monthTotals).reduce((a, b) => a + b, 0) / months.length) : 0;
                        // Total Performance
                        const totalPerformance = Object.values(monthTotals).reduce((a, b) => a + b, 0);
                        return (
                          <div className="mb-6 flex flex-col items-center justify-center">
                            <div className="flex flex-wrap gap-6 mb-6 justify-center">
                              <div className="bg-green-100 rounded-xl p-4 min-w-[180px] flex flex-col items-center">
                                <div className="font-semibold text-green-800 text-xl mb-1">Best Month</div>
                                <div className="text-xl font-bold mb-1">{bestMonth}</div>
                                <div className="text-green-700 text-lg">{bestMonthValue.toLocaleString()}</div>
                              </div>
                              <div className="bg-yellow-100 rounded-xl p-4 min-w-[180px] flex flex-col items-center">
                                <div className="font-semibold text-yellow-800 text-xl mb-1">Monthly Average</div>
                                <div className="text-xl font-bold text-yellow-700 mb-1">{monthlyAvg.toLocaleString(undefined, {maximumFractionDigits: 2})}</div>
                              </div>
                              <div className="bg-blue-100 rounded-xl p-4 min-w-[180px] flex flex-col items-center">
                                <div className="font-semibold text-blue-800 text-xl mb-1">Total Performance</div>
                                <div className="text-xl font-bold text-blue-700 mb-1">{totalPerformance.toLocaleString()}</div>
                              </div>
                            </div>
                          </div>
                        );
                      })()}
                    </div>
                    {/* End restore */}
                  </>
                )}
              </>
            )}
          </div>
        </div>
      </div>
    </div>
  );
}

export default App;
