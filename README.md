# Excel Dashboard - Data Table & Visualizations

This project provides an interactive dashboard for analyzing Excel files containing sales and performance data. It enables users to upload Excel files, select sheets and tables, apply filters, and visualize data through various charts and tables. The dashboard also supports exporting data and charts to CSV and PowerPoint (PPTX) formats.

## Features

- Upload Excel files and select sheets/tables for analysis
- Filter data by month, year, branch, or product
- View data tables with formatting and download as CSV
- Visualize data using Bar, Line, and Pie charts (Plotly)
- Download individual or master PowerPoint presentations with charts
- Metrics and highlights for top/bottom performers
- Responsive UI with tabs for different analyses

## Usage

1. **Install dependencies:**
   - Python 3.8+
   - Install required packages:
     ```
     pip install streamlit pandas numpy matplotlib seaborn plotly python-pptx openpyxl
     ```

2. **Run the dashboard:**
   ```
   streamlit run Dashboard(1).py
   ```

3. **Interact with the dashboard:**
   - Upload your Excel file
   - Select the desired sheet and table
   - Apply filters as needed
   - Explore tables and charts in different tabs
   - Download CSV or PPTX files as required

## Notes

- The dashboard is optimized for sales and performance data with specific table structures.
- For best results, use Excel files with clear headers and consistent formatting.
- All chart exports use clean, presentation-ready visuals.

---

# React + Vite

This template provides a minimal setup to get React working in Vite with HMR and some ESLint rules.

Currently, two official plugins are available:

- [@vitejs/plugin-react](https://github.com/vitejs/vite-plugin-react/blob/main/packages/plugin-react) uses [Babel](https://babeljs.io/) for Fast Refresh
- [@vitejs/plugin-react-swc](https://github.com/vitejs/vite-plugin-react/blob/main/packages/plugin-react-swc) uses [SWC](https://swc.rs/) for Fast Refresh

## Expanding the ESLint configuration

If you are developing a production application, we recommend using TypeScript with type-aware lint rules enabled. Check out the [TS template](https://github.com/vitejs/vite/tree/main/packages/create-vite/template-react-ts) for information on how to integrate TypeScript and [`typescript-eslint`](https://typescript-eslint.io) in your project.
