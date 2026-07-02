# PAGASA 5-Day Weather Tool

Run `Run PAGASA Weather Tool.bat`. On first launch it creates a local Python environment and installs the required packages. The browser opens at `http://127.0.0.1:8877`.

The app refreshes the Extended Weather Outlook for Selected Cities from PAGASA, arranges it in the same Region / Site / five-day structure as the supplied workbook, and exports the current view to Excel.

Forecast data refreshes automatically every 30 minutes while the page is open. The header displays a countdown to the next refresh, and the refresh button remains available for an immediate update.

Some workbook sites are not separate PAGASA selected-city entries. Those rows use the nearest available PAGASA city outlook and display that mapping in the interface.

## Vercel

This repository includes a Vercel Python entry point and routing configuration. Import the GitHub repository as a new Vercel project; no environment variables are required.
