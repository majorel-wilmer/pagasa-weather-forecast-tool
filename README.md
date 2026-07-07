# 5-Day Weather Tool (Open-Meteo)

Run `Run PAGASA Weather Tool.bat`. On first launch it creates a local Python environment and installs the required packages. The browser opens at `http://127.0.0.1:8877`.

The 5-day forecast matrix, rain timing and rain-intensity classification are all computed from **Open-Meteo** (https://open-meteo.com/), queried per site from each site's own coordinates — no site borrows another city's numbers. For every day, the app scans Open-Meteo's hourly precipitation data to find the most likely contiguous rain period and shows it as an exact clock window (for example, "Rain likely 2:00 PM – 5:00 PM") instead of a vague part of the day. Rain intensity (green/yellow/orange) is classified from modeled hourly precipitation rate (light <2.5 mm/hr, moderate 2.5–7.5 mm/hr, heavy >7.5 mm/hr), with thunderstorm codes escalating the rating once rain is measurable.

Historical comparisons (the Historical tab) also use Open-Meteo's Historical Weather API (ERA5 reanalysis).

Tropical cyclone/typhoon status on the Live Storm Map is the one exception: Open-Meteo does not publish cyclone advisories, so that page still reads PAGASA's official daily bulletin, which is the authoritative source for named-storm tracking in the Philippines.

Forecast data refreshes automatically every 30 minutes while the page is open. The header displays a countdown to the next refresh, and the refresh button remains available for an immediate update.

Open-Meteo data is licensed under CC BY 4.0 and requires no API key for non-commercial use.

## Vercel

This repository includes a Vercel Python entry point and routing configuration. Import the GitHub repository as a new Vercel project; no environment variables are required.
