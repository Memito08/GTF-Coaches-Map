# GTF Coaches Map

Interactive world map showing how many participants in the Global Talent Fund
GTF Coaches program come from each country.

GTF Coaches is an online training initiative focused on Olympiad mathematics
problem-solving and teaching methods. It is designed for coaches and organizers
of Olympiad mathematics competitions, helping them strengthen mathematical,
leadership, and organizational skills.

## Overview

This is a static single-page app built in `index.html`. It uses OpenLayers for
the interactive map, Proj4js for a Robinson projection, and SheetJS to read the
local Excel data file in the browser.

The app loads:

- `data/GTF Coaches Data.xlsx`: coach counts by country.
- `data/country_codes.json`: canonical country labels.
- `data/geometry.json`: country polygon geometry.

The spreadsheet data is matched to the canonical country labels, aggregated by
country, and rendered as a green choropleth map. Countries with no data or zero
coaches are shown in white. Tooltips show each country's coach count on hover or
tap.

## Running Locally

Because the app fetches local data files, serve it over HTTP:

```bash
python -m http.server 8000
```

Then open:

```text
http://localhost:8000/
```

## Updating Data

To update the map, replace or edit `data/GTF Coaches Data.xlsx`, then check that
country names still match the map labels:

```bash
python data/check_matches.py
```

If needed, update the `COUNTRY_ALIASES` list in both `index.html` and
`data/check_matches.py`.

## Notes

Map boundaries come from the IMF DataMapper source referenced in the in-app
disclaimer. They are used for visualization only and do not imply any position
on the legal status of any territory.
