# Excel: matcha företagsnamn

Webbverktyg som läser två Excel-filer i webbläsaren (filerna lämnar inte din dator), hittar rader i fil B där företagsnamnet från lista A (text efter första ` - ` i kolumn A; prefixet får innehålla streck) förekommer i någon cell, och låter dig ladda ner en ny `.xlsx` med alla träffar.

## GitHub Pages

Efter första lyckade deploy blir appen tillgänglig på:

`https://<ditt-användarnamn>.github.io/excel-company-matcher/`

(byt om du valde annat repo-namn)

### Engångsinställning i GitHub

1. Repo → **Settings** → **Pages**
2. Under **Build and deployment**, **Source**: **GitHub Actions**

Därefter räcker det att pusha till `main`; workflow **Deploy to GitHub Pages** bygger och publicerar.

## Lokal användning

Öppna `index.html` i en webbläsare eller följ `INSTRUKTIONER.txt`.

## Villkor (SheetJS)

`vendor/xlsx.full.min.js` är SheetJS Community Edition — följ licensen i deras dokumentation om du redistribuerar.
