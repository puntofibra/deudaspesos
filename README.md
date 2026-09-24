# DeudasPesos

Control de préstamos y deudas en **pesos colombianos (COP)**, con **interés mensual opcional**.
Web app de Google Apps Script + Google Sheets, publicada como PWA instalable en GitHub Pages.

## Cómo se calcula el interés
- Interés simple **por mes cumplido** desde la fecha de la deuda, sobre el **capital pendiente**.
- Los pagos (parciales o totales) cubren **primero los intereses** y después el capital.
- Todo se redondea a pesos enteros y se recalcula desde la fecha y la lista de pagos.

## Archivos
- `index.html`, `manifest.webmanifest`, `sw.js`, iconos: la app instalable (carga la web app de Apps Script).
- `apps-script/`: copia de respaldo del código del proyecto de Apps Script
  (en `Llave.gs` el secreto de la Central de Llaves está quitado a propósito).
