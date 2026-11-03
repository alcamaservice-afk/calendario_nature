# Note sulle icone

Le icone SVG in questa cartella (`icon-192.svg`, `icon-512.svg`) sono le icone sorgente dell'app.

## Come usarle per la PWA su GitHub Pages

Se usi GitHub Pages per ospitare il manifest/icone:
1. Converti i file SVG in PNG (puoi usare [svgToPng online](https://svgtopng.com/))
2. Salva come `icon-192.png` (192×192 px) e `icon-512.png` (512×512 px)
3. Aggiorna il `manifest.json` con il percorso corretto delle icone

## Come usarle su Google Apps Script

Google Apps Script non serve file binari. Il modo corretto per il logo PWA è:
1. Carica `icon-192.png` e `icon-512.png` su GitHub Pages o un CDN pubblico
2. Aggiorna il tag `<link rel="apple-touch-icon">` nel `Login.html` con l'URL pubblico delle PNG
3. Aggiorna il `manifest.json` con gli URL pubblici

## Nota per la scheda di Chrome

Su Google Chrome Desktop/Android, il logo che appare nella scheda del browser
e nell'icona di installazione è quello definito nel `<link rel="icon">`.
GAS usa automaticamente un'icona generica, ma puoi personalizzarla aggiungendo
nel `<head>` del `Login.html`:

```html
<link rel="icon" type="image/png" href="URL_PUBBLICA_ICON_192.png">
<link rel="shortcut icon" href="URL_PUBBLICA_ICON_192.png">
```
