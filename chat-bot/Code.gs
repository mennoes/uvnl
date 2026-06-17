// ──────────────────────────────────────────────────────────────
// UvNL Titelbalk — Google Chat bot
//
// Werkwijze:
//  1. Gebruiker stuurt bv. "Deze truc laat je 20% meer onthouden 9:16"
//     of "Even denken! 4:5 caps".
//  2. Bot parsed tekst + formaat + style.
//  3. Bot bouwt een headless URL naar de GitHub-Pages generator:
//     https://mennoes.github.io/uvnl/uvnl-titelbalk.html
//       ?headless=1&t=...&format=916&style=bold
//  4. Bot vraagt screenshotone.com (gratis tier, 100/mnd) om een PNG
//     van die URL met transparante achtergrond op exacte format-maat.
//  5. PNG wordt als image-widget in een Chat-card teruggestuurd.
//
// Eenmalige setup:
//   - Script Properties → voeg key SCREENSHOTONE_KEY toe met je
//     screenshotone.com access key.
//   - Optioneel: GENERATOR_URL overrulen als je een eigen mirror host.
// ──────────────────────────────────────────────────────────────

const DEFAULT_GENERATOR_URL = 'https://mennoes.github.io/uvnl/uvnl-titelbalk.html';

const FORMATS = {
  '9:16': { id: '916', w: 1080, h: 1920 },
  '4:5':  { id: '45',  w: 1080, h: 1350 },
  '16:9': { id: '169', w: 1920, h: 1080 },
  '1:1':  { id: '11',  w: 1080, h: 1080 }
};

// ── Chat events ──────────────────────────────────────────────
function onMessage(event) {
  try {
    const raw = (event.message && event.message.argumentText
      ? event.message.argumentText
      : event.message.text || '').trim();

    if (!raw) {
      return textReply(helpText());
    }
    if (/^(help|hi|hoi|hallo|\?)$/i.test(raw)) {
      return textReply(helpText());
    }

    const parsed = parseInput(raw);
    if (!parsed.text) {
      return textReply('Geen titeltekst gevonden. ' + helpText());
    }

    const fmt = FORMATS[parsed.format];
    const targetUrl = buildHeadlessUrl(parsed.text, fmt.id, parsed.style);
    const imageUrl  = renderViaScreenshotOne(targetUrl, fmt.w, fmt.h);

    return cardReply(parsed, imageUrl, targetUrl);
  } catch (err) {
    console.error(err);
    return textReply('🚧 Er ging iets mis bij het renderen: ' + (err.message || err));
  }
}

function onAddToSpace(event) {
  return textReply(helpText());
}

function onRemoveFromSpace(event) {
  // niets te doen
}

// ── Input parsing ────────────────────────────────────────────
// Accepteert (in willekeurige volgorde, aan einde van bericht):
//   - formaat: 9:16 | 4:5 | 16:9 | 1:1 (default 9:16)
//   - style:   "caps" of "condensed" → condensed FULL CAPS
//              standaard: bold (Oldschool Grotesk Bold, mixed case)
function parseInput(raw) {
  const tokens = raw.split(/\s+/);
  let format = '9:16';
  let style  = 'bold';

  // Loop van achter naar voor en pak format/style-tokens af.
  while (tokens.length) {
    const last = tokens[tokens.length - 1].toLowerCase();
    if (FORMATS[last]) {
      format = last;
      tokens.pop();
      continue;
    }
    if (last === 'caps' || last === 'condensed') {
      style = 'condensed';
      tokens.pop();
      continue;
    }
    if (last === 'bold') {
      style = 'bold';
      tokens.pop();
      continue;
    }
    break;
  }

  return { text: tokens.join(' ').trim(), format: format, style: style };
}

// ── URL voor headless render ─────────────────────────────────
function buildHeadlessUrl(text, fmtId, style) {
  const base = scriptProp('GENERATOR_URL') || DEFAULT_GENERATOR_URL;
  const params = [
    'headless=1',
    'format=' + fmtId,
    'style=' + encodeURIComponent(style),
    't=' + encodeURIComponent(text)
  ].join('&');
  return base + '?' + params;
}

// ── Screenshotone.com render ─────────────────────────────────
// API docs: https://screenshotone.com/docs/
// Free tier: 100 renders/maand, geen credit card. Maak een account
// op https://screenshotone.com en zet de access key in Script
// Properties onder SCREENSHOTONE_KEY.
function renderViaScreenshotOne(targetUrl, w, h) {
  const key = scriptProp('SCREENSHOTONE_KEY');
  if (!key) {
    throw new Error('SCREENSHOTONE_KEY ontbreekt in Script Properties.');
  }
  const params = {
    access_key: key,
    url: targetUrl,
    viewport_width: String(w),
    viewport_height: String(h),
    device_scale_factor: '1',
    format: 'png',
    omit_background: 'true',
    block_ads: 'true',
    block_cookie_banners: 'true',
    wait_until: 'network_idle',
    delay: '1',
    cache: 'true',
    cache_ttl: '3600',
    response_type: 'by_format'
  };
  const qs = Object.keys(params)
    .map(k => k + '=' + encodeURIComponent(params[k]))
    .join('&');
  const apiUrl = 'https://api.screenshotone.com/take?' + qs;

  // Screenshotone met response_type=by_format levert direct image/png
  // bytes. We fetchen 'm in Apps Script en uploaden naar Drive zodat
  // Chat de afbeelding direct kan tonen zonder op een live render te
  // wachten. Drive-link is publiek voor "ANYONE_WITH_LINK".
  const resp = UrlFetchApp.fetch(apiUrl, { muteHttpExceptions: true });
  const code = resp.getResponseCode();
  if (code >= 400) {
    throw new Error('screenshotone HTTP ' + code + ' — ' + resp.getContentText().slice(0, 200));
  }
  const blob = resp.getBlob().setContentType('image/png')
    .setName('uvnl-titelbalk-' + Date.now() + '.png');
  const folder = getOrCreateBotFolder();
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return 'https://drive.google.com/uc?export=view&id=' + file.getId();
}

function getOrCreateBotFolder() {
  const name = 'UvNL Titelbalk Bot — renders';
  const it = DriveApp.getFoldersByName(name);
  return it.hasNext() ? it.next() : DriveApp.createFolder(name);
}

// ── Chat-replies ─────────────────────────────────────────────
function textReply(text) {
  return { text: text };
}

function cardReply(parsed, imageUrl, sourceUrl) {
  return {
    cardsV2: [{
      cardId: 'titelbalk-' + Date.now(),
      card: {
        header: {
          title: parsed.text,
          subtitle: parsed.format + ' • ' + (parsed.style === 'condensed' ? 'Condensed CAPS' : 'Oldschool Bold')
        },
        sections: [{
          widgets: [
            { image: { imageUrl: imageUrl, altText: parsed.text } },
            { buttonList: { buttons: [
              { text: 'Open in generator',
                onClick: { openLink: { url: sourceUrl.replace('&headless=1', '').replace('headless=1&', '') } } },
              { text: 'Download PNG',
                onClick: { openLink: { url: imageUrl } } }
            ] } }
          ]
        }]
      }
    }]
  };
}

function helpText() {
  return [
    '*UvNL Titelbalk bot* — typ een titel en (optioneel) een formaat.',
    '',
    'Voorbeelden:',
    '• `Deze truc laat je 20% meer onthouden`  → 9:16 standaard',
    '• `Even denken! 4:5`',
    '• `Wat slaap doet 16:9 caps`  → Condensed FULL CAPS',
    '',
    'Formaten: 9:16, 4:5, 16:9, 1:1',
    'Style: voeg `caps` of `condensed` toe voor de Condensed-variant; default is Oldschool Grotesk Bold.'
  ].join('\n');
}

// ── Helpers ──────────────────────────────────────────────────
function scriptProp(key) {
  return PropertiesService.getScriptProperties().getProperty(key);
}
