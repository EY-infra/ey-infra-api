const express = require('express');
const { ClientSecretCredential } = require('@azure/identity');
const { Client } = require('@microsoft/microsoft-graph-client');
const { TokenCredentialAuthenticationProvider } = require('@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials');

const app = express();
app.use(express.json());

app.use((req, res, next) => {
  res.header('Access-Control-Allow-Origin', '*');
  res.header('Access-Control-Allow-Headers', 'Content-Type');
  res.header('Access-Control-Allow-Methods', 'POST, GET, OPTIONS, DELETE');
  if (req.method === 'OPTIONS') return res.sendStatus(200);
  next();
});

const credential = new ClientSecretCredential(
  process.env.TENANT_ID,
  process.env.CLIENT_ID,
  process.env.CLIENT_SECRET
);

const AFZENDER = process.env.MAIL_AFZENDER || 'administratie@eyinfrasupport.nl';

// ── GEDEELDE OPSLAG (in geheugen — blijft zolang server draait) ──
let gedeeldeData = {
  projecten: ['Intern'],
  aanvragen: [],    // vakantieaanvragen voor manager overzicht
  vragen: [],       // vragen voor manager overzicht
  documenten: []    // documenten voor manager overzicht
};

// ── HOOFDPAGINA ──
app.get('/', (req, res) => {
  res.send(`<!DOCTYPE html><html lang="nl"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>EY Infra Support API</title>
  <style>body{font-family:Arial,sans-serif;background:#f0f4fb;display:flex;align-items:center;justify-content:center;min-height:100vh;margin:0}
  .box{background:#fff;border-radius:16px;padding:32px 40px;max-width:440px;width:100%;box-shadow:0 4px 20px rgba(26,42,94,.1);text-align:center}
  h1{color:#1a2a5e;font-size:20px;margin:0 0 4px}.sub{color:#7a90b8;font-size:13px;margin-bottom:24px}
  .ok{background:#e6f9f0;border:1px solid #a8e6c8;border-radius:10px;padding:12px 16px;color:#1a6b44;font-weight:bold}
  .ep{background:#f0f4fb;border-radius:8px;padding:8px 12px;font-family:monospace;font-size:12px;color:#1a2a5e;margin-top:10px;text-align:left}
  .tijd{font-size:11px;color:#aaa;margin-top:16px}</style></head>
  <body><div class="box"><h1>⚡ EY Infra Support</h1><div class="sub">HR Portaal — Microsoft 365 API</div>
  <div class="ok">✅ Server is online</div>
  <div class="ep">GET  /status<br>GET  /projecten<br>POST /projecten<br>POST /send-mail</div>
  <div class="tijd">⏱ ${new Date().toLocaleString('nl-NL')}</div></div></body></html>`);
});

// ── STATUS ──
app.get('/status', (req, res) => {
  res.json({ status: 'online', naam: 'EY Infra Support — HR Portaal API', tijd: new Date().toLocaleString('nl-NL'), projecten: gedeeldeData.projecten.length });
});

// ── PROJECTEN OPHALEN ──
app.get('/projecten', (req, res) => {
  res.json({ projecten: gedeeldeData.projecten });
});

// ── PROJECTEN OPSLAAN ──
app.post('/projecten', (req, res) => {
  const { projecten } = req.body;
  if (!Array.isArray(projecten)) return res.status(400).json({ error: 'projecten moet een lijst zijn' });
  gedeeldeData.projecten = projecten;
  console.log(`📋 Projecten bijgewerkt: ${projecten.join(', ')}`);
  res.json({ success: true, projecten });
});

// ── MAIL VERSTUREN ──
app.post('/send-mail', async (req, res) => {
  const { aan, onderwerp, tekst, htmlTekst } = req.body;
  if (!aan || !onderwerp || !tekst) return res.status(400).json({ error: 'Velden aan, onderwerp en tekst zijn verplicht.' });

  try {
    const authProvider = new TokenCredentialAuthenticationProvider(credential, {
      scopes: ['https://graph.microsoft.com/.default']
    });
    const client = Client.initWithMiddleware({ authProvider });
    await client.api(`/users/${AFZENDER}/sendMail`).post({
      message: {
        subject: onderwerp,
        body: { contentType: htmlTekst ? 'HTML' : 'Text', content: htmlTekst || tekst },
        toRecipients: [{ emailAddress: { address: aan } }],
        from: { emailAddress: { name: 'HR Portaal — EY Infra Support', address: AFZENDER } }
      },
      saveToSentItems: true
    });
    console.log(`✅ Mail verstuurd naar ${aan}: ${onderwerp}`);
    res.json({ success: true });
  } catch (err) {
    console.error('❌ Mail fout:', err.message);
    res.status(500).json({ error: err.message });
  }
});

const PORT = process.env.PORT || 3001;
app.listen(PORT, () => {
  console.log(`🚀 EY Infra API draait op poort ${PORT}`);
  console.log(`📧 Afzender: ${AFZENDER}`);
});

const app = express();
app.use(express.json());

// CORS — zodat Netlify de server mag aanroepen
app.use((req, res, next) => {
  res.header('Access-Control-Allow-Origin', '*');
  res.header('Access-Control-Allow-Headers', 'Content-Type');
  res.header('Access-Control-Allow-Methods', 'POST, GET, OPTIONS');
  if (req.method === 'OPTIONS') return res.sendStatus(200);
  next();
});

// Verbinding met Azure via omgevingsvariabelen
const credential = new ClientSecretCredential(
  process.env.TENANT_ID,
  process.env.CLIENT_ID,
  process.env.CLIENT_SECRET
);

// Afzender — het Outlook-account dat de mails verstuurt
const AFZENDER = process.env.MAIL_AFZENDER || 'ahmet.erol@eyinfrasupport.nl';

// ── MAIL VERSTUREN ──
app.post('/send-mail', async (req, res) => {
  const { aan, onderwerp, tekst, htmlTekst } = req.body;

  if (!aan || !onderwerp || !tekst) {
    return res.status(400).json({ error: 'Velden aan, onderwerp en tekst zijn verplicht.' });
  }

  try {
    const authProvider = new TokenCredentialAuthenticationProvider(credential, {
      scopes: ['https://graph.microsoft.com/.default']
    });

    const client = Client.initWithMiddleware({ authProvider });

    await client.api(`/users/${AFZENDER}/sendMail`).post({
      message: {
        subject: onderwerp,
        body: {
          contentType: htmlTekst ? 'HTML' : 'Text',
          content: htmlTekst || tekst
        },
        toRecipients: [
          { emailAddress: { address: aan } }
        ],
        from: {
          emailAddress: {
            name: 'HR Portaal — EY Infra Support',
            address: AFZENDER
          }
        }
      },
      saveToSentItems: true
    });

    console.log(`✅ Mail verstuurd naar ${aan}: ${onderwerp}`);
    res.json({ success: true, bericht: `Mail verstuurd naar ${aan}` });

  } catch (err) {
    console.error('❌ Fout bij versturen mail:', err.message);
    res.status(500).json({ error: err.message });
  }
});

// ── HOOFDPAGINA (status dashboard) ──
app.get('/', (req, res) => {
  res.send(`
    <!DOCTYPE html>
    <html lang="nl">
    <head>
      <meta charset="UTF-8">
      <meta name="viewport" content="width=device-width, initial-scale=1">
      <title>EY Infra Support — HR API</title>
      <style>
        body { font-family: Arial, sans-serif; background: #f0f4fb; display: flex; align-items: center; justify-content: center; min-height: 100vh; margin: 0; }
        .box { background: #fff; border-radius: 16px; padding: 32px 40px; max-width: 440px; width: 100%; box-shadow: 0 4px 20px rgba(26,42,94,.1); text-align: center; }
        .bolt { background: #1a2a5e; width: 56px; height: 56px; border-radius: 14px; display: flex; align-items: center; justify-content: center; margin: 0 auto 16px; font-size: 28px; }
        h1 { color: #1a2a5e; font-size: 20px; margin: 0 0 4px; }
        .sub { color: #7a90b8; font-size: 13px; margin-bottom: 24px; }
        .status { background: #e6f9f0; border: 1px solid #a8e6c8; border-radius: 10px; padding: 12px 16px; margin-bottom: 16px; }
        .status-dot { display: inline-block; width: 10px; height: 10px; background: #1a6b44; border-radius: 50%; margin-right: 8px; }
        .status span { font-weight: bold; color: #1a6b44; font-size: 14px; }
        .info { font-size: 12px; color: #7a90b8; line-height: 1.6; }
        .endpoint { background: #f0f4fb; border-radius: 8px; padding: 8px 12px; font-family: monospace; font-size: 12px; color: #1a2a5e; margin-top: 12px; text-align: left; }
        .tijd { font-size: 11px; color: #aaa; margin-top: 16px; }
      </style>
    </head>
    <body>
      <div class="box">
        <div class="bolt">⚡</div>
        <h1>EY Infra Support</h1>
        <div class="sub">HR Portaal — Microsoft 365 API</div>
        <div class="status">
          <span class="status-dot"></span>
          <span>Server is online en actief</span>
        </div>
        <div class="endpoint">POST /send-mail — mail versturen</div>
        <div class="endpoint">GET &nbsp;/status &nbsp;&nbsp;&nbsp;— status controleren</div>
        <div class="tijd">⏱ ${new Date().toLocaleString('nl-NL')}</div>
      </div>
    </body>
    </html>
  `);
});

// ── STATUS CHECK ──
app.get('/status', (req, res) => {
  res.json({
    status: 'online',
    naam: 'EY Infra Support — HR Portaal API',
    afzender: AFZENDER,
    tijd: new Date().toLocaleString('nl-NL')
  });
});

const PORT = process.env.PORT || 3001;
app.listen(PORT, () => {
  console.log(`🚀 EY Infra API draait op poort ${PORT}`);
  console.log(`📧 Afzender: ${AFZENDER}`);
});
