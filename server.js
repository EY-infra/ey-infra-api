const express = require('express');
const { ClientSecretCredential } = require('@azure/identity');
const { Client } = require('@microsoft/microsoft-graph-client');
const { TokenCredentialAuthenticationProvider } = require('@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials');

const app = express();
app.use(express.json());

// CORS
app.use((req, res, next) => {
  res.header('Access-Control-Allow-Origin', '*');
  res.header('Access-Control-Allow-Headers', 'Content-Type');
  res.header('Access-Control-Allow-Methods', 'POST, GET, OPTIONS, DELETE');
  if (req.method === 'OPTIONS') return res.sendStatus(200);
  next();
});

// Azure verbinding
const credential = new ClientSecretCredential(
  process.env.TENANT_ID,
  process.env.CLIENT_ID,
  process.env.CLIENT_SECRET
);

const AFZENDER = process.env.MAIL_AFZENDER || 'administratie@eyinfrasupport.nl';

// Gedeelde opslag (in geheugen)
let gedeeldeData = {
  projecten: ['Heerle', 'W-Papendrecht', 'W-Dintelweg', 'W-Vijfhuizen', 'Infra overige', 'Intern']
};

// ── HOOFDPAGINA ──
app.get('/', (req, res) => {
  res.send(`<!DOCTYPE html>
<html lang="nl">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width,initial-scale=1">
  <title>EY Infra Support API</title>
  <style>
    body{font-family:Arial,sans-serif;background:#f0f4fb;display:flex;align-items:center;justify-content:center;min-height:100vh;margin:0}
    .box{background:#fff;border-radius:16px;padding:32px 40px;max-width:440px;width:100%;box-shadow:0 4px 20px rgba(26,42,94,.1);text-align:center}
    h1{color:#1a2a5e;font-size:20px;margin:0 0 4px}
    .sub{color:#7a90b8;font-size:13px;margin-bottom:20px}
    .ok{background:#e6f9f0;border:1px solid #a8e6c8;border-radius:10px;padding:12px 16px;color:#1a6b44;font-weight:bold;margin-bottom:12px}
    .ep{background:#f0f4fb;border-radius:8px;padding:8px 12px;font-family:monospace;font-size:12px;color:#1a2a5e;margin-top:8px;text-align:left}
    .tijd{font-size:11px;color:#aaa;margin-top:14px}
  </style>
</head>
<body>
  <div class="box">
    <h1>⚡ EY Infra Support</h1>
    <div class="sub">HR Portaal — Microsoft 365 API</div>
    <div class="ok">✅ Server is online en actief</div>
    <div class="ep">GET  /status<br>GET  /projecten<br>POST /projecten<br>POST /send-mail</div>
    <div class="tijd">⏱ ${new Date().toLocaleString('nl-NL')}</div>
  </div>
</body>
</html>`);
});

// ── STATUS ──
app.get('/status', (req, res) => {
  res.json({
    status: 'online',
    naam: 'EY Infra Support — HR Portaal API',
    afzender: AFZENDER,
    tijd: new Date().toLocaleString('nl-NL'),
    projecten: gedeeldeData.projecten.length
  });
});

// ── PROJECTEN OPHALEN ──
app.get('/projecten', (req, res) => {
  res.json({ projecten: gedeeldeData.projecten });
});

// ── PROJECTEN OPSLAAN ──
app.post('/projecten', (req, res) => {
  const { projecten } = req.body;
  if (!Array.isArray(projecten)) {
    return res.status(400).json({ error: 'projecten moet een lijst zijn' });
  }
  gedeeldeData.projecten = projecten;
  console.log(`📋 Projecten bijgewerkt: ${projecten.join(', ')}`);
  res.json({ success: true, projecten });
});

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
        toRecipients: [{ emailAddress: { address: aan } }],
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
    res.json({ success: true });

  } catch (err) {
    console.error('❌ Mail fout:', err.message);
    res.status(500).json({ error: err.message });
  }
});

const PORT = process.env.PORT || 3001;
app.listen(PORT, () => {
  console.log(`🚀 EY Infra Support API draait op poort ${PORT}`);
  console.log(`📧 Afzender: ${AFZENDER}`);
});
