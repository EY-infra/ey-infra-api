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
  res.header('Access-Control-Allow-Methods', 'POST, GET, OPTIONS, DELETE, PUT');
  if (req.method === 'OPTIONS') return res.sendStatus(200);
  next();
});

// Azure
const credential = new ClientSecretCredential(
  process.env.TENANT_ID,
  process.env.CLIENT_ID,
  process.env.CLIENT_SECRET
);
const AFZENDER = process.env.MAIL_AFZENDER || 'administratie@eyinfrasupport.nl';

// Supabase
const SUPABASE_URL  = process.env.SUPABASE_URL;
const SUPABASE_KEY  = process.env.SUPABASE_KEY;

// Hulpfunctie voor Supabase REST API
async function sbFetch(path, method='GET', body=null) {
  if (!SUPABASE_URL || !SUPABASE_KEY) throw new Error('Supabase niet geconfigureerd');
  const opts = {
    method,
    headers: {
      'Content-Type': 'application/json',
      'apikey': SUPABASE_KEY,
      'Authorization': `Bearer ${SUPABASE_KEY}`,
      'Prefer': method === 'POST' ? 'return=representation' : ''
    }
  };
  if (body) opts.body = JSON.stringify(body);
  const res = await fetch(`${SUPABASE_URL}/rest/v1/${path}`, opts);
  if (!res.ok) {
    const err = await res.text();
    throw new Error(`Supabase fout: ${err}`);
  }
  return res.status === 204 ? null : res.json();
}

// Fallback in-memory (als Supabase niet geconfigureerd)
const geheugen = {};

// ── HOOFDPAGINA ──
app.get('/', (req, res) => {
  const dbStatus = SUPABASE_URL ? '✅ Supabase database verbonden' : '⚠️ Geen database (lokaal geheugen)';
  res.send(`<!DOCTYPE html>
<html lang="nl"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>EY Infra Support API</title>
<style>body{font-family:Arial,sans-serif;background:#f0f4fb;display:flex;align-items:center;justify-content:center;min-height:100vh;margin:0}
.box{background:#fff;border-radius:16px;padding:32px 40px;max-width:440px;width:100%;box-shadow:0 4px 20px rgba(26,42,94,.1);text-align:center}
h1{color:#1a2a5e;font-size:20px;margin:0 0 4px}.sub{color:#7a90b8;font-size:13px;margin-bottom:20px}
.ok{background:#e6f9f0;border:1px solid #a8e6c8;border-radius:10px;padding:10px 14px;color:#1a6b44;font-weight:bold;margin-bottom:8px}
.ep{background:#f0f4fb;border-radius:8px;padding:8px 12px;font-family:monospace;font-size:11px;color:#1a2a5e;margin-top:8px;text-align:left}
.tijd{font-size:11px;color:#aaa;margin-top:12px}</style></head>
<body><div class="box">
  <h1>⚡ EY Infra Support</h1>
  <div class="sub">HR Portaal — Microsoft 365 API</div>
  <div class="ok">✅ Server online</div>
  <div class="ok" style="font-size:13px">${dbStatus}</div>
  <div class="ep">GET  /status<br>GET  /data/:userId/:key<br>POST /data/:userId/:key<br>POST /send-mail</div>
  <div class="tijd">⏱ ${new Date().toLocaleString('nl-NL')}</div>
</div></body></html>`);
});

// ── STATUS ──
app.get('/status', (req, res) => {
  res.json({
    status: 'online',
    naam: 'EY Infra Support — HR Portaal API',
    database: SUPABASE_URL ? 'Supabase' : 'in-memory (tijdelijk)',
    tijd: new Date().toLocaleString('nl-NL')
  });
});

// ── DATA OPHALEN ──
app.get('/data/:userId/:key', async (req, res) => {
  const { userId, key } = req.params;
  try {
    if (SUPABASE_URL) {
      const rows = await sbFetch(
        `user_data?user_id=eq.${encodeURIComponent(userId)}&data_key=eq.${encodeURIComponent(key)}&select=data_value`
      );
      const waarde = rows?.[0]?.data_value ?? null;
      return res.json({ waarde });
    } else {
      const waarde = geheugen[`${userId}:${key}`] ?? null;
      return res.json({ waarde });
    }
  } catch (err) {
    console.error('Data ophalen fout:', err.message);
    res.status(500).json({ error: err.message });
  }
});

// ── DATA OPSLAAN ──
app.post('/data/:userId/:key', async (req, res) => {
  const { userId, key } = req.params;
  const { waarde } = req.body;
  try {
    if (SUPABASE_URL) {
      // Upsert (aanmaken of bijwerken)
      await sbFetch('user_data', 'POST', {
        user_id: userId,
        data_key: key,
        data_value: waarde,
        bijgewerkt: new Date().toISOString()
      });
      // Als al bestaat → update
      await fetch(`${SUPABASE_URL}/rest/v1/user_data?user_id=eq.${encodeURIComponent(userId)}&data_key=eq.${encodeURIComponent(key)}`, {
        method: 'PATCH',
        headers: {
          'Content-Type': 'application/json',
          'apikey': SUPABASE_KEY,
          'Authorization': `Bearer ${SUPABASE_KEY}`
        },
        body: JSON.stringify({ data_value: waarde, bijgewerkt: new Date().toISOString() })
      });
    } else {
      geheugen[`${userId}:${key}`] = waarde;
    }
    res.json({ success: true });
  } catch (err) {
    // Probeer alleen update als insert faalt (duplicate key)
    try {
      if (SUPABASE_URL) {
        await fetch(`${SUPABASE_URL}/rest/v1/user_data?user_id=eq.${encodeURIComponent(userId)}&data_key=eq.${encodeURIComponent(key)}`, {
          method: 'PATCH',
          headers: { 'Content-Type': 'application/json', 'apikey': SUPABASE_KEY, 'Authorization': `Bearer ${SUPABASE_KEY}` },
          body: JSON.stringify({ data_value: waarde, bijgewerkt: new Date().toISOString() })
        });
        return res.json({ success: true });
      }
    } catch(e2) {}
    console.error('Data opslaan fout:', err.message);
    res.status(500).json({ error: err.message });
  }
});

// ── GEDEELDE DATA (voor alle gebruikers, bijv. vakantieaanvragen) ──
app.get('/gedeeld/:key', async (req, res) => {
  const { key } = req.params;
  try {
    if (SUPABASE_URL) {
      const rows = await sbFetch(
        `user_data?user_id=eq._gedeeld&data_key=eq.${encodeURIComponent(key)}&select=data_value`
      );
      return res.json({ waarde: rows?.[0]?.data_value ?? null });
    } else {
      return res.json({ waarde: geheugen[`_gedeeld:${key}`] ?? null });
    }
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.post('/gedeeld/:key', async (req, res) => {
  const { key } = req.params;
  const { waarde } = req.body;
  try {
    if (SUPABASE_URL) {
      await fetch(`${SUPABASE_URL}/rest/v1/user_data`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', 'apikey': SUPABASE_KEY, 'Authorization': `Bearer ${SUPABASE_KEY}`, 'Prefer': 'resolution=merge-duplicates' },
        body: JSON.stringify({ user_id: '_gedeeld', data_key: key, data_value: waarde, bijgewerkt: new Date().toISOString() })
      });
    } else {
      geheugen[`_gedeeld:${key}`] = waarde;
    }
    res.json({ success: true });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// ── PROJECTEN ──
const STANDAARD_PROJECTEN = ['Heerle','W-Papendrecht','W-Dintelweg','W-Vijfhuizen','Infra overige','Intern'];
app.get('/projecten', async (req, res) => {
  try {
    const r = await fetch(`http://localhost:${process.env.PORT||3001}/gedeeld/projecten`);
    const d = await r.json();
    res.json({ projecten: d.waarde || STANDAARD_PROJECTEN });
  } catch(e) {
    res.json({ projecten: STANDAARD_PROJECTEN });
  }
});
app.post('/projecten', async (req, res) => {
  const { projecten } = req.body;
  if (!Array.isArray(projecten)) return res.status(400).json({ error: 'projecten moet een lijst zijn' });
  try {
    await fetch(`http://localhost:${process.env.PORT||3001}/gedeeld/projecten`, {
      method:'POST', headers:{'Content-Type':'application/json'},
      body: JSON.stringify({waarde: projecten})
    });
    res.json({ success: true, projecten });
  } catch(e) {
    res.status(500).json({ error: e.message });
  }
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
  console.log(`🚀 EY Infra Support API draait op poort ${PORT}`);
  console.log(`📧 Afzender: ${AFZENDER}`);
  console.log(`🗄️  Database: ${SUPABASE_URL ? 'Supabase ✅' : 'In-memory (voeg SUPABASE_URL toe voor persistentie)'}`);
});
