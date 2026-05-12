const express = require('express');
const { ClientSecretCredential } = require('@azure/identity');
const { Client } = require('@microsoft/microsoft-graph-client');
const { TokenCredentialAuthenticationProvider } = require('@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials');

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
