const express = require('express');
const { ImapFlow } = require('imapflow');
const { simpleParser } = require('mailparser');
const crypto = require('crypto');
const path = require('path');

const app = express();
app.use(express.json({ limit: '10mb' }));
app.use(express.static(path.join(__dirname, 'public')));

// ── Session store ──────────────────────────────────────────────
const sessions = new Map(); // id -> { client, email, lastUsed }

setInterval(() => {
  const now = Date.now();
  for (const [id, sess] of sessions) {
    if (now - sess.lastUsed > 30 * 60 * 1000) {
      sess.client.logout().catch(() => {});
      sessions.delete(id);
    }
  }
}, 60 * 1000);

// ── Helpers ────────────────────────────────────────────────────
function getSession(req, res) {
  const id = req.headers['x-session'] || req.query.session;
  const sess = sessions.get(id);
  if (!sess) { res.status(401).json({ error: 'Session invalide ou expirée' }); return null; }
  sess.lastUsed = Date.now();
  return sess;
}

function fmtAddr(addr) {
  if (!addr) return null;
  return { name: addr.name || '', address: addr.address || '' };
}

// ── IMAP Presets ───────────────────────────────────────────────
const PRESETS = {
  'gmail.com':     { host: 'imap.gmail.com',    port: 993, secure: true },
  'googlemail.com':{ host: 'imap.gmail.com',    port: 993, secure: true },
  'outlook.com':   { host: 'outlook.office365.com', port: 993, secure: true },
  'hotmail.com':   { host: 'outlook.office365.com', port: 993, secure: true },
  'live.com':      { host: 'outlook.office365.com', port: 993, secure: true },
  'yahoo.com':     { host: 'imap.mail.yahoo.com',   port: 993, secure: true },
  'yahoo.fr':      { host: 'imap.mail.yahoo.com',   port: 993, secure: true },
  'icloud.com':    { host: 'imap.mail.me.com',      port: 993, secure: true },
  'me.com':        { host: 'imap.mail.me.com',      port: 993, secure: true },
  'proton.me':     { host: '127.0.0.1',             port: 1143, secure: false },
  'protonmail.com':{ host: '127.0.0.1',             port: 1143, secure: false },
  'ovh.com':       { host: 'ssl0.ovh.net',          port: 993, secure: true },
  'laposte.net':   { host: 'imap.laposte.net',      port: 993, secure: true },
  'orange.fr':     { host: 'imap.orange.fr',        port: 993, secure: true },
  'free.fr':       { host: 'imap.free.fr',          port: 993, secure: true },
  'sfr.fr':        { host: 'imap.sfr.fr',           port: 993, secure: true },
};

app.get('/api/preset', (req, res) => {
  const domain = (req.query.email || '').split('@')[1]?.toLowerCase();
  const preset = domain ? PRESETS[domain] : null;
  res.json(preset || {});
});

// ── Connect ────────────────────────────────────────────────────
app.post('/api/connect', async (req, res) => {
  const { host, port, secure, user, pass } = req.body;
  if (!host || !user || !pass) return res.status(400).json({ error: 'Paramètres manquants' });

  const client = new ImapFlow({
    host,
    port: parseInt(port) || (secure ? 993 : 143),
    secure: secure !== false,
    auth: { user, pass },
    logger: false,
    tls: { rejectUnauthorized: false },
    connectionTimeout: 10000,
    greetingTimeout: 5000,
    socketTimeout: 30000,
  });

  try {
    await client.connect();
    const id = crypto.randomUUID();
    sessions.set(id, { client, email: user, lastUsed: Date.now() });

    // Get folders
    const rawFolders = await client.list();
    const folders = rawFolders.map(mb => ({
      path: mb.path,
      name: mb.name,
      delimiter: mb.delimiter,
      flags: mb.flags ? [...mb.flags] : [],
      specialUse: mb.specialUse || null,
      subscribed: mb.subscribed,
    }));

    res.json({ sessionId: id, email: user, folders });
  } catch (err) {
    try { await client.logout(); } catch {}
    res.status(401).json({ error: err.message || 'Connexion échouée' });
  }
});

// ── Folders ────────────────────────────────────────────────────
app.get('/api/folders', async (req, res) => {
  const sess = getSession(req, res); if (!sess) return;
  try {
    const raw = await sess.client.list();
    const folders = await Promise.all(raw.map(async (mb) => {
      let count = null;
      try {
        const s = await sess.client.status(mb.path, { messages: true, unseen: true });
        count = { total: s.messages, unseen: s.unseen };
      } catch {}
      return {
        path: mb.path, name: mb.name,
        flags: mb.flags ? [...mb.flags] : [],
        specialUse: mb.specialUse || null,
        count,
      };
    }));
    res.json(folders);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// ── Messages list ──────────────────────────────────────────────
app.get('/api/messages', async (req, res) => {
  const sess = getSession(req, res); if (!sess) return;
  const { folder = 'INBOX', page = '1', limit = '40' } = req.query;
  const pageNum = Math.max(1, parseInt(page));
  const pageSize = Math.max(1, Math.min(100, parseInt(limit)));

  try {
    const lock = await sess.client.getMailboxLock(folder);
    try {
      const status = await sess.client.status(folder, { messages: true, unseen: true, recent: true });
      const total = status.messages || 0;

      if (total === 0) return res.json({ messages: [], total: 0, page: pageNum, pages: 0, unseen: 0 });

      const end = total - (pageNum - 1) * pageSize;
      const start = Math.max(1, end - pageSize + 1);
      if (end < 1) return res.json({ messages: [], total, page: pageNum, pages: Math.ceil(total / pageSize), unseen: status.unseen || 0 });

      const messages = [];
      for await (const msg of sess.client.fetch(`${start}:${end}`, {
        uid: true, flags: true, envelope: true, size: true,
        headers: ['content-type', 'x-priority'],
      })) {
        messages.push({
          seq: msg.seq,
          uid: msg.uid,
          seen: msg.flags.has('\\Seen'),
          flagged: msg.flags.has('\\Flagged'),
          answered: msg.flags.has('\\Answered'),
          draft: msg.flags.has('\\Draft'),
          subject: msg.envelope?.subject || '(Sans objet)',
          from: fmtAddr(msg.envelope?.from?.[0]),
          to: (msg.envelope?.to || []).map(fmtAddr),
          date: msg.envelope?.date?.toISOString() || null,
          size: msg.size || 0,
        });
      }
      messages.reverse();
      res.json({
        messages,
        total,
        page: pageNum,
        pages: Math.ceil(total / pageSize),
        unseen: status.unseen || 0,
      });
    } finally {
      lock.release();
    }
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// ── Full message ───────────────────────────────────────────────
app.get('/api/message', async (req, res) => {
  const sess = getSession(req, res); if (!sess) return;
  const { folder = 'INBOX', uid } = req.query;
  if (!uid) return res.status(400).json({ error: 'UID requis' });

  try {
    const lock = await sess.client.getMailboxLock(folder);
    try {
      let source = null;
      for await (const msg of sess.client.fetch(uid, { source: true }, { uid: true })) {
        source = msg.source;
      }
      if (!source) return res.status(404).json({ error: 'Message introuvable' });

      // Mark as seen
      await sess.client.messageFlagsAdd(uid, ['\\Seen'], { uid: true });

      const parsed = await simpleParser(source);
      res.json({
        uid: parseInt(uid),
        subject: parsed.subject || '(Sans objet)',
        from: parsed.from?.value?.map(a => ({ name: a.name || '', address: a.address || '' })) || [],
        to: parsed.to?.value?.map(a => ({ name: a.name || '', address: a.address || '' })) || [],
        cc: parsed.cc?.value?.map(a => ({ name: a.name || '', address: a.address || '' })) || [],
        date: parsed.date?.toISOString() || null,
        html: parsed.html || null,
        text: parsed.text || null,
        textAsHtml: parsed.textAsHtml || null,
        attachments: (parsed.attachments || []).map(a => ({
          filename: a.filename || 'fichier',
          contentType: a.contentType || 'application/octet-stream',
          size: a.size || 0,
          contentId: a.contentId,
        })),
      });
    } finally {
      lock.release();
    }
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// ── Flags ──────────────────────────────────────────────────────
app.post('/api/flags', async (req, res) => {
  const sess = getSession(req, res); if (!sess) return;
  const { folder, uid, flag, set } = req.body;
  if (!folder || !uid || !flag) return res.status(400).json({ error: 'Paramètres manquants' });

  try {
    const lock = await sess.client.getMailboxLock(folder);
    try {
      if (set) {
        await sess.client.messageFlagsAdd(String(uid), [flag], { uid: true });
      } else {
        await sess.client.messageFlagsRemove(String(uid), [flag], { uid: true });
      }
      res.json({ ok: true });
    } finally {
      lock.release();
    }
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// ── Delete ─────────────────────────────────────────────────────
app.delete('/api/message', async (req, res) => {
  const sess = getSession(req, res); if (!sess) return;
  const { folder, uid } = req.query;
  if (!folder || !uid) return res.status(400).json({ error: 'Paramètres manquants' });

  try {
    const lock = await sess.client.getMailboxLock(folder);
    try {
      await sess.client.messageDelete(String(uid), { uid: true });
      res.json({ ok: true });
    } finally {
      lock.release();
    }
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// ── Move ───────────────────────────────────────────────────────
app.post('/api/move', async (req, res) => {
  const sess = getSession(req, res); if (!sess) return;
  const { fromFolder, toFolder, uid } = req.body;
  if (!fromFolder || !toFolder || !uid) return res.status(400).json({ error: 'Paramètres manquants' });

  try {
    const lock = await sess.client.getMailboxLock(fromFolder);
    try {
      await sess.client.messageMove(String(uid), toFolder, { uid: true });
      res.json({ ok: true });
    } finally {
      lock.release();
    }
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// ── Folder status ──────────────────────────────────────────────
app.get('/api/status', async (req, res) => {
  const sess = getSession(req, res); if (!sess) return;
  const { folder = 'INBOX' } = req.query;
  try {
    const s = await sess.client.status(folder, { messages: true, unseen: true, recent: true });
    res.json({ total: s.messages, unseen: s.unseen, recent: s.recent });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// ── Disconnect ─────────────────────────────────────────────────
app.post('/api/disconnect', async (req, res) => {
  const { sessionId } = req.body;
  const sess = sessions.get(sessionId);
  if (sess) {
    sess.client.logout().catch(() => {});
    sessions.delete(sessionId);
  }
  res.json({ ok: true });
});

// ── Start ──────────────────────────────────────────────────────
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`\n  MailFlow — http://localhost:${PORT}\n`);
});