const express = require('express');
const cors = require('cors');
const path = require('path');
const session = require('express-session');
const generateRoute = require('./routes/generate');
const templatesRoute = require('./routes/templates');

const app = express();
const PORT = process.env.PORT || 3000;

// ─── Credentials (set via env vars, default for dev) ───
const APP_USERNAME = process.env.APP_USERNAME || 'admin';
const APP_PASSWORD = process.env.APP_PASSWORD || 'skypec2026';
const SESSION_SECRET = process.env.SESSION_SECRET || 'skypec-cert-secret-key-2026';

// ─── Middleware ────────────────────────────────────────
app.use(cors());
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

app.use(session({
  secret: SESSION_SECRET,
  resave: false,
  saveUninitialized: false,
  cookie: {
    secure: false, // set true if HTTPS
    maxAge: 8 * 60 * 60 * 1000 // 8 hours
  }
}));

// ─── Auth middleware ───────────────────────────────────
function requireAuth(req, res, next) {
  if (req.session && req.session.loggedIn) return next();
  // AJAX/API requests → 401
  if (req.path.startsWith('/api/')) return res.status(401).json({ error: 'Unauthorized' });
  res.redirect('/login');
}

// ─── Public routes (no auth needed) ───────────────────
app.get('/login', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'login.html'));
});

app.post('/login', (req, res) => {
  const { username, password } = req.body;
  if (username === APP_USERNAME && password === APP_PASSWORD) {
    req.session.loggedIn = true;
    req.session.username = username;
    return res.redirect('/');
  }
  res.redirect('/login?error=1');
});

app.get('/logout', (req, res) => {
  req.session.destroy(() => {
    res.redirect('/login');
  });
});

// ─── Protected routes ──────────────────────────────────
app.use('/api', requireAuth, generateRoute);
app.use('/api', requireAuth, templatesRoute);

app.get('/', requireAuth, (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

// Serve static files (login.html, style.css, app.js - accessible publicly for login page)
app.use(express.static(path.join(__dirname, 'public')));

app.listen(PORT, () => {
  console.log(`✅ Certificate Generator running at http://localhost:${PORT}`);
  console.log(`👤 Login: ${APP_USERNAME} / ${APP_PASSWORD}`);
});
