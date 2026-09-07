import 'dotenv/config';
import express from 'express';
import helmet from 'helmet';
import cookieParser from 'cookie-parser';
import rateLimit from 'express-rate-limit';
import bcrypt from 'bcryptjs';
import crypto from 'crypto';
import pg from 'pg';

const { Pool } = pg;
const app = express();
const pool = new Pool({ connectionString: process.env.DATABASE_URL });
const PORT = Number(process.env.PORT || 3001);
const SESSION_DAYS = 7;
const SESSION_HOURS = 12;
const REMEMBER_DAYS = 30;
const cookieName = 'malidesk_session';
const cookieSameSite = process.env.COOKIE_SAMESITE || ((process.env.CORS_ORIGIN && process.env.COOKIE_SECURE !== 'false') ? 'none' : 'lax');
const sessionCookieOptions = (extra = {}) => ({ httpOnly:true, secure:process.env.COOKIE_SECURE !== 'false', sameSite:cookieSameSite, path:'/', ...extra });

if (!process.env.DATABASE_URL) {
  console.error('DATABASE_URL is missing. Create server/.env from server/.env.example before starting MaliDesk.');
}

app.set('trust proxy', 1);
app.use(helmet({ contentSecurityPolicy: false }));
app.use(express.json({ limit: '2mb' }));
app.use(cookieParser());

// Allow the separately deployed frontend to call the API with session cookies.
// Keep the origin explicit; wildcard origins are incompatible with credentialed requests.
const allowedOrigins = String(process.env.CORS_ORIGIN || '')
  .split(',')
  .map(v => v.trim())
  .filter(Boolean);
app.use((req, res, next) => {
  const origin = req.get('origin');
  if (origin && (allowedOrigins.length === 0 || allowedOrigins.includes(origin))) {
    res.setHeader('Access-Control-Allow-Origin', origin);
    res.setHeader('Vary', 'Origin');
    res.setHeader('Access-Control-Allow-Credentials', 'true');
    res.setHeader('Access-Control-Allow-Headers', 'Content-Type, Authorization');
    res.setHeader('Access-Control-Allow-Methods', 'GET,POST,PUT,PATCH,DELETE,OPTIONS');
  }
  if (req.method === 'OPTIONS') return res.sendStatus(204);
  next();
});
app.use(rateLimit({ windowMs: 15 * 60 * 1000, limit: 100, standardHeaders: true, legacyHeaders: false }));

app.get('/', (req, res) => res.json({ ok: true, service: 'malidesk-api' }));
app.get('/api/health', async (req,res) => { try { await pool.query('SELECT 1'); res.json({ok:true,service:'malidesk-api'}); } catch (e) { console.error('Health check failed:', e.message); res.status(503).json({ok:false,service:'malidesk-api',error:'Database unavailable'}); } });

const authLimiter = rateLimit({ windowMs: 15 * 60 * 1000, limit: 10, standardHeaders: true, legacyHeaders: false, message: { error: 'Authentication temporarily unavailable. Try again later.' } });
const hashToken = (token) => crypto.createHash('sha256').update(token).digest('hex');
const newToken = () => crypto.randomBytes(32).toString('base64url');

async function audit(req, action, entity = null, entityId = null, details = {}) {
  const actorId = req.user?.id || null;
  const ip = req.ip || null;
  const ua = (typeof req.get === 'function' ? req.get('user-agent') : '') || '';
  await pool.query(
    `INSERT INTO audit_logs(actor_user_id,action,entity,entity_id,ip_address,user_agent,details) 
     VALUES($1,$2,$3,$4,$5,$6,$7)`,
    [actorId, action, entity, entityId, ip, ua, JSON.stringify(details)]
  );
}

async function requireAuth(req, res, next) {
  try {
    const raw = req.cookies[cookieName];
    if (!raw) return res.status(401).json({ error: 'Authentication required' });
    const { rows } = await pool.query(`SELECT u.id,u.full_name,u.username,u.email,u.role,u.status,u.must_change_password,s.id AS session_id FROM sessions s JOIN app_users u ON u.id=s.user_id WHERE s.token_hash=$1 AND s.revoked_at IS NULL AND s.expires_at>now()`, [hashToken(raw)]);
    if (!rows[0] || rows[0].status !== 'active') return res.status(401).json({ error: 'Authentication required' });
    req.user = rows[0];
    next();
  } catch { res.status(500).json({ error: 'Authentication service error' }); }
}

const roleDefaults = {
  Administrator: '*',
  Manager: new Set(['view_dashboard','view_units','create_unit','edit_unit','view_tenants','create_tenant','edit_tenant','record_payment','edit_payment','delete_payment','manage_monthly_charges','edit_monthly_charges','delete_monthly_charges','manage_adjustments','close_tenancy','view_reports','export_data']),
  Staff: new Set(['view_dashboard','view_units','view_tenants','record_payment','manage_monthly_charges','manage_adjustments','view_reports']),
  Viewer: new Set(['view_dashboard','view_units','view_tenants','view_reports'])
};

async function hasPermission(user, key) {
  if (user.role === 'Administrator') return true;
  const { rows } = await pool.query('SELECT 1 FROM user_permissions WHERE user_id=$1 AND permission_key=$2', [user.id, key]);
  return rows.length > 0 || roleDefaults[user.role]?.has(key) === true;
}

const requirePermission = (key) => async (req, res, next) => {
  if (!(await hasPermission(req.user, key))) return res.status(403).json({ error: 'Access denied' });
  next();
};



const ALL_DATA_KEYS = ['version','units','tenants','transactions','archive','auditLog','settings','deletedImportKeys'];
const arrayOrEmpty = (v) => Array.isArray(v) ? v : [];
const objectOrEmpty = (v) => v && typeof v === 'object' && !Array.isArray(v) ? v : {};
const clone = (v) => JSON.parse(JSON.stringify(v));

function normaliseServerData(raw) {
  const d = objectOrEmpty(raw);
  return {
    version: 3,
    units: arrayOrEmpty(d.units),
    tenants: arrayOrEmpty(d.tenants),
    transactions: arrayOrEmpty(d.transactions).filter(t => t && typeof t.date === 'string' && t.date.trim() !== ''),
    archive: arrayOrEmpty(d.archive),
    auditLog: [],
    settings: objectOrEmpty(d.settings),
    deletedImportKeys: arrayOrEmpty(d.deletedImportKeys),
  };
}

function mapById(list) {
  return new Map(list.filter(x => x && x.id).map(x => [String(x.id), x]));
}

function dataChangeRequirements(oldData, newData) {
  const required = new Set();
  const changes = [];
  const oldUnits = mapById(arrayOrEmpty(oldData.units));
  const newUnits = mapById(arrayOrEmpty(newData.units));
  for (const [id] of oldUnits) if (!newUnits.has(id)) { required.add('delete_unit'); changes.push(['UNIT_DELETED', id]); }
  for (const [id] of newUnits) if (!oldUnits.has(id)) { required.add('create_unit'); changes.push(['UNIT_CREATED', id]); }
  for (const [id, u] of newUnits) if (oldUnits.has(id) && JSON.stringify(u) !== JSON.stringify(oldUnits.get(id))) { required.add('edit_unit'); changes.push(['UNIT_UPDATED', id]); }

  const oldTenants = mapById(arrayOrEmpty(oldData.tenants));
  const newTenants = mapById(arrayOrEmpty(newData.tenants));
  for (const [id] of oldTenants) if (!newTenants.has(id)) { required.add('delete_tenant'); changes.push(['TENANT_DELETED', id]); }
  for (const [id] of newTenants) if (!oldTenants.has(id)) { required.add('create_tenant'); changes.push(['TENANT_CREATED', id]); }
  for (const [id, t] of newTenants) if (oldTenants.has(id) && JSON.stringify(t) !== JSON.stringify(oldTenants.get(id))) { required.add('edit_tenant'); changes.push(['TENANT_UPDATED', id]); }

  const oldTx = mapById(arrayOrEmpty(oldData.transactions));
  const newTx = mapById(arrayOrEmpty(newData.transactions));
  for (const [id] of oldTx) if (!newTx.has(id)) {
    required.add(oldTx.get(id)?.type === 'payment' ? 'delete_payment' : 'delete_monthly_charges');
    changes.push([oldTx.get(id)?.type === 'payment' ? 'PAYMENT_DELETED' : 'CHARGE_DELETED', id]);
  }
  for (const [id, t] of newTx) if (!oldTx.has(id)) {
    required.add(t?.type === 'payment' ? 'record_payment' : 'manage_monthly_charges');
    changes.push([t?.type === 'payment' ? 'PAYMENT_CREATED' : 'CHARGE_CREATED', id]);
  }
  for (const [id, t] of newTx) if (oldTx.has(id) && JSON.stringify(t) !== JSON.stringify(oldTx.get(id))) {
    required.add(t?.type === 'payment' ? 'edit_payment' : 'edit_monthly_charges');
    changes.push([t?.type === 'payment' ? 'PAYMENT_UPDATED' : 'CHARGE_UPDATED', id]);
  }

  if (JSON.stringify(oldData.archive) !== JSON.stringify(newData.archive)) { required.add('close_tenancy'); changes.push(['ARCHIVE_CHANGED', null]); }
  if (JSON.stringify(oldData.settings) !== JSON.stringify(newData.settings)) { required.add('manage_settings'); changes.push(['SETTINGS_CHANGED', null]); }
  // deletedImportKeys is derived local import bookkeeping; it is not an independent privilege.
  // Import/replace operations are protected by the dedicated import_data permission.

  return { required, changes };
}

async function loadRentalState() {
  const { rows } = await pool.query('SELECT version,data FROM rental_state WHERE id=1');
  return rows[0] || { version: 1, data: normaliseServerData({}) };
}

app.get('/api/data', requireAuth, async (req, res) => {
  try {
    const state = await loadRentalState();
    res.json({ version: Number(state.version), data: normaliseServerData(state.data) });
  } catch (e) { console.error(e); res.status(500).json({ error: 'Unable to load rental data' }); }
});

app.put('/api/data', requireAuth, async (req, res) => {
  const expectedVersion = Number(req.body?.expectedVersion);
  const incoming = normaliseServerData(req.body?.data);
  if (!Number.isInteger(expectedVersion) || expectedVersion < 1) return res.status(400).json({ error: 'A valid expectedVersion is required' });
  if (incoming.units.length > 10000 || incoming.tenants.length > 20000 || incoming.transactions.length > 200000 || incoming.archive.length > 20000) return res.status(413).json({ error: 'Rental dataset is too large' });
  if (JSON.stringify(incoming).length > 8 * 1024 * 1024) return res.status(413).json({ error: 'Rental dataset exceeds the maximum size' });
  try {
    const client = await pool.connect();
    try {
      await client.query('BEGIN');
      const current = await client.query('SELECT version,data FROM rental_state WHERE id=1 FOR UPDATE');
      const row = current.rows[0] || { version: 1, data: normaliseServerData({}) };
      if (Number(row.version) !== expectedVersion) {
        await client.query('ROLLBACK');
        return res.status(409).json({ error: 'Rental data changed elsewhere. Reload before saving.', version: Number(row.version), data: normaliseServerData(row.data) });
      }
      const oldData = normaliseServerData(row.data);
      const { required, changes } = dataChangeRequirements(oldData, incoming);
      for (const permission of required) {
        if (!(await hasPermission(req.user, permission))) {
          await client.query('ROLLBACK');
          return res.status(403).json({ error: `Access denied: ${permission}` });
        }
      }
      incoming.auditLog = [];
      const nextVersion = Number(row.version) + 1;
      await client.query('UPDATE rental_state SET version=$1,data=$2,updated_by=$3,updated_at=now() WHERE id=1', [nextVersion, JSON.stringify(incoming), req.user.id]);
      await client.query('COMMIT');
      for (const [action, entityId] of changes.slice(0, 50)) await audit(req, action, 'rental_data', entityId, { version: nextVersion });
      res.json({ ok: true, version: nextVersion, data: incoming });
    } catch (e) { await client.query('ROLLBACK'); throw e; } finally { client.release(); }
  } catch (e) { console.error(e); res.status(500).json({ error: 'Unable to save rental data' }); }
});

app.post('/api/data/bootstrap', requireAuth, requirePermission('import_data'), async (req, res) => {
  const incoming = normaliseServerData(req.body?.data);
  try {
    const current = await loadRentalState();
    const hasData = arrayOrEmpty(current.data?.units).length || arrayOrEmpty(current.data?.tenants).length || arrayOrEmpty(current.data?.transactions).length;
    if (hasData) return res.status(409).json({ error: 'Rental data already exists' });
    const nextVersion = Number(current.version) + 1;
    await pool.query('UPDATE rental_state SET version=$1,data=$2,updated_by=$3,updated_at=now() WHERE id=1', [nextVersion, JSON.stringify(incoming), req.user.id]);
    await audit(req, 'DATA_BOOTSTRAPPED', 'rental_data', '1', { version: nextVersion });
    res.json({ ok:true, version:nextVersion, data:incoming });
  } catch (e) { console.error(e); res.status(500).json({ error:'Unable to initialise rental data' }); }
});

app.post('/api/auth/login', authLimiter, async (req, res) => {
  const username = String(req.body?.username || '').trim();
  const password = String(req.body?.password || '');
  const remember = Boolean(req.body?.remember);
  try {
    const { rows } = await pool.query('SELECT * FROM app_users WHERE lower(username)=lower($1)', [username]);
    const user = rows[0];
    const valid = user && user.status === 'active' && await bcrypt.compare(password, user.password_hash);
    if (!valid) {
      await audit(req, 'LOGIN_FAILED', 'user', user?.id || null);
      return res.status(401).json({ error: 'Invalid username or password' });
    }
    const token = newToken();
    await pool.query('INSERT INTO sessions(user_id,token_hash,expires_at) VALUES($1,$2,now()+($3 || \' days\')::interval)', [user.id, hashToken(token), SESSION_DAYS]);
    await pool.query('UPDATE app_users SET last_login_at=now(),updated_at=now() WHERE id=$1', [user.id]);
    await audit({ ...req, user: { id: user.id } }, 'LOGIN_SUCCESS', 'user', user.id);
    res.cookie(cookieName, token, { httpOnly: true, secure: process.env.COOKIE_SECURE !== 'false', sameSite: 'lax', ...(remember ? { maxAge: REMEMBER_DAYS * 86400000 } : { maxAge: SESSION_HOURS * 60 * 60 * 1000 }), path: '/' });
    res.json({ user: { id:user.id, fullName:user.full_name, username:user.username, email:user.email, role:user.role, mustChangePassword:user.must_change_password } });
  } catch (e) { console.error(e); res.status(500).json({ error: 'Unable to sign in' }); }
});

app.post('/api/auth/logout', requireAuth, async (req, res) => {
  await pool.query('UPDATE sessions SET revoked_at=now() WHERE id=$1', [req.user.session_id]);
  await audit(req, 'LOGOUT', 'user', req.user.id);
  res.clearCookie(cookieName, sessionCookieOptions());
  res.json({ ok:true });
});

app.get('/api/auth/me', requireAuth, async (req,res) => res.json({ user:{ id:req.user.id,fullName:req.user.full_name,username:req.user.username,email:req.user.email,role:req.user.role,mustChangePassword:req.user.must_change_password } }));

app.get('/api/users', requireAuth, requirePermission('manage_users'), async (req,res) => {
  const { rows } = await pool.query('SELECT id,full_name,username,email,role,status,last_login_at,created_at,must_change_password FROM app_users ORDER BY full_name');
  res.json({ users: rows });
});

app.post('/api/users', requireAuth, requirePermission('manage_users'), async (req,res) => {
  const { fullName, username, email, role='Viewer', password } = req.body || {};
  if (!fullName?.trim() || !username?.trim() || !password) return res.status(400).json({ error:'Full name, username and password are required' });
  if (password.length < 8 || !/[a-z]/.test(password) || !/[A-Z]/.test(password) || !/[0-9]/.test(password)) return res.status(400).json({ error:'Password must be at least 8 characters and include upper, lower and number' });
  if (role === 'Administrator' && req.user.role !== 'Administrator') return res.status(403).json({ error:'Only an Administrator can assign Administrator role' });
  const hash = await bcrypt.hash(password, 12);
  const client = await pool.connect();
  try {
    await client.query('BEGIN');
    const { rows } = await client.query('INSERT INTO app_users(full_name,username,email,password_hash,role,must_change_password) VALUES($1,$2,$3,$4,$5,true) RETURNING id,full_name,username,email,role,status,created_at', [fullName.trim(),username.trim(),email?.trim()||null,hash,role]);
    await client.query('COMMIT');
    await audit(req,'USER_CREATED','user',rows[0].id,{role});
    res.status(201).json({ user:rows[0] });
  } catch(e) { await client.query('ROLLBACK'); if(e.code==='23505') return res.status(409).json({error:'Username or email already exists'}); res.status(500).json({error:'Unable to create user'}); } finally { client.release(); }
});

app.patch('/api/users/:id', requireAuth, requirePermission('manage_users'), async (req,res) => {
  const target = req.params.id;
  const { fullName, username, email, role, status, password } = req.body || {};
  if (!fullName?.trim() || !username?.trim()) return res.status(400).json({error:'Full name and username are required'});
  if (!['Administrator','Manager','Staff','Viewer'].includes(role)) return res.status(400).json({error:'Invalid role'});
  if (!['active','inactive'].includes(status)) return res.status(400).json({error:'Invalid status'});
  if (role === 'Administrator' && req.user.role !== 'Administrator') return res.status(403).json({error:'Only an Administrator can assign Administrator role'});
  const { rows: existing } = await pool.query('SELECT id,role,status FROM app_users WHERE id=$1',[target]);
  if (!existing[0]) return res.status(404).json({error:'User not found'});
  if (existing[0].role === 'Administrator' && role !== 'Administrator') {
    const { rows:a } = await pool.query("SELECT count(*)::int AS n FROM app_users WHERE role='Administrator' AND status='active'");
    if (existing[0].status === 'active' && a[0].n <= 1) return res.status(409).json({error:'The last active Administrator cannot be demoted'});
  }
  if (existing[0].role === 'Administrator' && status === 'inactive' && existing[0].status === 'active') {
    const { rows:a } = await pool.query("SELECT count(*)::int AS n FROM app_users WHERE role='Administrator' AND status='active'");
    if (a[0].n <= 1) return res.status(409).json({error:'The last active Administrator cannot be deactivated'});
  }
  if (target === req.user.id && status === 'inactive') return res.status(409).json({error:'You cannot deactivate your own account'});
  if (password !== undefined && password !== '') {
    if (String(password).length < 8 || !/[a-z]/.test(password) || !/[A-Z]/.test(password) || !/[0-9]/.test(password)) return res.status(400).json({error:'Password must be at least 8 characters and include upper, lower and number'});
  }
  try {
    const fields = ['full_name=$1','username=$2','email=$3','role=$4','status=$5','updated_at=now()'];
    const values = [fullName.trim(), username.trim(), email?.trim() || null, role, status];
    if (password !== undefined && password !== '') { fields.push(`password_hash=$${values.length+1}`); values.push(await bcrypt.hash(password,12)); fields.push('must_change_password=true'); }
    values.push(target);
    const { rows } = await pool.query(`UPDATE app_users SET ${fields.join(',')} WHERE id=$${values.length} RETURNING id,full_name,username,email,role,status,last_login_at,created_at,must_change_password`, values);
    if (password !== undefined && password !== '') await pool.query('UPDATE sessions SET revoked_at=now() WHERE user_id=$1 AND id<>$2',[target,req.user.session_id]);
    await audit(req,'USER_UPDATED','user',target,{role,status,passwordChanged:Boolean(password)});
    res.json({ok:true,user:rows[0]});
  } catch(e) { if(e.code==='23505') return res.status(409).json({error:'Username or email already exists'}); console.error(e); res.status(500).json({error:'Unable to update user'}); }
});

app.patch('/api/users/:id/status', requireAuth, requirePermission('manage_users'), async (req,res) => {
  const target = req.params.id;
  const status = req.body?.status;
  if (!['active','inactive'].includes(status)) return res.status(400).json({error:'Invalid status'});
  const { rows } = await pool.query('SELECT id,role,status FROM app_users WHERE id=$1',[target]);
  if(!rows[0]) return res.status(404).json({error:'User not found'});
  if (rows[0].role==='Administrator' && status==='inactive') {
    const { rows: admins } = await pool.query("SELECT count(*)::int AS n FROM app_users WHERE role='Administrator' AND status='active'");
    if (admins[0].n <= 1) return res.status(409).json({error:'The last active Administrator cannot be deactivated'});
  }
  await pool.query('UPDATE app_users SET status=$1,updated_at=now() WHERE id=$2',[status,target]);
  await audit(req,status==='active'?'USER_ACTIVATED':'USER_DEACTIVATED','user',target);
  res.json({ok:true});
});




app.get('/api/permissions', requireAuth, requirePermission('manage_permissions'), async (req,res) => {
  const { rows } = await pool.query('SELECT key,description FROM permissions ORDER BY key');
  res.json({ permissions: rows });
});

app.get('/api/users/:id/permissions', requireAuth, requirePermission('manage_permissions'), async (req,res) => {
  const { rows } = await pool.query('SELECT permission_key FROM user_permissions WHERE user_id=$1 ORDER BY permission_key', [req.params.id]);
  res.json({ permissions: rows.map(r=>r.permission_key) });
});

app.put('/api/users/:id/permissions', requireAuth, requirePermission('manage_permissions'), async (req,res) => {
  const target=req.params.id;
  const keys=Array.isArray(req.body?.permissions) ? req.body.permissions.map(String) : [];
  if (req.user.id===target && req.user.role!=='Administrator') return res.status(403).json({error:'Only an Administrator may change permissions'});
  const { rows: valid } = await pool.query('SELECT key FROM permissions WHERE key = ANY($1::text[])', [keys]);
  if (valid.length !== keys.length) return res.status(400).json({error:'One or more permissions are invalid'});
  const client=await pool.connect();
  try {
    await client.query('BEGIN');
    await client.query('DELETE FROM user_permissions WHERE user_id=$1',[target]);
    if(keys.length) await client.query('INSERT INTO user_permissions(user_id,permission_key) SELECT $1,key FROM unnest($2::text[]) AS key',[target,keys]);
    await client.query('COMMIT');
    await audit(req,'USER_PERMISSIONS_CHANGED','user',target,{permissions:keys});
    res.json({ok:true,permissions:keys});
  } catch(e) { await client.query('ROLLBACK'); console.error(e); res.status(500).json({error:'Unable to update permissions'}); } finally { client.release(); }
});

app.post('/api/users/:id/reset-password', requireAuth, requirePermission('manage_users'), async (req,res) => {
  const target=req.params.id;
  if(target===req.user.id) return res.status(409).json({error:'Use Change Password for your own account'});
  const { rows } = await pool.query('SELECT id FROM app_users WHERE id=$1',[target]);
  if(!rows[0]) return res.status(404).json({error:'User not found'});
  const temporaryPassword = crypto.randomBytes(9).toString('base64url');
  await pool.query('UPDATE app_users SET password_hash=$1,must_change_password=true,updated_at=now() WHERE id=$2',[await bcrypt.hash(temporaryPassword,12),target]);
  await pool.query('UPDATE sessions SET revoked_at=now() WHERE user_id=$1',[target]);
  await audit(req,'PASSWORD_RESET','user',target);
  res.json({ok:true,temporaryPassword});
});

app.patch('/api/users/:id/role', requireAuth, requirePermission('manage_roles'), async (req,res) => {
  const target=req.params.id, role=req.body?.role;
  if(!['Administrator','Manager','Staff','Viewer'].includes(role)) return res.status(400).json({error:'Invalid role'});
  if(role==='Administrator' && req.user.role!=='Administrator') return res.status(403).json({error:'Only an Administrator can assign Administrator role'});
  const {rows}=await pool.query('SELECT id,role,status FROM app_users WHERE id=$1',[target]);
  if(!rows[0]) return res.status(404).json({error:'User not found'});
  if(rows[0].role==='Administrator' && role!=='Administrator') {
    const {rows:a}=await pool.query("SELECT count(*)::int AS n FROM app_users WHERE role='Administrator' AND status='active'");
    if(a[0].n<=1) return res.status(409).json({error:'The last active Administrator cannot be demoted'});
  }
  await pool.query('UPDATE app_users SET role=$1,updated_at=now() WHERE id=$2',[role,target]);
  await audit(req,'USER_ROLE_CHANGED','user',target,{role});
  res.json({ok:true});
});

app.delete('/api/users/:id', requireAuth, requirePermission('manage_users'), async (req,res) => {
  const target=req.params.id;
  if(target===req.user.id) return res.status(409).json({error:'You cannot delete your own account'});
  const {rows}=await pool.query('SELECT id,role,status FROM app_users WHERE id=$1',[target]);
  if(!rows[0]) return res.status(404).json({error:'User not found'});
  if(rows[0].role==='Administrator') {
    const {rows:a}=await pool.query("SELECT count(*)::int AS n FROM app_users WHERE role='Administrator' AND status='active'");
    if(rows[0].status==='active' && a[0].n<=1) return res.status(409).json({error:'The last active Administrator cannot be deleted'});
  }
  await pool.query('DELETE FROM app_users WHERE id=$1',[target]);
  await audit(req,'USER_DELETED','user',target);
  res.json({ok:true});
});

app.post('/api/auth/change-password', requireAuth, async (req,res) => {
  const current=String(req.body?.currentPassword||''), next=String(req.body?.newPassword||'');
  if(next.length<8 || !/[a-z]/.test(next) || !/[A-Z]/.test(next) || !/[0-9]/.test(next)) return res.status(400).json({error:'Password must be at least 8 characters and include upper, lower and number'});
  const {rows}=await pool.query('SELECT password_hash FROM app_users WHERE id=$1',[req.user.id]);
  if(!rows[0] || !(await bcrypt.compare(current,rows[0].password_hash))) return res.status(400).json({error:'Current password is incorrect'});
  await pool.query('UPDATE app_users SET password_hash=$1,must_change_password=false,updated_at=now() WHERE id=$2',[await bcrypt.hash(next,12),req.user.id]);
  await pool.query('UPDATE sessions SET revoked_at=now() WHERE user_id=$1 AND id<>$2',[req.user.id,req.user.session_id]);
  await audit(req,'PASSWORD_CHANGED','user',req.user.id);
  res.json({ok:true});
});

app.get('/api/audit-logs', requireAuth, requirePermission('view_audit_log'), async (req,res) => {
  const { rows } = await pool.query('SELECT a.id,a.action,a.entity,a.entity_id,a.created_at,u.full_name AS actor FROM audit_logs a LEFT JOIN app_users u ON u.id=a.actor_user_id ORDER BY a.created_at DESC LIMIT 500');
  res.json({logs:rows});
});

app.use((err, req, res, next) => { console.error(err); res.status(500).json({error:'Unexpected server error'}); });
app.listen(PORT, () => console.log(`MaliDesk auth server listening on ${PORT}`));
