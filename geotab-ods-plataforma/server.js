const express = require('express'), cors = require('cors'), multer = require('multer'), bcrypt = require('bcryptjs'),
  jwt = require('jsonwebtoken'), XLSX = require('xlsx'), PDFDocument = require('pdfkit'), fs = require('fs'),
  path = require('path'), crypto = require('crypto');

const PORT = process.env.PORT || 3000;
const DIR = process.env.DATA_DIR || __dirname;          // En Render/Railway apunta a un disco persistente
const FILE = path.join(DIR, 'db.json'), ODS = path.join(DIR, 'ods');
const SECRET = process.env.JWT_SECRET || crypto.randomBytes(32).toString('hex');
fs.mkdirSync(ODS, { recursive: true });

let db = { usuarios: [], servicios: [] };
try { db = JSON.parse(fs.readFileSync(FILE, 'utf8')); } catch (e) {}
const save = () => fs.writeFileSync(FILE, JSON.stringify(db));
const t = v => String(v ?? '').trim(), nm = v => t(v).toUpperCase();

// Primer arranque: crea al coordinador
if (!db.usuarios.some(u => u.rol === 'coord')) {
  const pass = process.env.ADMIN_PASS || crypto.randomBytes(5).toString('hex');
  db.usuarios.push({ usuario: t(process.env.ADMIN_USER || 'cesar.admin').toLowerCase(), nombre: 'Cesar Sanchez', rol: 'coord', hash: bcrypt.hashSync(pass, 10) });
  save();
  if (!process.env.ADMIN_PASS) console.log('Contraseña inicial del coordinador:', pass);
}

const ok = (res, x = {}) => res.json({ exito: true, success: true, ...x });
const bad = (res, c, e) => res.status(c).json({ exito: false, success: false, error: e });
const auth = (req, res, next) => { try { req.u = jwt.verify((req.headers.authorization || '').slice(7), SECRET); next(); } catch (e) { bad(res, 401, 'Sesión no válida'); } };
const coord = (req, res, next) => req.u.rol === 'coord' ? next() : bad(res, 403, 'Solo el coordinador puede hacer esto');
const up = multer({ storage: multer.memoryStorage(), limits: { fileSize: 12 * 1024 * 1024 } });

const app = express();
app.use(cors()); app.use(express.json({ limit: '2mb' }));
app.use('/ods', express.static(ODS));
app.use(express.static(path.join(__dirname, 'public')));

// ---------- Sesión ----------
app.post('/api/login', (req, res) => {
  const u = db.usuarios.find(x => x.usuario === t(req.body.usuario).toLowerCase());
  if (!u || !bcrypt.compareSync(String(req.body.password || ''), u.hash)) return bad(res, 401, 'Usuario o contraseña incorrectos');
  const usuario = { usuario: u.usuario, nombre: u.nombre, rol: u.rol };
  ok(res, { token: jwt.sign(usuario, SECRET, { expiresIn: '12h' }), usuario });
});
app.get('/api/yo', auth, (req, res) => ok(res, { usuario: { usuario: req.u.usuario, nombre: req.u.nombre, rol: req.u.rol } }));

// ---------- Instaladores ----------
app.get('/api/instaladores', auth, coord, (req, res) => ok(res, { instaladores: db.usuarios.filter(u => u.rol === 'tec').map(({ hash, ...u }) => u) }));
app.post('/api/instaladores', auth, coord, (req, res) => {
  const usuario = t(req.body.usuario).toLowerCase(), nombre = t(req.body.nombre), pass = String(req.body.password || '');
  if (!usuario || !nombre || pass.length < 6) return bad(res, 400, 'Falta usuario, nombre o contraseña (mínimo 6 caracteres)');
  if (db.usuarios.some(u => u.usuario === usuario)) return bad(res, 409, 'Ese usuario ya existe');
  db.usuarios.push({ usuario, nombre, rol: 'tec', hash: bcrypt.hashSync(pass, 10) }); save(); ok(res);
});
app.put('/api/instaladores/:usuario', auth, coord, (req, res) => {
  const u = db.usuarios.find(x => x.usuario === req.params.usuario);
  if (!u || String(req.body.password || '').length < 6) return bad(res, 400, 'Usuario no encontrado o contraseña muy corta');
  u.hash = bcrypt.hashSync(req.body.password, 10); save(); ok(res);
});
app.delete('/api/instaladores/:usuario', auth, coord, (req, res) => {
  db.usuarios = db.usuarios.filter(u => !(u.usuario === req.params.usuario && u.rol === 'tec')); save(); ok(res);
});

// ---------- Servicios ----------
const vista = s => ({ ...s, unidad: s.cliente, operador: s.tecnico });
app.get('/api/servicios', auth, (req, res) => {
  const l = req.u.rol === 'coord' ? db.servicios : db.servicios.filter(s => nm(s.tecnico) === nm(req.u.nombre));
  ok(res, { servicios: l.slice().reverse().map(vista) });
});
app.post('/api/servicios', auth, coord, (req, res) => {
  const b = req.body, folio = t(b.folioAcuerdo);
  if (!folio || !t(b.tecnicoAsignado)) return bad(res, 400, 'Falta el folio o el técnico');
  if (db.servicios.some(s => nm(s.folio) === nm(folio))) return bad(res, 409, 'Ya existe ese folio');
  db.servicios.push({ folio, tecnico: t(b.tecnicoAsignado), cliente: t(b.cliente), razon: t(b.razonSocial), lugar: t(b.lugarAtencion),
    recibe: t(b.recibe), contacto: t(b.contacto), ciudad: t(b.ciudad), edo: t(b.estado), motivo: t(b.motivo) || 'Instalación',
    fecha: t(b.fecha) || new Date().toLocaleDateString('es-MX'), hora: t(b.hora), camara: t(b.camara) || 'No', accesorios: t(b.accesorios),
    proveedor: t(b.proveedor) || 'ADMINISTRACION DE FLOTAS', instalador: t(b.instalador) || 'CS142W', estatus: 'Enviado', urlDoc: '' });
  save(); ok(res);
});
app.delete('/api/servicios/:folio', auth, coord, (req, res) => {
  db.servicios = db.servicios.filter(s => nm(s.folio) !== nm(req.params.folio)); save(); ok(res);
});
// Alta masiva con tu Plantilla_servicios_ATT.xlsx
app.post('/api/servicios/importar', auth, coord, up.single('archivo'), (req, res) => {
  try {
    const wb = XLSX.read(req.file.buffer, { type: 'buffer' });
    const rows = XLSX.utils.sheet_to_json(wb.Sheets['Servicios'] || wb.Sheets[wb.SheetNames[0]], { header: 1, raw: false, defval: '' }).slice(1);
    let n = 0, dup = 0; const malas = [];
    rows.forEach((r, i) => {
      const [folio, tecnico, fecha, hora, motivo, camara, accesorios, cliente, razon, lugar, recibe, contacto, ciudad, edo, proveedor, instalador] = r.map(t);
      if (!folio || nm(folio) === 'EJEMPLO') return;
      if (![tecnico, fecha, hora, motivo, cliente, razon, lugar, recibe, contacto, ciudad, edo].every(Boolean)) return malas.push(i + 2);
      if (db.servicios.some(s => nm(s.folio) === nm(folio))) return dup++;
      db.servicios.push({ folio, tecnico, fecha, hora, motivo, camara: /^s/i.test(camara) ? 'Sí' : 'No', accesorios, cliente, razon, lugar, recibe, contacto,
        ciudad, edo, proveedor: proveedor || 'ADMINISTRACION DE FLOTAS', instalador: instalador || 'CS142W', estatus: 'Enviado', urlDoc: '' });
      n++;
    });
    save(); ok(res, { cargados: n, repetidos: dup, filasIncompletas: malas });
  } catch (e) { bad(res, 400, 'No se pudo leer el Excel'); }
});

// ---------- ODS en PDF ----------
const MOT = ['Instalación', 'Desinstalación', 'Reinstalación', 'Revisión', 'Reemplazo', 'Intercambio', 'Remover', 'Visita en falso'];
const CY = '#00A9E0';
function band(d, txt) {
  if (d.y > 690) d.addPage(); d.y += 6; const y = d.y;
  d.rect(40, y, 532, 18).fill(CY); d.fillColor('#fff').font('Helvetica-Bold').fontSize(10).text(txt, 40, y + 5, { width: 532, align: 'center' });
  d.fillColor('#111'); d.x = 40; d.y = y + 26;
}
function row(d, a) {
  if (d.y > 720) d.addPage(); const y = d.y, w = 532 / a.length; let e = y + 24;
  a.forEach(([l, v], i) => {
    const x = 40 + i * w;
    d.font('Helvetica').fontSize(7.5).fillColor('#666').text(l, x, y, { width: w - 8, lineBreak: false });
    d.fontSize(10).fillColor('#111').text(t(v) || ' ', x, y + 10, { width: w - 10 }); e = Math.max(e, d.y);
  });
  a.forEach((_, i) => { const x = 40 + i * (532 / a.length); d.moveTo(x, e + 1).lineTo(x + w - 8, e + 1).lineWidth(.4).strokeColor('#111').stroke(); });
  d.x = 40; d.y = e + 9;
}
function dataImg(u) { return u && u.startsWith('data:image') ? Buffer.from(u.split(',')[1], 'base64') : null; }

function renderODS(d, s, files, b) {
  d.font('Helvetica-Bold').fontSize(9).fillColor(CY).text('AT&T', 40, 40); d.fillColor('#111').font('Helvetica').fontSize(15).text('Orden de servicio de la flotilla', 40, 38, { width: 532, align: 'center' });
  d.y = 70; d.x = 40;
  row(d, [['Proveedor', s.proveedor], ['Folio acuerdo', s.folio]]); row(d, [['Instalador', s.instalador], ['Fecha de visita', `${s.fecha} ${s.hora || ''}`]]);
  band(d, 'MOTIVO DE LA VISITA');
  const y0 = d.y;
  MOT.forEach((m, i) => {
    const x = 40 + (i % 4) * 133, y = y0 + Math.floor(i / 4) * 20;
    d.rect(x, y, 11, 11).lineWidth(.8).strokeColor('#111').stroke();
    if (s.motivo === m) d.font('Helvetica-Bold').fontSize(10).fillColor('#111').text('X', x + 2.5, y + 1.5, { lineBreak: false });
    d.font('Helvetica').fontSize(9.5).fillColor('#111').text(m, x + 16, y + 1.5, { lineBreak: false });
  });
  d.x = 40; d.y = y0 + 46;
  row(d, [['Instalación accesorios', [s.camara === 'Sí' ? 'Cámara' : '', s.accesorios].filter(Boolean).join(', ')]]);
  band(d, 'DATOS GENERALES');
  [['Nombre del cliente', s.cliente], ['Razón social', s.razon], ['Lugar de atención', s.lugar], ['Nombre de quien recibe al técnico', s.recibe], ['Números de contacto', s.contacto], ['Ciudad', s.ciudad], ['Estado', s.edo]]
    .forEach(r => row(d, [r]));
  band(d, 'EQUIPO INSTALADO');
  row(d, [['Serie GO9', s.serie], ['IMEI cámara', s.imei]]); row(d, [['Arnés', s.arnes], ['Número de SIM', s.sim]]);
  row(d, [['Marca', s.marca], ['Modelo', s.modelo], ['Placas', s.placas]]); row(d, [['VIN', s.vin]]); row(d, [['Comentario del técnico / Observaciones', s.comentario]]);

  const fotos = [['Foto del adherible GO9', files.foto_adherible], ['Foto IMEI y SIM de cámara', files.foto_imei_camara], ['Foto del VIN', files.foto_vin]]
    .map(([l, a]) => [l, a && a[0]]).concat(Object.keys(files).filter(k => !['foto_adherible', 'foto_imei_camara', 'foto_vin'].includes(k))
      .flatMap(k => files[k].map((f, i) => [k === 'foto_tablero' ? 'Instalación ' + (i + 1) : k.replace(/^foto_/, '').replace(/_/g, ' '), f]))).filter(x => x[1]);
  if (fotos.length) {
    d.addPage(); band(d, 'EVIDENCIA FOTOGRÁFICA'); let col = 0, y = d.y;
    fotos.forEach(([l, f]) => {
      if (col === 0) { if (d.y > 540) d.addPage(); y = d.y; }
      const x = 40 + col * 272;
      try { d.image(f.buffer, x, y, { fit: [256, 175] }); } catch (e) { d.font('Helvetica').fontSize(8).fillColor('#a00').text('(formato de imagen no compatible, usa JPG o PNG)', x, y, { width: 256 }); }
      d.font('Helvetica').fontSize(8).fillColor('#555').text(l, x, y + 180, { width: 256 });
      if (++col === 2) { col = 0; d.y = y + 200; }
    });
    if (col) d.y = y + 200; d.x = 40;
  }
  if (d.y > 560) d.addPage(); band(d, 'FIRMAS'); const fy = d.y + 4;
  [[dataImg(b.firmaCliente), 'Nombre y firma de conformidad del cliente', b.nombreCliente || s.recibe], [dataImg(b.firmaTecnico), 'Nombre y firma del técnico', s.tecnico]].forEach(([img, l, n], i) => {
    const x = 40 + i * 280;
    if (img) try { d.image(img, x, fy, { fit: [240, 70] }); } catch (e) {}
    d.moveTo(x, fy + 76).lineTo(x + 240, fy + 76).lineWidth(.5).strokeColor('#111').stroke();
    d.font('Helvetica').fontSize(10).fillColor('#111').text(t(n), x, fy + 80, { width: 240 }); d.fontSize(7.5).fillColor('#666').text(l, x, fy + 94, { width: 240 });
  });
  d.x = 40; d.y = fy + 125;
  d.font('Helvetica').fontSize(7.5).fillColor('#555').text('Nota: Recuerda que antes de concluir tu visita, es necesario comunicarte a Call Center para que nuestro cliente responda la encuesta de satisfacción. Posteriormente debes compartir el estatus de tu visita al ejecutivo. CC: 55-7100-2854. Soporte Técnico CLIENTE comunicarse al 01 800 211 4012', 40, d.y, { width: 532 });
}

const campos = up.any();
const agrupa = l => (l || []).reduce((o, f) => ((o[f.fieldname] = o[f.fieldname] || []).push(f), o), {});
app.post('/api/generar-ods', auth, campos, async (req, res) => {
  try {
    const b = req.body, s = db.servicios.find(x => nm(x.folio) === nm(b.folioAcuerdo));
    if (!s) return bad(res, 404, 'Folio no encontrado');
    if (req.u.rol !== 'coord' && nm(s.tecnico) !== nm(req.u.nombre)) return bad(res, 403, 'Ese servicio no está asignado a ti');
    Object.assign(s, { serie: t(b.serialGo9), imei: t(b.imeiCamara), sim: t(b.numeroSim), arnes: t(b.arnes), marca: t(b.marcaVehiculo), modelo: t(b.modeloVehiculo),
      vin: t(b.vinVehiculo).toUpperCase(), placas: t(b.placasVehiculo).toUpperCase(), comentario: t(b.comentarioTecnico) });
    const name = `ODS_${s.folio.replace(/[^\w-]/g, '')}_${crypto.randomBytes(4).toString('hex')}.pdf`;
    const doc = new PDFDocument({ size: 'LETTER', margin: 40 }), out = fs.createWriteStream(path.join(ODS, name));
    doc.pipe(out); renderODS(doc, s, agrupa(req.files), b); doc.end();
    await new Promise((r, j) => { out.on('finish', r); out.on('error', j); });
    s.estatus = 'Completado'; s.urlDoc = '/ods/' + name; save();
    ok(res, { descargaUrl: s.urlDoc });
  } catch (e) { console.error(e); bad(res, 500, 'No se pudo generar el PDF: ' + e.message); }
});

app.listen(PORT, () => console.log('FLOTAS en puerto ' + PORT));
