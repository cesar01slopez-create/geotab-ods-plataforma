const express = require('express');
const multer = require('multer');
const xlsx = require('xlsx');
const PizZip = require('pizzip');
const Docxtemplater = require('docxtemplater');
const fs = require('fs');
const path = require('path');
const cors = require('cors');
const jwt = require('jsonwebtoken');
const bcrypt = require('bcryptjs');

const app = express();
app.use(express.json());
app.use(cors());

const JWT_SECRET = process.env.JWT_SECRET || 'clave_secreta_super_segura_geotab';

// Lista de usuarios dinámica (ahora permite registrar nuevos desde el panel)
let usuariosRegistrados = [
    { id: 1, usuario: 'cesar.admin', passwordHash: bcrypt.hashSync('admin123', 8), rol: 'admin' },
    { id: 2, usuario: 'tecnico.campo', passwordHash: bcrypt.hashSync('tec123', 8), rol: 'tecnico' }
];

let historialInstalaciones = [
    { fecha: '2026-10-12', eco: 'ECO-101', estatus: 'Óptimo', arnes: '16 Pines' },
    { fecha: '2026-10-13', eco: 'ECO-102', estatus: 'Óptimo', arnes: '16 Pines' }
];

const upload = multer({ storage: multer.memoryStorage() });

// --- RUTA 1: LOGIN ---
app.post('/api/login', (req, res) => {
    const { usuario, password } = req.body;
    const user = usuariosRegistrados.find(u => u.usuario === usuario);

    if (!user || !bcrypt.compareSync(password, user.passwordHash)) {
        return res.status(401).json({ exito: false, error: 'Usuario o contraseña incorrectos.' });
    }

    const token = jwt.sign({ id: user.id, usuario: user.usuario, rol: user.rol }, JWT_SECRET, { expiresIn: '8h' });
    res.json({ exito: true, token, rol: user.rol });
});

// Middleware de seguridad para Administradores
function verificarAdmin(req, res, next) {
    const authHeader = req.headers['authorization'];
    const token = authHeader && authHeader.split(' ')[1];
    if (!token) return res.status(403).json({ exito: false, error: 'Acceso denegado. Token requerido.' });

    jwt.verify(token, JWT_SECRET, (err, user) => {
        if (err) return res.status(403).json({ exito: false, error: 'Token inválido o expirado.' });
        if (user.rol !== 'admin') return res.status(403).json({ exito: false, error: 'Permisos insuficientes.' });
        req.user = user;
        next();
    });
}

// --- RUTA NUEVA: CREAR USUARIO DESDE EL PANEL (Solo Admin) ---
app.post('/api/usuarios/crear', verificarAdmin, (req, res) => {
    const { nuevoUsuario, nuevaPassword, rol } = req.body;

    if (!nuevoUsuario || !nuevaPassword || !rol) {
        return res.status(400).json({ exito: false, error: 'Faltan datos obligatorios para crear el usuario.' });
    }

    // Verificar si el usuario ya existe
    const existe = usuariosRegistrados.find(u => u.usuario === nuevoUsuario);
    if (existe) {
        return res.status(400).json({ exito: false, error: 'El nombre de usuario ya está registrado.' });
    }

    const nuevoId = usuariosRegistrados.length + 1;
    const passwordHash = bcrypt.hashSync(nuevaPassword, 8);

    usuariosRegistrados.push({
        id: nuevoId,
        usuario: nuevoUsuario,
        passwordHash,
        rol
    });

    res.json({ exito: true, mensaje: `Usuario '${nuevoUsuario}' creado correctamente como ${rol}.` });
});

// --- RUTA 2: GENERAR ODS ---
app.post('/api/generar-ods', upload.fields([
    { name: 'foto_adherible', maxCount: 1 },
    { name: 'foto_imei_camara', maxCount: 1 },
    { name: 'foto_vin', maxCount: 1 },
    { name: 'foto_tablero', maxCount: 1 },
    { name: 'foto_arnes', maxCount: 1 }
]), async (req, res) => {
    try {
        const { folioAcuerdo, serialGo9, imeiCamara, numeroSim, vinVehiculo, comentarioTecnico } = req.body;
        if (!folioAcuerdo) return res.status(400).json({ exito: false, error: 'El folio de acuerdo es obligatorio.' });

        const excelPath = path.join(__dirname, 'Plantilla_servicios_ATT.xlsx');
        if (!fs.existsSync(excelPath)) return res.status(500).json({ exito: false, error: 'Falta el archivo Excel institucional.' });

        const workbook = xlsx.readFile(excelPath);
        const rows = xlsx.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]]);
        const servicioData = rows.find(row => String(row['Folio acuerdo *']).trim() === String(folioAcuerdo).trim());

        if (!servicioData) return res.status(404).json({ exito: false, error: `No se encontró el folio ${folioAcuerdo} en el Excel.` });

        const templatePath = path.join(__dirname, 'plantilla_base.docx');
        if (!fs.existsSync(templatePath)) return res.status(500).json({ exito: false, error: 'Falta la plantilla base Word.' });

        const content = fs.readFileSync(templatePath, 'binary');
        const zip = new PizZip(content);
        const doc = new Docxtemplater(zip, { paragraphLoop: true, linebreaks: true });

        doc.setData({
            proveedor: servicioData['Proveedor'] || 'ADMINISTRACION DE FLOTAS',
            instalador: servicioData['Instalador (clave)'] || 'CS142W',
            folio_acuerdo: folioAcuerdo,
            fecha_visita: servicioData['Fecha de visita * (dd/mm/aaaa)'] || '',
            cliente: servicioData['Nombre del cliente *'] || '',
            razon_social: servicioData['Razón social *'] || '',
            lugar_atencion: servicioData['Lugar de atención *'] || '',
            recibe: servicioData['Nombre de quien recibe al técnico *'] || '',
            contacto: servicioData['Números de contacto *'] || '',
            ciudad: servicioData['Ciudad *'] || '',
            estado: servicioData['Estado *'] || '',
            serie_go9: serialGo9 || '',
            imei_camara: imeiCamara || '',
            sim_camara: numeroSim || '',
            vin: vinVehiculo || '',
            comentario_tecnico: comentarioTecnico || 'Ok'
        });

        doc.render();

        const buf = doc.getZip().generate({ type: 'nodebuffer' });
        const outputFilename = `ODS_${folioAcuerdo}_${Date.now()}.docx`;
        const publicDir = path.join(__dirname, 'public');
        if (!fs.existsSync(publicDir)) fs.mkdirSync(publicDir);

        fs.writeFileSync(path.join(publicDir, outputFilename), buf);

        historialInstalaciones.push({
            fecha: new Date().toISOString().split('T')[0],
            eco: vinVehiculo || 'N/A',
            estatus: 'Óptimo',
            arnes: '16 Pines'
        });

        const hostUrl = req.protocol + '://' + req.get('host');
        return res.json({
            exito: true,
            mensaje: 'Orden generada y métricas actualizadas con éxito.',
            descargaUrl: `${hostUrl}/public/${outputFilename}`
        });

    } catch (error) {
        return res.status(500).json({ exito: false, error: error.message });
    }
});

// --- RUTA 3: OBTENER KPIS ---
app.get('/api/kpis', verificarAdmin, (req, res) => {
    const totalInstalaciones = historialInstalaciones.length;
    const optimas = historialInstalaciones.filter(i => i.estatus === 'Óptimo').length;
    const porcentajeSalud = totalInstalaciones > 0 ? Math.round((optimas / totalInstalaciones) * 100) : 100;

    res.json({
        totalInstalaciones,
        saludOptimaPct: porcentajeSalud,
        enRevision: historialInstalaciones.filter(i => i.estatus === 'Revisión').length,
        ciegas: 0,
        historial: historialInstalaciones
    });
});

app.use('/public', express.static(path.join(__dirname, 'public')));

const PORT = process.env.PORT || 3000;
// Servir la carpeta 'public' (donde está index.html)
app.use(express.static(path.join(__dirname, 'public')));

// Ruta principal que entrega la interfaz visual
app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

const multer = require('multer');
const upload = multer({ dest: 'uploads/' });

// 1. Ruta para Carga Masiva de Servicios vía Excel
app.post('/api/servicios/masivo', upload.single('file'), (req, res) => {
    try {
        if (!req.file) {
            return res.status(400).json({ success: false, error: "No se subió ningún archivo Excel." });
        }
        const workbook = xlsx.readFile(req.file.path);
        const sheetName = workbook.SheetNames[0];
        const data = xlsx.utils.sheet_to_json(workbook.Sheets[sheetName]);
        
        console.log(`[EXCEL MASIVO] Se procesaron ${data.length} registros.`);
        res.json({ success: true, message: `¡Carga masiva exitosa! Se procesaron ${data.length} servicios correctamente.` });
    } catch (error) {
        res.status(500).json({ success: false, error: error.message });
    }
});

// 2. Ruta para Exportar Datos, Salud y KPIs a Excel
app.get('/api/exportar/excel', (req, res) => {
    try {
        const datosSaludYKPIs = [
            { Folio: "ODS-1001", Unidad: "Unit-01", Operador: "Cesar", EstadoSalud: "Óptima", CumplimientoKPI: "98%", Fecha: "2026-10-15" }
        ];

        const worksheet = xlsx.utils.json_to_sheet(datosSaludYKPIs);
        const workbook = xlsx.utils.book_new();
        xlsx.utils.book_append_sheet(workbook, worksheet, "Salud_y_KPIs");

        const buffer = xlsx.write(workbook, { bookType: 'xlsx', type: 'buffer' });
        
        res.setHeader('Content-Disposition', 'attachment; filename="Reporte_Salud_KPIs_Geotab.xlsx"');
        res.type('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.send(buffer);
    } catch (error) {
        res.status(500).json({ success: false, error: error.message });
    }
});

// Almacenamiento autónomo en la memoria del servidor
let listaServicios = [
    { id: 1, folio: "ODS-1001", unidad: "Unit-01", operador: "Cesar", detalle: "Mantenimiento preventivo", salud: "Óptima", kpi: "98%", fecha: "2026-10-15" }
];

// 1. Obtener todos los servicios registrados
app.get('/api/servicios', (req, res) => {
    res.json({ success: true, servicios: listaServicios });
});

// 2. Registrar un servicio de forma manual y autónoma desde la plataforma
app.post('/api/servicios', express.json(), (req, res) => {
    try {
        const nuevoServicio = {
            id: listaServicios.length + 1,
            folio: `ODS-2026-${String(listaServicios.length + 1).padStart(3, '0')}`,
            unidad: req.body.unidad || "Sin unidad",
            operador: req.body.operador || "General",
            detalle: req.body.detalle || "Servicio general de flota",
            salud: req.body.salud || "Óptima",
            kpi: req.body.kpi || "95%",
            fecha: new Date().toISOString().split('T')[0]
        };
        
        listaServicios.push(nuevoServicio);
        res.json({ success: true, message: "¡Servicio registrado exitosamente!", servicio: nuevoServicio });
    } catch (error) {
        res.status(500).json({ success: false, error: error.message });
    }
});

// 3. Modificar la Carga Masiva para que guarde los datos en la lista autónoma
app.post('/api/servicios/masivo', upload.single('file'), (req, res) => {
    try {
        if (!req.file) {
            return res.status(400).json({ success: false, error: "No se subió ningún archivo Excel." });
        }
        const workbook = xlsx.readFile(req.file.path);
        const sheetName = workbook.SheetNames[0];
        const data = xlsx.utils.sheet_to_json(workbook.Sheets[sheetName]);
        
        // Agregar cada fila del Excel a nuestra lista autónoma
        data.forEach((item, index) => {
            listaServicios.push({
                id: listaServicios.length + 1,
                folio: item.Folio || `ODS-MASIVO-${index + 1}`,
                unidad: item.Unidad || "Desconocida",
                operador: item.Operador || "N/A",
                detalle: item.Detalle || item.Diagnostico || "Carga masiva Excel",
                salud: item.Salud || "Óptima",
                kpi: item.KPI || "90%",
                fecha: new Date().toISOString().split('T')[0]
            });
        });

        res.json({ success: true, message: `¡Carga masiva exitosa! Se añadieron ${data.length} servicios al sistema autónomo.` });
    } catch (error) {
        res.status(500).json({ success: false, error: error.message });
    }
});

// 4. Hacer que el reporte de Excel descargable use la información real autónoma
app.get('/api/exportar/excel', (req, res) => {
    try {
        const worksheet = xlsx.utils.json_to_sheet(listaServicios);
        const workbook = xlsx.utils.book_new();
        xlsx.utils.book_append_sheet(workbook, worksheet, "Flota_Servicios_KPIs");

        const buffer = xlsx.write(workbook, { bookType: 'xlsx', type: 'buffer' });
        
        res.setHeader('Content-Disposition', 'attachment; filename="Reporte_General_Geotab.xlsx"');
        res.type('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.send(buffer);
    } catch (error) {
        res.status(500).json({ success: false, error: error.message });
    }
});

app.listen(PORT, () => console.log(`Servidor activo en el puerto ${PORT}`));
