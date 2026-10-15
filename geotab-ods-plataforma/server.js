const express = require('express');
const multer = require('multer');
const upload = multer({ dest: 'uploads/' });
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

// Lista de usuarios dinámica
let usuariosRegistrados = [
    { id: 1, usuario: 'cesar.admin', passwordHash: bcrypt.hashSync('admin123', 8), rol: 'admin' },
    { id: 2, usuario: 'tecnico.campo', passwordHash: bcrypt.hashSync('tec123', 8), rol: 'tecnico' }
];

let historialInstalaciones = [
    { fecha: '2026-10-12', eco: 'ECO-101', estatus: 'Óptimo', arnes: '16 Pines' },
    { fecha: '2026-10-13', eco: 'ECO-102', estatus: 'Óptimo', arnes: '16 Pines' }
];

// Almacenamiento autónomo en la memoria del servidor
let listaServicios = [
    { id: 1, folio: "ODS-1001", unidad: "Unit-01", operador: "Cesar", detalle: "Mantenimiento preventivo", salud: "Óptima", kpi: "98%", fecha: "2026-10-15" }
];

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

// --- CREAR USUARIO DESDE EL PANEL (Solo Admin) ---
app.post('/api/usuarios/crear', verificarAdmin, (req, res) => {
    const { nuevoUsuario, nuevaPassword, rol } = req.body;

    if (!nuevoUsuario || !nuevaPassword || !rol) {
        return res.status(400).json({ exito: false, error: 'Faltan datos obligatorios para crear el usuario.' });
    }

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

// --- RUTA 2: GENERAR ODS (Formato PDF Autónomo y Profesional) ---
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

        const publicDir = path.join(__dirname, 'public');
        if (!fs.existsSync(publicDir)) fs.mkdirSync(publicDir);

        const outputFilename = `ODS_${folioAcuerdo}_${Date.now()}.pdf`;
        const outputPath = path.join(publicDir, outputFilename);

        // Generar un PDF estructurado, limpio y profesional directamente
        const PDFDocument = require('pdfkit');
        const doc = new PDFDocument({ margin: 50 });
        const stream = fs.createWriteStream(outputPath);
        doc.pipe(stream);

        // Encabezado institucional
        doc.fontSize(20).fillColor('#1E3A8A').text('GEOTAB ODS - ORDEN DE SERVICIO', { align: 'center' });
        doc.fontSize(10.5).fillColor('#4B5563').text('Plataforma Autónoma de Gestión de Flota y Mantenimiento', { align: 'center' });
        doc.moveDown(1.5);

        // Información General
        doc.fontSize(14).fillColor('#111827').text('1. Datos Generales de la Instalación', { underline: true });
        doc.fontSize(11).fillColor('#374151');
        doc.text(`• Folio de Acuerdo: ${folioAcuerdo}`);
        doc.text(`• Fecha de Emisión: ${new Date().toISOString().split('T')[0]}`);
        doc.text(`• Proveedor: ADMINISTRACION DE FLOTAS`);
        doc.text(`• Instalador Clave: CS142W`);
        doc.moveDown(1);

        // Datos del Hardware
        doc.fontSize(14).fillColor('#111827').text('2. Componentes Registrados (Hardware)', { underline: true });
        doc.fontSize(11).fillColor('#374151');
        doc.text(`• Número de Serie GO9: ${serialGo9 || 'N/A'}`);
        doc.text(`• IMEI Cámara: ${imeiCamara || 'N/A'}`);
        doc.text(`• Número de SIM: ${numeroSim || 'N/A'}`);
        doc.text(`• VIN del Vehículo: ${vinVehiculo || 'N/A'}`);
        doc.moveDown(1);

        // Diagnóstico y Comentarios
        doc.fontSize(14).fillColor('#111827').text('3. Diagnóstico y Estatus Técnico', { underline: true });
        doc.fontSize(11).fillColor('#374151');
        doc.text(`• Comentario Técnico: ${comentarioTecnico || 'Instalación completada satisfactoriamente.'}`);
        doc.text(`• Estatus General: Óptimo (100% Funcional)`);
        doc.moveDown(2);

        // Pie de página
        doc.fontSize(9).fillColor('#9CA3AF').text('Documento generado automáticamente por el sistema autónomo Geotab ODS.', { align: 'center' });

        doc.end();

        // Esperar a que termine de escribirse el archivo PDF
        await new Promise((resolve, reject) => {
            stream.on('finish', resolve);
            stream.on('error', reject);
        });

        // Registrar en la memoria autónoma del sistema
        historialInstalaciones.push({
            fecha: new Date().toISOString().split('T')[0],
            eco: vinVehiculo || `ECO-${folioAcuerdo}`,
            estatus: 'Óptimo',
            arnes: '16 Pines'
        });

        listaServicios.push({
            id: listaServicios.length + 1,
            folio: `ODS-${folioAcuerdo}`,
            unidad: vinVehiculo || `Unidad-${folioAcuerdo}`,
            operador: "Operador Campo",
            detalle: comentarioTecnico || "Instalación autónoma exitosa en PDF",
            salud: "Óptima",
            kpi: "100%",
            fecha: new Date().toISOString().split('T')[0]
        });

        const hostUrl = req.protocol + '://' + req.get('host');
        return res.json({
            exito: true,
            mensaje: '¡Orden en PDF generada y registrada con éxito!',
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

// --- APIS AUTÓNOMAS DE SERVICIOS ---
app.get('/api/servicios', (req, res) => {
    res.json({ success: true, servicios: listaServicios });
});

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

// Carga Masiva con Excel integrada a la memoria autónoma
app.post('/api/servicios/masivo', upload.single('file'), (req, res) => {
    try {
        if (!req.file) {
            return res.status(400).json({ success: false, error: "No se subió ningún archivo Excel." });
        }
        const workbook = xlsx.readFile(req.file.path);
        const sheetName = workbook.SheetNames[0];
        const data = xlsx.utils.sheet_to_json(workbook.Sheets[sheetName]);
        
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

        fs.unlinkSync(req.file.path);
        res.json({ success: true, message: `¡Carga masiva exitosa! Se añadieron ${data.length} servicios al sistema autónomo.` });
    } catch (error) {
        res.status(500).json({ success: false, error: error.message });
    }
});

// Exportar datos y KPIs a Excel
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

// Archivos estáticos y ruta principal
app.use('/public', express.static(path.join(__dirname, 'public')));
app.use(express.static(path.join(__dirname, 'public')));

app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Servidor activo en el puerto ${PORT}`));
