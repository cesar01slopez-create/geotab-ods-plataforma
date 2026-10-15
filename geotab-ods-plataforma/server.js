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

// --- RUTA: GENERAR ODS ---
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

        const templatePath = path.join(__dirname, 'plantilla_base.docx');
        const publicDir = path.join(__dirname, 'public');
        if (!fs.existsSync(publicDir)) fs.mkdirSync(publicDir);

        const outputFilename = `ODS_${folioAcuerdo}_${Date.now()}.docx`;
        const outputPath = path.join(publicDir, outputFilename);

        if (fs.existsSync(templatePath)) {
            const content = fs.readFileSync(templatePath, 'binary');
            const zip = new PizZip(content);
            const doc = new Docxtemplater(zip, { paragraphLoop: true, linebreaks: true });

            doc.setData({
                proveedor: 'ADMINISTRACION DE FLOTAS',
                instalador: 'CS142W',
                folio_acuerdo: folioAcuerdo,
                fecha_visita: new Date().toISOString().split('T')[0],
                cliente: req.body.cliente || 'CLIENTE GENERAL',
                razon_social: req.body.razonSocial || 'EMPRESA',
                lugar_atencion: req.body.lugarAtencion || 'MEXICO',
                recibe: req.body.recibe || 'PERSONAL',
                contacto: req.body.contacto || 'N/A',
                ciudad: req.body.ciudad || 'CDMX',
                estado: req.body.estado || 'CDMX',
                serie_go9: serialGo9 || '',
                imei_camara: imeiCamara || '',
                sim_camara: numeroSim || '',
                vin: vinVehiculo || '',
                comentario_tecnico: comentarioTecnico || 'Instalación sin novedad'
            });

            doc.render();
            const buf = doc.getZip().generate({ type: 'nodebuffer' });
            fs.writeFileSync(outputPath, buf);
        } else {
            const zip = new PizZip();
            zip.file("word/document.xml", `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>ODS - Folio: ${folioAcuerdo}</w:t></w:r></w:p></w:body></w:document>`);
            fs.writeFileSync(outputPath, zip.generate({ type: 'nodebuffer' }));
        }

        const hostUrl = req.protocol + '://' + req.get('host');
        const fileUrl = `${hostUrl}/${outputFilename}`;

        // Registrar en la lista de servicios con su URL de descarga
        listaServicios.push({
            id: listaServicios.length + 1,
            folio: `ODS-${folioAcuerdo}`,
            unidad: vinVehiculo || `Unidad-${folioAcuerdo}`,
            operador: "ARMANDO GUTIERREZ",
            detalle: comentarioTecnico || "Instalación completada",
            salud: "Óptima",
            kpi: "100%",
            fecha: new Date().toISOString().split('T')[0],
            urlDoc: fileUrl
        });

        return res.json({
            exito: true,
            mensaje: '¡Orden institucional generada y registrada con éxito!',
            descargaUrl: fileUrl
        });

    } catch (error) {
        return res.status(500).json({ exito: false, error: error.message });
    }
});

        // Registrar en memoria autónoma
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
            operador: "ARMANDO GUTIERREZ",
            detalle: comentarioTecnico || "Instalación completada",
            salud: "Óptima",
            kpi: "100%",
            fecha: new Date().toISOString().split('T')[0]
        });

        const hostUrl = req.protocol + '://' + req.get('host');
        return res.json({
            exito: true,
            mensaje: '¡Orden institucional generada y registrada con éxito!',
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
const hostUrl = req.protocol + '://' + req.get('host');
        const fileUrl = `${hostUrl}/${outputFilename}`;

        // Registrar en el historial para que aparezca en la tabla con su link real
        listaServicios.push({
            id: listaServicios.length + 1,
            folio: `ODS-${folioAcuerdo}`,
            unidad: vinVehiculo || `Unidad-${folioAcuerdo}`,
            operador: "ARMANDO GUTIERREZ",
            detalle: comentarioTecnico || "Instalación completada",
            salud: "Óptima",
            kpi: "100%",
            fecha: new Date().toISOString().split('T')[0],
            urlDoc: fileUrl
        });

        return res.json({
            exito: true,
            mensaje: '¡Orden institucional generada y registrada con éxito!',
            descargaUrl: fileUrl
        });

app.listen(PORT, () => console.log(`Servidor activo en el puerto ${PORT}`));
