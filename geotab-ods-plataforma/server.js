const express = require('express');
const multer = require('multer');
const path = require('path');
const fs = require('fs');
const PizZip = require('pizzip');
const Docxtemplater = require('docxtemplater');

const app = express();
const PORT = process.env.PORT || 3000;

// Configuración de middlewares
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

// Servir la carpeta 'public' en la raíz de los archivos estáticos
app.use(express.static(path.join(__dirname, 'public')));

// Configuración de Multer para archivos
const upload = multer({ dest: 'uploads/' });

// Memoria del sistema
let listaServicios = [
    {
        id: 1,
        folio: "ODS-1001",
        unidad: "Unit-01",
        operador: "Cesar",
        detalle: "Instalación completada",
        salud: "Óptima",
        kpi: "100%",
        fecha: new Date().toISOString().split('T')[0],
        urlDoc: "#"
    }
];

// --- RUTA: LOGIN ---
app.post('/api/login', (req, res) => {
    const { usuario, password } = req.body;
    if (!usuario || !password) {
        return res.status(400).json({ exito: false, error: 'Usuario y contraseña obligatorios.' });
    }
    const userLower = usuario.toLowerCase();
    const esTecnico = userLower.includes('tec') || userLower.includes('armando') || userLower.includes('instalador');
    return res.json({
        exito: true,
        mensaje: 'Sesión iniciada correctamente',
        rol: esTecnico ? 'tecnico' : 'coordinador'
    });
});

// --- RUTA: OBTENER SERVICIOS ---
app.get('/api/servicios', (req, res) => {
    return res.json({ success: true, servicios: listaServicios });
});

// --- RUTA: CREAR SERVICIO (Coordinador) ---
app.post('/api/servicios', (req, res) => {
    const { folioAcuerdo, tecnicoAsignado, cliente, ciudad } = req.body;
    if (!folioAcuerdo) {
        return res.status(400).json({ success: false, error: 'El folio es obligatorio' });
    }

    listaServicios.push({
        id: listaServicios.length + 1,
        folio: folioAcuerdo,
        unidad: cliente || 'Cliente General',
        operador: tecnicoAsignado || 'Técnico Campo',
        detalle: 'Asignado para instalación',
        salud: 'Pendiente',
        kpi: '0%',
        fecha: new Date().toISOString().split('T')[0],
        urlDoc: '#'
    });

    return res.json({ success: true, mensaje: 'Servicio creado exitosamente' });
});

// --- RUTA: REGISTRAR INSTALADOR ---
app.post('/api/instaladores', (req, res) => {
    const { nombre, usuario, password } = req.body;
    if (!nombre || !usuario) {
        return res.status(400).json({ exito: false, error: 'Datos incompletos' });
    }
    return res.json({ exito: true, mensaje: 'Instalador registrado con éxito' });
});

// --- RUTA: GENERAR ODS (Word Oficial) ---
app.post('/api/generar-ods', upload.fields([
    { name: 'foto_adherible', maxCount: 1 },
    { name: 'foto_imei_camara', maxCount: 1 },
    { name: 'foto_vin', maxCount: 1 },
    { name: 'foto_tablero', maxCount: 1 },
    { name: 'foto_arnes', maxCount: 1 }
]), async (req, res) => {
    try {
        const { folioAcuerdo, serialGo9, imeiCamara, numeroSim, vinVehiculo, comentarioTecnico } = req.body;
        if (!folioAcuerdo) {
            return res.status(400).json({ exito: false, error: 'El folio de acuerdo es obligatorio.' });
        }

        const templatePath = path.join(__dirname, 'plantilla_base.docx');
        const publicDir = path.join(__dirname, 'public');
        if (!fs.existsSync(publicDir)) {
            fs.mkdirSync(publicDir);
        }

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
        // AQUÍ ESTABA EL ERROR: No debemos agregar '/public/' a la URL final
        const fileUrl = `${hostUrl}/${outputFilename}`;

        // Guardar o actualizar en lista
        const index = listaServicios.findIndex(s => s.folio.includes(folioAcuerdo));
        if (index !== -1) {
            listaServicios[index].urlDoc = fileUrl;
            listaServicios[index].salud = "Completado";
        } else {
            listaServicios.push({
                id: listaServicios.length + 1,
                folio: `ODS-${folioAcuerdo}`,
                unidad: vinVehiculo || `Unidad-${folioAcuerdo}`,
                operador: "ARMANDO GUTIERREZ",
                detalle: comentarioTecnico || "Instalación completada",
                salud: "Completado",
                kpi: "100%",
                fecha: new Date().toISOString().split('T')[0],
                urlDoc: fileUrl
            });
        }

        return res.json({
            exito: true,
            mensaje: '¡Orden institucional generada y registrada con éxito!',
            descargaUrl: fileUrl
        });

    } catch (error) {
        return res.status(500).json({ exito: false, error: error.message });
    }
});

// Inicializar servidor
app.listen(PORT, () => {
    console.log(`Servidor corriendo en el puerto ${PORT}`);
});
