"""Genera las capturas del diálogo de correo a partir de src/Correo.html, con datos ficticios.

Uso, desde la raíz del repositorio:  python docs/capturas/generar.py
Requiere Microsoft Edge (o Chrome: cambie NAVEGADOR).
"""
import json
import pathlib
import subprocess

NAVEGADOR = r'C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe'
RAIZ = pathlib.Path(__file__).resolve().parents[2]
SALIDA = RAIZ / 'docs'

DATOS = {
    'codigo': 'ENV-2026-0007',
    'filas': [
        {'dni': '12345678', 'nombre': 'PÉREZ GÓMEZ, Juan', 'tipo': 'Certificado de Defunción',
         'pdf': 'PEREZ-GOMEZ-JUAN_001234_12345678_2026-11-18.pdf'},
        {'dni': '87654321', 'nombre': 'QUISPE MAMANI, Rosa', 'tipo': 'Informe de Defunción',
         'pdf': 'QUISPE-MAMANI-ROSA_005678_87654321_2026-11-20.pdf'},
    ],
    'personal': ['Personal 1', 'Personal 2'],
    'medios': ['WhatsApp', 'Correo institucional', 'Presencial'],
    'solicitantes': [{'nombre': 'Dra. Ana López', 'area': 'Medicina Interna', 'correo': 'ana.lopez@institucion.gob.pe'}],
    'cc': 'unidad@institucion.gob.pe; jefatura@institucion.gob.pe',
    'hoy': '2026-11-25',
}

STUB = ("<script>var R={withSuccessHandler:function(f){return R},withFailureHandler:function(){return R},"
        "crearBorrador:function(){},registrarEnvio:function(){}};"
        "var google={script:{host:{close:function(){}},run:R}};</script>")

LLENAR = ("el('sol').value='Dra. Ana López'; elegirSolicitante(); el('medio').value='WhatsApp';"
          "el('fecha').value='2026-11-24'; el('por').value='Personal 1'; armar();")

ESTADOS = {
    'dialogo-paso1.png': LLENAR,
    'dialogo-vista-previa.png': LLENAR + "el('vista').open=true;"
                                         "document.querySelector('main').scrollTop=170;",
    'dialogo-paso2.png': LLENAR + "pasarAPaso2('ok','Borrador creado ✓ Adjunte el PDF en Gmail, envíelo y registre aquí.');",
}

html = (RAIZ / 'src' / 'Correo.html').read_text(encoding='utf-8')
html = html.replace('<?!= datos ?>', json.dumps(DATOS, ensure_ascii=False)).replace('<script>', STUB + '<script>', 1)
html = html.replace('<body>', '<body style="padding:12px 16px 14px;box-sizing:border-box">')
def capturar(nombre, contenido, ancho, alto):
    pagina = SALIDA / 'capturas' / ('_' + nombre.replace('.png', '.html'))
    pagina.write_text(contenido, encoding='utf-8')
    subprocess.run([NAVEGADOR, '--headless=new', '--hide-scrollbars', '--force-device-scale-factor=2',
                    '--window-size=%d,%d' % (ancho, alto), '--screenshot=' + str(SALIDA / nombre), pagina.as_uri()],
                   check=True)
    pagina.unlink()
    print('generada', nombre)


for nombre, js in ESTADOS.items():
    capturar(nombre, html.replace('</body>', '<script>' + js + '</script></body>'), 744, 521)

# Diálogo "Copiar nombre del PDF": lo arma copiarNombrePdf() de Codigo.gs; se ejecuta con Node y datos ficticios.
NODE = """
const vm = require('vm'), fs = require('fs');
let html = '', alto = 0;
const salida = { setWidth() { return salida; }, setHeight(h) { alto = h; return salida; } };
const ctx = { SpreadsheetApp: { getUi: () => ({ showModalDialog() {} }) },
              HtmlService: { createHtmlOutput: h => { html = h; return salida; } } };
vm.createContext(ctx);
vm.runInContext(fs.readFileSync(process.argv[1], 'utf8'), ctx);
ctx.filas = JSON.parse(process.argv[2]);
vm.runInContext('filasSeleccionadas_ = () => filas; copiarNombrePdf();', ctx);
console.log(JSON.stringify({ html, alto }));
"""
r = json.loads(subprocess.run(['node', '-e', NODE, str(RAIZ / 'src' / 'Codigo.gs'), json.dumps(DATOS['filas'])],
                              capture_output=True, text=True, encoding='utf-8', check=True).stdout)
capturar('copiar-nombre-pdf.png',
         '<body style="margin:0;padding:12px 16px;box-sizing:border-box">' + r['html'] + '</body>', 600, r['alto'])
