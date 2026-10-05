
    // CONFIGURACIÓN DE WORKER DE PDF.JS
    pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/2.16.105/pdf.worker.min.js';

    // COMPRUEBA LAS LIBRERIAS AL CARGAR
    window.addEventListener('DOMContentLoaded', () => {
        const estadoDiv = document.getElementById('estado');
        const mammothCargado = typeof mammoth !== 'undefined';
        const pdfjsCargado = typeof pdfjsLib !== 'undefined';
        const xlsxCargado = typeof XLSX !== 'undefined';

        if (mammothCargado && pdfjsCargado && xlsxCargado) {
            estadoDiv.textContent = "✓ Librerías listas para procesar Word, PDF y Excel.";
            estadoDiv.style.color = "#28a745";
        } else {
            estadoDiv.textContent = "✗ Error al cargar una o más librerías. Revisa tu conexión a internet.";
            estadoDiv.style.color = "#d9534f";
        }
    });

    // PROCESAMIENTO DEL DOCUMENTO   - evidentemente todos los formatos juntos con la validación
    document.getElementById('archivo-doc').addEventListener('change', function(evento) {
        if (!evento.target.files || evento.target.files.length === 0) {
            return; 
        }
        //este es el selector
        const archivoSeleccionado = evento.target.files[0];
    
        const visor = document.getElementById('visor-contenido');
        const nombreArchivo = archivoSeleccionado.name.toLowerCase();
        
        visor.innerHTML = "<em>Procesando y convirtiendo documento...</em>";

        // PROCESA UN ARCHIVO PDF
        if (archivoSeleccionado.type === "application/pdf" || nombreArchivo.endsWith(".pdf")) {
            procesarPDF(archivoSeleccionado, visor);
        } 
        // PROCESAR ARCHIVO WORD (.DOCX)
        else if (nombreArchivo.endsWith(".docx")) {
            procesarWord(archivoSeleccionado, visor);
        } 
        // PROCESAR ARCHIVO EXCEL (.XLSX)
        else if (nombreArchivo.endsWith(".xlsx")) {
            procesarExcel(archivoSeleccionado, visor);
        }
        else if (nombreArchivo.endsWith(".xls")) {
            procesarExcel(archivoSeleccionado, visor);
        }
        else {
            visor.innerHTML = "<span style='color:red;'> Formato no soportado. Por favor, sube un archivo .docx, .pdf o .xlsx, xls.</span>";
        }
    });

    // FUNCIÓN PARA PROCESAR PDF
    function procesarPDF(archivo, visor) {
        const lector = new FileReader();

        lector.onload = function(e) {
            const typedarray = new Uint8Array(e.target.result);

            pdfjsLib.getDocument(typedarray).promise.then(async function(pdf) {
                visor.innerHTML = ""; 

                for (let numPagina = 1; numPagina <= pdf.numPages; numPagina++) {
                    try {
                        const pagina = await pdf.getPage(numPagina);
                        const canvas = document.createElement('canvas');
                        canvas.className = 'pagina-pdf';
                        const contexto = canvas.getContext('2d');
                        
                        const escala = 1.5; 
                        const viewport = pagina.getViewport({ scale: escala });

                        canvas.height = viewport.height;
                        canvas.width = viewport.width;

                        visor.appendChild(canvas);

                        await pagina.render({
                            canvasContext: contexto,
                            viewport: viewport
                        }).promise;
                    } catch (err) {
                        console.error("Error en la página " + numPagina, err);
                    }
                }
            }).catch(function(error) {
                console.error(error);
                visor.innerHTML = "<span style='color:red;'>Error al procesar el archivo PDF. Es posible que esté protegido o dañado.</span>";
            });
        };

        lector.readAsArrayBuffer(archivo);
    }

    // FUNCIÓN PARA RENDERIZAR WORD CON MAMMOTH.JS
    function procesarWord(archivo, visor) {
        const lector = new FileReader();

        lector.onload = function(eventoLectura) {
            const arrayBuffer = eventoLectura.target.result;

            mammoth.convertToHtml({ arrayBuffer: arrayBuffer })
                .then(function(resultado) {
                    if (resultado.value.trim() === "") {
                        visor.innerHTML = "<span style='color:orange;'>El documento está vacío o no contiene texto compatible.</span>";
                    } else {
                        visor.innerHTML = resultado.value;
                    }
                })
                .catch(function(error) {
                    console.error(error);
                    visor.innerHTML = "<span style='color:red;'>Error crítico al procesar la estructura del documento Word.</span>";
                });
        };

        lector.onerror = function() {
            visor.innerHTML = "<span style='color:red;'>Error físico al leer el archivo desde el disco duro.</span>";
        };

        lector.readAsArrayBuffer(archivo);
    }

    // FUNCIÓN PARA PROCESAR EXCEL CON SHEETJS

function procesarExcel(archivo, visor) {
    const lector = new FileReader();

    // 1. IMPORTANTE: Reconstruimos el evento onload que se había borrado
    lector.onload = function(e) {
        try {
            const datos = new Uint8Array(e.target.result);
            const libroTrabajo = XLSX.read(datos, { type: 'array' });
            
            visor.innerHTML = ""; // Limpiar visor el mensaje de "Procesando..."

            // Recorrer todas las hojas del libro de Excel
            libroTrabajo.SheetNames.forEach(function(nombreHoja) {
                try {
                    const hoja = libroTrabajo.Sheets[nombreHoja];
                    
                    // Si la hoja no tiene celdas o rango (!ref) definido,
                    // le creamos uno por defecto para que XLSX.utils.sheet_to_html no tire el error 'indexOf'
                    if (!hoja) return;
                    if (!hoja['!ref']) {
                        hoja['!ref'] = "A1:A1"; // Rango mínimo ficticio para evitar el colapso de la librería
                        hoja['A1'] = { t: 's', v: 'Hoja vacía o sin formato' };
                    }

                    // Convertir la hoja a formato HTML limpio
                    const htmlTabla = XLSX.utils.sheet_to_html(hoja);
                    
                    // Crear el contenedor de la pestaña actual
                    const contenedorHoja = document.createElement('div');
                    contenedorHoja.className = 'contenedor-pestaña-excel';
                    
                    // Título de la pestaña
                    const titulo = document.createElement('div');
                    titulo.className = 'titulo-hoja';
                    titulo.textContent = "Hoja: " + nombreHoja;
                    contenedorHoja.appendChild(titulo);
                    
                    // Contenedor intermedio para inyectar el HTML seguro
                    const wrapperTabla = document.createElement('div');
                    wrapperTabla.innerHTML = htmlTabla;
                    
                    // Forzar la clase de la tabla o enmendar si el string vino plano
                    const tabla = wrapperTabla.querySelector('table');
                    if (tabla) {
                        tabla.className = 'tabla-excel';
                    } else {
                        const tablaManual = document.createElement('table');
                        tablaManual.className = 'tabla-excel';
                        tablaManual.innerHTML = htmlTabla;
                        wrapperTabla.innerHTML = "";
                        wrapperTabla.appendChild(tablaManual);
                    }

                    contenedorHoja.appendChild(wrapperTabla);
                    visor.appendChild(contenedorHoja);

                } catch (errorHoja) {
                    console.error(`Error específico en la hoja ${nombreHoja}:`, errorHoja);
                    
                    // Si falla una sola hoja, mostramos un aviso elegante pero no rompemos el visor
                    const contenedorError = document.createElement('div');
                    contenedorError.innerHTML = `<p style="color:orange;">⚠️ No se pudo renderizar la hoja "${nombreHoja}".</p>`;
                    visor.appendChild(contenedorError);
                }
            });
        } catch (error) {
            console.error("Error crítico general en SheetJS:", error);
            visor.innerHTML = "<span style='color:red;'>Error al procesar el estructura del archivo Excel.</span>";
        }
    };

    lector.onerror = function() {
        visor.innerHTML = "<span style='color:red;'>Error al leer el archivo Excel desde el disco.</span>";
    };

    // 2. Ejecutar la lectura binaria
    lector.readAsArrayBuffer(archivo); 
}
 

 // Boton  
 
    document.addEventListener("DOMContentLoaded", function() {
    
    // 1. Traer los elementos con nombres claros y descriptivos
    const botonLimpiar = document.getElementById("clean-doc");
    const selectorArchivo = document.getElementById("archivo-doc");
    const visorContenido = document.getElementById("visor-contenido");

    // 2. Controlar la acción del botón de limpieza
    if (botonLimpiar) {
        botonLimpiar.addEventListener("click", function() {
            
            // A. Reset del estado del selector  
            if (selectorArchivo) {
                selectorArchivo.value = ""; // reseteo el string
            }

            // B. Limpieza y restauración visual del visor
            if (visorContenido) {
                // Volvemos a colocar el mensaje inicial o indicador de vacío
                visorContenido.innerHTML = '<p class="placeholder-text">VISUALIZAR DOCUMENTO...</p>';
            }

            console.log("Visor e input reiniciados al estado inicial con éxito.");
        });
    }
});
 