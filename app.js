// ============================================
// TELEMEDICIÓN
// Aplicación de lectura de medidores
// ============================================

let ruta = [];
let lecturas = {};
let medidorSeleccionado = null;
let modoEdicion = false;

let rutaKey = "";
let nombreArchivoRuta = "";


// ============================================
// ELEMENTOS HTML
// ============================================

const excelFile = document.getElementById("excelFile");
const archivo = document.getElementById("archivo");

const total = document.getElementById("total");
const medidos = document.getElementById("medidos");
const pendientes = document.getElementById("pendientes");

const progreso = document.getElementById("progreso");
const porcentaje = document.getElementById("porcentaje");

const buscar = document.getElementById("buscar");
const resultados = document.getElementById("resultados");

const clienteCard = document.getElementById("clienteCard");
const datosCliente = document.getElementById("datosCliente");

const lectura = document.getElementById("lectura");
const guardar = document.getElementById("guardar");
const editar = document.getElementById("editar");

const respaldo = document.getElementById("respaldo");
const exportar = document.getElementById("exportar");


// ============================================
// CLAVES DE ALMACENAMIENTO
// ============================================

const CLAVE_RUTA = "telemedicion_ruta_actual";


// ============================================
// NORMALIZAR TEXTO
// ============================================

function normalizarTexto(texto) {

    return String(texto || "")
        .normalize("NFD")
        .replace(/[\u0300-\u036f]/g, "")
        .toLowerCase()
        .trim();

}


// ============================================
// NORMALIZAR MEDIDOR
// ============================================

function normalizarMedidor(valor) {

    if (valor === null || valor === undefined) {
        return "";
    }

    return String(valor)
        .trim()
        .replace(/\s+/g, "");

}


// ============================================
// SOLO NÚMEROS
// ============================================

function soloNumeros(valor) {

    return normalizarMedidor(valor)
        .replace(/\D/g, "");

}


// ============================================
// BUSCAR COLUMNA
// ============================================

function buscarColumna(fila, nombre) {

    const objetivo = normalizarTexto(nombre);

    return Object.keys(fila).find(columna => {

        const columnaNormalizada =
            normalizarTexto(columna);

        return (
            columnaNormalizada === objetivo ||
            columnaNormalizada.includes(objetivo)
        );

    });

}


// ============================================
// CREAR IDENTIFICADOR ÚNICO
// ============================================

function crearIdRegistro(indice) {

    return "registro_" + indice;

}


// ============================================
// CARGAR EXCEL
// ============================================

excelFile.addEventListener("change", function(event) {

    const archivoSeleccionado =
        event.target.files[0];

    if (!archivoSeleccionado) {
        return;
    }


    const lector = new FileReader();


    lector.onload = function(e) {

        try {

            const datos =
                new Uint8Array(e.target.result);


            const workbook =
                XLSX.read(datos, {
                    type: "array",
                    cellDates: false
                });


            const nombreHoja =
                workbook.SheetNames[0];


            const hoja =
                workbook.Sheets[nombreHoja];


            const datosExcel =
                XLSX.utils.sheet_to_json(
                    hoja,
                    {
                        defval: "",
                        raw: false
                    }
                );


            if (datosExcel.length === 0) {

                alert(
                    "El Excel no contiene registros."
                );

                return;
            }


            // Buscar columna Medidor
            const columnaMedidor =
                buscarColumna(
                    datosExcel[0],
                    "medidor"
                );


            if (!columnaMedidor) {

                alert(
                    "No se encontró la columna 'Medidor'."
                );

                return;
            }


            // Buscar columna Secuencia
            const columnaSecuencia =
                buscarColumna(
                    datosExcel[0],
                    "secuencia"
                );


            // ========================================
            // CREAR RUTA
            // ========================================

            ruta = datosExcel.map(
                (fila, indice) => {

                    const medidor =
                        normalizarMedidor(
                            fila[columnaMedidor]
                        );


                    const secuencia =
                        columnaSecuencia
                            ? fila[columnaSecuencia]
                            : indice + 1;


                    return {

                        // ID ÚNICO REAL
                        id:
                            crearIdRegistro(indice),

                        fila:
                            fila,

                        medidor:
                            medidor,

                        medidorBusqueda:
                            soloNumeros(medidor),

                        secuencia:
                            secuencia

                    };

                }
            );


            // ========================================
            // IDENTIFICADOR DE LA RUTA
            // ========================================

            rutaKey =
                "telemedicion_" +
                archivoSeleccionado.name +
                "_" +
                archivoSeleccionado.size;


            nombreArchivoRuta =
                archivoSeleccionado.name;


            // ========================================
            // RECUPERAR LECTURAS DE ESTA RUTA
            // ========================================

            cargarLecturas();


            // ========================================
            // GUARDAR LA RUTA COMPLETA
            // ========================================

            guardarRutaCompleta();


            // ========================================
            // ACTUALIZAR INTERFAZ
            // ========================================

            archivo.textContent =
                "Ruta cargada: " +
                nombreArchivoRuta;


            buscar.value = "";

            resultados.innerHTML = "";

            clienteCard.classList.add(
                "oculto"
            );

            medidorSeleccionado = null;

            modoEdicion = false;


            actualizarProgreso();


            alert(
                "Ruta cargada correctamente.\n\n" +
                "Registros: " +
                ruta.length +
                "\n\n" +
                "La ruta quedó guardada en el teléfono."
            );


        } catch (error) {

            console.error(error);

            alert(
                "No se pudo leer el Excel."
            );

        }

    };


    lector.readAsArrayBuffer(
        archivoSeleccionado
    );

});


// ============================================
// GUARDAR RUTA COMPLETA
// ============================================

function guardarRutaCompleta() {

    try {

        const datosRuta = {

            nombreArchivo:
                nombreArchivoRuta,

            rutaKey:
                rutaKey,

            ruta:
                ruta,

            lecturas:
                lecturas,

            fechaGuardado:
                new Date()
                    .toLocaleString("es-EC")

        };


        localStorage.setItem(
            CLAVE_RUTA,
            JSON.stringify(datosRuta)
        );


        return true;


    } catch (error) {

        console.error(
            "Error guardando la ruta",
            error
        );


        alert(
            "⚠️ No se pudo guardar la ruta completa."
        );


        return false;

    }

}


// ============================================
// RECUPERAR RUTA AUTOMÁTICAMENTE
// ============================================

function recuperarRuta() {

    try {

        const datosGuardados =
            localStorage.getItem(
                CLAVE_RUTA
            );


        if (!datosGuardados) {

            actualizarProgreso();

            return;

        }


        const datos =
            JSON.parse(
                datosGuardados
            );


        if (
            !datos.ruta ||
            !Array.isArray(datos.ruta) ||
            datos.ruta.length === 0
        ) {

            actualizarProgreso();

            return;

        }


        ruta =
            datos.ruta;


        lecturas =
            datos.lecturas || {};


        rutaKey =
            datos.rutaKey || "";


        nombreArchivoRuta =
            datos.nombreArchivo || "";


        archivo.textContent =
            "Ruta recuperada: " +
            nombreArchivoRuta;


        actualizarProgreso();


        console.log(
            "Ruta recuperada automáticamente:",
            ruta.length,
            "registros"
        );


    } catch (error) {

        console.error(
            "Error recuperando ruta",
            error
        );


        ruta = [];

        lecturas = {};

        actualizarProgreso();

    }

}


// ============================================
// CARGAR LECTURAS
// ============================================

function cargarLecturas() {

    lecturas = {};


    try {

        const principal =
            localStorage.getItem(
                rutaKey
            );


        if (principal) {

            lecturas =
                JSON.parse(
                    principal
                );

        }

    } catch (error) {

        console.error(
            "Error leyendo lecturas",
            error
        );

    }


    // Segunda copia
    try {

        const secundario =
            localStorage.getItem(
                rutaKey +
                "_backup"
            );


        if (secundario) {

            const respaldoGuardado =
                JSON.parse(
                    secundario
                );


            Object.keys(
                respaldoGuardado
            ).forEach(id => {

                if (!lecturas[id]) {

                    lecturas[id] =
                        respaldoGuardado[id];

                }

            });

        }

    } catch (error) {

        console.error(
            "Error leyendo respaldo secundario",
            error
        );

    }

}


// ============================================
// BUSCAR MEDIDOR
// ============================================

buscar.addEventListener(
    "input",
    function() {

        const texto =
            soloNumeros(
                buscar.value
            );


        resultados.innerHTML = "";

        clienteCard.classList.add(
            "oculto"
        );

        medidorSeleccionado = null;


        if (texto.length < 3) {
            return;
        }


        if (ruta.length === 0) {

            resultados.innerHTML = `
                <div class="resultado">
                    Primero cargue una ruta.
                </div>
            `;

            return;
        }


        const encontrados =
            ruta.filter(
                registro => {

                    return registro
                        .medidorBusqueda
                        .includes(texto);

                }
            );


        if (encontrados.length === 0) {

            resultados.innerHTML = `
                <div class="resultado pendiente">
                    No se encontraron medidores.
                </div>
            `;

            return;
        }


        encontrados.forEach(
            registro => {

                const medicion =
                    lecturas[
                        registro.id
                    ];


                const div =
                    document.createElement(
                        "div"
                    );


                div.className =
                    "resultado";


                let estado = "";


                if (medicion) {

                    estado = `
                        <div class="estado ya-medido">
                            ✓ YA MEDIDO<br>
                            Lectura:
                            ${medicion.lectura}
                        </div>
                    `;

                } else {

                    estado = `
                        <div class="estado pendiente">
                            Pendiente de lectura
                        </div>
                    `;

                }


                div.innerHTML = `

                    <strong>
                        Medidor:
                        ${registro.medidor}
                    </strong>

                    <small>
                        Secuencia:
                        ${registro.secuencia}
                    </small>

                    ${estado}

                    <button>
                        Seleccionar
                    </button>

                `;


                div.querySelector(
                    "button"
                ).addEventListener(
                    "click",
                    function() {

                        seleccionarMedidor(
                            registro
                        );

                    }
                );


                resultados.appendChild(
                    div
                );

            }
        );

    }
);


// ============================================
// SELECCIONAR MEDIDOR
// ============================================

function seleccionarMedidor(
    registro
) {

    medidorSeleccionado =
        registro;


    modoEdicion = false;


    resultados.innerHTML = "";


    clienteCard.classList.remove(
        "oculto"
    );


    datosCliente.innerHTML = `

        <strong>Medidor:</strong>
        ${registro.medidor}

        <br>

        <strong>Secuencia:</strong>
        ${registro.secuencia}

    `;


    const medicion =
        lecturas[
            registro.id
        ];


    if (medicion) {

        lectura.value =
            medicion.lectura;


        lectura.disabled =
            true;


        guardar.classList.add(
            "oculto"
        );


        editar.classList.remove(
            "oculto"
        );


        datosCliente.innerHTML += `

            <div class="estado ya-medido">

                ✓ ESTE MEDIDOR YA FUE MEDIDO

                <br>

                Lectura:
                ${medicion.lectura}

                <br>

                Guardado:
                ${medicion.fecha}

            </div>

        `;

    } else {

        lectura.value = "";

        lectura.disabled =
            false;


        guardar.classList.remove(
            "oculto"
        );


        editar.classList.add(
            "oculto"
        );


        guardar.textContent =
            "💾 Guardar lectura";


        datosCliente.innerHTML += `

            <div class="estado pendiente">

                Medidor pendiente de lectura

            </div>

        `;

    }


    lectura.focus();

}


// ============================================
// EDITAR
// ============================================

editar.addEventListener(
    "click",
    function() {

        if (!medidorSeleccionado) {
            return;
        }


        modoEdicion = true;


        lectura.disabled =
            false;


        lectura.focus();


        guardar.classList.remove(
            "oculto"
        );


        guardar.textContent =
            "💾 Guardar cambio";


        editar.classList.add(
            "oculto"
        );


        datosCliente.innerHTML += `

            <div class="estado pendiente">

                Modo edición activado.
                Puede corregir la lectura.

            </div>

        `;

    }
);


// ============================================
// GUARDAR LECTURA
// ============================================

guardar.addEventListener(
    "click",
    function() {

        if (!medidorSeleccionado) {

            alert(
                "Seleccione un medidor."
            );

            return;
        }


        const valor =
            lectura.value.trim();


        if (valor === "") {

            alert(
                "Ingrese una lectura."
            );

            lectura.focus();

            return;
        }


        if (isNaN(valor)) {

            alert(
                "La lectura debe ser numérica."
            );

            lectura.focus();

            return;
        }


        const id =
            medidorSeleccionado.id;


        const yaExistia =
            lecturas[id];


        // Evitar sobrescritura accidental
        if (
            yaExistia &&
            !modoEdicion
        ) {

            alert(
                "Este medidor ya fue medido.\n\n" +
                "Presione 'Editar lectura' " +
                "para modificarlo."
            );

            return;
        }


        // ========================================
        // GUARDAR LECTURA
        // ========================================

        lecturas[id] = {

            lectura:
                valor,

            fecha:
                new Date()
                    .toLocaleString("es-EC")

        };


        // ========================================
        // GUARDADO INMEDIATO
        // ========================================

        guardarDatos();


        // Actualizar también la ruta completa
        guardarRutaCompleta();


        modoEdicion = false;


        lectura.disabled =
            true;


        guardar.classList.add(
            "oculto"
        );


        editar.classList.remove(
            "oculto"
        );


        datosCliente.innerHTML = `

            <strong>Medidor:</strong>
            ${medidorSeleccionado.medidor}

            <br>

            <strong>Secuencia:</strong>
            ${medidorSeleccionado.secuencia}

            <div class="estado ya-medido">

                ✓ LECTURA REGISTRADA

                <br>

                Lectura:
                ${valor}

                <br>

                ✓ Guardado automáticamente

            </div>

        `;


        actualizarProgreso();

    }
);


// ============================================
// GUARDAR LECTURAS EN DOS COPIAS
// ============================================

function guardarDatos() {

    if (!rutaKey) {
        return false;
    }


    const datos =
        JSON.stringify(
            lecturas
        );


    try {

        // Copia principal
        localStorage.setItem(
            rutaKey,
            datos
        );


        // Copia secundaria
        localStorage.setItem(
            rutaKey +
            "_backup",
            datos
        );


        return true;


    } catch (error) {

        console.error(
            "Error guardando lecturas",
            error
        );


        alert(
            "⚠️ No se pudo guardar la lectura.\n\n" +
            "Cree un respaldo manual."
        );


        return false;

    }

}


// ============================================
// CREAR RESPALDO MANUAL
// ============================================

respaldo.addEventListener(
    "click",
    function() {

        if (ruta.length === 0) {

            alert(
                "Primero cargue una ruta."
            );

            return;
        }


        const datosRespaldo =
            ruta.map(
                registro => {

                    const medicion =
                        lecturas[
                            registro.id
                        ];


                    return {

                        Secuencia:
                            registro.secuencia,

                        Medidor:
                            registro.medidor,

                        Lectura:
                            medicion
                                ? medicion.lectura
                                : "",

                        Estado:
                            medicion
                                ? "MEDIDO"
                                : "PENDIENTE",

                        Fecha:
                            medicion
                                ? medicion.fecha
                                : ""

                    };

                }
            );


        const hoja =
            XLSX.utils.json_to_sheet(
                datosRespaldo
            );


        const libro =
            XLSX.utils.book_new();


        XLSX.utils.book_append_sheet(
            libro,
            hoja,
            "Respaldo"
        );


        const fecha =
            new Date()
                .toISOString()
                .slice(0, 10);


        const nombreBase =
            obtenerNombreSinExtension(
                nombreArchivoRuta
            );


        const nombre =
            nombreBase +
            "_RESPALDO_" +
            fecha +
            ".xlsx";


        XLSX.writeFile(
            libro,
            nombre
        );


        alert(
            "✓ Respaldo creado correctamente.\n\n" +
            "Registros: " +
            ruta.length
        );

    }
);


// ============================================
// EXPORTAR EXCEL FINAL
// ============================================

exportar.addEventListener(
    "click",
    function() {

        if (ruta.length === 0) {

            alert(
                "Primero cargue una ruta."
            );

            return;
        }


        const datosExportar =
            ruta.map(
                registro => {

                    // Copiar TODA la fila original
                    const fila = {
                        ...registro.fila
                    };


                    const medicion =
                        lecturas[
                            registro.id
                        ];


                    fila[
                        "Lectura Registrada"
                    ] =
                        medicion
                            ? medicion.lectura
                            : "";


                    fila[
                        "Estado"
                    ] =
                        medicion
                            ? "MEDIDO"
                            : "PENDIENTE";


                    fila[
                        "Fecha lectura"
                    ] =
                        medicion
                            ? medicion.fecha
                            : "";


                    return fila;

                }
            );


        const hoja =
            XLSX.utils.json_to_sheet(
                datosExportar
            );


        const libro =
            XLSX.utils.book_new();


        XLSX.utils.book_append_sheet(
            libro,
            hoja,
            "Lecturas"
        );


        // Nombre de la ruta
        const nombreBase =
            obtenerNombreSinExtension(
                nombreArchivoRuta
            );


        const nombreFinal =
            nombreBase +
            "_FINAL.xlsx";


        XLSX.writeFile(
            libro,
            nombreFinal
        );


        alert(
            "✓ Excel final generado correctamente.\n\n" +
            "Registros: " +
            ruta.length
        );

    }
);


// ============================================
// OBTENER NOMBRE SIN .XLSX
// ============================================

function obtenerNombreSinExtension(
    nombre
) {

    if (!nombre) {
        return "ruta";
    }


    return nombre.replace(
        /\.(xlsx|xls)$/i,
        ""
    );

}


// ============================================
// ACTUALIZAR PROGRESO
// ============================================

function actualizarProgreso() {

    const totalRegistros =
        ruta.length;


    const totalMedidos =
        Object.keys(lecturas)
            .filter(
                id => {

                    return ruta.some(
                        registro =>
                            registro.id === id
                    );

                }
            )
            .length;


    const totalPendientes =
        totalRegistros -
        totalMedidos;


    total.textContent =
        totalRegistros;


    medidos.textContent =
        totalMedidos;


    pendientes.textContent =
        totalPendientes;


    let porcentajeActual =
        0;


    if (totalRegistros > 0) {

        porcentajeActual =
            (
                totalMedidos /
                totalRegistros
            ) * 100;

    }


    progreso.style.width =
        porcentajeActual + "%";


    porcentaje.textContent =
        porcentajeActual.toFixed(1) +
        "%";

}


// ============================================
// ACTIVAR MODO OFFLINE
// ============================================

if ("serviceWorker" in navigator) {

    window.addEventListener(
        "load",
        function() {

            navigator.serviceWorker
                .register("./sw.js")
                .then(
                    function() {

                        console.log(
                            "Modo offline activado."
                        );

                    }
                )
                .catch(
                    function(error) {

                        console.error(
                            "Error activando modo offline:",
                            error
                        );

                    }
                );

        }
    );

}


// ============================================
// RECUPERAR RUTA AL ABRIR LA APLICACIÓN
// ============================================

recuperarRuta();