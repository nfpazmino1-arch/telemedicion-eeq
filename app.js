// ============================================
// TELEMEDICIÓN EEQ
// Sistema de lectura con guardado automático
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


            const columnaSecuencia =
                buscarColumna(
                    datosExcel[0],
                    "secuencia"
                );


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

                        id: String(
                            secuencia ||
                            indice + 1
                        ),

                        fila: fila,

                        medidor: medidor,

                        medidorBusqueda:
                            soloNumeros(medidor),

                        secuencia:
                            secuencia

                    };

                }
            );


            // Identificador de la ruta
            rutaKey =
                "telemedicion_" +
                archivoSeleccionado.name +
                "_" +
                archivoSeleccionado.size;

            nombreArchivoRuta =
                archivoSeleccionado.name;


            // Recuperar lecturas guardadas
            cargarRespaldo();


            archivo.textContent =
                "Ruta cargada: " +
                archivoSeleccionado.name;


            buscar.value = "";

            resultados.innerHTML = "";

            clienteCard.classList.add("oculto");

            medidorSeleccionado = null;

            modoEdicion = false;


            actualizarProgreso();


            alert(
                "Ruta cargada correctamente.\n\n" +
                "Registros: " +
                ruta.length +
                "\n\n" +
                "Las lecturas guardadas anteriormente " +
                "se conservarán."
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
// CARGAR LECTURAS GUARDADAS
// ============================================

function cargarRespaldo() {

    lecturas = {};

    try {

        const principal =
            localStorage.getItem(rutaKey);

        if (principal) {

            lecturas =
                JSON.parse(principal);

        }

    } catch (error) {

        console.error(
            "Error leyendo respaldo principal",
            error
        );

    }


    // Recuperar segunda copia
    try {

        const secundario =
            localStorage.getItem(
                rutaKey + "_backup"
            );

        if (secundario) {

            const respaldoGuardado =
                JSON.parse(secundario);


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
            soloNumeros(buscar.value);

        resultados.innerHTML = "";

        clienteCard.classList.add("oculto");

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
            ruta.filter(registro => {

                return registro.medidorBusqueda
                    .includes(texto);

            });


        if (encontrados.length === 0) {

            resultados.innerHTML = `
                <div class="resultado pendiente">
                    No se encontraron medidores.
                </div>
            `;

            return;
        }


        encontrados.forEach(registro => {

            const medicion =
                lecturas[registro.id];

            const div =
                document.createElement("div");

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


            div.querySelector("button")
                .addEventListener(
                    "click",
                    function() {

                        seleccionarMedidor(
                            registro
                        );

                    }
                );


            resultados.appendChild(div);

        });

    }
);


// ============================================
// SELECCIONAR MEDIDOR
// ============================================

function seleccionarMedidor(registro) {

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
        lecturas[registro.id];


    if (medicion) {

        lectura.value =
            medicion.lectura;

        lectura.disabled = true;

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

        lectura.disabled = false;

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
// EDITAR LECTURA
// ============================================

editar.addEventListener(
    "click",
    function() {

        if (!medidorSeleccionado) {
            return;
        }

        modoEdicion = true;

        lectura.disabled = false;

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


        const yaExistia =
            lecturas[
                medidorSeleccionado.id
            ];


        // No permitir sobrescritura accidental
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


        // Registrar lectura
        lecturas[
            medidorSeleccionado.id
        ] = {

            lectura: valor,

            fecha:
                new Date()
                .toLocaleString(
                    "es-EC"
                )

        };


        // GUARDADO AUTOMÁTICO
        guardarDatos();


        modoEdicion = false;

        lectura.disabled = true;

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
// GUARDADO AUTOMÁTICO
// ============================================

function guardarDatos() {

    if (!rutaKey) {
        return false;
    }


    const datos =
        JSON.stringify(lecturas);


    try {

        // Primera copia
        localStorage.setItem(
            rutaKey,
            datos
        );


        // Segunda copia
        localStorage.setItem(
            rutaKey + "_backup",
            datos
        );


        return true;


    } catch (error) {

        console.error(
            "Error guardando datos",
            error
        );


        alert(
            "⚠️ No se pudo guardar la lectura.\n\n" +
            "Cree un respaldo manual inmediatamente."
        );


        return false;

    }

}


// ============================================
// BOTÓN: CREAR RESPALDO MANUAL
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
            ruta.map(registro => {

                const medicion =
                    lecturas[registro.id];


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

            });


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


        const nombre =
            "RESPALDO_" +
            fecha +
            "_" +
            nombreArchivoRuta;


        XLSX.writeFile(
            libro,
            nombre
        );


        alert(
            "✓ Respaldo creado correctamente."
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
            ruta.map(registro => {

                const fila = {
                    ...registro.fila
                };


                const medicion =
                    lecturas[registro.id];


                fila["Lectura Registrada"] =
                    medicion
                        ? medicion.lectura
                        : "";


                fila["Estado"] =
                    medicion
                        ? "MEDIDO"
                        : "PENDIENTE";


                fila["Fecha lectura"] =
                    medicion
                        ? medicion.fecha
                        : "";


                return fila;

            });


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


        XLSX.writeFile(
            libro,
            "lecturas_telemedicion.xlsx"
        );


        alert(
            "✓ Excel final generado correctamente."
        );

    }
);


// ============================================
// ACTUALIZAR PROGRESO
// ============================================

function actualizarProgreso() {

    const totalRegistros =
        ruta.length;


    const totalMedidos =
        Object.keys(lecturas)
            .filter(id => {

                return ruta.some(
                    registro =>
                        registro.id === id
                );

            })
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


    let porcentajeActual = 0;


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
// INICIO
// ============================================

actualizarProgreso();
// ============================================
// ACTIVAR MODO OFFLINE
// ============================================

if ("serviceWorker" in navigator) {

    window.addEventListener(
        "load",
        function() {

            navigator.serviceWorker
                .register("./sw.js")
                .then(function() {

                    console.log(
                        "Telemedición EEQ: modo offline activado."
                    );

                })
                .catch(function(error) {

                    console.error(
                        "Error activando modo offline:",
                        error
                    );

                });

        }
    );

}