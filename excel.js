let datosExcel = [];
let datosFiltrados = [];

// Los valores se rellenan solos al leer el Excel (ver resolverColumnas)
const columnas = {
    fecha: "Marca temporal",
    edad: "Edad",
    genero: "Género",
    carrera: "Carrera",
    prendas: "",
    presupuesto: "",
    marcaReconocida: "",
    precio: "",
    marcasFrecuentes: "",
    tipoRopa: "",
    criterio: "",
    lugarCompra: "",
    marcaPreferida: "",
    ropaNoUsa: ""
};

document.addEventListener("DOMContentLoaded", cargarExcel);

async function cargarExcel() {
    try {
        const respuesta = await fetch("datos.xlsx");

        if (!respuesta.ok) {
            throw new Error("No se pudo encontrar datos.xlsx");
        }

        const archivo = await respuesta.arrayBuffer();
        const libro = XLSX.read(archivo, { type: "array" });
        const hoja = libro.Sheets[libro.SheetNames[0]];

        datosExcel = XLSX.utils.sheet_to_json(hoja, { defval: "" });

        if (datosExcel.length === 0) {
            throw new Error("El Excel no tiene filas de datos");
        }

        resolverColumnas(datosExcel[0]);
        datosFiltrados = [...datosExcel];
        iniciarPagina();

    } catch (error) {
        console.error("Error al cargar Excel:", error);
        mostrarError(error);
    }
}

/* Busca cada columna por su encabezado real, sin importar tildes,
   mayúsculas ni el texto completo de la pregunta. */
function resolverColumnas(fila) {
    const encabezados = Object.keys(fila);

    const buscar = (inicio, porDefecto) => {
        const clave = normalizar(inicio);
        return encabezados.find(h => normalizar(h).startsWith(clave)) || porDefecto;
    };

    columnas.fecha = buscar("Marca temporal", columnas.fecha);
    columnas.edad = buscar("Edad", columnas.edad);
    columnas.genero = buscar("Genero", columnas.genero);
    columnas.carrera = buscar("Carrera", columnas.carrera);
    columnas.prendas = buscar("1.", columnas.prendas);
    columnas.presupuesto = buscar("2.", columnas.presupuesto);
    columnas.marcaReconocida = buscar("3.", columnas.marcaReconocida);
    columnas.precio = buscar("4.", columnas.precio);
    columnas.marcasFrecuentes = buscar("5.", columnas.marcasFrecuentes);
    columnas.tipoRopa = buscar("6.", columnas.tipoRopa);
    columnas.criterio = buscar("7.", columnas.criterio);
    columnas.lugarCompra = buscar("8.", columnas.lugarCompra);
    columnas.marcaPreferida = buscar("9.", columnas.marcaPreferida);
    columnas.ropaNoUsa = buscar("10.", columnas.ropaNoUsa);
}

function mostrarError(error) {
    const contenido = document.querySelector(".Contenido");
    if (!contenido) return;

    const aviso = document.createElement("div");
    aviso.style.cssText =
        "background:#ffebee;color:#b71c1c;padding:14px 18px;border-radius:8px;margin-bottom:20px;font-size:14px;";

    aviso.textContent =
        location.protocol === "file:"
            ? "No se pudo leer datos.xlsx porque la página se abrió con doble clic (file://). " +
              "Ábrela con un servidor local, por ejemplo la extensión Live Server de VS Code."
            : "Error al cargar los datos: " + error.message;

    contenido.prepend(aviso);
}

function iniciarPagina() {
    cargarFiltros();

    if (document.getElementById("cuerpoTabla")) {
        mostrarDatos();
    }

    if (document.getElementById("totalEncuestados")) {
        actualizarDashboard(datosFiltrados);
    }
}

/* ---------- FILTROS ---------- */

function cargarFiltros() {
    const filtroCarrera = document.getElementById("filtroCarrera");
    const filtroEdad = document.getElementById("filtroEdad");
    const filtroGenero = document.getElementById("filtroGenero");

    if (filtroCarrera) {
        llenarFiltro(filtroCarrera, columnas.carrera, "Todas", "Todas las Carreras");
        filtroCarrera.addEventListener("change", aplicarFiltros);
    }

    if (filtroEdad) {
        llenarFiltro(filtroEdad, columnas.edad, "Todas", "Todas las edades");
        filtroEdad.addEventListener("change", aplicarFiltros);
    }

    if (filtroGenero) {
        llenarFiltro(filtroGenero, columnas.genero, "Todos", "Todos los Géneros");
        filtroGenero.addEventListener("change", aplicarFiltros);
    }

    const botonLimpiar = document.getElementById("limpiarFiltros");

    if (botonLimpiar) {
        botonLimpiar.addEventListener("click", limpiarFiltros);
    }
}

function llenarFiltro(select, columna, valorInicial, textoInicial) {
    select.innerHTML = "";

    const primera = document.createElement("option");
    primera.value = valorInicial;
    primera.textContent = textoInicial;
    select.appendChild(primera);

    obtenerValoresUnicos(columna).forEach(valor => {
        const opcion = document.createElement("option");
        opcion.value = valor;
        opcion.textContent = valor;
        select.appendChild(opcion);
    });
}

function obtenerValoresUnicos(columna) {
    const valores = datosExcel
        .map(fila => String(fila[columna] ?? "").trim())
        .filter(valor => valor !== "");

    return [...new Set(valores)].sort((a, b) =>
        a.localeCompare(b, "es", { numeric: true })
    );
}

function valorFiltro(id, porDefecto) {
    const elemento = document.getElementById(id);
    return elemento ? elemento.value : porDefecto;
}

function aplicarFiltros() {
    const carrera = valorFiltro("filtroCarrera", "Todas");
    const edad = valorFiltro("filtroEdad", "Todas");
    const genero = valorFiltro("filtroGenero", "Todos");

    datosFiltrados = datosExcel.filter(fila => {
        const valorCarrera = String(fila[columnas.carrera] ?? "").trim();
        const valorEdad = String(fila[columnas.edad] ?? "").trim();
        const valorGenero = String(fila[columnas.genero] ?? "").trim();

        return (
            (carrera === "Todas" || valorCarrera === carrera) &&
            (edad === "Todas" || valorEdad === edad) &&
            (genero === "Todos" || valorGenero === genero)
        );
    });

    refrescarVista();
}

function limpiarFiltros() {
    const restablecer = (id, valor) => {
        const elemento = document.getElementById(id);
        if (elemento) elemento.value = valor;
    };

    restablecer("filtroCarrera", "Todas");
    restablecer("filtroEdad", "Todas");
    restablecer("filtroGenero", "Todos");

    datosFiltrados = [...datosExcel];
    refrescarVista();
}

function refrescarVista() {
    if (document.getElementById("cuerpoTabla")) {
        mostrarDatos();
    }

    if (document.getElementById("totalEncuestados")) {
        actualizarDashboard(datosFiltrados);
    }
}

/* ---------- TABLA (explorador.html) ---------- */

function mostrarDatos() {
    const cuerpo = document.getElementById("cuerpoTabla");

    if (!cuerpo) return;

    cuerpo.innerHTML = "";

    datosFiltrados.forEach(fila => {
        const tr = document.createElement("tr");

        const valores = [
            convertirFecha(fila[columnas.fecha]),
            fila[columnas.edad],
            fila[columnas.genero],
            fila[columnas.carrera],
            fila[columnas.prendas],
            fila[columnas.presupuesto],
            fila[columnas.marcaReconocida],
            fila[columnas.precio],
            fila[columnas.marcasFrecuentes],
            fila[columnas.tipoRopa],
            fila[columnas.criterio],
            fila[columnas.lugarCompra],
            fila[columnas.marcaPreferida],
            fila[columnas.ropaNoUsa]
        ];

        valores.forEach(valor => {
            const td = document.createElement("td");
            td.textContent = valor ?? "";
            tr.appendChild(td);
        });

        cuerpo.appendChild(tr);
    });

    actualizarContadores();
}

function convertirFecha(valor) {
    if (typeof valor === "number") {
        const fecha = XLSX.SSF.parse_date_code(valor);

        if (fecha) {
            const dos = n => String(n).padStart(2, "0");

            return `${dos(fecha.d)}/${dos(fecha.m)}/${fecha.y} ` +
                   `${dos(fecha.H)}:${dos(fecha.M)}:${dos(Math.floor(fecha.S))}`;
        }
    }

    return valor ?? "";
}

function actualizarContadores() {
    actualizarTexto("totalRegistros", datosExcel.length);
    actualizarTexto("registrosMostrados", datosFiltrados.length);
    actualizarTexto("cantidadResultados", datosFiltrados.length + " resultados");
}

/* ---------- DASHBOARD (prueba1.html) ---------- */

function actualizarDashboard(datos) {
    const total = datos.length;

    actualizarTexto("totalEncuestados", total);
    actualizarTexto("totalDona", total);
    actualizarTexto("respuestasTipo", total + " respuestas");
    actualizarTexto("respuestasCriterio", total + " respuestas");
    actualizarTexto("registrosAnalizados", total);

    // Tipo de ropa
    const sostenible = calcularPorcentaje(datos, columnas.tipoRopa, ["Moda sostenible"]);
    const noFijo = calcularPorcentaje(datos, columnas.tipoRopa, ["No me fijo"]);
    const rapida = calcularPorcentaje(datos, columnas.tipoRopa, ["Moda rapida"]);

    actualizarTexto("porcentajeSostenible", sostenible + "%");
    actualizarTexto("datoSostenible", sostenible + "%");
    actualizarTexto("datoNofijo", noFijo + "%");
    actualizarTexto("datoRapida", rapida + "%");

    actualizarDona(total, sostenible, noFijo);

    // Criterio principal
    const calidad = calcularPorcentaje(datos, columnas.criterio, ["Buena calidad"]);
    const diseno = calcularPorcentaje(datos, columnas.criterio, ["Diseño y estilo"]);
    const precio = calcularPorcentaje(datos, columnas.criterio, ["Precio bajo"]);

    actualizarTexto("porcentajeDiseno", diseno + "%");
    actualizarTexto("datoCalidad", calidad + "%");
    actualizarTexto("datoDiseno", diseno + "%");
    actualizarTexto("datoPrecio", precio + "%");

    actualizarBarra("barraCalidad", calidad);
    actualizarBarra("barraDiseno", diseno);
    actualizarBarra("barraPrecio", precio);

    // Presupuesto: % de encuestados que destina entre 5 % y 10 % a ropa
    const presupuesto = calcularPorcentaje(datos, columnas.presupuesto, [
        "Entre el 5"   // "Entre el 5 % y el 10 %"
    ]);

    actualizarTexto("porcentajePresupuesto", presupuesto + "%");

    actualizarPrendas(datos);
    actualizarMarcas(datos);
}

/* Cuenta las filas cuya respuesta CONTIENE alguna de las opciones
   (ignora tildes y mayúsculas), así "Moda sostenible" también
   coincide con "Moda sostenible (ecológica/duradera)". */
function calcularPorcentaje(datos, columna, opciones) {
    if (!datos.length) return 0;

    const buscadas = opciones.map(normalizar);
    let cantidad = 0;

    datos.forEach(fila => {
        const respuesta = normalizar(fila[columna]);

        if (buscadas.some(opcion => respuesta.includes(opcion))) {
            cantidad++;
        }
    });

    return redondear(cantidad / datos.length * 100);
}

function actualizarPrendas(datos) {
    const grupos = {
        prendasNinguna: ["Ninguna"],
        prendas12: ["1 a 2"],
        prendas35: ["3 a 5"],
        prendas68: ["6 a 8"],
        prendas9: ["Mas de 9"]
    };

    Object.keys(grupos).forEach(id => {
        const porcentaje = calcularPorcentaje(datos, columnas.prendas, grupos[id]);
        actualizarTexto(id, porcentaje + "%");
    });
}

function actualizarMarcas(datos) {
    const marcas = {
        marcaAdidas: ["Adidas"],
        marcaPuma: ["Puma"],
        marcaNike: ["Nike"],
        marcaZara: ["Zara"],
        marcaHM: ["H&M", "H & M"]
    };

    Object.keys(marcas).forEach(id => {
        const porcentaje = calcularPorcentaje(datos, columnas.marcaPreferida, marcas[id]);
        actualizarTexto(id, porcentaje + "%");
    });
}

/* Actualiza el gráfico de dona (antes tenía 40/30/30 fijo en el CSS) */
function actualizarDona(total, sostenible, noFijo) {
    const dona = document.getElementById("donaTipoRopa");

    if (!dona) return;

    if (!total) {
        dona.style.background = "#eceff1";
        return;
    }

    const corte1 = sostenible;
    const corte2 = redondear(sostenible + noFijo);

    dona.style.background =
        `conic-gradient(#4caf50 0% ${corte1}%, #2196f3 ${corte1}% ${corte2}%, #f44336 ${corte2}% 100%)`;
}

function actualizarBarra(id, porcentaje) {
    const elemento = document.getElementById(id);

    if (elemento) {
        elemento.style.width = porcentaje + "%";
    }
}

function actualizarTexto(id, texto) {
    const elemento = document.getElementById(id);

    if (elemento) {
        elemento.textContent = texto;
    }
}

function normalizar(valor) {
    return String(valor ?? "")
        .trim()
        .toLowerCase()
        .normalize("NFD")
        .replace(/[\u0300-\u036f]/g, "");
}

function redondear(numero) {
    return Number(numero.toFixed(1));
}