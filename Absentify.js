const XLSX = require('xlsx');

// Función para procesar el archivo Excel
function procesarExcel(inputFilePath, outputFilePath) {
    // Leer el archivo Excel
    const workbook = XLSX.readFile(inputFilePath);
    const sheetName = workbook.SheetNames[0]; // Asumimos que el archivo tiene una sola hoja
    const worksheet = workbook.Sheets[sheetName];

    // Convertir la hoja de cálculo a JSON
    const data = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

    // Eliminar las filas no deseadas (primera, segunda y cuarta)
    data.splice(0, 1); // Eliminar primera fila
    data.splice(0, 1); // Eliminar segunda fila (ahora es la primera después de eliminar la anterior)
    data.splice(1, 1); // Eliminar cuarta fila (ahora es la segunda después de eliminar las anteriores)

    // Obtener la fecha de creación (tercera fila)
    const fechaCreacion = data[0][1];

    // Obtener los nombres de las asignaturas (quinta fila)
    const asignaturas = data[1].slice(1); // Ignorar la primera columna (nombre del alumno)

    // Crear un nuevo array para los datos procesados
    const nuevosDatos = [];

    // Añadir la fila de la fecha de creación
    nuevosDatos.push(['Fecha de creación:', fechaCreacion]);

    // Crear la cabecera de las nuevas columnas
    const cabeceraSuperior = ['Alumno/a']; // Cabecera superior (celdas combinadas)
    const cabeceraInferior = ['']; // Cabecera inferior (J, I, TOTAL)

    // Añadir las columnas de "Día completo"
    //cabeceraSuperior.push('Día completo');
    cabeceraInferior.push('J', 'I', 'TOTAL');

    // Añadir las columnas de las asignaturas
    asignaturas.forEach(asignatura => {
        cabeceraSuperior.push("", "", asignatura); // Celda combinada con el nombre de la asignatura
        cabeceraInferior.push('J', 'I', 'TOTAL'); // Subcabecera para Justificadas, Injustificadas y Total
    });

    // Añadir las cabeceras al array de nuevos datos
    nuevosDatos.push(cabeceraSuperior, cabeceraInferior);

    // Procesar cada fila de datos
    for (let i = 2; i < data.length; i++) {
        const fila = data[i];
        const nuevaFila = [fila[0]]; // Nombre del alumno

        // Procesar cada columna
        for (let j = 1; j < fila.length; j++) {
            if (fila[j]) {
                const [justificadas, injustificadas] = fila[j].split('-').slice(0, 2);
                const total = parseInt(justificadas.replace('J', '')) + parseInt(injustificadas.replace('I', ''));

                nuevaFila.push(
                    justificadas.replace('J', '').replace('I', ''),
                    injustificadas.replace('J', '').replace('I', ''),
                    total
                );
            } else {
                nuevaFila.push('', '', ''); // Si no hay datos, añadir celdas vacías
            }
        }
        nuevosDatos.push(nuevaFila);
    }

    // Crear un nuevo libro de Excel
    const nuevoWorkbook = XLSX.utils.book_new();
    const nuevoWorksheet = XLSX.utils.aoa_to_sheet(nuevosDatos);

    // Combinar celdas para la cabecera superior
    let colIndex = 1; // Empezar desde la segunda columna (la primera es "Alumno/a")
    cabeceraSuperior.forEach((header, index) => {
        if (index > 0) { // Ignorar la primera columna ("Alumno/a")
            const numSubColumns = 1; // Cada grupo tiene 3 columnas (J, I, TOTAL)
            nuevoWorksheet['!merges'] = nuevoWorksheet['!merges'] || [];
            nuevoWorksheet['!merges'].push({
                s: { r: 1, c: colIndex }, // Fila 1 (cabecera superior), columna actual
                e: { r: 1, c: colIndex + numSubColumns - 1 } // Fila 1, columna final del grupo
            });
            colIndex += numSubColumns; // Mover al siguiente grupo de columnas
        }
    });

    XLSX.utils.book_append_sheet(nuevoWorkbook, nuevoWorksheet, 'Datos Procesados');

    // Guardar el nuevo archivo
    XLSX.writeFile(nuevoWorkbook, outputFilePath);

    console.log(`Archivo procesado guardado como: ${outputFilePath}`);
}

// Obtener los nombres de los archivos de entrada y salida desde los parámetros
const inputFilePath = process.argv[2];
const outputFilePath = process.argv[3];

// Verificar que se hayan proporcionado los parámetros
if (!inputFilePath || !outputFilePath) {
    console.error('Uso: node procesarExcel.js <archivo_entrada> <archivo_salida>');
    process.exit(1);
}

// Ejecutar la función con los archivos proporcionados
procesarExcel(inputFilePath, outputFilePath);