// ============================================================================
// Parametro: RutaDatos
// ----------------------------------------------------------------------------
// Resuelve el problema de la ruta de origen "hardcodeada": en lugar de repetir
// C:\...\datos dentro de cada consulta, el path vive en un unico lugar y las
// cuatro consultas lo concatenan con el nombre de archivo.
//
// POR QUE UN PARAMETRO Y NO UNA CONSULTA QUE LEA a UNA CELDA
// ----------------------------------------------------------------------------
// La alternativa elegante seria leer la ruta desde una celda del libro con
// Excel.CurrentWorkbook(). No se usa a proposito: Ventas quedaria referenciando
// otra consulta Y accediendo a un origen externo con File.Contents en el mismo
// paso, que es la condicion exacta que dispara el error
//
//     Formula.Firewall: la consulta 'Ventas' hace referencia a otras consultas
//     o pasos, por lo que puede que no tenga acceso directo a un origen de datos
//
// Un parametro, en cambio, el motor lo resuelve como valor literal antes de
// evaluar la consulta, asi que no cuenta como referencia y el firewall no salta.
// Es la solucion idiomatica al problema, no un rodeo.
//
// COMO SE CREA
// ----------------------------------------------------------------------------
// Opcion A (recomendada): Datos > Obtener datos > Iniciar el editor de Power
//   Query > Inicio > Administrar parametros > Nuevo. Nombre: RutaDatos,
//   Tipo: Texto, Valor actual: la ruta de la carpeta 'datos' terminada en \
//
// Opcion B: crear una consulta en blanco, pegar la linea de abajo en el Editor
//   avanzado y renombrarla RutaDatos. El metadato IsParameterQuery hace que
//   Excel la trate como parametro, no como consulta.
// ============================================================================

"C:\Users\TU_USUARIO\Documents\ModeloBI\datos\"
    meta
    [
        IsParameterQuery = true,
        List = null,
        DefaultValue = null,
        Type = "Text",
        IsParameterQueryRequired = true
    ]
