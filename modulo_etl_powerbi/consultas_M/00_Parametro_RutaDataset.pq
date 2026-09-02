// =============================================================================
//  PARAMETRO: RutaDataset
// -----------------------------------------------------------------------------
//  Crealo en Power BI: Inicio > Administrar parametros > Nuevo.
//    - Nombre : RutaDataset
//    - Tipo   : Texto
//    - Valor actual: la ruta COMPLETA a tu archivo, por ejemplo:
//        C:\Users\TuUsuario\Documents\modulo_etl_powerbi\Pipeline_ETL_Dataset.xlsx
//
//  Por que un parametro y no la ruta escrita en cada consulta:
//  evita el "hardcoding" — si moves el archivo, cambias UN solo valor y las 4
//  consultas siguen funcionando. Es el error comun #2 del enunciado.
//
//  (Este archivo es solo la referencia del valor; el parametro se crea desde el
//   menu, no pegando este texto en el Editor Avanzado.)
// =============================================================================
"C:\ruta\a\tu\carpeta\Pipeline_ETL_Dataset.xlsx" meta [IsParameterQuery=true, Type="Text", IsParameterQueryRequired=true]
