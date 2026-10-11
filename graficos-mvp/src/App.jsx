import React, { useState, useMemo } from 'react';
import * as XLSX from 'xlsx';
import {
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer,
  PieChart, Pie, Cell
} from 'recharts';
import { 
  Upload, Download, Database, AlertTriangle, CheckCircle, 
  Truck, Package, FileSpreadsheet, ShieldAlert, Filter 
} from 'lucide-react';

// Datos iniciales de simulación (Plantas Agroindustriales)
const DATOS_SIMULACION = [
  { ID_Registro: 'REG-101', Planta: 'Viña del Mar', Insumo: 'Maíz Grano', Stock_Silo_Tn: 1250, En_Transito_GPS_Tn: 350, Capacidad_Max_Tn: 2000, Contratos_Comprometidos_Tn: 1400 },
  { ID_Registro: 'REG-102', Planta: 'Viña del Mar', Insumo: 'Harina de Soya', Stock_Silo_Tn: 320, En_Transito_GPS_Tn: 80, Capacidad_Max_Tn: 1200, Contratos_Comprometidos_Tn: 550 },
  { ID_Registro: 'REG-103', Planta: 'San Antonio', Insumo: 'Trigo Pan', Stock_Silo_Tn: 1800, En_Transito_GPS_Tn: 450, Capacidad_Max_Tn: 2500, Contratos_Comprometidos_Tn: 1900 },
  { ID_Registro: 'REG-104', Planta: 'San Antonio', Insumo: 'Fertilizante NPK', Stock_Silo_Tn: 600, En_Transito_GPS_Tn: 120, Capacidad_Max_Tn: 1000, Contratos_Comprometidos_Tn: 800 },
  { ID_Registro: 'REG-105', Planta: 'Temuco', Insumo: 'Avena Industrial', Stock_Silo_Tn: 950, En_Transito_GPS_Tn: 200, Capacidad_Max_Tn: 1500, Contratos_Comprometidos_Tn: 900 },
  { ID_Registro: 'REG-106', Planta: 'Temuco', Insumo: 'Aceite Vegetal', Stock_Silo_Tn: 410, En_Transito_GPS_Tn: 90, Capacidad_Max_Tn: 800, Contratos_Comprometidos_Tn: 480 },
];

const COLORES_PIE = ['#2563eb', '#16a34a', '#d97706', '#9333ea', '#0891b2', '#dc2626'];

export default function App() {
  const [datos, setDatos] = useState(DATOS_SIMULACION);
  const [origen, setOrigen] = useState('Simulación en Memoria (Datos por Defecto)');
  const [plantaFiltro, setPlantaFiltro] = useState('TODAS');
  const [auditoria, setAuditoria] = useState({
    procesados: 6,
    omitidos: 0,
    erroresDetalle: []
  });

  // Función Validadora y Limpiadora de Datos (Tu argumento clave en la entrevista)
  const validarYLimpiarDatos = (filasCrudas, nombreOrigen) => {
    const idsVistos = new Set();
    const filasValidas = [];
    const errores = [];

    filasCrudas.forEach((fila, index) => {
      const numFila = index + 2; // Considerando el encabezado en fila 1
      const id = String(fila.ID_Registro || '').trim();
      const stock = Number(fila.Stock_Silo_Tn);
      const transito = Number(fila.En_Transito_GPS_Tn);
      const capacidad = Number(fila.Capacidad_Max_Tn);
      const contratos = Number(fila.Contratos_Comprometidos_Tn);

      // 1. Validar campos vacíos o corruptos
      if (!id || !fila.Planta || !fila.Insumo) {
        errores.push(`Fila ${numFila}: Campos obligatorios vacíos (ID, Planta o Insumo).`);
        return;
      }

      // 2. Validar IDs Duplicados (Primary Key)
      if (idsVistos.has(id)) {
        errores.push(`Fila ${numFila}: ID duplicado detectado (${id}). Se omitió para evitar doble conteo.`);
        return;
      }

      // 3. Validar valores numéricos lógicos (No negativos)
      if (isNaN(stock) || isNaN(transito) || stock < 0 || transito < 0) {
        errores.push(`Fila ${numFila} (${id}): Toneladas con valores negativos o formato inválido.`);
        return;
      }

      idsVistos.add(id);

      const stockTotalProyectado = stock + transito;
      const deficitContrato = contratos - stockTotalProyectado;

      filasValidas.push({
        ID_Registro: id,
        Planta: String(fila.Planta).trim(),
        Insumo: String(fila.Insumo).trim(),
        Stock_Silo_Tn: stock,
        En_Transito_GPS_Tn: transito,
        Stock_Proyectado_Tn: stockTotalProyectado,
        Capacidad_Max_Tn: capacidad || 1000,
        Contratos_Comprometidos_Tn: contratos || 0,
        Cumplimiento_OK: stockTotalProyectado >= contratos,
        Deficit_Tn: deficitContrato > 0 ? deficitContrato : 0,
        Ocupacion_Porcentaje: Math.round((stockTotalProyectado / (capacidad || 1)) * 100)
      });
    });

    setDatos(filasValidas);
    setOrigen(nombreOrigen);
    setAuditoria({
      procesados: filasValidas.length,
      omitidos: errores.length,
      erroresDetalle: errores
    });
  };

  // Procesar subida de archivo Excel (.xlsx / .csv)
  const manejarSubidaExcel = (e) => {
    const archivo = e.target.files[0];
    if (!archivo) return;

    const reader = new FileReader();
    reader.onload = (evento) => {
      const bstr = evento.target.result;
      const workbook = XLSX.read(bstr, { type: 'binary' });
      const nombreHoja = workbook.SheetNames[0];
      const hoja = workbook.Sheets[nombreHoja];
      const datosJson = XLSX.utils.sheet_to_json(hoja);

      validarYLimpiarDatos(datosJson, `Archivo Excel: ${archivo.name}`);
    };
    reader.readAsBinaryString(archivo);
    e.target.value = null; // Permitir subir el mismo archivo de nuevo
  };

  // Generar y descargar plantilla Excel de prueba (Incluye errores intencionales para demostrar la auditoría)
  const descargarPlantillaPrueba = () => {
    const datosPlantilla = [
      ...DATOS_SIMULACION,
      // Filas con errores intencionales para lucirse en la demostración en vivo
      { ID_Registro: 'REG-101', Planta: 'Viña del Mar', Insumo: 'Maíz Grano (DUPLICADO)', Stock_Silo_Tn: 1250, En_Transito_GPS_Tn: 350, Capacidad_Max_Tn: 2000, Contratos_Comprometidos_Tn: 1400 },
      { ID_Registro: 'REG-999', Planta: 'San Antonio', Insumo: 'Azúcar Refinada (ERROR NEGATIVO)', Stock_Silo_Tn: -450, En_Transito_GPS_Tn: 100, Capacidad_Max_Tn: 1500, Contratos_Comprometidos_Tn: 500 }
    ];

    const hoja = XLSX.utils.json_to_sheet(datosPlantilla);
    const libro = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(libro, hoja, 'Inventario_GPS');
    XLSX.writeFile(libro, 'Plantilla_Control_Inventario_Soprodi.xlsx');
  };

  // Filtrado por planta y cálculo de KPIs en tiempo real
  const plantasDisponibles = useMemo(() => {
    const unicas = new Set(datos.map(d => d.Planta));
    return ['TODAS', ...Array.from(unicas)];
  }, [datos]);

  const datosFiltrados = useMemo(() => {
    const base = datos.map(fila => {
      const stockTotalProyectado = fila.Stock_Silo_Tn + fila.En_Transito_GPS_Tn;
      return {
        ...fila,
        Stock_Proyectado_Tn: stockTotalProyectado,
        Cumplimiento_OK: stockTotalProyectado >= fila.Contratos_Comprometidos_Tn,
        Etiqueta_Grafico: `${fila.Insumo} (${fila.Planta})`
      };
    });
    return plantaFiltro === 'TODAS' ? base : base.filter(d => d.Planta === plantaFiltro);
  }, [datos, plantaFiltro]);

  const kpis = useMemo(() => {
    return datosFiltrados.reduce((acc, item) => {
      acc.stockSilos += item.Stock_Silo_Tn;
      acc.enTransitoGps += item.En_Transito_GPS_Tn;
      acc.contratosEnRiesgo += !item.Cumplimiento_OK ? 1 : 0;
      return acc;
    }, { stockSilos: 0, enTransitoGps: 0, contratosEnRiesgo: 0 });
  }, [datosFiltrados]);

  return (
    <div className="min-h-screen bg-slate-50 text-slate-800 p-4 md:p-8 font-sans">
      {/* Encabezado Principal */}
      <header className="max-w-7xl mx-auto mb-8 bg-white p-6 rounded-xl shadow-sm border border-slate-200 flex flex-col md:flex-row justify-between items-start md:items-center gap-4">
        <div>
          <div className="flex items-center gap-2 text-blue-700 font-bold text-sm uppercase tracking-wider mb-1">
            <Database className="w-4 h-4" /> Automatización & Control Comercial
          </div>
          <h1 className="text-2xl md:text-3xl font-extrabold text-slate-900">
            Monitor de Existencias, Flota GPS y Contratos
          </h1>
          <p className="text-slate-500 text-sm mt-1">
            Origen activo: <span className="font-semibold text-slate-700">{origen}</span>
          </p>
        </div>

        {/* Botonera de Acciones */}
        <div className="flex flex-wrap items-center gap-3">
          <label className="cursor-pointer bg-blue-600 hover:bg-blue-700 text-white px-4 py-2.5 rounded-lg font-medium text-sm flex items-center gap-2 shadow-sm transition">
            <Upload className="w-4 h-4" />
            Subir Reporte Excel (.xlsx)
            <input type="file" accept=".xlsx, .xls, .csv" onChange={manejarSubidaExcel} className="hidden" />
          </label>

          <button
            onClick={descargarPlantillaPrueba}
            className="bg-emerald-600 hover:bg-emerald-700 text-white px-4 py-2.5 rounded-lg font-medium text-sm flex items-center gap-2 shadow-sm transition"
            title="Descarga un Excel con datos de prueba (incluye errores intencionales para probar el validador)"
          >
            <Download className="w-4 h-4" />
            Descargar Plantilla Demo
          </button>

          <button
            onClick={() => validarYLimpiarDatos(DATOS_SIMULACION, 'Simulación en Memoria (Restaurado)')}
            className="bg-slate-200 hover:bg-slate-300 text-slate-700 px-3 py-2.5 rounded-lg font-medium text-sm transition"
          >
            Restaurar Demo
          </button>
        </div>
      </header>

      <main className="max-w-7xl mx-auto space-y-6">
        {/* Banner de Auditoría Automática de Datos */}
        <div className={`p-4 rounded-xl border flex flex-col md:flex-row justify-between items-start md:items-center gap-4 ${
          auditoria.omitidos > 0 ? 'bg-amber-50 border-amber-300 text-amber-900' : 'bg-emerald-50 border-emerald-200 text-emerald-900'
        }`}>
          <div className="flex items-start gap-3">
            {auditoria.omitidos > 0 ? (
              <ShieldAlert className="w-6 h-6 text-amber-600 shrink-0 mt-0.5" />
            ) : (
              <CheckCircle className="w-6 h-6 text-emerald-600 shrink-0 mt-0.5" />
            )}
            <div>
              <h3 className="font-bold text-sm md:text-base">
                Motor de Validación de Integridad: {auditoria.procesados} registros válidos cargados
                {auditoria.omitidos > 0 && ` (${auditoria.omitidos} filas corruptas o duplicadas fueron bloqueadas)`}
              </h3>
              {auditoria.omitidos > 0 ? (
                <ul className="text-xs mt-1 space-y-0.5 list-disc list-inside text-amber-800">
                  {auditoria.erroresDetalle.map((err, i) => <li key={i}>{err}</li>)}
                </ul>
              ) : (
                <p className="text-xs text-emerald-700 mt-0.5">
                  Verificación de claves únicas (ID) y consistencia matemática de tonelaje completada sin anomalías.
                </p>
              )}
            </div>
          </div>

          {/* Filtro por Planta */}
          <div className="flex items-center gap-2 self-end md:self-center bg-white px-3 py-1.5 rounded-lg border border-slate-200">
            <Filter className="w-4 h-4 text-slate-500" />
            <span className="text-xs font-semibold text-slate-600">Planta:</span>
            <select
              value={plantaFiltro}
              onChange={(e) => setPlantaFiltro(e.target.value)}
              className="text-sm font-bold text-slate-800 bg-transparent focus:outline-none cursor-pointer"
            >
              {plantasDisponibles.map(planta => (
                <option key={planta} value={planta}>{planta}</option>
              ))}
            </select>
          </div>
        </div>

        {/* Tarjetas de KPIs Superiores */}
        <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
          <div className="bg-white p-6 rounded-xl shadow-sm border border-slate-200 flex items-center justify-between">
            <div>
              <p className="text-sm font-medium text-slate-500">Existencias Físicas en Silos</p>
              <p className="text-3xl font-extrabold text-slate-900 mt-1">
                {kpis.stockSilos.toLocaleString('es-CL')} <span className="text-base font-normal text-slate-500">Tn</span>
              </p>
              <p className="text-xs text-blue-600 mt-2 font-medium">Disponibilidad inmediata en planta</p>
            </div>
            <div className="p-4 bg-blue-50 rounded-full text-blue-600">
              <Package className="w-8 h-8" />
            </div>
          </div>

          <div className="bg-white p-6 rounded-xl shadow-sm border border-slate-200 flex items-center justify-between">
            <div>
              <p className="text-sm font-medium text-slate-500">Abastecimiento en Ruta (GPS)</p>
              <p className="text-3xl font-extrabold text-emerald-600 mt-1">
                +{kpis.enTransitoGps.toLocaleString('es-CL')} <span className="text-base font-normal text-slate-500">Tn</span>
              </p>
              <p className="text-xs text-emerald-700 mt-2 font-medium">Total Proyectado: {(kpis.stockSilos + kpis.enTransitoGps).toLocaleString('es-CL')} Tn</p>
            </div>
            <div className="p-4 bg-emerald-50 rounded-full text-emerald-600">
              <Truck className="w-8 h-8" />
            </div>
          </div>

          <div className="bg-white p-6 rounded-xl shadow-sm border border-slate-200 flex items-center justify-between">
            <div>
              <p className="text-sm font-medium text-slate-500">Alertas de Contratos Comerciales</p>
              <p className={`text-3xl font-extrabold mt-1 ${kpis.contratosEnRiesgo > 0 ? 'text-red-600' : 'text-slate-900'}`}>
                {kpis.contratosEnRiesgo} <span className="text-base font-normal text-slate-500">insumos en déficit</span>
              </p>
              <p className="text-xs text-slate-500 mt-2">Cruce Stock + GPS vs Compromisos</p>
            </div>
            <div className={`p-4 rounded-full ${kpis.contratosEnRiesgo > 0 ? 'bg-red-50 text-red-600' : 'bg-slate-100 text-slate-500'}`}>
              <AlertTriangle className="w-8 h-8" />
            </div>
          </div>
        </div>

        {/* Sección de Gráficos */}
        <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
          {/* Gráfico 1: Barras Apiladas (Stock Silo + Camiones GPS vs Contratos) */}
          <div className="bg-white p-6 rounded-xl shadow-sm border border-slate-200 lg:col-span-2">
            <h2 className="text-lg font-bold text-slate-800 mb-1">
              Proyección de Stock (Silos + GPS) vs Contratos Comprometidos
            </h2>
            <p className="text-xs text-slate-500 mb-6">
              Permite detectar quiebres de stock antes de que ocurran cruzando existencias físicas y camiones en tránsito.
            </p>
            <div className="h-80 w-full">
              <ResponsiveContainer width="100%" height="100%">
                <BarChart data={datosFiltrados} margin={{ top: 10, right: 20, left: 0, bottom: 20 }}>
                  <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#e2e8f0" />
                  <XAxis dataKey="Insumo" tick={{ fontSize: 12 }} />
                  <YAxis tick={{ fontSize: 12 }} unit=" Tn" />
                  <Tooltip 
                    formatter={(value) => [`${value.toLocaleString('es-CL')} Tn`, '']}
                    contentStyle={{ borderRadius: '8px', border: '1px solid #cbd5e1' }}
                  />
                  <Legend wrapperStyle={{ paddingTop: '10px' }} />
                  <Bar dataKey="Stock_Silo_Tn" name="Stock en Silo" stackId="a" fill="#2563eb" radius={[0, 0, 4, 4]} />
                  <Bar dataKey="En_Transito_GPS_Tn" name="En Ruta (GPS)" stackId="a" fill="#10b981" radius={[4, 4, 0, 0]} />
                  <Bar dataKey="Contratos_Comprometidos_Tn" name="Contrato Comprometido" fill="#f59e0b" radius={[4, 4, 0, 0]} />
                </BarChart>
              </ResponsiveContainer>
            </div>
          </div>

          {/* Gráfico 2: Distribución de Materias Primas */}
          <div className="bg-white p-6 rounded-xl shadow-sm border border-slate-200">
            <h2 className="text-lg font-bold text-slate-800 mb-1">
              Distribución por Materia Prima
            </h2>
            <p className="text-xs text-slate-500 mb-4">
              Volumen total proyectado (Silo + Ruta) en toneladas.
            </p>
            <div className="h-80 w-full">
              <ResponsiveContainer width="100%" height="100%">
                <PieChart>
                  <Pie
                    data={datosFiltrados}
                    dataKey="Stock_Proyectado_Tn"
                    nameKey="Insumo"
                    cx="50%"
                    cy="45%"
                    innerRadius={55}
                    outerRadius={90}
                    paddingAngle={3}
                    label={({ percent }) => `${(percent * 100).toFixed(0)}%`}
                  >
                    {datosFiltrados.map((entry, index) => (
                      <Cell key={`cell-${index}`} fill={COLORES_PIE[index % COLORES_PIE.length]} />
                    ))}
                  </Pie>
                  <Tooltip formatter={(value) => `${value.toLocaleString('es-CL')} Tn`} />
                  <Legend layout="horizontal" verticalAlign="bottom" align="center" wrapperStyle={{ fontSize: '12px' }} />
                </PieChart>
              </ResponsiveContainer>
            </div>
          </div>
        </div>

        {/* Tabla de Control de Contratos y Ocupación de Planta */}
        <div className="bg-white rounded-xl shadow-sm border border-slate-200 overflow-hidden">
          <div className="p-5 border-b border-slate-200 flex items-center justify-between">
            <div className="flex items-center gap-2">
              <FileSpreadsheet className="w-5 h-5 text-blue-600" />
              <h2 className="text-lg font-bold text-slate-800">Detalle de Cumplimiento de Contratos por Planta</h2>
            </div>
            <span className="text-xs text-slate-500">Actualizado al instante en memoria local</span>
          </div>
          <div className="overflow-x-auto">
            <table className="w-full text-left border-collapse">
              <thead>
                <tr className="bg-slate-50 text-slate-600 text-xs uppercase border-b border-slate-200">
                  <th className="py-3 px-4">ID</th>
                  <th className="py-3 px-4">Planta</th>
                  <th className="py-3 px-4">Insumo</th>
                  <th className="py-3 px-4 text-right">Silo (Tn)</th>
                  <th className="py-3 px-4 text-right">Ruta GPS (Tn)</th>
                  <th className="py-3 px-4 text-right">Total Proyectado</th>
                  <th className="py-3 px-4 text-right">Contrato (Tn)</th>
                  <th className="py-3 px-4 text-center">Estado Contrato</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-200 text-sm">
                {datosFiltrados.map((item) => (
                  <tr key={item.ID_Registro} className="hover:bg-slate-50">
                    <td className="py-3 px-4 font-mono text-xs text-slate-500">{item.ID_Registro}</td>
                    <td className="py-3 px-4 font-medium">{item.Planta}</td>
                    <td className="py-3 px-4 font-semibold text-slate-900">{item.Insumo}</td>
                    <td className="py-3 px-4 text-right">{item.Stock_Silo_Tn.toLocaleString('es-CL')}</td>
                    <td className="py-3 px-4 text-right text-emerald-600 font-medium">+{item.En_Transito_GPS_Tn.toLocaleString('es-CL')}</td>
                    <td className="py-3 px-4 text-right font-bold">{item.Stock_Proyectado_Tn.toLocaleString('es-CL')}</td>
                    <td className="py-3 px-4 text-right">{item.Contratos_Comprometidos_Tn.toLocaleString('es-CL')}</td>
                    <td className="py-3 px-4 text-center">
                      {item.Cumplimiento_OK ? (
                        <span className="inline-flex items-center gap-1 px-2.5 py-1 rounded-full text-xs font-semibold bg-emerald-100 text-emerald-800">
                          Cubierto
                        </span>
                      ) : (
                        <span className="inline-flex items-center gap-1 px-2.5 py-1 rounded-full text-xs font-semibold bg-red-100 text-red-800">
                          Déficit (-{item.Deficit_Tn} Tn)
                        </span>
                      )}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      </main>
    </div>
  );
}