import React, { useState, useMemo } from 'react';
import * as XLSX from 'xlsx';
import {
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer,
  PieChart, Pie, Cell
} from 'recharts';
import { 
  Upload, Download, Database, AlertTriangle, 
  Truck, Package, FileSpreadsheet, ShieldAlert,
  Activity, Grid
} from 'lucide-react';

const DATOS_SIMULACION = [
  { ID_Registro: 'REG-101', Planta: 'Viña del Mar', Insumo: 'Maíz Grano', Stock_Silo_Tn: 1250, En_Transito_GPS_Tn: 350, Capacidad_Max_Tn: 2000, Contratos_Comprometidos_Tn: 1400 },
  { ID_Registro: 'REG-102', Planta: 'Viña del Mar', Insumo: 'Harina de Soya', Stock_Silo_Tn: 320, En_Transito_GPS_Tn: 80, Capacidad_Max_Tn: 1200, Contratos_Comprometidos_Tn: 550 },
  { ID_Registro: 'REG-103', Planta: 'San Antonio', Insumo: 'Trigo Pan', Stock_Silo_Tn: 1800, En_Transito_GPS_Tn: 450, Capacidad_Max_Tn: 2500, Contratos_Comprometidos_Tn: 1900 },
  { ID_Registro: 'REG-104', Planta: 'San Antonio', Insumo: 'Fertilizante NPK', Stock_Silo_Tn: 600, En_Transito_GPS_Tn: 120, Capacidad_Max_Tn: 1000, Contratos_Comprometidos_Tn: 800 },
  { ID_Registro: 'REG-105', Planta: 'Temuco', Insumo: 'Avena Industrial', Stock_Silo_Tn: 950, En_Transito_GPS_Tn: 200, Capacidad_Max_Tn: 1500, Contratos_Comprometidos_Tn: 900 },
  { ID_Registro: 'REG-106', Planta: 'Temuco', Insumo: 'Aceite Vegetal', Stock_Silo_Tn: 410, En_Transito_GPS_Tn: 90, Capacidad_Max_Tn: 800, Contratos_Comprometidos_Tn: 480 },
];

const COLORES_GLOW = ['#3b82f6', '#10b981', '#f59e0b', '#8b5cf6', '#06b6d4', '#ef4444'];

export default function App() {
  const [datos, setDatos] = useState(DATOS_SIMULACION);
  const [origen, setOrigen] = useState('Por Defecto');
  const [plantaFiltro, setPlantaFiltro] = useState('TODAS');
  const [auditoria, setAuditoria] = useState({
    procesados: 6,
    omitidos: 0,
    erroresDetalle: []
  });

  const validarYLimpiarDatos = (filasCrudas, nombreOrigen) => {
    const idsVistos = new Set();
    const filasValidas = [];
    const errores = [];

    filasCrudas.forEach((fila, index) => {
      const numFila = index + 2;
      const id = String(fila.ID_Registro || '').trim();
      const stock = Number(fila.Stock_Silo_Tn);
      const transito = Number(fila.En_Transito_GPS_Tn);
      const capacidad = Number(fila.Capacidad_Max_Tn);
      const contratos = Number(fila.Contratos_Comprometidos_Tn);

      if (!id || !fila.Planta || !fila.Insumo) {
        errores.push(`Fila ${numFila}: Campos vacíos (ID, Planta o Insumo).`);
        return;
      }
      if (idsVistos.has(id)) {
        errores.push(`Fila ${numFila}: ID duplicado detectado (${id}).`);
        return;
      }
      if (isNaN(stock) || isNaN(transito) || stock < 0 || transito < 0) {
        errores.push(`Fila ${numFila} (${id}): Toneladas inválidas o negativas.`);
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

      validarYLimpiarDatos(datosJson, archivo.name);
    };
    reader.readAsBinaryString(archivo);
    e.target.value = null;
  };

  const descargarPlantillaPrueba = () => {
    const datosPlantilla = [
      ...DATOS_SIMULACION,
      { ID_Registro: 'REG-101', Planta: 'Viña del Mar', Insumo: 'Maíz Grano (DUPLICADO)', Stock_Silo_Tn: 1250, En_Transito_GPS_Tn: 350, Capacidad_Max_Tn: 2000, Contratos_Comprometidos_Tn: 1400 },
      { ID_Registro: 'REG-999', Planta: 'San Antonio', Insumo: 'Azúcar Refinada (ERROR NEGATIVO)', Stock_Silo_Tn: -450, En_Transito_GPS_Tn: 100, Capacidad_Max_Tn: 1500, Contratos_Comprometidos_Tn: 500 }
    ];

    const hoja = XLSX.utils.json_to_sheet(datosPlantilla);
    const libro = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(libro, hoja, 'Inventario_GPS');
    XLSX.writeFile(libro, 'Soprodi_Inventario_Auditado.xlsx');
  };

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
    <>
      {/* Reset CSS forzado para eliminar el centrado y márgenes por defecto de Vite */}
      <style>{`
        html, body, #root {
          margin: 0 !important;
          padding: 0 !important;
          width: 100% !important;
          max-width: 100% !important;
          height: 100vh !important;
          overflow-x: hidden !important;
          display: block !important;
          text-align: left !important;
          background-color: #020617 !important;
        }
      `}</style>

      <div className="flex h-screen w-full bg-slate-950 text-slate-100 font-sans overflow-hidden">
        {/* Barra Lateral Dashboard */}
        <aside className="w-64 bg-slate-900 border-r border-slate-800 flex flex-col justify-between shrink-0 hidden lg:flex h-full">
          <div>
            {/* Logo Soprodi */}
            <div className="p-5 border-b border-slate-800/60 flex items-center gap-3">
              <div className="p-2.5 bg-gradient-to-tr from-blue-600 to-indigo-500 rounded-lg shadow-lg shadow-blue-500/20 text-white font-bold tracking-tight">
                SP
              </div>
              <div>
                <h2 className="font-extrabold text-white tracking-wider text-lg leading-none">SOPRODI</h2>
                <span className="text-[10px] text-blue-400 font-semibold tracking-wide uppercase">Operaciones Logísticas</span>
              </div>
            </div>

            <div className="p-4">
              <span className="text-[10px] font-bold text-slate-500 tracking-wider uppercase block mb-2 px-2">Principal</span>
              <nav className="space-y-1">
                <button className="w-full flex items-center gap-3 px-3 py-2.5 rounded-xl bg-gradient-to-r from-blue-600/15 to-blue-600/5 text-blue-400 text-sm font-semibold border border-blue-500/20">
                  <Grid className="w-4 h-4 text-blue-400" />
                  Panel de Control
                </button>
                <button className="w-full flex items-center gap-3 px-3 py-2.5 rounded-xl text-slate-400 hover:text-white hover:bg-slate-800/40 text-sm transition">
                  <Activity className="w-4 h-4 text-slate-400" />
                  Auditoría en Tiempo Real
                </button>
              </nav>
            </div>
          </div>

          {/* Auditoría en Lateral */}
          <div className="p-4 border-t border-slate-800 bg-slate-950/50 m-3 rounded-2xl">
            <div className="flex items-center gap-2 mb-3">
              <ShieldAlert className="w-4 h-4 text-emerald-400" />
              <h3 className="font-bold text-xs text-slate-300 uppercase">Estado Auditoría</h3>
            </div>
            <div className="grid grid-cols-2 gap-2 text-center">
              <div className="bg-slate-900 border border-slate-800 rounded-xl p-2">
                <span className="text-[10px] font-medium text-slate-500 uppercase block">Válidos</span>
                <span className="text-lg font-bold text-emerald-400">{auditoria.procesados}</span>
              </div>
              <div className="bg-slate-900 border border-slate-800 rounded-xl p-2">
                <span className="text-[10px] font-medium text-slate-500 uppercase block">Errores</span>
                <span className={`text-lg font-bold ${auditoria.omitidos > 0 ? 'text-amber-400' : 'text-slate-500'}`}>
                  {auditoria.omitidos}
                </span>
              </div>
            </div>
            {auditoria.omitidos > 0 && (
              <div className="mt-2.5 text-[11px] text-amber-400/90 font-medium bg-amber-500/5 p-2 rounded-lg border border-amber-500/10">
                Se omitieron {auditoria.omitidos} registros con ID duplicado o datos corruptos.
              </div>
            )}
          </div>
        </aside>

        {/* Contenido Principal */}
        <div className="flex-1 flex flex-col min-w-0 h-full overflow-y-auto overflow-x-hidden bg-slate-950">
          <header className="bg-slate-900/60 backdrop-blur-md border-b border-slate-800 px-6 py-4 sticky top-0 z-30 flex flex-wrap justify-between items-center gap-4">
            <div className="min-w-0">
              <div className="flex items-center gap-2 mb-1">
                <span className="inline-flex items-center gap-1.5 px-2.5 py-0.5 rounded-full text-xs font-semibold bg-blue-500/10 text-blue-400 border border-blue-500/20">
                  <Database className="w-3 h-3" /> Automático
                </span>
                <span className="text-xs text-slate-500">•</span>
                <span className="text-xs text-slate-400 font-medium">Viña del Mar / San Antonio / Temuco</span>
              </div>
              <h1 className="text-xl md:text-2xl font-black text-white tracking-tight">
                Monitor de Existencias, Flotas y Contratos
              </h1>
            </div>

            <div className="flex flex-wrap items-center gap-2.5">
              <label className="cursor-pointer bg-blue-600 hover:bg-blue-500 text-white px-4 py-2 rounded-xl font-bold text-xs flex items-center gap-2 shadow-lg shadow-blue-600/20 transition border border-blue-500/30">
                <Upload className="w-3.5 h-3.5" />
                Subir Excel (.xlsx)
                <input type="file" accept=".xlsx, .xls, .csv" onChange={manejarSubidaExcel} className="hidden" />
              </label>

              <button
                onClick={descargarPlantillaPrueba}
                className="bg-emerald-600/10 hover:bg-emerald-600/20 text-emerald-400 px-4 py-2 rounded-xl font-bold text-xs flex items-center gap-2 border border-emerald-500/20 transition"
              >
                <Download className="w-3.5 h-3.5" />
                Plantilla Demo
              </button>

              <button
                onClick={() => validarYLimpiarDatos(DATOS_SIMULACION, 'Por Defecto')}
                className="bg-slate-800 hover:bg-slate-700 text-slate-300 px-3 py-2 rounded-xl font-bold text-xs transition"
              >
                Restablecer
              </button>
            </div>
          </header>

          {/* Cuerpo del Dashboard */}
          <main className="p-6 space-y-6 flex-1 min-w-0">
            {auditoria.omitidos > 0 && (
              <div className="bg-amber-500/10 border border-amber-500/20 p-4 rounded-2xl flex items-start gap-3">
                <ShieldAlert className="w-5 h-5 text-amber-500 mt-0.5 shrink-0" />
                <div>
                  <h4 className="font-bold text-sm text-amber-400">Auditoría: Registro corrupto o duplicado omitido</h4>
                  <p className="text-xs text-slate-400 mt-0.5">
                    El sistema ha filtrado {auditoria.omitidos} registros defectuosos ({auditoria.erroresDetalle.join(' ')}) para asegurar el 100% de consistencia.
                  </p>
                </div>
              </div>
            )}

            {/* Tarjetas KPI */}
            <div className="grid grid-cols-1 md:grid-cols-3 gap-5">
              <div className="bg-slate-900 border border-slate-800 p-5 rounded-2xl flex items-center justify-between">
                <div>
                  <p className="text-xs font-semibold text-slate-400 uppercase tracking-wide">Stock Físico en Silos</p>
                  <p className="text-3xl font-black text-white mt-1 font-mono">
                    {kpis.stockSilos.toLocaleString('es-CL')} <span className="text-sm font-normal text-slate-500">Tn</span>
                  </p>
                  <div className="flex items-center gap-1.5 text-[11px] text-blue-400 mt-2 font-semibold uppercase">
                    <span className="w-2 h-2 rounded-full bg-blue-400 animate-pulse" />
                    Disponibilidad de Planta
                  </div>
                </div>
                <div className="p-3.5 bg-blue-500/10 border border-blue-500/20 rounded-xl text-blue-400 shrink-0">
                  <Package className="w-6 h-6" />
                </div>
              </div>

              <div className="bg-slate-900 border border-slate-800 p-5 rounded-2xl flex items-center justify-between">
                <div>
                  <p className="text-xs font-semibold text-slate-400 uppercase tracking-wide">Tránsito en Ruta (GPS)</p>
                  <p className="text-3xl font-black text-emerald-400 mt-1 font-mono">
                    +{kpis.enTransitoGps.toLocaleString('es-CL')} <span className="text-sm font-normal text-slate-500">Tn</span>
                  </p>
                  <p className="text-[11px] text-slate-400 mt-2 font-semibold">
                    Silo + GPS: <span className="text-emerald-400 font-bold">{(kpis.stockSilos + kpis.enTransitoGps).toLocaleString('es-CL')} Tn</span>
                  </p>
                </div>
                <div className="p-3.5 bg-emerald-500/10 border border-emerald-500/20 rounded-xl text-emerald-400 shrink-0">
                  <Truck className="w-6 h-6" />
                </div>
              </div>

              <div className="bg-slate-900 border border-slate-800 p-5 rounded-2xl flex items-center justify-between">
                <div>
                  <p className="text-xs font-semibold text-slate-400 uppercase tracking-wide">Contratos Comprometidos</p>
                  <p className={`text-3xl font-black mt-1 font-mono ${kpis.contratosEnRiesgo > 0 ? 'text-amber-400' : 'text-slate-300'}`}>
                    {kpis.contratosEnRiesgo} <span className="text-sm font-normal text-slate-500">en riesgo</span>
                  </p>
                  <p className="text-[11px] text-slate-400 mt-2 font-semibold uppercase">
                    Auditoría automática de déficit
                  </p>
                </div>
                <div className={`p-3.5 rounded-xl border shrink-0 ${kpis.contratosEnRiesgo > 0 ? 'bg-amber-500/10 border-amber-500/20 text-amber-400' : 'bg-slate-800/40 border-slate-700/50 text-slate-500'}`}>
                  <AlertTriangle className="w-6 h-6" />
                </div>
              </div>
            </div>

            {/* Gráficos */}
            <div className="grid grid-cols-1 lg:grid-cols-3 gap-5">
              <div className="bg-slate-900 border border-slate-800/80 p-5 rounded-2xl lg:col-span-2 min-w-0">
                <div className="flex flex-wrap justify-between items-center gap-4 mb-4">
                  <div>
                    <h2 className="text-base font-extrabold text-white">Existencias vs Requisitos de Contratos</h2>
                    <p className="text-xs text-slate-400 mt-0.5">Distribución y proyección de stock en base a contratos existentes.</p>
                  </div>

                  <div className="flex items-center gap-2 bg-slate-950 p-1.5 rounded-xl border border-slate-800">
                    <span className="text-[10px] font-bold text-slate-400 px-2 uppercase">Planta:</span>
                    <select
                      value={plantaFiltro}
                      onChange={(e) => setPlantaFiltro(e.target.value)}
                      className="text-xs font-bold text-blue-400 bg-transparent focus:outline-none cursor-pointer pr-2"
                    >
                      {plantasDisponibles.map(planta => (
                        <option key={planta} value={planta} className="bg-slate-900 text-slate-200">{planta}</option>
                      ))}
                    </select>
                  </div>
                </div>

                <div className="h-72 w-full">
                  <ResponsiveContainer width="100%" height="100%">
                    <BarChart data={datosFiltrados} margin={{ top: 10, right: 10, left: -15, bottom: 0 }}>
                      <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#1e293b" />
                      <XAxis dataKey="Insumo" stroke="#64748b" fontSize={11} tickLine={false} />
                      <YAxis stroke="#64748b" fontSize={11} tickLine={false} />
                      <Tooltip 
                        contentStyle={{ backgroundColor: '#0f172a', borderColor: '#334155', borderRadius: '12px' }}
                        itemStyle={{ color: '#f1f5f9' }}
                      />
                      <Legend wrapperStyle={{ fontSize: '11px', paddingTop: '10px' }} />
                      <Bar dataKey="Stock_Silo_Tn" name="Stock en Silo (Tn)" stackId="a" fill="#3b82f6" radius={[0, 0, 4, 4]} />
                      <Bar dataKey="En_Transito_GPS_Tn" name="Tránsito GPS (Tn)" stackId="a" fill="#10b981" radius={[4, 4, 0, 0]} />
                      <Bar dataKey="Contratos_Comprometidos_Tn" name="Contrato Comprometido (Tn)" fill="#f59e0b" radius={[4, 4, 4, 4]} />
                    </BarChart>
                  </ResponsiveContainer>
                </div>
              </div>

              <div className="bg-slate-900 border border-slate-800/80 p-5 rounded-2xl min-w-0 flex flex-col justify-between">
                <div>
                  <h2 className="text-base font-extrabold text-white mb-1">Volumen Total</h2>
                  <p className="text-xs text-slate-400 mb-2">Proporción general de inventario por insumo.</p>
                </div>
                <div className="h-56 w-full">
                  <ResponsiveContainer width="100%" height="100%">
                    <PieChart>
                      <Pie
                        data={datosFiltrados}
                        dataKey="Stock_Proyectado_Tn"
                        nameKey="Insumo"
                        cx="50%"
                        cy="50%"
                        innerRadius={50}
                        outerRadius={75}
                        paddingAngle={3}
                        label={({ percent }) => `${(percent * 100).toFixed(0)}%`}
                      >
                        {datosFiltrados.map((entry, index) => (
                          <Cell key={`cell-${index}`} fill={COLORES_GLOW[index % COLORES_GLOW.length]} />
                        ))}
                      </Pie>
                      <Tooltip 
                        contentStyle={{ backgroundColor: '#0f172a', borderColor: '#334155', borderRadius: '12px' }}
                      />
                    </PieChart>
                  </ResponsiveContainer>
                </div>
                <div className="grid grid-cols-2 gap-2 mt-2">
                  {datosFiltrados.map((item, index) => (
                    <div key={item.ID_Registro} className="flex items-center gap-2">
                      <span className="w-2.5 h-2.5 rounded-full shrink-0" style={{ backgroundColor: COLORES_GLOW[index % COLORES_GLOW.length] }} />
                      <span className="text-[11px] text-slate-300 truncate">{item.Insumo}</span>
                    </div>
                  ))}
                </div>
              </div>
            </div>

            {/* Tabla de Control de Contratos */}
            <div className="bg-slate-900 border border-slate-800 rounded-2xl overflow-hidden shadow-xl">
              <div className="p-5 border-b border-slate-800 flex items-center gap-3">
                <FileSpreadsheet className="w-5 h-5 text-blue-400" />
                <div>
                  <h2 className="text-base font-extrabold text-white">Detalle de Existencias y Contratos por Planta</h2>
                  <p className="text-xs text-slate-400 mt-0.5">Base de datos de existencias físicas cruzada en tiempo real con camiones GPS.</p>
                </div>
              </div>
              <div className="overflow-x-auto">
                <table className="w-full text-left border-collapse">
                  <thead>
                    <tr className="bg-slate-950/40 text-slate-400 text-xs font-bold uppercase border-b border-slate-800/80">
                      <th className="py-3.5 px-5">ID</th>
                      <th className="py-3.5 px-5">Planta</th>
                      <th className="py-3.5 px-5">Materia Prima</th>
                      <th className="py-3.5 px-5 text-right">Existencia Física</th>
                      <th className="py-3.5 px-5 text-right">Abastecimiento en Ruta</th>
                      <th className="py-3.5 px-5 text-right">Proyectado</th>
                      <th className="py-3.5 px-5 text-right">Contratos</th>
                      <th className="py-3.5 px-5 text-center">Auditoría Contrato</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-800 text-sm">
                    {datosFiltrados.map((item) => (
                      <tr key={item.ID_Registro} className="hover:bg-slate-800/20 transition">
                        <td className="py-3.5 px-5 font-mono text-xs text-slate-500">{item.ID_Registro}</td>
                        <td className="py-3.5 px-5 font-medium text-slate-300">{item.Planta}</td>
                        <td className="py-3.5 px-5 font-bold text-white">{item.Insumo}</td>
                        <td className="py-3.5 px-5 text-right font-mono text-slate-300">{item.Stock_Silo_Tn.toLocaleString('es-CL')} Tn</td>
                        <td className="py-3.5 px-5 text-right font-mono text-emerald-400 font-medium">+{item.En_Transito_GPS_Tn.toLocaleString('es-CL')} Tn</td>
                        <td className="py-3.5 px-5 text-right font-mono font-bold text-white">{item.Stock_Proyectado_Tn.toLocaleString('es-CL')} Tn</td>
                        <td className="py-3.5 px-5 text-right font-mono text-slate-400">{item.Contratos_Comprometidos_Tn.toLocaleString('es-CL')} Tn</td>
                        <td className="py-3.5 px-5 text-center">
                          {item.Cumplimiento_OK ? (
                            <span className="inline-flex items-center gap-1.5 px-3 py-1 rounded-full text-xs font-bold bg-emerald-500/10 text-emerald-400 border border-emerald-500/20">
                              Cubierto
                            </span>
                          ) : (
                            <span className="inline-flex items-center gap-1.5 px-3 py-1 rounded-full text-xs font-bold bg-amber-500/10 text-amber-400 border border-amber-500/20">
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
      </div>
    </>
  );
}