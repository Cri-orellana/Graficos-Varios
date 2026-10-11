import React, { useState, useMemo } from 'react';
import * as XLSX from 'xlsx';
import {
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer,
  PieChart, Pie, Cell, AreaChart, Area
} from 'recharts';
import {
  Upload, Download, Database, AlertTriangle, CheckCircle,
  Truck, Package, FileSpreadsheet, ShieldAlert,
  Activity, Grid, Search, FileCheck, Sliders, Calculator, PlusCircle
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

// Buscador flexible de columnas para soportar distintos formatos de Excel en la industria
const obtenerValorFlexible = (fila, posiblesNombres, valorPorDefecto = '') => {
  const llavesFila = Object.keys(fila);
  for (const nombre of posiblesNombres) {
    const llaveEncontrada = llavesFila.find(
      k => k.toLowerCase().replace(/[\s_-]/g, '') === nombre.toLowerCase().replace(/[\s_-]/g, '')
    );
    if (llaveEncontrada !== undefined && fila[llaveEncontrada] !== undefined) {
      return fila[llaveEncontrada];
    }
  }
  return valorPorDefecto;
};

export default function App() {
  const [datos, setDatos] = useState(DATOS_SIMULACION);
  const [origen, setOrigen] = useState('Datos de Simulación');
  const [vistaActiva, setVistaActiva] = useState('dashboard'); // 'dashboard' | 'auditoria' | 'simulador'

  // Filtros y controles de versatilidad
  const [plantaFiltro, setPlantaFiltro] = useState('TODAS');
  const [busqueda, setBusqueda] = useState('');
  const [soloDeficit, setSoloDeficit] = useState(false);
  const [tipoGrafico, setTipoGrafico] = useState('barras'); // 'barras' | 'area'

  // Estado del Simulador Comercial de nuevos contratos
  const [simInsumoId, setSimInsumoId] = useState('REG-101');
  const [simToneladasExtra, setSimToneladasExtra] = useState(250);

  const [auditoria, setAuditoria] = useState({
    procesados: 6,
    omitidos: 0,
    erroresDetalle: []
  });

  // Validador universal con soporte de columnas flexibles
  const validarYLimpiarDatos = (filasCrudas, nombreOrigen) => {
    const idsVistos = new Set();
    const filasValidas = [];
    const errores = [];

    filasCrudas.forEach((fila, index) => {
      const numFila = index + 2;
      const id = String(obtenerValorFlexible(fila, ['ID_Registro', 'ID', 'Codigo', 'SKU'], `REG-${100 + index}`)).trim();
      const planta = String(obtenerValorFlexible(fila, ['Planta', 'Sucursal', 'Bodega', 'Centro'], '')).trim();
      const insumo = String(obtenerValorFlexible(fila, ['Insumo', 'Producto', 'MateriaPrima', 'Material', 'Item'], '')).trim();

      const stock = Number(obtenerValorFlexible(fila, ['Stock_Silo_Tn', 'Stock', 'Silo', 'Existencia', 'Fisico'], 0));
      const transito = Number(obtenerValorFlexible(fila, ['En_Transito_GPS_Tn', 'Transito', 'GPS', 'Ruta', 'EnRuta'], 0));
      const capacidad = Number(obtenerValorFlexible(fila, ['Capacidad_Max_Tn', 'Capacidad', 'Max'], 2000));
      const contratos = Number(obtenerValorFlexible(fila, ['Contratos_Comprometidos_Tn', 'Contratos', 'Comprometido', 'Ventas', 'Demanda'], 0));

      if (!id || !planta || !insumo) {
        errores.push(`Fila ${numFila}: Campos clave vacíos (Planta o Insumo no identificados).`);
        return;
      }
      if (idsVistos.has(id)) {
        errores.push(`Fila ${numFila}: ID duplicado bloqueado (${id} - ${insumo}).`);
        return;
      }
      if (isNaN(stock) || isNaN(transito) || stock < 0 || transito < 0) {
        errores.push(`Fila ${numFila} (${id}): Toneladas negativas o corruptas (${stock} Tn).`);
        return;
      }

      idsVistos.add(id);

      filasValidas.push({
        ID_Registro: id,
        Planta: planta,
        Insumo: insumo,
        Stock_Silo_Tn: stock,
        En_Transito_GPS_Tn: transito,
        Capacidad_Max_Tn: capacidad > 0 ? capacidad : 2000,
        Contratos_Comprometidos_Tn: contratos >= 0 ? contratos : 0,
      });
    });

    setDatos(filasValidas);
    if (filasValidas.length > 0) setSimInsumoId(filasValidas[0].ID_Registro);
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

  // Editar valores directamente en la tabla (Simulación en vivo)
  const actualizarValorEnVivo = (idRegistro, campo, nuevoValor) => {
    const valorNumerico = Math.max(0, Number(nuevoValor) || 0);
    setDatos(prev => prev.map(item =>
      item.ID_Registro === idRegistro ? { ...item, [campo]: valorNumerico } : item
    ));
    setOrigen('Editado en Vivo (Simulación Activa)');
  };

  // Agregar nuevo registro rápido desde la interfaz
  const agregarFilaRapida = () => {
    const nuevoId = `REG-${Math.floor(200 + Math.random() * 799)}`;
    const nuevaFila = {
      ID_Registro: nuevoId,
      Planta: plantaFiltro === 'TODAS' ? 'Viña del Mar' : plantaFiltro,
      Insumo: 'Nuevo Insumo',
      Stock_Silo_Tn: 500,
      En_Transito_GPS_Tn: 150,
      Capacidad_Max_Tn: 1500,
      Contratos_Comprometidos_Tn: 600
    };
    setDatos(prev => [nuevaFila, ...prev]);
    setAuditoria(prev => ({ ...prev, procesados: prev.procesados + 1 }));
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
    XLSX.writeFile(libro, 'Plantilla_Prueba_Con_Errores.xlsx');
  };

  // Exportar el reporte ya limpio y con los cálculos hechos a un nuevo Excel
  const exportarReporteAuditado = () => {
    const datosExportar = datosFiltrados.map(item => ({
      ID_Registro: item.ID_Registro,
      Planta: item.Planta,
      Insumo: item.Insumo,
      Stock_Fisico_Silo_Tn: item.Stock_Silo_Tn,
      Abastecimiento_Ruta_GPS_Tn: item.En_Transito_GPS_Tn,
      Total_Proyectado_Tn: item.Stock_Proyectado_Tn,
      Contratos_Comprometidos_Tn: item.Contratos_Comprometidos_Tn,
      Balance_Final_Tn: item.Stock_Proyectado_Tn - item.Contratos_Comprometidos_Tn,
      Estado_Auditoría: item.Cumplimiento_OK ? 'CUBIERTO' : `DÉFICIT (-${item.Deficit_Tn} Tn)`,
      Ocupacion_Silo_Porcentaje: `${item.Ocupacion_Porcentaje}%`
    }));

    const hoja = XLSX.utils.json_to_sheet(datosExportar);
    const libro = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(libro, hoja, 'Reporte_Auditado');
    XLSX.writeFile(libro, `Reporte_Limpio_Soprodi_${new Date().toISOString().slice(0, 10)}.xlsx`);
  };

  const plantasDisponibles = useMemo(() => {
    const unicas = new Set(datos.map(d => d.Planta));
    return ['TODAS', ...Array.from(unicas)];
  }, [datos]);

  // Cálculo enriquecido y filtrado multicriterio
  const datosFiltrados = useMemo(() => {
    return datos
      .map(fila => {
        const stockTotalProyectado = fila.Stock_Silo_Tn + fila.En_Transito_GPS_Tn;
        const deficit = fila.Contratos_Comprometidos_Tn - stockTotalProyectado;
        const ocupacion = Math.round((stockTotalProyectado / (fila.Capacidad_Max_Tn || 1)) * 100);
        return {
          ...fila,
          Stock_Proyectado_Tn: stockTotalProyectado,
          Cumplimiento_OK: stockTotalProyectado >= fila.Contratos_Comprometidos_Tn,
          Deficit_Tn: deficit > 0 ? deficit : 0,
          Holgura_Tn: deficit < 0 ? Math.abs(deficit) : 0,
          Ocupacion_Porcentaje: ocupacion
        };
      })
      .filter(d => plantaFiltro === 'TODAS' || d.Planta === plantaFiltro)
      .filter(d => !soloDeficit || !d.Cumplimiento_OK)
      .filter(d =>
        d.Insumo.toLowerCase().includes(busqueda.toLowerCase()) ||
        d.Planta.toLowerCase().includes(busqueda.toLowerCase()) ||
        d.ID_Registro.toLowerCase().includes(busqueda.toLowerCase())
      );
  }, [datos, plantaFiltro, soloDeficit, busqueda]);

  const kpis = useMemo(() => {
    return datosFiltrados.reduce((acc, item) => {
      acc.stockSilos += item.Stock_Silo_Tn;
      acc.enTransitoGps += item.En_Transito_GPS_Tn;
      acc.contratosEnRiesgo += !item.Cumplimiento_OK ? 1 : 0;
      acc.silosCriticos += item.Ocupacion_Porcentaje >= 90 ? 1 : 0;
      return acc;
    }, { stockSilos: 0, enTransitoGps: 0, contratosEnRiesgo: 0, silosCriticos: 0 });
  }, [datosFiltrados]);

  // Datos para el Simulador Comercial
  const itemSimulado = useMemo(() => {
    const encontrado = datos.find(d => d.ID_Registro === simInsumoId) || datos[0];
    if (!encontrado) return null;
    const proyectado = encontrado.Stock_Silo_Tn + encontrado.En_Transito_GPS_Tn;
    const nuevoContratoTotal = encontrado.Contratos_Comprometidos_Tn + Number(simToneladasExtra || 0);
    const saldoFinal = proyectado - nuevoContratoTotal;
    return {
      ...encontrado,
      proyectado,
      nuevoContratoTotal,
      saldoFinal,
      viable: saldoFinal >= 0
    };
  }, [datos, simInsumoId, simToneladasExtra]);

  return (
    <>
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
            <div className="p-5 border-b border-slate-800/60 flex items-center gap-3">
              <div className="p-2.5 bg-gradient-to-tr from-blue-600 to-indigo-500 rounded-lg shadow-lg shadow-blue-500/20 text-white font-bold tracking-tight">
                SP
              </div>
              <div>
                <h2 className="font-extrabold text-white tracking-wider text-lg leading-none">SOPRODI</h2>
                <span className="text-[10px] text-blue-400 font-semibold tracking-wide uppercase">Suite Logística & BI</span>
              </div>
            </div>

            <div className="p-4">
              <span className="text-[10px] font-bold text-slate-500 tracking-wider uppercase block mb-2 px-2">Módulos Operativos</span>
              <nav className="space-y-1.5">
                <button
                  onClick={() => setVistaActiva('dashboard')}
                  className={`w-full flex items-center gap-3 px-3 py-2.5 rounded-xl text-sm font-semibold transition ${vistaActiva === 'dashboard'
                      ? 'bg-gradient-to-r from-blue-600/20 to-blue-600/5 text-blue-400 border border-blue-500/30'
                      : 'text-slate-400 hover:text-white hover:bg-slate-800/40'
                    }`}
                >
                  <Grid className="w-4 h-4" />
                  Panel de Control
                </button>

                <button
                  onClick={() => setVistaActiva('auditoria')}
                  className={`w-full flex items-center justify-between px-3 py-2.5 rounded-xl text-sm font-semibold transition ${vistaActiva === 'auditoria'
                      ? 'bg-gradient-to-r from-blue-600/20 to-blue-600/5 text-blue-400 border border-blue-500/30'
                      : 'text-slate-400 hover:text-white hover:bg-slate-800/40'
                    }`}
                >
                  <span className="flex items-center gap-3">
                    <Activity className="w-4 h-4" />
                    Auditoría & Silos
                  </span>
                  {auditoria.omitidos > 0 && (
                    <span className="px-1.5 py-0.5 text-[10px] bg-amber-500/20 text-amber-400 rounded-full border border-amber-500/30">
                      {auditoria.omitidos}
                    </span>
                  )}
                </button>

                <button
                  onClick={() => setVistaActiva('simulador')}
                  className={`w-full flex items-center gap-3 px-3 py-2.5 rounded-xl text-sm font-semibold transition ${vistaActiva === 'simulador'
                      ? 'bg-gradient-to-r from-blue-600/20 to-blue-600/5 text-blue-400 border border-blue-500/30'
                      : 'text-slate-400 hover:text-white hover:bg-slate-800/40'
                    }`}
                >
                  <Calculator className="w-4 h-4" />
                  Simulador Comercial
                </button>
              </nav>
            </div>
          </div>

          {/* Tarjeta Resumen de Auditoría en Lateral */}
          <div className="p-4 border-t border-slate-800 bg-slate-950/50 m-3 rounded-2xl">
            <div className="flex items-center gap-2 mb-3">
              <ShieldAlert className="w-4 h-4 text-emerald-400" />
              <h3 className="font-bold text-xs text-slate-300 uppercase">Integridad de Datos</h3>
            </div>
            <div className="grid grid-cols-2 gap-2 text-center">
              <div className="bg-slate-900 border border-slate-800 rounded-xl p-2">
                <span className="text-[10px] font-medium text-slate-500 uppercase block">Válidos</span>
                <span className="text-lg font-bold text-emerald-400">{auditoria.procesados}</span>
              </div>
              <div className="bg-slate-900 border border-slate-800 rounded-xl p-2">
                <span className="text-[10px] font-medium text-slate-500 uppercase block">Bloqueados</span>
                <span className={`text-lg font-bold ${auditoria.omitidos > 0 ? 'text-amber-400' : 'text-slate-500'}`}>
                  {auditoria.omitidos}
                </span>
              </div>
            </div>
            <p className="text-[10px] text-slate-500 mt-2.5 leading-relaxed">
              💡 Puedes **editar cualquier número directamente en la tabla** para simular escenarios en vivo.
            </p>
          </div>
        </aside>

        {/* Contenido Principal */}
        <div className="flex-1 flex flex-col min-w-0 h-full overflow-y-auto overflow-x-hidden bg-slate-950">
          <header className="bg-slate-900/60 backdrop-blur-md border-b border-slate-800 px-6 py-4 sticky top-0 z-30 flex flex-wrap justify-between items-center gap-4">
            <div className="min-w-0">
              <div className="flex items-center gap-2 mb-1">
                <span className="inline-flex items-center gap-1.5 px-2.5 py-0.5 rounded-full text-xs font-semibold bg-blue-500/10 text-blue-400 border border-blue-500/20">
                  <Database className="w-3 h-3" /> {origen}
                </span>
                <span className="text-xs text-slate-500">•</span>
                <span className="text-xs text-slate-400 font-medium">Mapeo Inteligente de Columnas Activo</span>
              </div>
              <h1 className="text-xl md:text-2xl font-black text-white tracking-tight">
                Monitor de Existencias, Flotas GPS y Contratos
              </h1>
            </div>

            {/* Botonera Superior Versátil */}
            <div className="flex flex-wrap items-center gap-2">
              <label className="cursor-pointer bg-blue-600 hover:bg-blue-500 text-white px-3.5 py-2 rounded-xl font-bold text-xs flex items-center gap-2 shadow-lg shadow-blue-600/20 transition border border-blue-500/30">
                <Upload className="w-3.5 h-3.5" />
                Cargar Excel / CSV
                <input type="file" accept=".xlsx, .xls, .csv" onChange={manejarSubidaExcel} className="hidden" />
              </label>

              <button
                onClick={exportarReporteAuditado}
                className="bg-indigo-600/20 hover:bg-indigo-600/30 text-indigo-300 px-3.5 py-2 rounded-xl font-bold text-xs flex items-center gap-2 border border-indigo-500/30 transition"
                title="Descarga un archivo Excel limpio con los cálculos actuales"
              >
                <FileCheck className="w-3.5 h-3.5" />
                Exportar Limpio (.xlsx)
              </button>

              <button
                onClick={descargarPlantillaPrueba}
                className="bg-emerald-600/10 hover:bg-emerald-600/20 text-emerald-400 px-3.5 py-2 rounded-xl font-bold text-xs flex items-center gap-2 border border-emerald-500/20 transition"
              >
                <Download className="w-3.5 h-3.5" />
                Plantilla Demo
              </button>

              <button
                onClick={() => validarYLimpiarDatos(DATOS_SIMULACION, 'Datos de Simulación')}
                className="bg-slate-800 hover:bg-slate-700 text-slate-300 px-3 py-2 rounded-xl font-bold text-xs transition"
              >
                Restablecer
              </button>
            </div>
          </header>

          {/* Barra de Filtros Interactivos */}
          <div className="bg-slate-900/30 border-b border-slate-800/80 px-6 py-3 flex flex-wrap items-center justify-between gap-4">
            <div className="flex flex-wrap items-center gap-3 flex-1 min-w-[260px]">
              {/* Buscador en tiempo real */}
              <div className="relative flex-1 max-w-xs">
                <Search className="w-3.5 h-3.5 text-slate-500 absolute left-3 top-1/2 -translate-y-1/2" />
                <input
                  type="text"
                  value={busqueda}
                  onChange={(e) => setBusqueda(e.target.value)}
                  placeholder="Buscar insumo, planta o ID..."
                  className="w-full bg-slate-900 border border-slate-800 rounded-xl pl-9 pr-3 py-1.5 text-xs text-slate-200 placeholder-slate-500 focus:outline-none focus:border-blue-500"
                />
              </div>

              {/* Selector de Planta */}
              <div className="flex items-center gap-2 bg-slate-900 px-3 py-1.5 rounded-xl border border-slate-800">
                <span className="text-[10px] font-bold text-slate-400 uppercase">Planta:</span>
                <select
                  value={plantaFiltro}
                  onChange={(e) => setPlantaFiltro(e.target.value)}
                  className="text-xs font-bold text-blue-400 bg-transparent focus:outline-none cursor-pointer"
                >
                  {plantasDisponibles.map(planta => (
                    <option key={planta} value={planta} className="bg-slate-900 text-slate-200">{planta}</option>
                  ))}
                </select>
              </div>

              {/* Switch Solo Déficit */}
              <button
                onClick={() => setSoloDeficit(!soloDeficit)}
                className={`px-3 py-1.5 rounded-xl text-xs font-bold border transition flex items-center gap-1.5 ${soloDeficit
                    ? 'bg-amber-500/20 text-amber-300 border-amber-500/40'
                    : 'bg-slate-900 text-slate-400 border-slate-800 hover:text-slate-200'
                  }`}
              >
                <AlertTriangle className="w-3.5 h-3.5" />
                Solo en Déficit
              </button>
            </div>

            {/* Selector de Tipo de Gráfico */}
            <div className="flex items-center gap-1 bg-slate-900 p-1 rounded-xl border border-slate-800">
              <button
                onClick={() => setTipoGrafico('barras')}
                className={`px-2.5 py-1 rounded-lg text-xs font-semibold transition ${tipoGrafico === 'barras' ? 'bg-blue-600 text-white' : 'text-slate-400 hover:text-white'
                  }`}
              >
                Barras Apiladas
              </button>
              <button
                onClick={() => setTipoGrafico('area')}
                className={`px-2.5 py-1 rounded-lg text-xs font-semibold transition ${tipoGrafico === 'area' ? 'bg-blue-600 text-white' : 'text-slate-400 hover:text-white'
                  }`}
              >
                Curva de Área
              </button>
            </div>
          </div>

          {/* Cuerpo Dinámico según la Vista Seleccionada */}
          <main className="p-6 space-y-6 flex-1 min-w-0">
            {/* VISTA 1: PANEL DE CONTROL PRINCIPAL */}
            {vistaActiva === 'dashboard' && (
              <>
                {auditoria.omitidos > 0 && (
                  <div className="bg-amber-500/10 border border-amber-500/20 p-4 rounded-2xl flex items-center justify-between gap-3">
                    <div className="flex items-start gap-3">
                      <ShieldAlert className="w-5 h-5 text-amber-500 mt-0.5 shrink-0" />
                      <div>
                        <h4 className="font-bold text-sm text-amber-400">Auditoría Activa: {auditoria.omitidos} filas corruptas o duplicadas fueron bloqueadas</h4>
                        <p className="text-xs text-slate-400 mt-0.5">
                          {auditoria.erroresDetalle.join(' | ')}
                        </p>
                      </div>
                    </div>
                    <button
                      onClick={() => setVistaActiva('auditoria')}
                      className="text-xs font-bold text-amber-400 underline shrink-0"
                    >
                      Ver detalle
                    </button>
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
                        Disponibilidad inmediata
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
                        Proyectado Total: <span className="text-emerald-400 font-bold">{(kpis.stockSilos + kpis.enTransitoGps).toLocaleString('es-CL')} Tn</span>
                      </p>
                    </div>
                    <div className="p-3.5 bg-emerald-500/10 border border-emerald-500/20 rounded-xl text-emerald-400 shrink-0">
                      <Truck className="w-6 h-6" />
                    </div>
                  </div>

                  <div className="bg-slate-900 border border-slate-800 p-5 rounded-2xl flex items-center justify-between">
                    <div>
                      <p className="text-xs font-semibold text-slate-400 uppercase tracking-wide">Contratos en Riesgo</p>
                      <p className={`text-3xl font-black mt-1 font-mono ${kpis.contratosEnRiesgo > 0 ? 'text-amber-400' : 'text-emerald-400'}`}>
                        {kpis.contratosEnRiesgo} <span className="text-sm font-normal text-slate-500">insumos en déficit</span>
                      </p>
                      <p className="text-[11px] text-slate-400 mt-2 font-semibold uppercase">
                        {kpis.silosCriticos} silos sobre 90% de capacidad
                      </p>
                    </div>
                    <div className={`p-3.5 rounded-xl border shrink-0 ${kpis.contratosEnRiesgo > 0 ? 'bg-amber-500/10 border-amber-500/20 text-amber-400' : 'bg-emerald-500/10 border-emerald-500/20 text-emerald-400'}`}>
                      <AlertTriangle className="w-6 h-6" />
                    </div>
                  </div>
                </div>

                {/* Gráficos Dinámicos */}
                <div className="grid grid-cols-1 lg:grid-cols-3 gap-5">
                  <div className="bg-slate-900 border border-slate-800/80 p-5 rounded-2xl lg:col-span-2 min-w-0">
                    <div className="mb-4">
                      <h2 className="text-base font-extrabold text-white">Cruce de Existencias (Silos + GPS) vs Contratos Comerciales</h2>
                      <p className="text-xs text-slate-400 mt-0.5">Modifica los números en la tabla inferior para simular llegadas de camiones o nuevas ventas.</p>
                    </div>

                    <div className="h-72 w-full">
                      <ResponsiveContainer width="100%" height="100%">
                        {tipoGrafico === 'barras' ? (
                          <BarChart data={datosFiltrados} margin={{ top: 10, right: 10, left: -15, bottom: 0 }}>
                            <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#1e293b" />
                            <XAxis dataKey="Insumo" stroke="#64748b" fontSize={11} tickLine={false} />
                            <YAxis stroke="#64748b" fontSize={11} tickLine={false} />
                            <Tooltip contentStyle={{ backgroundColor: '#0f172a', borderColor: '#334155', borderRadius: '12px' }} />
                            <Legend wrapperStyle={{ fontSize: '11px', paddingTop: '10px' }} />
                            <Bar dataKey="Stock_Silo_Tn" name="Stock en Silo (Tn)" stackId="a" fill="#3b82f6" radius={[0, 0, 4, 4]} />
                            <Bar dataKey="En_Transito_GPS_Tn" name="Tránsito GPS (Tn)" stackId="a" fill="#10b981" radius={[4, 4, 0, 0]} />
                            <Bar dataKey="Contratos_Comprometidos_Tn" name="Contrato Comprometido (Tn)" fill="#f59e0b" radius={[4, 4, 4, 4]} />
                          </BarChart>
                        ) : (
                          <AreaChart data={datosFiltrados} margin={{ top: 10, right: 10, left: -15, bottom: 0 }}>
                            <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#1e293b" />
                            <XAxis dataKey="Insumo" stroke="#64748b" fontSize={11} />
                            <YAxis stroke="#64748b" fontSize={11} />
                            <Tooltip contentStyle={{ backgroundColor: '#0f172a', borderColor: '#334155', borderRadius: '12px' }} />
                            <Legend wrapperStyle={{ fontSize: '11px', paddingTop: '10px' }} />
                            <Area type="monotone" dataKey="Stock_Proyectado_Tn" name="Stock Total Proyectado (Silo+GPS)" stroke="#10b981" fill="#10b981" fillOpacity={0.25} />
                            <Area type="monotone" dataKey="Contratos_Comprometidos_Tn" name="Contratos Comprometidos" stroke="#f59e0b" fill="#f59e0b" fillOpacity={0.2} />
                          </AreaChart>
                        )}
                      </ResponsiveContainer>
                    </div>
                  </div>

                  <div className="bg-slate-900 border border-slate-800/80 p-5 rounded-2xl min-w-0 flex flex-col justify-between">
                    <div>
                      <h2 className="text-base font-extrabold text-white mb-1">Distribución de Volumen</h2>
                      <p className="text-xs text-slate-400 mb-2">Participación por materia prima proyectada.</p>
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
                          <Tooltip contentStyle={{ backgroundColor: '#0f172a', borderColor: '#334155', borderRadius: '12px' }} />
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

                {/* Tabla Interactiva con Edición "What-If" en Vivo */}
                <div className="bg-slate-900 border border-slate-800 rounded-2xl overflow-hidden shadow-xl">
                  <div className="p-5 border-b border-slate-800 flex flex-wrap items-center justify-between gap-3">
                    <div className="flex items-center gap-3">
                      <FileSpreadsheet className="w-5 h-5 text-blue-400" />
                      <div>
                        <h2 className="text-base font-extrabold text-white">Matriz Interactiva de Existencias y Contratos (Edición en Vivo)</h2>
                        <p className="text-xs text-slate-400 mt-0.5">Haz clic en las celdas numéricas para simular llegada de camiones o aumento de demanda.</p>
                      </div>
                    </div>
                    <button
                      onClick={agregarFilaRapida}
                      className="bg-blue-600/15 hover:bg-blue-600/25 text-blue-400 px-3 py-1.5 rounded-xl text-xs font-bold border border-blue-500/30 flex items-center gap-1.5 transition"
                    >
                      <PlusCircle className="w-3.5 h-3.5" />
                      Añadir Fila de Prueba
                    </button>
                  </div>
                  <div className="overflow-x-auto">
                    <table className="w-full text-left border-collapse">
                      <thead>
                        <tr className="bg-slate-950/40 text-slate-400 text-xs font-bold uppercase border-b border-slate-800/80">
                          <th className="py-3.5 px-4">ID</th>
                          <th className="py-3.5 px-4">Planta</th>
                          <th className="py-3.5 px-4">Materia Prima</th>
                          <th className="py-3.5 px-4 text-right">Silo Físico (Editable)</th>
                          <th className="py-3.5 px-4 text-right">Ruta GPS (Editable)</th>
                          <th className="py-3.5 px-4 text-right">Total Proyectado</th>
                          <th className="py-3.5 px-4 text-right">Contratos (Editable)</th>
                          <th className="py-3.5 px-4 text-center">Ocupación Silo</th>
                          <th className="py-3.5 px-4 text-center">Estado Contrato</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-slate-800 text-sm">
                        {datosFiltrados.map((item) => (
                          <tr key={item.ID_Registro} className="hover:bg-slate-800/20 transition">
                            <td className="py-3 px-4 font-mono text-xs text-slate-500">{item.ID_Registro}</td>
                            <td className="py-3 px-4 font-medium text-slate-300">{item.Planta}</td>
                            <td className="py-3 px-4 font-bold text-white">{item.Insumo}</td>
                            <td className="py-3 px-4 text-right">
                              <input
                                type="number"
                                value={item.Stock_Silo_Tn}
                                onChange={(e) => actualizarValorEnVivo(item.ID_Registro, 'Stock_Silo_Tn', e.target.value)}
                                className="w-24 bg-slate-950 border border-slate-700/80 focus:border-blue-500 rounded-lg px-2 py-1 text-right font-mono text-white text-xs"
                              />
                            </td>
                            <td className="py-3 px-4 text-right">
                              <input
                                type="number"
                                value={item.En_Transito_GPS_Tn}
                                onChange={(e) => actualizarValorEnVivo(item.ID_Registro, 'En_Transito_GPS_Tn', e.target.value)}
                                className="w-24 bg-slate-950 border border-emerald-700/50 focus:border-emerald-400 rounded-lg px-2 py-1 text-right font-mono text-emerald-400 text-xs"
                              />
                            </td>
                            <td className="py-3 px-4 text-right font-mono font-bold text-white">
                              {item.Stock_Proyectado_Tn.toLocaleString('es-CL')} Tn
                            </td>
                            <td className="py-3 px-4 text-right">
                              <input
                                type="number"
                                value={item.Contratos_Comprometidos_Tn}
                                onChange={(e) => actualizarValorEnVivo(item.ID_Registro, 'Contratos_Comprometidos_Tn', e.target.value)}
                                className="w-24 bg-slate-950 border border-amber-700/50 focus:border-amber-400 rounded-lg px-2 py-1 text-right font-mono text-amber-300 text-xs"
                              />
                            </td>
                            <td className="py-3 px-4">
                              <div className="flex items-center gap-2 justify-center">
                                <div className="w-16 h-2 bg-slate-800 rounded-full overflow-hidden">
                                  <div
                                    className={`h-full ${item.Ocupacion_Porcentaje > 90 ? 'bg-red-500' : 'bg-blue-500'}`}
                                    style={{ width: `${Math.min(100, item.Ocupacion_Porcentaje)}%` }}
                                  />
                                </div>
                                <span className="font-mono text-xs text-slate-400">{item.Ocupacion_Porcentaje}%</span>
                              </div>
                            </td>
                            <td className="py-3 px-4 text-center">
                              {item.Cumplimiento_OK ? (
                                <span className="inline-flex items-center gap-1 px-2.5 py-1 rounded-full text-xs font-bold bg-emerald-500/10 text-emerald-400 border border-emerald-500/20">
                                  Cubierto (+{item.Holgura_Tn} Tn)
                                </span>
                              ) : (
                                <span className="inline-flex items-center gap-1 px-2.5 py-1 rounded-full text-xs font-bold bg-amber-500/10 text-amber-400 border border-amber-500/20">
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
              </>
            )}

            {/* VISTA 2: CONSOLA DE AUDITORÍA Y CAPACIDAD DE SILOS */}
            {vistaActiva === 'auditoria' && (
              <div className="space-y-6">
                <div className="bg-slate-900 border border-slate-800 p-6 rounded-2xl">
                  <h2 className="text-lg font-bold text-white mb-2 flex items-center gap-2">
                    <ShieldAlert className="w-5 h-5 text-emerald-400" />
                    Registro de Limpieza y Validación de Datos (Data Quality)
                  </h2>
                  <p className="text-xs text-slate-400 mb-4">
                    Este módulo garantiza que ninguna planilla Excel con errores humanos corrompa los indicadores de gerencia.
                  </p>
                  {auditoria.erroresDetalle.length > 0 ? (
                    <div className="space-y-2">
                      {auditoria.erroresDetalle.map((err, idx) => (
                        <div key={idx} className="p-3 bg-amber-500/10 border border-amber-500/30 rounded-xl text-xs text-amber-300 font-mono">
                          ⚠️ {err}
                        </div>
                      ))}
                    </div>
                  ) : (
                    <div className="p-4 bg-emerald-500/10 border border-emerald-500/20 rounded-xl text-xs text-emerald-300">
                      ✅ Todos los registros actuales cumplen con validación de Primary Key (ID único) y tonelajes positivos. Descarga la "Plantilla Demo" y súbela para poner a prueba el bloqueo de errores.
                    </div>
                  )}
                </div>

                <div className="bg-slate-900 border border-slate-800 p-6 rounded-2xl">
                  <h2 className="text-lg font-bold text-white mb-4">Capacidad de Almacenamiento por Silo / Planta</h2>
                  <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                    {datosFiltrados.map(item => (
                      <div key={item.ID_Registro} className="p-4 bg-slate-950 border border-slate-800 rounded-xl">
                        <div className="flex justify-between text-sm font-bold mb-1">
                          <span>{item.Insumo} ({item.Planta})</span>
                          <span className={item.Ocupacion_Porcentaje > 90 ? 'text-red-400' : 'text-blue-400'}>
                            {item.Stock_Proyectado_Tn} / {item.Capacidad_Max_Tn} Tn ({item.Ocupacion_Porcentaje}%)
                          </span>
                        </div>
                        <div className="w-full h-3 bg-slate-800 rounded-full overflow-hidden">
                          <div
                            className={`h-full transition-all ${item.Ocupacion_Porcentaje > 90 ? 'bg-red-500' : 'bg-blue-500'}`}
                            style={{ width: `${Math.min(100, item.Ocupacion_Porcentaje)}%` }}
                          />
                        </div>
                      </div>
                    ))}
                  </div>
                </div>
              </div>
            )}

            {/* VISTA 3: SIMULADOR COMERCIAL DE NUEVOS CONTRATOS */}
            {vistaActiva === 'simulador' && itemSimulado && (
              <div className="bg-slate-900 border border-slate-800 p-6 rounded-2xl max-w-3xl">
                <div className="flex items-center gap-2 text-blue-400 text-xs font-bold uppercase mb-1">
                  <Calculator className="w-4 h-4" /> Toma de Decisiones Comerciales
                </div>
                <h2 className="text-xl font-black text-white mb-2">Simulador de Factibilidad de Nuevo Contrato</h2>
                <p className="text-xs text-slate-400 mb-6">
                  Verifica en segundos si una planta puede comprometer una nueva venta considerando el stock en silo más los camiones GPS en ruta.
                </p>

                <div className="grid grid-cols-1 md:grid-cols-2 gap-4 mb-6">
                  <div>
                    <label className="block text-xs font-bold text-slate-400 uppercase mb-2">Seleccionar Materia Prima / Planta</label>
                    <select
                      value={simInsumoId}
                      onChange={(e) => setSimInsumoId(e.target.value)}
                      className="w-full bg-slate-950 border border-slate-800 rounded-xl p-3 text-sm text-white font-semibold"
                    >
                      {datos.map(d => (
                        <option key={d.ID_Registro} value={d.ID_Registro}>
                          {d.Insumo} — {d.Planta} (Disp: {d.Stock_Silo_Tn + d.En_Transito_GPS_Tn} Tn)
                        </option>
                      ))}
                    </select>
                  </div>

                  <div>
                    <label className="block text-xs font-bold text-slate-400 uppercase mb-2">Toneladas del Nuevo Pedido Cliente</label>
                    <input
                      type="number"
                      value={simToneladasExtra}
                      onChange={(e) => setSimToneladasExtra(e.target.value)}
                      className="w-full bg-slate-950 border border-blue-500/40 rounded-xl p-3 text-sm text-white font-mono font-bold"
                    />
                  </div>
                </div>

                <div className={`p-5 rounded-2xl border ${itemSimulado.viable ? 'bg-emerald-500/10 border-emerald-500/30' : 'bg-red-500/10 border-red-500/30'}`}>
                  <h3 className={`text-base font-black ${itemSimulado.viable ? 'text-emerald-400' : 'text-red-400'}`}>
                    {itemSimulado.viable
                      ? '✅ CONTRATO VIABLE: Stock proyectado suficiente para cubrir la venta'
                      : '❌ RIESGO DE QUIEBRE: El pedido supera la suma de Silo + Flota GPS'}
                  </h3>
                  <div className="grid grid-cols-3 gap-4 mt-4 text-xs">
                    <div>
                      <span className="text-slate-400 block">Stock Total (Silo+GPS)</span>
                      <span className="text-lg font-mono font-bold text-white">{itemSimulado.proyectado} Tn</span>
                    </div>
                    <div>
                      <span className="text-slate-400 block">Contratos Actuales + Nuevo</span>
                      <span className="text-lg font-mono font-bold text-amber-400">{itemSimulado.nuevoContratoTotal} Tn</span>
                    </div>
                    <div>
                      <span className="text-slate-400 block">Balance Resultante</span>
                      <span className={`text-lg font-mono font-bold ${itemSimulado.viable ? 'text-emerald-400' : 'text-red-400'}`}>
                        {itemSimulado.saldoFinal > 0 ? `+${itemSimulado.saldoFinal}` : itemSimulado.saldoFinal} Tn
                      </span>
                    </div>
                  </div>
                </div>
              </div>
            )}
          </main>
        </div>
      </div>
    </>
  );
}