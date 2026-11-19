/* Períodos estándar de La Asturiana para los filtros de fecha de una página.
   Pegar tal cual en un <script>. Uso:
     Periodos.lista(new Date())  ->  [{ clave, etiqueta, desde, hasta, parcial }, ...]
   desde y hasta son Date locales (sin hora), ambos inclusive.
   Quincena: 1Q = días 1 al 15, 2Q = del 16 al último día del mes. */

var Periodos = (function () {
  var MESES = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio',
               'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'];

  function dia(a, m, d) { return new Date(a, m, d); }
  function ultimoDia(a, m) { return new Date(a, m + 1, 0).getDate(); }
  function dosDig(n) { return (n < 10 ? '0' : '') + n; }

  function fecha(d) {
    return dosDig(d.getDate()) + '/' + dosDig(d.getMonth() + 1) + '/' + d.getFullYear();
  }
  function nombreMes(a, m) {
    return MESES[m].charAt(0).toUpperCase() + MESES[m].slice(1) + ' ' + a;
  }

  // La quincena que contiene la fecha d.
  function quincenaDe(d) {
    var a = d.getFullYear(), m = d.getMonth(), q = d.getDate() <= 15 ? 1 : 2;
    return {
      nombre: dosDig(m + 1) + '-' + q + 'Q ' + a,
      desde: dia(a, m, q === 1 ? 1 : 16),
      hasta: dia(a, m, q === 1 ? 15 : ultimoDia(a, m))
    };
  }

  function lista(hoy) {
    var h = dia(hoy.getFullYear(), hoy.getMonth(), hoy.getDate());
    var a = h.getFullYear(), m = h.getMonth();

    var qAct = quincenaDe(h);
    var qAnt = quincenaDe(dia(qAct.desde.getFullYear(), qAct.desde.getMonth(), qAct.desde.getDate() - 1));
    var mAntDesde = dia(a, m - 1, 1);
    var mAntHasta = dia(a, m, 0);
    var tresDesde = dia(a, m - 3, 1);

    return [
      { clave: 'quincena-actual', tipo: 'quincena', etiqueta: 'Quincena en curso (' + qAct.nombre + ')',
        desde: qAct.desde, hasta: h, parcial: h < qAct.hasta },
      { clave: 'quincena-anterior', tipo: 'quincena', etiqueta: 'Quincena anterior (' + qAnt.nombre + ')',
        desde: qAnt.desde, hasta: qAnt.hasta, parcial: false },
      { clave: 'mes-actual', tipo: 'mes', etiqueta: 'Mes en curso (' + nombreMes(a, m) + ')',
        desde: dia(a, m, 1), hasta: h, parcial: h.getDate() < ultimoDia(a, m) },
      { clave: 'mes-anterior', tipo: 'mes',
        etiqueta: 'Mes anterior (' + nombreMes(mAntDesde.getFullYear(), mAntDesde.getMonth()) + ')',
        desde: mAntDesde, hasta: mAntHasta, parcial: false },
      { clave: 'ultimos-3-meses', tipo: 'rango', etiqueta: 'Últimos 3 meses cerrados',
        desde: tresDesde, hasta: mAntHasta, parcial: false },
      { clave: 'anio-actual', tipo: 'anio', etiqueta: 'Año en curso (' + a + ')',
        desde: dia(a, 0, 1), hasta: h, parcial: !(m === 11 && h.getDate() === 31) },
      { clave: 'rango', tipo: 'rango', etiqueta: 'Elegir fechas…', desde: null, hasta: null, parcial: false }
    ];
  }

  // Mismo tramo del período anterior, para comparar un período en curso:
  // del 1 al 12 de noviembre se compara con el 1 al 12 de octubre, no con octubre entero.
  // tipo: 'quincena', 'mes' o 'anio' (el campo tipo de cada período de lista()).
  function mismoTramoAnterior(desde, hasta, tipo) {
    var a = desde.getFullYear(), m = desde.getMonth();
    var largo = hasta.getDate() - desde.getDate();
    var ini, tope;
    if (tipo === 'quincena' && desde.getDate() === 16) {
      ini = dia(a, m, 1);  tope = 15;
    } else if (tipo === 'quincena') {
      ini = dia(a, m - 1, 16);  tope = ultimoDia(ini.getFullYear(), ini.getMonth());
    } else if (tipo === 'mes') {
      ini = dia(a, m - 1, 1);  tope = ultimoDia(ini.getFullYear(), ini.getMonth());
    } else {
      var mh = hasta.getMonth();
      return { desde: dia(a - 1, 0, 1),
               hasta: dia(a - 1, mh, Math.min(hasta.getDate(), ultimoDia(a - 1, mh))) };
    }
    // Si el período ya está completo, se compara contra el anterior completo.
    var finDelPeriodo = tipo === 'quincena' && desde.getDate() === 1 ? 15 : ultimoDia(a, m);
    var fin = hasta.getDate() === finDelPeriodo ? tope : Math.min(ini.getDate() + largo, tope);
    return { desde: ini, hasta: dia(ini.getFullYear(), ini.getMonth(), fin) };
  }

  return { lista: lista, quincenaDe: quincenaDe, mismoTramoAnterior: mismoTramoAnterior,
           fecha: fecha, nombreMes: nombreMes };
})();

if (typeof module !== 'undefined') { module.exports = Periodos; }
