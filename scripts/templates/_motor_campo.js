/* Motor de "Asistencias a campo" para el filtro de fechas.

   La regla de oro: ACA NO SE DECIDE NADA. El robot (motor_agregacion.py) ya
   resolvio en Python a que salida pertenece cada linea y cual no cuenta como
   asistencia, y mando las filas con esa decision adentro. Este archivo solo
   filtra por fecha, agrupa y suma.

   Si alguna vez hay que escribir aca un criterio del negocio -- que es un
   repuesto, que maquina es de terceros, que no cuenta como asistencia -- es
   que algo salio mal: esa regla quedaria escrita dos veces, en Python y aca,
   y el dia que cambie una sola de las dos el informe sigue mostrando numeros
   pero deja de ser el mismo calculo. Va en Python y viene resuelta en la fila.

   Columnas de cada fila (ver exportar_filas_campo):
     0 d    dias desde datos.desde
     1 sal  indice de la salida
     2 mec  3 sup  4 maq  5 sub  6 rub  7 lug  8 con   indices del diccionario
     9 h   10 ht  11 hp   horas de trabajo, traslado y preparacion
    12 x    1 = no cuenta como asistencia                                      */

var MotorCampo = (function () {
  var C = null;
  var base = null;

  function aDia(iso) {
    var p = iso.split('-');
    return new Date(+p[0], +p[1] - 1, +p[2]);
  }
  function offset(fecha) {
    return Math.round((fecha - base) / 86400000);
  }
  function fechaDe(off) {
    return new Date(base.getTime() + off * 86400000);
  }
  function dosDig(n) { return (n < 10 ? '0' : '') + n; }
  function ddmmaaaa(d) {
    return dosDig(d.getDate()) + '/' + dosDig(d.getMonth() + 1) + '/' + d.getFullYear();
  }
  function r1(x) { return Math.round(x * 10) / 10; }

  function cargar(datos) {
    C = datos;
    base = C && C.desde ? aDia(C.desde) : null;
    return !!(C && C.filas && C.filas.length);
  }

  function hayDatos() { return !!(C && C.filas && C.filas.length); }

  /* Ultimo dia con datos cargados. La regla de la casa es decir hasta cuando
     hay datos y no suponer que es hoy. */
  function ultimoDia() { return C && C.hasta ? aDia(C.hasta) : null; }
  function primerDia() { return base; }

  function texto(dic, i) { return i >= 0 ? C.dic[dic][i] : null; }

  /* Mismo armado que construir_top_supervisores() en Python: supervisor ->
     fecha -> salida -> lineas. */
  function supervisores(desde, hasta) {
    if (!hayDatos()) return [];
    var d0 = offset(desde), d1 = offset(hasta);
    var F = C.filas, sups = {}, i;

    for (i = 0; i < F.length; i++) {
      var r = F[i];
      if (r[0] < d0 || r[0] > d1) continue;
      if (r[12] === 1) continue;        // lo descarto el robot, no se rediscute
      if (r[3] < 0) continue;           // sin supervisor cargado

      var nombre = texto('sup', r[3]);
      var s = sups[nombre];
      if (!s) {
        s = sups[nombre] = { n: nombre, salidas: 0, maqs: {}, t: 0, tr: 0, p: 0,
                             porClave: {}, orden: [] };
      }
      s.t += r[9]; s.tr += r[10]; s.p += r[11];
      if (r[4] >= 0) s.maqs[r[4]] = 1;

      var clave = r[1] + '~' + r[0];
      var sal = s.porClave[clave];
      if (!sal) {
        sal = s.porClave[clave] = { d: r[0], mec: '', maqs: {}, hs: 0, hst: 0, hsp: 0,
                                    lineas: [] };
        s.orden.push(clave);
        s.salidas++;
      }
      if (!sal.mec && r[2] >= 0) sal.mec = texto('mec', r[2]);
      if (r[4] >= 0) sal.maqs[r[4]] = 1;
      sal.hs += r[9]; sal.hst += r[10]; sal.hsp += r[11];
      sal.lineas.push({ maq: texto('maq', r[4]) || '', sub: texto('sub', r[5]),
                        lug: texto('lug', r[7]),
                        hs: r1(r[9]), hst: r1(r[10]), hsp: r1(r[11]) });
    }

    var salida = [];
    Object.keys(sups).sort().forEach(function (nombre) {
      var s = sups[nombre], porFecha = {}, claves = [];
      s.orden.forEach(function (clave) {
        var sal = s.porClave[clave];
        sal.lineas.sort(function (a, b) { return (a.maq || '') < (b.maq || '') ? -1 : 1; });
        var f = fechaDe(sal.d);
        var k = f.getFullYear() + '-' + dosDig(f.getMonth() + 1) + '-' + dosDig(f.getDate());
        if (!porFecha[k]) { porFecha[k] = { f: ddmmaaaa(f), salidas: [] }; claves.push(k); }
        porFecha[k].salidas.push({
          mec: sal.mec, maqs: Object.keys(sal.maqs).length,
          hs: r1(sal.hs), hst: r1(sal.hst), hsp: r1(sal.hsp), lineas: sal.lineas
        });
      });
      // Fechas de la mas reciente a la mas vieja; adentro, la salida mas larga
      // primero, igual que en Python.
      claves.sort().reverse();
      var fechas = claves.map(function (k) {
        porFecha[k].salidas.sort(function (a, b) {
          return (b.hs + b.hst + b.hsp) - (a.hs + a.hst + a.hsp);
        });
        return porFecha[k];
      });
      salida.push({ n: s.n, salidas: s.salidas, maqs: Object.keys(s.maqs).length,
                    t: r1(s.t), tr: r1(s.tr), p: r1(s.p), fechas: fechas });
    });
    return salida;
  }

  /* Autocontrol: el robot ya calculo la semana por defecto con sus propias
     cuentas. Si lo que sale de aca para ese mismo rango no coincide, hay algo
     desincronizado y la pagina lo tiene que gritar, no disimularlo.
     Las horas se comparan con tolerancia: Python redondea al par y JavaScript
     hacia arriba, y 0,05 de diferencia no es un error de criterio. */
  function comparar(robot, desde, hasta) {
    var mio = supervisores(desde, hasta);
    var dif = [];
    var porNombre = {};
    mio.forEach(function (s) { porNombre[s.n] = s; });
    (robot || []).forEach(function (r) {
      var m = porNombre[r.n];
      if (!m) { dif.push(r.n + ': no aparece'); return; }
      if (m.salidas !== r.salidas) dif.push(r.n + ': ' + r.salidas + ' vs ' + m.salidas + ' salidas');
      if (m.maqs !== r.maqs) dif.push(r.n + ': ' + r.maqs + ' vs ' + m.maqs + ' máquinas');
      ['t', 'tr', 'p'].forEach(function (c) {
        if (Math.abs(m[c] - r[c]) > 0.051) dif.push(r.n + ': ' + r[c] + ' vs ' + m[c] + ' hs (' + c + ')');
      });
    });
    if (mio.length !== (robot || []).length) {
      dif.push('supervisores: ' + (robot || []).length + ' vs ' + mio.length);
    }
    return dif;
  }

  return { cargar: cargar, hayDatos: hayDatos, supervisores: supervisores,
           comparar: comparar, ultimoDia: ultimoDia, primerDia: primerDia,
           ddmmaaaa: ddmmaaaa, aDia: aDia };
})();
