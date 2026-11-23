/**
 * ALIMENTACIÓN SOLIDARIA - Backend v1.1 (unificado)
 * Google Apps Script + Google Sheets
 *
 * Ejecutar primero: configurarSistema()
 * Luego: crearAdministradorInicial("correo@ejemplo.com", "ContraseñaSegura", "Administrador")
 *
 * Para HTMLService, llamar estas funciones con google.script.run.
 * No publicar contraseñas ni tokens en hojas visibles.
 */

const AS = {
  VERSION: "1.1.0",
  TOKEN_TTL_SECONDS: 21600, // 6 horas
  MAX_INTENTOS_LOGIN: 5,
  BLOQUEO_LOGIN_SECONDS: 900, // 15 minutos
  SHEETS: {
    USUARIOS: ["id", "nombre", "correo", "password_hash", "salt", "rol", "estado", "fecha_creacion"],
    COMEDORES: ["id", "usuario_id", "nombre", "responsable", "telefono", "direccion", "distrito", "beneficiarios", "estado", "fecha_creacion"],
    DONANTES: ["id", "usuario_id", "nombre", "tipo", "contacto", "telefono", "direccion", "estado", "fecha_creacion"],
    ALIMENTOS: ["id", "nombre", "categoria", "unidad", "descripcion", "estado", "fecha_creacion"],
    DONACIONES: ["id", "donante_id", "alimento_id", "cantidad_publicada", "cantidad_disponible", "unidad", "fecha_vencimiento", "fecha_recojo", "direccion_recojo", "observaciones", "estado", "fecha_creacion"],
    SOLICITUDES: ["id", "donacion_id", "comedor_id", "cantidad_solicitada", "estado", "observaciones", "fecha_creacion", "fecha_actualizacion"],
    MOVIMIENTOS: ["id", "comedor_id", "alimento_id", "tipo", "cantidad", "unidad", "referencia_id", "observaciones", "fecha"],
    NUTRICION: ["alimento_id", "calorias_100g", "proteinas_100g", "carbohidratos_100g", "grasas_100g", "fuente", "actualizado"],
    PREPARACIONES: ["id", "comedor_id", "nombre_plato", "fecha", "porciones", "ingredientes_json", "calorias_por_porcion", "proteinas_por_porcion", "carbohidratos_por_porcion", "grasas_por_porcion", "observaciones", "fecha_creacion"]
  }
};

/** Crea las hojas y encabezados sin borrar datos existentes. */
function configurarSistema() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error("Abre este proyecto desde la hoja de cálculo o configura un ID de Spreadsheet.");

  Object.keys(AS.SHEETS).forEach(function(key) {
    const name = nombreHoja_(key);
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);

    const headers = AS.SHEETS[key];
    if (sh.getLastRow() === 0) {
      sh.getRange(1, 1, 1, headers.length).setValues([headers]);
      sh.setFrozenRows(1);
    } else {
      const current = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), headers.length))
        .getValues()[0].filter(String);
      const missing = headers.filter(h => current.indexOf(h) === -1);
      if (missing.length) {
        sh.getRange(1, current.length + 1, 1, missing.length).setValues([missing]);
      }
    }
  });

  return { ok: true, message: "Estructura de Alimentación Solidaria creada/verificada.", version: AS.VERSION };
}

/** Ejecutar manualmente una sola vez para crear el primer administrador. */
function crearAdministradorInicial(correo, contrasena, nombre) {
  configurarSistema();
  correo = normalizarCorreo_(correo);
  if (!correo || !contrasena || String(contrasena).length < 10) {
    throw new Error("Indica un correo y una contraseña de al menos 10 caracteres.");
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const usuarios = leerHoja_("USUARIOS");
    if (usuarios.some(u => normalizarCorreo_(u.correo) === correo)) {
      throw new Error("Ya existe una cuenta con ese correo.");
    }

    const salt = Utilities.getUuid();
    const hash = hashPassword_(contrasena, salt);
    insertar_("USUARIOS", {
      id: nuevoId_("USR"), nombre: nombre || "Administrador", correo: correo,
      password_hash: hash, salt: salt, rol: "ADMIN", estado: "ACTIVO",
      fecha_creacion: new Date()
    });
  } finally {
    lock.releaseLock();
  }
  return { ok: true, message: "Administrador creado. Inicia sesión desde la aplicación." };
}

/** Inicio de sesión. Devuelve un token temporal que el frontend debe conservar. */
function iniciarSesion(correo, contrasena) {
  const email = normalizarCorreo_(correo);
  if (!email || !contrasena) throw new Error("Indica tu correo y contraseña.");

  const cache = CacheService.getScriptCache();
  const claveFallos = "AS_FAIL_" + hashPassword_(email, "intentos");
  const fallos = Number(cache.get(claveFallos) || 0);
  if (fallos >= AS.MAX_INTENTOS_LOGIN) {
    throw new Error("Demasiados intentos fallidos. Espera unos minutos e inténtalo nuevamente.");
  }

  const user = leerHoja_("USUARIOS").find(u =>
    normalizarCorreo_(u.correo) === email && String(u.estado).toUpperCase() === "ACTIVO"
  );
  if (!user || hashPassword_(contrasena, user.salt) !== user.password_hash) {
    cache.put(claveFallos, String(fallos + 1), AS.BLOQUEO_LOGIN_SECONDS);
    throw new Error("Correo o contraseña incorrectos.");
  }
  cache.remove(claveFallos);

  const token = Utilities.getUuid() + Utilities.getUuid();
  const safeUser = {
    id: user.id, nombre: user.nombre, correo: user.correo, rol: String(user.rol).toUpperCase()
  };
  cache.put("AS_TOKEN_" + token, JSON.stringify(safeUser), AS.TOKEN_TTL_SECONDS);
  return { ok: true, token: token, usuario: safeUser, expiraEnSegundos: AS.TOKEN_TTL_SECONDS };
}

function cerrarSesion(token) {
  if (token) CacheService.getScriptCache().remove("AS_TOKEN_" + token);
  return { ok: true };
}

function obtenerSesion(token) {
  return { ok: true, usuario: usuarioDesdeToken_(token) };
}


/** Crear una cuenta de COMEDOR o DONANTE. Solo ADMIN puede hacerlo. */
function crearUsuario(token, datos) {
  usuarioDesdeToken_(token, ["ADMIN"]);
  validarRequeridos_(datos, ["nombre", "correo", "contrasena", "rol"]);
  const rol = String(datos.rol).toUpperCase();
  if (["COMEDOR", "DONANTE"].indexOf(rol) === -1) {
    throw new Error("El rol debe ser COMEDOR o DONANTE. El ADMIN inicial se crea con crearAdministradorInicial().");
  }
  const correo = normalizarCorreo_(datos.correo);
  if (String(datos.contrasena).length < 10) throw new Error("La contraseña debe tener al menos 10 caracteres.");

  const salt = Utilities.getUuid();
  const user = {
    id: nuevoId_("USR"), nombre: limpiar_(datos.nombre), correo: correo,
    password_hash: hashPassword_(datos.contrasena, salt), salt: salt,
    rol: rol, estado: "ACTIVO", fecha_creacion: new Date()
  };

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    if (leerHoja_("USUARIOS").some(u => normalizarCorreo_(u.correo) === correo)) {
      throw new Error("Ya existe una cuenta con ese correo.");
    }
    insertar_("USUARIOS", user);
  } finally {
    lock.releaseLock();
  }
  return { ok: true, id: user.id, nombre: user.nombre, correo: user.correo, rol: user.rol };
}


/* ============================================================
   GESTIÓN DE COMEDORES
   ============================================================ */

/**
 * Registra un comedor y lo vincula con una cuenta de usuario COMEDOR.
 * datos: { usuario_id, nombre, responsable, telefono, direccion,
 *          distrito, beneficiarios }
 */
function crearComedor(token, datos) {
  usuarioDesdeToken_(token, ["ADMIN"]);
  validarRequeridos_(datos, ["usuario_id", "nombre", "responsable", "direccion"]);

  const usuario = exigirRegistro_("USUARIOS", datos.usuario_id);
  if (String(usuario.rol).toUpperCase() !== "COMEDOR") {
    throw new Error("La cuenta vinculada debe tener el rol COMEDOR.");
  }
  if (String(usuario.estado).toUpperCase() !== "ACTIVO") {
    throw new Error("La cuenta del comedor no está activa.");
  }

  const comedor = {
    id: nuevoId_("COM"),
    usuario_id: limpiar_(datos.usuario_id),
    nombre: limpiar_(datos.nombre),
    responsable: limpiar_(datos.responsable),
    telefono: limpiar_(datos.telefono),
    direccion: limpiar_(datos.direccion),
    distrito: limpiar_(datos.distrito),
    beneficiarios: datos.beneficiarios === "" || datos.beneficiarios === undefined || datos.beneficiarios === null
      ? 0 : numeroNoNegativo_(datos.beneficiarios),
    estado: "ACTIVO",
    fecha_creacion: new Date()
  };

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const existente = leerHoja_("COMEDORES").some(c =>
      String(c.usuario_id) === String(datos.usuario_id)
    );
    if (existente) throw new Error("Esta cuenta ya está vinculada a un comedor.");
    insertar_("COMEDORES", comedor);
  } finally {
    lock.releaseLock();
  }
  return { ok: true, id: comedor.id, data: comedor, message: "Comedor registrado correctamente." };
}

/** El administrador puede actualizar los datos de un comedor. */
function actualizarComedor(token, comedorId, cambios) {
  usuarioDesdeToken_(token, ["ADMIN"]);
  if (!cambios || typeof cambios !== "object") throw new Error("No se recibieron cambios.");

  const permitidos = ["nombre", "responsable", "telefono", "direccion", "distrito", "beneficiarios"];
  const limpios = {};
  Object.keys(cambios).forEach(k => {
    if (permitidos.indexOf(k) !== -1) {
      limpios[k] = k === "beneficiarios" ? numeroNoNegativo_(cambios[k]) : limpiar_(cambios[k]);
    }
  });
  if (!Object.keys(limpios).length) throw new Error("No hay campos válidos para actualizar.");

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    actualizar_("COMEDORES", comedorId, limpios);
  } finally {
    lock.releaseLock();
  }
  return { ok: true, id: comedorId, message: "Comedor actualizado correctamente." };
}

/** Lista de comedores. ADMIN ve todos; COMEDOR solo ve su propio registro. */
function listarComedores(token, filtros) {
  const user = usuarioDesdeToken_(token, ["ADMIN", "COMEDOR"]);
  filtros = filtros || {};
  let lista = leerHoja_("COMEDORES");

  if (user.rol === "COMEDOR") {
    lista = lista.filter(c => String(c.usuario_id) === String(user.id));
  } else {
    if (filtros.estado) {
      lista = lista.filter(c => String(c.estado).toUpperCase() === String(filtros.estado).toUpperCase());
    }
    if (filtros.distrito) {
      lista = lista.filter(c => String(c.distrito).toLowerCase().indexOf(String(filtros.distrito).toLowerCase()) !== -1);
    }
  }
  return { ok: true, data: lista };
}

/** Cambia el estado de un comedor a ACTIVO o INACTIVO. */
function cambiarEstadoComedor(token, comedorId, estado) {
  usuarioDesdeToken_(token, ["ADMIN"]);
  estado = String(estado || "").toUpperCase();
  if (["ACTIVO", "INACTIVO"].indexOf(estado) === -1) {
    throw new Error("El estado debe ser ACTIVO o INACTIVO.");
  }
  actualizar_("COMEDORES", comedorId, { estado: estado });
  return { ok: true, id: comedorId, message: "Estado del comedor actualizado." };
}


/* ============================================================
   GESTIÓN DE DONANTES
   ============================================================ */

/**
 * Registra un donante vinculado a una cuenta DONANTE.
 * datos: { usuario_id, nombre, tipo, contacto, telefono, direccion }
 * tipo sugerido: PERSONA, MERCADO, RESTAURANTE, EMPRESA u OTRO.
 */
function crearDonante(token, datos) {
  usuarioDesdeToken_(token, ["ADMIN"]);
  validarRequeridos_(datos, ["usuario_id", "nombre", "tipo"]);

  const usuario = exigirRegistro_("USUARIOS", datos.usuario_id);
  if (String(usuario.rol).toUpperCase() !== "DONANTE") {
    throw new Error("La cuenta vinculada debe tener el rol DONANTE.");
  }
  if (String(usuario.estado).toUpperCase() !== "ACTIVO") {
    throw new Error("La cuenta del donante no está activa.");
  }

  const donante = {
    id: nuevoId_("DON"),
    usuario_id: limpiar_(datos.usuario_id),
    nombre: limpiar_(datos.nombre),
    tipo: limpiar_(datos.tipo).toUpperCase(),
    contacto: limpiar_(datos.contacto),
    telefono: limpiar_(datos.telefono),
    direccion: limpiar_(datos.direccion),
    estado: "ACTIVO",
    fecha_creacion: new Date()
  };

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    if (leerHoja_("DONANTES").some(d => String(d.usuario_id) === String(datos.usuario_id))) {
      throw new Error("Esta cuenta ya está vinculada a un donante.");
    }
    insertar_("DONANTES", donante);
  } finally {
    lock.releaseLock();
  }
  return { ok: true, id: donante.id, data: donante, message: "Donante registrado correctamente." };
}

function actualizarDonante(token, donanteId, cambios) {
  usuarioDesdeToken_(token, ["ADMIN"]);
  if (!cambios || typeof cambios !== "object") throw new Error("No se recibieron cambios.");

  const permitidos = ["nombre", "tipo", "contacto", "telefono", "direccion"];
  const limpios = {};
  Object.keys(cambios).forEach(k => {
    if (permitidos.indexOf(k) !== -1) {
      limpios[k] = k === "tipo" ? limpiar_(cambios[k]).toUpperCase() : limpiar_(cambios[k]);
    }
  });
  if (!Object.keys(limpios).length) throw new Error("No hay campos válidos para actualizar.");

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    actualizar_("DONANTES", donanteId, limpios);
  } finally {
    lock.releaseLock();
  }
  return { ok: true, id: donanteId, message: "Donante actualizado correctamente." };
}

/** ADMIN ve todos; DONANTE solo ve su propio registro. */
function listarDonantes(token, filtros) {
  const user = usuarioDesdeToken_(token, ["ADMIN", "DONANTE"]);
  filtros = filtros || {};
  let lista = leerHoja_("DONANTES");

  if (user.rol === "DONANTE") {
    lista = lista.filter(d => String(d.usuario_id) === String(user.id));
  } else {
    if (filtros.estado) {
      lista = lista.filter(d => String(d.estado).toUpperCase() === String(filtros.estado).toUpperCase());
    }
    if (filtros.tipo) {
      lista = lista.filter(d => String(d.tipo).toUpperCase() === String(filtros.tipo).toUpperCase());
    }
  }
  return { ok: true, data: lista };
}

function cambiarEstadoDonante(token, donanteId, estado) {
  usuarioDesdeToken_(token, ["ADMIN"]);
  estado = String(estado || "").toUpperCase();
  if (["ACTIVO", "INACTIVO"].indexOf(estado) === -1) {
    throw new Error("El estado debe ser ACTIVO o INACTIVO.");
  }
  actualizar_("DONANTES", donanteId, { estado: estado });
  return { ok: true, id: donanteId, message: "Estado del donante actualizado." };
}


/* ============================================================
   CATÁLOGO DE ALIMENTOS
   ============================================================ */

/** ADMIN crea alimentos del catálogo. */
function crearAlimento(token, datos) {
  usuarioDesdeToken_(token, ["ADMIN"]);
  validarRequeridos_(datos, ["nombre", "categoria", "unidad"]);

  const nombre = limpiar_(datos.nombre);
  const categoria = limpiar_(datos.categoria);
  const unidad = limpiar_(datos.unidad).toLowerCase();

  const alimento = {
    id: nuevoId_("ALI"),
    nombre: nombre,
    categoria: categoria,
    unidad: unidad,
    descripcion: limpiar_(datos.descripcion),
    estado: "ACTIVO",
    fecha_creacion: new Date()
  };

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const duplicado = leerHoja_("ALIMENTOS").some(a =>
      String(a.nombre).trim().toLowerCase() === nombre.toLowerCase() &&
      String(a.estado).toUpperCase() !== "INACTIVO"
    );
    if (duplicado) throw new Error("Ya existe un alimento activo con ese nombre.");
    insertar_("ALIMENTOS", alimento);
  } finally {
    lock.releaseLock();
  }
  return { ok: true, id: alimento.id, data: alimento, message: "Alimento creado correctamente." };
}

function actualizarAlimento(token, alimentoId, cambios) {
  usuarioDesdeToken_(token, ["ADMIN"]);
  if (!cambios || typeof cambios !== "object") throw new Error("No se recibieron cambios.");

  const permitidos = ["nombre", "categoria", "unidad", "descripcion"];
  const limpios = {};
  Object.keys(cambios).forEach(k => {
    if (permitidos.indexOf(k) !== -1) {
      limpios[k] = k === "unidad" ? limpiar_(cambios[k]).toLowerCase() : limpiar_(cambios[k]);
    }
  });
  if (!Object.keys(limpios).length) throw new Error("No hay campos válidos para actualizar.");

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    actualizar_("ALIMENTOS", alimentoId, limpios);
  } finally {
    lock.releaseLock();
  }
  return { ok: true, id: alimentoId, message: "Alimento actualizado correctamente." };
}

/** ADMIN, DONANTE y COMEDOR pueden consultar el catálogo activo. */
function listarAlimentos(token, filtros) {
  usuarioDesdeToken_(token, ["ADMIN", "DONANTE", "COMEDOR"]);
  filtros = filtros || {};
  let lista = leerHoja_("ALIMENTOS");

  lista = lista.filter(a => String(a.estado).toUpperCase() === "ACTIVO");
  if (filtros.categoria) {
    lista = lista.filter(a => String(a.categoria).toLowerCase() === String(filtros.categoria).toLowerCase());
  }
  if (filtros.buscar) {
    const q = String(filtros.buscar).toLowerCase();
    lista = lista.filter(a => String(a.nombre).toLowerCase().indexOf(q) !== -1 ||
      String(a.descripcion).toLowerCase().indexOf(q) !== -1);
  }
  return { ok: true, data: lista };
}

function cambiarEstadoAlimento(token, alimentoId, estado) {
  usuarioDesdeToken_(token, ["ADMIN"]);
  estado = String(estado || "").toUpperCase();
  if (["ACTIVO", "INACTIVO"].indexOf(estado) === -1) {
    throw new Error("El estado debe ser ACTIVO o INACTIVO.");
  }
  actualizar_("ALIMENTOS", alimentoId, { estado: estado });
  return { ok: true, id: alimentoId, message: "Estado del alimento actualizado." };
}


/* ============================================================
   PUBLICACIÓN Y GESTIÓN DE DONACIONES
   ============================================================ */

/**
 * El donante publica un alimento disponible.
 * datos: { alimento_id, cantidad, unidad, fecha_vencimiento,
 *          fecha_recojo, direccion_recojo, observaciones }
 */
function crearDonacion(token, datos) {
  const user = usuarioDesdeToken_(token, ["DONANTE", "ADMIN"]);
  validarRequeridos_(datos, ["alimento_id", "cantidad", "unidad", "fecha_vencimiento"]);

  const alimento = exigirRegistro_("ALIMENTOS", datos.alimento_id);
  if (String(alimento.estado).toUpperCase() !== "ACTIVO") {
    throw new Error("El alimento seleccionado no está activo.");
  }

  let donante;
  if (user.rol === "DONANTE") {
    donante = leerHoja_("DONANTES").find(d =>
      String(d.usuario_id) === String(user.id) && String(d.estado).toUpperCase() === "ACTIVO"
    );
    if (!donante) throw new Error("Tu usuario no tiene un perfil de donante activo.");
  } else {
    validarRequeridos_(datos, ["donante_id"]);
    donante = exigirRegistro_("DONANTES", datos.donante_id);
  }
  if (String(donante.estado).toUpperCase() !== "ACTIVO") {
    throw new Error("El perfil del donante está inactivo.");
  }

  const cantidad = numeroPositivo_(datos.cantidad);
  const vencimiento = fechaValida_(datos.fecha_vencimiento);
  const hoy = new Date();
  hoy.setHours(0, 0, 0, 0);
  if (vencimiento < hoy) throw new Error("La fecha de vencimiento no puede estar en el pasado.");

  const fechaRecojo = datos.fecha_recojo ? fechaValida_(datos.fecha_recojo) : "";
  const donacion = {
    id: nuevoId_("DNC"),
    donante_id: donante.id,
    alimento_id: alimento.id,
    cantidad_publicada: cantidad,
    cantidad_disponible: cantidad,
    unidad: limpiar_(datos.unidad).toLowerCase(),
    fecha_vencimiento: vencimiento,
    fecha_recojo: fechaRecojo,
    direccion_recojo: limpiar_(datos.direccion_recojo),
    observaciones: limpiar_(datos.observaciones),
    estado: "DISPONIBLE",
    fecha_creacion: new Date()
  };

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    insertar_("DONACIONES", donacion);
  } finally {
    lock.releaseLock();
  }
  return { ok: true, id: donacion.id, data: donacion, message: "Donación publicada correctamente." };
}

/** Consulta donaciones. ADMIN ve todas; DONANTE ve las propias; COMEDOR ve disponibles y vigentes. */
function listarDonaciones(token, filtros) {
  const user = usuarioDesdeToken_(token, ["ADMIN", "DONANTE", "COMEDOR"]);
  filtros = filtros || {};
  let lista = leerHoja_("DONACIONES");

  if (user.rol === "DONANTE") {
    const perfil = leerHoja_("DONANTES").find(d => String(d.usuario_id) === String(user.id));
    if (!perfil) return { ok: true, data: [] };
    lista = lista.filter(d => String(d.donante_id) === String(perfil.id));
  } else if (user.rol === "COMEDOR") {
    const hoy = new Date();
    hoy.setHours(0, 0, 0, 0);
    lista = lista.filter(d => {
      if (String(d.estado).toUpperCase() !== "DISPONIBLE" || Number(d.cantidad_disponible) <= 0) return false;
      const venc = new Date(d.fecha_vencimiento);
      return isNaN(venc.getTime()) || venc >= hoy;
    });
  }

  if (filtros.estado && user.rol === "ADMIN") {
    lista = lista.filter(d => String(d.estado).toUpperCase() === String(filtros.estado).toUpperCase());
  }
  if (filtros.alimento_id) {
    lista = lista.filter(d => String(d.alimento_id) === String(filtros.alimento_id));
  }
  if (filtros.donante_id && user.rol === "ADMIN") {
    lista = lista.filter(d => String(d.donante_id) === String(filtros.donante_id));
  }
  return { ok: true, data: lista };
}

/** ADMIN puede cambiar estado; DONANTE puede cancelar su propia publicación si aún no fue aprobada/entregada. */
function cambiarEstadoDonacion(token, donacionId, estado) {
  const user = usuarioDesdeToken_(token, ["ADMIN", "DONANTE"]);
  estado = String(estado || "").toUpperCase();
  const estados = ["DISPONIBLE", "AGOTADA", "VENCIDA", "CANCELADA"];
  if (estados.indexOf(estado) === -1) throw new Error("Estado de donación no válido.");

  const donacion = exigirRegistro_("DONACIONES", donacionId);
  if (user.rol === "DONANTE") {
    const perfil = leerHoja_("DONANTES").find(d => String(d.usuario_id) === String(user.id));
    if (!perfil || String(donacion.donante_id) !== String(perfil.id)) {
      throw new Error("No tienes permiso para modificar esta donación.");
    }
    if (estado !== "CANCELADA") throw new Error("El donante solo puede cancelar su publicación.");
    const solicitudes = leerHoja_("SOLICITUDES").filter(s =>
      String(s.donacion_id) === String(donacionId) &&
      ["APROBADA", "ENTREGADA"].indexOf(String(s.estado).toUpperCase()) !== -1
    );
    if (solicitudes.length) throw new Error("No se puede cancelar una donación con solicitudes aprobadas o entregadas.");
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    actualizar_("DONACIONES", donacionId, { estado: estado });
  } finally {
    lock.releaseLock();
  }
  return { ok: true, id: donacionId, message: "Estado de la donación actualizado." };
}


/* ============================================================
   SOLICITUDES DE ALIMENTOS
   ============================================================ */

/**
 * El comedor solicita una cantidad de una donación disponible.
 * datos: { donacion_id, cantidad_solicitada, observaciones }
 */
function crearSolicitud(token, datos) {
  const user = usuarioDesdeToken_(token, ["COMEDOR", "ADMIN"]);
  validarRequeridos_(datos, ["donacion_id", "cantidad_solicitada"]);

  const cantidad = numeroPositivo_(datos.cantidad_solicitada);

  let comedor;
  if (user.rol === "COMEDOR") {
    comedor = leerHoja_("COMEDORES").find(c =>
      String(c.usuario_id) === String(user.id) && String(c.estado).toUpperCase() === "ACTIVO"
    );
    if (!comedor) throw new Error("Tu usuario no tiene un perfil de comedor activo.");
  } else {
    validarRequeridos_(datos, ["comedor_id"]);
    comedor = exigirRegistro_("COMEDORES", datos.comedor_id);
  }
  if (String(comedor.estado).toUpperCase() !== "ACTIVO") {
    throw new Error("El comedor está inactivo.");
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    // Leer la donación dentro del bloqueo para evitar solicitudes que excedan el saldo.
    const actual = exigirRegistro_("DONACIONES", datos.donacion_id);
    if (String(actual.estado).toUpperCase() !== "DISPONIBLE") {
      throw new Error("Esta donación no está disponible.");
    }
    const hoy = new Date();
    hoy.setHours(0, 0, 0, 0);
    const venc = new Date(actual.fecha_vencimiento);
    if (!isNaN(venc.getTime()) && venc < hoy) {
      throw new Error("Esta donación ya venció.");
    }
    if (cantidad > Number(actual.cantidad_disponible)) {
      throw new Error("La cantidad solicitada supera la cantidad disponible.");
    }

    const duplicada = leerHoja_("SOLICITUDES").some(s =>
      String(s.donacion_id) === String(datos.donacion_id) &&
      String(s.comedor_id) === String(comedor.id) &&
      ["PENDIENTE", "APROBADA"].indexOf(String(s.estado).toUpperCase()) !== -1
    );
    if (duplicada) throw new Error("Este comedor ya tiene una solicitud pendiente o aprobada para esta donación.");

    const solicitud = {
      id: nuevoId_("SOL"),
      donacion_id: String(datos.donacion_id),
      comedor_id: comedor.id,
      cantidad_solicitada: cantidad,
      estado: "PENDIENTE",
      observaciones: limpiar_(datos.observaciones),
      fecha_creacion: new Date(),
      fecha_actualizacion: new Date()
    };
    insertar_("SOLICITUDES", solicitud);
    return { ok: true, id: solicitud.id, data: solicitud, message: "Solicitud registrada y pendiente de revisión." };
  } finally {
    lock.releaseLock();
  }
}

/** ADMIN revisa solicitudes. COMEDOR solo consulta las propias. */
function listarSolicitudes(token, filtros) {
  const user = usuarioDesdeToken_(token, ["ADMIN", "COMEDOR"]);
  filtros = filtros || {};
  let lista = leerHoja_("SOLICITUDES");

  if (user.rol === "COMEDOR") {
    const comedor = leerHoja_("COMEDORES").find(c => String(c.usuario_id) === String(user.id));
    if (!comedor) return { ok: true, data: [] };
    lista = lista.filter(s => String(s.comedor_id) === String(comedor.id));
  } else {
    if (filtros.estado) {
      lista = lista.filter(s => String(s.estado).toUpperCase() === String(filtros.estado).toUpperCase());
    }
    if (filtros.comedor_id) {
      lista = lista.filter(s => String(s.comedor_id) === String(filtros.comedor_id));
    }
    if (filtros.donacion_id) {
      lista = lista.filter(s => String(s.donacion_id) === String(filtros.donacion_id));
    }
  }
  return { ok: true, data: lista };
}

/**
 * ADMIN aprueba o rechaza una solicitud.
 * estados permitidos: APROBADA, RECHAZADA.
 * La aprobación reserva la cantidad restándola de cantidad_disponible.
 */
function resolverSolicitud(token, solicitudId, decision, observaciones) {
  usuarioDesdeToken_(token, ["ADMIN"]);
  decision = String(decision || "").toUpperCase();
  if (["APROBADA", "RECHAZADA"].indexOf(decision) === -1) {
    throw new Error("La decisión debe ser APROBADA o RECHAZADA.");
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const solicitud = exigirRegistro_("SOLICITUDES", solicitudId);
    if (String(solicitud.estado).toUpperCase() !== "PENDIENTE") {
      throw new Error("Solo se pueden resolver solicitudes pendientes.");
    }
    const donacion = exigirRegistro_("DONACIONES", solicitud.donacion_id);

    if (decision === "APROBADA") {
      const cantidad = Number(solicitud.cantidad_solicitada);
      const disponible = Number(donacion.cantidad_disponible);
      if (String(donacion.estado).toUpperCase() !== "DISPONIBLE" || cantidad > disponible) {
        throw new Error("La donación ya no tiene cantidad suficiente disponible.");
      }
      const restante = redondear_(disponible - cantidad);
      actualizar_("DONACIONES", donacion.id, {
        cantidad_disponible: restante,
        estado: restante <= 0 ? "AGOTADA" : "DISPONIBLE"
      });
    }

    actualizar_("SOLICITUDES", solicitudId, {
      estado: decision,
      observaciones: observaciones === undefined ? solicitud.observaciones : limpiar_(observaciones),
      fecha_actualizacion: new Date()
    });
    return { ok: true, id: solicitudId, message: "Solicitud " + decision.toLowerCase() + " correctamente." };
  } finally {
    lock.releaseLock();
  }
}

/** Un comedor puede cancelar su propia solicitud si sigue pendiente. */
function cancelarSolicitud(token, solicitudId) {
  const user = usuarioDesdeToken_(token, ["COMEDOR"]);
  const comedor = leerHoja_("COMEDORES").find(c => String(c.usuario_id) === String(user.id));
  if (!comedor) throw new Error("No se encontró el perfil del comedor.");

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const solicitud = exigirRegistro_("SOLICITUDES", solicitudId);
    if (String(solicitud.comedor_id) !== String(comedor.id)) {
      throw new Error("No tienes permiso para cancelar esta solicitud.");
    }
    if (String(solicitud.estado).toUpperCase() !== "PENDIENTE") {
      throw new Error("Solo se pueden cancelar solicitudes pendientes.");
    }
    actualizar_("SOLICITUDES", solicitudId, {
      estado: "CANCELADA",
      fecha_actualizacion: new Date()
    });
    return { ok: true, id: solicitudId, message: "Solicitud cancelada." };
  } finally {
    lock.releaseLock();
  }
}


/* ============================================================
   REGISTRO DE ENTREGAS
   ============================================================ */

/**
 * Registra la entrega de una solicitud APROBADA al comedor.
 * La aprobación de la solicitud ya reservó/restó la cantidad de la donación.
 * Aquí se ingresa esa cantidad al inventario del comedor, sin volver a
 * descontarla de DONACIONES.
 */
function registrarEntrega(token, solicitudId, observaciones) {
  usuarioDesdeToken_(token, ["ADMIN"]);

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const solicitud = exigirRegistro_("SOLICITUDES", solicitudId);
    if (String(solicitud.estado).toUpperCase() !== "APROBADA") {
      throw new Error("Solo se pueden entregar solicitudes aprobadas.");
    }

    const yaEntregada = leerHoja_("MOVIMIENTOS").some(m =>
      String(m.referencia_id) === String(solicitudId) &&
      String(m.tipo).toUpperCase() === "INGRESO_ENTREGA"
    );
    if (yaEntregada) throw new Error("Esta solicitud ya tiene una entrega registrada.");

    const donacion = exigirRegistro_("DONACIONES", solicitud.donacion_id);
    const alimento = exigirRegistro_("ALIMENTOS", donacion.alimento_id);
    const comedor = exigirRegistro_("COMEDORES", solicitud.comedor_id);
    const cantidad = numeroPositivo_(solicitud.cantidad_solicitada);

    const movimiento = {
      id: nuevoId_("MOV"),
      comedor_id: comedor.id,
      alimento_id: alimento.id,
      tipo: "INGRESO_ENTREGA",
      cantidad: cantidad,
      unidad: limpiar_(donacion.unidad || alimento.unidad).toLowerCase(),
      referencia_id: String(solicitudId),
      observaciones: limpiar_(observaciones || "Entrega de solicitud aprobada"),
      fecha: new Date()
    };

    insertar_("MOVIMIENTOS", movimiento);
    actualizar_("SOLICITUDES", solicitudId, {
      estado: "ENTREGADA",
      observaciones: observaciones === undefined ? solicitud.observaciones : limpiar_(observaciones),
      fecha_actualizacion: new Date()
    });

    return {
      ok: true,
      id: movimiento.id,
      data: movimiento,
      message: "Entrega registrada e inventario del comedor actualizado."
    };
  } finally {
    lock.releaseLock();
  }
}


/* ============================================================
   INVENTARIO Y MOVIMIENTOS
   ============================================================ */

/**
 * Calcula inventario a partir del historial MOVIMIENTOS.
 * Los ingresos suman; consumos y salidas restan.
 * No mezcla unidades distintas: convierte g/kg a gramos para consolidar.
 */
function obtenerInventario(token, filtros) {
  const user = usuarioDesdeToken_(token, ["ADMIN", "COMEDOR"]);
  filtros = filtros || {};

  let comedorId = filtros.comedor_id || "";
  if (user.rol === "COMEDOR") {
    const perfil = leerHoja_("COMEDORES").find(c => String(c.usuario_id) === String(user.id));
    if (!perfil) throw new Error("No se encontró el perfil del comedor.");
    comedorId = perfil.id;
  } else if (!comedorId) {
    throw new Error("El administrador debe indicar comedor_id para consultar un inventario.");
  }

  const movimientos = leerHoja_("MOVIMIENTOS").filter(m =>
    String(m.comedor_id) === String(comedorId)
  );
  const alimentos = leerHoja_("ALIMENTOS");
  const acumulado = {};

  movimientos.forEach(m => {
    const alimento = alimentos.find(a => String(a.id) === String(m.alimento_id));
    const nombre = alimento ? alimento.nombre : "Alimento eliminado o no identificado";
    const unidad = String(m.unidad || (alimento ? alimento.unidad : "")).toLowerCase();
    const conversion = convertirAUnidadBase_(Number(m.cantidad), unidad);
    const key = String(m.alimento_id) + "|" + conversion.unidadBase;

    if (!acumulado[key]) {
      acumulado[key] = {
        alimento_id: m.alimento_id,
        alimento: nombre,
        cantidad: 0,
        unidad: conversion.unidadBase,
        comedor_id: comedorId
      };
    }

    const tipo = String(m.tipo).toUpperCase();
    if (["INGRESO_ENTREGA", "INGRESO", "AJUSTE_POSITIVO"].indexOf(tipo) !== -1) {
      acumulado[key].cantidad += conversion.cantidadBase;
    } else if (["CONSUMO_PREPARACION", "SALIDA", "AJUSTE_NEGATIVO"].indexOf(tipo) !== -1) {
      acumulado[key].cantidad -= conversion.cantidadBase;
    }
  });

  let resultado = Object.keys(acumulado).map(k => {
    const item = acumulado[k];
    item.cantidad = redondear_(item.cantidad);
    return item;
  }).filter(x => Math.abs(x.cantidad) > 0.000001);

  if (filtros.alimento_id) {
    resultado = resultado.filter(x => String(x.alimento_id) === String(filtros.alimento_id));
  }
  return { ok: true, data: resultado };
}

/** Devuelve el historial de movimientos. COMEDOR solo ve los propios. */
function listarMovimientos(token, filtros) {
  const user = usuarioDesdeToken_(token, ["ADMIN", "COMEDOR"]);
  filtros = filtros || {};
  let lista = leerHoja_("MOVIMIENTOS");

  if (user.rol === "COMEDOR") {
    const perfil = leerHoja_("COMEDORES").find(c => String(c.usuario_id) === String(user.id));
    if (!perfil) return { ok: true, data: [] };
    lista = lista.filter(m => String(m.comedor_id) === String(perfil.id));
  } else {
    if (filtros.comedor_id) lista = lista.filter(m => String(m.comedor_id) === String(filtros.comedor_id));
  }

  if (filtros.alimento_id) lista = lista.filter(m => String(m.alimento_id) === String(filtros.alimento_id));
  if (filtros.tipo) lista = lista.filter(m => String(m.tipo).toUpperCase() === String(filtros.tipo).toUpperCase());

  lista.sort((a, b) => new Date(b.fecha).getTime() - new Date(a.fecha).getTime());
  return { ok: true, data: lista };
}

/**
 * Permite al ADMIN registrar un ajuste positivo o negativo justificado.
 * datos: { comedor_id, alimento_id, tipo, cantidad, unidad, observaciones }
 */
function registrarAjusteInventario(token, datos) {
  usuarioDesdeToken_(token, ["ADMIN"]);
  validarRequeridos_(datos, ["comedor_id", "alimento_id", "tipo", "cantidad", "unidad", "observaciones"]);

  const comedor = exigirRegistro_("COMEDORES", datos.comedor_id);
  const alimento = exigirRegistro_("ALIMENTOS", datos.alimento_id);
  if (String(comedor.estado).toUpperCase() !== "ACTIVO") throw new Error("El comedor está inactivo.");
  if (String(alimento.estado).toUpperCase() !== "ACTIVO") throw new Error("El alimento está inactivo.");

  const tipo = String(datos.tipo).toUpperCase();
  if (["AJUSTE_POSITIVO", "AJUSTE_NEGATIVO"].indexOf(tipo) === -1) {
    throw new Error("El tipo debe ser AJUSTE_POSITIVO o AJUSTE_NEGATIVO.");
  }
  const cantidad = numeroPositivo_(datos.cantidad);
  const unidad = limpiar_(datos.unidad).toLowerCase();

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    if (tipo === "AJUSTE_NEGATIVO") {
      const stock = calcularStockAlimento_(datos.comedor_id, datos.alimento_id);
      const solicitadaBase = convertirAUnidadBase_(cantidad, unidad);
      if (stock.unidadBase !== solicitadaBase.unidadBase ||
          solicitadaBase.cantidadBase > stock.cantidadBase + 0.000001) {
        throw new Error("El ajuste dejaría el inventario negativo o usa una unidad incompatible.");
      }
    }

    const movimiento = {
      id: nuevoId_("MOV"),
      comedor_id: comedor.id,
      alimento_id: alimento.id,
      tipo: tipo,
      cantidad: cantidad,
      unidad: unidad,
      referencia_id: "",
      observaciones: limpiar_(datos.observaciones),
      fecha: new Date()
    };
    insertar_("MOVIMIENTOS", movimiento);
    return { ok: true, id: movimiento.id, data: movimiento, message: "Ajuste de inventario registrado." };
  } finally {
    lock.releaseLock();
  }
}


/* ============================================================
   INFORMACIÓN NUTRICIONAL
   ============================================================ */

/**
 * Registra o actualiza valores nutricionales por cada 100 gramos.
 * datos: { alimento_id, calorias_100g, proteinas_100g,
 *          carbohidratos_100g, grasas_100g, fuente }
 */
function guardarInformacionNutricional(token, datos) {
  usuarioDesdeToken_(token, ["ADMIN"]);
  validarRequeridos_(datos, [
    "alimento_id", "calorias_100g", "proteinas_100g",
    "carbohidratos_100g", "grasas_100g"
  ]);

  const alimento = exigirRegistro_("ALIMENTOS", datos.alimento_id);
  if (String(alimento.estado).toUpperCase() !== "ACTIVO") {
    throw new Error("No se puede registrar nutrición para un alimento inactivo.");
  }

  const nutricion = {
    alimento_id: String(datos.alimento_id),
    calorias_100g: numeroNoNegativo_(datos.calorias_100g),
    proteinas_100g: numeroNoNegativo_(datos.proteinas_100g),
    carbohidratos_100g: numeroNoNegativo_(datos.carbohidratos_100g),
    grasas_100g: numeroNoNegativo_(datos.grasas_100g),
    fuente: limpiar_(datos.fuente || "Fuente no especificada"),
    actualizado: new Date()
  };

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const existente = leerHoja_("NUTRICION").some(n =>
      String(n.alimento_id) === String(datos.alimento_id)
    );
    if (existente) {
      actualizar_("NUTRICION", datos.alimento_id, nutricion, "alimento_id");
    } else {
      insertar_("NUTRICION", nutricion);
    }
  } finally {
    lock.releaseLock();
  }

  return { ok: true, data: nutricion, message: "Información nutricional guardada por cada 100 gramos." };
}

/** ADMIN, DONANTE y COMEDOR pueden consultar los datos nutricionales. */
function obtenerInformacionNutricional(token, alimentoId) {
  usuarioDesdeToken_(token, ["ADMIN", "DONANTE", "COMEDOR"]);
  const item = leerHoja_("NUTRICION").find(n => String(n.alimento_id) === String(alimentoId));
  if (!item) return { ok: true, data: null, message: "Este alimento aún no tiene información nutricional registrada." };
  return { ok: true, data: item };
}


/* ============================================================
   PREPARACIONES Y CÁLCULO NUTRICIONAL
   ============================================================ */

/**
 * Registra una preparación, descuenta sus ingredientes del inventario
 * y calcula nutrientes por preparación y por porción.
 *
 * datos: {
 *   comedor_id (solo ADMIN; el COMEDOR usa su propio perfil),
 *   nombre_plato, fecha, porciones,
 *   ingredientes: [{ alimento_id, cantidad_gramos }],
 *   observaciones
 * }
 *
 * La cantidad de cada ingrediente siempre se expresa en gramos.
 * El inventario del alimento debe estar en g o kg.
 */
function registrarPreparacion(token, datos) {
  const user = usuarioDesdeToken_(token, ["ADMIN", "COMEDOR"]);
  validarRequeridos_(datos, ["nombre_plato", "porciones", "ingredientes"]);
  if (!Array.isArray(datos.ingredientes) || datos.ingredientes.length === 0) {
    throw new Error("Agrega al menos un ingrediente a la preparación.");
  }

  let comedor;
  if (user.rol === "COMEDOR") {
    comedor = leerHoja_("COMEDORES").find(c =>
      String(c.usuario_id) === String(user.id) && String(c.estado).toUpperCase() === "ACTIVO"
    );
    if (!comedor) throw new Error("Tu usuario no tiene un comedor activo.");
  } else {
    validarRequeridos_(datos, ["comedor_id"]);
    comedor = exigirRegistro_("COMEDORES", datos.comedor_id);
  }

  const porciones = numeroPositivo_(datos.porciones);
  const nutriciones = leerHoja_("NUTRICION");
  const alimentos = leerHoja_("ALIMENTOS");

  // Consolidar ingredientes repetidos y validar antes de modificar hojas.
  const ingredientesMap = {};
  datos.ingredientes.forEach(ing => {
    validarRequeridos_(ing, ["alimento_id", "cantidad_gramos"]);
    const gramos = numeroPositivo_(ing.cantidad_gramos);
    const alimento = alimentos.find(a => String(a.id) === String(ing.alimento_id));
    if (!alimento || String(alimento.estado).toUpperCase() !== "ACTIVO") {
      throw new Error("Uno de los ingredientes no existe o está inactivo: " + ing.alimento_id);
    }
    if (!nutriciones.some(n => String(n.alimento_id) === String(ing.alimento_id))) {
      throw new Error("Falta registrar información nutricional para: " + alimento.nombre);
    }
    if (!ingredientesMap[ing.alimento_id]) {
      ingredientesMap[ing.alimento_id] = {
        alimento_id: String(ing.alimento_id),
        nombre: alimento.nombre,
        cantidad_gramos: 0,
        unidad_inventario: String(alimento.unidad || "").toLowerCase()
      };
    }
    ingredientesMap[ing.alimento_id].cantidad_gramos += gramos;
  });

  const ingredientes = Object.keys(ingredientesMap).map(k => ingredientesMap[k]);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    // Comprobar existencias de todos los ingredientes antes de descontar alguno.
    ingredientes.forEach(ing => {
      const stock = calcularStockAlimento_(comedor.id, ing.alimento_id);
      const requerido = convertirAUnidadBase_(ing.cantidad_gramos, "g");
      if (stock.unidadBase !== requerido.unidadBase ||
          stock.cantidadBase + 0.000001 < requerido.cantidadBase) {
        throw new Error(
          "Inventario insuficiente para " + ing.nombre +
          ". Disponible: " + redondear_(stock.cantidadBase) + " " + stock.unidadBase +
          "; requerido: " + redondear_(requerido.cantidadBase) + " " + requerido.unidadBase + "."
        );
      }
    });

    let calorias = 0, proteinas = 0, carbohidratos = 0, grasas = 0;
    ingredientes.forEach(ing => {
      const n = nutriciones.find(x => String(x.alimento_id) === String(ing.alimento_id));
      const factor = ing.cantidad_gramos / 100;
      calorias += Number(n.calorias_100g) * factor;
      proteinas += Number(n.proteinas_100g) * factor;
      carbohidratos += Number(n.carbohidratos_100g) * factor;
      grasas += Number(n.grasas_100g) * factor;
    });

    const fecha = datos.fecha ? fechaValida_(datos.fecha) : new Date();
    const prepId = nuevoId_("PRE");
    const preparacion = {
      id: prepId,
      comedor_id: comedor.id,
      nombre_plato: limpiar_(datos.nombre_plato),
      fecha: fecha,
      porciones: porciones,
      ingredientes_json: JSON.stringify(ingredientes),
      calorias_por_porcion: redondear_(calorias / porciones),
      proteinas_por_porcion: redondear_(proteinas / porciones),
      carbohidratos_por_porcion: redondear_(carbohidratos / porciones),
      grasas_por_porcion: redondear_(grasas / porciones),
      observaciones: limpiar_(datos.observaciones),
      fecha_creacion: new Date()
    };

    // Primero registrar la preparación y luego movimientos; si falla un movimiento,
    // se informa el error para que el administrador revise la operación.
    insertar_("PREPARACIONES", preparacion);

    ingredientes.forEach(ing => {
      const movimiento = {
        id: nuevoId_("MOV"),
        comedor_id: comedor.id,
        alimento_id: ing.alimento_id,
        tipo: "CONSUMO_PREPARACION",
        cantidad: ing.cantidad_gramos,
        unidad: "g",
        referencia_id: prepId,
        observaciones: "Consumo para preparación: " + preparacion.nombre_plato,
        fecha: new Date()
      };
      insertar_("MOVIMIENTOS", movimiento);
    });

    return {
      ok: true,
      id: prepId,
      data: preparacion,
      resumen_nutricional: {
        calorias_totales: redondear_(calorias),
        proteinas_totales: redondear_(proteinas),
        carbohidratos_totales: redondear_(carbohidratos),
        grasas_totales: redondear_(grasas),
        calorias_por_porcion: preparacion.calorias_por_porcion,
        proteinas_por_porcion: preparacion.proteinas_por_porcion,
        carbohidratos_por_porcion: preparacion.carbohidratos_por_porcion,
        grasas_por_porcion: preparacion.grasas_por_porcion
      },
      message: "Preparación registrada y consumo descontado del inventario."
    };
  } finally {
    lock.releaseLock();
  }
}

/** Lista preparaciones. COMEDOR solo ve las propias. */
function listarPreparaciones(token, filtros) {
  const user = usuarioDesdeToken_(token, ["ADMIN", "COMEDOR"]);
  filtros = filtros || {};
  let lista = leerHoja_("PREPARACIONES");

  if (user.rol === "COMEDOR") {
    const perfil = leerHoja_("COMEDORES").find(c => String(c.usuario_id) === String(user.id));
    if (!perfil) return { ok: true, data: [] };
    lista = lista.filter(p => String(p.comedor_id) === String(perfil.id));
  } else if (filtros.comedor_id) {
    lista = lista.filter(p => String(p.comedor_id) === String(filtros.comedor_id));
  }

  if (filtros.buscar) {
    const q = String(filtros.buscar).toLowerCase();
    lista = lista.filter(p => String(p.nombre_plato).toLowerCase().indexOf(q) !== -1);
  }
  lista.sort((a, b) => new Date(b.fecha).getTime() - new Date(a.fecha).getTime());
  return { ok: true, data: lista };
}


/* ============================================================
   DASHBOARD Y REPORTES
   ============================================================ */

/**
 * Dashboard general para ADMIN. Opcional: filtrar por comedor_id.
 * Los totales de kilos solo incluyen movimientos expresados en g o kg.
 */
function obtenerDashboard(token, filtros) {
  usuarioDesdeToken_(token, ["ADMIN"]);
  filtros = filtros || {};

  const comedorId = filtros.comedor_id || "";
  const movimientos = leerHoja_("MOVIMIENTOS").filter(m =>
    !comedorId || String(m.comedor_id) === String(comedorId)
  );
  const donaciones = leerHoja_("DONACIONES");
  const solicitudes = leerHoja_("SOLICITUDES");
  const comedores = leerHoja_("COMEDORES");
  const donantes = leerHoja_("DONANTES");
  const preparaciones = leerHoja_("PREPARACIONES");

  const ingresosKg = movimientos.reduce((total, m) => {
    if (["INGRESO_ENTREGA", "INGRESO", "AJUSTE_POSITIVO"].indexOf(String(m.tipo).toUpperCase()) === -1) return total;
    const c = convertirAUnidadBase_(Number(m.cantidad), String(m.unidad || "").toLowerCase());
    return c.unidadBase === "g" ? total + c.cantidadBase / 1000 : total;
  }, 0);

  const consumidosKg = movimientos.reduce((total, m) => {
    if (["CONSUMO_PREPARACION", "SALIDA", "AJUSTE_NEGATIVO"].indexOf(String(m.tipo).toUpperCase()) === -1) return total;
    const c = convertirAUnidadBase_(Number(m.cantidad), String(m.unidad || "").toLowerCase());
    return c.unidadBase === "g" ? total + c.cantidadBase / 1000 : total;
  }, 0);

  const hoy = new Date();
  hoy.setHours(0, 0, 0, 0);
  const limite = new Date(hoy);
  limite.setDate(limite.getDate() + 7);
  const proximosVencer = donaciones.filter(d => {
    if (String(d.estado).toUpperCase() !== "DISPONIBLE" || Number(d.cantidad_disponible) <= 0) return false;
    const f = new Date(d.fecha_vencimiento);
    return !isNaN(f.getTime()) && f >= hoy && f <= limite;
  });

  const solicitudesPendientes = solicitudes.filter(s => String(s.estado).toUpperCase() === "PENDIENTE");
  const entregasRealizadas = movimientos.filter(m => String(m.tipo).toUpperCase() === "INGRESO_ENTREGA");
  const donacionesDisponibles = donaciones.filter(d =>
    String(d.estado).toUpperCase() === "DISPONIBLE" && Number(d.cantidad_disponible) > 0
  );

  return {
    ok: true,
    data: {
      total_donaciones: donaciones.length,
      donaciones_disponibles: donacionesDisponibles.length,
      total_comedores: comedores.filter(c => String(c.estado).toUpperCase() === "ACTIVO").length,
      total_donantes: donantes.filter(d => String(d.estado).toUpperCase() === "ACTIVO").length,
      solicitudes_pendientes: solicitudesPendientes.length,
      entregas_realizadas: entregasRealizadas.length,
      alimentos_proximos_a_vencer: proximosVencer.length,
      preparaciones_registradas: preparaciones.length,
      kilos_ingresados_a_comedores: redondear_(ingresosKg),
      kilos_consumidos_en_preparaciones: redondear_(consumidosKg),
      kilos_estimados_disponibles: redondear_(ingresosKg - consumidosKg),
      solicitudes: solicitudesPendientes,
      proximos_vencer: proximosVencer
    }
  };
}


/* ============================================================
   PUNTO DE ENTRADA DE LA APLICACIÓN WEB
   ============================================================ */

function doGet(e) {
  return HtmlService
    .createHtmlOutputFromFile("Index")
    .setTitle("Alimentación Solidaria")
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
}

/** Ejecutar manualmente una sola vez. Cambia la contraseña después del primer ingreso. */
function crearAdminQarana() {
  crearAdministradorInicial(
    "administrador@qarana.com",
    "Qarana2026*",
    "Administrador"
  );
}


/* ============================================================
   FUNCIONES INTERNAS
   ============================================================ */

function nombreHoja_(key) {
  const map = {
    USUARIOS: "USUARIOS", COMEDORES: "COMEDORES", DONANTES: "DONANTES",
    ALIMENTOS: "ALIMENTOS", DONACIONES: "DONACIONES", SOLICITUDES: "SOLICITUDES",
    MOVIMIENTOS: "MOVIMIENTOS", NUTRICION: "NUTRICION",
    PREPARACIONES: "PREPARACIONES"
  };
  return map[key];
}

function hoja_(key) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nombreHoja_(key));
  if (!sh) throw new Error("No existe la hoja " + nombreHoja_(key) + ". Ejecuta configurarSistema().");
  return sh;
}

function leerHoja_(key) {
  const sh = hoja_(key);
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0];
  return values.slice(1).filter(row => row.some(v => v !== "")).map(row => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i] instanceof Date ? row[i].toISOString() : row[i]);
    return obj;
  });
}

/** Evita que textos que empiezan con = + - @ se interpreten como fórmulas. */
function valorSeguro_(v) {
  if (typeof v === "string" && /^[=+\-@]/.test(v)) return "'" + v;
  return v;
}

function insertar_(key, objeto) {
  const sh = hoja_(key);
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const row = headers.map(h => Object.prototype.hasOwnProperty.call(objeto, h) ? valorSeguro_(objeto[h]) : "");
  sh.appendRow(row);
  return { ok: true, id: objeto.id || objeto.alimento_id, data: objeto };
}

function actualizar_(key, id, cambios, campoId) {
  const sh = hoja_(key);
  const values = sh.getDataRange().getValues();
  const headers = values[0];
  const idCol = headers.indexOf(campoId || "id");
  if (idCol < 0) throw new Error("No se encontró la columna identificadora.");
  for (let r = 1; r < values.length; r++) {
    if (String(values[r][idCol]) === String(id)) {
      Object.keys(cambios).forEach(k => {
        const c = headers.indexOf(k);
        if (c >= 0) sh.getRange(r + 1, c + 1).setValue(valorSeguro_(cambios[k]));
      });
      return { ok: true, id: id };
    }
  }
  throw new Error("No se encontró el registro: " + id);
}

function exigirRegistro_(key, id) {
  const item = leerHoja_(key).find(x => String(x.id) === String(id) ||
    (key === "NUTRICION" && String(x.alimento_id) === String(id)));
  if (!item) throw new Error("No existe el registro " + id + " en " + nombreHoja_(key));
  return item;
}

function usuarioDesdeToken_(token, rolesPermitidos) {
  if (!token) throw new Error("Sesión no válida. Inicia sesión nuevamente.");
  const raw = CacheService.getScriptCache().get("AS_TOKEN_" + token);
  if (!raw) throw new Error("La sesión expiró. Inicia sesión nuevamente.");
  const user = JSON.parse(raw);
  if (rolesPermitidos && rolesPermitidos.indexOf(user.rol) === -1) {
    throw new Error("No tienes permisos para realizar esta operación.");
  }
  return user;
}

function hashPassword_(password, salt) {
  const bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256, String(salt) + String(password), Utilities.Charset.UTF_8
  );
  return Utilities.base64Encode(bytes);
}

function nuevoId_(prefijo) {
  return prefijo + "-" + Utilities.getUuid().split("-")[0].toUpperCase();
}

function validarRequeridos_(obj, campos) {
  if (!obj || typeof obj !== "object") throw new Error("Faltan los datos de entrada.");
  campos.forEach(c => {
    if (obj[c] === undefined || obj[c] === null || obj[c] === "") {
      throw new Error("El campo " + c + " es obligatorio.");
    }
  });
}

function limpiar_(v) {
  return v === undefined || v === null ? "" : String(v).trim();
}

function normalizarCorreo_(v) {
  return String(v || "").trim().toLowerCase();
}

function numeroPositivo_(v) {
  const n = Number(v);
  if (!isFinite(n) || n <= 0) throw new Error("La cantidad debe ser un número mayor que cero.");
  return n;
}

function numeroNoNegativo_(v) {
  const n = Number(v);
  if (!isFinite(n) || n < 0) throw new Error("El valor debe ser un número igual o mayor que cero.");
  return n;
}

/** Interpreta fechas AAAA-MM-DD en hora local (evita desfase de un día por UTC). */
function fechaValida_(v) {
  if (!v) return "";
  let d;
  if (v instanceof Date) {
    d = v;
  } else {
    const s = String(v).trim();
    const m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
    d = m ? new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3])) : new Date(s);
  }
  if (isNaN(d.getTime())) throw new Error("La fecha indicada no es válida.");
  return d;
}

function redondear_(n) {
  return Math.round(Number(n) * 100) / 100;
}

/**
 * Convierte cantidades de kg a g. Las unidades no masivas se conservan
 * en su propia unidad y no se mezclan con gramos/kilogramos.
 */
function convertirAUnidadBase_(cantidad, unidad) {
  const u = String(unidad || "").trim().toLowerCase();
  const n = Number(cantidad);
  if (!isFinite(n) || n < 0) throw new Error("Cantidad de movimiento no válida.");

  if (["kg", "kilo", "kilos", "kilogramo", "kilogramos"].indexOf(u) !== -1) {
    return { cantidadBase: n * 1000, unidadBase: "g" };
  }
  if (["g", "gr", "gramo", "gramos"].indexOf(u) !== -1) {
    return { cantidadBase: n, unidadBase: "g" };
  }
  return { cantidadBase: n, unidadBase: u || "unidad" };
}

/** Calcula stock de un alimento para un comedor en unidad base. */
function calcularStockAlimento_(comedorId, alimentoId) {
  const movimientos = leerHoja_("MOVIMIENTOS").filter(m =>
    String(m.comedor_id) === String(comedorId) &&
    String(m.alimento_id) === String(alimentoId)
  );

  let cantidadBase = 0;
  let unidadBase = "";
  movimientos.forEach(m => {
    const convertido = convertirAUnidadBase_(Number(m.cantidad), String(m.unidad || "").toLowerCase());
    if (!unidadBase) unidadBase = convertido.unidadBase;
    if (unidadBase !== convertido.unidadBase) {
      throw new Error("Hay movimientos con unidades incompatibles para el mismo alimento. Revisa el historial.");
    }

    const tipo = String(m.tipo).toUpperCase();
    if (["INGRESO_ENTREGA", "INGRESO", "AJUSTE_POSITIVO"].indexOf(tipo) !== -1) {
      cantidadBase += convertido.cantidadBase;
    } else if (["CONSUMO_PREPARACION", "SALIDA", "AJUSTE_NEGATIVO"].indexOf(tipo) !== -1) {
      cantidadBase -= convertido.cantidadBase;
    }
  });

  return { cantidadBase: redondear_(cantidadBase), unidadBase: unidadBase || "g" };
}