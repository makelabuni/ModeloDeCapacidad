const pptxgen = require("pptxgenjs");

const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.title  = "Modelo de Capacidad";

// ─── Paleta ──────────────────────────────────────────────────────────────────
const C = {
  navy:    "1E2761",
  blue:    "2D5FA3",
  ice:     "CADCFC",
  white:   "FFFFFF",
  gray:    "64748B",
  light:   "F8FAFF",
  bau:     "6B7280",
  norm:    "3B82F6",
  estr:    "8B5CF6",
  evol:    "22C55E",
  libre:   "D1D5DB",
  warn:    "F59E0B",
  danger:  "EF4444",
};

const makeShadow = () => ({ type: "outer", blur: 8, offset: 3, angle: 135, color: "000000", opacity: 0.10 });

// ════════════════════════════════════════════════════════════════════════════
// SLIDE 1 — Portada
// ════════════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.navy };

  // Bloque lateral decorativo
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 0.18, h: 5.625, fill: { color: C.ice } });

  // Título principal
  s.addText("Modelo de Capacidad", {
    x: 0.5, y: 1.4, w: 9, h: 1.2,
    fontSize: 44, fontFace: "Calibri", bold: true, color: C.white,
    margin: 0,
  });

  // Subtítulo
  s.addText("Una herramienta para planificar y entender\nla capacidad real de un equipo de desarrollo", {
    x: 0.5, y: 2.75, w: 8, h: 1.0,
    fontSize: 18, fontFace: "Calibri", color: C.ice,
    margin: 0,
  });

  // Línea separadora
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 3.85, w: 3.5, h: 0.04, fill: { color: C.ice } });

  // Pie
  s.addText("Link · Proyecto ModeloDeCapacidad", {
    x: 0.5, y: 4.1, w: 9, h: 0.4,
    fontSize: 12, fontFace: "Calibri", color: C.ice, margin: 0,
  });
}

// ════════════════════════════════════════════════════════════════════════════
// SLIDE 2 — El problema
// ════════════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.light };

  s.addText("El problema que todos vivimos", {
    x: 0.5, y: 0.35, w: 9, h: 0.7,
    fontSize: 30, fontFace: "Calibri", bold: true, color: C.navy, margin: 0,
  });

  // Tres tarjetas
  const cards = [
    { x: 0.4,  icon: "📦", title: "El software crece", body: "Cada proyecto entregado suma componentes que deben ser mantenidos." },
    { x: 3.55, icon: "🔧", title: "El mantenimiento escala", body: "Bugs, cambios normativos y mejoras consumen capacidad en forma creciente." },
    { x: 6.7,  icon: "📉", title: "La capacidad se erosiona", body: "Cada trimestre, menos FTEs quedan disponibles para generar nuevo valor." },
  ];

  cards.forEach(c => {
    s.addShape(pres.shapes.RECTANGLE, {
      x: c.x, y: 1.2, w: 3.0, h: 3.5,
      fill: { color: C.white }, shadow: makeShadow(),
    });
    s.addText(c.icon, { x: c.x + 0.2, y: 1.35, w: 2.6, h: 0.6, fontSize: 28, margin: 0 });
    s.addText(c.title, {
      x: c.x + 0.2, y: 2.0, w: 2.6, h: 0.5,
      fontSize: 15, fontFace: "Calibri", bold: true, color: C.navy, margin: 0,
    });
    s.addText(c.body, {
      x: c.x + 0.2, y: 2.6, w: 2.6, h: 1.8,
      fontSize: 13, fontFace: "Calibri", color: C.gray, margin: 0,
    });
  });

  // Cita destacada
  s.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 4.85, w: 9.2, h: 0.55, fill: { color: C.ice } });
  s.addText("Sin visibilidad sobre este fenómeno, la planificación de proyectos parte de una premisa incorrecta.", {
    x: 0.6, y: 4.9, w: 8.8, h: 0.45,
    fontSize: 13, fontFace: "Calibri", bold: true, color: C.navy, margin: 0,
  });
}

// ════════════════════════════════════════════════════════════════════════════
// SLIDE 3 — La pregunta clave
// ════════════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.navy };

  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 0.18, h: 5.625, fill: { color: C.ice } });

  s.addText("La pregunta que queremos responder", {
    x: 0.5, y: 0.5, w: 9, h: 0.6,
    fontSize: 22, fontFace: "Calibri", color: C.ice, margin: 0,
  });

  s.addText("¿Cuánta capacidad real tiene\nel equipo para generar valor\nen los próximos trimestres?", {
    x: 0.5, y: 1.3, w: 9, h: 2.4,
    fontSize: 38, fontFace: "Calibri", bold: true, color: C.white, margin: 0,
  });

  s.addText([
    { text: "Considerando el crecimiento del BAU, la calidad del software y los proyectos en cartera.", options: {} }
  ], {
    x: 0.5, y: 3.9, w: 9, h: 0.7,
    fontSize: 16, fontFace: "Calibri", color: C.ice, margin: 0,
  });
}

// ════════════════════════════════════════════════════════════════════════════
// SLIDE 4 — Qué es el BAU
// ════════════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.light };

  s.addText("BAU: el costo oculto del éxito", {
    x: 0.5, y: 0.35, w: 9, h: 0.7,
    fontSize: 30, fontFace: "Calibri", bold: true, color: C.navy, margin: 0,
  });

  // Definición
  s.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 1.15, w: 4.2, h: 3.8, fill: { color: C.white }, shadow: makeShadow() });
  s.addText("Business As Usual (BAU)", {
    x: 0.6, y: 1.3, w: 3.8, h: 0.45,
    fontSize: 15, fontFace: "Calibri", bold: true, color: C.navy, margin: 0,
  });
  s.addText([
    { text: "Es el esfuerzo que el equipo dedica a mantener el software ya en producción:", options: { breakLine: true } },
    { text: " ", options: { breakLine: true } },
    { text: "• Corrección de bugs", options: { bullet: false, breakLine: true } },
    { text: "• Cambios normativos y regulatorios", options: { bullet: false, breakLine: true } },
    { text: "• Soporte y estabilización", options: { bullet: false, breakLine: true } },
    { text: "• Deuda técnica acumulada", options: { bullet: false, breakLine: true } },
    { text: " ", options: { breakLine: true } },
    { text: "No genera valor nuevo, pero es ineludible.", options: { bold: true } },
  ], {
    x: 0.6, y: 1.85, w: 3.8, h: 2.9,
    fontSize: 13, fontFace: "Calibri", color: C.gray, margin: 0,
  });

  // Fórmula
  s.addShape(pres.shapes.RECTANGLE, { x: 5.0, y: 1.15, w: 4.6, h: 1.7, fill: { color: C.navy }, shadow: makeShadow() });
  s.addText("Fórmula del modelo", {
    x: 5.2, y: 1.25, w: 4.2, h: 0.4,
    fontSize: 13, fontFace: "Calibri", bold: true, color: C.ice, margin: 0,
  });
  s.addText("BAU = Componentes × Factor × Modificador_calidad", {
    x: 5.2, y: 1.7, w: 4.2, h: 0.55,
    fontSize: 12, fontFace: "Consolas", color: C.white, margin: 0,
  });
  s.addText("El Factor se calibra con datos históricos reales del equipo.", {
    x: 5.2, y: 2.35, w: 4.2, h: 0.4,
    fontSize: 11, fontFace: "Calibri", color: C.ice, margin: 0,
  });

  // Modificador
  s.addShape(pres.shapes.RECTANGLE, { x: 5.0, y: 3.0, w: 4.6, h: 2.0, fill: { color: C.white }, shadow: makeShadow() });
  s.addText("Modificador de calidad", {
    x: 5.2, y: 3.1, w: 4.2, h: 0.4,
    fontSize: 13, fontFace: "Calibri", bold: true, color: C.navy, margin: 0,
  });
  const factores = [
    { label: "Complejidad ciclomática", dir: "↑ sube BAU" },
    { label: "Deuda técnica",           dir: "↑ sube BAU" },
    { label: "Cobertura de tests",      dir: "↑ baja BAU" },
  ];
  factores.forEach((f, i) => {
    s.addText(`${f.label}`, { x: 5.2, y: 3.6 + i * 0.35, w: 3.0, h: 0.3, fontSize: 12, fontFace: "Calibri", color: C.gray, margin: 0 });
    s.addText(f.dir, { x: 8.3, y: 3.6 + i * 0.35, w: 1.1, h: 0.3, fontSize: 11, fontFace: "Calibri", bold: true, color: i < 2 ? C.danger : C.evol, margin: 0 });
  });
}

// ════════════════════════════════════════════════════════════════════════════
// SLIDE 5 — Cómo funciona el modelo
// ════════════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.light };

  s.addText("Cómo funciona el modelo", {
    x: 0.5, y: 0.35, w: 9, h: 0.7,
    fontSize: 30, fontFace: "Calibri", bold: true, color: C.navy, margin: 0,
  });

  // Tres columnas: Inputs → Modelo → Outputs
  const cols = [
    {
      x: 0.4, title: "Entradas", color: C.blue,
      items: ["Cantidad de Programadores, Testers y Analistas", "Complejidad, deuda técnica, cobertura de tests", "Componentes iniciales + crecimiento trimestral", "Proyectos en cartera (FTEs, duración, componentes entregados)", "BAU % actual conocido del equipo"],
    },
    {
      x: 3.55, title: "Modelo", color: C.navy,
      items: ["Deriva el factor BAU desde datos reales", "Calcula BAU por trimestre a medida que el portfolio crece", "Asigna capacidad por prioridad: Normativo → Estratégico → Evolución", "Detecta déficit y puntos de inflexión"],
    },
    {
      x: 6.7, title: "Salidas", color: C.evol,
      items: ["Distribución trimestral de capacidad en FTE y %", "Gráfico de evolución por vertical", "% BAU con umbrales de alerta (50%) y crítico (70%)", "3 escenarios: Optimista · Base · Pesimista"],
    },
  ];

  cols.forEach(col => {
    s.addShape(pres.shapes.RECTANGLE, { x: col.x, y: 1.1, w: 0.06, h: 3.9, fill: { color: col.color } });
    s.addText(col.title, {
      x: col.x + 0.15, y: 1.1, w: 2.9, h: 0.45,
      fontSize: 16, fontFace: "Calibri", bold: true, color: col.color, margin: 0,
    });
    col.items.forEach((item, i) => {
      s.addText(`· ${item}`, {
        x: col.x + 0.15, y: 1.65 + i * 0.72, w: 2.9, h: 0.68,
        fontSize: 12, fontFace: "Calibri", color: C.gray, margin: 0,
      });
    });
  });

  // Flechas texto
  s.addText("→", { x: 3.35, y: 2.6, w: 0.3, h: 0.4, fontSize: 22, color: C.gray, margin: 0 });
  s.addText("→", { x: 6.5,  y: 2.6, w: 0.3, h: 0.4, fontSize: 22, color: C.gray, margin: 0 });
}

// ════════════════════════════════════════════════════════════════════════════
// SLIDE 6 — Los tres escenarios
// ════════════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.light };

  s.addText("Tres escenarios, una misma realidad", {
    x: 0.5, y: 0.35, w: 9, h: 0.7,
    fontSize: 30, fontFace: "Calibri", bold: true, color: C.navy, margin: 0,
  });

  s.addText("El factor BAU se calibra con datos reales del equipo. Los escenarios modelan cómo varía la calidad del software a futuro.", {
    x: 0.5, y: 1.0, w: 9.2, h: 0.45,
    fontSize: 13, fontFace: "Calibri", color: C.gray, margin: 0,
  });

  const scenarios = [
    { x: 0.4, color: "22C55E", label: "Optimista  ×0.70", emoji: "📈",
      title: "Inversión en calidad",
      body: "El equipo aumenta la cobertura de tests y reduce la deuda técnica activamente. El BAU crece más lentamente a medida que se suman componentes." },
    { x: 3.55, color: "3B82F6", label: "Base  ×1.00", emoji: "📊",
      title: "Calidad estable",
      body: "La calidad del software se mantiene en los niveles actuales. El BAU crece al ritmo calibrado con la experiencia real del equipo." },
    { x: 6.7, color: "EF4444", label: "Pesimista  ×1.40", emoji: "⚠️",
      title: "Degradación de calidad",
      body: "La deuda técnica se acumula y la cobertura baja. El BAU crece más rápido, consumiendo capacidad que podría destinarse a proyectos de valor." },
  ];

  scenarios.forEach(sc => {
    s.addShape(pres.shapes.RECTANGLE, {
      x: sc.x, y: 1.55, w: 3.0, h: 3.8,
      fill: { color: C.white }, shadow: makeShadow(),
    });
    s.addShape(pres.shapes.RECTANGLE, { x: sc.x, y: 1.55, w: 3.0, h: 0.45, fill: { color: sc.color } });
    s.addText(sc.label, {
      x: sc.x + 0.12, y: 1.58, w: 2.8, h: 0.38,
      fontSize: 13, fontFace: "Calibri", bold: true, color: C.white, margin: 0,
    });
    s.addText(sc.emoji, { x: sc.x + 0.12, y: 2.1, w: 0.6, h: 0.5, fontSize: 26, margin: 0 });
    s.addText(sc.title, {
      x: sc.x + 0.12, y: 2.7, w: 2.75, h: 0.4,
      fontSize: 14, fontFace: "Calibri", bold: true, color: C.navy, margin: 0,
    });
    s.addText(sc.body, {
      x: sc.x + 0.12, y: 3.2, w: 2.75, h: 2.0,
      fontSize: 12, fontFace: "Calibri", color: C.gray, margin: 0,
    });
  });
}

// ════════════════════════════════════════════════════════════════════════════
// SLIDE 7 — Ejemplo numérico
// ════════════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.light };

  s.addText("Un ejemplo concreto", {
    x: 0.5, y: 0.35, w: 9, h: 0.7,
    fontSize: 30, fontFace: "Calibri", bold: true, color: C.navy, margin: 0,
  });

  // Parámetros del equipo
  s.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 1.1, w: 4.0, h: 4.2, fill: { color: C.white }, shadow: makeShadow() });
  s.addText("Parámetros del equipo", {
    x: 0.6, y: 1.2, w: 3.6, h: 0.4,
    fontSize: 14, fontFace: "Calibri", bold: true, color: C.navy, margin: 0,
  });

  const params = [
    ["Equipo", "6 prog + 3 testers + 1 analista = 10 FTE"],
    ["Portfolio", "70 componentes en producción"],
    ["Complejidad", "3/5  ·  Deuda técnica: 3/5"],
    ["Cobertura de tests", "70%"],
    ["BAU medido hoy", "40% del equipo"],
    ["Nuevos comp/trim.", "+5 por trimestre"],
  ];
  params.forEach(([k, v], i) => {
    s.addText(k, { x: 0.6, y: 1.75 + i * 0.53, w: 1.6, h: 0.45, fontSize: 12, fontFace: "Calibri", bold: true, color: C.gray, margin: 0 });
    s.addText(v, { x: 2.2, y: 1.75 + i * 0.53, w: 2.0, h: 0.45, fontSize: 12, fontFace: "Calibri", color: C.navy, margin: 0 });
  });

  // Gráfico de barras apiladas nativo
  const labels = ["Hoy", "Q1", "Q2", "Q3", "Q4", "Q5", "Q6", "Q7", "Q8"];

  // BAU crece: factor≈0.038, mod=1.49. BAU_hoy=70*0.038*1.49≈3.96
  // Por trimestre +5 comp: BAU_Q1=75*0.038*1.49≈4.24, etc.
  const bauData   = [3.96, 4.24, 4.53, 4.81, 5.09, 5.37, 5.66, 5.94, 6.22];
  const normData  = [0,    1.5,  1.5,  1.5,  1.0,  1.0,  0,    0,    0   ];
  const estrData  = [0,    1.0,  1.0,  0.5,  0.5,  0,    0,    0,    0   ];
  const evolData  = [0,    0.5,  0.5,  0.5,  0,    0,    0,    0,    0   ];
  const libreData = bauData.map((b, i) => Math.max(0, 10 - b - normData[i] - estrData[i] - evolData[i]).toFixed(2) * 1);

  s.addChart(
    pres.charts.BAR,
    [
      { name: "BAU",         labels, values: bauData   },
      { name: "Normativo",   labels, values: normData  },
      { name: "Estratégico", labels, values: estrData  },
      { name: "Evolución",   labels, values: evolData  },
      { name: "Libre",       labels, values: libreData },
    ],
    {
      x: 4.6, y: 1.1, w: 5.0, h: 4.2,
      barDir: "col",
      barGrouping: "stacked",
      chartColors: [C.bau, C.norm, C.estr, C.evol, C.libre],
      showLegend: true, legendPos: "b", legendFontSize: 10,
      showValue: false,
      chartArea: { fill: { color: C.white } },
      valAxisLabelColor: C.gray,
      catAxisLabelColor: C.gray,
      valGridLine: { color: "E2E8F0", size: 0.5 },
      catGridLine: { style: "none" },
      valAxisMaxVal: 12,
    }
  );
}

// ════════════════════════════════════════════════════════════════════════════
// SLIDE 8 — Qué podemos decidir con esto
// ════════════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.light };

  s.addText("¿Qué decisiones habilita el modelo?", {
    x: 0.5, y: 0.35, w: 9, h: 0.7,
    fontSize: 30, fontFace: "Calibri", bold: true, color: C.navy, margin: 0,
  });

  const decisiones = [
    { emoji: "🗓️", title: "Planificación realista",
      body: "Proyectar la capacidad disponible para nuevos proyectos trimestre a trimestre, sin sobrestimar." },
    { emoji: "⚖️", title: "Priorización con datos",
      body: "Visualizar cómo el BAU compite con los proyectos Normativos, Estratégicos y de Evolución." },
    { emoji: "🔍", title: "Detección temprana de riesgo",
      body: "Identificar el trimestre en que el BAU supera el 50% o 70% antes de que se convierta en un problema." },
    { emoji: "🛠️", title: "Impacto de invertir en calidad",
      body: "Cuantificar cuánta capacidad se recupera si se mejora la cobertura de tests o se reduce la deuda técnica." },
    { emoji: "📋", title: "Defensa de la cartera",
      body: "Argumentar con números por qué ciertos proyectos no caben en un trimestre dado." },
    { emoji: "👥", title: "Dimensionamiento del equipo",
      body: "Evaluar cuántos FTEs adicionales se necesitan para mantener la capacidad de generación de valor." },
  ];

  decisiones.forEach((d, i) => {
    const col = i % 2 === 0 ? 0.4 : 5.1;
    const row = Math.floor(i / 2);
    const y   = 1.15 + row * 1.45;

    s.addShape(pres.shapes.RECTANGLE, { x: col, y, w: 4.5, h: 1.32, fill: { color: C.white }, shadow: makeShadow() });
    s.addText(d.emoji, { x: col + 0.15, y: y + 0.1, w: 0.55, h: 0.55, fontSize: 22, margin: 0 });
    s.addText(d.title, {
      x: col + 0.75, y: y + 0.1, w: 3.55, h: 0.38,
      fontSize: 13, fontFace: "Calibri", bold: true, color: C.navy, margin: 0,
    });
    s.addText(d.body, {
      x: col + 0.75, y: y + 0.52, w: 3.55, h: 0.72,
      fontSize: 11, fontFace: "Calibri", color: C.gray, margin: 0,
    });
  });
}

// ════════════════════════════════════════════════════════════════════════════
// SLIDE 9 — Próximos pasos
// ════════════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.light };

  s.addText("Próximos pasos", {
    x: 0.5, y: 0.35, w: 9, h: 0.7,
    fontSize: 30, fontFace: "Calibri", bold: true, color: C.navy, margin: 0,
  });

  const pasos = [
    { n: "01", title: "Calibrar con datos reales",
      body: "Cada equipo ingresa su BAU % medido hoy. El modelo deriva el factor automáticamente." },
    { n: "02", title: "Cargar la cartera de proyectos",
      body: "Ingresar los proyectos planificados con FTEs, trimestre de inicio, duración y componentes a entregar." },
    { n: "03", title: "Analizar escenarios",
      body: "Comparar Base vs. Optimista para cuantificar el beneficio de invertir en calidad." },
    { n: "04", title: "Revisar en cada ciclo de planificación",
      body: "Actualizar los parámetros y la cartera al inicio de cada trimestre para mantener el modelo vivo." },
  ];

  pasos.forEach((p, i) => {
    const y = 1.2 + i * 1.05;
    s.addShape(pres.shapes.RECTANGLE, { x: 0.4, y, w: 0.65, h: 0.65, fill: { color: C.navy } });
    s.addText(p.n, {
      x: 0.4, y, w: 0.65, h: 0.65,
      fontSize: 18, fontFace: "Calibri", bold: true, color: C.white, align: "center", valign: "middle", margin: 0,
    });
    s.addText(p.title, {
      x: 1.2, y: y + 0.03, w: 8.4, h: 0.35,
      fontSize: 15, fontFace: "Calibri", bold: true, color: C.navy, margin: 0,
    });
    s.addText(p.body, {
      x: 1.2, y: y + 0.42, w: 8.4, h: 0.55,
      fontSize: 13, fontFace: "Calibri", color: C.gray, margin: 0,
    });
  });
}

// ════════════════════════════════════════════════════════════════════════════
// SLIDE 10 — Cierre
// ════════════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.navy };

  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 0.18, h: 5.625, fill: { color: C.ice } });

  s.addText("El modelo no predice el futuro.", {
    x: 0.5, y: 1.2, w: 9, h: 0.7,
    fontSize: 28, fontFace: "Calibri", color: C.ice, margin: 0,
  });

  s.addText("Hace visible lo que ya está pasando.", {
    x: 0.5, y: 2.0, w: 9, h: 0.9,
    fontSize: 40, fontFace: "Calibri", bold: true, color: C.white, margin: 0,
  });

  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 3.1, w: 3.5, h: 0.04, fill: { color: C.ice } });

  s.addText("Modelo de Capacidad · Link · makelabuni", {
    x: 0.5, y: 3.3, w: 9, h: 0.4,
    fontSize: 12, fontFace: "Calibri", color: C.ice, margin: 0,
  });
}

// ─── Exportar ────────────────────────────────────────────────────────────────
pres.writeFile({ fileName: "D:/ClaudeCode/ModeloDeCapacidad/ModeloDeCapacidad.pptx" })
  .then(() => console.log("✅ Presentación generada: ModeloDeCapacidad.pptx"))
  .catch(e => console.error("❌ Error:", e));
