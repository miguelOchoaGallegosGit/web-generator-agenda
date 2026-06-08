import {
  AlignmentType,
  BorderStyle,
  Document,
  HeightRule,
  PageOrientation,
  Packer,
  Paragraph,
  Table,
  TableCell,
  TableLayoutType,
  TableRow,
  TextRun,
  VerticalAlign,
  WidthType,
} from "docx";

export type AgendaModelId = "clasica" | "color-pop" | "semana-vista";

export type AgendaModel = {
  id: AgendaModelId;
  label: string;
  description: string;
  image: string;
};

export const agendaModels: AgendaModel[] = [
  {
    id: "clasica",
    label: "Clásica sobria",
    description: "Dos días por página, líneas amplias y estilo limpio.",
    image: "/modelos/modelo-1.jpeg",
  },
  {
    id: "color-pop",
    label: "Color pop",
    description: "Bloques alegres de cuatro días para organizar rápido.",
    image: "/modelos/modelo-2.jpeg",
  },
  {
    id: "semana-vista",
    label: "Semana vista",
    description: "Una semana por página con acentos de color.",
    image: "/modelos/modelo-3.jpeg",
  },
];

type AgendaDay = {
  date: Date;
  dayNumber: string;
  weekday: string;
  month: string;
  year: number;
};

const monthFormatter = new Intl.DateTimeFormat("es-PE", { month: "long" });
const weekdayFormatter = new Intl.DateTimeFormat("es-PE", { weekday: "long" });
const longDateFormatter = new Intl.DateTimeFormat("es-PE", {
  day: "2-digit",
  month: "long",
  year: "numeric",
});

const palette = ["F5C84C", "52C7B8", "5EA8F2", "F06292", "8D6DF2", "F59F68", "E25858"];
const A4_PORTRAIT = { width: 11906, height: 16838 };
const A4_MARGIN = { top: 280, right: 280, bottom: 280, left: 280 };
const DAYS_BY_MODEL: Record<AgendaModelId, number> = {
  clasica: 2,
  "color-pop": 4,
  "semana-vista": 7,
};

const emptyBorders = {
  top: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
  bottom: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
  left: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
  right: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
};

const softBorders = {
  top: { style: BorderStyle.SINGLE, size: 4, color: "D9DEE7" },
  bottom: { style: BorderStyle.SINGLE, size: 4, color: "D9DEE7" },
  left: { style: BorderStyle.SINGLE, size: 4, color: "D9DEE7" },
  right: { style: BorderStyle.SINGLE, size: 4, color: "D9DEE7" },
};

const lineBorders = {
  top: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
  bottom: { style: BorderStyle.SINGLE, size: 3, color: "D7DEE8" },
  left: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
  right: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
};

export function parseInputDate(value: string): Date {
  const [year, month, day] = value.split("-").map(Number);
  return new Date(year, month - 1, day);
}

export function formatInputDate(date: Date): string {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

export function diffInDays(start: Date, end: Date): number {
  const cleanStart = new Date(start.getFullYear(), start.getMonth(), start.getDate());
  const cleanEnd = new Date(end.getFullYear(), end.getMonth(), end.getDate());
  return Math.round((cleanEnd.getTime() - cleanStart.getTime()) / 86_400_000);
}

export function validateDateRange(startValue: string, endValue: string): string | null {
  if (!startValue || !endValue) {
    return "Ingresa una fecha de inicio y una fecha de fin.";
  }

  const start = parseInputDate(startValue);
  const end = parseInputDate(endValue);
  const days = diffInDays(start, end);

  if (Number.isNaN(start.getTime()) || Number.isNaN(end.getTime())) {
    return "Revisa el formato de las fechas.";
  }

  if (days < 0) {
    return "La fecha fin no puede ser anterior a la fecha de inicio.";
  }

  if (days > 365) {
    return "El rango máximo permitido es de un año.";
  }

  return null;
}

export async function generateAgendaDocx(
  startValue: string,
  endValue: string,
  modelId: AgendaModelId,
): Promise<Blob> {
  const start = parseInputDate(startValue);
  const end = parseInputDate(endValue);
  const days = buildDays(start, end);
  const selectedModel = agendaModels.find((model) => model.id === modelId) ?? agendaModels[0];

  const doc = new Document({
    creator: "Generador de Agendas",
    title: `Agenda ${selectedModel.label}`,
    description: "Agenda imprimible generada desde la app web.",
    styles: {
      default: {
        document: {
          run: {
            font: modelId === "color-pop" ? "Comic Sans MS" : "Aptos",
            color: "1F2937",
          },
          paragraph: {
            spacing: { after: 45 },
          },
        },
      },
    },
    sections: buildSections(days, modelId, selectedModel),
  });

  return Packer.toBlob(doc);
}

function buildDays(start: Date, end: Date): AgendaDay[] {
  const result: AgendaDay[] = [];
  const cursor = new Date(start);

  while (cursor <= end) {
    result.push(buildDay(cursor));
    cursor.setDate(cursor.getDate() + 1);
  }

  return result;
}

function buildSections(days: AgendaDay[], modelId: AgendaModelId, selectedModel: AgendaModel) {
  const pageGroups = buildPageGroups(days, modelId);

  // Build individual A5 page content
  const a5Pages: Table[] = pageGroups.map((group, index) => {
    if (modelId === "clasica") return classicPage(group);
    if (modelId === "color-pop") return colorPopPage(group);
    return weekPage(group, selectedModel, index + 1, pageGroups.length);
  });

  // Pair A5 pages into diptych spreads (2 per A4 landscape sheet)
  const spreads: [Table, Table | null][] = [];
  for (let i = 0; i < a5Pages.length; i += 2) {
    spreads.push([a5Pages[i], a5Pages[i + 1] ?? null]);
  }

  return spreads.map(([left, right]) => ({
    properties: {
      page: {
        size: {
          ...A4_PORTRAIT,
          orientation: PageOrientation.LANDSCAPE,
        },
        margin: A4_MARGIN,
      },
    },
    children: [diptychTable(left, right)],
  }));
}

function diptychTable(left: Table, right: Table | null): Table {
  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    layout: TableLayoutType.FIXED,
    borders: emptyBorders,
    rows: [
      new TableRow({
        children: [
          new TableCell({
            width: { size: 50, type: WidthType.PERCENTAGE },
            verticalAlign: VerticalAlign.TOP,
            margins: { top: 100, right: 250, bottom: 100, left: 100 },
            borders: {
              top: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
              bottom: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
              left: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
              right: { style: BorderStyle.DOTTED, size: 2, color: "D0D5DD" },
            },
            children: [left],
          }),
          new TableCell({
            width: { size: 50, type: WidthType.PERCENTAGE },
            verticalAlign: VerticalAlign.TOP,
            margins: { top: 100, right: 100, bottom: 100, left: 250 },
            borders: emptyBorders,
            children: right ? [right] : [new Paragraph("")],
          }),
        ],
      }),
    ],
  });
}

function classicPage(days: AgendaDay[]): Table {
  const cells = days.map((day) => dayCellClassic(day));
  while (cells.length < 2) {
    cells.push(blankCell());
  }

  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    layout: TableLayoutType.FIXED,
    borders: emptyBorders,
    rows: [
      new TableRow({
        children: cells,
      }),
    ],
  });
}

function colorPopPage(days: AgendaDay[]): Table {
  const cells = days.map((day, index) => dayCellColor(day, palette[index % palette.length]));
  while (cells.length < 4) {
    cells.push(blankCell());
  }
  const firstDay = days[0];
  const titleMonth = firstDay ? firstDay.month.toUpperCase() : "";
  const titleYear = firstDay ? String(firstDay.year) : "";

  const colorfulText = (text: string, size: number) =>
    text.split("").map((char, i) => new TextRun({ text: char, bold: true, size, color: palette[i % palette.length] }));

  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    layout: TableLayoutType.FIXED,
    borders: emptyBorders,
    rows: [
      new TableRow({
        cantSplit: true,
        children: [
          new TableCell({
            width: { size: 46, type: WidthType.PERCENTAGE },
            borders: emptyBorders,
            margins: { top: 0, right: 70, bottom: 50, left: 70 },
            children: [
              new Paragraph({
                children: colorfulText(titleMonth, 34),
              }),
            ],
          }),
          gutterCell(),
          new TableCell({
            width: { size: 46, type: WidthType.PERCENTAGE },
            borders: emptyBorders,
            margins: { top: 0, right: 70, bottom: 50, left: 70 },
            children: [
              new Paragraph({
                alignment: AlignmentType.RIGHT,
                children: colorfulText(titleYear, 34),
              }),
            ],
          }),
        ],
      }),
      new TableRow({ cantSplit: true, children: [cells[0], gutterCell(), cells[1]] }),
      new TableRow({
        height: { value: 140, rule: HeightRule.EXACT },
        children: [blankCell(), gutterCell(), blankCell()],
      }),
      new TableRow({ cantSplit: true, children: [cells[2], gutterCell(), cells[3]] }),
    ],
  });
}

function weekPage(days: AgendaDay[], model: AgendaModel, page: number, total: number): Table {
  const cells = days.map((day, index) => dayCellWeek(day, palette[index % palette.length]));
  while (cells.length < 7) {
    cells.push(blankCell());
  }

  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    layout: TableLayoutType.FIXED,
    borders: emptyBorders,
    rows: [
      new TableRow({
        cantSplit: true,
        children: [
          new TableCell({
            columnSpan: 2,
            borders: emptyBorders,
            margins: { top: 0, right: 70, bottom: 40, left: 70 },
            children: [modelHeader(model, days, page, total)],
          }),
        ],
      }),
      new TableRow({ children: [cells[0], cells[1]] }),
      new TableRow({ children: [cells[2], cells[3]] }),
      new TableRow({ children: [cells[4], cells[5]] }),
      new TableRow({ children: [cells[6], monthSummary(days)] }),
    ],
  });
}

function dayCellClassic(day: AgendaDay): TableCell {
  return new TableCell({
    width: { size: 50, type: WidthType.PERCENTAGE },
    verticalAlign: VerticalAlign.TOP,
    margins: { top: 160, right: 220, bottom: 180, left: 220 },
    borders: emptyBorders,
    children: [
      new Paragraph({
        alignment: AlignmentType.RIGHT,
        children: [new TextRun({ text: `${day.month} ${day.year}`, bold: true, size: 22 })],
      }),
      new Paragraph({
        children: [
          new TextRun({ text: day.dayNumber, bold: true, size: 30 }),
          new TextRun({ text: `  ${day.weekday}`, bold: true, size: 22 }),
        ],
      }),
      ...lines(12, "9CA3AF", 18),
    ],
  });
}

function dayCellColor(day: AgendaDay, color: string): TableCell {
  return new TableCell({
    width: { size: 46, type: WidthType.PERCENTAGE },
    verticalAlign: VerticalAlign.TOP,
    margins: { top: 80, right: 80, bottom: 90, left: 80 },
    borders: {
      top: { style: BorderStyle.SINGLE, size: 6, color },
      bottom: { style: BorderStyle.SINGLE, size: 6, color },
      left: { style: BorderStyle.SINGLE, size: 6, color },
      right: { style: BorderStyle.SINGLE, size: 6, color },
    },
    children: [
      new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        layout: TableLayoutType.FIXED,
        borders: emptyBorders,
        rows: [
          dayRibbonRow(day, color),
          ...ruledLineRows(7),
        ],
      }),
    ],
  });
}

function dayCellWeek(day: AgendaDay, color: string): TableCell {
  return new TableCell({
    width: { size: 50, type: WidthType.PERCENTAGE },
    verticalAlign: VerticalAlign.TOP,
    margins: { top: 120, right: 150, bottom: 140, left: 150 },
    borders: emptyBorders,
    children: [
      new Paragraph({
        children: [
          new TextRun({ text: day.dayNumber, bold: true, color, size: 28 }),
          new TextRun({ text: `  ${day.weekday.toLowerCase()}`, bold: true, color, size: 18 }),
          new TextRun({ text: `  ${day.month}`, size: 16, color: "6B7280" }),
        ],
      }),
      ...lines(day.date.getDay() === 0 || day.date.getDay() === 6 ? 3 : 4, "111827", 18),
    ],
  });
}

function monthSummary(days: AgendaDay[]): TableCell {
  const first = days[0];
  const month = first ? `${first.month} ${first.year}` : "";

  return new TableCell({
    width: { size: 50, type: WidthType.PERCENTAGE },
    verticalAlign: VerticalAlign.CENTER,
    margins: { top: 100, right: 120, bottom: 100, left: 120 },
    borders: softBorders,
    children: [
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 60 },
        children: [new TextRun({ text: month, bold: true, size: 22, color: "111827" })],
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 30 },
        children: [new TextRun({ text: "Lu   Ma   Mi   Ju   Vi   Sa   Do", bold: true, size: 15, color: "0F766E", font: "Consolas" })],
      }),
      ...miniCalendar(first?.date ?? new Date()),
    ],
  });
}

function blankCell(): TableCell {
  return new TableCell({
    width: { size: 50, type: WidthType.PERCENTAGE },
    borders: emptyBorders,
    children: [new Paragraph("")],
  });
}

function gutterCell(): TableCell {
  return new TableCell({
    width: { size: 8, type: WidthType.PERCENTAGE },
    borders: emptyBorders,
    children: [new Paragraph("")],
  });
}

function modelHeader(model: AgendaModel, days: AgendaDay[], page: number, total: number): Table {
  const firstDay = days[0];
  const titleMonth = firstDay ? firstDay.month.toUpperCase() : "";
  const titleYear = firstDay ? String(firstDay.year) : "";

  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    layout: TableLayoutType.FIXED,
    borders: emptyBorders,
    rows: [
      new TableRow({
        cantSplit: true,
        children: [
          new TableCell({
            width: { size: 50, type: WidthType.PERCENTAGE },
            borders: emptyBorders,
            margins: { top: 0, right: 70, bottom: 50, left: 70 },
            children: [
              new Paragraph({
                children: [new TextRun({ text: titleMonth, bold: true, size: 36, color: "EA5EC7" })],
              }),
            ],
          }),
          new TableCell({
            width: { size: 50, type: WidthType.PERCENTAGE },
            borders: emptyBorders,
            margins: { top: 0, right: 70, bottom: 50, left: 70 },
            children: [
              new Paragraph({
                alignment: AlignmentType.RIGHT,
                children: [new TextRun({ text: titleYear, bold: true, size: 36, color: "C46CE8" })],
              }),
            ],
          }),
        ],
      }),
    ],
  });
}

function dayRibbonRow(day: AgendaDay, color: string): TableRow {
  return new TableRow({
    cantSplit: true,
    children: [
      new TableCell({
        width: { size: 18, type: WidthType.PERCENTAGE },
        shading: { fill: color },
        borders: emptyBorders,
        margins: { top: 25, right: 25, bottom: 25, left: 25 },
        children: [
          new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [new TextRun({ text: day.dayNumber, bold: true, color: "FFFFFF", size: 20 })],
          }),
        ],
      }),
      new TableCell({
        width: { size: 82, type: WidthType.PERCENTAGE },
        shading: { fill: color },
        borders: emptyBorders,
        margins: { top: 25, right: 45, bottom: 25, left: 45 },
        children: [
          new Paragraph({
            alignment: AlignmentType.RIGHT,
            children: [
              new TextRun({ text: day.weekday.toUpperCase(), bold: true, color: "FFFFFF", size: 20 }),
            ],
          }),
        ],
      }),
    ],
  });
}

function ruledLineRows(count: number): TableRow[] {
  return Array.from(
    { length: count },
    () =>
      new TableRow({
        height: { value: 210, rule: HeightRule.EXACT },
        children: [
          new TableCell({
            columnSpan: 2,
            borders: lineBorders,
            margins: { top: 0, right: 0, bottom: 0, left: 0 },
            children: [new Paragraph("")],
          }),
        ],
      }),
  );
}

function buildPageGroups(days: AgendaDay[], modelId: AgendaModelId): AgendaDay[][] {
  if (modelId === "semana-vista") {
    return groupByWeeks(days);
  }

  const perPage = DAYS_BY_MODEL[modelId];
  const groups: AgendaDay[][] = [];

  for (let index = 0; index < days.length; index += perPage) {
    const pageDays = days.slice(index, index + perPage);
    groups.push(fillMissingDays(pageDays, perPage));
  }

  return groups;
}

function fillMissingDays(days: AgendaDay[], count: number): AgendaDay[] {
  if (days.length === 0) {
    return days;
  }

  const result = [...days];
  const cursor = new Date(days[days.length - 1].date);

  while (result.length < count) {
    cursor.setDate(cursor.getDate() + 1);
    result.push(buildDay(cursor));
  }

  return result;
}

function buildDay(date: Date): AgendaDay {
  return {
    date: new Date(date),
    dayNumber: String(date.getDate()),
    weekday: capitalize(weekdayFormatter.format(date)),
    month: capitalize(monthFormatter.format(date)),
    year: date.getFullYear(),
  };
}

function lines(count: number, color: string, size: number): Paragraph[] {
  return Array.from({ length: count }, () =>
    new Paragraph({
      spacing: { before: 70, after: 70 },
      children: [new TextRun({ text: "______________________________________", color, size })],
    }),
  );
}

function miniCalendar(date: Date): Paragraph[] {
  const year = date.getFullYear();
  const month = date.getMonth();
  const first = new Date(year, month, 1);
  const last = new Date(year, month + 1, 0);
  const offset = (first.getDay() + 6) % 7;
  const slots = Array.from({ length: offset }, () => "  ");

  for (let day = 1; day <= last.getDate(); day += 1) {
    slots.push(String(day).padStart(2, " "));
  }

  while (slots.length % 7 !== 0) {
    slots.push("  ");
  }

  return chunk(slots, 7).map(
    (week) =>
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 20, after: 20 },
        children: [new TextRun({ text: week.join("  "), size: 16, color: "374151", font: "Consolas" })],
      }),
  );
}

function groupByWeeks(days: AgendaDay[]): AgendaDay[][] {
  const groups: AgendaDay[][] = [];
  let current: AgendaDay[] = [];

  days.forEach((day) => {
    const weekDay = (day.date.getDay() + 6) % 7;

    if (weekDay === 0 && current.length > 0) {
      groups.push(current);
      current = [];
    }

    current.push(day);

    if (current.length === 7) {
      groups.push(current);
      current = [];
    }
  });

  if (current.length > 0) {
    groups.push(current);
  }

  return groups;
}

function chunk<T>(items: T[], size: number): T[][] {
  const result: T[][] = [];

  for (let index = 0; index < items.length; index += size) {
    result.push(items.slice(index, index + size));
  }

  return result;
}

function capitalize(value: string): string {
  return value.charAt(0).toUpperCase() + value.slice(1);
}
