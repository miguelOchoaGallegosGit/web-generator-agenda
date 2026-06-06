import {
  AlignmentType,
  BorderStyle,
  Document,
  HeadingLevel,
  Packer,
  Paragraph,
  Table,
  TableCell,
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

const emptyBorders = {
  top: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
  bottom: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
  left: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
  right: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
};

const softBorders = {
  top: { style: BorderStyle.SINGLE, size: 2, color: "D9DEE7" },
  bottom: { style: BorderStyle.SINGLE, size: 2, color: "D9DEE7" },
  left: { style: BorderStyle.SINGLE, size: 2, color: "D9DEE7" },
  right: { style: BorderStyle.SINGLE, size: 2, color: "D9DEE7" },
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
            font: "Aptos",
            color: "1F2937",
          },
          paragraph: {
            spacing: { after: 90 },
          },
        },
      },
      paragraphStyles: [
        {
          id: "Heading1",
          name: "Heading 1",
          basedOn: "Normal",
          next: "Normal",
          quickFormat: true,
          run: {
            size: 34,
            bold: true,
            color: "111827",
          },
          paragraph: {
            spacing: { after: 220 },
          },
        },
      ],
    },
    sections: buildSections(days, modelId, selectedModel),
  });

  return Packer.toBlob(doc);
}

function buildDays(start: Date, end: Date): AgendaDay[] {
  const result: AgendaDay[] = [];
  const cursor = new Date(start);

  while (cursor <= end) {
    result.push({
      date: new Date(cursor),
      dayNumber: String(cursor.getDate()),
      weekday: capitalize(weekdayFormatter.format(cursor)),
      month: capitalize(monthFormatter.format(cursor)),
      year: cursor.getFullYear(),
    });
    cursor.setDate(cursor.getDate() + 1);
  }

  return result;
}

function buildSections(days: AgendaDay[], modelId: AgendaModelId, selectedModel: AgendaModel) {
  const pageGroups =
    modelId === "semana-vista"
      ? groupByWeeks(days)
      : chunk(days, modelId === "clasica" ? 2 : 4);

  return pageGroups.map((group, index) => ({
    properties: {
      page: {
        margin: { top: 560, right: 520, bottom: 520, left: 520 },
      },
    },
    children: [
      coverHeader(selectedModel, group, index + 1, pageGroups.length),
      modelId === "clasica"
        ? classicPage(group)
        : modelId === "color-pop"
          ? colorPopPage(group)
          : weekPage(group),
    ],
  }));
}

function coverHeader(model: AgendaModel, days: AgendaDay[], page: number, total: number): Paragraph {
  const firstDay = days[0];
  const lastDay = days[days.length - 1] ?? firstDay;
  const range =
    firstDay && lastDay
      ? `${longDateFormatter.format(firstDay.date)} - ${longDateFormatter.format(lastDay.date)}`
      : "";

  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    alignment: AlignmentType.CENTER,
    children: [
      new TextRun({ text: "Agenda ", bold: true }),
      new TextRun({ text: model.label, color: "0F766E", bold: true }),
      new TextRun({ text: `  |  ${range}`, size: 22, color: "6B7280" }),
      new TextRun({ text: `  |  ${page}/${total}`, size: 18, color: "9CA3AF" }),
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

  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: emptyBorders,
    rows: [
      new TableRow({ children: [cells[0], cells[1]] }),
      new TableRow({ children: [cells[2], cells[3]] }),
    ],
  });
}

function weekPage(days: AgendaDay[]): Table {
  const cells = days.map((day, index) => dayCellWeek(day, palette[index % palette.length]));
  while (cells.length < 7) {
    cells.push(blankCell());
  }

  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: emptyBorders,
    rows: [
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
        children: [new TextRun({ text: `${day.month} ${day.year}`, bold: true, size: 26 })],
      }),
      new Paragraph({
        children: [
          new TextRun({ text: day.dayNumber, bold: true, size: 34 }),
          new TextRun({ text: `  ${day.weekday}`, bold: true, size: 24 }),
        ],
      }),
      ...lines(14, "9CA3AF", 22),
    ],
  });
}

function dayCellColor(day: AgendaDay, color: string): TableCell {
  return new TableCell({
    width: { size: 50, type: WidthType.PERCENTAGE },
    verticalAlign: VerticalAlign.TOP,
    margins: { top: 150, right: 150, bottom: 180, left: 150 },
    borders: softBorders,
    children: [
      new Paragraph({
        shading: { fill: color },
        spacing: { after: 120 },
        children: [
          new TextRun({ text: day.weekday.toUpperCase(), bold: true, color: "FFFFFF", size: 21 }),
          new TextRun({ text: `   ${day.dayNumber}`, bold: true, color: "FFFFFF", size: 22 }),
        ],
      }),
      new Paragraph({
        children: [new TextRun({ text: `${day.month} ${day.year}`, bold: true, color, size: 18 })],
      }),
      ...lines(9, "CBD5E1", 19),
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
          new TextRun({ text: day.dayNumber, bold: true, color, size: 32 }),
          new TextRun({ text: `  ${day.weekday.toLowerCase()}`, bold: true, color, size: 20 }),
          new TextRun({ text: `  ${day.month}`, size: 16, color: "6B7280" }),
        ],
      }),
      ...lines(day.weekday === "Sábado" || day.weekday === "Domingo" ? 3 : 5, "111827", 20),
    ],
  });
}

function monthSummary(days: AgendaDay[]): TableCell {
  const first = days[0];
  const month = first ? `${first.month} ${first.year}` : "";

  return new TableCell({
    width: { size: 50, type: WidthType.PERCENTAGE },
    verticalAlign: VerticalAlign.TOP,
    margins: { top: 140, right: 150, bottom: 140, left: 150 },
    borders: softBorders,
    children: [
      new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: month, bold: true, size: 22, color: "111827" })],
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: "Lu  Ma  Mi  Ju  Vi  Sa  Do", bold: true, size: 16, color: "0F766E" })],
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

function lines(count: number, color: string, size: number): Paragraph[] {
  return Array.from({ length: count }, () =>
    new Paragraph({
      spacing: { before: 70, after: 70 },
      children: [new TextRun({ text: "________________________________________________", color, size })],
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
        children: [new TextRun({ text: week.join("   "), size: 16, color: "374151" })],
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
